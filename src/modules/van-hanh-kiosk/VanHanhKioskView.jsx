import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Monitor, CreditCard, Clock, RefreshCw, Download, 
  Search, Maximize2, Minimize2, CheckCircle2, ChevronLeft, 
  ChevronRight, Building2, QrCode, FileText, Pill, 
  Settings, X, Sparkles, AlertCircle, ArrowUpRight, UploadCloud, Layers
} from 'lucide-react';

import {
  DEFAULT_SHEET_CONFIG,
  DEFAULT_KIOSK_LOCATIONS,
  TARGET_AREAS,
  generateSampleKioskData,
  aggregateKioskData,
  fetchGoogleSheetData,
  exportKioskReportToExcel,
  normalizeKioskName
} from './engine';

export default function VanHanhKioskView() {
  const [rawRows, setRawRows] = useState([]);
  const [locationsMap, setLocationsMap] = useState({});
  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState(null);
  const [syncError, setSyncError] = useState(null);

  // Bộ lọc
  const [dateFilter, setDateFilter] = useState('ALL');
  const [selectedShift, setSelectedShift] = useState('ALL');
  const [selectedArea, setSelectedArea] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');

  // Tab góc nhìn (Tất cả / Bảng chi tiết)
  const [viewMode, setViewMode] = useState('DASHBOARD');

  // Phân trang
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 15;

  // Toàn màn hình TV
  const [isFullscreen, setIsFullscreen] = useState(false);

  // Modal Cấu hình
  const [showConfigModal, setShowConfigModal] = useState(false);
  const [sheetConfig, setSheetConfig] = useState(DEFAULT_SHEET_CONFIG);

  useEffect(() => {
    const map = {};
    DEFAULT_KIOSK_LOCATIONS.forEach(loc => {
      map[normalizeKioskName(loc.machine)] = loc;
    });
    setLocationsMap(map);

    const sample = generateSampleKioskData();
    setRawRows(sample);
    setLastSyncTime('16/09/2026 16:30');
  }, []);

  const handleSyncGoogleSheet = async () => {
    setIsSyncing(true);
    setSyncError(null);

    try {
      let currentLocMap = { ...locationsMap };
      if (sheetConfig.gidViTriMay) {
        try {
          const locRows = await fetchGoogleSheetData(sheetConfig.spreadsheetId, sheetConfig.gidViTriMay);
          if (locRows && locRows.length > 0) {
            locRows.forEach(r => {
              const machine = normalizeKioskName(r['Tên máy'] || r['Tên KIOS'] || '');
              const area = r['Khu vực'] || '';
              const config = r['Cấu hình'] || 'CSKH hỗ trợ';
              if (machine) currentLocMap[machine] = { area, machine, config };
            });
            setLocationsMap(currentLocMap);
          }
        } catch (e) {
          console.warn('Lỗi đọc sheet vị trí:', e);
        }
      }

      const dataRows = await fetchGoogleSheetData(sheetConfig.spreadsheetId, sheetConfig.gidDuLieuTong);
      if (dataRows && dataRows.length > 0) {
        setRawRows(dataRows);
        setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
      }
    } catch (err) {
      setSyncError(err.message || 'Lỗi kết nối Google Sheet.');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleUploadExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet);
        if (rows && rows.length > 0) {
          setRawRows(rows);
          setLastSyncTime(`${new Date().toLocaleTimeString('vi-VN')} (${file.name})`);
          setSyncError(null);
        }
      } catch (err) {
        alert('Lỗi đọc tệp Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const analytics = useMemo(() => {
    return aggregateKioskData(rawRows, locationsMap, {
      dateFilter,
      selectedShift,
      selectedArea,
      searchTerm
    });
  }, [rawRows, locationsMap, dateFilter, selectedShift, selectedArea, searchTerm]);

  const totalPages = Math.ceil(analytics.filteredRows.length / pageSize) || 1;
  const paginatedItems = analytics.filteredRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  // Chiều cao tối đa cho biểu đồ lịch sử ngày
  const maxDayVal = useMemo(() => {
    return Math.max(...analytics.historicalDays.map(d => d.checkin + d.vienPhi + d.banThuoc), 3200);
  }, [analytics.historicalDays]);

  // Chiều cao tối đa cho biểu đồ xếp chồng khu vực
  const maxAreaVal = useMemo(() => {
    return Math.max(...analytics.areaBreakdown.map(a => a.total), 350);
  }, [analytics.areaBreakdown]);

  return (
    <div className={`space-y-6 pb-12 transition-all font-sans ${isFullscreen ? 'fixed inset-0 z-50 bg-slate-900 text-slate-100 p-6 overflow-y-auto' : ''}`}>
      {/* 1. HEADER CHÍNH */}
      <div className={`p-5 rounded-2xl border shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 ${
        isFullscreen ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
      }`}>
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-teal-600 text-white flex items-center justify-center font-bold shadow-xs">
            <Monitor className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-xl font-extrabold text-blue-600 tracking-tight">
              [TB] BÁO CÁO HOẠT ĐỘNG KIOSK
            </h1>
            <p className="text-xs text-slate-500 font-medium">
              Theo dõi trực tiếp dữ liệu hoạt động tiếp đón, viện phí và máy tự phục vụ
            </p>
          </div>
        </div>

        {/* Cụm Điều Khiển */}
        <div className="flex flex-wrap items-center gap-2">
          {lastSyncTime && (
            <span className="text-xs px-3 py-1.5 rounded-lg border bg-slate-50 border-slate-200 text-slate-600 font-medium flex items-center gap-1.5">
              <Clock className="w-3.5 h-3.5 text-teal-600" />
              {lastSyncTime}
            </span>
          )}

          <button
            onClick={handleSyncGoogleSheet}
            disabled={isSyncing}
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold text-white bg-teal-600 hover:bg-teal-700 transition cursor-pointer disabled:opacity-50"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
            {isSyncing ? 'Đang tải...' : 'Lấy từ Google Sheet'}
          </button>

          <button
            onClick={() => setShowConfigModal(true)}
            className="p-1.5 rounded-xl border bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700 transition cursor-pointer"
            title="Cài đặt link Google Sheet"
          >
            <Settings className="w-4 h-4" />
          </button>

          <button
            onClick={() => exportKioskReportToExcel(analytics.filteredRows)}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl border bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700 text-xs font-semibold transition cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            Tải Excel
          </button>

          <button
            onClick={() => setIsFullscreen(!isFullscreen)}
            className="p-1.5 rounded-xl border bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700 transition cursor-pointer"
            title={isFullscreen ? 'Thu nhỏ' : 'Toàn màn hình (chiếu TV)'}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* CẢNH BÁO NẾU SHEET CHƯA MỞ QUYỀN */}
      {syncError && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-3.5 text-amber-900 flex items-center justify-between text-xs">
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-amber-600 flex-shrink-0" />
            <span><b>Google Sheet chưa mở chia sẻ:</b> Hãy mở sheet chọn Chia sẻ &gt; "Bất kỳ ai có đường liên kết đều có thể xem". Hoặc tải file Excel trực tiếp.</span>
          </div>
          <label className="inline-flex items-center gap-1.5 px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded-lg font-semibold cursor-pointer">
            <UploadCloud className="w-3.5 h-3.5" />
            Nạp file Excel
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleUploadExcel} />
          </label>
        </div>
      )}

      {/* 2. KHU VỰC TỔNG QUAN TRONG NGÀY (TRỰC QUAN NHƯ DATA STUDIO CỦA BẠN) */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-5">
        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
          <h2 className="text-base font-black text-blue-700 uppercase tracking-wide">
            TỔNG QUAN TRONG NGÀY
          </h2>
          <div className="flex items-center gap-2 text-xs">
            <span className="font-semibold text-slate-500">Mục tiêu ngày:</span>
            <span className="px-2 py-0.5 rounded-md bg-blue-50 text-blue-700 font-bold font-mono">
              3,000 lượt
            </span>
          </div>
        </div>

        {/* HÀNG 1: TỔNG LƯỢT - TỶ LỆ CHUNG - SỐ LƯỢT THEO MỤC TIÊU */}
        <div className="grid grid-cols-1 md:grid-cols-12 gap-4">
          {/* Ô 1: TỔNG LƯỢT */}
          <div className="md:col-span-3 border-2 border-rose-300 bg-rose-50/30 rounded-2xl p-4 flex flex-col items-center justify-center text-center">
            <span className="text-xs font-extrabold uppercase text-blue-700 tracking-wider">
              TỔNG LƯỢT
            </span>
            <span className="text-4xl lg:text-5xl font-black font-mono text-rose-500 mt-1">
              {analytics.summary.totalOps.toLocaleString()}
            </span>
          </div>

          {/* Ô 2: TỶ LỆ CHUNG */}
          <div className="md:col-span-3 border-2 border-amber-300 bg-amber-50/30 rounded-2xl p-4 flex flex-col items-center justify-center text-center">
            <span className="text-xs font-extrabold uppercase text-blue-700 tracking-wider">
              TỶ LỆ CHUNG
            </span>
            <span className="text-4xl lg:text-5xl font-black font-mono text-amber-500 mt-1">
              {analytics.summary.overallRate}%
            </span>
          </div>

          {/* Ô 3: SỐ LƯỢT THEO CÁC MỤC TIÊU (THANH NGANG CHIA 3 MÀU + MỐC 3000) */}
          <div className="md:col-span-6 border border-slate-200 rounded-2xl p-4 flex flex-col justify-between">
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-extrabold uppercase text-blue-700 tracking-wider">
                SỐ LƯỢT THEO CÁC MỤC TIÊU
              </span>
              <div className="flex items-center gap-3 text-xs font-semibold">
                <span className="flex items-center gap-1.5 text-blue-600">
                  <span className="w-3 h-3 rounded-sm bg-blue-500"></span> Checkin
                </span>
                <span className="flex items-center gap-1.5 text-amber-600">
                  <span className="w-3 h-3 rounded-sm bg-amber-500"></span> Thu viện phí
                </span>
                <span className="flex items-center gap-1.5 text-emerald-600">
                  <span className="w-3 h-3 rounded-sm bg-emerald-500"></span> Bán thuốc
                </span>
              </div>
            </div>

            {/* Thanh thanh đo phân chia tỷ lệ */}
            <div className="space-y-1.5 pt-1">
              <div className="w-full bg-slate-100 h-9 rounded-xl overflow-hidden flex relative border border-slate-200">
                {/* Checkin */}
                <div 
                  className="bg-blue-500 h-full flex items-center justify-center text-white text-xs font-bold font-mono transition-all"
                  style={{ width: `${(analytics.summary.totalCheckin / 3000) * 100}%` }}
                  title={`Checkin: ${analytics.summary.totalCheckin.toLocaleString()}`}
                >
                  {analytics.summary.totalCheckin.toLocaleString()}
                </div>

                {/* Thu viện phí */}
                <div 
                  className="bg-amber-500 h-full flex items-center justify-center text-white text-xs font-bold font-mono transition-all"
                  style={{ width: `${(analytics.summary.totalVienPhi / 3000) * 100}%` }}
                  title={`Thu viện phí: ${analytics.summary.totalVienPhi.toLocaleString()}`}
                >
                  {analytics.summary.totalVienPhi.toLocaleString()}
                </div>

                {/* Bán thuốc */}
                <div 
                  className="bg-emerald-600 h-full flex items-center justify-center text-white text-xs font-bold font-mono transition-all"
                  style={{ width: `${Math.max(2, (analytics.summary.totalBanThuoc / 3000) * 100)}%` }}
                  title={`Bán thuốc: ${analytics.summary.totalBanThuoc.toLocaleString()}`}
                >
                  {analytics.summary.totalBanThuoc}
                </div>

                {/* Vạch mốc Mục tiêu (3,000) */}
                <div 
                  className="absolute top-0 bottom-0 border-r-2 border-dashed border-amber-500 z-10"
                  style={{ left: '100%' }}
                />
              </div>

              {/* Thước đo số lượt */}
              <div className="flex justify-between text-[11px] text-slate-400 font-mono pt-0.5">
                <span>0</span>
                <span>500</span>
                <span>1K</span>
                <span>1.5K</span>
                <span>2K</span>
                <span>2.5K</span>
                <span className="font-bold text-amber-600">Mục tiêu (3000)</span>
              </div>
            </div>
          </div>
        </div>

        {/* HÀNG 2: TỶ LỆ HOÀN THÀNH MỤC TIÊU & TỶ LỆ HÌNH THỨC THANH TOÁN */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* Cụm 1: TỶ LỆ HOÀN THÀNH MỤC TIÊU */}
          <div className="border border-slate-200 rounded-2xl p-4">
            <span className="text-xs font-extrabold uppercase text-blue-700 tracking-wider block mb-3 text-center">
              TỶ LỆ HOÀN THÀNH MỤC TIÊU
            </span>
            <div className="grid grid-cols-2 gap-3">
              {/* Đăng ký & Thu CLS */}
              <div className="border border-rose-300 bg-rose-50/20 rounded-xl p-3 flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-blue-100 text-blue-700 flex items-center justify-center flex-shrink-0">
                  <FileText className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] font-bold text-rose-700 block uppercase">Đăng ký &amp; Thu CLS</span>
                  <span className="text-2xl font-black font-mono text-rose-600">
                    {analytics.summary.regCompletedRate}%
                  </span>
                </div>
              </div>

              {/* Bán thuốc */}
              <div className="border border-emerald-300 bg-emerald-50/20 rounded-xl p-3 flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center flex-shrink-0">
                  <Pill className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] font-bold text-emerald-800 block uppercase">Bán thuốc</span>
                  <span className="text-2xl font-black font-mono text-emerald-700">
                    {analytics.summary.medCompletedRate}%
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Cụm 2: TỶ LỆ HÌNH THỨC THANH TOÁN */}
          <div className="border border-slate-200 rounded-2xl p-4">
            <span className="text-xs font-extrabold uppercase text-blue-700 tracking-wider block mb-3 text-center">
              TỶ LỆ HÌNH THỨC THANH TOÁN
            </span>
            <div className="grid grid-cols-2 gap-3">
              {/* Cà thẻ */}
              <div className="border border-teal-300 bg-teal-50/20 rounded-xl p-3 flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-teal-100 text-teal-700 flex items-center justify-center flex-shrink-0">
                  <CreditCard className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] font-bold text-teal-800 block uppercase">CÀ THẺ (POS)</span>
                  <span className="text-2xl font-black font-mono text-teal-700">
                    {analytics.summary.ratePOS}%
                  </span>
                </div>
              </div>

              {/* Chuyển khoản */}
              <div className="border border-purple-300 bg-purple-50/20 rounded-xl p-3 flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-purple-100 text-purple-700 flex items-center justify-center flex-shrink-0">
                  <QrCode className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] font-bold text-purple-800 block uppercase">CHUYỂN KHOẢN (QR)</span>
                  <span className="text-2xl font-black font-mono text-purple-700">
                    {analytics.summary.rateCK}%
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* 3. KHU VỰC ĐẶC BIỆT THEO YÊU CẦU: PHÂN TÍCH RIÊNG CÁC MÁY "TỰ THỰC HIỆN" */}
      <div className="bg-gradient-to-r from-purple-50 via-indigo-50/40 to-white p-5 rounded-2xl border-2 border-purple-200 shadow-xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-purple-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-purple-600 text-white flex items-center justify-center font-bold">
              <Sparkles className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-black text-purple-900 uppercase tracking-wide">
                PHÂN TÍCH CHUYÊN BIỆT: MÁY CẤU HÌNH "TỰ THỰC HIỆN" (BỆNH NHÂN TỰ BẤM MÁY)
              </h3>
              <p className="text-xs text-purple-700">
                Theo dõi mức độ người bệnh tự thực hiện độc lập tại các máy Kiosk (ví dụ KIOS17 ở Sảnh chính)
              </p>
            </div>
          </div>

          <span className="px-3 py-1 bg-purple-100 text-purple-800 rounded-full text-xs font-bold font-mono">
            Tỷ lệ tự làm: {analytics.summary.selfShareRate}% toàn viện
          </span>
        </div>

        {/* 4 Thẻ chỉ số máy Tự Thực Hiện */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <div className="bg-white p-3.5 rounded-xl border border-purple-200">
            <span className="text-[11px] font-bold text-slate-500 uppercase">Tổng lượt Tự làm</span>
            <p className="text-2xl font-black font-mono text-purple-700 mt-1">
              {analytics.summary.selfTotal.toLocaleString()}
            </p>
            <span className="text-[10px] text-slate-400">Không cần CSKH can thiệp</span>
          </div>

          <div className="bg-white p-3.5 rounded-xl border border-blue-200">
            <span className="text-[11px] font-bold text-blue-700 uppercase">Tự Check-in</span>
            <p className="text-2xl font-black font-mono text-blue-600 mt-1">
              {analytics.summary.selfCheckin.toLocaleString()}
            </p>
            <span className="text-[10px] text-slate-400">
              {analytics.summary.selfTotal > 0 ? ((analytics.summary.selfCheckin / analytics.summary.selfTotal) * 100).toFixed(1) : 0}% trên tổng tự làm
            </span>
          </div>

          <div className="bg-white p-3.5 rounded-xl border border-amber-200">
            <span className="text-[11px] font-bold text-amber-700 uppercase">Tự nộp Viện phí</span>
            <p className="text-2xl font-black font-mono text-amber-600 mt-1">
              {analytics.summary.selfVienPhi.toLocaleString()}
            </p>
            <span className="text-[10px] text-slate-400">
              {analytics.summary.selfTotal > 0 ? ((analytics.summary.selfVienPhi / analytics.summary.selfTotal) * 100).toFixed(1) : 0}% trên tổng tự làm
            </span>
          </div>

          <div className="bg-white p-3.5 rounded-xl border border-emerald-200">
            <span className="text-[11px] font-bold text-emerald-700 uppercase">Thanh toán tự làm</span>
            <p className="text-lg font-black font-mono text-emerald-700 mt-1">
              CK {analytics.summary.selfRateCK}% - POS {analytics.summary.selfRatePOS}%
            </p>
            <span className="text-[10px] text-slate-400">Tự quét mã QR tại máy</span>
          </div>
        </div>

        {/* Biểu đồ Nhịp Giờ: Lúc nào người bệnh tự bấm máy đông nhất? */}
        <div className="bg-white p-4 rounded-xl border border-purple-200">
          <div className="flex items-center justify-between mb-3 text-xs">
            <span className="font-bold text-purple-900">
              Lưu lượng người bệnh tự thao tác theo từng khung giờ (06:00 - 18:00)
            </span>
            <span className="text-slate-500 font-medium">
              Đỉnh tự làm nhiều nhất: <b>08:00 - 10:00</b>
            </span>
          </div>

          <div className="h-28 flex items-end gap-1.5 pt-4 pb-1 border-b border-slate-100">
            {analytics.hourlySelfList.map((item) => {
              const maxSelf = Math.max(...analytics.hourlySelfList.map(s => s.total), 1);
              const pct = Math.round((item.total / maxSelf) * 100);

              return (
                <div key={item.hour} className="flex-1 flex flex-col items-center h-full justify-end group relative">
                  <div className="absolute -top-7 opacity-0 group-hover:opacity-100 transition bg-slate-800 text-white text-[10px] rounded px-1.5 py-0.5 z-20 whitespace-nowrap pointer-events-none">
                    {item.hour}: {item.total} lượt tự làm
                  </div>
                  <div 
                    className="w-full max-w-[20px] bg-purple-600 rounded-t-sm transition-all"
                    style={{ height: `${Math.max(4, pct)}%` }}
                  />
                  <span className="text-[9px] font-mono text-slate-400 mt-1">
                    {item.hour.split(':')[0]}h
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* 4. HAI BIỂU ĐỒ NGANG THEO KHU VỰC (CHECKIN & VIỆN PHÍ VS BÁN THUỐC) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
        {/* Biểu đồ trái: Checkin vs Viện phí tại 11 khu vực */}
        <div className="lg:col-span-8 bg-white p-5 rounded-2xl border border-slate-200 shadow-xs">
          <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-2">
            <h3 className="text-xs font-extrabold uppercase text-blue-700">
              SỐ LƯỢT CHECK-IN &amp; THU VIỆN PHÍ THEO KHU VỰC
            </h3>
            <div className="flex items-center gap-3 text-xs font-semibold">
              <span className="flex items-center gap-1.5 text-blue-600">
                <span className="w-3 h-3 rounded-sm bg-blue-500"></span> Checkin
              </span>
              <span className="flex items-center gap-1.5 text-amber-600">
                <span className="w-3 h-3 rounded-sm bg-amber-500"></span> Thu viện phí
              </span>
            </div>
          </div>

          {/* Các thanh ngang cho 11 khu vực */}
          <div className="space-y-2.5">
            {analytics.areaBreakdown.map((item) => {
              const maxSingle = 350;
              const wCheckin = Math.min(100, (item.checkin / maxSingle) * 100);
              const wVienPhi = Math.min(100, (item.vienPhi / maxSingle) * 100);

              return (
                <div key={item.area} className="space-y-1">
                  <div className="flex items-center justify-between text-xs font-medium">
                    <span className="text-slate-800 truncate max-w-[190px]">{item.area}</span>
                    <span className="text-[11px] font-mono text-slate-500">
                      Checkin: <b className="text-blue-600">{item.checkin}</b> | VP: <b className="text-amber-600">{item.vienPhi}</b>
                    </span>
                  </div>
                  <div className="flex h-5 bg-slate-100 rounded-md overflow-hidden">
                    <div 
                      className="bg-blue-500 text-white text-[10px] font-bold font-mono flex items-center justify-center transition-all"
                      style={{ width: `${wCheckin}%` }}
                    >
                      {item.checkin > 30 ? item.checkin : ''}
                    </div>
                    <div 
                      className="bg-amber-500 text-white text-[10px] font-bold font-mono flex items-center justify-center transition-all"
                      style={{ width: `${wVienPhi}%` }}
                    >
                      {item.vienPhi > 30 ? item.vienPhi : ''}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Biểu đồ phải: Bán thuốc theo khu vực */}
        <div className="lg:col-span-4 bg-white p-5 rounded-2xl border border-slate-200 shadow-xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-2">
              <h3 className="text-xs font-extrabold uppercase text-blue-700">
                SỐ LƯỢT BÁN THUỐC
              </h3>
              <span className="flex items-center gap-1.5 text-emerald-600 text-xs font-semibold">
                <span className="w-3 h-3 rounded-sm bg-emerald-600"></span> Bán thuốc
              </span>
            </div>

            <div className="space-y-2.5">
              {analytics.areaBreakdown.map((item) => {
                const maxMed = 35;
                const w = Math.min(100, (item.banThuoc / maxMed) * 100);

                return (
                  <div key={item.area} className="space-y-1">
                    <div className="flex items-center justify-between text-xs font-medium">
                      <span className="text-slate-800 truncate max-w-[130px]">{item.area}</span>
                      <span className="text-[11px] font-mono font-bold text-emerald-700">
                        {item.banThuoc}
                      </span>
                    </div>
                    <div className="h-5 bg-slate-100 rounded-md overflow-hidden flex">
                      <div 
                        className="bg-emerald-600 text-white text-[10px] font-bold font-mono flex items-center justify-center transition-all"
                        style={{ width: `${Math.max(item.banThuoc > 0 ? 8 : 0, w)}%` }}
                      >
                        {item.banThuoc > 0 ? item.banThuoc : ''}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      {/* 5. BIỂU ĐỒ CỘT XẾP CHỒNG DỌC THEO 11 KHU VỰC (ĐÚNG NHƯ DATA STUDIO CỦA BẠN) */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-2">
          <h3 className="text-xs font-extrabold uppercase text-blue-700">
            TỔNG HỢP CỘT XẾP CHỒNG THEO KHU VỰC (CHECKIN + THU VIỆN PHÍ + BÁN THUỐC)
          </h3>
          <div className="flex items-center gap-3 text-xs font-semibold">
            <span className="flex items-center gap-1.5 text-blue-600">
              <span className="w-3 h-3 rounded-sm bg-blue-500"></span> Checkin
            </span>
            <span className="flex items-center gap-1.5 text-amber-600">
              <span className="w-3 h-3 rounded-sm bg-amber-500"></span> Thu viện phí
            </span>
            <span className="flex items-center gap-1.5 text-emerald-600">
              <span className="w-3 h-3 rounded-sm bg-emerald-600"></span> Bán thuốc
            </span>
          </div>
        </div>

        {/* Khung Biểu Đồ Cột Đứng */}
        <div className="h-64 flex items-end gap-2 pt-6 pb-2 px-1 border-b border-slate-100">
          {analytics.areaBreakdown.map((item) => {
            const hCheckin = (item.checkin / maxAreaVal) * 100;
            const hVienPhi = (item.vienPhi / maxAreaVal) * 100;
            const hBanThuoc = (item.banThuoc / maxAreaVal) * 100;

            return (
              <div key={item.area} className="flex-1 flex flex-col items-center h-full justify-end group relative">
                {/* Số tổng trên đỉnh cột */}
                <span className="text-[10px] font-bold font-mono text-slate-700 mb-1">
                  {item.total}
                </span>

                {/* Cột xếp chồng */}
                <div className="w-full max-w-[42px] rounded-t-md overflow-hidden flex flex-col justify-end bg-slate-100">
                  {item.banThuoc > 0 && (
                    <div 
                      className="bg-emerald-600 text-white text-[9px] font-bold font-mono flex items-center justify-center transition-all"
                      style={{ height: `${hBanThuoc}%`, minHeight: '14px' }}
                      title={`Bán thuốc: ${item.banThuoc}`}
                    >
                      {item.banThuoc}
                    </div>
                  )}
                  {item.vienPhi > 0 && (
                    <div 
                      className="bg-amber-500 text-white text-[9px] font-bold font-mono flex items-center justify-center transition-all"
                      style={{ height: `${hVienPhi}%`, minHeight: '18px' }}
                      title={`Viện phí: ${item.vienPhi}`}
                    >
                      {item.vienPhi}
                    </div>
                  )}
                  {item.checkin > 0 && (
                    <div 
                      className="bg-blue-500 text-white text-[10px] font-bold font-mono flex items-center justify-center transition-all"
                      style={{ height: `${hCheckin}%`, minHeight: '22px' }}
                      title={`Checkin: ${item.checkin}`}
                    >
                      {item.checkin}
                    </div>
                  )}
                </div>

                {/* Tên khu vực dưới chân cột */}
                <span 
                  className="text-[10px] font-medium text-slate-600 mt-2 truncate max-w-[50px] text-center" 
                  title={item.area}
                >
                  {item.area.split('-')[0].trim()}
                </span>
              </div>
            );
          })}
        </div>
      </div>

      {/* 6. BIỂU ĐỒ LỊCH SỬ CHUỖI NGÀY & ĐƯỜNG MỤC TIÊU 3,000 (KHỚP 100% VỚI BẢN CŨ CỦA BẠN) */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-2">
          <div className="flex items-center gap-2">
            <h3 className="text-xs font-extrabold uppercase text-blue-700">
              LỊCH SỬ HOẠT ĐỘNG THEO NGÀY (THÁNG 8 - THÁNG 9)
            </h3>
            <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold">
              Vạch chỉ tiêu: 3,000 lượt/ngày
            </span>
          </div>

          <div className="flex items-center gap-3 text-xs font-semibold">
            <span className="flex items-center gap-1.5 text-blue-600">
              <span className="w-3 h-3 rounded-sm bg-blue-500"></span> Checkin
            </span>
            <span className="flex items-center gap-1.5 text-amber-600">
              <span className="w-3 h-3 rounded-sm bg-amber-500"></span> Thu viện phí
            </span>
            <span className="flex items-center gap-1.5 text-emerald-600">
              <span className="w-3 h-3 rounded-sm bg-emerald-600"></span> Bán thuốc
            </span>
          </div>
        </div>

        {/* Biểu đồ Cột Chuỗi Ngày Kèm Đường Mục Tiêu 3,000 */}
        <div className="relative pt-6">
          {/* Đường mục tiêu 3,000 gạch đứt màu vàng */}
          <div 
            className="absolute left-0 right-0 border-t-2 border-dashed border-amber-500 z-10 flex items-center justify-between pointer-events-none"
            style={{ top: '22%' }}
          >
            <span className="bg-amber-500 text-white text-[10px] font-bold font-mono px-2 py-0.5 rounded-r shadow-xs">
              Mục tiêu (3,000)
            </span>
          </div>

          <div className="h-64 flex items-end gap-1.5 pb-2 overflow-x-auto">
            {analytics.historicalDays.map((d) => {
              const total = d.checkin + d.vienPhi + d.banThuoc;
              const hTotal = (total / maxDayVal) * 100;
              const isOverTarget = total >= 3000;

              return (
                <div key={d.date} className="min-w-[34px] flex-1 flex flex-col items-center h-full justify-end group relative">
                  {/* Tooltip khi di chuột */}
                  <div className="absolute -top-12 opacity-0 group-hover:opacity-100 transition bg-slate-900 text-white text-[10px] rounded p-1.5 whitespace-nowrap z-30 pointer-events-none shadow-lg">
                    <p className="font-bold">{d.date}: {total.toLocaleString()} lượt</p>
                    <p>Checkin: {d.checkin} | VP: {d.vienPhi} | Thuốc: {d.banThuoc}</p>
                  </div>

                  {/* Số thuốc trên đỉnh */}
                  <span className="text-[8px] font-bold font-mono text-emerald-700">
                    {d.banThuoc}
                  </span>

                  {/* Cột xếp chồng */}
                  <div 
                    className={`w-full rounded-t-sm overflow-hidden flex flex-col justify-end ${
                      isOverTarget ? 'ring-2 ring-emerald-500' : ''
                    }`}
                    style={{ height: `${hTotal}%` }}
                  >
                    <div 
                      className="bg-emerald-600 text-white text-[8px] font-bold flex items-center justify-center"
                      style={{ height: `${(d.banThuoc / total) * 100}%`, minHeight: '6px' }}
                    />
                    <div 
                      className="bg-amber-500 text-white text-[8px] font-bold flex items-center justify-center font-mono"
                      style={{ height: `${(d.vienPhi / total) * 100}%`, minHeight: '12px' }}
                    >
                      {d.vienPhi}
                    </div>
                    <div 
                      className="bg-blue-500 text-white text-[9px] font-bold flex items-center justify-center font-mono"
                      style={{ height: `${(d.checkin / total) * 100}%`, minHeight: '24px' }}
                    >
                      {d.checkin}
                    </div>
                  </div>

                  {/* Nhãn ngày bên dưới */}
                  <span className="text-[9px] font-mono text-slate-500 mt-2 truncate max-w-[32px] text-center rotate-45 origin-left">
                    {d.date}
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* 7. BẢNG DỮ LIỆU TRA CỨU CHI TIẾT (DRILL-DOWN) */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h3 className="text-sm font-bold text-slate-900">Nhật Ký Thao Tác Kiosk Chi Tiết</h3>
            <p className="text-xs text-slate-500">Tra cứu nhanh theo mã bệnh nhân (PID), tên bệnh nhân hoặc Kiosk</p>
          </div>

          <div className="relative w-full sm:w-64">
            <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
            <input
              type="text"
              placeholder="PID, Tên, Kiosk, CSKH..."
              value={searchTerm}
              onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
              className="w-full text-xs pl-8 pr-3 py-1.5 rounded-xl border border-slate-300 focus:outline-none focus:ring-1 focus:ring-teal-500"
            />
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs text-slate-600">
            <thead className="bg-slate-50 text-slate-700 uppercase font-semibold text-[11px] border-b border-slate-200">
              <tr>
                <th className="py-2.5 px-3 w-12 text-center">STT</th>
                <th className="py-2.5 px-3">Thời Gian</th>
                <th className="py-2.5 px-3">PID</th>
                <th className="py-2.5 px-4">Họ và Tên</th>
                <th className="py-2.5 px-3">Kiosk</th>
                <th className="py-2.5 px-3">Khu Vực</th>
                <th className="py-2.5 px-3 text-center">Dịch Vụ</th>
                <th className="py-2.5 px-3 text-center">Hình Thức TT</th>
                <th className="py-2.5 px-4">Người Thao Tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {paginatedItems.map((r, i) => {
                const stt = (currentPage - 1) * pageSize + i + 1;
                const isSelf = String(r['Họ tên user'] || '').toLowerCase().includes('tự thực hiện');

                return (
                  <tr key={r['STT'] || stt} className="hover:bg-slate-50/80 transition">
                    <td className="py-2 px-3 text-center text-slate-400 font-mono text-[11px]">{stt}</td>
                    <td className="py-2 px-3 font-mono text-slate-700">{r['Ngày giờ thao tác']}</td>
                    <td className="py-2 px-3 font-mono font-bold text-slate-900">{r['PID']}</td>
                    <td className="py-2 px-4 font-medium text-slate-800">{r['Họ tên']}</td>
                    <td className="py-2 px-3 font-mono font-semibold text-teal-700">{r['Tên KIOS']}</td>
                    <td className="py-2 px-3 text-slate-600 truncate max-w-[130px]">{r['KHU VỰC']}</td>
                    <td className="py-2 px-3 text-center">
                      <div className="flex items-center justify-center gap-1">
                        {r['Checkin'] && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-100 text-blue-800">Checkin</span>}
                        {r['Thu viện phí'] && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800">Viện phí</span>}
                        {r['Bán thuốc'] && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">Thuốc</span>}
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono font-bold">
                      {r['HÌNH THỨC TT'] || '--'}
                    </td>
                    <td className="py-2 px-4 text-slate-700">
                      {isSelf ? (
                        <span className="font-semibold text-purple-700 bg-purple-50 px-2 py-0.5 rounded text-[11px]">
                          Tự thực hiện
                        </span>
                      ) : (
                        r['Họ tên user']
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Phân trang */}
        <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <span>
            Hiển thị <b>{(currentPage - 1) * pageSize + 1}</b> - <b>{Math.min(currentPage * pageSize, analytics.filteredRows.length)}</b> / <b>{analytics.filteredRows.length.toLocaleString()}</b> dòng
          </span>
          <div className="flex items-center gap-1">
            <button
              onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
              disabled={currentPage === 1}
              className="p-1.5 rounded border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 cursor-pointer"
            >
              <ChevronLeft className="w-3.5 h-3.5" />
            </button>
            <span className="px-2 font-medium">{currentPage} / {totalPages}</span>
            <button
              onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
              disabled={currentPage === totalPages}
              className="p-1.5 rounded border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 cursor-pointer"
            >
              <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      </div>

      {/* MODAL CẤU HÌNH GOOGLE SHEET */}
      {showConfigModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-lg w-full p-6 text-slate-800 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-bold text-sm">Cấu Hình Google Sheet</h3>
              <button onClick={() => setShowConfigModal(false)} className="text-slate-400 hover:text-slate-600 cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold mb-1">Spreadsheet ID</label>
                <input
                  type="text"
                  value={sheetConfig.spreadsheetId}
                  onChange={(e) => setSheetConfig({ ...sheetConfig, spreadsheetId: e.target.value.trim() })}
                  className="w-full p-2 border rounded-xl font-mono text-xs"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold mb-1">GID Dữ Liệu Tổng</label>
                  <input
                    type="text"
                    value={sheetConfig.gidDuLieuTong}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidDuLieuTong: e.target.value.trim() })}
                    className="w-full p-2 border rounded-xl font-mono text-xs"
                  />
                </div>
                <div>
                  <label className="block font-semibold mb-1">GID Vị Trí Máy</label>
                  <input
                    type="text"
                    value={sheetConfig.gidViTriMay}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidViTriMay: e.target.value.trim() })}
                    className="w-full p-2 border rounded-xl font-mono text-xs"
                  />
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                onClick={() => setShowConfigModal(false)}
                className="px-4 py-2 border rounded-xl text-xs font-semibold"
              >
                Đóng
              </button>
              <button
                onClick={() => {
                  setShowConfigModal(false);
                  handleSyncGoogleSheet();
                }}
                className="px-5 py-2 bg-teal-600 text-white rounded-xl text-xs font-semibold"
              >
                Lưu &amp; Lấy Dữ Liệu
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
