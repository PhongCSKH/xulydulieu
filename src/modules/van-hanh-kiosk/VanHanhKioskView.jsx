import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Activity, Monitor, Users, CreditCard, Clock, RefreshCw, Download, 
  Search, Filter, Maximize2, Minimize2, CheckCircle2, ChevronLeft, 
  ChevronRight, ArrowUpRight, TrendingUp, Cpu, Settings, ShieldCheck, 
  Layers, UploadCloud, X, HelpCircle, Sparkles, Building2, QrCode
} from 'lucide-react';

import {
  DEFAULT_SHEET_CONFIG,
  DEFAULT_KIOSK_LOCATIONS,
  generateSampleKioskData,
  aggregateKioskData,
  fetchGoogleSheetData,
  exportKioskReportToExcel,
  normalizeKioskName
} from './engine';

export default function VanHanhKioskView() {
  // Trạng thái dữ liệu
  const [rawRows, setRawRows] = useState([]);
  const [locationsMap, setLocationsMap] = useState({});
  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState(null);
  const [syncError, setSyncError] = useState(null);

  // Bộ lọc
  const [dateFilter, setDateFilter] = useState('ALL');
  const [selectedShift, setSelectedShift] = useState('ALL');
  const [selectedArea, setSelectedArea] = useState('ALL');
  const [selectedService, setSelectedService] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');

  // Tab góc nhìn (Dashboard / Thiết bị / Nhân sự / Dữ liệu chi tiết)
  const [activeTab, setActiveTab] = useState('OVERVIEW');

  // Phân trang cho bảng chi tiết
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 15;

  // Chế độ Command Center (Toàn màn hình)
  const [isFullscreen, setIsFullscreen] = useState(false);

  // Modal Cấu hình Google Sheet
  const [showConfigModal, setShowConfigModal] = useState(false);
  const [sheetConfig, setSheetConfig] = useState(DEFAULT_SHEET_CONFIG);

  // Khởi tạo Locations Map mặc định
  useEffect(() => {
    const map = {};
    DEFAULT_KIOSK_LOCATIONS.forEach(loc => {
      map[normalizeKioskName(loc.machine)] = loc;
    });
    setLocationsMap(map);

    // Nạp dữ liệu mẫu ban đầu để giao diện hiển thị ngay lập tức
    const sample = generateSampleKioskData();
    setRawRows(sample);
    setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
  }, []);

  // Hàm đồng bộ dữ liệu từ Google Sheet
  const handleSyncGoogleSheet = async () => {
    setIsSyncing(true);
    setSyncError(null);

    try {
      // 1. Tải Sheet Vị trí bố trí máy nếu có
      let currentLocMap = { ...locationsMap };
      if (sheetConfig.gidViTriMay) {
        try {
          const locRows = await fetchGoogleSheetData(sheetConfig.spreadsheetId, sheetConfig.gidViTriMay);
          if (locRows && locRows.length > 0) {
            locRows.forEach(r => {
              const machine = normalizeKioskName(r['Tên máy'] || r['Tên KIOS'] || r['KIOSK'] || '');
              const area = r['Khu vực'] || r['Khu Vực'] || '';
              const config = r['Cấu hình'] || r['Cấu Hình'] || 'CSKH hỗ trợ';
              if (machine) {
                currentLocMap[machine] = { area, machine, config };
              }
            });
            setLocationsMap(currentLocMap);
          }
        } catch (e) {
          console.warn('Lỗi đọc sheet Vị trí bố trí máy:', e);
        }
      }

      // 2. Tải Sheet Dữ liệu tổng
      const dataRows = await fetchGoogleSheetData(sheetConfig.spreadsheetId, sheetConfig.gidDuLieuTong);
      if (dataRows && dataRows.length > 0) {
        setRawRows(dataRows);
        setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
      } else {
        throw new Error('Dữ liệu từ Google Sheet rỗng.');
      }
    } catch (err) {
      console.error(err);
      setSyncError(err.message || 'Không thể đồng bộ dữ liệu từ Google Sheet.');
    } finally {
      setIsSyncing(false);
    }
  };

  // Upload file Excel thủ công dự phòng
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
          setLastSyncTime(`${new Date().toLocaleTimeString('vi-VN')} (Từ tệp ${file.name})`);
          setSyncError(null);
        }
      } catch (err) {
        alert('Lỗi đọc tệp Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Tổng hợp dữ liệu phân tích thông qua engine
  const analytics = useMemo(() => {
    return aggregateKioskData(rawRows, locationsMap, {
      dateFilter,
      selectedShift,
      selectedArea,
      selectedService,
      searchTerm
    });
  }, [rawRows, locationsMap, dateFilter, selectedShift, selectedArea, selectedService, searchTerm]);

  // Danh sách các ngày có trong dữ liệu
  const availableDates = useMemo(() => {
    const set = new Set();
    rawRows.forEach(r => {
      const d = r['NGÀY'] || (r['Ngày giờ thao tác'] || '').split(' ')[0];
      if (d && d.length >= 8) set.add(d);
    });
    return Array.from(set).sort().reverse();
  }, [rawRows]);

  // Danh sách khu vực
  const availableAreas = useMemo(() => {
    const set = new Set();
    rawRows.forEach(r => {
      const a = r['KHU VỰC'];
      if (a) set.add(a);
    });
    return Array.from(set).sort();
  }, [rawRows]);

  // Phân trang bảng chi tiết
  const totalPages = Math.ceil(analytics.filteredRows.length / pageSize) || 1;
  const paginatedItems = analytics.filteredRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  // Tính chiều cao tối đa của biểu đồ giờ
  const maxHourlyVal = useMemo(() => {
    return Math.max(...analytics.hourlyData.map(h => h.total), 1);
  }, [analytics.hourlyData]);

  return (
    <div className={`space-y-6 pb-12 transition-all ${isFullscreen ? 'fixed inset-0 z-50 bg-slate-900 text-slate-100 p-6 overflow-y-auto' : ''}`}>
      {/* HEADER COMMAND CENTER */}
      <div className={`flex flex-col md:flex-row md:items-center justify-between gap-4 p-5 rounded-2xl border shadow-sm transition ${
        isFullscreen ? 'bg-slate-800/80 border-slate-700 text-white' : 'bg-white border-slate-200/90 text-slate-900'
      }`}>
        <div className="flex items-center gap-3.5">
          <div className="w-11 h-11 rounded-xl bg-gradient-to-tr from-teal-500 to-emerald-400 flex items-center justify-center text-slate-950 font-bold shadow-sm">
            <Monitor className="w-5 h-5 text-slate-950" />
          </div>
          <div>
            <div className="flex items-center gap-2.5">
              <h1 className="text-xl font-bold tracking-tight">Trung Tâm Vận Hành Kiosk</h1>
              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/10 text-emerald-600 border border-emerald-500/20">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                Trực Tiếp
              </span>
            </div>
            <p className={`text-xs mt-0.5 ${isFullscreen ? 'text-slate-400' : 'text-slate-500'}`}>
              Giám sát lưu lượng tiếp đón, thu viện phí và năng suất hỗ trợ Kiosk toàn bệnh viện
            </p>
          </div>
        </div>

        {/* Cụm Nút Hành Động */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Trạng thái cập nhật */}
          {lastSyncTime && (
            <span className={`text-[11px] px-3 py-1.5 rounded-lg border font-mono flex items-center gap-1.5 ${
              isFullscreen ? 'bg-slate-700/60 border-slate-600 text-slate-300' : 'bg-slate-50 border-slate-200 text-slate-600'
            }`}>
              <Clock className="w-3.5 h-3.5 text-teal-500" />
              Cập nhật: {lastSyncTime}
            </span>
          )}

          {/* Nút Đồng Bộ Google Sheet */}
          <button
            onClick={handleSyncGoogleSheet}
            disabled={isSyncing}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-semibold text-white bg-teal-600 hover:bg-teal-700 transition shadow-sm cursor-pointer disabled:opacity-50"
            title="Kéo dữ liệu mới nhất từ Google Sheet"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
            {isSyncing ? 'Đang đồng bộ...' : 'Đồng bộ Sheet'}
          </button>

          {/* Nút Cấu hình Google Sheet */}
          <button
            onClick={() => setShowConfigModal(true)}
            className={`p-2 rounded-xl border text-xs font-medium transition cursor-pointer ${
              isFullscreen ? 'bg-slate-700 border-slate-600 hover:bg-slate-600 text-slate-200' : 'bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700'
            }`}
            title="Cài đặt liên kết Google Sheet"
          >
            <Settings className="w-4 h-4" />
          </button>

          {/* Nút Xuất Excel */}
          <button
            onClick={() => exportKioskReportToExcel(analytics.filteredRows)}
            className={`inline-flex items-center gap-1.5 px-3 py-2 rounded-xl border text-xs font-semibold transition cursor-pointer ${
              isFullscreen ? 'bg-slate-700 border-slate-600 hover:bg-slate-600 text-slate-200' : 'bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700'
            }`}
          >
            <Download className="w-3.5 h-3.5" />
            Xuất Excel
          </button>

          {/* Nút Toàn màn hình Command Center */}
          <button
            onClick={() => setIsFullscreen(!isFullscreen)}
            className={`p-2 rounded-xl border text-xs font-medium transition cursor-pointer ${
              isFullscreen ? 'bg-teal-600 border-teal-500 text-white' : 'bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700'
            }`}
            title={isFullscreen ? 'Thu nhỏ cửa sổ' : 'Chế độ Wallboard Toàn màn hình (Treo TV)'}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* CẢNH BÁO LỖI KẾT NỐI (NẾU CÓ) */}
      {syncError && (
        <div className="bg-amber-500/10 border border-amber-500/30 rounded-2xl p-4 text-amber-700 flex items-start justify-between">
          <div className="flex items-start gap-3">
            <HelpCircle className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
            <div className="text-xs space-y-1">
              <p className="font-bold">Chưa thể tải trực tiếp từ Google Sheet:</p>
              <p>{syncError}</p>
              <p className="text-[11px] text-amber-800">
                👉 <b>Khắc phục nhanh:</b> Hãy mở Google Sheet của bạn &gt; Nhấn nút <b>Chia sẻ</b> &gt; Đổi quyền thành <b>"Bất kỳ ai có đường liên kết đều có thể xem"</b>. Hoặc bạn có thể tải tệp Excel lên trực tiếp bên dưới.
              </p>
            </div>
          </div>
          <label className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-semibold cursor-pointer transition">
            <UploadCloud className="w-3.5 h-3.5" />
            Nạp tệp Excel
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleUploadExcel} />
          </label>
        </div>
      )}

      {/* THANH ĐIỀU KHIỂN & BỘ LỌC ĐA CHIỀU */}
      <div className={`p-4 rounded-2xl border shadow-sm flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3.5 ${
        isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
      }`}>
        {/* Nhóm Filter Ngày, Ca, Khu vực */}
        <div className="flex flex-wrap items-center gap-2 text-xs">
          {/* Lọc Ngày */}
          <select
            value={dateFilter}
            onChange={(e) => { setDateFilter(e.target.value); setCurrentPage(1); }}
            className={`px-3 py-1.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-1 focus:ring-teal-500 ${
              isFullscreen ? 'bg-slate-700 border-slate-600 text-white' : 'bg-slate-50 border-slate-300 text-slate-800'
            }`}
          >
            <option value="ALL">Tất cả ngày ({rawRows.length.toLocaleString()} GD)</option>
            {availableDates.map(d => (
              <option key={d} value={d}>Ngày: {d}</option>
            ))}
          </select>

          {/* Lọc Ca Trực */}
          <div className={`inline-flex rounded-xl p-1 border ${
            isFullscreen ? 'bg-slate-700/60 border-slate-600' : 'bg-slate-100 border-slate-200'
          }`}>
            <button
              onClick={() => { setSelectedShift('ALL'); setCurrentPage(1); }}
              className={`px-2.5 py-1 rounded-lg text-xs font-medium transition ${
                selectedShift === 'ALL'
                  ? 'bg-teal-600 text-white shadow-xs'
                  : isFullscreen ? 'text-slate-300 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Cả ngày
            </button>
            <button
              onClick={() => { setSelectedShift('Buổi sáng'); setCurrentPage(1); }}
              className={`px-2.5 py-1 rounded-lg text-xs font-medium transition ${
                selectedShift === 'Buổi sáng'
                  ? 'bg-teal-600 text-white shadow-xs'
                  : isFullscreen ? 'text-slate-300 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Sáng (06h-12h)
            </button>
            <button
              onClick={() => { setSelectedShift('Buổi chiều'); setCurrentPage(1); }}
              className={`px-2.5 py-1 rounded-lg text-xs font-medium transition ${
                selectedShift === 'Buổi chiều'
                  ? 'bg-teal-600 text-white shadow-xs'
                  : isFullscreen ? 'text-slate-300 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Chiều (12h-18h)
            </button>
          </div>

          {/* Lọc Khu Vực */}
          <select
            value={selectedArea}
            onChange={(e) => { setSelectedArea(e.target.value); setCurrentPage(1); }}
            className={`px-3 py-1.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-1 focus:ring-teal-500 max-w-[170px] truncate ${
              isFullscreen ? 'bg-slate-700 border-slate-600 text-white' : 'bg-slate-50 border-slate-300 text-slate-800'
            }`}
          >
            <option value="ALL">Tất cả Khu vực</option>
            {availableAreas.map(a => (
              <option key={a} value={a}>{a}</option>
            ))}
          </select>

          {/* Lọc Dịch Vụ */}
          <select
            value={selectedService}
            onChange={(e) => { setSelectedService(e.target.value); setCurrentPage(1); }}
            className={`px-3 py-1.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-1 focus:ring-teal-500 ${
              isFullscreen ? 'bg-slate-700 border-slate-600 text-white' : 'bg-slate-50 border-slate-300 text-slate-800'
            }`}
          >
            <option value="ALL">Mọi loại dịch vụ</option>
            <option value="CHECKIN">Chỉ Check-in</option>
            <option value="VIEN_PHI">Chỉ Thu viện phí</option>
            <option value="BAN_THUOC">Chỉ Bán thuốc</option>
          </select>
        </div>

        {/* Ô Tìm Kiếm Nhanh */}
        <div className="relative w-full lg:w-64">
          <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
          <input
            type="text"
            placeholder="PID, Tên, Kiosk, CSKH..."
            value={searchTerm}
            onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
            className={`w-full text-xs pl-8 pr-3 py-2 rounded-xl border focus:outline-none focus:ring-1 focus:ring-teal-500 ${
              isFullscreen ? 'bg-slate-700 border-slate-600 text-white placeholder-slate-400' : 'bg-slate-50 border-slate-300 text-slate-900 placeholder-slate-400'
            }`}
          />
        </div>
      </div>

      {/* DẢI THẺ CHỈ SỐ CHIẾN LƯỢC (EXECUTIVE KPI CARDS) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Tổng Lượt Thao Tác */}
        <div className={`p-4 rounded-2xl border shadow-sm relative overflow-hidden flex flex-col justify-between ${
          isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Tổng Lượt Phục Vụ</span>
              <div className="w-8 h-8 rounded-lg bg-teal-500/10 text-teal-600 flex items-center justify-center">
                <Activity className="w-4 h-4" />
              </div>
            </div>
            <p className="text-3xl font-extrabold font-mono tracking-tight mt-2 text-teal-600">
              {analytics.summary.totalOps.toLocaleString()}
            </p>
          </div>
          <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">Mục tiêu ngày: 1,200</span>
            <span className="font-semibold text-emerald-600">{analytics.summary.targetPercent}% KPI</span>
          </div>
        </div>

        {/* Card 2: Check-in Tiếp Đón */}
        <div className={`p-4 rounded-2xl border shadow-sm relative overflow-hidden flex flex-col justify-between ${
          isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Check-in Tiếp Đón</span>
              <div className="w-8 h-8 rounded-lg bg-blue-500/10 text-blue-600 flex items-center justify-center">
                <QrCode className="w-4 h-4" />
              </div>
            </div>
            <p className="text-3xl font-extrabold font-mono tracking-tight mt-2 text-blue-600">
              {analytics.summary.totalCheckin.toLocaleString()}
            </p>
          </div>
          <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">Tỷ trọng lưu lượng</span>
            <span className="font-semibold text-blue-600">
              {analytics.summary.totalOps > 0 ? ((analytics.summary.totalCheckin / analytics.summary.totalOps) * 100).toFixed(1) : 0}%
            </span>
          </div>
        </div>

        {/* Card 3: Thu Viện Phí Tại Kiosk */}
        <div className={`p-4 rounded-2xl border shadow-sm relative overflow-hidden flex flex-col justify-between ${
          isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Thu Viện Phí Kiosk</span>
              <div className="w-8 h-8 rounded-lg bg-emerald-500/10 text-emerald-600 flex items-center justify-center">
                <CreditCard className="w-4 h-4" />
              </div>
            </div>
            <p className="text-3xl font-extrabold font-mono tracking-tight mt-2 text-emerald-600">
              {analytics.summary.totalVienPhi.toLocaleString()}
            </p>
          </div>
          <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">CK: {analytics.summary.countCK}</span>
            <span className="font-semibold text-teal-600">POS: {analytics.summary.countPOS}</span>
          </div>
        </div>

        {/* Card 4: Tự Thực Hiện vs CSKH Hỗ Trợ */}
        <div className={`p-4 rounded-2xl border shadow-sm relative overflow-hidden flex flex-col justify-between ${
          isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Tự Động Hóa Hành Vi</span>
              <div className="w-8 h-8 rounded-lg bg-purple-500/10 text-purple-600 flex items-center justify-center">
                <Users className="w-4 h-4" />
              </div>
            </div>
            <div className="flex items-baseline justify-between mt-2">
              <span className="text-2xl font-bold font-mono text-purple-600">
                {analytics.summary.selfServiceRate}%
              </span>
              <span className="text-xs font-medium text-slate-500">Tự thực hiện</span>
            </div>
          </div>
          <div className="mt-3 pt-3 border-t border-slate-100 space-y-1.5">
            <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden flex">
              <div 
                className="bg-purple-600 h-full" 
                style={{ width: `${analytics.summary.selfServiceRate}%` }} 
                title={`Tự thực hiện: ${analytics.summary.countSelfService}`}
              />
              <div 
                className="bg-teal-500 h-full" 
                style={{ width: `${analytics.summary.assistedRate}%` }} 
                title={`CSKH hỗ trợ: ${analytics.summary.countAssisted}`}
              />
            </div>
            <div className="flex justify-between text-[10px] text-slate-500">
              <span>Tự làm: {analytics.summary.countSelfService}</span>
              <span>CSKH: {analytics.summary.countAssisted}</span>
            </div>
          </div>
        </div>
      </div>

      {/* BIỂU ĐỒ NHỊP SINH HỌC THEO GIỜ & PHÂN PHỐI DỊCH VỤ */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        {/* Biểu đồ Nhịp Giờ Cao Điểm (2/3 chiều rộng) */}
        <div className={`lg:col-span-2 p-5 rounded-2xl border shadow-sm flex flex-col justify-between ${
          isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
        }`}>
          <div className="flex items-center justify-between mb-4">
            <div>
              <h3 className="text-sm font-bold tracking-tight">Lưu Lượng Thao Tác Theo Khung Giờ (06:00 - 18:00)</h3>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Nhận diện đỉnh cao điểm phục vụ để điều động nhân sự tiếp đón hợp lý
              </p>
            </div>
            <div className="flex items-center gap-3 text-xs">
              <span className="flex items-center gap-1.5 text-blue-600 font-medium">
                <span className="w-2.5 h-2.5 rounded-sm bg-blue-500"></span>
                Check-in
              </span>
              <span className="flex items-center gap-1.5 text-emerald-600 font-medium">
                <span className="w-2.5 h-2.5 rounded-sm bg-emerald-500"></span>
                Viện phí
              </span>
              <span className="flex items-center gap-1.5 text-purple-600 font-medium">
                <span className="w-2.5 h-2.5 rounded-sm bg-purple-500"></span>
                Bán thuốc
              </span>
            </div>
          </div>

          {/* Trục Cột Trực Quan SVG/Tailwind */}
          <div className="h-48 flex items-end gap-2 pt-6 pb-2 px-2 border-b border-slate-100">
            {analytics.hourlyData.map((item) => {
              const heightPct = Math.round((item.total / maxHourlyVal) * 100);
              const isPeak = heightPct > 75;

              return (
                <div key={item.hour} className="flex-1 flex flex-col items-center h-full justify-end group relative">
                  {/* Tooltip khi hover */}
                  <div className="absolute -top-12 opacity-0 group-hover:opacity-100 transition pointer-events-none bg-slate-900 text-white text-[10px] rounded-lg p-1.5 shadow-lg whitespace-nowrap z-20">
                    <p className="font-bold">{item.hour}: {item.total} lượt</p>
                    <p>Checkin: {item.checkin} | VP: {item.vienPhi}</p>
                  </div>

                  {/* Cột dữ liệu xếp chồng */}
                  <div 
                    className="w-full max-w-[28px] rounded-t-md overflow-hidden flex flex-col justify-end transition-all duration-300"
                    style={{ height: `${Math.max(6, heightPct)}%` }}
                  >
                    <div 
                      className="bg-blue-500 transition-all" 
                      style={{ height: `${item.total > 0 ? (item.checkin / item.total) * 100 : 0}%` }} 
                    />
                    <div 
                      className="bg-emerald-500 transition-all" 
                      style={{ height: `${item.total > 0 ? (item.vienPhi / item.total) * 100 : 0}%` }} 
                    />
                    <div 
                      className="bg-purple-500 transition-all" 
                      style={{ height: `${item.total > 0 ? (item.banThuoc / item.total) * 100 : 0}%` }} 
                    />
                  </div>

                  {/* Nhãn giờ */}
                  <span className={`text-[10px] mt-2 font-mono ${isPeak ? 'font-bold text-teal-600' : 'text-slate-400'}`}>
                    {item.hour.split(':')[0]}h
                  </span>
                </div>
              );
            })}
          </div>

          <div className="mt-3 flex items-center justify-between text-xs text-slate-500">
            <span className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-teal-500"></span>
              Đỉnh cao điểm thường rơi vào: <b>07:30 - 10:00</b> và <b>13:30 - 14:30</b>
            </span>
            <span className="font-mono text-[11px]">Cao nhất: {maxHourlyVal} lượt/giờ</span>
          </div>
        </div>

        {/* Top Khu Vực Tải Trọng Lớn (1/3 chiều rộng) */}
        <div className={`p-5 rounded-2xl border shadow-sm flex flex-col justify-between ${
          isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between mb-3">
              <h3 className="text-sm font-bold tracking-tight">Khu Vực Đông Đúc Nhất</h3>
              <Building2 className="w-4 h-4 text-slate-400" />
            </div>
            <p className="text-[11px] text-slate-500 mb-3">
              Xếp hạng theo tổng số lượt giao dịch phát sinh
            </p>

            <div className="space-y-3">
              {analytics.areaRankings.slice(0, 5).map((ar, idx) => {
                const maxArea = analytics.areaRankings[0]?.total || 1;
                const pct = Math.round((ar.total / maxArea) * 100);

                return (
                  <div key={ar.area} className="space-y-1">
                    <div className="flex items-center justify-between text-xs">
                      <span className="font-medium truncate max-w-[170px]" title={ar.area}>
                        {idx + 1}. {ar.area}
                      </span>
                      <span className="font-mono font-bold text-teal-600">
                        {ar.total.toLocaleString()} <span className="text-[10px] text-slate-400 font-normal">lượt</span>
                      </span>
                    </div>
                    <div className="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                      <div 
                        className="bg-teal-600 h-full rounded-full transition-all duration-300" 
                        style={{ width: `${pct}%` }} 
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
            <span>Tổng cộng: {analytics.areaRankings.length} khu vực</span>
            <span className="text-teal-600 font-semibold">{analytics.summary.activeKiosksCount} máy hoạt động</span>
          </div>
        </div>
      </div>

      {/* CHUYỂN TAB GÓC NHÌN (MA TRẬN KIOSK / NHÂN SỰ CSKH / BẢNG CHI TIẾT) */}
      <div className="flex items-center gap-2 border-b border-slate-200 pb-2">
        <button
          onClick={() => setActiveTab('OVERVIEW')}
          className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-semibold transition cursor-pointer ${
            activeTab === 'OVERVIEW'
              ? 'bg-teal-600 text-white shadow-xs'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <Cpu className="w-4 h-4" />
          Ma Trận Thiết Bị Kiosk ({analytics.kioskRankings.length})
        </button>

        <button
          onClick={() => setActiveTab('STAFF')}
          className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-semibold transition cursor-pointer ${
            activeTab === 'STAFF'
              ? 'bg-teal-600 text-white shadow-xs'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <Users className="w-4 h-4" />
          Bảng Xếp Hạng CSKH ({analytics.staffRankings.length})
        </button>

        <button
          onClick={() => setActiveTab('DETAILS')}
          className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-semibold transition cursor-pointer ${
            activeTab === 'DETAILS'
              ? 'bg-teal-600 text-white shadow-xs'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <Layers className="w-4 h-4" />
          Nhật Ký Giao Dịch Chi Tiết ({analytics.filteredRows.length.toLocaleString()})
        </button>
      </div>

      {/* TAB 1: MA TRẬN THIẾT BỊ KIOSK (DEVICE MATRIX GRID) */}
      {activeTab === 'OVERVIEW' && (
        <div className="space-y-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3.5">
            {analytics.kioskRankings.map((kiosk) => {
              const isHigh = kiosk.count >= 40;
              const isModerate = kiosk.count >= 15 && kiosk.count < 40;
              const isSelf = kiosk.config.toLowerCase().includes('tự thực hiện');

              return (
                <div 
                  key={kiosk.machine} 
                  className={`p-4 rounded-2xl border transition duration-150 hover:shadow-md flex flex-col justify-between ${
                    isFullscreen ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'
                  }`}
                >
                  <div>
                    {/* Header Card */}
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <span className="font-mono font-bold text-sm tracking-tight text-slate-900 block">
                          {kiosk.machine}
                        </span>
                        <span className="text-[11px] text-slate-500 truncate block max-w-[140px]" title={kiosk.area}>
                          {kiosk.area}
                        </span>
                      </div>

                      {/* Trạng thái hiện đại (Dùng HTML indicator, không dùng emoji) */}
                      {isHigh ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                          Tải cao
                        </span>
                      ) : isModerate ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                          <span className="w-1.5 h-1.5 rounded-full bg-blue-500"></span>
                          Ổn định
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">
                          <span className="w-1.5 h-1.5 rounded-full bg-amber-500"></span>
                          Tải thấp
                        </span>
                      )}
                    </div>

                    {/* Số lượng thao tác */}
                    <div className="mt-3 flex items-baseline justify-between">
                      <span className="text-2xl font-black font-mono text-teal-600">
                        {kiosk.count}
                      </span>
                      <span className="text-[11px] text-slate-400">lượt hôm nay</span>
                    </div>

                    {/* Chi tiết Checkin vs Viện phí */}
                    <div className="mt-2 pt-2 border-t border-slate-100 flex items-center justify-between text-[10px] text-slate-500">
                      <span>Checkin: <b>{kiosk.checkin}</b></span>
                      <span>Viện phí: <b>{kiosk.vienPhi}</b></span>
                      <span>Thuốc: <b>{kiosk.banThuoc}</b></span>
                    </div>
                  </div>

                  {/* Footer Card */}
                  <div className="mt-3 pt-2 border-t border-slate-100 flex items-center justify-between text-[10px]">
                    <span className={`px-1.5 py-0.5 rounded font-medium ${
                      isSelf ? 'bg-purple-50 text-purple-700 border border-purple-200' : 'bg-slate-100 text-slate-600'
                    }`}>
                      {kiosk.config}
                    </span>
                    <span className="text-slate-400 font-mono">
                      {kiosk.lastActiveTime.split(' ')[1] || kiosk.lastActiveTime}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* TAB 2: BẢNG XẾP HẠNG NHÂN SỰ CSKH (STAFF PRODUCTIVITY LEADERBOARD) */}
      {activeTab === 'STAFF' && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="p-4 border-b border-slate-100 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-900">Bảng Vinh Danh Năng Suất Nhân Sự CSKH Trực Máy</h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Thống kê số lượt hỗ trợ người bệnh tại các điểm Kiosk
              </p>
            </div>
            <span className="text-xs font-semibold px-2.5 py-1 bg-teal-50 text-teal-700 rounded-full border border-teal-200">
              {analytics.staffRankings.length} Nhân sự tham gia
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-600">
              <thead className="bg-slate-50 text-slate-700 uppercase font-semibold text-[11px] tracking-wider border-b border-slate-200">
                <tr>
                  <th className="py-3 px-4 w-14 text-center">Hạng</th>
                  <th className="py-3 px-4">Họ và Tên Nhân Viên</th>
                  <th className="py-3 px-4 text-right">Tổng Lượt Hỗ Trợ</th>
                  <th className="py-3 px-4 text-center">Lượt Check-in</th>
                  <th className="py-3 px-4 text-center">Lượt Viện Phí</th>
                  <th className="py-3 px-4 text-center">Lượt Bán Thuốc</th>
                  <th className="py-3 px-4">Phân Bổ Ca Trực</th>
                  <th className="py-3 px-4 text-center">Đánh Giá Hiệu Suất</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {analytics.staffRankings.map((staff, idx) => {
                  const isTop1 = idx === 0;
                  const isTop2 = idx === 1;
                  const isTop3 = idx === 2;

                  return (
                    <tr key={staff.name} className="hover:bg-slate-50/80 transition">
                      <td className="py-3 px-4 text-center">
                        {isTop1 ? (
                          <span className="w-6 h-6 rounded-full bg-amber-500 text-white font-bold inline-flex items-center justify-center text-xs shadow-xs">1</span>
                        ) : isTop2 ? (
                          <span className="w-6 h-6 rounded-full bg-slate-400 text-white font-bold inline-flex items-center justify-center text-xs shadow-xs">2</span>
                        ) : isTop3 ? (
                          <span className="w-6 h-6 rounded-full bg-amber-700 text-white font-bold inline-flex items-center justify-center text-xs shadow-xs">3</span>
                        ) : (
                          <span className="font-mono text-slate-400 font-semibold">{idx + 1}</span>
                        )}
                      </td>

                      <td className="py-3 px-4 font-bold text-slate-900">
                        {staff.name}
                      </td>

                      <td className="py-3 px-4 text-right font-mono font-extrabold text-teal-600 text-sm">
                        {staff.total.toLocaleString()}
                      </td>

                      <td className="py-3 px-4 text-center font-mono">
                        <span className="px-2 py-0.5 rounded bg-blue-50 text-blue-700 font-semibold text-[11px]">
                          {staff.checkin}
                        </span>
                      </td>

                      <td className="py-3 px-4 text-center font-mono">
                        <span className="px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 font-semibold text-[11px]">
                          {staff.vienPhi}
                        </span>
                      </td>

                      <td className="py-3 px-4 text-center font-mono">
                        <span className="px-2 py-0.5 rounded bg-purple-50 text-purple-700 font-semibold text-[11px]">
                          {staff.banThuoc}
                        </span>
                      </td>

                      <td className="py-3 px-4 text-slate-500 text-[11px]">
                        <span>Sáng: <b>{staff.morningCount}</b></span> | <span>Chiều: <b>{staff.afternoonCount}</b></span>
                      </td>

                      <td className="py-3 px-4 text-center">
                        {staff.total >= 100 ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800">
                            <Sparkles className="w-3 h-3 text-emerald-600" /> Xuất sắc
                          </span>
                        ) : staff.total >= 40 ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-blue-100 text-blue-800">
                            Tích cực
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-700">
                            Đạt chuẩn
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* TAB 3: BẢNG DỮ LIỆU DRILL-DOWN CHI TIẾT */}
      {activeTab === 'DETAILS' && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-600">
              <thead className="bg-slate-50 text-slate-700 uppercase font-semibold text-[11px] tracking-wider border-b border-slate-200">
                <tr>
                  <th className="py-3 px-3 w-12 text-center">STT</th>
                  <th className="py-3 px-3">Thời Gian</th>
                  <th className="py-3 px-3">PID</th>
                  <th className="py-3 px-4">Họ và Tên Bệnh Nhân</th>
                  <th className="py-3 px-3">Tên Kiosk</th>
                  <th className="py-3 px-3">Khu Vực</th>
                  <th className="py-3 px-3 text-center">Dịch Vụ</th>
                  <th className="py-3 px-3 text-center">Hình Thức TT</th>
                  <th className="py-3 px-4">Nhân Viên CSKH Hỗ Trợ</th>
                  <th className="py-3 px-3">Buổi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginatedItems.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="py-12 text-center text-slate-400">
                      Không tìm thấy bản ghi nào phù hợp với bộ lọc hiện tại.
                    </td>
                  </tr>
                ) : (
                  paginatedItems.map((row, i) => {
                    const stt = (currentPage - 1) * pageSize + i + 1;
                    const isSelf = String(row['Họ tên user'] || '').toLowerCase().includes('tự thực hiện');

                    return (
                      <tr key={row['STT'] || stt} className="hover:bg-slate-50/80 transition">
                        <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">{stt}</td>
                        <td className="py-2.5 px-3 font-mono text-slate-700">{row['Ngày giờ thao tác']}</td>
                        <td className="py-2.5 px-3 font-mono font-bold text-slate-900">{row['PID']}</td>
                        <td className="py-2.5 px-4 font-medium text-slate-800">{row['Họ tên']}</td>
                        <td className="py-2.5 px-3 font-mono font-semibold text-teal-700">{row['Tên KIOS']}</td>
                        <td className="py-2.5 px-3 text-slate-600 max-w-[140px] truncate" title={row['KHU VỰC']}>
                          {row['KHU VỰC']}
                        </td>

                        {/* Cột Dịch Vụ */}
                        <td className="py-2.5 px-3 text-center">
                          <div className="flex items-center justify-center gap-1">
                            {row['Checkin'] && (
                              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-100 text-blue-800">
                                Checkin
                              </span>
                            )}
                            {row['Thu viện phí'] && (
                              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">
                                Viện phí
                              </span>
                            )}
                            {row['Bán thuốc'] && (
                              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-purple-100 text-purple-800">
                                Thuốc
                              </span>
                            )}
                          </div>
                        </td>

                        {/* Hình thức TT */}
                        <td className="py-2.5 px-3 text-center">
                          {row['HÌNH THỨC TT'] ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-mono font-bold bg-slate-100 text-slate-800 border border-slate-200">
                              {row['HÌNH THỨC TT']}
                            </span>
                          ) : (
                            <span className="text-slate-300">--</span>
                          )}
                        </td>

                        {/* Nhân sự CSKH */}
                        <td className="py-2.5 px-4 text-slate-700 font-medium">
                          {isSelf ? (
                            <span className="text-purple-600 text-[11px] font-semibold">Tự thực hiện</span>
                          ) : (
                            row['Họ tên user']
                          )}
                        </td>

                        <td className="py-2.5 px-3 text-slate-500 text-[11px]">{row['BUỔI']}</td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Thanh Phân Trang */}
          <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
            <span>
              Hiển thị <b>{analytics.filteredRows.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}</b> - <b>{Math.min(currentPage * pageSize, analytics.filteredRows.length)}</b> / <b>{analytics.filteredRows.length.toLocaleString()}</b> lượt
            </span>

            <div className="flex items-center gap-1">
              <button
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="px-2 font-medium">Trang {currentPage} / {totalPages}</span>
              <button
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL CẤU HÌNH GOOGLE SHEET */}
      {showConfigModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-lg w-full p-6 text-slate-800 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <Settings className="w-5 h-5 text-teal-600" />
                <h3 className="font-bold text-base">Cấu Hình Kết Nối Google Sheet</h3>
              </div>
              <button 
                onClick={() => setShowConfigModal(false)}
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Spreadsheet ID</label>
                <input
                  type="text"
                  value={sheetConfig.spreadsheetId}
                  onChange={(e) => setSheetConfig({ ...sheetConfig, spreadsheetId: e.target.value.trim() })}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-xs focus:outline-none focus:ring-1 focus:ring-teal-500"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">GID Dữ Liệu Tổng</label>
                  <input
                    type="text"
                    value={sheetConfig.gidDuLieuTong}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidDuLieuTong: e.target.value.trim() })}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-xs focus:outline-none focus:ring-1 focus:ring-teal-500"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">GID Vị Trí Máy</label>
                  <input
                    type="text"
                    value={sheetConfig.gidViTriMay}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidViTriMay: e.target.value.trim() })}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-xs focus:outline-none focus:ring-1 focus:ring-teal-500"
                  />
                </div>
              </div>

              <div className="p-3 bg-blue-50/70 border border-blue-200 rounded-xl text-blue-900 space-y-1">
                <p className="font-bold">Lưu ý để kết nối tự động thành công:</p>
                <p className="text-[11px] leading-relaxed">
                  Trong Google Sheet, nhấn nút <b>Chia sẻ</b> &gt; Mục <i>Quyền truy cập chung</i> chọn <b>"Bất kỳ ai có đường liên kết"</b> với vai trò <b>Người xem</b>.
                </p>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                onClick={() => setShowConfigModal(false)}
                className="px-4 py-2 border border-slate-300 text-slate-700 rounded-xl text-xs font-medium hover:bg-slate-50 cursor-pointer"
              >
                Đóng
              </button>
              <button
                onClick={() => {
                  setShowConfigModal(false);
                  handleSyncGoogleSheet();
                }}
                className="px-5 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-xl text-xs font-semibold cursor-pointer shadow-xs"
              >
                Lưu &amp; Đồng Bộ Ngay
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
