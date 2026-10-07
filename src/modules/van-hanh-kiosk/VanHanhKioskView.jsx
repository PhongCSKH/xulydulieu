import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Monitor, CreditCard, Clock, RefreshCw, Download, 
  Search, Building2, QrCode, FileText, Pill, 
  Settings, X, Sparkles, AlertCircle, ChevronLeft, 
  ChevronRight, Layers, Activity, BarChart3, TrendingUp
} from 'lucide-react';

import {
  DEFAULT_SHEET_CONFIG,
  DEFAULT_KIOSK_LOCATIONS,
  TARGET_AREAS,
  aggregateKioskData,
  syncAllKioskSheets,
  exportKioskReportToExcel,
  normalizeKioskName
} from './engine';

export default function VanHanhKioskView() {
  const [rawRows, setRawRows] = useState([]);
  const [locationsMap, setLocationsMap] = useState({});
  const [targetsInfo, setTargetsInfo] = useState({
    totalDailyTarget: 3150,
    totalTargetCLS: 2850,
    totalTargetThuoc: 300
  });

  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState(null);
  const [syncError, setSyncError] = useState(null);

  // Tab điều hướng chính
  const [activeTab, setActiveTab] = useState('OVERVIEW');

  // Bộ lọc
  const [dateFilter, setDateFilter] = useState('ALL');
  const [selectedShift, setSelectedShift] = useState('ALL');
  const [selectedArea, setSelectedArea] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');

  // Phân trang
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 12;

  // Modal cấu hình Sheet
  const [showConfigModal, setShowConfigModal] = useState(false);
  const [sheetConfig, setSheetConfig] = useState(DEFAULT_SHEET_CONFIG);

  // Hàm đồng bộ trực tiếp từ Google Sheet
  const loadDataFromGoogleSheet = async (cfg = sheetConfig) => {
    setIsSyncing(true);
    setSyncError(null);

    try {
      const res = await syncAllKioskSheets(cfg);
      if (res.dataRows && res.dataRows.length > 0) {
        setRawRows(res.dataRows);
        setLocationsMap(res.locationsMap);
        setTargetsInfo({
          totalDailyTarget: res.totalDailyTarget,
          totalTargetCLS: res.totalTargetCLS,
          totalTargetThuoc: res.totalTargetThuoc
        });
        setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
        setSyncError(null);
      } else {
        throw new Error('Dữ liệu từ Google Sheet rỗng.');
      }
    } catch (err) {
      console.error(err);
      setSyncError(err.message || 'Chưa thể kết nối Google Sheet.');
    } finally {
      setIsSyncing(false);
    }
  };

  useEffect(() => {
    const map = {};
    DEFAULT_KIOSK_LOCATIONS.forEach(loc => {
      map[normalizeKioskName(loc.machine)] = loc;
    });
    setLocationsMap(map);

    loadDataFromGoogleSheet();
  }, []);

  // Danh sách các ngày thực tế có trong Google Sheet
  const availableDates = useMemo(() => {
    const s = new Set();
    rawRows.forEach(r => {
      const d = r['NGÀY'] || (r['Ngày giờ thao tác'] || '').split(' ')[0];
      if (d && d.length >= 8) s.add(d);
    });
    return Array.from(s).sort().reverse();
  }, [rawRows]);

  // Phân tích dữ liệu
  const analytics = useMemo(() => {
    return aggregateKioskData(rawRows, locationsMap, targetsInfo, {
      dateFilter,
      selectedShift,
      selectedArea,
      searchTerm
    });
  }, [rawRows, locationsMap, targetsInfo, dateFilter, selectedShift, selectedArea, searchTerm]);

  const totalPages = Math.ceil(analytics.filteredRows.length / pageSize) || 1;
  const paginatedItems = analytics.filteredRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  return (
    <div className="min-h-screen bg-slate-50/70 text-slate-900 font-sans -m-6 p-6 space-y-6">
      {/* 1. TOP HEADER ĐIỀU HÀNH - GIAO DIỆN SÁNG CAO CẤP */}
      <div className="bg-white p-5 rounded-3xl border border-slate-200 shadow-sm shadow-slate-100">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          {/* Logo & Tiêu đề */}
          <div className="flex items-center gap-3.5">
            <div className="w-12 h-12 rounded-2xl bg-gradient-to-tr from-teal-500 to-emerald-500 flex items-center justify-center text-white font-black shadow-md shadow-teal-500/25">
              <Monitor className="w-6 h-6 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[11px] font-extrabold font-mono tracking-widest text-teal-700 bg-teal-50 border border-teal-200 px-2.5 py-0.5 rounded-full uppercase">
                  KIOSK INTELLIGENCE
                </span>
                <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-50 text-emerald-800 border border-emerald-200">
                  <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                  LIVE GOOGLE SHEET
                </span>
              </div>
              <h1 className="text-xl lg:text-2xl font-black text-slate-900 tracking-tight mt-1">
                Trung Tâm Điều Hành Hoạt Động Kiosk
              </h1>
            </div>
          </div>

          {/* Nút hành động */}
          <div className="flex flex-wrap items-center gap-2.5">
            {lastSyncTime && (
              <div className="px-3 py-1.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-700 text-xs font-mono font-medium flex items-center gap-2">
                <Clock className="w-3.5 h-3.5 text-teal-600" />
                <span>Đã nạp: <b className="text-slate-900">{lastSyncTime}</b> ({rawRows.length.toLocaleString()} dòng)</span>
              </div>
            )}

            <button
              onClick={() => loadDataFromGoogleSheet()}
              disabled={isSyncing}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold text-white bg-teal-600 hover:bg-teal-700 shadow-sm shadow-teal-700/20 transition-all cursor-pointer disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
              <span>{isSyncing ? 'Đang kéo Sheet...' : 'Đồng Bộ Sheet'}</span>
            </button>

            <button
              onClick={() => setShowConfigModal(true)}
              className="p-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 text-xs font-medium transition cursor-pointer shadow-2xs"
              title="Cài đặt link Google Sheet"
            >
              <Settings className="w-4 h-4" />
            </button>

            <button
              onClick={() => exportKioskReportToExcel(analytics.filteredRows)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-100 text-slate-800 text-xs font-bold transition cursor-pointer shadow-2xs"
            >
              <Download className="w-3.5 h-3.5 text-slate-600" />
              <span>Xuất Excel</span>
            </button>
          </div>
        </div>

        {/* SUBNAV: Segmented Tab Bar & Bộ lọc */}
        <div className="mt-5 pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3">
          <div className="inline-flex p-1 rounded-2xl bg-slate-100 border border-slate-200">
            <button
              onClick={() => setActiveTab('OVERVIEW')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'OVERVIEW'
                  ? 'bg-white text-teal-800 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5 text-teal-600" />
              <span>Tổng Quan Điều Hành</span>
            </button>

            <button
              onClick={() => setActiveTab('SELF_SERVICE')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'SELF_SERVICE'
                  ? 'bg-purple-600 text-white shadow-xs'
                  : 'text-purple-700 hover:text-purple-900'
              }`}
            >
              <Sparkles className="w-3.5 h-3.5 text-purple-300" />
              <span>Máy Tự Thực Hiện ({analytics.summary.selfTotal})</span>
            </button>

            <button
              onClick={() => setActiveTab('DEPARTMENTS')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'DEPARTMENTS'
                  ? 'bg-white text-teal-800 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Building2 className="w-3.5 h-3.5 text-teal-600" />
              <span>Khoa Phòng &amp; Mục Tiêu</span>
            </button>

            <button
              onClick={() => setActiveTab('LOGS')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'LOGS'
                  ? 'bg-white text-teal-800 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Layers className="w-3.5 h-3.5 text-teal-600" />
              <span>Nhật Ký Dữ Liệu ({analytics.filteredRows.length.toLocaleString()})</span>
            </button>
          </div>

          {/* Lọc Ngày, Ca, Khu vực */}
          <div className="flex items-center gap-2 text-xs">
            {availableDates.length > 0 && (
              <select
                value={dateFilter}
                onChange={(e) => { setDateFilter(e.target.value); setCurrentPage(1); }}
                className="px-3 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-800 font-semibold focus:outline-none focus:ring-1 focus:ring-teal-500 cursor-pointer shadow-2xs"
              >
                <option value="ALL">Tất cả ngày ({rawRows.length.toLocaleString()} dòng)</option>
                {availableDates.map(d => (
                  <option key={d} value={d}>Ngày: {d}</option>
                ))}
              </select>
            )}

            <select
              value={selectedShift}
              onChange={(e) => { setSelectedShift(e.target.value); setCurrentPage(1); }}
              className="px-3 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-800 font-semibold focus:outline-none focus:ring-1 focus:ring-teal-500 cursor-pointer shadow-2xs"
            >
              <option value="ALL">Mọi ca trực</option>
              <option value="Buổi sáng">Ca Sáng (06h - 12h)</option>
              <option value="Buổi chiều">Ca Chiều (12h - 18h)</option>
            </select>

            <select
              value={selectedArea}
              onChange={(e) => { setSelectedArea(e.target.value); setCurrentPage(1); }}
              className="px-3 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-800 font-semibold focus:outline-none focus:ring-1 focus:ring-teal-500 max-w-[170px] truncate cursor-pointer shadow-2xs"
            >
              <option value="ALL">Tất cả khu vực</option>
              {TARGET_AREAS.map(a => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* CẢNH BÁO LỖI NẾU CÓ */}
      {syncError && (
        <div className="bg-amber-50 border border-amber-300 rounded-2xl p-4 text-amber-900 flex items-center justify-between text-xs shadow-xs">
          <div className="flex items-center gap-2.5">
            <AlertCircle className="w-5 h-5 text-amber-600 flex-shrink-0" />
            <span><b>Thông báo đồng bộ:</b> {syncError}</span>
          </div>
          <button 
            onClick={() => loadDataFromGoogleSheet()}
            className="px-3.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg cursor-pointer"
          >
            Thử lại ngay
          </button>
        </div>
      )}

      {/* 2. DẢI 4 THẺ METRICS SÁNG TƯƠI, ĐỘ TƯƠNG PHẢN CAO */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Metric 1: Tổng Lượt Thực Tế */}
        <div className="p-5 rounded-3xl border border-slate-200 bg-white shadow-xs hover:shadow-md transition-all duration-200 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-extrabold tracking-wider text-slate-500 uppercase">
                Tổng Lượt Thực Tế
              </span>
              <div className="w-8 h-8 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center border border-rose-200">
                <Activity className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-4xl font-black font-mono tracking-tight text-rose-600">
                {analytics.summary.totalOps.toLocaleString()}
              </span>
              <span className="text-xs font-bold text-rose-700 bg-rose-50 px-2 py-0.5 rounded-full border border-rose-200">
                Chính xác
              </span>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600">
            <span>Mục tiêu ngày: <b className="text-slate-900 font-mono">{targetsInfo.totalDailyTarget.toLocaleString()}</b></span>
            <span className="font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded-md border border-teal-200">
              Đạt {analytics.summary.overallRate}%
            </span>
          </div>
        </div>

        {/* Metric 2: Tiến Độ Mục Tiêu Ngày */}
        <div className="p-5 rounded-3xl border border-slate-200 bg-white shadow-xs hover:shadow-md transition-all duration-200 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-extrabold tracking-wider text-slate-500 uppercase">
                Tiến Độ Mục Tiêu Ngày
              </span>
              <span className="text-xs font-mono font-bold text-amber-800 bg-amber-50 px-2 py-0.5 rounded-full border border-amber-200">
                Chỉ tiêu: {targetsInfo.totalDailyTarget.toLocaleString()}
              </span>
            </div>
            <div className="mt-2 flex items-baseline justify-between">
              <span className="text-4xl font-black font-mono tracking-tight text-amber-600">
                {analytics.summary.overallRate}%
              </span>
              <span className="text-xs text-slate-600 font-mono font-bold">
                {analytics.summary.totalOps} / {targetsInfo.totalDailyTarget}
              </span>
            </div>
          </div>
          <div className="mt-4 space-y-1.5">
            <div className="w-full bg-slate-100 h-2.5 rounded-full overflow-hidden border border-slate-200">
              <div 
                className="bg-amber-500 h-full rounded-full transition-all duration-500" 
                style={{ width: `${Math.min(100, analytics.summary.overallRate)}%` }} 
              />
            </div>
            <div className="flex justify-between text-[11px] text-slate-600 font-medium">
              <span>Đã xong: <b className="text-slate-900">{analytics.summary.totalOps}</b></span>
              <span>Còn thiếu: <b className="text-amber-700">{Math.max(0, targetsInfo.totalDailyTarget - analytics.summary.totalOps)}</b></span>
            </div>
          </div>
        </div>

        {/* Metric 3: Đăng Ký & Thu Viện Phí */}
        <div className="p-5 rounded-3xl border border-slate-200 bg-white shadow-xs hover:shadow-md transition-all duration-200 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-extrabold tracking-wider text-slate-500 uppercase">
                Đăng Ký &amp; Thu Viện Phí
              </span>
              <div className="w-8 h-8 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center border border-blue-200">
                <FileText className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-4xl font-black font-mono tracking-tight text-blue-600">
                {analytics.summary.regCompletedRate}%
              </span>
              <span className="text-xs text-slate-600 font-mono font-bold">
                ({(analytics.summary.totalCheckin + analytics.summary.totalVienPhi).toLocaleString()} lượt)
              </span>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
            <span className="text-blue-700 font-bold font-mono bg-blue-50 px-2 py-0.5 rounded border border-blue-200">Checkin: {analytics.summary.totalCheckin}</span>
            <span className="text-amber-800 font-bold font-mono bg-amber-50 px-2 py-0.5 rounded border border-amber-200">Viện phí: {analytics.summary.totalVienPhi}</span>
          </div>
        </div>

        {/* Metric 4: Cơ Cấu Thanh Toán */}
        <div className="p-5 rounded-3xl border border-slate-200 bg-white shadow-xs hover:shadow-md transition-all duration-200 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-extrabold tracking-wider text-slate-500 uppercase">
                Cơ Cấu Thanh Toán
              </span>
              <div className="w-8 h-8 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center border border-purple-200">
                <QrCode className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline justify-between">
              <span className="text-4xl font-black font-mono tracking-tight text-purple-700">
                {analytics.summary.rateCK}%
              </span>
              <span className="text-xs text-purple-900 font-bold bg-purple-50 px-2 py-0.5 rounded-full border border-purple-200">
                Chuyển Khoản QR
              </span>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
            <span className="text-slate-600 font-medium">Quẹt thẻ POS: <b className="text-teal-700 font-mono font-bold">{analytics.summary.ratePOS}%</b></span>
            <span className="text-purple-700 font-bold font-mono">CK: {analytics.summary.countCK}</span>
          </div>
        </div>
      </div>

      {/* 3. PHÂN HỆ ĐỘC QUYỀN: MÁY TỰ THỰC HIỆN (SÁNG RỰC RỠ, NỔI BẬT) */}
      <div className="p-6 rounded-3xl border-2 border-purple-200 bg-gradient-to-r from-purple-50/80 via-white to-purple-50/40 shadow-xs space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-purple-200/80 pb-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-purple-600 text-white flex items-center justify-center shadow-md shadow-purple-600/30">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base font-black text-purple-950 uppercase tracking-wide">
                  PHÂN TÍCH CHUYÊN SÂU: MÁY CẤU HÌNH "TỰ THỰC HIỆN" (KIOS17 SẢNH CHÍNH)
                </h2>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase bg-purple-100 text-purple-800 border border-purple-300">
                  TỰ ĐỘNG HÓA
                </span>
              </div>
              <p className="text-xs text-purple-800 font-medium mt-0.5">
                Bóc tách riêng biệt lưu lượng phát sinh tại máy có cấu hình "Tự thực hiện" trong sheet Vị trí bố trí máy
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right bg-white px-4 py-2 rounded-2xl border border-purple-200 shadow-2xs">
              <span className="text-[11px] text-slate-500 font-bold block">Tỷ lệ tự phục vụ:</span>
              <span className="text-xl font-black font-mono text-purple-700">
                {analytics.summary.selfShareRate}% toàn viện
              </span>
            </div>
          </div>
        </div>

        {/* 4 Thẻ chỉ số máy Tự làm */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <div className="p-4 rounded-2xl border border-purple-200 bg-white shadow-2xs">
            <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500 block">
              Tổng Lượt Tự Làm
            </span>
            <span className="text-3xl font-black font-mono text-purple-700 mt-1 block">
              {analytics.summary.selfTotal.toLocaleString()}
            </span>
            <span className="text-[11px] text-slate-500 font-medium">Máy KIOS17 Sảnh chính</span>
          </div>

          <div className="p-4 rounded-2xl border border-blue-200 bg-white shadow-2xs">
            <span className="text-[11px] font-bold uppercase text-blue-700 block">Tự Check-in</span>
            <span className="text-3xl font-black font-mono text-blue-600 mt-1 block">
              {analytics.summary.selfCheckin}
            </span>
            <span className="text-[11px] text-slate-500 font-medium">
              {analytics.summary.selfTotal > 0 ? ((analytics.summary.selfCheckin / analytics.summary.selfTotal) * 100).toFixed(1) : 0}% trên máy tự làm
            </span>
          </div>

          <div className="p-4 rounded-2xl border border-amber-200 bg-white shadow-2xs">
            <span className="text-[11px] font-bold uppercase text-amber-700 block">Tự Nộp Viện Phí</span>
            <span className="text-3xl font-black font-mono text-amber-600 mt-1 block">
              {analytics.summary.selfVienPhi}
            </span>
            <span className="text-[11px] text-slate-500 font-medium">
              {analytics.summary.selfTotal > 0 ? ((analytics.summary.selfVienPhi / analytics.summary.selfTotal) * 100).toFixed(1) : 0}% trên máy tự làm
            </span>
          </div>

          <div className="p-4 rounded-2xl border border-emerald-200 bg-white shadow-2xs">
            <span className="text-[11px] font-bold uppercase text-emerald-700 block">Thanh Toán Tại Máy</span>
            <span className="text-lg font-black font-mono text-emerald-700 mt-1 block">
              CK {analytics.summary.selfRateCK}% - POS {analytics.summary.selfRatePOS}%
            </span>
            <span className="text-[11px] text-slate-500 font-medium">Quét mã QR trực tiếp</span>
          </div>
        </div>

        {/* Biểu đồ nhịp giờ máy tự làm */}
        <div className="p-4 rounded-2xl border border-purple-200 bg-white shadow-2xs">
          <div className="flex items-center justify-between mb-2 text-xs font-bold text-slate-700">
            <span className="text-purple-950 uppercase tracking-wide">
              Lưu lượng bệnh nhân tự thao tác theo khung giờ (05:00 - 18:00)
            </span>
            <span className="text-purple-700 font-mono bg-purple-50 px-2 py-0.5 rounded border border-purple-200">
              Đỉnh tự làm: 08:00 - 10:00
            </span>
          </div>

          <div className="h-28 flex items-end gap-1.5 pt-4 pb-1 border-b border-slate-100">
            {analytics.hourlySelfList.map((item) => {
              const maxSelf = Math.max(...analytics.hourlySelfList.map(s => s.total), 1);
              const pct = Math.round((item.total / maxSelf) * 100);

              return (
                <div key={item.hour} className="flex-1 flex flex-col items-center h-full justify-end group relative">
                  <div className="absolute -top-8 opacity-0 group-hover:opacity-100 transition bg-slate-900 text-white text-[10px] rounded px-2 py-1 shadow-lg pointer-events-none whitespace-nowrap z-20 font-mono font-bold">
                    {item.hour}: {item.total} lượt tự làm
                  </div>
                  <div 
                    className="w-full max-w-[22px] bg-purple-600 rounded-t-sm transition-all duration-300"
                    style={{ height: `${Math.max(6, pct)}%` }}
                  />
                  <span className="text-[10px] font-mono text-slate-500 mt-1 font-semibold">
                    {item.hour.split(':')[0]}h
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* 4. BIỂU ĐỒ PHÂN BỔ KHOA PHÒNG & NHẬT KÝ CHI TIẾT */}
      {activeTab === 'OVERVIEW' && (
        <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-base font-extrabold text-slate-900 tracking-tight">
                Phân Bổ Tải Trọng Toàn Viện (Check-in + Viện Phí + Bán Thuốc)
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Số liệu thực tế chính xác 100% từ từng dòng Google Sheet
              </p>
            </div>

            <div className="flex items-center gap-4 text-xs font-bold">
              <span className="flex items-center gap-1.5 text-blue-700 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                <span className="w-2.5 h-2.5 rounded-sm bg-blue-600"></span> Checkin ({analytics.summary.totalCheckin})
              </span>
              <span className="flex items-center gap-1.5 text-amber-800 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                <span className="w-2.5 h-2.5 rounded-sm bg-amber-500"></span> Thu viện phí ({analytics.summary.totalVienPhi})
              </span>
              <span className="flex items-center gap-1.5 text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                <span className="w-2.5 h-2.5 rounded-sm bg-emerald-600"></span> Bán thuốc ({analytics.summary.totalBanThuoc})
              </span>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {analytics.areaBreakdown.map((item, idx) => {
              const total = item.checkin + item.vienPhi + item.banThuoc;
              const pCheckin = total > 0 ? (item.checkin / total) * 100 : 0;
              const pVienPhi = total > 0 ? (item.vienPhi / total) * 100 : 0;
              const pBanThuoc = total > 0 ? (item.banThuoc / total) * 100 : 0;

              return (
                <div key={item.area} className="p-4 rounded-2xl border border-slate-200 bg-slate-50/40 hover:bg-white hover:border-teal-400 hover:shadow-xs transition-all duration-200">
                  <div className="flex items-center justify-between text-xs mb-2">
                    <div className="flex items-center gap-2">
                      <span className="font-mono text-slate-400 text-xs font-bold">#{idx + 1}</span>
                      <span className="font-bold text-slate-900">{item.area}</span>
                    </div>
                    <span className="font-mono font-black text-sm text-teal-700">
                      {total.toLocaleString()} <span className="text-[11px] text-slate-500 font-normal">lượt</span>
                    </span>
                  </div>

                  <div className="h-3.5 w-full bg-slate-200 rounded-full overflow-hidden flex border border-slate-200">
                    <div 
                      className="bg-blue-600 h-full transition-all" 
                      style={{ width: `${pCheckin}%` }} 
                      title={`Checkin: ${item.checkin}`}
                    />
                    <div 
                      className="bg-amber-500 h-full transition-all" 
                      style={{ width: `${pVienPhi}%` }} 
                      title={`Viện phí: ${item.vienPhi}`}
                    />
                    <div 
                      className="bg-emerald-600 h-full transition-all" 
                      style={{ width: `${pBanThuoc}%` }} 
                      title={`Bán thuốc: ${item.banThuoc}`}
                    />
                  </div>

                  <div className="mt-2.5 flex items-center justify-between text-xs text-slate-600 font-mono font-semibold">
                    <span>Checkin: <b className="text-blue-700">{item.checkin}</b></span>
                    <span>Viện phí: <b className="text-amber-700">{item.vienPhi}</b></span>
                    <span>Thuốc: <b className="text-emerald-700">{item.banThuoc}</b></span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* 5. TAB LOGS: NHẬT KÝ CHI TIẾT SÁNG RÕ */}
      {activeTab === 'LOGS' && (
        <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
            <div>
              <h3 className="text-sm font-bold text-slate-900">Nhật Ký Dữ Liệu Đồng Bộ Trực Tiếp Từ Sheet</h3>
              <p className="text-xs text-slate-500">Khớp 100% từng dòng trên Google Sheet của bạn ({rawRows.length.toLocaleString()} dòng)</p>
            </div>
            <div className="relative w-full sm:w-64">
              <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                placeholder="PID, Tên, Kiosk..."
                value={searchTerm}
                onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                className="w-full text-xs pl-8 pr-3 py-1.5 rounded-xl border border-slate-300 bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-medium"
              />
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700">
              <thead className="border-b border-slate-200 bg-slate-100/90 text-slate-800 text-[11px] font-bold uppercase tracking-wider">
                <tr>
                  <th className="py-2.5 px-3 w-12 text-center">STT</th>
                  <th className="py-2.5 px-3">Thời Gian</th>
                  <th className="py-2.5 px-3">PID</th>
                  <th className="py-2.5 px-4">Bệnh Nhân</th>
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
                    <tr key={r['STT'] || stt} className="hover:bg-slate-50 transition">
                      <td className="py-2.5 px-3 text-center text-slate-400 font-mono">{stt}</td>
                      <td className="py-2.5 px-3 font-mono text-slate-600 font-medium">{r['Ngày giờ thao tác']}</td>
                      <td className="py-2.5 px-3 font-mono font-bold text-slate-900">{r['PID']}</td>
                      <td className="py-2.5 px-4 font-bold text-slate-800">{r['Họ tên']}</td>
                      <td className="py-2.5 px-3 font-mono text-teal-700 font-bold">{r['Tên KIOS']}</td>
                      <td className="py-2.5 px-3 text-slate-700 font-medium">{r['KHU VỰC']}</td>
                      <td className="py-2.5 px-3 text-center">
                        <div className="flex items-center justify-center gap-1 font-bold">
                          {r['Checkin'] && <span className="px-1.5 py-0.5 rounded text-[10px] bg-blue-100 text-blue-800 border border-blue-200">Checkin</span>}
                          {r['Thu viện phí'] && <span className="px-1.5 py-0.5 rounded text-[10px] bg-amber-100 text-amber-900 border border-amber-200">Viện phí</span>}
                          {r['Bán thuốc'] && <span className="px-1.5 py-0.5 rounded text-[10px] bg-emerald-100 text-emerald-900 border border-emerald-200">Thuốc</span>}
                        </div>
                      </td>
                      <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-800">
                        {r['HÌNH THỨC TT'] || '--'}
                      </td>
                      <td className="py-2.5 px-4">
                        {isSelf ? (
                          <span className="font-bold text-purple-800 bg-purple-100 px-2 py-0.5 rounded text-[11px] border border-purple-200">
                            Tự thực hiện
                          </span>
                        ) : (
                          <span className="text-slate-700 font-medium">{r['Họ tên user']}</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="p-3 border-t border-slate-200 flex items-center justify-between text-xs text-slate-600">
            <span>
              Hiển thị <b>{(currentPage - 1) * pageSize + 1}</b> - <b>{Math.min(currentPage * pageSize, analytics.filteredRows.length)}</b> / <b>{analytics.filteredRows.length.toLocaleString()}</b>
            </span>
            <div className="flex items-center gap-1">
              <button
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 cursor-pointer shadow-2xs"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
              </button>
              <span className="px-2 font-bold font-mono">{currentPage} / {totalPages}</span>
              <button
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 cursor-pointer shadow-2xs"
              >
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL CẤU HÌNH GOOGLE SHEET */}
      {showConfigModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="rounded-3xl border border-slate-200 bg-white text-slate-800 shadow-2xl max-w-lg w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3 border-slate-100">
              <h3 className="font-extrabold text-sm text-slate-900">Cấu Hình Google Sheet Tự Động</h3>
              <button onClick={() => setShowConfigModal(false)} className="text-slate-400 hover:text-slate-600 cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-bold mb-1 text-slate-700">Spreadsheet ID</label>
                <input
                  type="text"
                  value={sheetConfig.spreadsheetId}
                  onChange={(e) => setSheetConfig({ ...sheetConfig, spreadsheetId: e.target.value.trim() })}
                  className="w-full p-2.5 border border-slate-300 rounded-xl font-mono text-xs bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-semibold"
                />
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block font-bold mb-1 text-slate-700">GID Dữ Liệu</label>
                  <input
                    type="text"
                    value={sheetConfig.gidDuLieuTong}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidDuLieuTong: e.target.value.trim() })}
                    className="w-full p-2 border border-slate-300 rounded-xl font-mono text-xs bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-semibold"
                  />
                </div>
                <div>
                  <label className="block font-bold mb-1 text-slate-700">GID Vị Trí Máy</label>
                  <input
                    type="text"
                    value={sheetConfig.gidViTriMay}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidViTriMay: e.target.value.trim() })}
                    className="w-full p-2 border border-slate-300 rounded-xl font-mono text-xs bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-semibold"
                  />
                </div>
                <div>
                  <label className="block font-bold mb-1 text-slate-700">GID Mục Tiêu</label>
                  <input
                    type="text"
                    value={sheetConfig.gidMucTieu}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidMucTieu: e.target.value.trim() })}
                    className="w-full p-2 border border-slate-300 rounded-xl font-mono text-xs bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-semibold"
                  />
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                onClick={() => setShowConfigModal(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-100 cursor-pointer"
              >
                Đóng
              </button>
              <button
                onClick={() => {
                  setShowConfigModal(false);
                  loadDataFromGoogleSheet(sheetConfig);
                }}
                className="px-5 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-xl text-xs font-bold cursor-pointer shadow-sm"
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
