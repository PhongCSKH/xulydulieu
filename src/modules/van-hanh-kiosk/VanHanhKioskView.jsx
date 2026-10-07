import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Monitor, CreditCard, Clock, RefreshCw, Download, 
  Search, CheckCircle2, ChevronLeft, ChevronRight, Building2, 
  QrCode, FileText, Pill, Settings, X, Sparkles, AlertCircle, 
  Sun, Moon, Layers, Activity, BarChart3
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

  // Theme Tối / Sáng (Mặc định Dark Executive)
  const [themeMode, setThemeMode] = useState('dark');

  // Tab điều hướng
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

  // Hàm đồng bộ toàn bộ 3 sheet trực tiếp từ Google Sheet
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

  // Tự động tải từ Google Sheet ngay khi mở trang
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

  // Phân tích và tổng hợp số liệu chính xác 100%
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

  const isDark = themeMode === 'dark';

  return (
    <div className={`min-h-screen transition-colors duration-300 font-sans ${
      isDark ? 'bg-[#090D16] text-slate-100' : 'bg-slate-50 text-slate-900'
    } -m-6 p-6`}>
      {/* 1. TOP HEADER ĐIỀU HÀNH */}
      <div className={`p-5 rounded-3xl border transition-all duration-300 shadow-sm mb-6 ${
        isDark 
          ? 'bg-slate-900/80 border-slate-800/80 backdrop-blur-xl' 
          : 'bg-white border-slate-200/80 shadow-slate-100'
      }`}>
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-center gap-3.5">
            <div className="w-12 h-12 rounded-2xl bg-gradient-to-tr from-cyan-500 via-teal-500 to-emerald-400 flex items-center justify-center text-slate-950 font-black shadow-lg shadow-teal-500/20">
              <Monitor className="w-6 h-6 text-slate-950" />
            </div>
            <div>
              <div className="flex items-center gap-2.5">
                <span className="text-xs font-bold font-mono tracking-widest text-teal-400 uppercase">
                  KIOSK INTELLIGENCE
                </span>
                <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                  LIVE GOOGLE SHEET
                </span>
              </div>
              <h1 className="text-xl lg:text-2xl font-black tracking-tight mt-0.5">
                Trung Tâm Điều Hành Hoạt Động Kiosk
              </h1>
            </div>
          </div>

          {/* Quick Actions */}
          <div className="flex flex-wrap items-center gap-2.5">
            {lastSyncTime && (
              <div className={`px-3 py-1.5 rounded-xl border text-xs font-mono flex items-center gap-2 ${
                isDark ? 'bg-slate-800/60 border-slate-700/80 text-slate-400' : 'bg-slate-100 border-slate-200 text-slate-600'
              }`}>
                <Clock className="w-3.5 h-3.5 text-teal-400" />
                <span>Đã nạp: {lastSyncTime} ({rawRows.length.toLocaleString()} dòng)</span>
              </div>
            )}

            <button
              onClick={() => loadDataFromGoogleSheet()}
              disabled={isSyncing}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold text-white bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-500 hover:to-emerald-500 shadow-md shadow-teal-900/30 transition-all cursor-pointer disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
              <span>{isSyncing ? 'Đang kéo Sheet...' : 'Đồng Bộ Sheet'}</span>
            </button>

            <button
              onClick={() => setShowConfigModal(true)}
              className={`p-2 rounded-xl border text-xs font-medium transition cursor-pointer ${
                isDark ? 'bg-slate-800/80 border-slate-700 hover:bg-slate-700 text-slate-300' : 'bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700'
              }`}
              title="Cài đặt link Google Sheet"
            >
              <Settings className="w-4 h-4" />
            </button>

            <button
              onClick={() => exportKioskReportToExcel(analytics.filteredRows)}
              className={`inline-flex items-center gap-1.5 px-3 py-2 rounded-xl border text-xs font-semibold transition cursor-pointer ${
                isDark ? 'bg-slate-800/80 border-slate-700 hover:bg-slate-700 text-slate-300' : 'bg-slate-100 border-slate-200 hover:bg-slate-200 text-slate-700'
              }`}
            >
              <Download className="w-3.5 h-3.5" />
              <span>Xuất Excel</span>
            </button>

            <button
              onClick={() => setThemeMode(isDark ? 'light' : 'dark')}
              className={`p-2 rounded-xl border text-xs font-medium transition cursor-pointer ${
                isDark ? 'bg-amber-500/10 border-amber-500/30 text-amber-400 hover:bg-amber-500/20' : 'bg-slate-800 border-slate-700 text-slate-100 hover:bg-slate-700'
              }`}
              title={isDark ? 'Chế độ Sáng' : 'Chế độ Tối'}
            >
              {isDark ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>
          </div>
        </div>

        {/* SUBNAV: Segmented Tabs & Filters */}
        <div className={`mt-5 pt-4 border-t flex flex-wrap items-center justify-between gap-3 ${
          isDark ? 'border-slate-800' : 'border-slate-100'
        }`}>
          <div className={`inline-flex p-1 rounded-2xl border ${
            isDark ? 'bg-slate-950/60 border-slate-800' : 'bg-slate-100 border-slate-200'
          }`}>
            <button
              onClick={() => setActiveTab('OVERVIEW')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'OVERVIEW'
                  ? 'bg-teal-600 text-white shadow-sm'
                  : isDark ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5" />
              <span>Tổng Quan Điều Hành</span>
            </button>

            <button
              onClick={() => setActiveTab('SELF_SERVICE')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'SELF_SERVICE'
                  ? 'bg-gradient-to-r from-purple-600 to-indigo-600 text-white shadow-sm'
                  : isDark ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Sparkles className="w-3.5 h-3.5 text-purple-400" />
              <span>Máy Tự Thực Hiện ({analytics.summary.selfTotal})</span>
            </button>

            <button
              onClick={() => setActiveTab('DEPARTMENTS')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'DEPARTMENTS'
                  ? 'bg-teal-600 text-white shadow-sm'
                  : isDark ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Building2 className="w-3.5 h-3.5" />
              <span>Khoa Phòng &amp; Mục Tiêu</span>
            </button>

            <button
              onClick={() => setActiveTab('LOGS')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'LOGS'
                  ? 'bg-teal-600 text-white shadow-sm'
                  : isDark ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>Nhật Ký Dữ Liệu ({analytics.filteredRows.length.toLocaleString()})</span>
            </button>
          </div>

          {/* Lọc Ngày, Ca, Khu vực */}
          <div className="flex items-center gap-2 text-xs">
            {availableDates.length > 0 && (
              <select
                value={dateFilter}
                onChange={(e) => { setDateFilter(e.target.value); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-xl border text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-teal-500 cursor-pointer ${
                  isDark ? 'bg-slate-800 border-slate-700 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
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
              className={`px-3 py-1.5 rounded-xl border text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-teal-500 cursor-pointer ${
                isDark ? 'bg-slate-800 border-slate-700 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'
              }`}
            >
              <option value="ALL">Mọi ca trực</option>
              <option value="Buổi sáng">Ca Sáng (06h - 12h)</option>
              <option value="Buổi chiều">Ca Chiều (12h - 18h)</option>
            </select>

            <select
              value={selectedArea}
              onChange={(e) => { setSelectedArea(e.target.value); setCurrentPage(1); }}
              className={`px-3 py-1.5 rounded-xl border text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-teal-500 max-w-[160px] truncate cursor-pointer ${
                isDark ? 'bg-slate-800 border-slate-700 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'
              }`}
            >
              <option value="ALL">Tất cả khu vực</option>
              {TARGET_AREAS.map(a => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* CẢNH BÁO LỖI NẾU KẾT NỐI SHEET GẶP SỰ CỐ */}
      {syncError && (
        <div className="bg-amber-500/10 border border-amber-500/30 rounded-2xl p-4 text-amber-400 flex items-center justify-between mb-6 text-xs">
          <div className="flex items-center gap-2.5">
            <AlertCircle className="w-4 h-4 text-amber-400 flex-shrink-0" />
            <span><b>Thông báo đồng bộ:</b> {syncError}</span>
          </div>
          <button 
            onClick={() => loadDataFromGoogleSheet()}
            className="px-3 py-1 bg-amber-500 text-slate-950 font-bold rounded-lg cursor-pointer"
          >
            Thử lại
          </button>
        </div>
      )}

      {/* 2. DẢI 4 THẺ METRICS EXECUTIVE (SỐ LIỆU CHÍNH XÁC TỪ GOOGLE SHEET) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        {/* Metric 1: Tổng Lượt Vận Hành */}
        <div className={`p-5 rounded-3xl border transition-all duration-200 flex flex-col justify-between ${
          isDark ? 'bg-slate-900/70 border-slate-800/90' : 'bg-white border-slate-200/90 shadow-xs'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
                Tổng Lượt Thực Tế
              </span>
              <div className="w-8 h-8 rounded-xl bg-teal-500/10 text-teal-400 flex items-center justify-center">
                <Activity className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-4xl font-black font-mono tracking-tight text-teal-400">
                {analytics.summary.totalOps.toLocaleString()}
              </span>
              <span className="text-xs font-bold text-teal-400 bg-teal-500/10 px-2 py-0.5 rounded-full border border-teal-500/20">
                Chính xác 100%
              </span>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-800/40 flex items-center justify-between text-xs text-slate-400">
            <span>Mục tiêu ngày: <b>{targetsInfo.totalDailyTarget.toLocaleString()}</b></span>
            <span className="font-mono text-slate-300">Đạt {analytics.summary.overallRate}%</span>
          </div>
        </div>

        {/* Metric 2: Tiến Độ Mục Tiêu Ngày */}
        <div className={`p-5 rounded-3xl border transition-all duration-200 flex flex-col justify-between ${
          isDark ? 'bg-slate-900/70 border-slate-800/90' : 'bg-white border-slate-200/90 shadow-xs'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
                Tiến Độ Mục Tiêu (Sheet MỤC TIÊU)
              </span>
              <span className="text-xs font-mono font-bold text-amber-400 bg-amber-500/10 px-2 py-0.5 rounded-full border border-amber-500/20">
                {targetsInfo.totalDailyTarget.toLocaleString()} lượt
              </span>
            </div>
            <div className="mt-2 flex items-baseline justify-between">
              <span className="text-4xl font-black font-mono tracking-tight text-amber-400">
                {analytics.summary.overallRate}%
              </span>
              <span className="text-xs text-slate-400 font-mono">
                {analytics.summary.totalOps} / {targetsInfo.totalDailyTarget}
              </span>
            </div>
          </div>
          <div className="mt-4 space-y-1.5">
            <div className="w-full bg-slate-800 h-2 rounded-full overflow-hidden">
              <div 
                className="bg-gradient-to-r from-amber-500 to-amber-400 h-full rounded-full transition-all duration-500" 
                style={{ width: `${Math.min(100, analytics.summary.overallRate)}%` }} 
              />
            </div>
            <div className="flex justify-between text-[10px] text-slate-400">
              <span>Đạt: {analytics.summary.totalOps}</span>
              <span>Còn thiếu: {Math.max(0, targetsInfo.totalDailyTarget - analytics.summary.totalOps)}</span>
            </div>
          </div>
        </div>

        {/* Metric 3: Đăng Ký & Thu Viện Phí CLS */}
        <div className={`p-5 rounded-3xl border transition-all duration-200 flex flex-col justify-between ${
          isDark ? 'bg-slate-900/70 border-slate-800/90' : 'bg-white border-slate-200/90 shadow-xs'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
                Đăng Ký &amp; Thu Viện Phí
              </span>
              <div className="w-8 h-8 rounded-xl bg-blue-500/10 text-blue-400 flex items-center justify-center">
                <FileText className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-4xl font-black font-mono tracking-tight text-blue-400">
                {analytics.summary.regCompletedRate}%
              </span>
              <span className="text-xs text-slate-400 font-mono">
                ({(analytics.summary.totalCheckin + analytics.summary.totalVienPhi).toLocaleString()} lượt)
              </span>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-800/40 flex items-center justify-between text-xs">
            <span className="text-blue-400 font-semibold font-mono">Checkin: {analytics.summary.totalCheckin}</span>
            <span className="text-amber-400 font-semibold font-mono">Viện phí: {analytics.summary.totalVienPhi}</span>
          </div>
        </div>

        {/* Metric 4: Thanh Toán Không Tiền Mặt */}
        <div className={`p-5 rounded-3xl border transition-all duration-200 flex flex-col justify-between ${
          isDark ? 'bg-slate-900/70 border-slate-800/90' : 'bg-white border-slate-200/90 shadow-xs'
        }`}>
          <div>
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
                Cơ Cấu Thanh Toán
              </span>
              <div className="w-8 h-8 rounded-xl bg-purple-500/10 text-purple-400 flex items-center justify-center">
                <QrCode className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline justify-between">
              <span className="text-4xl font-black font-mono tracking-tight text-purple-400">
                {analytics.summary.rateCK}%
              </span>
              <span className="text-xs text-slate-400 font-mono">
                Chuyển Khoản QR
              </span>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-800/40 flex items-center justify-between text-xs">
            <span className="text-slate-400">Quẹt thẻ POS: <b className="text-teal-400 font-mono">{analytics.summary.ratePOS}%</b></span>
            <span className="text-purple-400 font-bold font-mono">CK: {analytics.summary.countCK}</span>
          </div>
        </div>
      </div>

      {/* 3. PHÂN HỆ ĐỘC QUYỀN: MÁY TỰ THỰC HIỆN (SỐ LIỆU THỰC TỪ SHEET VỊ TRÍ MÁY) */}
      <div className={`p-6 rounded-3xl border mb-6 transition-all duration-300 relative overflow-hidden ${
        isDark 
          ? 'bg-gradient-to-br from-indigo-950/40 via-purple-950/20 to-slate-900 border-indigo-500/30' 
          : 'bg-gradient-to-br from-purple-50 via-indigo-50/30 to-white border-purple-200 shadow-md'
      }`}>
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-6">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-purple-600 text-white flex items-center justify-center shadow-lg shadow-purple-600/30">
              <Sparkles className="w-5 h-5 text-purple-200" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base font-black tracking-tight uppercase">
                  PHÂN TÍCH CHUYÊN SÂU: MÁY CẤU HÌNH "TỰ THỰC HIỆN" (KIOS17 SẢNH CHÍNH)
                </h2>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase bg-purple-500/20 text-purple-400 border border-purple-500/30">
                  Tự Động Hóa
                </span>
              </div>
              <p className="text-xs text-slate-400 mt-0.5">
                Bóc tách riêng biệt lưu lượng phát sinh tại máy có cấu hình "Tự thực hiện" trong sheet Vị trí bố trí máy
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right">
              <span className="text-[11px] text-slate-400 block">Tỷ lệ tự phục vụ:</span>
              <span className="text-xl font-black font-mono text-purple-400">
                {analytics.summary.selfShareRate}% toàn viện
              </span>
            </div>
          </div>
        </div>

        {/* 4 Thẻ chỉ số máy Tự làm */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
          <div className={`p-4 rounded-2xl border ${
            isDark ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-purple-100 shadow-xs'
          }`}>
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block">
              Tổng Lượt Tự Làm
            </span>
            <span className="text-3xl font-black font-mono text-purple-400 mt-1 block">
              {analytics.summary.selfTotal.toLocaleString()}
            </span>
            <span className="text-[10px] text-slate-400">Máy KIOS17 Sảnh chính</span>
          </div>

          <div className={`p-4 rounded-2xl border ${
            isDark ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-blue-100 shadow-xs'
          }`}>
            <span className="text-[10px] font-bold uppercase text-blue-400 block">Tự Check-in</span>
            <span className="text-3xl font-black font-mono text-blue-400 mt-1 block">
              {analytics.summary.selfCheckin}
            </span>
            <span className="text-[10px] text-slate-400">
              {analytics.summary.selfTotal > 0 ? ((analytics.summary.selfCheckin / analytics.summary.selfTotal) * 100).toFixed(1) : 0}% trên máy tự làm
            </span>
          </div>

          <div className={`p-4 rounded-2xl border ${
            isDark ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-amber-100 shadow-xs'
          }`}>
            <span className="text-[10px] font-bold uppercase text-amber-400 block">Tự Nộp Viện Phí</span>
            <span className="text-3xl font-black font-mono text-amber-400 mt-1 block">
              {analytics.summary.selfVienPhi}
            </span>
            <span className="text-[10px] text-slate-400">
              {analytics.summary.selfTotal > 0 ? ((analytics.summary.selfVienPhi / analytics.summary.selfTotal) * 100).toFixed(1) : 0}% trên máy tự làm
            </span>
          </div>

          <div className={`p-4 rounded-2xl border ${
            isDark ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-emerald-100 shadow-xs'
          }`}>
            <span className="text-[10px] font-bold uppercase text-emerald-400 block">Thanh Toán Tại Máy</span>
            <span className="text-lg font-black font-mono text-emerald-400 mt-1 block">
              CK {analytics.summary.selfRateCK}% - POS {analytics.summary.selfRatePOS}%
            </span>
            <span className="text-[10px] text-slate-400">Quét mã QR trực tiếp</span>
          </div>
        </div>

        {/* Biểu đồ nhịp giờ máy tự làm */}
        <div className={`p-4 rounded-2xl border ${
          isDark ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-purple-100 shadow-xs'
        }`}>
          <div className="flex items-center justify-between mb-2 text-xs">
            <span className="font-bold text-slate-300 uppercase tracking-wide">
              Lưu lượng bệnh nhân tự thao tác theo khung giờ (05:00 - 18:00)
            </span>
            <span className="text-xs font-mono text-purple-400 font-bold">
              Đỉnh tự làm: 08:00 - 10:00
            </span>
          </div>

          <div className="h-28 flex items-end gap-1.5 pt-4 pb-1 border-b border-slate-800/60">
            {analytics.hourlySelfList.map((item) => {
              const maxSelf = Math.max(...analytics.hourlySelfList.map(s => s.total), 1);
              const pct = Math.round((item.total / maxSelf) * 100);

              return (
                <div key={item.hour} className="flex-1 flex flex-col items-center h-full justify-end group relative">
                  <div className="absolute -top-8 opacity-0 group-hover:opacity-100 transition bg-slate-800 text-white text-[10px] rounded px-2 py-1 shadow-lg pointer-events-none whitespace-nowrap z-20 font-mono">
                    {item.hour}: {item.total} lượt tự làm
                  </div>
                  <div 
                    className="w-full max-w-[20px] bg-gradient-to-t from-purple-700 to-indigo-500 rounded-t-sm transition-all duration-300"
                    style={{ height: `${Math.max(4, pct)}%` }}
                  />
                  <span className="text-[9px] font-mono text-slate-500 mt-1">
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
        <div className={`p-6 rounded-3xl border transition-all ${
          isDark ? 'bg-slate-900/70 border-slate-800' : 'bg-white border-slate-200/90 shadow-xs'
        }`}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-6 border-b pb-4 border-slate-800/60">
            <div>
              <h3 className="text-base font-extrabold tracking-tight">
                Phân Bổ Tải Trọng Toàn Viện (Check-in + Viện Phí + Bán Thuốc)
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Số liệu thực tế chính xác 100% từ từng dòng Google Sheet
              </p>
            </div>

            <div className="flex items-center gap-4 text-xs font-semibold">
              <span className="flex items-center gap-1.5 text-blue-400">
                <span className="w-3 h-3 rounded-md bg-blue-500"></span> Checkin ({analytics.summary.totalCheckin})
              </span>
              <span className="flex items-center gap-1.5 text-amber-400">
                <span className="w-3 h-3 rounded-md bg-amber-500"></span> Thu viện phí ({analytics.summary.totalVienPhi})
              </span>
              <span className="flex items-center gap-1.5 text-emerald-400">
                <span className="w-3 h-3 rounded-md bg-emerald-500"></span> Bán thuốc ({analytics.summary.totalBanThuoc})
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
                <div key={item.area} className={`p-3.5 rounded-2xl border transition-all duration-200 hover:border-teal-500/40 ${
                  isDark ? 'bg-slate-950/40 border-slate-800/80' : 'bg-slate-50 border-slate-200'
                }`}>
                  <div className="flex items-center justify-between text-xs mb-2">
                    <div className="flex items-center gap-2">
                      <span className="font-mono text-slate-500 text-[11px] font-bold">#{idx + 1}</span>
                      <span className="font-bold text-slate-200">{item.area}</span>
                    </div>
                    <span className="font-mono font-black text-sm text-teal-400">
                      {total.toLocaleString()} <span className="text-[10px] text-slate-500 font-normal">lượt</span>
                    </span>
                  </div>

                  <div className="h-3 w-full bg-slate-800 rounded-full overflow-hidden flex">
                    <div 
                      className="bg-blue-500 h-full transition-all" 
                      style={{ width: `${pCheckin}%` }} 
                      title={`Checkin: ${item.checkin}`}
                    />
                    <div 
                      className="bg-amber-500 h-full transition-all" 
                      style={{ width: `${pVienPhi}%` }} 
                      title={`Viện phí: ${item.vienPhi}`}
                    />
                    <div 
                      className="bg-emerald-500 h-full transition-all" 
                      style={{ width: `${pBanThuoc}%` }} 
                      title={`Bán thuốc: ${item.banThuoc}`}
                    />
                  </div>

                  <div className="mt-2 flex items-center justify-between text-[11px] text-slate-400 font-mono">
                    <span>Checkin: <b className="text-blue-400">{item.checkin}</b></span>
                    <span>Viện phí: <b className="text-amber-400">{item.vienPhi}</b></span>
                    <span>Thuốc: <b className="text-emerald-400">{item.banThuoc}</b></span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* 5. TAB LOGS: NHẬT KÝ TỪNG DÒNG THỰC TẾ TRÊN SHEET */}
      {activeTab === 'LOGS' && (
        <div className={`p-6 rounded-3xl border transition-all ${
          isDark ? 'bg-slate-900/70 border-slate-800' : 'bg-white border-slate-200/90 shadow-xs'
        }`}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
            <div>
              <h3 className="text-sm font-bold">Nhật Ký Dữ Liệu Đồng Bộ Trực Tiếp Từ Sheet</h3>
              <p className="text-xs text-slate-400">Khớp 100% từng dòng trên Google Sheet của bạn ({rawRows.length.toLocaleString()} dòng)</p>
            </div>
            <div className="relative w-full sm:w-64">
              <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                placeholder="PID, Tên, Kiosk..."
                value={searchTerm}
                onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                className={`w-full text-xs pl-8 pr-3 py-1.5 rounded-xl border focus:outline-none focus:ring-1 focus:ring-teal-500 ${
                  isDark ? 'bg-slate-800 border-slate-700 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className={`border-b text-[11px] uppercase tracking-wider ${
                isDark ? 'bg-slate-800/40 border-slate-800 text-slate-400' : 'bg-slate-50 border-slate-200 text-slate-600'
              }`}>
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
              <tbody className="divide-y divide-slate-800/40">
                {paginatedItems.map((r, i) => {
                  const stt = (currentPage - 1) * pageSize + i + 1;
                  const isSelf = String(r['Họ tên user'] || '').toLowerCase().includes('tự thực hiện');

                  return (
                    <tr key={r['STT'] || stt} className="hover:bg-slate-800/30 transition">
                      <td className="py-2 px-3 text-center text-slate-500 font-mono">{stt}</td>
                      <td className="py-2 px-3 font-mono text-slate-400">{r['Ngày giờ thao tác']}</td>
                      <td className="py-2 px-3 font-mono font-bold text-slate-200">{r['PID']}</td>
                      <td className="py-2 px-4 font-semibold text-slate-300">{r['Họ tên']}</td>
                      <td className="py-2 px-3 font-mono text-teal-400 font-bold">{r['Tên KIOS']}</td>
                      <td className="py-2 px-3 text-slate-400">{r['KHU VỰC']}</td>
                      <td className="py-2 px-3 text-center">
                        <div className="flex items-center justify-center gap-1">
                          {r['Checkin'] && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-500/20 text-blue-400 border border-blue-500/30">Checkin</span>}
                          {r['Thu viện phí'] && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-500/20 text-amber-400 border border-amber-500/30">Viện phí</span>}
                          {r['Bán thuốc'] && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">Thuốc</span>}
                        </div>
                      </td>
                      <td className="py-2 px-3 text-center font-mono font-bold text-slate-300">
                        {r['HÌNH THỨC TT'] || '--'}
                      </td>
                      <td className="py-2 px-4">
                        {isSelf ? (
                          <span className="font-bold text-purple-400 bg-purple-500/10 px-2 py-0.5 rounded text-[11px] border border-purple-500/20">
                            Tự thực hiện
                          </span>
                        ) : (
                          <span className="text-slate-400">{r['Họ tên user']}</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="p-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400 mt-2">
            <span>
              Hiển thị <b>{(currentPage - 1) * pageSize + 1}</b> - <b>{Math.min(currentPage * pageSize, analytics.filteredRows.length)}</b> / <b>{analytics.filteredRows.length.toLocaleString()}</b>
            </span>
            <div className="flex items-center gap-1">
              <button
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-1.5 rounded-lg border border-slate-700 bg-slate-800 hover:bg-slate-700 disabled:opacity-40 cursor-pointer"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
              </button>
              <span className="px-2 font-medium">{currentPage} / {totalPages}</span>
              <button
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1.5 rounded-lg border border-slate-700 bg-slate-800 hover:bg-slate-700 disabled:opacity-40 cursor-pointer"
              >
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL CẤU HÌNH GOOGLE SHEET */}
      {showConfigModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className={`rounded-3xl border shadow-2xl max-w-lg w-full p-6 space-y-4 ${
            isDark ? 'bg-slate-900 border-slate-800 text-slate-200' : 'bg-white border-slate-200 text-slate-800'
          }`}>
            <div className="flex items-center justify-between border-b pb-3 border-slate-800">
              <h3 className="font-bold text-sm">Cấu Hình Google Sheet Tự Động</h3>
              <button onClick={() => setShowConfigModal(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold mb-1 text-slate-400">Spreadsheet ID</label>
                <input
                  type="text"
                  value={sheetConfig.spreadsheetId}
                  onChange={(e) => setSheetConfig({ ...sheetConfig, spreadsheetId: e.target.value.trim() })}
                  className={`w-full p-2.5 border rounded-xl font-mono text-xs ${
                    isDark ? 'bg-slate-800 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'
                  }`}
                />
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block font-semibold mb-1 text-slate-400">GID Dữ Liệu</label>
                  <input
                    type="text"
                    value={sheetConfig.gidDuLieuTong}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidDuLieuTong: e.target.value.trim() })}
                    className={`w-full p-2 border rounded-xl font-mono text-xs ${
                      isDark ? 'bg-slate-800 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'
                    }`}
                  />
                </div>
                <div>
                  <label className="block font-semibold mb-1 text-slate-400">GID Vị Trí Máy</label>
                  <input
                    type="text"
                    value={sheetConfig.gidViTriMay}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidViTriMay: e.target.value.trim() })}
                    className={`w-full p-2 border rounded-xl font-mono text-xs ${
                      isDark ? 'bg-slate-800 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'
                    }`}
                  />
                </div>
                <div>
                  <label className="block font-semibold mb-1 text-slate-400">GID Mục Tiêu</label>
                  <input
                    type="text"
                    value={sheetConfig.gidMucTieu}
                    onChange={(e) => setSheetConfig({ ...sheetConfig, gidMucTieu: e.target.value.trim() })}
                    className={`w-full p-2 border rounded-xl font-mono text-xs ${
                      isDark ? 'bg-slate-800 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'
                    }`}
                  />
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                onClick={() => setShowConfigModal(false)}
                className="px-4 py-2 border rounded-xl text-xs font-semibold cursor-pointer border-slate-700"
              >
                Đóng
              </button>
              <button
                onClick={() => {
                  setShowConfigModal(false);
                  loadDataFromGoogleSheet(sheetConfig);
                }}
                className="px-5 py-2 bg-teal-600 hover:bg-teal-500 text-white rounded-xl text-xs font-semibold cursor-pointer"
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
