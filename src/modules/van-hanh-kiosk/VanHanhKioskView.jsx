import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Monitor, CreditCard, Clock, RefreshCw, Download, 
  Search, Building2, QrCode, FileText, Pill, 
  Settings, X, Sparkles, AlertCircle, ChevronLeft, 
  ChevronRight, Layers, Activity, BarChart3, Plus, 
  Edit3, Trash2, RotateCcw, Check, Users, Cpu,
  TrendingUp, Award, CheckCircle2
} from 'lucide-react';

import {
  DEFAULT_SHEET_CONFIG,
  DEFAULT_KIOSK_LOCATIONS,
  TARGET_AREAS,
  aggregateKioskData,
  syncAllKioskSheets,
  saveKioskLocations,
  resetKioskLocationsToDefault,
  exportKioskReportToExcel,
  normalizeKioskName
} from './engine';

export default function VanHanhKioskView() {
  const [rawRows, setRawRows] = useState([]);
  const [locationsList, setLocationsList] = useState([]);
  const [locationsMap, setLocationsMap] = useState({});
  const [targetsInfo, setTargetsInfo] = useState({
    totalDailyTarget: 0,
    totalTargetCLS: 0,
    totalTargetThuoc: 0,
    targetsMap: {}
  });

  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState(null);
  const [syncError, setSyncError] = useState(null);

  // Tab điều hướng chính
  const [activeTab, setActiveTab] = useState('OVERVIEW'); // 'OVERVIEW', 'SELF_SERVICE', 'DEPARTMENTS', 'KIOSK_CONFIG', 'LOGS'

  // Bộ lọc Dashboard
  const [dateFilter, setDateFilter] = useState('ALL');
  const [selectedShift, setSelectedShift] = useState('ALL');
  const [selectedArea, setSelectedArea] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');

  // Bộ lọc riêng cho Tab Quản Lý Kiosk
  const [kioskSearchTerm, setKioskSearchTerm] = useState('');
  const [kioskConfigFilter, setKioskConfigFilter] = useState('ALL');

  // Modal Thêm / Chỉnh Sửa Kiosk
  const [showKioskModal, setShowKioskModal] = useState(false);
  const [kioskFormData, setKioskFormData] = useState({
    originalMachine: '',
    machine: '',
    area: 'Trệt A - Sảnh chính',
    config: 'CSKH hỗ trợ'
  });

  // Phân trang
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 12;

  // Modal cấu hình Google Sheet
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
        setLocationsList(res.locationsList);
        setLocationsMap(res.locationsMap);
        setTargetsInfo({
          totalDailyTarget: res.totalDailyTarget,
          totalTargetCLS: res.totalTargetCLS,
          totalTargetThuoc: res.totalTargetThuoc,
          targetsMap: res.targetsMap
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
    loadDataFromGoogleSheet();
  }, []);

  // Cập nhật lại Map khi locationsList thay đổi
  const updateLocationsState = (newList) => {
    setLocationsList(newList);
    saveKioskLocations(newList);
    const newMap = {};
    newList.forEach(loc => {
      const norm = normalizeKioskName(loc.machine);
      if (norm) newMap[norm] = { area: loc.area, machine: loc.machine, config: loc.config };
    });
    setLocationsMap(newMap);
  };

  // Mở modal thêm Kiosk mới
  const handleOpenAddKiosk = () => {
    setKioskFormData({
      originalMachine: '',
      machine: '',
      area: TARGET_AREAS[0] || 'Trệt A - Sảnh chính',
      config: 'CSKH hỗ trợ'
    });
    setShowKioskModal(true);
  };

  // Mở modal sửa Kiosk
  const handleOpenEditKiosk = (item) => {
    setKioskFormData({
      originalMachine: item.machine,
      machine: item.machine,
      area: item.area,
      config: item.config
    });
    setShowKioskModal(true);
  };

  // Lưu thông tin Kiosk (Thêm hoặc Sửa)
  const handleSaveKiosk = (e) => {
    e.preventDefault();
    const cleanMachine = kioskFormData.machine.trim().toUpperCase();
    if (!cleanMachine) {
      alert('Vui lòng nhập tên máy Kiosk (ví dụ KIOS17, KIOS99)');
      return;
    }

    let updatedList = [...locationsList];

    if (kioskFormData.originalMachine) {
      // Đang sửa Kiosk hiện tại
      const normOrig = normalizeKioskName(kioskFormData.originalMachine);
      updatedList = updatedList.map(item => {
        if (normalizeKioskName(item.machine) === normOrig) {
          return {
            machine: cleanMachine,
            area: kioskFormData.area.trim(),
            config: kioskFormData.config
          };
        }
        return item;
      });
    } else {
      // Thêm Kiosk mới
      const normNew = normalizeKioskName(cleanMachine);
      if (updatedList.some(k => normalizeKioskName(k.machine) === normNew)) {
        alert(`Tên máy ${cleanMachine} đã tồn tại trong danh sách! Vui lòng chọn tên khác.`);
        return;
      }
      updatedList.push({
        machine: cleanMachine,
        area: kioskFormData.area.trim(),
        config: kioskFormData.config
      });
    }

    updateLocationsState(updatedList);
    setShowKioskModal(false);
  };

  // Xóa Kiosk
  const handleDeleteKiosk = (machineName) => {
    if (confirm(`Bạn có chắc muốn xóa máy Kiosk "${machineName}" khỏi danh sách cấu hình?`)) {
      const norm = normalizeKioskName(machineName);
      const updatedList = locationsList.filter(item => normalizeKioskName(item.machine) !== norm);
      updateLocationsState(updatedList);
    }
  };

  // Khôi phục danh sách gốc từ Sheet
  const handleResetKioskToDefault = () => {
    if (confirm('Khôi phục lại danh sách máy và cấu hình gốc ban đầu từ Google Sheet?')) {
      resetKioskLocationsToDefault();
      loadDataFromGoogleSheet();
    }
  };

  // Danh sách các ngày thực tế có trong Google Sheet
  const availableDates = useMemo(() => {
    const s = new Set();
    rawRows.forEach(r => {
      const d = r['NGÀY'] || (r['Ngày giờ thao tác'] || '').split(' ')[0];
      if (d && d.length >= 8) s.add(d);
    });
    return Array.from(s).sort().reverse();
  }, [rawRows]);

  // Phân tích dữ liệu Dashboard 100% từ Sheet thực tế
  const analytics = useMemo(() => {
    return aggregateKioskData(rawRows, locationsMap, targetsInfo, {
      dateFilter,
      selectedShift,
      selectedArea,
      searchTerm
    });
  }, [rawRows, locationsMap, targetsInfo, dateFilter, selectedShift, selectedArea, searchTerm]);

  // Lọc danh sách máy trong Tab Quản lý Kiosk
  const filteredKioskList = useMemo(() => {
    return locationsList.filter(item => {
      if (kioskConfigFilter !== 'ALL' && item.config !== kioskConfigFilter) return false;
      if (kioskSearchTerm) {
        const q = kioskSearchTerm.toLowerCase().trim();
        const matchMachine = item.machine.toLowerCase().includes(q);
        const matchArea = item.area.toLowerCase().includes(q);
        return matchMachine || matchArea;
      }
      return true;
    });
  }, [locationsList, kioskConfigFilter, kioskSearchTerm]);

  const totalPages = Math.ceil(analytics.filteredRows.length / pageSize) || 1;
  const paginatedItems = analytics.filteredRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  // Đếm nhanh số máy tự làm vs có hỗ trợ
  const countSelfKiosks = useMemo(() => {
    return locationsList.filter(k => k.config === 'Tự thực hiện').length;
  }, [locationsList]);

  // Tìm giá trị max cho biểu đồ năng suất theo giờ toàn viện
  const maxHourlyGeneral = useMemo(() => {
    return Math.max(...analytics.hourlyGeneralList.map(h => h.total), 1);
  }, [analytics.hourlyGeneralList]);

  // Tìm giá trị max cho biểu đồ năng suất máy tự thực hiện
  const maxHourlySelf = useMemo(() => {
    return Math.max(...analytics.hourlySelfList.map(h => h.total), 1);
  }, [analytics.hourlySelfList]);

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
              <Sparkles className="w-3.5 h-3.5" />
              <span>Máy Tự Thực Hiện ({analytics.summary.selfTotal.toLocaleString()})</span>
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
              <span>Khoa Phòng &amp; Mục Tiêu ({analytics.areaBreakdown.length})</span>
            </button>

            <button
              onClick={() => setActiveTab('KIOSK_CONFIG')}
              className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                activeTab === 'KIOSK_CONFIG'
                  ? 'bg-teal-700 text-white shadow-xs'
                  : 'text-teal-700 hover:text-teal-900'
              }`}
            >
              <Cpu className="w-3.5 h-3.5" />
              <span>Cấu Hình Kiosk ({locationsList.length})</span>
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

          {/* Lọc Ngày, Ca, Khu vực cho màn hình Dashboard */}
          {activeTab !== 'KIOSK_CONFIG' && (
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
          )}
        </div>
      </div>

      {/* CẢNH BÁO LỖI NẾU CÓ */}
      {syncError && (
        <div className="bg-amber-50 border border-amber-300 rounded-2xl p-4 text-amber-900 flex items-center justify-between text-xs shadow-xs">
          <div className="flex items-center gap-2.5">
            <AlertCircle className="w-4 h-4 text-amber-600 flex-shrink-0" />
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

      {/* ============================================================ */}
      {/* 2. TAB TỔNG QUAN ĐIỀU HÀNH (OVERVIEW) */}
      {/* ============================================================ */}
      {activeTab === 'OVERVIEW' && (
        <div className="space-y-6">
          {/* 4 THẺ METRICS SÁNG TƯƠI, ĐỘ TƯƠNG PHẢN CAO, SỐ LIỆU CHÍNH XÁC */}
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
                    lượt
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600">
                <span>Mục tiêu ngày: <b className="text-slate-900 font-mono">{targetsInfo.totalDailyTarget > 0 ? `${targetsInfo.totalDailyTarget.toLocaleString()} lượt` : 'Chờ nạp'}</b></span>
                {targetsInfo.totalDailyTarget > 0 && (
                  <span className="font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded-md border border-teal-200">
                    Đạt {analytics.summary.overallRate}%
                  </span>
                )}
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
                    Chỉ tiêu: {targetsInfo.totalDailyTarget.toLocaleString()} lượt
                  </span>
                </div>
                <div className="mt-2 flex items-baseline justify-between">
                  <span className="text-4xl font-black font-mono tracking-tight text-amber-600">
                    {analytics.summary.overallRate}%
                  </span>
                  <span className="text-xs text-slate-600 font-mono font-bold">
                    {analytics.summary.totalOps.toLocaleString()} / {targetsInfo.totalDailyTarget.toLocaleString()} lượt
                  </span>
                </div>
              </div>
              <div className="mt-4 space-y-1.5">
                <div className="w-full bg-slate-100 h-3 rounded-full overflow-hidden border border-slate-200 relative">
                  <div 
                    className="bg-amber-500 h-full rounded-full transition-all duration-500" 
                    style={{ width: `${Math.min(100, Number(analytics.summary.overallRate))}%` }} 
                  />
                </div>
                <div className="flex justify-between text-[11px] text-slate-600 font-medium">
                  <span>Đã thực hiện: <b className="text-slate-900 font-bold">{analytics.summary.totalOps.toLocaleString()} lượt</b></span>
                  <span>Còn lại: <b className="text-amber-700 font-bold">{Math.max(0, targetsInfo.totalDailyTarget - analytics.summary.totalOps).toLocaleString()} lượt</b></span>
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
                    {(analytics.summary.totalCheckin + analytics.summary.totalVienPhi).toLocaleString()}
                  </span>
                  <span className="text-xs text-slate-600 font-mono font-bold">
                    lượt ({analytics.summary.regCompletedRate}% chỉ tiêu)
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                <span className="text-blue-700 font-bold font-mono bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                  Checkin: {analytics.summary.totalCheckin.toLocaleString()} lượt
                </span>
                <span className="text-amber-800 font-bold font-mono bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                  Viện phí: {analytics.summary.totalVienPhi.toLocaleString()} lượt
                </span>
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
                  <span className="text-3xl font-black font-mono tracking-tight text-purple-700">
                    {analytics.summary.rateCK}%
                  </span>
                  <span className="text-xs text-purple-900 font-bold bg-purple-50 px-2.5 py-0.5 rounded-full border border-purple-200">
                    Chuyển Khoản QR
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                <span className="text-slate-600 font-medium">
                  Quẹt POS: <b className="text-teal-700 font-mono font-bold">{analytics.summary.ratePOS}%</b> ({analytics.summary.countPOS.toLocaleString()} lượt)
                </span>
                <span className="text-purple-700 font-bold font-mono">
                  CK: {analytics.summary.countCK.toLocaleString()} lượt
                </span>
              </div>
            </div>
          </div>

          {/* ============================================================ */}
          {/* BIỂU ĐỒ NĂNG SUẤT THEO KHUNG GIỜ TRONG NGÀY (TOÀN VIỆN) */}
          {/* CÓ KÈM NHÃN DỮ LIỆU HIỆN TRỰC TIẾP TRÊN TỪNG CỘT */}
          {/* ============================================================ */}
          <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs space-y-5">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-xl bg-teal-50 text-teal-700 flex items-center justify-center border border-teal-200">
                    <BarChart3 className="w-4 h-4" />
                  </div>
                  <h2 className="text-base font-black text-slate-900 tracking-tight uppercase">
                    BIỂU ĐỒ NĂNG SUẤT THEO KHUNG GIỜ TRONG NGÀY (TOÀN VIỆN)
                  </h2>
                </div>
                <p className="text-xs text-slate-500 mt-1">
                  Phân tích lưu lượng xử lý thực tế theo từng giờ phát sinh giao dịch trong ngày. <b>Đơn vị tiêu chuẩn: Lượt / Giờ</b>.
                </p>
              </div>

              {/* Mốc dữ liệu thực tế: Đỉnh tải, Năng suất bình quân, Thời gian hoạt động */}
              <div className="flex flex-wrap items-center gap-2 text-xs">
                {analytics.summary.peakHourGeneral.total > 0 && (
                  <div className="px-3 py-1.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-800 font-bold flex items-center gap-1.5">
                    <TrendingUp className="w-3.5 h-3.5 text-rose-600" />
                    <span>Đỉnh tải thực tế: <b className="font-mono text-rose-950">{analytics.summary.peakHourGeneral.hour}</b> (<b className="font-mono">{analytics.summary.peakHourGeneral.total.toLocaleString()} lượt/giờ</b>)</span>
                  </div>
                )}

                <div className="px-3 py-1.5 rounded-xl bg-teal-50 border border-teal-200 text-teal-800 font-bold flex items-center gap-1.5">
                  <Activity className="w-3.5 h-3.5 text-teal-600" />
                  <span>Năng suất bình quân: <b className="font-mono text-teal-950">{analytics.summary.avgHourlyGeneral} lượt/giờ</b></span>
                </div>

                {analytics.summary.firstActivityTime !== '--:--' && (
                  <div className="px-3 py-1.5 rounded-xl bg-slate-100 border border-slate-200 text-slate-700 font-mono text-[11px] font-semibold flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5 text-slate-500" />
                    <span>Thời gian giao dịch: <b>{analytics.summary.firstActivityTime} → {analytics.summary.lastActivityTime}</b></span>
                  </div>
                )}
              </div>
            </div>

            {/* Chú thích màu sắc */}
            <div className="flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-4 font-bold">
                <span className="flex items-center gap-1.5 text-blue-700 bg-blue-50 px-2.5 py-1 rounded-lg border border-blue-200">
                  <span className="w-2.5 h-2.5 rounded-xs bg-blue-600"></span> Check-in ({analytics.summary.totalCheckin.toLocaleString()} lượt)
                </span>
                <span className="flex items-center gap-1.5 text-amber-800 bg-amber-50 px-2.5 py-1 rounded-lg border border-amber-200">
                  <span className="w-2.5 h-2.5 rounded-xs bg-amber-500"></span> Thu viện phí ({analytics.summary.totalVienPhi.toLocaleString()} lượt)
                </span>
                <span className="flex items-center gap-1.5 text-emerald-800 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200">
                  <span className="w-2.5 h-2.5 rounded-xs bg-emerald-600"></span> Bán thuốc ({analytics.summary.totalBanThuoc.toLocaleString()} lượt)
                </span>
              </div>
              <span className="text-[11px] text-slate-400 italic">
                * Nhãn số trên đỉnh mỗi cột hiển thị tổng lượt giao dịch thực tế của khung giờ đó
              </span>
            </div>

            {/* KHUNG BIỂU ĐỒ CỘT CÓ NHÃN SỐ HIỂN THỊ TRỰC TIẾP TRÊN ĐẦU CỘT */}
            <div className="pt-6 pb-2">
              <div className="h-64 flex items-end gap-2 sm:gap-3 border-b-2 border-slate-200 px-2">
                {analytics.hourlyGeneralList.map((item) => {
                  const pct = Math.round((item.total / maxHourlyGeneral) * 100);
                  const isPeak = item.total > 0 && item.total === analytics.summary.peakHourGeneral.total;

                  // Tính tỷ lệ chiều cao từng phân đoạn trong cột
                  const pCheckin = item.total > 0 ? (item.checkin / item.total) * 100 : 0;
                  const pVienPhi = item.total > 0 ? (item.vienPhi / item.total) * 100 : 0;
                  const pBanThuoc = item.total > 0 ? (item.banThuoc / item.total) * 100 : 0;

                  return (
                    <div key={item.hour} className="flex-1 flex flex-col items-center h-full min-w-[28px]">
                      {/* ZONE 1: NHÃN ĐỈNH CỘT (Chiều cao cố định h-9 để vùng vẽ cột của mọi giờ đều bằng nhau 100%) */}
                      <div className="h-9 w-full flex flex-col items-center justify-end pb-1">
                        {isPeak && (
                          <span className="text-[9px] font-black uppercase text-rose-600 bg-rose-50 border border-rose-200 px-1 py-0.2 rounded mb-0.5 tracking-tighter whitespace-nowrap">
                            Đỉnh
                          </span>
                        )}
                        <span className={`text-[11px] sm:text-xs font-mono font-black ${
                          isPeak ? 'text-rose-600 font-extrabold' : item.total > 0 ? 'text-slate-800' : 'text-slate-300'
                        }`}>
                          {item.total > 0 ? item.total : '0'}
                        </span>
                      </div>

                      {/* ZONE 2: KHU VỰC VẼ THÂN CỘT (Chiếm toàn bộ không gian còn lại flex-1, tỷ lệ chiều cao tuyệt đối chuẩn xác) */}
                      <div className="flex-1 w-full flex items-end justify-center">
                        <div 
                          className={`w-full max-w-[36px] rounded-t-md overflow-hidden flex flex-col justify-end transition-all duration-300 border border-b-0 ${
                            isPeak ? 'border-rose-400 ring-2 ring-rose-200' : 'border-slate-300'
                          } ${item.total === 0 ? 'bg-slate-100' : ''}`}
                          style={{ height: `${Math.max(item.total > 0 ? 3 : 1, pct)}%` }}
                          title={`${item.hour}: Tổng ${item.total} lượt (Checkin: ${item.checkin}, Viện phí: ${item.vienPhi}, Thuốc: ${item.banThuoc})`}
                        >
                          {item.total > 0 && (
                            <>
                              {item.banThuoc > 0 && (
                                <div 
                                  className="bg-emerald-600 w-full" 
                                  style={{ height: `${pBanThuoc}%` }} 
                                />
                              )}
                              {item.vienPhi > 0 && (
                                <div 
                                  className="bg-amber-500 w-full" 
                                  style={{ height: `${pVienPhi}%` }} 
                                />
                              )}
                              {item.checkin > 0 && (
                                <div 
                                  className="bg-blue-600 w-full" 
                                  style={{ height: `${pCheckin}%` }} 
                                />
                              )}
                            </>
                          )}
                        </div>
                      </div>

                      {/* ZONE 3: NHÃN TRỤC X DƯỚI CHÂN CỘT (Chiều cao cố định h-6) */}
                      <div className="h-6 w-full flex items-center justify-center pt-1.5 border-t border-slate-200">
                        <span className={`text-[11px] font-mono font-bold block ${
                          isPeak ? 'text-rose-600 font-black' : 'text-slate-600'
                        }`}>
                          {item.label}
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* ============================================================ */}
          {/* BẢNG PHÂN BỔ TẢI TRỌNG KHOA PHÒNG & KHU VỰC */}
          {/* CÓ NHÃN SỐ TRỰC TIẾP TRÊN TỪNG KHỐI PHÂN ĐOẠN */}
          {/* ============================================================ */}
          <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs space-y-5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
              <div>
                <h3 className="text-base font-extrabold text-slate-900 tracking-tight">
                  Phân Bổ Tải Trọng Các Khoa Phòng &amp; Khu Vực
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Số lượt xử lý thực tế (Check-in + Viện Phí + Bán Thuốc) tại từng vị trí đặt máy
                </p>
              </div>

              <div className="flex items-center gap-3 text-xs font-bold">
                <span className="flex items-center gap-1.5 text-blue-700 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                  <span className="w-2.5 h-2.5 rounded-xs bg-blue-600"></span> Check-in
                </span>
                <span className="flex items-center gap-1.5 text-amber-800 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                  <span className="w-2.5 h-2.5 rounded-xs bg-amber-500"></span> Thu viện phí
                </span>
                <span className="flex items-center gap-1.5 text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                  <span className="w-2.5 h-2.5 rounded-xs bg-emerald-600"></span> Bán thuốc
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
                  <div key={item.area} className="p-4 rounded-2xl border border-slate-200 bg-slate-50/40 hover:bg-white hover:border-teal-400 hover:shadow-xs transition-all duration-200 space-y-2.5">
                    <div className="flex items-center justify-between text-xs">
                      <div className="flex items-center gap-2">
                        <span className="font-mono text-slate-400 text-xs font-bold">#{idx + 1}</span>
                        <span className="font-bold text-slate-900 text-sm">{item.area}</span>
                      </div>
                      <span className="font-mono font-black text-sm text-teal-800 bg-teal-50 px-2.5 py-0.5 rounded-lg border border-teal-200">
                        {total.toLocaleString()} lượt
                      </span>
                    </div>

                    {/* Thanh phân bổ xếp chồng */}
                    <div className="h-4 w-full bg-slate-200 rounded-full overflow-hidden flex border border-slate-200">
                      {item.checkin > 0 && (
                        <div 
                          className="bg-blue-600 h-full transition-all" 
                          style={{ width: `${pCheckin}%` }} 
                        />
                      )}
                      {item.vienPhi > 0 && (
                        <div 
                          className="bg-amber-500 h-full transition-all" 
                          style={{ width: `${pVienPhi}%` }} 
                        />
                      )}
                      {item.banThuoc > 0 && (
                        <div 
                          className="bg-emerald-600 h-full transition-all" 
                          style={{ width: `${pBanThuoc}%` }} 
                        />
                      )}
                    </div>

                    {/* NHÃN DỮ LIỆU HIỆN TRỰC TIẾP CHO TỪNG PHÂN ĐOẠN */}
                    <div className="flex items-center justify-between text-xs font-mono font-semibold pt-0.5">
                      <span className="text-blue-700 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                        Checkin: <b>{item.checkin.toLocaleString()} lượt</b>
                      </span>
                      <span className="text-amber-800 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                        Viện phí: <b>{item.vienPhi.toLocaleString()} lượt</b>
                      </span>
                      <span className="text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                        Thuốc: <b>{item.banThuoc.toLocaleString()} lượt</b>
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* ============================================================ */}
      {/* 3. TAB MÁY TỰ THỰC HIỆN (SELF_SERVICE) */}
      {/* ============================================================ */}
      {activeTab === 'SELF_SERVICE' && (
        <div className="space-y-6">
          {/* 4 THẺ METRICS DÀNH RIÊNG CHO MÁY TỰ LÀM */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="p-5 rounded-3xl border border-purple-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-purple-700">Tổng Lượt Tự Làm</span>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-4xl font-black font-mono text-purple-700">
                  {analytics.summary.selfTotal.toLocaleString()}
                </span>
                <span className="text-xs font-bold text-purple-800 bg-purple-50 px-2 py-0.5 rounded-full border border-purple-200">
                  lượt
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-purple-100 font-medium">
                Chiếm <b className="text-purple-700 font-mono">{analytics.summary.selfShareRate}%</b> tổng giao dịch toàn viện
              </p>
            </div>

            <div className="p-5 rounded-3xl border border-blue-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-blue-700">Tự Check-in</span>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-4xl font-black font-mono text-blue-600">
                  {analytics.summary.selfCheckin.toLocaleString()}
                </span>
                <span className="text-xs font-bold text-blue-700 bg-blue-50 px-2 py-0.5 rounded-full border border-blue-200">
                  lượt
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-blue-100 font-medium">
                Tỷ trọng: <b className="text-blue-700 font-mono">{analytics.summary.selfTotal > 0 ? ((analytics.summary.selfCheckin / analytics.summary.selfTotal) * 100).toFixed(1) : 0}%</b> trên máy tự làm
              </p>
            </div>

            <div className="p-5 rounded-3xl border border-amber-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-amber-700">Tự Nộp Viện Phí</span>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-4xl font-black font-mono text-amber-600">
                  {analytics.summary.selfVienPhi.toLocaleString()}
                </span>
                <span className="text-xs font-bold text-amber-800 bg-amber-50 px-2 py-0.5 rounded-full border border-amber-200">
                  lượt
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-amber-100 font-medium">
                Tỷ trọng: <b className="text-amber-700 font-mono">{analytics.summary.selfTotal > 0 ? ((analytics.summary.selfVienPhi / analytics.summary.selfTotal) * 100).toFixed(1) : 0}%</b> trên máy tự làm
              </p>
            </div>

            <div className="p-5 rounded-3xl border border-emerald-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-emerald-700">Thanh Toán Tại Máy</span>
              <div className="mt-2 flex items-baseline justify-between">
                <span className="text-2xl font-black font-mono text-emerald-700">
                  CK {analytics.summary.selfRateCK}%
                </span>
                <span className="text-xs font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                  POS {analytics.summary.selfRatePOS}%
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-emerald-100 font-medium">
                Bệnh nhân quét mã QR hoặc chạm thẻ trực tiếp
              </p>
            </div>
          </div>

          {/* BIỂU ĐỒ NĂNG SUẤT THEO KHUNG GIỜ - MÁY TỰ THỰC HIỆN */}
          <div className="p-6 rounded-3xl border-2 border-purple-200 bg-white shadow-xs space-y-5">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-purple-100 pb-4">
              <div>
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-xl bg-purple-100 text-purple-700 flex items-center justify-center border border-purple-200">
                    <Sparkles className="w-4 h-4 text-purple-600" />
                  </div>
                  <h2 className="text-base font-black text-purple-950 uppercase tracking-tight">
                    BIỂU ĐỒ NĂNG SUẤT THEO KHUNG GIỜ - MÁY TỰ THỰC HIỆN
                  </h2>
                </div>
                <p className="text-xs text-purple-800 font-medium mt-1">
                  Đo lường năng suất người bệnh tự thao tác độc lập theo từng khung giờ trong ngày. <b>Đơn vị: Lượt tự làm / Giờ</b>.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2 text-xs">
                {analytics.summary.peakHourSelf.total > 0 && (
                  <div className="px-3 py-1.5 rounded-xl bg-purple-100 border border-purple-300 text-purple-950 font-bold flex items-center gap-1.5">
                    <TrendingUp className="w-3.5 h-3.5 text-purple-700" />
                    <span>Đỉnh tải tự làm thực tế: <b className="font-mono">{analytics.summary.peakHourSelf.hour}</b> (<b className="font-mono">{analytics.summary.peakHourSelf.total.toLocaleString()} lượt/giờ</b>)</span>
                  </div>
                )}
                <div className="px-3 py-1.5 rounded-xl bg-purple-50 border border-purple-200 text-purple-800 font-bold">
                  Năng suất bình quân: <b className="font-mono">{analytics.summary.avgHourlySelf} lượt/giờ</b>
                </div>
              </div>
            </div>

            {/* Khung cột biểu đồ tự làm có nhãn số trực tiếp */}
            <div className="pt-6 pb-2">
              <div className="h-56 flex items-end gap-2 sm:gap-3 border-b-2 border-purple-200 px-2">
                {analytics.hourlySelfList.map((item) => {
                  const pct = Math.round((item.total / maxHourlySelf) * 100);
                  const isPeak = item.total > 0 && item.total === analytics.summary.peakHourSelf.total;

                  return (
                    <div key={item.hour} className="flex-1 flex flex-col items-center h-full min-w-[28px]">
                      {/* ZONE 1: NHÃN ĐỈNH CỘT (Chiều cao cố định h-9) */}
                      <div className="h-9 w-full flex flex-col items-center justify-end pb-1">
                        {isPeak && (
                          <span className="text-[9px] font-black uppercase text-purple-700 bg-purple-100 border border-purple-300 px-1 py-0.2 rounded mb-0.5 tracking-tighter whitespace-nowrap">
                            Đỉnh
                          </span>
                        )}
                        <span className={`text-[11px] sm:text-xs font-mono font-black ${
                          isPeak ? 'text-purple-700 font-extrabold' : item.total > 0 ? 'text-slate-800' : 'text-slate-300'
                        }`}>
                          {item.total > 0 ? item.total : '0'}
                        </span>
                      </div>

                      {/* ZONE 2: KHU VỰC VẼ THÂN CỘT (Chiếm toàn bộ không gian còn lại flex-1) */}
                      <div className="flex-1 w-full flex items-end justify-center">
                        <div 
                          className={`w-full max-w-[36px] bg-purple-600 rounded-t-md transition-all duration-300 border border-b-0 ${
                            isPeak ? 'border-purple-800 ring-2 ring-purple-300' : 'border-purple-700'
                          }`}
                          style={{ height: `${Math.max(item.total > 0 ? 3 : 1, pct)}%` }}
                          title={`${item.hour}: ${item.total} lượt tự làm`}
                        />
                      </div>

                      {/* ZONE 3: NHÃN TRỤC X DƯỚI CHÂN CỘT (Chiều cao cố định h-6) */}
                      <div className="h-6 w-full flex items-center justify-center pt-1.5 border-t border-purple-200">
                        <span className={`text-[11px] font-mono font-bold block ${
                          isPeak ? 'text-purple-700 font-black' : 'text-slate-600'
                        }`}>
                          {item.label}
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* BẢNG XẾP HẠNG CÁC MÁY TỰ THỰC HIỆN */}
          <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div>
                <h3 className="text-sm font-extrabold text-slate-900 uppercase">
                  DANH SÁCH &amp; NĂNG SUẤT CÁC MÁY CẤU HÌNH "TỰ THỰC HIỆN"
                </h3>
                <p className="text-xs text-slate-500">Hiệu quả vận hành của từng máy tự phục vụ bệnh nhân</p>
              </div>
              <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-purple-100 text-purple-900 border border-purple-200 font-mono">
                {analytics.selfKiosks.length} Máy hoạt động
              </span>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-slate-200">
              <table className="w-full text-left text-xs text-slate-700">
                <thead className="border-b border-slate-200 bg-slate-100 text-slate-800 font-bold uppercase tracking-wider text-[11px]">
                  <tr>
                    <th className="py-3 px-3 w-12 text-center">STT</th>
                    <th className="py-3 px-4">Tên Máy Kiosk</th>
                    <th className="py-3 px-4">Vị Trí Đặt Máy</th>
                    <th className="py-3 px-4 text-center">Tổng Lượt Tự Làm</th>
                    <th className="py-3 px-4 text-center">Tự Check-in</th>
                    <th className="py-3 px-4 text-center">Tự Viện Phí</th>
                    <th className="py-3 px-4 text-center">Tự Bán Thuốc</th>
                    <th className="py-3 px-4 text-center">Giao Dịch Gần Nhất</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {analytics.selfKiosks.map((k, i) => (
                    <tr key={k.machine} className="hover:bg-purple-50/40 transition">
                      <td className="py-3 px-3 text-center text-slate-400 font-mono font-bold">{i + 1}</td>
                      <td className="py-3 px-4 font-mono font-black text-sm text-purple-800">{k.machine}</td>
                      <td className="py-3 px-4 font-medium text-slate-800">{k.area}</td>
                      <td className="py-3 px-4 text-center font-mono font-black text-sm text-purple-700 bg-purple-50/70">
                        {k.count.toLocaleString()} lượt
                      </td>
                      <td className="py-3 px-4 text-center font-mono font-bold text-blue-700">
                        {k.checkin.toLocaleString()}
                      </td>
                      <td className="py-3 px-4 text-center font-mono font-bold text-amber-700">
                        {k.vienPhi.toLocaleString()}
                      </td>
                      <td className="py-3 px-4 text-center font-mono font-bold text-emerald-700">
                        {k.banThuoc.toLocaleString()}
                      </td>
                      <td className="py-3 px-4 text-center font-mono text-slate-500">
                        {k.lastActiveTime || '--'}
                      </td>
                    </tr>
                  ))}
                  {analytics.selfKiosks.length === 0 && (
                    <tr>
                      <td colSpan="8" className="py-8 text-center text-slate-400">
                        Chưa có máy Kiosk nào cấu hình "Tự thực hiện" hoặc chưa có giao dịch phát sinh.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ============================================================ */}
      {/* 4. TAB KHOA PHÒNG & MỤC TIÊU (DEPARTMENTS) */}
      {/* ĐỐI CHIẾU THỰC TẾ VS SHEET MỤC TIÊU 100% */}
      {/* ============================================================ */}
      {activeTab === 'DEPARTMENTS' && (
        <div className="space-y-6">
          {/* TỔNG KẾT MỤC TIÊU TOÀN VIỆN TỪ SHEET */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="p-5 rounded-3xl border border-slate-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-slate-500">Tổng Chỉ Tiêu Toàn Viện</span>
              <div className="mt-2 flex items-baseline justify-between">
                <span className="text-3xl font-black font-mono text-slate-900">
                  {targetsInfo.totalDailyTarget.toLocaleString()} lượt
                </span>
                <span className="text-xs font-bold text-teal-800 bg-teal-50 px-2.5 py-0.5 rounded-full border border-teal-200">
                  Đạt {analytics.summary.overallRate}%
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-slate-100">
                Thực tế đạt được: <b className="text-teal-700 font-mono font-bold">{analytics.summary.totalOps.toLocaleString()} lượt</b>
              </p>
            </div>

            <div className="p-5 rounded-3xl border border-blue-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-blue-700">Chỉ Tiêu Đăng Ký &amp; Thu CLS</span>
              <div className="mt-2 flex items-baseline justify-between">
                <span className="text-3xl font-black font-mono text-blue-700">
                  {targetsInfo.totalTargetCLS.toLocaleString()} lượt
                </span>
                <span className="text-xs font-bold text-blue-800 bg-blue-50 px-2.5 py-0.5 rounded-full border border-blue-200">
                  Đạt {analytics.summary.regCompletedRate}%
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-blue-100">
                Thực tế: <b className="text-blue-700 font-mono font-bold">{(analytics.summary.totalCheckin + analytics.summary.totalVienPhi).toLocaleString()} lượt</b>
              </p>
            </div>

            <div className="p-5 rounded-3xl border border-emerald-200 bg-white shadow-xs">
              <span className="text-xs font-bold uppercase text-emerald-700">Chỉ Tiêu Bán Thuốc</span>
              <div className="mt-2 flex items-baseline justify-between">
                <span className="text-3xl font-black font-mono text-emerald-700">
                  {targetsInfo.totalTargetThuoc.toLocaleString()} lượt
                </span>
                <span className="text-xs font-bold text-emerald-800 bg-emerald-50 px-2.5 py-0.5 rounded-full border border-emerald-200">
                  Đạt {analytics.summary.medCompletedRate}%
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2 pt-2 border-t border-emerald-100">
                Thực tế: <b className="text-emerald-700 font-mono font-bold">{analytics.summary.totalBanThuoc.toLocaleString()} lượt</b>
              </p>
            </div>
          </div>

          {/* BẢNG ĐỐI CHIẾU TIẾN ĐỘ TỪNG KHOA PHÒNG VS SHEET MỤC TIÊU */}
          <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
              <div>
                <h3 className="text-base font-extrabold text-slate-900 uppercase">
                  BẢNG ĐỐI CHIẾU TIẾN ĐỘ THỰC TẾ VS MỤC TIÊU TỪNG KHOA PHÒNG
                </h3>
                <p className="text-xs text-slate-500">
                  Khớp nối 100% giữa Sheet DỮ LIỆU TỔNG và Sheet MỤC TIÊU (Có chia mục tiêu Ca Sáng / Ca Chiều)
                </p>
              </div>
              <span className="text-xs font-mono font-bold text-slate-600 bg-slate-100 px-3 py-1 rounded-xl">
                {analytics.areaBreakdown.length} Khoa phòng
              </span>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-slate-200">
              <table className="w-full text-left text-xs text-slate-700">
                <thead className="border-b border-slate-200 bg-slate-100 text-slate-800 font-bold uppercase tracking-wider text-[11px]">
                  <tr>
                    <th className="py-3 px-3 w-12 text-center">STT</th>
                    <th className="py-3 px-4">Khoa Phòng / Khu Vực</th>
                    <th className="py-3 px-4 text-center">Thực Tế CLS / Chỉ Tiêu</th>
                    <th className="py-3 px-4 text-center w-36">% Đạt CLS</th>
                    <th className="py-3 px-4 text-center">Thực Tế Thuốc / Chỉ Tiêu</th>
                    <th className="py-3 px-4 text-center w-36">% Đạt Thuốc</th>
                    <th className="py-3 px-4 text-center">Tổng Thực Tế / Chỉ Tiêu</th>
                    <th className="py-3 px-4 text-center">Chỉ Tiêu Ca Sáng/Chiều</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-medium">
                  {analytics.areaBreakdown.map((item, idx) => {
                    const rateTotalNum = Number(item.rateTotal);
                    const isCompleted = item.totalTarget > 0 && rateTotalNum >= 100;

                    return (
                      <tr key={item.area} className="hover:bg-slate-50 transition">
                        <td className="py-3 px-3 text-center text-slate-400 font-mono font-bold">{idx + 1}</td>
                        <td className="py-3 px-4 font-bold text-slate-900 text-sm">
                          <div className="flex items-center gap-1.5">
                            {isCompleted && <CheckCircle2 className="w-4 h-4 text-emerald-600 flex-shrink-0" />}
                            <span>{item.area}</span>
                          </div>
                        </td>

                        {/* CLS Thực tế / Mục tiêu */}
                        <td className="py-3 px-4 text-center font-mono font-bold">
                          <span className="text-blue-700">{item.clsActual.toLocaleString()}</span>
                          <span className="text-slate-400"> / </span>
                          <span className="text-slate-700">{item.targetCLS > 0 ? item.targetCLS.toLocaleString() : '--'}</span>
                        </td>

                        {/* % Đạt CLS có thanh progress và nhãn số trực tiếp */}
                        <td className="py-3 px-4 text-center">
                          <div className="space-y-1">
                            <span className="font-mono font-bold text-xs text-blue-700 block">
                              {item.targetCLS > 0 ? `${item.rateCLS}%` : '--'}
                            </span>
                            {item.targetCLS > 0 && (
                              <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden border border-slate-200">
                                <div 
                                  className="bg-blue-600 h-full rounded-full" 
                                  style={{ width: `${Math.min(100, Number(item.rateCLS))}%` }} 
                                />
                              </div>
                            )}
                          </div>
                        </td>

                        {/* Thuốc Thực tế / Mục tiêu */}
                        <td className="py-3 px-4 text-center font-mono font-bold">
                          <span className="text-emerald-700">{item.medActual.toLocaleString()}</span>
                          <span className="text-slate-400"> / </span>
                          <span className="text-slate-700">{item.targetThuoc > 0 ? item.targetThuoc.toLocaleString() : '--'}</span>
                        </td>

                        {/* % Đạt Thuốc có thanh progress và nhãn số trực tiếp */}
                        <td className="py-3 px-4 text-center">
                          <div className="space-y-1">
                            <span className="font-mono font-bold text-xs text-emerald-700 block">
                              {item.targetThuoc > 0 ? `${item.rateThuoc}%` : '--'}
                            </span>
                            {item.targetThuoc > 0 && (
                              <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden border border-slate-200">
                                <div 
                                  className="bg-emerald-600 h-full rounded-full" 
                                  style={{ width: `${Math.min(100, Number(item.rateThuoc))}%` }} 
                                />
                              </div>
                            )}
                          </div>
                        </td>

                        {/* Tổng Thực tế / Chỉ tiêu & % Hoàn thành tổng */}
                        <td className="py-3 px-4 text-center font-mono">
                          <div className="font-black text-slate-900 text-sm">
                            {item.totalActual.toLocaleString()} <span className="text-slate-400 font-normal">/</span> {item.totalTarget > 0 ? item.totalTarget.toLocaleString() : '--'}
                          </div>
                          {item.totalTarget > 0 && (
                            <span className={`text-[11px] font-bold px-2 py-0.5 rounded-full border inline-block mt-0.5 ${
                              isCompleted 
                                ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
                                : 'bg-amber-50 text-amber-800 border-amber-200'
                            }`}>
                              {item.rateTotal}% tiến độ
                            </span>
                          )}
                        </td>

                        {/* Chỉ tiêu theo ca từ Sheet */}
                        <td className="py-3 px-4 text-center font-mono text-[11px] text-slate-600">
                          <div>Sáng: <b>{item.clsSang + item.thuocSang}</b> (CLS: {item.clsSang}, Thuốc: {item.thuocSang})</div>
                          <div>Chiều: <b>{item.clsChieu + item.thuocChieu}</b> (CLS: {item.clsChieu}, Thuốc: {item.thuocChieu})</div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ============================================================ */}
      {/* 5. TAB QUẢN LÝ CẤU HÌNH KIOSK (KIOSK_CONFIG) */}
      {/* ============================================================ */}
      {activeTab === 'KIOSK_CONFIG' && (
        <div className="space-y-5">
          <div className="p-6 rounded-3xl border border-slate-200 bg-white shadow-xs">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-base font-extrabold text-slate-900 uppercase tracking-tight">
                    QUẢN LÝ DANH SÁCH &amp; CẤU HÌNH MÁY KIOSK
                  </h2>
                  <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-teal-50 text-teal-800 border border-teal-200 font-mono">
                    {locationsList.length} Máy
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-1">
                  Chỉnh sửa tên máy, vị trí khu vực và công năng (<b>Tự thực hiện</b> hoặc <b>CSKH hỗ trợ</b>). Thay đổi sẽ được lưu và áp dụng ngay vào Dashboard!
                </p>
              </div>

              <div className="flex items-center gap-2.5">
                <button
                  onClick={handleResetKioskToDefault}
                  className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold transition cursor-pointer shadow-2xs"
                  title="Xóa tùy chỉnh và tải lại danh sách gốc từ Sheet"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>Đặt Lại Từ Sheet</span>
                </button>

                <button
                  onClick={handleOpenAddKiosk}
                  className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-bold text-white bg-teal-600 hover:bg-teal-700 shadow-sm shadow-teal-700/20 transition cursor-pointer"
                >
                  <Plus className="w-4 h-4" />
                  <span>+ Thêm Kiosk Mới</span>
                </button>
              </div>
            </div>

            {/* Thống kê nhanh công năng máy */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-4 pt-1">
              <div className="p-3.5 rounded-2xl bg-slate-50 border border-slate-200">
                <span className="text-[11px] font-bold text-slate-500 uppercase">Tổng Kiosk</span>
                <p className="text-2xl font-black font-mono text-slate-900 mt-1">{locationsList.length}</p>
              </div>

              <div className="p-3.5 rounded-2xl bg-purple-50/70 border border-purple-200">
                <span className="text-[11px] font-bold text-purple-700 uppercase">Tự Thực Hiện</span>
                <p className="text-2xl font-black font-mono text-purple-700 mt-1">{countSelfKiosks}</p>
                <span className="text-[10px] text-purple-600">Bệnh nhân tự làm</span>
              </div>

              <div className="p-3.5 rounded-2xl bg-blue-50/70 border border-blue-200">
                <span className="text-[11px] font-bold text-blue-700 uppercase">CSKH Hỗ Trợ</span>
                <p className="text-2xl font-black font-mono text-blue-700 mt-1">{locationsList.length - countSelfKiosks}</p>
                <span className="text-[10px] text-blue-600">Có nhân sự đứng máy</span>
              </div>

              <div className="p-3.5 rounded-2xl bg-emerald-50/70 border border-emerald-200">
                <span className="text-[11px] font-bold text-emerald-700 uppercase">Vị Trí Đặt Máy</span>
                <p className="text-2xl font-black font-mono text-emerald-700 mt-1">
                  {new Set(locationsList.map(k => k.area)).size}
                </p>
                <span className="text-[10px] text-emerald-600">Khu vực khoa phòng</span>
              </div>
            </div>

            {/* Thanh tìm kiếm & lọc công năng */}
            <div className="mt-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-600">Lọc công năng:</span>
                <div className="inline-flex p-0.5 rounded-xl bg-slate-100 border border-slate-200 text-xs">
                  <button
                    onClick={() => setKioskConfigFilter('ALL')}
                    className={`px-3 py-1 rounded-lg font-bold transition cursor-pointer ${
                      kioskConfigFilter === 'ALL' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
                    }`}
                  >
                    Tất cả ({locationsList.length})
                  </button>
                  <button
                    onClick={() => setKioskConfigFilter('Tự thực hiện')}
                    className={`px-3 py-1 rounded-lg font-bold transition cursor-pointer ${
                      kioskConfigFilter === 'Tự thực hiện' ? 'bg-purple-600 text-white shadow-2xs' : 'text-purple-700'
                    }`}
                  >
                    Tự thực hiện ({countSelfKiosks})
                  </button>
                  <button
                    onClick={() => setKioskConfigFilter('CSKH hỗ trợ')}
                    className={`px-3 py-1 rounded-lg font-bold transition cursor-pointer ${
                      kioskConfigFilter === 'CSKH hỗ trợ' ? 'bg-blue-600 text-white shadow-2xs' : 'text-blue-700'
                    }`}
                  >
                    CSKH hỗ trợ ({locationsList.length - countSelfKiosks})
                  </button>
                </div>
              </div>

              <div className="relative w-full sm:w-64">
                <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Tìm theo tên máy, khu vực..."
                  value={kioskSearchTerm}
                  onChange={(e) => setKioskSearchTerm(e.target.value)}
                  className="w-full text-xs pl-8 pr-3 py-1.5 rounded-xl border border-slate-300 bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-medium"
                />
              </div>
            </div>

            {/* BẢNG DANH SÁCH MÁY KIOSK */}
            <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-200">
              <table className="w-full text-left text-xs text-slate-700">
                <thead className="border-b border-slate-200 bg-slate-100/80 text-slate-800 text-[11px] font-bold uppercase tracking-wider">
                  <tr>
                    <th className="py-3 px-3 w-12 text-center">STT</th>
                    <th className="py-3 px-4">Tên Máy Kiosk</th>
                    <th className="py-3 px-4">Vị Trí / Khu Vực</th>
                    <th className="py-3 px-4 text-center">Công Năng / Cấu Hình</th>
                    <th className="py-3 px-4 text-center">Lượt GD Hôm Nay</th>
                    <th className="py-3 px-4 text-right">Thao Tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredKioskList.map((item, i) => {
                    const normMachine = normalizeKioskName(item.machine);
                    const kioskStat = analytics.kiosks.find(k => normalizeKioskName(k.machine) === normMachine);
                    const txCount = kioskStat?.count || 0;
                    const isSelf = item.config === 'Tự thực hiện';

                    return (
                      <tr key={item.machine} className="hover:bg-slate-50 transition">
                        <td className="py-3 px-3 text-center text-slate-400 font-mono">{i + 1}</td>
                        <td className="py-3 px-4 font-mono font-black text-sm text-teal-800">
                          {item.machine}
                        </td>
                        <td className="py-3 px-4 font-medium text-slate-800">
                          {item.area}
                        </td>
                        <td className="py-3 px-4 text-center">
                          {isSelf ? (
                            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold bg-purple-100 text-purple-900 border border-purple-300 shadow-2xs">
                              <Sparkles className="w-3.5 h-3.5 text-purple-600" />
                              Tự thực hiện
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-blue-50 text-blue-800 border border-blue-200">
                              <Users className="w-3.5 h-3.5 text-blue-600" />
                              CSKH hỗ trợ
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-4 text-center font-mono font-bold">
                          {txCount > 0 ? (
                            <span className="text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">
                              {txCount.toLocaleString()} lượt
                            </span>
                          ) : (
                            <span className="text-slate-400 font-normal">--</span>
                          )}
                        </td>
                        <td className="py-3 px-4 text-right">
                          <div className="flex items-center justify-end gap-1.5">
                            <button
                              onClick={() => handleOpenEditKiosk(item)}
                              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 font-bold text-xs transition cursor-pointer shadow-2xs"
                              title="Sửa tên, vị trí hoặc công năng"
                            >
                              <Edit3 className="w-3.5 h-3.5 text-teal-600" />
                              <span>Sửa</span>
                            </button>

                            <button
                              onClick={() => handleDeleteKiosk(item.machine)}
                              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg border border-rose-200 bg-rose-50/60 hover:bg-rose-100 text-rose-700 font-bold text-xs transition cursor-pointer"
                              title="Xóa máy này"
                            >
                              <Trash2 className="w-3.5 h-3.5 text-rose-600" />
                              <span>Xóa</span>
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ============================================================ */}
      {/* 6. TAB LOGS: NHẬT KÝ CHI TIẾT SÁNG RÕ */}
      {/* ============================================================ */}
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

      {/* ============================================================ */}
      {/* MODAL THÊM / SỬA CẤU HÌNH KIOSK (POPUP CHUYÊN NGHIỆP) */}
      {/* ============================================================ */}
      {showKioskModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="rounded-3xl border border-slate-200 bg-white text-slate-800 shadow-2xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3 border-slate-100">
              <div className="flex items-center gap-2">
                <Cpu className="w-5 h-5 text-teal-600" />
                <h3 className="font-extrabold text-sm text-slate-900">
                  {kioskFormData.originalMachine ? `Chỉnh Sửa Kiosk: ${kioskFormData.originalMachine}` : 'Thêm Máy Kiosk Mới'}
                </h3>
              </div>
              <button 
                onClick={() => setShowKioskModal(false)} 
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveKiosk} className="space-y-4 text-xs">
              <div>
                <label className="block font-bold mb-1.5 text-slate-700">
                  Tên máy Kiosk <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  required
                  placeholder="Ví dụ: KIOS17, KIOS99, KIOSK-23..."
                  value={kioskFormData.machine}
                  onChange={(e) => setKioskFormData({ ...kioskFormData, machine: e.target.value.toUpperCase() })}
                  className="w-full p-2.5 border border-slate-300 rounded-xl font-mono text-xs bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-bold"
                />
                <span className="text-[10px] text-slate-400 mt-0.5 block">Hệ thống sẽ tự động viết hoa và chuẩn hóa tên</span>
              </div>

              <div>
                <label className="block font-bold mb-1.5 text-slate-700">
                  Vị trí / Khu vực đặt máy <span className="text-rose-500">*</span>
                </label>
                <select
                  value={kioskFormData.area}
                  onChange={(e) => setKioskFormData({ ...kioskFormData, area: e.target.value })}
                  className="w-full p-2.5 border border-slate-300 rounded-xl text-xs bg-slate-50 text-slate-900 focus:outline-none focus:ring-1 focus:ring-teal-500 font-semibold cursor-pointer mb-2"
                >
                  {TARGET_AREAS.map(a => (
                    <option key={a} value={a}>{a}</option>
                  ))}
                  <option value="Khác">-- Khu vực khác (Tự nhập bên dưới) --</option>
                </select>

                {kioskFormData.area === 'Khác' && (
                  <input
                    type="text"
                    required
                    placeholder="Nhập tên khu vực mới..."
                    onChange={(e) => setKioskFormData({ ...kioskFormData, area: e.target.value })}
                    className="w-full p-2.5 border border-slate-300 rounded-xl text-xs bg-white text-slate-900 font-medium"
                  />
                )}
              </div>

              <div>
                <label className="block font-bold mb-1.5 text-slate-700">
                  Công năng / Cấu hình vận hành <span className="text-rose-500">*</span>
                </label>
                <div className="grid grid-cols-2 gap-3">
                  <label className={`p-3 rounded-2xl border flex flex-col cursor-pointer transition ${
                    kioskFormData.config === 'CSKH hỗ trợ'
                      ? 'border-blue-400 bg-blue-50/70 text-blue-900 shadow-2xs'
                      : 'border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700'
                  }`}>
                    <div className="flex items-center justify-between mb-1">
                      <span className="font-extrabold text-xs">CSKH hỗ trợ</span>
                      <input
                        type="radio"
                        name="configRadio"
                        checked={kioskFormData.config === 'CSKH hỗ trợ'}
                        onChange={() => setKioskFormData({ ...kioskFormData, config: 'CSKH hỗ trợ' })}
                        className="text-blue-600 cursor-pointer"
                      />
                    </div>
                    <span className="text-[10px] text-slate-500">Nhân sự túc trực hướng dẫn người bệnh</span>
                  </label>

                  <label className={`p-3 rounded-2xl border flex flex-col cursor-pointer transition ${
                    kioskFormData.config === 'Tự thực hiện'
                      ? 'border-purple-400 bg-purple-50/70 text-purple-900 shadow-2xs'
                      : 'border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700'
                  }`}>
                    <div className="flex items-center justify-between mb-1">
                      <span className="font-extrabold text-xs flex items-center gap-1">
                        <Sparkles className="w-3 h-3 text-purple-600" />
                        Tự thực hiện
                      </span>
                      <input
                        type="radio"
                        name="configRadio"
                        checked={kioskFormData.config === 'Tự thực hiện'}
                        onChange={() => setKioskFormData({ ...kioskFormData, config: 'Tự thực hiện' })}
                        className="text-purple-600 cursor-pointer"
                      />
                    </div>
                    <span className="text-[10px] text-slate-500">Bệnh nhân tự thao tác độc lập 100%</span>
                  </label>
                </div>
              </div>

              <div className="flex justify-end gap-2.5 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowKioskModal(false)}
                  className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-100 cursor-pointer"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-xl text-xs font-bold cursor-pointer shadow-sm"
                >
                  Lưu Cấu Hình
                </button>
              </div>
            </form>
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
