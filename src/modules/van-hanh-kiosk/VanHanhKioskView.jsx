import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Monitor, CreditCard, Clock, RefreshCw, Download, 
  Search, Building2, QrCode, FileText, Pill, 
  Settings, X, Sparkles, AlertCircle, ChevronLeft, 
  ChevronRight, Layers, Activity, BarChart3, Plus, 
  Edit3, Trash2, RotateCcw, Check, Users, Cpu
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
    totalDailyTarget: 3150,
    totalTargetCLS: 2850,
    totalTargetThuoc: 300
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

  // Phân tích dữ liệu Dashboard
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

            {/* TAB MỚI: QUẢN LÝ CẤU HÌNH KIOSK */}
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
      {/* 2. TAB QUẢN LÝ CẤU HÌNH KIOSK (SỬA, THÊM, XÓA, ĐỔI CÔNG NĂNG) */}
      {/* ============================================================ */}
      {activeTab === 'KIOSK_CONFIG' && (
        <div className="space-y-5">
          {/* Header Quản Lý Kiosk & Các Nút Thao Tác */}
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
                              {txCount} lượt
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
      {/* CÁC PHÂN HỆ DASHBOARD KHÁC (OVERVIEW / SELF_SERVICE / DEPARTMENTS / LOGS) */}
      {/* ============================================================ */}
      {activeTab !== 'KIOSK_CONFIG' && (
        <>
          {/* DẢI 4 THẺ METRICS SÁNG TƯƠI, ĐỘ TƯƠNG PHẢN CAO */}
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

          {/* 3. PHÂN HỆ ĐỘC QUYỀN: MÁY TỰ THỰC HIỆN */}
          <div className="p-6 rounded-3xl border-2 border-purple-200 bg-gradient-to-r from-purple-50/80 via-white to-purple-50/40 shadow-xs space-y-4">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-purple-200/80 pb-4">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-purple-600 text-white flex items-center justify-center shadow-md shadow-purple-600/30">
                  <Sparkles className="w-5 h-5 text-white" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <h2 className="text-base font-black text-purple-950 uppercase tracking-wide">
                      PHÂN TÍCH CHUYÊN SÂU: CÁC MÁY CẤU HÌNH "TỰ THỰC HIỆN"
                    </h2>
                    <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase bg-purple-100 text-purple-800 border border-purple-300">
                      {countSelfKiosks} MÁY TỰ LÀM
                    </span>
                  </div>
                  <p className="text-xs text-purple-800 font-medium mt-0.5">
                    Tự động nhận diện tất cả máy có cấu hình "Tự thực hiện" do bạn thiết lập để tổng hợp số liệu
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

            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <div className="p-4 rounded-2xl border border-purple-200 bg-white shadow-2xs">
                <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500 block">
                  Tổng Lượt Tự Làm
                </span>
                <span className="text-3xl font-black font-mono text-purple-700 mt-1 block">
                  {analytics.summary.selfTotal.toLocaleString()}
                </span>
                <span className="text-[11px] text-slate-500 font-medium">Bệnh nhân tự thao tác</span>
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

          {/* 4. BIỂU ĐỒ PHÂN BỔ KHOA PHÒNG */}
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
        </>
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
