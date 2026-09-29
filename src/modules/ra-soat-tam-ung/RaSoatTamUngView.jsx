import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle, Clock, 
  Download, Search, RefreshCw, Trash2, Layers, AlertTriangle, 
  ChevronLeft, ChevronRight, CreditCard
} from 'lucide-react';
import { 
  detectFileType, 
  mergeAndDeduplicateBC140, 
  mergeAndDeduplicateBoVe, 
  processReconciliationTamUng, 
  exportTamUngToExcel,
  formatCurrency,
  parseLargeBC140FileAsync,
  cleanStr,
  getRowVal
} from './engine';

export default function RaSoatTamUngView() {
  const [filesBoVe, setFilesBoVe] = useState([]);
  const [files140, setFiles140] = useState([]);

  const [mergedBoVe, setMergedBoVe] = useState({ rows: [], alerts: [] });
  const [merged140, setMerged140] = useState({ rows: [], alerts: [] });

  const [isProcessing, setIsProcessing] = useState(false);
  const [progressInfo, setProgressInfo] = useState(null);
  const [resultData, setResultData] = useState(null);
  const [showDuplicateModal, setShowDuplicateModal] = useState(false);

  const [activeTab, setActiveTab] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedKhoa, setSelectedKhoa] = useState('ALL');

  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 15;

  const targetMaBNSet = useMemo(() => {
    const s = new Set();
    mergedBoVe.rows.forEach(r => {
      const bn = cleanStr(getRowVal(r, ['MÃ BỆNH NHÂN', 'Mã bệnh nhân', 'Mã BN']));
      if (bn) s.add(bn);
    });
    return s;
  }, [mergedBoVe.rows]);

  const handleUploadBoVe = async (event) => {
    const selectedFiles = Array.from(event.target.files || []);
    if (selectedFiles.length === 0) return;

    setIsProcessing(true);
    const parsedList = [];

    for (const file of selectedFiles) {
      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        parsedList.push({
          id: `${file.name}_${file.size}_${Date.now()}`,
          name: file.name,
          size: (file.size / 1024).toFixed(1) + ' KB',
          rows
        });
      } catch (err) {
        console.error('Lỗi đọc file:', file.name, err);
      }
    }

    const updated = [...filesBoVe, ...parsedList];
    setFilesBoVe(updated);
    const res = mergeAndDeduplicateBoVe(updated.map(f => ({ fileName: f.name, rows: f.rows })));
    setMergedBoVe({ rows: res.mergedRows, alerts: res.duplicateAlerts });
    setIsProcessing(false);
    event.target.value = '';
  };

  const handleUploadBC140 = async (event) => {
    const selectedFiles = Array.from(event.target.files || []);
    if (selectedFiles.length === 0) return;

    setIsProcessing(true);
    const parsedList = [];

    for (const file of selectedFiles) {
      try {
        const parsed = await parseLargeBC140FileAsync(file, targetMaBNSet, (p) => {
          setProgressInfo(p);
        });

        parsedList.push({
          id: `${file.name}_${file.size}_${Date.now()}`,
          name: file.name,
          size: (file.size / 1024).toFixed(1) + ' KB',
          totalRows: parsed.totalRows,
          rows: parsed.rows
        });
      } catch (err) {
        console.error('Lỗi đọc file BC140:', file.name, err);
      }
    }

    const updated = [...files140, ...parsedList];
    setFiles140(updated);
    const res = mergeAndDeduplicateBC140(updated.map(f => ({ fileName: f.name, rows: f.rows })));
    setMerged140({ rows: res.mergedRows, alerts: res.duplicateAlerts });

    setProgressInfo(null);
    setIsProcessing(false);
    event.target.value = '';
  };

  const handleRemoveFile = (id, targetType) => {
    if (targetType === 'BO_VE') {
      const updated = filesBoVe.filter(f => f.id !== id);
      setFilesBoVe(updated);
      const res = mergeAndDeduplicateBoVe(updated.map(f => ({ fileName: f.name, rows: f.rows })));
      setMergedBoVe({ rows: res.mergedRows, alerts: res.duplicateAlerts });
    } else {
      const updated = files140.filter(f => f.id !== id);
      setFiles140(updated);
      const res = mergeAndDeduplicateBC140(updated.map(f => ({ fileName: f.name, rows: f.rows })));
      setMerged140({ rows: res.mergedRows, alerts: res.duplicateAlerts });
    }
  };

  const handleRunReconcile = () => {
    if (mergedBoVe.rows.length === 0 || merged140.rows.length === 0) return;

    setIsProcessing(true);
    setTimeout(() => {
      const output = processReconciliationTamUng(mergedBoVe.rows, merged140.rows);
      setResultData(output);
      setIsProcessing(false);
      setCurrentPage(1);
    }, 200);
  };

  const handleResetAll = () => {
    setFilesBoVe([]);
    setFiles140([]);
    setMergedBoVe({ rows: [], alerts: [] });
    setMerged140({ rows: [], alerts: [] });
    setResultData(null);
    setProgressInfo(null);
    setCurrentPage(1);
  };

  const departmentList = useMemo(() => {
    if (!resultData) return [];
    const set = new Set(resultData.items.map(i => i.khoaPhongBoVe).filter(Boolean));
    return Array.from(set).sort();
  }, [resultData]);

  const totalDuplicates = mergedBoVe.alerts.length + merged140.alerts.length;

  const filteredItems = useMemo(() => {
    if (!resultData) return [];
    return resultData.items.filter(item => {
      if (activeTab === 'MATCHED' && item.status !== 'MATCHED_RE_ADVANCE') return false;
      if (activeTab === 'TRANSFER_REFUND' && item.status !== 'MATCHED_TRANSFER_REFUND') return false;
      if (activeTab === 'DIFF_DATE' && item.status !== 'MATCHED_DIFF_DATE') return false;
      if (activeTab === 'NOT_FOUND' && item.status !== 'NOT_FOUND') return false;

      if (selectedKhoa !== 'ALL' && item.khoaPhongBoVe !== selectedKhoa) return false;

      if (searchTerm) {
        const q = searchTerm.toLowerCase().trim();
        const matchBN = String(item.maBN || '').toLowerCase().includes(q);
        const matchName = String(item.hoTen || '').toLowerCase().includes(q);
        const matchBL = String(item.soBienLai || '').toLowerCase().includes(q);
        const matchTN = String(item.maThuNgan || '').toLowerCase().includes(q) || String(item.tenThuNgan || '').toLowerCase().includes(q);
        const matchKhoa = String(item.khoaPhongBoVe || '').toLowerCase().includes(q);
        return matchBN || matchName || matchBL || matchTN || matchKhoa;
      }
      return true;
    });
  }, [resultData, activeTab, selectedKhoa, searchTerm]);

  const totalPages = Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = filteredItems.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const handleExport = () => {
    if (!resultData || resultData.items.length === 0) return;
    exportTamUngToExcel(resultData.items, resultData.summary);
  };

  return (
    <div className="space-y-6 pb-12">
      {/* HEADER */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-5 rounded-xl border border-slate-200/80 shadow-xs">
        <div>
          <h1 className="text-xl font-bold text-slate-900 tracking-tight">
            Rà Soát Tạm Ứng Lại
          </h1>
        </div>

        <div className="flex items-center gap-2">
          {resultData && (
            <button
              onClick={handleExport}
              className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold shadow-xs transition cursor-pointer"
            >
              <Download className="w-4 h-4" />
              Xuất Excel
            </button>
          )}
          {(filesBoVe.length > 0 || files140.length > 0) && (
            <button
              onClick={handleResetAll}
              className="inline-flex items-center gap-1.5 px-3 py-2 border border-slate-300 hover:bg-slate-100 text-slate-700 rounded-lg text-xs font-medium transition cursor-pointer"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              Làm mới
            </button>
          )}
        </div>
      </div>

      {/* THANH TIẾN TRÌNH XỬ LÝ 200K DÒNG */}
      {progressInfo && (
        <div className="bg-white border border-teal-200 rounded-xl p-4 shadow-sm space-y-2">
          <div className="flex items-center justify-between text-xs font-semibold text-teal-900">
            <span>{progressInfo.text}</span>
            <span>{progressInfo.percent}%</span>
          </div>
          <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
            <div 
              className="h-full bg-teal-600 transition-all duration-150 rounded-full"
              style={{ width: `${progressInfo.percent}%` }}
            />
          </div>
        </div>
      )}

      {/* CẢNH BÁO TRÙNG LẶP */}
      {totalDuplicates > 0 && (
        <div className="bg-amber-50 border border-amber-300 rounded-xl p-4 text-amber-900 flex items-start justify-between shadow-xs">
          <div className="flex items-start gap-3">
            <AlertTriangle className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
            <div>
              <h4 className="text-xs font-bold uppercase tracking-wide">
                Phát hiện {totalDuplicates} dòng dữ liệu trùng lặp
              </h4>
            </div>
          </div>
          <button
            onClick={() => setShowDuplicateModal(!showDuplicateModal)}
            className="text-xs font-semibold underline text-amber-800 hover:text-amber-950 flex-shrink-0 ml-4 cursor-pointer"
          >
            {showDuplicateModal ? 'Thu gọn' : 'Xem chi tiết'}
          </button>
        </div>
      )}

      {/* CHI TIẾT TRÙNG LẶP */}
      {showDuplicateModal && totalDuplicates > 0 && (
        <div className="bg-white border border-amber-200 rounded-xl p-4 shadow-sm space-y-3">
          <div className="max-h-48 overflow-y-auto space-y-1.5 text-xs text-slate-600">
            {mergedBoVe.alerts.map((a, i) => (
              <div key={`bv_${i}`} className="p-2 bg-amber-50/50 rounded border border-amber-100 flex items-center justify-between">
                <span>[BC KH Bỏ Về] Mã BN: <b>{a.maBN}</b> ({a.ngay}, {a.khoaPhong})</span>
                <span className="text-slate-500 font-mono text-[11px]">{a.fileName}</span>
              </div>
            ))}
            {merged140.alerts.map((a, i) => (
              <div key={`bc_${i}`} className="p-2 bg-amber-50/50 rounded border border-amber-100 flex items-center justify-between">
                <span>[BC140] Biên lai: <b>{a.receiptNumber}</b> - Mã BN: <b>{a.maBN}</b></span>
                <span className="text-slate-500 font-mono text-[11px]">{a.fileName}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* UPLOAD FILES */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
        {/* BOX 1: BC KH BỎ VỀ */}
        <div className="bg-white p-5 rounded-xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                <FileSpreadsheet className="w-4 h-4 text-blue-600" />
                1. Báo Cáo KH Bỏ Về
              </span>
              <span className="text-[11px] px-2 py-0.5 bg-blue-50 text-blue-700 font-semibold rounded-full">
                {filesBoVe.length} file ({mergedBoVe.rows.length} hồ sơ)
              </span>
            </div>

            <label className="border-2 border-dashed border-blue-200 hover:border-blue-400 bg-blue-50/30 hover:bg-blue-50/60 rounded-xl p-4 flex flex-col items-center justify-center cursor-pointer transition text-center group">
              <UploadCloud className="w-7 h-7 text-blue-500 group-hover:scale-110 transition duration-150 mb-1" />
              <span className="text-xs font-semibold text-blue-700">
                Chọn tệp Excel
              </span>
              <input 
                type="file" 
                multiple 
                accept=".xlsx, .xls" 
                className="hidden" 
                onChange={handleUploadBoVe}
              />
            </label>

            {filesBoVe.length > 0 && (
              <div className="mt-3 space-y-1.5 max-h-36 overflow-y-auto pr-1">
                {filesBoVe.map((file) => (
                  <div key={file.id} className="flex items-center justify-between p-2 bg-slate-50 rounded-lg border border-slate-200/60 text-xs">
                    <div className="flex items-center gap-2 truncate mr-2">
                      <FileSpreadsheet className="w-3.5 h-3.5 text-blue-600 flex-shrink-0" />
                      <span className="truncate font-medium text-slate-800">{file.name}</span>
                      <span className="text-[10px] text-slate-500">({file.rows.length})</span>
                    </div>
                    <button
                      onClick={() => handleRemoveFile(file.id, 'BO_VE')}
                      className="text-slate-400 hover:text-red-500 p-1 transition cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* BOX 2: BC140 */}
        <div className="bg-white p-5 rounded-xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                <CreditCard className="w-4 h-4 text-teal-600" />
                2. Báo Cáo BC140
              </span>
              <span className="text-[11px] px-2 py-0.5 bg-teal-50 text-teal-700 font-semibold rounded-full">
                {files140.length} file ({merged140.rows.length} GD)
              </span>
            </div>

            <label className="border-2 border-dashed border-teal-200 hover:border-teal-400 bg-teal-50/30 hover:bg-teal-50/60 rounded-xl p-4 flex flex-col items-center justify-center cursor-pointer transition text-center group">
              <UploadCloud className="w-7 h-7 text-teal-500 group-hover:scale-110 transition duration-150 mb-1" />
              <span className="text-xs font-semibold text-teal-700">
                Chọn tệp BC140 (hỗ trợ file 200k dòng)
              </span>
              <input 
                type="file" 
                multiple 
                accept=".xlsx, .xls" 
                className="hidden" 
                onChange={handleUploadBC140}
              />
            </label>

            {files140.length > 0 && (
              <div className="mt-3 space-y-1.5 max-h-36 overflow-y-auto pr-1">
                {files140.map((file) => (
                  <div key={file.id} className="flex items-center justify-between p-2 bg-slate-50 rounded-lg border border-slate-200/60 text-xs">
                    <div className="flex items-center gap-2 truncate mr-2">
                      <FileSpreadsheet className="w-3.5 h-3.5 text-teal-600 flex-shrink-0" />
                      <span className="truncate font-medium text-slate-800">{file.name}</span>
                      <span className="text-[10px] text-slate-500">
                        {file.totalRows ? `(${file.totalRows.toLocaleString()} dòng)` : `(${file.rows.length})`}
                      </span>
                    </div>
                    <button
                      onClick={() => handleRemoveFile(file.id, 'BC140')}
                      className="text-slate-400 hover:text-red-500 p-1 transition cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ACTION BUTTON */}
      <div className="flex justify-center">
        <button
          onClick={handleRunReconcile}
          disabled={mergedBoVe.rows.length === 0 || merged140.rows.length === 0 || isProcessing}
          className={`px-8 py-3 rounded-xl font-bold text-xs uppercase tracking-wider flex items-center gap-2.5 shadow-sm transition ${
            mergedBoVe.rows.length > 0 && merged140.rows.length > 0 && !isProcessing
              ? 'bg-gradient-to-r from-teal-600 to-emerald-600 text-white hover:from-teal-700 hover:to-emerald-700 cursor-pointer'
              : 'bg-slate-200 text-slate-400 cursor-not-allowed'
          }`}
        >
          {isProcessing ? (
            <>
              <RefreshCw className="w-4 h-4 animate-spin" />
              Đang đối chiếu...
            </>
          ) : (
            <>
              <Layers className="w-4 h-4" />
              Thực Hiện Rà Soát
            </>
          )}
        </button>
      </div>

      {/* DASHBOARD KẾT QUẢ */}
      {resultData && (
        <div className="space-y-5">
          {/* STATS CARDS */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-3.5">
            <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs">
              <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Tổng Hồ Sơ</span>
              <div className="mt-1 flex items-baseline justify-between">
                <span className="text-2xl font-bold text-slate-900">{resultData.summary.total}</span>
              </div>
            </div>

            <div className="bg-white p-4 rounded-xl border border-emerald-200/80 shadow-xs bg-emerald-50/20">
              <span className="text-[11px] font-semibold text-emerald-700 uppercase tracking-wider">Khớp Tạm Ứng</span>
              <div className="mt-1 flex items-baseline justify-between">
                <span className="text-2xl font-bold text-emerald-700">{resultData.summary.countMatched}</span>
              </div>
              <p className="text-[10px] text-emerald-600 mt-1 truncate">
                {formatCurrency(resultData.summary.totalReAdvanceAmount)}
              </p>
            </div>

            <div className="bg-white p-4 rounded-xl border border-purple-200/80 shadow-xs bg-purple-50/20">
              <span className="text-[11px] font-semibold text-purple-700 uppercase tracking-wider">Hoàn Chuyển Khoản</span>
              <div className="mt-1 flex items-baseline justify-between">
                <span className="text-2xl font-bold text-purple-700">{resultData.summary.countTransferRefund}</span>
              </div>
              <p className="text-[10px] text-purple-600 mt-1 truncate">
                {formatCurrency(resultData.summary.totalRefundAmount)}
              </p>
            </div>

            <div className="bg-white p-4 rounded-xl border border-amber-200/80 shadow-xs bg-amber-50/20">
              <span className="text-[11px] font-semibold text-amber-700 uppercase tracking-wider">Lệch Ngày</span>
              <div className="mt-1 flex items-baseline justify-between">
                <span className="text-2xl font-bold text-amber-700">{resultData.summary.countDiffDate}</span>
              </div>
            </div>

            <div className="bg-white p-4 rounded-xl border border-rose-200/80 shadow-xs bg-rose-50/20">
              <span className="text-[11px] font-semibold text-rose-700 uppercase tracking-wider">Chưa Tìm Thấy</span>
              <div className="mt-1 flex items-baseline justify-between">
                <span className="text-2xl font-bold text-rose-700">{resultData.summary.countNotFound}</span>
              </div>
            </div>
          </div>

          {/* FILTER TOOLBAR */}
          <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-1.5 border-b md:border-b-0 pb-2 md:pb-0 border-slate-100">
              <button
                onClick={() => { setActiveTab('ALL'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer ${
                  activeTab === 'ALL'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                Tất cả ({resultData.summary.total})
              </button>
              <button
                onClick={() => { setActiveTab('MATCHED'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer ${
                  activeTab === 'MATCHED'
                    ? 'bg-emerald-600 text-white'
                    : 'text-emerald-700 hover:bg-emerald-50'
                }`}
              >
                Khớp Tạm Ứng ({resultData.summary.countMatched})
              </button>
              <button
                onClick={() => { setActiveTab('TRANSFER_REFUND'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer ${
                  activeTab === 'TRANSFER_REFUND'
                    ? 'bg-purple-600 text-white'
                    : 'text-purple-700 hover:bg-purple-50'
                }`}
              >
                Hoàn CK ({resultData.summary.countTransferRefund})
              </button>
              <button
                onClick={() => { setActiveTab('DIFF_DATE'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer ${
                  activeTab === 'DIFF_DATE'
                    ? 'bg-amber-600 text-white'
                    : 'text-amber-700 hover:bg-amber-50'
                }`}
              >
                Lệch Ngày ({resultData.summary.countDiffDate})
              </button>
              <button
                onClick={() => { setActiveTab('NOT_FOUND'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer ${
                  activeTab === 'NOT_FOUND'
                    ? 'bg-rose-600 text-white'
                    : 'text-rose-700 hover:bg-rose-50'
                }`}
              >
                Chưa Thấy ({resultData.summary.countNotFound})
              </button>
            </div>

            <div className="flex items-center gap-2">
              <select
                value={selectedKhoa}
                onChange={(e) => { setSelectedKhoa(e.target.value); setCurrentPage(1); }}
                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-700 focus:outline-none focus:ring-1 focus:ring-teal-500 max-w-[160px] truncate"
              >
                <option value="ALL">Khoa phòng</option>
                {departmentList.map(k => (
                  <option key={k} value={k}>{k}</option>
                ))}
              </select>

              <div className="relative">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Mã BN, tên, BL, thu ngân..."
                  value={searchTerm}
                  onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                  className="pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:ring-1 focus:ring-teal-500 w-52"
                />
              </div>
            </div>
          </div>

          {/* BẢNG KẾT QUẢ */}
          <div className="bg-white rounded-xl border border-slate-200/80 shadow-xs overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-slate-600">
                <thead className="bg-slate-100/80 text-slate-700 uppercase font-semibold text-[11px] tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-3 w-10 text-center">STT</th>
                    <th className="py-3 px-3">Ngày</th>
                    <th className="py-3 px-3">Mã BN</th>
                    <th className="py-3 px-4">Họ và Tên</th>
                    <th className="py-3 px-3">Khoa Phòng</th>
                    <th className="py-3 px-3 text-right">Số Tiền</th>
                    <th className="py-3 px-3">Số BL</th>
                    <th className="py-3 px-3">Thu Ngân</th>
                    <th className="py-3 px-3">Thời Gian</th>
                    <th className="py-3 px-3">Hoàn CK</th>
                    <th className="py-3 px-3">Trạng Thái</th>
                    <th className="py-3 px-4">Chi Tiết</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginatedItems.length === 0 ? (
                    <tr>
                      <td colSpan="12" className="py-8 text-center text-slate-400">
                        Không có dữ liệu phù hợp.
                      </td>
                    </tr>
                  ) : (
                    paginatedItems.map((item) => (
                      <tr key={item.stt} className="hover:bg-slate-50/70 transition">
                        <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">{item.stt}</td>
                        <td className="py-2.5 px-3 font-mono text-slate-700">{item.ngayBoVe || '--/--/----'}</td>
                        <td className="py-2.5 px-3 font-mono font-semibold text-slate-900">{item.maBN}</td>
                        <td className="py-2.5 px-4 font-medium text-slate-800">{item.hoTen}</td>
                        <td className="py-2.5 px-3 text-slate-600 max-w-[160px] truncate" title={item.khoaPhongBoVe}>
                          {item.khoaPhongBoVe}
                        </td>
                        <td className="py-2.5 px-3 text-right font-bold text-slate-900 font-mono">
                          {formatCurrency(item.targetAmount)}
                        </td>
                        <td className="py-2.5 px-3 font-mono font-bold text-teal-800">
                          {item.soBienLai || <span className="text-slate-300 font-normal">--</span>}
                        </td>
                        <td className="py-2.5 px-3 text-slate-700">
                          {item.maThuNgan ? (
                            <div>
                              <span className="font-semibold text-slate-900">{item.maThuNgan}</span>
                              {item.tenThuNgan && <span className="text-[11px] text-slate-500 block">{item.tenThuNgan}</span>}
                            </div>
                          ) : (
                            <span className="text-slate-300">--</span>
                          )}
                        </td>
                        <td className="py-2.5 px-3 font-mono text-slate-600">
                          {item.thoiGianBL ? (
                            <div>
                              <span className="font-semibold text-slate-800">{item.thoiGianBL}</span>
                              {item.ngayBL && <span className="text-[10px] text-slate-400 block">{item.ngayBL}</span>}
                            </div>
                          ) : (
                            <span className="text-slate-300">--:--</span>
                          )}
                        </td>
                        <td className="py-2.5 px-3">
                          {item.hoanChuyenKhoan ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-bold bg-purple-100 text-purple-800 border border-purple-200">
                              {item.hoanChuyenKhoan}
                            </span>
                          ) : (
                            <span className="text-slate-300">--</span>
                          )}
                        </td>
                        <td className="py-2.5 px-3">
                          {item.status === 'MATCHED_RE_ADVANCE' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800">
                              <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                              {item.statusText}
                            </span>
                          )}
                          {item.status === 'MATCHED_TRANSFER_REFUND' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-purple-100 text-purple-800">
                              <CreditCard className="w-3 h-3 text-purple-600" />
                              Hoàn CK
                            </span>
                          )}
                          {item.status === 'MATCHED_DIFF_DATE' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-800">
                              <Clock className="w-3 h-3 text-amber-600" />
                              Lệch Ngày
                            </span>
                          )}
                          {item.status === 'NOT_FOUND' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-rose-100 text-rose-800">
                              <AlertCircle className="w-3 h-3 text-rose-600" />
                              Chưa Thấy
                            </span>
                          )}
                        </td>
                        <td className="py-2.5 px-4 text-slate-500 text-[11px] max-w-[220px] truncate" title={item.auditNotes}>
                          {item.auditNotes}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* PHÂN TRANG */}
            <div className="p-3 bg-slate-50 border-t border-slate-200/80 flex items-center justify-between text-xs text-slate-500">
              <span>
                <b>{filteredItems.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}</b> - <b>{Math.min(currentPage * pageSize, filteredItems.length)}</b> / <b>{filteredItems.length}</b>
              </span>

              <div className="flex items-center gap-1">
                <button
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="p-1 rounded border border-slate-300 hover:bg-slate-200 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <span className="px-2 font-medium">Trang {currentPage} / {totalPages}</span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="p-1 rounded border border-slate-300 hover:bg-slate-200 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
