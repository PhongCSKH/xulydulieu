import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle, Clock, 
  Download, Search, RefreshCw, ArrowRight, ChevronLeft, ChevronRight
} from 'lucide-react';
import { detectFileType, processReconciliation, exportToExcel } from './engine';

export default function BV01ModuleView() {
  const [file121, setFile121] = useState(null);
  const [file140, setFile140] = useState(null);
  const [rows121, setRows121] = useState([]);
  const [rows140, setRows140] = useState([]);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [resultData, setResultData] = useState(null);
  
  // Filter & Search states
  const [activeTab, setActiveTab] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedKhoa, setSelectedKhoa] = useState('ALL');
  
  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 15;

  // Xử lý đọc file Excel vào JSON
  const handleFileUpload = (file, target) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const detectedType = detectFileType(firstSheet);
        const rows = XLSX.utils.sheet_to_json(firstSheet);

        if (target === '121' || detectedType === 'BC121') {
          setFile121(file);
          setRows121(rows);
        } else if (target === '140' || detectedType === 'BC140') {
          setFile140(file);
          setRows140(rows);
        }
      } catch (err) {
        console.error('Lỗi đọc file:', err);
        alert('Không thể đọc file Excel này.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Khởi chạy đối soát
  const handleRunReconcile = () => {
    if (rows121.length === 0 || rows140.length === 0) return;
    setIsProcessing(true);
    setTimeout(() => {
      const output = processReconciliation(rows121, rows140);
      setResultData(output);
      setIsProcessing(false);
      setCurrentPage(1);
    }, 200);
  };

  // Danh sách Khoa phòng
  const departmentList = useMemo(() => {
    if (!resultData) return [];
    const set = new Set(resultData.items.map(i => i.khoaPhong).filter(Boolean));
    return Array.from(set).sort();
  }, [resultData]);

  // Lọc dữ liệu
  const filteredItems = useMemo(() => {
    if (!resultData) return [];
    return resultData.items.filter(item => {
      if (activeTab !== 'ALL' && item.status !== activeTab) return false;
      if (selectedKhoa !== 'ALL' && item.khoaPhong !== selectedKhoa) return false;
      if (searchTerm) {
        const query = searchTerm.toLowerCase().trim();
        const matchBN = String(item.maBN || '').toLowerCase().includes(query);
        const matchName = String(item.hoTen || '').toLowerCase().includes(query);
        const matchKhoa = String(item.khoaPhong || '').toLowerCase().includes(query);
        const matchBL = String(item.soBienLai || '').toLowerCase().includes(query);
        return matchBN || matchName || matchKhoa || matchBL;
      }
      return true;
    });
  }, [resultData, activeTab, selectedKhoa, searchTerm]);

  // Phân trang
  const totalPages = Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = filteredItems.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  // Nạp dữ liệu mẫu
  const handleLoadSampleData = async () => {
    setIsProcessing(true);
    try {
      const res121 = await fetch('./samples/BC121_Mau.xlsx');
      const buf121 = await res121.arrayBuffer();
      const wb121 = XLSX.read(new Uint8Array(buf121), { type: 'array', cellDates: true });
      const s121 = wb121.Sheets[wb121.SheetNames[0]];
      const r121 = XLSX.utils.sheet_to_json(s121);
      setFile121({ name: 'BC121 02.09.xlsx', size: buf121.byteLength });
      setRows121(r121);

      const res140 = await fetch('./samples/BC140_Mau.xlsx');
      const buf140 = await res140.arrayBuffer();
      const wb140 = XLSX.read(new Uint8Array(buf140), { type: 'array', cellDates: true });
      const s140 = wb140.Sheets[wb140.SheetNames[0]];
      const r140 = XLSX.utils.sheet_to_json(s140);
      setFile140({ name: 'BC140 02-04.09.xlsx', size: buf140.byteLength });
      setRows140(r140);

      const output = processReconciliation(r121, r140);
      setResultData(output);
      setCurrentPage(1);
    } catch (err) {
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="space-y-5">
      {/* Header Chức Năng */}
      <div className="bg-white border border-slate-200 rounded-xl px-5 py-4 shadow-sm flex items-center justify-between">
        <h1 className="text-xl font-bold text-slate-900">Đối Soát Chốt BV01 Chưa Ra Bill</h1>
        
        <div className="flex items-center gap-2">
          <button
            onClick={handleLoadSampleData}
            disabled={isProcessing}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 transition"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isProcessing ? 'animate-spin' : ''}`} />
            Dữ liệu mẫu
          </button>

          {resultData && (
            <button
              onClick={() => exportToExcel(filteredItems)}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold text-white bg-teal-600 hover:bg-teal-700 transition shadow-sm"
            >
              <Download className="w-3.5 h-3.5" />
              Xuất Excel ({filteredItems.length})
            </button>
          )}
        </div>
      </div>

      {/* Vùng Tải Lên 2 Tệp */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* Khung File 1: BC121 */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center gap-2">
              <span className="w-7 h-7 rounded bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-xs border border-blue-200">
                121
              </span>
              <h3 className="text-sm font-semibold text-slate-900">Báo cáo 121</h3>
            </div>
            {file121 && (
              <span className="text-xs font-medium text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full border border-blue-200 font-mono">
                {rows121.length.toLocaleString()} dòng
              </span>
            )}
          </div>

          <label className={`border-2 border-dashed rounded-lg p-4 flex flex-col items-center justify-center cursor-pointer transition ${
            file121 ? 'border-blue-300 bg-blue-50/20' : 'border-slate-300 hover:border-blue-400 hover:bg-slate-50'
          }`}>
            <input 
              type="file" 
              accept=".xlsx,.xls,.csv" 
              className="hidden" 
              onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0], '121')}
            />
            {file121 ? (
              <div className="flex items-center gap-2.5 text-slate-700">
                <FileSpreadsheet className="w-6 h-6 text-blue-600 flex-shrink-0" />
                <div className="text-left">
                  <p className="text-xs font-medium text-slate-900 truncate max-w-xs">{file121.name}</p>
                  <p className="text-[11px] text-slate-400 font-mono">{(file121.size / 1024).toFixed(1)} KB</p>
                </div>
              </div>
            ) : (
              <div className="text-center py-1">
                <UploadCloud className="w-6 h-6 text-slate-400 mx-auto mb-1" />
                <p className="text-xs text-slate-600">Chọn tệp BC121</p>
              </div>
            )}
          </label>
        </div>

        {/* Khung File 2: BC140 */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center gap-2">
              <span className="w-7 h-7 rounded bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-xs border border-indigo-200">
                140
              </span>
              <h3 className="text-sm font-semibold text-slate-900">Báo cáo 140</h3>
            </div>
            {file140 && (
              <span className="text-xs font-medium text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-full border border-indigo-200 font-mono">
                {rows140.length.toLocaleString()} dòng
              </span>
            )}
          </div>

          <label className={`border-2 border-dashed rounded-lg p-4 flex flex-col items-center justify-center cursor-pointer transition ${
            file140 ? 'border-indigo-300 bg-indigo-50/20' : 'border-slate-300 hover:border-indigo-400 hover:bg-slate-50'
          }`}>
            <input 
              type="file" 
              accept=".xlsx,.xls,.csv" 
              className="hidden" 
              onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0], '140')}
            />
            {file140 ? (
              <div className="flex items-center gap-2.5 text-slate-700">
                <FileSpreadsheet className="w-6 h-6 text-indigo-600 flex-shrink-0" />
                <div className="text-left">
                  <p className="text-xs font-medium text-slate-900 truncate max-w-xs">{file140.name}</p>
                  <p className="text-[11px] text-slate-400 font-mono">{(file140.size / 1024).toFixed(1)} KB</p>
                </div>
              </div>
            ) : (
              <div className="text-center py-1">
                <UploadCloud className="w-6 h-6 text-slate-400 mx-auto mb-1" />
                <p className="text-xs text-slate-600">Chọn tệp BC140</p>
              </div>
            )}
          </label>
        </div>
      </div>

      {/* Nút Thực Hiện */}
      <div className="flex justify-end">
        <button
          onClick={handleRunReconcile}
          disabled={!file121 || !file140 || isProcessing}
          className={`inline-flex items-center gap-2 px-6 py-2 rounded-lg text-xs font-semibold transition shadow-sm ${
            !file121 || !file140
              ? 'bg-slate-200 text-slate-400 cursor-not-allowed'
              : isProcessing
              ? 'bg-teal-500 text-white cursor-wait'
              : 'bg-teal-600 hover:bg-teal-700 text-white'
          }`}
        >
          {isProcessing ? (
            <>
              <RefreshCw className="w-3.5 h-3.5 animate-spin" />
              Đang xử lý...
            </>
          ) : (
            <>
              Đối Soát
              <ArrowRight className="w-3.5 h-3.5" />
            </>
          )}
        </button>
      </div>

      {/* KHU VỰC KẾT QUẢ ĐỐI SOÁT */}
      {resultData && (
        <div className="space-y-4">
          {/* Thẻ Thống Kê Tổng Quan */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {/* Card 1 */}
            <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs text-slate-500 font-medium">Tổng hồ sơ BV01</span>
              <p className="text-2xl font-bold text-slate-900 mt-1 font-mono">{resultData.summary.total.toLocaleString()}</p>
            </div>

            {/* Card 2 */}
            <div className="bg-white border border-rose-200 rounded-xl p-4 shadow-sm bg-rose-50/20">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-rose-700">Chưa Ra Bill</span>
                <AlertCircle className="w-4 h-4 text-rose-600" />
              </div>
              <p className="text-2xl font-bold text-rose-600 mt-1 font-mono">{resultData.summary.countPendingBill.toLocaleString()}</p>
            </div>

            {/* Card 3 */}
            <div className="bg-white border border-amber-200 rounded-xl p-4 shadow-sm bg-amber-50/20">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-amber-700">Ra Bill Trễ</span>
                <Clock className="w-4 h-4 text-amber-600" />
              </div>
              <p className="text-2xl font-bold text-amber-600 mt-1 font-mono">{resultData.summary.countMatchedLate.toLocaleString()}</p>
            </div>

            {/* Card 4 */}
            <div className="bg-white border border-emerald-200 rounded-xl p-4 shadow-sm bg-emerald-50/20">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-emerald-700">Đã Ra Bill</span>
                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              </div>
              <p className="text-2xl font-bold text-emerald-600 mt-1 font-mono">{resultData.summary.countMatchedOnTime.toLocaleString()}</p>
            </div>
          </div>

          {/* Thanh Lọc & Tìm Kiếm */}
          <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-sm flex flex-col md:flex-row items-center justify-between gap-3">
            {/* Tabs */}
            <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-lg w-full md:w-auto">
              <button
                onClick={() => { setActiveTab('ALL'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-md text-xs font-semibold transition ${
                  activeTab === 'ALL' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Tất cả ({resultData.summary.total})
              </button>
              <button
                onClick={() => { setActiveTab('PENDING_BILL'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-md text-xs font-semibold transition ${
                  activeTab === 'PENDING_BILL' ? 'bg-rose-600 text-white shadow-sm' : 'text-rose-600 hover:bg-rose-50'
                }`}
              >
                Chưa ra Bill ({resultData.summary.countPendingBill})
              </button>
              <button
                onClick={() => { setActiveTab('MATCHED_LATE'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-md text-xs font-semibold transition ${
                  activeTab === 'MATCHED_LATE' ? 'bg-amber-600 text-white shadow-sm' : 'text-amber-600 hover:bg-amber-50'
                }`}
              >
                Ra Bill trễ ({resultData.summary.countMatchedLate})
              </button>
              <button
                onClick={() => { setActiveTab('MATCHED_ON_TIME'); setCurrentPage(1); }}
                className={`px-3 py-1.5 rounded-md text-xs font-semibold transition ${
                  activeTab === 'MATCHED_ON_TIME' ? 'bg-emerald-600 text-white shadow-sm' : 'text-emerald-700 hover:bg-emerald-50'
                }`}
              >
                Đúng hạn ({resultData.summary.countMatchedOnTime})
              </button>
            </div>

            {/* Tìm kiếm & Khoa */}
            <div className="flex items-center gap-2.5 w-full md:w-auto">
              <select
                value={selectedKhoa}
                onChange={(e) => { setSelectedKhoa(e.target.value); setCurrentPage(1); }}
                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-700 focus:outline-none focus:ring-1 focus:ring-teal-500"
              >
                <option value="ALL">Tất cả Khoa / Phòng</option>
                {departmentList.map(dep => (
                  <option key={dep} value={dep}>{dep}</option>
                ))}
              </select>

              <div className="relative flex-1 md:w-56">
                <Search className="w-3.5 h-3.5 absolute left-2.5 top-2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Mã BN, Họ tên, Số BL..."
                  value={searchTerm}
                  onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                  className="w-full text-xs pl-7 pr-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-teal-500"
                />
              </div>
            </div>
          </div>

          {/* BẢNG DỮ LIỆU */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-600 uppercase font-semibold border-b border-slate-200">
                  <tr>
                    <th className="py-2.5 px-3 w-12 text-center">STT</th>
                    <th className="py-2.5 px-3">Mã BN</th>
                    <th className="py-2.5 px-3">Họ và Tên</th>
                    <th className="py-2.5 px-3">Khoa Phòng</th>
                    <th className="py-2.5 px-3">Ngày Chốt</th>
                    <th className="py-2.5 px-3 text-center">Trạng Thái</th>
                    <th className="py-2.5 px-3">Số Biên Lai</th>
                    <th className="py-2.5 px-3">Ngày Ra Bill</th>
                    <th className="py-2.5 px-3 text-right">BN Trả (121)</th>
                    <th className="py-2.5 px-3 text-right">Thực Thu (140)</th>
                    <th className="py-2.5 px-3">Ghi Chú</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginatedItems.length === 0 ? (
                    <tr>
                      <td colSpan={11} className="py-8 text-center text-slate-400">
                        Không có dữ liệu.
                      </td>
                    </tr>
                  ) : (
                    paginatedItems.map((item) => (
                      <tr key={item.stt} className="hover:bg-slate-50/80 transition">
                        <td className="py-2 px-3 text-center text-slate-400 font-mono">{item.stt}</td>
                        <td className="py-2 px-3 font-semibold text-slate-900 font-mono">{item.maBN}</td>
                        <td className="py-2 px-3 font-medium text-slate-800">{item.hoTen}</td>
                        <td className="py-2 px-3 text-slate-600">{item.khoaPhong}</td>
                        <td className="py-2 px-3 text-slate-600">{item.ngayChot}</td>
                        
                        <td className="py-2 px-3 text-center">
                          {item.status === 'MATCHED_ON_TIME' && (
                            <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                              Đã ra Bill
                            </span>
                          )}
                          {item.status === 'MATCHED_LATE' && (
                            <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">
                              Trễ {item.soNgayTre} ngày
                            </span>
                          )}
                          {item.status === 'PENDING_BILL' && (
                            <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                              Chưa ra Bill
                            </span>
                          )}
                        </td>

                        <td className="py-2 px-3 font-mono text-slate-700">{item.soBienLai || '-'}</td>
                        <td className="py-2 px-3 text-slate-600">{item.ngayBill || '-'}</td>
                        <td className="py-2 px-3 text-right font-mono text-slate-700">
                          {item.tienBNTra121.toLocaleString()} đ
                        </td>
                        <td className="py-2 px-3 text-right font-mono font-medium text-slate-900">
                          {item.tienThanhToan140 > 0 ? `${item.tienThanhToan140.toLocaleString()} đ` : '-'}
                        </td>
                        <td className="py-2 px-3 text-slate-500 text-[11px]">
                          {item.notes}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Phân trang */}
            <div className="bg-slate-50 px-3 py-2 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
              <div>
                <span className="font-semibold text-slate-700">{Math.min(filteredItems.length, (currentPage - 1) * pageSize + 1)}</span> -{' '}
                <span className="font-semibold text-slate-700">{Math.min(filteredItems.length, currentPage * pageSize)}</span> /{' '}
                <span className="font-semibold text-slate-700">{filteredItems.length}</span>
              </div>

              <div className="flex items-center gap-1">
                <button
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="p-1 rounded border border-slate-300 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                </button>
                <span className="px-2 font-medium">{currentPage} / {totalPages}</span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="p-1 rounded border border-slate-300 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
