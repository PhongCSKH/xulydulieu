import * as XLSX from 'xlsx';

/**
 * Chuẩn hóa chuỗi (loại bỏ khoảng trắng thừa, đưa về chữ thường)
 */
export function cleanStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Chuẩn hóa tên khoa phòng để so khớp linh hoạt
 * Loại bỏ các tiền tố thông dụng như "phòng khám", "pk", "khoa", dấu gạch ngang
 */
export function normalizeDeptName(val) {
  let s = cleanStr(val);
  if (!s) return '';
  s = s.replace(/^(phòng khám|phong kham|pk|khoa)\s+/gi, '');
  s = s.replace(/[\-\–]/g, ' ');
  return s.trim().replace(/\s+/g, ' ');
}

/**
 * Làm sạch chuỗi tiền tệ và chuyển đổi thành Number
 * Xử lý: "500.800 đ", "89,200", "140.800", "-50.000", null -> 500800
 */
export function cleanMoney(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : Math.round(val);
  
  let s = String(val).trim().replace(/[đ₫\s]/gi, '');
  // Kiểm tra số âm
  const isNegative = s.startsWith('-') || (s.startsWith('(') && s.endsWith(')'));
  s = s.replace(/[()\-]/g, '');

  // Nếu dạng 500.800 hoặc 500,800
  s = s.replace(/[.,]/g, '');
  const num = parseInt(s, 10);
  if (isNaN(num)) return 0;
  return isNegative ? -num : num;
}

/**
 * Định dạng tiền tệ hiển thị đẹp mắt (500.800 đ)
 */
export function formatCurrency(amount) {
  if (amount === null || amount === undefined || isNaN(amount)) return '0 đ';
  return new Intl.NumberFormat('vi-VN').format(amount) + ' đ';
}

/**
 * Phân tích ô Date/Time từ Excel, bóc tách Ngày (DD/MM/YYYY) và Thời gian (HH:mm:ss)
 */
export function parseExcelDateAndTime(val) {
  let dateStr = '';
  let timeStr = '';

  if (!val && val !== 0) {
    return { dateStr: '', timeStr: '', raw: val };
  }

  // 1. Trường hợp val là đối tượng Date (khi đọc sheet với cellDates: true)
  if (val instanceof Date && !isNaN(val.getTime())) {
    const d = val.getDate().toString().padStart(2, '0');
    const m = (val.getMonth() + 1).toString().padStart(2, '0');
    const y = val.getFullYear();
    dateStr = `${d}/${m}/${y}`;

    const hh = val.getHours().toString().padStart(2, '0');
    const mm = val.getMinutes().toString().padStart(2, '0');
    const ss = val.getSeconds().toString().padStart(2, '0');
    timeStr = `${hh}:${mm}:${ss}`;
    return { dateStr, timeStr, raw: val };
  }

  // 2. Trường hợp val là số serial của Excel
  if (typeof val === 'number') {
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed) {
      const d = String(parsed.d).padStart(2, '0');
      const m = String(parsed.m).padStart(2, '0');
      const y = parsed.y;
      dateStr = `${d}/${m}/${y}`;

      const hh = String(parsed.H || 0).padStart(2, '0');
      const mm = String(parsed.M || 0).padStart(2, '0');
      const ss = String(Math.floor(parsed.S || 0)).padStart(2, '0');
      timeStr = `${hh}:${mm}:${ss}`;
      return { dateStr, timeStr, raw: val };
    }
  }

  // 3. Trường hợp val là chuỗi String (vd: "06/06/2026 18:18:00" hoặc "06/06/2026")
  if (typeof val === 'string') {
    const str = val.trim();
    // Bóc tách ngày DD/MM/YYYY hoặc YYYY-MM-DD
    const matchDateDMY = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    const matchDateYMD = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);

    if (matchDateDMY) {
      dateStr = `${matchDateDMY[1].padStart(2, '0')}/${matchDateDMY[2].padStart(2, '0')}/${matchDateDMY[3]}`;
    } else if (matchDateYMD) {
      dateStr = `${matchDateYMD[3].padStart(2, '0')}/${matchDateYMD[2].padStart(2, '0')}/${matchDateYMD[1]}`;
    }

    // Bóc tách giờ HH:mm:ss hoặc HH:mm
    const matchTime = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (matchTime) {
      const hh = matchTime[1].padStart(2, '0');
      const mm = matchTime[2].padStart(2, '0');
      const ss = (matchTime[3] || '00').padStart(2, '0');
      timeStr = `${hh}:${mm}:${ss}`;
    }

    return { dateStr, timeStr, raw: val };
  }

  return { dateStr: String(val), timeStr: '', raw: val };
}

/**
 * Tự động nhận diện loại file dựa trên tiêu đề cột
 */
export function detectFileType(sheet) {
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (!json || json.length === 0) return 'UNKNOWN';

  // Quét 3 dòng đầu để tìm header (tránh file có tiêu đề ở dòng 1-2)
  for (let r = 0; r < Math.min(3, json.length); r++) {
    const headerRow = (json[r] || []).map(cleanStr);
    
    // File KH Bỏ Về đặc trưng: Bác sĩ, Xử trí, Thuốc BHYT, Nhân viên xử lý L1/L2/L3, Trạng thái tiền, KH bỏ về
    const hasBoVeKeywords = headerRow.some(h => 
      h.includes('xử trí') || h.includes('xu tri') || 
      h.includes('thuốc bhyt') || h.includes('thuoc bhyt') ||
      h.includes('nhân viên xử lý') || h.includes('nhan vien xu ly') ||
      h.includes('trạng thái tiền') || h.includes('trang thai tien') ||
      h.includes('gọi sms') || h.includes('nhật ký hệ thống')
    );
    if (hasBoVeKeywords) return 'BC_BO_VE';

    // File BC140 đặc trưng: Mã thu ngân, Hoàn chuyển khoản, Bệnh nhân thanh toán, Số biên lai, Quyển số
    const has140Keywords = headerRow.some(h => 
      h.includes('mã thu ngân') || h.includes('ma thu ngan') ||
      h.includes('họ tên thu ngân') || h.includes('ho ten thu ngan') ||
      h.includes('hoàn chuyển khoản') || h.includes('hoan chuyen khoan') ||
      h.includes('số biên lai') || h.includes('so bien lai') ||
      h.includes('quyển số') || h.includes('quyen so')
    );
    if (has140Keywords) return 'BC140';
  }

  return 'UNKNOWN';
}

/**
 * Tìm kiếm giá trị cột linh hoạt theo danh sách tên cột ứng cử viên
 */
export function getRowVal(row, candidateNames) {
  if (!row) return undefined;
  for (const name of candidateNames) {
    if (row[name] !== undefined) return row[name];
    // Tìm không phân biệt hoa thường
    const cleanCandidate = cleanStr(name);
    for (const key of Object.keys(row)) {
      if (cleanStr(key) === cleanCandidate) return row[key];
    }
  }
  return undefined;
}

/**
 * Gộp và phát hiện trùng lặp trong danh sách file BC140
 */
export function mergeAndDeduplicateBC140(fileRecords) {
  const mergedRows = [];
  const duplicateAlerts = [];
  const seenReceipts = new Map(); // Key -> info

  fileRecords.forEach(fileRec => {
    const { fileName, rows } = fileRec;
    rows.forEach((row, rowIndex) => {
      const maBN = cleanStr(getRowVal(row, ['Mã bệnh nhân', 'Mã BN', 'mã bệnh nhân', 'Mã người bệnh']));
      const soBL = String(getRowVal(row, ['Số biên lai', 'Số BL', 'so bien lai', 'Số BL/HĐ']) || '').trim();
      const ngayRaw = getRowVal(row, ['Ngày', 'Ngay', 'Ngày lập', 'Ngày thanh toán']);
      const dt = parseExcelDateAndTime(ngayRaw);
      
      // Tạo composite key duy nhất: Ưu tiên Số biên lai, nếu trống thì dùng Mã BN + Ngày + Giờ
      const uniqueKey = soBL ? `BL_${soBL}` : `BN_${maBN}_${dt.dateStr}_${dt.timeStr}`;

      if (uniqueKey && seenReceipts.has(uniqueKey)) {
        const prev = seenReceipts.get(uniqueKey);
        duplicateAlerts.push({
          type: 'BC140_DUPLICATE',
          key: uniqueKey,
          receiptNumber: soBL || '(Không có số BL)',
          maBN,
          fileName,
          rowIndex: rowIndex + 2,
          originalFile: prev.fileName,
          originalRow: prev.rowIndex
        });
      } else {
        seenReceipts.set(uniqueKey, { fileName, rowIndex: rowIndex + 2 });
        // Bổ sung các trường chuẩn hóa trực tiếp vào row
        row.__parsedDate = dt.dateStr;
        row.__parsedTime = dt.timeStr;
        row.__sourceFile = fileName;
        mergedRows.push(row);
      }
    });
  });

  return { mergedRows, duplicateAlerts };
}

/**
 * Gộp và phát hiện trùng lặp trong danh sách file BC KH Bỏ Về
 */
export function mergeAndDeduplicateBoVe(fileRecords) {
  const mergedRows = [];
  const duplicateAlerts = [];
  const seenCases = new Map();

  fileRecords.forEach(fileRec => {
    const { fileName, rows } = fileRec;
    rows.forEach((row, rowIndex) => {
      const maBN = cleanStr(getRowVal(row, ['MÃ BỆNH NHÂN', 'Mã bệnh nhân', 'Mã BN']));
      const ngayRaw = getRowVal(row, ['NGÀY', 'Ngày']);
      const dt = parseExcelDateAndTime(ngayRaw);
      const khoaPhong = cleanStr(getRowVal(row, ['KHOA PHÒNG', 'Khoa phòng']));
      
      if (!maBN) return;

      const uniqueKey = `${maBN}_${dt.dateStr}_${khoaPhong}`;
      if (seenCases.has(uniqueKey)) {
        const prev = seenCases.get(uniqueKey);
        duplicateAlerts.push({
          type: 'BO_VE_DUPLICATE',
          key: uniqueKey,
          maBN,
          ngay: dt.dateStr,
          khoaPhong,
          fileName,
          rowIndex: rowIndex + 2,
          originalFile: prev.fileName,
          originalRow: prev.rowIndex
        });
      } else {
        seenCases.set(uniqueKey, { fileName, rowIndex: rowIndex + 2 });
      }

      row.__parsedDate = dt.dateStr;
      row.__parsedTime = dt.timeStr;
      row.__sourceFile = fileName;
      mergedRows.push(row);
    });
  });

  return { mergedRows, duplicateAlerts };
}

/**
 * THUẬT TOÁN ĐỐI SOÁT TẠM ỨNG LẠI (Re-advance Reconciliation)
 */
export function processReconciliationTamUng(rowsBoVe, rows140) {
  // 1. Lập chỉ mục BC140 theo Mã Bệnh Nhân để tìm kiếm cực nhanh O(1)
  const map140ByBN = new Map();
  rows140.forEach(row => {
    const maBN = cleanStr(getRowVal(row, ['Mã bệnh nhân', 'Mã BN', 'mã bệnh nhân', 'Mã người bệnh']));
    if (!maBN) return;
    if (!map140ByBN.has(maBN)) {
      map140ByBN.set(maBN, []);
    }
    map140ByBN.get(maBN).push(row);
  });

  const results = [];
  let countMatched = 0;
  let countTransferRefund = 0;
  let countDiffDate = 0;
  let countNotFound = 0;
  let totalReAdvanceAmount = 0;
  let totalRefundAmount = 0;

  // 2. Duyệt từng dòng trong file BC KH Bỏ Về
  rowsBoVe.forEach((rBoVe, idx) => {
    const maBN = cleanStr(getRowVal(rBoVe, ['MÃ BỆNH NHÂN', 'Mã bệnh nhân', 'Mã BN']));
    const hoTen = String(getRowVal(rBoVe, ['HỌ TÊN', 'Họ tên', 'Họ Và Tên']) || '').trim();
    const ngayBoVe = rBoVe.__parsedDate || parseExcelDateAndTime(getRowVal(rBoVe, ['NGÀY', 'Ngày'])).dateStr;
    const khoaPhongBoVe = String(getRowVal(rBoVe, ['KHOA PHÒNG', 'Khoa phòng']) || '').trim();
    const normDeptBoVe = normalizeDeptName(khoaPhongBoVe);
    
    // Số tiền mục tiêu cần đối chiếu
    const rawSoTien = getRowVal(rBoVe, ['SỐ TIỀN', 'Số tiền', 'Số Tiền']);
    const targetAmount = cleanMoney(rawSoTien);

    // Lấy các giao dịch phát sinh của BN này trong BC140
    const candidateRows = map140ByBN.get(maBN) || [];

    let bestMatch = null;
    let matchQuality = 'NONE'; // EXACT | APPROX_DEPT | DIFF_DATE | BY_MONEY
    let foundTransferRefundVal = null;
    let foundTransferRefundRow = null;

    // Kiểm tra xem BN có bất kỳ dòng nào có "Hoàn chuyển khoản" không (Cột AF)
    for (const cRow of candidateRows) {
      const hckVal = getRowVal(cRow, ['Hoàn chuyển khoản', 'Hoan chuyen khoan', 'HOÀN CHUYỂN KHOẢN']);
      if (hckVal !== undefined && hckVal !== null && String(hckVal).trim() !== '' && String(hckVal).trim() !== '0') {
        foundTransferRefundVal = hckVal;
        foundTransferRefundRow = cRow;
        break;
      }
    }

    // Nếu có giao dịch phát sinh, tìm dòng khớp số tiền
    if (candidateRows.length > 0 && targetAmount > 0) {
      // Hàm kiểm tra 1 dòng BC140 có khớp số tiền không
      const matchesAmount = (r) => {
        // Kiểm tra các cột tiền khả dĩ trong BC140
        const possibleAmounts = [
          cleanMoney(getRowVal(r, ['Tạm ứng', 'Tam ung', 'Tiền tạm ứng'])),
          cleanMoney(getRowVal(r, ['Thu tạm ứng', 'Thu tam ung'])),
          cleanMoney(getRowVal(r, ['Bệnh nhân thanh toán', 'Bệnh nhân trả', 'BN thanh toán'])),
          cleanMoney(getRowVal(r, ['Tổng cộng', 'Tổng thanh toán', 'Thực thu'])),
          cleanMoney(getRowVal(r, ['Số tiền', 'Số Tiền', 'Thành tiền']))
        ];
        // So khớp với sai số nhỏ hơn 100đ
        return possibleAmounts.some(amt => Math.abs(amt - targetAmount) < 100);
      };

      // Tầng 1: Khớp tuyệt đối: Cùng Ngày + Chuẩn hóa Khoa Phòng + Khớp Số tiền
      bestMatch = candidateRows.find(r => {
        const rDate = r.__parsedDate || parseExcelDateAndTime(getRowVal(r, ['Ngày', 'Ngay'])).dateStr;
        const rDept = normalizeDeptName(getRowVal(r, ['Khoa phòng', 'Phòng chỉ định', 'Khoa']) || '');
        return rDate === ngayBoVe && (rDept === normDeptBoVe || !normDeptBoVe) && matchesAmount(r);
      });

      if (bestMatch) {
        matchQuality = 'EXACT';
      } else {
        // Tầng 2: Khớp cùng Ngày + Đúng Số tiền (Khoa phòng có thể khác do điều chuyển/gom bill)
        bestMatch = candidateRows.find(r => {
          const rDate = r.__parsedDate || parseExcelDateAndTime(getRowVal(r, ['Ngày', 'Ngay'])).dateStr;
          return rDate === ngayBoVe && matchesAmount(r);
        });

        if (bestMatch) {
          matchQuality = 'APPROX_DEPT';
        } else {
          // Tầng 3: Khớp đúng Số tiền nhưng khác ngày (bệnh nhân đến tạm ứng lại vào ngày hôm sau)
          bestMatch = candidateRows.find(r => matchesAmount(r));
          if (bestMatch) {
            matchQuality = 'DIFF_DATE';
          }
        }
      }
    }

    // Nếu không khớp dòng tiền cụ thể nhưng có dòng Hoàn chuyển khoản
    if (!bestMatch && foundTransferRefundRow) {
      bestMatch = foundTransferRefundRow;
    }

    // Tổng hợp kết quả
    let status = 'NOT_FOUND';
    let statusText = 'Chưa tìm thấy biên lai';
    let maThuNgan = '';
    let tenThuNgan = '';
    let soBienLai = '';
    let thoiGianBL = '';
    let ngayBL = '';
    let hoanChuyenKhoanStr = '';
    let auditNotes = '';

    // Lấy thông tin Hoàn chuyển khoản nếu có
    if (foundTransferRefundVal) {
      hoanChuyenKhoanStr = String(foundTransferRefundVal).trim();
      const refundNum = cleanMoney(foundTransferRefundVal);
      if (refundNum > 0) totalRefundAmount += refundNum;
    }

    if (bestMatch) {
      maThuNgan = String(getRowVal(bestMatch, ['Mã thu ngân', 'Mã TN', 'Ma thu ngan']) || '').trim();
      tenThuNgan = String(getRowVal(bestMatch, ['Họ tên thu ngân', 'Họ và tên thu ngân', 'Thu ngân', 'Tên thu ngân']) || '').trim();
      soBienLai = String(getRowVal(bestMatch, ['Số biên lai', 'Số BL', 'Số phiếu']) || '').trim();
      
      const dt = parseExcelDateAndTime(getRowVal(bestMatch, ['Ngày', 'Ngay']));
      thoiGianBL = bestMatch.__parsedTime || dt.timeStr;
      ngayBL = bestMatch.__parsedDate || dt.dateStr;

      if (hoanChuyenKhoanStr) {
        status = 'MATCHED_TRANSFER_REFUND';
        statusText = 'Có Hoàn Chuyển Khoản';
        countTransferRefund++;
        auditNotes = `Tìm thấy GD hoàn CK: ${hoanChuyenKhoanStr} (TN: ${maThuNgan || 'N/A'})`;
      } else if (matchQuality === 'EXACT') {
        status = 'MATCHED_RE_ADVANCE';
        statusText = 'Đã khớp tạm ứng lại';
        countMatched++;
        totalReAdvanceAmount += targetAmount;
        auditNotes = `Khớp chuẩn xác ngày ${ngayBL} lúc ${thoiGianBL || '--:--'} (BL: ${soBienLai})`;
      } else if (matchQuality === 'APPROX_DEPT') {
        status = 'MATCHED_RE_ADVANCE';
        statusText = 'Đã khớp tiền (Khác khoa)';
        countMatched++;
        totalReAdvanceAmount += targetAmount;
        const dept140 = String(getRowVal(bestMatch, ['Khoa phòng', 'Khoa']) || '').trim();
        auditNotes = `Khớp đúng tiền ${formatCurrency(targetAmount)}, phòng BC140 ghi: "${dept140}"`;
      } else if (matchQuality === 'DIFF_DATE') {
        status = 'MATCHED_DIFF_DATE';
        statusText = 'Khớp tiền - Lệch ngày';
        countDiffDate++;
        totalReAdvanceAmount += targetAmount;
        auditNotes = `Khám ${ngayBoVe} nhưng biên lai lập ngày ${ngayBL} lúc ${thoiGianBL} (BL: ${soBienLai})`;
      } else {
        status = 'MATCHED_RE_ADVANCE';
        statusText = 'Đã đối chiếu';
        countMatched++;
      }
    } else {
      status = 'NOT_FOUND';
      statusText = 'Không tìm thấy trên BC140';
      countNotFound++;
      auditNotes = candidateRows.length > 0 
        ? `Có ${candidateRows.length} GD của BN trên BC140 nhưng không khớp số tiền ${formatCurrency(targetAmount)}`
        : 'Không có dữ liệu của BN này trong các file BC140 đã nạp';
    }

    // Kết quả trả về chứa dòng gốc + các trường bổ sung
    results.push({
      stt: idx + 1,
      originalRow: rBoVe,
      maBN,
      hoTen,
      ngayBoVe,
      khoaPhongBoVe,
      targetAmount,
      rawAmount: rawSoTien,
      status,
      statusText,
      maThuNgan,
      tenThuNgan,
      soBienLai,
      thoiGianBL,
      ngayBL,
      hoanChuyenKhoan: hoanChuyenKhoanStr,
      auditNotes
    });
  });

  const total = results.length;
  const percentCompleted = total > 0 ? Math.round(((countMatched + countTransferRefund + countDiffDate) / total) * 100) : 0;

  return {
    summary: {
      total,
      countMatched,
      countTransferRefund,
      countDiffDate,
      countNotFound,
      percentCompleted,
      totalReAdvanceAmount,
      totalRefundAmount
    },
    items: results
  };
}

/**
 * Xuất kết quả rà soát ra tệp Excel bảo toàn cột gốc của file KH Bỏ Về
 */
export function exportTamUngToExcel(items, summary, fileName = 'Ket_Qua_Ra_Soat_Tam_Ung_Lai.xlsx') {
  const exportData = items.map(item => {
    const orig = item.originalRow || {};
    // Nhân bản đối tượng gốc để giữ nguyên thứ tự các cột ban đầu
    const rowObj = { ...orig };

    // Cập nhật hoặc ghi đè cột SỐ BIÊN LAI nếu có kết quả
    if (item.soBienLai) {
      rowObj['SỐ BIÊN LAI'] = item.soBienLai;
    }

    // Bổ sung các cột trích xuất từ BC140 vào phía sau
    rowObj['MÃ THU NGÂN (BC140)'] = item.maThuNgan;
    rowObj['HỌ TÊN THU NGÂN (BC140)'] = item.tenThuNgan;
    rowObj['THỜI GIAN BIÊN LAI (BC140)'] = item.thoiGianBL;
    rowObj['NGÀY BIÊN LAI (BC140)'] = item.ngayBL;
    rowObj['HOÀN CHUYỂN KHOẢN (BC140)'] = item.hoanChuyenKhoan;
    rowObj['TRẠNG THÁI RÀ SOÁT'] = item.statusText;
    rowObj['GHI CHÚ ĐỐI SOÁT CHI TIẾT'] = item.auditNotes;

    // Loại bỏ các trường tạm thời do engine tạo ra (nếu có)
    delete rowObj['__parsedDate'];
    delete rowObj['__parsedTime'];
    delete rowObj['__sourceFile'];

    return rowObj;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'RaSoatTamUngLai');

  XLSX.writeFile(workbook, fileName);
}
