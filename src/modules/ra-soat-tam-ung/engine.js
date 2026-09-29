import * as XLSX from 'xlsx';

export function cleanStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

export function normalizeDeptName(val) {
  let s = cleanStr(val);
  if (!s) return '';
  s = s.replace(/^(phòng khám|phong kham|pk|khoa)\s+/gi, '');
  s = s.replace(/[\-\–]/g, ' ');
  return s.trim().replace(/\s+/g, ' ');
}

export function cleanMoney(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : Math.round(val);
  
  let s = String(val).trim().replace(/[đ₫\s]/gi, '');
  const isNegative = s.startsWith('-') || (s.startsWith('(') && s.endsWith(')'));
  s = s.replace(/[()\-]/g, '');

  s = s.replace(/[.,]/g, '');
  const num = parseInt(s, 10);
  if (isNaN(num)) return 0;
  return isNegative ? -num : num;
}

export function formatCurrency(amount) {
  if (amount === null || amount === undefined || isNaN(amount)) return '0 đ';
  return new Intl.NumberFormat('vi-VN').format(amount) + ' đ';
}

export function parseExcelDateAndTime(val) {
  let dateStr = '';
  let timeStr = '';

  if (!val && val !== 0) {
    return { dateStr: '', timeStr: '', raw: val };
  }

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

  if (typeof val === 'string') {
    const str = val.trim();
    const matchDateDMY = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    const matchDateYMD = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);

    if (matchDateDMY) {
      dateStr = `${matchDateDMY[1].padStart(2, '0')}/${matchDateDMY[2].padStart(2, '0')}/${matchDateDMY[3]}`;
    } else if (matchDateYMD) {
      dateStr = `${matchDateYMD[3].padStart(2, '0')}/${matchDateYMD[2].padStart(2, '0')}/${matchDateYMD[1]}`;
    }

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

export function detectFileType(sheet) {
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (!json || json.length === 0) return 'UNKNOWN';

  for (let r = 0; r < Math.min(3, json.length); r++) {
    const headerRow = (json[r] || []).map(cleanStr);
    
    const hasBoVeKeywords = headerRow.some(h => 
      h.includes('xử trí') || h.includes('xu tri') || 
      h.includes('thuốc bhyt') || h.includes('thuoc bhyt') ||
      h.includes('nhân viên xử lý') || h.includes('nhan vien xu ly') ||
      h.includes('trạng thái tiền') || h.includes('trang thai tien') ||
      h.includes('nhật ký hệ thống')
    );
    if (hasBoVeKeywords) return 'BC_BO_VE';

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

export function getRowVal(row, candidateNames) {
  if (!row) return undefined;
  for (const name of candidateNames) {
    if (row[name] !== undefined) return row[name];
    const cleanCandidate = cleanStr(name);
    for (const key of Object.keys(row)) {
      if (cleanStr(key) === cleanCandidate) return row[key];
    }
  }
  return undefined;
}

/**
 * Xử lý file BC140 dung lượng lớn (lên tới 200.000 dòng)
 * Tối ưu: dense mode, lọc index, early filter theo targetMaBNSet, chunking bất đồng bộ
 */
export async function parseLargeBC140FileAsync(file, targetMaBNSet = null, onProgress = null) {
  return new Promise(async (resolve, reject) => {
    try {
      const buffer = await file.arrayBuffer();
      if (onProgress) onProgress({ percent: 15, text: `Đang tải ${file.name}...` });

      await new Promise(r => setTimeout(r, 20));

      const workbook = XLSX.read(new Uint8Array(buffer), { 
        type: 'array', 
        dense: true, 
        cellDates: true 
      });

      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      if (!rawRows || rawRows.length === 0) {
        return resolve({ fileName: file.name, totalRows: 0, rows: [] });
      }

      const totalRows = rawRows.length;
      if (onProgress) onProgress({ percent: 30, text: `Đang phân tích cấu trúc (${totalRows.toLocaleString()} dòng)...` });

      let headerRowIndex = 0;
      let headerRow = [];
      for (let r = 0; r < Math.min(5, rawRows.length); r++) {
        const row = (rawRows[r] || []).map(cleanStr);
        if (row.some(c => c.includes('mã thu ngân') || c.includes('mã bn') || c.includes('mã bệnh nhân') || c.includes('số biên lai') || c.includes('hoàn chuyển khoản'))) {
          headerRowIndex = r;
          headerRow = row;
          break;
        }
      }

      let idxNgay = -1;
      let idxMaTN = -1;
      let idxTenTN = -1;
      let idxMaBN = -1;
      let idxKhoa = -1;
      let idxSoBL = -1;
      let idxHoanCK = -1;
      const idxTienCols = [];

      headerRow.forEach((colName, idx) => {
        if (!colName) return;
        if (idxNgay === -1 && (colName === 'ngày' || colName === 'ngay' || colName.includes('ngày thanh toán') || colName.includes('ngày lập'))) {
          idxNgay = idx;
        }
        if (idxMaTN === -1 && (colName.includes('mã thu ngân') || colName === 'mã tn')) {
          idxMaTN = idx;
        }
        if (idxTenTN === -1 && (colName.includes('họ tên thu ngân') || colName.includes('tên thu ngân') || colName === 'thu ngân')) {
          idxTenTN = idx;
        }
        if (idxMaBN === -1 && (colName.includes('mã bệnh nhân') || colName.includes('mã bn') || colName.includes('mã người bệnh'))) {
          idxMaBN = idx;
        }
        if (idxKhoa === -1 && (colName.includes('khoa phòng') || colName.includes('phòng chỉ định') || colName === 'khoa')) {
          idxKhoa = idx;
        }
        if (idxSoBL === -1 && (colName.includes('số biên lai') || colName.includes('số bl') || colName.includes('số phiếu'))) {
          idxSoBL = idx;
        }
        if (idxHoanCK === -1 && (colName.includes('hoàn chuyển khoản') || colName.includes('hoan chuyen khoan'))) {
          idxHoanCK = idx;
        }
        if (
          colName.includes('tạm ứng') || 
          colName.includes('thu tạm ứng') || 
          colName.includes('bệnh nhân thanh toán') || 
          colName.includes('bệnh nhân trả') || 
          colName.includes('tổng cộng') || 
          colName.includes('tổng thanh toán') ||
          colName === 'số tiền'
        ) {
          idxTienCols.push(idx);
        }
      });

      if (idxNgay === -1) idxNgay = 0;
      if (idxMaTN === -1 && headerRow.length > 2) idxMaTN = 2;
      if (idxTenTN === -1 && headerRow.length > 3) idxTenTN = 3;
      if (idxHoanCK === -1 && headerRow.length > 31) idxHoanCK = 31;

      const extractedRows = [];
      const chunkSize = 20000;
      const startRow = headerRowIndex + 1;

      for (let i = startRow; i < totalRows; i += chunkSize) {
        const endRow = Math.min(i + chunkSize, totalRows);
        
        for (let r = i; r < endRow; r++) {
          const row = rawRows[r];
          if (!row || row.length === 0) continue;

          const rawMaBN = idxMaBN !== -1 ? row[idxMaBN] : null;
          const maBN = cleanStr(rawMaBN);
          if (!maBN) continue;

          const rawHoanCK = idxHoanCK !== -1 ? row[idxHoanCK] : null;
          const hasHoanCK = rawHoanCK !== undefined && rawHoanCK !== null && String(rawHoanCK).trim() !== '' && String(rawHoanCK).trim() !== '0';

          if (targetMaBNSet && targetMaBNSet.size > 0) {
            if (!targetMaBNSet.has(maBN) && !hasHoanCK) {
              continue;
            }
          }

          const rawNgay = idxNgay !== -1 ? row[idxNgay] : null;
          const dt = parseExcelDateAndTime(rawNgay);

          const soBL = idxSoBL !== -1 ? String(row[idxSoBL] || '').trim() : '';
          const maTN = idxMaTN !== -1 ? String(row[idxMaTN] || '').trim() : '';
          const tenTN = idxTenTN !== -1 ? String(row[idxTenTN] || '').trim() : '';
          const khoa = idxKhoa !== -1 ? String(row[idxKhoa] || '').trim() : '';

          const amounts = [];
          if (idxTienCols.length > 0) {
            for (const colIdx of idxTienCols) {
              const val = row[colIdx];
              if (val !== undefined && val !== null && val !== '') {
                const amt = cleanMoney(val);
                if (amt !== 0) amounts.push(amt);
              }
            }
          }

          extractedRows.push({
            maBN,
            soBL,
            maTN,
            tenTN,
            ngay: dt.dateStr,
            gio: dt.timeStr,
            khoa,
            amounts,
            hoanCK: hasHoanCK ? String(rawHoanCK).trim() : '',
            fileName: file.name,
            rowIndex: r + 1
          });
        }

        const pct = Math.min(95, Math.round(30 + ((endRow - startRow) / (totalRows - startRow)) * 65));
        if (onProgress) {
          onProgress({ 
            percent: pct, 
            text: `Đang quét: ${endRow.toLocaleString()} / ${totalRows.toLocaleString()} dòng (${extractedRows.length.toLocaleString()} dòng liên quan)...` 
          });
        }
        await new Promise(res => setTimeout(res, 5));
      }

      if (onProgress) {
        onProgress({ 
          percent: 100, 
          text: `Hoàn tất (${totalRows.toLocaleString()} dòng, lọc ${extractedRows.length.toLocaleString()} dòng liên quan).` 
        });
      }

      resolve({
        fileName: file.name,
        totalRows,
        extractedCount: extractedRows.length,
        rows: extractedRows
      });
    } catch (err) {
      console.error('Lỗi phân tích file lớn:', err);
      reject(err);
    }
  });
}

export function mergeAndDeduplicateBC140(fileRecords) {
  const mergedRows = [];
  const duplicateAlerts = [];
  const seenReceipts = new Map();

  fileRecords.forEach(fileRec => {
    const { fileName, rows } = fileRec;
    rows.forEach((row) => {
      const maBN = row.maBN;
      const soBL = row.soBL;
      const uniqueKey = soBL ? `BL_${soBL}` : `BN_${maBN}_${row.ngay}_${row.gio}`;

      if (uniqueKey && seenReceipts.has(uniqueKey)) {
        const prev = seenReceipts.get(uniqueKey);
        duplicateAlerts.push({
          type: 'BC140_DUPLICATE',
          key: uniqueKey,
          receiptNumber: soBL || '(Trống BL)',
          maBN,
          fileName,
          rowIndex: row.rowIndex,
          originalFile: prev.fileName,
          originalRow: prev.rowIndex
        });
      } else {
        seenReceipts.set(uniqueKey, { fileName, rowIndex: row.rowIndex });
        mergedRows.push(row);
      }
    });
  });

  return { mergedRows, duplicateAlerts };
}

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

export function processReconciliationTamUng(rowsBoVe, rows140) {
  const map140ByBN = new Map();
  rows140.forEach(row => {
    const maBN = row.maBN;
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

  rowsBoVe.forEach((rBoVe, idx) => {
    const maBN = cleanStr(getRowVal(rBoVe, ['MÃ BỆNH NHÂN', 'Mã bệnh nhân', 'Mã BN']));
    const hoTen = String(getRowVal(rBoVe, ['HỌ TÊN', 'Họ tên', 'Họ Và Tên']) || '').trim();
    const ngayBoVe = rBoVe.__parsedDate || parseExcelDateAndTime(getRowVal(rBoVe, ['NGÀY', 'Ngày'])).dateStr;
    const khoaPhongBoVe = String(getRowVal(rBoVe, ['KHOA PHÒNG', 'Khoa phòng']) || '').trim();
    const normDeptBoVe = normalizeDeptName(khoaPhongBoVe);
    
    const rawSoTien = getRowVal(rBoVe, ['SỐ TIỀN', 'Số tiền', 'Số Tiền']);
    const targetAmount = cleanMoney(rawSoTien);

    const candidateRows = map140ByBN.get(maBN) || [];

    let bestMatch = null;
    let matchQuality = 'NONE';
    let foundTransferRefundVal = null;
    let foundTransferRefundRow = null;

    for (const cRow of candidateRows) {
      if (cRow.hoanCK) {
        foundTransferRefundVal = cRow.hoanCK;
        foundTransferRefundRow = cRow;
        break;
      }
    }

    if (candidateRows.length > 0 && targetAmount > 0) {
      const matchesAmount = (r) => {
        return r.amounts && r.amounts.some(amt => Math.abs(amt - targetAmount) < 100);
      };

      bestMatch = candidateRows.find(r => {
        const rDept = normalizeDeptName(r.khoa);
        return r.ngay === ngayBoVe && (rDept === normDeptBoVe || !normDeptBoVe) && matchesAmount(r);
      });

      if (bestMatch) {
        matchQuality = 'EXACT';
      } else {
        bestMatch = candidateRows.find(r => r.ngay === ngayBoVe && matchesAmount(r));
        if (bestMatch) {
          matchQuality = 'APPROX_DEPT';
        } else {
          bestMatch = candidateRows.find(r => matchesAmount(r));
          if (bestMatch) {
            matchQuality = 'DIFF_DATE';
          }
        }
      }
    }

    if (!bestMatch && foundTransferRefundRow) {
      bestMatch = foundTransferRefundRow;
    }

    let status = 'NOT_FOUND';
    let statusText = 'Chưa tìm thấy biên lai';
    let maThuNgan = '';
    let tenThuNgan = '';
    let soBienLai = '';
    let thoiGianBL = '';
    let ngayBL = '';
    let hoanChuyenKhoanStr = '';
    let auditNotes = '';

    if (foundTransferRefundVal) {
      hoanChuyenKhoanStr = String(foundTransferRefundVal).trim();
      const refundNum = cleanMoney(foundTransferRefundVal);
      if (refundNum > 0) totalRefundAmount += refundNum;
    }

    if (bestMatch) {
      maThuNgan = bestMatch.maTN || '';
      tenThuNgan = bestMatch.tenTN || '';
      soBienLai = bestMatch.soBL || '';
      thoiGianBL = bestMatch.gio || '';
      ngayBL = bestMatch.ngay || '';

      if (hoanChuyenKhoanStr) {
        status = 'MATCHED_TRANSFER_REFUND';
        statusText = 'Có Hoàn Chuyển Khoản';
        countTransferRefund++;
        auditNotes = `Hoàn CK: ${hoanChuyenKhoanStr} (TN: ${maThuNgan || 'N/A'})`;
      } else if (matchQuality === 'EXACT') {
        status = 'MATCHED_RE_ADVANCE';
        statusText = 'Đã khớp tạm ứng lại';
        countMatched++;
        totalReAdvanceAmount += targetAmount;
        auditNotes = `Khớp ${ngayBL} ${thoiGianBL} (BL: ${soBienLai})`;
      } else if (matchQuality === 'APPROX_DEPT') {
        status = 'MATCHED_RE_ADVANCE';
        statusText = 'Đã khớp tiền (Khác khoa)';
        countMatched++;
        totalReAdvanceAmount += targetAmount;
        auditNotes = `Khớp tiền ${formatCurrency(targetAmount)} (Khoa: ${bestMatch.khoa})`;
      } else if (matchQuality === 'DIFF_DATE') {
        status = 'MATCHED_DIFF_DATE';
        statusText = 'Khớp tiền - Lệch ngày';
        countDiffDate++;
        totalReAdvanceAmount += targetAmount;
        auditNotes = `Khám ${ngayBoVe}, ra BL ${ngayBL} ${thoiGianBL} (BL: ${soBienLai})`;
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
        ? `Có ${candidateRows.length} GD nhưng không khớp tiền ${formatCurrency(targetAmount)}`
        : 'Không có dữ liệu trong BC140';
    }

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

export function exportTamUngToExcel(items, summary, fileName = 'Ket_Qua_Ra_Soat_Tam_Ung_Lai.xlsx') {
  const exportData = items.map(item => {
    const orig = item.originalRow || {};
    const rowObj = { ...orig };

    if (item.soBienLai) {
      rowObj['SỐ BIÊN LAI'] = item.soBienLai;
    }

    rowObj['MÃ THU NGÂN (BC140)'] = item.maThuNgan;
    rowObj['HỌ TÊN THU NGÂN (BC140)'] = item.tenThuNgan;
    rowObj['THỜI GIAN BIÊN LAI (BC140)'] = item.thoiGianBL;
    rowObj['NGÀY BIÊN LAI (BC140)'] = item.ngayBL;
    rowObj['HOÀN CHUYỂN KHOẢN (BC140)'] = item.hoanChuyenKhoan;
    rowObj['TRẠNG THÁI RÀ SOÁT'] = item.statusText;
    rowObj['GHI CHÚ ĐỐI SOÁT CHI TIẾT'] = item.auditNotes;

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
