import * as XLSX from 'xlsx';

/**
 * Chuẩn hóa chuỗi (loại bỏ khoảng trắng thừa, đưa về chữ thường)
 */
export function cleanStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Định dạng ngày sang DD/MM/YYYY chuẩn mực
 */
export function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const d = val.getDate().toString().padStart(2, '0');
    const m = (val.getMonth() + 1).toString().padStart(2, '0');
    const y = val.getFullYear();
    return `${d}/${m}/${y}`;
  }
  if (typeof val === 'number') {
    // Excel serial date to JS date
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed) {
      const d = String(parsed.d).padStart(2, '0');
      const m = String(parsed.m).padStart(2, '0');
      const y = parsed.y;
      return `${d}/${m}/${y}`;
    }
  }
  if (typeof val === 'string') {
    const m = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) {
      const d = m[1].padStart(2, '0');
      const mo = m[2].padStart(2, '0');
      return `${d}/${mo}/${m[3]}`;
    }
  }
  return String(val);
}

/**
 * Chuyển DD/MM/YYYY thành timestamp để tính số ngày chênh lệch
 */
function parseDateToTime(dateStr) {
  if (!dateStr) return null;
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0])).getTime();
  }
  return null;
}

/**
 * Tự động phát hiện loại file (BC121 hay BC140) dựa trên tiêu đề cột
 */
export function detectFileType(sheet) {
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (!json || json.length === 0) return 'UNKNOWN';
  const header = (json[0] || []).map(cleanStr);
  
  if (header.includes('user chốt bv01') || header.includes('ngày chốt bv') || header.includes('kh bỏ về')) {
    return 'BC121';
  }
  if (header.includes('quyển số') || header.includes('số biên lai') || header.includes('bệnh nhân thanh toán') || header.includes('tổng thanh toán')) {
    return 'BC140';
  }
  return 'UNKNOWN';
}

/**
 * Thuật toán đối soát thông minh (Phương án 3)
 */
export function processReconciliation(rows121, rows140) {
  // 1. Lọc BC140 chỉ lấy nhóm 'Viện phí' và lập chỉ mục theo Mã bệnh nhân
  const map140ByBN = new Map();
  rows140.forEach(row => {
    const nhom = cleanStr(row['Nhóm'] || row['nhóm']);
    if (nhom === 'viện phí' || nhom === 'vien phi') {
      const bn = cleanStr(row['Mã bệnh nhân'] || row['Mã BN'] || row['mã bệnh nhân']);
      if (bn) {
        if (!map140ByBN.has(bn)) map140ByBN.set(bn, []);
        map140ByBN.get(bn).push(row);
      }
    }
  });

  const results = [];
  let countMatchedOnTime = 0;
  let countMatchedLate = 0;
  let countPendingBill = 0;
  let totalPendingMoney = 0;
  let totalCollectedMoney = 0;

  // 2. Duyệt từng dòng trong BC121
  rows121.forEach((r121, index) => {
    const maBN = cleanStr(r121['Mã bệnh nhân'] || r121['Mã BN']);
    if (!maBN) return;

    const hoTen = String(r121['Họ tên'] || r121['Họ và tên'] || '').trim();
    const ngayVao = formatDate(r121['Ngày vào']);
    const ngayRa = formatDate(r121['Ngày ra']);
    const ngayChotRaw = r121['Ngày chốt BV'] || r121['Ngày'] || r121['Ngày giờ'];
    const ngayChot = formatDate(ngayChotRaw);
    const khoaPhong = String(r121['Khoa phòng'] || '').trim();
    const khoaClean = cleanStr(khoaPhong);
    const tienBNTra121 = Number(r121['Bệnh nhân trả'] || 0);

    const candidateRows = map140ByBN.get(maBN) || [];

    // Tìm kiếm biên lai khớp
    let bestMatch = null;
    let matchType = 'PENDING_BILL'; // Mặc định là Chưa ra bill

    if (candidateRows.length > 0) {
      // Ưu tiên 1: Khớp chuẩn xác cả Khoa phòng và Ngày (Cùng ngày chốt)
      bestMatch = candidateRows.find(r => {
        const ngay140 = formatDate(r['Ngày']);
        const khoa140 = cleanStr(r['Khoa phòng']);
        return (ngay140 === ngayChot || ngay140 === ngayVao) && khoa140 === khoaClean;
      });

      if (bestMatch) {
        matchType = 'MATCHED_ON_TIME';
      } else {
        // Ưu tiên 2: Khớp cùng Khoa phòng nhưng khác ngày (Ra bill trễ hạn)
        bestMatch = candidateRows.find(r => {
          const khoa140 = cleanStr(r['Khoa phòng']);
          return khoa140 === khoaClean;
        });

        if (bestMatch) {
          const tChot = parseDateToTime(ngayChot);
          const t140 = parseDateToTime(formatDate(bestMatch['Ngày']));
          if (tChot && t140 && t140 > tChot) {
            matchType = 'MATCHED_LATE';
          } else {
            matchType = 'MATCHED_ON_TIME';
          }
        } else {
          // Ưu tiên 3: Nếu khoa phòng có biến thể nhỏ nhưng trùng mã BN trong ngày chốt
          bestMatch = candidateRows.find(r => {
            const ngay140 = formatDate(r['Ngày']);
            return ngay140 === ngayChot;
          });
          if (bestMatch) {
            matchType = 'MATCHED_ON_TIME';
          }
        }
      }
    }

    // Tổng hợp thông tin
    let status = matchType;
    let statusText = 'Chưa ra Bill';
    let soBienLai = '';
    let ngayBill = '';
    let tienThanhToan140 = 0;
    let soNgayTre = 0;
    let notes = '';

    if (bestMatch) {
      soBienLai = bestMatch['Số biên lai'] || bestMatch['Số BL'] || '';
      ngayBill = formatDate(bestMatch['Ngày']);
      tienThanhToan140 = Number(bestMatch['Bệnh nhân thanh toán'] || bestMatch['Tổng cộng biên lai'] || 0);
      totalCollectedMoney += tienThanhToan140;

      if (status === 'MATCHED_ON_TIME') {
        statusText = 'Đã ra Bill';
        countMatchedOnTime++;
      } else {
        statusText = 'Ra Bill trễ';
        countMatchedLate++;
        const tChot = parseDateToTime(ngayChot);
        const t140 = parseDateToTime(ngayBill);
        if (tChot && t140) {
          soNgayTre = Math.max(1, Math.round((t140 - tChot) / (1000 * 60 * 60 * 24)));
        } else {
          soNgayTre = 1;
        }
        notes = `Chốt ${ngayChot}, ra bill ${ngayBill} (Trễ ${soNgayTre} ngày)`;
      }
    } else {
      status = 'PENDING_BILL';
      statusText = 'Chưa ra Bill';
      countPendingBill++;
      totalPendingMoney += tienBNTra121;
      notes = 'Không tìm thấy biên lai trong BC140 (Cần kiểm tra thu ngân)';
    }

    results.push({
      stt: index + 1,
      maBN: r121['Mã bệnh nhân'] || r121['Mã BN'],
      hoTen,
      ngaySinh: formatDate(r121['Ngày sinh']),
      khoaPhong,
      ngayVao,
      ngayRa,
      ngayChot,
      userChot: r121['User chốt BV01'] || '',
      nhanVienChot: r121['Nhân viên chốt BV01'] || '',
      tienBNTra121,
      status,
      statusText,
      soBienLai,
      ngayBill,
      tienThanhToan140,
      soNgayTre,
      chenhLechTien: tienThanhToan140 - tienBNTra121,
      notes
    });
  });

  const total = results.length;
  const percentCompleted = total > 0 ? Math.round(((countMatchedOnTime + countMatchedLate) / total) * 100) : 0;

  return {
    summary: {
      total,
      countMatchedOnTime,
      countMatchedLate,
      countPendingBill,
      percentCompleted,
      totalPendingMoney,
      totalCollectedMoney
    },
    items: results
  };
}

/**
 * Xuất kết quả đối soát ra tệp Excel chuẩn mực
 */
export function exportToExcel(items, fileName = 'Ket_Qua_Doi_Soat_BV01_Chua_Ra_Bill.xlsx') {
  const exportData = items.map(item => ({
    'STT': item.stt,
    'Mã Bệnh Nhân': item.maBN,
    'Họ Và Tên': item.hoTen,
    'Khoa Phòng': item.khoaPhong,
    'Ngày Vào': item.ngayVao,
    'Ngày Chốt BV01': item.ngayChot,
    'Trạng Thái Bill': item.statusText,
    'Số Biên Lai (BC140)': item.soBienLai,
    'Ngày Ra Bill': item.ngayBill,
    'Số Ngày Trễ': item.soNgayTre > 0 ? item.soNgayTre : '',
    'Bệnh Nhân Trả (BC121)': item.tienBNTra121,
    'Thực Thu (BC140)': item.tienThanhToan140,
    'Chênh Lệch': item.chenhLechTien,
    'User Chốt BV01': item.userChot,
    'Ghi Chú Nghiệp Vụ': item.notes
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'KetQuaDoiSoat');

  // Đặt độ rộng cột tối ưu
  worksheet['!cols'] = [
    { wch: 6 },  // STT
    { wch: 14 }, // Mã BN
    { wch: 24 }, // Họ tên
    { wch: 26 }, // Khoa phòng
    { wch: 12 }, // Ngày vào
    { wch: 14 }, // Ngày chốt
    { wch: 16 }, // Trạng thái
    { wch: 16 }, // Số biên lai
    { wch: 12 }, // Ngày ra bill
    { wch: 12 }, // Số ngày trễ
    { wch: 18 }, // BN trả 121
    { wch: 18 }, // Thực thu 140
    { wch: 14 }, // Chênh lệch
    { wch: 16 }, // User chốt
    { wch: 38 }  // Ghi chú
  ];

  XLSX.writeFile(workbook, fileName);
}
