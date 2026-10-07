import * as XLSX from 'xlsx';

/**
 * CẤU HÌNH GOOGLE SHEET MẶC ĐỊNH
 */
export const DEFAULT_SHEET_CONFIG = {
  spreadsheetId: '1K6JDDbI4A8smzlAb7OjFbZaDKP4vgF_ZQBwFM6zdK8Y',
  gidDuLieuTong: '414639832',     // Sheet DỮ LIỆU TỔNG
  gidViTriMay: '491763635',       // Sheet VỊ TRÍ BỐ TRÍ MÁY
  gidMucTieu: '',                 // Sheet MỤC TIÊU
};

/**
 * Chuẩn hóa chuỗi an toàn
 */
export function cleanStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Chuẩn hóa tên máy Kiosk (bỏ ký tự thừa, chữ hoa)
 * Ví dụ: "kiosk-23" -> "KIOSK-23", "kios 17" -> "KIOS17"
 */
export function normalizeKioskName(val) {
  if (!val) return '';
  return String(val).trim().toUpperCase().replace(/\s+/g, '');
}

/**
 * Parser đơn giản cho dòng CSV trả về từ Google Sheets
 */
export function parseCSV(csvText) {
  const rows = [];
  let currentRow = [];
  let currentField = '';
  let insideQuotes = false;

  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];

    if (char === '"') {
      if (insideQuotes && nextChar === '"') {
        currentField += '"';
        i++;
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (char === ',' && !insideQuotes) {
      currentRow.push(currentField);
      currentField = '';
    } else if ((char === '\r' || char === '\n') && !insideQuotes) {
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
      currentRow.push(currentField);
      if (currentRow.some(col => col.trim() !== '')) {
        rows.push(currentRow);
      }
      currentRow = [];
      currentField = '';
    } else {
      currentField += char;
    }
  }

  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField);
    if (currentRow.some(col => col.trim() !== '')) {
      rows.push(currentRow);
    }
  }

  if (rows.length === 0) return [];
  const headers = rows[0].map(h => h.trim());
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = r[idx] !== undefined ? r[idx].trim() : '';
    });
    return obj;
  });
}

/**
 * Tải dữ liệu trực tiếp từ Google Sheet thông qua Google Visualization API CSV Export
 */
export async function fetchGoogleSheetData(spreadsheetId, gid) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
  const response = await fetch(url, {
    method: 'GET',
    headers: { 'Accept': 'text/csv,text/plain' }
  });

  if (!response.ok) {
    throw new Error(`Không thể kết nối Google Sheet (HTTP ${response.status})`);
  }

  const text = await response.text();
  // Kiểm tra nếu trả về trang HTML đăng nhập
  if (text.includes('<!DOCTYPE html>') || text.includes('accounts.google.com') || text.includes('Sign in')) {
    throw new Error('Google Sheet chưa được bật quyền công khai ("Bất kỳ ai có liên kết đều có thể xem").');
  }

  return parseCSV(text);
}

/**
 * Danh sách vị trí bố trí máy mặc định (đối chiếu nếu sheet vị trí chưa kịp tải)
 */
export const DEFAULT_KIOSK_LOCATIONS = [
  { area: 'VIP', machine: 'KIOS108', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS110', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS117', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS121', config: 'CSKH hỗ trợ' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOS17', config: 'Tự thực hiện' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOS3', config: 'CSKH hỗ trợ' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOS5', config: 'CSKH hỗ trợ' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOSK-23', config: 'CSKH hỗ trợ' },
  { area: 'Trệt B - CĐHA', machine: 'KIOS35', config: 'CSKH hỗ trợ' },
  { area: 'Trệt B - CĐHA', machine: 'KIOS13', config: 'CSKH hỗ trợ' },
  { area: '1B - Sản', machine: 'KIOSK-25', config: 'CSKH hỗ trợ' },
  { area: '1B - Sản', machine: 'KIOS16', config: 'CSKH hỗ trợ' },
  { area: '1B - Sản', machine: 'KIOS40', config: 'CSKH hỗ trợ' },
  { area: '1A - Tiêu hóa', machine: 'KIOS28', config: 'CSKH hỗ trợ' },
  { area: '1A - Tiêu hóa', machine: 'KIOS29', config: 'CSKH hỗ trợ' },
  { area: '2A - Tim Mạch', machine: 'KIOS256', config: 'CSKH hỗ trợ' },
  { area: '2A - Tim Mạch', machine: 'KIOS260', config: 'CSKH hỗ trợ' },
  { area: '2A - Tim Mạch', machine: 'KIOS39', config: 'CSKH hỗ trợ' },
  { area: '2B - CTCH', machine: 'KIOS34', config: 'CSKH hỗ trợ' },
  { area: '2B - CTCH', machine: 'KIOS41', config: 'CSKH hỗ trợ' },
  { area: '2B - CTCH', machine: 'KIOS43', config: 'CSKH hỗ trợ' },
  { area: 'Trệt D - Nội tiết', machine: 'KIOS42', config: 'CSKH hỗ trợ' },
  { area: 'Trệt D - Nội tiết', machine: 'KIOSK-26', config: 'CSKH hỗ trợ' },
  { area: '2D - Nội thần kinh', machine: 'KIOS30', config: 'CSKH hỗ trợ' },
  { area: '2D - Nội thần kinh', machine: 'KIOS31', config: 'CSKH hỗ trợ' },
  { area: '2D - Nội thần kinh', machine: 'KIOS36', config: 'CSKH hỗ trợ' },
  { area: '2D - Mắt', machine: 'KIOSK-24', config: 'CSKH hỗ trợ' },
  { area: '3D - Nhi - Tai mũi họng', machine: 'KIOSK-27', config: 'CSKH hỗ trợ' },
  { area: '3D - Nhi - Tai mũi họng', machine: 'KIOS37', config: 'CSKH hỗ trợ' },
  { area: '3D - Nhi - Tai mũi họng', machine: 'KIOS38', config: 'CSKH hỗ trợ' },
  { area: 'Trệt F - Da liễu', machine: 'KIOS44', config: 'CSKH hỗ trợ' },
  { area: 'Trệt C - Khám Ngoại', machine: 'TB-KIOSK-33', config: 'CSKH hỗ trợ' },
];

/**
 * Sinh bộ dữ liệu mẫu ban đầu mô phỏng chính xác từ Google Sheet
 */
export function generateSampleKioskData() {
  const staffList = [
    'Hoàng Thị Huyền Trang',
    'Phạm Hoài Linh',
    'Nguyễn Thụy Trúc Quỳnh',
    'Phạm Ngọc Hải',
    'Trần Thanh Thảo',
    'Lê Minh Quân',
    'Võ Thị Ngọc Mai',
    'Đặng Thu Hà'
  ];

  const patients = [
    { pid: '2610153896', name: 'TA HỒ PHƯỚC MINH' },
    { pid: '2530211605', name: 'NGUYỄN LÊ HƯNG THỊNH' },
    { pid: '22972232', name: 'VÕ THỊ MỸ PHƯƠNG' },
    { pid: '2510043796', name: 'TRƯƠNG QUỲNH MY' },
    { pid: '2510082940', name: 'HỨA THỊ ÚT NHI' },
    { pid: '21037640', name: 'TRỊNH HẢI PHƯƠNG' },
    { pid: '2410255651', name: 'NGUYỄN HOÀNG VƯƠNG' },
    { pid: '2510280502', name: 'ĐOÀN NGỌC TÚ LINH' },
    { pid: '2610130113', name: 'DƯƠNG KHAI NGUYÊN' },
    { pid: '2510245881', name: 'PHẠM DƯƠNG QUỐC KHIÊM' },
    { pid: '22906609', name: 'BÙI QUANG MINH' },
    { pid: '2410236780', name: 'TRẦN NGUYỄN TUỆ AN' },
    { pid: '2410047966', name: 'HUỲNH PHƯƠNG ANH' },
    { pid: '2610026610', name: 'LÊ CÁT HẠ MY' },
    { pid: '2510280701', name: 'NGUYỄN GIA BÁCH' },
    { pid: '2410011761', name: 'THÁI CAO ĐỊNH' },
    { pid: '2510075123', name: 'TRẦN THỊ THU' },
    { pid: '2510307849', name: 'NGUYỄN MINH QUÂN' },
    { pid: '2610233458', name: 'TRƯƠNG THỊ CHÁNH' },
    { pid: '2610233365', name: 'QUÁCH LÝ KIM THẢO' },
    { pid: '2610160271', name: 'LÊ THỊ ÁI VY' },
    { pid: '2410150927', name: 'ĐINH NGUYỄN PHƯƠNG LINH' }
  ];

  const rawRows = [];
  const currentDate = '07/10/2026';
  let stt = 1;

  // Tạo khoảng 1200 records giao dịch mẫu phân bổ theo ngày hôm nay và hôm qua
  for (let hour = 6; hour <= 17; hour++) {
    // Lưu lượng tăng cao từ 7-10h sáng và 13-15h chiều
    const trafficWeight = (hour >= 7 && hour <= 10) ? 90 : (hour >= 13 && hour <= 15) ? 65 : 25;

    for (let i = 0; i < trafficWeight; i++) {
      const min = Math.floor(Math.random() * 60).toString().padStart(2, '0');
      const timeStr = `${hour.toString().padStart(2, '0')}:${min}`;
      const dateTimeStr = `${currentDate} ${timeStr}`;

      const pIdx = Math.floor(Math.random() * patients.length);
      const p = patients[pIdx];

      const kIdx = Math.floor(Math.random() * DEFAULT_KIOSK_LOCATIONS.length);
      const kInfo = DEFAULT_KIOSK_LOCATIONS[kIdx];

      const staff = kInfo.config === 'Tự thực hiện' ? 'Bệnh nhân tự thực hiện' : staffList[Math.floor(Math.random() * staffList.length)];
      const buoi = hour < 12 ? 'Buổi sáng' : 'Buổi chiều';

      const serviceRoll = Math.random();
      const isCheckin = serviceRoll < 0.55 ? 'x' : '';
      const isVienPhi = serviceRoll >= 0.45 && serviceRoll < 0.90 ? 'x' : '';
      const isBanThuoc = serviceRoll >= 0.85 ? 'x' : '';

      const hinhThucTT = (isVienPhi || isBanThuoc) ? (Math.random() > 0.35 ? 'CK' : 'POS') : '';

      rawRows.push({
        'STT': stt++,
        'Ngày giờ thao tác': dateTimeStr,
        'PID': p.pid,
        'Họ tên': p.name,
        'Checkin': isCheckin || (serviceRoll < 0.3 ? 'x' : ''),
        'Thu viện phí': isVienPhi,
        'Bán thuốc': isBanThuoc,
        'Họ tên user': staff,
        'Tên KIOS': kInfo.machine,
        'KHU VỰC': kInfo.area,
        'HÌNH THỨC TT': hinhThucTT,
        'NGÀY': currentDate,
        'BUỔI': buoi,
        '_hour': hour,
        '_time': timeStr,
        '_config': kInfo.config
      });
    }
  }

  return rawRows;
}

/**
 * Xử lý tính toán Aggregation toàn diện cho Executive Dashboard
 */
export function aggregateKioskData(rawRows, locationsMap = {}, filterOptions = {}) {
  const {
    dateFilter = 'ALL',
    selectedShift = 'ALL',
    selectedArea = 'ALL',
    selectedService = 'ALL',
    searchTerm = ''
  } = filterOptions;

  // Lọc dữ liệu theo tiêu chí
  const filteredRows = rawRows.filter(row => {
    // Lọc theo ngày
    if (dateFilter !== 'ALL') {
      const rowDate = row['NGÀY'] || (row['Ngày giờ thao tác'] || '').split(' ')[0];
      if (rowDate !== dateFilter) return false;
    }

    // Lọc theo buổi
    if (selectedShift !== 'ALL') {
      const buoi = row['BUỔI'] || '';
      if (buoi !== selectedShift) return false;
    }

    // Lọc theo khu vực
    if (selectedArea !== 'ALL') {
      const area = row['KHU VỰC'] || '';
      if (area !== selectedArea) return false;
    }

    // Lọc theo loại hình dịch vụ
    if (selectedService === 'CHECKIN' && !row['Checkin']) return false;
    if (selectedService === 'VIEN_PHI' && !row['Thu viện phí']) return false;
    if (selectedService === 'BAN_THUOC' && !row['Bán thuốc']) return false;

    // Lọc tìm kiếm
    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      const matchPid = String(row['PID'] || '').toLowerCase().includes(q);
      const matchName = String(row['Họ tên'] || '').toLowerCase().includes(q);
      const matchKiosk = String(row['Tên KIOS'] || '').toLowerCase().includes(q);
      const matchStaff = String(row['Họ tên user'] || '').toLowerCase().includes(q);
      return matchPid || matchName || matchKiosk || matchStaff;
    }

    return true;
  });

  // KPI Cơ bản
  let totalCheckin = 0;
  let totalVienPhi = 0;
  let totalBanThuoc = 0;
  let countCK = 0;
  let countPOS = 0;
  let countSelfService = 0;
  let countAssisted = 0;

  // Map theo giờ (06:00 - 18:00)
  const hourlyMap = {};
  for (let h = 6; h <= 18; h++) {
    hourlyMap[h] = { hour: `${h}:00`, total: 0, checkin: 0, vienPhi: 0, banThuoc: 0 };
  }

  // Map theo Kiosk
  const kioskMap = {};
  // Map theo Nhân sự CSKH
  const staffMap = {};
  // Map theo Khu vực
  const areaMap = {};

  filteredRows.forEach(row => {
    const hasCheckin = Boolean(row['Checkin'] && String(row['Checkin']).toLowerCase() === 'x');
    const hasVienPhi = Boolean(row['Thu viện phí'] && String(row['Thu viện phí']).toLowerCase() === 'x');
    const hasBanThuoc = Boolean(row['Bán thuốc'] && String(row['Bán thuốc']).toLowerCase() === 'x');

    if (hasCheckin) totalCheckin++;
    if (hasVienPhi) totalVienPhi++;
    if (hasBanThuoc) totalBanThuoc++;

    // Hình thức TT
    const ht = String(row['HÌNH THỨC TT'] || '').toUpperCase();
    if (ht.includes('CK')) countCK++;
    else if (ht.includes('POS')) countPOS++;

    // Cấu hình máy (Tự thực hiện vs CSKH)
    const machine = normalizeKioskName(row['Tên KIOS']);
    const locConfig = locationsMap[machine]?.config || row['_config'] || 'CSKH hỗ trợ';
    const isSelf = locConfig.toLowerCase().includes('tự thực hiện') || String(row['Họ tên user']).toLowerCase().includes('tự thực hiện');

    if (isSelf) {
      countSelfService++;
    } else {
      countAssisted++;
    }

    // Phân tích theo giờ
    const rawTime = row['Ngày giờ thao tác'] || '';
    let hour = -1;
    if (row['_hour'] !== undefined) {
      hour = Number(row['_hour']);
    } else if (rawTime.includes(':')) {
      const timePart = rawTime.split(' ')[1] || rawTime;
      const hStr = timePart.split(':')[0];
      hour = parseInt(hStr, 10);
    }

    if (hour >= 6 && hour <= 18 && hourlyMap[hour]) {
      hourlyMap[hour].total++;
      if (hasCheckin) hourlyMap[hour].checkin++;
      if (hasVienPhi) hourlyMap[hour].vienPhi++;
      if (hasBanThuoc) hourlyMap[hour].banThuoc++;
    }

    // Kiosk Grouping
    const area = row['KHU VỰC'] || locationsMap[machine]?.area || 'Chưa định vị';
    if (!kioskMap[machine]) {
      kioskMap[machine] = {
        machine,
        area,
        config: locConfig,
        count: 0,
        checkin: 0,
        vienPhi: 0,
        banThuoc: 0,
        lastActiveTime: rawTime
      };
    }
    kioskMap[machine].count++;
    if (hasCheckin) kioskMap[machine].checkin++;
    if (hasVienPhi) kioskMap[machine].vienPhi++;
    if (hasBanThuoc) kioskMap[machine].banThuoc++;
    kioskMap[machine].lastActiveTime = rawTime;

    // Staff Grouping (bỏ qua bệnh nhân tự thực hiện)
    const staff = row['Họ tên user'] || 'Không xác định';
    if (!staff.toLowerCase().includes('tự thực hiện')) {
      if (!staffMap[staff]) {
        staffMap[staff] = {
          name: staff,
          total: 0,
          checkin: 0,
          vienPhi: 0,
          banThuoc: 0,
          morningCount: 0,
          afternoonCount: 0
        };
      }
      staffMap[staff].total++;
      if (hasCheckin) staffMap[staff].checkin++;
      if (hasVienPhi) staffMap[staff].vienPhi++;
      if (hasBanThuoc) staffMap[staff].banThuoc++;
      if (row['BUỔI'] === 'Buổi sáng') staffMap[staff].morningCount++;
      else staffMap[staff].afternoonCount++;
    }

    // Area Grouping
    if (!areaMap[area]) {
      areaMap[area] = { area, total: 0, machines: new Set() };
    }
    areaMap[area].total++;
    areaMap[area].machines.add(machine);
  });

  const totalOps = filteredRows.length;
  const selfServiceRate = totalOps > 0 ? ((countSelfService / totalOps) * 100).toFixed(1) : 0;
  const assistedRate = totalOps > 0 ? ((countAssisted / totalOps) * 100).toFixed(1) : 0;

  // Sắp xếp Kiosk theo số lượng giao dịch giảm dần
  const kioskRankings = Object.values(kioskMap).sort((a, b) => b.count - a.count);

  // Sắp xếp Nhân sự theo năng suất giảm dần
  const staffRankings = Object.values(staffMap).sort((a, b) => b.total - a.total);

  // Sắp xếp Khu vực
  const areaRankings = Object.values(areaMap)
    .map(a => ({ area: a.area, total: a.total, machineCount: a.machines.size }))
    .sort((a, b) => b.total - a.total);

  // Danh sách các khung giờ
  const hourlyData = Object.keys(hourlyMap)
    .map(Number)
    .sort((a, b) => a - b)
    .map(h => hourlyMap[h]);

  return {
    summary: {
      totalOps,
      totalCheckin,
      totalVienPhi,
      totalBanThuoc,
      countCK,
      countPOS,
      countSelfService,
      countAssisted,
      selfServiceRate,
      assistedRate,
      activeKiosksCount: Object.keys(kioskMap).length,
      activeStaffCount: Object.keys(staffMap).length,
      // Mục tiêu KPI mặc định (ví dụ: 1200 giao dịch/ngày)
      targetOps: 1200,
      targetPercent: Math.min(100, Math.round((totalOps / 1200) * 100))
    },
    hourlyData,
    kioskRankings,
    staffRankings,
    areaRankings,
    filteredRows
  };
}

/**
 * Xuất dữ liệu Kiosk ra file Excel chuẩn mực
 */
export function exportKioskReportToExcel(rows, fileName = 'Bao_Cao_Van_Hanh_Kiosk.xlsx') {
  const exportData = rows.map((r, i) => ({
    'STT': r['STT'] || (i + 1),
    'Ngày giờ thao tác': r['Ngày giờ thao tác'] || '',
    'PID': r['PID'] || '',
    'Họ tên': r['Họ tên'] || '',
    'Checkin': r['Checkin'] || '',
    'Thu viện phí': r['Thu viện phí'] || '',
    'Bán thuốc': r['Bán thuốc'] || '',
    'Họ tên user': r['Họ tên user'] || '',
    'Tên KIOS': r['Tên KIOS'] || '',
    'Khu vực': r['KHU VỰC'] || '',
    'Hình thức TT': r['HÌNH THỨC TT'] || '',
    'Ngày': r['NGÀY'] || '',
    'Buổi': r['BUỔI'] || ''
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'VanHanhKiosk');

  worksheet['!cols'] = [
    { wch: 8 },
    { wch: 20 },
    { wch: 14 },
    { wch: 26 },
    { wch: 10 },
    { wch: 14 },
    { wch: 12 },
    { wch: 26 },
    { wch: 16 },
    { wch: 24 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 }
  ];

  XLSX.writeFile(workbook, fileName);
}
