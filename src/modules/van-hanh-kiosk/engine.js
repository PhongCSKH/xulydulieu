import * as XLSX from 'xlsx';

export const DEFAULT_SHEET_CONFIG = {
  spreadsheetId: '1K6JDDbI4A8smzlAb7OjFbZaDKP4vgF_ZQBwFM6zdK8Y',
  gidDuLieuTong: '414639832',     // Sheet DỮ LIỆU TỔNG
  gidViTriMay: '491763635',       // Sheet VỊ TRÍ BỐ TRÍ MÁY
  gidMucTieu: '697870575',        // Sheet MỤC TIÊU
};

export const KIOSK_STORAGE_KEY = 'CSKH_KIOSK_LOCATIONS_CUSTOM';

export function cleanStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

export function normalizeKioskName(val) {
  if (!val) return '';
  return String(val).trim().toUpperCase().replace(/\s+/g, '');
}

/**
 * Parser CSV an toàn xử lý cả trường có dấu phẩy và ngoặc kép
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
      if (char === '\r' && nextChar === '\n') i++;
      currentRow.push(currentField);
      if (currentRow.some(col => col.trim() !== '')) rows.push(currentRow);
      currentRow = [];
      currentField = '';
    } else {
      currentField += char;
    }
  }

  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField);
    if (currentRow.some(col => col.trim() !== '')) rows.push(currentRow);
  }

  if (rows.length === 0) return [];
  const headers = rows[0].map(h => h.trim());
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => {
      if (h) obj[h] = r[idx] !== undefined ? r[idx].trim() : '';
    });
    return obj;
  });
}

/**
 * Kéo dữ liệu trực tiếp từ Google Sheet qua Google Visualization API CSV Export
 */
export async function fetchGoogleSheetData(spreadsheetId, gid) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
  const response = await fetch(url, {
    method: 'GET',
    headers: { 'Accept': 'text/csv,text/plain' }
  });

  if (!response.ok) {
    throw new Error(`Lỗi kết nối Google Sheet (HTTP ${response.status})`);
  }

  const text = await response.text();
  if (text.includes('<!DOCTYPE html>') || text.includes('accounts.google.com') || text.includes('Sign in')) {
    throw new Error('Google Sheet chưa mở quyền "Bất kỳ ai có liên kết đều có thể xem".');
  }

  return parseCSV(text);
}

/**
 * Danh sách vị trí bố trí máy chuẩn gốc
 */
export const DEFAULT_KIOSK_LOCATIONS = [
  { area: 'Trệt A - Sảnh chính', machine: 'KIOS17', config: 'Tự thực hiện' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOS3', config: 'CSKH hỗ trợ' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOS5', config: 'CSKH hỗ trợ' },
  { area: 'Trệt A - Sảnh chính', machine: 'KIOSK-23', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS108', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS110', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS117', config: 'CSKH hỗ trợ' },
  { area: 'VIP', machine: 'KIOS121', config: 'CSKH hỗ trợ' },
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
  { area: 'Trệt B - CĐHA', machine: 'TB-KIOSK-33', config: 'CSKH hỗ trợ' },
];

/**
 * 11 Khu vực chuẩn
 */
export const TARGET_AREAS = [
  '3D - Nhi - Tai mũi họng',
  '2D - Nội thần kinh',
  '2A - Tim Mạch',
  'Trệt A - Sảnh chính',
  '2B - CTCH',
  '1B - Sản',
  '1A - Tiêu hóa',
  'Trệt B - CĐHA',
  'Trệt D - Nội tiết',
  'Trệt F - Da liễu',
  '2D - Mắt',
  'VIP'
];

/**
 * Lưu & nạp danh sách Kiosk tùy chỉnh từ LocalStorage
 */
export function getSavedKioskLocations() {
  try {
    const saved = localStorage.getItem(KIOSK_STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    }
  } catch (e) {
    console.warn('Lỗi đọc KIOSK_STORAGE_KEY:', e);
  }
  return DEFAULT_KIOSK_LOCATIONS;
}

export function saveKioskLocations(locationsList) {
  try {
    localStorage.setItem(KIOSK_STORAGE_KEY, JSON.stringify(locationsList));
  } catch (e) {
    console.warn('Lỗi ghi KIOSK_STORAGE_KEY:', e);
  }
}

export function resetKioskLocationsToDefault() {
  try {
    localStorage.removeItem(KIOSK_STORAGE_KEY);
  } catch (e) {}
  return DEFAULT_KIOSK_LOCATIONS;
}

/**
 * Tải toàn bộ 3 sheet cùng lúc và liên kết dữ liệu
 */
export async function syncAllKioskSheets(config = DEFAULT_SHEET_CONFIG) {
  const [dataRows, locRows, targetRows] = await Promise.all([
    fetchGoogleSheetData(config.spreadsheetId, config.gidDuLieuTong),
    fetchGoogleSheetData(config.spreadsheetId, config.gidViTriMay).catch(() => []),
    fetchGoogleSheetData(config.spreadsheetId, config.gidMucTieu).catch(() => [])
  ]);

  // Kiểm tra xem người dùng có cấu hình tùy chỉnh riêng trên máy không
  const customSaved = localStorage.getItem(KIOSK_STORAGE_KEY);
  let baseLocations = [];

  if (customSaved) {
    try {
      baseLocations = JSON.parse(customSaved);
    } catch (e) {
      baseLocations = [];
    }
  }

  // Nếu chưa có tùy chỉnh riêng thì lấy từ Sheet Vị trí máy
  if (baseLocations.length === 0 && locRows.length > 0) {
    baseLocations = locRows.map(r => ({
      area: r['Khu vực'] || 'Khác',
      machine: r['Tên máy'] || r['Tên KIOS'] || '',
      config: r['Cấu hình'] || 'CSKH hỗ trợ'
    })).filter(k => Boolean(k.machine));
  }

  if (baseLocations.length === 0) {
    baseLocations = DEFAULT_KIOSK_LOCATIONS;
  }

  const locationsMap = {};
  baseLocations.forEach(loc => {
    const machine = normalizeKioskName(loc.machine);
    if (machine) {
      locationsMap[machine] = { area: loc.area, machine: loc.machine, config: loc.config };
    }
  });

  // Phân tích sheet Mục tiêu
  const targetsMap = {};
  let totalTargetCLS = 0;
  let totalTargetThuoc = 0;

  targetRows.forEach(r => {
    const khuVuc = cleanStr(r['KHU VỰC'] || '');
    const cls = parseInt(r['ĐĂNG KÝ & THU CLS'] || r['DANG KY & THU CLS'] || '0', 10) || 0;
    const thuoc = parseInt(r['BÁN THUỐC'] || r['BAN THUOC'] || '0', 10) || 0;
    const clsSang = parseInt(r['Checkin & Thu viện phí (Sáng)'] || '0', 10) || 0;
    const thuocSang = parseInt(r['Bán thuốc (Sáng)'] || '0', 10) || 0;
    const clsChieu = parseInt(r['Checkin & Thu viện phí (Chiều)'] || '0', 10) || 0;
    const thuocChieu = parseInt(r['Bán thuốc (Chiều)'] || '0', 10) || 0;

    if (khuVuc.includes('mục tiêu') || khuVuc.includes('muc tieu') || khuVuc.includes('tổng')) {
      if (cls > 0) totalTargetCLS = cls;
      if (thuoc > 0) totalTargetThuoc = thuoc;
    } else if (khuVuc) {
      targetsMap[khuVuc] = {
        areaRaw: r['KHU VỰC'],
        targetCLS: cls,
        targetThuoc: thuoc,
        totalTarget: cls + thuoc,
        clsSang,
        thuocSang,
        clsChieu,
        thuocChieu
      };
    }
  });

  // Nếu sheet Mục tiêu không có dòng tổng "MỤC TIÊU", tự cộng từ các khoa phòng thực tế trong Sheet
  if (totalTargetCLS === 0 && totalTargetThuoc === 0) {
    Object.values(targetsMap).forEach(t => {
      totalTargetCLS += t.targetCLS;
      totalTargetThuoc += t.targetThuoc;
    });
  }

  return {
    dataRows,
    locationsList: baseLocations,
    locationsMap,
    targetsMap,
    totalTargetCLS,
    totalTargetThuoc,
    totalDailyTarget: totalTargetCLS + totalTargetThuoc
  };
}

/**
 * Thuật toán tổng hợp dữ liệu chính xác 100%
 */
export function aggregateKioskData(rawRows, locationsMap = {}, targetsInfo = {}, filterOptions = {}) {
  const {
    dateFilter = 'ALL',
    selectedShift = 'ALL',
    selectedArea = 'ALL',
    searchTerm = ''
  } = filterOptions;

  const targetDaily = Number(targetsInfo.totalDailyTarget || 0);
  const targetCLS = Number(targetsInfo.totalTargetCLS || 0);
  const targetThuoc = Number(targetsInfo.totalTargetThuoc || 0);
  const targetsMap = targetsInfo.targetsMap || {};

  const filteredRows = rawRows.filter(row => {
    if (dateFilter !== 'ALL') {
      const d = row['NGÀY'] || (row['Ngày giờ thao tác'] || '').split(' ')[0];
      if (d !== dateFilter) return false;
    }
    if (selectedShift !== 'ALL' && row['BUỔI'] !== selectedShift) return false;
    if (selectedArea !== 'ALL' && row['KHU VỰC'] !== selectedArea) return false;

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

  let totalCheckin = 0;
  let totalVienPhi = 0;
  let totalBanThuoc = 0;
  let countCK = 0;
  let countPOS = 0;

  // Dành riêng cho các máy TỰ THỰC HIỆN
  let selfTotal = 0;
  let selfCheckin = 0;
  let selfVienPhi = 0;
  let selfBanThuoc = 0;
  let selfCountCK = 0;
  let selfCountPOS = 0;

  // Quét tìm khung giờ thực tế có trong dữ liệu (không bỏ sót bất kỳ giao dịch nào)
  let minHour = 24;
  let maxHour = -1;

  filteredRows.forEach(row => {
    const rawTime = row['Ngày giờ thao tác'] || '';
    let hour = -1;
    if (row['_hour'] !== undefined) {
      hour = Number(row['_hour']);
    } else if (rawTime.includes(':')) {
      const timePart = rawTime.split(' ')[1] || rawTime;
      hour = parseInt(timePart.split(':')[0], 10);
    }
    if (!isNaN(hour) && hour >= 0 && hour <= 23) {
      if (hour < minHour) minHour = hour;
      if (hour > maxHour) maxHour = hour;
    }
  });

  if (minHour === 24 || maxHour === -1) {
    minHour = 4;
    maxHour = 17;
  } else {
    minHour = Math.min(minHour, 5);
    maxHour = Math.max(maxHour, 17);
  }

  const hourlyGeneral = {};
  const hourlySelf = {};
  for (let h = minHour; h <= maxHour; h++) {
    const label = `${String(h).padStart(2, '0')}:00`;
    hourlyGeneral[h] = { hour: label, label: `${String(h).padStart(2, '0')}h`, h, total: 0, checkin: 0, vienPhi: 0, banThuoc: 0 };
    hourlySelf[h] = { hour: label, label: `${String(h).padStart(2, '0')}h`, h, total: 0, checkin: 0, vienPhi: 0, banThuoc: 0 };
  }

  const areaMap = {};
  TARGET_AREAS.forEach(a => {
    areaMap[a] = { area: a, checkin: 0, vienPhi: 0, banThuoc: 0, total: 0 };
  });

  const kioskMap = {};
  const staffMap = {};

  filteredRows.forEach(row => {
    const hasCheckin = Boolean(row['Checkin'] && String(row['Checkin']).toLowerCase() === 'x');
    const hasVienPhi = Boolean(row['Thu viện phí'] && String(row['Thu viện phí']).toLowerCase() === 'x');
    const hasBanThuoc = Boolean(row['Bán thuốc'] && String(row['Bán thuốc']).toLowerCase() === 'x');

    if (hasCheckin) totalCheckin++;
    if (hasVienPhi) totalVienPhi++;
    if (hasBanThuoc) totalBanThuoc++;

    const ht = String(row['HÌNH THỨC TT'] || '').toUpperCase();
    if (ht.includes('CK')) countCK++;
    else if (ht.includes('POS')) countPOS++;

    // Nhận diện cấu hình máy
    const machine = normalizeKioskName(row['Tên KIOS']);
    const locInfo = locationsMap[machine];
    const locConfig = locInfo?.config || row['_config'] || 'CSKH hỗ trợ';
    const isSelfService = locConfig.toLowerCase().includes('tự thực hiện') || String(row['Họ tên user']).toLowerCase().includes('tự thực hiện');

    const rawTime = row['Ngày giờ thao tác'] || '';
    let hour = -1;
    if (row['_hour'] !== undefined) {
      hour = Number(row['_hour']);
    } else if (rawTime.includes(':')) {
      const timePart = rawTime.split(' ')[1] || rawTime;
      hour = parseInt(timePart.split(':')[0], 10);
    }

    if (!isNaN(hour) && hour >= 0 && hour <= 23) {
      if (!hourlyGeneral[hour]) {
        const label = `${String(hour).padStart(2, '0')}:00`;
        hourlyGeneral[hour] = { hour: label, label: `${String(hour).padStart(2, '0')}h`, h: hour, total: 0, checkin: 0, vienPhi: 0, banThuoc: 0 };
        hourlySelf[hour] = { hour: label, label: `${String(hour).padStart(2, '0')}h`, h: hour, total: 0, checkin: 0, vienPhi: 0, banThuoc: 0 };
      }
      hourlyGeneral[hour].total++;
      if (hasCheckin) hourlyGeneral[hour].checkin++;
      if (hasVienPhi) hourlyGeneral[hour].vienPhi++;
      if (hasBanThuoc) hourlyGeneral[hour].banThuoc++;

      if (isSelfService) {
        hourlySelf[hour].total++;
        if (hasCheckin) hourlySelf[hour].checkin++;
        if (hasVienPhi) hourlySelf[hour].vienPhi++;
        if (hasBanThuoc) hourlySelf[hour].banThuoc++;
      }
    }

    if (isSelfService) {
      selfTotal++;
      if (hasCheckin) selfCheckin++;
      if (hasVienPhi) selfVienPhi++;
      if (hasBanThuoc) selfBanThuoc++;
      if (ht.includes('CK')) selfCountCK++;
      else if (ht.includes('POS')) selfCountPOS++;
    }

    let area = row['KHU VỰC'] || '';
    if ((!area || area === 'Không thấy khu vực') && locInfo?.area) {
      area = locInfo.area;
    }
    if (!area) area = 'Khác';

    if (!areaMap[area]) {
      areaMap[area] = { area, checkin: 0, vienPhi: 0, banThuoc: 0, total: 0 };
    }
    if (hasCheckin) areaMap[area].checkin++;
    if (hasVienPhi) areaMap[area].vienPhi++;
    if (hasBanThuoc) areaMap[area].banThuoc++;
    areaMap[area].total++;

    if (!kioskMap[machine]) {
      kioskMap[machine] = {
        machine,
        area: locInfo?.area || area,
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

    const staff = row['Họ tên user'] || 'Không xác định';
    if (!staff.toLowerCase().includes('tự thực hiện')) {
      if (!staffMap[staff]) {
        staffMap[staff] = { name: staff, total: 0, checkin: 0, vienPhi: 0, banThuoc: 0 };
      }
      staffMap[staff].total++;
      if (hasCheckin) staffMap[staff].checkin++;
      if (hasVienPhi) staffMap[staff].vienPhi++;
      if (hasBanThuoc) staffMap[staff].banThuoc++;
    }
  });

  const totalOps = filteredRows.length;
  const overallRate = targetDaily > 0 ? ((totalOps / targetDaily) * 100).toFixed(1) : '0';
  const regCompletedRate = targetCLS > 0 ? (((totalCheckin + totalVienPhi) / targetCLS) * 100).toFixed(1) : '0';
  const medCompletedRate = targetThuoc > 0 ? ((totalBanThuoc / targetThuoc) * 100).toFixed(1) : '0';

  const totalPaid = (countCK + countPOS) || 1;
  const rateCK = ((countCK / totalPaid) * 100).toFixed(1);
  const ratePOS = ((countPOS / totalPaid) * 100).toFixed(1);

  const hourlyGeneralList = Object.keys(hourlyGeneral).map(Number).sort((a, b) => a - b).map(h => hourlyGeneral[h]);
  const hourlySelfList = Object.keys(hourlySelf).map(Number).sort((a, b) => a - b).map(h => hourlySelf[h]);

  // Đỉnh tải thực tế 100% từ dữ liệu
  const peakHourGeneral = hourlyGeneralList.reduce((max, h) => h.total > max.total ? h : max, { hour: '--', label: '--', total: 0 });
  const peakHourSelf = hourlySelfList.reduce((max, h) => h.total > max.total ? h : max, { hour: '--', label: '--', total: 0 });

  // Năng suất trung bình mỗi giờ có hoạt động
  const activeHoursGeneral = hourlyGeneralList.filter(h => h.total > 0).length || 1;
  const avgHourlyGeneral = (totalOps / activeHoursGeneral).toFixed(1);

  const activeHoursSelf = hourlySelfList.filter(h => h.total > 0).length || 1;
  const avgHourlySelf = (selfTotal / activeHoursSelf).toFixed(1);

  // Mốc thời gian giao dịch sớm nhất và muộn nhất
  let firstActivityTime = '--:--';
  let lastActivityTime = '--:--';
  const validTimes = filteredRows
    .map(r => r['Ngày giờ thao tác'])
    .filter(Boolean)
    .sort();
  if (validTimes.length > 0) {
    firstActivityTime = validTimes[0].split(' ')[1] || validTimes[0];
    lastActivityTime = validTimes[validTimes.length - 1].split(' ')[1] || validTimes[validTimes.length - 1];
  }

  // Phân tích Khoa phòng gắn kèm Mục tiêu thực tế từ Sheet
  const areaBreakdown = Object.values(areaMap).map(item => {
    const key = cleanStr(item.area);
    const targetObj = targetsMap[key] || {
      targetCLS: 0,
      targetThuoc: 0,
      totalTarget: 0,
      clsSang: 0,
      thuocSang: 0,
      clsChieu: 0,
      thuocChieu: 0
    };

    const clsActual = item.checkin + item.vienPhi;
    const medActual = item.banThuoc;
    const totalActual = clsActual + medActual;
    const targetCLS = targetObj.targetCLS || 0;
    const targetThuoc = targetObj.targetThuoc || 0;
    const totalTarget = targetObj.totalTarget || (targetCLS + targetThuoc);

    const rateCLS = targetCLS > 0 ? ((clsActual / targetCLS) * 100).toFixed(1) : (targetCLS === 0 && clsActual > 0 ? '100.0' : '0.0');
    const rateThuoc = targetThuoc > 0 ? ((medActual / targetThuoc) * 100).toFixed(1) : (targetThuoc === 0 && medActual > 0 ? '100.0' : '0.0');
    const rateTotal = totalTarget > 0 ? ((totalActual / totalTarget) * 100).toFixed(1) : '0.0';

    return {
      ...item,
      clsActual,
      medActual,
      totalActual,
      targetCLS,
      targetThuoc,
      totalTarget,
      clsSang: targetObj.clsSang || 0,
      thuocSang: targetObj.thuocSang || 0,
      clsChieu: targetObj.clsChieu || 0,
      thuocChieu: targetObj.thuocChieu || 0,
      rateCLS,
      rateThuoc,
      rateTotal
    };
  }).sort((a, b) => b.total - a.total);

  const kiosks = Object.values(kioskMap).sort((a, b) => b.count - a.count);
  const selfKiosks = kiosks.filter(k => k.config.toLowerCase().includes('tự thực hiện'));

  return {
    summary: {
      totalOps,
      targetDaily,
      overallRate,
      totalCheckin,
      totalVienPhi,
      totalBanThuoc,
      regCompletedRate,
      medCompletedRate,
      rateCK,
      ratePOS,
      countCK,
      countPOS,
      selfTotal,
      selfCheckin,
      selfVienPhi,
      selfBanThuoc,
      selfShareRate: totalOps > 0 ? ((selfTotal / totalOps) * 100).toFixed(1) : '0',
      selfRateCK: (selfCountCK + selfCountPOS) > 0 ? ((selfCountCK / (selfCountCK + selfCountPOS)) * 100).toFixed(1) : '0',
      selfRatePOS: (selfCountCK + selfCountPOS) > 0 ? ((selfCountPOS / (selfCountCK + selfCountPOS)) * 100).toFixed(1) : '0',
      peakHourGeneral,
      peakHourSelf,
      avgHourlyGeneral,
      avgHourlySelf,
      firstActivityTime,
      lastActivityTime,
      activeHoursGeneral,
      activeHoursSelf
    },
    areaBreakdown,
    hourlyGeneralList,
    hourlySelfList,
    kiosks,
    selfKiosks,
    staff: Object.values(staffMap).sort((a, b) => b.total - a.total),
    filteredRows
  };
}

export function exportKioskReportToExcel(rows, fileName = 'Bao_Cao_Hoat_Dong_Kiosk.xlsx') {
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
  XLSX.utils.book_append_sheet(workbook, worksheet, 'HoatDongKiosk');
  XLSX.writeFile(workbook, fileName);
}
