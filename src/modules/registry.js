import bv01Module from './bv01-chua-ra-bill';
import raSoatTamUngModule from './ra-soat-tam-ung';
import vanHanhKioskModule from './van-hanh-kiosk';

/**
 * Danh bạ Đăng ký Module Độc lập (Pluggable Registry)
 * Khi có chức năng mới, chỉ cần tạo thư mục riêng trong src/modules/ và import vào danh sách bên dưới.
 * Đảm bảo 100% không làm ảnh hưởng đến các chức năng khác!
 */
export const registeredModules = [
  bv01Module,
  raSoatTamUngModule,
  vanHanhKioskModule,
];

export function getModuleById(id) {
  return registeredModules.find(m => m.id === id) || registeredModules[0];
}
