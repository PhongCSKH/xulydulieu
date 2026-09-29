import RaSoatTamUngView from './RaSoatTamUngView';
import { ReceiptText } from 'lucide-react';

export default {
  id: 'MOD_RA_SOAT_TAM_UNG',
  name: 'Rà Soát Tạm Ứng Lại',
  shortCode: 'RSTU',
  category: 'Đối Soát Viện Phí',
  description: 'Đối chiếu hồ sơ KH Bỏ Về với BC140 theo Ngày, Mã BN, Khoa phòng, Số tiền; trích xuất Mã/Tên thu ngân, Số BL, Thời gian và Hoàn chuyển khoản.',
  icon: ReceiptText,
  component: RaSoatTamUngView
};
