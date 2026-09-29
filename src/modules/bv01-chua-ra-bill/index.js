import BV01ModuleView from './BV01ModuleView';
import { FileCheck } from 'lucide-react';

export default {
  id: 'MOD_RECON_BV01',
  name: 'Chốt BV01 Chưa Ra Bill',
  shortCode: 'BV01',
  category: 'Đối Soát Viện Phí',
  description: 'Đối chiếu hồ sơ chốt viện phí (BC121) với biên lai thực thu (BC140) theo cơ chế phân luồng thông minh 3 trạng thái.',
  icon: FileCheck,
  component: BV01ModuleView
};
