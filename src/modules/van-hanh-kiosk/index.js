import VanHanhKioskView from './VanHanhKioskView';
import { MonitorCheck } from 'lucide-react';

export default {
  id: 'MOD_VAN_HANH_KIOSK',
  name: 'Vận Hành Kiosk',
  shortCode: 'KIOSK',
  category: 'Điều Hành Dịch Vụ',
  description: 'Trung tâm giám sát vận hành Kiosk thông minh: đồng bộ Google Sheet thời gian thực, lưu lượng giờ cao điểm, ma trận thiết bị và năng suất nhân sự.',
  icon: MonitorCheck,
  component: VanHanhKioskView
};
