import React from 'react';
import { Monitor } from 'lucide-react';

export default function VanHanhKioskView() {
  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="p-3 bg-teal-50 text-teal-600 rounded-lg">
            <Monitor className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900">Phân hệ Vận hành Kiosk</h1>
            <p className="text-sm text-slate-500">
              Khu vực đang sẵn sàng tiếp nhận các chức năng và nội dung cấu hình tiếp theo.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
