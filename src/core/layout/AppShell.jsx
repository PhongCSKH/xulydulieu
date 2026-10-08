import React, { useState } from 'react';
import { registeredModules } from '../../modules/registry';
import { APP_VERSION } from '../../version';
import { Cpu, ChevronRight, Menu, PanelLeftClose } from 'lucide-react';

export default function AppShell() {
  // Mặc định ẩn thanh danh mục chức năng
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [activeModuleId, setActiveModuleId] = useState(registeredModules[0]?.id);
  const currentModule = registeredModules.find(m => m.id === activeModuleId) || registeredModules[0];
  const ActiveComponent = currentModule.component;

  const handleSelectModule = (id) => {
    setActiveModuleId(id);
    // Tự động thu gọn trên màn hình nhỏ sau khi chọn
    if (typeof window !== 'undefined' && window.innerWidth < 768) {
      setIsSidebarOpen(false);
    }
  };

  return (
    <div className="flex h-screen bg-slate-50 text-slate-900 font-sans overflow-hidden">
      {/* Lớp nền mờ trên màn hình nhỏ khi mở danh mục */}
      {isSidebarOpen && (
        <div
          className="fixed inset-0 bg-slate-950/40 z-30 md:hidden backdrop-blur-xs transition-opacity duration-300"
          onClick={() => setIsSidebarOpen(false)}
        />
      )}

      {/* THANH DANH MỤC (SIDEBAR) - MẶC ĐỊNH ẨN, CHUYỂN CẢNH MƯỢT KHÔNG BỊ TRÀN VIỀN */}
      <aside
        className={`fixed md:relative inset-y-0 left-0 z-40 bg-slate-900 text-slate-300 flex flex-col shrink-0 border-r border-slate-800 transition-[width] duration-300 ease-in-out overflow-hidden select-none ${
          isSidebarOpen
            ? 'w-64 shadow-2xl md:shadow-none'
            : 'w-0 border-r-0 pointer-events-none'
        }`}
      >
        {/* Khung nội dung cố định 256px bên trong để không bị bóp méo chữ trong quá trình mở/đóng */}
        <div className="w-64 min-w-[16rem] flex flex-col h-full flex-shrink-0">
          {/* Header của Sidebar */}
          <div className="p-4 border-b border-slate-800/80 flex items-center justify-between">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-tr from-teal-500 to-emerald-400 flex items-center justify-center text-slate-950 font-bold shadow-sm">
                <Cpu className="w-4 h-4 text-slate-950" />
              </div>
              <div className="flex items-center gap-1.5">
                <span className="font-bold text-white tracking-tight text-sm">CSKH HUB</span>
                <span className="px-1.5 py-0.2 rounded text-[10px] font-semibold bg-teal-950 text-teal-400 border border-teal-800">
                  NỘI BỘ
                </span>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setIsSidebarOpen(false)}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition cursor-pointer"
              title="Thu gọn danh mục"
            >
              <PanelLeftClose className="w-4 h-4" />
            </button>
          </div>

          {/* Danh mục Modules */}
          <div className="flex-1 overflow-y-auto px-3 py-4 space-y-4">
            <div>
              <p className="px-3 text-[11px] font-semibold uppercase tracking-wider text-slate-500 mb-2">
                Danh Mục Chức Năng
              </p>
              <div className="space-y-1">
                {registeredModules.map((module) => {
                  const IconComponent = module.icon;
                  const isActive = module.id === activeModuleId;
                  return (
                    <button
                      key={module.id}
                      onClick={() => handleSelectModule(module.id)}
                      className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs font-medium transition cursor-pointer ${
                        isActive
                          ? 'bg-teal-600 text-white shadow-sm'
                          : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                      }`}
                    >
                      <div className="flex items-center gap-2.5 truncate">
                        <IconComponent className={`w-4 h-4 flex-shrink-0 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                        <span className="truncate">{module.name}</span>
                      </div>
                      {isActive && <ChevronRight className="w-3.5 h-3.5 text-teal-200" />}
                    </button>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Footer Sidebar */}
          <div className="p-3 border-t border-slate-800/80 text-[11px] text-slate-400 flex items-center justify-between">
            <span className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
              Sẵn sàng
            </span>
            <span className="font-mono text-teal-400 font-semibold">{APP_VERSION}</span>
          </div>
        </div>
      </aside>

      {/* KHU VỰC NỘI DUNG CHÍNH */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Top Navbar */}
        <header className="h-12 bg-white border-b border-slate-200 flex items-center justify-between px-4 sm:px-6 flex-shrink-0 shadow-xs">
          <div className="flex items-center gap-3">
            {/* Chỉ hiển thị nút mở khi sidebar đang thu gọn, tránh trùng lặp 2 nút cạnh nhau */}
            {!isSidebarOpen && (
              <button
                type="button"
                onClick={() => setIsSidebarOpen(true)}
                className="p-1.5 rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition border border-slate-200 shadow-2xs flex items-center gap-1.5 text-xs font-medium cursor-pointer"
                title="Mở danh mục chức năng"
              >
                <Menu className="w-4 h-4 text-teal-600" />
                <span className="text-slate-700 hidden sm:inline">Danh mục</span>
              </button>
            )}

            <div className="flex items-center gap-2 text-xs text-slate-500">
              <span className="font-medium text-slate-600">Phòng CSKH</span>
              <span>/</span>
              <span className="font-semibold text-slate-900">{currentModule.name}</span>
            </div>
          </div>

          {/* Phiên bản ứng dụng trên Navbar */}
          <div className="flex items-center gap-2">
            <span className="px-2 py-0.5 rounded-md text-[11px] font-mono font-semibold bg-slate-100 text-slate-700 border border-slate-200">
              {APP_VERSION}
            </span>
          </div>
        </header>

        {/* Viewport chính */}
        <main className="flex-1 overflow-y-auto p-6 bg-slate-50/50">
          <div className="max-w-7xl mx-auto">
            <ActiveComponent />
          </div>
        </main>
      </div>
    </div>
  );
}
