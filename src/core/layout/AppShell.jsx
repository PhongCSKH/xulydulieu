import React, { useState } from 'react';
import { registeredModules } from '../../modules/registry';
import { Cpu, ChevronRight } from 'lucide-react';

export default function AppShell() {
  const [activeModuleId, setActiveModuleId] = useState(registeredModules[0]?.id);
  const currentModule = registeredModules.find(m => m.id === activeModuleId) || registeredModules[0];
  const ActiveComponent = currentModule.component;

  return (
    <div className="flex h-screen bg-slate-50 text-slate-900 font-sans overflow-hidden">
      {/* SIDEBAR BÊN TRÁI */}
      <aside className="w-64 bg-slate-900 text-slate-300 flex flex-col flex-shrink-0 border-r border-slate-800">
        {/* Brand Header */}
        <div className="p-4 border-b border-slate-800/80">
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
                    onClick={() => setActiveModuleId(module.id)}
                    className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs font-medium transition ${
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
          <span className="font-mono text-slate-400">v1.0</span>
        </div>
      </aside>

      {/* KHU VỰC NỘI DUNG CHÍNH */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Top Navbar */}
        <header className="h-12 bg-white border-b border-slate-200 flex items-center justify-between px-6 flex-shrink-0">
          <div className="flex items-center gap-2 text-xs text-slate-500">
            <span className="font-medium text-slate-600">Phòng CSKH</span>
            <span>/</span>
            <span className="font-semibold text-slate-900">{currentModule.name}</span>
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
