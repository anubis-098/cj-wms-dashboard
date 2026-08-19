import { LayoutDashboard, Upload, Settings, SlidersHorizontal } from "lucide-react";

import type { DashboardResponse } from "../types/dashboard";

type AdminDashboardPageProps = {
  dashboard: DashboardResponse | null;
  onNavigate: (path: string) => void;
};

export function AdminDashboardPage({ dashboard, onNavigate }: AdminDashboardPageProps) {
  return (
    <main className="min-h-screen bg-screen-bg p-6">
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-black text-cj-navy">Admin Dashboard</h1>
          <p className="font-bold text-slate-500">Upload Excel, review data status, and customize TV display.</p>
        </div>
        <button className="rounded-md bg-cj-navy px-4 py-2 font-black text-white" onClick={() => onNavigate("/tv")}>
          Open TV Mode
        </button>
      </header>

      <section className="grid grid-cols-4 gap-5">
        <button className="rounded-lg bg-white p-5 text-left shadow-panel" onClick={() => onNavigate("/upload")}>
          <Upload className="mb-4 h-7 w-7 text-cj-blue" />
          <h2 className="text-xl font-black">Upload Excel</h2>
          <p className="mt-2 text-sm font-bold text-slate-500">Upload .xlsx file for Inbound, Pick, and Outbound dashboard data.</p>
        </button>
        <button className="rounded-lg bg-white p-5 text-left shadow-panel" onClick={() => onNavigate("/settings")}>
          <Settings className="mb-4 h-7 w-7 text-cj-yellow" />
          <h2 className="text-xl font-black">Settings</h2>
          <p className="mt-2 text-sm font-bold text-slate-500">Adjust refresh interval and section visibility.</p>
        </button>
        <button className="rounded-lg bg-white p-5 text-left shadow-panel" onClick={() => onNavigate("/customize")}>
          <SlidersHorizontal className="mb-4 h-7 w-7 text-cj-red" />
          <h2 className="text-xl font-black">Customize</h2>
          <p className="mt-2 text-sm font-bold text-slate-500">Arrange TV dashboard widgets and save the layout.</p>
        </button>
        <div className="rounded-lg bg-white p-5 shadow-panel">
          <LayoutDashboard className="mb-4 h-7 w-7 text-cj-red" />
          <h2 className="text-xl font-black">Current Status</h2>
          <p className="mt-2 text-sm font-bold text-slate-500">Backend: {dashboard?.status || "waiting"}</p>
          <p className="text-sm font-bold text-slate-500">Updated: {dashboard?.updated_at || "-"}</p>
        </div>
      </section>
    </main>
  );
}
