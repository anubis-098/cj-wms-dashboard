import { useState } from "react";

import { saveDashboardSettings } from "../services/api";
import type { DashboardSettings } from "../types/dashboard";

type SettingsPageProps = {
  settings: DashboardSettings;
  onNavigate: (path: string) => void;
};

export function SettingsPage({ settings, onNavigate }: SettingsPageProps) {
  const [form, setForm] = useState(settings);
  const [status, setStatus] = useState("Ready");

  async function handleSave() {
    setStatus("Saving...");
    try {
      await saveDashboardSettings(form);
      setStatus("Saved");
    } catch {
      setStatus("Save failed");
    }
  }

  return (
    <main className="min-h-screen bg-screen-bg p-6">
      <button className="mb-5 rounded-md bg-white px-4 py-2 font-black shadow-panel" onClick={() => onNavigate("/admin")}>
        Back
      </button>
      <section className="max-w-2xl rounded-lg bg-white p-6 shadow-panel">
        <h1 className="text-3xl font-black text-cj-navy">Dashboard Settings</h1>
        <label className="mt-6 block">
          <span className="text-sm font-black uppercase text-slate-500">Refresh seconds</span>
          <input
            className="mt-1 w-full rounded-md border border-slate-200 px-3 py-2"
            min={15}
            type="number"
            value={form.refresh_seconds}
            onChange={(event) => setForm({ ...form, refresh_seconds: Number(event.target.value) })}
          />
        </label>
        <div className="mt-5 grid grid-cols-3 gap-3">
          {(["show_inbound", "show_pick", "show_outbound"] as const).map((key) => (
            <label key={key} className="rounded-md border border-slate-200 p-3 font-bold">
              <input className="mr-2" checked={Boolean(form[key])} type="checkbox" onChange={(event) => setForm({ ...form, [key]: event.target.checked })} />
              {key.replace("show_", "").toUpperCase()}
            </label>
          ))}
        </div>
        <button className="mt-6 rounded-md bg-cj-navy px-4 py-3 font-black text-white" onClick={handleSave}>
          Save Settings
        </button>
        <p className="mt-4 font-black text-slate-600">{status}</p>
      </section>
    </main>
  );
}
