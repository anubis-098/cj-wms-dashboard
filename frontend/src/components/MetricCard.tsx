import type { WorkGroupKey, WorkGroupSummary } from "../types/dashboard";

const labels: Record<WorkGroupKey, string> = {
  inbound: "Inbound",
  pick: "Pick",
  outbound: "Outbound",
};

const accents: Record<WorkGroupKey, string> = {
  inbound: "border-cj-blue",
  pick: "border-cj-yellow",
  outbound: "border-cj-red",
};

type MetricCardProps = {
  group: WorkGroupKey;
  summary: WorkGroupSummary;
};

export function MetricCard({ group, summary }: MetricCardProps) {
  return (
    <section className={`rounded-lg border-l-4 ${accents[group]} bg-white p-5 shadow-panel`}>
      <div className="flex items-start justify-between gap-4">
        <div>
          <p className="text-sm font-black uppercase tracking-wide text-slate-500">{labels[group]}</p>
          <p className="mt-2 text-4xl font-black text-cj-navy">{summary.progress.toFixed(0)}%</p>
        </div>
        <div className="text-right">
          <p className="text-xs font-bold uppercase text-slate-400">Total</p>
          <p className="text-2xl font-black">{summary.total.toLocaleString()}</p>
        </div>
      </div>
      <div className="mt-4 h-3 overflow-hidden rounded-full bg-slate-100">
        <div className="h-full rounded-full bg-cj-blue transition-all duration-700" style={{ width: `${summary.progress}%` }} />
      </div>
      <div className="mt-4 grid grid-cols-3 gap-2 text-center text-sm font-black">
        <div className="rounded-md bg-emerald-50 px-2 py-2 text-emerald-700">{summary.completed.toLocaleString()} Done</div>
        <div className="rounded-md bg-amber-50 px-2 py-2 text-amber-700">{summary.in_progress.toLocaleString()} Doing</div>
        <div className="rounded-md bg-rose-50 px-2 py-2 text-rose-700">{summary.pending.toLocaleString()} Pending</div>
      </div>
    </section>
  );
}
