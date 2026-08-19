import Chart from "react-apexcharts";

import type { DashboardData, WorkGroupKey } from "../types/dashboard";

type WorkGroupChartProps = {
  data: DashboardData;
};

const groups: WorkGroupKey[] = ["inbound", "pick", "outbound"];

export function WorkGroupChart({ data }: WorkGroupChartProps) {
  const options: ApexCharts.ApexOptions = {
    chart: {
      toolbar: { show: false },
      fontFamily: "Inter, system-ui, sans-serif",
    },
    colors: ["#003b71", "#ffb000", "#d71920"],
    plotOptions: {
      bar: {
        borderRadius: 4,
        columnWidth: "45%",
      },
    },
    dataLabels: {
      enabled: true,
      formatter: (value) => `${Number(value).toFixed(0)}%`,
    },
    xaxis: {
      categories: ["Inbound", "Pick", "Outbound"],
      labels: { style: { fontWeight: 800 } },
    },
    yaxis: {
      min: 0,
      max: 100,
      labels: { formatter: (value) => `${value.toFixed(0)}%` },
    },
  };

  const series = [
    {
      name: "Progress",
      data: groups.map((group) => data.work_groups[group].progress),
    },
  ];

  return (
    <section className="rounded-lg bg-white p-5 shadow-panel">
      <div className="mb-4 flex items-center justify-between">
        <h2 className="text-lg font-black text-cj-navy">Work Progress</h2>
        <span className="text-xs font-bold uppercase text-slate-400">Inbound / Pick / Outbound</span>
      </div>
      <Chart options={options} series={series} type="bar" height={280} />
    </section>
  );
}
