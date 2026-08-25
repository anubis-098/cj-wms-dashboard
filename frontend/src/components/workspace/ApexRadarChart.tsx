import type { ApexOptions } from "apexcharts";
import Chart from "react-apexcharts";

type ApexRadarChartProps = {
  options: ApexOptions;
  series: Array<{ name: string; data: Array<number | null> }>;
};

export default function ApexRadarChart({ options, series }: ApexRadarChartProps) {
  return <Chart type="radar" width="100%" height="100%" options={options} series={series} />;
}
