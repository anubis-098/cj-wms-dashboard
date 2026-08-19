import type { ApexOptions } from "apexcharts";
import Chart from "react-apexcharts";

type ApexLineChartProps = {
  options: ApexOptions;
  series: Array<{ name: string; data: Array<number | null> }>;
};

export default function ApexLineChart({ options, series }: ApexLineChartProps) {
  return <Chart type="line" width="100%" height="100%" options={options} series={series} />;
}
