import type { ApexOptions } from "apexcharts";
import Chart from "react-apexcharts";

type ApexPieChartProps = {
  type: "pie" | "donut";
  data: number[];
  options: ApexOptions;
};

export default function ApexPieChart({ type, data, options }: ApexPieChartProps) {
  return <Chart type={type} width="100%" height="100%" options={options} series={data} />;
}
