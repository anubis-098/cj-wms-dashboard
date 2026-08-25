import type { ApexOptions } from "apexcharts";
import Chart from "react-apexcharts";

type ApexBarChartProps = {
  data?: number[];
  markerData?: Array<{
    x: string;
    y: number;
    goals: Array<{ name: string; value: number; strokeColor: string; strokeHeight: number; strokeWidth: number }>;
  }>;
  series?: Array<{ name: string; data: Array<number | null> }>;
  options: ApexOptions;
};

export default function ApexBarChart({ data = [], markerData, series, options }: ApexBarChartProps) {
  const markerRenderKey = markerData
    ? markerData.map((item) => `${item.y}:${item.goals.map((goal) => `${goal.value}:${goal.strokeColor}:${goal.strokeHeight}:${goal.strokeWidth}`).join(",")}`).join("|")
    : "standard-bar";
  const barShape = options.plotOptions?.bar;
  const shapeRenderKey = [
    barShape?.borderRadius ?? 0,
    barShape?.borderRadiusApplication ?? "end",
    barShape?.borderRadiusWhenStacked ?? "last",
    barShape?.horizontal ? "horizontal" : "vertical",
  ].join(":");
  const appearanceRenderKey = JSON.stringify({
    annotations: options.annotations,
    colors: options.colors,
    events: Object.keys(options.chart?.events ?? {}),
    legend: options.legend,
    max: options.xaxis && !Array.isArray(options.xaxis) ? options.xaxis.max : undefined,
  });

  return (
    <Chart
      key={`${markerRenderKey}:${shapeRenderKey}:${appearanceRenderKey}`}
      type="bar"
      width="100%"
      height="100%"
      options={options}
      series={series ?? [{ name: "Value", data: markerData ?? data }]}
    />
  );
}
