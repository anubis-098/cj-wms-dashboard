import type { ApexOptions } from "apexcharts";
import { Component, type ErrorInfo, type ReactNode } from "react";
import Chart from "react-apexcharts";

type ApexRadarChartProps = {
  options: ApexOptions;
  series: Array<{ name: string; data: number[] }>;
};

type RadarChartBoundaryProps = {
  children: ReactNode;
  resetKey: string;
};

class RadarChartBoundary extends Component<RadarChartBoundaryProps, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("Radar chart render failed", error, info);
  }

  componentDidUpdate(previousProps: RadarChartBoundaryProps) {
    if (this.state.failed && previousProps.resetKey !== this.props.resetKey) {
      this.setState({ failed: false });
    }
  }

  render() {
    if (this.state.failed) {
      return <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">Cannot display radar data</div>;
    }
    return this.props.children;
  }
}

export default function ApexRadarChart({ options, series }: ApexRadarChartProps) {
  const resetKey = JSON.stringify(series);
  return (
    <RadarChartBoundary resetKey={resetKey}>
      <Chart type="radar" width="100%" height="100%" options={options} series={series} />
    </RadarChartBoundary>
  );
}
