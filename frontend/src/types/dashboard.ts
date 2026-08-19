export type WorkGroupKey = "inbound" | "pick" | "outbound";

export type WorkGroupSummary = {
  total: number;
  completed: number;
  pending: number;
  in_progress: number;
  progress: number;
};

export type DashboardData = {
  work_groups: Record<WorkGroupKey, WorkGroupSummary>;
  sheet_names: string[];
};

export type DashboardSettings = {
  refresh_seconds: number;
  theme: "light" | "dark" | string;
  show_inbound: boolean;
  show_pick: boolean;
  show_outbound: boolean;
};

export type DashboardResponse = {
  status: string;
  data: DashboardData;
  updated_at: string;
  system_id: string;
  settings: DashboardSettings;
};

export type DashboardWidgetSize = "1x1" | "2x1" | "2x2" | "1x2";

export type DashboardWidgetItemType =
  | "metric-total"
  | "metric-inbound"
  | "metric-pick"
  | "metric-outbound"
  | "progress-chart"
  | "last-upload-status"
  | "workload-table";

export type DashboardWidgetItem = {
  id: string;
  type: DashboardWidgetItemType;
  slot: number;
  label: string;
};

export type DashboardWidget = {
  id: string;
  title: string;
  size: DashboardWidgetSize;
  items: DashboardWidgetItem[];
};

export type DashboardWidgetLayout = {
  widgets: DashboardWidget[];
};

export type WorkspaceWidget = {
  id: string;
  type: "title" | "chart" | "bar" | "text" | string;
  label: string;
  slot: number;
  width?: number;
  height?: number;
  content?: string;
  fontSize?: number;
  fontFamily?: string;
  fontWeight?: number;
  fontStyle?: "normal" | "italic";
  textColor?: string;
  backgroundColor?: string;
  useBackgroundColor?: boolean;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "center" | "bottom";
  sourceUploadId?: string;
  sourceFilename?: string;
  sheetName?: string;
  cellRange?: string;
  textQueryCell?: string;
  tableFontFamily?: string;
  tableFontSize?: number;
  tableTextAlign?: "left" | "center" | "right";
  tableVerticalAlign?: "top" | "middle" | "bottom";
  tableColumnWidths?: number[];
  tableRowHeights?: number[];
  tableCellStyles?: Record<string, { textColor?: string; backgroundColor?: string }>;
  barItems?: Array<{
    id: string;
    label: string;
    cell: string;
    markerCell?: string;
  }>;
  barMax?: number;
  barMaxInput?: string;
  barDisplayPercentage?: boolean;
  barMarkerColor?: string;
  barMarkerHeight?: number;
  barMarkerWidth?: number;
  barMarkerShowValue?: boolean;
  barMarkerFontSize?: number;
  barBorderRadius?: number;
  stackCategories?: Array<{
    id: string;
    label: string;
  }>;
  stackSeries?: Array<{
    id: string;
    label: string;
    cells: string[];
  }>;
  chartShowLegend?: boolean;
  chartLegendPosition?: "top" | "bottom";
  chartColors?: string[];
  chartFontSize?: number;
  pieShowValueCallouts?: boolean;
  iconName?: string;
  iconColor?: string;
  iconSize?: number;
  gradientStartColor?: string;
  gradientEndColor?: string;
  gradientDirection?: string;
  gradientStartPosition?: number;
  gradientEndPosition?: number;
  gradientOpacity?: number;
  gradientBorderRadius?: number;
  gradientStops?: Array<{
    id: string;
    color: string;
    position: number;
  }>;
  lineCurve?: "smooth" | "straight" | "stepline";
  lineStrokeWidth?: number;
  lineShowMarkers?: boolean;
  lineNullMissing?: boolean;
  lineAnnotations?: Array<{
    id: string;
    label: string;
    axis: "x" | "y";
    value: string;
    color: string;
  }>;
  columnLabelRotation?: number;
  columnWidth?: number;
  columnBorderRadius?: number;
  columnShowDataLabels?: boolean;
  columnShowDataLabelBackground?: boolean;
};

export type WorkspaceBox = {
  id: string;
  size: string;
  title: string;
  cell: number;
  columns: number;
  rows: number;
  widgets: WorkspaceWidget[];
};

export type WorkspacePage = {
  id: string;
  name: string;
  boxes: WorkspaceBox[];
  isMain?: boolean;
  locked?: boolean;
};

export type WorkspaceLayout = {
  boxes: WorkspaceBox[];
  pages?: WorkspacePage[];
  activePageId?: string;
};
