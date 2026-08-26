import {
  CollisionDetection,
  DndContext,
  DragEndEvent,
  DragOverlay,
  DragStartEvent,
  PointerSensor,
  DragOverEvent,
  useDraggable,
  useDroppable,
  useSensor,
  useSensors,
  pointerWithin,
  rectIntersection,
  closestCenter,
} from "@dnd-kit/core";
import { arrayMove, SortableContext, useSortable, verticalListSortingStrategy } from "@dnd-kit/sortable";
import { CSS as DndCSS } from "@dnd-kit/utilities";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownToLine,
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  ArrowUpToLine,
  AlertTriangle,
  AlarmClock,
  BarChart3,
  Box,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  ClipboardList,
  Clock3,
  CloudMoon,
  CloudSun,
  Copy,
  Database,
  FileSpreadsheet,
  GripVertical,
  Heading,
  Italic,
  Lock,
  LockOpen,
  Link,
  MapPin,
  Menu,
  Moon,
  Plus,
  Package,
  Palette,
  Pencil,
  RefreshCw,
  Save,
  Sunrise,
  Sunset,
  Sun,
  Table2,
  Triangle,
  Timer,
  Trash2,
  Truck,
  Type,
  Users,
  Warehouse,
  X,
} from "lucide-react";
import type { LucideIcon } from "lucide-react";
import { CSSProperties, KeyboardEvent as ReactKeyboardEvent, lazy, PointerEvent, ReactNode, Suspense, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type ApexCharts from "apexcharts";
import type { ApexOptions } from "apexcharts";

import { fetchExcelCells, fetchExcelRange, fetchExcelSheets, fetchExcelUploads, fetchWorkspaceLayout, saveWorkspaceLayout } from "../../services/api";
import type { ExcelUploadRecord } from "../../services/api";
import type { WorkspaceBox, WorkspaceLayout, WorkspacePage, WorkspaceWidget } from "../../types/dashboard";

const ApexBarChart = lazy(() => import("./ApexBarChart"));
const ApexPieChart = lazy(() => import("./ApexPieChart"));
const ApexLineChart = lazy(() => import("./ApexLineChart"));
const ApexRadarChart = lazy(() => import("./ApexRadarChart"));

type BoxSize = "1x1" | "2x1" | "2x2" | "1x2";
type WidgetType = "title" | "chart" | "bar" | "bar-markers" | "stack-bar" | "stack-column" | "stack-100-bar" | "stack-100-column" | "column-rotated-labels" | "basic-line" | "line-annotations" | "radar-polygon" | "simple-pie" | "simple-donut" | "icon" | "gradient-color" | "text" | "text-query" | "excel-table";

type BoxTemplate = {
  size: BoxSize;
  label: string;
  columns: number;
  rows: number;
};

type WidgetTemplate = {
  type: WidgetType;
  label: string;
  icon: ReactNode;
};

type DragData =
  | { kind: "box-template"; template: BoxTemplate }
  | { kind: "box"; box: WorkspaceBox }
  | { kind: "widget-template"; template: WidgetTemplate }
  | { kind: "widget-copy"; sourceBoxId: string; widget: WorkspaceWidget }
  | { kind: "widget"; boxId: string; widget: WorkspaceWidget };

function SortableChartRow({ id, className, children }: { id: string; className: string; children: ReactNode }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id });

  return (
    <div
      ref={setNodeRef}
      className={`${className} ${isDragging ? "relative z-10 opacity-60 shadow-panel" : ""}`}
      style={{ transform: DndCSS.Transform.toString(transform), transition }}
    >
      <button
        {...attributes}
        {...listeners}
        aria-label="Reorder series"
        className="grid h-8 w-7 touch-none cursor-grab place-items-center rounded text-slate-400 hover:bg-white hover:text-cj-blue active:cursor-grabbing"
        title="Drag to reorder"
        type="button"
      >
        <GripVertical className="h-4 w-4" />
      </button>
      {children}
    </div>
  );
}

function SortablePageRow({ id, active, children }: { id: string; active: boolean; children: ReactNode }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id });

  return (
    <div
      ref={setNodeRef}
      className={`mb-1 flex items-center gap-1 rounded p-1 transition ${active ? "bg-blue-50" : "hover:bg-slate-50"} ${isDragging ? "relative z-20 bg-white opacity-75 shadow-panel" : ""}`}
      style={{ transform: DndCSS.Transform.toString(transform), transition }}
    >
      <button
        {...attributes}
        {...listeners}
        aria-label="Drag to reorder page"
        className="grid h-7 w-6 shrink-0 touch-none cursor-grab place-items-center rounded text-slate-400 hover:bg-white hover:text-cj-blue active:cursor-grabbing"
        title="Drag to reorder page"
        type="button"
      >
        <GripVertical className="h-4 w-4" />
      </button>
      {children}
    </div>
  );
}

function PageCopyDropTarget({ page, available }: { page: WorkspacePage; available: boolean }) {
  const { isOver, setNodeRef } = useDroppable({
    id: `page-copy-target:${page.id}`,
    data: { kind: "page-copy-target", pageId: page.id },
    disabled: !available,
  });

  return (
    <div
      ref={setNodeRef}
      className={`flex min-h-20 items-center gap-3 rounded-md border-2 px-4 py-3 transition-colors ${
        !available
          ? "border-slate-200 bg-slate-50 opacity-45"
          : isOver
            ? "border-cj-blue bg-blue-100 text-cj-blue shadow-panel ring-2 ring-cj-blue/25"
            : "border-slate-200 bg-white text-slate-600 hover:border-cj-blue/50 hover:bg-blue-50"
      }`}
    >
      <span className="grid h-8 w-8 shrink-0 place-items-center rounded bg-slate-100 text-cj-blue"><Copy className="h-4 w-4" /></span>
      <span className="min-w-0">
        <strong className="block truncate text-xs font-black">{page.name}</strong>
        <small className="block text-[9px] font-bold text-slate-400">{available ? "Drop to copy Widget" : "No available Box space"}</small>
      </span>
    </div>
  );
}

function PageCopyPanel({ children }: { children: ReactNode }) {
  const { setNodeRef } = useDroppable({
    id: "page-copy-panel",
    data: { kind: "page-copy-panel" },
  });

  return (
    <aside
      ref={setNodeRef}
      className="fixed bottom-4 right-3 top-24 z-[110] flex w-80 flex-col rounded-md border border-slate-200 bg-white p-3 shadow-panel"
    >
      {children}
    </aside>
  );
}

const workspaceCollisionDetection: CollisionDetection = (args) => {
  const activeKind = (args.active.data.current as DragData | undefined)?.kind;
  const targetKinds = activeKind === "widget-copy"
    ? ["box-slot", "page-copy-target", "page-copy-panel"]
    : activeKind === "widget" || activeKind === "widget-template"
      ? ["box-slot"]
      : ["workspace-cell"];
  const droppableContainers = args.droppableContainers.filter(
    (container) => targetKinds.includes(container.data.current?.kind),
  );
  const filteredArgs = { ...args, droppableContainers };
  const pointerCollisions = pointerWithin(filteredArgs);

  if (activeKind === "widget-copy" && pointerCollisions.length > 0) {
    const pageTarget = pointerCollisions.find((collision) => collision.data?.droppableContainer.data.current?.kind === "page-copy-target");
    if (pageTarget) return [pageTarget];

    const pagePanel = pointerCollisions.find((collision) => collision.data?.droppableContainer.data.current?.kind === "page-copy-panel");
    if (pagePanel) return [pagePanel];
  }

  if (pointerCollisions.length > 0) return pointerCollisions;

  const rectangleCollisions = rectIntersection(filteredArgs);
  if (activeKind === "widget-copy") {
    const pageTarget = rectangleCollisions.find((collision) => collision.data?.droppableContainer.data.current?.kind === "page-copy-target");
    if (pageTarget) return [pageTarget];

    const pagePanel = rectangleCollisions.find((collision) => collision.data?.droppableContainer.data.current?.kind === "page-copy-panel");
    if (pagePanel) return [pagePanel];
  }

  return rectangleCollisions;
};

const workspaceColumns = 12;
const workspaceRows = 6;
const workspaceCellCount = workspaceColumns * workspaceRows;
const slotsPerCellColumn = 4;
const slotsPerCellRow = 4;
const defaultChartColors = ["#0080c6", "#e42f44", "#ec8922", "#16a085", "#7c3aed", "#64748b", "#84cc16", "#0891b2"];
const isTvRuntime = /SMART-TV|Tizen|SamsungBrowser/i.test(window.navigator.userAgent);
const chartAnimations = {
  animateGradually: { delay: isTvRuntime ? 0 : 55, enabled: !isTvRuntime },
  dynamicAnimation: { enabled: true, speed: isTvRuntime ? 400 : 1100 },
  easing: "easeinout" as const,
  enabled: true,
  speed: isTvRuntime ? 550 : 1400,
};

function FontSizeInput({
  allowAuto = false,
  disabled = false,
  max,
  min,
  onCommit,
  value,
}: {
  allowAuto?: boolean;
  disabled?: boolean;
  max: number;
  min: number;
  onCommit: (value: number | undefined) => void;
  value: number | undefined;
}) {
  const [draft, setDraft] = useState(value === undefined ? "" : String(value));

  useEffect(() => setDraft(value === undefined ? "" : String(value)), [value]);

  function commit() {
    if (draft.trim() === "" && allowAuto) {
      onCommit(undefined);
      return;
    }
    const parsed = Number(draft);
    if (!Number.isFinite(parsed)) {
      setDraft(value === undefined ? "" : String(value));
      return;
    }
    const nextValue = Math.min(max, Math.max(min, parsed));
    setDraft(String(nextValue));
    onCommit(nextValue);
  }

  return (
    <input
      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-100"
      disabled={disabled}
      inputMode="numeric"
      max={max}
      min={min}
      placeholder={allowAuto ? "Auto" : undefined}
      step="1"
      type="number"
      value={draft}
      onBlur={commit}
      onChange={(event) => setDraft(event.target.value)}
      onKeyDown={(event) => {
        if (event.key === "Enter") event.currentTarget.blur();
      }}
    />
  );
}
const iconSizeOptions = [16, 20, 24, 28, 32, 40, 48, 56, 64, 72];
const iconOptions: Array<{ name: string; label: string; icon: LucideIcon }> = [
  { name: "arrow-up", label: "Arrow up", icon: ArrowUp },
  { name: "arrow-down", label: "Arrow down", icon: ArrowDown },
  { name: "arrow-left", label: "Arrow left", icon: ArrowLeft },
  { name: "arrow-right", label: "Arrow right", icon: ArrowRight },
  { name: "link", label: "Link", icon: Link },
  { name: "package", label: "Package", icon: Package },
  { name: "box", label: "Box", icon: Box },
  { name: "warehouse", label: "Warehouse", icon: Warehouse },
  { name: "truck", label: "Truck", icon: Truck },
  { name: "clipboard", label: "Clipboard", icon: ClipboardList },
  { name: "sun", label: "Day / Sun", icon: Sun },
  { name: "moon", label: "Night / Moon", icon: Moon },
  { name: "sunrise", label: "Sunrise", icon: Sunrise },
  { name: "sunset", label: "Sunset", icon: Sunset },
  { name: "cloud-sun", label: "Day shift", icon: CloudSun },
  { name: "cloud-moon", label: "Night shift", icon: CloudMoon },
  { name: "clock", label: "Clock", icon: Clock3 },
  { name: "alarm-clock", label: "Alarm clock", icon: AlarmClock },
  { name: "timer", label: "Timer", icon: Timer },
  { name: "complete", label: "Complete", icon: CheckCircle2 },
  { name: "warning", label: "Warning", icon: AlertTriangle },
  { name: "users", label: "Users", icon: Users },
  { name: "database", label: "Database", icon: Database },
  { name: "spreadsheet", label: "Spreadsheet", icon: FileSpreadsheet },
  { name: "refresh", label: "Refresh", icon: RefreshCw },
  { name: "location", label: "Location", icon: MapPin },
];
const fontFamilyOptions = ["Inter", "Arial", "Tahoma", "Verdana", "Georgia", "Times New Roman"];
const fontWeightOptions = [
  { label: "Thin", value: 300 },
  { label: "Normal", value: 400 },
  { label: "Bold", value: 700 },
  { label: "Extra Bold", value: 900 },
];

function isLargeDataWidget(type: string) {
  return type === "excel-table" || type === "bar" || type === "bar-markers" || type === "stack-bar" || type === "stack-column" || type === "stack-100-bar" || type === "stack-100-column" || type === "column-rotated-labels" || type === "basic-line" || type === "line-annotations" || type === "radar-polygon" || type === "simple-pie" || type === "simple-donut";
}

function excelColumnToNumber(column: string) {
  return column.toUpperCase().split("").reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0);
}

function excelNumberToColumn(value: number) {
  let column = "";
  let current = value;
  while (current > 0) {
    current -= 1;
    column = String.fromCharCode(65 + (current % 26)) + column;
    current = Math.floor(current / 26);
  }
  return column;
}

function getExcelCellReference(cellRange: string | undefined, rowOffset: number, columnOffset: number) {
  const match = cellRange?.trim().toUpperCase().match(/^\$?([A-Z]+)\$?(\d+)/);
  if (!match) return "";
  return `${excelNumberToColumn(excelColumnToNumber(match[1]) + columnOffset)}${Number(match[2]) + rowOffset}`;
}

function buildRunningCellReferences(value: string, count: number) {
  const match = value.trim().toUpperCase().match(/^\$?([A-Z]+)\$?(\d+)$/);
  if (!match) return null;
  const startRow = Number(match[2]);
  if (!Number.isSafeInteger(startRow) || startRow < 1) return null;
  return Array.from({ length: count }, (_, index) => `${match[1]}${startRow + index}`);
}

function parseNumericLiteral(input: string | undefined) {
  const normalized = (input ?? "").trim().replace(/,/g, "");
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;
  const value = Number(normalized);
  return Number.isFinite(value) ? value : null;
}

function formatExcelDateValue(value: unknown, numberFormat: string | undefined) {
  const match = String(value ?? "").match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ].*)?$/);
  if (!match || !numberFormat) return null;
  const [, year, month, day] = match;
  let format = numberFormat.split(";")[0]
    .replace(/\[\$-[^\]]+]/g, "")
    .replace(/\\(.)/g, "$1")
    .replace(/"([^"]*)"/g, "$1");
  if (!/[dy]/i.test(format)) return null;

  const dateTokens = [...format.matchAll(/m{1,4}|d{1,2}|y{2,4}/gi)];
  const monthToken = dateTokens.find((token) => /^m{1,2}$/i.test(token[0]));
  const dayToken = dateTokens.find((token) => /^d{1,2}$/i.test(token[0]));
  if (monthToken?.index !== undefined && dayToken?.index !== undefined && monthToken.index < dayToken.index) {
    format = `${format.slice(0, monthToken.index)}${dayToken[0]}${format.slice(monthToken.index + monthToken[0].length, dayToken.index)}${monthToken[0]}${format.slice(dayToken.index + dayToken[0].length)}`;
  }

  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  const replacements: Record<string, string> = {
    d: String(Number(day)),
    dd: day,
    m: String(Number(month)),
    mm: month,
    mmm: date.toLocaleDateString("en-GB", { month: "short", timeZone: "UTC" }),
    mmmm: date.toLocaleDateString("en-GB", { month: "long", timeZone: "UTC" }),
    yy: year.slice(-2),
    yyyy: year,
  };
  return format.replace(/yyyy|mmmm|mmm|yy|mm|dd|m|d/gi, (token) => replacements[token.toLowerCase()] ?? token);
}

function formatTextQueryValue(value: unknown, numberFormat?: string) {
  if (typeof value === "number" && Number.isFinite(value)) {
    const integer = Math.trunc(value);
    const formatted = Math.abs(integer).toLocaleString("en-US");
    return integer < 0 ? `(${formatted})` : formatted;
  }
  const excelDate = formatExcelDateValue(value, numberFormat);
  if (excelDate !== null) return excelDate;
  const text = String(value ?? "");
  const isoDate = text.match(/^(\d{4}-\d{2}-\d{2})(?:[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?$/);
  return isoDate?.[1] ?? text;
}

function formatExcelTableValue(value: unknown, numberFormat?: string) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return formatExcelDateValue(value, numberFormat) ?? String(value ?? "");
  }
  const formatSections = (numberFormat ?? "").split(";");
  const numericFormat = formatSections[value < 0 ? 1 : 0] ?? numberFormat ?? "";
  const decimalPattern = numericFormat.match(/\.([0#?]+)/);
  const minimumFractionDigits = Math.min(1, decimalPattern?.[1].replace(/[^0]/g, "").length ?? 0);
  const maximumFractionDigits = Math.min(1, decimalPattern?.[1].length ?? (String(value).includes(".") ? 1 : 0));
  const useGrouping = numericFormat.includes(",");
  const formatted = Math.abs(value).toLocaleString("en-US", {
    maximumFractionDigits,
    minimumFractionDigits,
    useGrouping,
  });
  if (value >= 0) return formatted;
  const usesParentheses = numericFormat.replace(/\\([()])/g, "$1").includes("(") && numericFormat.replace(/\\([()])/g, "$1").includes(")");
  return usesParentheses ? `(${formatted})` : `-${formatted}`;
}

function formatChartTooltipValue(value: number) {
  return Number(value).toLocaleString("en-US", { maximumFractionDigits: 2 });
}

function excelZeroFormatDisplaysDash(numberFormat: string | undefined) {
  if (!numberFormat) return false;
  const sections = numberFormat.split(";");
  if (sections.length < 3) return false;
  const zeroSection = sections[2]
    .replace(/\[[^\]]*]/g, "")
    .replace(/_.|\*./g, "")
    .replace(/\\(.)/g, "$1")
    .replace(/"([^"]*)"/g, "$1")
    .replace(/\s+/g, "");
  // In Excel accounting formats, ? reserves character width; it is not a numeric zero value.
  return zeroSection.includes("-") && !/[0#]/.test(zeroSection);
}

function isExcelInput(input: string | undefined) {
  return /\$?[A-Z]+\$?\d+/i.test(input ?? "");
}

type ExcelCellLookup = Record<string, { value: unknown; number_format: string }>;

async function fetchNumericInputLookup(uploadId: string | undefined, sheetName: string | undefined, inputs: Array<string | undefined>) {
  const references = [...new Set(inputs.flatMap((input) => input?.match(/\$?[A-Z]+\$?\d+/gi) ?? []).map((cell) => cell.replace(/\$/g, "").toUpperCase()))];
  if (references.length === 0) return {} as ExcelCellLookup;
  if (!uploadId || !sheetName) throw new Error("Excel source required");
  try {
    return (await fetchExcelCells(uploadId, sheetName, references)).data;
  } catch (batchError) {
    console.warn("Excel batch Cell lookup failed; falling back to Range requests", batchError);
    const entries = await Promise.all(references.map(async (reference) => {
      const response = await fetchExcelRange(uploadId, sheetName, reference);
      return [reference, {
        value: response.data[0]?.[0],
        number_format: response.number_formats?.[0]?.[0] ?? "General",
      }] as const;
    }));
    return Object.fromEntries(entries) as ExcelCellLookup;
  }
}

function tokenizeNumericExpression(input: string) {
  const expression = input
    .trim()
    .replace(/^=/, "")
    .replace(/(\d|\))\s*[xX×]\s*(?=[$A-Z(\d.])/g, "$1*")
    .replace(/\s+/g, "")
    .toUpperCase();
  const tokens = expression.match(/\$?[A-Z]+\$?\d+|(?:\d+\.?\d*|\.\d+)|[()+\-*/]/g);
  if (!tokens || tokens.join("") !== expression) throw new Error("Invalid formula");
  return tokens;
}

async function evaluateNumericExpression(uploadId: string | undefined, sheetName: string | undefined, input: string, cellLookup?: ExcelCellLookup) {
  const tokens = tokenizeNumericExpression(input);
  let position = 0;

  async function parsePrimary(): Promise<number> {
    const token = tokens[position];
    if (token === "(") {
      position += 1;
      const value = await parseAddition();
      if (tokens[position] !== ")") throw new Error("Missing closing parenthesis");
      position += 1;
      return value;
    }
    if (token === "+" || token === "-") {
      position += 1;
      const value = await parsePrimary();
      return token === "-" ? -value : value;
    }
    if (token && /^\$?[A-Z]+\$?\d+$/.test(token)) {
      if (!uploadId || !sheetName) throw new Error("Excel source required");
      position += 1;
      const normalizedCell = token.replace(/\$/g, "");
      const hasCachedCell = cellLookup && Object.prototype.hasOwnProperty.call(cellLookup, normalizedCell);
      const rawValue = hasCachedCell ? cellLookup[normalizedCell].value : (await fetchExcelRange(uploadId, sheetName, normalizedCell)).data[0]?.[0];
      const value = Number(String(rawValue ?? 0).replace(/,/g, ""));
      if (!Number.isFinite(value)) throw new Error(`Cell ${token} is not numeric`);
      return value;
    }
    if (token && /^(?:\d+\.?\d*|\.\d+)$/.test(token)) {
      position += 1;
      return Number(token);
    }
    throw new Error("Expected a number, Cell, or parenthesis");
  }

  async function parseMultiplication(): Promise<number> {
    let value = await parsePrimary();
    while (tokens[position] === "*" || tokens[position] === "/") {
      const operator = tokens[position];
      position += 1;
      const right = await parsePrimary();
      if (operator === "/" && right === 0) throw new Error("Division by zero");
      value = operator === "*" ? value * right : value / right;
    }
    return value;
  }

  async function parseAddition(): Promise<number> {
    let value = await parseMultiplication();
    while (tokens[position] === "+" || tokens[position] === "-") {
      const operator = tokens[position];
      position += 1;
      const right = await parseMultiplication();
      value = operator === "+" ? value + right : value - right;
    }
    return value;
  }

  const result = await parseAddition();
  if (position !== tokens.length || !Number.isFinite(result)) throw new Error("Invalid formula result");
  return result;
}

async function resolveNumericInput(uploadId: string | undefined, sheetName: string | undefined, input: string | undefined, cellLookup?: ExcelCellLookup) {
  const literal = parseNumericLiteral(input);
  if (literal !== null) return literal;
  if (!input?.trim()) throw new Error("Value required");
  return evaluateNumericExpression(uploadId, sheetName, input, cellLookup);
}

async function resolveNullableNumericInput(uploadId: string | undefined, sheetName: string | undefined, input: string | undefined, preserveMissing: boolean, cellLookup?: ExcelCellLookup) {
  const normalized = input?.trim().toUpperCase() ?? "";
  if (preserveMissing && /^\$?[A-Z]+\$?\d+$/.test(normalized)) {
    if (!uploadId || !sheetName) throw new Error("Excel source required");
    const normalizedCell = normalized.replace(/\$/g, "");
    const cachedCell = cellLookup?.[normalizedCell];
    const response = cachedCell ? null : await fetchExcelRange(uploadId, sheetName, normalizedCell);
    const rawValue = cachedCell?.value ?? response?.data[0]?.[0];
    const numberFormat = cachedCell?.number_format ?? response?.number_formats?.[0]?.[0];
    const normalizedValue = String(rawValue ?? "").trim();
    if (rawValue === null || rawValue === undefined || normalizedValue === "" || /^[-\u2013\u2014]$/.test(normalizedValue)) return null;
    if (Number(rawValue) === 0 && excelZeroFormatDisplaysDash(numberFormat)) return null;
  }
  return resolveNumericInput(uploadId, sheetName, input, cellLookup);
}

function splitFormulaArguments(input: string) {
  const argumentsList: string[] = [];
  let depth = 0;
  let start = 0;
  for (let index = 0; index < input.length; index += 1) {
    if (input[index] === "(") depth += 1;
    if (input[index] === ")") depth -= 1;
    if (input[index] === "," && depth === 0) {
      argumentsList.push(input.slice(start, index));
      start = index + 1;
    }
  }
  argumentsList.push(input.slice(start));
  return argumentsList.map((item) => item.trim()).filter(Boolean);
}

async function resolveMaximumInput(uploadId: string | undefined, sheetName: string | undefined, input: string) {
  const normalized = input.trim();
  const maxMatch = normalized.match(/^MAX\((.*)\)$/i);
  if (!maxMatch) return resolveNumericInput(uploadId, sheetName, normalized);

  const values: number[] = [];
  for (const argument of splitFormulaArguments(maxMatch[1])) {
    const rangeMatch = argument.toUpperCase().match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
    if (rangeMatch) {
      if (!uploadId || !sheetName) throw new Error("Excel source required");
      const cellRange = `${rangeMatch[1]}${rangeMatch[2]}:${rangeMatch[3]}${rangeMatch[4]}`;
      const response = await fetchExcelRange(uploadId, sheetName, cellRange);
      response.data.flat().forEach((cell) => {
        const value = Number(String(cell ?? "").replace(/,/g, ""));
        if (Number.isFinite(value)) values.push(value);
      });
    } else {
      values.push(await resolveNumericInput(uploadId, sheetName, argument));
    }
  }
  if (values.length === 0) throw new Error("MAX requires at least one numeric value");
  return Math.max(...values);
}

function roundMaximumUp(value: number) {
  if (!Number.isFinite(value) || value <= 0) return value;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  return Math.ceil(value / magnitude) * magnitude;
}

function roundColumnDataLabelBackgrounds(chartContext: ApexCharts, fixedSize?: number) {
  window.requestAnimationFrame(() => {
    const chartRoot = (chartContext as unknown as { el?: Element }).el;
    chartRoot?.querySelectorAll<SVGRectElement>(".apexcharts-data-labels rect").forEach((rect) => {
      const width = Number(rect.getAttribute("width")) || 0;
      const height = Number(rect.getAttribute("height")) || 0;
      const x = Number(rect.getAttribute("x")) || 0;
      const y = Number(rect.getAttribute("y")) || 0;
      const size = fixedSize ?? Math.max(width, height);
      if (size <= 0) return;
      rect.setAttribute("x", String(x - (size - width) / 2));
      rect.setAttribute("y", String(y - (size - height) / 2));
      rect.setAttribute("width", String(size));
      rect.setAttribute("height", String(size));
      rect.setAttribute("rx", String(size / 2));
      rect.setAttribute("ry", String(size / 2));
    });
  });
}

function drawPieValueCallouts(chartContext: ApexCharts) {
  window.requestAnimationFrame(() => {
    const chartRoot = (chartContext as unknown as { el?: Element }).el;
    const svg = chartRoot?.querySelector<SVGSVGElement>("svg.apexcharts-svg");
    const pie = svg?.querySelector<SVGGraphicsElement>(".apexcharts-pie-series");
    if (!svg || !pie) return;

    svg.querySelector("g[data-pie-value-callouts]")?.remove();
    const labels = Array.from(svg.querySelectorAll<SVGTextElement>(".apexcharts-data-labels text"));
    if (labels.length === 0) return;

    const svgMatrix = svg.getScreenCTM();
    if (!svgMatrix) return;
    const inverseMatrix = svgMatrix.inverse();
    const pieRect = pie.getBoundingClientRect();
    const centerX = pieRect.left + pieRect.width / 2;
    const centerY = pieRect.top + pieRect.height / 2;
    const radius = Math.min(pieRect.width, pieRect.height) / 2;
    const group = document.createElementNS("http://www.w3.org/2000/svg", "g");
    group.setAttribute("data-pie-value-callouts", "true");
    group.setAttribute("fill", "none");
    group.setAttribute("stroke", "#64748b");
    group.setAttribute("stroke-width", "1.25");

    const toSvgPoint = (x: number, y: number) => {
      const point = svg.createSVGPoint();
      point.x = x;
      point.y = y;
      return point.matrixTransform(inverseMatrix);
    };

    labels.forEach((label) => {
      const labelRect = label.getBoundingClientRect();
      const labelX = labelRect.left + labelRect.width / 2;
      const labelY = labelRect.top + labelRect.height / 2;
      const deltaX = labelX - centerX;
      const deltaY = labelY - centerY;
      const distance = Math.hypot(deltaX, deltaY) || 1;
      const unitX = deltaX / distance;
      const unitY = deltaY / distance;
      const start = toSvgPoint(centerX + unitX * radius * 0.78, centerY + unitY * radius * 0.78);
      const elbow = toSvgPoint(centerX + unitX * radius * 1.02, centerY + unitY * radius * 1.02);
      const end = toSvgPoint(labelX - unitX * (labelRect.width / 2 + 3), labelY);
      const line = document.createElementNS("http://www.w3.org/2000/svg", "polyline");
      line.setAttribute("points", `${start.x},${start.y} ${elbow.x},${elbow.y} ${end.x},${end.y}`);
      group.appendChild(line);
    });

    pie.parentElement?.insertBefore(group, pie.nextSibling);
  });
}

let stackClipSequence = 0;
const stackClipNamespaces = new WeakMap<Element, string>();
const stackRoundTimers = new WeakMap<ApexCharts, number>();

function getStackClipNamespace(chartRoot: Element) {
  const existing = stackClipNamespaces.get(chartRoot);
  if (existing) return existing;
  stackClipSequence += 1;
  const namespace = `stack-chart-${stackClipSequence}`;
  stackClipNamespaces.set(chartRoot, namespace);
  return namespace;
}

function clearStackBarSegmentClips(chartContext: ApexCharts) {
  const chartRoot = (chartContext as unknown as { el?: Element }).el;
  const svg = chartRoot?.querySelector<SVGSVGElement>("svg.apexcharts-svg");
  if (!svg) return;
  svg.querySelector("defs[data-stack-rounded]")?.remove();
  svg.querySelectorAll(".apexcharts-bar-series .apexcharts-bar-area").forEach((segment) => {
    segment.removeAttribute("clip-path");
  });
}

function scheduleStackBarRounding(chartContext: ApexCharts, radius: number) {
  clearStackBarSegmentClips(chartContext);
  const activeTimer = stackRoundTimers.get(chartContext);
  if (activeTimer !== undefined) window.clearTimeout(activeTimer);
  const nextTimer = window.setTimeout(() => {
    stackRoundTimers.delete(chartContext);
    roundStackBarSegments(chartContext, radius);
  }, 2200);
  stackRoundTimers.set(chartContext, nextTimer);
}

function roundStackBarSegments(chartContext: ApexCharts, radius: number) {
  window.requestAnimationFrame(() => {
    const chartRoot = (chartContext as unknown as { el?: Element }).el;
    const svg = chartRoot?.querySelector<SVGSVGElement>("svg.apexcharts-svg");
    if (!chartRoot || !svg) return;

    svg.querySelector("defs[data-stack-rounded]")?.remove();
    const definitions = document.createElementNS("http://www.w3.org/2000/svg", "defs");
    definitions.setAttribute("data-stack-rounded", "true");
    const clipNamespace = getStackClipNamespace(chartRoot);

    svg.querySelectorAll<SVGGraphicsElement>(".apexcharts-bar-series .apexcharts-bar-area").forEach((segment, index) => {
      const bounds = segment.getBBox();
      if (bounds.width <= 0 || bounds.height <= 0) return;
      const appliedRadius = Math.min(radius, bounds.width / 2, bounds.height / 2);
      if (appliedRadius <= 0) {
        segment.removeAttribute("clip-path");
        return;
      }

      const clipId = `${clipNamespace}-segment-${index}`;
      const clipPath = document.createElementNS("http://www.w3.org/2000/svg", "clipPath");
      const clipRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      clipPath.setAttribute("id", clipId);
      clipRect.setAttribute("x", String(bounds.x));
      clipRect.setAttribute("y", String(bounds.y));
      clipRect.setAttribute("width", String(bounds.width));
      clipRect.setAttribute("height", String(bounds.height));
      clipRect.setAttribute("rx", String(appliedRadius));
      clipRect.setAttribute("ry", String(appliedRadius));
      clipPath.appendChild(clipRect);
      definitions.appendChild(clipPath);
      segment.setAttribute("clip-path", `url(#${clipId})`);
    });

    svg.prepend(definitions);
  });
}

const boxTemplates: BoxTemplate[] = [
  { size: "1x1", label: "Box 1x1", columns: 1, rows: 1 },
];

const widgetTemplates: WidgetTemplate[] = [
  { type: "title", label: "Title", icon: <Heading className="h-4 w-4" /> },
  { type: "chart", label: "Chart", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "bar", label: "Bar", icon: <GripVertical className="h-4 w-4 rotate-90" /> },
  { type: "bar-markers", label: "Bar with Markers", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "stack-bar", label: "Stack Bar", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "stack-column", label: "Stack Column", icon: <BarChart3 className="h-4 w-4 rotate-90" /> },
  { type: "stack-100-bar", label: "Stack 100% Bar", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "stack-100-column", label: "Stacked Column 100%", icon: <BarChart3 className="h-4 w-4 rotate-90" /> },
  { type: "column-rotated-labels", label: "Column with Rotated Labels", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "basic-line", label: "Basic Line Chart", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "line-annotations", label: "Line with Annotations", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "radar-polygon", label: "Radar with Polygon Fill", icon: <Triangle className="h-4 w-4" /> },
  { type: "simple-pie", label: "Simple Pie", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "simple-donut", label: "Simple Donut", icon: <BarChart3 className="h-4 w-4" /> },
  { type: "icon", label: "Icon", icon: <Package className="h-4 w-4" /> },
  { type: "gradient-color", label: "Gradient Color", icon: <Palette className="h-4 w-4" /> },
  { type: "text", label: "Text", icon: <Type className="h-4 w-4" /> },
  { type: "text-query", label: "Text Query", icon: <FileSpreadsheet className="h-4 w-4" /> },
  { type: "excel-table", label: "Excel Query", icon: <Table2 className="h-4 w-4" /> },
];

const initialBoxes: WorkspaceBox[] = [
  {
    id: "box-1",
    size: "1x1",
    title: "Box 1x1",
    cell: 0,
    columns: 1,
    rows: 1,
    widgets: [{
      id: "widget-1",
      type: "title",
      label: "Title",
      slot: 0,
      width: 3,
      height: 1,
      content: "Title",
      fontSize: 20,
      fontFamily: "Inter",
      fontWeight: 700,
      fontStyle: "normal",
      textColor: "#122033",
      backgroundColor: "#ffffff",
      useBackgroundColor: false,
      textAlign: "center",
      verticalAlign: "center",
    }],
  },
];

function makeId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function interpolateHexColor(start: string, end: string, ratio: number) {
  const parse = (color: string) => {
    const hex = color.replace("#", "");
    return [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16));
  };
  const from = parse(start);
  const to = parse(end);
  return `#${from.map((channel, index) => Math.round(channel + (to[index] - channel) * ratio).toString(16).padStart(2, "0")).join("")}`;
}

function duplicateWorkspaceWidget(widget: WorkspaceWidget, slot: number): WorkspaceWidget {
  return {
    ...widget,
    id: makeId("widget"),
    slot,
    barItems: widget.barItems?.map((item) => ({ ...item, id: makeId("bar-item") })),
    chartColors: widget.chartColors ? [...widget.chartColors] : undefined,
    lineAnnotations: widget.lineAnnotations?.map((annotation) => ({ ...annotation, id: makeId("line-annotation") })),
    stackCategories: widget.stackCategories?.map((category) => ({ ...category, id: makeId("stack-category") })),
    stackSeries: widget.stackSeries?.map((series) => ({ ...series, id: makeId("stack-series"), cells: [...series.cells] })),
    tableColumnWidths: widget.tableColumnWidths ? [...widget.tableColumnWidths] : undefined,
    tableRowHeights: widget.tableRowHeights ? [...widget.tableRowHeights] : undefined,
    tableCellStyles: widget.tableCellStyles ? Object.fromEntries(Object.entries(widget.tableCellStyles).map(([cell, style]) => [cell, { ...style }])) : undefined,
    gradientStops: widget.gradientStops?.map((stop) => ({ ...stop, id: makeId("gradient-stop") })),
  };
}

function getBoxTemplate(size: BoxSize) {
  return boxTemplates.find((template) => template.size === size) ?? boxTemplates[0];
}

function getCellPosition(cell: number) {
  return {
    column: (cell % workspaceColumns) + 1,
    row: Math.floor(cell / workspaceColumns) + 1,
  };
}

function getBoxFootprint(box: Pick<WorkspaceBox, "columns" | "rows">) {
  return { columns: box.columns, rows: box.rows };
}

function getBoxCells(box: Pick<WorkspaceBox, "cell" | "columns" | "rows">) {
  const footprint = getBoxFootprint(box);
  const start = getCellPosition(box.cell);
  const cells: number[] = [];

  for (let row = 0; row < footprint.rows; row += 1) {
    for (let column = 0; column < footprint.columns; column += 1) {
      const nextColumn = start.column + column;
      const nextRow = start.row + row;
      if (nextColumn <= workspaceColumns && nextRow <= workspaceRows) {
        cells.push((nextRow - 1) * workspaceColumns + (nextColumn - 1));
      }
    }
  }

  return cells;
}

function canPlaceBox(boxes: WorkspaceBox[], box: Pick<WorkspaceBox, "cell" | "columns" | "rows">, ignoreBoxId?: string) {
  const footprint = getBoxFootprint(box);
  const start = getCellPosition(box.cell);

  if (start.column + footprint.columns - 1 > workspaceColumns || start.row + footprint.rows - 1 > workspaceRows) {
    return false;
  }

  const targetCells = new Set(getBoxCells(box));
  return boxes
    .filter((currentBox) => currentBox.id !== ignoreBoxId)
    .every((currentBox) => getBoxCells(currentBox).every((cell) => !targetCells.has(cell)));
}

function mergeRemoteBoxes(currentBoxes: WorkspaceBox[], incomingBoxes: WorkspaceBox[]) {
  let changed = currentBoxes.length !== incomingBoxes.length;
  const currentById = new Map(currentBoxes.map((box) => [box.id, box]));
  const mergedBoxes = incomingBoxes.map((incomingBox) => {
    const currentBox = currentById.get(incomingBox.id);
    if (!currentBox) {
      changed = true;
      return incomingBox;
    }
    if (JSON.stringify(currentBox) === JSON.stringify(incomingBox)) return currentBox;

    const currentWidgetsById = new Map(currentBox.widgets.map((widget) => [widget.id, widget]));
    const mergedWidgets = incomingBox.widgets.map((incomingWidget) => {
      const currentWidget = currentWidgetsById.get(incomingWidget.id);
      return currentWidget && JSON.stringify(currentWidget) === JSON.stringify(incomingWidget)
        ? currentWidget
        : incomingWidget;
    });
    changed = true;
    return { ...incomingBox, widgets: mergedWidgets };
  });

  if (!changed && mergedBoxes.every((box, index) => box === currentBoxes[index])) return currentBoxes;
  return mergedBoxes;
}

function parsePageRange(input: string, pageCount: number) {
  const normalized = input.trim();
  if (!normalized) return Array.from({ length: pageCount }, (_, index) => index);
  const selected: number[] = [];

  for (const part of normalized.split(",")) {
    const token = part.trim();
    if (!token) return [];
    const rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);
      if (start < 1 || end < start || end > pageCount) return [];
      for (let page = start; page <= end; page += 1) {
        const index = page - 1;
        if (!selected.includes(index)) selected.push(index);
      }
      continue;
    }
    if (!/^\d+$/.test(token)) return [];
    const page = Number(token);
    if (page < 1 || page > pageCount) return [];
    if (!selected.includes(page - 1)) selected.push(page - 1);
  }

  return selected;
}

function getWidgetFootprint(widget: WorkspaceWidget) {
  const defaultWidth = isLargeDataWidget(widget.type) ? 4 : widget.type === "gradient-color" ? 4 : widget.type === "title" || widget.type === "text" || widget.type === "text-query" ? 3 : 1;
  const defaultHeight = isLargeDataWidget(widget.type) ? 4 : widget.type === "gradient-color" ? 2 : 1;
  return { width: widget.width ?? defaultWidth, height: widget.height ?? defaultHeight };
}

function getWidgetTextStyle(widget: WorkspaceWidget): CSSProperties {
  const horizontalPosition = widget.textAlign === "left" ? "start" : widget.textAlign === "right" ? "end" : "center";
  const verticalPosition = widget.verticalAlign === "top" ? "start" : widget.verticalAlign === "bottom" ? "end" : "center";

  return {
    alignSelf: verticalPosition,
    fontFamily: widget.fontFamily ?? "Inter",
    fontSize: `calc(${widget.fontSize ?? 16}px * var(--wms-display-scale, 1))`,
    fontStyle: widget.fontStyle ?? "normal",
    fontWeight: widget.fontWeight ?? 700,
    color: widget.textColor ?? "#122033",
    justifySelf: horizontalPosition,
    textAlign: widget.textAlign ?? "center",
  };
}

function AdaptiveWidgetText({ content: contentOverride, widget }: { content?: string; widget: WorkspaceWidget }) {
  const content = contentOverride ?? widget.content ?? widget.label;
  const allowOverflow = widget.type === "text";
  const hasCustomTextColor = Boolean(widget.textColor && widget.textColor.toLowerCase() !== "#122033");

  return (
    <span
      data-widget-custom-text-color={hasCustomTextColor ? "true" : undefined}
      className={`${allowOverflow ? "relative z-10 overflow-visible" : "max-h-full max-w-full overflow-hidden"} px-1 leading-tight text-cj-navy`}
      style={{
        ...getWidgetTextStyle(widget),
        "--widget-text-color": widget.textColor ?? "#122033",
        maxHeight: allowOverflow ? "none" : undefined,
        maxWidth: allowOverflow ? "none" : undefined,
        overflowWrap: allowOverflow ? "normal" : "anywhere",
        whiteSpace: "pre-wrap",
      } as CSSProperties}
    >
      {content}
    </span>
  );
}

function getWidgetCells(widget: WorkspaceWidget, slotColumns: number) {
  const { width, height } = getWidgetFootprint(widget);
  const startColumn = widget.slot % slotColumns;
  const startRow = Math.floor(widget.slot / slotColumns);
  const cells: number[] = [];

  for (let row = 0; row < height; row += 1) {
    for (let column = 0; column < width; column += 1) {
      cells.push((startRow + row) * slotColumns + startColumn + column);
    }
  }
  return cells;
}

function canPlaceWidget(box: WorkspaceBox, widget: WorkspaceWidget, ignoreWidgetId?: string) {
  const slotColumns = box.columns * slotsPerCellColumn;
  const slotRows = box.rows * slotsPerCellRow;
  const { width, height } = getWidgetFootprint(widget);
  const startColumn = widget.slot % slotColumns;
  const startRow = Math.floor(widget.slot / slotColumns);

  if (startColumn < 0 || startRow < 0 || startColumn + width > slotColumns || startRow + height > slotRows) return false;

  const targetCells = new Set(getWidgetCells(widget, slotColumns));
  return box.widgets
    .filter((item) => item.id !== ignoreWidgetId)
    .every((item) => getWidgetCells(item, slotColumns).every((cell) => !targetCells.has(cell)));
}

function findWidgetPlacement(box: WorkspaceBox, widget: WorkspaceWidget, requestedSlot: number, ignoreWidgetId?: string) {
  const slotColumns = box.columns * slotsPerCellColumn;
  const slotRows = box.rows * slotsPerCellRow;
  const { width, height } = getWidgetFootprint(widget);
  const requestedColumn = requestedSlot % slotColumns;
  const requestedRow = Math.floor(requestedSlot / slotColumns);
  const candidates: number[] = [];

  for (let row = 0; row <= slotRows - height; row += 1) {
    for (let column = 0; column <= slotColumns - width; column += 1) {
      candidates.push(row * slotColumns + column);
    }
  }

  candidates.sort((left, right) => {
    const leftDistance = Math.abs(Math.floor(left / slotColumns) - requestedRow) + Math.abs((left % slotColumns) - requestedColumn);
    const rightDistance = Math.abs(Math.floor(right / slotColumns) - requestedRow) + Math.abs((right % slotColumns) - requestedColumn);
    return leftDistance - rightDistance;
  });

  return candidates.find((slot) => canPlaceWidget(box, { ...widget, slot }, ignoreWidgetId)) ?? null;
}

function WorkspaceCell({ cell, occupied }: { cell: number; occupied: boolean }) {
  const { setNodeRef } = useDroppable({
    id: `workspace-cell:${cell}`,
    data: { kind: "workspace-cell", cell },
  });

  return (
    <div
      ref={setNodeRef}
      className={`min-h-20 rounded-md border border-dashed ${
        occupied ? "border-slate-200/50 bg-slate-200/35" : "border-slate-200 bg-white/55"
      }`}
    />
  );
}

function DraggableBoxTemplate({ template }: { template: BoxTemplate }) {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: `box-template:${template.size}`,
    data: { kind: "box-template", template },
  });

  return (
    <button
      ref={setNodeRef}
      className={`flex w-full items-center gap-2 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-left shadow-sm transition hover:border-cj-blue ${isDragging ? "opacity-40" : ""}`}
      type="button"
      {...attributes}
      {...listeners}
    >
      <Box className="h-4 w-4 shrink-0 text-cj-blue" />
      <span className="min-w-0">
        <strong className="block truncate text-xs font-black text-cj-navy">{template.label}</strong>
        <small className="block truncate text-[10px] font-bold text-slate-500">{template.columns} x {template.rows} cells</small>
      </span>
    </button>
  );
}

function DraggableWidgetTemplate({ template }: { template: WidgetTemplate }) {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: `widget-template:${template.type}`,
    data: { kind: "widget-template", template },
  });

  return (
    <button
      ref={setNodeRef}
      className={`flex min-w-0 items-center gap-2 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-left shadow-sm transition hover:border-cj-blue ${isDragging ? "opacity-40" : ""}`}
      type="button"
      {...attributes}
      {...listeners}
    >
      <span className="grid h-6 w-6 shrink-0 place-items-center rounded bg-slate-100 text-slate-600">{template.icon}</span>
      <strong className="min-w-0 truncate text-xs font-black text-cj-navy" title={template.label}>{template.label}</strong>
    </button>
  );
}

function WidgetSlot({
  boxId,
  dropEnabled,
  editable,
  onRemoveWidget,
  onResizeWidget,
  onUpdateWidget,
  slot,
  slotColumns,
  widget,
}: {
  boxId: string;
  dropEnabled: boolean;
  editable: boolean;
  onRemoveWidget: (boxId: string, widgetId: string) => void;
  onResizeWidget: (boxId: string, widgetId: string, slot: number, width: number, height: number) => void;
  onUpdateWidget: (boxId: string, widgetId: string, changes: Partial<WorkspaceWidget>) => void;
  slot: number;
  slotColumns: number;
  widget?: WorkspaceWidget;
}) {
  const [isEditing, setIsEditing] = useState(false);
  const [isSelected, setIsSelected] = useState(false);
  const [toolbarPosition, setToolbarPosition] = useState({ x: 16, y: 88 });
  const [excelUploads, setExcelUploads] = useState<ExcelUploadRecord[]>([]);
  const [excelSheets, setExcelSheets] = useState<string[]>([]);
  const [excelTableData, setExcelTableData] = useState<unknown[][]>([]);
  const [excelTableFormats, setExcelTableFormats] = useState<string[][]>([]);
  const [tableColumnWidthsDraft, setTableColumnWidthsDraft] = useState<number[]>([]);
  const [tableRowHeightsDraft, setTableRowHeightsDraft] = useState<number[]>([]);
  const [selectedExcelTableCell, setSelectedExcelTableCell] = useState<{ row: number; column: number; reference: string } | null>(null);
  const [excelTableStatus, setExcelTableStatus] = useState("Select a file, sheet, and range");
  const [textQueryValue, setTextQueryValue] = useState("");
  const [textQueryLoaded, setTextQueryLoaded] = useState(false);
  const [textQueryStatus, setTextQueryStatus] = useState("Select a file, sheet, and Cell");
  const [lineCellSuggestions, setLineCellSuggestions] = useState<Record<string, string[]>>({});
  const [excelCellTooltip, setExcelCellTooltip] = useState<{ reference: string; x: number; y: number; below: boolean } | null>(null);
  const [cellRangeDraft, setCellRangeDraft] = useState(widget?.cellRange ?? "A1:J7");
  const [barValues, setBarValues] = useState<number[]>([]);
  const [barMarkerValues, setBarMarkerValues] = useState<number[]>([]);
  const [pieValues, setPieValues] = useState<number[]>([]);
  const [barStatus, setBarStatus] = useState("Select a file and sheet");
  const [resolvedBarMax, setResolvedBarMax] = useState(widget?.barMax ?? 100);
  const [barMaxStatus, setBarMaxStatus] = useState("");
  const [stackValues, setStackValues] = useState<Array<Array<number | null>>>([]);
  const [updateNoticeKey, setUpdateNoticeKey] = useState(0);
  const [showUpdateNotice, setShowUpdateNotice] = useState(false);
  const [selectedGradientStopId, setSelectedGradientStopId] = useState<string | null>(null);
  const seriesSortSensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 4 } }));
  const mountedDuringEditRef = useRef(editable);
  const updateSignatureRef = useRef<string | null>(null);
  const { setNodeRef } = useDroppable({
    id: `box-slot:${boxId}:${slot}`,
    data: { kind: "box-slot", boxId, slot },
    disabled: !editable || !dropEnabled,
  });
  const draggable = useDraggable({
    id: widget ? `widget:${boxId}:${widget.id}` : `empty-widget:${boxId}:${slot}`,
    disabled: !widget || !editable,
    data: widget ? { kind: "widget", boxId, widget } : undefined,
  });
  const duplicateDraggable = useDraggable({
    id: widget ? `widget-copy:${boxId}:${widget.id}` : `empty-widget-copy:${boxId}:${slot}`,
    disabled: !widget || !editable,
    data: widget ? { kind: "widget-copy", sourceBoxId: boxId, widget } : undefined,
  });
  const footprint = widget ? getWidgetFootprint(widget) : { width: 1, height: 1 };
  const currentGradientStops = widget
    ? (widget.gradientStops?.length && widget.gradientStops.length >= 2
        ? [...widget.gradientStops]
        : [
            { id: `${widget.id}-gradient-start`, color: widget.gradientStartColor ?? "#0080c6", position: widget.gradientStartPosition ?? 0 },
            { id: `${widget.id}-gradient-end`, color: widget.gradientEndColor ?? "#e42f44", position: widget.gradientEndPosition ?? 100 },
          ]).sort((left, right) => left.position - right.position)
    : [];
  const selectedGradientStop = currentGradientStops.find((stop) => stop.id === selectedGradientStopId) ?? currentGradientStops[0];
  const gradientCss = widget
    ? `linear-gradient(${widget.gradientDirection ?? "to right"}, ${currentGradientStops.map((stop) => `${stop.color} ${stop.position}%`).join(", ")})`
    : "none";
  const slotStyle: CSSProperties = {
    gridColumn: `${(slot % slotColumns) + 1} / span ${footprint.width}`,
    gridRow: `${Math.floor(slot / slotColumns) + 1} / span ${footprint.height}`,
    zIndex: widget ? (isSelected ? 50 : widget.type === "text" ? 20 : 10) : 1,
  };
  const excelTableColumnCount = Math.max(1, ...excelTableData.map((row) => row.length));
  const equalTableColumnWidths = Array.from({ length: excelTableColumnCount }, () => 100 / excelTableColumnCount);
  const effectiveTableColumnWidths = tableColumnWidthsDraft.length === excelTableColumnCount
    ? tableColumnWidthsDraft
    : widget?.tableColumnWidths?.length === excelTableColumnCount
      ? widget.tableColumnWidths
      : equalTableColumnWidths;
  const tableWidthUnits = Math.max(100, effectiveTableColumnWidths.reduce((total, width) => total + width, 0));
  const excelTableRowCount = Math.max(1, excelTableData.length);
  const equalTableRowHeights = Array.from({ length: excelTableRowCount }, () => 100 / excelTableRowCount);
  const effectiveTableRowHeights = tableRowHeightsDraft.length === excelTableRowCount
    ? tableRowHeightsDraft
    : widget?.tableRowHeights?.length === excelTableRowCount
      ? widget.tableRowHeights
      : equalTableRowHeights;
  const tableHeightUnits = Math.max(100, effectiveTableRowHeights.reduce((total, height) => total + height, 0));
  const excelTableFontSize = excelTableColumnCount >= 16 ? 6 : excelTableColumnCount >= 12 ? 7 : excelTableColumnCount >= 8 ? 8 : 10;
  const excelTableCellPadding = excelTableColumnCount >= 12 ? "1px" : excelTableColumnCount >= 8 ? "2px" : "4px";
  const excelTableCellStyle: CSSProperties = {
    padding: excelTableCellPadding,
    textAlign: widget?.tableTextAlign ?? "left",
    verticalAlign: widget?.tableVerticalAlign ?? "middle",
  };
  const currentBarItems = widget?.barItems?.length ? widget.barItems : [{ id: "bar-1", label: "Bar 1", cell: "C7", markerCell: "D7" }];
  const currentStackCategories = widget?.stackCategories?.length
    ? widget.stackCategories
    : [{ id: "category-1", label: "Bar 1" }];
  const currentStackSeries = widget?.stackSeries?.length
    ? widget.stackSeries
    : [
        { id: "series-1", label: "Complete", cells: ["C7"] },
        { id: "series-2", label: "Pending", cells: ["D7"] },
      ];
  const chartColors = widget?.chartColors?.length ? widget.chartColors : defaultChartColors;
  const chartFontSize = widget?.chartFontSize ?? 10;
  const chartLegendPosition = widget?.chartLegendPosition ?? "bottom";
  const chartShowLegend = widget?.chartShowLegend ?? true;
  const chartBarBorderRadius = widget?.barBorderRadius ?? widget?.columnBorderRadius ?? 6;
  const isBarMarkers = widget?.type === "bar-markers";
  const barPercentageMode = (widget?.type === "bar" && (widget.barDisplayPercentage ?? false))
    || (isBarMarkers && (widget?.barMarkerDisplayPercentage ?? false));
  const effectiveBarMax = barPercentageMode ? 100 : Math.max(1, resolvedBarMax || widget?.barMax || 100);
  const markerZoneLowEnd = Math.max(0, Math.min(100, widget?.barMarkerZoneLowEnd ?? 50));
  const markerZoneMidEnd = Math.max(markerZoneLowEnd, Math.min(100, widget?.barMarkerZoneMidEnd ?? 80));
  const markerZoneOpacity = Math.max(0, Math.min(100, widget?.barMarkerZoneOpacity ?? 32));
  const markerZoneBackground = `linear-gradient(to right, ${widget?.barMarkerZoneLowColor ?? "#fee2e2"} 0%, ${widget?.barMarkerZoneLowColor ?? "#fee2e2"} ${markerZoneLowEnd}%, ${widget?.barMarkerZoneMidColor ?? "#fef3c7"} ${markerZoneLowEnd}%, ${widget?.barMarkerZoneMidColor ?? "#fef3c7"} ${markerZoneMidEnd}%, ${widget?.barMarkerZoneHighColor ?? "#dcfce7"} ${markerZoneMidEnd}%, ${widget?.barMarkerZoneHighColor ?? "#dcfce7"} 100%)`;
  const currentLineAnnotations = widget?.lineAnnotations ?? [];
  const SelectedIcon = iconOptions.find((option) => option.name === widget?.iconName)?.icon ?? Package;
  const barChartOptions: ApexOptions = {
    chart: {
      animations: chartAnimations,
      background: "transparent",
      parentHeightOffset: 0,
      toolbar: { show: false },
    },
    colors: chartColors,
    dataLabels: {
      enabled: true,
      formatter: (value, options) => {
        const actual = barPercentageMode ? `${Number(value).toFixed(1)}%` : Number(value).toLocaleString();
        if (widget?.type !== "bar-markers" || !(widget.barMarkerShowValue ?? true)) return actual;
        const target = Number(barMarkerValues[options?.dataPointIndex ?? 0] ?? 0);
        return `${actual} | T:${barPercentageMode ? `${target.toFixed(1)}%` : target.toLocaleString()}`;
      },
      style: { fontSize: `${widget?.type === "bar-markers" ? (widget.barMarkerFontSize ?? 10) : chartFontSize}px`, fontWeight: 800 },
    },
    grid: { borderColor: "#e2e8f0", padding: { left: 4, right: 8, top: -8, bottom: -8 } },
    legend: {
      show: false,
    },
    plotOptions: { bar: { borderRadius: chartBarBorderRadius, borderRadiusApplication: "around", distributed: true, horizontal: true } },
    states: { active: { filter: { type: "none" } }, hover: { filter: { type: "none" } } },
    tooltip: { enabled: true, y: { formatter: (value) => barPercentageMode ? `${Number(value).toFixed(2)}%` : formatChartTooltipValue(value) } },
    xaxis: {
      categories: currentBarItems.map((item) => item.label || item.cell),
      max: effectiveBarMax,
      min: 0,
      labels: { formatter: barPercentageMode ? (value) => `${Math.round(Number(value))}%` : undefined, style: { fontSize: `${chartFontSize}px` } },
    },
    yaxis: { labels: { maxWidth: 90, style: { fontSize: `${chartFontSize}px`, fontWeight: 700 } } },
  };
  const isStackColumn100 = widget?.type === "stack-100-column";
  const isStackColumn = widget?.type === "stack-column" || isStackColumn100;
  const isStackColumnPercentage = isStackColumn100 || (widget?.type === "stack-column" && (widget.stackColumnPercentage ?? false));
  const isStackBar = widget?.type === "stack-bar" || widget?.type === "stack-column" || widget?.type === "stack-100-bar" || isStackColumn100;
  const isLineChart = widget?.type === "basic-line" || widget?.type === "line-annotations";
  const isAnnotatedLine = widget?.type === "line-annotations";
  const isColumnChart = widget?.type === "column-rotated-labels";
  const isRadarChart = widget?.type === "radar-polygon";
  const isSimplePie = widget?.type === "simple-pie" || widget?.type === "simple-donut";
  const isTextQuery = widget?.type === "text-query";
  const usesMaximumValue = (widget?.type === "bar" && !barPercentageMode) || widget?.type === "bar-markers" || widget?.type === "stack-bar" || widget?.type === "stack-column" || isLineChart || isColumnChart || isRadarChart;

  function updateGradientStops(stops: WorkspaceWidget["gradientStops"]) {
    if (!widget || !stops?.length) return;
    const sortedStops = [...stops].sort((left, right) => left.position - right.position);
    onUpdateWidget(boxId, widget.id, {
      gradientStops: sortedStops,
      gradientStartColor: sortedStops[0].color,
      gradientStartPosition: sortedStops[0].position,
      gradientEndColor: sortedStops[sortedStops.length - 1].color,
      gradientEndPosition: sortedStops[sortedStops.length - 1].position,
    });
  }

  function addGradientStop(position: number) {
    if (!widget || currentGradientStops.length < 2) return;
    const boundedPosition = Math.max(0, Math.min(100, Math.round(position)));
    const rightIndex = currentGradientStops.findIndex((stop) => stop.position >= boundedPosition);
    const right = currentGradientStops[rightIndex < 0 ? currentGradientStops.length - 1 : rightIndex];
    const left = currentGradientStops[Math.max(0, (rightIndex < 0 ? currentGradientStops.length : rightIndex) - 1)];
    const span = Math.max(1, right.position - left.position);
    const nextStop = {
      id: makeId("gradient-stop"),
      color: interpolateHexColor(left.color, right.color, Math.max(0, Math.min(1, (boundedPosition - left.position) / span))),
      position: boundedPosition,
    };
    updateGradientStops([...currentGradientStops, nextStop]);
    setSelectedGradientStopId(nextStop.id);
  }

  useEffect(() => {
    if (!widget || !usesMaximumValue) return;
    const input = widget.barMaxInput?.trim() || String(widget.barMax ?? 100);
    const sourceUploadId = widget.sourceUploadId;
    const sheetName = widget.sheetName;
    const fallbackMaximum = widget.barMax ?? 100;
    const shouldRoundMaximum = (widget.type === "basic-line" || widget.type === "column-rotated-labels" || widget.type === "stack-column") && /^MAX\s*\(/i.test(input);
    let active = true;

    function loadMaximum() {
      setBarMaxStatus("Resolving maximum...");
      resolveMaximumInput(sourceUploadId, sheetName, input)
        .then((value) => {
          if (!active) return;
          if (!Number.isFinite(value) || value <= 0) throw new Error("Maximum must be greater than zero");
          const maximum = shouldRoundMaximum ? roundMaximumUp(value) : value;
          setResolvedBarMax(maximum);
          setBarMaxStatus(shouldRoundMaximum && maximum !== value
            ? `Maximum: ${formatChartTooltipValue(maximum)} (from ${formatChartTooltipValue(value)})`
            : `Maximum: ${formatChartTooltipValue(maximum)}`);
        })
        .catch(() => {
          if (!active) return;
          setResolvedBarMax(fallbackMaximum);
          setBarMaxStatus("Cannot resolve maximum value");
        });
    }

    function handleUploadReplaced(event: Event) {
      if ((event as CustomEvent<{ uploadId: string }>).detail.uploadId === sourceUploadId) loadMaximum();
    }

    loadMaximum();
    window.addEventListener("excel-upload-replaced", handleUploadReplaced);
    return () => {
      active = false;
      window.removeEventListener("excel-upload-replaced", handleUploadReplaced);
    };
  }, [usesMaximumValue, widget?.barMax, widget?.barMaxInput, widget?.sheetName, widget?.sourceUploadId, widget?.type]);
  const pieChartOptions: ApexOptions = {
    chart: {
      animations: chartAnimations,
      background: "transparent",
      ...(widget?.type === "simple-pie" && widget.pieShowValueCallouts
        ? { events: {
            animationEnd: drawPieValueCallouts,
            mounted: drawPieValueCallouts,
            updated: drawPieValueCallouts,
          } }
        : {}),
      parentHeightOffset: 0,
      toolbar: { show: false },
    },
    colors: chartColors,
    dataLabels: {
      enabled: true,
      formatter: (percentage, options) => {
        if (widget?.type === "simple-pie" && widget.pieShowValueCallouts) {
          const index = options?.seriesIndex ?? 0;
          return `${currentBarItems[index]?.label ?? `Series ${index + 1}`}: ${formatChartTooltipValue(pieValues[index] ?? 0)}`;
        }
        return `${Number(percentage).toFixed(1)}%`;
      },
      style: { colors: [widget?.type === "simple-pie" && widget.pieShowValueCallouts ? "#334155" : "#ffffff"], fontSize: `${chartFontSize}px`, fontWeight: 800 },
      dropShadow: { enabled: true, blur: 2, opacity: 0.7 },
    },
    labels: currentBarItems.map((item) => item.label || item.cell),
    legend: { fontSize: `${chartFontSize}px`, fontWeight: 700, position: chartLegendPosition, show: chartShowLegend },
    plotOptions: {
      pie: {
        customScale: widget?.type === "simple-donut" ? 1 : 1,
        dataLabels: {
          offset: widget?.type === "simple-pie" ? (widget.pieShowValueCallouts ? 24 : -12) : 0,
          minAngleToShowLabel: 8,
        },
        expandOnClick: widget?.type !== "simple-donut",
        donut: {
          size: "56%",
          labels: {
            show: widget?.type === "simple-donut",
            name: { show: true },
            total: { show: true, label: "Total", formatter: (chart) => chart.globals.seriesTotals.reduce((sum: number, value: number) => sum + value, 0).toLocaleString() },
            value: { show: true, formatter: (value) => Number(value).toLocaleString() },
          },
        },
      },
    },
    states: { active: { filter: { type: "none" } }, hover: { filter: { type: "lighten" } } },
    stroke: {
      colors: ["var(--chart-surface)"],
      lineCap: "round",
      width: widget?.type === "simple-donut" ? 7 : 3,
    },
    tooltip: { y: { formatter: formatChartTooltipValue } },
  };
  const stackChartOptions: ApexOptions = {
    chart: {
      animations: chartAnimations,
      background: "transparent",
      events: {
        animationEnd: (chart) => {
          roundStackBarSegments(chart, chartBarBorderRadius);
        },
        mounted: (chart) => {
          scheduleStackBarRounding(chart, chartBarBorderRadius);
        },
        updated: (chart) => {
          scheduleStackBarRounding(chart, chartBarBorderRadius);
        },
      },
      parentHeightOffset: 0,
      stacked: true,
      stackType: widget?.type === "stack-100-bar" || isStackColumnPercentage ? "100%" : "normal",
      toolbar: { show: false },
    },
    colors: chartColors,
    dataLabels: {
      enabled: true,
      background: { enabled: false },
      ...(isStackColumn || isStackColumnPercentage ? {
        formatter: (value: number, options) => {
          const categoryIndex = options?.dataPointIndex ?? 0;
          const total = currentStackSeries.reduce((sum, _, seriesIndex) => sum + Number(stackValues[seriesIndex]?.[categoryIndex] ?? 0), 0);
          const percentage = total > 0 ? (Number(value) / total) * 100 : 0;
          return `${percentage.toFixed(1)}%`;
        },
      } : {}),
      style: { fontSize: `${chartFontSize}px`, fontWeight: 800 },
    },
    grid: {
      borderColor: "#e2e8f0",
      padding: {
        left: 4,
        right: 8,
        top: widget?.type === "stack-100-bar" ? -16 : isStackColumn100 ? 4 : isStackColumn ? 4 : -4,
        bottom: widget?.type === "stack-100-bar" ? -12 : isStackColumn100 ? 0 : isStackColumn ? 4 : -8,
      },
    },
    legend: {
      fontSize: `${chartFontSize}px`,
      fontWeight: 700,
      offsetY: widget?.type === "stack-100-bar" ? 6 : 0,
      position: chartLegendPosition,
      show: chartShowLegend,
    },
    plotOptions: {
      bar: {
        borderRadius: chartBarBorderRadius,
        borderRadiusApplication: "around",
        borderRadiusWhenStacked: "all",
        ...(isStackColumn
          ? { columnWidth: "64%" }
          : { barHeight: widget?.type === "stack-100-bar" ? "92%" : "70%" }),
        horizontal: !isStackColumn,
      },
    },
    states: { active: { filter: { type: "none" } }, hover: { filter: { type: "none" } } },
    stroke: { colors: ["var(--chart-surface)"], lineCap: "round", show: true, width: 2 },
    tooltip: { enabled: true, y: { formatter: formatChartTooltipValue } },
    xaxis: {
      categories: currentStackCategories.map((category) => category.label),
      max: isStackColumn ? undefined : widget?.type === "stack-100-bar" ? 100 : effectiveBarMax,
      min: isStackColumn ? undefined : 0,
      labels: {
        formatter: widget?.type === "stack-100-bar" ? (value) => `${Math.round(Number(value))}%` : undefined,
        style: { fontSize: `${chartFontSize}px` },
      },
    },
    yaxis: isStackColumnPercentage
      ? { max: 100, min: 0, labels: { formatter: (value) => isStackColumn100 ? `${Math.round(Number(value))}%` : `${Math.round(Number(value))}`, style: { fontSize: `${chartFontSize}px`, fontWeight: 700 } } }
      : isStackColumn
        ? { max: effectiveBarMax, min: 0, labels: { formatter: (value) => Math.round(Number(value) / 1000) * 1000 === 0 ? "0" : (Math.round(Number(value) / 1000) * 1000).toLocaleString(), style: { fontSize: `${chartFontSize}px`, fontWeight: 700 } } }
        : { labels: { maxWidth: 90, style: { fontSize: `${chartFontSize}px`, fontWeight: 700 } } },
  };
  const lineChartOptions: ApexOptions = {
    annotations: {
      xaxis: currentLineAnnotations.filter((item) => item.axis === "x" && item.value).map((item) => ({
        x: item.value,
        borderColor: item.color,
        strokeDashArray: 4,
        label: { text: item.label, style: { background: item.color, color: "#ffffff", fontSize: `${chartFontSize}px` } },
      })),
      yaxis: currentLineAnnotations.filter((item) => item.axis === "y" && Number.isFinite(Number(item.value))).map((item) => ({
        y: Number(item.value),
        borderColor: item.color,
        strokeDashArray: 4,
        label: { text: item.label, style: { background: item.color, color: "#ffffff", fontSize: `${chartFontSize}px` } },
      })),
    },
    chart: { animations: chartAnimations, background: "transparent", parentHeightOffset: 0, toolbar: { show: false }, zoom: { enabled: false } },
    colors: chartColors,
    dataLabels: { enabled: false },
    grid: { borderColor: "#e2e8f0", padding: { left: 4, right: 10, top: -4, bottom: 0 } },
    legend: { fontSize: `${chartFontSize}px`, fontWeight: 700, position: chartLegendPosition, show: chartShowLegend },
    markers: { size: widget?.lineShowMarkers ?? true ? 4 : 0, strokeColors: "var(--chart-surface)", strokeWidth: 2, hover: { sizeOffset: 2 } },
    stroke: { curve: widget?.lineCurve ?? "smooth", lineCap: "round", width: widget?.lineStrokeWidth ?? 3 },
    tooltip: { shared: true, intersect: false, y: { formatter: formatChartTooltipValue } },
    xaxis: { categories: currentStackCategories.map((category) => category.label), labels: { style: { fontSize: `${chartFontSize}px` } } },
    yaxis: {
      max: effectiveBarMax,
      min: 0,
      labels: { formatter: (value) => Number(value).toLocaleString(), style: { fontSize: `${chartFontSize}px` } },
    },
  };
  const columnChartOptions: ApexOptions = {
    chart: {
      animations: chartAnimations,
      background: "transparent",
      events: {
        mounted: (chart) => roundColumnDataLabelBackgrounds(chart),
        updated: (chart) => roundColumnDataLabelBackgrounds(chart),
      },
      parentHeightOffset: 0,
      toolbar: { show: false },
      zoom: { enabled: false },
    },
    colors: chartColors,
    dataLabels: {
      background: {
        borderColor: "var(--column-label-border)",
        borderRadius: 999,
        borderWidth: 1,
        enabled: widget?.columnShowDataLabelBackground ?? true,
        foreColor: "var(--column-label-text)",
        opacity: 1,
        padding: 6,
      },
      enabled: widget?.columnShowDataLabels ?? true,
      formatter: (value) => Number(value).toLocaleString(),
      offsetY: -10,
      style: { fontSize: `${chartFontSize}px`, fontWeight: 800 },
    },
    grid: { borderColor: "#e2e8f0", padding: { left: 4, right: 8, top: 8, bottom: 4 } },
    legend: { fontSize: `${chartFontSize}px`, fontWeight: 700, position: chartLegendPosition, show: chartShowLegend },
    plotOptions: {
      bar: {
        borderRadius: chartBarBorderRadius,
        borderRadiusApplication: "end",
        columnWidth: `${widget?.columnWidth ?? 55}%`,
        dataLabels: { position: "top" },
        horizontal: false,
      },
    },
    states: { active: { filter: { type: "none" } }, hover: { filter: { type: "none" } } },
    tooltip: { shared: true, intersect: false, y: { formatter: formatChartTooltipValue } },
    xaxis: {
      categories: currentStackCategories.map((category) => category.label),
      labels: {
        hideOverlappingLabels: false,
        rotate: widget?.columnLabelRotation ?? -45,
        rotateAlways: true,
        style: { fontSize: `${chartFontSize}px` },
        trim: false,
      },
    },
    yaxis: {
      max: effectiveBarMax,
      min: 0,
      labels: { formatter: (value) => Number(value).toLocaleString(), style: { fontSize: `${chartFontSize}px` } },
    },
  };
  const radarChartOptions: ApexOptions = {
    chart: { animations: chartAnimations, background: "transparent", fontFamily: "Arial, sans-serif", offsetY: 2, parentHeightOffset: 0, toolbar: { show: false } },
    colors: chartColors,
    dataLabels: { enabled: false },
    fill: { opacity: widget?.radarFillOpacity ?? 0.2 },
    grid: { padding: { bottom: -14, left: -12, right: -12, top: -14 } },
    legend: { fontSize: `${chartFontSize}px`, fontWeight: 700, position: chartLegendPosition, show: chartShowLegend },
    markers: {
      size: widget?.radarMarkerSize ?? 4,
      strokeColors: ["#ffffff"],
      strokeWidth: 2,
      hover: { sizeOffset: 2 },
    },
    plotOptions: {
      radar: {
        polygons: {
          fill: { colors: [widget?.radarPolygonColor1 ?? "#f8fafc", widget?.radarPolygonColor2 ?? "#eef2f6"] },
          strokeColors: widget?.radarPolygonStrokeColor ?? "#cbd5e1",
        },
      },
    },
    stroke: { width: widget?.radarStrokeWidth ?? 2 },
    tooltip: { intersect: false, shared: false, y: { formatter: formatChartTooltipValue } },
    xaxis: {
      categories: currentStackCategories.map((category) => category.label),
      labels: { style: { fontFamily: "Arial, sans-serif", fontSize: `${chartFontSize}px`, fontWeight: 800 } },
    },
    yaxis: { max: effectiveBarMax, min: 0, show: false, tickAmount: 5 },
  };
  const radarCategoryCount = currentStackCategories.length;
  const radarSeries = currentStackSeries.map((series, seriesIndex) => ({
    name: series.label || `Series ${seriesIndex + 1}`,
    data: Array.from({ length: radarCategoryCount }, (_, categoryIndex) => {
      const value = Number(stackValues[seriesIndex]?.[categoryIndex]);
      return Number.isFinite(value) ? Math.max(0, value) : 0;
    }),
  }));
  const canRenderRadar = radarCategoryCount >= 3 && radarSeries.length > 0 && stackValues.length > 0;
  const barChartViewportStyle: CSSProperties = chartShowLegend
    ? chartLegendPosition === "top"
      ? { height: "calc(100% - 20px)", top: 20 }
      : { height: "calc(100% - 20px)", top: 0 }
    : { height: "100%", top: 0 };

  function renderBarLegend() {
    if (!chartShowLegend) return null;
    return (
      <div className={`absolute inset-x-2 z-20 flex min-h-5 flex-wrap items-center justify-center gap-x-3 gap-y-0.5 overflow-hidden text-[9px] font-bold text-slate-600 ${chartLegendPosition === "top" ? "top-0" : "bottom-0"}`}>
        {currentBarItems.map((item, index) => (
          <span key={item.id} className="inline-flex min-w-0 items-center gap-1">
            <i className="h-2 w-2 shrink-0 rounded-sm" style={{ backgroundColor: chartColors[index % chartColors.length] }} />
            <span className="max-w-24 truncate">{item.label || item.cell}</span>
          </span>
        ))}
        {isBarMarkers ? (
          <span className="inline-flex items-center gap-1">
            <i className="h-0 w-3 shrink-0 border-t-2" style={{ borderColor: widget?.barMarkerColor ?? "#e42f44" }} />
            <span>{widget?.barMarkerLabelText || "Target"}</span>
          </span>
        ) : null}
      </div>
    );
  }

  function renderBasicBarChart() {
    const axisSteps = [0, 25, 50, 75, 100];
    return (
      <div className="flex h-full min-h-0 flex-col pb-4 pt-1">
        <div className="flex min-h-0 flex-1 flex-col justify-evenly gap-1">
          {currentBarItems.map((item, index) => {
            const value = Math.max(0, Number(barValues[index] ?? 0));
            const widthPercent = Math.max(0, Math.min(100, (value / effectiveBarMax) * 100));
            const displayValue = barPercentageMode ? `${value.toFixed(1)}%` : value.toLocaleString(undefined, { maximumFractionDigits: 2 });
            return (
              <div key={item.id} className="grid min-h-0 flex-1 grid-cols-[clamp(42px,15%,72px)_minmax(0,1fr)] items-center gap-1.5 px-1">
                <div className="truncate text-right font-bold text-slate-600" style={{ fontSize: `${chartFontSize}px` }} title={item.label || item.cell}>
                  {item.label || item.cell}
                </div>
                <div className="relative h-[62%] min-h-3 overflow-visible rounded-sm" title={`${item.label || item.cell}: ${displayValue}`}>
                  <div className="pointer-events-none absolute inset-0 z-0 rounded-sm" aria-hidden="true">
                    {axisSteps.map((step) => (
                      <span key={step} className="absolute inset-y-0 w-px bg-slate-200" style={{ left: `${step}%` }} />
                    ))}
                  </div>
                  <div
                    className="relative z-10 flex h-full min-w-[2px] items-center justify-center rounded-[inherit] px-1.5 text-center font-extrabold text-white transition-[width] duration-700 ease-out"
                    style={{ backgroundColor: chartColors[index % chartColors.length], borderRadius: `${chartBarBorderRadius}px`, fontSize: `${chartFontSize}px`, width: `${widthPercent}%` }}
                  >
                    {widthPercent >= 18 ? displayValue : null}
                  </div>
                  {widthPercent < 18 ? (
                    <span className="absolute top-1/2 -translate-y-1/2 whitespace-nowrap font-extrabold text-slate-700" style={{ fontSize: `${chartFontSize}px`, left: `calc(${widthPercent}% + 4px)` }}>
                      {displayValue}
                    </span>
                  ) : null}
                </div>
              </div>
            );
          })}
        </div>
        <div className="grid grid-cols-[clamp(42px,15%,72px)_minmax(0,1fr)] gap-1.5 px-1 text-[8px] font-bold text-slate-400">
          <span />
          <div className="flex justify-between border-t border-slate-300 pt-0.5">
            {axisSteps.map((step) => <span key={step}>{barPercentageMode ? `${step}%` : Math.round((effectiveBarMax * step) / 100).toLocaleString()}</span>)}
          </div>
        </div>
      </div>
    );
  }

  function renderBarMarkerLabels() {
    if (!isBarMarkers || !(widget?.barMarkerShowLabel ?? true)) return null;
    const markerColor = widget?.barMarkerColor ?? "#e42f44";
    const labelText = widget?.barMarkerLabelText || "Target";
    return (
      <div className="pointer-events-none absolute bottom-0 right-2 top-0 z-30 overflow-visible" style={{ left: "clamp(52px, 22%, 96px)" }} aria-hidden="true">
        {currentBarItems.map((item, index) => {
          const target = Math.max(0, Number(barMarkerValues[index] ?? 0));
          const targetPercent = Math.max(0, Math.min(100, (target / effectiveBarMax) * 100));
          const rowPercent = ((index + 0.5) / Math.max(1, currentBarItems.length)) * 100;
          return (
            <span
              key={`${item.id}-target-label`}
              className="absolute whitespace-nowrap rounded px-1 py-px font-extrabold leading-tight text-white shadow-sm"
              style={{
                backgroundColor: markerColor,
                fontSize: `${widget?.barMarkerLabelFontSize ?? 9}px`,
                left: `${targetPercent}%`,
                top: `${rowPercent}%`,
                transform: "translate(-50%, -145%)",
              }}
            >
              {labelText}
            </span>
          );
        })}
      </div>
    );
  }

  useEffect(() => {
    if (!editable) {
      setIsSelected(false);
      setIsEditing(false);
      setExcelCellTooltip(null);
    }
  }, [editable]);

  useEffect(() => {
    if (!isSelected || !widget) return;
    const widgetId = widget.id;

    function handleOutsidePointer(event: globalThis.PointerEvent) {
      const target = event.target as Element | null;
      const selectedWidget = target?.closest("[data-workspace-widget]");
      const selectedToolbar = target?.closest("[data-widget-toolbar]");
      const isCurrentWidget = selectedWidget?.getAttribute("data-workspace-widget") === widgetId;
      const isCurrentToolbar = selectedToolbar?.getAttribute("data-widget-toolbar") === widgetId;
      if (!isCurrentWidget && !isCurrentToolbar) {
        setIsSelected(false);
        setIsEditing(false);
      }
    }

    document.addEventListener("pointerdown", handleOutsidePointer);
    return () => document.removeEventListener("pointerdown", handleOutsidePointer);
  }, [isSelected, widget]);

  useEffect(() => {
    if (!isSelected || (widget?.type !== "excel-table" && widget?.type !== "bar" && !isTextQuery && !isBarMarkers && !isStackBar && !isLineChart && !isColumnChart && !isRadarChart && !isSimplePie)) return;
    fetchExcelUploads()
      .then(setExcelUploads)
      .catch(() => setExcelTableStatus("Cannot load uploaded files"));
  }, [isBarMarkers, isColumnChart, isLineChart, isRadarChart, isSelected, isSimplePie, isStackBar, isTextQuery, widget?.type]);

  useEffect(() => {
    if ((widget?.type !== "excel-table" && widget?.type !== "bar" && !isTextQuery && !isBarMarkers && !isStackBar && !isLineChart && !isColumnChart && !isRadarChart && !isSimplePie) || !widget.sourceUploadId) {
      setExcelSheets([]);
      return;
    }
    fetchExcelSheets(widget.sourceUploadId)
      .then((response) => setExcelSheets(response.data))
      .catch(() => setExcelTableStatus("Cannot load workbook sheets"));
  }, [isBarMarkers, isColumnChart, isLineChart, isRadarChart, isSimplePie, isStackBar, isTextQuery, widget?.sourceUploadId, widget?.type]);

  useEffect(() => {
    setCellRangeDraft(widget?.cellRange ?? "A1:J7");
    if (widget?.type !== "excel-table" || !widget.sourceUploadId || !widget.sheetName || !widget.cellRange) {
      setExcelTableData([]);
      setExcelTableFormats([]);
      return;
    }

    const uploadId = widget.sourceUploadId;
    const sheetName = widget.sheetName;
    const cellRange = widget.cellRange;

    function loadConfiguredRange() {
      setExcelTableStatus("Loading range...");
      fetchExcelRange(uploadId, sheetName, cellRange)
        .then((response) => {
          setExcelTableData(response.data);
          setExcelTableFormats(response.number_formats ?? []);
          setExcelTableStatus(`${response.sheet} | ${response.cell_range}`);
        })
        .catch(() => {
          setExcelTableData([]);
          setExcelTableFormats([]);
          setExcelTableStatus("Cannot load the selected range");
        });
    }

    function handleUploadReplaced(event: Event) {
      const replacedUploadId = (event as CustomEvent<{ uploadId: string }>).detail.uploadId;
      if (replacedUploadId !== uploadId) return;
      fetchExcelSheets(uploadId).then((response) => setExcelSheets(response.data)).catch(() => setExcelSheets([]));
      loadConfiguredRange();
    }

    loadConfiguredRange();
    window.addEventListener("excel-upload-replaced", handleUploadReplaced);
    return () => window.removeEventListener("excel-upload-replaced", handleUploadReplaced);
  }, [widget?.cellRange, widget?.sheetName, widget?.sourceUploadId, widget?.type]);

  useEffect(() => {
    setTableColumnWidthsDraft(widget?.tableColumnWidths ?? []);
  }, [widget?.id, widget?.tableColumnWidths]);

  useEffect(() => {
    setTableRowHeightsDraft(widget?.tableRowHeights ?? []);
  }, [widget?.id, widget?.tableRowHeights]);

  useEffect(() => {
    setSelectedExcelTableCell(null);
  }, [widget?.cellRange, widget?.id, widget?.sheetName, widget?.sourceUploadId]);

  useEffect(() => {
    if (!isTextQuery || !widget?.sourceUploadId || !widget.sheetName || !widget.textQueryCell?.trim()) {
      setTextQueryValue("");
      setTextQueryLoaded(false);
      setTextQueryStatus("Select a file, sheet, and Cell");
      return;
    }

    const uploadId = widget.sourceUploadId;
    const sheetName = widget.sheetName;
    const cell = widget.textQueryCell.trim().toUpperCase();

    function loadTextQuery() {
      setTextQueryStatus("Loading Cell...");
      const directCell = /^\$?[A-Z]+\$?\d+$/.test(cell);
      const valueRequest = directCell
        ? fetchExcelRange(uploadId, sheetName, cell).then((response) => ({ value: response.data[0]?.[0], numberFormat: response.number_formats?.[0]?.[0], status: `${response.sheet} | ${cell}` }))
        : evaluateNumericExpression(uploadId, sheetName, cell).then((value) => ({ value, numberFormat: undefined, status: `${sheetName} | Formula` }));
      valueRequest
        .then(({ value, numberFormat, status }) => {
          setTextQueryValue(formatTextQueryValue(value, numberFormat));
          setTextQueryLoaded(true);
          setTextQueryStatus(status);
        })
        .catch(() => {
          setTextQueryValue("");
          setTextQueryLoaded(false);
          setTextQueryStatus("Cannot evaluate the Cell or formula");
        });
    }

    function handleUploadReplaced(event: Event) {
      if ((event as CustomEvent<{ uploadId: string }>).detail.uploadId !== uploadId) return;
      fetchExcelSheets(uploadId).then((response) => setExcelSheets(response.data)).catch(() => setExcelSheets([]));
      loadTextQuery();
    }

    const loadTimer = window.setTimeout(loadTextQuery, 250);
    window.addEventListener("excel-upload-replaced", handleUploadReplaced);
    return () => {
      window.clearTimeout(loadTimer);
      window.removeEventListener("excel-upload-replaced", handleUploadReplaced);
    };
  }, [isTextQuery, widget?.sheetName, widget?.sourceUploadId, widget?.textQueryCell]);

  useEffect(() => {
    if (widget?.type !== "bar" && !isBarMarkers) {
      setBarValues([]);
      setBarMarkerValues([]);
      return;
    }

    const uploadId = widget.sourceUploadId;
    const sheetName = widget.sheetName;
    const items = widget.barItems?.length ? widget.barItems : [{ id: "bar-1", label: "Bar 1", cell: "C7" }];
    const needsTargets = isBarMarkers || barPercentageMode;
    const inputs = items.flatMap((item) => needsTargets ? [item.cell, item.markerCell || "100"] : [item.cell]);
    if (inputs.some(isExcelInput) && (!uploadId || !sheetName)) {
      setBarValues([]);
      setBarMarkerValues([]);
      setBarStatus("Select a file and sheet for Cell references");
      return;
    }

    function loadBarValues() {
      setBarStatus("Loading bar values...");
      fetchNumericInputLookup(uploadId, sheetName, inputs)
        .then((cellLookup) => {
          const valueRequests = items.map((item) => resolveNumericInput(uploadId, sheetName, item.cell, cellLookup));
          const markerRequests = needsTargets
            ? items.map((item) => resolveNumericInput(uploadId, sheetName, item.markerCell || "100", cellLookup))
            : [];
          return Promise.all([...valueRequests, ...markerRequests]);
        })
        .then((resolvedValues) => {
          const rawValues = resolvedValues.slice(0, items.length).map((value) => Math.max(0, value));
          const markers = resolvedValues.slice(items.length).map((value) => Math.max(0, value));
          const percentageMaximum = Math.max(1, resolvedBarMax || widget?.barMax || 100);
          const values = barPercentageMode
            ? isBarMarkers
              ? rawValues.map((value) => (value / percentageMaximum) * 100)
              : rawValues.map((value, index) => markers[index] > 0 ? (value / markers[index]) * 100 : 0)
            : rawValues;
          const displayMarkers = isBarMarkers && barPercentageMode
            ? markers.map((value) => (value / percentageMaximum) * 100)
            : markers;
          setBarValues(values);
          setBarMarkerValues(displayMarkers);
          setBarStatus(`${items.length} bar${items.length === 1 ? "" : "s"} loaded`);
        })
        .catch(() => {
          setBarValues([]);
          setBarMarkerValues([]);
          setBarStatus("Cannot load one or more cells");
        });
    }

    function handleUploadReplaced(event: Event) {
      const replacedUploadId = (event as CustomEvent<{ uploadId: string }>).detail.uploadId;
      if (replacedUploadId !== uploadId) return;
      fetchExcelSheets(uploadId).then((response) => setExcelSheets(response.data)).catch(() => setExcelSheets([]));
      loadBarValues();
    }

    const loadTimer = window.setTimeout(loadBarValues, 350);
    window.addEventListener("excel-upload-replaced", handleUploadReplaced);
    return () => {
      window.clearTimeout(loadTimer);
      window.removeEventListener("excel-upload-replaced", handleUploadReplaced);
    };
  }, [barPercentageMode, isBarMarkers, resolvedBarMax, widget?.barItems, widget?.barMax, widget?.sheetName, widget?.sourceUploadId, widget?.type]);

  useEffect(() => {
    if (!isSimplePie || !widget) {
      setPieValues([]);
      return;
    }

    const uploadId = widget.sourceUploadId;
    const sheetName = widget.sheetName;
    const items = currentBarItems;
    if (items.some((item) => isExcelInput(item.cell)) && (!uploadId || !sheetName)) {
      setPieValues([]);
      setBarStatus("Select a file and sheet for Cell references");
      return;
    }

    function loadPieValues() {
      setBarStatus("Loading slice values...");
      fetchNumericInputLookup(uploadId, sheetName, items.map((item) => item.cell))
        .then((cellLookup) => Promise.all(items.map((item) => resolveNumericInput(uploadId, sheetName, item.cell, cellLookup))))
        .then((resolvedValues) => {
          const values = resolvedValues.map((value) => value >= 0 ? value : 0);
          setPieValues(values);
          setBarStatus(values.some((value) => value > 0) ? `${items.length} slice${items.length === 1 ? "" : "s"} loaded` : "No positive values to display");
        })
        .catch(() => {
          setPieValues([]);
          setBarStatus("Cannot load one or more cells");
        });
    }

    function handleUploadReplaced(event: Event) {
      if ((event as CustomEvent<{ uploadId: string }>).detail.uploadId === uploadId) loadPieValues();
    }

    const loadTimer = window.setTimeout(loadPieValues, 350);
    window.addEventListener("excel-upload-replaced", handleUploadReplaced);
    return () => {
      window.clearTimeout(loadTimer);
      window.removeEventListener("excel-upload-replaced", handleUploadReplaced);
    };
  }, [isSimplePie, widget?.barItems, widget?.sheetName, widget?.sourceUploadId]);

  useEffect(() => {
    if ((!isStackBar && !isLineChart && !isColumnChart && !isRadarChart) || !widget) {
      setStackValues([]);
      return;
    }

    const uploadId = widget.sourceUploadId;
    const sheetName = widget.sheetName;
    const widgetType = widget.type;
    const series = currentStackSeries;
    const cells = series.flatMap((item) => currentStackCategories.map((_, index) => item.cells[index] || "A1"));
    if (cells.some(isExcelInput) && (!uploadId || !sheetName)) {
      setStackValues([]);
      setBarStatus("Select a file and sheet for Cell references");
      return;
    }

    function loadStackValues() {
      setBarStatus(isLineChart ? "Loading line values..." : isColumnChart ? "Loading column values..." : isRadarChart ? "Loading radar values..." : "Loading stacked values...");
      fetchNumericInputLookup(uploadId, sheetName, cells)
        .then((cellLookup) => Promise.all(cells.map((cell) => resolveNullableNumericInput(uploadId, sheetName, cell, widgetType === "basic-line" && (widget?.lineNullMissing ?? false), cellLookup))))
        .then((values) => {
          const rawChartValues = widgetType === "basic-line"
            ? values.map((value) => value === null ? null : Math.abs(value))
            : isLineChart
              ? values
              : values.map((value) => Math.max(0, value ?? 0));
          const chartValues = widgetType === "stack-column" && (widget?.stackColumnPercentage ?? false)
            ? rawChartValues.map((_, valueIndex) => {
              const categoryIndex = valueIndex % currentStackCategories.length;
              const total = rawChartValues.reduce<number>((sum, value, seriesValueIndex) => seriesValueIndex % currentStackCategories.length === categoryIndex ? sum + Number(value ?? 0) : sum, 0);
              return total > 0 ? (Number(rawChartValues[valueIndex] ?? 0) / total) * 100 : 0;
            })
            : rawChartValues;
          setStackValues(series.map((_, index) => chartValues.slice(index * currentStackCategories.length, (index + 1) * currentStackCategories.length)));
          setBarStatus(`${series.length} series loaded`);
        })
        .catch(() => {
          setStackValues([]);
          setBarStatus("Cannot load one or more cells");
        });
    }

    function handleUploadReplaced(event: Event) {
      if ((event as CustomEvent<{ uploadId: string }>).detail.uploadId === uploadId) loadStackValues();
    }

    const loadTimer = window.setTimeout(loadStackValues, 350);
    window.addEventListener("excel-upload-replaced", handleUploadReplaced);
    return () => {
      window.clearTimeout(loadTimer);
      window.removeEventListener("excel-upload-replaced", handleUploadReplaced);
    };
  }, [isColumnChart, isLineChart, isRadarChart, isStackBar, widget?.lineNullMissing, widget?.sheetName, widget?.sourceUploadId, widget?.stackCategories, widget?.stackSeries, widget?.stackColumnPercentage]);

  useEffect(() => {
    if (!widget) return;
    if (editable) {
      setShowUpdateNotice(false);
      return;
    }
    let resolvedData: unknown = null;

    if (widget.type === "excel-table") {
      if (excelTableData.length === 0) return;
      resolvedData = excelTableData;
    } else if (isTextQuery) {
      if (!textQueryLoaded) return;
      resolvedData = textQueryValue;
    } else if (widget.type === "bar") {
      if (barValues.length === 0) return;
      resolvedData = barValues;
    } else if (isBarMarkers) {
      if (barValues.length === 0 || barMarkerValues.length !== barValues.length) return;
      resolvedData = [barValues, barMarkerValues];
    } else if (isStackBar || isLineChart || isColumnChart || isRadarChart) {
      if (stackValues.length === 0) return;
      resolvedData = stackValues;
    } else if (isSimplePie) {
      if (pieValues.length === 0) return;
      resolvedData = pieValues;
    }

    const signature = JSON.stringify({ data: resolvedData, widget });
    if (updateSignatureRef.current === null) {
      updateSignatureRef.current = signature;
      if (!mountedDuringEditRef.current) return;
      mountedDuringEditRef.current = false;
    } else {
      if (updateSignatureRef.current === signature) return;
      updateSignatureRef.current = signature;
    }

    setUpdateNoticeKey((current) => current + 1);
    setShowUpdateNotice(true);
    const hideTimer = window.setTimeout(() => setShowUpdateNotice(false), 4000);
    return () => window.clearTimeout(hideTimer);
  }, [barMarkerValues, barValues, editable, excelTableData, isBarMarkers, isColumnChart, isLineChart, isRadarChart, isSimplePie, isStackBar, isTextQuery, pieValues, stackValues, textQueryLoaded, textQueryValue, widget]);

  function startWidgetResize(direction: "left" | "right" | "top" | "bottom", event: PointerEvent<HTMLButtonElement>) {
    if (!widget) return;
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startY = event.clientY;
    const widgetId = widget.id;
    const startSlot = widget.slot;
    const startWidth = footprint.width;
    const startHeight = footprint.height;
    const widgetRect = event.currentTarget.parentElement?.getBoundingClientRect();
    const slotWidth = (widgetRect?.width ?? startWidth * 20) / startWidth;
    const slotHeight = (widgetRect?.height ?? startHeight * 20) / startHeight;

    function handleMove(moveEvent: globalThis.PointerEvent) {
      const horizontalDelta = Math.round((moveEvent.clientX - startX) / slotWidth);
      const verticalDelta = Math.round((moveEvent.clientY - startY) / slotHeight);
      if (direction === "right") {
        onResizeWidget(boxId, widgetId, startSlot, Math.max(1, startWidth + horizontalDelta), startHeight);
      } else if (direction === "left") {
        const nextWidth = Math.max(1, startWidth - horizontalDelta);
        const appliedDelta = startWidth - nextWidth;
        onResizeWidget(boxId, widgetId, startSlot + appliedDelta, nextWidth, startHeight);
      } else if (direction === "bottom") {
        onResizeWidget(boxId, widgetId, startSlot, startWidth, Math.max(1, startHeight + verticalDelta));
      } else {
        const nextHeight = Math.max(1, startHeight - verticalDelta);
        const appliedDelta = startHeight - nextHeight;
        onResizeWidget(boxId, widgetId, startSlot + appliedDelta * slotColumns, startWidth, nextHeight);
      }
    }

    function handleUp() {
      window.removeEventListener("pointermove", handleMove);
      window.removeEventListener("pointerup", handleUp);
    }

    window.addEventListener("pointermove", handleMove);
    window.addEventListener("pointerup", handleUp);
  }

  function startTableColumnResize(columnIndex: number, event: PointerEvent<HTMLButtonElement>) {
    if (!widget || columnIndex >= excelTableColumnCount - 1) return;
    event.preventDefault();
    event.stopPropagation();
    setExcelCellTooltip(null);
    const widgetId = widget.id;
    const table = event.currentTarget.closest("table") ?? event.currentTarget.parentElement?.querySelector("table");
    const viewportWidth = table?.parentElement?.getBoundingClientRect().width ?? 1;
    const startX = event.clientX;
    const startWidths = [...effectiveTableColumnWidths];
    let nextWidths = startWidths;

    function handleMove(moveEvent: globalThis.PointerEvent) {
      const delta = ((moveEvent.clientX - startX) / viewportWidth) * 100;
      nextWidths = [...startWidths];
      nextWidths[columnIndex] = Math.max(4, startWidths[columnIndex] + delta);
      setTableColumnWidthsDraft(nextWidths);
    }

    function handleUp() {
      onUpdateWidget(boxId, widgetId, { tableColumnWidths: nextWidths.map((width) => Number(width.toFixed(3))) });
      window.removeEventListener("pointermove", handleMove);
      window.removeEventListener("pointerup", handleUp);
    }

    window.addEventListener("pointermove", handleMove);
    window.addEventListener("pointerup", handleUp);
  }

  function startTableRowResize(rowIndex: number, event: PointerEvent<HTMLButtonElement>) {
    if (!widget || rowIndex >= excelTableRowCount) return;
    event.preventDefault();
    event.stopPropagation();
    setExcelCellTooltip(null);
    const widgetId = widget.id;
    const table = event.currentTarget.closest("table") ?? event.currentTarget.parentElement?.querySelector("table");
    const viewportHeight = table?.parentElement?.getBoundingClientRect().height ?? 1;
    const startY = event.clientY;
    const startHeights = [...effectiveTableRowHeights];
    let nextHeights = startHeights;

    function handleMove(moveEvent: globalThis.PointerEvent) {
      const delta = ((moveEvent.clientY - startY) / viewportHeight) * 100;
      nextHeights = [...startHeights];
      nextHeights[rowIndex] = Math.max(4, startHeights[rowIndex] + delta);
      setTableRowHeightsDraft(nextHeights);
    }

    function handleUp() {
      onUpdateWidget(boxId, widgetId, { tableRowHeights: nextHeights.map((height) => Number(height.toFixed(3))) });
      window.removeEventListener("pointermove", handleMove);
      window.removeEventListener("pointerup", handleUp);
    }

    window.addEventListener("pointermove", handleMove);
    window.addEventListener("pointerup", handleUp);
  }

  function updateSelectedExcelCellStyle(changes: { textColor?: string; backgroundColor?: string }) {
    if (!widget || !selectedExcelTableCell) return;
    const reference = selectedExcelTableCell.reference;
    onUpdateWidget(boxId, widget.id, {
      tableCellStyles: {
        ...(widget.tableCellStyles ?? {}),
        [reference]: { ...(widget.tableCellStyles?.[reference] ?? {}), ...changes },
      },
    });
  }

  function placeToolbarBesideWidget(element: HTMLElement) {
    const rect = element.getBoundingClientRect();
    const toolbarWidth = widget?.type === "bar" || isBarMarkers || isStackBar || isLineChart || isColumnChart || isRadarChart || isSimplePie ? 384 : widget?.type === "excel-table" || isTextQuery || widget?.type === "icon" || widget?.type === "gradient-color" ? 320 : 288;
    const toolbarHeight = widget?.type === "bar" || isBarMarkers || isStackBar || isLineChart || isColumnChart || isRadarChart || isSimplePie ? Math.min(560, window.innerHeight - 96) : widget?.type === "excel-table" || isTextQuery ? 430 : widget?.type === "icon" ? 410 : widget?.type === "gradient-color" ? 390 : 260;
    const preferredX = rect.right + 8;
    const fallbackX = rect.left - toolbarWidth - 8;

    setToolbarPosition({
      x: Math.max(8, Math.min(window.innerWidth - toolbarWidth - 8, preferredX + toolbarWidth <= window.innerWidth ? preferredX : fallbackX)),
      y: Math.max(88, Math.min(window.innerHeight - toolbarHeight - 8, rect.top)),
    });
  }

  function startToolbarDrag(event: PointerEvent<HTMLButtonElement>) {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startY = event.clientY;
    const startPosition = toolbarPosition;
    const toolbarWidth = widget?.type === "bar" || isBarMarkers || isStackBar || isLineChart || isColumnChart || isRadarChart || isSimplePie ? 384 : widget?.type === "excel-table" || isTextQuery || widget?.type === "icon" || widget?.type === "gradient-color" ? 320 : 288;
    const toolbarHeight = widget?.type === "bar" || isBarMarkers || isStackBar || isLineChart || isColumnChart || isRadarChart || isSimplePie ? Math.min(560, window.innerHeight - 96) : widget?.type === "excel-table" || isTextQuery ? 430 : widget?.type === "icon" ? 410 : widget?.type === "gradient-color" ? 390 : 260;

    function handleMove(moveEvent: globalThis.PointerEvent) {
      setToolbarPosition({
        x: Math.max(8, Math.min(window.innerWidth - toolbarWidth - 8, startPosition.x + moveEvent.clientX - startX)),
        y: Math.max(88, Math.min(window.innerHeight - toolbarHeight - 8, startPosition.y + moveEvent.clientY - startY)),
      });
    }

    function handleUp() {
      window.removeEventListener("pointermove", handleMove);
      window.removeEventListener("pointerup", handleUp);
    }

    window.addEventListener("pointermove", handleMove);
    window.addEventListener("pointerup", handleUp);
  }

  function applyExcelRange() {
    if (!widget || !widget.sourceUploadId || !widget.sheetName || !cellRangeDraft.trim()) {
      setExcelTableStatus("Select a file, sheet, and range");
      return;
    }
    onUpdateWidget(boxId, widget.id, { cellRange: cellRangeDraft.trim().toUpperCase() });
  }

  function showExcelCellTooltip(rowOffset: number, columnOffset: number, element: HTMLElement) {
    const reference = getExcelCellReference(widget?.cellRange, rowOffset, columnOffset);
    if (!reference) return;
    const rect = element.getBoundingClientRect();
    const below = rect.top < 46;
    setExcelCellTooltip({
      reference,
      x: Math.max(28, Math.min(window.innerWidth - 28, rect.left + rect.width / 2)),
      y: below ? rect.bottom + 7 : rect.top - 7,
      below,
    });
  }

  function updateBarItem(itemId: string, changes: Partial<(typeof currentBarItems)[number]>) {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      barItems: currentBarItems.map((item) => item.id === itemId ? { ...item, ...changes } : item),
    });
  }

  function addBarItem() {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      barItems: [...currentBarItems, { id: makeId("bar-item"), label: `Bar ${currentBarItems.length + 1}`, cell: "A1", markerCell: widget.type === "bar" ? "100" : "A1" }],
    });
  }

  function removeBarItem(itemId: string) {
    if (!widget || currentBarItems.length <= 1) return;
    onUpdateWidget(boxId, widget.id, { barItems: currentBarItems.filter((item) => item.id !== itemId) });
  }

  function reorderBarItems(event: DragEndEvent) {
    if (!widget || !event.over || event.active.id === event.over.id) return;
    const oldIndex = currentBarItems.findIndex((item) => item.id === event.active.id);
    const newIndex = currentBarItems.findIndex((item) => item.id === event.over?.id);
    if (oldIndex < 0 || newIndex < 0) return;
    onUpdateWidget(boxId, widget.id, { barItems: arrayMove(currentBarItems, oldIndex, newIndex) });
  }

  function updateStackCategory(categoryId: string, label: string) {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, { stackCategories: currentStackCategories.map((item) => item.id === categoryId ? { ...item, label } : item) });
  }

  function addStackCategory() {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      stackCategories: [...currentStackCategories, { id: makeId("stack-category"), label: `Bar ${currentStackCategories.length + 1}` }],
      stackSeries: currentStackSeries.map((series) => ({ ...series, cells: [...series.cells, "A1"] })),
    });
  }

  function removeStackCategory(categoryIndex: number) {
    if (!widget || currentStackCategories.length <= 1) return;
    onUpdateWidget(boxId, widget.id, {
      stackCategories: currentStackCategories.filter((_, index) => index !== categoryIndex),
      stackSeries: currentStackSeries.map((series) => ({ ...series, cells: series.cells.filter((_, index) => index !== categoryIndex) })),
    });
  }

  function updateStackSeries(seriesId: string, changes: Partial<(typeof currentStackSeries)[number]>) {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, { stackSeries: currentStackSeries.map((series) => series.id === seriesId ? { ...series, ...changes } : series) });
  }

  function updateSeriesCell(series: (typeof currentStackSeries)[number], categoryIndex: number, value: string) {
    const normalizedValue = value.toUpperCase();
    const cells = currentStackCategories.map((_, index) => series.cells[index] ?? "A1");
    cells[categoryIndex] = normalizedValue;
    updateStackSeries(series.id, { cells });

    if (widget?.type !== "basic-line" || categoryIndex !== 0) {
      setLineCellSuggestions((current) => {
        if (!current[series.id]) return current;
        const next = { ...current };
        delete next[series.id];
        return next;
      });
      return;
    }

    const suggestions = buildRunningCellReferences(normalizedValue, currentStackCategories.length);
    setLineCellSuggestions((current) => {
      if (!suggestions || suggestions.length < 2) {
        const next = { ...current };
        delete next[series.id];
        return next;
      }
      return { ...current, [series.id]: suggestions };
    });
  }

  function confirmRunningSeriesCells(seriesId: string) {
    const suggestions = lineCellSuggestions[seriesId];
    if (!suggestions) return;
    updateStackSeries(seriesId, { cells: suggestions });
    setLineCellSuggestions((current) => {
      const next = { ...current };
      delete next[seriesId];
      return next;
    });
  }

  function addStackSeries() {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      stackSeries: [...currentStackSeries, { id: makeId("stack-series"), label: `Series ${currentStackSeries.length + 1}`, cells: currentStackCategories.map(() => "A1") }],
    });
  }

  function removeStackSeries(seriesId: string) {
    if (!widget || currentStackSeries.length <= 1) return;
    onUpdateWidget(boxId, widget.id, { stackSeries: currentStackSeries.filter((series) => series.id !== seriesId) });
  }

  function reorderStackSeries(event: DragEndEvent) {
    if (!widget || !event.over || event.active.id === event.over.id) return;
    const oldIndex = currentStackSeries.findIndex((series) => series.id === event.active.id);
    const newIndex = currentStackSeries.findIndex((series) => series.id === event.over?.id);
    if (oldIndex < 0 || newIndex < 0) return;
    onUpdateWidget(boxId, widget.id, { stackSeries: arrayMove(currentStackSeries, oldIndex, newIndex) });
  }

  function addLineAnnotation() {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      lineAnnotations: [
        ...currentLineAnnotations,
        { id: makeId("line-annotation"), label: `Target ${currentLineAnnotations.length + 1}`, axis: "y", value: "80", color: defaultChartColors[(currentLineAnnotations.length + 1) % defaultChartColors.length] },
      ],
    });
  }

  function updateLineAnnotation(annotationId: string, changes: Partial<(typeof currentLineAnnotations)[number]>) {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      lineAnnotations: currentLineAnnotations.map((annotation) => annotation.id === annotationId ? { ...annotation, ...changes } : annotation),
    });
  }

  function removeLineAnnotation(annotationId: string) {
    if (!widget) return;
    onUpdateWidget(boxId, widget.id, {
      lineAnnotations: currentLineAnnotations.filter((annotation) => annotation.id !== annotationId),
    });
  }

  function updateChartColor(index: number, color: string) {
    if (!widget) return;
    const itemCount = isStackBar || isLineChart || isColumnChart || isRadarChart ? currentStackSeries.length : currentBarItems.length;
    const colors = Array.from({ length: Math.max(itemCount, index + 1) }, (_, colorIndex) => chartColors[colorIndex] ?? defaultChartColors[colorIndex % defaultChartColors.length]);
    colors[index] = color;
    onUpdateWidget(boxId, widget.id, { chartColors: colors });
  }

  function renderChartAppearance() {
    if (!widget) return null;
    const colorLabels = isStackBar || isLineChart || isColumnChart || isRadarChart
      ? currentStackSeries.map((series) => series.label)
      : currentBarItems.map((item) => item.label);

    return (
      <div className="space-y-2 border-t border-slate-200 pt-3">
        <div className="flex items-center justify-between">
          <div><h3 className="text-xs font-black text-cj-navy">Appearance</h3><p className="text-[10px] font-bold text-slate-400">Legend, font, and data colors.</p></div>
          <label className="flex cursor-pointer items-center gap-2 text-[10px] font-black uppercase text-slate-500">
            <input
              checked={chartShowLegend}
              className="h-4 w-4 accent-cj-blue"
              type="checkbox"
              onChange={(event) => onUpdateWidget(boxId, widget.id, { chartShowLegend: event.target.checked })}
            />
            Legend
          </label>
        </div>
        <div className="grid grid-cols-2 gap-2">
          <label className="text-[10px] font-black uppercase text-slate-500">
            Legend position
            <select
              className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-50"
              disabled={!chartShowLegend}
              value={chartLegendPosition}
              onChange={(event) => onUpdateWidget(boxId, widget.id, { chartLegendPosition: event.target.value as "top" | "bottom" })}
            >
              <option value="top">Top</option>
              <option value="bottom">Bottom</option>
            </select>
          </label>
          <label className="text-[10px] font-black uppercase text-slate-500">
            Font size
            <FontSizeInput
              max={96}
              min={6}
              value={chartFontSize}
              onCommit={(value) => onUpdateWidget(boxId, widget.id, { chartFontSize: value ?? 10 })}
            />
          </label>
        </div>
        <div className="grid grid-cols-2 gap-1.5">
          {colorLabels.map((label, index) => (
            <label key={index} className="flex min-w-0 items-center gap-2 rounded border border-slate-200 bg-slate-50 p-1.5 text-[10px] font-bold text-slate-600">
              <input
                aria-label={`${label || `Item ${index + 1}`} color`}
                className="h-7 w-8 shrink-0 cursor-pointer rounded border-0 bg-transparent p-0"
                type="color"
                value={chartColors[index] ?? defaultChartColors[index % defaultChartColors.length]}
                onChange={(event) => updateChartColor(index, event.target.value)}
              />
              <span className="truncate" title={label || `Item ${index + 1}`}>{label || `Item ${index + 1}`}</span>
            </label>
          ))}
        </div>
        {widget.type === "bar" || isBarMarkers || isStackBar || isColumnChart ? (
          <label className="block rounded border border-slate-200 bg-slate-50 p-2 text-[10px] font-black uppercase text-slate-500">
            <span className="flex items-center justify-between">
              <span>Corner radius</span>
              <output className="rounded bg-white px-2 py-0.5 text-xs font-black normal-case text-cj-navy">{chartBarBorderRadius}px</output>
            </span>
            <input
              aria-label="Bar corner radius"
              className="mt-2 h-2 w-full cursor-pointer accent-cj-blue"
              max="16"
              min="0"
              step="1"
              type="range"
              value={chartBarBorderRadius}
              onChange={(event) => onUpdateWidget(boxId, widget.id, { barBorderRadius: Number(event.target.value) })}
            />
            <span className="mt-1 flex justify-between text-[9px] font-bold normal-case text-slate-400">
              <span>Square</span>
              <span>Rounded</span>
            </span>
          </label>
        ) : null}
        {isBarMarkers ? (
          <div className="space-y-2 rounded border border-slate-200 bg-slate-50 p-2">
            <div className="flex items-center justify-between">
              <div><h4 className="text-[10px] font-black uppercase text-cj-navy">Target marker</h4><p className="text-[9px] font-bold text-slate-400">Marker shape and target values.</p></div>
              <label className="flex cursor-pointer items-center gap-2 text-[9px] font-black uppercase text-slate-500">
                <input className="h-4 w-4 accent-cj-blue" type="checkbox" checked={widget.barMarkerShowValue ?? true} onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerShowValue: event.target.checked })} />
                Show target value
              </label>
            </div>
            <div className="rounded border border-slate-200 bg-white p-2">
              <label className="flex items-center justify-between text-[9px] font-black uppercase text-slate-600">
                Target label
                <input
                  className="h-4 w-4 accent-cj-blue"
                  type="checkbox"
                  checked={widget.barMarkerShowLabel ?? true}
                  onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerShowLabel: event.target.checked })}
                />
              </label>
              <div className={`mt-2 grid grid-cols-2 gap-2 ${widget.barMarkerShowLabel ?? true ? "" : "pointer-events-none opacity-45"}`}>
                <label className="text-[8px] font-black uppercase text-slate-500">
                  Label text
                  <input
                    className="mt-1 h-8 w-full rounded border border-slate-200 px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                    maxLength={24}
                    value={widget.barMarkerLabelText ?? "Target"}
                    onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerLabelText: event.target.value })}
                  />
                </label>
                <label className="text-[8px] font-black uppercase text-slate-500">
                  Font size
                  <FontSizeInput
                    max={32}
                    min={6}
                    value={widget.barMarkerLabelFontSize ?? 9}
                    onCommit={(value) => onUpdateWidget(boxId, widget.id, { barMarkerLabelFontSize: value ?? 9 })}
                  />
                </label>
                <label className="text-[8px] font-black uppercase text-slate-500">
                  Text color
                  <input
                    aria-label="Target label text color"
                    className="mt-1 h-8 w-full cursor-pointer rounded border-0 bg-transparent p-0"
                    type="color"
                    value={widget.barMarkerLabelTextColor ?? "#ffffff"}
                    onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerLabelTextColor: event.target.value })}
                  />
                </label>
                <label className="text-[8px] font-black uppercase text-slate-500">
                  Gap {widget.barMarkerLabelOffsetY ?? 4}px
                  <input
                    className="mt-2 w-full accent-cj-blue"
                    type="range"
                    min={0}
                    max={24}
                    value={widget.barMarkerLabelOffsetY ?? 4}
                    onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerLabelOffsetY: Number(event.target.value) })}
                  />
                </label>
              </div>
              <p className="mt-1 text-[8px] font-bold text-slate-400">Background follows Marker color.</p>
            </div>
            <label className="flex items-center justify-between rounded border border-slate-200 bg-white px-2 py-1.5 text-[9px] font-black uppercase text-slate-600">
              <span>
                Display as % of maximum
                <span className="mt-0.5 block text-[8px] font-bold normal-case text-slate-400">Actual and target divided by Maximum Value.</span>
              </span>
              <input
                className="h-4 w-4 accent-cj-blue"
                type="checkbox"
                checked={widget.barMarkerDisplayPercentage ?? false}
                onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerDisplayPercentage: event.target.checked })}
              />
            </label>
            <div className="grid grid-cols-2 gap-2">
              <label className="text-[9px] font-black uppercase text-slate-500">
                Marker color
                <div className="mt-1 flex h-8 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                  <input aria-label="Target marker color" className="h-6 w-8 cursor-pointer border-0 bg-transparent p-0" type="color" value={widget.barMarkerColor ?? "#e42f44"} onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerColor: event.target.value })} />
                  <span className="text-[9px] font-bold uppercase text-slate-500">{widget.barMarkerColor ?? "#e42f44"}</span>
                </div>
              </label>
              <label className="text-[9px] font-black uppercase text-slate-500">
                Target font
                <FontSizeInput
                  disabled={!(widget.barMarkerShowValue ?? true)}
                  max={96}
                  min={6}
                  value={widget.barMarkerFontSize ?? 10}
                  onCommit={(value) => onUpdateWidget(boxId, widget.id, { barMarkerFontSize: value ?? 10 })}
                />
              </label>
              <label className="text-[9px] font-black uppercase text-slate-500">
                Marker height
                <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold text-cj-navy" value={widget.barMarkerHeight ?? 5} onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerHeight: Number(event.target.value) })}>
                  {[2, 3, 4, 5, 6, 8, 10, 12, 16, 20, 24, 32].map((height) => <option key={height} value={height}>{height}px</option>)}
                </select>
              </label>
              <label className="text-[9px] font-black uppercase text-slate-500">
                Marker width
                <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold text-cj-navy" value={widget.barMarkerWidth ?? 3} onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerWidth: Number(event.target.value) })}>
                  {[1, 2, 3, 4, 5, 6, 8, 10, 12, 16, 20].map((width) => <option key={width} value={width}>{width}px</option>)}
                </select>
              </label>
            </div>
            <div className="rounded border border-slate-200 bg-white p-2">
              <label className="flex items-center justify-between text-[9px] font-black uppercase text-slate-600">
                Background zones
                <input
                  className="h-4 w-4 accent-cj-blue"
                  type="checkbox"
                  checked={widget.barMarkerZoneEnabled ?? false}
                  onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerZoneEnabled: event.target.checked })}
                />
              </label>
              <div className={`mt-2 space-y-2 ${widget.barMarkerZoneEnabled ?? false ? "" : "pointer-events-none opacity-45"}`}>
                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[8px] font-black uppercase text-slate-500">
                    Low ends (%)
                    <input
                      className="mt-1 h-8 w-full rounded border border-slate-200 px-2 text-xs font-bold text-cj-navy outline-none focus:border-cj-blue"
                      type="number"
                      min={0}
                      max={markerZoneMidEnd}
                      step={1}
                      value={markerZoneLowEnd}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerZoneLowEnd: Math.max(0, Math.min(markerZoneMidEnd, Number(event.target.value))) })}
                    />
                    <input
                      className="mt-1 w-full accent-cj-blue"
                      type="range"
                      min={0}
                      max={markerZoneMidEnd}
                      value={markerZoneLowEnd}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerZoneLowEnd: Number(event.target.value) })}
                    />
                  </label>
                  <label className="text-[8px] font-black uppercase text-slate-500">
                    Mid ends (%)
                    <input
                      className="mt-1 h-8 w-full rounded border border-slate-200 px-2 text-xs font-bold text-cj-navy outline-none focus:border-cj-blue"
                      type="number"
                      min={markerZoneLowEnd}
                      max={100}
                      step={1}
                      value={markerZoneMidEnd}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerZoneMidEnd: Math.max(markerZoneLowEnd, Math.min(100, Number(event.target.value))) })}
                    />
                    <input
                      className="mt-1 w-full accent-cj-blue"
                      type="range"
                      min={markerZoneLowEnd}
                      max={100}
                      value={markerZoneMidEnd}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerZoneMidEnd: Number(event.target.value) })}
                    />
                  </label>
                </div>
                <div className="grid grid-cols-3 gap-1.5">
                  {([
                    ["Low", "barMarkerZoneLowColor", widget.barMarkerZoneLowColor ?? "#fee2e2"],
                    ["Mid", "barMarkerZoneMidColor", widget.barMarkerZoneMidColor ?? "#fef3c7"],
                    ["High", "barMarkerZoneHighColor", widget.barMarkerZoneHighColor ?? "#dcfce7"],
                  ] as const).map(([label, key, color]) => (
                    <label key={key} className="text-center text-[8px] font-black uppercase text-slate-500">
                      {label}
                      <input
                        aria-label={`${label} zone color`}
                        className="mt-1 h-7 w-full cursor-pointer rounded border-0 bg-transparent p-0"
                        type="color"
                        value={color}
                        onChange={(event) => onUpdateWidget(boxId, widget.id, { [key]: event.target.value })}
                      />
                    </label>
                  ))}
                </div>
                <label className="block text-[8px] font-black uppercase text-slate-500">
                  Zone opacity {markerZoneOpacity}%
                  <input
                    className="mt-1 w-full accent-cj-blue"
                    type="range"
                    min={5}
                    max={80}
                    step={1}
                    value={markerZoneOpacity}
                    onChange={(event) => onUpdateWidget(boxId, widget.id, { barMarkerZoneOpacity: Number(event.target.value) })}
                  />
                </label>
              </div>
            </div>
          </div>
        ) : null}
      </div>
    );
  }

  function renderWidgetToolbarActions() {
    if (!widget) return null;
    return (
      <div className="grid grid-cols-2 gap-2 border-b border-slate-100 pb-3">
        <button
          ref={duplicateDraggable.setNodeRef}
          className="flex h-9 touch-none cursor-grab items-center justify-center gap-2 rounded border border-cj-blue/30 bg-blue-50 text-[10px] font-black uppercase text-cj-blue hover:border-cj-blue active:cursor-grabbing"
          title="Drag to duplicate widget"
          type="button"
          {...duplicateDraggable.attributes}
          {...duplicateDraggable.listeners}
        >
          <Copy className="h-4 w-4" />
          Duplicate
        </button>
        <button
          className="flex h-9 items-center justify-center gap-2 rounded border border-red-200 bg-red-50 text-[10px] font-black uppercase text-red-600 hover:border-red-400 hover:bg-red-100"
          title="Delete widget"
          type="button"
          onClick={() => onRemoveWidget(boxId, widget.id)}
        >
          <Trash2 className="h-4 w-4" />
          Delete
        </button>
      </div>
    );
  }

  function handleToolbarFieldNavigation(event: ReactKeyboardEvent<HTMLDivElement>) {
    const currentField = event.target as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement;
    if (!currentField.matches("input, select, textarea")) return;

    const isTextEntry = currentField.matches('input:not([type="checkbox"]):not([type="color"]):not([type="range"]), textarea');
    if (event.key === "Tab") return;
    const isTabNavigation = false;
    const isArrowNavigation = isTextEntry && (event.key === "ArrowDown" || event.key === "ArrowUp");
    if (!isTabNavigation && !isArrowNavigation) return;

    const toolbar = event.currentTarget;
    const fields = Array.from(toolbar.querySelectorAll<HTMLElement>(
      'input:not([type="checkbox"]):not([type="color"]):not([type="range"]):not(:disabled), select:not(:disabled), textarea:not(:disabled)',
    )).filter((field) => field.offsetParent !== null);
    const currentIndex = fields.indexOf(currentField);
    if (currentIndex < 0 || fields.length < 2) return;

    const moveBackward = event.shiftKey || event.key === "ArrowUp";
    const nextIndex = (currentIndex + (moveBackward ? -1 : 1) + fields.length) % fields.length;
    event.preventDefault();
    fields[nextIndex].focus();
    if (fields[nextIndex] instanceof HTMLInputElement && fields[nextIndex].type === "text") fields[nextIndex].select();
  }

  return (
    <div
      ref={setNodeRef}
      style={slotStyle}
      className={`relative grid min-h-0 place-items-center rounded text-center hover:!z-[55] ${
        widget
          ? editable ? "border border-dashed border-slate-200 bg-white" : "border border-dashed border-transparent bg-white"
          : "border border-transparent bg-transparent"
      }`}
    >
      {widget ? (
        <div
          ref={draggable.setNodeRef}
          data-workspace-widget={widget.id}
          data-widget-selected={editable && isSelected ? "true" : undefined}
          data-widget-custom-background={widget.type === "title" || widget.type === "text" || isTextQuery ? String(Boolean(widget.useBackgroundColor)) : undefined}
          className={`group relative grid h-full min-h-0 w-full place-items-center overflow-visible rounded bg-white p-0.5 ${
            editable ? isSelected ? "outline outline-2 outline-offset-2 outline-cj-blue shadow-panel" : "shadow-sm" : ""
          } ${showUpdateNotice && !editable ? "animate-widget-content-update" : ""}`}
          style={{
            "--widget-chart-font-size": `${chartFontSize}px`,
            "--widget-background-color": widget.backgroundColor ?? "#ffffff",
            ...(widget.type === "title" || widget.type === "text" || isTextQuery
              ? { backgroundColor: widget.useBackgroundColor ? (widget.backgroundColor ?? "#ffffff") : undefined }
              : {}),
          } as CSSProperties}
          {...draggable.attributes}
          onClick={(event) => {
            event.stopPropagation();
            if (editable) {
              if (!isSelected && (widget.type === "title" || widget.type === "text" || isTextQuery || widget.type === "icon" || widget.type === "gradient-color" || isLargeDataWidget(widget.type))) placeToolbarBesideWidget(event.currentTarget);
              setIsSelected(true);
              setIsEditing(widget.type === "title" || widget.type === "text");
            }
          }}
        >
          {editable ? (
            <>
              <button
                className={`absolute z-20 grid h-5 w-5 cursor-grab place-items-center rounded bg-white text-slate-500 shadow-sm ${widget.type === "icon" ? "-left-2 -top-2" : "left-1 top-1"}`}
                title="Move widget"
                type="button"
                {...draggable.listeners}
                onClick={(event) => event.stopPropagation()}
              >
                <GripVertical className="h-3 w-3" />
              </button>
              <button
                aria-label="Expand widget left"
                className="absolute bottom-0 left-0 top-0 z-20 w-1.5 cursor-ew-resize bg-transparent hover:bg-cj-blue/30"
                type="button"
                onPointerDown={(event) => startWidgetResize("left", event)}
              />
              <button
                aria-label="Expand widget right"
                className="absolute bottom-0 right-0 top-0 z-20 w-1.5 cursor-ew-resize bg-transparent hover:bg-cj-blue/30"
                type="button"
                onPointerDown={(event) => startWidgetResize("right", event)}
              />
              <button
                aria-label="Expand widget top"
                className="absolute left-0 right-0 top-0 z-20 h-1.5 cursor-ns-resize bg-transparent hover:bg-cj-blue/30"
                type="button"
                onPointerDown={(event) => startWidgetResize("top", event)}
              />
              <button
                aria-label="Expand widget bottom"
                className="absolute bottom-0 left-0 right-0 z-20 h-1.5 cursor-ns-resize bg-transparent hover:bg-cj-blue/30"
                type="button"
                onPointerDown={(event) => startWidgetResize("bottom", event)}
              />
            </>
          ) : null}

          {editable && isEditing && (widget.type === "title" || widget.type === "text") ? (
            <div className="absolute inset-1 z-30 flex items-center gap-1 rounded bg-white px-1 shadow-sm" onClick={(event) => event.stopPropagation()}>
              <input
                autoFocus
                data-widget-custom-text-color={widget.textColor && widget.textColor.toLowerCase() !== "#122033" ? "true" : undefined}
                className="min-w-0 flex-1 border-0 bg-transparent px-1 text-cj-navy outline-none"
                style={{ ...getWidgetTextStyle(widget), "--widget-text-color": widget.textColor ?? "#122033" } as CSSProperties}
                value={widget.content ?? widget.label}
                onChange={(event) => onUpdateWidget(boxId, widget.id, { content: event.target.value })}
                onKeyDown={(event) => event.key === "Enter" && setIsEditing(false)}
              />
              <button className="h-7 px-1 text-xs font-black text-cj-blue" type="button" onClick={() => setIsEditing(false)}>Done</button>
            </div>
          ) : widget.type === "excel-table" ? (
            <div className="absolute inset-1 overflow-hidden rounded bg-white">
              {excelTableData.length > 0 ? (
                <>
                <table
                  className="h-full w-full table-fixed border-collapse leading-tight text-slate-700"
                  style={{
                    fontFamily: widget.tableFontFamily ?? "Inter",
                    fontSize: `calc(${widget.tableFontSize ?? excelTableFontSize}px * var(--wms-display-scale, 1))`,
                    height: `${tableHeightUnits}%`,
                    width: `${tableWidthUnits}%`,
                  }}
                >
                  <colgroup>
                    {effectiveTableColumnWidths.map((width, columnIndex) => <col key={columnIndex} style={{ width: `${(width / tableWidthUnits) * 100}%` }} />)}
                  </colgroup>
                  <thead className="relative z-30" style={{ height: `${(effectiveTableRowHeights[0] / tableHeightUnits) * 100}%` }}>
                    <tr style={{ height: `${(effectiveTableRowHeights[0] / tableHeightUnits) * 100}%` }}>
                      {excelTableData[0].map((value, columnIndex) => {
                        const reference = getExcelCellReference(widget.cellRange, 0, columnIndex);
                        const cellStyle = widget.tableCellStyles?.[reference];
                        const selected = editable && selectedExcelTableCell?.reference === reference;
                        return (
                        <th
                          key={columnIndex}
                          data-excel-custom-text-color={cellStyle?.textColor ? "true" : undefined}
                          className={`${editable ? "cursor-help" : ""} relative z-10 overflow-visible break-words border border-slate-300 bg-slate-100 font-black text-cj-navy ${selected ? "outline outline-2 outline-inset outline-cj-blue" : ""}`}
                          style={{ ...excelTableCellStyle, backgroundColor: cellStyle?.backgroundColor, color: cellStyle?.textColor, "--excel-cell-text-color": cellStyle?.textColor } as CSSProperties}
                          onClick={editable ? (event) => { event.stopPropagation(); setSelectedExcelTableCell({ row: 0, column: columnIndex, reference }); } : undefined}
                          onPointerEnter={editable ? (event) => showExcelCellTooltip(0, columnIndex, event.currentTarget) : undefined}
                          onPointerLeave={editable ? () => setExcelCellTooltip(null) : undefined}
                        >
                          <span className="line-clamp-3 overflow-hidden">{formatExcelTableValue(value, excelTableFormats[0]?.[columnIndex])}</span>
                          {editable && columnIndex < excelTableColumnCount - 1 ? (
                            <button
                              aria-label={`Resize column ${columnIndex + 1}`}
                              className="absolute -right-1 top-0 z-20 h-full w-2 touch-none cursor-col-resize bg-transparent hover:bg-cj-blue/40"
                              title="Drag to resize column"
                              type="button"
                              onPointerDown={(event) => startTableColumnResize(columnIndex, event)}
                            />
                          ) : null}
                        </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {excelTableData.slice(1).map((row, rowIndex) => (
                      <tr key={rowIndex} style={{ height: `${(effectiveTableRowHeights[rowIndex + 1] / tableHeightUnits) * 100}%` }}>
                        {row.map((value, columnIndex) => {
                          const tableRowIndex = rowIndex + 1;
                          const reference = getExcelCellReference(widget.cellRange, tableRowIndex, columnIndex);
                          const cellStyle = widget.tableCellStyles?.[reference];
                          const selected = editable && selectedExcelTableCell?.reference === reference;
                          return (
                          <td
                            key={columnIndex}
                            data-excel-custom-text-color={cellStyle?.textColor ? "true" : undefined}
                            className={`${editable ? "cursor-help" : ""} relative overflow-visible break-words border border-slate-200 ${selected ? "z-10 outline outline-2 outline-inset outline-cj-blue" : ""}`}
                            style={{ ...excelTableCellStyle, backgroundColor: cellStyle?.backgroundColor, color: cellStyle?.textColor, "--excel-cell-text-color": cellStyle?.textColor } as CSSProperties}
                            onClick={editable ? (event) => { event.stopPropagation(); setSelectedExcelTableCell({ row: tableRowIndex, column: columnIndex, reference }); } : undefined}
                            onPointerEnter={editable ? (event) => showExcelCellTooltip(tableRowIndex, columnIndex, event.currentTarget) : undefined}
                            onPointerLeave={editable ? () => setExcelCellTooltip(null) : undefined}
                          >
                            <span className="line-clamp-3 overflow-hidden">{formatExcelTableValue(value, excelTableFormats[tableRowIndex]?.[columnIndex])}</span>
                            {editable ? <button aria-label={`Resize row ${tableRowIndex + 1}`} className="absolute -bottom-1 left-0 z-20 h-2 w-full touch-none cursor-row-resize bg-transparent hover:bg-cj-blue/40" title="Drag to resize row" type="button" onPointerDown={(event) => startTableRowResize(tableRowIndex, event)} /> : null}
                          </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {editable ? (
                  <button
                    aria-label="Resize Header row"
                    className="absolute left-0 right-0 z-40 h-3 -translate-y-1/2 touch-none cursor-row-resize bg-transparent hover:bg-cj-blue/50"
                    style={{ top: `${effectiveTableRowHeights[0]}%` }}
                    title="Drag to resize Header row"
                    type="button"
                    onPointerDown={(event) => startTableRowResize(0, event)}
                  />
                ) : null}
                </>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{excelTableStatus}</div>
              )}
            </div>
          ) : widget.type === "bar" ? (
            <div className="absolute inset-1 overflow-visible rounded bg-white">
              {renderBarLegend()}
              <div className="workspace-chart absolute inset-x-0 min-h-0 overflow-visible" style={barChartViewportStyle}>
              {barValues.length > 0 ? (
                renderBasicBarChart()
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{barStatus}</div>
              )}
              </div>
            </div>
          ) : isBarMarkers ? (
            <div className="absolute inset-1 overflow-visible rounded bg-white">
              {renderBarLegend()}
              <div className="absolute inset-x-0 min-h-0 overflow-hidden" style={barChartViewportStyle}>
                {widget.barMarkerZoneEnabled ?? false ? (
                  <div
                    className="pointer-events-none absolute bottom-7 right-2 top-2 z-0 rounded-sm"
                    style={{ background: markerZoneBackground, left: "clamp(52px, 22%, 96px)", opacity: markerZoneOpacity / 100 }}
                  />
                ) : null}
                <div className="workspace-chart absolute inset-0 z-10 overflow-visible">
              {barValues.length > 0 && barMarkerValues.length === barValues.length ? (
                <Suspense fallback={<div className="grid h-full place-items-center text-[10px] font-bold text-slate-400">Loading chart...</div>}>
                  <ApexBarChart
                    options={barChartOptions}
                    markerData={currentBarItems.map((item, index) => ({
                      x: item.label || item.cell,
                      y: barValues[index] ?? 0,
                      goals: [{
                        name: "Target",
                        value: barMarkerValues[index] ?? 0,
                        strokeColor: widget.barMarkerColor ?? "#e42f44",
                        strokeHeight: widget.barMarkerHeight ?? 5,
                        strokeWidth: widget.barMarkerWidth ?? 3,
                      }],
                    }))}
                  />
                </Suspense>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{barStatus}</div>
              )}
                </div>
                {barValues.length > 0 && barMarkerValues.length === barValues.length ? renderBarMarkerLabels() : null}
              </div>
            </div>
          ) : isStackBar ? (
            <div className="workspace-chart absolute inset-1 overflow-visible rounded bg-white">
              {stackValues.length > 0 ? (
                <Suspense fallback={<div className="grid h-full place-items-center text-[10px] font-bold text-slate-400">Loading chart...</div>}>
                  <ApexBarChart
                    options={stackChartOptions}
                    series={currentStackSeries.map((series, index) => ({ name: series.label, data: stackValues[index] ?? [] }))}
                  />
                </Suspense>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{barStatus}</div>
              )}
            </div>
          ) : isLineChart ? (
            <div className="workspace-chart absolute inset-1 overflow-visible rounded bg-white">
              {stackValues.length > 0 ? (
                <Suspense fallback={<div className="grid h-full place-items-center text-[10px] font-bold text-slate-400">Loading chart...</div>}>
                  <ApexLineChart
                    options={lineChartOptions}
                    series={currentStackSeries.map((series, index) => ({ name: series.label, data: stackValues[index] ?? [] }))}
                  />
                </Suspense>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{barStatus}</div>
              )}
            </div>
          ) : isColumnChart ? (
            <div className="column-rotated-chart workspace-chart absolute inset-1 overflow-visible rounded bg-white">
              {stackValues.length > 0 ? (
                <Suspense fallback={<div className="grid h-full place-items-center text-[10px] font-bold text-slate-400">Loading chart...</div>}>
                  <ApexBarChart
                    options={columnChartOptions}
                    series={currentStackSeries.map((series, index) => ({ name: series.label, data: stackValues[index] ?? [] }))}
                  />
                </Suspense>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{barStatus}</div>
              )}
            </div>
          ) : isRadarChart ? (
            <div className="workspace-chart absolute inset-0 overflow-visible rounded bg-white">
              {canRenderRadar ? (
                <Suspense fallback={<div className="grid h-full place-items-center text-[10px] font-bold text-slate-400">Loading chart...</div>}>
                  <ApexRadarChart
                    options={radarChartOptions}
                    series={radarSeries}
                  />
                </Suspense>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">
                  {radarCategoryCount < 3 ? "Radar requires at least 3 categories" : barStatus}
                </div>
              )}
            </div>
          ) : isSimplePie ? (
            <div className={`workspace-chart absolute inset-1 overflow-visible rounded bg-white ${widget.type === "simple-pie" && widget.pieShowValueCallouts ? "pie-value-callout-chart" : ""} ${widget.type === "simple-donut" ? "rounded-spaced-donut-chart" : ""}`}>
              {pieValues.some((value) => value > 0) ? (
                <Suspense fallback={<div className="grid h-full place-items-center text-[10px] font-bold text-slate-400">Loading chart...</div>}>
                  <ApexPieChart options={pieChartOptions} data={pieValues} type={widget.type === "simple-donut" ? "donut" : "pie"} />
                </Suspense>
              ) : (
                <div className="grid h-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{barStatus}</div>
              )}
            </div>
          ) : isTextQuery ? (
            textQueryLoaded ? <AdaptiveWidgetText content={textQueryValue} widget={widget} /> : (
              <div className="grid h-full w-full place-items-center px-2 text-center text-[10px] font-bold text-slate-400">{textQueryStatus}</div>
            )
          ) : widget.type === "gradient-color" ? (
            <div
              className="absolute inset-0"
              style={{
                background: gradientCss,
                borderRadius: `${widget.gradientBorderRadius ?? 8}px`,
                opacity: (widget.gradientOpacity ?? 100) / 100,
              }}
            />
          ) : widget.type === "icon" ? (
            <div className="absolute inset-0 grid place-items-center overflow-hidden">
              <SelectedIcon
                aria-label={iconOptions.find((option) => option.name === widget.iconName)?.label ?? "Package"}
                className="block shrink-0"
                role="img"
                style={{
                  color: widget.iconColor ?? "#1473e6",
                  height: `calc(${widget.iconSize ?? 32}px * var(--wms-display-scale, 1))`,
                  maxHeight: "90%",
                  maxWidth: "90%",
                  width: `calc(${widget.iconSize ?? 32}px * var(--wms-display-scale, 1))`,
                }}
                strokeWidth={2}
              />
            </div>
          ) : (
            <AdaptiveWidgetText widget={widget} />
          )}

          {showUpdateNotice ? (
            <div key={updateNoticeKey} className="pointer-events-none absolute inset-0 z-[60] grid place-items-center overflow-hidden rounded">
              <span className="animate-widget-update-glow absolute inset-0 rounded border-2 border-cj-blue bg-cj-blue/10" />
              <span className="animate-widget-update-notice relative rounded-md border border-cj-blue/30 bg-white/95 px-3 py-1.5 text-xs font-black text-cj-blue shadow-panel">
                New update!
              </span>
            </div>
          ) : null}

          {editable && isSelected ? createPortal((
            <div
              aria-hidden="true"
              className="pointer-events-none fixed inset-0 z-[75] bg-slate-950/15 transition-opacity duration-200 dark:bg-black/25"
              data-edit-focus-backdrop
            />
          ), document.body) : null}

          {editable && excelCellTooltip ? createPortal((
            <div
              data-excel-cell-tooltip
              className="pointer-events-none fixed z-[120] rounded border border-slate-200 bg-white px-2 py-1 font-mono text-[11px] font-black text-cj-navy shadow-panel"
              style={{
                left: excelCellTooltip.x,
                top: excelCellTooltip.y,
                transform: excelCellTooltip.below ? "translateX(-50%)" : "translate(-50%, -100%)",
              }}
            >
              {excelCellTooltip.reference}
            </div>
          ), document.body) : null}

          {editable && isSelected && (widget.type === "title" || widget.type === "text" || isTextQuery) ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className={`fixed z-[100] max-h-[calc(100vh-6rem)] overflow-y-auto ${isTextQuery ? "w-80" : "w-72"}`}
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  {widget.type === "title" ? "Title" : isTextQuery ? "Text Query" : "Text"} tools
                </button>
                {renderWidgetToolbarActions()}
                {isTextQuery ? (
                  <div className="space-y-2 rounded border border-slate-200 bg-slate-50 p-2">
                    <label className="block text-[9px] font-black uppercase text-slate-500">
                      Excel file
                      <select
                        className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                        value={widget.sourceUploadId ?? ""}
                        onChange={(event) => {
                          const upload = excelUploads.find((item) => item.id === event.target.value);
                          onUpdateWidget(boxId, widget.id, {
                            sourceUploadId: upload?.id ?? "",
                            sourceFilename: upload?.filename ?? "",
                            sheetName: "",
                          });
                        }}
                      >
                        <option value="">Select uploaded file</option>
                        {excelUploads.map((upload) => <option key={upload.id} value={upload.id}>{upload.filename}</option>)}
                      </select>
                    </label>
                    <div className="grid grid-cols-2 gap-2">
                      <label className="text-[9px] font-black uppercase text-slate-500">
                        Sheet
                        <select
                          className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-100"
                          disabled={!widget.sourceUploadId || excelSheets.length === 0}
                          value={widget.sheetName ?? ""}
                          onChange={(event) => onUpdateWidget(boxId, widget.id, { sheetName: event.target.value })}
                        >
                          <option value="">Select sheet</option>
                          {excelSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                        </select>
                      </label>
                      <label className="text-[9px] font-black uppercase text-slate-500">
                        Cell / formula
                        <input
                          className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold uppercase text-cj-navy outline-none focus:border-cj-blue"
                          maxLength={120}
                          placeholder="A1 / D7+E7"
                          value={widget.textQueryCell ?? "A1"}
                          onChange={(event) => onUpdateWidget(boxId, widget.id, { textQueryCell: event.target.value.toUpperCase() })}
                        />
                      </label>
                    </div>
                    <p className="truncate text-[9px] font-bold text-slate-400" title={textQueryStatus}>{textQueryStatus}</p>
                  </div>
                ) : null}
                <div className="grid grid-cols-[1fr_84px] gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Font
                    <select
                      aria-label="Font family"
                      className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.fontFamily ?? "Inter"}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { fontFamily: event.target.value })}
                    >
                      {fontFamilyOptions.map((font) => <option key={font} value={font}>{font}</option>)}
                    </select>
                  </label>
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Size
                    <FontSizeInput
                      max={240}
                      min={6}
                      value={widget.fontSize ?? 16}
                      onCommit={(value) => onUpdateWidget(boxId, widget.id, { fontSize: value ?? 16 })}
                    />
                  </label>
                </div>

                <div className="grid grid-cols-[1fr_36px] gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Weight
                    <select
                      aria-label="Font weight"
                      className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.fontWeight ?? 700}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { fontWeight: Number(event.target.value) })}
                    >
                      {fontWeightOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                    </select>
                  </label>
                  <button
                    aria-label="Italic"
                    className={`mt-4 grid h-8 w-9 place-items-center rounded border ${widget.fontStyle === "italic" ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 text-slate-500"}`}
                    title="Italic"
                    type="button"
                    onClick={() => onUpdateWidget(boxId, widget.id, { fontStyle: widget.fontStyle === "italic" ? "normal" : "italic" })}
                  >
                    <Italic className="h-4 w-4" />
                  </button>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Text color
                    <span className="mt-1 flex h-8 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                      <input
                        aria-label="Text color"
                        className="h-5 w-7 cursor-pointer border-0 bg-transparent p-0"
                        type="color"
                        value={widget.textColor ?? "#122033"}
                        onChange={(event) => onUpdateWidget(boxId, widget.id, { textColor: event.target.value })}
                      />
                      <span className="text-[9px] font-bold normal-case text-slate-500">{widget.textColor ?? "#122033"}</span>
                    </span>
                  </label>
                  <div className="text-[10px] font-black uppercase text-slate-500">
                    <label className="flex items-center gap-1.5">
                      <input
                        className="h-3.5 w-3.5 accent-cj-blue"
                        type="checkbox"
                        checked={widget.useBackgroundColor ?? false}
                        onChange={(event) => onUpdateWidget(boxId, widget.id, { useBackgroundColor: event.target.checked })}
                      />
                      Background
                    </label>
                    <span className="mt-1 flex h-8 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                      <input
                        aria-label="Background color"
                        className="h-5 w-7 cursor-pointer border-0 bg-transparent p-0 disabled:cursor-not-allowed disabled:opacity-40"
                        disabled={!(widget.useBackgroundColor ?? false)}
                        type="color"
                        value={widget.backgroundColor ?? "#ffffff"}
                        onChange={(event) => onUpdateWidget(boxId, widget.id, { backgroundColor: event.target.value })}
                      />
                      <span className="text-[9px] font-bold normal-case text-slate-500">{widget.backgroundColor ?? "#ffffff"}</span>
                    </span>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <span className="text-[10px] font-black uppercase text-slate-500">Horizontal</span>
                    <div className="mt-1 flex">
                      {([
                        ["left", AlignLeft],
                        ["center", AlignCenter],
                        ["right", AlignRight],
                      ] as const).map(([alignment, Icon]) => (
                        <button
                          key={alignment}
                          aria-label={`Align ${alignment}`}
                          className={`grid h-8 flex-1 place-items-center border first:rounded-l last:rounded-r ${widget.textAlign === alignment || (!widget.textAlign && alignment === "center") ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 text-slate-500"}`}
                          type="button"
                          onClick={() => onUpdateWidget(boxId, widget.id, { textAlign: alignment })}
                        >
                          <Icon className="h-4 w-4" />
                        </button>
                      ))}
                    </div>
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase text-slate-500">Vertical</span>
                    <div className="mt-1 flex">
                      {([
                        ["top", ArrowUpToLine],
                        ["center", AlignCenter],
                        ["bottom", ArrowDownToLine],
                      ] as const).map(([alignment, Icon]) => (
                        <button
                          key={alignment}
                          aria-label={`Align ${alignment}`}
                          className={`grid h-8 flex-1 place-items-center border first:rounded-l last:rounded-r ${widget.verticalAlign === alignment || (!widget.verticalAlign && alignment === "center") ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 text-slate-500"}`}
                          type="button"
                          onClick={() => onUpdateWidget(boxId, widget.id, { verticalAlign: alignment })}
                        >
                          <Icon className={`h-4 w-4 ${alignment === "center" ? "rotate-90" : ""}`} />
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ), document.body) : null}

          {editable && isSelected && widget.type === "excel-table" ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className="fixed z-[100] max-h-[calc(100vh-6rem)] w-80 overflow-y-auto"
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  Excel Query tools
                </button>
                {renderWidgetToolbarActions()}

                <label className="block text-[10px] font-black uppercase text-slate-500">
                  Excel file
                  <select
                    aria-label="Excel file"
                    className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                    value={widget.sourceUploadId ?? ""}
                    onChange={(event) => {
                      const upload = excelUploads.find((item) => item.id === event.target.value);
                      onUpdateWidget(boxId, widget.id, {
                        sourceUploadId: upload?.id ?? "",
                        sourceFilename: upload?.filename ?? "",
                        sheetName: "",
                        cellRange: "",
                      });
                      setCellRangeDraft("A1:J7");
                      setExcelTableStatus("Select a sheet and range");
                    }}
                  >
                    <option value="">Select uploaded file</option>
                    {excelUploads.map((upload) => <option key={upload.id} value={upload.id}>{upload.filename}</option>)}
                  </select>
                </label>

                <label className="block text-[10px] font-black uppercase text-slate-500">
                  Sheet
                  <select
                    aria-label="Excel sheet"
                    className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-50 disabled:text-slate-400"
                    disabled={!widget.sourceUploadId || excelSheets.length === 0}
                    value={widget.sheetName ?? ""}
                    onChange={(event) => onUpdateWidget(boxId, widget.id, { sheetName: event.target.value, cellRange: "" })}
                  >
                    <option value="">Select sheet</option>
                    {excelSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                  </select>
                </label>

                <div className="grid grid-cols-[1fr_96px] gap-2">
                  <label className="block text-[10px] font-black uppercase text-slate-500">
                    Table font
                    <select
                      aria-label="Table font family"
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.tableFontFamily ?? "Inter"}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { tableFontFamily: event.target.value })}
                    >
                      {fontFamilyOptions.map((font) => <option key={font} value={font}>{font}</option>)}
                    </select>
                  </label>
                  <label className="block text-[10px] font-black uppercase text-slate-500">
                    Size
                    <FontSizeInput
                      allowAuto
                      max={72}
                      min={4}
                      value={widget.tableFontSize}
                      onCommit={(value) => onUpdateWidget(boxId, widget.id, { tableFontSize: value })}
                    />
                  </label>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <span className="text-[10px] font-black uppercase text-slate-500">Horizontal</span>
                    <div className="mt-1 flex">
                      {([
                        ["left", AlignLeft],
                        ["center", AlignCenter],
                        ["right", AlignRight],
                      ] as const).map(([alignment, Icon]) => (
                        <button
                          key={alignment}
                          aria-label={`Align table text ${alignment}`}
                          className={`grid h-8 flex-1 place-items-center border first:rounded-l last:rounded-r ${(widget.tableTextAlign ?? "left") === alignment ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 text-slate-500"}`}
                          type="button"
                          onClick={() => onUpdateWidget(boxId, widget.id, { tableTextAlign: alignment })}
                        >
                          <Icon className="h-4 w-4" />
                        </button>
                      ))}
                    </div>
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase text-slate-500">Vertical</span>
                    <div className="mt-1 flex">
                      {([
                        ["top", ArrowUpToLine],
                        ["middle", AlignCenter],
                        ["bottom", ArrowDownToLine],
                      ] as const).map(([alignment, Icon]) => (
                        <button
                          key={alignment}
                          aria-label={`Align table text ${alignment}`}
                          className={`grid h-8 flex-1 place-items-center border first:rounded-l last:rounded-r ${(widget.tableVerticalAlign ?? "middle") === alignment ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 text-slate-500"}`}
                          type="button"
                          onClick={() => onUpdateWidget(boxId, widget.id, { tableVerticalAlign: alignment })}
                        >
                          <Icon className={`h-4 w-4 ${alignment === "middle" ? "rotate-90" : ""}`} />
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="rounded border border-slate-200 bg-slate-50 p-2">
                  <div className="mb-2 flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <h3 className="text-[10px] font-black uppercase text-slate-600">Selected Cell</h3>
                      <p className="truncate text-[10px] font-bold text-slate-400">{selectedExcelTableCell?.reference ?? "Click a table Cell"}</p>
                    </div>
                    <button
                      className="h-7 rounded border border-slate-200 bg-white px-2 text-[9px] font-black uppercase text-slate-500 hover:border-red-300 hover:text-red-600 disabled:opacity-30"
                      disabled={!selectedExcelTableCell}
                      type="button"
                      onClick={() => {
                        if (!selectedExcelTableCell) return;
                        const nextStyles = { ...(widget.tableCellStyles ?? {}) };
                        delete nextStyles[selectedExcelTableCell.reference];
                        onUpdateWidget(boxId, widget.id, { tableCellStyles: nextStyles });
                      }}
                    >
                      Reset
                    </button>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <label className="text-[9px] font-black uppercase text-slate-500">
                      Text color
                      <span className="mt-1 flex h-8 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                        <input
                          aria-label="Selected Cell text color"
                          className="h-5 w-7 cursor-pointer border-0 bg-transparent p-0 disabled:opacity-30"
                          disabled={!selectedExcelTableCell}
                          type="color"
                          value={selectedExcelTableCell ? (widget.tableCellStyles?.[selectedExcelTableCell.reference]?.textColor ?? "#334155") : "#334155"}
                          onChange={(event) => updateSelectedExcelCellStyle({ textColor: event.target.value })}
                        />
                      </span>
                    </label>
                    <label className="text-[9px] font-black uppercase text-slate-500">
                      Background
                      <span className="mt-1 flex h-8 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                        <input
                          aria-label="Selected Cell background color"
                          className="h-5 w-7 cursor-pointer border-0 bg-transparent p-0 disabled:opacity-30"
                          disabled={!selectedExcelTableCell}
                          type="color"
                          value={selectedExcelTableCell ? (widget.tableCellStyles?.[selectedExcelTableCell.reference]?.backgroundColor ?? "#ffffff") : "#ffffff"}
                          onChange={(event) => updateSelectedExcelCellStyle({ backgroundColor: event.target.value })}
                        />
                      </span>
                    </label>
                  </div>
                </div>

                <label className="block text-[10px] font-black uppercase text-slate-500">
                  Cell range
                  <div className="mt-1 flex gap-2">
                    <input
                      className="h-9 min-w-0 flex-1 rounded border border-slate-200 px-3 text-sm font-bold uppercase text-cj-navy outline-none focus:border-cj-blue"
                      placeholder="A1:J7"
                      value={cellRangeDraft}
                      onChange={(event) => setCellRangeDraft(event.target.value)}
                      onKeyDown={(event) => event.key === "Enter" && applyExcelRange()}
                    />
                    <button
                      className="h-9 rounded bg-cj-blue px-4 text-xs font-black text-white transition hover:bg-cj-navy disabled:cursor-not-allowed disabled:opacity-40"
                      disabled={!widget.sourceUploadId || !widget.sheetName || !cellRangeDraft.trim()}
                      type="button"
                      onClick={applyExcelRange}
                    >
                      Apply
                    </button>
                  </div>
                </label>
                <p className="truncate text-[10px] font-bold text-slate-400" title={excelTableStatus}>{excelTableStatus}</p>
              </div>
            </div>
          ), document.body) : null}

          {editable && isSelected && (widget.type === "bar" || isBarMarkers) ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className="fixed z-[100] max-h-[calc(100vh-6rem)] w-96 overflow-y-auto"
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  {isBarMarkers ? "Bar with Markers" : "Basic Bar"} tools
                </button>
                {renderWidgetToolbarActions()}

                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Excel file
                    <select
                      aria-label="Bar Excel file"
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.sourceUploadId ?? ""}
                      onChange={(event) => {
                        const upload = excelUploads.find((item) => item.id === event.target.value);
                        onUpdateWidget(boxId, widget.id, {
                          sourceUploadId: upload?.id ?? "",
                          sourceFilename: upload?.filename ?? "",
                          sheetName: "",
                        });
                      }}
                    >
                      <option value="">Select file</option>
                      {excelUploads.map((upload) => <option key={upload.id} value={upload.id}>{upload.filename}</option>)}
                    </select>
                  </label>
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Sheet
                    <select
                      aria-label="Bar Excel sheet"
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-50 disabled:text-slate-400"
                      disabled={!widget.sourceUploadId || excelSheets.length === 0}
                      value={widget.sheetName ?? ""}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { sheetName: event.target.value })}
                    >
                      <option value="">Select sheet</option>
                      {excelSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                    </select>
                  </label>
                </div>

                {widget.type === "bar" ? (
                  <label className="flex items-center justify-between rounded border border-slate-200 bg-slate-50 px-3 py-2 text-[10px] font-black uppercase text-slate-600">
                    <span>
                      Display as % of target
                      <span className="mt-0.5 block text-[9px] font-bold normal-case text-slate-400">Actual value divided by each Series target.</span>
                    </span>
                    <input
                      className="h-4 w-4 accent-cj-blue"
                      type="checkbox"
                      checked={widget.barDisplayPercentage ?? false}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barDisplayPercentage: event.target.checked })}
                    />
                  </label>
                ) : null}

                {!barPercentageMode || isBarMarkers ? (
                  <label className="block text-[10px] font-black uppercase text-slate-500">
                    Maximum value
                    <input
                      className="mt-1 h-9 w-full rounded border border-slate-200 px-3 text-sm font-bold uppercase text-cj-navy outline-none focus:border-cj-blue"
                      placeholder="100 / C7 / MAX(C7:C12)"
                      type="text"
                      value={widget.barMaxInput ?? String(widget.barMax ?? 100)}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barMaxInput: event.target.value.toUpperCase() })}
                    />
                    <span className={`mt-1 block text-[9px] font-bold normal-case ${barMaxStatus.startsWith("Cannot") ? "text-red-500" : "text-slate-400"}`}>{barMaxStatus}</span>
                  </label>
                ) : null}

                <div>
                  <div className="flex items-center justify-between">
                    <div>
                      <h3 className="text-xs font-black text-cj-navy">Series</h3>
                      <p className="text-[10px] font-bold text-slate-400">One cell per horizontal bar.</p>
                    </div>
                    <button
                      aria-label="Add bar"
                      className="grid h-8 w-8 place-items-center rounded-md bg-cj-navy text-white transition hover:bg-cj-blue"
                      title="Add bar"
                      type="button"
                      onClick={addBarItem}
                    >
                      <Plus className="h-4 w-4" />
                    </button>
                  </div>

                  <DndContext collisionDetection={closestCenter} onDragEnd={reorderBarItems}>
                    <SortableContext items={currentBarItems.map((item) => item.id)} strategy={verticalListSortingStrategy}>
                      <div className="mt-2 space-y-2">
                    {currentBarItems.map((item, index) => (
                      <SortableChartRow key={item.id} id={item.id} className={`grid items-end gap-2 rounded border border-slate-200 bg-slate-50 p-2 ${isBarMarkers || barPercentageMode ? "grid-cols-[28px_minmax(0,1fr)_62px_62px_28px]" : "grid-cols-[28px_minmax(0,1fr)_76px_28px]"}`}>
                        <label className="text-[9px] font-black uppercase text-slate-500">
                          Label
                          <input
                            className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                            value={item.label}
                            onChange={(event) => updateBarItem(item.id, { label: event.target.value })}
                          />
                        </label>
                        <label className="text-[9px] font-black uppercase text-slate-500">
                          Cell / formula
                          <input
                            className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold uppercase text-cj-navy outline-none focus:border-cj-blue"
                            placeholder="D7+E7 / 80"
                            value={item.cell}
                            onChange={(event) => updateBarItem(item.id, { cell: event.target.value.toUpperCase() })}
                          />
                        </label>
                        {isBarMarkers || barPercentageMode ? (
                          <label className="text-[9px] font-black uppercase text-slate-500">
                            Target source
                            <input
                              className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold uppercase text-cj-navy outline-none focus:border-cj-blue"
                              placeholder="D7 / 80"
                              value={item.markerCell ?? (barPercentageMode ? "100" : "")}
                              onChange={(event) => updateBarItem(item.id, { markerCell: event.target.value.toUpperCase() })}
                            />
                          </label>
                        ) : null}
                        <button
                          aria-label={`Remove bar ${index + 1}`}
                          className="grid h-8 w-7 place-items-center rounded text-red-500 transition hover:bg-red-50 disabled:opacity-30"
                          disabled={currentBarItems.length <= 1}
                          title="Remove bar"
                          type="button"
                          onClick={() => removeBarItem(item.id)}
                        >
                          <X className="h-3.5 w-3.5" />
                        </button>
                      </SortableChartRow>
                    ))}
                      </div>
                    </SortableContext>
                  </DndContext>
                </div>
                {widget.type === "simple-pie" ? (
                  <label className="flex items-center justify-between rounded border border-slate-200 bg-slate-50 px-3 py-2 text-[10px] font-black uppercase text-slate-600">
                    <span>
                      Show value callouts
                      <span className="mt-0.5 block text-[9px] font-bold normal-case text-slate-400">Keep each Series value visible outside the Pie.</span>
                    </span>
                    <input
                      className="h-4 w-4 accent-cj-blue"
                      type="checkbox"
                      checked={widget.pieShowValueCallouts ?? false}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { pieShowValueCallouts: event.target.checked })}
                    />
                  </label>
                ) : null}
                {renderChartAppearance()}
                <p className="text-[10px] font-bold text-slate-400">{barStatus}</p>
              </div>
            </div>
          ), document.body) : null}

          {editable && isSelected && (isStackBar || isLineChart || isColumnChart || isRadarChart) ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className="fixed z-[100] max-h-[calc(100vh-6rem)] w-96 overflow-y-auto"
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  {isAnnotatedLine ? "Line with Annotations" : isLineChart ? "Basic Line Chart" : isColumnChart ? "Column with Rotated Labels" : isRadarChart ? "Radar with Polygon Fill" : isStackColumn100 ? "Stacked Column 100%" : widget.type === "stack-column" ? "Stack Column" : widget.type === "stack-100-bar" ? "Stack 100% Bar" : "Stack Bar"} tools
                </button>
                {renderWidgetToolbarActions()}

                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Excel file
                    <select
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.sourceUploadId ?? ""}
                      onChange={(event) => {
                        const upload = excelUploads.find((item) => item.id === event.target.value);
                        onUpdateWidget(boxId, widget.id, { sourceUploadId: upload?.id ?? "", sourceFilename: upload?.filename ?? "", sheetName: "" });
                      }}
                    >
                      <option value="">Select file</option>
                      {excelUploads.map((upload) => <option key={upload.id} value={upload.id}>{upload.filename}</option>)}
                    </select>
                  </label>
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Sheet
                    <select
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-50"
                      disabled={!widget.sourceUploadId || excelSheets.length === 0}
                      value={widget.sheetName ?? ""}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { sheetName: event.target.value })}
                    >
                      <option value="">Select sheet</option>
                      {excelSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                    </select>
                  </label>
                </div>

                {widget.type === "stack-bar" || widget.type === "stack-column" || isLineChart || isColumnChart || isRadarChart ? (
                  <label className="block text-[10px] font-black uppercase text-slate-500">
                    Maximum value
                    <input
                      className="mt-1 h-9 w-full rounded border border-slate-200 px-3 text-sm font-bold uppercase text-cj-navy outline-none focus:border-cj-blue"
                      placeholder="100 / C7 / MAX(C7:C12)"
                      type="text"
                      value={widget.barMaxInput ?? String(widget.barMax ?? 100)}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { barMaxInput: event.target.value.toUpperCase() })}
                    />
                    <span className={`mt-1 block text-[9px] font-bold normal-case ${barMaxStatus.startsWith("Cannot") ? "text-red-500" : "text-slate-400"}`}>{barMaxStatus}</span>
                  </label>
                ) : (
                  <p className="rounded bg-blue-50 px-3 py-2 text-[10px] font-bold text-cj-blue">Scale is normalized automatically from 0% to 100%.</p>
                )}

                {widget.type === "stack-column" ? (
                  <label className="flex items-center justify-between rounded border border-slate-200 bg-slate-50 p-2 text-[10px] font-black uppercase text-slate-500">
                    <span>
                      Normalize bars to 100%
                      <span className="mt-0.5 block text-[9px] font-bold normal-case text-slate-400">Series labels always show %. Enable this to scale each stack to 100%.</span>
                    </span>
                    <input
                      className="h-4 w-4 accent-cj-blue"
                      type="checkbox"
                      checked={widget.stackColumnPercentage ?? false}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { stackColumnPercentage: event.target.checked })}
                    />
                  </label>
                ) : null}

                {isLineChart ? (
                  <div className="grid grid-cols-3 gap-2 rounded border border-slate-200 bg-slate-50 p-2">
                    <label className="text-[9px] font-black uppercase text-slate-500">
                      Curve
                      <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-1 text-xs font-bold normal-case text-cj-navy" value={widget.lineCurve ?? "smooth"} onChange={(event) => onUpdateWidget(boxId, widget.id, { lineCurve: event.target.value as "smooth" | "straight" | "stepline" })}>
                        <option value="smooth">Smooth</option>
                        <option value="straight">Straight</option>
                        <option value="stepline">Step</option>
                      </select>
                    </label>
                    <label className="text-[9px] font-black uppercase text-slate-500">
                      Line width
                      <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-1 text-xs font-bold normal-case text-cj-navy" value={widget.lineStrokeWidth ?? 3} onChange={(event) => onUpdateWidget(boxId, widget.id, { lineStrokeWidth: Number(event.target.value) })}>
                        {[1, 2, 3, 4, 5, 6].map((width) => <option key={width} value={width}>{width}px</option>)}
                      </select>
                    </label>
                    <label className="flex cursor-pointer flex-col text-[9px] font-black uppercase text-slate-500">
                      Markers
                      <span className="mt-1 flex h-8 items-center justify-center gap-2 rounded border border-slate-200 bg-white text-xs normal-case text-cj-navy">
                        <input className="h-4 w-4 accent-cj-blue" type="checkbox" checked={widget.lineShowMarkers ?? true} onChange={(event) => onUpdateWidget(boxId, widget.id, { lineShowMarkers: event.target.checked })} />
                        Show
                      </span>
                    </label>
                    {widget.type === "basic-line" ? (
                      <label className="col-span-3 flex cursor-pointer items-center justify-between rounded border border-slate-200 bg-white px-3 py-2 text-[9px] font-black uppercase text-slate-500">
                        <span>
                          Null / Missing values
                          <span className="mt-0.5 block text-[9px] font-bold normal-case text-slate-400">Leave gaps in the line when an Excel Cell is empty.</span>
                        </span>
                        <input
                          className="h-4 w-4 accent-cj-blue"
                          type="checkbox"
                          checked={widget.lineNullMissing ?? false}
                          onChange={(event) => onUpdateWidget(boxId, widget.id, { lineNullMissing: event.target.checked })}
                        />
                      </label>
                    ) : null}
                  </div>
                ) : null}

                {isColumnChart ? (
                  <div className="grid grid-cols-2 gap-2 rounded border border-slate-200 bg-slate-50 p-2">
                    <label className="text-[9px] font-black uppercase text-slate-500">
                      Label rotation
                      <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy" value={widget.columnLabelRotation ?? -45} onChange={(event) => onUpdateWidget(boxId, widget.id, { columnLabelRotation: Number(event.target.value) })}>
                        {[-90, -75, -60, -45, -30, -15, 0].map((angle) => <option key={angle} value={angle}>{angle}°</option>)}
                      </select>
                    </label>
                    <label className="text-[9px] font-black uppercase text-slate-500">
                      Column width
                      <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy" value={widget.columnWidth ?? 55} onChange={(event) => onUpdateWidget(boxId, widget.id, { columnWidth: Number(event.target.value) })}>
                        {[30, 40, 50, 55, 60, 70, 80, 90].map((width) => <option key={width} value={width}>{width}%</option>)}
                      </select>
                    </label>
                    <label className="flex cursor-pointer flex-col text-[9px] font-black uppercase text-slate-500">
                      Data labels
                      <span className="mt-1 flex h-8 items-center justify-center gap-2 rounded border border-slate-200 bg-white text-xs normal-case text-cj-navy">
                        <input className="h-4 w-4 accent-cj-blue" type="checkbox" checked={widget.columnShowDataLabels ?? true} onChange={(event) => onUpdateWidget(boxId, widget.id, { columnShowDataLabels: event.target.checked })} />
                        Show
                      </span>
                    </label>
                    <label className={`flex flex-col text-[9px] font-black uppercase text-slate-500 ${(widget.columnShowDataLabels ?? true) ? "cursor-pointer" : "cursor-not-allowed opacity-50"}`}>
                      Label background
                      <span className="mt-1 flex h-8 items-center justify-center gap-2 rounded border border-slate-200 bg-white text-xs normal-case text-cj-navy">
                        <input
                          className="h-4 w-4 accent-cj-blue"
                          type="checkbox"
                          disabled={!(widget.columnShowDataLabels ?? true)}
                          checked={widget.columnShowDataLabelBackground ?? true}
                          onChange={(event) => onUpdateWidget(boxId, widget.id, { columnShowDataLabelBackground: event.target.checked })}
                        />
                        Circle
                      </span>
                    </label>
                  </div>
                ) : null}

                {isRadarChart ? (
                  <div className="space-y-2 rounded border border-slate-200 bg-slate-50 p-2">
                    <div>
                      <h3 className="text-xs font-black text-cj-navy">Radar appearance</h3>
                      <p className="text-[9px] font-bold text-slate-400">Polygon, line, fill, and marker styling.</p>
                    </div>
                    <div className="grid grid-cols-3 gap-2">
                      <label className="text-[8px] font-black uppercase text-slate-500">
                        Line {widget.radarStrokeWidth ?? 2}px
                        <input className="mt-2 w-full accent-cj-blue" type="range" min={1} max={8} value={widget.radarStrokeWidth ?? 2} onChange={(event) => onUpdateWidget(boxId, widget.id, { radarStrokeWidth: Number(event.target.value) })} />
                      </label>
                      <label className="text-[8px] font-black uppercase text-slate-500">
                        Fill {Math.round((widget.radarFillOpacity ?? 0.2) * 100)}%
                        <input className="mt-2 w-full accent-cj-blue" type="range" min={0} max={80} value={Math.round((widget.radarFillOpacity ?? 0.2) * 100)} onChange={(event) => onUpdateWidget(boxId, widget.id, { radarFillOpacity: Number(event.target.value) / 100 })} />
                      </label>
                      <label className="text-[8px] font-black uppercase text-slate-500">
                        Marker {widget.radarMarkerSize ?? 4}px
                        <input className="mt-2 w-full accent-cj-blue" type="range" min={0} max={12} value={widget.radarMarkerSize ?? 4} onChange={(event) => onUpdateWidget(boxId, widget.id, { radarMarkerSize: Number(event.target.value) })} />
                      </label>
                    </div>
                    <div className="grid grid-cols-3 gap-2">
                      {([
                        ["Polygon A", "radarPolygonColor1", widget.radarPolygonColor1 ?? "#f8fafc"],
                        ["Polygon B", "radarPolygonColor2", widget.radarPolygonColor2 ?? "#eef2f6"],
                        ["Grid line", "radarPolygonStrokeColor", widget.radarPolygonStrokeColor ?? "#cbd5e1"],
                      ] as const).map(([label, key, color]) => (
                        <label key={key} className="text-center text-[8px] font-black uppercase text-slate-500">
                          {label}
                          <input aria-label={`${label} color`} className="mt-1 h-7 w-full cursor-pointer rounded border-0 bg-transparent p-0" type="color" value={color} onChange={(event) => onUpdateWidget(boxId, widget.id, { [key]: event.target.value })} />
                        </label>
                      ))}
                    </div>
                  </div>
                ) : null}

                {isAnnotatedLine ? (
                  <div className="rounded border border-slate-200 bg-slate-50 p-2">
                    <div className="flex items-center justify-between">
                      <div><h3 className="text-xs font-black text-cj-navy">Annotations</h3><p className="text-[10px] font-bold text-slate-400">Reference lines on the X or Y axis.</p></div>
                      <button className="grid h-8 w-8 place-items-center rounded-md bg-cj-navy text-white hover:bg-cj-blue" title="Add annotation" type="button" onClick={addLineAnnotation}><Plus className="h-4 w-4" /></button>
                    </div>
                    <div className="mt-2 space-y-2">
                      {currentLineAnnotations.length > 0 ? currentLineAnnotations.map((annotation) => (
                        <div key={annotation.id} className="grid grid-cols-[minmax(0,1fr)_58px_66px_30px_28px] items-end gap-1.5">
                          <label className="text-[9px] font-black uppercase text-slate-500">
                            Label
                            <input className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy" value={annotation.label} onChange={(event) => updateLineAnnotation(annotation.id, { label: event.target.value })} />
                          </label>
                          <label className="text-[9px] font-black uppercase text-slate-500">
                            Axis
                            <select className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-1 text-xs font-bold text-cj-navy" value={annotation.axis} onChange={(event) => updateLineAnnotation(annotation.id, { axis: event.target.value as "x" | "y" })}>
                              <option value="y">Y</option>
                              <option value="x">X</option>
                            </select>
                          </label>
                          <label className="text-[9px] font-black uppercase text-slate-500">
                            Value
                            <input className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold text-cj-navy" placeholder={annotation.axis === "x" ? "Week 2" : "80"} value={annotation.value} onChange={(event) => updateLineAnnotation(annotation.id, { value: event.target.value })} />
                          </label>
                          <label className="text-[9px] font-black uppercase text-slate-500">
                            Color
                            <input aria-label={`${annotation.label} color`} className="mt-1 h-8 w-full cursor-pointer border-0 bg-transparent p-0" type="color" value={annotation.color} onChange={(event) => updateLineAnnotation(annotation.id, { color: event.target.value })} />
                          </label>
                          <button className="grid h-8 w-7 place-items-center rounded text-red-500 hover:bg-red-50" title="Remove annotation" type="button" onClick={() => removeLineAnnotation(annotation.id)}><X className="h-3.5 w-3.5" /></button>
                        </div>
                      )) : <p className="py-2 text-center text-[10px] font-bold text-slate-400">No annotations</p>}
                    </div>
                  </div>
                ) : null}

                <div>
                  <div className="flex items-center justify-between">
                    <div><h3 className="text-xs font-black text-cj-navy">{isLineChart || isColumnChart || isRadarChart || isStackColumn ? "Categories" : "Bars"}</h3><p className="text-[10px] font-bold text-slate-400">Labels on the chart axis.</p></div>
                    <button className="grid h-8 w-8 place-items-center rounded-md bg-cj-navy text-white hover:bg-cj-blue" title="Add category" type="button" onClick={addStackCategory}><Plus className="h-4 w-4" /></button>
                  </div>
                  <div className="mt-2 space-y-1.5">
                    {currentStackCategories.map((category, index) => (
                      <div key={category.id} className="grid grid-cols-[24px_1fr_28px] items-center gap-2">
                        <span className="text-center text-[10px] font-black text-slate-400">{index + 1}</span>
                        <input className="h-8 rounded border border-slate-200 px-2 text-xs font-bold text-cj-navy outline-none focus:border-cj-blue" value={category.label} onChange={(event) => updateStackCategory(category.id, event.target.value)} />
                        <button className="grid h-8 w-7 place-items-center rounded text-red-500 hover:bg-red-50 disabled:opacity-30" disabled={currentStackCategories.length <= 1} title="Remove bar" type="button" onClick={() => removeStackCategory(index)}><X className="h-3.5 w-3.5" /></button>
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <div className="flex items-center justify-between">
                    <div><h3 className="text-xs font-black text-cj-navy">Series data</h3><p className="text-[10px] font-bold text-slate-400">Use a Cell, formula, or raw number.</p></div>
                    <button className="grid h-8 w-8 place-items-center rounded-md bg-cj-navy text-white hover:bg-cj-blue" title="Add series" type="button" onClick={addStackSeries}><Plus className="h-4 w-4" /></button>
                  </div>
                  <DndContext sensors={seriesSortSensors} collisionDetection={closestCenter} onDragEnd={reorderStackSeries}>
                    <SortableContext items={currentStackSeries.map((series) => series.id)} strategy={verticalListSortingStrategy}>
                  <div className="mt-2 space-y-2">
                    {currentStackSeries.map((series, seriesIndex) => (
                      <SortableChartRow key={series.id} id={series.id} className="grid grid-cols-[28px_minmax(0,1fr)] items-start gap-1 rounded border border-slate-200 bg-slate-50 p-2">
                        <div className="min-w-0">
                        <div className="grid grid-cols-[20px_1fr_28px] items-center gap-2">
                          <span className="text-center text-[10px] font-black text-slate-400">{seriesIndex + 1}</span>
                          <input className="h-8 rounded border border-slate-200 bg-white px-2 text-xs font-bold text-cj-navy outline-none focus:border-cj-blue" placeholder="Series name" value={series.label} onChange={(event) => updateStackSeries(series.id, { label: event.target.value })} />
                          <button className="grid h-8 w-7 place-items-center rounded text-red-500 hover:bg-red-50 disabled:opacity-30" disabled={currentStackSeries.length <= 1} title="Remove series" type="button" onClick={() => removeStackSeries(series.id)}><X className="h-3.5 w-3.5" /></button>
                        </div>
                        <div className="mt-2 grid grid-cols-2 gap-1.5">
                          {currentStackCategories.map((category, categoryIndex) => {
                            const suggestedCell = widget.type === "basic-line" && categoryIndex > 0
                              ? lineCellSuggestions[series.id]?.[categoryIndex]
                              : undefined;
                            return (
                            <label key={category.id} className="text-[9px] font-black uppercase text-slate-500">
                              <span className="block truncate" title={category.label}>{category.label}</span>
                              <input
                                className={`mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold uppercase outline-none focus:border-cj-blue ${suggestedCell ? "text-slate-400 italic" : "text-cj-navy"}`}
                                placeholder="D7+E7 / 80"
                                value={suggestedCell ?? series.cells[categoryIndex] ?? ""}
                                onChange={(event) => updateSeriesCell(series, categoryIndex, event.target.value)}
                                onKeyDown={(event) => {
                                  if (widget.type === "basic-line" && event.key === "Tab" && !event.shiftKey && categoryIndex === 0) {
                                    confirmRunningSeriesCells(series.id);
                                  }
                                }}
                              />
                            </label>
                            );
                          })}
                        </div>
                        </div>
                      </SortableChartRow>
                    ))}
                  </div>
                    </SortableContext>
                  </DndContext>
                </div>
                {renderChartAppearance()}
                <p className="text-[10px] font-bold text-slate-400">{barStatus}</p>
              </div>
            </div>
          ), document.body) : null}

          {editable && isSelected && isSimplePie ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className="fixed z-[100] max-h-[calc(100vh-6rem)] w-96 overflow-y-auto"
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  {widget.type === "simple-donut" ? "Simple Donut" : "Simple Pie"} tools
                </button>
                {renderWidgetToolbarActions()}

                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Excel file
                    <select
                      aria-label="Pie Excel file"
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.sourceUploadId ?? ""}
                      onChange={(event) => {
                        const upload = excelUploads.find((item) => item.id === event.target.value);
                        onUpdateWidget(boxId, widget.id, { sourceUploadId: upload?.id ?? "", sourceFilename: upload?.filename ?? "", sheetName: "" });
                      }}
                    >
                      <option value="">Select file</option>
                      {excelUploads.map((upload) => <option key={upload.id} value={upload.id}>{upload.filename}</option>)}
                    </select>
                  </label>
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Sheet
                    <select
                      aria-label="Pie Excel sheet"
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-50"
                      disabled={!widget.sourceUploadId || excelSheets.length === 0}
                      value={widget.sheetName ?? ""}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { sheetName: event.target.value })}
                    >
                      <option value="">Select sheet</option>
                      {excelSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                    </select>
                  </label>
                </div>

                <div>
                  <div className="flex items-center justify-between">
                    <div>
                      <h3 className="text-xs font-black text-cj-navy">Slices</h3>
                      <p className="text-[10px] font-bold text-slate-400">One non-negative Excel value per slice.</p>
                    </div>
                    <button className="grid h-8 w-8 place-items-center rounded-md bg-cj-navy text-white hover:bg-cj-blue" title="Add slice" type="button" onClick={addBarItem}><Plus className="h-4 w-4" /></button>
                  </div>

                  <DndContext collisionDetection={closestCenter} onDragEnd={reorderBarItems}>
                    <SortableContext items={currentBarItems.map((item) => item.id)} strategy={verticalListSortingStrategy}>
                  <div className="mt-2 space-y-2">
                    {currentBarItems.map((item) => (
                      <SortableChartRow key={item.id} id={item.id} className="grid grid-cols-[28px_minmax(0,1fr)_76px_28px] items-end gap-2 rounded border border-slate-200 bg-slate-50 p-2">
                        <label className="text-[9px] font-black uppercase text-slate-500">
                          Label
                          <input className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue" value={item.label} onChange={(event) => updateBarItem(item.id, { label: event.target.value })} />
                        </label>
                        <label className="text-[9px] font-black uppercase text-slate-500">
                          Cell / formula
                          <input className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold uppercase text-cj-navy outline-none focus:border-cj-blue" placeholder="D7+E7 / 80" value={item.cell} onChange={(event) => updateBarItem(item.id, { cell: event.target.value.toUpperCase() })} />
                        </label>
                        <button className="grid h-8 w-7 place-items-center rounded text-red-500 hover:bg-red-50 disabled:opacity-30" disabled={currentBarItems.length <= 1} title="Remove slice" type="button" onClick={() => removeBarItem(item.id)}><X className="h-3.5 w-3.5" /></button>
                      </SortableChartRow>
                    ))}
                  </div>
                    </SortableContext>
                  </DndContext>
                </div>
                {renderChartAppearance()}
                <p className="text-[10px] font-bold text-slate-400">{barStatus}</p>
              </div>
            </div>
          ), document.body) : null}

          {editable && isSelected && widget.type === "gradient-color" ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className="fixed z-[100] max-h-[calc(100vh-6rem)] w-80 overflow-y-auto"
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  Gradient Color tools
                </button>
                {renderWidgetToolbarActions()}

                <div className="rounded border border-slate-200 bg-slate-50 p-2">
                  <div className="mb-2 flex items-center justify-between">
                    <div>
                      <h3 className="text-[10px] font-black uppercase text-slate-600">Color stops</h3>
                      <p className="text-[9px] font-bold text-slate-400">Click the strip to add a point.</p>
                    </div>
                    <button
                      className="grid h-7 w-7 place-items-center rounded bg-cj-navy text-white hover:bg-cj-blue"
                      title="Add color stop"
                      type="button"
                      onClick={() => {
                        const gaps = currentGradientStops.slice(0, -1).map((stop, index) => ({
                          size: currentGradientStops[index + 1].position - stop.position,
                          position: (stop.position + currentGradientStops[index + 1].position) / 2,
                        }));
                        addGradientStop(gaps.sort((left, right) => right.size - left.size)[0]?.position ?? 50);
                      }}
                    >
                      <Plus className="h-3.5 w-3.5" />
                    </button>
                  </div>
                  <div className="relative mb-7">
                    <button
                      aria-label="Add gradient color stop"
                      className="block h-9 w-full rounded border border-slate-300 shadow-inner"
                      style={{ background: gradientCss }}
                      title="Click to add a color stop"
                      type="button"
                      onClick={(event) => {
                        const bounds = event.currentTarget.getBoundingClientRect();
                        addGradientStop(((event.clientX - bounds.left) / bounds.width) * 100);
                      }}
                    />
                    {currentGradientStops.map((stop) => (
                      <button
                        key={stop.id}
                        aria-label={`Select color stop at ${stop.position}%`}
                        className={`absolute top-[calc(100%+2px)] -translate-x-1/2 transition-transform hover:scale-110 ${selectedGradientStop?.id === stop.id ? "z-10 scale-110" : ""}`}
                        style={{ left: `${stop.position}%`, color: selectedGradientStop?.id === stop.id ? "#1473e6" : "#64748b" }}
                        title={`${stop.position}% - ${stop.color}`}
                        type="button"
                        onClick={() => setSelectedGradientStopId(stop.id)}
                      >
                        <Triangle className="h-5 w-5" fill={stop.color} strokeWidth={selectedGradientStop?.id === stop.id ? 3 : 2} />
                      </button>
                    ))}
                  </div>

                  {selectedGradientStop ? (
                    <div className="grid grid-cols-[1fr_1fr_auto] items-end gap-2">
                      <label className="text-[9px] font-black uppercase text-slate-500">
                        Stop color
                        <span className="mt-1 flex h-9 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                          <input
                            aria-label="Selected stop color"
                            className="h-6 w-8 cursor-pointer border-0 bg-transparent p-0"
                            type="color"
                            value={selectedGradientStop.color}
                            onChange={(event) => updateGradientStops(currentGradientStops.map((stop) => stop.id === selectedGradientStop.id ? { ...stop, color: event.target.value } : stop))}
                          />
                          <span className="text-[9px] font-bold uppercase text-slate-500">{selectedGradientStop.color}</span>
                        </span>
                      </label>
                      <label className="text-[9px] font-black uppercase text-slate-500">
                        Position
                        <span className="mt-1 flex h-9 items-center rounded border border-slate-200 bg-white px-2">
                          <input
                            aria-label="Selected stop position"
                            className="w-full text-xs font-bold text-cj-navy outline-none"
                            max="100"
                            min="0"
                            type="number"
                            value={selectedGradientStop.position}
                            onChange={(event) => updateGradientStops(currentGradientStops.map((stop) => stop.id === selectedGradientStop.id ? { ...stop, position: Math.max(0, Math.min(100, Number(event.target.value))) } : stop))}
                          />
                          <span className="text-[10px] font-black text-slate-400">%</span>
                        </span>
                      </label>
                      <button
                        className="grid h-9 w-9 place-items-center rounded border border-red-200 text-red-600 hover:bg-red-50 disabled:cursor-not-allowed disabled:opacity-30"
                        disabled={currentGradientStops.length <= 2}
                        title="Delete selected color stop"
                        type="button"
                        onClick={() => {
                          const remaining = currentGradientStops.filter((stop) => stop.id !== selectedGradientStop.id);
                          updateGradientStops(remaining);
                          setSelectedGradientStopId(remaining[0]?.id ?? null);
                        }}
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </div>
                  ) : null}
                </div>

                <label className="block text-[9px] font-black uppercase text-slate-500">
                  Direction
                  <select
                    className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                    value={widget.gradientDirection ?? "to right"}
                    onChange={(event) => onUpdateWidget(boxId, widget.id, { gradientDirection: event.target.value })}
                  >
                    <option value="to right">Left to right</option>
                    <option value="to left">Right to left</option>
                    <option value="to bottom">Top to bottom</option>
                    <option value="to top">Bottom to top</option>
                    <option value="135deg">Diagonal down</option>
                    <option value="45deg">Diagonal up</option>
                  </select>
                </label>

                {([
                  ["Opacity", "gradientOpacity", widget.gradientOpacity ?? 100],
                  ["Corner radius", "gradientBorderRadius", widget.gradientBorderRadius ?? 8],
                ] as const).map(([label, field, value]) => (
                  <label key={field} className="block rounded border border-slate-200 bg-slate-50 p-2 text-[9px] font-black uppercase text-slate-500">
                    <span className="flex items-center justify-between">
                      {label}
                      <output className="rounded bg-white px-2 py-0.5 text-[10px] font-black normal-case text-cj-navy">{value}{field === "gradientBorderRadius" ? "px" : "%"}</output>
                    </span>
                    <input
                      aria-label={label}
                      className="mt-2 h-2 w-full cursor-pointer accent-cj-blue"
                      max={field === "gradientBorderRadius" ? 32 : 100}
                      min="0"
                      step="1"
                      type="range"
                      value={value}
                      onChange={(event) => {
                        const nextValue = Number(event.target.value);
                        onUpdateWidget(boxId, widget.id, { [field]: nextValue });
                      }}
                    />
                  </label>
                ))}
              </div>
            </div>
          ), document.body) : null}

          {editable && isSelected && widget.type === "icon" ? createPortal((
            <div
              data-widget-toolbar={widget.id}
              className="fixed z-[100] max-h-[calc(100vh-6rem)] w-80 overflow-y-auto"
              style={{ left: toolbarPosition.x, top: toolbarPosition.y }}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={handleToolbarFieldNavigation}
            >
              <div className="space-y-3 rounded-md border border-slate-200 bg-white p-3 text-left shadow-panel">
                <button
                  className="-mx-1 -mt-1 flex h-8 w-[calc(100%+0.5rem)] cursor-grab items-center gap-2 rounded bg-slate-50 px-2 text-xs font-black text-cj-navy active:cursor-grabbing"
                  title="Move toolbar"
                  type="button"
                  onPointerDown={startToolbarDrag}
                >
                  <GripVertical className="h-4 w-4 text-slate-400" />
                  Icon tools
                </button>
                {renderWidgetToolbarActions()}

                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Color
                    <div className="mt-1 flex h-9 items-center gap-2 rounded border border-slate-200 bg-white px-2">
                      <input
                        aria-label="Icon color"
                        className="h-6 w-8 cursor-pointer border-0 bg-transparent p-0"
                        type="color"
                        value={widget.iconColor ?? "#1473e6"}
                        onChange={(event) => onUpdateWidget(boxId, widget.id, { iconColor: event.target.value })}
                      />
                      <span className="text-[10px] font-bold uppercase text-slate-500">{widget.iconColor ?? "#1473e6"}</span>
                    </div>
                  </label>
                  <label className="text-[10px] font-black uppercase text-slate-500">
                    Size
                    <select
                      className="mt-1 h-9 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={widget.iconSize ?? 32}
                      onChange={(event) => onUpdateWidget(boxId, widget.id, { iconSize: Number(event.target.value) })}
                    >
                      {iconSizeOptions.map((size) => <option key={size} value={size}>{size}px</option>)}
                    </select>
                  </label>
                </div>

                <div>
                  <h3 className="text-xs font-black text-cj-navy">Choose icon</h3>
                  <p className="text-[10px] font-bold text-slate-400">Select a symbol for this widget.</p>
                  <div className="mt-2 grid grid-cols-6 gap-1.5">
                    {iconOptions.map((option) => {
                      const IconChoice = option.icon;
                      const selected = (widget.iconName ?? "package") === option.name;
                      return (
                        <button
                          key={option.name}
                          aria-label={option.label}
                          className={`grid aspect-square place-items-center rounded border transition ${selected ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 bg-white text-slate-500 hover:border-slate-400 hover:text-cj-navy"}`}
                          title={option.label}
                          type="button"
                          onClick={() => onUpdateWidget(boxId, widget.id, { iconName: option.name })}
                        >
                          <IconChoice className="h-5 w-5" />
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          ), document.body) : null}
        </div>
      ) : null}
    </div>
  );
}

function WorkspaceBoxCard({
  box,
  editable,
  showWidgetSlots,
  widgetPlacementPreview,
  onRemoveBox,
  onRemoveWidget,
  onResizeBox,
  onResizeWidget,
  onUpdateWidget,
}: {
  box: WorkspaceBox;
  editable: boolean;
  showWidgetSlots: boolean;
  widgetPlacementPreview?: { slot: number; width: number; height: number; valid: boolean };
  onRemoveBox: (boxId: string) => void;
  onRemoveWidget: (boxId: string, widgetId: string) => void;
  onResizeBox: (boxId: string, columns: number, rows: number, cell?: number) => void;
  onResizeWidget: (boxId: string, widgetId: string, slot: number, width: number, height: number) => void;
  onUpdateWidget: (boxId: string, widgetId: string, changes: Partial<WorkspaceWidget>) => void;
}) {
  const [isSelected, setIsSelected] = useState(false);
  const position = getCellPosition(box.cell);
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: `box:${box.id}`,
    data: { kind: "box", box },
    disabled: !editable,
  });
  const style: CSSProperties = {
    gridColumn: `${position.column} / span ${box.columns}`,
    gridRow: `${position.row} / span ${box.rows}`,
  };
  const slotColumns = box.columns * slotsPerCellColumn;
  const slotRows = box.rows * slotsPerCellRow;
  const slotCount = slotColumns * slotRows;

  useEffect(() => {
    if (!editable) setIsSelected(false);
  }, [editable]);

  useEffect(() => {
    if (!isSelected) return;

    function handleOutsidePointer(event: globalThis.PointerEvent) {
      const target = event.target as Element | null;
      const selectedBox = target?.closest("[data-workspace-box]");
      if (selectedBox?.getAttribute("data-workspace-box") !== box.id) setIsSelected(false);
    }

    document.addEventListener("pointerdown", handleOutsidePointer);
    return () => document.removeEventListener("pointerdown", handleOutsidePointer);
  }, [box.id, isSelected]);

  function startResize(direction: "left" | "top" | "right" | "bottom" | "corner" | "top-left", event: PointerEvent<HTMLButtonElement>) {
    event.preventDefault();
    event.stopPropagation();

    const startX = event.clientX;
    const startY = event.clientY;
    const startColumns = box.columns;
    const startRows = box.rows;
    const startPosition = getCellPosition(box.cell);
    const cellStep = 82;

    function handleMove(moveEvent: globalThis.PointerEvent) {
      const rawColumnDelta = Math.round((moveEvent.clientX - startX) / cellStep);
      const rawRowDelta = Math.round((moveEvent.clientY - startY) / cellStep);
      let nextColumn = startPosition.column;
      let nextRow = startPosition.row;
      let nextColumns = startColumns;
      let nextRows = startRows;

      if (direction === "right" || direction === "corner") {
        nextColumns = startColumns + rawColumnDelta;
      }

      if (direction === "bottom" || direction === "corner") {
        nextRows = startRows + rawRowDelta;
      }

      if (direction === "left" || direction === "top-left") {
        nextColumn = startPosition.column + rawColumnDelta;
        nextColumns = startColumns - rawColumnDelta;
      }

      if (direction === "top" || direction === "top-left") {
        nextRow = startPosition.row + rawRowDelta;
        nextRows = startRows - rawRowDelta;
      }

      nextColumn = Math.max(1, Math.min(workspaceColumns, nextColumn));
      nextRow = Math.max(1, Math.min(workspaceRows, nextRow));
      nextColumns = Math.max(1, Math.min(workspaceColumns - nextColumn + 1, nextColumns));
      nextRows = Math.max(1, Math.min(workspaceRows - nextRow + 1, nextRows));

      const nextCell = (nextRow - 1) * workspaceColumns + (nextColumn - 1);
      onResizeBox(box.id, nextColumns, nextRows, nextCell);
    }

    function handleUp() {
      window.removeEventListener("pointermove", handleMove);
      window.removeEventListener("pointerup", handleUp);
    }

    window.addEventListener("pointermove", handleMove);
    window.addEventListener("pointerup", handleUp);
  }

  return (
    <section
      ref={setNodeRef}
      data-workspace-box={box.id}
      className={`relative rounded-[10px] border bg-white shadow-panel transition hover:z-30 ${editable && isSelected ? "z-40 border-cj-blue" : "z-10 border-slate-200"} ${isDragging ? "opacity-50" : ""}`}
      style={style}
      onClick={() => editable && setIsSelected(true)}
    >
      {editable ? <div className="absolute -right-3 -top-3 z-30 flex items-center gap-1">
        <button
          className="grid h-7 w-7 cursor-grab place-items-center rounded-full border border-slate-200 bg-white text-slate-500 shadow-sm"
          title="Move box"
          type="button"
          {...attributes}
          {...listeners}
        >
          <GripVertical className="h-4 w-4" />
        </button>
        {isSelected ? <button
          className="grid h-7 w-7 place-items-center rounded-full bg-red-600 text-white shadow-sm"
          title="Remove box"
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            onRemoveBox(box.id);
          }}
        >
          <X className="h-3.5 w-3.5" />
        </button> : null}
      </div> : null}
      <div
        className="absolute inset-2 grid gap-1"
        style={{
          gridTemplateColumns: `repeat(${slotColumns}, minmax(0, 1fr))`,
          gridTemplateRows: `repeat(${slotRows}, minmax(0, 1fr))`,
        }}
      >
        {Array.from({ length: slotCount }, (_, slot) => {
          const widget = box.widgets.find((item) => item.slot === slot);
          if (!showWidgetSlots && !widget) return null;

          return (
            <WidgetSlot
              key={`${box.id}-${slot}`}
              boxId={box.id}
              dropEnabled={showWidgetSlots}
              editable={editable}
              slot={slot}
              slotColumns={slotColumns}
              widget={widget}
              onRemoveWidget={onRemoveWidget}
              onResizeWidget={onResizeWidget}
              onUpdateWidget={onUpdateWidget}
            />
          );
        })}
        {widgetPlacementPreview ? (
          <div
            className={`pointer-events-none z-30 rounded border-2 border-dashed ${
              widgetPlacementPreview.valid ? "border-cj-blue bg-blue-100/70" : "border-red-500 bg-red-100/70"
            }`}
            style={{
              gridColumn: `${(widgetPlacementPreview.slot % slotColumns) + 1} / span ${widgetPlacementPreview.width}`,
              gridRow: `${Math.floor(widgetPlacementPreview.slot / slotColumns) + 1} / span ${widgetPlacementPreview.height}`,
            }}
          />
        ) : null}
      </div>
      {editable ? <>
      <button
        aria-label="Resize from left"
        className="absolute left-0 top-0 z-20 h-full w-2 cursor-ew-resize rounded-l-lg bg-transparent hover:bg-cj-blue/20"
        type="button"
        onPointerDown={(event) => startResize("left", event)}
      />
      <button
        aria-label="Resize from top"
        className="absolute left-0 top-0 z-20 h-2 w-full cursor-ns-resize rounded-t-lg bg-transparent hover:bg-cj-blue/20"
        type="button"
        onPointerDown={(event) => startResize("top", event)}
      />
      <button
        aria-label="Resize width"
        className="absolute right-0 top-0 z-20 h-full w-2 cursor-ew-resize rounded-r-lg bg-transparent hover:bg-cj-blue/20"
        type="button"
        onPointerDown={(event) => startResize("right", event)}
      />
      <button
        aria-label="Resize height"
        className="absolute bottom-0 left-0 z-20 h-2 w-full cursor-ns-resize rounded-b-lg bg-transparent hover:bg-cj-blue/20"
        type="button"
        onPointerDown={(event) => startResize("bottom", event)}
      />
      <button
        aria-label="Resize from top left"
        className="absolute left-0 top-0 z-30 h-4 w-4 cursor-nwse-resize rounded-tl-lg border-l-2 border-t-2 border-cj-blue/60 bg-white/80"
        type="button"
        onPointerDown={(event) => startResize("top-left", event)}
      />
      <button
        aria-label="Resize box"
        className="absolute bottom-0 right-0 z-30 h-4 w-4 cursor-nwse-resize rounded-br-lg border-b-2 border-r-2 border-cj-blue/60 bg-white/80"
        type="button"
        onPointerDown={(event) => startResize("corner", event)}
      />
      </> : null}
    </section>
  );
}

function BoxSkeletonPreview({ columns, label, rows }: { columns: number; label: string; rows: number }) {
  return (
    <div className="rounded-lg border-2 border-dashed border-cj-blue bg-blue-50/80 p-3 shadow-panel">
      <div className="mb-2 text-xs font-black uppercase text-cj-blue">{label}</div>
      <div className="grid gap-2" style={{ gridTemplateColumns: `repeat(${columns}, 44px)` }}>
        {Array.from({ length: columns * rows }, (_, index) => (
          <span key={index} className="h-10 rounded border border-dashed border-cj-blue/45 bg-white/70" />
        ))}
      </div>
    </div>
  );
}

function PlacementPreview({
  cell,
  columns,
  rows,
  valid,
}: {
  cell: number;
  columns: number;
  rows: number;
  valid: boolean;
}) {
  const position = getCellPosition(cell);

  return (
    <div
      className={`pointer-events-none z-20 rounded-lg border-2 border-dashed ${
        valid ? "border-cj-blue bg-blue-100/45" : "border-red-500 bg-red-100/45"
      }`}
      style={{
        gridColumn: `${position.column} / span ${columns}`,
        gridRow: `${position.row} / span ${rows}`,
      }}
    />
  );
}

function WidgetCursorPreview({ copy = false }: { copy?: boolean }) {
  const Icon = copy ? Copy : Box;

  return (
    <div className="grid h-11 w-11 place-items-center rounded-md border-2 border-cj-blue bg-white text-cj-blue shadow-panel">
      <Icon className="h-5 w-5" />
    </div>
  );
}

export function BoxWorkspaceEditor({
  cancelRequestId = 0,
  editable,
  onCancelComplete,
  onSaveComplete,
  saveRequestId = 0,
}: {
  cancelRequestId?: number;
  editable: boolean;
  onCancelComplete?: () => void;
  onSaveComplete?: (saved: boolean) => void;
  saveRequestId?: number;
}) {
  const [boxes, setBoxes] = useState<WorkspaceBox[]>(initialBoxes);
  const [pages, setPages] = useState<WorkspacePage[]>([{ id: "page-1", name: "Main", boxes: initialBoxes, isMain: true, locked: true }]);
  const [activePageId, setActivePageId] = useState("page-1");
  const [isPageMenuOpen, setIsPageMenuOpen] = useState(false);
  const [isCreatingPage, setIsCreatingPage] = useState(false);
  const [isToolTrayExpanded, setIsToolTrayExpanded] = useState(true);
  const [newPageName, setNewPageName] = useState("");
  const [editingPageId, setEditingPageId] = useState<string | null>(null);
  const [editingPageName, setEditingPageName] = useState("");
  const [activeDrag, setActiveDrag] = useState<DragData | null>(null);
  const [overCell, setOverCell] = useState<number | null>(null);
  const [overWidgetTarget, setOverWidgetTarget] = useState<{ boxId: string; slot: number } | null>(null);
  const [saveStatus, setSaveStatus] = useState("Ready");
  const [autoSlideEnabled, setAutoSlideEnabled] = useState(
    () => window.localStorage.getItem("cj-wms-auto-slide-enabled") === "true",
  );
  const [autoSlideRange, setAutoSlideRange] = useState(
    () => window.localStorage.getItem("cj-wms-auto-slide-range") ?? "",
  );
  const [autoSlideSeconds, setAutoSlideSeconds] = useState(
    () => Math.max(5, Number(window.localStorage.getItem("cj-wms-auto-slide-seconds")) || 10),
  );
  const [edgeNavigationEnabled, setEdgeNavigationEnabled] = useState(
    () => window.localStorage.getItem("cj-wms-edge-page-navigation") !== "false",
  );
  const [edgeDirection, setEdgeDirection] = useState<"previous" | "next" | null>(null);
  const [edgeCycle, setEdgeCycle] = useState(0);
  const latestBoxesRef = useRef(boxes);
  const pagesRef = useRef(pages);
  const activePageIdRef = useRef(activePageId);
  const previousEditableRef = useRef(editable);
  const processedCancelRequestRef = useRef(cancelRequestId);
  const processedSaveRequestRef = useRef(saveRequestId);
  const skipNextExitSaveRef = useRef(false);
  const editSnapshotRef = useRef<WorkspaceLayout | null>(null);
  const edgeDirectionRef = useRef<"previous" | "next" | null>(null);
  const edgeTimerRef = useRef<number | null>(null);
  const pageMenuRef = useRef<HTMLDivElement>(null);
  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 6 } }));
  const occupiedCells = useMemo(() => new Set(boxes.flatMap((box) => getBoxCells(box))), [boxes]);
  const autoSlidePageIndexes = useMemo(() => parsePageRange(autoSlideRange, pages.length), [autoSlideRange, pages.length]);
  latestBoxesRef.current = boxes;
  pagesRef.current = pages;
  activePageIdRef.current = activePageId;

  useEffect(() => {
    window.localStorage.setItem("cj-wms-edge-page-navigation", String(edgeNavigationEnabled));
  }, [edgeNavigationEnabled]);

  useEffect(() => {
    window.localStorage.setItem("cj-wms-auto-slide-enabled", String(autoSlideEnabled));
    window.localStorage.setItem("cj-wms-auto-slide-range", autoSlideRange);
    window.localStorage.setItem("cj-wms-auto-slide-seconds", String(autoSlideSeconds));
  }, [autoSlideEnabled, autoSlideRange, autoSlideSeconds]);

  useEffect(() => {
    if (saveRequestId <= processedSaveRequestRef.current) return;
    processedSaveRequestRef.current = saveRequestId;
    void handleSaveLayout().then((saved) => {
      if (saved) skipNextExitSaveRef.current = true;
      onSaveComplete?.(saved);
    });
  }, [onSaveComplete, saveRequestId]);

  useEffect(() => {
    if (cancelRequestId <= processedCancelRequestRef.current) return;
    processedCancelRequestRef.current = cancelRequestId;
    const snapshot = editSnapshotRef.current;
    if (snapshot) {
      const snapshotPages = snapshot.pages ?? [];
      const snapshotActivePageId = snapshot.activePageId ?? snapshotPages[0]?.id ?? "page-1";
      skipNextExitSaveRef.current = true;
      pagesRef.current = snapshotPages;
      activePageIdRef.current = snapshotActivePageId;
      latestBoxesRef.current = snapshot.boxes;
      setPages(snapshotPages);
      setActivePageId(snapshotActivePageId);
      setBoxes(snapshot.boxes);
    }
    onCancelComplete?.();
  }, [cancelRequestId, onCancelComplete]);

  useEffect(() => {
    fetchWorkspaceLayout()
      .then((layout) => {
        const rawPages = layout.pages && layout.pages.length > 0
          ? layout.pages
          : [{ id: "page-1", name: "Main", boxes: layout.boxes.length > 0 ? layout.boxes : initialBoxes, isMain: true, locked: true }];
        const loadedPages = rawPages.map((page, index) => {
          const isMain = page.isMain ?? index === 0;
          return { ...page, isMain, locked: isMain ? true : (page.locked ?? false) };
        });
        const loadedActivePageId = loadedPages.some((page) => page.id === layout.activePageId)
          ? layout.activePageId!
          : loadedPages[0].id;
        setPages(loadedPages);
        setActivePageId(loadedActivePageId);
        setBoxes(loadedPages.find((page) => page.id === loadedActivePageId)?.boxes ?? []);
        setSaveStatus("Loaded from database");
      })
      .catch(() => setSaveStatus("Database layout unavailable"));
  }, []);

  useEffect(() => {
    const workspaceEvents = new EventSource("/api/workspace/events");

    async function patchWorkspace() {
      if (editable) return;
      try {
        const layout = await fetchWorkspaceLayout();
        const rawPages = layout.pages && layout.pages.length > 0
          ? layout.pages
          : [{ id: "page-1", name: "Main", boxes: layout.boxes, isMain: true, locked: true }];
        const incomingPages = rawPages.map((page, index) => {
          const isMain = page.isMain ?? index === 0;
          return { ...page, isMain, locked: isMain ? true : (page.locked ?? false) };
        });
        const currentPagesById = new Map(pagesRef.current.map((page) => [page.id, page]));
        const mergedPages = incomingPages.map((incomingPage) => {
          const currentPage = currentPagesById.get(incomingPage.id);
          const mergedBoxes = mergeRemoteBoxes(currentPage?.boxes ?? [], incomingPage.boxes);
          if (currentPage && JSON.stringify({ ...currentPage, boxes: [] }) === JSON.stringify({ ...incomingPage, boxes: [] }) && mergedBoxes === currentPage.boxes) {
            return currentPage;
          }
          return { ...incomingPage, boxes: mergedBoxes };
        });
        const currentActivePageId = activePageIdRef.current;
        const nextActivePage = mergedPages.find((page) => page.id === currentActivePageId)
          ?? mergedPages.find((page) => page.id === layout.activePageId)
          ?? mergedPages[0];
        if (!nextActivePage) return;

        pagesRef.current = mergedPages;
        activePageIdRef.current = nextActivePage.id;
        latestBoxesRef.current = nextActivePage.boxes;
        setPages(mergedPages);
        setActivePageId(nextActivePage.id);
        setBoxes(nextActivePage.boxes);
        setSaveStatus("Updated live");
      } catch {
        setSaveStatus("Live update unavailable");
      }
    }

    function refreshExcelWidgets(event: MessageEvent<string>) {
      window.dispatchEvent(new CustomEvent("excel-upload-replaced", { detail: { uploadId: event.data } }));
    }

    workspaceEvents.addEventListener("workspace-layout-saved", patchWorkspace);
    workspaceEvents.addEventListener("excel-upload-replaced", refreshExcelWidgets as EventListener);
    return () => {
      workspaceEvents.removeEventListener("workspace-layout-saved", patchWorkspace);
      workspaceEvents.removeEventListener("excel-upload-replaced", refreshExcelWidgets as EventListener);
      workspaceEvents.close();
    };
  }, [editable]);

  useEffect(() => {
    const wasEditable = previousEditableRef.current;
    previousEditableRef.current = editable;
    if (!wasEditable && editable) {
      editSnapshotRef.current = {
        activePageId: activePageIdRef.current,
        boxes: latestBoxesRef.current,
        pages: pagesRef.current.map((page) => (
          page.id === activePageIdRef.current ? { ...page, boxes: latestBoxesRef.current } : page
        )),
      };
      return;
    }
    if (!wasEditable || editable) return;
    if (skipNextExitSaveRef.current) {
      skipNextExitSaveRef.current = false;
      setSaveStatus("Changes discarded");
      return;
    }

    const activeId = activePageIdRef.current;
    const currentBoxes = latestBoxesRef.current;
    const syncedPages = pagesRef.current.map((page) => (
      page.id === activeId ? { ...page, boxes: currentBoxes } : page
    ));

    pagesRef.current = syncedPages;
    setPages(syncedPages);
    setSaveStatus("Saving...");
    saveWorkspaceLayout({ boxes: currentBoxes, pages: syncedPages, activePageId: activeId })
      .then(() => setSaveStatus("Saved to database"))
      .catch(() => setSaveStatus("Save failed: database unavailable"));
  }, [editable]);

  useEffect(() => {
    if (!isPageMenuOpen) return;

    function handleOutsidePointer(event: globalThis.PointerEvent) {
      if (!pageMenuRef.current?.contains(event.target as Node)) {
        setIsPageMenuOpen(false);
        setIsCreatingPage(false);
        setNewPageName("");
      }
    }

    document.addEventListener("pointerdown", handleOutsidePointer);
    return () => document.removeEventListener("pointerdown", handleOutsidePointer);
  }, [isPageMenuOpen]);

  useEffect(() => {
    const edgeWidth = 10;

    function clearEdgeDwell() {
      if (edgeTimerRef.current === null && edgeDirectionRef.current === null) return;
      if (edgeTimerRef.current !== null) {
        window.clearTimeout(edgeTimerRef.current);
        edgeTimerRef.current = null;
      }
      edgeDirectionRef.current = null;
      setEdgeDirection(null);
    }

    function schedulePageChange(direction: "previous" | "next") {
      edgeTimerRef.current = window.setTimeout(() => {
        const currentPages = pagesRef.current;
        if (currentPages.length <= 1) {
          clearEdgeDwell();
          return;
        }

        const currentIndex = Math.max(0, currentPages.findIndex((page) => page.id === activePageIdRef.current));
        const offset = direction === "next" ? 1 : -1;
        const targetIndex = (currentIndex + offset + currentPages.length) % currentPages.length;
        const syncedPages = currentPages.map((page) => (
          page.id === activePageIdRef.current ? { ...page, boxes: latestBoxesRef.current } : page
        ));
        const targetPage = syncedPages[targetIndex];

        pagesRef.current = syncedPages;
        activePageIdRef.current = targetPage.id;
        setPages(syncedPages);
        setActivePageId(targetPage.id);
        setBoxes(targetPage.boxes);
        setEdgeCycle((current) => current + 1);
        schedulePageChange(direction);
      }, 1000);
    }

    function startEdgeDwell(direction: "previous" | "next") {
      if (edgeDirectionRef.current === direction && edgeTimerRef.current !== null) return;
      if (edgeTimerRef.current !== null) window.clearTimeout(edgeTimerRef.current);
      edgeDirectionRef.current = direction;
      setEdgeDirection(direction);
      setEdgeCycle((current) => current + 1);
      schedulePageChange(direction);
    }

    function handlePointerMove(event: globalThis.PointerEvent) {
      if (event.pointerType === "touch") {
        clearEdgeDwell();
        return;
      }
      if (event.clientX <= edgeWidth) {
        startEdgeDwell("previous");
      } else if (event.clientX >= window.innerWidth - edgeWidth) {
        startEdgeDwell("next");
      } else {
        clearEdgeDwell();
      }
    }

    if (!edgeNavigationEnabled || editable || pages.length <= 1) {
      clearEdgeDwell();
      return;
    }

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("blur", clearEdgeDwell);
    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("blur", clearEdgeDwell);
      if (edgeTimerRef.current !== null) window.clearTimeout(edgeTimerRef.current);
      edgeTimerRef.current = null;
    };
  }, [edgeNavigationEnabled, editable, pages.length]);

  useEffect(() => {
    if (!autoSlideEnabled || editable || autoSlidePageIndexes.length <= 1) return;

    const slideTimer = window.setInterval(() => {
      const currentPages = pagesRef.current;
      const slidePages = autoSlidePageIndexes
        .map((index) => currentPages[index])
        .filter((page): page is WorkspacePage => Boolean(page));
      if (slidePages.length <= 1) return;

      const currentSlideIndex = slidePages.findIndex((page) => page.id === activePageIdRef.current);
      const targetPage = slidePages[(currentSlideIndex + 1 + slidePages.length) % slidePages.length];
      const syncedPages = currentPages.map((page) => (
        page.id === activePageIdRef.current ? { ...page, boxes: latestBoxesRef.current } : page
      ));
      const syncedTargetPage = syncedPages.find((page) => page.id === targetPage.id) ?? targetPage;

      pagesRef.current = syncedPages;
      activePageIdRef.current = syncedTargetPage.id;
      latestBoxesRef.current = syncedTargetPage.boxes;
      setPages(syncedPages);
      setActivePageId(syncedTargetPage.id);
      setBoxes(syncedTargetPage.boxes);
    }, autoSlideSeconds * 1000);

    return () => window.clearInterval(slideTimer);
  }, [autoSlideEnabled, autoSlidePageIndexes, autoSlideSeconds, editable]);

  function removeBox(boxId: string) {
    setBoxes((current) => current.filter((box) => box.id !== boxId));
  }

  function removeWidget(boxId: string, widgetId: string) {
    setBoxes((current) =>
      current.map((box) => (box.id === boxId ? { ...box, widgets: box.widgets.filter((widget) => widget.id !== widgetId) } : box)),
    );
  }

  function updateWidget(boxId: string, widgetId: string, changes: Partial<WorkspaceWidget>) {
    setBoxes((current) => current.map((box) => (
      box.id === boxId
        ? { ...box, widgets: box.widgets.map((widget) => widget.id === widgetId ? { ...widget, ...changes } : widget) }
        : box
    )));
  }

  function resizeWidget(boxId: string, widgetId: string, slot: number, width: number, height: number) {
    const box = latestBoxesRef.current.find((item) => item.id === boxId);
    const widget = box?.widgets.find((item) => item.id === widgetId);
    if (!box || !widget) return;

    const resizedWidget = { ...widget, slot, width, height };
    if (!canPlaceWidget(box, resizedWidget, widgetId)) return;
    updateWidget(boxId, widgetId, { slot, width, height });
  }

  function putWidget(boxId: string, slot: number, widget: WorkspaceWidget) {
    setBoxes((current) =>
      current.map((box) => {
        if (box.id !== boxId) return box;
        const placement = findWidgetPlacement(box, widget, slot, widget.id);
        if (placement === null) return box;
        const nextWidget = { ...widget, slot: placement };
        return { ...box, widgets: [...box.widgets.filter((item) => item.id !== widget.id), nextWidget] };
      }),
    );
  }

  function moveWidget(sourceBoxId: string, targetBoxId: string, slot: number, widget: WorkspaceWidget) {
    setBoxes((current) => {
      const targetBox = current.find((box) => box.id === targetBoxId);
      if (!targetBox) return current;
      const placement = findWidgetPlacement(targetBox, widget, slot, widget.id);
      if (placement === null) return current;
      const nextWidget = { ...widget, slot: placement };

      return current.map((box) => {
        if (box.id === sourceBoxId && sourceBoxId !== targetBoxId) {
          return { ...box, widgets: box.widgets.filter((item) => item.id !== widget.id) };
        }
        if (box.id === targetBoxId) {
          return { ...box, widgets: [...box.widgets.filter((item) => item.id !== widget.id), nextWidget] };
        }
        return box;
      });
    });
  }

  function resizeBox(boxId: string, columns: number, rows: number, cell?: number) {
    const currentBox = latestBoxesRef.current.find((box) => box.id === boxId);
    if (!currentBox) return;

    const resizedBox = { ...currentBox, columns, rows, cell: cell ?? currentBox.cell };
    if (!canPlaceBox(latestBoxesRef.current, resizedBox, boxId)) return;

    setBoxes((current) => current.map((box) => (box.id === boxId ? resizedBox : box)));
  }

  function getSyncedPages() {
    return pages.map((page) => page.id === activePageId ? { ...page, boxes: latestBoxesRef.current } : page);
  }

  function switchPage(pageId: string) {
    const syncedPages = getSyncedPages();
    const targetPage = syncedPages.find((page) => page.id === pageId);
    if (!targetPage) return;

    setPages(syncedPages);
    setActivePageId(pageId);
    setBoxes(targetPage.boxes);
    setIsPageMenuOpen(false);
  }

  function reorderPages(event: DragEndEvent) {
    if (!event.over || event.active.id === event.over.id) return;
    const syncedPages = getSyncedPages();
    const oldIndex = syncedPages.findIndex((page) => page.id === event.active.id);
    const newIndex = syncedPages.findIndex((page) => page.id === event.over?.id);
    if (oldIndex < 0 || newIndex < 0) return;
    setPages(arrayMove(syncedPages, oldIndex, newIndex));
    setSaveStatus("Page order changed - save layout");
  }

  function startRenamingPage(page: WorkspacePage) {
    setEditingPageId(page.id);
    setEditingPageName(page.name);
  }

  function commitPageName() {
    const nextName = editingPageName.trim();
    if (!editingPageId || !nextName) return;
    setPages((current) => current.map((page) => page.id === editingPageId ? { ...page, name: nextName } : page));
    setEditingPageId(null);
    setEditingPageName("");
    setSaveStatus("Page name changed - save layout");
  }

  function cancelPageRename() {
    setEditingPageId(null);
    setEditingPageName("");
  }

  function createPage() {
    const name = newPageName.trim();
    if (!name) return;

    const syncedPages = getSyncedPages();
    const newPage: WorkspacePage = { id: makeId("page"), name, boxes: [], isMain: false, locked: false };
    setPages([...syncedPages, newPage]);
    setActivePageId(newPage.id);
    setBoxes([]);
    setNewPageName("");
    setIsCreatingPage(false);
    setIsPageMenuOpen(false);
    setSaveStatus("New page ready to save");
  }

  function togglePageLock(pageId: string) {
    setPages((current) => current.map((page) => (
      page.id === pageId && !page.isMain ? { ...page, locked: !page.locked } : page
    )));
    setSaveStatus("Page lock changed - save layout");
  }

  function deletePage(pageId: string) {
    const syncedPages = getSyncedPages();
    const targetPage = syncedPages.find((page) => page.id === pageId);
    if (!targetPage || targetPage.isMain || targetPage.locked) return;

    const remainingPages = syncedPages.filter((page) => page.id !== pageId);
    setPages(remainingPages);
    if (pageId === activePageId) {
      const fallbackPage = remainingPages.find((page) => page.isMain) ?? remainingPages[0];
      setActivePageId(fallbackPage.id);
      setBoxes(fallbackPage.boxes);
    }
    setSaveStatus("Page deleted - save layout");
  }

  async function handleSaveLayout() {
    setSaveStatus("Saving...");
    try {
      const syncedPages = getSyncedPages();
      const savedLayout: WorkspaceLayout = await saveWorkspaceLayout({ boxes, pages: syncedPages, activePageId });
      const savedPages = savedLayout.pages && savedLayout.pages.length > 0 ? savedLayout.pages : syncedPages;
      setPages(savedPages);
      setBoxes(savedPages.find((page) => page.id === activePageId)?.boxes ?? savedLayout.boxes);
      setSaveStatus("Saved to database");
      return true;
    } catch {
      setSaveStatus("Save failed: database unavailable");
      return false;
    }
  }

  function handleDragStart(event: DragStartEvent) {
    setActiveDrag((event.active.data.current as DragData | undefined) ?? null);
    setOverCell(null);
    setOverWidgetTarget(null);
  }

  function handleDragOver(event: DragOverEvent) {
    const overData = event.over?.data.current as { kind?: string; cell?: number; boxId?: string; pageId?: string; slot?: number } | undefined;
    setOverCell(overData?.kind === "workspace-cell" && typeof overData.cell === "number" ? overData.cell : null);
    setOverWidgetTarget(
      overData?.kind === "box-slot" && overData.boxId && typeof overData.slot === "number"
        ? { boxId: overData.boxId, slot: overData.slot }
        : null,
    );
  }

  function handleDragEnd(event: DragEndEvent) {
    const dragData = event.active.data.current as DragData | undefined;
    const overData = event.over?.data.current as { kind?: string; cell?: number; boxId?: string; pageId?: string; slot?: number } | undefined;
    setActiveDrag(null);
    setOverCell(null);
    setOverWidgetTarget(null);

    if (!dragData || !overData) return;

    if (dragData.kind === "box-template" && overData.kind === "workspace-cell" && typeof overData.cell === "number") {
      const nextBox: WorkspaceBox = {
        id: makeId("box"),
        size: dragData.template.size,
        title: dragData.template.label,
        cell: overData.cell,
        columns: dragData.template.columns,
        rows: dragData.template.rows,
        widgets: [],
      };
      if (canPlaceBox(boxes, nextBox)) {
        setBoxes((current) => [...current, nextBox]);
      }
      return;
    }

    if (dragData.kind === "box" && overData.kind === "workspace-cell" && typeof overData.cell === "number") {
      const movedBox = { ...dragData.box, cell: overData.cell };
      if (canPlaceBox(boxes, movedBox, dragData.box.id)) {
        setBoxes((current) => current.map((box) => (box.id === dragData.box.id ? movedBox : box)));
      }
      return;
    }

    if (dragData.kind === "widget-copy" && overData.kind === "page-copy-target" && overData.pageId) {
      const syncedPages = getSyncedPages();
      const targetPage = syncedPages.find((page) => page.id === overData.pageId);
      const sourceBox = latestBoxesRef.current.find((box) => box.id === dragData.sourceBoxId);
      if (!targetPage) return;

      const orderedBoxes = [...targetPage.boxes].sort((left, right) => {
        const leftPreferred = sourceBox && left.cell === sourceBox.cell ? 1 : 0;
        const rightPreferred = sourceBox && right.cell === sourceBox.cell ? 1 : 0;
        return rightPreferred - leftPreferred;
      });
      const copiedWidget = duplicateWorkspaceWidget(dragData.widget, dragData.widget.slot);
      let targetBoxId: string | null = null;
      let targetSlot: number | null = null;
      for (const targetBox of orderedBoxes) {
        const placement = findWidgetPlacement(targetBox, copiedWidget, dragData.widget.slot);
        if (placement !== null) {
          targetBoxId = targetBox.id;
          targetSlot = placement;
          break;
        }
      }
      if (!targetBoxId || targetSlot === null) {
        setSaveStatus(`Cannot copy to ${targetPage.name}: no available Box space`);
        return;
      }

      const nextPages = syncedPages.map((page) => page.id === targetPage.id
        ? {
            ...page,
            boxes: page.boxes.map((box) => box.id === targetBoxId
              ? { ...box, widgets: [...box.widgets, { ...copiedWidget, slot: targetSlot! }] }
              : box),
          }
        : page);
      setPages(nextPages);
      if (targetPage.id === activePageId) {
        const nextBoxes = nextPages.find((page) => page.id === activePageId)?.boxes ?? latestBoxesRef.current;
        setBoxes(nextBoxes);
      }
      setSaveStatus(`Widget copied to ${targetPage.name} - save layout`);
      return;
    }

    if (overData.kind === "box-slot" && overData.boxId && typeof overData.slot === "number") {
      if (dragData.kind === "widget-template") {
        putWidget(overData.boxId, overData.slot, {
          id: makeId("widget"),
          type: dragData.template.type,
          label: dragData.template.label,
          slot: overData.slot,
          width: isLargeDataWidget(dragData.template.type) ? 4 : dragData.template.type === "gradient-color" ? 4 : dragData.template.type === "title" || dragData.template.type === "text" || dragData.template.type === "text-query" ? 3 : 1,
          height: isLargeDataWidget(dragData.template.type) ? 4 : dragData.template.type === "gradient-color" ? 2 : 1,
          content: dragData.template.label,
          fontSize: dragData.template.type === "title" ? 20 : 16,
          fontFamily: "Inter",
          fontWeight: 700,
          fontStyle: "normal",
          textColor: "#122033",
          backgroundColor: "#ffffff",
          useBackgroundColor: false,
          textAlign: "center",
          verticalAlign: "center",
          textQueryCell: dragData.template.type === "text-query" ? "A1" : undefined,
          iconName: dragData.template.type === "icon" ? "package" : undefined,
          iconColor: dragData.template.type === "icon" ? "#1473e6" : undefined,
          iconSize: dragData.template.type === "icon" ? 32 : undefined,
          gradientStartColor: dragData.template.type === "gradient-color" ? "#0080c6" : undefined,
          gradientEndColor: dragData.template.type === "gradient-color" ? "#e42f44" : undefined,
          gradientDirection: dragData.template.type === "gradient-color" ? "to right" : undefined,
          gradientStartPosition: dragData.template.type === "gradient-color" ? 0 : undefined,
          gradientEndPosition: dragData.template.type === "gradient-color" ? 100 : undefined,
          gradientStops: dragData.template.type === "gradient-color"
            ? [
                { id: makeId("gradient-stop"), color: "#0080c6", position: 0 },
                { id: makeId("gradient-stop"), color: "#e42f44", position: 100 },
              ]
            : undefined,
          gradientOpacity: dragData.template.type === "gradient-color" ? 100 : undefined,
          gradientBorderRadius: dragData.template.type === "gradient-color" ? 8 : undefined,
          barItems: dragData.template.type === "bar"
            ? [{ id: makeId("bar-item"), label: "Bar 1", cell: "C7", markerCell: "100" }]
            : dragData.template.type === "bar-markers"
              ? [
                  { id: makeId("bar-marker"), label: "Actual 1", cell: "C7", markerCell: "D7" },
                  { id: makeId("bar-marker"), label: "Actual 2", cell: "C8", markerCell: "D8" },
                ]
            : dragData.template.type === "simple-pie" || dragData.template.type === "simple-donut"
              ? [
                  { id: makeId("pie-slice"), label: "Complete", cell: "C7" },
                  { id: makeId("pie-slice"), label: "Pending", cell: "D7" },
                ]
              : undefined,
          barMax: dragData.template.type === "bar" || dragData.template.type === "bar-markers" || dragData.template.type === "stack-bar" || dragData.template.type === "stack-column" || dragData.template.type === "column-rotated-labels" || dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" || dragData.template.type === "radar-polygon" ? (dragData.template.type === "stack-column" ? 10000 : 100) : undefined,
          barMaxInput: dragData.template.type === "bar" || dragData.template.type === "bar-markers" || dragData.template.type === "stack-bar" || dragData.template.type === "stack-column" || dragData.template.type === "column-rotated-labels" || dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" || dragData.template.type === "radar-polygon" ? (dragData.template.type === "stack-column" ? "10000" : "100") : undefined,
          barDisplayPercentage: dragData.template.type === "bar" ? false : undefined,
          barMarkerColor: dragData.template.type === "bar-markers" ? "#e42f44" : undefined,
          barMarkerHeight: dragData.template.type === "bar-markers" ? 5 : undefined,
          barMarkerWidth: dragData.template.type === "bar-markers" ? 3 : undefined,
          barMarkerShowValue: dragData.template.type === "bar-markers" ? true : undefined,
          barMarkerFontSize: dragData.template.type === "bar-markers" ? 10 : undefined,
          barMarkerShowLabel: dragData.template.type === "bar-markers" ? true : undefined,
          barMarkerLabelText: dragData.template.type === "bar-markers" ? "Target" : undefined,
          barMarkerLabelFontSize: dragData.template.type === "bar-markers" ? 9 : undefined,
          barMarkerLabelTextColor: dragData.template.type === "bar-markers" ? "#ffffff" : undefined,
          barMarkerLabelOffsetY: dragData.template.type === "bar-markers" ? 4 : undefined,
          barMarkerDisplayPercentage: dragData.template.type === "bar-markers" ? false : undefined,
          barMarkerZoneEnabled: dragData.template.type === "bar-markers" ? false : undefined,
          barMarkerZoneLowEnd: dragData.template.type === "bar-markers" ? 50 : undefined,
          barMarkerZoneMidEnd: dragData.template.type === "bar-markers" ? 80 : undefined,
          barMarkerZoneLowColor: dragData.template.type === "bar-markers" ? "#fee2e2" : undefined,
          barMarkerZoneMidColor: dragData.template.type === "bar-markers" ? "#fef3c7" : undefined,
          barMarkerZoneHighColor: dragData.template.type === "bar-markers" ? "#dcfce7" : undefined,
          barMarkerZoneOpacity: dragData.template.type === "bar-markers" ? 32 : undefined,
          barBorderRadius: dragData.template.type === "bar" || dragData.template.type === "bar-markers" || dragData.template.type === "stack-bar" || dragData.template.type === "stack-column" || dragData.template.type === "stack-100-bar" || dragData.template.type === "stack-100-column" || dragData.template.type === "column-rotated-labels" ? 6 : undefined,
          stackValueRounding: dragData.template.type === "stack-column" ? 10000 : undefined,
          stackColumnPercentage: dragData.template.type === "stack-column" ? false : undefined,
          radarStrokeWidth: dragData.template.type === "radar-polygon" ? 2 : undefined,
          radarFillOpacity: dragData.template.type === "radar-polygon" ? 0.2 : undefined,
          radarMarkerSize: dragData.template.type === "radar-polygon" ? 4 : undefined,
          radarPolygonColor1: dragData.template.type === "radar-polygon" ? "#f8fafc" : undefined,
          radarPolygonColor2: dragData.template.type === "radar-polygon" ? "#eef2f6" : undefined,
          radarPolygonStrokeColor: dragData.template.type === "radar-polygon" ? "#cbd5e1" : undefined,
          stackCategories: dragData.template.type === "stack-bar" || dragData.template.type === "stack-column" || dragData.template.type === "stack-100-bar" || dragData.template.type === "stack-100-column" || dragData.template.type === "column-rotated-labels" || dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" || dragData.template.type === "radar-polygon"
            ? dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" || dragData.template.type === "column-rotated-labels" || dragData.template.type === "stack-column" || dragData.template.type === "stack-100-column" || dragData.template.type === "radar-polygon"
              ? [
                  { id: makeId("chart-category"), label: "Week 1" },
                  { id: makeId("chart-category"), label: "Week 2" },
                  { id: makeId("chart-category"), label: "Week 3" },
                ]
              : [{ id: makeId("stack-category"), label: "Bar 1" }]
            : undefined,
          stackSeries: dragData.template.type === "stack-bar" || dragData.template.type === "stack-column" || dragData.template.type === "stack-100-bar" || dragData.template.type === "stack-100-column" || dragData.template.type === "column-rotated-labels" || dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" || dragData.template.type === "radar-polygon"
            ? dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" || dragData.template.type === "column-rotated-labels" || dragData.template.type === "stack-column" || dragData.template.type === "stack-100-column" || dragData.template.type === "radar-polygon"
              ? [
                  { id: makeId("chart-series"), label: "Complete", cells: ["C7", "C8", "C9"] },
                  { id: makeId("chart-series"), label: "Pending", cells: ["D7", "D8", "D9"] },
                ]
              : [
                  { id: makeId("stack-series"), label: "Complete", cells: ["C7"] },
                  { id: makeId("stack-series"), label: "Pending", cells: ["D7"] },
                ]
            : undefined,
          lineCurve: dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" ? "smooth" : undefined,
          lineStrokeWidth: dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" ? 3 : undefined,
          lineShowMarkers: dragData.template.type === "basic-line" || dragData.template.type === "line-annotations" ? true : undefined,
          lineNullMissing: dragData.template.type === "basic-line" ? false : undefined,
          lineAnnotations: dragData.template.type === "line-annotations"
            ? [{ id: makeId("line-annotation"), label: "Target", axis: "y", value: "80", color: "#e42f44" }]
            : undefined,
          columnLabelRotation: dragData.template.type === "column-rotated-labels" ? -45 : undefined,
          columnWidth: dragData.template.type === "column-rotated-labels" ? 55 : undefined,
          columnBorderRadius: dragData.template.type === "column-rotated-labels" ? 6 : undefined,
          columnShowDataLabels: dragData.template.type === "column-rotated-labels" ? true : undefined,
          columnShowDataLabelBackground: dragData.template.type === "column-rotated-labels" ? true : undefined,
        });
      }

      if (dragData.kind === "widget") {
        moveWidget(dragData.boxId, overData.boxId, overData.slot, dragData.widget);
      }

      if (dragData.kind === "widget-copy") {
        putWidget(overData.boxId, overData.slot, duplicateWorkspaceWidget(dragData.widget, overData.slot));
      }
    }
  }

  const placementPreview = useMemo(() => {
    if (overCell === null) return null;

    if (activeDrag?.kind === "box-template") {
      const box = {
        cell: overCell,
        columns: activeDrag.template.columns,
        rows: activeDrag.template.rows,
      };
      return {
        ...box,
        valid: canPlaceBox(boxes, box),
      };
    }

    if (activeDrag?.kind === "box") {
      const box = {
        cell: overCell,
        columns: activeDrag.box.columns,
        rows: activeDrag.box.rows,
      };
      return {
        ...box,
        valid: canPlaceBox(boxes, box, activeDrag.box.id),
      };
    }

    return null;
  }, [activeDrag, boxes, overCell]);

  const widgetPlacementPreview = useMemo(() => {
    if (!overWidgetTarget || (activeDrag?.kind !== "widget-template" && activeDrag?.kind !== "widget" && activeDrag?.kind !== "widget-copy")) return null;
    const box = boxes.find((item) => item.id === overWidgetTarget.boxId);
    if (!box) return null;

    const widget: WorkspaceWidget = activeDrag.kind === "widget" || activeDrag.kind === "widget-copy"
      ? activeDrag.widget
      : {
          id: "widget-preview",
          type: activeDrag.template.type,
          label: activeDrag.template.label,
          slot: overWidgetTarget.slot,
          width: isLargeDataWidget(activeDrag.template.type) ? 4 : activeDrag.template.type === "gradient-color" ? 4 : activeDrag.template.type === "title" || activeDrag.template.type === "text" || activeDrag.template.type === "text-query" ? 3 : 1,
          height: isLargeDataWidget(activeDrag.template.type) ? 4 : activeDrag.template.type === "gradient-color" ? 2 : 1,
        };
    const footprint = getWidgetFootprint(widget);
    const placement = findWidgetPlacement(box, widget, overWidgetTarget.slot, activeDrag.kind === "widget" ? widget.id : undefined);
    const slotColumns = box.columns * slotsPerCellColumn;
    const requestedRow = Math.floor(overWidgetTarget.slot / slotColumns);
    const fallbackColumn = Math.max(0, Math.min(slotColumns - footprint.width, overWidgetTarget.slot % slotColumns));

    return {
      boxId: box.id,
      slot: placement ?? requestedRow * slotColumns + fallbackColumn,
      width: footprint.width,
      height: footprint.height,
      valid: placement !== null,
    };
  }, [activeDrag, boxes, overWidgetTarget]);

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={workspaceCollisionDetection}
      onDragEnd={handleDragEnd}
      onDragOver={handleDragOver}
      onDragStart={handleDragStart}
    >
      {edgeDirection ? (
        <div
          key={`${edgeDirection}-${edgeCycle}`}
          aria-hidden="true"
          className={`pointer-events-none fixed bottom-0 top-20 z-[65] w-1 origin-bottom animate-edge-dwell bg-cj-blue ${edgeDirection === "next" ? "right-0" : "left-0"}`}
        />
      ) : null}

      <div ref={pageMenuRef} className="fixed right-[7.5rem] top-[1.375rem] z-[70]">
        <div className="cj-logo-plate absolute right-12 top-0 hidden h-9 w-[132px] items-center justify-center rounded-md border border-slate-200 bg-white px-2 shadow-sm lg:flex" aria-label="CJ Logistics">
          <img className="block h-auto max-h-7 w-auto max-w-full object-contain" src="/cj-logistics-logo.png" alt="CJ Logistics" />
        </div>
        <button
          aria-label="Pages"
          className={`grid h-9 w-9 place-items-center rounded-md border bg-white transition ${isPageMenuOpen ? "border-cj-blue text-cj-blue" : "border-slate-200 text-slate-600 hover:border-cj-blue hover:text-cj-blue"}`}
          title="Pages"
          type="button"
          onClick={() => setIsPageMenuOpen((current) => !current)}
        >
          <Menu className="h-4 w-4" />
        </button>

        {isPageMenuOpen ? (
          <div className="absolute right-0 top-12 w-80 overflow-hidden rounded-md border border-slate-200 bg-white shadow-panel">
            <div className="flex items-center justify-between border-b border-slate-100 px-3 py-2">
              <div>
                <h2 className="text-sm font-black text-cj-navy">Pages</h2>
                <p className="text-[10px] font-bold text-slate-400">Choose or create a dashboard page.</p>
              </div>
              <button
                aria-label="Create page"
                className="grid h-8 w-8 place-items-center rounded-md bg-cj-navy text-white transition hover:bg-cj-blue"
                title="Create page"
                type="button"
                onClick={() => setIsCreatingPage(true)}
              >
                <Plus className="h-4 w-4" />
              </button>
            </div>

            <div className="border-b border-slate-100 p-2">
              <button
                aria-checked={edgeNavigationEnabled}
                className="flex w-full items-center justify-between gap-3 rounded-md bg-slate-50 px-3 py-2 text-left transition hover:bg-slate-100"
                role="switch"
                type="button"
                onClick={() => setEdgeNavigationEnabled((current) => !current)}
              >
                <span className="min-w-0">
                  <strong className="block text-xs font-black text-cj-navy">Change Pages</strong>
                  <span className="block text-[10px] font-bold text-slate-400">{edgeNavigationEnabled ? "On" : "Off"}</span>
                </span>
                <span className={`relative h-6 w-11 shrink-0 rounded-full transition ${edgeNavigationEnabled ? "bg-cj-blue" : "bg-slate-300"}`}>
                  <span className={`absolute top-1 h-4 w-4 rounded-full bg-white shadow-sm transition-transform ${edgeNavigationEnabled ? "translate-x-6" : "translate-x-1"}`} />
                </span>
              </button>
            </div>

            <div className="space-y-2 border-b border-slate-100 p-2">
              <button
                aria-checked={autoSlideEnabled}
                className="flex w-full items-center justify-between gap-3 rounded-md bg-slate-50 px-3 py-2 text-left transition hover:bg-slate-100"
                role="switch"
                type="button"
                onClick={() => setAutoSlideEnabled((current) => !current)}
              >
                <span className="min-w-0">
                  <strong className="block text-xs font-black text-cj-navy">Auto slide</strong>
                  <span className="block text-[10px] font-bold text-slate-400">{autoSlideEnabled ? "On" : "Off"}</span>
                </span>
                <span className={`relative h-6 w-11 shrink-0 rounded-full transition ${autoSlideEnabled ? "bg-cj-blue" : "bg-slate-300"}`}>
                  <span className={`absolute top-1 h-4 w-4 rounded-full bg-white shadow-sm transition-transform ${autoSlideEnabled ? "translate-x-6" : "translate-x-1"}`} />
                </span>
              </button>

              {autoSlideEnabled ? (
                <div className="grid grid-cols-[minmax(0,1fr)_92px] gap-2 px-1 pb-1">
                  <label className="text-[9px] font-black uppercase text-slate-500">
                    Page range
                    <input
                      aria-invalid={autoSlidePageIndexes.length === 0}
                      className={`mt-1 h-8 w-full rounded border bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none ${autoSlidePageIndexes.length === 0 ? "border-red-400 focus:border-red-500" : "border-slate-200 focus:border-cj-blue"}`}
                      placeholder="1-3 or 1, 3"
                      value={autoSlideRange}
                      onChange={(event) => setAutoSlideRange(event.target.value)}
                    />
                  </label>
                  <label className="text-[9px] font-black uppercase text-slate-500">
                    Interval
                    <select
                      className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-xs font-bold normal-case text-cj-navy outline-none focus:border-cj-blue"
                      value={autoSlideSeconds}
                      onChange={(event) => setAutoSlideSeconds(Number(event.target.value))}
                    >
                      {[5, 10, 15, 30, 60].map((seconds) => <option key={seconds} value={seconds}>{seconds}s</option>)}
                    </select>
                  </label>
                  <p className={`col-span-2 text-[9px] font-bold ${autoSlidePageIndexes.length === 0 ? "text-red-500" : "text-slate-400"}`}>
                    {autoSlidePageIndexes.length === 0 ? `Use page numbers from 1 to ${pages.length}.` : `${autoSlidePageIndexes.length} page${autoSlidePageIndexes.length === 1 ? "" : "s"} selected`}
                  </p>
                </div>
              ) : null}
            </div>

            <div className="max-h-64 overflow-y-auto p-2">
              <DndContext collisionDetection={closestCenter} onDragEnd={reorderPages}>
                <SortableContext items={pages.map((page) => page.id)} strategy={verticalListSortingStrategy}>
                  {pages.map((page, pageIndex) => (
                <SortablePageRow key={page.id} id={page.id} active={page.id === activePageId}>
                  {editingPageId === page.id ? (
                    <form className="flex min-w-0 flex-1 items-center gap-1" onSubmit={(event) => { event.preventDefault(); commitPageName(); }}>
                      <input
                        autoFocus
                        aria-label={`Rename ${page.name}`}
                        className="h-8 min-w-0 flex-1 rounded border border-cj-blue bg-white px-2 text-xs font-bold text-cj-navy outline-none"
                        maxLength={40}
                        value={editingPageName}
                        onChange={(event) => setEditingPageName(event.target.value)}
                        onKeyDown={(event) => {
                          if (event.key === "Escape") {
                            event.preventDefault();
                            cancelPageRename();
                          }
                        }}
                      />
                      <button className="grid h-7 w-7 shrink-0 place-items-center rounded text-emerald-600 hover:bg-emerald-50 disabled:opacity-30" disabled={!editingPageName.trim()} title="Confirm page name" type="submit">
                        <CheckCircle2 className="h-4 w-4" />
                      </button>
                      <button className="grid h-7 w-7 shrink-0 place-items-center rounded text-slate-400 hover:bg-slate-100 hover:text-red-600" title="Cancel rename" type="button" onClick={cancelPageRename}>
                        <X className="h-4 w-4" />
                      </button>
                    </form>
                  ) : (
                    <>
                      <button
                        className={`flex min-w-0 flex-1 items-center gap-2 px-2 py-1.5 text-left text-sm font-bold ${page.id === activePageId ? "text-cj-blue" : "text-slate-600"}`}
                        type="button"
                        onClick={() => switchPage(page.id)}
                      >
                        <span className="grid h-5 w-5 shrink-0 place-items-center rounded bg-slate-100 text-[9px] font-black text-slate-500">{pageIndex + 1}</span>
                        <span className="truncate">{page.name}</span>
                        {page.id === activePageId ? <span className="h-2 w-2 shrink-0 rounded-full bg-cj-blue" /> : null}
                      </button>
                      <button
                        aria-label={`Rename ${page.name}`}
                        className="grid h-7 w-7 shrink-0 place-items-center rounded text-slate-400 transition hover:bg-white hover:text-cj-blue"
                        title="Rename page"
                        type="button"
                        onClick={() => startRenamingPage(page)}
                      >
                        <Pencil className="h-3.5 w-3.5" />
                      </button>

                  {page.isMain ? (
                    <span className="grid h-7 w-7 shrink-0 place-items-center text-slate-400" title="Main page is permanently locked">
                      <Lock className="h-3.5 w-3.5" />
                    </span>
                  ) : (
                    <>
                      <button
                        aria-label={`Save ${page.name}`}
                        className="grid h-7 w-7 shrink-0 place-items-center rounded text-slate-500 transition hover:bg-white hover:text-cj-blue"
                        title="Save pages"
                        type="button"
                        onClick={handleSaveLayout}
                      >
                        <Save className="h-3.5 w-3.5" />
                      </button>
                      <button
                        aria-label={`${page.locked ? "Unlock" : "Lock"} ${page.name}`}
                        className={`grid h-7 w-7 shrink-0 place-items-center rounded transition hover:bg-white ${page.locked ? "text-amber-600" : "text-slate-400"}`}
                        title={page.locked ? "Unlock page" : "Lock page"}
                        type="button"
                        onClick={() => togglePageLock(page.id)}
                      >
                        {page.locked ? <Lock className="h-3.5 w-3.5" /> : <LockOpen className="h-3.5 w-3.5" />}
                      </button>
                      {!page.locked ? (
                        <button
                          aria-label={`Delete ${page.name}`}
                          className="grid h-7 w-7 shrink-0 place-items-center rounded text-red-500 transition hover:bg-red-50 hover:text-red-700"
                          title="Delete page"
                          type="button"
                          onClick={() => deletePage(page.id)}
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      ) : null}
                    </>
                  )}
                    </>
                  )}
                </SortablePageRow>
                  ))}
                </SortableContext>
              </DndContext>
            </div>

            {isCreatingPage ? (
              <form className="border-t border-slate-100 p-3" onSubmit={(event) => { event.preventDefault(); createPage(); }}>
                <label className="text-[10px] font-black uppercase text-slate-500" htmlFor="new-page-name">Page name</label>
                <div className="mt-1 flex gap-2">
                  <input
                    id="new-page-name"
                    autoFocus
                    className="h-9 min-w-0 flex-1 rounded border border-slate-200 px-3 text-sm font-bold text-cj-navy outline-none focus:border-cj-blue"
                    maxLength={40}
                    placeholder="New page"
                    value={newPageName}
                    onChange={(event) => setNewPageName(event.target.value)}
                  />
                  <button
                    className="h-9 rounded bg-cj-blue px-3 text-sm font-black text-white disabled:cursor-not-allowed disabled:opacity-40"
                    disabled={!newPageName.trim()}
                    type="submit"
                  >
                    Add
                  </button>
                </div>
              </form>
            ) : null}
          </div>
        ) : null}
      </div>

      <section className={`flex min-h-[calc(100vh-5rem)] flex-col gap-4 p-4 ${editable && isToolTrayExpanded ? "pb-64" : "pb-12"}`}>
        {editable ? <div className="fixed bottom-4 left-4 right-4 z-40">
          <button
            aria-expanded={isToolTrayExpanded}
            aria-label={isToolTrayExpanded ? "Collapse tools" : "Expand tools"}
            className="absolute -top-9 left-1/2 grid h-8 w-12 -translate-x-1/2 place-items-center rounded-t-md border border-b-0 border-slate-200 bg-white/95 text-slate-500 shadow-sm backdrop-blur-sm transition hover:text-cj-blue"
            title={isToolTrayExpanded ? "Collapse tools" : "Expand tools"}
            type="button"
            onClick={() => setIsToolTrayExpanded((expanded) => !expanded)}
          >
            {isToolTrayExpanded ? <ChevronDown className="h-5 w-5" /> : <ChevronUp className="h-5 w-5" />}
          </button>
          {isToolTrayExpanded ? <aside className="tool-tray-scroll grid max-h-[220px] grid-cols-[170px_160px_minmax(0,1fr)] items-start gap-3 overflow-x-hidden overflow-y-auto rounded-lg border border-slate-200 bg-white/95 p-3 shadow-panel backdrop-blur-sm">
          <section className="sticky top-0 border-r border-slate-200 pr-3">
            <div className="mb-2">
              <h2 className="text-sm font-black text-cj-navy">Layout</h2>
            </div>
            <button
              className="inline-flex h-8 w-full items-center justify-center gap-2 rounded-md bg-cj-navy px-2 text-xs font-black text-white transition hover:bg-cj-blue"
              type="button"
              onClick={handleSaveLayout}
            >
              <Save className="h-4 w-4" />
              Save Layout
            </button>
            <p className="mt-1 truncate text-[10px] font-bold text-slate-500" title={saveStatus}>{saveStatus}</p>
          </section>

          <section className="sticky top-0 border-r border-slate-200 pr-3">
            <div>
              <h2 className="text-sm font-black text-cj-navy">Box</h2>
            </div>
            <div className="mt-2 grid gap-1.5">
              {boxTemplates.map((template) => (
                <DraggableBoxTemplate key={template.size} template={template} />
              ))}
            </div>
          </section>

          <section className="min-w-0">
            <div>
              <h2 className="text-sm font-black text-cj-navy">Widget</h2>
            </div>
            <div className="mt-2 grid grid-cols-3 items-start gap-3">
              <div className="min-w-0">
                <h3 className="text-[10px] font-black uppercase text-slate-400">Content</h3>
                <div className="mt-1 grid gap-1.5">
                  {widgetTemplates.filter((template) => template.type === "title" || template.type === "text" || template.type === "text-query" || template.type === "icon" || template.type === "gradient-color").map((template) => (
                    <DraggableWidgetTemplate key={template.type} template={template} />
                  ))}
                </div>
              </div>
              <div className="min-w-0 border-l border-slate-200 pl-3">
                <h3 className="text-[10px] font-black uppercase text-slate-400">Visualization</h3>
                <div className="mt-1 grid grid-cols-2 gap-1.5">
                  {widgetTemplates.filter((template) => template.type === "chart" || template.type === "bar" || template.type === "bar-markers" || template.type === "stack-bar" || template.type === "stack-column" || template.type === "stack-100-bar" || template.type === "stack-100-column" || template.type === "column-rotated-labels" || template.type === "basic-line" || template.type === "line-annotations" || template.type === "radar-polygon" || template.type === "simple-pie" || template.type === "simple-donut").map((template) => (
                    <DraggableWidgetTemplate key={template.type} template={template} />
                  ))}
                </div>
              </div>
              <div className="min-w-0 border-l border-slate-200 pl-3">
                <h3 className="text-[10px] font-black uppercase text-slate-400">Data</h3>
                <div className="mt-1 grid gap-1.5">
                  {widgetTemplates.filter((template) => template.type === "excel-table").map((template) => (
                    <DraggableWidgetTemplate key={template.type} template={template} />
                  ))}
                </div>
              </div>
            </div>
          </section>
          </aside> : null}
        </div> : null}

        <div className={`relative order-1 min-h-[calc(100vh-7rem)] rounded-lg border p-3 ${editable ? "border-slate-200 bg-slate-100/70 shadow-panel" : "border-transparent bg-transparent"}`}>
          {editable ? <div
            className="absolute inset-3 grid gap-2"
            style={{
              gridTemplateColumns: `repeat(${workspaceColumns}, minmax(0, 1fr))`,
              gridTemplateRows: `repeat(${workspaceRows}, minmax(72px, 1fr))`,
            }}
          >
            {Array.from({ length: workspaceCellCount }, (_, cell) => (
              <WorkspaceCell key={cell} cell={cell} occupied={occupiedCells.has(cell)} />
            ))}
          </div> : null}

          <div
            className="absolute inset-3 grid gap-2"
            style={{
              gridTemplateColumns: `repeat(${workspaceColumns}, minmax(0, 1fr))`,
              gridTemplateRows: `repeat(${workspaceRows}, minmax(72px, 1fr))`,
            }}
          >
            {boxes.map((box) => (
              <WorkspaceBoxCard
                key={box.id}
                box={box}
                editable={editable}
                showWidgetSlots={activeDrag?.kind === "widget" || activeDrag?.kind === "widget-template" || activeDrag?.kind === "widget-copy"}
                widgetPlacementPreview={widgetPlacementPreview?.boxId === box.id ? widgetPlacementPreview : undefined}
                onRemoveBox={removeBox}
                onRemoveWidget={removeWidget}
                onResizeBox={resizeBox}
                onResizeWidget={resizeWidget}
                onUpdateWidget={updateWidget}
              />
            ))}
            {editable && placementPreview ? (
              <PlacementPreview
                cell={placementPreview.cell}
                columns={placementPreview.columns}
                rows={placementPreview.rows}
                valid={placementPreview.valid}
              />
            ) : null}
          </div>
        </div>
      </section>

      {activeDrag?.kind === "widget-copy" ? (
        <PageCopyPanel>
          <div className="mb-3 border-b border-slate-100 pb-2">
            <h2 className="text-xs font-black text-cj-navy">Copy Widget to Page</h2>
            <p className="mt-0.5 truncate text-[10px] font-bold text-slate-400">Dragging: {activeDrag.widget.label}</p>
          </div>
          <div className="min-h-0 flex-1 space-y-1.5 overflow-y-auto pr-1">
            {pages.filter((page) => page.id !== activePageId).map((page) => {
              const available = page.boxes.some((box) => findWidgetPlacement(box, activeDrag.widget, activeDrag.widget.slot) !== null);
              return <PageCopyDropTarget key={page.id} page={page} available={available} />;
            })}
            {pages.every((page) => page.id === activePageId) ? (
              <div className="rounded border border-dashed border-slate-300 px-3 py-5 text-center text-[10px] font-bold text-slate-400">Create another Page first</div>
            ) : null}
          </div>
        </PageCopyPanel>
      ) : null}

      <footer className="pointer-events-none relative z-10 -mt-16 translate-y-[3px] px-4 pb-4 text-center text-[10px] font-medium text-slate-400">
        © 2026 copyright reserved , CJ Logistics , IT Auttha.
      </footer>

      <DragOverlay>
        {activeDrag?.kind === "box-template" ? <BoxSkeletonPreview columns={activeDrag.template.columns} rows={activeDrag.template.rows} label={activeDrag.template.label} /> : null}
        {activeDrag?.kind === "box" ? <BoxSkeletonPreview columns={activeDrag.box.columns} rows={activeDrag.box.rows} label={activeDrag.box.title} /> : null}
        {activeDrag?.kind === "widget-template" ? <WidgetCursorPreview /> : null}
        {activeDrag?.kind === "widget" ? <WidgetCursorPreview /> : null}
        {activeDrag?.kind === "widget-copy" ? <WidgetCursorPreview copy /> : null}
      </DragOverlay>
    </DndContext>
  );
}
