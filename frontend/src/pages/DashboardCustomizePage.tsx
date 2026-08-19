import { CSSProperties, useEffect, useMemo, useState } from "react";
import {
  closestCenter,
  DndContext,
  DragEndEvent,
  DragOverlay,
  DragStartEvent,
  PointerSensor,
  useDraggable,
  useDroppable,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import {
  arrayMove,
  rectSortingStrategy,
  SortableContext,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { GripVertical, Plus, Save, Trash2 } from "lucide-react";

import { fetchDashboardWidgets, saveDashboardWidgets } from "../services/api";
import type {
  DashboardWidget,
  DashboardWidgetItem,
  DashboardWidgetItemType,
  DashboardWidgetSize,
} from "../types/dashboard";

type WidgetTemplate = {
  label: string;
  size: DashboardWidgetSize;
  description: string;
};

type ItemTemplate = {
  type: DashboardWidgetItemType;
  label: string;
  description: string;
};

type ActiveDrag =
  | { kind: "widget"; widget: DashboardWidget }
  | { kind: "palette-item"; template: ItemTemplate }
  | { kind: "widget-item"; item: DashboardWidgetItem; widgetId: string };

const widgetTemplates: WidgetTemplate[] = [
  { label: "1x1", size: "1x1", description: "Single KPI widget" },
  { label: "2x1", size: "2x1", description: "Wide chart or table" },
  { label: "2x2", size: "2x2", description: "Large operational panel" },
  { label: "1x2", size: "1x2", description: "Tall status list" },
];

const itemTemplates: ItemTemplate[] = [
  { type: "metric-total", label: "Total Workload", description: "All WMS tasks" },
  { type: "metric-inbound", label: "Inbound KPI", description: "Inbound progress" },
  { type: "metric-pick", label: "Pick KPI", description: "Pick progress" },
  { type: "metric-outbound", label: "Outbound KPI", description: "Outbound progress" },
  { type: "progress-chart", label: "Progress Chart", description: "Compare work groups" },
  { type: "workload-table", label: "Workload Table", description: "Operational rows" },
  { type: "last-upload-status", label: "Upload Status", description: "Last file update" },
];

function getGridDimensions(size: DashboardWidgetSize) {
  if (size === "2x2") return { columns: 4, rows: 4 };
  if (size === "2x1") return { columns: 4, rows: 2 };
  if (size === "1x2") return { columns: 2, rows: 4 };
  return { columns: 2, rows: 2 };
}

function getWidgetSpan(size: DashboardWidgetSize) {
  if (size === "2x2") return "md:col-span-2 md:row-span-2";
  if (size === "2x1") return "md:col-span-2";
  if (size === "1x2") return "md:row-span-2";
  return "";
}

function makeId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function PaletteItem({ template }: { template: ItemTemplate }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id: `palette:${template.type}`,
    data: { kind: "palette-item", template },
  });
  const style = transform ? { transform: CSS.Translate.toString(transform) } : undefined;

  return (
    <button
      ref={setNodeRef}
      className={`rounded-md border border-slate-200 bg-white p-3 text-left shadow-sm transition ${isDragging ? "opacity-40" : "hover:border-cj-blue"}`}
      style={style}
      type="button"
      {...attributes}
      {...listeners}
    >
      <strong className="block text-sm font-black text-cj-navy">{template.label}</strong>
      <span className="mt-1 block text-xs font-bold text-slate-500">{template.description}</span>
    </button>
  );
}

function DroppableSlot({
  item,
  slot,
  widgetId,
  onRemoveItem,
}: {
  item?: DashboardWidgetItem;
  slot: number;
  widgetId: string;
  onRemoveItem: (widgetId: string, itemId: string) => void;
}) {
  const { isOver, setNodeRef } = useDroppable({
    id: `slot:${widgetId}:${slot}`,
    data: { widgetId, slot },
  });
  const draggable = useDraggable({
    id: item ? `item:${widgetId}:${item.id}` : `empty:${widgetId}:${slot}`,
    disabled: !item,
    data: item ? { kind: "widget-item", item, widgetId } : undefined,
  });

  return (
    <div
      ref={setNodeRef}
      className={`relative grid min-h-16 place-items-center rounded-md border border-dashed p-2 text-center transition ${
        isOver ? "border-cj-blue bg-blue-50" : item ? "border-slate-200 bg-white" : "border-slate-200 bg-slate-50"
      }`}
    >
      {item ? (
        <div
          ref={draggable.setNodeRef}
          className="grid h-full w-full cursor-grab place-items-center rounded bg-white px-2 py-3 shadow-sm"
          style={draggable.transform ? { transform: CSS.Translate.toString(draggable.transform) } : undefined}
          {...draggable.attributes}
          {...draggable.listeners}
        >
          <button
            className="absolute right-1 top-1 grid h-5 w-5 place-items-center rounded bg-rose-50 text-rose-600"
            title="Remove item"
            type="button"
            onClick={(event) => {
              event.stopPropagation();
              onRemoveItem(widgetId, item.id);
            }}
          >
            <Trash2 className="h-3 w-3" />
          </button>
          <span className="text-xs font-black text-cj-navy">{item.label}</span>
        </div>
      ) : (
        <Plus className="h-4 w-4 text-slate-300" />
      )}
    </div>
  );
}

function SortableWidgetCard({
  widget,
  onRemoveItem,
}: {
  widget: DashboardWidget;
  onRemoveItem: (widgetId: string, itemId: string) => void;
}) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
    id: widget.id,
    data: { kind: "widget", widget },
  });
  const dimensions = getGridDimensions(widget.size);
  const slotCount = dimensions.columns * dimensions.rows;
  const style: CSSProperties = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  return (
    <section
      ref={setNodeRef}
      className={`rounded-lg border border-slate-200 bg-white p-4 shadow-panel ${getWidgetSpan(widget.size)} ${isDragging ? "opacity-40" : ""}`}
      style={style}
    >
      <header className="mb-3 flex items-center justify-between gap-3">
        <div>
          <h3 className="text-lg font-black text-cj-navy">{widget.title}</h3>
          <p className="text-xs font-bold uppercase text-slate-400">{widget.size}</p>
        </div>
        <button className="grid h-9 w-9 cursor-grab place-items-center rounded-md border border-slate-200 bg-slate-50 text-slate-500" type="button" {...attributes} {...listeners}>
          <GripVertical className="h-5 w-5" />
        </button>
      </header>
      <div
        className="grid gap-2"
        style={{
          gridTemplateColumns: `repeat(${dimensions.columns}, minmax(0, 1fr))`,
          gridTemplateRows: `repeat(${dimensions.rows}, minmax(52px, 1fr))`,
        }}
      >
        {Array.from({ length: slotCount }, (_, slot) => (
          <DroppableSlot
            key={`${widget.id}-${slot}`}
            item={widget.items.find((item) => item.slot === slot)}
            slot={slot}
            widgetId={widget.id}
            onRemoveItem={onRemoveItem}
          />
        ))}
      </div>
    </section>
  );
}

type DashboardCustomizePageProps = {
  onNavigate: (path: string) => void;
};

export function DashboardCustomizePage({ onNavigate }: DashboardCustomizePageProps) {
  const [widgets, setWidgets] = useState<DashboardWidget[]>([]);
  const [activeDrag, setActiveDrag] = useState<ActiveDrag | null>(null);
  const [status, setStatus] = useState("Loading layout...");
  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 6 } }));

  useEffect(() => {
    fetchDashboardWidgets()
      .then((layout) => {
        setWidgets(layout.widgets);
        setStatus("Ready");
      })
      .catch(() => setStatus("Cannot load dashboard layout"));
  }, []);

  const widgetIds = useMemo(() => widgets.map((widget) => widget.id), [widgets]);

  function addWidget(template: WidgetTemplate) {
    setWidgets((current) => [
      ...current,
      {
        id: makeId("widget"),
        title: `Widget ${current.length + 1}`,
        size: template.size,
        items: [],
      },
    ]);
  }

  function removeItem(widgetId: string, itemId: string) {
    setWidgets((current) =>
      current.map((widget) =>
        widget.id === widgetId ? { ...widget, items: widget.items.filter((item) => item.id !== itemId) } : widget,
      ),
    );
  }

  function putItemInSlot(widgetId: string, slot: number, item: DashboardWidgetItem) {
    setWidgets((current) =>
      current.map((widget) => {
        if (widget.id !== widgetId) return widget;
        return {
          ...widget,
          items: [...widget.items.filter((existing) => existing.id !== item.id && existing.slot !== slot), { ...item, slot }],
        };
      }),
    );
  }

  function moveItemBetweenWidgets(sourceWidgetId: string, targetWidgetId: string, slot: number, item: DashboardWidgetItem) {
    setWidgets((current) =>
      current.map((widget) => {
        if (widget.id === sourceWidgetId && sourceWidgetId !== targetWidgetId) {
          return { ...widget, items: widget.items.filter((existing) => existing.id !== item.id) };
        }
        if (widget.id === targetWidgetId) {
          return {
            ...widget,
            items: [...widget.items.filter((existing) => existing.id !== item.id && existing.slot !== slot), { ...item, slot }],
          };
        }
        return widget;
      }),
    );
  }

  function handleDragStart(event: DragStartEvent) {
    const data = event.active.data.current as ActiveDrag | undefined;
    setActiveDrag(data ?? null);
  }

  function handleDragEnd(event: DragEndEvent) {
    const data = event.active.data.current as ActiveDrag | undefined;
    const over = event.over;
    setActiveDrag(null);

    if (!data || !over) return;

    if (data.kind === "widget") {
      const oldIndex = widgets.findIndex((widget) => widget.id === event.active.id);
      const newIndex = widgets.findIndex((widget) => widget.id === over.id);
      if (oldIndex >= 0 && newIndex >= 0 && oldIndex !== newIndex) {
        setWidgets((current) => arrayMove(current, oldIndex, newIndex));
      }
      return;
    }

    const overData = over.data.current as { widgetId?: string; slot?: number } | undefined;
    if (!overData?.widgetId || typeof overData.slot !== "number") return;

    if (data.kind === "palette-item") {
      putItemInSlot(overData.widgetId, overData.slot, {
        id: makeId("item"),
        type: data.template.type,
        slot: overData.slot,
        label: data.template.label,
      });
      return;
    }

    if (data.kind === "widget-item") {
      moveItemBetweenWidgets(data.widgetId, overData.widgetId, overData.slot, data.item);
    }
  }

  async function handleSave() {
    setStatus("Saving layout...");
    try {
      const saved = await saveDashboardWidgets({ widgets });
      setWidgets(saved.widgets);
      setStatus("Layout saved");
    } catch {
      setStatus("Save failed");
    }
  }

  return (
    <main className="min-h-screen bg-screen-bg p-6">
      <header className="mb-5 flex items-start justify-between gap-4">
        <div>
          <h1 className="text-3xl font-black text-cj-navy">Dashboard Customization</h1>
          <p className="font-bold text-slate-500">Drag widgets to reorder. Drag items into widget slots and save layout.</p>
        </div>
        <div className="flex gap-2">
          <button className="rounded-md bg-white px-4 py-2 font-black shadow-panel" type="button" onClick={() => onNavigate("/admin")}>
            Back
          </button>
          <button className="inline-flex items-center gap-2 rounded-md bg-cj-navy px-4 py-2 font-black text-white shadow-panel" type="button" onClick={handleSave}>
            <Save className="h-5 w-5" />
            Save Layout
          </button>
        </div>
      </header>

      <DndContext collisionDetection={closestCenter} sensors={sensors} onDragEnd={handleDragEnd} onDragStart={handleDragStart}>
        <div className="grid grid-cols-[280px_minmax(0,1fr)] gap-5">
          <aside className="space-y-5">
            <section className="rounded-lg bg-white p-4 shadow-panel">
              <h2 className="text-lg font-black text-cj-navy">Add Widget</h2>
              <div className="mt-3 grid gap-2">
                {widgetTemplates.map((template) => (
                  <button key={template.size} className="rounded-md border border-slate-200 p-3 text-left hover:border-cj-blue" type="button" onClick={() => addWidget(template)}>
                    <strong className="block text-sm font-black">{template.label}</strong>
                    <span className="text-xs font-bold text-slate-500">{template.description}</span>
                  </button>
                ))}
              </div>
            </section>

            <section className="rounded-lg bg-white p-4 shadow-panel">
              <h2 className="text-lg font-black text-cj-navy">Widget Items</h2>
              <div className="mt-3 grid gap-2">
                {itemTemplates.map((template) => (
                  <PaletteItem key={template.type} template={template} />
                ))}
              </div>
            </section>

            <p className="rounded-lg bg-white p-3 text-sm font-black text-slate-600 shadow-panel">{status}</p>
          </aside>

          <SortableContext items={widgetIds} strategy={rectSortingStrategy}>
            <section className="grid auto-rows-[minmax(220px,auto)] grid-cols-1 gap-5 lg:grid-cols-2">
              {widgets.map((widget) => (
                <SortableWidgetCard key={widget.id} widget={widget} onRemoveItem={removeItem} />
              ))}
            </section>
          </SortableContext>
        </div>

        <DragOverlay>
          {activeDrag?.kind === "palette-item" ? (
            <div className="rounded-md bg-cj-navy px-4 py-3 text-sm font-black text-white shadow-panel">{activeDrag.template.label}</div>
          ) : null}
          {activeDrag?.kind === "widget-item" ? (
            <div className="rounded-md bg-cj-blue px-4 py-3 text-sm font-black text-white shadow-panel">{activeDrag.item.label}</div>
          ) : null}
          {activeDrag?.kind === "widget" ? (
            <div className="rounded-lg bg-white p-4 text-lg font-black text-cj-navy shadow-panel">{activeDrag.widget.title}</div>
          ) : null}
        </DragOverlay>
      </DndContext>
    </main>
  );
}
