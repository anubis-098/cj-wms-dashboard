import { FileSpreadsheet, LoaderCircle, Moon, Pencil, Plus, RefreshCw, Save, Settings, Sun, Trash2, X } from "lucide-react";
import { ChangeEvent, useEffect, useRef, useState } from "react";

import { BoxWorkspaceEditor } from "../components/workspace/BoxWorkspaceEditor";
import { deleteExcelUpload, fetchExcelSheets, fetchExcelUploads, fetchFileServerSyncStatus, fetchWorkspaceLayout, replaceExcelUpload, switchWorkspaceUploadSheet, syncFileServerNow, uploadExcel } from "../services/api";
import type { ExcelUploadRecord, FileServerSyncStatus } from "../services/api";

function formatDateTime(date: Date) {
  const hours = date.getHours().toString().padStart(2, "0");
  const minutes = date.getMinutes().toString().padStart(2, "0");

  return {
    day: date.getDate().toString().padStart(2, "0"),
    weekday: date.toLocaleDateString("en-GB", { weekday: "long" }),
    month: date.toLocaleDateString("en-GB", { month: "long" }),
    year: date.getFullYear(),
    hours,
    minutes,
  };
}

export function TvDashboardPage() {
  const [now, setNow] = useState(new Date());
  const [isEditMode, setIsEditMode] = useState(false);
  const [isCancellingEdit, setIsCancellingEdit] = useState(false);
  const [isSavingEdit, setIsSavingEdit] = useState(false);
  const [cancelRequestId, setCancelRequestId] = useState(0);
  const [saveRequestId, setSaveRequestId] = useState(0);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [theme, setTheme] = useState<"light" | "dark">(() => window.localStorage.getItem("cj-wms-theme") === "dark" ? "dark" : "light");
  const [excelUploads, setExcelUploads] = useState<ExcelUploadRecord[]>([]);
  const [uploadStatus, setUploadStatus] = useState("Ready");
  const [isUploading, setIsUploading] = useState(false);
  const [deletingUploadId, setDeletingUploadId] = useState<string | null>(null);
  const [switchingSheetUploadId, setSwitchingSheetUploadId] = useState<string | null>(null);
  const [fileServerStatus, setFileServerStatus] = useState<FileServerSyncStatus | null>(null);
  const [isSyncingFileServer, setIsSyncingFileServer] = useState(false);
  const [uploadSheets, setUploadSheets] = useState<Record<string, string[]>>({});
  const [activeUploadSheets, setActiveUploadSheets] = useState<Record<string, string>>({});
  const [uploadTarget, setUploadTarget] = useState<{ category: ExcelUploadRecord["category"]; replaceId?: string } | null>(null);
  const settingsRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { day, weekday, month, year, hours, minutes } = formatDateTime(now);

  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    function updateDisplayScale() {
      const scale = Math.max(0.5, Math.min(2, window.innerWidth / 1920, window.innerHeight / 1080));
      document.documentElement.style.setProperty("--wms-display-scale", scale.toFixed(4));
    }

    updateDisplayScale();
    window.addEventListener("resize", updateDisplayScale);
    window.visualViewport?.addEventListener("resize", updateDisplayScale);
    return () => {
      window.removeEventListener("resize", updateDisplayScale);
      window.visualViewport?.removeEventListener("resize", updateDisplayScale);
      document.documentElement.style.removeProperty("--wms-display-scale");
    };
  }, []);

  useEffect(() => {
    document.documentElement.classList.toggle("dark", theme === "dark");
    window.localStorage.setItem("cj-wms-theme", theme);
    return () => document.documentElement.classList.remove("dark");
  }, [theme]);

  useEffect(() => {
    if (!isSettingsOpen) return;
    let active = true;
    Promise.all([fetchExcelUploads(), fetchWorkspaceLayout(), fetchFileServerSyncStatus().catch(() => null)])
      .then(async ([uploads, layout, serverStatus]) => {
        if (!active) return;
        setExcelUploads(uploads);
        setFileServerStatus(serverStatus);
        const sheetEntries = await Promise.all(uploads.map(async (upload) => {
          try {
            return [upload.id, (await fetchExcelSheets(upload.id)).data] as const;
          } catch {
            return [upload.id, []] as const;
          }
        }));
        if (!active) return;
        const availableSheets = Object.fromEntries(sheetEntries) as Record<string, string[]>;
        setUploadSheets(availableSheets);

        const currentSheets: Record<string, string> = {};
        const pages = layout.pages?.length ? layout.pages : [{ boxes: layout.boxes }];
        pages.forEach((page) => page.boxes.forEach((box) => box.widgets.forEach((widget) => {
          if (widget.sourceUploadId && widget.sheetName && !currentSheets[widget.sourceUploadId] && availableSheets[widget.sourceUploadId]?.includes(widget.sheetName)) {
            currentSheets[widget.sourceUploadId] = widget.sheetName;
          }
        })));
        setActiveUploadSheets(currentSheets);
      })
      .catch(() => setUploadStatus("Database unavailable"));

    function handleOutsidePointer(event: globalThis.PointerEvent) {
      if (!settingsRef.current?.contains(event.target as Node)) setIsSettingsOpen(false);
    }

    document.addEventListener("pointerdown", handleOutsidePointer);
    return () => {
      active = false;
      document.removeEventListener("pointerdown", handleOutsidePointer);
    };
  }, [isSettingsOpen]);

  async function handleExcelUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setUploadStatus("Only .xlsx files are supported");
      return;
    }

    setIsUploading(true);
    setUploadStatus(`Uploading ${file.name}...`);
    try {
      const replaceId = uploadTarget?.replaceId;
      if (replaceId) {
        await replaceExcelUpload(replaceId, file);
      } else {
        await uploadExcel(file, uploadTarget?.category ?? "inbound");
      }
      setExcelUploads(await fetchExcelUploads());
      if (replaceId) {
        const sheets = (await fetchExcelSheets(replaceId)).data;
        setUploadSheets((current) => ({ ...current, [replaceId]: sheets }));
        setActiveUploadSheets((current) => ({
          ...current,
          [replaceId]: sheets.includes(current[replaceId]) ? current[replaceId] : "",
        }));
        window.dispatchEvent(new CustomEvent("excel-upload-replaced", { detail: { uploadId: replaceId } }));
      }
      setUploadStatus(replaceId ? "File replaced" : "Upload completed");
    } catch {
      setUploadStatus("Upload failed");
    } finally {
      setIsUploading(false);
      setUploadTarget(null);
    }
  }

  function openExcelPicker(category: ExcelUploadRecord["category"], replaceId?: string) {
    setUploadTarget({ category, replaceId });
    fileInputRef.current?.click();
  }

  async function handleDeleteUpload(uploadId: string) {
    setDeletingUploadId(uploadId);
    try {
      await deleteExcelUpload(uploadId);
      setExcelUploads((current) => current.filter((upload) => upload.id !== uploadId));
      setUploadStatus("File deleted");
    } catch {
      setUploadStatus("Delete failed");
    } finally {
      setDeletingUploadId(null);
    }
  }

  async function handleSwitchUploadSheet(uploadId: string, sheet: string) {
    if (!sheet) return;
    const previousSheet = activeUploadSheets[uploadId] ?? "";
    setSwitchingSheetUploadId(uploadId);
    setActiveUploadSheets((current) => ({ ...current, [uploadId]: sheet }));
    setUploadStatus(`Switching widgets to ${sheet}...`);
    try {
      const result = await switchWorkspaceUploadSheet(uploadId, sheet);
      setActiveUploadSheets((current) => ({ ...current, [uploadId]: result.sheet }));
      setUploadStatus(`${result.updated_widgets} widget${result.updated_widgets === 1 ? "" : "s"} switched to ${result.sheet}`);
    } catch (error) {
      setActiveUploadSheets((current) => ({ ...current, [uploadId]: previousSheet }));
      const detail = (error as { response?: { data?: { detail?: string } } }).response?.data?.detail;
      setUploadStatus(detail ? `Sheet switch failed: ${detail}` : "Sheet switch failed");
    } finally {
      setSwitchingSheetUploadId(null);
    }
  }

  async function handleFileServerSync() {
    setIsSyncingFileServer(true);
    setUploadStatus("Checking SMD File Server...");
    try {
      const result = await syncFileServerNow();
      const uploads = await fetchExcelUploads();
      const sheetEntries = await Promise.all(uploads.map(async (upload) => {
        try {
          return [upload.id, (await fetchExcelSheets(upload.id)).data] as const;
        } catch {
          return [upload.id, []] as const;
        }
      }));
      setExcelUploads(uploads);
      setUploadSheets(Object.fromEntries(sheetEntries));
      setFileServerStatus(await fetchFileServerSyncStatus());
      setUploadStatus(result.changed ? `Synchronized ${result.filename}` : `${result.filename} is already up to date`);
    } catch (error) {
      const detail = (error as { response?: { data?: { detail?: string } } }).response?.data?.detail;
      setUploadStatus(detail ?? "File Server sync failed");
      setFileServerStatus(await fetchFileServerSyncStatus().catch(() => fileServerStatus));
    } finally {
      setIsSyncingFileServer(false);
    }
  }

  function formatFileSize(bytes: number) {
    return bytes < 1024 * 1024 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  return (
    <main className={`tv-dashboard h-screen bg-screen-bg ${isEditMode ? "overflow-y-auto" : "overflow-hidden"} ${theme === "dark" ? "theme-dark" : "theme-light"}`} aria-label="Workspace">
      <header className="cj-brand-header sticky top-0 z-50 flex h-20 shrink-0 items-center border-b border-slate-200 bg-white px-8 shadow-sm">
        <div className="text-sm font-black tracking-wide text-cj-navy">CJ WMS</div>
        <div className="absolute left-1/2 top-1/2 flex -translate-x-1/2 -translate-y-1/2 items-center gap-4">
          <div
            className="flex h-14 items-center font-black tabular-nums leading-none text-cj-navy"
            style={{
              fontSize: "calc(50px * var(--wms-display-scale, 1))",
              textShadow: "0 2px 10px rgba(18, 112, 219, 0.1)",
            }}
          >
            {hours}<span className="animate-clock-colon mx-1">:</span>{minutes}
          </div>
          <div className="flex h-12 flex-col justify-center border-l-2 border-slate-200 pl-4">
            <span className="text-sm font-black leading-tight tracking-wide text-slate-500">{day} {weekday}</span>
            <span className="mt-0.5 text-xs font-bold leading-tight tracking-widest text-slate-400">{month} {year}</span>
          </div>
        </div>
        <div className="ml-auto flex items-center gap-2">
          <div ref={settingsRef} className="relative">
            <button
              aria-label="Settings"
              className={`grid h-9 w-9 place-items-center rounded-md border bg-white transition ${isSettingsOpen ? "border-cj-blue text-cj-blue" : "border-slate-200 text-slate-600 hover:border-cj-blue hover:text-cj-blue"}`}
              title="Settings"
              type="button"
              onClick={() => setIsSettingsOpen((current) => !current)}
            >
              <Settings className="h-4 w-4" />
            </button>

            {isSettingsOpen ? (
              <div className="absolute right-0 top-12 z-[70] max-h-[calc(100vh-6rem)] w-96 overflow-y-auto rounded-md border border-slate-200 bg-white shadow-panel">
                <div className="flex items-center justify-between gap-3 border-b border-slate-100 px-4 py-3">
                  <div className="min-w-0">
                    <h2 className="text-sm font-black text-cj-navy">Excel files</h2>
                    <p className="text-[10px] font-bold text-slate-400">{excelUploads.length} file{excelUploads.length === 1 ? "" : "s"}</p>
                  </div>
                  <button
                    aria-label="Upload Excel"
                    className="grid h-8 w-8 shrink-0 place-items-center rounded-md bg-cj-navy text-white transition hover:bg-cj-blue disabled:cursor-not-allowed disabled:opacity-50"
                    disabled={isUploading}
                    title="Upload Excel"
                    type="button"
                    onClick={() => openExcelPicker("inbound")}
                  >
                    {isUploading && !uploadTarget?.replaceId ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <Plus className="h-4 w-4" />}
                  </button>
                  <input
                    ref={fileInputRef}
                    className="hidden"
                    type="file"
                    accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    onChange={handleExcelUpload}
                  />
                </div>

                <div className="max-h-[calc(100vh-24rem)] overflow-y-auto p-2">
                  <section className="overflow-hidden rounded-md border border-slate-200">
                    <div className="p-1">
                          {excelUploads.length > 0 ? excelUploads.map((upload) => (
                            <div key={upload.id} className="flex items-center gap-2 rounded-md px-2 py-2 hover:bg-slate-50">
                              <span className="grid h-9 w-9 shrink-0 place-items-center rounded bg-emerald-50 text-emerald-600">
                                <FileSpreadsheet className="h-4 w-4" />
                              </span>
                              <div className="min-w-0 flex-1">
                                <div className="flex min-w-0 items-center gap-1.5">
                                  <p className="truncate text-sm font-black text-cj-navy" title={upload.filename}>{upload.filename}</p>
                                  {upload.managed ? <span className="shrink-0 rounded bg-blue-50 px-1.5 py-0.5 text-[8px] font-black uppercase text-cj-blue">Auto Sync</span> : null}
                                </div>
                                <p className="text-[10px] font-bold text-slate-400">
                                  {formatFileSize(upload.file_size)}{upload.uploaded_at ? ` | ${new Date(upload.uploaded_at).toLocaleString("en-GB")}` : ""}
                                </p>
                                <select
                                  aria-label={`Sheet used by widgets for ${upload.filename}`}
                                  className="mt-1 h-8 w-full rounded border border-slate-200 bg-white px-2 text-[11px] font-bold text-cj-navy outline-none focus:border-cj-blue disabled:bg-slate-50 disabled:text-slate-400"
                                  disabled={switchingSheetUploadId === upload.id || (uploadSheets[upload.id]?.length ?? 0) === 0}
                                  title="Switch every widget using this file to another Sheet"
                                  value={activeUploadSheets[upload.id] ?? ""}
                                  onChange={(event) => handleSwitchUploadSheet(upload.id, event.target.value)}
                                >
                                  <option value="">Select Sheet for all widgets</option>
                                  {(uploadSheets[upload.id] ?? []).map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                                </select>
                              </div>
                              {switchingSheetUploadId === upload.id ? <LoaderCircle className="h-4 w-4 shrink-0 animate-spin text-cj-blue" /> : null}
                              {!upload.managed ? <button
                                aria-label={`Replace ${upload.filename}`}
                                className="grid h-8 w-8 shrink-0 place-items-center rounded-md text-cj-blue transition hover:bg-blue-50 disabled:opacity-40"
                                disabled={isUploading}
                                title="Replace file"
                                type="button"
                                onClick={() => openExcelPicker(upload.category, upload.id)}
                              >
                                {isUploading && uploadTarget?.replaceId === upload.id ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <RefreshCw className="h-4 w-4" />}
                              </button> : null}
                              {!upload.managed ? <button
                                aria-label={`Delete ${upload.filename}`}
                                className="grid h-8 w-8 shrink-0 place-items-center rounded-md text-red-500 transition hover:bg-red-50 hover:text-red-700 disabled:opacity-40"
                                disabled={deletingUploadId === upload.id || isUploading}
                                title="Delete file"
                                type="button"
                                onClick={() => handleDeleteUpload(upload.id)}
                              >
                                {deletingUploadId === upload.id ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <Trash2 className="h-4 w-4" />}
                              </button> : null}
                            </div>
                          )) : <div className="px-3 py-4 text-center text-[10px] font-bold text-slate-400">No files</div>}
                    </div>
                  </section>
                </div>

                <section className="border-t border-slate-100 px-4 py-3">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <h2 className="text-sm font-black text-cj-navy">SMD File Server</h2>
                      <p className="mt-0.5 truncate text-[10px] font-bold text-slate-400" title={fileServerStatus?.path}>{fileServerStatus?.path ?? "Status unavailable"}</p>
                    </div>
                    <button
                      aria-label="Sync latest Excel from SMD File Server"
                      className="inline-flex h-8 shrink-0 items-center gap-1.5 rounded-md border border-cj-blue px-2 text-[10px] font-black text-cj-blue transition hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-40"
                      disabled={!fileServerStatus?.enabled || isSyncingFileServer}
                      title="Sync latest Excel now"
                      type="button"
                      onClick={handleFileServerSync}
                    >
                      <RefreshCw className={`h-4 w-4 ${isSyncingFileServer ? "animate-spin" : ""}`} />
                      Sync latest Excel now
                    </button>
                  </div>
                  <div className="mt-2 grid grid-cols-[72px_minmax(0,1fr)] gap-x-2 gap-y-1 rounded border border-slate-200 bg-slate-50 p-2 text-[10px] font-bold">
                    <span className="text-slate-400">Status</span>
                    <span className={fileServerStatus?.state === "error" ? "text-red-600" : fileServerStatus?.enabled ? "text-emerald-600" : "text-slate-500"}>{fileServerStatus?.message ?? "Unavailable"}</span>
                    <span className="text-slate-400">Latest file</span>
                    <span className="truncate text-cj-navy" title={fileServerStatus?.latest_filename ?? undefined}>{fileServerStatus?.latest_filename ?? "-"}</span>
                    <span className="text-slate-400">Last sync</span>
                    <span className="text-cj-navy">{fileServerStatus?.last_synced_at ?? "-"}</span>
                    <span className="text-slate-400">Interval</span>
                    <span className="text-cj-navy">{Math.round((fileServerStatus?.interval_seconds ?? 1800) / 60)} minutes</span>
                  </div>
                </section>

                <section className="border-t border-slate-100 px-4 py-3">
                  <div className="mb-2">
                    <h2 className="text-sm font-black text-cj-navy">Theme</h2>
                    <p className="text-[10px] font-bold text-slate-400">Choose the dashboard appearance for this display.</p>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      aria-pressed={theme === "light"}
                      className={`flex h-12 items-center gap-3 rounded-md border px-3 text-left transition ${theme === "light" ? "border-cj-blue bg-blue-50 text-cj-blue" : "border-slate-200 bg-white text-slate-500 hover:border-slate-400"}`}
                      type="button"
                      onClick={() => setTheme("light")}
                    >
                      <Sun className="h-5 w-5 shrink-0" />
                      <span><strong className="block text-xs font-black">Light</strong><small className="block text-[9px] font-bold opacity-70">Bright workspace</small></span>
                    </button>
                    <button
                      aria-pressed={theme === "dark"}
                      className={`flex h-12 items-center gap-3 rounded-md border px-3 text-left transition ${theme === "dark" ? "border-cj-blue bg-slate-800 text-white" : "border-slate-200 bg-white text-slate-500 hover:border-slate-400"}`}
                      type="button"
                      onClick={() => setTheme("dark")}
                    >
                      <Moon className="h-5 w-5 shrink-0" />
                      <span><strong className="block text-xs font-black">Dark</strong><small className="block text-[9px] font-bold opacity-70">Low-light display</small></span>
                    </button>
                  </div>
                </section>

                <div className="border-t border-slate-100 px-4 py-2 text-[10px] font-bold text-slate-500">{uploadStatus}</div>
              </div>
            ) : null}
          </div>
          {isEditMode ? (
            <div className="relative">
              <button
                aria-label="Save layout and exit"
                className="grid h-9 w-9 place-items-center rounded-md border border-cj-blue bg-cj-blue text-white transition hover:border-cj-navy hover:bg-cj-navy disabled:cursor-wait disabled:opacity-60"
                disabled={isSavingEdit}
                title="Save and exit"
                type="button"
                onClick={() => {
                  setIsSavingEdit(true);
                  setSaveRequestId((current) => current + 1);
                }}
              >
                {isSavingEdit ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <Save className="h-4 w-4" />}
              </button>
              <button
                aria-label="Discard changes"
                className="absolute -right-2 -top-2 grid h-5 w-5 place-items-center rounded-full border border-slate-200 bg-white text-slate-500 shadow-sm transition hover:border-red-300 hover:text-red-600 disabled:cursor-wait disabled:opacity-50"
                disabled={isSavingEdit || isCancellingEdit}
                title="Discard changes"
                type="button"
                onClick={() => {
                  setIsCancellingEdit(true);
                  setCancelRequestId((current) => current + 1);
                }}
              >
                {isCancellingEdit ? <LoaderCircle className="h-3 w-3 animate-spin" /> : <X className="h-3 w-3" />}
              </button>
            </div>
          ) : (
            <button
              aria-label="Enter edit mode"
              className="grid h-9 w-9 place-items-center rounded-md border border-slate-200 bg-white text-slate-600 transition hover:border-cj-blue hover:text-cj-blue"
              title="Edit"
              type="button"
              onClick={() => setIsEditMode(true)}
            >
              <Pencil className="h-4 w-4" />
            </button>
          )}
        </div>
      </header>

      <BoxWorkspaceEditor
        cancelRequestId={cancelRequestId}
        editable={isEditMode}
        saveRequestId={saveRequestId}
        onCancelComplete={() => {
          setIsCancellingEdit(false);
          setIsEditMode(false);
        }}
        onSaveComplete={(saved) => {
          setIsSavingEdit(false);
          if (saved) setIsEditMode(false);
        }}
      />
    </main>
  );
}
