import axios from "axios";

import type { DashboardResponse, DashboardSettings, DashboardWidgetLayout, WorkspaceLayout } from "../types/dashboard";

export const api = axios.create({
  baseURL: "/api",
  timeout: 30000,
});

export type ExcelUploadRecord = {
  id: string;
  category: "inbound" | "pick" | "outbound";
  filename: string;
  content_type: string;
  file_size: number;
  uploaded_at: string | null;
  managed?: boolean;
};

export type FileServerSyncStatus = {
  enabled: boolean;
  path: string;
  interval_seconds: number;
  state: "disabled" | "waiting" | "checking" | "idle" | "error";
  last_checked_at: string | null;
  last_synced_at: string | null;
  latest_filename: string | null;
  upload_id: string | null;
  message: string;
};

export async function fetchDashboard() {
  const response = await api.get<DashboardResponse>("/data", {
    params: { _ts: Date.now() },
    headers: {
      "Cache-Control": "no-cache",
      Pragma: "no-cache",
    },
  });
  return response.data;
}

export async function uploadExcel(file: File, category: ExcelUploadRecord["category"] = "inbound") {
  const formData = new FormData();
  formData.append("file", file);
  formData.append("category", category);
  const response = await api.post("/upload/excel", formData, {
    headers: { "Content-Type": "multipart/form-data" },
  });
  return response.data;
}

export async function replaceExcelUpload(uploadId: string, file: File) {
  const formData = new FormData();
  formData.append("file", file);
  const response = await api.put(`/uploads/excel/${uploadId}`, formData, {
    headers: { "Content-Type": "multipart/form-data" },
  });
  return response.data;
}

export async function fetchExcelUploads() {
  const response = await api.get<{ status: string; data: ExcelUploadRecord[] }>("/uploads/excel");
  return response.data.data;
}

export async function fetchFileServerSyncStatus() {
  const response = await api.get<{ status: string; data: FileServerSyncStatus }>("/file-server/status");
  return response.data.data;
}

export async function syncFileServerNow() {
  const response = await api.post<{ status: string; data: { changed: boolean; upload_id: string; filename: string } }>("/file-server/sync");
  return response.data.data;
}

export async function deleteExcelUpload(uploadId: string) {
  const response = await api.delete<{ status: string; message: string }>(`/uploads/excel/${uploadId}`);
  return response.data;
}

export async function fetchExcelSheets(uploadId: string) {
  const response = await api.get<{ status: string; filename: string; data: string[] }>(`/uploads/excel/${uploadId}/sheets`);
  return response.data;
}

export async function fetchExcelRange(uploadId: string, sheet: string, cellRange: string) {
  const response = await api.get<{
    status: string;
    filename: string;
    sheet: string;
    cell_range: string;
    rows: number;
    columns: number;
    data: unknown[][];
    number_formats?: string[][];
  }>(`/uploads/excel/${uploadId}/range`, { params: { sheet, cell_range: cellRange } });
  return response.data;
}

export async function fetchExcelCells(uploadId: string, sheet: string, cells: string[]) {
  const response = await api.post<{
    status: string;
    filename: string;
    sheet: string;
    data: Record<string, { value: unknown; number_format: string }>;
  }>(`/uploads/excel/${uploadId}/cells`, { sheet, cells });
  return response.data;
}

export async function saveDashboardSettings(settings: DashboardSettings) {
  const response = await api.put("/settings/dashboard", settings);
  return response.data;
}

export async function fetchDashboardWidgets() {
  const response = await api.get<{ status: string; data: DashboardWidgetLayout }>("/dashboard/widgets");
  return response.data.data;
}

export async function saveDashboardWidgets(layout: DashboardWidgetLayout) {
  const response = await api.put<{ status: string; data: DashboardWidgetLayout }>("/dashboard/widgets", layout);
  return response.data.data;
}

export async function fetchWorkspaceLayout() {
  const response = await api.get<{ status: string; data: WorkspaceLayout }>("/workspace/layout");
  return response.data.data;
}

export async function saveWorkspaceLayout(layout: WorkspaceLayout) {
  const response = await api.put<{ status: string; data: WorkspaceLayout }>("/workspace/layout", layout);
  return response.data.data;
}

export async function switchWorkspaceUploadSheet(uploadId: string, sheet: string) {
  const response = await api.put<{
    status: string;
    sheet: string;
    updated_widgets: number;
    data: WorkspaceLayout;
  }>("/workspace/excel-sheet", { upload_id: uploadId, sheet });
  return response.data;
}

export async function login(username: string, password: string) {
  const response = await api.post("/auth/login", { username, password });
  return response.data;
}

export async function verifyEditPin(pin: string) {
  const response = await api.post<{ status: string }>("/auth/edit-pin", { pin });
  return response.data;
}
