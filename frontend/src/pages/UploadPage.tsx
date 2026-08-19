import { ChangeEvent, useState } from "react";
import { FileSpreadsheet, Upload } from "lucide-react";

import { uploadExcel } from "../services/api";

type UploadPageProps = {
  onNavigate: (path: string) => void;
};

export function UploadPage({ onNavigate }: UploadPageProps) {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState("Ready");
  const [uploading, setUploading] = useState(false);

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    setFile(event.target.files?.[0] || null);
  }

  async function handleUpload() {
    if (!file) {
      setStatus("Please select .xlsx file");
      return;
    }
    setUploading(true);
    setStatus("Uploading...");
    try {
      await uploadExcel(file);
      setStatus("Upload completed");
    } catch {
      setStatus("Upload failed");
    } finally {
      setUploading(false);
    }
  }

  return (
    <main className="min-h-screen bg-screen-bg p-6">
      <button className="mb-5 rounded-md bg-white px-4 py-2 font-black shadow-panel" onClick={() => onNavigate("/admin")}>
        Back
      </button>
      <section className="max-w-2xl rounded-lg bg-white p-6 shadow-panel">
        <FileSpreadsheet className="mb-4 h-9 w-9 text-emerald-600" />
        <h1 className="text-3xl font-black text-cj-navy">Upload Excel</h1>
        <p className="mt-2 font-bold text-slate-500">รองรับไฟล์ `.xlsx` สำหรับข้อมูล Inbound, Pick และ Outbound</p>
        <div className="mt-6 rounded-lg border-2 border-dashed border-slate-300 p-6">
          <input type="file" accept=".xlsx" onChange={handleFileChange} />
          <p className="mt-3 text-sm font-bold text-slate-500">{file?.name || "No file selected"}</p>
        </div>
        <button className="mt-5 inline-flex items-center gap-2 rounded-md bg-cj-blue px-4 py-3 font-black text-white disabled:opacity-50" disabled={uploading} onClick={handleUpload}>
          <Upload className="h-5 w-5" />
          {uploading ? "Uploading" : "Upload"}
        </button>
        <p className="mt-4 font-black text-slate-600">{status}</p>
      </section>
    </main>
  );
}
