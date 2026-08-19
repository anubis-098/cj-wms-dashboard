import { FormEvent, useState } from "react";
import { LockKeyhole } from "lucide-react";

import { login } from "../services/api";

type LoginPageProps = {
  onLogin: () => void;
};

export function LoginPage({ onLogin }: LoginPageProps) {
  const [username, setUsername] = useState("admin");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function handleSubmit(event: FormEvent) {
    event.preventDefault();
    setError("");
    try {
      await login(username, password || "dev");
      window.localStorage.setItem("cj-wms-auth", "dev-token");
      onLogin();
    } catch {
      setError("Login failed");
    }
  }

  return (
    <main className="flex min-h-screen items-center justify-center bg-screen-bg p-6">
      <form className="w-full max-w-md rounded-lg bg-white p-6 shadow-panel" onSubmit={handleSubmit}>
        <LockKeyhole className="mb-4 h-9 w-9 text-cj-blue" />
        <h1 className="text-3xl font-black text-cj-navy">CJ WMS Login</h1>
        <div className="mt-6 space-y-4">
          <label className="block">
            <span className="text-sm font-black uppercase text-slate-500">Username</span>
            <input className="mt-1 w-full rounded-md border border-slate-200 px-3 py-2" value={username} onChange={(event) => setUsername(event.target.value)} />
          </label>
          <label className="block">
            <span className="text-sm font-black uppercase text-slate-500">Password</span>
            <input className="mt-1 w-full rounded-md border border-slate-200 px-3 py-2" type="password" value={password} onChange={(event) => setPassword(event.target.value)} />
          </label>
        </div>
        {error && <p className="mt-4 rounded-md bg-rose-50 px-3 py-2 text-sm font-bold text-rose-700">{error}</p>}
        <button className="mt-6 w-full rounded-md bg-cj-navy px-4 py-3 font-black text-white">Login</button>
      </form>
    </main>
  );
}
