import { useEffect, useState } from "react";

import { AdminDashboardPage } from "./pages/AdminDashboardPage";
import { DashboardCustomizePage } from "./pages/DashboardCustomizePage";
import { LoginPage } from "./pages/LoginPage";
import { SettingsPage } from "./pages/SettingsPage";
import { TvDashboardPage } from "./pages/TvDashboardPage";
import { UploadPage } from "./pages/UploadPage";
import { fetchDashboard } from "./services/api";
import type { DashboardResponse } from "./types/dashboard";

const defaultSettings = {
  refresh_seconds: 60,
  theme: "light",
  show_inbound: true,
  show_pick: true,
  show_outbound: true,
};

function App() {
  const [path, setPath] = useState(window.location.pathname);
  const [dashboard, setDashboard] = useState<DashboardResponse | null>(null);
  const [isAuthenticated, setIsAuthenticated] = useState(() => Boolean(window.localStorage.getItem("cj-wms-auth")));

  useEffect(() => {
    const handlePopState = () => setPath(window.location.pathname);
    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, []);

  useEffect(() => {
    fetchDashboard()
      .then(setDashboard)
      .catch(() => setDashboard(null));
  }, []);

  function navigate(nextPath: string) {
    window.history.pushState({}, "", nextPath);
    setPath(nextPath);
  }

  if (path === "/tv" || path === "/") {
    return <TvDashboardPage />;
  }

  if (!isAuthenticated) {
    return <LoginPage onLogin={() => setIsAuthenticated(true)} />;
  }

  if (path === "/upload") {
    return <UploadPage onNavigate={navigate} />;
  }

  if (path === "/settings") {
    return <SettingsPage settings={dashboard?.settings || defaultSettings} onNavigate={navigate} />;
  }

  if (path === "/customize") {
    return <DashboardCustomizePage onNavigate={navigate} />;
  }

  return <AdminDashboardPage dashboard={dashboard} onNavigate={navigate} />;
}

export default App;
