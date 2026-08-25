import { lazy, Suspense, useEffect, useState } from "react";

import { TvDashboardPage } from "./pages/TvDashboardPage";
import { fetchDashboard } from "./services/api";
import type { DashboardResponse } from "./types/dashboard";

const AdminDashboardPage = lazy(() => import("./pages/AdminDashboardPage").then((module) => ({ default: module.AdminDashboardPage })));
const DashboardCustomizePage = lazy(() => import("./pages/DashboardCustomizePage").then((module) => ({ default: module.DashboardCustomizePage })));
const LoginPage = lazy(() => import("./pages/LoginPage").then((module) => ({ default: module.LoginPage })));
const SettingsPage = lazy(() => import("./pages/SettingsPage").then((module) => ({ default: module.SettingsPage })));
const UploadPage = lazy(() => import("./pages/UploadPage").then((module) => ({ default: module.UploadPage })));

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
    if (path === "/tv" || path === "/") return;
    fetchDashboard()
      .then(setDashboard)
      .catch(() => setDashboard(null));
  }, [path]);

  function navigate(nextPath: string) {
    window.history.pushState({}, "", nextPath);
    setPath(nextPath);
  }

  if (path === "/tv" || path === "/") {
    return <TvDashboardPage />;
  }

  return (
    <Suspense fallback={<main className="min-h-screen bg-screen-bg" />}>
      {!isAuthenticated ? <LoginPage onLogin={() => setIsAuthenticated(true)} />
        : path === "/upload" ? <UploadPage onNavigate={navigate} />
          : path === "/settings" ? <SettingsPage settings={dashboard?.settings || defaultSettings} onNavigate={navigate} />
            : path === "/customize" ? <DashboardCustomizePage onNavigate={navigate} />
              : <AdminDashboardPage dashboard={dashboard} onNavigate={navigate} />}
    </Suspense>
  );
}

export default App;
