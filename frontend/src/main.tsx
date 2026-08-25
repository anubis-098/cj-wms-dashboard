import React, { ErrorInfo, ReactNode } from "react";
import ReactDOM from "react-dom/client";

import App from "./App";
import "./index.css";

const isTvRuntime = /SMART-TV|Tizen|SamsungBrowser/i.test(window.navigator.userAgent);
document.documentElement.classList.toggle("tv-runtime", isTvRuntime);
window.clearTimeout((window as typeof window & { __cjWmsBootTimer?: number }).__cjWmsBootTimer);

class RuntimeErrorBoundary extends React.Component<{ children: ReactNode }, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("CJ WMS frontend runtime error", error, info);
  }

  render() {
    if (this.state.failed) {
      return (
        <main className="grid min-h-screen place-items-center bg-white p-8 text-center text-slate-700">
          <div>
            <strong className="block text-xl">CJ WMS cannot start</strong>
            <span className="mt-2 block text-sm">Refresh the TV browser or contact IT Auttha.</span>
          </div>
        </main>
      );
    }
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <RuntimeErrorBoundary>
    <React.StrictMode>
      <App />
    </React.StrictMode>
  </RuntimeErrorBoundary>,
);
