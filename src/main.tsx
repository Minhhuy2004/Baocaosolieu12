import React from "react"
import ReactDOM from "react-dom/client"
import App from "./App"
import "./index.css"

window.addEventListener("vite:preloadError", (event) => {
  event.preventDefault()

  const url = new URL(window.location.href)
  if (url.searchParams.has("__asset_reload")) return

  url.searchParams.set("__asset_reload", String(Date.now()))
  window.location.replace(url.toString())
})

class AppErrorBoundary extends React.Component<React.PropsWithChildren, {
  hasError: boolean
}> {
  state = { hasError: false }

  static getDerivedStateFromError() {
    return { hasError: true }
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error("Application rendering failed.", error, info.componentStack)
  }

  render() {
    if (this.state.hasError) {
      return (
        <main
          role="alert"
          style={{
            minHeight: "100vh",
            display: "grid",
            placeItems: "center",
            padding: 24,
            color: "#475569",
            background: "#eef2f7",
            textAlign: "center",
          }}
        >
          <div
            style={{
              maxWidth: 440,
              padding: "24px 28px",
              border: "1px solid #e2e8f0",
              borderRadius: 12,
              background: "#fff",
              boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
            }}
          >
            Nội dung tạm thời không khả dụng.
            <button
              type="button"
              onClick={() => window.location.reload()}
              style={{
                display: "block",
                margin: "16px auto 0",
                padding: "9px 16px",
                border: 0,
                borderRadius: 6,
                color: "#fff",
                background: "#006b68",
                font: "inherit",
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              Tải lại trang
            </button>
          </div>
        </main>
      )
    }

    return this.props.children
  }
}

const rootElement = document.getElementById("root")
if (!rootElement) {
  throw new Error("Missing root element.")
}

ReactDOM.createRoot(rootElement).render(
  <React.StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </React.StrictMode>,
)
