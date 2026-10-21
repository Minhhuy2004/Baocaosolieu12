import React from "react"
import ReactDOM from "react-dom/client"
import App from "./App"
import "./index.css"

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
