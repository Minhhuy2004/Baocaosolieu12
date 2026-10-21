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
          style={{
            maxWidth: 560,
            margin: "12vh auto",
            padding: 32,
            color: "#1f2937",
            fontFamily: "Arial, sans-serif",
            textAlign: "center",
          }}
        >
          <h1>Ứng dụng gặp lỗi khi hiển thị</h1>
          <p>Hãy tải lại trang để sử dụng phiên bản mới nhất.</p>
          <a
            href={`${window.location.pathname}?refresh=${Date.now()}`}
            style={{
              display: "inline-block",
              padding: "12px 20px",
              borderRadius: 8,
              background: "#006b54",
              color: "#fff",
              fontWeight: 700,
              textDecoration: "none",
            }}
          >
            Tải lại ứng dụng
          </a>
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
