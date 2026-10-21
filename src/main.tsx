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
          aria-hidden="true"
          style={{
            minHeight: "100vh",
            background: "#eef2f7",
          }}
        />
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
