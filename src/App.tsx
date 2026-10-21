import {
  Component,
  lazy,
  Suspense,
  useState,
  useMemo,
  useRef,
  useEffect,
  type ErrorInfo,
  type ReactNode,
} from "react"
import { createPortal } from "react-dom"
import type { WorkBook } from "xlsx"
import initialData from "./initialData.json"

const DashboardChart = lazy(() =>
  import("./chart-renderer").then((module) => ({
    default: module.DashboardChart,
  })),
)
const BieuDoChart = lazy(() =>
  import("./chart-renderer").then((module) => ({
    default: module.BieuDoChart,
  })),
)

// ── TYPES ─────────────────────────────────────────────────────────────────────

export type Metric = {
  dk: number
  kh: number
  giaoTang: number
  th: number
  conTh: number
  soDk: number
  soDauThang: number
  soTuanTruoc: number
  pctHT: number | string | null
}

export type Member = {
  code?: string
  name: string
  fullName: string
  unit: string
  hdv: Metric
  td: Metric
}

export type ReportRow = {
  code?: string
  name: string
  fullName: string
  unit: string
  hdv: Metric
  td: Metric
  kind: "member" | "unit" | "manager" | "total"
}

type Page = "dashboard" | "bieudo" | "solieu" | "xuatpdf"

class PageErrorBoundary extends Component<
  { children: ReactNode; pageName: Page },
  { hasError: boolean }
> {
  state = { hasError: false }

  static getDerivedStateFromError() {
    return { hasError: true }
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error(
      `Page rendering failed: ${this.props.pageName}.`,
      error,
      info.componentStack,
    )
  }

  render() {
    if (this.state.hasError) {
      return (
        <div
          aria-hidden="true"
          style={{ minHeight: 320, background: "#f1f5f9" }}
        />
      )
    }

    return this.props.children
  }
}

// ── FORMATTERS ────────────────────────────────────────────────────────────────

const fmt1 = (val: number | null | undefined): string => {
  if (val === null || val === undefined || isNaN(val)) return "0.0"
  return val.toLocaleString("en-US", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })
}

const fmt2 = (val: number | null | undefined): string => {
  if (val === null || val === undefined || isNaN(val)) return "0.00"
  return val.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })
}

const fmtPct = (val: number | string | null | undefined): string => {
  if (val === null || val === undefined || val === "") return ""
  if (typeof val === "string") {
    if (val.includes("#")) return val.trim()
    const parsed = parseFloat(val.replace(/%/g, "").replace(/,/g, ""))
    if (isNaN(parsed)) return val
    val = parsed
  }
  const pct =
    Math.abs(val) < 50 && val !== 0 ? Math.round(val * 100) : Math.round(val)
  return `${pct}%`
}

const getNumericPct = (m: Metric): number => {
  if (typeof m.pctHT === "number") {
    return Math.abs(m.pctHT) < 50 && m.pctHT !== 0 ? m.pctHT * 100 : m.pctHT
  }
  if (typeof m.pctHT === "string") {
    const parsed = parseFloat(m.pctHT.replace(/%/g, ""))
    return Number.isFinite(parsed) ? parsed : 0
  }
  return m.kh > 0 ? (m.th / m.kh) * 100 : 0
}

const toNumber = (value: unknown): number => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0
  const normalized = String(value ?? "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim()
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : 0
}

const pctColor = (p: number) =>
  p >= 100 ? "#059669" : p >= 70 ? "#d97706" : "#dc2626"
const diffColor = (n: number) => (n >= 0 ? "#059669" : "#dc2626")

const isUnitRow = (name: string) => /^(P\.|PGD\s)/i.test(name.trim())
const isManagerRow = (name: string) => /^(CHỊ\s|A\.CƯỜNG$)/i.test(name.trim())
const isTotalRow = (name: string) => /^TC$/i.test(name.trim())
const getMemberId = (member: Member, index: number) =>
  `${member.code?.trim() || member.fullName}-${index}`

// ── ROBUST EXCEL PARSER ───────────────────────────────────────────────────────

export async function parseExcelWorkbook(
  workbook: WorkBook,
): Promise<{ members: Member[]; rows: ReportRow[] }> {
  const XLSX = await import("xlsx")
  const importedMembers: Member[] = []
  const importedRows: ReportRow[] = []

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) continue

    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
    })
    if (!rows || rows.length < 2) continue

    let headerIdx = -1
    let hdvDkCol = -1
    let tdDkCol = -1

    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r]
      if (!Array.isArray(row)) continue
      const hdvIdx = row.findIndex((c) => /h[dđ]v\s*[dđ]k/i.test(String(c)))
      if (hdvIdx >= 0) {
        headerIdx = r
        hdvDkCol = hdvIdx
        tdDkCol = row.findIndex((c) => /t[dđ]\s*[dđ]k/i.test(String(c)))
        break
      }
    }

    if (headerIdx < 0 || hdvDkCol < 0) continue
    if (tdDkCol < 0) tdDkCol = hdvDkCol + 9

    const nameCol = hdvDkCol >= 2 ? 1 : 0
    const codeCol = hdvDkCol >= 2 ? 0 : -1

    let memberBuffer: Member[] = []

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r]
      if (!Array.isArray(row) || !row.length) continue

      const rawName = String(row[nameCol] ?? "").trim()
      if (!rawName || rawName.toLowerCase() === "tên") continue

      let fullName = rawName
      if (row[22] && String(row[22]).trim()) {
        fullName = String(row[22]).trim()
      } else if (
        row[row.length - 1] &&
        typeof row[row.length - 1] === "string" &&
        String(row[row.length - 1]).length > rawName.length
      ) {
        fullName = String(row[row.length - 1]).trim()
      }

      const parseMetric = (startCol: number): Metric => {
        const dk = toNumber(row[startCol])
        const kh = toNumber(row[startCol + 1])
        const giaoTang = toNumber(row[startCol + 2])
        const th = toNumber(row[startCol + 3])
        const conTh = toNumber(row[startCol + 4])
        const soDk = toNumber(row[startCol + 5])
        const soDauThang = toNumber(row[startCol + 6])
        const soTuanTruoc = toNumber(row[startCol + 7])

        const rawPct = row[startCol + 8]
        let pctHT: number | string | null = null
        if (rawPct !== "" && rawPct !== null && rawPct !== undefined) {
          if (typeof rawPct === "string" && rawPct.includes("#")) {
            pctHT = rawPct.trim()
          } else {
            pctHT = toNumber(rawPct)
          }
        }
        return {
          dk,
          kh,
          giaoTang,
          th,
          conTh,
          soDk,
          soDauThang,
          soTuanTruoc,
          pctHT,
        }
      }

      const hdv = parseMetric(hdvDkCol)
      const td = parseMetric(tdDkCol)
      const code = codeCol >= 0 ? String(row[codeCol] ?? "").trim() : ""

      if (isUnitRow(rawName)) {
        const unitName = rawName
        memberBuffer.forEach((m) => {
          m.unit = unitName
          importedMembers.push(m)
          importedRows.push({ ...m, kind: "member" })
        })
        memberBuffer = []
        importedRows.push({
          code,
          name: unitName,
          fullName: unitName,
          unit: unitName,
          hdv,
          td,
          kind: "unit",
        })
      } else if (isManagerRow(rawName)) {
        importedRows.push({
          code,
          name: rawName,
          fullName: rawName,
          unit: rawName,
          hdv,
          td,
          kind: "manager",
        })
      } else if (isTotalRow(rawName)) {
        importedRows.push({
          code,
          name: rawName,
          fullName: rawName,
          unit: rawName,
          hdv,
          td,
          kind: "total",
        })
      } else {
        memberBuffer.push({ code, name: rawName, fullName, unit: "", hdv, td })
      }
    }

    if (memberBuffer.length) {
      memberBuffer.forEach((m) => {
        if (!m.unit) m.unit = sheetName
        importedMembers.push(m)
        importedRows.push({ ...m, kind: "member" })
      })
      memberBuffer = []
    }

    if (importedMembers.length > 0) break
  }

  return { members: importedMembers, rows: importedRows }
}

// ── AGGREGATE HELPERS ─────────────────────────────────────────────────────────

const sumMetric = (
  members: Member[],
  type: "hdv" | "td",
  field: keyof Metric,
) =>
  members.reduce((acc, m) => {
    const val = m[type][field]
    return acc + (typeof val === "number" ? val : 0)
  }, 0)

const avgMetricPct = (members: Member[], type: "hdv" | "td") => {
  if (!members.length) return 0
  return (
    members.reduce((acc, m) => acc + getNumericPct(m[type]), 0) / members.length
  )
}

// ── SUB-COMPONENTS ────────────────────────────────────────────────────────────

function KPICard({
  label,
  value,
  sub,
  subColor,
  accent,
}: {
  label: string
  value: string
  sub?: string
  subColor?: string
  accent: string
}) {
  return (
    <div
      style={{
        background: "#fff",
        borderRadius: 10,
        padding: "18px 20px",
        boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
        borderTop: `3px solid ${accent}`,
        display: "flex",
        flexDirection: "column",
        gap: 5,
        minWidth: 0,
      }}
    >
      <div
        style={{
          fontSize: 11,
          fontWeight: 700,
          color: "#9ca3af",
          letterSpacing: 0.8,
          textTransform: "uppercase",
        }}
      >
        {label}
      </div>
      <div
        className="mono"
        style={{
          fontSize: 21,
          fontWeight: 700,
          color: "#111827",
          lineHeight: 1.1,
        }}
      >
        {value}
      </div>
      {sub && (
        <div
          style={{
            fontSize: 12,
            color: subColor || "#6b7280",
            fontWeight: 500,
          }}
        >
          {sub}
        </div>
      )}
    </div>
  )
}

function PctBar({ pct }: { pct: number }) {
  const color = pctColor(pct)
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div
        style={{
          flex: 1,
          height: 6,
          background: "#e5e7eb",
          borderRadius: 3,
          overflow: "hidden",
        }}
      >
        <div
          style={{
            width: `${Math.min(Math.max(pct, 0), 100)}%`,
            height: "100%",
            background: color,
            borderRadius: 3,
            transition: "width 0.3s",
          }}
        />
      </div>
      <span
        className="mono"
        style={{
          fontSize: 12,
          color,
          fontWeight: 700,
          minWidth: 54,
          textAlign: "right",
        }}
      >
        {pct.toFixed(1)}%
      </span>
    </div>
  )
}

function SectionTitle({ children }: { children: React.ReactNode }) {
  return (
    <div
      style={{
        fontSize: 13,
        fontWeight: 800,
        color: "#1f2937",
        letterSpacing: 0.5,
        marginBottom: 12,
        textTransform: "uppercase",
      }}
    >
      {children}
    </div>
  )
}

function FilterBar({
  selectedDonVi,
  setSelectedDonVi,
  selectedCBKH,
  setSelectedCBKH,
  availableCBKH,
  units,
}: {
  selectedDonVi: string
  setSelectedDonVi: (v: string) => void
  selectedCBKH: string
  setSelectedCBKH: (v: string) => void
  availableCBKH: { id: string; label: string }[]
  units: string[]
}) {
  const sel: React.CSSProperties = {
    border: "1px solid #d1d5db",
    borderRadius: 7,
    padding: "7px 30px 7px 11px",
    fontSize: 13,
    color: "#111827",
    background: "#fff",
    cursor: "pointer",
    fontFamily: "'Be Vietnam Pro', sans-serif",
    outline: "none",
    boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
  }
  const lbl: React.CSSProperties = {
    fontSize: 11,
    color: "#6b7280",
    fontWeight: 700,
    letterSpacing: 0.5,
    textTransform: "uppercase",
  }

  return (
    <div
      style={{
        display: "flex",
        gap: 14,
        alignItems: "center",
        flexWrap: "wrap",
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
        <span style={lbl}>Đơn vị</span>
        <select
          value={selectedDonVi}
          onChange={(e) => {
            setSelectedDonVi(e.target.value)
            setSelectedCBKH("Tất cả")
          }}
          style={sel}
        >
          <option value="Tất cả">Tất cả đơn vị ({units.length})</option>
          {units.map((u) => (
            <option key={u} value={u}>
              {u}
            </option>
          ))}
        </select>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
        <span style={lbl}>CBKH</span>
        <select
          value={selectedCBKH}
          onChange={(e) => setSelectedCBKH(e.target.value)}
          style={sel}
        >
          <option value="Tất cả">Tất cả CBKH ({availableCBKH.length})</option>
          {availableCBKH.map((member) => (
            <option key={member.id} value={member.id}>
              {member.label}
            </option>
          ))}
        </select>
      </div>
    </div>
  )
}

function ChartLoadingPlaceholder({ height }: { height: number }) {
  return (
    <div
      role="status"
      aria-label="Đang tải biểu đồ"
      style={{
        height,
        display: "grid",
        placeItems: "center",
        background: "#fff",
        color: "#64748b",
        fontSize: 13,
      }}
    >
      Đang tải biểu đồ...
    </div>
  )
}

// ── DASHBOARD PAGE ────────────────────────────────────────────────────────────

function DashboardPage({ members }: { members: Member[] }) {
  const [tab, setTab] = useState<"hdv" | "td">("hdv")
  const [shouldRenderCharts, setShouldRenderCharts] = useState(false)

  useEffect(() => {
    const timeoutId = window.setTimeout(() => setShouldRenderCharts(true), 250)
    return () => window.clearTimeout(timeoutId)
  }, [])

  const hdvTh = sumMetric(members, "hdv", "th")
  const hdvDk = sumMetric(members, "hdv", "dk")
  const hdvKh = sumMetric(members, "hdv", "kh")
  const hdvConTh = sumMetric(members, "hdv", "conTh")
  const hdvPct = avgMetricPct(members, "hdv")

  const tdTh = sumMetric(members, "td", "th")
  const tdDk = sumMetric(members, "td", "dk")
  const tdKh = sumMetric(members, "td", "kh")
  const tdConTh = sumMetric(members, "td", "conTh")
  const tdPct = avgMetricPct(members, "td")

  const barDataHDV = members.slice(0, 18).map((m) => ({
    name: m.name,
    fullName: m.fullName,
    ĐK: m.hdv.dk,
    KH: m.hdv.kh,
    TH: m.hdv.th,
  }))

  const barDataTD = members.slice(0, 18).map((m) => ({
    name: m.name,
    fullName: m.fullName,
    ĐK: m.td.dk,
    KH: m.td.kh,
    TH: m.td.th,
  }))

  const pctData = useMemo(() => {
    return [...members]
      .sort((a, b) => getNumericPct(b[tab]) - getNumericPct(a[tab]))
      .slice(0, 12)
      .map((m) => ({
        name: m.name,
        fullName: m.fullName,
        pct: getNumericPct(m[tab]),
        th: m[tab].th,
      }))
  }, [members, tab])

  return (
    <div
      style={{
        padding: "22px 26px",
        display: "flex",
        flexDirection: "column",
        gap: 20,
      }}
    >
      {/* KPI Row */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
          gap: 14,
        }}
      >
        <KPICard
          label="HĐV Thực hiện"
          value={`${fmt1(hdvTh)} tỷ`}
          sub={`ĐK: ${fmt1(hdvDk)} · KH: ${fmt1(hdvKh)}`}
          accent="#006b68"
        />
        <KPICard
          label="HĐV % Hoàn thành"
          value={`${hdvPct.toFixed(1)}%`}
          sub="Trung bình CBKH"
          subColor={pctColor(hdvPct)}
          accent="#006b68"
        />
        <KPICard
          label="HĐV Còn lại"
          value={`${fmt1(hdvConTh)} tỷ`}
          sub="Cần thực hiện thêm"
          subColor="#d97706"
          accent="#006b68"
        />
        <KPICard
          label="TD Thực hiện"
          value={`${fmt1(tdTh)} tỷ`}
          sub={`ĐK: ${fmt1(tdDk)} · KH: ${fmt1(tdKh)}`}
          accent="#fdb71a"
        />
        <KPICard
          label="TD % Hoàn thành"
          value={`${tdPct.toFixed(1)}%`}
          sub="Trung bình CBKH"
          subColor={pctColor(tdPct)}
          accent="#fdb71a"
        />
        <KPICard
          label="TD Còn lại"
          value={`${fmt1(tdConTh)} tỷ`}
          sub="Cần thực hiện thêm"
          subColor="#d97706"
          accent="#fdb71a"
        />
      </div>

      {/* Tab Switch */}
      <div
        style={{
          display: "flex",
          gap: 0,
          background: "#e2e8f0",
          borderRadius: 8,
          padding: 3,
          width: "fit-content",
        }}
      >
        {(["hdv", "td"] as const).map((t) => (
          <button
            key={t}
            onClick={() => setTab(t)}
            style={{
              padding: "8px 24px",
              border: "none",
              borderRadius: 6,
              cursor: "pointer",
              fontSize: 13,
              fontWeight: 700,
              background: tab === t ? "#fff" : "transparent",
              color: tab === t ? "#111827" : "#64748b",
              boxShadow: tab === t ? "0 1px 3px rgba(0,0,0,0.1)" : "none",
              transition: "all 0.15s",
            }}
          >
            {t === "hdv" ? "Huy Động Vốn (HĐV)" : "Tín Dụng (TD)"}
          </button>
        ))}
      </div>

      {/* Charts Row */}
      <div
        style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16 }}
      >
        {/* Bar Chart */}
        <div
          style={{
            background: "#fff",
            borderRadius: 10,
            padding: "20px 16px 16px 8px",
            boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
          }}
        >
          <div style={{ paddingLeft: 12, marginBottom: 14 }}>
            <SectionTitle>
              {tab === "hdv"
                ? "HĐV – So sánh ĐK / KH / TH (CBKH)"
                : "TD – So sánh ĐK / KH / TH (CBKH)"}
            </SectionTitle>
            <div style={{ fontSize: 12, color: "#9ca3af" }}>
              Đơn vị tính: tỷ đồng (hiển thị 18 CBKH đầu)
            </div>
          </div>
          {shouldRenderCharts ? (
            <Suspense fallback={<ChartLoadingPlaceholder height={280} />}>
              <DashboardChart
                data={tab === "hdv" ? barDataHDV : barDataTD}
                actualFill={tab === "hdv" ? "#005855" : "#fdb71a"}
              />
            </Suspense>
          ) : (
            <ChartLoadingPlaceholder height={280} />
          )}
        </div>

        {/* % Hoàn thành Top List */}
        <div
          style={{
            background: "#fff",
            borderRadius: 10,
            padding: "20px 20px 16px",
            boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
            display: "flex",
            flexDirection: "column",
          }}
        >
          <div style={{ marginBottom: 14 }}>
            <SectionTitle>
              {tab === "hdv"
                ? "HĐV – % Hoàn thành dẫn đầu"
                : "TD – % Hoàn thành dẫn đầu"}
            </SectionTitle>
            <div style={{ fontSize: 12, color: "#9ca3af" }}>
              Top CBKH đạt tiến độ cao nhất
            </div>
          </div>
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 10,
              overflowY: "auto",
              maxHeight: 280,
            }}
          >
            {pctData.map((d) => (
              <div key={d.fullName}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    marginBottom: 3,
                  }}
                >
                  <span
                    style={{ fontSize: 12, color: "#1e293b", fontWeight: 600 }}
                  >
                    {d.fullName}
                  </span>
                  <span style={{ fontSize: 11, color: "#64748b" }}>
                    TH: {fmt1(d.th)} tỷ
                  </span>
                </div>
                <PctBar pct={d.pct} />
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Summary Table */}
      <div
        style={{
          background: "#fff",
          borderRadius: 10,
          boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
          overflow: "hidden",
        }}
      >
        <div
          style={{
            padding: "16px 20px",
            borderBottom: "1px solid #f1f5f9",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <SectionTitle>
            {tab === "hdv"
              ? "Bảng tổng hợp HĐV theo CBKH"
              : "Bảng tổng hợp TD theo CBKH"}
          </SectionTitle>
          <span style={{ fontSize: 12, color: "#94a3b8" }}>
            Tổng cộng {members.length} CBKH
          </span>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: "#f8fafc" }}>
                {[
                  "CBKH",
                  "Họ và tên",
                  "Đơn vị",
                  "ĐK",
                  "KH",
                  "Giao tăng",
                  "TH",
                  "Còn TH",
                  "+/- ĐK",
                  "+/- ĐT",
                  "+/- T.Trước",
                  "% HT",
                ].map((h) => (
                  <th
                    key={h}
                    style={{
                      padding: "10px 12px",
                      textAlign:
                        h === "CBKH" || h === "Họ và tên" || h === "Đơn vị"
                          ? "left"
                          : "right",
                      fontSize: 11,
                      fontWeight: 700,
                      color: "#64748b",
                      letterSpacing: 0.3,
                      whiteSpace: "nowrap",
                      borderBottom: "1px solid #e2e8f0",
                    }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {members.map((m, i) => {
                const d = m[tab]
                return (
                  <tr
                    key={`${m.name}-${i}`}
                    style={{
                      background: i % 2 === 0 ? "#fff" : "#fafafa",
                      transition: "background 0.1s",
                    }}
                  >
                    <td
                      style={{
                        padding: "9px 12px",
                        fontSize: 13,
                        fontWeight: 700,
                        color: "#111827",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {m.name}
                    </td>
                    <td
                      style={{
                        padding: "9px 12px",
                        fontSize: 12,
                        color: "#475569",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {m.fullName}
                    </td>
                    <td
                      style={{
                        padding: "9px 12px",
                        fontSize: 12,
                        color: "#64748b",
                      }}
                    >
                      <span
                        style={{
                          background: "#eff6ff",
                          color: "#1d4ed8",
                          borderRadius: 4,
                          padding: "2px 7px",
                          fontSize: 11,
                          fontWeight: 600,
                        }}
                      >
                        {m.unit}
                      </span>
                    </td>
                    {[d.dk, d.kh, d.giaoTang, d.th, d.conTh].map((v, vi) => (
                      <td
                        key={vi}
                        className="mono"
                        style={{
                          padding: "9px 12px",
                          fontSize: 12,
                          textAlign: "right",
                          color: "#111827",
                        }}
                      >
                        {fmt1(v)}
                      </td>
                    ))}
                    {[d.soDk, d.soDauThang, d.soTuanTruoc].map((v, vi) => (
                      <td
                        key={vi}
                        className="mono"
                        style={{
                          padding: "9px 12px",
                          fontSize: 12,
                          textAlign: "right",
                          color: diffColor(v),
                          fontWeight: 600,
                        }}
                      >
                        {v >= 0 ? "+" : ""}
                        {fmt1(v)}
                      </td>
                    ))}
                    <td style={{ padding: "9px 12px", textAlign: "right" }}>
                      <span
                        className="mono"
                        style={{
                          fontSize: 11,
                          fontWeight: 800,
                          color: pctColor(getNumericPct(d)),
                          background: `${pctColor(getNumericPct(d))}18`,
                          borderRadius: 4,
                          padding: "2px 6px",
                        }}
                      >
                        {fmtPct(d.pctHT)}
                      </span>
                    </td>
                  </tr>
                )
              })}
            </tbody>
            <tfoot>
              <tr style={{ background: "#0f172a", color: "#fff" }}>
                <td
                  style={{
                    padding: "11px 12px",
                    fontSize: 13,
                    fontWeight: 800,
                  }}
                  colSpan={3}
                >
                  TỔNG CỘNG ({members.length} CBKH)
                </td>
                {[
                  sumMetric(members, tab, "dk"),
                  sumMetric(members, tab, "kh"),
                  sumMetric(members, tab, "giaoTang"),
                  sumMetric(members, tab, "th"),
                  sumMetric(members, tab, "conTh"),
                  sumMetric(members, tab, "soDk"),
                  sumMetric(members, tab, "soDauThang"),
                  sumMetric(members, tab, "soTuanTruoc"),
                ].map((v, i) => (
                  <td
                    key={i}
                    className="mono"
                    style={{
                      padding: "11px 12px",
                      fontSize: 12,
                      fontWeight: 800,
                      textAlign: "right",
                      color: i >= 5 ? (v >= 0 ? "#6ee7b7" : "#fca5a5") : "#fff",
                    }}
                  >
                    {i >= 5 && v >= 0 ? "+" : ""}
                    {fmt1(v)}
                  </td>
                ))}
                <td
                  className="mono"
                  style={{
                    padding: "11px 12px",
                    fontSize: 12,
                    fontWeight: 800,
                    textAlign: "right",
                    color: "#6ee7b7",
                  }}
                >
                  {avgMetricPct(members, tab).toFixed(1)}%
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  )
}

// ── BIỂU ĐỒ PAGE ──────────────────────────────────────────────────────────────

type ChartTabId = "compare" | "unit" | "pct" | "conth" | "perday"

/** Tính số ngày làm việc (Thứ 2–6) còn lại từ ngày mai đến 31/12 của năm hiện tại. */
function calcWorkingDaysLeft(): number {
  const today = new Date()
  const yearEnd = new Date(today.getFullYear(), 11, 31)
  let count = 0
  const cur = new Date(today)
  cur.setDate(cur.getDate() + 1) // bắt đầu tính từ ngày mai
  while (cur <= yearEnd) {
    const dow = cur.getDay()
    if (dow !== 0 && dow !== 6) count++ // bỏ Chủ nhật (0) và Thứ 7 (6)
    cur.setDate(cur.getDate() + 1)
  }
  return Math.max(count, 1)
}

function BieuDoPage({ members, units }: { members: Member[]; units: string[] }) {
  const [chartType, setChartType] = useState<ChartTabId>("compare")
  const [metric, setMetric] = useState<"hdv" | "td">("hdv")
  const [workingDays, setWorkingDays] = useState<number>(calcWorkingDaysLeft)
  const [groupByUnit, setGroupByUnit] = useState<boolean>(false)

  const safeDays = Math.max(workingDays, 1)
  const accentColor = metric === "hdv" ? "#006b68" : "#fdb71a"

  // ── Data: các tab cũ ──────────────────────────────────────────────────────
  const compareData = members.slice(0, 20).map((m) => ({
    name: m.name,
    ĐK: m[metric].dk,
    KH: m[metric].kh,
    TH: m[metric].th,
  }))

  const unitData = units
    .map((u) => {
      const ums = members.filter((m) => m.unit === u)
      return {
        name: u,
        KH: ums.reduce((acc, m) => acc + m[metric].kh, 0),
        TH: ums.reduce((acc, m) => acc + m[metric].th, 0),
      }
    })
    .filter((d) => d.TH > 0 || d.KH > 0)

  const pctData = [...members]
    .sort((a, b) => getNumericPct(b[metric]) - getNumericPct(a[metric]))
    .slice(0, 16)
    .map((m) => ({
      name: m.name,
      fullName: m.fullName,
      "% HT": parseFloat(getNumericPct(m[metric]).toFixed(1)),
    }))

  // ── Data: Còn thực hiện ───────────────────────────────────────────────────
  const conThByMember = useMemo(
    () =>
      [...members]
        .sort((a, b) => b[metric].conTh - a[metric].conTh)
        .map((m) => ({
          name: m.name,
          fullName: m.fullName,
          unit: m.unit,
          "Còn TH": parseFloat(m[metric].conTh.toFixed(2)),
        })),
    [members, metric],
  )

  const conThByUnit = useMemo(
    () =>
      units
        .map((u) => {
          const ums = members.filter((m) => m.unit === u)
          return {
            name: u,
            "Còn TH": parseFloat(
              ums.reduce((acc, m) => acc + m[metric].conTh, 0).toFixed(2),
            ),
          }
        })
        .filter((d) => d["Còn TH"] > 0)
        .sort((a, b) => b["Còn TH"] - a["Còn TH"]),
    [units, members, metric],
  )

  // ── Data: Phấn đấu/ngày ───────────────────────────────────────────────────
  const perDayByMember = useMemo(
    () =>
      [...members]
        .sort((a, b) => b[metric].conTh - a[metric].conTh)
        .map((m) => ({
          name: m.name,
          fullName: m.fullName,
          unit: m.unit,
          "Tỷ/ngày": parseFloat((m[metric].conTh / safeDays).toFixed(4)),
          conTh: m[metric].conTh,
        })),
    [members, metric, safeDays],
  )

  const perDayByUnit = useMemo(
    () =>
      units
        .map((u) => {
          const ums = members.filter((m) => m.unit === u)
          const total = ums.reduce((acc, m) => acc + m[metric].conTh, 0)
          return {
            name: u,
            "Tỷ/ngày": parseFloat((total / safeDays).toFixed(4)),
            conTh: total,
          }
        })
        .filter((d) => d.conTh > 0)
        .sort((a, b) => b["Tỷ/ngày"] - a["Tỷ/ngày"]),
    [units, members, metric, safeDays],
  )

  // ── Active data (dựa vào groupByUnit toggle) ──────────────────────────────
  const activeConTh = groupByUnit ? conThByUnit : conThByMember
  const activePerDay = groupByUnit ? perDayByUnit : perDayByMember

  const totalConTh = conThByMember.reduce((a, d) => a + d["Còn TH"], 0)
  const avgConTh =
    activeConTh.reduce((a, d) => a + d["Còn TH"], 0) /
    Math.max(activeConTh.length, 1)
  const avgPerDay =
    activePerDay.reduce((a, d) => a + d["Tỷ/ngày"], 0) /
    Math.max(activePerDay.length, 1)

  // ── Tab definitions ───────────────────────────────────────────────────────
  const CHART_TABS: { id: ChartTabId; label: string; emoji: string }[] = [
    { id: "compare", label: "ĐK / KH / TH", emoji: "📊" },
    { id: "unit", label: "Theo đơn vị", emoji: "🏢" },
    { id: "pct", label: "% Hoàn thành", emoji: "✅" },
    { id: "conth", label: "Còn thực hiện", emoji: "⏳" },
    { id: "perday", label: "Phấn đấu/ngày", emoji: "🎯" },
  ]

  const isNewTab = chartType === "conth" || chartType === "perday"

  return (
    <div
      style={{
        padding: "22px 26px",
        display: "flex",
        flexDirection: "column",
        gap: 20,
      }}
    >
      {/* ── Header controls ── */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          flexWrap: "wrap",
          gap: 12,
        }}
      >
        {/* Chart type tabs */}
        <div
          style={{
            display: "flex",
            gap: 0,
            background: "#e2e8f0",
            borderRadius: 8,
            padding: 3,
            flexWrap: "wrap",
          }}
        >
          {CHART_TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => setChartType(t.id)}
              style={{
                padding: "8px 14px",
                border: "none",
                borderRadius: 6,
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 700,
                background: chartType === t.id ? "#fff" : "transparent",
                color: chartType === t.id ? "#111827" : "#64748b",
                boxShadow:
                  chartType === t.id ? "0 1px 3px rgba(0,0,0,0.1)" : "none",
                transition: "all 0.15s",
                whiteSpace: "nowrap",
              }}
            >
              {t.emoji} {t.label}
            </button>
          ))}
        </div>

        {/* Right-side controls */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            flexWrap: "wrap",
          }}
        >
          {/* Working days input — chỉ hiện cho tab Còn TH / Phấn đấu/ngày */}
          {isNewTab && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                background: "#fffbeb",
                border: "1px solid #fbbf24",
                borderRadius: 8,
                padding: "6px 12px",
              }}
            >
              <span style={{ fontSize: 12, color: "#92400e", fontWeight: 600 }}>
                📅 Còn
              </span>
              <input
                type="number"
                min={1}
                max={365}
                value={workingDays}
                onChange={(e) =>
                  setWorkingDays(Math.max(1, parseInt(e.target.value) || 1))
                }
                style={{
                  width: 52,
                  border: "1px solid #fbbf24",
                  borderRadius: 5,
                  padding: "3px 6px",
                  fontSize: 13,
                  fontWeight: 800,
                  color: "#92400e",
                  background: "transparent",
                  textAlign: "center",
                  outline: "none",
                  fontFamily: "'JetBrains Mono', monospace",
                }}
              />
              <span style={{ fontSize: 12, color: "#92400e", fontWeight: 600 }}>
                ngày làm việc
              </span>
            </div>
          )}

          {/* Theo CBKH / Theo đơn vị toggle — chỉ hiện cho tab mới */}
          {isNewTab && (
            <div
              style={{
                display: "flex",
                gap: 0,
                background: "#e2e8f0",
                borderRadius: 8,
                padding: 3,
              }}
            >
              {[
                { val: false, label: "👤 Theo CBKH" },
                { val: true, label: "🏢 Theo đơn vị" },
              ].map((opt) => (
                <button
                  key={String(opt.val)}
                  onClick={() => setGroupByUnit(opt.val)}
                  style={{
                    padding: "6px 12px",
                    border: "none",
                    borderRadius: 6,
                    cursor: "pointer",
                    fontSize: 12,
                    fontWeight: 700,
                    background:
                      groupByUnit === opt.val ? accentColor : "transparent",
                    color: groupByUnit === opt.val ? "#fff" : "#64748b",
                    transition: "all 0.15s",
                    whiteSpace: "nowrap",
                  }}
                >
                  {opt.label}
                </button>
              ))}
            </div>
          )}

          {/* HĐV / TD metric toggle */}
          <div
            style={{
              display: "flex",
              gap: 0,
              background: "#e2e8f0",
              borderRadius: 8,
              padding: 3,
            }}
          >
            {(["hdv", "td"] as const).map((t) => (
              <button
                key={t}
                onClick={() => setMetric(t)}
                style={{
                  padding: "8px 20px",
                  border: "none",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 13,
                  fontWeight: 700,
                  background:
                    metric === t
                      ? t === "hdv"
                        ? "#006b68"
                        : "#fdb71a"
                      : "transparent",
                  color: metric === t
                    ? t === "hdv" ? "#fff" : "#1a1a1a"
                    : "#64748b",
                  transition: "all 0.15s",
                }}
              >
                {t === "hdv" ? "Huy Động Vốn" : "Tín Dụng"}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* ── Main chart card ── */}
      <div
        style={{
          background: "#fff",
          borderRadius: 10,
          padding: "22px 16px 16px 8px",
          boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
        }}
      >
        <div style={{ paddingLeft: 12, marginBottom: 16 }}>
          <SectionTitle>
            {metric === "hdv" ? "HUY ĐỘNG VỐN" : "TÍN DỤNG"} –{" "}
            {chartType === "compare"
              ? "So sánh ĐK / KH / TH theo CBKH"
              : chartType === "unit"
                ? "KH và TH theo đơn vị"
                : chartType === "pct"
                  ? "% Hoàn thành kế hoạch (xếp hạng)"
                  : chartType === "conth"
                    ? `Số còn thực hiện ${groupByUnit ? "theo đơn vị" : "theo từng CBKH"}`
                    : `Chỉ tiêu phấn đấu mỗi ngày ${groupByUnit ? "theo đơn vị" : "theo CBKH"}`}
          </SectionTitle>
          <div style={{ fontSize: 12, color: "#9ca3af" }}>
            {chartType === "compare" && "Đơn vị: tỷ đồng — Top 20 CBKH"}
            {chartType === "unit" && "Đơn vị: tỷ đồng — tổng hợp theo phòng/PGD"}
            {chartType === "pct" && "Xếp hạng từ cao → thấp — Top 16 CBKH"}
            {chartType === "conth" &&
              `Đơn vị: tỷ đồng — 🔴 > 150% TB · 🟡 > TB · 🟢 ≤ TB`}
            {chartType === "perday" &&
              `Tỷ/ngày = Còn TH ÷ ${workingDays} ngày làm việc · 🔴 > 150% TB · 🟡 > TB · 🟢 ≤ TB`}
          </div>
        </div>

        {/* ── Tab: ĐK / KH / TH ── */}
        {chartType === "compare" && (
          <Suspense fallback={<ChartLoadingPlaceholder height={340} />}>
            <BieuDoChart
              kind="compare"
              data={compareData}
              actualFill={metric === "hdv" ? "#005855" : "#fdb71a"}
            />
          </Suspense>
        )}

        {/* ── Tab: Theo đơn vị ── */}
        {chartType === "unit" && (
          <Suspense fallback={<ChartLoadingPlaceholder height={340} />}>
            <BieuDoChart
              kind="unit"
              data={unitData}
              actualFill={metric === "hdv" ? "#005855" : "#fdb71a"}
            />
          </Suspense>
        )}

        {/* ── Tab: % Hoàn thành ── */}
        {chartType === "pct" && (
          <div
            style={{
              padding: "0 14px",
              display: "flex",
              flexDirection: "column",
              gap: 12,
            }}
          >
            {pctData.map((d, i) => (
              <div key={`${d.name}-${i}`}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    marginBottom: 4,
                  }}
                >
                  <span
                    style={{ fontSize: 13, color: "#1e293b", fontWeight: 600 }}
                  >
                    {d.fullName} ({d.name})
                  </span>
                  <span
                    className="mono"
                    style={{
                      fontSize: 12,
                      color: pctColor(d["% HT"]),
                      fontWeight: 800,
                    }}
                  >
                    {d["% HT"]}%
                  </span>
                </div>
                <div
                  style={{
                    height: 8,
                    background: "#e2e8f0",
                    borderRadius: 4,
                    overflow: "hidden",
                  }}
                >
                  <div
                    style={{
                      width: `${Math.min(Math.max(d["% HT"], 0), 100)}%`,
                      height: "100%",
                      borderRadius: 4,
                      background: pctColor(d["% HT"]),
                      transition: "width 0.4s",
                    }}
                  />
                </div>
              </div>
            ))}
          </div>
        )}

        {/* ── Tab: Còn thực hiện (horizontal bar) ── */}
        {chartType === "conth" && (
          <Suspense
            fallback={
              <ChartLoadingPlaceholder
                height={Math.max(320, activeConTh.length * 26)}
              />
            }
          >
            <BieuDoChart
              kind="conth"
              data={activeConTh}
              groupByUnit={groupByUnit}
              average={avgConTh}
            />
          </Suspense>
        )}

        {/* ── Tab: Phấn đấu/ngày (horizontal bar) ── */}
        {chartType === "perday" && (
          <Suspense
            fallback={
              <ChartLoadingPlaceholder
                height={Math.max(320, activePerDay.length * 26)}
              />
            }
          >
            <BieuDoChart
              kind="perday"
              data={activePerDay}
              groupByUnit={groupByUnit}
              average={avgPerDay}
            />
          </Suspense>
        )}
      </div>

      {/* ── Summary strip (chỉ hiện khi tab mới) ── */}
      {isNewTab && (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))",
            gap: 14,
          }}
        >
          {/* Tổng còn TH */}
          <div
            style={{
              background: "#fff",
              borderRadius: 10,
              padding: "16px 20px",
              boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
              borderTop: `3px solid ${accentColor}`,
            }}
          >
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: "#9ca3af",
                letterSpacing: 0.8,
                textTransform: "uppercase",
              }}
            >
              Tổng còn thực hiện
            </div>
            <div
              className="mono"
              style={{ fontSize: 22, fontWeight: 800, color: "#111827", marginTop: 4 }}
            >
              {fmt1(totalConTh)} tỷ
            </div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 3 }}>
              {metric === "hdv" ? "HĐV" : "Tín dụng"} — toàn chi nhánh
            </div>
          </div>

          {/* Ngày còn lại */}
          <div
            style={{
              background: "#fff",
              borderRadius: 10,
              padding: "16px 20px",
              boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
              borderTop: "3px solid #f59e0b",
            }}
          >
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: "#9ca3af",
                letterSpacing: 0.8,
                textTransform: "uppercase",
              }}
            >
              Ngày làm việc còn lại
            </div>
            <div
              className="mono"
              style={{ fontSize: 22, fontWeight: 800, color: "#92400e", marginTop: 4 }}
            >
              {workingDays} ngày
            </div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 3 }}>
              Từ hôm nay → 31/12/{new Date().getFullYear()} (T2–T6)
            </div>
          </div>

          {/* TB/ngày toàn CN */}
          <div
            style={{
              background: "#fff",
              borderRadius: 10,
              padding: "16px 20px",
              boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
              borderTop: "3px solid #dc2626",
            }}
          >
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: "#9ca3af",
                letterSpacing: 0.8,
                textTransform: "uppercase",
              }}
            >
              Phấn đấu TB toàn CN/ngày
            </div>
            <div
              className="mono"
              style={{ fontSize: 22, fontWeight: 800, color: "#dc2626", marginTop: 4 }}
            >
              {(totalConTh / safeDays).toFixed(3)} tỷ/ngày
            </div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 3 }}>
              {fmt1(totalConTh)} tỷ ÷ {workingDays} ngày
            </div>
          </div>

          {/* Top 3 áp lực */}
          <div
            style={{
              background: "#fff",
              borderRadius: 10,
              padding: "16px 20px",
              boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
              borderTop: "3px solid #7c3aed",
            }}
          >
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: "#9ca3af",
                letterSpacing: 0.8,
                textTransform: "uppercase",
                marginBottom: 6,
              }}
            >
              Top 3 CBKH còn nhiều nhất
            </div>
            {conThByMember.slice(0, 3).map((d, i) => (
              <div
                key={`${d.name}-${i}`}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  marginTop: i === 0 ? 0 : 5,
                }}
              >
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    color: i === 0 ? "#dc2626" : i === 1 ? "#d97706" : "#374151",
                  }}
                >
                  {i + 1}. {d.name}
                </span>
                <span
                  className="mono"
                  style={{
                    fontSize: 12,
                    color: i === 0 ? "#dc2626" : "#64748b",
                    fontWeight: 700,
                  }}
                >
                  {fmt1(d["Còn TH"])} tỷ
                </span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}


// ── SỐ LIỆU PAGE ──────────────────────────────────────────────────────────────

function SoLieuPage({ members, units }: { members: Member[]; units: string[] }) {
  const [tab, setTab] = useState<"hdv" | "td">("hdv")

  return (
    <div
      style={{
        padding: "22px 26px",
        display: "flex",
        flexDirection: "column",
        gap: 20,
      }}
    >
      <div
        style={{
          display: "flex",
          gap: 0,
          background: "#e2e8f0",
          borderRadius: 8,
          padding: 3,
          width: "fit-content",
        }}
      >
        {(["hdv", "td"] as const).map((t) => (
          <button
            key={t}
            onClick={() => setTab(t)}
            style={{
              padding: "8px 24px",
              border: "none",
              borderRadius: 6,
              cursor: "pointer",
              fontSize: 13,
              fontWeight: 700,
              background: tab === t ? "#fff" : "transparent",
              color: tab === t ? "#111827" : "#64748b",
              boxShadow: tab === t ? "0 1px 3px rgba(0,0,0,0.1)" : "none",
              transition: "all 0.15s",
            }}
          >
            {t === "hdv" ? "Huy Động Vốn" : "Tín Dụng"}
          </button>
        ))}
      </div>

      <div
        style={{
          background: "#fff",
          borderRadius: 10,
          boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
          overflow: "hidden",
        }}
      >
        <div
          style={{
            padding: "16px 20px",
            borderBottom: "1px solid #f1f5f9",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <SectionTitle>
            Bảng số liệu chi tiết phân bổ theo đơn vị –{" "}
            {tab === "hdv" ? "Huy Động Vốn" : "Tín Dụng"}
          </SectionTitle>
          <span style={{ fontSize: 12, color: "#94a3b8" }}>
            Đơn vị: tỷ đồng · {members.length} CBKH
          </span>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: "#f8fafc" }}>
                <th
                  style={{
                    padding: "10px 12px",
                    textAlign: "left",
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#64748b",
                    letterSpacing: 0.3,
                    borderBottom: "1px solid #e2e8f0",
                    whiteSpace: "nowrap",
                  }}
                >
                  CBKH
                </th>
                <th
                  style={{
                    padding: "10px 12px",
                    textAlign: "left",
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#64748b",
                    letterSpacing: 0.3,
                    borderBottom: "1px solid #e2e8f0",
                  }}
                >
                  HỌ VÀ TÊN
                </th>
                <th
                  style={{
                    padding: "10px 12px",
                    textAlign: "left",
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#64748b",
                    letterSpacing: 0.3,
                    borderBottom: "1px solid #e2e8f0",
                  }}
                >
                  ĐƠN VỊ
                </th>
                {[
                  "ĐK",
                  "KH",
                  "GIAO TĂNG",
                  "TH",
                  "CÒN TH",
                  "+/- ĐK",
                  "+/- ĐẦU THÁNG",
                  "+/- TUẦN TRƯỚC",
                  "% HT",
                ].map((h) => (
                  <th
                    key={h}
                    style={{
                      padding: "10px 12px",
                      textAlign: "right",
                      fontSize: 11,
                      fontWeight: 700,
                      color: "#64748b",
                      letterSpacing: 0.3,
                      borderBottom: "1px solid #e2e8f0",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {units.flatMap((unit) => {
                const ums = members.filter((m) => m.unit === unit)
                if (!ums.length) return []
                const unitRow = (
                  <tr
                    key={`unit-${unit}`}
                    style={{ background: "#006b68", color: "#fff" }}
                  >
                    <td
                      colSpan={3}
                      style={{
                        padding: "8px 12px",
                        fontSize: 12,
                        fontWeight: 800,
                        letterSpacing: 0.8,
                      }}
                    >
                      {unit}
                    </td>
                    {[
                      sumMetric(ums, tab, "dk"),
                      sumMetric(ums, tab, "kh"),
                      sumMetric(ums, tab, "giaoTang"),
                      sumMetric(ums, tab, "th"),
                      sumMetric(ums, tab, "conTh"),
                      sumMetric(ums, tab, "soDk"),
                      sumMetric(ums, tab, "soDauThang"),
                      sumMetric(ums, tab, "soTuanTruoc"),
                    ].map((v, i) => (
                      <td
                        key={i}
                        className="mono"
                        style={{
                          padding: "8px 12px",
                          fontSize: 12,
                          fontWeight: 800,
                          textAlign: "right",
                          color:
                            i >= 5 ? (v >= 0 ? "#6ee7b7" : "#fca5a5") : "#fff",
                        }}
                      >
                        {i >= 5 && v >= 0 ? "+" : ""}
                        {fmt1(v)}
                      </td>
                    ))}
                    <td
                      className="mono"
                      style={{
                        padding: "8px 12px",
                        fontSize: 12,
                        fontWeight: 800,
                        textAlign: "right",
                        color: "#6ee7b7",
                      }}
                    >
                      {avgMetricPct(ums, tab).toFixed(1)}%
                    </td>
                  </tr>
                )

                const memberRows = ums.map((m, i) => {
                  const d = m[tab]
                  return (
                    <tr
                      key={`${unit}-${m.name}-${i}`}
                      style={{ background: i % 2 === 0 ? "#fff" : "#fafafa" }}
                    >
                      <td
                        style={{
                          padding: "9px 12px",
                          fontSize: 13,
                          fontWeight: 700,
                          color: "#111827",
                          whiteSpace: "nowrap",
                        }}
                      >
                        {m.name}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          fontSize: 12,
                          color: "#475569",
                          whiteSpace: "nowrap",
                        }}
                      >
                        {m.fullName}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          fontSize: 11,
                          color: "#64748b",
                        }}
                      >
                        {m.unit}
                      </td>
                      {[d.dk, d.kh, d.giaoTang, d.th, d.conTh].map((v, vi) => (
                        <td
                          key={vi}
                          className="mono"
                          style={{
                            padding: "9px 12px",
                            fontSize: 12,
                            textAlign: "right",
                            color: "#111827",
                          }}
                        >
                          {fmt1(v)}
                        </td>
                      ))}
                      {[d.soDk, d.soDauThang, d.soTuanTruoc].map((v, vi) => (
                        <td
                          key={vi}
                          className="mono"
                          style={{
                            padding: "9px 12px",
                            fontSize: 12,
                            textAlign: "right",
                            color: diffColor(v),
                            fontWeight: 600,
                          }}
                        >
                          {v >= 0 ? "+" : ""}
                          {fmt1(v)}
                        </td>
                      ))}
                      <td style={{ padding: "9px 12px", textAlign: "right" }}>
                        <span
                          className="mono"
                          style={{
                            fontSize: 11,
                            fontWeight: 800,
                            color: pctColor(getNumericPct(d)),
                            background: `${pctColor(getNumericPct(d))}18`,
                            borderRadius: 4,
                            padding: "2px 6px",
                          }}
                        >
                          {fmtPct(d.pctHT)}
                        </span>
                      </td>
                    </tr>
                  )
                })

                return [unitRow, ...memberRows]
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

// ── EXACT PDF REPORT TABLE (MATCHING PDF SAMPLE 100%) ─────────────────────────

function SingleReportTable({
  rows,
  pageTitle,
  isPortrait = true,
}: {
  rows: ReportRow[]
  pageTitle?: string
  isPortrait?: boolean
}) {
  const border = "1px solid #000000"
  const baseCell: React.CSSProperties = {
    border,
    padding: isPortrait ? "2px 1px" : "2px 1.5px",
    textAlign: "right",
    fontSize: isPortrait ? 7.2 : 8.5,
    lineHeight: 1.15,
    whiteSpace: "nowrap",
    verticalAlign: "middle",
    letterSpacing: isPortrait ? "-0.2px" : "normal",
  }
  const nameCellLeft: React.CSSProperties = {
    ...baseCell,
    textAlign: "left",
    fontWeight: 600,
    fontSize: isPortrait ? 7.2 : 8.5,
    paddingLeft: isPortrait ? 2 : 3,
  }
  const nameCellRight: React.CSSProperties = {
    ...baseCell,
    textAlign: "right",
    fontWeight: 600,
    fontSize: isPortrait ? 7.2 : 8.5,
    paddingRight: isPortrait ? 2 : 3,
  }

  const headers = [
    { label: "Tên", width: "5.8%", align: "center" as const },
    { label: "HDV ĐK", width: "4.95%", align: "center" as const },
    { label: "HDV KH", width: "4.95%", align: "center" as const },
    {
      label: "Giao\ntăng (*)",
      width: "4.95%",
      align: "center" as const,
      color: "#d90099",
    },
    { label: "HĐV TH", width: "4.95%", align: "center" as const },
    { label: "Còn TH\nHĐV", width: "4.95%", align: "center" as const },
    { label: "'+/- so\nĐK", width: "4.95%", align: "center" as const },
    { label: "'+/- so\nĐầu\ntháng", width: "4.95%", align: "center" as const },
    {
      label: "'+/- so\ntuần\ntrước",
      width: "4.95%",
      align: "center" as const,
      bg: "#ffff00",
    },
    { label: "% HT", width: "4.4%", align: "center" as const },
    { label: "TD ĐK", width: "4.95%", align: "center" as const },
    { label: "TD KH", width: "4.95%", align: "center" as const },
    {
      label: "Giao\ntăng (*)",
      width: "4.95%",
      align: "center" as const,
      color: "#d90099",
    },
    { label: "TD TH", width: "4.95%", align: "center" as const, bg: "#ffff00" },
    { label: "Còn\nTHTD", width: "4.95%", align: "center" as const },
    { label: "'+/- so\nĐK", width: "4.95%", align: "center" as const },
    { label: "'+/- so\nĐầu\ntháng", width: "4.95%", align: "center" as const },
    {
      label: "'+/- so\ntuần\ntrước",
      width: "4.95%",
      align: "center" as const,
      bg: "#ffff00",
    },
    { label: "% HT", width: "4.4%", align: "center" as const },
    { label: "Tên", width: "5.8%", align: "center" as const },
  ]

  const renderMetric = (m: Metric, prefix: string, kind: ReportRow["kind"]) => {
    const isUnit = kind === "unit"
    const isMgr = kind === "manager"
    const isTot = kind === "total"
    const isSummary = isUnit || isMgr || isTot

    const rowBg = isTot
      ? "#ffff00"
      : isMgr
        ? "#088f8f"
        : isUnit
          ? "#00b074"
          : undefined
    const rowColor = isTot ? "#075b64" : isSummary ? "#ffffff" : undefined

    // Values: 0:dk, 1:kh, 2:giaoTang, 3:th, 4:conTh, 5:soDk, 6:soDauThang, 7:soTuanTruoc
    const vals = [
      m.dk,
      m.kh,
      m.giaoTang,
      m.th,
      m.conTh,
      m.soDk,
      m.soDauThang,
      m.soTuanTruoc,
    ]

    return [
      ...vals.map((val, idx) => {
        let cellBg = rowBg
        let cellColor = rowColor
        let cellWeight = isSummary ? 800 : 400

        if (!isSummary) {
          // Member row specific styles:
          if (idx === 2) {
            // Giao tăng (*)
            cellColor = "#d90099"
          } else if (idx === 4) {
            // Còn TH
            cellBg = "#48c774"
            cellColor = "#ffffff"
          } else if (idx === 5) {
            // +/- so ĐK: amber if > 0
            if (val > 0) cellBg = "#fec84e"
          } else if (idx === 7) {
            // +/- so tuần trước: yellow
            cellBg = "#ffff00"
          }
        }

        return (
          <td
            key={`${prefix}-${idx}`}
            className="mono"
            style={{
              ...baseCell,
              background: cellBg,
              color: cellColor,
              fontWeight: cellWeight,
            }}
          >
            {fmt1(val)}
          </td>
        )
      }),
      // % HT column
      <td
        key={`${prefix}-pct`}
        className="mono"
        style={{
          ...baseCell,
          background: rowBg,
          color: isSummary ? rowColor : "#008000",
          fontWeight: 800,
        }}
      >
        {fmtPct(m.pctHT)}
      </td>,
    ]
  }

  return (
    <div
      className="report-table-wrapper"
      style={{ width: "100%", overflowX: "auto" }}
    >
      <table
        className="report-table"
        style={{
          width: "100%",
          borderCollapse: "collapse",
          tableLayout: "fixed",
          fontFamily: "Arial, sans-serif",
          background: "#fff",
        }}
      >
        <colgroup>
          {headers.map((h, i) => (
            <col key={i} style={{ width: h.width }} />
          ))}
        </colgroup>
        <thead>
          <tr style={{ background: "#ffffff" }}>
            {headers.map((h, i) => (
              <th
                key={i}
                style={{
                  ...baseCell,
                  textAlign: "center",
                  fontSize: isPortrait ? 6.8 : 7.8,
                  fontWeight: 800,
                  height: isPortrait ? 30 : 28,
                  padding: isPortrait ? "1px 0.5px" : "1px 1px",
                  lineHeight: 1.05,
                  whiteSpace: "pre-line",
                  background: h.bg || "#ffffff",
                  color: h.color || "#000000",
                }}
              >
                {h.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rIdx) => {
            const isUnit = row.kind === "unit"
            const isMgr = row.kind === "manager"
            const isTot = row.kind === "total"
            const isSummary = isUnit || isMgr || isTot

            const rowBg = isTot
              ? "#ffff00"
              : isMgr
                ? "#088f8f"
                : isUnit
                  ? "#00b074"
                  : undefined
            const rowColor = isTot
              ? "#075b64"
              : isSummary
                ? "#ffffff"
                : undefined

            return (
              <tr
                key={`row-${rIdx}-${row.name}`}
                style={rowBg ? { background: rowBg } : undefined}
              >
                <td
                  style={{
                    ...nameCellLeft,
                    background: rowBg,
                    color: rowColor,
                    fontWeight: isSummary ? 800 : 600,
                  }}
                >
                  {row.name}
                </td>
                {renderMetric(row.hdv, `hdv-${rIdx}`, row.kind)}
                {renderMetric(row.td, `td-${rIdx}`, row.kind)}
                <td
                  style={{
                    ...nameCellRight,
                    background: rowBg,
                    color: rowColor,
                    fontWeight: isSummary ? 800 : 600,
                  }}
                >
                  {row.name}
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

// ── SPLIT ROWS BY PAGES (EXACTLY 3 PAGES MATCHING THE MASTER SAMPLE) ─────────

function splitPages(
  rows: ReportRow[],
): { page1: ReportRow[]; page2: ReportRow[]; page3: ReportRow[] } {
  const p1: ReportRow[] = []
  const p2: ReportRow[] = []
  const p3: ReportRow[] = []
  let curPage = 1

  for (const r of rows) {
    if (curPage === 1) {
      p1.push(r)
      if (r.kind === "manager" && /CHỊ\s*THẢO/i.test(r.name)) {
        curPage = 2
      }
    } else if (curPage === 2) {
      p2.push(r)
      if (r.kind === "manager" && /CHỊ\s*HOA/i.test(r.name)) {
        curPage = 3
      }
    } else {
      p3.push(r)
    }
  }

  if (curPage === 1 && p1.length > 30) {
    const third = Math.ceil(p1.length / 3)
    return {
      page1: p1.slice(0, third),
      page2: p1.slice(third, third * 2),
      page3: p1.slice(third * 2),
    }
  }

  return { page1: p1, page2: p2, page3: p3 }
}

// ── XUẤT PDF PAGE ─────────────────────────────────────────────────────────────

function XuatPDFPage({
  rows,
  sourceName,
  onUploadExcel,
}: {
  rows: ReportRow[]
  sourceName: string
  onUploadExcel: (e: React.ChangeEvent<HTMLInputElement>) => void
}) {
  const [pageView, setPageView] = useState<"all" | "page1" | "page2" | "page3">(
    "all",
  )
  const [printScope, setPrintScope] = useState<"all" | "current">("all")
  const [orientation, setOrientation] = useState<"portrait" | "landscape">("portrait")
  const [zoom, setZoom] = useState<number>(100)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    previewRef.current?.scrollTo({ top: 0, left: 0, behavior: "auto" })
  }, [pageView, orientation, printScope, zoom])

  const { page1, page2, page3 } = useMemo(() => splitPages(rows), [rows])

  const handlePrint = () => {
    window.print()
  }

  const pagesToRenderScreen = useMemo(() => {
    if (pageView === "page1")
      return [
        {
          num: 1,
          title: "Trang 1: Khối Chị Thảo (P.KHDN, PGD TM, PGD TB)",
          rows: page1,
        },
      ]
    if (pageView === "page2")
      return [
        {
          num: 2,
          title: "Trang 2: Khối Chị Hoa (P.KHCN, PGD HN, PGD LTK)",
          rows: page2,
        },
      ]
    if (pageView === "page3")
      return [
        {
          num: 3,
          title: "Trang 3: Khối Anh Cường (PGD TN, PGD MT, PGD TH) & Tổng cộng",
          rows: page3,
        },
      ]
    return [
      {
        num: 1,
        title: "Trang 1/3: Khối Chị Thảo (P.KHDN, PGD TM, PGD TB)",
        rows: page1,
      },
      {
        num: 2,
        title: "Trang 2/3: Khối Chị Hoa (P.KHCN, PGD HN, PGD LTK)",
        rows: page2,
      },
      {
        num: 3,
        title: "Trang 3/3: Khối Anh Cường (PGD TN, PGD MT, PGD TH) & Tổng cộng",
        rows: page3,
      },
    ]
  }, [pageView, page1, page2, page3])

  // Dedicated pages for print output: either ALL 3 pages (default) or current page
  const pagesToPrint = useMemo(() => {
    if (printScope === "current") {
      return pagesToRenderScreen
    }
    return [
      {
        num: 1,
        title: "Trang 1: Khối Chị Thảo (P.KHDN, PGD TM, PGD TB)",
        rows: page1,
      },
      {
        num: 2,
        title: "Trang 2: Khối Chị Hoa (P.KHCN, PGD HN, PGD LTK)",
        rows: page2,
      },
      {
        num: 3,
        title: "Trang 3: Khối Anh Cường (PGD TN, PGD MT, PGD TH) & Tổng cộng",
        rows: page3,
      },
    ]
  }, [printScope, pagesToRenderScreen, page1, page2, page3])
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        overflow: "auto",
      }}
    >
      {/* Dynamic page orientation style injection */}
      <style>{`
        @page {
          size: ${orientation === "portrait" ? "A4 portrait" : "A4 landscape"};
          margin: ${orientation === "portrait" ? "3mm" : "4mm"};
        }
        @page :left {
          size: ${orientation === "portrait" ? "A4 portrait" : "A4 landscape"};
        }
        @page :right {
          size: ${orientation === "portrait" ? "A4 portrait" : "A4 landscape"};
        }
      `}</style>

      {/* ── DEDICATED PRINT PORTAL ── Render toàn bộ trang in ngoài #root ── */}
      {typeof document !== "undefined" &&
        createPortal(
          <div className="print-portal">
            {pagesToPrint.map((p, idx) => (
              <div
                key={`print-${p.num}`}
                className="pdf-page-sheet"
                style={{
                  pageBreakAfter: idx === pagesToPrint.length - 1 ? "auto" : "always",
                  breakAfter: idx === pagesToPrint.length - 1 ? "auto" : "page",
                }}
              >
                <SingleReportTable
                  rows={p.rows}
                  pageTitle={p.title}
                  isPortrait={orientation === "portrait"}
                />
              </div>
            ))}
          </div>,
          document.body,
        )}

      {/* ── Control Bar (Hidden when printing) ── */}

      <div
        className="no-print"
        style={{
          background: "#ffffff",
          borderBottom: "1px solid #e2e8f0",
          padding: "10px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          flexWrap: "wrap",
          gap: 12,
          boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
          zIndex: 10,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div>
            <div
              style={{
                fontSize: 16,
                fontWeight: 800,
                color: "#0f172a",
                display: "flex",
                alignItems: "center",
                gap: 8,
              }}
            >
              <span>Xem trước & Xuất PDF</span>
              <span
                style={{
                  fontSize: 11,
                  background: orientation === "portrait" ? "#006b68" : "#005855",
                  color: "#ffffff",
                  padding: "2px 8px",
                  borderRadius: 12,
                  fontWeight: 800,
                }}
              >
                {orientation === "portrait"
                  ? "Khổ dọc A4 (3 trang - 70 dòng)"
                  : "Khổ ngang A4 (3 trang - 70 dòng)"}
              </span>
            </div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>
              Đầy đủ 57 CBKH · 9 Đơn vị · 3 Khối Quản lý · 1 Tổng cộng (
              {sourceName})
            </div>
          </div>
        </div>

        {/* Action Controls */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            flexWrap: "wrap",
          }}
        >
          {/* Orientation switch tabs */}
          <div
            style={{
              display: "flex",
              background: "#f1f5f9",
              borderRadius: 8,
              padding: 3,
            }}
          >
            {[
              { id: "portrait" as const, label: "📄 Khổ dọc" },
              { id: "landscape" as const, label: "📑 Khổ ngang" },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setOrientation(tab.id)}
                style={{
                  padding: "6px 12px",
                  border: "none",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 12,
                  fontWeight: 700,
                  background: orientation === tab.id ? "#006b68" : "transparent",
                  color: orientation === tab.id ? "#ffffff" : "#64748b",
                  boxShadow:
                    orientation === tab.id ? "0 1px 2px rgba(0,0,0,0.12)" : "none",
                  transition: "all 0.15s",
                }}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* Page switch tabs */}
          <div
            style={{
              display: "flex",
              background: "#f1f5f9",
              borderRadius: 8,
              padding: 3,
            }}
          >
            {[
              { id: "all" as const, label: "Tất cả 3 trang" },
              { id: "page1" as const, label: "Trang 1 (Chị Thảo)" },
              { id: "page2" as const, label: "Trang 2 (Chị Hoa)" },
              { id: "page3" as const, label: "Trang 3 (A.Cường & TC)" },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setPageView(tab.id)}
                style={{
                  padding: "6px 12px",
                  border: "none",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 12,
                  fontWeight: 700,
                  background: pageView === tab.id ? "#ffffff" : "transparent",
                  color: pageView === tab.id ? "#0f172a" : "#64748b",
                  boxShadow:
                    pageView === tab.id ? "0 1px 2px rgba(0,0,0,0.08)" : "none",
                  transition: "all 0.15s",
                }}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* Print Scope Selector */}
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <select
              value={printScope}
              onChange={(e) =>
                setPrintScope(e.target.value as "all" | "current")
              }
              style={{
                border: "1px solid #cbd5e1",
                borderRadius: 7,
                padding: "6px 28px 6px 10px",
                fontSize: 12,
                fontWeight: 700,
                color: "#004040",
                background: "#f8fafc",
                cursor: "pointer",
              }}
            >
              <option value="all">In toàn bộ số liệu (3 trang)</option>
              <option value="current">Chỉ in trang đang xem</option>
            </select>
          </div>

          {/* Zoom buttons */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 4,
              background: "#f1f5f9",
              borderRadius: 8,
              padding: "3px 8px",
            }}
          >
            <span
              style={{
                fontSize: 11,
                color: "#64748b",
                fontWeight: 600,
                marginRight: 2,
              }}
            >
              Thu phóng:
            </span>
            {[80, 90, 100, 110].map((z) => (
              <button
                key={z}
                onClick={() => setZoom(z)}
                style={{
                  padding: "4px 8px",
                  border: "none",
                  borderRadius: 4,
                  cursor: "pointer",
                  fontSize: 11,
                  fontWeight: 700,
                  background: zoom === z ? "#ffffff" : "transparent",
                  color: zoom === z ? "#005855" : "#64748b",
                  boxShadow: zoom === z ? "0 1px 2px rgba(0,0,0,0.08)" : "none",
                }}
              >
                {z}%
              </button>
            ))}
          </div>

          {/* Nạp Excel Button */}
          <button
            onClick={() => fileInputRef.current?.click()}
            style={{
              background: "#006b68",
              color: "#ffffff",
              border: "none",
              borderRadius: 8,
              padding: "7px 12px",
              fontSize: 12,
              fontWeight: 700,
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              gap: 6,
              boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
            }}
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="17 8 12 3 7 8" />
              <line x1="12" y1="3" x2="12" y2="15" />
            </svg>
            Nạp Excel mới
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={onUploadExcel}
            style={{ display: "none" }}
          />

          {/* Print / Export Button */}
          <button
            onClick={handlePrint}
            style={{
              background: "#006b68",
              color: "#ffffff",
              border: "none",
              borderRadius: 8,
              padding: "8px 18px",
              fontSize: 13,
              fontWeight: 800,
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              gap: 8,
              boxShadow: "0 2px 5px rgba(0,107,104,0.35)",
              transition: "background 0.15s",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.background = "#004d4b")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "#006b68")}
          >
            <svg
              width="16"
              height="16"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <polyline points="6 9 6 2 18 2 18 9" />
              <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
              <rect x="6" y="14" width="12" height="8" />
            </svg>
            {printScope === "all"
              ? orientation === "portrait"
                ? "In khổ dọc (3 trang)"
                : "In khổ ngang (3 trang)"
              : orientation === "portrait"
                ? "In trang đang xem (Khổ dọc)"
                : "In trang đang xem (Khổ ngang)"}
          </button>
        </div>
      </div>

      {/* ── Screen Document Viewport ── */}
      <div
        ref={previewRef}
        className="pdf-preview-container"
        style={{
          flex: 1,
          overflowY: "auto",
          background: "#eef2f7",
          padding: "24px 20px 40px",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          gap: 24,
          minHeight: 0,
        }}
      >
        <div
          style={{
            width: "100%",
            maxWidth: orientation === "portrait" ? 1000 : 1400,
            display: "flex",
            flexDirection: "column",
            gap: 0,
            transition: "max-width 0.25s ease",
            transformOrigin: "top center",
            transform: zoom === 100 ? "none" : `scale(${zoom / 100})`,
          }}
        >
          {pagesToRenderScreen.map((page) => (
            <div
              key={page.num}
              className="pdf-page-sheet"
              style={{
                background: "#ffffff",
                border: "none",
                borderRadius: 0,
                boxShadow: "none",
                padding: 0,
                width: "100%",
                boxSizing: "border-box",
                display: "flex",
                flexDirection: "column",
                gap: 0,
                minHeight: orientation === "portrait" ? 1120 : 760,
              }}
            >
              {/* The exact report table matching the reference PDF sample */}
              <SingleReportTable
                rows={page.rows}
                pageTitle={page.title}
                isPortrait={orientation === "portrait"}
              />
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

// ── SIDEBAR ───────────────────────────────────────────────────────────────────

const NAV: { id: Page; label: string; icon: React.ReactElement }[] = [
  {
    id: "dashboard",
    label: "Tổng quan",
    icon: (
      <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
      >
        <rect x="3" y="3" width="7" height="7" rx="1" />
        <rect x="14" y="3" width="7" height="7" rx="1" />
        <rect x="14" y="14" width="7" height="7" rx="1" />
        <rect x="3" y="14" width="7" height="7" rx="1" />
      </svg>
    ),
  },
  {
    id: "bieudo",
    label: "Biểu đồ",
    icon: (
      <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
      >
        <line x1="18" y1="20" x2="18" y2="10" />
        <line x1="12" y1="20" x2="12" y2="4" />
        <line x1="6" y1="20" x2="6" y2="14" />
        <line x1="2" y1="20" x2="22" y2="20" />
      </svg>
    ),
  },
  {
    id: "solieu",
    label: "Số liệu chi tiết",
    icon: (
      <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
      >
        <path d="M9 3H5a2 2 0 00-2 2v4m6-6h10a2 2 0 012 2v4M9 3v18m0 0h10a2 2 0 002-2V9M9 21H5a2 2 0 01-2-2V9m0 0h18" />
      </svg>
    ),
  },
  {
    id: "xuatpdf",
    label: "Xuất PDF",
    icon: (
      <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
      >
        <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
        <polyline points="14 2 14 8 20 8" />
        <line x1="16" y1="13" x2="8" y2="13" />
        <line x1="16" y1="17" x2="8" y2="17" />
        <polyline points="10 9 9 9 8 9" />
      </svg>
    ),
  },
]

function Sidebar({ page, setPage }: { page: Page; setPage: (p: Page) => void }) {
  return (
    <aside
      className="no-print"
      style={{
        width: 225,
        background: "#004845",
        display: "flex",
        flexDirection: "column",
        flexShrink: 0,
        height: "100vh",
        position: "sticky",
        top: 0,
        borderRight: "1px solid rgba(0,0,0,0.1)",
      }}
    >
      {/* Header Sidebar nhận diện thương hiệu BIDV */}
      <div
        style={{
          padding: "20px 18px 16px",
          borderBottom: "1px solid rgba(255,255,255,0.08)",
        }}
      >
        <div
          style={{
            fontSize: 12,
            fontWeight: 900,
            letterSpacing: 2,
            color: "#fdb71a",
            marginBottom: 4,
            display: "flex",
            alignItems: "center",
            gap: 6,
          }}
        >
          <span>BIDV</span>
          {/* Biểu tượng hoa mai vàng BIDV */}
          <img
            src={`${import.meta.env.BASE_URL}bidv-flower.png`}
            alt="BIDV"
            style={{ width: 20, height: 20, objectFit: "contain" }}
          />
        </div>
        <div
          style={{
            fontSize: 14,
            fontWeight: 800,
            color: "#ffffff",
            lineHeight: 1.35,
          }}
        >
          Hệ thống Báo cáo
          <br />
          Số liệu CBKH
        </div>
      </div>

      {/* Danh sách 4 menu cơ bản */}
      <nav style={{ flex: 1, padding: "10px 0" }}>
        {NAV.map((item) => {
          const isActive = page === item.id
          return (
            <button
              key={item.id}
              onClick={() => setPage(item.id)}
              aria-current={isActive ? "page" : undefined}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                width: "100%",
                padding: "11px 18px",
                border: "none",
                cursor: "pointer",
                textAlign: "left",
                background: isActive ? "rgba(253,183,26,0.15)" : "transparent",
                borderLeft: isActive
                  ? "3.5px solid #fdb71a"
                  : "3.5px solid transparent",
                color: isActive ? "#fdb71a" : "#b2d6d4",
                fontSize: 13,
                fontWeight: isActive ? 700 : 500,
                transition: "all 0.15s",
                fontFamily: "'Be Vietnam Pro', sans-serif",
              }}
              onMouseEnter={(e) => {
                if (!isActive) {
                  e.currentTarget.style.color = "#ffffff"
                  e.currentTarget.style.background = "rgba(255,255,255,0.06)"
                }
              }}
              onMouseLeave={(e) => {
                if (!isActive) {
                  e.currentTarget.style.color = "#b2d6d4"
                  e.currentTarget.style.background = "transparent"
                }
              }}
            >
              <span style={{ opacity: isActive ? 1 : 0.75 }}>
                {item.icon}
              </span>
              {item.label}
            </button>
          )
        })}
      </nav>

      <div
        style={{
          padding: "12px 18px",
          borderTop: "1px solid rgba(255,255,255,0.08)",
          color: "#80b8b5",
          fontSize: 11,
          fontWeight: 600,
        }}
      >
        Báo cáo HĐV & TD chuẩn
      </div>
    </aside>
  )
}

// ── MAIN APPLICATION ──────────────────────────────────────────────────────────

export default function App() {
  const [page, setPage] = useState<Page>("dashboard")
  const mainContentRef = useRef<HTMLElement>(null)
  const [selectedDonVi, setSelectedDonVi] = useState("Tất cả")
  const [selectedCBKH, setSelectedCBKH] = useState("Tất cả")

  // Load verified authentic initial data directly matching the master Excel/PDF
  const [members, setMembers] = useState<Member[]>(
    initialData.members as Member[],
  )
  const [sourceRows, setSourceRows] = useState<ReportRow[]>(
    initialData.rows as ReportRow[],
  )
  const [sourceName, setSourceName] = useState(
    "Dữ liệu báo cáo chuẩn (SO LIEU.xlsx)",
  )
  const [importError, setImportError] = useState("")
  const [importSuccess, setImportSuccess] = useState("")

  const units = useMemo(
    () => [...new Set(members.map((member) => member.unit))],
    [members],
  )

  const handleExcelUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0]
    event.target.value = ""
    if (!file) return

    try {
      const XLSX = await import("xlsx")
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: "array" })
      const parsed = await parseExcelWorkbook(workbook)

      if (!parsed.members.length) {
        throw new Error(
          "Không thể tìm thấy cấu trúc bảng hợp lệ (cột HDV ĐK, TD ĐK).",
        )
      }

      setMembers(parsed.members)
      setSourceRows(parsed.rows)
      setSelectedDonVi("Tất cả")
      setSelectedCBKH("Tất cả")
      setSourceName(file.name)
      setImportError("")
      setImportSuccess(
        `Đã nạp thành công ${parsed.members.length} CBKH và ${parsed.rows.length} dòng báo cáo từ file ${file.name}!`,
      )
      setTimeout(() => setImportSuccess(""), 6000)
    } catch (error) {
      setImportError(
        error instanceof Error ? error.message : "Không thể đọc file Excel.",
      )
      setTimeout(() => setImportError(""), 6000)
    }
  }

  const availableCBKH = useMemo(() => {
    const source = members
      .map((member, index) => ({ member, index }))
      .filter(
        ({ member }) =>
          selectedDonVi === "Tất cả" || member.unit === selectedDonVi,
      )
    const nameCounts = new Map<string, number>()
    source.forEach(({ member }) => {
      nameCounts.set(member.name, (nameCounts.get(member.name) ?? 0) + 1)
    })
    return source.map(({ member, index }) => ({
      id: getMemberId(member, index),
      label:
        (nameCounts.get(member.name) ?? 0) > 1
          ? `${member.name} — ${member.fullName}${member.code ? ` (${member.code})` : ""}`
          : member.name,
    }))
  }, [members, selectedDonVi])

  const filteredMembers = useMemo(() => {
    return members.filter((m, index) => {
      const unitOk = selectedDonVi === "Tất cả" || m.unit === selectedDonVi
      const cbkhOk =
        selectedCBKH === "Tất cả" || getMemberId(m, index) === selectedCBKH
      return unitOk && cbkhOk
    })
  }, [members, selectedDonVi, selectedCBKH])

  const navigateToPage = (nextPage: Page) => {
    if (nextPage !== page) {
      setPage(nextPage)
    }
  }

  useEffect(() => {
    if (mainContentRef.current) {
      mainContentRef.current.scrollTop = 0
    }
  }, [page])

  return (
    <div style={{ display: "flex", height: "100vh", overflow: "hidden" }}>
      {/* Sidebar nhận diện thương hiệu BIDV */}
      <Sidebar page={page} setPage={navigateToPage} />

      <div
        style={{
          flex: 1,
          display: "flex",
          flexDirection: "column",
          overflow: "hidden",
        }}
      >
        {/* Header (Hidden when printing) */}
        <header
          className="no-print"
          style={{
            background: "#fff",
            borderBottom: "1px solid #e2e8f0",
            padding: "12px 24px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexShrink: 0,
            flexWrap: "wrap",
            gap: 12,
          }}
        >
          <div>
            <div style={{ fontSize: 16, fontWeight: 800, color: "#0f172a" }}>
              {page === "dashboard" && "Tổng quan số liệu"}
              {page === "bieudo" && "Biểu đồ phân tích"}
              {page === "solieu" && "Bảng số liệu chi tiết"}
              {page === "xuatpdf" && "Mẫu xuất báo cáo PDF"}
            </div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 1 }}>
              {filteredMembers.length} CBKH ·{" "}
              {selectedDonVi !== "Tất cả" ? selectedDonVi : "Tất cả đơn vị"} ·{" "}
              {sourceName}
            </div>
          </div>

          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              flexWrap: "wrap",
            }}
          >
            {importSuccess && (
              <span
                style={{
                  color: "#059669",
                  fontSize: 12,
                  fontWeight: 600,
                  background: "#ecfdf5",
                  padding: "4px 10px",
                  borderRadius: 6,
                }}
              >
                ✓ {importSuccess}
              </span>
            )}
            {importError && (
              <span
                style={{
                  color: "#dc2626",
                  fontSize: 12,
                  fontWeight: 600,
                  background: "#fef2f2",
                  padding: "4px 10px",
                  borderRadius: 6,
                }}
              >
                ✕ {importError}
              </span>
            )}

            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                background: "#006b68",
                color: "#fff",
                borderRadius: 7,
                padding: "7px 14px",
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 700,
                boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
                transition: "background 0.15s",
              }}
              onMouseEnter={(e) => (e.currentTarget.style.background = "#005553")}
              onMouseLeave={(e) => (e.currentTarget.style.background = "#006b68")}
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.5"
              >
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="17 8 12 3 7 8" />
                <line x1="12" y1="3" x2="12" y2="15" />
              </svg>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleExcelUpload}
                style={{ display: "none" }}
              />
              Nạp Excel
            </label>

            <FilterBar
              selectedDonVi={selectedDonVi}
              setSelectedDonVi={setSelectedDonVi}
              selectedCBKH={selectedCBKH}
              setSelectedCBKH={setSelectedCBKH}
              availableCBKH={availableCBKH}
              units={units}
            />
          </div>
        </header>

        {/* Main Content Area */}
        <main
          ref={mainContentRef}
          style={{
            flex: 1,
            overflowY: "auto",
            background: page === "xuatpdf" ? "#f8fafc" : "#f1f5f9",
          }}
        >
          <PageErrorBoundary key={page} pageName={page}>
            <div>
              {page === "dashboard" && (
                <DashboardPage members={filteredMembers} />
              )}
              {page === "bieudo" && (
                <BieuDoPage members={filteredMembers} units={units} />
              )}
              {page === "solieu" && (
                <SoLieuPage members={filteredMembers} units={units} />
              )}
              {page === "xuatpdf" && (
                <XuatPDFPage
                  rows={sourceRows}
                  sourceName={sourceName}
                  onUploadExcel={handleExcelUpload}
                />
              )}
            </div>
          </PageErrorBoundary>
        </main>
      </div>
    </div>
  )
}
