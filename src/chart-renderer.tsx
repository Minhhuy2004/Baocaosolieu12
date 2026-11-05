import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Rectangle,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts"
import type { ComponentProps } from "react"

type ChartDatum = {
  name: string
  [key: string]: unknown
}

type DashboardDatum = {
  name: string
  fullName: string
  "ĐK": number
  KH: number
  TH: number
}

type ChartTooltipProps = {
  readonly active?: boolean
  readonly payload?: readonly { name: string; value: number; color: string }[]
  readonly label?: string
}

type BieuDoChartProps = {
  readonly kind: "compare" | "unit" | "conth" | "perday"
  readonly data?: ChartDatum[]
  readonly actualFill?: string
  readonly groupByUnit?: boolean
  readonly average?: number
}

function formatNumber(value: unknown): string {
  const num = typeof value === "number" ? value : Number(value)
  if (!Number.isFinite(num)) return "0.0"
  return num.toLocaleString("en-US", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })
}

function barColor(value: unknown, average: unknown): string {
  const v = typeof value === "number" ? value : Number(value)
  const a = typeof average === "number" ? average : Number(average)
  if (!Number.isFinite(v) || !Number.isFinite(a)) return "#059669"
  if (v > a * 1.5) return "#dc2626"
  if (v > a) return "#d97706"
  return "#059669"
}

type ColoredBarShapeProps = ComponentProps<typeof Rectangle> & {
  payload?: ChartDatum
  dataKey: string
  average: number
}

function ColoredBarShape({
  payload,
  dataKey,
  average,
  ...props
}: ColoredBarShapeProps) {
  return (
    <Rectangle {...props} fill={barColor(payload?.[dataKey], average)} />
  )
}

function ChartTooltip({ active, payload, label }: ChartTooltipProps) {
  if (!active || !payload?.length) return null

  return (
    <div
      style={{
        background: "#111827",
        border: "1px solid rgba(255,255,255,0.15)",
        borderRadius: 8,
        padding: "10px 14px",
        boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
      }}
    >
      <div
        style={{
          fontSize: 12,
          color: "#9ca3af",
          marginBottom: 6,
          fontWeight: 700,
        }}
      >
        {label}
      </div>
      {payload.map((item) => (
        <div
          key={item.name}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            fontSize: 12,
            color: "#f9fafb",
            marginBottom: 2,
          }}
        >
          <span
            style={{
              width: 8,
              height: 8,
              borderRadius: "50%",
              background: item.color,
              flexShrink: 0,
            }}
          />
          <span style={{ color: "#9ca3af" }}>{item.name}:</span>
          <span className="mono" style={{ fontWeight: 700 }}>
            {formatNumber(item.value)} tỷ
          </span>
        </div>
      ))}
    </div>
  )
}

export function DashboardChart({
  data = [],
  actualFill = "#005855",
}: {
  readonly data?: DashboardDatum[]
  readonly actualFill?: string
}) {
  if (!data?.length) {
    return (
      <div
        style={{
          height: 280,
          display: "grid",
          placeItems: "center",
          color: "#94a3b8",
          fontSize: 13,
        }}
      >
        Chưa có dữ liệu biểu đồ
      </div>
    )
  }

  return (
    <ResponsiveContainer width="100%" height={280}>
      <BarChart data={data} margin={{ left: 0, right: 8 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
        <XAxis
          dataKey="name"
          tick={{ fontSize: 11, fill: "#64748b" }}
          axisLine={false}
          tickLine={false}
        />
        <YAxis
          tick={{
            fontSize: 11,
            fill: "#64748b",
            fontFamily: "JetBrains Mono",
          }}
          axisLine={false}
          tickLine={false}
          width={48}
        />
        <Tooltip content={<ChartTooltip />} />
        <Legend wrapperStyle={{ fontSize: 12, paddingTop: 6 }} />
        <Bar dataKey="ĐK" fill="#7ecbc9" radius={[3, 3, 0, 0]} />
        <Bar dataKey="KH" fill="#006b68" radius={[3, 3, 0, 0]} />
        <Bar dataKey="TH" fill={actualFill} radius={[3, 3, 0, 0]} />
      </BarChart>
    </ResponsiveContainer>
  )
}

function renderCompareChart(
  kind: "compare" | "unit",
  data: ChartDatum[],
  actualFill: string,
) {
  return (
    <ResponsiveContainer width="100%" height={340}>
      <BarChart
        data={data}
        margin={{
          left: 0,
          right: 16,
          ...(kind === "compare" ? { bottom: 20 } : {}),
        }}
      >
        <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
        <XAxis
          dataKey="name"
          tick={{
            fontSize: kind === "unit" ? 12 : 11,
            fill: "#64748b",
            ...(kind === "unit" ? { fontWeight: 600 } : {}),
          }}
          axisLine={false}
          tickLine={false}
        />
        <YAxis
          tick={{
            fontSize: 11,
            fill: "#64748b",
            fontFamily: "JetBrains Mono",
          }}
          axisLine={false}
          tickLine={false}
          width={50}
        />
        <Tooltip content={<ChartTooltip />} />
        <Legend wrapperStyle={{ fontSize: 12 }} />
        {kind === "compare" ? (
          <>
            <Bar dataKey="ĐK" fill="#7ecbc9" radius={[3, 3, 0, 0]} />
            <Bar dataKey="KH" fill="#006b68" radius={[3, 3, 0, 0]} />
            <Bar dataKey="TH" fill={actualFill} radius={[3, 3, 0, 0]} />
          </>
        ) : (
          <>
            <Bar dataKey="KH" fill="#7ecbc9" radius={[3, 3, 0, 0]} />
            <Bar dataKey="TH" fill={actualFill} radius={[3, 3, 0, 0]} />
          </>
        )}
      </BarChart>
    </ResponsiveContainer>
  )
}

export function BieuDoChart({
  kind,
  data = [],
  actualFill = "#005855",
  groupByUnit = false,
  average = 0,
}: BieuDoChartProps) {
  if (!data?.length) {
    return (
      <div
        style={{
          height: 320,
          display: "grid",
          placeItems: "center",
          color: "#94a3b8",
          fontSize: 13,
        }}
      >
        Chưa có dữ liệu biểu đồ
      </div>
    )
  }

  if (kind === "compare" || kind === "unit") {
    return renderCompareChart(kind, data, actualFill)
  }

  const dataKey = kind === "conth" ? "Còn TH" : "Tỷ/ngày"
  const chartHeight = Math.max(320, data.length * 26)
  const formatValue = (value: unknown) => {
    const num = typeof value === "number" ? value : Number(value)
    if (!Number.isFinite(num)) return "0.0"
    return kind === "conth" ? formatNumber(num) : num.toFixed(3)
  }
  const formatTooltipValue = (value: unknown): [string, string] => {
    const num = typeof value === "number" ? value : Number(value)
    if (!Number.isFinite(num)) {
      return ["0 tỷ", kind === "conth" ? "Còn thực hiện" : "Cần đạt mỗi ngày"]
    }
    return kind === "conth"
      ? [`${formatNumber(num)} tỷ`, "Còn thực hiện"]
      : [`${num.toFixed(4)} tỷ/ngày`, "Cần đạt mỗi ngày"]
  }

  return (
    <ResponsiveContainer width="100%" height={chartHeight}>
      <BarChart
        data={data}
        layout="vertical"
        margin={{
          left: groupByUnit ? 64 : 44,
          right: kind === "conth" ? 64 : 80,
          top: 8,
          bottom: 8,
        }}
      >
        <CartesianGrid
          strokeDasharray="3 3"
          stroke="#f1f5f9"
          horizontal={false}
        />
        <XAxis
          type="number"
          tick={{
            fontSize: 11,
            fill: "#64748b",
            fontFamily: "JetBrains Mono",
          }}
          axisLine={false}
          tickLine={false}
          tickFormatter={formatValue}
        />
        <YAxis
          type="category"
          dataKey="name"
          tick={{
            fontSize: groupByUnit ? 12 : 10,
            fill: "#374151",
            fontWeight: 600,
          }}
          axisLine={false}
          tickLine={false}
          width={groupByUnit ? 62 : 40}
        />
        <Tooltip formatter={(value) => formatTooltipValue(value)} />
        <ReferenceLine
          x={average}
          stroke="#f59e0b"
          strokeDasharray="5 3"
          strokeWidth={2}
          label={{
            value:
              kind === "conth"
                ? `TB: ${formatNumber(average)} tỷ`
                : `TB: ${(Number(average) || 0).toFixed(3)} tỷ/ngày`,
            position: "insideTopRight",
            fontSize: 11,
            fill: "#b45309",
            fontWeight: 700,
          }}
        />
        <Bar
          dataKey={dataKey}
          radius={[0, 4, 4, 0]}
          maxBarSize={20}
          shape={<ColoredBarShape dataKey={dataKey} average={average} />}
        />
      </BarChart>
    </ResponsiveContainer>
  )
}
