import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts"

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
  active?: boolean
  payload?: { name: string; value: number; color: string }[]
  label?: string
}

type BieuDoChartProps = {
  kind: "compare" | "unit" | "conth" | "perday"
  data: ChartDatum[]
  actualFill?: string
  groupByUnit?: boolean
  average?: number
}

function formatNumber(value: number): string {
  return value.toLocaleString("en-US", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })
}

function barColor(value: number, average: number): string {
  if (value > average * 1.5) return "#dc2626"
  if (value > average) return "#d97706"
  return "#059669"
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
  data,
  actualFill,
}: {
  data: DashboardDatum[]
  actualFill: string
}) {
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

export function BieuDoChart({
  kind,
  data,
  actualFill = "#005855",
  groupByUnit = false,
  average = 0,
}: BieuDoChartProps) {
  if (kind === "compare" || kind === "unit") {
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

  const dataKey = kind === "conth" ? "Còn TH" : "Tỷ/ngày"
  const chartHeight = Math.max(320, data.length * 26)
  const formatValue = (value: number) =>
    kind === "conth" ? formatNumber(value) : value.toFixed(3)
  const formatTooltipValue = (value: number): [string, string] =>
    kind === "conth"
      ? [`${formatNumber(value)} tỷ`, "Còn thực hiện"]
      : [`${value.toFixed(4)} tỷ/ngày`, "Cần đạt mỗi ngày"]

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
        <Tooltip formatter={(value) => formatTooltipValue(Number(value) || 0)} />
        <ReferenceLine
          x={average}
          stroke="#f59e0b"
          strokeDasharray="5 3"
          strokeWidth={2}
          label={{
            value:
              kind === "conth"
                ? `TB: ${formatNumber(average)} tỷ`
                : `TB: ${average.toFixed(3)} tỷ/ngày`,
            position: "insideTopRight",
            fontSize: 11,
            fill: "#b45309",
            fontWeight: 700,
          }}
        />
        <Bar dataKey={dataKey} radius={[0, 4, 4, 0]} maxBarSize={20}>
          {data.map((item, index) => (
            <Cell
              key={`${kind}-cell-${index}`}
              fill={barColor(Number(item[dataKey]) || 0, average)}
            />
          ))}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  )
}
