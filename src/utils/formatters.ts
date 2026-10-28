// ── FORMATTERS & STYLE HELPERS ────────────────────────────────────────────────
// Pure functions for number formatting, color mapping, and row classification.

import type { Metric } from "../types"

// ── Number formatters ─────────────────────────────────────────────────────────

/** Format số với 1 chữ số thập phân (ví dụ: 1,234.5). */
export const fmt1 = (val: number | null | undefined): string => {
  if (val === null || val === undefined || isNaN(val)) return "0.0"
  return val.toLocaleString("en-US", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })
}

/** Format số với 2 chữ số thập phân (ví dụ: 1,234.56). */
export const fmt2 = (val: number | null | undefined): string => {
  if (val === null || val === undefined || isNaN(val)) return "0.00"
  return val.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })
}

/** Format tỷ lệ % (ví dụ: 95%). Hỗ trợ cả string lẫn number. */
export const fmtPct = (val: number | string | null | undefined): string => {
  if (val === null || val === undefined || val === "") return ""
  let hasPercentSign = false
  if (typeof val === "string") {
    if (val.includes("#")) return val.trim()
    hasPercentSign = val.includes("%")
    const parsed = parseLocalizedNumber(val)
    if (parsed === null) return val
    val = parsed
  }
  const pct =
    !hasPercentSign && Math.abs(val) < 50 && val !== 0
      ? Math.round(val * 100)
      : Math.round(val)
  return `${pct}%`
}

/** Lấy giá trị % hoàn thành dưới dạng số (để sort, tính toán). */
export const getNumericPct = (m: Metric): number => {
  if (typeof m.pctHT === "number") {
    return Math.abs(m.pctHT) < 50 && m.pctHT !== 0 ? m.pctHT * 100 : m.pctHT
  }
  if (typeof m.pctHT === "string") {
    const parsed = parseLocalizedNumber(m.pctHT)
    if (parsed === null) return 0
    return m.pctHT.includes("%")
      ? parsed
      : Math.abs(parsed) < 50 && parsed !== 0
        ? parsed * 100
        : parsed
  }
  return m.kh > 0 ? (m.th / m.kh) * 100 : 0
}

const parseLocalizedNumber = (value: string): number | null => {
  let normalized = String(value ?? "")
    .replace(/[%\s\u00a0]/g, "")
    .trim()
  const isParenthesizedNegative =
    normalized.startsWith("(") && normalized.endsWith(")")
  if (isParenthesizedNegative) normalized = normalized.slice(1, -1)

  const lastComma = normalized.lastIndexOf(",")
  const lastDot = normalized.lastIndexOf(".")
  if (lastComma >= 0 && lastDot >= 0) {
    const decimalSeparator = lastComma > lastDot ? "," : "."
    const groupingSeparator = decimalSeparator === "," ? "." : ","
    normalized = normalized
      .replaceAll(groupingSeparator, "")
      .replace(decimalSeparator, ".")
  } else if (lastComma >= 0) {
    const commaCount = (normalized.match(/,/g) ?? []).length
    const trailingDigits = normalized.length - lastComma - 1
    if (commaCount > 1 || trailingDigits === 3) {
      normalized = normalized.replaceAll(",", "")
    } else {
      normalized = normalized.replace(",", ".")
    }
  }

  const parsed = Number(normalized)
  if (!Number.isFinite(parsed)) return null
  return isParenthesizedNegative ? -Math.abs(parsed) : parsed
}

/** Chuyển bất kỳ giá trị nào thành number an toàn (không NaN). */
export const toNumber = (value: unknown): number => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0
  return parseLocalizedNumber(String(value ?? "")) ?? 0
}

// ── Color helpers ─────────────────────────────────────────────────────────────

/** Màu sắc theo mức % hoàn thành: xanh ≥100%, vàng ≥70%, đỏ <70%. */
export const pctColor = (p: number) =>
  p >= 100 ? "#059669" : p >= 70 ? "#d97706" : "#dc2626"

/** Màu sắc số +/- : xanh nếu dương, đỏ nếu âm. */
export const diffColor = (n: number) => (n >= 0 ? "#059669" : "#dc2626")

// ── Row classification helpers ────────────────────────────────────────────────

/** Kiểm tra dòng là Phòng/PGD (dòng tổng đơn vị). */
export const isUnitRow = (name: string) => /^(P\.|PGD\s)/i.test(name.trim())

/** Kiểm tra dòng là Cấp quản lý (Chị Thảo, Chị Hoa, A.CƯỜNG). */
export const isManagerRow = (name: string) =>
  /^(CHỊ\s|A\.CƯỜNG$)/i.test(name.trim())

/** Kiểm tra dòng Tổng cộng (TC). */
export const isTotalRow = (name: string) => /^TC$/i.test(name.trim())
