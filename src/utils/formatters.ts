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

/** Lấy giá trị % hoàn thành dưới dạng số (để sort, tính toán). */
export const getNumericPct = (m: Metric): number => {
  if (typeof m.pctHT === "number") {
    return Math.abs(m.pctHT) < 50 && m.pctHT !== 0 ? m.pctHT * 100 : m.pctHT
  }
  if (typeof m.pctHT === "string") {
    const parsed = parseFloat(m.pctHT.replace(/%/g, ""))
    return Number.isFinite(parsed) ? parsed : 0
  }
  return m.kh > 0 ? (m.th / m.kh) * 100 : 0
}

/** Chuyển bất kỳ giá trị nào thành number an toàn (không NaN). */
export const toNumber = (value: unknown): number => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0
  const normalized = String(value ?? "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim()
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : 0
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

