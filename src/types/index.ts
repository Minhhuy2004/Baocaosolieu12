// ── SHARED TYPES ──────────────────────────────────────────────────────────────
// All TypeScript interfaces and type aliases used across the application.

/** Dữ liệu chỉ tiêu HĐV hoặc TD của một CBKH. */
export type Metric = {
  dk: number;
  kh: number;
  giaoTang: number;
  th: number;
  conTh: number;
  soDk: number;
  soDauThang: number;
  soTuanTruoc: number;
  pctHT: number | string | null;
}

/** Thông tin một Cán bộ Khách hàng (CBKH). */
export type Member = {
  code?: string;
  name: string;
  fullName: string;
  unit: string;
  hdv: Metric;
  td: Metric;
}

/** Một dòng trong bảng báo cáo tổng hợp (bao gồm cả dòng tổng và dòng tiêu đề). */
export type ReportRow = {
  code?: string;
  name: string;
  fullName: string;
  unit: string;
  hdv: Metric;
  td: Metric;
  kind: "member" | "unit" | "manager" | "total";
}

/** Các trang điều hướng trong ứng dụng. */
export type Page = "dashboard" | "bieudo" | "solieu" | "xuatpdf"

