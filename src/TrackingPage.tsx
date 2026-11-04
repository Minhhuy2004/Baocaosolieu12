import { useCallback, useEffect, useMemo, useState } from "react"

import type { FormEvent } from "react"

import type { Member, Metric } from "./App"

import { getNumericPct } from "./utils/formatters"

import {
  createFollowUpNote,
  deleteFollowUpNote,
  loadAuditEntries,
  loadFollowUpNotes,
  loadReportSnapshots,
  updateFollowUpNote,
  type AuditEntry,
  type FollowUpNote,
  type ReportSnapshot,
} from "./lib/reportStore"

type TrackingTab = "history" | "alerts" | "notes" | "audit"

type MetricKey = "hdv" | "td"

const TABS: { id: TrackingTab; label: string }[] = [
  { id: "history", label: "Lịch sử & so sánh" },

  { id: "alerts", label: "Cảnh báo tiến độ" },

  { id: "notes", label: "Ghi chú & việc cần làm" },

  { id: "audit", label: "Nhật ký thay đổi" },
]

const panelStyle: React.CSSProperties = {
  background: "#fff",

  border: "1px solid #e2e8f0",

  borderRadius: 12,

  padding: 20,

  boxShadow: "0 1px 3px rgba(15, 23, 42, 0.04)",
}

const buttonStyle: React.CSSProperties = {
  border: 0,

  borderRadius: 7,

  padding: "8px 13px",

  background: "#006b68",

  color: "#fff",

  fontSize: 12,

  fontWeight: 700,

  cursor: "pointer",
}

const inputStyle: React.CSSProperties = {
  border: "1px solid #cbd5e1",

  borderRadius: 7,

  padding: "9px 10px",

  color: "#0f172a",

  font: "inherit",

  fontSize: 13,

  minWidth: 0,
}

const fmt = (value: number) =>
  value.toLocaleString("vi-VN", { maximumFractionDigits: 1 })

const dateTime = (value: string) =>
  new Date(value).toLocaleString("vi-VN", {
    dateStyle: "short",
    timeStyle: "short",
  })

const localDateKey = (date: Date) =>
  `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`

function total(members: Member[], metric: MetricKey, field: keyof Metric) {
  return members.reduce((sum, member) => {
    const value = member[metric][field]

    return sum + (typeof value === "number" ? value : 0)
  }, 0)
}

function totals(members: Member[]) {
  return {
    hdv: total(members, "hdv", "th"),

    hdvPct: members.length
      ? members.reduce((sum, member) => sum + getNumericPct(member.hdv), 0) /
        members.length
      : 0,

    td: total(members, "td", "th"),

    tdPct: members.length
      ? members.reduce((sum, member) => sum + getNumericPct(member.td), 0) /
        members.length
      : 0,
  }
}

function TrackingPage({
  members,

  visibleMembers,

  sourceName,

  currentUserEmail,

  refreshToken,
}: {
  members: Member[];

  visibleMembers: Member[];

  sourceName: string;

  currentUserEmail: string;

  refreshToken: number
}) {
  const [tab, setTab] = useState<TrackingTab>("history")

  const [metric, setMetric] = useState<MetricKey>("hdv")

  const [snapshots, setSnapshots] = useState<ReportSnapshot[]>([])

  const [notes, setNotes] = useState<FollowUpNote[]>([])

  const [auditEntries, setAuditEntries] = useState<AuditEntry[]>([])

  const [loadError, setLoadError] = useState("")

  const [actionError, setActionError] = useState("")

  const [isLoading, setIsLoading] = useState(true)

  const [isSaving, setIsSaving] = useState(false)

  const [selectedSnapshotId, setSelectedSnapshotId] = useState("")

  const [memberId, setMemberId] = useState("")

  const [noteContent, setNoteContent] = useState("")

  const [dueDate, setDueDate] = useState("")

  const reload = useCallback(async () => {
    setIsLoading(true)

    setLoadError("")

    const results = await Promise.allSettled([
      loadReportSnapshots(),

      loadFollowUpNotes(),

      loadAuditEntries(),
    ])

    const failed = results.find(
      (result): result is PromiseRejectedResult => result.status === "rejected",
    )

    if (failed) {
      console.error("Unable to load report tracking data.", failed.reason)

      setLoadError(
        failed.reason instanceof Error
          ? failed.reason.message
          : "Không thể tải dữ liệu theo dõi.",
      )
    } else {
      setSnapshots(
        (results[0] as PromiseFulfilledResult<ReportSnapshot[]>).value,
      )

      setNotes((results[1] as PromiseFulfilledResult<FollowUpNote[]>).value)

      setAuditEntries(
        (results[2] as PromiseFulfilledResult<AuditEntry[]>).value,
      )
    }

    setIsLoading(false)
  }, [])

  useEffect(() => {
    void reload()
  }, [reload, refreshToken])

  useEffect(() => {
    if (
      selectedSnapshotId &&
      !snapshots.some((snapshot) => snapshot.id === selectedSnapshotId)
    ) {
      setSelectedSnapshotId("")
    }
  }, [snapshots, selectedSnapshotId])

  useEffect(() => {
    setMemberId("")
  }, [visibleMembers])

  const selectedMember = useMemo(() => {
    if (!memberId) return undefined

    const index = Number(memberId)

    return Number.isInteger(index) ? visibleMembers[index] : undefined
  }, [memberId, visibleMembers])

  const comparedSnapshot = snapshots.find(
    (snapshot) => snapshot.id === selectedSnapshotId,
  )

  const currentTotals = totals(members)

  const comparedTotals = comparedSnapshot
    ? totals(comparedSnapshot.members)
    : null

  const activeAlerts = useMemo(
    () =>
      visibleMembers

        .flatMap((member) => {
          const pct = getNumericPct(member[metric])

          return pct < 100
            ? [
                {
                  member,
                  pct,
                  gap: Math.max(0, member[metric].kh - member[metric].th),
                },
              ]
            : []
        })

        .sort((a, b) => a.pct - b.pct),

    [visibleMembers, metric],
  )

  const addNote = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()

    if (!selectedMember || !noteContent.trim()) return

    setIsSaving(true)

    setActionError("")

    try {
      await createFollowUpNote({
        member_code: selectedMember.code?.trim() ?? "",

        member_name: selectedMember.fullName || selectedMember.name,

        unit: selectedMember.unit,

        content: noteContent.trim(),

        due_date: dueDate || null,
      })

      setNoteContent("")

      setDueDate("")

      setMemberId("")

      await reload()
    } catch (error) {
      console.error("Unable to create follow-up note.", error)

      setActionError(
        error instanceof Error ? error.message : "Không thể lưu ghi chú.",
      )
    } finally {
      setIsSaving(false)
    }
  }

  const changeNote = async (
    note: FollowUpNote,

    changes: Pick<FollowUpNote, "status" | "content" | "due_date">,
  ) => {
    setActionError("")

    try {
      await updateFollowUpNote(note.id, changes)

      await reload()
    } catch (error) {
      console.error("Unable to update follow-up note.", error)

      setActionError(
        error instanceof Error ? error.message : "Không thể cập nhật ghi chú.",
      )
    }
  }

  const removeNote = async (note: FollowUpNote) => {
    if (!window.confirm(`Xóa ghi chú của ${note.member_name}?`)) return

    setActionError("")

    try {
      await deleteFollowUpNote(note.id)

      await reload()
    } catch (error) {
      console.error("Unable to delete follow-up note.", error)

      setActionError(
        error instanceof Error ? error.message : "Không thể xóa ghi chú.",
      )
    }
  }

  return (
    <div style={{ padding: 24, display: "grid", gap: 18 }}>
      <div
        style={{
          display: "flex",

          justifyContent: "space-between",

          alignItems: "center",

          gap: 12,

          flexWrap: "wrap",
        }}
      >
        <div>
          <h1 style={{ margin: 0, fontSize: 21, color: "#0f172a" }}>
            Theo dõi báo cáo
          </h1>
          <p style={{ margin: "5px 0 0", color: "#64748b", fontSize: 13 }}>
            Nguồn hiện tại: {sourceName} · {currentUserEmail}
          </p>
        </div>
        <button type="button" onClick={() => void reload()} style={buttonStyle}>
          Làm mới dữ liệu
        </button>
      </div>

      <div
        role="tablist"
        aria-label="Chức năng theo dõi"
        style={{ display: "flex", flexWrap: "wrap", gap: 8 }}
      >
        {TABS.map((item) => (
          <button
            type="button"
            role="tab"
            aria-selected={tab === item.id}
            key={item.id}
            onClick={() => setTab(item.id)}
            style={{
              ...buttonStyle,

              background: tab === item.id ? "#006b68" : "#e2e8f0",

              color: tab === item.id ? "#fff" : "#334155",
            }}
          >
            {item.label}
          </button>
        ))}
      </div>

      {loadError && (
        <div
          role="alert"
          style={{
            ...panelStyle,

            borderColor: "#fca5a5",

            background: "#fef2f2",

            color: "#991b1b",

            fontSize: 13,
          }}
        >
          Không thể tải dữ liệu dùng chung: {loadError}. Kiểm tra migration và
          chính sách RLS trong Supabase.
        </div>
      )}
      {actionError && (
        <div role="alert" style={{ color: "#b91c1c", fontSize: 13 }}>
          Thao tác chưa hoàn tất: {actionError}
        </div>
      )}
      {isLoading ? (
        <div role="status" style={{ ...panelStyle, color: "#64748b" }}>
          Đang tải dữ liệu theo dõi...
        </div>
      ) : (
        <>
          {tab === "history" && (
            <section style={{ display: "grid", gap: 16 }}>
              <div style={panelStyle}>
                <h2 style={{ margin: "0 0 5px", fontSize: 16 }}>
                  So sánh với báo cáo đã lưu
                </h2>
                <p
                  style={{ margin: "0 0 14px", fontSize: 12, color: "#64748b" }}
                >
                  Báo cáo hiện tại so với một kỳ đã nạp trước đó.
                </p>
                <select
                  aria-label="Chọn kỳ báo cáo để so sánh"
                  value={selectedSnapshotId}
                  onChange={(event) =>
                    setSelectedSnapshotId(event.target.value)
                  }
                  style={inputStyle}
                >
                  <option value="">Chọn kỳ báo cáo</option>
                  {snapshots.map((snapshot) => (
                    <option key={snapshot.id} value={snapshot.id}>
                      {dateTime(snapshot.created_at)} · {snapshot.source_name}
                    </option>
                  ))}
                </select>
                {comparedSnapshot && comparedTotals && (
                  <div
                    style={{
                      display: "grid",

                      gridTemplateColumns:
                        "repeat(auto-fit, minmax(190px, 1fr))",

                      gap: 10,

                      marginTop: 16,
                    }}
                  >
                    {([
                      ["HĐV thực hiện", currentTotals.hdv, comparedTotals.hdv],

                      [
                        "HĐV hoàn thành",
                        currentTotals.hdvPct,
                        comparedTotals.hdvPct,
                      ],

                      ["TD thực hiện", currentTotals.td, comparedTotals.td],

                      [
                        "TD hoàn thành",
                        currentTotals.tdPct,
                        comparedTotals.tdPct,
                      ],
                    ] as const)

                      .map(([label, current, previous]) => (
                        <div
                          key={label}
                          style={{
                            padding: 12,

                            borderRadius: 8,

                            background: "#f8fafc",
                          }}
                        >
                          <div style={{ color: "#64748b", fontSize: 12 }}>
                            {label}
                          </div>
                          <strong style={{ display: "block", marginTop: 6 }}>
                            {fmt(current)}
                            {label.includes("hoàn thành") ? "%" : " tỷ"}
                          </strong>
                          <span style={{ color: "#475569", fontSize: 12 }}>
                            Kỳ trước {fmt(previous)}
                            {label.includes("hoàn thành") ? "%" : " tỷ"} · Chênh
                            lệch {current - previous >= 0 ? "+" : ""}
                            {fmt(current - previous)}
                            {label.includes("hoàn thành") ? " điểm %" : " tỷ"}
                          </span>
                        </div>
                      ))}
                  </div>
                )}
              </div>

              <div style={{ ...panelStyle, overflowX: "auto" }}>
                <h2 style={{ margin: "0 0 12px", fontSize: 16 }}>
                  Lịch sử nạp Excel ({snapshots.length})
                </h2>
                {snapshots.length === 0 ? (
                  <p style={{ color: "#64748b", fontSize: 13 }}>
                    Chưa có báo cáo được lưu. Sau khi xác nhận nạp Excel, báo
                    cáo sẽ xuất hiện tại đây.
                  </p>
                ) : (
                  <table
                    style={{
                      width: "100%",

                      borderCollapse: "collapse",

                      textAlign: "left",

                      fontSize: 12,
                    }}
                  >
                    <thead>
                      <tr style={{ color: "#64748b" }}>
                        <th style={{ padding: 9 }}>Thời điểm</th>
                        <th style={{ padding: 9 }}>Tệp</th>
                        <th style={{ padding: 9 }}>CBKH</th>
                        <th style={{ padding: 9 }}>Người nạp</th>
                        <th style={{ padding: 9 }}>So sánh</th>
                      </tr>
                    </thead>
                    <tbody>
                      {snapshots.map((snapshot) => (
                        <tr
                          key={snapshot.id}
                          style={{ borderTop: "1px solid #e2e8f0" }}
                        >
                          <td style={{ padding: 9 }}>
                            {dateTime(snapshot.created_at)}
                          </td>
                          <td style={{ padding: 9 }}>{snapshot.source_name}</td>
                          <td style={{ padding: 9 }}>
                            {snapshot.members.length}
                          </td>
                          <td style={{ padding: 9 }}>
                            {auditEntries.find(
                              (entry) =>
                                entry.entity_type === "report" &&
                                entry.entity_id === snapshot.id,
                            )?.actor_email ||
                              snapshot.created_by?.slice(0, 8) ||
                              "—"}
                          </td>
                          <td style={{ padding: 9 }}>
                            <button
                              type="button"
                              onClick={() => setSelectedSnapshotId(snapshot.id)}
                              style={{
                                border: 0,

                                background: "transparent",

                                color: "#006b68",

                                cursor: "pointer",

                                fontWeight: 700,
                              }}
                            >
                              Chọn kỳ
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </section>
          )}

          {tab === "alerts" && (
            <section style={{ display: "grid", gap: 14 }}>
              <div style={panelStyle}>
                <div
                  style={{
                    display: "flex",

                    justifyContent: "space-between",

                    alignItems: "center",

                    gap: 10,

                    flexWrap: "wrap",
                  }}
                >
                  <div>
                    <h2 style={{ margin: 0, fontSize: 16 }}>
                      CBKH chưa đạt 100% kế hoạch
                    </h2>
                    <p
                      style={{
                        margin: "5px 0 0",
                        color: "#64748b",
                        fontSize: 12,
                      }}
                    >
                      Sắp xếp từ tỷ lệ hoàn thành thấp nhất; “còn thiếu” tính
                      theo KH trừ TH.
                    </p>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    {(["hdv", "td"] as const).map((value) => (
                      <button
                        type="button"
                        key={value}
                        onClick={() => setMetric(value)}
                        style={{
                          ...buttonStyle,

                          background: metric === value ? "#006b68" : "#e2e8f0",

                          color: metric === value ? "#fff" : "#334155",
                        }}
                      >
                        {value === "hdv" ? "HĐV" : "Tín dụng"}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <div style={{ ...panelStyle, overflowX: "auto" }}>
                {activeAlerts.length ? (
                  <table
                    style={{
                      width: "100%",

                      borderCollapse: "collapse",

                      textAlign: "left",

                      fontSize: 13,
                    }}
                  >
                    <thead>
                      <tr style={{ color: "#64748b" }}>
                        <th style={{ padding: 9 }}>CBKH</th>
                        <th style={{ padding: 9 }}>Đơn vị</th>
                        <th style={{ padding: 9 }}>% hoàn thành</th>
                        <th style={{ padding: 9 }}>Còn thiếu (tỷ)</th>
                        <th style={{ padding: 9 }}>Mức cảnh báo</th>
                      </tr>
                    </thead>
                    <tbody>
                      {activeAlerts.map(({ member, pct, gap }, index) => (
                        <tr
                          key={`${member.code ?? member.name}-${index}`}
                          style={{ borderTop: "1px solid #e2e8f0" }}
                        >
                          <td style={{ padding: 9 }}>
                            {member.fullName || member.name}
                          </td>
                          <td style={{ padding: 9 }}>{member.unit}</td>
                          <td style={{ padding: 9, fontWeight: 700 }}>
                            {pct.toFixed(1)}%
                          </td>
                          <td style={{ padding: 9 }}>{fmt(gap)}</td>
                          <td style={{ padding: 9 }}>
                            <span
                              style={{
                                color: pct < 70 ? "#b91c1c" : "#a16207",

                                background: pct < 70 ? "#fef2f2" : "#fffbeb",

                                borderRadius: 999,

                                padding: "4px 8px",

                                fontWeight: 700,
                              }}
                            >
                              {pct < 70 ? "Cần chú ý" : "Chưa đạt kế hoạch"}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <p style={{ margin: 0, color: "#047857", fontSize: 13 }}>
                    Không có CBKH nào dưới 100% ở chỉ tiêu này.
                  </p>
                )}
              </div>
            </section>
          )}

          {tab === "notes" && (
            <section style={{ display: "grid", gap: 14 }}>
              <form onSubmit={addNote} style={panelStyle}>
                <h2 style={{ margin: "0 0 12px", fontSize: 16 }}>
                  Thêm ghi chú / việc cần làm
                </h2>
                <div
                  style={{
                    display: "grid",

                    gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",

                    gap: 10,
                  }}
                >
                  <select
                    required
                    aria-label="Chọn CBKH"
                    value={memberId}
                    onChange={(event) => setMemberId(event.target.value)}
                    style={inputStyle}
                  >
                    <option value="">Chọn CBKH</option>
                    {visibleMembers.map((member, index) => (
                      <option key={`${member.code}-${index}`} value={index}>
                        {member.fullName || member.name} · {member.unit}
                      </option>
                    ))}
                  </select>
                  <input
                    required
                    aria-label="Nội dung công việc"
                    placeholder="Nội dung cần theo dõi"
                    value={noteContent}
                    onChange={(event) => setNoteContent(event.target.value)}
                    style={inputStyle}
                  />
                  <input
                    aria-label="Hạn hoàn thành"
                    type="date"
                    value={dueDate}
                    onChange={(event) => setDueDate(event.target.value)}
                    style={inputStyle}
                  />
                  <button disabled={isSaving} type="submit" style={buttonStyle}>
                    {isSaving ? "Đang lưu..." : "Thêm ghi chú"}
                  </button>
                </div>
              </form>
              <div style={{ display: "grid", gap: 10 }}>
                {notes.length === 0 ? (
                  <div
                    style={{ ...panelStyle, color: "#64748b", fontSize: 13 }}
                  >
                    Chưa có ghi chú hoặc việc cần làm.
                  </div>
                ) : (
                  notes.map((note) => {
                    const overdue =
                      note.status === "open" &&
                      note.due_date !== null &&
                      note.due_date < localDateKey(new Date())

                    return (
                      <article key={note.id} style={panelStyle}>
                        <div
                          style={{
                            display: "flex",

                            justifyContent: "space-between",

                            gap: 12,

                            flexWrap: "wrap",
                          }}
                        >
                          <div>
                            <strong style={{ color: "#0f172a", fontSize: 14 }}>
                              {note.member_name} · {note.unit}
                            </strong>
                            <p
                              style={{
                                margin: "6px 0",

                                color: "#334155",

                                fontSize: 13,

                                whiteSpace: "pre-wrap",
                              }}
                            >
                              {note.content}
                            </p>
                            <span
                              style={{
                                color: overdue ? "#b91c1c" : "#64748b",
                                fontSize: 11,
                              }}
                            >
                              {note.due_date
                                ? `${
                                    overdue ? "Quá hạn" : "Hạn"
                                  }: ${note.due_date}`
                                : "Chưa đặt hạn"}
                              {" · Tạo "}
                              {dateTime(note.created_at)}
                            </span>
                          </div>
                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                            }}
                          >
                            <button
                              type="button"
                              onClick={() =>
                                void changeNote(note, {
                                  status:
                                    note.status === "open" ? "done" : "open",

                                  content: note.content,

                                  due_date: note.due_date,
                                })
                              }
                              style={{
                                ...buttonStyle,

                                background:
                                  note.status === "done"
                                    ? "#e2e8f0"
                                    : "#059669",

                                color:
                                  note.status === "done" ? "#334155" : "#fff",
                              }}
                            >
                              {note.status === "done"
                                ? "Đã hoàn thành"
                                : "Đánh dấu xong"}
                            </button>
                            <button
                              type="button"
                              onClick={() => void removeNote(note)}
                              style={{
                                ...buttonStyle,

                                background: "#fff1f2",

                                color: "#b91c1c",
                              }}
                            >
                              Xóa
                            </button>
                          </div>
                        </div>
                      </article>
                    )
                  })
                )}
              </div>
            </section>
          )}

          {tab === "audit" && (
            <section style={{ ...panelStyle, overflowX: "auto" }}>
              <h2 style={{ margin: "0 0 12px", fontSize: 16 }}>
                Nhật ký gần đây (100 mục)
              </h2>
              <p style={{ margin: "0 0 12px", color: "#64748b", fontSize: 12 }}>
                Nhật ký được ghi tự động ở cơ sở dữ liệu khi tạo báo cáo hoặc
                thay đổi ghi chú; tài khoản ứng dụng chỉ có quyền xem.
              </p>
              {auditEntries.length === 0 ? (
                <p style={{ color: "#64748b", fontSize: 13 }}>
                  Chưa có hoạt động được ghi nhận.
                </p>
              ) : (
                <table
                  style={{
                    width: "100%",

                    borderCollapse: "collapse",

                    textAlign: "left",

                    fontSize: 12,
                  }}
                >
                  <thead>
                    <tr style={{ color: "#64748b" }}>
                      <th style={{ padding: 9 }}>Thời điểm</th>
                      <th style={{ padding: 9 }}>Tài khoản</th>
                      <th style={{ padding: 9 }}>Hoạt động</th>
                      <th style={{ padding: 9 }}>Đối tượng</th>
                      <th style={{ padding: 9 }}>Chi tiết</th>
                    </tr>
                  </thead>
                  <tbody>
                    {auditEntries.map((entry) => (
                      <tr
                        key={entry.id}
                        style={{ borderTop: "1px solid #e2e8f0" }}
                      >
                        <td style={{ padding: 9 }}>
                          {dateTime(entry.created_at)}
                        </td>
                        <td style={{ padding: 9 }}>
                          {entry.actor_email || "—"}
                        </td>
                        <td style={{ padding: 9 }}>{entry.action}</td>
                        <td style={{ padding: 9 }}>{entry.entity_type}</td>
                        <td style={{ padding: 9 }}>
                          {Object.entries(entry.details ?? {})

                            .map(([key, value]) => `${key}: ${String(value)}`)

                            .join(" · ")}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </section>
          )}
        </>
      )}
    </div>
  )
}

export default TrackingPage
