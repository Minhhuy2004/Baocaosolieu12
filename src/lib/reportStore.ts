import type { Member, ReportRow } from "../App"

export type ReportSnapshot = {
  id: string;

  source_name: string;

  created_at: string;

  created_by: string | null;

  members: Member[];

  rows: ReportRow[]
}

export type FollowUpNote = {
  id: string;

  member_code: string;

  member_name: string;

  unit: string;

  content: string;

  status: "open" | "done";

  due_date: string | null;

  created_at: string;

  updated_at: string
}

export type AuditEntry = {
  id: number;

  actor_email: string | null;

  action: string;

  entity_type: string;

  entity_id: string | null;

  created_at: string;

  details: Record<string, unknown>
}

async function getClient() {
  const { supabase } = await import("./supabase")

  if (!supabase) {
    throw new Error("Supabase chưa được cấu hình.")
  }

  return supabase
}

export async function saveReportSnapshot(
  sourceName: string,

  members: Member[],

  rows: ReportRow[],
) {
  const client = await getClient()

  const { error } = await client.from("report_snapshots").insert({
    source_name: sourceName,

    members,

    rows,
  })

  if (error) throw error
}

export async function loadReportSnapshots() {
  const client = await getClient()

  const { data, error } = await client

    .from("report_snapshots")

    .select("id, source_name, created_at, created_by, members, rows")

    .order("created_at", { ascending: false })

    .limit(50)

  if (error) throw error

  return data as unknown as ReportSnapshot[]
}

export async function loadFollowUpNotes() {
  const client = await getClient()

  const { data, error } = await client

    .from("follow_up_notes")

    .select(
      "id, member_code, member_name, unit, content, status, due_date, created_at, updated_at",
    )

    .order("created_at", { ascending: false })

    .limit(500)

  if (error) throw error

  return data as unknown as FollowUpNote[]
}

export async function loadAuditEntries() {
  const client = await getClient()

  const { data, error } = await client

    .from("tracking_audit_log")

    .select(
      "id, actor_email, action, entity_type, entity_id, created_at, details",
    )

    .order("created_at", { ascending: false })

    .limit(100)

  if (error) throw error

  return data as unknown as AuditEntry[]
}

export async function createFollowUpNote(note: {
  member_code: string;

  member_name: string;

  unit: string;

  content: string;

  due_date: string | null
}) {
  const client = await getClient()

  const { error } = await client.from("follow_up_notes").insert(note)

  if (error) throw error
}

export async function updateFollowUpNote(
  id: string,

  changes: Pick<FollowUpNote, "status" | "content" | "due_date">,
) {
  const client = await getClient()

  const { error } = await client

    .from("follow_up_notes")

    .update(changes)

    .eq("id", id)

  if (error) throw error
}

export async function deleteFollowUpNote(id: string) {
  const client = await getClient()

  const { error } = await client

    .from("follow_up_notes")

    .delete()

    .eq("id", id)

  if (error) throw error
}
