create table if not exists public.report_snapshots (
  id uuid primary key default gen_random_uuid(),
  source_name text not null,
  members jsonb not null check (jsonb_typeof(members) = 'array'),
  rows jsonb not null default '[]'::jsonb check (jsonb_typeof(rows) = 'array'),
  created_by uuid references auth.users(id) on delete set null default auth.uid(),
  created_at timestamptz not null default now()
);

create index if not exists report_snapshots_created_at_idx
  on public.report_snapshots (created_at desc);

create table if not exists public.follow_up_notes (
  id uuid primary key default gen_random_uuid(),
  member_code text not null default '',
  member_name text not null,
  unit text not null default '',
  content text not null check (length(trim(content)) > 0),
  status text not null default 'open' check (status in ('open', 'done')),
  due_date date,
  created_by uuid references auth.users(id) on delete set null default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists follow_up_notes_created_at_idx
  on public.follow_up_notes (created_at desc);

create table if not exists public.tracking_audit_log (
  id bigint generated always as identity primary key,
  actor_id uuid,
  actor_email text,
  action text not null,
  entity_type text not null,
  entity_id text,
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists tracking_audit_log_created_at_idx
  on public.tracking_audit_log (created_at desc);

alter table public.report_snapshots enable row level security;
alter table public.follow_up_notes enable row level security;
alter table public.tracking_audit_log enable row level security;

drop policy if exists "Authenticated users can read report snapshots"
  on public.report_snapshots;
create policy "Authenticated users can read report snapshots"
  on public.report_snapshots for select to authenticated
  using (true);

drop policy if exists "Authenticated users can create report snapshots"
  on public.report_snapshots;
create policy "Authenticated users can create report snapshots"
  on public.report_snapshots for insert to authenticated
  with check (created_by = (select auth.uid()));

drop policy if exists "Authenticated users can read follow-up notes"
  on public.follow_up_notes;
create policy "Authenticated users can read follow-up notes"
  on public.follow_up_notes for select to authenticated
  using (true);

drop policy if exists "Authenticated users can create follow-up notes"
  on public.follow_up_notes;
create policy "Authenticated users can create follow-up notes"
  on public.follow_up_notes for insert to authenticated
  with check (created_by = (select auth.uid()));

drop policy if exists "Authenticated users can update follow-up notes"
  on public.follow_up_notes;
create policy "Authenticated users can update follow-up notes"
  on public.follow_up_notes for update to authenticated
  using (true)
  with check (true);

drop policy if exists "Authenticated users can delete follow-up notes"
  on public.follow_up_notes;
create policy "Authenticated users can delete follow-up notes"
  on public.follow_up_notes for delete to authenticated
  using (true);

drop policy if exists "Authenticated users can read tracking audit log"
  on public.tracking_audit_log;
create policy "Authenticated users can read tracking audit log"
  on public.tracking_audit_log for select to authenticated
  using (true);

create or replace function public.set_follow_up_note_updated_at()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists follow_up_notes_set_updated_at
  on public.follow_up_notes;
create trigger follow_up_notes_set_updated_at
  before update on public.follow_up_notes
  for each row execute function public.set_follow_up_note_updated_at();

create or replace function public.log_tracking_change()
returns trigger
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  changed_row jsonb;
  action_name text;
  record_id text;
  detail_values jsonb;
begin
  if tg_op = 'DELETE' then
    changed_row = to_jsonb(old);
    record_id = old.id::text;
  else
    changed_row = to_jsonb(new);
    record_id = new.id::text;
  end if;

  if tg_table_name = 'report_snapshots' then
    action_name = 'report_created';
    detail_values = jsonb_build_object(
      'source_name', changed_row -> 'source_name',
      'member_count', jsonb_array_length(changed_row -> 'members')
    );
  else
    action_name = lower('note_' || tg_op);
    detail_values = jsonb_build_object(
      'member_name', changed_row -> 'member_name',
      'unit', changed_row -> 'unit',
      'status', changed_row -> 'status',
      'due_date', changed_row -> 'due_date'
    );
  end if;

  insert into public.tracking_audit_log (
    actor_id,
    actor_email,
    action,
    entity_type,
    entity_id,
    details
  )
  values (
    auth.uid(),
    auth.jwt() ->> 'email',
    action_name,
    case when tg_table_name = 'report_snapshots' then 'report' else 'note' end,
    record_id,
    detail_values
  );

  if tg_op = 'DELETE' then
    return old;
  end if;
  return new;
end;
$$;

revoke all on function public.log_tracking_change() from public, anon, authenticated;

drop trigger if exists report_snapshots_audit_insert
  on public.report_snapshots;
create trigger report_snapshots_audit_insert
  after insert on public.report_snapshots
  for each row execute function public.log_tracking_change();

drop trigger if exists follow_up_notes_audit_change
  on public.follow_up_notes;
create trigger follow_up_notes_audit_change
  after insert or update or delete on public.follow_up_notes
  for each row execute function public.log_tracking_change();

grant select, insert on public.report_snapshots to authenticated;
grant select, insert, update, delete on public.follow_up_notes to authenticated;
grant select on public.tracking_audit_log to authenticated;
