create extension if not exists pgcrypto;

create table if not exists public.planning_workspaces (
  id text primary key,
  file_name text not null,
  csv_rows jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.planning_tasks (
  id text primary key,
  workspace_id text not null references public.planning_workspaces(id) on delete cascade,
  row_index integer not null,
  department text not null,
  epic text not null,
  task text not null,
  assignee text not null,
  start_date date not null,
  end_date date not null,
  updated_at timestamptz not null default now()
);

create index if not exists planning_tasks_workspace_idx
  on public.planning_tasks (workspace_id, row_index);

alter table public.planning_workspaces enable row level security;
alter table public.planning_tasks enable row level security;

drop policy if exists "workspace read" on public.planning_workspaces;
create policy "workspace read"
on public.planning_workspaces
for select
to anon, authenticated
using (true);

drop policy if exists "workspace write" on public.planning_workspaces;
create policy "workspace write"
on public.planning_workspaces
for insert
to anon, authenticated
with check (true);

drop policy if exists "workspace update" on public.planning_workspaces;
create policy "workspace update"
on public.planning_workspaces
for update
to anon, authenticated
using (true)
with check (true);

drop policy if exists "task read" on public.planning_tasks;
create policy "task read"
on public.planning_tasks
for select
to anon, authenticated
using (true);

drop policy if exists "task write" on public.planning_tasks;
create policy "task write"
on public.planning_tasks
for insert
to anon, authenticated
with check (true);

drop policy if exists "task update" on public.planning_tasks;
create policy "task update"
on public.planning_tasks
for update
to anon, authenticated
using (true)
with check (true);

drop policy if exists "task delete" on public.planning_tasks;
create policy "task delete"
on public.planning_tasks
for delete
to anon, authenticated
using (true);

alter publication supabase_realtime add table public.planning_tasks;
alter publication supabase_realtime add table public.planning_workspaces;
