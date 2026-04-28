create extension if not exists pgcrypto;

create table if not exists public.planning_collaborators (
  id text primary key,
  username text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now()
);

create table if not exists public.planning_workspaces (
  id text primary key,
  file_name text not null,
  csv_rows jsonb not null,
  owner_name text,
  owner_client text,
  owner_collaborator_id text references public.planning_collaborators(id) on delete set null,
  last_published_by_name text,
  last_published_by_client text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.planning_workspace_members (
  workspace_id text not null references public.planning_workspaces(id) on delete cascade,
  collaborator_id text not null references public.planning_collaborators(id) on delete cascade,
  username_snapshot text not null,
  role text not null default 'collaborator',
  joined_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now(),
  primary key (workspace_id, collaborator_id)
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
  editing_by_name text,
  editing_by_client text,
  editing_started_at timestamptz,
  updated_by_name text,
  updated_by_client text,
  updated_at timestamptz not null default now()
);

alter table public.planning_workspaces
  add column if not exists owner_name text,
  add column if not exists owner_client text,
  add column if not exists owner_collaborator_id text references public.planning_collaborators(id) on delete set null,
  add column if not exists last_published_by_name text,
  add column if not exists last_published_by_client text;

alter table public.planning_collaborators
  add column if not exists username text,
  add column if not exists updated_at timestamptz not null default now(),
  add column if not exists last_seen_at timestamptz not null default now();

alter table public.planning_tasks
  add column if not exists editing_by_name text,
  add column if not exists editing_by_client text,
  add column if not exists editing_started_at timestamptz,
  add column if not exists updated_by_name text,
  add column if not exists updated_by_client text;

alter table public.planning_workspace_members
  add column if not exists username_snapshot text,
  add column if not exists role text not null default 'collaborator',
  add column if not exists joined_at timestamptz not null default now(),
  add column if not exists last_seen_at timestamptz not null default now();

create index if not exists planning_tasks_workspace_idx
  on public.planning_tasks (workspace_id, row_index);

create index if not exists planning_workspace_members_collaborator_idx
  on public.planning_workspace_members (collaborator_id, last_seen_at desc);

alter table public.planning_collaborators enable row level security;
alter table public.planning_workspaces enable row level security;
alter table public.planning_workspace_members enable row level security;
alter table public.planning_tasks enable row level security;

drop policy if exists "collaborator read" on public.planning_collaborators;
create policy "collaborator read"
on public.planning_collaborators
for select
to anon, authenticated
using (true);

drop policy if exists "collaborator write" on public.planning_collaborators;
create policy "collaborator write"
on public.planning_collaborators
for insert
to anon, authenticated
with check (true);

drop policy if exists "collaborator update" on public.planning_collaborators;
create policy "collaborator update"
on public.planning_collaborators
for update
to anon, authenticated
using (true)
with check (true);

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

drop policy if exists "workspace delete" on public.planning_workspaces;
create policy "workspace delete"
on public.planning_workspaces
for delete
to anon, authenticated
using (true);

drop policy if exists "workspace member read" on public.planning_workspace_members;
create policy "workspace member read"
on public.planning_workspace_members
for select
to anon, authenticated
using (true);

drop policy if exists "workspace member write" on public.planning_workspace_members;
create policy "workspace member write"
on public.planning_workspace_members
for insert
to anon, authenticated
with check (true);

drop policy if exists "workspace member update" on public.planning_workspace_members;
create policy "workspace member update"
on public.planning_workspace_members
for update
to anon, authenticated
using (true)
with check (true);

drop policy if exists "workspace member delete" on public.planning_workspace_members;
create policy "workspace member delete"
on public.planning_workspace_members
for delete
to anon, authenticated
using (true);

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

do $$
begin
  alter publication supabase_realtime add table public.planning_tasks;
exception
  when duplicate_object then null;
end
$$;

do $$
begin
  alter publication supabase_realtime add table public.planning_workspaces;
exception
  when duplicate_object then null;
end
$$;

do $$
begin
  alter publication supabase_realtime add table public.planning_workspace_members;
exception
  when duplicate_object then null;
end
$$;
