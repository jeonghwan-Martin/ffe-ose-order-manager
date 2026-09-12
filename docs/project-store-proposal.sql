-- REVIEW PROPOSAL ONLY: not a migration and not applied to the live database.
-- Import and validate the backup before switching the application to this store.
-- Access policies are deliberately not granted here. Configure on the designated
-- test/production environment with the approved team access policy at cutover.
begin;
create table public.project_documents (
  project_id uuid primary key references public.projects(id) on delete restrict,
  payload jsonb not null check (jsonb_typeof(payload)='object'),
  version bigint not null default 1 check (version>0),
  updated_at timestamptz not null default now()
);
create table public.project_document_history (
  project_id uuid not null references public.project_documents(project_id) on delete restrict,
  version bigint not null,
  payload jsonb not null,
  archived_at timestamptz not null default now(),
  primary key (project_id,version)
);
alter table public.project_documents enable row level security;
alter table public.project_document_history enable row level security;

create function public.save_project_document(p_project_id uuid,p_payload jsonb,p_expected_version bigint)
returns jsonb language plpgsql security invoker set search_path = '' as $$
declare current_doc public.project_documents%rowtype;
begin
  if jsonb_typeof(p_payload) is distinct from 'object' then
    raise exception '프로젝트 데이터 형식 오류' using errcode='22023';
  end if;
  select * into current_doc from public.project_documents where project_id=p_project_id for update;
  if not found then raise exception '이전되지 않은 프로젝트입니다' using errcode='P0002'; end if;
  if p_expected_version is null or current_doc.version<>p_expected_version then
    raise exception '다른 팀원이 먼저 저장했습니다' using errcode='40001';
  end if;
  insert into public.project_document_history(project_id,version,payload)
    values(p_project_id,current_doc.version,current_doc.payload);
  update public.project_documents set payload=p_payload,version=version+1,updated_at=now()
    where project_id=p_project_id;
  return jsonb_build_object('version',current_doc.version+1,'payload',p_payload);
end;
$$;
revoke all on function public.save_project_document(uuid,jsonb,bigint) from public;
commit;
