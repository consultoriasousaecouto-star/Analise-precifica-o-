-- Schema do Precificador Reforma — produtos + custos extras.
-- Cole isso no SQL Editor do Supabase (painel do projeto > SQL Editor > New query > Run).

create table if not exists produtos (
  id uuid primary key default gen_random_uuid(),
  sku text not null,
  venda numeric not null,
  custo numeric not null default 0,
  created_at timestamptz not null default now()
);

create table if not exists custos_extras (
  id uuid primary key default gen_random_uuid(),
  produto_id uuid not null references produtos(id) on delete cascade,
  nome text not null,
  valor numeric not null default 0
);

-- RLS ligado, com política aberta: como é uma ferramenta interna sem login,
-- a chave anon (pública) pode ler/gravar livremente. Se algum dia isso for
-- exposto publicamente ou precisar de multiusuário, trocar por políticas
-- restritas a um usuário autenticado.
alter table produtos enable row level security;
alter table custos_extras enable row level security;

create policy "allow all (app interno)" on produtos for all using (true) with check (true);
create policy "allow all (app interno)" on custos_extras for all using (true) with check (true);
