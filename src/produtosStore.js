// Persistência dos produtos: usa o Supabase quando as chaves estão
// configuradas (.env — VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY); sem
// isso, cai automaticamente para o localStorage do navegador (mesmo
// comportamento de antes) — o app nunca fica sem salvar nada.
import { supabase, supabaseConfigurado } from './supabaseClient';

const LS_KEY = 'precificador-reforma:produtos';

export const gerarId = () =>
  (typeof crypto !== 'undefined' && crypto.randomUUID)
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;

// custo de compra + soma de todos os custos extras adicionados no produto.
export const custoTotal = (p) => (p.custo || 0) + (p.custosExtras || []).reduce((acc, c) => acc + (c.valor || 0), 0);

const carregarProdutosLocal = () => {
  try {
    const raw = localStorage.getItem(LS_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    return Array.isArray(parsed) && parsed.length ? parsed : null;
  } catch {
    return null;
  }
};

export const salvarSnapshotLocal = (produtos) => {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify(produtos));
  } catch {
    // localStorage indisponível (modo privado, quota cheia) — segue sem salvar.
  }
};

// Ao iniciar: Supabase primeiro (fonte de verdade, se configurado); sem
// linhas lá ainda (banco novo) ou sem Supabase configurado, cai pro
// localStorage; sem nada em nenhum dos dois, devolve null (chamador semeia
// a partir da planilha-base).
export const carregarProdutos = async () => {
  if (supabaseConfigurado) {
    const { data, error } = await supabase
      .from('produtos')
      .select('id, sku, venda, custo, custos_extras(id, nome, valor)')
      .order('created_at', { ascending: true });
    if (!error && data && data.length) {
      return data.map((p) => ({
        id: p.id, sku: p.sku, venda: p.venda, custo: p.custo,
        custosExtras: (p.custos_extras || []).map((c) => ({ id: c.id, nome: c.nome, valor: c.valor })),
      }));
    }
  }
  return carregarProdutosLocal();
};

// Semeia o banco (Supabase, se configurado) com a lista inicial — usado só
// na primeira carga, quando não há nada salvo em lugar nenhum ainda.
export const semearProdutos = async (produtos) => {
  if (!supabaseConfigurado) return produtos;
  const { data, error } = await supabase
    .from('produtos')
    .insert(produtos.map((p) => ({ sku: p.sku, venda: p.venda, custo: p.custo })))
    .select('id, sku, venda, custo');
  if (error || !data) return produtos;
  return data.map((row) => ({ ...row, custosExtras: [] }));
};

// Carga inicial memoizada num promise em nível de módulo — o React
// StrictMode (dev) roda o efeito de montagem duas vezes de propósito; sem
// isso, as duas chamadas assíncronas corririam em paralelo e, achando o
// banco vazio nas duas, semeariam a planilha-base duas vezes (produtos
// duplicados). O módulo só existe uma vez por carregamento de página, então
// esse cache garante uma única execução mesmo com o duplo-efeito do
// StrictMode ou cliques repetidos.
let cargaInicialPromise = null;
export const carregarOuSemear = (buscarSeed) => {
  if (!cargaInicialPromise) {
    cargaInicialPromise = (async () => {
      const salvos = await carregarProdutos();
      if (salvos) return salvos;
      const seed = await buscarSeed();
      return semearProdutos(seed);
    })();
  }
  return cargaInicialPromise;
};

export const criarProduto = async ({ sku, custo, venda }) => {
  if (supabaseConfigurado) {
    const { data, error } = await supabase.from('produtos').insert({ sku, venda, custo }).select().single();
    if (!error && data) return { id: data.id, sku: data.sku, venda: data.venda, custo: data.custo, custosExtras: [] };
  }
  return { id: gerarId(), sku, venda, custo, custosExtras: [] };
};

export const excluirProduto = async (id) => {
  if (supabaseConfigurado) await supabase.from('produtos').delete().eq('id', id);
};

export const atualizarCustoBase = async (id, custo) => {
  if (supabaseConfigurado) await supabase.from('produtos').update({ custo }).eq('id', id);
};

export const criarCustoExtra = async (produtoId, nome, valor) => {
  if (supabaseConfigurado) {
    const { data, error } = await supabase
      .from('custos_extras').insert({ produto_id: produtoId, nome, valor }).select().single();
    if (!error && data) return { id: data.id, nome: data.nome, valor: data.valor };
  }
  return { id: gerarId(), nome, valor };
};

export const excluirCustoExtra = async (custoId) => {
  if (supabaseConfigurado) await supabase.from('custos_extras').delete().eq('id', custoId);
};

// Importar planilha substitui a lista inteira — apaga tudo e insere de novo.
export const substituirTodosProdutos = async (produtos) => {
  if (!supabaseConfigurado) return produtos.map((p) => ({ ...p, id: gerarId(), custosExtras: [] }));
  await supabase.from('produtos').delete().not('id', 'is', null);
  const { data, error } = await supabase
    .from('produtos')
    .insert(produtos.map((p) => ({ sku: p.sku, venda: p.venda, custo: p.custo })))
    .select('id, sku, venda, custo');
  if (error || !data) return produtos.map((p) => ({ ...p, id: gerarId(), custosExtras: [] }));
  return data.map((row) => ({ ...row, custosExtras: [] }));
};
