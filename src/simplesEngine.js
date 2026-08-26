// Motor de cálculo — Simples Nacional (Anexo I / Comércio) e Reforma Tributária
// (CBS/IBS, LC 214/2025). Tabelas e fórmulas reaproveitadas (mesmos números,
// já validados) do simplesEngine.js do CBS SIMPLES — aqui simplificado para
// uma única empresa de comércio, com entrada 100% manual (sem XML/NCM/ICMS-ST).

export const SIMPLES_DB = {
  '2027': {
    anexo1: {
      nome: 'Anexo I — Comércio',
      tributos: ['IRPJ', 'CSLL', 'CBS', 'CPP', 'ICMS', 'IBS'],
      faixas: [
        { limite: 180000, nominal: 0.0400, deducao: 0, rep: [0.0550, 0.0350, 0.1533, 0.4150, 0.3400, 0.0017] },
        { limite: 360000, nominal: 0.0730, deducao: 5940, rep: [0.0550, 0.0350, 0.1533, 0.4150, 0.3400, 0.0017] },
        { limite: 720000, nominal: 0.0950, deducao: 13860, rep: [0.0550, 0.0350, 0.1533, 0.4200, 0.3350, 0.0017] },
        { limite: 1800000, nominal: 0.1070, deducao: 22500, rep: [0.0550, 0.0350, 0.1533, 0.4200, 0.3350, 0.0017] },
        { limite: 3600000, nominal: 0.1430, deducao: 87300, rep: [0.0550, 0.0350, 0.1533, 0.4200, 0.3350, 0.0017] },
        { limite: 4800000, nominal: 0.1890, deducao: 378000, rep: [0.1358, 0.1006, 0.3402, 0.4234, 0.0000, 0.0000] },
      ],
    },
  },
  '2029': {
    anexo1: {
      nome: 'Anexo I — Comércio',
      tributos: ['IRPJ', 'CSLL', 'CBS', 'CPP', 'ICMS', 'IBS'],
      faixas: [
        { limite: 180000, nominal: 0.0400, deducao: 0, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.3060, 0.0340] },
        { limite: 360000, nominal: 0.0730, deducao: 5940, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.3060, 0.0340] },
        { limite: 720000, nominal: 0.0950, deducao: 13860, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.3015, 0.0335] },
        { limite: 1800000, nominal: 0.1070, deducao: 22500, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.3015, 0.0335] },
        { limite: 3600000, nominal: 0.1430, deducao: 87300, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.3015, 0.0335] },
        { limite: 4800000, nominal: 0.1900, deducao: 378000, rep: [0.1350, 0.1000, 0.3440, 0.4210, 0.0000, 0.0000] },
      ],
    },
  },
  '2030': {
    anexo1: {
      nome: 'Anexo I — Comércio',
      tributos: ['IRPJ', 'CSLL', 'CBS', 'CPP', 'ICMS', 'IBS'],
      faixas: [
        { limite: 180000, nominal: 0.0400, deducao: 0, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.2720, 0.0680] },
        { limite: 360000, nominal: 0.0730, deducao: 5940, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.2720, 0.0680] },
        { limite: 720000, nominal: 0.0950, deducao: 13860, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2680, 0.0670] },
        { limite: 1800000, nominal: 0.1070, deducao: 22500, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2680, 0.0670] },
        { limite: 3600000, nominal: 0.1430, deducao: 87300, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2680, 0.0670] },
        { limite: 4800000, nominal: 0.1900, deducao: 378000, rep: [0.1350, 0.1000, 0.3440, 0.4210, 0.0000, 0.0000] },
      ],
    },
  },
  '2031': {
    anexo1: {
      nome: 'Anexo I — Comércio',
      tributos: ['IRPJ', 'CSLL', 'CBS', 'CPP', 'ICMS', 'IBS'],
      faixas: [
        { limite: 180000, nominal: 0.0400, deducao: 0, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.2380, 0.1020] },
        { limite: 360000, nominal: 0.0730, deducao: 5940, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.2380, 0.1020] },
        { limite: 720000, nominal: 0.0950, deducao: 13860, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2345, 0.1005] },
        { limite: 1800000, nominal: 0.1070, deducao: 22500, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2345, 0.1005] },
        { limite: 3600000, nominal: 0.1430, deducao: 87300, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2345, 0.1005] },
        { limite: 4800000, nominal: 0.1900, deducao: 378000, rep: [0.1350, 0.1000, 0.3440, 0.4210, 0.0000, 0.0000] },
      ],
    },
  },
  '2032': {
    anexo1: {
      nome: 'Anexo I — Comércio',
      tributos: ['IRPJ', 'CSLL', 'CBS', 'CPP', 'ICMS', 'IBS'],
      faixas: [
        { limite: 180000, nominal: 0.0400, deducao: 0, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.2040, 0.1360] },
        { limite: 360000, nominal: 0.0730, deducao: 5940, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.2040, 0.1360] },
        { limite: 720000, nominal: 0.0950, deducao: 13860, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2010, 0.1340] },
        { limite: 1800000, nominal: 0.1070, deducao: 22500, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2010, 0.1340] },
        { limite: 3600000, nominal: 0.1430, deducao: 87300, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.2010, 0.1340] },
        { limite: 4800000, nominal: 0.1900, deducao: 378000, rep: [0.1350, 0.1000, 0.3440, 0.4210, 0.0000, 0.0000] },
      ],
    },
  },
  '2033': {
    anexo1: {
      nome: 'Anexo I — Comércio',
      tributos: ['IRPJ', 'CSLL', 'CBS', 'CPP', 'IBS'],
      faixas: [
        { limite: 180000, nominal: 0.0400, deducao: 0, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.3400] },
        { limite: 360000, nominal: 0.0730, deducao: 5940, rep: [0.0550, 0.0350, 0.1550, 0.4150, 0.3400] },
        { limite: 720000, nominal: 0.0950, deducao: 13860, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.3350] },
        { limite: 1800000, nominal: 0.1070, deducao: 22500, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.3350] },
        { limite: 3600000, nominal: 0.1430, deducao: 87300, rep: [0.0550, 0.0350, 0.1550, 0.4200, 0.3350] },
        { limite: 4800000, nominal: 0.1900, deducao: 378000, rep: [0.1350, 0.1000, 0.3440, 0.4210, 0.0000] },
      ],
    },
  },
};

// Alíquotas nominais de CBS/IBS por ano — regime regular (fora do Simples),
// mesmos números da LC 214/2025 usados no CBS SIMPLES (constants.js).
export const REFORM_SCHEDULE = {
  '2027': { cbs: 8.8, ibs: 0.1 },
  '2028': { cbs: 8.8, ibs: 0.1 },
  '2029': { cbs: 8.8, ibs: 1.9 },
  '2030': { cbs: 8.8, ibs: 3.7 },
  '2031': { cbs: 8.8, ibs: 5.6 },
  '2032': { cbs: 8.8, ibs: 7.4 },
  '2033': { cbs: 8.8, ibs: 18.5 },
};

// Tabela vigente do Anexo I não muda de estrutura (faixas/nominal/dedução) com a
// reforma — a LC 214/2025 só altera a REPARTIÇÃO interna do DAS (o que passa a
// ser CBS/IBS), não os limites nem as alíquotas nominais. Por isso 2026 (hoje,
// pré-reforma) usa a mesma tabela de faixas de 2027 — é a tabela oficial vigente.
const anoTabelaValido = (ano) => {
  const anos = Object.keys(SIMPLES_DB).sort();
  return SIMPLES_DB[ano] ? ano : anos.filter((a) => a <= ano).pop() || anos[0];
};

const resolveFaixa = (rbt12, ano) => {
  const anoTab = anoTabelaValido(String(ano));
  const tab = SIMPLES_DB[anoTab].anexo1;
  const fi = tab.faixas.findIndex((f) => rbt12 <= f.limite);
  const faixa = tab.faixas[fi >= 0 ? fi : tab.faixas.length - 1];
  const faixaIndex = fi >= 0 ? fi : tab.faixas.length - 1;
  const aliqEf = rbt12 > 0 ? (rbt12 * faixa.nominal - faixa.deducao) / rbt12 : 0;
  return { tab, faixa, faixaIndex, aliqEf, anoTab };
};

// DAS de hoje (2026, pré-reforma): alíquota efetiva cheia da tabela — não existe
// CBS/IBS separado ainda, é um único DAS.
export const calcDasHoje = (rbt12) => {
  const { faixa, faixaIndex, aliqEf } = resolveFaixa(rbt12, 2026);
  return { aliqEf, faixa, faixaIndex };
};

// "Por Dentro" (regime unificado): alíquota efetiva cheia da tabela do ano da
// reforma — CBS e IBS continuam embutidos no DAS, junto com os demais tributos.
// `detalhe`: quebra da alíquota efetiva por tributo (IRPJ, CSLL, CBS, CPP,
// ICMS, IBS), na mesma ordem/nomes da tabela oficial do Simples Nacional.
export const calcDasPorDentroRate = (rbt12, ano) => {
  const { tab, faixa, faixaIndex, aliqEf, anoTab } = resolveFaixa(rbt12, ano);
  const detalhe = tab.tributos.map((nome, i) => ({ nome, rate: aliqEf * (faixa.rep[i] || 0) }));
  return { aliqEf, detalhe, faixa, faixaIndex, anoTab };
};

// "Por Fora" (regime híbrido/regular de IBS-CBS): DAS residual (sem a fatia de
// CBS/IBS) + CBS integral + IBS integral, recolhidos separadamente com as
// alíquotas nominais da reforma (REFORM_SCHEDULE) sobre o valor da venda.
// `cbsPctOverride`: alíquota de CBS em % (ex: 8.8) informada manualmente pelo
// usuário na aba CBS e IBS — substitui a nominal de REFORM_SCHEDULE quando
// preenchida, para simular alíquotas diferentes da LC 214/2025 vigente.
// `detalhe`: quebra do DAS residual por tributo (IRPJ, CSLL, CPP, ICMS — tudo
// que sobra no DAS depois que CBS e IBS saem por completo).
export const calcDasPorForaRate = (rbt12, ano, cbsPctOverride) => {
  const { tab, faixa, faixaIndex, aliqEf, anoTab } = resolveFaixa(rbt12, ano);
  const ibsIdx = tab.tributos.indexOf('IBS');
  const cbsIdx = tab.tributos.indexOf('CBS');
  const detalhe = tab.tributos.reduce((acc, nome, i) => {
    if (i === ibsIdx || i === cbsIdx) return acc;
    acc.push({ nome, rate: aliqEf * (faixa.rep[i] || 0) });
    return acc;
  }, []);
  const rateResidual = detalhe.reduce((acc, d) => acc + d.rate, 0);
  const anos = Object.keys(REFORM_SCHEDULE).sort();
  const anoSched = REFORM_SCHEDULE[String(ano)] ? String(ano) : anos.filter((a) => a <= String(ano)).pop() || anos[0];
  const sched = REFORM_SCHEDULE[anoSched];
  const cbsRate = (Number.isFinite(cbsPctOverride) ? cbsPctOverride : sched.cbs) / 100;
  const ibsRate = sched.ibs / 100;
  return { rateResidual, detalhe, cbsRate, ibsRate, totalRate: rateResidual + cbsRate + ibsRate, faixa, faixaIndex, anoTab };
};

export const fmtPct = (v, casas = 2) => `${(v * 100).toFixed(casas)}%`;
export const fmtR = (v) =>
  (v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
