import { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, TrendingUp, Calculator, FileSpreadsheet, LineChart } from 'lucide-react';
import {
  calcDasHoje,
  calcDasPorDentroRate,
  calcDasPorForaRate,
  REFORM_SCHEDULE,
  fmtPct,
  fmtR,
} from './simplesEngine';
import produtosBaseUrl from './assets/produtos-base.xlsx?url';

const ANOS_REFORMA = Object.keys(REFORM_SCHEDULE);

// ─── Inputs numéricos em BRL ────────────────────────────────────────────────
// Aceita tanto "8,8" (vírgula BR) quanto "8.8" (ponto, comum ao digitar
// alíquotas) e "1.234,56" (milhar BR) — vírgula, quando presente, é sempre o
// separador decimal; sem vírgula, o ponto é tratado como decimal.
const parseBR = (v) => {
  const cleaned = String(v ?? '').replace(/[^\d,.-]/g, '');
  if (!cleaned) return 0;
  const normalized = cleaned.includes(',') ? cleaned.replace(/\./g, '').replace(',', '.') : cleaned;
  const n = parseFloat(normalized);
  return Number.isFinite(n) ? n : 0;
};

function CurrencyInput({ label, value, onChange, placeholder }) {
  return (
    <label className="block">
      <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">{label}</span>
      <input
        type="text"
        inputMode="decimal"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
      />
    </label>
  );
}

// ─── Parser da planilha de produtos ────────────────────────────────────────
function parseProdutos(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!rows.length) return [];
  const header = rows[0].map((h) => String(h || '').toLowerCase());
  const findCol = (needle, fallbackIdx) => {
    const idx = header.findIndex((h) => h.includes(needle));
    return idx >= 0 ? idx : fallbackIdx;
  };
  const skuCol = findCol(header.some((h) => h.includes('produto')) ? 'produto' : 'sku', 0);
  const vendaCol = findCol('venda', 1);
  const custoCol = findCol('custo', 7);

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[skuCol]) continue;
    const venda = parseFloat(r[vendaCol]);
    const custo = parseFloat(r[custoCol]);
    if (!Number.isFinite(venda) || !Number.isFinite(custo) || venda <= 0) continue;
    out.push({ sku: String(r[skuCol]), venda, custo });
  }
  return out;
}

// ─── Gráfico: margem com repasse x margem se o preço não mudar ─────────────
// Duas séries com uma cor fixa cada (identidade, não ranking): "com repasse"
// (verde/aqua — alvo preservado) e "sem repasse" (laranja — erosão se o preço
// não acompanhar a alíquota real). Par validado (ΔE adjacente ≥8 CVD / ≥15
// visão normal) na paleta de referência do skill de dataviz.
const COR_COM_REPASSE = '#1baf7a';
const COR_SEM_REPASSE = '#eb6834';

function niceTicks(min, max, count = 4) {
  const span = max - min || 0.01;
  const step = span / count;
  return Array.from({ length: count + 1 }, (_, i) => min + step * i);
}

function MargemChart({ pontos }) {
  const svgRef = useRef(null);
  const [hoverIdx, setHoverIdx] = useState(null);
  const W = 760, H = 260, M = { top: 16, right: 16, bottom: 28, left: 46 };
  const innerW = W - M.left - M.right, innerH = H - M.top - M.bottom;
  const meses = pontos.length - 1;

  const vals = pontos.flatMap((p) => [p.margemComRepasse, p.margemSemRepasse]);
  const rawMin = Math.min(...vals), rawMax = Math.max(...vals);
  const pad = Math.max((rawMax - rawMin) * 0.15, 0.01);
  const yMin = rawMin - pad, yMax = rawMax + pad;

  const x = (m) => M.left + (meses > 0 ? (m / meses) * innerW : 0);
  const y = (v) => M.top + innerH - ((v - yMin) / (yMax - yMin)) * innerH;

  const pathFor = (key) => pontos.map((p, i) => `${i === 0 ? 'M' : 'L'} ${x(p.mes)} ${y(p[key])}`).join(' ');
  const areaGap = [
    ...pontos.map((p) => `${p.mes === 0 ? 'M' : 'L'} ${x(p.mes)} ${y(p.margemComRepasse)}`),
    ...[...pontos].reverse().map((p) => `L ${x(p.mes)} ${y(p.margemSemRepasse)}`),
    'Z',
  ].join(' ');

  const ticks = niceTicks(yMin, yMax, 4);
  const hovered = hoverIdx != null ? pontos[hoverIdx] : null;

  const handleMove = (e) => {
    const rect = svgRef.current.getBoundingClientRect();
    const px = ((e.clientX - rect.left) / rect.width) * W;
    const m = Math.round(((px - M.left) / innerW) * meses);
    setHoverIdx(Math.min(meses, Math.max(0, m)));
  };

  return (
    <div className="relative">
      <svg
        ref={svgRef} viewBox={`0 0 ${W} ${H}`} className="w-full h-auto" data-chart="margem"
        onMouseMove={handleMove} onMouseLeave={() => setHoverIdx(null)}
      >
        {ticks.map((t, i) => (
          <g key={i}>
            <line x1={M.left} x2={W - M.right} y1={y(t)} y2={y(t)} stroke="#e1e0d9" strokeWidth="1" />
            <text x={M.left - 8} y={y(t)} textAnchor="end" dominantBaseline="middle" fontSize="10" fill="#898781">
              {(t * 100).toFixed(0)}%
            </text>
          </g>
        ))}
        <line x1={M.left} x2={W - M.right} y1={M.top + innerH} y2={M.top + innerH} stroke="#c3c2b7" strokeWidth="1" />
        {pontos.map((p) => (
          <text key={p.mes} x={x(p.mes)} y={H - 6} textAnchor="middle" fontSize="10" fill="#898781">
            {p.mes === 0 ? 'hoje' : `m${p.mes}`}
          </text>
        ))}

        <path d={areaGap} fill={COR_SEM_REPASSE} opacity="0.10" />
        <path d={pathFor('margemSemRepasse')} fill="none" stroke={COR_SEM_REPASSE} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <path d={pathFor('margemComRepasse')} fill="none" stroke={COR_COM_REPASSE} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />

        <circle cx={x(meses)} cy={y(pontos[meses].margemComRepasse)} r="4" fill={COR_COM_REPASSE} stroke="#fcfcfb" strokeWidth="2" />
        <circle cx={x(meses)} cy={y(pontos[meses].margemSemRepasse)} r="4" fill={COR_SEM_REPASSE} stroke="#fcfcfb" strokeWidth="2" />
        <text x={x(meses) - 6} y={y(pontos[meses].margemComRepasse) - 8} textAnchor="end" fontSize="11" fontWeight="700" fill="#0b0b0b">
          {fmtPct(pontos[meses].margemComRepasse)}
        </text>
        <text x={x(meses) - 6} y={y(pontos[meses].margemSemRepasse) + 16} textAnchor="end" fontSize="11" fontWeight="700" fill="#0b0b0b">
          {fmtPct(pontos[meses].margemSemRepasse)}
        </text>

        {hovered && (
          <>
            <line x1={x(hovered.mes)} x2={x(hovered.mes)} y1={M.top} y2={M.top + innerH} stroke="#c3c2b7" strokeWidth="1" />
            <circle cx={x(hovered.mes)} cy={y(hovered.margemComRepasse)} r="4" fill={COR_COM_REPASSE} stroke="#fcfcfb" strokeWidth="2" />
            <circle cx={x(hovered.mes)} cy={y(hovered.margemSemRepasse)} r="4" fill={COR_SEM_REPASSE} stroke="#fcfcfb" strokeWidth="2" />
          </>
        )}
      </svg>

      {hovered && (
        <div
          className="absolute top-2 bg-white border border-slate-200 rounded-lg shadow-md px-3 py-2 text-xs pointer-events-none"
          style={{ left: `${Math.min(82, Math.max(2, (x(hovered.mes) / W) * 100))}%` }}
        >
          <p className="font-semibold text-slate-700 mb-1">{hovered.mes === 0 ? 'Hoje' : `Mês ${hovered.mes}`}</p>
          <p className="flex items-center gap-1.5"><span className="inline-block w-2.5 h-0.5 rounded" style={{ background: COR_COM_REPASSE }} />Com repasse: <b>{fmtPct(hovered.margemComRepasse)}</b></p>
          <p className="flex items-center gap-1.5"><span className="inline-block w-2.5 h-0.5 rounded" style={{ background: COR_SEM_REPASSE }} />Sem repasse: <b>{fmtPct(hovered.margemSemRepasse)}</b></p>
          <p className="text-slate-400 mt-1">alíquota {fmtPct(hovered.aliquotaMes, 2)}</p>
        </div>
      )}

      <div className="flex items-center gap-5 mt-2 text-xs text-slate-600">
        <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-0.5 rounded" style={{ background: COR_COM_REPASSE }} />Margem com repasse (alvo preservado)</span>
        <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-0.5 rounded" style={{ background: COR_SEM_REPASSE }} />Margem se o preço não mudar</span>
      </div>
    </div>
  );
}

// ─── Aba 4: Repasse (visual) ────────────────────────────────────────────────
// A alíquota REAL (pelo RBT12 informado) costuma vir maior que a alíquota que
// os preços atuais já embutem (o negócio cresceu de faixa e o preço ainda não
// acompanhou). Em vez de repassar tudo de uma vez, distribui o aumento em
// etapas mensais, preservando a margem que os preços de hoje já consideram —
// não a margem real de hoje, que já está mais apertada do que parece.
function RepasseTab({ rbt12Raw, produtos, aliquotaUsadaRaw, setAliquotaUsadaRaw, mesesRaw, setMesesRaw }) {
  const rbt12 = parseBR(rbt12Raw);
  const aliquotaReal = useMemo(() => (rbt12 > 0 ? calcDasHoje(rbt12).aliqEf : 0), [rbt12]);
  const aliquotaUsada = parseBR(aliquotaUsadaRaw) / 100;
  const meses = Math.max(1, parseInt(mesesRaw, 10) || 0);
  const gap = aliquotaReal - aliquotaUsada;

  const produtosValidos = useMemo(() => produtos.filter((p) => p.venda > 0), [produtos]);
  const margemAlvoMedia = useMemo(() => {
    if (!produtosValidos.length) return 0;
    const soma = produtosValidos.reduce((acc, p) => acc + (1 - p.custo / p.venda - aliquotaUsada), 0);
    return soma / produtosValidos.length;
  }, [produtosValidos, aliquotaUsada]);

  const pontos = useMemo(() => Array.from({ length: meses + 1 }, (_, mes) => {
    const aliquotaMes = aliquotaUsada + gap * (mes / meses);
    return { mes, aliquotaMes, margemComRepasse: margemAlvoMedia, margemSemRepasse: margemAlvoMedia - gap * (mes / meses) };
  }), [aliquotaUsada, gap, meses, margemAlvoMedia]);

  const [mesSelecionado, setMesSelecionado] = useState(meses);
  useEffect(() => { setMesSelecionado(meses); }, [meses]);
  const aliquotaSelecionada = aliquotaUsada + gap * (mesSelecionado / meses);

  const linhasMes = useMemo(() => produtosValidos.map((p) => {
    const margemAlvo = 1 - p.custo / p.venda - aliquotaUsada;
    const denom = 1 - margemAlvo - aliquotaSelecionada;
    const precoSugerido = denom > 0 ? p.custo / denom : null;
    const margemSemRepasse = 1 - p.custo / p.venda - aliquotaSelecionada;
    return { ...p, margemAlvo, precoSugerido, margemSemRepasse };
  }), [produtosValidos, aliquotaUsada, aliquotaSelecionada]);

  const temDados = !!aliquotaUsadaRaw && rbt12 > 0;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-1">Repasse suave até a alíquota real</h3>
        <p className="text-[11px] text-slate-400 mb-4">
          Se os preços de hoje ainda consideram uma alíquota menor do que a real (a empresa cresceu de faixa),
          em vez de repassar tudo de uma vez, distribui o aumento em etapas mensais até a alíquota real.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <CurrencyInput label="Alíquota já embutida nos preços hoje (%)" value={aliquotaUsadaRaw} onChange={setAliquotaUsadaRaw} placeholder="Ex: 6,00" />
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Meses até o repasse completo</span>
            <input
              type="number" min="1" value={mesesRaw} onChange={(e) => setMesesRaw(e.target.value)}
              className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
            />
          </label>
        </div>
        {rbt12 <= 0 && <p className="text-sm text-slate-400 italic mt-3">Informe o RBT12 na aba Simples Nacional primeiro.</p>}
      </div>

      {temDados && (
        <>
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <div className="flex flex-wrap items-baseline gap-x-6 gap-y-1 mb-4">
              <p className="text-sm text-slate-600">Hoje nos preços: <b>{fmtPct(aliquotaUsada)}</b></p>
              <p className="text-sm text-slate-600">Real (RBT12 informado): <b>{fmtPct(aliquotaReal, 4)}</b></p>
              <p className="text-sm text-slate-600">Gap: <b className={gap > 0 ? 'text-amber-600' : 'text-slate-500'}>{fmtPct(gap, 4)}</b></p>
              <p className="text-sm text-slate-600">Margem-alvo média: <b>{fmtPct(margemAlvoMedia)}</b></p>
            </div>
            {gap <= 0 ? (
              <p className="text-sm text-slate-400 italic">Sem gap a repassar — a alíquota real já é igual ou menor que a usada nos preços.</p>
            ) : produtosValidos.length === 0 ? (
              <p className="text-sm text-slate-400 italic">Carregue produtos na aba Precificação para ver a margem projetada.</p>
            ) : (
              <MargemChart pontos={pontos} />
            )}
          </div>

          {gap > 0 && produtosValidos.length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
              <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Escolha o mês</h4>
              <div className="flex flex-wrap gap-2 mb-5">
                {pontos.map((p) => (
                  <button
                    key={p.mes}
                    onClick={() => setMesSelecionado(p.mes)}
                    className={`px-3 py-1.5 rounded-lg text-sm font-semibold border transition ${
                      mesSelecionado === p.mes
                        ? 'bg-[#2F6F4E] text-white border-[#2F6F4E]'
                        : 'bg-white text-slate-600 border-slate-200 hover:border-[#2F6F4E]'
                    }`}
                  >
                    {p.mes === 0 ? 'Hoje' : `Mês ${p.mes}`}
                  </button>
                ))}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-5">
                <div className="bg-slate-50 rounded-xl p-4">
                  <p className="text-xs font-semibold text-slate-500 uppercase">Alíquota do mês</p>
                  <p className="text-2xl font-black text-slate-800 mt-1">{fmtPct(aliquotaSelecionada, 4)}</p>
                </div>
                <div className="bg-emerald-50 rounded-xl p-4">
                  <p className="text-xs font-semibold text-emerald-700 uppercase">Margem atual — com repasse</p>
                  <p className="text-2xl font-black text-emerald-700 mt-1">{fmtPct(margemAlvoMedia)}</p>
                  <p className="text-[11px] text-emerald-600 mt-1">preservada, seguindo o cronograma</p>
                </div>
                <div className="bg-amber-50 rounded-xl p-4">
                  <p className="text-xs font-semibold text-amber-700 uppercase">Margem atual — sem repasse</p>
                  <p className="text-2xl font-black text-amber-700 mt-1">{fmtPct(pontos[mesSelecionado].margemSemRepasse)}</p>
                  <p className="text-[11px] text-amber-600 mt-1">se o preço não mudar até esse mês</p>
                </div>
              </div>

              <h4 className="text-xs font-bold text-slate-500 uppercase mb-2">Detalhe por produto — {mesSelecionado === 0 ? 'hoje' : `mês ${mesSelecionado}`}</h4>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-slate-50 text-left text-[11px] uppercase text-slate-500">
                      <th className="px-4 py-2">SKU</th>
                      <th className="px-4 py-2 text-right">Preço atual</th>
                      <th className="px-4 py-2 text-right">Preço sugerido no mês</th>
                      <th className="px-4 py-2 text-right">Margem com repasse</th>
                      <th className="px-4 py-2 text-right">Margem se não mudar o preço</th>
                    </tr>
                  </thead>
                  <tbody>
                    {linhasMes.map((l, i) => (
                      <tr key={l.sku + i} className="border-t border-slate-100">
                        <td className="px-4 py-2 font-mono text-xs text-slate-600">{l.sku}</td>
                        <td className="px-4 py-2 text-right">{fmtR(l.venda)}</td>
                        <td className="px-4 py-2 text-right font-semibold text-emerald-700">{l.precoSugerido != null ? fmtR(l.precoSugerido) : '—'}</td>
                        <td className="px-4 py-2 text-right text-emerald-700">{fmtPct(l.margemAlvo)}</td>
                        <td className={`px-4 py-2 text-right ${l.margemSemRepasse < l.margemAlvo ? 'text-amber-600' : 'text-slate-600'}`}>{fmtPct(l.margemSemRepasse)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ─── Aba 1: Simples Nacional (hoje, 2026) ──────────────────────────────────
function SimplesNacionalTab({ rbt12Raw, setRbt12Raw, faturamentoMesRaw, setFaturamentoMesRaw }) {
  const rbt12 = parseBR(rbt12Raw);
  const fatMes = parseBR(faturamentoMesRaw);
  const resultado = useMemo(() => (rbt12 > 0 ? calcDasHoje(rbt12) : null), [rbt12]);

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-4">Dados da empresa (Anexo I — Comércio)</h3>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <CurrencyInput label="RBT12 (faturamento últimos 12 meses)" value={rbt12Raw} onChange={setRbt12Raw} placeholder="Ex: 480000,00" />
          <CurrencyInput label="Faturamento do mês" value={faturamentoMesRaw} onChange={setFaturamentoMesRaw} placeholder="Ex: 40000,00" />
        </div>
      </div>

      {resultado && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <p className="text-xs font-semibold text-slate-500 uppercase">Faixa</p>
            <p className="text-2xl font-black text-slate-800 mt-1">{resultado.faixaIndex + 1}ª faixa</p>
            <p className="text-[11px] text-slate-400 mt-1">até {fmtR(resultado.faixa.limite)}</p>
          </div>
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <p className="text-xs font-semibold text-slate-500 uppercase">Alíquota nominal</p>
            <p className="text-2xl font-black text-slate-800 mt-1">{fmtPct(resultado.faixa.nominal)}</p>
            <p className="text-[11px] text-slate-400 mt-1">dedução {fmtR(resultado.faixa.deducao)}</p>
          </div>
          <div className="bg-[#2F6F4E] text-white rounded-2xl border border-[#2F6F4E] p-5 shadow-sm">
            <p className="text-xs font-semibold text-emerald-100 uppercase">Alíquota efetiva hoje (2026)</p>
            <p className="text-2xl font-black mt-1">{fmtPct(resultado.aliqEf, 4)}</p>
            <p className="text-[11px] text-emerald-100 mt-1">DAS do mês: {fmtR(fatMes * resultado.aliqEf)}</p>
          </div>
        </div>
      )}

      {resultado && (
        <p className="text-[11px] text-slate-400 italic">
          * As faixas, alíquotas nominais e deduções do Anexo I não mudam com a reforma — a LC 214/2025 só altera
          a composição interna do DAS a partir de 2027 (o que passa a ser CBS/IBS). Por isso esse cálculo vale
          tanto para hoje quanto de base para a simulação da aba "CBS e IBS".
        </p>
      )}

      {!resultado && (
        <p className="text-sm text-slate-400 italic">Informe o RBT12 para calcular a faixa e a alíquota efetiva do DAS.</p>
      )}
    </div>
  );
}

// ─── Aba 2: CBS e IBS (Reforma, a partir de 2027) ──────────────────────────
function CbsIbsTab({
  rbt12Raw, faturamentoMesRaw, ano, setAno, regime, setRegime, valorVendaRaw, setValorVendaRaw,
  cbsPctRaw, setCbsPctRaw, creditoCbsRaw, setCreditoCbsRaw, creditoIbsRaw, setCreditoIbsRaw,
}) {
  const rbt12 = parseBR(rbt12Raw);
  const fatMes = parseBR(faturamentoMesRaw);
  const valorVenda = parseBR(valorVendaRaw);
  const sched = REFORM_SCHEDULE[ano];
  const cbsPct = parseBR(cbsPctRaw);
  const creditoCbs = parseBR(creditoCbsRaw);
  const creditoIbs = parseBR(creditoIbsRaw);

  const dentro = useMemo(() => (rbt12 > 0 ? calcDasPorDentroRate(rbt12, ano) : null), [rbt12, ano]);
  const fora = useMemo(() => (rbt12 > 0 ? calcDasPorForaRate(rbt12, ano, cbsPct) : null), [rbt12, ano, cbsPct]);

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-4">Simulação da reforma</h3>
        <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Ano simulado</span>
            <select value={ano} onChange={(e) => setAno(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]">
              {ANOS_REFORMA.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
          </label>
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Regime</span>
            <select value={regime} onChange={(e) => setRegime(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]">
              <option value="dentro">Simples Por Dentro (sem crédito)</option>
              <option value="fora">Simples Por Fora (regime regular IBS/CBS)</option>
            </select>
          </label>
          <CurrencyInput label="Alíquota de CBS (%)" value={cbsPctRaw} onChange={setCbsPctRaw} placeholder={sched.cbs.toFixed(1)} />
          <CurrencyInput label="Valor de venda p/ simular (R$)" value={valorVendaRaw} onChange={setValorVendaRaw} placeholder="Ex: 100,00" />
        </div>
        <p className="text-[11px] text-slate-400 mt-3">
          Nominal LC 214/2025 em {ano}: CBS {sched.cbs.toFixed(1)}% · IBS {sched.ibs.toFixed(1)}% (IBS não é editável aqui).
          RBT12 usado: <b>{fmtR(rbt12)}</b> — o mesmo informado na aba Simples Nacional.
        </p>
      </div>

      {rbt12 <= 0 && <p className="text-sm text-slate-400 italic">Informe o RBT12 na aba Simples Nacional para simular a reforma.</p>}

      {rbt12 > 0 && regime === 'dentro' && dentro && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">DAS reconstituído — {dentro.faixa && `${dentro.faixaIndex + 1}ª faixa`}</h4>
          <div className="flex items-baseline gap-3">
            <span className="text-3xl font-black text-[#2F6F4E]">{fmtPct(dentro.aliqEf, 4)}</span>
            <span className="text-sm text-slate-400">alíquota efetiva total (CBS + IBS embutidos)</span>
          </div>
          <p className="text-sm text-slate-600 mt-3">Sobre {fmtR(valorVenda)}: <b>{fmtR(valorVenda * dentro.aliqEf)}</b> de DAS</p>
        </div>
      )}

      {rbt12 > 0 && regime === 'fora' && fora && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Composição — {fora.faixa && `${fora.faixaIndex + 1}ª faixa`}</h4>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
            <div>
              <p className="text-xs text-slate-500">DAS residual</p>
              <p className="text-xl font-black text-slate-800">{fmtPct(fora.rateResidual, 4)}</p>
              <p className="text-[11px] text-slate-400">{fmtR(valorVenda * fora.rateResidual)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">CBS ({(cbsPct || sched.cbs).toFixed(1)}%)</p>
              <p className="text-xl font-black text-slate-800">{fmtPct(fora.cbsRate)}</p>
              <p className="text-[11px] text-slate-400">{fmtR(valorVenda * fora.cbsRate)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">IBS ({sched.ibs.toFixed(1)}%)</p>
              <p className="text-xl font-black text-slate-800">{fmtPct(fora.ibsRate)}</p>
              <p className="text-[11px] text-slate-400">{fmtR(valorVenda * fora.ibsRate)}</p>
            </div>
          </div>
          <div className="border-t border-slate-100 pt-3 flex items-baseline gap-3">
            <span className="text-2xl font-black text-[#2F6F4E]">{fmtPct(fora.totalRate, 4)}</span>
            <span className="text-sm text-slate-400">carga total sobre a venda (sem créditos)</span>
          </div>
        </div>
      )}

      {rbt12 > 0 && regime === 'fora' && fora && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Créditos manuais (a apurar no mês)</h4>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <CurrencyInput label="Crédito de CBS no mês (R$)" value={creditoCbsRaw} onChange={setCreditoCbsRaw} placeholder="Ex: 500,00" />
            <CurrencyInput label="Crédito de IBS no mês (R$)" value={creditoIbsRaw} onChange={setCreditoIbsRaw} placeholder="Ex: 20,00" />
          </div>
          <p className="text-[11px] text-slate-400 mt-2">
            Créditos de CBS/IBS pagos nas compras (notas de fornecedores no regime regular), a abater do valor devido no mês.
          </p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4 border-t border-slate-100 pt-4">
            <div>
              <p className="text-xs text-slate-500">CBS a pagar no mês (líquido)</p>
              <p className="text-xl font-black text-slate-800">{fmtR(Math.max(0, fatMes * fora.cbsRate - creditoCbs))}</p>
              <p className="text-[11px] text-slate-400">bruto {fmtR(fatMes * fora.cbsRate)} − crédito {fmtR(creditoCbs)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">IBS a pagar no mês (líquido)</p>
              <p className="text-xl font-black text-slate-800">{fmtR(Math.max(0, fatMes * fora.ibsRate - creditoIbs))}</p>
              <p className="text-[11px] text-slate-400">bruto {fmtR(fatMes * fora.ibsRate)} − crédito {fmtR(creditoIbs)}</p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Aba 3: Precificação ────────────────────────────────────────────────────
function cenario(venda, custo, margemHoje, rate) {
  const denom = 1 - margemHoje - rate;
  const precoSugerido = denom > 0 ? custo / denom : null;
  const aumentoValor = precoSugerido != null ? precoSugerido - venda : null;
  const aumentoPct = precoSugerido != null && venda > 0 ? aumentoValor / venda : null;
  return { precoSugerido, aumentoValor, aumentoPct };
}

function AumentoCell({ c }) {
  if (c.aumentoPct == null) return <span className="text-red-500">inviável</span>;
  return (
    <span className={c.aumentoPct > 0 ? 'text-amber-600 font-semibold' : 'text-slate-500'}>
      {c.aumentoValor >= 0 ? '+' : ''}{fmtR(c.aumentoValor)} ({fmtPct(c.aumentoPct)})
    </span>
  );
}

function PrecificacaoTab({ rbt12Raw, ano, setAno, cbsPctRaw, produtos, setProdutos }) {
  const fileInputRef = useRef(null);
  const rbt12 = parseBR(rbt12Raw);
  const cbsPct = parseBR(cbsPctRaw);

  const [novoSku, setNovoSku] = useState('');
  const [novoCusto, setNovoCusto] = useState('');
  const [novoVenda, setNovoVenda] = useState('');

  const hoje = useMemo(() => (rbt12 > 0 ? calcDasHoje(rbt12) : null), [rbt12]);
  const dentroInfo = useMemo(() => (rbt12 > 0 ? calcDasPorDentroRate(rbt12, ano) : null), [rbt12, ano]);
  const foraInfo = useMemo(() => (rbt12 > 0 ? calcDasPorForaRate(rbt12, ano, cbsPct) : null), [rbt12, ano, cbsPct]);

  const linhas = useMemo(() => {
    if (!hoje || !dentroInfo || !foraInfo) return [];
    return produtos.map((p) => {
      const dasHoje = p.venda * hoje.aliqEf;
      const lucroHoje = p.venda - p.custo - dasHoje;
      const margemHoje = p.venda > 0 ? lucroHoje / p.venda : 0;
      return {
        ...p,
        margemHoje,
        dentro: cenario(p.venda, p.custo, margemHoje, dentroInfo.aliqEf),
        fora: cenario(p.venda, p.custo, margemHoje, foraInfo.totalRate),
      };
    });
  }, [produtos, hoje, dentroInfo, foraInfo]);

  const handleFile = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseProdutos(e.target.result);
        setProdutos(parsed);
      } catch (err) {
        alert('Não foi possível ler essa planilha: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, [setProdutos]);

  const adicionarProduto = useCallback(() => {
    const custo = parseBR(novoCusto);
    const venda = parseBR(novoVenda);
    if (!novoSku.trim() || venda <= 0) return;
    setProdutos((prev) => [...prev, { sku: novoSku.trim(), custo, venda }]);
    setNovoSku('');
    setNovoCusto('');
    setNovoVenda('');
  }, [novoSku, novoCusto, novoVenda, setProdutos]);

  const removerProduto = useCallback((idx) => {
    setProdutos((prev) => prev.filter((_, i) => i !== idx));
  }, [setProdutos]);

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <div>
          <h3 className="text-sm font-bold text-slate-700">Produtos ({produtos.length})</h3>
          <p className="text-[11px] text-slate-400 mt-1">Comparando DAS de hoje x Simples Por Dentro x Simples Por Fora em</p>
        </div>
        <div className="flex items-center gap-3">
          <select value={ano} onChange={(e) => setAno(e.target.value)} className="rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]">
            {ANOS_REFORMA.map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 text-sm font-semibold bg-[#2F6F4E] text-white px-4 py-2 rounded-lg hover:bg-[#265c40] transition"
          >
            <Upload size={16} /> Importar planilha
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Adicionar produto manualmente</h4>
        <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 items-end">
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">SKU / Produto</span>
            <input type="text" value={novoSku} onChange={(e) => setNovoSku(e.target.value)} placeholder="Ex: KIT2NOVO-M"
              className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none" />
          </label>
          <CurrencyInput label="Preço de compra (custo)" value={novoCusto} onChange={setNovoCusto} placeholder="Ex: 15,00" />
          <CurrencyInput label="Preço de venda" value={novoVenda} onChange={setNovoVenda} placeholder="Ex: 35,00" />
          <button onClick={adicionarProduto} className="h-[38px] text-sm font-semibold bg-[#2F6F4E] text-white px-4 py-2 rounded-lg hover:bg-[#265c40] transition">
            Adicionar
          </button>
        </div>
      </div>

      {!hoje && <p className="text-sm text-slate-400 italic">Informe o RBT12 na aba Simples Nacional para calcular a precificação.</p>}

      {hoje && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase text-slate-500">
                <th className="px-4 py-2 bg-slate-50" rowSpan={2}>SKU</th>
                <th className="px-4 py-2 text-right bg-slate-50" rowSpan={2}>Custo</th>
                <th className="px-4 py-2 text-right bg-slate-50" colSpan={2}>Hoje</th>
                <th className="px-4 py-2 text-right bg-emerald-50" colSpan={2}>Por Dentro ({ano})</th>
                <th className="px-4 py-2 text-right bg-amber-50" colSpan={2}>Por Fora ({ano})</th>
                <th className="px-4 py-2 bg-slate-50" rowSpan={2}></th>
              </tr>
              <tr className="text-left text-[11px] uppercase text-slate-500">
                <th className="px-4 py-1 text-right bg-slate-50">Preço atual</th>
                <th className="px-4 py-1 text-right bg-slate-50">Margem líquida</th>
                <th className="px-4 py-1 text-right bg-emerald-50">Preço sugerido</th>
                <th className="px-4 py-1 text-right bg-emerald-50">Aumento</th>
                <th className="px-4 py-1 text-right bg-amber-50">Preço sugerido</th>
                <th className="px-4 py-1 text-right bg-amber-50">Aumento</th>
              </tr>
            </thead>
            <tbody>
              {linhas.map((l, i) => (
                <tr key={l.sku + i} className="border-t border-slate-100">
                  <td className="px-4 py-2 font-mono text-xs text-slate-600">{l.sku}</td>
                  <td className="px-4 py-2 text-right">{fmtR(l.custo)}</td>
                  <td className="px-4 py-2 text-right">{fmtR(l.venda)}</td>
                  <td className={`px-4 py-2 text-right font-semibold ${l.margemHoje < 0 ? 'text-red-500' : 'text-slate-700'}`}>{fmtPct(l.margemHoje)}</td>
                  <td className="px-4 py-2 text-right font-semibold text-emerald-700">{l.dentro.precoSugerido != null ? fmtR(l.dentro.precoSugerido) : '—'}</td>
                  <td className="px-4 py-2 text-right"><AumentoCell c={l.dentro} /></td>
                  <td className="px-4 py-2 text-right font-semibold text-amber-700">{l.fora.precoSugerido != null ? fmtR(l.fora.precoSugerido) : '—'}</td>
                  <td className="px-4 py-2 text-right"><AumentoCell c={l.fora} /></td>
                  <td className="px-4 py-2 text-right">
                    <button onClick={() => removerProduto(i)} className="text-xs text-slate-300 hover:text-red-500">remover</button>
                  </td>
                </tr>
              ))}
              {linhas.length === 0 && (
                <tr><td colSpan={9} className="px-4 py-8 text-center text-slate-400 italic">Nenhum produto carregado. Importe uma planilha ou adicione manualmente.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState('simples');
  const [rbt12Raw, setRbt12Raw] = useState('');
  const [faturamentoMesRaw, setFaturamentoMesRaw] = useState('');
  const [ano, setAno] = useState('2027');
  const [regime, setRegime] = useState('fora');
  const [valorVendaRaw, setValorVendaRaw] = useState('100,00');
  const [cbsPctRaw, setCbsPctRaw] = useState(String(REFORM_SCHEDULE['2027'].cbs));
  const [creditoCbsRaw, setCreditoCbsRaw] = useState('');
  const [creditoIbsRaw, setCreditoIbsRaw] = useState('');
  const [aliquotaUsadaRaw, setAliquotaUsadaRaw] = useState('');
  const [mesesRaw, setMesesRaw] = useState('5');
  const [produtos, setProdutos] = useState([]);

  // Ao trocar o ano simulado, repõe a alíquota de CBS para a nominal oficial
  // daquele ano — o usuário pode sobrescrever de novo em seguida.
  useEffect(() => {
    setCbsPctRaw(String(REFORM_SCHEDULE[ano].cbs));
  }, [ano]);

  useEffect(() => {
    fetch(produtosBaseUrl)
      .then((r) => r.arrayBuffer())
      .then((buf) => setProdutos(parseProdutos(buf)))
      .catch(() => {});
  }, []);

  const tabs = [
    { id: 'simples', label: 'Simples Nacional', icon: Calculator },
    { id: 'reforma', label: 'CBS e IBS', icon: TrendingUp },
    { id: 'precificacao', label: 'Precificação', icon: FileSpreadsheet },
    { id: 'repasse', label: 'Repasse', icon: LineChart },
  ];

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-200 px-6 py-4">
        <h1 className="text-lg font-black text-slate-800">Precificador — Simples x Reforma Tributária</h1>
        <p className="text-xs text-slate-400 mt-0.5">Comércio · Anexo I · entrada manual</p>
      </header>

      <nav className="bg-white border-b border-slate-200 px-6">
        <div className="flex gap-1">
          {tabs.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              onClick={() => setTab(id)}
              className={`flex items-center gap-2 px-4 py-3 text-sm font-semibold border-b-2 transition ${
                tab === id ? 'border-[#2F6F4E] text-[#2F6F4E]' : 'border-transparent text-slate-400 hover:text-slate-600'
              }`}
            >
              <Icon size={16} /> {label}
            </button>
          ))}
        </div>
      </nav>

      <main className="max-w-6xl mx-auto px-6 py-6">
        {tab === 'simples' && (
          <SimplesNacionalTab
            rbt12Raw={rbt12Raw}
            setRbt12Raw={setRbt12Raw}
            faturamentoMesRaw={faturamentoMesRaw}
            setFaturamentoMesRaw={setFaturamentoMesRaw}
          />
        )}
        {tab === 'reforma' && (
          <CbsIbsTab
            rbt12Raw={rbt12Raw}
            faturamentoMesRaw={faturamentoMesRaw}
            ano={ano}
            setAno={setAno}
            regime={regime}
            setRegime={setRegime}
            valorVendaRaw={valorVendaRaw}
            setValorVendaRaw={setValorVendaRaw}
            cbsPctRaw={cbsPctRaw}
            setCbsPctRaw={setCbsPctRaw}
            creditoCbsRaw={creditoCbsRaw}
            setCreditoCbsRaw={setCreditoCbsRaw}
            creditoIbsRaw={creditoIbsRaw}
            setCreditoIbsRaw={setCreditoIbsRaw}
          />
        )}
        {tab === 'precificacao' && (
          <PrecificacaoTab
            rbt12Raw={rbt12Raw}
            ano={ano}
            setAno={setAno}
            cbsPctRaw={cbsPctRaw}
            produtos={produtos}
            setProdutos={setProdutos}
          />
        )}
        {tab === 'repasse' && (
          <RepasseTab
            rbt12Raw={rbt12Raw}
            produtos={produtos}
            aliquotaUsadaRaw={aliquotaUsadaRaw}
            setAliquotaUsadaRaw={setAliquotaUsadaRaw}
            mesesRaw={mesesRaw}
            setMesesRaw={setMesesRaw}
          />
        )}
      </main>
    </div>
  );
}
