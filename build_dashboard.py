"""
Lê os xlsx gerados por fetch_pipedrive.py / fetch_previsao.py, agrega métricas
por closer + consolidado do time, e gera publish/index.html a partir do template.
"""
import base64, json, os
from collections import Counter
from datetime import date, timedelta
import pandas as pd

import config

WORKDIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(WORKDIR, 'data_pipedrive')
PUBLISH_DIR = os.path.join(WORKDIR, 'publish')
os.makedirs(PUBLISH_DIR, exist_ok=True)


def load(name):
    path = os.path.join(DATA_DIR, name)
    if not os.path.exists(path):
        return pd.DataFrame()
    return pd.read_excel(path)


cc1_re = load('activities_cc1_realizada.xlsx')
cc1_fut = load('activities_cc1_futura.xlsx')
cc2_re = load('activities_cc2_realizada.xlsx')
cc2_fut = load('activities_cc2_futura.xlsx')
cc3_re = load('activities_cc3_realizada.xlsx')
cc3_fut = load('activities_cc3_futura.xlsx')
fpq_re = load('activities_fpq_realizada.xlsx')
contrato_enviado_re = load('activities_contrato_enviado_realizada.xlsx')
deals = load('deals.xlsx')
deals_p62 = load('deals_pipeline62.xlsx')
previsao = load('previsao.xlsx')

if not os.path.exists(os.path.join(WORKDIR, config.METAS_FILE)):
    raise RuntimeError(
        f'Arquivo de meta "{config.METAS_FILE}" não existe — crie um novo com a meta de '
        f'contratos do mês (copie um mês anterior como modelo).'
    )
with open(os.path.join(WORKDIR, config.METAS_FILE)) as f:
    metas = json.load(f)


def count_by_closer(df, nome):
    if df.empty or 'closer_nome' not in df.columns:
        return 0
    return int((df['closer_nome'] == nome).sum())


def pct(a, b):
    if not b:
        return None
    return round(100 * a / b, 1)


def deals_leads_count(nome):
    if deals.empty:
        return 0
    d = deals[deals['closer_nome'] == nome].copy()
    d['add_time'] = pd.to_datetime(d['add_time'], errors='coerce')
    mask = (d['add_time'].dt.date >= config.MONTH_START) & (d['add_time'].dt.date <= config.MONTH_END)
    return int(mask.sum())


def deals_contratos_count(nome):
    if deals.empty:
        return 0
    d = deals[(deals['closer_nome'] == nome) & (deals['status'] == 'won')].copy()
    d['won_time'] = pd.to_datetime(d['won_time'], errors='coerce')
    mask = (d['won_time'].dt.date >= config.MONTH_START) & (d['won_time'].dt.date <= config.MONTH_END)
    return int(mask.sum())


def per_closer_daily(df, date_col, nome, status=None):
    if df.empty or 'closer_nome' not in df.columns or date_col not in df.columns:
        return {}
    d = df[df['closer_nome'] == nome]
    if status is not None:
        d = d[d.get('status') == status]
    d = d.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors='coerce').dt.date
    out = {}
    for day in d[date_col].dropna():
        if config.MONTH_START <= day <= config.MONTH_END:
            out[day.isoformat()] = out.get(day.isoformat(), 0) + 1
    return out


def closer_metrics(nome):
    m = {
        'cc1_realizada': count_by_closer(cc1_re, nome),
        'cc1_futura': count_by_closer(cc1_fut, nome),
        'fpq_realizada': count_by_closer(fpq_re, nome),
        'cc2_realizada': count_by_closer(cc2_re, nome),
        'cc2_futura': count_by_closer(cc2_fut, nome),
        'cc3_realizada': count_by_closer(cc3_re, nome),
        'cc3_futura': count_by_closer(cc3_fut, nome),
        'contrato_enviado': count_by_closer(contrato_enviado_re, nome),
        'leads': deals_leads_count(nome),
        'contratos': deals_contratos_count(nome),
    }
    m['conversao'] = {
        'cc1_para_fpq': pct(m['fpq_realizada'], m['cc1_realizada']),
        'fpq_para_cc2': pct(m['cc2_realizada'], m['fpq_realizada']),
        'cc2_para_cc3': pct(m['cc3_realizada'], m['cc2_realizada']),
    }
    return m


closers_out = []
for c in config.CLOSERS:
    metrics = closer_metrics(c['nome'])
    diario = {
        'leads': per_closer_daily(deals, 'add_time', c['nome']),
        'cc1_realizada': per_closer_daily(cc1_re, 'marked_as_done_time', c['nome']),
        'fpq_realizada': per_closer_daily(fpq_re, 'marked_as_done_time', c['nome']),
        'cc2_realizada': per_closer_daily(cc2_re, 'marked_as_done_time', c['nome']),
        'cc3_realizada': per_closer_daily(cc3_re, 'marked_as_done_time', c['nome']),
        'contrato_enviado': per_closer_daily(contrato_enviado_re, 'marked_as_done_time', c['nome']),
        'contratos': per_closer_daily(deals, 'won_time', c['nome'], status='won'),
    }
    closers_out.append({'nome': c['nome'], 'foto': c['foto'], **metrics, 'diario': diario})

consolidado = {
    'cc1_realizada': sum(c['cc1_realizada'] for c in closers_out),
    'cc1_futura': sum(c['cc1_futura'] for c in closers_out),
    'fpq_realizada': sum(c['fpq_realizada'] for c in closers_out),
    'cc2_realizada': sum(c['cc2_realizada'] for c in closers_out),
    'cc2_futura': sum(c['cc2_futura'] for c in closers_out),
    'cc3_realizada': sum(c['cc3_realizada'] for c in closers_out),
    'cc3_futura': sum(c['cc3_futura'] for c in closers_out),
    'contrato_enviado': sum(c['contrato_enviado'] for c in closers_out),
    'leads': sum(c['leads'] for c in closers_out),  # negócios sob responsabilidade dos closers (não é o total do funil — ver 'leads_funil')
    'contratos': sum(c['contratos'] for c in closers_out),
}
consolidado['conversao'] = {
    'cc1_para_fpq': pct(consolidado['fpq_realizada'], consolidado['cc1_realizada']),
    'fpq_para_cc2': pct(consolidado['cc2_realizada'], consolidado['fpq_realizada']),
    'cc2_para_cc3': pct(consolidado['cc3_realizada'], consolidado['cc2_realizada']),
}

# ============= RITMO (série diária) =============
def daily_series(df, date_col):
    out = {}
    if df.empty or date_col not in df.columns:
        return out
    d = pd.to_datetime(df[date_col], errors='coerce').dt.date
    for day in d.dropna():
        if config.MONTH_START <= day <= config.MONTH_END:
            out[day.isoformat()] = out.get(day.isoformat(), 0) + 1
    return out


# "Negócios iniciados" (leads) = TODO negócio criado no funil STUDIO FISCAL FRANQUIA (pipeline
# 62) no mês, independente do responsável — o funil ainda é majoritariamente da Qualificadora,
# então filtrar só pelos 4 closers subestimava muito o volume real de entrada.
daily_leads = {}
if not deals_p62.empty:
    d = deals_p62.copy()
    d['add_time'] = pd.to_datetime(d['add_time'], errors='coerce').dt.date
    for day in d['add_time'].dropna():
        if config.MONTH_START <= day <= config.MONTH_END:
            daily_leads[day.isoformat()] = daily_leads.get(day.isoformat(), 0) + 1
consolidado['leads_funil'] = sum(daily_leads.values())
consolidado['leads_atribuidos_closers'] = consolidado['leads']
consolidado['leads'] = consolidado['leads_funil']

daily_contrato_enviado = daily_series(contrato_enviado_re, 'marked_as_done_time')

daily_contratos = {}
if not deals.empty:
    dw = deals[deals['status'] == 'won'].copy()
    dw['won_time'] = pd.to_datetime(dw['won_time'], errors='coerce').dt.date
    for day in dw['won_time'].dropna():
        if config.MONTH_START <= day <= config.MONTH_END:
            daily_contratos[day.isoformat()] = daily_contratos.get(day.isoformat(), 0) + 1

daily_cc1 = daily_series(cc1_re, 'marked_as_done_time')
daily_fpq = daily_series(fpq_re, 'marked_as_done_time')
daily_cc2 = daily_series(cc2_re, 'marked_as_done_time')
daily_cc3 = daily_series(cc3_re, 'marked_as_done_time')

n_dias_uteis_total = sum(
    1 for i in range((config.MONTH_END - config.MONTH_START).days + 1)
    if (config.MONTH_START + timedelta(days=i)).weekday() < 5
)
n_dias_uteis_passados = sum(
    1 for i in range((min(date.today(), config.MONTH_END) - config.MONTH_START).days + 1)
    if (config.MONTH_START + timedelta(days=i)).weekday() < 5
)

ritmo = {
    'dias_uteis_total': n_dias_uteis_total,
    'dias_uteis_passados': max(n_dias_uteis_passados, 0),
}


def ideal_cumulative(meta, dias_uteis_total):
    """Curva de ritmo ideal: reparte a meta pelos dias úteis do mês (acumulado)."""
    if not meta or not dias_uteis_total:
        return {}
    increment = meta / dias_uteis_total
    out = {}
    util_count = 0
    cum = 0.0
    for i in range((config.MONTH_END - config.MONTH_START).days + 1):
        day = config.MONTH_START + timedelta(days=i)
        if day.weekday() < 5:
            util_count += 1
            cum = increment * util_count
        out[day.isoformat()] = round(cum, 1)
    return out


# Metas explícitas: leads (informada pelo time) e contratos (metas_*.json).
# Metas de CC1/FPQ/CC2/CC3 são estimadas em cascata a partir da meta de leads,
# usando a conversão de referência do time (config.BENCHMARK_*, "conta pra cada
# contrato": 100 leads -> 15 CC1 -> 8 FPQ -> 4 CC2 -> 2 CC3 -> 1 fechado) em vez
# da conversão real do mês — que é instável/ruidosa no início do mês, com poucos dados.
# Contrato enviado não está nesse quadro de referência, então continua estimado pela
# conversão real do mês (CC3 -> Contrato enviado).
meta_leads = metas.get('meta_leads')
meta_cc1_calc = round(meta_leads * config.BENCHMARK_LEADS_PARA_CC1) if meta_leads else None
meta_fpq_calc = round(meta_cc1_calc * config.BENCHMARK_CC1_PARA_FPQ) if meta_cc1_calc else None
meta_cc2_calc = round(meta_fpq_calc * config.BENCHMARK_FPQ_PARA_CC2) if meta_fpq_calc else None
meta_cc3_calc = round(meta_cc2_calc * config.BENCHMARK_CC2_PARA_CC3) if meta_cc2_calc else None
cc3_para_contrato_enviado = pct(consolidado['contrato_enviado'], consolidado['cc3_realizada'])
meta_contrato_enviado_calc = (
    round(meta_cc3_calc * cc3_para_contrato_enviado / 100)
    if meta_cc3_calc and cc3_para_contrato_enviado else None
)

pace = {
    'leads': {
        'label': 'Leads', 'meta': meta_leads, 'estimado': False,
        'daily': daily_leads, 'ideal_cumulative': ideal_cumulative(meta_leads, n_dias_uteis_total),
    },
    'cc1': {
        'label': 'CC1 realizada', 'meta': meta_cc1_calc, 'estimado': True,
        'daily': daily_cc1, 'ideal_cumulative': ideal_cumulative(meta_cc1_calc, n_dias_uteis_total),
    },
    'fpq': {
        'label': 'FPQ recebida', 'meta': meta_fpq_calc, 'estimado': True,
        'daily': daily_fpq, 'ideal_cumulative': ideal_cumulative(meta_fpq_calc, n_dias_uteis_total),
    },
    'cc2': {
        'label': 'CC2 realizada', 'meta': meta_cc2_calc, 'estimado': True,
        'daily': daily_cc2, 'ideal_cumulative': ideal_cumulative(meta_cc2_calc, n_dias_uteis_total),
    },
    'cc3': {
        'label': 'CC3 realizada', 'meta': meta_cc3_calc, 'estimado': True,
        'daily': daily_cc3, 'ideal_cumulative': ideal_cumulative(meta_cc3_calc, n_dias_uteis_total),
    },
    'contrato_enviado': {
        'label': 'Contrato enviado', 'meta': meta_contrato_enviado_calc, 'estimado': True,
        'daily': daily_contrato_enviado, 'ideal_cumulative': ideal_cumulative(meta_contrato_enviado_calc, n_dias_uteis_total),
    },
    'contratos': {
        'label': 'Contratos fechados', 'meta': metas['meta_contratos'], 'estimado': False,
        'daily': daily_contratos, 'ideal_cumulative': ideal_cumulative(metas['meta_contratos'], n_dias_uteis_total),
    },
}

# ============= PERDIDOS (funil Studio Fiscal Franquia) =============
perdidos = {'total': 0, 'pela_qualificadora': 0, 'apos_repasse': 0, 'por_motivo': []}
if not deals_p62.empty:
    dl = deals_p62[deals_p62['status'] == 'lost'].copy()
    dl['lost_time'] = pd.to_datetime(dl['lost_time'], errors='coerce').dt.date
    dl = dl[(dl['lost_time'] >= config.MONTH_START) & (dl['lost_time'] <= config.MONTH_END)]
    perdidos['total'] = int(len(dl))
    # Quem perdeu o negócio define a atribuição, não a etapa em que ele estava —
    # o closer às vezes faz a CC1 e já marca como perdido ainda na etapa de MQL
    # (Qualificadora), então a etapa sozinha classificaria errado. Se o dono do
    # negócio é um dos 4 closers, é perdido pela Expansão; senão, é SDR/Qualificadora.
    dl['pela_qualificadora'] = ~dl['owner_id'].isin(config.CLOSER_IDS)
    perdidos['pela_qualificadora'] = int(dl['pela_qualificadora'].sum())
    perdidos['apos_repasse'] = perdidos['total'] - perdidos['pela_qualificadora']
    dl['motivo'] = dl['lost_reason'].fillna('Não informado').replace('', 'Não informado')
    motivo_totais = Counter(dl['motivo']).most_common()
    motivo_qualificadora = Counter(dl[dl['pela_qualificadora']]['motivo'])
    motivo_expansao = Counter(dl[~dl['pela_qualificadora']]['motivo'])
    perdidos['por_motivo'] = [
        {'motivo': k, 'total': int(v), 'qualificadora': int(motivo_qualificadora.get(k, 0)), 'expansao': int(motivo_expansao.get(k, 0))}
        for k, v in motivo_totais
    ]

# ============= RANKING CAMPANHAS / CRIATIVOS (por CC1 realizada) =============
def ranking_utm(col):
    leads_count = Counter()
    if not deals_p62.empty:
        dm = deals_p62.copy()
        dm['add_time'] = pd.to_datetime(dm['add_time'], errors='coerce').dt.date
        dm = dm[(dm['add_time'] >= config.MONTH_START) & (dm['add_time'] <= config.MONTH_END)]
        for v in dm[col].dropna():
            v = str(v).strip()
            if v:
                leads_count[v] += 1
    cc1_count = Counter()
    if not cc1_re.empty and col in cc1_re.columns:
        cc1_p62 = cc1_re[cc1_re.get('pipeline_id') == config.PIPELINE_STUDIO_FISCAL_FRANQUIA] if 'pipeline_id' in cc1_re.columns else cc1_re
        for v in cc1_p62[col].dropna():
            v = str(v).strip()
            if v:
                cc1_count[v] += 1
    itens = []
    for nome in set(leads_count) | set(cc1_count):
        leads_n = leads_count.get(nome, 0)
        cc1_n = cc1_count.get(nome, 0)
        itens.append({'nome': nome, 'leads': leads_n, 'cc1': cc1_n, 'conversao': pct(cc1_n, leads_n)})
    melhores = sorted([i for i in itens if i['cc1'] > 0], key=lambda i: (-i['cc1'], -(i['conversao'] or 0)))[:5]
    piores_pool = [i for i in itens if i['leads'] >= 3] or [i for i in itens if i['leads'] >= 1]
    piores = sorted(piores_pool, key=lambda i: (i['conversao'] if i['conversao'] is not None else 999, i['cc1']))[:5]
    return {'melhores': melhores, 'piores': piores}


ranking = {
    'campanhas': ranking_utm('utm_campanha'),
    'criativos': ranking_utm('utm_criativo'),
}

# ============= PREVISÃO =============
previsao_out = []
if not previsao.empty:
    for _, r in previsao.iterrows():
        row = r.to_dict()
        for k, v in row.items():
            if pd.isna(v):
                row[k] = None
        previsao_out.append(row)

# Placar com base no preenchimento da planilha de previsão: só conta como "fechado"
# quando o negócio vinculado ("Link do pipe") já está Ganho no Pipedrive.
scoreboard = {
    'contratos_fechados': sum(1 for r in previsao_out if r.get('venda_realizada') is not None),
    'venda_realizada_total': sum(r.get('venda_realizada') or 0 for r in previsao_out),
    'caixa_realizado_total': sum(r.get('caixa_realizado') or 0 for r in previsao_out),
    'contratos_previstos': len(previsao_out),
    'venda_prevista_total': sum(r.get('venda') or 0 for r in previsao_out),
    'caixa_prevista_total': sum(r.get('caixa') or 0 for r in previsao_out),
}

dashboard = {
    'mes_label': config.MONTH_LABEL,
    'periodo_inicio': config.MONTH_START.isoformat(),
    'periodo_fim': config.MONTH_END.isoformat(),
    'gerado_em': pd.Timestamp.now().isoformat(timespec='minutes'),
    'meta_contratos': metas['meta_contratos'],
    'meta_leads': meta_leads,
    'closers': closers_out,
    'consolidado': consolidado,
    'ritmo': ritmo,
    'pace': pace,
    'perdidos': perdidos,
    'ranking': ranking,
    'scoreboard': scoreboard,
    'previsao': previsao_out,
}

template_path = os.path.join(WORKDIR, 'dashboard_template.html')
with open(template_path, encoding='utf-8') as f:
    html = f.read()

html = html.replace('__DATA_PLACEHOLDER__', json.dumps(dashboard, ensure_ascii=False))
html = html.replace('__MONTH_LABEL__', config.MONTH_LABEL)

with open(os.path.join(WORKDIR, 'logo_studio_fiscal.png'), 'rb') as f:
    logo_b64 = base64.b64encode(f.read()).decode('ascii')
html = html.replace('__LOGO_DATA_URI__', f'data:image/png;base64,{logo_b64}')

out_path = os.path.join(PUBLISH_DIR, 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Dashboard gerado em {out_path}')
print(f"Consolidado: CC1 realizada={consolidado['cc1_realizada']} | FPQ={consolidado['fpq_realizada']} | "
      f"CC2 realizada={consolidado['cc2_realizada']} | CC3 realizada={consolidado['cc3_realizada']} | "
      f"Contratos={consolidado['contratos']}/{metas['meta_contratos']}")
