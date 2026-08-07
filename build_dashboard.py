"""
Lê os xlsx gerados por fetch_pipedrive.py / fetch_previsao.py, agrega métricas
por closer + consolidado do time, e gera publish/index.html (funil Studio Fiscal,
pipeline 62 — "inbound") e publish/outbound.html (funil Outbound, pipeline 66) a
partir do mesmo template — duas páginas separadas pra não misturar os números dos
dois funis (CC1/FPQ/CC2/CC3/Contrato enviado/Contratos de um não conta no outro).
"""
import base64, json, os
from collections import Counter
from datetime import date, timedelta
from zoneinfo import ZoneInfo
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
deals = load('deals.xlsx')
deals_p62 = load('deals_pipeline62.xlsx')
deals_p66 = load('deals_pipeline66.xlsx')
previsao = load('previsao.xlsx')

if not deals.empty:
    deals = deals[~deals['deal_id'].isin(config.DEALS_EXCLUIDOS)]
if not deals_p62.empty:
    deals_p62 = deals_p62[~deals_p62['deal_id'].isin(config.DEALS_EXCLUIDOS)]
if not deals_p66.empty:
    deals_p66 = deals_p66[~deals_p66['deal_id'].isin(config.DEALS_EXCLUIDOS)]
if not previsao.empty:
    previsao = previsao[~previsao['deal_id'].isin(config.PREVISAO_EXCLUIDOS)]

if not os.path.exists(os.path.join(WORKDIR, config.METAS_FILE)):
    raise RuntimeError(
        f'Arquivo de meta "{config.METAS_FILE}" não existe — crie um novo com a meta de '
        f'contratos do mês (copie um mês anterior como modelo).'
    )
with open(os.path.join(WORKDIR, config.METAS_FILE)) as f:
    metas = json.load(f)

_closer_nome_por_id = {c['id']: c['nome'] for c in config.CLOSERS}


def count_by_closer(df, nome):
    if df.empty or 'closer_nome' not in df.columns:
        return 0
    return int((df['closer_nome'] == nome).sum())


def pct(a, b):
    if not b:
        return None
    return round(100 * a / b, 1)


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


def daily_series(df, date_col):
    out = {}
    if df.empty or date_col not in df.columns:
        return out
    d = pd.to_datetime(df[date_col], errors='coerce').dt.date
    for day in d.dropna():
        if config.MONTH_START <= day <= config.MONTH_END:
            out[day.isoformat()] = out.get(day.isoformat(), 0) + 1
    return out


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

# ============= PREVISÃO (planilha manual — não é específica de um funil, compartilhada) =============
previsao_out = []
if not previsao.empty:
    for _, r in previsao.iterrows():
        row = r.to_dict()
        for k, v in row.items():
            if pd.isna(v):
                row[k] = None
        previsao_out.append(row)

# Placar do mês: "Previsto" continua vindo da planilha de Previsão (compartilhado,
# não é por pipeline — é a lista de trabalho manual do time). "Realizado" (Contratos
# fechados/Venda/Caixa) agora vem DIRETO do Pipedrive — mesmo negócio Ganho no mês
# que o card "Metas do mês" já conta — em vez de depender de alguém preencher a
# planilha com o link certo. Isso é calculado dentro de build_page() (por pipeline).
scoreboard_previsto = {
    'contratos_previstos': len(previsao_out),
    'venda_prevista_total': sum(r.get('venda') or 0 for r in previsao_out),
    'caixa_prevista_total': sum(r.get('caixa') or 0 for r in previsao_out),
}


def build_page(*, pipeline_id, deals_p, stage_contrato_enviado, leads_mode, meta_leads, meta_contratos_val, tem_meta):
    """Monta o dict de dados de UMA página (inbound ou outbound), filtrando CC1/CC2/
    CC3/FPQ pelo pipeline_id do negócio vinculado — pra não misturar atividade de um
    funil no número do outro.

    leads_mode:
      'cc1'    — "Leads" conta negócios distintos que o closer levou a CC1 realizada
                 no mês (funil Studio Fiscal, já em regime — ver comentário antigo).
      'criado' — "Leads" conta negócios que o closer já criou/tem como dono, criados
                 no mês (funil Outbound, onde cada closer adiciona os próprios leads).
    """
    def pf(df):
        if df.empty or 'pipeline_id' not in df.columns:
            return df.iloc[0:0]
        return df[df['pipeline_id'] == pipeline_id]

    cc1_p, cc1_fut_p = pf(cc1_re), pf(cc1_fut)
    cc2_p, cc2_fut_p = pf(cc2_re), pf(cc2_fut)
    cc3_p, cc3_fut_p = pf(cc3_re), pf(cc3_fut)
    fpq_p = pf(fpq_re)
    deals_pipe = (
        deals[deals['pipeline_id'] == pipeline_id].copy()
        if not deals.empty and 'pipeline_id' in deals.columns else deals.iloc[0:0]
    )

    deals_p = deals_p.copy()
    if not deals_p.empty:
        deals_p['closer_nome'] = deals_p['owner_id'].map(_closer_nome_por_id)

    # "Contrato enviado" contado pela etapa do negócio, não por atividade separada —
    # só pega negócios que estão *atualmente* nessa etapa (stage_change_time = quando
    # entrou).
    contrato_enviado_re = pd.DataFrame()
    if not deals_p.empty and 'stage_id' in deals_p.columns:
        contrato_enviado_re = deals_p[deals_p['stage_id'] == stage_contrato_enviado].copy()
        contrato_enviado_re = contrato_enviado_re[contrato_enviado_re['closer_nome'].notna()]

    def negocios_iniciados_por_deal(nome):
        """"Negócios iniciados" = união deduplicada de negócios que esse closer levou
        a CC1 realizada no mês OU que foram adicionados no funil Expansão Franquia
        (deals_p) já com ele como dono no mês — pega a data mais antiga entre as duas
        fontes, pra não contar o mesmo negócio duas vezes nem no dia errado."""
        by_deal = {}
        if not cc1_p.empty and 'closer_nome' in cc1_p.columns:
            d = cc1_p[cc1_p['closer_nome'] == nome].copy()
            d['marked_as_done_time'] = pd.to_datetime(d['marked_as_done_time'], errors='coerce')
            for _, r in d.dropna(subset=['marked_as_done_time']).iterrows():
                deal_id = r.get('negocio_id')
                day = r['marked_as_done_time'].date()
                if pd.notna(deal_id) and config.MONTH_START <= day <= config.MONTH_END:
                    by_deal[deal_id] = min(day, by_deal.get(deal_id, day))
        if not deals_p.empty:
            d = deals_p[deals_p['closer_nome'] == nome].copy()
            d['add_time'] = pd.to_datetime(d['add_time'], errors='coerce')
            for _, r in d.dropna(subset=['add_time']).iterrows():
                deal_id = r.get('deal_id')
                day = r['add_time'].date()
                if pd.notna(deal_id) and config.MONTH_START <= day <= config.MONTH_END:
                    by_deal[deal_id] = min(day, by_deal.get(deal_id, day))
        return by_deal

    def leads_count(nome):
        if leads_mode == 'iniciados':
            return len(negocios_iniciados_por_deal(nome))
        return sum(per_closer_daily(deals_p, 'add_time', nome).values())

    def leads_diario(nome):
        if leads_mode == 'iniciados':
            out = {}
            for day in negocios_iniciados_por_deal(nome).values():
                out[day.isoformat()] = out.get(day.isoformat(), 0) + 1
            return out
        return per_closer_daily(deals_p, 'add_time', nome)

    def closer_metrics(nome):
        m = {
            'cc1_realizada': count_by_closer(cc1_p, nome),
            'cc1_futura': count_by_closer(cc1_fut_p, nome),
            'fpq_realizada': count_by_closer(fpq_p, nome),
            'cc2_realizada': count_by_closer(cc2_p, nome),
            'cc2_futura': count_by_closer(cc2_fut_p, nome),
            'cc3_realizada': count_by_closer(cc3_p, nome),
            'cc3_futura': count_by_closer(cc3_fut_p, nome),
            'contrato_enviado': count_by_closer(contrato_enviado_re, nome),
            'leads': leads_count(nome),
            'contratos': sum(per_closer_daily(deals_pipe, 'won_time', nome, status='won').values()),
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
            'leads': leads_diario(c['nome']),
            'cc1_realizada': per_closer_daily(cc1_p, 'marked_as_done_time', c['nome']),
            'cc1_futura': per_closer_daily(cc1_fut_p, 'due_date', c['nome']),
            'fpq_realizada': per_closer_daily(fpq_p, 'marked_as_done_time', c['nome']),
            'cc2_realizada': per_closer_daily(cc2_p, 'marked_as_done_time', c['nome']),
            'cc2_futura': per_closer_daily(cc2_fut_p, 'due_date', c['nome']),
            'cc3_realizada': per_closer_daily(cc3_p, 'marked_as_done_time', c['nome']),
            'cc3_futura': per_closer_daily(cc3_fut_p, 'due_date', c['nome']),
            'contrato_enviado': per_closer_daily(contrato_enviado_re, 'stage_change_time', c['nome']),
            'contratos': per_closer_daily(deals_pipe, 'won_time', c['nome'], status='won'),
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
        'leads': sum(c['leads'] for c in closers_out),  # ver override logo abaixo (leads_funil)
        'contratos': sum(c['contratos'] for c in closers_out),
    }
    consolidado['conversao'] = {
        'cc1_para_fpq': pct(consolidado['fpq_realizada'], consolidado['cc1_realizada']),
        'fpq_para_cc2': pct(consolidado['cc2_realizada'], consolidado['fpq_realizada']),
        'cc2_para_cc3': pct(consolidado['cc3_realizada'], consolidado['cc2_realizada']),
    }

    # Total do funil (todo negócio criado no mês, qualquer dono — inclui SDR/
    # Qualificadora) — guardado só como referência (leads_funil), não usado como
    # "Negócios iniciados" no modo 'iniciados' (ver abaixo).
    daily_leads_funil = {}
    if not deals_p.empty:
        d = deals_p.copy()
        d['add_time'] = pd.to_datetime(d['add_time'], errors='coerce').dt.date
        for day in d['add_time'].dropna():
            if config.MONTH_START <= day <= config.MONTH_END:
                daily_leads_funil[day.isoformat()] = daily_leads_funil.get(day.isoformat(), 0) + 1
    consolidado['leads_funil'] = sum(daily_leads_funil.values())
    consolidado['leads_atribuidos_closers'] = consolidado['leads']

    if leads_mode == 'iniciados':
        # "Negócios iniciados" no Ritmo/Metas usa a MESMA definição do Relatório por
        # closer (só o que entrou na propriedade de um dos 4 closers — SDR/
        # Qualificadora não conta), não o total do funil. consolidado['leads'] já é
        # a soma dos closers (calculada acima), não sobrescreve.
        daily_leads = {}
        for c in closers_out:
            for day, count in c['diario']['leads'].items():
                daily_leads[day] = daily_leads.get(day, 0) + count
    else:
        daily_leads = daily_leads_funil
        consolidado['leads'] = consolidado['leads_funil']

    daily_contrato_enviado = daily_series(contrato_enviado_re, 'stage_change_time')

    daily_contratos = {}
    if not deals_pipe.empty:
        dw = deals_pipe[deals_pipe['status'] == 'won'].copy()
        dw['won_time'] = pd.to_datetime(dw['won_time'], errors='coerce').dt.date
        for day in dw['won_time'].dropna():
            if config.MONTH_START <= day <= config.MONTH_END:
                daily_contratos[day.isoformat()] = daily_contratos.get(day.isoformat(), 0) + 1

    daily_cc1 = daily_series(cc1_p, 'marked_as_done_time')
    daily_fpq = daily_series(fpq_p, 'marked_as_done_time')
    daily_cc2 = daily_series(cc2_p, 'marked_as_done_time')
    daily_cc3 = daily_series(cc3_p, 'marked_as_done_time')

    # Metas de CC1/FPQ/CC2/CC3 estimadas em cascata a partir da meta de leads
    # (config.BENCHMARK_*). Contrato enviado continua estimado pela conversão real
    # do mês (CC3 -> Contrato enviado). Sem meta_leads (ex: Outbound ainda sem meta),
    # tudo cascateia pra None.
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
            'label': 'Negócios iniciados', 'meta': meta_leads, 'estimado': False,
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
            'label': 'Contratos fechados', 'meta': meta_contratos_val, 'estimado': False,
            'daily': daily_contratos, 'ideal_cumulative': ideal_cumulative(meta_contratos_val, n_dias_uteis_total),
        },
    }

    # ============= PERDIDOS =============
    perdidos = {'total': 0, 'pela_qualificadora': 0, 'apos_repasse': 0, 'por_motivo': []}
    if not deals_p.empty:
        dl = deals_p[deals_p['status'] == 'lost'].copy()
        dl['lost_time'] = pd.to_datetime(dl['lost_time'], errors='coerce').dt.date
        dl = dl[(dl['lost_time'] >= config.MONTH_START) & (dl['lost_time'] <= config.MONTH_END)]
        perdidos['total'] = int(len(dl))
        # Quem perdeu o negócio define a atribuição, não a etapa em que ele estava.
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
        leads_count_ = Counter()
        if not deals_p.empty:
            dm = deals_p.copy()
            dm['add_time'] = pd.to_datetime(dm['add_time'], errors='coerce').dt.date
            dm = dm[(dm['add_time'] >= config.MONTH_START) & (dm['add_time'] <= config.MONTH_END)]
            for v in dm[col].dropna():
                v = str(v).strip()
                if v:
                    leads_count_[v] += 1
        cc1_count = Counter()
        if not cc1_p.empty and col in cc1_p.columns:
            for v in cc1_p[col].dropna():
                v = str(v).strip()
                if v:
                    cc1_count[v] += 1
        itens = []
        for nome in set(leads_count_) | set(cc1_count):
            leads_n = leads_count_.get(nome, 0)
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

    # ============= MAPA DE CALOR — HORÁRIOS DE CC1 REALIZADA (seg-sex, 09h-20h) =============
    HEATMAP_DIAS = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
    HEATMAP_HORAS = list(range(9, 21))
    heatmap_matriz = [[0] * len(HEATMAP_HORAS) for _ in HEATMAP_DIAS]
    if not cc1_p.empty and 'marked_as_done_time' in cc1_p.columns:
        hm = cc1_p.copy()
        hm['marked_as_done_time'] = pd.to_datetime(hm['marked_as_done_time'], errors='coerce')
        for _, r in hm.dropna(subset=['marked_as_done_time']).iterrows():
            wd = r['marked_as_done_time'].weekday()
            if wd > 4:
                continue
            hora = r['marked_as_done_time'].hour
            if hora in HEATMAP_HORAS:
                heatmap_matriz[wd][HEATMAP_HORAS.index(hora)] += 1
    heatmap_cc1 = {'dias': HEATMAP_DIAS, 'horas': HEATMAP_HORAS, 'matriz': heatmap_matriz}

    # ============= FPQS RECEBIDAS (lista, com link pro Pipedrive) =============
    fpqs_recebidas = []
    if not fpq_p.empty:
        fq = fpq_p.copy()
        fq['marked_as_done_time'] = pd.to_datetime(fq['marked_as_done_time'], errors='coerce')
        fq = fq.dropna(subset=['marked_as_done_time']).sort_values('marked_as_done_time', ascending=False)
        for _, r in fq.iterrows():
            fpqs_recebidas.append({
                'data': r['marked_as_done_time'].strftime('%Y-%m-%d %H:%M'),
                'closer_nome': None if pd.isna(r.get('closer_nome')) else r.get('closer_nome'),
                'pessoa_nome': None if pd.isna(r.get('pessoa_nome')) else r.get('pessoa_nome'),
                'negocio_titulo': None if pd.isna(r.get('negocio_titulo')) else r.get('negocio_titulo'),
                'deal_link': None if pd.isna(r.get('deal_link')) else r.get('deal_link'),
            })

    # ============= PLACAR DO MÊS — "Realizado" direto do Pipedrive =============
    # Mesmo escopo do "Contratos fechados" do Metas do mês (Ganho, won_time no mês
    # corrente, negócio desse pipeline, excluídos os de config.DEALS_EXCLUIDOS) —
    # não depende mais de alguém preencher a planilha com o link certo.
    dw = deals_pipe[deals_pipe['status'] == 'won'].copy() if not deals_pipe.empty else deals_pipe
    if not dw.empty:
        dw['won_time'] = pd.to_datetime(dw['won_time'], errors='coerce')
        dw = dw[(dw['won_time'].dt.date >= config.MONTH_START) & (dw['won_time'].dt.date <= config.MONTH_END)]
    scoreboard = {
        **scoreboard_previsto,
        'contratos_fechados': int(len(dw)),
        'venda_realizada_total': float(dw['valor'].fillna(0).sum()) if not dw.empty else 0.0,
        'caixa_realizado_total': float(dw['caixa'].fillna(0).sum()) if not dw.empty else 0.0,
    }

    return {
        'meta_contratos': meta_contratos_val,
        'meta_leads': meta_leads,
        'tem_meta': tem_meta,
        'closers': closers_out,
        'consolidado': consolidado,
        'pace': pace,
        'perdidos': perdidos,
        'ranking': ranking,
        'heatmap_cc1': heatmap_cc1,
        'fpqs_recebidas': fpqs_recebidas,
        'scoreboard': scoreboard,
    }


common = {
    'mes_label': config.MONTH_LABEL,
    'periodo_inicio': config.MONTH_START.isoformat(),
    'periodo_fim': config.MONTH_END.isoformat(),
    'gerado_em': pd.Timestamp.now(tz=ZoneInfo('America/Sao_Paulo')).strftime('%Y-%m-%dT%H:%M'),
    'ritmo': ritmo,
    'previsao': previsao_out,
}

inbound_data = build_page(
    pipeline_id=config.PIPELINE_STUDIO_FISCAL_FRANQUIA,
    deals_p=deals_p62,
    stage_contrato_enviado=config.PIPELINE_STAGE_CONTRATO_ENVIADO,
    leads_mode='iniciados',
    meta_leads=metas.get('meta_leads'),
    meta_contratos_val=metas['meta_contratos'],
    tem_meta=True,
)
outbound_data = build_page(
    pipeline_id=config.PIPELINE_OUTBOUND_STUDIO_FISCAL,
    deals_p=deals_p66,
    stage_contrato_enviado=config.PIPELINE_STAGE_CONTRATO_ENVIADO_OUTBOUND,
    leads_mode='criado',
    meta_leads=None,
    meta_contratos_val=None,
    tem_meta=False,
)

# ============= ARQUIVAMENTO AUTOMÁTICO DE MÊS FECHADO =============
# Todo ciclo (mesmo os automáticos, sem ninguém olhando) grava o mês corrente TAMBÉM
# num arquivo próprio (ex: agosto-2026.html), além do index.html normal — com os
# MESMOS dados. Isso é o que garante o congelamento: assim que o mês vira, o próximo
# ciclo passa a calcular um MONTH_START/END diferente (via config.py) e escreve num
# nome de arquivo diferente (setembro-2026.html) — o arquivo do mês anterior simplesmente
# para de ser tocado, ficando congelado sozinho com o último estado gravado antes da
# virada (a planilha de Previsão pode continuar a mesma, sendo reaproveitada mês a mês,
# sem afetar o que já foi congelado). Não depende de ninguém lembrar de rodar nada
# manualmente no fim do mês.
_MES_NOMES = {1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 6: 'junho',
              7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'}


def _arquivo_mes(ano, mes, outbound=False):
    nome = _MES_NOMES[mes]
    return f'outbound-{nome}-{ano}.html' if outbound else f'{nome}-{ano}.html'


_cur_ano, _cur_mes = config.MONTH_START.year, config.MONTH_START.month
_prev_mes = 12 if _cur_mes == 1 else _cur_mes - 1
_prev_ano = _cur_ano - 1 if _cur_mes == 1 else _cur_ano

ARQUIVO_MES_ATUAL = _arquivo_mes(_cur_ano, _cur_mes)
ARQUIVO_MES_ATUAL_OUTBOUND = _arquivo_mes(_cur_ano, _cur_mes, outbound=True)

_arq_anterior = _arquivo_mes(_prev_ano, _prev_mes)
_arq_anterior_ob = _arquivo_mes(_prev_ano, _prev_mes, outbound=True)
# Só linka pro mês anterior se o arquivo já existir (evita link morto no primeiro
# ciclo de um mês novo, antes do mês anterior ainda ter sido congelado nenhuma vez).
MES_ANTERIOR_ARQUIVO = _arq_anterior if os.path.exists(os.path.join(PUBLISH_DIR, _arq_anterior)) else None
MES_ANTERIOR_ARQUIVO_OUTBOUND = _arq_anterior_ob if os.path.exists(os.path.join(PUBLISH_DIR, _arq_anterior_ob)) else None
MES_ANTERIOR_LABEL = f'Ver {_MES_NOMES[_prev_mes].capitalize()} (fechado) →'

dashboard_inbound = {
    **common,
    **inbound_data,
    'pagina_titulo': 'Equipe de Vendas — Studio Fiscal',
    'pagina_outra_url': 'outbound.html',
    'pagina_outra_label': 'Ver Outbound →',
    'mes_outro_url': MES_ANTERIOR_ARQUIVO,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'perdidos_hint': '— negócios perdidos no funil Studio Fiscal Franquia no mês',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. "Negócios iniciados" conta a união '
        '(sem duplicar) de negócios que o closer levou a CC1 realizada no mês OU que '
        'entraram no funil Expansão Franquia Studio Fiscal já com ele como dono no mês'
    ),
    'relatorio_metrics': [
        ['leads', 'Negócios iniciados'],
        ['cc1_realizada', 'CC1 realizada'],
        ['cc1_futura', 'CC1 a acontecer'],
        ['fpq_realizada', 'FPQ recebida'],
        ['cc2_realizada', 'CC2 realizada'],
        ['cc2_futura', 'CC2 a acontecer'],
        ['cc3_realizada', 'CC3 realizada'],
        ['cc3_futura', 'CC3 a acontecer'],
        ['contrato_enviado', 'Contrato enviado'],
        ['contratos', 'Contratos fechados'],
    ],
}
dashboard_outbound = {
    **common,
    **outbound_data,
    'pagina_titulo': 'Equipe de Vendas — Outbound',
    'pagina_outra_url': 'index.html',
    'pagina_outra_label': '← Ver Studio Fiscal (Inbound)',
    'mes_outro_url': MES_ANTERIOR_ARQUIVO_OUTBOUND,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'perdidos_hint': '— negócios perdidos no funil Outbound Studio Fiscal no mês',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. "Leads" conta negócios do funil '
        'Outbound criados no mês já com um closer como dono'
    ),
    # Aqui "Leads" continua — ainda não tem CC1 nesse funil, então é a única
    # métrica de produtividade disponível por enquanto.
    'relatorio_metrics': [
        ['leads', 'Leads'],
        ['cc1_realizada', 'CC1 realizada'],
        ['cc1_futura', 'CC1 a acontecer'],
        ['fpq_realizada', 'FPQ recebida'],
        ['cc2_realizada', 'CC2 realizada'],
        ['cc2_futura', 'CC2 a acontecer'],
        ['cc3_realizada', 'CC3 realizada'],
        ['cc3_futura', 'CC3 a acontecer'],
        ['contrato_enviado', 'Contrato enviado'],
        ['contratos', 'Contratos fechados'],
    ],
}


def render(dashboard, out_name):
    template_path = os.path.join(WORKDIR, 'dashboard_template.html')
    with open(template_path, encoding='utf-8') as f:
        html = f.read()
    html = html.replace('__DATA_PLACEHOLDER__', json.dumps(dashboard, ensure_ascii=False))
    html = html.replace('__MONTH_LABEL__', config.MONTH_LABEL)
    with open(os.path.join(WORKDIR, 'logo_studio_fiscal.png'), 'rb') as f:
        logo_b64 = base64.b64encode(f.read()).decode('ascii')
    html = html.replace('__LOGO_DATA_URI__', f'data:image/png;base64,{logo_b64}')
    out_path = os.path.join(PUBLISH_DIR, out_name)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Dashboard gerado em {out_path}')
    return out_path


render(dashboard_inbound, 'index.html')
render(dashboard_outbound, 'outbound.html')

# Snapshot do mês corrente (ver comentário acima, "ARQUIVAMENTO AUTOMÁTICO DE MÊS
# FECHADO") — mesmo conteúdo do index.html/outbound.html, só troca o link de
# navegação: em vez de apontar pro mês anterior, aponta de volta pro mês atual
# (genérico, sem citar o mês por nome, pra não ficar desatualizado quando esse
# arquivo congelar e passar a ser visto meses depois).
render({**dashboard_inbound, 'mes_outro_url': 'index.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL)
render({**dashboard_outbound, 'mes_outro_url': 'outbound.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL_OUTBOUND)

ci = dashboard_inbound['consolidado']
print(f"Inbound (Studio Fiscal) — CC1={ci['cc1_realizada']} | FPQ={ci['fpq_realizada']} | "
      f"CC2={ci['cc2_realizada']} | CC3={ci['cc3_realizada']} | "
      f"Contratos={ci['contratos']}/{metas['meta_contratos']}")
co = dashboard_outbound['consolidado']
print(f"Outbound — CC1={co['cc1_realizada']} | FPQ={co['fpq_realizada']} | "
      f"CC2={co['cc2_realizada']} | CC3={co['cc3_realizada']} | "
      f"Leads criados no mês={co['leads_atribuidos_closers']} | Contratos={co['contratos']}")
