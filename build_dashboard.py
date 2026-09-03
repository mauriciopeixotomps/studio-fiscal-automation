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
from math import ceil
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
deals_p65 = load('deals_pipeline65.xlsx')  # REBARBA STUDIO FISCAL FRANQUIA
deals_p69 = load('deals_pipeline69.xlsx')  # OUTBOUND - EVENTOS (reportado junto com o 66)
deals_produtos = load('deals_produtos.xlsx')  # tipo_contrato: 'franquia' (tem valor) / 'partner' (de graça)
previsao = load('previsao.xlsx')
agendadores_leads = load('agendadores_leads_recebidos.xlsx')
agendadores_agendamentos = load('agendadores_agendamentos.xlsx')
qualificacao_re = load('activities_qualificacao_realizada.xlsx')
no_show_re = load('activities_no_show.xlsx')

if not deals.empty:
    deals = deals[~deals['deal_id'].isin(config.DEALS_EXCLUIDOS)]
    if not deals_produtos.empty:
        deals = deals.merge(deals_produtos[['deal_id', 'tipo_contrato']], on='deal_id', how='left')
    else:
        deals['tipo_contrato'] = None
if not deals_p62.empty:
    deals_p62 = deals_p62[~deals_p62['deal_id'].isin(config.DEALS_EXCLUIDOS)]
if not deals_p66.empty:
    deals_p66 = deals_p66[~deals_p66['deal_id'].isin(config.DEALS_EXCLUIDOS)]
if not deals_p65.empty:
    deals_p65 = deals_p65[~deals_p65['deal_id'].isin(config.DEALS_EXCLUIDOS)]
if not deals_p69.empty:
    deals_p69 = deals_p69[~deals_p69['deal_id'].isin(config.DEALS_EXCLUIDOS)]
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

# Placar do mês: "Previsto" vem da planilha de Previsão, filtrado pela coluna FUNIL
# (Inbound/Outbound/Rebarba) — calculado dentro de build_page(), por página, não
# mais aqui (ver comentário lá). "Realizado" (Contratos fechados/Venda/Caixa) vem
# DIRETO do Pipedrive — mesmo negócio Ganho no mês que o card "Metas do mês" já
# conta — em vez de depender de alguém preencher a planilha com o link certo.

# ============= TIME DE AGENDAMENTO (SDRs dedicados à Studio Fiscal) =============
# "Realizada"/"No-show" cruzam o negócio do agendamento com os CC1 Realizada e os
# no-show já buscados pros closers (cc1_re / no_show_re) — a CC1 em si (e o no-show
# dela) ficam na agenda do closer, não do agendador. Restrito ao pipeline 62 (funil
# Studio Fiscal) — é o único que os agendadores atendem.
_cc1_re_p62 = cc1_re[cc1_re['pipeline_id'] == config.PIPELINE_STUDIO_FISCAL_FRANQUIA] if not cc1_re.empty and 'pipeline_id' in cc1_re.columns else cc1_re
_no_show_p62 = no_show_re[no_show_re['pipeline_id'] == config.PIPELINE_STUDIO_FISCAL_FRANQUIA] if not no_show_re.empty and 'pipeline_id' in no_show_re.columns else no_show_re
_cc1_realizado_deal_ids = set(_cc1_re_p62['negocio_id'].dropna()) if not _cc1_re_p62.empty and 'negocio_id' in _cc1_re_p62.columns else set()
_no_show_deal_ids = set(_no_show_p62['negocio_id'].dropna()) if not _no_show_p62.empty and 'negocio_id' in _no_show_p62.columns else set()

agendadores_out = []
for a in config.AGENDADORES:
    # "Leads recebidos" = negócios do funil 62 que já têm esse agendador como dono,
    # adicionados no mês — mesma definição de "Negócios iniciados" usada em todo o
    # resto do painel. Antes contava a tarefa "1ª tentativa de contato" (uma etapa
    # seguinte, criada só depois que alguém começa a trabalhar o lead) — o time
    # confirmou que quer ver o negócio assim que ele entra no funil, não só quando
    # essa tarefa aparece (01/09/2026).
    leads_n = 0
    if not deals_p62.empty and 'owner_id' in deals_p62.columns and 'add_time' in deals_p62.columns:
        _dd = deals_p62[deals_p62['owner_id'] == a['id']].copy()
        _dd['add_time'] = pd.to_datetime(_dd['add_time'], errors='coerce').dt.date
        leads_n = int(_dd['add_time'].apply(
            lambda x: pd.notna(x) and config.MONTH_START <= x <= config.MONTH_END
        ).sum())
    agendamento_deal_ids = set()
    if not agendadores_agendamentos.empty and 'agendador_nome' in agendadores_agendamentos.columns:
        agendamento_deal_ids = set(
            agendadores_agendamentos.loc[agendadores_agendamentos['agendador_nome'] == a['nome'], 'negocio_id'].dropna()
        )
    n_agendamentos = len(agendamento_deal_ids)
    realizadas_n = len(agendamento_deal_ids & _cc1_realizado_deal_ids)
    no_show_n = len(agendamento_deal_ids & _no_show_deal_ids)
    meta_agendamentos = leads_n * config.META_AGENDAMENTO_PCT
    meta_realizadas = n_agendamentos * config.META_AGENDADOR_REALIZADA_PCT
    agendadores_out.append({
        'nome': a['nome'],
        'leads_recebidos': int(leads_n),
        'agendamentos': n_agendamentos,
        'realizadas': realizadas_n,
        'no_show': no_show_n,
        'pct_agendamento': pct(n_agendamentos, leads_n),
        'pct_realizada': pct(realizadas_n, n_agendamentos),
        'pct_no_show': pct(no_show_n, n_agendamentos),
        # Quanto falta pra bater a própria meta: agendamentos até chegar em 30% dos
        # leads recebidos, e realizadas até chegar em 65% dos agendamentos já feitos
        # (não recalcula com agendamentos hipotéticos futuros).
        'faltam_agendamentos': max(0, ceil(meta_agendamentos - n_agendamentos)),
        'faltam_realizadas': max(0, ceil(meta_realizadas - realizadas_n)),
    })
agendadores_meta = {
    'meta_agendamento_pct': round(config.META_AGENDAMENTO_PCT * 100),
    'meta_realizada_pct': round(config.META_AGENDADOR_REALIZADA_PCT * 100),
}


def build_page(*, pipeline_ids, deals_p, stages_contrato_enviado, leads_mode, meta_leads, meta_contratos_val, tem_meta, funil_nome,
                scoreboard_pipeline_ids=None, previsao_todos_funis=False):
    """Monta o dict de dados de UMA página (inbound, outbound, rebarba...), filtrando
    CC1/CC2/CC3/FPQ pelo pipeline_id do negócio vinculado — pra não misturar atividade
    de um funil no número do outro. pipeline_ids é uma lista pra dar pra juntar mais de
    um pipeline numa mesma página (ex: Outbound = funil próprio + Eventos, ver chamada
    mais abaixo); stages_contrato_enviado é a lista das etapas "CONTRATO ENVIADO"
    correspondentes (uma por pipeline, IDs diferentes em cada um).

    leads_mode:
      'cc1'    — "Leads" conta negócios distintos que o closer levou a CC1 realizada
                 no mês (funil Studio Fiscal, já em regime — ver comentário antigo).
      'criado' — "Leads" conta negócios que o closer já criou/tem como dono, criados
                 no mês (funil Outbound, onde cada closer adiciona os próprios leads).
    """
    def pf(df):
        if df.empty or 'pipeline_id' not in df.columns:
            return df.iloc[0:0]
        return df[df['pipeline_id'].isin(pipeline_ids)]

    cc1_p, cc1_fut_p = pf(cc1_re), pf(cc1_fut)
    cc2_p, cc2_fut_p = pf(cc2_re), pf(cc2_fut)
    cc3_p, cc3_fut_p = pf(cc3_re), pf(cc3_fut)
    fpq_p = pf(fpq_re)
    no_show_p = pf(no_show_re)
    deals_pipe = (
        deals[deals['pipeline_id'].isin(pipeline_ids)].copy()
        if not deals.empty and 'pipeline_id' in deals.columns else deals.iloc[0:0]
    )
    # Escopo à parte só pro Placar do mês (Realizado) — o Inbound passa a somar
    # Ganho dos 4 pipelines da Fiscal juntos (01/09/2026, o time quer ver todo
    # fechamento — inclusive Rebarba/Outbound — no painel principal), enquanto
    # Metas/Consolidado/Comparativo/Relatório continuam só do pipeline dessa
    # página (senão um contrato da Rebarba contaria em dois lugares diferentes).
    _scoreboard_pipeline_ids = scoreboard_pipeline_ids or pipeline_ids
    deals_pipe_scoreboard = (
        deals[deals['pipeline_id'].isin(_scoreboard_pipeline_ids)].copy()
        if not deals.empty and 'pipeline_id' in deals.columns else deals.iloc[0:0]
    )

    deals_p = deals_p.copy()
    if not deals_p.empty:
        deals_p['closer_nome'] = deals_p['owner_id'].map(_closer_nome_por_id)

    # "Contrato enviado" contado pela etapa do negócio, não por atividade separada —
    # só pega negócios que estão *atualmente* nessa etapa E ainda em aberto
    # (stage_change_time = quando entrou). Sem o status=='open', um negócio que
    # passou por essa etapa e depois foi Ganho ou Perdido continuava contando pra
    # sempre (o stage_id gravado não muda depois que o negócio fecha) — corrigido
    # 29/08/2026.
    contrato_enviado_re = pd.DataFrame()
    if not deals_p.empty and 'stage_id' in deals_p.columns:
        contrato_enviado_re = deals_p[
            deals_p['stage_id'].isin(stages_contrato_enviado) & (deals_p['status'] == 'open')
        ].copy()
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

    def contratos_por_tipo(nome, tipo):
        """Conta contratos (Ganho, won_time no mês) do closer, filtrando por
        tipo_contrato ('franquia' tem valor / 'partner' de graça / None = sem
        produto vinculado ou não reconhecido — ver fetch_pipedrive.py)."""
        if deals_pipe.empty or 'closer_nome' not in deals_pipe.columns:
            return 0
        d = deals_pipe[(deals_pipe['closer_nome'] == nome) & (deals_pipe['status'] == 'won')].copy()
        d = d[d['tipo_contrato'].isna()] if tipo is None else d[d['tipo_contrato'] == tipo]
        if d.empty:
            return 0
        d['won_time'] = pd.to_datetime(d['won_time'], errors='coerce').dt.date
        return int(d['won_time'].apply(lambda x: pd.notna(x) and config.MONTH_START <= x <= config.MONTH_END).sum())

    def contratos_diario_por_tipo(nome, tipo):
        if deals_pipe.empty or 'closer_nome' not in deals_pipe.columns:
            return {}
        d = deals_pipe[(deals_pipe['closer_nome'] == nome) & (deals_pipe['status'] == 'won')].copy()
        d = d[d['tipo_contrato'].isna()] if tipo is None else d[d['tipo_contrato'] == tipo]
        d['won_time'] = pd.to_datetime(d['won_time'], errors='coerce').dt.date
        out = {}
        for day in d['won_time'].dropna():
            if config.MONTH_START <= day <= config.MONTH_END:
                out[day.isoformat()] = out.get(day.isoformat(), 0) + 1
        return out

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
            'contratos_franquia': contratos_por_tipo(nome, 'franquia'),
            'contratos_partner': contratos_por_tipo(nome, 'partner'),
            'no_show': count_by_closer(no_show_p, nome),
        }
        m['conversao'] = {
            'cc1_para_fpq': pct(m['fpq_realizada'], m['cc1_realizada']),
            'fpq_para_cc2': pct(m['cc2_realizada'], m['fpq_realizada']),
            'cc2_para_cc3': pct(m['cc3_realizada'], m['cc2_realizada']),
        }
        # % no-show = do total de CC1 que deveriam acontecer (realizada + no-show),
        # quantas viraram no-show — não usa cc1_futura (essas ainda não aconteceram).
        m['pct_no_show'] = pct(m['no_show'], m['cc1_realizada'] + m['no_show'])
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
            'contratos_franquia': contratos_diario_por_tipo(c['nome'], 'franquia'),
            'contratos_partner': contratos_diario_por_tipo(c['nome'], 'partner'),
            'no_show': per_closer_daily(no_show_p, 'marked_as_done_time', c['nome']),
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
        'contratos_franquia': sum(c['contratos_franquia'] for c in closers_out),
        'contratos_partner': sum(c['contratos_partner'] for c in closers_out),
        'no_show': sum(c['no_show'] for c in closers_out),
    }
    consolidado['conversao'] = {
        'cc1_para_fpq': pct(consolidado['fpq_realizada'], consolidado['cc1_realizada']),
        'fpq_para_cc2': pct(consolidado['cc2_realizada'], consolidado['fpq_realizada']),
        'cc2_para_cc3': pct(consolidado['cc3_realizada'], consolidado['cc2_realizada']),
    }
    consolidado['pct_no_show'] = pct(consolidado['no_show'], consolidado['cc1_realizada'] + consolidado['no_show'])

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
        # "Negócios iniciados" no Ritmo/Metas conta todo negócio do funil marcado pela
        # Qualificadora com a etiqueta da equipe (config.LABEL_STUDIO_FISCAL — "Lead
        # Studio Fiscal"), adicionado no mês, independente de já ter dono entre os 4
        # closers ou ainda estar com a Qualificadora — é o volume real de geração de
        # lead que a meta (ex: 1500) foi calibrada pra medir. Diferente do "Negócios
        # iniciados" por closer no Relatório (CC1 realizada OU adicionado já com o
        # closer como dono), que mede outra coisa (produtividade individual) e continua
        # como estava.
        def _tem_etiqueta(v):
            # O Excel/pandas pode ler o campo como texto ("97") ou como número
            # (97.0, quando toda a coluna só tem valores de etiqueta única) —
            # normaliza os dois casos antes de comparar. Também cobre o formato
            # multi-select original da API ("69,97").
            if pd.isna(v):
                return False
            s = str(v)
            if s.endswith('.0'):
                s = s[:-2]
            return config.LABEL_STUDIO_FISCAL in [p.strip() for p in s.split(',')]

        daily_leads = {}
        if not deals_p.empty and 'label' in deals_p.columns:
            d = deals_p.copy()
            d['add_time'] = pd.to_datetime(d['add_time'], errors='coerce').dt.date
            tem_etiqueta = d['label'].apply(_tem_etiqueta)
            for day in d.loc[tem_etiqueta, 'add_time'].dropna():
                if config.MONTH_START <= day <= config.MONTH_END:
                    daily_leads[day.isoformat()] = daily_leads.get(day.isoformat(), 0) + 1
        consolidado['leads'] = sum(daily_leads.values())
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
    # Ganho, won_time no mês corrente, excluídos os de config.DEALS_EXCLUIDOS —
    # não depende mais de alguém preencher a planilha com o link certo. Usa
    # deals_pipe_scoreboard (não deals_pipe) pra poder somar mais de um pipeline
    # aqui sem misturar com o resto da página (ver comentário acima).
    dw = deals_pipe_scoreboard[deals_pipe_scoreboard['status'] == 'won'].copy() if not deals_pipe_scoreboard.empty else deals_pipe_scoreboard
    if not dw.empty:
        dw['won_time'] = pd.to_datetime(dw['won_time'], errors='coerce')
        dw = dw[(dw['won_time'].dt.date >= config.MONTH_START) & (dw['won_time'].dt.date <= config.MONTH_END)]

    # "Previsto" é filtrado pela coluna FUNIL da planilha (Inbound/Outbound/Rebarba)
    # — cada página só vê os negócios previstos do próprio funil (31/08/2026),
    # EXCETO o Inbound, que a partir de 01/09/2026 mostra a lista inteira (todos os
    # funis juntos) com o rodapé quebrando por funil — é o painel principal, o time
    # quer ver tudo ali. Linha sem FUNIL preenchido não aparece em nenhuma página
    # (não dá pra saber de qual é) até alguém preencher.
    previsao_funil = previsao_out if previsao_todos_funis else [r for r in previsao_out if r.get('funil') == funil_nome]
    previsao_pendente = [r for r in previsao_funil if r.get('deal_status') != 'won']
    scoreboard_previsto = {
        'contratos_previstos': len(previsao_pendente),
        'venda_prevista_total': sum(r.get('venda') or 0 for r in previsao_pendente),
        'caixa_prevista_total': sum(r.get('caixa') or 0 for r in previsao_pendente),
    }
    # Franquia (tem valor) x Partner (de graça) dentro do mesmo "Contratos
    # fechados" do Placar — mesma classificação (produto do Pipedrive, com
    # fallback pro Modelo da planilha) já usada no Relatório por closer
    # (01/09/2026, pedido do time: ver de cara quantos são franquia x parceria).
    dw_franquia = dw[dw['tipo_contrato'] == 'franquia'] if not dw.empty and 'tipo_contrato' in dw.columns else dw.iloc[0:0]
    dw_partner = dw[dw['tipo_contrato'] == 'partner'] if not dw.empty and 'tipo_contrato' in dw.columns else dw.iloc[0:0]
    scoreboard = {
        **scoreboard_previsto,
        'contratos_fechados': int(len(dw)),
        'contratos_fechados_franquia': int(len(dw_franquia)),
        'contratos_fechados_partner': int(len(dw_partner)),
        'venda_realizada_total': float(dw['valor'].fillna(0).sum()) if not dw.empty else 0.0,
        'venda_realizada_franquia': float(dw_franquia['valor'].fillna(0).sum()) if not dw_franquia.empty else 0.0,
        'venda_realizada_partner': float(dw_partner['valor'].fillna(0).sum()) if not dw_partner.empty else 0.0,
        'caixa_realizado_total': float(dw['caixa'].fillna(0).sum()) if not dw.empty else 0.0,
        'caixa_realizado_franquia': float(dw_franquia['caixa'].fillna(0).sum()) if not dw_franquia.empty else 0.0,
        'caixa_realizado_partner': float(dw_partner['caixa'].fillna(0).sum()) if not dw_partner.empty else 0.0,
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
        'previsao': previsao_funil,
        'scoreboard': scoreboard,
    }


common = {
    'mes_label': config.MONTH_LABEL,
    'periodo_inicio': config.MONTH_START.isoformat(),
    'periodo_fim': config.MONTH_END.isoformat(),
    'gerado_em': pd.Timestamp.now(tz=ZoneInfo('America/Sao_Paulo')).strftime('%Y-%m-%dT%H:%M'),
    'ritmo': ritmo,
}

inbound_data = build_page(
    pipeline_ids=[config.PIPELINE_STUDIO_FISCAL_FRANQUIA],
    deals_p=deals_p62,
    stages_contrato_enviado=[config.PIPELINE_STAGE_CONTRATO_ENVIADO],
    leads_mode='iniciados',
    meta_leads=metas.get('meta_leads'),
    meta_contratos_val=metas['meta_contratos'],
    tem_meta=True,
    funil_nome='INBOUND',
    # Pedido do time (01/09/2026): o Inbound é o painel principal, então o Placar
    # do mês (Realizado) soma Ganho dos 4 pipelines da Fiscal juntos — Inbound,
    # Rebarba e os dois de Outbound — não só o 62. A Previsão também mostra a
    # lista inteira (todos os funis), com o rodapé quebrando por funil.
    scoreboard_pipeline_ids=[
        config.PIPELINE_STUDIO_FISCAL_FRANQUIA,
        config.PIPELINE_REBARBA_STUDIO_FISCAL,
        config.PIPELINE_OUTBOUND_STUDIO_FISCAL,
        config.PIPELINE_OUTBOUND_EVENTOS,
    ],
    previsao_todos_funis=True,
)
# "Consolidado do time" (Inbound) ganha 2 etapas antes da CC1 — Leads iniciados (já
# calculado acima) e Leads qualificados (atividade Qualificação Realizada, feita pelo
# time de qualificação) — mais "Agendado"/"Realizado" no lugar do que era só "CC1
# realizada" (31/08/2026). Só no Inbound: é a etapa que antecede o funil 62.
_qualif_p62 = (
    qualificacao_re[qualificacao_re['pipeline_id'] == config.PIPELINE_STUDIO_FISCAL_FRANQUIA]
    if not qualificacao_re.empty and 'pipeline_id' in qualificacao_re.columns else qualificacao_re.iloc[0:0]
)
inbound_data['consolidado']['leads_qualificados'] = int(len(_qualif_p62))
inbound_data['consolidado']['cc1_agendado'] = (
    inbound_data['consolidado']['cc1_realizada'] + inbound_data['consolidado']['cc1_futura']
)
# Outbound = funil próprio (66) + Eventos (69) juntos numa página só, a pedido do
# time (29/08/2026) — "Eventos" não tem etiqueta de marca nenhuma e é 100% Fiscal,
# não precisa filtrar por dono.
deals_p_outbound = pd.concat([deals_p66, deals_p69], ignore_index=True) if not deals_p66.empty or not deals_p69.empty else deals_p66
outbound_data = build_page(
    pipeline_ids=[config.PIPELINE_OUTBOUND_STUDIO_FISCAL, config.PIPELINE_OUTBOUND_EVENTOS],
    deals_p=deals_p_outbound,
    stages_contrato_enviado=[config.PIPELINE_STAGE_CONTRATO_ENVIADO_OUTBOUND, config.PIPELINE_STAGE_CONTRATO_ENVIADO_EVENTOS],
    leads_mode='criado',
    meta_leads=None,
    meta_contratos_val=None,
    tem_meta=False,
    funil_nome='OUTBOUND',
)
# Rebarba (65) — reativação de perdidos, página própria, ainda sem meta (29/08/2026).
rebarba_data = build_page(
    pipeline_ids=[config.PIPELINE_REBARBA_STUDIO_FISCAL],
    deals_p=deals_p65,
    stages_contrato_enviado=[config.PIPELINE_STAGE_CONTRATO_ENVIADO_REBARBA],
    leads_mode='criado',
    meta_leads=None,
    meta_contratos_val=None,
    tem_meta=False,
    funil_nome='REBARBA',
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


def _arquivo_mes(ano, mes, prefixo=None):
    nome = _MES_NOMES[mes]
    return f'{prefixo}-{nome}-{ano}.html' if prefixo else f'{nome}-{ano}.html'


_cur_ano, _cur_mes = config.MONTH_START.year, config.MONTH_START.month
_prev_mes = 12 if _cur_mes == 1 else _cur_mes - 1
_prev_ano = _cur_ano - 1 if _cur_mes == 1 else _cur_ano

ARQUIVO_MES_ATUAL = _arquivo_mes(_cur_ano, _cur_mes)
ARQUIVO_MES_ATUAL_OUTBOUND = _arquivo_mes(_cur_ano, _cur_mes, prefixo='outbound')
ARQUIVO_MES_ATUAL_REBARBA = _arquivo_mes(_cur_ano, _cur_mes, prefixo='rebarba')

# Calculado só por matemática de data (NÃO checa se o arquivo existe no disco local):
# a automação roda tanto na nuvem (GitHub Actions, checkout novo a cada execução, sem
# nenhum HTML antigo local) quanto no Mac (pasta publish/ persistente) — se o link
# dependesse de os.path.exists() local, a nuvem sempre acharia que o mês anterior não
# existe e ia sobrescrever o link correto que o Mac gerou, alternando entre os dois a
# cada ciclo (foi exatamente o que causou o link "Ver Julho" sumir em 10/08/2026).
MES_ANTERIOR_ARQUIVO = _arquivo_mes(_prev_ano, _prev_mes)
MES_ANTERIOR_ARQUIVO_OUTBOUND = _arquivo_mes(_prev_ano, _prev_mes, prefixo='outbound')
MES_ANTERIOR_ARQUIVO_REBARBA = _arquivo_mes(_prev_ano, _prev_mes, prefixo='rebarba')
MES_ANTERIOR_LABEL = f'Ver {_MES_NOMES[_prev_mes].capitalize()} (fechado) →'


def _historico(inicio, prefixo, url_atual):
    """Lista TODOS os meses de `inicio` (ano, mes) até o mês corrente — não só o
    anterior. Mês corrente aponta pro arquivo "vivo" (index.html/outbound.html/
    rebarba.html); os passados apontam pro arquivo congelado de cada um. Mais
    recente primeiro (fica no topo do menu "Histórico")."""
    itens = []
    ano, mes = inicio
    while (ano, mes) <= (_cur_ano, _cur_mes):
        label = f'{_MES_NOMES[mes].capitalize()} {ano}'
        if (ano, mes) == (_cur_ano, _cur_mes):
            itens.append({'url': url_atual, 'label': label + ' (atual)'})
        else:
            itens.append({'url': _arquivo_mes(ano, mes, prefixo=prefixo), 'label': label})
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
    return list(reversed(itens))


HISTORICO_INBOUND = _historico(config.HISTORICO_INICIO_INBOUND, None, 'index.html')
HISTORICO_OUTBOUND = _historico(config.HISTORICO_INICIO_OUTBOUND, 'outbound', 'outbound.html')
HISTORICO_REBARBA = _historico(config.HISTORICO_INICIO_REBARBA, 'rebarba', 'rebarba.html')

PAGINAS_IRMAS = {
    'inbound': {'url': 'index.html', 'label': 'Studio Fiscal (Inbound)'},
    'outbound': {'url': 'outbound.html', 'label': 'Outbound'},
    'rebarba': {'url': 'rebarba.html', 'label': 'Rebarba (reativação)'},
}


def _paginas_irmas(exceto):
    return [v for k, v in PAGINAS_IRMAS.items() if k != exceto]


dashboard_inbound = {
    **common,
    **inbound_data,
    'agendadores': agendadores_out,
    'agendadores_meta': agendadores_meta,
    'pagina_titulo': 'Equipe de Vendas — Studio Fiscal',
    'paginas_irmas': _paginas_irmas('inbound'),
    'historico': HISTORICO_INBOUND,
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
        ['contratos_franquia', 'Contratos Franquia (c/ valor)'],
        ['contratos_partner', 'Contratos Partner (grátis)'],
        ['no_show', 'No-show'],
    ],
}
dashboard_outbound = {
    **common,
    **outbound_data,
    'pagina_titulo': 'Equipe de Vendas — Outbound',
    'paginas_irmas': _paginas_irmas('outbound'),
    'historico': HISTORICO_OUTBOUND,
    'mes_outro_url': MES_ANTERIOR_ARQUIVO_OUTBOUND,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'perdidos_hint': '— negócios perdidos nos funis Outbound + Eventos da Studio Fiscal no mês',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. Junta o funil Outbound próprio (66) '
        'com o Outbound - Eventos (69). "Leads" conta negócios criados no mês já com '
        'um closer como dono'
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
        ['contratos_franquia', 'Contratos Franquia (c/ valor)'],
        ['contratos_partner', 'Contratos Partner (grátis)'],
        ['no_show', 'No-show'],
    ],
}
dashboard_rebarba = {
    **common,
    **rebarba_data,
    'pagina_titulo': 'Equipe de Vendas — Rebarba (reativação)',
    'paginas_irmas': _paginas_irmas('rebarba'),
    'historico': HISTORICO_REBARBA,
    'mes_outro_url': MES_ANTERIOR_ARQUIVO_REBARBA,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'perdidos_hint': '— negócios perdidos no funil Rebarba Studio Fiscal Franquia no mês',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. Reativação de negócios perdidos do '
        'funil Studio Fiscal Franquia. "Leads" conta negócios criados no mês já com '
        'um closer como dono'
    ),
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
        ['contratos_franquia', 'Contratos Franquia (c/ valor)'],
        ['contratos_partner', 'Contratos Partner (grátis)'],
        ['no_show', 'No-show'],
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
render(dashboard_rebarba, 'rebarba.html')

# Snapshot do mês corrente (ver comentário acima, "ARQUIVAMENTO AUTOMÁTICO DE MÊS
# FECHADO") — mesmo conteúdo do index.html/outbound.html/rebarba.html, só troca o
# link de navegação: em vez de apontar pro mês anterior, aponta de volta pro mês
# atual (genérico, sem citar o mês por nome, pra não ficar desatualizado quando esse
# arquivo congelar e passar a ser visto meses depois).
render({**dashboard_inbound, 'mes_outro_url': 'index.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL)
render({**dashboard_outbound, 'mes_outro_url': 'outbound.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL_OUTBOUND)
render({**dashboard_rebarba, 'mes_outro_url': 'rebarba.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL_REBARBA)

ci = dashboard_inbound['consolidado']
print(f"Inbound (Studio Fiscal) — CC1={ci['cc1_realizada']} | FPQ={ci['fpq_realizada']} | "
      f"CC2={ci['cc2_realizada']} | CC3={ci['cc3_realizada']} | "
      f"Contratos={ci['contratos']}/{metas['meta_contratos']}")
co = dashboard_outbound['consolidado']
print(f"Outbound — CC1={co['cc1_realizada']} | FPQ={co['fpq_realizada']} | "
      f"CC2={co['cc2_realizada']} | CC3={co['cc3_realizada']} | "
      f"Leads criados no mês={co['leads_atribuidos_closers']} | Contratos={co['contratos']}")
