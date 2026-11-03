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
deals_p72 = load('deals_pipeline72.xlsx')  # FENALAW 2026 (compartilhado com a Agro)
deals_produtos = load('deals_produtos.xlsx')  # tipo_contrato: 'franquia' (tem valor) / 'partner' (de graça)
previsao = load('previsao.xlsx')
agendadores_leads = load('agendadores_leads_recebidos.xlsx')
deals_agendadores = load('deals_agendadores.xlsx')  # negócio do agendador em QUALQUER pipeline
agendadores_agendamentos = load('agendadores_agendamentos.xlsx')
qualificacao_re = load('activities_qualificacao_realizada.xlsx')
no_show_re = load('activities_no_show.xlsx')


def _dedup_mesmo_dia(df):
    """Atividade duplicada no Pipedrive (mesmo negócio, mesmo tipo, marcada no
    mesmo dia — ex: FPQ lançada duas vezes) contava duas vezes no painel. Fica
    uma por negócio por dia (16/10/2026: divergência de CC1/FPQ entre o painel,
    o Pipedrive e a conta do closer)."""
    if df.empty or 'negocio_id' not in df.columns or 'marked_as_done_time' not in df.columns:
        return df
    dia = pd.to_datetime(df['marked_as_done_time'], errors='coerce').dt.date
    sem_negocio = df['negocio_id'].isna()
    chave = df.assign(_dia=dia)
    return pd.concat([chave[sem_negocio], chave[~sem_negocio].drop_duplicates(subset=['negocio_id', '_dia'])]).drop(columns=['_dia']).sort_index()


cc1_re = _dedup_mesmo_dia(cc1_re)
cc2_re = _dedup_mesmo_dia(cc2_re)
cc3_re = _dedup_mesmo_dia(cc3_re)
fpq_re = _dedup_mesmo_dia(fpq_re)
no_show_re = _dedup_mesmo_dia(no_show_re)

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
if not deals_p72.empty:
    deals_p72 = deals_p72[~deals_p72['deal_id'].isin(config.DEALS_EXCLUIDOS)]
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


def per_closer_daily_sum(df, date_col, value_col, nome, status=None):
    """Igual a per_closer_daily, mas SOMA value_col por dia em vez de contar
    linhas — usado pro Venda/Caixa realizado diário (página Período
    personalizado, 12/09/2026: "todos os indicadores refletirem o período
    escolhido" incluía o Placar, que até então só existia como total do mês)."""
    if df.empty or 'closer_nome' not in df.columns or date_col not in df.columns or value_col not in df.columns:
        return {}
    d = df[df['closer_nome'] == nome]
    if status is not None:
        d = d[d.get('status') == status]
    d = d.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors='coerce').dt.date
    out = {}
    for day, val in zip(d[date_col], d[value_col]):
        if pd.notna(day) and config.MONTH_START <= day <= config.MONTH_END and pd.notna(val):
            key = day.isoformat()
            out[key] = out.get(key, 0) + float(val)
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


def _dia_by_dealid(df, deal_col, date_col):
    """{negocio_id: 'AAAA-MM-DD'} — 1ª data válida por negócio (usado pra
    dar granularidade diária ao Time de agendamento; ver 'diario' abaixo)."""
    if df.empty or deal_col not in df.columns or date_col not in df.columns:
        return {}
    d = df[[deal_col, date_col]].copy()
    d[date_col] = pd.to_datetime(d[date_col], errors='coerce').dt.date
    out = {}
    for _, r in d.dropna(subset=[date_col]).iterrows():
        out.setdefault(r[deal_col], r[date_col].isoformat())
    return out


_agendamento_dia_by_deal = _dia_by_dealid(agendadores_agendamentos, 'negocio_id', 'marked_as_done_time')
_cc1_dia_by_deal = _dia_by_dealid(_cc1_re_p62, 'negocio_id', 'marked_as_done_time')
_no_show_dia_by_deal = _dia_by_dealid(_no_show_p62, 'negocio_id', 'marked_as_done_time')


def _contar_por_dia(deal_ids, dia_by_deal):
    """{'AAAA-MM-DD': n} pra um conjunto de negócios, só dias dentro do mês
    corrente — mesmo padrão do 'diario' dos closers (ver per_closer_daily)."""
    out = {}
    for deal_id in deal_ids:
        dia = dia_by_deal.get(deal_id)
        if not dia:
            continue
        d = date.fromisoformat(dia)
        if config.MONTH_START <= d <= config.MONTH_END:
            out[dia] = out.get(dia, 0) + 1
    return out

agendadores_out = []
for a in config.AGENDADORES:
    # "Leads recebidos" = TODO negócio que já tem esse agendador como dono, em
    # QUALQUER pipeline (não só o 62/Studio Fiscal) — adicionado no mês. Antes
    # contava só o funil 62 (01/09/2026), mas o time confirmou que o SDR pode
    # receber lead que ainda não foi roteado pro funil da Fiscal (ex: ainda numa
    # etapa de qualificação compartilhada como a FRANQUIA/42) e isso tem que
    # contar igual (03/09/2026) — antes disso contava a tarefa "1ª tentativa de
    # contato" (uma etapa seguinte, só depois de alguém começar a trabalhar).
    leads_n = 0
    leads_dia = {}
    if not deals_agendadores.empty and 'agendador_id' in deals_agendadores.columns and 'add_time' in deals_agendadores.columns:
        _dd = deals_agendadores[deals_agendadores['agendador_id'] == a['id']].copy()
        add_time_dt = pd.to_datetime(_dd['add_time'], errors='coerce')
        # Lead REATRIBUÍDO de outro SDR (o dono mudou muito tempo depois de
        # criado) não conta como "recebido" por quem é dono agora — senão
        # infla o denominador de quem recebeu o lote e derruba a conversão
        # dele à toa. Caso real (19/10/2026): 337 negócios transferidos de
        # Leonardo Soares pra Jonatans Junior de uma vez (bulk update
        # confirmado no histórico do Pipedrive), 5 deles criados em outubro.
        if 'update_time' in _dd.columns:
            update_time_dt = pd.to_datetime(_dd['update_time'], errors='coerce')
            gap_dias = (update_time_dt - add_time_dt).dt.days
            nao_reatribuido = gap_dias.isna() | (gap_dias.abs() <= 7)
        else:
            nao_reatribuido = True
        _dd['add_time'] = add_time_dt.dt.date
        _dd_mes = _dd[(add_time_dt.dt.date.apply(lambda x: pd.notna(x) and config.MONTH_START <= x <= config.MONTH_END)) & nao_reatribuido]
        leads_n = int(len(_dd_mes))
        for d in _dd_mes['add_time']:
            leads_dia[d.isoformat()] = leads_dia.get(d.isoformat(), 0) + 1
    agendamento_deal_ids = set()
    if not agendadores_agendamentos.empty and 'agendador_nome' in agendadores_agendamentos.columns:
        agendamento_deal_ids = set(
            agendadores_agendamentos.loc[agendadores_agendamentos['agendador_nome'] == a['nome'], 'negocio_id'].dropna()
        )
    n_agendamentos = len(agendamento_deal_ids)
    realizadas_ids = agendamento_deal_ids & _cc1_realizado_deal_ids
    no_show_ids = agendamento_deal_ids & _no_show_deal_ids
    realizadas_n = len(realizadas_ids)
    no_show_n = len(no_show_ids)
    # Granularidade diária de tudo acima — pra "Ver o dia" (Relatório por
    # closer) também filtrar o Time de agendamento, não só o resto da página
    # (pedido do time, 16/09/2026: "não aparece" o histórico do SDR num
    # período específico).
    diario_agendador = {
        'leads_recebidos': leads_dia,
        'agendamentos': _contar_por_dia(agendamento_deal_ids, _agendamento_dia_by_deal),
        'realizadas': _contar_por_dia(realizadas_ids, _cc1_dia_by_deal),
        'no_show': _contar_por_dia(no_show_ids, _no_show_dia_by_deal),
    }
    meta_agendamentos = leads_n * config.META_AGENDAMENTO_PCT
    # Meta de realizadas usa o agendamento ESPERADO (meta_agendamentos), não o
    # que o SDR já agendou de fato — senão quem agenda pouco "ganha" uma meta
    # de realizadas artificialmente baixa, mascarando o atraso. Pedido do
    # time (09/10/2026): "a meta de realizadas e o que falta deve andar junto
    # com o que ele deveria agendar, não com o que já tem agendado".
    meta_realizadas = meta_agendamentos * config.META_AGENDADOR_REALIZADA_PCT
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
        # leads recebidos, e realizadas até chegar em 65% do que DEVERIA ter sido
        # agendado (meta_agendamentos) — não do que já foi agendado de fato.
        'faltam_agendamentos': max(0, ceil(meta_agendamentos - n_agendamentos)),
        'faltam_realizadas': max(0, ceil(meta_realizadas - realizadas_n)),
        'diario': diario_agendador,
    })
agendadores_meta = {
    'meta_agendamento_pct': round(config.META_AGENDAMENTO_PCT * 100),
    'meta_realizada_pct': round(config.META_AGENDADOR_REALIZADA_PCT * 100),
}


def build_page(*, pipeline_ids, deals_p, stages_contrato_enviado, leads_mode, meta_leads, meta_contratos_val, tem_meta, funil_nome,
                scoreboard_pipeline_ids=None, previsao_todos_funis=False, contrato_enviado_extra=None):
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

    # "Contrato enviado" contado pela etapa do negócio, não por atividade separada.
    # Duas versões, usos diferentes:
    # - contrato_enviado_re (status=='open'): snapshot ATUAL de quem está parado
    #   nessa etapa agora, aguardando assinatura — usado no "Relatório por closer"
    #   (count_by_closer, sem filtro de data). Sem o status=='open' aqui, um
    #   negócio que passou pela etapa e depois foi Ganho/Perdido continuava
    #   contando pra sempre (o stage_id gravado não muda depois que fecha) —
    #   corrigido 29/08/2026.
    # - contrato_enviado_mes_re (qualquer status): usado no ritmo (daily_series/
    #   per_closer_daily, já filtrados por stage_change_time dentro do mês) — todo
    #   negócio Ganho necessariamente passou por "Contrato enviado" antes, então
    #   excluir quem já não está mais 'open' subcontava o ritmo do mês (pedido do
    #   time, 21/10/2026: "todo contrato ganho necessariamente teve contrato
    #   enviado"). O filtro de mês já vem do stage_change_time dentro de
    #   daily_series/per_closer_daily, não precisa do status aqui.
    contrato_enviado_re = pd.DataFrame()
    contrato_enviado_mes_re = pd.DataFrame()
    if not deals_p.empty and 'stage_id' in deals_p.columns:
        _ce = deals_p[deals_p['stage_id'].isin(stages_contrato_enviado) & deals_p['closer_nome'].notna()].copy()
        contrato_enviado_re = _ce[_ce['status'] == 'open']
        contrato_enviado_mes_re = _ce

    # Painel principal (Inbound) soma "Contrato enviado" de TODOS os funis da
    # Fiscal, não só o próprio (pedido do time, 21/10/2026: "todos que passam
    # pela etapa de contrato enviado tem que constar aqui, de qualquer funil") —
    # mesmo espírito do Placar de Ganhos (scoreboard_pipeline_ids), mas aqui
    # precisa de stage_id/stage_change_time por funil (não tem em deals.xlsx),
    # por isso recebe os dataframes por pipeline já carregados em vez de reusar
    # deals_pipe_scoreboard.
    if contrato_enviado_extra:
        extras_atual, extras_mes = [], []
        for _df_extra, _stages_extra in contrato_enviado_extra:
            if _df_extra.empty or 'stage_id' not in _df_extra.columns:
                continue
            _e = _df_extra[_df_extra['stage_id'].isin(_stages_extra)].copy()
            _e['closer_nome'] = _e['owner_id'].map(_closer_nome_por_id)
            _e = _e[_e['closer_nome'].notna()]
            extras_mes.append(_e)
            extras_atual.append(_e[_e['status'] == 'open'])
        if extras_mes:
            contrato_enviado_re = pd.concat([contrato_enviado_re] + extras_atual, ignore_index=True)
            contrato_enviado_mes_re = pd.concat([contrato_enviado_mes_re] + extras_mes, ignore_index=True)

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
            'contrato_enviado': per_closer_daily(contrato_enviado_mes_re, 'stage_change_time', c['nome']),
            'contratos': per_closer_daily(deals_pipe, 'won_time', c['nome'], status='won'),
            'contratos_franquia': contratos_diario_por_tipo(c['nome'], 'franquia'),
            'contratos_partner': contratos_diario_por_tipo(c['nome'], 'partner'),
            'no_show': per_closer_daily(no_show_p, 'marked_as_done_time', c['nome']),
            'venda': per_closer_daily_sum(deals_pipe, 'won_time', 'valor', c['nome'], status='won'),
            'caixa': per_closer_daily_sum(deals_pipe, 'won_time', 'caixa', c['nome'], status='won'),
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
            add_time_dt = pd.to_datetime(d['add_time'], errors='coerce')
            tem_etiqueta = d['label'].apply(_tem_etiqueta)
            # Negócio ANTIGO que uma automação (ex: n8n) reabre/reseta pra etapa
            # inicial e/ou marca com a etiqueta "Lead Studio Fiscal" muito tempo
            # depois de criado não é um lead NOVO — mesmo com add_time dentro do
            # mês, ele nunca tinha contado até esse reset, e reaparece de uma vez
            # só inflando o total retroativamente. Caso real (29/09/2026): 202
            # negócios de agosto/início de setembro, todos resetados pra etapa
            # "LEAD" no mesmo minuto, só então passando no filtro. Lead genuíno
            # tem a troca de etapa bem perto da criação (mesmo dia); gap grande
            # entre add_time e a entrada na etapa atual = reativação, não
            # geração nova.
            if 'stage_change_time' in d.columns:
                stage_change_dt = pd.to_datetime(d['stage_change_time'], errors='coerce')
                gap_dias = (stage_change_dt - add_time_dt).dt.days
                nao_reativado = gap_dias.isna() | (gap_dias.abs() <= 7)
            else:
                nao_reativado = True
            d['add_time'] = add_time_dt.dt.date
            for day in d.loc[tem_etiqueta & nao_reativado, 'add_time'].dropna():
                if config.MONTH_START <= day <= config.MONTH_END:
                    daily_leads[day.isoformat()] = daily_leads.get(day.isoformat(), 0) + 1
        consolidado['leads'] = sum(daily_leads.values())
    else:
        daily_leads = daily_leads_funil
        consolidado['leads'] = consolidado['leads_funil']

    daily_contrato_enviado = daily_series(contrato_enviado_mes_re, 'stage_change_time')

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

    # Metas de CC1/FPQ/CC2/CC3/Contrato enviado estimadas em cascata a partir da
    # meta de leads (config.BENCHMARK_*) — taxas fixas acumuladas de meses
    # anteriores, não uma razão ao vivo do mês corrente (essa ficava instável no
    # início do mês e pior ainda depois que Contrato enviado passou a somar os 4
    # funis da Fiscal contra uma CC3 só do Inbound, 21/10/2026). Sem meta_leads
    # (ex: Outbound ainda sem meta), tudo cascateia pra None.
    meta_cc1_calc = round(meta_leads * config.BENCHMARK_LEADS_PARA_CC1) if meta_leads else None
    meta_fpq_calc = round(meta_cc1_calc * config.BENCHMARK_CC1_PARA_FPQ) if meta_cc1_calc else None
    meta_cc2_calc = round(meta_fpq_calc * config.BENCHMARK_FPQ_PARA_CC2) if meta_fpq_calc else None
    meta_cc3_calc = round(meta_cc2_calc * config.BENCHMARK_CC2_PARA_CC3) if meta_cc2_calc else None
    meta_contrato_enviado_calc = (
        round(meta_cc3_calc * config.BENCHMARK_CC3_PARA_CONTRATO_ENVIADO) if meta_cc3_calc else None
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
    perdidos = {'total': 0, 'pela_qualificadora': 0, 'apos_repasse': 0, 'por_motivo': [], 'detalhe': []}
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
        # Lista negócio a negócio (data + motivo) — só o total do mês não dava pra
        # recortar um período customizado (página Período personalizado,
        # 12/09/2026: "todos os indicadores refletirem o período escolhido").
        perdidos['detalhe'] = [
            {
                'data': r['lost_time'].isoformat() if pd.notna(r['lost_time']) else None,
                'motivo': r['motivo'],
                'pela_qualificadora': bool(r['pela_qualificadora']),
                'closer_nome': None if pd.isna(r.get('closer_nome')) else r.get('closer_nome'),
            }
            for _, r in dl.iterrows()
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

    # Lista bruta (data + utm_campanha + utm_criativo) de leads e CC1 realizada —
    # o ranking acima já vem pronto pro mês inteiro, mas a página Período
    # personalizado precisa recalcular pra um recorte de datas arbitrário
    # (12/09/2026: "todos os indicadores refletirem o período escolhido").
    def _ranking_bruto(df, date_col):
        if df.empty or date_col not in df.columns:
            return []
        d = df.copy()
        d[date_col] = pd.to_datetime(d[date_col], errors='coerce').dt.date
        out = []
        for _, r in d.dropna(subset=[date_col]).iterrows():
            if not (config.MONTH_START <= r[date_col] <= config.MONTH_END):
                continue
            out.append({
                'data': r[date_col].isoformat(),
                'utm_campanha': None if pd.isna(r.get('utm_campanha')) else str(r.get('utm_campanha')).strip() or None,
                'utm_criativo': None if pd.isna(r.get('utm_criativo')) else str(r.get('utm_criativo')).strip() or None,
            })
        return out

    ranking = {
        'campanhas': ranking_utm('utm_campanha'),
        'criativos': ranking_utm('utm_criativo'),
        'detalhe_leads': _ranking_bruto(deals_p, 'add_time'),
        'detalhe_cc1': _ranking_bruto(cc1_p, 'marked_as_done_time'),
    }

    # ============= MAPA DE CALOR — HORÁRIOS DE CC1 REALIZADA (seg-sex, 09h-20h) =============
    HEATMAP_DIAS = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
    HEATMAP_HORAS = list(range(9, 21))
    heatmap_matriz = [[0] * len(HEATMAP_HORAS) for _ in HEATMAP_DIAS]
    heatmap_detalhe = []  # (data, hora) de cada CC1 — pra "Ver o dia" recalcular só daquele dia
    if not cc1_p.empty and 'marked_as_done_time' in cc1_p.columns:
        hm = cc1_p.copy()
        hm['marked_as_done_time'] = pd.to_datetime(hm['marked_as_done_time'], errors='coerce')
        for _, r in hm.dropna(subset=['marked_as_done_time']).iterrows():
            wd = r['marked_as_done_time'].weekday()
            hora = r['marked_as_done_time'].hour
            if hora in HEATMAP_HORAS:
                heatmap_detalhe.append({'data': r['marked_as_done_time'].date().isoformat(), 'hora': hora})
            if wd > 4:
                continue
            if hora in HEATMAP_HORAS:
                heatmap_matriz[wd][HEATMAP_HORAS.index(hora)] += 1
    heatmap_cc1 = {'dias': HEATMAP_DIAS, 'horas': HEATMAP_HORAS, 'matriz': heatmap_matriz, 'detalhe': heatmap_detalhe}

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

    def _prev_perdido(r):
        # "Perdido" = deu lost no Pipedrive OU o time marcou PERDIDO na coluna
        # Confiança da planilha (às vezes o time marca antes do Pipedrive, ou o
        # negócio nem tem link válido) — pedido do time (10/09/2026).
        return r.get('deal_status') == 'lost' or 'perdid' in str(r.get('confianca') or '').lower()

    def _prev_ordem(r):
        # Três blocos, nessa ordem: Ganhos, depois Perdidos, depois o resto
        # (pipeline vivo) — dentro de cada bloco, por data de fechamento
        # esperada do menor pro maior (sem data vai pro fim do bloco). Pedido
        # do time (28/09/2026, revertendo a tentativa anterior de misturar
        # Ganho+Perdido num bloco só).
        grupo = 0 if r.get('deal_status') == 'won' else (1 if _prev_perdido(r) else 2)
        return (grupo, r.get('data_fechamento_esperada') or '9999-12-31')

    previsao_funil = sorted(previsao_funil, key=_prev_ordem)

    def _prev_proximo_mes(r):
        # Negócio com data de fechamento esperada JÁ EMPURRADA pro mês
        # seguinte não deveria inflar o "previsto" deste mês (o time claramente
        # não espera mais fechar isso agora) — pedido do time (28/09/2026,
        # casos reais: Bruna Bressani e Cyro Neto com fechamento esperado em
        # outubro, aparecendo somados no total de setembro).
        d = r.get('data_fechamento_esperada')
        return bool(d) and str(d) > config.MONTH_END.isoformat()

    # Negócio Ganho OU Perdido não conta no "previsto" (venda/caixa/contratos) —
    # um negócio morto não é mais uma previsão de venda futura (09-10/09/2026,
    # casos reais: Nilza, Alberto Moron, Daniel Thiago). Postergado pro mês
    # seguinte também sai do "previsto" deste mês (vira sua própria lista,
    # 'previsao_postergada', só pra mostrar/somar à parte).
    previsao_pendente = [
        r for r in previsao_funil
        if r.get('deal_status') != 'won' and not _prev_perdido(r) and not _prev_proximo_mes(r)
    ]
    previsao_postergada = [
        r for r in previsao_funil
        if r.get('deal_status') != 'won' and not _prev_perdido(r) and _prev_proximo_mes(r)
    ]
    # Previsto por TIPO de contrato (21/09/2026): o alvo principal é Franquia, não
    # dá pra somar Franquia com Partner (de graça) num número só. "Pendente" =
    # ainda em aberto na planilha (não Ganho, não Perdido); linha sem tipo
    # (sem produto nem Modelo) conta como Franquia. O front soma com o
    # Realizado de cada tipo pra montar "Previsto" e "faltam".
    _pend_partner = [r for r in previsao_pendente if r.get('tipo_contrato') == 'partner']
    scoreboard_previsto = {
        'contratos_previstos_franquia': len(previsao_pendente) - len(_pend_partner),
        'contratos_previstos_partner': len(_pend_partner),
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

    # Negócios fechados no mês, um por um (16/10/2026: "quero ver quem são,
    # quem foram") — mesmo escopo do Placar (dw), com dono, funil, tipo e link.
    ganhos = []
    if not dw.empty:
        for _, r in dw.sort_values('won_time').iterrows():
            def _v(x):
                return None if pd.isna(x) else x
            ganhos.append({
                'deal_id': int(r['deal_id']) if pd.notna(r.get('deal_id')) else None,
                'titulo': _v(r.get('titulo')),
                'pessoa_nome': _v(r.get('pessoa_nome')),
                'closer_nome': _v(r.get('closer_nome')),
                'won_time': r['won_time'].strftime('%Y-%m-%d %H:%M') if pd.notna(r.get('won_time')) else None,
                'tipo_contrato': _v(r.get('tipo_contrato')),
                'valor': float(r['valor']) if pd.notna(r.get('valor')) else 0.0,
                'caixa': float(r['caixa']) if pd.notna(r.get('caixa')) else 0.0,
                'funil': config.FUNIL_POR_PIPELINE.get(int(r['pipeline_id'])) if pd.notna(r.get('pipeline_id')) else None,
                'deal_link': f"https://{config.PIPEDRIVE_DOMAIN}/deal/{int(r['deal_id'])}" if pd.notna(r.get('deal_id')) else None,
            })
    _escopo_multi = set(_scoreboard_pipeline_ids) != set(pipeline_ids)
    placar_hint = (
        '— Realizado direto do Pipedrive somando todos os funis da Fiscal (Inbound, Rebarba e Outbound); Previsto vem da planilha de previsão'
        if _escopo_multi else
        '— Realizado direto do Pipedrive, só deste funil (mesmo escopo do Metas do mês); Previsto vem da planilha de previsão'
    )
    metas_hint = (
        '— contratos aqui contam só os Ganhos deste funil com dono entre os closers atuais; o Placar acima soma também Rebarba e Outbound'
        if _escopo_multi else
        '— contratos aqui contam só os Ganhos deste funil com dono entre os closers atuais (mesmo número do Placar do mês)'
    )

    return {
        'ganhos': ganhos,
        'placar_hint': placar_hint,
        'metas_hint': metas_hint,
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
        # Fenalaw 2026 (pipeline 72) — compartilhado com a Agro, só conta aqui
        # quando um dos 4 closers da Fiscal é o dono do negócio (ver deals.xlsx,
        # que já cobre qualquer pipeline que um closer assuma); pedido do time
        # (30/10/2026).
        config.PIPELINE_FENALAW_2026,
    ],
    previsao_todos_funis=True,
    contrato_enviado_extra=[
        (deals_p65, [config.PIPELINE_STAGE_CONTRATO_ENVIADO_REBARBA]),
        (deals_p66, [config.PIPELINE_STAGE_CONTRATO_ENVIADO_OUTBOUND]),
        (deals_p69, [config.PIPELINE_STAGE_CONTRATO_ENVIADO_EVENTOS]),
        (deals_p72, [config.PIPELINE_STAGE_CONTRATO_ENVIADO_FENALAW]),
    ],
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
    # Meta combinada (Circular 13/2026 + acordo com Adalberto, 05/10/2026): 1
    # Franquia + 10 Partner = 11 contratos/mês — só pra popular o card
    # "Contratos fechados" do Ritmo (não ativa "Metas do mês", que depende de
    # meta_leads).
    meta_contratos_val=config.META_CONTRATOS_OUTBOUND_FRANQUIA + config.META_CONTRATOS_OUTBOUND_PARTNER,
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
# Fenalaw 2026 (72) — pipeline compartilhado com a Studio Agro (evento único);
# aqui só entra pela ótica da Fiscal, com os closers da Fiscal como dono (os
# mesmos negócios já contam no Placar/Contrato enviado do Inbound via
# scoreboard_pipeline_ids/contrato_enviado_extra acima). Ganha seção própria
# no painel unificado a pedido do time (30/10/2026), ainda sem meta.
fenalaw_data = build_page(
    pipeline_ids=[config.PIPELINE_FENALAW_2026],
    deals_p=deals_p72,
    stages_contrato_enviado=[config.PIPELINE_STAGE_CONTRATO_ENVIADO_FENALAW],
    leads_mode='criado',
    meta_leads=None,
    meta_contratos_val=None,
    tem_meta=False,
    funil_nome='FENALAW',
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
    'inbound': {'url': 'index.html', 'label': 'Studio Fiscal (painel unificado)'},
    'outbound': {'url': 'outbound.html', 'label': 'Outbound'},
    'rebarba': {'url': 'rebarba.html', 'label': 'Rebarba (reativação)'},
}


def _paginas_irmas(exceto):
    return [v for k, v in PAGINAS_IRMAS.items() if k != exceto]


# Métricas padrão do "Relatório por closer" quando o funil não tem CC1 como
# porta de entrada (Outbound/Rebarba/Fenalaw — cada closer já entra com o
# próprio lead, "Leads" é que mede produtividade aqui, não "Negócios
# iniciados" como no Inbound).
_RELATORIO_METRICS_PADRAO = [
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
]

# ============= Dict de dados de CADA funil (seção do painel unificado E,
# pra Outbound/Rebarba, também a página single-funil própria) — "dados_X"
# nunca ganha a chave 'funis' nele mesmo (só os dicts de PÁGINA, montados
# depois, ganham isso), pra não criar referência circular no JSON. =============
dados_inbound = {
    **common,
    **inbound_data,
    'agendadores': agendadores_out,
    'agendadores_meta': agendadores_meta,
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
dados_outbound = {
    **common,
    **outbound_data,
    'perdidos_hint': '— negócios perdidos nos funis Outbound + Eventos da Studio Fiscal no mês',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. Junta o funil Outbound próprio (66) '
        'com o Outbound - Eventos (69). "Leads" conta negócios criados no mês já com '
        'um closer como dono'
    ),
    'relatorio_metrics': _RELATORIO_METRICS_PADRAO,
}
dados_rebarba = {
    **common,
    **rebarba_data,
    'perdidos_hint': '— negócios perdidos no funil Rebarba Studio Fiscal Franquia no mês',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. Reativação de negócios perdidos do '
        'funil Studio Fiscal Franquia. "Leads" conta negócios criados no mês já com '
        'um closer como dono'
    ),
    'relatorio_metrics': _RELATORIO_METRICS_PADRAO,
}
dados_fenalaw = {
    **common,
    **fenalaw_data,
    'perdidos_hint': '— negócios perdidos no pipeline Fenalaw 2026 no mês, entre os negócios já com dono entre os closers da Fiscal',
    'relatorio_hint': (
        '— diário, semanal e consolidado do mês. Pipeline compartilhado com a Studio '
        'Agro (evento Fenalaw 2026) — só entra aqui quando o dono é um dos closers da '
        'Fiscal. "Leads" conta negócios criados no mês já com um closer como dono'
    ),
    'relatorio_metrics': _RELATORIO_METRICS_PADRAO,
}

dashboard_outbound = {
    **dados_outbound,
    'pagina_titulo': 'Equipe de Vendas — Outbound',
    'paginas_irmas': _paginas_irmas('outbound'),
    'historico': HISTORICO_OUTBOUND,
    'mes_outro_url': MES_ANTERIOR_ARQUIVO_OUTBOUND,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'funis': [{'sufixo': 'outbound', 'cor': '#38bdf8', 'titulo': 'OUTBOUND', 'badge': '+ Eventos', 'dados': dados_outbound}],
}
dashboard_rebarba = {
    **dados_rebarba,
    'pagina_titulo': 'Equipe de Vendas — Rebarba (reativação)',
    'paginas_irmas': _paginas_irmas('rebarba'),
    'historico': HISTORICO_REBARBA,
    'mes_outro_url': MES_ANTERIOR_ARQUIVO_REBARBA,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'funis': [{'sufixo': 'rebarba', 'cor': '#22c55e', 'titulo': 'REBARBA', 'badge': 'reativação', 'dados': dados_rebarba}],
}

# ============= Painel unificado (index.html) — as 4 seções (Inbound,
# Outbound, Rebarba, Fenalaw) empilhadas numa página só, cada uma com seu
# próprio card completo (Placar/Metas/Ritmo/Consolidado/Perdidos/
# Comparativo/Ranking/Relatório/Previsão/Mapa de calor/FPQs), em vez de 3
# páginas separadas. Decidido com o time via mockup (Opção B — "tudo
# empilhado, uma seção por funil"), 03/11/2026. =============
dashboard_unificada = {
    **common,
    'pagina_titulo': 'Equipe de Vendas — Studio Fiscal (painel unificado)',
    # Não tem mais "página irmã" pra linkar — o ponto é não precisar navegar
    # entre páginas pra ver os 4 funis. outbound.html/rebarba.html continuam
    # existindo (linkadas lá, não daqui) pra quem tiver o link salvo.
    'paginas_irmas': [],
    'historico': HISTORICO_INBOUND,
    'mes_outro_url': MES_ANTERIOR_ARQUIVO,
    'mes_outro_label': MES_ANTERIOR_LABEL,
    'funis': [
        {'sufixo': 'inbound', 'cor': '#a855f7', 'titulo': 'INBOUND', 'badge': 'painel principal', 'dados': dados_inbound},
        {'sufixo': 'outbound', 'cor': '#38bdf8', 'titulo': 'OUTBOUND', 'badge': '+ Eventos', 'dados': dados_outbound},
        {'sufixo': 'rebarba', 'cor': '#22c55e', 'titulo': 'REBARBA', 'badge': 'reativação', 'dados': dados_rebarba},
        {'sufixo': 'fenalaw', 'cor': '#f97316', 'titulo': 'FENALAW 2026', 'badge': 'novo · compartilhado com Agro', 'dados': dados_fenalaw},
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


render(dashboard_unificada, 'index.html')
render(dashboard_outbound, 'outbound.html')
render(dashboard_rebarba, 'rebarba.html')

# Snapshot do mês corrente (ver comentário acima, "ARQUIVAMENTO AUTOMÁTICO DE MÊS
# FECHADO") — mesmo conteúdo do index.html/outbound.html/rebarba.html, só troca o
# link de navegação: em vez de apontar pro mês anterior, aponta de volta pro mês
# atual (genérico, sem citar o mês por nome, pra não ficar desatualizado quando esse
# arquivo congelar e passar a ser visto meses depois).
render({**dashboard_unificada, 'mes_outro_url': 'index.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL)
render({**dashboard_outbound, 'mes_outro_url': 'outbound.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL_OUTBOUND)
render({**dashboard_rebarba, 'mes_outro_url': 'rebarba.html', 'mes_outro_label': '← Ver mês atual'}, ARQUIVO_MES_ATUAL_REBARBA)

ci = dados_inbound['consolidado']
print(f"Inbound (Studio Fiscal) — CC1={ci['cc1_realizada']} | FPQ={ci['fpq_realizada']} | "
      f"CC2={ci['cc2_realizada']} | CC3={ci['cc3_realizada']} | "
      f"Contratos={ci['contratos']}/{metas['meta_contratos']}")
co = dados_outbound['consolidado']
print(f"Outbound — CC1={co['cc1_realizada']} | FPQ={co['fpq_realizada']} | "
      f"CC2={co['cc2_realizada']} | CC3={co['cc3_realizada']} | "
      f"Leads criados no mês={co['leads_atribuidos_closers']} | Contratos={co['contratos']}")
cf = dados_fenalaw['consolidado']
print(f"Fenalaw 2026 — CC1={cf['cc1_realizada']} | FPQ={cf['fpq_realizada']} | "
      f"CC2={cf['cc2_realizada']} | CC3={cf['cc3_realizada']} | "
      f"Leads criados no mês={cf['leads_atribuidos_closers']} | Contratos={cf['contratos']}")
