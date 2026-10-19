"""
Busca dados do Pipedrive via API para o dashboard Studio Fiscal (equipe de closers).

Conta CC1/CC2/CC3/FPQ e leads/contratos SEMPRE pelo responsável da atividade/negócio
(um dos 4 closers), qualquer funil — o funil STUDIO FISCAL FRANQUIA ainda está em
transição, então não dá pra filtrar só por ele.

Gera em /data_pipedrive/:
- activities_cc1_realizada.xlsx / activities_cc1_futura.xlsx
- activities_cc2_realizada.xlsx / activities_cc2_futura.xlsx
- activities_cc3_realizada.xlsx
- activities_fpq_realizada.xlsx
- deals.xlsx  — negócios com um dos 4 closers como proprietário (p/ leads e contratos)
"""
import urllib.request, urllib.parse, urllib.error, json, os, time
from datetime import timedelta
import pandas as pd
import concurrent.futures

import config

WORKDIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(WORKDIR, '.pipedrive_token')
OUT_DIR = os.path.join(WORKDIR, 'data_pipedrive')
os.makedirs(OUT_DIR, exist_ok=True)

TOKEN = os.environ.get('PIPEDRIVE_TOKEN') or (
    open(TOKEN_FILE).read().strip() if os.path.exists(TOKEN_FILE) else None
)
if not TOKEN:
    raise RuntimeError('PIPEDRIVE_TOKEN não encontrado (env var ou .pipedrive_token)')

BASE = 'https://api.pipedrive.com/v1'


def fetch_all(endpoint, params=None):
    """Itera paginação até pegar todos os resultados, com retry em rate-limit/erro."""
    params = dict(params or {})
    params['api_token'] = TOKEN
    params['limit'] = 500
    start = 0
    all_data = []
    while True:
        params['start'] = start
        qs = urllib.parse.urlencode(params, doseq=True)
        url = f'{BASE}/{endpoint}?{qs}'
        body = None
        for attempt in range(10):
            try:
                with urllib.request.urlopen(url, timeout=120) as r:
                    body = json.loads(r.read())
                break
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    wait = 15 * (attempt + 1)
                    print(f'     rate limit (429), esperando {wait}s...')
                    time.sleep(wait)
                elif attempt == 9:
                    raise
                else:
                    time.sleep(3)
            except Exception:
                if attempt == 9:
                    raise
                time.sleep(3)
        if body is None:
            raise RuntimeError(f'Falha ao buscar {endpoint} após 10 tentativas')
        data = body.get('data') or []
        all_data.extend(data)
        pag = body.get('additional_data', {}).get('pagination', {})
        if not pag.get('more_items_in_collection'):
            break
        start = pag.get('next_start', start + len(data))
        if len(data) == 0:
            break
    return all_data


def _in_range(dt_str, start, end):
    if not dt_str:
        return False
    try:
        d = pd.to_datetime(dt_str).date()
        return start <= d <= end
    except Exception:
        return False


# A conta Pipedrive do Grupo Studio devolve TODO campo de data/hora 3h à frente do
# horário real de Brasília (confirmado comparando due_time/marked_as_done_time de
# atividades reais contra o relatório oficial de Insights do Pipedrive p/ as mesmas
# atividades — provavelmente timezone da conta mal configurado). Corrige subtraindo
# 3h de tudo, direto na busca, pra todo o resto do pipeline já trabalhar com o
# horário certo.
TZ_OFFSET = timedelta(hours=3)


def _fix_dt(dt_str):
    """Corrige um datetime completo ('2026-07-07 19:47:54') -3h."""
    if not dt_str:
        return dt_str
    try:
        return (pd.to_datetime(dt_str) - TZ_OFFSET).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return dt_str


def _fix_due(due_date, due_time):
    """Corrige o par due_date+due_time (-3h), com possível virada de dia."""
    if not due_date:
        return due_date, due_time
    try:
        combined = pd.to_datetime(f"{due_date} {due_time or '00:00:00'}") - TZ_OFFSET
        return combined.strftime('%Y-%m-%d'), (combined.strftime('%H:%M') if due_time else due_time)
    except Exception:
        return due_date, due_time


def _correct_activity_tz(a):
    a = dict(a)
    a['due_date'], a['due_time'] = _fix_due(a.get('due_date'), a.get('due_time'))
    a['marked_as_done_time'] = _fix_dt(a.get('marked_as_done_time'))
    return a


CLOSER_NAME = {c['id']: c['nome'] for c in config.CLOSERS}

# ============= ACTIVITIES =============
# (activity_type, tem_para_acontecer)
# "Contrato enviado" não entra aqui — é contado pela etapa do negócio (ver config.py
# PIPELINE_STAGE_CONTRATO_ENVIADO), o time não registrava a atividade separada de forma
# confiável.
ACT_STAGES = [
    ('cc1', config.ACT_CC1, True),
    ('cc2', config.ACT_CC2, True),
    ('cc3', config.ACT_CC3, True),
    ('fpq', config.ACT_FPQ, False),
]


def _fetch_activities(stage_key, act_type, closer_id):
    rows = fetch_all('activities', {
        'type': act_type,
        'user_id': closer_id,
        'start_date': config.FETCH_START.isoformat(),
        'end_date': config.FETCH_END.isoformat(),
    })
    return stage_key, closer_id, rows


def act_row(a):
    return {
        'closer_id': a.get('user_id'),
        'closer_nome': CLOSER_NAME.get(a.get('user_id')),
        'assunto': a.get('subject'),
        'negocio_id': a.get('deal_id'),
        'due_date': a.get('due_date'),
        'due_time': a.get('due_time'),
        'marked_as_done_time': a.get('marked_as_done_time'),
        'done': a.get('done'),
    }


print('-> Fetching activities (CC1/CC2/CC3/FPQ) por closer, em paralelo...')
tasks = [(stage_key, act_type, cid) for stage_key, act_type, _ in ACT_STAGES for cid in config.CLOSER_IDS]
acts_by_stage = {stage_key: [] for stage_key, _, _ in ACT_STAGES}
with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
    futures = [executor.submit(_fetch_activities, stage_key, act_type, cid) for stage_key, act_type, cid in tasks]
    for f in concurrent.futures.as_completed(futures):
        stage_key, closer_id, rows = f.result()
        acts_by_stage[stage_key].extend(rows)
        print(f'   {stage_key} / {CLOSER_NAME.get(closer_id)}: {len(rows)}')

cc1_realizada_rows = []
all_realizada = {}
all_futura = {}
for stage_key, act_type, tem_futura in ACT_STAGES:
    raw = [_correct_activity_tz(a) for a in acts_by_stage[stage_key]]
    done_no_mes = [a for a in raw
                   if a.get('done') and _in_range(a.get('marked_as_done_time'), config.MONTH_START, config.MONTH_END)]
    if stage_key == 'fpq':
        # FPQ recebida: guarda pessoa/negócio + link direto pro Pipedrive, pra dar
        # pra abrir e ver a FPQ (fica registrada como nota da atividade lá).
        realizada = []
        for a in done_no_mes:
            row = act_row(a)
            row['pessoa_nome'] = a.get('person_name')
            row['negocio_titulo'] = a.get('deal_title')
            deal_id = a.get('deal_id')
            row['deal_link'] = f'https://{config.PIPEDRIVE_DOMAIN}/deal/{deal_id}' if deal_id else None
            realizada.append(row)
    else:
        realizada = [act_row(a) for a in done_no_mes]
    all_realizada[stage_key] = realizada
    if stage_key == 'cc1':
        cc1_realizada_rows = realizada
    print(f'   {stage_key} realizada: {len(realizada)}')
    if tem_futura:
        futura = [act_row(a) for a in raw
                  if not a.get('done') and _in_range(a.get('due_date'), config.MONTH_START, config.MONTH_END)]
        all_futura[stage_key] = futura
        print(f'   {stage_key} para acontecer: {len(futura)}')

# ============= NO-SHOW (tarefa do CLOSER — ver config.py) =============
print('-> Fetching no-show (imprevisto/desinteresse) por closer...')


def _fetch_no_show(act_type, closer_id):
    rows = fetch_all('activities', {
        'type': act_type,
        'user_id': closer_id,
        'start_date': config.FETCH_START.isoformat(),
        'end_date': config.FETCH_END.isoformat(),
    })
    return act_type, closer_id, rows


no_show_tasks = [
    (act_type, cid)
    for act_type in (config.ACT_NO_SHOW_IMPREVISTO, config.ACT_NO_SHOW_DESINTERESSE)
    for cid in config.CLOSER_IDS
]
no_show_raw = {act_type: [] for act_type in (config.ACT_NO_SHOW_IMPREVISTO, config.ACT_NO_SHOW_DESINTERESSE)}
with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
    futures = [executor.submit(_fetch_no_show, act_type, cid) for act_type, cid in no_show_tasks]
    for f in concurrent.futures.as_completed(futures):
        act_type, closer_id, rows = f.result()
        no_show_raw[act_type].extend(rows)

no_show_rows = []
for act_type, rows in no_show_raw.items():
    for a in rows:
        a = _correct_activity_tz(a)
        if a.get('done') and _in_range(a.get('marked_as_done_time'), config.MONTH_START, config.MONTH_END):
            row = act_row(a)
            row['tipo'] = act_type
            no_show_rows.append(row)
print(f'   no-show no mês (todos closers): {len(no_show_rows)}')

# ============= LEADS QUALIFICADOS (Qualificação Realizada) =============
# Etapa anterior à CC1 no "Consolidado do time" (31/08/2026) — feita pelo time de
# qualificação (SDRs, ~12 pessoas, não só os 3 agendadores da Fiscal), não por um
# dos 4 closers. user_id=0 é o valor que a API do Pipedrive usa pra "todo mundo" —
# sem ele o endpoint /activities só devolve as atividades do dono do token
# (Mauricio), por isso vinha sempre 0 (31/08/2026).
print('-> Fetching Qualificação Realizada (time de qualificação inteiro, user_id=0)...')
qualificacao_raw = fetch_all('activities', {
    'type': config.ACT_QUALIFICACAO_REALIZADA,
    'user_id': 0,
    'start_date': config.FETCH_START.isoformat(),
    'end_date': config.FETCH_END.isoformat(),
})
qualificacao_rows = []
for a in qualificacao_raw:
    a = _correct_activity_tz(a)
    if a.get('done') and _in_range(a.get('marked_as_done_time'), config.MONTH_START, config.MONTH_END):
        qualificacao_rows.append(act_row(a))
print(f'   qualificação realizada no mês: {len(qualificacao_rows)}')

# ============= PIPELINE (e UTM da CC1) do negócio vinculado =============
# Enriquece TODA atividade (CC1/CC2/CC3/FPQ, realizada e futura) com o pipeline_id
# do negócio vinculado — pra dar pra separar Studio Fiscal (62) de Outbound (66)
# sem misturar os números. CC1 realizada também ganha campanha/criativo (UTM) pra
# ranking. Um único batch deduplicado pra não repetir fetch do mesmo negócio.
print('-> Enriquecendo atividades com pipeline (e UTM da CC1) do negócio vinculado...')


def _fetch_deal_utm(deal_id):
    url = f'{BASE}/deals/{deal_id}?api_token={TOKEN}'
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=30) as r:
                body = json.loads(r.read())
            d = body.get('data') or {}
            return deal_id, d.get(config.CF_UTM_CAMPAIGN), d.get(config.CF_UTM_TERM), d.get('pipeline_id')
        except Exception:
            if attempt == 3:
                return deal_id, None, None, None
            time.sleep(2 * (attempt + 1))
    return deal_id, None, None, None


all_rows = [r for rows in all_realizada.values() for r in rows] + [r for rows in all_futura.values() for r in rows] + no_show_rows + qualificacao_rows
unique_deal_ids = sorted({r['negocio_id'] for r in all_rows if r.get('negocio_id')})
utm_by_deal = {}
if unique_deal_ids:
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        for deal_id, campanha, criativo, pipeline_id in executor.map(_fetch_deal_utm, unique_deal_ids):
            utm_by_deal[deal_id] = (campanha, criativo, pipeline_id)

for r in all_rows:
    _, _, pipeline_id = utm_by_deal.get(r.get('negocio_id'), (None, None, None))
    r['pipeline_id'] = pipeline_id
for r in cc1_realizada_rows:
    campanha, criativo, _ = utm_by_deal.get(r.get('negocio_id'), (None, None, None))
    r['utm_campanha'] = campanha
    r['utm_criativo'] = criativo

for stage_key, act_type, tem_futura in ACT_STAGES:
    if stage_key != 'cc1':
        pd.DataFrame(all_realizada[stage_key]).to_excel(os.path.join(OUT_DIR, f'activities_{stage_key}_realizada.xlsx'), index=False)
    if tem_futura:
        pd.DataFrame(all_futura[stage_key]).to_excel(os.path.join(OUT_DIR, f'activities_{stage_key}_futura.xlsx'), index=False)
pd.DataFrame(cc1_realizada_rows).to_excel(os.path.join(OUT_DIR, 'activities_cc1_realizada.xlsx'), index=False)
pd.DataFrame(no_show_rows).to_excel(os.path.join(OUT_DIR, 'activities_no_show.xlsx'), index=False)
pd.DataFrame(qualificacao_rows).to_excel(os.path.join(OUT_DIR, 'activities_qualificacao_realizada.xlsx'), index=False)

# ============= AGENDADORES (leads recebidos / agendamentos / realizadas) =============
# "Leads recebidos" = 1ª tentativa de contato criada pro agendador (due_date = quando
# o lead caiu pra ele). "Agendamento" = CC1 marcada como agendada por ele (manual ou
# bot), concluída (done) no mês. "Realizada" cruza o negocio_id do agendamento com os
# CC1 Realizada já buscados acima (activities_cc1_realizada, cc1_realizada_rows) — a
# CC1 em si fica na agenda do closer, não do agendador.
print('-> Fetching atividades dos agendadores (leads recebidos / agendamentos)...')
AGENDADOR_NOME = {a['id']: a['nome'] for a in config.AGENDADORES}


def _fetch_agendador_activities(act_type, agendador_id):
    rows = fetch_all('activities', {
        'type': act_type,
        'user_id': agendador_id,
        'start_date': config.FETCH_START.isoformat(),
        'end_date': config.FETCH_END.isoformat(),
    })
    return act_type, agendador_id, rows


agendador_tasks = [
    (act_type, aid)
    for act_type in (config.ACT_TENTATIVA_1, config.ACT_AGENDAMENTO_CC1, config.ACT_AGENDAMENTO_CC1_BOT)
    for aid in config.AGENDADOR_IDS
]
agendador_raw = {act_type: [] for act_type in (config.ACT_TENTATIVA_1, config.ACT_AGENDAMENTO_CC1, config.ACT_AGENDAMENTO_CC1_BOT)}
with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
    futures = [executor.submit(_fetch_agendador_activities, act_type, aid) for act_type, aid in agendador_tasks]
    for f in concurrent.futures.as_completed(futures):
        act_type, agendador_id, rows = f.result()
        agendador_raw[act_type].extend(rows)
        print(f'   {act_type} / {AGENDADOR_NOME.get(agendador_id)}: {len(rows)}')

tentativa1_rows = []
for a in agendador_raw[config.ACT_TENTATIVA_1]:
    a = _correct_activity_tz(a)
    due = a.get('due_date')
    if due and _in_range(due, config.MONTH_START, config.MONTH_END):
        tentativa1_rows.append({
            'agendador_id': a.get('user_id'),
            'agendador_nome': AGENDADOR_NOME.get(a.get('user_id')),
            'negocio_id': a.get('deal_id'),
            'due_date': due,
        })
pd.DataFrame(tentativa1_rows).to_excel(os.path.join(OUT_DIR, 'agendadores_leads_recebidos.xlsx'), index=False)

agendamento_rows = []
for act_type in (config.ACT_AGENDAMENTO_CC1, config.ACT_AGENDAMENTO_CC1_BOT):
    for a in agendador_raw[act_type]:
        a = _correct_activity_tz(a)
        if a.get('done') and _in_range(a.get('marked_as_done_time'), config.MONTH_START, config.MONTH_END):
            agendamento_rows.append({
                'agendador_id': a.get('user_id'),
                'agendador_nome': AGENDADOR_NOME.get(a.get('user_id')),
                'negocio_id': a.get('deal_id'),
                'tipo': act_type,
                'marked_as_done_time': a.get('marked_as_done_time'),
            })
pd.DataFrame(agendamento_rows).to_excel(os.path.join(OUT_DIR, 'agendadores_agendamentos.xlsx'), index=False)
print(f'   leads recebidos (mês, todos agendadores): {len(tentativa1_rows)} | agendamentos concluídos: {len(agendamento_rows)}')

# ============= DEALS (leads + contratos) =============
print('-> Fetching deals por closer (proprietário)...')


def _fetch_deals(closer_id):
    rows = fetch_all('deals', {'user_id': closer_id})
    return closer_id, rows


deals_rows = []
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
    futures = [executor.submit(_fetch_deals, cid) for cid in config.CLOSER_IDS]
    for f in concurrent.futures.as_completed(futures):
        closer_id, rows = f.result()
        print(f'   {CLOSER_NAME.get(closer_id)}: {len(rows)} deals')
        for d in rows:
            pessoa = d.get('person_id')
            deals_rows.append({
                'deal_id': d.get('id'),
                'titulo': d.get('title'),
                'closer_id': closer_id,
                'closer_nome': CLOSER_NAME.get(closer_id),
                'pipeline_id': d.get('pipeline_id'),
                'status': d.get('status'),
                'valor': d.get('value'),
                'caixa': d.get(config.CF_CAIXA),
                'add_time': _fix_dt(d.get('add_time')),
                'won_time': _fix_dt(d.get('won_time')),
                # Campos usados só pelo sync automático da planilha de Previsão
                # (sync_previsao_auto_add.py, 09/09/2026) — "Data de fechamento
                # esperada" é campo nativo do negócio; "pessoa_nome" é o contato
                # vinculado (o que a planilha chama de "Lead"), vem aninhado no
                # payload do negócio (person_id.name), não achatado pela API v1.
                'expected_close_date': d.get('expected_close_date'),
                'pessoa_nome': pessoa.get('name') if isinstance(pessoa, dict) else None,
                'stage_id': d.get('stage_id'),
                # Probabilidade (0-100) que o Pipedrive calcula/o closer ajusta no
                # negócio — usada pelo sync da planilha de Previsão pra definir a
                # Confiança automaticamente (>80% Quente, 50-80% Morno, <=50% Frio;
                # pedido do time, 08/10/2026).
                'probability': d.get('probability'),
            })

df_deals = pd.DataFrame(deals_rows)
df_deals.to_excel(os.path.join(OUT_DIR, 'deals.xlsx'), index=False)

# ============= DEALS DOS AGENDADORES (leads recebidos, TODO pipeline) =============
# "Leads recebidos" do agendador conta negócio dele em QUALQUER funil, não só o 62
# (Studio Fiscal Franquia) — o time confirmou que o SDR pode receber lead que ainda
# não foi roteado pro funil da Fiscal (ex: ainda na FRANQUIA/42 compartilhada) e
# isso tem que contar igual (03/09/2026). Só os campos que interessam pra contar.
print('-> Fetching deals dos agendadores (todo pipeline)...')


def _fetch_deals_agendador(agendador_id):
    rows = fetch_all('deals', {'user_id': agendador_id})
    return agendador_id, rows


agendadores_deals_rows = []
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
    futures = [executor.submit(_fetch_deals_agendador, aid) for aid in config.AGENDADOR_IDS]
    for f in concurrent.futures.as_completed(futures):
        agendador_id, rows = f.result()
        print(f'   {AGENDADOR_NOME.get(agendador_id)}: {len(rows)} deals (todo pipeline)')
        for d in rows:
            agendadores_deals_rows.append({
                'deal_id': d.get('id'),
                'agendador_id': agendador_id,
                'agendador_nome': AGENDADOR_NOME.get(agendador_id),
                'pipeline_id': d.get('pipeline_id'),
                'add_time': _fix_dt(d.get('add_time')),
                # Usado só pra descartar lead REATRIBUÍDO de outro SDR (o dono
                # mudou muito tempo depois de criado) da contagem de "leads
                # recebidos" do dono atual — ver build_dashboard.py.
                'update_time': _fix_dt(d.get('update_time')),
            })
pd.DataFrame(agendadores_deals_rows).to_excel(os.path.join(OUT_DIR, 'deals_agendadores.xlsx'), index=False)
n_leads = df_deals['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df_deals) else 0
won_no_mes_mask = (df_deals['status'] == 'won') & df_deals['won_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)) if len(df_deals) else pd.Series(dtype=bool)
n_contratos = won_no_mes_mask.sum() if len(df_deals) else 0
print(f'   leads no mês: {n_leads} | contratos (Ganho) no mês: {n_contratos}')

# ============= PRODUTOS dos contratos (Franquia com valor x Partner de graça) =============
# Só busca produto pro que já é contrato fechado no mês (29/08/2026) — o time quer
# separar, no relatório, quantos contratos são de Franquia (produto com valor) e
# quantos são "Contrato Partner" (produto de parceria, sem valor de aquisição).
print('-> Fetching produtos dos contratos fechados no mês...')


def _fetch_deal_products(deal_id):
    url = f'{BASE}/deals/{deal_id}/products?api_token={TOKEN}'
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=30) as r:
                body = json.loads(r.read())
            return deal_id, body.get('data') or []
        except Exception:
            if attempt == 3:
                return deal_id, []
            time.sleep(2 * (attempt + 1))
    return deal_id, []


def _classifica_produtos(produtos):
    nomes = [p.get('name') or '' for p in produtos]
    if any(n.startswith('Franquia') for n in nomes):
        return 'franquia'
    if any(n == 'Contrato Partner' for n in nomes):
        return 'partner'
    return None


# Fallback pro Modelo da planilha de Previsão quando o negócio não tem produto
# vinculado no Pipedrive ainda — confirmado com o time (31/08/2026): "GS Partner"
# é um modelo de FRANQUIA (tem valor), só "Partner" sozinho (sem "GS" na frente) é
# o modelo de parceria (de graça). Só funciona se fetch_previsao.py já rodou antes
# nesse ciclo (previsao.xlsx) — se não existir ainda, o fallback fica vazio.
_modelo_por_deal = {}
_previsao_path = os.path.join(OUT_DIR, 'previsao.xlsx')
if os.path.exists(_previsao_path):
    _prev = pd.read_excel(_previsao_path)
    if 'deal_id' in _prev.columns and 'modelo' in _prev.columns:
        for _, _r in _prev.dropna(subset=['deal_id']).iterrows():
            _modelo_por_deal[int(_r['deal_id'])] = _r.get('modelo')


def _tipo_por_modelo(deal_id):
    modelo = str(_modelo_por_deal.get(deal_id) or '').strip().lower()
    if not modelo:
        return None
    return 'partner' if modelo == 'partner' else 'franquia'


produto_rows = []
if len(df_deals) and won_no_mes_mask.any():
    deal_ids_contratos = df_deals.loc[won_no_mes_mask, 'deal_id'].tolist()
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        for deal_id, produtos in executor.map(_fetch_deal_products, deal_ids_contratos):
            produto_rows.append({
                'deal_id': deal_id,
                'produtos': ', '.join(p.get('name') or '' for p in produtos) or None,
                'tipo_contrato': _classifica_produtos(produtos) or _tipo_por_modelo(deal_id),
            })
    n_franquia = sum(1 for r in produto_rows if r['tipo_contrato'] == 'franquia')
    n_partner = sum(1 for r in produto_rows if r['tipo_contrato'] == 'partner')
    n_sem = sum(1 for r in produto_rows if r['tipo_contrato'] is None)
    print(f'   contratos: franquia={n_franquia} | partner={n_partner} | sem produto/não classificado={n_sem}')
pd.DataFrame(produto_rows).to_excel(os.path.join(OUT_DIR, 'deals_produtos.xlsx'), index=False)

# ============= NEGÓCIOS DO FUNIL STUDIO FISCAL FRANQUIA (pipeline 62, todos) =============
# "Negócios iniciados" (leads) considera TODO negócio criado no funil, independente
# do responsável — não só os 4 closers (o funil ainda é majoritariamente da Qualificadora).
# Também usado pra apurar perdidos (por estágio/motivo) e base de leads por campanha/criativo.
# OBS: /deals?pipeline_id=X é um filtro não confiável nessa conta (retorna negócios de
# outros funis) — usar sempre /pipelines/{id}/deals, com status=all_not_deleted pra
# trazer também os perdidos (status=lost sozinho retorna vazio nesse endpoint).
print('-> Fetching negócios do funil STUDIO FISCAL FRANQUIA (pipeline 62, todos)...')
p62_raw = fetch_all(f'pipelines/{config.PIPELINE_STUDIO_FISCAL_FRANQUIA}/deals', {'status': 'all_not_deleted'})

# `/pipelines/{id}/deals?status=all_not_deleted` só devolve deals OPEN nessa conta
# (bug/particularidade da API já confirmada em teste manual) — perdidos precisam vir
# de `/deals?status=lost`, que não aceita filtro de pipeline confiável, então buscamos
# ordenado por lost_time DESC e paramos assim que sair da janela do mês (evita puxar
# as ~58 mil perdas de todos os funis/meses).
print('-> Fetching negócios perdidos (todos os funis, parando na janela do mês)...')
lost_rows_all = []
start = 0
stop = False
while not stop:
    params = {'status': 'lost', 'sort': 'lost_time DESC', 'limit': 500, 'start': start, 'api_token': TOKEN}
    qs = urllib.parse.urlencode(params)
    body = None
    for attempt in range(10):
        try:
            with urllib.request.urlopen(f'{BASE}/deals?{qs}', timeout=120) as r:
                body = json.loads(r.read())
            break
        except urllib.error.HTTPError as e:
            if e.code == 429:
                wait = 15 * (attempt + 1)
                print(f'     rate limit (429), esperando {wait}s...')
                time.sleep(wait)
            elif attempt == 9:
                raise
            else:
                time.sleep(3)
    if body is None:
        break
    data = body.get('data') or []
    if not data:
        break
    for d in data:
        lt = _fix_dt(d.get('lost_time'))
        if lt and _in_range(lt, config.MONTH_START, config.MONTH_END):
            lost_rows_all.append(d)
        elif lt and pd.to_datetime(lt).date() < config.MONTH_START:
            stop = True
    pag = body.get('additional_data', {}).get('pagination', {})
    if not pag.get('more_items_in_collection'):
        break
    start = pag.get('next_start', start + len(data))

p62_lost = [d for d in lost_rows_all if d.get('pipeline_id') == config.PIPELINE_STUDIO_FISCAL_FRANQUIA]
print(f'   perdidos no mês (todos funis): {len(lost_rows_all)} | no funil 62: {len(p62_lost)}')

p62_rows = []
for d in p62_raw + p62_lost:
    owner = d.get('user_id') or {}
    p62_rows.append({
        'deal_id': d.get('id'),
        'titulo': d.get('title'),
        'stage_id': d.get('stage_id'),
        'stage_change_time': _fix_dt(d.get('stage_change_time')),
        'status': d.get('status'),
        'add_time': _fix_dt(d.get('add_time')),
        'lost_time': _fix_dt(d.get('lost_time')),
        'lost_reason': d.get('lost_reason'),
        'owner_id': owner.get('id') if isinstance(owner, dict) else owner,
        'owner_nome': owner.get('name') if isinstance(owner, dict) else None,
        'utm_campanha': d.get(config.CF_UTM_CAMPAIGN),
        'utm_criativo': d.get(config.CF_UTM_TERM),
        # Sempre string (não deixar o pandas/Excel inferir número) — o campo é
        # multi-select e vem como "97" ou "69,97"; se virasse float perderíamos a
        # comparação exata (97 -> "97.0") e a etiqueta pra achar quem é Studio Fiscal.
        'label': str(d.get('label')) if d.get('label') is not None else None,
    })
df_p62 = pd.DataFrame(p62_rows)
df_p62.to_excel(os.path.join(OUT_DIR, 'deals_pipeline62.xlsx'), index=False)
n_p62_leads = df_p62['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df_p62) else 0
n_p62_lost = ((df_p62['status'] == 'lost') & df_p62['lost_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END))).sum() if len(df_p62) else 0
print(f'   negócios no funil (todos, {len(df_p62)}) | negócios iniciados no mês: {n_p62_leads} | perdidos no mês: {n_p62_lost}')

# ============= NEGÓCIOS DO FUNIL OUTBOUND - STUDIO FISCAL FRANQUIA (pipeline 66, todos) =============
# Mesma lógica do funil 62 — leads brutos importados nascem com o Mauricio como
# owner até um closer criar/assumir o negócio (owner_id vira o closer). "Leads" do
# outbound é contado no build_dashboard.py pelo owner_id atual + add_time no mês.
print('-> Fetching negócios do funil OUTBOUND - STUDIO FISCAL FRANQUIA (pipeline 66, todos)...')
p66_raw = fetch_all(f'pipelines/{config.PIPELINE_OUTBOUND_STUDIO_FISCAL}/deals', {'status': 'all_not_deleted'})
p66_lost = [d for d in lost_rows_all if d.get('pipeline_id') == config.PIPELINE_OUTBOUND_STUDIO_FISCAL]
print(f'   perdidos no funil 66 (dentre os já buscados no mês): {len(p66_lost)}')

p66_rows = []
for d in p66_raw + p66_lost:
    owner = d.get('user_id') or {}
    p66_rows.append({
        'deal_id': d.get('id'),
        'titulo': d.get('title'),
        'stage_id': d.get('stage_id'),
        'stage_change_time': _fix_dt(d.get('stage_change_time')),
        'status': d.get('status'),
        'add_time': _fix_dt(d.get('add_time')),
        'lost_time': _fix_dt(d.get('lost_time')),
        'lost_reason': d.get('lost_reason'),
        'owner_id': owner.get('id') if isinstance(owner, dict) else owner,
        'owner_nome': owner.get('name') if isinstance(owner, dict) else None,
        'utm_campanha': d.get(config.CF_UTM_CAMPAIGN),
        'utm_criativo': d.get(config.CF_UTM_TERM),
    })
df_p66 = pd.DataFrame(p66_rows)
df_p66.to_excel(os.path.join(OUT_DIR, 'deals_pipeline66.xlsx'), index=False)
n_p66_leads = df_p66['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df_p66) else 0
n_p66_leads_closers = (
    df_p66['owner_id'].isin(config.CLOSER_IDS)
    & df_p66['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END))
).sum() if len(df_p66) else 0
print(f'   negócios no funil outbound (todos, {len(df_p66)}) | criados no mês: {n_p66_leads} | criados no mês já com um closer como dono: {n_p66_leads_closers}')


def _fetch_pipeline_deals(pipeline_id, nome_log):
    """Mesma lógica de p62/p66: busca todo negócio do pipeline (aberto via endpoint
    próprio, perdido via lost_rows_all já buscado) e devolve as linhas no formato
    padrão usado pelas páginas do dashboard."""
    print(f'-> Fetching negócios do funil {nome_log} (pipeline {pipeline_id}, todos)...')
    raw = fetch_all(f'pipelines/{pipeline_id}/deals', {'status': 'all_not_deleted'})
    lost = [d for d in lost_rows_all if d.get('pipeline_id') == pipeline_id]
    print(f'   perdidos no funil {pipeline_id} (dentre os já buscados no mês): {len(lost)}')
    rows = []
    for d in raw + lost:
        owner = d.get('user_id') or {}
        rows.append({
            'deal_id': d.get('id'),
            'titulo': d.get('title'),
            'stage_id': d.get('stage_id'),
            'stage_change_time': _fix_dt(d.get('stage_change_time')),
            'status': d.get('status'),
            'add_time': _fix_dt(d.get('add_time')),
            'lost_time': _fix_dt(d.get('lost_time')),
            'lost_reason': d.get('lost_reason'),
            'owner_id': owner.get('id') if isinstance(owner, dict) else owner,
            'owner_nome': owner.get('name') if isinstance(owner, dict) else None,
            'utm_campanha': d.get(config.CF_UTM_CAMPAIGN),
            'utm_criativo': d.get(config.CF_UTM_TERM),
            'label': str(d.get('label')) if d.get('label') is not None else None,
        })
    df = pd.DataFrame(rows)
    df.to_excel(os.path.join(OUT_DIR, f'deals_pipeline{pipeline_id}.xlsx'), index=False)
    n_leads = df['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df) else 0
    print(f'   negócios no funil (todos, {len(df)}) | criados no mês: {n_leads}')
    return df


# ============= REBARBA STUDIO FISCAL FRANQUIA (pipeline 65) — reativação de perdidos =============
df_p65 = _fetch_pipeline_deals(config.PIPELINE_REBARBA_STUDIO_FISCAL, 'REBARBA STUDIO FISCAL FRANQUIA')

# ============= OUTBOUND - EVENTOS (pipeline 69) — reportado junto com o Outbound =============
df_p69 = _fetch_pipeline_deals(config.PIPELINE_OUTBOUND_EVENTOS, 'OUTBOUND - EVENTOS')

print(f'\nDados salvos em {OUT_DIR}/')
