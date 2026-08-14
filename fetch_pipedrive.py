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
            return deal_id, d.get(config.CF_UTM_CAMPAIGN), d.get(config.CF_UTM_CONTENT), d.get('pipeline_id')
        except Exception:
            if attempt == 3:
                return deal_id, None, None, None
            time.sleep(2 * (attempt + 1))
    return deal_id, None, None, None


all_rows = [r for rows in all_realizada.values() for r in rows] + [r for rows in all_futura.values() for r in rows]
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
            })

df_deals = pd.DataFrame(deals_rows)
df_deals.to_excel(os.path.join(OUT_DIR, 'deals.xlsx'), index=False)
n_leads = df_deals['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df_deals) else 0
n_contratos = ((df_deals['status'] == 'won') & df_deals['won_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END))).sum() if len(df_deals) else 0
print(f'   leads no mês: {n_leads} | contratos (Ganho) no mês: {n_contratos}')

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
        'utm_criativo': d.get(config.CF_UTM_CONTENT),
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
        'utm_criativo': d.get(config.CF_UTM_CONTENT),
    })
df_p66 = pd.DataFrame(p66_rows)
df_p66.to_excel(os.path.join(OUT_DIR, 'deals_pipeline66.xlsx'), index=False)
n_p66_leads = df_p66['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df_p66) else 0
n_p66_leads_closers = (
    df_p66['owner_id'].isin(config.CLOSER_IDS)
    & df_p66['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END))
).sum() if len(df_p66) else 0
print(f'   negócios no funil outbound (todos, {len(df_p66)}) | criados no mês: {n_p66_leads} | criados no mês já com um closer como dono: {n_p66_leads_closers}')

print(f'\nDados salvos em {OUT_DIR}/')
