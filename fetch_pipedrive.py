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
for stage_key, act_type, tem_futura in ACT_STAGES:
    raw = acts_by_stage[stage_key]
    realizada = [act_row(a) for a in raw
                 if a.get('done') and _in_range(a.get('marked_as_done_time'), config.MONTH_START, config.MONTH_END)]
    if stage_key == 'cc1':
        cc1_realizada_rows = realizada
    else:
        pd.DataFrame(realizada).to_excel(os.path.join(OUT_DIR, f'activities_{stage_key}_realizada.xlsx'), index=False)
    print(f'   {stage_key} realizada: {len(realizada)}')
    if tem_futura:
        futura = [act_row(a) for a in raw
                  if not a.get('done') and _in_range(a.get('due_date'), config.MONTH_START, config.MONTH_END)]
        pd.DataFrame(futura).to_excel(os.path.join(OUT_DIR, f'activities_{stage_key}_futura.xlsx'), index=False)
        print(f'   {stage_key} para acontecer: {len(futura)}')

# ============= UTM (campanha/criativo) das CC1 realizada =============
# Enriquece cada CC1 realizada com a campanha/criativo do negócio vinculado,
# pra dar pra rankear campanhas/criativos por CC1 gerada.
print('-> Enriquecendo CC1 realizada com UTM campanha/criativo do negócio...')


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


unique_deal_ids = sorted({r['negocio_id'] for r in cc1_realizada_rows if r.get('negocio_id')})
utm_by_deal = {}
if unique_deal_ids:
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        for deal_id, campanha, criativo, pipeline_id in executor.map(_fetch_deal_utm, unique_deal_ids):
            utm_by_deal[deal_id] = (campanha, criativo, pipeline_id)

for r in cc1_realizada_rows:
    campanha, criativo, pipeline_id = utm_by_deal.get(r.get('negocio_id'), (None, None, None))
    r['utm_campanha'] = campanha
    r['utm_criativo'] = criativo
    r['pipeline_id'] = pipeline_id

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
                'add_time': d.get('add_time'),
                'won_time': d.get('won_time'),
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
        lt = d.get('lost_time')
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
        'stage_change_time': d.get('stage_change_time'),
        'status': d.get('status'),
        'add_time': d.get('add_time'),
        'lost_time': d.get('lost_time'),
        'lost_reason': d.get('lost_reason'),
        'owner_id': owner.get('id') if isinstance(owner, dict) else owner,
        'owner_nome': owner.get('name') if isinstance(owner, dict) else None,
        'utm_campanha': d.get(config.CF_UTM_CAMPAIGN),
        'utm_criativo': d.get(config.CF_UTM_CONTENT),
    })
df_p62 = pd.DataFrame(p62_rows)
df_p62.to_excel(os.path.join(OUT_DIR, 'deals_pipeline62.xlsx'), index=False)
n_p62_leads = df_p62['add_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END)).sum() if len(df_p62) else 0
n_p62_lost = ((df_p62['status'] == 'lost') & df_p62['lost_time'].apply(lambda x: _in_range(x, config.MONTH_START, config.MONTH_END))).sum() if len(df_p62) else 0
print(f'   negócios no funil (todos, {len(df_p62)}) | negócios iniciados no mês: {n_p62_leads} | perdidos no mês: {n_p62_lost}')

print(f'\nDados salvos em {OUT_DIR}/')
