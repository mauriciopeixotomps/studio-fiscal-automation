"""
Baixa a planilha de Previsão (Google Sheets, CSV público) e cruza com o Pipedrive
pra saber Venda/Caixa realizado quando o negócio referenciado em "Link do pipe"
já estiver Ganho.

Gera /data_pipedrive/previsao.xlsx
"""
import urllib.request, urllib.parse, json, os, re, time
import pandas as pd
import io

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
DEAL_LINK_RE = re.compile(r'/deal/(\d+)')


def _closer_key(nome_planilha):
    if not isinstance(nome_planilha, str):
        return None
    n = nome_planilha.strip().lower()
    for c in config.CLOSERS:
        if n in c['apelidos_planilha']:
            return c['nome']
    return nome_planilha.strip()  # não reconhecido — mantém como veio


def _parse_valor(v):
    """'R$ 35.000,00' -> 35000.0"""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace('R$', '').strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def _parse_data(v):
    if not v or not str(v).strip():
        return None
    try:
        return pd.to_datetime(str(v).strip(), dayfirst=True, errors='coerce').date().isoformat()
    except Exception:
        return None


def _fetch_deal(deal_id):
    url = f'{BASE}/deals/{deal_id}?api_token={TOKEN}'
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=30) as r:
                body = json.loads(r.read())
            return body.get('data')
        except Exception:
            if attempt == 3:
                return None
            time.sleep(2 * (attempt + 1))
    return None


print('-> Baixando planilha de previsão (CSV)...')
with urllib.request.urlopen(config.PREVISAO_CSV_URL, timeout=30) as r:
    csv_text = r.read().decode('utf-8')

lines = csv_text.splitlines()
header_idx = next((i for i, l in enumerate(lines) if l.startswith('DATA DO REGISTRO')), None)
if header_idx is None:
    raise RuntimeError('Não achei a linha de cabeçalho "DATA DO REGISTRO" na planilha')

df = pd.read_csv(io.StringIO('\n'.join(lines[header_idx:])))
df.columns = [c.strip() for c in df.columns]
df = df.dropna(how='all')
print(f'   {len(df)} linhas de previsão')

col_modelo = next((c for c in df.columns if c.lower() == 'modelo'), None)
col_bandeira = next((c for c in df.columns if c.lower() == 'bandeira'), None)

rows = []
for _, r in df.iterrows():
    link = str(r.get('Link do pipe') or '')
    m = DEAL_LINK_RE.search(link)
    deal_id = int(m.group(1)) if m else None

    venda_realizada = None
    caixa_realizado = None
    deal_status = None
    if deal_id:
        d = _fetch_deal(deal_id)
        if d:
            deal_status = d.get('status')
            if deal_status == 'won':
                venda_realizada = d.get('value')
                caixa_realizado = d.get(config.CF_CAIXA)

    rows.append({
        'data_registro': _parse_data(r.get('DATA DO REGISTRO')),
        'closer': _closer_key(r.get('Closer')),
        'lead': r.get('Lead'),
        'modelo': r.get(col_modelo) if col_modelo else None,
        'bandeira': r.get(col_bandeira) if col_bandeira else None,
        'venda': _parse_valor(r.get('Venda')),
        'caixa': _parse_valor(r.get('Caixa')),
        'data_fechamento_esperada': _parse_data(r.get('Data fechamento Esperada')),
        'status': r.get('Satus') or r.get('Status'),
        'deal_id': deal_id,
        'deal_link': link if link and link != 'nan' else None,
        'deal_status': deal_status,
        'venda_realizada': venda_realizada,
        'caixa_realizado': caixa_realizado,
    })

df_out = pd.DataFrame(rows)
df_out.to_excel(os.path.join(OUT_DIR, 'previsao.xlsx'), index=False)
n_realizados = df_out['venda_realizada'].notna().sum()
print(f'   {n_realizados} negócios já Ganhos (venda/caixa realizado preenchido)')
print(f'Salvo em {OUT_DIR}/previsao.xlsx')

if not col_modelo or not col_bandeira:
    faltando = [n for n, c in (('Modelo', col_modelo), ('Bandeira', col_bandeira)) if not c]
    print(f'   AVISO: coluna(s) {faltando} ainda não existem na planilha — ficarão em branco até serem adicionadas.')
