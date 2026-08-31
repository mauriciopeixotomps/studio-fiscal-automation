"""
Baixa a planilha de Previsão (Google Sheets, CSV público) e cruza com o Pipedrive
pra saber Venda/Caixa realizado quando o negócio referenciado em "Link do pipe"
já estiver Ganho.

Gera /data_pipedrive/previsao.xlsx
"""
import urllib.request, urllib.parse, json, os, re, time
from datetime import timedelta
import pandas as pd
import io

import config

# Mesma correção de fuso (-3h) do fetch_pipedrive.py — a API do Pipedrive devolve
# TODO campo de data/hora 3h à frente do horário real de Brasília nessa conta.
TZ_OFFSET = timedelta(hours=3)


def _fix_dt(dt_str):
    if not dt_str:
        return dt_str
    try:
        return (pd.to_datetime(dt_str) - TZ_OFFSET).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return dt_str

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
DEAL_LINK_RE = re.compile(r'https?://([^/\s]+)/deal/(\d+)')


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


def _fetch_deal_produtos(deal_id):
    url = f'{BASE}/deals/{deal_id}/products?api_token={TOKEN}'
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=30) as r:
                body = json.loads(r.read())
            return body.get('data') or []
        except Exception:
            if attempt == 3:
                return []
            time.sleep(2 * (attempt + 1))
    return []


def _tipo_contrato(deal_id, modelo_texto=None):
    """'franquia' (produto com valor) / 'partner' (Contrato Partner, de graça) /
    None (sem produto vinculado nem Modelo preenchido). Prioridade: produto de
    verdade vinculado no Pipedrive (mesma classificação usada pros contratos
    fechados em fetch_pipedrive.py); se não tiver produto vinculado ainda, cai
    pro texto da coluna "Modelo" da planilha — confirmado com o time (31/08/2026):
    "GS Partner" É um modelo de franquia (tem valor); só "Partner" sozinho (sem o
    "GS" na frente) é o modelo de parceria (de graça)."""
    if deal_id:
        nomes = [p.get('name') or '' for p in _fetch_deal_produtos(deal_id)]
        if any(n.startswith('Franquia') for n in nomes):
            return 'franquia'
        if any(n == 'Contrato Partner' for n in nomes):
            return 'partner'
    modelo = (modelo_texto or '').strip().lower()
    if not modelo:
        return None
    return 'partner' if modelo == 'partner' else 'franquia'


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
# Ignora linha de resumo/total (ex: "TOTAL,,,,,R$ x,R$ y,,,") — só conta linhas com Closer e Lead preenchidos
df = df[df['Closer'].notna() & df['Lead'].notna()]
print(f'   {len(df)} linhas de previsão')

col_modelo = next((c for c in df.columns if c.lower() == 'modelo'), None)
col_bandeira = next((c for c in df.columns if c.lower() == 'bandeira'), None)
col_confianca = next((c for c in df.columns if c.lower() in ('confiança', 'confianca', 'termômetro', 'termometro')), None)
col_funil = next((c for c in df.columns if c.lower() == 'funil'), None)

rows = []
for _, r in df.iterrows():
    link = str(r.get('Link do pipe') or '')
    m = DEAL_LINK_RE.search(link)
    deal_id = int(m.group(2)) if m else None
    # Reconstrói um link limpo a partir do domínio/id extraídos, evitando URLs de
    # redirecionamento do Google (que às vezes envolvem o link real quando colado no Sheets).
    clean_link = f'https://{m.group(1)}/deal/{m.group(2)}' if m else None

    venda_realizada = None
    caixa_realizado = None
    deal_status = None
    won_time = None
    if deal_id:
        d = _fetch_deal(deal_id)
        if d:
            deal_status = d.get('status')
            if deal_status == 'won':
                venda_realizada = d.get('value')
                caixa_realizado = d.get(config.CF_CAIXA)
                won_time = _fix_dt(d.get('won_time'))
    else:
        # Sem link válido do Pipedrive (célula tem nome/e-mail em vez da URL) —
        # usa o status preenchido manualmente na planilha como fallback. Sem
        # won_time nesse caso (não tem negócio no Pipedrive pra consultar), então
        # não entra no Placar do mês (que exige won_time no mês corrente).
        status_texto = str(r.get('Satus') or r.get('Status') or '').strip().lower()
        if 'fechado' in status_texto or 'ganho' in status_texto:
            deal_status = 'won'
            venda_realizada = _parse_valor(r.get('Venda'))
            caixa_realizado = _parse_valor(r.get('Caixa'))

    rows.append({
        'data_registro': _parse_data(r.get('DATA DO REGISTRO')),
        'closer': _closer_key(r.get('Closer')),
        'lead': r.get('Lead'),
        'modelo': r.get(col_modelo) if col_modelo else None,
        'bandeira': r.get(col_bandeira) if col_bandeira else None,
        'confianca': r.get(col_confianca) if col_confianca else None,
        'venda': _parse_valor(r.get('Venda')),
        'caixa': _parse_valor(r.get('Caixa')),
        'data_fechamento_esperada': _parse_data(r.get('Data fechamento Esperada')),
        'status': r.get('Satus') or r.get('Status'),
        'deal_id': deal_id,
        'deal_link': clean_link or (link if link and link != 'nan' else None),
        'deal_status': deal_status,
        'venda_realizada': venda_realizada,
        'caixa_realizado': caixa_realizado,
        'won_time': won_time,
        # 'Modelo' na planilha é o NOME comercial do plano (TAX/Platinum/Corporate/GS
        # Partner/GS Black) — não indica se é o "Contrato Partner" de graça, então
        # classifica pelo produto de verdade vinculado no Pipedrive (mesma lógica dos
        # contratos fechados).
        'tipo_contrato': _tipo_contrato(deal_id, r.get(col_modelo) if col_modelo else None),
        # Qual página (Inbound/Outbound/Rebarba) esse negócio previsto pertence —
        # normaliza maiúscula/minúscula ("INBOUND", "Rebarba" etc. já vêm misturados
        # na planilha) pra comparar sem erro (31/08/2026).
        'funil': str(r.get(col_funil)).strip().upper() if col_funil and pd.notna(r.get(col_funil)) else None,
    })

df_out = pd.DataFrame(rows)
df_out.to_excel(os.path.join(OUT_DIR, 'previsao.xlsx'), index=False)
n_realizados = df_out['venda_realizada'].notna().sum() if not df_out.empty else 0
print(f'   {n_realizados} negócios já Ganhos (venda/caixa realizado preenchido)')
print(f'Salvo em {OUT_DIR}/previsao.xlsx')

if not col_modelo or not col_bandeira or not col_confianca:
    faltando = [n for n, c in (('Modelo', col_modelo), ('Bandeira', col_bandeira), ('Confiança', col_confianca)) if not c]
    print(f'   AVISO: coluna(s) {faltando} ainda não existem na planilha — ficarão em branco até serem adicionadas.')
