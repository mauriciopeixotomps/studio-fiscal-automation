"""
Detecta negócios no Pipedrive que ganharam Produto + Valor + Data de Fechamento
Esperada (os 3 campos que, juntos, indicam que o closer já está tratando o
negócio como uma previsão de venda) e que ainda NÃO estão na planilha de
Previsão — e adiciona uma linha nova pra cada um, seguindo o padrão de
preenchimento das linhas manuais.

Roda depois de fetch_pipedrive.py (usa data_pipedrive/deals.xlsx, que precisa
ter as colunas 'expected_close_date' e 'pessoa_nome') e antes de
fetch_previsao.py, pra a linha nova já entrar no mesmo ciclo do dashboard.

Pedido do time (09/09/2026): "cada vez que alguém da equipe adicionar produto
e valor e data de fechamento esperada em algum negócio, ele seja adicionado na
planilha conforme o padrão de preenchimento". Confirmado com o usuário: a
linha nova entra DIRETO na tabela principal (não numa área separada pra
revisão), marcada na coluna "Satus" como "Novo (auto)" pra o time saber que
foi o robô e ainda precisa revisar/completar; "CONFIANÇA" começa como MORNO
(meio-termo conservador) até o time reclassificar pra Quente/Perdido/Fechado.

Requer:
- .google_service_account.json (conta de serviço do Google com permissão de
  Editor na planilha de Previsão) — ver GOOGLE_SERVICE_ACCOUNT_SETUP.md.
- Bibliotecas gspread + google-auth (requirements.txt).

Não-fatal por design: se faltar a credencial, ou a API do Sheets falhar por
qualquer motivo, avisa e sai com código 0 — não derruba o resto do ciclo de
automação (fetch_previsao.py / build / deploy continuam normalmente, a
planilha segue funcionando manualmente como sempre até isso ser resolvido).
"""
import json
import os
import re
import sys
import time
import urllib.request
from datetime import date

import pandas as pd

import config

WORKDIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(WORKDIR, 'data_pipedrive')
TOKEN_FILE = os.path.join(WORKDIR, '.pipedrive_token')
TOKEN = os.environ.get('PIPEDRIVE_TOKEN') or (
    open(TOKEN_FILE).read().strip() if os.path.exists(TOKEN_FILE) else None
)
BASE = 'https://api.pipedrive.com/v1'

# Trava de segurança: se detectar mais candidatos que isso num ciclo só, algo
# está errado na lógica de detecção (ou a planilha ficou fora do ar um tempo
# grande) — melhor parar e avisar do que arriscar poluir a planilha do time.
MAX_NOVOS_POR_CICLO = 20

CRED_FILE = os.environ.get('GOOGLE_SERVICE_ACCOUNT_FILE') or config.GOOGLE_SERVICE_ACCOUNT_FILE
CRED_JSON_ENV = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')  # GitHub Actions manda o conteúdo direto, não um arquivo


def _pular(motivo):
    print(f'-> Sync automático da Previsão: PULADO ({motivo})')
    sys.exit(0)


try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    _pular('faltam as libs gspread/google-auth — rode `pip install -r requirements.txt`')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

if CRED_JSON_ENV:
    creds = Credentials.from_service_account_info(json.loads(CRED_JSON_ENV), scopes=SCOPES)
elif os.path.exists(CRED_FILE):
    creds = Credentials.from_service_account_file(CRED_FILE, scopes=SCOPES)
else:
    _pular(
        f'credencial não encontrada — salve o JSON da conta de serviço em {CRED_FILE} '
        f'(ou defina a env var GOOGLE_SERVICE_ACCOUNT_JSON) — ver GOOGLE_SERVICE_ACCOUNT_SETUP.md'
    )

gc = gspread.authorize(creds)

try:
    sh = gc.open_by_key(config.PREVISAO_SHEET_ID)
    ws = sh.get_worksheet(0)  # mesma aba (gid=0) que fetch_previsao.py já lê
except Exception as e:
    _pular(f'sem acesso de edição à planilha ainda ({e}) — compartilhe com o e-mail da conta de serviço, ver GOOGLE_SERVICE_ACCOUNT_SETUP.md')

DEAL_LINK_RE = re.compile(r'/deal/(\d+)')


def _closer_primeiro_nome(closer_nome):
    return closer_nome.split()[0] if closer_nome else ''


def _modelo_from_produto(nome_produto):
    """'Franquia Studio Fiscal Tax' -> 'TAX'; 'Contrato Partner' -> 'PARTNER'
    (mesmo padrão curto usado na coluna Modelo da planilha)."""
    n = (nome_produto or '').strip()
    if not n:
        return ''
    if n.lower() == 'contrato partner':
        return 'PARTNER'
    low = n.lower()
    for prefixo in ('franquia studio fiscal ', 'franquia '):
        if low.startswith(prefixo):
            return n[len(prefixo):].strip().upper()
    return n.upper()


def _fmt_data_br(data_iso):
    if not data_iso or pd.isna(data_iso):
        return ''
    try:
        return pd.to_datetime(data_iso).strftime('%d/%m/%Y')
    except Exception:
        return ''


def _fetch_deal_products(deal_id):
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


# ============= 1) Quem já está na planilha (pra não duplicar) =============
print('-> Sync automático da Previsão: lendo planilha atual...')
existing = ws.get_all_values()
header_row_idx = next((i for i, row in enumerate(existing) if row and row[0].strip() == 'DATA DO REGISTRO'), None)
if header_row_idx is None:
    _pular('não achei a linha de cabeçalho "DATA DO REGISTRO" na planilha')

header = [c.strip() for c in existing[header_row_idx]]
col_link = next((i for i, c in enumerate(header) if c.lower().startswith('link do pipe')), None)
if col_link is None:
    _pular('não achei a coluna "Link do pipe" na planilha')

deal_ids_existentes = set()
ultima_linha_com_dado = header_row_idx  # índice 0-based da última linha com Closer+Lead preenchidos
for i in range(header_row_idx + 1, len(existing)):
    row = existing[i]
    link = row[col_link] if col_link < len(row) else ''
    m = DEAL_LINK_RE.search(link or '')
    if m:
        deal_ids_existentes.add(int(m.group(1)))
    closer_cell = row[1] if len(row) > 1 else ''
    lead_cell = row[2] if len(row) > 2 else ''
    if (closer_cell or '').strip() and (lead_cell or '').strip():
        ultima_linha_com_dado = i

print(f'   {len(deal_ids_existentes)} negócio(s) já na planilha')

# ============= 2) Candidatos no Pipedrive (aberto + valor + data esperada) =============
deals_path = os.path.join(OUT_DIR, 'deals.xlsx')
if not os.path.exists(deals_path):
    _pular('data_pipedrive/deals.xlsx não existe ainda — rode fetch_pipedrive.py primeiro')

df = pd.read_excel(deals_path)
if df.empty or 'expected_close_date' not in df.columns:
    print('-> Sync automático da Previsão: nenhum negócio pra checar (deals.xlsx sem expected_close_date — rode fetch_pipedrive.py atualizado).')
    sys.exit(0)

candidatos = df[
    (df['status'] == 'open')
    & df['pipeline_id'].isin(config.FUNIL_POR_PIPELINE.keys())
    & df['valor'].fillna(0).gt(0)
    & df['expected_close_date'].notna()
    & ~df['deal_id'].isin(deal_ids_existentes)
].copy()

print(f'   {len(candidatos)} negócio(s) aberto(s) com valor + data esperada, ainda não na planilha (antes de checar produto vinculado)')

# ============= 3) Confirma produto vinculado (o 3º critério pedido) e monta as linhas =============
# A trava de segurança (MAX_NOVOS_POR_CICLO) é checada só DEPOIS desse filtro —
# testado em 09/09/2026 com o backlog real: 28 negócios tinham valor+data
# esperada (muitos Rebarba antigos, reativação que nunca avançou), mas só 3
# tinham produto de verdade vinculado. Cortar antes do filtro de produto ia
# travar até esse punhado de 3 legítimos.
novas_linhas = []
for _, d in candidatos.iterrows():
    deal_id = int(d['deal_id'])
    produtos = _fetch_deal_products(deal_id)
    if not produtos:
        continue  # tem valor + data esperada, mas ainda sem produto vinculado — não conta ainda
    modelo = _modelo_from_produto(produtos[0].get('name'))
    hoje = date.today().strftime('%d/%m/%Y')
    caixa = d.get('caixa')
    novas_linhas.append([
        hoje,                                                        # DATA DO REGISTRO
        _closer_primeiro_nome(d.get('closer_nome')),                  # Closer
        d.get('pessoa_nome') or d.get('titulo') or '',                # Lead
        config.FUNIL_POR_PIPELINE.get(int(d['pipeline_id']), ''),     # FUNIL
        modelo,                                                        # Modelo
        'Studio Fiscal',                                               # Bandeira
        float(caixa) if pd.notna(caixa) else '',                       # Caixa (só se o campo de Caixa estiver preenchido no Pipedrive)
        float(d['valor']),                                             # Venda
        _fmt_data_br(d.get('expected_close_date')),                   # Data fechamento Esperada
        'Novo (auto)',                                                 # Satus
        hoje,                                                           # DATA DA ATUALIZAÇÃO
        'MORNO',                                                        # CONFIANÇA
        f'https://{config.PIPEDRIVE_DOMAIN}/deal/{deal_id}',           # Link do pipe
    ])
    print(f'   + {d.get("pessoa_nome") or d.get("titulo")} (deal {deal_id}, {modelo}) — R$ {d["valor"]:,.2f}')

if not novas_linhas:
    print('-> Sync automático da Previsão: nada novo pra adicionar (candidatos sem produto vinculado ainda).')
    sys.exit(0)

if len(novas_linhas) > MAX_NOVOS_POR_CICLO:
    print(f'   AVISO: {len(novas_linhas)} negócio(s) prontos pra entrar é mais que o esperado (trava em {MAX_NOVOS_POR_CICLO}) '
          f'— pulando esse ciclo pra não arriscar poluir a planilha. Confira manualmente.')
    sys.exit(0)

# ============= 4) Escreve na planilha =============
# Insere logo depois da última linha com dado (não usa "append" simples) —
# empurra pra baixo o que tiver depois (linhas em branco, linha TOTAL com
# fórmula de soma), igual um Inserir Linha manual faria. O Google Sheets
# expande sozinho o range de uma SUM() que cobria a linha de cima/baixo do
# ponto de inserção, então o TOTAL continua somando tudo certo.
insert_at = ultima_linha_com_dado + 2  # +1 pra ir pra próxima linha, +1 pra converter índice 0-based -> linha 1-based do gspread
ws.insert_rows(novas_linhas, row=insert_at, value_input_option='USER_ENTERED')
print(f'-> Sync automático da Previsão: {len(novas_linhas)} linha(s) nova(s) adicionada(s) na planilha.')
