"""
Detecta negócios no Pipedrive que ganharam Produto + Data de Fechamento
Esperada (os 2 campos que, juntos, indicam que o closer já está tratando o
negócio como uma previsão de venda) e que ainda NÃO estão na planilha de
Previsão — e adiciona uma linha nova pra cada um, seguindo o padrão de
preenchimento das linhas manuais. Negócio modelo Parceria (Venda R$0) entra
igual, desde que já tenha o produto "Contrato Partner" vinculado — pedido do
time (03/11/2026): antes só entrava quando Ganho, mas Parceria em aberto
também é previsão (de fechamento, não de receita) e tem que aparecer.

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


def _norm_nome(nome):
    """Normaliza nome do lead pra comparar sem falso-negativo: minúsculo, sem
    acento, espaços colapsados, E só primeiro + último nome (ignora nomes do
    meio) — pra "Cyro Neto" casar com "Cyro Galvani Neto" (mesma pessoa, o time
    digitou de um jeito na planilha e o Pipedrive tem de outro). Usado pra NÃO
    adicionar um lead que já consta na planilha, mesmo que seja outro negócio /
    outro deal_id / outro funil. Pedido do time (10/09/2026 — casos Nilza
    Bernardes e Cyro Galvani Neto)."""
    import unicodedata
    n = unicodedata.normalize('NFD', str(nome or ''))
    n = ''.join(c for c in n if unicodedata.category(c) != 'Mn')
    toks = n.lower().split()
    if not toks:
        return ''
    return toks[0] if len(toks) == 1 else f'{toks[0]} {toks[-1]}'


def _fmt_data_br(data_iso):
    if not data_iso or pd.isna(data_iso):
        return ''
    try:
        return pd.to_datetime(data_iso).strftime('%d/%m/%Y')
    except Exception:
        return ''


def _parse_valor_cel(txt):
    """Lê de volta um valor que a própria planilha mostra formatado (ex: "R$
    38.250,00") pra comparar com o número cru do Pipedrive. None se a célula
    estiver vazia ou não for um número reconhecível."""
    if not txt or not str(txt).strip():
        return None
    s = str(txt).strip().replace('R$', '').strip()
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


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
col_satus = next((i for i, c in enumerate(header) if c.lower().startswith('satus') or c.lower().startswith('status')), None)
col_confianca = next((i for i, c in enumerate(header) if c.lower().startswith('confian')), None)

col_lead = next((i for i, c in enumerate(header) if c.lower().startswith('lead')), 2)
col_data_esperada = next((i for i, c in enumerate(header) if 'fechamento' in c.lower() and 'esperad' in c.lower()), None)
col_venda = next((i for i, c in enumerate(header) if c.lower().strip() == 'venda'), None)
col_caixa = next((i for i, c in enumerate(header) if c.lower().strip() == 'caixa'), None)

deal_ids_existentes = set()
nomes_existentes = set()  # nomes de lead (normalizados) que já estão na planilha — não adiciona de novo, mesmo que seja outro negócio
linhas_auto = []  # (linha 1-based, deal_id) das linhas que ESTE script já criou (Satus começa com "Novo (auto)")
linhas_com_deal = []  # (linha 1-based, deal_id, satus_cell, confianca_cell, data_esperada_cell, venda_cell, caixa_cell) de TODA linha com link válido, mesmo editada à mão pelo time
ultima_linha_com_dado = header_row_idx  # índice 0-based da última linha com Closer+Lead preenchidos
for i in range(header_row_idx + 1, len(existing)):
    row = existing[i]
    link = row[col_link] if col_link < len(row) else ''
    m = DEAL_LINK_RE.search(link or '')
    if m:
        deal_ids_existentes.add(int(m.group(1)))
    closer_cell = row[1] if len(row) > 1 else ''
    lead_cell = row[col_lead] if col_lead < len(row) else ''
    if (lead_cell or '').strip():
        nomes_existentes.add(_norm_nome(lead_cell))
    if (closer_cell or '').strip() and (lead_cell or '').strip():
        ultima_linha_com_dado = i
    satus_cell = row[col_satus] if col_satus is not None and col_satus < len(row) else ''
    confianca_cell = row[col_confianca] if col_confianca is not None and col_confianca < len(row) else ''
    data_esperada_cell = row[col_data_esperada] if col_data_esperada is not None and col_data_esperada < len(row) else ''
    venda_cell = row[col_venda] if col_venda is not None and col_venda < len(row) else ''
    caixa_cell = row[col_caixa] if col_caixa is not None and col_caixa < len(row) else ''
    if m and (satus_cell or '').strip().startswith('Novo (auto)'):
        linhas_auto.append((i + 1, int(m.group(1))))  # +1: gspread é 1-based
    if m:
        linhas_com_deal.append((i + 1, int(m.group(1)), satus_cell, confianca_cell, data_esperada_cell, venda_cell, caixa_cell))

print(f'   {len(deal_ids_existentes)} negócio(s) / {len(nomes_existentes)} nome(s) já na planilha ({len(linhas_auto)} ainda "Novo (auto)")')

# ============= 2) Candidatos no Pipedrive (aberto + valor + data esperada) =============
deals_path = os.path.join(OUT_DIR, 'deals.xlsx')
if not os.path.exists(deals_path):
    _pular('data_pipedrive/deals.xlsx não existe ainda — rode fetch_pipedrive.py primeiro')

df = pd.read_excel(deals_path)

# ============= 1.5) Atualiza status/data de TODA linha com negócio vinculado =============
# Antes só mexia em linha ainda marcada "Novo (auto)" — assim que o time
# escrevia qualquer nota na coluna Satus (ex: "SUMIU", "Aguardando
# pagamento"), a linha ficava travada pra sempre: se o negócio depois fechasse
# ou morresse de verdade no Pipedrive, a planilha nunca mais refletia isso.
# Caso real (28/09/2026): Luis Pacheco (deal 58331) foi marcado "lost" no
# Pipedrive, mas a linha continuou com Confiança=FRIO porque o Satus já tinha
# sido editado à mão pro time. Corrigido: agora TODA linha com link válido é
# checada. A coluna "Satus" (texto livre do time) só é reescrita quando ainda
# está em "Novo (auto)" (mesmo comportamento de antes); "CONFIANÇA" (o
# termômetro que vira o badge Ganho/Perdido no dashboard) e "Data fechamento
# Esperada" (fato que vem do Pipedrive, não opinião do time) são mantidas
# sincronizadas em QUALQUER linha, editada ou não — só não mexe se o valor já
# está igual, pra não gerar barulho de edição à toa.
if linhas_com_deal and not df.empty and col_confianca is not None:
    status_por_deal = dict(zip(df['deal_id'], df['status']))
    data_esperada_por_deal = dict(zip(df['deal_id'], df['expected_close_date'])) if 'expected_close_date' in df.columns else {}
    valor_por_deal = dict(zip(df['deal_id'], df['valor'])) if 'valor' in df.columns else {}
    caixa_por_deal = dict(zip(df['deal_id'], df['caixa'])) if 'caixa' in df.columns else {}
    probabilidade_por_deal = dict(zip(df['deal_id'], df['probability'])) if 'probability' in df.columns else {}
    updates = []  # (a1, valor) pra um único batch_update no final
    logs = []
    for linha, deal_id, satus_cell, confianca_cell, data_esperada_cell, venda_cell, caixa_cell in linhas_com_deal:
        status_atual = status_por_deal.get(deal_id)
        is_novo_auto = (satus_cell or '').strip().startswith('Novo (auto)')
        partes_log = []

        if status_atual == 'won':
            nova_confianca = 'FECHADO'
        elif status_atual == 'lost':
            nova_confianca = 'PERDIDO'
        elif status_atual == 'open':
            # Confiança do negócio em aberto vem da Probabilidade do Pipedrive
            # (campo que o closer já ajusta no negócio) — fato, não mais uma
            # estimativa manual do time. Pedido do time (08/10/2026): >80%
            # Quente, 50-80% Morno, <=50% Frio. Sem probabilidade definida no
            # Pipedrive, não mexe na Confiança que já está na planilha.
            prob = probabilidade_por_deal.get(deal_id)
            if prob is None or pd.isna(prob):
                nova_confianca = None
            elif prob > 80:
                nova_confianca = 'QUENTE'
            elif prob > 50:
                nova_confianca = 'MORNO'
            else:
                nova_confianca = 'FRIO'
        else:
            nova_confianca = None
        if nova_confianca and (confianca_cell or '').strip().upper() != nova_confianca:
            if col_confianca is not None:
                updates.append((gspread.utils.rowcol_to_a1(linha, col_confianca + 1), nova_confianca))
                partes_log.append(f'Confiança={nova_confianca}')
            if is_novo_auto and col_satus is not None and nova_confianca in ('FECHADO', 'PERDIDO'):
                novo_satus = 'Ganho (auto)' if nova_confianca == 'FECHADO' else 'Perdido (auto)'
                updates.append((gspread.utils.rowcol_to_a1(linha, col_satus + 1), novo_satus))
                partes_log.append(f'Satus="{novo_satus}"')

        if col_data_esperada is not None:
            data_nova = data_esperada_por_deal.get(deal_id)
            data_nova_fmt = _fmt_data_br(data_nova)
            if data_nova_fmt:
                try:
                    igual = pd.to_datetime(data_nova).date() == pd.to_datetime(data_esperada_cell, dayfirst=True).date()
                except Exception:
                    igual = False
                if not igual:
                    updates.append((gspread.utils.rowcol_to_a1(linha, col_data_esperada + 1), data_nova_fmt))
                    partes_log.append(f'Data fechamento esperada={data_nova_fmt}')

        # Venda/Caixa também são fato do Pipedrive (o valor do negócio), não
        # opinião do time — mantém sincronizado em qualquer linha, editada ou
        # não, igual à data esperada. Pedido do time (08/10/2026, casos reais:
        # Mirian e Bruna Bressani com o valor corrigido no Pipedrive depois
        # que a linha já tinha entrado na planilha, e a planilha nunca mais
        # acompanhou).
        for col_idx, deal_dict, cel_atual, rotulo in (
            (col_venda, valor_por_deal, venda_cell, 'Venda'),
            (col_caixa, caixa_por_deal, caixa_cell, 'Caixa'),
        ):
            if col_idx is None:
                continue
            valor_novo = deal_dict.get(deal_id)
            if valor_novo is None or pd.isna(valor_novo):
                continue
            valor_atual = _parse_valor_cel(cel_atual)
            if valor_atual is not None and abs(float(valor_novo) - valor_atual) < 0.5:
                continue
            updates.append((gspread.utils.rowcol_to_a1(linha, col_idx + 1), float(valor_novo)))
            partes_log.append(f'{rotulo}=R$ {float(valor_novo):,.2f}')

        if partes_log:
            logs.append(f'   ~ linha {linha} (deal {deal_id}): ' + ', '.join(partes_log))

    if updates:
        ws.batch_update([{'range': a1, 'values': [[valor]]} for a1, valor in updates], value_input_option='USER_ENTERED')
        for l in logs:
            print(l)

if df.empty or 'expected_close_date' not in df.columns:
    print('-> Sync automático da Previsão: nenhum negócio novo pra checar (deals.xlsx sem expected_close_date — rode fetch_pipedrive.py atualizado).')
    sys.exit(0)

candidatos = df[
    (df['status'] == 'open')
    & df['pipeline_id'].isin(config.FUNIL_POR_PIPELINE.keys())
    & df['expected_close_date'].notna()
    & ~df['deal_id'].isin(deal_ids_existentes)
].copy()

# Negócio com fechamento esperado num mês FUTURO (depois do mês corrente) não
# entra na planilha — pedido do time (08/10/2026): a planilha é só pro mês
# atual (+ negócios em aberto atrasados de meses passados). Se continuar
# aberto quando o mês dele chegar, entra normalmente nessa hora.
_data_esperada_dt = pd.to_datetime(candidatos['expected_close_date'], errors='coerce')
candidatos = candidatos[_data_esperada_dt.dt.date.apply(lambda x: pd.isna(x) or x <= config.MONTH_END)]

# A Rebarba começa numa etapa literalmente chamada "PERDIDO" (reativação de
# negócio perdido) — valor/data esperada ali podem ser resquício do negócio
# antigo, não o closer de fato trabalhando de novo. Só entra a partir de
# "OPORTUNIDADE (FPQ RECEBIDA)" ou depois (pedido do time, 09/09/2026).
_rebarba = candidatos['pipeline_id'] == config.PIPELINE_REBARBA_STUDIO_FISCAL
candidatos = candidatos[~_rebarba | (candidatos['stage_id'] >= config.PIPELINE_STAGE_OPORTUNIDADE_REBARBA)]

# Se o NOME do lead já consta na planilha (por qualquer negócio, mesmo outro
# deal_id / outro funil), não adiciona de novo — pedido do time (10/09/2026).
if not candidatos.empty and nomes_existentes:
    _nome_lead = candidatos.apply(lambda r: _norm_nome(r.get('pessoa_nome') or r.get('titulo') or ''), axis=1)
    candidatos = candidatos[~_nome_lead.isin(nomes_existentes)]

print(f'   {len(candidatos)} negócio(s) aberto(s) com data esperada, ainda não na planilha (antes de checar produto vinculado)')

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

# ============= 3.5) Backfill de negócios já GANHOS no mês que nunca passaram
# pela planilha =============
# O filtro de "candidato" acima exige Venda > 0 — mas um negócio de modelo
# Partner (parceria, sem cobrança) sempre tem Venda R$0 no Pipedrive, então
# NUNCA entra como candidato "aberto", mesmo depois de fechado. Resultado:
# esse fechamento aparecia certo no Placar do mês (que lê direto do Pipedrive,
# sem depender da planilha) mas sumia da tabela de Previsão — o time via
# "3 contratos" no Placar e só 2 "Ganho" aqui embaixo. Corrigido (11/09/2026,
# caso real: Antonio Santana, Partner, Venda R$0, nunca virou linha). Aqui não
# depende de Venda>0 nem de Data esperada (o negócio já fechou, não precisa
# mais de previsão) — só de ter produto vinculado (pra saber o Modelo) e ter
# sido Ganho ESTE mês num dos pipelines da Fiscal.
if not df.empty and 'won_time' in df.columns:
    _wt = pd.to_datetime(df['won_time'], errors='coerce')
    ganhos = df[
        (df['status'] == 'won')
        & df['pipeline_id'].isin(config.FUNIL_POR_PIPELINE.keys())
        & _wt.notna()
        & (_wt.dt.date >= config.MONTH_START) & (_wt.dt.date <= config.MONTH_END)
        & ~df['deal_id'].isin(deal_ids_existentes)
    ].copy()
    if not ganhos.empty and nomes_existentes:
        _nomes_g = ganhos.apply(lambda r: _norm_nome(r.get('pessoa_nome') or r.get('titulo') or ''), axis=1)
        ganhos = ganhos[~_nomes_g.isin(nomes_existentes)]
    if not ganhos.empty:
        print(f'   {len(ganhos)} negócio(s) Ganho(s) este mês, ainda não na planilha (checando produto vinculado)')
    for _, d in ganhos.iterrows():
        deal_id = int(d['deal_id'])
        produtos = _fetch_deal_products(deal_id)
        if not produtos:
            continue  # ganho mas sem produto vinculado — não dá pra saber o Modelo
        modelo = _modelo_from_produto(produtos[0].get('name'))
        hoje = date.today().strftime('%d/%m/%Y')
        caixa = d.get('caixa')
        data_fech = _fmt_data_br(d.get('expected_close_date')) or _fmt_data_br(d.get('won_time'))
        novas_linhas.append([
            hoje, _closer_primeiro_nome(d.get('closer_nome')), d.get('pessoa_nome') or d.get('titulo') or '',
            config.FUNIL_POR_PIPELINE.get(int(d['pipeline_id']), ''), modelo, 'Studio Fiscal',
            float(caixa) if pd.notna(caixa) else '', float(d['valor']), data_fech,
            'Ganho (auto)', hoje, 'FECHADO', f'https://{config.PIPEDRIVE_DOMAIN}/deal/{deal_id}',
        ])
        print(f'   + [Ganho] {d.get("pessoa_nome") or d.get("titulo")} (deal {deal_id}, {modelo}) — R$ {d["valor"]:,.2f}')

if not novas_linhas:
    print('-> Sync automático da Previsão: nada novo pra adicionar (candidatos sem produto vinculado ainda).')
    sys.exit(0)

# Dedup dentro do próprio lote — se o mesmo lead tiver 2 negócios abertos que
# passaram por todos os filtros, entra 1 só (o 1º).
_vistos, _dedup = set(), []
for ln in novas_linhas:
    chave = _norm_nome(ln[2])
    if chave in _vistos:
        continue
    _vistos.add(chave)
    _dedup.append(ln)
novas_linhas = _dedup

if len(novas_linhas) > MAX_NOVOS_POR_CICLO:
    print(f'   AVISO: {len(novas_linhas)} negócio(s) prontos pra entrar é mais que o esperado (trava em {MAX_NOVOS_POR_CICLO}) '
          f'— pulando esse ciclo pra não arriscar poluir a planilha. Confira manualmente.')
    sys.exit(0)

# ============= 4) Re-checa a planilha AGORA (o motor local e o da nuvem rodam
# no mesmo horário — sem isso, os dois liam a planilha, viam o mesmo negócio
# faltando e adicionavam a MESMA linha 2x; caso real 10/09/2026: Lourivelthon
# Moreira Barros entrou duplicado) e escreve. =============
_ex2 = ws.get_all_values()
_hdr2 = next((i for i, r in enumerate(_ex2) if r and r[0].strip() == 'DATA DO REGISTRO'), header_row_idx)
_ids2, _nomes2 = set(), set()
_ultima2 = _hdr2
for i in range(_hdr2 + 1, len(_ex2)):
    r = _ex2[i]
    mm = DEAL_LINK_RE.search(r[col_link] if col_link < len(r) else '')
    if mm:
        _ids2.add(int(mm.group(1)))
    ld = r[col_lead] if col_lead < len(r) else ''
    if (ld or '').strip():
        _nomes2.add(_norm_nome(ld))
    cl = r[1] if len(r) > 1 else ''
    if (cl or '').strip() and (ld or '').strip():
        _ultima2 = i

# novas_linhas é sempre montada na ordem canônica: [0]=Data [2]=Lead [12]=Link
novas_linhas = [
    ln for ln in novas_linhas
    if int(DEAL_LINK_RE.search(ln[12]).group(1)) not in _ids2
    and _norm_nome(ln[2]) not in _nomes2
]
if not novas_linhas:
    print('-> Sync automático da Previsão: outro ciclo já adicionou esses negócios nesse meio-tempo — nada a fazer.')
    sys.exit(0)

# Insere logo depois da última linha com dado (não usa "append" simples) —
# empurra pra baixo o que tiver depois (linhas em branco, linha TOTAL com
# fórmula de soma), igual um Inserir Linha manual faria. O Google Sheets
# expande sozinho o range de uma SUM() que cobria a linha de cima/baixo do
# ponto de inserção, então o TOTAL continua somando tudo certo.
insert_at = _ultima2 + 2  # +1 pra ir pra próxima linha, +1 pra converter índice 0-based -> linha 1-based do gspread
ws.insert_rows(novas_linhas, row=insert_at, value_input_option='USER_ENTERED')
print(f'-> Sync automático da Previsão: {len(novas_linhas)} linha(s) nova(s) adicionada(s) na planilha.')
