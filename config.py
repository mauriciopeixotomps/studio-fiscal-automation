"""Configuração do Dashboard Studio Fiscal — equipe de vendas (closers)."""
from datetime import date
import calendar

# ============= PERÍODO (mês corrente) =============
_today = date.today()
MONTH_START = date(_today.year, _today.month, 1)
MONTH_END = date(_today.year, _today.month, calendar.monthrange(_today.year, _today.month)[1])
MONTH_LABEL = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
    7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro',
}[_today.month] + f' {_today.year}'

# Janela ampla de fetch de atividades (pega due_date futuro além do fim do mês,
# pra "para acontecer" continuar aparecendo mesmo perto da virada do mês)
FETCH_START = date(_today.year, _today.month, 1)
FETCH_END = date(_today.year, _today.month, 28) if _today.month == 12 else date(_today.year, _today.month + 1, 28)

# ============= CLOSERS =============
CLOSERS = [
    {
        'id': 22760768,
        'nome': 'Leonardo Ribeiro',
        'email': 'leonardo.ribeiro@grupostudio.com.br',
        'foto': 'https://usericons.pipedrive.com/profile_120x120_22760768_5329c507a3b64e959d03f5aea84e92a7.jpg',
        'apelidos_planilha': {'leo', 'leonardo'},
    },
    {
        'id': 24688510,
        'nome': 'Adalberto Neto',
        'email': 'adalberto.neto@grupostudio.com.br',
        'foto': 'https://usericons.pipedrive.com/profile_120x120_24688510_1a239cdc0558254144dce6bc90463a9a.jpg',
        'apelidos_planilha': {'adalberto', 'adalberto neto'},
    },
    {
        'id': 23882023,
        'nome': 'Fred Garcia',
        'email': 'frederico.garcia@grupostudio.com.br',
        'foto': 'https://usericons.pipedrive.com/profile_120x120_23882023_7a0d02ab9e79475246f7e835ffbcd52a.jpg',
        'apelidos_planilha': {'fred', 'fred garcia', 'frederico'},
    },
    {
        'id': 25500365,
        'nome': 'Luciana Issler',
        'email': 'luciana.issler@grupostudio.com.br',
        'foto': 'https://usericons.pipedrive.com/profile_120x120_25500365_4489819785fd00c507f273ac2dbf8339.jpg',
        'apelidos_planilha': {'luciana', 'luciana issler', 'luciana krewer issler'},
    },
]
CLOSER_IDS = [c['id'] for c in CLOSERS]

# ============= PIPEDRIVE: tipos de atividade =============
ACT_CC1 = 'meeting'             # id 2  — CC1 Realizada (done=true) / CC1 Para Acontecer (done=false)
ACT_FPQ = 'fpq_recebida'        # id 13 — FPQ Recebida
ACT_CC2 = 'cc2_realizada1'      # id 10 — CC2 Realizada (done=true) / CC2 Para Acontecer (done=false)
ACT_CC3 = 'cc3_realizada'       # id 32 — CC3 Realizada
ACT_CONTRATO_ENVIADO = 'contrato_enviado'  # id 11 — Contrato enviado

PIPELINE_STUDIO_FISCAL_FRANQUIA = 62

# Estágios do pipeline 62 que ainda pertencem à Qualificadora (SDR) — antes do
# repasse pro closer. Confirmado com o time (etapas de qualificação vs. etapas do
# closer) e via IDs reais do pipeline 62: repasse acontece na transição de MQL
# FRANQUIA (596, CC1 pode ou não estar agendada) pra AGUARDANDO FPQ (597, CC1 já
# realizada) — a primeira etapa do closer.
PIPELINE62_QUALIFICADORA_STAGES = {593, 594, 595, 596}  # LEAD, PRÉ-QUALIFICAÇÃO, EM ATENDIMENTO, MQL FRANQUIA

# Campo customizado de negócio: Valor de Caixa (Entrada + Taxa)
CF_CAIXA = 'a11287b4a40b38a06e050e24b1bbc78e91a6da53'

# Campos customizados de negócio: origem UTM (campanha/criativo de marketing)
CF_UTM_CAMPAIGN = 'ee93fbc62c1804b82721486630058f3467651cec'
CF_UTM_CONTENT = '6673bb1656201df819dcd47be1cf751a966d2229'

# ============= CONVERSÃO DE REFERÊNCIA (quadro do time, "conta pra cada contrato") =============
# 100 leads -> 15 CC1 realizada -> 8 FPQ -> 4 CC2 -> 2 CC3 -> 1 contrato fechado.
# Usado pra estimar as metas de CC1/FPQ/CC2/CC3 em cascata a partir da meta de leads,
# em vez da conversão real do mês corrente (que fica instável/ruidosa no início do mês,
# com poucos dados). Contrato enviado não está no quadro — continua sendo estimado pela
# conversão real do mês (CC3 -> Contrato enviado).
BENCHMARK_LEADS_PARA_CC1 = 15 / 100
BENCHMARK_CC1_PARA_FPQ = 8 / 15
BENCHMARK_FPQ_PARA_CC2 = 4 / 8
BENCHMARK_CC2_PARA_CC3 = 2 / 4

# ============= META =============
# Um arquivo de metas por mês (mesmo padrão do projeto Qualificadora) — criar um novo
# metas_AAAA_MM.json todo mês com a meta de contratos daquele mês.
METAS_FILE = f'metas_{_today.year}_{_today.month:02d}.json'

# ============= PLANILHA DE PREVISÃO (Google Sheets) =============
PREVISAO_SHEET_ID = '1PJrFhbCAEkcLYeWH0MvRujkWQqpUl9AM0Nwrc2cgYcM'
PREVISAO_CSV_URL = f'https://docs.google.com/spreadsheets/d/{PREVISAO_SHEET_ID}/export?format=csv&gid=0'
