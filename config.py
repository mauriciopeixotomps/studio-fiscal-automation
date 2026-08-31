"""Configuração do Dashboard Studio Fiscal — equipe de vendas (closers)."""
from datetime import date
import calendar
import os

# ============= PERÍODO (mês corrente) =============
# DASHBOARD_MONTH (env var, formato 'AAAA-MM') força o mês pra gerar uma página
# congelada de um mês já fechado (ex: arquivar julho antes de virar pra agosto),
# sem depender da data real do sistema. Uso normal (automação) não define isso.
_month_override = os.environ.get('DASHBOARD_MONTH')
if _month_override:
    _override_year, _override_month = map(int, _month_override.split('-'))
    _today = date(_override_year, _override_month, 1)
else:
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

# ============= AGENDADORES (equipe de agendamento, a partir de 11/08/2026) =============
# SDRs dedicados a agendar CC1 só pro time da Studio Fiscal a partir de 11/08/2026
# (antes atendiam vários times — ver histórico de chat). Não são closers: não têm
# negócio como "dono" no sentido de fechamento, o painel só mostra a produtividade
# deles (leads recebidos / agendamentos / realizadas).
AGENDADORES = [
    {
        'id': 24196535,
        'nome': 'Lorenzo Meneghetti',
        'email': 'lorenzo.meneghetti@grupostudio.com.br',
    },
    {
        'id': 24560514,
        'nome': 'Jonatans Junior',
        'email': 'jonatans.junior@grupostudio.com.br',
    },
    {
        'id': 24165647,
        'nome': 'Leonardo Soares',
        'email': 'leonardo.soares@grupostudio.com.br',
    },
]
AGENDADOR_IDS = [a['id'] for a in AGENDADORES]

# Metas do time de agendamento (confirmado com o time, 11/08/2026):
# - % agendamento = agendamentos / leads recebidos
# - % realizada = CC1s realizadas / agendamentos
META_AGENDAMENTO_PCT = 0.30
META_AGENDADOR_REALIZADA_PCT = 0.65

# ============= PIPEDRIVE: tipos de atividade =============
ACT_CC1 = 'meeting'             # id 2  — CC1 Realizada (done=true) / CC1 Para Acontecer (done=false)
ACT_FPQ = 'fpq_recebida'        # id 13 — FPQ Recebida
ACT_CC2 = 'cc2_realizada1'      # id 10 — CC2 Realizada (done=true) / CC2 Para Acontecer (done=false)
ACT_CC3 = 'cc3_realizada'       # id 32 — CC3 Realizada

# "Leads qualificados" (Consolidado do time, antes da CC1) — existem DUAS chaves de
# atividade chamadas "Qualificação Realizada" na conta ('qualificacao_realizada' e
# 'qualificacao_realizada1'); confirmado em 31/08/2026 que só a com "1" no fim tem
# uso real (a outra tem zero atividades em toda a história da conta).
ACT_QUALIFICACAO_REALIZADA = 'qualificacao_realizada1'

# Atividades do agendador: "Leads recebidos" = 1ª tentativa de contato criada pra ele
# (task que nasce quando o lead é roteado pra ele); "Agendamento" = CC1 marcada como
# agendada por ele (manual ou pelo bot) — a CC1 em si (config.ACT_CC1) fica na agenda
# do CLOSER, não do agendador, então "Realizada" é calculada cruzando o negocio_id do
# agendamento com os CC1 Realizada já buscados pros closers (ver build_dashboard.py).
ACT_TENTATIVA_1 = '1o_tentativa_de_contato'
ACT_AGENDAMENTO_CC1 = 'agendamento_de_cc1'
ACT_AGENDAMENTO_CC1_BOT = 'agendamento_cc1___bot'

# No-show: o time passou a usar essas duas tarefas a partir de 29/08/2026 (antes,
# 0 registros em toda a conta). O "responsável" (user_id) da tarefa é o CLOSER, não
# o agendador — assim dá pra ver de quem foi o no-show pelo lado do closer (dono da
# tarefa) e, cruzando o negocio_id com os agendamentos, também de qual agendador
# (ver build_dashboard.py — mesmo cruzamento já usado pra achar "Realizada").
ACT_NO_SHOW_IMPREVISTO = 'no_show_imprevisto'
ACT_NO_SHOW_DESINTERESSE = 'no_show_desinteresse'

PIPELINE_STUDIO_FISCAL_FRANQUIA = 62

# Funil novo de outbound (pipeline própria, dashboard separado — ver publish/outbound.html)
# — mesma estrutura de 11 etapas do funil 62, com PROSPECT/ABORDAGEM INICIAL/EM
# ATENDIMENTO/REUNIÃO AGENDADA no lugar de LEAD/PRÉ-QUALIFICAÇÃO/EM ATENDIMENTO/MQL
# FRANQUIA; de AGUARDANDO FPQ em diante os nomes e a ordem são idênticos. Cada closer
# cria/adiciona os próprios leads ali (import inicial foi só pontual, sob o Mauricio).
# Ainda sem meta, só produtividade.
PIPELINE_OUTBOUND_STUDIO_FISCAL = 66

# Etapas do pipeline 62, pra referência (IDs reais, confirmados com o time):
# 593 LEAD, 594 PRÉ-QUALIFICAÇÃO, 595 EM ATENDIMENTO, 596 MQL FRANQUIA (qualificação,
# CC1 pode ou não estar agendada) -> 597 AGUARDANDO FPQ (CC1 realizada, 1ª etapa do
# closer), 598 OPORTUNIDADE (FPQ recebida), 599 OPORTUNIDADE QUENTE (CC2 agendada),
# 600 PROPOSTA (CC2 realizada), 601 NEGOCIAÇÃO, 602 ADITIVO JURIDICO, 603 CONTRATO
# ENVIADO. Repasse não é definido pela etapa (o closer às vezes faz a CC1 e já marca
# como perdido ainda em MQL FRANQUIA) — quem perdeu o negócio (owner_id) é que decide
# se foi a Qualificadora (SDR) ou a Expansão (um dos 4 closers). Ver build_dashboard.py.

# "Contrato enviado" é contado pela etapa do negócio (chegou em CONTRATO ENVIADO),
# não por uma atividade separada — o time não estava registrando essa atividade de
# forma confiável, então a etapa do funil é a fonte de verdade (usa stage_change_time
# pra saber quando entrou; só pega negócios que estão *atualmente* nessa etapa).
PIPELINE_STAGE_CONTRATO_ENVIADO = 603

# Etapa "CONTRATO ENVIADO" equivalente no funil Outbound (pipeline 66) — mesma
# lógica, offset consistente de +92 confirmado nas 7 últimas etapas dos dois funis.
PIPELINE_STAGE_CONTRATO_ENVIADO_OUTBOUND = 695

# ============= NOVOS FUNIS (confirmado com o time, 29/08/2026) =============
# "REBARBA STUDIO FISCAL FRANQUIA" — reativação de negócios perdidos do funil 62
# (etapas: PERDIDO -> REATIVAÇÃO -> EM ATENDIMENTO -> REUNIÃO AGENDADA -> AGUARDANDO
# FPQ -> ... -> CONTRATO ENVIADO, de "REUNIÃO AGENDADA" em diante é o mesmo desenho
# do 62). Página própria (publish/rebarba.html), como o Outbound: ainda sem meta, só
# produtividade (CC1/FPQ/CC2/CC3/Contrato enviado/Contratos).
PIPELINE_REBARBA_STUDIO_FISCAL = 65
PIPELINE_STAGE_CONTRATO_ENVIADO_REBARBA = 633

# "OUTBOUND - EVENTOS" — outro motor de outbound (convite/parceria em eventos), sem
# etiqueta de marca nenhuma (diferente do 62/65, que usam LABEL_STUDIO_FISCAL) — o
# time confirmou que esse pipeline é 100% da Fiscal, conta tudo sem filtrar por dono
# nem etiqueta. Reportado JUNTO com o Outbound (mesma página, publish/outbound.html)
# — pro time ele é "outbound", não um funil separado.
PIPELINE_OUTBOUND_EVENTOS = 69
PIPELINE_STAGE_CONTRATO_ENVIADO_EVENTOS = 721

# Domínio Pipedrive da conta — usado pra montar link direto pro negócio (ex: nas FPQs recebidas)
PIPEDRIVE_DOMAIN = 'grupostudioexp.pipedrive.com'

# Campo customizado de negócio: Valor de Caixa (Entrada + Taxa)
CF_CAIXA = 'a11287b4a40b38a06e050e24b1bbc78e91a6da53'

# Campos customizados de negócio: origem UTM (campanha/criativo de marketing)
CF_UTM_CAMPAIGN = 'ee93fbc62c1804b82721486630058f3467651cec'
CF_UTM_CONTENT = '6673bb1656201df819dcd47be1cf751a966d2229'

# Etiqueta (campo "label" nativo do Pipedrive, multi-select) que a Qualificadora usa
# pra marcar um negócio como sendo da Studio Fiscal — id 97 "Lead Studio Fiscal"
# (conferido em 11/08/2026: 162 dos 170 negócios adicionados no funil 62 em agosto
# tinham essa etiqueta; a "Studio Fiscal" sem "Lead" na frente, id 69, é usada 1x só,
# não é essa). "Negócios iniciados" (Ritmo/Metas do mês) usa essa etiqueta pra contar
# todo lead gerado pra Studio Fiscal no mês, independente de já ter dono entre os
# closers (SDR/Qualificadora inclusos) — diferente do "Negócios iniciados" por closer
# no Relatório, que continua sendo CC1 realizada OU adicionado já com o closer como dono.
LABEL_STUDIO_FISCAL = '97'

# ============= NEGÓCIOS EXCLUÍDOS =============
# deal_id que aparece com um dos 4 closers como "owner" no Pipedrive, mas que o time
# confirmou não ser fechamento real. Excluído de leads/contratos por closer e do
# consolidado.
# - 57206: negócio do Émerson, de outra equipe, mal atribuído no Pipedrive.
# - 84764, 92183: marcados "Ganho" (valor R$0) ainda em estágio inicial da
#   Qualificadora (LEAD/PRÉ-QUALIFICAÇÃO), no mesmo minuto — parece marcação em
#   massa/erro de dados, não fechamento real, confirmado com o time (03/07/2026).
# - 95176 (Nelson/Leonardo): fechado de verdade em julho (won_time 31/07), o
#   Pipedrive reabriu sozinho em 05/08 e o Leonardo deu Ganho de novo no mesmo dia
#   — já contado no arquivo congelado de julho (julho-2026.html), excluído aqui
#   pra não contar de novo como fechamento de agosto (confirmado com o time,
#   05/08/2026).
DEALS_EXCLUIDOS = {57206, 84764, 92183, 95176}

# deal_id da planilha de Previsão pra desconsiderar do Placar do mês (e da própria
# tabela de Previsão) — não mexe na planilha em si, só filtra na leitura.
# - 90987 (Sandro/Émerson): venda da Studio Fiscal, mas Émerson já é da Studio Agro
#   agora — o time pediu pra não contar mais aqui na Fiscal (21/07/2026).
PREVISAO_EXCLUIDOS = {90987}

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
