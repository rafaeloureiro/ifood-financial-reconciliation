import pandas as pd
import openpyxl

# ==========================================================
# CONFIGURAÇÃO
# ==========================================================

CAMINHO_PLANILHA = r"C:\Users\rafae\OneDrive\Desktop\Lab. de testes\Relatorios Ifood\conciliacao-de-faturamento\conciliacao_02\pedidos.xlsx"

# Taxas contratuais por modalidade de entrega
COMISSAO_ENTREGA_PROPRIA     = 0.09    # SELF_DELIVERY — base: valor dos itens bruto
COMISSAO_ENTREGA_FLEX        = 0.20    # Entrega Flex (referência contratual)
COMISSAO_SOB_DEMANDA         = 0.09    # Sob Demanda ON
COMISSAO_RETIRADA_LOJA       = 0.0875  # Pra Retirar
TAXA_TRANSACAO_PAGAMENTO_APP = 0.03    # Pagamentos via APP, exceto VR/SODEXO/ALELO

# Tolerância para considerar uma divergência relevante (em R$)
TOLERANCIA_DIVERGENCIA = 1.00

# Formas de pagamento que não geram taxa de transação nem entram no repasse iFood.
# VR/SODEXO/ALELO via APP são isentos de taxa E recebidos direto na loja —
# por isso aparecem em ambas as listas abaixo (intencional).
FORMAS_PAGAMENTO_ISENTAS_TAXA = [
    "Pgto via APP - Vale Refeição (VR)",
    "Pgto via APP - Vale Refeição (SODEXO)",
    "Pgto via APP - Vale Refeição (ALELO)",
]

FORMAS_PAGAMENTO_DIRETO_LOJA = [
    "Pgto via APP - Vale Refeição (VR)",
    "Pgto via APP - Vale Refeição (SODEXO)",
    "Pgto via APP - Vale Refeição (ALELO)",
]


def pagamento_recebido_na_loja(forma_de_pagamento: str) -> bool:
    return (
        forma_de_pagamento == "Dinheiro"
        or forma_de_pagamento == "Outros vales"
        or forma_de_pagamento.startswith("Pgto na Entrega")
        or forma_de_pagamento in FORMAS_PAGAMENTO_DIRETO_LOJA
    )


# ==========================================================
# LEITURA E LIMPEZA
# ==========================================================

def carregar_planilha(caminho: str) -> pd.DataFrame:
    with open(caminho, "rb") as arquivo:
        workbook  = openpyxl.load_workbook(arquivo)
        planilha  = workbook.active
        cabecalhos = [celula.value for celula in planilha[1]]
        linhas = [
            dict(zip(cabecalhos, linha))
            for linha in planilha.iter_rows(min_row=2, values_only=True)
        ]
    return pd.DataFrame(linhas)


def limpar_e_classificar_pedidos(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()

    colunas_monetarias = [
        "VALOR DOS ITENS (R$)",
        "TOTAL PAGO PELO CLIENTE (R$)",
        "TAXA DE ENTREGA PAGA PELO CLIENTE (R$)",
        "INCENTIVO PROMOCIONAL DO IFOOD (R$)",
        "INCENTIVO PROMOCIONAL DA LOJA (R$)",
        "INCENTIVO PROMOCIONAL DA REDE (R$)",
        "TAXA DE SERVIÇO (R$)",
        "TAXAS E COMISSOES (R$)",
        "VALOR LIQUIDO (R$)",
    ]
    for coluna in colunas_monetarias:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

    df["PRODUTO LOGISTICO"]      = df["PRODUTO LOGISTICO"].fillna("").str.strip()
    df["FORMA DE PAGAMENTO"]     = df["FORMA DE PAGAMENTO"].fillna("").str.strip()
    df["STATUS FINAL DO PEDIDO"] = df["STATUS FINAL DO PEDIDO"].fillna("").str.strip()

    # Status do pedido
    df["pedido_concluido"] = df["STATUS FINAL DO PEDIDO"].isin(["CONCLUIDO", "CANCELAMENTO PARCIAL"])
    df["pedido_cancelado"] = df["STATUS FINAL DO PEDIDO"] == "CANCELADO"

    # Modalidade de entrega
    df["entrega_propria"] = df["PRODUTO LOGISTICO"] == "SELF_DELIVERY_PARTIAL_AREA"
    df["entrega_flex"]    = df["PRODUTO LOGISTICO"] == "ENTREGA FLEX"
    df["sob_demanda"]     = df["PRODUTO LOGISTICO"] == "SOB DEMANDA ON"
    df["retirada_loja"]   = df["PRODUTO LOGISTICO"] == "RETIRADA"

    # Forma de pagamento
    df["pagamento_isento_taxa"]    = df["FORMA DE PAGAMENTO"].isin(FORMAS_PAGAMENTO_ISENTAS_TAXA)
    df["pagamento_via_app"]        = df["FORMA DE PAGAMENTO"].str.contains("Pgto via APP", case=False, na=False)
    df["pagamento_sujeito_taxa"]   = df["pagamento_via_app"] & ~df["pagamento_isento_taxa"]
    df["pagamento_direto_na_loja"] = df["FORMA DE PAGAMENTO"].apply(pagamento_recebido_na_loja)

    return df


# ==========================================================
# CÁLCULO POR LOJA
# ==========================================================

def calcular_metricas_da_loja(pedidos_loja: pd.DataFrame) -> dict:
    pedidos_concluidos = pedidos_loja[pedidos_loja["pedido_concluido"]]
    pedidos_cancelados = pedidos_loja[pedidos_loja["pedido_cancelado"]]
    todos_pedidos      = pedidos_loja

    # Por modalidade — todos os status (o portal contabiliza todos)
    todos_entrega_propria = todos_pedidos[todos_pedidos["entrega_propria"]]
    todos_entrega_flex    = todos_pedidos[todos_pedidos["entrega_flex"]]
    todos_sob_demanda     = todos_pedidos[todos_pedidos["sob_demanda"]]

    # Flex cancelados: base do reembolso de comissão
    cancelados_entrega_flex = pedidos_cancelados[pedidos_cancelados["entrega_flex"]]

    # Pagamentos sujeitos à taxa de transação — todos os status
    todos_com_taxa_transacao = todos_pedidos[todos_pedidos["pagamento_sujeito_taxa"]]

    # Repasse — apenas concluídos, separado por quem recebe
    concluidos_pagamento_direto = pedidos_concluidos[pedidos_concluidos["pagamento_direto_na_loja"]]
    concluidos_repasse_ifood    = pedidos_concluidos[~pedidos_concluidos["pagamento_direto_na_loja"]]

    # Entrega própria concluída: inclui taxa de entrega no valor das vendas
    concluidos_entrega_propria = pedidos_concluidos[pedidos_concluidos["entrega_propria"]]

    # ── Valor das vendas ──────────────────────────────────────────
    valor_total_vendas = (
        pedidos_concluidos["VALOR DOS ITENS (R$)"].sum()
        + concluidos_entrega_propria["TAXA DE ENTREGA PAGA PELO CLIENTE (R$)"].sum()
    )

    # ── Comissões cobradas pelo iFood (conforme CSV) ──────────────
    comissao_cobrada_entrega_propria   = todos_entrega_propria["TAXAS E COMISSOES (R$)"].sum()
    comissao_cobrada_entrega_flex      = todos_entrega_flex["TAXAS E COMISSOES (R$)"].sum()
    comissao_cobrada_sob_demanda       = todos_sob_demanda["TAXAS E COMISSOES (R$)"].sum()
    reembolso_comissao_flex_cancelados = cancelados_entrega_flex["TAXAS E COMISSOES (R$)"].sum()

    # ── Comissões esperadas pelo contrato ─────────────────────────
    base_calculo_entrega_propria = todos_entrega_propria["VALOR DOS ITENS (R$)"].sum()
    base_calculo_entrega_flex    = todos_entrega_flex["VALOR DOS ITENS (R$)"].sum()
    base_calculo_sob_demanda     = todos_sob_demanda["VALOR DOS ITENS (R$)"].sum()

    comissao_esperada_entrega_propria = -(base_calculo_entrega_propria * COMISSAO_ENTREGA_PROPRIA)
    comissao_esperada_entrega_flex    = -(base_calculo_entrega_flex    * COMISSAO_ENTREGA_FLEX)
    comissao_esperada_sob_demanda     = -(base_calculo_sob_demanda     * COMISSAO_SOB_DEMANDA)

    # ── Taxa de transação ─────────────────────────────────────────
    taxa_transacao_cobrada      = todos_pedidos["TAXA DE SERVIÇO (R$)"].sum()
    base_calculo_taxa_transacao = todos_com_taxa_transacao["TOTAL PAGO PELO CLIENTE (R$)"].sum()
    taxa_transacao_esperada     = -(base_calculo_taxa_transacao * TAXA_TRANSACAO_PAGAMENTO_APP)

    # ── Serviços ──────────────────────────────────────────────────
    custo_logistico_sob_demanda = -todos_sob_demanda["TAXA DE ENTREGA PAGA PELO CLIENTE (R$)"].sum()

    # ── Promoções (apenas pedidos concluídos) ─────────────────────
    desconto_promocional_loja   = -pedidos_concluidos["INCENTIVO PROMOCIONAL DA LOJA (R$)"].sum()
    incentivo_promocional_ifood =  pedidos_concluidos["INCENTIVO PROMOCIONAL DO IFOOD (R$)"].sum()

    # ── Repasse ───────────────────────────────────────────────────
    total_recebido_direto_na_loja = concluidos_pagamento_direto["TOTAL PAGO PELO CLIENTE (R$)"].sum()
    total_repasse_ifood           = concluidos_repasse_ifood["VALOR LIQUIDO (R$)"].sum()

    return dict(
        valor_total_vendas                   = valor_total_vendas,
        comissao_cobrada_entrega_propria      = comissao_cobrada_entrega_propria,
        comissao_cobrada_entrega_flex         = comissao_cobrada_entrega_flex,
        comissao_cobrada_sob_demanda          = comissao_cobrada_sob_demanda,
        comissao_entrega_parceira             = comissao_cobrada_entrega_flex + comissao_cobrada_sob_demanda,
        reembolso_comissao_flex_cancelados    = reembolso_comissao_flex_cancelados,
        taxa_transacao_cobrada                = taxa_transacao_cobrada,
        comissao_esperada_entrega_propria     = comissao_esperada_entrega_propria,
        comissao_esperada_entrega_flex        = comissao_esperada_entrega_flex,
        comissao_esperada_sob_demanda         = comissao_esperada_sob_demanda,
        taxa_transacao_esperada               = taxa_transacao_esperada,
        base_calculo_entrega_propria          = base_calculo_entrega_propria,
        base_calculo_entrega_flex             = base_calculo_entrega_flex,
        base_calculo_taxa_transacao           = base_calculo_taxa_transacao,
        custo_logistico_sob_demanda           = custo_logistico_sob_demanda,
        desconto_promocional_loja             = desconto_promocional_loja,
        incentivo_promocional_ifood           = incentivo_promocional_ifood,
        total_recebido_direto_na_loja         = total_recebido_direto_na_loja,
        total_repasse_ifood                   = total_repasse_ifood,
        qtd_pedidos_concluidos                = len(pedidos_concluidos),
        qtd_pedidos_cancelados                = len(pedidos_cancelados),
        qtd_pedidos_entrega_propria           = len(todos_entrega_propria),
        qtd_pedidos_entrega_flex              = len(todos_entrega_flex),
        qtd_pedidos_sob_demanda               = len(todos_sob_demanda),
        qtd_pedidos_entrega_parceira          = len(todos_entrega_flex) + len(todos_sob_demanda),
        qtd_pedidos_direto_na_loja            = len(concluidos_pagamento_direto),
        qtd_pedidos_com_taxa_transacao        = len(todos_com_taxa_transacao),
    )


# ==========================================================
# AVALIAÇÃO DE DIVERGÊNCIAS
# ==========================================================

# Cada observação é um dict com:
#   "status"    → "ok" | "aviso" | "erro"
#   "mensagem"  → texto exibido no relatório
#   "detalhe"   → linha complementar opcional (ou None)

def avaliar_divergencias(metricas: dict) -> list[dict]:
    observacoes = []

    divergencia_entrega_propria = (
        metricas["comissao_cobrada_entrega_propria"]
        - metricas["comissao_esperada_entrega_propria"]
    )
    divergencia_taxa_transacao = (
        metricas["taxa_transacao_cobrada"]
        + metricas["taxa_transacao_esperada"]
    )

    # Comissão de entrega própria
    if metricas["qtd_pedidos_entrega_propria"] == 0:
        observacoes.append({
            "status": "ok",
            "mensagem": "Entrega própria: sem pedidos neste período",
            "detalhe": None,
        })
    elif abs(divergencia_entrega_propria) <= TOLERANCIA_DIVERGENCIA:
        observacoes.append({
            "status": "ok",
            "mensagem": "Comissão de entrega própria dentro do contrato",
            "detalhe": None,
        })
    else:
        valor_cobrado_a_mais = -divergencia_entrega_propria
        observacoes.append({
            "status": "erro",
            "mensagem": f"Comissão de entrega própria divergente: R$ {valor_cobrado_a_mais:,.2f} cobrado a mais",
            "detalhe": (
                f"{COMISSAO_ENTREGA_PROPRIA*100:.0f}% contratado "
                f"sobre base R$ {metricas['base_calculo_entrega_propria']:,.2f} "
                f"= esperado R$ {abs(metricas['comissao_esperada_entrega_propria']):,.2f}"
            ),
        })

    # Taxa de transação
    if metricas["qtd_pedidos_com_taxa_transacao"] == 0:
        observacoes.append({
            "status": "ok",
            "mensagem": "Taxa de transação: sem pagamentos via APP sujeitos à taxa",
            "detalhe": None,
        })
    elif abs(divergencia_taxa_transacao) <= TOLERANCIA_DIVERGENCIA:
        observacoes.append({
            "status": "ok",
            "mensagem": "Taxa de transação online dentro do contrato",
            "detalhe": None,
        })
    else:
        observacoes.append({
            "status": "erro",
            "mensagem": f"Taxa de transação divergente: Δ R$ {divergencia_taxa_transacao:+,.2f}",
            "detalhe": (
                f"{TAXA_TRANSACAO_PAGAMENTO_APP*100:.0f}% contratado "
                f"sobre base R$ {metricas['base_calculo_taxa_transacao']:,.2f} "
                f"= esperado R$ {abs(metricas['taxa_transacao_esperada']):,.2f}"
            ),
        })

    # Entrega Flex: apenas aviso informativo (custo variável, não auditável por %)
    if metricas["qtd_pedidos_entrega_flex"] > 0:
        observacoes.append({
            "status": "aviso",
            "mensagem": "Comissão de entrega flex não conciliável por percentual fixo",
            "detalhe": "Custo logístico variável — conferir via Portal do Parceiro",
        })

    # Itens estruturalmente não conciliáveis pelo relatório de pedidos
    observacoes.append({
        "status": "aviso",
        "mensagem": "Pacote de anúncios não disponível no relatório de pedidos",
        "detalhe": None,
    })
    observacoes.append({
        "status": "aviso",
        "mensagem": "Ajustes financeiros exigem conferência no Portal do Parceiro",
        "detalhe": None,
    })

    return observacoes


# ==========================================================
# FORMATAÇÃO E IMPRESSÃO
# ==========================================================

def formatar_brl(valor: float) -> str:
    negativo = valor < 0
    formatado = f"R$ {abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-{formatado}" if negativo else formatado


def imprimir_linha(label: str, valor: float) -> None:
    print(f"  {label:<50} {formatar_brl(valor):>18}")


def imprimir_subtitulo(texto: str) -> None:
    print(f"    {texto}")


def imprimir_cabecalho_secao(titulo: str) -> None:
    print(f"\n{titulo}")
    print("─" * 70)


ICONE_STATUS = {"ok": "✓", "aviso": "⚠", "erro": "✗"}

def imprimir_observacoes(observacoes: list[dict]) -> None:
    print("\nOBSERVAÇÕES")
    print("─" * 70)
    for obs in observacoes:
        icone = ICONE_STATUS[obs["status"]]
        print(f"  {icone} {obs['mensagem']}")
        if obs["detalhe"]:
            print(f"    {obs['detalhe']}")


# ==========================================================
# EXECUÇÃO
# ==========================================================

df = carregar_planilha(CAMINHO_PLANILHA)
df = limpar_e_classificar_pedidos(df)

nomes_das_lojas = sorted(
    df["NOME DA LOJA"].dropna().astype(str).str.strip().unique()
)

SEPARADOR = "═" * 70

print(f"\n{SEPARADOR}")
print("CONCILIAÇÃO DE FATURAMENTO IFOOD")
print(SEPARADOR)

divergencia_entrega_propria_acumulada = 0.0

for nome_loja in nomes_das_lojas:
    pedidos_loja = df[df["NOME DA LOJA"].astype(str).str.strip() == nome_loja]
    metricas     = calcular_metricas_da_loja(pedidos_loja)
    observacoes  = avaliar_divergencias(metricas)

    divergencia_entrega_propria_acumulada += (
        metricas["comissao_cobrada_entrega_propria"]
        - metricas["comissao_esperada_entrega_propria"]
    )

    tag_loja = nome_loja.split(" - ")[0]

    print(f"\n\n{SEPARADOR}")
    print(tag_loja.upper())
    print(
        f"{metricas['qtd_pedidos_concluidos']} pedidos concluídos | "
        f"{metricas['qtd_pedidos_cancelados']} pedidos cancelados"
    )
    print(SEPARADOR)

    imprimir_cabecalho_secao("VALOR DAS VENDAS")
    imprimir_linha(
        "Valor dos itens e entrega própria da loja",
        metricas["valor_total_vendas"],
    )

    imprimir_cabecalho_secao("TAXAS E COMISSÕES")
    imprimir_linha(
        f"Comissão iFood - entrega parceira ({metricas['qtd_pedidos_entrega_parceira']} ped)",
        metricas["comissao_entrega_parceira"],
    )
    imprimir_linha(
        f"Taxa de transação online ({metricas['qtd_pedidos_com_taxa_transacao']} ped)",
        metricas["taxa_transacao_cobrada"],
    )
    imprimir_linha(
        f"Comissão iFood - entrega própria ({metricas['qtd_pedidos_entrega_propria']} ped)",
        metricas["comissao_cobrada_entrega_propria"],
    )

    imprimir_cabecalho_secao("PROMOÇÕES")
    imprimir_linha("Incentivadas pela loja",  metricas["desconto_promocional_loja"])
    if metricas["incentivo_promocional_ifood"] != 0:
        imprimir_linha("Incentivadas pelo iFood", metricas["incentivo_promocional_ifood"])

    imprimir_cabecalho_secao("SERVIÇOS")
    print("  Pacote de anúncios")
    imprimir_subtitulo("⚠ Não é possível conciliar pelo relatório de pedidos")
    if metricas["custo_logistico_sob_demanda"] != 0:
        imprimir_linha(
            f"Solicitação sob demanda ({metricas['qtd_pedidos_sob_demanda']} ped)",
            metricas["custo_logistico_sob_demanda"],
        )

    imprimir_cabecalho_secao("AJUSTES")
    print("  Reembolso da taxa de serviço cobrada do cliente")
    imprimir_subtitulo("⚠ Não identificado no relatório de pedidos — conferir no Portal")

    imprimir_cabecalho_secao("TOTAL DO REPASSE")
    imprimir_linha(
        f"Valores recebidos diretamente pela loja ({metricas['qtd_pedidos_direto_na_loja']} ped)",
        metricas["total_recebido_direto_na_loja"],
    )
    imprimir_linha("Valor do repasse iFood", metricas["total_repasse_ifood"])

    imprimir_observacoes(observacoes)

# ── Resumo consolidado das 3 lojas ────────────────────────────────
print(f"\n\n{SEPARADOR}")
print("RESUMO — DIVERGÊNCIAS DA SEMANA (3 lojas)")
print("─" * 70)

if abs(divergencia_entrega_propria_acumulada) > TOLERANCIA_DIVERGENCIA:
    valor_cobrado_a_mais = -divergencia_entrega_propria_acumulada
    print(f"  ✗ Comissão própria cobrada a mais:   {formatar_brl(valor_cobrado_a_mais):>18}")
    print(f"    Projeção mensal  (~4 semanas):      {formatar_brl(valor_cobrado_a_mais * 4):>18}")
    print(f"    Projeção anual   (~52 semanas):     {formatar_brl(valor_cobrado_a_mais * 52):>18}")
else:
    print("  ✓ Comissões dentro do esperado contratual.")

print(f"{SEPARADOR}\n")