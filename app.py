import pandas as pd
import openpyxl

# ==========================================================
# CONFIGURAÇÃO
# ==========================================================

CAMINHO_PLANILHA = r"C:\Users\rafae\OneDrive\Desktop\Lab. de testes\Relatorios Ifood\conciliacao-de-faturamento\conciliacao_02\pedidos.xlsx"

# Taxas contratuais por modalidade de entrega
COMISSAO_ENTREGA_PROPRIA  = 0.09    # SELF_DELIVERY — base: valor dos itens bruto
COMISSAO_ENTREGA_FLEX     = 0.20    # Entrega Flex (referência contratual)
COMISSAO_SOB_DEMANDA      = 0.09    # Sob Demanda ON
COMISSAO_RETIRADA_LOJA    = 0.0875  # Pra Retirar
TAXA_TRANSACAO_PAGAMENTO_APP = 0.03 # Pagamentos via APP, exceto VR/SODEXO/ALELO

# Formas de pagamento isentas da taxa de transação
FORMAS_PAGAMENTO_ISENTAS_TAXA = [
    "Pgto via APP - Vale Refeição (VR)",
    "Pgto via APP - Vale Refeição (SODEXO)",
    "Pgto via APP - Vale Refeição (ALELO)",
]

# Formas de pagamento recebidas diretamente na loja (não entram no repasse iFood)
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
        workbook = openpyxl.load_workbook(arquivo)
        planilha = workbook.active
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
    df["pagamento_isento_taxa"]   = df["FORMA DE PAGAMENTO"].isin(FORMAS_PAGAMENTO_ISENTAS_TAXA)
    df["pagamento_via_app"]       = df["FORMA DE PAGAMENTO"].str.contains("Pgto via APP", case=False, na=False)
    df["pagamento_sujeito_taxa"]  = df["pagamento_via_app"] & ~df["pagamento_isento_taxa"]
    df["pagamento_direto_na_loja"] = df["FORMA DE PAGAMENTO"].apply(pagamento_recebido_na_loja)

    return df


# ==========================================================
# CÁLCULO POR LOJA
# ==========================================================

def calcular_metricas_da_loja(pedidos_loja: pd.DataFrame) -> dict:
    pedidos_concluidos  = pedidos_loja[pedidos_loja["pedido_concluido"]]
    pedidos_cancelados  = pedidos_loja[pedidos_loja["pedido_cancelado"]]
    todos_pedidos       = pedidos_loja

    # Filtros por modalidade — todos os status (o portal conta todos)
    todos_entrega_propria = todos_pedidos[todos_pedidos["entrega_propria"]]
    todos_entrega_flex    = todos_pedidos[todos_pedidos["entrega_flex"]]
    todos_sob_demanda     = todos_pedidos[todos_pedidos["sob_demanda"]]

    # Filtros por modalidade — apenas concluídos (base de receita e promoções)
    concluidos_entrega_propria = pedidos_concluidos[pedidos_concluidos["entrega_propria"]]
    concluidos_entrega_flex    = pedidos_concluidos[pedidos_concluidos["entrega_flex"]]
    concluidos_sob_demanda     = pedidos_concluidos[pedidos_concluidos["sob_demanda"]]

    # Flex cancelados: usados para calcular o reembolso de comissão
    cancelados_entrega_flex = pedidos_cancelados[pedidos_cancelados["entrega_flex"]]

    # Pagamentos sujeitos à taxa de transação — todos os status (portal cobra mesmo em cancelados)
    todos_com_taxa_transacao = todos_pedidos[todos_pedidos["pagamento_sujeito_taxa"]]

    # Repasse: apenas concluídos, separado por quem recebe
    concluidos_pagamento_direto = pedidos_concluidos[pedidos_concluidos["pagamento_direto_na_loja"]]
    concluidos_repasse_ifood    = pedidos_concluidos[~pedidos_concluidos["pagamento_direto_na_loja"]]

    # ── Valor das vendas ──────────────────────────────────────────
    valor_total_vendas = (
        pedidos_concluidos["VALOR DOS ITENS (R$)"].sum()
        + concluidos_entrega_propria["TAXA DE ENTREGA PAGA PELO CLIENTE (R$)"].sum()
    )

    # ── Comissões cobradas pelo iFood (conforme CSV) ──────────────
    comissao_cobrada_entrega_propria = todos_entrega_propria["TAXAS E COMISSOES (R$)"].sum()
    comissao_cobrada_entrega_flex    = todos_entrega_flex["TAXAS E COMISSOES (R$)"].sum()
    comissao_cobrada_sob_demanda     = todos_sob_demanda["TAXAS E COMISSOES (R$)"].sum()
    reembolso_comissao_flex_cancelados = cancelados_entrega_flex["TAXAS E COMISSOES (R$)"].sum()

    # ── Comissões esperadas pelo contrato (base: valor dos itens bruto) ──
    base_calculo_entrega_propria = todos_entrega_propria["VALOR DOS ITENS (R$)"].sum()
    base_calculo_entrega_flex    = todos_entrega_flex["VALOR DOS ITENS (R$)"].sum()
    base_calculo_sob_demanda     = todos_sob_demanda["VALOR DOS ITENS (R$)"].sum()

    comissao_esperada_entrega_propria = -(base_calculo_entrega_propria * COMISSAO_ENTREGA_PROPRIA)
    comissao_esperada_entrega_flex    = -(base_calculo_entrega_flex    * COMISSAO_ENTREGA_FLEX)
    comissao_esperada_sob_demanda     = -(base_calculo_sob_demanda     * COMISSAO_SOB_DEMANDA)

    # ── Taxa de transação ─────────────────────────────────────────
    taxa_transacao_cobrada  = todos_pedidos["TAXA DE SERVIÇO (R$)"].sum()
    base_calculo_taxa_transacao = todos_com_taxa_transacao["TOTAL PAGO PELO CLIENTE (R$)"].sum()
    taxa_transacao_esperada = -(base_calculo_taxa_transacao * TAXA_TRANSACAO_PAGAMENTO_APP)

    # ── Serviços ──────────────────────────────────────────────────
    # Taxa de entrega cobrada nos pedidos Sob Demanda ON
    custo_logistico_sob_demanda = -todos_sob_demanda["TAXA DE ENTREGA PAGA PELO CLIENTE (R$)"].sum()

    # ── Promoções (apenas pedidos concluídos) ─────────────────────
    desconto_promocional_loja  = -pedidos_concluidos["INCENTIVO PROMOCIONAL DA LOJA (R$)"].sum()
    incentivo_promocional_ifood =  pedidos_concluidos["INCENTIVO PROMOCIONAL DO IFOOD (R$)"].sum()

    # ── Repasse ───────────────────────────────────────────────────
    total_recebido_direto_na_loja = concluidos_pagamento_direto["TOTAL PAGO PELO CLIENTE (R$)"].sum()
    total_repasse_ifood           = concluidos_repasse_ifood["VALOR LIQUIDO (R$)"].sum()

    return dict(
        valor_total_vendas                   = valor_total_vendas,
        comissao_cobrada_entrega_propria      = comissao_cobrada_entrega_propria,
        comissao_cobrada_entrega_flex         = comissao_cobrada_entrega_flex,
        comissao_cobrada_sob_demanda          = comissao_cobrada_sob_demanda,
        reembolso_comissao_flex_cancelados    = reembolso_comissao_flex_cancelados,
        taxa_transacao_cobrada               = taxa_transacao_cobrada,
        comissao_esperada_entrega_propria     = comissao_esperada_entrega_propria,
        comissao_esperada_entrega_flex        = comissao_esperada_entrega_flex,
        comissao_esperada_sob_demanda         = comissao_esperada_sob_demanda,
        taxa_transacao_esperada              = taxa_transacao_esperada,
        base_calculo_entrega_propria          = base_calculo_entrega_propria,
        base_calculo_entrega_flex             = base_calculo_entrega_flex,
        base_calculo_taxa_transacao           = base_calculo_taxa_transacao,
        custo_logistico_sob_demanda           = custo_logistico_sob_demanda,
        desconto_promocional_loja            = desconto_promocional_loja,
        incentivo_promocional_ifood          = incentivo_promocional_ifood,
        total_recebido_direto_na_loja        = total_recebido_direto_na_loja,
        total_repasse_ifood                  = total_repasse_ifood,
        qtd_pedidos_concluidos               = len(pedidos_concluidos),
        qtd_pedidos_cancelados               = len(pedidos_cancelados),
        qtd_pedidos_entrega_propria           = len(todos_entrega_propria),
        qtd_pedidos_entrega_flex              = len(todos_entrega_flex),
        qtd_pedidos_sob_demanda               = len(todos_sob_demanda),
        qtd_pedidos_direto_na_loja            = len(concluidos_pagamento_direto),
        qtd_pedidos_com_taxa_transacao        = len(todos_com_taxa_transacao),
    )


# ==========================================================
# IMPRESSÃO DO RELATÓRIO
# ==========================================================

def formatar_brl(valor: float) -> str:
    sinal = "-" if valor < 0 else " "
    return f"{sinal}R$ {abs(valor):>10,.2f}"

def imprimir_linha(label: str, valor: float, nota: str = "") -> None:
    print(f"  {label:<50} {formatar_brl(valor)}{nota}")

def imprimir_detalhe(label: str, complemento: str = "") -> None:
    print(f"    {label:<48} {complemento}")

def imprimir_cabecalho_secao(titulo: str) -> None:
    print(f"\n  {titulo}")
    print(f"  {'─' * 66}")


SEPARADOR = "=" * 70

df = carregar_planilha(CAMINHO_PLANILHA)
df = limpar_e_classificar_pedidos(df)

nomes_das_lojas = sorted(df["NOME DA LOJA"].dropna().astype(str).str.strip().unique())
divergencia_entrega_propria_total = 0.0

print(f"\n{SEPARADOR}")
print(f"  CONCILIAÇÃO DE FATURAMENTO iFood")
print(f"  Período: 01/06/2026 a 07/06/2026")
print(f"{SEPARADOR}")

for nome_loja in nomes_das_lojas:
    pedidos_loja = df[df["NOME DA LOJA"].astype(str).str.strip() == nome_loja]
    metricas = calcular_metricas_da_loja(pedidos_loja)

    divergencia_entrega_propria = (
        metricas["comissao_cobrada_entrega_propria"]
        - metricas["comissao_esperada_entrega_propria"]
    )
    divergencia_entrega_flex = (
        metricas["comissao_cobrada_entrega_flex"]
        - metricas["comissao_esperada_entrega_flex"]
    )
    divergencia_taxa_transacao = (
        metricas["taxa_transacao_cobrada"]
        + metricas["taxa_transacao_esperada"]
    )
    divergencia_entrega_propria_total += divergencia_entrega_propria

    tag_loja = nome_loja.split(" - ")[0]
    print(f"\n{'─' * 70}")
    print(f"  {tag_loja.upper()}   ({metricas['qtd_pedidos_concluidos']} concluídos | {metricas['qtd_pedidos_cancelados']} cancelados)")
    print(f"{'─' * 70}")

    # ── Valor das vendas ──────────────────────────────────────────
    imprimir_cabecalho_secao("VALOR DAS VENDAS")
    imprimir_linha("Valor dos itens + entrega própria", metricas["valor_total_vendas"])
    imprimir_detalhe("itens (concluídos) + taxa entrega SELF_DELIVERY")

    # ── Taxas e comissões ─────────────────────────────────────────
    imprimir_cabecalho_secao("TAXAS E COMISSÕES")

    flag_entrega_propria = "  ✓" if abs(divergencia_entrega_propria) < 1 else f"  *** Δ {divergencia_entrega_propria:+,.2f}"
    imprimir_linha(
        f"Comissão entrega própria ({metricas['qtd_pedidos_entrega_propria']} ped)",
        metricas["comissao_cobrada_entrega_propria"],
        flag_entrega_propria,
    )
    imprimir_detalhe(
        "base: VALOR_ITENS bruto (todos status)",
        f"R$ {metricas['base_calculo_entrega_propria']:>10,.2f}",
    )
    imprimir_detalhe(
        f"esperado {COMISSAO_ENTREGA_PROPRIA*100:.0f}% (Escalonada contrato)",
        formatar_brl(metricas["comissao_esperada_entrega_propria"]),
    )

    flag_entrega_flex = "  ✓" if abs(divergencia_entrega_flex) < 5 else f"  Δ {divergencia_entrega_flex:+,.2f}"
    imprimir_linha(
        f"Comissão entrega flex ({metricas['qtd_pedidos_entrega_flex']} ped)",
        metricas["comissao_cobrada_entrega_flex"],
        flag_entrega_flex,
    )
    imprimir_detalhe(
        "base: VALOR_ITENS bruto (todos status)",
        f"R$ {metricas['base_calculo_entrega_flex']:>10,.2f}",
    )
    imprimir_detalhe(
        f"esperado {COMISSAO_ENTREGA_FLEX*100:.0f}% (ref. contrato)",
        formatar_brl(metricas["comissao_esperada_entrega_flex"]),
    )
    imprimir_detalhe("ℹ  Flex inclui custo logístico variável (não reproduzível por %)")

    if metricas["reembolso_comissao_flex_cancelados"] != 0:
        imprimir_linha(
            "Reembolso comissão flex cancelados",
            metricas["reembolso_comissao_flex_cancelados"],
            "  ← TAXAS E COMISSOES pedidos FLEX cancelados",
        )

    if metricas["qtd_pedidos_sob_demanda"] > 0:
        divergencia_sob_demanda = (
            metricas["comissao_cobrada_sob_demanda"]
            - metricas["comissao_esperada_sob_demanda"]
        )
        flag_sob_demanda = "  ✓" if abs(divergencia_sob_demanda) < 1 else f"  Δ {divergencia_sob_demanda:+,.2f}"
        imprimir_linha(
            f"Comissão sob demanda ({metricas['qtd_pedidos_sob_demanda']} ped)",
            metricas["comissao_cobrada_sob_demanda"],
            flag_sob_demanda,
        )

    flag_taxa_transacao = "  ✓" if abs(divergencia_taxa_transacao) < 1 else f"  Δ {divergencia_taxa_transacao:+,.2f}"
    imprimir_linha(
        f"Taxa transação online ({metricas['qtd_pedidos_com_taxa_transacao']} ped)",
        metricas["taxa_transacao_cobrada"],
        flag_taxa_transacao,
    )
    imprimir_detalhe(
        "base: Pgto APP excl. VR/SODEXO/ALELO (todos status)",
        f"R$ {metricas['base_calculo_taxa_transacao']:>10,.2f}",
    )
    imprimir_detalhe(
        f"esperado {TAXA_TRANSACAO_PAGAMENTO_APP*100:.0f}% (contrato)",
        formatar_brl(metricas["taxa_transacao_esperada"]),
    )

    # ── Serviços ──────────────────────────────────────────────────
    imprimir_cabecalho_secao("SERVIÇOS")
    print(f"  {'Pacote de anúncios':<50} {'[não conciliável via pedidos]':>22}")
    if metricas["custo_logistico_sob_demanda"] != 0:
        imprimir_linha(
            f"Solicitação sob demanda ({metricas['qtd_pedidos_sob_demanda']} ped)",
            metricas["custo_logistico_sob_demanda"],
            "  ← taxa entrega SOB DEMANDA ON",
        )
    else:
        print(f"  {'Solicitação sob demanda':<50} {'R$       0,00':>14}  ← sem pedidos")

    # ── Promoções ─────────────────────────────────────────────────
    imprimir_cabecalho_secao("PROMOÇÕES")
    imprimir_linha(
        "Incentivadas pela loja",
        metricas["desconto_promocional_loja"],
        "  ← INCENTIVO PROMOCIONAL DA LOJA (concluídos)",
    )
    imprimir_linha(
        "Incentivadas pelo iFood",
        metricas["incentivo_promocional_ifood"],
        "  ← informativo (custo iFood)",
    )

    # ── Total do repasse ──────────────────────────────────────────
    imprimir_cabecalho_secao("TOTAL DO REPASSE")
    imprimir_linha(
        f"Recebido direto na loja ({metricas['qtd_pedidos_direto_na_loja']} ped)",
        metricas["total_recebido_direto_na_loja"],
        "  ← Dinheiro + Pgto na Entrega + VR/SODEXO/ALELO APP",
    )
    imprimir_linha(
        "Valor do repasse iFood",
        metricas["total_repasse_ifood"],
        "  ← VALOR LIQUIDO pedidos não-diretos concluídos",
    )

    # ── Alerta de divergência ──────────────────────────────────────
    if abs(divergencia_entrega_propria) > 1:
        print(f"\n  {'─' * 66}")
        print(f"  ⚠  COBRANÇA INDEVIDA — Entrega Própria:  R$ {-divergencia_entrega_propria:,.2f}")
        print(
            f"     {COMISSAO_ENTREGA_PROPRIA*100:.0f}% contratado vs cobrado "
            f"sobre R$ {metricas['base_calculo_entrega_propria']:,.2f}"
        )

# ── Resumo final ───────────────────────────────────────────────────
print(f"\n{SEPARADOR}")
print(f"  RESUMO — DIVERGÊNCIAS DA SEMANA (3 lojas)")
print(f"{'─' * 70}")
if abs(divergencia_entrega_propria_total) > 1:
    print(f"  ⚠  Comissão própria cobrada a mais:   R$ {-divergencia_entrega_propria_total:>10,.2f}")
    print(f"     Projeção mensal  (~4 semanas):      R$ {-divergencia_entrega_propria_total*4:>10,.2f}")
    print(f"     Projeção anual   (~52 semanas):     R$ {-divergencia_entrega_propria_total*52:>10,.2f}")
else:
    print(f"  ✓  Comissões dentro do esperado contratual.")
print(f"{SEPARADOR}\n")