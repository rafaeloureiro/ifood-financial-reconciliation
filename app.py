import pandas as pd
import openpyxl

# ==========================================================
# CONFIGURAÇÃO
# ==========================================================

ARQUIVO = "Data/pedidos.xlsx"

# Taxas contratuais
TAXA_COMISSAO_ESCALONADA   = 0.09     # Entrega própria — base: VALOR_ITENS bruto
TAXA_COMISSAO_FLEX         = 0.20     # Entrega Flex
TAXA_COMISSAO_SOB_DEMANDA  = 0.09     # Sob Demanda ON
TAXA_COMISSAO_RETIRADA     = 0.0875   # Pra Retirar
TAXA_TRANSACAO_ONLINE      = 0.03     # Pgto via APP — exceto VR/SODEXO/ALELO via APP

# Pagamentos com taxa de transação ZERO (iFood isenta)
# IFOOD MEAL VOUCHER e Outros vales NÃO estão aqui — pagam 3% e entram no repasse
ISENCAO_TAXA_TRANSACAO = [
    "Pgto via APP - Vale Refeição (VR)",
    "Pgto via APP - Vale Refeição (SODEXO)",
    "Pgto via APP - Vale Refeição (ALELO)",
]

# Pagamentos recebidos DIRETO NA LOJA (não entram no repasse iFood)
# Regra: Dinheiro + qualquer "Pgto na Entrega" + VR/SODEXO/ALELO via APP
PAGAMENTOS_DIRETO_LOJA = [
    "Dinheiro",
    "Outros vales",
    "Pgto na Entrega",                          # cobre todos os subtipos via str.contains
    "Pgto via APP - Vale Refeição (VR)",
    "Pgto via APP - Vale Refeição (SODEXO)",
    "Pgto via APP - Vale Refeição (ALELO)",
]

# ==========================================================
# LEITURA E LIMPEZA
# ==========================================================

with open(ARQUIVO, "rb") as f:
    wb = openpyxl.load_workbook(f)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    rows = [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]

df = pd.DataFrame(rows)
df.columns = df.columns.str.strip()

COLUNAS_MONETARIAS = [
    "VALOR DOS ITENS (R$)",
    "TAXA DE ENTREGA PAGA PELO CLIENTE (R$)",
    "TOTAL PAGO PELO CLIENTE (R$)",
    "INCENTIVO PROMOCIONAL DO IFOOD (R$)",
    "INCENTIVO PROMOCIONAL DA LOJA (R$)",
    "INCENTIVO PROMOCIONAL DA REDE (R$)",
    "TAXA DE SERVIÇO (R$)",
    "TAXAS E COMISSOES (R$)",
    "VALOR LIQUIDO (R$)",
]
for col in COLUNAS_MONETARIAS:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

df["PRODUTO LOGISTICO"]      = df["PRODUTO LOGISTICO"].fillna("").str.strip()
df["FORMA DE PAGAMENTO"]     = df["FORMA DE PAGAMENTO"].fillna("").str.strip()
df["STATUS FINAL DO PEDIDO"] = df["STATUS FINAL DO PEDIDO"].fillna("").str.strip()

# Flags de status
df["_concluido"] = df["STATUS FINAL DO PEDIDO"].isin(["CONCLUIDO", "CANCELAMENTO PARCIAL"])
df["_cancelado"] = df["STATUS FINAL DO PEDIDO"] == "CANCELADO"

# Flags de logística
df["_own_delivery"] = df["PRODUTO LOGISTICO"] == "SELF_DELIVERY_PARTIAL_AREA"
df["_flex"]         = df["PRODUTO LOGISTICO"] == "ENTREGA FLEX"
df["_sob_demanda"]  = df["PRODUTO LOGISTICO"] == "SOB DEMANDA ON"
df["_retirada"]     = df["PRODUTO LOGISTICO"] == "RETIRADA"

# Flag: isento de taxa de transação (VR/SODEXO/ALELO via APP)
df["_isento_taxa"] = df["FORMA DE PAGAMENTO"].isin(ISENCAO_TAXA_TRANSACAO)

# Flag: pago via APP (inclui IFOOD MEAL VOUCHER — paga 3% de transação)
df["_pgto_app"] = df["FORMA DE PAGAMENTO"].str.contains("Pgto via APP", case=False, na=False)

# Base da taxa de transação: APP exceto isentos
df["_base_transacao"] = df["_pgto_app"] & ~df["_isento_taxa"]

# Flag: pagamento direto na loja (não entra no repasse iFood)
def eh_direto(pgto):
    return (
        pgto == "Dinheiro"
        or pgto == "Outros vales"
        or pgto.startswith("Pgto na Entrega")
        or pgto in [
            "Pgto via APP - Vale Refeição (VR)",
            "Pgto via APP - Vale Refeição (SODEXO)",
            "Pgto via APP - Vale Refeição (ALELO)",
        ]
    )

df["_direto_loja"] = df["FORMA DE PAGAMENTO"].apply(eh_direto)

# ==========================================================
# CÁLCULO POR LOJA
# ==========================================================

def calcular_loja(df_loja):
    c  = df_loja[df_loja["_concluido"]]
    ca = df_loja[df_loja["_cancelado"]]

    own      = c[c["_own_delivery"]]
    flex     = c[c["_flex"]]
    sob      = c[c["_sob_demanda"]]
    flex_canc = ca[ca["_flex"]]

    base_transacao = c[c["_base_transacao"]]
    diretos        = c[c["_direto_loja"]]
    repasse_orders = c[~c["_direto_loja"]]

    # ── Valor das vendas ──────────────────────────────────────────
    valor_itens_entrega = (
        c["VALOR DOS ITENS (R$)"].sum()
        + own["TAXA DE ENTREGA PAGA PELO CLIENTE (R$)"].sum()
    )

    # ── Comissões cobradas (CSV) ──────────────────────────────────
    comissao_own_cobrada  = own["TAXAS E COMISSOES (R$)"].sum()
    comissao_flex_cobrada = flex["TAXAS E COMISSOES (R$)"].sum()
    comissao_sob_cobrada  = sob["TAXAS E COMISSOES (R$)"].sum()
    reembolso_flex_canc   = flex_canc["TAXAS E COMISSOES (R$)"].sum()

    # ── Comissões esperadas (contrato) ────────────────────────────
    # Base confirmada pelo portal: VALOR_ITENS bruto (sem descontar promoções)
    base_own  = own["VALOR DOS ITENS (R$)"].sum()
    base_flex = flex["VALOR DOS ITENS (R$)"].sum()
    base_sob  = sob["VALOR DOS ITENS (R$)"].sum()

    comissao_own_esp  = -(base_own  * TAXA_COMISSAO_ESCALONADA)
    comissao_flex_esp = -(base_flex * TAXA_COMISSAO_FLEX)
    comissao_sob_esp  = -(base_sob  * TAXA_COMISSAO_SOB_DEMANDA)

    # ── Taxa de transação ─────────────────────────────────────────
    taxa_trans_cobrada  = c["TAXA DE SERVIÇO (R$)"].sum()
    base_trans_valor    = base_transacao["TOTAL PAGO PELO CLIENTE (R$)"].sum()
    taxa_trans_esperada = -(base_trans_valor * TAXA_TRANSACAO_ONLINE)

    # ── Serviços ──────────────────────────────────────────────────
    sob_demanda_servico = -sob["TAXA DE ENTREGA PAGA PELO CLIENTE (R$)"].sum()

    # ── Promoções ─────────────────────────────────────────────────
    promo_loja  = -c["INCENTIVO PROMOCIONAL DA LOJA (R$)"].sum()
    promo_ifood =  c["INCENTIVO PROMOCIONAL DO IFOOD (R$)"].sum()

    # ── Repasse ───────────────────────────────────────────────────
    direto_loja   = diretos["TOTAL PAGO PELO CLIENTE (R$)"].sum()
    valor_repasse = repasse_orders["VALOR LIQUIDO (R$)"].sum()

    return dict(
        valor_itens_entrega     = valor_itens_entrega,
        comissao_own_cobrada    = comissao_own_cobrada,
        comissao_flex_cobrada   = comissao_flex_cobrada,
        comissao_sob_cobrada    = comissao_sob_cobrada,
        reembolso_flex_canc     = reembolso_flex_canc,
        taxa_trans_cobrada      = taxa_trans_cobrada,
        comissao_own_esp        = comissao_own_esp,
        comissao_flex_esp       = comissao_flex_esp,
        comissao_sob_esp        = comissao_sob_esp,
        taxa_trans_esperada     = taxa_trans_esperada,
        base_own                = base_own,
        base_flex               = base_flex,
        base_trans_valor        = base_trans_valor,
        sob_demanda_servico     = sob_demanda_servico,
        promo_loja              = promo_loja,
        promo_ifood             = promo_ifood,
        direto_loja             = direto_loja,
        valor_repasse           = valor_repasse,
        n_concluidos            = len(c),
        n_cancelados            = len(ca),
        n_own                   = len(own),
        n_flex                  = len(flex),
        n_sob                   = len(sob),
        n_diretos               = len(diretos),
        n_base_transacao        = len(base_transacao),
    )

# ==========================================================
# IMPRESSÃO DO RELATÓRIO
# ==========================================================

def brl(v):
    sinal = "-" if v < 0 else " "
    return f"{sinal}R$ {abs(v):>10,.2f}"

def linha(label, valor, nota=""):
    print(f"  {label:<50} {brl(valor)}{nota}")

def sublinha(label, detalhe=""):
    print(f"    {label:<48} {detalhe}")

def secao(titulo):
    print(f"\n  {titulo}")
    print(f"  {'─' * 66}")

SEP = "=" * 70
lojas = sorted(df["NOME DA LOJA"].dropna().astype(str).str.strip().unique())
totais = dict(div_own=0.0)

print(f"\n{SEP}")
print(f"  CONCILIAÇÃO DE FATURAMENTO iFood")
print(f"  Período: 01/06/2026 a 07/06/2026")
print(f"{SEP}")

for loja in lojas:
    df_loja = df[df["NOME DA LOJA"].astype(str).str.strip() == loja]
    r = calcular_loja(df_loja)

    div_own  = r["comissao_own_cobrada"]  - r["comissao_own_esp"]
    div_flex = r["comissao_flex_cobrada"] - r["comissao_flex_esp"]
    div_taxa = r["taxa_trans_cobrada"]    - (-r["taxa_trans_esperada"])
    totais["div_own"] += div_own

    tag = loja.split(" - ")[0]
    print(f"\n{'─' * 70}")
    print(f"  {tag.upper()}   ({r['n_concluidos']} concluídos | {r['n_cancelados']} cancelados)")
    print(f"{'─' * 70}")

    # ── Valor das vendas ──────────────────────────────────────────
    secao("VALOR DAS VENDAS")
    linha("Valor dos itens + entrega própria", r["valor_itens_entrega"])
    sublinha("itens (concluídos) + taxa entrega SELF_DELIVERY")

    # ── Taxas e comissões ─────────────────────────────────────────
    secao("TAXAS E COMISSÕES")

    flag_own = "  ✓" if abs(div_own) < 1 else f"  *** Δ {div_own:+,.2f}"
    linha(f"Comissão entrega própria ({r['n_own']} ped)",
          r["comissao_own_cobrada"], flag_own)
    sublinha(f"base: VALOR_ITENS bruto",
             f"R$ {r['base_own']:>10,.2f}")
    sublinha(f"esperado {TAXA_COMISSAO_ESCALONADA*100:.0f}% (Escalonada contrato)",
             brl(r["comissao_own_esp"]))

    flag_flex = "  ✓" if abs(div_flex) < 5 else f"  Δ {div_flex:+,.2f}"
    linha(f"Comissão entrega flex ({r['n_flex']} ped)",
          r["comissao_flex_cobrada"], flag_flex)
    sublinha(f"base: VALOR_ITENS bruto",
             f"R$ {r['base_flex']:>10,.2f}")
    sublinha(f"esperado {TAXA_COMISSAO_FLEX*100:.0f}% (ref. contrato)",
             brl(r["comissao_flex_esp"]))
    sublinha("ℹ  Flex inclui custo logístico variável (não reproduzível por %)")

    if r["n_sob"] > 0:
        div_sob = r["comissao_sob_cobrada"] - r["comissao_sob_esp"]
        flag_sob = "  ✓" if abs(div_sob) < 1 else f"  Δ {div_sob:+,.2f}"
        linha(f"Comissão sob demanda ({r['n_sob']} ped)",
              r["comissao_sob_cobrada"], flag_sob)

    if r["reembolso_flex_canc"] != 0:
        linha("Reembolso comissão flex (cancelados)",
              r["reembolso_flex_canc"],
              "  ← TAXAS E COMISSOES pedidos FLEX cancelados")

    flag_taxa = "  ✓" if abs(div_taxa) < 1 else f"  Δ {div_taxa:+,.2f}"
    linha(f"Taxa transação online ({r['n_base_transacao']} ped)",
          r["taxa_trans_cobrada"], flag_taxa)
    sublinha(f"base: Pgto APP (excl. VR/SODEXO/ALELO via APP)",
             f"R$ {r['base_trans_valor']:>10,.2f}")
    sublinha(f"esperado {TAXA_TRANSACAO_ONLINE*100:.0f}% (contrato)",
             brl(r["taxa_trans_esperada"]))

    # ── Serviços ──────────────────────────────────────────────────
    secao("SERVIÇOS")
    print(f"  {'Pacote de anúncios':<50} {'[não conciliável via pedidos]':>22}")
    if r["sob_demanda_servico"] != 0:
        linha(f"Solicitação sob demanda ({r['n_sob']} ped)",
              r["sob_demanda_servico"],
              "  ← taxa entrega pedidos SOB DEMANDA ON")
    else:
        print(f"  {'Solicitação sob demanda':<50} {'R$       0,00':>14}  ← sem pedidos")

    # ── Promoções ─────────────────────────────────────────────────
    secao("PROMOÇÕES")
    linha("Incentivadas pela loja",  r["promo_loja"],
          "  ← INCENTIVO PROMOCIONAL DA LOJA")
    linha("Incentivadas pelo iFood", r["promo_ifood"],
          "  ← informativo (custo iFood, não da loja)")

    # ── Total do repasse ──────────────────────────────────────────
    secao("TOTAL DO REPASSE")
    linha(f"Recebido direto na loja ({r['n_diretos']} ped)",
          r["direto_loja"],
          "  ← Dinheiro + Pgto na Entrega + VR/SODEXO/ALELO APP")
    linha("Valor do repasse iFood",
          r["valor_repasse"],
          "  ← VALOR LIQUIDO pedidos não-diretos")

    # ── Divergência da loja ───────────────────────────────────────
    if abs(div_own) > 1:
        print(f"\n  {'─' * 66}")
        print(f"  ⚠  COBRANÇA INDEVIDA — Entrega Própria:  R$ {-div_own:,.2f}")
        print(f"     {TAXA_COMISSAO_ESCALONADA*100:.0f}% esperado vs cobrado sobre "
              f"R$ {r['base_own']:,.2f} (VALOR_ITENS bruto)")

# ── Resumo geral ──────────────────────────────────────────────────
print(f"\n{SEP}")
print(f"  RESUMO — DIVERGÊNCIAS DA SEMANA (3 lojas)")
print(f"{'─' * 70}")
if abs(totais["div_own"]) > 1:
    print(f"  ⚠  Comissão própria cobrada a mais:   R$ {-totais['div_own']:>10,.2f}")
    print(f"     Projeção mensal  (~4 semanas):      R$ {-totais['div_own']*4:>10,.2f}")
    print(f"     Projeção anual   (~52 semanas):     R$ {-totais['div_own']*52:>10,.2f}")
else:
    print(f"  ✓  Comissões dentro do esperado contratual.")
print(f"{SEP}\n")
