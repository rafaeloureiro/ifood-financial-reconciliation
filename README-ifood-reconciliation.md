# 💰 ifood-reconciliation

Weekly financial reconciliation pipeline for iFood operations.

Reads the order-level report exported from the iFood Partner Portal and validates gross revenue, commissions, delivery fees, promotional discounts, and net payout against the values shown on the iFood settlement screen — catching any discrepancy before it goes unnoticed.

---

## 🏪 Context

Three delivery brands operating on iFood in Santos, Brazil:

| Brand | iFood Store |
|---|---|
| O Burguês | – |
| Ex-Touro | Ex-touro - Burger Santos |
| Seu Vidal | – |

The reconciliation runs **weekly**, covering Sunday through Saturday, using the `relatorio-pedidos` XLSX export as its single source of truth.

---

## 🧾 What It Reconciles

The pipeline derives each line of the iFood settlement screen from order-level data and flags any mismatch:

| Settlement Item (Portal iFood) | Derived From |
|---|---|
| Valor dos itens e entrega própria | Sum of `VALOR DOS ITENS` — concluded orders |
| Reembolso de pedidos cancelados | `VALOR LÍQUIDO` of cancelled orders |
| Comissão iFood – entrega própria | `TAXAS E COMISSOES` where logistics = `SELF_DELIVERY_PARTIAL_AREA` |
| Comissão iFood – entrega parceira | `TAXAS E COMISSOES` where logistics = `ENTREGA FLEX` |
| Taxa de transação de pagamento online | Sum of `TAXA DE SERVIÇO` — online payment orders |
| Incentivos da loja | Sum of `INCENTIVO PROMOCIONAL DA LOJA` |
| Incentivos do iFood | Sum of `INCENTIVO PROMOCIONAL DO IFOOD` |
| Valor do repasse | Sum of `VALOR LÍQUIDO` — concluded, online payment orders |

---

## 📁 Project Structure

```
ifood-reconciliation/
├── data/
│   └── raw/                    # Weekly XLSX files from iFood (not versioned)
├── reconciliation/
│   ├── __init__.py
│   ├── loader.py               # Reads and validates the XLSX input
│   ├── reconciler.py           # Derives settlement totals from order data
│   └── report.py               # Prints formatted reconciliation summary
├── utils/
│   ├── __init__.py
│   ├── constants.py            # Column name aliases, status/logistics values
│   └── formatting.py           # Brazilian currency and percentage formatting
├── tests/
│   ├── __init__.py
│   └── test_reconciler.py
├── main.py                     # Entry point
├── requirements.txt
├── .gitignore
└── README.md
```

---

## 📄 Input

**Weekly order report** exported from the iFood Partner Portal.

File naming pattern:
```
relatorio-pedidos_<hash>_<YYYY-MM-DD>-<YYYY-MM-DD>.xlsx
```

Place the file in `data/raw/` before running.

### Columns used

| Column | Description |
|---|---|
| `ID COMPLETO DO PEDIDO` | Unique order UUID |
| `STATUS FINAL DO PEDIDO` | `CONCLUIDO` or `CANCELADO` |
| `VALOR DOS ITENS (R$)` | Gross item value |
| `TAXA DE SERVIÇO (R$)` | iFood service fee charged to customer |
| `TAXAS E COMISSOES (R$)` | Total commissions and fees (negative) |
| `VALOR LIQUIDO (R$)` | Net value to store after all deductions |
| `INCENTIVO PROMOCIONAL DO IFOOD (R$)` | iFood-funded discount |
| `INCENTIVO PROMOCIONAL DA LOJA (R$)` | Store-funded discount |
| `PRODUTO LOGISTICO` | `SELF_DELIVERY_PARTIAL_AREA` or `ENTREGA FLEX` |
| `FORMA DE PAGAMENTO` | Payment method — used to identify direct payments |

---

## 🚀 Usage

```bash
pip install -r requirements.txt

python main.py --file data/raw/relatorio-pedidos_<hash>_2026-06-01-2026-06-07.xlsx
```

### Example output

```
============================================================
  CONCILIAÇÃO FINANCEIRA iFood — 2026-06-01 a 2026-06-07
============================================================

Pedidos concluídos          326
Pedidos cancelados            2

VALOR DAS VENDAS
  Valor dos itens (bruto)    R$ 19.412,84
  Reembolso de cancelados        R$ 31,61

TAXAS E COMISSÕES
  Comissão entrega própria   -R$ 1.463,05
  Comissão entrega parceira    -R$ 219,48
  Taxa de transação online     -R$ 404,29

PROMOÇÕES
  Incentivo da loja          -R$ 1.955,24
  Incentivo do iFood           R$ 825,79

TOTAL LÍQUIDO (VALOR LÍQUIDO)  R$ 11.454,37
============================================================
```

---

## 🛠️ Requirements

```
pandas>=2.0
openpyxl>=3.1
tabulate>=0.9
```

---

## 📌 Notes

- Raw XLSX files are excluded from version control via `.gitignore`.
- Cancelled orders are excluded from commission and payout totals but included in reimbursement reconciliation.
- Orders paid with cash, VR, or directly at the store (`Pgto na Entrega`, `Dinheiro`) are excluded from the iFood repasse total.
- Monetary values use Brazilian formatting throughout: `R$ 1.234,56`.

---

## 🔗 Related

- [`ifood-sales-analytics`](https://github.com/seu-usuario/ifood-sales-analytics) — weekly KPIs and revenue breakdown using the same data source.
