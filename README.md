# TCS Financial Model — Three-Statement + DCF Valuation

![Excel](https://img.shields.io/badge/Excel-12_Sheets-217346?logo=microsoftexcel)
![Python](https://img.shields.io/badge/Python-Automation-3776AB?logo=python)
![Sheets](https://img.shields.io/badge/Model-3_Statement-0D2137)
![Status](https://img.shields.io/badge/Status-Complete-success)

> Industry-ready financial model for **Tata Consultancy Services (NSE: TCS)**.
> Covers FY2011–FY2025 historical (15 years) and FY2026E–FY2030E forecast (5 years).
> Includes corrected FCFF DCF, full Cash Flow Statement, computed sensitivity
> heat maps, peer comps, Football Field valuation, and a Python automation layer.

## Key findings

| Metric | Value |
|--------|-------|
| Base case IV per share | ₹2,397 |
| Bull case IV per share | ₹3,938 |
| Bear case IV per share | ₹1,787 |
| Current market price (CMP) | ₹3,350 |
| Base case implied downside | (28.4%) |
| 15-year revenue CAGR (FY11–FY25, USD) | 9.8% |
| FY25 EBIT margin | 24.3% |
| FY25 EBITDA margin | 26.4% |
| FY25 Net cash position | USD 5,527 Mn |
| FY25 Return on equity | 50.6% |
| WACC (CAPM, India Rf) | 12.89% |
| Terminal growth rate | 4.5% |

## What was audited and corrected (v1.0 → v2.0)

Three critical errors were found and corrected from the original model:

- **FCFF bug** — Free cash flow equalled NOPAT. D&A was not added back and
  Capex / ΔWC were not deducted. Fixed: FCFF = NOPAT + D&A − Capex − ΔWC
- **Year label bug** — Forecast was labelled FY27E–FY31E but actually represented
  FY26E–FY30E (one year off throughout). Fixed: corrected to FY26E–FY30E
- **WACC methodology** — India G-Sec Rf used for a USD model with no explanation.
  Fixed: methodology note added; dual approach shown in Sensitivity sheet

Full 13-item audit log is documented in the Audit_Log sheet.

## Project structure

```
tcs-financial-model/
├── TCS_Financial_Model_v2.xlsx     ← Main model (12 sheets)
├── data/
│   └── TCS-Data-Sheet-Q2FY26.xlsx ← Source: TCS Q2 FY26 investor data
├── python/
│   ├── tcs_model_refresh.py        ← Scenario DCF + sensitivity + charts
│   └── requirements.txt
└── README.md
```

## Excel model — 12 sheets

| Sheet | Purpose |
|-------|---------|
| Cover | Metadata, navigation, colour code guide |
| Assumptions | All forecast drivers — Bear / Base / Bull for every input |
| Income_Statement | Full P&L FY11–FY25 + FY26E–FY30E forecast |
| Cash_Flow | **NEW** — Indirect method CF: Operations, Investing, Financing |
| Balance_Sheet | Historical BS FY11–FY25 with balance check and ROE |
| Revenue_Decomp | Geographic mix, employee metrics, CAGR analysis |
| Margin_Analysis | EBIT margin bridge, cost waterfall FY11–FY25 |
| DCF_Valuation | Corrected FCFF DCF — WACC, FCF, TV, equity bridge |
| Sensitivity | **Computed** heat maps — WACC×TGR and RevCAGR×Margin |
| Comps | Peer trading multiples — NTM labelled, 6 peers |
| Football_Field | Valuation range summary — formula-linked values |
| Audit_Log | 13-item issue tracker — all v1.0 errors documented |

## Python automation

Runs Bear / Base / Bull DCF, computes two sensitivity tables,
outputs CSV files, and generates a matplotlib chart.

```bash
cd python
pip install -r requirements.txt
python tcs_model_refresh.py
```

**Outputs generated:**
- Scenario summary table (printed to console)
- `tcs_scenarios_YYYYMMDD.csv` — all scenario outputs
- `tcs_sensitivity_wacc_tgr_YYYYMMDD.csv` — WACC × TGR table
- `tcs_model_output.png` — scenario bar chart + sensitivity heat map

## Methodology

- **Model type:** Three-statement (Income Statement + Balance Sheet + Cash Flow)
- **Valuation:** FCFF DCF + EV/EBITDA comps + P/E comps + Football Field
- **Historical period:** FY2011–FY2025 (15 years, IFRS, USD Millions)
- **Forecast period:** FY2026E–FY2030E (5 years)
- **WACC:** 12.89% (CAPM: Rf 7.1% India G-Sec + β 0.85 × ERP 7.0%)
- **TGR:** 4.5% (India IT sector long-run nominal; WACC–g spread = 8.4%)
- **Mid-year convention** used for discount factors
- **Source:** TCS Q2 FY2026 Earnings Data Sheet (IFRS)

## Tools used

Microsoft Excel · Python · pandas · matplotlib · openpyxl
Financial Modelling · Equity Research · Three-Statement Modelling

---
*For educational and portfolio purposes only. Not investment advice.
Consult a SEBI-registered advisor before investing.*
