# Advanced Sales Analysis — Portuguese Market & Product Performance

> Multi-dimensional statistical analysis of sales performance data, delivering 12 interactive visualisations and publication-quality reports across product, geographic, price, and correlation dimensions.

[![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org)
[![Plotly](https://img.shields.io/badge/Plotly-3F4F75?style=for-the-badge&logo=plotly&logoColor=white)](https://plotly.com)
[![SciPy](https://img.shields.io/badge/SciPy-8CAAE6?style=for-the-badge&logo=scipy&logoColor=white)](https://scipy.org)

---

## Overview

This project performs a comprehensive statistical analysis of a multi-country sales dataset, with a primary focus on the **Portuguese market**. It answers five core business questions through rigorous statistical methods and produces outputs suitable for executive reporting, stakeholder presentations, and academic portfolios.

Two scripts are provided:
| Script | Output | Use Case |
|---|---|---|
| `enhanced_sales_analysis.py` | Static PNG plots (300 DPI) + Excel | Formal reports, publications |
| `enhanced_sales_analysis_interactive.py` | Interactive HTML plots + Excel | Web portfolios, dashboards |

---

## Business Questions Answered

### Q1 — Product Performance in Portugal
Which products sell most in the Portuguese market, and how consistent is that performance?

- Ranks all products by total quantity sold in Portugal
- Calculates **coefficient of variation** to flag volatile performers
- Identifies revenue vs. volume mismatches

### Q2 — Extreme Performers: Top vs. Bottom
What separates the best-selling products from the worst?

- Side-by-side comparison of **Top 5 vs. Bottom 5** products
- Visualises the gap in units sold and revenue contribution
- Highlights portfolio imbalance for strategic review

### Q3 — Geographic Distribution
How does demand for a flagship product distribute across markets?

- Maps sales volume by country with **95% confidence intervals**
- Applies **Pareto (80/20) analysis** to identify key revenue-driving markets
- Reveals geographic concentration risk

### Q4 — Price Variation by Country
Does pricing differ significantly across markets for the same product?

- Box plots and violin plots showing price distribution shape by country
- Calculates standard deviation, IQR, and confidence intervals per market
- Exports full descriptive statistics to Excel for stakeholder use

### Q5 — Quantity-Price Correlation
Is there a statistically significant relationship between volume and price?

- Computes **Pearson** (linear) and **Spearman** (rank-based) correlation coefficients
- Performs significance testing (p-values) across markets
- Visualises relationships with scatter plots and regression lines

---

## Visualisations Produced (12 Total)

| # | Plot | Type |
|---|---|---|
| 1 | Top 15 Products by Quantity — Portugal | Bar chart |
| 2 | Mean vs. Median Sales Comparison | Grouped bar chart |
| 3 | Portfolio Performance Matrix | Bubble chart |
| 4 | Top 5 vs. Bottom 5 Performers | Comparative bar |
| 5 | Sales Distribution Analysis | Histogram + KDE |
| 6 | Geographic Performance by Country | Bar with CI |
| 7 | Pareto Analysis — Key Markets | Pareto chart |
| 8 | Unit Price Distribution by Country | Box plots |
| 9 | Unit Price Density by Country | Violin plots |
| 10 | Mean Price with Confidence Intervals | Error bar chart |
| 11 | Quantity vs. Price Scatter Regression | Scatter + regression |
| 12 | Correlation Heatmap | Heatmap |

---

## Technical Stack

| Component | Technology |
|---|---|
| Data manipulation | `pandas`, `numpy` |
| Statistical analysis | `scipy` (pearsonr, spearmanr, skew, kurtosis) |
| Static visualisation | `matplotlib`, `seaborn` |
| Interactive visualisation | `plotly.express`, `plotly.graph_objects` |
| Export | `openpyxl` (Excel), HTML |

---

## Project Structure

```
Sales-Analysis/
├── enhanced_sales_analysis.py            # Static outputs (matplotlib/seaborn)
├── enhanced_sales_analysis_interactive.py # Interactive outputs (Plotly)
├── data/
│   └── online_retail.xlsx                # Source dataset
└── outputs/
    ├── plots/                            # PNG exports (300 DPI)
    ├── interactive/                      # HTML interactive plots
    └── reports/                          # Excel statistical summaries
```

---

## Getting Started

### Prerequisites

```bash
pip install pandas numpy matplotlib seaborn scipy plotly openpyxl
```

### Run Static Analysis

```bash
python enhanced_sales_analysis.py
```

Generates 12 PNG plots at 300 DPI resolution and Excel summary reports.

### Run Interactive Analysis

```bash
python enhanced_sales_analysis_interactive.py
```

Generates 12 interactive HTML visualisations with hover, zoom, pan, and export capabilities.

---

## Dataset

The analysis uses a transactional retail dataset containing sales records across multiple countries. The Portuguese market subset is the primary focus, with cross-country comparisons for geographic and price analysis.

**Key fields used:** `Description`, `Quantity`, `UnitPrice`, `Country`, `InvoiceDate`, `CustomerID`

---

## Statistical Methods

| Method | Application |
|---|---|
| Coefficient of Variation | Measuring sales consistency across products |
| Confidence Intervals (95%) | Geographic and price comparison reliability |
| Pearson Correlation | Linear relationship between quantity and price |
| Spearman Correlation | Rank-based relationship (robust to outliers) |
| Pareto Analysis | 80/20 identification of key markets |
| Skewness & Kurtosis | Distribution shape characterisation |

---

## Outputs

| File | Description |
|---|---|
| `Plot_1_Top_15_Products.html` | Interactive top product ranking |
| `Plot_6_Geographic_Performance.html` | Country-level demand map |
| `Plot_7_Pareto_Analysis.html` | 80/20 market concentration |
| `Plot_12_Correlation_Heatmap.html` | Full correlation matrix |
| `products_portugal_advanced.xlsx` | Portuguese market statistics |
| `unit_price_summary_stats.xlsx` | Price analysis by country |
| `correlation_analysis_stats.xlsx` | Correlation coefficients and p-values |

---

## Author

**Ailton D'Alcântara** — Data Engineer & Analytics Engineer  
📍 Olten, Switzerland · [LinkedIn](https://linkedin.com/in/ailton-d-alcantara-3a4681195) · [Portfolio](https://succulent-warlock-c69.notion.site/Hey-there-I-m-Ailton-D-Alcantara-1d194e30fb42802991fed77a1de216d6)
