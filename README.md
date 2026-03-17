# Multi-Currency Excel Automation

A Python script that takes messy, multi-currency expense exports from multiple sources — Stripe, QuickBooks, internal ledgers — cleans them automatically, normalizes everything to USD using live FX rates, and outputs a polished, investor-ready Excel report in under 15 seconds.

Replaces 8 hours of manual monthly work. Permanently.

---

## Live Demo

> [View the before & after walkthrough](https://aryanture1000-prog.github.io/excel-automation-multicurrency)

---

## The Problem It Solves

Every finance team running multi-currency operations deals with this every month:

- Date formats inconsistent across source systems (MM/DD, DD-MM, text)
- Currency symbols mixed into amount columns ($, £, €)
- No FX normalization — USD and EUR figures sitting side by side unreconciled
- Duplicate entries from overlapping exports
- Categories named differently across Stripe, QuickBooks, and internal tools
- No budget vs actual comparison, no overspend flags

This script fixes all of it automatically.

---

## What It Does

```
Input:   3 messy CSV exports (Stripe + QuickBooks + Ledger)
         Mixed currencies: USD, EUR, GBP
         Inconsistent dates, symbols, categories, duplicates

Output:  1 clean Excel report with 2 sheets:
         Sheet 1 — Clean Data (normalized, deduplicated, categorized)
         Sheet 2 — Summary (budget vs actual, overspend flags, all in USD)

Runtime: ~11 seconds
```

---

## Features

- **Live FX conversion** — fetches real-time exchange rates at runtime, not stale manual lookups
- **Multi-source merging** — combines any number of CSV files in one pass
- **Date normalization** — handles any date format automatically
- **Category standardization** — maps inconsistent labels to a unified taxonomy
- **Duplicate removal** — deduplicates across overlapping source exports
- **Budget vs actual** — compares actuals to preset budgets and flags overspend automatically
- **Audit trail** — original currency and amount preserved alongside USD conversion

---

## Tech Stack

- Python 3.10+
- pandas — data cleaning and transformation
- requests — live FX rate fetching
- openpyxl — Excel report generation
- exchangerate-api — FX data source

---

## How to Use

```bash
git clone https://github.com/aryanture1000-prog/excel-automation-multicurrency.git
cd excel-automation-multicurrency
pip install -r requirements.txt
```

Drop your CSV files into the `/data` folder, set your budget targets in `config.py`, then run:

```bash
python automate.py
```

Output report saved to `/output/report_clean.xlsx`.

---

## Supported Currencies

USD · EUR · GBP · JPY · AED · SGD · INR — easily extendable to any currency via the config file.

---

## Impact

| Before | After |
|---|---|
| 8 hours manual work per month | 11 seconds per run |
| Human errors in every report | Zero errors |
| FX rates looked up manually | Live rates fetched automatically |
| 3 separate files to reconcile | 1 clean unified output |
| No overspend visibility | Automatic budget flags |

**96 hours returned to the finance team every year.**

---

## Services

This script is part of my freelance financial data practice. I build custom financial automation tools, reporting pipelines, and data systems for fintech startups and finance teams globally.

**Available for freelance engagements.**
Connect on [LinkedIn](https://linkedin.com/in/aryan-kasture) or reach out via GitHub.

---

## Author

**Aryan Kasture**
Financial Data Automation & Python Specialist
Mumbai, India

> Helping fintech startups turn messy financial data into clear decisions.
