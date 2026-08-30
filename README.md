# Fiscal Depreciation Automation

> An end-to-end desktop tool that automates the calculation of **inflation-adjusted fiscal (tax) depreciation** for fixed assets under Mexican tax law — from raw data cleansing and validation, through the full depreciation engine, to audit-ready Excel and Power BI outputs.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-data%20wrangling-150458?logo=pandas&logoColor=white)
![MS Access](https://img.shields.io/badge/MS%20Access-ODBC%20backend-A4373A?logo=microsoftaccess&logoColor=white)
![Power BI](https://img.shields.io/badge/Power%20BI-reporting-F2C811?logo=powerbi&logoColor=black)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)

---

## Table of Contents

- [Why this project exists](#why-this-project-exists)
- [What it does](#what-it-does)
- [Key features](#key-features)
- [How it works: the pipeline](#how-it-works-the-pipeline)
- [The calculation engine (C1 → C2 → C3)](#the-calculation-engine-c1--c2--c3)
- [Asset disposals module](#asset-disposals-module-bajaspy)
- [Data controls & auditability](#data-controls--auditability)
- [Tech stack](#tech-stack)
- [Project structure](#project-structure)
- [Installation](#installation)
- [Usage](#usage)
- [Data model](#data-model)
- [Notes, assumptions & limitations](#notes-assumptions--limitations)
- [Possible extensions](#possible-extensions)

---

## Why this project exists

In Mexico, companies can deduct the depreciation of their fixed assets from taxable income — but the tax figure is **not** the same as the accounting figure. Two things make it tricky:

1. **Inflation adjustment.** The tax deduction is restated for inflation using the *National Consumer Price Index* (INPC). Each asset carries its own "update factor," derived from the ratio between the INPC at the moment of purchase and the INPC at the calculation date.
2. **Special rules & caps.** Certain asset classes (notably vehicles) have deduction ceilings, hybrid/electric units get a higher ceiling, and assets acquired before 2008 follow a different accumulation rule than those acquired afterward.

Doing this by hand across thousands of assets and multiple legal entities means a small army of spreadsheets, VLOOKUPs, and copy-paste — slow, and easy to get subtly wrong.

Think of this tool as a **calculator with an audit trail built in**: you feed it the raw asset register, and it walks every asset through the same disciplined set of rules, showing its work at each step so any number can be traced back to its inputs. The result is a repeatable process that reconciled to the legacy manual method with **near-zero variance**.

---

## What it does

The project is a single **Tkinter desktop application** that orchestrates the entire monthly/annual depreciation cycle from one control panel. On launch it asks you to point it at a **Microsoft Access** database (the working data store), then exposes each stage of the process as a button:

| Stage | Button | What it does |
|-------|--------|--------------|
| 1 | **Data Cleansing** | Normalizes and standardizes the incoming `.xlsx` asset files |
| 2 | **File Validation** | Checks key fields for nulls, wrong types, and bad date formats |
| 3 | **Set Calculation Month** | Defines the calculation date and prior-year cutoff |
| 4 | **INPC Management** | CRUD interface for the inflation-index table |
| 5 | **Identify Hybrid/EV Vehicles** | Flags hybrid & electric units (they get a higher deduction cap) |
| 6 | **Identify Keywords** | Catalog-based tagging of specific asset/account combinations |
| 7 | **Calculate Depreciation** | Runs the core engine and builds the calculation tables |
| 8 | **Export by Entity** | One Excel workbook per legal entity, one sheet per account |
| 9 | **Export for Power BI** | A single flat table ready for BI consumption |

---

## Key features

- **One-click orchestration.** A GUI wraps every step, so a finance/tax analyst can run the whole cycle without touching code or SQL.
- **Inflation-aware calculation.** Implements the INPC update-factor logic, including a 4-decimal truncation convention and guardrails against implausible factors.
- **Regulatory edge cases handled explicitly**, including:
  - Deduction **ceilings** for vehicles, with a higher cap for **hybrid/electric** units.
  - **Pre-2008 vs. post-2008** depreciation-month accumulation.
  - **Reclassification** of executive vehicles.
  - **Exception assets** that carry a special remaining-balance treatment.
- **Smart asset identification.** A tuned regular expression detects hybrid/electric vehicles from free-text asset descriptions (a field notoriously full of typos and abbreviations), plus a catalog-matching fallback for specific accounts.
- **Two export modes.** Human-facing workbooks split by entity/account for review, and a single tidy table optimized for a Power BI data model.
- **Standalone disposals tool.** A companion module computes gain/loss on asset retirements ("bajas") using the same inflation logic.

---

## How it works: the pipeline

```
  ┌─────────────────────────────────────────────────────────────────────┐
  │                      Tkinter control panel                           │
  └─────────────────────────────────────────────────────────────────────┘
             │
   ┌─────────┴──────────┐
   ▼                    ▼
[Raw .xlsx files]   [Access database]
   │                    │
   │  1. Cleanse        │
   │  ───────────────►  │  standardized files
   │                    │
   │                    │  2. Validate ──► "fields to fix" report (.xlsx)
   │                    │
   │                    │  3. Set calculation date
   │                    │  4. Maintain INPC table
   │                    │  5. Flag hybrid/EV assets
   │                    │  6. Tag keyword assets
   │                    │
   │                    ▼
   │            7. CALCULATE ──►  Sociedades ─► C1 ─► C2 ─► C3
   │                    │                        (staged tables)
   │                    ▼
   │            8/9. EXPORT ──►  • Excel per entity (review)
   │                             • Flat table (Power BI)
   ▼
[Disposals tool] ──► merge disposal income + recompute factor ──► gain/loss ──► Excel
```

Each stage writes its results back into the Access database as a named table, so the process is **restartable** and every intermediate result is inspectable.

---

## The calculation engine (C1 → C2 → C3)

The heart of the tool is a chain of three staged tables. Splitting the math into layers is deliberate: each layer adds one conceptual step, so an auditor can open any table and see exactly where a value came from — like showing every line of a long-division problem instead of just the answer.

**Base — `Sociedades`**
A consolidated snapshot of the asset register across all legal entities, enriched via joins to the rate table, hybrid/EV flags, executive-vehicle reclassification, keyword tags, and the exceptions list.

**Layer 1 — `C1_ValorLibros` (book values)**
For every asset it derives:
- The **deductible acquisition value**, after applying the relevant ceiling (a higher one for hybrid/electric vehicles).
- **Months of use**, split into a *pre-2008* segment and a *post-2008* segment, each capped at the asset's maximum depreciation months.
- **Accumulated depreciation** through the prior-year cutoff, **current-year depreciation**, and the **book value** at both the opening and closing of the period.

**Layer 2 — `C2_INPC_Compra` (purchase-side inflation)**
Attaches the **INPC at the time of purchase**, computes the non-deductible portion (acquisition value above the cap), the number of months in the exercise, and the reference month used for the update factor (a mid-period convention).

**Layer 3 — `C3_Depreciacion` (final adjusted figures)**
Attaches the **INPC at the calculation date**, then computes the **update factor** (`INPC_calc / INPC_purchase`, truncated to four decimals and guarded against outliers) and applies it to produce the final **inflation-adjusted fiscal depreciation** and the **updated remaining balance**. A parallel "deferred INPC" variant supports scenarios where the latest official index isn't published yet and an estimate must be used.

---

## Asset disposals module (`Bajas.py`)

When an asset is retired or sold, the company recognizes a **gain or loss** for tax purposes — and that comparison must also be done on an inflation-adjusted basis. This standalone Tkinter tool:

1. Loads **two Excel files** (e.g., the disposals list and the depreciation base) and lets the user pick the sheet and the columns that form a **composite matching key**.
2. **Merges** them on that key (join type is configurable: inner / left / right / outer).
3. Prompts for the **latest known INPC** and recomputes the update factor, the adjusted depreciation, and the **updated remaining balance**.
4. Computes **Gain** = `max(0, disposal income − updated balance)` and **Loss** = `max(0, updated balance − disposal income)`.
5. Exports a clean, ordered Excel workbook.

> Nice touch for data integrity: every column is read as text (`dtype=object`) so that **leading zeros are never silently dropped** from entity and asset identifiers — a classic Excel foot-gun.

---

## Data controls & auditability

This project was built with an **auditor's mindset**, and several design choices reflect that:

- **Validation gate before calculation.** The validation step scans key fields for nulls, type mismatches, and malformed dates, and emits a *"fields to fix"* workbook pinpointing the offending row and column. Bad data is caught **before** it can contaminate the numbers.
- **Full traceability.** The layered `C1 → C2 → C3` design means no figure is a black box — each is reproducible from the layer beneath it.
- **Deterministic, restartable runs.** Calculation tables are dropped and rebuilt from scratch on each run, so results don't depend on hidden prior state.
- **Reconciliation-driven.** The output was validated against the legacy manual process and reached **near-zero variance**, the practical proof that the automation faithfully reproduces the intended tax treatment.
- **Explicit exception handling.** Special cases (executive vehicles, ceilings, redemption-balance exceptions) are modeled as named tables and joins rather than buried in ad-hoc overrides.

---

## Tech stack

| Layer | Technology |
|-------|-----------|
| Language | **Python 3** |
| GUI | **Tkinter** (`tkinter`, `ttk`) |
| Data processing | **pandas** |
| Database | **Microsoft Access** via **pyodbc** (ODBC) |
| Excel I/O | **openpyxl** |
| Text matching | **regular expressions** (`re`) |
| Downstream reporting | **Power BI** (consumes the exported flat table) |

---

## Project structure

```
.
├── Depreciacion_Fiscal_V4.py   # Main application: GUI + full depreciation pipeline
├── Bajas.py                    # Standalone tool for asset disposals (gain/loss)
├── requirements.txt            # Python dependencies
├── .gitignore
└── README.md
```

---

## Installation

> **Platform note:** the tool relies on the **Microsoft Access ODBC driver**, so it is designed to run on **Windows** with Microsoft Access (or the *Microsoft Access Database Engine* redistributable) installed.

```bash
# 1. Clone the repository
git clone <your-repo-url>
cd <your-repo-folder>

# 2. (Recommended) create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

`requirements.txt`:

```
pyodbc
pandas
openpyxl
```

`tkinter` ships with the standard CPython installer on Windows, so no separate install is needed.

---

## Usage

**Main application**

```bash
python Depreciacion_Fiscal_V4.py
```

1. When prompted, select the **Access database** (`.accdb` / `.mdb`) that holds the working tables.
2. Use the control-panel buttons **in order** for a full cycle:
   *Data Cleansing → File Validation → Set Calculation Month → INPC Management → Identify Hybrid/EV → Identify Keywords → Calculate Depreciation → Export.*
3. Collect the generated Excel workbooks (per entity, and/or the Power BI base) from the database's folder.

**Disposals tool**

```bash
python Bajas.py
```

Follow the guided windows to pick the two files, define the matching key and export columns, enter the latest INPC, and save the resulting gain/loss workbook.

---

## Data model

The tool reads and writes a set of named tables in the Access database. Identifiers below are shown **generically**:

| Table | Role |
|-------|------|
| `Unión_Sociedades` | Consolidated raw asset register (source) |
| `Sociedades` | Working snapshot of the register |
| `Tasas_Depreciacion` | Maximum annual depreciation rates & month caps per account/entity |
| `INPC` / `INPC1` | Inflation index values by month |
| `Fechas_Calculo` | Calculation date and prior-year cutoff |
| `Vehiculos_Electricos` | Assets flagged as hybrid/electric |
| `Activos_Keywords` | Assets tagged via the keyword catalog |
| `Reclasificacion_VehiculosEjecutivos` | Executive-vehicle reclassification |
| `Excepcion_SaldoRedimir` | Assets with special remaining-balance treatment |
| `C1_ValorLibros` / `C2_INPC_Compra` / `C3_Depreciacion` | Staged calculation layers (outputs) |

---

## Notes, assumptions & limitations

- The tool is **Windows/Access-oriented** by design; porting the backend to a server-based database (e.g., SQL Server or PostgreSQL) would remove the ODBC/driver dependency.
- Business rules (ceilings, rates, exception lists) live in configuration tables, so they can be adjusted without changing the code — but they **do** encode a specific tax interpretation and should be reviewed against current regulation each period.
- The GUI is optimized for a single analyst running the process interactively; batch/headless execution would be a natural next step.

---
