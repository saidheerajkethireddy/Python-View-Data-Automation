# Python-View-Data-Automation

Feeds respondent-level rNPS, pNPS, and cNPS view data CSVs into weighted NPS calculation scripts, then writes computed outputs into mapped cells of a shared Excel workbook — structured so each chart's values land in predefined cell references, ready for ThinkCell macro automation.

---

## Overview

This repo contains three standalone Python scripts for computing NPS scores from raw Desjardins CAN Life & Wealth NPS Prism survey exports. Each script reads respondent-level view data, applies weighted NPS calculations, and writes results directly into a pre-structured shared Excel workbook at specific cell references — enabling ThinkCell to pick them up automatically for chart rendering.

---

## Scripts

| Script | NPS Type | Input CSV | What it does |
|---|---|---|---|
| `rnps_code.py` | Relationship NPS | `Q1'26_RNPS_View_Data.csv` | Computes weighted rNPS by provider, region (All Canada / Quebec / Rest of Canada), and product (Auto, Home, Auto+Home). Writes 6 blocks (YoY + QoQ × 3 regions) to mapped cells. |
| `cnps_code.py` | Channel NPS | `Q1'26_Channel_nps_view_data.csv` | Computes rolling-window cNPS by channel (Digital, Call Center, Broker/Agent) and region. Writes YoY and QoQ delta blocks to mapped cells. |
| `pnps_code.py` | Product NPS | `q1'26_pnps_data.csv` | Computes rolling-window pNPS by product (Car, Home) and region. Writes results to mapped cells in the shared output workbook. |

---

## How it works

1. Each script reads a respondent-level CSV export from NPS Prism view data
2. Applies weighted NPS formula: `NPS = (Σ weight × category_value) / Σ weight × 100`
3. Segments results by region and product/channel
4. Writes computed values into predefined cell references in a shared Excel output file
5. The Excel cell map mirrors the ThinkCell chart layout — so running a ThinkCell macro after the scripts populates all charts automatically

---

## Setup

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## Configuration

At the top of each script, set these two variables before running:

```python
CSV_PATH    = "your_input_file.csv"       # path to your view data CSV
OUTPUT_PATH = r"CNI_Desjardins_Output.xlsx"  # path to your shared Excel file
```

For `cnps_code.py`, also set the rolling window:

```python
ROLLING = 4   # change to 1 / 2 / 3 / 4 / 5 quarters
```

---

## Running the scripts

```bash
python rnps_code.py
python cnps_code.py
python pnps_code.py
```

Each script will print a confirmation with the sheet, cell reference, and row written upon completion.

---

## Output

All three scripts write to a **single shared Excel workbook** (`CNI_Desjardins_Output.xlsx`). Each script targets a specific set of cell references corresponding to chart positions in the ThinkCell-linked PowerPoint deck.

> **Note:** Input CSV files and the output Excel workbook are not included in this repo — place them locally and update the paths in each script before running.

---

## Repository Structure

```
Python-View-Data-Automation/
├── README.md               # This file
├── rnps_code.py            # Relationship NPS calculator
├── cnps_code.py            # Channel NPS calculator
└── pnps_code.py            # Product NPS calculator
```

---

*Maintained by saidheerajkethireddy*
