# CIS Hardening Compliance Analyzer

Compares Nessus compliance scan results against CIS hardening guide Excel templates and generates two Excel reports:

- **CIS Compliance Report.xlsx** — one sheet per host with PASSED / FAILED / N/A per control
- **CIS Review.xlsx** — controls missing from the scan (manual verification needed) + FYI sheet for newer CIS version checks

## Supported Platforms

- Windows Server (2016 / 2019 / 2022)
- Red Hat Enterprise Linux (RHEL)

## Requirements

```bash
pip install openpyxl colorama
```

## Usage

```bash
python build_compliance_report.py
```

The tool will prompt you for:
1. Folder containing CIS hardening guide `.xlsx` templates
2. Folder containing Nessus scan `.csv` files
3. Output folder for the generated reports

## Guide Template Format

Each Excel guide must have columns:

| S/N | Title | Adoption | ... | Justification |

`Adoption` values: `Yes` / `No` / `NA`

## How Matching Works

Controls are matched between the guide and Nessus scan using a three-level lookup:

1. **Exact title** — after Unicode normalisation and editorial parenthetical stripping
2. **Same S/N verified** — if the same S/N exists in the scan with ≥ 40% title similarity (handles phrasing differences like *"is set to 'Success and Failure'"* vs *"is set to include 'Failure'"*)
3. **Fuzzy cross-S/N** — word-Jaccard + perfect-coverage metric (handles CIS S/N renumbering between guide and audit versions, and Nessus appending `(MS only)` or IIS-specific conditions)

## Output Legend

| Result | Meaning |
|---|---|
| `PASSED` | Nessus check passed |
| `FAILED` | Nessus check failed |
| `NOT IN SCAN` | Adoption=YES but no matching Nessus check found |
| `Not Adopted (Intentional)` | Excluded per business justification |
| `N/A` | Not applicable for this server role |

## Authors

Jasveer · Chen Yung
