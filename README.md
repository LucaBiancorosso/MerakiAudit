# Meraki Toolkit

A collection of CLI scripts to interact with the Cisco Meraki Dashboard API.  
Supports listing organizations, networks, inventory, SSIDs, and generating End-of-Life (EoL) forecasts.

---

## Requirements

- Python 3.9+
- A Meraki Dashboard API key ([how to generate one](https://documentation.meraki.com/General_Administration/Other_Topics/Cisco_Meraki_Dashboard_API#Enable_API_Access))

---

## Setup

```bash
# 1. Clone the repository
git clone <your-repo-url>
cd meraki_toolkit

# 2. Create and activate a virtual environment
python -m venv venv
source venv/bin/activate        # Linux / macOS
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure environment variables
cp .env.example .env
# Edit .env and set your MERAKI_DASHBOARD_API_KEY
```

---

## Project Structure

```
meraki_toolkit/
├── config/
│   └── settings.py             # Env-based configuration
├── lib/
│   ├── meraki_client.py        # Authenticated DashboardAPI factory
│   ├── output.py               # CSV / JSON helpers
│   └── standards.py            # Standards loader and field comparator
├── tools/
│   ├── list_organizations.py   # List all orgs
│   ├── list_networks.py        # List networks in an org
│   ├── list_inventory.py       # List devices with EoX data
│   ├── list_ssid.py            # List SSIDs in a network
│   ├── list_ssid_settings.py   # Detailed SSID settings
│   ├── audit_ssid.py           # SSID compliance audit
│   └── forecast_eol.py         # Generate EoL forecast Excel report
├── standards/                  # Compliance standard templates
│   └── ssid_standards_template.xlsx
├── output/                     # Generated files (git-ignored)
├── .env.example
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Usage

All tools are run from the **project root**.

### List Organizations
```bash
python -m tools.list_organizations
python -m tools.list_organizations --csv --json
```

### List Networks
```bash
python -m tools.list_networks --org-id <ORG_ID>
python -m tools.list_networks --org-id <ORG_ID> --csv --json
```

### List Inventory (with EoX)
```bash
python -m tools.list_inventory --org-id <ORG_ID>
python -m tools.list_inventory --org-id <ORG_ID> --csv --json

# Filter by EoX status: endOfSale | endOfSupport | nearEndOfSupport | null
python -m tools.list_inventory --org-id <ORG_ID> --eox-status endOfSupport --csv
```

### List SSIDs
```bash
python -m tools.list_ssid --network-id <NETWORK_ID>
python -m tools.list_ssid --network-id <NETWORK_ID> --csv --json
```

### List SSID Settings (detailed)
```bash
python -m tools.list_ssid_settings --network-id <NETWORK_ID>
python -m tools.list_ssid_settings --network-id <NETWORK_ID> --csv --json
```

### SSID Compliance Audit
Requires a filled-in standards file (see `standards/ssid_standards_template.xlsx`).

```bash
python -m tools.audit_ssid --org-id <ORG_ID> --standards-file standards/ssid_standards_template.xlsx --xlsx
# Audit a single network only
python -m tools.audit_ssid --org-id <ORG_ID> --standards-file standards/ssid_standards_template.xlsx --network-id <NETWORK_ID> --xlsx
```

### EoL Forecast (Excel report)
Requires inventory CSV files to be already generated (run `list_inventory` first for each org).

```bash
python -m tools.forecast_eol
python -m tools.forecast_eol --date-field endOfSaleAt --start-year 2024 --end-year 2031
python -m tools.forecast_eol --only-eol --only-network-associated
```

| Argument | Default | Description |
|---|---|---|
| `--date-field` | `entOfSupportAt` | Lifecycle date to use (`entOfSupportAt` or `endOfSaleAt`) |
| `--start-year` | current year | First year column in forecast |
| `--end-year` | current year + 7 | Last year column in forecast |
| `--org-file` | `output/organizations.csv` | Path to organizations CSV |
| `--only-eol` | false | Include only devices with a date in the forecast range |
| `--only-network-associated` | false | Exclude unassigned devices |
| `--output-file` | `output/eol_forecast.xlsx` | Output path |

---

## Output Files

All generated files are saved in the `output/` directory and are git-ignored by default.

| Tool | Output files |
|---|---|
| `list_organizations` | `organizations.csv`, `organizations.json` |
| `list_networks` | `networks_<org_id>.csv`, `networks_<org_id>.json` |
| `list_inventory` | `inventory_<org_id>.csv`, `inventory_<org_id>.json` |
| `list_ssid` | `ssids_<network_id>.csv`, `ssids_<network_id>.json` |
| `list_ssid_settings` | `ssid_setting_<network_id>.csv`, `ssid_setting_<network_id>.json` |
| `audit_ssid` | `audit_ssid_<org_id>.xlsx`, `audit_ssid_<org_id>_detail.csv`, `audit_ssid_<org_id>_summary.csv` |
| `forecast_eol` | `eol_forecast.xlsx` |

---

## License

MIT
