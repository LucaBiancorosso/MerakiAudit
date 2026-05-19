# Meraki Toolkit

A collection of CLI tools to interact with the Cisco Meraki Dashboard API.
Covers inventory listing, SSID/RF/AP configuration auditing, and End-of-Life forecasting — all org-wide or filtered to a single network.

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
│   └── settings.py                 # Env-based configuration (API key, log path)
├── lib/
│   ├── meraki_client.py            # Authenticated DashboardAPI factory
│   ├── output.py                   # CSV / JSON file writers
│   ├── standards.py                # Standards loader, field comparator, subnet checker
│   └── audit_common.py             # Shared audit helpers (rows, summary, XLSX output)
├── tools/
│   ├── list_organizations.py       # List all orgs accessible by the API key
│   ├── list_networks.py            # List networks in an org
│   ├── list_inventory.py           # List devices with EoX lifecycle data
│   ├── list_ssid.py                # List SSIDs in a network
│   ├── list_ssid_settings.py       # Detailed SSID settings (auth, VLAN, RADIUS, etc.)
│   ├── audit_ssid.py               # SSID compliance audit
│   ├── audit_rf_profile.py         # RF profile compliance audit
│   ├── audit_ap.py                 # AP management & connectivity audit
│   ├── audit_switch_interfaces.py  # Switch port role-based audit
│   └── forecast_eol.py             # End-of-Life forecast Excel report
├── standards/
│   └── standard_audit_fields.xlsx   # Standards template (fill in and use for audits)
├── output/                         # Generated files — git-ignored
├── .env.example
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Standards File

All three audit tools share a single Excel file: `standards/standard_audit_fields.xlsx`.
It contains the following sheets:

| Sheet | Used by | Key columns |
|---|---|---|
| `Standards` | `audit_ssid` | `org_id`, `network_id` *(optional)*, `ssid_name` |
| `RFProfiles` | `audit_rf_profile` | `org_id`, `network_id` *(optional)*, `profile_name` |
| `APConfig` | `audit_ap` | `org_id`, `network_id` *(optional)* |
| `SwitchInterfaces` | `audit_switch_interfaces` | `org_id`, `network_id` *(optional)*, `role` |
| `VlanRoles` | `audit_switch_interfaces` | `org_id`, `network_id` *(optional)*, `vlan_id` → `role` |
| `Ignore` | all audit tools | `field_name` — fields to skip globally |

**Network-level overrides:** for all sheets, leaving `network_id` blank defines an org-wide default. Filling it in creates a network-specific standard that takes priority over the org-wide row for that network. This allows most networks to share one standard while a specific network can deviate.

**Tip:** format `org_id` and `network_id` cells as **Text** in Excel to prevent large IDs from being converted to scientific notation.

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

### List Inventory (with EoX lifecycle data)
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

# Include unconfigured placeholder slots (hidden by default)
python -m tools.list_ssid --network-id <NETWORK_ID> --include-unconfigured
```

### List SSID Settings (detailed)
Includes auth mode, encryption, VLAN tagging, client IP assignment, RADIUS hosts, bandwidth limits, splash page, and more.

```bash
python -m tools.list_ssid_settings --network-id <NETWORK_ID>
python -m tools.list_ssid_settings --network-id <NETWORK_ID> --csv --json
python -m tools.list_ssid_settings --network-id <NETWORK_ID> --include-unconfigured
```

---

## Audit Tools

All audit tools share the same behaviour:
- **`--network-id`** limits the audit to a single network (MISSING checks are suppressed since a standard SSID/profile may simply live in a different network).
- **`--xlsx`** produces a colour-coded Excel report with a Summary sheet and a Detail sheet.
- **`--csv`** writes a `_detail.csv` and `_summary.csv` to the `output/` directory.
- **Unconfigured SSIDs** (`Unconfigured SSID N`) are silently skipped in all wireless audits.
- Detail rows carry a `standard_level` column (`org` or `network`) showing which standard was applied.

### Result codes

| Result | Meaning |
|---|---|
| `PASS` | Field matches the standard |
| `FAIL` | Field does not match the standard |
| `NOT_DEFINED` | Standard cell was left blank — field is not audited |
| `IGNORED` | Field is listed in the `Ignore` sheet |
| `NON_STANDARD` | Entity exists live but has no entry in the standard |
| `MISSING` / `MISSING_SSID` | Entity is defined in the standard but not found live |

---

### SSID Compliance Audit

Audits every SSID in every wireless network against the `Standards` sheet.

Checked fields include: `enabled`, `authMode`, `encryptionMode`, `wpaEncryptionMode`,
`ipAssignmentMode`, `defaultVlanId`, `useVlanTagging`, `concentratorNetworkId`,
`lanIsolationEnabled`, `perClientBandwidthLimitUp`, `perClientBandwidthLimitDown`,
`bandSelection`, `minBitrate`, `ssidAdminAccessible`, `radiusEnabled`,
`radiusAccountingEnabled`, `radiusHosts` *(compared as unordered set)*, `splashPage`,
`walledGardenEnabled`, `visible`.

```bash
# Audit all networks in an org
python -m tools.audit_ssid --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx --xlsx

# Audit a single network
python -m tools.audit_ssid --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx \
    --network-id <NETWORK_ID> --xlsx --csv

# Skip disabled SSIDs (only audit enabled ones)
python -m tools.audit_ssid --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx \
    --enabled-only --xlsx
```

| Argument | Default | Description |
|---|---|---|
| `--org-id` | required | Meraki Organization ID |
| `--standards-file` | required | Path to the standards Excel file |
| `--network-id` | all networks | Limit audit to a single network |
| `--enabled-only` | false | Skip disabled SSIDs; audit only currently enabled ones |
| `--csv` | false | Write detail and summary CSV files |
| `--xlsx` | false | Write colour-coded Excel report |

---

### Switch Interface Audit

Audits every enabled switch port in the org. Ports are assigned a role automatically then checked against the standard for that role.

**Role detection order:**
1. CDP/LLDP neighbour platform matches `cdp_ap_pattern` (regex) → `trunk_ap`
2. CDP/LLDP neighbour platform matches `cdp_cpe_pattern` (regex) → `trunk_cpe`
3. No CDP match + trunk port → `trunk_unknown` *(flagged as NON_STANDARD, not audited)*
4. No CDP match + access port → VLAN looked up in `VlanRoles` sheet → e.g. `access_data`, `access_voice`
5. VLAN not in map → `access_unknown` *(flagged as NON_STANDARD, not audited)*

The regex patterns (`cdp_ap_pattern`, `cdp_cpe_pattern`) are defined per role row in `SwitchInterfaces` and can be different per network. For example `^MR|^CW` for APs, `^MX|Versa|VOS` for CPEs.

**Audited fields per port:**

| Field | Notes |
|---|---|
| `enabled` | Admin state |
| `type` | `access` / `trunk` |
| `vlan` | Access VLAN or native VLAN |
| `voiceVlan` | Access ports only |
| `allowedVlans` | Trunk only — compared as **expanded VLAN set** (handles ranges like `20-30`) |
| `poeEnabled` | |
| `isolationEnabled` | Client isolation |
| `rstpEnabled` | Rapid STP |
| `stpGuard` | `disabled` / `root guard` / `bpdu guard` / `loop guard` |
| `stpPortFastTrunk` | Trunk only |
| `udld` | `Disabled` / `Alert only` / `Enforce` |
| `stormControlEnabled` | |
| `daiTrusted` | Dynamic ARP Inspection trust — trunk only |
| `accessPolicyType` | `Open` / `Custom access policy` / `MAC allow list` / `Sticky MAC allow list` |
| `linkNegotiation` | |
| `dot3az_enabled` | Energy Efficient Ethernet |

```bash
python -m tools.audit_switch_interfaces --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx --xlsx

# Single network
python -m tools.audit_switch_interfaces --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx \
    --network-id <NETWORK_ID> --xlsx --csv

# Include admin-disabled ports (skipped by default)
python -m tools.audit_switch_interfaces --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx \
    --include-disabled --xlsx
```

| Argument | Default | Description |
|---|---|---|
| `--org-id` | required | Meraki Organization ID |
| `--standards-file` | required | Path to standards Excel file |
| `--network-id` | all networks | Limit audit to a single network |
| `--include-disabled` | false | Include admin-disabled ports |
| `--csv` | false | Write detail and summary CSV files |
| `--xlsx` | false | Write colour-coded Excel report |

---

### RF Profile Audit

Audits RF profiles in every wireless network against the `RFProfiles` sheet.

Checked fields include (all three bands where applicable):
`clientBalancingEnabled`, `minBitrateType`, `bandSelectionType`, `transmission_enabled`,
`isIndoorDefault`, `isOutdoorDefault`, `ap_bandOperationMode`, `ap_bandSteeringEnabled`,
`ap_bands_enabled`, `2g_maxPower`, `2g_minPower`, `2g_minBitrate`, `2g_validAutoChannels`,
`2g_axEnabled`, `2g_rxsop`, `5g_maxPower`, `5g_minPower`, `5g_minBitrate`,
`5g_channelWidth`, `5g_validAutoChannels`, `5g_rxsop`, and the equivalent `6g_*` fields.

Channel and band lists are compared as **unordered sets** — `"1,6,11"` matches `"11,6,1"`.

```bash
python -m tools.audit_rf_profile --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx --xlsx

python -m tools.audit_rf_profile --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx \
    --network-id <NETWORK_ID> --xlsx --csv
```

---

### AP Configuration Audit

Audits every wireless AP in the org against the `APConfig` sheet.
Combines three API calls per network: device inventory, alternate management interface, and mesh status.

| Field | Check type | Description |
|---|---|---|
| `mgmt_ip_mode` | exact | `static` or `dhcp` |
| `mgmt_ip_in_subnet` | CIDR membership | AP management IP must fall within one of the comma-separated CIDR ranges defined in the standard (e.g. `10.0.0.0/8, 192.168.1.0/24`) |
| `mgmt_vlan` | exact | Management VLAN ID from alternate management interface |
| `mgmt_dns1` | exact | Primary DNS server |
| `mgmt_dns2` | exact | Secondary DNS server |
| `connection_mode` | exact | `gateway` or `mesh` (derived from mesh status API) |
| `tags_subset` | subset | All expected tags must be present on the AP; additional tags are allowed |

```bash
python -m tools.audit_ap --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx --xlsx

python -m tools.audit_ap --org-id <ORG_ID> \
    --standards-file standards/standard_audit_fields.xlsx \
    --network-id <NETWORK_ID> --xlsx --csv
```

---

### EoL Forecast (Excel report)

Generates a multi-sheet Excel workbook forecasting device End-of-Life by model and year.
Requires inventory CSV files to be generated first (`list_inventory --csv` for each org).

```bash
python -m tools.forecast_eol
python -m tools.forecast_eol --date-field endOfSaleAt --start-year 2024 --end-year 2031
python -m tools.forecast_eol --only-eol --only-network-associated
```

| Argument | Default | Description |
|---|---|---|
| `--date-field` | `entOfSupportAt` | Lifecycle date field: `entOfSupportAt` or `endOfSaleAt` |
| `--start-year` | current year | First year column in the forecast |
| `--end-year` | current year + 7 | Last year column in the forecast |
| `--org-file` | `output/organizations.csv` | Path to the organizations CSV (run `list_organizations` first) |
| `--only-eol` | false | Only include devices with a lifecycle date within the forecast range |
| `--only-network-associated` | false | Exclude devices not assigned to any network |
| `--output-file` | `output/eol_forecast.xlsx` | Output file path |

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
| `audit_rf_profile` | `audit_rf_<org_id>.xlsx`, `audit_rf_<org_id>_detail.csv`, `audit_rf_<org_id>_summary.csv` |
| `audit_ap` | `audit_ap_<org_id>.xlsx`, `audit_ap_<org_id>_detail.csv`, `audit_ap_<org_id>_summary.csv` |
| `audit_switch_interfaces` | `audit_switch_<org_id>.xlsx`, `audit_switch_<org_id>_detail.csv`, `audit_switch_<org_id>_summary.csv` |
| `forecast_eol` | `eol_forecast.xlsx` |

---

## License

MIT

---

## Roadmap

See [`ROADMAP.md`](ROADMAP.md) for planned audit coverage across switching, security appliance, wireless, and org-level compliance.
