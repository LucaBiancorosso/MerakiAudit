"""
audit_switch_dhcp_snooping.py
Audit DHCP snooping and Dynamic ARP Inspection (DAI) network-level settings
against standards defined in the SwitchDHCPSnooping sheet.

API: getNetworkSwitchDhcpServerPolicy(networkId)
Response:
  defaultPolicy             — "allow" or "block"
  allowedServers            — list of MACs (only used when defaultPolicy=block)
  blockedServers            — list of MACs (only used when defaultPolicy=allow)
  alwaysAllowedServers      — list of MACs always allowed regardless of policy
  alerts.email.enabled      — bool
  arpInspection.enabled     — bool (DAI)
  arpInspection.unsupportedModels — informational only

Standard fields:
  defaultPolicy                   — allow / block
  arpInspection_enabled           — True / False
  alerts_email_enabled            — True / False
  required_always_allowed_servers — comma-sep MACs that must be in alwaysAllowedServers

Usage:
    python -m tools.audit_switch_dhcp_snooping \\
        --org-id <ORG_ID> \\
        --standards-file standards/standard_audit_fields.xlsx --xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv
from lib.standards import (
    NOT_DEFINED,
    compare_field,
    load_dhcp_snooping_standards,
    load_ignored_fields,
    resolve_dhcp_snooping_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

_AUDITED_FIELDS = [
    "defaultPolicy",
    "arpInspection_enabled",
    "alerts_email_enabled",
    "required_always_allowed_servers",  # subset check — all required MACs must be present
]

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "standard_level", "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]


def _flatten_dhcp(config: dict) -> dict[str, str]:
    alerts      = config.get("alerts") or {}
    email       = alerts.get("email") or {}
    arp         = config.get("arpInspection") or {}
    always_srv  = config.get("alwaysAllowedServers") or []

    return {
        "defaultPolicy":                   str(config.get("defaultPolicy", "")),
        "arpInspection_enabled":            str(arp.get("enabled", "")),
        "alerts_email_enabled":             str(email.get("enabled", "")),
        "required_always_allowed_servers":  ",".join(
            sorted(m.lower() for m in always_srv if m)
        ),
    }


def audit_network_dhcp(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    dhcp_config: dict,
    standard: dict | None,
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    flat = _flatten_dhcp(dhcp_config)
    detail_rows: list[dict] = []

    def _row(field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=network_id, entity_label=network_name,
            field=field, expected=expected, actual=actual,
            result=result, standard_level=std_level,
        )

    if standard is None:
        for field in _AUDITED_FIELDS:
            detail_rows.append(_row(field, "", flat.get(field, ""), "NON_STANDARD"))
        return detail_rows

    for field in _AUDITED_FIELDS:
        if field in ignored_fields:
            detail_rows.append(_row(field, "", flat.get(field, ""), "IGNORED"))
            continue

        expected = standard.get(field, NOT_DEFINED)
        actual   = flat.get(field, "")

        if expected == NOT_DEFINED:
            detail_rows.append(_row(field, "", actual, NOT_DEFINED))
            continue

        if field == "required_always_allowed_servers":
            # Subset check: all required MACs must be present (normalised lowercase)
            req_macs  = {m.strip().lower() for m in expected.split(",") if m.strip()}
            live_macs = {m.strip().lower() for m in actual.split(",")   if m.strip()}
            missing   = req_macs - live_macs
            result    = "PASS" if not missing else "FAIL"
            detail_rows.append(_row(field, expected, actual, result))
        else:
            check = compare_field(field, expected, actual, ignored_fields)
            detail_rows.append(_row(field, check["expected"], check["actual"], check["result"]))

    return detail_rows


def _print_results(network_name: str, detail_rows: list[dict]) -> None:
    pass_c = sum(1 for r in detail_rows if r["result"] == "PASS")
    fail_c = sum(1 for r in detail_rows if r["result"] == "FAIL")
    nd_c   = sum(1 for r in detail_rows if r["result"] in (NOT_DEFINED, "NOT_DEFINED"))
    print(f"  [{network_name}]  PASS={pass_c}  FAIL={fail_c}  NOT_DEFINED={nd_c}")
    for r in detail_rows:
        result = r["result"]
        if result == "FAIL":
            print(f"    ✗ FAIL    {r['field']:<36} "
                  f"expected={r['expected']!r}  actual={r['actual']!r}")
        elif result == "NON_STANDARD":
            print(f"    ⚠ NON_STANDARD  {r['field']:<36} actual={r['actual']!r}")
        elif result == "PASS":
            print(f"    ✓ PASS    {r['field']:<36} value={r['actual']!r}")
        else:
            print(f"    – {result:<14} {r['field']}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit switch DHCP snooping and DAI configuration"
    )
    parser.add_argument("--org-id",         required=True)
    parser.add_argument("--standards-file", required=True)
    parser.add_argument("--network-id",     default=None)
    parser.add_argument("--csv",  action="store_true")
    parser.add_argument("--xlsx", action="store_true")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: {standards_path} not found"); sys.exit(1)

    print(f"Loading DHCP snooping standards from {standards_path.name}...")
    standards      = load_dhcp_snooping_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key  = args.org_id.strip().lower()
    org_stds = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_stds:
        print(f"WARNING: no DHCP snooping standards for org '{args.org_id}'. "
              "Add rows to 'SwitchDHCPSnooping'.")
    else:
        print(f"  {len(org_stds)} standard(s) found.")
    print()

    dashboard      = get_dashboard()
    network_lookup = build_network_lookup(dashboard, args.org_id)

    org_name = ""
    try:
        for o in dashboard.organizations.getOrganizations():
            if str(o.get("id")) == str(args.org_id):
                org_name = o.get("name", ""); break
    except Exception:
        pass

    if args.network_id:
        if args.network_id not in network_lookup:
            print(f"ERROR: network '{args.network_id}' not found"); sys.exit(1)
        networks_to_audit = {args.network_id: network_lookup[args.network_id]}
    else:
        networks_to_audit = dict(network_lookup)

    print(f"Auditing {len(networks_to_audit)} network(s)...")
    print()

    all_detail: list[dict] = []

    for network_id, network_name in sorted(networks_to_audit.items(), key=lambda x: x[1]):
        try:
            dhcp_config = dashboard.switch.getNetworkSwitchDhcpServerPolicy(network_id)
        except Exception as e:
            print(f"  [{network_name}] skipped — {e}"); continue

        standard  = resolve_dhcp_snooping_standard(standards, args.org_id, network_id)
        oid_l, nid_l = args.org_id.strip().lower(), network_id.strip().lower()
        std_level = "network" if standards.get((oid_l, nid_l)) else (
                    "org"     if standards.get((oid_l, ""))    else "")

        rows = audit_network_dhcp(
            args.org_id, org_name, network_id, network_name,
            dhcp_config, standard, std_level, ignored_fields,
        )
        _print_results(network_name, rows)
        all_detail.extend(rows)

    summary_rows = build_summary(all_detail)
    print()
    print(f"TOTAL — PASS={sum(r['tot_pass'] for r in summary_rows)}  "
          f"FAIL={sum(r['tot_fail'] for r in summary_rows)}  "
          f"NOT_DEFINED={sum(r['tot_not_defined'] for r in summary_rows)}")

    suffix = f"_{args.network_id}" if args.network_id else ""
    if args.csv:
        write_csv(OUTPUT_DIR / f"audit_dhcp_snooping_{args.org_id}{suffix}_detail.csv",  all_detail,   _DETAIL_HEADERS)
        write_csv(OUTPUT_DIR / f"audit_dhcp_snooping_{args.org_id}{suffix}_summary.csv", summary_rows, _SUMMARY_HEADERS)
    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_dhcp_snooping_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp, _DETAIL_HEADERS, _SUMMARY_HEADERS, "DHCPSnoopAudit")
        print(f"XLSX → {xp}")


if __name__ == "__main__":
    main()
