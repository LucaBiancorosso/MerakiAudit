"""
audit_switch_mtu.py
Audit switch MTU configuration against standards in the SwitchMTU sheet.

API: getNetworkSwitchMtu(networkId)
Response:
  defaultMtuSize: int   (e.g. 9578 for jumbo, 1500 for standard)
  overrides:      list of {switches, switchProfiles, mtuSize}

Standard fields:
  defaultMtuSize  — expected default MTU (e.g. 9578 or 1500)

Usage:
    python -m tools.audit_switch_mtu \\
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
    load_ignored_fields,
    load_mtu_standards,
    resolve_mtu_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

_AUDITED_FIELDS = ["defaultMtuSize"]

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


def audit_network_mtu(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    mtu_config: dict,
    standard: dict | None,
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    detail_rows: list[dict] = []

    def _row(field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=network_id, entity_label=network_name,
            field=field, expected=expected, actual=actual,
            result=result, standard_level=std_level,
        )

    actual_mtu = str(mtu_config.get("defaultMtuSize", ""))

    if standard is None:
        detail_rows.append(_row("defaultMtuSize", "", actual_mtu, "NON_STANDARD"))
        return detail_rows

    for field in _AUDITED_FIELDS:
        if field in ignored_fields:
            detail_rows.append(_row(field, "", actual_mtu, "IGNORED"))
            continue
        expected = standard.get(field, NOT_DEFINED)
        check    = compare_field(field, expected, actual_mtu, ignored_fields)
        detail_rows.append(_row(field, check["expected"], check["actual"], check["result"]))

    return detail_rows


def _print_results(network_name: str, network_id: str, detail_rows: list[dict]) -> None:
    for r in detail_rows:
        result = r["result"]
        icon   = {"PASS": "✓", "FAIL": "✗"}.get(result, "⚠" if result == "NON_STANDARD" else "–")
        if result == "FAIL":
            print(f"  {icon} FAIL     [{network_name}]  {r['field']:<20} "
                  f"expected={r['expected']!r}  actual={r['actual']!r}")
        elif result == "NON_STANDARD":
            print(f"  {icon} NON_STANDARD  [{network_name}]  actual={r['actual']!r}")
        elif result == "PASS":
            print(f"  {icon} PASS     [{network_name}]  {r['field']:<20} value={r['actual']!r}")
        else:
            print(f"  – {result:<14} [{network_name}]  {r['field']}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit switch MTU configuration")
    parser.add_argument("--org-id",         required=True)
    parser.add_argument("--standards-file", required=True)
    parser.add_argument("--network-id",     default=None)
    parser.add_argument("--csv",  action="store_true")
    parser.add_argument("--xlsx", action="store_true")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: {standards_path} not found"); sys.exit(1)

    print(f"Loading MTU standards from {standards_path.name}...")
    standards      = load_mtu_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key  = args.org_id.strip().lower()
    org_stds = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_stds:
        print(f"WARNING: no MTU standards for org '{args.org_id}'. Add rows to 'SwitchMTU'.")
    else:
        print(f"  {len(org_stds)} MTU standard(s) found.")
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
            mtu_config = dashboard.switch.getNetworkSwitchMtu(network_id)
        except Exception as e:
            print(f"  [{network_name}] skipped — {e}"); continue

        standard  = resolve_mtu_standard(standards, args.org_id, network_id)
        oid_l, nid_l = args.org_id.strip().lower(), network_id.strip().lower()
        std_level = "network" if standards.get((oid_l, nid_l)) else (
                    "org"     if standards.get((oid_l, ""))    else "")

        rows = audit_network_mtu(
            args.org_id, org_name, network_id, network_name,
            mtu_config, standard, std_level, ignored_fields,
        )
        for r in rows:
            _print_results(network_name, network_id, [r])
        all_detail.extend(rows)

    summary_rows = build_summary(all_detail)
    print()
    print(f"TOTAL — PASS={sum(r['tot_pass'] for r in summary_rows)}  "
          f"FAIL={sum(r['tot_fail'] for r in summary_rows)}  "
          f"NOT_DEFINED={sum(r['tot_not_defined'] for r in summary_rows)}")

    suffix = f"_{args.network_id}" if args.network_id else ""
    if args.csv:
        write_csv(OUTPUT_DIR / f"audit_mtu_{args.org_id}{suffix}_detail.csv",  all_detail,   _DETAIL_HEADERS)
        write_csv(OUTPUT_DIR / f"audit_mtu_{args.org_id}{suffix}_summary.csv", summary_rows, _SUMMARY_HEADERS)
    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_mtu_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp, _DETAIL_HEADERS, _SUMMARY_HEADERS, "MTUAudit")
        print(f"XLSX → {xp}")


if __name__ == "__main__":
    main()
