"""
audit_switch_stp.py
Audit network-level STP configuration and per-switch bridge priority
against standards defined in the SwitchSTP sheet.

API: getNetworkSwitchStp(networkId)
Response:
  rstpEnabled:       bool
  stpBridgePriority: list of {switches, stacks, switchProfiles, stpPriority}

Standard fields:
  rstpEnabled           — True / False
  default_stp_priority  — expected priority for switches not explicitly listed
                          (must be multiple of 4096, range 0–61440)

Usage:
    python -m tools.audit_switch_stp \\
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
    load_stp_standards,
    resolve_stp_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

_AUDITED_FIELDS = [
    "rstpEnabled",
    "default_stp_priority",
    "root_stp_priority",    # only applied to serials listed in root_serials
    "root_serials",         # defines which switches are expected to be root
]

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key",    # network_id (one row-group per network)
    "entity_label",  # network_name
    "standard_level", "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]


def _effective_priority(serial: str, stp_bridge_priority: list[dict]) -> str | None:
    """Return the stpPriority assigned to a switch serial, or None if not listed."""
    for group in stp_bridge_priority:
        if serial in (group.get("switches") or []):
            return str(group.get("stpPriority", ""))
    return None


def audit_network_stp(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    stp_config: dict,
    switches: list[dict],
    standard: dict | None,
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    detail_rows: list[dict] = []

    def _row(field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=network_id,
            entity_label=network_name,
            field=field, expected=expected, actual=actual,
            result=result, standard_level=std_level,
        )

    if standard is None:
        for field in _AUDITED_FIELDS:
            actuals = {
                "rstpEnabled":          str(stp_config.get("rstpEnabled", "")),
                "default_stp_priority": "",
            }
            detail_rows.append(_row(field, "", actuals.get(field, ""), "NON_STANDARD"))
        return detail_rows

    # ── rstpEnabled ──────────────────────────────────────────────────────────
    field = "rstpEnabled"
    if field not in ignored_fields:
        expected = standard.get(field, NOT_DEFINED)
        actual   = str(stp_config.get("rstpEnabled", ""))
        check    = compare_field(field, expected, actual, ignored_fields)
        detail_rows.append(_row(field, check["expected"], check["actual"], check["result"]))

    # ── per-switch priority — with root vs non-root differentiation ──────────
    stp_bp          = stp_config.get("stpBridgePriority") or []
    default_prio    = standard.get("default_stp_priority", NOT_DEFINED)
    root_prio       = standard.get("root_stp_priority",    NOT_DEFINED)
    root_serials_v  = standard.get("root_serials",         NOT_DEFINED)

    root_serials: set[str] = set()
    if root_serials_v and root_serials_v != NOT_DEFINED:
        root_serials = {s.strip() for s in root_serials_v.split(",") if s.strip()}

    for sw in switches:
        serial  = sw.get("serial", "")
        sw_name = sw.get("name", serial)

        if "default_stp_priority" in ignored_fields and "root_stp_priority" in ignored_fields:
            continue

        is_root      = serial in root_serials
        expected_prio = (root_prio    if is_root else default_prio)
        field_name    = ("root_stp_priority" if is_root else "default_stp_priority")

        if field_name in ignored_fields:
            continue

        actual_prio = _effective_priority(serial, stp_bp)
        # A switch with no explicit entry uses the Meraki default (32768);
        # treat absence as the default rather than the expected value.
        if actual_prio is None:
            actual_prio = "32768"

        if expected_prio == NOT_DEFINED:
            result   = NOT_DEFINED
            expected_str = ""
        else:
            expected_str = str(expected_prio)
            result       = "PASS" if actual_prio == expected_str else "FAIL"

        detail_rows.append(make_row(
            org_id, org_name, network_id, network_name,
            entity_key=f"{network_id}:{serial}",
            entity_label=f"{network_name} / {sw_name}",
            field=field_name,
            expected=expected_str,
            actual=actual_prio,
            result=result,
            standard_level=std_level,
        ))

    return detail_rows


def _print_results(network_name: str, network_id: str, detail_rows: list[dict]) -> None:
    pass_c = sum(1 for r in detail_rows if r["result"] == "PASS")
    fail_c = sum(1 for r in detail_rows if r["result"] == "FAIL")
    nd_c   = sum(1 for r in detail_rows if r["result"] in (NOT_DEFINED, "NOT_DEFINED"))
    print(f"  [{network_name}]  {network_id}")
    print(f"    PASS={pass_c}  FAIL={fail_c}  NOT_DEFINED={nd_c}")
    for r in detail_rows:
        result = r["result"]
        if result == "FAIL":
            print(f"    ✗ FAIL    {r['field']:<28} entity={r['entity_label']:<30} "
                  f"expected={r['expected']!r}  actual={r['actual']!r}")
        elif result == "NON_STANDARD":
            print(f"    ⚠ NON_STANDARD  {r['field']:<28} actual={r['actual']!r}")
        elif result == "PASS":
            print(f"    ✓ PASS    {r['field']:<28} value={r['actual']!r}")
        else:
            print(f"    – {result:<14} {r['field']:<28}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit switch STP configuration")
    parser.add_argument("--org-id",         required=True)
    parser.add_argument("--standards-file", required=True)
    parser.add_argument("--network-id",     default=None)
    parser.add_argument("--csv",  action="store_true")
    parser.add_argument("--xlsx", action="store_true")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: {standards_path} not found"); sys.exit(1)

    print(f"Loading STP standards from {standards_path.name}...")
    standards      = load_stp_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key = args.org_id.strip().lower()
    org_stds = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_stds:
        print(f"WARNING: no STP standards for org '{args.org_id}'. "
              "Add rows to the 'SwitchSTP' sheet.")
    else:
        print(f"  {len(org_stds)} STP standard(s) found.")
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

    # Fetch switch inventory once
    try:
        all_switches = dashboard.organizations.getOrganizationDevices(
            args.org_id, total_pages="all", productTypes=["switch"]
        )
    except Exception as e:
        print(f"ERROR fetching switches: {e}"); sys.exit(1)

    switches_by_network: dict[str, list[dict]] = {}
    for sw in all_switches:
        nid = str(sw.get("networkId", "")).strip()
        if nid:
            switches_by_network.setdefault(nid, []).append(sw)

    print(f"Auditing {len(networks_to_audit)} network(s)...")
    print()

    all_detail: list[dict] = []

    for network_id, network_name in sorted(networks_to_audit.items(), key=lambda x: x[1]):
        switches = switches_by_network.get(network_id, [])
        if not switches:
            continue
        try:
            stp_config = dashboard.switch.getNetworkSwitchStp(network_id)
        except Exception as e:
            print(f"  [{network_name}] skipped — {e}"); continue

        standard  = resolve_stp_standard(standards, args.org_id, network_id)
        oid_l, nid_l = args.org_id.strip().lower(), network_id.strip().lower()
        std_level = "network" if standards.get((oid_l, nid_l)) else (
                    "org"     if standards.get((oid_l, ""))    else "")

        rows = audit_network_stp(
            args.org_id, org_name, network_id, network_name,
            stp_config, switches, standard, std_level, ignored_fields,
        )
        _print_results(network_name, network_id, rows)
        all_detail.extend(rows)

    summary_rows = build_summary(all_detail)
    print()
    print(f"TOTAL — PASS={sum(r['tot_pass'] for r in summary_rows)}  "
          f"FAIL={sum(r['tot_fail'] for r in summary_rows)}  "
          f"NOT_DEFINED={sum(r['tot_not_defined'] for r in summary_rows)}")

    suffix = f"_{args.network_id}" if args.network_id else ""
    if args.csv:
        write_csv(OUTPUT_DIR / f"audit_stp_{args.org_id}{suffix}_detail.csv",  all_detail,   _DETAIL_HEADERS)
        write_csv(OUTPUT_DIR / f"audit_stp_{args.org_id}{suffix}_summary.csv", summary_rows, _SUMMARY_HEADERS)
    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_stp_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp, _DETAIL_HEADERS, _SUMMARY_HEADERS, "STPAudit")
        print(f"XLSX → {xp}")


if __name__ == "__main__":
    main()
