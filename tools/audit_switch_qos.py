"""
audit_switch_qos.py
Audit switch QoS rules against required rules defined in the SwitchQoS sheet.

API: getNetworkSwitchQosRules(networkId)
Response per rule: {id, vlan, protocol, srcPort, srcPortRange,
                    dstPort, dstPortRange, dscp}

Note: QoS rules have no comment field — they are identified by rule_name
defined in the standard sheet (used only for matching, not sent to the API).
Matching is done by (vlan + protocol + dstPort) combination.

Audit logic:
  - Each SwitchQoS row defines a REQUIRED rule identified by rule_name.
  - Matching: find the live rule where vlan+protocol+dstPort all match.
  - If found: check dscp and port range fields.
  - If not found: flag as MISSING.
  - Live rules not matched to any standard rule are flagged as NON_STANDARD.

Standard sheet columns:
  org_id, network_id (optional), rule_name (key — label only),
  vlan, protocol, srcPort, srcPortRange, dstPort, dstPortRange, dscp

Usage:
    python -m tools.audit_switch_qos \\
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
    load_ignored_fields,
    load_qos_standards,
    resolve_qos_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

# Fields used as the matching key
_MATCH_FIELDS  = {"vlan", "protocol", "dstPort"}
# Fields checked once a rule is matched
_AUDITED_FIELDS = ["vlan", "protocol", "srcPort", "srcPortRange",
                   "dstPort", "dstPortRange", "dscp"]

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key",    # rule_name
    "entity_label",  # rule_name
    "standard_level", "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]


def _s(v) -> str:
    if v is None:
        return ""
    return str(v).strip().lower()


def _match_key(rule: dict) -> tuple[str, str, str]:
    def _mk(v):
        """Treat NOT_DEFINED as empty so standard wildcards match live None."""
        s = _s(v)
        return "" if s in ("not_defined", "not defined") else s
    return (_mk(rule.get("vlan")), _mk(rule.get("protocol")), _mk(rule.get("dstPort")))


def audit_network_qos(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    live_rules: list[dict],
    required_rules: list[dict],
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    detail_rows: list[dict] = []

    # Build live lookup by match key — last one wins if duplicate
    live_by_key: dict[tuple, dict] = {}
    for r in live_rules:
        live_by_key[_match_key(r)] = r

    matched_keys: set[tuple] = set()

    def _row(rule_name, field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=rule_name, entity_label=rule_name,
            field=field, expected=expected, actual=actual,
            result=result, standard_level=std_level,
        )

    for req in required_rules:
        rule_name = req.get("rule_name", "")
        req_key   = _match_key(req)
        live_rule = live_by_key.get(req_key)

        if live_rule is None:
            for field in _AUDITED_FIELDS:
                detail_rows.append(_row(rule_name, field,
                                        str(req.get(field, "") or ""), "", "MISSING"))
            continue

        matched_keys.add(req_key)

        for field in _AUDITED_FIELDS:
            if field in ignored_fields:
                detail_rows.append(_row(rule_name, field, "",
                                        str(live_rule.get(field, "") or ""), "IGNORED"))
                continue
            expected = req.get(field, NOT_DEFINED)
            actual   = str(live_rule.get(field, "") or "")
            if expected == NOT_DEFINED:
                detail_rows.append(_row(rule_name, field, "", actual, NOT_DEFINED))
            else:
                result = "PASS" if _s(expected) == _s(actual) else "FAIL"
                detail_rows.append(_row(rule_name, field, str(expected), actual, result))

    # Flag live rules not matched to any standard rule
    for key, live_rule in live_by_key.items():
        if key not in matched_keys:
            label = f"UNEXPECTED vlan={key[0]} proto={key[1]} dstPort={key[2]}"
            for field in _AUDITED_FIELDS:
                detail_rows.append(_row(label, field,
                                        "", str(live_rule.get(field, "") or ""), "NON_STANDARD"))

    return detail_rows


def _print_results(network_name: str, network_id: str, detail_rows: list[dict]) -> None:
    pass_c  = sum(1 for r in detail_rows if r["result"] == "PASS")
    fail_c  = sum(1 for r in detail_rows if r["result"] == "FAIL")
    miss_c  = len({r["entity_key"] for r in detail_rows if r["result"] == "MISSING"})
    ns_c    = len({r["entity_key"] for r in detail_rows if r["result"] == "NON_STANDARD"})
    print(f"  [{network_name}]  {network_id}")
    print(f"    PASS={pass_c}  FAIL={fail_c}  MISSING_RULES={miss_c}  UNEXPECTED_RULES={ns_c}")

    rules_seen: list[str] = []
    for r in detail_rows:
        if r["entity_key"] not in rules_seen:
            rules_seen.append(r["entity_key"])

    for rule_name in rules_seen:
        rule_rows = [r for r in detail_rows if r["entity_key"] == rule_name]
        has_issue = any(r["result"] not in ("PASS", NOT_DEFINED, "NOT_DEFINED", "IGNORED")
                        for r in rule_rows)
        if not has_issue:
            continue
        print(f"\n    Rule: {rule_name!r}")
        for r in rule_rows:
            result = r["result"]
            if result == "FAIL":
                print(f"      ✗ FAIL        {r['field']:<14} "
                      f"expected={r['expected']!r}  actual={r['actual']!r}")
            elif result == "MISSING":
                print(f"      ✗ MISSING     {r['field']:<14} expected={r['expected']!r}")
            elif result == "NON_STANDARD":
                print(f"      ⚠ UNEXPECTED  {r['field']:<14} actual={r['actual']!r}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit switch QoS rules")
    parser.add_argument("--org-id",         required=True)
    parser.add_argument("--standards-file", required=True)
    parser.add_argument("--network-id",     default=None)
    parser.add_argument("--csv",  action="store_true")
    parser.add_argument("--xlsx", action="store_true")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: {standards_path} not found"); sys.exit(1)

    print(f"Loading QoS standards from {standards_path.name}...")
    standards      = load_qos_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key  = args.org_id.strip().lower()
    org_stds = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_stds:
        print(f"WARNING: no QoS standards for org '{args.org_id}'. Add rows to 'SwitchQoS'.")
    else:
        total_rules = sum(len(v) for v in org_stds.values())
        print(f"  {total_rules} required QoS rule(s) across {len(org_stds)} standard(s).")
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
            live_rules = dashboard.switch.getNetworkSwitchQosRules(network_id)
        except Exception as e:
            print(f"  [{network_name}] skipped — {e}"); continue

        required_rules = resolve_qos_standard(standards, args.org_id, network_id)

        oid_l, nid_l = args.org_id.strip().lower(), network_id.strip().lower()
        std_level = "network" if standards.get((oid_l, nid_l)) else (
                    "org"     if standards.get((oid_l, ""))    else "")

        rows = audit_network_qos(
            args.org_id, org_name, network_id, network_name,
            live_rules, required_rules, std_level, ignored_fields,
        )
        _print_results(network_name, network_id, rows)
        all_detail.extend(rows)

    summary_rows = build_summary(all_detail)
    print()
    print(f"TOTAL — PASS={sum(r['tot_pass'] for r in summary_rows)}  "
          f"FAIL={sum(r['tot_fail'] for r in summary_rows)}  "
          f"WARNINGS={sum(1 for r in summary_rows if r['overall'] in ('NON_STANDARD','MISSING'))}")

    suffix = f"_{args.network_id}" if args.network_id else ""
    if args.csv:
        write_csv(OUTPUT_DIR / f"audit_qos_{args.org_id}{suffix}_detail.csv",  all_detail,   _DETAIL_HEADERS)
        write_csv(OUTPUT_DIR / f"audit_qos_{args.org_id}{suffix}_summary.csv", summary_rows, _SUMMARY_HEADERS)
    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_qos_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp, _DETAIL_HEADERS, _SUMMARY_HEADERS, "QoSAudit")
        print(f"XLSX → {xp}")


if __name__ == "__main__":
    main()
