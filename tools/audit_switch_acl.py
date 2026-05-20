"""
audit_switch_acl.py
Audit switch ACL rules against required rules defined in the SwitchACL sheet.

API: getNetworkSwitchAccessControlLists(networkId)
Response: {"rules": [{comment, policy, ipVersion, protocol, srcCidr,
                       srcPort, dstCidr, dstPort, vlan}, ...]}

The last rule is always the implicit default allow-any added by Meraki.

Audit logic:
  - Each row in SwitchACL defines a REQUIRED rule identified by rule_comment.
  - For every required rule: check it exists in live and all fields match.
  - Any live rule whose comment is NOT in the standard is flagged as NON_STANDARD
    (unexpected rule — may indicate a policy drift).
  - The implicit default allow-any is always skipped.

Standard sheet columns:
  org_id, network_id (optional), rule_comment (key), policy, ipVersion,
  protocol, srcCidr, srcPort, dstCidr, dstPort, vlan

Usage:
    python -m tools.audit_switch_acl \\
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
    load_acl_standards,
    load_ignored_fields,
    resolve_acl_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

_RULE_FIELDS = ["policy", "ipVersion", "protocol", "srcCidr", "srcPort",
                "dstCidr", "dstPort", "vlan"]

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key",    # rule_comment
    "entity_label",  # rule_comment (same — rules have no numeric id)
    "standard_level", "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]

# Meraki always appends this implicit rule — never treat it as unexpected
_DEFAULT_RULE_COMMENT = "default rule"


def _is_default_rule(rule: dict) -> bool:
    comment = (rule.get("comment") or "").strip().lower()
    if comment == _DEFAULT_RULE_COMMENT:
        return True
    # Also detect by content — allow any/any on all ports (the true implicit default)
    return (
        rule.get("policy") == "allow"
        and str(rule.get("srcCidr",  "")).lower() == "any"
        and str(rule.get("dstCidr",  "")).lower() == "any"
        and str(rule.get("dstPort",  "")).lower() in ("any", "")
        and str(rule.get("srcPort",  "")).lower() in ("any", "")
        and str(rule.get("protocol", "")).lower() in ("any", "")
    )


def _normalise_rule_field(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def audit_network_acl(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    live_rules: list[dict],
    required_rules: list[dict],
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    detail_rows: list[dict] = []

    # Build lookup: comment (lower) → live rule
    live_by_comment = {
        (r.get("comment") or "").strip().lower(): r
        for r in live_rules
        if not _is_default_rule(r)
    }

    def _row(comment, field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=comment, entity_label=comment,
            field=field, expected=expected, actual=actual,
            result=result, standard_level=std_level,
        )

    # ── Check each required rule ─────────────────────────────────────────────
    required_comments = set()
    for req in required_rules:
        comment = req.get("rule_comment", "").strip()
        comment_l = comment.lower()
        required_comments.add(comment_l)
        live_rule = live_by_comment.get(comment_l)

        if live_rule is None:
            # Rule is required but missing from live config
            for field in _RULE_FIELDS:
                detail_rows.append(_row(comment, field, req.get(field, ""), "", "MISSING"))
            continue

        # Rule exists — check each field
        for field in _RULE_FIELDS:
            if field in ignored_fields:
                detail_rows.append(_row(comment, field, "", str(live_rule.get(field, "")), "IGNORED"))
                continue
            expected = req.get(field, NOT_DEFINED)
            actual   = str(live_rule.get(field, "") or "")
            if expected == NOT_DEFINED:
                detail_rows.append(_row(comment, field, "", actual, NOT_DEFINED))
            else:
                result = "PASS" if _normalise_rule_field(expected) == _normalise_rule_field(actual) else "FAIL"
                detail_rows.append(_row(comment, field, expected, actual, result))

    # ── Flag unexpected live rules (not in standard, not the default) ────────
    for comment_l, live_rule in live_by_comment.items():
        if comment_l not in required_comments:
            comment_display = live_rule.get("comment") or comment_l
            for field in _RULE_FIELDS:
                detail_rows.append(_row(
                    comment_display, field,
                    "", str(live_rule.get(field, "") or ""),
                    "NON_STANDARD",
                ))

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

    for rule_comment in rules_seen:
        rule_rows = [r for r in detail_rows if r["entity_key"] == rule_comment]
        has_issue = any(r["result"] not in ("PASS", NOT_DEFINED, "NOT_DEFINED", "IGNORED")
                        for r in rule_rows)
        if not has_issue:
            continue
        print(f"\n    Rule: {rule_comment!r}")
        for r in rule_rows:
            result = r["result"]
            if result == "FAIL":
                print(f"      ✗ FAIL        {r['field']:<12} "
                      f"expected={r['expected']!r}  actual={r['actual']!r}")
            elif result == "MISSING":
                print(f"      ✗ MISSING     {r['field']:<12} expected={r['expected']!r}")
            elif result == "NON_STANDARD":
                print(f"      ⚠ UNEXPECTED  {r['field']:<12} actual={r['actual']!r}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit switch ACL rules")
    parser.add_argument("--org-id",         required=True)
    parser.add_argument("--standards-file", required=True)
    parser.add_argument("--network-id",     default=None)
    parser.add_argument("--csv",  action="store_true")
    parser.add_argument("--xlsx", action="store_true")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: {standards_path} not found"); sys.exit(1)

    print(f"Loading ACL standards from {standards_path.name}...")
    standards      = load_acl_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key  = args.org_id.strip().lower()
    org_stds = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_stds:
        print(f"WARNING: no ACL standards for org '{args.org_id}'. Add rows to 'SwitchACL'.")
    else:
        total_rules = sum(len(v) for v in org_stds.values())
        print(f"  {total_rules} required rule(s) across {len(org_stds)} standard(s).")
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
            acl_resp   = dashboard.switch.getNetworkSwitchAccessControlLists(network_id)
            live_rules = acl_resp.get("rules", [])
        except Exception as e:
            print(f"  [{network_name}] skipped — {e}"); continue

        required_rules = resolve_acl_standard(standards, args.org_id, network_id)

        oid_l, nid_l = args.org_id.strip().lower(), network_id.strip().lower()
        std_level = "network" if standards.get((oid_l, nid_l)) else (
                    "org"     if standards.get((oid_l, ""))    else "")

        rows = audit_network_acl(
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
        write_csv(OUTPUT_DIR / f"audit_acl_{args.org_id}{suffix}_detail.csv",  all_detail,   _DETAIL_HEADERS)
        write_csv(OUTPUT_DIR / f"audit_acl_{args.org_id}{suffix}_summary.csv", summary_rows, _SUMMARY_HEADERS)
    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_acl_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp, _DETAIL_HEADERS, _SUMMARY_HEADERS, "ACLAudit")
        print(f"XLSX → {xp}")


if __name__ == "__main__":
    main()
