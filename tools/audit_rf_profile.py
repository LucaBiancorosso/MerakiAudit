"""
audit_rf_profile.py
Audit RF profile configuration against standards defined in the RFProfiles
sheet of the standards Excel file.

Usage:
    python -m tools.audit_rf_profile --org-id <ORG_ID> \
        --standards-file standards/ssid_standards_template.xlsx --xlsx
    python -m tools.audit_rf_profile --org-id <ORG_ID> \
        --standards-file standards/ssid_standards_template.xlsx \
        --network-id <NET_ID> --xlsx
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
    load_rf_standards,
    resolve_rf_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

# ---------------------------------------------------------------------------
# Fields audited — flat names after we flatten the nested API response
# ---------------------------------------------------------------------------
_AUDITED_FIELDS = [
    # General
    "clientBalancingEnabled",
    "minBitrateType",
    "bandSelectionType",
    "transmission_enabled",
    "isIndoorDefault",
    "isOutdoorDefault",
    # AP band settings
    "ap_bandOperationMode",
    "ap_bandSteeringEnabled",
    "ap_bands_enabled",        # comma-sep sorted list, e.g. "2.4,5,6"
    # 2.4 GHz
    "2g_maxPower",
    "2g_minPower",
    "2g_minBitrate",
    "2g_validAutoChannels",    # comma-sep sorted, e.g. "1,6,11"
    "2g_axEnabled",
    "2g_rxsop",
    # 5 GHz
    "5g_maxPower",
    "5g_minPower",
    "5g_minBitrate",
    "5g_channelWidth",
    "5g_validAutoChannels",
    "5g_rxsop",
    # 6 GHz (Wi-Fi 6E / 6 GHz capable APs only)
    "6g_maxPower",
    "6g_minPower",
    "6g_minBitrate",
    "6g_channelWidth",
    "6g_validAutoChannels",
    "6g_rxsop",
]

# Fields where we compare as unordered comma-separated sets (channels, bands)
_SET_FIELDS = {
    "ap_bands_enabled",
    "2g_validAutoChannels",
    "5g_validAutoChannels",
    "6g_validAutoChannels",
}

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key",    # profile_name
    "entity_label",  # profile_id
    "standard_level", "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]


# ---------------------------------------------------------------------------
# Flatten nested RF profile dict → flat field dict
# ---------------------------------------------------------------------------

def flatten_rf_profile(p: dict) -> dict[str, str]:
    """Convert the nested API response into the flat field names we audit."""

    def _s(v) -> str:
        """Safe string — handles None, bool, int, list."""
        if v is None:
            return ""
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, list):
            # Sort numerically when possible (channel numbers: 1,6,11 not 1,11,6)
            try:
                return ",".join(str(i) for i in sorted(v, key=lambda x: int(str(x))))
            except (ValueError, TypeError):
                return ",".join(sorted(str(i) for i in v))
        if isinstance(v, float):
            return str(int(v)) if v == int(v) else str(v)
        return str(v)

    ap  = p.get("apBandSettings") or {}
    g2  = p.get("twoFourGhzSettings") or {}
    g5  = p.get("fiveGhzSettings") or {}
    g6  = p.get("sixGhzSettings") or {}
    tx  = p.get("transmission") or {}

    return {
        "clientBalancingEnabled": _s(p.get("clientBalancingEnabled")),
        "minBitrateType":         _s(p.get("minBitrateType")),
        "bandSelectionType":      _s(p.get("bandSelectionType")),
        "transmission_enabled":   _s(tx.get("enabled")),
        "isIndoorDefault":        _s(p.get("isIndoorDefault")),
        "isOutdoorDefault":       _s(p.get("isOutdoorDefault")),
        # AP band
        "ap_bandOperationMode":   _s(ap.get("bandOperationMode")),
        "ap_bandSteeringEnabled": _s(ap.get("bandSteeringEnabled")),
        "ap_bands_enabled":       _s((ap.get("bands") or {}).get("enabled")),
        # 2.4 GHz
        "2g_maxPower":            _s(g2.get("maxPower")),
        "2g_minPower":            _s(g2.get("minPower")),
        "2g_minBitrate":          _s(g2.get("minBitrate")),
        "2g_validAutoChannels":   _s(g2.get("validAutoChannels")),
        "2g_axEnabled":           _s(g2.get("axEnabled")),
        "2g_rxsop":               _s(g2.get("rxsop")),
        # 5 GHz
        "5g_maxPower":            _s(g5.get("maxPower")),
        "5g_minPower":            _s(g5.get("minPower")),
        "5g_minBitrate":          _s(g5.get("minBitrate")),
        "5g_channelWidth":        _s(g5.get("channelWidth")),
        "5g_validAutoChannels":   _s(g5.get("validAutoChannels")),
        "5g_rxsop":               _s(g5.get("rxsop")),
        # 6 GHz
        "6g_maxPower":            _s(g6.get("maxPower")),
        "6g_minPower":            _s(g6.get("minPower")),
        "6g_minBitrate":          _s(g6.get("minBitrate")),
        "6g_channelWidth":        _s(g6.get("channelWidth")),
        "6g_validAutoChannels":   _s(g6.get("validAutoChannels")),
        "6g_rxsop":               _s(g6.get("rxsop")),
    }


# ---------------------------------------------------------------------------
# Compare with set-aware override for channel/band fields
# ---------------------------------------------------------------------------

def _compare(field: str, expected: str, actual: str, ignored: set[str]) -> dict:
    if field in _SET_FIELDS and expected not in (NOT_DEFINED, "") and field not in ignored:
        # Compare as unordered sets
        a_set = {v.strip() for v in actual.split(",") if v.strip()}
        e_set = {v.strip() for v in expected.split(",") if v.strip()}
        result = "PASS" if a_set == e_set else "FAIL"
        return {"field": field, "expected": expected, "actual": actual, "result": result}
    return compare_field(field, expected, actual, ignored)


# ---------------------------------------------------------------------------
# Audit one network's RF profiles
# ---------------------------------------------------------------------------

def audit_network_rf(
    org_id: str,
    org_name: str,
    network_id: str,
    network_name: str,
    rf_profiles: list[dict],
    standards: dict,
    ignored_fields: set[str],
    single_network_mode: bool = False,
) -> list[dict]:
    detail_rows: list[dict] = []

    oid = org_id.strip().lower()
    nid = network_id.strip().lower()

    # Names defined in standard for this org+network (for MISSING check)
    standard_names = {
        pname
        for (o, n, pname) in standards
        if o == oid and n in ("", nid)
    }

    live_by_name = {p["name"].lower(): p for p in rf_profiles if p.get("name")}

    for profile in rf_profiles:
        profile_name = profile.get("name", "")
        profile_id   = str(profile.get("id", ""))
        flat         = flatten_rf_profile(profile)
        standard     = resolve_rf_standard(standards, org_id, network_id, profile_name)

        # Determine standard level for reporting
        pname_l = profile_name.strip().lower()
        if standards.get((oid, nid, pname_l)):
            std_level = "network"
        elif standards.get((oid, "", pname_l)):
            std_level = "org"
        else:
            std_level = ""

        if standard is None:
            # RF profile exists live but not in standards → NON_STANDARD
            for field in _AUDITED_FIELDS:
                detail_rows.append(make_row(
                    org_id, org_name, network_id, network_name,
                    entity_key=profile_name, entity_label=profile_id,
                    field=field, expected="", actual=flat.get(field, ""),
                    result="NON_STANDARD", standard_level="",
                ))
            continue

        for field in _AUDITED_FIELDS:
            if field in ignored_fields:
                continue
            expected = standard.get(field, NOT_DEFINED)
            actual   = flat.get(field, "")
            check    = _compare(field, expected, actual, ignored_fields)
            detail_rows.append(make_row(
                org_id, org_name, network_id, network_name,
                entity_key=profile_name, entity_label=profile_id,
                field=field,
                expected=check["expected"],
                actual=check["actual"],
                result=check["result"],
                standard_level=std_level,
            ))

    # MISSING check: profile in standard but not found live
    if not single_network_mode:
        for pname in standard_names:
            if pname not in live_by_name:
                ml = "network" if standards.get((oid, nid, pname)) else "org"
                detail_rows.append(make_row(
                    org_id, org_name, network_id, network_name,
                    entity_key=pname, entity_label="",
                    field="—", expected="", actual="",
                    result="MISSING", standard_level=ml,
                ))

    return detail_rows


# ---------------------------------------------------------------------------
# Console output
# ---------------------------------------------------------------------------

def _print_network_results(network_name: str, network_id: str, detail_rows: list[dict]) -> None:
    pass_c  = sum(1 for r in detail_rows if r["result"] == "PASS")
    fail_c  = sum(1 for r in detail_rows if r["result"] == "FAIL")
    nd_c    = sum(1 for r in detail_rows if r["result"] in (NOT_DEFINED, "NOT_DEFINED"))
    ign_c   = sum(1 for r in detail_rows if r["result"] == "IGNORED")
    warn_profiles = {r["entity_key"] for r in detail_rows
                     if r["result"] in ("NON_STANDARD", "MISSING")}

    print(f"  [{network_name}]  {network_id}")
    print(f"    PASS={pass_c}  FAIL={fail_c}  NOT_DEFINED={nd_c}  "
          f"IGNORED={ign_c}  WARNINGS={len(warn_profiles)}")

    profiles_seen: list[str] = []
    for r in detail_rows:
        if r["entity_key"] not in profiles_seen:
            profiles_seen.append(r["entity_key"])

    for pname in profiles_seen:
        rows = [r for r in detail_rows if r["entity_key"] == pname]
        pid  = rows[0]["entity_label"]
        print(f"\n    RF Profile {pname!r}  (id={pid})")
        for r in rows:
            result = r["result"]
            icon   = {"PASS": "✓", "FAIL": "✗"}.get(result, "⚠" if result in ("NON_STANDARD","MISSING") else "–")
            if result == "MISSING":
                print(f"      {icon} MISSING")
            elif result == "NON_STANDARD":
                print(f"      {icon} NON_STANDARD   {r['field']:<32} actual={r['actual']!r}")
            elif result == "FAIL":
                print(f"      {icon} FAIL           {r['field']:<32} "
                      f"expected={r['expected']!r}  actual={r['actual']!r}")
            elif result in (NOT_DEFINED, "NOT_DEFINED"):
                print(f"      – NOT_DEFINED   {r['field']:<32} (not audited)")
            elif result == "IGNORED":
                print(f"      – IGNORED       {r['field']:<32}")
            else:
                print(f"      ✓ PASS          {r['field']:<32} value={r['actual']!r}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit RF profile configuration against compliance standards"
    )
    parser.add_argument("--org-id",         required=True,  help="Meraki Organization ID")
    parser.add_argument("--standards-file", required=True,  help="Path to standards Excel file")
    parser.add_argument("--network-id",     default=None,   help="Optional: audit a single network")
    parser.add_argument("--csv",            action="store_true", help="Write CSV output")
    parser.add_argument("--xlsx",           action="store_true", help="Write XLSX report")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: standards file not found: {standards_path}")
        sys.exit(1)

    print(f"Loading RF profile standards from {standards_path.name}...")
    standards      = load_rf_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key = args.org_id.strip().lower()
    org_standards = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_standards:
        print(f"WARNING: no RF profile standards for org '{args.org_id}' in template.")
        print(f"  Add rows to the 'RFProfiles' sheet.")
    else:
        print(f"  {len(org_standards)} RF profile standard(s) found.")
        for (o, n, pname) in org_standards:
            level = f"network={n}" if n else "org-wide"
            print(f"    profile={pname!r}  [{level}]")
    print()

    dashboard      = get_dashboard()
    network_lookup = build_network_lookup(dashboard, args.org_id)

    # Resolve org name
    org_name = ""
    try:
        for o in dashboard.organizations.getOrganizations():
            if str(o.get("id")) == str(args.org_id):
                org_name = o.get("name", ""); break
    except Exception:
        pass

    if args.network_id:
        if args.network_id not in network_lookup:
            print(f"ERROR: network_id '{args.network_id}' not found.")
            sys.exit(1)
        networks_to_audit = {args.network_id: network_lookup[args.network_id]}
    else:
        networks_to_audit = dict(network_lookup)

    print(f"Auditing {len(networks_to_audit)} network(s) in org '{org_name or args.org_id}'...")
    print()

    all_detail: list[dict] = []

    for network_id, network_name in sorted(networks_to_audit.items(), key=lambda x: x[1]):
        try:
            profiles = dashboard.wireless.getNetworkWirelessRfProfiles(network_id)
        except Exception as e:
            print(f"  [{network_name}] skipped — {e}")
            continue

        if not profiles:
            print(f"  [{network_name}] no RF profiles found, skipping.")
            continue

        rows = audit_network_rf(
            org_id=args.org_id, org_name=org_name,
            network_id=network_id, network_name=network_name,
            rf_profiles=profiles, standards=standards,
            ignored_fields=ignored_fields,
            single_network_mode=bool(args.network_id),
        )
        _print_network_results(network_name, network_id, rows)
        all_detail.extend(rows)

    summary_rows = build_summary(all_detail)

    print()
    tot_pass = sum(r["tot_pass"] for r in summary_rows)
    tot_fail = sum(r["tot_fail"] for r in summary_rows)
    tot_nd   = sum(r["tot_not_defined"] for r in summary_rows)
    tot_warn = sum(1 for r in summary_rows if r["overall"] in ("NON_STANDARD", "MISSING"))
    print(f"TOTAL — PASS={tot_pass}  FAIL={tot_fail}  NOT_DEFINED={tot_nd}  WARNINGS={tot_warn}")

    suffix = f"_{args.network_id}" if args.network_id else ""

    if args.csv:
        dp = OUTPUT_DIR / f"audit_rf_{args.org_id}{suffix}_detail.csv"
        sp = OUTPUT_DIR / f"audit_rf_{args.org_id}{suffix}_summary.csv"
        write_csv(dp, all_detail,    _DETAIL_HEADERS)
        write_csv(sp, summary_rows,  _SUMMARY_HEADERS)
        print(f"\nCSV detail  → {dp}")
        print(f"CSV summary → {sp}")

    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_rf_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp,
                          _DETAIL_HEADERS, _SUMMARY_HEADERS, "RFAudit")
        print(f"XLSX report → {xp}")


if __name__ == "__main__":
    main()
