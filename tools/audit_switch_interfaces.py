"""
audit_switch_interfaces.py
Audit switch port configuration against role-based standards.

Role detection priority per port:
  1. CDP/LLDP neighbour platform matches cdp_ap_pattern  → trunk_ap
  2. CDP/LLDP neighbour platform matches cdp_cpe_pattern → trunk_cpe
  3. No CDP match + type==trunk                          → trunk_unknown  (NON_STANDARD)
  4. No CDP match + type==access → VLAN looked up in VlanRoles sheet
       match found  → access_<role>  (e.g. access_data, access_voice)
       no match     → access_unknown (NON_STANDARD)

Standards are keyed by (org_id, network_id, role) with the usual
network-specific → org-wide fallback from resolve_switch_standard().

Usage:
    python -m tools.audit_switch_interfaces \\
        --org-id <ORG_ID> \\
        --standards-file standards/standard_audit_fields.xlsx --xlsx

    python -m tools.audit_switch_interfaces \\
        --org-id <ORG_ID> \\
        --standards-file standards/standard_audit_fields.xlsx \\
        --network-id <NETWORK_ID> --xlsx --csv
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv
from lib.standards import (
    NOT_DEFINED,
    compare_field,
    expand_vlan_ranges,
    load_ignored_fields,
    load_switch_standards,
    load_vlan_roles,
    resolve_switch_standard,
    resolve_vlan_role,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

# ---------------------------------------------------------------------------
# Fields audited — directly from getDeviceSwitchPorts response
# ---------------------------------------------------------------------------
_AUDITED_FIELDS = [
    # Port mode
    "enabled",
    "type",
    # VLAN assignment
    "vlan",
    "voiceVlan",
    "allowedVlans",       # trunk — compared as expanded VLAN set
    # Layer 2 features
    "poeEnabled",
    "isolationEnabled",
    "rstpEnabled",
    "stpGuard",           # disabled / root guard / bpdu guard / loop guard
    "stpPortFastTrunk",   # trunk only
    "udld",               # Disabled / Alert only / Enforce
    "stormControlEnabled",
    "daiTrusted",         # Dynamic ARP Inspection trust — trunk only
    # Access control
    "accessPolicyType",   # Open / Custom access policy / MAC allow list / Sticky MAC allow list
    # Physical
    "linkNegotiation",
    "dot3az_enabled",     # Energy Efficient Ethernet (flattened from dot3az.enabled)
]

# Fields compared as expanded VLAN sets
_VLAN_SET_FIELDS = {"allowedVlans"}

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "switch_serial", "switch_name",
    "port_id", "port_name",
    "detected_role", "standard_level",
    "cdp_platform", "cdp_system_name",
    "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key",    # switch_serial:port_id
    "entity_label",  # switch_name / port_name
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]


# ---------------------------------------------------------------------------
# VLAN set comparison
# ---------------------------------------------------------------------------

def _vlan_sets_equal(expected_str: str, actual_str: str) -> bool:
    return expand_vlan_ranges(expected_str) == expand_vlan_ranges(actual_str)


# ---------------------------------------------------------------------------
# CDP / LLDP role detection
# ---------------------------------------------------------------------------

def _cdp_platform(status: dict) -> str:
    """Return the best available neighbour platform string from CDP then LLDP."""
    cdp  = status.get("cdp")  or {}
    lldp = status.get("lldp") or {}
    return (
        cdp.get("platform")
        or lldp.get("systemDescription")
        or cdp.get("systemName")
        or lldp.get("systemName")
        or ""
    )


def _cdp_system_name(status: dict) -> str:
    cdp  = status.get("cdp")  or {}
    lldp = status.get("lldp") or {}
    return cdp.get("systemName") or lldp.get("systemName") or ""


def _detect_role(
    port_config: dict,
    port_status: dict,
    ap_pattern:  re.Pattern | None,
    cpe_pattern: re.Pattern | None,
    vlan_roles:  dict,
    org_id:      str,
    network_id:  str,
) -> str:
    """Return the inferred role string for a port."""
    platform = _cdp_platform(port_status)

    if platform:
        if ap_pattern  and ap_pattern.search(platform):
            return "trunk_ap"
        if cpe_pattern and cpe_pattern.search(platform):
            return "trunk_cpe"

    port_type = (port_config.get("type") or "").lower()

    if port_type == "trunk":
        return "trunk_unknown"

    # Access port — look up VLAN in VlanRoles
    vlan = port_config.get("vlan")
    if vlan is not None:
        role = resolve_vlan_role(vlan_roles, org_id, network_id, vlan)
        if role:
            return role

    return "access_unknown"


# ---------------------------------------------------------------------------
# Compile patterns from standards rows
# ---------------------------------------------------------------------------

def _compile_patterns(
    standards: dict,
    org_id: str,
    network_id: str,
) -> tuple[re.Pattern | None, re.Pattern | None]:
    """
    Pull cdp_ap_pattern and cdp_cpe_pattern from the trunk_ap / trunk_cpe
    standard rows and compile them. Network-specific rows take priority.
    """
    def _get_field(role: str, field: str) -> str:
        std = resolve_switch_standard(standards, org_id, network_id, role)
        if std:
            val = std.get(field, NOT_DEFINED)
            if val and val != NOT_DEFINED:
                return val
        return ""

    ap_pat_str  = _get_field("trunk_ap",  "cdp_ap_pattern")
    cpe_pat_str = _get_field("trunk_cpe", "cdp_cpe_pattern")

    ap_pat  = None
    cpe_pat = None

    try:
        if ap_pat_str:
            ap_pat = re.compile(ap_pat_str, re.IGNORECASE)
    except re.error as e:
        print(f"  WARNING: invalid cdp_ap_pattern regex {ap_pat_str!r}: {e}")

    try:
        if cpe_pat_str:
            cpe_pat = re.compile(cpe_pat_str, re.IGNORECASE)
    except re.error as e:
        print(f"  WARNING: invalid cdp_cpe_pattern regex {cpe_pat_str!r}: {e}")

    return ap_pat, cpe_pat


# ---------------------------------------------------------------------------
# Flatten port config → flat field dict
# ---------------------------------------------------------------------------

def _flatten_port(port: dict) -> dict[str, str]:
    def _s(v) -> str:
        if v is None:
            return ""
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, float):
            return str(int(v)) if v == int(v) else str(v)
        return str(v).strip()

    dot3az = port.get("dot3az") or {}

    return {
        "enabled":          _s(port.get("enabled")),
        "type":             _s(port.get("type")),
        "vlan":             _s(port.get("vlan")),
        "voiceVlan":        _s(port.get("voiceVlan")),
        "allowedVlans":     _s(port.get("allowedVlans")),
        "poeEnabled":       _s(port.get("poeEnabled")),
        "isolationEnabled": _s(port.get("isolationEnabled")),
        "rstpEnabled":      _s(port.get("rstpEnabled")),
        "stpGuard":         _s(port.get("stpGuard")),
        "stpPortFastTrunk": _s(port.get("stpPortFastTrunk")),
        "udld":             _s(port.get("udld")),
        "stormControlEnabled": _s(port.get("stormControlEnabled")),
        "daiTrusted":       _s(port.get("daiTrusted")),
        "accessPolicyType": _s(port.get("accessPolicyType")),
        "linkNegotiation":  _s(port.get("linkNegotiation")),
        "dot3az_enabled":   _s(dot3az.get("enabled")),
    }


# ---------------------------------------------------------------------------
# Compare one field — vlan-set aware
# ---------------------------------------------------------------------------

def _compare(field: str, expected: str, actual: str, ignored: set[str]) -> dict:
    if field in _VLAN_SET_FIELDS and expected not in (NOT_DEFINED, "") and field not in ignored:
        result = "PASS" if _vlan_sets_equal(expected, actual) else "FAIL"
        return {"field": field, "expected": expected, "actual": actual, "result": result}
    return compare_field(field, expected, actual, ignored)


# ---------------------------------------------------------------------------
# Audit one port
# ---------------------------------------------------------------------------

def _audit_port(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    switch_serial: str, switch_name: str,
    port_config: dict,
    port_status: dict,
    role: str,
    standard: dict | None,
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    port_id   = str(port_config.get("portId", ""))
    port_name = port_config.get("name", "") or port_id
    flat      = _flatten_port(port_config)
    cdp_plat  = _cdp_platform(port_status)
    cdp_name  = _cdp_system_name(port_status)

    entity_key   = f"{switch_serial}:{port_id}"
    entity_label = f"{switch_name} / {port_name}"

    extra = {
        "switch_serial":   switch_serial,
        "switch_name":     switch_name,
        "port_id":         port_id,
        "port_name":       port_name,
        "detected_role":   role,
        "cdp_platform":    cdp_plat,
        "cdp_system_name": cdp_name,
    }

    def _row(field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=entity_key,
            entity_label=entity_label,
            field=field,
            expected=expected,
            actual=actual,
            result=result,
            standard_level=std_level,
            extra=extra,
        )

    if standard is None:
        # Port has no standard for its role
        rows_out = []
        for field in _AUDITED_FIELDS:
            rows_out.append(_row(field, "", flat.get(field, ""), "NON_STANDARD"))
        return rows_out

    rows_out = []
    for field in _AUDITED_FIELDS:
        if field in ignored_fields:
            rows_out.append(_row(field, "", flat.get(field, ""), "IGNORED"))
            continue
        expected = standard.get(field, NOT_DEFINED)
        actual   = flat.get(field, "")
        check    = _compare(field, expected, actual, ignored_fields)
        rows_out.append(_row(field, check["expected"], check["actual"], check["result"]))

    return rows_out


# ---------------------------------------------------------------------------
# Audit one switch
# ---------------------------------------------------------------------------

def audit_switch(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    switch_serial: str, switch_name: str,
    port_configs: list[dict],
    port_statuses: list[dict],
    standards: dict,
    vlan_roles: dict,
    ignored_fields: set[str],
) -> list[dict]:
    # Build status lookup by portId
    status_by_port = {str(s.get("portId", "")): s for s in port_statuses}

    # Compile regex patterns for this network
    ap_pat, cpe_pat = _compile_patterns(standards, org_id, network_id)

    detail_rows: list[dict] = []

    for port in port_configs:
        port_id = str(port.get("portId", ""))
        status  = status_by_port.get(port_id, {})

        role = _detect_role(port, status, ap_pat, cpe_pat, vlan_roles, org_id, network_id)

        standard  = resolve_switch_standard(standards, org_id, network_id, role)
        oid_l     = org_id.strip().lower()
        nid_l     = network_id.strip().lower()
        role_l    = role.strip().lower()
        if standards.get((oid_l, nid_l, role_l)):
            std_level = "network"
        elif standards.get((oid_l, "", role_l)):
            std_level = "org"
        else:
            std_level = ""

        rows = _audit_port(
            org_id, org_name, network_id, network_name,
            switch_serial, switch_name,
            port, status, role, standard, std_level, ignored_fields,
        )
        detail_rows.extend(rows)

    return detail_rows


# ---------------------------------------------------------------------------
# Console output
# ---------------------------------------------------------------------------

def _print_network_results(
    network_name: str,
    network_id: str,
    detail_rows: list[dict],
) -> None:
    pass_c = sum(1 for r in detail_rows if r["result"] == "PASS")
    fail_c = sum(1 for r in detail_rows if r["result"] == "FAIL")
    nd_c   = sum(1 for r in detail_rows if r["result"] in (NOT_DEFINED, "NOT_DEFINED"))
    ign_c  = sum(1 for r in detail_rows if r["result"] == "IGNORED")
    warn_p = {r["entity_key"] for r in detail_rows if r["result"] == "NON_STANDARD"}

    print(f"  [{network_name}]  {network_id}")
    print(f"    PASS={pass_c}  FAIL={fail_c}  NOT_DEFINED={nd_c}  "
          f"IGNORED={ign_c}  WARNINGS={len(warn_p)}")

    # Group by switch then port
    switches_seen: list[str] = []
    for r in detail_rows:
        sw = r["switch_serial"]
        if sw not in switches_seen:
            switches_seen.append(sw)

    for sw_serial in switches_seen:
        sw_rows = [r for r in detail_rows if r["switch_serial"] == sw_serial]
        sw_name = sw_rows[0]["switch_name"]
        sw_fail = sum(1 for r in sw_rows if r["result"] == "FAIL")
        sw_warn = len({r["entity_key"] for r in sw_rows if r["result"] == "NON_STANDARD"})
        print(f"\n    Switch {sw_name!r}  serial={sw_serial}  "
              f"FAIL={sw_fail}  WARNINGS={sw_warn}")

        ports_seen: list[str] = []
        for r in sw_rows:
            if r["entity_key"] not in ports_seen:
                ports_seen.append(r["entity_key"])

        for ek in ports_seen:
            port_rows = [r for r in sw_rows if r["entity_key"] == ek]
            first = port_rows[0]
            role  = first["detected_role"]
            cdp   = first["cdp_platform"]
            cdp_s = f"  cdp={cdp!r}" if cdp else ""
            has_issue = any(r["result"] in ("FAIL", "NON_STANDARD") for r in port_rows)

            # Only print ports with issues unless verbose wanted
            if not has_issue:
                continue

            print(f"      Port {first['port_id']:>4}  {first['port_name']:<20}  "
                  f"role={role}{cdp_s}")
            for r in port_rows:
                result = r["result"]
                if result == "NON_STANDARD":
                    print(f"        ⚠ NON_STANDARD  {r['field']:<22} actual={r['actual']!r}")
                elif result == "FAIL":
                    print(f"        ✗ FAIL          {r['field']:<22} "
                          f"expected={r['expected']!r}  actual={r['actual']!r}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit switch port configuration against role-based standards"
    )
    parser.add_argument("--org-id",         required=True,
                        help="Meraki Organization ID")
    parser.add_argument("--standards-file", required=True,
                        help="Path to standards Excel file")
    parser.add_argument("--network-id",     default=None,
                        help="Optional: limit audit to a single network")
    parser.add_argument("--include-disabled", action="store_true",
                        help="Include admin-disabled ports (skipped by default)")
    parser.add_argument("--csv",  action="store_true", help="Write CSV output")
    parser.add_argument("--xlsx", action="store_true", help="Write XLSX report")
    args = parser.parse_args()

    standards_path = Path(args.standards_file)
    if not standards_path.exists():
        print(f"ERROR: standards file not found: {standards_path}")
        sys.exit(1)

    print(f"Loading switch standards from {standards_path.name}...")
    standards      = load_switch_standards(standards_path)
    vlan_roles     = load_vlan_roles(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key = args.org_id.strip().lower()
    org_standards = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_standards:
        print(f"WARNING: no switch standards for org '{args.org_id}'. "
              "Add rows to the 'SwitchInterfaces' sheet.")
    else:
        print(f"  {len(org_standards)} role standard(s) found:")
        for (o, n, role) in sorted(org_standards):
            level = f"network={n}" if n else "org-wide"
            print(f"    role={role:<20} [{level}]")

    org_vlan_roles = {k: v for k, v in vlan_roles.items() if k[0] == org_key}
    print(f"  {len(org_vlan_roles)} VLAN role mapping(s) found.")
    if ignored_fields:
        print(f"  Ignored fields: {', '.join(sorted(ignored_fields))}")
    print()

    dashboard      = get_dashboard()
    network_lookup = build_network_lookup(dashboard, args.org_id)

    # Org name
    org_name = ""
    try:
        for o in dashboard.organizations.getOrganizations():
            if str(o.get("id")) == str(args.org_id):
                org_name = o.get("name", "")
                break
    except Exception:
        pass

    if args.network_id:
        if args.network_id not in network_lookup:
            print(f"ERROR: network_id '{args.network_id}' not found in org.")
            sys.exit(1)
        networks_to_audit = {args.network_id: network_lookup[args.network_id]}
    else:
        networks_to_audit = dict(network_lookup)

    # Fetch all switch serials per network
    print(f"Fetching switch inventory for org '{org_name or args.org_id}'...")
    try:
        all_switches = dashboard.organizations.getOrganizationDevices(
            args.org_id, total_pages="all", productTypes=["switch"]
        )
    except Exception as e:
        print(f"ERROR fetching devices: {e}")
        sys.exit(1)

    # Group by network
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
            print(f"  [{network_name}] no switches, skipping.")
            continue

        net_rows: list[dict] = []

        for sw in switches:
            serial  = sw.get("serial", "")
            sw_name = sw.get("name", serial)

            try:
                port_configs = dashboard.switch.getDeviceSwitchPorts(serial)
            except Exception as e:
                print(f"    [{sw_name}] could not fetch ports: {e}")
                continue

            try:
                port_statuses = dashboard.switch.getDeviceSwitchPortsStatuses(serial)
            except Exception as e:
                print(f"    [{sw_name}] could not fetch port statuses: {e}")
                port_statuses = []

            # Filter disabled ports unless --include-disabled
            if not args.include_disabled:
                port_configs = [p for p in port_configs if p.get("enabled", True)]

            sw_rows = audit_switch(
                org_id=args.org_id, org_name=org_name,
                network_id=network_id, network_name=network_name,
                switch_serial=serial, switch_name=sw_name,
                port_configs=port_configs,
                port_statuses=port_statuses,
                standards=standards,
                vlan_roles=vlan_roles,
                ignored_fields=ignored_fields,
            )
            net_rows.extend(sw_rows)

        _print_network_results(network_name, network_id, net_rows)
        all_detail.extend(net_rows)

    summary_rows = build_summary(all_detail)

    print()
    tot_pass = sum(r["tot_pass"]           for r in summary_rows)
    tot_fail = sum(r["tot_fail"]           for r in summary_rows)
    tot_nd   = sum(r["tot_not_defined"]    for r in summary_rows)
    tot_warn = sum(1 for r in summary_rows if r["overall"] == "NON_STANDARD")
    print(f"TOTAL — PASS={tot_pass}  FAIL={tot_fail}  "
          f"NOT_DEFINED={tot_nd}  WARNINGS={tot_warn}")

    suffix = f"_{args.network_id}" if args.network_id else ""

    if args.csv:
        dp = OUTPUT_DIR / f"audit_switch_{args.org_id}{suffix}_detail.csv"
        sp = OUTPUT_DIR / f"audit_switch_{args.org_id}{suffix}_summary.csv"
        write_csv(dp, all_detail,    _DETAIL_HEADERS)
        write_csv(sp, summary_rows,  _SUMMARY_HEADERS)
        print(f"\nCSV detail  → {dp}")
        print(f"CSV summary → {sp}")

    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_switch_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(
            all_detail, summary_rows, xp,
            _DETAIL_HEADERS, _SUMMARY_HEADERS,
            "SwitchAudit",
        )
        print(f"XLSX report → {xp}")


if __name__ == "__main__":
    main()
