"""
audit_ap.py
Audit AP management configuration and connectivity mode against standards
defined in the APConfig sheet of the standards Excel file.

Data sources per AP:
  - getOrganizationDevices           → serial, name, model, lanIp, firmware, tags, networkId
  - getNetworkWirelessAlternateManagementInterface → per-AP static IP config
  - getNetworkWirelessMeshStatuses   → which APs are mesh repeaters (vs gateway)

Checks performed:
  mgmt_ip_mode         — "static" or "dhcp"
  mgmt_ip_in_subnet    — AP static/lan IP falls within allowed_mgmt_subnets (CIDR list)
  mgmt_vlan            — management VLAN ID
  mgmt_dns1            — primary DNS
  mgmt_dns2            — secondary DNS
  connection_mode      — "gateway" or "mesh"
  tags_subset          — expected_tags are a subset of the AP's actual tags

Usage:
    python -m tools.audit_ap --org-id <ORG_ID> \
        --standards-file standards/ssid_standards_template.xlsx --xlsx
    python -m tools.audit_ap --org-id <ORG_ID> \
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
    ip_in_allowed_subnets,
    load_ap_standards,
    load_ignored_fields,
    resolve_ap_standard,
)
from lib.audit_common import (
    build_network_lookup,
    build_summary,
    make_row,
    write_xlsx_report,
)

# ---------------------------------------------------------------------------
# Fields and headers
# ---------------------------------------------------------------------------

_AUDITED_FIELDS = [
    "mgmt_ip_mode",        # "static" or "dhcp"
    "mgmt_ip_in_subnet",   # PASS/FAIL against allowed_mgmt_subnets
    "mgmt_vlan",           # management VLAN
    "mgmt_dns1",           # primary DNS
    "mgmt_dns2",           # secondary DNS
    "connection_mode",     # "gateway" or "mesh"
    "tags_subset",         # expected tags present on AP
]

_DETAIL_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key",    # serial
    "entity_label",  # AP name
    "model", "firmware", "lanIp",
    "standard_level", "field", "expected", "actual", "result",
]

_SUMMARY_HEADERS = [
    "org_id", "org_name", "network_id", "network_name",
    "entity_key", "entity_label",
    "tot_checks", "tot_pass", "tot_fail", "tot_not_defined",
    "tot_ignored", "non_standard", "missing", "overall",
]


# ---------------------------------------------------------------------------
# Per-network data collectors
# ---------------------------------------------------------------------------

def _fetch_alternate_mgmt(dashboard, network_id: str) -> dict:
    """
    Return alternateManagementInterface dict, or {} on error.
    Structure: {enabled, vlanId, protocols, accessPoints: [{serial, alternateManagementIp, ...}]}
    """
    try:
        return dashboard.wireless.getNetworkWirelessAlternateManagementInterface(network_id) or {}
    except Exception:
        return {}


def _fetch_mesh_serials(dashboard, network_id: str) -> set[str]:
    """Return set of serials that are mesh repeaters."""
    try:
        statuses = dashboard.wireless.getNetworkWirelessMeshStatuses(
            network_id, total_pages="all"
        )
        return {s["serial"] for s in statuses if s.get("serial")}
    except Exception:
        return set()


def _build_static_ip_map(alt_mgmt: dict) -> dict[str, dict]:
    """Return {serial: {ip, subnetMask, gateway, dns1, dns2, vlanId}} from alternateManagementInterface."""
    result = {}
    for ap in alt_mgmt.get("accessPoints") or []:
        serial = ap.get("serial", "")
        if serial:
            result[serial] = {
                "ip":         ap.get("alternateManagementIp", ""),
                "subnetMask": ap.get("subnetMask", ""),
                "gateway":    ap.get("gateway", ""),
                "dns1":       ap.get("dns1", ""),
                "dns2":       ap.get("dns2", ""),
                "vlanId":     str(alt_mgmt.get("vlanId", "")),
            }
    return result


# ---------------------------------------------------------------------------
# Audit one AP
# ---------------------------------------------------------------------------

def _audit_ap(
    org_id: str, org_name: str,
    network_id: str, network_name: str,
    ap: dict,
    static_map: dict[str, dict],
    mesh_serials: set[str],
    standard: dict | None,
    std_level: str,
    ignored_fields: set[str],
) -> list[dict]:
    serial   = ap.get("serial", "")
    ap_name  = ap.get("name", serial)
    model    = ap.get("model", "")
    firmware = ap.get("firmware", "")
    lan_ip   = ap.get("lanIp", "")
    tags     = ap.get("tags") or []
    tags_str = ",".join(sorted(t.strip() for t in tags if t.strip()))

    static   = static_map.get(serial)
    is_static = static is not None

    # Derive live values for each checked field
    mgmt_ip          = static["ip"]   if is_static else lan_ip
    mgmt_dns1        = static["dns1"] if is_static else ""
    mgmt_dns2        = static["dns2"] if is_static else ""
    mgmt_vlan        = static["vlanId"] if is_static else ""
    live_ip_mode     = "static" if is_static else "dhcp"
    live_conn_mode   = "mesh" if serial in mesh_serials else "gateway"

    extra = {"model": model, "firmware": firmware, "lanIp": lan_ip}
    rows: list[dict] = []

    def _row(field, expected, actual, result):
        return make_row(
            org_id, org_name, network_id, network_name,
            entity_key=serial, entity_label=ap_name,
            field=field, expected=expected, actual=actual,
            result=result, standard_level=std_level, extra=extra,
        )

    if standard is None:
        # AP's network has no standard → report actual values as NON_STANDARD
        for field in _AUDITED_FIELDS:
            actuals = {
                "mgmt_ip_mode":     live_ip_mode,
                "mgmt_ip_in_subnet": mgmt_ip,
                "mgmt_vlan":         mgmt_vlan,
                "mgmt_dns1":         mgmt_dns1,
                "mgmt_dns2":         mgmt_dns2,
                "connection_mode":   live_conn_mode,
                "tags_subset":       tags_str,
            }
            rows.append(_row(field, "", actuals.get(field, ""), "NON_STANDARD"))
        return rows

    for field in _AUDITED_FIELDS:
        if field in ignored_fields:
            rows.append(_row(field, "", "", "IGNORED"))
            continue

        expected = standard.get(field, NOT_DEFINED)

        if expected == NOT_DEFINED:
            actuals = {
                "mgmt_ip_mode":     live_ip_mode,
                "mgmt_ip_in_subnet": mgmt_ip,
                "mgmt_vlan":         mgmt_vlan,
                "mgmt_dns1":         mgmt_dns1,
                "mgmt_dns2":         mgmt_dns2,
                "connection_mode":   live_conn_mode,
                "tags_subset":       tags_str,
            }
            rows.append(_row(field, "", actuals.get(field, ""), NOT_DEFINED))
            continue

        # ── Special comparisons ──────────────────────────────────────────────
        if field == "mgmt_ip_mode":
            result = "PASS" if live_ip_mode.lower() == expected.lower() else "FAIL"
            rows.append(_row(field, expected, live_ip_mode, result))

        elif field == "mgmt_ip_in_subnet":
            # expected is a comma-sep list of CIDR subnets
            in_subnet = ip_in_allowed_subnets(mgmt_ip, expected)
            rows.append(_row(field, expected, mgmt_ip,
                             "PASS" if in_subnet else "FAIL"))

        elif field == "mgmt_vlan":
            result = "PASS" if str(mgmt_vlan) == str(expected) else "FAIL"
            rows.append(_row(field, expected, mgmt_vlan, result))

        elif field == "mgmt_dns1":
            result = "PASS" if mgmt_dns1.strip().lower() == expected.strip().lower() else "FAIL"
            rows.append(_row(field, expected, mgmt_dns1, result))

        elif field == "mgmt_dns2":
            result = "PASS" if mgmt_dns2.strip().lower() == expected.strip().lower() else "FAIL"
            rows.append(_row(field, expected, mgmt_dns2, result))

        elif field == "connection_mode":
            result = "PASS" if live_conn_mode.lower() == expected.lower() else "FAIL"
            rows.append(_row(field, expected, live_conn_mode, result))

        elif field == "tags_subset":
            # expected is a comma-sep list; all expected tags must be present on AP
            expected_tags = {t.strip().lower() for t in expected.split(",") if t.strip()}
            actual_tags   = {t.strip().lower() for t in tags if t.strip()}
            missing_tags  = expected_tags - actual_tags
            result = "PASS" if not missing_tags else "FAIL"
            rows.append(_row(field, expected, tags_str, result))

    return rows


# ---------------------------------------------------------------------------
# Console output
# ---------------------------------------------------------------------------

def _print_network_results(network_name: str, network_id: str, detail_rows: list[dict]) -> None:
    pass_c = sum(1 for r in detail_rows if r["result"] == "PASS")
    fail_c = sum(1 for r in detail_rows if r["result"] == "FAIL")
    nd_c   = sum(1 for r in detail_rows if r["result"] in (NOT_DEFINED, "NOT_DEFINED"))
    ign_c  = sum(1 for r in detail_rows if r["result"] == "IGNORED")
    warn_ap = {r["entity_key"] for r in detail_rows if r["result"] == "NON_STANDARD"}

    print(f"  [{network_name}]  {network_id}")
    print(f"    PASS={pass_c}  FAIL={fail_c}  NOT_DEFINED={nd_c}  "
          f"IGNORED={ign_c}  WARNINGS={len(warn_ap)}")

    aps_seen: list[str] = []
    for r in detail_rows:
        if r["entity_key"] not in aps_seen:
            aps_seen.append(r["entity_key"])

    for serial in aps_seen:
        rows = [r for r in detail_rows if r["entity_key"] == serial]
        first = rows[0]
        print(f"\n    AP {first['entity_label']!r}  serial={serial}  "
              f"model={first.get('model','')}  ip={first.get('lanIp','')}")
        for r in rows:
            result = r["result"]
            icon   = {"PASS": "✓", "FAIL": "✗"}.get(result, "–")
            if result == "NON_STANDARD":
                print(f"      ⚠ NON_STANDARD  {r['field']:<22} actual={r['actual']!r}")
            elif result == "FAIL":
                print(f"      {icon} FAIL          {r['field']:<22} "
                      f"expected={r['expected']!r}  actual={r['actual']!r}")
            elif result in (NOT_DEFINED, "NOT_DEFINED"):
                print(f"      – NOT_DEFINED  {r['field']:<22} (not audited)")
            elif result == "IGNORED":
                print(f"      – IGNORED      {r['field']:<22}")
            else:
                print(f"      {icon} PASS          {r['field']:<22} value={r['actual']!r}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit AP management and connectivity configuration"
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

    print(f"Loading AP standards from {standards_path.name}...")
    standards      = load_ap_standards(standards_path)
    ignored_fields = load_ignored_fields(standards_path)

    org_key = args.org_id.strip().lower()
    org_standards = {k: v for k, v in standards.items() if k[0] == org_key}
    if not org_standards:
        print(f"WARNING: no AP standards for org '{args.org_id}'. "
              f"Add rows to the 'APConfig' sheet.")
    else:
        print(f"  {len(org_standards)} AP standard(s) found.")
        for (o, n) in org_standards:
            level = f"network={n}" if n else "org-wide"
            print(f"    [{level}]")
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

    # Fetch all wireless APs for the org at once (more efficient than per-network)
    print(f"Fetching wireless AP inventory for org '{org_name or args.org_id}'...")
    try:
        all_aps = dashboard.organizations.getOrganizationDevices(
            args.org_id, total_pages="all", productTypes=["wireless"]
        )
    except Exception as e:
        print(f"ERROR fetching devices: {e}")
        sys.exit(1)

    # Group APs by network
    aps_by_network: dict[str, list[dict]] = {}
    for ap in all_aps:
        nid = str(ap.get("networkId", "")).strip()
        if nid:
            aps_by_network.setdefault(nid, []).append(ap)

    print(f"Auditing {len(networks_to_audit)} network(s)...")
    print()

    all_detail: list[dict] = []

    for network_id, network_name in sorted(networks_to_audit.items(), key=lambda x: x[1]):
        aps = aps_by_network.get(network_id, [])
        if not aps:
            print(f"  [{network_name}] no wireless APs, skipping.")
            continue

        alt_mgmt     = _fetch_alternate_mgmt(dashboard, network_id)
        static_map   = _build_static_ip_map(alt_mgmt)
        mesh_serials = _fetch_mesh_serials(dashboard, network_id)

        standard  = resolve_ap_standard(standards, args.org_id, network_id)
        oid_l, nid_l = args.org_id.strip().lower(), network_id.strip().lower()
        if standards.get((oid_l, nid_l)):
            std_level = "network"
        elif standards.get((oid_l, "")):
            std_level = "org"
        else:
            std_level = ""

        net_rows: list[dict] = []
        for ap in aps:
            net_rows.extend(_audit_ap(
                org_id=args.org_id, org_name=org_name,
                network_id=network_id, network_name=network_name,
                ap=ap, static_map=static_map, mesh_serials=mesh_serials,
                standard=standard, std_level=std_level,
                ignored_fields=ignored_fields,
            ))

        _print_network_results(network_name, network_id, net_rows)
        all_detail.extend(net_rows)

    summary_rows = build_summary(all_detail)

    print()
    tot_pass = sum(r["tot_pass"] for r in summary_rows)
    tot_fail = sum(r["tot_fail"] for r in summary_rows)
    tot_nd   = sum(r["tot_not_defined"] for r in summary_rows)
    tot_warn = sum(1 for r in summary_rows if r["overall"] in ("NON_STANDARD",))
    print(f"TOTAL — PASS={tot_pass}  FAIL={tot_fail}  NOT_DEFINED={tot_nd}  WARNINGS={tot_warn}")

    suffix = f"_{args.network_id}" if args.network_id else ""

    if args.csv:
        dp = OUTPUT_DIR / f"audit_ap_{args.org_id}{suffix}_detail.csv"
        sp = OUTPUT_DIR / f"audit_ap_{args.org_id}{suffix}_summary.csv"
        write_csv(dp, all_detail,   _DETAIL_HEADERS)
        write_csv(sp, summary_rows, _SUMMARY_HEADERS)
        print(f"\nCSV detail  → {dp}")
        print(f"CSV summary → {sp}")

    if args.xlsx:
        xp = OUTPUT_DIR / f"audit_ap_{args.org_id}{suffix}.xlsx"
        write_xlsx_report(all_detail, summary_rows, xp,
                          _DETAIL_HEADERS, _SUMMARY_HEADERS, "APAudit")
        print(f"XLSX report → {xp}")


if __name__ == "__main__":
    main()
