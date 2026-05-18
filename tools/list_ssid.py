from __future__ import annotations

import argparse
import re

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json

_UNCONFIGURED_RE = re.compile(r'^unconfigured\s+ssid\s+\d+$', re.IGNORECASE)

def _is_unconfigured(name: str) -> bool:
    return bool(_UNCONFIGURED_RE.match((name or "").strip()))

def main() -> None:
    parser = argparse.ArgumentParser(description="List SSIDs in a Meraki Network")
    parser.add_argument("--network-id", required=True, help="Meraki Network ID")
    parser.add_argument("--json", action="store_true", help="Write JSON output")
    parser.add_argument("--csv", action="store_true", help="Write CSV output")
    parser.add_argument("--include-unconfigured", action="store_true",
                        help="Include 'Unconfigured SSID N' slots (hidden by default)")
    args = parser.parse_args()

    dashboard = get_dashboard()
    ssids = dashboard.wireless.getNetworkWirelessSsids(args.network_id)

    rows = []
    for ssid in ssids:
        if not args.include_unconfigured and _is_unconfigured(ssid.get("name", "")):
            continue
        rows.append(
            {
                "number":           ssid.get("number"),
                "name":             ssid.get("name"),
                "enabled":          ssid.get("enabled"),
                "authMode":         ssid.get("authMode"),
                "ipAssignmentMode": ssid.get("ipAssignmentMode"),
                "defaultVlanId":    ssid.get("defaultVlanId"),
                "useVlanTagging":   ssid.get("useVlanTagging"),
                "bandSelection":    ssid.get("bandSelection"),
                "splashPage":       ssid.get("splashPage"),
            }
        )

    for row in rows:
        vlan_info = (
            f"vlan={row['defaultVlanId']}" if row["useVlanTagging"]
            else f"ip={row['ipAssignmentMode']}"
        )
        print(
            f'{row["number"]}\t{row["name"]}\t'
            f'enabled={row["enabled"]}\tauth={row["authMode"]}\t{vlan_info}'
        )

    if args.json:
        write_json(OUTPUT_DIR / f"ssids_{args.network_id}.json", ssids)

    if args.csv:
        write_csv(
            OUTPUT_DIR / f"ssids_{args.network_id}.csv",
            rows,
            ["number", "name", "enabled", "authMode", "ipAssignmentMode",
             "defaultVlanId", "useVlanTagging", "bandSelection", "splashPage"],
        )

if __name__ == "__main__":
    main()
