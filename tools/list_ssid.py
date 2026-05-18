from __future__ import annotations

import argparse

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="List SSID in a Meraki Network")
    parser.add_argument("--network-id", required=True, help="Meraki Organization ID")
    parser.add_argument("--json", action="store_true", help="Write JSON output")
    parser.add_argument("--csv", action="store_true", help="Write CSV output")
    args = parser.parse_args()

    dashboard = get_dashboard()
    ssids = dashboard.wireless.getNetworkWirelessSsids(args.network_id)

    rows = []
    for ssid in ssids:
        rows.append(
            {
                "number":ssid.get("number"),
                "name": ssid.get("name"),
                "enabled": ssid.get("enabled"),
                "authMode": ssid.get("authMode"),
                "ipAssignmentMode": ssid.get("ipAssignmentMode"),
                "bandSelection": ssid.get("bandSelection"),
                "spashPage": ssid.get("spashPage"),
            }
        )

    for row in rows:
        print(
            f'{row["number"]}\t{row["name"]}\t'
            f'enabled={row["enabled"]}\tauth={row["authMode"]}'
        )

    if args.json:
        write_json(OUTPUT_DIR / f"ssids_{args.network_id}.json", ssids)

    if args.csv:
        write_csv(
            OUTPUT_DIR / f"ssids_{args.network_id}.csv",
            rows,
            ["number", "name", "enabled", "authMode", "ipAssignmentMode", "bandSelection", "spashPage"],
        )

if __name__ == "__main__":
    main()
