from __future__ import annotations

import argparse

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json

def main() -> None:
    parser = argparse.ArgumentParser(description="List SSID setting in a Meraki Network")
    parser.add_argument("--network-id", required=True, help="Meraki Organization ID")
    parser.add_argument("--json", action="store_true", help="Write JSON output")
    parser.add_argument("--csv", action="store_true", help="Write CSV output")
    args = parser.parse_args()

    dashboard = get_dashboard()
    ssids = dashboard.wireless.getNetworkWirelessSsids(args.network_id)

    rows = []
    for ssid in ssids:
        radius_servers = ssid.get("radiusServers", [])
        radius_hosts = ",".join([srv.get("host", "") for srv in radius_servers if isinstance(srv, dict)])

        rows.append(
            {
                "ssid_number": ssid.get("number"),
                "ssid_name": ssid.get("name"),
                "enabled": ssid.get("enabled"),
                "authMode": ssid.get("authMode"),
                "encryptionMode": ssid.get ("encryptionMode"),
                "wpaEncryptionMode": ssid.get ("wpaEncryptionMode"),
                "ipAssignmentMode": ssid.get ("ipAssignmentMode"),
                "bandSelection": ssid.get ("bandSelection"),
                "minBitrate": ssid.get("minBitrate"),
                "ssidAdminAccessible": ssid.get("ssidAdminAccessible"),
                "radiusEnabled": ssid.get("radiusEnabled"),
                "radiusAccountingEnabled": ssid.get ("radiusAccountingEnabled"),
                "radiusHosts": radius_hosts,
                "splashPage": ssid.get("splashPage"),
                "walledGardenEnabled": ssid.get("walledGardenEnabled"),
                "visible": ssid.get("visible"),
            }
        )

    for row in rows:
        print(
            f'{row["ssid_number"]}\t{row["ssid_name"]}\t'
            f'auth={row["authMode"]}\tip={row["ipAssignmentMode"]}\tradius{row["radiusEnabled"]}'
        )

    if args.json:
        write_json(OUTPUT_DIR / f"ssid_setting_{args.network_id}.json", ssids)

    if args.csv:
        write_csv(
            OUTPUT_DIR / f"ssid_setting_{args.network_id}.csv",
            rows,
            ["ssid_number","ssid_name","enabled","authMode","encryptionMode","wpaEncryptionMode","ipAssignmentMode","bandSelection","minBitrate","ssidAdminAccessible","radiusEnabled","radiusAccountingEnabled","radiusHosts","splashPage", "walledGardenEnabled","visible"],
        )

if __name__ == "__main__":
    main()
