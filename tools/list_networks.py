from __future__ import annotations

import argparse

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json

def main() -> None:
    parser = argparse.ArgumentParser(description="List Meraki Networks in Organizations")
    parser.add_argument("--org-id", required=True, help="Meraki Organization ID")
    parser.add_argument("--json", action="store_true", help="Write JSON output")
    parser.add_argument("--csv", action="store_true", help="Write CSV output")
    args = parser.parse_args()
    print("DebugArgs", args)

    dashboard = get_dashboard()
    networks = dashboard.organizations.getOrganizationNetworks(args.org_id, total_pages="all")

    rows = []
    for net in networks:
        rows.append(
            {
                "id": net.get("id"),
                "organizationId": net.get("organizationId"),
                "name": net.get("name"),
                "productTypes": ",".join(net.get("productTypes", [])),
                "timeZone": net.get("timeZone"),
                "tags": net.get("tags"),
                "isBoundToConfigTemplate": net.get("isBoundToConfigTemplate"),
            }
        )

    for row in rows:
        print(f'{row["id"]}\t{row["name"]}\t{row["productTypes"]}')

    if args.json:
        write_json(OUTPUT_DIR / f"networks_{args.org_id}.json",networks)

    if args.csv:
        write_csv(
            OUTPUT_DIR / f"networks_{args.org_id}.csv",
            rows,
            ["id", "organizationId", "name", "productTypes", "timeZone", "tags", "isBoundToConfigTemplate"],
        )

if __name__ == "__main__":
    main()
