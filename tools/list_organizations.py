from __future__ import annotations

import argparse
from pathlib import Path
from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="List Meraki Organizations")
    parser.add_argument("--json", action="store_true", help="Write JSON output")
    parser.add_argument("--csv", action="store_true", help="Write CSV output")
    args = parser.parse_args()

    dashboard = get_dashboard()
    orgs = dashboard.organizations.getOrganizations()

    rows = []

    for org in orgs:
        rows.append(
            {
                "id": org.get("id"),
                "name": org.get("name"),
                "api_enabled": org.get("api", {}).get("enabled") if isinstance(org.get("api"),dict) else None,
            }
        )

    for row in rows:
        print(f'{row["id"]}\t{row["name"]}')

    if args.json:
        write_json(OUTPUT_DIR / "organizations.json", orgs)

    if args.csv:
        write_csv(
            OUTPUT_DIR / "organizations.csv",
            rows,
            ["id","name", "api_enabled"],
        )

if __name__ == "__main__":
    main()
