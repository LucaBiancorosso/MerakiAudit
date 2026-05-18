from __future__ import annotations

import argparse

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json

def build_network_lookup(dashboard, org_id: str) -> dict[str, str]:
    networks = dashboard.organizations.getOrganizationNetworks(org_id, total_pages="all")
    lookup = {}

    for net in networks:
        if not isinstance(net, dict):
            print(f"DEBUG Unexpected Type -> {type(net)} : {net}")
            continue
        net_id = net.get("id")
        net_name = net.get("name")
        if net_id:
            lookup[str(net_id).strip()] = net_name

    return lookup

def normalize_device_family(product_type: str | None, model: str | None ) -> tuple[str, str]:
    if not model:
        if product_type == "switch":
            return "Unknown", "Switch"
        if product_type == "wireless":
            return "Unknown", "Access Point"
        if product_type == "wirelessController":
            return "Unknown", "Wireless Controller"
        if product_type == "appliance":
            return "Unknown", "Security Appliance"
        return "Unknown", "Unknown"

    m = model.upper()

    if m.startswith("MS"):
        return "Meraki", "Meraki Switch"
    if m.startswith("MR"):
        return "Meraki", "Meraki Access Point"
    if m.startswith("MX"):
        return "Meraki", "Meraki Security Appliance"
    if m.startswith("MV"):
        return "Meraki", "Meraki Camera"
    if m.startswith("MG"):
        return "Meraki", "Meraki Cellular Gateway"
    if m.startswith("MT"):
        return "Meraki", "Meraki Sensor"
    if m.startswith("VMX"):
        return "Meraki", "Meraki Virutal Appliance"

    if m.startswith("CW"):
        return "Catalyst", "Catalyst Access Point"
    if m.startswith("C9800"):
        return "Catalyst", "Catalyst Wireless Controller"
    if (m.startswith("C9200")
        or m.startswith("C9300")
        or m.startswith("C9400")
        or m.startswith("C9500")
        or m.startswith("C9600")):
        return "Catalyst", "Catalyst Switch"

    if product_type == "switch":
        return "Unknown", "Switch"
    if product_type == "wireless":
        return "Unknown", "Access Point"
    if product_type == "wirelessController":
        return "Unknown", "Wireless Controller"
    if product_type == "appliance":
        return "Unknown", "Security Appliance"
    return "Unknown", "Unknown"

def date_only(value: str | None) -> str:
    if not value:
        return ""
    return value.split("T")[0]

def main() -> None:
    parser = argparse.ArgumentParser(description="List Meraki Inventory in Organizations with EoX")
    parser.add_argument("--org-id", required=True, help="Meraki Organization ID")
    parser.add_argument("--json", action="store_true", help="Write JSON output")
    parser.add_argument("--csv", action="store_true", help="Write CSV output")
    parser.add_argument("--eox-status", choices=["endOfSale","endOfSupport","nearEndOfSupport", "null"], help="Optional filter For EOX")
    args = parser.parse_args()

    dashboard = get_dashboard()

    networ_lookup = build_network_lookup(dashboard, args.org_id)

    inventory_kwargs = {"total_pages": "all"}
    if args.eox_status:
        inventory_kwargs["eoxStatus"] = args.eox_status

    devices = dashboard.organizations.getOrganizationInventoryDevices(args.org_id,**inventory_kwargs)

    rows = []
    for d in devices:
        eox = d.get("eox") or {}
        network_id = d.get("networkId")
        network_id = str(network_id).strip() if network_id else ""
        network_name = networ_lookup.get(network_id, "") if network_id else ""
        product_type = d.get("productType")
        model = d.get("model")
        platform, family = normalize_device_family(product_type, model)

        rows.append(
            {
                "serial": d.get("serial"),
                "name": d.get("name"),
                "model": model,
                "product_type": product_type,
                "devicePlatform": platform,
                "deviceFamily": family,
                "networkId": network_id,
                "networkName": network_name,
                "isAssociatedToNetwork": bool(network_id),
                "eoxStatus": eox.get("status"),
                "endOfSaleAt": date_only(eox.get("endOfSaleAt")),
                "entOfSupportAt": date_only(eox.get("endOfSupportAt")),
            }
        )

    for r in rows:
        print(
            f'{r["serial"]}\t'
            f'{r["deviceFamily"]}\t'
            f'{r["networkName"] or "UNASSIGNED"}\t'
            f'{r["eoxStatus"] or "NO_EOX"}\t'
            f'{r["networkId"]}\t'
        )

    suffix = f"_{args.eox_status}" if args.eox_status else ""

    if args.json:
        write_json(OUTPUT_DIR / f"inventory_{args.org_id}{suffix}.json",devices)

    if args.csv:
        write_csv(
            OUTPUT_DIR / f"inventory_{args.org_id}{suffix}.csv",
            rows,
            [
            "serial",
            "name",
            "model",
            "product_type",
            "devicePlatform",
            "deviceFamily",
            "networkId",
            "networkName",
            "isAssociatedToNetwork",
            "eoxStatus",
            "endOfSaleAt",
            "entOfSupportAt",
            ]
        )

if __name__ == "__main__":
    main()
