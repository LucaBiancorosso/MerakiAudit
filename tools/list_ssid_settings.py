from __future__ import annotations

import argparse
import re

from config.settings import OUTPUT_DIR
from lib.meraki_client import get_dashboard
from lib.output import write_csv, write_json

_UNCONFIGURED_RE = re.compile(r'^unconfigured\s+ssid\s+\d+$', re.IGNORECASE)

def _is_unconfigured(name: str) -> bool:
    return bool(_UNCONFIGURED_RE.match((name or "").strip()))

_FIELDNAMES = [
    "ssid_number",
    "ssid_name",
    "enabled",
    "authMode",
    "encryptionMode",
    "wpaEncryptionMode",
    # ── Client IP / addressing ──────────────────────────────────────────────
    "ipAssignmentMode",
    "defaultVlanId",
    "useVlanTagging",
    "concentratorNetworkId",
    "lanIsolationEnabled",
    # ── Per-client bandwidth limits ─────────────────────────────────────────
    "perClientBandwidthLimitUp",
    "perClientBandwidthLimitDown",
    # ── Radio / band ────────────────────────────────────────────────────────
    "bandSelection",
    "minBitrate",
    # ── Auth / RADIUS ────────────────────────────────────────────────────────
    "ssidAdminAccessible",
    "radiusEnabled",
    "radiusAccountingEnabled",
    "radiusHosts",
    # ── Splash / visibility ──────────────────────────────────────────────────
    "splashPage",
    "walledGardenEnabled",
    "visible",
]

def main() -> None:
    parser = argparse.ArgumentParser(description="List SSID settings in a Meraki Network")
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

        radius_servers = ssid.get("radiusServers") or []
        radius_hosts = ",".join(
            srv.get("host", "")
            for srv in radius_servers
            if isinstance(srv, dict)
        )

        rows.append(
            {
                "ssid_number":                  ssid.get("number"),
                "ssid_name":                    ssid.get("name"),
                "enabled":                      ssid.get("enabled"),
                "authMode":                     ssid.get("authMode"),
                "encryptionMode":               ssid.get("encryptionMode"),
                "wpaEncryptionMode":            ssid.get("wpaEncryptionMode"),
                # ── Client IP / addressing ──────────────────────────────
                "ipAssignmentMode":             ssid.get("ipAssignmentMode"),
                "defaultVlanId":                ssid.get("defaultVlanId"),
                "useVlanTagging":               ssid.get("useVlanTagging"),
                "concentratorNetworkId":        ssid.get("concentratorNetworkId") or "",
                "lanIsolationEnabled":          ssid.get("lanIsolationEnabled"),
                # ── Per-client bandwidth limits ─────────────────────────
                "perClientBandwidthLimitUp":    ssid.get("perClientBandwidthLimitUp", 0),
                "perClientBandwidthLimitDown":  ssid.get("perClientBandwidthLimitDown", 0),
                # ── Radio / band ────────────────────────────────────────
                "bandSelection":                ssid.get("bandSelection"),
                "minBitrate":                   ssid.get("minBitrate"),
                # ── Auth / RADIUS ────────────────────────────────────────
                "ssidAdminAccessible":          ssid.get("ssidAdminAccessible"),
                "radiusEnabled":                ssid.get("radiusEnabled"),
                "radiusAccountingEnabled":      ssid.get("radiusAccountingEnabled"),
                "radiusHosts":                  radius_hosts,
                # ── Splash / visibility ──────────────────────────────────
                "splashPage":                   ssid.get("splashPage"),
                "walledGardenEnabled":          ssid.get("walledGardenEnabled"),
                "visible":                      ssid.get("visible"),
            }
        )

    for row in rows:
        vlan_info = (
            f"vlan={row['defaultVlanId']}" if row["useVlanTagging"]
            else f"ip={row['ipAssignmentMode']}"
        )
        print(
            f'{row["ssid_number"]}\t{row["ssid_name"]}\t'
            f'auth={row["authMode"]}\t{vlan_info}\t'
            f'radius={row["radiusEnabled"]}'
        )

    if args.json:
        write_json(OUTPUT_DIR / f"ssid_setting_{args.network_id}.json", ssids)

    if args.csv:
        write_csv(
            OUTPUT_DIR / f"ssid_setting_{args.network_id}.csv",
            rows,
            _FIELDNAMES,
        )

if __name__ == "__main__":
    main()
