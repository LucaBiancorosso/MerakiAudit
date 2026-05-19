# Audit Roadmap

Planned audit coverage for a full Meraki architecture.
Each item maps to a new tool in `tools/` and one or more new sheets in `standards/standard_audit_fields.xlsx`.

Status legend: ✅ Done · 🔜 Next · 📋 Planned · 💡 Idea

---

## ✅ Completed

| Tool | Sheet(s) | What it audits |
|---|---|---|
| `audit_ssid` | `Standards` | SSID auth, encryption, VLAN, RADIUS, bandwidth, splash |
| `audit_rf_profile` | `RFProfiles` | RF profile per band — power, channels, bitrate, axEnabled, RXSOP |
| `audit_ap` | `APConfig` | AP management IP mode, subnet, VLAN, DNS, mesh vs gateway, tags |
| `audit_switch_interfaces` | `SwitchInterfaces`, `VlanRoles` | Per-port role (access/trunk_ap/trunk_cpe) with STP, DAI, UDLD, storm control |

---

## 🔜 Phase 2 — LAN / Switching (continue)

### `audit_switch_stp` — Switch STP global settings
- **API:** `getNetworkSwitchStp`
- **Sheet:** `SwitchSTP`
- **Fields:** `rstpEnabled`, per-switch `stpPriority`, RSTP bridge priority hierarchy

### `audit_switch_acl` — Switch ACLs
- **API:** `getNetworkSwitchAccessControlLists`
- **Sheet:** `SwitchACL`
- **Approach:** compare rule count, check that specific deny/permit rules exist by policy (allow/deny, protocol, srcCidr, dstCidr, dstPort)
- **Note:** order-sensitive — rules are evaluated top to bottom

### `audit_switch_qos` — QoS rules
- **API:** `getNetworkSwitchQosRules`
- **Sheet:** `SwitchQoS`
- **Fields:** per-rule `vlan`, `protocol`, `srcPort`, `dscp`

### `audit_switch_mtu` — MTU and jumbo frames
- **API:** `getNetworkSwitchMtu`
- **Sheet:** `SwitchMTU`
- **Fields:** `defaultMtuSize`, per-switch overrides

### `audit_switch_dhcp_snooping` — DHCP snooping
- **API:** `getNetworkSwitchDhcpServerPolicy`
- **Sheet:** `SwitchDHCPSnooping`
- **Fields:** `defaultPolicy` (allow/block), `allowedServers`, `blockedServers`

---

## 📋 Phase 3 — Security Appliance (MX)

### `audit_mx_firewall` — L3 and L7 firewall rules
- **API:** `getNetworkApplianceFirewallL3FirewallRules`, `getNetworkApplianceFirewallL7FirewallRules`
- **Sheet:** `MXFirewall`
- **Approach:** check that mandatory deny/permit rules exist; validate default outbound policy; flag rules with `destCidr=any` and `policy=allow`

### `audit_mx_vpn` — Site-to-site VPN
- **API:** `getNetworkApplianceVpnSiteToSiteVpn`
- **Sheet:** `MXVPN`
- **Fields:** `mode` (hub/spoke/none), `hubs` list, `subnets` advertised, `psk` presence (not value)

### `audit_mx_content_filtering` — Content filtering
- **API:** `getNetworkApplianceContentFiltering`
- **Sheet:** `MXContentFiltering`
- **Fields:** `allowedUrlPatterns`, `blockedUrlPatterns`, `blockedUrlCategories`, `urlCategoryListSize`

### `audit_mx_threat` — Threat protection / IDS-IPS
- **API:** `getNetworkApplianceSecurity`  (Threat Protection endpoint)
- **Sheet:** `MXThreat`
- **Fields:** `mode` (detection/prevention), `rulesetType` (connectivity/balanced/security)

### `audit_mx_wan` — WAN uplink configuration
- **API:** `getNetworkApplianceUplinksUsageHistory`, `getDeviceApplianceUplinksSettings`
- **Sheet:** `MXWan`
- **Fields:** per-uplink `enabled`, `vlan`, `pppoeEnabled`, `svisEnabled`, DNS, static vs DHCP

### `audit_mx_vlans` — MX VLAN configuration
- **API:** `getNetworkApplianceVlans`
- **Sheet:** `MXVlans`
- **Fields:** per-VLAN `id`, `name`, `subnet`, `applianceIp`, `dhcpHandling`, `dhcpLeaseTime`, `dnsNameservers`, `reservedIpRanges`

---

## 📋 Phase 4 — Wireless (continue)

### `audit_wireless_settings` — Network-level wireless settings
- **API:** `getNetworkWirelessSettings`
- **Sheet:** `WirelessSettings`
- **Fields:** `meshingEnabled`, `ipv6BridgeEnabled`, `locationAnalyticsEnabled`, `upgradeStrategy`, `ledLightsOn`

### `audit_bluetooth` — Bluetooth / BLE settings
- **API:** `getNetworkWirelessBluetoothSettings`
- **Sheet:** `Bluetooth`
- **Fields:** `scanningEnabled`, `advertisingEnabled`, `uuid`, `major`, `minor`

### `audit_splash` — Splash page configuration
- **API:** `getNetworkWirelessSsidSplashSettings` (per SSID)
- **Sheet:** `SplashSettings`
- **Fields:** `splashUrl`, `useSplashUrl`, `splashTimeout`, `redirectUrl`, `welcomeMessage`, `allowAllExemptions`

---

## 📋 Phase 5 — Network-wide / Organisation

### `audit_admin` — Dashboard admin accounts
- **API:** `getOrganizationAdmins`
- **Sheet:** `Admins`
- **Approach:** flag admins with `orgAccess=full` that are not in an approved list; check MFA enforcement; check for stale accounts (lastActive > N days)

### `audit_syslog` — Syslog configuration
- **API:** `getNetworkSyslogServers`
- **Sheet:** `Syslog`
- **Fields:** check that required syslog server IPs and roles (URLs, events types) are configured per network

### `audit_snmp` — SNMP settings
- **API:** `getOrganizationSnmp`
- **Sheet:** `SNMP`
- **Fields:** `v2cEnabled`, `v3Enabled`, `v3AuthMode`, `v3PrivMode` — flag v2c if org policy requires v3-only

### `audit_alerts` — Alert profiles
- **API:** `getNetworkAlertsSettings`
- **Sheet:** `Alerts`
- **Approach:** check that mandatory alert types (rogue AP, uplink down, VPN connectivity) are enabled with the correct destinations

### `audit_firmware` — Firmware compliance
- **API:** `getOrganizationFirmwareUpgrades`, `getOrganizationDevices`
- **Sheet:** `Firmware`
- **Approach:** compare running firmware against a minimum acceptable version per product family; flag devices below the threshold

### `audit_network_tags` — Network tagging consistency
- **API:** `getOrganizationNetworks`
- **Sheet:** `NetworkTags`
- **Approach:** check that required tags are present on each network (e.g. `prod`, `site-type:branch`); useful for policy grouping consistency

---

## 💡 Phase 6 — Operational / Compliance reporting

### `report_eol_summary` — End-of-Life executive summary
- Extend `forecast_eol` with a cover sheet showing counts by platform and urgency band (already EoS / EoL within 12 months / 12-24 months / beyond)

### `report_compliance_score` — Cross-tool compliance score
- Aggregate results from all audit XLSX outputs into a single workbook
- One row per network with pass/fail/warn counts per audit domain
- RAG (Red/Amber/Green) overall score per network and org

### `report_inventory_delta` — Inventory change detection
- Compare two inventory CSV snapshots (current vs previous run)
- Highlight new devices, removed devices, moved devices (network change), firmware changes

---

## Notes on `standard_audit_fields.xlsx` growth

As new audit tools are added, the standards file gains new sheets. The current sheet list:

| Sheet | Status |
|---|---|
| `Standards` | ✅ |
| `RFProfiles` | ✅ |
| `APConfig` | ✅ |
| `SwitchInterfaces` | ✅ |
| `VlanRoles` | ✅ |
| `Ignore` | ✅ |
| `SwitchSTP` | 🔜 |
| `SwitchACL` | 🔜 |
| `MXFirewall` | 📋 |
| `MXVlans` | 📋 |
| `MXVPN` | 📋 |
| `WirelessSettings` | 📋 |
| `Admins` | 📋 |
| `Firmware` | 📋 |
| `Alerts` | 📋 |

The `Ignore` sheet remains global — fields listed there are skipped across **all** audit tools.
