from __future__ import annotations

import meraki
from config.settings import (
    MERAKI_DASHBOARD_API_KEY,
    MERAKI_LOG_PATH,
    MERAKI_PRINT_CONSOLE,
)

def get_dashboard() ->meraki.DashboardAPI:
    if not MERAKI_DASHBOARD_API_KEY:
        raise RuntimeError ("Missing Meraki APi key in environment variables")

    return meraki.DashboardAPI(
        api_key=MERAKI_DASHBOARD_API_KEY,
        suppress_logging=not MERAKI_PRINT_CONSOLE,
        output_log=bool(MERAKI_LOG_PATH),
        log_path=MERAKI_LOG_PATH or None,
        print_console=MERAKI_PRINT_CONSOLE,
    )
