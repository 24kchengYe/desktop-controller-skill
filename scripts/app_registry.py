"""
App Registry — maps application names to their automation profiles.
Usage:
    python app_registry.py identify "WeChat"
    python app_registry.py list
"""

import json
import sys

APPS = {
    "weixin": {
        "name": "WeChat",
        "aliases": ["wechat", "weixin", "微信"],
        "process": "Weixin",
        "mode": "win32",
        "search_key": "^f",  # Ctrl+F in SendKeys format
        "search_key_display": "Ctrl+F",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.12},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 2000,
            "contact_select": 2500,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "wxwork": {
        "name": "WeCom (企业微信)",
        "aliases": ["wecom", "wxwork", "企业微信"],
        "process": "WXWork",
        "mode": "win32",
        "search_key": "^f",
        "search_key_display": "Ctrl+F",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.12},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 2000,
            "contact_select": 2500,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "dingtalk": {
        "name": "DingTalk (钉钉)",
        "aliases": ["dingtalk", "dingding", "钉钉"],
        "process": "DingTalk",
        "mode": "win32",
        "search_key": "^k",
        "search_key_display": "Ctrl+K",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.12},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 2000,
            "contact_select": 2500,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "feishu": {
        "name": "Feishu/Lark (飞书)",
        "aliases": ["feishu", "lark", "飞书"],
        "process": "Feishu",
        "mode": "win32",
        "search_key": "^k",
        "search_key_display": "Ctrl+K",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.12},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 2000,
            "contact_select": 2500,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "qq": {
        "name": "QQ",
        "aliases": ["qq", "腾讯qq"],
        "process": "QQ",
        "mode": "win32",
        "search_key": "^f",
        "search_key_display": "Ctrl+F",
        "input_area": {"x_ratio": 0.60, "y_ratio_from_bottom": 0.15},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 2000,
            "contact_select": 2500,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "telegram": {
        "name": "Telegram",
        "aliases": ["telegram", "tg"],
        "process": "Telegram",
        "mode": "win32",
        "search_key": "^k",
        "search_key_display": "Ctrl+K",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.08},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 1500,
            "contact_select": 2000,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "slack": {
        "name": "Slack",
        "aliases": ["slack"],
        "process": "slack",
        "mode": "win32",
        "search_key": "^k",
        "search_key_display": "Ctrl+K",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.08},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 1500,
            "contact_select": 2000,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
    "teams": {
        "name": "Microsoft Teams",
        "aliases": ["teams", "ms-teams", "微软teams"],
        "process": "ms-teams",
        "mode": "win32",
        "search_key": "^e",
        "search_key_display": "Ctrl+E",
        "input_area": {"x_ratio": 0.65, "y_ratio_from_bottom": 0.10},
        "timings": {
            "activate": 1000,
            "search_open": 600,
            "search_paste": 2000,
            "contact_select": 2500,
            "input_click": 1000,
            "message_paste": 800,
        },
    },
}


def identify(query: str) -> dict | None:
    """Find app profile by name or alias."""
    q = query.lower().strip()
    for key, app in APPS.items():
        if q == key or q in [a.lower() for a in app["aliases"]]:
            return {"key": key, **app}
    # Fuzzy match
    for key, app in APPS.items():
        for alias in app["aliases"]:
            if q in alias.lower() or alias.lower() in q:
                return {"key": key, **app}
    return None


def list_apps():
    """List all supported apps."""
    result = []
    for key, app in APPS.items():
        result.append({
            "key": key,
            "name": app["name"],
            "process": app["process"],
            "mode": app["mode"],
            "search_key": app["search_key_display"],
        })
    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python app_registry.py [identify|list] [query]")
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "list":
        for app in list_apps():
            print(f"  {app['key']:<12} {app['name']:<25} process={app['process']:<12} mode={app['mode']:<10} search={app['search_key']}")

    elif cmd == "identify":
        if len(sys.argv) < 3:
            print("Usage: python app_registry.py identify <app_name>")
            sys.exit(1)
        result = identify(sys.argv[2])
        if result:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(f"App not found: {sys.argv[2]}")
            print("Available apps:")
            for app in list_apps():
                print(f"  {app['key']}: {app['name']}")
            sys.exit(1)
