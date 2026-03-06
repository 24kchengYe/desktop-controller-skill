# Universal Desktop Controller

A Claude Code skill for automating **any** Windows desktop application — send messages, click buttons, type text, take screenshots, and interact with native apps using Win32 API and Playwright.

Inspired by [OpenAI's playwright-interactive](https://github.com/openai/skills/tree/main/skills/.curated/playwright-interactive) Codex skill, adapted for Claude Code with native Windows desktop support.

## Architecture

```
┌──────────────────────────────────────────────────┐
│             Desktop Controller Skill              │
├──────────────┬──────────────┬────────────────────┤
│  Win32 Layer │  Playwright  │  Visual Feedback   │
│  (Native)    │  (Web/Elec.) │  (Screenshot+AI)   │
├──────────────┼──────────────┼────────────────────┤
│ FindWindow   │ Browser      │ CaptureScreen      │
│ SendKeys     │ Page.click   │ CaptureWindow      │
│ SetCursorPos │ Page.fill    │ → Claude Vision    │
│ mouse_event  │ Page.goto    │ → Verify State     │
│ Clipboard    │ Locator      │ → Decide Next Step │
└──────────────┴──────────────┴────────────────────┘
```

## Supported Apps

| App | Process | Mode | Search Key | Status |
|-----|---------|------|------------|--------|
| **WeChat** | Weixin | Win32 | Ctrl+F | Tested |
| **WeCom** | WXWork | Win32 | Ctrl+F | Ready |
| **DingTalk** | DingTalk | Win32 | Ctrl+K | Ready |
| **Feishu/Lark** | Feishu | Win32 | Ctrl+K | Ready |
| **QQ** | QQ | Win32 | Ctrl+F | Ready |
| **Telegram** | Telegram | Win32 | Ctrl+K | Ready |
| **Slack** | slack | Win32 | Ctrl+K | Ready |
| **Teams** | ms-teams | Win32 | Ctrl+E | Ready |

## Quick Start

### Via Claude Code (Recommended)

Just tell Claude Code what you want in natural language:

```
"Send a WeChat message to 张三 saying 你好"
"给张三发钉钉消息"
"Take a screenshot of WeChat"
"Click at position 500,400 in DingTalk"
"帮我操控电脑发消息"
```

### Via Command Line

```bash
# Send a message
python scripts/desktop_control.py send-message --app weixin --contact "张三" --message "你好"

# Take a screenshot of an app window
python scripts/desktop_control.py screenshot --app weixin --output wechat.png

# Take a full screen screenshot
python scripts/desktop_control.py screenshot --output fullscreen.png

# Click at specific coordinates
python scripts/desktop_control.py click --app weixin --x 500 --y 400

# Type text into an app
python scripts/desktop_control.py type --app weixin --text "Hello World"

# Find an app window
python scripts/desktop_control.py find-window --app weixin

# List all supported apps
python scripts/desktop_control.py list-apps
```

## Key Technical Insights

### 1. Mouse Click is Critical for Input Focus

After searching and selecting a contact in chat apps, the message input area does **NOT** automatically receive keyboard focus. You **must** use Win32 mouse automation (`SetCursorPos` + `mouse_event`) to physically click on the input area.

### 2. Unicode Handling

Chinese/CJK text is converted to Unicode code point arrays to avoid PowerShell encoding issues:

```python
# Python: convert text to char codes
",".join(str(ord(c)) for c in "你好")  # "20320,22909"

# PowerShell: reconstruct from codes
[string]::new([char[]]@(20320,22909))  # "你好"
```

### 3. Clipboard Safety

Windows clipboard can be locked. Always use a retry loop with `Clear()` before `SetText()`.

### 4. Visual Feedback Loop

Like OpenAI's playwright-interactive, the key pattern is:
```
Execute action → Screenshot → Analyze → Verify → Next action
```

## Prerequisites

- Windows OS with PowerShell
- Python 3.8+
- Target application running and logged in

## How It Differs from OpenAI's playwright-interactive

| Feature | playwright-interactive | Desktop Controller |
|---------|----------------------|-------------------|
| Platform | Codex (cloud) | Claude Code (local) |
| Runtime | js_repl | PowerShell + Python |
| Web apps | Playwright | Playwright |
| **Native desktop apps** | Not supported | **Win32 API** |
| Chat apps (WeChat, etc.) | Not supported | **Full support** |
| Visual feedback | Screenshots in REPL | Screenshots + Claude Vision |
| Session persistence | Kernel-based | Script-based |

The key advantage: **native Windows desktop app control** — something playwright-interactive cannot do since it only supports web/Electron through a browser runtime.

## Extending

To add a new app, edit `scripts/app_registry.py`:

1. Find the process name: `Get-Process | Where-Object { $_.MainWindowTitle -like "*AppName*" }`
2. Identify the search shortcut
3. Determine input area position
4. Add to the `APPS` dictionary

## License

MIT
