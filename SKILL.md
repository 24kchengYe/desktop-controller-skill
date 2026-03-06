---
name: desktop-controller
description: |
  Universal desktop application controller for Windows. Automate any desktop app — send messages, click buttons, type text, take screenshots, and interact with native Windows applications.
  Use this skill when the user asks to: "control my computer", "automate desktop", "send a message on [any chat app]", "click on [something]", "type into [app]", "take a screenshot", "操控电脑", "自动化桌面", "给某某发消息", "打开某某软件", "截个屏", or any task involving interacting with Windows desktop applications.
  Supports: WeChat, WeCom, DingTalk, Feishu/Lark, QQ, Telegram, Slack, Teams, and any other Windows desktop application.
  Inspired by OpenAI's playwright-interactive skill, combining code-based automation with visual screenshot feedback loops.
---

# Universal Desktop Controller

Automate any Windows desktop application using a combination of **Win32 API**, **Playwright** (for web/Electron apps), and **screenshot-based visual feedback**. Inspired by OpenAI's `playwright-interactive` Codex skill, adapted for Claude Code.

## Architecture

```
┌──────────────────────────────────────────────────┐
│             Desktop Controller Skill              │
├──────────────┬──────────────┬────────────────────┤
│  Win32 Layer │ Playwright   │  Visual Feedback   │
│  (Native)    │ (Web/Electron)│  (Screenshot+AI)   │
├──────────────┼──────────────┼────────────────────┤
│ FindWindow   │ Browser      │ CaptureScreen      │
│ SendKeys     │ Page.click   │ CaptureWindow      │
│ SetCursorPos │ Page.fill    │ → Claude Vision    │
│ mouse_event  │ Page.goto    │ → Verify State     │
│ Clipboard    │ Locator      │ → Decide Next Step │
└──────────────┴──────────────┴────────────────────┘
```

## Two Automation Modes

### Mode 1: Win32 Native (for desktop apps)
Best for: WeChat, QQ, DingTalk, WeCom, Feishu, Notepad, File Explorer, etc.

**How it works:**
1. Find the app window by process name
2. Bring to foreground (ShowWindow + SetForegroundWindow)
3. Use keyboard (SendKeys) and mouse (SetCursorPos + mouse_event) automation
4. Use clipboard for text input (handles Unicode/Chinese perfectly)
5. Take screenshots for visual verification

### Mode 2: Playwright (for web & Electron apps)
Best for: Slack, Discord, Teams (web), VS Code, Notion, any browser-based or Electron app.

**How it works:**
1. Launch or connect to browser/Electron app via Playwright
2. Use CSS selectors or coordinates for interaction
3. Built-in screenshot capture
4. DOM inspection for precise element targeting

## Supported Applications

| App | Process Name | Mode | Search Key | Notes |
|-----|-------------|------|------------|-------|
| WeChat | Weixin | Win32 | Ctrl+F | Tested and verified |
| WeCom | WXWork | Win32 | Ctrl+F | Enterprise WeChat |
| DingTalk | DingTalk | Win32 | Ctrl+K | Alibaba's chat |
| Feishu/Lark | Feishu | Win32 | Ctrl+K | ByteDance's chat |
| QQ | QQ | Win32 | Ctrl+F | Tencent QQ |
| Telegram | Telegram | Win32 | Ctrl+K | - |
| Slack | slack | Playwright | Ctrl+K | Electron app |
| Teams | ms-teams | Win32/Playwright | Ctrl+E | - |
| VS Code | Code | Playwright | Ctrl+P | Electron app |
| Any browser | chrome/msedge/firefox | Playwright | - | Via CDP |

## Core Workflow

### Step 1: Identify the target app and automation mode

```python
# Use the app registry to determine process name and mode
python scripts/app_registry.py identify "WeChat"
# Output: { "process": "Weixin", "mode": "win32", "search_key": "Ctrl+F" }
```

### Step 2: Execute the action

**For Win32 native apps:**
```bash
# Send a message to a contact in a chat app
python scripts/desktop_control.py send-message --app weixin --contact "张三" --message "你好"

# Click at specific coordinates
python scripts/desktop_control.py click --app weixin --x 500 --y 400

# Type text into the focused app
python scripts/desktop_control.py type --app weixin --text "Hello World"

# Take a screenshot of a specific app window
python scripts/desktop_control.py screenshot --app weixin --output screenshot.png
```

**For Playwright/Electron apps:**
```bash
# Interact with a web page
python scripts/desktop_control.py web-click --url "http://localhost:3000" --selector "#submit-btn"

# Fill a form field
python scripts/desktop_control.py web-fill --url "http://localhost:3000" --selector "input[name=email]" --text "test@example.com"
```

### Step 3: Visual verification (screenshot feedback loop)

After every action, optionally capture a screenshot and analyze it to verify the action succeeded. This is the key insight from OpenAI's playwright-interactive: **always verify visually**.

```bash
# Take screenshot and return for Claude to analyze
python scripts/desktop_control.py screenshot --app weixin --output verify.png
# Claude reads the screenshot and decides next steps
```

## Key Technical Patterns

### Pattern 1: Chat App Message Sending (Win32)

The universal pattern for sending messages in any chat app:

```
1. FindProcess(process_name) → window handle
2. ShowWindow(handle, SW_RESTORE) + SetForegroundWindow(handle)
3. SendKeys(search_shortcut)        # Open search (Ctrl+F, Ctrl+K, etc.)
4. Clipboard.SetText(contact_name)  # Set contact name
5. SendKeys(Ctrl+V)                 # Paste contact name
6. Sleep(2000)                      # Wait for search results
7. SendKeys(Enter)                  # Select contact
8. Sleep(2500)                      # Wait for chat to load
9. ClickAt(input_area_x, input_area_y)  # CRITICAL: Mouse click to focus input
10. Clipboard.SetText(message)      # Set message
11. SendKeys(Ctrl+V)               # Paste message
12. SendKeys(Enter)                 # Send
```

**Critical insight**: After search+Enter selects a contact, the message input area does NOT automatically get keyboard focus. You MUST use mouse click automation to click on the input area. This was discovered empirically with WeChat and applies to most chat apps.

### Pattern 2: Visual Feedback Loop

```
while not task_complete:
    1. Execute action (click, type, etc.)
    2. Take screenshot
    3. Analyze screenshot (Claude vision)
    4. Determine if action succeeded
    5. If failed → adjust and retry
    6. If succeeded → next action
```

### Pattern 3: Unicode Text Handling

For Chinese/CJK text, always use Unicode code points to avoid encoding issues:

```python
def text_to_char_codes(text):
    return ",".join(str(ord(c)) for c in text)

# In PowerShell, reconstruct from codes:
# [string]::new([char[]]@(20320,22909))  → "你好"
```

### Pattern 4: Clipboard Safety

```powershell
function Set-ClipboardSafe($text) {
    for ($i = 0; $i -lt 5; $i++) {
        try {
            [System.Windows.Forms.Clipboard]::Clear()
            Start-Sleep -Milliseconds 100
            [System.Windows.Forms.Clipboard]::SetText($text)
            return $true
        } catch {
            Start-Sleep -Milliseconds 300
        }
    }
    return $false
}
```

### Pattern 5: Window Position Calculation

Click coordinates are calculated relative to the app window:

```powershell
# Get window rect
$rect = New-Object Win32.RECT
GetWindowRect($hwnd, [ref]$rect)
$winW = $rect.Right - $rect.Left
$winH = $rect.Bottom - $rect.Top

# Chat app input area is typically at bottom-center
$inputX = $rect.Left + [int]($winW * 0.65)
$inputY = $rect.Bottom - [int]($winH * 0.12)
```

## Timing Guidelines

| Operation | Recommended Delay |
|-----------|------------------|
| After ShowWindow/SetForeground | 1000ms |
| After opening search | 600ms |
| After pasting search text | 2000ms |
| After selecting contact (Enter) | 2500ms |
| After clicking input area | 1000ms |
| After pasting message | 800ms |
| Between clipboard Clear and Set | 100ms |

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Window not found | Check process name with `Get-Process` |
| Window won't come to foreground | App may be in system tray; activate manually first |
| Clipboard errors | Use retry loop; close other clipboard-heavy apps |
| Wrong contact selected | Use more specific search term (full name/remark) |
| Input area not focused | Adjust click coordinates for the specific app |
| Chinese text garbled | Use Unicode char code arrays instead of literal strings |
| Screenshot is black | Some apps use hardware acceleration; try with software rendering |

## Extending to New Apps

To add support for a new app:

1. Find the process name: `Get-Process | Where-Object { $_.MainWindowTitle -like "*AppName*" }`
2. Identify the search shortcut (usually Ctrl+F or Ctrl+K)
3. Determine the input area position (take a screenshot and measure)
4. Add to the app registry in `scripts/app_registry.py`
5. Test the full send-message flow
6. Add visual verification screenshots
