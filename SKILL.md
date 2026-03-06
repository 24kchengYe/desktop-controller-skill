---
name: desktop-controller
description: |
  Universal desktop application controller for Windows. Automate any desktop app — send messages, click buttons, type text, take screenshots, and interact with native Windows applications.
  Use this skill when the user asks to: "control my computer", "automate desktop", "send a message on [any chat app]", "click on [something]", "type into [app]", "take a screenshot", "操控电脑", "自动化桌面", "给某某发消息", "打开某某软件", "截个屏", or any task involving interacting with Windows desktop applications.
  WeChat triggers: "send a WeChat message", "message someone on WeChat", "打开微信发消息", "给某某发微信", "微信发送", "微信发消息", "用微信发", "帮我发微信".
  Email/Web triggers: "打开邮箱", "发邮件", "send email", "open browser", "操控浏览器", "网页操作".
  Supports: WeChat, WeCom, DingTalk, Feishu/Lark, QQ, Telegram, Slack, Teams, any browser-based web app, and any other Windows desktop application.
  Inspired by OpenAI's playwright-interactive skill, combining code-based automation with visual screenshot feedback loops.
  ROUTING: For web/Electron apps (email, Slack, browser), prefer Playwright with CSS selectors for speed and precision. For native desktop apps (WeChat, QQ, DingTalk), use Win32 API automation.
---

# Universal Desktop Controller

Automate any Windows desktop application using a combination of **Win32 API**, **Playwright** (for web/Electron apps), and **screenshot-based visual feedback**. Inspired by OpenAI's `playwright-interactive` Codex skill, adapted for Claude Code.

## Smart Routing: Choose the Right Engine

**CRITICAL DECISION**: Before any action, determine which engine to use:

```
User Request → Is it a web/browser/Electron app?
                 ├── YES → Playwright (CSS selectors, instant, precise)
                 │         Examples: email webmail, Slack, browser tasks, web forms
                 │         ✅ page.click('#compose')  — instant, 100% accurate
                 │
                 └── NO  → Win32 API (native desktop automation)
                           Examples: WeChat, QQ, DingTalk, Notepad, File Explorer
                           ✅ SetCursorPos + mouse_event — works for any native app
```

**Why this matters**: For the Tsinghua email (a web app), using Playwright CSS selectors is 10-50x faster than screenshot→analyze→click. OpenAI's playwright-interactive proved this: DOM selectors beat vision-based coordinate guessing every time for web content.

| Engine | Speed | Precision | Use For |
|--------|-------|-----------|---------|
| Playwright | <1s per action | 100% (DOM selector) | Web apps, Electron apps, browser tasks |
| Win32 API | 2-3s per action | 95% (coordinate-based) | Native desktop apps (WeChat, QQ, etc.) |
| Screenshot+Vision | 10-15s per action | 80% (AI guessing) | Last resort / unknown UI |

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

### Mode 1: Win32 Native (for desktop apps — WeChat, QQ, etc.)
Best for: WeChat, QQ, DingTalk, WeCom, Feishu, Notepad, File Explorer, etc.

**How it works:**
1. Find the app window by process name
2. Bring to foreground (ShowWindow + SetForegroundWindow)
3. Use keyboard (SendKeys) and mouse (SetCursorPos + mouse_event) automation
4. Use clipboard for text input (handles Unicode/Chinese perfectly)
5. Take screenshots for visual verification

### Mode 2: Playwright (for web & Electron apps — email, Slack, etc.)
Best for: Email webmail, Slack, Discord, Teams (web), VS Code, Notion, any browser-based or Electron app.

**How it works:**
1. Connect to running Chrome via CDP (Chrome DevTools Protocol) or launch new browser
2. Use CSS selectors for precise, instant element targeting — NO screenshots needed
3. `page.click()`, `page.fill()`, `page.goto()` for all interactions
4. DOM inspection for finding the right selectors
5. 10-50x faster than screenshot+coordinate approach

**Playwright Quick Start for Web Apps (e.g., sending email):**
```javascript
// Connect to existing Chrome (must be launched with --remote-debugging-port=9222)
const browser = await chromium.connectOverCDP('http://127.0.0.1:9222');
const page = browser.contexts()[0].pages()[0]; // Get current tab

// Example: Compose email in Coremail (Tsinghua email)
await page.click('span:has-text("写信")');           // Click compose — instant!
await page.fill('input[name="to"]', '575860760@qq.com');  // Fill recipient
await page.fill('.compose-body', '今天不去吃饭了');          // Fill body
await page.click('button:has-text("发送")');                // Send
```

**vs. Screenshot approach (old, slow):**
```
Step 1: Take screenshot (2s)
Step 2: AI analyzes image to find "写信" button position (5s)
Step 3: Click at guessed coordinates (80, 145) — might miss! (1s)
Step 4: Take screenshot again to verify (2s)
Step 5: Repeat for each UI element...
Total: 30-60 seconds for a simple email
```

**How to connect to existing Chrome:**
```bash
# First, restart Chrome with debugging port enabled:
chrome.exe --remote-debugging-port=9222
# Then Playwright can connect to it and control any tab
```

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
