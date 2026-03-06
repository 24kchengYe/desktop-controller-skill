"""
Universal Desktop Controller — main automation script.
Combines Win32 API automation with screenshot feedback.

Usage:
    python desktop_control.py send-message --app weixin --contact "张三" --message "你好"
    python desktop_control.py screenshot --app weixin --output screenshot.png
    python desktop_control.py click --app weixin --x 500 --y 400
    python desktop_control.py type --app weixin --text "Hello"
    python desktop_control.py list-apps
    python desktop_control.py find-window --app weixin
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
from app_registry import identify, list_apps


def text_to_char_codes(text: str) -> str:
    """Convert text to comma-separated Unicode code points."""
    return ",".join(str(ord(c)) for c in text)


def generate_ps_header() -> str:
    """Generate PowerShell Win32 API declarations."""
    return r'''
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
public class DesktopCtrl {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern void SetCursorPos(int x, int y);
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left; public int Top; public int Right; public int Bottom;
    }

    public static void ClickAt(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }

    public static void ScreenshotRegion(int x, int y, int w, int h, string path) {
        var bmp = new Bitmap(w, h);
        var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(x, y, 0, 0, new Size(w, h));
        bmp.Save(path, ImageFormat.Png);
        g.Dispose();
        bmp.Dispose();
    }
}
"@

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

function Find-AppWindow($processName) {
    $procs = Get-Process -Name $processName -ErrorAction SilentlyContinue
    $hwnd = [IntPtr]::Zero
    foreach ($p in $procs) {
        if ($p.MainWindowHandle -ne [IntPtr]::Zero) {
            $hwnd = $p.MainWindowHandle
            break
        }
    }
    return $hwnd
}

function Activate-Window($hwnd) {
    [DesktopCtrl]::ShowWindow($hwnd, 9)
    Start-Sleep -Milliseconds 300
    [DesktopCtrl]::SetForegroundWindow($hwnd)
}

function Get-WindowRect($hwnd) {
    $rect = New-Object DesktopCtrl+RECT
    [DesktopCtrl]::GetWindowRect($hwnd, [ref]$rect)
    return $rect
}
'''


def generate_send_message_ps(app_profile: dict, contact_codes: str, message_codes: str) -> str:
    """Generate PowerShell script for sending a message."""
    timings = app_profile["timings"]
    input_area = app_profile["input_area"]

    return generate_ps_header() + f'''
$hwnd = Find-AppWindow "{app_profile['process']}"
if ($hwnd -eq [IntPtr]::Zero) {{
    Write-Error "{app_profile['name']} is not running"
    exit 1
}}

Activate-Window $hwnd
Start-Sleep -Milliseconds {timings['activate']}

$rect = Get-WindowRect $hwnd
$winW = $rect.Right - $rect.Left
$winH = $rect.Bottom - $rect.Top
Write-Output "Window: ${{winW}}x${{winH}}"

# Open search
[System.Windows.Forms.SendKeys]::SendWait("{app_profile['search_key']}")
Start-Sleep -Milliseconds {timings['search_open']}

# Paste contact name
$contactName = [string]::new([char[]]@({contact_codes}))
Set-ClipboardSafe $contactName | Out-Null
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait("^v")
Start-Sleep -Milliseconds {timings['search_paste']}

# Select contact
[System.Windows.Forms.SendKeys]::SendWait("{{ENTER}}")
Start-Sleep -Milliseconds {timings['contact_select']}

# Click input area
$clickX = $rect.Left + [int]($winW * {input_area['x_ratio']})
$clickY = $rect.Bottom - [int]($winH * {input_area['y_ratio_from_bottom']})
Write-Output "Clicking input at: ($clickX, $clickY)"
[DesktopCtrl]::ClickAt($clickX, $clickY)
Start-Sleep -Milliseconds {timings['input_click']}

[DesktopCtrl]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds 500

# Paste message
$message = [string]::new([char[]]@({message_codes}))
Set-ClipboardSafe $message | Out-Null
Start-Sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait("^v")
Start-Sleep -Milliseconds {timings['message_paste']}

# Send
[System.Windows.Forms.SendKeys]::SendWait("{{ENTER}}")
Start-Sleep -Milliseconds 500

Write-Output "Message sent to $contactName via {app_profile['name']}!"
'''


def generate_screenshot_ps(process_name: str, output_path: str, full_screen: bool = False) -> str:
    """Generate PowerShell script for taking a screenshot."""
    if full_screen:
        return generate_ps_header() + f'''
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
[DesktopCtrl]::ScreenshotRegion($screen.X, $screen.Y, $screen.Width, $screen.Height, "{output_path}")
Write-Output "Full screen screenshot saved to {output_path}"
'''
    return generate_ps_header() + f'''
$hwnd = Find-AppWindow "{process_name}"
if ($hwnd -eq [IntPtr]::Zero) {{
    Write-Error "App not running: {process_name}"
    exit 1
}}

Activate-Window $hwnd
Start-Sleep -Milliseconds 500

$rect = Get-WindowRect $hwnd
$w = $rect.Right - $rect.Left
$h = $rect.Bottom - $rect.Top
[DesktopCtrl]::ScreenshotRegion($rect.Left, $rect.Top, $w, $h, "{output_path}")
Write-Output "Window screenshot saved to {output_path} (${{w}}x${{h}})"
'''


def generate_click_ps(process_name: str, x: int, y: int) -> str:
    """Generate PowerShell script for clicking at coordinates."""
    return generate_ps_header() + f'''
$hwnd = Find-AppWindow "{process_name}"
if ($hwnd -eq [IntPtr]::Zero) {{
    Write-Error "App not running: {process_name}"
    exit 1
}}

Activate-Window $hwnd
Start-Sleep -Milliseconds 500

[DesktopCtrl]::ClickAt({x}, {y})
Write-Output "Clicked at ({x}, {y})"
'''


def generate_type_ps(process_name: str, text_codes: str) -> str:
    """Generate PowerShell script for typing text."""
    return generate_ps_header() + f'''
$hwnd = Find-AppWindow "{process_name}"
if ($hwnd -eq [IntPtr]::Zero) {{
    Write-Error "App not running: {process_name}"
    exit 1
}}

Activate-Window $hwnd
Start-Sleep -Milliseconds 500

$text = [string]::new([char[]]@({text_codes}))
Set-ClipboardSafe $text | Out-Null
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait("^v")
Write-Output "Typed text into {process_name}"
'''


def generate_find_window_ps(process_name: str) -> str:
    """Generate PowerShell script for finding a window."""
    return generate_ps_header() + f'''
$hwnd = Find-AppWindow "{process_name}"
if ($hwnd -eq [IntPtr]::Zero) {{
    Write-Output '{{"found": false, "process": "{process_name}"}}'
    exit 0
}}

$rect = Get-WindowRect $hwnd
$w = $rect.Right - $rect.Left
$h = $rect.Bottom - $rect.Top
Write-Output '{{"found": true, "process": "{process_name}", "left": ' + $rect.Left + ', "top": ' + $rect.Top + ', "width": ' + $w + ', "height": ' + $h + '}}'
'''


def run_ps_script(script_content: str, timeout: int = 30) -> tuple[int, str, str]:
    """Write and execute a PowerShell script."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as f:
        f.write(script_content)
        ps_path = f.name

    try:
        result = subprocess.run(
            ["powershell", "-STA", "-File", ps_path],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return 1, "", "Timeout"
    finally:
        os.unlink(ps_path)


def main():
    parser = argparse.ArgumentParser(description="Universal Desktop Controller")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # send-message
    p_send = subparsers.add_parser("send-message", help="Send a message via a chat app")
    p_send.add_argument("--app", required=True, help="App name (e.g., weixin, dingtalk)")
    p_send.add_argument("--contact", required=True, help="Contact name to search for")
    p_send.add_argument("--message", required=True, help="Message to send")

    # screenshot
    p_ss = subparsers.add_parser("screenshot", help="Take a screenshot")
    p_ss.add_argument("--app", help="App name (omit for full screen)")
    p_ss.add_argument("--output", default="screenshot.png", help="Output file path")

    # click
    p_click = subparsers.add_parser("click", help="Click at coordinates")
    p_click.add_argument("--app", required=True, help="App name")
    p_click.add_argument("--x", type=int, required=True, help="X coordinate")
    p_click.add_argument("--y", type=int, required=True, help="Y coordinate")

    # type
    p_type = subparsers.add_parser("type", help="Type text into an app")
    p_type.add_argument("--app", required=True, help="App name")
    p_type.add_argument("--text", required=True, help="Text to type")

    # find-window
    p_find = subparsers.add_parser("find-window", help="Find an app window")
    p_find.add_argument("--app", required=True, help="App name")

    # list-apps
    subparsers.add_parser("list-apps", help="List supported apps")

    args = parser.parse_args()

    if args.command == "list-apps":
        print("Supported applications:")
        for app in list_apps():
            print(f"  {app['key']:<12} {app['name']:<25} process={app['process']:<12} mode={app['mode']:<10} search={app['search_key']}")
        return

    if args.command == "send-message":
        profile = identify(args.app)
        if not profile:
            print(f"Unknown app: {args.app}")
            sys.exit(1)

        contact_codes = text_to_char_codes(args.contact)
        message_codes = text_to_char_codes(args.message)
        script = generate_send_message_ps(profile, contact_codes, message_codes)

        print(f"Sending message via {profile['name']}...")
        print(f"  Contact: {args.contact}")
        print(f"  Message: {args.message}")

        rc, stdout, stderr = run_ps_script(script)
        if stdout:
            print(stdout)
        if stderr:
            print(f"Error: {stderr}", file=sys.stderr)
        sys.exit(rc)

    elif args.command == "screenshot":
        output = os.path.abspath(args.output)
        if args.app:
            profile = identify(args.app)
            if not profile:
                print(f"Unknown app: {args.app}")
                sys.exit(1)
            script = generate_screenshot_ps(profile["process"], output)
        else:
            script = generate_screenshot_ps("", output, full_screen=True)

        rc, stdout, stderr = run_ps_script(script)
        if stdout:
            print(stdout)
        if stderr:
            print(f"Error: {stderr}", file=sys.stderr)
        sys.exit(rc)

    elif args.command == "click":
        profile = identify(args.app)
        if not profile:
            print(f"Unknown app: {args.app}")
            sys.exit(1)
        script = generate_click_ps(profile["process"], args.x, args.y)
        rc, stdout, stderr = run_ps_script(script)
        if stdout:
            print(stdout)
        if stderr:
            print(f"Error: {stderr}", file=sys.stderr)
        sys.exit(rc)

    elif args.command == "type":
        profile = identify(args.app)
        if not profile:
            print(f"Unknown app: {args.app}")
            sys.exit(1)
        text_codes = text_to_char_codes(args.text)
        script = generate_type_ps(profile["process"], text_codes)
        rc, stdout, stderr = run_ps_script(script)
        if stdout:
            print(stdout)
        if stderr:
            print(f"Error: {stderr}", file=sys.stderr)
        sys.exit(rc)

    elif args.command == "find-window":
        profile = identify(args.app)
        if not profile:
            print(f"Unknown app: {args.app}")
            sys.exit(1)
        script = generate_find_window_ps(profile["process"])
        rc, stdout, stderr = run_ps_script(script)
        if stdout:
            print(stdout)
        sys.exit(rc)


if __name__ == "__main__":
    main()
