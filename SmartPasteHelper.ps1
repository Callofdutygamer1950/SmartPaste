# SmartPaste Helper - runs in the background (system tray) and shows a small
# SmartPaste button over the bottom-right corner of the Raiser's Edge "Edit Field
# Criteria" box whenever that box is the active window. Clicking it opens SmartPaste.
# Raiser's Edge runs in Citrix, so the button floats above the box rather than
# being part of it.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public static class SmartPasteHelperWin32 {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int maxCount);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder text, int maxCount);
    public static string GetWindowTitle(IntPtr hWnd) {
        var sb = new System.Text.StringBuilder(256);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
    public static string GetWindowClass(IntPtr hWnd) {
        var sb = new System.Text.StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}
// A window that stays on top and never takes focus, so clicking it leaves the
// Raiser's Edge box as the active window.
public class SmartPasteOverlay : Form {
    protected override bool ShowWithoutActivation { get { return true; } }
    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x08000000 | 0x00000080 | 0x00000008; // NOACTIVATE | TOOLWINDOW | TOPMOST
            return cp;
        }
    }
}
"@

# Only one helper at a time, so repeated starts don't stack up tray icons.
$mutex = New-Object System.Threading.Mutex($false, 'Local\SmartPasteHelper')
if (-not $mutex.WaitOne(0)) { exit }

# Work in real screen pixels so the button lines up with the Citrix window on
# scaled displays.
[void][SmartPasteHelperWin32]::SetProcessDPIAware()

$smartPastePath = Join-Path $PSScriptRoot 'SmartPaste.ps1'
$startupShortcut = Join-Path ([Environment]::GetFolderPath('Startup')) 'SmartPaste Helper.lnk'

# Citrix shows each Raiser's Edge window as its own local window, titled like
# "Edit Field Criteria - \\Remote".
$criteriaClass = 'Transparent Windows Client'
$criteriaTitle = 'Edit Field Criteria*'

# ---- Opening SmartPaste ------------------------------------------------------

$script:smartPasteProcess = $null

function Open-SmartPaste {
    $p = $script:smartPasteProcess
    if ($p -and -not $p.HasExited) {
        # Already open: bring it to the front instead of opening a second copy.
        $p.Refresh()
        $hWnd = $p.MainWindowHandle
        if ($hWnd -ne [IntPtr]::Zero) {
            if ([SmartPasteHelperWin32]::IsIconic($hWnd)) { [void][SmartPasteHelperWin32]::ShowWindow($hWnd, 9) }  # SW_RESTORE
            [void][SmartPasteHelperWin32]::SetForegroundWindow($hWnd)
        }
        return
    }
    $script:smartPasteProcess = Start-Process powershell.exe -PassThru -ArgumentList @(
        '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
        '-File', "`"$smartPastePath`"")
}

# ---- Start with Windows ------------------------------------------------------

function Test-StartWithWindows { Test-Path $startupShortcut }

function Set-StartWithWindows([bool]$enabled) {
    if ($enabled) {
        $shell = New-Object -ComObject WScript.Shell
        $lnk = $shell.CreateShortcut($startupShortcut)
        $lnk.TargetPath = Join-Path $PSHOME 'powershell.exe'
        $lnk.Arguments = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
        $lnk.WorkingDirectory = $PSScriptRoot
        $lnk.WindowStyle = 7  # minimized, so the console barely flashes at sign-in
        $lnk.Description = 'SmartPaste Helper'
        $lnk.Save()
    } elseif (Test-Path $startupShortcut) {
        Remove-Item $startupShortcut
    }
}

# ---- Tray icon ---------------------------------------------------------------

function New-SmartPasteIcon {
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 103, 192))), 1, 1, 30, 30)
    $font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = 'Center'
    $format.LineAlignment = 'Center'
    $g.DrawString('SP', $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0, 0, 32, 32)), $format)
    $g.Dispose()
    [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon = New-SmartPasteIcon
$tray.Text = 'SmartPaste Helper'

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $menu.Items.Add('Open SmartPaste')
$miOpen.Font = New-Object System.Drawing.Font($miOpen.Font, [System.Drawing.FontStyle]::Bold)
$miStartup = New-Object System.Windows.Forms.ToolStripMenuItem('Start with Windows')
$miStartup.Checked = Test-StartWithWindows
[void]$menu.Items.Add($miStartup)
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$miExit = $menu.Items.Add('Exit')
$tray.ContextMenuStrip = $menu

$menu.Add_Opening({ $miStartup.Checked = Test-StartWithWindows })
$miOpen.Add_Click({ Open-SmartPaste })
$tray.Add_MouseDoubleClick({ Open-SmartPaste })
$miStartup.Add_Click({
    Set-StartWithWindows (-not $miStartup.Checked)
    $miStartup.Checked = Test-StartWithWindows
})

# ---- Floating button ---------------------------------------------------------

$overlay = New-Object SmartPasteOverlay
$overlay.FormBorderStyle = 'None'
$overlay.ShowInTaskbar = $false
$overlay.StartPosition = 'Manual'
$overlay.TopMost = $true

$scale = $overlay.DeviceDpi / 96
$overlay.Size = New-Object System.Drawing.Size([int](92 * $scale), [int](26 * $scale))

$btnOverlay = New-Object System.Windows.Forms.Button
$btnOverlay.Text = 'SmartPaste'
$btnOverlay.Dock = 'Fill'
$btnOverlay.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$btnOverlay.FlatStyle = 'Flat'
$btnOverlay.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
$btnOverlay.ForeColor = [System.Drawing.Color]::White
$btnOverlay.TabStop = $false
$overlay.Controls.Add($btnOverlay)
$btnOverlay.Add_Click({ Open-SmartPaste })

# Keeps the button on the box's bottom-right corner while the box is active,
# and hidden the rest of the time.
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250
$timer.Add_Tick({
    $hWnd = [SmartPasteHelperWin32]::GetForegroundWindow()
    $isCriteriaBox = ([SmartPasteHelperWin32]::GetWindowClass($hWnd) -eq $criteriaClass) -and
        ([SmartPasteHelperWin32]::GetWindowTitle($hWnd) -like $criteriaTitle)
    if (-not $isCriteriaBox) {
        if ($overlay.Visible) { $overlay.Hide() }
        return
    }
    $rect = New-Object SmartPasteHelperWin32+RECT
    [void][SmartPasteHelperWin32]::GetWindowRect($hWnd, [ref]$rect)
    $margin = [int](10 * $scale)
    $location = New-Object System.Drawing.Point(
        ($rect.Right - $overlay.Width - $margin), ($rect.Bottom - $overlay.Height - $margin))
    if ($overlay.Location -ne $location) { $overlay.Location = $location }
    if (-not $overlay.Visible) { $overlay.Show() }
})

# ---- Run ---------------------------------------------------------------------

$miExit.Add_Click({
    $timer.Stop()
    $tray.Visible = $false
    [System.Windows.Forms.Application]::Exit()
})

$tray.Visible = $true
$timer.Start()
[System.Windows.Forms.Application]::Run()
$tray.Dispose()
$mutex.ReleaseMutex()
