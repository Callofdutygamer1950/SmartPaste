# SmartPaste - types a list of values (e.g. copied from an Excel column) into the
# focused field one at a time, pressing a key such as Tab after each value.
# Built for the "one of" criterion in the Raiser's Edge database view query screen.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class SmartPasteWin32 {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int maxCount);
    public static string GetWindowTitle(IntPtr hWnd) {
        var sb = new System.Text.StringBuilder(256);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr extraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr extraInfo; }
    [StructLayout(LayoutKind.Explicit)]
    struct INPUTUNION { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
    [StructLayout(LayoutKind.Sequential)]
    struct INPUT { public uint type; public INPUTUNION u; }
    // ---- Esc guard: a low-level keyboard hook that catches Esc before any program
    // (Citrix, Raiser's Edge, or whoever owns Esc as a hotkey) sees it. It runs on
    // its own thread with its own message loop, because Windows skips a hook whose
    // thread is busy, and the typing thread often is.
    [StructLayout(LayoutKind.Sequential)]
    struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y; }
    delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern IntPtr SetWindowsHookEx(int idHook, HookProc proc, IntPtr hMod, uint threadId);
    [DllImport("user32.dll")] static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);
    [DllImport("user32.dll")] static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll")] static extern IntPtr GetModuleHandle(string name);

    public static volatile bool EscPressed;
    static HookProc hookProc = EscHook;  // kept in a field so it isn't garbage-collected
    static System.Threading.Thread hookThread;
    static uint hookThreadId;

    static IntPtr EscHook(int code, IntPtr wParam, IntPtr lParam) {
        if (code >= 0 && Marshal.ReadInt32(lParam) == 0x1B) {  // VK_ESCAPE
            int msg = wParam.ToInt32();
            if (msg == 0x100 || msg == 0x104) EscPressed = true;  // WM_KEYDOWN, WM_SYSKEYDOWN
            return (IntPtr)1;  // swallow both key down and key up
        }
        return CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
    }

    // Starts catching Esc. Returns false if Windows refused the hook.
    public static bool StartEscGuard() {
        EscPressed = false;
        bool ok = false;
        var ready = new System.Threading.ManualResetEvent(false);
        hookThread = new System.Threading.Thread(() => {
            hookThreadId = GetCurrentThreadId();
            IntPtr hook = SetWindowsHookEx(13, hookProc, GetModuleHandle(null), 0);  // WH_KEYBOARD_LL
            ok = hook != IntPtr.Zero;
            ready.Set();
            if (!ok) return;
            MSG m;
            while (GetMessage(out m, IntPtr.Zero, 0, 0) > 0) { }
            UnhookWindowsHookEx(hook);
        });
        hookThread.IsBackground = true;
        hookThread.Start();
        ready.WaitOne();
        return ok;
    }

    public static void StopEscGuard() {
        if (hookThread == null) return;
        PostThreadMessage(hookThreadId, 0x0012, IntPtr.Zero, IntPtr.Zero);  // WM_QUIT
        hookThread.Join(1000);
        hookThread = null;
    }

    [DllImport("user32.dll")] static extern uint SendInput(uint count, INPUT[] inputs, int size);
    [DllImport("user32.dll")] static extern short VkKeyScan(char ch);
    [DllImport("user32.dll")] static extern uint MapVirtualKey(uint code, uint mapType);

    static INPUT Key(ushort vk, bool up, bool extended) {
        var i = new INPUT();
        i.type = 1;  // INPUT_KEYBOARD
        i.u.ki.wVk = vk;
        i.u.ki.wScan = (ushort)MapVirtualKey(vk, 0);
        i.u.ki.dwFlags = (up ? 2u : 0u) | (extended ? 1u : 0u);  // KEYUP, EXTENDEDKEY
        return i;
    }

    // Presses the keys for text, then trailingVk if it isn't 0, in one SendInput call.
    // Returns the number of keys pressed, or -1 (sending nothing) if a character has
    // no plain key on this keyboard layout, such as one needing AltGr.
    public static int TypeText(string text, ushort trailingVk) {
        var inputs = new List<INPUT>();
        int keys = 0;
        foreach (char ch in text) {
            short scan = VkKeyScan(ch);
            if (scan == -1 || (scan & 0x600) != 0) return -1;  // no key, or needs Ctrl/Alt
            ushort vk = (ushort)(scan & 0xFF);
            bool shift = (scan & 0x100) != 0;
            if (shift) inputs.Add(Key(0x10, false, false));
            inputs.Add(Key(vk, false, false));
            inputs.Add(Key(vk, true, false));
            if (shift) inputs.Add(Key(0x10, true, false));
            keys++;
        }
        if (trailingVk != 0) {
            bool extended = trailingVk >= 0x21 && trailingVk <= 0x28;  // navigation keys
            inputs.Add(Key(trailingVk, false, extended));
            inputs.Add(Key(trailingVk, true, extended));
            keys++;
        }
        SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));
        return keys;
    }
}
"@

$VK_ESCAPE = 0x1B

# SendKeys treats these characters as commands, so they must be wrapped in braces
# to be typed literally.
function ConvertTo-SendKeysText([string]$text) {
    $special = '+^%~(){}[]'
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $text.ToCharArray()) {
        if ($special.IndexOf($ch) -ge 0) { [void]$sb.Append('{').Append($ch).Append('}') }
        else { [void]$sb.Append($ch) }
    }
    $sb.ToString()
}

# Splits pasted text into values: one per line, first column only if several
# columns were copied, trimmed, blanks dropped.
function Get-ListValues([string]$text, [bool]$removeDuplicates) {
    $values = @($text -split "\r?\n" |
        ForEach-Object { ($_ -split "`t")[0].Trim() } |
        Where-Object { $_ -ne '' })
    if ($removeDuplicates) { $values = @($values | Select-Object -Unique) }
    $values
}

# While typing, SmartPaste catches Esc itself (Register-EscapeKey), so pressing it
# stops SmartPaste and never reaches Raiser's Edge, where it would act as Cancel and
# throw away the open box. Once pressed, it stays "pressed" until the next run.
$script:escPressed = $false

function Register-EscapeKey {
    $script:escPressed = $false
    # If Windows refuses, fall back to watching the key, which still stops typing
    # but lets Esc through to Raiser's Edge.
    $script:escGuarded = [SmartPasteWin32]::StartEscGuard()
}

function Unregister-EscapeKey {
    [SmartPasteWin32]::StopEscGuard()
    $script:escGuarded = $false
}

function Test-EscapePressed {
    if ([SmartPasteWin32]::EscPressed -or
        ([SmartPasteWin32]::GetAsyncKeyState($VK_ESCAPE) -band 0x8000) -ne 0) {
        $script:escPressed = $true
    }
    $script:escPressed
}

# Sleeps for $ms, watching for Esc. Returns $true if Esc was pressed.
function Wait-UnlessEscape([int]$ms) {
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    while ($clock.ElapsedMilliseconds -lt $ms) {
        if (Test-EscapePressed) { return $true }
        Start-Sleep -Milliseconds ([Math]::Min(50, [Math]::Max(1, $ms - $clock.ElapsedMilliseconds)))
    }
    Test-EscapePressed
}

# Keys from the "keys after each value" list that can go in the same batch as the
# value itself.
$trailingKeys = @{ '{TAB}' = 0x09; '{DOWN}' = 0x28 }

# Types one value and the key after it. The keys go in a single batch, which is far
# faster through Citrix than SendKeys (that waits after every key), then this waits
# $keyPauseMs per key so Raiser's Edge keeps up and keys never pile up in a queue
# (otherwise Esc and the window check would act too late). Values with characters
# that have no plain key fall back to SendKeys.
function Send-Value([string]$value, [string]$after, [int]$keyPauseMs) {
    $vk = if ($trailingKeys.ContainsKey($after)) { $trailingKeys[$after] } else { 0 }
    $keys = [SmartPasteWin32]::TypeText($value, $vk)
    if ($keys -lt 0) {
        [System.Windows.Forms.SendKeys]::SendWait((ConvertTo-SendKeysText $value) + $after)
        return
    }
    if ($vk -eq 0) { [System.Windows.Forms.SendKeys]::SendWait($after) }
    [void](Wait-UnlessEscape ($keys * $keyPauseMs))
}

# Citrix shows the Raiser's Edge criteria box as its own window, titled like
# "Edit Field Criteria - \\Remote".
function Test-CriteriaBox([IntPtr]$hWnd) {
    [SmartPasteWin32]::GetWindowTitle($hWnd) -like 'Edit Field Criteria*'
}

# Polls the active window until $condition (given its handle) is true. Returns the
# handle, or IntPtr.Zero after $timeoutMs or if Esc is pressed.
function Wait-ForActiveWindow([scriptblock]$condition, [int]$timeoutMs) {
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    while ($clock.ElapsedMilliseconds -lt $timeoutMs) {
        $hWnd = [SmartPasteWin32]::GetForegroundWindow()
        if (& $condition $hWnd) { return $hWnd }
        if (Wait-UnlessEscape 100) { break }
    }
    [IntPtr]::Zero
}

# Closes a full one-of box and opens a new one for the same field, ready to type
# in its first row. Returns the new box's handle, or a string saying why it could
# not. Raiser's Edge joins the new box to the old one with OR by itself. Esc stops
# it between steps, never partway through a keystroke sequence.
function Open-NextCriteriaBox([IntPtr]$box) {
    $esc = 'Esc was pressed'
    # Let the last values finish arriving before closing the box.
    if (Wait-UnlessEscape 1000) { return $esc }
    # After Tab on the last row, focus is on OK, so Enter saves the box.
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $query = Wait-ForActiveWindow { param($h) $h -ne $box -and -not (Test-CriteriaBox $h) } 5000
    if ($script:escPressed) { return "$esc (after the full box was saved)" }
    if ($query -eq [IntPtr]::Zero) { return 'the full one-of box did not close' }
    # Focus is back on the field in the Available Fields list; Enter opens a new box.
    if (Wait-UnlessEscape 300) { return "$esc (after the full box was saved)" }
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $newBox = Wait-ForActiveWindow { param($h) Test-CriteriaBox $h } 10000
    $emptyBox = "$esc while a new box was opening. The full boxes were saved; click Cancel on the new, empty box"
    if ($script:escPressed) { return $emptyBox }
    if ($newBox -eq [IntPtr]::Zero) { return 'a new one-of box did not open' }
    # Operator gets focus: "o" picks "one of", then Tab and Shift+Tab land in row 1.
    if (Wait-UnlessEscape 300) { return $emptyBox }
    [System.Windows.Forms.SendKeys]::SendWait('o{TAB}+{TAB}')
    Start-Sleep -Milliseconds 300
    $newBox
}

# ---- Window ----------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SmartPaste'
$form.Size = New-Object System.Drawing.Size(480, 592)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = 'Values to enter (one per line):'
$lblList.Location = New-Object System.Drawing.Point(12, 10)
$lblList.AutoSize = $true
$form.Controls.Add($lblList)

$txtList = New-Object System.Windows.Forms.TextBox
$txtList.Multiline = $true
$txtList.ScrollBars = 'Vertical'
$txtList.WordWrap = $false
$txtList.Location = New-Object System.Drawing.Point(12, 30)
$txtList.Size = New-Object System.Drawing.Size(440, 260)
$form.Controls.Add($txtList)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point(12, 296)
$lblCount.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($lblCount)

$btnClipboard = New-Object System.Windows.Forms.Button
$btnClipboard.Text = 'Load from clipboard'
$btnClipboard.Location = New-Object System.Drawing.Point(312, 294)
$btnClipboard.Size = New-Object System.Drawing.Size(140, 26)
$form.Controls.Add($btnClipboard)

$chkDedupe = New-Object System.Windows.Forms.CheckBox
$chkDedupe.Text = 'Remove duplicate values'
$chkDedupe.Checked = $true
$chkDedupe.Location = New-Object System.Drawing.Point(12, 326)
$chkDedupe.AutoSize = $true
$form.Controls.Add($chkDedupe)

$lblAfter = New-Object System.Windows.Forms.Label
$lblAfter.Text = 'Keys to press after each value:'
$lblAfter.Location = New-Object System.Drawing.Point(12, 360)
$lblAfter.AutoSize = $true
$form.Controls.Add($lblAfter)

$cmbAfter = New-Object System.Windows.Forms.ComboBox
# Tab saves the value and moves to the next row of the one-of grid. Enter is not
# offered: it presses the dialog's OK button, closing it and dropping the value
# being typed.
$cmbAfter.Items.AddRange(@('{TAB}', '{DOWN}'))
$cmbAfter.Text = '{TAB}'
$cmbAfter.Location = New-Object System.Drawing.Point(240, 356)
$cmbAfter.Size = New-Object System.Drawing.Size(212, 24)
$form.Controls.Add($cmbAfter)

$lblDelay = New-Object System.Windows.Forms.Label
$lblDelay.Text = 'Pause per key typed (milliseconds):'
$lblDelay.Location = New-Object System.Drawing.Point(12, 392)
$lblDelay.AutoSize = $true
$form.Controls.Add($lblDelay)

# Raiser's Edge through Citrix took about 8 ms per key in testing; 12 leaves room.
$numDelay = New-Object System.Windows.Forms.NumericUpDown
$numDelay.Minimum = 1
$numDelay.Maximum = 200
$numDelay.Increment = 2
$numDelay.Value = 12
$numDelay.Location = New-Object System.Drawing.Point(240, 388)
$numDelay.Size = New-Object System.Drawing.Size(80, 24)
$form.Controls.Add($numDelay)

$lblCountdown = New-Object System.Windows.Forms.Label
$lblCountdown.Text = 'Seconds to click into Raiser''s Edge:'
$lblCountdown.Location = New-Object System.Drawing.Point(12, 424)
$lblCountdown.AutoSize = $true
$form.Controls.Add($lblCountdown)

$numCountdown = New-Object System.Windows.Forms.NumericUpDown
$numCountdown.Minimum = 1
$numCountdown.Maximum = 30
$numCountdown.Value = 5
$numCountdown.Location = New-Object System.Drawing.Point(240, 420)
$numCountdown.Size = New-Object System.Drawing.Size(80, 24)
$form.Controls.Add($numCountdown)

$lblPerBox = New-Object System.Windows.Forms.Label
$lblPerBox.Text = 'Values per one-of box (then open a new box):'
$lblPerBox.Location = New-Object System.Drawing.Point(12, 456)
$lblPerBox.AutoSize = $true
$form.Controls.Add($lblPerBox)

# The one-of grid holds at most 500 rows.
$numPerBox = New-Object System.Windows.Forms.NumericUpDown
$numPerBox.Minimum = 1
$numPerBox.Maximum = 500
$numPerBox.Value = 500
$numPerBox.Location = New-Object System.Drawing.Point(280, 452)
$numPerBox.Size = New-Object System.Drawing.Size(80, 24)
$form.Controls.Add($numPerBox)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start typing'
$btnStart.Location = New-Object System.Drawing.Point(12, 494)
$btnStart.Size = New-Object System.Drawing.Size(440, 34)
$form.Controls.Add($btnStart)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = 'Press Esc to stop at any time (the one-of box keeps what was typed).'
$lblHint.Location = New-Object System.Drawing.Point(12, 532)
$lblHint.AutoSize = $true
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

# ---- Behaviour -------------------------------------------------------------

function Update-Count {
    $n = (Get-ListValues $txtList.Text $chkDedupe.Checked).Count
    $lblCount.Text = "$n value(s) ready"
}

$txtList.Add_TextChanged({ Update-Count })
$chkDedupe.Add_CheckedChanged({ Update-Count })

$btnClipboard.Add_Click({
    $txtList.Text = [System.Windows.Forms.Clipboard]::GetText()
})

$btnStart.Add_Click({
    $values = Get-ListValues $txtList.Text $chkDedupe.Checked
    if ($values.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('There are no values to type.', 'SmartPaste')
        return
    }
    $after = $cmbAfter.Text
    $delay = [int]$numDelay.Value
    $perBox = [int]$numPerBox.Value
    if ($values.Count -gt $perBox -and $after -ne '{TAB}') {
        [void][System.Windows.Forms.MessageBox]::Show(
            "There are more than $perBox values, so SmartPaste will need to open new one-of boxes. That only works with {TAB} as the key after each value.",
            'SmartPaste')
        return
    }

    $form.TopMost = $false
    $form.WindowState = 'Minimized'
    $typed = 0
    $inBox = 0
    $boxes = 1
    $stopReason = $null
    # Esc stops SmartPaste from here until the end, including during the countdown.
    Register-EscapeKey
    try {
        for ($i = [int]$numCountdown.Value; $i -gt 0; $i--) {
            $form.Text = "SmartPaste - starting in $i"
            if (Wait-UnlessEscape 1000) { $stopReason = 'Esc was pressed before typing started'; break }
        }
        $form.Text = 'SmartPaste - typing...'

        # Whatever window has focus when the countdown ends is the target; typing
        # stops if focus moves elsewhere so keystrokes never land in the wrong program.
        $target = [SmartPasteWin32]::GetForegroundWindow()
        foreach ($value in $values) {
            if ($stopReason) { break }
            if (Test-EscapePressed) { $stopReason = 'Esc was pressed'; break }
            if ($inBox -eq $perBox) {
                if (-not (Test-CriteriaBox $target)) {
                    $stopReason = "$perBox values were typed and a new one-of box is needed, but the window typed into was not a Raiser's Edge criteria box"
                    break
                }
                $form.Text = 'SmartPaste - opening a new box...'
                $next = Open-NextCriteriaBox $target
                if ($next -is [string]) { $stopReason = $next; break }
                $target = $next
                $inBox = 0
                $boxes++
                $form.Text = 'SmartPaste - typing...'
            }
            $current = [SmartPasteWin32]::GetForegroundWindow()
            if ($current -ne $target) {
                $stopReason = "the active window changed from `"$([SmartPasteWin32]::GetWindowTitle($target))`" to `"$([SmartPasteWin32]::GetWindowTitle($current))`""
                break
            }
            Send-Value $value $after $delay
            $typed++
            $inBox++
        }
    } finally {
        Unregister-EscapeKey
    }

    $form.Text = 'SmartPaste'
    $form.WindowState = 'Normal'
    $form.TopMost = $true
    $form.Activate()
    $message = "Typed $typed of $($values.Count) value(s)"
    $message += if ($boxes -gt 1) { " across $boxes one-of boxes." } else { '.' }
    if ($stopReason) { $message += "`n`nStopped because $stopReason." }
    [void][System.Windows.Forms.MessageBox]::Show($message, 'SmartPaste')
})

$form.Add_Shown({
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $txtList.Text = [System.Windows.Forms.Clipboard]::GetText()
    }
    Update-Count
})

[void]$form.ShowDialog()
