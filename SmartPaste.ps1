# SmartPaste - types a list of values (e.g. copied from an Excel column) into the
# focused field one at a time, pressing a key such as Tab after each value.
# Built for the "one of" criterion in the Raiser's Edge database view query screen.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
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

function Test-EscapePressed {
    ([SmartPasteWin32]::GetAsyncKeyState($VK_ESCAPE) -band 0x8001) -ne 0
}

# ---- Window ----------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SmartPaste'
$form.Size = New-Object System.Drawing.Size(480, 560)
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
$lblDelay.Text = 'Pause between values (milliseconds):'
$lblDelay.Location = New-Object System.Drawing.Point(12, 392)
$lblDelay.AutoSize = $true
$form.Controls.Add($lblDelay)

$numDelay = New-Object System.Windows.Forms.NumericUpDown
$numDelay.Minimum = 0
$numDelay.Maximum = 5000
$numDelay.Increment = 50
$numDelay.Value = 200
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

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start typing'
$btnStart.Location = New-Object System.Drawing.Point(12, 462)
$btnStart.Size = New-Object System.Drawing.Size(440, 34)
$form.Controls.Add($btnStart)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = 'Press Esc or switch windows to stop.'
$lblHint.Location = New-Object System.Drawing.Point(12, 500)
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

    $form.TopMost = $false
    $form.WindowState = 'Minimized'
    for ($i = [int]$numCountdown.Value; $i -gt 0; $i--) {
        $form.Text = "SmartPaste - starting in $i"
        Start-Sleep -Seconds 1
    }
    $form.Text = 'SmartPaste - typing...'

    # Whatever window has focus when the countdown ends is the target; typing stops
    # if focus moves elsewhere so keystrokes never land in the wrong program.
    $target = [SmartPasteWin32]::GetForegroundWindow()
    [void](Test-EscapePressed)  # clear any earlier Esc press
    $typed = 0
    $stopReason = $null
    foreach ($value in $values) {
        if (Test-EscapePressed) { $stopReason = 'Esc was pressed'; break }
        $current = [SmartPasteWin32]::GetForegroundWindow()
        if ($current -ne $target) {
            $stopReason = "the active window changed from `"$([SmartPasteWin32]::GetWindowTitle($target))`" to `"$([SmartPasteWin32]::GetWindowTitle($current))`""
            break
        }
        [System.Windows.Forms.SendKeys]::SendWait((ConvertTo-SendKeysText $value) + $after)
        $typed++
        Start-Sleep -Milliseconds $delay
    }

    $form.Text = 'SmartPaste'
    $form.WindowState = 'Normal'
    $form.TopMost = $true
    $form.Activate()
    $message = "Typed $typed of $($values.Count) value(s)."
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
