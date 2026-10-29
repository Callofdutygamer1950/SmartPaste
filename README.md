# SmartPaste

A small replacement for pasting long lists into a **one of** criterion in the
Raiser's Edge database view query screen. It types each value from your list into
the field that has focus and presses a key (Enter by default) after each one, just as
you would by hand.

Nothing to install: it is a PowerShell script, which comes with Windows.

## Files

- `SmartPaste.ps1` - the tool
- `SmartPaste.bat` - double-click this to start it

Keep both files in the same folder.

## How to use

1. In Excel, select the column of values (constituent IDs, codes, ...) and copy it.
2. Double-click `SmartPaste.bat`. The list from your clipboard appears in the window.
   Check the count looks right.
3. In Raiser's Edge, add your criterion, choose **one of**, and get the value box
   ready the way you would to type the first value yourself.
4. Click **Start typing** in SmartPaste, then click into the Raiser's Edge value box
   before the countdown ends.
5. SmartPaste types the values. Press **Esc** or switch to another window to stop.
   A message at the end says how many were typed.

## Settings

- **Keys to press after each value** - what you would press yourself after typing one
  value to add it to the list. `{ENTER}` is the default; `{TAB}`, `{DOWN}` and
  combinations such as `{TAB}{ENTER}` are also available, and you can type your own
  (it uses [SendKeys](https://learn.microsoft.com/dotnet/api/system.windows.forms.sendkeys)
  notation).
- **Pause between values** - raise this if Raiser's Edge drops or mixes up values.
- **Remove duplicate values** - on by default.

If several columns were copied, only the first column is used. Blank lines are skipped.

## Tips

- Try it first with a short list (3-5 values) to confirm the keys setting is right.
- If the .bat file is blocked, right-click `SmartPaste.ps1` > **Properties** and tick
  **Unblock** if shown, or run it from PowerShell:
  `powershell -ExecutionPolicy Bypass -File .\SmartPaste.ps1`
