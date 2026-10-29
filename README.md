# SmartPaste

A small replacement for pasting long lists into a **one of** criterion in the
Raiser's Edge database view query screen. It types each value from your list into
the field that has focus and presses a key (Tab by default, to move to the next row)
after each one, just as you would by hand.

Nothing to install: it is a PowerShell script, which comes with Windows.

## Files

- `SmartPaste.ps1` - the tool
- `SmartPaste.bat` - double-click this to start it

Keep both files in the same folder.

## How to use

1. In Excel, select the column of values (constituent IDs, codes, ...) and copy it.
2. Double-click `SmartPaste.bat`. The list from your clipboard appears in the window.
   Check the count looks right.
3. In Raiser's Edge, add your criterion and choose **one of**.
4. Click **Start typing** in SmartPaste, then click the first empty row in the
   Raiser's Edge one-of box before the countdown ends.
5. SmartPaste types the values. Press **Esc** or switch to another window to stop.
   A message at the end says how many were typed.
6. Check the list in the one-of box, then click **OK** yourself.

## Settings

- **Keys to press after each value** - what moves you to the next row after typing one
  value. `{TAB}` is the default and is what works in the one-of box: it keeps the
  value and moves to the next empty row. `{DOWN}` is also listed, and you can type
  your own (it uses
  [SendKeys](https://learn.microsoft.com/dotnet/api/system.windows.forms.sendkeys)
  notation). Don't use `{ENTER}` in the one-of box: it acts as OK, closing the box
  after the first value and dropping the value that was being typed.
- **Pause between values** - raise this if Raiser's Edge drops or mixes up values.
- **Remove duplicate values** - on by default.

If several columns were copied, only the first column is used. Blank lines are skipped.

## Tips

- Try it first with a short list (3-5 values) to confirm the keys setting is right.
- If it stops early, the message names the window that took over, which shows what
  Raiser's Edge did after the last value (for example, closed the box or opened a
  popup).
- If the .bat file is blocked, right-click `SmartPaste.ps1` > **Properties** and tick
  **Unblock** if shown, or run it from PowerShell:
  `powershell -ExecutionPolicy Bypass -File .\SmartPaste.ps1`
