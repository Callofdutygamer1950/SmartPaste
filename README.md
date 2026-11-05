# SmartPaste

A small replacement for pasting long lists into a **one of** criterion in the
Raiser's Edge database view query screen. It types each value from your list into
the field that has focus and presses a key (Tab by default, to move to the next row)
after each one, just as you would by hand.

Nothing to install: it is a PowerShell script, which comes with Windows.

## Files

- `SmartPaste.ps1` - the tool
- `SmartPaste.bat` - double-click this to start it
- `SmartPasteHelper.ps1` - optional background helper (see below)
- `SmartPasteHelper.bat` - double-click this to start the helper

Keep all the files in the same folder.

## Helper: button on the one-of box

The helper sits in the system tray (a blue **SP** icon) and shows a blue
**SmartPaste** button in the bottom-right corner of the Raiser's Edge
**Edit Field Criteria** box whenever that box is the active window. Click it to open
SmartPaste, or click it again to bring an already open SmartPaste to the front.
Raiser's Edge runs in Citrix, so the button floats over the box rather than being
part of it, and it shows for every operator, not only **one of**.

Right-click the tray icon for:

- **Open SmartPaste** (double-clicking the icon does the same)
- **Start with Windows** - tick to start the helper each time you sign in to Windows.
  This adds a shortcut named `SmartPaste Helper` to your Startup folder; untick to
  remove it.
- **Exit**

If you move the SmartPaste folder, untick and re-tick **Start with Windows** so the
shortcut points to the new place.

## How to use

1. In Excel, select the column of values (constituent IDs, codes, ...) and copy it.
2. In Raiser's Edge, double-click the field in **Available Fields** and choose
   **one of**. (Open it from the field list, not with **Change...** on an existing
   filter: for lists over 500, SmartPaste reopens the field from that list.)
3. Open SmartPaste: click the **SmartPaste** button on the box if the helper is
   running, or double-click `SmartPaste.bat`. The list from your clipboard appears
   in the window. Check the count looks right.
4. Click **Start typing** in SmartPaste, then click the first empty row in the
   Raiser's Edge one-of box before the countdown ends.
5. SmartPaste types the values. Press **Esc** or switch to another window to stop.
   A message at the end says how many were typed.
6. Check the list in the one-of box, then click **OK** yourself.

### Lists over 500 values

A one-of box holds at most 500 values. When a box is full, SmartPaste clicks OK,
opens the same field again from **Available Fields**, picks **one of**, and carries on
in the new box. Raiser's Edge joins the boxes with **OR** by itself, so the query
matches any value from any box. The last box is left open for you to check and
click **OK**. If a box doesn't open or close as expected, SmartPaste stops and says
why.

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
- **Values per one-of box** - 500 by default, the most a box holds. When more values
  remain, SmartPaste opens a new box (see above); this needs `{TAB}` as the key.

If several columns were copied, only the first column is used. Blank lines are skipped.

## Tips

- Try it first with a short list (3-5 values) to confirm the keys setting is right.
- If it stops early, the message names the window that took over, which shows what
  Raiser's Edge did after the last value (for example, closed the box or opened a
  popup).
- If the helper's SP icon doesn't appear, it shows a message saying why and saves it
  to `%TEMP%\SmartPasteHelper.log`.
- If the .bat file is blocked, right-click `SmartPaste.ps1` > **Properties** and tick
  **Unblock** if shown, or run it from PowerShell:
  `powershell -ExecutionPolicy Bypass -File .\SmartPaste.ps1`
