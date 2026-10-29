# FS2Tools – Excel productivity add-in

A Windows Excel add-in (`.xlam`) with FactSet-style productivity shortcuts: format cycles, banking colour conventions, a precedent/dependent navigator, and formula tools. It adds its own **FS2** ribbon tab.

It covers the productivity side only. FactSet data functions (`=FDS(...)`, screening, downloads) need a FactSet subscription and are not included. This project is not affiliated with FactSet.

## Install

1. **Download:** <https://github.com/marcosreyuttone/factset2.0/archive/refs/heads/claude/excel-shortcuts-addin.zip>
2. Right-click the downloaded zip, choose **Properties**, tick **Unblock** if you see it, and click **OK**. Then right-click it again and choose **Extract All…**.
3. Close Excel. Open the extracted folder and double-click **`build.cmd`**. If Windows shows "Windows protected your PC", click **More info › Run anyway**. The build creates `dist\FS2Tools.xlam` and copies it to Excel's add-ins folder.
4. Open Excel. Go to **File › Options › Add-ins**, set **Manage: Excel Add-ins**, click **Go…**, tick **FS2Tools** and click **OK**.

The **FS2** tab appears and the shortcuts are active. You only need step 4 once. After that, rebuilding replaces the add-in in place.

> While it runs, the build briefly turns on *Trust access to the VBA project object model* so it can write the code into the workbook. It restores your previous setting afterwards.

## Shortcuts

Keys follow FactSet's Excel add-in wherever FactSet has the same feature. Every action is also a button on the **FS2** ribbon tab.

| Shortcut | Action |
|---|---|
| **Number formats** | |
| Ctrl+Shift+1 | Number cycle: `#,##0` › `#,##0.0` › `#,##0.00` › General |
| Ctrl+Shift+4 | Currency cycle: `$#,##0` › `$#,##0.0` › `$#,##0.00` |
| Ctrl+Shift+5 | Percent cycle: `0.0%` › `0.00%` › `0%` |
| Ctrl+Shift+8 | Multiples cycle: `0.0x` › `0.00x` › `0x` |
| Ctrl+Shift+3 (or Ctrl+Shift+') | Date cycle: `dd-mmm-yyyy` (11-Jun-2026) › `mmm-dd-yyyy` (Jun-11-2026) › `dd-mmm-yy` › `mmm-yy` › `m/d/yyyy` › `yyyy` › `FYyy` |
| Ctrl+Shift+Y | Binary cycle: Yes/No › True/False › On/Off › General |
| Ctrl+Shift+, / . | Increase / decrease decimals (also works on General cells) |
| **Cell style** | |
| Ctrl+: (Ctrl+Shift+;) or Ctrl+Shift+9 | Font colour cycle: blue › black › green › red |
| Ctrl+Shift+K | Cell back colour cycle: yellow › grey › light blue › none |
| Ctrl+Shift+B | Border cycle: bottom › top › total (top + double bottom) › outline › none |
| Ctrl+Shift+S | Sum bar: thin line across the top of the selection (on/off) |
| Ctrl+Shift+U | Underline cycle: single › double › single accounting › double accounting › none |
| Ctrl+Shift+F / G | Increase / decrease font size |
| Ctrl+Shift+C | Center cycle: center › center across selection › general |
| Ctrl+Shift+H | Horizontal alignment cycle: left › center › right › general |
| Ctrl+Shift+V | Vertical alignment cycle: top › middle › bottom |
| Ctrl+Shift+W | Wrap text (on/off) |
| Ctrl+Alt+A | Auto-colour selection: inputs blue, formulas black, other-sheet links green, other-workbook links red |
| Ctrl+Alt+S | Auto-colour the entire worksheet |
| Ctrl+Alt+Shift+7 | Sheet tab colour cycle |
| **Formulas** | |
| Ctrl+Shift+D | Smart copy down: fills down as far as the data in the next column goes (or copies the cell above) |
| Ctrl+Shift+R | Smart copy right: fills right as far as the data in the row above goes (or copies the cell to the left) |
| Ctrl+Shift+N | Positive/negative toggle: `5` ↔ `-5`, `=A1+B1` ↔ `=-(A1+B1)` |
| Ctrl+Shift+E | Check for errors: selects every `#N/A`, `#REF!`, `#DIV/0!`… (whole sheet if one cell is selected) |
| Ctrl+Alt+Shift+R | Anchoring cycle on every selected formula: `$A$1` › `A$1` › `$A1` › `A1` |
| Ctrl+Alt+Shift+E | Wrap selected formulas in `IFERROR(…, 0)` |
| Ctrl+Alt+Shift+V | Paste values |
| **Auditing** | |
| Ctrl+Shift+[ (or Ctrl+Alt+Shift+P) | Trace precedents navigator |
| Ctrl+Shift+] (or Ctrl+Alt+Shift+D) | Trace dependents navigator |
| Ctrl+Alt+Shift+H | Select formulas with hardcoded numbers (whole sheet if one cell is selected) |
| Ctrl+Alt+Shift+X | Clear trace arrows |
| **Refresh** | |
| Alt+Shift+Q | Refresh workbook: data connections, then a full recalculation |
| Alt+Shift+X | Recalculate worksheet |
| Alt+Shift+S | Recalculate selection |
| **View** | |
| Ctrl+Alt+Shift+W | Autofit columns |
| Ctrl+Alt+Shift+G | Toggle gridlines |
| Ctrl+Alt+Shift+F | Freeze / unfreeze panes at the active cell |
| Ctrl+Alt+Shift+S | Show the shortcut list |

Each cycle steps forward each time you press it again on the same selection.

**Not included:** FactSet-only features that need a FactSet login. That means the sidebar (Ctrl+M, Ctrl+J), `=FDS` data formulas, and refreshing FactSet data. The refresh shortcuts above refresh Excel's own data connections and recalculate instead.

### The trace navigator

The precedents/dependents window shows a table of every linked cell, including cells on other sheets and in other open workbooks. For each one it shows the sheet, address, value and formula. Moving through the list jumps to each cell so you can see it in place.

| Key | Action |
|---|---|
| ↑ / ↓ | Preview a cell |
| Enter / double-click | Go to the cell and close |
| → | Drill in: trace that cell's own links |
| ← | Back to the previous level |
| Esc | Close and return to where you started |

The **Dependents/Precedents** button switches direction for the current cell.

## Customising

Everything you'd want to change is in `src/modules/modConfig.bas`: key bindings, format lists, colours and the `IFERROR` fallback value. Edit it, close Excel, and run `build.cmd` again.

Key codes are `^` for Ctrl, `+` for Shift and `%` for Alt. For example, `"^+1"` means Ctrl+Shift+1.

## Things to know

- **These shortcuts replace some of Excel's own.** For example, Ctrl+Shift+1 normally applies Excel's number format, Ctrl+: inserts the current time, Ctrl+Shift+V pastes values (Excel 365), Ctrl+Shift+F opens Format Cells, Ctrl+Shift+U expands the formula bar, Ctrl+Shift+[ / ] select all precedents / dependents, Ctrl+Shift+8 selects the current region and Ctrl+Shift+9 unhides rows. Use the **Shortcuts** toggle on the ribbon to turn them all off and get Excel's back. The setting is remembered.
- **Macros clear Excel's undo history.** You can't press Ctrl+Z after running one of these actions. This is how Excel macros work, and FactSet's and Macabacus's add-ins behave the same way.
- Shortcuts don't fire while you are typing in a cell.
- Ctrl+Alt is the same as AltGr on many non-US keyboards. If Ctrl+Alt+A or Ctrl+Alt+S clash with a character you type, change them in `modConfig.bas`.

## Project layout

```
build.cmd / build.ps1   build the .xlam (uses Excel through COM)
src/modules/*.bas       VBA modules (config, formatting, auditing, tools, shortcuts, ribbon)
src/forms/*.vba         UserForms (controls are created in code)
src/ThisWorkbook.vba    registers shortcuts when the add-in loads
src/customUI14.xml      the FS2 ribbon tab (generated by tools/gen_ribbon.py)
```
