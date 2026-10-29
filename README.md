# FS2Tools – Excel productivity add-in

A Windows Excel add-in (`.xlam`) with FactSet-style productivity shortcuts: format cycles, banking colour conventions, a precedent/dependent navigator, and formula tools. It adds its own **FS2** ribbon tab.

It covers the productivity side only. FactSet data functions (`=FDS(...)`, screening, downloads) need a FactSet subscription and are not included. This project is not affiliated with FactSet.

## Install

1. Download or clone this repository on a Windows PC with Excel (2016, 2019, 2021 or 365).
2. Close Excel, then double-click **`build.cmd`**. It builds `dist\FS2Tools.xlam` and copies it to `%APPDATA%\Microsoft\AddIns`.
3. Open Excel. Go to **File › Options › Add-ins**, set **Manage: Excel Add-ins**, click **Go…**, and tick **FS2Tools**.

The **FS2** tab appears and the shortcuts are active. You only need step 3 once. After that, rebuilding replaces the add-in in place.

> While it runs, the build briefly turns on *Trust access to the VBA project object model* so it can write the code into the workbook. It restores your previous setting afterwards.

## Shortcuts

| Shortcut | Action |
|---|---|
| **Formatting** | |
| Ctrl+Shift+1 | Number format cycle: `#,##0` › `#,##0.0` › `#,##0.00` › General |
| Ctrl+Shift+3 | Date cycle: `m/d/yyyy` › `mmm-yy` › `d-mmm-yy` › `yyyy` › `FYyy` |
| Ctrl+Shift+4 | Currency cycle: `$#,##0` › `$#,##0.0` › `$#,##0.00` |
| Ctrl+Shift+5 | Percent cycle: `0.0%` › `0.00%` › `0%` |
| Ctrl+Shift+8 | Multiple cycle: `0.0x` › `0.00x` › `0x` |
| Ctrl+Alt+Shift+. / , | Increase / decrease decimals |
| Ctrl+Shift+9 | Font colour cycle: blue › black › green › red |
| Ctrl+Shift+K | Fill colour cycle: yellow › grey › light blue › none |
| Ctrl+Shift+B | Border cycle: bottom › top › total (top + double bottom) › outline › none |
| Ctrl+Alt+Shift+C | Center across selection (on/off) |
| Ctrl+Alt+Shift+A | Auto-colour: inputs blue, formulas black, other-sheet links green, other-workbook links red |
| **Auditing** | |
| Ctrl+Alt+Shift+P | Trace precedents navigator |
| Ctrl+Alt+Shift+D | Trace dependents navigator |
| Ctrl+Alt+Shift+H | Select formulas with hardcoded numbers (whole sheet if one cell is selected) |
| Ctrl+Alt+Shift+X | Clear trace arrows |
| **Formulas** | |
| Ctrl+Alt+Shift+R | Anchoring cycle on every selected formula: `$A$1` › `A$1` › `$A1` › `A1` |
| Ctrl+Alt+Shift+E | Wrap selected formulas in `IFERROR(…, 0)` |
| Ctrl+Alt+Shift+V | Paste values |
| **View** | |
| Ctrl+Alt+Shift+W | Autofit columns |
| Ctrl+Alt+Shift+G | Toggle gridlines |
| Ctrl+Alt+Shift+F | Freeze / unfreeze panes at the active cell |
| Ctrl+Alt+Shift+S | Show the shortcut list |

Each format cycle steps forward each time you press it again on the same selection.

### The trace navigator

The precedents/dependents window lists every linked cell, including cells on other sheets and in other open workbooks. For each one it shows the sheet, address, value and formula. Moving through the list jumps to each cell so you can see it in place.

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

- **These shortcuts replace some of Excel's own.** For example, Ctrl+Shift+1 normally applies Excel's number format, Ctrl+Shift+8 selects the current region and Ctrl+Shift+9 unhides rows. Use the **Shortcuts** toggle on the ribbon to turn them all off and get Excel's back. The setting is remembered.
- **Macros clear Excel's undo history.** You can't press Ctrl+Z after running one of these actions. This is how Excel macros work, and FactSet's and Macabacus's add-ins behave the same way.
- Shortcuts don't fire while you are typing in a cell.
- Ctrl+Alt is the same as AltGr on many non-US keyboards. That's why most of the bindings use Ctrl+Alt+**Shift**.

## Project layout

```
build.cmd / build.ps1   build the .xlam (uses Excel through COM)
src/modules/*.bas       VBA modules (config, formatting, auditing, tools, shortcuts, ribbon)
src/forms/*.vba         UserForms (controls are created in code)
src/ThisWorkbook.vba    registers shortcuts when the add-in loads
src/customUI14.xml      the FS2 ribbon tab
```
