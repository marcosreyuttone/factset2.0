Option Explicit
' ---------------------------------------------------------------------------
' All keyboard shortcuts, formats and colours live in this module.
' Edit what you like, then run build.cmd again to rebuild the add-in.
'
' OnKey codes:  ^ = Ctrl   + = Shift   % = Alt
' Examples:     "^+1" = Ctrl+Shift+1    "^%+p" = Ctrl+Alt+Shift+P
' Special keys go in braces, e.g. "^+{F4}".
' ---------------------------------------------------------------------------

Public Const APP_NAME As String = "FS2Tools"

' Value used by "Wrap in IFERROR": =IFERROR(formula, IFERROR_FALLBACK)
Public Const IFERROR_FALLBACK As String = "0"

' Banking colour convention used by Auto-colour and the font colour cycle.
Public Function ColorInput() As Long: ColorInput = RGB(0, 0, 255): End Function          ' hardcoded numbers
Public Function ColorFormula() As Long: ColorFormula = RGB(0, 0, 0): End Function        ' formulas
Public Function ColorOtherSheet() As Long: ColorOtherSheet = RGB(0, 128, 0): End Function ' links to other sheets
Public Function ColorExternal() As Long: ColorExternal = RGB(192, 0, 0): End Function     ' links to other workbooks

' Each entry: key, macro name, description (shown in the shortcut list and ribbon tooltips).
' Keys follow FactSet's Excel add-in where FactSet has an equivalent.
Public Function Bindings() As Collection
    Dim c As New Collection
    ' Number formats
    AddBinding c, "^+1", "CycleNumber", "Number format cycle"
    AddBinding c, "^+4", "CycleCurrency", "Currency format cycle"
    AddBinding c, "^+5", "CyclePercent", "Percent format cycle"
    AddBinding c, "^+8", "CycleMultiple", "Multiples (0.0x) format cycle"
    AddBinding c, "^+'", "CycleDate", "Date format cycle"
    AddBinding c, "^+3", "CycleDate", "Date format cycle (alternative key)"
    AddBinding c, "^+y", "CycleBinary", "Binary (Yes/No, True/False, On/Off) format cycle"
    AddBinding c, "^+,", "IncreaseDecimal", "Increase decimals"
    AddBinding c, "^+.", "DecreaseDecimal", "Decrease decimals"
    ' Cell style
    AddBinding c, "^+9", "CycleFontColor", "Font colour cycle (blue / black / green / red)"
    AddBinding c, "^+k", "CycleFillColor", "Cell back colour cycle"
    AddBinding c, "^+b", "CycleBorder", "Border cycle"
    AddBinding c, "^+s", "ToggleSumBar", "Sum bar (line above a total)"
    AddBinding c, "^+u", "CycleUnderline", "Underline cycle"
    AddBinding c, "^+f", "IncreaseFontSize", "Increase font size"
    AddBinding c, "^+g", "DecreaseFontSize", "Decrease font size"
    AddBinding c, "^+c", "CycleCenter", "Center cycle (center / center across / general)"
    AddBinding c, "^+h", "CycleHorizontal", "Horizontal alignment cycle"
    AddBinding c, "^+v", "CycleVertical", "Vertical alignment cycle"
    AddBinding c, "^+w", "ToggleWrapText", "Wrap text"
    AddBinding c, "^%a", "AutoColor", "Auto-colour selection (inputs / formulas / links)"
    AddBinding c, "^%s", "AutoColorSheet", "Auto-colour entire worksheet"
    AddBinding c, "^%+7", "CycleTabColor", "Sheet tab colour cycle"
    ' Formulas
    AddBinding c, "^+d", "SmartCopyDown", "Smart copy down"
    AddBinding c, "^+r", "SmartCopyRight", "Smart copy right"
    AddBinding c, "^+n", "ToggleSign", "Positive / negative toggle"
    AddBinding c, "^+e", "CheckErrors", "Check for errors (selects error cells)"
    AddBinding c, "^%+r", "CycleAnchoring", "Anchoring cycle ($A$1 / A$1 / $A1 / A1)"
    AddBinding c, "^%+e", "WrapIfError", "Wrap formulas in IFERROR"
    AddBinding c, "^%+v", "PasteValues", "Paste values"
    ' Auditing
    AddBinding c, "^+{[}", "TracePrecedents", "Trace precedents (navigator)"
    AddBinding c, "^+{]}", "TraceDependents", "Trace dependents (navigator)"
    AddBinding c, "^%+p", "TracePrecedents", "Trace precedents (navigator, alternative key)"
    AddBinding c, "^%+d", "TraceDependents", "Trace dependents (navigator, alternative key)"
    AddBinding c, "^%+h", "SelectHardcodes", "Select formulas with hardcoded numbers"
    AddBinding c, "^%+x", "ClearArrows", "Clear trace arrows"
    ' Refresh / recalculate
    AddBinding c, "%+q", "RefreshWorkbook", "Refresh workbook (data connections + full recalc)"
    AddBinding c, "%+x", "RefreshWorksheet", "Recalculate worksheet"
    AddBinding c, "%+s", "RefreshSelection", "Recalculate selection"
    ' View
    AddBinding c, "^%+w", "AutofitColumns", "Autofit columns"
    AddBinding c, "^%+g", "ToggleGridlines", "Toggle gridlines"
    AddBinding c, "^%+f", "ToggleFreezePanes", "Freeze / unfreeze panes at cell"
    AddBinding c, "^%+s", "ShowShortcuts", "Show this shortcut list"
    Set Bindings = c
End Function

Private Sub AddBinding(ByVal c As Collection, ByVal key As String, ByVal macro As String, ByVal description As String)
    c.Add Array(key, macro, description)
End Sub

' Format cycles. Each press moves to the next entry; the list wraps around.
Public Function NumberFormats() As Variant
    NumberFormats = Array( _
        "#,##0_);(#,##0);""-""_)", _
        "#,##0.0_);(#,##0.0);""-""_)", _
        "#,##0.00_);(#,##0.00);""-""_)", _
        "General")
End Function

Public Function CurrencyFormats() As Variant
    CurrencyFormats = Array( _
        "$#,##0_);($#,##0);""-""_)", _
        "$#,##0.0_);($#,##0.0);""-""_)", _
        "$#,##0.00_);($#,##0.00);""-""_)")
End Function

Public Function PercentFormats() As Variant
    PercentFormats = Array( _
        "0.0%_);(0.0%);""-""_)", _
        "0.00%_);(0.00%);""-""_)", _
        "0%_);(0%);""-""_)")
End Function

Public Function MultipleFormats() As Variant
    MultipleFormats = Array( _
        "0.0""x""_);(0.0""x"")", _
        "0.00""x""_);(0.00""x"")", _
        "0""x""_);(0""x"")")
End Function

Public Function DateFormats() As Variant
    DateFormats = Array("m/d/yyyy", "mmm-yy", "d-mmm-yy", "yyyy", """FY""yy")
End Function

Public Function FontColors() As Variant
    FontColors = Array(ColorInput(), ColorFormula(), ColorOtherSheet(), ColorExternal())
End Function

' -1 means "no fill".
Public Function FillColors() As Variant
    FillColors = Array(RGB(255, 255, 153), RGB(242, 242, 242), RGB(221, 235, 247), -1)
End Function

'"Binary" formats: show the first word for any non-zero value, the second for zero.
Public Function BinaryFormats() As Variant
    BinaryFormats = Array( _
        """Yes"";""Yes"";""No""", _
        """True"";""True"";""False""", _
        """On"";""On"";""Off""", _
        "General")
End Function

' Sheet tab colours. -1 means "no colour".
Public Function TabColors() As Variant
    TabColors = Array(RGB(255, 192, 0), RGB(0, 112, 192), RGB(0, 176, 80), RGB(166, 166, 166), -1)
End Function

' Border styles, applied in this order: see ApplyBorder in modFormat.
Public Function BorderStyles() As Variant
    BorderStyles = Array("bottom", "top", "total", "outline", "none")
End Function
