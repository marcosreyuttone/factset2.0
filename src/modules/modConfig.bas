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
Public Function Bindings() As Collection
    Dim c As New Collection
    ' Formatting
    AddBinding c, "^+1", "CycleNumber", "Number format cycle"
    AddBinding c, "^+3", "CycleDate", "Date format cycle"
    AddBinding c, "^+4", "CycleCurrency", "Currency format cycle"
    AddBinding c, "^+5", "CyclePercent", "Percent format cycle"
    AddBinding c, "^+8", "CycleMultiple", "Multiple (0.0x) format cycle"
    AddBinding c, "^%+.", "IncreaseDecimal", "Increase decimals"
    AddBinding c, "^%+,", "DecreaseDecimal", "Decrease decimals"
    AddBinding c, "^+9", "CycleFontColor", "Font colour cycle (blue / black / green / red)"
    AddBinding c, "^+k", "CycleFillColor", "Fill colour cycle"
    AddBinding c, "^+b", "CycleBorder", "Border cycle"
    AddBinding c, "^%+c", "ToggleCenterAcross", "Center across selection"
    AddBinding c, "^%+a", "AutoColor", "Auto-colour selection (inputs / formulas / links)"
    ' Auditing
    AddBinding c, "^%+p", "TracePrecedents", "Trace precedents (navigator)"
    AddBinding c, "^%+d", "TraceDependents", "Trace dependents (navigator)"
    AddBinding c, "^%+h", "SelectHardcodes", "Select formulas with hardcoded numbers"
    AddBinding c, "^%+x", "ClearArrows", "Clear trace arrows"
    ' Formula tools
    AddBinding c, "^%+r", "CycleAnchoring", "Anchoring cycle ($A$1 / A$1 / $A1 / A1)"
    AddBinding c, "^%+e", "WrapIfError", "Wrap formulas in IFERROR"
    AddBinding c, "^%+v", "PasteValues", "Paste values"
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

' Border styles, applied in this order: see ApplyBorder in modFormat.
Public Function BorderStyles() As Variant
    BorderStyles = Array("bottom", "top", "total", "outline", "none")
End Function
