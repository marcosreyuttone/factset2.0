Option Explicit
' Formula and view tools.

Private mAnchorSelection As String
Private mAnchorStep As Long

Public Sub Notify(ByVal message As String)
    MsgBox message, vbExclamation, APP_NAME
End Sub

' Formula2 (dynamic arrays) only exists in Excel 365, so it is reached late-bound.
Private Function ReadFormula(ByVal c As Range, ByRef useFormula2 As Boolean) As String
    On Error Resume Next
    ReadFormula = CallByName(c, "Formula2", VbGet)
    useFormula2 = (Err.Number = 0)
    If Not useFormula2 Then
        Err.Clear
        ReadFormula = c.Formula
    End If
End Function

Private Sub WriteFormula(ByVal c As Range, ByVal formula As String, ByVal useFormula2 As Boolean)
    If useFormula2 Then
        CallByName c, "Formula2", VbLet, formula
    Else
        c.Formula = formula
    End If
End Sub

Private Function SelectedFormulas() As Range
    On Error Resume Next
    If Selection.CountLarge = 1 Then
        If ActiveCell.HasFormula Then Set SelectedFormulas = ActiveCell
    Else
        Set SelectedFormulas = Selection.SpecialCells(xlCellTypeFormulas)
    End If
End Function

' Each press steps every reference in the selected formulas through
' $A$1 -> A$1 -> $A1 -> A1, like pressing F4 on all of them at once.
Public Sub CycleAnchoring()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim steps As Variant, rng As Range, c As Range, f As String, f2 As Boolean, key As String
    Set rng = SelectedFormulas()
    If rng Is Nothing Then Exit Sub
    steps = Array(xlAbsolute, xlAbsRowRelColumn, xlRelRowAbsColumn, xlRelative)
    key = ActiveSheet.Name & "|" & Selection.Address
    If key = mAnchorSelection Then mAnchorStep = (mAnchorStep + 1) Mod 4 Else mAnchorStep = 0
    mAnchorSelection = key
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        f = ReadFormula(c, f2)
        On Error Resume Next ' ConvertFormula rejects formulas over 255 characters
        f = Application.ConvertFormula(f, xlA1, xlA1, steps(mAnchorStep), c)
        If Err.Number = 0 Then WriteFormula c, f, f2
        Err.Clear
        On Error GoTo Fail
    Next c
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify Err.Description
End Sub

Public Sub WrapIfError()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim rng As Range, c As Range, f As String, f2 As Boolean
    Set rng = SelectedFormulas()
    If rng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        f = ReadFormula(c, f2)
        If UCase$(Left$(f, 9)) <> "=IFERROR(" Then
            WriteFormula c, "=IFERROR(" & Mid$(f, 2) & "," & IFERROR_FALLBACK & ")", f2
        End If
    Next c
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify Err.Description
End Sub

Public Sub PasteValues()
    If Not HasRangeSelection() Then Exit Sub
    If Application.CutCopyMode = False Then
        Notify "Copy something first."
        Exit Sub
    End If
    On Error GoTo Fail
    Selection.PasteSpecial Paste:=xlPasteValues
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub AutofitColumns()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Selection.EntireColumn.AutoFit
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub ToggleGridlines()
    On Error Resume Next
    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub

Public Sub ToggleFreezePanes()
    On Error Resume Next
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
End Sub

Public Sub ShowShortcuts()
    Dim f As frmShortcuts
    Set f = New frmShortcuts
    f.Show vbModal
End Sub
