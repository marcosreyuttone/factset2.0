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

' Fills the selection down. With a single row selected, fills down as far as the
' data in the neighbouring column goes (like double-clicking the fill handle);
' with nothing to follow, copies the cell above (like Ctrl+D).
Public Sub SmartCopyDown()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim sel As Range, lastRow As Long
    Set sel = Selection
    If sel.Rows.Count > 1 Then
        sel.FillDown
        Exit Sub
    End If
    lastRow = DataEnd(sel.Cells(1, 1), 0, -1, xlDown, True)
    If lastRow = 0 Then lastRow = DataEnd(sel.Cells(1, sel.Columns.Count), 0, 1, xlDown, True)
    If lastRow > sel.Row Then
        sel.Resize(lastRow - sel.Row + 1).FillDown
    ElseIf sel.Row > 1 Then
        sel.Offset(-1).Resize(2).FillDown
    End If
    Exit Sub
Fail:
    Notify Err.Description
End Sub

' Same as SmartCopyDown, but across: follows the data in the row above (or below).
Public Sub SmartCopyRight()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim sel As Range, lastCol As Long
    Set sel = Selection
    If sel.Columns.Count > 1 Then
        sel.FillRight
        Exit Sub
    End If
    lastCol = DataEnd(sel.Cells(1, 1), -1, 0, xlToRight, False)
    If lastCol = 0 Then lastCol = DataEnd(sel.Cells(sel.Rows.Count, 1), 1, 0, xlToRight, False)
    If lastCol > sel.Column Then
        sel.Resize(, lastCol - sel.Column + 1).FillRight
    ElseIf sel.Column > 1 Then
        sel.Offset(, -1).Resize(, 2).FillRight
    End If
    Exit Sub
Fail:
    Notify Err.Description
End Sub

' Row (or column) where the block of data next to `start` ends, or 0 if there is none.
Private Function DataEnd(ByVal start As Range, ByVal rowOff As Long, ByVal colOff As Long, _
                         ByVal direction As XlDirection, ByVal wantRow As Boolean) As Long
    Dim guide As Range
    On Error Resume Next
    Set guide = start.Offset(rowOff, colOff)
    If guide Is Nothing Then Exit Function
    If IsEmpty(guide.Value2) Then Exit Function
    If direction = xlDown Then
        If IsEmpty(guide.Offset(1, 0).Value2) Then Exit Function
    Else
        If IsEmpty(guide.Offset(0, 1).Value2) Then Exit Function
    End If
    If wantRow Then DataEnd = guide.End(direction).Row Else DataEnd = guide.End(direction).Column
End Function

' Flips the sign of numbers and formulas: 5 <-> -5, =A1+B1 <-> =-(A1+B1).
Public Sub ToggleSign()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim rng As Range, c As Range, f As String, f2 As Boolean
    Set rng = Intersect(Selection, ActiveSheet.UsedRange)
    If rng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        If c.HasFormula Then
            f = ReadFormula(c, f2)
            If Left$(f, 3) = "=-(" And ClosesAtEnd(f, 3) Then
                f = "=" & Mid$(f, 4, Len(f) - 4)
            Else
                f = "=-(" & Mid$(f, 2) & ")"
            End If
            WriteFormula c, f, f2
        ElseIf VarType(c.Value2) = vbDouble Then
            c.Value2 = -c.Value2
        End If
    Next c
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify Err.Description
End Sub

' True when the bracket at position `openPos` is closed by the formula's last character.
Private Function ClosesAtEnd(ByVal f As String, ByVal openPos As Long) As Boolean
    Dim i As Long, depth As Long, inText As Boolean, ch As String
    For i = openPos To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then
            inText = Not inText
        ElseIf Not inText Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    ClosesAtEnd = (i = Len(f))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' Selects every cell showing an error (#N/A, #REF!, #DIV/0! ...) in the selection,
' or on the whole sheet when only one cell is selected.
Public Sub CheckErrors()
    If Not HasRangeSelection() Then Exit Sub
    Dim scope As Range, found As Range, part As Range
    If Selection.CountLarge = 1 Then Set scope = ActiveSheet.UsedRange Else Set scope = Selection
    On Error Resume Next
    Set found = scope.SpecialCells(xlCellTypeFormulas, xlErrors)
    Set part = scope.SpecialCells(xlCellTypeConstants, xlErrors)
    On Error GoTo 0
    If Not part Is Nothing Then
        If found Is Nothing Then Set found = part Else Set found = Union(found, part)
    End If
    If found Is Nothing Then
        MsgBox "No errors found.", vbInformation, APP_NAME
    Else
        found.Select
        MsgBox found.Count & " cell(s) with errors selected.", vbExclamation, APP_NAME
    End If
End Sub

Public Sub RefreshWorkbook()
    On Error Resume Next
    ActiveWorkbook.RefreshAll
    Application.CalculateFull
End Sub

Public Sub RefreshWorksheet()
    On Error Resume Next
    ActiveSheet.Calculate
End Sub

Public Sub RefreshSelection()
    If Not HasRangeSelection() Then Exit Sub
    On Error Resume Next
    Selection.Calculate
End Sub
