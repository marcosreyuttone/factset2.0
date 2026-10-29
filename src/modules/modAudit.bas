Option Explicit
' Auditing: the precedent / dependent navigator and hardcode finder.

Private Const MAX_ARROWS As Long = 1000

Public Sub TracePrecedents(): ShowTrace True: End Sub
Public Sub TraceDependents(): ShowTrace False: End Sub

Private Sub ShowTrace(ByVal precedents As Boolean)
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim f As frmTrace
    Set f = New frmTrace
    f.Init ActiveCell, precedents
    f.Show vbModal
    Exit Sub
Fail:
    Notify Err.Description
End Sub

' Returns every cell/range that feeds into (precedents = True) or depends on
' (precedents = False) the given cell, including links on other sheets and in
' other open workbooks. Uses Excel's trace arrows, so it sees exactly what
' Formulas > Trace Precedents would show. Moves the selection: callers restore it.
Public Function TraceLinks(ByVal target As Range, ByVal precedents As Boolean) As Collection
    Dim res As New Collection, seen As Object
    Dim c As Range, r As Range, ws As Worksheet
    Dim arrowNum As Long, linkNum As Long, found As Boolean
    Dim originKey As String, key As String
    Dim screen As Boolean, alerts As Boolean

    Set TraceLinks = res
    Set c = target.Cells(1, 1)
    If precedents And Not c.HasFormula Then Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = c.Worksheet
    originKey = c.Address(External:=True)
    screen = Application.ScreenUpdating
    alerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo Done
    ws.Parent.Activate
    ws.Activate
    ws.ClearArrows
    If precedents Then c.ShowPrecedents Else c.ShowDependents

    For arrowNum = 1 To MAX_ARROWS
        found = False
        For linkNum = 1 To MAX_ARROWS
            Set r = Nothing
            On Error Resume Next
            Set r = c.NavigateArrow(precedents, arrowNum, linkNum)
            Err.Clear
            On Error GoTo Done
            If r Is Nothing Then Exit For
            key = r.Address(External:=True)
            If key = originKey Then Exit For
            found = True
            ' Excel repeats the last link once the arrow has no more.
            If seen.Exists(key) Then Exit For
            seen.Add key, True
            res.Add r
        Next linkNum
        If Not found Then Exit For
    Next arrowNum

Done:
    On Error Resume Next
    ws.ClearArrows
    Application.DisplayAlerts = alerts
    Application.ScreenUpdating = screen
End Function

Public Sub ClearArrows()
    On Error Resume Next
    ActiveSheet.ClearArrows
End Sub

' True when a formula contains a typed-in number, e.g. =A1*1.05 or =SUM(A1:A5)+100.
Public Function HasHardcodedNumber(ByVal formula As String) As Boolean
    Dim re As Object, s As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    s = formula
    re.Pattern = """[^""]*""": s = re.Replace(s, "")                       ' text literals
    re.Pattern = "'[^']*'!": s = re.Replace(s, "")                          ' 'Sheet name'!
    re.Pattern = "\[[^\]]*\]": s = re.Replace(s, "")                        ' [Book.xlsx], Table[Column]
    re.Pattern = "\$?[A-Za-z]{1,3}\$?\d+": s = re.Replace(s, "")            ' A1, $B$2
    re.Pattern = "\$?\d+:\$?\d+": s = re.Replace(s, "")                     ' 1:1 row ranges
    re.Pattern = "[A-Za-z_\\][A-Za-z0-9_\.]*": s = re.Replace(s, "")        ' names, functions, sheets
    re.Pattern = "\d"
    HasHardcodedNumber = re.Test(s)
End Function

' Selects every formula with a hardcoded number in the selection
' (or on the whole sheet when only one cell is selected).
Public Sub SelectHardcodes()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim scope As Range, formulas As Range, c As Range, hits As Range
    If Selection.CountLarge = 1 Then Set scope = ActiveSheet.UsedRange Else Set scope = Selection
    On Error Resume Next
    Set formulas = scope.SpecialCells(xlCellTypeFormulas)
    On Error GoTo Fail
    If Not formulas Is Nothing Then
        For Each c In formulas.Cells
            If HasHardcodedNumber(c.Formula) Then
                If hits Is Nothing Then Set hits = c Else Set hits = Union(hits, c)
            End If
        Next c
    End If
    If hits Is Nothing Then
        MsgBox "No formulas with hardcoded numbers found.", vbInformation, APP_NAME
    Else
        hits.Select
        Application.StatusBar = APP_NAME & ": " & hits.Count & " formula(s) with hardcoded numbers selected"
        Application.OnTime Now + TimeSerial(0, 0, 5), "'" & ThisWorkbook.Name & "'!ResetStatusBar"
    End If
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub
