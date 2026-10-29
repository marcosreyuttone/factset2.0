Option Explicit
' Formatting shortcuts: format cycles, colours, borders, alignment.

' Remembers the last cycle step so repeated presses on the same selection
' keep advancing even when Excel rewrites the format string we applied.
Private mLastCycle As String
Private mLastSelection As String
Private mLastIndex As Long

Public Function HasRangeSelection() As Boolean
    On Error Resume Next
    HasRangeSelection = (TypeName(Selection) = "Range")
End Function

Private Function SelectionKey() As String
    SelectionKey = ActiveWorkbook.Name & "|" & ActiveSheet.Name & "|" & Selection.Address
End Function

Private Function NextIndex(ByVal cycleName As String, ByVal current As Variant, ByVal items As Variant) As Long
    Dim i As Long, j As Long, key As String
    key = SelectionKey()
    If cycleName = mLastCycle And key = mLastSelection Then
        i = mLastIndex + 1
    Else
        i = LBound(items)
        For j = LBound(items) To UBound(items)
            If CStr(items(j)) = CStr(current) Then i = j + 1: Exit For
        Next j
    End If
    If i > UBound(items) Then i = LBound(items)
    mLastCycle = cycleName
    mLastSelection = key
    mLastIndex = i
    NextIndex = i
End Function

Private Sub CycleFormats(ByVal cycleName As String, ByVal formats As Variant)
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Selection.NumberFormat = formats(NextIndex(cycleName, ActiveCell.NumberFormat, formats))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleNumber(): CycleFormats "number", NumberFormats(): End Sub
Public Sub CycleCurrency(): CycleFormats "currency", CurrencyFormats(): End Sub
Public Sub CyclePercent(): CycleFormats "percent", PercentFormats(): End Sub
Public Sub CycleMultiple(): CycleFormats "multiple", MultipleFormats(): End Sub
Public Sub CycleDate(): CycleFormats "date", DateFormats(): End Sub

Public Sub IncreaseDecimal(): ChangeDecimals True: End Sub
Public Sub DecreaseDecimal(): ChangeDecimals False: End Sub

Private Sub ChangeDecimals(ByVal increase As Boolean)
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim fmt As String, re As Object
    fmt = ActiveCell.NumberFormat
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    If fmt = "General" Then
        If Not increase Then Exit Sub
        fmt = "#,##0.0"
    ElseIf increase Then
        If InStr(fmt, ".") > 0 Then
            re.Pattern = "\.(?=[0#])"
            fmt = re.Replace(fmt, ".0")
        Else
            re.Pattern = "0(?![0#,.])"
            fmt = re.Replace(fmt, "0.0")
        End If
    Else
        re.Pattern = "\.([0#]*)[0#](?![0#])"
        fmt = re.Replace(fmt, ".$1")
        re.Pattern = "\.(?![0#])"
        fmt = re.Replace(fmt, "")
    End If
    Selection.NumberFormat = fmt
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleFontColor()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim colors As Variant
    colors = FontColors()
    Selection.Font.Color = colors(NextIndex("font", ActiveCell.Font.Color, colors))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleFillColor()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim colors As Variant, current As Long, c As Long
    colors = FillColors()
    If ActiveCell.Interior.Pattern = xlNone Then current = -1 Else current = ActiveCell.Interior.Color
    c = colors(NextIndex("fill", current, colors))
    If c = -1 Then
        Selection.Interior.Pattern = xlNone
    Else
        Selection.Interior.Color = c
    End If
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleBorder()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim styles As Variant
    styles = BorderStyles()
    ApplyBorder Selection, CStr(styles(NextIndex("border", "", styles)))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Private Sub ApplyBorder(ByVal rng As Range, ByVal style As String)
    rng.Borders.LineStyle = xlNone
    Select Case style
        Case "bottom"
            SetEdge rng, xlEdgeBottom, xlContinuous
        Case "top"
            SetEdge rng, xlEdgeTop, xlContinuous
        Case "total"
            SetEdge rng, xlEdgeTop, xlContinuous
            SetEdge rng, xlEdgeBottom, xlDouble
        Case "outline"
            rng.BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End Select
End Sub

Private Sub SetEdge(ByVal rng As Range, ByVal edge As XlBordersIndex, ByVal edgeStyle As XlLineStyle)
    With rng.Borders(edge)
        .LineStyle = edgeStyle
        If edgeStyle = xlContinuous Then .Weight = xlThin
    End With
End Sub

Public Sub CycleBinary(): CycleFormats "binary", BinaryFormats(): End Sub

Public Sub CycleCenter()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim items As Variant
    items = Array(xlCenter, xlCenterAcrossSelection, xlGeneral)
    Selection.HorizontalAlignment = items(NextIndex("center", ActiveCell.HorizontalAlignment, items))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleHorizontal()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim items As Variant
    items = Array(xlLeft, xlCenter, xlRight, xlGeneral)
    Selection.HorizontalAlignment = items(NextIndex("horizontal", ActiveCell.HorizontalAlignment, items))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleVertical()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim items As Variant
    items = Array(xlTop, xlCenter, xlBottom)
    Selection.VerticalAlignment = items(NextIndex("vertical", ActiveCell.VerticalAlignment, items))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleUnderline()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim items As Variant
    items = Array(xlUnderlineStyleSingle, xlUnderlineStyleDouble, xlUnderlineStyleSingleAccounting, _
        xlUnderlineStyleDoubleAccounting, xlUnderlineStyleNone)
    Selection.Font.Underline = items(NextIndex("underline", ActiveCell.Font.Underline, items))
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub IncreaseFontSize(): ChangeFontSize 1: End Sub
Public Sub DecreaseFontSize(): ChangeFontSize -1: End Sub

Private Sub ChangeFontSize(ByVal delta As Long)
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Dim size As Double
    size = ActiveCell.Font.Size + delta
    If size >= 1 Then Selection.Font.Size = size
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub ToggleWrapText()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    Selection.WrapText = Not ActiveCell.WrapText
    Exit Sub
Fail:
    Notify Err.Description
End Sub

' Thin line along the top of the selection, i.e. above a total row.
Public Sub ToggleSumBar()
    If Not HasRangeSelection() Then Exit Sub
    On Error GoTo Fail
    With Selection.Borders(xlEdgeTop)
        If .LineStyle = xlContinuous Then
            .LineStyle = xlNone
        Else
            .LineStyle = xlContinuous
            .Weight = xlThin
        End If
    End With
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub CycleTabColor()
    On Error GoTo Fail
    Dim colors As Variant, current As Long, c As Long
    colors = TabColors()
    If ActiveSheet.Tab.ColorIndex = xlColorIndexNone Then current = -1 Else current = ActiveSheet.Tab.Color
    c = colors(NextIndex("tab", current, colors))
    If c = -1 Then
        ActiveSheet.Tab.ColorIndex = xlColorIndexNone
    Else
        ActiveSheet.Tab.Color = c
    End If
    Exit Sub
Fail:
    Notify Err.Description
End Sub

' Colours every cell by what it contains: hardcoded number = blue, formula = black,
' other-sheet link = green, other-workbook link = red.
Public Sub AutoColor()
    If Not HasRangeSelection() Then Exit Sub
    AutoColorRange Intersect(Selection, ActiveSheet.UsedRange)
End Sub

Public Sub AutoColorSheet()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    AutoColorRange ActiveSheet.UsedRange
End Sub

Private Sub AutoColorRange(ByVal rng As Range)
    If rng Is Nothing Then Exit Sub
    On Error GoTo Fail
    Dim c As Range, re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\[[^\]]+\.xl\w*\]"
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        If c.HasFormula Then
            If re.Test(c.Formula) Then
                c.Font.Color = ColorExternal()
            ElseIf InStr(c.Formula, "!") > 0 Then
                c.Font.Color = ColorOtherSheet()
            Else
                c.Font.Color = ColorFormula()
            End If
        ElseIf VarType(c.Value2) = vbDouble Then
            c.Font.Color = ColorInput()
        End If
    Next c
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify Err.Description
End Sub
