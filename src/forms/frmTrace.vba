Option Explicit
' Precedent / dependent navigator. Shows a table of every linked cell; moving
' through it jumps to each cell so you can see it. Controls are created in code
' so the form can be built from plain text.
'
' Keys:  Up/Down = preview   Enter = go there   Right = drill into cell
'        Left = back         Esc = return to where you started

Private Const MARGIN As Single = 10
Private Const TABLE_WIDTH As Single = 760
Private Const COL_SHEET As Single = 130
Private Const COL_CELL As Single = 70
Private Const COL_VALUE As Single = 130
Private Const COL_FORMULA As Single = 410

Private WithEvents lstLinks As MSForms.ListBox
Private WithEvents btnGo As MSForms.CommandButton
Private WithEvents btnDrill As MSForms.CommandButton
Private WithEvents btnBack As MSForms.CommandButton
Private WithEvents btnSwitch As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private lblTitle As MSForms.Label
Private txtFormula As MSForms.TextBox
Private txtSelected As MSForms.TextBox

Private mStart As Range         ' cell selected when the navigator opened
Private mCurrent As Range       ' cell whose links are listed
Private mPrecedents As Boolean
Private mLinks As Collection
Private mHistory As Collection
Private mLoading As Boolean

Private Sub UserForm_Initialize()
    Dim x As Single
    Me.Caption = APP_NAME & " - Trace"
    Me.Width = TABLE_WIDTH + 2 * MARGIN + 6
    Me.Height = 450
    Me.BackColor = RGB(255, 255, 255)

    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Left = MARGIN: .Top = 8: .Width = TABLE_WIDTH: .Height = 18
        .Font.Name = "Segoe UI": .Font.Size = 11: .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    Set txtFormula = AddReadOnlyBox("txtFormula", 30, 36)
    txtFormula.MultiLine = True

    x = MARGIN
    AddHeader "Sheet", x, COL_SHEET: x = x + COL_SHEET
    AddHeader "Cell", x, COL_CELL: x = x + COL_CELL
    AddHeader "Value", x, COL_VALUE: x = x + COL_VALUE
    AddHeader "Formula", x, TABLE_WIDTH - (x - MARGIN)

    Set lstLinks = Me.Controls.Add("Forms.ListBox.1", "lstLinks")
    With lstLinks
        .Left = MARGIN: .Top = 90: .Width = TABLE_WIDTH: .Height = 240
        .Font.Name = "Segoe UI": .Font.Size = 9
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(191, 191, 191)
        .ColumnCount = 4
        .ColumnWidths = COL_SHEET & " pt;" & COL_CELL & " pt;" & COL_VALUE & " pt;" & COL_FORMULA & " pt"
    End With

    Set txtSelected = AddReadOnlyBox("txtSelected", 336, 20)

    Dim hint As MSForms.Label
    Set hint = Me.Controls.Add("Forms.Label.1", "lblHint")
    With hint
        .Left = MARGIN: .Top = 370: .Width = 330: .Height = 30
        .Caption = "Up/Down: preview    Enter: go    Right: drill in" & vbLf & "Left: back    Esc: return to start"
        .Font.Name = "Segoe UI": .Font.Size = 8
        .ForeColor = RGB(118, 118, 118)
        .BackStyle = fmBackStyleTransparent
    End With

    Set btnGo = AddButton("btnGo", "Go", 380)
    Set btnDrill = AddButton("btnDrill", "Drill in  >", 456)
    Set btnBack = AddButton("btnBack", "<  Back", 532)
    Set btnSwitch = AddButton("btnSwitch", "Dependents", 608)
    Set btnCancel = AddButton("btnCancel", "Close", 684)
    btnCancel.Cancel = True   ' Esc
End Sub

Private Function AddReadOnlyBox(ByVal ctlName As String, ByVal y As Single, ByVal h As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = Me.Controls.Add("Forms.TextBox.1", ctlName)
    With t
        .Left = MARGIN: .Top = y: .Width = TABLE_WIDTH: .Height = h
        .Font.Name = "Consolas": .Font.Size = 10
        .Locked = True
        .TabStop = False
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(191, 191, 191)
        .BackColor = RGB(248, 248, 248)
    End With
    Set AddReadOnlyBox = t
End Function

Private Sub AddHeader(ByVal text As String, ByVal x As Single, ByVal w As Single)
    Dim header As MSForms.Label
    Set header = Me.Controls.Add("Forms.Label.1")
    With header
        .Caption = " " & text
        .Left = x: .Top = 72: .Width = w: .Height = 18
        .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(0, 51, 102)
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(255, 255, 255)
    End With
End Sub

Private Function AddButton(ByVal ctlName As String, ByVal text As String, ByVal x As Single) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = Me.Controls.Add("Forms.CommandButton.1", ctlName)
    b.Caption = text
    b.Left = x: b.Top = 370: b.Width = 70: b.Height = 24
    b.Font.Name = "Segoe UI"
    Set AddButton = b
End Function

Public Sub Init(ByVal target As Range, ByVal precedents As Boolean)
    Set mStart = target
    Set mHistory = New Collection
    mPrecedents = precedents
    LoadLinks target
End Sub

Private Sub LoadLinks(ByVal target As Range)
    Dim r As Variant, i As Long
    Set mCurrent = target.Cells(1, 1)
    Set mLinks = TraceLinks(mCurrent, mPrecedents)
    Application.Goto mCurrent

    mLoading = True
    lstLinks.Clear
    For Each r In mLinks
        lstLinks.AddItem " " & SheetLabel(r)
        i = lstLinks.ListCount - 1
        lstLinks.List(i, 1) = r.Address(False, False)
        lstLinks.List(i, 2) = DisplayValue(r)
        lstLinks.List(i, 3) = DisplayFormula(r)
    Next r
    mLoading = False

    lblTitle.Caption = IIf(mPrecedents, "Precedents", "Dependents") & " of " & _
        SheetLabel(mCurrent) & "!" & mCurrent.Address(False, False) & _
        "   (" & mLinks.Count & " found)"
    If mCurrent.HasFormula Then
        txtFormula.Text = mCurrent.Formula
    Else
        txtFormula.Text = DisplayValue(mCurrent)
    End If
    txtSelected.Text = ""
    btnSwitch.Caption = IIf(mPrecedents, "Dependents", "Precedents")
    btnBack.Enabled = (mHistory.Count > 0)
    btnGo.Enabled = (mLinks.Count > 0)
    btnDrill.Enabled = (mLinks.Count > 0)

    If lstLinks.ListCount > 0 Then
        lstLinks.ListIndex = 0
    Else
        txtSelected.Text = "No " & IIf(mPrecedents, "precedents", "dependents") & " found."
    End If
    On Error Resume Next ' SetFocus fails before the form is shown
    lstLinks.SetFocus
End Sub

Private Function SheetLabel(ByVal r As Range) As String
    If r.Worksheet.Parent.Name = mStart.Worksheet.Parent.Name Then
        SheetLabel = r.Worksheet.Name
    Else
        SheetLabel = "[" & r.Worksheet.Parent.Name & "]" & r.Worksheet.Name
    End If
End Function

Private Function DisplayValue(ByVal r As Range) As String
    On Error Resume Next
    If r.CountLarge = 1 Then
        DisplayValue = r.Text
        ' Column too narrow to show the number: fall back to the raw value.
        If Len(DisplayValue) > 0 And Replace(DisplayValue, "#", "") = "" Then DisplayValue = CStr(r.Value2)
    Else
        DisplayValue = "(" & r.CountLarge & " cells)"
    End If
End Function

Private Function DisplayFormula(ByVal r As Range) As String
    On Error Resume Next
    If r.CountLarge = 1 Then
        If r.HasFormula Then DisplayFormula = r.Formula
    End If
End Function

Private Function SelectedLink() As Range
    If lstLinks.ListIndex >= 0 Then Set SelectedLink = mLinks(lstLinks.ListIndex + 1)
End Function

Private Sub lstLinks_Change()
    Dim r As Range
    If mLoading Then Exit Sub
    On Error Resume Next
    Set r = SelectedLink()
    If r Is Nothing Then Exit Sub
    txtSelected.Text = SheetLabel(r) & "!" & r.Address(False, False) & "   " & _
        IIf(DisplayFormula(r) = "", DisplayValue(r), DisplayFormula(r))
    Application.Goto r
End Sub

Private Sub lstLinks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnGo_Click
End Sub

Private Sub lstLinks_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn: KeyCode = 0: btnGo_Click
        Case vbKeyRight: KeyCode = 0: btnDrill_Click
        Case vbKeyLeft: KeyCode = 0: btnBack_Click
    End Select
End Sub

Private Sub btnGo_Click()
    Dim r As Range
    Set r = SelectedLink()
    If r Is Nothing Then Exit Sub
    Application.Goto r
    Unload Me
End Sub

Private Sub btnDrill_Click()
    Dim r As Range
    Set r = SelectedLink()
    If r Is Nothing Then Exit Sub
    mHistory.Add mCurrent
    LoadLinks r
End Sub

Private Sub btnBack_Click()
    Dim r As Range
    If mHistory.Count = 0 Then Exit Sub
    Set r = mHistory(mHistory.Count)
    mHistory.Remove mHistory.Count
    LoadLinks r
End Sub

Private Sub btnSwitch_Click()
    mPrecedents = Not mPrecedents
    LoadLinks mCurrent
End Sub

Private Sub btnCancel_Click()
    ReturnToStart
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ReturnToStart
End Sub

Private Sub ReturnToStart()
    On Error Resume Next
    Application.Goto mStart
End Sub
