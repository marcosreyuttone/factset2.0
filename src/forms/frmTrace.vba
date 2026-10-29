Option Explicit
' Precedent / dependent navigator. Lists every linked cell; moving through the
' list jumps to each cell so you can see it. Controls are created in code so the
' form can be built from plain text.
'
' Keys:  Up/Down = preview   Enter = go there   Right = drill into cell
'        Left = back         Esc = return to where you started

Private WithEvents lstLinks As MSForms.ListBox
Private WithEvents btnGo As MSForms.CommandButton
Private WithEvents btnDrill As MSForms.CommandButton
Private WithEvents btnBack As MSForms.CommandButton
Private WithEvents btnSwitch As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private lblHeader As MSForms.Label

Private mStart As Range         ' cell selected when the navigator opened
Private mCurrent As Range       ' cell whose links are listed
Private mPrecedents As Boolean
Private mLinks As Collection
Private mHistory As Collection
Private mLoading As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " - Trace"
    Me.Width = 600
    Me.Height = 345

    Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader")
    With lblHeader
        .Left = 6: .Top = 6: .Width = 580: .Height = 30
        .WordWrap = True
    End With

    Set lstLinks = Me.Controls.Add("Forms.ListBox.1", "lstLinks")
    With lstLinks
        .Left = 6: .Top = 40: .Width = 580: .Height = 228
        .ColumnCount = 4
        .ColumnWidths = "110 pt;60 pt;110 pt;290 pt"
    End With

    Set btnGo = AddButton("btnGo", "Go", 6)
    Set btnDrill = AddButton("btnDrill", "Drill in  >", 92)
    Set btnBack = AddButton("btnBack", "<  Back", 178)
    Set btnSwitch = AddButton("btnSwitch", "Dependents", 264)
    Set btnCancel = AddButton("btnCancel", "Cancel", 506)
    btnCancel.Cancel = True   ' Esc
End Sub

Private Function AddButton(ByVal ctlName As String, ByVal text As String, ByVal x As Single) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = Me.Controls.Add("Forms.CommandButton.1", ctlName)
    b.Caption = text
    b.Left = x: b.Top = 276: b.Width = 80: b.Height = 24
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
        lstLinks.AddItem SheetLabel(r)
        i = lstLinks.ListCount - 1
        lstLinks.List(i, 1) = r.Address(False, False)
        lstLinks.List(i, 2) = DisplayValue(r)
        lstLinks.List(i, 3) = DisplayFormula(r)
    Next r
    mLoading = False

    lblHeader.Caption = IIf(mPrecedents, "Precedents", "Dependents") & " of " & _
        SheetLabel(mCurrent) & "!" & mCurrent.Address(False, False) & _
        "  (" & mLinks.Count & " found)" & vbLf & DisplayFormula(mCurrent)
    btnSwitch.Caption = IIf(mPrecedents, "Dependents", "Precedents")
    btnBack.Enabled = (mHistory.Count > 0)
    btnGo.Enabled = (mLinks.Count > 0)
    btnDrill.Enabled = (mLinks.Count > 0)

    If lstLinks.ListCount > 0 Then lstLinks.ListIndex = 0
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
    If mLoading Then Exit Sub
    On Error Resume Next
    If Not SelectedLink() Is Nothing Then Application.Goto SelectedLink()
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
