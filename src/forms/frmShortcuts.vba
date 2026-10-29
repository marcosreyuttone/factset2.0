Option Explicit
' Lists every keyboard shortcut defined in modConfig.

Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim lst As MSForms.ListBox, b As Variant
    Me.Caption = APP_NAME & " shortcuts (" & IIf(ShortcutsEnabled(), "on", "off") & ")"
    Me.Width = 400
    Me.Height = 420

    Set lst = Me.Controls.Add("Forms.ListBox.1", "lstKeys")
    With lst
        .Left = 6: .Top = 6: .Width = 380: .Height = 348
        .ColumnCount = 2
        .ColumnWidths = "120 pt;240 pt"
        For Each b In Bindings()
            .AddItem KeyLabel(b(0))
            .List(.ListCount - 1, 1) = b(2)
        Next b
    End With

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = 306: .Top = 362: .Width = 80: .Height = 24
        .Default = True
        .Cancel = True
    End With
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
