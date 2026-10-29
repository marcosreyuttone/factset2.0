Option Explicit

Private Sub Workbook_Open()
    If ShortcutsEnabled() Then RegisterShortcuts
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UnregisterShortcuts
End Sub

Private Sub Workbook_AddinUninstall()
    UnregisterShortcuts
End Sub
