Option Explicit
' Registers the keyboard shortcuts from modConfig with Excel.

Public Function ShortcutsEnabled() As Boolean
    ShortcutsEnabled = (GetSetting(APP_NAME, "General", "ShortcutsOn", "1") = "1")
End Function

Public Sub SetShortcutsEnabled(ByVal enabled As Boolean)
    SaveSetting APP_NAME, "General", "ShortcutsOn", IIf(enabled, "1", "0")
    If enabled Then RegisterShortcuts Else UnregisterShortcuts
End Sub

Public Sub RegisterShortcuts()
    Dim b As Variant
    On Error Resume Next
    Dim k As Variant
    For Each b In Bindings()
        For Each k In Split(b(0), "|")
            Application.OnKey CStr(k), "'" & ThisWorkbook.Name & "'!" & b(1)
        Next k
    Next b
End Sub

Public Sub UnregisterShortcuts()
    Dim b As Variant
    On Error Resume Next
    Dim k As Variant
    For Each b In Bindings()
        For Each k In Split(b(0), "|")
            Application.OnKey CStr(k)
        Next k
    Next b
End Sub

' "^%+a" -> "Ctrl+Alt+Shift+A"
Public Function KeyLabel(ByVal key As String) As String
    Dim i As Long, s As String
    key = Split(key, "|")(0)
    For i = 1 To Len(key)
        Select Case Mid$(key, i, 1)
            Case "^": s = s & "Ctrl+"
            Case "%": s = s & "Alt+"
            Case "+": s = s & "Shift+"
            Case Else
                s = s & UCase$(Replace(Replace(Mid$(key, i), "{", ""), "}", ""))
                Exit For
        End Select
    Next i
    KeyLabel = s
End Function

' Tooltip text for a macro, e.g. "Number format cycle (Ctrl+Shift+1)".
Public Function MacroTip(ByVal macro As String) As String
    Dim b As Variant
    For Each b In Bindings()
        If b(1) = macro Then
            MacroTip = b(2) & " (" & KeyLabel(b(0)) & ")"
            Exit Function
        End If
    Next b
End Function
