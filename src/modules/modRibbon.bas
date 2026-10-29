Option Explicit
' Callbacks for the ribbon tab defined in src/customUI14.xml.
' Every button's tag is the name of the macro it runs.

Private mRibbon As IRibbonUI

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub RibbonAction(control As IRibbonControl)
    On Error GoTo Fail
    Application.Run "'" & ThisWorkbook.Name & "'!" & control.Tag
    Exit Sub
Fail:
    Notify Err.Description
End Sub

Public Sub RibbonSupertip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = MacroTip(control.Tag)
End Sub

Public Sub ShortcutsGetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ShortcutsEnabled()
End Sub

Public Sub ShortcutsToggle(control As IRibbonControl, pressed As Boolean)
    SetShortcutsEnabled pressed
End Sub
