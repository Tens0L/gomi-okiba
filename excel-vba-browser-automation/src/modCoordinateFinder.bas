Attribute VB_Name = "modCoordinateFinder"
Option Explicit

' Helper tool for building automation scripts: shows the live mouse
' position in Excel's status bar so you can read off the X/Y values
' to put into the AutomationSteps sheet. Move the mouse over the
' target field in the browser and read the coordinates before they
' disappear (or press Esc to stop early).
Public Sub ShowCursorPositionLive(Optional durationSeconds As Long = 30)
    Dim pt As POINTAPI
    Dim startTime As Single
    startTime = Timer

    Do While Timer - startTime < durationSeconds
        GetCursorPos pt
        Application.StatusBar = "Mouse position: X=" & pt.x & "  Y=" & pt.y & "   (Escで停止)"
        DoEvents

        If (GetAsyncKeyState(vbKeyEscape) And &H8000) <> 0 Then Exit Do
        Sleep 100
    Loop

    Application.StatusBar = False
End Sub
