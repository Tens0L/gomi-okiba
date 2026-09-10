Attribute VB_Name = "modMouseKeyboard"
Option Explicit

' Moves the real Windows cursor and simulates clicks/keystrokes.
' Coordinates are absolute screen pixels (as shown by modCoordinateFinder).

' Moves the cursor from its current position to (x, y).
' When smooth is True the cursor glides in small steps instead of
' teleporting, which behaves more like a human-driven mouse.
Public Sub MoveMouseTo(x As Long, y As Long, _
                        Optional smooth As Boolean = True, _
                        Optional steps As Long = 15, _
                        Optional stepDelayMs As Long = 5)
    If Not smooth Then
        SetCursorPos x, y
        Exit Sub
    End If

    Dim startPt As POINTAPI
    GetCursorPos startPt

    Dim i As Long, ix As Long, iy As Long
    For i = 1 To steps
        ix = startPt.x + CLng((x - startPt.x) * i / steps)
        iy = startPt.y + CLng((y - startPt.y) * i / steps)
        SetCursorPos ix, iy
        Sleep stepDelayMs
    Next i

    SetCursorPos x, y
End Sub

Public Sub ClickAt(x As Long, y As Long, _
                    Optional smooth As Boolean = True, _
                    Optional afterDelayMs As Long = 150)
    MoveMouseTo x, y, smooth
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Sleep 40
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Sleep afterDelayMs
End Sub

Public Sub DoubleClickAt(x As Long, y As Long, Optional smooth As Boolean = True)
    ClickAt x, y, smooth, 40
    ClickAt x, y, False, 150
End Sub

Public Sub RightClickAt(x As Long, y As Long, _
                         Optional smooth As Boolean = True, _
                         Optional afterDelayMs As Long = 150)
    MoveMouseTo x, y, smooth
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    Sleep 40
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep afterDelayMs
End Sub

' Types text into whatever field currently has focus (numbers and
' strings both work, since they are sent as literal keystrokes).
Public Sub TypeText(text As String, Optional afterDelayMs As Long = 100)
    Application.SendKeys EscapeForSendKeys(CStr(text))
    DoEvents
    Sleep afterDelayMs
End Sub

' Clicks the field at (x, y), clears any existing content (Ctrl+A, Delete),
' then types the given value. Handy for both numeric and text inputs.
Public Sub ClearFieldAndType(x As Long, y As Long, text As String)
    ClickAt x, y
    Application.SendKeys "^a"
    DoEvents
    Sleep 50
    Application.SendKeys "{DEL}"
    DoEvents
    Sleep 50
    TypeText text
End Sub

' SendKeys treats + ^ % ~ ( ) { } [ ] as special characters;
' wrap each one in braces so literal text (e.g. "3+4", "a(b)") is typed as-is.
Public Function EscapeForSendKeys(text As String) As String
    Const specials As String = "+^%~(){}[]"
    Dim result As String, i As Long, ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If InStr(specials, ch) > 0 Then
            result = result & "{" & ch & "}"
        Else
            result = result & ch
        End If
    Next i
    EscapeForSendKeys = result
End Function
