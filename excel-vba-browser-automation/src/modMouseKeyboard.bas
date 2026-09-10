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

' Scrolls the mouse wheel by the given number of notches (one notch is one
' typical wheel click). Positive scrolls up, negative scrolls down.
' Pass x/y to move the cursor over a specific area (e.g. a scrollable panel)
' before scrolling; omit them to scroll wherever the cursor already is.
Public Sub ScrollMouse(notches As Long, Optional x As Long = -1, Optional y As Long = -1, _
                        Optional afterDelayMs As Long = 100)
    If x <> -1 And y <> -1 Then MoveMouseTo x, y, False
    mouse_event MOUSEEVENTF_WHEEL, 0, 0, notches * WHEEL_DELTA, 0
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

' Presses Enter, e.g. to submit a form after typing into its last field.
Public Sub PressEnter(Optional afterDelayMs As Long = 100)
    Application.SendKeys "{ENTER}"
    DoEvents
    Sleep afterDelayMs
End Sub

' Presses the Space bar, e.g. to toggle a checkbox or an active button.
Public Sub PressSpace(Optional afterDelayMs As Long = 100)
    Application.SendKeys " "
    DoEvents
    Sleep afterDelayMs
End Sub

' Presses Tab, e.g. to move focus to the next field without clicking it.
Public Sub PressTab(Optional afterDelayMs As Long = 100)
    Application.SendKeys "{TAB}"
    DoEvents
    Sleep afterDelayMs
End Sub

' Presses one of the arrow keys, optionally repeated (e.g. to move N steps
' through a dropdown list or a slider). direction is "UP", "DOWN", "LEFT" or
' "RIGHT" (case-insensitive).
Public Sub PressArrow(direction As String, Optional times As Long = 1, _
                       Optional afterDelayMs As Long = 100)
    Dim code As String
    Select Case UCase$(direction)
        Case "UP": code = "{UP}"
        Case "DOWN": code = "{DOWN}"
        Case "LEFT": code = "{LEFT}"
        Case "RIGHT": code = "{RIGHT}"
        Case Else
            Err.Raise vbObjectError + 3, "PressArrow", _
                "direction must be UP, DOWN, LEFT, or RIGHT."
    End Select

    Dim i As Long
    For i = 1 To times
        Application.SendKeys code
        DoEvents
    Next i
    Sleep afterDelayMs
End Sub

Public Sub PressUp(Optional times As Long = 1, Optional afterDelayMs As Long = 100)
    PressArrow "UP", times, afterDelayMs
End Sub

Public Sub PressDown(Optional times As Long = 1, Optional afterDelayMs As Long = 100)
    PressArrow "DOWN", times, afterDelayMs
End Sub

Public Sub PressLeft(Optional times As Long = 1, Optional afterDelayMs As Long = 100)
    PressArrow "LEFT", times, afterDelayMs
End Sub

Public Sub PressRight(Optional times As Long = 1, Optional afterDelayMs As Long = 100)
    PressArrow "RIGHT", times, afterDelayMs
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
