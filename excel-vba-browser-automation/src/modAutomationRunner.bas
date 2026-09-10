Attribute VB_Name = "modAutomationRunner"
Option Explicit

' Runs a sequence of automation steps listed on a worksheet, so the
' script itself can be edited without touching VBA code.
'
' Expected sheet layout (header row 1, data from row 2), sheet name
' "AutomationSteps" by default:
'
'   A: Action   B: X     C: Y     D: Text        E: WaitMs   F: Url/ChromePath
'   OPEN                          -              -           https://example.com
'   MOVE        640      400
'   CLICK       640      400
'   TYPE                          12345
'   KEY                           {ENTER}
'   WAIT                                          1000
'
' Supported actions: OPEN, MOVE, CLICK, DBLCLICK, RIGHTCLICK,
'                    TYPE, CLEARTYPE (clicks B/C then clears+types D),
'                    KEY, SCROLL (D = notches, positive up / negative down;
'                    B/C optional to move the cursor there first), WAIT.
Public Sub RunAutomationSheet(Optional sheetName As String = "AutomationSteps")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim action As String
        action = UCase$(Trim$(ws.Cells(r, "A").Value))
        If Len(action) = 0 Then GoTo NextRow

        Select Case action
            Case "OPEN"
                LaunchChromeFixedSize CStr(ws.Cells(r, "F").Value), _
                                      CLng(NzNum(ws.Cells(r, "B").Value, 1280)), _
                                      CLng(NzNum(ws.Cells(r, "C").Value, 800))

            Case "MOVE"
                MoveMouseTo CLng(ws.Cells(r, "B").Value), CLng(ws.Cells(r, "C").Value)

            Case "CLICK"
                ClickAt CLng(ws.Cells(r, "B").Value), CLng(ws.Cells(r, "C").Value)

            Case "DBLCLICK"
                DoubleClickAt CLng(ws.Cells(r, "B").Value), CLng(ws.Cells(r, "C").Value)

            Case "RIGHTCLICK"
                RightClickAt CLng(ws.Cells(r, "B").Value), CLng(ws.Cells(r, "C").Value)

            Case "TYPE"
                TypeText CStr(ws.Cells(r, "D").Value)

            Case "CLEARTYPE"
                ClearFieldAndType CLng(ws.Cells(r, "B").Value), _
                                   CLng(ws.Cells(r, "C").Value), _
                                   CStr(ws.Cells(r, "D").Value)

            Case "KEY"
                Application.SendKeys CStr(ws.Cells(r, "D").Value)
                DoEvents

            Case "SCROLL"
                If Len(Trim$(ws.Cells(r, "B").Value)) > 0 And Len(Trim$(ws.Cells(r, "C").Value)) > 0 Then
                    ScrollMouse CLng(ws.Cells(r, "D").Value), _
                                CLng(ws.Cells(r, "B").Value), CLng(ws.Cells(r, "C").Value)
                Else
                    ScrollMouse CLng(ws.Cells(r, "D").Value)
                End If

            Case "WAIT"
                Sleep CLng(NzNum(ws.Cells(r, "E").Value, 500))

            Case Else
                Debug.Print "RunAutomationSheet: skipped unsupported Action '" & action & "' (row " & r & ")"
        End Select

        DoEvents
NextRow:
    Next r
End Sub

Private Function NzNum(v As Variant, defaultValue As Double) As Double
    If IsEmpty(v) Or Len(CStr(v)) = 0 Then
        NzNum = defaultValue
    Else
        NzNum = CDbl(v)
    End If
End Function
