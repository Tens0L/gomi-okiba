Attribute VB_Name = "modBrowserWindow"
Option Explicit

' Launches Chrome and locks its window to a fixed size/position so that
' screen coordinates recorded once stay valid for the whole automation run.

Private mTargetPid As Long
Private mFoundHwnd As LongPtr
Private mSearchTitle As String

' Launches Chrome with the given URL, waits for its window to appear,
' then removes the resize border and forces the requested size/position.
' Returns the window handle (0 if the window could not be located).
Public Function LaunchChromeFixedSize(url As String, _
                                       Optional width As Long = 1280, _
                                       Optional height As Long = 800, _
                                       Optional left As Long = 100, _
                                       Optional top As Long = 100, _
                                       Optional chromePath As String = "") As LongPtr
    Dim path As String
    path = chromePath
    If path = "" Then path = DefaultChromePath()

    Dim cmd As String
    cmd = """" & path & """ --new-window " & _
          "--window-size=" & width & "," & height & " " & _
          "--window-position=" & left & "," & top & " " & _
          """" & url & """"

    Dim pid As Double
    pid = Shell(cmd, vbNormalFocus)

    Dim hwnd As LongPtr
    Dim tries As Long
    Do
        Sleep 200
        hwnd = FindMainWindowByPid(CLng(pid))
        tries = tries + 1
    Loop While hwnd = 0 And tries < 50 ' ~10 seconds max

    If hwnd <> 0 Then
        LockWindowSize hwnd, width, height, left, top
        SetForegroundWindow hwnd
    End If

    LaunchChromeFixedSize = hwnd
End Function

' Locates chrome.exe. First asks Windows itself (the same "App Paths" registry
' key the Start Menu / Run dialog use to resolve "chrome"), which works no
' matter where Chrome was installed. Falls back to the common install
' locations, then raises a clear error if none of that finds it.
' Pass chromePath explicitly to LaunchChromeFixedSize to skip this entirely.
Public Function DefaultChromePath() As String
    Dim fromRegistry As String
    fromRegistry = ReadAppPathFromRegistry("chrome.exe")
    If fromRegistry <> "" And Dir$(fromRegistry) <> "" Then
        DefaultChromePath = fromRegistry
        Exit Function
    End If

    Dim candidates(2) As String
    candidates(0) = Environ$("ProgramFiles") & "\Google\Chrome\Application\chrome.exe"
    candidates(1) = Environ$("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe"
    candidates(2) = Environ$("LocalAppData") & "\Google\Chrome\Application\chrome.exe"

    Dim i As Long
    For i = 0 To UBound(candidates)
        If Dir$(candidates(i)) <> "" Then
            DefaultChromePath = candidates(i)
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 1, "DefaultChromePath", _
        "chrome.exe was not found. Pass the chromePath argument to " & _
        "LaunchChromeFixedSize explicitly, or verify that Chrome is installed."
End Function

' Reads HKLM\...\App Paths\<exeName>\(Default), the registry key Windows
' uses to resolve a bare executable name to its full path. Returns "" (never
' raises) if the key is missing or unreadable.
Private Function ReadAppPathFromRegistry(exeName As String) As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ReadAppPathFromRegistry = wsh.RegRead( _
        "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & exeName & "\")
    On Error GoTo 0
End Function

' Finds the main top-level window belonging to the given process id.
Public Function FindMainWindowByPid(pid As Long) As LongPtr
    mTargetPid = pid
    mFoundHwnd = 0
    EnumWindows AddressOf EnumWindowsByPidCallback, 0
    FindMainWindowByPid = mFoundHwnd
End Function

Public Function EnumWindowsByPidCallback(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
    Dim windowPid As Long
    GetWindowThreadProcessId hwnd, windowPid

    If windowPid = mTargetPid Then
        If IsWindowVisible(hwnd) <> 0 And GetParent(hwnd) = 0 Then
            If GetWindowTextLength(hwnd) > 0 Then
                mFoundHwnd = hwnd
                EnumWindowsByPidCallback = 0 ' stop enumeration
                Exit Function
            End If
        End If
    End If

    EnumWindowsByPidCallback = 1 ' continue
End Function

' Fallback: locate an already-open window whose title contains titleSubstring
' (useful if the browser was opened outside of LaunchChromeFixedSize).
Public Function FindWindowByTitle(titleSubstring As String) As LongPtr
    mSearchTitle = titleSubstring
    mFoundHwnd = 0
    EnumWindows AddressOf EnumWindowsByTitleCallback, 0
    FindWindowByTitle = mFoundHwnd
End Function

Public Function EnumWindowsByTitleCallback(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
    Dim length As Long
    length = GetWindowTextLength(hwnd)
    If length > 0 And IsWindowVisible(hwnd) <> 0 Then
        Dim buf As String
        buf = Space$(length + 1)
        GetWindowText hwnd, buf, length + 1
        buf = Left$(buf, length)
        If InStr(1, buf, mSearchTitle, vbTextCompare) > 0 Then
            mFoundHwnd = hwnd
            EnumWindowsByTitleCallback = 0 ' stop
            Exit Function
        End If
    End If
    EnumWindowsByTitleCallback = 1 ' continue
End Function

' Removes the resizable/maximize window styles and forces the given
' size and position, so the user cannot resize or maximize the window
' after the automation script has started.
Public Sub LockWindowSize(hwnd As LongPtr, width As Long, height As Long, _
                           Optional left As Long = -1, Optional top As Long = -1)
    ShowWindow hwnd, SW_RESTORE

    Dim style As Long
    style = GetWindowLong(hwnd, GWL_STYLE)
    style = style And Not WS_THICKFRAME
    style = style And Not WS_MAXIMIZEBOX
    SetWindowLong hwnd, GWL_STYLE, style

    Dim x As Long, y As Long
    If left = -1 Or top = -1 Then
        Dim r As RECT
        GetWindowRect hwnd, r
        x = r.Left
        y = r.Top
    Else
        x = left
        y = top
    End If

    SetWindowPos hwnd, 0, x, y, width, height, SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
