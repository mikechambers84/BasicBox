Attribute VB_Name = "modDebugLog"
Option Explicit

Public Const DEBUG_NONE As Byte = 0&
Public Const DEBUG_ERROR As Byte = 1&
Public Const DEBUG_INFO As Byte = 2&
Public Const DEBUG_DETAIL As Byte = 3&

Private Const ATTACH_PARENT_PROCESS As Long = -1&
Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const INVALID_HANDLE_VALUE As Long = -1&

Private Declare Function AttachConsole Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteConsoleW Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Long, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, ByVal lpReserved As Long) As Long

Private debug_level As Byte
Private debug_console_ready As Byte
Private debug_console_handle As Long

Public Sub debug_init()
    debug_level = DEBUG_INFO
End Sub

Public Sub debug_setLevel(ByVal level As Byte)
    If level > DEBUG_DETAIL Then
        Exit Sub
    End If
    debug_level = level
End Sub

Private Sub debug_consoleEnsure()
    If debug_console_ready <> 0& Then
        Exit Sub
    End If

    On Error Resume Next
    If AttachConsole(ATTACH_PARENT_PROCESS) = 0& Then
        Call AllocConsole
    End If
    On Error GoTo 0&

    debug_console_handle = GetStdHandle(STD_OUTPUT_HANDLE)
    debug_console_ready = 1&
End Sub

Public Sub debug_consoleWrite(ByVal message As String)
    Dim charsWritten As Long

    debug_consoleEnsure

    If (debug_console_handle <> 0&) And (debug_console_handle <> INVALID_HANDLE_VALUE) Then
        If Len(message) > 0& Then
            If WriteConsoleW(debug_console_handle, StrPtr(message), Len(message), charsWritten, 0&) = 0& Then
                Debug.Print message;
            End If
        End If
    Else
        Debug.Print message;
    End If
End Sub

Public Sub debug_log(ByVal level As Byte, ByVal message As String)
    If level > debug_level Then
        Exit Sub
    End If

    If right$(message, 2&) <> vbCrLf Then message = message & vbCrLf
    debug_consoleWrite message
End Sub
