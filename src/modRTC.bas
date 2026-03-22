Attribute VB_Name = "modRTC"
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)

Private Function RTC_Bcd(ByVal value As Byte) As Byte
    RTC_Bcd = CByte((((value \ 10&) Mod 10&) * &H10&) Or (value Mod 10&))
End Function

Public Function rtc_read(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim ret As Byte
    Dim tdata As SYSTEMTIME

    ret = &HFF&
    GetLocalTime tdata

    Select Case (addr And &H1F&)
        Case 1&
            ret = CByte((tdata.wMilliseconds \ 10&) And &HFF&)
        Case 2&
            ret = CByte(tdata.wSecond And &HFF&)
        Case 3&
            ret = CByte(tdata.wMinute And &HFF&)
        Case 4&
            ret = CByte(tdata.wHour And &HFF&)
        Case 5&
            ret = CByte(tdata.wDayOfWeek And &HFF&)
        Case 6&
            ret = CByte(tdata.wDay And &HFF&)
        Case 7&
            ret = CByte(tdata.wMonth And &HFF&)
        Case 9&
            ret = CByte((tdata.wYear Mod 100&) And &HFF&)
    End Select

    If ret <> &HFF& Then
        ret = RTC_Bcd(ret)
    End If

    rtc_read = ret
End Function

Public Sub rtc_write(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    ' Intentionally empty for parity with rtc.c.
End Sub

Public Sub rtc_init(ByRef cpu As CPU_t)
    debug_log DEBUG_INFO, "[RTC] Initializing real time clock"
    ports_cbRegister &H240&, &H18&, PORTS_CB_RTC, PORTS_CB_NONE, PORTS_CB_RTC, PORTS_CB_NONE, 0&
End Sub

