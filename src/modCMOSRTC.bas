Attribute VB_Name = "modCMOSRTC"
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

Private cmosrtc_index As Byte
Private cmosrtc_nvram(0& To 127&) As Byte
Private cmosrtc_ext_index As Byte
Private cmosrtc_ext_nvram(0& To 127&) As Byte
Private cmosrtc_i8259Slot As Long
Private cmosrtc_timer As Long
Private cmosfileNum As Integer
Private cmosfileOpen As Byte

Public Function cmosrtc_bcd(ByVal value As Byte) As Byte
    cmosrtc_bcd = CByte((value Mod 10&) Or (((value \ 10&) Mod 10&) * &H10&))
End Function

Public Function cmosrtc_rate_hz(ByVal value As Byte) As Double
    Dim rateBits As Byte

    rateBits = (value And &HF&)
    If (rateBits < 1&) Or (rateBits > 15&) Then
        cmosrtc_rate_hz = 0#
        Exit Function
    End If

    cmosrtc_rate_hz = 32768# / (2# ^ (rateBits - 1&))
End Function

Public Sub cmosrtc_tick(ByVal dummy As Long)
    Dim tdata As SYSTEMTIME
    Dim picSlot As Long

    GetLocalTime tdata

    If (cmosrtc_nvram(&HB&) And &H40&) <> 0& Then
        cmosrtc_nvram(&HC&) = &HC0&
        picSlot = CMOSRTC_GetI8259Slot()
        i8259_doirq picSlot, 0&
    End If

    If (cmosrtc_nvram(&HB&) And &H80&) <> 0& Then
        Exit Sub
    End If

    cmosrtc_nvram(0&) = cmosrtc_bcd(CByte(tdata.wSecond And &HFF&))
    cmosrtc_nvram(2&) = cmosrtc_bcd(CByte(tdata.wMinute And &HFF&))
    cmosrtc_nvram(4&) = cmosrtc_bcd(CByte(tdata.wHour And &HFF&))
    cmosrtc_nvram(6&) = cmosrtc_bcd(CByte(tdata.wDayOfWeek And &HFF&))
    cmosrtc_nvram(7&) = cmosrtc_bcd(CByte(tdata.wDay And &HFF&))
    cmosrtc_nvram(8&) = cmosrtc_bcd(CByte(tdata.wMonth And &HFF&))
    cmosrtc_nvram(9&) = cmosrtc_bcd(CByte((tdata.wYear Mod 100&) And &HFF&))
End Sub

Public Function cmosrtc_read(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim tdata As SYSTEMTIME
    Dim ret As Byte

    ret = &HFF&
    GetLocalTime tdata

    If (addr And &HFFFF&) = &H71& Then
        ret = cmosrtc_nvram(cmosrtc_index And &H7F&)

        If cmosrtc_index = &HC& Then
            cmosrtc_nvram(&HC&) = 0&
        ElseIf cmosrtc_index = &HA& Then
            ret = cmosrtc_nvram(&HA&) And &H7F&
            If tdata.wMilliseconds < 10& Then
                ret = ret Or &H80&
            End If
        ElseIf cmosrtc_index = &HD& Then
            If cmosrtc_nvram(&HD&) = 0& Then
                cmosrtc_nvram(&HD&) = &H80&
                ret = &H80&
            End If
        ElseIf cmosrtc_index = &H3F& Then
            ret = cmosrtc_ext_nvram(cmosrtc_nvram(&H3D&) And &H7F&)
        End If
    End If

    cmosrtc_read = ret
End Function

Public Sub cmosrtc_write(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    Select Case (addr And &HFFFF&)
        Case &H3F&
            cmosrtc_ext_nvram(cmosrtc_nvram(&H3D&) And &H7F&) = value

        Case &H70&
            cmosrtc_index = value And &H7F&

        Case &H71&
            cmosrtc_nvram(cmosrtc_index And &H7F&) = value
            Select Case cmosrtc_index
                Case &HA&
                    timing_updateIntervalFreq cmosrtc_timer, cmosrtc_rate_hz(value)
            End Select

            If cmosfileOpen <> 0& Then
                CMOSRTC_FlushFile
            End If
    End Select
End Sub

Public Sub cmosrtc_init(ByVal cmosfilename As String, ByVal i8259_slave As Long)
    Dim i As Long

    debug_log DEBUG_INFO, "[CMOS] Initializing CMOS + real time clock"

    cmosrtc_index = 0&
    cmosrtc_ext_index = 0&
    cmosrtc_i8259Slot = i8259_slave

    For i = 0& To 127&
        cmosrtc_nvram(i) = 0&
        cmosrtc_ext_nvram(i) = 0&
    Next i

    ports_cbRegister &H70&, 2&, PORTS_CB_CMOSRTC, PORTS_CB_NONE, PORTS_CB_CMOSRTC, PORTS_CB_NONE, 0&

    cmosfileOpen = 0&
    cmosfileNum = 0&

    On Error Resume Next

    cmosfileNum = FreeFile
    Open cmosfilename For Binary Access Read Write As #cmosfileNum
    If Err.Number = 0& Then
        Get #cmosfileNum, 1&, cmosrtc_nvram
        cmosfileOpen = 1&
    Else
        Err.Clear
        cmosfileNum = FreeFile
        Open cmosfilename For Binary Access Write As #cmosfileNum
        If Err.Number = 0& Then
            Put #cmosfileNum, 1&, cmosrtc_nvram
            cmosfileOpen = 1&
        Else
            debug_log DEBUG_INFO, "[CMOS] WARNING: Unable to open file cmos.bin for write. CMOS data will not be preserved!"
        End If
    End If

    On Error GoTo 0&

    cmosrtc_timer = timing_addTimer(TIMER_CB_CMOSRTC_TICK, 0&, 64#, TIMING_ENABLED)
End Sub

Private Function CMOSRTC_GetI8259Slot() As Long
    If cmosrtc_i8259Slot >= 0& Then
        CMOSRTC_GetI8259Slot = cmosrtc_i8259Slot
    Else
        CMOSRTC_GetI8259Slot = machine.i8259_slave
    End If
End Function

Private Sub CMOSRTC_FlushFile()
    On Error Resume Next
    Seek #cmosfileNum, 1&
    Put #cmosfileNum, 1&, cmosrtc_nvram
    On Error GoTo 0&
End Sub

