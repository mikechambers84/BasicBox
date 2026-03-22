Attribute VB_Name = "modI8253"
Option Explicit

Public Const PIT_MODE_LATCHCOUNT As Byte = 0&
Public Const PIT_MODE_LOBYTE As Byte = 1&
Public Const PIT_MODE_HIBYTE As Byte = 2&
Public Const PIT_MODE_TOGGLE As Byte = 3&

Private i8253_chandata(0& To 2&) As Integer
Private i8253_accessmode(0& To 2&) As Byte
Private i8253_bytetoggle(0& To 2&) As Byte
Private i8253_effectivedata(0& To 2&) As Long
Private i8253_chanfreq(0& To 2&) As Single

Private i8253_active(0& To 2&) As Byte
Private i8253_counter(0& To 2&) As Long
Private i8253_reload(0& To 2&) As Long
Private i8253_mode(0& To 2&) As Byte
Private i8253_dataflipflop(0& To 2&) As Byte
Private i8253_bcd(0& To 2&) As Byte
Private i8253_rlmode(0& To 2&) As Byte
Private i8253_latch(0& To 2&) As Integer
Private i8253_out(0& To 2&) As Byte

Private i8253_i8259Slot As Long

Public Sub i8253_write(ByVal dummy As Long, ByVal portnum As Integer, ByVal value As Byte)
    Dim sel As Byte
    Dim rl As Byte
    Dim loaded As Byte
    Dim p As Long

    p = (portnum And 3&)
    loaded = 0&

    Select Case p
        Case 0& To 2&
            Select Case i8253_rlmode(p)
                Case 1&
                    i8253_reload(p) = CLng(value) * &H100&
                    i8253_active(p) = 1&
                    loaded = 1&

                Case 2&
                    i8253_reload(p) = CLng(value)
                    i8253_active(p) = 1&
                    loaded = 1&

                Case 3&
                    If i8253_dataflipflop(p) = 0& Then
                        i8253_reload(p) = (i8253_reload(p) And &HFF00&) Or CLng(value)
                    Else
                        i8253_reload(p) = (i8253_reload(p) And &HFF&) Or (CLng(value) * &H100&)
                        i8253_counter(p) = i8253_reload(p)
                        If i8253_reload(p) = 0& Then
                            i8253_reload(p) = 65536&
                        End If
                        i8253_active(p) = 1&
                        loaded = 1&
                    End If
                    i8253_dataflipflop(p) = i8253_dataflipflop(p) Xor 1&
            End Select

            If loaded <> 0& Then
                Select Case i8253_mode(p)
                    Case 0&, 1&
                        i8253_out(p) = 0&
                    Case 2&, 3&
                        i8253_out(p) = 1&
                End Select
            End If

        Case 3&
            sel = (value And &HC0&) \ &H40&
            If sel = 3& Then Exit Sub

            rl = (value And &H30&) \ &H10&
            If rl = 0& Then
                i8253_latch(sel) = CInt(i8253_counter(sel) And &HFFFF&)
            Else
                i8253_rlmode(sel) = rl
                i8253_mode(sel) = (value And &HE&) \ 2&
                If (i8253_mode(sel) And &H2&) <> 0& Then
                    i8253_mode(sel) = i8253_mode(sel) And 3&
                End If
                i8253_bcd(sel) = (value And 1&)
            End If
            i8253_dataflipflop(sel) = 0&
    End Select
End Sub

Public Function i8253_read(ByVal dummy As Long, ByVal portnum As Integer) As Byte
    Dim p As Long
    Dim ret As Byte

    p = (portnum And 3&)
    If p = 3& Then
        i8253_read = &HFF&
        Exit Function
    End If

    If i8253_active(p) = 0& Then
        i8253_read = &HFF&
        Exit Function
    End If

    Select Case i8253_rlmode(p)
        Case 1&
            i8253_read = CByte((i8253_latch(p) And &HFF00&) \ &H100&)

        Case 2&
            i8253_read = CByte(i8253_latch(p) And &HFF&)

        Case Else
            If i8253_dataflipflop(p) = 0& Then
                ret = CByte(i8253_latch(p) And &HFF&)
            Else
                ret = CByte((i8253_latch(p) And &HFF00&) \ &H100&)
            End If
            i8253_dataflipflop(p) = i8253_dataflipflop(p) Xor 1&
            i8253_read = ret
    End Select
End Function

Private Sub i8253_timerCallback0()
    diag_count_irq0
    i8259_doirq i8253_i8259Slot, 0&
End Sub

Private Sub i8253_timerCallback1()
End Sub

Private Sub i8253_timerCallback2()
End Sub

Public Sub i8253_tickCallback(ByVal dummy As Long)
    Dim i As Long
    Dim gateVal As Byte

    For i = 0& To 2&
        If (i = 2&) And (i8253_mode(2&) <> 3&) Then
            pcspeaker_setGateState 0&, PC_SPEAKER_GATE_TIMER2, 0&
        End If

        If i8253_active(i) <> 0& Then
            Select Case i8253_mode(i)
                Case 0&
                    i8253_counter(i) = i8253_counter(i) - 25&
                    If i8253_counter(i) <= 0& Then
                        i8253_counter(i) = 0&
                        i8253_out(i) = 1&
                        If i = 0& Then i8253_timerCallback0
                    End If

                Case 2&
                    i8253_counter(i) = i8253_counter(i) - 25&
                    If i8253_counter(i) <= 0& Then
                        i8253_out(i) = i8253_out(i) Xor 1&
                        If i = 0& Then i8253_timerCallback0
                        i8253_counter(i) = i8253_counter(i) + i8253_reload(i)
                    End If

                Case 3&
                    i8253_counter(i) = i8253_counter(i) - 50&
                    If i8253_counter(i) <= 0& Then
                        i8253_out(i) = i8253_out(i) Xor 1&
                        If i8253_out(i) = 0& Then
                            If i = 0& Then i8253_timerCallback0
                        End If

                        If i = 2& Then
                            If i8253_reload(i) < 50& Then
                                gateVal = 0&
                            Else
                                gateVal = i8253_out(i)
                            End If
                            pcspeaker_setGateState 0&, PC_SPEAKER_GATE_TIMER2, gateVal
                        End If

                        i8253_counter(i) = i8253_counter(i) + i8253_reload(i)
                    End If
            End Select
        End If
    Next i
End Sub

Public Sub i8253_init(ByRef machineRef As MACHINE_t, ByVal i8259Slot As Long)
    Dim i As Long

    For i = 0& To 2&
        i8253_chandata(i) = 0&
        i8253_accessmode(i) = 0&
        i8253_bytetoggle(i) = 0&
        i8253_effectivedata(i) = 0&
        i8253_chanfreq(i) = 0!

        i8253_active(i) = 0&
        i8253_counter(i) = 0&
        i8253_reload(i) = 0&
        i8253_mode(i) = 0&
        i8253_dataflipflop(i) = 0&
        i8253_bcd(i) = 0&
        i8253_rlmode(i) = 0&
        i8253_latch(i) = -1&
        i8253_out(i) = 0&
    Next i

    i8253_i8259Slot = i8259Slot

    timing_addTimer TIMER_CB_I8253_TICK, 0&, 48000#, TIMING_ENABLED
    ports_cbRegister &H40&, 4&, PORTS_CB_I8253, PORTS_CB_NONE, PORTS_CB_I8253, PORTS_CB_NONE, 0&
End Sub

