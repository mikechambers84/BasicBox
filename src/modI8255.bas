Attribute VB_Name = "modI8255"
Option Explicit

Private i8255_sw2 As Byte
Private i8255_portA As Byte
Private i8255_portB As Byte
Private i8255_portC As Byte

Public Function i8255_readport(ByVal dummy As Long, ByVal portnum As Integer) As Byte
    Dim p As Long

    p = (portnum And 7&)

    Select Case p
        Case 0&
            i8255_readport = machine.KeyState.scancode

        Case 1&
            i8255_readport = (i8255_portB And &H3F&)

        Case 2&
            If (i8255_portB And &H8&) <> 0& Then
                i8255_readport = (i8255_sw2 And &HF0&) \ &H10&
            Else
                i8255_readport = (i8255_sw2 And &HF&)
            End If

        Case Else
            i8255_readport = &HFF&
    End Select
End Function

Public Sub i8255_writeport(ByVal dummy As Long, ByVal portnum As Integer, ByVal value As Byte)
    Dim p As Long

    p = (portnum And 7&)

    Select Case p
        Case 0&
            machine.KeyState.scancode = &HAA&

        Case 1&
            If (value And &H1&) <> 0& Then
                pcspeaker_selectGate 0&, PC_SPEAKER_USE_TIMER2
            Else
                pcspeaker_selectGate 0&, PC_SPEAKER_USE_DIRECT
            End If

            pcspeaker_setGateState 0&, PC_SPEAKER_GATE_DIRECT, CByte((value And &H2&) \ 2&)

            If ((value And &H40&) <> 0&) And ((i8255_portB And &H40&) = 0&) Then
                machine.KeyState.scancode = &HAA&
            End If

            i8255_portB = (value And &HEF&) Or (i8255_portB And &H10&)
    End Select
End Sub

Public Sub i8255_refreshToggle(ByVal dummy As Long)
    i8255_portB = i8255_portB Xor &H10&
End Sub

Public Sub i8255_init(ByRef machineRef As MACHINE_t)
    i8255_sw2 = 0&
    i8255_portA = 0&
    i8255_portB = 0&
    i8255_portC = 0&

    i8255_sw2 = &H46&

    ports_cbRegister &H60&, 6&, PORTS_CB_I8255, PORTS_CB_NONE, PORTS_CB_I8255, PORTS_CB_NONE, 0&
    timing_addTimer TIMER_CB_I8255_REFRESH, 0&, 66667#, TIMING_ENABLED
End Sub
