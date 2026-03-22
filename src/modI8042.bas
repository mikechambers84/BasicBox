Attribute VB_Name = "modI8042"
Option Explicit

Public Type i8042_t
    data_buffer(0& To 7&) As Byte
    buflen As Byte
    status As Byte
    config As Byte
    command As Byte
    reset_requested As Byte
    self_test_done As Byte
    keyboard_enabled As Byte
    i8259 As Long
End Type

Public kbc As i8042_t

Public Sub i8042_buffer_key_data(ByRef data() As Byte, ByVal datalen As Byte, ByVal doirq As Byte)
    Dim i As Long

    kbc.keyboard_enabled = 1&

    If (kbc.keyboard_enabled = 0&) Or (datalen = 0&) Or (datalen > (UBound(kbc.data_buffer) + 1&)) Then
        Exit Sub
    End If

    For i = 0& To datalen - 1&
        kbc.data_buffer(i) = data(i)
    Next i

    kbc.buflen = datalen

    If doirq <> 0& Then
        i8259_doirq kbc.i8259, 1&
    End If
End Sub

Public Function i8042_read_0x60() As Byte
    Static ret As Byte
    Dim i As Long

    If kbc.buflen > 0& Then
        ret = kbc.data_buffer(0&)

        For i = 0& To UBound(kbc.data_buffer) - 1&
            kbc.data_buffer(i) = kbc.data_buffer(i + 1&)
        Next i

        kbc.buflen = kbc.buflen - 1&
    End If

    If kbc.buflen > 0& Then
        i8259_doirq kbc.i8259, 1&
    End If

    i8042_read_0x60 = ret
End Function

Public Sub i8042_write_0x60(ByVal value As Byte)
    Dim outBytes(0& To 0&) As Byte

    kbc.status = kbc.status Or &H2&

    Select Case kbc.command
        Case &H0&
            Select Case value
                Case &HFF&
                    outBytes(0&) = &HFA&
                    i8042_buffer_key_data outBytes, 1&, 1&

                    outBytes(0&) = &HAA&
                    i8042_buffer_key_data outBytes, 1&, 1&

                Case Else
                    outBytes(0&) = &HFA&
                    i8042_buffer_key_data outBytes, 1&, 1&
            End Select

        Case &H60&
            kbc.config = value

        Case &HD1&
            If (value And &H2&) <> 0& Then
                machine.CPU.a20_gate = 1&
            Else
                machine.CPU.a20_gate = 0&
            End If

        Case &HFF&
            outBytes(0&) = &HFA&
            kbc.status = kbc.status And (Not &H4&)
            i8042_buffer_key_data outBytes, 1&, 0&

        Case Else
            ' Unhandled command.
    End Select

    kbc.command = 0&
End Sub

Public Function i8042_read_0x64() As Byte
    Dim status As Byte

    status = (kbc.status Or &H10&)
    If kbc.buflen > 0& Then
        status = status Or &H1&
    End If

    kbc.status = kbc.status And (Not &H2&)

    i8042_read_0x64 = status
End Function

Public Sub i8042_write_0x64(ByVal value As Byte)
    Dim outBytes(0& To 0&) As Byte

    kbc.status = kbc.status Or &H2&

    Select Case value
        Case &H20&
            outBytes(0&) = kbc.config
            i8042_buffer_key_data outBytes, 1&, 0&

        Case &H60&
            kbc.command = &H60&

        Case &HA1&
            outBytes(0&) = Asc("N")
            i8042_buffer_key_data outBytes, 1&, 0&

        Case &HAA&
            kbc.status = kbc.status Or &H4&
            outBytes(0&) = &H55&
            i8042_buffer_key_data outBytes, 1&, 0&

        Case &HAB&
            outBytes(0&) = &H0&
            i8042_buffer_key_data outBytes, 1&, 0&

        Case &HAD&
            kbc.keyboard_enabled = 0&

        Case &HAE&
            kbc.keyboard_enabled = 1&
            outBytes(0&) = &HFA&
            i8042_buffer_key_data outBytes, 1&, 0&

        Case &HC0&
            outBytes(0&) = (&H80& Or &H20&)
            i8042_buffer_key_data outBytes, 1&, 0&

        Case &HD1&
            kbc.command = value

        Case Else
            kbc.command = 0&
    End Select
End Sub

Public Function i8042_readport(ByVal dummy As Long, ByVal portnum As Integer) As Byte
    Select Case (portnum And &HFFFF&)
        Case &H60&
            i8042_readport = i8042_read_0x60()
        Case &H64&
            i8042_readport = i8042_read_0x64()
        Case Else
            i8042_readport = &HFF&
    End Select
End Function

Public Sub i8042_writeport(ByVal dummy As Long, ByVal portnum As Integer, ByVal value As Byte)
    Select Case (portnum And &HFFFF&)
        Case &H60&
            i8042_write_0x60 value
        Case &H64&
            i8042_write_0x64 value
        Case Else
            ' Invalid write for controller.
    End Select
End Sub

Public Sub i8042_init(ByRef cpu As CPU_t, ByVal i8259Slot As Long)
    Dim i As Long

    kbc.status = &H0&
    kbc.buflen = 0&
    kbc.command = 0&
    kbc.reset_requested = 0&
    kbc.self_test_done = 0&
    kbc.i8259 = i8259Slot

    For i = 0& To UBound(kbc.data_buffer)
        kbc.data_buffer(i) = 0&
    Next i

    cpu.a20_gate = 0&
    machine.CPU.a20_gate = 0&

    ports_cbRegister &H60&, 1&, PORTS_CB_I8042, PORTS_CB_NONE, PORTS_CB_I8042, PORTS_CB_NONE, 0&
    ports_cbRegister &H64&, 1&, PORTS_CB_I8042, PORTS_CB_NONE, PORTS_CB_I8042, PORTS_CB_NONE, 0&

    debug_log DEBUG_INFO, "[KBC] Initialized AT-style keyboard controller"
End Sub
