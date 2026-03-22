Attribute VB_Name = "modPorts"
Option Explicit

Private Type PORTMAP_t
    start As Long
    size As Long
    readcb As Long
    readcbW As Long
    readcbL As Long
    writecb As Long
    writecbW As Long
    writecbL As Long
    udata As Long
    used As Integer
End Type

Public Const PORTS_CB_NONE As Long = 0&
Public Const PORTS_CB_I8259 As Long = 1&
Public Const PORTS_CB_I8042 As Long = 2&
Public Const PORTS_CB_I8253 As Long = 3&
Public Const PORTS_CB_I8237_PORT As Long = 4&
Public Const PORTS_CB_I8237_PAGE As Long = 5&
Public Const PORTS_CB_I8255 As Long = 6&
Public Const PORTS_CB_RTC As Long = 7&
Public Const PORTS_CB_CMOSRTC As Long = 8&
Public Const PORTS_CB_RABBIT As Long = 9&
Public Const PORTS_CB_OPTI495 As Long = 10&
Public Const PORTS_CB_OPTI5X7 As Long = 11&
Public Const PORTS_CB_UART As Long = 12&
Public Const PORTS_CB_NE2000_REG As Long = 13&
Public Const PORTS_CB_NE2000_ASIC As Long = 14&
Public Const PORTS_CB_NE2000_RESET As Long = 15&
Public Const PORTS_CB_ATA_PORT As Long = 16&
Public Const PORTS_CB_ATA_DATA As Long = 17&
Public Const PORTS_CB_FDC As Long = 18&
Public Const PORTS_CB_VGA As Long = 21&
Public Const PORTS_CB_BLASTER As Long = 22&
Public Const PORTS_CB_OPL2 As Long = 23&
Public Const PORTS_CB_I8237_PAGEH As Long = 24&
Public Const PORTS_CB_BUSLOGIC As Long = 25&
Public Const PORTS_CB_ET4000 As Long = 26&

Private ports(0& To 63&) As PORTMAP_t
Private lastportmap As Long

Private Function getportmap(ByVal addr32 As Long) As Long
    Dim i As Long

    For i = lastportmap To 0& Step -1&
        If ports(i).used <> 0& Then
            If (addr32 >= ports(i).start) And (addr32 < (ports(i).start + ports(i).size)) Then
                getportmap = i
                Exit Function
            End If
        End If
    Next i

    getportmap = -1&
End Function

Public Sub port_write(ByRef cpu As CPU_t, ByVal portnum As Long, ByVal value As Byte)
    Dim map As Long

    map = getportmap(portnum And &HFFFF&)

    If (portnum And &HFFFF&) = &H92& Then
        If (value And &H2&) <> 0& Then
            cpu.a20_gate = 1&
        Else
            cpu.a20_gate = 0&
        End If
        Exit Sub
    End If

    If map <> -1& Then
        If ports(map).writecb <> PORTS_CB_NONE Then
            Ports_DispatchWriteB ports(map).writecb, ports(map).udata, CInt(portnum And &HFFFF&), value
            Exit Sub
        End If
    End If
End Sub

Public Sub port_writew(ByRef cpu As CPU_t, ByVal portnum As Long, ByVal value As Long)
    Dim map As Long

    map = getportmap(portnum And &HFFFF&)

    If map <> -1& Then
        If ports(map).writecbW <> PORTS_CB_NONE Then
            Ports_DispatchWriteW ports(map).writecbW, ports(map).udata, CInt(portnum And &HFFFF&), (value And &HFFFF&)
            Exit Sub
        End If
    End If

    port_write cpu, portnum, CByte(value And &HFF&)
    port_write cpu, portnum + 1&, CByte((value \ &H100&) And &HFF&)
End Sub

Public Sub port_writel(ByRef cpu As CPU_t, ByVal portnum As Long, ByVal value As Long)
    Dim map As Long

    map = getportmap(portnum And &HFFFF&)

    If map <> -1& Then
        If ports(map).writecbL <> PORTS_CB_NONE Then
            Ports_DispatchWriteL ports(map).writecbL, ports(map).udata, CInt(portnum And &HFFFF&), value
            Exit Sub
        End If
    End If

    port_write cpu, portnum, CByte(value And &HFF&)
    port_write cpu, portnum + 1&, CByte((U32Shr(value, 8&)) And &HFF&)
    port_write cpu, portnum + 2&, CByte((U32Shr(value, 16&)) And &HFF&)
    port_write cpu, portnum + 3&, CByte((U32Shr(value, 24&)) And &HFF&)
End Sub

Public Function port_read(ByRef cpu As CPU_t, ByVal portnum As Long) As Byte
    Dim map As Long

    map = getportmap(portnum And &HFFFF&)

    If map <> -1& Then
        If ports(map).readcb <> PORTS_CB_NONE Then
            port_read = Ports_DispatchReadB(ports(map).readcb, ports(map).udata, CInt(portnum And &HFFFF&))
            Exit Function
        End If
    End If

    port_read = &HFF&
End Function

Public Function port_readw(ByRef cpu As CPU_t, ByVal portnum As Long) As Long
    Dim map As Long
    Dim ret As Long

    map = getportmap(portnum And &HFFFF&)

    If map <> -1& Then
        If ports(map).readcbW <> PORTS_CB_NONE Then
            port_readw = Ports_DispatchReadW(ports(map).readcbW, ports(map).udata, CInt(portnum And &HFFFF&)) And &HFFFF&
            Exit Function
        End If
    End If

    ret = CLng(port_read(cpu, portnum))
    ret = ret Or (CLng(port_read(cpu, portnum + 1&)) * &H100&)
    port_readw = ret And &HFFFF&
End Function

Public Function port_readl(ByRef cpu As CPU_t, ByVal portnum As Long) As Long
    Dim map As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long

    map = getportmap(portnum And &HFFFF&)

    If map <> -1& Then
        If ports(map).readcbL <> PORTS_CB_NONE Then
            port_readl = Ports_DispatchReadL(ports(map).readcbL, ports(map).udata, CInt(portnum And &HFFFF&))
            Exit Function
        End If
    End If

    b0 = CLng(port_read(cpu, portnum))
    b1 = CLng(port_read(cpu, portnum + 1&))
    b2 = CLng(port_read(cpu, portnum + 2&))
    b3 = CLng(port_read(cpu, portnum + 3&))

    port_readl = U32FromDouble(CDbl(b0) + CDbl(b1) * 256# + CDbl(b2) * 65536# + CDbl(b3) * 16777216#)
End Function

Public Sub ports_cbRegister(ByVal start As Long, ByVal count As Long, ByVal readb As Long, ByVal readw As Long, ByVal writeb As Long, ByVal writew As Long, ByVal udata As Long)
    Dim i As Long

    For i = 0& To 63&
        If ports(i).used = 0& Then Exit For
    Next i

    If i = 64& Then
        debug_log DEBUG_ERROR, "[PORTS] Out of port map structs!"
        End
    End If

    ports(i).readcb = readb
    ports(i).writecb = writeb
    ports(i).readcbW = readw
    ports(i).writecbW = writew
    ports(i).readcbL = PORTS_CB_NONE
    ports(i).writecbL = PORTS_CB_NONE
    ports(i).start = start And &HFFFF&
    ports(i).size = count And &HFFFF&
    ports(i).udata = udata
    ports(i).used = 1&

    lastportmap = i
End Sub

Public Sub ports_init()
    Dim i As Long

    For i = 0& To 63&
        ports(i).readcb = PORTS_CB_NONE
        ports(i).writecb = PORTS_CB_NONE
        ports(i).readcbW = PORTS_CB_NONE
        ports(i).writecbW = PORTS_CB_NONE
        ports(i).readcbL = PORTS_CB_NONE
        ports(i).writecbL = PORTS_CB_NONE
        ports(i).used = 0&
        ports(i).start = 0&
        ports(i).size = 0&
        ports(i).udata = 0&
    Next i

    lastportmap = -1&
End Sub

Private Function Ports_DispatchReadB(ByVal cbid As Long, ByVal udata As Long, ByVal portnum As Integer) As Byte
    Select Case cbid
        Case PORTS_CB_I8259
            Ports_DispatchReadB = i8259_read(udata, portnum)
        Case PORTS_CB_I8042
            Ports_DispatchReadB = i8042_readport(udata, portnum)
        Case PORTS_CB_I8253
            Ports_DispatchReadB = i8253_read(udata, portnum)
        Case PORTS_CB_I8237_PORT
            Ports_DispatchReadB = i8237_readport(udata, portnum)
        Case PORTS_CB_I8237_PAGE
            Ports_DispatchReadB = i8237_readpage(udata, portnum)
        Case PORTS_CB_I8237_PAGEH
            Ports_DispatchReadB = i8237_readpageh(udata, portnum)
        Case PORTS_CB_I8255
            Ports_DispatchReadB = i8255_readport(udata, portnum)
        Case PORTS_CB_RTC
            Ports_DispatchReadB = rtc_read(udata, portnum)
        Case PORTS_CB_CMOSRTC
            Ports_DispatchReadB = cmosrtc_read(udata, portnum)
        Case PORTS_CB_RABBIT
            Ports_DispatchReadB = rabbit_read(udata, portnum)
        Case PORTS_CB_OPTI495
            Ports_DispatchReadB = opti495_read(udata, portnum)
        Case PORTS_CB_OPTI5X7
            Ports_DispatchReadB = opti5x7_read(udata, portnum)
        Case PORTS_CB_UART
            Ports_DispatchReadB = uart_readport(udata, portnum)
        Case PORTS_CB_NE2000_REG
            Ports_DispatchReadB = ne2000_read(udata, portnum)
        Case PORTS_CB_NE2000_ASIC
            Ports_DispatchReadB = ne2000_asic_read_b(udata, portnum)
        Case PORTS_CB_NE2000_RESET
            Ports_DispatchReadB = ne2000_reset_read(udata, portnum)
        Case PORTS_CB_ATA_PORT
            Ports_DispatchReadB = ata_read_port(udata, portnum)
        Case PORTS_CB_FDC
            Ports_DispatchReadB = fdc_read(udata, portnum)
        Case PORTS_CB_VGA
            Ports_DispatchReadB = vga_readport(udata, portnum)
        Case PORTS_CB_ET4000
            Ports_DispatchReadB = et4000_readport(udata, portnum)
        Case PORTS_CB_BLASTER
            Ports_DispatchReadB = blaster_read(udata, portnum)
        Case PORTS_CB_OPL2
            Ports_DispatchReadB = opl2_read(udata, portnum)
        Case PORTS_CB_BUSLOGIC
            Ports_DispatchReadB = buslogic_readport(udata, portnum)
        Case Else
            Ports_DispatchReadB = &HFF&
    End Select
End Function

Private Function Ports_DispatchReadW(ByVal cbid As Long, ByVal udata As Long, ByVal portnum As Integer) As Long
    Select Case cbid
        Case PORTS_CB_NE2000_ASIC
            Ports_DispatchReadW = ne2000_asic_read_w(udata, portnum) And &HFFFF&
        Case PORTS_CB_ATA_DATA
            Ports_DispatchReadW = ata_read_data(udata, portnum) And &HFFFF&
        Case Else
            Ports_DispatchReadW = 0&
    End Select
End Function

Private Function Ports_DispatchReadL(ByVal cbid As Long, ByVal udata As Long, ByVal portnum As Integer) As Long
    Ports_DispatchReadL = 0&
End Function

Private Sub Ports_DispatchWriteB(ByVal cbid As Long, ByVal udata As Long, ByVal portnum As Integer, ByVal value As Byte)
    Select Case cbid
        Case PORTS_CB_I8259
            i8259_write udata, portnum, value
        Case PORTS_CB_I8042
            i8042_writeport udata, portnum, value
        Case PORTS_CB_I8253
            i8253_write udata, portnum, value
        Case PORTS_CB_I8237_PORT
            i8237_writeport udata, portnum, value
        Case PORTS_CB_I8237_PAGE
            i8237_writepage udata, portnum, value
        Case PORTS_CB_I8237_PAGEH
            i8237_writepageh udata, portnum, value
        Case PORTS_CB_I8255
            i8255_writeport udata, portnum, value
        Case PORTS_CB_RTC
            rtc_write udata, portnum, value
        Case PORTS_CB_CMOSRTC
            cmosrtc_write udata, portnum, value
        Case PORTS_CB_RABBIT
            rabbit_write udata, portnum, value
        Case PORTS_CB_OPTI495
            opti495_write udata, portnum, value
        Case PORTS_CB_OPTI5X7
            opti5x7_write udata, portnum, value
        Case PORTS_CB_UART
            uart_writeport udata, portnum, value
        Case PORTS_CB_NE2000_REG
            ne2000_write udata, portnum, value
        Case PORTS_CB_NE2000_ASIC
            ne2000_asic_write_b udata, portnum, value
        Case PORTS_CB_NE2000_RESET
            ne2000_reset_write udata, portnum, value
        Case PORTS_CB_ATA_PORT
            ata_write_port udata, portnum, value
        Case PORTS_CB_FDC
            fdc_write udata, portnum, value
        Case PORTS_CB_VGA
            vga_writeport udata, portnum, value
        Case PORTS_CB_ET4000
            et4000_writeport udata, portnum, value
        Case PORTS_CB_BLASTER
            blaster_write udata, portnum, value
        Case PORTS_CB_OPL2
            opl2_write udata, portnum, value
        Case PORTS_CB_BUSLOGIC
            buslogic_writeport udata, portnum, value
    End Select
End Sub

Private Sub Ports_DispatchWriteW(ByVal cbid As Long, ByVal udata As Long, ByVal portnum As Integer, ByVal value As Long)
    Select Case cbid
        Case PORTS_CB_NE2000_ASIC
            ne2000_asic_write_w udata, portnum, value
        Case PORTS_CB_ATA_DATA
            ata_write_data udata, portnum, value
        Case Else
            ' No-op.
    End Select
End Sub

Private Sub Ports_DispatchWriteL(ByVal cbid As Long, ByVal udata As Long, ByVal portnum As Integer, ByVal value As Long)
    ' No 32-bit port callbacks currently used.
End Sub


