Attribute VB_Name = "modBlaster"
Option Explicit

Private Type BLASTER_t
    i8259Slot As Long
    dspenable As Byte
    sample As Integer
    readbuf(0& To 15&) As Byte
    readlen As Byte
    readready As Byte
    writebuf As Byte
    timeconst As Byte
    samplerate As Double
    timer As Long
    dmalen As Long
    dmachan As Byte
    irq As Byte
    lastcmd As Byte
    writehilo As Byte
    dmacount As Long
    autoinit As Byte
    testreg As Byte
    silencedsp As Byte
    dorecord As Byte
    activedma As Byte
    dmaLastTransferred As Long
    dmaBlockDone As Byte
End Type

Private blaster_dev As BLASTER_t

Private Function blaster_cmdE2Table(ByVal idx As Long) As Integer
    Select Case idx
        Case 0&: blaster_cmdE2Table = &H1&
        Case 1&: blaster_cmdE2Table = -&H2&
        Case 2&: blaster_cmdE2Table = -&H4&
        Case 3&: blaster_cmdE2Table = &H8&
        Case 4&: blaster_cmdE2Table = -&H10&
        Case 5&: blaster_cmdE2Table = &H20&
        Case 6&: blaster_cmdE2Table = &H40&
        Case 7&: blaster_cmdE2Table = -&H80&
        Case Else: blaster_cmdE2Table = -106&
    End Select
End Function

Private Sub blaster_putreadbuf(ByVal value As Byte)
    If blaster_dev.readlen = 16& Then Exit Sub
    blaster_dev.readbuf(blaster_dev.readlen) = value
    blaster_dev.readlen = blaster_dev.readlen + 1&
End Sub

Private Function blaster_getreadbuf() As Byte
    Dim ret As Byte
    Dim i As Long

    ret = blaster_dev.readbuf(0&)
    If blaster_dev.readlen > 0& Then
        blaster_dev.readlen = blaster_dev.readlen - 1&
    End If

    For i = 0& To 14&
        blaster_dev.readbuf(i) = blaster_dev.readbuf(i + 1&)
    Next i

    blaster_getreadbuf = ret
End Function

Private Sub blaster_reset()
    blaster_dev.dspenable = 0&
    blaster_dev.sample = 0&
    blaster_dev.readlen = 0&
    blaster_dev.dmaLastTransferred = 0&
    blaster_dev.dmaBlockDone = 0&
    blaster_putreadbuf &HAA&
End Sub

Private Sub blaster_writecmd(ByVal value As Byte)
    Dim val16 As Integer
    Dim i As Long

    Select Case blaster_dev.lastcmd
        Case &H9&, &H10&
            blaster_dev.sample = CInt((CLng(value) - 128&) * 256&)
            blaster_dev.lastcmd = 0&
            Exit Sub

        Case &H14&, &H24&
            If blaster_dev.writehilo = 0& Then
                blaster_dev.dmalen = value
                blaster_dev.writehilo = 1&
            Else
                blaster_dev.dmalen = blaster_dev.dmalen Or (CLng(value) * &H100&)
                blaster_dev.dmalen = blaster_dev.dmalen + 1&
                blaster_dev.lastcmd = 0&
                blaster_dev.dmacount = 0&
                blaster_dev.silencedsp = 0&
                blaster_dev.autoinit = 0&
                If blaster_dev.lastcmd = &H24& Then
                    blaster_dev.dorecord = 1&
                Else
                    blaster_dev.dorecord = 0&
                End If
                blaster_dev.activedma = 1&
                timing_timerEnable blaster_dev.timer
            End If
            Exit Sub

        Case &H40&
            blaster_dev.timeconst = value
            blaster_dev.samplerate = 1000000# / (256# - CDbl(value))
            timing_updateIntervalFreq blaster_dev.timer, blaster_dev.samplerate
            blaster_dev.lastcmd = 0&
            Exit Sub

        Case &H48&
            If blaster_dev.writehilo = 0& Then
                blaster_dev.dmalen = value
                blaster_dev.writehilo = 1&
            Else
                blaster_dev.dmalen = blaster_dev.dmalen Or (CLng(value) * &H100&)
                blaster_dev.dmalen = blaster_dev.dmalen + 1&
                blaster_dev.lastcmd = 0&
            End If
            Exit Sub

        Case &H80&
            If blaster_dev.writehilo = 0& Then
                blaster_dev.dmalen = value
                blaster_dev.writehilo = 1&
            Else
                blaster_dev.dmalen = blaster_dev.dmalen Or (CLng(value) * &H100&)
                blaster_dev.dmalen = blaster_dev.dmalen + 1&
                blaster_dev.lastcmd = 0&
                blaster_dev.dmacount = 0&
                blaster_dev.silencedsp = 1&
                blaster_dev.autoinit = 0&
                timing_timerEnable blaster_dev.timer
            End If
            Exit Sub

        Case &HE0&
            blaster_putreadbuf CByte((Not value) And &HFF&)
            blaster_dev.lastcmd = 0&
            Exit Sub

        Case &HE2&
            val16 = &HAA&
            For i = 0& To 7&
                If ((value \ U32Shl(1&, i)) And 1&) <> 0& Then
                    val16 = val16 + blaster_cmdE2Table(i)
                End If
            Next i
            val16 = val16 + blaster_cmdE2Table(8&)
            i8237_write blaster_dev.dmachan, CByte(val16 And &HFF&)
            blaster_dev.lastcmd = 0&
            Exit Sub

        Case &HE4&
            blaster_dev.testreg = value
            blaster_dev.lastcmd = 0&
            Exit Sub
    End Select

    Select Case value
        Case &H10&
            ' direct DAC

        Case &H14&, &H24&
            blaster_dev.writehilo = 0&

        Case &H1C&, &H2C&
            blaster_dev.dmacount = 0&
            blaster_dev.silencedsp = 0&
            blaster_dev.autoinit = 1&
            If value = &H2C& Then
                blaster_dev.dorecord = 1&
            Else
                blaster_dev.dorecord = 0&
            End If
            blaster_dev.activedma = 1&
            timing_timerEnable blaster_dev.timer

        Case &H20&
            blaster_putreadbuf 128&

        Case &H40&
            ' set time constant command prefix

        Case &H48&
            blaster_dev.writehilo = 0&

        Case &H80&
            blaster_dev.writehilo = 0&

        Case &HD0&
            blaster_dev.activedma = 0&
            timing_timerDisable blaster_dev.timer

        Case &HD1&
            blaster_dev.dspenable = 1&

        Case &HD3&
            blaster_dev.dspenable = 0&

        Case &HD4&
            blaster_dev.activedma = 1&
            timing_timerEnable blaster_dev.timer

        Case &HDA&
            blaster_dev.activedma = 0&
            blaster_dev.autoinit = 0&

        Case &HE0&
            ' DSP identification prefix

        Case &HE1&
            blaster_putreadbuf 2&
            blaster_putreadbuf 1&

        Case &HE2&
            ' DMA identification write prefix

        Case &HE4&
            ' write test register prefix

        Case &HE8&
            blaster_putreadbuf blaster_dev.testreg

        Case &HF2&
            i8259_doirq blaster_dev.i8259Slot, blaster_dev.irq

        Case &HF8&
            blaster_putreadbuf 0&

        Case Else
            debug_log DEBUG_ERROR, "[BLASTER] Unrecognized command: 0x" & Right$("00" & Hex$(value And &HFF&), 2&)
    End Select

    blaster_dev.lastcmd = value
End Sub

Public Sub blaster_write(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    Dim a As Long

    a = (addr And &HF&)

    Select Case a
        Case &H6&
            If value = 0& Then
                blaster_reset
            End If
        Case &HC&
            blaster_writecmd value
    End Select
End Sub

Public Function blaster_read(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim ret As Byte
    Dim a As Long

    ret = &HFF&
    a = (addr And &HF&)

    Select Case a
        Case &HA&
            blaster_read = blaster_getreadbuf()
            Exit Function
        Case &HC&
            blaster_read = &H0&
            Exit Function
        Case &HE&
            If blaster_dev.readlen > 0& Then
                blaster_read = &H80&
            Else
                blaster_read = &H0&
            End If
            Exit Function
    End Select

    blaster_read = ret
End Function

Public Function blaster_dmaTransferCallback(ByVal callbackData As Long, ByVal channel As Long, ByVal dmaPos As Long, ByVal dmaLen As Long) As Long
    Dim dataBuffer(0& To 0&) As Byte
    Dim transferred As Long

    blaster_dev.dmaLastTransferred = 0&
    blaster_dev.dmaBlockDone = 0&

    If blaster_dev.activedma = 0& Then
        blaster_dmaTransferCallback = dmaPos
        Exit Function
    End If

    If dmaPos >= dmaLen Then
        blaster_dev.dmaBlockDone = 1&
        blaster_dmaTransferCallback = dmaPos
        Exit Function
    End If

    If blaster_dev.dorecord = 0& Then
        dataBuffer(0&) = &H80&
        transferred = i8237_dma_readMemory(channel, dataBuffer, dmaPos, 1&)
        If transferred > 0& Then
            blaster_dev.sample = CInt((CLng(dataBuffer(0&)) - 128&) * 256&)
        End If
    Else
        dataBuffer(0&) = 128&
        transferred = i8237_dma_writeMemory(channel, dataBuffer, dmaPos, 1&)
    End If

    blaster_dev.dmaLastTransferred = transferred
    If transferred > 0& Then
        blaster_dev.dmacount = blaster_dev.dmacount + transferred
    End If
    If blaster_dev.dmacount >= blaster_dev.dmalen Then
        blaster_dev.dmaBlockDone = 1&
    End If

    blaster_dmaTransferCallback = (dmaPos + transferred)
End Function

Public Sub blaster_generateSample(ByVal dummy As Long)
    If blaster_dev.activedma = 0& Then
        If blaster_dev.dspenable = 0& Then
            blaster_dev.sample = 0&
        End If
        Exit Sub
    End If

    blaster_dev.dmaLastTransferred = 0&
    blaster_dev.dmaBlockDone = 0&

    If blaster_dev.silencedsp = 0& Then
        i8237_dma_holdDreq blaster_dev.dmachan
        i8237_dma_releaseDreq blaster_dev.dmachan
        If (blaster_dev.dorecord = 0&) And (blaster_dev.dmaLastTransferred <= 0&) Then
            blaster_dev.sample = 0&
        End If
    Else
        blaster_dev.sample = 0&
        blaster_dev.dmacount = blaster_dev.dmacount + 1&
        If blaster_dev.dmacount >= blaster_dev.dmalen Then
            blaster_dev.dmaBlockDone = 1&
        End If
    End If

    If blaster_dev.dmaBlockDone <> 0& Then
        blaster_dev.dmacount = 0&
        i8259_doirq blaster_dev.i8259Slot, blaster_dev.irq
        If blaster_dev.autoinit = 0& Then
            blaster_dev.activedma = 0&
            timing_timerDisable blaster_dev.timer
        End If
    End If

    If blaster_dev.dspenable = 0& Then
        blaster_dev.sample = 0&
    End If
End Sub

Public Function blaster_getSample(ByVal dummy As Long) As Integer
    blaster_getSample = blaster_dev.sample
End Function

Public Sub blaster_init(ByRef machineRef As MACHINE_t)
    Dim zeroBlaster As BLASTER_t
    Const BLASTER_BASE As Long = &H220&
    Const BLASTER_DMA As Byte = 1&
    Const BLASTER_IRQ As Byte = 5&

    debug_log DEBUG_INFO, "[BLASTER] Initializing Sound Blaster 2.0 at base port 0x" & Right$("000" & Hex$(BLASTER_BASE And &HFFFF&), 3&) & ", IRQ " & CStr(BLASTER_IRQ) & ", DMA " & CStr(BLASTER_DMA)

    blaster_dev = zeroBlaster
    blaster_dev.i8259Slot = machineRef.i8259
    blaster_dev.dmachan = BLASTER_DMA
    blaster_dev.irq = BLASTER_IRQ
    i8237_dma_registerChannel blaster_dev.dmachan, DMA_CB_BLASTER, 0&

    ports_cbRegister BLASTER_BASE, 16&, PORTS_CB_BLASTER, PORTS_CB_NONE, PORTS_CB_BLASTER, PORTS_CB_NONE, 0&

    blaster_dev.timer = timing_addTimer(TIMER_CB_BLASTER_DMA, 0&, 22050#, TIMING_DISABLED)
End Sub
