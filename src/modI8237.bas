Attribute VB_Name = "modI8237"
Option Explicit

Private Const DMA_CMD_MEMORY_TO_MEMORY As Long = &H1&
Private Const DMA_CMD_FIXED_ADDRESS As Long = &H2&
Private Const DMA_CMD_BLOCK_CONTROLLER As Long = &H4&
Private Const DMA_CMD_COMPRESSED_TIME As Long = &H8&
Private Const DMA_CMD_CYCLIC_PRIORITY As Long = &H10&
Private Const DMA_CMD_EXTENDED_WRITE As Long = &H20&
Private Const DMA_CMD_LOW_DREQ As Long = &H40&
Private Const DMA_CMD_LOW_DACK As Long = &H80&
Private Const DMA_CMD_NOT_SUPPORTED As Long = DMA_CMD_MEMORY_TO_MEMORY Or DMA_CMD_FIXED_ADDRESS Or DMA_CMD_COMPRESSED_TIME Or DMA_CMD_CYCLIC_PRIORITY Or DMA_CMD_EXTENDED_WRITE Or DMA_CMD_LOW_DREQ Or DMA_CMD_LOW_DACK

Public Const DMA_OP_VERIFY As Long = 0&
Public Const DMA_OP_WRITEMEM As Long = 1&
Public Const DMA_OP_READMEM As Long = 2&
Public Const DMA_OP_ILLEGAL As Long = 3&

Public Const DMA_CB_NONE As Long = 0&
Public Const DMA_CB_FDC As Long = 1&
Public Const DMA_CB_BLASTER As Long = 2&

Private Type I8257_REG_t
    nowAddr As Long
    nowCount As Long
    baseAddr As Long
    baseCount As Long
    mode As Byte
    page As Byte
    pageHigh As Byte
    dack As Byte
    eop As Byte
End Type

Private Type I8257_CTL_t
    base As Long
    pageBase As Long
    pageHighBase As Long
    dshift As Long
    status As Byte
    command As Byte
    mask As Byte
    flipFlop As Byte
    running As Byte
    dmaBhScheduled As Byte
    regs(0& To 3&) As I8257_REG_t
End Type

Private Type DMA_HANDLER_t
    callbackId As Long
    callbackData As Long
End Type

Private dmaCtl(0& To 1&) As I8257_CTL_t
Private dmaHandlers(0& To 7&) As DMA_HANDLER_t
Private dmaPageShadow(0& To 15&) As Byte
Private dmaPageHighShadow(0& To 15&) As Byte

Private Function DMA_U8(ByVal value As Long) As Byte
    DMA_U8 = CByte(value And &HFF&)
End Function

Private Function DMA_U16(ByVal value As Long) As Long
    DMA_U16 = (value And &HFFFF&)
End Function

Private Function DMA_MaskBit(ByVal localChan As Long) As Long
    DMA_MaskBit = U32Shl(1&, localChan)
End Function

Private Function DMA_ControlSpan(ByVal ctlIndex As Long) As Long
    If dmaCtl(ctlIndex).dshift = 0& Then
        DMA_ControlSpan = 8&
    Else
        DMA_ControlSpan = 16&
    End If
End Function

Private Sub DMA_InitController(ByVal ctlIndex As Long, ByVal baseAddr As Long, ByVal pageBase As Long, ByVal pageHighBase As Long, ByVal dshift As Long)
    Dim ch As Long

    dmaCtl(ctlIndex).base = baseAddr
    dmaCtl(ctlIndex).pageBase = pageBase
    dmaCtl(ctlIndex).pageHighBase = pageHighBase
    dmaCtl(ctlIndex).dshift = dshift
    dmaCtl(ctlIndex).status = 0&
    dmaCtl(ctlIndex).command = 0&
    dmaCtl(ctlIndex).mask = &HF&
    dmaCtl(ctlIndex).flipFlop = 0&
    dmaCtl(ctlIndex).running = 0&
    dmaCtl(ctlIndex).dmaBhScheduled = 0&

    For ch = 0& To 3&
        dmaCtl(ctlIndex).regs(ch).nowAddr = 0&
        dmaCtl(ctlIndex).regs(ch).nowCount = 0&
        dmaCtl(ctlIndex).regs(ch).baseAddr = 0&
        dmaCtl(ctlIndex).regs(ch).baseCount = 0&
        dmaCtl(ctlIndex).regs(ch).mode = 0&
        dmaCtl(ctlIndex).regs(ch).page = 0&
        dmaCtl(ctlIndex).regs(ch).pageHigh = 0&
        dmaCtl(ctlIndex).regs(ch).dack = 0&
        dmaCtl(ctlIndex).regs(ch).eop = 0&
    Next ch
End Sub

Private Sub DMA_InitChannel(ByVal ctlIndex As Long, ByVal localChan As Long)
    dmaCtl(ctlIndex).regs(localChan).nowAddr = U32Shl(DMA_U16(dmaCtl(ctlIndex).regs(localChan).baseAddr), dmaCtl(ctlIndex).dshift)
    dmaCtl(ctlIndex).regs(localChan).nowCount = 0&
End Sub

Private Function DMA_GetFlipFlop(ByVal ctlIndex As Long) As Long
    DMA_GetFlipFlop = dmaCtl(ctlIndex).flipFlop
    If dmaCtl(ctlIndex).flipFlop = 0& Then
        dmaCtl(ctlIndex).flipFlop = 1&
    Else
        dmaCtl(ctlIndex).flipFlop = 0&
    End If
End Function

Private Function DMA_PagePortChannel(ByVal addr As Long) As Long
    Select Case (addr And &HFFFF&)
        Case &H81&, &H481&
            DMA_PagePortChannel = 2&
        Case &H82&, &H482&
            DMA_PagePortChannel = 3&
        Case &H83&, &H483&
            DMA_PagePortChannel = 1&
        Case &H87&, &H487&
            DMA_PagePortChannel = 0&
        Case &H89&, &H489&
            DMA_PagePortChannel = 6&
        Case &H8A&, &H48A&
            DMA_PagePortChannel = 7&
        Case &H8B&, &H48B&
            DMA_PagePortChannel = 5&
        Case &H8F&, &H48F&
            DMA_PagePortChannel = 4&
        Case Else
            DMA_PagePortChannel = -1&
    End Select
End Function

Private Function DMA_ChannelController(ByVal channel As Long) As Long
    DMA_ChannelController = ((channel And 7&) \ 4&)
End Function

Private Function DMA_ChannelLocal(ByVal channel As Long) As Long
    DMA_ChannelLocal = (channel And 3&)
End Function

Private Function DMA_TotalLenBytes(ByVal ctlIndex As Long, ByVal localChan As Long) As Long
    DMA_TotalLenBytes = (DMA_U16(dmaCtl(ctlIndex).regs(localChan).baseCount) + 1&)
    If dmaCtl(ctlIndex).dshift <> 0& Then
        DMA_TotalLenBytes = DMA_TotalLenBytes * 2&
    End If
End Function

Private Function DMA_IsVerifyTransfer(ByVal mode As Byte) As Long
    If (mode And &HC&) = 0& Then
        DMA_IsVerifyTransfer = 1&
    Else
        DMA_IsVerifyTransfer = 0&
    End If
End Function

Private Function DMA_ComposeAddr(ByVal pageHigh As Byte, ByVal page As Byte, ByVal addr As Long) As Long
    DMA_ComposeAddr = U32Add(U32Shl((CLng(pageHigh) And &H7F&), 24&), U32Add(U32Shl((CLng(page) And &HFF&), 16&), (addr And &H1FFFF&)))
End Function

Private Function DMA_StepAddress(ByVal baseAddr As Long, ByVal delta As Long, ByVal decrement As Long) As Long
    If decrement <> 0& Then
        DMA_StepAddress = U32Sub(baseAddr, delta)
    Else
        DMA_StepAddress = U32Add(baseAddr, delta)
    End If
End Function

Private Function DMA_DispatchTransfer(ByVal callbackId As Long, ByVal callbackData As Long, ByVal channel As Long, ByVal dmaPos As Long, ByVal dmaLen As Long) As Long
    Select Case callbackId
        Case DMA_CB_FDC
            DMA_DispatchTransfer = fdc_dmaTransferCallback(callbackData, channel, dmaPos, dmaLen)
        Case DMA_CB_BLASTER
            DMA_DispatchTransfer = blaster_dmaTransferCallback(callbackData, channel, dmaPos, dmaLen)
        Case Else
            DMA_DispatchTransfer = dmaPos
    End Select
End Function

Private Sub DMA_CommitPosition(ByVal ctlIndex As Long, ByVal localChan As Long, ByVal newPos As Long)
    Dim totalLen As Long

    totalLen = DMA_TotalLenBytes(ctlIndex, localChan)
    If newPos > totalLen Then newPos = totalLen
    dmaCtl(ctlIndex).regs(localChan).nowCount = newPos
    If newPos >= totalLen Then
        dmaCtl(ctlIndex).status = CByte((dmaCtl(ctlIndex).status Or DMA_MaskBit(localChan)) And &HFF&)
        If (dmaCtl(ctlIndex).regs(localChan).mode And &H10&) <> 0& Then
            DMA_InitChannel ctlIndex, localChan
        End If
    End If
End Sub

Private Function DMA_ChannelRun(ByVal ctlIndex As Long, ByVal localChan As Long) As Long
    Dim channel As Long
    Dim dmaPos As Long
    Dim dmaLen As Long
    Dim newPos As Long

    channel = (localChan Or (ctlIndex * 4&))
    dmaPos = dmaCtl(ctlIndex).regs(localChan).nowCount
    dmaLen = DMA_TotalLenBytes(ctlIndex, localChan)
    newPos = DMA_DispatchTransfer(dmaHandlers(channel).callbackId, dmaHandlers(channel).callbackData, channel, dmaPos, dmaLen)

    If newPos < dmaPos Then newPos = dmaPos
    DMA_CommitPosition ctlIndex, localChan, newPos
    DMA_ChannelRun = newPos
End Function

Private Sub DMA_RunController(ByVal ctlIndex As Long)
    Dim localChan As Long
    Dim maskBit As Long
    Dim rearm As Byte

    If dmaCtl(ctlIndex).running <> 0& Then
        dmaCtl(ctlIndex).dmaBhScheduled = 1&
        Exit Sub
    End If

    dmaCtl(ctlIndex).running = 1&
    rearm = 0&

    For localChan = 0& To 3&
        maskBit = DMA_MaskBit(localChan)
        If ((dmaCtl(ctlIndex).mask And maskBit) = 0&) And ((dmaCtl(ctlIndex).status And (maskBit * 16&)) <> 0&) Then
            Call DMA_ChannelRun(ctlIndex, localChan)
            rearm = 1&
        End If
    Next localChan

    dmaCtl(ctlIndex).running = 0&
    dmaCtl(ctlIndex).dmaBhScheduled = rearm
End Sub

Private Function DMA_ControllerFromPort(ByVal addr As Long, ByRef ctlIndex As Long, ByRef localPort As Long) As Long
    addr = (addr And &HFFFF&)

    If (addr >= dmaCtl(0).base) And (addr < (dmaCtl(0).base + (DMA_ControlSpan(0&) * 2&))) Then
        ctlIndex = 0&
        localPort = (addr - dmaCtl(0).base)
        DMA_ControllerFromPort = 1&
        Exit Function
    End If

    If (addr >= dmaCtl(1).base) And (addr < (dmaCtl(1).base + (DMA_ControlSpan(1&) * 2&))) Then
        ctlIndex = 1&
        localPort = (addr - dmaCtl(1).base)
        DMA_ControllerFromPort = 1&
        Exit Function
    End If

    ctlIndex = -1&
    localPort = 0&
    DMA_ControllerFromPort = 0&
End Function

Private Function DMA_ReadChanPort(ByVal ctlIndex As Long, ByVal localPort As Long) As Byte
    Dim iport As Long
    Dim localChan As Long
    Dim regIndex As Long
    Dim ff As Long
    Dim value32 As Long

    iport = (U32Shr(localPort, dmaCtl(ctlIndex).dshift) And &HF&)
    localChan = U32Shr(iport, 1&)
    regIndex = (iport And 1&)
    ff = DMA_GetFlipFlop(ctlIndex)

    If regIndex <> 0& Then
        value32 = (U32Shl(DMA_U16(dmaCtl(ctlIndex).regs(localChan).baseCount), dmaCtl(ctlIndex).dshift) - dmaCtl(ctlIndex).regs(localChan).nowCount)
    ElseIf (dmaCtl(ctlIndex).regs(localChan).mode And &H20&) <> 0& Then
        value32 = U32Sub(dmaCtl(ctlIndex).regs(localChan).nowAddr, dmaCtl(ctlIndex).regs(localChan).nowCount)
    Else
        value32 = U32Add(dmaCtl(ctlIndex).regs(localChan).nowAddr, dmaCtl(ctlIndex).regs(localChan).nowCount)
    End If

    DMA_ReadChanPort = DMA_U8(U32Shr(value32, dmaCtl(ctlIndex).dshift + (ff * 8&)))
End Function

Private Sub DMA_WriteChanPort(ByVal ctlIndex As Long, ByVal localPort As Long, ByVal value As Byte)
    Dim iport As Long
    Dim localChan As Long
    Dim regIndex As Long
    Dim ff As Long

    iport = (U32Shr(localPort, dmaCtl(ctlIndex).dshift) And &HF&)
    localChan = U32Shr(iport, 1&)
    regIndex = (iport And 1&)
    ff = DMA_GetFlipFlop(ctlIndex)

    If ff <> 0& Then
        If regIndex = 0& Then
            dmaCtl(ctlIndex).regs(localChan).baseAddr = ((dmaCtl(ctlIndex).regs(localChan).baseAddr And &HFF&) Or (CLng(value) * &H100&))
        Else
            dmaCtl(ctlIndex).regs(localChan).baseCount = ((dmaCtl(ctlIndex).regs(localChan).baseCount And &HFF&) Or (CLng(value) * &H100&))
        End If
        DMA_InitChannel ctlIndex, localChan
    Else
        If regIndex = 0& Then
            dmaCtl(ctlIndex).regs(localChan).baseAddr = ((dmaCtl(ctlIndex).regs(localChan).baseAddr And &HFF00&) Or CLng(value))
        Else
            dmaCtl(ctlIndex).regs(localChan).baseCount = ((dmaCtl(ctlIndex).regs(localChan).baseCount And &HFF00&) Or CLng(value))
        End If
    End If
End Sub

Private Function DMA_ReadContPort(ByVal ctlIndex As Long, ByVal localPort As Long) As Byte
    Dim iport As Long

    iport = (U32Shr(localPort, dmaCtl(ctlIndex).dshift) And &HF&)
    Select Case iport
        Case 0&
            DMA_ReadContPort = dmaCtl(ctlIndex).status
            dmaCtl(ctlIndex).status = CByte(dmaCtl(ctlIndex).status And &HF0&)
        Case 1&
            DMA_ReadContPort = CByte(dmaCtl(ctlIndex).mask And &HFF&)
        Case Else
            DMA_ReadContPort = 0&
    End Select
End Function

Private Sub DMA_WriteContPort(ByVal ctlIndex As Long, ByVal localPort As Long, ByVal value As Byte)
    Dim iport As Long
    Dim localChan As Long
    Dim maskBit As Long

    iport = (U32Shr(localPort, dmaCtl(ctlIndex).dshift) And &HF&)

    Select Case iport
        Case 0&
            If (value <> 0&) And ((value And DMA_CMD_NOT_SUPPORTED) <> 0&) Then Exit Sub
            dmaCtl(ctlIndex).command = value

        Case 1&
            localChan = (value And 3&)
            maskBit = DMA_MaskBit(localChan)
            If (value And 4&) <> 0& Then
                dmaCtl(ctlIndex).status = CByte((dmaCtl(ctlIndex).status Or (maskBit * 16&)) And &HFF&)
            Else
                dmaCtl(ctlIndex).status = CByte(dmaCtl(ctlIndex).status And (Not (maskBit * 16&)))
            End If
            dmaCtl(ctlIndex).status = CByte(dmaCtl(ctlIndex).status And (Not maskBit))
            DMA_RunController ctlIndex

        Case 2&
            localChan = (value And 3&)
            maskBit = DMA_MaskBit(localChan)
            If (value And 4&) <> 0& Then
                dmaCtl(ctlIndex).mask = CByte((dmaCtl(ctlIndex).mask Or maskBit) And &HFF&)
            Else
                dmaCtl(ctlIndex).mask = CByte(dmaCtl(ctlIndex).mask And (Not maskBit))
            End If
            DMA_RunController ctlIndex

        Case 3&
            localChan = (value And 3&)
            dmaCtl(ctlIndex).regs(localChan).mode = value

        Case 4&
            dmaCtl(ctlIndex).flipFlop = 0&

        Case 5&
            dmaCtl(ctlIndex).flipFlop = 0&
            dmaCtl(ctlIndex).mask = &HF&
            dmaCtl(ctlIndex).status = 0&
            dmaCtl(ctlIndex).command = 0&

        Case 6&
            dmaCtl(ctlIndex).mask = 0&
            DMA_RunController ctlIndex

        Case 7&
            dmaCtl(ctlIndex).mask = CByte(value And &HF&)
            DMA_RunController ctlIndex
    End Select
End Sub

Private Function DMA_ChannelCompatRead(ByVal channel As Long, ByRef value As Byte) As Byte
    Dim buffer(0& To 0&) As Byte
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim transferred As Long
    Dim pos As Long
    Dim totalLen As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)

    pos = dmaCtl(ctlIndex).regs(localChan).nowCount
    totalLen = DMA_TotalLenBytes(ctlIndex, localChan)
    If pos >= totalLen Then
        value = 0&
        DMA_ChannelCompatRead = 1&
        Exit Function
    End If

    transferred = i8237_dma_readMemory(channel, buffer, pos, 1&)
    If transferred <= 0& Then
        value = 0&
        DMA_ChannelCompatRead = 0&
        Exit Function
    End If

    value = buffer(0&)
    DMA_CommitPosition ctlIndex, localChan, pos + transferred
    If (dmaCtl(ctlIndex).status And DMA_MaskBit(localChan)) <> 0& Then
        DMA_ChannelCompatRead = 1&
    Else
        DMA_ChannelCompatRead = 0&
    End If
End Function

Private Function DMA_ChannelCompatWrite(ByVal channel As Long, ByVal value As Byte) As Byte
    Dim buffer(0& To 0&) As Byte
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim transferred As Long
    Dim pos As Long
    Dim totalLen As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    pos = dmaCtl(ctlIndex).regs(localChan).nowCount
    totalLen = DMA_TotalLenBytes(ctlIndex, localChan)
    If pos >= totalLen Then
        DMA_ChannelCompatWrite = 1&
        Exit Function
    End If

    buffer(0&) = value
    transferred = i8237_dma_writeMemory(channel, buffer, pos, 1&)
    If transferred <= 0& Then
        DMA_ChannelCompatWrite = 0&
        Exit Function
    End If

    DMA_CommitPosition ctlIndex, localChan, pos + transferred
    If (dmaCtl(ctlIndex).status And DMA_MaskBit(localChan)) <> 0& Then
        DMA_ChannelCompatWrite = 1&
    Else
        DMA_ChannelCompatWrite = 0&
    End If
End Function

Public Sub i8237_dma_registerChannel(ByVal channel As Long, ByVal callbackId As Long, ByVal callbackData As Long)
    channel = (channel And 7&)
    dmaHandlers(channel).callbackId = callbackId
    dmaHandlers(channel).callbackData = callbackData
End Sub

Public Sub i8237_dma_holdDreq(ByVal channel As Long)
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim maskBit As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    maskBit = DMA_MaskBit(localChan)

    dmaCtl(ctlIndex).status = CByte((dmaCtl(ctlIndex).status Or (maskBit * 16&)) And &HFF&)
    DMA_RunController ctlIndex
End Sub

Public Sub i8237_dma_releaseDreq(ByVal channel As Long)
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim maskBit As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    maskBit = DMA_MaskBit(localChan)

    dmaCtl(ctlIndex).status = CByte(dmaCtl(ctlIndex).status And (Not (maskBit * 16&)))
    DMA_RunController ctlIndex
End Sub

Public Function i8237_dma_readMemory(ByVal channel As Long, ByRef buffer() As Byte, ByVal pos As Long, ByVal length As Long) As Long
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim baseAddr As Long
    Dim physAddr As Long
    Dim i As Long
    Dim decrement As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    baseAddr = DMA_ComposeAddr(dmaCtl(ctlIndex).regs(localChan).pageHigh, dmaCtl(ctlIndex).regs(localChan).page, dmaCtl(ctlIndex).regs(localChan).nowAddr)

    If DMA_IsVerifyTransfer(dmaCtl(ctlIndex).regs(localChan).mode) <> 0& Then
        i8237_dma_readMemory = length
        Exit Function
    End If

    decrement = (dmaCtl(ctlIndex).regs(localChan).mode And &H20&)
    For i = 0& To length - 1&
        If decrement <> 0& Then
            physAddr = DMA_StepAddress(baseAddr, pos + length - 1& - i, 1&)
        Else
            physAddr = DMA_StepAddress(baseAddr, pos + i, 0&)
        End If
        buffer(i) = cpu_read_linear(machine.CPU, physAddr)
    Next i

    i8237_dma_readMemory = length
End Function

Public Function i8237_dma_writeMemory(ByVal channel As Long, ByRef buffer() As Byte, ByVal pos As Long, ByVal length As Long) As Long
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim baseAddr As Long
    Dim physAddr As Long
    Dim i As Long
    Dim decrement As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    baseAddr = DMA_ComposeAddr(dmaCtl(ctlIndex).regs(localChan).pageHigh, dmaCtl(ctlIndex).regs(localChan).page, dmaCtl(ctlIndex).regs(localChan).nowAddr)

    If DMA_IsVerifyTransfer(dmaCtl(ctlIndex).regs(localChan).mode) <> 0& Then
        i8237_dma_writeMemory = length
        Exit Function
    End If

    decrement = (dmaCtl(ctlIndex).regs(localChan).mode And &H20&)
    For i = 0& To length - 1&
        If decrement <> 0& Then
            physAddr = DMA_StepAddress(baseAddr, pos + length - 1& - i, 1&)
        Else
            physAddr = DMA_StepAddress(baseAddr, pos + i, 0&)
        End If
        cpu_write_linear machine.CPU, physAddr, buffer(i)
        wrcache_flush
    Next i

    i8237_dma_writeMemory = length
End Function

Public Sub i8237_dma_run()
    DMA_RunController 0&
    DMA_RunController 1&
End Sub

Public Sub i8237_set_drq(ByVal channel As Long, ByVal asserted As Byte)
    If asserted <> 0& Then
        i8237_dma_holdDreq channel
    Else
        i8237_dma_releaseDreq channel
    End If
End Sub

Public Function i8237_get_drq(ByVal channel As Long) As Byte
    Dim ctlIndex As Long
    Dim localChan As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    If (dmaCtl(ctlIndex).status And (DMA_MaskBit(localChan) * 16&)) <> 0& Then
        i8237_get_drq = 1&
    Else
        i8237_get_drq = 0&
    End If
End Function

Public Function i8237_get_mode(ByVal channel As Long) As Byte
    i8237_get_mode = dmaCtl(DMA_ChannelController(channel)).regs(DMA_ChannelLocal(channel)).mode
End Function

Public Function i8237_get_operation(ByVal channel As Long) As Long
    i8237_get_operation = ((CLng(i8237_get_mode(channel)) And &HC&) \ 4&)
End Function

Public Function i8237_get_terminal(ByVal channel As Long) As Byte
    Dim ctlIndex As Long
    Dim localChan As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    If (dmaCtl(ctlIndex).status And DMA_MaskBit(localChan)) <> 0& Then
        i8237_get_terminal = 1&
    Else
        i8237_get_terminal = 0&
    End If
End Function

Public Sub i8237_clear_terminal(ByVal channel As Long)
    Dim ctlIndex As Long
    Dim localChan As Long

    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    dmaCtl(ctlIndex).status = CByte(dmaCtl(ctlIndex).status And (Not DMA_MaskBit(localChan)))
End Sub

Public Function i8237_fdc_writeByte(ByVal channel As Long, ByVal value As Byte) As Byte
    i8237_fdc_writeByte = DMA_ChannelCompatWrite(channel, value)
End Function

Public Function i8237_fdc_readByte(ByVal channel As Long, ByRef value As Byte) As Byte
    i8237_fdc_readByte = DMA_ChannelCompatRead(channel, value)
End Function

Public Function i8237_read(ByVal channel As Byte) As Byte
    Dim value As Byte

    Call DMA_ChannelCompatRead(channel, value)
    i8237_read = value
End Function

Public Sub i8237_write(ByVal channel As Byte, ByVal value As Byte)
    Call DMA_ChannelCompatWrite(channel, value)
End Sub

Public Function i8237_readpage(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim channel As Long
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim offset As Long

    offset = ((addr And &HFFFF&) - &H80&) And &HF&
    i8237_readpage = dmaPageShadow(offset)

    channel = DMA_PagePortChannel(addr)
    If channel < 0& Then Exit Function
    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    i8237_readpage = dmaCtl(ctlIndex).regs(localChan).page
End Function

Public Sub i8237_writepage(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    Dim channel As Long
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim offset As Long

    offset = ((addr And &HFFFF&) - &H80&) And &HF&
    dmaPageShadow(offset) = value

    channel = DMA_PagePortChannel(addr)
    If channel < 0& Then Exit Sub
    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    dmaCtl(ctlIndex).regs(localChan).page = value
End Sub

Public Function i8237_readpageh(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim channel As Long
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim offset As Long

    offset = ((addr And &HFFFF&) - &H480&) And &HF&
    i8237_readpageh = dmaPageHighShadow(offset)

    channel = DMA_PagePortChannel(addr)
    If channel < 0& Then Exit Function
    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    i8237_readpageh = dmaCtl(ctlIndex).regs(localChan).pageHigh
End Function

Public Sub i8237_writepageh(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    Dim channel As Long
    Dim ctlIndex As Long
    Dim localChan As Long
    Dim offset As Long

    offset = ((addr And &HFFFF&) - &H480&) And &HF&
    dmaPageHighShadow(offset) = value

    channel = DMA_PagePortChannel(addr)
    If channel < 0& Then Exit Sub
    ctlIndex = DMA_ChannelController(channel)
    localChan = DMA_ChannelLocal(channel)
    dmaCtl(ctlIndex).regs(localChan).pageHigh = value
End Sub

Public Function i8237_readport(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim ctlIndex As Long
    Dim localPort As Long
    Dim span As Long

    If DMA_ControllerFromPort(addr, ctlIndex, localPort) = 0& Then
        i8237_readport = &HFF&
        Exit Function
    End If

    span = DMA_ControlSpan(ctlIndex)
    If localPort < span Then
        i8237_readport = DMA_ReadChanPort(ctlIndex, localPort)
    Else
        i8237_readport = DMA_ReadContPort(ctlIndex, localPort - span)
    End If
End Function

Public Sub i8237_writeport(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    Dim ctlIndex As Long
    Dim localPort As Long
    Dim span As Long

    If DMA_ControllerFromPort(addr, ctlIndex, localPort) = 0& Then Exit Sub

    span = DMA_ControlSpan(ctlIndex)
    If localPort < span Then
        DMA_WriteChanPort ctlIndex, localPort, value
    Else
        DMA_WriteContPort ctlIndex, localPort - span, value
    End If
End Sub

Public Sub i8237_init(ByRef machineRef As MACHINE_t)
    Dim channel As Long
    Dim i As Long

    Call DMA_InitController(0&, &H0&, &H80&, &H480&, 0&)
    Call DMA_InitController(1&, &HC0&, &H88&, &H488&, 1&)

    For channel = 0& To 7&
        dmaHandlers(channel).callbackId = DMA_CB_NONE
        dmaHandlers(channel).callbackData = 0&
    Next channel
    For i = 0& To 15&
        dmaPageShadow(i) = 0&
        dmaPageHighShadow(i) = 0&
    Next i

    ports_cbRegister &H0&, &H10&, PORTS_CB_I8237_PORT, PORTS_CB_NONE, PORTS_CB_I8237_PORT, PORTS_CB_NONE, 0&
    ports_cbRegister &HC0&, &H20&, PORTS_CB_I8237_PORT, PORTS_CB_NONE, PORTS_CB_I8237_PORT, PORTS_CB_NONE, 0&
    ports_cbRegister &H80&, &H10&, PORTS_CB_I8237_PAGE, PORTS_CB_NONE, PORTS_CB_I8237_PAGE, PORTS_CB_NONE, 0&
    ports_cbRegister &H480&, &H10&, PORTS_CB_I8237_PAGEH, PORTS_CB_NONE, PORTS_CB_I8237_PAGEH, PORTS_CB_NONE, 0&
End Sub
Public Sub dma_bm_read(ByVal phys_addr As Long, ByRef data() As Byte, ByVal total_size As Long, ByVal transfer_size As Long)
    Dim aligned As Long
    Dim tail As Long
    Dim bytes(0 To 3) As Byte
    Dim i As Long
    Dim j As Long
    Dim idx As Long

    If total_size <= 0& Then Exit Sub
    If transfer_size <= 0& Then transfer_size = 1&
    If transfer_size > 4& Then transfer_size = 4&

    aligned = (total_size And (Not (transfer_size - 1&)))
    tail = (total_size - aligned)

    For i = 0& To aligned - 1& Step transfer_size
        For j = 0& To transfer_size - 1&
            data(i + j) = cpu_read_linear(machine.CPU, U32Add(phys_addr, i + j))
        Next j
    Next i

    If tail <> 0& Then
        For j = 0& To transfer_size - 1&
            bytes(j) = cpu_read_linear(machine.CPU, U32Add(phys_addr, aligned + j))
        Next j
        For idx = 0& To tail - 1&
            data(aligned + idx) = bytes(idx)
        Next idx
    End If
End Sub

Public Sub dma_bm_write(ByVal phys_addr As Long, ByRef data() As Byte, ByVal total_size As Long, ByVal transfer_size As Long)
    Dim aligned As Long
    Dim tail As Long
    Dim bytes(0 To 3) As Byte
    Dim i As Long
    Dim j As Long
    Dim idx As Long

    If total_size <= 0& Then Exit Sub
    If transfer_size <= 0& Then transfer_size = 1&
    If transfer_size > 4& Then transfer_size = 4&

    aligned = (total_size And (Not (transfer_size - 1&)))
    tail = (total_size - aligned)

    For i = 0& To aligned - 1& Step transfer_size
        For j = 0& To transfer_size - 1&
            cpu_write_linear machine.CPU, U32Add(phys_addr, i + j), data(i + j)
            wrcache_flush
        Next j
    Next i

    If tail <> 0& Then
        For j = 0& To transfer_size - 1&
            bytes(j) = cpu_read_linear(machine.CPU, U32Add(phys_addr, aligned + j))
        Next j
        For idx = 0& To tail - 1&
            bytes(idx) = data(aligned + idx)
        Next idx
        For j = 0& To transfer_size - 1&
            cpu_write_linear machine.CPU, U32Add(phys_addr, aligned + j), bytes(j)
            wrcache_flush
        Next j
    End If
End Sub
