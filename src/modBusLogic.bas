Attribute VB_Name = "modBusLogic"
Option Explicit

Private Const ROM_SIZE As Long = &H4000&

Private Const BUSLOGIC_FLAG_MBX_24BIT As Byte = &H1&
Private Const BUSLOGIC_FLAG_CDROM_BOOT As Byte = &H2&
Private Const BUSLOGIC_FLAG_INT_GEOM_WRITABLE As Byte = &H4&

Private Const BUSLOGIC_RESET_DELAY_US As Double = 50000#

Private Const CTRL_HRST As Byte = &H80&
Private Const CTRL_SRST As Byte = &H40&
Private Const CTRL_IRST As Byte = &H20&
Private Const CTRL_SCRST As Byte = &H10&

Private Const STAT_STST As Byte = &H80&
Private Const STAT_DFAIL As Byte = &H40&
Private Const STAT_INIT As Byte = &H20&
Private Const STAT_IDLE As Byte = &H10&
Private Const STAT_CDFULL As Byte = &H8&
Private Const STAT_DFULL As Byte = &H4&
Private Const STAT_INVCMD As Byte = &H1&

Private Const CMD_NOP As Byte = &H0&
Private Const CMD_MBINIT As Byte = &H1&
Private Const CMD_START_SCSI As Byte = &H2&
Private Const CMD_BIOSCMD As Byte = &H3&
Private Const CMD_INQUIRY As Byte = &H4&
Private Const CMD_EMBOI As Byte = &H5&
Private Const CMD_SELTIMEOUT As Byte = &H6&
Private Const CMD_BUSON_TIME As Byte = &H7&
Private Const CMD_BUSOFF_TIME As Byte = &H8&
Private Const CMD_DMASPEED As Byte = &H9&
Private Const CMD_RETDEVS As Byte = &HA&
Private Const CMD_RETCONF As Byte = &HB&
Private Const CMD_TARGET As Byte = &HC&
Private Const CMD_RETSETUP As Byte = &HD&
Private Const CMD_WRITE_CH2 As Byte = &H1A&
Private Const CMD_READ_CH2 As Byte = &H1B&
Private Const CMD_ECHO As Byte = &H1F&
Private Const CMD_OPTIONS As Byte = &H21&

Private Const INTR_ANY As Byte = &H80&
Private Const INTR_HACC As Byte = &H4&
Private Const INTR_MBOA As Byte = &H2&
Private Const INTR_MBIF As Byte = &H1&

Private Const MBO_FREE As Byte = &H0&
Private Const MBO_START As Byte = &H1&
Private Const MBO_ABORT As Byte = &H2&

Private Const MBI_SUCCESS As Byte = &H1&
Private Const MBI_NOT_FOUND As Byte = &H3&
Private Const MBI_ERROR As Byte = &H4&

Private Const SCSI_INITIATOR_COMMAND As Byte = &H0&
Private Const TARGET_MODE_COMMAND As Byte = &H1&
Private Const SCATTER_GATHER_COMMAND As Byte = &H2&
Private Const SCSI_INITIATOR_COMMAND_RES As Byte = &H3&
Private Const SCATTER_GATHER_COMMAND_RES As Byte = &H4&
Private Const BUS_RESET_OPCODE As Byte = &H81&

Private Const CCB_DATA_XFER_IN As Byte = &H1&
Private Const CCB_DATA_XFER_OUT As Byte = &H2&

Private Const CCB_COMPLETE As Byte = &H0&
Private Const CCB_SELECTION_TIMEOUT As Byte = &H11&
Private Const CCB_INVALID_OP_CODE As Byte = &H16&
Private Const CCB_INVALID_CCB As Byte = &H1A&
Private Const CCB_ABORTED As Byte = &H26&

Private Const BL_MAILBOX24_SIZE As Long = 4&
Private Const BL_MAILBOX32_SIZE As Long = 8&
Private Const BL_CCB24_SIZE As Long = 30&
Private Const BL_CCB32_SIZE As Long = 40&
Private Const BL_SGE24_SIZE As Long = 6&
Private Const BL_SGE32_SIZE As Long = 8&
Private Const BL_MAILBOXINIT24_SIZE As Long = 4&
Private Const BL_MAILBOXINIT32_SIZE As Long = 5&
Private Const BL_ESCMD_HEADER_SIZE As Long = 12&
Private Const BL_ESCMD_MAX_SIZE As Long = 24&

Private Type BUSLOGIC_REQ_t
    CCB(0 To BL_CCB32_SIZE - 1) As Byte
    CCBPointer As Long
    Is24Bit As Byte
    TargetID As Byte
    LUN As Byte
    HostStatus As Byte
    TargetStatus As Byte
    MailboxCompletionCode As Byte
End Type

Private Type BUSLOGIC_t
    bus As Byte
    Base As Long
    Irq As Byte
    DmaChannel As Byte
    HostID As Byte
    flags As Byte
    Status As Byte
    Interrupt As Byte
    IrqEnabled As Byte
    Geometry As Byte
    Command As Byte
    CmdParam As Byte
    CmdParamLeft As Long
    DataReply As Long
    DataReplyLeft As Long
    CmdBuf(0 To 127) As Byte
    dma_buffer(0 To 63) As Byte
    temp_cdb(0 To 11) As Byte
    fw_rev(0 To 7) As Byte
    BusOnTime As Byte
    BusOffTime As Byte
    ATBusSpeed As Byte
    target_data_len As Long
    scsi_cmd_phase As Byte
    aggressive_round_robin As Byte
    ExtendedLUNCCBFormat As Byte
    transfer_size As Long
    MailboxInit As Long
    MailboxCount As Long
    MailboxOutAddr As Long
    MailboxOutPosCur As Long
    MailboxInAddr As Long
    MailboxInPosCur As Long
    MailboxReq As Long
    MailboxOutInterrupts As Byte
    PendingInterrupt As Byte
    ToRaise As Byte
    callback_sub_phase As Byte
    Outgoing As Long
    mail_timer_id As Long
    reset_timer_id As Long
    rom_loaded As Byte
    rom_addr As Long
    rom_path As String * 512
    nvr_path As String * 512
    LocalRAM(0 To 255) As Byte
    Req As BUSLOGIC_REQ_t
    pic_slave As Long
    initialized As Byte
End Type

Private buslogic As BUSLOGIC_t
Private buslogicDataBuf() As Byte
Private buslogicRomData() As Byte

Private Function BL_Min(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        BL_Min = a
    Else
        BL_Min = b
    End If
End Function

Private Function BL_FixedString(ByVal fixedValue As String) As String
    Dim nulPos As Long

    nulPos = InStr(1&, fixedValue, vbNullChar)
    If nulPos > 0& Then
        BL_FixedString = Left$(fixedValue, nulPos - 1&)
    Else
        BL_FixedString = RTrim$(fixedValue)
    End If
End Function

Private Function BL_ReadU16LE(ByRef data() As Byte, ByVal offset As Long) As Long
    BL_ReadU16LE = ((CLng(data(offset + 1&)) * &H100&) Or CLng(data(offset))) And &HFFFF&
End Function

Private Sub BL_WriteU16LE(ByRef data() As Byte, ByVal offset As Long, ByVal value As Long)
    data(offset) = CByte(value And &HFF&)
    data(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
End Sub

Private Function BL_ReadU32LE(ByRef data() As Byte, ByVal offset As Long) As Long
    BL_ReadU32LE = U32FromDouble(CDbl(data(offset)) + CDbl(data(offset + 1&)) * 256# + CDbl(data(offset + 2&)) * 65536# + CDbl(data(offset + 3&)) * 16777216#)
End Function

Private Sub BL_WriteU32LE(ByRef data() As Byte, ByVal offset As Long, ByVal value As Long)
    data(offset) = CByte(value And &HFF&)
    data(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
    data(offset + 2&) = CByte(U32Shr(value, 16&) And &HFF&)
    data(offset + 3&) = CByte(U32Shr(value, 24&) And &HFF&)
End Sub

Private Function BL_ReadAddr24(ByRef data() As Byte, ByVal offset As Long) As Long
    BL_ReadAddr24 = U32FromDouble(CDbl(data(offset + 2&)) + CDbl(data(offset + 1&)) * 256# + CDbl(data(offset)) * 65536#)
End Function

Private Sub BL_WriteAddr24(ByRef data() As Byte, ByVal offset As Long, ByVal value As Long)
    data(offset) = CByte(U32Shr(value, 16&) And &HFF&)
    data(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
    data(offset + 2&) = CByte(value And &HFF&)
End Sub

Private Function BL_ReadU32LE_ReqCCB(ByVal offset As Long) As Long
    BL_ReadU32LE_ReqCCB = U32FromDouble(CDbl(buslogic.Req.CCB(offset)) + CDbl(buslogic.Req.CCB(offset + 1&)) * 256# + CDbl(buslogic.Req.CCB(offset + 2&)) * 65536# + CDbl(buslogic.Req.CCB(offset + 3&)) * 16777216#)
End Function

Private Function BL_ReadAddr24_ReqCCB(ByVal offset As Long) As Long
    BL_ReadAddr24_ReqCCB = U32FromDouble(CDbl(buslogic.Req.CCB(offset + 2&)) + CDbl(buslogic.Req.CCB(offset + 1&)) * 256# + CDbl(buslogic.Req.CCB(offset)) * 65536#)
End Function

Private Function BL_ReadU32LE_CmdBuf(ByVal offset As Long) As Long
    BL_ReadU32LE_CmdBuf = U32FromDouble(CDbl(buslogic.CmdBuf(offset)) + CDbl(buslogic.CmdBuf(offset + 1&)) * 256# + CDbl(buslogic.CmdBuf(offset + 2&)) * 65536# + CDbl(buslogic.CmdBuf(offset + 3&)) * 16777216#)
End Function

Private Function BL_ReadAddr24_CmdBuf(ByVal offset As Long) As Long
    BL_ReadAddr24_CmdBuf = U32FromDouble(CDbl(buslogic.CmdBuf(offset + 2&)) + CDbl(buslogic.CmdBuf(offset + 1&)) * 256# + CDbl(buslogic.CmdBuf(offset)) * 65536#)
End Function

Private Sub BL_WriteU16LE_DataBuf(ByVal offset As Long, ByVal value As Long)
    buslogicDataBuf(offset) = CByte(value And &HFF&)
    buslogicDataBuf(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
End Sub

Private Sub BL_WriteU32LE_DataBuf(ByVal offset As Long, ByVal value As Long)
    buslogicDataBuf(offset) = CByte(value And &HFF&)
    buslogicDataBuf(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
    buslogicDataBuf(offset + 2&) = CByte(U32Shr(value, 16&) And &HFF&)
    buslogicDataBuf(offset + 3&) = CByte(U32Shr(value, 24&) And &HFF&)
End Sub

Private Sub BL_WriteAddr24_DataBuf(ByVal offset As Long, ByVal value As Long)
    buslogicDataBuf(offset) = CByte(U32Shr(value, 16&) And &HFF&)
    buslogicDataBuf(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
    buslogicDataBuf(offset + 2&) = CByte(value And &HFF&)
End Sub

Private Sub BL_CopyCmdBuf(ByRef dst() As Byte, ByVal count As Long)
    Dim i As Long

    If count <= 0& Then
        ReDim dst(0 To 0) As Byte
        Exit Sub
    End If

    ReDim dst(0 To count - 1&) As Byte
    For i = 0& To count - 1&
        dst(i) = buslogic.CmdBuf(i)
    Next i
End Sub

Private Sub BL_CopyTempCdb(ByRef dst() As Byte)
    Dim i As Long

    ReDim dst(0 To 11&) As Byte
    For i = 0& To 11&
        dst(i) = buslogic.temp_cdb(i)
    Next i
End Sub

Private Sub BL_ClearRange(ByRef data() As Byte, ByVal count As Long)
    Dim i As Long

    For i = 0& To count - 1&
        data(i) = 0&
    Next i
End Sub

Private Sub BL_CopyBytes(ByRef dst() As Byte, ByVal dstOffset As Long, ByRef src() As Byte, ByVal srcOffset As Long, ByVal count As Long)
    Dim i As Long

    For i = 0& To count - 1&
        dst(dstOffset + i) = src(srcOffset + i)
    Next i
End Sub

Private Sub BL_DMAReadBytes(ByVal phys_addr As Long, ByRef data() As Byte, ByVal total_size As Long)
    dma_bm_read phys_addr, data, total_size, buslogic.transfer_size
End Sub

Private Sub BL_DMAWriteBytes(ByVal phys_addr As Long, ByRef data() As Byte, ByVal total_size As Long)
    dma_bm_write phys_addr, data, total_size, buslogic.transfer_size
End Sub

Private Function buslogic_interval_from_us(ByVal usec As Double) As Double
    Dim ticks As Double

    ticks = (usec / 1000000#) * timing_getFreq()
    If ticks < 1# Then ticks = 1#
    buslogic_interval_from_us = ticks
End Function

Private Sub buslogic_schedule_mail()
    timing_updateInterval buslogic.mail_timer_id, buslogic_interval_from_us(5#)
    timing_timerEnable buslogic.mail_timer_id
End Sub

Private Sub buslogic_set_localram_u16(ByVal offset As Long, ByVal value As Long)
    buslogic.LocalRAM(offset) = CByte(value And &HFF&)
    buslogic.LocalRAM(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
End Sub

Private Function buslogic_get_localram_u16(ByVal offset As Long) As Long
    buslogic_get_localram_u16 = ((CLng(buslogic.LocalRAM(offset + 1&)) * &H100&) Or CLng(buslogic.LocalRAM(offset))) And &HFFFF&
End Function

Private Sub buslogic_save_nvr()
    Dim path As String
    Dim fn As Integer

    path = BL_FixedString(buslogic.nvr_path)
    If LenB(path) = 0& Then Exit Sub

    On Error GoTo SaveFail
    fn = FreeFile
    Open path For Binary Access Write As #fn
    Put #fn, 1, buslogic.LocalRAM
    Close #fn
    Exit Sub

SaveFail:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    On Error GoTo 0
End Sub

Private Sub buslogic_autoscsi_defaults(ByVal safe As Byte)
    Dim i As Long
    Dim biosConfig As Long

    For i = 0& To 255&
        buslogic.LocalRAM(i) = 0&
    Next i

    buslogic.LocalRAM(64) = Asc("F")
    buslogic.LocalRAM(65) = Asc("A")
    buslogic.LocalRAM(66) = 64&
    buslogic.LocalRAM(68) = Asc("5")
    buslogic.LocalRAM(69) = Asc("4")
    buslogic.LocalRAM(70) = Asc("5")
    buslogic.LocalRAM(71) = Asc("S")

    Select Case buslogic.DmaChannel
        Case 5&
            buslogic.LocalRAM(75) = 1&
        Case 6&
            buslogic.LocalRAM(75) = 2&
        Case 7&
            buslogic.LocalRAM(75) = 3&
        Case Else
            buslogic.LocalRAM(75) = 0&
    End Select

    Select Case buslogic.Irq
        Case 9&
            buslogic.LocalRAM(76) = 1&
        Case 10&
            buslogic.LocalRAM(76) = 2&
        Case 11&
            buslogic.LocalRAM(76) = 3&
        Case 12&
            buslogic.LocalRAM(76) = 4&
        Case 14&
            buslogic.LocalRAM(76) = 5&
        Case 15&
            buslogic.LocalRAM(76) = 6&
        Case Else
            buslogic.LocalRAM(76) = 0&
    End Select

    buslogic.LocalRAM(77) = 1&
    buslogic.LocalRAM(78) = 7&
    buslogic.LocalRAM(79) = &H3F&
    buslogic.LocalRAM(80) = 7&
    buslogic.LocalRAM(81) = 4&

    biosConfig = IIf(buslogic.rom_addr <> 0&, &H33&, &H32&)
    If safe = 0& Then biosConfig = (biosConfig Or &H4&)
    buslogic.LocalRAM(82) = CByte(biosConfig And &HFF&)

    buslogic_set_localram_u16 83&, &HFFFF&
    buslogic_set_localram_u16 87&, &HFFFF&
    buslogic_set_localram_u16 91&, &HFFFF&
    buslogic.LocalRAM(97) = IIf(safe <> 0&, &H10&, 0&)
    buslogic.LocalRAM(105) = 7&
    buslogic.LocalRAM(107) = IIf(safe <> 0&, 0&, (&H1 Or &H4 Or &H20))
    buslogic.LocalRAM(109) = 1&
End Sub

Private Sub buslogic_load_nvr()
    Dim path As String
    Dim fn As Integer
    Dim fileLen As Long

    buslogic_autoscsi_defaults 0&
    path = BL_FixedString(buslogic.nvr_path)
    If LenB(path) = 0& Then Exit Sub

    On Error GoTo LoadFail
    fn = FreeFile
    Open path For Binary Access Read As #fn
    fileLen = LOF(fn)
    If fileLen < 256& Then GoTo LoadFail
    Get #fn, 1, buslogic.LocalRAM
    Close #fn
    Exit Sub

LoadFail:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    On Error GoTo 0
    buslogic_autoscsi_defaults 0&
    buslogic_save_nvr
End Sub
Private Function BL_ReqControlByte() As Byte
    If buslogic.Req.Is24Bit <> 0& Then
        BL_ReqControlByte = CByte((U32Shr(CLng(buslogic.Req.CCB(1)), 3&) And &H3&) And &HFF&)
    Else
        BL_ReqControlByte = CByte((U32Shr(CLng(buslogic.Req.CCB(1)), 3&) And &H3&) And &HFF&)
    End If
End Function

Private Function BL_ReqOpcode() As Byte
    BL_ReqOpcode = buslogic.Req.CCB(0)
End Function

Private Function BL_ReqCdbLength() As Byte
    BL_ReqCdbLength = buslogic.Req.CCB(2)
End Function

Private Function BL_ReqRequestSenseLength() As Byte
    BL_ReqRequestSenseLength = buslogic.Req.CCB(3)
End Function

Private Function BL_ReqDataLength() As Long
    If buslogic.Req.Is24Bit <> 0& Then
        BL_ReqDataLength = BL_ReadAddr24_ReqCCB(4&)
    Else
        BL_ReqDataLength = BL_ReadU32LE_ReqCCB(4&)
    End If
End Function

Private Function BL_ReqDataPointer() As Long
    If buslogic.Req.Is24Bit <> 0& Then
        BL_ReqDataPointer = BL_ReadAddr24_ReqCCB(7&)
    Else
        BL_ReqDataPointer = BL_ReadU32LE_ReqCCB(8&)
    End If
End Function

Private Function BL_ReqSensePointer() As Long
    If buslogic.Req.Is24Bit <> 0& Then
        BL_ReqSensePointer = U32Add(buslogic.Req.CCBPointer, &H1E&)
    Else
        BL_ReqSensePointer = BL_ReadU32LE_ReqCCB(36&)
    End If
End Function

Private Function BL_RequestSenseLength(ByVal value As Byte) As Byte
    If value = 0& Then
        BL_RequestSenseLength = 14&
    ElseIf value = 1& Then
        BL_RequestSenseLength = 0&
    Else
        BL_RequestSenseLength = value
    End If
End Function

Private Function BL_CompletionCode(ByVal targetId As Long) As Byte
    Select Case scsi_devices(buslogic.bus, targetId).sense(12)
        Case ASC_NONE
            BL_CompletionCode = &H0&
        Case ASC_ILLEGAL_OPCODE, ASC_INV_FIELD_IN_CMD_PACKET, ASC_INV_FIELD_IN_PARAMETER_LIST
            BL_CompletionCode = &H1&
        Case ASC_LBA_OUT_OF_RANGE
            BL_CompletionCode = &H2&
        Case ASC_WRITE_PROTECTED
            BL_CompletionCode = &H3&
        Case ASC_INCOMPATIBLE_FORMAT
            BL_CompletionCode = &HC&
        Case ASC_NOT_READY, ASC_MEDIUM_MAY_HAVE_CHANGED, ASC_CAPACITY_DATA_CHANGED, ASC_MEDIUM_NOT_PRESENT
            BL_CompletionCode = &HAA&
        Case Else
            BL_CompletionCode = &HFF&
    End Select
End Function

Private Sub buslogic_cmd_phase1()
    If (buslogic.Command = &H90&) And (buslogic.CmdParam = 2&) Then
        buslogic.CmdParamLeft = buslogic.CmdBuf(1)
    End If
End Sub

Private Sub buslogic_set_irq_line(ByVal setLine As Long)
    If buslogic.Irq >= 8& Then
        If buslogic.pic_slave < 0& Then Exit Sub
        If setLine <> 0& Then
            i8259_doirq buslogic.pic_slave, CByte((buslogic.Irq - 8&) And &HFF&)
        Else
            i8259_clearirq buslogic.pic_slave, CByte((buslogic.Irq - 8&) And &HFF&)
        End If
    Else
        If setLine <> 0& Then
            i8259_doirq machine.i8259, buslogic.Irq
        Else
            i8259_clearirq machine.i8259, buslogic.Irq
        End If
    End If
End Sub

Private Sub buslogic_raise_irq(ByVal suppress As Long, ByVal interruptValue As Byte)
    If (interruptValue And (INTR_MBIF Or INTR_MBOA)) <> 0& Then
        If (buslogic.Interrupt And INTR_HACC) = 0& Then
            buslogic.Interrupt = CByte((buslogic.Interrupt Or interruptValue) And &HFF&)
        Else
            buslogic.PendingInterrupt = CByte((buslogic.PendingInterrupt Or interruptValue) And &HFF&)
        End If
    ElseIf (interruptValue And INTR_HACC) <> 0& Then
        buslogic.Interrupt = CByte((buslogic.Interrupt Or interruptValue) And &HFF&)
    End If

    buslogic.Interrupt = CByte((buslogic.Interrupt Or INTR_ANY) And &HFF&)
    If (buslogic.IrqEnabled <> 0&) And (suppress = 0&) Then
        buslogic_set_irq_line 1&
    End If
End Sub

Private Sub buslogic_clear_irq()
    Dim pending As Byte

    buslogic.Interrupt = 0&
    buslogic_set_irq_line 0&
    If buslogic.PendingInterrupt <> 0& Then
        pending = buslogic.PendingInterrupt
        buslogic.PendingInterrupt = 0&
        buslogic_raise_irq 0&, pending
    End If
End Sub

Private Sub buslogic_register_io()
    ports_cbRegister buslogic.Base, 4&, PORTS_CB_BUSLOGIC, PORTS_CB_NONE, PORTS_CB_BUSLOGIC, PORTS_CB_NONE, 0&
End Sub

Private Sub buslogic_reset(ByVal hard_reset As Long)
    Dim i As Long

    buslogic_clear_irq
    buslogic.Geometry = &H90&
    buslogic.Command = &HFF&
    buslogic.CmdParam = 0&
    buslogic.CmdParamLeft = 0&
    buslogic.DataReply = 0&
    buslogic.DataReplyLeft = 0&
    buslogic.flags = CByte((buslogic.flags Or BUSLOGIC_FLAG_MBX_24BIT) And &HFF&)
    buslogic.MailboxOutInterrupts = 0&
    buslogic.PendingInterrupt = 0&
    buslogic.MailboxInPosCur = 0&
    buslogic.MailboxOutPosCur = 0&
    buslogic.MailboxCount = 0&
    buslogic.MailboxReq = 0&
    buslogic.IrqEnabled = 1&
    buslogic.target_data_len = 0&
    buslogic.ToRaise = 0&
    buslogic.callback_sub_phase = 0&
    buslogic.Outgoing = 0&
    timing_timerDisable buslogic.mail_timer_id

    For i = 0& To SCSI_ID_MAX - 1&
        scsi_device_reset buslogic.bus, i
    Next i

    If hard_reset <> 0& Then
        buslogic.Status = STAT_STST
        timing_updateInterval buslogic.reset_timer_id, buslogic_interval_from_us(BUSLOGIC_RESET_DELAY_US)
        timing_timerEnable buslogic.reset_timer_id
    Else
        buslogic.Status = (STAT_INIT Or STAT_IDLE)
    End If
End Sub

Private Sub buslogic_rd_sge(ByVal is24bit As Long, ByVal address As Long, ByRef segmentLen As Long, ByRef segmentPtr As Long)
    Dim bytes() As Byte

    If is24bit <> 0& Then
        ReDim bytes(0 To BL_SGE24_SIZE - 1&) As Byte
        BL_DMAReadBytes address, bytes, BL_SGE24_SIZE
        segmentLen = BL_ReadAddr24(bytes, 0&)
        segmentPtr = BL_ReadAddr24(bytes, 3&)
    Else
        ReDim bytes(0 To BL_SGE32_SIZE - 1&) As Byte
        BL_DMAReadBytes address, bytes, BL_SGE32_SIZE
        segmentLen = BL_ReadU32LE(bytes, 0&)
        segmentPtr = BL_ReadU32LE(bytes, 4&)
    End If
End Sub

Private Function buslogic_get_length(ByVal is24bit As Long) As Long
    Dim data_pointer As Long
    Dim data_length As Long
    Dim sg_entry_len As Long
    Dim total As Long
    Dim segLen As Long
    Dim segPtr As Long
    Dim opcode As Byte
    Dim i As Long

    data_pointer = BL_ReqDataPointer()
    data_length = BL_ReqDataLength()
    sg_entry_len = IIf(is24bit <> 0&, BL_SGE24_SIZE, BL_SGE32_SIZE)
    opcode = BL_ReqOpcode()

    If (data_length = 0&) Or (BL_ReqControlByte() = &H3&) Then
        buslogic_get_length = 0&
        Exit Function
    End If

    If (opcode = SCATTER_GATHER_COMMAND) Or (opcode = SCATTER_GATHER_COMMAND_RES) Then
        total = 0&
        For i = 0& To data_length - 1& Step sg_entry_len
            buslogic_rd_sge is24bit, U32Add(data_pointer, i), segLen, segPtr
            total = U32Add(total, segLen)
        Next i
        buslogic_get_length = total
        Exit Function
    End If

    If (opcode = SCSI_INITIATOR_COMMAND) Or (opcode = SCSI_INITIATOR_COMMAND_RES) Then
        buslogic_get_length = data_length
    Else
        buslogic_get_length = 0&
    End If
End Function

Private Sub buslogic_set_residue(ByVal transfer_length As Long)
    Dim residue As Long
    Dim buf_len As Long
    Dim opcode As Byte
    Dim bytes() As Byte

    opcode = BL_ReqOpcode()
    If (opcode <> SCSI_INITIATOR_COMMAND_RES) And (opcode <> SCATTER_GATHER_COMMAND_RES) Then Exit Sub

    residue = 0&
    buf_len = scsi_devices(buslogic.bus, buslogic.Req.TargetID).buffer_length
    If (transfer_length > 0&) And (BL_ReqControlByte() < &H3&) Then
        transfer_length = transfer_length - buf_len
        If transfer_length > 0& Then residue = transfer_length
    End If

    If buslogic.Req.Is24Bit <> 0& Then
        ReDim bytes(0 To 3&) As Byte
        BL_DMAReadBytes U32Add(buslogic.Req.CCBPointer, 4&), bytes, 4&
        BL_WriteAddr24 bytes, 0&, residue
        BL_DMAWriteBytes U32Add(buslogic.Req.CCBPointer, 4&), bytes, 4&
    Else
        ReDim bytes(0 To 3&) As Byte
        BL_WriteU32LE bytes, 0&, residue
        BL_DMAWriteBytes U32Add(buslogic.Req.CCBPointer, 4&), bytes, 4&
    End If
End Sub

Private Sub buslogic_buf_dma_transfer(ByVal is24bit As Long, ByVal transfer_length As Long, ByVal dirDataOut As Long)
    Dim data_pointer As Long
    Dim data_length As Long
    Dim sg_entry_len As Long
    Dim buf_len As Long
    Dim read_from_host As Long
    Dim write_to_host As Long
    Dim opcode As Byte
    Dim i As Long
    Dim sg_pos As Long
    Dim segLen As Long
    Dim segPtr As Long
    Dim data_to_transfer As Long
    Dim temp() As Byte

    data_pointer = BL_ReqDataPointer()
    data_length = BL_ReqDataLength()
    sg_entry_len = IIf(is24bit <> 0&, BL_SGE24_SIZE, BL_SGE32_SIZE)
    buf_len = scsi_devices(buslogic.bus, buslogic.Req.TargetID).buffer_length
    read_from_host = IIf((dirDataOut <> 0&) And ((BL_ReqControlByte() = CCB_DATA_XFER_OUT) Or (BL_ReqControlByte() = 0&)), 1&, 0&)
    write_to_host = IIf((dirDataOut = 0&) And ((BL_ReqControlByte() = CCB_DATA_XFER_IN) Or (BL_ReqControlByte() = 0&)), 1&, 0&)
    opcode = BL_ReqOpcode()

    If (BL_ReqControlByte() = &H3&) Or (transfer_length = 0&) Or (buf_len <= 0&) Then Exit Sub
    If buf_len > scsi_temp_buffer_sz Then buf_len = scsi_temp_buffer_sz

    If (opcode = SCATTER_GATHER_COMMAND) Or (opcode = SCATTER_GATHER_COMMAND_RES) Then
        sg_pos = 0&
        For i = 0& To data_length - 1& Step sg_entry_len
            buslogic_rd_sge is24bit, U32Add(data_pointer, i), segLen, segPtr
            data_to_transfer = BL_Min(buf_len, segLen)
            If data_to_transfer <> 0& Then
                ReDim temp(0 To data_to_transfer - 1&) As Byte
                If read_from_host <> 0& Then
                    BL_DMAReadBytes segPtr, temp, data_to_transfer
                    BL_CopyBytes scsi_temp_buffer, sg_pos, temp, 0&, data_to_transfer
                ElseIf write_to_host <> 0& Then
                    BL_CopyBytes temp, 0&, scsi_temp_buffer, sg_pos, data_to_transfer
                    BL_DMAWriteBytes segPtr, temp, data_to_transfer
                End If
            End If
            sg_pos = sg_pos + data_to_transfer
            buf_len = buf_len - data_to_transfer
            If buf_len <= 0& Then Exit For
        Next i
    ElseIf (opcode = SCSI_INITIATOR_COMMAND) Or (opcode = SCSI_INITIATOR_COMMAND_RES) Then
        data_to_transfer = BL_Min(buf_len, data_length)
        If data_to_transfer <> 0& Then
            ReDim temp(0 To data_to_transfer - 1&) As Byte
            If read_from_host <> 0& Then
                BL_DMAReadBytes data_pointer, temp, data_to_transfer
                BL_CopyBytes scsi_temp_buffer, 0&, temp, 0&, data_to_transfer
            ElseIf write_to_host <> 0& Then
                BL_CopyBytes temp, 0&, scsi_temp_buffer, 0&, data_to_transfer
                BL_DMAWriteBytes data_pointer, temp, data_to_transfer
            End If
        End If
    End If
End Sub

Private Sub buslogic_sense_buffer_write(ByVal copySense As Long)
    Dim sense_len As Byte
    Dim sense_addr As Long
    Dim tempSense() As Byte

    sense_len = BL_RequestSenseLength(BL_ReqRequestSenseLength())
    If (copySense = 0&) Or (sense_len = 0&) Then Exit Sub

    scsi_device_request_sense buslogic.bus, buslogic.Req.TargetID, tempSense, sense_len
    sense_addr = BL_ReqSensePointer()
    BL_DMAWriteBytes sense_addr, tempSense, sense_len
End Sub

Private Sub buslogic_mbi_setup(ByVal host_status As Byte, ByVal target_status As Byte, ByVal mbcc As Byte)
    buslogic.Req.HostStatus = host_status
    buslogic.Req.TargetStatus = target_status
    buslogic.Req.MailboxCompletionCode = mbcc
End Sub

Private Sub buslogic_write_ccb_status()
    Dim bytes() As Byte

    If buslogic.Req.MailboxCompletionCode = MBI_NOT_FOUND Then Exit Sub

    ReDim bytes(0 To 3&) As Byte
    BL_DMAReadBytes U32Add(buslogic.Req.CCBPointer, &HC&), bytes, 4&
    bytes(2) = buslogic.Req.HostStatus
    bytes(3) = buslogic.Req.TargetStatus
    BL_DMAWriteBytes U32Add(buslogic.Req.CCBPointer, &HC&), bytes, 4&
End Sub

Private Sub buslogic_write_mbo_free()
    Dim outgoing As Long
    Dim offset As Long
    Dim freeCode() As Byte

    ReDim freeCode(0 To 0) As Byte
    freeCode(0) = MBO_FREE
    outgoing = buslogic.Outgoing
    offset = IIf((buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0&, 0&, 7&)

    If outgoing = 0& Then
        If (buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0& Then
            outgoing = U32Add(buslogic.MailboxOutAddr, (buslogic.MailboxOutPosCur * BL_MAILBOX24_SIZE))
        Else
            outgoing = U32Add(buslogic.MailboxOutAddr, (buslogic.MailboxOutPosCur * BL_MAILBOX32_SIZE))
        End If
    End If

    BL_DMAWriteBytes U32Add(outgoing, offset), freeCode, 1&
End Sub

Private Sub buslogic_write_mbi()
    Dim incoming As Long
    Dim bytes() As Byte

    If (buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0& Then
        incoming = U32Add(buslogic.MailboxInAddr, (buslogic.MailboxInPosCur * BL_MAILBOX24_SIZE))
    Else
        incoming = U32Add(buslogic.MailboxInAddr, (buslogic.MailboxInPosCur * BL_MAILBOX32_SIZE))
    End If

    If buslogic.Req.MailboxCompletionCode <> MBI_NOT_FOUND Then
        buslogic_write_ccb_status
    End If

    If (buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0& Then
        ReDim bytes(0 To BL_MAILBOX24_SIZE - 1&) As Byte
        bytes(0) = buslogic.Req.MailboxCompletionCode
        BL_WriteAddr24 bytes, 1&, buslogic.Req.CCBPointer
        BL_DMAWriteBytes incoming, bytes, BL_MAILBOX24_SIZE
    Else
        ReDim bytes(0 To BL_MAILBOX32_SIZE - 1&) As Byte
        BL_WriteU32LE bytes, 0&, buslogic.Req.CCBPointer
        bytes(4) = buslogic.Req.HostStatus
        bytes(5) = buslogic.Req.TargetStatus
        bytes(6) = 0&
        bytes(7) = buslogic.Req.MailboxCompletionCode
        BL_DMAWriteBytes incoming, bytes, BL_MAILBOX32_SIZE
    End If

    buslogic.MailboxInPosCur = buslogic.MailboxInPosCur + 1&
    If buslogic.MailboxInPosCur >= buslogic.MailboxCount Then buslogic.MailboxInPosCur = 0&

    buslogic.ToRaise = (INTR_MBIF Or INTR_ANY)
    If buslogic.MailboxOutInterrupts <> 0& Then
        buslogic.ToRaise = CByte((buslogic.ToRaise Or INTR_MBOA) And &HFF&)
    End If
End Sub
Private Function BL_BiosCmdToScsi(ByVal biosCommand As Byte) As Byte
    Select Case biosCommand
        Case &H2&
            BL_BiosCmdToScsi = GPCMD_READ_10
        Case &H3&
            BL_BiosCmdToScsi = GPCMD_WRITE_10
        Case &H4&
            BL_BiosCmdToScsi = GPCMD_VERIFY_10
        Case &H7&
            BL_BiosCmdToScsi = GPCMD_FORMAT_UNIT
        Case &HC&
            BL_BiosCmdToScsi = &H2B&
        Case &H10&
            BL_BiosCmdToScsi = GPCMD_TEST_UNIT_READY
        Case &H11&
            BL_BiosCmdToScsi = &H1&
        Case Else
            BL_BiosCmdToScsi = 0&
    End Select
End Function

Private Sub buslogic_request_sense_phase()
    Dim sense_len As Byte
    Dim sense_addr As Long

    If (buslogic.scsi_cmd_phase <> SCSI_PHASE_STATUS) And (buslogic.temp_cdb(0) = GPCMD_REQUEST_SENSE) And (BL_ReqControlByte() = &H3&) Then
        sense_len = BL_RequestSenseLength(BL_ReqRequestSenseLength())
        If (scsi_devices(buslogic.bus, buslogic.Req.TargetID).status <> SCSI_STATUS_OK) And (sense_len > 0&) Then
            sense_addr = BL_ReqSensePointer()
            BL_DMAWriteBytes sense_addr, scsi_temp_buffer, sense_len
        End If
        scsi_device_command_phase1 buslogic.bus, buslogic.Req.TargetID
    Else
        buslogic_sense_buffer_write IIf(scsi_devices(buslogic.bus, buslogic.Req.TargetID).status <> SCSI_STATUS_OK, 1&, 0&)
    End If

    buslogic_set_residue buslogic.target_data_len
    If scsi_devices(buslogic.bus, buslogic.Req.TargetID).status = SCSI_STATUS_OK Then
        buslogic_mbi_setup CCB_COMPLETE, SCSI_STATUS_OK, MBI_SUCCESS
    Else
        buslogic_mbi_setup CCB_COMPLETE, SCSI_STATUS_CHECK_CONDITION, MBI_ERROR
    End If

    buslogic.callback_sub_phase = 4&
End Sub

Private Sub buslogic_scsi_cmd()
    Dim cdb_len As Long
    Dim cdbBytes() As Byte
    Dim i As Long

    buslogic.target_data_len = buslogic_get_length(IIf(buslogic.Req.Is24Bit <> 0&, 1&, 0&))
    cdb_len = BL_ReqCdbLength()
    For i = 0& To 11&
        buslogic.temp_cdb(i) = 0&
    Next i
    For i = 0& To BL_Min(cdb_len, 12&) - 1&
        buslogic.temp_cdb(i) = buslogic.Req.CCB(18& + i)
    Next i

    BL_CopyTempCdb cdbBytes
    scsi_devices(buslogic.bus, buslogic.Req.TargetID).buffer_length = buslogic.target_data_len
    scsi_device_command_phase0 buslogic.bus, buslogic.Req.TargetID, cdbBytes
    buslogic.scsi_cmd_phase = scsi_devices(buslogic.bus, buslogic.Req.TargetID).phase
    If buslogic.scsi_cmd_phase = SCSI_PHASE_STATUS Then
        buslogic.callback_sub_phase = 3&
    Else
        buslogic.callback_sub_phase = 2&
    End If
End Sub

Private Sub buslogic_scsi_cmd_phase1()
    If buslogic.scsi_cmd_phase <> SCSI_PHASE_STATUS Then
        If (buslogic.temp_cdb(0) <> GPCMD_REQUEST_SENSE) Or (BL_ReqControlByte() <> &H3&) Then
            buslogic_buf_dma_transfer IIf(buslogic.Req.Is24Bit <> 0&, 1&, 0&), buslogic.target_data_len, IIf(buslogic.scsi_cmd_phase = SCSI_PHASE_DATA_OUT, 1&, 0&)
            scsi_device_command_phase1 buslogic.bus, buslogic.Req.TargetID
        End If
    End If

    buslogic.callback_sub_phase = 3&
End Sub

Private Sub buslogic_read_req(ByVal ccb_pointer As Long)
    Dim newReq As BUSLOGIC_REQ_t
    Dim ccb_size As Long
    Dim bytes() As Byte
    Dim i As Long

    buslogic.Req = newReq
    buslogic.Req.Is24Bit = IIf((buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0&, 1&, 0&)
    buslogic.Req.CCBPointer = ccb_pointer
    ccb_size = IIf(buslogic.Req.Is24Bit <> 0&, BL_CCB24_SIZE, BL_CCB32_SIZE)

    ReDim bytes(0 To ccb_size - 1&) As Byte
    BL_DMAReadBytes ccb_pointer, bytes, ccb_size
    For i = 0& To ccb_size - 1&
        buslogic.Req.CCB(i) = bytes(i)
    Next i

    If buslogic.Req.Is24Bit <> 0& Then
        buslogic.Req.TargetID = CByte(U32Shr(CLng(buslogic.Req.CCB(1)), 5&) And &H7&)
        buslogic.Req.LUN = CByte(buslogic.Req.CCB(1) And &H7&)
    Else
        buslogic.Req.TargetID = buslogic.Req.CCB(16)
        buslogic.Req.LUN = CByte(buslogic.Req.CCB(17) And &H1F&)
    End If
End Sub

Private Sub buslogic_req_setup(ByVal ccb_pointer As Long)
    buslogic_read_req ccb_pointer
    buslogic.target_data_len = 0&
    buslogic.scsi_cmd_phase = SCSI_PHASE_STATUS

    If (buslogic.Req.TargetID > 7&) Or (buslogic.Req.LUN > 7&) Or (scsi_device_present(buslogic.bus, buslogic.Req.TargetID) = 0&) Or (buslogic.Req.LUN > 0&) Then
        buslogic_mbi_setup CCB_SELECTION_TIMEOUT, SCSI_STATUS_OK, MBI_ERROR
        buslogic.callback_sub_phase = 4&
        Exit Sub
    End If

    scsi_device_identify buslogic.bus, buslogic.Req.TargetID, buslogic.Req.LUN
    If BL_ReqOpcode() = BUS_RESET_OPCODE Then
        scsi_device_reset buslogic.bus, buslogic.Req.TargetID
        buslogic_mbi_setup CCB_COMPLETE, SCSI_STATUS_OK, MBI_SUCCESS
        buslogic.callback_sub_phase = 4&
        Exit Sub
    End If

    If (BL_ReqOpcode() = TARGET_MODE_COMMAND) Or (BL_ReqOpcode() > SCATTER_GATHER_COMMAND_RES) Then
        buslogic_mbi_setup CCB_INVALID_OP_CODE, SCSI_STATUS_OK, MBI_ERROR
        buslogic.callback_sub_phase = 4&
        Exit Sub
    End If

    buslogic.callback_sub_phase = 1&
End Sub

Private Sub buslogic_notify()
    buslogic_write_mbo_free
    buslogic_write_mbi

    If scsi_device_present(buslogic.bus, buslogic.Req.TargetID) <> 0& Then
        scsi_device_identify buslogic.bus, buslogic.Req.TargetID, SCSI_LUN_USE_CDB
    End If

    buslogic.Outgoing = 0&
    buslogic.callback_sub_phase = 0&
    If buslogic.ToRaise <> 0& Then
        buslogic_raise_irq 0&, buslogic.ToRaise
    End If
End Sub

Private Function buslogic_process_mbo(ByVal outgoing As Long, ByVal ccbPointer As Long, ByVal mailboxCode As Byte) As Long
    buslogic.ToRaise = 0&
    buslogic.Outgoing = outgoing

    If mailboxCode = MBO_START Then
        buslogic_req_setup ccbPointer
        buslogic_process_mbo = 1&
        Exit Function
    End If

    If mailboxCode = MBO_ABORT Then
        buslogic_read_req ccbPointer
        buslogic_mbi_setup CCB_ABORTED, SCSI_STATUS_OK, MBI_NOT_FOUND
        buslogic.callback_sub_phase = 4&
        buslogic_process_mbo = 1&
        Exit Function
    End If

    buslogic_process_mbo = 0&
End Function

Private Sub buslogic_process_mail()
    Dim idx As Long
    Dim outgoing As Long
    Dim bytes() As Byte
    Dim ccbPointer As Long
    Dim mailboxCode As Byte

    If ((buslogic.Status And STAT_INIT) <> 0&) Or (buslogic.MailboxInit = 0&) Or (buslogic.MailboxCount = 0&) Or (buslogic.MailboxReq = 0&) Then Exit Sub

    If buslogic.aggressive_round_robin <> 0& Then
        For idx = 0& To buslogic.MailboxCount - 1&
            If (buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0& Then
                outgoing = U32Add(buslogic.MailboxOutAddr, (idx * BL_MAILBOX24_SIZE))
                ReDim bytes(0 To BL_MAILBOX24_SIZE - 1&) As Byte
                BL_DMAReadBytes outgoing, bytes, BL_MAILBOX24_SIZE
                mailboxCode = bytes(0)
                ccbPointer = BL_ReadAddr24(bytes, 1&)
            Else
                outgoing = U32Add(buslogic.MailboxOutAddr, (idx * BL_MAILBOX32_SIZE))
                ReDim bytes(0 To BL_MAILBOX32_SIZE - 1&) As Byte
                BL_DMAReadBytes outgoing, bytes, BL_MAILBOX32_SIZE
                ccbPointer = BL_ReadU32LE(bytes, 0&)
                mailboxCode = bytes(7)
            End If

            buslogic.MailboxOutPosCur = idx
            If buslogic_process_mbo(outgoing, ccbPointer, mailboxCode) <> 0& Then
                If buslogic.MailboxReq > 0& Then buslogic.MailboxReq = buslogic.MailboxReq - 1&
                Exit For
            End If
        Next idx
    Else
        If (buslogic.flags And BUSLOGIC_FLAG_MBX_24BIT) <> 0& Then
            outgoing = U32Add(buslogic.MailboxOutAddr, (buslogic.MailboxOutPosCur * BL_MAILBOX24_SIZE))
            ReDim bytes(0 To BL_MAILBOX24_SIZE - 1&) As Byte
            BL_DMAReadBytes outgoing, bytes, BL_MAILBOX24_SIZE
            mailboxCode = bytes(0)
            ccbPointer = BL_ReadAddr24(bytes, 1&)
        Else
            outgoing = U32Add(buslogic.MailboxOutAddr, (buslogic.MailboxOutPosCur * BL_MAILBOX32_SIZE))
            ReDim bytes(0 To BL_MAILBOX32_SIZE - 1&) As Byte
            BL_DMAReadBytes outgoing, bytes, BL_MAILBOX32_SIZE
            ccbPointer = BL_ReadU32LE(bytes, 0&)
            mailboxCode = bytes(7)
        End If

        If buslogic_process_mbo(outgoing, ccbPointer, mailboxCode) <> 0& Then
            If buslogic.MailboxReq > 0& Then buslogic.MailboxReq = buslogic.MailboxReq - 1&
            buslogic.MailboxOutPosCur = buslogic.MailboxOutPosCur + 1&
            If buslogic.MailboxOutPosCur >= buslogic.MailboxCount Then buslogic.MailboxOutPosCur = 0&
        End If
    End If
End Sub
Private Sub buslogic_cmd_done(ByVal suppress As Long)
    buslogic.DataReply = 0&
    buslogic.Status = CByte((buslogic.Status Or STAT_IDLE) And &HFF&)
    If buslogic.Command <> CMD_START_SCSI Then
        buslogic.Status = CByte((buslogic.Status And (&HFF& Xor STAT_DFULL)) And &HFF&)
        buslogic_raise_irq suppress, INTR_HACC
    End If
    buslogic.Command = &HFF&
    buslogic.CmdParam = 0&
    buslogic.CmdParamLeft = 0&
End Sub

Private Function buslogic_get_param_len(ByVal command As Byte) As Byte
    Select Case command
        Case &H25&, &H8B&, &H8C&, &H8D&, &H8F&, &H96&
            buslogic_get_param_len = 1&
        Case &H81&
            buslogic_get_param_len = BL_MAILBOXINIT32_SIZE
        Case &H83&
            buslogic_get_param_len = 12&
        Case &H90&, &H91&
            buslogic_get_param_len = 2&
        Case &HFB&
            buslogic_get_param_len = 3&
        Case Else
            buslogic_get_param_len = 0&
    End Select
End Function

Private Sub buslogic_setup_data()
    Dim i As Long

    buslogicDataBuf(0) = IIf(buslogic_get_localram_u16(89&) <> 0&, 1&, 0&)
    buslogicDataBuf(1) = buslogic.ATBusSpeed
    buslogicDataBuf(2) = buslogic.BusOnTime
    buslogicDataBuf(3) = buslogic.BusOffTime
    buslogicDataBuf(4) = CByte(buslogic.MailboxCount And &HFF&)
    buslogicDataBuf(5) = CByte(U32Shr(buslogic.MailboxOutAddr, 16&) And &HFF&)
    buslogicDataBuf(6) = CByte(U32Shr(buslogic.MailboxOutAddr, 8&) And &HFF&)
    buslogicDataBuf(7) = CByte(buslogic.MailboxOutAddr And &HFF&)
    For i = 8& To 44&
        buslogicDataBuf(i) = 0&
    Next i
    buslogicDataBuf(17) = Asc("B")
    buslogicDataBuf(18) = Asc("D")
    buslogicDataBuf(19) = Asc("A")
End Sub

Private Function buslogic_bios_scsi_command(ByVal targetId As Long, ByRef cdb() As Byte, ByRef buf() As Byte, ByVal useBuffer As Long, ByVal dataLen As Long, ByVal dmaAddr As Long) As Byte
    Dim actualLen As Long
    Dim i As Long

    scsi_devices(buslogic.bus, targetId).buffer_length = -1&
    scsi_device_command_phase0 buslogic.bus, targetId, cdb
    If scsi_devices(buslogic.bus, targetId).phase = SCSI_PHASE_STATUS Then
        buslogic_bios_scsi_command = BL_CompletionCode(targetId)
        Exit Function
    End If

    If dataLen > 0& Then
        actualLen = scsi_devices(buslogic.bus, targetId).buffer_length
        If actualLen < 0& Then actualLen = 0&
        If scsi_devices(buslogic.bus, targetId).phase = SCSI_PHASE_DATA_IN Then
            If useBuffer <> 0& Then
                For i = 0& To actualLen - 1&
                    buf(i) = scsi_temp_buffer(i)
                Next i
            Else
                BL_DMAWriteBytes dmaAddr, scsi_temp_buffer, actualLen
            End If
        ElseIf scsi_devices(buslogic.bus, targetId).phase = SCSI_PHASE_DATA_OUT Then
            If useBuffer <> 0& Then
                For i = 0& To actualLen - 1&
                    scsi_temp_buffer(i) = buf(i)
                Next i
            Else
                BL_DMAReadBytes dmaAddr, scsi_temp_buffer, actualLen
            End If
        End If
    End If

    scsi_device_command_phase1 buslogic.bus, targetId
    buslogic_bios_scsi_command = BL_CompletionCode(targetId)
End Function

Private Function buslogic_bios_read_capacity(ByVal targetId As Long, ByRef buf() As Byte) As Byte
    Dim cdb() As Byte
    Dim i As Long

    ReDim cdb(0 To 11&) As Byte
    For i = 0& To 11&
        cdb(i) = 0&
    Next i
    cdb(0) = GPCMD_READ_CDROM_CAPACITY
    ReDim buf(0 To 7&) As Byte
    For i = 0& To 7&
        buf(i) = 0&
    Next i
    buslogic_bios_read_capacity = buslogic_bios_scsi_command(targetId, cdb, buf, 1&, 8&, 0&)
End Function

Private Function buslogic_bios_inquiry(ByVal targetId As Long, ByRef buf() As Byte) As Byte
    Dim cdb() As Byte
    Dim i As Long

    ReDim cdb(0 To 11&) As Byte
    For i = 0& To 11&
        cdb(i) = 0&
    Next i
    cdb(0) = GPCMD_INQUIRY
    cdb(4) = 36&
    ReDim buf(0 To 35&) As Byte
    For i = 0& To 35&
        buf(i) = 0&
    Next i
    buslogic_bios_inquiry = buslogic_bios_scsi_command(targetId, cdb, buf, 1&, 36&, 0&)
End Function

Private Function buslogic_bios_command(ByRef cmdBytes() As Byte, ByVal islba As Long) As Byte
    Dim targetId As Long
    Dim lun As Long
    Dim dma_address As Long
    Dim lba As Long
    Dim cylinder As Long
    Dim head As Long
    Dim sec As Long
    Dim transfer_len As Long
    Dim ret As Byte
    Dim cdb() As Byte
    Dim buf() As Byte
    Dim rcbuf() As Byte
    Dim i As Long

    targetId = U32Shr(CLng(cmdBytes(1)), 5&) And &H7&
    lun = cmdBytes(1) And &H7&

    If (targetId > 7&) Or (lun <> 0&) Then
        buslogic_bios_command = &H80&
        Exit Function
    End If

    If scsi_device_present(buslogic.bus, targetId) = 0& Then
        buslogic_bios_command = &H80&
        Exit Function
    End If

    scsi_device_identify buslogic.bus, targetId, &HFF&
    If (scsi_devices(buslogic.bus, targetId).deviceType = SCSI_REMOVABLE_CDROM) And ((buslogic.flags And BUSLOGIC_FLAG_CDROM_BOOT) = 0&) Then
        buslogic_bios_command = &H80&
        Exit Function
    End If

    ReDim cdb(0 To 11&) As Byte
    For i = 0& To 11&
        cdb(i) = 0&
    Next i

    dma_address = BL_ReadAddr24(cmdBytes, 7&)
    If islba <> 0& Then
        lba = U32FromDouble(CDbl(cmdBytes(2)) * 16777216# + CDbl(cmdBytes(3)) * 65536# + CDbl(cmdBytes(4)) * 256# + CDbl(cmdBytes(5)))
    Else
        cylinder = ((CLng(cmdBytes(2)) * &H100&) Or CLng(cmdBytes(3))) And &HFFFF&
        head = cmdBytes(4) And &HF&
        sec = cmdBytes(5) And &H1F&
        lba = U32FromDouble(CDbl(cylinder) * 512# + CDbl(head) * 32# + CDbl(sec))
    End If

    Select Case cmdBytes(0)
        Case &H0&, &H6&, &H9&, &HD&, &HE&, &HF&, &H14&
            buslogic_bios_command = 0&
            Exit Function
        Case &H1&
            ReDim buf(0 To 13&) As Byte
            For i = 0& To 13&
                buf(i) = scsi_devices(buslogic.bus, targetId).sense(i)
            Next i
            BL_DMAWriteBytes dma_address, buf, 14&
            buslogic_bios_command = 0&
            Exit Function
        Case &H2&, &H3&, &H4&, &HC&
            cdb(0) = BL_BiosCmdToScsi(cmdBytes(0))
            cdb(1) = CByte((lun And &H7&) * &H20&)
            cdb(2) = CByte(U32Shr(lba, 24&) And &HFF&)
            cdb(3) = CByte(U32Shr(lba, 16&) And &HFF&)
            cdb(4) = CByte(U32Shr(lba, 8&) And &HFF&)
            cdb(5) = CByte(lba And &HFF&)
            If cmdBytes(0) <> &HC& Then
                cdb(8) = cmdBytes(6)
                If (cmdBytes(0) = &H2&) Or (cmdBytes(0) = &H3&) Then
                    transfer_len = U32FromDouble(CDbl(cmdBytes(6)) * CDbl(scsi_devices(buslogic.bus, targetId).block_len))
                Else
                    transfer_len = 0&
                End If
            Else
                transfer_len = 0&
            End If
            ret = buslogic_bios_scsi_command(targetId, cdb, buf, 0&, transfer_len, dma_address)
            If cmdBytes(0) = &HC& Then
                buslogic_bios_command = IIf(ret <> 0&, 1&, 0&)
            Else
                buslogic_bios_command = ret
            End If
            Exit Function
        Case &H7&, &H10&, &H11&
            cdb(0) = BL_BiosCmdToScsi(cmdBytes(0))
            cdb(1) = CByte((lun And &H7&) * &H20&)
            buslogic_bios_command = buslogic_bios_scsi_command(targetId, cdb, buf, 0&, cmdBytes(6), dma_address)
            Exit Function
        Case &H8&, &H15&
            If cmdBytes(0) = &H8& Then
                ret = buslogic_bios_read_capacity(targetId, rcbuf)
                If ret <> 0& Then
                    buslogic_bios_command = ret
                    Exit Function
                End If
                ReDim buf(0 To 5&) As Byte
                For i = 0& To 5&
                    buf(i) = 0&
                Next i
                For i = 0& To 3&
                    buf(i) = rcbuf(i)
                Next i
                buf(4) = rcbuf(7)
                buf(5) = rcbuf(6)
                BL_DMAWriteBytes dma_address, buf, 4&
                buslogic_bios_command = 0&
                Exit Function
            Else
                ret = buslogic_bios_inquiry(targetId, buf)
                If ret <> 0& Then
                    buslogic_bios_command = ret
                    Exit Function
                End If
                ret = buslogic_bios_read_capacity(targetId, rcbuf)
                If ret <> 0& Then
                    buslogic_bios_command = ret
                    Exit Function
                End If
                buf(4) = buf(0)
                buf(5) = buf(1)
                For i = 0& To 3&
                    buf(i) = rcbuf(i)
                Next i
                BL_DMAWriteBytes dma_address, buf, 4&
                buslogic_bios_command = 0&
                Exit Function
            End If
        Case Else
            buslogic_bios_command = &H1&
            Exit Function
    End Select
End Function
Private Sub buslogic_bios_dma_transfer(ByRef cmdBytes() As Byte, ByVal targetId As Long, ByVal dirDataOut As Long)
    Dim data_direction As Long
    Dim transfer_length As Long
    Dim dataLength As Long
    Dim dataPointer As Long

    data_direction = U32Shr(CLng(cmdBytes(10)), 3&) And &H3&
    dataLength = BL_ReadU32LE(cmdBytes, 0&)
    dataPointer = BL_ReadU32LE(cmdBytes, 4&)

    If (data_direction = &H3&) Or (dataLength = 0&) Or (scsi_devices(buslogic.bus, targetId).buffer_length <= 0&) Then Exit Sub

    transfer_length = BL_Min(dataLength, scsi_devices(buslogic.bus, targetId).buffer_length)
    If (dirDataOut <> 0&) And ((data_direction = CCB_DATA_XFER_OUT) Or (data_direction = 0&)) Then
        BL_DMAReadBytes dataPointer, scsi_temp_buffer, transfer_length
    ElseIf (dirDataOut = 0&) And ((data_direction = CCB_DATA_XFER_IN) Or (data_direction = 0&)) Then
        BL_DMAWriteBytes dataPointer, scsi_temp_buffer, transfer_length
    End If
End Sub

Private Sub buslogic_bios_request_setup(ByRef cmdBytes() As Byte, ByVal reply_len As Long)
    Dim targetId As Long
    Dim logicalUnit As Long
    Dim cdbLength As Long
    Dim phase As Byte
    Dim temp_cdb() As Byte
    Dim dataLength As Long
    Dim i As Long

    For i = 0& To 3&
        buslogicDataBuf(i) = 0&
    Next i

    targetId = cmdBytes(8)
    logicalUnit = cmdBytes(9)
    cdbLength = cmdBytes(11)
    If (targetId > 15&) Or (logicalUnit > 7&) Then
        buslogicDataBuf(2) = CCB_INVALID_CCB
        buslogicDataBuf(3) = SCSI_STATUS_OK
        buslogic.DataReplyLeft = reply_len
        Exit Sub
    End If
    If (scsi_device_present(buslogic.bus, targetId) = 0&) Or (logicalUnit > 0&) Then
        buslogicDataBuf(2) = CCB_SELECTION_TIMEOUT
        buslogicDataBuf(3) = SCSI_STATUS_OK
        buslogic.DataReplyLeft = reply_len
        Exit Sub
    End If

    scsi_device_identify buslogic.bus, targetId, logicalUnit
    ReDim temp_cdb(0 To 11&) As Byte
    For i = 0& To 11&
        temp_cdb(i) = 0&
    Next i
    For i = 0& To BL_Min(cdbLength, 12&) - 1&
        temp_cdb(i) = cmdBytes(12& + i)
    Next i
    dataLength = BL_ReadU32LE(cmdBytes, 0&)
    scsi_devices(buslogic.bus, targetId).buffer_length = dataLength
    scsi_device_command_phase0 buslogic.bus, targetId, temp_cdb
    phase = scsi_devices(buslogic.bus, targetId).phase
    If phase <> SCSI_PHASE_STATUS Then
        buslogic_bios_dma_transfer cmdBytes, targetId, IIf(phase = SCSI_PHASE_DATA_OUT, 1&, 0&)
        scsi_device_command_phase1 buslogic.bus, targetId
    End If
    scsi_device_identify buslogic.bus, targetId, SCSI_LUN_USE_CDB
    buslogicDataBuf(2) = CCB_COMPLETE
    buslogicDataBuf(3) = scsi_devices(buslogic.bus, targetId).status
    buslogic.DataReplyLeft = reply_len
End Sub

Private Function buslogic_vendor_cmd() As Long
    Dim targets_present_mask As Long
    Dim offset As Long
    Dim count As Long
    Dim i As Long

    Select Case buslogic.Command
        Case &H20&
            buslogic_reset 1&
            buslogic.DataReplyLeft = 0&
        Case &H23&
            For i = 0& To 7&
                buslogicDataBuf(i) = 0&
            Next i
            buslogic.DataReplyLeft = 8&
        Case &H24&
            targets_present_mask = 0&
            For i = 0& To 14&
                If (i <> buslogic.HostID) And (scsi_device_present(buslogic.bus, i) <> 0&) Then
                    targets_present_mask = (targets_present_mask Or U32Shl(1&, i))
                End If
            Next i
            buslogicDataBuf(0) = CByte(targets_present_mask And &HFF&)
            buslogicDataBuf(1) = CByte(U32Shr(targets_present_mask, 8&) And &HFF&)
            buslogic.DataReplyLeft = 2&
        Case &H25&
            buslogic.IrqEnabled = IIf(buslogic.CmdBuf(0) <> 0&, 1&, 0&)
            buslogic.DataReplyLeft = 0&
            buslogic_vendor_cmd = 1&
            Exit Function
        Case &H81&
            buslogic.flags = CByte((buslogic.flags And (&HFF& Xor BUSLOGIC_FLAG_MBX_24BIT)) And &HFF&)
            buslogic.MailboxInit = 1&
            buslogic.MailboxCount = buslogic.CmdBuf(0)
            buslogic.MailboxOutAddr = BL_ReadU32LE_CmdBuf(1&)
            buslogic.MailboxInAddr = U32Add(buslogic.MailboxOutAddr, (buslogic.MailboxCount * BL_MAILBOX32_SIZE))
            buslogic.Status = CByte((buslogic.Status And (&HFF& Xor STAT_INIT)) And &HFF&)
            buslogic.DataReplyLeft = 0&
        Case &H83&
            If buslogic.CmdParam = 12& Then
                buslogic.CmdParamLeft = buslogic.CmdBuf(11)
                buslogic_vendor_cmd = 0&
                Exit Function
            End If
            Dim biosReq() As Byte
            BL_CopyCmdBuf biosReq, 24&
            buslogic_bios_request_setup biosReq, 4&
        Case &H84&
            buslogicDataBuf(0) = buslogic.fw_rev(4)
            buslogic.DataReplyLeft = 1&
        Case &H85&
            If buslogic.fw_rev(5) <> 0& Then
                buslogicDataBuf(0) = buslogic.fw_rev(5)
            Else
                buslogicDataBuf(0) = Asc(" ")
            End If
            buslogic.DataReplyLeft = 1&
        Case &H8B&
            buslogic.DataReplyLeft = buslogic.CmdBuf(0)
            For i = 0& To buslogic.DataReplyLeft - 1&
                buslogicDataBuf(i) = 0&
            Next i
            If buslogic.DataReplyLeft > 0& Then buslogicDataBuf(0) = Asc("5")
            If buslogic.DataReplyLeft > 1& Then buslogicDataBuf(1) = Asc("4")
            If buslogic.DataReplyLeft > 2& Then buslogicDataBuf(2) = Asc("5")
            If buslogic.DataReplyLeft > 3& Then buslogicDataBuf(3) = Asc("S")
        Case &H8C&
            buslogic.DataReplyLeft = buslogic.CmdBuf(0)
            For i = 0& To buslogic.DataReplyLeft - 1&
                buslogicDataBuf(i) = 0&
            Next i
        Case &H8D&
            buslogic.DataReplyLeft = buslogic.CmdBuf(0)
            For i = 0& To buslogic.DataReplyLeft - 1&
                buslogicDataBuf(i) = 0&
            Next i
            If buslogic.DataReplyLeft > 0& Then buslogicDataBuf(0) = Asc("A")
            If buslogic.DataReplyLeft > 1& Then buslogicDataBuf(1) = CByte(U32Shr(buslogic.rom_addr, 12&) And &HFF&)
            If buslogic.DataReplyLeft > 3& Then BL_WriteU16LE_DataBuf 2&, 8192&
            If buslogic.DataReplyLeft > 4& Then buslogicDataBuf(4) = CByte(buslogic.MailboxCount And &HFF&)
            If buslogic.DataReplyLeft > 8& Then BL_WriteU32LE_DataBuf 5&, buslogic.MailboxOutAddr
            If buslogic.DataReplyLeft > 10& Then buslogicDataBuf(10) = buslogic.fw_rev(2)
            If buslogic.DataReplyLeft > 11& Then buslogicDataBuf(11) = buslogic.fw_rev(3)
            If buslogic.DataReplyLeft > 12& Then buslogicDataBuf(12) = buslogic.fw_rev(4)
        Case &H8F&
            buslogic.aggressive_round_robin = CByte(buslogic.CmdBuf(0) And &H1&)
            buslogic.DataReplyLeft = 0&
        Case &H90&
            offset = buslogic.CmdBuf(0)
            count = BL_Min(buslogic.CmdBuf(1), (256& - BL_Min(offset, 256&)))
            If count <> 0& Then
                For i = 0& To count - 1&
                    buslogic.LocalRAM(offset + i) = buslogic.CmdBuf(2& + i)
                Next i
                buslogic_save_nvr
            End If
            buslogic.DataReplyLeft = 0&
        Case &H91&
            offset = buslogic.CmdBuf(0)
            count = BL_Min(buslogic.CmdBuf(1), (256& - BL_Min(offset, 256&)))
            buslogic.DataReplyLeft = count
            If count <> 0& Then
                For i = 0& To count - 1&
                    buslogicDataBuf(i) = buslogic.LocalRAM(offset + i)
                Next i
            End If
        Case &H96&
            buslogic.ExtendedLUNCCBFormat = IIf(buslogic.CmdBuf(0) <> 0&, 1&, 0&)
            buslogic.DataReplyLeft = 0&
        Case &HFB&
            buslogic.DataReplyLeft = buslogic.CmdBuf(2)
        Case Else
            buslogic.Status = CByte((buslogic.Status Or STAT_INVCMD) And &HFF&)
            buslogic.DataReplyLeft = 0&
    End Select

    buslogic_vendor_cmd = 0&
End Function
Private Sub buslogic_handle_command()
    Dim fifo_buf As Long
    Dim id As Long
    Dim suppress As Long
    Dim replyLen As Long
    Dim tempBuf() As Byte

    suppress = 0&

    Select Case buslogic.Command
        Case CMD_NOP
            buslogic.DataReplyLeft = 0&
        Case CMD_MBINIT
            buslogic.flags = CByte((buslogic.flags Or BUSLOGIC_FLAG_MBX_24BIT) And &HFF&)
            buslogic.MailboxInit = 1&
            buslogic.MailboxCount = buslogic.CmdBuf(0)
            buslogic.MailboxOutAddr = BL_ReadAddr24_CmdBuf(1&)
            buslogic.MailboxInAddr = U32Add(buslogic.MailboxOutAddr, (buslogic.MailboxCount * BL_MAILBOX24_SIZE))
            buslogic.Status = CByte((buslogic.Status And (&HFF& Xor STAT_INIT)) And &HFF&)
            buslogic.DataReplyLeft = 0&
        Case CMD_BIOSCMD
            Dim biosCmd() As Byte
            BL_CopyCmdBuf biosCmd, 10&
            buslogicDataBuf(0) = buslogic_bios_command(biosCmd, 0&)
            buslogic.DataReplyLeft = 1&
        Case CMD_INQUIRY
            buslogicDataBuf(0) = buslogic.fw_rev(0)
            buslogicDataBuf(1) = buslogic.fw_rev(1)
            buslogicDataBuf(2) = buslogic.fw_rev(2)
            buslogicDataBuf(3) = buslogic.fw_rev(3)
            buslogic.DataReplyLeft = 4&
        Case CMD_EMBOI
            If buslogic.CmdBuf(0) <= 1& Then
                buslogic.MailboxOutInterrupts = buslogic.CmdBuf(0)
                suppress = 1&
            Else
                buslogic.Status = CByte((buslogic.Status Or STAT_INVCMD) And &HFF&)
            End If
            buslogic.DataReplyLeft = 0&
        Case CMD_SELTIMEOUT
            buslogic.DataReplyLeft = 0&
        Case CMD_BUSON_TIME
            buslogic.BusOnTime = buslogic.CmdBuf(0)
            buslogic.DataReplyLeft = 0&
        Case CMD_BUSOFF_TIME
            buslogic.BusOffTime = buslogic.CmdBuf(0)
            buslogic.DataReplyLeft = 0&
        Case CMD_DMASPEED
            buslogic.ATBusSpeed = buslogic.CmdBuf(0)
            buslogic.DataReplyLeft = 0&
        Case CMD_RETDEVS
            For id = 0& To 7&
                buslogicDataBuf(id) = 0&
            Next id
            For id = 0& To 7&
                If (id <> buslogic.HostID) And (scsi_device_present(buslogic.bus, id) <> 0&) Then
                    buslogicDataBuf(id) = 1&
                End If
            Next id
            buslogic.DataReplyLeft = 8&
        Case CMD_RETCONF
            buslogicDataBuf(0) = CByte(U32Shl(1&, buslogic.DmaChannel) And &HFF&)
            If buslogic.Irq >= 9& Then
                buslogicDataBuf(1) = CByte(U32Shl(1&, buslogic.Irq - 9&) And &HFF&)
            Else
                buslogicDataBuf(1) = 0&
            End If
            buslogicDataBuf(2) = buslogic.HostID
            buslogic.DataReplyLeft = 3&
        Case CMD_RETSETUP
            replyLen = buslogic.CmdBuf(0)
            For id = 0& To replyLen - 1&
                buslogicDataBuf(id) = 0&
            Next id
            If replyLen > 1& Then buslogicDataBuf(1) = buslogic.ATBusSpeed
            If replyLen > 2& Then buslogicDataBuf(2) = buslogic.BusOnTime
            If replyLen > 3& Then buslogicDataBuf(3) = buslogic.BusOffTime
            If replyLen > 4& Then buslogicDataBuf(4) = CByte(buslogic.MailboxCount And &HFF&)
            If replyLen > 7& Then BL_WriteAddr24_DataBuf 5&, buslogic.MailboxOutAddr
            buslogic_setup_data
            buslogic.DataReplyLeft = replyLen
        Case CMD_ECHO
            buslogicDataBuf(0) = buslogic.CmdBuf(0)
            buslogic.DataReplyLeft = 1&
        Case CMD_WRITE_CH2
            fifo_buf = BL_ReadAddr24_CmdBuf(0&)
            ReDim tempBuf(0 To 63&) As Byte
            BL_DMAReadBytes fifo_buf, tempBuf, 64&
            For id = 0& To 63&
                buslogic.dma_buffer(id) = tempBuf(id)
            Next id
            buslogic.DataReplyLeft = 0&
        Case CMD_READ_CH2
            fifo_buf = BL_ReadAddr24_CmdBuf(0&)
            ReDim tempBuf(0 To 63&) As Byte
            For id = 0& To 63&
                tempBuf(id) = buslogic.dma_buffer(id)
            Next id
            BL_DMAWriteBytes fifo_buf, tempBuf, 64&
            buslogic.DataReplyLeft = 0&
        Case CMD_OPTIONS
            If buslogic.CmdParam = 1& Then
                buslogic.CmdParamLeft = buslogic.CmdBuf(0)
            End If
            buslogic.DataReplyLeft = 0&
        Case Else
            suppress = buslogic_vendor_cmd()
    End Select

    If buslogic.DataReplyLeft <> 0& Then
        buslogic.Status = CByte((buslogic.Status Or STAT_DFULL) And &HFF&)
    ElseIf buslogic.CmdParamLeft = 0& Then
        buslogic_cmd_done suppress
    End If
End Sub

Private Function buslogic_port_read(ByVal port As Long) As Byte
    Dim ret As Byte

    ret = &HFF&
    Select Case (port And &H3&)
        Case 0&
            ret = buslogic.Status
        Case 1&
            If buslogic.DataReplyLeft <> 0& Then
                ret = buslogicDataBuf(buslogic.DataReply)
                buslogic.DataReply = buslogic.DataReply + 1&
                buslogic.DataReplyLeft = buslogic.DataReplyLeft - 1&
                If buslogic.DataReplyLeft = 0& Then
                    buslogic_cmd_done 0&
                End If
            Else
                ret = 0&
            End If
        Case 2&
            ret = buslogic.Interrupt
        Case 3&
            ret = buslogic.Geometry
    End Select

    buslogic_port_read = ret
End Function

Private Sub buslogic_port_write(ByVal port As Long, ByVal value As Byte)
    Dim i As Long

    Select Case (port And &H3&)
        Case 0&
            If ((value And CTRL_HRST) <> 0&) Or ((value And CTRL_SRST) <> 0&) Then
                buslogic_reset 1&
                Exit Sub
            End If
            If (value And CTRL_SCRST) <> 0& Then
                For i = 0& To SCSI_ID_MAX - 1&
                    scsi_device_reset buslogic.bus, i
                Next i
            End If
            If (value And CTRL_IRST) <> 0& Then
                buslogic_clear_irq
            End If
            Exit Sub
        Case 1&
            If (value = CMD_START_SCSI) And (buslogic.Command = &HFF&) Then
                buslogic.MailboxReq = buslogic.MailboxReq + 1&
                buslogic_schedule_mail
                Exit Sub
            End If

            If buslogic.Command = &HFF& Then
                buslogic.Command = value
                buslogic.CmdParam = 0&
                buslogic.CmdParamLeft = 0&
                buslogic.Status = CByte((buslogic.Status And (&HFF& Xor (STAT_INVCMD Or STAT_IDLE))) And &HFF&)
                Select Case buslogic.Command
                    Case CMD_MBINIT
                        buslogic.CmdParamLeft = BL_MAILBOXINIT24_SIZE
                    Case CMD_BIOSCMD
                        buslogic.CmdParamLeft = 10&
                    Case CMD_EMBOI, CMD_BUSON_TIME, CMD_BUSOFF_TIME, CMD_DMASPEED, CMD_RETSETUP, CMD_ECHO, CMD_OPTIONS
                        buslogic.CmdParamLeft = 1&
                    Case CMD_SELTIMEOUT
                        buslogic.CmdParamLeft = 4&
                    Case CMD_WRITE_CH2, CMD_READ_CH2
                        buslogic.CmdParamLeft = 3&
                    Case Else
                        buslogic.CmdParamLeft = buslogic_get_param_len(buslogic.Command)
                End Select
            Else
                buslogic.CmdBuf(buslogic.CmdParam) = value
                buslogic.CmdParam = CByte((buslogic.CmdParam + 1&) And &HFF&)
                If buslogic.CmdParamLeft > 0& Then buslogic.CmdParamLeft = buslogic.CmdParamLeft - 1&
                buslogic_cmd_phase1
            End If

            If buslogic.CmdParamLeft = 0& Then
                buslogic_handle_command
            End If
            Exit Sub
        Case 2&
            If (buslogic.flags And BUSLOGIC_FLAG_INT_GEOM_WRITABLE) <> 0& Then
                buslogic.Interrupt = value
            End If
            Exit Sub
        Case 3&
            If (buslogic.flags And BUSLOGIC_FLAG_INT_GEOM_WRITABLE) <> 0& Then
                buslogic.Geometry = value
            End If
            Exit Sub
    End Select
End Sub

Public Function buslogic_readport(ByVal udata As Long, ByVal portnum As Integer) As Byte
    If buslogic.initialized = 0& Then
        buslogic_readport = &HFF&
    Else
        buslogic_readport = buslogic_port_read(portnum And &HFFFF&)
    End If
End Function

Public Sub buslogic_writeport(ByVal udata As Long, ByVal portnum As Integer, ByVal value As Byte)
    If buslogic.initialized = 0& Then Exit Sub
    buslogic_port_write portnum And &HFFFF&, value
End Sub

Public Sub buslogic_process_mail_cb(ByVal data As Long)
    timing_timerDisable buslogic.mail_timer_id
    Select Case buslogic.callback_sub_phase
        Case 0&
            buslogic_process_mail
        Case 1&
            buslogic_scsi_cmd
        Case 2&
            buslogic_scsi_cmd_phase1
        Case 3&
            buslogic_request_sense_phase
        Case 4&
            buslogic_notify
        Case Else
            buslogic.callback_sub_phase = 0&
    End Select

    If (buslogic.callback_sub_phase <> 0&) Or (buslogic.MailboxReq <> 0&) Then
        buslogic_schedule_mail
    End If
End Sub

Public Sub buslogic_reset_timer_cb(ByVal data As Long)
    timing_timerDisable buslogic.reset_timer_id
    buslogic.Status = (STAT_INIT Or STAT_IDLE)
End Sub
Public Function buslogic_readrom(ByVal udata As Long, ByVal addr32 As Long) As Byte
    If (buslogic.rom_loaded = 0&) Or (U32Lt(addr32, buslogic.rom_addr) <> 0&) Or (U32Lt(addr32, U32Add(buslogic.rom_addr, ROM_SIZE)) = 0&) Then
        buslogic_readrom = &HFF&
    Else
        buslogic_readrom = buslogicRomData(U32Sub(addr32, buslogic.rom_addr))
    End If
End Function

Public Sub buslogic_writerom(ByVal udata As Long, ByVal addr32 As Long, ByVal value As Byte)
End Sub

Public Function buslogic_init(ByRef machineArg As MACHINE_t, ByVal picSlave As Long) As Long
    Dim i As Long
    Dim fn As Integer
    Dim romLen As Long
    Dim romPath As String

    If machineArg.buslogic_enabled = 0& Then
        buslogic_init = 0&
        Exit Function
    End If

    If (machineArg.buslogic_irq >= 8&) And (picSlave < 0&) Then
        debug_log DEBUG_ERROR, "[SCSI] BusLogic IRQ " & CStr(machineArg.buslogic_irq) & " requires a slave PIC on this machine" & vbCrLf
        buslogic_init = -1&
        Exit Function
    End If

    scsi_reset
    scsi_device_init

    buslogic.initialized = 0&
    buslogic.pic_slave = picSlave
    buslogic.Base = machineArg.buslogic_base
    buslogic.Irq = machineArg.buslogic_irq
    buslogic.DmaChannel = machineArg.buslogic_dma
    buslogic.HostID = 7&
    buslogic.transfer_size = 2&
    buslogic.aggressive_round_robin = 1&
    buslogic.flags = (BUSLOGIC_FLAG_INT_GEOM_WRITABLE Or BUSLOGIC_FLAG_MBX_24BIT Or BUSLOGIC_FLAG_CDROM_BOOT)
    buslogic.fw_rev(0) = Asc("A")
    buslogic.fw_rev(1) = Asc("A")
    buslogic.fw_rev(2) = Asc("3")
    buslogic.fw_rev(3) = Asc("3")
    buslogic.fw_rev(4) = Asc("1")
    buslogic.fw_rev(5) = 0&
    buslogic.fw_rev(6) = 0&
    buslogic.fw_rev(7) = 0&
    buslogic.BusOnTime = 7&
    buslogic.BusOffTime = 4&
    buslogic.ATBusSpeed = 1&
    buslogic.bus = scsi_get_bus()
    buslogic.rom_addr = machineArg.buslogic_bios_addr
    buslogic.rom_path = machineArg.buslogic_rom_path
    buslogic.nvr_path = machineArg.buslogic_nvr_path

    If buslogic.bus = &HFF& Then
        buslogic_init = -1&
        Exit Function
    End If
    ReDim buslogicDataBuf(0 To 65535&) As Byte
    ReDim buslogicRomData(0 To ROM_SIZE - 1&) As Byte
    scsi_bus_set_speed buslogic.bus, 5000000#
    buslogic_load_nvr

    If buslogic.rom_addr <> 0& Then
        romPath = BL_FixedString(buslogic.rom_path)
        On Error GoTo RomFail
        fn = FreeFile
        Open romPath For Binary Access Read As #fn
        romLen = LOF(fn)
        If romLen < ROM_SIZE Then GoTo RomFail
        Get #fn, 1, buslogicRomData
        Close #fn
        buslogic.rom_loaded = 1&
        memory_mapCallbackRegister buslogic.rom_addr, ROM_SIZE, MEMORY_CB_BUSLOGIC_ROM, MEMORY_CB_BUSLOGIC_ROM, 0&
    End If

    For i = 0& To BUSLOGIC_MAX_TARGETS - 1&
        If machineArg.scsi_targets(i).present = 0& Then GoTo NextTarget
        If machineArg.scsi_targets(i).targetType = BUSLOGIC_TARGET_DISK Then
            If scsi_disk_attach(buslogic.bus, i, BL_FixedString(machineArg.scsi_targets(i).path)) <> 0& Then
                debug_log DEBUG_ERROR, "[SCSI] Failed to attach disk target " & CStr(i) & vbCrLf
                buslogic_init = -1&
                Exit Function
            End If
        ElseIf machineArg.scsi_targets(i).targetType = BUSLOGIC_TARGET_CDROM Then
            If LenB(BL_FixedString(machineArg.scsi_targets(i).path)) = 0& Then
                scsi_cdrom_eject buslogic.bus, i
            ElseIf scsi_cdrom_attach(buslogic.bus, i, BL_FixedString(machineArg.scsi_targets(i).path)) <> 0& Then
                debug_log DEBUG_ERROR, "[SCSI] Failed to attach CD-ROM target " & CStr(i) & vbCrLf
                buslogic_init = -1&
                Exit Function
            End If
        End If
NextTarget:
    Next i

    buslogic.mail_timer_id = timing_addTimer(TIMER_CB_BUSLOGIC_MAIL, 0&, 1000#, TIMING_DISABLED)
    buslogic.reset_timer_id = timing_addTimer(TIMER_CB_BUSLOGIC_RESET, 0&, 1000#, TIMING_DISABLED)
    buslogic_register_io
    buslogic_reset 1&
    buslogic.initialized = 1&

    debug_log DEBUG_INFO, "[SCSI] BusLogic BT-545S initialized at 0x" & Hex$(buslogic.Base) & " IRQ " & CStr(buslogic.Irq) & " DMA " & CStr(buslogic.DmaChannel) & " BIOS " & IIf(buslogic.rom_addr <> 0&, "enabled", "disabled") & vbCrLf
    buslogic_init = 0&
    Exit Function

RomFail:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    On Error GoTo 0
    debug_log DEBUG_ERROR, "[SCSI] BusLogic BIOS ROM not found or invalid: " & romPath & vbCrLf
    buslogic_init = -1&
End Function

Public Function buslogic_is_initialized() As Long
    buslogic_is_initialized = buslogic.initialized
End Function

Public Function buslogic_get_bus_id() As Long
    If buslogic.initialized = 0& Then
        buslogic_get_bus_id = -1&
    Else
        buslogic_get_bus_id = buslogic.bus
    End If
End Function
