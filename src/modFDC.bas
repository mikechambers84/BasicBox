Attribute VB_Name = "modFDC"
Option Explicit

Private Const FDC_IRQ As Byte = 6&
Private Const FDC_DMA_CH As Byte = 2&

Private Const FDC_PORT_DOR As Long = 2&
Private Const FDC_PORT_MSR As Long = 4&
Private Const FDC_PORT_DATA As Long = 5&
Private Const FDC_PORT_DIR As Long = 7&

Private Const FDC_CMD_MODE As Byte = 1&
Private Const FDC_CMD_READ_TRACK As Byte = 2&
Private Const FDC_CMD_SPECIFY As Byte = 3&
Private Const FDC_CMD_SENSE_DRIVE_STATUS As Byte = 4&
Private Const FDC_CMD_WRITE_DATA As Byte = 5&
Private Const FDC_CMD_READ_DATA As Byte = 6&
Private Const FDC_CMD_RECALIBRATE As Byte = 7&
Private Const FDC_CMD_SENSE_INTERRUPT As Byte = 8&
Private Const FDC_CMD_WRITE_DELETED As Byte = 9&
Private Const FDC_CMD_READ_ID As Byte = 10&
Private Const FDC_CMD_READ_DELETED As Byte = 12&
Private Const FDC_CMD_FORMAT_TRACK As Byte = 13&
Private Const FDC_CMD_DUMPREG As Byte = 14&
Private Const FDC_CMD_SEEK As Byte = 15&
Private Const FDC_CMD_VERSION As Byte = 16&
Private Const FDC_CMD_SCAN_EQUAL As Byte = 17&
Private Const FDC_CMD_PERPENDICULAR As Byte = 18&
Private Const FDC_CMD_CONFIGURE As Byte = 19&
Private Const FDC_CMD_UNLOCK As Byte = 20&
Private Const FDC_CMD_VERIFY As Byte = 22&
Private Const FDC_CMD_POWERDOWN As Byte = 23&
Private Const FDC_CMD_NSC As Byte = 24&
Private Const FDC_CMD_SCAN_LOW_EQUAL As Byte = 25&
Private Const FDC_CMD_SCAN_HIGH_EQUAL As Byte = 29&
Private Const FDC_CMD_LOCK As Byte = &H94&

Private Const FDC_ST0_HEAD As Byte = &H4&
Private Const FDC_ST0_NOT_READY As Byte = &H8&
Private Const FDC_ST0_SEEK_END As Byte = &H20&
Private Const FDC_ST0_ABNORMAL As Byte = &H40&
Private Const FDC_ST0_INVALID As Byte = &H80&

Private Const FDC_ST1_NO_ID As Byte = &H1&
Private Const FDC_ST1_WRITE_PROTECT As Byte = &H2&
Private Const FDC_ST1_NO_DATA As Byte = &H4&

Private Const FDC_ST2_SCAN_NOT_SATISFIED As Byte = &H4&
Private Const FDC_ST2_SCAN_EQUAL_HIT As Byte = &H8&

Private Const FDC_ST3_TWO_SIDE As Byte = &H8&
Private Const FDC_ST3_TRACK0 As Byte = &H10&
Private Const FDC_ST3_READY As Byte = &H20&
Private Const FDC_ST3_WRITE_PROTECT As Byte = &H40&

Private Const FDC_EXEC_NONE As Byte = 0&
Private Const FDC_EXEC_RESET As Byte = 1&
Private Const FDC_EXEC_READ As Byte = 2&
Private Const FDC_EXEC_WRITE As Byte = 3&
Private Const FDC_EXEC_READID As Byte = 4&
Private Const FDC_EXEC_FORMAT As Byte = 5&
Private Const FDC_EXEC_SEEK As Byte = 6&
Private Const FDC_EXEC_RECAL As Byte = 7&
Private Const FDC_EXEC_READTRACK As Byte = 8&
Private Const FDC_EXEC_VERIFY As Byte = 9&
Private Const FDC_EXEC_SCAN As Byte = 10&

Private Const FDC_SCAN_NONE As Byte = 0&
Private Const FDC_SCAN_EQUAL_MODE As Byte = 1&
Private Const FDC_SCAN_LOW_EQUAL_MODE As Byte = 2&
Private Const FDC_SCAN_HIGH_EQUAL_MODE As Byte = 3&

Private Const FDC_EXEC_DELAY_USEC As Long = 2000&
Private Const FDC_RESET_DELAY_USEC As Long = 500&
Private Const FDC_DMA_BUFFER_NONE As Byte = 0&
Private Const FDC_DMA_BUFFER_SECTOR As Byte = 1&
Private Const FDC_DMA_BUFFER_FORMAT As Byte = 2&

Private Type FDC_t
    inited As Byte
    i8259Slot As Long
    dor As Byte
    msr As Byte
    dsr As Byte
    rate As Byte
    irqPending As Byte
    dmaEnabled As Byte
    command As Byte
    commandBase As Byte
    params(0& To 15&) As Byte
    paramCount As Long
    paramExpected As Long
    results(0& To 15&) As Byte
    resultCount As Long
    resultIndex As Long
    pcn(0& To 3&) As Byte
    selectedDrive As Byte
    commandDrive As Byte
    rwDrive As Byte
    fintr As Byte
    resetPending As Byte
    lastSenseSt0 As Byte
    lastSenseDrive As Byte
    execState As Byte
    waitingSeek As Byte
    opTrack As Byte
    opHead As Byte
    opSector As Byte
    opSizeCode As Byte
    opEot As Byte
    eot(0& To 3&) As Byte
    opGap As Byte
    opDtl As Byte
    specify(0& To 1&) As Byte
    denselForce As Byte
    perp As Byte
    config As Byte
    pretrk As Byte
    controllerLock As Byte
    powerDownMode As Byte
    compareMode As Byte
    compareSatisfied As Long
    compareTotal As Long
    verifyRemaining As Long
    formatFill As Byte
    formatSectorCount As Byte
    dmaBufferKind As Byte
    dmaLength As Long
    dmaLastTransferred As Long
    timerController As Long
End Type

Private fdc As FDC_t
Private fdc_dmaSectorBuffer(0& To 511&) As Byte
Private fdc_dmaFormatBuffer(0& To 255&) As Byte

Private Function FDC_BusyMask(ByVal drive As Long) As Byte
    Select Case drive
        Case 0&: FDC_BusyMask = 1&
        Case 1&: FDC_BusyMask = 2&
        Case 2&: FDC_BusyMask = 4&
        Case Else: FDC_BusyMask = 8&
    End Select
End Function

Private Function FDC_MotorMask(ByVal drive As Long) As Byte
    Select Case drive
        Case 0&: FDC_MotorMask = &H10&
        Case 1&: FDC_MotorMask = &H20&
        Case 2&: FDC_MotorMask = &H40&
        Case Else: FDC_MotorMask = &H80&
    End Select
End Function

Private Function FDC_HeadBit(ByVal head As Long) As Byte
    If head <> 0& Then
        FDC_HeadBit = FDC_ST0_HEAD
    Else
        FDC_HeadBit = 0&
    End If
End Function

Private Function FDC_PhysicalDrive() As Long
    FDC_PhysicalDrive = (fdc.commandDrive And 3&)
End Function

Private Function FDC_IsMotorOn(ByVal drive As Long) As Boolean
    FDC_IsMotorOn = ((fdc.dor And FDC_MotorMask(drive)) <> 0&)
End Function

Private Sub FDC_ClearDrq()
    i8237_set_drq FDC_DMA_CH, 0&
End Sub

Private Sub FDC_ResetDmaState()
    fdc.dmaBufferKind = FDC_DMA_BUFFER_NONE
    fdc.dmaLength = 0&
    fdc.dmaLastTransferred = 0&
End Sub

Private Function FDC_RunDmaRequest() As Byte
    Dim i As Long
    Dim tc As Byte

    fdc.dmaLastTransferred = 0&
    If fdc.dmaLength <= 0& Then Exit Function

    i8237_clear_terminal FDC_DMA_CH

    Select Case fdc.execState
        Case FDC_EXEC_READ, FDC_EXEC_READTRACK
            For i = 0& To fdc.dmaLength - 1&
                tc = i8237_fdc_writeByte(FDC_DMA_CH, fdc_dmaSectorBuffer(i))
                fdc.dmaLastTransferred = i + 1&
                If tc <> 0& Then Exit For
            Next i

        Case FDC_EXEC_WRITE, FDC_EXEC_SCAN, FDC_EXEC_FORMAT
            If fdc.dmaBufferKind = FDC_DMA_BUFFER_FORMAT Then
                For i = 0& To fdc.dmaLength - 1&
                    tc = i8237_fdc_readByte(FDC_DMA_CH, fdc_dmaFormatBuffer(i))
                    fdc.dmaLastTransferred = i + 1&
                    If tc <> 0& Then Exit For
                Next i
            Else
                For i = 0& To fdc.dmaLength - 1&
                    tc = i8237_fdc_readByte(FDC_DMA_CH, fdc_dmaSectorBuffer(i))
                    fdc.dmaLastTransferred = i + 1&
                    If tc <> 0& Then Exit For
                Next i
            End If
    End Select

    FDC_RunDmaRequest = i8237_get_terminal(FDC_DMA_CH)
End Function

Private Sub FDC_LowerIRQ()
    If (fdc.i8259Slot >= 0&) And (fdc.irqPending <> 0&) Then
        i8259_clearirq fdc.i8259Slot, FDC_IRQ
    End If
    fdc.irqPending = 0&
End Sub

Private Sub FDC_RaiseIRQ()
    If (fdc.dor And &H8&) = 0& Then Exit Sub
    If fdc.i8259Slot < 0& Then Exit Sub

    i8259_doirq fdc.i8259Slot, FDC_IRQ
    fdc.irqPending = 1&
End Sub

Private Sub FDC_StopExecution()
    If fdc.timerController >= 0& Then
        timing_timerDisable fdc.timerController
    End If

    FDC_ClearDrq
    FDC_ResetDmaState
    fdc.execState = FDC_EXEC_NONE
    fdc.waitingSeek = 0&
End Sub

Private Sub FDC_SetIdle()
    fdc.command = 0&
    fdc.commandBase = 0&
    fdc.paramCount = 0&
    fdc.paramExpected = 0&
    fdc.resultCount = 0&
    fdc.resultIndex = 0&
    fdc.commandDrive = fdc.selectedDrive
    fdc.compareMode = FDC_SCAN_NONE
    fdc.compareSatisfied = 0&
    fdc.compareTotal = 0&
    fdc.verifyRemaining = 0&
    FDC_ResetDmaState
    fdc.msr = &H80&
End Sub

Private Sub FDC_SetTimerUsec(ByVal usecDelay As Long)
    If fdc.timerController < 0& Then Exit Sub
    timing_updateInterval fdc.timerController, timing_getFreq() * (usecDelay / 1000000)
    timing_timerEnable fdc.timerController
End Sub

Private Sub FDC_EnterResult(ByVal count As Long, ByVal raiseIrq As Byte)
    fdc.resultCount = count
    fdc.resultIndex = 0&
    fdc.paramCount = 0&
    fdc.paramExpected = 0&
    FDC_ResetDmaState
    fdc.execState = FDC_EXEC_NONE
    fdc.waitingSeek = 0&
    fdc.msr = &HD0&
    FDC_ClearDrq

    If raiseIrq <> 0& Then
        FDC_RaiseIRQ
    End If
End Sub

Private Sub FDC_ControllerReset(ByVal raiseInterrupt As Byte)
    Dim drive As Long

    FDC_StopExecution
    FDC_LowerIRQ

    fdc.command = 0&
    fdc.commandBase = 0&
    fdc.paramCount = 0&
    fdc.paramExpected = 0&
    fdc.resultCount = 0&
    fdc.resultIndex = 0&
    fdc.rate = 2&
    fdc.dsr = &H80&
    fdc.dmaEnabled = 1&
    fdc.powerDownMode = 0&
    fdc.compareMode = FDC_SCAN_NONE
    fdc.compareSatisfied = 0&
    fdc.compareTotal = 0&
    fdc.verifyRemaining = 0&
    fdc.fintr = 0&
    fdc.lastSenseSt0 = 0&
    fdc.lastSenseDrive = 0&
    fdc.selectedDrive = fdc.dor And 3&
    fdc.commandDrive = fdc.selectedDrive
    FDC_ResetDmaState
    fdc.msr = &H80&

    For drive = 0& To 3&
        fdc.pcn(drive) = 0&
    Next drive

    If raiseInterrupt <> 0& Then
        fdc.resetPending = 4&
        fdc.fintr = 1&
        FDC_RaiseIRQ
    Else
        fdc.resetPending = 0&
    End If
End Sub

Private Sub FDC_RequestReset()
    FDC_StopExecution
    fdc.execState = FDC_EXEC_RESET
    fdc.msr = 0&
    FDC_SetTimerUsec FDC_RESET_DELAY_USEC
End Sub

Private Function FDC_ParamBytes(ByVal commandByte As Byte) As Long
    Select Case commandByte
        Case FDC_CMD_LOCK
            FDC_ParamBytes = 0&
            Exit Function
    End Select

    Select Case (commandByte And &H1F&)
        Case FDC_CMD_SPECIFY
            FDC_ParamBytes = 2&
        Case FDC_CMD_MODE
            FDC_ParamBytes = 4&
        Case FDC_CMD_SENSE_DRIVE_STATUS
            FDC_ParamBytes = 1&
        Case FDC_CMD_READ_TRACK, FDC_CMD_WRITE_DATA, FDC_CMD_READ_DATA, FDC_CMD_WRITE_DELETED, FDC_CMD_READ_DELETED, FDC_CMD_SCAN_EQUAL, FDC_CMD_VERIFY, FDC_CMD_SCAN_LOW_EQUAL, FDC_CMD_SCAN_HIGH_EQUAL
            FDC_ParamBytes = 8&
        Case FDC_CMD_RECALIBRATE
            FDC_ParamBytes = 1&
        Case FDC_CMD_SENSE_INTERRUPT
            FDC_ParamBytes = 0&
        Case FDC_CMD_READ_ID
            FDC_ParamBytes = 1&
        Case FDC_CMD_FORMAT_TRACK
            FDC_ParamBytes = 5&
        Case FDC_CMD_DUMPREG
            FDC_ParamBytes = 0&
        Case FDC_CMD_SEEK
            FDC_ParamBytes = 2&
        Case FDC_CMD_VERSION
            FDC_ParamBytes = 0&
        Case FDC_CMD_PERPENDICULAR
            FDC_ParamBytes = 1&
        Case FDC_CMD_CONFIGURE
            FDC_ParamBytes = 3&
        Case FDC_CMD_POWERDOWN
            FDC_ParamBytes = 1&
        Case FDC_CMD_NSC
            FDC_ParamBytes = 0&
        Case FDC_CMD_UNLOCK
            FDC_ParamBytes = 0&
        Case Else
            FDC_ParamBytes = -1&
    End Select
End Function

Private Sub FDC_InvalidCommand()
    fdc.results(0&) = FDC_ST0_INVALID
    FDC_EnterResult 1&, 0&
End Sub

Private Sub FDC_SetTransferResult(ByVal drive As Long, ByVal headBit As Byte, ByVal st0Extra As Byte, ByVal st1 As Byte, ByVal st2 As Byte, ByVal track As Byte, ByVal head As Byte, ByVal sector As Byte, ByVal sizeCode As Byte, ByVal abnormal As Byte)
    Dim st0 As Long

    st0 = (drive And 3&) Or headBit Or st0Extra
    If abnormal <> 0& Then st0 = st0 Or FDC_ST0_ABNORMAL

    fdc.results(0&) = CByte(st0 And &HFF&)
    fdc.results(1&) = 0&
    fdc.results(2&) = 0&
    fdc.results(3&) = track
    fdc.results(4&) = head
    fdc.results(5&) = sector
    fdc.results(6&) = sizeCode
    fdc.results(1&) = st1
    fdc.results(2&) = st2
    FDC_EnterResult 7&, 1&
End Sub

Private Sub FDC_SetReadWriteSuccess(ByVal drive As Long, ByVal track As Byte, ByVal head As Byte, ByVal sector As Byte, ByVal sizeCode As Byte)
    FDC_SetTransferResult drive, FDC_HeadBit(fdd_getHead(FDC_PhysicalDrive())), 0&, 0&, 0&, track, head, sector, sizeCode, 0&
End Sub

Private Sub FDC_SetReadWriteError(ByVal drive As Long, ByVal head As Byte, ByVal st0Extra As Byte, ByVal st1 As Byte, ByVal st2 As Byte, ByVal track As Byte, ByVal sector As Byte, ByVal sizeCode As Byte)
    FDC_SetTransferResult drive, FDC_HeadBit(fdd_getHead(FDC_PhysicalDrive())), st0Extra, st1, st2, track, head, sector, sizeCode, 1&
End Sub

Private Function FDC_CommandMsr(ByVal execState As Byte) As Byte
    Select Case execState
        Case FDC_EXEC_READ, FDC_EXEC_READTRACK, FDC_EXEC_VERIFY
            FDC_CommandMsr = &H50&
        Case Else
            FDC_CommandMsr = &H10&
    End Select
End Function

Private Function FDC_IsScanByteSatisfied(ByVal compareMode As Byte, ByVal memoryValue As Byte, ByVal diskValue As Byte) As Boolean
    If memoryValue = &HFF& Then
        FDC_IsScanByteSatisfied = True
        Exit Function
    End If

    Select Case compareMode
        Case FDC_SCAN_EQUAL_MODE
            FDC_IsScanByteSatisfied = (memoryValue = diskValue)
        Case FDC_SCAN_LOW_EQUAL_MODE
            FDC_IsScanByteSatisfied = (memoryValue <= diskValue)
        Case FDC_SCAN_HIGH_EQUAL_MODE
            FDC_IsScanByteSatisfied = (memoryValue >= diskValue)
    End Select
End Function

Private Sub FDC_SetSenseInterruptResult()
    Dim drive As Long

    If fdc.resetPending > 0& Then
        drive = 4& - fdc.resetPending
        fdc.results(0&) = CByte(&HC0& Or (drive And 3&))
        fdc.results(1&) = 0&
        fdc.resetPending = fdc.resetPending - 1&
        If fdc.resetPending = 0& Then fdc.fintr = 0&
        FDC_EnterResult 2&, 0&
        Exit Sub
    End If

    If fdc.fintr <> 0& Then
        fdc.results(0&) = fdc.lastSenseSt0
        fdc.results(1&) = fdc.pcn(fdc.lastSenseDrive And 3&)
        fdc.fintr = 0&
        FDC_EnterResult 2&, 0&
        Exit Sub
    End If

    fdc.results(0&) = &H80&
    FDC_EnterResult 1&, 0&
End Sub

Private Sub FDC_SetSenseDriveResult()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim head As Long
    Dim st3 As Long

    drive = (fdc.params(0&) And 3&)
    physicalDrive = (fdc.selectedDrive And 3&)
    head = ((fdc.params(0&) And 4&) \ 4&)

    fdd_setHead physicalDrive, head

    st3 = (drive And 3&)
    If head <> 0& Then st3 = st3 Or FDC_ST0_HEAD
    If fdd_is_double_sided(physicalDrive) <> 0& Then st3 = st3 Or FDC_ST3_TWO_SIDE
    If fdd_track0(physicalDrive) <> 0& Then st3 = st3 Or FDC_ST3_TRACK0
    If fdd_get_flags(physicalDrive) <> 0& Then st3 = st3 Or FDC_ST3_READY
    If fdd_isWriteProtected(physicalDrive) <> 0& Then st3 = st3 Or FDC_ST3_WRITE_PROTECT

    fdc.results(0&) = CByte(st3 And &HFF&)
    FDC_EnterResult 1&, 0&
End Sub

Private Sub FDC_FinishSeekInterrupt(ByVal drive As Long, ByVal st0 As Byte)
    fdc.lastSenseSt0 = st0
    fdc.lastSenseDrive = CByte(drive And 3&)
    fdc.fintr = 1&
    fdc.execState = FDC_EXEC_NONE
    fdc.waitingSeek = 0&
    fdc.msr = &H80&
    FDC_RaiseIRQ
End Sub

Private Function FDC_SetupDataCommand(ByVal execState As Byte) As Boolean
    Dim physicalDrive As Long
    Dim seekDiff As Long

    physicalDrive = FDC_PhysicalDrive()

    If fdc.opSizeCode <> 2& Then
        FDC_SetReadWriteError fdc.rwDrive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Function
    End If

    If (fdd_get_flags(physicalDrive) = 0&) Or (FDC_IsMotorOn(physicalDrive) = False) Or (fdd_hasMedia(physicalDrive) = False) Then
        FDC_SetReadWriteError fdc.rwDrive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Function
    End If

    fdc.execState = execState

    If (fdc.config And &H40&) <> 0& Then
        If fdc.pcn(fdc.rwDrive And 3&) <> fdc.opTrack Then
            seekDiff = CLng(fdc.opTrack) - CLng(fdc.pcn(fdc.rwDrive And 3&))
            fdc.pcn(fdc.rwDrive And 3&) = fdc.opTrack
            fdc.waitingSeek = 1&
            fdc.msr = CByte(&H10& Or FDC_BusyMask(physicalDrive))
            fdd_seek physicalDrive, seekDiff
            FDC_SetupDataCommand = True
            Exit Function
        End If
    End If

    fdc.msr = FDC_CommandMsr(execState)
    FDC_SetTimerUsec FDC_EXEC_DELAY_USEC
    FDC_SetupDataCommand = True
End Function

Private Sub FDC_StartExplicitSeek(ByVal drive As Long, ByVal head As Long, ByVal targetTrack As Long, ByVal recalibrate As Byte)
    Dim physicalDrive As Long
    Dim st0 As Byte
    Dim seekDiff As Long

    physicalDrive = (fdc.selectedDrive And 3&)
    fdc.commandDrive = CByte(physicalDrive And 3&)
    fdd_setHead physicalDrive, head
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opHead = CByte(head And 1&)
    fdc.opTrack = CByte(targetTrack And &HFF&)

    st0 = CByte(&H20& Or (drive And 3&))

    If (fdd_get_flags(physicalDrive) = 0&) Or (FDC_IsMotorOn(physicalDrive) = False) Then
        If recalibrate <> 0& Then
            st0 = CByte(&H70& Or (drive And 3&))
            fdc.pcn(drive) = 0&
        Else
            st0 = CByte(&H20& Or (drive And 3&))
            fdc.pcn(drive) = CByte(targetTrack And &HFF&)
        End If
        FDC_FinishSeekInterrupt drive, st0
        Exit Sub
    End If

    If recalibrate <> 0& Then
        fdc.execState = FDC_EXEC_RECAL
        fdc.opTrack = 0&
        If fdd_track0(physicalDrive) <> 0& Then
            fdc.pcn(drive) = 0&
            FDC_FinishSeekInterrupt drive, CByte(&H20& Or (drive And 3&))
            Exit Sub
        End If
        seekDiff = -1024&
    Else
        fdc.execState = FDC_EXEC_SEEK
        If targetTrack = fdc.pcn(drive) Then
            fdc.pcn(drive) = CByte(targetTrack And &HFF&)
            FDC_FinishSeekInterrupt drive, CByte(&H20& Or (drive And 3&))
            Exit Sub
        End If
        seekDiff = CLng(targetTrack) - CLng(fdc.pcn(drive))
        fdc.pcn(drive) = CByte(targetTrack And &HFF&)
    End If

    fdc.waitingSeek = 1&
    fdc.msr = CByte(&H10& Or FDC_BusyMask(physicalDrive))
    fdd_seek physicalDrive, seekDiff
End Sub

Private Sub FDC_StartReadWrite(ByVal writeCommand As Byte)
    Dim drive As Long
    Dim selectedHead As Long
    Dim execState As Byte

    drive = (fdc.params(0&) And 3&)
    selectedHead = ((fdc.params(0&) And 4&) \ 4&)

    fdc.commandDrive = fdc.selectedDrive
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opTrack = fdc.params(1&)
    fdc.opHead = CByte(fdc.params(2&) And &HFF&)
    fdc.opSector = fdc.params(3&)
    fdc.opSizeCode = fdc.params(4&)
    fdc.opEot = fdc.params(5&)
    fdc.eot(drive) = fdc.opEot
    fdc.opGap = fdc.params(6&)
    fdc.opDtl = fdc.params(7&)
    fdc.compareMode = FDC_SCAN_NONE
    fdc.compareSatisfied = 0&
    fdc.compareTotal = 0&
    fdc.verifyRemaining = 0&

    fdd_setHead FDC_PhysicalDrive(), selectedHead

    If writeCommand <> 0& Then
        execState = FDC_EXEC_WRITE
    ElseIf (fdc.dmaEnabled <> 0&) And (i8237_get_operation(FDC_DMA_CH) = DMA_OP_VERIFY) Then
        execState = FDC_EXEC_VERIFY
        fdc.verifyRemaining = 1&
    Else
        execState = FDC_EXEC_READ
    End If

    Call FDC_SetupDataCommand(execState)
End Sub

Private Sub FDC_StartReadTrack()
    Dim drive As Long
    Dim selectedHead As Long

    drive = (fdc.params(0&) And 3&)
    selectedHead = ((fdc.params(0&) And 4&) \ 4&)

    fdc.commandDrive = fdc.selectedDrive
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opTrack = fdc.params(1&)
    fdc.opHead = CByte(fdc.params(2&) And &HFF&)
    fdc.opSector = 1&
    fdc.opSizeCode = fdc.params(4&)
    fdc.opEot = fdc.params(5&)
    fdc.eot(drive) = fdc.opEot
    fdc.opGap = fdc.params(6&)
    fdc.opDtl = fdc.params(7&)
    fdc.compareMode = FDC_SCAN_NONE
    fdc.compareSatisfied = 0&
    fdc.compareTotal = 0&
    fdc.verifyRemaining = 0&

    fdd_setHead FDC_PhysicalDrive(), selectedHead
    Call FDC_SetupDataCommand(FDC_EXEC_READTRACK)
End Sub

Private Sub FDC_StartVerify()
    Dim drive As Long
    Dim selectedHead As Long

    drive = (fdc.params(0&) And 3&)
    selectedHead = ((fdc.params(0&) And 4&) \ 4&)

    fdc.commandDrive = fdc.selectedDrive
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opTrack = fdc.params(1&)
    fdc.opHead = CByte(fdc.params(2&) And &HFF&)
    fdc.opSector = fdc.params(3&)
    fdc.opSizeCode = fdc.params(4&)
    fdc.opEot = fdc.params(5&)
    fdc.eot(drive) = fdc.opEot
    fdc.opGap = fdc.params(6&)
    fdc.opDtl = fdc.params(7&)
    fdc.compareMode = FDC_SCAN_NONE
    fdc.compareSatisfied = 0&
    fdc.compareTotal = 0&

    If (fdc.params(0&) And &H80&) <> 0& Then
        fdc.verifyRemaining = CLng(fdc.params(7&) And &HFF&)
    Else
        fdc.verifyRemaining = 1&
    End If
    If fdc.verifyRemaining <= 0& Then fdc.verifyRemaining = 1&

    fdd_setHead FDC_PhysicalDrive(), selectedHead
    Call FDC_SetupDataCommand(FDC_EXEC_VERIFY)
End Sub

Private Sub FDC_StartScan(ByVal compareMode As Byte)
    Dim drive As Long
    Dim selectedHead As Long

    drive = (fdc.params(0&) And 3&)
    selectedHead = ((fdc.params(0&) And 4&) \ 4&)

    fdc.commandDrive = fdc.selectedDrive
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opTrack = fdc.params(1&)
    fdc.opHead = CByte(fdc.params(2&) And &HFF&)
    fdc.opSector = fdc.params(3&)
    fdc.opSizeCode = fdc.params(4&)
    fdc.opEot = fdc.params(5&)
    fdc.eot(drive) = fdc.opEot
    fdc.opGap = fdc.params(6&)
    fdc.opDtl = fdc.params(7&)
    fdc.compareMode = compareMode
    fdc.compareSatisfied = 0&
    fdc.compareTotal = 0&
    fdc.verifyRemaining = 0&

    fdd_setHead FDC_PhysicalDrive(), selectedHead
    Call FDC_SetupDataCommand(FDC_EXEC_SCAN)
End Sub

Private Sub FDC_StartReadId()
    Dim drive As Long
    Dim head As Long

    drive = (fdc.params(0&) And 3&)
    head = ((fdc.params(0&) And 4&) \ 4&)

    fdc.commandDrive = fdc.selectedDrive
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opHead = CByte(head And 1&)
    fdd_setHead FDC_PhysicalDrive(), head

    If (fdd_get_flags(FDC_PhysicalDrive()) = 0&) Or (FDC_IsMotorOn(FDC_PhysicalDrive()) = False) Or (fdd_hasMedia(FDC_PhysicalDrive()) = False) Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, FDC_ST1_NO_ID, 0&, fdc.pcn(drive), 1&, 2&
        Exit Sub
    End If

    fdc.execState = FDC_EXEC_READID
    fdc.msr = &H10&
    FDC_SetTimerUsec FDC_EXEC_DELAY_USEC
End Sub

Private Sub FDC_StartFormat()
    Dim drive As Long
    Dim head As Long

    drive = (fdc.params(0&) And 3&)
    head = ((fdc.params(0&) And 4&) \ 4&)

    fdc.commandDrive = fdc.selectedDrive
    fdc.rwDrive = CByte(drive And 3&)
    fdc.opTrack = fdc.pcn(drive)
    fdc.opHead = CByte(head And 1&)
    fdc.opSizeCode = fdc.params(1&)
    fdc.formatSectorCount = fdc.params(2&)
    fdc.opGap = fdc.params(3&)
    fdc.formatFill = fdc.params(4&)

    fdd_setHead FDC_PhysicalDrive(), head

    If fdc.opSizeCode <> 2& Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    If (fdd_get_flags(FDC_PhysicalDrive()) = 0&) Or (FDC_IsMotorOn(FDC_PhysicalDrive()) = False) Or (fdd_hasMedia(FDC_PhysicalDrive()) = False) Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    fdc.execState = FDC_EXEC_FORMAT
    fdc.msr = &H10&
    FDC_SetTimerUsec FDC_EXEC_DELAY_USEC
End Sub

Private Function FDC_PrepareNextSector(ByVal drive As Long, ByVal terminalCount As Byte) As Long
    Dim currentHead As Long

    currentHead = fdd_getHead(drive)

    If terminalCount <> 0& Then
        If fdc.opSector = fdc.opEot Then
            If (fdc.command And &H80&) = 0& Then
                fdc.opTrack = CByte((CLng(fdc.opTrack) + 1&) And &HFF&)
                fdc.opSector = 1&
            Else
                If currentHead <> 0& Then
                    fdc.opTrack = CByte((CLng(fdc.opTrack) + 1&) And &HFF&)
                End If
                fdc.opHead = CByte((currentHead Xor 1&) And 1&)
                fdd_setHead drive, fdc.opHead
                fdc.opSector = 1&
            End If
        Else
            fdc.opSector = CByte((CLng(fdc.opSector) + 1&) And &HFF&)
        End If

        FDC_PrepareNextSector = 1&
        Exit Function
    End If

    If fdc.opSector = fdc.opEot Then
        If (fdc.command And &H80&) = 0& Then
            fdc.opTrack = CByte((CLng(fdc.opTrack) + 1&) And &HFF&)
            fdc.opSector = 1&
            FDC_PrepareNextSector = 1&
            Exit Function
        End If

        If currentHead <> 0& Then
            fdc.opTrack = CByte((CLng(fdc.opTrack) + 1&) And &HFF&)
            fdc.opSector = 1&
            fdc.opHead = 0&
            fdd_setHead drive, 0&
            FDC_PrepareNextSector = 1&
            Exit Function
        End If

        fdc.opSector = 1&
        fdc.opHead = 1&
        fdd_setHead drive, 1&

        If fdd_is_double_sided(drive) = 0& Then
            FDC_PrepareNextSector = 2&
            Exit Function
        End If

        Exit Function
    End If

    If (fdc.opSector < fdc.opEot) Or (fdc.opEot = 0&) Then
        fdc.opSector = CByte((CLng(fdc.opSector) + 1&) And &HFF&)
        Exit Function
    End If

    FDC_PrepareNextSector = 1&
End Function

Private Sub FDC_HandleDriveIoError(ByVal drive As Long, ByVal head As Byte, ByVal track As Byte, ByVal sector As Byte, ByVal sizeCode As Byte, ByVal status As Long)
    Select Case status
        Case FDD_IO_NO_MEDIA
            FDC_SetReadWriteError drive, head, FDC_ST0_NOT_READY, 0&, 0&, track, sector, sizeCode
        Case FDD_IO_NOT_FOUND
            FDC_SetReadWriteError drive, head, 0&, FDC_ST1_NO_ID, 0&, track, sector, sizeCode
        Case FDD_IO_WRITE_PROTECT
            FDC_SetReadWriteError drive, head, 0&, FDC_ST1_WRITE_PROTECT, 0&, track, sector, sizeCode
        Case FDD_IO_INVALID_SIZE
            FDC_SetReadWriteError drive, head, 0&, FDC_ST1_NO_DATA, 0&, track, sector, sizeCode
        Case Else
            FDC_SetReadWriteError drive, head, 0&, FDC_ST1_NO_DATA, 0&, track, sector, sizeCode
    End Select
End Sub

Public Function fdc_dmaTransferCallback(ByVal callbackData As Long, ByVal channel As Long, ByVal dmaPos As Long, ByVal dmaLen As Long) As Long
    Dim actualLen As Long
    Dim transferred As Long

    actualLen = fdc.dmaLength
    If actualLen > (dmaLen - dmaPos) Then actualLen = (dmaLen - dmaPos)
    If actualLen <= 0& Then
        fdc.dmaLastTransferred = 0&
        fdc_dmaTransferCallback = dmaPos
        Exit Function
    End If

    Select Case fdc.execState
        Case FDC_EXEC_READ, FDC_EXEC_READTRACK
            transferred = i8237_dma_writeMemory(channel, fdc_dmaSectorBuffer, dmaPos, actualLen)
        Case FDC_EXEC_WRITE, FDC_EXEC_SCAN, FDC_EXEC_FORMAT
            If fdc.dmaBufferKind = FDC_DMA_BUFFER_FORMAT Then
                transferred = i8237_dma_readMemory(channel, fdc_dmaFormatBuffer, dmaPos, actualLen)
            Else
                transferred = i8237_dma_readMemory(channel, fdc_dmaSectorBuffer, dmaPos, actualLen)
            End If
        Case Else
            transferred = 0&
    End Select

    fdc.dmaLastTransferred = transferred
    fdc_dmaTransferCallback = (dmaPos + transferred)
End Function

Private Sub FDC_ExecuteRead()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim status As Long
    Dim nextState As Long
    Dim tc As Byte

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    If fdc.dmaEnabled = 0& Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    If i8237_get_operation(FDC_DMA_CH) <> DMA_OP_WRITEMEM Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    Do
        status = fdd_readSector(physicalDrive, fdc.opSector, fdc.opTrack, fdc.opHead, fdc.opSizeCode, fdc_dmaSectorBuffer)
        If status <> FDD_IO_OK Then
            FDC_HandleDriveIoError drive, fdc.opHead, fdc.opTrack, fdc.opSector, fdc.opSizeCode, status
            Exit Sub
        End If

        fdc.dmaBufferKind = FDC_DMA_BUFFER_SECTOR
        fdc.dmaLength = 512&
        tc = FDC_RunDmaRequest()
        If fdc.dmaLastTransferred <= 0& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If

        nextState = FDC_PrepareNextSector(physicalDrive, tc)
        If nextState = 2& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_ID, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If
        If (tc <> 0&) Or (nextState <> 0&) Then Exit Do
    Loop

    FDC_SetReadWriteSuccess drive, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode
End Sub

Private Sub FDC_ExecuteWrite()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim status As Long
    Dim nextState As Long
    Dim i As Long
    Dim tc As Byte

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    If fdc.dmaEnabled = 0& Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    If i8237_get_operation(FDC_DMA_CH) <> DMA_OP_READMEM Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    Do
        For i = 0& To 511&
            fdc_dmaSectorBuffer(i) = 0&
        Next i

        fdc.dmaBufferKind = FDC_DMA_BUFFER_SECTOR
        fdc.dmaLength = 512&
        tc = FDC_RunDmaRequest()
        If fdc.dmaLastTransferred <= 0& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If

        status = fdd_writeSector(physicalDrive, fdc.opSector, fdc.opTrack, fdc.opHead, fdc.opSizeCode, fdc_dmaSectorBuffer)
        If status <> FDD_IO_OK Then
            FDC_HandleDriveIoError drive, fdc.opHead, fdc.opTrack, fdc.opSector, fdc.opSizeCode, status
            Exit Sub
        End If

        nextState = FDC_PrepareNextSector(physicalDrive, tc)
        If nextState = 2& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_ID, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If
        If (tc <> 0&) Or (nextState <> 0&) Then Exit Do
    Loop

    FDC_SetReadWriteSuccess drive, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode
End Sub

Private Sub FDC_ExecuteReadTrack()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim status As Long
    Dim tc As Byte

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    If fdc.dmaEnabled = 0& Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    If i8237_get_operation(FDC_DMA_CH) <> DMA_OP_WRITEMEM Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    Do
        status = fdd_readSector(physicalDrive, fdc.opSector, fdc.opTrack, fdc.opHead, fdc.opSizeCode, fdc_dmaSectorBuffer)
        If status <> FDD_IO_OK Then
            FDC_HandleDriveIoError drive, fdc.opHead, fdc.opTrack, fdc.opSector, fdc.opSizeCode, status
            Exit Sub
        End If

        fdc.dmaBufferKind = FDC_DMA_BUFFER_SECTOR
        fdc.dmaLength = 512&
        tc = FDC_RunDmaRequest()
        If fdc.dmaLastTransferred <= 0& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If

        fdc.opSector = CByte((CLng(fdc.opSector) + 1&) And &HFF&)
        If (tc <> 0&) Or (fdc.opSector > fdc.opEot) Then Exit Do
    Loop

    FDC_SetReadWriteSuccess drive, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode
End Sub

Private Sub FDC_ExecuteVerify()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim status As Long
    Dim nextState As Long
    Dim buffer(0& To 511&) As Byte

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    Do
        status = fdd_readSector(physicalDrive, fdc.opSector, fdc.opTrack, fdc.opHead, fdc.opSizeCode, buffer)
        If status <> FDD_IO_OK Then
            FDC_HandleDriveIoError drive, fdc.opHead, fdc.opTrack, fdc.opSector, fdc.opSizeCode, status
            Exit Sub
        End If

        If fdc.verifyRemaining > 0& Then
            fdc.verifyRemaining = fdc.verifyRemaining - 1&
        End If

        nextState = FDC_PrepareNextSector(physicalDrive, 0&)
        If nextState = 2& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_ID, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If
        If (fdc.verifyRemaining <= 0&) Or (nextState <> 0&) Then Exit Do
    Loop

    FDC_SetReadWriteSuccess drive, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode
End Sub

Private Sub FDC_ExecuteScan()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim status As Long
    Dim nextState As Long
    Dim i As Long
    Dim tc As Byte
    Dim bytesCompared As Long
    Dim sectorSatisfied As Boolean
    Dim diskBuffer(0& To 511&) As Byte

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    If fdc.dmaEnabled = 0& Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    If i8237_get_operation(FDC_DMA_CH) <> DMA_OP_READMEM Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
        Exit Sub
    End If

    Do
        status = fdd_readSector(physicalDrive, fdc.opSector, fdc.opTrack, fdc.opHead, fdc.opSizeCode, diskBuffer)
        If status <> FDD_IO_OK Then
            FDC_HandleDriveIoError drive, fdc.opHead, fdc.opTrack, fdc.opSector, fdc.opSizeCode, status
            Exit Sub
        End If

        sectorSatisfied = True
        bytesCompared = 0&

        For i = 0& To 511&
            fdc_dmaSectorBuffer(i) = 0&
        Next i

        fdc.dmaBufferKind = FDC_DMA_BUFFER_SECTOR
        fdc.dmaLength = 512&
        tc = FDC_RunDmaRequest()
        bytesCompared = fdc.dmaLastTransferred
        If bytesCompared <= 0& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If

        For i = 0& To bytesCompared - 1&
            If Not FDC_IsScanByteSatisfied(fdc.compareMode, fdc_dmaSectorBuffer(i), diskBuffer(i)) Then
                sectorSatisfied = False
            End If
        Next i

        If bytesCompared = 512& Then
            fdc.compareTotal = fdc.compareTotal + 1&
            If sectorSatisfied Then
                fdc.compareSatisfied = fdc.compareSatisfied + 1&
            End If
        End If

        nextState = FDC_PrepareNextSector(physicalDrive, tc)
        If nextState = 2& Then
            FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_ID, 0&, fdc.opTrack, fdc.opSector, fdc.opSizeCode
            Exit Sub
        End If
        If (tc <> 0&) Or (nextState <> 0&) Then Exit Do
    Loop

    If fdc.compareMode <> FDC_SCAN_NONE Then
        If fdc.compareSatisfied = 0& Then
            FDC_SetTransferResult drive, FDC_HeadBit(fdd_getHead(physicalDrive)), 0&, 0&, FDC_ST2_SCAN_NOT_SATISFIED, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode, 0&
        ElseIf fdc.compareSatisfied = fdc.compareTotal Then
            FDC_SetTransferResult drive, FDC_HeadBit(fdd_getHead(physicalDrive)), 0&, 0&, FDC_ST2_SCAN_EQUAL_HIT, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode, 0&
        Else
            FDC_SetReadWriteSuccess drive, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode
        End If
    Else
        FDC_SetReadWriteSuccess drive, fdc.opTrack, fdc.opHead, fdc.opSector, fdc.opSizeCode
    End If
End Sub

Private Sub FDC_ExecuteReadId()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim trackVal As Byte
    Dim headVal As Byte
    Dim sectorVal As Byte
    Dim sizeVal As Byte
    Dim status As Long

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    status = fdd_readAddress(physicalDrive, fdc.opHead, trackVal, headVal, sectorVal, sizeVal)
    If status <> FDD_IO_OK Then
        FDC_HandleDriveIoError drive, fdc.opHead, fdc.pcn(drive), 1&, 2&, status
        Exit Sub
    End If

    FDC_SetReadWriteSuccess drive, trackVal, headVal, sectorVal, sizeVal
End Sub

Private Sub FDC_ExecuteFormat()
    Dim drive As Long
    Dim physicalDrive As Long
    Dim totalBytes As Long
    Dim i As Long
    Dim status As Long
    Dim lastTrack As Byte
    Dim lastHead As Byte
    Dim lastSector As Byte
    Dim lastSize As Byte
    Dim tc As Byte

    drive = (fdc.rwDrive And 3&)
    physicalDrive = FDC_PhysicalDrive()

    If fdc.dmaEnabled = 0& Then
        FDC_SetReadWriteError drive, fdc.opHead, FDC_ST0_NOT_READY, 0&, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    If i8237_get_operation(FDC_DMA_CH) <> DMA_OP_READMEM Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    If fdc.formatSectorCount = 0& Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    totalBytes = CLng(fdc.formatSectorCount) * 4&
    If totalBytes > 256& Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    For i = 0& To totalBytes - 1&
        fdc_dmaFormatBuffer(i) = 0&
    Next i

    fdc.dmaBufferKind = FDC_DMA_BUFFER_FORMAT
    fdc.dmaLength = totalBytes
    tc = FDC_RunDmaRequest()
    If fdc.dmaLastTransferred < totalBytes Then
        FDC_SetReadWriteError drive, fdc.opHead, 0&, FDC_ST1_NO_DATA, 0&, fdc.opTrack, 1&, fdc.opSizeCode
        Exit Sub
    End If

    status = fdd_formatTrack(physicalDrive, fdc.opTrack, fdc.opHead, fdc.formatFill, fdc.formatSectorCount, fdc_dmaFormatBuffer)
    If status <> FDD_IO_OK Then
        FDC_HandleDriveIoError drive, fdc.opHead, fdc.opTrack, 1&, fdc.opSizeCode, status
        Exit Sub
    End If

    lastTrack = fdc_dmaFormatBuffer(totalBytes - 4&)
    lastHead = fdc_dmaFormatBuffer(totalBytes - 3&)
    lastSector = fdc_dmaFormatBuffer(totalBytes - 2&)
    lastSize = fdc_dmaFormatBuffer(totalBytes - 1&)

    FDC_SetReadWriteSuccess drive, lastTrack, lastHead, lastSector, lastSize
End Sub

Private Sub FDC_ProcessImmediateCommand(ByVal commandBase As Byte)
    Dim drive As Long

    Select Case fdc.command
        Case FDC_CMD_LOCK
            fdc.controllerLock = 1&
            fdc.results(0&) = &H10&
            FDC_EnterResult 1&, 0&
            Exit Sub
    End Select

    Select Case (commandBase And &H1F&)
        Case FDC_CMD_SENSE_INTERRUPT
            FDC_SetSenseInterruptResult
        Case FDC_CMD_DUMPREG
            drive = (fdc.selectedDrive And 3&)
            fdc.results(0&) = fdc.pcn(0&)
            fdc.results(1&) = fdc.pcn(1&)
            fdc.results(2&) = fdc.pcn(2&)
            fdc.results(3&) = fdc.pcn(3&)
            fdc.results(4&) = fdc.specify(0&)
            fdc.results(5&) = fdc.specify(1&)
            fdc.results(6&) = fdc.eot(drive)
            fdc.results(7&) = CByte((fdc.perp And &H7F&) Or IIf(fdc.controllerLock <> 0&, &H80&, 0&))
            fdc.results(8&) = fdc.config
            fdc.results(9&) = fdc.pretrk
            FDC_EnterResult 10&, 0&
        Case FDC_CMD_VERSION
            fdc.results(0&) = &H90&
            FDC_EnterResult 1&, 0&
        Case FDC_CMD_NSC
            fdc.results(0&) = &H73&
            FDC_EnterResult 1&, 0&
        Case FDC_CMD_UNLOCK
            fdc.controllerLock = 0&
            fdc.results(0&) = 0&
            FDC_EnterResult 1&, 0&
        Case Else
            FDC_InvalidCommand
    End Select
End Sub

Private Sub FDC_ProcessCommand()
    Select Case fdc.commandBase
        Case FDC_CMD_MODE
            fdc.denselForce = CByte((fdc.params(2&) And &HC0&) \ &H40&)
            FDC_SetIdle

        Case FDC_CMD_SPECIFY
            fdc.specify(0&) = fdc.params(0&)
            fdc.specify(1&) = fdc.params(1&)
            fdc.dmaEnabled = CByte((fdc.params(1&) And 1&) Xor 1&)
            FDC_SetIdle

        Case FDC_CMD_SENSE_DRIVE_STATUS
            FDC_SetSenseDriveResult

        Case FDC_CMD_CONFIGURE
            fdc.config = fdc.params(1&)
            fdc.pretrk = fdc.params(2&)
            FDC_SetIdle

        Case FDC_CMD_PERPENDICULAR
            If (fdc.params(0&) And &H80&) <> 0& Then
                fdc.perp = CByte(fdc.params(0&) And &H3F&)
            Else
                fdc.perp = CByte((fdc.perp And &HFC&) Or (fdc.params(0&) And 3&))
            End If
            FDC_SetIdle

        Case FDC_CMD_POWERDOWN
            fdc.powerDownMode = fdc.params(0&)
            fdc.results(0&) = fdc.powerDownMode
            FDC_EnterResult 1&, 0&

        Case FDC_CMD_RECALIBRATE
            Call FDC_StartExplicitSeek((fdc.params(0&) And 3&), 0&, 0&, 1&)

        Case FDC_CMD_SEEK
            Call FDC_StartExplicitSeek((fdc.params(0&) And 3&), ((fdc.params(0&) And 4&) \ 4&), fdc.params(1&), 0&)

        Case FDC_CMD_READ_TRACK
            FDC_StartReadTrack

        Case FDC_CMD_READ_DATA
            FDC_StartReadWrite 0&

        Case FDC_CMD_WRITE_DATA
            FDC_StartReadWrite 1&

        Case FDC_CMD_WRITE_DELETED
            FDC_StartReadWrite 1&

        Case FDC_CMD_READ_DELETED
            FDC_StartReadWrite 0&

        Case FDC_CMD_VERIFY
            FDC_StartVerify

        Case FDC_CMD_SCAN_EQUAL
            FDC_StartScan FDC_SCAN_EQUAL_MODE

        Case FDC_CMD_SCAN_LOW_EQUAL
            FDC_StartScan FDC_SCAN_LOW_EQUAL_MODE

        Case FDC_CMD_SCAN_HIGH_EQUAL
            FDC_StartScan FDC_SCAN_HIGH_EQUAL_MODE

        Case FDC_CMD_READ_ID
            FDC_StartReadId

        Case FDC_CMD_FORMAT_TRACK
            FDC_StartFormat

        Case Else
            FDC_InvalidCommand
    End Select
End Sub

Private Sub FDC_WriteDataRegister(ByVal value As Byte)
    If (fdc.dor And &H4&) = 0& Then Exit Sub
    If (fdc.msr And &H40&) <> 0& Then Exit Sub

    If fdc.paramExpected <> 0& Then
        fdc.params(fdc.paramCount) = value
        fdc.paramCount = fdc.paramCount + 1&
        fdc.msr = &H90&

        If fdc.paramCount >= fdc.paramExpected Then
            fdc.paramExpected = 0&
            fdc.paramCount = 0&
            FDC_ProcessCommand
        End If
        Exit Sub
    End If

    FDC_LowerIRQ
    fdc.command = value
    fdc.commandBase = (value And &H1F&)
    fdc.paramExpected = FDC_ParamBytes(value)
    fdc.paramCount = 0&
    fdc.resultCount = 0&
    fdc.resultIndex = 0&

    If fdc.paramExpected < 0& Then
        FDC_InvalidCommand
        Exit Sub
    End If

    If fdc.paramExpected = 0& Then
        FDC_ProcessImmediateCommand fdc.commandBase
    Else
        fdc.msr = &H90&
    End If
End Sub

Private Function FDC_ReadDataRegister() As Byte
    If fdc.resultIndex < fdc.resultCount Then
        FDC_ReadDataRegister = fdc.results(fdc.resultIndex)
        fdc.resultIndex = fdc.resultIndex + 1&
        If fdc.resultIndex >= fdc.resultCount Then
            FDC_SetIdle
        Else
            fdc.msr = &HD0&
        End If
        Exit Function
    End If

    FDC_ReadDataRegister = 0&
End Function

Public Sub fdc_controllerCallback(ByVal dummy As Long)
    If fdc.inited = 0& Then Exit Sub
    If fdc.timerController >= 0& Then timing_timerDisable fdc.timerController

    Select Case fdc.execState
        Case FDC_EXEC_RESET
            FDC_ControllerReset 1&
        Case FDC_EXEC_READ
            FDC_ExecuteRead
        Case FDC_EXEC_WRITE
            FDC_ExecuteWrite
        Case FDC_EXEC_READTRACK
            FDC_ExecuteReadTrack
        Case FDC_EXEC_VERIFY
            FDC_ExecuteVerify
        Case FDC_EXEC_SCAN
            FDC_ExecuteScan
        Case FDC_EXEC_READID
            FDC_ExecuteReadId
        Case FDC_EXEC_FORMAT
            FDC_ExecuteFormat
    End Select
End Sub

Public Sub fdc_seek_complete_interrupt(ByVal drive As Long)
    Dim st0 As Byte

    If fdc.inited = 0& Then Exit Sub
    If (drive < 0&) Or (drive > 3&) Then Exit Sub

    If (fdc.waitingSeek <> 0&) And (FDC_PhysicalDrive() = drive) Then
        fdc.pcn(fdc.rwDrive And 3&) = CByte(fdd_getCurrentTrack(drive) And &HFF&)
    Else
        fdc.pcn(drive) = CByte(fdd_getCurrentTrack(drive) And &HFF&)
    End If

    If (fdc.waitingSeek <> 0&) And (FDC_PhysicalDrive() = drive) Then
        If (fdc.execState = FDC_EXEC_SEEK) Or (fdc.execState = FDC_EXEC_RECAL) Then
            st0 = CByte(FDC_ST0_SEEK_END Or (drive And 3&) Or FDC_HeadBit(fdd_getHead(drive)))
            fdc.lastSenseSt0 = st0
            fdc.lastSenseDrive = CByte(drive And 3&)
            fdc.fintr = 1&
            fdc.waitingSeek = 0&
            fdc.execState = FDC_EXEC_NONE
            fdc.msr = &H80&
            FDC_RaiseIRQ
            Exit Sub
        End If

        fdc.waitingSeek = 0&
        If fdc.execState <> FDC_EXEC_NONE Then
            fdc.msr = FDC_CommandMsr(fdc.execState)
            FDC_SetTimerUsec FDC_EXEC_DELAY_USEC
            Exit Sub
        End If
    End If

    st0 = CByte(FDC_ST0_SEEK_END Or (drive And 3&) Or FDC_HeadBit(fdd_getHead(drive)))
    fdc.lastSenseSt0 = st0
    fdc.lastSenseDrive = CByte(drive And 3&)
    fdc.fintr = 1&
    fdc.msr = &H80&
    FDC_RaiseIRQ
End Sub

Public Sub fdc_write(ByVal dummy As Long, ByVal addr As Integer, ByVal value As Byte)
    Dim drive As Long
    Dim prevDor As Byte

    If fdc.inited = 0& Then Exit Sub

    Select Case (addr And 7&)
        Case FDC_PORT_DOR
            prevDor = fdc.dor
            fdc.dor = value
            fdc.selectedDrive = CByte(value And 3&)

            For drive = 0& To 3&
                If (value And FDC_MotorMask(drive)) <> 0& Then
                    fdd_setMotorEnable drive, 1&
                Else
                    fdd_setMotorEnable drive, 0&
                End If
            Next drive

            If (value And &H4&) = 0& Then
                FDC_StopExecution
                FDC_LowerIRQ
                fdc.msr = 0&
                Exit Sub
            End If

            If ((prevDor And &H4&) = 0&) And ((value And &H4&) <> 0&) Then
                FDC_RequestReset
            End If

            If (value And &H8&) = 0& Then
                FDC_LowerIRQ
            End If

        Case FDC_PORT_MSR
            fdc.dsr = value
            fdc.rate = (value And 3&)

        Case FDC_PORT_DATA
            FDC_WriteDataRegister value

        Case FDC_PORT_DIR
            fdc.rate = (value And 3&)
    End Select
End Sub

Public Function fdc_read(ByVal dummy As Long, ByVal addr As Integer) As Byte
    Dim drive As Long
    Dim ret As Long

    If fdc.inited = 0& Then
        fdc_read = &HFF&
        Exit Function
    End If

    Select Case (addr And 7&)
        Case 2&
            fdc_read = fdc.dor

        Case 3&
            fdc_read = &HFF&

        Case FDC_PORT_MSR
            fdc_read = fdc.msr

        Case FDC_PORT_DATA
            fdc_read = FDC_ReadDataRegister()

        Case FDC_PORT_DIR
            drive = (fdc.dor And 3&)
            ret = &H7F&
            If ((fdc.dor And FDC_MotorMask(drive)) <> 0&) And ((fdd_getChanged(drive) <> 0&) Or (fdd_hasMedia(drive) = False)) Then
                ret = ret Or &H80&
            End If
            fdc_read = CByte(ret And &HFF&)

        Case Else
            fdc_read = &HFF&
    End Select
End Function

Public Function fdc_init(ByRef cpu As CPU_t, ByVal i8259Slot As Long) As Long
    fdc.inited = 1&
    fdc.i8259Slot = i8259Slot
    fdc.dor = &HC&
    fdc.dsr = &H80&
    fdc.rate = 2&
    fdc.config = &H20&
    fdc.pretrk = 0&
    fdc.controllerLock = 0&
    fdc.perp = 0&
    fdc.irqPending = 0&
    fdc.dmaEnabled = 1&
    fdc.timerController = timing_addTimerUsingInterval(TIMER_CB_FDC_CONTROLLER, 0&, timing_getFreq() * (FDC_EXEC_DELAY_USEC / 1000000), TIMING_DISABLED)

    ports_cbRegister &H3F0&, 6&, PORTS_CB_FDC, PORTS_CB_NONE, PORTS_CB_FDC, PORTS_CB_NONE, 0&
    ports_cbRegister &H3F7&, 1&, PORTS_CB_FDC, PORTS_CB_NONE, PORTS_CB_FDC, PORTS_CB_NONE, 0&
    i8237_dma_registerChannel FDC_DMA_CH, DMA_CB_NONE, 0&

    FDC_ControllerReset 0&
    fdc_init = 0&
End Function
