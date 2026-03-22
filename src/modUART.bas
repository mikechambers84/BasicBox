Attribute VB_Name = "modUART"
Option Explicit

Public Const UART_IRQ_MSR_ENABLE As Byte = &H8&
Public Const UART_IRQ_LSR_ENABLE As Byte = &H4&
Public Const UART_IRQ_TX_ENABLE As Byte = &H2&
Public Const UART_IRQ_RX_ENABLE As Byte = &H1&

Public Const UART_PENDING_RX As Byte = &H1&
Public Const UART_PENDING_TX As Byte = &H2&
Public Const UART_PENDING_MSR As Byte = &H4&
Public Const UART_PENDING_LSR As Byte = &H8&

Public Const UART_MSR_CTS As Byte = &H10&
Public Const UART_MSR_DSR As Byte = &H20&
Public Const UART_MSR_RI As Byte = &H40&
Public Const UART_MSR_DCD As Byte = &H80&

Private Const UART_LSR_DR As Byte = &H1&
Private Const UART_LSR_OE As Byte = &H2&
Private Const UART_LSR_THRE As Byte = &H20&
Private Const UART_LSR_TEMT As Byte = &H40&

Private Const UART_MAX As Long = 4&
Private Const UART_RX_BUFFER_LEN As Long = 64&
Private Const UART_MCR_OUT2 As Byte = &H8&

Private Const UART_TXCB_NONE As Byte = 0&
Private Const UART_TXCB_TCPMODEM As Byte = 1&

Private Const UART_MCRCB_NONE As Byte = 0&
Private Const UART_MCRCB_MOUSE_TOGGLE As Byte = 1&

Private Type UART_t
    rx As Byte
    rxStage As Byte
    tx As Byte
    rxnew As Byte
    rxStageValid As Byte
    dlab As Byte
    ien As Byte
    iir As Byte
    lcr As Byte
    mcr As Byte
    lsr As Byte
    msr As Byte
    lastmsr As Byte
    scratch As Byte
    divisor As Long
    irq As Byte
    pendirq As Byte
    i8259Slot As Long
    txCbKind As Byte
    mcrCbKind As Byte
    udata As Long
    udata2 As Long
    rxTimerId As Long
    rxTimerActive As Byte
    rxHead As Long
    rxTail As Long
    rxCount As Long
    irqLineActive As Byte
End Type

Private uart_devs(0& To UART_MAX - 1&) As UART_t
Private uart_used(0& To UART_MAX - 1&) As Byte
Private uart_wordmask(0& To 3&) As Byte
Private uart_masksInit As Byte
Private uart_rxbuf(0& To UART_MAX - 1&, 0& To UART_RX_BUFFER_LEN - 1&) As Byte

Private Sub UART_InitMasks()
    If uart_masksInit <> 0& Then Exit Sub

    uart_wordmask(0&) = &H1F&
    uart_wordmask(1&) = &H3F&
    uart_wordmask(2&) = &H7F&
    uart_wordmask(3&) = &HFF&

    uart_masksInit = 1&
End Sub

Private Function UART_IsValid(ByVal uartnum As Long) As Boolean
    UART_IsValid = ((uartnum >= 0&) And (uartnum < UART_MAX) And (uart_used(uartnum) <> 0&))
End Function

Private Function UART_RxBufferedCount(ByVal uartnum As Long) As Long
    If Not UART_IsValid(uartnum) Then Exit Function

    UART_RxBufferedCount = uart_devs(uartnum).rxCount
    If uart_devs(uartnum).rxnew <> 0& Then UART_RxBufferedCount = UART_RxBufferedCount + 1&
End Function

Private Function UART_IsMouse(ByVal uartnum As Long) As Boolean
    If Not UART_IsValid(uartnum) Then Exit Function
    UART_IsMouse = (uart_devs(uartnum).mcrCbKind = UART_MCRCB_MOUSE_TOGGLE)
End Function

Private Function UART_GetReceiveBitCount(ByVal uartnum As Long) As Double
    Dim bitCount As Double

    If Not UART_IsValid(uartnum) Then Exit Function

    bitCount = CDbl(5& + (uart_devs(uartnum).lcr And &H3&))
    bitCount = bitCount + 1# ' Start bit.
    bitCount = bitCount + 1# ' First stop bit.
    If (uart_devs(uartnum).lcr And &H4&) <> 0& Then bitCount = bitCount + 1#
    If (uart_devs(uartnum).lcr And &H8&) <> 0& Then bitCount = bitCount + 1#

    UART_GetReceiveBitCount = bitCount
End Function

Private Function UART_GetReceiveBaud(ByVal uartnum As Long) As Double
    Dim divisor As Long

    If Not UART_IsValid(uartnum) Then Exit Function

    divisor = (uart_devs(uartnum).divisor And &HFFFF&)
    If divisor <> 0& Then
        UART_GetReceiveBaud = 115200# / CDbl(divisor)
    ElseIf UART_IsMouse(uartnum) Then
        UART_GetReceiveBaud = MOUSE_DEFAULT_BAUD
    Else
        UART_GetReceiveBaud = 115200#
    End If
End Function

Private Sub UART_UpdateReceiveTiming(ByVal uartnum As Long)
    Dim baud As Double
    Dim bitCount As Double

    If Not UART_IsValid(uartnum) Then Exit Sub
    If uart_devs(uartnum).rxTimerId = TIMING_ERROR Then Exit Sub

    baud = UART_GetReceiveBaud(uartnum)
    bitCount = UART_GetReceiveBitCount(uartnum)

    If baud < 1# Then baud = 1#
    If bitCount < 1# Then bitCount = 1#

    timing_updateIntervalFreq uart_devs(uartnum).rxTimerId, baud / bitCount
End Sub

Private Sub UART_EnableReceiveTimer(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub
    If uart_devs(uartnum).rxTimerId = TIMING_ERROR Then Exit Sub

    UART_UpdateReceiveTiming uartnum
    If uart_devs(uartnum).rxTimerActive = 0& Then
        timing_timerEnable uart_devs(uartnum).rxTimerId
        uart_devs(uartnum).rxTimerActive = 1&
    End If
End Sub

Private Sub UART_DisableReceiveTimer(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub
    If uart_devs(uartnum).rxTimerId = TIMING_ERROR Then Exit Sub

    timing_timerDisable uart_devs(uartnum).rxTimerId
    uart_devs(uartnum).rxTimerActive = 0&
End Sub

Private Sub UART_QueueMouseReceiveByte(ByVal uartnum As Long, ByVal value As Byte)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        If .rxStageValid = 0& Then
            .rxStage = value
            .rxStageValid = 1&
        ElseIf .rxCount < UART_RX_BUFFER_LEN Then
            uart_rxbuf(uartnum, .rxTail) = value
            .rxTail = (.rxTail + 1&) Mod UART_RX_BUFFER_LEN
            .rxCount = .rxCount + 1&
        End If
    End With

    UART_EnableReceiveTimer uartnum
End Sub

Private Sub UART_AdvanceMouseStage(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        If .rxCount > 0& Then
            .rxStage = uart_rxbuf(uartnum, .rxHead)
            .rxHead = (.rxHead + 1&) Mod UART_RX_BUFFER_LEN
            .rxCount = .rxCount - 1&
            .rxStageValid = 1&
            UART_EnableReceiveTimer uartnum
        Else
            .rxStageValid = 0&
            UART_DisableReceiveTimer uartnum
        End If
    End With
End Sub

Private Sub UART_UpdateInterruptLine(ByVal uartnum As Long)
    Dim wantsIrq As Byte

    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        wantsIrq = CByte((.pendirq <> 0&) And ((.mcr And UART_MCR_OUT2) <> 0&))

        If wantsIrq <> 0& Then
            If .irqLineActive = 0& Then
                .irqLineActive = 1&
                i8259_setlevelirq .i8259Slot, .irq, 1&
            End If
        ElseIf .irqLineActive <> 0& Then
            .irqLineActive = 0&
            i8259_setlevelirq .i8259Slot, .irq, 0&
        End If
    End With
End Sub

Private Sub UART_TriggerRxIrq(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        If (.ien And UART_IRQ_RX_ENABLE) <> 0& Then
            .pendirq = (.pendirq Or UART_PENDING_RX)
        End If
    End With

    UART_UpdateInterruptLine uartnum
End Sub

Private Sub UART_TriggerLsrIrq(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        If (.ien And UART_IRQ_LSR_ENABLE) <> 0& Then
            .pendirq = (.pendirq Or UART_PENDING_LSR)
        End If
    End With

    UART_UpdateInterruptLine uartnum
End Sub

Private Sub UART_SetCurrentRxByte(ByVal uartnum As Long, ByVal value As Byte)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        .rx = value
        .rxnew = 1&
        .lsr = CByte((.lsr Or UART_LSR_DR) And &HFF&)
    End With

    UART_TriggerRxIrq uartnum
End Sub

Private Sub UART_LoadNextRxByte(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        .pendirq = CByte(.pendirq And ((Not UART_PENDING_RX) And &HFF&))
        UART_UpdateInterruptLine uartnum
        If .rxCount > 0& Then
            .rx = uart_rxbuf(uartnum, .rxHead)
            .rxHead = (.rxHead + 1&) Mod UART_RX_BUFFER_LEN
            .rxCount = .rxCount - 1&
            .rxnew = 1&
            .lsr = CByte((.lsr Or UART_LSR_DR) And &HFF&)
        Else
            .rxnew = 0&
            .pendirq = CByte(.pendirq And ((Not UART_PENDING_RX) And &HFF&))
            .lsr = CByte(.lsr And ((Not UART_LSR_DR) And &HFF&))
        End If
    End With

    If uart_devs(uartnum).rxnew <> 0& Then
        UART_TriggerRxIrq uartnum
    End If
End Sub

Public Sub uart_writeport(ByVal uartnum As Long, ByVal addr As Integer, ByVal value As Byte)
    Dim oldIen As Byte

    If Not UART_IsValid(uartnum) Then Exit Sub

    addr = (addr And &H7&)

    With uart_devs(uartnum)
        Select Case addr
            Case &H0&
                If .dlab = 0& Then
                    .tx = (value And uart_wordmask(.lcr And &H3&))
                    debug_log DEBUG_DETAIL, Chr$(.tx)

                    If (.mcr And &H10&) <> 0& Then
                        uart_rxdata uartnum, .tx
                    Else
                        If .txCbKind <> UART_TXCB_NONE Then
                            Select Case .txCbKind
                                Case UART_TXCB_TCPMODEM
                                    tcpmodem_tx .udata, .tx
                            End Select

                            If (.ien And UART_IRQ_TX_ENABLE) <> 0& Then
                                .pendirq = (.pendirq Or UART_PENDING_TX)
                                UART_UpdateInterruptLine uartnum
                            End If
                        End If
                    End If
                Else
                    .divisor = ((.divisor And &HFF00&) Or (value And &HFF&))
                    UART_UpdateReceiveTiming uartnum
                    If .mcrCbKind = UART_MCRCB_MOUSE_TOGGLE Then
                        mouse_updateSerialDivisor uartnum, .divisor
                    End If
                End If

            Case &H1&
                If .dlab = 0& Then
                    oldIen = .ien
                    .ien = value
                    If (.ien And UART_IRQ_RX_ENABLE) = 0& Then
                        .pendirq = CByte(.pendirq And ((Not UART_PENDING_RX) And &HFF&))
                    End If
                    If (.ien And UART_IRQ_TX_ENABLE) = 0& Then
                        .pendirq = CByte(.pendirq And ((Not UART_PENDING_TX) And &HFF&))
                    End If
                    If (.ien And UART_IRQ_MSR_ENABLE) = 0& Then
                        .pendirq = CByte(.pendirq And ((Not UART_PENDING_MSR) And &HFF&))
                    End If
                    If (.ien And UART_IRQ_LSR_ENABLE) = 0& Then
                        .pendirq = CByte(.pendirq And ((Not UART_PENDING_LSR) And &HFF&))
                    End If
                    If ((oldIen And UART_IRQ_RX_ENABLE) = 0&) And ((.ien And UART_IRQ_RX_ENABLE) <> 0&) And (.rxnew <> 0&) Then
                        UART_TriggerRxIrq uartnum
                    End If
                    If ((oldIen And UART_IRQ_MSR_ENABLE) = 0&) And ((.ien And UART_IRQ_MSR_ENABLE) <> 0&) Then
                        If (.msr And &HF0&) <> (.lastmsr And &HF0&) Then
                            .pendirq = (.pendirq Or UART_PENDING_MSR)
                        End If
                    End If
                    If ((oldIen And UART_IRQ_LSR_ENABLE) = 0&) And ((.ien And UART_IRQ_LSR_ENABLE) <> 0&) Then
                        If (.lsr And &H1E&) <> 0& Then
                            .pendirq = (.pendirq Or UART_PENDING_LSR)
                        End If
                    End If
                    UART_UpdateInterruptLine uartnum
                Else
                    .divisor = ((.divisor And &HFF&) Or ((CLng(value) And &HFF&) * &H100&))
                    UART_UpdateReceiveTiming uartnum
                    If .mcrCbKind = UART_MCRCB_MOUSE_TOGGLE Then
                        mouse_updateSerialDivisor uartnum, .divisor
                    End If
                End If

            Case &H3&
                .lcr = value
                .dlab = CByte((value And &H80&) \ &H80&)
                UART_UpdateReceiveTiming uartnum

            Case &H4&
                .mcr = value
                If .mcrCbKind <> UART_MCRCB_NONE Then
                    Select Case .mcrCbKind
                        Case UART_MCRCB_MOUSE_TOGGLE
                            mouse_togglereset .udata2, value
                    End Select
                End If
                UART_UpdateInterruptLine uartnum

            Case &H7&
                .scratch = value
        End Select
    End With
End Sub

Public Function uart_readport(ByVal uartnum As Long, ByVal addr As Integer) As Byte
    Dim ret As Byte

    If Not UART_IsValid(uartnum) Then
        uart_readport = 0&
        Exit Function
    End If

    ret = 0&
    addr = (addr And &H7&)

    With uart_devs(uartnum)
        Select Case addr
            Case &H0&
                If .dlab = 0& Then
                    ret = .rx
                    .lsr = CByte(.lsr And ((Not UART_LSR_DR) And &HFF&))
                    .pendirq = CByte(.pendirq And ((Not UART_PENDING_RX) And &HFF&))
                    UART_UpdateInterruptLine uartnum
                    If UART_IsMouse(uartnum) Then
                        .rxnew = 0&
                    Else
                        UART_LoadNextRxByte uartnum
                    End If
                Else
                    ret = CByte(.divisor And &HFF&)
                End If

            Case &H1&
                If .dlab = 0& Then
                    ret = .ien
                Else
                    ret = CByte((.divisor And &HFF00&) \ &H100&)
                End If

            Case &H2&
                ret = IIf(.pendirq <> 0&, 0&, 1&)
                If (.pendirq And UART_PENDING_LSR) <> 0& Then
                    ret = (ret Or &H6&)
                ElseIf (.pendirq And UART_PENDING_RX) <> 0& Then
                    ret = (ret Or &H4&)
                ElseIf (.pendirq And UART_PENDING_TX) <> 0& Then
                    ret = (ret Or &H2&)
                    .pendirq = CByte(.pendirq And ((Not UART_PENDING_TX) And &HFF&))
                ElseIf (.pendirq And UART_PENDING_MSR) <> 0& Then
                    ' No-op, mirrors C behavior.
                End If

                UART_UpdateInterruptLine uartnum

            Case &H3&
                ret = .lcr

            Case &H4&
                ret = .mcr

            Case &H5&
                ret = .lsr
                .pendirq = CByte(.pendirq And ((Not UART_PENDING_LSR) And &HFF&))
                If (.lsr And &H1F&) <> 0& Then
                    .lsr = CByte(.lsr And &HE1&)
                End If
                UART_UpdateInterruptLine uartnum

            Case &H6&
                ret = (.msr And &HF0&)
                If (.msr And &H80&) <> (.lastmsr And &H80&) Then ret = (ret Or &H8&)
                If (.msr And &H20&) <> (.lastmsr And &H20&) Then ret = (ret Or &H2&)
                If (.msr And &H10&) <> (.lastmsr And &H10&) Then ret = (ret Or &H1&)

                .lastmsr = .msr
                .pendirq = CByte(.pendirq And ((Not UART_PENDING_MSR) And &HFF&))
                UART_UpdateInterruptLine uartnum

            Case &H7&
                ret = .scratch
        End Select
    End With

    uart_readport = ret
End Function

Public Sub uart_rxdata(ByVal uartnum As Long, ByVal value As Byte)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        If .mcrCbKind = UART_MCRCB_MOUSE_TOGGLE Then
            UART_QueueMouseReceiveByte uartnum, value
        ElseIf .rxnew = 0& Then
            UART_SetCurrentRxByte uartnum, value
        ElseIf UART_RxBufferedCount(uartnum) < UART_RX_BUFFER_LEN Then
            uart_rxbuf(uartnum, .rxTail) = value
            .rxTail = (.rxTail + 1&) Mod UART_RX_BUFFER_LEN
            .rxCount = .rxCount + 1&
        Else
            .lsr = CByte((.lsr Or UART_LSR_OE) And &HFF&)
            UART_TriggerLsrIrq uartnum
        End If
    End With
End Sub

Public Sub uart_receiveTick(ByVal uartnum As Long)
    Dim deliverRx As Byte

    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        If .rxStageValid = 0& Then
            UART_DisableReceiveTimer uartnum
            Exit Sub
        End If

        If .rxnew <> 0& Then
            .lsr = CByte((.lsr Or UART_LSR_OE Or UART_LSR_DR) And &HFF&)
            .rxStageValid = 0&
        Else
            .rx = .rxStage
            .rxnew = 1&
            .rxStageValid = 0&
            .lsr = CByte((.lsr Or UART_LSR_DR) And &HFF&)
            deliverRx = 1&
        End If
    End With

    If (uart_devs(uartnum).lsr And UART_LSR_OE) <> 0& Then
        UART_TriggerLsrIrq uartnum
    End If
    If deliverRx <> 0& Then UART_TriggerRxIrq uartnum

    UART_AdvanceMouseStage uartnum
End Sub

Public Function uart_isInitialized(ByVal uartnum As Long) As Byte
    uart_isInitialized = IIf(UART_IsValid(uartnum), 1&, 0&)
End Function
Public Function uart_isRxNew(ByVal uartnum As Long) As Byte
    If Not UART_IsValid(uartnum) Then
        uart_isRxNew = 0&
        Exit Function
    End If

    uart_isRxNew = uart_devs(uartnum).rxnew
End Function

Public Function uart_canAcceptRx(ByVal uartnum As Long) As Byte
    If Not UART_IsValid(uartnum) Then
        uart_canAcceptRx = 0&
        Exit Function
    End If

    If UART_RxBufferedCount(uartnum) < UART_RX_BUFFER_LEN Then
        uart_canAcceptRx = 1&
    Else
        uart_canAcceptRx = 0&
    End If
End Function

Public Sub uart_flushRx(ByVal uartnum As Long)
    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        .rx = 0&
        .rxStage = 0&
        .rxnew = 0&
        .rxStageValid = 0&
        .rxTimerActive = 0&
        .rxHead = 0&
        .rxTail = 0&
        .rxCount = 0&
        .pendirq = CByte(.pendirq And ((Not UART_PENDING_RX) And &HFF&))
        .pendirq = CByte(.pendirq And ((Not UART_PENDING_LSR) And &HFF&))
        .lsr = (UART_LSR_THRE Or UART_LSR_TEMT)
    End With

    UART_DisableReceiveTimer uartnum
    UART_UpdateInterruptLine uartnum
End Sub

Public Function uart_getIen(ByVal uartnum As Long) As Byte
    If Not UART_IsValid(uartnum) Then
        uart_getIen = 0&
        Exit Function
    End If

    uart_getIen = uart_devs(uartnum).ien
End Function

Public Function uart_getMcr(ByVal uartnum As Long) As Byte
    If Not UART_IsValid(uartnum) Then
        uart_getMcr = 0&
        Exit Function
    End If

    uart_getMcr = uart_devs(uartnum).mcr
End Function

Public Function uart_getMsr(ByVal uartnum As Long) As Byte
    If Not UART_IsValid(uartnum) Then
        uart_getMsr = 0&
        Exit Function
    End If

    uart_getMsr = uart_devs(uartnum).msr
End Function

Public Sub uart_setMsrSignals(ByVal uartnum As Long, ByVal value As Byte)
    Dim oldSignals As Byte
    Dim newSignals As Byte

    If Not UART_IsValid(uartnum) Then Exit Sub

    With uart_devs(uartnum)
        oldSignals = (.msr And &HF0&)
        newSignals = (value And &HF0&)

        If oldSignals = newSignals Then Exit Sub

        .msr = CByte((.msr And &HF&) Or newSignals)

        If (.ien And UART_IRQ_MSR_ENABLE) <> 0& Then
            .pendirq = (.pendirq Or UART_PENDING_MSR)
        End If
    End With

    UART_UpdateInterruptLine uartnum
End Sub

Public Sub uart_setMsr(ByVal uartnum As Long, ByVal value As Byte)
    If Not UART_IsValid(uartnum) Then Exit Sub
    uart_devs(uartnum).msr = value
End Sub

Public Sub uart_orPendIrq(ByVal uartnum As Long, ByVal mask As Byte)
    If Not UART_IsValid(uartnum) Then Exit Sub
    uart_devs(uartnum).pendirq = (uart_devs(uartnum).pendirq Or mask)
    UART_UpdateInterruptLine uartnum
End Sub

Public Sub uart_init(ByRef machine As MACHINE_t, ByVal uartnum As Long, ByVal base As Long, ByVal irq As Byte, ByVal mode As String)
    Dim initDev As UART_t

    If (uartnum < 0&) Or (uartnum >= UART_MAX) Then
        debug_log DEBUG_ERROR, "[UART] Invalid UART index: " & CStr(uartnum)
        Exit Sub
    End If

    UART_InitMasks

    debug_log DEBUG_INFO, "[UART] Initializing 8250 UART at base port 0x" & Right$("000" & Hex$(base And &HFFFF&), 3&) & ", IRQ " & CStr(irq)

    uart_devs(uartnum) = initDev
    uart_used(uartnum) = 1&

    uart_devs(uartnum).i8259Slot = machine.i8259
    uart_devs(uartnum).irq = irq
    uart_devs(uartnum).lsr = (UART_LSR_THRE Or UART_LSR_TEMT)
    uart_devs(uartnum).msr = &H30&
    uart_devs(uartnum).lastmsr = &H30&

    Select Case LCase$(mode)
        Case "mouse"
            uart_devs(uartnum).mcrCbKind = UART_MCRCB_MOUSE_TOGGLE
            uart_devs(uartnum).udata2 = 0&
            uart_devs(uartnum).msr = 0&
            uart_devs(uartnum).lastmsr = 0&

        Case "tcpmodem"
            uart_devs(uartnum).txCbKind = UART_TXCB_TCPMODEM
            uart_devs(uartnum).udata = uartnum
    End Select

    uart_devs(uartnum).rxTimerId = timing_addTimer(TIMER_CB_UART_RX, uartnum, 1#, TIMING_DISABLED)
    UART_UpdateReceiveTiming uartnum

    ports_cbRegister base, 8&, PORTS_CB_UART, PORTS_CB_NONE, PORTS_CB_UART, PORTS_CB_NONE, uartnum
End Sub


