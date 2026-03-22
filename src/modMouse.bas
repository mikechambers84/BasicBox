Attribute VB_Name = "modMouse"
Option Explicit

Public Const MOUSE_ACTION_MOVE As Byte = 0&
Public Const MOUSE_ACTION_LEFT As Byte = 1&
Public Const MOUSE_ACTION_RIGHT As Byte = 2&

Public Const MOUSE_PRESSED As Byte = 0&
Public Const MOUSE_UNPRESSED As Byte = 1&
Public Const MOUSE_NEITHER As Byte = 2&

Public Const MOUSE_DEFAULT_BAUD As Long = 1200&
Private Const MOUSE_BITS_PER_BYTE As Double = 10#
Private Const MOUSE_PACKET_LEN As Long = 3&
Private Const MOUSE_TX_BUFFER_LEN As Long = 4&

Private Type MOUSE_t
    left As Byte
    right As Byte
End Type

Private mouse_state As MOUSE_t
Private mouse_uart As Long
Private mouse_lasttoggle As Byte
Private mouse_pendingX As Long
Private mouse_pendingY As Long
Private mouse_buttonsDirty As Byte
Private mouse_timerId As Long
Private mouse_timerRunning As Byte
Private mouse_txbuf(0& To MOUSE_TX_BUFFER_LEN - 1&) As Byte
Private mouse_txLen As Long
Private mouse_txPos As Long

Private Function mouse_isPowered() As Boolean
    mouse_isPowered = True
End Function

Private Sub mouse_enableTimer()
    If mouse_timerId = TIMING_ERROR Then Exit Sub
    If mouse_timerRunning <> 0& Then Exit Sub

    timing_timerEnable mouse_timerId
    mouse_timerRunning = 1&
End Sub

Private Sub mouse_disableTimer()
    If mouse_timerId = TIMING_ERROR Then Exit Sub

    timing_timerDisable mouse_timerId
    mouse_timerRunning = 0&
End Sub

Private Sub mouse_clearPendingState()
    mouse_txLen = 0&
    mouse_txPos = 0&
    mouse_pendingX = 0&
    mouse_pendingY = 0&
    mouse_buttonsDirty = 0&
    mouse_disableTimer
    Call uart_flushRx(mouse_uart)
End Sub

Private Function mouse_hasPendingPacket() As Boolean
    mouse_hasPendingPacket = ((mouse_pendingX <> 0&) Or (mouse_pendingY <> 0&) Or (mouse_buttonsDirty <> 0&))
End Function

Private Function mouse_isTransmitting() As Boolean
    mouse_isTransmitting = ((mouse_txLen > 0&) And (mouse_txPos < mouse_txLen))
End Function

Private Function mouse_clampPacketDelta(ByVal value As Long) As Long
    If value < -128& Then
        mouse_clampPacketDelta = -128&
    ElseIf value > 127& Then
        mouse_clampPacketDelta = 127&
    Else
        mouse_clampPacketDelta = value
    End If
End Function

Private Sub mouse_beginTransmit(ByVal byteCount As Long)
    mouse_txLen = byteCount
    mouse_txPos = 0&
    mouse_enableTimer
End Sub

Private Sub mouse_prepareResetId()
    mouse_txbuf(0&) = CByte(Asc("M"))
    mouse_beginTransmit 1&
End Sub

Private Sub mouse_preparePacket(ByVal xrel As Long, ByVal yrel As Long)
    Dim b0 As Byte

    b0 = &H40&
    b0 = b0 Or CByte(((yrel And &HC0&) \ &H10&) And &HC&)
    b0 = b0 Or CByte(((xrel And &HC0&) \ &H40&) And &H3&)
    If mouse_state.left <> 0& Then b0 = b0 Or &H20&
    If mouse_state.right <> 0& Then b0 = b0 Or &H10&

    mouse_txbuf(0&) = b0
    mouse_txbuf(1&) = CByte(xrel And &H3F&)
    mouse_txbuf(2&) = CByte(yrel And &H3F&)
    mouse_beginTransmit MOUSE_PACKET_LEN
End Sub

Private Sub mouse_preparePendingPacket()
    Dim xrel As Long
    Dim yrel As Long

    If mouse_isTransmitting() Then Exit Sub
    If mouse_hasPendingPacket() = False Then Exit Sub

    xrel = mouse_clampPacketDelta(mouse_pendingX)
    yrel = mouse_clampPacketDelta(mouse_pendingY)

    mouse_pendingX = mouse_pendingX - xrel
    mouse_pendingY = mouse_pendingY - yrel
    mouse_buttonsDirty = 0&

    Call mouse_preparePacket(xrel, yrel)
End Sub

Public Sub mouse_updateSerialDivisor(ByVal uartnum As Long, ByVal divisor As Long)
    If uartnum <> mouse_uart Then Exit Sub
    ' The Microsoft-compatible serial mouse runs at a fixed 1200 baud.
    ' 86Box only lets Logitech-style devices track host baud changes.
End Sub

Public Sub mouse_togglereset(ByVal dummy As Long, ByVal value As Byte)
    Dim newToggle As Byte

    newToggle = (value And &H3&)

    If ((newToggle And &H2&) <> 0&) And ((mouse_lasttoggle And &H2&) = 0&) Then
        Call mouse_clearPendingState
        Call mouse_prepareResetId
    End If

    mouse_lasttoggle = newToggle
    uart_setMsrSignals mouse_uart, (UART_MSR_DCD Or UART_MSR_DSR Or UART_MSR_CTS)
End Sub

Public Sub mouse_action(ByVal action As Byte, ByVal state As Byte, ByVal xrel As Long, ByVal yrel As Long)
    If uart_isInitialized(mouse_uart) = 0& Then Exit Sub

    Select Case action
        Case MOUSE_ACTION_MOVE
            If mouse_isPowered() Then
                mouse_pendingX = mouse_pendingX + xrel
                mouse_pendingY = mouse_pendingY + yrel
            End If
        Case MOUSE_ACTION_LEFT
            mouse_state.left = IIf(state = MOUSE_PRESSED, 1&, 0&)
            If mouse_isPowered() Then mouse_buttonsDirty = 1&
        Case MOUSE_ACTION_RIGHT
            mouse_state.right = IIf(state = MOUSE_PRESSED, 1&, 0&)
            If mouse_isPowered() Then mouse_buttonsDirty = 1&
    End Select

    If mouse_hasPendingPacket() Or mouse_isTransmitting() Then
        mouse_enableTimer
    End If
End Sub

Public Sub mouse_syncButtons(ByVal leftDown As Byte, ByVal rightDown As Byte)
    Dim changed As Byte

    If uart_isInitialized(mouse_uart) = 0& Then Exit Sub

    leftDown = CByte(leftDown And 1&)
    rightDown = CByte(rightDown And 1&)

    If mouse_state.left <> leftDown Then
        mouse_state.left = leftDown
        changed = 1&
    End If

    If mouse_state.right <> rightDown Then
        mouse_state.right = rightDown
        changed = 1&
    End If

    If (changed <> 0&) And mouse_isPowered() Then
        mouse_buttonsDirty = 1&
    End If

    If mouse_hasPendingPacket() Or mouse_isTransmitting() Then
        mouse_enableTimer
    End If
End Sub

Public Sub mouse_rxpoll(ByVal dummy As Long)
    If uart_isInitialized(mouse_uart) = 0& Then Exit Sub
    If (mouse_isPowered() = False) And (mouse_isTransmitting() = False) Then Exit Sub

    If mouse_isTransmitting() = False Then
        Call mouse_preparePendingPacket
        If mouse_isTransmitting() = False Then
            mouse_disableTimer
        End If
        Exit Sub
    End If

    uart_rxdata mouse_uart, mouse_txbuf(mouse_txPos)
    mouse_txPos = mouse_txPos + 1&

    If mouse_txPos >= mouse_txLen Then
        mouse_txPos = 0&
        mouse_txLen = 0&
        Call mouse_preparePendingPacket
        If mouse_isTransmitting() = False Then mouse_disableTimer
    End If
End Sub

Public Sub mouse_init(ByVal uartnum As Long, ByVal timerId As Long)
    debug_log DEBUG_INFO, "[MOUSE] Initializing Microsoft-compatible serial mouse"

    mouse_uart = uartnum
    mouse_timerId = timerId
    mouse_state.left = 0&
    mouse_state.right = 0&
    mouse_lasttoggle = 0&
    mouse_pendingX = 0&
    mouse_pendingY = 0&
    mouse_buttonsDirty = 0&
    mouse_timerRunning = 0&
    mouse_txLen = 0&
    mouse_txPos = 0&
    uart_setMsrSignals mouse_uart, (UART_MSR_DCD Or UART_MSR_DSR Or UART_MSR_CTS)
    If mouse_timerId <> TIMING_ERROR Then
        timing_updateIntervalFreq mouse_timerId, MOUSE_DEFAULT_BAUD / MOUSE_BITS_PER_BYTE
        mouse_disableTimer
    End If
End Sub

