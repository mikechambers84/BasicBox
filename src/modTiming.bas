Attribute VB_Name = "modTiming"
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private timing_threadStarted As Byte
Private timing_threadHandle As Long
Private timing_threadId As Long
Public timing_pendingDispatch As Boolean

Public Const TIMING_ENABLED As Byte = 1&
Public Const TIMING_DISABLED As Byte = 0&
Public Const TIMING_ERROR As Long = -1&
Public Const TIMING_RINGSIZE As Long = 1024&
Public Const TIMING_MAX_DISPATCH_BURST As Long = 25&

Public Const TIMER_CB_NONE As Long = 0&
Public Const TIMER_CB_OPTIMER As Long = 1&
Public Const TIMER_CB_CPUTIMER As Long = 2&
Public Const TIMER_CB_AUDIO_GENERATE As Long = 3&
Public Const TIMER_CB_CONSOLE_KEYREPEAT As Long = 4&
Public Const TIMER_CB_I8253_TICK As Long = 5&
Public Const TIMER_CB_I8255_REFRESH As Long = 6&
Public Const TIMER_CB_PCSPEAKER As Long = 7&
Public Const TIMER_CB_CMOSRTC_TICK As Long = 8&
Public Const TIMER_CB_MOUSE_RXPOLL As Long = 9&
Public Const TIMER_CB_TCPMODEM_RXPOLL As Long = 10&
Public Const TIMER_CB_TCPMODEM_RINGER As Long = 11&
Public Const TIMER_CB_NE2000_TX As Long = 12&
Public Const TIMER_CB_ATA_DELAYED_IRQ As Long = 13&
Public Const TIMER_CB_ATA_RESET As Long = 14&
Public Const TIMER_CB_FDC_CONTROLLER As Long = 15&
Public Const TIMER_CB_FDD_POLL As Long = 16&
Public Const TIMER_CB_MENUS_RESET As Long = 17&
Public Const TIMER_CB_VGA_BLINK As Long = 22&
Public Const TIMER_CB_VGA_HBLANK As Long = 23&
Public Const TIMER_CB_VGA_HBLANK_END As Long = 24&
Public Const TIMER_CB_VGA_DRAW As Long = 25&
Public Const TIMER_CB_BLASTER_DMA As Long = 26&
Public Const TIMER_CB_OPL2_TICK As Long = 27&
Public Const TIMER_CB_PCAP_POLL As Long = 28&
Public Const TIMER_CB_FDD_SEEK_COMPLETE As Long = 29&
Public Const TIMER_CB_BUSLOGIC_MAIL As Long = 30&
Public Const TIMER_CB_BUSLOGIC_RESET As Long = 31&
Public Const TIMER_CB_CONSOLE_CTRLALTDEL As Long = 32&
Public Const TIMER_CB_UART_RX As Long = 33&
Public Const TIMER_CB_ET4000_BLINK As Long = 34&
Public Const TIMER_CB_ET4000_HBLANK As Long = 35&
Public Const TIMER_CB_ET4000_HBLANK_END As Long = 36&
Public Const TIMER_CB_ET4000_DRAW As Long = 37&

Private Type TIMER_t
    interval As Double
    previous As Double
    enabled As Byte
    callbackId As Long
    data As Long
End Type

Public timing_cur As Double
Public timing_freq As Double

Private timers() As TIMER_t
Private timers_count As Long

Private Sub timing_startPollThread()
    If timing_threadStarted <> 0& Then Exit Sub

    timing_threadHandle = CreateThread(ByVal 0&, 0&, AddressOf timing_pollThreadProc, ByVal 0&, 0&, timing_threadId)
    If timing_threadHandle = 0& Then
        timing_threadStarted = 0&
        debug_log DEBUG_INFO, "[TIMING] Polling thread start failed, using main thread polling"
        Exit Sub
    End If

    timing_threadStarted = 1&
    CloseHandle timing_threadHandle
    timing_threadHandle = 0&
End Sub

Public Function timing_pollThreadProc(ByVal lpParam As Long) As Long
    Dim i As Long
    Dim cur As Currency

    Do While running <> 0&
        If QueryPerformanceCounter(cur) <> 0& Then
            timing_cur = cur * 10000#
        Else
            timing_cur = timer * 1000000#
        End If
        
        For i = 0& To timers_count - 1&
            If timing_cur >= (timers(i).previous + timers(i).interval) Then
                If timers(i).enabled <> TIMING_DISABLED Then
                    timing_pendingDispatch = True
                End If
            End If
        Next i
    Loop

    timing_pollThreadProc = 0&
End Function

Public Function timing_init() As Long
    Dim freq As Currency

    If QueryPerformanceFrequency(freq) <> 0& Then
        timing_freq = freq * 10000#
    Else
        timing_freq = 1000000#
    End If

    timing_cur = timing_getCur()
    timers_count = 0&
    ReDim timers(0& To 0&) As TIMER_t
    
    timing_pendingDispatch = False
    'timing_startPollThread

    timing_init = 0&
End Function

Public Function timing_loop(ByVal doDispatch As Boolean) As Boolean
    Dim i As Long
    Dim dispatchCount As Long

    timing_loop = False

    timing_cur = timing_getCur()
    For i = 0& To timers_count - 1&
        If timers(i).enabled <> TIMING_DISABLED Then
            If timers(i).interval > 0# Then
                If timing_cur >= (timers(i).previous + timers(i).interval) Then
                    timing_loop = True
                    If (doDispatch = False) Then
                        Exit Function
                    End If

                    dispatchCount = 0&
                    Do While (timers(i).enabled <> TIMING_DISABLED) And (timers(i).interval > 0#) And (timing_cur >= (timers(i).previous + timers(i).interval))
                        Timing_InvokeTimer timers(i).callbackId, timers(i).data
                        timers(i).previous = timers(i).previous + timers(i).interval
                        dispatchCount = dispatchCount + 1&
                        If dispatchCount >= TIMING_MAX_DISPATCH_BURST Then
                            timers(i).previous = timing_cur
                            Exit Do
                        End If
                    Loop

                    If (timing_cur - timers(i).previous) >= (timers(i).interval * 100#) Then
                        timers(i).previous = timing_cur
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Sub timing_speedTest()
    ' Stub for parity with timing.c; not needed for runtime behavior.
End Sub

Public Function timing_addTimerUsingInterval(ByVal callbackId As Long, ByVal data As Long, ByVal interval As Double, ByVal enabled As Byte) As Long
    If interval <= 0# Then
        timing_addTimerUsingInterval = TIMING_ERROR
        Exit Function
    End If

    If timers_count = 0& Then
        ReDim timers(0& To 0&) As TIMER_t
    Else
        ReDim Preserve timers(0& To timers_count) As TIMER_t
    End If

    timers(timers_count).previous = timing_getCur()
    timers(timers_count).interval = interval
    timers(timers_count).callbackId = callbackId
    timers(timers_count).data = data
    timers(timers_count).enabled = enabled

    timing_addTimerUsingInterval = timers_count
    timers_count = timers_count + 1&
End Function

Public Function timing_addTimer(ByVal callbackId As Long, ByVal data As Long, ByVal frequency As Double, ByVal enabled As Byte) As Long
    If frequency <= 0# Then
        timing_addTimer = TIMING_ERROR
        Exit Function
    End If
    timing_addTimer = timing_addTimerUsingInterval(callbackId, data, timing_freq / frequency, enabled)
End Function

Public Sub timing_updateInterval(ByVal tnum As Long, ByVal interval As Double)
    If (tnum < 0&) Or (tnum >= timers_count) Then
        debug_log DEBUG_ERROR, "[ERROR] timing_updateInterval() asked to operate on invalid timer"
        Exit Sub
    End If

    timers(tnum).interval = interval
End Sub

Public Sub timing_updateIntervalFreq(ByVal tnum As Long, ByVal frequency As Double)
    If (tnum < 0&) Or (tnum >= timers_count) Then
        debug_log DEBUG_ERROR, "[ERROR] timing_updateIntervalFreq() asked to operate on invalid timer"
        Exit Sub
    End If

    If frequency <= 0# Then Exit Sub

    timers(tnum).interval = timing_freq / frequency
End Sub

Public Sub timing_timerEnable(ByVal tnum As Long)
    If (tnum < 0&) Or (tnum >= timers_count) Then
        debug_log DEBUG_ERROR, "[ERROR] timing_timerEnable() asked to operate on invalid timer"
        Exit Sub
    End If

    timers(tnum).enabled = TIMING_ENABLED
    timers(tnum).previous = timing_getCur()
End Sub

Public Sub timing_timerDisable(ByVal tnum As Long)
    If (tnum < 0&) Or (tnum >= timers_count) Then
        debug_log DEBUG_ERROR, "[ERROR] timing_timerDisable() asked to operate on invalid timer"
        Exit Sub
    End If

    timers(tnum).enabled = TIMING_DISABLED
End Sub

Public Function timing_getFreq() As Double
    timing_getFreq = timing_freq
End Function

Public Function timing_getCur() As Double
    Dim cur As Currency

    If QueryPerformanceCounter(cur) <> 0& Then
        timing_cur = cur * 10000#
    Else
        timing_cur = timer * 1000000#
    End If

    timing_getCur = timing_cur
End Function

Private Sub Timing_InvokeTimer(ByVal callbackId As Long, ByVal data As Long)
    Select Case callbackId
        Case TIMER_CB_OPTIMER
            optimer data
        Case TIMER_CB_CPUTIMER
            cputimer data
        Case TIMER_CB_AUDIO_GENERATE
            audio_generateSample data
        Case TIMER_CB_CONSOLE_KEYREPEAT
            console_keyRepeat data
        Case TIMER_CB_I8253_TICK
            i8253_tickCallback data
        Case TIMER_CB_I8255_REFRESH
            i8255_refreshToggle data
        Case TIMER_CB_PCSPEAKER
            pcspeaker_callback data
        Case TIMER_CB_CMOSRTC_TICK
            cmosrtc_tick data
        Case TIMER_CB_MOUSE_RXPOLL
            mouse_rxpoll data
        Case TIMER_CB_TCPMODEM_RXPOLL
            tcpmodem_rxpoll data
        Case TIMER_CB_TCPMODEM_RINGER
            tcpmodem_ringer data
        Case TIMER_CB_NE2000_TX
            ne2000_tx_timer data
        Case TIMER_CB_ATA_DELAYED_IRQ
            ata_delayed_irq data
        Case TIMER_CB_ATA_RESET
            ata_reset_cb data
        Case TIMER_CB_FDC_CONTROLLER
            fdc_controllerCallback data
        Case TIMER_CB_FDD_POLL
            fdd_pollCallback data
        Case TIMER_CB_MENUS_RESET
            menus_resetCallback data
        Case TIMER_CB_VGA_BLINK
            vga_blinkCallback data
        Case TIMER_CB_VGA_HBLANK
            vga_hblankCallback data
        Case TIMER_CB_VGA_HBLANK_END
            vga_hblankEndCallback data
        Case TIMER_CB_VGA_DRAW
            vga_drawCallback data
        Case TIMER_CB_BLASTER_DMA
            blaster_generateSample data
        Case TIMER_CB_OPL2_TICK
            opl2_tickOperator data
        Case TIMER_CB_PCAP_POLL
            pcap_check_packets data
        Case TIMER_CB_FDD_SEEK_COMPLETE
            fdd_seekCompleteCallback data
        Case TIMER_CB_BUSLOGIC_MAIL
            buslogic_process_mail_cb data
        Case TIMER_CB_BUSLOGIC_RESET
            buslogic_reset_timer_cb data
        Case TIMER_CB_CONSOLE_CTRLALTDEL
            console_ctrlAltDelStep data
        Case TIMER_CB_UART_RX
            uart_receiveTick data
        Case TIMER_CB_ET4000_BLINK
            et4000_blinkCallback data
        Case TIMER_CB_ET4000_HBLANK
            et4000_hblankCallback data
        Case TIMER_CB_ET4000_HBLANK_END
            et4000_hblankEndCallback data
        Case TIMER_CB_ET4000_DRAW
            et4000_drawCallback data
        Case Else
            ' Unused callback slot.
    End Select
End Sub


