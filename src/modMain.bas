Attribute VB_Name = "modMain"
Option Explicit

Public title As String
Public ops As Double
Public instructionsperloop As Long
Public cpuLimitTimer As Long
Public goCPU As Byte
Public limitCPU As Byte
Public machine As MACHINE_t

Public Sub optimer(ByVal dummy As Long)
    ops = ops / 10000#

    If showMIPS <> 0& Then
        debug_log DEBUG_INFO, CStr(Fix(ops / 10#)) & "." & CStr(Fix(ops) Mod 10&) & " MIPS"
    End If

    ops = 0#
    diag_tick machine
End Sub

Public Sub cputimer(ByVal dummy As Long)
    goCPU = 1&
End Sub

Public Sub setspeed(ByVal mhz As Double)
    If mhz > 0# Then
        speed = mhz
        instructionsperloop = CLng((speed * 1000000#) / 140000#)
        If instructionsperloop < 1& Then instructionsperloop = 1&
        limitCPU = 1&
        debug_log DEBUG_INFO, "[MACHINE] Throttling speed to approximately a " & Format$(speed, "0.00") & " MHz 8088 (" & CStr(instructionsperloop * 10000&) & " instructions/sec)"
        timing_timerEnable cpuLimitTimer
    Else
        speed = 0#
        instructionsperloop = 100&
        limitCPU = 0&
        timing_timerDisable cpuLimitTimer
    End If
End Sub

Public Sub Main()
    Dim curloop As Long
    Dim ev As Long

    debug_init

    useMachine = "award495"
    cmosOverride = vbNullString
    baudrate = 115200
    guestRamMB = DEFAULT_GUEST_RAM_MB
    instructionsperloop = 50&
    videocard = &HFF&
    showMIPS = 0&
    goCPU = 1&
    limitCPU = 0&
    speed = 0#
    speedarg = 0#
    vga_lockFPS = 0#
    running = 1&

    title = STR_TITLE & " v" & STR_VERSION

    debug_log DEBUG_INFO, title & " (c)2025 Mike Chambers" & vbCrLf
    debug_log DEBUG_INFO, "[An x86 PC emulator written in Visual Basic 6]" & vbCrLf
    debug_log DEBUG_INFO, vbCrLf

    ports_init
    Call timing_init
    Call memory_init
    menus_setMachine machine

    machine.i8259 = -1&
    machine.i8259_slave = -1&
    machine.pcap_if = -1&
    machine.hwflags_lo = 0&
    machine.hwflags_hi = 0&
    machine.extmem = 0&

    If args_parse(machine) <> 0& Then
        Exit Sub
    End If

    On Error GoTo ConsoleLoadErr
    Load frmConsole
    frmConsole.Show
    DoEvents
    On Error GoTo 0

    If console_init(title) <> 0& Then
        debug_log DEBUG_ERROR, "[ERROR] Console initialization failure"
        Exit Sub
    End If

    menus_refreshScsiMenu

    If audio_init(machine) <> 0& Then
        debug_log DEBUG_INFO, "[WARNING] Audio initialization failure"
    End If

    diag_init machine

    If machine_init(machine, useMachine) < 0& Then
        debug_log DEBUG_ERROR, "[ERROR] Machine initialization failure"
        debug_log DEBUG_ERROR, "[ERROR] Verify BIOS/CMOS files exist and relative paths resolve (e.g. roms/... and cmos/...)"
        running = 0&
        pcap_shutdown
        audio_shutdown
        Unload frmConsole
        Exit Sub
    End If

    Call timing_addTimer(TIMER_CB_OPTIMER, 0&, 10#, TIMING_ENABLED)
    cpuLimitTimer = timing_addTimer(TIMER_CB_CPUTIMER, 0&, 10000#, TIMING_DISABLED)

    If speed > 0# Then
        setspeed speed
    End If

    curloop = 0&

    Do While running <> 0&
        If limitCPU = 0& Then
            goCPU = 1&
        End If

        If goCPU <> 0& Then
            cpu_exec machine, instructionsperloop
            ops = ops + instructionsperloop
            goCPU = 0&
        End If

        timing_loop True
        console_pump
        audio_updateSampleTiming
        pcap_check_packets 0&

        curloop = curloop + 1&
        If curloop = 100& Then
            ev = console_loop()

            Select Case ev
                Case CONSOLE_EVENT_KEY
                    machine.KeyState.scancode = console_getScancode()
                    machine.KeyState.isNew = 1&
                    i8259_doirq machine.i8259, 1&

                Case CONSOLE_EVENT_QUIT
                    running = 0&

                Case CONSOLE_EVENT_DEBUG_1
                    diag_set_vga_verbose (diag_vga_verbose Xor 1&)
                    If diag_vga_verbose <> 0& Then
                        debug_log DEBUG_INFO, "[DIAG] VGA verbose diagnostics enabled"
                    Else
                        debug_log DEBUG_INFO, "[DIAG] VGA verbose diagnostics disabled"
                    End If

                Case CONSOLE_EVENT_DEBUG_2
                    showops = showops Xor 1&
            End Select

            curloop = 0&
        End If

        If curloop = 0& Then
            Select Case videocard
                Case VIDEO_CARD_ET4000
                    If et4000_doBlitNow = True Then
                        et4000_sendBlit
                        et4000_doBlitNow = False
                    End If
                Case VIDEO_CARD_VGA
                    If vga_doBlitNow = True Then
                        vga_sendBlit
                        vga_doBlitNow = False
                    End If
            End Select
        End If
    Loop

    pcap_shutdown
    audio_shutdown
    Unload frmConsole
    Exit Sub

ConsoleLoadErr:
    debug_log DEBUG_ERROR, "[ERROR] Unable to load display console form"
End Sub



