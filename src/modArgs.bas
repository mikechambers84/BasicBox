Attribute VB_Name = "modArgs"
Option Explicit

Public Function args_isMatch(ByVal s1 As String, ByVal s2 As String) As Long
    args_isMatch = IIf(StrComp(s1, s2, vbTextCompare) = 0&, 1&, 0&)
End Function

Public Sub args_showHelp()
    args_logInfo STR_TITLE & " command line parameters:"
    args_logBlank
    args_logInfo "Machine options:"
    args_logInfo "  -machine <id>          Emulate machine definition defined by <id>."
    args_logInfo "  -cmos <file>           Override the machine's default CMOS file."
    args_logInfo "  -speed <mhz>           Run the emulated CPU at approximately <mhz> MHz."
    args_logBlank
    args_logInfo "Disk options:"
    args_logInfo "  -fd0 <file>            Insert <file> disk image as floppy 0."
    args_logInfo "  -fd1 <file>            Insert <file> disk image as floppy 1."
    args_logInfo "  -hd0 <file>            Insert <file> disk image as IDE hard disk 0."
    args_logInfo "  -hd1 <file>            Insert <file> disk image as IDE hard disk 1."
    args_logInfo "  -buslogic              Enable BusLogic BT-545S ISA SCSI controller emulation."
    args_logInfo "  -buslogic-base <hex>   Set BusLogic base port. (Default is 334h)"
    args_logInfo "  -buslogic-irq <num>    Set BusLogic IRQ. (Default is 11)"
    args_logInfo "  -buslogic-dma <num>    Set BusLogic DMA channel. (Default is 6)"
    args_logInfo "  -buslogic-bios <addr>  Set BusLogic BIOS address: off, c800, d000 or d800."
    args_logInfo "  -scsi-hd <id> <file>   Attach <file> as SCSI hard disk target <id> (0-6)."
    args_logInfo "  -scsi-cd <id> <file>   Attach <file> as SCSI CD-ROM target <id> (0-6). Use . for an empty drive."
    args_logBlank
    args_logInfo "Video options:"
    args_logInfo "  -video <type>          Use <type> (stdvga, et4000) video card emulation."
    args_logInfo "  -fpslock <FPS>         Attempt to lock video refresh to <FPS>."
    args_logBlank
    args_logInfo "Serial options:"
    args_logInfo "  -baud <value>          Baud rate for serial/tcpmodem emulation."
    args_logInfo "  -uart0 <type> [port]   UART on COM1 attached to none/mouse/tcpmodem."
    args_logInfo "  -uart1 <type> [port]   UART on COM2 attached to none/mouse/tcpmodem."
    args_logBlank
    args_logInfo "Miscellaneous options:"
    args_logInfo "  -mem <size>            Total guest RAM in MB (" & CStr(MIN_GUEST_RAM_MB) & " to " & CStr(MAX_GUEST_RAM_MB) & ", default " & CStr(DEFAULT_GUEST_RAM_MB) & ")."
    args_logInfo "  -debug <level>         NONE, ERROR, INFO, DETAIL."
    args_logInfo "  -mips                  Display live MIPS."
    args_logInfo "  -hw <name>             opl/noopl/blaster/noblaster/rtc/nortc."
    args_logInfo "  -net <id|list|user>    Configure NE2000 backend (-net user = built-in usermode gateway)."
    args_logInfo "  -h                     Show this help screen."
End Sub

Private Sub args_logInfo(ByVal message As String)
    debug_log DEBUG_INFO, message & vbCrLf
End Sub

Private Sub args_logError(ByVal message As String)
    debug_log DEBUG_ERROR, message & vbCrLf
End Sub

Private Sub args_logBlank()
    debug_log DEBUG_INFO, vbCrLf
End Sub

Private Sub args_init_buslogic_defaults(ByRef machine As MACHINE_t)
    If machine.buslogic_base = 0& Then
        machine.buslogic_base = &H334&
    End If
    If machine.buslogic_irq = 0& Then
        machine.buslogic_irq = 11&
    End If
    If machine.buslogic_dma = 0& Then
        machine.buslogic_dma = 6&
    End If
    If LenB(Trim$(Replace$(machine.buslogic_rom_path, vbNullChar, vbNullString))) = 0& Then
        machine.buslogic_rom_path = "roms/scsi/buslogic/BT-545S_BIOS.rom"
    End If
    If LenB(Trim$(Replace$(machine.buslogic_nvr_path, vbNullChar, vbNullString))) = 0& Then
        machine.buslogic_nvr_path = "nvr/bt545s.nvr"
    End If
End Sub

Private Function Args_ParseNonNegativeLong(ByVal text As String, ByRef value As Long) As Long
    Dim s As String
    Dim base As Long
    Dim i As Long
    Dim ch As String
    Dim digit As Long
    Dim accum As Double

    s = Trim$(text)
    If LenB(s) = 0& Then
        Args_ParseNonNegativeLong = 0&
        Exit Function
    End If

    base = 10&
    If Len(s) > 2& And LCase$(Left$(s, 2&)) = "0x" Then
        base = 16&
        s = Mid$(s, 3&)
    ElseIf Len(s) > 2& And LCase$(Left$(s, 2&)) = "&h" Then
        base = 16&
        s = Mid$(s, 3&)
    ElseIf Right$(LCase$(s), 1&) = "h" Then
        base = 16&
        s = Left$(s, Len(s) - 1&)
    End If

    If LenB(s) = 0& Then
        Args_ParseNonNegativeLong = 0&
        Exit Function
    End If

    accum = 0#
    For i = 1& To Len(s)
        ch = Mid$(s, i, 1&)
        If ch >= "0" And ch <= "9" Then
            digit = asc(ch) - asc("0")
        ElseIf base = 16& And ch >= "A" And ch <= "F" Then
            digit = 10& + asc(ch) - asc("A")
        ElseIf base = 16& And ch >= "a" And ch <= "f" Then
            digit = 10& + asc(ch) - asc("a")
        Else
            Args_ParseNonNegativeLong = 0&
            Exit Function
        End If
        accum = accum * CDbl(base) + CDbl(digit)
        If accum > 2147483647# Then
            Args_ParseNonNegativeLong = 0&
            Exit Function
        End If
    Next i

    value = CLng(accum)
    Args_ParseNonNegativeLong = 1&
End Function

Private Function args_parse_scsi_target_id(ByVal text As String) As Long
    Dim value As Long

    If Args_ParseNonNegativeLong(text, value) = 0& Then
        args_parse_scsi_target_id = -1&
        Exit Function
    End If
    If (value < 0&) Or (value >= BUSLOGIC_MAX_TARGETS) Then
        args_parse_scsi_target_id = -1&
    Else
        args_parse_scsi_target_id = value
    End If
End Function

Public Function args_parse(ByRef machine As MACHINE_t) As Long
    Dim argv() As String
    Dim argc As Long
    Dim i As Long

    argv = Args_BuildArgv(command$)
    argc = UBound(argv) - LBound(argv) + 1&

    args_init_buslogic_defaults machine

    For i = 1& To argc - 1&
        If args_isMatch(argv(i), "-h") <> 0& Then
            args_showHelp
            args_parse = -1&
            Exit Function

        ElseIf args_isMatch(argv(i), "-machine") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -machine. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            i = i + 1&
            If args_isMatch(argv(i), "list") <> 0& Then
                machine_list
                args_parse = -1&
                Exit Function
            End If
            useMachine = argv(i)

        ElseIf args_isMatch(argv(i), "-cmos") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -cmos. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            cmosOverride = argv(i)

        ElseIf args_isMatch(argv(i), "-speed") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -speed. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            speedarg = CDbl(val(argv(i)))

        ElseIf args_isMatch(argv(i), "-fd0") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -fd0. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            fdd_stageImage 0&, argv(i)

        ElseIf args_isMatch(argv(i), "-fd1") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -fd1. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            fdd_stageImage 1&, argv(i)

        ElseIf args_isMatch(argv(i), "-hd0") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -hd0. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            ata_insert_disk 0&, argv(i)

        ElseIf args_isMatch(argv(i), "-hd1") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -hd1. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            ata_insert_disk 1&, argv(i)

        ElseIf args_isMatch(argv(i), "-buslogic") <> 0& Then
            machine.buslogic_enabled = 1&

        ElseIf args_isMatch(argv(i), "-buslogic-base") <> 0& Then
            Dim blBase As Long

            If (i + 1&) >= argc Then
                args_logError "Parameter required for -buslogic-base. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            If Args_ParseNonNegativeLong(argv(i), blBase) = 0& Or blBase = 0& Or blBase > &HFFFF& Then
                args_logError argv(i) & " is an invalid BusLogic base port"
                args_parse = -1&
                Exit Function
            End If
            machine.buslogic_base = blBase
            machine.buslogic_enabled = 1&

        ElseIf args_isMatch(argv(i), "-buslogic-irq") <> 0& Then
            Dim blIrq As Long

            If (i + 1&) >= argc Then
                args_logError "Parameter required for -buslogic-irq. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            If Args_ParseNonNegativeLong(argv(i), blIrq) = 0& Or blIrq < 3& Or blIrq > 15& Then
                args_logError argv(i) & " is an invalid BusLogic IRQ"
                args_parse = -1&
                Exit Function
            End If
            machine.buslogic_irq = CByte(blIrq And &HFF&)
            machine.buslogic_enabled = 1&

        ElseIf args_isMatch(argv(i), "-buslogic-dma") <> 0& Then
            Dim blDma As Long

            If (i + 1&) >= argc Then
                args_logError "Parameter required for -buslogic-dma. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            If Args_ParseNonNegativeLong(argv(i), blDma) = 0& Or blDma < 5& Or blDma > 7& Then
                args_logError argv(i) & " is an invalid BusLogic DMA channel"
                args_parse = -1&
                Exit Function
            End If
            machine.buslogic_dma = CByte(blDma And &HFF&)
            machine.buslogic_enabled = 1&

        ElseIf args_isMatch(argv(i), "-buslogic-bios") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -buslogic-bios. Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            If args_isMatch(argv(i), "off") <> 0& Then
                machine.buslogic_bios_addr = 0&
            ElseIf args_isMatch(argv(i), "c800") <> 0& Then
                machine.buslogic_bios_addr = &HC8000
            ElseIf args_isMatch(argv(i), "d000") <> 0& Then
                machine.buslogic_bios_addr = &HD0000
            ElseIf args_isMatch(argv(i), "d800") <> 0& Then
                machine.buslogic_bios_addr = &HD8000
            Else
                args_logError argv(i) & " is an invalid BusLogic BIOS address"
                args_parse = -1&
                Exit Function
            End If
            machine.buslogic_enabled = 1&

        ElseIf (args_isMatch(argv(i), "-scsi-hd") <> 0&) Or (args_isMatch(argv(i), "-scsi-cd") <> 0&) Then
            Dim targetId As Long

            If (i + 2&) >= argc Then
                args_logError "Parameters required for " & argv(i) & ". Use -h for help."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&
            targetId = args_parse_scsi_target_id(argv(i))
            If targetId < 0& Then
                args_logError argv(i) & " is an invalid SCSI target ID"
                args_parse = -1&
                Exit Function
            End If
            machine.scsi_targets(targetId).present = 1&
            If args_isMatch(argv(i - 1&), "-scsi-hd") <> 0& Then
                machine.scsi_targets(targetId).targetType = BUSLOGIC_TARGET_DISK
            Else
                machine.scsi_targets(targetId).targetType = BUSLOGIC_TARGET_CDROM
            End If
            i = i + 1&
            If (machine.scsi_targets(targetId).targetType = BUSLOGIC_TARGET_CDROM) And (StrComp(argv(i), ".", vbBinaryCompare) = 0&) Then
                machine.scsi_targets(targetId).path = vbNullString
            Else
                machine.scsi_targets(targetId).path = argv(i)
            End If
            machine.buslogic_enabled = 1&

        ElseIf args_isMatch(argv(i), "-video") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -video. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            If args_isMatch(argv(i + 1&), "stdvga") <> 0& Then
                videocard = VIDEO_CARD_VGA
            ElseIf args_isMatch(argv(i + 1&), "et4000") <> 0& Then
                videocard = VIDEO_CARD_ET4000
            Else
                args_logError argv(i + 1&) & " is an invalid video card option. Valid values are stdvga, et4000."
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&

        ElseIf args_isMatch(argv(i), "-fpslock") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -fpslock. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            i = i + 1&
            vga_lockFPS = CDbl(val(argv(i)))
            If (vga_lockFPS < 1#) Or (vga_lockFPS > 144#) Then
                args_logError CStr(vga_lockFPS) & " is an invalid FPS option, valid range is 1 to 144"
                args_parse = -1&
                Exit Function
            End If

        ElseIf args_isMatch(argv(i), "-mem") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -mem. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            i = i + 1&
            guestRamMB = CLng(Fix(val(argv(i))))
            If (guestRamMB < MIN_GUEST_RAM_MB) Or (guestRamMB > MAX_GUEST_RAM_MB) Then
                args_logError "-mem must be between " & CStr(MIN_GUEST_RAM_MB) & " and " & CStr(MAX_GUEST_RAM_MB) & " MB."
                args_parse = -1&
                Exit Function
            End If

        ElseIf args_isMatch(argv(i), "-debug") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -debug. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            If args_isMatch(argv(i + 1&), "none") <> 0& Then
                debug_setLevel DEBUG_NONE
            ElseIf args_isMatch(argv(i + 1&), "error") <> 0& Then
                debug_setLevel DEBUG_ERROR
            ElseIf args_isMatch(argv(i + 1&), "info") <> 0& Then
                debug_setLevel DEBUG_INFO
            ElseIf args_isMatch(argv(i + 1&), "detail") <> 0& Then
                debug_setLevel DEBUG_DETAIL
            Else
                args_logError argv(i + 1&) & " is an invalid debug option"
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&

        ElseIf args_isMatch(argv(i), "-mips") <> 0& Then
            showMIPS = 1&

        ElseIf args_isMatch(argv(i), "-baud") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -baud. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            i = i + 1&
            baudrate = CLng(val(argv(i)))
            If (baudrate < 300&) Or (baudrate > 115200) Then
                args_logError "Baud rate must be between 300 and 115200."
                args_parse = -1&
                Exit Function
            End If

        ElseIf (args_isMatch(argv(i), "-uart0") <> 0&) Or (args_isMatch(argv(i), "-uart1") <> 0&) Then
            Dim uartnum As Long
            Dim base As Long
            Dim irq As Byte
            Dim listenPort As Long

            uartnum = CLng(Right$(argv(i), 1&))
            If uartnum = 0& Then
                base = &H3F8&
                irq = 4&
                machine_hwflag_set machine, 0&, MACHINE_HW_SKIP_UART0_HI
            Else
                base = &H2F8&
                irq = 3&
                machine_hwflag_set machine, 0&, MACHINE_HW_SKIP_UART1_HI
            End If

            If (i + 1&) >= argc Then
                args_logError "Parameter required for -uart" & CStr(uartnum) & ". Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            If args_isMatch(argv(i + 1&), "tcpmodem") <> 0& Then
                i = i + 1&
                If (i + 1&) >= argc Then
                    listenPort = 23&
                ElseIf Left$(argv(i + 1&), 1&) = "-" Then
                    listenPort = 23&
                Else
                    i = i + 1&
                    listenPort = CLng(val(argv(i)))
                End If

                uart_init machine, uartnum, base, irq, "tcpmodem"
                tcpmodem_init uartnum, listenPort
                timing_addTimer TIMER_CB_TCPMODEM_RXPOLL, uartnum, baudrate / 9#, TIMING_ENABLED

            ElseIf args_isMatch(argv(i + 1&), "mouse") <> 0& Then
                i = i + 1&
                uart_init machine, uartnum, base, irq, "mouse"
                Call mouse_init(uartnum, timing_addTimer(TIMER_CB_MOUSE_RXPOLL, uartnum, MOUSE_DEFAULT_BAUD / 10#, TIMING_DISABLED))

            ElseIf args_isMatch(argv(i + 1&), "none") <> 0& Then
                i = i + 1&
                uart_init machine, uartnum, base, irq, "none"

            Else
                args_logError argv(i + 1&) & " is not a valid parameter for -uart" & CStr(uartnum) & ". Use -h for help."
                args_parse = -1&
                Exit Function
            End If

        ElseIf args_isMatch(argv(i), "-hw") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -hw. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            If args_isMatch(argv(i + 1&), "opl") <> 0& Then
                machine_hwflag_set machine, MACHINE_HW_OPL_LO, 0&
            ElseIf args_isMatch(argv(i + 1&), "noopl") <> 0& Then
                machine_hwflag_set machine, 0&, MACHINE_HW_SKIP_OPL_HI
            ElseIf args_isMatch(argv(i + 1&), "blaster") <> 0& Then
                machine_hwflag_set machine, MACHINE_HW_BLASTER_LO, 0&
            ElseIf args_isMatch(argv(i + 1&), "noblaster") <> 0& Then
                machine_hwflag_set machine, 0&, MACHINE_HW_SKIP_BLASTER_HI
            ElseIf args_isMatch(argv(i + 1&), "rtc") <> 0& Then
                machine_hwflag_set machine, MACHINE_HW_RTC_LO, 0&
            ElseIf args_isMatch(argv(i + 1&), "nortc") <> 0& Then
                machine_hwflag_set machine, 0&, MACHINE_HW_SKIP_RTC_HI
            Else
                args_logError argv(i + 1&) & " is an invalid hardware option"
                args_parse = -1&
                Exit Function
            End If
            i = i + 1&

        ElseIf args_isMatch(argv(i), "-net") <> 0& Then
            If (i + 1&) >= argc Then
                args_logError "Parameter required for -net. Use -h for help."
                args_parse = -1&
                Exit Function
            End If

            i = i + 1&
            If args_isMatch(argv(i), "list") <> 0& Then
                pcap_listdevs
                args_parse = -1&
                Exit Function
            End If

            If args_isMatch(argv(i), "user") <> 0& Then
                machine.pcap_if = PCAP_IF_USERNET
            Else
                machine.pcap_if = CLng(val(argv(i)))
            End If
            machine_hwflag_set machine, MACHINE_HW_NE2000_LO, 0&

        Else
            args_logError argv(i) & " is not a valid parameter. Use -h for help."
            args_parse = -1&
            Exit Function
        End If
    Next i

    args_parse = 0&
End Function

Private Function Args_BuildArgv(ByVal cmd As String) As String()
    Dim args() As String
    Dim cur As String
    Dim inQuote As Boolean
    Dim i As Long
    Dim ch As String
    Dim count As Long

    ReDim args(0& To 0&) As String
    args(0&) = "BasicBox"
    count = 1&

    cur = vbNullString
    inQuote = False

    For i = 1& To Len(cmd)
        ch = Mid$(cmd, i, 1&)

        If ch = """" Then
            inQuote = Not inQuote
        ElseIf (ch = " " Or ch = vbTab) And (Not inQuote) Then
            If LenB(cur) <> 0& Then
                ReDim Preserve args(0& To count) As String
                args(count) = cur
                count = count + 1&
                cur = vbNullString
            End If
        Else
            cur = cur & ch
        End If
    Next i

    If LenB(cur) <> 0& Then
        ReDim Preserve args(0& To count) As String
        args(count) = cur
    End If

    Args_BuildArgv = args
End Function


