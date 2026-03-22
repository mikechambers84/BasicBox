Attribute VB_Name = "modMachine"
Option Explicit

Public Const MACHINE_MEM_RAM As Byte = 0&
Public Const MACHINE_MEM_ROM As Byte = 1&
Public Const MACHINE_MEM_ROM_INTERLEAVED_LOW As Byte = 2&
Public Const MACHINE_MEM_ROM_INTERLEAVED_HIGH As Byte = 3&
Public Const MACHINE_MEM_ENDLIST As Byte = 4&

Public Const MACHINE_ROM_OPTIONAL As Byte = 0&
Public Const MACHINE_ROM_REQUIRED As Byte = 1&
Public Const MACHINE_ROM_ISNOTROM As Byte = 2&

Public Const MACHINE_HW_OPL_LO As Long = &H1&
Public Const MACHINE_HW_BLASTER_LO As Long = &H2&
Public Const MACHINE_HW_UART0_NONE_LO As Long = &H4&
Public Const MACHINE_HW_UART0_MOUSE_LO As Long = &H8&
Public Const MACHINE_HW_UART0_TCPMODEM_LO As Long = &H10&
Public Const MACHINE_HW_UART1_NONE_LO As Long = &H20&
Public Const MACHINE_HW_UART1_MOUSE_LO As Long = &H40&
Public Const MACHINE_HW_UART1_TCPMODEM_LO As Long = &H80&
Public Const MACHINE_HW_RTC_LO As Long = &H100&
Public Const MACHINE_HW_NE2000_LO As Long = &H400&

Public Const MACHINE_HW_SKIP_OPL_HI As Long = &H80000000
Public Const MACHINE_HW_SKIP_BLASTER_HI As Long = &H40000000
Public Const MACHINE_HW_SKIP_UART0_HI As Long = &H20000000
Public Const MACHINE_HW_SKIP_UART1_HI As Long = &H10000000
Public Const MACHINE_HW_SKIP_RTC_HI As Long = &H4000000

Private Const MACHDEF_INIT_GENERIC_XT As Long = 0&
Private Const MACHDEF_INIT_ASUS_386 As Long = 1&
Private Const MACHDEF_INIT_OPTI495 As Long = 2&
Private Const MACHDEF_INIT_OPTI5X7 As Long = 3&
Private Const MACHINE_CONVENTIONAL_RAM_BYTES As Long = &HA0000
Private Const MACHINE_EXTENDED_RAM_BASE As Long = &H100000

Private Function Machine_DefaultHwLo() As Long
    Machine_DefaultHwLo = (MACHINE_HW_BLASTER_LO Or MACHINE_HW_UART1_MOUSE_LO Or MACHINE_HW_RTC_LO)
End Function

Private Function Machine_InitBusLogicController(ByRef machine As MACHINE_t, ByVal picSlave As Long) As Long
    If machine.buslogic_enabled = 0& Then
        Machine_InitBusLogicController = 0&
        Exit Function
    End If

    If buslogic_init(machine, picSlave) <> 0& Then
        debug_log DEBUG_ERROR, "[SCSI] Failed to initialize BusLogic BT-545S" & vbCrLf
        Machine_InitBusLogicController = -1&
    Else
        Machine_InitBusLogicController = 0&
    End If
End Function

Public Sub machine_hwflag_set(ByRef machine As MACHINE_t, ByVal flagLo As Long, ByVal flagHi As Long)
    machine.hwflags_lo = (machine.hwflags_lo Or flagLo)
    machine.hwflags_hi = (machine.hwflags_hi Or flagHi)
End Sub

Public Function machine_hwflag_has(ByRef machine As MACHINE_t, ByVal flagLo As Long, ByVal flagHi As Long) As Boolean
    Dim okLo As Boolean
    Dim okHi As Boolean

    If flagLo = 0& Then
        okLo = True
    Else
        okLo = ((machine.hwflags_lo And flagLo) = flagLo)
    End If

    If flagHi = 0& Then
        okHi = True
    Else
        okHi = ((machine.hwflags_hi And flagHi) = flagHi)
    End If

    machine_hwflag_has = (okLo And okHi)
End Function

Private Function machine_hwflag_has_not_skip(ByRef machine As MACHINE_t, ByVal flagLo As Long, ByVal flagHi As Long, ByVal skipLo As Long, ByVal skipHi As Long) As Boolean
    machine_hwflag_has_not_skip = machine_hwflag_has(machine, flagLo, flagHi) And (Not machine_hwflag_has(machine, skipLo, skipHi))
End Function

Public Function machine_init_generic_xt(ByRef machine As MACHINE_t) As Long
    machine_init_generic_xt = Machine_InitCommon(machine, MACHDEF_INIT_GENERIC_XT, 2&, 0&, 0&)
End Function

Public Function machine_init_asus_386(ByRef machine As MACHINE_t) As Long
    machine_init_asus_386 = Machine_InitCommon(machine, MACHDEF_INIT_ASUS_386, 2&, 1&, 1&)
End Function

Public Function machine_init_opti495(ByRef machine As MACHINE_t) As Long
    machine_init_opti495 = Machine_InitCommon(machine, MACHDEF_INIT_OPTI495, 7&, 1&, 1&)
End Function

Public Function machine_init_opti5x7(ByRef machine As MACHINE_t) As Long
    machine_init_opti5x7 = Machine_InitCommon(machine, MACHDEF_INIT_OPTI5X7, 7&, 1&, 1&)
End Function

Private Function Machine_InitCommon(ByRef machine As MACHINE_t, ByVal boardKind As Long, ByVal ne2000Irq As Byte, ByVal hasSlave As Byte, ByVal resetShowOps As Byte) As Long
    Dim buslogicPicSlave As Long
    If resetShowOps <> 0& Then showops = 0&

    console_changeScancodes 2&

    machine.i8259 = -1&
    machine.i8259_slave = -1&

    i8259_init machine.i8259, &H20&, -1&
    If hasSlave <> 0& Then
        i8259_init machine.i8259_slave, &HA0&, machine.i8259
    End If

    i8253_init machine, machine.i8259
    i8237_init machine
    i8255_init machine
    pcspeaker_init machine
    i8042_init machine.cpu, machine.i8259

    Select Case boardKind
        Case 0&
            kbc.config = &H40&
            kbc.keyboard_enabled = 1&
            machine.cpu.a20_gate = 1&
        Case MACHDEF_INIT_ASUS_386
            rabbit_init
        Case MACHDEF_INIT_OPTI495
            opti495_init
        Case MACHDEF_INIT_OPTI5X7
            opti5x7_init
    End Select

    If machine_hwflag_has_not_skip(machine, MACHINE_HW_BLASTER_LO, 0&, 0&, MACHINE_HW_SKIP_BLASTER_HI) Then
        blaster_init machine
        OPL3_init machine
        machine.mixBlaster = 1&
        machine.mixOPL = 1&
    ElseIf machine_hwflag_has_not_skip(machine, MACHINE_HW_OPL_LO, 0&, 0&, MACHINE_HW_SKIP_OPL_HI) Then
        OPL3_init machine
        machine.mixOPL = 1&
    End If

    If machine_hwflag_has_not_skip(machine, MACHINE_HW_RTC_LO, 0&, 0&, MACHINE_HW_SKIP_RTC_HI) Then
        rtc_init machine.cpu
    End If

    If machine_hwflag_has_not_skip(machine, MACHINE_HW_UART0_NONE_LO, 0&, 0&, MACHINE_HW_SKIP_UART0_HI) Then
        uart_init machine, 0&, &H3F8&, 4&, "none"
    ElseIf machine_hwflag_has_not_skip(machine, MACHINE_HW_UART0_MOUSE_LO, 0&, 0&, MACHINE_HW_SKIP_UART0_HI) Then
        uart_init machine, 0&, &H3F8&, 4&, "mouse"
        Call mouse_init(0&, timing_addTimer(TIMER_CB_MOUSE_RXPOLL, 0&, MOUSE_DEFAULT_BAUD / 10#, TIMING_DISABLED))
    ElseIf machine_hwflag_has_not_skip(machine, MACHINE_HW_UART0_TCPMODEM_LO, 0&, 0&, MACHINE_HW_SKIP_UART0_HI) Then
        uart_init machine, 0&, &H3F8&, 4&, "tcpmodem"
        tcpmodem_init 0&, 23&
        timing_addTimer TIMER_CB_TCPMODEM_RXPOLL, 0&, baudrate / 9#, TIMING_ENABLED
    End If

    If machine_hwflag_has_not_skip(machine, MACHINE_HW_UART1_NONE_LO, 0&, 0&, MACHINE_HW_SKIP_UART1_HI) Then
        uart_init machine, 1&, &H2F8&, 3&, "none"
    ElseIf machine_hwflag_has_not_skip(machine, MACHINE_HW_UART1_MOUSE_LO, 0&, 0&, MACHINE_HW_SKIP_UART1_HI) Then
        uart_init machine, 1&, &H2F8&, 3&, "mouse"
        Call mouse_init(1&, timing_addTimer(TIMER_CB_MOUSE_RXPOLL, 1&, MOUSE_DEFAULT_BAUD / 10#, TIMING_DISABLED))
    ElseIf machine_hwflag_has_not_skip(machine, MACHINE_HW_UART1_TCPMODEM_LO, 0&, 0&, MACHINE_HW_SKIP_UART1_HI) Then
        uart_init machine, 1&, &H2F8&, 3&, "tcpmodem"
        tcpmodem_init 1&, 23&
        timing_addTimer TIMER_CB_TCPMODEM_RXPOLL, 1&, baudrate / 9#, TIMING_ENABLED
    End If

    If machine_hwflag_has(machine, MACHINE_HW_NE2000_LO, 0&) Then
        ne2000_init machine, &H300&, ne2000Irq
        If machine.pcap_if <> -1& Then
            If pcap_init(machine.pcap_if) <> 0& Then
                Machine_InitCommon = -1&
                Exit Function
            End If
        End If
    End If

    cpu_reset machine.cpu

    If boardKind <> MACHDEF_INIT_GENERIC_XT Then
        ata_init machine.i8259_slave
        fdd_init
        fdc_init machine.cpu, machine.i8259
    End If

    buslogicPicSlave = -1&
    If hasSlave <> 0& Then
        buslogicPicSlave = machine.i8259_slave
    End If
    If Machine_InitBusLogicController(machine, buslogicPicSlave) <> 0& Then
        Machine_InitCommon = -1&
        Exit Function
    End If

    Select Case videocard
        Case VIDEO_CARD_ET4000
            If et4000_init() <> 0& Then
                Machine_InitCommon = -1&
                Exit Function
            End If
        Case VIDEO_CARD_VGA
            If vga_init() <> 0& Then
                Machine_InitCommon = -1&
                Exit Function
            End If
        Case Else
            debug_log DEBUG_ERROR, "[VIDEO] Unsupported video card selection"
            Machine_InitCommon = -1&
            Exit Function
    End Select

    Machine_InitCommon = 0&
End Function

Public Function machine_init(ByRef machine As MACHINE_t, ByVal machineId As String) As Long
    Dim idx As Long
    Dim defDescription As String
    Dim defInitKind As Long
    Dim defCmos As String
    Dim defVideo As Byte
    Dim defSpeed As Double
    Dim defHwLo As Long
    Dim defHwHi As Long
    Dim ret As Long

    idx = Machine_GetDefIndex(machineId)
    If idx < 0& Then
        debug_log DEBUG_ERROR, "[MACHINE] ERROR: Machine definition not found: " & machineId
        machine_init = -1&
        Exit Function
    End If

    Machine_GetDef idx, defDescription, defInitKind, defCmos, defVideo, defSpeed, defHwLo, defHwHi

    debug_log DEBUG_INFO, "[MACHINE] Initializing machine: """ & defDescription & """ (" & Machine_GetDefId(idx) & ")"

    If Machine_LoadMemoryMap(machine, idx) <> 0& Then
        machine_init = -1&
        Exit Function
    End If

    If LenB(defCmos) <> 0& Then
        If LenB(cmosOverride) <> 0& Then
            defCmos = cmosOverride
        End If
        cmosrtc_init defCmos, machine.i8259_slave
    End If

    machine_hwflag_set machine, defHwLo, defHwHi

    If videocard = &HFF& Then
        videocard = defVideo
    End If

    If speedarg > 0# Then
        speed = speedarg
    ElseIf speedarg < 0# Then
        speed = -1#
    Else
        speed = defSpeed
    End If

    Select Case defInitKind
        Case MACHDEF_INIT_GENERIC_XT
            ret = machine_init_generic_xt(machine)
        Case MACHDEF_INIT_ASUS_386
            ret = machine_init_asus_386(machine)
        Case MACHDEF_INIT_OPTI495
            ret = machine_init_opti495(machine)
        Case MACHDEF_INIT_OPTI5X7
            ret = machine_init_opti5x7(machine)
        Case Else
            ret = -1&
    End Select

    If ret <> 0& Then
        machine_init = -1&
        Exit Function
    End If

    machine_init = idx
End Function

Public Sub machine_list()
    Dim i As Long

    debug_log DEBUG_INFO, "Valid " & STR_TITLE & " machines:" & vbCrLf
    For i = 0& To 9&
        debug_log DEBUG_INFO, Machine_GetDefId(i) & ": """ & Machine_GetDefDescription(i) & """" & vbCrLf
    Next i
End Sub

Private Function Machine_LoadMemoryMap(ByRef machine As MACHINE_t, ByVal idx As Long) As Long
    Dim highRamBytes As Long

    If Machine_MapRam(&H0&, MACHINE_CONVENTIONAL_RAM_BYTES) <> 0& Then GoTo Fail

    machine.extmem = 0&
    If guestRamMB > 1& Then
        machine.extmem = (guestRamMB - 1&) * 1024&
        highRamBytes = (guestRamMB - 1&) * MACHINE_EXTENDED_RAM_BASE
    End If

    Select Case idx
        Case 0& ' ami386
            If Machine_MapRom(&HF0000, &H10000, "roms/machine/ami386/ami386.bin", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 1& ' award495
            If Machine_MapRom(&HF0000, &H10000, "roms/machine/award495/opt495s.awa", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 2& ' seabios
            If Machine_MapRom(&HE0000, &H20000, "roms/machine/seabios/bios.bin", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 3& ' award486
            If Machine_MapRom(&HF0000, &H10000, "roms/machine/award486/ah4-02-5eed236c7c5ab670872178.bin", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 4& ' mrbios486
            If Machine_MapRom(&HF0000, &H10000, "roms/machine/mrbios486/opt495sx.mr", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 5& ' 4saw2
            If Machine_MapRom(&HE0000, &H20000, "roms/machine/4saw2/4saw0911.bin", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 6& ' hot543
            If Machine_MapRom(&HE0000, &H20000, "roms/machine/hot543/543_R21.BIN", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 7& ' sp97xv
            If Machine_MapRom(&HE0000, &H20000, "roms/machine/sp97xv/0109XVJ2.BIN", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 8& ' p5sp4
            If Machine_MapRom(&HE0000, &H20000, "roms/machine/p5sp4/0106.001", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case 9& ' genericxt
            If Machine_MapRom(&HFE000, &H2000&, "roms/machine/generic_xt/pcxtbios.bin", MACHINE_ROM_REQUIRED) <> 0& Then GoTo Fail

        Case Else
            GoTo Fail
    End Select

    If highRamBytes > 0& Then
        If Machine_MapRam(MACHINE_EXTENDED_RAM_BASE, highRamBytes) <> 0& Then GoTo Fail
    End If

    Machine_LoadMemoryMap = 0&
    Exit Function

Fail:
    Machine_LoadMemoryMap = -1&
End Function

Private Function Machine_MapRam(ByVal start As Long, ByVal size As Long) As Long
    Dim buf() As Byte
    Dim v As Variant

    On Error GoTo MapFail

    ReDim buf(0& To size - 1&) As Byte
    v = buf
    memory_mapRegister start, size, v, v

    Machine_MapRam = 0&
    Exit Function

MapFail:
    debug_log DEBUG_ERROR, "[MACHINE] ERROR: Unable to allocate " & CStr(size) & " bytes of memory"
    Machine_MapRam = -1&
End Function

Private Function Machine_MapRom(ByVal start As Long, ByVal size As Long, ByVal filename As String, ByVal required As Byte) As Long
    Dim buf() As Byte
    Dim ret As Long
    Dim vRead As Variant
    Dim vWrite As Variant

    On Error GoTo MapFail

    ReDim buf(0& To size - 1&) As Byte
    ret = utility_loadFile(buf, size, filename)
    If (required = MACHINE_ROM_REQUIRED) And (ret <> 0&) Then
        debug_log DEBUG_ERROR, "[MACHINE] Could not open file, or size is less than expected: " & filename
        Machine_MapRom = -1&
        Exit Function
    End If

    If ret <> 0& Then ReDim buf(0& To size - 1&) As Byte

    vRead = buf
    vWrite = Empty
    memory_mapRegister start, size, vRead, vWrite

    Machine_MapRom = 0&
    Exit Function

MapFail:
    debug_log DEBUG_ERROR, "[MACHINE] Could not open file, or size is less than expected: " & filename
    Machine_MapRom = -1&
End Function

Private Function Machine_GetDefIndex(ByVal machineId As String) As Long
    Select Case LCase$(machineId)
        Case "ami386": Machine_GetDefIndex = 0&
        Case "award495": Machine_GetDefIndex = 1&
        Case "seabios": Machine_GetDefIndex = 2&
        Case "award486": Machine_GetDefIndex = 3&
        Case "mrbios486": Machine_GetDefIndex = 4&
        Case "4saw2": Machine_GetDefIndex = 5&
        Case "hot543": Machine_GetDefIndex = 6&
        Case "sp97xv": Machine_GetDefIndex = 7&
        Case "p5sp4": Machine_GetDefIndex = 8&
        Case "genericxt": Machine_GetDefIndex = 9&
        Case Else: Machine_GetDefIndex = -1&
    End Select
End Function

Private Sub Machine_GetDef(ByVal idx As Long, ByRef description As String, ByRef initKind As Long, ByRef cmosfile As String, ByRef video As Byte, ByRef machineSpeed As Double, ByRef hwLo As Long, ByRef hwHi As Long)
    description = Machine_GetDefDescription(idx)
    hwLo = Machine_DefaultHwLo()
    hwHi = 0&
    video = VIDEO_CARD_VGA
    machineSpeed = -1#

    Select Case idx
        Case 0&
            initKind = MACHDEF_INIT_ASUS_386
            cmosfile = "cmos/ami386.bin"
        Case 1&
            initKind = MACHDEF_INIT_OPTI495
            cmosfile = "cmos/award495.bin"
        Case 2&
            initKind = MACHDEF_INIT_OPTI495
            cmosfile = "cmos/seabios.bin"
        Case 3&
            initKind = MACHDEF_INIT_ASUS_386
            cmosfile = "cmos/award486.bin"
        Case 4&
            initKind = MACHDEF_INIT_OPTI495
            cmosfile = "cmos/mrbios486.bin"
        Case 5&
            initKind = MACHDEF_INIT_ASUS_386
            cmosfile = "cmos/4saw2.bin"
        Case 6&
            initKind = MACHDEF_INIT_OPTI5X7
            cmosfile = "cmos/hot543.bin"
        Case 7&
            initKind = MACHDEF_INIT_OPTI5X7
            cmosfile = "cmos/sp97xv.bin"
        Case 8&
            initKind = MACHDEF_INIT_OPTI495
            cmosfile = "cmos/p5sp4.bin"
        Case 9&
            initKind = MACHDEF_INIT_GENERIC_XT
            cmosfile = vbNullString
        Case Else
            initKind = 0&
            cmosfile = vbNullString
    End Select
End Sub

Private Function Machine_GetDefId(ByVal idx As Long) As String
    Select Case idx
        Case 0&: Machine_GetDefId = "ami386"
        Case 1&: Machine_GetDefId = "award495"
        Case 2&: Machine_GetDefId = "seabios"
        Case 3&: Machine_GetDefId = "award486"
        Case 4&: Machine_GetDefId = "mrbios486"
        Case 5&: Machine_GetDefId = "4saw2"
        Case 6&: Machine_GetDefId = "hot543"
        Case 7&: Machine_GetDefId = "sp97xv"
        Case 8&: Machine_GetDefId = "p5sp4"
        Case 9&: Machine_GetDefId = "genericxt"
        Case Else: Machine_GetDefId = vbNullString
    End Select
End Function

Private Function Machine_GetDefDescription(ByVal idx As Long) As String
    Select Case idx
        Case 0&: Machine_GetDefDescription = "AMI 386"
        Case 1&: Machine_GetDefDescription = "OPTi 495 Award 486 clone"
        Case 2&: Machine_GetDefDescription = "SeaBIOS (from QEMU)"
        Case 3&: Machine_GetDefDescription = "Award 4.50G 486 clone"
        Case 4&: Machine_GetDefDescription = "OPTi 495 MR-BIOS 486"
        Case 5&: Machine_GetDefDescription = "Soyo 4SAW2"
        Case 6&: Machine_GetDefDescription = "Shuttle HOT-543"
        Case 7&: Machine_GetDefDescription = "Asus SP97-XV"
        Case 8&: Machine_GetDefDescription = "ASUS PCI/I-P5SP4"
        Case 9&: Machine_GetDefDescription = "Generic XT"
        Case Else: Machine_GetDefDescription = ""
    End Select
End Function

