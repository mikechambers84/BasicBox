Attribute VB_Name = "modTypes"
Option Explicit

Public Const CPU_TLB_ENTRY_VALID As Long = &H1&
Public Const CPU_TLB_ENTRY_USER_OK As Long = &H2&
Public Const CPU_TLB_ENTRY_WRITE_OK As Long = &H4&
Public Const CPU_TLB_ENTRY_DIRTY As Long = &H8&

Public Type CPU_TLB_ENTRY_t
    tag As Long
    phys_base As Long
    pte_addr As Long
    flags As Byte
End Type

Public Type CPU_TLB_SET_t
    way(0& To 1&) As CPU_TLB_ENTRY_t
    mru As Byte
End Type

Public Type CPU_t
    regs_long(0& To 7&) As Long

    opcode As Byte
    segoverride As Byte
    reptype As Byte

    a20_gate As Byte
    nowrite As Byte
    currentseg As Byte
    startcpl As Byte
    cpl As Byte
    doexception As Byte
    exceptionval As Byte
    hltstate As Byte
    isaddr32 As Byte
    ifl As Byte
    isCS32 As Byte
    protected_mode As Byte
    paging As Byte
    usegdt As Byte
    tr As Byte
    have387 As Byte
    v86f As Byte

    sib As Byte
    sib_scale As Byte
    sib_index As Byte
    sib_base As Byte

    tf As Byte
    tempcf As Byte
    oldcf As Byte
    cf As Byte
    pf As Byte
    af As Byte
    zf As Byte
    sf As Byte
    df As Byte
    ofl As Byte
    rf As Byte
    acf As Byte
    idf As Byte
    mode As Byte
    reg As Byte
    rm As Byte
    nt As Byte
    iopl As Byte
    isoper32 As Byte

    trap_toggle As Long
    totalexec_lo As Long
    totalexec_hi As Long
    ip As Long
    saveip As Long
    savecs As Long
    exceptionip As Long
    useseg As Long
    oldsp As Long

    segregs(0& To 5&) As Long
    segcache(0& To 5&) As Long
    segis32(0& To 5&) As Byte
    seglimit(0& To 5&) As Long

    oper1 As Long
    oper2 As Long
    res16 As Long
    disp16 As Long
    temp16 As Long
    dummy As Long
    stacksize As Long
    frametemp As Long

    oper1_32 As Long
    oper2_32 As Long
    res32 As Long
    disp32 As Long

    oper1b As Long
    oper2b As Long
    res8 As Long
    disp8 As Long
    temp8 As Long
    nestlev As Long
    addrbyte As Long

    sib_val As Long
    temp1 As Long
    temp2 As Long
    temp3 As Long
    temp4 As Long
    temp5 As Long
    temp32 As Long
    tempaddr32 As Long
    frametemp32 As Long
    ea As Long

    exceptionerr As Long
    cr(0& To 7&) As Long
    dr(0& To 7&) As Long

    gdtr As Long
    gdtl As Long
    idtr As Long
    idtl As Long
    ldtr As Long
    ldtl As Long
    trbase As Long
    trlimit As Long
    shadow_esp As Long
    trtype As Byte
    bypass_paging As Byte
    interrupt_inhibit As Byte
    ldt_selector As Long
    tr_selector As Long
    result As Long

    int_callback(0& To 255&) As Long
    tlb(0& To 255&) As CPU_TLB_SET_t
End Type

Public Type I8259_t
    imr As Byte
    irr As Byte
    isr As Byte
    lineActive(0& To 7&) As Byte
    icwstep As Byte
    icw(0& To 4&) As Byte
    ocw(0& To 4&) As Byte
    intoffset As Byte
    priority As Byte
    autoeoi As Byte
    readmode As Byte
    vector As Byte
    lastintr As Byte
    enabled As Byte
    master As Long
End Type

Public Type KEYSTATE_t
    scancode As Byte
    isNew As Byte
End Type

Public Const BUSLOGIC_MAX_TARGETS As Long = 7&
Public Const BUSLOGIC_TARGET_NONE As Long = 0&
Public Const BUSLOGIC_TARGET_DISK As Long = 1&
Public Const BUSLOGIC_TARGET_CDROM As Long = 2&

Public Type BUSLOGIC_TARGET_t
    present As Byte
    targetType As Byte
    path As String * 512
End Type
Public Type MACHINE_t
    CPU As CPU_t
    i8259 As Long
    i8259_slave As Long
    KeyState As KEYSTATE_t
    mixOPL As Byte
    mixBlaster As Byte
    pcap_if As Long
    hwflags_lo As Long
    hwflags_hi As Long
    extmem As Long
    buslogic_enabled As Byte
    buslogic_base As Long
    buslogic_irq As Byte
    buslogic_dma As Byte
    buslogic_bios_addr As Long
    buslogic_rom_path As String * 512
    buslogic_nvr_path As String * 512
    scsi_targets(0& To BUSLOGIC_MAX_TARGETS - 1&) As BUSLOGIC_TARGET_t
End Type

