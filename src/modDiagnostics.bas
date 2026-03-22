Attribute VB_Name = "modDiagnostics"
Option Explicit

Private Const U64_MOD_D As Double = 1.84467440737096E+19
Private Const U32_MOD_D As Double = 4294967296#

Public diag_enabled As Byte
Public diag_vga_verbose As Byte

Private diag_tick_div As Long
Private diag_prev_exec As Double
Private diag_lastIPS As String

Private diag_sec_illegal As Long
Private diag_sec_exception As Long
Private diag_sec_exception1 As Long
Private diag_sec_exception6 As Long
Private diag_sec_exception7 As Long
Private diag_sec_exception13 As Long
Private diag_sec_exception14 As Long
Private diag_sec_exception_by_num(0& To 255&) As Long
Private diag_sec_irq0 As Long
Private diag_sec_cga_port As Long
Private diag_sec_cga_mem As Long
Private diag_sec_vga_port As Long
Private diag_sec_vga_mem As Long

Private diag_sec_de_opcode_counts(0& To 255&) As Long
Private diag_sec_de_last_opcode As Long
Private diag_sec_de_last_next As Long
Private diag_sec_de_last_cs As Long
Private diag_sec_de_last_ip As Long
Private diag_sec_de_last_is32 As Long
Private diag_sec_de_last_dividend_lo As Long
Private diag_sec_de_last_dividend_hi As Long
Private diag_sec_de_last_divisor As Long
Private diag_sec_de_last_reason As Long

Private diag_total_illegal As Double
Private diag_total_exception As Double
Private diag_total_exception1 As Double
Private diag_total_exception6 As Double
Private diag_total_exception7 As Double
Private diag_total_exception13 As Double
Private diag_total_exception14 As Double
Private diag_total_irq0 As Double
Private diag_total_cga_port As Double
Private diag_total_cga_mem As Double
Private diag_total_vga_port As Double
Private diag_total_vga_mem As Double

Public diag_cga_inited As Byte
Public diag_vga_inited As Byte
Public diag_et4000_inited As Byte

Public Sub diag_init(ByRef machineRef As MACHINE_t)
    Dim curExec As Double
    Dim i As Long

    diag_enabled = 1&
    diag_vga_verbose = 0&
    diag_tick_div = 0&
    diag_lastIPS = vbNullString

    diag_sec_illegal = 0&
    diag_sec_exception = 0&
    diag_sec_exception1 = 0&
    diag_sec_exception6 = 0&
    diag_sec_exception7 = 0&
    diag_sec_exception13 = 0&
    diag_sec_exception14 = 0&
    diag_sec_irq0 = 0&
    diag_sec_cga_port = 0&
    diag_sec_cga_mem = 0&
    diag_sec_vga_port = 0&
    diag_sec_vga_mem = 0&

    diag_total_illegal = 0#
    diag_total_exception = 0#
    diag_total_exception1 = 0#
    diag_total_exception6 = 0#
    diag_total_exception7 = 0#
    diag_total_exception13 = 0#
    diag_total_exception14 = 0#
    diag_total_irq0 = 0#
    diag_total_cga_port = 0#
    diag_total_cga_mem = 0#
    diag_total_vga_port = 0#
    diag_total_vga_mem = 0#

    For i = 0& To 255&
        diag_sec_exception_by_num(i) = 0&
        diag_sec_de_opcode_counts(i) = 0&
    Next i

    diag_sec_de_last_opcode = 0&
    diag_sec_de_last_next = 0&
    diag_sec_de_last_cs = 0&
    diag_sec_de_last_ip = 0&
    diag_sec_de_last_is32 = 0&
    diag_sec_de_last_dividend_lo = 0&
    diag_sec_de_last_dividend_hi = 0&
    diag_sec_de_last_divisor = 0&
    diag_sec_de_last_reason = 0&

    diag_cga_inited = 0&
    diag_vga_inited = 0&
    diag_et4000_inited = 0&

    curExec = U32ToDouble(machineRef.cpu.totalexec_hi) * U32_MOD_D
    curExec = curExec + U32ToDouble(machineRef.cpu.totalexec_lo)
    diag_prev_exec = curExec
End Sub

Public Sub diag_tick(ByRef machineRef As MACHINE_t)
    Dim curExec As Double
    Dim deltaExec As Double
    Dim msg As String
    Dim csip As String
    Dim i As Long
    Dim topEx As Long
    Dim topExCount As Long
    Dim topDeOpcode As Long
    Dim topDeCount As Long
    Dim deInfo As String
    Dim vgaInfo As String

    If diag_enabled = 0& Then Exit Sub

    diag_tick_div = diag_tick_div + 1&
    If diag_tick_div < 10& Then Exit Sub
    diag_tick_div = 0&

    curExec = U32ToDouble(machineRef.cpu.totalexec_hi) * U32_MOD_D
    curExec = curExec + U32ToDouble(machineRef.cpu.totalexec_lo)
    deltaExec = curExec - diag_prev_exec
    If deltaExec < 0# Then deltaExec = deltaExec + U64_MOD_D
    diag_prev_exec = curExec

    csip = right$("0000" & Hex$(machineRef.cpu.segregs(1&)), 4&) & ":" & right$("00000000" & Hex$(machineRef.cpu.ip), 8&)

    topEx = 0&
    topExCount = 0&
    topDeOpcode = 0&
    topDeCount = 0&

    For i = 0& To 255&
        If diag_sec_exception_by_num(i) > topExCount Then
            topExCount = diag_sec_exception_by_num(i)
            topEx = i
        End If

        If diag_sec_de_opcode_counts(i) > topDeCount Then
            topDeCount = diag_sec_de_opcode_counts(i)
            topDeOpcode = i
        End If
    Next i

    deInfo = ""
    If topDeCount > 0& Then
        deInfo = " deOp=" & right$("00" & Hex$(topDeOpcode), 2&) & ":" & CStr(topDeCount) & _
                 " deLast=" & right$("0000" & Hex$(diag_sec_de_last_cs), 4&) & ":" & right$("00000000" & Hex$(diag_sec_de_last_ip), 8&) & _
                 " " & right$("00" & Hex$(diag_sec_de_last_opcode), 2&) & " " & right$("00" & Hex$(diag_sec_de_last_next), 2&) & _
                 " deMath=" & IIf(diag_sec_de_last_is32 <> 0&, "32", "16") & _
                 " d=" & right$("00000000" & Hex$(diag_sec_de_last_dividend_hi), 8&) & ":" & right$("00000000" & Hex$(diag_sec_de_last_dividend_lo), 8&) & _
                 " /" & right$("00000000" & Hex$(diag_sec_de_last_divisor), 8&) & _
                 " r=" & CStr(diag_sec_de_last_reason)
    End If

    vgaInfo = ""
    If (diag_vga_verbose <> 0&) And (diag_et4000_inited <> 0&) Then
        vgaInfo = et4000_diagSnapshotAndReset()
    ElseIf (diag_vga_verbose <> 0&) And (diag_vga_inited <> 0&) Then
        vgaInfo = vga_diagSnapshotAndReset()
    End If

    diag_lastIPS = CStr(Fix(deltaExec))

    'msg = "[DIAG] ips=" & diag_lastIPS & _
          " ex=" & CStr(diag_sec_exception) & _
          " (#1=" & CStr(diag_sec_exception1) & ",#6=" & CStr(diag_sec_exception6) & ",#7=" & CStr(diag_sec_exception7) & ",#13=" & CStr(diag_sec_exception13) & ",#14=" & CStr(diag_sec_exception14) & ",top=" & CStr(topEx) & ":" & CStr(topExCount) & ")" & _
          " ill=" & CStr(diag_sec_illegal) & _
          " irq0=" & CStr(diag_sec_irq0) & _
          " tf=" & CStr(machineRef.cpu.tf) & " tt=" & CStr(machineRef.cpu.trap_toggle) & " pm=" & CStr(machineRef.cpu.protected_mode) & _
          " cs:ip=" & csip & deInfo & _
          " cgaInit=" & CStr(diag_cga_inited) & " cgaP=" & CStr(diag_sec_cga_port) & " cgaM=" & CStr(diag_sec_cga_mem) & _
          " vgaInit=" & CStr(diag_vga_inited) & " vgaP=" & CStr(diag_sec_vga_port) & " vgaM=" & CStr(diag_sec_vga_mem) & vgaInfo

    'debug_log DEBUG_INFO, msg

    diag_sec_illegal = 0&
    diag_sec_exception = 0&
    diag_sec_exception1 = 0&
    diag_sec_exception6 = 0&
    diag_sec_exception7 = 0&
    diag_sec_exception13 = 0&
    diag_sec_exception14 = 0&
    diag_sec_irq0 = 0&
    diag_sec_cga_port = 0&
    diag_sec_cga_mem = 0&
    diag_sec_vga_port = 0&
    diag_sec_vga_mem = 0&

    For i = 0& To 255&
        diag_sec_exception_by_num(i) = 0&
        diag_sec_de_opcode_counts(i) = 0&
    Next i

    diag_sec_de_last_opcode = 0&
    diag_sec_de_last_next = 0&
    diag_sec_de_last_cs = 0&
    diag_sec_de_last_ip = 0&
    diag_sec_de_last_is32 = 0&
    diag_sec_de_last_dividend_lo = 0&
    diag_sec_de_last_dividend_hi = 0&
    diag_sec_de_last_divisor = 0&
    diag_sec_de_last_reason = 0&
End Sub

Public Function diag_getIPSString() As String
    diag_getIPSString = diag_lastIPS
End Function

Public Sub diag_mark_cga_init()
    diag_cga_inited = 1&
End Sub

Public Sub diag_mark_vga_init()
    diag_vga_inited = 1&
End Sub

Public Sub diag_mark_et4000_init()
    diag_et4000_inited = 1&
End Sub

Public Sub diag_set_vga_verbose(ByVal enabled As Long)
    If enabled <> 0& Then
        diag_vga_verbose = 1&
    Else
        diag_vga_verbose = 0&
    End If
End Sub

Public Sub diag_count_illegal()
    diag_sec_illegal = diag_sec_illegal + 1&
    diag_total_illegal = diag_total_illegal + 1#
End Sub

Public Sub diag_note_divide_fault(ByVal opcode As Long, ByVal nextByte As Long, ByVal cs As Long, ByVal ip As Long)
    Dim idx As Long

    idx = (opcode And &HFF&)
    diag_sec_de_opcode_counts(idx) = diag_sec_de_opcode_counts(idx) + 1&

    diag_sec_de_last_opcode = idx
    diag_sec_de_last_next = (nextByte And &HFF&)
    diag_sec_de_last_cs = (cs And &HFFFF&)
    diag_sec_de_last_ip = ip
End Sub

Public Sub diag_note_divide_math(ByVal is32 As Long, ByVal dividendLo As Long, ByVal dividendHi As Long, ByVal divisor As Long, ByVal reason As Long)
    diag_sec_de_last_is32 = (is32 And &H1&)
    diag_sec_de_last_dividend_lo = dividendLo
    diag_sec_de_last_dividend_hi = dividendHi
    diag_sec_de_last_divisor = divisor
    diag_sec_de_last_reason = reason
End Sub

Public Sub diag_count_exception(ByVal exnum As Long)
    Dim idx As Long

    idx = (exnum And &HFF&)

    diag_sec_exception = diag_sec_exception + 1&
    diag_total_exception = diag_total_exception + 1#
    diag_sec_exception_by_num(idx) = diag_sec_exception_by_num(idx) + 1&

    Select Case idx
        Case 1&
            diag_sec_exception1 = diag_sec_exception1 + 1&
            diag_total_exception1 = diag_total_exception1 + 1#
        Case 6&
            diag_sec_exception6 = diag_sec_exception6 + 1&
            diag_total_exception6 = diag_total_exception6 + 1#
        Case 7&
            diag_sec_exception7 = diag_sec_exception7 + 1&
            diag_total_exception7 = diag_total_exception7 + 1#
        Case 13&
            diag_sec_exception13 = diag_sec_exception13 + 1&
            diag_total_exception13 = diag_total_exception13 + 1#
        Case 14&
            diag_sec_exception14 = diag_sec_exception14 + 1&
            diag_total_exception14 = diag_total_exception14 + 1#
    End Select
End Sub

Public Sub diag_count_irq0()
    diag_sec_irq0 = diag_sec_irq0 + 1&
    diag_total_irq0 = diag_total_irq0 + 1#
End Sub

Public Sub diag_count_cga_port()
    diag_sec_cga_port = diag_sec_cga_port + 1&
    diag_total_cga_port = diag_total_cga_port + 1#
End Sub

Public Sub diag_count_cga_mem()
    diag_sec_cga_mem = diag_sec_cga_mem + 1&
    diag_total_cga_mem = diag_total_cga_mem + 1#
End Sub

Public Sub diag_count_vga_port()
    diag_sec_vga_port = diag_sec_vga_port + 1&
    diag_total_vga_port = diag_total_vga_port + 1#
End Sub

Public Sub diag_count_vga_mem()
    diag_sec_vga_mem = diag_sec_vga_mem + 1&
    diag_total_vga_mem = diag_total_vga_mem + 1#
End Sub
