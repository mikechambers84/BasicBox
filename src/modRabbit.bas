Attribute VB_Name = "modRabbit"
Option Explicit

Private Type rabbit_t
    cur_reg As Byte
    tries As Byte
    regs(0& To 257&) As Byte
End Type

Private rabbit As rabbit_t

Private Sub rabbit_recalcmapping(ByRef dev As rabbit_t)
    Dim shread As Long
    Dim shwrite As Long
    Dim shflags As Long

    shread = Abs((dev.regs(&H101&) And &H40&) <> 0&)
    shwrite = Abs((dev.regs(&H100&) And &H2&) <> 0&)
    shflags = 0&

    Select Case (dev.regs(&H100&) And &H9&)
        Case &H1&
            ' Shadow mapping hooks intentionally omitted (same as current C source path).
        Case &H0&
            ' Shadow mapping hooks intentionally omitted.
        Case &H9&
            ' Shadow mapping hooks intentionally omitted.
        Case &H8&
            ' Shadow mapping hooks intentionally omitted.
        Case Else
            ' No-op.
    End Select
End Sub

Public Sub rabbit_write(ByVal priv As Long, ByVal addr As Integer, ByVal val As Byte)
    Select Case (addr And &HFFFF&)
        Case &H22&
            rabbit.cur_reg = val
            rabbit.tries = 0&

        Case &H23&
            If rabbit.cur_reg = &H83& Then
                If rabbit.tries < 2& Then
                    rabbit.regs((rabbit.tries Or &H100&)) = val
                    rabbit.tries = CByte(rabbit.tries + 1&)
                    If rabbit.tries = 2& Then
                        rabbit_recalcmapping rabbit
                    End If
                End If
            End If

            rabbit.regs(rabbit.cur_reg) = val
    End Select
End Sub

Public Function rabbit_read(ByVal priv As Long, ByVal addr As Integer) As Byte
    Dim ret As Byte

    ret = &HFF&

    Select Case (addr And &HFFFF&)
        Case &H23&
            If rabbit.cur_reg = &H83& Then
                If rabbit.tries < 2& Then
                    ret = rabbit.regs((rabbit.tries Or &H100&))
                    rabbit.tries = CByte(rabbit.tries + 1&)
                End If
            Else
                ret = rabbit.regs(rabbit.cur_reg)
            End If
    End Select

    rabbit_read = ret
End Function

Public Sub rabbit_init()
    debug_log DEBUG_INFO, "[SIS 85C310] Init SiS 310 ""Rabbit"" chipset"
    ports_cbRegister &H22&, 2&, PORTS_CB_RABBIT, PORTS_CB_NONE, PORTS_CB_RABBIT, PORTS_CB_NONE, 0&
End Sub
