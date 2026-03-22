Attribute VB_Name = "modOPTI5X7"
Option Explicit

Private Type opti5x7_t
    idx As Byte
    is_pci As Byte
    regs(0& To 17&) As Byte
End Type

Private opti5x7 As opti5x7_t

Private Sub opti5x7_shadow_map(ByVal cur_reg As Byte)
    Dim i As Long

    If cur_reg = &H6& Then
        If opti5x7.is_pci <> 0& Then
            ' Shadow mapping hooks intentionally omitted.
        Else
            ' Shadow mapping hooks intentionally omitted.
        End If
    Else
        For i = 0& To 3&
            ' Shadow mapping hooks intentionally omitted.
        Next i
    End If
End Sub

Public Sub opti5x7_write(ByVal priv As Long, ByVal addr As Integer, ByVal val As Byte)
    Select Case (addr And &HFFFF&)
        Case &H22&
            opti5x7.idx = val

        Case &H24&
            Select Case opti5x7.idx
                Case &H0&
                    opti5x7.regs(opti5x7.idx) = (val And &H7F&)

                Case &H1&, &H2&, &H3&, &H7&, &H8&, &H9&, &HA&, &HB&, &HD&, &HE&, &HF&, &H10&, &H11&
                    opti5x7.regs(opti5x7.idx) = val

                Case &H4&, &H5&, &H6&
                    opti5x7.regs(opti5x7.idx) = val
                    opti5x7_shadow_map opti5x7.idx

                Case &HC&
                    opti5x7.regs(opti5x7.idx) = (val And &HCF&)

                Case Else
                    ' No-op.
            End Select
    End Select
End Sub

Public Function opti5x7_read(ByVal priv As Long, ByVal addr As Integer) As Byte
    If ((addr And &HFFFF&) = &H24&) And (opti5x7.idx < 18&) Then
        opti5x7_read = opti5x7.regs(opti5x7.idx)
    Else
        opti5x7_read = &HFF&
    End If
End Function

Public Sub opti5x7_init()
    Dim i As Long

    opti5x7.idx = 0&
    opti5x7.is_pci = 0&
    For i = 0& To 17&
        opti5x7.regs(i) = 0&
    Next i

    ports_cbRegister &H22&, 1&, PORTS_CB_OPTI5X7, PORTS_CB_NONE, PORTS_CB_OPTI5X7, PORTS_CB_NONE, 0&
    ports_cbRegister &H24&, 1&, PORTS_CB_OPTI5X7, PORTS_CB_NONE, PORTS_CB_OPTI5X7, PORTS_CB_NONE, 0&
End Sub
