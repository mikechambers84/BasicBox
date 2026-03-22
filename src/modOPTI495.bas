Attribute VB_Name = "modOPTI495"
Option Explicit

Private Type opti495_t
    devtype As Byte
    max As Byte
    idx As Byte
    regs(0& To 255&) As Byte
    scratch(0& To 1&) As Byte
End Type

Private Const OPTI493 As Byte = 0&
Private Const OPTI495_TYPE_495 As Byte = 1&
Private Const OPTI495SLC As Byte = 2&
Private Const OPTI495SX As Byte = 3&
Private Const OPTI495XLC As Byte = 4&

Private opti495_dev As opti495_t
Private opti495_masks(0& To 4&, 0& To 11&) As Byte
Private opti495_masksInited As Byte

Private Sub opti495_init_masks()
    If opti495_masksInited <> 0& Then Exit Sub

    opti495_masks(0&, 0&) = &H3F&: opti495_masks(0&, 1&) = &HFF&: opti495_masks(0&, 2&) = &HFF&: opti495_masks(0&, 3&) = &HFF&
    opti495_masks(0&, 4&) = &HF7&: opti495_masks(0&, 5&) = &HFB&: opti495_masks(0&, 6&) = &H7F&: opti495_masks(0&, 7&) = &H9F&
    opti495_masks(0&, 8&) = &HE3&: opti495_masks(0&, 9&) = &HFF&: opti495_masks(0&, 10&) = &HE3&: opti495_masks(0&, 11&) = &HFF&

    opti495_masks(1&, 0&) = &H3A&: opti495_masks(1&, 1&) = &H7F&: opti495_masks(1&, 2&) = &HFF&: opti495_masks(1&, 3&) = &HFF&
    opti495_masks(1&, 4&) = &HF0&: opti495_masks(1&, 5&) = &HFB&: opti495_masks(1&, 6&) = &H7F&: opti495_masks(1&, 7&) = &HBF&
    opti495_masks(1&, 8&) = &HE3&: opti495_masks(1&, 9&) = &HFF&: opti495_masks(1&, 10&) = &H0&: opti495_masks(1&, 11&) = &H0&

    opti495_masks(2&, 0&) = &H3A&: opti495_masks(2&, 1&) = &H7F&: opti495_masks(2&, 2&) = &HFC&: opti495_masks(2&, 3&) = &HFF&
    opti495_masks(2&, 4&) = &HF0&: opti495_masks(2&, 5&) = &HFB&: opti495_masks(2&, 6&) = &HFF&: opti495_masks(2&, 7&) = &HBF&
    opti495_masks(2&, 8&) = &HE3&: opti495_masks(2&, 9&) = &HFF&: opti495_masks(2&, 10&) = &H0&: opti495_masks(2&, 11&) = &H0&

    opti495_masks(3&, 0&) = &H3A&: opti495_masks(3&, 1&) = &HFF&: opti495_masks(3&, 2&) = &HFD&: opti495_masks(3&, 3&) = &HFF&
    opti495_masks(3&, 4&) = &HF0&: opti495_masks(3&, 5&) = &HFB&: opti495_masks(3&, 6&) = &H7F&: opti495_masks(3&, 7&) = &HBF&
    opti495_masks(3&, 8&) = &HE3&: opti495_masks(3&, 9&) = &HFF&: opti495_masks(3&, 10&) = &H0&: opti495_masks(3&, 11&) = &H0&

    opti495_masks(4&, 0&) = &H3A&: opti495_masks(4&, 1&) = &HFF&: opti495_masks(4&, 2&) = &HFC&: opti495_masks(4&, 3&) = &HFF&
    opti495_masks(4&, 4&) = &HF0&: opti495_masks(4&, 5&) = &HFB&: opti495_masks(4&, 6&) = &HFF&: opti495_masks(4&, 7&) = &HBF&
    opti495_masks(4&, 8&) = &HE3&: opti495_masks(4&, 9&) = &HFF&: opti495_masks(4&, 10&) = &H0&: opti495_masks(4&, 11&) = &H0&

    opti495_masksInited = 1&
End Sub

Private Function opti495_mask(ByVal devtype As Byte, ByVal idxOffset As Long) As Byte
    If (devtype > 4&) Or (idxOffset < 0&) Or (idxOffset > 11&) Then
        opti495_mask = &HFF&
    Else
        opti495_mask = opti495_masks(devtype, idxOffset)
    End If
End Function

Private Sub opti495_recalc(ByRef dev As opti495_t)
    Dim base As Long
    Dim shadowbios As Long
    Dim shadowbios_write As Long
    Dim shflags As Long
    Dim i As Long

    shadowbios = 0&
    shadowbios_write = 0&

    If (dev.regs(&H22&) And &H80&) <> 0& Then
        shadowbios = 1&
        shadowbios_write = 0&
        shflags = 1& Or 2&
    Else
        shadowbios = 0&
        shadowbios_write = 1&
        shflags = 4& Or 8&
    End If

    For i = 0& To 7&
        base = &HD0000& + (i * &H4000&)

        If ((dev.regs(&H22&) And IIf(base >= &HE0000&, &H20&, &H40&)) <> 0&) And ((dev.regs(&H23&) And (2& ^ i)) <> 0&) Then
            shflags = 4&
            If (dev.regs(&H22&) And IIf(base >= &HE0000&, &H8&, &H10&)) <> 0& Then
                shflags = shflags Or 8&
            Else
                shflags = shflags Or 2&
            End If
        Else
            If (dev.regs(&H26&) And &H40&) <> 0& Then
                shflags = 1&
                If (dev.regs(&H22&) And IIf(base >= &HE0000&, &H8&, &H10&)) <> 0& Then
                    shflags = shflags Or 8&
                Else
                    shflags = shflags Or 2&
                End If
            Else
                shflags = 1& Or 16&
            End If
        End If
    Next i

    For i = 0& To 3&
        base = &HC0000& + (i * &H4000&)

        If ((dev.regs(&H26&) And &H10&) <> 0&) And ((dev.regs(&H26&) And (2& ^ i)) <> 0&) Then
            shflags = 4&
            If (dev.regs(&H26&) And &H20&) <> 0& Then
                shflags = shflags Or 8&
            Else
                shflags = shflags Or 2&
            End If
        Else
            If (dev.regs(&H26&) And &H40&) <> 0& Then
                shflags = 1&
                If (dev.regs(&H26&) And &H20&) <> 0& Then
                    shflags = shflags Or 8&
                Else
                    shflags = shflags Or 2&
                End If
            Else
                shflags = 1& Or 16&
            End If
        End If
    Next i
End Sub

Public Sub opti495_write(ByVal priv As Long, ByVal addr As Integer, ByVal val As Byte)
    Select Case (addr And &HFFFF&)
        Case &H22&
            opti495_dev.idx = val

        Case &H23&, &H24&
            If (opti495_dev.idx >= &H20&) And (opti495_dev.idx <= opti495_dev.max) Then
                opti495_dev.regs(opti495_dev.idx) = (val And opti495_mask(opti495_dev.devtype, CLng(opti495_dev.idx) - &H20&))
                If (opti495_dev.devtype = OPTI493) And (opti495_dev.idx = &H20&) Then
                    val = val Or &H40&
                End If

                Select Case opti495_dev.idx
                    Case &H21&
                        ' CPU cache/waitstate hooks intentionally omitted.
                    Case &H22&, &H23&, &H26&
                        opti495_recalc opti495_dev
                End Select
            End If

            opti495_dev.idx = &HFF&

        Case &HE1&, &HE2&
            opti495_dev.scratch((Not (addr And &HFFFF&)) And 1&) = val
    End Select
End Sub

Public Function opti495_read(ByVal priv As Long, ByVal addr As Integer) As Byte
    Dim ret As Byte

    ret = &HFF&

    Select Case (addr And &HFFFF&)
        Case &H23&, &H24&
            If (opti495_dev.idx >= &H20&) And (opti495_dev.idx <= opti495_dev.max) Then
                ret = opti495_dev.regs(opti495_dev.idx)
            End If
            opti495_dev.idx = &HFF&

        Case &HE1&, &HE2&
            ret = opti495_dev.scratch((Not (addr And &HFFFF&)) And 1&)
    End Select

    opti495_read = ret
End Function

Public Sub opti495_init()
    Dim i As Long

    opti495_init_masks

    opti495_dev.devtype = OPTI495_TYPE_495
    opti495_dev.max = 0&
    opti495_dev.idx = 0&

    For i = 0& To 255&
        opti495_dev.regs(i) = 0&
    Next i

    opti495_dev.scratch(0&) = &HFF&
    opti495_dev.scratch(1&) = &HFF&

    ports_cbRegister &H22&, 3&, PORTS_CB_OPTI495, PORTS_CB_NONE, PORTS_CB_OPTI495, PORTS_CB_NONE, 0&
    ports_cbRegister &HE1&, 2&, PORTS_CB_OPTI495, PORTS_CB_NONE, PORTS_CB_OPTI495, PORTS_CB_NONE, 0&

    opti495_dev.max = &H29&
    opti495_dev.regs(&H20&) = &H2&
    opti495_dev.regs(&H21&) = &H20&
    opti495_dev.regs(&H22&) = &HE4&
    opti495_dev.regs(&H25&) = &HF0&
    opti495_dev.regs(&H26&) = &H80&
    opti495_dev.regs(&H27&) = &HB1&
    opti495_dev.regs(&H28&) = &H80&
    opti495_dev.regs(&H29&) = &H10&

    opti495_recalc opti495_dev

    ports_cbRegister &HE1&, 2&, PORTS_CB_OPTI495, PORTS_CB_NONE, PORTS_CB_OPTI495, PORTS_CB_NONE, 0&
End Sub

