Attribute VB_Name = "modI8259"
Option Explicit

Private Const I8259_MAX As Long = 8&

Public i8259_devices(0& To I8259_MAX - 1&) As I8259_t 'public to check it from cpu_exec
Private i8259_used(0& To I8259_MAX - 1&) As Byte

Public Sub i8259_init(ByRef i8259 As Long, ByVal portbase As Long, ByVal master As Long)
    Dim idx As Long
    Dim zeroPic As I8259_t

    idx = i8259
    If (idx < 0&) Or (idx >= I8259_MAX) Or (i8259_used(idx) = 0&) Then
        idx = I8259_AllocSlot()
        If idx < 0& Then
            debug_log DEBUG_ERROR, "[I8259] Out of PIC slots"
            End
        End If
    End If

    i8259_devices(idx) = zeroPic
    i8259_used(idx) = 1&
    i8259_devices(idx).intoffset = 0&
    i8259_devices(idx).IMR = &HFF&
    i8259_devices(idx).master = master

    ports_cbRegister portbase, 2&, PORTS_CB_I8259, PORTS_CB_NONE, PORTS_CB_I8259, PORTS_CB_NONE, idx

    i8259 = idx
End Sub

Public Sub i8259_doirq(ByVal i8259 As Long, ByVal irqnum As Byte)
    Dim irr As Long
    Dim bitMask As Long

    If Not I8259_IsValid(i8259) Then Exit Sub

    irr = i8259_devices(i8259).irr
    bitMask = (2& ^ (irqnum And &H7&))
    irr = irr Or bitMask

    i8259_devices(i8259).irr = CByte(irr And &HFF&)
End Sub

Public Sub i8259_clearirq(ByVal i8259 As Long, ByVal irqnum As Byte)
    Dim irr As Long
    Dim bitMask As Long

    If Not I8259_IsValid(i8259) Then Exit Sub

    irr = i8259_devices(i8259).irr
    bitMask = (2& ^ (irqnum And &H7&))
    irr = irr And ((Not bitMask) And &HFF&)

    i8259_devices(i8259).irr = CByte(irr And &HFF&)
End Sub

Public Sub i8259_setlevelirq(ByVal i8259 As Long, ByVal irqnum As Byte, ByVal state As Byte)
    Dim irr As Long
    Dim bitMask As Long
    Dim irqLine As Long

    If Not I8259_IsValid(i8259) Then Exit Sub

    irqLine = (irqnum And &H7&)
    bitMask = (2& ^ irqLine)
    irr = i8259_devices(i8259).irr

    If state <> 0& Then
        i8259_devices(i8259).lineActive(irqLine) = 1&
        irr = irr Or bitMask
    Else
        i8259_devices(i8259).lineActive(irqLine) = 0&
        irr = irr And ((Not bitMask) And &HFF&)
    End If

    i8259_devices(i8259).irr = CByte(irr And &HFF&)
End Sub

Public Function i8259_nextintr(ByVal i8259 As Long) As Byte
    Dim offset As Long
    Dim irq As Long
    Dim tmpirr As Long
    Dim currentIsr As Long
    Dim bitMask As Long
    Dim startPriority As Long

    If Not I8259_IsValid(i8259) Then
        i8259_nextintr = &HFF&
        Exit Function
    End If

    tmpirr = i8259_devices(i8259).irr And ((Not i8259_devices(i8259).IMR) And &HFF&)
    currentIsr = i8259_devices(i8259).ISR
    startPriority = (i8259_devices(i8259).priority And &H7&)

    For offset = 0& To 7&
        irq = ((offset + startPriority) And &H7&)
        bitMask = (2& ^ irq)

        If (currentIsr And bitMask) <> 0& Then Exit For

        If (tmpirr And bitMask) <> 0& Then
            If i8259_devices(i8259).lineActive(irq) = 0& Then
                i8259_devices(i8259).irr = CByte(i8259_devices(i8259).irr And ((Not bitMask) And &HFF&))
            End If
            i8259_devices(i8259).ISR = CByte((i8259_devices(i8259).ISR Or bitMask) And &HFF&)
            i8259_devices(i8259).lastintr = CByte(irq And &HFF&)
            i8259_nextintr = CByte(irq And &HFF&)
            Exit Function
        End If
    Next offset

    i8259_nextintr = &HFF&
End Function

Public Sub i8259_write(ByVal i8259 As Long, ByVal portnum As Integer, ByVal value As Byte)
    Dim idx As Long
    Dim current As Long
    Dim bitMask As Long

    If Not I8259_IsValid(i8259) Then Exit Sub

    idx = i8259

    Select Case (portnum And 1&)
        Case 0&
            If (value And &H10&) <> 0& Then
                i8259_devices(idx).IMR = &H0&
                i8259_devices(idx).icw(1&) = value
                i8259_devices(idx).icwstep = 2&
                i8259_devices(idx).readmode = 0&
            ElseIf (value And &H8&) = 0& Then
                i8259_devices(idx).ocw(2&) = value
                Select Case (value And &HE0&)
                    Case &H60&
                        i8259_devices(idx).irr = CByte(i8259_devices(idx).irr And ((Not (2& ^ (value And &H7&))) And &HFF&))
                        i8259_devices(idx).ISR = CByte(i8259_devices(idx).ISR And ((Not (2& ^ (value And &H7&))) And &HFF&))
                    Case &H40&
                        ' No operation.
                    Case &H20&
                        For current = 0& To 7&
                            bitMask = (2& ^ ((current + i8259_devices(idx).priority) And &H7&))
                            If (i8259_devices(idx).ISR And bitMask) <> 0& Then
                                i8259_devices(idx).ISR = CByte(i8259_devices(idx).ISR And ((Not bitMask) And &HFF&))
                                Exit For
                            End If
                        Next current
                    Case Else
                        ' Unhandled EOI type.
                End Select
            Else
                i8259_devices(idx).ocw(3&) = value
                If (value And &H2&) <> 0& Then
                    i8259_devices(idx).readmode = value And &H1&
                End If
            End If

        Case 1&
            Select Case i8259_devices(idx).icwstep
                Case 2&
                    i8259_devices(idx).icw(2&) = value
                    i8259_devices(idx).intoffset = value And &HF8&
                    If (i8259_devices(idx).icw(1&) And &H2&) <> 0& Then
                        i8259_devices(idx).icwstep = 4&
                    Else
                        i8259_devices(idx).icwstep = 3&
                    End If

                Case 3&
                    i8259_devices(idx).icw(3&) = value
                    If (i8259_devices(idx).icw(1&) And &H1&) <> 0& Then
                        i8259_devices(idx).icwstep = 4&
                    Else
                        i8259_devices(idx).icwstep = 5&
                    End If

                Case 4&
                    i8259_devices(idx).icw(4&) = value
                    i8259_devices(idx).icwstep = 5&

                Case 5&
                    i8259_devices(idx).IMR = value
            End Select
    End Select
End Sub

Public Function i8259_read(ByVal i8259 As Long, ByVal portnum As Integer) As Byte
    Dim idx As Long

    If Not I8259_IsValid(i8259) Then
        i8259_read = 0&
        Exit Function
    End If

    idx = i8259

    Select Case (portnum And 1&)
        Case 0&
            If i8259_devices(idx).readmode = 0& Then
                i8259_read = i8259_devices(idx).irr
            Else
                i8259_read = i8259_devices(idx).ISR
            End If
        Case 1&
            i8259_read = i8259_devices(idx).IMR
        Case Else
            i8259_read = 0&
    End Select
End Function

Public Function i8259_irqInService(ByVal i8259 As Long, ByVal irqnum As Byte) As Byte
    Dim bitMask As Long

    If Not I8259_IsValid(i8259) Then
        i8259_irqInService = 0&
        Exit Function
    End If

    bitMask = (2& ^ (irqnum And &H7&))
    If (i8259_devices(i8259).ISR And bitMask) <> 0& Then
        i8259_irqInService = 1&
    Else
        i8259_irqInService = 0&
    End If
End Function

Public Function i8259_getIntOffset(ByVal i8259 As Long) As Byte
    If Not I8259_IsValid(i8259) Then
        i8259_getIntOffset = 0&
        Exit Function
    End If

    i8259_getIntOffset = i8259_devices(i8259).intoffset
End Function
Private Function I8259_AllocSlot() As Long
    Dim i As Long

    For i = 0& To I8259_MAX - 1&
        If i8259_used(i) = 0& Then
            i8259_used(i) = 1&
            I8259_AllocSlot = i
            Exit Function
        End If
    Next i

    I8259_AllocSlot = -1&
End Function

Private Function I8259_IsValid(ByVal idx As Long) As Boolean
    I8259_IsValid = ((idx >= 0&) And (idx < I8259_MAX) And (i8259_used(idx) <> 0&))
End Function

