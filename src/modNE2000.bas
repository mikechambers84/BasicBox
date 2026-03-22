Attribute VB_Name = "modNE2000"
Option Explicit

Private Const NE2K_MEMSIZ As Long = (32& * 1024&)
Private Const NE2K_MEMSTART As Long = (16& * 1024&)
Private Const NE2K_MEMEND As Long = (NE2K_MEMSTART + NE2K_MEMSIZ)

Private Const NE2K_RESET_HARDWARE As Long = 0&
Private Const NE2K_RESET_SOFTWARE As Long = 1&
Private Const NE2K_NEVER_FULL_RING As Long = 1&

Private Const NE2K_MAX As Long = 2&

Public Const NE2000_RX_RETRY As Long = 0&
Public Const NE2000_RX_ACCEPTED As Long = 1&
Public Const NE2000_RX_DROP As Long = -1&

Private Type NE2K_CR_t
    stopBit As Byte
    startBit As Byte
    tx_packet As Byte
    rdma_cmd As Byte
    pgsel As Byte
End Type

Private Type NE2K_ISR_t
    pkt_rx As Byte
    pkt_tx As Byte
    rx_err As Byte
    tx_err As Byte
    overwrite As Byte
    cnt_oflow As Byte
    rdma_done As Byte
    resetBit As Byte
End Type

Private Type NE2K_IMR_t
    rx_inte As Byte
    tx_inte As Byte
    rxerr_inte As Byte
    txerr_inte As Byte
    overw_inte As Byte
    cofl_inte As Byte
    rdma_inte As Byte
End Type

Private Type NE2K_DCR_t
    wdsize As Byte
    endian As Byte
    longaddr As Byte
    loopBit As Byte
    auto_rx As Byte
    fifo_size As Byte
End Type

Private Type NE2K_TCR_t
    crc_disable As Byte
    loop_cntl As Byte
    ext_stoptx As Byte
    coll_prio As Byte
End Type

Private Type NE2K_TSR_t
    tx_ok As Byte
    collided As Byte
    aborted As Byte
    no_carrier As Byte
    fifo_ur As Byte
    cd_hbeat As Byte
    ow_coll As Byte
End Type

Private Type NE2K_RCR_t
    errors_ok As Byte
    runts_ok As Byte
    broadcast As Byte
    multicast As Byte
    promisc As Byte
    monitor As Byte
End Type

Private Type NE2K_RSR_t
    rx_ok As Byte
    bad_crc As Byte
    bad_falign As Byte
    fifo_or As Byte
    rx_missed As Byte
    rx_mbit As Byte
    rx_disabled As Byte
    deferred As Byte
End Type

Private Type NE2000_t
    CR As NE2K_CR_t
    ISR As NE2K_ISR_t
    IMR As NE2K_IMR_t
    DCR As NE2K_DCR_t
    TCR As NE2K_TCR_t
    TSR As NE2K_TSR_t
    RCR As NE2K_RCR_t
    RSR As NE2K_RSR_t

    local_dma As Long
    page_start As Byte
    page_stop As Byte
    bound_ptr As Byte
    tx_page_start As Byte
    num_coll As Byte
    tx_bytes As Long
    fifo As Byte
    remote_dma As Long
    remote_start As Long
    remote_bytes As Long
    tallycnt_0 As Byte
    tallycnt_1 As Byte
    tallycnt_2 As Byte

    i8029id0 As Byte
    i8029id1 As Byte

    physaddr(0& To 5&) As Byte
    curr_page As Byte
    mchash(0& To 7&) As Byte

    rempkt_ptr As Byte
    localpkt_ptr As Byte
    address_cnt As Long

    crRaw As Byte
    i9346cr As Byte
    config0 As Byte
    config2 As Byte
    config3 As Byte
    hltclk As Byte
    i8029asid0 As Byte
    i8029asid1 As Byte

    macaddr(0& To 31&) As Byte
    mem(0& To NE2K_MEMSIZ - 1&) As Byte

    base_address As Long
    base_irq As Long
    tx_timer_active As Byte
    i8259Slot As Long
    tx_timer As Long
    used As Byte
End Type

Private ne2k_devs(0& To NE2K_MAX - 1&) As NE2000_t
Private ne2k_primary As Long

Private Function NE2K_IsValid(ByVal devId As Long) As Boolean
    NE2K_IsValid = ((devId >= 0&) And (devId < NE2K_MAX) And (ne2k_devs(devId).used <> 0&))
End Function

Public Function ne2000_getPrimary() As Long
    ne2000_getPrimary = ne2k_primary
End Function

Private Function NE2K_Flag(ByVal cond As Boolean) As Byte
    If cond Then
        NE2K_Flag = 1&
    Else
        NE2K_Flag = 0&
    End If
End Function

Private Function NE2K_PktByte(ByRef pkt() As Byte, ByVal idx As Long) As Byte
    If idx < LBound(pkt) Or idx > UBound(pkt) Then
        NE2K_PktByte = 0&
    Else
        NE2K_PktByte = pkt(idx)
    End If
End Function

Private Function NE2K_MemByte(ByVal devId As Long, ByVal idx As Long) As Byte
    If (idx < 0&) Or (idx >= NE2K_MEMSIZ) Then
        NE2K_MemByte = 0&
    Else
        NE2K_MemByte = ne2k_devs(devId).mem(idx)
    End If
End Function

Private Sub NE2K_MemSetByte(ByVal devId As Long, ByVal idx As Long, ByVal value As Byte)
    If (idx < 0&) Or (idx >= NE2K_MEMSIZ) Then Exit Sub
    ne2k_devs(devId).mem(idx) = value
End Sub

Private Function NE2K_MemWordLE(ByVal devId As Long, ByVal idx As Long) As Long
    NE2K_MemWordLE = (CLng(NE2K_MemByte(devId, idx)) Or (CLng(NE2K_MemByte(devId, idx + 1&)) * &H100&)) And &HFFFF&
End Function

Private Sub NE2K_MemSetWordLE(ByVal devId As Long, ByVal idx As Long, ByVal value As Long)
    NE2K_MemSetByte devId, idx, CByte(value And &HFF&)
    NE2K_MemSetByte devId, idx + 1&, CByte((value \ &H100&) And &HFF&)
End Sub

Private Sub ne2000_setirq(ByVal devId As Long, ByVal irq As Byte)
    ne2k_devs(devId).base_irq = irq
End Sub

Private Sub NE2K_Clear(ByVal devId As Long)
    Dim i As Long

    ne2k_devs(devId).CR.stopBit = 0&
    ne2k_devs(devId).CR.startBit = 0&
    ne2k_devs(devId).CR.tx_packet = 0&
    ne2k_devs(devId).CR.rdma_cmd = 0&
    ne2k_devs(devId).CR.pgsel = 0&

    ne2k_devs(devId).ISR.pkt_rx = 0&
    ne2k_devs(devId).ISR.pkt_tx = 0&
    ne2k_devs(devId).ISR.rx_err = 0&
    ne2k_devs(devId).ISR.tx_err = 0&
    ne2k_devs(devId).ISR.overwrite = 0&
    ne2k_devs(devId).ISR.cnt_oflow = 0&
    ne2k_devs(devId).ISR.rdma_done = 0&
    ne2k_devs(devId).ISR.resetBit = 0&

    ne2k_devs(devId).IMR.rx_inte = 0&
    ne2k_devs(devId).IMR.tx_inte = 0&
    ne2k_devs(devId).IMR.rxerr_inte = 0&
    ne2k_devs(devId).IMR.txerr_inte = 0&
    ne2k_devs(devId).IMR.overw_inte = 0&
    ne2k_devs(devId).IMR.cofl_inte = 0&
    ne2k_devs(devId).IMR.rdma_inte = 0&

    ne2k_devs(devId).DCR.wdsize = 0&
    ne2k_devs(devId).DCR.endian = 0&
    ne2k_devs(devId).DCR.longaddr = 0&
    ne2k_devs(devId).DCR.loopBit = 0&
    ne2k_devs(devId).DCR.auto_rx = 0&
    ne2k_devs(devId).DCR.fifo_size = 0&

    ne2k_devs(devId).TCR.crc_disable = 0&
    ne2k_devs(devId).TCR.loop_cntl = 0&
    ne2k_devs(devId).TCR.ext_stoptx = 0&
    ne2k_devs(devId).TCR.coll_prio = 0&

    ne2k_devs(devId).TSR.tx_ok = 0&
    ne2k_devs(devId).TSR.collided = 0&
    ne2k_devs(devId).TSR.aborted = 0&
    ne2k_devs(devId).TSR.no_carrier = 0&
    ne2k_devs(devId).TSR.fifo_ur = 0&
    ne2k_devs(devId).TSR.cd_hbeat = 0&
    ne2k_devs(devId).TSR.ow_coll = 0&

    ne2k_devs(devId).RCR.errors_ok = 0&
    ne2k_devs(devId).RCR.runts_ok = 0&
    ne2k_devs(devId).RCR.broadcast = 0&
    ne2k_devs(devId).RCR.multicast = 0&
    ne2k_devs(devId).RCR.promisc = 0&
    ne2k_devs(devId).RCR.monitor = 0&

    ne2k_devs(devId).RSR.rx_ok = 0&
    ne2k_devs(devId).RSR.bad_crc = 0&
    ne2k_devs(devId).RSR.bad_falign = 0&
    ne2k_devs(devId).RSR.fifo_or = 0&
    ne2k_devs(devId).RSR.rx_missed = 0&
    ne2k_devs(devId).RSR.rx_mbit = 0&
    ne2k_devs(devId).RSR.rx_disabled = 0&
    ne2k_devs(devId).RSR.deferred = 0&

    ne2k_devs(devId).tx_timer_active = 0&
    ne2k_devs(devId).local_dma = 0&
    ne2k_devs(devId).page_start = 0&
    ne2k_devs(devId).page_stop = 0&
    ne2k_devs(devId).bound_ptr = 0&
    ne2k_devs(devId).tx_page_start = 0&
    ne2k_devs(devId).num_coll = 0&
    ne2k_devs(devId).tx_bytes = 0&
    ne2k_devs(devId).fifo = 0&
    ne2k_devs(devId).remote_dma = 0&
    ne2k_devs(devId).remote_start = 0&
    ne2k_devs(devId).remote_bytes = 0&
    ne2k_devs(devId).tallycnt_0 = 0&
    ne2k_devs(devId).tallycnt_1 = 0&
    ne2k_devs(devId).tallycnt_2 = 0&

    ne2k_devs(devId).curr_page = 0&
    ne2k_devs(devId).rempkt_ptr = 0&
    ne2k_devs(devId).localpkt_ptr = 0&
    ne2k_devs(devId).address_cnt = 0&

    For i = 0& To NE2K_MEMSIZ - 1&
        ne2k_devs(devId).mem(i) = 0&
    Next i
End Sub

Private Sub ne2000_reset(ByVal devId As Long, ByVal resetType As Long)
    Dim i As Long

    If Not NE2K_IsValid(devId) Then Exit Sub

    ne2k_devs(devId).macaddr(0&) = ne2k_devs(devId).physaddr(0&)
    ne2k_devs(devId).macaddr(1&) = ne2k_devs(devId).physaddr(0&)
    ne2k_devs(devId).macaddr(2&) = ne2k_devs(devId).physaddr(1&)
    ne2k_devs(devId).macaddr(3&) = ne2k_devs(devId).physaddr(1&)
    ne2k_devs(devId).macaddr(4&) = ne2k_devs(devId).physaddr(2&)
    ne2k_devs(devId).macaddr(5&) = ne2k_devs(devId).physaddr(2&)
    ne2k_devs(devId).macaddr(6&) = ne2k_devs(devId).physaddr(3&)
    ne2k_devs(devId).macaddr(7&) = ne2k_devs(devId).physaddr(3&)
    ne2k_devs(devId).macaddr(8&) = ne2k_devs(devId).physaddr(4&)
    ne2k_devs(devId).macaddr(9&) = ne2k_devs(devId).physaddr(4&)
    ne2k_devs(devId).macaddr(10&) = ne2k_devs(devId).physaddr(5&)
    ne2k_devs(devId).macaddr(11&) = ne2k_devs(devId).physaddr(5&)

    For i = 12& To 31&
        ne2k_devs(devId).macaddr(i) = &H57&
    Next i

    NE2K_Clear devId

    ne2k_devs(devId).CR.stopBit = 1&
    ne2k_devs(devId).CR.rdma_cmd = 4&
    ne2k_devs(devId).ISR.resetBit = 1&
    ne2k_devs(devId).DCR.longaddr = 1&

    i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
End Sub
Public Function ne2000_chipmem_read_b(ByVal devId As Long, ByVal address As Long) As Byte
    If Not NE2K_IsValid(devId) Then
        ne2000_chipmem_read_b = &HFF&
        Exit Function
    End If

    If (address >= 0&) And (address <= 31&) Then
        ne2000_chipmem_read_b = ne2k_devs(devId).macaddr(address)
        Exit Function
    End If

    If (address >= NE2K_MEMSTART) And (address < NE2K_MEMEND) Then
        ne2000_chipmem_read_b = ne2k_devs(devId).mem(address - NE2K_MEMSTART)
    Else
        ne2000_chipmem_read_b = &HFF&
    End If
End Function

Public Function ne2000_chipmem_read_w(ByVal devId As Long, ByVal address As Long) As Long
    Dim lo As Long
    Dim hi As Long

    If Not NE2K_IsValid(devId) Then
        ne2000_chipmem_read_w = &HFFFF&
        Exit Function
    End If

    If (address >= 0&) And (address <= 31&) Then
        lo = ne2k_devs(devId).macaddr(address)
        If address < 31& Then hi = ne2k_devs(devId).macaddr(address + 1&)
        ne2000_chipmem_read_w = (lo Or (hi * &H100&)) And &HFFFF&
        Exit Function
    End If

    If (address >= NE2K_MEMSTART) And (address < NE2K_MEMEND) Then
        ne2000_chipmem_read_w = NE2K_MemWordLE(devId, address - NE2K_MEMSTART)
    Else
        ne2000_chipmem_read_w = &HFFFF&
    End If
End Function

Public Sub ne2000_chipmem_write_b(ByVal devId As Long, ByVal address As Long, ByVal value As Byte)
    If Not NE2K_IsValid(devId) Then Exit Sub
    If (address >= NE2K_MEMSTART) And (address < NE2K_MEMEND) Then
        ne2k_devs(devId).mem(address - NE2K_MEMSTART) = value
    End If
End Sub

Public Sub ne2000_chipmem_write_w(ByVal devId As Long, ByVal address As Long, ByVal value As Long)
    If Not NE2K_IsValid(devId) Then Exit Sub
    If (address >= NE2K_MEMSTART) And (address < NE2K_MEMEND) Then
        NE2K_MemSetWordLE devId, address - NE2K_MEMSTART, value
    End If
End Sub

Public Function ne2000_dma_read(ByVal devId As Long, ByVal io_len As Long) As Long
    If Not NE2K_IsValid(devId) Then
        ne2000_dma_read = 0&
        Exit Function
    End If

    With ne2k_devs(devId)
        .remote_dma = (.remote_dma + io_len) And &HFFFF&
        If .remote_dma = ((CLng(.page_stop) And &HFF&) * &H100&) Then
            .remote_dma = (CLng(.page_start) And &HFF&) * &H100&
        End If

        If .remote_bytes > 1& Then
            .remote_bytes = (.remote_bytes - io_len) And &HFFFF&
        Else
            .remote_bytes = 0&
        End If

        If .remote_bytes = 0& Then
            .ISR.rdma_done = 1&
            If .IMR.rdma_inte <> 0& Then
                i8259_doirq .i8259Slot, CByte(.base_irq And &HFF&)
            End If
        End If
    End With

    ne2000_dma_read = 0&
End Function

Public Function ne2000_asic_read_w(ByVal devId As Long, ByVal offset As Integer) As Long
    Dim ret As Long

    If Not NE2K_IsValid(devId) Then
        ne2000_asic_read_w = 0&
        Exit Function
    End If

    If (ne2k_devs(devId).DCR.wdsize And &H1&) <> 0& Then
        ret = ne2000_chipmem_read_w(devId, ne2k_devs(devId).remote_dma)
        ne2000_dma_read devId, 2&
    Else
        ret = ne2000_chipmem_read_b(devId, ne2k_devs(devId).remote_dma)
        ne2000_dma_read devId, 1&
    End If

    ne2000_asic_read_w = ret And &HFFFF&
End Function

Public Sub ne2000_dma_write(ByVal devId As Long, ByVal io_len As Long)
    If Not NE2K_IsValid(devId) Then Exit Sub

    With ne2k_devs(devId)
        .remote_dma = (.remote_dma + io_len) And &HFFFF&
        If .remote_dma = ((CLng(.page_stop) And &HFF&) * &H100&) Then
            .remote_dma = (CLng(.page_start) And &HFF&) * &H100&
        End If

        .remote_bytes = (.remote_bytes - io_len) And &HFFFF&
        If .remote_bytes > NE2K_MEMSIZ Then .remote_bytes = 0&

        If .remote_bytes = 0& Then
            .ISR.rdma_done = 1&
            If .IMR.rdma_inte <> 0& Then
                i8259_doirq .i8259Slot, CByte(.base_irq And &HFF&)
            End If
        End If
    End With
End Sub

Public Sub ne2000_asic_write_w(ByVal devId As Long, ByVal offset As Integer, ByVal value As Long)
    If Not NE2K_IsValid(devId) Then Exit Sub
    If ne2k_devs(devId).remote_bytes = 0& Then Exit Sub

    If (ne2k_devs(devId).DCR.wdsize And &H1&) <> 0& Then
        ne2000_chipmem_write_w devId, ne2k_devs(devId).remote_dma, value
        ne2000_dma_write devId, 2&
    Else
        ne2000_chipmem_write_b devId, ne2k_devs(devId).remote_dma, CByte(value And &HFF&)
        ne2000_dma_write devId, 1&
    End If
End Sub

Public Function ne2000_asic_read_b(ByVal devId As Long, ByVal offset As Integer) As Byte
    Dim w As Long

    If (offset And 1&) <> 0& Then
        w = ne2000_asic_read_w(devId, offset And &HFFFE&)
        ne2000_asic_read_b = CByte((w \ &H100&) And &HFF&)
    Else
        w = ne2000_asic_read_w(devId, offset)
        ne2000_asic_read_b = CByte(w And &HFF&)
    End If
End Function

Public Sub ne2000_asic_write_b(ByVal devId As Long, ByVal offset As Integer, ByVal value As Byte)
    If (offset And 1&) <> 0& Then
        ne2000_asic_write_w devId, (offset And &HFFFE&), ((CLng(value) And &HFF&) * &H100&)
    Else
        ne2000_asic_write_w devId, offset, value
    End If
End Sub

Public Function ne2000_reset_read(ByVal devId As Long, ByVal offset As Integer) As Byte
    ne2000_reset devId, NE2K_RESET_SOFTWARE
    ne2000_reset_read = 0&
End Function

Public Sub ne2000_reset_write(ByVal devId As Long, ByVal offset As Integer, ByVal value As Byte)
    ' no-op in C
End Sub

Private Function NE2K_ISRValue(ByVal devId As Long) As Byte
    With ne2k_devs(devId).ISR
        NE2K_ISRValue = CByte(((.resetBit And 1&) * &H80&) Or ((.rdma_done And 1&) * &H40&) Or ((.cnt_oflow And 1&) * &H20&) Or ((.overwrite And 1&) * &H10&) Or ((.tx_err And 1&) * &H8&) Or ((.rx_err And 1&) * &H4&) Or ((.pkt_tx And 1&) * &H2&) Or (.pkt_rx And 1&))
    End With
End Function

Private Function NE2K_IMRValue(ByVal devId As Long) As Byte
    With ne2k_devs(devId).IMR
        NE2K_IMRValue = CByte(((.rdma_inte And 1&) * &H40&) Or ((.cofl_inte And 1&) * &H20&) Or ((.overw_inte And 1&) * &H10&) Or ((.txerr_inte And 1&) * &H8&) Or ((.rxerr_inte And 1&) * &H4&) Or ((.tx_inte And 1&) * &H2&) Or (.rx_inte And 1&))
    End With
End Function

Public Function ne2000_read(ByVal devId As Long, ByVal address As Integer) As Byte
    Dim ret As Long

    If Not NE2K_IsValid(devId) Then
        ne2000_read = &HFF&
        Exit Function
    End If

    ret = 0&
    address = (address And &HF&)

    If address = 0& Then
        With ne2k_devs(devId).CR
            ret = (((.pgsel And &H3&) * &H40&) Or ((.rdma_cmd And &H7&) * &H8&) Or ((.tx_packet And 1&) * &H4&) Or ((.startBit And 1&) * &H2&) Or (.stopBit And 1&))
        End With
    Else
        Select Case ne2k_devs(devId).CR.pgsel
            Case &H0&
                Select Case address
                    Case &H1&: ret = ne2k_devs(devId).local_dma And &HFF&
                    Case &H2&: ret = (ne2k_devs(devId).local_dma \ &H100&) And &HFF&
                    Case &H3&: ret = ne2k_devs(devId).bound_ptr
                    Case &H4&
                        With ne2k_devs(devId).TSR
                            ret = ((.ow_coll And 1&) * &H80&) Or ((.cd_hbeat And 1&) * &H40&) Or ((.fifo_ur And 1&) * &H20&) Or ((.no_carrier And 1&) * &H10&) Or ((.aborted And 1&) * &H8&) Or ((.collided And 1&) * &H4&) Or (.tx_ok And 1&)
                        End With
                    Case &H5&: ret = ne2k_devs(devId).num_coll
                    Case &H6&: ret = ne2k_devs(devId).fifo
                    Case &H7&: ret = NE2K_ISRValue(devId)
                    Case &H8&: ret = ne2k_devs(devId).remote_dma And &HFF&
                    Case &H9&: ret = (ne2k_devs(devId).remote_dma \ &H100&) And &HFF&
                    Case &HA&, &HB&: ret = &HFF&
                    Case &HC&
                        With ne2k_devs(devId).RSR
                            ret = ((.deferred And 1&) * &H80&) Or ((.rx_disabled And 1&) * &H40&) Or ((.rx_mbit And 1&) * &H20&) Or ((.rx_missed And 1&) * &H10&) Or ((.fifo_or And 1&) * &H8&) Or ((.bad_falign And 1&) * &H4&) Or ((.bad_crc And 1&) * &H2&) Or (.rx_ok And 1&)
                        End With
                    Case &HD&: ret = ne2k_devs(devId).tallycnt_0
                    Case &HE&: ret = ne2k_devs(devId).tallycnt_1
                    Case &HF&: ret = ne2k_devs(devId).tallycnt_2
                End Select

            Case &H1&
                Select Case address
                    Case &H1& To &H6&: ret = ne2k_devs(devId).physaddr(address - 1&)
                    Case &H7&: ret = ne2k_devs(devId).curr_page
                    Case &H8& To &HF&: ret = ne2k_devs(devId).mchash(address - 8&)
                End Select

            Case &H2&
                Select Case address
                    Case &H1&: ret = ne2k_devs(devId).page_start
                    Case &H2&: ret = ne2k_devs(devId).page_stop
                    Case &H3&: ret = ne2k_devs(devId).rempkt_ptr
                    Case &H4&: ret = ne2k_devs(devId).tx_page_start
                    Case &H5&: ret = ne2k_devs(devId).localpkt_ptr
                    Case &H6&: ret = (ne2k_devs(devId).address_cnt \ &H100&) And &HFF&
                    Case &H7&: ret = ne2k_devs(devId).address_cnt And &HFF&
                    Case &HC&
                        With ne2k_devs(devId).RCR
                            ret = ((.monitor And 1&) * &H20&) Or ((.promisc And 1&) * &H10&) Or ((.multicast And 1&) * &H8&) Or ((.broadcast And 1&) * &H4&) Or ((.runts_ok And 1&) * &H2&) Or (.errors_ok And 1&)
                        End With
                    Case &HD&
                        With ne2k_devs(devId).TCR
                            ret = ((.coll_prio And 1&) * &H10&) Or ((.ext_stoptx And 1&) * &H8&) Or ((.loop_cntl And &H3&) * &H2&) Or (.crc_disable And 1&)
                        End With
                    Case &HE&
                        With ne2k_devs(devId).DCR
                            ret = ((.fifo_size And &H3&) * &H20&) Or ((.auto_rx And 1&) * &H10&) Or ((.loopBit And 1&) * &H8&) Or ((.longaddr And 1&) * &H4&) Or ((.endian And 1&) * &H2&) Or (.wdsize And 1&)
                        End With
                    Case &HF&
                        ret = NE2K_IMRValue(devId)
                End Select

            Case &H3&
                Select Case address
                    Case 0&: ret = ne2k_devs(devId).crRaw
                    Case 1&: ret = ne2k_devs(devId).i9346cr
                    Case 3&: ret = ne2k_devs(devId).config0
                    Case 5&: ret = ne2k_devs(devId).config2
                    Case 6&: ret = ne2k_devs(devId).config3
                    Case 9&: ret = &HFF&
                    Case &HE&: ret = ne2k_devs(devId).i8029asid0
                    Case &HF&: ret = ne2k_devs(devId).i8029asid1
                End Select
        End Select
    End If

    ne2000_read = CByte(ret And &HFF&)
End Function
Private Sub NE2K_CopyFromMem(ByVal devId As Long, ByVal srcOff As Long, ByRef outBuf() As Byte, ByVal count As Long)
    Dim i As Long

    If count <= 0& Then
        ReDim outBuf(0& To 0&) As Byte
        Exit Sub
    End If

    ReDim outBuf(0& To count - 1&) As Byte
    For i = 0& To count - 1&
        outBuf(i) = NE2K_MemByte(devId, srcOff + i)
    Next i
End Sub

Public Sub ne2000_write(ByVal devId As Long, ByVal address As Integer, ByVal value As Byte)
    Dim oldStart As Byte
    Dim memOff As Long
    Dim txbuf() As Byte
    Dim pending As Byte
    Dim microsecs As Double

    If Not NE2K_IsValid(devId) Then Exit Sub

    address = (address And &HF&)

    If address = 0& Then
        If (value And &H38&) = 0& Then value = (value Or &H20&)

        If (value And &H1&) <> 0& Then
            ne2k_devs(devId).ISR.resetBit = 1&
            ne2k_devs(devId).CR.stopBit = 1&
        Else
            ne2k_devs(devId).CR.stopBit = 0&
        End If

        ne2k_devs(devId).CR.rdma_cmd = CByte((value And &H38&) \ &H8&)

        oldStart = ne2k_devs(devId).CR.startBit
        If ((value And &H2&) <> 0&) And (oldStart = 0&) Then
            ne2k_devs(devId).ISR.resetBit = 0&
        End If

        ne2k_devs(devId).CR.startBit = NE2K_Flag((value And &H2&) = &H2&)
        ne2k_devs(devId).CR.pgsel = CByte((value And &HC0&) \ &H40&)

        If ne2k_devs(devId).CR.rdma_cmd = 3& Then
            ne2k_devs(devId).remote_start = (CLng(ne2k_devs(devId).bound_ptr) And &HFF&) * &H100&
            ne2k_devs(devId).remote_dma = ne2k_devs(devId).remote_start
            memOff = ((CLng(ne2k_devs(devId).bound_ptr) And &HFF&) * &H100&) + 2& - NE2K_MEMSTART
            ne2k_devs(devId).remote_bytes = NE2K_MemWordLE(devId, memOff)
        End If

        If ((value And &H4&) <> 0&) And (ne2k_devs(devId).TCR.loop_cntl <> 0&) Then
            If ne2k_devs(devId).TCR.loop_cntl = 1& Then
                memOff = ((CLng(ne2k_devs(devId).tx_page_start) And &HFF&) * &H100&) - NE2K_MEMSTART
                NE2K_CopyFromMem devId, memOff, txbuf, ne2k_devs(devId).tx_bytes
                ne2000_rx_frame devId, txbuf, ne2k_devs(devId).tx_bytes

                If (ne2k_devs(devId).IMR.tx_inte <> 0&) And (ne2k_devs(devId).ISR.pkt_tx = 0&) Then
                    i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
                End If
                ne2k_devs(devId).ISR.pkt_tx = 1&
            End If

        ElseIf (value And &H4&) <> 0& Then
            memOff = ((CLng(ne2k_devs(devId).tx_page_start) And &HFF&) * &H100&) - NE2K_MEMSTART
            NE2K_CopyFromMem devId, memOff, txbuf, ne2k_devs(devId).tx_bytes
            pcap_txPacket txbuf, ne2k_devs(devId).tx_bytes

            microsecs = (64# + 96# + 4# * 8# + CDbl(ne2k_devs(devId).tx_bytes) * 8#) / 10#
            microsecs = microsecs * (timing_getFreq() / 1000000#)
            ne2000_tx_event devId, microsecs
        End If

        If (ne2k_devs(devId).CR.rdma_cmd = 1&) And (ne2k_devs(devId).CR.startBit <> 0&) And (ne2k_devs(devId).remote_bytes = 0&) Then
            ne2k_devs(devId).ISR.rdma_done = 1&
            If ne2k_devs(devId).IMR.rdma_inte <> 0& Then
                i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
            End If
        End If

    Else
        Select Case ne2k_devs(devId).CR.pgsel
            Case &H0&
                Select Case address
                    Case &H1&: ne2k_devs(devId).page_start = value
                    Case &H2&: ne2k_devs(devId).page_stop = value
                    Case &H3&: ne2k_devs(devId).bound_ptr = value
                    Case &H4&: ne2k_devs(devId).tx_page_start = value
                    Case &H5&: ne2k_devs(devId).tx_bytes = ((ne2k_devs(devId).tx_bytes And &HFF00&) Or (value And &HFF&))
                    Case &H6&: ne2k_devs(devId).tx_bytes = ((ne2k_devs(devId).tx_bytes And &HFF&) Or ((CLng(value) And &HFF&) * &H100&))
                    Case &H7&
                        value = (value And &H7F&)
                        If (value And &H1&) <> 0& Then ne2k_devs(devId).ISR.pkt_rx = 0&
                        If (value And &H2&) <> 0& Then ne2k_devs(devId).ISR.pkt_tx = 0&
                        If (value And &H4&) <> 0& Then ne2k_devs(devId).ISR.rx_err = 0&
                        If (value And &H8&) <> 0& Then ne2k_devs(devId).ISR.tx_err = 0&
                        If (value And &H10&) <> 0& Then ne2k_devs(devId).ISR.overwrite = 0&
                        If (value And &H20&) <> 0& Then ne2k_devs(devId).ISR.cnt_oflow = 0&
                        If (value And &H40&) <> 0& Then ne2k_devs(devId).ISR.rdma_done = 0&

                        pending = (NE2K_ISRValue(devId) And &H7F&)
                        pending = (pending And NE2K_IMRValue(devId))
                        If pending <> 0& Then
                            i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
                        End If

                    Case &H8&
                        ne2k_devs(devId).remote_start = ((ne2k_devs(devId).remote_start And &HFF00&) Or (value And &HFF&))
                        ne2k_devs(devId).remote_dma = ne2k_devs(devId).remote_start
                    Case &H9&
                        ne2k_devs(devId).remote_start = ((ne2k_devs(devId).remote_start And &HFF&) Or ((CLng(value) And &HFF&) * &H100&))
                        ne2k_devs(devId).remote_dma = ne2k_devs(devId).remote_start
                    Case &HA&
                        ne2k_devs(devId).remote_bytes = ((ne2k_devs(devId).remote_bytes And &HFF00&) Or (value And &HFF&))
                    Case &HB&
                        ne2k_devs(devId).remote_bytes = ((ne2k_devs(devId).remote_bytes And &HFF&) Or ((CLng(value) And &HFF&) * &H100&))
                    Case &HC&
                        ne2k_devs(devId).RCR.errors_ok = NE2K_Flag((value And &H1&) = &H1&)
                        ne2k_devs(devId).RCR.runts_ok = NE2K_Flag((value And &H2&) = &H2&)
                        ne2k_devs(devId).RCR.broadcast = NE2K_Flag((value And &H4&) = &H4&)
                        ne2k_devs(devId).RCR.multicast = NE2K_Flag((value And &H8&) = &H8&)
                        ne2k_devs(devId).RCR.promisc = NE2K_Flag((value And &H10&) = &H10&)
                        ne2k_devs(devId).RCR.monitor = NE2K_Flag((value And &H20&) = &H20&)
                    Case &HD&
                        If (value And &H6&) <> 0& Then
                            ne2k_devs(devId).TCR.loop_cntl = CByte((value And &H6&) \ &H2&)
                        Else
                            ne2k_devs(devId).TCR.loop_cntl = 0&
                        End If
                        If (value And &H1&) <> 0& Then Exit Sub
                        ne2k_devs(devId).TCR.coll_prio = NE2K_Flag((value And &H8&) = &H8&)
                    Case &HE&
                        ne2k_devs(devId).DCR.wdsize = NE2K_Flag((value And &H1&) = &H1&)
                        ne2k_devs(devId).DCR.endian = NE2K_Flag((value And &H2&) = &H2&)
                        ne2k_devs(devId).DCR.longaddr = NE2K_Flag((value And &H4&) = &H4&)
                        ne2k_devs(devId).DCR.loopBit = NE2K_Flag((value And &H8&) = &H8&)
                        ne2k_devs(devId).DCR.auto_rx = NE2K_Flag((value And &H10&) = &H10&)
                        ne2k_devs(devId).DCR.fifo_size = CByte((value And &H50&) \ &H20&)
                    Case &HF&
                        ne2k_devs(devId).IMR.rx_inte = NE2K_Flag((value And &H1&) = &H1&)
                        ne2k_devs(devId).IMR.tx_inte = NE2K_Flag((value And &H2&) = &H2&)
                        ne2k_devs(devId).IMR.rxerr_inte = NE2K_Flag((value And &H4&) = &H4&)
                        ne2k_devs(devId).IMR.txerr_inte = NE2K_Flag((value And &H8&) = &H8&)
                        ne2k_devs(devId).IMR.overw_inte = NE2K_Flag((value And &H10&) = &H10&)
                        ne2k_devs(devId).IMR.cofl_inte = NE2K_Flag((value And &H20&) = &H20&)
                        ne2k_devs(devId).IMR.rdma_inte = NE2K_Flag((value And &H40&) = &H40&)
                        pending = (NE2K_ISRValue(devId) And &H7F&)
                        pending = (pending And NE2K_IMRValue(devId))
                        If pending <> 0& Then
                            i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
                        End If
                End Select

            Case &H1&
                Select Case address
                    Case &H1& To &H6&: ne2k_devs(devId).physaddr(address - 1&) = value
                    Case &H7&: ne2k_devs(devId).curr_page = value
                    Case &H8& To &HF&: ne2k_devs(devId).mchash(address - 8&) = value
                End Select

            Case &H2&
                Select Case address
                    Case &H1&: ne2k_devs(devId).local_dma = ((ne2k_devs(devId).local_dma And &HFF00&) Or (value And &HFF&))
                    Case &H2&: ne2k_devs(devId).local_dma = ((ne2k_devs(devId).local_dma And &HFF&) Or ((CLng(value) And &HFF&) * &H100&))
                    Case &H3&: ne2k_devs(devId).rempkt_ptr = value
                    Case &H5&: ne2k_devs(devId).localpkt_ptr = value
                    Case &H6&: ne2k_devs(devId).address_cnt = ((ne2k_devs(devId).address_cnt And &HFF&) Or ((CLng(value) And &HFF&) * &H100&))
                    Case &H7&: ne2k_devs(devId).address_cnt = ((ne2k_devs(devId).address_cnt And &HFF00&) Or (value And &HFF&))
                End Select

            Case &H3&
                Select Case address
                    Case 0&: ne2k_devs(devId).crRaw = value
                    Case 1&: ne2k_devs(devId).i9346cr = value
                    Case 5&: If (ne2k_devs(devId).i9346cr And &HC0&) = &HC0& Then ne2k_devs(devId).config2 = value
                    Case 6&: If (ne2k_devs(devId).i9346cr And &HC0&) = &HC0& Then ne2k_devs(devId).config3 = value
                    Case 9&: ne2k_devs(devId).hltclk = value
                End Select
        End Select
    End If
End Sub

Private Function mcast_index(ByRef dst() As Byte) As Long
    Const POLYNOMIAL As Long = &H4C11DB6&

    Dim crc As Long
    Dim carry As Long
    Dim i As Long
    Dim j As Long
    Dim b As Long

    crc = -1&

    For i = 0& To 5&
        b = NE2K_PktByte(dst, i)
        For j = 0& To 7&
            carry = NE2K_Flag(((crc And &H80000000&) <> 0&) Xor ((b And &H1&) <> 0&))
            crc = U32Shl(crc, 1&)
            b = (b \ 2&) And &H7FFFFFFF&
            If carry <> 0& Then
                crc = ((crc Xor POLYNOMIAL) Or carry)
            End If
        Next j
    Next i

    mcast_index = (U32Shr(crc, 26&) And &H3F&)
End Function

Private Function NE2K_IsBroadcast(ByRef pkt() As Byte) As Boolean
    Dim i As Long
    For i = 0& To 5&
        If NE2K_PktByte(pkt, i) <> &HFF& Then
            NE2K_IsBroadcast = False
            Exit Function
        End If
    Next i
    NE2K_IsBroadcast = True
End Function

Private Function NE2K_IsPhysAddr(ByVal devId As Long, ByRef pkt() As Byte) As Boolean
    Dim i As Long
    For i = 0& To 5&
        If NE2K_PktByte(pkt, i) <> ne2k_devs(devId).physaddr(i) Then
            NE2K_IsPhysAddr = False
            Exit Function
        End If
    Next i
    NE2K_IsPhysAddr = True
End Function

Private Sub NE2K_CopyPacketToMem(ByVal devId As Long, ByVal dstOff As Long, ByRef src() As Byte, ByVal srcOff As Long, ByVal count As Long)
    Dim i As Long
    For i = 0& To count - 1&
        NE2K_MemSetByte devId, dstOff + i, NE2K_PktByte(src, srcOff + i)
    Next i
End Sub
Public Function ne2000_rx_frame_try(ByVal devId As Long, ByRef pkt() As Byte, ByVal io_len As Long) As Long
    Dim pages As Long
    Dim avail As Long
    Dim idx As Long
    Dim nextpage As Long
    Dim pkthdr() As Byte
    Dim startOff As Long
    Dim endbytes As Long

    ne2000_rx_frame_try = NE2000_RX_RETRY

    If Not NE2K_IsValid(devId) Then
        ne2000_rx_frame_try = NE2000_RX_DROP
        Exit Function
    End If
    If (ne2k_devs(devId).CR.stopBit <> 0&) Or (ne2k_devs(devId).page_start = 0&) Then
        ne2000_rx_frame_try = NE2000_RX_RETRY
        Exit Function
    End If

    pages = (io_len + 4& + 4& + 255&) \ 256&

    If ne2k_devs(devId).curr_page < ne2k_devs(devId).bound_ptr Then
        avail = ne2k_devs(devId).bound_ptr - ne2k_devs(devId).curr_page
    Else
        avail = (ne2k_devs(devId).page_stop - ne2k_devs(devId).page_start) - (ne2k_devs(devId).curr_page - ne2k_devs(devId).bound_ptr)
    End If

    If avail < pages Then Exit Function
    If (NE2K_NEVER_FULL_RING <> 0&) And (avail = pages) Then Exit Function

    If (io_len < 40&) And (ne2k_devs(devId).RCR.runts_ok = 0&) Then
        ne2000_rx_frame_try = NE2000_RX_DROP
        Exit Function
    End If
    If io_len < 60& Then io_len = 60&

    If ne2k_devs(devId).RCR.promisc = 0& Then
        If NE2K_IsBroadcast(pkt) Then
            If ne2k_devs(devId).RCR.broadcast = 0& Then
                ne2000_rx_frame_try = NE2000_RX_DROP
                Exit Function
            End If
        ElseIf (NE2K_PktByte(pkt, 0&) And &H1&) <> 0& Then
            If ne2k_devs(devId).RCR.multicast = 0& Then
                ne2000_rx_frame_try = NE2000_RX_DROP
                Exit Function
            End If
            idx = mcast_index(pkt)
            If (ne2k_devs(devId).mchash(idx \ 8&) And (2& ^ (idx And 7&))) = 0& Then
                ne2000_rx_frame_try = NE2000_RX_DROP
                Exit Function
            End If
        ElseIf Not NE2K_IsPhysAddr(devId, pkt) Then
            ne2000_rx_frame_try = NE2000_RX_DROP
            Exit Function
        End If
    End If

    nextpage = ne2k_devs(devId).curr_page + pages
    If nextpage >= ne2k_devs(devId).page_stop Then
        nextpage = nextpage - (ne2k_devs(devId).page_stop - ne2k_devs(devId).page_start)
    End If

    ReDim pkthdr(0& To 3&) As Byte
    pkthdr(0&) = 1&
    If (NE2K_PktByte(pkt, 0&) And &H1&) <> 0& Then pkthdr(0&) = (pkthdr(0&) Or &H20&)
    pkthdr(1&) = CByte(nextpage And &HFF&)
    pkthdr(2&) = CByte((io_len + 4&) And &HFF&)
    pkthdr(3&) = CByte(((io_len + 4&) \ &H100&) And &HFF&)

    startOff = (CLng(ne2k_devs(devId).curr_page) * 256&) - NE2K_MEMSTART

    If (nextpage > ne2k_devs(devId).curr_page) Or ((ne2k_devs(devId).curr_page + pages) = ne2k_devs(devId).page_stop) Then
        NE2K_CopyPacketToMem devId, startOff, pkthdr, 0&, 4&
        NE2K_CopyPacketToMem devId, startOff + 4&, pkt, 0&, io_len
        ne2k_devs(devId).curr_page = CByte(nextpage And &HFF&)
    Else
        endbytes = (ne2k_devs(devId).page_stop - ne2k_devs(devId).curr_page) * 256&
        NE2K_CopyPacketToMem devId, startOff, pkthdr, 0&, 4&
        If (endbytes - 4&) > 0& Then
            NE2K_CopyPacketToMem devId, startOff + 4&, pkt, 0&, (endbytes - 4&)
        End If
        startOff = (CLng(ne2k_devs(devId).page_start) * 256&) - NE2K_MEMSTART
        NE2K_CopyPacketToMem devId, startOff, pkt, (endbytes - 4&), (io_len - endbytes + 8&)
        ne2k_devs(devId).curr_page = CByte(nextpage And &HFF&)
    End If

    ne2k_devs(devId).RSR.rx_ok = 1&
    If (NE2K_PktByte(pkt, 0&) And &H80&) <> 0& Then ne2k_devs(devId).RSR.rx_mbit = 1&

    ne2k_devs(devId).ISR.pkt_rx = 1&
    If ne2k_devs(devId).IMR.rx_inte <> 0& Then
        i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
    End If

    ne2000_rx_frame_try = NE2000_RX_ACCEPTED
End Function

Public Sub ne2000_rx_frame(ByVal devId As Long, ByRef pkt() As Byte, ByVal io_len As Long)
    Call ne2000_rx_frame_try(devId, pkt, io_len)
End Sub

Public Sub ne2000_tx_timer(ByVal devId As Long)
    If Not NE2K_IsValid(devId) Then Exit Sub

    timing_timerDisable ne2k_devs(devId).tx_timer
    ne2k_devs(devId).TSR.tx_ok = 1&

    If (ne2k_devs(devId).IMR.tx_inte <> 0&) And (ne2k_devs(devId).ISR.pkt_tx = 0&) Then
        i8259_doirq ne2k_devs(devId).i8259Slot, CByte(ne2k_devs(devId).base_irq And &HFF&)
    End If

    ne2k_devs(devId).ISR.pkt_tx = 1&
    ne2k_devs(devId).tx_timer_active = 0&
End Sub

Public Sub ne2000_tx_event(ByVal devId As Long, ByVal interval As Double)
    If Not NE2K_IsValid(devId) Then Exit Sub
    timing_updateInterval ne2k_devs(devId).tx_timer, interval
    timing_timerEnable ne2k_devs(devId).tx_timer
End Sub

Public Sub ne2000_init(ByRef machine As MACHINE_t, ByVal baseport As Long, ByVal irq As Byte)
    Dim devId As Long
    Dim i As Long

    devId = -1&
    For i = 0& To NE2K_MAX - 1&
        If ne2k_devs(i).used = 0& Then
            devId = i
            Exit For
        End If
    Next i

    If devId < 0& Then
        debug_log DEBUG_ERROR, "[NE2000] Out of NE2000 slots"
        Exit Sub
    End If

    debug_log DEBUG_INFO, "[NE2000] Initializing NE2000 Ethernet adapter at 0x" & Right$("000" & Hex$(baseport And &HFFFF&), 3&) & ", IRQ " & CStr(irq)

    ne2k_devs(devId).used = 1&
    ne2k_devs(devId).i8259Slot = machine.i8259
    ne2k_devs(devId).base_address = (baseport And &HFFFF&)

    ne2k_devs(devId).physaddr(0&) = &HAC&
    ne2k_devs(devId).physaddr(1&) = &HDE&
    ne2k_devs(devId).physaddr(2&) = &H48&
    ne2k_devs(devId).physaddr(3&) = &H88&
    ne2k_devs(devId).physaddr(4&) = &HBB&
    ne2k_devs(devId).physaddr(5&) = &HAB&

    ports_cbRegister (baseport And &HFFFF&), &H10&, PORTS_CB_NE2000_REG, PORTS_CB_NONE, PORTS_CB_NE2000_REG, PORTS_CB_NONE, devId
    ports_cbRegister ((baseport + &H10&) And &HFFFF&), &H10&, PORTS_CB_NE2000_ASIC, PORTS_CB_NE2000_ASIC, PORTS_CB_NE2000_ASIC, PORTS_CB_NE2000_ASIC, devId
    ports_cbRegister ((baseport + &H1F&) And &HFFFF&), &H1&, PORTS_CB_NE2000_RESET, PORTS_CB_NONE, PORTS_CB_NE2000_RESET, PORTS_CB_NONE, devId

    ne2000_setirq devId, irq
    ne2000_reset devId, NE2K_RESET_HARDWARE

    ne2k_devs(devId).tx_timer = timing_addTimer(TIMER_CB_NE2000_TX, devId, 1000#, TIMING_DISABLED)

    ne2k_primary = devId
End Sub


