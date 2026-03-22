Attribute VB_Name = "modMemory"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)

Public Const MEMORY_RANGE As Long = &H4000000&
Public Const MEMORY_MASK As Long = &H3FFFFFF&

Public Const MEMORY_CB_NONE As Long = 0&
Public Const MEMORY_CB_VGA As Long = 3&
Public Const MEMORY_CB_BUSLOGIC_ROM As Long = 4&
Public Const MEMORY_CB_ET4000 As Long = 5&
Public Const MEM_ACCESS_USER_READ As Long = 0&
Public Const MEM_ACCESS_USER_WRITE As Long = 1&
Public Const MEM_ACCESS_SUPERVISOR_READ As Long = 2&
Public Const MEM_ACCESS_SUPERVISOR_WRITE As Long = 3&
Private Const WRCACHE_INDEX_THRESHOLD As Long = 16&
Private Const WRCACHE_INDEX_BUCKETS As Long = 1024&
Private Const WRCACHE_INDEX_MASK As Long = WRCACHE_INDEX_BUCKETS - 1&

Private Type WRCACHE_t
    addr As Long
    value As Byte
End Type

Private Type PAGE_WALK_RESULT_t
    phys_base As Long
    pte_addr As Long
    entry_flags As Byte
End Type

Private mem_start(0& To 31&) As Long
Private mem_size(0& To 31&) As Long
Private mem_hasRead(0& To 31&) As Byte
Private mem_hasWrite(0& To 31&) As Byte
Private mem_readcb(0& To 31&) As Long
Private mem_writecb(0& To 31&) As Long
Private mem_udata(0& To 31&) As Long
Private mem_used(0& To 31&) As Byte
Private mem_readOffset(0& To 31&) As Long
Private mem_writeOffset(0& To 31&) As Long

Private mem_pool() As Byte
Private mem_poolSize As Long

Private mem_map_lookup(0& To (2& ^ 20&) - 1&) As Byte
Private wrcache(0& To 511&) As WRCACHE_t
Private wrcache_next(0& To 511&) As Long
Private wrcache_bucketHead(0& To WRCACHE_INDEX_BUCKETS - 1&) As Long
Private wrcache_bucketTail(0& To WRCACHE_INDEX_BUCKETS - 1&) As Long
Private wrcache_bucketStamp(0& To WRCACHE_INDEX_BUCKETS - 1&) As Long
Public wrcache_count As Long
Private wrcache_indexEnabled As Long
Private wrcache_epoch As Long

Private Function getmap(ByVal addr32 As Long) As Byte
    getmap = mem_map_lookup((((addr32 And &HFFFFF000) \ &H1000&) And &HFFFFF&))
End Function

Private Sub wrcache_advance_epoch()
    Dim i As Long

    If wrcache_epoch = &H7FFFFFFF Then
        For i = 0& To WRCACHE_INDEX_BUCKETS - 1&
            wrcache_bucketStamp(i) = 0&
        Next i
        wrcache_epoch = 1&
    Else
        wrcache_epoch = wrcache_epoch + 1&
    End If
End Sub

Private Sub wrcache_reset_state()
    wrcache_count = 0&
    wrcache_indexEnabled = 0&
    wrcache_advance_epoch
End Sub

Private Function wrcache_getBucket(ByVal addr32 As Long) As Long
    wrcache_getBucket = (addr32 And WRCACHE_INDEX_MASK)
End Function

Private Sub wrcache_indexAppend(ByVal idx As Long)
    Dim bucket As Long

    bucket = wrcache_getBucket(wrcache(idx).addr)
    wrcache_next(idx) = -1&

    If wrcache_bucketStamp(bucket) <> wrcache_epoch Then
        wrcache_bucketStamp(bucket) = wrcache_epoch
        wrcache_bucketHead(bucket) = idx
        wrcache_bucketTail(bucket) = idx
    Else
        wrcache_next(wrcache_bucketTail(bucket)) = idx
        wrcache_bucketTail(bucket) = idx
    End If
End Sub

Private Sub wrcache_buildIndex()
    Dim i As Long

    wrcache_advance_epoch
    wrcache_indexEnabled = 1&

    For i = 0& To wrcache_count - 1&
        wrcache_indexAppend i
    Next i
End Sub

Private Sub wrcache_append(ByVal addr32 As Long, ByVal value As Byte)
    Dim idx As Long

    idx = wrcache_count
    wrcache(idx).addr = addr32
    wrcache(idx).value = value
    wrcache_count = idx + 1&

    If wrcache_indexEnabled <> 0& Then
        wrcache_indexAppend idx
    ElseIf wrcache_count > WRCACHE_INDEX_THRESHOLD Then
        wrcache_buildIndex
    End If

    If wrcache_count = 512& Then
        debug_log DEBUG_ERROR, "FATAL: wrcache_count == 512"
        End
    End If
End Sub

Public Sub wrcache_init()
    wrcache_reset_state
End Sub

Public Sub wrcache_flush()
    Dim i As Long
    Dim map As Byte

    For i = 0& To wrcache_count - 1&
        map = getmap(wrcache(i).addr)
        If map = &HFF& Then
            GoTo NextWrite
        End If

        If mem_hasWrite(map) <> 0& Then
            Memory_WriteMappedByte map, wrcache(i).addr, wrcache(i).value
            GoTo NextWrite
        End If

        If mem_writecb(map) <> MEMORY_CB_NONE Then
            Memory_DispatchWriteCB mem_writecb(map), mem_udata(map), wrcache(i).addr, wrcache(i).value
            GoTo NextWrite
        End If

NextWrite:
    Next i

    wrcache_reset_state
End Sub

Public Sub wrcache_write(ByVal addr32 As Long, ByVal value As Byte)
    wrcache_append addr32, value
End Sub

Private Sub wrcache_writew(ByVal addr0 As Long, ByVal addr1 As Long, ByVal b0 As Byte, ByVal b1 As Byte)
    wrcache_append addr0, b0
    wrcache_append addr1, b1
End Sub

Private Sub wrcache_writel(ByVal addr0 As Long, ByVal addr1 As Long, ByVal addr2 As Long, ByVal addr3 As Long, ByVal b0 As Byte, ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte)
    wrcache_append addr0, b0
    wrcache_append addr1, b1
    wrcache_append addr2, b2
    wrcache_append addr3, b3
End Sub

Public Function wrcache_read(ByVal addr32 As Long, ByRef dst As Byte) As Long
    Dim bucket As Long
    Dim i As Long

    If wrcache_count = 0& Then
        wrcache_read = 0&
        Exit Function
    End If

    If wrcache_indexEnabled = 0& Then
        For i = 0& To wrcache_count - 1&
            If wrcache(i).addr = addr32 Then
                dst = wrcache(i).value
                wrcache_read = 1&
                Exit Function
            End If
        Next i
    Else
        bucket = wrcache_getBucket(addr32)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr32 Then
                    dst = wrcache(i).value
                    wrcache_read = 1&
                    Exit Function
                End If
                i = wrcache_next(i)
            Loop
        End If
    End If

    wrcache_read = 0&
End Function

Private Function wrcache_readw(ByVal addr0 As Long, ByVal addr1 As Long, ByRef b0 As Long, ByRef b1 As Long) As Long
    Dim bucket As Long
    Dim i As Long
    Dim mask As Long

    If wrcache_count = 0& Then
        wrcache_readw = 0&
        Exit Function
    End If

    If wrcache_indexEnabled = 0& Then
        For i = 0& To wrcache_count - 1&
            If ((mask And 1&) = 0&) And (wrcache(i).addr = addr0) Then
                b0 = (wrcache(i).value And &HFF&)
                mask = (mask Or 1&)
            End If
            If ((mask And 2&) = 0&) And (wrcache(i).addr = addr1) Then
                b1 = (wrcache(i).value And &HFF&)
                mask = (mask Or 2&)
            End If
            If mask = 3& Then Exit For
        Next i
    Else
        bucket = wrcache_getBucket(addr0)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr0 Then
                    b0 = (wrcache(i).value And &HFF&)
                    mask = 1&
                    Exit Do
                End If
                i = wrcache_next(i)
            Loop
        End If

        bucket = wrcache_getBucket(addr1)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr1 Then
                    b1 = (wrcache(i).value And &HFF&)
                    mask = (mask Or 2&)
                    Exit Do
                End If
                i = wrcache_next(i)
            Loop
        End If
    End If

    wrcache_readw = mask
End Function

Private Function wrcache_readl(ByVal addr0 As Long, ByVal addr1 As Long, ByVal addr2 As Long, ByVal addr3 As Long, ByRef b0 As Long, ByRef b1 As Long, ByRef b2 As Long, ByRef b3 As Long) As Long
    Dim bucket As Long
    Dim i As Long
    Dim mask As Long

    If wrcache_count = 0& Then
        wrcache_readl = 0&
        Exit Function
    End If

    If wrcache_indexEnabled = 0& Then
        For i = 0& To wrcache_count - 1&
            If ((mask And 1&) = 0&) And (wrcache(i).addr = addr0) Then
                b0 = (wrcache(i).value And &HFF&)
                mask = (mask Or 1&)
            End If
            If ((mask And 2&) = 0&) And (wrcache(i).addr = addr1) Then
                b1 = (wrcache(i).value And &HFF&)
                mask = (mask Or 2&)
            End If
            If ((mask And 4&) = 0&) And (wrcache(i).addr = addr2) Then
                b2 = (wrcache(i).value And &HFF&)
                mask = (mask Or 4&)
            End If
            If ((mask And 8&) = 0&) And (wrcache(i).addr = addr3) Then
                b3 = (wrcache(i).value And &HFF&)
                mask = (mask Or 8&)
            End If
            If mask = 15& Then Exit For
        Next i
    Else
        bucket = wrcache_getBucket(addr0)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr0 Then
                    b0 = (wrcache(i).value And &HFF&)
                    mask = 1&
                    Exit Do
                End If
                i = wrcache_next(i)
            Loop
        End If

        bucket = wrcache_getBucket(addr1)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr1 Then
                    b1 = (wrcache(i).value And &HFF&)
                    mask = (mask Or 2&)
                    Exit Do
                End If
                i = wrcache_next(i)
            Loop
        End If

        bucket = wrcache_getBucket(addr2)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr2 Then
                    b2 = (wrcache(i).value And &HFF&)
                    mask = (mask Or 4&)
                    Exit Do
                End If
                i = wrcache_next(i)
            Loop
        End If

        bucket = wrcache_getBucket(addr3)
        If wrcache_bucketStamp(bucket) = wrcache_epoch Then
            i = wrcache_bucketHead(bucket)
            Do While i <> -1&
                If wrcache(i).addr = addr3 Then
                    b3 = (wrcache(i).value And &HFF&)
                    mask = (mask Or 8&)
                    Exit Do
                End If
                i = wrcache_next(i)
            Loop
        End If
    End If

    wrcache_readl = mask
End Function

Public Function memory_paging_enabled(ByRef cpu As CPU_t) As Long
    If (cpu.protected_mode <> 0&) And ((cpu.cr(0&) And &H80000000) <> 0&) Then
        memory_paging_enabled = 1&
    Else
        memory_paging_enabled = 0&
    End If
End Function

Private Function mem_access_is_write(ByVal access As Long) As Long
    mem_access_is_write = (access And 1&)
End Function

Private Function mem_access_is_user(ByVal access As Long) As Long
    If access <= MEM_ACCESS_USER_WRITE Then
        mem_access_is_user = 1&
    Else
        mem_access_is_user = 0&
    End If
End Function

Private Function cpu_default_access(ByRef cpu As CPU_t, ByVal iswrite As Long) As Long
    If cpu.startcpl = 3& Then
        If iswrite <> 0& Then
            cpu_default_access = MEM_ACCESS_USER_WRITE
        Else
            cpu_default_access = MEM_ACCESS_USER_READ
        End If
        Exit Function
    End If

    If iswrite <> 0& Then
        cpu_default_access = MEM_ACCESS_SUPERVISOR_WRITE
    Else
        cpu_default_access = MEM_ACCESS_SUPERVISOR_READ
    End If
End Function

Private Function cpu_apply_a20(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    If cpu.a20_gate = 0& Then
        cpu_apply_a20 = (addr32 And &HFFFFF&)
    Else
        cpu_apply_a20 = addr32
    End If
End Function

Private Function cpu_read_phys_u8(ByRef cpu As CPU_t, ByVal addr32 As Long) As Byte
    Dim map As Byte
    Dim cacheval As Byte

    addr32 = cpu_apply_a20(cpu, addr32)

    If wrcache_read(addr32, cacheval) <> 0& Then
        cpu_read_phys_u8 = cacheval
        Exit Function
    End If

    map = getmap(addr32)
    If map = &HFF& Then
        cpu_read_phys_u8 = &HFF&
        Exit Function
    End If

    If mem_hasRead(map) <> 0& Then
        cpu_read_phys_u8 = Memory_ReadMappedByte(map, addr32)
        Exit Function
    End If

    If mem_readcb(map) <> MEMORY_CB_NONE Then
        cpu_read_phys_u8 = Memory_DispatchReadCB(mem_readcb(map), mem_udata(map), addr32)
        Exit Function
    End If

    cpu_read_phys_u8 = &HFF&
End Function

Public Function cpu_read_linear(ByRef cpu As CPU_t, ByVal addr32 As Long) As Byte
    cpu_read_linear = cpu_read_phys_u8(cpu, addr32)
End Function

Private Sub cpu_write_phys_u8_direct(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Byte)
    Dim map As Byte

    addr32 = cpu_apply_a20(cpu, addr32)
    map = getmap(addr32)
    If map = &HFF& Then Exit Sub

    If mem_hasWrite(map) <> 0& Then
        Memory_WriteMappedByte map, addr32, value
        Exit Sub
    End If

    If mem_writecb(map) <> MEMORY_CB_NONE Then
        Memory_DispatchWriteCB mem_writecb(map), mem_udata(map), addr32, value
    End If
End Sub

Private Function cpu_read_phys_backing_u8(ByVal addr32 As Long) As Byte
    Dim map As Byte

    map = getmap(addr32)
    If map = &HFF& Then
        cpu_read_phys_backing_u8 = &HFF&
        Exit Function
    End If

    If mem_hasRead(map) <> 0& Then
        cpu_read_phys_backing_u8 = Memory_ReadMappedByte(map, addr32)
        Exit Function
    End If

    If mem_readcb(map) <> MEMORY_CB_NONE Then
        cpu_read_phys_backing_u8 = Memory_DispatchReadCB(mem_readcb(map), mem_udata(map), addr32)
        Exit Function
    End If

    cpu_read_phys_backing_u8 = &HFF&
End Function

Private Function cpu_read_phys_u16(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    Dim addr1 As Long
    Dim phys0 As Long
    Dim phys1 As Long
    Dim map As Byte
    Dim idx0 As Long
    Dim idx1 As Long
    Dim mask As Long
    Dim b0 As Long
    Dim b1 As Long

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    phys0 = cpu_apply_a20(cpu, addr32)
    phys1 = cpu_apply_a20(cpu, addr1)

    mask = wrcache_readw(phys0, phys1, b0, b1)
    If mask <> 3& Then
        map = getmap(phys0)
        If (map <> &HFF&) And (mem_hasRead(map) <> 0&) And (mem_readOffset(map) >= 0&) Then
            idx0 = phys0 - mem_start(map)
            idx1 = phys1 - mem_start(map)
            If (idx0 >= 0&) And (idx0 < mem_size(map)) And (idx1 >= 0&) And (idx1 < mem_size(map)) Then
                If (mask And 1&) = 0& Then b0 = (mem_pool(mem_readOffset(map) + idx0) And &HFF&)
                If (mask And 2&) = 0& Then b1 = (mem_pool(mem_readOffset(map) + idx1) And &HFF&)
                cpu_read_phys_u16 = (b0 Or (b1 * &H100&))
                Exit Function
            End If
        End If

        If (mask And 1&) = 0& Then b0 = (cpu_read_phys_backing_u8(phys0) And &HFF&)
        If (mask And 2&) = 0& Then b1 = (cpu_read_phys_backing_u8(phys1) And &HFF&)
    End If

    cpu_read_phys_u16 = (b0 Or (b1 * &H100&))
End Function

Private Function cpu_read_phys_u32(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim phys0 As Long
    Dim phys1 As Long
    Dim phys2 As Long
    Dim phys3 As Long
    Dim map As Byte
    Dim idx0 As Long
    Dim idx1 As Long
    Dim idx2 As Long
    Dim idx3 As Long
    Dim mask As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If

    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    phys0 = cpu_apply_a20(cpu, addr32)
    phys1 = cpu_apply_a20(cpu, addr1)
    phys2 = cpu_apply_a20(cpu, addr2)
    phys3 = cpu_apply_a20(cpu, addr3)

    mask = wrcache_readl(phys0, phys1, phys2, phys3, b0, b1, b2, b3)
    If mask <> 15& Then
        map = getmap(phys0)
        If (map <> &HFF&) And (mem_hasRead(map) <> 0&) And (mem_readOffset(map) >= 0&) Then
            idx0 = phys0 - mem_start(map)
            idx1 = phys1 - mem_start(map)
            idx2 = phys2 - mem_start(map)
            idx3 = phys3 - mem_start(map)
            If (idx0 >= 0&) And (idx0 < mem_size(map)) And _
               (idx1 >= 0&) And (idx1 < mem_size(map)) And _
               (idx2 >= 0&) And (idx2 < mem_size(map)) And _
               (idx3 >= 0&) And (idx3 < mem_size(map)) Then
                If (mask And 1&) = 0& Then b0 = (mem_pool(mem_readOffset(map) + idx0) And &HFF&)
                If (mask And 2&) = 0& Then b1 = (mem_pool(mem_readOffset(map) + idx1) And &HFF&)
                If (mask And 4&) = 0& Then b2 = (mem_pool(mem_readOffset(map) + idx2) And &HFF&)
                If (mask And 8&) = 0& Then b3 = (mem_pool(mem_readOffset(map) + idx3) And &HFF&)
                GoTo BuildResult
            End If
        End If

        If (mask And 1&) = 0& Then b0 = (cpu_read_phys_backing_u8(phys0) And &HFF&)
        If (mask And 2&) = 0& Then b1 = (cpu_read_phys_backing_u8(phys1) And &HFF&)
        If (mask And 4&) = 0& Then b2 = (cpu_read_phys_backing_u8(phys2) And &HFF&)
        If (mask And 8&) = 0& Then b3 = (cpu_read_phys_backing_u8(phys3) And &HFF&)
    End If

BuildResult:
    cpu_read_phys_u32 = (b0 Or (b1 * &H100&) Or (b2 * &H10000))
    If (b3 And &H80&) <> 0& Then
        cpu_read_phys_u32 = (cpu_read_phys_u32 Or ((b3 - &H100&) * &H1000000))
    Else
        cpu_read_phys_u32 = (cpu_read_phys_u32 Or (b3 * &H1000000))
    End If
End Function

Private Sub cpu_or_phys_u8_direct(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal mask As Byte)
    Dim value As Byte

    value = cpu_read_phys_u8(cpu, addr32)
    If (value And mask) = mask Then Exit Sub
    cpu_write_phys_u8_direct cpu, addr32, CByte((value Or mask) And &HFF&)
End Sub

Private Function page_access_is_allowed(ByRef cpu As CPU_t, ByVal access As Long, ByVal effective_user As Long, ByVal effective_write As Long) As Long
    If (mem_access_is_user(access) <> 0&) And (effective_user = 0&) Then
        page_access_is_allowed = 0&
        Exit Function
    End If

    If (mem_access_is_write(access) <> 0&) And (effective_write = 0&) Then
        If (mem_access_is_user(access) <> 0&) Or ((cpu.cr(0&) And &H10000&) <> 0&) Then
            page_access_is_allowed = 0&
            Exit Function
        End If
    End If

    page_access_is_allowed = 1&
End Function

Private Function cpu_raise_page_fault(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal access As Long, ByVal present As Long) As Long
    Dim errCode As Long

    errCode = 0&
    If present <> 0& Then errCode = (errCode Or 1&)
    If mem_access_is_write(access) <> 0& Then errCode = (errCode Or 2&)
    If mem_access_is_user(access) <> 0& Then errCode = (errCode Or 4&)

    If cpu.doexception = 0& Then
        cpu.cr(2&) = addr32
    End If

    Memory_SetException cpu, 14&, errCode
    cpu_raise_page_fault = -1&
End Function

Public Sub memory_tlb_flush(ByRef cpu As CPU_t)
    Dim setIdx As Long
    Dim wayIdx As Long

    For setIdx = 0& To 255&
        cpu.tlb(setIdx).mru = 0&
        For wayIdx = 0& To 1&
            cpu.tlb(setIdx).way(wayIdx).tag = 0&
            cpu.tlb(setIdx).way(wayIdx).phys_base = 0&
            cpu.tlb(setIdx).way(wayIdx).pte_addr = 0&
            cpu.tlb(setIdx).way(wayIdx).flags = 0&
        Next wayIdx
    Next setIdx
End Sub

Public Sub memory_tlb_invalidate_page(ByRef cpu As CPU_t, ByVal linear_addr As Long)
    Dim linear_page As Long
    Dim setIdx As Long
    Dim i As Long

    linear_page = (((linear_addr And &HFFFFF000) \ &H1000&) And &HFFFFF&)
    setIdx = (linear_page And &HFF&)

    For i = 0& To 1&
        If ((cpu.tlb(setIdx).way(i).flags And CPU_TLB_ENTRY_VALID) <> 0&) And _
           (cpu.tlb(setIdx).way(i).tag = linear_page) Then
            cpu.tlb(setIdx).way(i).flags = 0&
        End If
    Next i
End Sub

Private Function page_walk_translate(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal access As Long, ByRef outWalk As PAGE_WALK_RESULT_t) As Long
    Dim dirIdx As Long
    Dim tableIdx As Long
    Dim dentry_addr As Long
    Dim dentry As Long
    Dim tentry_addr As Long
    Dim tentry As Long
    Dim effective_user As Long
    Dim effective_write As Long
    Dim entryFlags As Long

    dirIdx = (((addr32 And &HFFC00000) \ &H400000) And &H3FF&)
    tableIdx = (((addr32 And &H3FF000) \ &H1000&) And &H3FF&)

    dentry_addr = ((cpu.cr(3&) And &HFFFFF000) Or ((dirIdx And &H3FF&) * 4&))
    dentry = cpu_read_phys_u32(cpu, dentry_addr)
    If (dentry And 1&) = 0& Then
        Call cpu_raise_page_fault(cpu, addr32, access, 0&)
        page_walk_translate = 0&
        Exit Function
    End If

    cpu_or_phys_u8_direct cpu, dentry_addr, &H20&

    tentry_addr = ((dentry And &HFFFFF000) Or ((tableIdx And &H3FF&) * 4&))
    tentry = cpu_read_phys_u32(cpu, tentry_addr)
    If (tentry And 1&) = 0& Then
        Call cpu_raise_page_fault(cpu, addr32, access, 0&)
        page_walk_translate = 0&
        Exit Function
    End If

    cpu_or_phys_u8_direct cpu, tentry_addr, &H20&

    If ((dentry And &H4&) <> 0&) And ((tentry And &H4&) <> 0&) Then
        effective_user = 1&
    Else
        effective_user = 0&
    End If

    If ((dentry And &H2&) <> 0&) And ((tentry And &H2&) <> 0&) Then
        effective_write = 1&
    Else
        effective_write = 0&
    End If

    If page_access_is_allowed(cpu, access, effective_user, effective_write) = 0& Then
        Call cpu_raise_page_fault(cpu, addr32, access, 1&)
        page_walk_translate = 0&
        Exit Function
    End If

    entryFlags = 0&
    If effective_user <> 0& Then entryFlags = (entryFlags Or CPU_TLB_ENTRY_USER_OK)
    If effective_write <> 0& Then entryFlags = (entryFlags Or CPU_TLB_ENTRY_WRITE_OK)
    If (tentry And &H40&) <> 0& Then entryFlags = (entryFlags Or CPU_TLB_ENTRY_DIRTY)

    If (mem_access_is_write(access) <> 0&) And ((entryFlags And CPU_TLB_ENTRY_DIRTY) = 0&) Then
        cpu_or_phys_u8_direct cpu, tentry_addr, &H40&
        entryFlags = (entryFlags Or CPU_TLB_ENTRY_DIRTY)
    End If

    outWalk.phys_base = (tentry And &HFFFFF000)
    outWalk.pte_addr = tentry_addr
    outWalk.entry_flags = CByte(entryFlags And &HFF&)
    page_walk_translate = 1&
End Function

Private Function translate_page(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal access As Long) As Long
    Dim linear_page As Long
    Dim setIdx As Long
    Dim idx As Long
    Dim other As Long
    Dim walk As PAGE_WALK_RESULT_t
    Dim entryFlags As Long
    Dim isWrite As Long
    Dim isUser As Long

    linear_page = (((addr32 And &HFFFFF000) \ &H1000&) And &HFFFFF&)
    setIdx = (linear_page And &HFF&)
    idx = (cpu.tlb(setIdx).mru And 1&)
    isWrite = (access And 1&)
    If access <= MEM_ACCESS_USER_WRITE Then
        isUser = 1&
    Else
        isUser = 0&
    End If

    If ((cpu.tlb(setIdx).way(idx).flags And CPU_TLB_ENTRY_VALID) <> 0&) And _
       (cpu.tlb(setIdx).way(idx).tag = linear_page) Then
        GoTo tlb_hit
    End If

    other = (idx Xor 1&)
    If ((cpu.tlb(setIdx).way(other).flags And CPU_TLB_ENTRY_VALID) <> 0&) And _
       (cpu.tlb(setIdx).way(other).tag = linear_page) Then
        idx = other
        cpu.tlb(setIdx).mru = CByte(idx And 1&)
        GoTo tlb_hit
    End If

    If page_walk_translate(cpu, addr32, access, walk) = 0& Then
        translate_page = -1&
        Exit Function
    End If

    If (cpu.tlb(setIdx).way(0&).flags And CPU_TLB_ENTRY_VALID) = 0& Then
        idx = 0&
    ElseIf (cpu.tlb(setIdx).way(1&).flags And CPU_TLB_ENTRY_VALID) = 0& Then
        idx = 1&
    Else
        idx = ((cpu.tlb(setIdx).mru Xor 1&) And 1&)
    End If

    cpu.tlb(setIdx).way(idx).tag = linear_page
    cpu.tlb(setIdx).way(idx).phys_base = walk.phys_base
    cpu.tlb(setIdx).way(idx).pte_addr = walk.pte_addr
    cpu.tlb(setIdx).way(idx).flags = CByte((walk.entry_flags Or CPU_TLB_ENTRY_VALID) And &HFF&)
    cpu.tlb(setIdx).mru = CByte(idx And 1&)

    translate_page = ((cpu.tlb(setIdx).way(idx).phys_base And &HFFFFF000) Or (addr32 And &HFFF&))
    Exit Function

tlb_hit:
    entryFlags = cpu.tlb(setIdx).way(idx).flags

    If isUser <> 0& Then
        If (entryFlags And CPU_TLB_ENTRY_USER_OK) = 0& Then
            translate_page = cpu_raise_page_fault(cpu, addr32, access, 1&)
            Exit Function
        End If
        If (isWrite <> 0&) And ((entryFlags And CPU_TLB_ENTRY_WRITE_OK) = 0&) Then
            translate_page = cpu_raise_page_fault(cpu, addr32, access, 1&)
            Exit Function
        End If
    ElseIf (isWrite <> 0&) And ((entryFlags And CPU_TLB_ENTRY_WRITE_OK) = 0&) Then
        If (cpu.cr(0&) And &H10000&) <> 0& Then
            translate_page = cpu_raise_page_fault(cpu, addr32, access, 1&)
            Exit Function
        End If
    End If

    If (isWrite <> 0&) And ((entryFlags And CPU_TLB_ENTRY_DIRTY) = 0&) Then
        cpu_or_phys_u8_direct cpu, cpu.tlb(setIdx).way(idx).pte_addr, &H40&
        entryFlags = (entryFlags Or CPU_TLB_ENTRY_DIRTY)
        cpu.tlb(setIdx).way(idx).flags = CByte(entryFlags And &HFF&)
    End If

    cpu.tlb(setIdx).mru = CByte(idx And 1&)
    translate_page = ((cpu.tlb(setIdx).way(idx).phys_base And &HFFFFF000) Or (addr32 And &HFFF&))
End Function

Private Function cpu_read_access_cached(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal access As Long, ByVal pagingEnabled As Long) As Byte
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_read_access_cached = &HFF&
        Exit Function
    End If

    If pagingEnabled <> 0& Then
        addr32 = translate_page(cpu, addr32, access)
        If addr32 = -1& Then
            cpu_read_access_cached = &HFF&
            Exit Function
        End If
    End If

    cpu_read_access_cached = cpu_read_linear(cpu, addr32)
End Function

Private Function cpu_read_access(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal access As Long) As Byte
    cpu_read_access = cpu_read_access_cached(cpu, addr32, access, memory_paging_enabled(cpu))
End Function

Private Sub cpu_write_access_cached(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Byte, ByVal access As Long, ByVal pagingEnabled As Long)
    If cpu.nowrite <> 0& Then Exit Sub

    If pagingEnabled <> 0& Then
        addr32 = translate_page(cpu, addr32, access)
        If addr32 = -1& Then Exit Sub
    End If

    cpu_write_linear cpu, addr32, value
End Sub

Private Sub cpu_write_access(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Byte, ByVal access As Long)
    cpu_write_access_cached cpu, addr32, value, access, memory_paging_enabled(cpu)
End Sub

Public Sub cpu_write_linear(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Byte)
    If cpu.nowrite <> 0& Then Exit Sub

    addr32 = cpu_apply_a20(cpu, addr32)

    wrcache_write addr32, value
End Sub

Public Sub cpu_write(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Byte)
    cpu_write_access cpu, addr32, value, cpu_default_access(cpu, 1&)
End Sub

Public Function cpu_read(ByRef cpu As CPU_t, ByVal addr32 As Long) As Byte
    cpu_read = cpu_read_access(cpu, addr32, cpu_default_access(cpu, 0&))
End Function

Public Sub cpu_write_sys(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Byte)
    cpu_write_access cpu, addr32, value, MEM_ACCESS_SUPERVISOR_WRITE
End Sub

Public Function cpu_read_sys(ByRef cpu As CPU_t, ByVal addr32 As Long) As Byte
    cpu_read_sys = cpu_read_access(cpu, addr32, MEM_ACCESS_SUPERVISOR_READ)
End Function

Public Sub cpu_writew(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Long)
    Dim access As Long
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim v1 As Byte

    If cpu.nowrite <> 0& Then Exit Sub

    access = cpu_default_access(cpu, 1&)
    pagingEnabled = memory_paging_enabled(cpu)
    v1 = CByte((value And &HFF00&) \ &H100&)

    If pagingEnabled = 0& Then
        cpu_writew_linear cpu, addr32, value
        Exit Sub
    End If

    If (addr32 And &HFFF&) <> &HFFF& Then
        phys = translate_page(cpu, addr32, access)
        If phys = -1& Then Exit Sub
        cpu_writew_linear cpu, phys, value
        Exit Sub
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    cpu_write_access_cached cpu, addr32, CByte(value And &HFF&), access, pagingEnabled
    cpu_write_access_cached cpu, addr1, v1, access, pagingEnabled
End Sub

Public Sub cpu_writel(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Long)
    Dim access As Long
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim v1 As Byte
    Dim v2 As Byte
    Dim v3 As Byte

    If cpu.nowrite <> 0& Then Exit Sub

    access = cpu_default_access(cpu, 1&)
    pagingEnabled = memory_paging_enabled(cpu)
    v1 = CByte((value And &HFF00&) \ &H100&)
    v2 = CByte((value And &HFF0000) \ &H10000)
    If (value And &H80000000) <> 0& Then
        v3 = CByte((((value And &H7F000000) \ &H1000000) Or &H80&) And &HFF&)
    Else
        v3 = CByte((value And &H7F000000) \ &H1000000)
    End If

    If pagingEnabled = 0& Then
        cpu_writel_linear cpu, addr32, value
        Exit Sub
    End If

    If (addr32 And &HFFF&) <= &HFFC& Then
        phys = translate_page(cpu, addr32, access)
        If phys = -1& Then Exit Sub
        cpu_writel_linear cpu, phys, value
        Exit Sub
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If
    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If
    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    cpu_write_access_cached cpu, addr32, CByte(value And &HFF&), access, pagingEnabled
    cpu_write_access_cached cpu, addr1, v1, access, pagingEnabled
    cpu_write_access_cached cpu, addr2, v2, access, pagingEnabled
    cpu_write_access_cached cpu, addr3, v3, access, pagingEnabled
End Sub

Public Sub cpu_writew_sys(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Long)
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim v1 As Byte

    If cpu.nowrite <> 0& Then Exit Sub

    pagingEnabled = memory_paging_enabled(cpu)
    v1 = CByte((value And &HFF00&) \ &H100&)

    If pagingEnabled = 0& Then
        cpu_writew_linear cpu, addr32, value
        Exit Sub
    End If

    If (addr32 And &HFFF&) <> &HFFF& Then
        phys = translate_page(cpu, addr32, MEM_ACCESS_SUPERVISOR_WRITE)
        If phys = -1& Then Exit Sub
        cpu_writew_linear cpu, phys, value
        Exit Sub
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    cpu_write_access_cached cpu, addr32, CByte(value And &HFF&), MEM_ACCESS_SUPERVISOR_WRITE, pagingEnabled
    cpu_write_access_cached cpu, addr1, v1, MEM_ACCESS_SUPERVISOR_WRITE, pagingEnabled
End Sub

Public Sub cpu_writel_sys(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Long)
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim v1 As Byte
    Dim v2 As Byte
    Dim v3 As Byte

    If cpu.nowrite <> 0& Then Exit Sub

    pagingEnabled = memory_paging_enabled(cpu)
    v1 = CByte((value And &HFF00&) \ &H100&)
    v2 = CByte((value And &HFF0000) \ &H10000)
    If (value And &H80000000) <> 0& Then
        v3 = CByte((((value And &H7F000000) \ &H1000000) Or &H80&) And &HFF&)
    Else
        v3 = CByte((value And &H7F000000) \ &H1000000)
    End If

    If pagingEnabled = 0& Then
        cpu_writel_linear cpu, addr32, value
        Exit Sub
    End If

    If (addr32 And &HFFF&) <= &HFFC& Then
        phys = translate_page(cpu, addr32, MEM_ACCESS_SUPERVISOR_WRITE)
        If phys = -1& Then Exit Sub
        cpu_writel_linear cpu, phys, value
        Exit Sub
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If
    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If
    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    cpu_write_access_cached cpu, addr32, CByte(value And &HFF&), MEM_ACCESS_SUPERVISOR_WRITE, pagingEnabled
    cpu_write_access_cached cpu, addr1, v1, MEM_ACCESS_SUPERVISOR_WRITE, pagingEnabled
    cpu_write_access_cached cpu, addr2, v2, MEM_ACCESS_SUPERVISOR_WRITE, pagingEnabled
    cpu_write_access_cached cpu, addr3, v3, MEM_ACCESS_SUPERVISOR_WRITE, pagingEnabled
End Sub

Public Sub cpu_writew_linear(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Long)
    Dim addr1 As Long
    Dim b0 As Byte
    Dim b1 As Byte

    If cpu.nowrite <> 0& Then Exit Sub

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    b0 = CByte(value And &HFF&)
    b1 = CByte((value And &HFF00&) \ &H100&)
    wrcache_writew cpu_apply_a20(cpu, addr32), cpu_apply_a20(cpu, addr1), b0, b1
End Sub

Public Sub cpu_writel_linear(ByRef cpu As CPU_t, ByVal addr32 As Long, ByVal value As Long)
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim v3 As Byte

    If cpu.nowrite <> 0& Then Exit Sub

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If
    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If
    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    If (value And &H80000000) <> 0& Then
        v3 = CByte((((value And &H7F000000) \ &H1000000) Or &H80&) And &HFF&)
    Else
        v3 = CByte((value And &H7F000000) \ &H1000000)
    End If

    wrcache_writel cpu_apply_a20(cpu, addr32), cpu_apply_a20(cpu, addr1), _
                   cpu_apply_a20(cpu, addr2), cpu_apply_a20(cpu, addr3), _
                   CByte(value And &HFF&), CByte((value And &HFF00&) \ &H100&), _
                   CByte((value And &HFF0000) \ &H10000), v3
End Sub

Public Function cpu_readw(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    Dim access As Long
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim value As Long

    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readw = &HFF&
        Exit Function
    End If

    access = cpu_default_access(cpu, 0&)
    pagingEnabled = memory_paging_enabled(cpu)

    If pagingEnabled = 0& Then
        cpu_readw = cpu_read_phys_u16(cpu, addr32)
        Exit Function
    End If

    If (addr32 And &HFFF&) <> &HFFF& Then
        phys = translate_page(cpu, addr32, access)
        If phys = -1& Then
            cpu_readw = &HFF&
            Exit Function
        End If
        cpu_readw = cpu_read_phys_u16(cpu, phys)
        Exit Function
    End If

    value = (cpu_read_access_cached(cpu, addr32, access, pagingEnabled) And &HFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readw = value
        Exit Function
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    cpu_readw = (value Or (CLng(cpu_read_access_cached(cpu, addr1, access, pagingEnabled) And &HFF&) * &H100&))
End Function

Public Function cpu_readl(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    Dim access As Long
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim b3 As Long
    Dim value As Long

    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl = &HFF&
        Exit Function
    End If

    access = cpu_default_access(cpu, 0&)
    pagingEnabled = memory_paging_enabled(cpu)

    If pagingEnabled = 0& Then
        cpu_readl = cpu_read_phys_u32(cpu, addr32)
        Exit Function
    End If

    If (addr32 And &HFFF&) <= &HFFC& Then
        phys = translate_page(cpu, addr32, access)
        If phys = -1& Then
            cpu_readl = &HFF&
            Exit Function
        End If
        cpu_readl = cpu_read_phys_u32(cpu, phys)
        Exit Function
    End If

    value = (cpu_read_access_cached(cpu, addr32, access, pagingEnabled) And &HFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl = value
        Exit Function
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If
    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If
    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    value = (value Or (CLng(cpu_read_access_cached(cpu, addr1, access, pagingEnabled) And &HFF&) * &H100&))
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl = value
        Exit Function
    End If
    value = (value Or (CLng(cpu_read_access_cached(cpu, addr2, access, pagingEnabled) And &HFF&) * &H10000))
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl = value
        Exit Function
    End If
    b3 = (cpu_read_access_cached(cpu, addr3, access, pagingEnabled) And &HFF&)
    If (b3 And &H80&) <> 0& Then
        value = (value Or ((b3 - &H100&) * &H1000000))
    Else
        value = (value Or (b3 * &H1000000))
    End If

    cpu_readl = value
End Function

Public Function cpu_readw_sys(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim value As Long

    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readw_sys = &HFF&
        Exit Function
    End If

    pagingEnabled = memory_paging_enabled(cpu)

    If pagingEnabled = 0& Then
        cpu_readw_sys = cpu_read_phys_u16(cpu, addr32)
        Exit Function
    End If

    If (addr32 And &HFFF&) <> &HFFF& Then
        phys = translate_page(cpu, addr32, MEM_ACCESS_SUPERVISOR_READ)
        If phys = -1& Then
            cpu_readw_sys = &HFF&
            Exit Function
        End If
        cpu_readw_sys = cpu_read_phys_u16(cpu, phys)
        Exit Function
    End If

    value = (cpu_read_access_cached(cpu, addr32, MEM_ACCESS_SUPERVISOR_READ, pagingEnabled) And &HFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readw_sys = value
        Exit Function
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If

    cpu_readw_sys = (value Or (CLng(cpu_read_access_cached(cpu, addr1, MEM_ACCESS_SUPERVISOR_READ, pagingEnabled) And &HFF&) * &H100&))
End Function

Public Function cpu_readl_sys(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    Dim pagingEnabled As Long
    Dim phys As Long
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim b3 As Long
    Dim value As Long

    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl_sys = &HFF&
        Exit Function
    End If

    pagingEnabled = memory_paging_enabled(cpu)

    If pagingEnabled = 0& Then
        cpu_readl_sys = cpu_read_phys_u32(cpu, addr32)
        Exit Function
    End If

    If (addr32 And &HFFF&) <= &HFFC& Then
        phys = translate_page(cpu, addr32, MEM_ACCESS_SUPERVISOR_READ)
        If phys = -1& Then
            cpu_readl_sys = &HFF&
            Exit Function
        End If
        cpu_readl_sys = cpu_read_phys_u32(cpu, phys)
        Exit Function
    End If

    value = (cpu_read_access_cached(cpu, addr32, MEM_ACCESS_SUPERVISOR_READ, pagingEnabled) And &HFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl_sys = value
        Exit Function
    End If

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If
    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If
    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    value = (value Or (CLng(cpu_read_access_cached(cpu, addr1, MEM_ACCESS_SUPERVISOR_READ, pagingEnabled) And &HFF&) * &H100&))
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl_sys = value
        Exit Function
    End If
    value = (value Or (CLng(cpu_read_access_cached(cpu, addr2, MEM_ACCESS_SUPERVISOR_READ, pagingEnabled) And &HFF&) * &H10000))
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_readl_sys = value
        Exit Function
    End If
    b3 = (cpu_read_access_cached(cpu, addr3, MEM_ACCESS_SUPERVISOR_READ, pagingEnabled) And &HFF&)
    If (b3 And &H80&) <> 0& Then
        value = (value Or ((b3 - &H100&) * &H1000000))
    Else
        value = (value Or (b3 * &H1000000))
    End If

    cpu_readl_sys = value
End Function

Public Function cpu_readw_linear(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    cpu_readw_linear = cpu_read_phys_u16(cpu, addr32)
End Function

Public Function cpu_readl_linear(ByRef cpu As CPU_t, ByVal addr32 As Long) As Long
    cpu_readl_linear = cpu_read_phys_u32(cpu, addr32)
End Function

Public Sub memory_mapRegister(ByVal start As Long, ByVal length As Long, ByRef readb As Variant, ByRef writeb As Variant)
    Dim i As Long
    Dim j As Long
    Dim pageStart As Long
    Dim pageCount As Long

    For i = 0& To 31&
        If mem_used(i) = 0& Then Exit For
    Next i

    If i = 32& Then
        debug_log DEBUG_ERROR, "[MEMORY] Out of memory map structs!"
        End
    End If

    mem_start(i) = start
    mem_size(i) = length
    mem_readcb(i) = MEMORY_CB_NONE
    mem_writecb(i) = MEMORY_CB_NONE
    mem_udata(i) = 0&
    mem_used(i) = 1&
    mem_readOffset(i) = -1&
    mem_writeOffset(i) = -1&

    mem_hasRead(i) = IIf(IsArray(readb), 1&, 0&)
    If mem_hasRead(i) <> 0& Then
        If Memory_CopyArrayToPool(readb, length, mem_readOffset(i)) <> 0& Then
            debug_log DEBUG_ERROR, "[MEMORY] Failed to map read buffer"
            End
        End If
    End If

    mem_hasWrite(i) = IIf(IsArray(writeb), 1&, 0&)
    If mem_hasWrite(i) <> 0& Then
        If mem_hasRead(i) <> 0& Then
            mem_writeOffset(i) = mem_readOffset(i)
        Else
            If Memory_CopyArrayToPool(writeb, length, mem_writeOffset(i)) <> 0& Then
                debug_log DEBUG_ERROR, "[MEMORY] Failed to map write buffer"
                End
            End If
        End If
    End If

    pageStart = (((start And &HFFFFF000) \ &H1000&) And &HFFFFF&)
    pageCount = (((length And &HFFFFF000) \ &H1000&) And &HFFFFF&)
    For j = pageStart To (pageStart + pageCount - 1&)
        mem_map_lookup(j) = CByte(i)
    Next j
End Sub

Public Sub memory_mapCallbackRegister(ByVal start As Long, ByVal count As Long, ByVal readb As Long, ByVal writeb As Long, ByVal udata As Long)
    Dim i As Long
    Dim j As Long
    Dim pageStart As Long
    Dim pageCount As Long

    For i = 0& To 31&
        If mem_used(i) = 0& Then Exit For
    Next i

    If i = 32& Then
        debug_log DEBUG_ERROR, "[MEMORY] Out of memory map structs!"
        End
    End If

    mem_start(i) = start
    mem_size(i) = count
    mem_readcb(i) = readb
    mem_writecb(i) = writeb
    mem_udata(i) = udata
    mem_used(i) = 1&
    mem_hasRead(i) = 0&
    mem_hasWrite(i) = 0&
    mem_readOffset(i) = -1&
    mem_writeOffset(i) = -1&

    pageStart = (((start And &HFFFFF000) \ &H1000&) And &HFFFFF&)
    pageCount = (((count And &HFFFFF000) \ &H1000&) And &HFFFFF&)
    For j = pageStart To (pageStart + pageCount - 1&)
        mem_map_lookup(j) = CByte(i)
    Next j
End Sub

Public Function memory_init() As Long
    Dim i As Long

    mem_poolSize = 0&
    ReDim mem_pool(0& To 0&) As Byte

    For i = 0& To 31&
        mem_start(i) = 0&
        mem_size(i) = 0&
        mem_hasRead(i) = 0&
        mem_hasWrite(i) = 0&
        mem_readcb(i) = MEMORY_CB_NONE
        mem_writecb(i) = MEMORY_CB_NONE
        mem_udata(i) = 0&
        mem_used(i) = 0&
        mem_readOffset(i) = -1&
        mem_writeOffset(i) = -1&
    Next i

    For i = 0& To UBound(mem_map_lookup)
        mem_map_lookup(i) = &HFF&
    Next i

    memory_init = 0&
End Function

Private Function Memory_ReadMappedByte(ByVal map As Byte, ByVal addr32 As Long) As Byte
    Dim idx As Long

    If (mem_hasRead(map) = 0&) Or (mem_readOffset(map) < 0&) Then
        Memory_ReadMappedByte = &HFF&
        Exit Function
    End If

    idx = addr32 - mem_start(map)
    If (idx < 0&) Or (idx >= mem_size(map)) Then
        Memory_ReadMappedByte = &HFF&
        Exit Function
    End If

    Memory_ReadMappedByte = mem_pool(mem_readOffset(map) + idx)
End Function

Private Sub Memory_WriteMappedByte(ByVal map As Byte, ByVal addr32 As Long, ByVal value As Byte)
    Dim idx As Long

    idx = addr32 - mem_start(map)
    If (idx < 0&) Or (idx >= mem_size(map)) Then Exit Sub

    If (mem_hasWrite(map) <> 0&) And (mem_writeOffset(map) >= 0&) Then
        mem_pool(mem_writeOffset(map) + idx) = value
    End If

    If (mem_hasWrite(map) <> 0&) And (mem_hasRead(map) <> 0&) And (mem_readOffset(map) >= 0&) Then
        If mem_readOffset(map) <> mem_writeOffset(map) Then
            mem_pool(mem_readOffset(map) + idx) = value
        End If
    End If
End Sub

Private Function Memory_AllocPool(ByVal length As Long) As Long
    Dim newSize As Long

    If length < 0& Then
        Memory_AllocPool = -1&
        Exit Function
    End If

    If length = 0& Then
        Memory_AllocPool = mem_poolSize
        Exit Function
    End If

    If mem_poolSize > (&H7FFFFFFF - length) Then
        Memory_AllocPool = -1&
        Exit Function
    End If

    Memory_AllocPool = mem_poolSize
    newSize = mem_poolSize + length

    If newSize > (UBound(mem_pool) + 1&) Then
        ReDim Preserve mem_pool(0& To newSize - 1&) As Byte
    End If

    mem_poolSize = newSize
End Function

Private Function Memory_CopyArrayToPool(ByRef srcVar As Variant, ByVal length As Long, ByRef poolOffset As Long) As Long
    Dim src() As Byte
    Dim srcL As Long
    Dim srcU As Long

    On Error GoTo CopyFail

    src = srcVar
    srcL = LBound(src)
    srcU = UBound(src)
    If (srcU - srcL + 1&) < length Then GoTo CopyFail

    poolOffset = Memory_AllocPool(length)
    If poolOffset < 0& Then GoTo CopyFail

    If length > 0& Then
        CopyMemory mem_pool(poolOffset), src(srcL), length
    End If

    Memory_CopyArrayToPool = 0&
    Exit Function

CopyFail:
    poolOffset = -1&
    Memory_CopyArrayToPool = -1&
End Function

Private Function Memory_ReadByteByAddr(ByVal addr32 As Long) As Byte
    Dim map As Byte

    map = getmap(addr32)
    If map = &HFF& Then
        Memory_ReadByteByAddr = &HFF&
        Exit Function
    End If

    If mem_hasRead(map) <> 0& Then
        Memory_ReadByteByAddr = Memory_ReadMappedByte(map, addr32)
    ElseIf mem_readcb(map) <> MEMORY_CB_NONE Then
        Memory_ReadByteByAddr = Memory_DispatchReadCB(mem_readcb(map), mem_udata(map), addr32)
    Else
        Memory_ReadByteByAddr = &HFF&
    End If
End Function

Private Sub Memory_WriteByteByAddr(ByVal addr32 As Long, ByVal value As Byte)
    Dim map As Byte

    map = getmap(addr32)
    If map = &HFF& Then Exit Sub

    If mem_hasWrite(map) <> 0& Then
        Memory_WriteMappedByte map, addr32, value
    ElseIf mem_writecb(map) <> MEMORY_CB_NONE Then
        Memory_DispatchWriteCB mem_writecb(map), mem_udata(map), addr32, value
    End If
End Sub

Private Function Memory_ReadDwordMapped(ByVal addr32 As Long) As Long
    Dim addr1 As Long
    Dim addr2 As Long
    Dim addr3 As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long

    addr1 = addr32
    If addr1 = &H7FFFFFFF Then
        addr1 = &H80000000
    Else
        addr1 = addr1 + 1&
    End If
    addr2 = addr1
    If addr2 = &H7FFFFFFF Then
        addr2 = &H80000000
    Else
        addr2 = addr2 + 1&
    End If
    addr3 = addr2
    If addr3 = &H7FFFFFFF Then
        addr3 = &H80000000
    Else
        addr3 = addr3 + 1&
    End If

    b0 = Memory_ReadByteByAddr(addr32)
    b1 = Memory_ReadByteByAddr(addr1)
    b2 = Memory_ReadByteByAddr(addr2)
    b3 = Memory_ReadByteByAddr(addr3)

    Memory_ReadDwordMapped = (b0 Or (b1 * &H100&) Or (b2 * &H10000))
    If (b3 And &H80&) <> 0& Then
        Memory_ReadDwordMapped = (Memory_ReadDwordMapped Or ((b3 - &H100&) * &H1000000))
    Else
        Memory_ReadDwordMapped = (Memory_ReadDwordMapped Or (b3 * &H1000000))
    End If
End Function

Private Sub Memory_SetException(ByRef cpu As CPU_t, ByVal exnum As Byte, ByVal exerr As Long)
    If cpu.doexception = 0& Then
        cpu.doexception = 1&
        cpu.exceptionval = exnum
        cpu.exceptionerr = exerr

        diag_count_exception CLng(exnum)

        If exnum = 14& Then cpu.nowrite = 1&

        If showops <> 0& Then
            debug_log DEBUG_DETAIL, "EX: " & CStr(exnum) & " (" & Hex$(exerr) & ")"
        End If
    End If
End Sub

Private Function Memory_DispatchReadCB(ByVal cbid As Long, ByVal udata As Long, ByVal addr32 As Long) As Byte
    Select Case cbid
        Case MEMORY_CB_VGA
            Memory_DispatchReadCB = vga_readmemory(udata, addr32)
        Case MEMORY_CB_BUSLOGIC_ROM
            Memory_DispatchReadCB = buslogic_readrom(udata, addr32)
        Case MEMORY_CB_ET4000
            Memory_DispatchReadCB = et4000_readmemory(udata, addr32)
        Case Else
            Memory_DispatchReadCB = &HFF&
    End Select
End Function

Private Sub Memory_DispatchWriteCB(ByVal cbid As Long, ByVal udata As Long, ByVal addr32 As Long, ByVal value As Byte)
    Select Case cbid
        Case MEMORY_CB_VGA
            vga_writememory udata, addr32, value
        Case MEMORY_CB_BUSLOGIC_ROM
            buslogic_writerom udata, addr32, value
        Case MEMORY_CB_ET4000
            et4000_writememory udata, addr32, value
        Case Else
            ' Callback backend not used.
    End Select
End Sub


