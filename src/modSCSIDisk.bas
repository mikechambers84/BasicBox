Attribute VB_Name = "modSCSIDisk"
Option Explicit

Private Function SCSIDisk_Be32(ByRef data() As Byte, ByVal offset As Long) As Long
    SCSIDisk_Be32 = U32FromDouble(CDbl(data(offset)) * 16777216# + CDbl(data(offset + 1&)) * 65536# + CDbl(data(offset + 2&)) * 256# + CDbl(data(offset + 3&)))
End Function

Private Function SCSIDisk_ReadLBA(ByRef cdb() As Byte) As Long
    Select Case cdb(0)
        Case GPCMD_READ_6, GPCMD_WRITE_6
            SCSIDisk_ReadLBA = ((CLng(cdb(1) And &H1F&) * &H10000) Or (CLng(cdb(2)) * &H100&) Or CLng(cdb(3)))
        Case GPCMD_READ_10, GPCMD_WRITE_10, GPCMD_VERIFY_10, GPCMD_READ_12, GPCMD_WRITE_12
            SCSIDisk_ReadLBA = SCSIDisk_Be32(cdb, 2&)
        Case Else
            SCSIDisk_ReadLBA = 0&
    End Select
End Function

Private Function SCSIDisk_ReadBlocks(ByRef cdb() As Byte) As Long
    Select Case cdb(0)
        Case GPCMD_READ_6, GPCMD_WRITE_6
            If cdb(4) <> 0& Then
                SCSIDisk_ReadBlocks = CLng(cdb(4))
            Else
                SCSIDisk_ReadBlocks = 256&
            End If
        Case GPCMD_READ_10, GPCMD_WRITE_10, GPCMD_VERIFY_10
            SCSIDisk_ReadBlocks = ((CLng(cdb(7)) * &H100&) Or CLng(cdb(8)))
        Case GPCMD_READ_12, GPCMD_WRITE_12
            SCSIDisk_ReadBlocks = SCSIDisk_Be32(cdb, 6&)
        Case Else
            SCSIDisk_ReadBlocks = 0&
    End Select
End Function

Private Function SCSIDisk_FileOffset(ByVal lba As Long) As U64_t
    Dim offset64 As U64_t

    offset64 = U64_FromU32(lba)
    SCSIDisk_FileOffset = U64_Shl(offset64, 9&)
End Function

Private Function SCSIDisk_FileBlocks(ByVal fileHandle As Long) As U64_t
    Dim size64 As U64_t

    If scsi_file_get_size(fileHandle, size64) <> 0& Then
        SCSIDisk_FileBlocks = U64_Zero()
        Exit Function
    End If

    SCSIDisk_FileBlocks = U64_Shr(size64, 9&)
End Function

Private Function SCSIDisk_TotalBlocksMinusOne(ByVal bus As Long, ByVal targetId As Long) As U64_t
    Dim one As U64_t

    one = U64_FromU32(1&)
    SCSIDisk_TotalBlocksMinusOne = U64_Sub(scsi_devices(bus, targetId).total_blocks, one)
End Function

Private Function SCSIDisk_LbaOutOfRange(ByVal bus As Long, ByVal targetId As Long, ByVal lba As Long, ByVal blocks As Long) As Long
    Dim lba64 As U64_t
    Dim blocks64 As U64_t
    Dim end64 As U64_t
    Dim total As U64_t

    total = scsi_devices(bus, targetId).total_blocks
    lba64 = U64_FromU32(lba)
    If (U64_Lt(total, lba64) <> 0&) Or (U64_Eq(total, lba64) <> 0&) Then
        SCSIDisk_LbaOutOfRange = 1&
        Exit Function
    End If

    blocks64 = U64_FromU32(blocks)
    end64 = U64_Add(lba64, blocks64)
    If U64_Lt(total, end64) <> 0& Then
        SCSIDisk_LbaOutOfRange = 1&
    Else
        SCSIDisk_LbaOutOfRange = 0&
    End If
End Function

Private Sub SCSIDisk_WriteAscii(ByVal offset As Long, ByVal text As String, ByVal maxLen As Long)
    Dim i As Long
    Dim ch As Long

    For i = 0& To maxLen - 1&
        If i < Len(text) Then
            ch = asc(Mid$(text, i + 1&, 1&))
            scsi_temp_buffer(offset + i) = CByte(ch And &HFF&)
        Else
            scsi_temp_buffer(offset + i) = 32&
        End If
    Next i
End Sub

Public Sub scsi_disk_request_sense(ByVal bus As Long, ByVal targetId As Long, ByRef buffer() As Byte, ByVal allocLength As Long)
    Dim copyLen As Long
    Dim i As Long

    copyLen = allocLength
    If copyLen > 18& Then copyLen = 18&
    If copyLen < 0& Then copyLen = 0&
    If copyLen = 0& Then Exit Sub

    ReDim buffer(0 To copyLen - 1&) As Byte
    For i = 0& To copyLen - 1&
        buffer(i) = scsi_devices(bus, targetId).sense(i)
    Next i
    scsi_common_clear_sense bus, targetId
End Sub

Public Sub scsi_disk_reset(ByVal bus As Long, ByVal targetId As Long)
    scsi_common_clear_sense bus, targetId
End Sub

Private Sub SCSIDisk_BuildInquiry(ByVal bus As Long, ByVal targetId As Long)
    Dim i As Long

    If scsi_common_ensure_buffer(36&) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
        scsi_devices(bus, targetId).buffer_length = 0&
        Exit Sub
    End If

    For i = 0& To 35&
        scsi_temp_buffer(i) = 0&
    Next i
    scsi_temp_buffer(0) = &H0&
    scsi_temp_buffer(1) = &H0&
    scsi_temp_buffer(2) = &H5&
    scsi_temp_buffer(3) = &H2&
    scsi_temp_buffer(4) = 31&
    SCSIDisk_WriteAscii 8&, "BasicBox", 8&
    SCSIDisk_WriteAscii 16&, "SCSI Disk      ", 16&
    SCSIDisk_WriteAscii 32&, "0001", 4&
    scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
    scsi_devices(bus, targetId).buffer_length = 36&
End Sub

Private Sub SCSIDisk_BuildModeSense(ByVal bus As Long, ByVal targetId As Long, ByRef cdb() As Byte)
    Dim pageCode As Byte
    Dim isTen As Long
    Dim totalLen As Long
    Dim i As Long

    pageCode = CByte(cdb(2) And &H3F&)
    isTen = IIf(cdb(0) = GPCMD_MODE_SENSE_10, 1&, 0&)
    totalLen = IIf(isTen <> 0&, 20&, 16&)

    If (pageCode <> &H3F&) And (pageCode <> &H8&) And (pageCode <> &H4&) Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
        scsi_devices(bus, targetId).buffer_length = 0&
        Exit Sub
    End If

    If scsi_common_ensure_buffer(totalLen) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
        scsi_devices(bus, targetId).buffer_length = 0&
        Exit Sub
    End If

    For i = 0& To totalLen - 1&
        scsi_temp_buffer(i) = 0&
    Next i

    If isTen <> 0& Then
        scsi_temp_buffer(0) = 0&
        scsi_temp_buffer(1) = CByte((totalLen - 2&) And &HFF&)
        scsi_temp_buffer(7) = 8&
        scsi_temp_buffer(8) = &H8&
        scsi_temp_buffer(9) = &HA&
        scsi_temp_buffer(10) = &H4&
        scsi_temp_buffer(16) = &H4&
        scsi_temp_buffer(17) = &H16&
        scsi_temp_buffer(18) = &H0&
        scsi_temp_buffer(19) = &H80&
    Else
        scsi_temp_buffer(0) = CByte((totalLen - 1&) And &HFF&)
        scsi_temp_buffer(3) = 8&
        scsi_temp_buffer(4) = &H8&
        scsi_temp_buffer(5) = &HA&
        scsi_temp_buffer(6) = &H4&
        scsi_temp_buffer(12) = &H4&
        scsi_temp_buffer(13) = &H16&
        scsi_temp_buffer(14) = &H0&
        scsi_temp_buffer(15) = &H80&
    End If

    scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
    scsi_devices(bus, targetId).buffer_length = totalLen
End Sub

Public Sub scsi_disk_command(ByVal bus As Long, ByVal targetId As Long, ByRef cdb() As Byte)
    Dim lba As Long
    Dim blocks As Long
    Dim dataLen As Long
    Dim i As Long
    Dim lastBlock As U64_t
    Dim fileOffset As U64_t

    scsi_common_clear_sense bus, targetId
    scsi_devices(bus, targetId).buffer_length = 0&
    scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS

    Select Case cdb(0)
        Case GPCMD_TEST_UNIT_READY
            Exit Sub

        Case GPCMD_INQUIRY
            If (cdb(1) And 1&) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            SCSIDisk_BuildInquiry bus, targetId
            Exit Sub

        Case GPCMD_REQUEST_SENSE
            If scsi_common_ensure_buffer(18&) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            For i = 0& To 17&
                scsi_temp_buffer(i) = scsi_devices(bus, targetId).sense(i)
            Next i
            scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
            scsi_devices(bus, targetId).buffer_length = 18&
            scsi_common_clear_sense bus, targetId
            Exit Sub

        Case GPCMD_MODE_SENSE_6, GPCMD_MODE_SENSE_10
            SCSIDisk_BuildModeSense bus, targetId, cdb
            Exit Sub

        Case GPCMD_MODE_SELECT_6, GPCMD_MODE_SELECT_10, GPCMD_START_STOP_UNIT, GPCMD_PREVENT_REMOVAL, GPCMD_SYNCHRONIZE_CACHE, GPCMD_VERIFY_10
            Exit Sub

        Case GPCMD_READ_CDROM_CAPACITY
            If scsi_common_ensure_buffer(8&) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            If U64_IsZero(scsi_devices(bus, targetId).total_blocks) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_NOT_READY, ASC_MEDIUM_NOT_PRESENT, ASCQ_NONE
                Exit Sub
            End If
            For i = 0& To 7&
                scsi_temp_buffer(i) = 0&
            Next i
            lastBlock = SCSIDisk_TotalBlocksMinusOne(bus, targetId)
            scsi_temp_buffer(0) = CByte(U32Shr(lastBlock.Lo, 24&) And &HFF&)
            scsi_temp_buffer(1) = CByte(U32Shr(lastBlock.Lo, 16&) And &HFF&)
            scsi_temp_buffer(2) = CByte(U32Shr(lastBlock.Lo, 8&) And &HFF&)
            scsi_temp_buffer(3) = CByte(lastBlock.Lo And &HFF&)
            scsi_temp_buffer(4) = CByte(U32Shr(scsi_devices(bus, targetId).block_len, 24&) And &HFF&)
            scsi_temp_buffer(5) = CByte(U32Shr(scsi_devices(bus, targetId).block_len, 16&) And &HFF&)
            scsi_temp_buffer(6) = CByte(U32Shr(scsi_devices(bus, targetId).block_len, 8&) And &HFF&)
            scsi_temp_buffer(7) = CByte(scsi_devices(bus, targetId).block_len And &HFF&)
            scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
            scsi_devices(bus, targetId).buffer_length = 8&
            Exit Sub

        Case GPCMD_READ_6, GPCMD_READ_10, GPCMD_READ_12
            lba = SCSIDisk_ReadLBA(cdb)
            blocks = SCSIDisk_ReadBlocks(cdb)
            If SCSIDisk_LbaOutOfRange(bus, targetId, lba, blocks) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_LBA_OUT_OF_RANGE, ASCQ_NONE
                Exit Sub
            End If
            dataLen = CLng(U32ToDouble(blocks) * CDbl(scsi_devices(bus, targetId).block_len))
            If scsi_common_ensure_buffer(dataLen) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            fileOffset = SCSIDisk_FileOffset(lba)
            If scsi_file_read_exact(scsi_devices(bus, targetId).fileHandle, fileOffset, scsi_temp_buffer, dataLen) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_MEDIUM_ERROR, ASC_UNRECOVERED_READ_ERROR, ASCQ_NONE
                Exit Sub
            End If
            scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
            scsi_devices(bus, targetId).buffer_length = dataLen
            Exit Sub

        Case GPCMD_WRITE_6, GPCMD_WRITE_10, GPCMD_WRITE_12
            If scsi_devices(bus, targetId).read_only <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_DATA_PROTECT, ASC_WRITE_PROTECTED, ASCQ_NONE
                Exit Sub
            End If
            lba = SCSIDisk_ReadLBA(cdb)
            blocks = SCSIDisk_ReadBlocks(cdb)
            If SCSIDisk_LbaOutOfRange(bus, targetId, lba, blocks) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_LBA_OUT_OF_RANGE, ASCQ_NONE
                Exit Sub
            End If
            dataLen = CLng(U32ToDouble(blocks) * CDbl(scsi_devices(bus, targetId).block_len))
            If scsi_common_ensure_buffer(dataLen) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            scsi_devices(bus, targetId).pending_lba = lba
            scsi_devices(bus, targetId).pending_blocks = blocks
            For i = 0& To dataLen - 1&
                scsi_temp_buffer(i) = 0&
            Next i
            scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_OUT
            scsi_devices(bus, targetId).buffer_length = dataLen
            Exit Sub

        Case Else
            scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_ILLEGAL_OPCODE, ASCQ_NONE
            Exit Sub
    End Select
End Sub

Public Function scsi_disk_phase_data_out(ByVal bus As Long, ByVal targetId As Long) As Byte
    Dim dataLen As Long
    Dim fileOffset As U64_t

    dataLen = CLng(U32ToDouble(scsi_devices(bus, targetId).pending_blocks) * CDbl(scsi_devices(bus, targetId).block_len))
    fileOffset = SCSIDisk_FileOffset(scsi_devices(bus, targetId).pending_lba)
    If scsi_file_write_exact(scsi_devices(bus, targetId).fileHandle, fileOffset, scsi_temp_buffer, dataLen) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_MEDIUM_ERROR, ASC_WRITE_ERROR, ASCQ_NONE
        scsi_disk_phase_data_out = 1&
        Exit Function
    End If

    scsi_disk_phase_data_out = 0&
End Function

Public Function scsi_disk_attach(ByVal bus As Byte, ByVal targetId As Byte, ByVal path As String) As Long
    Dim fileHandle As Long
    Dim readOnly As Byte

    If (bus >= SCSI_BUS_MAX) Or (targetId >= SCSI_ID_MAX) Then
        scsi_disk_attach = -1&
        Exit Function
    End If

    If scsi_file_open_readwrite_or_readonly(path, fileHandle, readOnly) <> 0& Then
        scsi_disk_attach = -1&
        Exit Function
    End If

    scsi_devices(bus, targetId).buffer_length = 0&
    scsi_devices(bus, targetId).status = 0&
    scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
    scsi_devices(bus, targetId).deviceType = SCSI_FIXED_DISK
    scsi_devices(bus, targetId).fileHandle = fileHandle
    scsi_devices(bus, targetId).openFlag = 1&
    scsi_devices(bus, targetId).block_len = 512&
    scsi_devices(bus, targetId).total_blocks = SCSIDisk_FileBlocks(fileHandle)
    scsi_devices(bus, targetId).path = path
    scsi_devices(bus, targetId).max_transfer_len = &HFFFF&
    scsi_devices(bus, targetId).id = targetId
    scsi_devices(bus, targetId).read_only = readOnly
    scsi_common_clear_sense bus, targetId

    debug_log DEBUG_INFO, "[SCSI] Attached disk target " & CStr(targetId) & " to bus " & CStr(bus) & " from " & path & vbCrLf
    scsi_disk_attach = 0&
End Function
