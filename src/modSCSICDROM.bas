Attribute VB_Name = "modSCSICDROM"
Option Explicit

Private Function SCSICDROM_Be32(ByRef data() As Byte, ByVal offset As Long) As Long
    SCSICDROM_Be32 = U32FromDouble(CDbl(data(offset)) * 16777216# + CDbl(data(offset + 1&)) * 65536# + CDbl(data(offset + 2&)) * 256# + CDbl(data(offset + 3&)))
End Function

Private Function SCSICDROM_ReadLBA(ByRef cdb() As Byte) As Long
    Select Case cdb(0)
        Case GPCMD_READ_10, GPCMD_READ_12
            SCSICDROM_ReadLBA = SCSICDROM_Be32(cdb, 2&)
        Case Else
            SCSICDROM_ReadLBA = 0&
    End Select
End Function

Private Function SCSICDROM_ReadBlocks(ByRef cdb() As Byte) As Long
    Select Case cdb(0)
        Case GPCMD_READ_10
            SCSICDROM_ReadBlocks = ((CLng(cdb(7)) * &H100&) Or CLng(cdb(8)))
        Case GPCMD_READ_12
            SCSICDROM_ReadBlocks = SCSICDROM_Be32(cdb, 6&)
        Case Else
            SCSICDROM_ReadBlocks = 0&
    End Select
End Function

Private Function SCSICDROM_FileOffset(ByVal lba As Long) As U64_t
    Dim offset64 As U64_t

    offset64 = U64_FromU32(lba)
    SCSICDROM_FileOffset = U64_Shl(offset64, 11&)
End Function

Private Function SCSICDROM_FileBlocks(ByVal fileHandle As Long) As U64_t
    Dim size64 As U64_t

    If scsi_file_get_size(fileHandle, size64) <> 0& Then
        SCSICDROM_FileBlocks = U64_Zero()
        Exit Function
    End If

    SCSICDROM_FileBlocks = U64_Shr(size64, 11&)
End Function

Private Function SCSICDROM_TotalBlocksMinusOne(ByVal bus As Long, ByVal targetId As Long) As U64_t
    Dim one As U64_t

    one = U64_FromU32(1&)
    SCSICDROM_TotalBlocksMinusOne = U64_Sub(scsi_devices(bus, targetId).total_blocks, one)
End Function

Private Function SCSICDROM_LbaOutOfRange(ByVal bus As Long, ByVal targetId As Long, ByVal lba As Long, ByVal blocks As Long) As Long
    Dim lba64 As U64_t
    Dim blocks64 As U64_t
    Dim end64 As U64_t
    Dim total As U64_t

    total = scsi_devices(bus, targetId).total_blocks
    lba64 = U64_FromU32(lba)
    If (U64_Lt(total, lba64) <> 0&) Or (U64_Eq(total, lba64) <> 0&) Then
        SCSICDROM_LbaOutOfRange = 1&
        Exit Function
    End If

    blocks64 = U64_FromU32(blocks)
    end64 = U64_Add(lba64, blocks64)
    If U64_Lt(total, end64) <> 0& Then
        SCSICDROM_LbaOutOfRange = 1&
    Else
        SCSICDROM_LbaOutOfRange = 0&
    End If
End Function

Private Sub SCSICDROM_WriteAscii(ByVal offset As Long, ByVal text As String, ByVal maxLen As Long)
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

Private Function SCSICDROM_CommandAllowsUnitAttention(ByVal opcode As Byte) As Long
    Select Case opcode
        Case GPCMD_INQUIRY, GPCMD_REQUEST_SENSE
            SCSICDROM_CommandAllowsUnitAttention = 1&
        Case Else
            SCSICDROM_CommandAllowsUnitAttention = 0&
    End Select
End Function

Private Sub SCSICDROM_SignalMediaChange(ByVal bus As Long, ByVal targetId As Long)
    scsi_devices(bus, targetId).unit_attention = CLng(ASC_MEDIUM_MAY_HAVE_CHANGED)
End Sub

Public Sub scsi_cdrom_request_sense(ByVal bus As Long, ByVal targetId As Long, ByRef buffer() As Byte, ByVal allocLength As Long)
    Dim copyLen As Long
    Dim i As Long

    If (scsi_devices(bus, targetId).status_byte = 0&) And (scsi_devices(bus, targetId).unit_attention <> 0&) Then
        scsi_common_set_sense bus, targetId, SENSE_UNIT_ATTENTION, CByte(scsi_devices(bus, targetId).unit_attention And &HFF&), ASCQ_NONE
    End If

    copyLen = allocLength
    If copyLen > 18& Then copyLen = 18&
    If copyLen < 0& Then copyLen = 0&
    If copyLen = 0& Then Exit Sub

    ReDim buffer(0 To copyLen - 1&) As Byte
    For i = 0& To copyLen - 1&
        buffer(i) = scsi_devices(bus, targetId).sense(i)
    Next i
    If buffer(2) = SENSE_UNIT_ATTENTION Then
        scsi_devices(bus, targetId).unit_attention = 0&
    End If
    scsi_common_clear_sense bus, targetId
End Sub

Public Sub scsi_cdrom_reset(ByVal bus As Long, ByVal targetId As Long)
    scsi_common_clear_sense bus, targetId
End Sub

Private Sub SCSICDROM_BuildInquiry(ByVal bus As Long, ByVal targetId As Long)
    Dim i As Long

    If scsi_common_ensure_buffer(36&) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        Exit Sub
    End If

    For i = 0& To 35&
        scsi_temp_buffer(i) = 0&
    Next i
    scsi_temp_buffer(0) = &H5&
    scsi_temp_buffer(1) = &H80&
    scsi_temp_buffer(2) = &H5&
    scsi_temp_buffer(3) = &H2&
    scsi_temp_buffer(4) = 31&
    SCSICDROM_WriteAscii 8&, "BasicBox", 8&
    SCSICDROM_WriteAscii 16&, "SCSI CD-ROM    ", 16&
    SCSICDROM_WriteAscii 32&, "0001", 4&
    scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
    scsi_devices(bus, targetId).buffer_length = 36&
End Sub

Private Sub SCSICDROM_BuildCapacity(ByVal bus As Long, ByVal targetId As Long)
    Dim lastBlock As U64_t
    Dim i As Long

    If U64_IsZero(scsi_devices(bus, targetId).total_blocks) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_NOT_READY, ASC_MEDIUM_NOT_PRESENT, ASCQ_NONE
        Exit Sub
    End If
    If scsi_common_ensure_buffer(8&) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        Exit Sub
    End If

    For i = 0& To 7&
        scsi_temp_buffer(i) = 0&
    Next i
    lastBlock = SCSICDROM_TotalBlocksMinusOne(bus, targetId)
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
End Sub

Private Sub SCSICDROM_BuildModeSense(ByVal bus As Long, ByVal targetId As Long, ByRef cdb() As Byte)
    Dim isTen As Long
    Dim totalLen As Long
    Dim i As Long

    isTen = IIf(cdb(0) = GPCMD_MODE_SENSE_10, 1&, 0&)
    totalLen = IIf(isTen <> 0&, 32&, 28&)
    If scsi_common_ensure_buffer(totalLen) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        Exit Sub
    End If

    For i = 0& To totalLen - 1&
        scsi_temp_buffer(i) = 0&
    Next i
    If isTen <> 0& Then
        scsi_temp_buffer(1) = CByte((totalLen - 2&) And &HFF&)
        scsi_temp_buffer(8) = &H2A&
        scsi_temp_buffer(9) = &H12&
        scsi_temp_buffer(10) = &H3&
        scsi_temp_buffer(12) = &H71&
        scsi_temp_buffer(13) = &H0&
        scsi_temp_buffer(14) = &H2&
    Else
        scsi_temp_buffer(0) = CByte((totalLen - 1&) And &HFF&)
        scsi_temp_buffer(4) = &H2A&
        scsi_temp_buffer(5) = &H12&
        scsi_temp_buffer(6) = &H3&
        scsi_temp_buffer(8) = &H71&
        scsi_temp_buffer(9) = &H0&
        scsi_temp_buffer(10) = &H2&
    End If
    scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
    scsi_devices(bus, targetId).buffer_length = totalLen
End Sub

Private Sub SCSICDROM_BuildTOC(ByVal bus As Long, ByVal targetId As Long, ByRef cdb() As Byte)
    Dim msf As Long
    Dim leadout As Long
    Dim i As Long

    If scsi_common_ensure_buffer(20&) <> 0& Then
        scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
        Exit Sub
    End If

    msf = (cdb(1) And &H2&)
    leadout = scsi_devices(bus, targetId).total_blocks.Lo
    For i = 0& To 19&
        scsi_temp_buffer(i) = 0&
    Next i
    scsi_temp_buffer(1) = 18&
    scsi_temp_buffer(2) = 1&
    scsi_temp_buffer(3) = 1&
    scsi_temp_buffer(5) = &H14&
    scsi_temp_buffer(6) = 1&
    scsi_temp_buffer(13) = &H16&
    scsi_temp_buffer(14) = &HAA&

    If msf <> 0& Then
        scsi_temp_buffer(9) = 0&
        scsi_temp_buffer(10) = 2&
        scsi_temp_buffer(11) = 0&
        scsi_temp_buffer(17) = CByte((leadout \ (75& * 60&)) And &HFF&)
        scsi_temp_buffer(18) = CByte(((leadout \ 75&) Mod 60&) And &HFF&)
        scsi_temp_buffer(19) = CByte((leadout Mod 75&) And &HFF&)
    Else
        scsi_temp_buffer(16) = CByte(U32Shr(leadout, 24&) And &HFF&)
        scsi_temp_buffer(17) = CByte(U32Shr(leadout, 16&) And &HFF&)
        scsi_temp_buffer(18) = CByte(U32Shr(leadout, 8&) And &HFF&)
        scsi_temp_buffer(19) = CByte(leadout And &HFF&)
    End If

    scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
    scsi_devices(bus, targetId).buffer_length = 20&
End Sub

Public Sub scsi_cdrom_command(ByVal bus As Long, ByVal targetId As Long, ByRef cdb() As Byte)
    Dim lba As Long
    Dim blocks As Long
    Dim dataLen As Long
    Dim i As Long
    Dim fileOffset As U64_t

    If cdb(0) <> GPCMD_REQUEST_SENSE Then
        scsi_common_clear_sense bus, targetId
    End If
    scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
    scsi_devices(bus, targetId).buffer_length = 0&

    If (scsi_devices(bus, targetId).unit_attention <> 0&) And (SCSICDROM_CommandAllowsUnitAttention(cdb(0)) = 0&) Then
        scsi_common_set_sense bus, targetId, SENSE_UNIT_ATTENTION, CByte(scsi_devices(bus, targetId).unit_attention And &HFF&), ASCQ_NONE
        Exit Sub
    End If

    Select Case cdb(0)
        Case GPCMD_TEST_UNIT_READY
            If U64_IsZero(scsi_devices(bus, targetId).total_blocks) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_NOT_READY, ASC_MEDIUM_NOT_PRESENT, ASCQ_NONE
            End If
            Exit Sub

        Case GPCMD_INQUIRY
            SCSICDROM_BuildInquiry bus, targetId
            Exit Sub

        Case GPCMD_REQUEST_SENSE
            If (scsi_devices(bus, targetId).status_byte = 0&) And (scsi_devices(bus, targetId).unit_attention <> 0&) Then
                scsi_common_set_sense bus, targetId, SENSE_UNIT_ATTENTION, CByte(scsi_devices(bus, targetId).unit_attention And &HFF&), ASCQ_NONE
            End If
            If scsi_common_ensure_buffer(18&) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            For i = 0& To 17&
                scsi_temp_buffer(i) = scsi_devices(bus, targetId).sense(i)
            Next i
            scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
            scsi_devices(bus, targetId).buffer_length = 18&
            If scsi_temp_buffer(2) = SENSE_UNIT_ATTENTION Then
                scsi_devices(bus, targetId).unit_attention = 0&
            End If
            scsi_common_clear_sense bus, targetId
            Exit Sub

        Case GPCMD_MODE_SENSE_6, GPCMD_MODE_SENSE_10
            SCSICDROM_BuildModeSense bus, targetId, cdb
            Exit Sub

        Case GPCMD_READ_CDROM_CAPACITY
            SCSICDROM_BuildCapacity bus, targetId
            Exit Sub

        Case GPCMD_READ_TOC_PMA_ATIP
            SCSICDROM_BuildTOC bus, targetId, cdb
            Exit Sub

        Case GPCMD_START_STOP_UNIT, GPCMD_PREVENT_REMOVAL
            Exit Sub

        Case GPCMD_READ_10, GPCMD_READ_12
            If U64_IsZero(scsi_devices(bus, targetId).total_blocks) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_NOT_READY, ASC_MEDIUM_NOT_PRESENT, ASCQ_NONE
                Exit Sub
            End If
            lba = SCSICDROM_ReadLBA(cdb)
            blocks = SCSICDROM_ReadBlocks(cdb)
            If SCSICDROM_LbaOutOfRange(bus, targetId, lba, blocks) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_LBA_OUT_OF_RANGE, ASCQ_NONE
                Exit Sub
            End If
            dataLen = CLng(U32ToDouble(blocks) * CDbl(scsi_devices(bus, targetId).block_len))
            If scsi_common_ensure_buffer(dataLen) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
                Exit Sub
            End If
            fileOffset = SCSICDROM_FileOffset(lba)
            If scsi_file_read_exact(scsi_devices(bus, targetId).fileHandle, fileOffset, scsi_temp_buffer, dataLen) <> 0& Then
                scsi_common_set_sense bus, targetId, SENSE_MEDIUM_ERROR, ASC_UNRECOVERED_READ_ERROR, ASCQ_NONE
                Exit Sub
            End If
            scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_IN
            scsi_devices(bus, targetId).buffer_length = dataLen
            Exit Sub

        Case Else
            scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_ILLEGAL_OPCODE, ASCQ_NONE
            Exit Sub
    End Select
End Sub

Private Sub SCSICDROM_ClearMediaState(ByVal bus As Byte, ByVal targetId As Byte)
    scsi_devices(bus, targetId).buffer_length = 0&
    scsi_devices(bus, targetId).status = 0&
    scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
    scsi_devices(bus, targetId).fileHandle = 0&
    scsi_devices(bus, targetId).openFlag = 0&
    scsi_devices(bus, targetId).deviceType = SCSI_REMOVABLE_CDROM
    scsi_devices(bus, targetId).block_len = 2048&
    scsi_devices(bus, targetId).total_blocks = U64_Zero()
    scsi_devices(bus, targetId).path = vbNullString
    scsi_devices(bus, targetId).max_transfer_len = &HFFFF&
    scsi_devices(bus, targetId).id = targetId
    scsi_common_clear_sense bus, targetId
End Sub

Public Sub scsi_cdrom_eject(ByVal bus As Byte, ByVal targetId As Byte)
    If (bus >= SCSI_BUS_MAX) Or (targetId >= SCSI_ID_MAX) Then Exit Sub

    If scsi_devices(bus, targetId).openFlag <> 0& Then
        scsi_file_close scsi_devices(bus, targetId).fileHandle
    End If

    Call SCSICDROM_ClearMediaState(bus, targetId)
    Call SCSICDROM_SignalMediaChange(bus, targetId)
    debug_log DEBUG_INFO, "[SCSI] Ejected CD-ROM target " & CStr(targetId) & " on bus " & CStr(bus) & vbCrLf
End Sub

Public Function scsi_cdrom_attach(ByVal bus As Byte, ByVal targetId As Byte, ByVal path As String) As Long
    Dim fileHandle As Long

    If (bus >= SCSI_BUS_MAX) Or (targetId >= SCSI_ID_MAX) Then
        scsi_cdrom_attach = -1&
        Exit Function
    End If

    If scsi_file_open_readonly(path, fileHandle) <> 0& Then
        scsi_cdrom_attach = -1&
        Exit Function
    End If

    If scsi_devices(bus, targetId).openFlag <> 0& Then
        scsi_file_close scsi_devices(bus, targetId).fileHandle
    End If

    Call SCSICDROM_ClearMediaState(bus, targetId)
    scsi_devices(bus, targetId).fileHandle = fileHandle
    scsi_devices(bus, targetId).openFlag = 1&
    scsi_devices(bus, targetId).total_blocks = SCSICDROM_FileBlocks(fileHandle)
    scsi_devices(bus, targetId).path = path
    scsi_common_clear_sense bus, targetId
    Call SCSICDROM_SignalMediaChange(bus, targetId)

    debug_log DEBUG_INFO, "[SCSI] Attached CD-ROM target " & CStr(targetId) & " to bus " & CStr(bus) & " from " & path & vbCrLf
    scsi_cdrom_attach = 0&
End Function
