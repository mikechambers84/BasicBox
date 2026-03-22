Attribute VB_Name = "modSCSIDevice"
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1&
Private Const FILE_SHARE_WRITE As Long = &H2&
Private Const OPEN_EXISTING As Long = 3&
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Private Const FILE_BEGIN As Long = 0&
Private Const INVALID_HANDLE_VALUE As Long = -1&
Private Const NO_ERROR As Long = 0&

Public Const SCSI_LUN_USE_CDB As Long = &HFF&

Public Const GPCMD_TEST_UNIT_READY As Byte = &H0&
Public Const GPCMD_REQUEST_SENSE As Byte = &H3&
Public Const GPCMD_FORMAT_UNIT As Byte = &H4&
Public Const GPCMD_READ_6 As Byte = &H8&
Public Const GPCMD_WRITE_6 As Byte = &HA&
Public Const GPCMD_INQUIRY As Byte = &H12&
Public Const GPCMD_MODE_SELECT_6 As Byte = &H15&
Public Const GPCMD_MODE_SENSE_6 As Byte = &H1A&
Public Const GPCMD_START_STOP_UNIT As Byte = &H1B&
Public Const GPCMD_PREVENT_REMOVAL As Byte = &H1E&
Public Const GPCMD_READ_CDROM_CAPACITY As Byte = &H25&
Public Const GPCMD_READ_10 As Byte = &H28&
Public Const GPCMD_WRITE_10 As Byte = &H2A&
Public Const GPCMD_VERIFY_10 As Byte = &H2F&
Public Const GPCMD_SYNCHRONIZE_CACHE As Byte = &H35&
Public Const GPCMD_READ_TOC_PMA_ATIP As Byte = &H43&
Public Const GPCMD_MODE_SELECT_10 As Byte = &H55&
Public Const GPCMD_MODE_SENSE_10 As Byte = &H5A&
Public Const GPCMD_READ_12 As Byte = &HA8&
Public Const GPCMD_WRITE_12 As Byte = &HAA&

Public Const SCSI_STATUS_OK As Byte = 0&
Public Const SCSI_STATUS_CHECK_CONDITION As Byte = 2&

Public Const SENSE_NONE As Byte = 0&
Public Const SENSE_NOT_READY As Byte = 2&
Public Const SENSE_MEDIUM_ERROR As Byte = 3&
Public Const SENSE_ILLEGAL_REQUEST As Byte = 5&
Public Const SENSE_UNIT_ATTENTION As Byte = 6&
Public Const SENSE_DATA_PROTECT As Byte = 7&

Public Const ASC_NONE As Byte = &H0&
Public Const ASC_NOT_READY As Byte = &H4&
Public Const ASC_WRITE_ERROR As Byte = &HC&
Public Const ASC_UNRECOVERED_READ_ERROR As Byte = &H11&
Public Const ASC_ILLEGAL_OPCODE As Byte = &H20&
Public Const ASC_LBA_OUT_OF_RANGE As Byte = &H21&
Public Const ASC_INV_FIELD_IN_CMD_PACKET As Byte = &H24&
Public Const ASC_INV_FIELD_IN_PARAMETER_LIST As Byte = &H26&
Public Const ASC_WRITE_PROTECTED As Byte = &H27&
Public Const ASC_MEDIUM_MAY_HAVE_CHANGED As Byte = &H28&
Public Const ASC_CAPACITY_DATA_CHANGED As Byte = &H2A&
Public Const ASC_INCOMPATIBLE_FORMAT As Byte = &H30&
Public Const ASC_MEDIUM_NOT_PRESENT As Byte = &H3A&

Public Const ASCQ_NONE As Byte = &H0&
Public Const ASCQ_UNIT_IN_PROCESS_OF_BECOMING_READY As Byte = &H1&

Public Const SCSI_PHASE_DATA_OUT As Byte = 0&
Public Const SCSI_PHASE_DATA_IN As Byte = 1&
Public Const SCSI_PHASE_COMMAND As Byte = 2&
Public Const SCSI_PHASE_STATUS As Byte = 3&

Public Const SCSI_NONE As Long = &H60&
Public Const SCSI_FIXED_DISK As Long = &H0&
Public Const SCSI_REMOVABLE_CDROM As Long = &H8005&

Public Type SCSI_DEVICE_t
    buffer_length As Long
    status As Byte
    phase As Byte
    deviceType As Long
    id As Byte
    cur_lun As Byte
    max_transfer_len As Long
    buffer_pos As Long
    total_length As Long
    unit_attention As Long
    block_len As Long
    callbackDelay As Double
    status_byte As Byte
    total_blocks As U64_t
    pending_lba As Long
    pending_blocks As Long
    read_only As Byte
    fileHandle As Long
    openFlag As Byte
    path As String * 512
    sense(0 To 255) As Byte
    current_cdb(0 To 15) As Byte
End Type

Public scsi_devices(0 To SCSI_BUS_MAX - 1, 0 To SCSI_ID_MAX - 1) As SCSI_DEVICE_t
Public scsi_temp_buffer() As Byte
Public scsi_temp_buffer_sz As Long

Private scsi_null_device_sense(0 To 17) As Byte

Private Sub SCSIDevice_InitNullSense()
    scsi_null_device_sense(0) = &H70&
    scsi_null_device_sense(1) = &H0&
    scsi_null_device_sense(2) = SENSE_ILLEGAL_REQUEST
    scsi_null_device_sense(3) = &H0&
    scsi_null_device_sense(4) = &H0&
    scsi_null_device_sense(5) = &H0&
    scsi_null_device_sense(6) = &H0&
    scsi_null_device_sense(7) = &HA&
    scsi_null_device_sense(8) = &H0&
    scsi_null_device_sense(9) = &H0&
    scsi_null_device_sense(10) = &H0&
    scsi_null_device_sense(11) = &H0&
    scsi_null_device_sense(12) = ASC_INV_FIELD_IN_CMD_PACKET
    scsi_null_device_sense(13) = &H0&
    scsi_null_device_sense(14) = &H0&
    scsi_null_device_sense(15) = &H0&
    scsi_null_device_sense(16) = &H0&
    scsi_null_device_sense(17) = &H0&
End Sub

Private Function SCSIDevice_SetFilePointer64(ByVal hFile As Long, ByRef position As U64_t, ByVal moveMethod As Long) As Long
    Dim highPart As Long
    Dim lowPart As Long
    Dim errNum As Long

    highPart = position.Hi
    SetLastError NO_ERROR
    lowPart = SetFilePointer(hFile, position.Lo, highPart, moveMethod)
    If lowPart = -1& Then
        errNum = GetLastError()
        If errNum <> NO_ERROR Then
            SCSIDevice_SetFilePointer64 = -1&
            Exit Function
        End If
    End If

    position.Lo = lowPart
    position.Hi = highPart
    SCSIDevice_SetFilePointer64 = 0&
End Function

Public Function scsi_file_open_readwrite_or_readonly(ByVal path As String, ByRef fileHandle As Long, ByRef readOnly As Byte) As Long
    fileHandle = CreateFile(path, (GENERIC_READ Or GENERIC_WRITE), (FILE_SHARE_READ Or FILE_SHARE_WRITE), 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    If fileHandle <> INVALID_HANDLE_VALUE Then
        readOnly = 0&
        scsi_file_open_readwrite_or_readonly = 0&
        Exit Function
    End If

    fileHandle = CreateFile(path, GENERIC_READ, (FILE_SHARE_READ Or FILE_SHARE_WRITE), 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    If fileHandle = INVALID_HANDLE_VALUE Then
        readOnly = 0&
        scsi_file_open_readwrite_or_readonly = -1&
    Else
        readOnly = 1&
        scsi_file_open_readwrite_or_readonly = 0&
    End If
End Function

Public Function scsi_file_open_readonly(ByVal path As String, ByRef fileHandle As Long) As Long
    fileHandle = CreateFile(path, GENERIC_READ, (FILE_SHARE_READ Or FILE_SHARE_WRITE), 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    If fileHandle = INVALID_HANDLE_VALUE Then
        scsi_file_open_readonly = -1&
    Else
        scsi_file_open_readonly = 0&
    End If
End Function

Public Sub scsi_file_close(ByVal fileHandle As Long)
    If (fileHandle <> 0&) And (fileHandle <> INVALID_HANDLE_VALUE) Then
        CloseHandle fileHandle
    End If
End Sub

Public Function scsi_file_get_size(ByVal fileHandle As Long, ByRef sizeOut As U64_t) As Long
    Dim highPart As Long
    Dim lowPart As Long
    Dim errNum As Long

    highPart = 0&
    SetLastError NO_ERROR
    lowPart = GetFileSize(fileHandle, highPart)
    If lowPart = -1& Then
        errNum = GetLastError()
        If errNum <> NO_ERROR Then
            scsi_file_get_size = -1&
            Exit Function
        End If
    End If

    sizeOut.Lo = lowPart
    sizeOut.Hi = highPart
    scsi_file_get_size = 0&
End Function

Public Function scsi_file_read_exact(ByVal fileHandle As Long, ByRef byteOffset As U64_t, ByRef buffer() As Byte, ByVal length As Long) As Long
    Dim bytesRead As Long

    If length < 0& Then
        scsi_file_read_exact = -1&
        Exit Function
    End If
    If length = 0& Then
        scsi_file_read_exact = 0&
        Exit Function
    End If
    If SCSIDevice_SetFilePointer64(fileHandle, byteOffset, FILE_BEGIN) <> 0& Then
        scsi_file_read_exact = -1&
        Exit Function
    End If

    bytesRead = 0&
    If ReadFile(fileHandle, buffer(0&), length, bytesRead, 0&) = 0& Then
        scsi_file_read_exact = -1&
        Exit Function
    End If
    If bytesRead <> length Then
        scsi_file_read_exact = -1&
        Exit Function
    End If

    scsi_file_read_exact = 0&
End Function

Public Function scsi_file_write_exact(ByVal fileHandle As Long, ByRef byteOffset As U64_t, ByRef buffer() As Byte, ByVal length As Long) As Long
    Dim bytesWritten As Long

    If length < 0& Then
        scsi_file_write_exact = -1&
        Exit Function
    End If
    If length = 0& Then
        scsi_file_write_exact = 0&
        Exit Function
    End If
    If SCSIDevice_SetFilePointer64(fileHandle, byteOffset, FILE_BEGIN) <> 0& Then
        scsi_file_write_exact = -1&
        Exit Function
    End If

    bytesWritten = 0&
    If WriteFile(fileHandle, buffer(0&), length, bytesWritten, 0&) = 0& Then
        scsi_file_write_exact = -1&
        Exit Function
    End If
    If bytesWritten <> length Then
        scsi_file_write_exact = -1&
        Exit Function
    End If
    If FlushFileBuffers(fileHandle) = 0& Then
        scsi_file_write_exact = -1&
        Exit Function
    End If

    scsi_file_write_exact = 0&
End Function

Public Function scsi_common_ensure_buffer(ByVal length As Long) As Long
    On Error GoTo EnsureFail

    If length < 0& Then length = 0&
    If scsi_temp_buffer_sz >= length Then
        scsi_common_ensure_buffer = 0&
        Exit Function
    End If

    If length = 0& Then
        ReDim scsi_temp_buffer(0 To 0) As Byte
    Else
        ReDim scsi_temp_buffer(0 To length - 1&) As Byte
    End If
    scsi_temp_buffer_sz = length
    scsi_common_ensure_buffer = 0&
    Exit Function

EnsureFail:
    debug_log DEBUG_ERROR, "[SCSI] Unable to allocate " & CStr(length) & " bytes for target buffer" & vbCrLf
    scsi_common_ensure_buffer = -1&
End Function

Public Function scsi_device_get_callback(ByVal bus As Long, ByVal targetId As Long) As Double
    If scsi_device_valid(bus, targetId) <> 0& Then
        scsi_device_get_callback = scsi_devices(bus, targetId).callbackDelay
    Else
        scsi_device_get_callback = -1#
    End If
End Function

Public Sub scsi_common_set_sense(ByVal bus As Long, ByVal targetId As Long, ByVal key As Byte, ByVal asc As Byte, ByVal ascq As Byte)
    Dim i As Long

    For i = 0& To 255&
        scsi_devices(bus, targetId).sense(i) = 0&
    Next i
    scsi_devices(bus, targetId).sense(0) = &H70&
    scsi_devices(bus, targetId).sense(2) = key
    scsi_devices(bus, targetId).sense(7) = &HA&
    scsi_devices(bus, targetId).sense(12) = asc
    scsi_devices(bus, targetId).sense(13) = ascq
    scsi_devices(bus, targetId).status_byte = 1&
End Sub

Public Sub scsi_common_clear_sense(ByVal bus As Long, ByVal targetId As Long)
    Dim i As Long

    For i = 0& To 255&
        scsi_devices(bus, targetId).sense(i) = 0&
    Next i
    scsi_devices(bus, targetId).sense(0) = &H70&
    scsi_devices(bus, targetId).sense(7) = &HA&
    scsi_devices(bus, targetId).status_byte = 0&
End Sub

Public Function scsi_device_present(ByVal bus As Long, ByVal targetId As Long) As Long
    If scsi_devices(bus, targetId).deviceType <> SCSI_NONE Then
        scsi_device_present = 1&
    Else
        scsi_device_present = 0&
    End If
End Function

Public Function scsi_device_valid(ByVal bus As Long, ByVal targetId As Long) As Long
    scsi_device_valid = scsi_device_present(bus, targetId)
End Function

Public Function scsi_device_cdb_length(ByVal bus As Long, ByVal targetId As Long) As Long
    scsi_device_cdb_length = 12&
End Function

Public Function scsi_device_copy_sense(ByVal bus As Long, ByVal targetId As Long, ByRef buffer() As Byte, ByVal allocLength As Long) As Long
    Dim copyLen As Long
    Dim i As Long

    If allocLength < 0& Then allocLength = 0&
    copyLen = allocLength
    If copyLen > 18& Then copyLen = 18&

    If copyLen <= 0& Then
        scsi_device_copy_sense = 0&
        Exit Function
    End If

    ReDim buffer(0 To copyLen - 1&) As Byte
    If scsi_device_valid(bus, targetId) <> 0& Then
        For i = 0& To copyLen - 1&
            buffer(i) = scsi_devices(bus, targetId).sense(i)
        Next i
    Else
        For i = 0& To copyLen - 1&
            buffer(i) = scsi_null_device_sense(i)
        Next i
    End If
    scsi_device_copy_sense = copyLen
End Function

Public Sub scsi_device_request_sense(ByVal bus As Long, ByVal targetId As Long, ByRef buffer() As Byte, ByVal allocLength As Long)
    Dim copyLen As Long
    Dim i As Long

    copyLen = allocLength
    If copyLen > 18& Then copyLen = 18&
    If copyLen < 0& Then copyLen = 0&

    If copyLen <= 0& Then Exit Sub

    ReDim buffer(0 To copyLen - 1&) As Byte

    If scsi_device_valid(bus, targetId) = 0& Then
        For i = 0& To copyLen - 1&
            buffer(i) = scsi_null_device_sense(i)
        Next i
        Exit Sub
    End If

    Select Case scsi_devices(bus, targetId).deviceType
        Case SCSI_FIXED_DISK
            scsi_disk_request_sense bus, targetId, buffer, copyLen
        Case SCSI_REMOVABLE_CDROM
            scsi_cdrom_request_sense bus, targetId, buffer, copyLen
        Case Else
            For i = 0& To copyLen - 1&
                buffer(i) = scsi_null_device_sense(i)
            Next i
    End Select
End Sub

Public Sub scsi_device_reset(ByVal bus As Long, ByVal targetId As Long)
    If scsi_device_valid(bus, targetId) = 0& Then Exit Sub

    Select Case scsi_devices(bus, targetId).deviceType
        Case SCSI_FIXED_DISK
            scsi_disk_reset bus, targetId
        Case SCSI_REMOVABLE_CDROM
            scsi_cdrom_reset bus, targetId
    End Select
End Sub

Public Sub scsi_device_identify(ByVal bus As Long, ByVal targetId As Long, ByVal lun As Long)
    If (bus < 0&) Or (bus >= SCSI_BUS_MAX) Then Exit Sub
    If (targetId < 0&) Or (targetId >= SCSI_ID_MAX) Then Exit Sub
    If scsi_device_valid(bus, targetId) = 0& Then Exit Sub

    scsi_devices(bus, targetId).cur_lun = CByte(lun And &HFF&)
End Sub

Public Sub scsi_device_command_phase0(ByVal bus As Long, ByVal targetId As Long, ByRef cdb() As Byte)
    Dim i As Long

    If scsi_device_valid(bus, targetId) = 0& Then
        scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
        scsi_devices(bus, targetId).status = SCSI_STATUS_CHECK_CONDITION
        Exit Sub
    End If

    scsi_devices(bus, targetId).phase = SCSI_PHASE_COMMAND
    scsi_devices(bus, targetId).status_byte = 0&
    For i = 0& To 15&
        If i <= UBound(cdb) Then
            scsi_devices(bus, targetId).current_cdb(i) = cdb(i)
        Else
            scsi_devices(bus, targetId).current_cdb(i) = 0&
        End If
    Next i

    Select Case scsi_devices(bus, targetId).deviceType
        Case SCSI_FIXED_DISK
            scsi_disk_command bus, targetId, cdb
        Case SCSI_REMOVABLE_CDROM
            scsi_cdrom_command bus, targetId, cdb
        Case Else
            scsi_common_set_sense bus, targetId, SENSE_ILLEGAL_REQUEST, ASC_INV_FIELD_IN_CMD_PACKET, ASCQ_NONE
            scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
            scsi_devices(bus, targetId).buffer_length = 0&
    End Select

    If scsi_devices(bus, targetId).status_byte <> 0& Then
        scsi_devices(bus, targetId).status = SCSI_STATUS_CHECK_CONDITION
    Else
        scsi_devices(bus, targetId).status = SCSI_STATUS_OK
    End If
End Sub

Public Sub scsi_device_command_stop(ByVal bus As Long, ByVal targetId As Long)
    If scsi_device_valid(bus, targetId) = 0& Then Exit Sub

    If scsi_devices(bus, targetId).status_byte <> 0& Then
        scsi_devices(bus, targetId).status = SCSI_STATUS_CHECK_CONDITION
    Else
        scsi_devices(bus, targetId).status = SCSI_STATUS_OK
    End If
End Sub

Public Sub scsi_device_command_phase1(ByVal bus As Long, ByVal targetId As Long)
    If scsi_device_valid(bus, targetId) = 0& Then Exit Sub

    If scsi_devices(bus, targetId).phase = SCSI_PHASE_DATA_OUT Then
        Select Case scsi_devices(bus, targetId).deviceType
            Case SCSI_FIXED_DISK
                Call scsi_disk_phase_data_out(bus, targetId)
            Case Else
                Call scsi_device_command_stop(bus, targetId)
        End Select
    Else
        Call scsi_device_command_stop(bus, targetId)
    End If

    If scsi_devices(bus, targetId).status_byte <> 0& Then
        scsi_devices(bus, targetId).status = SCSI_STATUS_CHECK_CONDITION
    Else
        scsi_devices(bus, targetId).status = SCSI_STATUS_OK
    End If
End Sub

Public Sub scsi_device_close_all()
    Dim bus As Long
    Dim targetId As Long

    For bus = 0& To SCSI_BUS_MAX - 1&
        For targetId = 0& To SCSI_ID_MAX - 1&
            ' Intentionally no-op: the C implementation only invokes command_stop handlers,
            ' and the current disk/CD targets do not provide one.
        Next targetId
    Next bus
End Sub

Public Sub scsi_device_init()
    Dim bus As Long
    Dim targetId As Long
    Dim i As Long

    SCSIDevice_InitNullSense

    For bus = 0& To SCSI_BUS_MAX - 1&
        For targetId = 0& To SCSI_ID_MAX - 1&
            scsi_devices(bus, targetId).buffer_length = 0&
            scsi_devices(bus, targetId).status = 0&
            scsi_devices(bus, targetId).phase = SCSI_PHASE_STATUS
            scsi_devices(bus, targetId).deviceType = SCSI_NONE
            scsi_devices(bus, targetId).id = CByte(targetId And &HFF&)
            scsi_devices(bus, targetId).cur_lun = 0&
            scsi_devices(bus, targetId).max_transfer_len = 0&
            scsi_devices(bus, targetId).buffer_pos = 0&
            scsi_devices(bus, targetId).total_length = 0&
            scsi_devices(bus, targetId).unit_attention = 0&
            scsi_devices(bus, targetId).block_len = 0&
            scsi_devices(bus, targetId).callbackDelay = -1#
            scsi_devices(bus, targetId).status_byte = 0&
            scsi_devices(bus, targetId).total_blocks = U64_Zero()
            scsi_devices(bus, targetId).pending_lba = 0&
            scsi_devices(bus, targetId).pending_blocks = 0&
            scsi_devices(bus, targetId).read_only = 0&
            scsi_devices(bus, targetId).fileHandle = 0&
            scsi_devices(bus, targetId).openFlag = 0&
            scsi_devices(bus, targetId).path = vbNullString
            For i = 0& To 255&
                scsi_devices(bus, targetId).sense(i) = 0&
            Next i
            For i = 0& To 15&
                scsi_devices(bus, targetId).current_cdb(i) = 0&
            Next i
            scsi_common_clear_sense bus, targetId
        Next targetId
    Next bus

    scsi_temp_buffer_sz = 0&
    ReDim scsi_temp_buffer(0 To 0) As Byte
End Sub
