Attribute VB_Name = "modFDD"
Option Explicit

Public Const FDD_NUM As Long = 4&
Public Const FDD_SECTOR_SIZE As Long = 512&

Public Const FDD_IO_OK As Long = 0&
Public Const FDD_IO_NO_MEDIA As Long = 1&
Public Const FDD_IO_NOT_FOUND As Long = 2&
Public Const FDD_IO_WRITE_PROTECT As Long = 3&
Public Const FDD_IO_INVALID_SIZE As Long = 4&

Private Const FDD_AT_DRIVES As Long = 2&
Private Const FDD_DEFAULT_MAX_TRACK As Long = 79&
Private Const FDD_POLL_HZ As Long = 50&
Private Const FDD_SEEK_DELAY_USEC As Long = 10000&
Private Const FDD_MAX_SIDES As Long = 2&
Private Const FDD_MAX_SECTORS_PER_TRACK As Long = 36&
Private Const FDD_TRACK_BUFFER_SIZE As Long = FDD_MAX_SECTORS_PER_TRACK * FDD_SECTOR_SIZE

Private Type FDD_GEOMETRY_t
    sizeBytes As Long
    tracks As Long
    sides As Long
    sectorsPerTrack As Long
    is525 As Byte
End Type

Private Type FDD_DRIVE_t
    installed As Byte
    inserted As Byte
    writeProtected As Byte
    changed As Byte
    motorEnabled As Byte
    track As Long
    head As Long
    seekInProgress As Byte
    seekTarget As Long
    tracks As Long
    sides As Long
    sectorsPerTrack As Long
    fileSize As Long
    is525 As Byte
    readAddressTrack As Long
    readAddressHead As Long
    readAddressSector As Long
    trackBufferValid As Byte
    trackBufferTrack As Long
    fileNum As Integer
    pollTimer As Long
    seekTimer As Long
    stagedPath As String
End Type

Private fdd(0& To FDD_NUM - 1&) As FDD_DRIVE_t
Private fdd_initialized As Byte
Private fdd_geometry(0& To 7&) As FDD_GEOMETRY_t
Private fdd_geometryInitialized As Byte
Private fdd_trackData(0& To FDD_NUM - 1&, 0& To FDD_MAX_SIDES - 1&, 0& To FDD_TRACK_BUFFER_SIZE - 1&) As Byte

Private Sub FDD_InitGeometryTable()
    If fdd_geometryInitialized <> 0& Then Exit Sub

    fdd_geometry(0&).sizeBytes = 160& * 1024&
    fdd_geometry(0&).tracks = 40&
    fdd_geometry(0&).sides = 1&
    fdd_geometry(0&).sectorsPerTrack = 8&
    fdd_geometry(0&).is525 = 1&

    fdd_geometry(1&).sizeBytes = 180& * 1024&
    fdd_geometry(1&).tracks = 40&
    fdd_geometry(1&).sides = 1&
    fdd_geometry(1&).sectorsPerTrack = 9&
    fdd_geometry(1&).is525 = 1&

    fdd_geometry(2&).sizeBytes = 320& * 1024&
    fdd_geometry(2&).tracks = 40&
    fdd_geometry(2&).sides = 2&
    fdd_geometry(2&).sectorsPerTrack = 8&
    fdd_geometry(2&).is525 = 1&

    fdd_geometry(3&).sizeBytes = 360& * 1024&
    fdd_geometry(3&).tracks = 40&
    fdd_geometry(3&).sides = 2&
    fdd_geometry(3&).sectorsPerTrack = 9&
    fdd_geometry(3&).is525 = 1&

    fdd_geometry(4&).sizeBytes = 720& * 1024&
    fdd_geometry(4&).tracks = 80&
    fdd_geometry(4&).sides = 2&
    fdd_geometry(4&).sectorsPerTrack = 9&
    fdd_geometry(4&).is525 = 0&

    fdd_geometry(5&).sizeBytes = 1200& * 1024&
    fdd_geometry(5&).tracks = 80&
    fdd_geometry(5&).sides = 2&
    fdd_geometry(5&).sectorsPerTrack = 15&
    fdd_geometry(5&).is525 = 1&

    fdd_geometry(6&).sizeBytes = 1440& * 1024&
    fdd_geometry(6&).tracks = 80&
    fdd_geometry(6&).sides = 2&
    fdd_geometry(6&).sectorsPerTrack = 18&
    fdd_geometry(6&).is525 = 0&

    fdd_geometry(7&).sizeBytes = 2880& * 1024&
    fdd_geometry(7&).tracks = 80&
    fdd_geometry(7&).sides = 2&
    fdd_geometry(7&).sectorsPerTrack = 36&
    fdd_geometry(7&).is525 = 0&

    fdd_geometryInitialized = 1&
End Sub

Private Function FDD_IsValidDrive(ByVal drive As Long) As Boolean
    FDD_IsValidDrive = ((drive >= 0&) And (drive < FDD_NUM))
End Function

Private Function FDD_GetBusyMask(ByVal drive As Long) As Byte
    Select Case drive
        Case 0&: FDD_GetBusyMask = 1&
        Case 1&: FDD_GetBusyMask = 2&
        Case 2&: FDD_GetBusyMask = 4&
        Case Else: FDD_GetBusyMask = 8&
    End Select
End Function

Private Function FDD_MaxTrack(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then
        FDD_MaxTrack = FDD_DEFAULT_MAX_TRACK
        Exit Function
    End If

    If fdd(drive).tracks > 0& Then
        FDD_MaxTrack = fdd(drive).tracks - 1&
    Else
        FDD_MaxTrack = FDD_DEFAULT_MAX_TRACK
    End If
End Function

Private Function FDD_ResolveGeometry(ByVal fileSize As Long, ByRef tracks As Long, ByRef sides As Long, ByRef sectorsPerTrack As Long, ByRef is525 As Byte) As Boolean
    Dim i As Long

    FDD_InitGeometryTable

    For i = 0& To UBound(fdd_geometry)
        If fdd_geometry(i).sizeBytes = fileSize Then
            tracks = fdd_geometry(i).tracks
            sides = fdd_geometry(i).sides
            sectorsPerTrack = fdd_geometry(i).sectorsPerTrack
            is525 = fdd_geometry(i).is525
            FDD_ResolveGeometry = True
            Exit Function
        End If
    Next i

    FDD_ResolveGeometry = False
End Function

Private Sub FDD_ResetReadAddressCursor(ByVal drive As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub

    fdd(drive).readAddressTrack = -1&
    fdd(drive).readAddressHead = -1&
    fdd(drive).readAddressSector = 0&
End Sub

Private Sub FDD_CloseDrive(ByVal drive As Long)
    On Error Resume Next
    If fdd(drive).fileNum <> 0& Then Close #fdd(drive).fileNum
    On Error GoTo 0&
    fdd(drive).fileNum = 0&
End Sub

Private Function FDD_CalcOffset(ByVal drive As Long, ByVal track As Long, ByVal side As Long, ByVal sector As Long) As Long
    Dim logicalSector As Long

    logicalSector = (((track * fdd(drive).sides) + side) * fdd(drive).sectorsPerTrack) + (sector - 1&)
    FDD_CalcOffset = logicalSector * FDD_SECTOR_SIZE
End Function

Private Sub FDD_ClearTrackBuffer(ByVal drive As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub

    fdd(drive).trackBufferValid = 0&
    fdd(drive).trackBufferTrack = -1&
End Sub

Private Sub FDD_FillTrackSide(ByVal drive As Long, ByVal side As Long, ByVal fillValue As Byte)
    Dim i As Long

    For i = 0& To FDD_TRACK_BUFFER_SIZE - 1&
        fdd_trackData(drive, side, i) = fillValue
    Next i
End Sub

Private Function FDD_LoadTrackBuffer(ByVal drive As Long, ByVal track As Long) As Long
    Dim side As Long
    Dim bytesPerSide As Long
    Dim sideBuffer() As Byte
    Dim i As Long

    If Not FDD_IsValidDrive(drive) Then Exit Function
    If fdd(drive).inserted = 0& Then Exit Function
    If fdd(drive).fileNum = 0& Then Exit Function
    If (track < 0&) Or (track >= fdd(drive).tracks) Then Exit Function

    bytesPerSide = fdd(drive).sectorsPerTrack * FDD_SECTOR_SIZE
    If bytesPerSide <= 0& Then Exit Function
    If bytesPerSide > FDD_TRACK_BUFFER_SIZE Then Exit Function

    For side = 0& To FDD_MAX_SIDES - 1&
        Call FDD_FillTrackSide(drive, side, &HF6&)
        If side < fdd(drive).sides Then
            ReDim sideBuffer(0& To bytesPerSide - 1&) As Byte
            On Error GoTo LoadFail
            Get #fdd(drive).fileNum, (FDD_CalcOffset(drive, track, side, 1&) + 1&), sideBuffer
            On Error GoTo 0&
            For i = 0& To bytesPerSide - 1&
                fdd_trackData(drive, side, i) = sideBuffer(i)
            Next i
        End If
    Next side

    fdd(drive).trackBufferTrack = track
    fdd(drive).trackBufferValid = 1&
    FDD_LoadTrackBuffer = 1&
    Exit Function

LoadFail:
    On Error GoTo 0&
    Call FDD_ClearTrackBuffer(drive)
End Function

Private Function FDD_EnsureTrackBuffer(ByVal drive As Long, ByVal track As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function

    If (fdd(drive).trackBufferValid = 0&) Or (fdd(drive).trackBufferTrack <> track) Then
        FDD_EnsureTrackBuffer = FDD_LoadTrackBuffer(drive, track)
    Else
        FDD_EnsureTrackBuffer = 1&
    End If
End Function

Private Function FDD_ValidateCHS(ByVal drive As Long, ByVal track As Long, ByVal side As Long, ByVal sector As Long, ByVal sizeCode As Byte) As Long
    If Not FDD_IsValidDrive(drive) Then
        FDD_ValidateCHS = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).installed = 0& Then
        FDD_ValidateCHS = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).inserted = 0& Then
        FDD_ValidateCHS = FDD_IO_NO_MEDIA
        Exit Function
    End If

    If sizeCode <> 2& Then
        FDD_ValidateCHS = FDD_IO_INVALID_SIZE
        Exit Function
    End If

    If (track < 0&) Or (track >= fdd(drive).tracks) Then
        FDD_ValidateCHS = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If (side < 0&) Or (side >= fdd(drive).sides) Then
        FDD_ValidateCHS = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If (sector <= 0&) Or (sector > fdd(drive).sectorsPerTrack) Then
        FDD_ValidateCHS = FDD_IO_NOT_FOUND
        Exit Function
    End If

    FDD_ValidateCHS = FDD_IO_OK
End Function

Private Function FDD_InterleavedSectorId(ByVal ordinal As Long, ByVal sectorsPerTrack As Long) As Long
    Dim skewedIndex As Long
    Dim adjustedSector As Long
    Dim addValue As Long
    Dim adjustValue As Long

    If sectorsPerTrack <= 0& Then Exit Function

    ordinal = ordinal Mod sectorsPerTrack
    addValue = (sectorsPerTrack And 1&)
    adjustValue = (sectorsPerTrack \ 2&)
    skewedIndex = (ordinal + 1&) Mod sectorsPerTrack
    adjustedSector = (skewedIndex \ 2&) + 1&
    If (skewedIndex And 1&) <> 0& Then
        adjustedSector = adjustedSector + adjustValue + addValue
    End If

    FDD_InterleavedSectorId = adjustedSector
End Function

Private Sub FDD_SetTimerUsec(ByVal timerNum As Long, ByVal usecDelay As Long)
    If timerNum < 0& Then Exit Sub
    timing_updateInterval timerNum, timing_getFreq() * (usecDelay / 1000000)
    timing_timerEnable timerNum
End Sub

Public Sub fdd_init()
    Dim drive As Long

    FDD_InitGeometryTable

    If fdd_initialized = 0& Then
        For drive = 0& To FDD_NUM - 1&
            fdd(drive).installed = 0&
            If drive < FDD_AT_DRIVES Then fdd(drive).installed = 1&
            fdd(drive).pollTimer = timing_addTimer(TIMER_CB_FDD_POLL, drive, FDD_POLL_HZ, TIMING_DISABLED)
            fdd(drive).seekTimer = timing_addTimerUsingInterval(TIMER_CB_FDD_SEEK_COMPLETE, drive, timing_getFreq() * (FDD_SEEK_DELAY_USEC / 1000000), TIMING_DISABLED)
        Next drive
        fdd_initialized = 1&
    End If

    fdd_reset

    For drive = 0& To FDD_AT_DRIVES - 1&
        If LenB(fdd(drive).stagedPath) <> 0& Then
            Call fdd_load(drive, fdd(drive).stagedPath)
        End If
    Next drive
End Sub

Public Sub fdd_reset()
    Dim drive As Long

    For drive = 0& To FDD_NUM - 1&
        If fdd(drive).pollTimer >= 0& Then
            timing_timerDisable fdd(drive).pollTimer
        End If
        If fdd(drive).seekTimer >= 0& Then
            timing_timerDisable fdd(drive).seekTimer
        End If

        fdd(drive).motorEnabled = 0&
        fdd(drive).head = 0&
        fdd(drive).seekInProgress = 0&
        fdd(drive).seekTarget = fdd(drive).track
        Call FDD_ClearTrackBuffer(drive)
        FDD_ResetReadAddressCursor drive

        If fdd(drive).track < 0& Then fdd(drive).track = 0&
        If fdd(drive).track > FDD_MaxTrack(drive) Then fdd(drive).track = FDD_MaxTrack(drive)
    Next drive
End Sub

Public Sub fdd_stageImage(ByVal drive As Long, ByVal filename As String)
    If Not FDD_IsValidDrive(drive) Then Exit Sub

    fdd(drive).stagedPath = filename
    If fdd_initialized <> 0& Then
        Call fdd_load(drive, filename)
    End If
End Sub

Public Function fdd_load(ByVal drive As Long, ByVal filename As String) As Long
    fdd_load = fdd_loadEx(drive, filename, 0&)
End Function

Public Function fdd_loadEx(ByVal drive As Long, ByVal filename As String, ByVal forceWriteProtected As Byte) As Long
    Dim fn As Integer
    Dim fileSize As Long
    Dim tracks As Long
    Dim sides As Long
    Dim sectorsPerTrack As Long
    Dim is525 As Byte
    Dim openedReadOnly As Byte

    If Not FDD_IsValidDrive(drive) Then
        fdd_loadEx = -1&
        Exit Function
    End If

    FDD_CloseDrive drive

    fdd(drive).inserted = 0&
    fdd(drive).writeProtected = 0&
    fdd(drive).tracks = 0&
    fdd(drive).sides = 0&
    fdd(drive).sectorsPerTrack = 0&
    fdd(drive).fileSize = 0&
    fdd(drive).is525 = 0&
    fdd(drive).track = 0&
    fdd(drive).head = 0&
    fdd(drive).seekInProgress = 0&
    fdd(drive).seekTarget = 0&
    Call FDD_ClearTrackBuffer(drive)
    FDD_ResetReadAddressCursor drive
    fdd(drive).stagedPath = vbNullString

    fn = FreeFile
    If forceWriteProtected <> 0& Then
        On Error GoTo LoadFail
        Open filename For Binary Access Read As #fn
        openedReadOnly = 1&
        GoTo Opened
    End If

    On Error GoTo OpenReadOnly
    Open filename For Binary Access Read Write As #fn
    GoTo Opened

OpenReadOnly:
    Err.Clear
    On Error GoTo LoadFail
    Open filename For Binary Access Read As #fn
    openedReadOnly = 1&

Opened:
    On Error GoTo LoadFail

    fileSize = LOF(fn)
    If FDD_ResolveGeometry(fileSize, tracks, sides, sectorsPerTrack, is525) = False Then
        GoTo LoadFail
    End If

    fdd(drive).fileNum = fn
    fdd(drive).inserted = 1&
    fdd(drive).writeProtected = openedReadOnly
    fdd(drive).changed = 1&
    fdd(drive).tracks = tracks
    fdd(drive).sides = sides
    fdd(drive).sectorsPerTrack = sectorsPerTrack
    fdd(drive).fileSize = fileSize
    fdd(drive).is525 = is525
    fdd(drive).stagedPath = filename
    If FDD_LoadTrackBuffer(drive, 0&) = 0& Then GoTo LoadFail

    fdd_loadEx = 0&
    Exit Function

LoadFail:
    On Error Resume Next
    If fn <> 0& Then Close #fn
    On Error GoTo 0&
    fdd(drive).fileNum = 0&
    fdd(drive).inserted = 0&
    fdd(drive).writeProtected = 0&
    fdd(drive).tracks = 0&
    fdd(drive).sides = 0&
    fdd(drive).sectorsPerTrack = 0&
    fdd(drive).fileSize = 0&
    fdd(drive).is525 = 0&
    fdd(drive).changed = 1&
    Call FDD_ClearTrackBuffer(drive)
    FDD_ResetReadAddressCursor drive
    fdd_loadEx = -1&
End Function

Public Sub fdd_eject(ByVal drive As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub

    FDD_CloseDrive drive

    If fdd(drive).pollTimer >= 0& Then timing_timerDisable fdd(drive).pollTimer
    If fdd(drive).seekTimer >= 0& Then timing_timerDisable fdd(drive).seekTimer

    fdd(drive).inserted = 0&
    fdd(drive).writeProtected = 0&
    fdd(drive).changed = 1&
    fdd(drive).motorEnabled = 0&
    fdd(drive).track = 0&
    fdd(drive).head = 0&
    fdd(drive).seekInProgress = 0&
    fdd(drive).seekTarget = 0&
    fdd(drive).tracks = 0&
    fdd(drive).sides = 0&
    fdd(drive).sectorsPerTrack = 0&
    fdd(drive).fileSize = 0&
    fdd(drive).is525 = 0&
    Call FDD_ClearTrackBuffer(drive)
    FDD_ResetReadAddressCursor drive
    fdd(drive).stagedPath = vbNullString
End Sub

Public Function fdd_hasMedia(ByVal drive As Long) As Boolean
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_hasMedia = (fdd(drive).inserted <> 0&)
End Function

Public Function fdd_isWriteProtected(ByVal drive As Long) As Boolean
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_isWriteProtected = (fdd(drive).writeProtected <> 0&)
End Function

Public Function fdd_getChanged(ByVal drive As Long) As Boolean
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_getChanged = (fdd(drive).changed <> 0&)
End Function

Public Sub fdd_clearChanged(ByVal drive As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub
    fdd(drive).changed = 0&
End Sub

Public Function fdd_get_flags(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_get_flags = fdd(drive).installed
End Function

Public Function fdd_getCurrentTrack(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_getCurrentTrack = fdd(drive).track
End Function

Public Sub fdd_setMotorEnable(ByVal drive As Long, ByVal enabled As Byte)
    If Not FDD_IsValidDrive(drive) Then Exit Sub
    If fdd(drive).installed = 0& Then Exit Sub

    fdd(drive).motorEnabled = enabled
    If fdd(drive).pollTimer < 0& Then Exit Sub

    If enabled <> 0& Then
        timing_timerEnable fdd(drive).pollTimer
    Else
        timing_timerDisable fdd(drive).pollTimer
    End If
End Sub

Public Sub fdd_setHead(ByVal drive As Long, ByVal head As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub

    If (head <= 0&) Or (fdd_is_double_sided(drive) = 0&) Then
        fdd(drive).head = 0&
    Else
        fdd(drive).head = 1&
    End If

    If fdd(drive).readAddressHead <> fdd(drive).head Then
        FDD_ResetReadAddressCursor drive
    End If
End Sub

Public Function fdd_getHead(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    If fdd_is_double_sided(drive) = 0& Then Exit Function
    fdd_getHead = fdd(drive).head
End Function

Public Function fdd_track0(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_track0 = IIf(fdd(drive).track = 0&, 1&, 0&)
End Function

Public Function fdd_is_double_sided(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    If fdd(drive).inserted = 0& Then
        fdd_is_double_sided = IIf(fdd(drive).installed <> 0&, 1&, 0&)
    Else
        fdd_is_double_sided = IIf(fdd(drive).sides > 1&, 1&, 0&)
    End If
End Function

Public Function fdd_is_525(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    fdd_is_525 = fdd(drive).is525
End Function

Public Function fdd_is_dd(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    If fdd(drive).inserted = 0& Then Exit Function
    fdd_is_dd = IIf((fdd(drive).sectorsPerTrack <= 9&), 1&, 0&)
End Function

Public Function fdd_is_hd(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    If fdd(drive).inserted = 0& Then Exit Function
    fdd_is_hd = IIf((fdd(drive).sectorsPerTrack = 15&) Or (fdd(drive).sectorsPerTrack = 18&), 1&, 0&)
End Function

Public Function fdd_is_ed(ByVal drive As Long) As Long
    If Not FDD_IsValidDrive(drive) Then Exit Function
    If fdd(drive).inserted = 0& Then Exit Function
    fdd_is_ed = IIf(fdd(drive).sectorsPerTrack = 36&, 1&, 0&)
End Function

Public Sub fdd_seek(ByVal drive As Long, ByVal trackDiff As Long)
    Dim maxTrack As Long
    Dim oldTrack As Long
    Dim targetTrack As Long

    If Not FDD_IsValidDrive(drive) Then Exit Sub
    If fdd(drive).installed = 0& Then Exit Sub
    If trackDiff = 0& Then Exit Sub

    maxTrack = FDD_MaxTrack(drive)
    If fdd(drive).seekInProgress <> 0& Then Exit Sub

    fdd(drive).changed = 0&
    oldTrack = fdd(drive).track
    targetTrack = oldTrack + trackDiff
    If targetTrack < 0& Then targetTrack = 0&
    If targetTrack > maxTrack Then targetTrack = maxTrack

    fdd(drive).track = targetTrack
    fdd(drive).seekInProgress = 1&
    fdd(drive).seekTarget = targetTrack
    Call FDD_ClearTrackBuffer(drive)

    If fdd(drive).seekTimer >= 0& Then
        FDD_SetTimerUsec fdd(drive).seekTimer, FDD_SEEK_DELAY_USEC
    Else
        fdd_seekCompleteCallback drive
    End If
End Sub

Public Sub fdd_pollCallback(ByVal drive As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub
    If fdd(drive).motorEnabled = 0& Then
        If fdd(drive).pollTimer >= 0& Then timing_timerDisable fdd(drive).pollTimer
    End If
End Sub

Public Sub fdd_seekCompleteCallback(ByVal drive As Long)
    If Not FDD_IsValidDrive(drive) Then Exit Sub

    If fdd(drive).seekTimer >= 0& Then timing_timerDisable fdd(drive).seekTimer

    fdd(drive).track = fdd(drive).seekTarget
    fdd(drive).seekInProgress = 0&
    Call FDD_LoadTrackBuffer(drive, fdd(drive).track)
    FDD_ResetReadAddressCursor drive
    fdc_seek_complete_interrupt drive
End Sub

Public Function fdd_readSector(ByVal drive As Long, ByVal sector As Long, ByVal track As Long, ByVal side As Long, ByVal sizeCode As Byte, ByRef buffer() As Byte) As Long
    Dim status As Long
    Dim startPos As Long
    Dim i As Long

    status = FDD_ValidateCHS(drive, track, side, sector, sizeCode)
    If status <> FDD_IO_OK Then
        fdd_readSector = status
        Exit Function
    End If

    If fdd(drive).seekInProgress <> 0& Then
        fdd_readSector = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If FDD_EnsureTrackBuffer(drive, track) = 0& Then
        fdd_readSector = FDD_IO_NOT_FOUND
        Exit Function
    End If

    startPos = (sector - 1&) * FDD_SECTOR_SIZE
    For i = 0& To FDD_SECTOR_SIZE - 1&
        buffer(i) = fdd_trackData(drive, side, startPos + i)
    Next i
    fdd_readSector = FDD_IO_OK
    Exit Function
End Function

Public Function fdd_writeSector(ByVal drive As Long, ByVal sector As Long, ByVal track As Long, ByVal side As Long, ByVal sizeCode As Byte, ByRef buffer() As Byte) As Long
    Dim status As Long
    Dim startPos As Long
    Dim i As Long

    status = FDD_ValidateCHS(drive, track, side, sector, sizeCode)
    If status <> FDD_IO_OK Then
        fdd_writeSector = status
        Exit Function
    End If

    If fdd(drive).seekInProgress <> 0& Then
        fdd_writeSector = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).writeProtected <> 0& Then
        fdd_writeSector = FDD_IO_WRITE_PROTECT
        Exit Function
    End If

    If FDD_EnsureTrackBuffer(drive, track) = 0& Then
        fdd_writeSector = FDD_IO_NOT_FOUND
        Exit Function
    End If

    startPos = (sector - 1&) * FDD_SECTOR_SIZE
    For i = 0& To FDD_SECTOR_SIZE - 1&
        fdd_trackData(drive, side, startPos + i) = buffer(i)
    Next i

    On Error GoTo WriteFail
    Put #fdd(drive).fileNum, (FDD_CalcOffset(drive, track, side, sector) + 1&), buffer
    fdd_writeSector = FDD_IO_OK
    Exit Function

WriteFail:
    fdd_writeSector = FDD_IO_NOT_FOUND
End Function

Public Function fdd_readAddress(ByVal drive As Long, ByVal side As Long, ByRef trackOut As Byte, ByRef headOut As Byte, ByRef sectorOut As Byte, ByRef sizeOut As Byte) As Long
    If Not FDD_IsValidDrive(drive) Then
        fdd_readAddress = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).inserted = 0& Then
        fdd_readAddress = FDD_IO_NO_MEDIA
        Exit Function
    End If

    If (side < 0&) Or (side >= fdd(drive).sides) Then
        fdd_readAddress = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).sectorsPerTrack <= 0& Then
        fdd_readAddress = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).seekInProgress <> 0& Then
        fdd_readAddress = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If FDD_EnsureTrackBuffer(drive, fdd(drive).track) = 0& Then
        fdd_readAddress = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If (fdd(drive).readAddressTrack <> fdd(drive).track) Or (fdd(drive).readAddressHead <> side) Then
        fdd(drive).readAddressTrack = fdd(drive).track
        fdd(drive).readAddressHead = side
        fdd(drive).readAddressSector = 0&
    Else
        fdd(drive).readAddressSector = fdd(drive).readAddressSector + 1&
        If fdd(drive).readAddressSector >= fdd(drive).sectorsPerTrack Then
            fdd(drive).readAddressSector = 0&
        End If
    End If

    trackOut = CByte(fdd(drive).track And &HFF&)
    headOut = CByte(side And &HFF&)
    sectorOut = CByte(FDD_InterleavedSectorId(fdd(drive).readAddressSector, fdd(drive).sectorsPerTrack) And &HFF&)
    sizeOut = 2&
    fdd_readAddress = FDD_IO_OK
End Function

Public Function fdd_formatTrack(ByVal drive As Long, ByVal track As Long, ByVal side As Long, ByVal fillByte As Byte, ByVal sectorCount As Long, ByRef sectorIds() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim sectorId As Long
    Dim sizeCode As Byte
    Dim fillBuffer(0& To FDD_SECTOR_SIZE - 1&) As Byte

    If Not FDD_IsValidDrive(drive) Then
        fdd_formatTrack = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).inserted = 0& Then
        fdd_formatTrack = FDD_IO_NO_MEDIA
        Exit Function
    End If

    If fdd(drive).writeProtected <> 0& Then
        fdd_formatTrack = FDD_IO_WRITE_PROTECT
        Exit Function
    End If

    If (track < 0&) Or (track >= fdd(drive).tracks) Then
        fdd_formatTrack = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If (side < 0&) Or (side >= fdd(drive).sides) Then
        fdd_formatTrack = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If sectorCount <= 0& Then
        fdd_formatTrack = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If fdd(drive).seekInProgress <> 0& Then
        fdd_formatTrack = FDD_IO_NOT_FOUND
        Exit Function
    End If

    If FDD_EnsureTrackBuffer(drive, track) = 0& Then
        fdd_formatTrack = FDD_IO_NOT_FOUND
        Exit Function
    End If

    For i = 0& To FDD_SECTOR_SIZE - 1&
        fillBuffer(i) = fillByte
    Next i

    For i = 0& To sectorCount - 1&
        sectorId = sectorIds((i * 4&) + 2&)
        sizeCode = sectorIds((i * 4&) + 3&)

        If sizeCode <> 2& Then
            fdd_formatTrack = FDD_IO_INVALID_SIZE
            Exit Function
        End If

        If (sectorId <= 0&) Or (sectorId > fdd(drive).sectorsPerTrack) Then
            fdd_formatTrack = FDD_IO_NOT_FOUND
            Exit Function
        End If

        For j = 0& To FDD_SECTOR_SIZE - 1&
            fdd_trackData(drive, side, ((sectorId - 1&) * FDD_SECTOR_SIZE) + j) = fillByte
        Next j

        On Error GoTo FormatFail
        Put #fdd(drive).fileNum, (FDD_CalcOffset(drive, track, side, sectorId) + 1&), fillBuffer
        On Error GoTo 0&
    Next i

    fdd_formatTrack = FDD_IO_OK
    Exit Function

FormatFail:
    fdd_formatTrack = FDD_IO_NOT_FOUND
End Function

Public Function fdd_get_inserted_count() As Byte
    Dim drive As Long
    Dim count As Long

    For drive = 0& To FDD_AT_DRIVES - 1&
        If fdd(drive).inserted <> 0& Then count = count + 1&
    Next drive

    fdd_get_inserted_count = CByte(count And &HFF&)
End Function
