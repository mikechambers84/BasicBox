Attribute VB_Name = "modAudio"
Option Explicit

Private Declare Function waveOutOpen Lib "winmm.dll" (ByRef lphWaveOut As Long, ByVal uDeviceID As Long, ByRef lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Type DSBUFFERDESC
    dwSize As Long
    dwFlags As Long
    dwBufferBytes As Long
    dwReserved As Long
    lpwfxFormat As Long
End Type

Private Const MMSYSERR_NOERROR As Long = 0&
Private Const CALLBACK_NULL As Long = 0&
Private Const WAVE_MAPPER As Long = -1&
Private Const WAVE_FORMAT_PCM As Integer = 1%

Private Const WHDR_DONE As Long = &H1&
Private Const WHDR_PREPARED As Long = &H2&

Private Const DSSCL_PRIORITY As Long = 2&
Private Const DSBCAPS_STICKYFOCUS As Long = &H4000&
Private Const DSBCAPS_GETCURRENTPOSITION2 As Long = &H10000
Private Const DSBPLAY_LOOPING As Long = &H1&

Private Const AUDIO_WINMM_BUFFERS As Long = 6&
Private Const AUDIO_WINMM_BUFFER_SAMPLES As Long = 1024&
Private Const AUDIO_WINMM_BUFFER_BYTES As Long = AUDIO_WINMM_BUFFER_SAMPLES * 2&
Private Const AUDIO_WINMM_TOTAL_SAMPLES As Long = AUDIO_WINMM_BUFFERS * AUDIO_WINMM_BUFFER_SAMPLES
Private Const AUDIO_DS_BUFFER_SAMPLES As Long = 8192&
Private Const AUDIO_DS_BUFFER_BYTES As Long = AUDIO_DS_BUFFER_SAMPLES * 2&
Private Const AUDIO_DS_TARGET_LEAD_BYTES As Long = 4096&

Private audio_buffer(0& To SAMPLE_BUFFER - 1&) As Integer
Private audio_bufferpos As Long
Private audio_rateFast As Double
Private audio_timer As Long
Private audio_updateTiming As Byte

Private audio_firstfill As Byte
Private audio_timeIdx As Byte
Private audio_genSampRate As Double
Private audio_genInterval As Double
Private audio_paused As Byte
Private audio_cbTime(0& To 9&) As Double

Private audio_waveOut As Long
Private audio_waveHdr(0& To AUDIO_WINMM_BUFFERS - 1&) As WAVEHDR
Private audio_waveData(0& To AUDIO_WINMM_TOTAL_SAMPLES - 1&) As Integer
Private audio_waveReady As Byte
Private audio_waveErrLogged As Byte
Private audio_useDirectSound As Byte
Private audio_ds As Long
Private audio_dsBuffer As Long
Private audio_dsWriteCursor As Long
Private audio_dsReady As Byte
Private audio_dsErrLogged As Byte

Private Function audio_backendInit() As Long
    audio_backendInit = -1&

    If audio_dsReady <> 0& Then
        audio_backendInit = 0&
        Exit Function
    End If

    If audio_waveReady <> 0& Then
        audio_backendInit = 0&
        Exit Function
    End If

    If audio_dsBackendInit() = 0& Then
        audio_backendInit = 0&
        Exit Function
    End If

    If audio_waveBackendInit() = 0& Then
        audio_backendInit = 0&
        Exit Function
    End If
End Function

Private Function audio_dsBackendInit() As Long
    Dim fmt As WAVEFORMATEX
    Dim desc As DSBUFFERDESC
    Dim hr As Long

    audio_dsBackendInit = -1&
    audio_useDirectSound = 0&
    audio_ds = 0&
    audio_dsBuffer = 0&
    audio_dsWriteCursor = 0&
    audio_dsReady = 0&
    audio_dsErrLogged = 0&

    fmt.wFormatTag = WAVE_FORMAT_PCM
    fmt.nChannels = 1%
    fmt.nSamplesPerSec = SAMPLE_RATE
    fmt.wBitsPerSample = 16%
    fmt.nBlockAlign = CInt((CLng(fmt.nChannels) * CLng(fmt.wBitsPerSample)) \ 8&)
    fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * CLng(fmt.nBlockAlign)
    fmt.cbSize = 0%

    hr = DirectSoundCreate(0&, audio_ds, 0&)
    If dxHrFailed(hr) Or (audio_ds = 0&) Then Exit Function

    hr = dxCallLong(audio_ds, IDX_IDIRECTSOUND_SETCOOPERATIVELEVEL, frmConsole.hWnd, DSSCL_PRIORITY)
    If dxHrFailed(hr) Then
        Call audio_dsBackendShutdown
        Exit Function
    End If

    dxZeroMemory VarPtr(desc), LenB(desc)
    desc.dwSize = LenB(desc)
    desc.dwFlags = DSBCAPS_STICKYFOCUS Or DSBCAPS_GETCURRENTPOSITION2
    desc.dwBufferBytes = AUDIO_DS_BUFFER_BYTES
    desc.lpwfxFormat = VarPtr(fmt)

    hr = dxCallLong(audio_ds, IDX_IDIRECTSOUND_CREATESOUNDBUFFER, VarPtr(desc), VarPtr(audio_dsBuffer), 0&)
    If dxHrFailed(hr) Or (audio_dsBuffer = 0&) Then
        Call audio_dsBackendShutdown
        Exit Function
    End If

    Call audio_dsWriteRegion(0&, AUDIO_DS_TARGET_LEAD_BYTES)
    audio_dsWriteCursor = AUDIO_DS_TARGET_LEAD_BYTES Mod AUDIO_DS_BUFFER_BYTES

    hr = dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_PLAY, 0&, 0&, DSBPLAY_LOOPING)
    If dxHrFailed(hr) Then
        Call audio_dsBackendShutdown
        Exit Function
    End If

    audio_useDirectSound = 1&
    audio_dsReady = 1&
    audio_dsBackendInit = 0&
End Function

Private Function audio_waveBackendInit() As Long
    Dim fmt As WAVEFORMATEX
    Dim i As Long
    Dim rc As Long
    Dim baseIdx As Long

    audio_waveBackendInit = -1&

    audio_waveOut = 0&
    audio_waveErrLogged = 0&

    fmt.wFormatTag = WAVE_FORMAT_PCM
    fmt.nChannels = 1%
    fmt.nSamplesPerSec = SAMPLE_RATE
    fmt.wBitsPerSample = 16%
    fmt.nBlockAlign = CInt((CLng(fmt.nChannels) * CLng(fmt.wBitsPerSample)) \ 8&)
    fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * CLng(fmt.nBlockAlign)
    fmt.cbSize = 0%

    rc = waveOutOpen(audio_waveOut, WAVE_MAPPER, fmt, 0&, 0&, CALLBACK_NULL)
    If (rc <> MMSYSERR_NOERROR) Or (audio_waveOut = 0&) Then
        debug_log DEBUG_ERROR, "[AUDIO] DirectSound and waveOut initialization failed"
        Exit Function
    End If

    For i = 0& To AUDIO_WINMM_BUFFERS - 1&
        baseIdx = i * AUDIO_WINMM_BUFFER_SAMPLES
        audio_waveHdr(i).lpData = VarPtr(audio_waveData(baseIdx))
        audio_waveHdr(i).dwBufferLength = AUDIO_WINMM_BUFFER_BYTES
        audio_waveHdr(i).dwBytesRecorded = 0&
        audio_waveHdr(i).dwUser = 0&
        audio_waveHdr(i).dwFlags = 0&
        audio_waveHdr(i).dwLoops = 0&
        audio_waveHdr(i).lpNext = 0&
        audio_waveHdr(i).Reserved = 0&

        rc = waveOutPrepareHeader(audio_waveOut, audio_waveHdr(i), LenB(audio_waveHdr(i)))
        If rc <> MMSYSERR_NOERROR Then
            debug_log DEBUG_ERROR, "[AUDIO] waveOutPrepareHeader failed"
            Call audio_waveBackendShutdown
            Exit Function
        End If

        Call audio_waveBackendFillBuffer(i)
        rc = waveOutWrite(audio_waveOut, audio_waveHdr(i), LenB(audio_waveHdr(i)))
        If rc <> MMSYSERR_NOERROR Then
            debug_log DEBUG_ERROR, "[AUDIO] waveOutWrite failed during init"
            Call audio_waveBackendShutdown
            Exit Function
        End If
    Next i

    audio_waveReady = 1&
    audio_waveBackendInit = 0&
End Function

Private Sub audio_waveBackendFillBuffer(ByVal bufferIdx As Long)
    Dim tmp(0& To AUDIO_WINMM_BUFFER_SAMPLES - 1&) As Integer
    Dim dstBase As Long

    audio_pullSamples tmp, AUDIO_WINMM_BUFFER_SAMPLES
    dstBase = bufferIdx * AUDIO_WINMM_BUFFER_SAMPLES
    CopyMemory audio_waveData(dstBase), tmp(0&), AUDIO_WINMM_BUFFER_BYTES
End Sub

Private Sub audio_backendPump()
    Dim i As Long
    Dim rc As Long

    If audio_dsReady <> 0& Then
        Call audio_dsBackendPump
        Exit Sub
    End If

    If audio_waveReady = 0& Then Exit Sub

    For i = 0& To AUDIO_WINMM_BUFFERS - 1&
        If (audio_waveHdr(i).dwFlags And WHDR_DONE) <> 0& Then
            Call audio_waveBackendFillBuffer(i)
            audio_waveHdr(i).dwFlags = (audio_waveHdr(i).dwFlags And (Not WHDR_DONE))
            rc = waveOutWrite(audio_waveOut, audio_waveHdr(i), LenB(audio_waveHdr(i)))
            If (rc <> MMSYSERR_NOERROR) And (audio_waveErrLogged = 0&) Then
                audio_waveErrLogged = 1&
                debug_log DEBUG_ERROR, "[AUDIO] waveOutWrite failed while streaming"
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub audio_dsBackendPump()
    Dim playCursor As Long
    Dim writeCursor As Long
    Dim desiredWrite As Long
    Dim bytesToWrite As Long
    Dim hr As Long

    If audio_dsReady = 0& Then Exit Sub

    hr = dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_GETCURRENTPOSITION, VarPtr(playCursor), VarPtr(writeCursor))
    If dxHrFailed(hr) Then
        Call dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_RESTORE)
        hr = dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_GETCURRENTPOSITION, VarPtr(playCursor), VarPtr(writeCursor))
        If dxHrFailed(hr) Then Exit Sub
    End If

    desiredWrite = playCursor + AUDIO_DS_TARGET_LEAD_BYTES
    If desiredWrite >= AUDIO_DS_BUFFER_BYTES Then desiredWrite = desiredWrite - AUDIO_DS_BUFFER_BYTES

    bytesToWrite = audio_dsCircularDistance(audio_dsWriteCursor, desiredWrite, AUDIO_DS_BUFFER_BYTES)
    If bytesToWrite <= 0& Then Exit Sub

    Call audio_dsWriteRegion(audio_dsWriteCursor, bytesToWrite)
    audio_dsWriteCursor = (audio_dsWriteCursor + bytesToWrite) Mod AUDIO_DS_BUFFER_BYTES
End Sub

Private Sub audio_backendShutdown()
    If audio_dsReady <> 0& Or audio_dsBuffer <> 0& Or audio_ds <> 0& Then
        Call audio_dsBackendShutdown
    End If

    If audio_waveReady <> 0& Or audio_waveOut <> 0& Then
        Call audio_waveBackendShutdown
    End If
End Sub

Private Sub audio_dsBackendShutdown()
    If audio_dsBuffer <> 0& Then
        Call dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_STOP)
    End If

    Call dxRelease(audio_dsBuffer)
    Call dxRelease(audio_ds)
    audio_useDirectSound = 0&
    audio_dsReady = 0&
    audio_dsWriteCursor = 0&
    audio_dsErrLogged = 0&
End Sub

Private Sub audio_waveBackendShutdown()
    Dim i As Long

    If audio_waveOut <> 0& Then
        waveOutReset audio_waveOut

        For i = 0& To AUDIO_WINMM_BUFFERS - 1&
            If (audio_waveHdr(i).dwFlags And WHDR_PREPARED) <> 0& Then
                waveOutUnprepareHeader audio_waveOut, audio_waveHdr(i), LenB(audio_waveHdr(i))
            End If
        Next i

        waveOutClose audio_waveOut
    End If

    audio_waveOut = 0&
    audio_waveReady = 0&
    audio_waveErrLogged = 0&
End Sub

Private Sub audio_dsWriteRegion(ByVal offsetBytes As Long, ByVal byteLen As Long)
    Dim ptr1 As Long
    Dim ptr2 As Long
    Dim bytes1 As Long
    Dim bytes2 As Long
    Dim hr As Long

    If (audio_dsBuffer = 0&) Or (byteLen <= 0&) Then Exit Sub

    hr = dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_LOCK, offsetBytes, byteLen, VarPtr(ptr1), VarPtr(bytes1), VarPtr(ptr2), VarPtr(bytes2), 0&)
    If dxHrFailed(hr) Then
        Call dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_RESTORE)
        hr = dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_LOCK, offsetBytes, byteLen, VarPtr(ptr1), VarPtr(bytes1), VarPtr(ptr2), VarPtr(bytes2), 0&)
        If dxHrFailed(hr) Then Exit Sub
    End If

    Call audio_dsFillPtr(ptr1, bytes1)
    Call audio_dsFillPtr(ptr2, bytes2)
    Call dxCallLong(audio_dsBuffer, IDX_IDIRECTSOUNDBUFFER_UNLOCK, ptr1, bytes1, ptr2, bytes2)
End Sub

Private Sub audio_dsFillPtr(ByVal dstPtr As Long, ByVal byteLen As Long)
    Dim tmp() As Integer
    Dim sampleCount As Long

    If (dstPtr = 0&) Or (byteLen <= 0&) Then Exit Sub

    sampleCount = byteLen \ 2&
    If sampleCount <= 0& Then Exit Sub

    ReDim tmp(0& To sampleCount - 1&) As Integer
    audio_pullSamples tmp, sampleCount
    dxCopyMemory dstPtr, VarPtr(tmp(0&)), byteLen
End Sub

Private Function audio_dsCircularDistance(ByVal startPos As Long, ByVal endPos As Long, ByVal sizeBytes As Long) As Long
    If endPos >= startPos Then
        audio_dsCircularDistance = endPos - startPos
    Else
        audio_dsCircularDistance = (sizeBytes - startPos) + endPos
    End If
End Function

Public Function audio_init(ByRef useMachine As MACHINE_t) As Long
    Dim i As Long

    audio_bufferpos = 0&
    audio_rateFast = CDbl(SAMPLE_RATE) * 1.01#
    audio_updateTiming = 0&
    audio_firstfill = 1&
    audio_timeIdx = 0&
    audio_genSampRate = SAMPLE_RATE
    audio_genInterval = 0#
    audio_paused = 1&
    audio_useDirectSound = 0&
    audio_dsReady = 0&
    audio_dsWriteCursor = 0&
    audio_dsErrLogged = 0&
    audio_waveReady = 0&
    audio_waveErrLogged = 0&

    For i = 0& To SAMPLE_BUFFER - 1&
        audio_buffer(i) = 0&
    Next i
    For i = 0& To 9&
        audio_cbTime(i) = 0#
    Next i

    audio_timer = timing_addTimer(TIMER_CB_AUDIO_GENERATE, 0&, SAMPLE_RATE, TIMING_ENABLED)
    If audio_timer = TIMING_ERROR Then
        audio_init = -1&
        Exit Function
    End If

    If audio_backendInit() <> 0& Then
        timing_timerDisable audio_timer
        audio_init = -1&
        Exit Function
    End If

    audio_init = 0&
End Function

Public Sub audio_updateSampleTiming()
    If audio_updateTiming = AUDIO_TIMING_FAST Then
        timing_updateIntervalFreq audio_timer, audio_rateFast
    ElseIf audio_updateTiming = AUDIO_TIMING_NORMAL Then
        timing_updateIntervalFreq audio_timer, SAMPLE_RATE
        audio_paused = 0&
    End If

    audio_updateTiming = 0&
    audio_backendPump
End Sub

Public Sub audio_generateSample(ByVal dummy As Long)
    Dim val As Long
    Dim OPLsample(0& To 1&) As Integer

    val = pcspeaker_getSample(0&) \ 3&

    If machine.mixOPL <> 0& Then
        OPL3_GenerateStream 0&, OPLsample, 1&
        val = val + (OPLsample(0&) \ 2&)
    End If

    If machine.mixBlaster <> 0& Then
        val = val + (blaster_getSample(0&) \ 3&)
    End If

    audio_bufferSample CInt(val)
End Sub

Private Sub audio_bufferSample(ByVal val As Integer)
    If audio_bufferpos = SAMPLE_BUFFER Then
        Exit Sub
    End If

    audio_buffer(audio_bufferpos) = val
    audio_bufferpos = audio_bufferpos + 1&

    If audio_bufferpos < CLng(CDbl(SAMPLE_BUFFER) * 0.5#) Then
        audio_updateTiming = AUDIO_TIMING_FAST
    ElseIf audio_bufferpos >= CLng(CDbl(SAMPLE_BUFFER) * 0.75#) Then
        audio_updateTiming = AUDIO_TIMING_NORMAL
    End If

    If audio_bufferpos = SAMPLE_BUFFER Then
        timing_timerDisable audio_timer
    End If
End Sub

Private Sub audio_moveBuffer(ByRef dst() As Integer, ByVal byteLen As Long)
    Dim i As Long
    Dim samplesToMove As Long
    Dim dstLow As Long
    Dim dstHigh As Long
    Dim dstCount As Long

    dstLow = LBound(dst)
    dstHigh = UBound(dst)
    dstCount = (dstHigh - dstLow) + 1&

    For i = dstLow To dstHigh
        dst(i) = 0&
    Next i

    If audio_bufferpos < CLng(CDbl(SAMPLE_BUFFER) * 0.75#) Then
        timing_timerEnable audio_timer
    End If

    If (audio_bufferpos * 2&) < byteLen Then
        audio_paused = 1&
        Exit Sub
    End If

    samplesToMove = (byteLen \ 2&)
    If samplesToMove > dstCount Then samplesToMove = dstCount
    If samplesToMove > audio_bufferpos Then samplesToMove = audio_bufferpos

    For i = 0& To samplesToMove - 1&
        dst(dstLow + i) = audio_buffer(i)
    Next i

    For i = samplesToMove To audio_bufferpos - 1&
        audio_buffer(i - samplesToMove) = audio_buffer(i)
    Next i

    audio_bufferpos = audio_bufferpos - samplesToMove
End Sub

Public Sub audio_fill(ByRef stream() As Integer, ByVal byteLen As Long)
    audio_moveBuffer stream, byteLen
End Sub

Public Sub audio_pullSamples(ByRef dst() As Integer, ByVal sampleCount As Long)
    If sampleCount <= 0& Then Exit Sub
    audio_moveBuffer dst, (sampleCount * 2&)
End Sub

Public Sub audio_shutdown()
    audio_backendShutdown
End Sub

