Attribute VB_Name = "modPCap"
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal cbCopy As Long)
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByVal pvargResult As Long) As Long

Public Const PCAP_IF_USERNET As Long = -2&

Private Const PCAP_BACKEND_NONE As Long = 0&
Private Const PCAP_BACKEND_USERNET As Long = 1&
Private Const PCAP_BACKEND_WPCAP As Long = 2&

Private Const CC_CDECL As Long = 1&
Private Const VT_EMPTY As Integer = 0
Private Const VT_I4 As Integer = 3
Private Const PCAP_ERRBUF_SIZE As Long = 256&
Private Const PCAP_RX_QUEUE As Long = 512&
Private Const PCAP_SNAPLEN As Long = 65535
Private Const PCAP_OPEN_PROMISC As Long = 1&
Private Const PCAP_OPEN_TIMEOUT_MS As Long = 1000&
Private Const PCAP_POLL_BUDGET As Long = 256&
Private Const PCAP_POLL_HZ As Double = 2000#

Private Type PCAP_IF_t
    nextPtr As Long
    namePtr As Long
    descriptionPtr As Long
    addressesPtr As Long
    flags As Long
End Type

Private Type PCAP_PKTHDR_t
    ts_sec As Long
    ts_usec As Long
    caplen As Long
    packetLen As Long
End Type

Private pcap_warned As Byte
Private pcap_backend As Long
Private pcap_timer As Long
Private pcap_timerActive As Byte

Private pcap_hModule As Long
Private pcap_proc_findalldevs As Long
Private pcap_proc_freealldevs As Long
Private pcap_proc_open_live As Long
Private pcap_proc_setnonblock As Long
Private pcap_proc_setmintocopy As Long
Private pcap_proc_next_ex As Long
Private pcap_proc_sendpacket As Long
Private pcap_proc_close As Long
Private pcap_proc_geterr As Long

Private pcap_handle As Long
Private pcap_devId As Long
Private pcap_rxFrames(0& To PCAP_RX_QUEUE - 1&) As Variant
Private pcap_rxLen(0& To PCAP_RX_QUEUE - 1&) As Long
Private pcap_rxHead As Long
Private pcap_rxTail As Long
Private pcap_rxCount As Long

Private Sub pcap_logUnavailable(ByVal message As String)
    If pcap_warned = 0& Then
        debug_log DEBUG_ERROR, message
        pcap_warned = 1&
    End If
End Sub

Private Sub Pcap_ResetQueue()
    Dim i As Long

    pcap_rxHead = 0&
    pcap_rxTail = 0&
    pcap_rxCount = 0&

    For i = 0& To PCAP_RX_QUEUE - 1&
        pcap_rxFrames(i) = Empty
        pcap_rxLen(i) = 0&
    Next i
End Sub

Private Sub Pcap_ResetState()
    pcap_handle = 0&
    pcap_devId = -1&
    Pcap_ResetQueue
    pcap_backend = PCAP_BACKEND_NONE
End Sub

Private Function Pcap_QueueFrame(ByRef frame() As Byte, ByVal frameLen As Long) As Long
    Dim slot As Long
    Dim copyBuf() As Byte

    If frameLen <= 0& Then
        Pcap_QueueFrame = 0&
        Exit Function
    End If

    If pcap_rxCount >= PCAP_RX_QUEUE Then
        debug_log DEBUG_INFO, "[PCAP] RX queue overflow; dropping packet"
        Pcap_QueueFrame = -1&
        Exit Function
    End If

    ReDim copyBuf(0& To frameLen - 1&) As Byte
    CopyMemory copyBuf(0&), VarPtr(frame(0&)), frameLen

    slot = pcap_rxTail
    pcap_rxFrames(slot) = copyBuf
    pcap_rxLen(slot) = frameLen
    pcap_rxTail = (pcap_rxTail + 1&) Mod PCAP_RX_QUEUE
    pcap_rxCount = pcap_rxCount + 1&

    Pcap_QueueFrame = 0&
End Function

Private Sub Pcap_DequeueFrame()
    If pcap_rxCount <= 0& Then Exit Sub

    pcap_rxFrames(pcap_rxHead) = Empty
    pcap_rxLen(pcap_rxHead) = 0&
    pcap_rxHead = (pcap_rxHead + 1&) Mod PCAP_RX_QUEUE
    pcap_rxCount = pcap_rxCount - 1&
End Sub

Private Function Pcap_DrainRxQueue() As Long
    Dim frame As Variant
    Dim pkt() As Byte
    Dim frameLen As Long
    Dim rxResult As Long

    If (pcap_backend <> PCAP_BACKEND_WPCAP) Or (pcap_devId < 0&) Then Exit Function

    Do While pcap_rxCount > 0&
        frame = pcap_rxFrames(pcap_rxHead)
        frameLen = pcap_rxLen(pcap_rxHead)

        If IsArray(frame) Then
            pkt = frame
            rxResult = ne2000_rx_frame_try(pcap_devId, pkt, frameLen)

            If rxResult = NE2000_RX_ACCEPTED Then
                Pcap_DequeueFrame
                Pcap_DrainRxQueue = Pcap_DrainRxQueue + 1&
            ElseIf rxResult = NE2000_RX_DROP Then
                Pcap_DequeueFrame
                Pcap_DrainRxQueue = Pcap_DrainRxQueue + 1&
            Else
                Exit Do
            End If
        Else
            Pcap_DequeueFrame
            Pcap_DrainRxQueue = Pcap_DrainRxQueue + 1&
        End If
    Loop
End Function

Private Function Pcap_CStringFromPtr(ByVal strPtr As Long) As String
    Dim strLen As Long
    Dim buf() As Byte

    If strPtr = 0& Then Exit Function

    strLen = lstrlenA(strPtr)
    If strLen <= 0& Then Exit Function

    ReDim buf(0& To strLen - 1&) As Byte
    CopyMemory buf(0&), strPtr, strLen
    Pcap_CStringFromPtr = StrConv(buf, vbUnicode)
End Function

Private Function Pcap_CStringFromBytes(ByRef buf() As Byte) As String
    Dim i As Long
    Dim textLen As Long
    Dim tmp() As Byte

    For i = LBound(buf) To UBound(buf)
        If buf(i) = 0& Then Exit For
        textLen = textLen + 1&
    Next i

    If textLen <= 0& Then Exit Function

    ReDim tmp(0& To textLen - 1&) As Byte
    CopyMemory tmp(0&), VarPtr(buf(LBound(buf))), textLen
    Pcap_CStringFromBytes = StrConv(tmp, vbUnicode)
End Function

Private Function Pcap_StringToAnsiZ(ByVal text As String) As Byte()
    Dim tmp() As Byte
    Dim outBuf() As Byte
    Dim i As Long

    If LenB(text) = 0& Then
        ReDim outBuf(0& To 0&) As Byte
        outBuf(0&) = 0&
        Pcap_StringToAnsiZ = outBuf
        Exit Function
    End If

    tmp = StrConv(text, vbFromUnicode)
    ReDim outBuf(0& To UBound(tmp) + 1&) As Byte

    For i = 0& To UBound(tmp)
        outBuf(i) = tmp(i)
    Next i

    outBuf(UBound(outBuf)) = 0&
    Pcap_StringToAnsiZ = outBuf
End Function

' DispCallFunc wants pointers to VARIANTARG values, so even plain Long arguments
' need to be boxed as Variants before we hand their addresses across the cdecl boundary.
Private Function Pcap_DispInvoke(ByVal procAddr As Long, ByVal vtReturn As Integer, ByVal argCount As Long, ByRef argValues() As Long, ByRef retValue As Variant) As Long
    Dim hr As Long
    Dim argTypes() As Integer
    Dim argVars() As Variant
    Dim argPtrs() As Long
    Dim i As Long

    If procAddr = 0& Then
        Pcap_DispInvoke = -1&
        Exit Function
    End If

    If argCount > 0& Then
        ReDim argTypes(0& To argCount - 1&) As Integer
        ReDim argVars(0& To argCount - 1&) As Variant
        ReDim argPtrs(0& To argCount - 1&) As Long

        For i = 0& To argCount - 1&
            argVars(i) = argValues(i)
            argTypes(i) = VarType(argVars(i))
            argPtrs(i) = VarPtr(argVars(i))
        Next i

        hr = DispCallFunc(0&, procAddr, CC_CDECL, vtReturn, argCount, VarPtr(argTypes(0&)), VarPtr(argPtrs(0&)), VarPtr(retValue))
    Else
        hr = DispCallFunc(0&, procAddr, CC_CDECL, vtReturn, 0&, 0&, 0&, VarPtr(retValue))
    End If

    Pcap_DispInvoke = hr
End Function

Private Function Pcap_CallLong(ByVal procAddr As Long, ParamArray args() As Variant) As Long
    Dim retVar As Variant
    Dim argCount As Long
    Dim argValues() As Long
    Dim i As Long
    Dim hr As Long

    argCount = 0&
    On Error Resume Next
    argCount = UBound(args) - LBound(args) + 1&
    If Err.Number <> 0& Then
        Err.Clear
        argCount = 0&
    End If
    On Error GoTo 0

    If argCount > 0& Then
        ReDim argValues(0& To argCount - 1&) As Long
        For i = 0& To argCount - 1&
            argValues(i) = CLng(args(i))
        Next i
    End If

    retVar = Empty
    hr = Pcap_DispInvoke(procAddr, VT_I4, argCount, argValues, retVar)
    If hr <> 0& Then
        Pcap_CallLong = 0&
    ElseIf IsEmpty(retVar) Then
        Pcap_CallLong = 0&
    Else
        Pcap_CallLong = CLng(retVar)
    End If
End Function

Private Sub Pcap_CallVoid(ByVal procAddr As Long, ParamArray args() As Variant)
    Dim retVar As Variant
    Dim argCount As Long
    Dim argValues() As Long
    Dim i As Long

    argCount = 0&
    On Error Resume Next
    argCount = UBound(args) - LBound(args) + 1&
    If Err.Number <> 0& Then
        Err.Clear
        argCount = 0&
    End If
    On Error GoTo 0

    If argCount > 0& Then
        ReDim argValues(0& To argCount - 1&) As Long
        For i = 0& To argCount - 1&
            argValues(i) = CLng(args(i))
        Next i
    End If

    retVar = Empty
    Call Pcap_DispInvoke(procAddr, VT_EMPTY, argCount, argValues, retVar)
End Sub

Private Function Pcap_LoadApi() As Long
    If pcap_hModule <> 0& Then
        Pcap_LoadApi = 0&
        Exit Function
    End If

    pcap_hModule = LoadLibrary("wpcap.dll")
    If pcap_hModule = 0& Then
        pcap_logUnavailable "[PCAP] Unable to load wpcap.dll. Install Npcap or WinPcap, or use -net user."
        Pcap_LoadApi = -1&
        Exit Function
    End If

    pcap_proc_findalldevs = GetProcAddress(pcap_hModule, "pcap_findalldevs")
    pcap_proc_freealldevs = GetProcAddress(pcap_hModule, "pcap_freealldevs")
    pcap_proc_open_live = GetProcAddress(pcap_hModule, "pcap_open_live")
    pcap_proc_setnonblock = GetProcAddress(pcap_hModule, "pcap_setnonblock")
    pcap_proc_setmintocopy = GetProcAddress(pcap_hModule, "pcap_setmintocopy")
    pcap_proc_next_ex = GetProcAddress(pcap_hModule, "pcap_next_ex")
    pcap_proc_sendpacket = GetProcAddress(pcap_hModule, "pcap_sendpacket")
    pcap_proc_close = GetProcAddress(pcap_hModule, "pcap_close")
    pcap_proc_geterr = GetProcAddress(pcap_hModule, "pcap_geterr")

    If (pcap_proc_findalldevs = 0&) Or (pcap_proc_freealldevs = 0&) Or (pcap_proc_open_live = 0&) Or _
       (pcap_proc_setnonblock = 0&) Or (pcap_proc_next_ex = 0&) Or (pcap_proc_sendpacket = 0&) Or _
       (pcap_proc_close = 0&) Or (pcap_proc_geterr = 0&) Then
        debug_log DEBUG_ERROR, "[PCAP] wpcap.dll is missing one or more required exports."
        Call FreeLibrary(pcap_hModule)
        pcap_hModule = 0&
        pcap_proc_findalldevs = 0&
        pcap_proc_freealldevs = 0&
        pcap_proc_open_live = 0&
        pcap_proc_setnonblock = 0&
        pcap_proc_setmintocopy = 0&
        pcap_proc_next_ex = 0&
        pcap_proc_sendpacket = 0&
        pcap_proc_close = 0&
        pcap_proc_geterr = 0&
        Pcap_LoadApi = -1&
        Exit Function
    End If

    pcap_warned = 0&
    Pcap_LoadApi = 0&
End Function

Private Sub Pcap_UnloadApi()
    If pcap_hModule <> 0& Then
        Call FreeLibrary(pcap_hModule)
    End If

    pcap_hModule = 0&
    pcap_proc_findalldevs = 0&
    pcap_proc_freealldevs = 0&
    pcap_proc_open_live = 0&
    pcap_proc_setnonblock = 0&
    pcap_proc_setmintocopy = 0&
    pcap_proc_next_ex = 0&
    pcap_proc_sendpacket = 0&
    pcap_proc_close = 0&
    pcap_proc_geterr = 0&
End Sub

Private Function Pcap_FindAllDevs(ByRef allDevsPtr As Long, ByRef errbuf() As Byte) As Long
    Pcap_FindAllDevs = Pcap_CallLong(pcap_proc_findalldevs, VarPtr(allDevsPtr), VarPtr(errbuf(0&)))
End Function

Private Function Pcap_OpenLive(ByVal namePtr As Long, ByRef errbuf() As Byte) As Long
    Pcap_OpenLive = Pcap_CallLong(pcap_proc_open_live, namePtr, PCAP_SNAPLEN, PCAP_OPEN_PROMISC, PCAP_OPEN_TIMEOUT_MS, VarPtr(errbuf(0&)))
End Function

Private Function Pcap_SetNonBlock(ByVal handle As Long, ByRef errbuf() As Byte) As Long
    Pcap_SetNonBlock = Pcap_CallLong(pcap_proc_setnonblock, handle, 1&, VarPtr(errbuf(0&)))
End Function

Private Function Pcap_SetMinToCopy(ByVal handle As Long, ByVal minBytes As Long) As Long
    If pcap_proc_setmintocopy = 0& Then
        Pcap_SetMinToCopy = 0&
    Else
        Pcap_SetMinToCopy = Pcap_CallLong(pcap_proc_setmintocopy, handle, minBytes)
    End If
End Function

Private Function Pcap_NextEx(ByVal handle As Long, ByRef headerPtr As Long, ByRef dataPtr As Long) As Long
    Pcap_NextEx = Pcap_CallLong(pcap_proc_next_ex, handle, VarPtr(headerPtr), VarPtr(dataPtr))
End Function

Private Function Pcap_SendPacket(ByVal handle As Long, ByVal dataPtr As Long, ByVal frameLen As Long) As Long
    Pcap_SendPacket = Pcap_CallLong(pcap_proc_sendpacket, handle, dataPtr, frameLen)
End Function

Private Sub Pcap_Close(ByVal handle As Long)
    Call Pcap_CallVoid(pcap_proc_close, handle)
End Sub

Private Sub Pcap_FreeAllDevs(ByVal allDevsPtr As Long)
    Call Pcap_CallVoid(pcap_proc_freealldevs, allDevsPtr)
End Sub

Private Function Pcap_GetLastError(ByVal handle As Long) As String
    Dim errPtr As Long

    If handle = 0& Then Exit Function

    errPtr = Pcap_CallLong(pcap_proc_geterr, handle)
    Pcap_GetLastError = Pcap_CStringFromPtr(errPtr)
End Function

Private Function Pcap_StartPollTimer(ByVal errorText As String) As Long
    pcap_timer = timing_addTimer(TIMER_CB_PCAP_POLL, 0&, PCAP_POLL_HZ, TIMING_ENABLED)
    If pcap_timer = TIMING_ERROR Then
        debug_log DEBUG_ERROR, errorText
        Pcap_StartPollTimer = -1&
        Exit Function
    End If

    pcap_timerActive = 1&
    Pcap_StartPollTimer = 0&
End Function

Public Sub pcap_listdevs()
    Dim errbuf(0& To PCAP_ERRBUF_SIZE - 1&) As Byte
    Dim allDevsPtr As Long
    Dim devPtr As Long
    Dim dev As PCAP_IF_t
    Dim idx As Long
    Dim devName As String
    Dim devDesc As String

    debug_log DEBUG_INFO, "[NET] Available backends:" & vbCrLf
    debug_log DEBUG_INFO, "[NET]   user    Built-in usermode gateway backend (QEMU-style baseline)." & vbCrLf

    If Pcap_LoadApi() <> 0& Then Exit Sub

    If Pcap_FindAllDevs(allDevsPtr, errbuf) <> 0& Then
        debug_log DEBUG_ERROR, "[PCAP] Error in pcap_findalldevs: " & Pcap_CStringFromBytes(errbuf)
        Pcap_UnloadApi
        Exit Sub
    End If

    devPtr = allDevsPtr
    idx = 0&

    Do While devPtr <> 0&
        CopyMemory dev, devPtr, LenB(dev)
        idx = idx + 1&

        devName = Pcap_CStringFromPtr(dev.namePtr)
        devDesc = Pcap_CStringFromPtr(dev.descriptionPtr)
        If LenB(devDesc) = 0& Then devDesc = "No description available"
        If LenB(devName) <> 0& Then
            debug_log DEBUG_INFO, "[NET]   " & CStr(idx) & "    " & devDesc & " [" & devName & "]" & vbCrLf
        Else
            debug_log DEBUG_INFO, "[NET]   " & CStr(idx) & "    " & devDesc & vbCrLf
        End If

        devPtr = dev.nextPtr
    Loop

    If idx = 0& Then
        debug_log DEBUG_INFO, "[NET]   No host capture adapters found." & vbCrLf
    End If

    If allDevsPtr <> 0& Then
        Pcap_FreeAllDevs allDevsPtr
    End If
    Pcap_UnloadApi
End Sub

Public Function pcap_init(ByVal ifIndex As Long) As Long
    Dim errbuf(0& To PCAP_ERRBUF_SIZE - 1&) As Byte
    Dim allDevsPtr As Long
    Dim devPtr As Long
    Dim dev As PCAP_IF_t
    Dim idx As Long
    Dim devName As String
    Dim devDesc As String
    Dim nameBytes() As Byte

    pcap_shutdown

    If ifIndex = PCAP_IF_USERNET Then
        If usernet_init(ne2000_getPrimary()) <> 0& Then
            pcap_init = -1&
            Exit Function
        End If

        pcap_backend = PCAP_BACKEND_USERNET
        If Pcap_StartPollTimer("[NET] Failed to create usernet poll timer") <> 0& Then
            usernet_shutdown
            pcap_backend = PCAP_BACKEND_NONE
            pcap_init = -1&
            Exit Function
        End If

        pcap_init = 0&
        Exit Function
    End If

    If ifIndex <= 0& Then
        debug_log DEBUG_ERROR, "[PCAP] Host capture adapter index must be 1 or greater, or use -net user."
        pcap_init = -1&
        Exit Function
    End If

    If Pcap_LoadApi() <> 0& Then
        pcap_init = -1&
        Exit Function
    End If

    If Pcap_FindAllDevs(allDevsPtr, errbuf) <> 0& Then
        debug_log DEBUG_ERROR, "[PCAP] Error in pcap_findalldevs: " & Pcap_CStringFromBytes(errbuf)
        Pcap_UnloadApi
        pcap_init = -1&
        Exit Function
    End If

    devPtr = allDevsPtr
    idx = 1&
    Do While (devPtr <> 0&) And (idx < ifIndex)
        CopyMemory dev, devPtr, LenB(dev)
        devPtr = dev.nextPtr
        idx = idx + 1&
    Loop

    If devPtr = 0& Then
        debug_log DEBUG_ERROR, "[PCAP] Host capture adapter index " & CStr(ifIndex) & " was not found."
        If allDevsPtr <> 0& Then
            Pcap_FreeAllDevs allDevsPtr
        End If
        Pcap_UnloadApi
        pcap_init = -1&
        Exit Function
    End If

    CopyMemory dev, devPtr, LenB(dev)
    devName = Pcap_CStringFromPtr(dev.namePtr)
    devDesc = Pcap_CStringFromPtr(dev.descriptionPtr)
    If LenB(devDesc) = 0& Then devDesc = "No description available"

    debug_log DEBUG_INFO, "[PCAP] Initializing host capture backend using device: """ & devDesc & """"

    nameBytes = Pcap_StringToAnsiZ(devName)
    pcap_handle = Pcap_OpenLive(VarPtr(nameBytes(0&)), errbuf)

    If allDevsPtr <> 0& Then
        Pcap_FreeAllDevs allDevsPtr
    End If

    If pcap_handle = 0& Then
        debug_log DEBUG_ERROR, "[PCAP] Unable to open the adapter: " & Pcap_CStringFromBytes(errbuf)
        Pcap_UnloadApi
        pcap_init = -1&
        Exit Function
    End If

    If Pcap_SetNonBlock(pcap_handle, errbuf) <> 0& Then
        debug_log DEBUG_ERROR, "[PCAP] Failed to set non-blocking mode: " & Pcap_CStringFromBytes(errbuf)
        Pcap_Close pcap_handle
        pcap_handle = 0&
        Pcap_UnloadApi
        pcap_init = -1&
        Exit Function
    End If

    Call Pcap_SetMinToCopy(pcap_handle, 0&)

    pcap_devId = ne2000_getPrimary()
    If pcap_devId < 0& Then
        debug_log DEBUG_ERROR, "[PCAP] No NE2000 device is available for host capture."
        Pcap_Close pcap_handle
        pcap_handle = 0&
        Pcap_UnloadApi
        pcap_init = -1&
        Exit Function
    End If
    Pcap_ResetQueue

    pcap_backend = PCAP_BACKEND_WPCAP
    If Pcap_StartPollTimer("[PCAP] Failed to create host capture poll timer") <> 0& Then
        Pcap_Close pcap_handle
        pcap_handle = 0&
        pcap_backend = PCAP_BACKEND_NONE
        Pcap_UnloadApi
        pcap_init = -1&
        Exit Function
    End If

    pcap_init = 0&
End Function

Public Sub pcap_shutdown()
    If pcap_timerActive <> 0& Then
        timing_timerDisable pcap_timer
        pcap_timerActive = 0&
        pcap_timer = 0&
    End If

    Select Case pcap_backend
        Case PCAP_BACKEND_USERNET
            usernet_shutdown

        Case PCAP_BACKEND_WPCAP
            If pcap_handle <> 0& Then
                Pcap_Close pcap_handle
            End If
            Pcap_UnloadApi
    End Select

    Pcap_ResetState
End Sub

Public Sub pcap_check_packets(ByVal dummy As Long)
    Dim headerPtr As Long
    Dim dataPtr As Long
    Dim ret As Long
    Dim hdr As PCAP_PKTHDR_t
    Dim frame() As Byte
    Dim frameLen As Long
    Dim rxResult As Long
    Dim errText As String
    Dim budget As Long

    Select Case pcap_backend
        Case PCAP_BACKEND_USERNET
            usernet_pollPackets

        Case PCAP_BACKEND_WPCAP
            If (pcap_handle = 0&) Or (pcap_devId < 0&) Then Exit Sub

            budget = PCAP_POLL_BUDGET
            Call Pcap_DrainRxQueue

            Do While budget > 0&
                If pcap_rxCount >= PCAP_RX_QUEUE Then Exit Do

                headerPtr = 0&
                dataPtr = 0&
                ret = Pcap_NextEx(pcap_handle, headerPtr, dataPtr)

                If ret = 0& Then Exit Do

                If ret < 0& Then
                    errText = Pcap_GetLastError(pcap_handle)
                    If LenB(errText) <> 0& Then
                        debug_log DEBUG_ERROR, "[PCAP] Receive error: " & errText
                    End If
                    Exit Do
                End If

                If (headerPtr = 0&) Or (dataPtr = 0&) Then Exit Do

                CopyMemory hdr, headerPtr, LenB(hdr)
                frameLen = hdr.caplen
                If frameLen > 0& Then
                    ReDim frame(0& To frameLen - 1&) As Byte
                    CopyMemory frame(0&), dataPtr, frameLen

                    rxResult = ne2000_rx_frame_try(pcap_devId, frame, frameLen)
                    If rxResult = NE2000_RX_RETRY Then
                        If Pcap_QueueFrame(frame, frameLen) <> 0& Then Exit Do
                    End If
                End If
                budget = budget - 1&
            Loop

            Call Pcap_DrainRxQueue
    End Select
End Sub

Public Sub pcap_txPacket(ByRef data() As Byte, ByVal length As Long)
    Dim sendRet As Long
    Dim errText As String

    Select Case pcap_backend
        Case PCAP_BACKEND_USERNET
            usernet_txPacket data, length

        Case PCAP_BACKEND_WPCAP
            If pcap_handle = 0& Then Exit Sub
            If length <= 0& Then Exit Sub

            sendRet = Pcap_SendPacket(pcap_handle, VarPtr(data(LBound(data))), length)
            If sendRet <> 0& Then
                errText = Pcap_GetLastError(pcap_handle)
                If LenB(errText) = 0& Then errText = "pcap_sendpacket failed"
                debug_log DEBUG_ERROR, "[PCAP] Transmit error: " & errText
            End If
    End Select
End Sub
