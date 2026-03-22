Attribute VB_Name = "modTCPModem"
Option Explicit

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByVal addr As Long, ByVal addrlen As Long) As Long
Private Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)

Private Const TCPMODEM_MAX As Long = 4&
Private Const TCPMODEM_RXBUF_LEN As Long = 1024&
Private Const TCPMODEM_TXBUF_LEN As Long = 1025&

Private Const AF_INET As Long = 2&
Private Const SOCK_STREAM As Long = 1&
Private Const IPPROTO_TCP As Long = 6&
Private Const INADDR_ANY As Long = 0&
Private Const INVALID_SOCKET As Long = -1&
Private Const SOCKET_ERROR As Long = -1&
Private Const FIONBIO As Long = &H8004667E&

Private Const WSAEWOULDBLOCK As Long = 10035&
Private Const WSAENETDOWN As Long = 10050&
Private Const WSAENETUNREACH As Long = 10051&
Private Const WSAENETRESET As Long = 10052&
Private Const WSAECONNABORTED As Long = 10053&
Private Const WSAECONNRESET As Long = 10054&
Private Const WSAENOTCONN As Long = 10057&
Private Const WSAEHOSTUNREACH As Long = 10065&

Private Type IN_ADDR
    s_addr As Long
End Type

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero(0& To 7&) As Byte
End Type

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0& To 256&) As Byte
    szSystemStatus(0& To 128&) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type TCPMODEM_t
    escaped As Byte
    livesocket As Byte
    listening As Byte
    ringing As Byte
    ringstate As Byte
    ringtimer As Long
    listenport As Integer
    echocmd As Byte
    rxbuf(0& To TCPMODEM_RXBUF_LEN - 1&) As Byte
    txbuf(0& To TCPMODEM_TXBUF_LEN - 1&) As Byte
    rxpos As Integer
    txpos As Integer
    lasttx(0& To 2&) As Byte
    wsa As WSADATA
    socket As Long
    serversocket As Long
    server As SOCKADDR_IN
    uartnum As Long
End Type

Private tcpmodem_devs(0& To TCPMODEM_MAX - 1&) As TCPMODEM_t
Private tcpmodem_used(0& To TCPMODEM_MAX - 1&) As Byte

Private Function TCPMODEM_IsValid(ByVal uartnum As Long) As Boolean
    TCPMODEM_IsValid = ((uartnum >= 0&) And (uartnum < TCPMODEM_MAX) And (tcpmodem_used(uartnum) <> 0&))
End Function

Private Sub tcpmodem_setrx(ByVal uartnum As Long, ByVal text As String)
    Dim n As Long
    Dim i As Long

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub

    n = Len(text)
    If n > (TCPMODEM_RXBUF_LEN - 1&) Then n = TCPMODEM_RXBUF_LEN - 1&

    For i = 0& To n - 1&
        tcpmodem_devs(uartnum).rxbuf(i) = CByte(Asc(Mid$(text, i + 1&, 1&)) And &HFF&)
    Next i

    tcpmodem_devs(uartnum).rxbuf(n) = 0&
End Sub

Private Sub tcpmodem_clearrx(ByVal uartnum As Long)
    Dim i As Long

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub

    For i = 0& To TCPMODEM_RXBUF_LEN - 1&
        tcpmodem_devs(uartnum).rxbuf(i) = 0&
    Next i
End Sub

Private Sub tcpmodem_msrirq(ByVal uartnum As Long)
    If (uart_getIen(uartnum) And UART_IRQ_MSR_ENABLE) <> 0& Then
        uart_orPendIrq uartnum, UART_PENDING_MSR
    End If
End Sub

Public Function tcpmodem_listen(ByVal uartnum As Long, ByVal port As Long) As Long
    Dim ret As Long
    Dim iMode As Long
    Dim addr As SOCKADDR_IN

    If Not TCPMODEM_IsValid(uartnum) Then
        tcpmodem_listen = -1&
        Exit Function
    End If

    iMode = 1&
    closesocket tcpmodem_devs(uartnum).serversocket

    tcpmodem_devs(uartnum).serversocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If tcpmodem_devs(uartnum).serversocket = INVALID_SOCKET Then
        tcpmodem_listen = -1&
        Exit Function
    End If

    ioctlsocket tcpmodem_devs(uartnum).serversocket, FIONBIO, iMode

    addr.sin_family = AF_INET
    addr.sin_port = htons(CInt(port And &HFFFF&))
    addr.sin_addr.s_addr = INADDR_ANY

    ret = bind(tcpmodem_devs(uartnum).serversocket, addr, LenB(addr))
    If ret = SOCKET_ERROR Then
        closesocket tcpmodem_devs(uartnum).socket
        tcpmodem_listen = -1&
        Exit Function
    End If

    ret = listen(tcpmodem_devs(uartnum).serversocket, 1&)
    If ret = SOCKET_ERROR Then
        closesocket tcpmodem_devs(uartnum).serversocket
        tcpmodem_listen = -1&
        Exit Function
    End If

    tcpmodem_devs(uartnum).listening = 1&
    tcpmodem_listen = 0&
End Function

Public Function tcpmodem_connect(ByVal uartnum As Long, ByVal host As String, ByVal port As Long) As Long
    Dim ret As Long
    Dim iMode As Long
    Dim hostentPtr As Long
    Dim hostentVal As HOSTENT
    Dim addrListPtr As Long
    Dim addrPtr As Long
    Dim msrVal As Byte

    If Not TCPMODEM_IsValid(uartnum) Then
        tcpmodem_connect = -1&
        Exit Function
    End If

    tcpmodem_devs(uartnum).rxpos = 0&

    tcpmodem_devs(uartnum).socket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If tcpmodem_devs(uartnum).socket = INVALID_SOCKET Then
        WSACleanup
        tcpmodem_connect = -1&
        Exit Function
    End If

    hostentPtr = gethostbyname(host)
    If hostentPtr = 0& Then
        tcpmodem_setrx uartnum, vbLf & "NO CARRIER" & vbCrLf
        tcpmodem_connect = -1&
        Exit Function
    End If

    CopyMemory hostentVal, ByVal hostentPtr, LenB(hostentVal)
    If hostentVal.h_addr_list = 0& Then
        tcpmodem_setrx uartnum, vbLf & "NO CARRIER" & vbCrLf
        tcpmodem_connect = -1&
        Exit Function
    End If

    CopyMemory addrPtr, ByVal hostentVal.h_addr_list, 4&
    If addrPtr = 0& Then
        tcpmodem_setrx uartnum, vbLf & "NO CARRIER" & vbCrLf
        tcpmodem_connect = -1&
        Exit Function
    End If

    tcpmodem_devs(uartnum).server.sin_family = AF_INET
    tcpmodem_devs(uartnum).server.sin_port = htons(CInt(port And &HFFFF&))
    CopyMemory tcpmodem_devs(uartnum).server.sin_addr.s_addr, ByVal addrPtr, 4&

    ret = connect(tcpmodem_devs(uartnum).socket, tcpmodem_devs(uartnum).server, LenB(tcpmodem_devs(uartnum).server))
    If ret = SOCKET_ERROR Then
        closesocket tcpmodem_devs(uartnum).socket
        tcpmodem_setrx uartnum, vbLf & "NO CARRIER" & vbCrLf
        tcpmodem_connect = -1&
        Exit Function
    End If

    iMode = 1&
    ioctlsocket tcpmodem_devs(uartnum).socket, FIONBIO, iMode

    tcpmodem_setrx uartnum, vbLf & "CONNECT" & vbCrLf
    tcpmodem_devs(uartnum).livesocket = 1&
    tcpmodem_devs(uartnum).escaped = 0&
    tcpmodem_devs(uartnum).listening = 0&
    closesocket tcpmodem_devs(uartnum).serversocket

    msrVal = uart_getMsr(uartnum)
    msrVal = (msrVal Or &H80&)
    msrVal = (msrVal And &HF7&)
    uart_setMsr uartnum, msrVal
    tcpmodem_msrirq uartnum

    tcpmodem_connect = 0&
End Function

Private Sub tcpmodem_setringmsr(ByVal uartnum As Long, ByVal state As Byte)
    Dim msrVal As Byte

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub

    msrVal = uart_getMsr(uartnum)

    If tcpmodem_devs(uartnum).ringstate <> 0& Then
        msrVal = (msrVal Or &H40&)
        msrVal = (msrVal And &HFB&)
    Else
        msrVal = (msrVal And &HBF&)
        msrVal = (msrVal Or &H4&)
    End If

    uart_setMsr uartnum, msrVal
    tcpmodem_msrirq uartnum
End Sub

Public Sub tcpmodem_offline(ByVal uartnum As Long)
    Dim msrVal As Byte

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub

    closesocket tcpmodem_devs(uartnum).socket
    tcpmodem_devs(uartnum).livesocket = 0&
    tcpmodem_devs(uartnum).escaped = 1&
    tcpmodem_setrx uartnum, vbLf & "NO CARRIER" & vbCrLf
    tcpmodem_devs(uartnum).rxpos = 0&
    tcpmodem_devs(uartnum).listening = 0&
    tcpmodem_devs(uartnum).ringing = 0&

    msrVal = uart_getMsr(uartnum)
    msrVal = (msrVal And &H7F&)
    msrVal = (msrVal Or &H8&)
    uart_setMsr uartnum, msrVal
    tcpmodem_msrirq uartnum

    tcpmodem_listen uartnum, (tcpmodem_devs(uartnum).listenport And &HFFFF&)
End Sub

Private Sub tcpmodem_parseAT(ByVal uartnum As Long)
    Dim i As Long
    Dim port As Long
    Dim hostpos As Long
    Dim isdial As Byte
    Dim host As String
    Dim cc As Byte

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub

    i = 0&
    port = 23&
    hostpos = 0&
    isdial = 0&
    host = vbNullString

    With tcpmodem_devs(uartnum)
        If (.txbuf(0&) <> Asc("A")) Or (.txbuf(1&) <> Asc("T")) Then Exit Sub

        If .txbuf(2&) = Asc("D") Then
            If .txbuf(3&) = Asc("T") Then
                i = 4&
            Else
                i = 3&
            End If
            isdial = 1&
        End If

        If .txbuf(2&) = Asc("H") Then
            If .livesocket <> 0& Then
                tcpmodem_offline uartnum
            Else
                tcpmodem_setrx uartnum, vbLf & "OK" & vbCrLf
                .rxpos = 0&
            End If
            Exit Sub
        End If

        If .txbuf(2&) = Asc("A") Then
            If .ringing <> 0& Then
                timing_timerDisable .ringtimer
                .ringing = 0&
                .escaped = 0&

                cc = uart_getMsr(uartnum)
                cc = (cc Or &H80&)
                cc = (cc And &HF7&)
                uart_setMsr uartnum, cc
                tcpmodem_msrirq uartnum

                tcpmodem_setrx uartnum, vbLf & "CONNECT" & vbCrLf
                .rxpos = 0&
            Else
                tcpmodem_setrx uartnum, vbLf & "OK" & vbCrLf
                .rxpos = 0&
            End If
            Exit Sub
        End If

        If .txbuf(2&) = Asc("E") Then
            .echocmd = CByte(((CLng(.txbuf(3&)) - Asc("0")) And 1&))
            tcpmodem_setrx uartnum, CStr(.echocmd) & vbCrLf
            .rxpos = 0&
        End If

        If isdial <> 0& Then
            For i = i To .txpos - 1&
                cc = .txbuf(i)
                If cc = Asc(":") Then
                    port = CLng(Val(tcpmodem_txTailString(uartnum, i + 1&))) And &HFFFF&
                    Exit For
                End If
                host = host & Chr$(cc)
                hostpos = hostpos + 1&
            Next i

            tcpmodem_connect uartnum, host, port
        Else
            tcpmodem_setrx uartnum, vbLf & "OK" & vbCrLf
            .rxpos = 0&
        End If
    End With
End Sub

Private Function tcpmodem_txTailString(ByVal uartnum As Long, ByVal startPos As Long) As String
    Dim i As Long
    Dim cc As Byte
    Dim s As String

    If Not TCPMODEM_IsValid(uartnum) Then
        tcpmodem_txTailString = vbNullString
        Exit Function
    End If

    s = vbNullString
    For i = startPos To TCPMODEM_TXBUF_LEN - 1&
        cc = tcpmodem_devs(uartnum).txbuf(i)
        If cc = 0& Then Exit For
        s = s & Chr$(cc)
    Next i

    tcpmodem_txTailString = s
End Function

Public Sub tcpmodem_rxpoll(ByVal uartnum As Long)
    Dim cc As Byte
    Dim ret As Long
    Dim wsaErr As Long

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub
    If uart_canAcceptRx(uartnum) = 0& Then Exit Sub

    With tcpmodem_devs(uartnum)
        If (.livesocket <> 0&) And ((uart_getMcr(uartnum) And 1&) = 0&) Then
            tcpmodem_offline uartnum
        End If

        If (.livesocket <> 0&) And (.escaped = 0&) And (.rxbuf(.rxpos) = 0&) Then
            ret = recv(.socket, cc, 1&, 0&)
            If ret > 0& Then
                uart_rxdata uartnum, cc
            ElseIf ret = SOCKET_ERROR Then
                wsaErr = WSAGetLastError()
                Select Case wsaErr
                    Case WSAENETDOWN, WSAENETUNREACH, WSAENETRESET, WSAECONNABORTED, WSAECONNRESET, WSAENOTCONN, WSAEHOSTUNREACH
                        tcpmodem_offline uartnum
                End Select
            End If
        Else
            cc = .rxbuf(.rxpos)
            If cc <> 0& Then
                uart_rxdata uartnum, cc
                .rxpos = CInt(.rxpos + 1&)
            Else
                tcpmodem_clearrx uartnum
                .rxpos = 0&
            End If

            If .listening <> 0& Then
                closesocket .socket
                .socket = accept(.serversocket, 0&, 0&)

                If (.socket = INVALID_SOCKET) And (WSAGetLastError() <> WSAEWOULDBLOCK) Then
                    tcpmodem_listen uartnum, (.listenport And &HFFFF&)
                ElseIf .socket <> INVALID_SOCKET Then
                    closesocket .serversocket
                    .livesocket = 1&
                    .escaped = 1&
                    .listening = 0&
                    .ringing = 1&
                    .ringstate = 0&
                    timing_timerEnable .ringtimer
                End If
            End If
        End If
    End With
End Sub

Public Sub tcpmodem_tx(ByVal uartnum As Long, ByVal value As Byte)
    Dim ret As Long
    Dim wsaErr As Long
    Dim txValue As Byte

    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub

    With tcpmodem_devs(uartnum)
        .lasttx(0&) = .lasttx(1&)
        .lasttx(1&) = .lasttx(2&)
        .lasttx(2&) = value

        If (.lasttx(0&) = Asc("+")) And (.lasttx(1&) = Asc("+")) And (.lasttx(2&) = Asc("+")) And (.livesocket <> 0&) Then
            .lasttx(0&) = 0&
            .lasttx(1&) = 0&
            .lasttx(2&) = 0&
            .escaped = CByte(.escaped Xor 1&)
            If .escaped <> 0& Then
                tcpmodem_setrx uartnum, vbLf & "OK" & vbCrLf
                .rxpos = 0&
            End If
        End If

        If (.livesocket <> 0&) And (.escaped = 0&) Then
            txValue = value
            ret = send(.socket, txValue, 1&, 0&)
            If ret = SOCKET_ERROR Then
                wsaErr = WSAGetLastError()
                Select Case wsaErr
                    Case WSAENETDOWN, WSAENETUNREACH, WSAENETRESET, WSAECONNABORTED, WSAECONNRESET, WSAENOTCONN, WSAEHOSTUNREACH
                        tcpmodem_offline uartnum
                End Select
            End If
        Else
            If .echocmd <> 0& Then
                uart_rxdata uartnum, value
            End If

            If value = 8& Then
                If .txpos > 0& Then
                    .txpos = CInt(.txpos - 1&)
                End If
            ElseIf value = 13& Then
                .txbuf(.txpos + 1&) = 0&
                tcpmodem_parseAT uartnum
                .txpos = 0&
            ElseIf .txpos < 1023& Then
                Select Case value
                    Case 0&, 32&, Asc("+")
                        ' no-op
                    Case Else
                        If (value >= Asc("a")) And (value <= Asc("z")) Then
                            value = CByte(value - (Asc("a") - Asc("A")))
                        End If
                        .txbuf(.txpos) = value
                        .txpos = CInt(.txpos + 1&)
                End Select
            End If
        End If
    End With
End Sub

Public Sub tcpmodem_ringer(ByVal uartnum As Long)
    If Not TCPMODEM_IsValid(uartnum) Then Exit Sub
    If tcpmodem_devs(uartnum).ringing = 0& Then Exit Sub

    tcpmodem_devs(uartnum).ringstate = CByte(tcpmodem_devs(uartnum).ringstate Xor 1&)

    If tcpmodem_devs(uartnum).ringstate <> 0& Then
        tcpmodem_setrx uartnum, "RING" & vbCrLf
        tcpmodem_devs(uartnum).rxpos = 0&
    End If

    tcpmodem_setringmsr uartnum, tcpmodem_devs(uartnum).ringstate
End Sub

Public Sub tcpmodem_init(ByVal uartnum As Long, ByVal listenPort As Long)
    Dim initDev As TCPMODEM_t

    If (uartnum < 0&) Or (uartnum >= TCPMODEM_MAX) Then Exit Sub

    debug_log DEBUG_INFO, "[TCPMODEM] Initializing TCP serial modem emulator (listen on port " & CStr(listenPort And &HFFFF&) & ")"

    tcpmodem_devs(uartnum) = initDev
    tcpmodem_used(uartnum) = 1&

    With tcpmodem_devs(uartnum)
        .uartnum = uartnum
        .escaped = 1&
        .echocmd = 1&
        .listenport = CInt(listenPort And &HFFFF&)

        WSAStartup &H202&, .wsa
        tcpmodem_listen uartnum, (.listenport And &HFFFF&)
        .ringtimer = timing_addTimer(TIMER_CB_TCPMODEM_RINGER, uartnum, 1#, TIMING_DISABLED)
    End With
End Sub

