Attribute VB_Name = "modUserNet"
Option Explicit

Private Const USERNET_RX_QUEUE As Long = 256&
Private Const USERNET_MAX_FRAME As Long = 2048&

Private Const USERNET_ETHERTYPE_IPV4 As Long = &H800&
Private Const USERNET_ETHERTYPE_ARP As Long = &H806&

Private Const USERNET_IPPROTO_ICMP As Long = 1&
Private Const USERNET_IPPROTO_UDP As Long = 17&

Private Const USERNET_DHCP_CLIENT_PORT As Long = 68&
Private Const USERNET_DHCP_SERVER_PORT As Long = 67&
Private Const USERNET_DHCPDISCOVER As Byte = 1&
Private Const USERNET_DHCPOFFER As Byte = 2&
Private Const USERNET_DHCPREQUEST As Byte = 3&
Private Const USERNET_DHCPACK As Byte = 5&
Private Const USERNET_UDP_MAX_PAYLOAD As Long = 1472&
Private Const USERNET_UDP_MAX_FLOWS As Long = 64&
Private Const USERNET_UDP_IDLE_SECONDS As Double = 120#
Private Const USERNET_ICMP_TIMEOUT_MS As Long = 250&
Private Const USERNET_IP_SUCCESS As Long = 0&

Private Const AF_INET As Long = 2&
Private Const SOCK_DGRAM As Long = 2&
Private Const IPPROTO_UDP_SOCKET As Long = 17&
Private Const INADDR_ANY As Long = 0&
Private Const INVALID_SOCKET As Long = -1&
Private Const SOCKET_ERROR As Long = -1&
Private Const FIONBIO As Long = &H8004667E&

Private Const WSAEWOULDBLOCK As Long = 10035&

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef toAddr As SOCKADDR_IN, ByVal tolen As Long) As Long
Private Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef fromAddr As SOCKADDR_IN, ByRef fromlen As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function IcmpCreateFile Lib "iphlpapi.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "iphlpapi.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "iphlpapi.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByRef RequestData As Any, ByVal RequestSize As Long, ByVal RequestOptions As Long, ByRef ReplyBuffer As Any, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)

Private Type IN_ADDR
    s_addr As Long
End Type

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero(0 To 7) As Byte
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type USERNET_IP_OPTION_INFORMATION_t
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Byte
    OptionsData As Long
End Type

Private Type USERNET_ICMP_ECHO_REPLY_t
    Address As Long
    status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    DataPtr As Long
    Options As USERNET_IP_OPTION_INFORMATION_t
End Type

Private Type USERNET_UDP_FLOW_t
    used As Byte
    dnsProxy As Byte
    hostSock As Long
    guestSrcIp As Long
    guestDstIp As Long
    hostDstIp As Long
    guestSrcPort As Long
    guestDstPort As Long
    hostDstPort As Long
    lastSeen As Double
End Type

Private usernet_active As Byte
Private usernet_devId As Long
Private usernet_guestMacKnown As Byte
Private usernet_guestMac(0 To 5) As Byte
Private usernet_gatewayMac(0 To 5) As Byte
Private usernet_gatewayIp(0 To 3) As Byte
Private usernet_offerIp(0 To 3) As Byte
Private usernet_dnsIp(0 To 3) As Byte
Private usernet_netmask(0 To 3) As Byte
Private usernet_broadcastIp(0 To 3) As Byte
Private usernet_ipId As Long

Private usernet_rxFrames(0 To USERNET_RX_QUEUE - 1) As Variant
Private usernet_rxLen(0 To USERNET_RX_QUEUE - 1) As Long
Private usernet_rxHead As Long
Private usernet_rxTail As Long
Private usernet_rxCount As Long
Private usernet_gatewayIpBE As Long
Private usernet_offerIpBE As Long
Private usernet_dnsIpBE As Long
Private usernet_dnsUpstreamIpBE As Long
Private usernet_wsa As WSADATA
Private usernet_wsaStarted As Byte
Private usernet_udpFlows(0 To USERNET_UDP_MAX_FLOWS - 1) As USERNET_UDP_FLOW_t
Private usernet_icmpHandle As Long
Private Function UserNet_Byte(ByRef data() As Byte, ByVal idx As Long) As Byte
    If idx < LBound(data) Or idx > UBound(data) Then
        UserNet_Byte = 0&
    Else
        UserNet_Byte = data(idx)
    End If
End Function

Private Function UserNet_ReadBE16(ByRef data() As Byte, ByVal idx As Long) As Long
    UserNet_ReadBE16 = ((CLng(UserNet_Byte(data, idx)) * &H100&) Or CLng(UserNet_Byte(data, idx + 1&))) And &HFFFF&
End Function

Private Function UserNet_ReadBE32(ByRef data() As Byte, ByVal idx As Long) As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long

    b0 = CLng(UserNet_Byte(data, idx))
    b1 = CLng(UserNet_Byte(data, idx + 1&))
    b2 = CLng(UserNet_Byte(data, idx + 2&))
    b3 = CLng(UserNet_Byte(data, idx + 3&))

    UserNet_ReadBE32 = U32FromDouble(CDbl(b0) * 16777216# + CDbl(b1) * 65536# + CDbl(b2) * 256# + CDbl(b3))
End Function

Private Sub UserNet_WriteBE16(ByRef data() As Byte, ByVal idx As Long, ByVal value As Long)
    data(idx) = CByte(U32Shr(value, 8&) And &HFF&)
    data(idx + 1&) = CByte(value And &HFF&)
End Sub

Private Sub UserNet_WriteBE32(ByRef data() As Byte, ByVal idx As Long, ByVal value As Long)
    data(idx) = CByte(U32Shr(value, 24&) And &HFF&)
    data(idx + 1&) = CByte(U32Shr(value, 16&) And &HFF&)
    data(idx + 2&) = CByte(U32Shr(value, 8&) And &HFF&)
    data(idx + 3&) = CByte(value And &HFF&)
End Sub

Private Function UserNet_MakeIpBE(ByVal b0 As Byte, ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte) As Long
    UserNet_MakeIpBE = U32FromDouble(CDbl(b0) * 16777216# + CDbl(b1) * 65536# + CDbl(b2) * 256# + CDbl(b3))
End Function

Private Function UserNet_BSwap32(ByVal value As Long) As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long

    b0 = U32Shr(value, 24&) And &HFF&
    b1 = U32Shr(value, 16&) And &HFF&
    b2 = U32Shr(value, 8&) And &HFF&
    b3 = value And &HFF&

    UserNet_BSwap32 = U32FromDouble(CDbl(b3) * 16777216# + CDbl(b2) * 65536# + CDbl(b1) * 256# + CDbl(b0))
End Function

Private Function UserNet_IsServiceIp(ByVal ipBe As Long) As Boolean
    UserNet_IsServiceIp = (ipBe = usernet_gatewayIpBE) Or (ipBe = usernet_dnsIpBE)
End Function

Private Function UserNet_FlowTimeoutTicks() As Double
    UserNet_FlowTimeoutTicks = timing_getFreq() * USERNET_UDP_IDLE_SECONDS
End Function

Private Function UserNet_DnsResolveLocal(ByRef query() As Byte, ByVal queryLen As Long, ByRef reply() As Byte, ByRef replyLen As Long) As Long
    Dim pos As Long
    Dim qdCount As Long
    Dim labelLen As Long
    Dim qName As String
    Dim i As Long
    Dim qType As Long
    Dim qClass As Long
    Dim qEnd As Long
    Dim headerFlags2 As Byte
    Dim hostentPtr As Long
    Dim hostentVal As HOSTENT
    Dim addrPtr As Long
    Dim ipRaw As Long
    Dim ansPos As Long
    Dim foundIp As Byte

    UserNet_DnsResolveLocal = 0&
    replyLen = 0&

    If queryLen < 17& Then Exit Function

    qdCount = UserNet_ReadBE16(query, 4&)
    If qdCount < 1& Then Exit Function

    pos = 12&
    qName = vbNullString

    Do While pos < queryLen
        labelLen = CLng(query(pos))

        If labelLen = 0& Then
            pos = pos + 1&
            Exit Do
        End If

        If (labelLen And &HC0&) <> 0& Then Exit Function
        If labelLen > 63& Then Exit Function
        If (pos + 1& + labelLen) > queryLen Then Exit Function

        If Len(qName) > 0& Then
            qName = qName & "."
        End If

        For i = 0& To labelLen - 1&
            qName = qName & Chr$(query(pos + 1& + i))
        Next i

        pos = pos + 1& + labelLen
    Loop

    If (pos + 4&) > queryLen Then Exit Function

    qType = UserNet_ReadBE16(query, pos)
    qClass = UserNet_ReadBE16(query, pos + 2&)
    qEnd = pos + 4&

    If Len(qName) = 0& Then Exit Function
    If (qType <> 1&) Or (qClass <> 1&) Then Exit Function

    headerFlags2 = CByte(&H80& Or (query(2&) And &H1&))

    hostentPtr = gethostbyname(qName)
    foundIp = 0&
    If hostentPtr <> 0& Then
        CopyMemory hostentVal, ByVal hostentPtr, LenB(hostentVal)
        If hostentVal.h_addr_list <> 0& Then
            CopyMemory addrPtr, ByVal hostentVal.h_addr_list, 4&
            If addrPtr <> 0& Then
                CopyMemory ipRaw, ByVal addrPtr, 4&
                foundIp = 1&
            End If
        End If
    End If

    If foundIp <> 0& Then
        ReDim reply(0 To (12& + (qEnd - 12&) + 16&) - 1&) As Byte
    Else
        ReDim reply(0 To (12& + (qEnd - 12&)) - 1&) As Byte
    End If

    reply(0&) = query(0&)
    reply(1&) = query(1&)
    reply(2&) = headerFlags2
    If foundIp <> 0& Then
        reply(3&) = &H80&
    Else
        reply(3&) = &H83&
    End If

    reply(4&) = 0&
    reply(5&) = 1&

    If foundIp <> 0& Then
        reply(6&) = 0&
        reply(7&) = 1&
    Else
        reply(6&) = 0&
        reply(7&) = 0&
    End If

    reply(8&) = 0&
    reply(9&) = 0&
    reply(10&) = 0&
    reply(11&) = 0&

    For i = 12& To qEnd - 1&
        reply(i) = query(i)
    Next i

    If foundIp <> 0& Then
        ansPos = qEnd
        reply(ansPos + 0&) = &HC0&
        reply(ansPos + 1&) = &HC&
        reply(ansPos + 2&) = 0&
        reply(ansPos + 3&) = 1&
        reply(ansPos + 4&) = 0&
        reply(ansPos + 5&) = 1&
        reply(ansPos + 6&) = 0&
        reply(ansPos + 7&) = 0&
        reply(ansPos + 8&) = 0&
        reply(ansPos + 9&) = 60&
        reply(ansPos + 10&) = 0&
        reply(ansPos + 11&) = 4&
        reply(ansPos + 12&) = CByte(ipRaw And &HFF&)
        reply(ansPos + 13&) = CByte(U32Shr(ipRaw, 8&) And &HFF&)
        reply(ansPos + 14&) = CByte(U32Shr(ipRaw, 16&) And &HFF&)
        reply(ansPos + 15&) = CByte(U32Shr(ipRaw, 24&) And &HFF&)
        replyLen = ansPos + 16&
    Else
        replyLen = qEnd
    End If

    UserNet_DnsResolveLocal = 1&
End Function

Private Function UserNet_Checksum(ByRef data() As Byte, ByVal startIdx As Long, ByVal byteCount As Long) As Long
    Dim i As Long
    Dim sum As Long
    Dim wordVal As Long

    sum = 0&
    i = 0&

    Do While i < byteCount
        If (i + 1&) < byteCount Then
            wordVal = (CLng(data(startIdx + i)) * &H100&) Or CLng(data(startIdx + i + 1&))
        Else
            wordVal = (CLng(data(startIdx + i)) * &H100&)
        End If

        sum = sum + wordVal
        Do While U32Shr(sum, 16&) <> 0&
            sum = (sum And &HFFFF&) + (U32Shr(sum, 16&) And &HFFFF&)
        Loop

        i = i + 2&
    Loop

    UserNet_Checksum = ((Not sum) And &HFFFF&)
End Function

Private Function UserNet_IpMatches(ByRef data() As Byte, ByVal idx As Long, ByVal ip0 As Byte, ByVal ip1 As Byte, ByVal ip2 As Byte, ByVal ip3 As Byte) As Boolean
    UserNet_IpMatches = (UserNet_Byte(data, idx) = ip0) And _
                       (UserNet_Byte(data, idx + 1&) = ip1) And _
                       (UserNet_Byte(data, idx + 2&) = ip2) And _
                       (UserNet_Byte(data, idx + 3&) = ip3)
End Function

Private Sub UserNet_CopyMac(ByRef dst() As Byte, ByVal dstIdx As Long, ByRef src() As Byte, ByVal srcIdx As Long)
    Dim i As Long

    For i = 0& To 5&
        dst(dstIdx + i) = UserNet_Byte(src, srcIdx + i)
    Next i
End Sub

Private Sub UserNet_CopyGuestMac(ByRef src() As Byte, ByVal srcIdx As Long)
    Dim i As Long

    For i = 0& To 5&
        usernet_guestMac(i) = UserNet_Byte(src, srcIdx + i)
    Next i

    usernet_guestMacKnown = 1&
End Sub

Private Sub UserNet_ResetQueue()
    Dim i As Long

    usernet_rxHead = 0&
    usernet_rxTail = 0&
    usernet_rxCount = 0&

    For i = 0& To USERNET_RX_QUEUE - 1&
        usernet_rxFrames(i) = Empty
        usernet_rxLen(i) = 0&
    Next i
End Sub

Private Sub UserNet_QueueFrame(ByRef frame() As Byte, ByVal frameLen As Long)
    Dim i As Long
    Dim slot As Long
    Dim copyBuf() As Byte

    If frameLen <= 0& Then Exit Sub
    If frameLen > USERNET_MAX_FRAME Then Exit Sub

    If usernet_rxCount >= USERNET_RX_QUEUE Then
        ' Preserve oldest queued frames for retry when the emulated NIC is backpressured.
        Exit Sub
    End If

    ReDim copyBuf(0 To frameLen - 1) As Byte
    For i = 0& To frameLen - 1&
        copyBuf(i) = UserNet_Byte(frame, i)
    Next i

    slot = usernet_rxTail
    usernet_rxFrames(slot) = copyBuf
    usernet_rxLen(slot) = frameLen
    usernet_rxTail = (usernet_rxTail + 1&) Mod USERNET_RX_QUEUE
    usernet_rxCount = usernet_rxCount + 1&
End Sub

Private Sub UserNet_PopFrame()
    If usernet_rxCount <= 0& Then Exit Sub

    usernet_rxFrames(usernet_rxHead) = Empty
    usernet_rxLen(usernet_rxHead) = 0&
    usernet_rxHead = (usernet_rxHead + 1&) Mod USERNET_RX_QUEUE
    usernet_rxCount = usernet_rxCount - 1&
End Sub

Private Function UserNet_DrainRxQueue() As Long
    Dim frame As Variant
    Dim pkt() As Byte
    Dim frameLen As Long
    Dim rxResult As Long

    UserNet_DrainRxQueue = 0&

    If (usernet_active = 0&) Or (usernet_devId < 0&) Then Exit Function

    Do While usernet_rxCount > 0&
        frame = usernet_rxFrames(usernet_rxHead)
        frameLen = usernet_rxLen(usernet_rxHead)

        If (frameLen <= 0&) Or (IsArray(frame) = False) Then
            UserNet_PopFrame
            UserNet_DrainRxQueue = UserNet_DrainRxQueue + 1&
        Else
            pkt = frame
            rxResult = ne2000_rx_frame_try(usernet_devId, pkt, frameLen)

            If rxResult = NE2000_RX_ACCEPTED Then
                UserNet_PopFrame
                UserNet_DrainRxQueue = UserNet_DrainRxQueue + 1&
            ElseIf rxResult = NE2000_RX_DROP Then
                UserNet_PopFrame
                UserNet_DrainRxQueue = UserNet_DrainRxQueue + 1&
            Else
                Exit Function
            End If
        End If
    Loop
End Function

Private Sub UserNet_DeliverFrame(ByRef frame() As Byte, ByVal frameLen As Long)
    Dim rxResult As Long

    If frameLen <= 0& Then Exit Sub

    If (usernet_active <> 0&) And (usernet_devId >= 0&) Then
        rxResult = ne2000_rx_frame_try(usernet_devId, frame, frameLen)
        If rxResult = NE2000_RX_ACCEPTED Then Exit Sub
        If rxResult = NE2000_RX_DROP Then Exit Sub
    End If

    UserNet_QueueFrame frame, frameLen
End Sub

Private Sub UserNet_SendArpReply(ByRef senderMac() As Byte, ByRef senderIp() As Byte, ByVal targetIp0 As Byte, ByVal targetIp1 As Byte, ByVal targetIp2 As Byte, ByVal targetIp3 As Byte)
    Dim frame(0 To 41) As Byte
    Dim i As Long

    For i = 0& To 5&
        frame(i) = senderMac(i)
        frame(6& + i) = usernet_gatewayMac(i)
    Next i

    frame(12&) = &H8&
    frame(13&) = &H6&

    frame(14&) = 0&
    frame(15&) = 1&
    frame(16&) = &H8&
    frame(17&) = 0&
    frame(18&) = 6&
    frame(19&) = 4&
    frame(20&) = 0&
    frame(21&) = 2&

    For i = 0& To 5&
        frame(22& + i) = usernet_gatewayMac(i)
    Next i

    frame(28&) = targetIp0
    frame(29&) = targetIp1
    frame(30&) = targetIp2
    frame(31&) = targetIp3

    For i = 0& To 5&
        frame(32& + i) = senderMac(i)
    Next i

    frame(38&) = senderIp(0&)
    frame(39&) = senderIp(1&)
    frame(40&) = senderIp(2&)
    frame(41&) = senderIp(3&)
    UserNet_DeliverFrame frame, 42&
End Sub

Private Sub UserNet_DhcpOptByte(ByRef data() As Byte, ByRef pos As Long, ByVal code As Byte, ByVal value As Byte)
    data(pos) = code
    data(pos + 1&) = 1&
    data(pos + 2&) = value
    pos = pos + 3&
End Sub

Private Sub UserNet_DhcpOpt4(ByRef data() As Byte, ByRef pos As Long, ByVal code As Byte, ByVal b0 As Byte, ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte)
    data(pos) = code
    data(pos + 1&) = 4&
    data(pos + 2&) = b0
    data(pos + 3&) = b1
    data(pos + 4&) = b2
    data(pos + 5&) = b3
    pos = pos + 6&
End Sub

Private Sub UserNet_DhcpOpt32(ByRef data() As Byte, ByRef pos As Long, ByVal code As Byte, ByVal value As Long)
    data(pos) = code
    data(pos + 1&) = 4&
    UserNet_WriteBE32 data, pos + 2&, value
    pos = pos + 6&
End Sub

Private Function UserNet_DhcpMessageType(ByRef data() As Byte, ByVal optStart As Long, ByVal optEnd As Long) As Byte
    Dim pos As Long
    Dim optCode As Long
    Dim optLen As Long

    pos = optStart
    UserNet_DhcpMessageType = 0&

    Do While pos < optEnd
        optCode = CLng(UserNet_Byte(data, pos))
        If optCode = 255& Then Exit Do

        If optCode = 0& Then
            pos = pos + 1&
        Else
            If (pos + 1&) >= optEnd Then Exit Do
            optLen = CLng(UserNet_Byte(data, pos + 1&))
            If (pos + 2& + optLen) > optEnd Then Exit Do

            If (optCode = 53&) And (optLen >= 1&) Then
                UserNet_DhcpMessageType = UserNet_Byte(data, pos + 2&)
                Exit Function
            End If

            pos = pos + 2& + optLen
        End If
    Loop
End Function

Private Sub UserNet_SendDhcpReply(ByRef clientMac() As Byte, ByVal xid As Long, ByVal flags As Long, ByVal msgType As Byte)
    Dim dhcpLen As Long
    Dim frameLen As Long
    Dim dhcp() As Byte
    Dim frame() As Byte
    Dim pos As Long
    Dim i As Long
    Dim ipCsum As Long
    Dim doBroadcast As Byte

    dhcpLen = 300&
    frameLen = 14& + 20& + 8& + dhcpLen

    ReDim dhcp(0 To dhcpLen - 1&) As Byte
    ReDim frame(0 To frameLen - 1&) As Byte

    dhcp(0&) = 2&
    dhcp(1&) = 1&
    dhcp(2&) = 6&
    dhcp(3&) = 0&
    UserNet_WriteBE32 dhcp, 4&, xid
    UserNet_WriteBE16 dhcp, 10&, flags And &HFFFF&

    dhcp(16&) = usernet_offerIp(0&)
    dhcp(17&) = usernet_offerIp(1&)
    dhcp(18&) = usernet_offerIp(2&)
    dhcp(19&) = usernet_offerIp(3&)

    dhcp(20&) = usernet_gatewayIp(0&)
    dhcp(21&) = usernet_gatewayIp(1&)
    dhcp(22&) = usernet_gatewayIp(2&)
    dhcp(23&) = usernet_gatewayIp(3&)

    For i = 0& To 5&
        dhcp(28& + i) = clientMac(i)
    Next i

    dhcp(236&) = &H63&
    dhcp(237&) = &H82&
    dhcp(238&) = &H53&
    dhcp(239&) = &H63&

    pos = 240&
    UserNet_DhcpOptByte dhcp, pos, 53&, msgType
    UserNet_DhcpOpt4 dhcp, pos, 54&, usernet_gatewayIp(0&), usernet_gatewayIp(1&), usernet_gatewayIp(2&), usernet_gatewayIp(3&)
    UserNet_DhcpOpt4 dhcp, pos, 1&, usernet_netmask(0&), usernet_netmask(1&), usernet_netmask(2&), usernet_netmask(3&)
    UserNet_DhcpOpt4 dhcp, pos, 3&, usernet_gatewayIp(0&), usernet_gatewayIp(1&), usernet_gatewayIp(2&), usernet_gatewayIp(3&)
    UserNet_DhcpOpt4 dhcp, pos, 6&, usernet_dnsIp(0&), usernet_dnsIp(1&), usernet_dnsIp(2&), usernet_dnsIp(3&)
    UserNet_DhcpOpt4 dhcp, pos, 28&, usernet_broadcastIp(0&), usernet_broadcastIp(1&), usernet_broadcastIp(2&), usernet_broadcastIp(3&)
    UserNet_DhcpOpt32 dhcp, pos, 51&, 86400&
    UserNet_DhcpOpt32 dhcp, pos, 58&, 43200&
    UserNet_DhcpOpt32 dhcp, pos, 59&, 75600&
    dhcp(pos) = 255&

    doBroadcast = CByte((flags And &H8000&) <> 0&)

    For i = 0& To 5&
        If doBroadcast <> 0& Then
            frame(i) = &HFF&
        Else
            frame(i) = clientMac(i)
        End If
        frame(6& + i) = usernet_gatewayMac(i)
    Next i

    frame(12&) = &H8&
    frame(13&) = 0&

    frame(14&) = &H45&
    frame(15&) = 0&
    UserNet_WriteBE16 frame, 16&, 20& + 8& + dhcpLen
    UserNet_WriteBE16 frame, 18&, usernet_ipId
    usernet_ipId = (usernet_ipId + 1&) And &HFFFF&
    UserNet_WriteBE16 frame, 20&, 0&
    frame(22&) = 64&
    frame(23&) = USERNET_IPPROTO_UDP
    frame(24&) = 0&
    frame(25&) = 0&

    frame(26&) = usernet_gatewayIp(0&)
    frame(27&) = usernet_gatewayIp(1&)
    frame(28&) = usernet_gatewayIp(2&)
    frame(29&) = usernet_gatewayIp(3&)

    If doBroadcast <> 0& Then
        frame(30&) = &HFF&
        frame(31&) = &HFF&
        frame(32&) = &HFF&
        frame(33&) = &HFF&
    Else
        frame(30&) = usernet_offerIp(0&)
        frame(31&) = usernet_offerIp(1&)
        frame(32&) = usernet_offerIp(2&)
        frame(33&) = usernet_offerIp(3&)
    End If

    ipCsum = UserNet_Checksum(frame, 14&, 20&)
    UserNet_WriteBE16 frame, 24&, ipCsum

    UserNet_WriteBE16 frame, 34&, USERNET_DHCP_SERVER_PORT
    UserNet_WriteBE16 frame, 36&, USERNET_DHCP_CLIENT_PORT
    UserNet_WriteBE16 frame, 38&, 8& + dhcpLen
    UserNet_WriteBE16 frame, 40&, 0&

    For i = 0& To dhcpLen - 1&
        frame(42& + i) = dhcp(i)
    Next i
    UserNet_DeliverFrame frame, frameLen
End Sub

Private Sub UserNet_HandleArp(ByRef data() As Byte, ByVal frameLen As Long)
    Dim opcode As Long
    Dim senderMac(0 To 5) As Byte
    Dim senderIp(0 To 3) As Byte
    Dim targetIp(0 To 3) As Byte
    Dim targetIpBe As Long
    Dim i As Long

    If frameLen < 42& Then Exit Sub
    If UserNet_ReadBE16(data, 14&) <> 1& Then Exit Sub
    If UserNet_ReadBE16(data, 16&) <> USERNET_ETHERTYPE_IPV4 Then Exit Sub
    If UserNet_Byte(data, 18&) <> 6& Then Exit Sub
    If UserNet_Byte(data, 19&) <> 4& Then Exit Sub

    opcode = UserNet_ReadBE16(data, 20&)
    If opcode <> 1& Then Exit Sub

    targetIp(0&) = UserNet_Byte(data, 38&)
    targetIp(1&) = UserNet_Byte(data, 39&)
    targetIp(2&) = UserNet_Byte(data, 40&)
    targetIp(3&) = UserNet_Byte(data, 41&)
    targetIpBe = UserNet_MakeIpBE(targetIp(0&), targetIp(1&), targetIp(2&), targetIp(3&))

    If UserNet_IsServiceIp(targetIpBe) = False Then Exit Sub

    For i = 0& To 5&
        senderMac(i) = UserNet_Byte(data, 22& + i)
    Next i

    For i = 0& To 3&
        senderIp(i) = UserNet_Byte(data, 28& + i)
    Next i

    UserNet_CopyGuestMac data, 22&
    UserNet_SendArpReply senderMac, senderIp, targetIp(0&), targetIp(1&), targetIp(2&), targetIp(3&)
End Sub

Private Sub UserNet_HandleIcmpEcho(ByRef data() As Byte, ByVal frameLen As Long, ByVal ihl As Long, ByVal totalLen As Long)
    Dim icmpOff As Long
    Dim icmpLen As Long
    Dim reply() As Byte
    Dim i As Long
    Dim ipCsum As Long
    Dim icmpCsum As Long

    If totalLen < (ihl + 8&) Then Exit Sub
    icmpOff = 14& + ihl
    If (icmpOff + 8&) > frameLen Then Exit Sub

    If UserNet_Byte(data, icmpOff) <> 8& Then Exit Sub

    ReDim reply(0 To (14& + totalLen) - 1&) As Byte
    For i = 0& To UBound(reply)
        reply(i) = UserNet_Byte(data, i)
    Next i

    For i = 0& To 5&
        reply(i) = UserNet_Byte(data, 6& + i)
        reply(6& + i) = usernet_gatewayMac(i)
    Next i

    reply(22&) = 64&

    reply(26&) = usernet_gatewayIp(0&)
    reply(27&) = usernet_gatewayIp(1&)
    reply(28&) = usernet_gatewayIp(2&)
    reply(29&) = usernet_gatewayIp(3&)

    reply(30&) = UserNet_Byte(data, 26&)
    reply(31&) = UserNet_Byte(data, 27&)
    reply(32&) = UserNet_Byte(data, 28&)
    reply(33&) = UserNet_Byte(data, 29&)

    reply(24&) = 0&
    reply(25&) = 0&
    ipCsum = UserNet_Checksum(reply, 14&, ihl)
    UserNet_WriteBE16 reply, 24&, ipCsum

    reply(icmpOff) = 0&
    reply(icmpOff + 2&) = 0&
    reply(icmpOff + 3&) = 0&
    icmpLen = totalLen - ihl
    icmpCsum = UserNet_Checksum(reply, icmpOff, icmpLen)
    UserNet_WriteBE16 reply, icmpOff + 2&, icmpCsum

    UserNet_DeliverFrame reply, UBound(reply) + 1&
End Sub

Private Sub UserNet_HandleIcmpExternalEcho(ByRef data() As Byte, ByVal frameLen As Long, ByVal ihl As Long, ByVal totalLen As Long)
    Dim icmpOff As Long
    Dim icmpLen As Long
    Dim payloadLen As Long
    Dim srcIpBe As Long
    Dim dstIpBe As Long
    Dim reqPayload() As Byte
    Dim reqByte As Byte
    Dim replyBuf() As Byte
    Dim replyBufLen As Long
    Dim echoReply As USERNET_ICMP_ECHO_REPLY_t
    Dim ret As Long
    Dim reply() As Byte
    Dim ttl As Long
    Dim ipCsum As Long
    Dim icmpCsum As Long
    Dim i As Long

    If usernet_icmpHandle = 0& Then Exit Sub

    If totalLen < (ihl + 8&) Then Exit Sub
    icmpOff = 14& + ihl
    If (icmpOff + 8&) > frameLen Then Exit Sub

    If UserNet_Byte(data, icmpOff) <> 8& Then Exit Sub
    If UserNet_Byte(data, icmpOff + 1&) <> 0& Then Exit Sub

    srcIpBe = UserNet_ReadBE32(data, 26&)
    dstIpBe = UserNet_ReadBE32(data, 30&)

    If srcIpBe <> usernet_offerIpBE Then Exit Sub
    If UserNet_IsServiceIp(dstIpBe) <> 0& Then Exit Sub

    icmpLen = totalLen - ihl
    If icmpLen < 8& Then Exit Sub

    payloadLen = icmpLen - 8&
    replyBufLen = LenB(echoReply) + payloadLen + 8&
    If replyBufLen < (LenB(echoReply) + 8&) Then
        replyBufLen = LenB(echoReply) + 8&
    End If

    ReDim replyBuf(0 To replyBufLen - 1&) As Byte

    If payloadLen > 0& Then
        ReDim reqPayload(0 To payloadLen - 1&) As Byte
        For i = 0& To payloadLen - 1&
            reqPayload(i) = UserNet_Byte(data, icmpOff + 8& + i)
        Next i
        ret = IcmpSendEcho(usernet_icmpHandle, UserNet_BSwap32(dstIpBe), reqPayload(0), payloadLen, 0&, replyBuf(0), replyBufLen, USERNET_ICMP_TIMEOUT_MS)
    Else
        reqByte = 0&
        ret = IcmpSendEcho(usernet_icmpHandle, UserNet_BSwap32(dstIpBe), reqByte, 0&, 0&, replyBuf(0), replyBufLen, USERNET_ICMP_TIMEOUT_MS)
    End If

    If ret <= 0& Then Exit Sub

    CopyMemory echoReply, replyBuf(0), LenB(echoReply)
    If echoReply.status <> USERNET_IP_SUCCESS Then Exit Sub

    UserNet_CopyGuestMac data, 6&

    ReDim reply(0 To (14& + totalLen) - 1&) As Byte
    For i = 0& To UBound(reply)
        reply(i) = UserNet_Byte(data, i)
    Next i

    For i = 0& To 5&
        reply(i) = UserNet_Byte(data, 6& + i)
        reply(6& + i) = usernet_gatewayMac(i)
    Next i

    ttl = CLng(echoReply.Options.Ttl) And &HFF&
    If ttl <= 0& Then ttl = 64&
    reply(22&) = CByte(ttl And &HFF&)

    UserNet_WriteBE32 reply, 26&, dstIpBe
    UserNet_WriteBE32 reply, 30&, srcIpBe

    reply(24&) = 0&
    reply(25&) = 0&
    ipCsum = UserNet_Checksum(reply, 14&, ihl)
    UserNet_WriteBE16 reply, 24&, ipCsum

    reply(icmpOff) = 0&
    reply(icmpOff + 1&) = 0&
    reply(icmpOff + 2&) = 0&
    reply(icmpOff + 3&) = 0&

    icmpCsum = UserNet_Checksum(reply, icmpOff, icmpLen)
    UserNet_WriteBE16 reply, icmpOff + 2&, icmpCsum

    UserNet_DeliverFrame reply, UBound(reply) + 1&
End Sub

Private Sub UserNet_HandleDhcp(ByRef data() As Byte, ByVal frameLen As Long, ByVal ihl As Long, ByVal totalLen As Long)
    Dim udpOff As Long
    Dim dhcpOff As Long
    Dim udpLen As Long
    Dim udpPayloadLen As Long
    Dim optStart As Long
    Dim optEnd As Long
    Dim msgType As Byte
    Dim xid As Long
    Dim flags As Long
    Dim clientMac(0 To 5) As Byte
    Dim i As Long

    udpOff = 14& + ihl
    If (udpOff + 8&) > frameLen Then Exit Sub

    If UserNet_ReadBE16(data, udpOff) <> USERNET_DHCP_CLIENT_PORT Then Exit Sub
    If UserNet_ReadBE16(data, udpOff + 2&) <> USERNET_DHCP_SERVER_PORT Then Exit Sub

    udpLen = UserNet_ReadBE16(data, udpOff + 4&)
    If udpLen < 8& Then Exit Sub

    dhcpOff = udpOff + 8&
    udpPayloadLen = udpLen - 8&
    If udpPayloadLen < 240& Then Exit Sub

    If (dhcpOff + udpPayloadLen) > (14& + totalLen) Then
        udpPayloadLen = (14& + totalLen) - dhcpOff
        If udpPayloadLen < 240& Then Exit Sub
    End If

    If (dhcpOff + udpPayloadLen) > frameLen Then Exit Sub

    If UserNet_Byte(data, dhcpOff + 236&) <> &H63& Then Exit Sub
    If UserNet_Byte(data, dhcpOff + 237&) <> &H82& Then Exit Sub
    If UserNet_Byte(data, dhcpOff + 238&) <> &H53& Then Exit Sub
    If UserNet_Byte(data, dhcpOff + 239&) <> &H63& Then Exit Sub

    msgType = UserNet_DhcpMessageType(data, dhcpOff + 240&, dhcpOff + udpPayloadLen)
    If (msgType <> USERNET_DHCPDISCOVER) And (msgType <> USERNET_DHCPREQUEST) Then Exit Sub

    xid = UserNet_ReadBE32(data, dhcpOff + 4&)
    flags = UserNet_ReadBE16(data, dhcpOff + 10&)

    For i = 0& To 5&
        clientMac(i) = UserNet_Byte(data, dhcpOff + 28& + i)
    Next i

    UserNet_CopyGuestMac data, dhcpOff + 28&

    If msgType = USERNET_DHCPDISCOVER Then
        UserNet_SendDhcpReply clientMac, xid, flags, USERNET_DHCPOFFER
    ElseIf msgType = USERNET_DHCPREQUEST Then
        UserNet_SendDhcpReply clientMac, xid, flags, USERNET_DHCPACK
    End If
End Sub

Private Sub UserNet_UdpCloseFlow(ByVal idx As Long)
    Dim emptyFlow As USERNET_UDP_FLOW_t

    If (idx < 0&) Or (idx >= USERNET_UDP_MAX_FLOWS) Then Exit Sub
    If usernet_udpFlows(idx).used = 0& Then Exit Sub

    If (usernet_udpFlows(idx).hostSock <> 0&) And (usernet_udpFlows(idx).hostSock <> INVALID_SOCKET) Then
        closesocket usernet_udpFlows(idx).hostSock
    End If

    usernet_udpFlows(idx) = emptyFlow
End Sub

Private Sub UserNet_UdpReset()
    Dim i As Long

    For i = 0& To USERNET_UDP_MAX_FLOWS - 1&
        UserNet_UdpCloseFlow i
    Next i
End Sub

Private Function UserNet_UdpStartup() As Long
    Dim ret As Long

    If usernet_wsaStarted <> 0& Then
        UserNet_UdpStartup = 0&
        Exit Function
    End If

    ret = WSAStartup(&H202&, usernet_wsa)
    If ret <> 0& Then
        debug_log DEBUG_ERROR, "[USERNET] WSAStartup failed with error " & CStr(ret)
        UserNet_UdpStartup = -1&
        Exit Function
    End If

    usernet_wsaStarted = 1&
    UserNet_UdpStartup = 0&
End Function

Private Function UserNet_IcmpStartup() As Long
    If usernet_icmpHandle <> 0& Then
        UserNet_IcmpStartup = 0&
        Exit Function
    End If

    usernet_icmpHandle = IcmpCreateFile()
    If (usernet_icmpHandle = 0&) Or (usernet_icmpHandle = INVALID_SOCKET) Then
        usernet_icmpHandle = 0&
        UserNet_IcmpStartup = -1&
        Exit Function
    End If

    UserNet_IcmpStartup = 0&
End Function

Private Sub UserNet_IcmpShutdown()
    If usernet_icmpHandle <> 0& Then
        IcmpCloseHandle usernet_icmpHandle
        usernet_icmpHandle = 0&
    End If
End Sub

Private Sub UserNet_UdpShutdown()
    UserNet_UdpReset

    If usernet_wsaStarted <> 0& Then
        WSACleanup
        usernet_wsaStarted = 0&
    End If
End Sub

Private Function UserNet_UdpFindFlow(ByVal guestSrcIp As Long, ByVal guestDstIp As Long, ByVal guestSrcPort As Long, ByVal guestDstPort As Long, ByVal hostDstIp As Long, ByVal hostDstPort As Long, ByVal dnsProxy As Byte) As Long
    Dim i As Long

    UserNet_UdpFindFlow = -1&

    For i = 0& To USERNET_UDP_MAX_FLOWS - 1&
        If usernet_udpFlows(i).used <> 0& Then
            If usernet_udpFlows(i).guestSrcIp = guestSrcIp And _
               usernet_udpFlows(i).guestDstIp = guestDstIp And _
               usernet_udpFlows(i).guestSrcPort = guestSrcPort And _
               usernet_udpFlows(i).guestDstPort = guestDstPort And _
               usernet_udpFlows(i).hostDstIp = hostDstIp And _
               usernet_udpFlows(i).hostDstPort = hostDstPort And _
               usernet_udpFlows(i).dnsProxy = dnsProxy Then
                UserNet_UdpFindFlow = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function UserNet_UdpAllocFlow(ByVal guestSrcIp As Long, ByVal guestDstIp As Long, ByVal guestSrcPort As Long, ByVal guestDstPort As Long, ByVal hostDstIp As Long, ByVal hostDstPort As Long, ByVal dnsProxy As Byte) As Long
    Dim i As Long
    Dim sock As Long
    Dim mode As Long
    Dim localAddr As SOCKADDR_IN
    Dim ret As Long

    UserNet_UdpAllocFlow = -1&

    For i = 0& To USERNET_UDP_MAX_FLOWS - 1&
        If usernet_udpFlows(i).used = 0& Then
            sock = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP_SOCKET)
            If sock = INVALID_SOCKET Then Exit Function

            mode = 1&
            ioctlsocket sock, FIONBIO, mode

            localAddr.sin_family = AF_INET
            localAddr.sin_port = 0&
            localAddr.sin_addr.s_addr = INADDR_ANY

            ret = bind(sock, localAddr, LenB(localAddr))
            If ret = SOCKET_ERROR Then
                closesocket sock
                Exit Function
            End If

            usernet_udpFlows(i).used = 1&
            usernet_udpFlows(i).dnsProxy = dnsProxy
            usernet_udpFlows(i).hostSock = sock
            usernet_udpFlows(i).guestSrcIp = guestSrcIp
            usernet_udpFlows(i).guestDstIp = guestDstIp
            usernet_udpFlows(i).hostDstIp = hostDstIp
            usernet_udpFlows(i).guestSrcPort = guestSrcPort
            usernet_udpFlows(i).guestDstPort = guestDstPort
            usernet_udpFlows(i).hostDstPort = hostDstPort
            usernet_udpFlows(i).lastSeen = timing_getCur()

            UserNet_UdpAllocFlow = i
            Exit Function
        End If
    Next i
End Function

Private Sub UserNet_SendUdpToGuest(ByVal dstIpBe As Long, ByVal srcIpBe As Long, ByVal dstPort As Long, ByVal srcPort As Long, ByRef payload() As Byte, ByVal payloadLen As Long)
    Dim frameLen As Long
    Dim frame() As Byte
    Dim i As Long
    Dim ipCsum As Long

    If usernet_guestMacKnown = 0& Then Exit Sub
    If payloadLen < 0& Then Exit Sub

    frameLen = 14& + 20& + 8& + payloadLen
    If frameLen > USERNET_MAX_FRAME Then Exit Sub

    ReDim frame(0 To frameLen - 1&) As Byte

    For i = 0& To 5&
        frame(i) = usernet_guestMac(i)
        frame(6& + i) = usernet_gatewayMac(i)
    Next i

    frame(12&) = &H8&
    frame(13&) = 0&

    frame(14&) = &H45&
    frame(15&) = 0&
    UserNet_WriteBE16 frame, 16&, 20& + 8& + payloadLen
    UserNet_WriteBE16 frame, 18&, usernet_ipId
    usernet_ipId = (usernet_ipId + 1&) And &HFFFF&
    UserNet_WriteBE16 frame, 20&, 0&
    frame(22&) = 64&
    frame(23&) = USERNET_IPPROTO_UDP
    frame(24&) = 0&
    frame(25&) = 0&

    UserNet_WriteBE32 frame, 26&, srcIpBe
    UserNet_WriteBE32 frame, 30&, dstIpBe

    ipCsum = UserNet_Checksum(frame, 14&, 20&)
    UserNet_WriteBE16 frame, 24&, ipCsum

    UserNet_WriteBE16 frame, 34&, srcPort
    UserNet_WriteBE16 frame, 36&, dstPort
    UserNet_WriteBE16 frame, 38&, 8& + payloadLen
    UserNet_WriteBE16 frame, 40&, 0&

    For i = 0& To payloadLen - 1&
        frame(42& + i) = payload(i)
    Next i

    UserNet_DeliverFrame frame, frameLen
End Sub

Private Sub UserNet_HandleUdpNat(ByRef data() As Byte, ByVal frameLen As Long, ByVal ihl As Long, ByVal totalLen As Long)
    Dim udpOff As Long
    Dim udpLen As Long
    Dim udpPayloadLen As Long
    Dim srcIpBe As Long
    Dim dstIpBe As Long
    Dim srcPort As Long
    Dim dstPort As Long
    Dim hostDstIpBe As Long
    Dim hostDstPort As Long
    Dim dnsProxy As Byte
    Dim flowIdx As Long
    Dim payload() As Byte
    Dim dnsReply() As Byte
    Dim dnsReplyLen As Long
    Dim i As Long
    Dim ret As Long
    Dim wsaErr As Long
    Dim toAddr As SOCKADDR_IN

    If usernet_wsaStarted = 0& Then Exit Sub

    udpOff = 14& + ihl
    If (udpOff + 8&) > frameLen Then Exit Sub

    udpLen = UserNet_ReadBE16(data, udpOff + 4&)
    If udpLen < 8& Then Exit Sub

    If (udpOff + udpLen) > (14& + totalLen) Then
        udpLen = (14& + totalLen) - udpOff
    End If
    If (udpOff + udpLen) > frameLen Then
        udpLen = frameLen - udpOff
    End If
    If udpLen < 8& Then Exit Sub

    udpPayloadLen = udpLen - 8&
    If udpPayloadLen <= 0& Then Exit Sub

    srcIpBe = UserNet_ReadBE32(data, 26&)
    dstIpBe = UserNet_ReadBE32(data, 30&)
    If srcIpBe <> usernet_offerIpBE Then Exit Sub

    srcPort = UserNet_ReadBE16(data, udpOff)
    dstPort = UserNet_ReadBE16(data, udpOff + 2&)
    If (srcPort <= 0&) Or (dstPort <= 0&) Then Exit Sub

    dnsProxy = 0&
    hostDstIpBe = dstIpBe
    hostDstPort = dstPort

    If (dstPort = 53&) And ((dstIpBe = usernet_dnsIpBE) Or (dstIpBe = usernet_gatewayIpBE)) Then
        dnsProxy = 1&
        hostDstIpBe = usernet_dnsUpstreamIpBE
        hostDstPort = 53&
    ElseIf dstIpBe = usernet_gatewayIpBE Then
        Exit Sub
    End If

    UserNet_CopyGuestMac data, 6&

    ReDim payload(0 To udpPayloadLen - 1&) As Byte
    For i = 0& To udpPayloadLen - 1&
        payload(i) = UserNet_Byte(data, udpOff + 8& + i)
    Next i

    If dnsProxy <> 0& Then
        If UserNet_DnsResolveLocal(payload, udpPayloadLen, dnsReply, dnsReplyLen) <> 0& Then
            If dnsReplyLen > 0& Then
                UserNet_SendUdpToGuest srcIpBe, dstIpBe, srcPort, dstPort, dnsReply, dnsReplyLen
                UserNet_SendUdpToGuest srcIpBe, dstIpBe, srcPort, dstPort, dnsReply, dnsReplyLen
                Exit Sub
            End If
        End If
    End If

    flowIdx = UserNet_UdpFindFlow(srcIpBe, dstIpBe, srcPort, dstPort, hostDstIpBe, hostDstPort, dnsProxy)
    If flowIdx < 0& Then
        flowIdx = UserNet_UdpAllocFlow(srcIpBe, dstIpBe, srcPort, dstPort, hostDstIpBe, hostDstPort, dnsProxy)
        If flowIdx < 0& Then Exit Sub
    End If

    toAddr.sin_family = AF_INET
    toAddr.sin_port = htons(CInt(hostDstPort And &HFFFF&))
    toAddr.sin_addr.s_addr = UserNet_BSwap32(hostDstIpBe)

    ret = sendto(usernet_udpFlows(flowIdx).hostSock, payload(0), udpPayloadLen, 0&, toAddr, LenB(toAddr))
    If ret = SOCKET_ERROR Then
        wsaErr = WSAGetLastError()
        If wsaErr <> WSAEWOULDBLOCK Then
            UserNet_UdpCloseFlow flowIdx
        End If
        Exit Sub
    End If

    usernet_udpFlows(flowIdx).lastSeen = timing_getCur()
End Sub

Private Sub UserNet_UdpPoll()
    Dim idx As Long
    Dim ret As Long
    Dim wsaErr As Long
    Dim nowTicks As Double
    Dim fromAddr As SOCKADDR_IN
    Dim fromLen As Long
    Dim rxBuf(0 To USERNET_UDP_MAX_PAYLOAD - 1) As Byte
    Dim payload() As Byte
    Dim i As Long
    Dim srcIpBe As Long
    Dim srcPort As Long

    If usernet_wsaStarted = 0& Then Exit Sub

    nowTicks = timing_getCur()

    For idx = 0& To USERNET_UDP_MAX_FLOWS - 1&
        If usernet_udpFlows(idx).used <> 0& Then
            If (nowTicks - usernet_udpFlows(idx).lastSeen) > UserNet_FlowTimeoutTicks() Then
                UserNet_UdpCloseFlow idx
            Else
                Do
                    fromLen = LenB(fromAddr)
                    ret = recvfrom(usernet_udpFlows(idx).hostSock, rxBuf(0), USERNET_UDP_MAX_PAYLOAD, 0&, fromAddr, fromLen)
                    If ret = SOCKET_ERROR Then
                        wsaErr = WSAGetLastError()
                        If wsaErr <> WSAEWOULDBLOCK Then
                            UserNet_UdpCloseFlow idx
                        End If
                        Exit Do
                    End If

                    If ret <= 0& Then Exit Do

                    usernet_udpFlows(idx).lastSeen = timing_getCur()

                    ReDim payload(0 To ret - 1&) As Byte
                    For i = 0& To ret - 1&
                        payload(i) = rxBuf(i)
                    Next i

                    If usernet_udpFlows(idx).dnsProxy <> 0& Then
                        srcIpBe = usernet_udpFlows(idx).guestDstIp
                        srcPort = usernet_udpFlows(idx).guestDstPort
                    Else
                        srcIpBe = UserNet_BSwap32(fromAddr.sin_addr.s_addr)
                        srcPort = CLng(ntohs(fromAddr.sin_port)) And &HFFFF&
                        If srcPort <= 0& Then srcPort = usernet_udpFlows(idx).hostDstPort
                    End If

                    UserNet_SendUdpToGuest usernet_udpFlows(idx).guestSrcIp, srcIpBe, usernet_udpFlows(idx).guestSrcPort, srcPort, payload, ret
                Loop
            End If
        End If
    Next idx
End Sub
Private Sub UserNet_HandleIPv4(ByRef data() As Byte, ByVal frameLen As Long)
    Dim ihl As Long
    Dim totalLen As Long
    Dim proto As Long
    Dim udpOff As Long
    Dim udpSrcPort As Long
    Dim udpDstPort As Long

    If frameLen < 34& Then Exit Sub

    ihl = (CLng(UserNet_Byte(data, 14&)) And &HF&) * 4&
    If ihl < 20& Then Exit Sub
    If frameLen < (14& + ihl) Then Exit Sub

    totalLen = UserNet_ReadBE16(data, 16&)
    If totalLen < ihl Then Exit Sub

    If totalLen > (frameLen - 14&) Then
        totalLen = frameLen - 14&
    End If

    proto = CLng(UserNet_Byte(data, 23&))

    If proto = USERNET_IPPROTO_ICMP Then
        If UserNet_IpMatches(data, 30&, usernet_gatewayIp(0&), usernet_gatewayIp(1&), usernet_gatewayIp(2&), usernet_gatewayIp(3&)) Then
            UserNet_HandleIcmpEcho data, frameLen, ihl, totalLen
        Else
            UserNet_HandleIcmpExternalEcho data, frameLen, ihl, totalLen
        End If
    ElseIf proto = USERNET_IPPROTO_UDP Then
        udpOff = 14& + ihl
        If (udpOff + 8&) > frameLen Then Exit Sub

        udpSrcPort = UserNet_ReadBE16(data, udpOff)
        udpDstPort = UserNet_ReadBE16(data, udpOff + 2&)

        If (udpSrcPort = USERNET_DHCP_CLIENT_PORT) And (udpDstPort = USERNET_DHCP_SERVER_PORT) Then
            UserNet_HandleDhcp data, frameLen, ihl, totalLen
        Else
            UserNet_HandleUdpNat data, frameLen, ihl, totalLen
        End If
    End If
End Sub

Public Function usernet_init(ByVal devId As Long) As Long
    usernet_devId = devId
    usernet_active = 1&
    usernet_guestMacKnown = 0&
    usernet_ipId = 1&

    usernet_gatewayMac(0&) = &H52&
    usernet_gatewayMac(1&) = &H54&
    usernet_gatewayMac(2&) = 0&
    usernet_gatewayMac(3&) = &H12&
    usernet_gatewayMac(4&) = &H34&
    usernet_gatewayMac(5&) = &H56&

    usernet_gatewayIp(0&) = 10&
    usernet_gatewayIp(1&) = 0&
    usernet_gatewayIp(2&) = 2&
    usernet_gatewayIp(3&) = 2&

    usernet_offerIp(0&) = 10&
    usernet_offerIp(1&) = 0&
    usernet_offerIp(2&) = 2&
    usernet_offerIp(3&) = 15&

    usernet_dnsIp(0&) = 10&
    usernet_dnsIp(1&) = 0&
    usernet_dnsIp(2&) = 2&
    usernet_dnsIp(3&) = 3&

    usernet_netmask(0&) = &HFF&
    usernet_netmask(1&) = &HFF&
    usernet_netmask(2&) = &HFF&
    usernet_netmask(3&) = 0&

    usernet_broadcastIp(0&) = 10&
    usernet_broadcastIp(1&) = 0&
    usernet_broadcastIp(2&) = 2&
    usernet_broadcastIp(3&) = &HFF&

    usernet_gatewayIpBE = UserNet_MakeIpBE(usernet_gatewayIp(0&), usernet_gatewayIp(1&), usernet_gatewayIp(2&), usernet_gatewayIp(3&))
    usernet_offerIpBE = UserNet_MakeIpBE(usernet_offerIp(0&), usernet_offerIp(1&), usernet_offerIp(2&), usernet_offerIp(3&))
    usernet_dnsIpBE = UserNet_MakeIpBE(usernet_dnsIp(0&), usernet_dnsIp(1&), usernet_dnsIp(2&), usernet_dnsIp(3&))
    usernet_dnsUpstreamIpBE = UserNet_MakeIpBE(8&, 8&, 8&, 8&)

    UserNet_ResetQueue
    UserNet_UdpReset

    If UserNet_UdpStartup() <> 0& Then
        debug_log DEBUG_INFO, "[USERNET] UDP NAT unavailable (Winsock startup failed); DHCP/ARP/ICMP will still work."
    Else
        debug_log DEBUG_INFO, "[USERNET] UDP NAT enabled (DNS proxy 10.0.2.3 -> 8.8.8.8)."
    End If

    If UserNet_IcmpStartup() <> 0& Then
        debug_log DEBUG_INFO, "[USERNET] External ICMP relay unavailable (IcmpCreateFile failed)."
    Else
        debug_log DEBUG_INFO, "[USERNET] External ICMP relay enabled."
    End If

    debug_log DEBUG_INFO, "[USERNET] Enabled built-in usermode gateway (DHCP/ARP/ICMP/UDP-NAT)."
    debug_log DEBUG_INFO, "[USERNET] Guest network: 10.0.2.0/24, gateway 10.0.2.2, DHCP lease 10.0.2.15"

    usernet_init = 0&
End Function

Public Sub usernet_shutdown()
    usernet_active = 0&
    usernet_guestMacKnown = 0&
    UserNet_IcmpShutdown
    UserNet_UdpShutdown
    UserNet_ResetQueue
End Sub

Public Sub usernet_txPacket(ByRef data() As Byte, ByVal length As Long)
    Dim etherType As Long

    If usernet_active = 0& Then Exit Sub
    If length < 14& Then Exit Sub

    etherType = UserNet_ReadBE16(data, 12&)

    Select Case etherType
        Case USERNET_ETHERTYPE_ARP
            UserNet_HandleArp data, length
        Case USERNET_ETHERTYPE_IPV4
            UserNet_HandleIPv4 data, length
    End Select
End Sub


Public Sub usernet_pollPackets()
    If usernet_active = 0& Then Exit Sub
    If usernet_devId < 0& Then Exit Sub

    UserNet_DrainRxQueue

    If usernet_rxCount > 0& Then Exit Sub

    UserNet_UdpPoll
    UserNet_DrainRxQueue
End Sub





