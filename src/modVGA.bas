Attribute VB_Name = "modVGA"
Option Explicit

Private Type VGADAC_t
    state As Byte
    index As Byte
    step As Byte
    pal(0& To 255&, 0& To 2&) As Byte
End Type

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const VGA_FB_MAX As Long = 1024&
Private Const VGA_FB_PIXELS As Long = VGA_FB_MAX * VGA_FB_MAX
Private Const VGA_RAM_PLANES As Long = 4&
Private Const VGA_RAM_PLANE_SIZE As Long = 65536

Public Const VGA_DAC_MODE_READ As Byte = 0&
Public Const VGA_DAC_MODE_WRITE As Byte = 3&

Public Const VGA_REG_DATA_CURSOR_BEGIN As Byte = &HA&
Public Const VGA_REG_DATA_CURSOR_END As Byte = &HB&

Public Const VGA_MODE_TEXT As Byte = 0&
Public Const VGA_MODE_GRAPHICS_8BPP As Byte = 1&
Public Const VGA_MODE_GRAPHICS_4BPP As Byte = 2&
Public Const VGA_MODE_GRAPHICS_2BPP As Byte = 3&
Public Const VGA_MODE_GRAPHICS_1BPP As Byte = 4&

Private VBIOS(0& To 32767&) As Byte
Public vga_palette(0& To 255&, 0& To 2&) As Byte

Private vga_DAC As VGADAC_t
Private vga_framebuffer(0& To VGA_FB_PIXELS - 1&) As Long
Private vga_textColorCache(0& To 15&) As Long
Private vga_glyphBitMask(0& To 7&) As Byte
Private vga_attrColorCache(0& To 15&) As Long

Public vga_w As Long
Public vga_h As Long
Private vga_dots As Long
Private vga_membase As Long
Private vga_memmask As Long

Private vga_cursorloc As Long
Private vga_dbl As Byte
Private vga_crtci As Byte
Private vga_crtcd(0& To &H18&) As Byte
Private vga_attri As Byte
Private vga_attrd(0& To &H14&) As Byte
Private vga_attrflipflop As Byte
Private vga_attrpal As Byte
Private vga_gfxi As Byte
Private vga_gfxd(0& To &H8&) As Byte
Private vga_seqi As Byte
Private vga_seqd(0& To &H4&) As Byte
Private vga_misc As Byte
Private vga_status0 As Byte
Private vga_status1 As Byte
Private vga_cursor_blink_state As Byte

Private vga_wmode As Byte
Private vga_rmode As Byte
Private vga_shiftmode As Byte
Private vga_rotate As Byte
Private vga_logicop As Byte
Private vga_enableplane As Byte
Private vga_readmap As Byte
Private vga_scandbl As Byte
Private vga_hdbl As Byte
Private vga_bpp As Byte
Private vga_latch(0& To 3&) As Byte

Private vga_RAM(0& To VGA_RAM_PLANES - 1&, 0& To VGA_RAM_PLANE_SIZE - 1&) As Byte

Private vga_hblankstart As Double
Private vga_hblankend As Double
Private vga_hblanklen As Double
Private vga_dispinterval As Double
Private vga_hblankinterval As Double
Private vga_htotal As Double
Private vga_vblankstart As Double
Private vga_vblankend As Double
Private vga_vblanklen As Double
Private vga_vblankinterval As Double
Private vga_frameinterval As Double

Private vga_doRender As Byte
Private vga_doBlit As Byte
Private vga_targetFPS As Double

Private vga_hblankTimer As Long
Private vga_hblankEndTimer As Long
Private vga_drawTimer As Long
Private vga_curScanline As Long
Private vga_chain4 As Byte
Private vga_modeY As Byte
Private vga_threadStarted As Byte
Private vga_threadHandle As Long
Private vga_threadId As Long
Public vga_doBlitNow As Boolean

Public vga_renderW As Long
Public vga_renderH As Long

Private vga_diagPort3C8Writes As Long
Private vga_diagPort3C9Writes As Long
Private vga_diagMemWrites As Long
Private vga_diagMemNonZeroWrites As Long
Private vga_diagLastMode As Long
Private vga_diag3DAReads As Long
Private vga_diag3DAVBlankReads As Long
Private vga_diag3DAHBlankReads As Long

Public Function vga_init() As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim j As Long
    Dim idx As Long
    Dim vRead As Variant
    Dim vWrite As Variant

    debug_log DEBUG_INFO, "[VGA] Initializing VGA video device"
    diag_mark_vga_init

    vga_w = 640&
    vga_h = 400&
    vga_dots = 8&
    vga_targetFPS = 60#
    vga_attrpal = &H20&
    vga_memmask = &HFFFF&
    vga_glyphBitMask(0&) = &H80&
    vga_glyphBitMask(1&) = &H40&
    vga_glyphBitMask(2&) = &H20&
    vga_glyphBitMask(3&) = &H10&
    vga_glyphBitMask(4&) = &H8&
    vga_glyphBitMask(5&) = &H4&
    vga_glyphBitMask(6&) = &H2&
    vga_glyphBitMask(7&) = &H1&

    For y = 0& To 399&
        idx = y * VGA_FB_MAX
        For x = 0& To 639&
            vga_framebuffer(idx + x) = vga_color(0&)
        Next x
    Next y
    console_blit VarPtr(vga_framebuffer(0&)), 640&, 400&, (VGA_FB_MAX * 4&)

    If vga_lockFPS >= 1# Then
        vga_targetFPS = vga_lockFPS
    End If

    vga_doRender = 0&
    vga_doBlit = 0&
    vga_doBlitNow = False
    vga_diagPort3C8Writes = 0&
    vga_diagPort3C9Writes = 0&
    vga_diagMemWrites = 0&
    vga_diagMemNonZeroWrites = 0&
    vga_diagLastMode = VGA_MODE_TEXT
    vga_diag3DAReads = 0&
    vga_diag3DAVBlankReads = 0&
    vga_diag3DAHBlankReads = 0&

    vga_curScanline = 0&
    vga_drawTimer = timing_addTimer(TIMER_CB_VGA_DRAW, 0&, vga_targetFPS, TIMING_ENABLED)
    timing_addTimer TIMER_CB_VGA_BLINK, 0&, 3.75, TIMING_ENABLED
    vga_hblankTimer = timing_addTimer(TIMER_CB_VGA_HBLANK, 0&, 10000#, TIMING_ENABLED)
    vga_hblankEndTimer = timing_addTimer(TIMER_CB_VGA_HBLANK_END, 0&, 100#, TIMING_ENABLED)

    For i = 0& To VGA_RAM_PLANES - 1&
        For j = 0& To VGA_RAM_PLANE_SIZE - 1&
            vga_RAM(i, j) = 0&
        Next j
    Next i

    ports_cbRegister &H3B4&, 39&, PORTS_CB_VGA, PORTS_CB_NONE, PORTS_CB_VGA, PORTS_CB_NONE, 0&
    memory_mapCallbackRegister &HA0000, &H20000, MEMORY_CB_VGA, MEMORY_CB_VGA, 0&

    If utility_loadFile(VBIOS, 32768, "roms/video/et4000.bin") <> 0& Then
        vga_init = -1&
        Exit Function
    End If

    vRead = VBIOS
    vWrite = Empty
    memory_mapRegister &HC0000, 32768, vRead, vWrite

    vga_threadStarted = 0&
    vga_startRenderThread

    vga_init = 0&
End Function

Public Sub vga_updateScanlineTiming()
    Dim pixelclock As Double
    Static lastw As Long
    Static lasth As Long
    Static lastFPS As Double
    Dim ratio As Double
    Dim vblankStartTmp As Long
    Dim vblankEndTmp As Long

    If (vga_misc And &H4&) <> 0& Then
        pixelclock = 28322000#
    Else
        pixelclock = 25175000#
    End If

    vga_hblankstart = CDbl(vga_crtcd(&H2&)) * CDbl(vga_dots)
    vga_hblankend = (CDbl(vga_crtcd(&H2&)) * CDbl(vga_dots)) + (CDbl(vga_crtcd(&H3&) And &H1F&) + 1#) * CDbl(vga_dots)
    vga_hblanklen = vga_hblankend - vga_hblankstart

    vblankStartTmp = CLng(vga_crtcd(&H10&)) Or CLng((vga_crtcd(&H7&) And &H4&) * 64&) Or CLng((vga_crtcd(&H7&) And &H80&) * 4&)
    vblankEndTmp = CLng(vga_crtcd(&H6&)) Or CLng((vga_crtcd(&H7&) And &H1&) * 256&) Or CLng((vga_crtcd(&H7&) And &H20&) * 16&)
    vga_vblankstart = CDbl(vblankStartTmp)
    vga_vblankend = CDbl(vblankEndTmp)
    vga_vblanklen = vga_vblankend - vga_vblankstart
    vga_htotal = CDbl(vga_crtcd(&H0&))

    If (vga_vblankend > 0#) And ((vga_htotal + 5#) > 0#) And (vga_dots > 0&) Then
        vga_targetFPS = pixelclock / ((vga_htotal + 5#) * CDbl(vga_dots) * vga_vblankend)
    Else
        vga_targetFPS = 60#
    End If

    ratio = timing_getFreq() / pixelclock
    vga_dispinterval = (vga_htotal + 5#) * CDbl(vga_dots) * ratio
    vga_hblankinterval = vga_hblanklen * ratio
    vga_vblankinterval = vga_hblankend * vga_vblanklen * ratio
    vga_frameinterval = vga_hblankend * vga_vblankend * ratio

    If (lastw <> vga_w) Or (lasth <> vga_h) Or (lastFPS <> vga_targetFPS) Then
        debug_log DEBUG_DETAIL, "[VGA] Mode switch: " & CStr(vga_w) & "x" & CStr(vga_h) & " (" & Format$(vga_targetFPS, "0.00") & " Hz)"
        lastw = vga_w
        lasth = vga_h
        lastFPS = vga_targetFPS
    End If

    timing_updateInterval vga_hblankTimer, vga_dispinterval
    timing_updateInterval vga_hblankEndTimer, vga_hblankinterval
    timing_timerEnable vga_hblankTimer
    timing_timerDisable vga_hblankEndTimer

    If vga_lockFPS = 0# Then
        If vga_targetFPS > 0# Then
            timing_updateIntervalFreq vga_drawTimer, vga_targetFPS
        End If
    End If
End Sub

Public Sub vga_update(ByVal start_x As Long, ByVal start_y As Long, ByVal end_x As Long, ByVal end_y As Long)
    Dim addr As Long
    Dim startaddr As Long
    Dim cursorloc As Long
    Dim cursor_x As Long
    Dim cursor_y As Long
    Dim fontbase As Long
    Dim color32 As Long
    Dim scx As Long
    Dim scy As Long
    Dim x As Long
    Dim y As Long
    Dim hchars As Long
    Dim divx As Long
    Dim yscanpixels As Long
    Dim xscanpixels As Long
    Dim xstride As Long
    Dim bpp As Long
    Dim pixelsperbyte As Long
    Dim shift As Long
    Dim cc As Long
    Dim attr As Long
    Dim fontdata As Long
    Dim blink As Long
    Dim mode As Long
    Dim blinkenable As Long
    Dim cursorenable As Long
    Dim dup9 As Long
    Dim maxscan As Long
    Dim charcolumn As Long
    Dim plane As Long
    Dim yadd As Long
    Dim xadd As Long
    Dim isodd As Long
    Dim paletteIdx As Long
    Dim glyphRow As Long
    Dim rowOffset As Long
    Dim cursorBegin As Long
    Dim cursorEnd As Long
    Dim cursorScan As Long
    Dim attrRaw As Long
    Dim attrMasked As Long
    Dim palettePage As Long
    Dim hiPage As Long
    Dim paletteHiMode As Long
    Dim byteX As Long
    Dim srcPixelX As Long
    Dim destX As Long
    Dim rowBase As Long
    Dim plane0Byte As Long
    Dim plane1Byte As Long
    Dim plane2Byte As Long
    Dim plane3Byte As Long
    Dim pixelBit As Long
    Dim byteStartX As Long
    Dim byteEndX As Long
    Dim bitMaskVal As Long

    If start_x < 0& Then start_x = 0&
    If start_y < 0& Then start_y = 0&
    If end_x >= vga_w Then end_x = vga_w - 1&
    If end_y >= vga_h Then end_y = vga_h - 1&
    If end_x > (VGA_FB_MAX - 1&) Then end_x = VGA_FB_MAX - 1&
    If end_y > (VGA_FB_MAX - 1&) Then end_y = VGA_FB_MAX - 1&
    If end_x < start_x Or end_y < start_y Then Exit Sub

    If (vga_attrd(&H10&) And 1&) <> 0& Then
        If (vga_shiftmode And &H2&) <> 0& Then
            xscanpixels = 2&
            yscanpixels = (vga_crtcd(&H9&) And &H1F&) + 1&
        Else
            If (vga_seqd(&H1&) And &H8&) <> 0& Then
                xscanpixels = 2&
            Else
                xscanpixels = 1&
            End If
            If (vga_crtcd(&H9&) And &H80&) <> 0& Then
                yscanpixels = 2&
            Else
                yscanpixels = 1&
            End If
        End If

        Select Case (vga_shiftmode And &H3&)
            Case 0&
                If (vga_attrd(&H12&) And &HF&) = 1& Then
                    bpp = 1&
                    pixelsperbyte = 8&
                    mode = VGA_MODE_GRAPHICS_1BPP
                Else
                    bpp = 4&
                    pixelsperbyte = 8&
                    mode = VGA_MODE_GRAPHICS_4BPP
                End If
            Case 1&
                bpp = 2&
                pixelsperbyte = 4&
                mode = VGA_MODE_GRAPHICS_2BPP
            Case Else
                bpp = 8&
                pixelsperbyte = 1&
                mode = VGA_MODE_GRAPHICS_8BPP
        End Select

        If xscanpixels <= 0& Then xscanpixels = 1&
        xstride = (vga_w \ xscanpixels) \ pixelsperbyte
    Else
        mode = VGA_MODE_TEXT
        If vga_dbl <> 0& Then
            hchars = 40&
            divx = vga_dots * 2&
        Else
            hchars = 80&
            divx = vga_dots
        End If
        If (vga_crtcd(&HA&) And &H20&) <> 0& Then
            cursorenable = 0&
        Else
            cursorenable = 1&
        End If
        blinkenable = 0&
        fontbase = vga_fontbase(vga_seqd(&H3&) And 7&)
        If (vga_attrd(&H10&) And &H4&) <> 0& Then
            dup9 = 0&
        Else
            dup9 = 1&
        End If
        vga_scandbl = 0&
    End If

    startaddr = (CLng(vga_crtcd(&HC&)) * &H100&) Or CLng(vga_crtcd(&HD&))
    cursorloc = (CLng(vga_crtcd(&HE&)) * &H100&) Or CLng(vga_crtcd(&HF&))

    Select Case mode
        Case VGA_MODE_TEXT
            If diag_vga_verbose <> 0& Then vga_diagLastMode = VGA_MODE_TEXT
            dup9 = 1&
            cursorloc = (cursorloc - startaddr) And &HFFFF&
            cursor_x = cursorloc Mod hchars
            cursor_y = cursorloc \ hchars
            maxscan = (vga_crtcd(&H9&) And &H1F&) + 1&
            If maxscan <= 0& Then maxscan = 1&
            cursorBegin = (vga_crtcd(VGA_REG_DATA_CURSOR_BEGIN) And 31&)
            cursorEnd = (vga_crtcd(VGA_REG_DATA_CURSOR_END) And 31&)
            palettePage = CLng(vga_attrd(&H14&)) * 16&
            hiPage = (CLng(vga_attrd(&H14&)) And 3&) * 16&
            If (vga_attrd(&H10&) And &H80&) <> 0& Then
                paletteHiMode = 1&
            Else
                paletteHiMode = 0&
            End If

            For x = 0& To 15&
                paletteIdx = vga_attrd(x) Or palettePage
                If paletteHiMode <> 0& Then
                    paletteIdx = (paletteIdx And &HCF&) Or hiPage
                End If
                paletteIdx = paletteIdx And &HFF&
                vga_textColorCache(x) = CLng(vga_palette(paletteIdx, 2&)) Or (CLng(vga_palette(paletteIdx, 1&)) * &H100&) Or (CLng(vga_palette(paletteIdx, 0&)) * &H10000)
            Next x

            For scy = start_y To end_y
                rowOffset = scy * VGA_FB_MAX
                y = scy \ maxscan
                glyphRow = scy Mod maxscan
                cursorScan = glyphRow

                For scx = start_x To end_x
                    x = scx \ divx
                    addr = startaddr + (y * hchars) + x
                    cc = vga_RAM(0&, addr And &HFFFF&)
                    attrRaw = vga_RAM(1&, addr And &HFFFF&)
                    blink = attrRaw \ &H80&
                    attrMasked = attrRaw
                    If blinkenable <> 0& Then attrMasked = attrMasked And &H7F&

                    fontdata = vga_RAM(2&, (fontbase + (cc * 32&) + glyphRow) And &HFFFF&)

                    If vga_dbl <> 0& Then
                        charcolumn = (scx \ 2&) Mod vga_dots
                    Else
                        charcolumn = scx Mod vga_dots
                    End If

                    If (attrRaw And &H80&) <> 0& Then
                        If vga_cursor_blink_state = 0& Then
                            fontdata = 0&
                        End If
                    End If

                    If (y = cursor_y) And (x = cursor_x) And _
                       (cursorScan >= cursorBegin) And _
                       (cursorScan <= cursorEnd) And _
                       (vga_cursor_blink_state <> 0&) And (cursorenable <> 0&) Then
                        color32 = vga_textColorCache(attrMasked And &HF&)
                    Else
                        If (blinkenable <> 0&) And (blink <> 0&) And (vga_cursor_blink_state = 0&) Then
                            fontdata = 0&
                        End If

                        If vga_dots = 9& Then
                            If charcolumn = 0& Then
                                If (dup9 <> 0&) And (cc >= &HC0&) And (cc <= &HDF&) Then
                                    charcolumn = 1&
                                Else
                                    charcolumn = -1&
                                End If
                            Else
                                charcolumn = charcolumn - 1&
                            End If
                        End If

                        If (charcolumn >= 0&) And (charcolumn <= 7&) And ((fontdata And vga_glyphBitMask(charcolumn)) <> 0&) Then
                            color32 = vga_textColorCache(attrMasked And &HF&)
                        Else
                            color32 = vga_textColorCache((attrMasked \ 16&) And &H7&)
                        End If
                    End If

                    vga_framebuffer(rowOffset + scx) = color32
                Next scx
            Next scy

        Case VGA_MODE_GRAPHICS_8BPP
            If diag_vga_verbose <> 0& Then vga_diagLastMode = VGA_MODE_GRAPHICS_8BPP
            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels

                For scx = start_x To end_x Step xscanpixels
                    x = scx \ xscanpixels

                    If (vga_chain4 <> 0&) Or ((vga_modeY <> 0&) And (vga_chain4 = 0&)) Then
                        addr = (y * xstride) + x
                        plane = addr And 3&
                        addr = (addr \ 4&) + startaddr
                        cc = vga_RAM(plane, addr And &HFFFF&)
                    Else
                        addr = ((y * CLng(vga_crtcd(&H13&)) * 2&) + (x \ 4&)) And &HFFFF&
                        plane = x And 3&
                        addr = addr + startaddr
                        cc = vga_RAM(plane, addr And &HFFFF&)
                    End If

                    paletteIdx = cc And &HFF&
                    color32 = CLng(vga_palette(paletteIdx, 2&)) Or (CLng(vga_palette(paletteIdx, 1&)) * &H100&) Or (CLng(vga_palette(paletteIdx, 0&)) * &H10000)
                    For yadd = 0& To yscanpixels - 1&
                        For xadd = 0& To xscanpixels - 1&
                            vga_framebuffer(((scy + yadd) * VGA_FB_MAX) + scx + xadd) = color32
                        Next xadd
                    Next yadd
                Next scx
            Next scy

        Case VGA_MODE_GRAPHICS_4BPP
            If diag_vga_verbose <> 0& Then vga_diagLastMode = VGA_MODE_GRAPHICS_4BPP
            palettePage = CLng(vga_attrd(&H14&)) * 16&
            hiPage = (CLng(vga_attrd(&H14&)) And 3&) * 16&
            If (vga_attrd(&H10&) And &H80&) <> 0& Then
                paletteHiMode = 1&
            Else
                paletteHiMode = 0&
            End If
            For cc = 0& To 15&
                paletteIdx = vga_attrd(cc) Or palettePage
                If paletteHiMode <> 0& Then
                    paletteIdx = (paletteIdx And &HCF&) Or hiPage
                End If
                paletteIdx = paletteIdx And &HFF&
                vga_attrColorCache(cc) = CLng(vga_palette(paletteIdx, 2&)) Or (CLng(vga_palette(paletteIdx, 1&)) * &H100&) Or (CLng(vga_palette(paletteIdx, 0&)) * &H10000)
            Next cc
            byteStartX = (start_x \ xscanpixels) \ 8&
            byteEndX = (end_x \ xscanpixels) \ 8&

            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels
                For byteX = byteStartX To byteEndX
                    addr = ((y * xstride) + byteX) And &HFFFF&
                    addr = (addr + startaddr) And &HFFFF&
                    plane0Byte = vga_RAM(0&, addr)
                    plane1Byte = vga_RAM(1&, addr)
                    plane2Byte = vga_RAM(2&, addr)
                    plane3Byte = vga_RAM(3&, addr)

                    For pixelBit = 0& To 7&
                        srcPixelX = (byteX * 8&) + pixelBit
                        destX = srcPixelX * xscanpixels
                        If destX > end_x Then Exit For
                        If (destX + xscanpixels - 1&) >= start_x Then
                            bitMaskVal = vga_glyphBitMask(pixelBit)
                            cc = 0&
                            If (plane0Byte And bitMaskVal) <> 0& Then cc = cc Or 1&
                            If (plane1Byte And bitMaskVal) <> 0& Then cc = cc Or 2&
                            If (plane2Byte And bitMaskVal) <> 0& Then cc = cc Or 4&
                            If (plane3Byte And bitMaskVal) <> 0& Then cc = cc Or 8&
                            color32 = vga_attrColorCache(cc And &HF&)

                            For yadd = 0& To yscanpixels - 1&
                                If (scy + yadd) > end_y Then Exit For
                                rowBase = ((scy + yadd) * VGA_FB_MAX) + destX
                                For xadd = 0& To xscanpixels - 1&
                                    If (destX + xadd) >= start_x Then
                                        If (destX + xadd) <= end_x Then
                                            vga_framebuffer(rowBase + xadd) = color32
                                        End If
                                    End If
                                Next xadd
                            Next yadd
                        End If
                    Next pixelBit
                Next byteX
            Next scy

        Case VGA_MODE_GRAPHICS_2BPP
            If diag_vga_verbose <> 0& Then vga_diagLastMode = VGA_MODE_GRAPHICS_2BPP
            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels
                isodd = y And 1&
                y = y \ 2&

                For scx = start_x To end_x Step xscanpixels
                    x = scx \ xscanpixels
                    addr = ((8192& * isodd) + (y * xstride) + (x \ pixelsperbyte)) And &HFFFF&
                    addr = addr + startaddr
                    shift = (3& - (x And 3&)) * 2&
                    cc = (vga_RAM(addr And 1&, (addr \ 2&) And &HFFFF&) \ (2& ^ shift)) And 3&

                    color32 = vga_attrd(cc And &HF&) Or (CLng(vga_attrd(&H14&)) * 16&)
                    If (vga_attrd(&H10&) And &H80&) <> 0& Then
                        color32 = (color32 And &HCF&) Or ((vga_attrd(&H14&) And 3&) * 16&)
                    End If
                    paletteIdx = color32 And &HFF&
                    color32 = CLng(vga_palette(paletteIdx, 2&)) Or (CLng(vga_palette(paletteIdx, 1&)) * &H100&) Or (CLng(vga_palette(paletteIdx, 0&)) * &H10000)

                    For yadd = 0& To yscanpixels - 1&
                        For xadd = 0& To xscanpixels - 1&
                            vga_framebuffer(((scy + yadd) * VGA_FB_MAX) + scx + xadd) = color32
                        Next xadd
                    Next yadd
                Next scx
            Next scy

        Case VGA_MODE_GRAPHICS_1BPP
            If diag_vga_verbose <> 0& Then vga_diagLastMode = VGA_MODE_GRAPHICS_1BPP
            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels
                isodd = y And 1&
                y = y \ 2&

                For scx = start_x To end_x Step xscanpixels
                    x = scx \ xscanpixels
                    addr = ((8192& * isodd) + (y * xstride) + (x \ pixelsperbyte)) And &HFFFF&
                    addr = addr + startaddr
                    shift = 7& - (x And 7&)
                    cc = (vga_RAM(0&, addr And &HFFFF&) \ (2& ^ shift)) And 1&

                    If cc <> 0& Then
                        color32 = &HFFFFFF
                    Else
                        color32 = 0&
                    End If

                    For yadd = 0& To yscanpixels - 1&
                        For xadd = 0& To xscanpixels - 1&
                            vga_framebuffer(((scy + yadd) * VGA_FB_MAX) + (scx + xadd)) = color32
                        Next xadd
                    Next yadd
                Next scx
            Next scy
    End Select
End Sub

Public Sub vga_sendBlit()
    console_blit VarPtr(vga_framebuffer(0&)), vga_renderW, vga_renderH, (VGA_FB_MAX * 4&)
End Sub

Public Sub vga_renderThread(ByVal dummy As Long)
    vga_renderW = vga_w
    vga_renderH = vga_h

    If vga_renderW < 1& Then vga_renderW = 1&
    If vga_renderH < 1& Then vga_renderH = 1&
    If vga_renderW > VGA_FB_MAX Then vga_renderW = VGA_FB_MAX
    If vga_renderH > VGA_FB_MAX Then vga_renderH = VGA_FB_MAX

    vga_update 0&, 0&, vga_renderW - 1&, vga_renderH - 1&
    vga_doBlitNow = True
End Sub

Private Sub vga_startRenderThread()
    If vga_threadStarted <> 0& Then Exit Sub

    vga_threadHandle = CreateThread(ByVal 0&, 0&, AddressOf vga_renderThreadProc, ByVal 0&, 0&, vga_threadId)
    If vga_threadHandle = 0& Then
        vga_threadStarted = 0&
        debug_log DEBUG_INFO, "[VGA] Render thread start failed, using main thread rendering"
        Exit Sub
    End If

    vga_threadStarted = 1&
    CloseHandle vga_threadHandle
    vga_threadHandle = 0&
End Sub

Public Function vga_renderThreadProc(ByVal lpParam As Long) As Long
    Do While running <> 0&
        If vga_doRender <> 0& Then
            vga_doRender = 0&
            vga_renderThread 0&
        End If

        If vga_doBlit <> 0& Then
            vga_doBlit = 0&
            vga_doBlitNow = True
        End If
    Loop

    vga_renderThreadProc = 0&
End Function

Private Sub vga_calcmemorymap()
    Select Case (vga_gfxd(&H6&) And &HC&)
        Case &H0&
            vga_membase = 0&
            vga_memmask = &HFFFF&
        Case &H4&
            vga_membase = 0&
            vga_memmask = &HFFFF&
        Case &H8&
            vga_membase = &H10000
            vga_memmask = &H7FFF&
        Case &HC&
            vga_membase = &H18000
            vga_memmask = &H7FFF&
    End Select
End Sub

Private Sub vga_calcscreensize()
    Dim h As Long

    vga_w = (1& + CLng(vga_crtcd(&H1&)) - ((CLng(vga_crtcd(&H5&) And &H60&)) \ 32&)) * vga_dots

    h = 1& + CLng(vga_crtcd(&H12&))
    If (vga_crtcd(&H7&) And 2&) <> 0& Then h = h Or &H100&
    If (vga_crtcd(&H7&) And &H40&) <> 0& Then h = h Or &H200&
    vga_h = h

    If ((vga_shiftmode And &H20&) = 0&) And ((vga_seqd(&H1&) And &H8&) <> 0&) Then
        vga_w = vga_w * 2&
    End If

    If vga_w < 1& Then vga_w = 1&
    If vga_h < 1& Then vga_h = 1&
    If vga_w > VGA_FB_MAX Then vga_w = VGA_FB_MAX
    If vga_h > VGA_FB_MAX Then vga_h = VGA_FB_MAX

    vga_updateScanlineTiming
End Sub

Private Function vga_readcrtci() As Byte
    vga_readcrtci = vga_crtci
End Function

Private Function vga_readcrtcd() As Byte
    If vga_crtci < &H19& Then
        vga_readcrtcd = vga_crtcd(vga_crtci)
    Else
        vga_readcrtcd = &HFF&
    End If
End Function

Private Sub vga_writecrtci(ByVal value As Byte)
    vga_crtci = value And &H1F&
End Sub

Private Sub vga_writecrtcd(ByVal value As Byte)
    If vga_crtci > &H18& Then Exit Sub

    vga_crtcd(vga_crtci) = value

    Select Case vga_crtci
        Case &H1&, &H12&, &H7&
            vga_calcscreensize
    End Select
End Sub

Public Sub vga_writeport(ByVal dummy As Long, ByVal port As Integer, ByVal value As Byte)
    diag_count_vga_port
    Dim p As Long
    Dim idx As Long

    p = port And &HFFFF&

    Select Case p
        Case &H3B4&
            If (vga_misc And 1&) = 0& Then vga_writecrtci value

        Case &H3B5&
            If (vga_misc And 1&) = 0& Then vga_writecrtcd value

        Case &H3C0&, &H3C1&
            If vga_attrflipflop = 0& Then
                vga_attri = value And &H1F&
                vga_attrpal = value And &H20&
            Else
                If vga_attri < &H15& Then vga_attrd(vga_attri) = value
            End If
            vga_attrflipflop = vga_attrflipflop Xor 1&

        Case &H3C7&
            vga_DAC.state = VGA_DAC_MODE_READ
            vga_DAC.index = value
            vga_DAC.step = 0&

        Case &H3C8&
            vga_DAC.state = VGA_DAC_MODE_WRITE
            vga_DAC.index = value
            vga_DAC.step = 0&
            If diag_vga_verbose <> 0& Then vga_diagPort3C8Writes = vga_diagPort3C8Writes + 1&

        Case &H3C9&
            If diag_vga_verbose <> 0& Then vga_diagPort3C9Writes = vga_diagPort3C9Writes + 1&
            idx = vga_DAC.index
            If idx >= 0& And idx <= 255& Then
                vga_DAC.pal(idx, vga_DAC.step) = value And &H3F&
            End If
            vga_DAC.step = vga_DAC.step + 1&
            If vga_DAC.step = 3& Then
                vga_palette(vga_DAC.index, 0&) = (vga_DAC.pal(vga_DAC.index, 0&) And &H3F&) * 4&
                vga_palette(vga_DAC.index, 1&) = (vga_DAC.pal(vga_DAC.index, 1&) And &H3F&) * 4&
                vga_palette(vga_DAC.index, 2&) = (vga_DAC.pal(vga_DAC.index, 2&) And &H3F&) * 4&
                vga_DAC.step = 0&
                vga_DAC.index = CByte((CLng(vga_DAC.index) + 1&) And &HFF&)
            End If

        Case &H3C2&
            vga_misc = value

        Case &H3C4&
            vga_seqi = value And &H1F&

        Case &H3C5&
            If vga_seqi < &H5& Then
                vga_seqd(vga_seqi) = value
                Select Case vga_seqi
                    Case &H1&
                        If (value And &H1&) <> 0& Then
                            vga_dots = 8&
                        Else
                            vga_dots = 9&
                        End If
                        If (value And &H8&) <> 0& Then
                            vga_dbl = 1&
                        Else
                            vga_dbl = 0&
                        End If
                        vga_calcscreensize
                    Case &H2&
                        vga_enableplane = value And &HF&
                    Case &H4&
                        If (vga_seqd(4&) And &H8&) <> 0& Then
                            vga_chain4 = 1&
                        Else
                            vga_chain4 = 0&
                        End If
                End Select
            End If

        Case &H3CE&
            vga_gfxi = value And &H1F&

        Case &H3CF&
            If vga_gfxi < &H9& Then
                vga_gfxd(vga_gfxi) = value
                Select Case vga_gfxi
                    Case &H3&
                        vga_rotate = value And 7&
                        vga_logicop = (value \ 8&) And 3&
                    Case &H4&
                        vga_readmap = value And 3&
                    Case &H5&
                        vga_wmode = value And 3&
                        vga_rmode = (value \ 8&) And 1&
                        vga_shiftmode = (value \ 32&) And 3&
                    Case &H6&
                        vga_calcmemorymap
                End Select
            End If

        Case &H3D4&
            If (vga_misc And 1&) = 1& Then vga_writecrtci value

        Case &H3D5&
            If (vga_misc And 1&) = 1& Then vga_writecrtcd value
    End Select

    If ((vga_seqd(4&) And &HC&) = 4&) And ((vga_gfxd(5&) And &HB&) = 0&) And ((vga_gfxd(6&) And &H2&) = 0&) And ((vga_crtcd(20&) And &H40&) = 0&) And ((vga_crtcd(23&) And &H40&) <> 0&) Then
        vga_modeY = 1&
    Else
        vga_modeY = 0&
    End If
End Sub

Public Function vga_readport(ByVal dummy As Long, ByVal port As Integer) As Byte
    Dim p As Long
    Dim ret As Byte

    p = port And &HFFFF&
    ret = &HFF&

    Select Case p
        Case &H3B4&
            If (vga_misc And 1&) = 0& Then vga_readport = vga_readcrtci(): Exit Function

        Case &H3B5&
            If (vga_misc And 1&) = 0& Then vga_readport = vga_readcrtcd(): Exit Function

        Case &H3C0&
            If vga_attrflipflop = 0& Then
                ret = vga_attri Or vga_attrpal
            Else
                If vga_attri < &H15& Then ret = vga_attrd(vga_attri)
            End If

        Case &H3C1&
            If vga_attri < &H15& Then vga_readport = vga_attrd(vga_attri): Exit Function

        Case &H3C4&
            vga_readport = vga_seqi
            Exit Function

        Case &H3C5&
            If vga_seqi < &H5& Then vga_readport = vga_seqd(vga_seqi): Exit Function

        Case &H3C7&
            vga_readport = vga_DAC.state
            Exit Function

        Case &H3C8&
            vga_readport = vga_DAC.index
            Exit Function

        Case &H3C9&
            ret = vga_DAC.pal(vga_DAC.index, vga_DAC.step)
            vga_DAC.step = vga_DAC.step + 1&
            If vga_DAC.step = 3& Then
                vga_DAC.step = 0&
                vga_DAC.index = CByte((CLng(vga_DAC.index) + 1&) And &HFF&)
            End If

        Case &H3CC&
            vga_readport = vga_misc
            Exit Function

        Case &H3CE&
            vga_readport = vga_gfxi
            Exit Function

        Case &H3CF&
            If vga_gfxi < &H9& Then vga_readport = vga_gfxd(vga_gfxi): Exit Function

        Case &H3D4&
            If (vga_misc And 1&) = 1& Then vga_readport = vga_readcrtci(): Exit Function

        Case &H3D5&
            If (vga_misc And 1&) = 1& Then vga_readport = vga_readcrtcd(): Exit Function

        Case &H3DA&
            If diag_vga_verbose <> 0& Then
                vga_diag3DAReads = vga_diag3DAReads + 1&
                If (vga_status1 And &H8&) <> 0& Then vga_diag3DAVBlankReads = vga_diag3DAVBlankReads + 1&
                If (vga_status1 And &H1&) <> 0& Then vga_diag3DAHBlankReads = vga_diag3DAHBlankReads + 1&
            End If
            vga_attrflipflop = 0&
            vga_readport = vga_status1
            Exit Function
    End Select

    vga_readport = ret
End Function

Private Function vga_dologic(ByVal value As Byte, ByVal latch As Byte) As Byte
    Select Case vga_logicop
        Case 0&
            vga_dologic = value
        Case 1&
            vga_dologic = value And latch
        Case 2&
            vga_dologic = value Or latch
        Case Else
            vga_dologic = value Xor latch
    End Select
End Function

Private Function vga_host_chain4_enabled() As Long
    If (vga_seqd(&H4&) And &H8&) <> 0& Then
        vga_host_chain4_enabled = 1&
    Else
        vga_host_chain4_enabled = 0&
    End If
End Function

Private Function vga_host_odd_even_write_enabled() As Long
    If (vga_host_chain4_enabled() = 0&) And ((vga_seqd(&H4&) And &H4&) = 0&) Then
        vga_host_odd_even_write_enabled = 1&
    Else
        vga_host_odd_even_write_enabled = 0&
    End If
End Function

Private Function vga_host_odd_even_read_enabled() As Long
    If (vga_host_chain4_enabled() = 0&) And ((vga_gfxd(&H5&) And &H10&) <> 0&) And ((vga_gfxd(&H6&) And &H2&) <> 0&) Then
        vga_host_odd_even_read_enabled = 1&
    Else
        vga_host_odd_even_read_enabled = 0&
    End If
End Function

Private Sub vga_loadlatches(ByVal addr As Long)
    addr = addr And &HFFFF&
    vga_latch(0&) = vga_RAM(0&, addr)
    vga_latch(1&) = vga_RAM(1&, addr)
    vga_latch(2&) = vga_RAM(2&, addr)
    vga_latch(3&) = vga_RAM(3&, addr)
End Sub

Public Sub vga_writememory(ByVal dummy As Long, ByVal addr As Long, ByVal value As Byte)
    diag_count_vga_mem
    If diag_vga_verbose <> 0& Then
        vga_diagMemWrites = vga_diagMemWrites + 1&
        If value <> 0& Then vga_diagMemNonZeroWrites = vga_diagMemNonZeroWrites + 1&
    End If
    Dim temp As Byte
    Dim plane As Long
    Dim a As Long
    Dim bitMask As Byte
    Dim parity As Long
    Dim planeMask As Long

    If (vga_misc And &H2&) = 0& Then GoTo WriteDone

    a = addr - &HA0000
    a = (a - vga_membase) And vga_memmask

    If vga_host_chain4_enabled() <> 0& Then
        plane = a And 3&
        a = (a \ 4&) And &HFFFF&
        planeMask = CLng(2 ^ plane)
        If (vga_enableplane And planeMask) <> 0& Then
            vga_RAM(plane, a) = value
        End If
        GoTo WriteDone
    End If

    If vga_host_odd_even_write_enabled() <> 0& Then
        parity = a And 1&
        a = (a \ 2&) And &HFFFF&

        planeMask = CLng(2 ^ parity)
        If (vga_enableplane And planeMask) <> 0& Then
            vga_RAM(parity, a) = value
        End If

        planeMask = CLng(2 ^ (parity + 2&))
        If (vga_enableplane And planeMask) <> 0& Then
            vga_RAM(parity + 2&, a) = value
        End If
        GoTo WriteDone
    End If

    a = a And &HFFFF&

    Select Case vga_wmode
        Case 0&
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (vga_enableplane And planeMask) <> 0& Then
                    If (vga_gfxd(&H1&) And planeMask) <> 0& Then
                        If (vga_gfxd(&H0&) And planeMask) <> 0& Then
                            temp = &HFF&
                        Else
                            temp = 0&
                        End If
                    Else
                        temp = vga_dorotate(value)
                    End If

                    temp = vga_dologic(temp, vga_latch(plane))
                    temp = (temp And vga_gfxd(&H8&)) Or (vga_latch(plane) And vga_notbyte(vga_gfxd(&H8&)))
                    vga_RAM(plane, a) = temp
                End If
            Next plane

        Case 1&
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (vga_enableplane And planeMask) <> 0& Then
                    vga_RAM(plane, a) = vga_latch(plane)
                End If
            Next plane

        Case 2&
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (vga_enableplane And planeMask) <> 0& Then
                    If (value And planeMask) <> 0& Then
                        temp = &HFF&
                    Else
                        temp = 0&
                    End If
                    temp = vga_dologic(temp, vga_latch(plane))
                    temp = (temp And vga_gfxd(&H8&)) Or (vga_latch(plane) And vga_notbyte(vga_gfxd(&H8&)))
                    vga_RAM(plane, a) = temp
                End If
            Next plane

        Case 3&
            bitMask = vga_dorotate(value) And vga_gfxd(&H8&)
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (vga_enableplane And planeMask) <> 0& Then
                    If (vga_gfxd(&H0&) And planeMask) <> 0& Then
                        temp = &HFF&
                    Else
                        temp = 0&
                    End If
                    temp = vga_dologic(temp, vga_latch(plane))
                    temp = (temp And bitMask) Or (vga_latch(plane) And vga_notbyte(bitMask))
                    vga_RAM(plane, a) = temp
                End If
            Next plane
    End Select

WriteDone:
End Sub

Public Function vga_readmemory(ByVal dummy As Long, ByVal addr As Long) As Byte
    Dim plane As Long
    Dim retL As Long
    Dim a As Long
    Dim compare As Long
    Dim planeMask As Long

    a = addr - &HA0000
    a = (a - vga_membase) And vga_memmask

    If vga_host_chain4_enabled() <> 0& Then
        plane = a And 3&
        a = (a \ 4&) And &HFFFF&
    ElseIf vga_host_odd_even_read_enabled() <> 0& Then
        plane = (a And 1&) Or (vga_readmap And &H2&)
        a = (a \ 2&) And &HFFFF&
    Else
        plane = vga_readmap And 3&
        a = a And &HFFFF&
    End If

    vga_loadlatches a

    If vga_rmode = 0& Then
        vga_readmemory = vga_latch(plane)
    Else
        retL = &HFF&
        For plane = 0& To 3&
            planeMask = CLng(2 ^ plane)
            If (vga_gfxd(&H7&) And planeMask) <> 0& Then
                If (vga_gfxd(&H2&) And planeMask) <> 0& Then
                    compare = &HFF&
                Else
                    compare = 0&
                End If
                retL = retL And (((CLng(vga_latch(plane)) Xor compare) Xor &HFF&) And &HFF&)
            End If
        Next plane
        vga_readmemory = CByte(retL And &HFF&)
    End If
End Function

Public Sub vga_drawCallback(ByVal dummy As Long)
    If vga_threadStarted = 0& Then
        vga_renderThread 0&
        vga_doBlitNow = True
    Else
        vga_doRender = 1&
        vga_doBlit = 1&
    End If
End Sub

Public Sub vga_blinkCallback(ByVal dummy As Long)
    vga_cursor_blink_state = vga_cursor_blink_state Xor 1&
End Sub

Public Sub vga_hblankCallback(ByVal dummy As Long)
    Dim vblankStartScan As Long
    Dim vblankEndScan As Long

    timing_timerEnable vga_hblankEndTimer
    vga_status1 = vga_status1 Or 1&

    vblankStartScan = (CLng(vga_vblankstart) And &HFFFF&)
    vblankEndScan = (CLng(vga_vblankend) And &HFFFF&)

    vga_curScanline = ((vga_curScanline + 1&) And &HFFFF&)

    If vga_curScanline = vblankStartScan Then
        vga_status1 = vga_status1 Or 8&
    ElseIf vga_curScanline = vblankEndScan Then
        vga_curScanline = 0&
        vga_status1 = vga_status1 And &HF7&
    End If
End Sub

Public Sub vga_hblankEndCallback(ByVal dummy As Long)
    timing_timerDisable vga_hblankEndTimer
    vga_status1 = vga_status1 And &HFE&
End Sub

Public Sub vga_dumpregs()
    ' Kept for parity; detailed register dump can be added behind a debug gate.
End Sub

Public Function vga_diagSnapshotAndReset() As String
    Dim startaddr As Long
    Dim modeVal As Long
    Dim outStr As String
    Dim vbStart As Long
    Dim vbEnd As Long

    startaddr = (CLng(vga_crtcd(&HC&)) * &H100&) Or CLng(vga_crtcd(&HD&))
    modeVal = (vga_attrd(&H10&) And 1&)
    vbStart = CLng(vga_vblankstart)
    vbEnd = CLng(vga_vblankend)

    outStr = ""
    outStr = outStr & " vmode=" & CStr(vga_diagLastMode)
    outStr = outStr & " chain4=" & CStr(vga_chain4)
    outStr = outStr & " modeY=" & CStr(vga_modeY)
    outStr = outStr & " sh=" & CStr(vga_shiftmode And &H3&)
    outStr = outStr & " wm=" & CStr(vga_wmode And &H3&)
    outStr = outStr & " rm=" & CStr(vga_rmode And &H1&)
    outStr = outStr & " em=" & right$("0" & Hex$(vga_enableplane And &HF&), 1&)
    outStr = outStr & " mm=" & right$("00" & Hex$(vga_gfxd(&H6&) And &HC&), 2&)
    outStr = outStr & " misc=" & right$("00" & Hex$(vga_misc), 2&)
    outStr = outStr & " gfx=" & CStr(modeVal)
    outStr = outStr & " sa=" & right$("0000" & Hex$(startaddr And &HFFFF&), 4&)
    outStr = outStr & " off=" & right$("00" & Hex$(vga_crtcd(&H13&)), 2&)
    outStr = outStr & " dac=" & CStr(vga_diagPort3C8Writes) & "/" & CStr(vga_diagPort3C9Writes)
    outStr = outStr & " vram=" & CStr(vga_diagMemWrites) & ":" & CStr(vga_diagMemNonZeroWrites)
    outStr = outStr & " rd3da=" & CStr(vga_diag3DAReads) & ":" & CStr(vga_diag3DAVBlankReads) & ":" & CStr(vga_diag3DAHBlankReads)
    outStr = outStr & " s1=" & right$("00" & Hex$(vga_status1), 2&) & " sl=" & CStr(vga_curScanline) & " vb=" & CStr(vbStart) & "-" & CStr(vbEnd)

    vga_diagPort3C8Writes = 0&
    vga_diagPort3C9Writes = 0&
    vga_diagMemWrites = 0&
    vga_diagMemNonZeroWrites = 0&
    vga_diag3DAReads = 0&
    vga_diag3DAVBlankReads = 0&
    vga_diag3DAHBlankReads = 0&

    vga_diagSnapshotAndReset = outStr
End Function

Private Function vga_color(ByVal c As Long) As Long
    Dim idx As Long

    idx = c And &HFF&
    vga_color = CLng(vga_palette(idx, 2&)) Or (CLng(vga_palette(idx, 1&)) * &H100&) Or (CLng(vga_palette(idx, 0&)) * &H10000)
End Function

Private Function vga_dorotate(ByVal v As Byte) As Byte
    Dim r As Long
    Dim src As Long
    Dim outv As Long

    r = vga_rotate And 7&
    src = v

    If r = 0& Then
        vga_dorotate = v
        Exit Function
    End If

    outv = ((src \ (2& ^ r)) Or ((src * (2& ^ (8& - r))) And &HFF&)) And &HFF&
    vga_dorotate = CByte(outv)
End Function

Private Function vga_notbyte(ByVal v As Byte) As Byte
    vga_notbyte = CByte(v Xor &HFF&)
End Function

Private Function vga_fontbase(ByVal idx As Long) As Long
    Select Case (idx And 7&)
        Case 0&: vga_fontbase = &H0&
        Case 1&: vga_fontbase = &H4000&
        Case 2&: vga_fontbase = &H8000&
        Case 3&: vga_fontbase = &HC000&
        Case 4&: vga_fontbase = &H2000&
        Case 5&: vga_fontbase = &H6000&
        Case 6&: vga_fontbase = &HA000&
        Case Else: vga_fontbase = &HE000&
    End Select
End Function
