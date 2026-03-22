Attribute VB_Name = "modET4000"
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
Private Const VGA_RAM_PLANE_SIZE As Long = 262144
Private Const ET4000_VRAM_SIZE As Long = 1048576
Private Const ET4000_VRAM_MASK As Long = ET4000_VRAM_SIZE - 1&
Private Const ET4000_CRTC_MAX As Long = &H3F&
Private Const ET4000_ATTR_MAX As Long = &H16&
Private Const ET4000_SEQ_MAX As Long = &HE&

Private Const VGA_DAC_MODE_READ As Byte = 0&
Private Const VGA_DAC_MODE_WRITE As Byte = 3&

Private Const VGA_REG_DATA_CURSOR_BEGIN As Byte = &HA&
Private Const VGA_REG_DATA_CURSOR_END As Byte = &HB&

Private Const VGA_MODE_TEXT As Byte = 0&
Private Const VGA_MODE_GRAPHICS_8BPP As Byte = 1&
Private Const VGA_MODE_GRAPHICS_4BPP As Byte = 2&
Private Const VGA_MODE_GRAPHICS_2BPP As Byte = 3&
Private Const VGA_MODE_GRAPHICS_1BPP As Byte = 4&
Private Const VGA_MODE_GRAPHICS_15BPP As Byte = 5&
Private Const VGA_MODE_GRAPHICS_16BPP As Byte = 6&

Private VBIOS(0& To 32767&) As Byte
Private et4000_palette(0& To 255&, 0& To 2&) As Byte
Private et4000_color15to32(0& To &HFFFF&) As Long
Private et4000_color16to32(0& To &HFFFF&) As Long

Private et4000_DAC As VGADAC_t
Private et4000_framebuffer(0& To VGA_FB_PIXELS - 1&) As Long
Private et4000_textColorCache(0& To 15&) As Long
Private et4000_glyphBitMask(0& To 7&) As Byte
Private et4000_attrColorCache(0& To 15&) As Long

Private et4000_w As Long
Private et4000_h As Long
Private et4000_dots As Long
Private et4000_membase As Long
Private et4000_memmask As Long

Private et4000_cursorloc As Long
Private et4000_dbl As Byte
Private et4000_crtci As Byte
Private et4000_crtcd(0& To ET4000_CRTC_MAX) As Byte
Private et4000_attri As Byte
Private et4000_attrd(0& To ET4000_ATTR_MAX) As Byte
Private et4000_attrflipflop As Byte
Private et4000_attrpal As Byte
Private et4000_gfxi As Byte
Private et4000_gfxd(0& To &H8&) As Byte
Private et4000_seqi As Byte
Private et4000_seqd(0& To ET4000_SEQ_MAX) As Byte
Private et4000_misc As Byte
Private et4000_status0 As Byte
Private et4000_status1 As Byte
Private et4000_cursor_blink_state As Byte
Private et4000_bankReg As Byte
Private et4000_dacMask As Byte
Private et4000_writeBank As Long
Private et4000_readBank As Long

Private et4000_wmode As Byte
Private et4000_rmode As Byte
Private et4000_shiftmode As Byte
Private et4000_rotate As Byte
Private et4000_logicop As Byte
Private et4000_enableplane As Byte
Private et4000_readmap As Byte
Private et4000_scandbl As Byte
Private et4000_hdbl As Byte
Private et4000_bpp As Byte
Private et4000_latch(0& To 3&) As Byte

Private et4000_RAM(0& To VGA_RAM_PLANES - 1&, 0& To VGA_RAM_PLANE_SIZE - 1&) As Byte

Private et4000_hblankstart As Double
Private et4000_hblankend As Double
Private et4000_hblanklen As Double
Private et4000_dispinterval As Double
Private et4000_hblankinterval As Double
Private et4000_htotal As Double
Private et4000_vblankstart As Double
Private et4000_vblankend As Double
Private et4000_vblanklen As Double
Private et4000_vblankinterval As Double
Private et4000_frameinterval As Double

Private et4000_doRender As Byte
Private et4000_doBlit As Byte
Private et4000_targetFPS As Double

Private et4000_hblankTimer As Long
Private et4000_hblankEndTimer As Long
Private et4000_drawTimer As Long
Private et4000_curScanline As Long
Private et4000_chain4 As Byte
Private et4000_modeY As Byte
Private et4000_threadStarted As Byte
Private et4000_threadHandle As Long
Private et4000_threadId As Long
Public et4000_doBlitNow As Boolean

Private et4000_renderW As Long
Private et4000_renderH As Long

Private et4000_diagPort3C8Writes As Long
Private et4000_diagPort3C9Writes As Long
Private et4000_diagMemWrites As Long
Private et4000_diagMemNonZeroWrites As Long
Private et4000_diagLastMode As Long
Private et4000_diag3DAReads As Long
Private et4000_diag3DAVBlankReads As Long
Private et4000_diag3DAHBlankReads As Long
Private et4000_colorTablesReady As Byte
Private et4000_ramdacCtrl As Byte
Private et4000_ramdacState As Byte
Private et4000_ramdacIndex As Byte
Private et4000_ramdacRegs(0& To 255&) As Byte
Private et4000_ramdacPixelMask As Long

'Some of these formerly local variables from et4000_update have to be outside of the sub, otherwise it can crash the thread
Private palettePage As Long
Private hiPage As Long
Private paletteHiMode As Long
Private byteX As Long
Private srcPixelX As Long
Private destX As Long
Private rowBase As Long
Private plane0Byte As Long
Private plane1Byte As Long
Private plane2Byte As Long
Private plane3Byte As Long
Private pixelBit As Long
Private byteStartX As Long
Private byteEndX As Long
Private bitMaskVal As Long

Public Function et4000_init() As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim j As Long
    Dim idx As Long
    Dim vRead As Variant
    Dim vWrite As Variant

    debug_log DEBUG_INFO, "[ET4000] Initializing ET4000AX video device"
    diag_mark_et4000_init

    et4000_w = 640&
    et4000_h = 400&
    et4000_dots = 8&
    et4000_targetFPS = 60#
    et4000_attrpal = &H20&
    et4000_memmask = &H1FFFF
    et4000_dacMask = &HFF&
    et4000_bpp = 8&
    et4000_bankReg = 0&
    et4000_writeBank = 0&
    et4000_readBank = 0&
    et4000_ramdacCtrl = 0&
    et4000_ramdacState = 0&
    et4000_ramdacIndex = 0&
    et4000_ramdacPixelMask = 0&
    For i = 0& To 255&
        et4000_ramdacRegs(i) = 0&
    Next i
    et4000_glyphBitMask(0&) = &H80&
    et4000_glyphBitMask(1&) = &H40&
    et4000_glyphBitMask(2&) = &H20&
    et4000_glyphBitMask(3&) = &H10&
    et4000_glyphBitMask(4&) = &H8&
    et4000_glyphBitMask(5&) = &H4&
    et4000_glyphBitMask(6&) = &H2&
    et4000_glyphBitMask(7&) = &H1&

    et4000_initColorTables

    For y = 0& To 399&
        idx = y * VGA_FB_MAX
        For x = 0& To 639&
            et4000_framebuffer(idx + x) = et4000_color(0&)
        Next x
    Next y
    console_blit VarPtr(et4000_framebuffer(0&)), 640&, 400&, (VGA_FB_MAX * 4&)

    If vga_lockFPS >= 1# Then
        et4000_targetFPS = vga_lockFPS
    End If

    et4000_doRender = 0&
    et4000_doBlit = 0&
    et4000_doBlitNow = False
    et4000_diagPort3C8Writes = 0&
    et4000_diagPort3C9Writes = 0&
    et4000_diagMemWrites = 0&
    et4000_diagMemNonZeroWrites = 0&
    et4000_diagLastMode = VGA_MODE_TEXT
    et4000_diag3DAReads = 0&
    et4000_diag3DAVBlankReads = 0&
    et4000_diag3DAHBlankReads = 0&
    et4000_curScanline = 0&
    et4000_drawTimer = timing_addTimer(TIMER_CB_ET4000_DRAW, 0&, et4000_targetFPS, TIMING_ENABLED)
    timing_addTimer TIMER_CB_ET4000_BLINK, 0&, 3.75, TIMING_ENABLED
    et4000_hblankTimer = timing_addTimer(TIMER_CB_ET4000_HBLANK, 0&, 10000#, TIMING_ENABLED)
    et4000_hblankEndTimer = timing_addTimer(TIMER_CB_ET4000_HBLANK_END, 0&, 100#, TIMING_ENABLED)

    For i = 0& To VGA_RAM_PLANES - 1&
        For j = 0& To VGA_RAM_PLANE_SIZE - 1&
            et4000_RAM(i, j) = 0&
        Next j
    Next i

    ports_cbRegister &H3B4&, 39&, PORTS_CB_ET4000, PORTS_CB_NONE, PORTS_CB_ET4000, PORTS_CB_NONE, 0&
    memory_mapCallbackRegister &HA0000, &H20000, MEMORY_CB_ET4000, MEMORY_CB_ET4000, 0&

    If utility_loadFile(VBIOS, 32768, "roms/video/et4000.bin") <> 0& Then
        et4000_init = -1&
        Exit Function
    End If

    vRead = VBIOS
    vWrite = Empty
    memory_mapRegister &HC0000, 32768, vRead, vWrite

    et4000_threadStarted = 0&
    et4000_startRenderThread

    et4000_init = 0&
End Function

Private Sub et4000_initColorTables()
    Dim c As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long

    If et4000_colorTablesReady <> 0& Then Exit Sub

    For c = 0& To &HFFFF&
        r = ((c \ &H400&) And &H1F&)
        g = ((c \ &H20&) And &H1F&)
        b = (c And &H1F&)
        et4000_color15to32(c) = et4000_expand5To8(b) Or (et4000_expand5To8(g) * &H100&) Or (et4000_expand5To8(r) * &H10000)

        r = ((c \ &H800&) And &H1F&)
        g = ((c \ &H20&) And &H3F&)
        b = (c And &H1F&)
        et4000_color16to32(c) = et4000_expand5To8(b) Or (et4000_expand6To8(g) * &H100&) Or (et4000_expand5To8(r) * &H10000)
    Next c

    et4000_colorTablesReady = 1&
End Sub

Private Function et4000_expand5To8(ByVal value As Long) As Long
    et4000_expand5To8 = ((value And &H1F&) * &HFF&) \ &H1F&
End Function

Private Function et4000_expand6To8(ByVal value As Long) As Long
    et4000_expand6To8 = ((value And &H3F&) * &HFF&) \ &H3F&
End Function

Private Sub et4000_updateRamdacBpp()
    Dim oldBpp As Long

    oldBpp = et4000_bpp

    If (et4000_ramdacCtrl And &H80&) <> 0& Then
        If (et4000_ramdacCtrl And &H40&) <> 0& Then
            et4000_bpp = 16&
        Else
            et4000_bpp = 15&
        End If
    ElseIf (et4000_ramdacCtrl And &H40&) <> 0& Then
        If (et4000_ramdacRegs(&H10&) And &H1&) <> 0& Then
            et4000_bpp = 32&
        ElseIf (et4000_ramdacCtrl And &H20&) <> 0& Then
            et4000_bpp = 24&
        Else
            et4000_bpp = 32&
        End If
    Else
        et4000_bpp = 8&
    End If

    If et4000_bpp <> oldBpp Then et4000_calcscreensize
End Sub

Private Sub et4000_writePaletteData(ByVal value As Byte)
    Dim idx As Long

    If diag_vga_verbose <> 0& Then et4000_diagPort3C9Writes = et4000_diagPort3C9Writes + 1&

    idx = et4000_DAC.index
    If idx >= 0& And idx <= 255& Then
        et4000_DAC.pal(idx, et4000_DAC.step) = value And &H3F&
    End If
    et4000_DAC.step = et4000_DAC.step + 1&
    If et4000_DAC.step = 3& Then
        et4000_palette(et4000_DAC.index, 0&) = (et4000_DAC.pal(et4000_DAC.index, 0&) And &H3F&) * 4&
        et4000_palette(et4000_DAC.index, 1&) = (et4000_DAC.pal(et4000_DAC.index, 1&) And &H3F&) * 4&
        et4000_palette(et4000_DAC.index, 2&) = (et4000_DAC.pal(et4000_DAC.index, 2&) And &H3F&) * 4&
        et4000_DAC.step = 0&
        et4000_DAC.index = CByte((CLng(et4000_DAC.index) + 1&) And &HFF&)
    End If
End Sub

Private Function et4000_readPaletteData() As Byte
    et4000_readPaletteData = et4000_DAC.pal(et4000_DAC.index, et4000_DAC.step)
    et4000_DAC.step = et4000_DAC.step + 1&
    If et4000_DAC.step = 3& Then
        et4000_DAC.step = 0&
        et4000_DAC.index = CByte((CLng(et4000_DAC.index) + 1&) And &HFF&)
    End If
End Function

Private Sub et4000_updateScanlineTiming()
    Dim pixelclock As Double
    Static lastw As Long
    Static lasth As Long
    Static lastFPS As Double
    Dim ratio As Double
    Dim vblankStartTmp As Long
    Dim vblankEndTmp As Long
    Dim clockSel As Long

    clockSel = ((CLng(et4000_misc) \ 4&) And 3&)
    clockSel = clockSel Or ((CLng(et4000_crtcd(&H34&)) * 2&) And &H4&)
    clockSel = clockSel Or ((CLng(et4000_crtcd(&H31&)) \ 8&) And &H8&)
    pixelclock = et4000_clockHz(clockSel)
    If pixelclock <= 0# Then pixelclock = 50000000#

    et4000_hblankstart = CDbl(CLng(et4000_crtcd(&H2&)) Or ((CLng(et4000_crtcd(&H3F&)) And &H4&) * 64&)) * CDbl(et4000_dots)
    et4000_hblankend = (CDbl(et4000_crtcd(&H2&)) * CDbl(et4000_dots)) + (CDbl(et4000_crtcd(&H3&) And &H1F&) + 1#) * CDbl(et4000_dots)
    et4000_hblanklen = et4000_hblankend - et4000_hblankstart

    vblankStartTmp = CLng(et4000_crtcd(&H10&)) Or CLng((et4000_crtcd(&H7&) And &H4&) * 64&) Or CLng((et4000_crtcd(&H7&) And &H80&) * 4&)
    vblankEndTmp = CLng(et4000_crtcd(&H6&)) Or CLng((et4000_crtcd(&H7&) And &H1&) * 256&) Or CLng((et4000_crtcd(&H7&) And &H20&) * 16&)
    If (et4000_crtcd(&H35&) And &H1&) <> 0& Then vblankStartTmp = vblankStartTmp Or &H400&
    If (et4000_crtcd(&H35&) And &H2&) <> 0& Then vblankEndTmp = vblankEndTmp Or &H400&
    et4000_vblankstart = CDbl(vblankStartTmp)
    et4000_vblankend = CDbl(vblankEndTmp)
    et4000_vblanklen = et4000_vblankend - et4000_vblankstart
    et4000_htotal = CDbl(CLng(et4000_crtcd(&H0&)) Or ((CLng(et4000_crtcd(&H3F&)) And &H1&) * 256&))

    If (et4000_vblankend > 0#) And ((et4000_htotal + 5#) > 0#) And (et4000_dots > 0&) Then
        et4000_targetFPS = pixelclock / ((et4000_htotal + 5#) * CDbl(et4000_dots) * et4000_vblankend)
    Else
        et4000_targetFPS = 60#
    End If

    ratio = timing_getFreq() / pixelclock
    et4000_dispinterval = (et4000_htotal + 5#) * CDbl(et4000_dots) * ratio
    et4000_hblankinterval = et4000_hblanklen * ratio
    et4000_vblankinterval = et4000_hblankend * et4000_vblanklen * ratio
    et4000_frameinterval = et4000_hblankend * et4000_vblankend * ratio

    If (lastw <> et4000_w) Or (lasth <> et4000_h) Or (lastFPS <> et4000_targetFPS) Then
        debug_log DEBUG_DETAIL, "[ET4000] Mode switch: " & CStr(et4000_w) & "x" & CStr(et4000_h) & " (" & Format$(et4000_targetFPS, "0.00") & " Hz)"
        lastw = et4000_w
        lasth = et4000_h
        lastFPS = et4000_targetFPS
    End If

    timing_updateInterval et4000_hblankTimer, et4000_dispinterval
    timing_updateInterval et4000_hblankEndTimer, et4000_hblankinterval
    timing_timerEnable et4000_hblankTimer
    timing_timerDisable et4000_hblankEndTimer

    If vga_lockFPS = 0# Then
        If et4000_targetFPS > 0# Then
            timing_updateIntervalFreq et4000_drawTimer, et4000_targetFPS
        End If
    End If
End Sub

Private Sub et4000_update(ByVal start_x As Long, ByVal start_y As Long, ByVal end_x As Long, ByVal end_y As Long)
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
    Dim lowresMode As Long
    Dim rowStrideBytes As Long
    Dim baseLinear As Long
    Dim srcY As Long
    Dim pixelValue As Long
    Dim packed8bpp As Long
    Dim packedStartScale As Long

    If start_x < 0& Then start_x = 0&
    If start_y < 0& Then start_y = 0&
    If end_x >= et4000_w Then end_x = et4000_w - 1&
    If end_y >= et4000_h Then end_y = et4000_h - 1&
    If end_x > (VGA_FB_MAX - 1&) Then end_x = VGA_FB_MAX - 1&
    If end_y > (VGA_FB_MAX - 1&) Then end_y = VGA_FB_MAX - 1&
    If end_x < start_x Or end_y < start_y Then Exit Sub

    If (et4000_attrd(&H10&) And 1&) <> 0& Then
        lowresMode = et4000_lowresEnabled()
        packed8bpp = 0&
        packedStartScale = 1&

        If (et4000_shiftmode And &H2&) <> 0& Then
            xscanpixels = 2&
            yscanpixels = (et4000_crtcd(&H9&) And &H1F&) + 1&
        Else
            If lowresMode <> 0& Then
                xscanpixels = 2&
            ElseIf (et4000_seqd(&H1&) And &H8&) <> 0& Then
                xscanpixels = 2&
            Else
                xscanpixels = 1&
            End If
            If (et4000_crtcd(&H9&) And &H80&) <> 0& Then
                yscanpixels = 2&
            Else
                yscanpixels = 1&
            End If
        End If

        If (et4000_gfxd(&H5&) And &H60&) >= &H40& Then
            rowStrideBytes = et4000_rowOffsetBytes()
            If lowresMode <> 0& Then
                xscanpixels = 2&
            Else
                xscanpixels = 1&
            End If

            Select Case et4000_bpp
                Case 15&
                    bpp = 15&
                    mode = VGA_MODE_GRAPHICS_15BPP
                Case 16&
                    bpp = 16&
                    mode = VGA_MODE_GRAPHICS_16BPP
                Case Else
                    bpp = 8&
                    mode = VGA_MODE_GRAPHICS_8BPP
                    packed8bpp = 1&
                    If (lowresMode <> 0&) And ((et4000_seqd(&HE&) And &H2&) <> 0&) Then
                        xscanpixels = 1&
                        rowStrideBytes = rowStrideBytes * 2&
                        packedStartScale = 2&
                    End If
            End Select
        Else
            Select Case (et4000_shiftmode And &H3&)
                Case 0&
                    If (et4000_attrd(&H12&) And &HF&) = 1& Then
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
            xstride = (et4000_w \ xscanpixels) \ pixelsperbyte
        End If
    Else
        mode = VGA_MODE_TEXT
        If et4000_dbl <> 0& Then
            hchars = 40&
            divx = et4000_dots * 2&
        Else
            hchars = 80&
            divx = et4000_dots
        End If
        If (et4000_crtcd(&HA&) And &H20&) <> 0& Then
            cursorenable = 0&
        Else
            cursorenable = 1&
        End If
        blinkenable = 0&
        fontbase = et4000_fontbase(et4000_seqd(&H3&) And 7&)
        If (et4000_attrd(&H10&) And &H4&) <> 0& Then
            dup9 = 0&
        Else
            dup9 = 1&
        End If
        et4000_scandbl = 0&
    End If

    startaddr = et4000_startAddress()
    cursorloc = et4000_cursorAddress()

    Select Case mode
        Case VGA_MODE_TEXT
            If diag_vga_verbose <> 0& Then et4000_diagLastMode = VGA_MODE_TEXT
            dup9 = 1&
            cursorloc = (cursorloc - startaddr) And (VGA_RAM_PLANE_SIZE - 1&)
            cursor_x = cursorloc Mod hchars
            cursor_y = cursorloc \ hchars
            maxscan = (et4000_crtcd(&H9&) And &H1F&) + 1&
            If maxscan <= 0& Then maxscan = 1&
            cursorBegin = (et4000_crtcd(VGA_REG_DATA_CURSOR_BEGIN) And 31&)
            cursorEnd = (et4000_crtcd(VGA_REG_DATA_CURSOR_END) And 31&)
            palettePage = CLng(et4000_attrd(&H14&)) * 16&
            hiPage = (CLng(et4000_attrd(&H14&)) And 3&) * 16&
            If (et4000_attrd(&H10&) And &H80&) <> 0& Then
                paletteHiMode = 1&
            Else
                paletteHiMode = 0&
            End If

            For x = 0& To 15&
                paletteIdx = et4000_attrd(x) Or palettePage
                If paletteHiMode <> 0& Then
                    paletteIdx = (paletteIdx And &HCF&) Or hiPage
                End If
                paletteIdx = paletteIdx And &HFF&
                et4000_textColorCache(x) = CLng(et4000_palette(paletteIdx, 2&)) Or (CLng(et4000_palette(paletteIdx, 1&)) * &H100&) Or (CLng(et4000_palette(paletteIdx, 0&)) * &H10000)
            Next x

            For scy = start_y To end_y
                rowOffset = scy * VGA_FB_MAX
                y = scy \ maxscan
                glyphRow = scy Mod maxscan
                cursorScan = glyphRow

                For scx = start_x To end_x
                    x = scx \ divx
                    addr = startaddr + (y * hchars) + x
                    cc = et4000_readPlaneByte(addr, 0&)
                    attrRaw = et4000_readPlaneByte(addr, 1&)
                    blink = attrRaw \ &H80&
                    attrMasked = attrRaw
                    If blinkenable <> 0& Then attrMasked = attrMasked And &H7F&

                    fontdata = et4000_readPlaneByte(fontbase + (cc * 32&) + glyphRow, 2&)

                    If et4000_dbl <> 0& Then
                        charcolumn = (scx \ 2&) Mod et4000_dots
                    Else
                        charcolumn = scx Mod et4000_dots
                    End If

                    If (attrRaw And &H80&) <> 0& Then
                        If et4000_cursor_blink_state = 0& Then
                            fontdata = 0&
                        End If
                    End If

                    If (y = cursor_y) And (x = cursor_x) And _
                       (cursorScan >= cursorBegin) And _
                       (cursorScan <= cursorEnd) And _
                       (et4000_cursor_blink_state <> 0&) And (cursorenable <> 0&) Then
                        color32 = et4000_textColorCache(attrMasked And &HF&)
                    Else
                        If (blinkenable <> 0&) And (blink <> 0&) And (et4000_cursor_blink_state = 0&) Then
                            fontdata = 0&
                        End If

                        If et4000_dots = 9& Then
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

                        If (charcolumn >= 0&) And (charcolumn <= 7&) And ((fontdata And et4000_glyphBitMask(charcolumn)) <> 0&) Then
                            color32 = et4000_textColorCache(attrMasked And &HF&)
                        Else
                            color32 = et4000_textColorCache((attrMasked \ 16&) And &H7&)
                        End If
                    End If

                    et4000_framebuffer(rowOffset + scx) = color32
                Next scx
            Next scy

        Case VGA_MODE_GRAPHICS_8BPP
            If diag_vga_verbose <> 0& Then et4000_diagLastMode = VGA_MODE_GRAPHICS_8BPP
            If packed8bpp <> 0& Then
                baseLinear = startaddr * 4& * packedStartScale
                For scy = start_y To end_y Step yscanpixels
                    y = scy \ yscanpixels
                    rowOffset = baseLinear + (y * rowStrideBytes)

                    For scx = start_x To end_x Step xscanpixels
                        x = scx \ xscanpixels
                        cc = et4000_readLinear(et4000_translateAddress(rowOffset + x))
                        If (et4000_attrd(&H10&) And &H80&) <> 0& Then
                            cc = (cc And &HF&) Or ((CLng(et4000_attrd(&H14&)) And &HF&) * &H10&)
                        End If

                        paletteIdx = cc And et4000_dacMask And &HFF&
                        color32 = CLng(et4000_palette(paletteIdx, 2&)) Or (CLng(et4000_palette(paletteIdx, 1&)) * &H100&) Or (CLng(et4000_palette(paletteIdx, 0&)) * &H10000)
                        For yadd = 0& To yscanpixels - 1&
                            For xadd = 0& To xscanpixels - 1&
                                et4000_framebuffer(((scy + yadd) * VGA_FB_MAX) + scx + xadd) = color32
                            Next xadd
                        Next yadd
                    Next scx
                Next scy
            Else
                For scy = start_y To end_y Step yscanpixels
                    y = scy \ yscanpixels

                    For scx = start_x To end_x Step xscanpixels
                        x = scx \ xscanpixels

                        If (et4000_chain4 <> 0&) Or ((et4000_modeY <> 0&) And (et4000_chain4 = 0&)) Then
                            addr = (y * xstride) + x
                            plane = addr And 3&
                            addr = (addr \ 4&) + startaddr
                            cc = et4000_readPlaneByte(addr, plane)
                        Else
                            addr = (y * CLng(et4000_crtcd(&H13&)) * 2&) + (x \ 4&)
                            plane = x And 3&
                            addr = addr + startaddr
                            cc = et4000_readPlaneByte(addr, plane)
                        End If

                        paletteIdx = cc And et4000_dacMask And &HFF&
                        color32 = CLng(et4000_palette(paletteIdx, 2&)) Or (CLng(et4000_palette(paletteIdx, 1&)) * &H100&) Or (CLng(et4000_palette(paletteIdx, 0&)) * &H10000)
                        For yadd = 0& To yscanpixels - 1&
                            For xadd = 0& To xscanpixels - 1&
                                et4000_framebuffer(((scy + yadd) * VGA_FB_MAX) + scx + xadd) = color32
                            Next xadd
                        Next yadd
                    Next scx
                Next scy
            End If

        Case VGA_MODE_GRAPHICS_15BPP, VGA_MODE_GRAPHICS_16BPP
            If diag_vga_verbose <> 0& Then et4000_diagLastMode = mode
            baseLinear = startaddr * 4&
            For scy = start_y To end_y Step yscanpixels
                srcY = scy \ yscanpixels
                rowOffset = baseLinear + (srcY * rowStrideBytes)

                For scx = start_x To end_x Step xscanpixels
                    x = scx \ xscanpixels
                    pixelValue = et4000_readPackedWord(rowOffset + (x * 2&))
                    If mode = VGA_MODE_GRAPHICS_15BPP Then
                        color32 = et4000_color15to32(pixelValue And &HFFFF&)
                    Else
                        color32 = et4000_color16to32(pixelValue And &HFFFF&)
                    End If

                    For yadd = 0& To yscanpixels - 1&
                        If (scy + yadd) > end_y Then Exit For
                        rowBase = ((scy + yadd) * VGA_FB_MAX) + scx
                        For xadd = 0& To xscanpixels - 1&
                            If (scx + xadd) <= end_x Then
                                et4000_framebuffer(rowBase + xadd) = color32
                            End If
                        Next xadd
                    Next yadd
                Next scx
            Next scy

        Case VGA_MODE_GRAPHICS_4BPP
            If diag_vga_verbose <> 0& Then et4000_diagLastMode = VGA_MODE_GRAPHICS_4BPP
            palettePage = CLng(et4000_attrd(&H14&)) * 16&
            hiPage = (CLng(et4000_attrd(&H14&)) And 3&) * 16&
            If (et4000_attrd(&H10&) And &H80&) <> 0& Then
                paletteHiMode = 1&
            Else
                paletteHiMode = 0&
            End If
            For cc = 0& To 15&
                paletteIdx = et4000_attrd(cc) Or palettePage
                If paletteHiMode <> 0& Then
                    paletteIdx = (paletteIdx And &HCF&) Or hiPage
                End If
                paletteIdx = paletteIdx And &HFF&
                et4000_attrColorCache(cc) = CLng(et4000_palette(paletteIdx, 2&)) Or (CLng(et4000_palette(paletteIdx, 1&)) * &H100&) Or (CLng(et4000_palette(paletteIdx, 0&)) * &H10000)
            Next cc
            byteStartX = (start_x \ xscanpixels) \ 8&
            byteEndX = (end_x \ xscanpixels) \ 8&

            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels
                For byteX = byteStartX To byteEndX
                    addr = (y * xstride) + byteX
                    addr = addr + startaddr
                    plane0Byte = et4000_readPlaneByte(addr, 0&)
                    plane1Byte = et4000_readPlaneByte(addr, 1&)
                    plane2Byte = et4000_readPlaneByte(addr, 2&)
                    plane3Byte = et4000_readPlaneByte(addr, 3&)

                    For pixelBit = 0& To 7&
                        srcPixelX = (byteX * 8&) + pixelBit
                        destX = srcPixelX * xscanpixels
                        If destX > end_x Then Exit For
                        If (destX + xscanpixels - 1&) >= start_x Then
                            bitMaskVal = et4000_glyphBitMask(pixelBit)
                            cc = 0&
                            If (plane0Byte And bitMaskVal) <> 0& Then cc = cc Or 1&
                            If (plane1Byte And bitMaskVal) <> 0& Then cc = cc Or 2&
                            If (plane2Byte And bitMaskVal) <> 0& Then cc = cc Or 4&
                            If (plane3Byte And bitMaskVal) <> 0& Then cc = cc Or 8&
                            color32 = et4000_attrColorCache(cc And &HF&)

                            For yadd = 0& To yscanpixels - 1&
                                If (scy + yadd) > end_y Then Exit For
                                rowBase = ((scy + yadd) * VGA_FB_MAX) + destX
                                For xadd = 0& To xscanpixels - 1&
                                    If (destX + xadd) >= start_x Then
                                        If (destX + xadd) <= end_x Then
                                            et4000_framebuffer(rowBase + xadd) = color32
                                        End If
                                    End If
                                Next xadd
                            Next yadd
                        End If
                    Next pixelBit
                Next byteX
            Next scy

        Case VGA_MODE_GRAPHICS_2BPP
            If diag_vga_verbose <> 0& Then et4000_diagLastMode = VGA_MODE_GRAPHICS_2BPP
            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels
                isodd = y And 1&
                y = y \ 2&

                For scx = start_x To end_x Step xscanpixels
                    x = scx \ xscanpixels
                    addr = (8192& * isodd) + (y * xstride) + (x \ pixelsperbyte)
                    addr = addr + startaddr
                    shift = (3& - (x And 3&)) * 2&
                    cc = (et4000_readLinear(et4000_translateAddress(addr)) \ (2& ^ shift)) And 3&

                    color32 = et4000_attrd(cc And &HF&) Or (CLng(et4000_attrd(&H14&)) * 16&)
                    If (et4000_attrd(&H10&) And &H80&) <> 0& Then
                        color32 = (color32 And &HCF&) Or ((et4000_attrd(&H14&) And 3&) * 16&)
                    End If
                    paletteIdx = color32 And &HFF&
                    color32 = CLng(et4000_palette(paletteIdx, 2&)) Or (CLng(et4000_palette(paletteIdx, 1&)) * &H100&) Or (CLng(et4000_palette(paletteIdx, 0&)) * &H10000)

                    For yadd = 0& To yscanpixels - 1&
                        For xadd = 0& To xscanpixels - 1&
                            et4000_framebuffer(((scy + yadd) * VGA_FB_MAX) + scx + xadd) = color32
                        Next xadd
                    Next yadd
                Next scx
            Next scy

        Case VGA_MODE_GRAPHICS_1BPP
            If diag_vga_verbose <> 0& Then et4000_diagLastMode = VGA_MODE_GRAPHICS_1BPP
            For scy = start_y To end_y Step yscanpixels
                y = scy \ yscanpixels
                isodd = y And 1&
                y = y \ 2&

                For scx = start_x To end_x Step xscanpixels
                    x = scx \ xscanpixels
                    addr = (8192& * isodd) + (y * xstride) + (x \ pixelsperbyte)
                    addr = addr + startaddr
                    shift = 7& - (x And 7&)
                    cc = (et4000_readLinear(et4000_translateAddress(addr)) \ (2& ^ shift)) And 1&

                    If cc <> 0& Then
                        color32 = &HFFFFFF
                    Else
                        color32 = 0&
                    End If

                    For yadd = 0& To yscanpixels - 1&
                        For xadd = 0& To xscanpixels - 1&
                            et4000_framebuffer(((scy + yadd) * VGA_FB_MAX) + (scx + xadd)) = color32
                        Next xadd
                    Next yadd
                Next scx
            Next scy
    End Select
End Sub

Public Sub et4000_sendBlit()
    console_blit VarPtr(et4000_framebuffer(0&)), et4000_renderW, et4000_renderH, (VGA_FB_MAX * 4&)
End Sub

Private Sub et4000_renderThread(ByVal dummy As Long)
    et4000_renderW = et4000_w
    et4000_renderH = et4000_h

    If et4000_renderW < 1& Then et4000_renderW = 1&
    If et4000_renderH < 1& Then et4000_renderH = 1&
    If et4000_renderW > VGA_FB_MAX Then et4000_renderW = VGA_FB_MAX
    If et4000_renderH > VGA_FB_MAX Then et4000_renderH = VGA_FB_MAX

    et4000_update 0&, 0&, et4000_renderW - 1&, et4000_renderH - 1&
    et4000_doBlitNow = True
End Sub

Private Sub et4000_startRenderThread()
    If et4000_threadStarted <> 0& Then Exit Sub

    et4000_threadHandle = CreateThread(ByVal 0&, 0&, AddressOf et4000_renderThreadProc, ByVal 0&, 0&, et4000_threadId)
    If et4000_threadHandle = 0& Then
        et4000_threadStarted = 0&
        debug_log DEBUG_INFO, "[ET4000] Render thread start failed, using main thread rendering"
        Exit Sub
    End If

    et4000_threadStarted = 1&
    CloseHandle et4000_threadHandle
    et4000_threadHandle = 0&
End Sub

Public Function et4000_renderThreadProc(ByVal lpParam As Long) As Long
    Do While running <> 0&
        If et4000_doRender <> 0& Then
            et4000_doRender = 0&
            et4000_renderThread 0&
        End If

        If et4000_doBlit <> 0& Then
            et4000_doBlit = 0&
            et4000_doBlitNow = True
        End If
    Loop

    et4000_renderThreadProc = 0&
End Function

Private Sub et4000_calcmemorymap()
    Select Case (et4000_gfxd(&H6&) And &HC&)
        Case &H0&
            et4000_membase = 0&
            et4000_memmask = &H1FFFF
        Case &H4&
            et4000_membase = 0&
            et4000_memmask = &HFFFF&
        Case &H8&
            et4000_membase = &H10000
            et4000_memmask = &H7FFF&
        Case &HC&
            et4000_membase = &H18000
            et4000_memmask = &H7FFF&
    End Select

    et4000_updateBanking
End Sub

Private Sub et4000_calcscreensize()
    Dim h As Long
    Dim hdisp As Long
    Dim textMode As Long

    hdisp = CLng(et4000_crtcd(&H1&))
    If (hdisp And 1&) <> 0& Then hdisp = hdisp + 1&

    If ((et4000_gfxd(&H6&) And 1&) = 0&) And ((et4000_attrd(&H10&) And 1&) = 0&) Then
        textMode = 1&
    Else
        textMode = 0&
    End If

    If textMode <> 0& Then
        If (et4000_seqd(&H1&) And &H8&) <> 0& Then
            If (et4000_seqd(&H1&) And &H1&) <> 0& Then
                hdisp = hdisp * 16&
            Else
                hdisp = hdisp * 18&
            End If
        Else
            If (et4000_seqd(&H1&) And &H1&) <> 0& Then
                hdisp = hdisp * 8&
            Else
                hdisp = hdisp * 9&
            End If
        End If
    Else
        If (et4000_seqd(&H1&) And &H8&) <> 0& Then
            hdisp = hdisp * 16&
        Else
            hdisp = hdisp * 8&
        End If
    End If

    If (et4000_attrd(&H16&) And &H20&) <> 0& Then hdisp = hdisp * 2&

    h = 1& + CLng(et4000_crtcd(&H12&))
    If (et4000_crtcd(&H7&) And 2&) <> 0& Then h = h Or &H100&
    If (et4000_crtcd(&H7&) And &H40&) <> 0& Then h = h Or &H200&
    If (et4000_crtcd(&H35&) And &H4&) <> 0& Then h = h Or &H400&
    et4000_h = h

    Select Case et4000_bpp
        Case 15&, 16&
            hdisp = hdisp \ 2&
        Case 24&
            hdisp = hdisp \ 3&
    End Select

    If (hdisp = 320&) And (et4000_h >= 400&) And ((et4000_attrd(&H10&) And 1&) <> 0&) Then
        If ((et4000_gfxd(&H5&) And &H60&) >= &H40&) Then hdisp = 640&
    End If

    et4000_w = hdisp

    If et4000_w < 1& Then et4000_w = 1&
    If et4000_h < 1& Then et4000_h = 1&
    If et4000_w > VGA_FB_MAX Then et4000_w = VGA_FB_MAX
    If et4000_h > VGA_FB_MAX Then et4000_h = VGA_FB_MAX

    et4000_updateScanlineTiming
End Sub

Private Function et4000_lowresEnabled() As Long
    If (et4000_attrd(&H10&) And &H40&) <> 0& Then
        et4000_lowresEnabled = 1&
    Else
        et4000_lowresEnabled = 0&
    End If
End Function

Private Sub et4000_updateChain4State()
    et4000_chain4 = et4000_seqd(&H4&) And &H8&
    If ((et4000_gfxd(&H5&) And &H40&) <> 0&) And (et4000_lowresEnabled() <> 0&) Then
        et4000_chain4 = et4000_chain4 Or (et4000_seqd(&HE&) And &H2&)
    End If
End Sub

Private Function et4000_readcrtci() As Byte
    et4000_readcrtci = et4000_crtci
End Function

Private Function et4000_readcrtcd() As Byte
    If et4000_crtci <= ET4000_CRTC_MAX Then
        et4000_readcrtcd = et4000_crtcd(et4000_crtci)
    Else
        et4000_readcrtcd = &HFF&
    End If
End Function

Private Sub et4000_writecrtci(ByVal value As Byte)
    et4000_crtci = value And &H3F&
End Sub

Private Sub et4000_writecrtcd(ByVal value As Byte)
    If et4000_crtci > ET4000_CRTC_MAX Then Exit Sub
    If (et4000_crtci < &H7&) And ((et4000_crtcd(&H11&) And &H80&) <> 0&) Then Exit Sub
    If (et4000_crtci = &H35&) And ((et4000_crtcd(&H11&) And &H80&) <> 0&) Then Exit Sub
    If (et4000_crtci = &H7&) And ((et4000_crtcd(&H11&) And &H80&) <> 0&) Then
        value = (et4000_crtcd(&H7&) And &HEF&) Or (value And &H10&)
    End If

    et4000_crtcd(et4000_crtci) = value

    Select Case et4000_crtci
        Case &H1&, &H2&, &H7&, &H12&, &H31&, &H33&, &H34&, &H35&, &H37&, &H3F&
            et4000_calcscreensize
        Case &HC&, &HD&
            et4000_updateScanlineTiming
        Case &H36&
            et4000_calcscreensize
            et4000_updateBanking
    End Select
End Sub

Public Sub et4000_writeport(ByVal dummy As Long, ByVal port As Integer, ByVal value As Byte)
    diag_count_vga_port
    Dim p As Long
    Dim idx As Long

    p = port And &HFFFF&

    Select Case p
        Case &H3B4&
            If (et4000_misc And 1&) = 0& Then et4000_writecrtci value

        Case &H3B5&
            If (et4000_misc And 1&) = 0& Then et4000_writecrtcd value

        Case &H3C0&, &H3C1&
            If et4000_attrflipflop = 0& Then
                et4000_attri = value And &H1F&
                et4000_attrpal = value And &H20&
            Else
                If et4000_attri <= ET4000_ATTR_MAX Then
                    et4000_attrd(et4000_attri) = value
                    Select Case et4000_attri
                        Case &H10&, &H12&, &H16&
                            If et4000_attri = &H10& Then et4000_updateChain4State
                            et4000_calcscreensize
                    End Select
                End If
            End If
            et4000_attrflipflop = et4000_attrflipflop Xor 1&

        Case &H3C6&
            If (et4000_ramdacState = 4&) Or ((et4000_ramdacCtrl And &H10&) <> 0&) Then
                et4000_ramdacState = 0&
                et4000_ramdacCtrl = value
                If value <> &HFF& Then et4000_updateRamdacBpp
            Else
                et4000_ramdacState = 0&
                et4000_dacMask = value
            End If

        Case &H3C7&
            If (et4000_ramdacCtrl And &H10&) <> 0& Then
                et4000_ramdacIndex = value
            Else
                et4000_DAC.state = VGA_DAC_MODE_READ
                et4000_DAC.index = value
                et4000_DAC.step = 0&
            End If
            et4000_ramdacState = 0&

        Case &H3C8&
            If (et4000_ramdacCtrl And &H10&) <> 0& Then
                Select Case et4000_ramdacIndex
                    Case 8&
                        et4000_ramdacRegs(8&) = value
                    Case &HD&
                        et4000_ramdacPixelMask = (et4000_ramdacPixelMask And &HFFFF00&) Or (value And et4000_dacMask)
                    Case &HE&
                        et4000_ramdacPixelMask = (et4000_ramdacPixelMask And &HFF00FF&) Or ((value And et4000_dacMask) * &H100&)
                    Case &HF&
                        et4000_ramdacPixelMask = (et4000_ramdacPixelMask And &HFFFF&) Or ((value And et4000_dacMask) * &H10000)
                    Case &H10&
                        et4000_ramdacRegs(&H10&) = value
                        et4000_updateRamdacBpp
                    Case Else
                        et4000_ramdacRegs(et4000_ramdacIndex) = value
                End Select
                et4000_ramdacState = 0&
            Else
                et4000_DAC.state = VGA_DAC_MODE_WRITE
                et4000_DAC.index = value
                et4000_DAC.step = 0&
                et4000_ramdacState = 0&
                If diag_vga_verbose <> 0& Then et4000_diagPort3C8Writes = et4000_diagPort3C8Writes + 1&
            End If

        Case &H3C9&
            et4000_ramdacState = 0&
            et4000_writePaletteData value

        Case &H3C2&
            et4000_misc = value

        Case &H3C4&
            et4000_seqi = value And &HF&

        Case &H3C5&
            If et4000_seqi <= ET4000_SEQ_MAX Then
                et4000_seqd(et4000_seqi) = value
                Select Case et4000_seqi
                    Case &H1&
                        If (value And &H1&) <> 0& Then
                            et4000_dots = 8&
                        Else
                            et4000_dots = 9&
                        End If
                        If (value And &H8&) <> 0& Then
                            et4000_dbl = 1&
                        Else
                            et4000_dbl = 0&
                        End If
                        et4000_calcscreensize
                    Case &H2&
                        et4000_enableplane = value And &HF&
                    Case &H4&
                        et4000_updateChain4State
                        et4000_calcscreensize
                    Case &HE&
                        et4000_updateChain4State
                        et4000_calcscreensize
                End Select
            End If

        Case &H3CE&
            et4000_gfxi = value And &H1F&

        Case &H3CF&
            If et4000_gfxi < &H9& Then
                et4000_gfxd(et4000_gfxi) = value
                Select Case et4000_gfxi
                    Case &H3&
                        et4000_rotate = value And 7&
                        et4000_logicop = (value \ 8&) And 3&
                    Case &H4&
                        et4000_readmap = value And 3&
                    Case &H5&
                        et4000_wmode = value And 3&
                        et4000_rmode = (value \ 8&) And 1&
                        et4000_shiftmode = (value \ 32&) And 3&
                        et4000_updateChain4State
                        et4000_calcscreensize
                    Case &H6&
                        et4000_calcmemorymap
                End Select
            End If

        Case &H3CD&
            et4000_bankReg = value
            et4000_updateBanking

        Case &H3D4&
            If (et4000_misc And 1&) = 1& Then et4000_writecrtci value

        Case &H3D5&
            If (et4000_misc And 1&) = 1& Then et4000_writecrtcd value
    End Select

    If ((et4000_seqd(4&) And &HC&) = 4&) And ((et4000_gfxd(5&) And &HB&) = 0&) And ((et4000_gfxd(6&) And &H2&) = 0&) And ((et4000_crtcd(20&) And &H40&) = 0&) And ((et4000_crtcd(23&) And &H40&) <> 0&) Then
        et4000_modeY = 1&
    ElseIf ((et4000_seqd(&HE&) And &H2&) <> 0&) And ((et4000_gfxd(5&) And &H60&) >= &H40&) And (et4000_lowresEnabled() <> 0&) Then
        et4000_modeY = 1&
    Else
        et4000_modeY = 0&
    End If
End Sub

Public Function et4000_readport(ByVal dummy As Long, ByVal port As Integer) As Byte
    Dim p As Long
    Dim ret As Byte

    p = port And &HFFFF&
    ret = &HFF&

    Select Case p
        Case &H3B4&
            If (et4000_misc And 1&) = 0& Then et4000_readport = et4000_readcrtci(): Exit Function

        Case &H3B5&
            If (et4000_misc And 1&) = 0& Then et4000_readport = et4000_readcrtcd(): Exit Function

        Case &H3C0&
            If et4000_attrflipflop = 0& Then
                ret = et4000_attri Or et4000_attrpal
            Else
                If et4000_attri <= ET4000_ATTR_MAX Then ret = et4000_attrd(et4000_attri)
            End If

        Case &H3C1&
            If et4000_attri <= ET4000_ATTR_MAX Then et4000_readport = et4000_attrd(et4000_attri): Exit Function

        Case &H3C6&
            If et4000_ramdacState = 4& Then
                et4000_readport = et4000_ramdacCtrl
            Else
                et4000_ramdacState = et4000_ramdacState + 1&
                et4000_readport = et4000_dacMask
            End If
            Exit Function

        Case &H3C4&
            et4000_readport = et4000_seqi
            Exit Function

        Case &H3C5&
            If et4000_seqi = 7& Then
                et4000_readport = (et4000_seqd(7&) Or &H4&)
                Exit Function
            End If
            If et4000_seqi <= ET4000_SEQ_MAX Then et4000_readport = et4000_seqd(et4000_seqi): Exit Function

        Case &H3C7&
            et4000_ramdacState = 0&
            et4000_readport = et4000_DAC.state
            Exit Function

        Case &H3C8&
            If (et4000_ramdacCtrl And &H10&) <> 0& Then
                Select Case et4000_ramdacIndex
                    Case 9&
                        ret = &H53&
                    Case &HA&
                        ret = &H3A&
                    Case &HB&
                        ret = &HB1&
                    Case &HC&
                        ret = &H41&
                    Case &HD&
                        ret = CByte(et4000_ramdacPixelMask And &HFF&)
                    Case &HE&
                        ret = CByte((et4000_ramdacPixelMask \ &H100&) And &HFF&)
                    Case &HF&
                        ret = CByte((et4000_ramdacPixelMask \ &H10000) And &HFF&)
                    Case Else
                        ret = et4000_ramdacRegs(et4000_ramdacIndex)
                End Select
                et4000_ramdacState = 0&
                et4000_readport = ret
            Else
                et4000_ramdacState = 0&
                et4000_readport = et4000_DAC.index
            End If
            Exit Function

        Case &H3C9&
            If (et4000_ramdacCtrl And &H10&) <> 0& Then
                ret = et4000_ramdacIndex
                et4000_ramdacState = 0&
            Else
                et4000_ramdacState = 0&
                ret = et4000_readPaletteData()
            End If

        Case &H3CC&
            et4000_readport = et4000_misc
            Exit Function

        Case &H3CE&
            et4000_readport = et4000_gfxi
            Exit Function

        Case &H3CF&
            If et4000_gfxi < &H9& Then et4000_readport = et4000_gfxd(et4000_gfxi): Exit Function

        Case &H3CD&
            et4000_readport = et4000_bankReg
            Exit Function

        Case &H3D4&
            If (et4000_misc And 1&) = 1& Then et4000_readport = et4000_readcrtci(): Exit Function

        Case &H3D5&
            If (et4000_misc And 1&) = 1& Then et4000_readport = et4000_readcrtcd(): Exit Function

        Case &H3DA&
            If diag_vga_verbose <> 0& Then
                et4000_diag3DAReads = et4000_diag3DAReads + 1&
                If (et4000_status1 And &H8&) <> 0& Then et4000_diag3DAVBlankReads = et4000_diag3DAVBlankReads + 1&
                If (et4000_status1 And &H1&) <> 0& Then et4000_diag3DAHBlankReads = et4000_diag3DAHBlankReads + 1&
            End If
            et4000_attrflipflop = 0&
            If (et4000_status1 And &H1&) <> 0& Then
                et4000_status1 = et4000_status1 And &HCF&
            Else
                et4000_status1 = et4000_status1 Xor &H30&
            End If
            ret = et4000_status1
            If (ret And &H8&) <> 0& Then
                ret = ret And &H7F&
            Else
                ret = ret Or &H80&
            End If
            et4000_readport = ret
            Exit Function
    End Select

    et4000_readport = ret
End Function

Private Function et4000_dologic(ByVal value As Byte, ByVal latch As Byte) As Byte
    Select Case et4000_logicop
        Case 0&
            et4000_dologic = value
        Case 1&
            et4000_dologic = value And latch
        Case 2&
            et4000_dologic = value Or latch
        Case Else
            et4000_dologic = value Xor latch
    End Select
End Function

Private Function et4000_host_chain4_enabled() As Long
    If et4000_chain4 <> 0& Then
        et4000_host_chain4_enabled = 1&
    Else
        et4000_host_chain4_enabled = 0&
    End If
End Function

Private Function et4000_host_odd_even_write_enabled() As Long
    If (et4000_host_chain4_enabled() = 0&) And ((et4000_seqd(&H4&) And &H4&) = 0&) Then
        et4000_host_odd_even_write_enabled = 1&
    Else
        et4000_host_odd_even_write_enabled = 0&
    End If
End Function

Private Function et4000_host_odd_even_read_enabled() As Long
    If (et4000_host_chain4_enabled() = 0&) And ((et4000_gfxd(&H5&) And &H10&) <> 0&) And ((et4000_gfxd(&H6&) And &H2&) <> 0&) Then
        et4000_host_odd_even_read_enabled = 1&
    Else
        et4000_host_odd_even_read_enabled = 0&
    End If
End Function

Public Sub et4000_writememory(ByVal dummy As Long, ByVal addr As Long, ByVal value As Byte)
    diag_count_vga_mem
    If diag_vga_verbose <> 0& Then
        et4000_diagMemWrites = et4000_diagMemWrites + 1&
        If value <> 0& Then et4000_diagMemNonZeroWrites = et4000_diagMemNonZeroWrites + 1&
    End If
    Dim temp As Byte
    Dim plane As Long
    Dim a As Long
    Dim bitMask As Byte
    Dim parity As Long
    Dim planeMask As Long
    Dim packedLane As Long
    Dim packedBase As Long

    If (et4000_misc And &H2&) = 0& Then GoTo WriteDone

    If et4000_decodeHostAddress(addr, 1&, a) = 0& Then GoTo WriteDone

    If et4000_host_chain4_enabled() <> 0& Then
        packedLane = a And 3&
        packedBase = et4000_translateAddress(a - packedLane)

        Select Case et4000_wmode
            Case 0&
                planeMask = CLng(2 ^ packedLane)
                If (et4000_gfxd(&H1&) And planeMask) <> 0& Then
                    If (et4000_gfxd(&H0&) And planeMask) <> 0& Then
                        temp = &HFF&
                    Else
                        temp = 0&
                    End If
                Else
                    temp = et4000_dorotate(value)
                End If

                temp = et4000_dologic(temp, et4000_latch(packedLane))
                temp = (temp And et4000_gfxd(&H8&)) Or (et4000_latch(packedLane) And et4000_notbyte(et4000_gfxd(&H8&)))
                et4000_writeLinear packedBase Or packedLane, temp

            Case 1&
                et4000_writeLinear packedBase Or packedLane, et4000_latch(packedLane)

            Case 2&
                planeMask = CLng(2 ^ packedLane)
                If (value And planeMask) <> 0& Then
                    temp = &HFF&
                Else
                    temp = 0&
                End If
                temp = et4000_dologic(temp, et4000_latch(packedLane))
                temp = (temp And et4000_gfxd(&H8&)) Or (et4000_latch(packedLane) And et4000_notbyte(et4000_gfxd(&H8&)))
                et4000_writeLinear packedBase Or packedLane, temp

            Case 3&
                planeMask = CLng(2 ^ packedLane)
                bitMask = et4000_dorotate(value) And et4000_gfxd(&H8&)
                If (et4000_gfxd(&H0&) And planeMask) <> 0& Then
                    temp = &HFF&
                Else
                    temp = 0&
                End If
                temp = et4000_dologic(temp, et4000_latch(packedLane))
                temp = (temp And bitMask) Or (et4000_latch(packedLane) And et4000_notbyte(bitMask))
                et4000_writeLinear packedBase Or packedLane, temp
        End Select
        GoTo WriteDone
    End If

    If et4000_host_odd_even_write_enabled() <> 0& Then
        parity = a And 1&
        a = (a \ 2&)

        planeMask = CLng(2 ^ parity)
        If (et4000_enableplane And planeMask) <> 0& Then
            et4000_writePlaneByte a, parity, value
        End If

        planeMask = CLng(2 ^ (parity + 2&))
        If (et4000_enableplane And planeMask) <> 0& Then
            et4000_writePlaneByte a, parity + 2&, value
        End If
        GoTo WriteDone
    End If

    Select Case et4000_wmode
        Case 0&
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (et4000_enableplane And planeMask) <> 0& Then
                    If (et4000_gfxd(&H1&) And planeMask) <> 0& Then
                        If (et4000_gfxd(&H0&) And planeMask) <> 0& Then
                            temp = &HFF&
                        Else
                            temp = 0&
                        End If
                    Else
                        temp = et4000_dorotate(value)
                    End If

                    temp = et4000_dologic(temp, et4000_latch(plane))
                    temp = (temp And et4000_gfxd(&H8&)) Or (et4000_latch(plane) And et4000_notbyte(et4000_gfxd(&H8&)))
                    et4000_writePlaneByte a, plane, temp
                End If
            Next plane

        Case 1&
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (et4000_enableplane And planeMask) <> 0& Then
                    et4000_writePlaneByte a, plane, et4000_latch(plane)
                End If
            Next plane

        Case 2&
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (et4000_enableplane And planeMask) <> 0& Then
                    If (value And planeMask) <> 0& Then
                        temp = &HFF&
                    Else
                        temp = 0&
                    End If
                    temp = et4000_dologic(temp, et4000_latch(plane))
                    temp = (temp And et4000_gfxd(&H8&)) Or (et4000_latch(plane) And et4000_notbyte(et4000_gfxd(&H8&)))
                    et4000_writePlaneByte a, plane, temp
                End If
            Next plane

        Case 3&
            bitMask = et4000_dorotate(value) And et4000_gfxd(&H8&)
            For plane = 0& To 3&
                planeMask = CLng(2 ^ plane)
                If (et4000_enableplane And planeMask) <> 0& Then
                    If (et4000_gfxd(&H0&) And planeMask) <> 0& Then
                        temp = &HFF&
                    Else
                        temp = 0&
                    End If
                    temp = et4000_dologic(temp, et4000_latch(plane))
                    temp = (temp And bitMask) Or (et4000_latch(plane) And et4000_notbyte(bitMask))
                    et4000_writePlaneByte a, plane, temp
                End If
            Next plane
    End Select

WriteDone:
End Sub

Public Function et4000_readmemory(ByVal dummy As Long, ByVal addr As Long) As Byte
    Dim plane As Long
    Dim retL As Long
    Dim a As Long
    Dim compare As Long
    Dim planeMask As Long

    If et4000_decodeHostAddress(addr, 0&, a) = 0& Then
        et4000_readmemory = &HFF&
        Exit Function
    End If

    If et4000_host_chain4_enabled() <> 0& Then
        et4000_loadPackedLatches a
        et4000_readmemory = et4000_readLinear(et4000_translateAddress(a))
        Exit Function
    ElseIf et4000_host_odd_even_read_enabled() <> 0& Then
        plane = (a And 1&) Or (et4000_readmap And &H2&)
        a = (a \ 2&)
    Else
        plane = et4000_readmap And 3&
    End If

    et4000_loadLatches a

    If et4000_rmode = 0& Then
        et4000_readmemory = et4000_latch(plane)
    Else
        retL = &HFF&
        For plane = 0& To 3&
            planeMask = CLng(2 ^ plane)
            If (et4000_gfxd(&H7&) And planeMask) <> 0& Then
                If (et4000_gfxd(&H2&) And planeMask) <> 0& Then
                    compare = &HFF&
                Else
                    compare = 0&
                End If
                retL = retL And (((CLng(et4000_latch(plane)) Xor compare) Xor &HFF&) And &HFF&)
            End If
        Next plane
        et4000_readmemory = CByte(retL And &HFF&)
    End If
End Function

Public Sub et4000_drawCallback(ByVal dummy As Long)
    If et4000_threadStarted = 0& Then
        et4000_renderThread 0&
        et4000_doBlitNow = True
    Else
        et4000_doRender = 1&
        et4000_doBlit = 1&
    End If
End Sub

Public Sub et4000_blinkCallback(ByVal dummy As Long)
    et4000_cursor_blink_state = et4000_cursor_blink_state Xor 1&
End Sub

Public Sub et4000_hblankCallback(ByVal dummy As Long)
    Dim vblankStartScan As Long
    Dim vblankEndScan As Long

    timing_timerEnable et4000_hblankEndTimer
    et4000_status1 = et4000_status1 Or 1&

    vblankStartScan = (CLng(et4000_vblankstart) And &H7FF&)
    vblankEndScan = (CLng(et4000_vblankend) And &H7FF&)

    et4000_curScanline = ((et4000_curScanline + 1&) And &H7FF&)

    If et4000_curScanline = vblankStartScan Then
        et4000_status1 = et4000_status1 Or 8&
    ElseIf et4000_curScanline = vblankEndScan Then
        et4000_curScanline = 0&
        et4000_status1 = et4000_status1 And &HF7&
    End If
End Sub

Public Sub et4000_hblankEndCallback(ByVal dummy As Long)
    timing_timerDisable et4000_hblankEndTimer
    et4000_status1 = et4000_status1 And &HFE&
End Sub

Public Sub et4000_dumpregs()
    ' Kept for parity; detailed register dump can be added behind a debug gate.
End Sub

Public Function et4000_diagSnapshotAndReset() As String
    Dim startaddr As Long
    Dim modeVal As Long
    Dim outStr As String
    Dim vbStart As Long
    Dim vbEnd As Long

    startaddr = et4000_startAddress()
    modeVal = (et4000_attrd(&H10&) And 1&)
    vbStart = CLng(et4000_vblankstart)
    vbEnd = CLng(et4000_vblankend)

    outStr = ""
    outStr = outStr & " vmode=" & CStr(et4000_diagLastMode)
    outStr = outStr & " chain4=" & CStr(et4000_chain4)
    outStr = outStr & " modeY=" & CStr(et4000_modeY)
    outStr = outStr & " sh=" & CStr(et4000_shiftmode And &H3&)
    outStr = outStr & " wm=" & CStr(et4000_wmode And &H3&)
    outStr = outStr & " rm=" & CStr(et4000_rmode And &H1&)
    outStr = outStr & " em=" & right$("0" & Hex$(et4000_enableplane And &HF&), 1&)
    outStr = outStr & " mm=" & right$("00" & Hex$(et4000_gfxd(&H6&) And &HC&), 2&)
    outStr = outStr & " misc=" & right$("00" & Hex$(et4000_misc), 2&)
    outStr = outStr & " gfx=" & CStr(modeVal)
    outStr = outStr & " bpp=" & CStr(et4000_bpp)
    outStr = outStr & " sa=" & right$("00000" & Hex$(startaddr And &H3FFFF), 5&)
    outStr = outStr & " off=" & right$("00" & Hex$(et4000_crtcd(&H13&)), 2&)
    outStr = outStr & " bank=" & right$("00" & Hex$(et4000_bankReg), 2&)
    outStr = outStr & " rdac=" & right$("00" & Hex$(et4000_ramdacCtrl), 2&) & ":" & right$("00" & Hex$(et4000_ramdacRegs(&H10&)), 2&)
    outStr = outStr & " dac=" & CStr(et4000_diagPort3C8Writes) & "/" & CStr(et4000_diagPort3C9Writes)
    outStr = outStr & " vram=" & CStr(et4000_diagMemWrites) & ":" & CStr(et4000_diagMemNonZeroWrites)
    outStr = outStr & " rd3da=" & CStr(et4000_diag3DAReads) & ":" & CStr(et4000_diag3DAVBlankReads) & ":" & CStr(et4000_diag3DAHBlankReads)
    outStr = outStr & " s1=" & right$("00" & Hex$(et4000_status1), 2&) & " sl=" & CStr(et4000_curScanline) & " vb=" & CStr(vbStart) & "-" & CStr(vbEnd)

    et4000_diagPort3C8Writes = 0&
    et4000_diagPort3C9Writes = 0&
    et4000_diagMemWrites = 0&
    et4000_diagMemNonZeroWrites = 0&
    et4000_diag3DAReads = 0&
    et4000_diag3DAVBlankReads = 0&
    et4000_diag3DAHBlankReads = 0&

    et4000_diagSnapshotAndReset = outStr
End Function

Private Function et4000_color(ByVal c As Long) As Long
    Dim idx As Long

    idx = c And CLng(et4000_dacMask)
    et4000_color = CLng(et4000_palette(idx, 2&)) Or (CLng(et4000_palette(idx, 1&)) * &H100&) Or (CLng(et4000_palette(idx, 0&)) * &H10000)
End Function

Private Function et4000_dorotate(ByVal v As Byte) As Byte
    Dim r As Long
    Dim src As Long
    Dim outv As Long

    r = et4000_rotate And 7&
    src = v

    If r = 0& Then
        et4000_dorotate = v
        Exit Function
    End If

    outv = ((src \ (2& ^ r)) Or ((src * (2& ^ (8& - r))) And &HFF&)) And &HFF&
    et4000_dorotate = CByte(outv)
End Function

Private Function et4000_notbyte(ByVal v As Byte) As Byte
    et4000_notbyte = CByte(v Xor &HFF&)
End Function

Private Function et4000_fontbase(ByVal idx As Long) As Long
    Select Case (idx And 7&)
        Case 0&: et4000_fontbase = &H0&
        Case 1&: et4000_fontbase = &H4000&
        Case 2&: et4000_fontbase = &H8000&
        Case 3&: et4000_fontbase = &HC000&
        Case 4&: et4000_fontbase = &H2000&
        Case 5&: et4000_fontbase = &H6000&
        Case 6&: et4000_fontbase = &HA000&
        Case Else: et4000_fontbase = &HE000&
    End Select
End Function

Private Function et4000_clockHz(ByVal clockSel As Long) As Double
    Select Case (clockSel And &HF&)
        Case &H0&: et4000_clockHz = 50000000#
        Case &H1&: et4000_clockHz = 56644000#
        Case &H2&: et4000_clockHz = 65000000#
        Case &H3&: et4000_clockHz = 72000000#
        Case &H4&: et4000_clockHz = 80000000#
        Case &H5&: et4000_clockHz = 89800000#
        Case &H6&: et4000_clockHz = 63000000#
        Case &H7&: et4000_clockHz = 75000000#
        Case &H8&: et4000_clockHz = 83078000#
        Case &H9&: et4000_clockHz = 93463000#
        Case &HA&: et4000_clockHz = 100000000#
        Case &HB&: et4000_clockHz = 104000000#
        Case &HC&: et4000_clockHz = 108000000#
        Case &HD&: et4000_clockHz = 120000000#
        Case &HE&: et4000_clockHz = 130000000#
        Case &HF&: et4000_clockHz = 134700000#
        Case Else: et4000_clockHz = 50000000#
    End Select
End Function

Private Sub et4000_updateBanking()
    If ((et4000_crtcd(&H36&) And &H10&) = 0&) And ((et4000_gfxd(&H6&) And &H8&) = 0&) Then
        et4000_writeBank = (CLng(et4000_bankReg And &HF&) * &H10000)
        et4000_readBank = (CLng((et4000_bankReg \ &H10&) And &HF&) * &H10000)
    Else
        et4000_writeBank = 0&
        et4000_readBank = 0&
    End If
End Sub

Private Function et4000_startAddress() As Long
    et4000_startAddress = (CLng(et4000_crtcd(&HC&)) * &H100&) Or CLng(et4000_crtcd(&HD&))
    et4000_startAddress = et4000_startAddress + ((CLng(et4000_crtcd(&H8&)) And &H60&) \ &H20&)
    et4000_startAddress = et4000_startAddress Or (CLng(et4000_crtcd(&H33&) And 3&) * &H10000)
End Function

Private Function et4000_cursorAddress() As Long
    et4000_cursorAddress = (CLng(et4000_crtcd(&HE&)) * &H100&) Or CLng(et4000_crtcd(&HF&))
    et4000_cursorAddress = et4000_cursorAddress + ((CLng(et4000_crtcd(&HB&)) And &H60&) \ &H20&)
End Function

Private Function et4000_lineCompare() As Long
    et4000_lineCompare = CLng(et4000_crtcd(&H18&))
    If (et4000_crtcd(&H7&) And &H10&) <> 0& Then et4000_lineCompare = et4000_lineCompare Or &H100&
    If (et4000_crtcd(&H9&) And &H40&) <> 0& Then et4000_lineCompare = et4000_lineCompare Or &H200&
    If (et4000_crtcd(&H35&) And &H10&) <> 0& Then et4000_lineCompare = et4000_lineCompare Or &H400&
End Function

Private Function et4000_translateAddress(ByVal addr As Long) As Long
    Dim nbank As Long

    Select Case (et4000_crtcd(&H37&) And &HB&)
        Case &H0&, &H1&
            nbank = 0&
            addr = addr And &HFFFF&
        Case &H2&
            nbank = (addr And 1&) * 2&
            addr = (addr \ 2&) And &HFFFF&
        Case &H3&
            nbank = addr And 3&
            addr = (addr \ 4&) And &HFFFF&
        Case &H8&, &H9&
            nbank = 0&
            addr = addr And &H3FFFF
        Case &HA&
            nbank = (addr And 1&) * 2&
            addr = (addr \ 2&) And &H3FFFF
        Case &HB&
            nbank = addr And 3&
            addr = (addr \ 4&) And &H3FFFF
        Case Else
            nbank = 0&
    End Select

    addr = (addr * 4&) Or (nbank And 3&)
    If (et4000_crtcd(&H37&) And 3&) = 2& Then
        addr = addr \ 2&
    ElseIf (et4000_crtcd(&H37&) And 3&) < 2& Then
        addr = addr \ 4&
    End If

    et4000_translateAddress = (addr And ET4000_VRAM_MASK)
End Function

Private Function et4000_linearIndex(ByVal linearAddr As Long) As Long
    et4000_linearIndex = ((linearAddr And ET4000_VRAM_MASK) \ 4&)
End Function

Private Function et4000_readLinear(ByVal linearAddr As Long) As Byte
    et4000_readLinear = et4000_RAM(linearAddr And 3&, et4000_linearIndex(linearAddr))
End Function

Private Sub et4000_writeLinear(ByVal linearAddr As Long, ByVal value As Byte)
    et4000_RAM(linearAddr And 3&, et4000_linearIndex(linearAddr)) = value
End Sub

Private Function et4000_rowOffsetBytes() As Long
    et4000_rowOffsetBytes = CLng(et4000_crtcd(&H13&))
    If et4000_rowOffsetBytes = 0& Then et4000_rowOffsetBytes = &H100&
    et4000_rowOffsetBytes = et4000_rowOffsetBytes * 8&
End Function

Private Function et4000_readPackedWord(ByVal byteAddr As Long) As Long
    et4000_readPackedWord = CLng(et4000_readLinear(et4000_translateAddress(byteAddr)))
    et4000_readPackedWord = et4000_readPackedWord Or (CLng(et4000_readLinear(et4000_translateAddress(byteAddr + 1&))) * &H100&)
End Function

Private Function et4000_readPlaneByte(ByVal cellAddr As Long, ByVal plane As Long) As Byte
    Dim baseLinear As Long

    baseLinear = et4000_translateAddress((cellAddr And (VGA_RAM_PLANE_SIZE - 1&)) * 4&)
    et4000_readPlaneByte = et4000_readLinear(baseLinear Or (plane And 3&))
End Function

Private Sub et4000_writePlaneByte(ByVal cellAddr As Long, ByVal plane As Long, ByVal value As Byte)
    Dim baseLinear As Long

    baseLinear = et4000_translateAddress((cellAddr And (VGA_RAM_PLANE_SIZE - 1&)) * 4&)
    et4000_writeLinear baseLinear Or (plane And 3&), value
End Sub

Private Sub et4000_loadLatches(ByVal cellAddr As Long)
    Dim baseLinear As Long

    baseLinear = et4000_translateAddress((cellAddr And (VGA_RAM_PLANE_SIZE - 1&)) * 4&)
    et4000_latch(0&) = et4000_readLinear(baseLinear Or 0&)
    et4000_latch(1&) = et4000_readLinear(baseLinear Or 1&)
    et4000_latch(2&) = et4000_readLinear(baseLinear Or 2&)
    et4000_latch(3&) = et4000_readLinear(baseLinear Or 3&)
End Sub

Private Sub et4000_loadPackedLatches(ByVal byteAddr As Long)
    Dim baseLinear As Long

    baseLinear = et4000_translateAddress(byteAddr)
    baseLinear = baseLinear - (baseLinear And 3&)
    et4000_latch(0&) = et4000_readLinear(baseLinear Or 0&)
    et4000_latch(1&) = et4000_readLinear(baseLinear Or 1&)
    et4000_latch(2&) = et4000_readLinear(baseLinear Or 2&)
    et4000_latch(3&) = et4000_readLinear(baseLinear Or 3&)
End Sub

Private Function et4000_decodeHostAddress(ByVal addr As Long, ByVal isWrite As Long, ByRef decoded As Long) As Long
    decoded = addr - &HA0000

    Select Case (et4000_gfxd(&H6&) And &HC&)
        Case &H0&
            If (decoded < 0&) Or (decoded >= &H20000) Then
                et4000_decodeHostAddress = 0&
                Exit Function
            End If
        Case &H4&
            If (decoded < 0&) Or (decoded >= &H10000) Then
                et4000_decodeHostAddress = 0&
                Exit Function
            End If
        Case &H8&
            decoded = decoded - &H10000
            If (decoded < 0&) Or (decoded >= &H8000&) Then
                et4000_decodeHostAddress = 0&
                Exit Function
            End If
        Case Else
            decoded = decoded - &H18000
            If (decoded < 0&) Or (decoded >= &H8000&) Then
                et4000_decodeHostAddress = 0&
                Exit Function
            End If
    End Select

    If (et4000_gfxd(&H6&) And &HC&) <= &H4& Then
        If isWrite <> 0& Then
            decoded = decoded + et4000_writeBank
        Else
            decoded = decoded + et4000_readBank
        End If
    End If

    et4000_decodeHostAddress = 1&
End Function

