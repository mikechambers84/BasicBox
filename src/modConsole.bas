Attribute VB_Name = "modConsole"
Option Explicit

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal lpBits As Long, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function InterlockedExchange Lib "kernel32" (ByRef target As Long, ByVal value As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (ByVal lpRect As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal cbCopy As Long)

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0& To 0&) As RGBQUAD
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type DDSCAPS
    dwCaps As Long
End Type

Private Type DDPIXELFORMAT
    dwSize As Long
    dwFlags As Long
    dwFourCC As Long
    dwRGBBitCount As Long
    dwRBitMask As Long
    dwGBitMask As Long
    dwBBitMask As Long
    dwRGBAlphaBitMask As Long
End Type

Private Type DDCOLORKEY
    dwColorSpaceLowValue As Long
    dwColorSpaceHighValue As Long
End Type

Private Type DDSURFACEDESC
    dwSize As Long
    dwFlags As Long
    dwHeight As Long
    dwWidth As Long
    lPitch As Long
    dwBackBufferCount As Long
    dwMipMapCount As Long
    dwAlphaBitDepth As Long
    dwReserved As Long
    lpSurface As Long
    ddckCKDestOverlay As DDCOLORKEY
    ddckCKDestBlt As DDCOLORKEY
    ddckCKSrcOverlay As DDCOLORKEY
    ddckCKSrcBlt As DDCOLORKEY
    ddpfPixelFormat As DDPIXELFORMAT
    ddsCaps As DDSCAPS
End Type

Private Type DIOBJECTDATAFORMAT
    pguid As Long
    dwOfs As Long
    dwType As Long
    dwFlags As Long
End Type

Private Type DIDATAFORMAT
    dwSize As Long
    dwObjSize As Long
    dwFlags As Long
    dwDataSize As Long
    dwNumObjs As Long
    rgodf As Long
End Type

Private Type DIMOUSESTATE
    lX As Long
    lY As Long
    lZ As Long
    rgbButtons(0& To 3&) As Byte
End Type

Private Type DIPROPHEADER
    dwSize As Long
    dwHeaderSize As Long
    dwObj As Long
    dwHow As Long
End Type

Private Type DIPROPDWORD
    diph As DIPROPHEADER
    dwData As Long
End Type

Private Type DIDEVICEOBJECTDATA
    dwOfs As Long
    dwData As Long
    dwTimeStamp As Long
    dwSequence As Long
End Type

Public Const CONSOLE_EVENT_NONE As Byte = 0&
Public Const CONSOLE_EVENT_KEY As Byte = 1&
Public Const CONSOLE_EVENT_QUIT As Byte = 2&
Public Const CONSOLE_EVENT_DEBUG_1 As Byte = 3&
Public Const CONSOLE_EVENT_DEBUG_2 As Byte = 4&

Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0&
Private Const GDI_ERROR As Long = -1&
Private Const SRCCOPY As Long = &HCC0020
Private Const CONSOLE_EVENT_QUEUE_MASK As Long = &HFF&
Private Const GWL_WNDPROC As Long = -4&
Private Const WM_SYSCOMMAND As Long = &H112&
Private Const WM_LBUTTONDOWN As Long = &H201&
Private Const WM_KEYDOWN As Long = &H100&
Private Const WM_KEYUP As Long = &H101&
Private Const WM_SYSKEYDOWN As Long = &H104&
Private Const WM_SYSKEYUP As Long = &H105&
Private Const SC_KEYMENU As Long = &HF100&
Private Const WH_KEYBOARD_LL As Long = 13&
Private Const HC_ACTION As Long = 0&
Private Const LLKHF_ALTDOWN As Long = &H20&

Private Const DDSCL_NORMAL As Long = &H8&
Private Const DDSD_CAPS As Long = &H1&
Private Const DDSD_HEIGHT As Long = &H2&
Private Const DDSD_WIDTH As Long = &H4&
Private Const DDSD_PIXELFORMAT As Long = &H1000&
Private Const DDSCAPS_PRIMARYSURFACE As Long = &H200&
Private Const DDSCAPS_OFFSCREENPLAIN As Long = &H40&
Private Const DDSCAPS_SYSTEMMEMORY As Long = &H800&
Private Const DDPF_RGB As Long = &H40&
Private Const DDBLT_WAIT As Long = &H1000000
Private Const DDLOCK_WAIT As Long = &H1&

Private Const DIDFT_AXIS As Long = &H3&
Private Const DIDFT_BUTTON As Long = &HC&
Private Const DIDFT_ANYINSTANCE As Long = &HFF00&
Private Const DIDFT_OPTIONAL As Long = &H80000000
Private Const DIDF_RELAXIS As Long = &H2&
Private Const DIPH_DEVICE As Long = 0&
Private Const DIPROP_BUFFERSIZE As Long = 1&
Private Const DI_BUFFERED_MOUSE_EVENTS As Long = 16&
Private Const DISCL_EXCLUSIVE As Long = &H1&
Private Const DISCL_NONEXCLUSIVE As Long = &H2&
Private Const DISCL_FOREGROUND As Long = &H4&

Private Const VK_OEM_1 As Long = &HBA&
Private Const VK_OEM_PLUS As Long = &HBB&
Private Const VK_OEM_COMMA As Long = &HBC&
Private Const VK_OEM_MINUS As Long = &HBD&
Private Const VK_OEM_PERIOD As Long = &HBE&
Private Const VK_OEM_2 As Long = &HBF&
Private Const VK_OEM_3 As Long = &HC0&
Private Const VK_OEM_4 As Long = &HDB&
Private Const VK_OEM_5 As Long = &HDC&
Private Const VK_OEM_6 As Long = &HDD&
Private Const VK_OEM_7 As Long = &HDE&
Private Const VK_LWIN As Long = &H5B&
Private Const VK_RWIN As Long = &H5C&
Private Const VK_APPS As Long = &H5D&

Private Const DIK_F11 As Long = &H57&
Private Const DIK_F12 As Long = &H58&
Private Const DIK_LCONTROL As Long = &H1D&
Private Const DIK_RCONTROL As Long = &H9D&
Private Const DIK_LMENU As Long = &H38&
Private Const DIK_RMENU As Long = &HB8&
Private Const DIK_NUMPADENTER As Long = &H9C&
Private Const DIK_DIVIDE As Long = &HB5&
Private Const DIK_SYSRQ As Long = &HB7&
Private Const DIK_APPS As Long = &HDD&

Private Const CONSOLE_DI_KEY_BYTES As Long = 256&
Private Const CONSOLE_DI_MOUSE_OBJECTS As Long = 7&
Private Const CONSOLE_DI_MOUSE_EVENT_BURST As Long = 64&
Private Const CONSOLE_CTRLALTDEL_STEPS As Long = 6&

Private console_frameTime(0& To 29&) As Double
Private console_keyTimer As Long
Private console_ctrlAltDelTimer As Long
Private console_scancodeSet As Byte
Private console_curkey As Byte
Private console_lastKey As Byte
Private console_frameIdx As Byte
Private console_grabbed As Byte
Private console_ctrl As Byte
Private console_alt As Byte
Private console_ctrlLeft As Byte
Private console_ctrlRight As Byte
Private console_altLeft As Byte
Private console_altRight As Byte
Private console_doRepeat As Byte
Private console_curw As Long
Private console_curh As Long
Private console_title As String
Private console_titleStatus As String
Private console_lastMouseX As Single
Private console_lastMouseY As Single
Private console_formActive As Byte
Private console_cursorHidden As Byte
Private console_prevMouseButtons As Byte
Private console_rawMouseButtons As Byte
Private console_suppressedMouseButtons As Byte
Private console_ctrlF11Shortcut As Byte
Private console_ctrlF12Shortcut As Byte
Private console_ctrlAltDelPos As Long
Private console_cursorClipped As Byte

Private console_eventQueue(0& To CONSOLE_EVENT_QUEUE_MASK) As Byte
Private console_eventData(0& To CONSOLE_EVENT_QUEUE_MASK) As Byte
Private console_eventQRead As Long
Private console_eventQWrite As Long

Private console_pendingPixelsPtr As Long
Private console_pendingW As Long
Private console_pendingH As Long
Private console_pendingStride As Long
Private console_pendingReady As Long

Private console_useDirectDraw As Byte
Private console_dd As Long
Private console_primarySurface As Long
Private console_backSurface As Long
Private console_clipper As Long
Private console_surfaceW As Long
Private console_surfaceH As Long

Private console_di As Long
Private console_keyboard As Long
Private console_mouse As Long
Private console_useDIKeyboard As Byte
Private console_useDIMouse As Byte
Private console_useBufferedDIMouse As Byte
Private console_keyboardAcquired As Byte
Private console_mouseAcquired As Byte
Private console_keyboardFmt As DIDATAFORMAT
Private console_keyboardObj(0& To CONSOLE_DI_KEY_BYTES - 1&) As DIOBJECTDATAFORMAT
Private console_keyboardState(0& To CONSOLE_DI_KEY_BYTES - 1&) As Byte
Private console_prevKeyboardState(0& To CONSOLE_DI_KEY_BYTES - 1&) As Byte
Private console_mouseFmt As DIDATAFORMAT
Private console_mouseObj(0& To CONSOLE_DI_MOUSE_OBJECTS - 1&) As DIOBJECTDATAFORMAT
Private console_guidXAxis As GUID
Private console_guidYAxis As GUID
Private console_guidZAxis As GUID
Private console_prevWndProc As Long
Private console_subclassHwnd As Long
Private console_keyboardHook As Long

Public Function Console_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hwnd = console_subclassHwnd Then
        If uMsg = WM_SYSCOMMAND Then
            If (wParam And &HFFF0&) = SC_KEYMENU Then
                Console_WindowProc = 0&
                Exit Function
            End If
        ElseIf uMsg = WM_LBUTTONDOWN Then
            If console_grabbed = 0& Then
                console_suppressedMouseButtons = console_suppressedMouseButtons Or 1&
                console_mousegrab
                Console_WindowProc = 0&
                Exit Function
            End If
        End If
    End If

    If console_prevWndProc <> 0& Then
        Console_WindowProc = CallWindowProc(console_prevWndProc, hwnd, uMsg, wParam, lParam)
    Else
        Console_WindowProc = 0&
    End If
End Function

Private Function Console_IsForegroundActive() As Long
    Dim fgHwnd As Long
    Dim consoleHwnd As Long

    If console_formActive = 0& Then Exit Function
    If console_isFormLoaded() = 0& Then Exit Function

    consoleHwnd = frmConsole.hWnd
    If consoleHwnd = 0& Then Exit Function

    fgHwnd = GetForegroundWindow()
    If fgHwnd = 0& Then Exit Function

    If fgHwnd = consoleHwnd Then
        Console_IsForegroundActive = 1&
    End If
End Function

Private Function Console_IsHostShortcutBlocked(ByVal vkCode As Long, ByVal flags As Long) As Long
    Dim altDown As Long
    Dim ctrlDown As Long

    If Console_IsForegroundActive() = 0& Then Exit Function

    altDown = IIf(((flags And LLKHF_ALTDOWN) <> 0&) Or ((CLng(GetAsyncKeyState(vbKeyMenu)) And &H8000&) <> 0&), 1&, 0&)
    ctrlDown = IIf((CLng(GetAsyncKeyState(vbKeyControl)) And &H8000&) <> 0&, 1&, 0&)

    Select Case vkCode
        Case VK_LWIN, VK_RWIN, VK_APPS
            Console_IsHostShortcutBlocked = 1&

        Case vbKeyTab
            If altDown <> 0& Then Console_IsHostShortcutBlocked = 1&

        Case vbKeyEscape
            If (altDown <> 0&) Or (ctrlDown <> 0&) Then Console_IsHostShortcutBlocked = 1&

        Case vbKeySpace, vbKeyF4
            If altDown <> 0& Then Console_IsHostShortcutBlocked = 1&
    End Select
End Function

Public Function Console_KeyboardHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim kb As KBDLLHOOKSTRUCT

    If nCode = HC_ACTION Then
        Select Case wParam
            Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
                CopyMemory kb, ByVal lParam, LenB(kb)
                If Console_IsHostShortcutBlocked(kb.vkCode, kb.flags) <> 0& Then
                    Console_KeyboardHookProc = 1&
                    Exit Function
                End If
        End Select
    End If

    Console_KeyboardHookProc = CallNextHookEx(console_keyboardHook, nCode, wParam, lParam)
End Function

Private Sub Console_InstallWindowHook()
    If console_subclassHwnd <> 0& Then Exit Sub
    If console_isFormLoaded() = 0& Then Exit Sub
    If frmConsole.hwnd = 0& Then Exit Sub

    console_subclassHwnd = frmConsole.hwnd
    console_prevWndProc = SetWindowLong(console_subclassHwnd, GWL_WNDPROC, AddressOf Console_WindowProc)
    If console_prevWndProc = 0& Then
        console_subclassHwnd = 0&
    End If
End Sub

Private Sub Console_InstallKeyboardHook()
    If console_keyboardHook <> 0& Then Exit Sub

    console_keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf Console_KeyboardHookProc, App.hInstance, 0&)
End Sub

Private Sub Console_RemoveWindowHook()
    If (console_subclassHwnd = 0&) Or (console_prevWndProc = 0&) Then Exit Sub

    Call SetWindowLong(console_subclassHwnd, GWL_WNDPROC, console_prevWndProc)
    console_prevWndProc = 0&
    console_subclassHwnd = 0&
End Sub

Private Sub Console_RemoveKeyboardHook()
    If console_keyboardHook = 0& Then Exit Sub

    Call UnhookWindowsHookEx(console_keyboardHook)
    console_keyboardHook = 0&
End Sub

Private Function console_isFormLoaded() As Long
    Dim f As Form

    For Each f In Forms
        If StrComp(f.Name, "frmConsole", vbTextCompare) = 0& Then
            console_isFormLoaded = 1&
            Exit Function
        End If
    Next f

    console_isFormLoaded = 0&
End Function

Public Function console_init(ByVal title As String) As Long
    On Error GoTo InitErr

    console_title = title
    console_eventQRead = 0&
    console_eventQWrite = 0&
    console_pendingPixelsPtr = 0&
    console_pendingW = 0&
    console_pendingH = 0&
    console_pendingStride = 0&
    console_pendingReady = 0&
    console_curw = 0&
    console_curh = 0&
    console_surfaceW = 0&
    console_surfaceH = 0&
    console_grabbed = 0&
    console_ctrl = 0&
    console_alt = 0&
    console_ctrlLeft = 0&
    console_ctrlRight = 0&
    console_altLeft = 0&
    console_altRight = 0&
    console_doRepeat = 0&
    console_formActive = 1&
    console_cursorHidden = 0&
    console_cursorClipped = 0&
    console_prevMouseButtons = 0&
    console_rawMouseButtons = 0&
    console_suppressedMouseButtons = 0&
    console_ctrlF11Shortcut = 0&
    console_ctrlF12Shortcut = 0&
    console_ctrlAltDelTimer = TIMING_ERROR
    console_ctrlAltDelPos = 0&
    console_keyboardAcquired = 0&
    console_mouseAcquired = 0&

    If console_isFormLoaded() = 0& Then Load frmConsole

    frmConsole.Caption = title
    frmConsole.ScaleMode = vbPixels
    If frmConsole.Visible = 0& Then frmConsole.Show
    Call Console_InstallWindowHook

    Call Console_BuildDataFormats
    Call Console_InitDirectDraw
    Call Console_InitDirectInput

    If console_setWindow(640&, 400&) <> 0& Then
        console_init = -1&
        Exit Function
    End If

    console_scancodeSet = 1&
    console_keyTimer = timing_addTimer(TIMER_CB_CONSOLE_KEYREPEAT, 0&, 2#, TIMING_DISABLED)
    console_ctrlAltDelTimer = timing_addTimer(TIMER_CB_CONSOLE_CTRLALTDEL, 0&, 20#, TIMING_DISABLED)
    Call Console_InstallKeyboardHook
    Call Console_UpdateInputAcquireState
    console_init = 0&
    Exit Function

InitErr:
    Call Console_RemoveKeyboardHook
    Call Console_RemoveWindowHook
    console_init = -1&
End Function

Public Sub console_shutdown()
    Call Console_SetGrabbed(0&)
    Call Console_ReleaseAllKeys
    Call Console_ReleaseMouseButtons
    If console_ctrlAltDelTimer <> TIMING_ERROR Then timing_timerDisable console_ctrlAltDelTimer
    Call Console_RemoveKeyboardHook
    Call Console_RemoveWindowHook
    Call Console_ReleaseDirectInput
    Call Console_ReleaseDirectDraw
    console_formActive = 0&
    console_cursorClipped = 0&
    console_ctrlF11Shortcut = 0&
    console_ctrlF12Shortcut = 0&
    console_ctrlAltDelPos = 0&
    console_ctrl = 0&
    console_alt = 0&
    console_ctrlLeft = 0&
    console_ctrlRight = 0&
    console_altLeft = 0&
    console_altRight = 0&
End Sub

Public Sub console_changeScancodes(ByVal keyset As Byte)
    console_scancodeSet = keyset
End Sub

Public Function console_setWindow(ByVal w As Long, ByVal h As Long) As Long
    On Error GoTo SetWindowErr

    Dim frameW As Long
    Dim frameH As Long

    frmConsole.ScaleMode = vbPixels

    frameW = (frmConsole.Width \ Screen.TwipsPerPixelX) - frmConsole.ScaleWidth
    frameH = (frmConsole.Height \ Screen.TwipsPerPixelY) - frmConsole.ScaleHeight

    frmConsole.Width = (w + frameW) * Screen.TwipsPerPixelX
    frmConsole.Height = (h + frameH) * Screen.TwipsPerPixelY

    console_curw = w
    console_curh = h

    If console_useDirectDraw <> 0& Then
        Call Console_RecreateBackSurface(w, h)
    End If

    If (console_useDIMouse <> 0&) And (console_grabbed <> 0&) Then
        Call Console_UpdateCursorLock
    End If

    console_setWindow = 0&
    Exit Function

SetWindowErr:
    console_setWindow = -1&
End Function

Public Sub console_setTitle(ByVal suffix As String)
    On Error Resume Next
    console_titleStatus = suffix
    Call Console_RefreshWindowTitle
End Sub

Public Sub console_blit(ByVal pixelsPtr As Long, ByVal w As Long, ByVal h As Long, ByVal stride As Long)
    Static lasttime As Double

    Dim curtime As Double

    curtime = timing_getCur()

    If (w <> console_curw) Or (h <> console_curh) Then
        Call console_setWindow(w, h)
    End If

    If Console_DrawPixelsDirectDraw(pixelsPtr, w, h, stride) <> 0& Then
        Console_DrawPixelsGDI pixelsPtr, w, h, stride
    End If

    Console_UpdateTitleTiming curtime, lasttime
    lasttime = curtime
End Sub

Public Sub console_queueBlit(ByVal pixelsPtr As Long, ByVal w As Long, ByVal h As Long, ByVal stride As Long)
    If pixelsPtr = 0& Then Exit Sub

    console_pendingPixelsPtr = pixelsPtr
    console_pendingW = w
    console_pendingH = h
    console_pendingStride = stride
    InterlockedExchange console_pendingReady, 1&
End Sub

Public Sub console_presentPending()
    If InterlockedExchange(console_pendingReady, 0&) = 0& Then Exit Sub
    console_blit console_pendingPixelsPtr, console_pendingW, console_pendingH, console_pendingStride
End Sub

Public Sub console_mousegrab()
    If console_grabbed <> 0& Then
        Call Console_SetGrabbed(0&)
    Else
        Call Console_SetGrabbed(1&)
    End If
End Sub

Public Sub console_pump()
    Call Console_PumpDirectInput
End Sub

Public Function console_loop() As Long
    Dim ev As Byte
    Dim idx As Long

    If console_doRepeat <> 0& Then
        console_doRepeat = 0&
        console_curkey = console_lastKey
        console_loop = CONSOLE_EVENT_KEY
        Exit Function
    End If

    DoEvents
    Call Console_PumpDirectInput

    If console_eventQRead = console_eventQWrite Then
        console_loop = CONSOLE_EVENT_NONE
        Exit Function
    End If

    idx = console_eventQRead
    ev = console_eventQueue(idx)

    If ev = CONSOLE_EVENT_KEY Then
        console_curkey = console_eventData(idx)
    End If

    console_eventQRead = (console_eventQRead + 1&) And CONSOLE_EVENT_QUEUE_MASK
    console_loop = ev
End Function

Public Function console_getScancode() As Byte
    console_getScancode = console_curkey
End Function

Public Function console_translateScancode(ByVal keyval As Long) As Byte
    Select Case keyval
        Case vbKeyEscape: console_translateScancode = &H1&
        Case 49&: console_translateScancode = &H2&
        Case 50&: console_translateScancode = &H3&
        Case 51&: console_translateScancode = &H4&
        Case 52&: console_translateScancode = &H5&
        Case 53&: console_translateScancode = &H6&
        Case 54&: console_translateScancode = &H7&
        Case 55&: console_translateScancode = &H8&
        Case 56&: console_translateScancode = &H9&
        Case 57&: console_translateScancode = &HA&
        Case 48&: console_translateScancode = &HB&
        Case VK_OEM_MINUS: console_translateScancode = &HC&
        Case VK_OEM_PLUS: console_translateScancode = &HD&
        Case vbKeyBack: console_translateScancode = &HE&
        Case vbKeyTab: console_translateScancode = &HF&
        Case 81&: console_translateScancode = &H10&
        Case 87&: console_translateScancode = &H11&
        Case 69&: console_translateScancode = &H12&
        Case 82&: console_translateScancode = &H13&
        Case 84&: console_translateScancode = &H14&
        Case 89&: console_translateScancode = &H15&
        Case 85&: console_translateScancode = &H16&
        Case 73&: console_translateScancode = &H17&
        Case 79&: console_translateScancode = &H18&
        Case 80&: console_translateScancode = &H19&
        Case VK_OEM_4: console_translateScancode = &H1A&
        Case VK_OEM_6: console_translateScancode = &H1B&
        Case vbKeyReturn: console_translateScancode = &H1C&
        Case vbKeyControl: console_translateScancode = &H1D&
        Case 65&: console_translateScancode = &H1E&
        Case 83&: console_translateScancode = &H1F&
        Case 68&: console_translateScancode = &H20&
        Case 70&: console_translateScancode = &H21&
        Case 71&: console_translateScancode = &H22&
        Case 72&: console_translateScancode = &H23&
        Case 74&: console_translateScancode = &H24&
        Case 75&: console_translateScancode = &H25&
        Case 76&: console_translateScancode = &H26&
        Case VK_OEM_1: console_translateScancode = &H27&
        Case VK_OEM_7: console_translateScancode = &H28&
        Case VK_OEM_3: console_translateScancode = &H29&
        Case vbKeyShift: console_translateScancode = &H2A&
        Case VK_OEM_5: console_translateScancode = &H2B&
        Case 90&: console_translateScancode = &H2C&
        Case 88&: console_translateScancode = &H2D&
        Case 67&: console_translateScancode = &H2E&
        Case 86&: console_translateScancode = &H2F&
        Case 66&: console_translateScancode = &H30&
        Case 78&: console_translateScancode = &H31&
        Case 77&: console_translateScancode = &H32&
        Case VK_OEM_COMMA: console_translateScancode = &H33&
        Case VK_OEM_PERIOD: console_translateScancode = &H34&
        Case VK_OEM_2: console_translateScancode = &H35&
        Case vbKeyMultiply: console_translateScancode = &H37&
        Case vbKeyMenu: console_translateScancode = &H38&
        Case vbKeySpace: console_translateScancode = &H39&
        Case vbKeyCapital: console_translateScancode = &H3A&
        Case vbKeyF1: console_translateScancode = &H3B&
        Case vbKeyF2: console_translateScancode = &H3C&
        Case vbKeyF3: console_translateScancode = &H3D&
        Case vbKeyF4: console_translateScancode = &H3E&
        Case vbKeyF5: console_translateScancode = &H3F&
        Case vbKeyF6: console_translateScancode = &H40&
        Case vbKeyF7: console_translateScancode = &H41&
        Case vbKeyF8: console_translateScancode = &H42&
        Case vbKeyF9: console_translateScancode = &H43&
        Case vbKeyF10: console_translateScancode = &H44&
        Case vbKeyNumlock: console_translateScancode = &H45&
        Case vbKeyScrollLock: console_translateScancode = &H46&
        Case vbKeyNumpad7, vbKeyHome: console_translateScancode = &H47&
        Case vbKeyNumpad8, vbKeyUp: console_translateScancode = &H48&
        Case vbKeyNumpad9, vbKeyPageUp: console_translateScancode = &H49&
        Case vbKeySubtract: console_translateScancode = &H4A&
        Case vbKeyNumpad4, vbKeyLeft: console_translateScancode = &H4B&
        Case vbKeyNumpad5: console_translateScancode = &H4C&
        Case vbKeyNumpad6, vbKeyRight: console_translateScancode = &H4D&
        Case vbKeyAdd: console_translateScancode = &H4E&
        Case vbKeyNumpad1, vbKeyEnd: console_translateScancode = &H4F&
        Case vbKeyNumpad2, vbKeyDown: console_translateScancode = &H50&
        Case vbKeyNumpad3, vbKeyPageDown: console_translateScancode = &H51&
        Case vbKeyNumpad0, vbKeyInsert: console_translateScancode = &H52&
        Case vbKeyDecimal, vbKeyDelete: console_translateScancode = &H53&
        Case Else: console_translateScancode = &H0&
    End Select
End Function

Public Function console_translateScancodeSet2(ByVal keyval As Long, ByVal isBreak As Long) As Byte
    console_translateScancodeSet2 = console_translateScancode(keyval)
End Function

Public Sub console_keyRepeat(ByVal dummy As Long)
    console_doRepeat = 1&
    timing_updateIntervalFreq console_keyTimer, 15#
End Sub

Public Sub Console_FormKeyDown(ByVal keyCode As Integer, ByVal shift As Integer)
    Dim key As Byte
    Dim bytes() As Byte
    Dim keyLen As Long

    If console_useDIKeyboard <> 0& Then Exit Sub

    If (keyCode = vbKeyF11) And (((shift And vbCtrlMask) <> 0&) Or (console_ctrl <> 0&)) Then
        If console_ctrlF11Shortcut = 0& Then
            console_ctrlF11Shortcut = 1&
            Call console_mousegrab
        End If
        Exit Sub
    End If

    If (keyCode = vbKeyF12) And (((shift And vbCtrlMask) <> 0&) Or (console_ctrl <> 0&)) Then
        If console_ctrlF12Shortcut = 0& Then
            console_ctrlF12Shortcut = 1&
            Call Console_StartCtrlAltDelete
        End If
        Exit Sub
    End If

    If keyCode = vbKeyF11 Then keyCode = vbKeyControl
    If keyCode = vbKeyF12 Then keyCode = vbKeyMenu

    If keyCode = vbKeyControl Then console_ctrl = 1&
    If keyCode = vbKeyMenu Then console_alt = 1&

    If console_scancodeSet = 1& Then
        key = console_translateScancode(keyCode)
        If key = &H0& Then Exit Sub

        console_curkey = key
        console_lastKey = key
        timing_updateIntervalFreq console_keyTimer, 2#
        timing_timerEnable console_keyTimer
        Console_QueueEvent CONSOLE_EVENT_KEY, key
    ElseIf console_scancodeSet = 2& Then
        If (kbc.config And &H40&) <> 0& Then
            key = console_translateScancode(keyCode)
            If key = 0& Then Exit Sub

            ReDim bytes(0& To 0&) As Byte
            bytes(0&) = key
            i8042_buffer_key_data bytes, 1&, 1&
        Else
            keyLen = Console_BuildSet2Bytes(keyCode, 0&, bytes)
            If keyLen > 0& Then i8042_buffer_key_data bytes, keyLen, 1&
        End If
    End If
End Sub

Public Sub Console_FormKeyUp(ByVal keyCode As Integer, ByVal shift As Integer)
    Dim key As Byte
    Dim bytes() As Byte
    Dim keyLen As Long

    If console_useDIKeyboard <> 0& Then Exit Sub

    If keyCode = vbKeyF11 Then
        If console_ctrlF11Shortcut <> 0& Then
            console_ctrlF11Shortcut = 0&
            Exit Sub
        End If
    End If

    If keyCode = vbKeyF12 Then
        If console_ctrlF12Shortcut <> 0& Then
            console_ctrlF12Shortcut = 0&
            Exit Sub
        End If
    End If

    If keyCode = vbKeyF11 Then keyCode = vbKeyControl
    If keyCode = vbKeyF12 Then keyCode = vbKeyMenu

    If keyCode = vbKeyControl Then console_ctrl = 0&
    If keyCode = vbKeyMenu Then console_alt = 0&

    If console_scancodeSet = 1& Then
        key = console_translateScancode(keyCode)
        If key = 0& Then Exit Sub

        key = key Or &H80&

        If (key And &H7F&) = console_lastKey Then
            timing_timerDisable console_keyTimer
        End If

        Console_QueueEvent CONSOLE_EVENT_KEY, key
    Else
        If (kbc.config And &H40&) <> 0& Then
            key = console_translateScancode(keyCode)
            If key = 0& Then Exit Sub

            ReDim bytes(0& To 0&) As Byte
            bytes(0&) = key Or &H80&
            i8042_buffer_key_data bytes, 1&, 1&
        Else
            keyLen = Console_BuildSet2Bytes(keyCode, 1&, bytes)
            If keyLen > 0& Then i8042_buffer_key_data bytes, keyLen, 1&
        End If
    End If
End Sub

Public Sub Console_FormMouseMove(ByVal x As Single, ByVal y As Single)
    Dim xrel As Long
    Dim yrel As Long

    If console_useDIMouse <> 0& Then Exit Sub

    If console_grabbed = 0& Then
        console_lastMouseX = x
        console_lastMouseY = y
        Exit Sub
    End If

    xrel = CLng(x - console_lastMouseX)
    yrel = CLng(y - console_lastMouseY)

    If xrel < -128& Then xrel = -128&
    If xrel > 127& Then xrel = 127&
    If yrel < -128& Then yrel = -128&
    If yrel > 127& Then yrel = 127&

    If (xrel <> 0&) Or (yrel <> 0&) Then
        mouse_action MOUSE_ACTION_MOVE, MOUSE_NEITHER, xrel, yrel
    End If

    console_lastMouseX = x
    console_lastMouseY = y
End Sub

Public Sub Console_FormMouseDown(ByVal Button As Integer)
    Dim action As Byte

    If console_useDIMouse <> 0& Then Exit Sub
    If console_grabbed = 0& Then Exit Sub

    If Button = vbLeftButton Then
        action = MOUSE_ACTION_LEFT
    ElseIf Button = vbRightButton Then
        action = MOUSE_ACTION_RIGHT
    Else
        Exit Sub
    End If

    If console_grabbed <> 0& Then
        mouse_action action, MOUSE_PRESSED, 0&, 0&
    End If
End Sub

Public Sub Console_FormMouseUp(ByVal Button As Integer)
    Dim action As Byte

    If console_useDIMouse <> 0& Then Exit Sub

    If Button = vbLeftButton Then
        action = MOUSE_ACTION_LEFT
    ElseIf Button = vbRightButton Then
        action = MOUSE_ACTION_RIGHT
    Else
        Exit Sub
    End If

    If console_grabbed <> 0& Then
        mouse_action action, MOUSE_UNPRESSED, 0&, 0&
    End If
End Sub

Public Sub Console_FormActivate()
    console_formActive = 1&
    Call Console_UpdateInputAcquireState
End Sub

Public Sub Console_FormDeactivate()
    console_formActive = 0&
    Call Console_UpdateInputAcquireState
End Sub

Public Sub Console_FormUnload()
    Call console_shutdown
    Console_QueueEvent CONSOLE_EVENT_QUIT, 0&
End Sub

Private Sub Console_BuildDataFormats()
    Dim i As Long

    Call dxGuidXAxis(console_guidXAxis)
    Call dxGuidYAxis(console_guidYAxis)
    Call dxGuidZAxis(console_guidZAxis)

    dxZeroMemory VarPtr(console_keyboardFmt), LenB(console_keyboardFmt)
    dxZeroMemory VarPtr(console_mouseFmt), LenB(console_mouseFmt)
    dxZeroMemory VarPtr(console_keyboardObj(0&)), LenB(console_keyboardObj(0&)) * CONSOLE_DI_KEY_BYTES
    dxZeroMemory VarPtr(console_mouseObj(0&)), LenB(console_mouseObj(0&)) * CONSOLE_DI_MOUSE_OBJECTS

    For i = 0& To CONSOLE_DI_KEY_BYTES - 1&
        console_keyboardObj(i).pguid = 0&
        console_keyboardObj(i).dwOfs = i
        console_keyboardObj(i).dwType = DIDFT_BUTTON Or DIDFT_OPTIONAL Or Console_DIDFT_MAKEINSTANCE(i)
    Next i

    With console_keyboardFmt
        .dwSize = LenB(console_keyboardFmt)
        .dwObjSize = LenB(console_keyboardObj(0&))
        .dwFlags = 0&
        .dwDataSize = CONSOLE_DI_KEY_BYTES
        .dwNumObjs = CONSOLE_DI_KEY_BYTES
        .rgodf = VarPtr(console_keyboardObj(0&))
    End With

    console_mouseObj(0&).pguid = VarPtr(console_guidXAxis)
    console_mouseObj(0&).dwOfs = 0&
    console_mouseObj(0&).dwType = DIDFT_AXIS Or DIDFT_ANYINSTANCE
    console_mouseObj(1&).pguid = VarPtr(console_guidYAxis)
    console_mouseObj(1&).dwOfs = 4&
    console_mouseObj(1&).dwType = DIDFT_AXIS Or DIDFT_ANYINSTANCE
    console_mouseObj(2&).pguid = VarPtr(console_guidZAxis)
    console_mouseObj(2&).dwOfs = 8&
    console_mouseObj(2&).dwType = DIDFT_OPTIONAL Or DIDFT_AXIS Or DIDFT_ANYINSTANCE
    console_mouseObj(3&).pguid = 0&
    console_mouseObj(3&).dwOfs = 12&
    console_mouseObj(3&).dwType = DIDFT_BUTTON Or DIDFT_ANYINSTANCE
    console_mouseObj(4&).pguid = 0&
    console_mouseObj(4&).dwOfs = 13&
    console_mouseObj(4&).dwType = DIDFT_BUTTON Or DIDFT_ANYINSTANCE
    console_mouseObj(5&).pguid = 0&
    console_mouseObj(5&).dwOfs = 14&
    console_mouseObj(5&).dwType = DIDFT_OPTIONAL Or DIDFT_BUTTON Or DIDFT_ANYINSTANCE
    console_mouseObj(6&).pguid = 0&
    console_mouseObj(6&).dwOfs = 15&
    console_mouseObj(6&).dwType = DIDFT_OPTIONAL Or DIDFT_BUTTON Or DIDFT_ANYINSTANCE

    With console_mouseFmt
        .dwSize = LenB(console_mouseFmt)
        .dwObjSize = LenB(console_mouseObj(0&))
        .dwFlags = DIDF_RELAXIS
        .dwDataSize = 16&
        .dwNumObjs = CONSOLE_DI_MOUSE_OBJECTS
        .rgodf = VarPtr(console_mouseObj(0&))
    End With

    dxZeroMemory VarPtr(console_keyboardState(0&)), CONSOLE_DI_KEY_BYTES
    dxZeroMemory VarPtr(console_prevKeyboardState(0&)), CONSOLE_DI_KEY_BYTES
End Sub

Private Sub Console_InitDirectDraw()
    Dim hr As Long
    Dim desc As DDSURFACEDESC

    console_useDirectDraw = 0&

    hr = DirectDrawCreate(0&, console_dd, 0&)
    If dxHrFailed(hr) Or (console_dd = 0&) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectDraw unavailable; using GDI blits"
        Call Console_ReleaseDirectDraw
        Exit Sub
    End If

    hr = dxCallLong(console_dd, IDX_IDIRECTDRAW_SETCOOPERATIVELEVEL, frmConsole.hWnd, DDSCL_NORMAL)
    If dxHrFailed(hr) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectDraw cooperative level failed; using GDI blits"
        Call Console_ReleaseDirectDraw
        Exit Sub
    End If

    dxZeroMemory VarPtr(desc), LenB(desc)
    desc.dwSize = LenB(desc)
    desc.dwFlags = DDSD_CAPS
    desc.ddsCaps.dwCaps = DDSCAPS_PRIMARYSURFACE

    hr = dxCallLong(console_dd, IDX_IDIRECTDRAW_CREATESURFACE, VarPtr(desc), VarPtr(console_primarySurface), 0&)
    If dxHrFailed(hr) Or (console_primarySurface = 0&) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectDraw primary surface creation failed; using GDI blits"
        Call Console_ReleaseDirectDraw
        Exit Sub
    End If

    hr = DirectDrawCreateClipper(0&, console_clipper, 0&)
    If dxHrFailed(hr) Or (console_clipper = 0&) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectDraw clipper creation failed; using GDI blits"
        Call Console_ReleaseDirectDraw
        Exit Sub
    End If

    hr = dxCallLong(console_clipper, IDX_IDIRECTDRAWCLIPPER_SETHWND, 0&, frmConsole.hWnd)
    If dxHrFailed(hr) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectDraw clipper setup failed; using GDI blits"
        Call Console_ReleaseDirectDraw
        Exit Sub
    End If

    hr = dxCallLong(console_primarySurface, IDX_IDIRECTDRAWSURFACE_SETCLIPPER, console_clipper)
    If dxHrFailed(hr) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectDraw clipper attach failed; using GDI blits"
        Call Console_ReleaseDirectDraw
        Exit Sub
    End If

    console_useDirectDraw = 1&
End Sub

Private Sub Console_ReleaseDirectDraw()
    Call dxRelease(console_backSurface)
    Call dxRelease(console_clipper)
    Call dxRelease(console_primarySurface)
    Call dxRelease(console_dd)
    console_useDirectDraw = 0&
    console_surfaceW = 0&
    console_surfaceH = 0&
End Sub

Private Function Console_RecreateBackSurface(ByVal w As Long, ByVal h As Long) As Long
    Dim hr As Long
    Dim desc As DDSURFACEDESC

    Console_RecreateBackSurface = -1&

    If console_useDirectDraw = 0& Then Exit Function

    If (console_backSurface <> 0&) And (console_surfaceW = w) And (console_surfaceH = h) Then
        Console_RecreateBackSurface = 0&
        Exit Function
    End If

    Call dxRelease(console_backSurface)
    console_surfaceW = 0&
    console_surfaceH = 0&

    dxZeroMemory VarPtr(desc), LenB(desc)
    desc.dwSize = LenB(desc)
    desc.dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_PIXELFORMAT
    desc.dwWidth = w
    desc.dwHeight = h
    desc.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    desc.ddpfPixelFormat.dwSize = LenB(desc.ddpfPixelFormat)
    desc.ddpfPixelFormat.dwFlags = DDPF_RGB
    desc.ddpfPixelFormat.dwRGBBitCount = 32&
    desc.ddpfPixelFormat.dwRBitMask = &HFF0000
    desc.ddpfPixelFormat.dwGBitMask = &HFF00&
    desc.ddpfPixelFormat.dwBBitMask = &HFF&

    hr = dxCallLong(console_dd, IDX_IDIRECTDRAW_CREATESURFACE, VarPtr(desc), VarPtr(console_backSurface), 0&)
    If dxHrFailed(hr) Or (console_backSurface = 0&) Then Exit Function

    console_surfaceW = w
    console_surfaceH = h
    Console_RecreateBackSurface = 0&
End Function

Private Sub Console_InitDirectInput()
    Dim hr As Long

    console_useDIKeyboard = 0&
    console_useDIMouse = 0&

    hr = DirectInputCreateA(App.hInstance, DIRECTINPUT_VERSION, console_di, 0&)
    If dxHrFailed(hr) Or (console_di = 0&) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput unavailable; using form input"
        Call Console_ReleaseDirectInput
        Exit Sub
    End If

    Call Console_InitDIKeyboard
    Call Console_InitDIMouse

    If (console_useDIKeyboard = 0&) And (console_useDIMouse = 0&) Then
        Call Console_ReleaseDirectInput
    End If
End Sub

Private Sub Console_InitDIKeyboard()
    Dim hr As Long
    Dim devGuid As GUID
    Dim devPtr As Long

    Call dxGuidSysKeyboard(devGuid)
    hr = dxCallLong(console_di, IDX_IDIRECTINPUT_CREATEDEVICE, VarPtr(devGuid), VarPtr(devPtr), 0&)
    If dxHrFailed(hr) Or (devPtr = 0&) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput keyboard init failed; using form keyboard"
        Exit Sub
    End If

    hr = dxCallLong(devPtr, IDX_IDIRECTINPUTDEVICE_SETDATAFORMAT, VarPtr(console_keyboardFmt))
    If dxHrFailed(hr) Then
        Call dxRelease(devPtr)
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput keyboard format failed; using form keyboard"
        Exit Sub
    End If

    hr = dxCallLong(devPtr, IDX_IDIRECTINPUTDEVICE_SETCOOPERATIVELEVEL, frmConsole.hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND)
    If dxHrFailed(hr) Then
        Call dxRelease(devPtr)
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput keyboard coop level failed; using form keyboard"
        Exit Sub
    End If

    console_keyboard = devPtr
    console_useDIKeyboard = 1&
    debug_log DEBUG_INFO, "[CONSOLE] DirectInput keyboard enabled"
End Sub

Private Sub Console_InitDIMouse()
    Dim hr As Long
    Dim devGuid As GUID
    Dim devPtr As Long

    Call dxGuidSysMouse(devGuid)
    hr = dxCallLong(console_di, IDX_IDIRECTINPUT_CREATEDEVICE, VarPtr(devGuid), VarPtr(devPtr), 0&)
    If dxHrFailed(hr) Or (devPtr = 0&) Then
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput mouse init failed; using form mouse"
        Exit Sub
    End If

    hr = dxCallLong(devPtr, IDX_IDIRECTINPUTDEVICE_SETDATAFORMAT, VarPtr(console_mouseFmt))
    If dxHrFailed(hr) Then
        Call dxRelease(devPtr)
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput mouse format failed; using form mouse"
        Exit Sub
    End If

    hr = dxCallLong(devPtr, IDX_IDIRECTINPUTDEVICE_SETCOOPERATIVELEVEL, frmConsole.hWnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND)
    If dxHrFailed(hr) Then
        Call dxRelease(devPtr)
        debug_log DEBUG_INFO, "[CONSOLE] DirectInput mouse coop level failed; using form mouse"
        Exit Sub
    End If

    console_mouse = devPtr
    console_useDIMouse = 1&
    console_useBufferedDIMouse = 0&
End Sub

Private Sub Console_ReleaseDirectInput()
    Call Console_ReleaseKeyboard
    Call Console_ReleaseMouse
    Call dxRelease(console_keyboard)
    Call dxRelease(console_mouse)
    Call dxRelease(console_di)
    console_useDIKeyboard = 0&
    console_useDIMouse = 0&
    console_useBufferedDIMouse = 0&
End Sub

Private Sub Console_UpdateInputAcquireState()
    If console_useDIKeyboard <> 0& Then
        If console_formActive <> 0& Then
            Call Console_AcquireKeyboard
        Else
            Call Console_ReleaseKeyboard
        End If
    End If

    If console_useDIMouse <> 0& Then
        If (console_formActive <> 0&) And (console_grabbed <> 0&) Then
            Call Console_AcquireMouse
        Else
            Call Console_ReleaseMouse
        End If
    End If

    Call Console_UpdateCursorLock
End Sub

Private Sub Console_AcquireKeyboard()
    Dim hr As Long

    If (console_useDIKeyboard = 0&) Or (console_keyboard = 0&) Or (console_keyboardAcquired <> 0&) Then Exit Sub

    hr = dxCallLong(console_keyboard, IDX_IDIRECTINPUTDEVICE_ACQUIRE)
    If dxHrFailed(hr) Then Exit Sub

    console_keyboardAcquired = 1&
End Sub

Private Sub Console_ReleaseKeyboard()
    If console_keyboard = 0& Then Exit Sub

    If console_keyboardAcquired <> 0& Then
        Call Console_ReleaseAllKeys
        Call dxCallLong(console_keyboard, IDX_IDIRECTINPUTDEVICE_UNACQUIRE)
    End If

    console_keyboardAcquired = 0&
End Sub

Private Sub Console_AcquireMouse()
    Dim hr As Long
    Dim state As DIMOUSESTATE

    If (console_useDIMouse = 0&) Or (console_mouse = 0&) Or (console_mouseAcquired <> 0&) Then Exit Sub

    hr = dxCallLong(console_mouse, IDX_IDIRECTINPUTDEVICE_ACQUIRE)
    If dxHrFailed(hr) Then Exit Sub

    console_mouseAcquired = 1&
    Call Console_FlushBufferedMouseEvents
    console_prevMouseButtons = 0&
    console_rawMouseButtons = 0&
    dxZeroMemory VarPtr(state), LenB(state)
    hr = dxCallLong(console_mouse, IDX_IDIRECTINPUTDEVICE_GETDEVICESTATE, LenB(state), VarPtr(state))
    If dxHrSucceeded(hr) Then
        console_suppressedMouseButtons = console_suppressedMouseButtons Or Console_GetRawMouseButtons(state)
        console_rawMouseButtons = Console_GetRawMouseButtons(state)
    End If
End Sub

Private Sub Console_ReleaseMouse()
    If console_mouse = 0& Then Exit Sub

    If console_mouseAcquired <> 0& Then
        Call Console_ReleaseMouseButtons
        Call dxCallLong(console_mouse, IDX_IDIRECTINPUTDEVICE_UNACQUIRE)
    End If

    console_mouseAcquired = 0&
    console_rawMouseButtons = 0&
    console_suppressedMouseButtons = 0&
End Sub

Private Sub Console_SetGrabbed(ByVal grabbed As Byte)
    If grabbed <> 0& Then
        console_grabbed = 1&
    Else
        console_grabbed = 0&
        Call Console_ReleaseMouseButtons
    End If

    Call Console_UpdateInputAcquireState
    Call Console_RefreshWindowTitle

    If console_useDIMouse = 0& Then
        If grabbed <> 0& Then
            Call Console_SetCursorVisible(0&)
        Else
            Call Console_SetCursorVisible(1&)
        End If
    End If
End Sub

Private Sub Console_SetCursorVisible(ByVal visible As Byte)
    Dim i As Long
    Dim showRet As Long

    If visible <> 0& Then
        If console_cursorHidden = 0& Then Exit Sub
        For i = 0& To 15&
            showRet = ShowCursor(1&)
            If showRet >= 0& Then Exit For
        Next i
        console_cursorHidden = 0&
    Else
        If console_cursorHidden <> 0& Then Exit Sub
        For i = 0& To 15&
            showRet = ShowCursor(0&)
            If showRet < 0& Then Exit For
        Next i
        console_cursorHidden = 1&
    End If
End Sub

Private Sub Console_UpdateCursorLock()
    Dim centerPt As POINTAPI
    Dim clipRect As RECT

    If console_useDIMouse = 0& Then
        If console_cursorClipped <> 0& Then
            ClipCursor 0&
            console_cursorClipped = 0&
        End If
        Exit Sub
    End If

    If (console_grabbed <> 0&) And (console_formActive <> 0&) Then
        If Console_GetClientCenterScreenPoint(centerPt) = 0& Then Exit Sub

        clipRect.Left = centerPt.x
        clipRect.Top = centerPt.y
        clipRect.Right = centerPt.x + 1&
        clipRect.Bottom = centerPt.y + 1&
        ClipCursor VarPtr(clipRect)
        SetCursorPos centerPt.x, centerPt.y
        console_cursorClipped = 1&
        Call Console_SetCursorVisible(0&)
    Else
        If console_cursorClipped <> 0& Then
            ClipCursor 0&
            console_cursorClipped = 0&
        End If
        Call Console_SetCursorVisible(1&)
    End If
End Sub

Private Sub Console_PumpDirectInput()
    If (console_useDIKeyboard <> 0&) And (console_formActive <> 0&) And (console_keyboardAcquired = 0&) Then
        Call Console_AcquireKeyboard
    End If

    If (console_useDIMouse <> 0&) And (console_formActive <> 0&) And (console_grabbed <> 0&) And (console_mouseAcquired = 0&) Then
        Call Console_AcquireMouse
    End If

    If console_useDIKeyboard <> 0& Then Call Console_PollKeyboard
    If console_useDIMouse <> 0& Then Call Console_PollMouse
End Sub

Private Sub Console_PollKeyboard()
    Dim hr As Long
    Dim i As Long

    If console_keyboard = 0& Then Exit Sub

    If console_keyboardAcquired = 0& Then
        If console_formActive <> 0& Then Call Console_AcquireKeyboard
        Exit Sub
    End If

    hr = dxCallLong(console_keyboard, IDX_IDIRECTINPUTDEVICE_GETDEVICESTATE, CONSOLE_DI_KEY_BYTES, VarPtr(console_keyboardState(0&)))
    If dxHrFailed(hr) Then
        Call Console_ReleaseAllKeys
        console_keyboardAcquired = 0&
        If console_formActive <> 0& Then Call Console_AcquireKeyboard
        Exit Sub
    End If

    For i = 0& To CONSOLE_DI_KEY_BYTES - 1&
        If (console_keyboardState(i) And &H80&) <> (console_prevKeyboardState(i) And &H80&) Then
            Call Console_HandleDIKeyChange(i, ((console_keyboardState(i) And &H80&) <> 0&))
            console_prevKeyboardState(i) = console_keyboardState(i)
        End If
    Next i
End Sub

Private Sub Console_PollMouse()
    Dim hr As Long
    Dim state As DIMOUSESTATE
    Dim objData As DIDEVICEOBJECTDATA
    Dim elemCount As Long
    Dim eventCount As Long
    Dim xrel As Long
    Dim yrel As Long
    Dim btnState As Byte

    If (console_mouse = 0&) Or (console_grabbed = 0&) Then Exit Sub

    If console_mouseAcquired = 0& Then
        If console_formActive <> 0& Then Call Console_AcquireMouse
        Exit Sub
    End If

    If console_useBufferedDIMouse <> 0& Then
        Do
            elemCount = 1&
            dxZeroMemory VarPtr(objData), LenB(objData)
            hr = dxCallLong(console_mouse, IDX_IDIRECTINPUTDEVICE_GETDEVICEDATA, LenB(objData), VarPtr(objData), VarPtr(elemCount), 0&)
            If dxHrFailed(hr) Then
                Call Console_ReleaseMouseButtons
                console_mouseAcquired = 0&
                If console_formActive <> 0& Then Call Console_AcquireMouse
                Exit Sub
            End If

            If elemCount = 0& Then Exit Do
            eventCount = eventCount + 1&

            Select Case objData.dwOfs
                Case 0&
                    xrel = xrel + objData.dwData
                Case 4&
                    yrel = yrel + objData.dwData
            End Select
            If eventCount >= CONSOLE_DI_MOUSE_EVENT_BURST Then Exit Do
        Loop
    End If

    dxZeroMemory VarPtr(state), LenB(state)
    hr = dxCallLong(console_mouse, IDX_IDIRECTINPUTDEVICE_GETDEVICESTATE, LenB(state), VarPtr(state))
    If dxHrFailed(hr) Then
        Call Console_ReleaseMouseButtons
        console_mouseAcquired = 0&
        If console_formActive <> 0& Then Call Console_AcquireMouse
        Exit Sub
    End If

    If console_useBufferedDIMouse = 0& Then
        xrel = state.lX
        yrel = state.lY
    End If

    btnState = Console_GetFilteredMouseButtons(state)
    Call mouse_syncButtons(IIf((btnState And 1&) <> 0&, 1&, 0&), IIf((btnState And 2&) <> 0&, 1&, 0&))

    If xrel < -128& Then xrel = -128&
    If xrel > 127& Then xrel = 127&
    If yrel < -128& Then yrel = -128&
    If yrel > 127& Then yrel = 127&

    If (xrel <> 0&) Or (yrel <> 0&) Then
        mouse_action MOUSE_ACTION_MOVE, MOUSE_NEITHER, xrel, yrel
    End If

    console_prevMouseButtons = btnState
End Sub

Private Sub Console_HandleDIKeyChange(ByVal dik As Long, ByVal isDown As Long)
    Dim mappedKey As Byte
    Dim bytes() As Byte
    Dim keyLen As Long

    If dik = DIK_F11 Then
        If isDown <> 0& Then
            If console_ctrl <> 0& Then
                If console_ctrlF11Shortcut = 0& Then
                    console_ctrlF11Shortcut = 1&
                    Call console_mousegrab
                End If
                Exit Sub
            End If
        ElseIf console_ctrlF11Shortcut <> 0& Then
            console_ctrlF11Shortcut = 0&
            Exit Sub
        End If
    End If

    If dik = DIK_F12 Then
        If isDown <> 0& Then
            If console_ctrl <> 0& Then
                If console_ctrlF12Shortcut = 0& Then
                    console_ctrlF12Shortcut = 1&
                    Call Console_StartCtrlAltDelete
                End If
                Exit Sub
            End If
        ElseIf console_ctrlF12Shortcut <> 0& Then
            console_ctrlF12Shortcut = 0&
            Exit Sub
        End If
    End If

    dik = Console_NormalizeDIK(dik)
    If dik = 0& Then Exit Sub

    Call Console_UpdateDIModifierState(dik, isDown)

    If console_scancodeSet = 1& Then
        mappedKey = Console_MapDIKToLegacySet1(dik)
        If mappedKey = 0& Then Exit Sub

        If isDown <> 0& Then
            console_curkey = mappedKey
            console_lastKey = mappedKey
            timing_updateIntervalFreq console_keyTimer, 2#
            timing_timerEnable console_keyTimer
            Console_QueueEvent CONSOLE_EVENT_KEY, mappedKey
        Else
            mappedKey = mappedKey Or &H80&
            If (mappedKey And &H7F&) = console_lastKey Then
                timing_timerDisable console_keyTimer
            End If
            Console_QueueEvent CONSOLE_EVENT_KEY, mappedKey
        End If

        Exit Sub
    End If

    If (kbc.config And &H40&) <> 0& Then
        mappedKey = Console_MapDIKToLegacySet1(dik)
        If mappedKey = 0& Then Exit Sub

        ReDim bytes(0& To 0&) As Byte
        bytes(0&) = mappedKey
        If isDown = 0& Then bytes(0&) = bytes(0&) Or &H80&
        i8042_buffer_key_data bytes, 1&, 1&
        Exit Sub
    End If

    keyLen = Console_BuildSet2BytesFromDIK(dik, (isDown = 0&), bytes)
    If keyLen > 0& Then i8042_buffer_key_data bytes, keyLen, 1&
End Sub

Private Sub Console_ReleaseAllKeys()
    Dim i As Long

    If console_useDIKeyboard = 0& Then Exit Sub

    For i = 0& To CONSOLE_DI_KEY_BYTES - 1&
        If (console_prevKeyboardState(i) And &H80&) <> 0& Then
            Call Console_HandleDIKeyChange(i, 0&)
            console_prevKeyboardState(i) = 0&
            console_keyboardState(i) = 0&
        End If
    Next i

    timing_timerDisable console_keyTimer
    console_ctrlF11Shortcut = 0&
    console_ctrlF12Shortcut = 0&
    console_ctrl = 0&
    console_alt = 0&
    console_ctrlLeft = 0&
    console_ctrlRight = 0&
    console_altLeft = 0&
    console_altRight = 0&
End Sub

Private Sub Console_ReleaseMouseButtons()
    mouse_syncButtons 0&, 0&

    console_prevMouseButtons = 0&
    console_rawMouseButtons = 0&
    console_suppressedMouseButtons = 0&
End Sub

Private Sub Console_FlushBufferedMouseEvents()
    Dim objData As DIDEVICEOBJECTDATA
    Dim elemCount As Long
    Dim flushCount As Long
    Dim hr As Long

    If console_useBufferedDIMouse = 0& Then Exit Sub

    Do
        elemCount = 1&
        dxZeroMemory VarPtr(objData), LenB(objData)
        hr = dxCallLong(console_mouse, IDX_IDIRECTINPUTDEVICE_GETDEVICEDATA, LenB(objData), VarPtr(objData), VarPtr(elemCount), 0&)
        If dxHrFailed(hr) Or (elemCount = 0&) Then Exit Do
        flushCount = flushCount + 1&
        If flushCount >= (CONSOLE_DI_MOUSE_EVENT_BURST * 4&) Then Exit Do
    Loop
End Sub

Private Function Console_FilterMouseButtonState(ByVal rawBtnState As Byte) As Byte
    console_suppressedMouseButtons = console_suppressedMouseButtons And rawBtnState
    Console_FilterMouseButtonState = rawBtnState And (&HFF& Xor console_suppressedMouseButtons)
End Function

Private Function Console_GetRawMouseButtons(ByRef state As DIMOUSESTATE) As Byte
    Dim rawBtnState As Byte

    rawBtnState = 0&
    If (state.rgbButtons(0&) And &H80&) <> 0& Then rawBtnState = rawBtnState Or 1&
    If (state.rgbButtons(1&) And &H80&) <> 0& Then rawBtnState = rawBtnState Or 2&

    Console_GetRawMouseButtons = rawBtnState
End Function

Private Function Console_GetFilteredMouseButtons(ByRef state As DIMOUSESTATE) As Byte
    Dim rawBtnState As Byte

    rawBtnState = Console_GetRawMouseButtons(state)
    console_rawMouseButtons = rawBtnState
    Console_GetFilteredMouseButtons = Console_FilterMouseButtonState(rawBtnState)
End Function

Private Sub Console_StartCtrlAltDelete()
    If console_ctrlAltDelTimer = TIMING_ERROR Then Exit Sub
    If console_ctrlAltDelPos <> 0& Then Exit Sub

    console_ctrlAltDelPos = 0&
    timing_timerEnable console_ctrlAltDelTimer
End Sub

Public Sub console_ctrlAltDelStep(ByVal dummy As Long)
    Dim keyCode As Long
    Dim isBreak As Long

    Select Case console_ctrlAltDelPos
        Case 0&
            keyCode = vbKeyControl
            isBreak = 0&
        Case 1&
            keyCode = vbKeyMenu
            isBreak = 0&
        Case 2&
            keyCode = vbKeyDelete
            isBreak = 0&
        Case 3&
            keyCode = vbKeyDelete
            isBreak = 1&
        Case 4&
            keyCode = vbKeyMenu
            isBreak = 1&
        Case 5&
            keyCode = vbKeyControl
            isBreak = 1&
        Case Else
            timing_timerDisable console_ctrlAltDelTimer
            console_ctrlAltDelPos = 0&
            Exit Sub
    End Select

    Call Console_SendSyntheticKey(keyCode, isBreak)
    console_ctrlAltDelPos = console_ctrlAltDelPos + 1&

    If console_ctrlAltDelPos >= CONSOLE_CTRLALTDEL_STEPS Then
        timing_timerDisable console_ctrlAltDelTimer
        console_ctrlAltDelPos = 0&
    End If
End Sub

Private Sub Console_SendSyntheticKey(ByVal keyCode As Long, ByVal isBreak As Long)
    Dim key As Byte
    Dim bytes() As Byte
    Dim keyLen As Long

    If console_scancodeSet = 1& Then
        key = console_translateScancode(keyCode)
        If key = 0& Then Exit Sub
        If isBreak <> 0& Then key = key Or &H80&
        Console_QueueEvent CONSOLE_EVENT_KEY, key
        Exit Sub
    End If

    If (kbc.config And &H40&) <> 0& Then
        key = console_translateScancode(keyCode)
        If key = 0& Then Exit Sub

        ReDim bytes(0& To 0&) As Byte
        bytes(0&) = key
        If isBreak <> 0& Then bytes(0&) = bytes(0&) Or &H80&
        i8042_buffer_key_data bytes, 1&, 1&
        Exit Sub
    End If

    keyLen = Console_BuildSet2Bytes(keyCode, isBreak, bytes)
    If keyLen > 0& Then i8042_buffer_key_data bytes, keyLen, 1&
End Sub

Private Function Console_NormalizeDIK(ByVal dik As Long) As Long
    Select Case dik
        Case DIK_F11
            Console_NormalizeDIK = &H1D&
        Case DIK_F12
            Console_NormalizeDIK = &H38&
        Case DIK_SYSRQ, DIK_APPS
            Console_NormalizeDIK = 0&
        Case Else
            Console_NormalizeDIK = dik
    End Select
End Function

Private Sub Console_UpdateDIModifierState(ByVal dik As Long, ByVal isDown As Long)
    Select Case dik
        Case DIK_LCONTROL
            console_ctrlLeft = CByte(isDown <> 0&)
        Case DIK_RCONTROL
            console_ctrlRight = CByte(isDown <> 0&)
        Case DIK_LMENU
            console_altLeft = CByte(isDown <> 0&)
        Case DIK_RMENU
            console_altRight = CByte(isDown <> 0&)
        Case Else
            Exit Sub
    End Select

    console_ctrl = CByte((console_ctrlLeft <> 0&) Or (console_ctrlRight <> 0&))
    console_alt = CByte((console_altLeft <> 0&) Or (console_altRight <> 0&))
End Sub

Private Function Console_MapDIKToLegacySet1(ByVal dik As Long) As Byte
    Select Case dik
        Case DIK_DIVIDE
            Console_MapDIKToLegacySet1 = 0&
        Case DIK_NUMPADENTER
            Console_MapDIKToLegacySet1 = &H1C&
        Case Else
            If (dik And &H80&) <> 0& Then
                Console_MapDIKToLegacySet1 = CByte(dik And &H7F&)
            ElseIf (dik >= 0&) And (dik <= &H7F&) Then
                Console_MapDIKToLegacySet1 = CByte(dik)
            Else
                Console_MapDIKToLegacySet1 = 0&
            End If
    End Select
End Function

Private Function Console_BuildSet2BytesFromDIK(ByVal dik As Long, ByVal isBreak As Long, ByRef outBytes() As Byte) As Long
    Dim set1Code As Byte
    Dim set2Code As Byte
    Dim isExtended As Byte

    If Console_MapDIKToSet1Base(dik, set1Code, isExtended) = 0& Then
        Console_BuildSet2BytesFromDIK = 0&
        Exit Function
    End If

    If Console_MapSet1ToSet2(set1Code, set2Code) = 0& Then
        Console_BuildSet2BytesFromDIK = 0&
        Exit Function
    End If

    If isExtended = 0& Then
        If isBreak = 0& Then
            ReDim outBytes(0& To 0&) As Byte
            outBytes(0&) = set2Code
            Console_BuildSet2BytesFromDIK = 1&
        Else
            ReDim outBytes(0& To 1&) As Byte
            outBytes(0&) = &HF0&
            outBytes(1&) = set2Code
            Console_BuildSet2BytesFromDIK = 2&
        End If
    Else
        If isBreak = 0& Then
            ReDim outBytes(0& To 1&) As Byte
            outBytes(0&) = &HE0&
            outBytes(1&) = set2Code
            Console_BuildSet2BytesFromDIK = 2&
        Else
            ReDim outBytes(0& To 2&) As Byte
            outBytes(0&) = &HE0&
            outBytes(1&) = &HF0&
            outBytes(2&) = set2Code
            Console_BuildSet2BytesFromDIK = 3&
        End If
    End If
End Function

Private Function Console_MapDIKToSet1Base(ByVal dik As Long, ByRef set1Code As Byte, ByRef isExtended As Byte) As Long
    Select Case dik
        Case 0&
            Console_MapDIKToSet1Base = 0&
            Exit Function
        Case DIK_NUMPADENTER
            set1Code = &H1C&
            isExtended = 1&
        Case DIK_DIVIDE
            set1Code = &H35&
            isExtended = 1&
        Case Else
            If (dik And &H80&) <> 0& Then
                set1Code = CByte(dik And &H7F&)
                isExtended = 1&
            ElseIf (dik >= 0&) And (dik <= &H7F&) Then
                set1Code = CByte(dik)
                isExtended = 0&
            Else
                Console_MapDIKToSet1Base = 0&
                Exit Function
            End If
    End Select

    Console_MapDIKToSet1Base = 1&
End Function

Private Function Console_MapSet1ToSet2(ByVal set1Code As Byte, ByRef scancode As Byte) As Long
    Select Case set1Code
        Case &H1&: scancode = &H76&
        Case &H2&: scancode = &H16&
        Case &H3&: scancode = &H1E&
        Case &H4&: scancode = &H26&
        Case &H5&: scancode = &H25&
        Case &H6&: scancode = &H2E&
        Case &H7&: scancode = &H36&
        Case &H8&: scancode = &H3D&
        Case &H9&: scancode = &H3E&
        Case &HA&: scancode = &H46&
        Case &HB&: scancode = &H45&
        Case &HC&: scancode = &H4E&
        Case &HD&: scancode = &H55&
        Case &HE&: scancode = &H66&
        Case &HF&: scancode = &HD&
        Case &H10&: scancode = &H15&
        Case &H11&: scancode = &H1D&
        Case &H12&: scancode = &H24&
        Case &H13&: scancode = &H2D&
        Case &H14&: scancode = &H2C&
        Case &H15&: scancode = &H35&
        Case &H16&: scancode = &H3C&
        Case &H17&: scancode = &H43&
        Case &H18&: scancode = &H44&
        Case &H19&: scancode = &H4D&
        Case &H1A&: scancode = &H54&
        Case &H1B&: scancode = &H5B&
        Case &H1C&: scancode = &H5A&
        Case &H1D&: scancode = &H14&
        Case &H1E&: scancode = &H1C&
        Case &H1F&: scancode = &H1B&
        Case &H20&: scancode = &H23&
        Case &H21&: scancode = &H2B&
        Case &H22&: scancode = &H34&
        Case &H23&: scancode = &H33&
        Case &H24&: scancode = &H3B&
        Case &H25&: scancode = &H42&
        Case &H26&: scancode = &H4B&
        Case &H27&: scancode = &H4C&
        Case &H28&: scancode = &H52&
        Case &H29&: scancode = &HE&
        Case &H2A&: scancode = &H12&
        Case &H2B&: scancode = &H5D&
        Case &H2C&: scancode = &H1A&
        Case &H2D&: scancode = &H22&
        Case &H2E&: scancode = &H21&
        Case &H2F&: scancode = &H2A&
        Case &H30&: scancode = &H32&
        Case &H31&: scancode = &H31&
        Case &H32&: scancode = &H3A&
        Case &H33&: scancode = &H41&
        Case &H34&: scancode = &H49&
        Case &H35&: scancode = &H4A&
        Case &H37&: scancode = &H7C&
        Case &H38&: scancode = &H11&
        Case &H39&: scancode = &H29&
        Case &H3A&: scancode = &H58&
        Case &H3B&: scancode = &H5&
        Case &H3C&: scancode = &H6&
        Case &H3D&: scancode = &H4&
        Case &H3E&: scancode = &HC&
        Case &H3F&: scancode = &H3&
        Case &H40&: scancode = &HB&
        Case &H41&: scancode = &H83&
        Case &H42&: scancode = &HA&
        Case &H43&: scancode = &H1&
        Case &H44&: scancode = &H9&
        Case &H45&: scancode = &H77&
        Case &H46&: scancode = &H7E&
        Case &H47&: scancode = &H6C&
        Case &H48&: scancode = &H75&
        Case &H49&: scancode = &H7D&
        Case &H4A&: scancode = &H7B&
        Case &H4B&: scancode = &H6B&
        Case &H4C&: scancode = &H73&
        Case &H4D&: scancode = &H74&
        Case &H4E&: scancode = &H79&
        Case &H4F&: scancode = &H69&
        Case &H50&: scancode = &H72&
        Case &H51&: scancode = &H7A&
        Case &H52&: scancode = &H70&
        Case &H53&: scancode = &H71&
        Case Else
            Console_MapSet1ToSet2 = 0&
            Exit Function
    End Select

    Console_MapSet1ToSet2 = 1&
End Function

Private Function Console_DrawPixelsDirectDraw(ByVal pixelsPtr As Long, ByVal w As Long, ByVal h As Long, ByVal stride As Long) As Long
    Dim bmi As BITMAPINFO
    Dim srcRect As RECT
    Dim dstRect As RECT
    Dim hr As Long
    Dim dibResult As Long
    Dim surfHdc As Long

    Console_DrawPixelsDirectDraw = -1&

    If (console_useDirectDraw = 0&) Or (pixelsPtr = 0&) Then Exit Function
    If Console_RecreateBackSurface(w, h) <> 0& Then Exit Function

    hr = dxCallLong(console_backSurface, IDX_IDIRECTDRAWSURFACE_GETDC, VarPtr(surfHdc))
    If dxHrFailed(hr) Then
        Call Console_RestoreSurfaces
        hr = dxCallLong(console_backSurface, IDX_IDIRECTDRAWSURFACE_GETDC, VarPtr(surfHdc))
        If dxHrFailed(hr) Then Exit Function
    End If

    Call Console_InitBitmapInfo(bmi, w, h, stride)

    dibResult = StretchDIBits(surfHdc, 0&, 0&, w, h, 0&, 0&, w, h, pixelsPtr, bmi, DIB_RGB_COLORS, SRCCOPY)
    Call dxCallLong(console_backSurface, IDX_IDIRECTDRAWSURFACE_RELEASEDC, surfHdc)
    If (dibResult = 0&) Or (dibResult = GDI_ERROR) Then Exit Function

    If Console_GetClientScreenRect(dstRect) = 0& Then Exit Function

    srcRect.Left = 0&
    srcRect.Top = 0&
    srcRect.Right = w
    srcRect.Bottom = h

    hr = dxCallLong(console_primarySurface, IDX_IDIRECTDRAWSURFACE_BLT, VarPtr(dstRect), console_backSurface, VarPtr(srcRect), DDBLT_WAIT, 0&)
    If dxHrFailed(hr) Then
        Call Console_RestoreSurfaces
        hr = dxCallLong(console_primarySurface, IDX_IDIRECTDRAWSURFACE_BLT, VarPtr(dstRect), console_backSurface, VarPtr(srcRect), DDBLT_WAIT, 0&)
        If dxHrFailed(hr) Then Exit Function
    End If

    Console_DrawPixelsDirectDraw = 0&
End Function

Private Sub Console_DrawPixelsGDI(ByVal pixelsPtr As Long, ByVal w As Long, ByVal h As Long, ByVal stride As Long)
    Dim bmi As BITMAPINFO

    If pixelsPtr = 0& Then Exit Sub

    Call Console_InitBitmapInfo(bmi, w, h, stride)

    StretchDIBits frmConsole.hdc, 0&, 0&, w, h, 0&, 0&, w, h, pixelsPtr, bmi, DIB_RGB_COLORS, SRCCOPY
    frmConsole.Refresh
End Sub

Private Sub Console_InitBitmapInfo(ByRef bmi As BITMAPINFO, ByVal w As Long, ByVal h As Long, ByVal stride As Long)
    Dim srcW As Long

    dxZeroMemory VarPtr(bmi), LenB(bmi)
    srcW = Console_SourceWidth(w, stride)

    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = srcW
        .biHeight = -h
        .biPlanes = 1&
        .biBitCount = 32&
        .biCompression = BI_RGB
        If stride > 0& Then
            .biSizeImage = stride * h
        Else
            .biSizeImage = w * 4& * h
        End If
    End With
End Sub

Private Function Console_SourceWidth(ByVal w As Long, ByVal stride As Long) As Long
    If stride > 0& Then
        Console_SourceWidth = stride \ 4&
    Else
        Console_SourceWidth = w
    End If

    If Console_SourceWidth < w Then
        Console_SourceWidth = w
    End If
End Function

Private Function Console_GetClientScreenRect(ByRef rc As RECT) As Long
    Dim pt As POINTAPI

    If GetClientRect(frmConsole.hWnd, rc) = 0& Then Exit Function

    pt.x = rc.Left
    pt.y = rc.Top
    If ClientToScreen(frmConsole.hWnd, pt) = 0& Then Exit Function
    rc.Left = pt.x
    rc.Top = pt.y

    pt.x = rc.Right
    pt.y = rc.Bottom
    If ClientToScreen(frmConsole.hWnd, pt) = 0& Then Exit Function
    rc.Right = pt.x
    rc.Bottom = pt.y

    Console_GetClientScreenRect = 1&
End Function

Private Function Console_GetClientCenterScreenPoint(ByRef pt As POINTAPI) As Long
    Dim rc As RECT

    If GetClientRect(frmConsole.hWnd, rc) = 0& Then Exit Function

    pt.x = (rc.Left + rc.Right) \ 2&
    pt.y = (rc.Top + rc.Bottom) \ 2&
    If ClientToScreen(frmConsole.hWnd, pt) = 0& Then Exit Function

    Console_GetClientCenterScreenPoint = 1&
End Function

Private Sub Console_RestoreSurfaces()
    If console_primarySurface <> 0& Then Call dxCallLong(console_primarySurface, IDX_IDIRECTDRAWSURFACE_RESTORE)
    If console_backSurface <> 0& Then Call dxCallLong(console_backSurface, IDX_IDIRECTDRAWSURFACE_RESTORE)
End Sub

Private Sub Console_UpdateTitleTiming(ByVal curtime As Double, ByRef lasttime As Double)
    Dim i As Long
    Dim avgcount As Long
    Dim curavg As Double
    Dim titleSuffix As String
    Dim ipsText As String

    If lasttime = 0# Then Exit Sub

    console_frameTime(console_frameIdx) = curtime - lasttime
    console_frameIdx = console_frameIdx + 1&

    If console_frameIdx <> 30& Then Exit Sub

    console_frameIdx = 0&

    For i = 0& To 29&
        If console_frameTime(i) <> 0# Then
            curavg = curavg + console_frameTime(i)
            avgcount = avgcount + 1&
        End If
    Next i

    If (avgcount <= 0&) Or (curavg <= 0#) Then Exit Sub

    curavg = curavg / avgcount
    titleSuffix = Format$((timing_getFreq() / curavg), "0.00") & " FPS"
    ipsText = diag_getIPSString()
    If LenB(ipsText) <> 0& Then titleSuffix = titleSuffix & " / " & ipsText & " IPS"
    console_setTitle titleSuffix
End Sub

Private Sub Console_RefreshWindowTitle()
    Dim captionText As String

    On Error Resume Next

    captionText = console_title
    If LenB(console_titleStatus) <> 0& Then
        captionText = captionText & " - " & console_titleStatus
    End If
    If console_grabbed <> 0& Then
        captionText = captionText & " [Ctrl-F11 releases mouse grab]"
    End If

    frmConsole.Caption = captionText
End Sub

Private Sub Console_QueueEvent(ByVal eventType As Byte, ByVal scancode As Byte)
    Dim nextPos As Long

    nextPos = (console_eventQWrite + 1&) And CONSOLE_EVENT_QUEUE_MASK
    If nextPos = console_eventQRead Then Exit Sub

    console_eventQueue(console_eventQWrite) = eventType
    console_eventData(console_eventQWrite) = scancode
    console_eventQWrite = nextPos
End Sub

Private Function Console_DIDFT_MAKEINSTANCE(ByVal n As Long) As Long
    Console_DIDFT_MAKEINSTANCE = (n And &HFF&) * &H100&
End Function

Private Function Console_BuildSet2Bytes(ByVal keyCode As Long, ByVal isBreak As Long, ByRef outBytes() As Byte) As Long
    Dim scancode As Byte
    Dim extended As Byte

    If Console_MapSet2Code(keyCode, scancode, extended) = 0& Then
        Console_BuildSet2Bytes = 0&
        Exit Function
    End If

    If extended = 0& Then
        If isBreak = 0& Then
            ReDim outBytes(0& To 0&) As Byte
            outBytes(0&) = scancode
            Console_BuildSet2Bytes = 1&
        Else
            ReDim outBytes(0& To 1&) As Byte
            outBytes(0&) = &HF0&
            outBytes(1&) = scancode
            Console_BuildSet2Bytes = 2&
        End If
    Else
        If isBreak = 0& Then
            ReDim outBytes(0& To 1&) As Byte
            outBytes(0&) = &HE0&
            outBytes(1&) = scancode
            Console_BuildSet2Bytes = 2&
        Else
            ReDim outBytes(0& To 2&) As Byte
            outBytes(0&) = &HE0&
            outBytes(1&) = &HF0&
            outBytes(2&) = scancode
            Console_BuildSet2Bytes = 3&
        End If
    End If
End Function

Private Function Console_MapSet2Code(ByVal keyCode As Long, ByRef scancode As Byte, ByRef extended As Byte) As Long
    extended = 0&

    Select Case keyCode
        Case 65&: scancode = &H1C&
        Case 66&: scancode = &H32&
        Case 67&: scancode = &H21&
        Case 68&: scancode = &H23&
        Case 69&: scancode = &H24&
        Case 70&: scancode = &H2B&
        Case 71&: scancode = &H34&
        Case 72&: scancode = &H33&
        Case 73&: scancode = &H43&
        Case 74&: scancode = &H3B&
        Case 75&: scancode = &H42&
        Case 76&: scancode = &H4B&
        Case 77&: scancode = &H3A&
        Case 78&: scancode = &H31&
        Case 79&: scancode = &H44&
        Case 80&: scancode = &H4D&
        Case 81&: scancode = &H15&
        Case 82&: scancode = &H2D&
        Case 83&: scancode = &H1B&
        Case 84&: scancode = &H2C&
        Case 85&: scancode = &H3C&
        Case 86&: scancode = &H2A&
        Case 87&: scancode = &H1D&
        Case 88&: scancode = &H22&
        Case 89&: scancode = &H35&
        Case 90&: scancode = &H1A&
        Case 48&: scancode = &H45&
        Case 49&: scancode = &H16&
        Case 50&: scancode = &H1E&
        Case 51&: scancode = &H26&
        Case 52&: scancode = &H25&
        Case 53&: scancode = &H2E&
        Case 54&: scancode = &H36&
        Case 55&: scancode = &H3D&
        Case 56&: scancode = &H3E&
        Case 57&: scancode = &H46&
        Case VK_OEM_3: scancode = &HE&
        Case VK_OEM_MINUS: scancode = &H4E&
        Case VK_OEM_PLUS: scancode = &H55&
        Case VK_OEM_5: scancode = &H5D&
        Case vbKeyBack: scancode = &H66&
        Case vbKeySpace: scancode = &H29&
        Case vbKeyTab: scancode = &HD&
        Case vbKeyCapital: scancode = &H58&
        Case vbKeyShift: scancode = &H12&
        Case vbKeyControl: scancode = &H14&
        Case vbKeyMenu: scancode = &H11&
        Case vbKeyReturn: scancode = &H5A&
        Case vbKeyEscape: scancode = &H76&
        Case vbKeyF1: scancode = &H5&
        Case vbKeyF2: scancode = &H6&
        Case vbKeyF3: scancode = &H4&
        Case vbKeyF4: scancode = &HC&
        Case vbKeyF5: scancode = &H3&
        Case vbKeyF6: scancode = &HB&
        Case vbKeyF7: scancode = &H83&
        Case vbKeyF8: scancode = &HA&
        Case vbKeyF9: scancode = &H1&
        Case vbKeyF10: scancode = &H9&
        Case vbKeyF11: scancode = &H78&
        Case vbKeyF12: scancode = &H7&
        Case VK_OEM_4: scancode = &H54&
        Case VK_OEM_6: scancode = &H5B&
        Case VK_OEM_1: scancode = &H4C&
        Case VK_OEM_7: scancode = &H52&
        Case VK_OEM_COMMA: scancode = &H41&
        Case VK_OEM_PERIOD: scancode = &H49&
        Case VK_OEM_2: scancode = &H4A&
        Case vbKeyInsert: scancode = &H70&: extended = 1&
        Case vbKeyDelete: scancode = &H71&: extended = 1&
        Case vbKeyHome: scancode = &H6C&: extended = 1&
        Case vbKeyEnd: scancode = &H69&: extended = 1&
        Case vbKeyPageUp: scancode = &H7D&: extended = 1&
        Case vbKeyPageDown: scancode = &H7A&: extended = 1&
        Case vbKeyUp: scancode = &H75&: extended = 1&
        Case vbKeyLeft: scancode = &H6B&: extended = 1&
        Case vbKeyDown: scancode = &H72&: extended = 1&
        Case vbKeyRight: scancode = &H74&: extended = 1&
        Case vbKeyNumlock: scancode = &H77&
        Case vbKeyDivide: scancode = &H4A&: extended = 1&
        Case vbKeyMultiply: scancode = &H7C&
        Case vbKeySubtract: scancode = &H7B&
        Case vbKeyAdd: scancode = &H79&
        Case vbKeyDecimal: scancode = &H71&
        Case vbKeyNumpad0: scancode = &H70&
        Case vbKeyNumpad1: scancode = &H69&
        Case vbKeyNumpad2: scancode = &H72&
        Case vbKeyNumpad3: scancode = &H7A&
        Case vbKeyNumpad4: scancode = &H6B&
        Case vbKeyNumpad5: scancode = &H73&
        Case vbKeyNumpad6: scancode = &H74&
        Case vbKeyNumpad7: scancode = &H6C&
        Case vbKeyNumpad8: scancode = &H75&
        Case vbKeyNumpad9: scancode = &H7D&
        Case Else
            Console_MapSet2Code = 0&
            Exit Function
    End Select

    Console_MapSet2Code = 1&
End Function
