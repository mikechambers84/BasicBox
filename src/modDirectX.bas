Attribute VB_Name = "modDirectX"
Option Explicit

Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByVal pvargResult As Long) As Long

Public Declare Function DirectDrawCreate Lib "ddraw.dll" (ByVal lpGuid As Long, ByRef lplpDD As Long, ByVal pUnkOuter As Long) As Long
Public Declare Function DirectDrawCreateClipper Lib "ddraw.dll" (ByVal dwFlags As Long, ByRef lplpDDClipper As Long, ByVal pUnkOuter As Long) As Long
Public Declare Function DirectInputCreateA Lib "dinput.dll" (ByVal hInst As Long, ByVal dwVersion As Long, ByRef ppDI As Long, ByVal pUnkOuter As Long) As Long
Public Declare Function DirectSoundCreate Lib "dsound.dll" (ByVal lpGuid As Long, ByRef ppDS As Long, ByVal pUnkOuter As Long) As Long

Public Declare Sub dxCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal cbCopy As Long)
Public Declare Sub dxZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal Destination As Long, ByVal Length As Long)

Private Const CC_STDCALL As Long = 4&
Private Const VT_EMPTY As Integer = 0
Private Const VT_I4 As Integer = 3

Public Const DIRECTINPUT_VERSION As Long = &H300&

Public Const IDX_IUNKNOWN_QUERYINTERFACE As Long = 0&
Public Const IDX_IUNKNOWN_ADDREF As Long = 1&
Public Const IDX_IUNKNOWN_RELEASE As Long = 2&

Public Const IDX_IDIRECTDRAW_SETCOOPERATIVELEVEL As Long = 20&
Public Const IDX_IDIRECTDRAW_CREATESURFACE As Long = 6&

Public Const IDX_IDIRECTDRAWCLIPPER_SETHWND As Long = 8&

Public Const IDX_IDIRECTDRAWSURFACE_BLT As Long = 5&
Public Const IDX_IDIRECTDRAWSURFACE_GETDC As Long = 17&
Public Const IDX_IDIRECTDRAWSURFACE_LOCK As Long = 25&
Public Const IDX_IDIRECTDRAWSURFACE_RELEASEDC As Long = 26&
Public Const IDX_IDIRECTDRAWSURFACE_RESTORE As Long = 27&
Public Const IDX_IDIRECTDRAWSURFACE_SETCLIPPER As Long = 28&
Public Const IDX_IDIRECTDRAWSURFACE_UNLOCK As Long = 32&

Public Const IDX_IDIRECTINPUT_CREATEDEVICE As Long = 3&

Public Const IDX_IDIRECTINPUTDEVICE_ACQUIRE As Long = 7&
Public Const IDX_IDIRECTINPUTDEVICE_UNACQUIRE As Long = 8&
Public Const IDX_IDIRECTINPUTDEVICE_GETDEVICESTATE As Long = 9&
Public Const IDX_IDIRECTINPUTDEVICE_GETDEVICEDATA As Long = 10&
Public Const IDX_IDIRECTINPUTDEVICE_SETPROPERTY As Long = 6&
Public Const IDX_IDIRECTINPUTDEVICE_SETDATAFORMAT As Long = 11&
Public Const IDX_IDIRECTINPUTDEVICE_SETCOOPERATIVELEVEL As Long = 13&

Public Const IDX_IDIRECTSOUND_CREATESOUNDBUFFER As Long = 3&
Public Const IDX_IDIRECTSOUND_SETCOOPERATIVELEVEL As Long = 6&

Public Const IDX_IDIRECTSOUNDBUFFER_GETCURRENTPOSITION As Long = 4&
Public Const IDX_IDIRECTSOUNDBUFFER_GETSTATUS As Long = 9&
Public Const IDX_IDIRECTSOUNDBUFFER_LOCK As Long = 11&
Public Const IDX_IDIRECTSOUNDBUFFER_PLAY As Long = 12&
Public Const IDX_IDIRECTSOUNDBUFFER_STOP As Long = 18&
Public Const IDX_IDIRECTSOUNDBUFFER_UNLOCK As Long = 19&
Public Const IDX_IDIRECTSOUNDBUFFER_RESTORE As Long = 20&

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0& To 7&) As Byte
End Type

Public Function dxHrFailed(ByVal hr As Long) As Long
    dxHrFailed = (hr < 0&)
End Function

Public Function dxHrSucceeded(ByVal hr As Long) As Long
    dxHrSucceeded = (hr >= 0&)
End Function

Public Sub dxGuidSysMouse(ByRef guid As GUID)
    guid.Data1 = &H6F1D2B60
    guid.Data2 = &HD5A0
    guid.Data3 = &H11CF
    guid.Data4(0&) = &HBF
    guid.Data4(1&) = &HC7
    guid.Data4(2&) = &H44
    guid.Data4(3&) = &H45
    guid.Data4(4&) = &H53
    guid.Data4(5&) = &H54
    guid.Data4(6&) = &H0
    guid.Data4(7&) = &H0
End Sub

Public Sub dxGuidSysKeyboard(ByRef guid As GUID)
    guid.Data1 = &H6F1D2B61
    guid.Data2 = &HD5A0
    guid.Data3 = &H11CF
    guid.Data4(0&) = &HBF
    guid.Data4(1&) = &HC7
    guid.Data4(2&) = &H44
    guid.Data4(3&) = &H45
    guid.Data4(4&) = &H53
    guid.Data4(5&) = &H54
    guid.Data4(6&) = &H0
    guid.Data4(7&) = &H0
End Sub

Public Sub dxGuidXAxis(ByRef guid As GUID)
    guid.Data1 = &HA36D02E0
    guid.Data2 = &HC9F3
    guid.Data3 = &H11CF
    guid.Data4(0&) = &HBF
    guid.Data4(1&) = &HC7
    guid.Data4(2&) = &H44
    guid.Data4(3&) = &H45
    guid.Data4(4&) = &H53
    guid.Data4(5&) = &H54
    guid.Data4(6&) = &H0
    guid.Data4(7&) = &H0
End Sub

Public Sub dxGuidYAxis(ByRef guid As GUID)
    guid.Data1 = &HA36D02E1
    guid.Data2 = &HC9F3
    guid.Data3 = &H11CF
    guid.Data4(0&) = &HBF
    guid.Data4(1&) = &HC7
    guid.Data4(2&) = &H44
    guid.Data4(3&) = &H45
    guid.Data4(4&) = &H53
    guid.Data4(5&) = &H54
    guid.Data4(6&) = &H0
    guid.Data4(7&) = &H0
End Sub

Public Sub dxGuidZAxis(ByRef guid As GUID)
    guid.Data1 = &HA36D02E2
    guid.Data2 = &HC9F3
    guid.Data3 = &H11CF
    guid.Data4(0&) = &HBF
    guid.Data4(1&) = &HC7
    guid.Data4(2&) = &H44
    guid.Data4(3&) = &H45
    guid.Data4(4&) = &H53
    guid.Data4(5&) = &H54
    guid.Data4(6&) = &H0
    guid.Data4(7&) = &H0
End Sub

Public Sub dxGuidButton(ByRef guid As GUID)
    guid.Data1 = &HA36D02F0
    guid.Data2 = &HC9F3
    guid.Data3 = &H11CF
    guid.Data4(0&) = &HBF
    guid.Data4(1&) = &HC7
    guid.Data4(2&) = &H44
    guid.Data4(3&) = &H45
    guid.Data4(4&) = &H53
    guid.Data4(5&) = &H54
    guid.Data4(6&) = &H0
    guid.Data4(7&) = &H0
End Sub

Private Function dxInvoke(ByVal pvInstance As Long, ByVal methodIndex As Long, ByVal vtReturn As Integer, ByVal argCount As Long, ByRef argValues() As Long, ByRef retValue As Variant) As Long
    Dim hr As Long
    Dim argTypes() As Integer
    Dim argVars() As Variant
    Dim argPtrs() As Long
    Dim i As Long

    If pvInstance = 0& Then
        dxInvoke = -1&
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

        hr = DispCallFunc(pvInstance, methodIndex * 4&, CC_STDCALL, vtReturn, argCount, VarPtr(argTypes(0&)), VarPtr(argPtrs(0&)), VarPtr(retValue))
    Else
        hr = DispCallFunc(pvInstance, methodIndex * 4&, CC_STDCALL, vtReturn, 0&, 0&, 0&, VarPtr(retValue))
    End If

    dxInvoke = hr
End Function

Public Function dxCallLong(ByVal pvInstance As Long, ByVal methodIndex As Long, ParamArray args() As Variant) As Long
    Dim argCount As Long
    Dim argValues() As Long
    Dim retVar As Variant
    Dim callHr As Long
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
    callHr = dxInvoke(pvInstance, methodIndex, VT_I4, argCount, argValues, retVar)
    If callHr <> 0& Then
        dxCallLong = callHr
    ElseIf IsEmpty(retVar) Then
        dxCallLong = 0&
    Else
        dxCallLong = CLng(retVar)
    End If
End Function

Public Sub dxCallVoid(ByVal pvInstance As Long, ByVal methodIndex As Long, ParamArray args() As Variant)
    Dim argCount As Long
    Dim argValues() As Long
    Dim retVar As Variant
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
    Call dxInvoke(pvInstance, methodIndex, VT_EMPTY, argCount, argValues, retVar)
End Sub

Public Sub dxRelease(ByRef pvInstance As Long)
    If pvInstance = 0& Then Exit Sub
    Call dxCallLong(pvInstance, IDX_IUNKNOWN_RELEASE)
    pvInstance = 0&
End Sub
