Attribute VB_Name = "modUtility"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME_t) As Long

Private Const OFN_HIDEREADONLY As Long = &H4&
Private Const OFN_NOCHANGEDIR As Long = &H8&
Private Const OFN_PATHMUSTEXIST As Long = &H800&
Private Const OFN_FILEMUSTEXIST As Long = &H1000&
Private Const OFN_EXPLORER As Long = &H80000

Private Type OPENFILENAME_t
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function utility_loadFile(ByRef dst() As Byte, ByVal length As Long, ByVal srcfile As String) As Long
    Dim h As Integer
    Dim fLen As Long

    On Error GoTo LoadErr

    If length <= 0& Then
        utility_loadFile = -1&
        Exit Function
    End If

    If Dir$(srcfile, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = vbNullString Then
        Erase dst
        utility_loadFile = -1&
        Exit Function
    End If

    h = FreeFile
    Open srcfile For Binary Access Read As #h
    fLen = LOF(h)
    If fLen < length Then
        Close #h
        Erase dst
        utility_loadFile = -1&
        Exit Function
    End If

    If (LBound(dst) <> 0&) Or ((UBound(dst) - LBound(dst) + 1&) < length) Then
        ReDim dst(0& To length - 1&) As Byte
    End If

    Get #h, 1&, dst
    Close #h

    utility_loadFile = 0&
    Exit Function

LoadErr:
    On Error Resume Next
    If h <> 0& Then Close #h
    Erase dst
    utility_loadFile = -1&
End Function

Public Sub utility_sleep(ByVal ms As Long)
    If ms <= 0& Then Exit Sub
    Sleep ms
End Sub

Public Function utility_openFileDialog(ByVal title As String, ByVal filterSpec As String, ByVal ownerHwnd As Long) As String
    Dim dialog As OPENFILENAME_t
    Dim fileBuf As String
    Dim fileTitleBuf As String
    Dim nulPos As Long

    fileBuf = String$(1024&, vbNullChar)
    fileTitleBuf = String$(260&, vbNullChar)

    dialog.lStructSize = Len(dialog)
    dialog.hwndOwner = ownerHwnd
    dialog.lpstrFilter = Replace$(filterSpec, "|", vbNullChar) & vbNullChar & vbNullChar
    dialog.nFilterIndex = 1&
    dialog.lpstrFile = fileBuf
    dialog.nMaxFile = Len(fileBuf)
    dialog.lpstrFileTitle = fileTitleBuf
    dialog.nMaxFileTitle = Len(fileTitleBuf)
    dialog.lpstrTitle = title
    dialog.Flags = (OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR)

    If GetOpenFileName(dialog) = 0& Then Exit Function

    nulPos = InStr(1&, dialog.lpstrFile, vbNullChar)
    If nulPos <= 1& Then Exit Function

    utility_openFileDialog = Left$(dialog.lpstrFile, nulPos - 1&)
End Function
