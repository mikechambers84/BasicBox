Attribute VB_Name = "modUInt32"
Option Explicit

Private Const U32_SIGN_BIT As Long = &H80000000
Private Const U32_LOW31_MASK As Long = &H7FFFFFFF
Private Const U16_MASK As Long = &HFFFF&
Private Const U16_SIGN_BIT As Long = &H8000&
Private Const U16_MOD As Long = &H10000

Private u32_pow2Init As Byte
Private u32_pow2(0& To 30&) As Long
Private u32_maskLow(0& To 30&) As Long

Private Const U32_MOD_D As Double = 4294967296#
Private Const U31_MOD_D As Double = 2147483648#

Private Sub U32_InitIntTables()
    Dim i As Long

    If u32_pow2Init <> 0& Then Exit Sub

    u32_pow2(0&) = 1&
    For i = 1& To 30&
        u32_pow2(i) = (u32_pow2(i - 1&) + u32_pow2(i - 1&))
    Next i

    u32_maskLow(0&) = 0&
    For i = 1& To 30&
        u32_maskLow(i) = (u32_pow2(i) - 1&)
    Next i

    u32_pow2Init = 1&
End Sub

Private Function U32_Hi16(ByVal v As Long) As Long
    U32_Hi16 = ((v And &H7FFF0000) \ U16_MOD)
    If v < 0& Then
        U32_Hi16 = (U32_Hi16 Or U16_SIGN_BIT)
    End If
End Function

Private Function U32_Combine16(ByVal hiPart As Long, ByVal loPart As Long) As Long
    Dim low31 As Long

    low31 = (((hiPart And &H7FFF&) * U16_MOD) Or (loPart And U16_MASK))
    If (hiPart And U16_SIGN_BIT) <> 0& Then
        U32_Combine16 = (low31 Or U32_SIGN_BIT)
    Else
        U32_Combine16 = low31
    End If
End Function

Public Function U32ToDouble(ByVal v As Long) As Double
    If v < 0& Then
        U32ToDouble = (CDbl(v And U32_LOW31_MASK) + U31_MOD_D)
    Else
        U32ToDouble = CDbl(v)
    End If
End Function

Public Function U32FromDouble(ByVal d As Double) As Long
    d = d - Fix(d / U32_MOD_D) * U32_MOD_D
    If d < 0# Then
        d = d + U32_MOD_D
    End If

    If d >= U31_MOD_D Then
        U32FromDouble = CLng(d - U32_MOD_D)
    Else
        U32FromDouble = CLng(d)
    End If
End Function

Public Function U32Shr(ByVal v As Long, ByVal bits As Long) As Long
    Dim divisor As Long

    If bits <= 0& Then
        U32Shr = v
        Exit Function
    End If

    If bits >= 32& Then
        U32Shr = 0&
        Exit Function
    End If

    If bits = 31& Then
        If v < 0& Then
            U32Shr = 1&
        Else
            U32Shr = 0&
        End If
        Exit Function
    End If

    U32_InitIntTables
    divisor = u32_pow2(bits)

    If v >= 0& Then
        U32Shr = (v \ divisor)
    Else
        U32Shr = ((v And U32_LOW31_MASK) \ divisor)
        U32Shr = (U32Shr + u32_pow2(31& - bits))
    End If
End Function

Public Function U32Shl(ByVal v As Long, ByVal bits As Long) As Long
    Dim sourceLowBits As Long
    Dim low31 As Long

    If bits <= 0& Then
        U32Shl = v
        Exit Function
    End If

    If bits >= 32& Then
        U32Shl = 0&
        Exit Function
    End If

    U32_InitIntTables

    If bits = 31& Then
        If (v And 1&) <> 0& Then
            U32Shl = U32_SIGN_BIT
        Else
            U32Shl = 0&
        End If
        Exit Function
    End If

    sourceLowBits = (31& - bits)
    low31 = ((v And u32_maskLow(sourceLowBits)) * u32_pow2(bits))

    If (v And u32_pow2(sourceLowBits)) <> 0& Then
        U32Shl = (low31 Or U32_SIGN_BIT)
    Else
        U32Shl = low31
    End If
End Function

Public Function U32Add(ByVal a As Long, ByVal b As Long) As Long
    Dim loSum As Long
    Dim hiSum As Long
    Dim loPart As Long
    Dim hiPart As Long
    Dim carry As Long

    loSum = ((a And U16_MASK) + (b And U16_MASK))
    carry = (loSum \ U16_MOD)
    loPart = (loSum And U16_MASK)

    hiSum = (U32_Hi16(a) + U32_Hi16(b) + carry)
    hiPart = (hiSum And U16_MASK)

    U32Add = U32_Combine16(hiPart, loPart)
End Function

Public Function U32Sub(ByVal a As Long, ByVal b As Long) As Long
    Dim aLo As Long
    Dim bLo As Long
    Dim aHi As Long
    Dim bHi As Long
    Dim loPart As Long
    Dim hiPart As Long
    Dim borrow As Long

    aLo = (a And U16_MASK)
    bLo = (b And U16_MASK)
    aHi = U32_Hi16(a)
    bHi = U32_Hi16(b)

    If aLo >= bLo Then
        loPart = (aLo - bLo)
        borrow = 0&
    Else
        loPart = ((aLo + U16_MOD) - bLo)
        borrow = 1&
    End If

    If aHi >= (bHi + borrow) Then
        hiPart = (aHi - (bHi + borrow))
    Else
        hiPart = ((aHi + U16_MOD) - (bHi + borrow))
    End If

    U32Sub = U32_Combine16((hiPart And U16_MASK), loPart)
End Function

Public Function U32Lt(ByVal a As Long, ByVal b As Long) As Long
    Dim aHi As Long
    Dim bHi As Long
    Dim aLo As Long
    Dim bLo As Long

    aHi = U32_Hi16(a)
    bHi = U32_Hi16(b)

    If aHi < bHi Then
        U32Lt = 1&
        Exit Function
    End If

    If aHi > bHi Then
        U32Lt = 0&
        Exit Function
    End If

    aLo = (a And U16_MASK)
    bLo = (b And U16_MASK)
    If aLo < bLo Then
        U32Lt = 1&
    Else
        U32Lt = 0&
    End If
End Function
