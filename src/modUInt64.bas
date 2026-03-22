Attribute VB_Name = "modUInt64"
Option Explicit

Public Type U64_t
    Lo As Long
    Hi As Long
End Type

Public Function U64_FromParts(ByVal loPart As Long, ByVal hiPart As Long) As U64_t
    U64_FromParts.Lo = loPart
    U64_FromParts.Hi = hiPart
End Function

Public Function U64_Zero() As U64_t
    U64_Zero.Lo = 0&
    U64_Zero.Hi = 0&
End Function

Public Function U64_FromU32(ByVal v As Long) As U64_t
    U64_FromU32.Lo = v
    U64_FromU32.Hi = 0&
End Function

Public Function U64_Add(ByRef a As U64_t, ByRef b As U64_t) As U64_t
    Dim sumLo As Long
    Dim carry As Long

    sumLo = U32Add(a.Lo, b.Lo)
    carry = U32Lt(sumLo, a.Lo)

    U64_Add.Lo = sumLo
    U64_Add.Hi = U32Add(U32Add(a.Hi, b.Hi), carry)
End Function

Public Function U64_Sub(ByRef a As U64_t, ByRef b As U64_t) As U64_t
    Dim diffLo As Long
    Dim borrow As Long

    diffLo = U32Sub(a.Lo, b.Lo)
    borrow = U32Lt(a.Lo, b.Lo)

    U64_Sub.Lo = diffLo
    U64_Sub.Hi = U32Sub(U32Sub(a.Hi, b.Hi), borrow)
End Function

Public Function U64_Shl(ByRef v As U64_t, ByVal bits As Long) As U64_t
    Dim carry As Long

    bits = bits And &H3F&

    If bits = 0& Then
        U64_Shl = v
        Exit Function
    End If

    If bits >= 32& Then
        U64_Shl.Hi = U32Shl(v.Lo, bits - 32&)
        U64_Shl.Lo = 0&
        Exit Function
    End If

    carry = U32Shr(v.Lo, 32& - bits)
    U64_Shl.Lo = U32Shl(v.Lo, bits)
    U64_Shl.Hi = U32Add(U32Shl(v.Hi, bits), carry)
End Function

Public Function U64_Shr(ByRef v As U64_t, ByVal bits As Long) As U64_t
    Dim carry As Long

    bits = bits And &H3F&

    If bits = 0& Then
        U64_Shr = v
        Exit Function
    End If

    If bits >= 32& Then
        U64_Shr.Lo = U32Shr(v.Hi, bits - 32&)
        U64_Shr.Hi = 0&
        Exit Function
    End If

    carry = U32Shl(v.Hi, 32& - bits)
    U64_Shr.Hi = U32Shr(v.Hi, bits)
    U64_Shr.Lo = U32Add(U32Shr(v.Lo, bits), carry)
End Function

Public Function U64_And(ByRef a As U64_t, ByRef b As U64_t) As U64_t
    U64_And.Lo = (a.Lo And b.Lo)
    U64_And.Hi = (a.Hi And b.Hi)
End Function

Public Function U64_Or(ByRef a As U64_t, ByRef b As U64_t) As U64_t
    U64_Or.Lo = (a.Lo Or b.Lo)
    U64_Or.Hi = (a.Hi Or b.Hi)
End Function

Public Function U64_Xor(ByRef a As U64_t, ByRef b As U64_t) As U64_t
    U64_Xor.Lo = (a.Lo Xor b.Lo)
    U64_Xor.Hi = (a.Hi Xor b.Hi)
End Function

Public Function U64_IsZero(ByRef v As U64_t) As Long
    If (v.Lo = 0&) And (v.Hi = 0&) Then
        U64_IsZero = 1&
    Else
        U64_IsZero = 0&
    End If
End Function

Public Function U64_Lt(ByRef a As U64_t, ByRef b As U64_t) As Long
    If U32Lt(a.Hi, b.Hi) <> 0& Then
        U64_Lt = 1&
        Exit Function
    End If

    If U32Lt(b.Hi, a.Hi) <> 0& Then
        U64_Lt = 0&
        Exit Function
    End If

    U64_Lt = U32Lt(a.Lo, b.Lo)
End Function

Public Function U64_ToDouble(ByRef v As U64_t) As Double
    U64_ToDouble = (U32ToDouble(v.Hi) * 4294967296#) + U32ToDouble(v.Lo)
End Function

Public Function U64_FromDoubleFloor(ByVal d As Double) As U64_t
    Dim hiD As Double
    Dim loD As Double

    If d <= 0# Then
        U64_FromDoubleFloor = U64_Zero()
        Exit Function
    End If

    hiD = Fix(d / 4294967296#)
    loD = d - (hiD * 4294967296#)

    U64_FromDoubleFloor.Hi = U32FromDouble(hiD)
    U64_FromDoubleFloor.Lo = U32FromDouble(Fix(loD))
End Function

Public Function U64_AddU32(ByRef a As U64_t, ByVal b As Long) As U64_t
    Dim temp As U64_t

    temp = U64_FromU32(b)
    U64_AddU32 = U64_Add(a, temp)
End Function
Public Function U64_Eq(ByRef a As U64_t, ByRef b As U64_t) As Long
    If (a.Lo = b.Lo) And (a.Hi = b.Hi) Then
        U64_Eq = 1&
    Else
        U64_Eq = 0&
    End If
End Function
