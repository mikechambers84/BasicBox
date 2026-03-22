Attribute VB_Name = "modFPU"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)

Private Const CPU_REG_EAX As Long = 0&

Private Const FPU_CW_DEFAULT As Long = &H37F&
Private Const FPU_SW_TOPLESS_MASK As Long = &HC7FF&
Private Const FPU_SW_CLEAR_EXC As Long = &H80FF&

Private Const FPU_C0 As Long = &H100&
Private Const FPU_C1 As Long = &H200&
Private Const FPU_C2 As Long = &H400&
Private Const FPU_C3 As Long = &H4000&

Private Const EFLAGS_CF As Long = &H1&
Private Const EFLAGS_PF As Long = &H4&
Private Const EFLAGS_AF As Long = &H10&
Private Const EFLAGS_ZF As Long = &H40&
Private Const EFLAGS_SF As Long = &H80&
Private Const EFLAGS_OF As Long = &H800&

Private Const BIAS80 As Long = 16383&
Private Const BIAS64 As Long = 1023&

Private Const PI_VAL As Double = 3.1415926535897931#
Private Const L2E_VAL As Double = 1.4426950408889634#
Private Const L2T_VAL As Double = 3.3219280948873623#
Private Const LN2_VAL As Double = 0.69314718055994531#
Private Const LG2_VAL As Double = 0.3010299956639812#

Private Const TWO_POW_32_D As Double = 4294967296#
Private Const TWO_POW_23_D As Double = 8388608#
Private Const I64_LIMIT_D As Double = 9223372036854775808#

Private Type F80_t
    mant0 As Long
    mant1 As Long
    high As Long
End Type

Private Type FPUSTATE_t
    cw As Long
    sw As Long
    top As Long
    st(0& To 7&) As Double
    rawst(0& To 7&) As F80_t
    rawtagr As Long
    rawtagw As Long
End Type

Private gFpu As FPUSTATE_t

Public Sub fpu_reset()
    Dim blank As FPUSTATE_t

    gFpu = blank
    gFpu.cw = FPU_CW_DEFAULT
End Sub

Private Function cpu_load16(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef value As Long) As Long
    value = (cpu_readw(cpu, addr) And &HFFFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_load16 = 0&
    Else
        cpu_load16 = 1&
    End If
End Function

Private Function cpu_load32(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef value As Long) As Long
    value = cpu_readl(cpu, addr)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_load32 = 0&
    Else
        cpu_load32 = 1&
    End If
End Function

Private Function cpu_store8(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Long) As Long
    cpu_write cpu, addr, (value And &HFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_store8 = 0&
    Else
        cpu_store8 = 1&
    End If
End Function

Private Function cpu_store16(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Long) As Long
    cpu_writew cpu, addr, (value And &HFFFF&)
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_store16 = 0&
    Else
        cpu_store16 = 1&
    End If
End Function

Private Function cpu_store32(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Long) As Long
    cpu_writel cpu, addr, value
    If (cpu.doexception <> 0&) And (cpu.nowrite <> 0&) Then
        cpu_store32 = 0&
    Else
        cpu_store32 = 1&
    End If
End Function

Private Function cpu_getflags(ByRef cpu As CPU_t) As Long
    cpu_getflags = makeflagsword(cpu)
End Function

Private Sub cpu_setflags(ByRef cpu As CPU_t, ByVal set_mask As Long, ByVal clear_mask As Long)
    Dim flags As Long

    flags = makeflagsword(cpu)
    flags = (flags Or set_mask)
    flags = (flags And (Not clear_mask))
    decodeflagsword cpu, flags
End Sub

Private Sub cpu_setax(ByRef cpu As CPU_t, ByVal value As Long)
    putreg16 cpu, CPU_REG_EAX, (value And &HFFFF&)
End Sub

Private Sub cpu_setexc(ByRef cpu As CPU_t, ByVal vector As Long, ByVal errCode As Long)
    cpu_raiseExceptionFromFirstIP cpu, vector, errCode
End Sub

Private Function getsw() As Long
    getsw = ((gFpu.sw And FPU_SW_TOPLESS_MASK) Or U32Shl((gFpu.top And &H7&), 11&))
End Function

Private Sub setsw(ByVal sw As Long)
    gFpu.sw = (sw And &HFFFF&)
    gFpu.top = (U32Shr(sw, 11&) And &H7&)
End Sub

Private Function fpget(ByVal i As Long) As Double
    Dim idx As Long
    Dim mask As Long

    idx = ((gFpu.top + i) And &H7&)
    mask = U32Shl(1&, idx)

    If ((gFpu.rawtagr And mask) = 0&) And ((gFpu.rawtagw And mask) = 0&) Then
        gFpu.st(idx) = F80ToDouble(gFpu.rawst(idx))
        gFpu.rawtagr = (gFpu.rawtagr Or mask)
    End If

    fpget = gFpu.st(idx)
End Function

Private Sub fpset(ByVal i As Long, ByVal value As Double)
    Dim idx As Long
    Dim mask As Long

    idx = ((gFpu.top + i) And &H7&)
    mask = U32Shl(1&, idx)

    gFpu.st(idx) = value
    gFpu.rawtagw = (gFpu.rawtagw Or mask)
End Sub

Private Sub fppush(ByVal value As Double)
    gFpu.top = ((gFpu.top - 1&) And &H7&)
    Call fpset(0&, value)
End Sub

Private Sub fppop()
    gFpu.top = ((gFpu.top + 1&) And &H7&)
End Sub

Private Function DoubleToBits(ByVal value As Double) As U64_t
    CopyMemory DoubleToBits, value, 8&
End Function

Private Function BitsToDouble(ByRef bits As U64_t) As Double
    CopyMemory BitsToDouble, bits, 8&
End Function

Private Function SignBit64(ByVal value As Double) As Long
    Dim bits As U64_t

    bits = DoubleToBits(value)
    SignBit64 = IIf((bits.Hi And &H80000000) <> 0&, 1&, 0&)
End Function

Private Function IsNaN64(ByVal value As Double) As Long
    Dim bits As U64_t

    bits = DoubleToBits(value)
    If (U32Shr(bits.Hi, 20&) And &H7FF&) = &H7FF& Then
        If ((bits.Hi And &HFFFFF&) <> 0&) Or (bits.Lo <> 0&) Then
            IsNaN64 = 1&
        End If
    End If
End Function

Private Function IsInf64(ByVal value As Double) As Long
    Dim bits As U64_t

    bits = DoubleToBits(value)
    If (U32Shr(bits.Hi, 20&) And &H7FF&) = &H7FF& Then
        If ((bits.Hi And &HFFFFF&) = 0&) And (bits.Lo = 0&) Then
            IsInf64 = 1&
        End If
    End If
End Function

Private Function IsFinite64(ByVal value As Double) As Long
    If (IsNaN64(value) = 0&) And (IsInf64(value) = 0&) Then
        IsFinite64 = 1&
    End If
End Function

Private Function IsZero64(ByVal value As Double) As Long
    Dim bits As U64_t

    bits = DoubleToBits(value)
    If ((bits.Hi And &H7FFFFFFF) = 0&) And (bits.Lo = 0&) Then
        IsZero64 = 1&
    End If
End Function

Private Function IsUnordered(ByVal a As Double, ByVal b As Double) As Long
    If (IsNaN64(a) <> 0&) Or (IsNaN64(b) <> 0&) Then
        IsUnordered = 1&
    End If
End Function

Private Function CopySign64(ByVal magnitude As Double, ByVal signSource As Double) As Double
    Dim a As U64_t
    Dim b As U64_t

    a = DoubleToBits(magnitude)
    b = DoubleToBits(signSource)

    a.Hi = ((a.Hi And &H7FFFFFFF) Or (b.Hi And &H80000000))
    CopySign64 = BitsToDouble(a)
End Function

Private Function MakeNaN64(ByVal negative As Long) As Double
    Dim bits As U64_t

    bits.Lo = 0&
    bits.Hi = &H7FF80000
    If negative <> 0& Then
        bits.Hi = (bits.Hi Or &H80000000)
    End If

    MakeNaN64 = BitsToDouble(bits)
End Function

Private Function MakeInfinity64(ByVal negative As Long) As Double
    Dim bits As U64_t

    bits.Lo = 0&
    bits.Hi = &H7FF00000
    If negative <> 0& Then
        bits.Hi = (bits.Hi Or &H80000000)
    End If

    MakeInfinity64 = BitsToDouble(bits)
End Function

Private Function MakeZero64(ByVal negative As Long) As Double
    Dim bits As U64_t

    bits.Lo = 0&
    bits.Hi = 0&
    If negative <> 0& Then
        bits.Hi = &H80000000
    End If

    MakeZero64 = BitsToDouble(bits)
End Function

Private Function SignExt16Local(ByVal value As Long) As Long
    value = (value And &HFFFF&)
    If (value And &H8000&) <> 0& Then
        SignExt16Local = (value Or &HFFFF0000)
    Else
        SignExt16Local = value
    End If
End Function

Private Function U64NegLocal(ByRef v As U64_t) As U64_t
    Dim inv As U64_t
    Dim one As U64_t

    inv.Lo = Not v.Lo
    inv.Hi = Not v.Hi

    one.Lo = 1&
    one.Hi = 0&

    U64NegLocal = U64_Add(inv, one)
End Function

Private Function I64IsNegative(ByRef v As U64_t) As Long
    If v.Hi < 0& Then
        I64IsNegative = 1&
    Else
        I64IsNegative = 0&
    End If
End Function

Private Function U64CompareLocal(ByRef a As U64_t, ByRef b As U64_t) As Long
    If U32Lt(a.Hi, b.Hi) <> 0& Then
        U64CompareLocal = -1&
        Exit Function
    End If
    If U32Lt(b.Hi, a.Hi) <> 0& Then
        U64CompareLocal = 1&
        Exit Function
    End If

    If U32Lt(a.Lo, b.Lo) <> 0& Then
        U64CompareLocal = -1&
    ElseIf U32Lt(b.Lo, a.Lo) <> 0& Then
        U64CompareLocal = 1&
    Else
        U64CompareLocal = 0&
    End If
End Function

Private Function U64GetBitLocal(ByRef v As U64_t, ByVal bitIdx As Long) As Long
    If bitIdx >= 32& Then
        U64GetBitLocal = (U32Shr(v.Hi, bitIdx - 32&) And &H1&)
    Else
        U64GetBitLocal = (U32Shr(v.Lo, bitIdx) And &H1&)
    End If
End Function

Private Sub U64SetBitLocal(ByRef v As U64_t, ByVal bitIdx As Long)
    If bitIdx >= 32& Then
        v.Hi = (v.Hi Or U32Shl(1&, bitIdx - 32&))
    Else
        v.Lo = (v.Lo Or U32Shl(1&, bitIdx))
    End If
End Sub

Private Sub U64DivModU32Local(ByRef dividend As U64_t, ByVal divisor As Long, ByRef quotient As U64_t, ByRef remOut As Long)
    Dim i As Long
    Dim rem64 As U64_t
    Dim div64 As U64_t
    Dim bitVal As Long

    quotient = U64_Zero()
    rem64 = U64_Zero()
    div64.Lo = divisor
    div64.Hi = 0&

    For i = 63& To 0& Step -1&
        rem64 = U64_Shl(rem64, 1&)
        bitVal = U64GetBitLocal(dividend, i)
        If bitVal <> 0& Then
            rem64.Lo = (rem64.Lo Or &H1&)
        End If

        If U64CompareLocal(rem64, div64) >= 0& Then
            rem64 = U64_Sub(rem64, div64)
            Call U64SetBitLocal(quotient, i)
        End If
    Next i

    remOut = rem64.Lo
End Sub

Private Function TinyFpuFfsll(ByRef x As U64_t) As Long
    Dim i As Long

    If U64_IsZero(x) <> 0& Then
        TinyFpuFfsll = 0&
        Exit Function
    End If

    For i = 0& To 63&
        If U64GetBitLocal(x, i) <> 0& Then
            TinyFpuFfsll = (i + 1&)
            Exit Function
        End If
    Next i
End Function

Private Function U64Mul100Add(ByRef value As U64_t, ByVal addend As Long) As U64_t
    Dim x64 As U64_t
    Dim x32 As U64_t
    Dim x4 As U64_t
    Dim sum As U64_t
    Dim add64 As U64_t

    x64 = U64_Shl(value, 6&)
    x32 = U64_Shl(value, 5&)
    x4 = U64_Shl(value, 2&)
    sum = U64_Add(x64, x32)
    sum = U64_Add(sum, x4)
    add64 = U64_FromU32(addend And &HFF&)
    U64Mul100Add = U64_Add(sum, add64)
End Function

Private Function UnsignedU64ToDouble(ByRef v As U64_t) As Double
    UnsignedU64ToDouble = ((U32ToDouble(v.Hi) * TWO_POW_32_D) + U32ToDouble(v.Lo))
End Function

Private Function I64ToDouble(ByRef v As U64_t) As Double
    Dim mag As U64_t
    Dim negative As Long

    mag = v
    negative = I64IsNegative(mag)
    If negative <> 0& Then
        mag = U64NegLocal(mag)
    End If

    I64ToDouble = ((U32ToDouble(mag.Hi) * TWO_POW_32_D) + U32ToDouble(mag.Lo))
    If negative <> 0& Then
        I64ToDouble = -I64ToDouble
    End If
End Function

Private Function DoubleToI64CastCore(ByVal value As Double) As U64_t
    Dim negative As Long
    Dim mag As Double

    If (IsFinite64(value) = 0&) Or (value >= I64_LIMIT_D) Or (value < -I64_LIMIT_D) Then
        DoubleToI64CastCore.Lo = 0&
        DoubleToI64CastCore.Hi = &H80000000
        Exit Function
    End If

    If value < 0# Then
        negative = 1&
        mag = -value
    Else
        negative = 0&
        mag = value
    End If

    DoubleToI64CastCore.Hi = U32FromDouble(Fix(mag / TWO_POW_32_D))
    DoubleToI64CastCore.Lo = U32FromDouble(mag - (U32ToDouble(DoubleToI64CastCore.Hi) * TWO_POW_32_D))

    If negative <> 0& Then
        DoubleToI64CastCore = U64NegLocal(DoubleToI64CastCore)
    End If
End Function

Private Function DoubleToI64Trunc(ByVal value As Double) As U64_t
    DoubleToI64Trunc = DoubleToI64CastCore(TruncLike(value))
End Function

Private Function DoubleToI64NearestEven(ByVal value As Double) As U64_t
    DoubleToI64NearestEven = DoubleToI64CastCore(NearbyIntEven(value))
End Function

Private Function F32BitsToDouble(ByVal bits As Long) As Double
    Dim signBit As Long
    Dim expv As Long
    Dim mant As Long
    Dim result As Double

    signBit = (U32Shr(bits, 31&) And &H1&)
    expv = (U32Shr(bits, 23&) And &HFF&)
    mant = (bits And &H7FFFFF)

    If expv = 0& Then
        If mant = 0& Then
            F32BitsToDouble = MakeZero64(signBit)
        Else
            result = (CDbl(mant) * (2# ^ -149#))
            If signBit <> 0& Then
                result = -result
            End If
            F32BitsToDouble = result
        End If
    ElseIf expv = &HFF& Then
        If mant = 0& Then
            F32BitsToDouble = MakeInfinity64(signBit)
        Else
            F32BitsToDouble = MakeNaN64(signBit)
        End If
    Else
        result = ((1# + (CDbl(mant) / TWO_POW_23_D)) * (2# ^ CDbl(expv - 127&)))
        If signBit <> 0& Then
            result = -result
        End If
        F32BitsToDouble = result
    End If
End Function

Private Function DoubleToFloat32Bits(ByVal value As Double) As Long
    Dim s As Single
    Dim bits As Long

    If IsNaN64(value) <> 0& Then
        DoubleToFloat32Bits = &H7FC00000
        If SignBit64(value) <> 0& Then
            DoubleToFloat32Bits = (DoubleToFloat32Bits Or &H80000000)
        End If
        Exit Function
    End If

    If IsInf64(value) <> 0& Then
        DoubleToFloat32Bits = &H7F800000
        If SignBit64(value) <> 0& Then
            DoubleToFloat32Bits = (DoubleToFloat32Bits Or &H80000000)
        End If
        Exit Function
    End If

    If IsZero64(value) <> 0& Then
        DoubleToFloat32Bits = 0&
        If SignBit64(value) <> 0& Then
            DoubleToFloat32Bits = &H80000000
        End If
        Exit Function
    End If

    On Error GoTo Overflowed
    s = CSng(value)
    CopyMemory bits, s, 4&
    DoubleToFloat32Bits = bits
    Exit Function

Overflowed:
    Err.Clear
    DoubleToFloat32Bits = &H7F800000
    If SignBit64(value) <> 0& Then
        DoubleToFloat32Bits = (DoubleToFloat32Bits Or &H80000000)
    End If
End Function

Private Function DoubleToF80(ByVal value As Double) As F80_t
    Dim bits As U64_t
    Dim signBit As Long
    Dim expv As Long
    Dim mant80 As U64_t
    Dim shift As Long

    bits = DoubleToBits(value)
    signBit = (U32Shr(bits.Hi, 31&) And &H1&)
    expv = (U32Shr(bits.Hi, 20&) And &H7FF&)

    mant80.Lo = bits.Lo
    mant80.Hi = (bits.Hi And &HFFFFF&)
    mant80 = U64_Shl(mant80, 11&)

    If expv = 0& Then
        If U64_IsZero(mant80) = 0& Then
            shift = (64& - TinyFpuFfsll(mant80))
            mant80 = U64_Shl(mant80, shift)
            expv = (expv + BIAS80 - BIAS64 + 1& - shift)
        End If
    ElseIf expv = &H7FF& Then
        Call U64SetBitLocal(mant80, 63&)
        expv = &H7FFF&
    Else
        Call U64SetBitLocal(mant80, 63&)
        expv = (expv + BIAS80 - BIAS64)
    End If

    DoubleToF80.high = (((signBit And &H1&) * &H8000&) Or (expv And &H7FFF&))
    DoubleToF80.mant1 = mant80.Hi
    DoubleToF80.mant0 = mant80.Lo
End Function

Private Function F80ToDouble(ByRef f80 As F80_t) As Double
    Dim signBit As Long
    Dim expv As Long
    Dim mant80 As U64_t
    Dim mant64 As U64_t
    Dim bits As U64_t

    signBit = (U32Shr(f80.high, 15&) And &H1&)
    expv = (f80.high And &H7FFF&)
    mant80.Lo = f80.mant0
    mant80.Hi = f80.mant1

    If expv = 0& Then
        mant64 = U64_Zero()
    ElseIf expv = &H7FFF& Then
        expv = &H7FF&
        mant64 = U64_Shr(mant80, 11&)
        If (U64_IsZero(mant64) <> 0&) And (U64_IsZero(mant80) = 0&) Then
            mant64 = U64_FromU32(1&)
        End If
    Else
        expv = (expv + BIAS64 - BIAS80)
        If expv <= -52& Then
            expv = 0&
            mant64 = U64_Zero()
        ElseIf expv <= 0& Then
            mant64 = U64_Shr(mant80, (12& - expv))
            expv = 0&
        ElseIf expv >= &H7FF& Then
            expv = &H7FF&
            mant64 = U64_Zero()
        Else
            mant64 = U64_Shr(mant80, 11&)
            mant64.Hi = (mant64.Hi And &HFFFFF&)
        End If
    End If

    bits.Lo = mant64.Lo
    bits.Hi = ((mant64.Hi And &HFFFFF&) Or U32Shl((expv And &H7FF&), 20&))
    If signBit <> 0& Then
        bits.Hi = (bits.Hi Or &H80000000)
    End If

    F80ToDouble = BitsToDouble(bits)
End Function

Private Function FloorLike(ByVal x As Double) As Double
    If (IsFinite64(x) = 0&) Or (IsZero64(x) <> 0&) Then
        FloorLike = x
    Else
        FloorLike = Int(x)
    End If
End Function

Private Function CeilLike(ByVal x As Double) As Double
    If (IsFinite64(x) = 0&) Or (IsZero64(x) <> 0&) Then
        CeilLike = x
    Else
        CeilLike = -Int(-x)
    End If
End Function

Private Function TruncLike(ByVal x As Double) As Double
    If (IsFinite64(x) = 0&) Or (IsZero64(x) <> 0&) Then
        TruncLike = x
    Else
        TruncLike = Fix(x)
    End If
End Function

Private Function NearbyIntEven(ByVal x As Double) As Double
    If (IsFinite64(x) = 0&) Or (IsZero64(x) <> 0&) Then
        NearbyIntEven = x
        Exit Function
    End If

    On Error GoTo Fail
    NearbyIntEven = Round(x, 0&)
    Exit Function

Fail:
    Err.Clear
    NearbyIntEven = x
End Function

Private Function fpround(ByVal x As Double, ByVal rc As Long) As Double
    Select Case (rc And &H3&)
        Case 0&
            fpround = NearbyIntEven(x)
        Case 1&
            fpround = FloorLike(x)
        Case 2&
            fpround = CeilLike(x)
        Case 3&
            fpround = TruncLike(x)
    End Select
End Function

Private Function Log2Like(ByVal x As Double) As Double
    If IsNaN64(x) <> 0& Then
        Log2Like = x
    ElseIf IsZero64(x) <> 0& Then
        Log2Like = MakeInfinity64(1&)
    ElseIf IsInf64(x) <> 0& Then
        If SignBit64(x) <> 0& Then
            Log2Like = MakeNaN64(0&)
        Else
            Log2Like = x
        End If
    ElseIf x < 0# Then
        Log2Like = MakeNaN64(0&)
    Else
        On Error GoTo Fail
        Log2Like = (Log(x) / Log(2#))
        Exit Function
    End If
    Exit Function

Fail:
    Err.Clear
    Log2Like = MakeNaN64(0&)
End Function

Private Function Pow2Like(ByVal x As Double) As Double
    If IsNaN64(x) <> 0& Then
        Pow2Like = x
        Exit Function
    End If

    If IsInf64(x) <> 0& Then
        If SignBit64(x) <> 0& Then
            Pow2Like = 0#
        Else
            Pow2Like = MakeInfinity64(0&)
        End If
        Exit Function
    End If

    On Error GoTo Overflowed
    Pow2Like = (2# ^ x)
    Exit Function

Overflowed:
    Err.Clear
    If x < 0# Then
        Pow2Like = 0#
    Else
        Pow2Like = MakeInfinity64(0&)
    End If
End Function

Private Function SqrtLike(ByVal x As Double) As Double
    If IsNaN64(x) <> 0& Then
        SqrtLike = x
    ElseIf IsZero64(x) <> 0& Then
        SqrtLike = x
    ElseIf IsInf64(x) <> 0& Then
        If SignBit64(x) <> 0& Then
            SqrtLike = MakeNaN64(0&)
        Else
            SqrtLike = x
        End If
    ElseIf x < 0# Then
        SqrtLike = MakeNaN64(0&)
    Else
        On Error GoTo Fail
        SqrtLike = Sqr(x)
        Exit Function
    End If
    Exit Function

Fail:
    Err.Clear
    SqrtLike = MakeNaN64(0&)
End Function

Private Function SinLike(ByVal x As Double) As Double
    If IsFinite64(x) = 0& Then
        SinLike = MakeNaN64(0&)
        Exit Function
    End If

    On Error GoTo Fail
    SinLike = Sin(x)
    Exit Function

Fail:
    Err.Clear
    SinLike = MakeNaN64(0&)
End Function

Private Function CosLike(ByVal x As Double) As Double
    If IsFinite64(x) = 0& Then
        CosLike = MakeNaN64(0&)
        Exit Function
    End If

    On Error GoTo Fail
    CosLike = Cos(x)
    Exit Function

Fail:
    Err.Clear
    CosLike = MakeNaN64(0&)
End Function

Private Function TanLike(ByVal x As Double) As Double
    If IsFinite64(x) = 0& Then
        TanLike = MakeNaN64(0&)
        Exit Function
    End If

    On Error GoTo Fail
    TanLike = Tan(x)
    Exit Function

Fail:
    Err.Clear
    TanLike = MakeNaN64(0&)
End Function

Private Function Atan2Like(ByVal y As Double, ByVal x As Double) As Double
    If (IsNaN64(x) <> 0&) Or (IsNaN64(y) <> 0&) Then
        Atan2Like = MakeNaN64(0&)
        Exit Function
    End If

    If (IsInf64(x) <> 0&) And (IsInf64(y) <> 0&) Then
        If SignBit64(x) <> 0& Then
            If SignBit64(y) <> 0& Then
                Atan2Like = (-3# * PI_VAL / 4#)
            Else
                Atan2Like = (3# * PI_VAL / 4#)
            End If
        Else
            If SignBit64(y) <> 0& Then
                Atan2Like = (-PI_VAL / 4#)
            Else
                Atan2Like = (PI_VAL / 4#)
            End If
        End If
        Exit Function
    End If

    If IsInf64(y) <> 0& Then
        If SignBit64(y) <> 0& Then
            Atan2Like = (-PI_VAL / 2#)
        Else
            Atan2Like = (PI_VAL / 2#)
        End If
        Exit Function
    End If

    If IsInf64(x) <> 0& Then
        If SignBit64(x) <> 0& Then
            If SignBit64(y) <> 0& Then
                Atan2Like = -PI_VAL
            Else
                Atan2Like = PI_VAL
            End If
        Else
            If SignBit64(y) <> 0& And (IsZero64(y) = 0&) Then
                Atan2Like = MakeZero64(1&)
            Else
                Atan2Like = 0#
            End If
        End If
        Exit Function
    End If

    If x > 0# Then
        Atan2Like = Atn(y / x)
    ElseIf x < 0# Then
        If y >= 0# Then
            Atan2Like = (Atn(y / x) + PI_VAL)
        Else
            Atan2Like = (Atn(y / x) - PI_VAL)
        End If
    Else
        If y > 0# Then
            Atan2Like = (PI_VAL / 2#)
        ElseIf y < 0# Then
            Atan2Like = (-PI_VAL / 2#)
        Else
            Atan2Like = 0#
        End If
    End If
End Function

Private Function FrexpLike(ByVal x As Double, ByRef expOut As Long) As Double
    Dim temp As Double

    If (IsFinite64(x) = 0&) Or (IsZero64(x) <> 0&) Then
        expOut = 0&
        FrexpLike = x
        Exit Function
    End If

    temp = Abs(x)
    expOut = 0&

    Do While temp >= 1#
        temp = (temp / 2#)
        expOut = expOut + 1&
    Loop

    Do While temp < 0.5
        temp = (temp * 2#)
        expOut = expOut - 1&
    Loop

    If SignBit64(x) <> 0& Then
        temp = -temp
    End If

    FrexpLike = temp
End Function

Private Function SafeAdd(ByVal a As Double, ByVal b As Double) As Double
    If IsNaN64(a) <> 0& Then
        SafeAdd = a
        Exit Function
    End If
    If IsNaN64(b) <> 0& Then
        SafeAdd = b
        Exit Function
    End If

    If (IsInf64(a) <> 0&) Or (IsInf64(b) <> 0&) Then
        If (IsInf64(a) <> 0&) And (IsInf64(b) <> 0&) Then
            If SignBit64(a) <> SignBit64(b) Then
                SafeAdd = MakeNaN64(0&)
            Else
                SafeAdd = a
            End If
        ElseIf IsInf64(a) <> 0& Then
            SafeAdd = a
        Else
            SafeAdd = b
        End If
        Exit Function
    End If

    On Error GoTo Overflowed
    SafeAdd = (a + b)
    Exit Function

Overflowed:
    Err.Clear
    SafeAdd = MakeInfinity64(SignBit64(a))
End Function

Private Function SafeSub(ByVal a As Double, ByVal b As Double) As Double
    If IsNaN64(a) <> 0& Then
        SafeSub = a
        Exit Function
    End If
    If IsNaN64(b) <> 0& Then
        SafeSub = b
        Exit Function
    End If

    If (IsInf64(a) <> 0&) Or (IsInf64(b) <> 0&) Then
        If (IsInf64(a) <> 0&) And (IsInf64(b) <> 0&) Then
            If SignBit64(a) = SignBit64(b) Then
                SafeSub = MakeNaN64(0&)
            Else
                SafeSub = a
            End If
        ElseIf IsInf64(a) <> 0& Then
            SafeSub = a
        ElseIf SignBit64(b) <> 0& Then
            SafeSub = MakeInfinity64(0&)
        Else
            SafeSub = MakeInfinity64(1&)
        End If
        Exit Function
    End If

    On Error GoTo Overflowed
    SafeSub = (a - b)
    Exit Function

Overflowed:
    Err.Clear
    SafeSub = MakeInfinity64(SignBit64(a))
End Function

Private Function SafeMul(ByVal a As Double, ByVal b As Double) As Double
    Dim negative As Long

    If IsNaN64(a) <> 0& Then
        SafeMul = a
        Exit Function
    End If
    If IsNaN64(b) <> 0& Then
        SafeMul = b
        Exit Function
    End If

    If ((IsInf64(a) <> 0&) And (IsZero64(b) <> 0&)) Or ((IsInf64(b) <> 0&) And (IsZero64(a) <> 0&)) Then
        SafeMul = MakeNaN64(0&)
        Exit Function
    End If

    negative = (SignBit64(a) Xor SignBit64(b))
    If (IsInf64(a) <> 0&) Or (IsInf64(b) <> 0&) Then
        SafeMul = MakeInfinity64(negative)
        Exit Function
    End If

    On Error GoTo Overflowed
    SafeMul = (a * b)
    Exit Function

Overflowed:
    Err.Clear
    SafeMul = MakeInfinity64(negative)
End Function

Private Function SafeDiv(ByVal a As Double, ByVal b As Double) As Double
    Dim negative As Long

    If IsNaN64(a) <> 0& Then
        SafeDiv = a
        Exit Function
    End If
    If IsNaN64(b) <> 0& Then
        SafeDiv = b
        Exit Function
    End If

    negative = (SignBit64(a) Xor SignBit64(b))

    If IsZero64(b) <> 0& Then
        If IsZero64(a) <> 0& Then
            SafeDiv = MakeNaN64(0&)
        Else
            SafeDiv = MakeInfinity64(negative)
        End If
        Exit Function
    End If

    If (IsInf64(a) <> 0&) And (IsInf64(b) <> 0&) Then
        SafeDiv = MakeNaN64(0&)
        Exit Function
    End If

    If IsInf64(a) <> 0& Then
        SafeDiv = MakeInfinity64(negative)
        Exit Function
    End If

    If IsInf64(b) <> 0& Then
        SafeDiv = MakeZero64(negative)
        Exit Function
    End If

    On Error GoTo Fail
    SafeDiv = (a / b)
    Exit Function

Fail:
    Err.Clear
    SafeDiv = MakeInfinity64(negative)
End Function

Private Function fploadf32(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim bits As Long

    If cpu_load32(cpu, seg, addr, bits) = 0& Then
        fploadf32 = 0&
        Exit Function
    End If

    res = F32BitsToDouble(bits)
    fploadf32 = 1&
End Function

Private Function fploadf64(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim loPart As Long
    Dim hiPart As Long
    Dim bits As U64_t

    If cpu_load32(cpu, seg, addr, loPart) = 0& Then
        fploadf64 = 0&
        Exit Function
    End If
    If cpu_load32(cpu, seg, U32Add(addr, 4&), hiPart) = 0& Then
        fploadf64 = 0&
        Exit Function
    End If

    bits.Lo = loPart
    bits.Hi = hiPart
    res = BitsToDouble(bits)
    fploadf64 = 1&
End Function

Private Function fploadf80(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim f80 As F80_t

    If cpu_load32(cpu, seg, addr, f80.mant0) = 0& Then
        fploadf80 = 0&
        Exit Function
    End If
    If cpu_load32(cpu, seg, U32Add(addr, 4&), f80.mant1) = 0& Then
        fploadf80 = 0&
        Exit Function
    End If
    If cpu_load16(cpu, seg, U32Add(addr, 8&), f80.high) = 0& Then
        fploadf80 = 0&
        Exit Function
    End If

    res = F80ToDouble(f80)
    fploadf80 = 1&
End Function

Private Function fploadi16(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim value As Long

    If cpu_load16(cpu, seg, addr, value) = 0& Then
        fploadi16 = 0&
        Exit Function
    End If

    res = CDbl(SignExt16Local(value))
    fploadi16 = 1&
End Function

Private Function fploadi32(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim value As Long

    If cpu_load32(cpu, seg, addr, value) = 0& Then
        fploadi32 = 0&
        Exit Function
    End If

    res = CDbl(value)
    fploadi32 = 1&
End Function

Private Function fploadi64(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim value As U64_t

    If cpu_load32(cpu, seg, addr, value.Lo) = 0& Then
        fploadi64 = 0&
        Exit Function
    End If
    If cpu_load32(cpu, seg, U32Add(addr, 4&), value.Hi) = 0& Then
        fploadi64 = 0&
        Exit Function
    End If

    res = I64ToDouble(value)
    fploadi64 = 1&
End Function

Private Function bcd100(ByVal b As Long) As Long
    b = (b And &HFF&)
    bcd100 = (b - (6& * U32Shr(b, 4&)))
End Function

Private Function fploadbcd(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByRef res As Double) As Long
    Dim loPart As Long
    Dim midPart As Long
    Dim hiPart As Long
    Dim value As U64_t
    Dim i As Long
    Dim signFlag As Long

    If cpu_load32(cpu, seg, addr, loPart) = 0& Then
        fploadbcd = 0&
        Exit Function
    End If
    If cpu_load32(cpu, seg, U32Add(addr, 4&), midPart) = 0& Then
        fploadbcd = 0&
        Exit Function
    End If
    If cpu_load16(cpu, seg, U32Add(addr, 8&), hiPart) = 0& Then
        fploadbcd = 0&
        Exit Function
    End If

    signFlag = (hiPart And &H8000&)
    hiPart = (hiPart And &H7FFF&)

    value = U64_Zero()
    For i = 0& To 3&
        value = U64Mul100Add(value, bcd100(loPart))
        loPart = U32Shr(loPart, 8&)
    Next i
    For i = 0& To 3&
        value = U64Mul100Add(value, bcd100(midPart))
        midPart = U32Shr(midPart, 8&)
    Next i
    For i = 0& To 1&
        value = U64Mul100Add(value, bcd100(hiPart))
        hiPart = U32Shr(hiPart, 8&)
    Next i

    res = UnsignedU64ToDouble(value)
    If signFlag <> 0& Then
        res = CopySign64(res, -1#)
    End If
    fploadbcd = 1&
End Function

Private Function fpstoref32(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    fpstoref32 = cpu_store32(cpu, seg, addr, DoubleToFloat32Bits(value))
End Function

Private Function fpstoref64(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    Dim bits As U64_t

    bits = DoubleToBits(value)
    If cpu_store32(cpu, seg, addr, bits.Lo) = 0& Then
        fpstoref64 = 0&
        Exit Function
    End If
    If cpu_store32(cpu, seg, U32Add(addr, 4&), bits.Hi) = 0& Then
        fpstoref64 = 0&
        Exit Function
    End If

    fpstoref64 = 1&
End Function

Private Function fpstoref80(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    Dim f80 As F80_t

    f80 = DoubleToF80(value)
    If cpu_store32(cpu, seg, addr, f80.mant0) = 0& Then
        fpstoref80 = 0&
        Exit Function
    End If
    If cpu_store32(cpu, seg, U32Add(addr, 4&), f80.mant1) = 0& Then
        fpstoref80 = 0&
        Exit Function
    End If
    If cpu_store16(cpu, seg, U32Add(addr, 8&), f80.high) = 0& Then
        fpstoref80 = 0&
        Exit Function
    End If

    fpstoref80 = 1&
End Function

Private Function fpstorei16(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    Dim outValue As Long

    If (IsFinite64(value) <> 0&) And (value < 32768#) And (value >= -32768#) Then
        outValue = CLng(Fix(value))
    Else
        outValue = &H8000&
    End If

    fpstorei16 = cpu_store16(cpu, seg, addr, outValue)
End Function

Private Function fpstorei32(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    Dim outValue As Long

    If (IsFinite64(value) <> 0&) And (value < 2147483648#) And (value >= -2147483648#) Then
        outValue = CLng(Fix(value))
    Else
        outValue = &H80000000
    End If

    fpstorei32 = cpu_store32(cpu, seg, addr, outValue)
End Function

Private Function fpstorei64(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    Dim outValue As U64_t

    outValue = DoubleToI64CastCore(value)
    If cpu_store32(cpu, seg, addr, outValue.Lo) = 0& Then
        fpstorei64 = 0&
        Exit Function
    End If
    If cpu_store32(cpu, seg, U32Add(addr, 4&), outValue.Hi) = 0& Then
        fpstorei64 = 0&
        Exit Function
    End If

    fpstorei64 = 1&
End Function

Private Function fpstorebcd(ByRef cpu As CPU_t, ByVal seg As Long, ByVal addr As Long, ByVal value As Double) As Long
    Dim signedValue As U64_t
    Dim mag As U64_t
    Dim quotient As U64_t
    Dim signFlag As Long
    Dim remVal As Long
    Dim outByte As Long
    Dim i As Long

    signedValue = DoubleToI64CastCore(value)
    mag = signedValue
    signFlag = I64IsNegative(signedValue)
    If signFlag <> 0& Then
        mag = U64NegLocal(mag)
    End If

    For i = 0& To 8&
        Call U64DivModU32Local(mag, 100&, quotient, remVal)
        outByte = ((remVal Mod 10&) Or U32Shl(((remVal \ 10&) And &HF&), 4&))
        If cpu_store8(cpu, seg, U32Add(addr, i), outByte) = 0& Then
            fpstorebcd = 0&
            Exit Function
        End If
        mag = quotient
    Next i

    Call U64DivModU32Local(mag, 100&, quotient, remVal)
    outByte = ((remVal Mod 10&) Or U32Shl(((remVal \ 10&) And &HF&), 4&))
    If signFlag <> 0& Then
        outByte = (outByte Or &H80&)
    End If
    If cpu_store8(cpu, seg, U32Add(addr, 9&), outByte) = 0& Then
        fpstorebcd = 0&
        Exit Function
    End If

    fpstorebcd = 1&
End Function

Private Sub UpdateCompareStatus(ByVal a As Double, ByVal b As Double)
    If IsUnordered(a, b) <> 0& Then
        gFpu.sw = (gFpu.sw Or (FPU_C0 Or FPU_C2 Or FPU_C3))
    ElseIf a = b Then
        gFpu.sw = (gFpu.sw Or FPU_C3)
        gFpu.sw = (gFpu.sw And (Not (FPU_C0 Or FPU_C2)))
    ElseIf a < b Then
        gFpu.sw = (gFpu.sw Or FPU_C0)
        gFpu.sw = (gFpu.sw And (Not (FPU_C2 Or FPU_C3)))
    Else
        gFpu.sw = (gFpu.sw And (Not (FPU_C0 Or FPU_C2 Or FPU_C3)))
    End If
End Sub

Private Sub UpdateQuotientBits(ByRef q As U64_t)
    If (q.Lo And &H1&) <> 0& Then
        gFpu.sw = (gFpu.sw Or FPU_C1)
    Else
        gFpu.sw = (gFpu.sw And (Not FPU_C1))
    End If

    If (q.Lo And &H2&) <> 0& Then
        gFpu.sw = (gFpu.sw Or FPU_C3)
    Else
        gFpu.sw = (gFpu.sw And (Not FPU_C3))
    End If

    If (q.Lo And &H4&) <> 0& Then
        gFpu.sw = (gFpu.sw Or FPU_C0)
    Else
        gFpu.sw = (gFpu.sw And (Not FPU_C0))
    End If
End Sub

Private Sub fparith(ByVal group As Long, ByVal destIdx As Long, ByVal a As Double, ByVal b As Double)
    Dim c As Double

    Select Case (group And &H7&)
        Case 0&
            c = SafeAdd(a, b)
        Case 1&
            c = SafeMul(a, b)
        Case 2&
            Call UpdateCompareStatus(a, b)
        Case 3&
            Call UpdateCompareStatus(a, b)
        Case 4&
            c = SafeSub(a, b)
        Case 5&
            c = SafeSub(b, a)
        Case 6&
            c = SafeDiv(a, b)
        Case 7&
            c = SafeDiv(b, a)
        Case Else
            Debug.Assert False
            Exit Sub
    End Select

    If (group And &H7&) = 3& Then
        Call fppop
    ElseIf (group And &H7&) <> 2& Then
        Call fpset(destIdx, c)
    End If
End Sub

Private Function cmov_cond(ByRef cpu As CPU_t, ByVal i As Long) As Long
    Dim flags As Long

    flags = cpu_getflags(cpu)
    Select Case (i And &H3&)
        Case 0&
            cmov_cond = IIf((flags And EFLAGS_CF) <> 0&, 1&, 0&)
        Case 1&
            cmov_cond = IIf((flags And EFLAGS_ZF) <> 0&, 1&, 0&)
        Case 2&
            cmov_cond = IIf((flags And (EFLAGS_CF Or EFLAGS_ZF)) <> 0&, 1&, 0&)
        Case 3&
            cmov_cond = IIf((flags And EFLAGS_PF) <> 0&, 1&, 0&)
        Case Else
            Debug.Assert False
            cmov_cond = 0&
    End Select
End Function

Private Sub ucomi(ByRef cpu As CPU_t, ByVal i As Long)
    Dim a As Double
    Dim b As Double

    a = fpget(0&)
    b = fpget(i)

    If IsUnordered(a, b) <> 0& Then
        Call cpu_setflags(cpu, (EFLAGS_ZF Or EFLAGS_PF Or EFLAGS_CF), 0&)
    ElseIf a = b Then
        Call cpu_setflags(cpu, EFLAGS_ZF, (EFLAGS_PF Or EFLAGS_CF))
    ElseIf a < b Then
        Call cpu_setflags(cpu, EFLAGS_CF, (EFLAGS_ZF Or EFLAGS_PF))
    Else
        Call cpu_setflags(cpu, 0&, (EFLAGS_ZF Or EFLAGS_PF Or EFLAGS_CF))
    End If
End Sub

Public Function fpu_exec2(ByRef cpu As CPU_t, ByVal opsz16 As Long, ByVal op As Long, ByVal group As Long, ByVal seg As Long, ByVal addr As Long) As Long
    Dim a As Double
    Dim b As Double
    Dim sw As Long
    Dim startAddr As Long
    Dim j As Long
    Dim rc As Long
    Dim mask As Long

    Select Case (op And &H7&)
        Case 0&
            a = fpget(0&)
            If fploadf32(cpu, seg, addr, b) = 0& Then
                fpu_exec2 = 0&
                Exit Function
            End If
            Call fparith(group, 0&, a, b)

        Case 1&
            Select Case (group And &H7&)
                Case 0&
                    If fploadf32(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 1&
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec2 = 0&
                    Exit Function

                Case 2&
                    a = fpget(0&)
                    If fpstoref32(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If

                Case 3&
                    a = fpget(0&)
                    If fpstoref32(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppop

                Case 4&
                    If opsz16 <> 0& Then
                        If cpu_load16(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_load16(cpu, seg, U32Add(addr, 2&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                    Else
                        If cpu_load16(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_load16(cpu, seg, U32Add(addr, 4&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                    End If
                    Call setsw(sw)

                Case 5&
                    If cpu_load16(cpu, seg, addr, gFpu.cw) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If

                Case 6&
                    If opsz16 <> 0& Then
                        If cpu_store16(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        sw = getsw()
                        If cpu_store16(cpu, seg, U32Add(addr, 2&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_store16(cpu, seg, U32Add(addr, 4&), 0&) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                    Else
                        If cpu_store32(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        sw = getsw()
                        If cpu_store32(cpu, seg, U32Add(addr, 4&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_store32(cpu, seg, U32Add(addr, 8&), 0&) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                    End If

                Case 7&
                    If cpu_store16(cpu, seg, addr, gFpu.cw) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
            End Select

        Case 2&
            a = fpget(0&)
            If fploadi32(cpu, seg, addr, b) = 0& Then
                fpu_exec2 = 0&
                Exit Function
            End If
            Call fparith(group, 0&, a, b)

        Case 3&
            Select Case (group And &H7&)
                Case 0&
                    If fploadi32(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 1&, 2&, 3&
                    a = fpget(0&)
                    If (group And &H7&) <> 1& Then
                        rc = (U32Shr(gFpu.cw, 10&) And &H3&)
                        a = fpround(a, rc)
                    End If
                    If fpstorei32(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    If (group And &H7&) <> 2& Then
                        Call fppop
                    End If

                Case 4&
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec2 = 0&
                    Exit Function

                Case 5&
                    If fploadf80(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 6&
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec2 = 0&
                    Exit Function

                Case 7&
                    a = fpget(0&)
                    If fpstoref80(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppop
            End Select

        Case 4&
            a = fpget(0&)
            If fploadf64(cpu, seg, addr, b) = 0& Then
                fpu_exec2 = 0&
                Exit Function
            End If
            Call fparith(group, 0&, a, b)

        Case 5&
            Select Case (group And &H7&)
                Case 0&
                    If fploadf64(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 1&
                    a = fpget(0&)
                    If fpstorei64(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppop

                Case 2&
                    a = fpget(0&)
                    If fpstoref64(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If

                Case 3&
                    a = fpget(0&)
                    If fpstoref64(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppop

                Case 4&
                    startAddr = addr
                    If opsz16 <> 0& Then
                        If cpu_load16(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_load16(cpu, seg, U32Add(addr, 2&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        startAddr = U32Add(startAddr, 14&)
                    Else
                        If cpu_load16(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_load16(cpu, seg, U32Add(addr, 4&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        startAddr = U32Add(startAddr, 28&)
                    End If
                    Call setsw(sw)

                    For j = 0& To 7&
                        If cpu_load32(cpu, seg, startAddr, gFpu.rawst(j).mant0) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_load32(cpu, seg, U32Add(startAddr, 4&), gFpu.rawst(j).mant1) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_load16(cpu, seg, U32Add(startAddr, 8&), gFpu.rawst(j).high) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        mask = U32Shl(1&, j)
                        gFpu.rawtagr = (gFpu.rawtagr And (Not mask))
                        gFpu.rawtagw = (gFpu.rawtagw And (Not mask))
                        startAddr = U32Add(startAddr, 10&)
                    Next j

                Case 5&
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec2 = 0&
                    Exit Function

                Case 6&
                    startAddr = addr
                    If opsz16 <> 0& Then
                        If cpu_store16(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        sw = getsw()
                        If cpu_store16(cpu, seg, U32Add(addr, 2&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_store16(cpu, seg, U32Add(addr, 4&), 0&) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        startAddr = U32Add(startAddr, 14&)
                    Else
                        If cpu_store32(cpu, seg, addr, gFpu.cw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        sw = getsw()
                        If cpu_store32(cpu, seg, U32Add(addr, 4&), sw) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_store32(cpu, seg, U32Add(addr, 8&), 0&) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        startAddr = U32Add(startAddr, 28&)
                    End If

                    For j = 0& To 7&
                        mask = U32Shl(1&, j)
                        If (gFpu.rawtagw And mask) <> 0& Then
                            gFpu.rawst(j) = DoubleToF80(gFpu.st(j))
                            gFpu.rawtagw = (gFpu.rawtagw And (Not mask))
                        End If

                        If cpu_store32(cpu, seg, startAddr, gFpu.rawst(j).mant0) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_store32(cpu, seg, U32Add(startAddr, 4&), gFpu.rawst(j).mant1) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        If cpu_store16(cpu, seg, U32Add(startAddr, 8&), gFpu.rawst(j).high) = 0& Then
                            fpu_exec2 = 0&
                            Exit Function
                        End If
                        startAddr = U32Add(startAddr, 10&)
                    Next j

                    gFpu.sw = 0&
                    gFpu.top = 0&
                    gFpu.cw = FPU_CW_DEFAULT

                Case 7&
                    If cpu_store16(cpu, seg, addr, getsw()) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
            End Select

        Case 6&
            a = fpget(0&)
            If fploadi16(cpu, seg, addr, b) = 0& Then
                fpu_exec2 = 0&
                Exit Function
            End If
            Call fparith(group, 0&, a, b)

        Case 7&
            Select Case (group And &H7&)
                Case 0&
                    If fploadi16(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 1&, 2&, 3&
                    a = fpget(0&)
                    If (group And &H7&) <> 1& Then
                        rc = (U32Shr(gFpu.cw, 10&) And &H3&)
                        a = fpround(a, rc)
                    End If
                    If fpstorei16(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    If (group And &H7&) <> 2& Then
                        Call fppop
                    End If

                Case 4&
                    If fploadbcd(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 5&
                    If fploadi64(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppush(a)

                Case 6&
                    rc = (U32Shr(gFpu.cw, 10&) And &H3&)
                    a = fpget(0&)
                    a = fpround(a, rc)
                    If fpstorebcd(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppop

                Case 7&
                    rc = (U32Shr(gFpu.cw, 10&) And &H3&)
                    a = fpget(0&)
                    a = fpround(a, rc)
                    If fpstorei64(cpu, seg, addr, a) = 0& Then
                        fpu_exec2 = 0&
                        Exit Function
                    End If
                    Call fppop
            End Select
    End Select

    fpu_exec2 = 1&
End Function

Public Function fpu_exec1(ByRef cpu As CPU_t, ByVal op As Long, ByVal group As Long, ByVal i As Long) As Long
    Dim a As Double
    Dim b As Double
    Dim temp As Double
    Dim temp2 As Double
    Dim q As U64_t
    Dim sw As Long
    Dim rc As Long
    Dim expv As Long
    Dim mant As Double

    Select Case (op And &H7&)
        Case 0&
            a = fpget(0&)
            b = fpget(i)
            Call fparith(group, 0&, a, b)

        Case 1&
            Select Case (group And &H7&)
                Case 0&
                    temp = fpget(i)
                    Call fppush(temp)

                Case 1&
                    temp = fpget(i)
                    temp2 = fpget(0&)
                    Call fpset(i, temp2)
                    Call fpset(0&, temp)

                Case 2&
                    ' FNOP

                Case 3&
                    temp = fpget(0&)
                    Call fpset(i, temp)
                    Call fppop

                Case 4&
                    temp = fpget(0&)
                    Select Case (i And &H7&)
                        Case 0&
                            If SignBit64(temp) <> 0& Then
                                Call fpset(0&, CopySign64(temp, 1#))
                            Else
                                Call fpset(0&, CopySign64(temp, -1#))
                            End If

                        Case 1&
                            Call fpset(0&, CopySign64(temp, 1#))

                        Case 2&, 3&, 6&, 7&
                            Call cpu_setexc(cpu, 6&, 0&)
                            fpu_exec1 = 0&
                            Exit Function

                        Case 4&
                            Call UpdateCompareStatus(temp, 0#)

                        Case 5&
                            If SignBit64(temp) <> 0& Then
                                gFpu.sw = (gFpu.sw Or FPU_C1)
                            Else
                                gFpu.sw = (gFpu.sw And (Not FPU_C1))
                            End If

                            If IsZero64(temp) <> 0& Then
                                gFpu.sw = (gFpu.sw Or FPU_C3)
                                gFpu.sw = (gFpu.sw And (Not (FPU_C0 Or FPU_C2)))
                            ElseIf IsNaN64(temp) <> 0& Then
                                gFpu.sw = (gFpu.sw Or FPU_C0)
                                gFpu.sw = (gFpu.sw And (Not (FPU_C2 Or FPU_C3)))
                            ElseIf IsFinite64(temp) <> 0& Then
                                gFpu.sw = (gFpu.sw Or FPU_C2)
                                gFpu.sw = (gFpu.sw And (Not (FPU_C0 Or FPU_C3)))
                            Else
                                gFpu.sw = (gFpu.sw Or (FPU_C0 Or FPU_C2))
                                gFpu.sw = (gFpu.sw And (Not FPU_C3))
                            End If
                    End Select

                Case 5&
                    Select Case (i And &H7&)
                        Case 0&: Call fppush(1#)
                        Case 1&: Call fppush(L2T_VAL)
                        Case 2&: Call fppush(L2E_VAL)
                        Case 3&: Call fppush(PI_VAL)
                        Case 4&: Call fppush(LG2_VAL)
                        Case 5&: Call fppush(LN2_VAL)
                        Case 6&: Call fppush(0#)
                        Case Else
                            Call cpu_setexc(cpu, 6&, 0&)
                            fpu_exec1 = 0&
                            Exit Function
                    End Select

                Case 6&
                    temp = fpget(0&)
                    Select Case (i And &H7&)
                        Case 0&
                            Call fpset(0&, (Pow2Like(temp) - 1#))

                        Case 1&
                            temp2 = fpget(1&)
                            Call fpset(1&, (temp2 * Log2Like(temp)))
                            Call fppop

                        Case 2&
                            Call fpset(0&, TanLike(temp))
                            Call fppush(1#)
                            gFpu.sw = (gFpu.sw And (Not FPU_C2))

                        Case 3&
                            temp2 = fpget(1&)
                            Call fpset(1&, Atan2Like(temp2, temp))
                            Call fppop

                        Case 4&
                            mant = FrexpLike(temp, expv)
                            mant = mant * 2#
                            expv = expv - 1&
                            Call fpset(0&, CDbl(expv))
                            Call fppush(mant)

                        Case 5&
                            temp2 = fpget(1&)
                            q = DoubleToI64NearestEven(SafeDiv(temp, temp2))
                            Call fpset(0&, (temp - (I64ToDouble(q) * temp2)))
                            gFpu.sw = (gFpu.sw And (Not FPU_C2))
                            Call UpdateQuotientBits(q)

                        Case 6&
                            gFpu.top = ((gFpu.top - 1&) And &H7&)

                        Case 7&
                            gFpu.top = ((gFpu.top + 1&) And &H7&)
                    End Select

                Case 7&
                    temp = fpget(0&)
                    Select Case (i And &H7&)
                        Case 0&
                            temp2 = fpget(1&)
                            q = DoubleToI64Trunc(SafeDiv(temp, temp2))
                            Call fpset(0&, (temp - (I64ToDouble(q) * temp2)))
                            gFpu.sw = (gFpu.sw And (Not FPU_C2))
                            Call UpdateQuotientBits(q)

                        Case 1&
                            temp2 = fpget(1&)
                            Call fpset(1&, (temp2 * Log2Like(1# + temp)))
                            Call fppop

                        Case 2&
                            Call fpset(0&, SqrtLike(temp))

                        Case 3&
                            Call fpset(0&, SinLike(temp))
                            Call fppush(CosLike(temp))
                            gFpu.sw = (gFpu.sw And (Not FPU_C2))

                        Case 4&
                            rc = (U32Shr(gFpu.cw, 10&) And &H3&)
                            Call fpset(0&, fpround(temp, rc))

                        Case 5&
                            Call fpset(0&, (temp * Pow2Like(TruncLike(fpget(1&)))))

                        Case 6&
                            Call fpset(0&, SinLike(temp))
                            gFpu.sw = (gFpu.sw And (Not FPU_C2))

                        Case 7&
                            Call fpset(0&, CosLike(temp))
                            gFpu.sw = (gFpu.sw And (Not FPU_C2))
                    End Select
            End Select

        Case 2&
            Select Case (group And &H7&)
                Case 5&
                    If (i And &H7&) = 1& Then
                        a = fpget(0&)
                        b = fpget(1&)
                        Call UpdateCompareStatus(a, b)
                        Call fppop
                        Call fppop
                    Else
                        Call cpu_setexc(cpu, 6&, 0&)
                        fpu_exec1 = 0&
                        Exit Function
                    End If

                Case 0&, 1&, 2&, 3&
                    If cmov_cond(cpu, group) <> 0& Then
                        Call fpset(0&, fpget(i))
                    End If

                Case Else
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec1 = 0&
                    Exit Function
            End Select

        Case 3&
            Select Case (group And &H7&)
                Case 4&
                    Select Case (i And &H7&)
                        Case 0&, 1&, 4&, 5&
                            ' no-op cases

                        Case 2&
                            gFpu.sw = (gFpu.sw And (Not FPU_SW_CLEAR_EXC))

                        Case 3&
                            gFpu.sw = 0&
                            gFpu.top = 0&
                            gFpu.cw = FPU_CW_DEFAULT

                        Case 6&, 7&
                            Call cpu_setexc(cpu, 6&, 0&)
                            fpu_exec1 = 0&
                            Exit Function
                    End Select

                Case 0&, 1&, 2&, 3&
                    If cmov_cond(cpu, group) = 0& Then
                        Call fpset(0&, fpget(i))
                    End If

                Case 5&, 6&
                    Call ucomi(cpu, i)

                Case Else
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec1 = 0&
                    Exit Function
            End Select

        Case 4&
            a = fpget(0&)
            b = fpget(i)
            Call fparith(group, i, a, b)

        Case 5&
            Select Case (group And &H7&)
                Case 0&
                    ' FFREE

                Case 1&
                    temp = fpget(i)
                    temp2 = fpget(0&)
                    Call fpset(i, temp2)
                    Call fpset(0&, temp)

                Case 2&
                    temp = fpget(0&)
                    Call fpset(i, temp)

                Case 3&
                    temp = fpget(0&)
                    Call fpset(i, temp)
                    Call fppop

                Case 4&, 5&
                    a = fpget(0&)
                    b = fpget(i)
                    Call UpdateCompareStatus(a, b)
                    If (group And &H7&) = 5& Then
                        Call fppop
                    End If

                Case Else
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec1 = 0&
                    Exit Function
            End Select

        Case 6&
            a = fpget(0&)
            b = fpget(i)
            Call fparith(group, i, a, b)
            Call fppop

        Case 7&
            Select Case (group And &H7&)
                Case 0&
                    Call fppop

                Case 1&
                    temp = fpget(i)
                    temp2 = fpget(0&)
                    Call fpset(i, temp2)
                    Call fpset(0&, temp)

                Case 2&, 3&
                    temp = fpget(0&)
                    Call fpset(i, temp)
                    Call fppop

                Case 4&
                    If (i And &H7&) = 0& Then
                        sw = getsw()
                        Call cpu_setax(cpu, sw)
                    Else
                        Call cpu_setexc(cpu, 6&, 0&)
                        fpu_exec1 = 0&
                        Exit Function
                    End If

                Case 5&, 6&
                    Call ucomi(cpu, i)
                    Call fppop

                Case Else
                    Call cpu_setexc(cpu, 6&, 0&)
                    fpu_exec1 = 0&
                    Exit Function
            End Select
    End Select

    fpu_exec1 = 1&
End Function
