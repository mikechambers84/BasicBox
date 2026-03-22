Attribute VB_Name = "modPCSpeaker"
Option Explicit

Public Const PC_SPEAKER_GATE_DIRECT As Byte = 0&
Public Const PC_SPEAKER_GATE_TIMER2 As Byte = 1&

Public Const PC_SPEAKER_USE_DIRECT As Byte = 0&
Public Const PC_SPEAKER_USE_TIMER2 As Byte = 1&

Public Const PC_SPEAKER_MOVEMENT As Long = 800&

Private pcspeaker_gateSelect As Byte
Private pcspeaker_gate(0& To 1&) As Byte
Private pcspeaker_amplitude As Integer

Public Sub pcspeaker_setGateState(ByVal spk As Long, ByVal gate As Byte, ByVal value As Byte)
    If gate > 1& Then Exit Sub
    pcspeaker_gate(gate) = value
End Sub

Public Sub pcspeaker_selectGate(ByVal spk As Long, ByVal value As Byte)
    pcspeaker_gateSelect = value
End Sub

Public Sub pcspeaker_callback(ByVal spk As Long)
    If pcspeaker_gateSelect = PC_SPEAKER_USE_TIMER2 Then
        If (pcspeaker_gate(PC_SPEAKER_GATE_TIMER2) <> 0&) And (pcspeaker_gate(PC_SPEAKER_GATE_DIRECT) <> 0&) Then
            If pcspeaker_amplitude < 15000& Then pcspeaker_amplitude = pcspeaker_amplitude + PC_SPEAKER_MOVEMENT
        Else
            If pcspeaker_amplitude > 0& Then pcspeaker_amplitude = pcspeaker_amplitude - PC_SPEAKER_MOVEMENT
        End If
    Else
        If pcspeaker_gate(PC_SPEAKER_GATE_DIRECT) <> 0& Then
            If pcspeaker_amplitude < 15000& Then pcspeaker_amplitude = pcspeaker_amplitude + PC_SPEAKER_MOVEMENT
        Else
            If pcspeaker_amplitude > 0& Then pcspeaker_amplitude = pcspeaker_amplitude - PC_SPEAKER_MOVEMENT
        End If
    End If

    If pcspeaker_amplitude > 15000& Then pcspeaker_amplitude = 15000&
    If pcspeaker_amplitude < 0& Then pcspeaker_amplitude = 0&
End Sub

Public Sub pcspeaker_init(ByRef machineRef As MACHINE_t)
    pcspeaker_gateSelect = PC_SPEAKER_GATE_DIRECT
    pcspeaker_gate(0&) = 0&
    pcspeaker_gate(1&) = 0&
    pcspeaker_amplitude = 0&

    timing_addTimer TIMER_CB_PCSPEAKER, 0&, SAMPLE_RATE, TIMING_ENABLED
End Sub

Public Function pcspeaker_getSample(ByVal spk As Long) As Integer
    pcspeaker_getSample = pcspeaker_amplitude
End Function
