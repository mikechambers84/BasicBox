Attribute VB_Name = "modSCSI"
Option Explicit

Public Const SCSI_BUS_MAX As Long = 1&
Public Const SCSI_ID_MAX As Long = 16&
Public Const SCSI_LUN_MAX As Long = 8&

Private scsi_next_bus As Byte
Private scsi_bus_speed(0 To SCSI_BUS_MAX - 1) As Double

Public Sub scsi_reset()
    Dim i As Long

    scsi_next_bus = 0&
    For i = 0& To SCSI_BUS_MAX - 1&
        scsi_bus_speed(i) = 0#
    Next i
End Sub

Public Function scsi_get_bus() As Byte
    If scsi_next_bus >= SCSI_BUS_MAX Then
        scsi_get_bus = &HFF&
        Exit Function
    End If

    scsi_get_bus = scsi_next_bus
    scsi_next_bus = CByte((scsi_next_bus + 1&) And &HFF&)
End Function

Public Sub scsi_bus_set_speed(ByVal bus As Byte, ByVal speed As Double)
    If bus < SCSI_BUS_MAX Then
        scsi_bus_speed(bus) = speed
    End If
End Sub

Public Function scsi_bus_get_speed(ByVal bus As Byte) As Double
    If bus >= SCSI_BUS_MAX Then
        scsi_bus_get_speed = 0#
    Else
        scsi_bus_get_speed = scsi_bus_speed(bus)
    End If
End Function
