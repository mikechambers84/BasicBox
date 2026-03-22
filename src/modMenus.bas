Attribute VB_Name = "modMenus"
Option Explicit

Private menus_resetTimer As Long
Private menus_resetPos As Byte
Private menus_ready As Byte
Private menus_ctrlaltdel(0& To 2&) As Byte

Private Function menus_isConsoleLoaded() As Long
    Dim f As Form

    For Each f In Forms
        If StrComp(f.Name, "frmConsole", vbTextCompare) = 0& Then
            menus_isConsoleLoaded = 1&
            Exit Function
        End If
    Next f
End Function

Private Sub menus_ensureInit()
    If menus_ready <> 0& Then Exit Sub

    menus_ctrlaltdel(0&) = &H1D&
    menus_ctrlaltdel(1&) = &H38&
    menus_ctrlaltdel(2&) = &H53&
    menus_resetPos = 0&

    menus_resetTimer = timing_addTimer(TIMER_CB_MENUS_RESET, 0&, 10#, TIMING_DISABLED)
    menus_ready = 1&
End Sub

Private Function menus_scsi_cd_slot_visible(ByVal targetId As Long) As Long
    If (targetId < 0&) Or (targetId >= BUSLOGIC_MAX_TARGETS) Then Exit Function
    If machine.buslogic_enabled = 0& Then Exit Function
    If machine.scsi_targets(targetId).present = 0& Then Exit Function
    If machine.scsi_targets(targetId).targetType <> BUSLOGIC_TARGET_CDROM Then Exit Function
    menus_scsi_cd_slot_visible = 1&
End Function

Public Function menus_init() As Long
    menus_ensureInit
    menus_init = 0&
End Function

Public Sub menus_setMachine(ByRef machineRef As MACHINE_t)
    menus_ensureInit
End Sub

Public Sub menus_resetCallback(ByVal dummy As Long)
    If menus_ready = 0& Then Exit Sub

    machine.KeyState.scancode = menus_ctrlaltdel(menus_resetPos)
    menus_resetPos = menus_resetPos + 1&
    machine.KeyState.isNew = 1&
    i8259_doirq machine.i8259, 1&

    If menus_resetPos = 3& Then
        timing_timerDisable menus_resetTimer
    End If
End Sub

Public Sub menus_exit()
    running = 0&
End Sub

Public Sub menus_openFloppyFile(ByVal disk As Byte)
    menus_openFloppyFileEx disk, 0&
End Sub

Public Sub menus_openFloppyFileEx(ByVal disk As Byte, ByVal forceWriteProtected As Byte)
    Dim filename As String

    filename = utility_openFileDialog("Open floppy image", "Floppy Images (*.img;*.ima;*.dsk)|*.img;*.ima;*.dsk|All Files (*.*)|*.*", frmConsole.hWnd)
    If LenB(filename) = 0& Then Exit Sub

    fdd_loadEx disk, filename, forceWriteProtected
End Sub

Public Sub menus_openHardFile(ByVal disk As Byte)
    Dim filename As String

    filename = InputBox$("Enter hard disk image path", "Open disk image")
    If LenB(filename) = 0& Then Exit Sub

    ata_insert_disk (disk - 2&), filename

    menus_reset
End Sub

Public Sub menus_refreshScsiMenu()
    Dim i As Long
    Dim anyVisible As Long
    Dim slotVisible As Long

    If menus_isConsoleLoaded() = 0& Then Exit Sub

    For i = 0& To 7&
        If menus_scsi_cd_slot_visible(i) <> 0& Then
            anyVisible = 1&
            Exit For
        End If
    Next i

    If anyVisible = 0& Then
        frmConsole.mnuSCSI.Visible = False
        Exit Sub
    End If

    frmConsole.mnuSCSI.Visible = True

    For i = 0& To 7&
        slotVisible = menus_scsi_cd_slot_visible(i)
        frmConsole.itmSCSICD(i).Visible = CBool(slotVisible)
        frmConsole.itmEjectSCSICD(i).Visible = CBool(slotVisible)
    Next i
End Sub

Public Sub menus_changeScsiCD(ByVal targetId As Long)
    Dim filename As String
    Dim busId As Long

    If (targetId < 0&) Or (targetId >= BUSLOGIC_MAX_TARGETS) Then Exit Sub
    If menus_scsi_cd_slot_visible(targetId) = 0& Then Exit Sub

    filename = utility_openFileDialog("Open SCSI CD image", "ISO Images (*.iso)|*.iso|All Files (*.*)|*.*", frmConsole.hWnd)
    If LenB(filename) = 0& Then Exit Sub

    busId = buslogic_get_bus_id()
    If busId < 0& Then Exit Sub

    If scsi_cdrom_attach(CByte(busId And &HFF&), CByte(targetId And &HFF&), filename) <> 0& Then
        debug_log DEBUG_ERROR, "[SCSI] Failed to attach CD-ROM target " & CStr(targetId) & vbCrLf
        Exit Sub
    End If

    machine.scsi_targets(targetId).present = 1&
    machine.scsi_targets(targetId).targetType = BUSLOGIC_TARGET_CDROM
    machine.scsi_targets(targetId).path = filename
    menus_refreshScsiMenu
End Sub

Public Sub menus_ejectScsiCD(ByVal targetId As Long)
    Dim busId As Long

    If (targetId < 0&) Or (targetId >= BUSLOGIC_MAX_TARGETS) Then Exit Sub
    If menus_scsi_cd_slot_visible(targetId) = 0& Then Exit Sub

    busId = buslogic_get_bus_id()
    If busId < 0& Then Exit Sub

    scsi_cdrom_eject CByte(busId And &HFF&), CByte(targetId And &HFF&)
    machine.scsi_targets(targetId).present = 1&
    machine.scsi_targets(targetId).targetType = BUSLOGIC_TARGET_CDROM
    machine.scsi_targets(targetId).path = vbNullString
    menus_refreshScsiMenu
End Sub

Public Sub menus_changeFloppy0()
    menus_openFloppyFile 0&
End Sub

Public Sub menus_changeFloppy0WP()
    menus_openFloppyFileEx 0&, 1&
End Sub

Public Sub menus_changeFloppy1()
    menus_openFloppyFile 1&
End Sub

Public Sub menus_changeFloppy1WP()
    menus_openFloppyFileEx 1&, 1&
End Sub

Public Sub menus_ejectFloppy0()
    fdd_eject 0&
End Sub

Public Sub menus_ejectFloppy1()
    fdd_eject 1&
End Sub

Public Sub menus_insertHard0()
    menus_openHardFile 2&
End Sub

Public Sub menus_insertHard1()
    menus_openHardFile 3&
End Sub

Public Sub menus_reset()
    menus_ensureInit
    menus_resetPos = 0&
    timing_timerEnable menus_resetTimer
End Sub

Public Sub menus_speed477()
    setspeed 4.77
End Sub

Public Sub menus_speed8()
    setspeed 8#
End Sub

Public Sub menus_speed10()
    setspeed 10#
End Sub

Public Sub menus_speed16()
    setspeed 16#
End Sub

Public Sub menus_speed25()
    setspeed 25#
End Sub

Public Sub menus_speed50()
    setspeed 50#
End Sub

Public Sub menus_speedunlimited()
    setspeed 0#
End Sub
