Attribute VB_Name = "modATA"
Option Explicit
Private Const ATA_PORT_DATA As Integer = &H1F0&
Private Const ATA_PORT_ERROR As Integer = &H1F1&
Private Const ATA_PORT_FEATURES As Integer = &H1F1&
Private Const ATA_PORT_SECTORS As Integer = &H1F2&
Private Const ATA_PORT_LBA_LOW As Integer = &H1F3&
Private Const ATA_PORT_LBA_MID As Integer = &H1F4&
Private Const ATA_PORT_LBA_HIGH As Integer = &H1F5&
Private Const ATA_PORT_DRIVE As Integer = &H1F6&
Private Const ATA_PORT_STATUS As Integer = &H1F7&
Private Const ATA_PORT_COMMAND As Integer = &H1F7&
Private Const ATA_PORT_ALTERNATE As Integer = &H3F6&
Private Const ATA_CMD_IDENTIFY As Byte = &HEC&
Private Const ATA_CMD_DIAGNOSTIC As Byte = &H90&
Private Const ATA_CMD_INITIALIZE_PARAMS As Byte = &H91&
Private Const ATA_CMD_RECALIBRATE As Byte = &H10&
Private Const ATA_CMD_IDLE_IMMEDIATE As Byte = &HE1&
Private Const ATA_CMD_READ_SECTORS As Byte = &H20&
Private Const ATA_CMD_WRITE_SECTORS As Byte = &H30&
Private Const ATA_CMD_DEVICE_RESET As Byte = &H8&
Private Const ATA_STATUS_BUSY As Byte = &H80&
Private Const ATA_STATUS_DRDY As Byte = &H40&
Private Const ATA_STATUS_DSC As Byte = &H10&
Private Const ATA_STATUS_DRQ As Byte = &H8&
Private Const ATA_STATUS_ERR As Byte = &H1&
Private Const ATA_FILEPOS_MAX As Long = 2147483646
Private Type ATA_REGS_t
    features As Byte
    sectors As Byte
    lba As Long
    drive As Byte
End Type
Private Type ATA_DISK_t
    fileNum As Integer
    openFlag As Byte
    sectors As Long
    heads As Long
    cylinders As Long
    spt As Long
    cursect As Long
    curhead As Long
    curcyl As Long
    iswriting As Byte
    isreading As Byte
    interrupt As Byte
    model As String
    errorCode As Byte
    status As Byte
    command As Byte
    lbamode As Byte
    lastcmd As Byte
    curreadsect As Long
    targetsect As Long
    buffer(0& To 511&) As Byte
    buffer_pos As Long
    regs As ATA_REGS_t
End Type
Private Type ATA_t
    selectDisk As Byte
    delay_irq As Byte
    inreset As Byte
    control As Byte
    irq_pending As Byte
    irq_drive As Byte
    timerNum As Long
    resettimer As Long
    readssincecommand As Byte
    savelba As Long
    dscflag As Byte
    i8259 As Long
    disk(0& To 1&) As ATA_DISK_t
End Type
Private ata As ATA_t
Private ata_swap(0& To 19&) As Byte
Private Function ATA_IsValidDisk(ByVal disk As Long) As Boolean
    ATA_IsValidDisk = ((disk >= 0&) And (disk <= 1&))
End Function
Private Function ATA_Read512(ByVal fileNum As Integer, ByVal fileOffset As Long, ByRef buf() As Byte) As Boolean
    On Error GoTo ReadFail
    If (fileOffset < 0&) Or (fileOffset > ATA_FILEPOS_MAX) Then
        ATA_Read512 = False
        Exit Function
    End If
    Get #fileNum, (fileOffset + 1&), buf
    ATA_Read512 = True
    Exit Function
ReadFail:
    ATA_Read512 = False
End Function
Private Function ATA_Write512(ByVal fileNum As Integer, ByVal fileOffset As Long, ByRef buf() As Byte) As Boolean
    On Error GoTo WriteFail
    If (fileOffset < 0&) Or (fileOffset > ATA_FILEPOS_MAX) Then
        ATA_Write512 = False
        Exit Function
    End If
    Put #fileNum, (fileOffset + 1&), buf
    ATA_Write512 = True
    Exit Function
WriteFail:
    ATA_Write512 = False
End Function
Private Sub ATA_CloseDisk(ByVal disk As Long)
    On Error Resume Next
    If ata.disk(disk).fileNum <> 0& Then Close #ata.disk(disk).fileNum
    On Error GoTo 0
    ata.disk(disk).fileNum = 0&
    ata.disk(disk).openFlag = 0&
End Sub
Private Sub ATA_LowerIRQ()
    If (ata.i8259 >= 0&) And (ata.irq_pending <> 0&) Then
        i8259_clearirq ata.i8259, 6&
    End If
    ata.irq_pending = 0&
End Sub
Private Sub ATA_SetInterruptEnable(ByVal enabled As Byte)
    ata.disk(0&).interrupt = enabled
    ata.disk(1&).interrupt = enabled
    If enabled = 0& Then
        ATA_LowerIRQ
    End If
End Sub
Private Sub ATA_ResetDrive(ByVal disk As Long)
    If Not ATA_IsValidDisk(disk) Then Exit Sub

    ata.disk(disk).errorCode = 1&
    ata.disk(disk).isreading = 0&
    ata.disk(disk).iswriting = 0&
    ata.disk(disk).lbamode = 0&
    ata.disk(disk).lastcmd = 0&
    ata.disk(disk).curreadsect = 0&
    ata.disk(disk).targetsect = 0&
    ata.disk(disk).buffer_pos = 512&
    ata.disk(disk).cursect = 1&
    ata.disk(disk).curhead = 0&
    ata.disk(disk).curcyl = 0&
    ata.disk(disk).regs.features = 0&
    ata.disk(disk).regs.sectors = 1&
    ata.disk(disk).regs.lba = 1&
End Sub
Private Sub ATA_CompleteReset()
    ata.delay_irq = 0&
    ata.inreset = 0&
    ata.selectDisk = 0&
    ata.dscflag = ATA_STATUS_DSC
    ATA_ResetDrive 0&
    ATA_ResetDrive 1&
    ATA_LowerIRQ
End Sub
Private Sub ATA_IdentifyStoreWord(ByVal disk As Long, ByVal wordIndex As Long, ByVal value As Long)
    Dim offset As Long

    offset = wordIndex * 2&
    ata.disk(disk).buffer(offset) = CByte(value And &HFF&)
    ata.disk(disk).buffer(offset + 1&) = CByte(U32Shr(value, 8&) And &HFF&)
End Sub
Private Sub ATA_IdentifyStoreString(ByVal disk As Long, ByVal wordIndex As Long, ByVal bytesLen As Long, ByVal value As String)
    Dim offset As Long
    Dim i As Long
    Dim a As Byte
    Dim b As Byte
    Dim lengthChars As Long

    offset = wordIndex * 2&
    lengthChars = Len(value)

    For i = 0& To bytesLen - 1&
        ata.disk(disk).buffer(offset + i) = asc(" ")
    Next i

    For i = 0& To bytesLen - 2& Step 2&
        If i < lengthChars Then
            a = CByte(asc(Mid$(value, i + 1&, 1&)) And &HFF&)
        Else
            a = asc(" ")
        End If

        If (i + 1&) < lengthChars Then
            b = CByte(asc(Mid$(value, i + 2&, 1&)) And &HFF&)
        Else
            b = asc(" ")
        End If

        ata.disk(disk).buffer(offset + i) = b
        ata.disk(disk).buffer(offset + i + 1&) = a
    Next i

    If ((bytesLen And 1&) <> 0&) And (bytesLen <= lengthChars) Then
        ata.disk(disk).buffer(offset + bytesLen - 1&) = CByte(asc(Mid$(value, bytesLen, 1&)) And &HFF&)
    End If
End Sub
Private Function ATA_U32Mul(ByVal a As Long, ByVal b As Long) As Long
    Dim i As Long
    Dim result As Long
    Dim multiplicand As Long
    Dim multiplier As Long

    result = 0&
    multiplicand = a
    multiplier = b

    For i = 0& To 31&
        If (multiplier And 1&) <> 0& Then
            result = U32Add(result, multiplicand)
        End If
        multiplier = U32Shr(multiplier, 1&)
        multiplicand = U32Shl(multiplicand, 1&)
    Next i

    ATA_U32Mul = result
End Function
Private Function ATA_CHSToLBA(ByVal cyl As Long, ByVal head As Long, ByVal sector As Long, ByVal heads As Long, ByVal spt As Long) As Long
    Dim lba As Long

    lba = ATA_U32Mul(cyl, heads)
    lba = U32Add(lba, head)
    lba = ATA_U32Mul(lba, spt)
    ATA_CHSToLBA = U32Add(lba, U32Sub(sector, 1&))
End Function
Public Sub ata_delayed_irq(ByVal dummy As Long)
    If ata.delay_irq = 0& Then Exit Sub
    If (ata.inreset = 0&) And (ata.disk(ata.irq_drive).interrupt <> 0&) Then
        i8259_doirq ata.i8259, 6&
        ata.irq_pending = 1&
    End If
    ata.delay_irq = 0&
    ata.dscflag = ATA_STATUS_DSC
    timing_timerDisable ata.timerNum
End Sub
Public Sub ata_reset_cb(ByVal dummy As Long)
    timing_timerDisable ata.resettimer
    ATA_CompleteReset
End Sub
Public Sub ata_irq()
    ata.delay_irq = 1&
    ata.irq_drive = ata.selectDisk
    ata_delayed_irq 0&
End Sub
Public Sub ata_swap_string(ByVal src As String)
    Dim i As Long
    Dim ch As Long
    Dim tmp As Byte
    For i = 0& To 19&
        If (i + 1&) <= Len(src) Then
            ch = asc(Mid$(src, i + 1&, 1&))
            ata_swap(i) = CByte(ch And &HFF&)
        Else
            ata_swap(i) = 0&
        End If
    Next i
    For i = 0& To 18& Step 2&
        tmp = ata_swap(i + 1&)
        ata_swap(i + 1&) = ata_swap(i)
        ata_swap(i) = tmp
    Next i
End Sub
Private Sub ATA_CopySwapToBuffer(ByVal bufOffset As Long, ByVal length As Long)
    Dim i As Long
    For i = 0& To length - 1&
        ata.disk(ata.selectDisk).buffer(bufOffset + i) = ata_swap(i)
    Next i
End Sub
Public Function ata_gen_status() As Byte
    If ata.inreset <> 0& Then
        ata_gen_status = ATA_STATUS_BUSY
        Exit Function
    End If

    If (ata.disk(ata.selectDisk).openFlag = 0&) And _
       (ata.disk(ata.selectDisk).iswriting = 0&) And _
       (ata.disk(ata.selectDisk).isreading = 0&) And _
       (ata.disk(ata.selectDisk).lastcmd = 0&) Then
        ata_gen_status = 0&
        Exit Function
    End If

    If ata.disk(ata.selectDisk).errorCode = 4& Then
        ata_gen_status = (ATA_STATUS_DRDY Or ATA_STATUS_DSC Or ATA_STATUS_ERR)
        Exit Function
    End If

    If ata.delay_irq <> 0& Then
        ata_gen_status = ATA_STATUS_BUSY
        Exit Function
    End If

    If (ata.disk(ata.selectDisk).iswriting <> 0&) Or (ata.disk(ata.selectDisk).isreading <> 0&) Then
        ata_gen_status = (ATA_STATUS_DRQ Or ATA_STATUS_DRDY Or ata.dscflag)
        Exit Function
    End If

    ata_gen_status = (ATA_STATUS_DRDY Or ATA_STATUS_DSC)
End Function
Public Sub ata_read_disk()
    Dim curLba As Long
    Dim fileOffset As Long
    If ata.disk(ata.selectDisk).openFlag = 0& Then
        ata.disk(ata.selectDisk).errorCode = 4&
        Exit Sub
    End If
    If ata.disk(ata.selectDisk).lbamode <> 0& Then
        If U32Shr(ata.disk(ata.selectDisk).regs.lba, 22&) <> 0& Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        fileOffset = U32Shl(ata.disk(ata.selectDisk).regs.lba, 9&)
        If Not ATA_Read512(ata.disk(ata.selectDisk).fileNum, fileOffset, ata.disk(ata.selectDisk).buffer) Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        ata.disk(ata.selectDisk).regs.lba = U32Add(ata.disk(ata.selectDisk).regs.lba, 1&)
        ata.disk(ata.selectDisk).errorCode = 0&
        ata.disk(ata.selectDisk).buffer_pos = 0&
        ata.disk(ata.selectDisk).curreadsect = ata.disk(ata.selectDisk).curreadsect + 1&
    Else
        curLba = ATA_CHSToLBA(ata.disk(ata.selectDisk).curcyl, ata.disk(ata.selectDisk).curhead, ata.disk(ata.selectDisk).cursect, ata.disk(ata.selectDisk).heads, ata.disk(ata.selectDisk).spt)
        If U32Shr(curLba, 22&) <> 0& Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        fileOffset = U32Shl(curLba, 9&)
        If Not ATA_Read512(ata.disk(ata.selectDisk).fileNum, fileOffset, ata.disk(ata.selectDisk).buffer) Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        ata.disk(ata.selectDisk).errorCode = 0&
        ata.disk(ata.selectDisk).buffer_pos = 0&
        ata.disk(ata.selectDisk).curreadsect = ata.disk(ata.selectDisk).curreadsect + 1&
        ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFFFFFFC0) Or (ata.disk(ata.selectDisk).cursect And 63&)
        ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFF0000FF) Or ((ata.disk(ata.selectDisk).curcyl And &HFFFF&) * &H100&)
        ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HF0FFFFFF) Or U32Shl((ata.disk(ata.selectDisk).curhead And &HFF&), 24&)
        ata.disk(ata.selectDisk).cursect = ata.disk(ata.selectDisk).cursect + 1&
        If ata.disk(ata.selectDisk).cursect > ata.disk(ata.selectDisk).spt Then
            ata.disk(ata.selectDisk).cursect = 1&
            ata.disk(ata.selectDisk).curhead = ata.disk(ata.selectDisk).curhead + 1&
            If ata.disk(ata.selectDisk).curhead = ata.disk(ata.selectDisk).heads Then
                ata.disk(ata.selectDisk).curhead = 0&
                ata.disk(ata.selectDisk).curcyl = ata.disk(ata.selectDisk).curcyl + 1&
            End If
        End If
    End If
    ata.dscflag = 0&
    If ata.disk(ata.selectDisk).interrupt <> 0& Then
        ata_irq
    End If
End Sub
Public Sub ata_write_disk()
    Dim curLba As Long
    Dim fileOffset As Long
    If ata.disk(ata.selectDisk).openFlag = 0& Then
        ata.disk(ata.selectDisk).errorCode = 4&
        Exit Sub
    End If
    If ata.disk(ata.selectDisk).lbamode <> 0& Then
        If U32Shr(ata.disk(ata.selectDisk).regs.lba, 22&) <> 0& Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        fileOffset = U32Shl(ata.disk(ata.selectDisk).regs.lba, 9&)
        If Not ATA_Write512(ata.disk(ata.selectDisk).fileNum, fileOffset, ata.disk(ata.selectDisk).buffer) Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        ata.disk(ata.selectDisk).regs.lba = U32Add(ata.disk(ata.selectDisk).regs.lba, 1&)
        ata.disk(ata.selectDisk).errorCode = 0&
        ata.disk(ata.selectDisk).buffer_pos = 0&
        ata.disk(ata.selectDisk).curreadsect = ata.disk(ata.selectDisk).curreadsect + 1&
    Else
        curLba = ATA_CHSToLBA(ata.disk(ata.selectDisk).curcyl, ata.disk(ata.selectDisk).curhead, ata.disk(ata.selectDisk).cursect, ata.disk(ata.selectDisk).heads, ata.disk(ata.selectDisk).spt)
        If U32Shr(curLba, 22&) <> 0& Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        fileOffset = U32Shl(curLba, 9&)
        If Not ATA_Write512(ata.disk(ata.selectDisk).fileNum, fileOffset, ata.disk(ata.selectDisk).buffer) Then
            ata.disk(ata.selectDisk).errorCode = 4&
            Exit Sub
        End If
        ata.disk(ata.selectDisk).errorCode = 0&
        ata.disk(ata.selectDisk).buffer_pos = 0&
        ata.disk(ata.selectDisk).curreadsect = ata.disk(ata.selectDisk).curreadsect + 1&
        ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFFFFFFC0) Or (ata.disk(ata.selectDisk).cursect And 63&)
        ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFF0000FF) Or ((ata.disk(ata.selectDisk).curcyl And &HFFFF&) * &H100&)
        ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HF0FFFFFF) Or U32Shl((ata.disk(ata.selectDisk).curhead And &HFF&), 24&)
        ata.disk(ata.selectDisk).cursect = ata.disk(ata.selectDisk).cursect + 1&
        If ata.disk(ata.selectDisk).cursect > ata.disk(ata.selectDisk).spt Then
            ata.disk(ata.selectDisk).cursect = 1&
            ata.disk(ata.selectDisk).curhead = ata.disk(ata.selectDisk).curhead + 1&
            If ata.disk(ata.selectDisk).curhead = ata.disk(ata.selectDisk).heads Then
                ata.disk(ata.selectDisk).curhead = 0&
                ata.disk(ata.selectDisk).curcyl = ata.disk(ata.selectDisk).curcyl + 1&
            End If
        End If
    End If
    ata.dscflag = 1&
    If ata.disk(ata.selectDisk).interrupt <> 0& Then
        ata_irq
    End If
End Sub
Public Sub ata_command_process()
    Dim sector As Long
    Dim cyl As Long
    Dim head As Long
    Dim total_chs As Long
    Dim i As Long
    ata.disk(ata.selectDisk).errorCode = 0&
    ata.disk(ata.selectDisk).iswriting = 0&
    ata.disk(ata.selectDisk).isreading = 0&
    ata.readssincecommand = 0&
    ata.disk(ata.selectDisk).lastcmd = ata.disk(ata.selectDisk).command
    Select Case ata.disk(ata.selectDisk).command
        Case ATA_CMD_IDENTIFY
            If ata.disk(ata.selectDisk).openFlag <> 0& Then
                total_chs = ATA_U32Mul(ATA_U32Mul(ata.disk(ata.selectDisk).cylinders, ata.disk(ata.selectDisk).heads), ata.disk(ata.selectDisk).spt)
                For i = 0& To 511&
                    ata.disk(ata.selectDisk).buffer(i) = 0&
                Next i
                ata.disk(ata.selectDisk).iswriting = 0&
                ata.disk(ata.selectDisk).isreading = 1&
                ata.disk(ata.selectDisk).buffer_pos = 0&
                ata.disk(ata.selectDisk).curreadsect = 1&
                ata.disk(ata.selectDisk).targetsect = 1&
                ata.disk(ata.selectDisk).errorCode = 0&
                ATA_IdentifyStoreWord ata.selectDisk, 0&, &H40&
                ATA_IdentifyStoreWord ata.selectDisk, 1&, (ata.disk(ata.selectDisk).cylinders And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 3&, (ata.disk(ata.selectDisk).heads And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 4&, ((ata.disk(ata.selectDisk).spt * 512&) And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 5&, 512&
                ATA_IdentifyStoreWord ata.selectDisk, 6&, (ata.disk(ata.selectDisk).spt And &HFFFF&)
                ATA_IdentifyStoreString ata.selectDisk, 10&, 20&, "BASICBOX0001"
                ATA_IdentifyStoreWord ata.selectDisk, 20&, 3&
                ATA_IdentifyStoreWord ata.selectDisk, 21&, 16&
                ATA_IdentifyStoreWord ata.selectDisk, 22&, 4&
                ATA_IdentifyStoreString ata.selectDisk, 23&, 8&, "1.0"
                ATA_IdentifyStoreString ata.selectDisk, 27&, 40&, "BasicBox virtual IDE disk"
                ATA_IdentifyStoreWord ata.selectDisk, 47&, 1&
                ATA_IdentifyStoreWord ata.selectDisk, 48&, 1&
                ATA_IdentifyStoreWord ata.selectDisk, 49&, U32Shl(1&, 9&)
                ATA_IdentifyStoreWord ata.selectDisk, 51&, &H200&
                ATA_IdentifyStoreWord ata.selectDisk, 52&, &H200&
                ATA_IdentifyStoreWord ata.selectDisk, 53&, &H7&
                ATA_IdentifyStoreWord ata.selectDisk, 54&, (ata.disk(ata.selectDisk).cylinders And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 55&, (ata.disk(ata.selectDisk).heads And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 56&, (ata.disk(ata.selectDisk).spt And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 57&, (total_chs And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 58&, (U32Shr(total_chs, 16&) And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 60&, (ata.disk(ata.selectDisk).sectors And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 61&, (U32Shr(ata.disk(ata.selectDisk).sectors, 16&) And &HFFFF&)
                ATA_IdentifyStoreWord ata.selectDisk, 80&, &H1E&
            Else
                ata.disk(ata.selectDisk).errorCode = 4&
            End If
            If ata.disk(ata.selectDisk).interrupt <> 0& Then
                ata_irq
            End If
        Case ATA_CMD_DIAGNOSTIC
            ata.disk(ata.selectDisk).errorCode = 0&
        Case ATA_CMD_DEVICE_RESET
            ata.disk(ata.selectDisk).errorCode = 4&
            ata.disk(0&).isreading = 0&
            ata.disk(0&).iswriting = 0&
            ata.disk(0&).regs.lba = 0&
            ata.disk(1&).isreading = 0&
            ata.disk(1&).iswriting = 0&
            ata.disk(1&).regs.lba = 0&
            ata.selectDisk = 0&
            ATA_LowerIRQ
            timing_timerDisable ata.timerNum
            ata.delay_irq = 0&
            If ata.disk(ata.selectDisk).interrupt <> 0& Then
                i8259_doirq ata.i8259, 6&
                ata.irq_pending = 1&
            End If
        Case ATA_CMD_INITIALIZE_PARAMS
            If ata.disk(ata.selectDisk).openFlag = 0& Then
                ata.disk(ata.selectDisk).errorCode = 4&
            Else
                ata.disk(ata.selectDisk).errorCode = 0&
                ata.disk(ata.selectDisk).spt = (ata.disk(ata.selectDisk).regs.sectors And 63&)
                ata.disk(ata.selectDisk).heads = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 24&) And &HF&) + 1&
                If ata.disk(ata.selectDisk).interrupt <> 0& Then
                    ata_irq
                End If
            End If
        Case &H41&, &H70&, &HEF&, &H40&, ATA_CMD_IDLE_IMMEDIATE
            ata.disk(ata.selectDisk).errorCode = 0&
            If ata.disk(ata.selectDisk).interrupt <> 0& Then
                ata_irq
            End If
        Case ATA_CMD_RECALIBRATE
            ata.disk(ata.selectDisk).errorCode = 0&
            If ata.disk(ata.selectDisk).interrupt <> 0& Then
                ata_irq
            End If
        Case ATA_CMD_READ_SECTORS, &H21&
            If ata.disk(ata.selectDisk).openFlag = 0& Then
                ata.disk(ata.selectDisk).errorCode = 4&
            Else
                ata.disk(ata.selectDisk).errorCode = 0&
                ata.disk(ata.selectDisk).iswriting = 0&
                ata.disk(ata.selectDisk).isreading = 1&
                ata.disk(ata.selectDisk).curreadsect = 0&
                If ata.disk(ata.selectDisk).regs.sectors = 0& Then
                    ata.disk(ata.selectDisk).targetsect = 256&
                Else
                    ata.disk(ata.selectDisk).targetsect = ata.disk(ata.selectDisk).regs.sectors
                End If
                If ata.disk(ata.selectDisk).lbamode <> 0& Then
                    ata.savelba = (ata.disk(ata.selectDisk).regs.lba And &HFFFFFF)
                Else
                    sector = (ata.disk(ata.selectDisk).regs.lba And 63&)
                    cyl = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 8&) And &HFFFF&)
                    head = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 24&) And &HF&)
                    ata.disk(ata.selectDisk).curcyl = cyl
                    ata.disk(ata.selectDisk).curhead = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 24&) And &HFF&)
                    ata.disk(ata.selectDisk).cursect = (ata.disk(ata.selectDisk).regs.lba And 63&)
                    ata.savelba = ATA_CHSToLBA(cyl, head, sector, ata.disk(ata.selectDisk).heads, ata.disk(ata.selectDisk).spt)
                End If
                ata_read_disk
            End If
        Case ATA_CMD_WRITE_SECTORS, &H31&
            If ata.disk(ata.selectDisk).openFlag = 0& Then
                ata.disk(ata.selectDisk).errorCode = 4&
            Else
                ata.disk(ata.selectDisk).errorCode = 0&
                ata.disk(ata.selectDisk).iswriting = 1&
                ata.disk(ata.selectDisk).isreading = 0&
                ata.disk(ata.selectDisk).curreadsect = 0&
                ata.disk(ata.selectDisk).buffer_pos = 0&
                If ata.disk(ata.selectDisk).regs.sectors = 0& Then
                    ata.disk(ata.selectDisk).targetsect = 256&
                Else
                    ata.disk(ata.selectDisk).targetsect = ata.disk(ata.selectDisk).regs.sectors
                End If
                If ata.disk(ata.selectDisk).lbamode <> 0& Then
                    ata.savelba = ata.disk(ata.selectDisk).regs.lba
                Else
                    sector = (ata.disk(ata.selectDisk).regs.lba And 63&)
                    cyl = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 8&) And &HFFFF&)
                    head = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 24&) And &HF&)
                    ata.disk(ata.selectDisk).curcyl = cyl
                    ata.disk(ata.selectDisk).curhead = (U32Shr(ata.disk(ata.selectDisk).regs.lba, 24&) And &HFF&)
                    ata.disk(ata.selectDisk).cursect = (ata.disk(ata.selectDisk).regs.lba And 63&)
                    ata.savelba = ATA_CHSToLBA(cyl, head, sector, ata.disk(ata.selectDisk).heads, ata.disk(ata.selectDisk).spt)
                End If
            End If
        Case Else
            debug_log DEBUG_DETAIL, "[ATA] Unimplemented command: 0x" & right$("00" & Hex$(ata.disk(ata.selectDisk).command), 2&)
    End Select
End Sub
Public Function ata_read_port(ByVal dummy As Long, ByVal portnum As Integer) As Byte
    Dim ret As Byte
    ret = 0&
    Select Case (portnum And &HFFFF&)
        Case ATA_PORT_DATA
            If ata.delay_irq <> 0& Then
                ata_read_port = 0&
                Exit Function
            End If
            If ata.disk(ata.selectDisk).buffer_pos >= 512& Then
                ata_read_port = 0&
                Exit Function
            End If
            ret = ata.disk(ata.selectDisk).buffer(ata.disk(ata.selectDisk).buffer_pos)
            ata.disk(ata.selectDisk).buffer_pos = ata.disk(ata.selectDisk).buffer_pos + 1&
            If ata.disk(ata.selectDisk).buffer_pos >= 512& Then
                If ata.disk(ata.selectDisk).curreadsect >= ata.disk(ata.selectDisk).targetsect Then
                    ata.disk(ata.selectDisk).isreading = 0&
                Else
                    ata_read_disk
                End If
            End If
        Case ATA_PORT_ERROR
            ret = ata.disk(ata.selectDisk).errorCode
        Case ATA_PORT_SECTORS
            ret = ata.disk(ata.selectDisk).regs.sectors
        Case ATA_PORT_LBA_LOW
            ret = CByte(ata.disk(ata.selectDisk).regs.lba And &HFF&)
        Case ATA_PORT_LBA_MID
            ret = CByte(U32Shr(ata.disk(ata.selectDisk).regs.lba, 8&) And &HFF&)
        Case ATA_PORT_LBA_HIGH
            ret = CByte(U32Shr(ata.disk(ata.selectDisk).regs.lba, 16&) And &HFF&)
        Case ATA_PORT_DRIVE
            ret = (&HA0& Or (U32Shr(ata.disk(ata.selectDisk).regs.lba, 24&) And &HF&) Or ((ata.selectDisk And 1&) * &H10&) Or ((ata.disk(ata.selectDisk).lbamode And 1&) * &H40&))
        Case ATA_PORT_STATUS
            ret = ata_gen_status()
            If ata.delay_irq = 0& Then
                ATA_LowerIRQ
            End If
        Case ATA_PORT_ALTERNATE
            ret = ata_gen_status()
    End Select
    ata_read_port = ret
End Function
Public Sub ata_write_port(ByVal dummy As Long, ByVal portnum As Integer, ByVal value As Byte)
    Dim old_control As Byte

    Select Case (portnum And &HFFFF&)
        Case ATA_PORT_DATA
            debug_log DEBUG_INFO, "ATA 8-bit data write"
        Case ATA_PORT_FEATURES
            ata.disk(ata.selectDisk).regs.features = value
        Case ATA_PORT_SECTORS
            ata.disk(ata.selectDisk).regs.sectors = value
        Case ATA_PORT_LBA_LOW
            ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFFFFFF00) Or (value And &HFF&)
        Case ATA_PORT_LBA_MID
            ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFFFF00FF) Or U32Shl((value And &HFF&), 8&)
        Case ATA_PORT_LBA_HIGH
            ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFF00FFFF) Or U32Shl((value And &HFF&), 16&)
        Case ATA_PORT_DRIVE
            ata.selectDisk = (U32Shr(value, 4&) And 1&)
            ata.disk(ata.selectDisk).lbamode = (U32Shr(value, 6&) And 1&)
            ata.disk(ata.selectDisk).regs.lba = (ata.disk(ata.selectDisk).regs.lba And &HFFFFFF) Or U32Shl((value And &HF&), 24&)
        Case ATA_PORT_COMMAND
            ata.disk(ata.selectDisk).command = value
            ata_command_process
        Case ATA_PORT_ALTERNATE
            old_control = ata.control
            ata.control = value
            ATA_SetInterruptEnable (((U32Shr(value, 1&) And 1&) Xor 1&) And 1&)
            If ((old_control And &H4&) = 0&) And ((value And &H4&) <> 0&) Then
                ata.inreset = 1&
                ata.delay_irq = 0&
                ata.dscflag = 0&
                timing_timerDisable ata.timerNum
                timing_timerDisable ata.resettimer
                ata.disk(0&).isreading = 0&
                ata.disk(0&).iswriting = 0&
                ata.disk(1&).isreading = 0&
                ata.disk(1&).iswriting = 0&
                ATA_LowerIRQ
            ElseIf ((old_control And &H4&) <> 0&) And ((value And &H4&) = 0&) Then
                timing_timerEnable ata.resettimer
            End If
    End Select
End Sub
Public Function ata_read_data(ByVal dummy As Long, ByVal portnum As Integer) As Long
    Dim ret As Long
    If (ata.delay_irq <> 0&) Or (ata.disk(ata.selectDisk).buffer_pos >= 512&) Or (ata.disk(ata.selectDisk).isreading = 0&) Then
        ata_read_data = 0&
        Exit Function
    End If
    ret = CLng(ata.disk(ata.selectDisk).buffer(ata.disk(ata.selectDisk).buffer_pos))
    ret = ret Or (CLng(ata.disk(ata.selectDisk).buffer(ata.disk(ata.selectDisk).buffer_pos + 1&)) * &H100&)
    ata.disk(ata.selectDisk).buffer_pos = ata.disk(ata.selectDisk).buffer_pos + 2&
    If ata.disk(ata.selectDisk).buffer_pos >= 512& Then
        If ata.disk(ata.selectDisk).curreadsect >= ata.disk(ata.selectDisk).targetsect Then
            ata.disk(ata.selectDisk).isreading = 0&
        Else
            ata_read_disk
        End If
    End If
    ata_read_data = (ret And &HFFFF&)
End Function
Public Sub ata_write_data(ByVal dummy As Long, ByVal portnum As Integer, ByVal value As Long)
    If (ata.disk(ata.selectDisk).buffer_pos >= 512&) Or (ata.disk(ata.selectDisk).iswriting = 0&) Then Exit Sub
    ata.disk(ata.selectDisk).buffer(ata.disk(ata.selectDisk).buffer_pos) = CByte(value And &HFF&)
    ata.disk(ata.selectDisk).buffer(ata.disk(ata.selectDisk).buffer_pos + 1&) = CByte((value \ &H100&) And &HFF&)
    ata.disk(ata.selectDisk).buffer_pos = ata.disk(ata.selectDisk).buffer_pos + 2&
    If ata.disk(ata.selectDisk).buffer_pos >= 512& Then
        ata_write_disk
        If ata.disk(ata.selectDisk).curreadsect >= ata.disk(ata.selectDisk).targetsect Then
            ata.disk(ata.selectDisk).iswriting = 0&
        End If
    End If
End Sub
Public Sub ata_insert_disk(ByVal diskSel As Long, ByVal filename As String)
    Dim fn As Integer
    Dim chs_total As Long
    If Not ATA_IsValidDisk(diskSel) Then Exit Sub
    ATA_CloseDisk diskSel
    fn = FreeFile
    On Error GoTo InsertFail
    Open filename For Binary Access Read Write As #fn
    ata.disk(diskSel).fileNum = fn
    ata.disk(diskSel).openFlag = 1&
    ata.disk(diskSel).sectors = (LOF(fn) \ 512&)
    ata.disk(diskSel).spt = 63&
    ata.disk(diskSel).heads = 16&
    ata.disk(diskSel).cylinders = (ata.disk(diskSel).sectors \ (16& * 63&))
    chs_total = ATA_U32Mul(ATA_U32Mul(ata.disk(diskSel).cylinders, ata.disk(diskSel).heads), ata.disk(diskSel).spt)
    If chs_total > ata.disk(diskSel).sectors Then
        ata.disk(diskSel).cylinders = ata.disk(diskSel).cylinders - 1&
    End If
    debug_log DEBUG_INFO, "[ATA] Inserted disk on " & IIf(diskSel <> 0&, "slave", "master") & " channel: " & filename
    Exit Sub
InsertFail:
    On Error Resume Next
    If fn <> 0& Then Close #fn
    On Error GoTo 0
    ata.disk(diskSel).fileNum = 0&
    ata.disk(diskSel).openFlag = 0&
End Sub
Public Sub ata_init(ByVal i8259_slave As Long)
    ata.selectDisk = 0&
    ata.i8259 = i8259_slave
    ata.control = 0&
    ata.delay_irq = 0&
    ata.inreset = 0&
    ata.dscflag = ATA_STATUS_DSC
    ata.irq_pending = 0&
    ata.irq_drive = 0&
    ATA_ResetDrive 0&
    ATA_ResetDrive 1&
    ATA_SetInterruptEnable 1&
    ata.timerNum = timing_addTimer(TIMER_CB_ATA_DELAYED_IRQ, 0&, 50000, TIMING_DISABLED)
    ata.resettimer = timing_addTimer(TIMER_CB_ATA_RESET, 0&, 4&, TIMING_DISABLED)
    ports_cbRegister &H1F0&, 1&, PORTS_CB_ATA_PORT, PORTS_CB_ATA_DATA, PORTS_CB_ATA_PORT, PORTS_CB_ATA_DATA, 0&
    ports_cbRegister &H1F1&, 7&, PORTS_CB_ATA_PORT, PORTS_CB_NONE, PORTS_CB_ATA_PORT, PORTS_CB_NONE, 0&
    ports_cbRegister &H3F6&, 1&, PORTS_CB_ATA_PORT, PORTS_CB_NONE, PORTS_CB_ATA_PORT, PORTS_CB_NONE, 0&
End Sub

Public Function ata_get_inserted_count() As Byte
    Dim disk As Long
    Dim count As Long

    For disk = 0& To 1&
        If ata.disk(disk).openFlag <> 0& Then count = count + 1&
    Next disk

    ata_get_inserted_count = CByte(count And &HFF&)
End Function

