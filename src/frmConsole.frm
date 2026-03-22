VERSION 5.00
Begin VB.Form frmConsole 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BasicBox"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   7995
   Icon            =   "frmConsole.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu itmFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFloppy 
      Caption         =   "Floppy"
      Begin VB.Menu itmFloppyInsert0 
         Caption         =   "Insert/change first floppy..."
      End
      Begin VB.Menu itmFloppyInsert0WP 
         Caption         =   "Insert/change first floppy... (Write-protected)"
      End
      Begin VB.Menu itmFloppyEject0 
         Caption         =   "Eject first floppy"
      End
      Begin VB.Menu itmDash 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu itmFloppyInsert1 
         Caption         =   "Insert/change second floppy..."
      End
      Begin VB.Menu itmFloppyInsert1WP 
         Caption         =   "Insert/change second floppy... (Write-protected)"
      End
      Begin VB.Menu itmFloppyEject1 
         Caption         =   "Eject second floppy"
      End
   End
   Begin VB.Menu mnuSCSI 
      Caption         =   "SCSI"
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 0)"
         Index           =   0
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 0)"
         Index           =   0
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 1)"
         Index           =   1
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 1)"
         Index           =   1
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 2)"
         Index           =   2
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 2)"
         Index           =   2
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 3)"
         Index           =   3
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 3)"
         Index           =   3
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 4)"
         Index           =   4
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 4)"
         Index           =   4
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 5)"
         Index           =   5
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 5)"
         Index           =   5
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 6)"
         Index           =   6
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 6)"
         Index           =   6
      End
      Begin VB.Menu itmSCSICD 
         Caption         =   "Insert/change SCSI-CD (ID 7)"
         Index           =   7
      End
      Begin VB.Menu itmEjectSCSICD 
         Caption         =   "Eject SCSI-CD (ID 7)"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private exitRequested As Byte

Private Function Console_ConfirmExit() As Boolean
    Dim res As Long

    If exitRequested <> 0& Then
        Console_ConfirmExit = True
        Exit Function
    End If

    res = MsgBox("Exit BasicBox, are you sure?", vbYesNo Or vbQuestion, "Exit")
    If res = vbNo Then Exit Function

    exitRequested = 1&
    Console_ConfirmExit = True
End Function

Private Sub Form_Activate()
    Console_FormActivate
End Sub

Private Sub Form_Deactivate()
    Console_FormDeactivate
End Sub

Private Sub Form_Load()
    ' Startup is handled by Sub Main in modMain.
    exitRequested = 0&
    mnuSCSI.visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormCode Then Exit Sub

    If Console_ConfirmExit() = False Then
        Cancel = 1
        Exit Sub
    End If

    running = 0&
    Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Console_FormUnload
    running = 0&
End Sub

Private Sub itmFileExit_Click()
    If Console_ConfirmExit() = False Then Exit Sub
    running = 0&
End Sub

Private Sub itmFloppyEject0_Click()
    menus_ejectFloppy0
End Sub

Private Sub itmFloppyEject1_Click()
    menus_ejectFloppy1
End Sub

Private Sub itmFloppyInsert0_Click()
    menus_changeFloppy0
End Sub

Private Sub itmFloppyInsert0WP_Click()
    menus_changeFloppy0WP
End Sub

Private Sub itmFloppyInsert1_Click()
    menus_changeFloppy1
End Sub

Private Sub itmFloppyInsert1WP_Click()
    menus_changeFloppy1WP
End Sub

Private Sub itmSCSICD_Click(index As Integer)
    menus_changeScsiCD index
End Sub

Private Sub itmEjectSCSICD_Click(index As Integer)
    menus_ejectScsiCD index
End Sub
