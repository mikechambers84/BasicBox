Attribute VB_Name = "modConfig"
Option Explicit

Public Const STR_TITLE As String = "BasicBox"
Public Const STR_VERSION As String = "0.5.0"

Public Const VIDEO_CARD_VGA As Long = 3&
Public Const VIDEO_CARD_ET4000 As Long = 4&

Public Const SAMPLE_RATE As Long = 44100
Public Const SAMPLE_BUFFER As Long = 4410&
Public Const DEFAULT_GUEST_RAM_MB As Long = 64&
Public Const MIN_GUEST_RAM_MB As Long = 1&
Public Const MAX_GUEST_RAM_MB As Long = 256&

Public Const AUDIO_TIMING_FAST As Long = 1&
Public Const AUDIO_TIMING_NORMAL As Long = 2&

Public running As Byte
Public videocard As Byte
Public showMIPS As Byte
Public speedarg As Double
Public speed As Double
Public baudrate As Long
Public guestRamMB As Long
Public useMachine As String
Public cmosOverride As String
Public vga_lockFPS As Double
