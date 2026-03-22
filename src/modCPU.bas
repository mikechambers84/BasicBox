Attribute VB_Name = "modCPU"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)

Private Const CPU_REG_ES As Long = 0&
Private Const CPU_REG_CS As Long = 1&
Private Const CPU_REG_SS As Long = 2&
Private Const CPU_REG_DS As Long = 3&
Private Const CPU_REG_FS As Long = 4&
Private Const CPU_REG_GS As Long = 5&

Private Const CPU_REG_EAX As Long = 0&
Private Const CPU_REG_ECX As Long = 1&
Private Const CPU_REG_EDX As Long = 2&
Private Const CPU_REG_EBX As Long = 3&
Private Const CPU_REG_ESP As Long = 4&
Private Const CPU_REG_EBP As Long = 5&
Private Const CPU_REG_ESI As Long = 6&
Private Const CPU_REG_EDI As Long = 7&

Private Const EFLAGS_CF As Long = &H1&
Private Const EFLAGS_PF As Long = &H4&
Private Const EFLAGS_AF As Long = &H10&
Private Const EFLAGS_ZF As Long = &H40&
Private Const EFLAGS_SF As Long = &H80&
Private Const EFLAGS_TF As Long = &H100&
Private Const EFLAGS_IF As Long = &H200&
Private Const EFLAGS_DF As Long = &H400&
Private Const EFLAGS_OF As Long = &H800&
Private Const EFLAGS_NT As Long = &H4000&
Private Const EFLAGS_IOPL As Long = &H3000&
Private Const EFLAGS_RF As Long = &H10000
Private Const EFLAGS_VM As Long = &H20000
Private Const EFLAGS_AC As Long = &H40000
Private Const EFLAGS_ID As Long = &H200000
Private Const CR0_EM As Long = &H4&
Private Const CR0_TS As Long = &H8&
Private Const CR0_NE As Long = &H20&
Private cpu_parityTable(0& To 255&) As Byte
Private cpu_parityInit As Byte

Public Const INT_SOURCE_EXCEPTION As Long = 0&
Public Const INT_SOURCE_SOFTWARE As Long = 1&
Public Const INT_SOURCE_HARDWARE As Long = 2&
Public Const INT_SOURCE_INT3 As Long = 3&
Public Const INT_SOURCE_INTO As Long = 4&

Private Const TASK_SWITCH_REASON_CALL As Long = 0&
Private Const TASK_SWITCH_REASON_GATE As Long = 1&
Private Const TASK_SWITCH_REASON_IRET As Long = 2&
Private Const TASK_SWITCH_REASON_JMP As Long = 3&

Public Const CPU_INTCB_NONE As Long = 0&

Private Type CPU_SEGDESC_t
    addr As Long
    access As Long
    flags As Long
    dpl As Long
End Type

Private Type CPU_RETURNCS_t
    selector As Long
    target_cpl As Long
    outer As Long
End Type

Private Type CPU_GATEDESC_t
    addr As Long
    target_selector As Long
    offset As Long
    access As Long
    flags As Long
    typeVal As Long
    dpl As Long
    present As Long
    param_count As Long
End Type

Private Type CPU_CODETARGET_t
    selector As Long
    target_cpl As Long
    outer As Long
    conforming As Long
End Type

Private Const OPH_ILLEGAL As Long = 0&
Private Const OPH_EXT_0F As Long = 1&
Private Const OPH_90 As Long = 2&
Private Const OPH_CC As Long = 3&
Private Const OPH_CD As Long = 4&
Private Const OPH_D0 As Long = 5&
Private Const OPH_D1 As Long = 6&
Private Const OPH_D2 As Long = 7&
Private Const OPH_D3 As Long = 8&
Private Const OPH_F1 As Long = 9&
Private Const OPH_F4 As Long = 10&
Private Const OPH_F5 As Long = 11&
Private Const OPH_F6 As Long = 12&
Private Const OPH_F7 As Long = 13&
Private Const OPH_F8 As Long = 14&
Private Const OPH_F9 As Long = 15&
Private Const OPH_FA As Long = 16&
Private Const OPH_FB As Long = 17&
Private Const OPH_FC As Long = 18&
Private Const OPH_FD As Long = 19&
Private Const OPH_FE As Long = 20&
Private Const OPH_FF As Long = 21&
Private Const OPH_E0 As Long = 22&
Private Const OPH_E1 As Long = 23&
Private Const OPH_E2 As Long = 24&
Private Const OPH_E3 As Long = 25&
Private Const OPH_E4 As Long = 26&
Private Const OPH_E5 As Long = 27&
Private Const OPH_E6 As Long = 28&
Private Const OPH_E7 As Long = 29&
Private Const OPH_E8 As Long = 30&
Private Const OPH_E9 As Long = 31&
Private Const OPH_EA As Long = 32&
Private Const OPH_EB As Long = 33&
Private Const OPH_EC As Long = 34&
Private Const OPH_ED As Long = 35&
Private Const OPH_EE As Long = 36&
Private Const OPH_EF As Long = 37&
Private Const OPH_68 As Long = 38&
Private Const OPH_69 As Long = 39&
Private Const OPH_6A As Long = 40&
Private Const OPH_6B As Long = 41&
Private Const OPH_C0 As Long = 42&
Private Const OPH_C1 As Long = 43&
Private Const OPH_C6 As Long = 44&
Private Const OPH_C7 As Long = 45&
Private Const OPH_C2 As Long = 46&
Private Const OPH_C3 As Long = 47&
Private Const OPH_C8 As Long = 48&
Private Const OPH_C9 As Long = 49&
Private Const OPH_CA As Long = 50&
Private Const OPH_CB As Long = 51&
Private Const OPH_CE As Long = 52&
Private Const OPH_CF As Long = 53&
Private Const OPH_C4 As Long = 54&
Private Const OPH_C5 As Long = 55&
Private Const OPH_A8 As Long = 56&
Private Const OPH_A9 As Long = 57&
Private Const OPH_B0 As Long = 58&
Private Const OPH_B1 As Long = 59&
Private Const OPH_B2 As Long = 60&
Private Const OPH_B3 As Long = 61&
Private Const OPH_B4 As Long = 62&
Private Const OPH_B5 As Long = 63&
Private Const OPH_B6 As Long = 64&
Private Const OPH_B7 As Long = 65&
Private Const OPH_MOV_REGIMM As Long = 66&
Private Const OPH_AA As Long = 67&
Private Const OPH_AB As Long = 68&
Private Const OPH_AC As Long = 69&
Private Const OPH_AD As Long = 70&
Private Const OPH_AE As Long = 71&
Private Const OPH_AF As Long = 72&
Private Const OPH_A0 As Long = 73&
Private Const OPH_A1 As Long = 74&
Private Const OPH_A2 As Long = 75&
Private Const OPH_A3 As Long = 76&
Private Const OPH_A4 As Long = 77&
Private Const OPH_A5 As Long = 78&
Private Const OPH_A6 As Long = 79&
Private Const OPH_A7 As Long = 80&
Private Const OPH_98 As Long = 81&
Private Const OPH_99 As Long = 82&
Private Const OPH_9A As Long = 83&
Private Const OPH_9B As Long = 84&
Private Const OPH_9C As Long = 85&
Private Const OPH_9D As Long = 86&
Private Const OPH_9E As Long = 87&
Private Const OPH_9F As Long = 88&
Private Const OPH_91 As Long = 89&
Private Const OPH_92 As Long = 90&
Private Const OPH_93 As Long = 91&
Private Const OPH_94 As Long = 92&
Private Const OPH_95 As Long = 93&
Private Const OPH_96 As Long = 94&
Private Const OPH_97 As Long = 95&
Private Const OPH_8C As Long = 96&
Private Const OPH_8D As Long = 97&
Private Const OPH_8E As Long = 98&
Private Const OPH_8F As Long = 99&
Private Const OPH_88 As Long = 100&
Private Const OPH_89 As Long = 101&
Private Const OPH_8A As Long = 102&
Private Const OPH_8B As Long = 103&
Private Const OPH_84 As Long = 104&
Private Const OPH_85 As Long = 105&
Private Const OPH_86 As Long = 106&
Private Const OPH_87 As Long = 107&
Private Const OPH_80 As Long = 108&
Private Const OPH_81 As Long = 109&
Private Const OPH_82 As Long = 110&
Private Const OPH_83 As Long = 111&
Private Const OPH_70 As Long = 112&
Private Const OPH_71 As Long = 113&
Private Const OPH_72 As Long = 114&
Private Const OPH_73 As Long = 115&
Private Const OPH_74 As Long = 116&
Private Const OPH_75 As Long = 117&
Private Const OPH_76 As Long = 118&
Private Const OPH_77 As Long = 119&
Private Const OPH_78 As Long = 120&
Private Const OPH_79 As Long = 121&
Private Const OPH_7A As Long = 122&
Private Const OPH_7B As Long = 123&
Private Const OPH_7C As Long = 124&
Private Const OPH_7D As Long = 125&
Private Const OPH_7E As Long = 126&
Private Const OPH_7F As Long = 127&
Private Const OPH_6C As Long = 128&
Private Const OPH_6D As Long = 129&
Private Const OPH_6E As Long = 130&
Private Const OPH_6F As Long = 131&
Private Const OPH_60 As Long = 132&
Private Const OPH_61 As Long = 133&
Private Const OPH_62 As Long = 134&
Private Const OPH_63 As Long = 135&
Private Const OPH_PUSH_REG As Long = 136&
Private Const OPH_POP_REG As Long = 137&
Private Const OPH_INC_REG As Long = 138&
Private Const OPH_DEC_REG As Long = 139&
Private Const OPH_30 As Long = 140&
Private Const OPH_31 As Long = 141&
Private Const OPH_32 As Long = 142&
Private Const OPH_33 As Long = 143&
Private Const OPH_34 As Long = 144&
Private Const OPH_35 As Long = 145&
Private Const OPH_37 As Long = 146&
Private Const OPH_38 As Long = 147&
Private Const OPH_39 As Long = 148&
Private Const OPH_3A As Long = 149&
Private Const OPH_3B As Long = 150&
Private Const OPH_3C As Long = 151&
Private Const OPH_3D As Long = 152&
Private Const OPH_3F As Long = 153&
Private Const OPH_20 As Long = 154&
Private Const OPH_21 As Long = 155&
Private Const OPH_22 As Long = 156&
Private Const OPH_23 As Long = 157&
Private Const OPH_24 As Long = 158&
Private Const OPH_25 As Long = 159&
Private Const OPH_27 As Long = 160&
Private Const OPH_28 As Long = 161&
Private Const OPH_29 As Long = 162&
Private Const OPH_2A As Long = 163&
Private Const OPH_2B As Long = 164&
Private Const OPH_2C As Long = 165&
Private Const OPH_2D As Long = 166&
Private Const OPH_2F As Long = 167&
Private Const OPH_10 As Long = 168&
Private Const OPH_11 As Long = 169&
Private Const OPH_12 As Long = 170&
Private Const OPH_13 As Long = 171&
Private Const OPH_14 As Long = 172&
Private Const OPH_15 As Long = 173&
Private Const OPH_16 As Long = 174&
Private Const OPH_17 As Long = 175&
Private Const OPH_18 As Long = 176&
Private Const OPH_19 As Long = 177&
Private Const OPH_1A As Long = 178&
Private Const OPH_1B As Long = 179&
Private Const OPH_1C As Long = 180&
Private Const OPH_1D As Long = 181&
Private Const OPH_1E As Long = 182&
Private Const OPH_1F As Long = 183&
Private Const OPH_00 As Long = 184&
Private Const OPH_01 As Long = 185&
Private Const OPH_02 As Long = 186&
Private Const OPH_03 As Long = 187&
Private Const OPH_04 As Long = 188&
Private Const OPH_05 As Long = 189&
Private Const OPH_06 As Long = 190&
Private Const OPH_07 As Long = 191&
Private Const OPH_08 As Long = 192&
Private Const OPH_09 As Long = 193&
Private Const OPH_0A As Long = 194&
Private Const OPH_0B As Long = 195&
Private Const OPH_0C As Long = 196&
Private Const OPH_0D As Long = 197&
Private Const OPH_0E As Long = 198&
Private Const OPH_D4 As Long = 199&
Private Const OPH_D5 As Long = 200&
Private Const OPH_D6_D7 As Long = 201&
Private Const OPH_FPU As Long = 202&
Private Const OPH_EXT_ILLEGAL As Long = 0&
Private Const OPHX_00 As Long = &H0&
Private Const OPHX_01 As Long = &H1&
Private Const OPHX_02 As Long = &H2&
Private Const OPHX_03 As Long = &H3&
Private Const OPHX_06 As Long = &H6&
Private Const OPHX_08_09 As Long = &H8&
Private Const OPHX_20 As Long = &H20&
Private Const OPHX_21 As Long = &H21&
Private Const OPHX_22 As Long = &H22&
Private Const OPHX_23 As Long = &H23&
Private Const OPHX_24_26 As Long = &H24&
Private Const OPHX_30 As Long = &H30&
Private Const OPHX_31 As Long = &H31&
Private Const OPHX_40 As Long = &H40&
Private Const OPHX_41 As Long = &H41&
Private Const OPHX_42 As Long = &H42&
Private Const OPHX_43 As Long = &H43&
Private Const OPHX_44 As Long = &H44&
Private Const OPHX_45 As Long = &H45&
Private Const OPHX_46 As Long = &H46&
Private Const OPHX_47 As Long = &H47&
Private Const OPHX_48 As Long = &H48&
Private Const OPHX_49 As Long = &H49&
Private Const OPHX_4A As Long = &H4A&
Private Const OPHX_4B As Long = &H4B&
Private Const OPHX_4C As Long = &H4C&
Private Const OPHX_4D As Long = &H4D&
Private Const OPHX_4E As Long = &H4E&
Private Const OPHX_4F As Long = &H4F&
Private Const OPHX_80 As Long = &H80&
Private Const OPHX_81 As Long = &H81&
Private Const OPHX_82 As Long = &H82&
Private Const OPHX_83 As Long = &H83&
Private Const OPHX_84 As Long = &H84&
Private Const OPHX_85 As Long = &H85&
Private Const OPHX_86 As Long = &H86&
Private Const OPHX_87 As Long = &H87&
Private Const OPHX_88 As Long = &H88&
Private Const OPHX_89 As Long = &H89&
Private Const OPHX_8A As Long = &H8A&
Private Const OPHX_8B As Long = &H8B&
Private Const OPHX_8C As Long = &H8C&
Private Const OPHX_8D As Long = &H8D&
Private Const OPHX_8E As Long = &H8E&
Private Const OPHX_8F As Long = &H8F&
Private Const OPHX_90 As Long = &H90&
Private Const OPHX_91 As Long = &H91&
Private Const OPHX_92 As Long = &H92&
Private Const OPHX_93 As Long = &H93&
Private Const OPHX_94 As Long = &H94&
Private Const OPHX_95 As Long = &H95&
Private Const OPHX_96 As Long = &H96&
Private Const OPHX_97 As Long = &H97&
Private Const OPHX_98 As Long = &H98&
Private Const OPHX_99 As Long = &H99&
Private Const OPHX_9A As Long = &H9A&
Private Const OPHX_9B As Long = &H9B&
Private Const OPHX_9C As Long = &H9C&
Private Const OPHX_9D As Long = &H9D&
Private Const OPHX_9E As Long = &H9E&
Private Const OPHX_9F As Long = &H9F&
Private Const OPHX_A0 As Long = &HA0&
Private Const OPHX_A1 As Long = &HA1&
Private Const OPHX_A2 As Long = &HA2&
Private Const OPHX_A3 As Long = &HA3&
Private Const OPHX_A4_A5 As Long = &HA4&
Private Const OPHX_A6_B0 As Long = &HA6&
Private Const OPHX_A8 As Long = &HA8&
Private Const OPHX_A9 As Long = &HA9&
Private Const OPHX_AA As Long = &HAA&
Private Const OPHX_AB As Long = &HAB&
Private Const OPHX_AC_AD As Long = &HAC&
Private Const OPHX_AF As Long = &HAF&
Private Const OPHX_B1 As Long = &HB1&
Private Const OPHX_B2_B4_B5 As Long = &HB2&
Private Const OPHX_B3 As Long = &HB3&
Private Const OPHX_B6 As Long = &HB6&
Private Const OPHX_B7 As Long = &HB7&
Private Const OPHX_BA As Long = &HBA&
Private Const OPHX_BB As Long = &HBB&
Private Const OPHX_BC As Long = &HBC&
Private Const OPHX_BD As Long = &HBD&
Private Const OPHX_BE As Long = &HBE&
Private Const OPHX_BF As Long = &HBF&
Private Const OPHX_C0 As Long = &HC0&
Private Const OPHX_C1 As Long = &HC1&
Private Const OPHX_C8_CF As Long = &HC8&

Private opcode_map_primary(0& To 255&) As Long
Private opcode_map_ext(0& To 255&) As Long
Private opcode_maps_ready As Byte

Private cpu_firstip As Long

Public showops As Long

Private Sub cpu_buildOpcodeMaps()
    Dim i As Long

    If opcode_maps_ready <> 0& Then Exit Sub

    For i = 0& To 255&
        opcode_map_primary(i) = OPH_ILLEGAL
        opcode_map_ext(i) = OPH_EXT_ILLEGAL
    Next i

    opcode_map_ext(&H0&) = OPHX_00
    opcode_map_ext(&H1&) = OPHX_01
    opcode_map_ext(&H2&) = OPHX_02
    opcode_map_ext(&H3&) = OPHX_03
    opcode_map_ext(&H6&) = OPHX_06
    opcode_map_ext(&H8&) = OPHX_08_09
    opcode_map_ext(&H9&) = OPHX_08_09
    opcode_map_ext(&H20&) = OPHX_20
    opcode_map_ext(&H21&) = OPHX_21
    opcode_map_ext(&H22&) = OPHX_22
    opcode_map_ext(&H23&) = OPHX_23
    opcode_map_ext(&H24&) = OPHX_24_26
    opcode_map_ext(&H26&) = OPHX_24_26
    opcode_map_ext(&H30&) = OPHX_30
    opcode_map_ext(&H31&) = OPHX_31
    opcode_map_ext(&H40&) = OPHX_40
    opcode_map_ext(&H41&) = OPHX_41
    opcode_map_ext(&H42&) = OPHX_42
    opcode_map_ext(&H43&) = OPHX_43
    opcode_map_ext(&H44&) = OPHX_44
    opcode_map_ext(&H45&) = OPHX_45
    opcode_map_ext(&H46&) = OPHX_46
    opcode_map_ext(&H47&) = OPHX_47
    opcode_map_ext(&H48&) = OPHX_48
    opcode_map_ext(&H49&) = OPHX_49
    opcode_map_ext(&H4A&) = OPHX_4A
    opcode_map_ext(&H4B&) = OPHX_4B
    opcode_map_ext(&H4C&) = OPHX_4C
    opcode_map_ext(&H4D&) = OPHX_4D
    opcode_map_ext(&H4E&) = OPHX_4E
    opcode_map_ext(&H4F&) = OPHX_4F
    opcode_map_ext(&H80&) = OPHX_80
    opcode_map_ext(&H81&) = OPHX_81
    opcode_map_ext(&H82&) = OPHX_82
    opcode_map_ext(&H83&) = OPHX_83
    opcode_map_ext(&H84&) = OPHX_84
    opcode_map_ext(&H85&) = OPHX_85
    opcode_map_ext(&H86&) = OPHX_86
    opcode_map_ext(&H87&) = OPHX_87
    opcode_map_ext(&H88&) = OPHX_88
    opcode_map_ext(&H89&) = OPHX_89
    opcode_map_ext(&H8A&) = OPHX_8A
    opcode_map_ext(&H8B&) = OPHX_8B
    opcode_map_ext(&H8C&) = OPHX_8C
    opcode_map_ext(&H8D&) = OPHX_8D
    opcode_map_ext(&H8E&) = OPHX_8E
    opcode_map_ext(&H8F&) = OPHX_8F
    opcode_map_ext(&H90&) = OPHX_90
    opcode_map_ext(&H91&) = OPHX_91
    opcode_map_ext(&H92&) = OPHX_92
    opcode_map_ext(&H93&) = OPHX_93
    opcode_map_ext(&H94&) = OPHX_94
    opcode_map_ext(&H95&) = OPHX_95
    opcode_map_ext(&H96&) = OPHX_96
    opcode_map_ext(&H97&) = OPHX_97
    opcode_map_ext(&H98&) = OPHX_98
    opcode_map_ext(&H99&) = OPHX_99
    opcode_map_ext(&H9A&) = OPHX_9A
    opcode_map_ext(&H9B&) = OPHX_9B
    opcode_map_ext(&H9C&) = OPHX_9C
    opcode_map_ext(&H9D&) = OPHX_9D
    opcode_map_ext(&H9E&) = OPHX_9E
    opcode_map_ext(&H9F&) = OPHX_9F
    opcode_map_ext(&HA0&) = OPHX_A0
    opcode_map_ext(&HA1&) = OPHX_A1
    opcode_map_ext(&HA2&) = OPHX_A2
    opcode_map_ext(&HA3&) = OPHX_A3
    opcode_map_ext(&HA4&) = OPHX_A4_A5
    opcode_map_ext(&HA5&) = OPHX_A4_A5
    opcode_map_ext(&HA6&) = OPHX_A6_B0
    opcode_map_ext(&HA8&) = OPHX_A8
    opcode_map_ext(&HA9&) = OPHX_A9
    opcode_map_ext(&HAA&) = OPHX_AA
    opcode_map_ext(&HAB&) = OPHX_AB
    opcode_map_ext(&HAC&) = OPHX_AC_AD
    opcode_map_ext(&HAD&) = OPHX_AC_AD
    opcode_map_ext(&HAF&) = OPHX_AF
    opcode_map_ext(&HB0&) = OPHX_A6_B0
    opcode_map_ext(&HB1&) = OPHX_B1
    opcode_map_ext(&HB2&) = OPHX_B2_B4_B5
    opcode_map_ext(&HB3&) = OPHX_B3
    opcode_map_ext(&HB4&) = OPHX_B2_B4_B5
    opcode_map_ext(&HB5&) = OPHX_B2_B4_B5
    opcode_map_ext(&HB6&) = OPHX_B6
    opcode_map_ext(&HB7&) = OPHX_B7
    opcode_map_ext(&HBA&) = OPHX_BA
    opcode_map_ext(&HBB&) = OPHX_BB
    opcode_map_ext(&HBC&) = OPHX_BC
    opcode_map_ext(&HBD&) = OPHX_BD
    opcode_map_ext(&HBE&) = OPHX_BE
    opcode_map_ext(&HBF&) = OPHX_BF
    opcode_map_ext(&HC0&) = OPHX_C0
    opcode_map_ext(&HC1&) = OPHX_C1
    opcode_map_ext(&HC8&) = OPHX_C8_CF
    opcode_map_ext(&HC9&) = OPHX_C8_CF
    opcode_map_ext(&HCA&) = OPHX_C8_CF
    opcode_map_ext(&HCB&) = OPHX_C8_CF
    opcode_map_ext(&HCC&) = OPHX_C8_CF
    opcode_map_ext(&HCD&) = OPHX_C8_CF
    opcode_map_ext(&HCE&) = OPHX_C8_CF
    opcode_map_ext(&HCF&) = OPHX_C8_CF

    opcode_map_primary(&H0&) = OPH_00
    opcode_map_primary(&H1&) = OPH_01
    opcode_map_primary(&H2&) = OPH_02
    opcode_map_primary(&H3&) = OPH_03
    opcode_map_primary(&H4&) = OPH_04
    opcode_map_primary(&H5&) = OPH_05
    opcode_map_primary(&H6&) = OPH_06
    opcode_map_primary(&H7&) = OPH_07
    opcode_map_primary(&H8&) = OPH_08
    opcode_map_primary(&H9&) = OPH_09
    opcode_map_primary(&HA&) = OPH_0A
    opcode_map_primary(&HB&) = OPH_0B
    opcode_map_primary(&HC&) = OPH_0C
    opcode_map_primary(&HD&) = OPH_0D
    opcode_map_primary(&HE&) = OPH_0E
    opcode_map_primary(&HF&) = OPH_EXT_0F
    opcode_map_primary(&H90&) = OPH_90
    opcode_map_primary(&H91&) = OPH_91
    opcode_map_primary(&H92&) = OPH_92
    opcode_map_primary(&H93&) = OPH_93
    opcode_map_primary(&H94&) = OPH_94
    opcode_map_primary(&H95&) = OPH_95
    opcode_map_primary(&H96&) = OPH_96
    opcode_map_primary(&H97&) = OPH_97
    opcode_map_primary(&H8C&) = OPH_8C
    opcode_map_primary(&H8D&) = OPH_8D
    opcode_map_primary(&H8E&) = OPH_8E
    opcode_map_primary(&H8F&) = OPH_8F
    opcode_map_primary(&H88&) = OPH_88
    opcode_map_primary(&H89&) = OPH_89
    opcode_map_primary(&H8A&) = OPH_8A
    opcode_map_primary(&H8B&) = OPH_8B
    opcode_map_primary(&H84&) = OPH_84
    opcode_map_primary(&H85&) = OPH_85
    opcode_map_primary(&H86&) = OPH_86
    opcode_map_primary(&H87&) = OPH_87
    opcode_map_primary(&H80&) = OPH_80
    opcode_map_primary(&H81&) = OPH_81
    opcode_map_primary(&H82&) = OPH_82
    opcode_map_primary(&H83&) = OPH_83
    opcode_map_primary(&H70&) = OPH_70
    opcode_map_primary(&H71&) = OPH_71
    opcode_map_primary(&H72&) = OPH_72
    opcode_map_primary(&H73&) = OPH_73
    opcode_map_primary(&H74&) = OPH_74
    opcode_map_primary(&H75&) = OPH_75
    opcode_map_primary(&H76&) = OPH_76
    opcode_map_primary(&H77&) = OPH_77
    opcode_map_primary(&H78&) = OPH_78
    opcode_map_primary(&H79&) = OPH_79
    opcode_map_primary(&H7A&) = OPH_7A
    opcode_map_primary(&H7B&) = OPH_7B
    opcode_map_primary(&H7C&) = OPH_7C
    opcode_map_primary(&H7D&) = OPH_7D
    opcode_map_primary(&H7E&) = OPH_7E
    opcode_map_primary(&H7F&) = OPH_7F
    opcode_map_primary(&H6C&) = OPH_6C
    opcode_map_primary(&H6D&) = OPH_6D
    opcode_map_primary(&H6E&) = OPH_6E
    opcode_map_primary(&H6F&) = OPH_6F
    opcode_map_primary(&H60&) = OPH_60
    opcode_map_primary(&H61&) = OPH_61
    opcode_map_primary(&H62&) = OPH_62
    opcode_map_primary(&H63&) = OPH_63
    opcode_map_primary(&H50&) = OPH_PUSH_REG
    opcode_map_primary(&H51&) = OPH_PUSH_REG
    opcode_map_primary(&H52&) = OPH_PUSH_REG
    opcode_map_primary(&H53&) = OPH_PUSH_REG
    opcode_map_primary(&H54&) = OPH_PUSH_REG
    opcode_map_primary(&H55&) = OPH_PUSH_REG
    opcode_map_primary(&H56&) = OPH_PUSH_REG
    opcode_map_primary(&H57&) = OPH_PUSH_REG
    opcode_map_primary(&H58&) = OPH_POP_REG
    opcode_map_primary(&H59&) = OPH_POP_REG
    opcode_map_primary(&H5A&) = OPH_POP_REG
    opcode_map_primary(&H5B&) = OPH_POP_REG
    opcode_map_primary(&H5C&) = OPH_POP_REG
    opcode_map_primary(&H5D&) = OPH_POP_REG
    opcode_map_primary(&H5E&) = OPH_POP_REG
    opcode_map_primary(&H5F&) = OPH_POP_REG
    opcode_map_primary(&H40&) = OPH_INC_REG
    opcode_map_primary(&H41&) = OPH_INC_REG
    opcode_map_primary(&H42&) = OPH_INC_REG
    opcode_map_primary(&H43&) = OPH_INC_REG
    opcode_map_primary(&H44&) = OPH_INC_REG
    opcode_map_primary(&H45&) = OPH_INC_REG
    opcode_map_primary(&H46&) = OPH_INC_REG
    opcode_map_primary(&H47&) = OPH_INC_REG
    opcode_map_primary(&H48&) = OPH_DEC_REG
    opcode_map_primary(&H49&) = OPH_DEC_REG
    opcode_map_primary(&H4A&) = OPH_DEC_REG
    opcode_map_primary(&H4B&) = OPH_DEC_REG
    opcode_map_primary(&H4C&) = OPH_DEC_REG
    opcode_map_primary(&H4D&) = OPH_DEC_REG
    opcode_map_primary(&H4E&) = OPH_DEC_REG
    opcode_map_primary(&H4F&) = OPH_DEC_REG
    opcode_map_primary(&H30&) = OPH_30
    opcode_map_primary(&H31&) = OPH_31
    opcode_map_primary(&H32&) = OPH_32
    opcode_map_primary(&H33&) = OPH_33
    opcode_map_primary(&H34&) = OPH_34
    opcode_map_primary(&H35&) = OPH_35
    opcode_map_primary(&H37&) = OPH_37
    opcode_map_primary(&H38&) = OPH_38
    opcode_map_primary(&H39&) = OPH_39
    opcode_map_primary(&H3A&) = OPH_3A
    opcode_map_primary(&H3B&) = OPH_3B
    opcode_map_primary(&H3C&) = OPH_3C
    opcode_map_primary(&H3D&) = OPH_3D
    opcode_map_primary(&H3F&) = OPH_3F
    opcode_map_primary(&H20&) = OPH_20
    opcode_map_primary(&H21&) = OPH_21
    opcode_map_primary(&H22&) = OPH_22
    opcode_map_primary(&H23&) = OPH_23
    opcode_map_primary(&H24&) = OPH_24
    opcode_map_primary(&H25&) = OPH_25
    opcode_map_primary(&H27&) = OPH_27
    opcode_map_primary(&H28&) = OPH_28
    opcode_map_primary(&H29&) = OPH_29
    opcode_map_primary(&H2A&) = OPH_2A
    opcode_map_primary(&H2B&) = OPH_2B
    opcode_map_primary(&H2C&) = OPH_2C
    opcode_map_primary(&H2D&) = OPH_2D
    opcode_map_primary(&H2F&) = OPH_2F
    opcode_map_primary(&H10&) = OPH_10
    opcode_map_primary(&H11&) = OPH_11
    opcode_map_primary(&H12&) = OPH_12
    opcode_map_primary(&H13&) = OPH_13
    opcode_map_primary(&H14&) = OPH_14
    opcode_map_primary(&H15&) = OPH_15
    opcode_map_primary(&H16&) = OPH_16
    opcode_map_primary(&H17&) = OPH_17
    opcode_map_primary(&H18&) = OPH_18
    opcode_map_primary(&H19&) = OPH_19
    opcode_map_primary(&H1A&) = OPH_1A
    opcode_map_primary(&H1B&) = OPH_1B
    opcode_map_primary(&H1C&) = OPH_1C
    opcode_map_primary(&H1D&) = OPH_1D
    opcode_map_primary(&H1E&) = OPH_1E
    opcode_map_primary(&H1F&) = OPH_1F
    opcode_map_primary(&H98&) = OPH_98
    opcode_map_primary(&H99&) = OPH_99
    opcode_map_primary(&H9A&) = OPH_9A
    opcode_map_primary(&H9B&) = OPH_9B
    opcode_map_primary(&H9C&) = OPH_9C
    opcode_map_primary(&H9D&) = OPH_9D
    opcode_map_primary(&H9E&) = OPH_9E
    opcode_map_primary(&H9F&) = OPH_9F
    opcode_map_primary(&HCC&) = OPH_CC
    opcode_map_primary(&HCD&) = OPH_CD
    opcode_map_primary(&HD0&) = OPH_D0
    opcode_map_primary(&HD1&) = OPH_D1
    opcode_map_primary(&HD2&) = OPH_D2
    opcode_map_primary(&HD3&) = OPH_D3
    opcode_map_primary(&HD4&) = OPH_D4
    opcode_map_primary(&HD5&) = OPH_D5
    opcode_map_primary(&HD6&) = OPH_D6_D7
    opcode_map_primary(&HD7&) = OPH_D6_D7
    opcode_map_primary(&HD8&) = OPH_FPU
    opcode_map_primary(&HD9&) = OPH_FPU
    opcode_map_primary(&HDA&) = OPH_FPU
    opcode_map_primary(&HDB&) = OPH_FPU
    opcode_map_primary(&HDC&) = OPH_FPU
    opcode_map_primary(&HDD&) = OPH_FPU
    opcode_map_primary(&HDE&) = OPH_FPU
    opcode_map_primary(&HDF&) = OPH_FPU
    opcode_map_primary(&HE0&) = OPH_E0
    opcode_map_primary(&HE1&) = OPH_E1
    opcode_map_primary(&HE2&) = OPH_E2
    opcode_map_primary(&HE3&) = OPH_E3
    opcode_map_primary(&HE4&) = OPH_E4
    opcode_map_primary(&HE5&) = OPH_E5
    opcode_map_primary(&HE6&) = OPH_E6
    opcode_map_primary(&HE7&) = OPH_E7
    opcode_map_primary(&HE8&) = OPH_E8
    opcode_map_primary(&HE9&) = OPH_E9
    opcode_map_primary(&HEA&) = OPH_EA
    opcode_map_primary(&HEB&) = OPH_EB
    opcode_map_primary(&HEC&) = OPH_EC
    opcode_map_primary(&HED&) = OPH_ED
    opcode_map_primary(&HEE&) = OPH_EE
    opcode_map_primary(&HEF&) = OPH_EF
    opcode_map_primary(&H68&) = OPH_68
    opcode_map_primary(&H69&) = OPH_69
    opcode_map_primary(&H6A&) = OPH_6A
    opcode_map_primary(&H6B&) = OPH_6B
    opcode_map_primary(&HC0&) = OPH_C0
    opcode_map_primary(&HC1&) = OPH_C1
    opcode_map_primary(&HC6&) = OPH_C6
    opcode_map_primary(&HC7&) = OPH_C7
    opcode_map_primary(&HC2&) = OPH_C2
    opcode_map_primary(&HC3&) = OPH_C3
    opcode_map_primary(&HC4&) = OPH_C4
    opcode_map_primary(&HC5&) = OPH_C5
    opcode_map_primary(&HA0&) = OPH_A0
    opcode_map_primary(&HA1&) = OPH_A1
    opcode_map_primary(&HA2&) = OPH_A2
    opcode_map_primary(&HA3&) = OPH_A3
    opcode_map_primary(&HA4&) = OPH_A4
    opcode_map_primary(&HA5&) = OPH_A5
    opcode_map_primary(&HA6&) = OPH_A6
    opcode_map_primary(&HA7&) = OPH_A7
    opcode_map_primary(&HA8&) = OPH_A8
    opcode_map_primary(&HA9&) = OPH_A9
    opcode_map_primary(&HB0&) = OPH_B0
    opcode_map_primary(&HB1&) = OPH_B1
    opcode_map_primary(&HB2&) = OPH_B2
    opcode_map_primary(&HB3&) = OPH_B3
    opcode_map_primary(&HB4&) = OPH_B4
    opcode_map_primary(&HB5&) = OPH_B5
    opcode_map_primary(&HB6&) = OPH_B6
    opcode_map_primary(&HB7&) = OPH_B7
    opcode_map_primary(&HB8&) = OPH_MOV_REGIMM
    opcode_map_primary(&HB9&) = OPH_MOV_REGIMM
    opcode_map_primary(&HBA&) = OPH_MOV_REGIMM
    opcode_map_primary(&HBB&) = OPH_MOV_REGIMM
    opcode_map_primary(&HBC&) = OPH_MOV_REGIMM
    opcode_map_primary(&HBD&) = OPH_MOV_REGIMM
    opcode_map_primary(&HBE&) = OPH_MOV_REGIMM
    opcode_map_primary(&HBF&) = OPH_MOV_REGIMM
    opcode_map_primary(&HAA&) = OPH_AA
    opcode_map_primary(&HAB&) = OPH_AB
    opcode_map_primary(&HAC&) = OPH_AC
    opcode_map_primary(&HAD&) = OPH_AD
    opcode_map_primary(&HAE&) = OPH_AE
    opcode_map_primary(&HAF&) = OPH_AF
    opcode_map_primary(&HC8&) = OPH_C8
    opcode_map_primary(&HC9&) = OPH_C9
    opcode_map_primary(&HCA&) = OPH_CA
    opcode_map_primary(&HCB&) = OPH_CB
    opcode_map_primary(&HCE&) = OPH_CE
    opcode_map_primary(&HCF&) = OPH_CF
    opcode_map_primary(&HF1&) = OPH_F1
    opcode_map_primary(&HF4&) = OPH_F4
    opcode_map_primary(&HF5&) = OPH_F5
    opcode_map_primary(&HF6&) = OPH_F6
    opcode_map_primary(&HF7&) = OPH_F7
    opcode_map_primary(&HF8&) = OPH_F8
    opcode_map_primary(&HF9&) = OPH_F9
    opcode_map_primary(&HFA&) = OPH_FA
    opcode_map_primary(&HFB&) = OPH_FB
    opcode_map_primary(&HFC&) = OPH_FC
    opcode_map_primary(&HFD&) = OPH_FD
    opcode_map_primary(&HFE&) = OPH_FE
    opcode_map_primary(&HFF&) = OPH_FF

    opcode_maps_ready = 1&
End Sub

Private Sub cpu_dispatchPrimaryOpcode(ByRef cpu As CPU_t, ByVal opcode As Long)
    opcode = (opcode And &HFF&)

    Select Case opcode_map_primary(opcode)
        Case OPH_EXT_0F
            cpu_extop cpu
        Case OPH_90
            op_90 cpu
        Case OPH_91
            op_91 cpu
        Case OPH_92
            op_92 cpu
        Case OPH_93
            op_93 cpu
        Case OPH_94
            op_94 cpu
        Case OPH_95
            op_95 cpu
        Case OPH_96
            op_96 cpu
        Case OPH_97
            op_97 cpu
        Case OPH_8C
            op_8C cpu
        Case OPH_8D
            op_8D cpu
        Case OPH_8E
            op_8E cpu
        Case OPH_8F
            op_8F cpu
        Case OPH_88
            op_88 cpu
        Case OPH_89
            op_89 cpu
        Case OPH_8A
            op_8A cpu
        Case OPH_8B
            op_8B cpu
        Case OPH_84
            op_84 cpu
        Case OPH_85
            op_85 cpu
        Case OPH_86
            op_86 cpu
        Case OPH_87
            op_87 cpu
        Case OPH_80
            op_80 cpu
        Case OPH_81
            op_81 cpu
        Case OPH_82
            op_82 cpu
        Case OPH_83
            op_83 cpu
        Case OPH_70
            op_70 cpu
        Case OPH_71
            op_71 cpu
        Case OPH_72
            op_72 cpu
        Case OPH_73
            op_73 cpu
        Case OPH_74
            op_74 cpu
        Case OPH_75
            op_75 cpu
        Case OPH_76
            op_76 cpu
        Case OPH_77
            op_77 cpu
        Case OPH_78
            op_78 cpu
        Case OPH_79
            op_79 cpu
        Case OPH_7A
            op_7A cpu
        Case OPH_7B
            op_7B cpu
        Case OPH_7C
            op_7C cpu
        Case OPH_7D
            op_7D cpu
        Case OPH_7E
            op_7E cpu
        Case OPH_7F
            op_7F cpu
        Case OPH_6C
            op_6C cpu
        Case OPH_6D
            op_6D cpu
        Case OPH_6E
            op_6E cpu
        Case OPH_6F
            op_6F cpu
        Case OPH_60
            op_60 cpu
        Case OPH_61
            op_61 cpu
        Case OPH_62
            op_62 cpu
        Case OPH_63
            op_63 cpu
        Case OPH_PUSH_REG
            op_pushreg cpu
        Case OPH_POP_REG
            op_popreg cpu
        Case OPH_INC_REG
            op_increg cpu
        Case OPH_DEC_REG
            op_decreg cpu
        Case OPH_30
            op_30 cpu
        Case OPH_31
            op_31 cpu
        Case OPH_32
            op_32 cpu
        Case OPH_33
            op_33 cpu
        Case OPH_34
            op_34 cpu
        Case OPH_35
            op_35 cpu
        Case OPH_37
            op_37 cpu
        Case OPH_38
            op_38 cpu
        Case OPH_39
            op_39 cpu
        Case OPH_3A
            op_3A cpu
        Case OPH_3B
            op_3B cpu
        Case OPH_3C
            op_3C cpu
        Case OPH_3D
            op_3D cpu
        Case OPH_3F
            op_3F cpu
        Case OPH_20
            op_20 cpu
        Case OPH_21
            op_21 cpu
        Case OPH_22
            op_22 cpu
        Case OPH_23
            op_23 cpu
        Case OPH_24
            op_24 cpu
        Case OPH_25
            op_25 cpu
        Case OPH_27
            op_27 cpu
        Case OPH_28
            op_28 cpu
        Case OPH_29
            op_29 cpu
        Case OPH_2A
            op_2A cpu
        Case OPH_2B
            op_2B cpu
        Case OPH_2C
            op_2C cpu
        Case OPH_2D
            op_2D cpu
        Case OPH_2F
            op_2F cpu
        Case OPH_10
            op_10 cpu
        Case OPH_11
            op_11 cpu
        Case OPH_12
            op_12 cpu
        Case OPH_13
            op_13 cpu
        Case OPH_14
            op_14 cpu
        Case OPH_15
            op_15 cpu
        Case OPH_16
            op_16 cpu
        Case OPH_17
            op_17 cpu
        Case OPH_18
            op_18 cpu
        Case OPH_19
            op_19 cpu
        Case OPH_1A
            op_1A cpu
        Case OPH_1B
            op_1B cpu
        Case OPH_1C
            op_1C cpu
        Case OPH_1D
            op_1D cpu
        Case OPH_1E
            op_1E cpu
        Case OPH_1F
            op_1F cpu
        Case OPH_00
            op_00 cpu
        Case OPH_01
            op_01 cpu
        Case OPH_02
            op_02 cpu
        Case OPH_03
            op_03 cpu
        Case OPH_04
            op_04 cpu
        Case OPH_05
            op_05 cpu
        Case OPH_06
            op_06 cpu
        Case OPH_07
            op_07 cpu
        Case OPH_08
            op_08 cpu
        Case OPH_09
            op_09 cpu
        Case OPH_0A
            op_0A cpu
        Case OPH_0B
            op_0B cpu
        Case OPH_0C
            op_0C cpu
        Case OPH_0D
            op_0D cpu
        Case OPH_0E
            op_0E cpu
        Case OPH_98
            op_98 cpu
        Case OPH_99
            op_99 cpu
        Case OPH_9A
            op_9A cpu
        Case OPH_9B
            op_9B cpu
        Case OPH_9C
            op_9C cpu
        Case OPH_9D
            op_9D cpu
        Case OPH_9E
            op_9E cpu
        Case OPH_9F
            op_9F cpu
        Case OPH_CC
            op_CC cpu
        Case OPH_CD
            op_CD cpu
        Case OPH_D0
            op_D0 cpu
        Case OPH_D1
            op_D1 cpu
        Case OPH_D2
            op_D2 cpu
        Case OPH_D3
            op_D3 cpu
        Case OPH_D4
            op_D4 cpu
        Case OPH_D5
            op_D5 cpu
        Case OPH_D6_D7
            op_D6_D7 cpu
        Case OPH_FPU
            op_fpu cpu
        Case OPH_E0
            op_E0 cpu
        Case OPH_E1
            op_E1 cpu
        Case OPH_E2
            op_E2 cpu
        Case OPH_E3
            op_E3 cpu
        Case OPH_E4
            op_E4 cpu
        Case OPH_E5
            op_E5 cpu
        Case OPH_E6
            op_E6 cpu
        Case OPH_E7
            op_E7 cpu
        Case OPH_E8
            op_E8 cpu
        Case OPH_E9
            op_E9 cpu
        Case OPH_EA
            op_EA cpu
        Case OPH_EB
            op_EB cpu
        Case OPH_EC
            op_EC cpu
        Case OPH_ED
            op_ED cpu
        Case OPH_EE
            op_EE cpu
        Case OPH_EF
            op_EF cpu
        Case OPH_68
            op_68 cpu
        Case OPH_69
            op_69 cpu
        Case OPH_6A
            op_6A cpu
        Case OPH_6B
            op_6B cpu
        Case OPH_C0
            op_C0 cpu
        Case OPH_C1
            op_C1 cpu
        Case OPH_C6
            op_C6 cpu
        Case OPH_C7
            op_C7 cpu
        Case OPH_C2
            op_C2 cpu
        Case OPH_C3
            op_C3 cpu
        Case OPH_C4
            op_C4 cpu
        Case OPH_C5
            op_C5 cpu
        Case OPH_A0
            op_A0 cpu
        Case OPH_A1
            op_A1 cpu
        Case OPH_A2
            op_A2 cpu
        Case OPH_A3
            op_A3 cpu
        Case OPH_A4
            op_A4 cpu
        Case OPH_A5
            op_A5 cpu
        Case OPH_A6
            op_A6 cpu
        Case OPH_A7
            op_A7 cpu
        Case OPH_A8
            op_A8 cpu
        Case OPH_A9
            op_A9 cpu
        Case OPH_AA
            op_AA cpu
        Case OPH_AB
            op_AB cpu
        Case OPH_AC
            op_AC cpu
        Case OPH_AD
            op_AD cpu
        Case OPH_AE
            op_AE cpu
        Case OPH_AF
            op_AF cpu
        Case OPH_B0
            op_B0 cpu
        Case OPH_B1
            op_B1 cpu
        Case OPH_B2
            op_B2 cpu
        Case OPH_B3
            op_B3 cpu
        Case OPH_B4
            op_B4 cpu
        Case OPH_B5
            op_B5 cpu
        Case OPH_B6
            op_B6 cpu
        Case OPH_B7
            op_B7 cpu
        Case OPH_MOV_REGIMM
            op_mov_regimm cpu
        Case OPH_C8
            op_C8 cpu
        Case OPH_C9
            op_C9 cpu
        Case OPH_CA
            op_CA cpu
        Case OPH_CB
            op_CB cpu
        Case OPH_CE
            op_CE cpu
        Case OPH_CF
            op_CF cpu
        Case OPH_F1
            op_F1 cpu
        Case OPH_F4
            op_F4 cpu
        Case OPH_F5
            op_F5 cpu
        Case OPH_F6
            op_F6 cpu
        Case OPH_F7
            op_F7 cpu
        Case OPH_F8
            op_F8 cpu
        Case OPH_F9
            op_F9 cpu
        Case OPH_FA
            op_FA cpu
        Case OPH_FB
            op_FB cpu
        Case OPH_FC
            op_FC cpu
        Case OPH_FD
            op_FD cpu
        Case OPH_FE
            op_FE cpu
        Case OPH_FF
            op_FF cpu
        Case Else
            op_illegal cpu
    End Select
End Sub

Private Sub cpu_dispatchExtOpcode(ByRef cpu As CPU_t, ByVal opcode As Long)
    opcode = (opcode And &HFF&)

    Select Case opcode_map_ext(opcode)
        Case OPHX_00
            op_ext_00 cpu
        Case OPHX_01
            op_ext_01 cpu
        Case OPHX_02
            op_ext_02 cpu
        Case OPHX_03
            op_ext_03 cpu
        Case OPHX_06
            op_ext_06 cpu
        Case OPHX_08_09
            op_ext_08_09 cpu
        Case OPHX_20
            op_ext_20 cpu
        Case OPHX_21
            op_ext_21 cpu
        Case OPHX_22
            op_ext_22 cpu
        Case OPHX_23
            op_ext_23 cpu
        Case OPHX_24_26
            op_ext_24_26 cpu
        Case OPHX_30
            op_ext_30 cpu
        Case OPHX_31
            op_ext_31 cpu
        Case OPHX_40
            op_ext_40 cpu
        Case OPHX_41
            op_ext_41 cpu
        Case OPHX_42
            op_ext_42 cpu
        Case OPHX_43
            op_ext_43 cpu
        Case OPHX_44
            op_ext_44 cpu
        Case OPHX_45
            op_ext_45 cpu
        Case OPHX_46
            op_ext_46 cpu
        Case OPHX_47
            op_ext_47 cpu
        Case OPHX_48
            op_ext_48 cpu
        Case OPHX_49
            op_ext_49 cpu
        Case OPHX_4A
            op_ext_4A cpu
        Case OPHX_4B
            op_ext_4B cpu
        Case OPHX_4C
            op_ext_4C cpu
        Case OPHX_4D
            op_ext_4D cpu
        Case OPHX_4E
            op_ext_4E cpu
        Case OPHX_4F
            op_ext_4F cpu
        Case OPHX_80
            op_ext_80 cpu
        Case OPHX_81
            op_ext_81 cpu
        Case OPHX_82
            op_ext_82 cpu
        Case OPHX_83
            op_ext_83 cpu
        Case OPHX_84
            op_ext_84 cpu
        Case OPHX_85
            op_ext_85 cpu
        Case OPHX_86
            op_ext_86 cpu
        Case OPHX_87
            op_ext_87 cpu
        Case OPHX_88
            op_ext_88 cpu
        Case OPHX_89
            op_ext_89 cpu
        Case OPHX_8A
            op_ext_8A cpu
        Case OPHX_8B
            op_ext_8B cpu
        Case OPHX_8C
            op_ext_8C cpu
        Case OPHX_8D
            op_ext_8D cpu
        Case OPHX_8E
            op_ext_8E cpu
        Case OPHX_8F
            op_ext_8F cpu
        Case OPHX_90
            op_ext_90 cpu
        Case OPHX_91
            op_ext_91 cpu
        Case OPHX_92
            op_ext_92 cpu
        Case OPHX_93
            op_ext_93 cpu
        Case OPHX_94
            op_ext_94 cpu
        Case OPHX_95
            op_ext_95 cpu
        Case OPHX_96
            op_ext_96 cpu
        Case OPHX_97
            op_ext_97 cpu
        Case OPHX_98
            op_ext_98 cpu
        Case OPHX_99
            op_ext_99 cpu
        Case OPHX_9A
            op_ext_9A cpu
        Case OPHX_9B
            op_ext_9B cpu
        Case OPHX_9C
            op_ext_9C cpu
        Case OPHX_9D
            op_ext_9D cpu
        Case OPHX_9E
            op_ext_9E cpu
        Case OPHX_9F
            op_ext_9F cpu
        Case OPHX_A0
            op_ext_A0 cpu
        Case OPHX_A1
            op_ext_A1 cpu
        Case OPHX_A2
            op_ext_A2 cpu
        Case OPHX_A3
            op_ext_A3 cpu
        Case OPHX_A4_A5
            op_ext_A4_A5 cpu
        Case OPHX_A6_B0
            op_ext_A6_B0 cpu
        Case OPHX_A8
            op_ext_A8 cpu
        Case OPHX_A9
            op_ext_A9 cpu
        Case OPHX_AA
            op_ext_AA cpu
        Case OPHX_AB
            op_ext_AB cpu
        Case OPHX_AC_AD
            op_ext_AC_AD cpu
        Case OPHX_AF
            op_ext_AF cpu
        Case OPHX_B1
            op_ext_B1 cpu
        Case OPHX_B2_B4_B5
            op_ext_B2_B4_B5 cpu
        Case OPHX_B3
            op_ext_B3 cpu
        Case OPHX_B6
            op_ext_B6 cpu
        Case OPHX_B7
            op_ext_B7 cpu
        Case OPHX_BA
            op_ext_BA cpu
        Case OPHX_BB
            op_ext_BB cpu
        Case OPHX_BC
            op_ext_BC cpu
        Case OPHX_BD
            op_ext_BD cpu
        Case OPHX_BE
            op_ext_BE cpu
        Case OPHX_BF
            op_ext_BF cpu
        Case OPHX_C0
            op_ext_C0 cpu
        Case OPHX_C1
            op_ext_C1 cpu
        Case OPHX_C8_CF
            op_ext_C8_C9_CA_CB_CC_CD_CE_CF cpu
        Case Else
            op_ext_illegal cpu
    End Select
End Sub

Public Sub cpu_extop(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim extOpcode As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    extOpcode = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    cpu.opcode = CByte(extOpcode)
    cpu_dispatchExtOpcode cpu, extOpcode
End Sub

Public Sub op_90(ByRef cpu As CPU_t)
    ' NOP
End Sub

Public Sub op_xchg(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, (cpu.opcode And &H7&))
        putreg32 cpu, (cpu.opcode And &H7&), cpu.regs_long(CPU_REG_EAX)
        cpu.regs_long(CPU_REG_EAX) = cpu.oper1_32
    Else
        cpu.oper1 = getreg16(cpu, (cpu.opcode And &H7&))
        putreg16 cpu, (cpu.opcode And &H7&), getreg16(cpu, CPU_REG_EAX)
        putreg16 cpu, CPU_REG_EAX, cpu.oper1
    End If
End Sub

Public Sub op_91(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_92(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_93(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_94(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_95(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_96(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_97(ByRef cpu As CPU_t)
    op_xchg cpu
End Sub

Public Sub op_6C(ByRef cpu As CPU_t)
    Dim addr As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        If (cpu.reptype <> 0&) And (getreg16(cpu, CPU_REG_ECX) = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    cpu_write cpu, addr, (port_read(cpu, getreg16(cpu, CPU_REG_EDX)) And &HFF&)

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_6D(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim stepVal As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            If getreg16(cpu, CPU_REG_ECX) = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        addr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        addr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    If cpu.isoper32 <> 0& Then
        cpu_writel cpu, addr, port_readl(cpu, getreg16(cpu, CPU_REG_EDX))
        stepVal = 4&
    Else
        cpu_writew cpu, addr, port_readw(cpu, getreg16(cpu, CPU_REG_EDX))
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_6E(ByRef cpu As CPU_t)
    Dim addr As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            If getreg16(cpu, CPU_REG_ECX) = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        addr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
    Else
        addr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
    End If

    port_write cpu, getreg16(cpu, CPU_REG_EDX), (cpu_read(cpu, addr) And &HFF&)

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_6F(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim stepVal As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            If getreg16(cpu, CPU_REG_ECX) = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        addr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
    Else
        addr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
    End If

    If cpu.isoper32 <> 0& Then
        port_writel cpu, getreg16(cpu, CPU_REG_EDX), cpu_readl(cpu, addr)
        stepVal = 4&
    Else
        port_writew cpu, getreg16(cpu, CPU_REG_EDX), (cpu_readw(cpu, addr) And &HFFFF&)
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_70(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.ofl <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_71(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.ofl = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_72(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.cf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_73(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.cf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_74(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.zf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_75(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.zf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_76(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If (cpu.cf <> 0&) Or (cpu.zf <> 0&) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_77(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If (cpu.cf = 0&) And (cpu.zf = 0&) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_78(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.sf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_79(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.sf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_7A(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.pf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_7B(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.pf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_7C(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.sf <> cpu.ofl Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_7D(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.sf = cpu.ofl Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_7E(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If (cpu.sf <> cpu.ofl) Or (cpu.zf <> 0&) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_7F(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If (cpu.zf = 0&) And (cpu.sf = cpu.ofl) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_80_82(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    Select Case (cpu.reg And &H7&)
        Case 0&
            op_add8 cpu
        Case 1&
            op_or8 cpu
        Case 2&
            op_adc8 cpu
        Case 3&
            op_sbb8 cpu
        Case 4&
            op_and8 cpu
        Case 5&
            op_sub8 cpu
        Case 6&
            op_xor8 cpu
        Case 7&
            flag_sub8 cpu, cpu.oper1b, cpu.oper2b
    End Select

    If (cpu.reg And &H7&) < 7& Then
        writerm8 cpu, cpu.rm, cpu.res8
    End If
End Sub

Public Sub op_81_83(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    modregrm cpu

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        If (cpu.opcode And &HFF&) = &H81& Then
            cpu.oper2_32 = cpu_readl(cpu, codeAddr)
            cpu_stepIP cpu, 4&
        Else
            cpu.oper2_32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
            cpu_stepIP cpu, 1&
        End If
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        If (cpu.opcode And &HFF&) = &H81& Then
            cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
            cpu_stepIP cpu, 2&
        Else
            cpu.oper2 = cpu_signext8to16(cpu_read(cpu, codeAddr))
            cpu_stepIP cpu, 1&
        End If
    End If

    Select Case (cpu.reg And &H7&)
        Case 0&
            If cpu.isoper32 <> 0& Then
                op_add32 cpu
            Else
                op_add16 cpu
            End If
        Case 1&
            If cpu.isoper32 <> 0& Then
                op_or32 cpu
            Else
                op_or16 cpu
            End If
        Case 2&
            If cpu.isoper32 <> 0& Then
                op_adc32 cpu
            Else
                op_adc16 cpu
            End If
        Case 3&
            If cpu.isoper32 <> 0& Then
                op_sbb32 cpu
            Else
                op_sbb16 cpu
            End If
        Case 4&
            If cpu.isoper32 <> 0& Then
                op_and32 cpu
            Else
                op_and16 cpu
            End If
        Case 5&
            If cpu.isoper32 <> 0& Then
                op_sub32 cpu
            Else
                op_sub16 cpu
            End If
        Case 6&
            If cpu.isoper32 <> 0& Then
                op_xor32 cpu
            Else
                op_xor16 cpu
            End If
        Case 7&
            If cpu.isoper32 <> 0& Then
                flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
            Else
                flag_sub16 cpu, cpu.oper1, cpu.oper2
            End If
    End Select

    If (cpu.reg And &H7&) < 7& Then
        If cpu.isoper32 <> 0& Then
            writerm32 cpu, cpu.rm, cpu.res32
        Else
            writerm16 cpu, cpu.rm, cpu.res16
        End If
    End If
End Sub

Public Sub op_80(ByRef cpu As CPU_t)
    op_80_82 cpu
End Sub

Public Sub op_81(ByRef cpu As CPU_t)
    op_81_83 cpu
End Sub

Public Sub op_82(ByRef cpu As CPU_t)
    op_80_82 cpu
End Sub

Public Sub op_83(ByRef cpu As CPU_t)
    op_81_83 cpu
End Sub

Public Sub op_84(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    flag_log8 cpu, ((cpu.oper1b And &HFF&) And (cpu.oper2b And &HFF&))
End Sub

Public Sub op_85(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        flag_log32 cpu, (cpu.oper1_32 And cpu.oper2_32)
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        flag_log16 cpu, ((cpu.oper1 And &HFFFF&) And (cpu.oper2 And &HFFFF&))
    End If
End Sub

Public Sub op_86(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    putreg8 cpu, cpu.reg, readrm8(cpu, cpu.rm)
    writerm8 cpu, cpu.rm, cpu.oper1b
End Sub

Public Sub op_87(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        writerm32 cpu, cpu.rm, cpu.oper1_32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        writerm16 cpu, cpu.rm, cpu.oper1
    End If
End Sub

Public Sub op_88(ByRef cpu As CPU_t)
    modregrm cpu
    writerm8 cpu, cpu.rm, getreg8(cpu, cpu.reg)
End Sub

Public Sub op_89(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        writerm32 cpu, cpu.rm, getreg32(cpu, cpu.reg)
    Else
        writerm16 cpu, cpu.rm, getreg16(cpu, cpu.reg)
    End If
End Sub

Public Sub op_8A(ByRef cpu As CPU_t)
    modregrm cpu
    putreg8 cpu, cpu.reg, readrm8(cpu, cpu.rm)
End Sub

Public Sub op_8B(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
    Else
        putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
    End If
End Sub

Public Sub op_8C(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.reg > CPU_REG_GS Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    writerm16 cpu, cpu.rm, getsegreg(cpu, cpu.reg)
End Sub

Public Sub op_8D(ByRef cpu As CPU_t)
    Dim offset As Long

    modregrm cpu

    If cpu.mode = 3& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    offset = cpu_effective_offset(cpu, cpu.rm)

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, offset
    Else
        putreg16 cpu, cpu.reg, (offset And &HFFFF&)
    End If
End Sub

Public Sub op_8E(ByRef cpu As CPU_t)
    Dim value As Long

    modregrm cpu

    If (cpu.reg = CPU_REG_CS) Or (cpu.reg > CPU_REG_GS) Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    value = (readrm16(cpu, cpu.rm) And &HFFFF&)
    If cpu.doexception <> 0& Then Exit Sub

    putsegreg cpu, cpu.reg, value
    If cpu.reg = CPU_REG_SS Then
        cpu_begin_interrupt_shadow cpu
    End If
End Sub

Public Sub op_8F(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim modrm As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    modrm = (cpu_read(cpu, codeAddr) And &HFF&)
    If (((modrm \ 8&) And &H7&) <> 0&) Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    If cpu.isaddr32 <> 0& Then
        If cpu.isoper32 <> 0& Then
            cpu.shadow_esp = U32Add(cpu.shadow_esp, 4&)
        Else
            cpu.shadow_esp = U32Add(cpu.shadow_esp, 2&)
        End If
    End If

    If cpu.isoper32 <> 0& Then
        modregrm cpu
        writerm32 cpu, cpu.rm, pop(cpu)
    Else
        modregrm cpu
        writerm16 cpu, cpu.rm, pop(cpu)
    End If
End Sub

Public Sub op_98(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        If (getreg16(cpu, CPU_REG_EAX) And &H8000&) <> 0& Then
            cpu.regs_long(CPU_REG_EAX) = (cpu.regs_long(CPU_REG_EAX) Or &HFFFF0000)
        Else
            cpu.regs_long(CPU_REG_EAX) = (cpu.regs_long(CPU_REG_EAX) And &HFFFF&)
        End If
    Else
        If (cpu_getReg8Low(cpu, CPU_REG_EAX) And &H80&) <> 0& Then
            cpu_setReg8High cpu, CPU_REG_EAX, &HFF&
        Else
            cpu_setReg8High cpu, CPU_REG_EAX, 0&
        End If
    End If
End Sub

Public Sub op_99(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        If (U32Shr(cpu.regs_long(CPU_REG_EAX), 31&) And &H1&) <> 0& Then
            cpu.regs_long(CPU_REG_EDX) = -1&
        Else
            cpu.regs_long(CPU_REG_EDX) = 0&
        End If
    Else
        If (cpu_getReg8High(cpu, CPU_REG_EAX) And &H80&) <> 0& Then
            putreg16 cpu, CPU_REG_EDX, &HFFFF&
        Else
            putreg16 cpu, CPU_REG_EDX, 0&
        End If
    End If
End Sub

Public Sub op_9A(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    Else
        cpu.oper1_32 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If

    cpu_callf cpu, cpu.oper2, cpu.oper1_32
End Sub

Private Function cpu_fpu_check_nm(ByRef cpu As CPU_t) As Long
    If (cpu.CR(0&) And (CR0_EM Or CR0_TS)) <> 0& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 7&, 0&
        cpu_fpu_check_nm = 1&
    Else
        cpu_fpu_check_nm = 0&
    End If
End Function

Public Sub op_9B(ByRef cpu As CPU_t)
    If cpu_fpu_check_nm(cpu) <> 0& Then
        Exit Sub
    End If
End Sub

Public Sub op_9C(ByRef cpu As CPU_t)
    Dim flags As Long

    If (cpu.v86f <> 0&) And (cpu.iopl < 3&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    flags = (cpu_makeflagsword(cpu) And Not (EFLAGS_VM Or EFLAGS_RF))
    push cpu, flags
End Sub

Public Sub op_9D(ByRef cpu As CPU_t)
    Dim old_flags As Long
    Dim raw_flags As Long
    Dim new_flags As Long
    Dim modifiable_mask As Long

    If (cpu.v86f <> 0&) And (cpu.iopl < 3&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    old_flags = cpu_makeflagsword(cpu)
    If cpu.isoper32 <> 0& Then
        raw_flags = pop(cpu)
    Else
        raw_flags = ((old_flags And &HFFFF0000) Or (pop(cpu) And &HFFFF&))
    End If

    modifiable_mask = (EFLAGS_CF Or EFLAGS_PF Or EFLAGS_AF Or EFLAGS_ZF Or EFLAGS_SF Or EFLAGS_TF Or EFLAGS_DF Or EFLAGS_OF Or EFLAGS_NT)
    If cpu.isoper32 <> 0& Then
        modifiable_mask = (modifiable_mask Or EFLAGS_AC Or EFLAGS_ID)
    End If

    new_flags = ((old_flags And Not modifiable_mask) Or (raw_flags And modifiable_mask))
    If (cpu.protected_mode = 0&) Or (cpu.cpl <= cpu.iopl) Then
        new_flags = ((new_flags And Not EFLAGS_IF) Or (raw_flags And EFLAGS_IF))
    Else
        new_flags = ((new_flags And Not EFLAGS_IF) Or (old_flags And EFLAGS_IF))
    End If

    If (cpu.protected_mode = 0&) Or (cpu.cpl = 0&) Then
        new_flags = ((new_flags And Not EFLAGS_IOPL) Or (raw_flags And EFLAGS_IOPL))
    Else
        new_flags = ((new_flags And Not EFLAGS_IOPL) Or (old_flags And EFLAGS_IOPL))
    End If

    If (new_flags And EFLAGS_CF) <> 0& Then cpu.cf = 1& Else cpu.cf = 0&
    If (new_flags And EFLAGS_PF) <> 0& Then cpu.pf = 1& Else cpu.pf = 0&
    If (new_flags And EFLAGS_AF) <> 0& Then cpu.af = 1& Else cpu.af = 0&
    If (new_flags And EFLAGS_ZF) <> 0& Then cpu.zf = 1& Else cpu.zf = 0&
    If (new_flags And EFLAGS_SF) <> 0& Then cpu.sf = 1& Else cpu.sf = 0&
    If (new_flags And EFLAGS_TF) <> 0& Then cpu.tf = 1& Else cpu.tf = 0&
    If (new_flags And EFLAGS_IF) <> 0& Then cpu.ifl = 1& Else cpu.ifl = 0&
    If (new_flags And EFLAGS_DF) <> 0& Then cpu.df = 1& Else cpu.df = 0&
    If (new_flags And EFLAGS_OF) <> 0& Then cpu.ofl = 1& Else cpu.ofl = 0&
    cpu.iopl = CByte((U32Shr(new_flags, 12&) And &H3&))
    If (new_flags And EFLAGS_NT) <> 0& Then cpu.nt = 1& Else cpu.nt = 0&
    If (old_flags And EFLAGS_RF) <> 0& Then cpu.rf = 1& Else cpu.rf = 0&
    If (old_flags And EFLAGS_VM) <> 0& Then cpu.v86f = 1& Else cpu.v86f = 0&
    If (new_flags And EFLAGS_AC) <> 0& Then cpu.acf = 1& Else cpu.acf = 0&
    If (new_flags And EFLAGS_ID) <> 0& Then cpu.idf = 1& Else cpu.idf = 0&
End Sub

Public Sub op_9E(ByRef cpu As CPU_t)
    decodeflagsword cpu, ((cpu_makeflagsword(cpu) And &HFFFFFF00) Or (cpu_getReg8High(cpu, CPU_REG_EAX) And &HFF&))
End Sub

Public Sub op_9F(ByRef cpu As CPU_t)
    cpu_setReg8High cpu, CPU_REG_EAX, (cpu_makeflagsword(cpu) And &HFF&)
End Sub

Public Sub op_CC(ByRef cpu As CPU_t)
    cpu_intcall cpu, 3&, INT_SOURCE_INT3, 0&
End Sub

Public Sub op_CD(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim intnum As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    intnum = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    cpu_intcall cpu, intnum, INT_SOURCE_SOFTWARE, 0&
End Sub

Public Sub op_D0(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    writerm8 cpu, cpu.rm, op_grp2_8(cpu, 1&)
End Sub

Public Sub op_D1(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        writerm32 cpu, cpu.rm, op_grp2_32(cpu, 1&)
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        writerm16 cpu, cpu.rm, op_grp2_16(cpu, 1&)
    End If
End Sub

Public Sub op_D2(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    writerm8 cpu, cpu.rm, op_grp2_8(cpu, cpu_getReg8Low(cpu, CPU_REG_ECX))
End Sub

Public Sub op_D3(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        writerm32 cpu, cpu.rm, op_grp2_32(cpu, cpu_getReg8Low(cpu, CPU_REG_ECX))
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        writerm16 cpu, cpu.rm, op_grp2_16(cpu, cpu_getReg8Low(cpu, CPU_REG_ECX))
    End If
End Sub

Public Sub op_D4(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim oldAl As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1 = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    If cpu.oper1 = 0& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    oldAl = (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
    cpu_setReg8High cpu, CPU_REG_EAX, ((oldAl \ cpu.oper1) And &HFF&)
    cpu_setReg8Low cpu, CPU_REG_EAX, ((oldAl Mod cpu.oper1) And &HFF&)
    flag_szp8 cpu, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
End Sub

Public Sub op_D5(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim oldAl As Long
    Dim oldAh As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1 = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    oldAl = (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
    oldAh = (cpu_getReg8High(cpu, CPU_REG_EAX) And &HFF&)
    cpu_setReg8Low cpu, CPU_REG_EAX, (((oldAh * cpu.oper1) + oldAl) And &HFF&)
    cpu_setReg8High cpu, CPU_REG_EAX, 0&
    flag_szp8 cpu, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
End Sub

Public Sub op_D6_D7(ByRef cpu As CPU_t)
    Dim addr As Long

    If cpu.isaddr32 <> 0& Then
        addr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_EBX))
    Else
        addr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_EBX))
    End If

    addr = U32Add(addr, cpu_getReg8Low(cpu, CPU_REG_EAX))
    cpu_setReg8Low cpu, CPU_REG_EAX, (cpu_read(cpu, addr) And &HFF&)
End Sub

Public Sub op_fpu(ByRef cpu As CPU_t)
    Dim op As Long

    If cpu_fpu_check_nm(cpu) <> 0& Then
        Exit Sub
    End If

    If cpu.have387 = 0& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    modregrm cpu
    If cpu.doexception <> 0& Then
        Exit Sub
    End If

    op = (cpu.opcode And &H7&)
    If cpu.mode = 3& Then
        Call fpu_exec1(cpu, op, cpu.reg, cpu.rm)
    Else
        getea cpu, cpu.rm
        If cpu.doexception <> 0& Then
            Exit Sub
        End If
        Call fpu_exec2(cpu, (cpu.isoper32 Xor 1&), op, cpu.reg, cpu.currentseg, cpu.ea)
    End If
End Sub
Public Sub op_E0(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim cxVal As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.isaddr32 <> 0& Then
        cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        If (cpu.regs_long(CPU_REG_ECX) <> 0&) And (cpu.zf = 0&) Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    Else
        cxVal = (getreg16(cpu, CPU_REG_ECX) - 1&) And &HFFFF&
        putreg16 cpu, CPU_REG_ECX, cxVal
        If (cxVal <> 0&) And (cpu.zf = 0&) Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    End If
End Sub

Public Sub op_E1(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim cxVal As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.isaddr32 <> 0& Then
        cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        If (cpu.regs_long(CPU_REG_ECX) <> 0&) And (cpu.zf <> 0&) Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    Else
        cxVal = (getreg16(cpu, CPU_REG_ECX) - 1&) And &HFFFF&
        putreg16 cpu, CPU_REG_ECX, cxVal
        If (cxVal <> 0&) And (cpu.zf <> 0&) Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    End If
End Sub

Public Sub op_E2(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim cxVal As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.isaddr32 <> 0& Then
        cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        If cpu.regs_long(CPU_REG_ECX) <> 0& Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    Else
        cxVal = (getreg16(cpu, CPU_REG_ECX) - 1&) And &HFFFF&
        putreg16 cpu, CPU_REG_ECX, cxVal
        If cxVal <> 0& Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    End If
End Sub

Public Sub op_E3(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.temp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&

    If cpu.isaddr32 <> 0& Then
        If cpu.regs_long(CPU_REG_ECX) = 0& Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    Else
        If getreg16(cpu, CPU_REG_ECX) = 0& Then
            cpu_apply_relative_branch cpu, cpu.temp32
        End If
    End If
End Sub

Public Sub op_E4(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    cpu_setReg8Low cpu, CPU_REG_EAX, (port_read(cpu, cpu.oper1b) And &HFF&)
End Sub

Public Sub op_E5(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    If cpu.isoper32 <> 0& Then
        cpu.regs_long(CPU_REG_EAX) = port_readl(cpu, cpu.oper1b)
    Else
        putreg16 cpu, CPU_REG_EAX, port_readw(cpu, cpu.oper1b)
    End If
End Sub

Public Sub op_E6(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    port_write cpu, cpu.oper1b, CByte(cpu_getReg8Low(cpu, CPU_REG_EAX))
End Sub

Public Sub op_E7(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    If cpu.isoper32 <> 0& Then
        port_writel cpu, cpu.oper1b, cpu.regs_long(CPU_REG_EAX)
    Else
        port_writew cpu, cpu.oper1b, getreg16(cpu, CPU_REG_EAX)
    End If
End Sub

Public Sub op_E8(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        push cpu, cpu.ip
    Else
        cpu.oper1_32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
        push cpu, cpu.ip
    End If

    cpu.ip = U32Add(cpu.ip, cpu.oper1_32)
    If cpu.isoper32 = 0& Then
        cpu.ip = (cpu.ip And &HFFFF&)
    End If
End Sub

Public Sub op_E9(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.oper1_32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    cpu.ip = U32Add(cpu.ip, cpu.oper1_32)
    If cpu.isoper32 = 0& Then
        cpu.ip = (cpu.ip And &HFFFF&)
    End If
End Sub

Public Sub op_EA(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        cpu_jmpf cpu, cpu.oper2, cpu.oper1_32
    Else
        cpu.oper1 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        cpu_jmpf cpu, cpu.oper2, cpu.oper1
    End If
End Sub

Public Sub op_EB(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1_32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&
    cpu.ip = U32Add(cpu.ip, cpu.oper1_32)
    If cpu.isoper32 = 0& Then
        cpu.ip = (cpu.ip And &HFFFF&)
    End If
End Sub

Public Sub op_EC(ByRef cpu As CPU_t)
    cpu.oper1 = getreg16(cpu, CPU_REG_EDX)
    cpu_setReg8Low cpu, CPU_REG_EAX, (port_read(cpu, cpu.oper1) And &HFF&)
End Sub

Public Sub op_ED(ByRef cpu As CPU_t)
    cpu.oper1 = getreg16(cpu, CPU_REG_EDX)

    If cpu.isoper32 <> 0& Then
        cpu.regs_long(CPU_REG_EAX) = port_readl(cpu, cpu.oper1)
    Else
        putreg16 cpu, CPU_REG_EAX, port_readw(cpu, cpu.oper1)
    End If
End Sub

Public Sub op_EE(ByRef cpu As CPU_t)
    cpu.oper1 = getreg16(cpu, CPU_REG_EDX)
    port_write cpu, cpu.oper1, CByte(cpu_getReg8Low(cpu, CPU_REG_EAX))
End Sub

Public Sub op_EF(ByRef cpu As CPU_t)
    cpu.oper1 = getreg16(cpu, CPU_REG_EDX)

    If cpu.isoper32 <> 0& Then
        port_writel cpu, cpu.oper1, cpu.regs_long(CPU_REG_EAX)
    Else
        port_writew cpu, cpu.oper1, getreg16(cpu, CPU_REG_EAX)
    End If
End Sub
Private Function cpu_signedMul32ToU64(ByVal lhs As Long, ByVal rhs As Long) As U64_t
    Dim a As Long
    Dim b As Long
    Dim signFlag As Long
    Dim prod As U64_t

    signFlag = 0&
    If ((lhs < 0&) Xor (rhs < 0&)) Then signFlag = 1&

    a = lhs
    b = rhs
    If a < 0& Then a = U32Add(Not a, 1&)
    If b < 0& Then b = U32Add(Not b, 1&)

    prod = cpu_u32MulToU64(a, b)
    If signFlag <> 0& Then
        prod = cpu_u64Neg(prod)
    End If

    cpu_signedMul32ToU64 = prod
End Function

Private Function cpu_u64SignExtMatches32(ByRef value64 As U64_t) As Long
    Dim expectedHi As Long

    If (U32Shr(value64.Lo, 31&) And &H1&) <> 0& Then
        expectedHi = -1&
    Else
        expectedHi = 0&
    End If

    If value64.Hi = expectedHi Then
        cpu_u64SignExtMatches32 = 1&
    Else
        cpu_u64SignExtMatches32 = 0&
    End If
End Function

Public Sub op_00(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_add8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_01(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_add32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_add16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_02(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_add8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_03(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_add32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_add16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_04(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_add8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_05(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_add32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_add16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_06(ByRef cpu As CPU_t)
    push cpu, cpu.segregs(CPU_REG_ES)
End Sub

Public Sub op_07(ByRef cpu As CPU_t)
    putsegreg cpu, CPU_REG_ES, pop(cpu)
End Sub

Public Sub op_08(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_or8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_09(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_or32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_or16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_0A(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_or8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_0B(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_or32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_or16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_0C(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_or8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_0D(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_or32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_or16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_0E(ByRef cpu As CPU_t)
    push cpu, cpu.segregs(CPU_REG_CS)
End Sub

Public Sub op_10(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_adc8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_11(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_adc32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_adc16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_12(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_adc8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_13(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_adc32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_adc16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_14(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_adc8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_15(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_adc32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_adc16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_16(ByRef cpu As CPU_t)
    push cpu, cpu.segregs(CPU_REG_SS)
End Sub

Public Sub op_17(ByRef cpu As CPU_t)
    putsegreg cpu, CPU_REG_SS, pop(cpu)
    cpu_begin_interrupt_shadow cpu
End Sub

Public Sub op_18(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_sbb8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_19(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_sbb32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_sbb16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_1A(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_sbb8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_1B(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_sbb32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_sbb16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_1C(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_sbb8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_1D(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_sbb32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_sbb16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_1E(ByRef cpu As CPU_t)
    push cpu, cpu.segregs(CPU_REG_DS)
End Sub

Public Sub op_1F(ByRef cpu As CPU_t)
    putsegreg cpu, CPU_REG_DS, pop(cpu)
End Sub

Public Sub op_20(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_and8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_21(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_and32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_and16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_22(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_and8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_23(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_and32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_and16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_24(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_and8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_25(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_and32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_and16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_27(ByRef cpu As CPU_t)
    Dim oldAl As Long
    Dim oldcf As Long
    Dim alVal As Long
    Dim tmp As Long

    alVal = (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
    oldAl = alVal
    oldcf = (cpu.cf And &H1&)
    cpu.cf = 0&

    If (((alVal And &HF&) > 9&) Or (cpu.af <> 0&)) Then
        tmp = (alVal + &H6&)
        cpu_setReg8Low cpu, CPU_REG_EAX, (tmp And &HFF&)
        If (oldcf <> 0&) Or ((tmp And &HFF00&) <> 0&) Then
            cpu.cf = 1&
        Else
            cpu.cf = 0&
        End If
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If

    If (oldAl > &H99&) Or (oldcf <> 0&) Then
        alVal = (((cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&) + &H60&) And &HFF&)
        cpu_setReg8Low cpu, CPU_REG_EAX, alVal
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    flag_szp8 cpu, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
End Sub

Public Sub op_28(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_sub8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_29(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_sub32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_sub16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_2A(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_sub8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_2B(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_sub32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_sub16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_2C(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_sub8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_2D(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_sub32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_sub16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_2F(ByRef cpu As CPU_t)
    Dim oldAl As Long
    Dim oldcf As Long
    Dim alVal As Long
    Dim tmp As Long

    alVal = (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
    oldAl = alVal
    oldcf = (cpu.cf And &H1&)
    cpu.cf = 0&

    If (((alVal And &HF&) > 9&) Or (cpu.af <> 0&)) Then
        tmp = (alVal - &H6&)
        cpu_setReg8Low cpu, CPU_REG_EAX, (tmp And &HFF&)
        If (oldcf <> 0&) Or ((tmp And &HFF00&) <> 0&) Then
            cpu.cf = 1&
        Else
            cpu.cf = 0&
        End If
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If

    If (oldAl > &H99&) Or (oldcf <> 0&) Then
        alVal = (((cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&) - &H60&) And &HFF&)
        cpu_setReg8Low cpu, CPU_REG_EAX, alVal
        cpu.cf = 1&
    End If

    flag_szp8 cpu, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)
End Sub

Public Sub op_30(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_xor8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
End Sub

Public Sub op_31(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_xor32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_xor16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
    End If
End Sub

Public Sub op_32(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_xor8 cpu
    putreg8 cpu, cpu.reg, cpu.res8
End Sub

Public Sub op_33(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_xor32 cpu
        putreg32 cpu, cpu.reg, cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_xor16 cpu
        putreg16 cpu, cpu.reg, cpu.res16
    End If
End Sub

Public Sub op_34(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    op_xor8 cpu
    cpu_setReg8Low cpu, CPU_REG_EAX, cpu.res8
End Sub

Public Sub op_35(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        op_xor32 cpu
        cpu.regs_long(CPU_REG_EAX) = cpu.res32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        op_xor16 cpu
        putreg16 cpu, CPU_REG_EAX, cpu.res16
    End If
End Sub

Public Sub op_37(ByRef cpu As CPU_t)
    Dim alVal As Long
    Dim ahVal As Long

    If ((cpu_getReg8Low(cpu, CPU_REG_EAX) And &HF&) > 9&) Or (cpu.af <> 0&) Then
        alVal = (((cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&) + &H6&) And &HF&)
        ahVal = (((cpu_getReg8High(cpu, CPU_REG_EAX) And &HFF&) + 1&) And &HFF&)
        cpu_setReg8Low cpu, CPU_REG_EAX, alVal
        cpu_setReg8High cpu, CPU_REG_EAX, ahVal
        cpu.af = 1&
        cpu.cf = 1&
    Else
        cpu.af = 0&
        cpu.cf = 0&
        cpu_setReg8Low cpu, CPU_REG_EAX, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HF&)
    End If
End Sub

Public Sub op_38(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    flag_sub8 cpu, cpu.oper1b, cpu.oper2b
End Sub

Public Sub op_39(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        flag_sub16 cpu, cpu.oper1, cpu.oper2
    End If
End Sub

Public Sub op_3A(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = getreg8(cpu, cpu.reg)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    flag_sub8 cpu, cpu.oper1b, cpu.oper2b
End Sub

Public Sub op_3B(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = getreg32(cpu, cpu.reg)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
    Else
        cpu.oper1 = getreg16(cpu, cpu.reg)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        flag_sub16 cpu, cpu.oper1, cpu.oper2
    End If
End Sub

Public Sub op_3C(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    flag_sub8 cpu, cpu.oper1b, cpu.oper2b
End Sub

Public Sub op_3D(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        flag_sub16 cpu, cpu.oper1, cpu.oper2
    End If
End Sub

Public Sub op_3F(ByRef cpu As CPU_t)
    Dim alVal As Long
    Dim ahVal As Long

    If ((cpu_getReg8Low(cpu, CPU_REG_EAX) And &HF&) > 9&) Or (cpu.af <> 0&) Then
        alVal = (((cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&) - &H6&) And &HF&)
        ahVal = (((cpu_getReg8High(cpu, CPU_REG_EAX) And &HFF&) - 1&) And &HFF&)
        cpu_setReg8Low cpu, CPU_REG_EAX, alVal
        cpu_setReg8High cpu, CPU_REG_EAX, ahVal
        cpu.af = 1&
        cpu.cf = 1&
    Else
        cpu.af = 0&
        cpu.cf = 0&
        cpu_setReg8Low cpu, CPU_REG_EAX, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HF&)
    End If
End Sub

Public Sub op_increg(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        cpu.oldcf = cpu.cf
        cpu.oper1_32 = getreg32(cpu, (cpu.opcode And &H7&))
        cpu.oper2_32 = 1&
        op_add32 cpu
        cpu.cf = cpu.oldcf
        putreg32 cpu, (cpu.opcode And &H7&), cpu.res32
    Else
        cpu.oldcf = cpu.cf
        cpu.oper1 = getreg16(cpu, (cpu.opcode And &H7&))
        cpu.oper2 = 1&
        op_add16 cpu
        cpu.cf = cpu.oldcf
        putreg16 cpu, (cpu.opcode And &H7&), cpu.res16
    End If
End Sub

Public Sub op_decreg(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        cpu.oldcf = cpu.cf
        cpu.oper1_32 = getreg32(cpu, (cpu.opcode And &H7&))
        cpu.oper2_32 = 1&
        op_sub32 cpu
        cpu.cf = cpu.oldcf
        putreg32 cpu, (cpu.opcode And &H7&), cpu.res32
    Else
        cpu.oldcf = cpu.cf
        cpu.oper1 = getreg16(cpu, (cpu.opcode And &H7&))
        cpu.oper2 = 1&
        op_sub16 cpu
        cpu.cf = cpu.oldcf
        putreg16 cpu, (cpu.opcode And &H7&), cpu.res16
    End If
End Sub

Public Sub op_pushreg(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        push cpu, getreg32(cpu, (cpu.opcode And &H7&))
    Else
        push cpu, getreg16(cpu, (cpu.opcode And &H7&))
    End If
End Sub

Public Sub op_popreg(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        putreg32 cpu, (cpu.opcode And &H7&), pop(cpu)
    Else
        putreg16 cpu, (cpu.opcode And &H7&), pop(cpu)
    End If
End Sub

Public Sub op_60(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        cpu.oldsp = cpu.regs_long(CPU_REG_ESP)
        push cpu, cpu.regs_long(CPU_REG_EAX)
        push cpu, cpu.regs_long(CPU_REG_ECX)
        push cpu, cpu.regs_long(CPU_REG_EDX)
        push cpu, cpu.regs_long(CPU_REG_EBX)
        push cpu, cpu.oldsp
        push cpu, cpu.regs_long(CPU_REG_EBP)
        push cpu, cpu.regs_long(CPU_REG_ESI)
        push cpu, cpu.regs_long(CPU_REG_EDI)
    Else
        cpu.oldsp = getreg16(cpu, CPU_REG_ESP)
        push cpu, getreg16(cpu, CPU_REG_EAX)
        push cpu, getreg16(cpu, CPU_REG_ECX)
        push cpu, getreg16(cpu, CPU_REG_EDX)
        push cpu, getreg16(cpu, CPU_REG_EBX)
        push cpu, cpu.oldsp
        push cpu, getreg16(cpu, CPU_REG_EBP)
        push cpu, getreg16(cpu, CPU_REG_ESI)
        push cpu, getreg16(cpu, CPU_REG_EDI)
    End If
End Sub

Public Sub op_61(ByRef cpu As CPU_t)
    Dim dummyVal As Long

    If cpu.isoper32 <> 0& Then
        cpu.regs_long(CPU_REG_EDI) = pop(cpu)
        cpu.regs_long(CPU_REG_ESI) = pop(cpu)
        cpu.regs_long(CPU_REG_EBP) = pop(cpu)
        dummyVal = pop(cpu)
        cpu.regs_long(CPU_REG_EBX) = pop(cpu)
        cpu.regs_long(CPU_REG_EDX) = pop(cpu)
        cpu.regs_long(CPU_REG_ECX) = pop(cpu)
        cpu.regs_long(CPU_REG_EAX) = pop(cpu)
    Else
        putreg16 cpu, CPU_REG_EDI, pop(cpu)
        putreg16 cpu, CPU_REG_ESI, pop(cpu)
        putreg16 cpu, CPU_REG_EBP, pop(cpu)
        dummyVal = pop(cpu)
        putreg16 cpu, CPU_REG_EBX, pop(cpu)
        putreg16 cpu, CPU_REG_EDX, pop(cpu)
        putreg16 cpu, CPU_REG_ECX, pop(cpu)
        putreg16 cpu, CPU_REG_EAX, pop(cpu)
    End If
End Sub

Public Sub op_62(ByRef cpu As CPU_t)
    Dim indexVal As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    modregrm cpu
    If cpu.mode = 3& Then
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If
    getea cpu, cpu.rm

    If cpu.isoper32 <> 0& Then
        indexVal = getreg32(cpu, cpu.reg)
        lowerBound = cpu_readl(cpu, cpu.ea)
        upperBound = cpu_readl(cpu, U32Add(cpu.ea, 4&))
    Else
        indexVal = cpu_signext16to32(getreg16(cpu, cpu.reg))
        lowerBound = cpu_signext16to32(cpu_readw(cpu, cpu.ea))
        upperBound = cpu_signext16to32(cpu_readw(cpu, U32Add(cpu.ea, 2&)))
    End If

    If (indexVal < lowerBound) Or (indexVal > upperBound) Then
        cpu_raiseException cpu, 5&, 0&
    End If
End Sub

Public Sub op_63(ByRef cpu As CPU_t)
    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    modregrm cpu
    cpu.oper1 = readrm16(cpu, cpu.rm)
    If cpu.doexception <> 0& Then Exit Sub
    cpu.oper2 = (getreg16(cpu, cpu.reg) And &H3&)

    If ((cpu.oper1 And &H3&) < cpu.oper2) Then
        cpu.res16 = ((cpu.oper1 And &HFFFC&) Or cpu.oper2)
        writerm16 cpu, cpu.rm, cpu.res16
        cpu.zf = 1&
    Else
        cpu.zf = 0&
    End If
End Sub

Public Sub op_68(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        push cpu, cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        push cpu, (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If
End Sub

Public Sub op_69(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim src1 As Long
    Dim src2 As Long
    Dim prod64 As U64_t

    modregrm cpu
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        src1 = readrm32(cpu, cpu.rm)
        src2 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&

        prod64 = cpu_signedMul32ToU64(src1, src2)
        putreg32 cpu, cpu.reg, prod64.Lo

        If cpu_u64SignExtMatches32(prod64) = 0& Then
            cpu.cf = 1&
            cpu.ofl = 1&
        Else
            cpu.cf = 0&
            cpu.ofl = 0&
        End If
    Else
        src1 = cpu_signext16to32(readrm16(cpu, cpu.rm))
        src2 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&

        cpu.temp3 = (src1 * src2)
        putreg16 cpu, cpu.reg, (cpu.temp3 And &HFFFF&)

        If (cpu.temp3 < -32768) Or (cpu.temp3 > 32767&) Then
            cpu.cf = 1&
            cpu.ofl = 1&
        Else
            cpu.cf = 0&
            cpu.ofl = 0&
        End If
    End If
End Sub

Public Sub op_6A(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    push cpu, cpu_signext8to32(cpu_read(cpu, codeAddr))
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_6B(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim src1 As Long
    Dim src2 As Long
    Dim prod64 As U64_t

    modregrm cpu
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        src1 = readrm32(cpu, cpu.rm)
        src2 = cpu_signext8to32(cpu_read(cpu, codeAddr))
        cpu_stepIP cpu, 1&

        prod64 = cpu_signedMul32ToU64(src1, src2)
        putreg32 cpu, cpu.reg, prod64.Lo

        If cpu_u64SignExtMatches32(prod64) = 0& Then
            cpu.cf = 1&
            cpu.ofl = 1&
        Else
            cpu.cf = 0&
            cpu.ofl = 0&
        End If
    Else
        src1 = cpu_signext16to32(readrm16(cpu, cpu.rm))
        src2 = cpu_signext16to32(cpu_signext8to16(cpu_read(cpu, codeAddr)))
        cpu_stepIP cpu, 1&

        cpu.temp3 = (src1 * src2)
        putreg16 cpu, cpu.reg, (cpu.temp3 And &HFFFF&)

        If (cpu.temp3 < -32768) Or (cpu.temp3 > 32767&) Then
            cpu.cf = 1&
            cpu.ofl = 1&
        Else
            cpu.cf = 0&
            cpu.ofl = 0&
        End If
    End If
End Sub

Public Sub op_A0(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim offs As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isaddr32 <> 0& Then
        offs = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        offs = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If

    cpu_setReg8Low cpu, CPU_REG_EAX, (cpu_read(cpu, U32Add(cpu.useseg, offs)) And &HFF&)
End Sub

Public Sub op_A1(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim offs As Long
    Dim addr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isaddr32 <> 0& Then
        offs = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        offs = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If

    addr = U32Add(cpu.useseg, offs)
    If cpu.isoper32 <> 0& Then
        cpu.regs_long(CPU_REG_EAX) = cpu_readl(cpu, addr)
    Else
        putreg16 cpu, CPU_REG_EAX, (cpu_readw(cpu, addr) And &HFFFF&)
    End If
End Sub

Public Sub op_A2(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim offs As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isaddr32 <> 0& Then
        offs = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        offs = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If

    cpu_write cpu, U32Add(cpu.useseg, offs), CByte(cpu_getReg8Low(cpu, CPU_REG_EAX))
End Sub

Public Sub op_A3(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim offs As Long
    Dim addr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isaddr32 <> 0& Then
        offs = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        offs = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If

    addr = U32Add(cpu.useseg, offs)
    If cpu.isoper32 <> 0& Then
        cpu_writel cpu, addr, cpu.regs_long(CPU_REG_EAX)
    Else
        cpu_writew cpu, addr, getreg16(cpu, CPU_REG_EAX)
    End If
End Sub

Public Sub op_A4(ByRef cpu As CPU_t)
    Dim srcAddr As Long
    Dim dstAddr As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        srcAddr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        If (cpu.reptype <> 0&) And (getreg16(cpu, CPU_REG_ECX) = 0&) Then Exit Sub
        srcAddr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    cpu_write cpu, dstAddr, (cpu_read(cpu, srcAddr) And &HFF&)

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), 1&)
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - 1&)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), 1&)
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + 1&)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_A5(ByRef cpu As CPU_t)
    Dim srcAddr As Long
    Dim dstAddr As Long
    Dim stepVal As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            If getreg16(cpu, CPU_REG_ECX) = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        srcAddr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        srcAddr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    If cpu.isoper32 <> 0& Then
        cpu_writel cpu, dstAddr, cpu_readl(cpu, srcAddr)
        stepVal = 4&
    Else
        cpu_writew cpu, dstAddr, (cpu_readw(cpu, srcAddr) And &HFFFF&)
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), stepVal)
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - stepVal)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), stepVal)
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + stepVal)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_A6(ByRef cpu As CPU_t)
    Dim srcAddr As Long
    Dim dstAddr As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        srcAddr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        If (cpu.reptype <> 0&) And (getreg16(cpu, CPU_REG_ECX) = 0&) Then Exit Sub
        srcAddr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    cpu.oper1b = (cpu_read(cpu, srcAddr) And &HFF&)
    cpu.oper2b = (cpu_read(cpu, dstAddr) And &HFF&)
    flag_sub8 cpu, cpu.oper1b, cpu.oper2b

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), 1&)
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - 1&)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), 1&)
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + 1&)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If (cpu.reptype = 1&) And (cpu.zf = 0&) Then Exit Sub
    If (cpu.reptype = 2&) And (cpu.zf <> 0&) Then Exit Sub
    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_A7(ByRef cpu As CPU_t)
    Dim srcAddr As Long
    Dim dstAddr As Long
    Dim stepVal As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            If getreg16(cpu, CPU_REG_ECX) = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        srcAddr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        srcAddr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
        dstAddr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_readl(cpu, srcAddr)
        cpu.oper2_32 = cpu_readl(cpu, dstAddr)
        flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
        stepVal = 4&
    Else
        cpu.oper1 = (cpu_readw(cpu, srcAddr) And &HFFFF&)
        cpu.oper2 = (cpu_readw(cpu, dstAddr) And &HFFFF&)
        flag_sub16 cpu, cpu.oper1, cpu.oper2
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), stepVal)
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - stepVal)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), stepVal)
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + stepVal)
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If (cpu.reptype = 1&) And (cpu.zf = 0&) Then Exit Sub
    If (cpu.reptype = 2&) And (cpu.zf <> 0&) Then Exit Sub
    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_A8(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    flag_log8 cpu, ((cpu.oper1b And &HFF&) And (cpu.oper2b And &HFF&))
End Sub

Public Sub op_A9(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        flag_log32 cpu, (cpu.oper1_32 And cpu.oper2_32)
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        flag_log16 cpu, ((cpu.oper1 And &HFFFF&) And (cpu.oper2 And &HFFFF&))
    End If
End Sub

Public Sub op_AA(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim countVal As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        countVal = getreg16(cpu, CPU_REG_ECX)
        If (cpu.reptype <> 0&) And (countVal = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    cpu_write cpu, addr, (cpu_getReg8Low(cpu, CPU_REG_EAX) And &HFF&)

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_AB(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim countVal As Long
    Dim stepVal As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            countVal = getreg16(cpu, CPU_REG_ECX)
            If countVal = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        addr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        addr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    If cpu.isoper32 <> 0& Then
        cpu_writel cpu, addr, cpu.regs_long(CPU_REG_EAX)
        stepVal = 4&
    Else
        cpu_writew cpu, addr, getreg16(cpu, CPU_REG_EAX)
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_AC(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim countVal As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        addr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
    Else
        countVal = getreg16(cpu, CPU_REG_ECX)
        If (cpu.reptype <> 0&) And (countVal = 0&) Then Exit Sub
        addr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
    End If

    cpu_setReg8Low cpu, CPU_REG_EAX, (cpu_read(cpu, addr) And &HFF&)

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), 1&)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_AD(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim countVal As Long
    Dim stepVal As Long

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            If cpu.regs_long(CPU_REG_ECX) = 0& Then Exit Sub
        Else
            countVal = getreg16(cpu, CPU_REG_ECX)
            If countVal = 0& Then Exit Sub
        End If
    End If

    If cpu.isaddr32 <> 0& Then
        addr = U32Add(cpu.useseg, cpu.regs_long(CPU_REG_ESI))
    Else
        addr = U32Add(cpu.useseg, getreg16(cpu, CPU_REG_ESI))
    End If

    If cpu.isoper32 <> 0& Then
        cpu.regs_long(CPU_REG_EAX) = cpu_readl(cpu, addr)
        stepVal = 4&
    Else
        putreg16 cpu, CPU_REG_EAX, (cpu_readw(cpu, addr) And &HFFFF&)
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Sub(cpu.regs_long(CPU_REG_ESI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ESI) = U32Add(cpu.regs_long(CPU_REG_ESI), stepVal)
        Else
            putreg16 cpu, CPU_REG_ESI, (getreg16(cpu, CPU_REG_ESI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_AE(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim countVal As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        countVal = getreg16(cpu, CPU_REG_ECX)
        If (cpu.reptype <> 0&) And (countVal = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = (cpu_read(cpu, addr) And &HFF&)
    flag_sub8 cpu, cpu.oper1b, cpu.oper2b

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - 1&)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), 1&)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + 1&)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If (cpu.reptype = 1&) And (cpu.zf = 0&) Then Exit Sub
    If (cpu.reptype = 2&) And (cpu.zf <> 0&) Then Exit Sub
    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_AF(ByRef cpu As CPU_t)
    Dim addr As Long
    Dim countVal As Long
    Dim stepVal As Long

    If cpu.isaddr32 <> 0& Then
        If (cpu.reptype <> 0&) And (cpu.regs_long(CPU_REG_ECX) = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), cpu.regs_long(CPU_REG_EDI))
    Else
        countVal = getreg16(cpu, CPU_REG_ECX)
        If (cpu.reptype <> 0&) And (countVal = 0&) Then Exit Sub
        addr = U32Add(cpu.segcache(CPU_REG_ES), getreg16(cpu, CPU_REG_EDI))
    End If

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = cpu_readl(cpu, addr)
        flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
        stepVal = 4&
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = (cpu_readw(cpu, addr) And &HFFFF&)
        flag_sub16 cpu, cpu.oper1, cpu.oper2
        stepVal = 2&
    End If

    If cpu.df <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Sub(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) - stepVal)
        End If
    Else
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_EDI) = U32Add(cpu.regs_long(CPU_REG_EDI), stepVal)
        Else
            putreg16 cpu, CPU_REG_EDI, (getreg16(cpu, CPU_REG_EDI) + stepVal)
        End If
    End If

    If cpu.reptype <> 0& Then
        If cpu.isaddr32 <> 0& Then
            cpu.regs_long(CPU_REG_ECX) = U32Sub(cpu.regs_long(CPU_REG_ECX), 1&)
        Else
            putreg16 cpu, CPU_REG_ECX, (getreg16(cpu, CPU_REG_ECX) - 1&)
        End If
    End If

    If (cpu.reptype = 1&) And (cpu.zf = 0&) Then Exit Sub
    If (cpu.reptype = 2&) And (cpu.zf <> 0&) Then Exit Sub
    If cpu.reptype = 0& Then Exit Sub
    cpu.ip = cpu_firstip
End Sub

Public Sub op_B0(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8Low cpu, CPU_REG_EAX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B1(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8Low cpu, CPU_REG_ECX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B2(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8Low cpu, CPU_REG_EDX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B3(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8Low cpu, CPU_REG_EBX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B4(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8High cpu, CPU_REG_EAX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B5(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8High cpu, CPU_REG_ECX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B6(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8High cpu, CPU_REG_EDX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_B7(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu_setReg8High cpu, CPU_REG_EBX, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_mov_regimm(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
        putreg32 cpu, (cpu.opcode And &H7&), cpu.oper1_32
    Else
        cpu.oper1 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
        putreg16 cpu, (cpu.opcode And &H7&), cpu.oper1
    End If
End Sub

Public Sub op_C0(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper2b = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
    writerm8 cpu, cpu.rm, op_grp2_8(cpu, cpu.oper2b)
End Sub

Public Sub op_C1(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    modregrm cpu
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = (cpu_read(cpu, codeAddr) And &HFF&)
        writerm32 cpu, cpu.rm, op_grp2_32(cpu, cpu.oper2_32)
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = (cpu_read(cpu, codeAddr) And &HFF&)
        writerm16 cpu, cpu.rm, op_grp2_16(cpu, cpu.oper2)
    End If

    cpu_stepIP cpu, 1&
End Sub

Public Sub op_C4(ByRef cpu As CPU_t)
    modregrm cpu
    getea cpu, cpu.rm

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, cpu_readl(cpu, cpu.ea)
        putsegreg cpu, CPU_REG_ES, (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&)
    Else
        putreg16 cpu, cpu.reg, (cpu_readw(cpu, cpu.ea) And &HFFFF&)
        putsegreg cpu, CPU_REG_ES, (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&)
    End If
End Sub

Public Sub op_C5(ByRef cpu As CPU_t)
    modregrm cpu
    getea cpu, cpu.rm

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, cpu_readl(cpu, cpu.ea)
        putsegreg cpu, CPU_REG_DS, (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&)
    Else
        putreg16 cpu, cpu.reg, (cpu_readw(cpu, cpu.ea) And &HFFFF&)
        putsegreg cpu, CPU_REG_DS, (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&)
    End If
End Sub

Public Sub op_C6(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    modregrm cpu
    If cpu.reg <> 0& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    writerm8 cpu, cpu.rm, (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&
End Sub

Public Sub op_C7(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    modregrm cpu
    If cpu.reg <> 0& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        writerm32 cpu, cpu.rm, cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        writerm16 cpu, cpu.rm, (cpu_readw(cpu, codeAddr) And &HFFFF&)
        cpu_stepIP cpu, 2&
    End If
End Sub

Private Function cpu_selector_offset(ByVal selector As Long) As Long
    cpu_selector_offset = (selector And &HFFF8&)
End Function

Private Function cpu_selector_error_code(ByVal selector As Long) As Long
    cpu_selector_error_code = (selector And &HFFFC&)
End Function

Private Function cpu_normalize_selector(ByVal selector As Long, ByVal rpl As Long) As Long
    cpu_normalize_selector = ((selector And &HFFFC&) Or (rpl And &H3&))
End Function

Private Function cpu_idt_error_code(ByVal intnum As Long, ByVal Source As Long) As Long
    cpu_idt_error_code = (U32Shl((intnum And &HFF&), 3&) Or 2&)
    If (Source <> INT_SOURCE_SOFTWARE) And (Source <> INT_SOURCE_INT3) And (Source <> INT_SOURCE_INTO) Then
        cpu_idt_error_code = (cpu_idt_error_code Or 1&)
    End If
End Function

Private Function cpu_selector_table_base(ByRef cpu As CPU_t, ByVal selector As Long) As Long
    If (selector And &H4&) <> 0& Then
        cpu_selector_table_base = cpu.ldtr
    Else
        cpu_selector_table_base = cpu.gdtr
    End If
End Function

Private Function cpu_selector_table_limit(ByRef cpu As CPU_t, ByVal selector As Long) As Long
    If (selector And &H4&) <> 0& Then
        cpu_selector_table_limit = cpu.ldtl
    Else
        cpu_selector_table_limit = cpu.gdtl
    End If
End Function

Private Function cpu_read_segdesc(ByRef cpu As CPU_t, ByVal selector As Long, ByRef desc As CPU_SEGDESC_t) As Long
    Dim offset As Long

    offset = cpu_selector_offset(selector)
    If U32Add(offset, 7&) > cpu_selector_table_limit(cpu, selector) Then
        cpu_read_segdesc = 0&
        Exit Function
    End If

    desc.addr = U32Add(cpu_selector_table_base(cpu, selector), offset)
    desc.access = (cpu_read_sys(cpu, U32Add(desc.addr, 5&)) And &HFF&)
    desc.flags = (cpu_read_sys(cpu, U32Add(desc.addr, 6&)) And &HFF&)
    desc.dpl = (U32Shr(desc.access, 5&) And &H3&)
    cpu_read_segdesc = 1&
End Function

Private Function cpu_desc_is_present(ByRef desc As CPU_SEGDESC_t) As Long
    If (desc.access And &H80&) <> 0& Then
        cpu_desc_is_present = 1&
    Else
        cpu_desc_is_present = 0&
    End If
End Function

Private Function cpu_desc_is_system(ByRef desc As CPU_SEGDESC_t) As Long
    If (desc.access And &H10&) = 0& Then
        cpu_desc_is_system = 1&
    Else
        cpu_desc_is_system = 0&
    End If
End Function

Private Function cpu_desc_is_code(ByRef desc As CPU_SEGDESC_t) As Long
    If (desc.access And &H18&) = &H18& Then
        cpu_desc_is_code = 1&
    Else
        cpu_desc_is_code = 0&
    End If
End Function

Private Function cpu_desc_is_conforming_code(ByRef desc As CPU_SEGDESC_t) As Long
    If (cpu_desc_is_code(desc) <> 0&) And ((desc.access And &H4&) <> 0&) Then
        cpu_desc_is_conforming_code = 1&
    Else
        cpu_desc_is_conforming_code = 0&
    End If
End Function

Private Function cpu_desc_is_readable_code(ByRef desc As CPU_SEGDESC_t) As Long
    If (cpu_desc_is_code(desc) <> 0&) And ((desc.access And &H2&) <> 0&) Then
        cpu_desc_is_readable_code = 1&
    Else
        cpu_desc_is_readable_code = 0&
    End If
End Function

Private Function cpu_desc_is_writable_data(ByRef desc As CPU_SEGDESC_t) As Long
    If ((desc.access And &H18&) = &H10&) And ((desc.access And &H2&) <> 0&) Then
        cpu_desc_is_writable_data = 1&
    Else
        cpu_desc_is_writable_data = 0&
    End If
End Function

Private Function cpu_desc_type(ByRef desc As CPU_SEGDESC_t) As Long
    cpu_desc_type = (desc.access And &HF&)
End Function

Private Function cpu_desc_is_call_gate(ByRef desc As CPU_SEGDESC_t) As Long
    If (cpu_desc_is_system(desc) <> 0&) And ((cpu_desc_type(desc) = &H4&) Or (cpu_desc_type(desc) = &HC&)) Then
        cpu_desc_is_call_gate = 1&
    Else
        cpu_desc_is_call_gate = 0&
    End If
End Function

Private Function cpu_desc_is_task_gate(ByRef desc As CPU_SEGDESC_t) As Long
    If (cpu_desc_is_system(desc) <> 0&) And (cpu_desc_type(desc) = &H5&) Then
        cpu_desc_is_task_gate = 1&
    Else
        cpu_desc_is_task_gate = 0&
    End If
End Function

Private Function cpu_desc_is_tss(ByRef desc As CPU_SEGDESC_t) As Long
    If (cpu_desc_is_system(desc) <> 0&) And ((cpu_desc_type(desc) = &H9&) Or (cpu_desc_type(desc) = &HB&)) Then
        cpu_desc_is_tss = 1&
    Else
        cpu_desc_is_tss = 0&
    End If
End Function

Private Function cpu_desc_is_busy_tss(ByRef desc As CPU_SEGDESC_t) As Long
    If (cpu_desc_is_system(desc) <> 0&) And (cpu_desc_type(desc) = &HB&) Then
        cpu_desc_is_busy_tss = 1&
    Else
        cpu_desc_is_busy_tss = 0&
    End If
End Function

Private Function cpu_gate_is_interrupt_or_trap(ByRef gate As CPU_GATEDESC_t) As Long
    Select Case gate.typeVal
        Case &H6&, &H7&, &HE&, &HF&
            cpu_gate_is_interrupt_or_trap = 1&
        Case Else
            cpu_gate_is_interrupt_or_trap = 0&
    End Select
End Function

Private Function cpu_gate_is_interrupt(ByRef gate As CPU_GATEDESC_t) As Long
    If (gate.typeVal = &H6&) Or (gate.typeVal = &HE&) Then
        cpu_gate_is_interrupt = 1&
    Else
        cpu_gate_is_interrupt = 0&
    End If
End Function

Private Function cpu_gate_is_32bit(ByRef gate As CPU_GATEDESC_t) As Long
    If (gate.typeVal = &HC&) Or (gate.typeVal = &HE&) Or (gate.typeVal = &HF&) Then
        cpu_gate_is_32bit = 1&
    Else
        cpu_gate_is_32bit = 0&
    End If
End Function

Private Function cpu_exception_has_error_code(ByVal intnum As Long) As Long
    Select Case (intnum And &HFF&)
        Case 8&, 10&, 11&, 12&, 13&, 14&, 17&
            cpu_exception_has_error_code = 1&
        Case Else
            cpu_exception_has_error_code = 0&
    End Select
End Function

Private Function cpu_stack_ptr_value(ByRef cpu As CPU_t) As Long
    If cpu.segis32(CPU_REG_SS) <> 0& Then
        cpu_stack_ptr_value = cpu.regs_long(CPU_REG_ESP)
    Else
        cpu_stack_ptr_value = (cpu.regs_long(CPU_REG_ESP) And &HFFFF&)
    End If
End Function

Private Function cpu_stack_ptr_add(ByRef cpu As CPU_t, ByVal sp As Long, ByVal addval As Long) As Long
    If cpu.segis32(CPU_REG_SS) <> 0& Then
        cpu_stack_ptr_add = U32Add(sp, addval)
    Else
        cpu_stack_ptr_add = ((sp + addval) And &HFFFF&)
    End If
End Function

Private Sub cpu_stack_pushw_sys(ByRef cpu As CPU_t, ByVal value As Long)
    Dim sp As Long
    Dim addr As Long

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        sp = U32Sub(cpu.regs_long(CPU_REG_ESP), 2&)
        cpu.regs_long(CPU_REG_ESP) = sp
    Else
        sp = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF&) - 2&) And &HFFFF&
        cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or sp)
    End If

    addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
    cpu_writew_sys cpu, addr, (value And &HFFFF&)
End Sub

Private Sub cpu_stack_pushl_sys(ByRef cpu As CPU_t, ByVal value As Long)
    Dim sp As Long
    Dim addr As Long

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        sp = U32Sub(cpu.regs_long(CPU_REG_ESP), 4&)
        cpu.regs_long(CPU_REG_ESP) = sp
    Else
        sp = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF&) - 4&) And &HFFFF&
        cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or sp)
    End If

    addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
    cpu_writel_sys cpu, addr, value
End Sub

Private Sub cpu_push_far_return16(ByRef cpu As CPU_t, ByVal old_cs As Long, ByVal old_ip As Long)
    pushw cpu, (old_cs And &HFFFF&)
    pushw cpu, (old_ip And &HFFFF&)
End Sub

Private Sub cpu_push_far_return32(ByRef cpu As CPU_t, ByVal old_cs As Long, ByVal old_ip As Long)
    pushl cpu, (old_cs And &HFFFF&)
    pushl cpu, old_ip
End Sub

Private Sub cpu_push_far_return16_sys(ByRef cpu As CPU_t, ByVal old_cs As Long, ByVal old_ip As Long)
    cpu_stack_pushw_sys cpu, (old_cs And &HFFFF&)
    cpu_stack_pushw_sys cpu, (old_ip And &HFFFF&)
End Sub

Private Sub cpu_push_far_return32_sys(ByRef cpu As CPU_t, ByVal old_cs As Long, ByVal old_ip As Long)
    cpu_stack_pushl_sys cpu, (old_cs And &HFFFF&)
    cpu_stack_pushl_sys cpu, old_ip
End Sub

Private Sub cpu_copy_call_gate_params_sys(ByRef cpu As CPU_t, ByVal old_ss_base As Long, ByVal old_esp As Long, ByVal old_stack_is32 As Long, ByVal param_count As Long, ByVal gate32 As Long)
    Dim params(0& To 30&) As Long
    Dim i As Long
    Dim offset As Long
    Dim param_size As Long
    Dim count As Long

    count = (param_count And &H1F&)
    If count <= 0& Then Exit Sub
    If count > 31& Then count = 31&

    If gate32 <> 0& Then
        param_size = 4&
    Else
        param_size = 2&
    End If

    For i = 0& To count - 1&
        If old_stack_is32 <> 0& Then
            offset = U32Add(old_esp, (i * param_size))
        Else
            offset = ((old_esp + (i * param_size)) And &HFFFF&)
        End If

        If gate32 <> 0& Then
            params(i) = cpu_readl_sys(cpu, U32Add(old_ss_base, offset))
        Else
            params(i) = (cpu_readw_sys(cpu, U32Add(old_ss_base, offset)) And &HFFFF&)
        End If
    Next i

    For i = count - 1& To 0& Step -1&
        If gate32 <> 0& Then
            cpu_stack_pushl_sys cpu, params(i)
        Else
            cpu_stack_pushw_sys cpu, params(i)
        End If
    Next i
End Sub

Private Sub cpu_load_gate_desc(ByRef cpu As CPU_t, ByVal addr As Long, ByVal access As Long, ByVal flags As Long, ByRef gate As CPU_GATEDESC_t)
    gate.addr = addr
    gate.access = (access And &HFF&)
    gate.flags = (flags And &HFF&)
    gate.typeVal = (gate.access And &HF&)
    gate.dpl = (U32Shr(gate.access, 5&) And &H3&)
    gate.present = (U32Shr(gate.access, 7&) And &H1&)
    gate.param_count = (cpu_read_sys(cpu, U32Add(addr, 4&)) And &H1F&)
    gate.target_selector = (cpu_readw_sys(cpu, U32Add(addr, 2&)) And &HFFFF&)
    gate.offset = (cpu_readw_sys(cpu, addr) And &HFFFF&)
    If cpu_gate_is_32bit(gate) <> 0& Then
        gate.offset = (gate.offset Or U32Shl((cpu_readw_sys(cpu, U32Add(addr, 6&)) And &HFFFF&), 16&))
    End If
End Sub

Private Function cpu_read_idt_gate(ByRef cpu As CPU_t, ByVal intnum As Long, ByRef gate As CPU_GATEDESC_t) As Long
    Dim offset As Long
    Dim gateAddr As Long

    offset = U32Shl((intnum And &HFF&), 3&)
    If U32Add(offset, 7&) > cpu.idtl Then
        cpu_read_idt_gate = 0&
        Exit Function
    End If

    gateAddr = U32Add(cpu.idtr, offset)
    cpu_load_gate_desc cpu, gateAddr, (cpu_read_sys(cpu, U32Add(gateAddr, 5&)) And &HFF&), (cpu_read_sys(cpu, U32Add(gateAddr, 6&)) And &HFF&), gate
    cpu_read_idt_gate = 1&
End Function

Private Function cpu_validate_stack_selector(ByRef cpu As CPU_t, ByVal selector As Long, ByVal target_cpl As Long, ByVal invalid_vector As Long, ByVal not_present_vector As Long) As Long
    Dim desc As CPU_SEGDESC_t

    selector = (selector And &HFFFF&)
    target_cpl = (target_cpl And &H3&)

    If cpu_selector_offset(selector) = 0& Then
        cpu_raiseException cpu, invalid_vector, 0&
        cpu_validate_stack_selector = 0&
        Exit Function
    End If

    If cpu_read_segdesc(cpu, selector, desc) = 0& Then
        cpu_raiseException cpu, invalid_vector, cpu_selector_error_code(selector)
        cpu_validate_stack_selector = 0&
        Exit Function
    End If

    If cpu_desc_is_writable_data(desc) = 0& Then
        cpu_raiseException cpu, invalid_vector, cpu_selector_error_code(selector)
        cpu_validate_stack_selector = 0&
        Exit Function
    End If

    If cpu_desc_is_present(desc) = 0& Then
        cpu_raiseException cpu, not_present_vector, cpu_selector_error_code(selector)
        cpu_validate_stack_selector = 0&
        Exit Function
    End If

    If ((selector And &H3&) <> target_cpl) Or (desc.dpl <> target_cpl) Then
        cpu_raiseException cpu, invalid_vector, cpu_selector_error_code(selector)
        cpu_validate_stack_selector = 0&
        Exit Function
    End If

    cpu_validate_stack_selector = 1&
End Function

Private Function cpu_read_code_desc(ByRef cpu As CPU_t, ByVal selector As Long, ByVal not_present_vector As Long, ByRef desc As CPU_SEGDESC_t) As Long
    selector = (selector And &HFFFF&)

    If cpu_selector_offset(selector) = 0& Then
        cpu_raiseException cpu, 13&, 0&
        cpu_read_code_desc = 0&
        Exit Function
    End If

    If cpu_read_segdesc(cpu, selector, desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        cpu_read_code_desc = 0&
        Exit Function
    End If

    If cpu_desc_is_code(desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        cpu_read_code_desc = 0&
        Exit Function
    End If

    If cpu_desc_is_present(desc) = 0& Then
        cpu_raiseException cpu, not_present_vector, cpu_selector_error_code(selector)
        cpu_read_code_desc = 0&
        Exit Function
    End If

    cpu_read_code_desc = 1&
End Function

Private Function cpu_validate_gate_target_code(ByRef cpu As CPU_t, ByVal selector As Long, ByRef target As CPU_CODETARGET_t) As Long
    Dim desc As CPU_SEGDESC_t

    selector = (selector And &HFFFF&)

    If cpu_read_code_desc(cpu, selector, 11&, desc) = 0& Then
        cpu_validate_gate_target_code = 0&
        Exit Function
    End If

    target.conforming = cpu_desc_is_conforming_code(desc)
    If target.conforming <> 0& Then
        If desc.dpl > cpu.cpl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_gate_target_code = 0&
            Exit Function
        End If
        target.target_cpl = cpu.cpl
        target.outer = 0&
    Else
        If desc.dpl > cpu.cpl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_gate_target_code = 0&
            Exit Function
        End If
        target.target_cpl = desc.dpl
        If desc.dpl < cpu.cpl Then
            target.outer = 1&
        Else
            target.outer = 0&
        End If
    End If

    target.selector = cpu_normalize_selector(selector, target.target_cpl)
    cpu_validate_gate_target_code = 1&
End Function

Private Function cpu_validate_direct_call_target(ByRef cpu As CPU_t, ByVal selector As Long, ByRef target As CPU_CODETARGET_t) As Long
    Dim desc As CPU_SEGDESC_t
    Dim rpl As Long
    Dim epl As Long

    selector = (selector And &HFFFF&)
    rpl = (selector And &H3&)
    If rpl > cpu.cpl Then
        epl = rpl
    Else
        epl = cpu.cpl
    End If

    If cpu_read_code_desc(cpu, selector, 11&, desc) = 0& Then
        cpu_validate_direct_call_target = 0&
        Exit Function
    End If

    target.conforming = cpu_desc_is_conforming_code(desc)
    If target.conforming <> 0& Then
        If desc.dpl > epl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_direct_call_target = 0&
            Exit Function
        End If
    Else
        If (desc.dpl <> cpu.cpl) Or (rpl > cpu.cpl) Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_direct_call_target = 0&
            Exit Function
        End If
    End If

    target.selector = cpu_normalize_selector(selector, cpu.cpl)
    target.target_cpl = cpu.cpl
    target.outer = 0&
    cpu_validate_direct_call_target = 1&
End Function

Private Function cpu_resolve_task_switch_target(ByRef cpu As CPU_t, ByVal selector As Long, ByRef first_desc As CPU_SEGDESC_t, ByRef task_selector As Long) As Long
    Dim desc As CPU_SEGDESC_t
    Dim gate As CPU_GATEDESC_t
    Dim epl As Long
    Dim target_selector As Long
    Dim rpl As Long

    selector = (selector And &HFFFF&)
    rpl = (selector And &H3&)
    If rpl > cpu.cpl Then
        epl = rpl
    Else
        epl = cpu.cpl
    End If

    desc = first_desc
    target_selector = selector

    If cpu_desc_is_task_gate(desc) <> 0& Then
        cpu_load_gate_desc cpu, desc.addr, desc.access, desc.flags, gate
        If epl > gate.dpl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If
        If gate.present = 0& Then
            cpu_raiseException cpu, 11&, cpu_selector_error_code(selector)
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If

        target_selector = (gate.target_selector And &HFFFF&)
        If cpu_selector_offset(target_selector) = 0& Then
            cpu_raiseException cpu, 13&, 0&
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If
        If (target_selector And &H4&) <> 0& Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(target_selector)
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If
        If cpu_read_segdesc(cpu, target_selector, desc) = 0& Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(target_selector)
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If
    Else
        If (target_selector And &H4&) <> 0& Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(target_selector)
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If
        If epl > desc.dpl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_resolve_task_switch_target = 0&
            Exit Function
        End If
    End If

    If cpu_desc_is_tss(desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(target_selector)
        cpu_resolve_task_switch_target = 0&
        Exit Function
    End If
    If cpu_desc_is_present(desc) = 0& Then
        cpu_raiseException cpu, 11&, cpu_selector_error_code(target_selector)
        cpu_resolve_task_switch_target = 0&
        Exit Function
    End If
    If cpu_desc_is_busy_tss(desc) <> 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(target_selector)
        cpu_resolve_task_switch_target = 0&
        Exit Function
    End If

    task_selector = (target_selector And &HFFFC&)
    cpu_resolve_task_switch_target = 1&
End Function

Private Function cpu_fetch_tss_stack(ByRef cpu As CPU_t, ByVal target_cpl As Long, ByRef new_ss As Long, ByRef new_esp As Long) As Long
    Dim addval As Long

    target_cpl = (target_cpl And &H3&)
    If target_cpl > 2& Then
        cpu_raiseException cpu, 10&, cpu.tr_selector
        cpu_fetch_tss_stack = 0&
        Exit Function
    End If

    addval = U32Shl(target_cpl, 3&)
    If (cpu.trlimit <> 0&) And (cpu.trlimit < (10& + addval)) Then
        cpu_raiseException cpu, 10&, cpu.tr_selector
        cpu_fetch_tss_stack = 0&
        Exit Function
    End If

    new_esp = cpu_readl_sys(cpu, U32Add(U32Add(cpu.trbase, 4&), addval))
    new_ss = (cpu_readw_sys(cpu, U32Add(U32Add(cpu.trbase, 8&), addval)) And &HFFFF&)
    cpu_fetch_tss_stack = cpu_validate_stack_selector(cpu, new_ss, target_cpl, 10&, 12&)
End Function

Private Function cpu_validate_return_cs(ByRef cpu As CPU_t, ByVal selector As Long, ByRef retinfo As CPU_RETURNCS_t) As Long
    Dim desc As CPU_SEGDESC_t
    Dim rpl As Long

    selector = (selector And &HFFFF&)

    If cpu_selector_offset(selector) = 0& Then
        cpu_raiseException cpu, 13&, 0&
        cpu_validate_return_cs = 0&
        Exit Function
    End If

    If cpu_read_segdesc(cpu, selector, desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        cpu_validate_return_cs = 0&
        Exit Function
    End If

    If cpu_desc_is_code(desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        cpu_validate_return_cs = 0&
        Exit Function
    End If

    If cpu_desc_is_present(desc) = 0& Then
        cpu_raiseException cpu, 11&, cpu_selector_error_code(selector)
        cpu_validate_return_cs = 0&
        Exit Function
    End If

    rpl = (selector And &H3&)
    If cpu_desc_is_conforming_code(desc) <> 0& Then
        If (rpl <> cpu.cpl) Or (desc.dpl > cpu.cpl) Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_return_cs = 0&
            Exit Function
        End If
        retinfo.target_cpl = cpu.cpl
        retinfo.outer = 0&
    Else
        If rpl < cpu.cpl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_return_cs = 0&
            Exit Function
        End If
        If desc.dpl <> rpl Then
            cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
            cpu_validate_return_cs = 0&
            Exit Function
        End If
        retinfo.target_cpl = rpl
        If rpl > cpu.cpl Then
            retinfo.outer = 1&
        Else
            retinfo.outer = 0&
        End If
    End If

    retinfo.selector = selector
    cpu_validate_return_cs = 1&
End Function

Private Function cpu_validate_return_ss(ByRef cpu As CPU_t, ByVal selector As Long, ByVal target_cpl As Long) As Long
    cpu_validate_return_ss = cpu_validate_stack_selector(cpu, selector, target_cpl, 13&, 12&)
End Function

Private Sub cpu_restore_iret_flags(ByRef cpu As CPU_t, ByVal new_eflags As Long, ByVal current_cpl As Long, ByVal target_cpl As Long)
    Dim old_eflags As Long
    Dim merged_flags As Long
    Dim modifiable_mask As Long

    old_eflags = makeflagsword(cpu)
    merged_flags = old_eflags
    modifiable_mask = (EFLAGS_CF Or EFLAGS_PF Or EFLAGS_AF Or EFLAGS_ZF Or EFLAGS_SF Or EFLAGS_TF Or EFLAGS_DF Or EFLAGS_OF Or EFLAGS_NT)

    If cpu.isoper32 <> 0& Then
        modifiable_mask = (modifiable_mask Or EFLAGS_AC Or EFLAGS_ID)
    Else
        new_eflags = ((new_eflags And &HFFFF&) Or (old_eflags And &HFFFF0000))
    End If

    merged_flags = ((merged_flags And Not modifiable_mask) Or (new_eflags And modifiable_mask))

    If (cpu.protected_mode = 0&) Or (current_cpl <= cpu.iopl) Then
        merged_flags = ((merged_flags And Not EFLAGS_IF) Or (new_eflags And EFLAGS_IF))
    End If
    If (cpu.protected_mode = 0&) Or (current_cpl = 0&) Then
        merged_flags = ((merged_flags And Not EFLAGS_IOPL) Or (new_eflags And EFLAGS_IOPL))
    End If

    If (merged_flags And EFLAGS_CF) <> 0& Then cpu.cf = 1& Else cpu.cf = 0&
    If (merged_flags And EFLAGS_PF) <> 0& Then cpu.pf = 1& Else cpu.pf = 0&
    If (merged_flags And EFLAGS_AF) <> 0& Then cpu.af = 1& Else cpu.af = 0&
    If (merged_flags And EFLAGS_ZF) <> 0& Then cpu.zf = 1& Else cpu.zf = 0&
    If (merged_flags And EFLAGS_SF) <> 0& Then cpu.sf = 1& Else cpu.sf = 0&
    If (merged_flags And EFLAGS_TF) <> 0& Then cpu.tf = 1& Else cpu.tf = 0&
    If (merged_flags And EFLAGS_IF) <> 0& Then cpu.ifl = 1& Else cpu.ifl = 0&
    If (merged_flags And EFLAGS_DF) <> 0& Then cpu.df = 1& Else cpu.df = 0&
    If (merged_flags And EFLAGS_OF) <> 0& Then cpu.ofl = 1& Else cpu.ofl = 0&
    cpu.iopl = CByte((U32Shr(merged_flags, 12&) And &H3&))
    If (merged_flags And EFLAGS_NT) <> 0& Then cpu.nt = 1& Else cpu.nt = 0&
    If (merged_flags And EFLAGS_RF) <> 0& Then cpu.rf = 1& Else cpu.rf = 0&
    If (merged_flags And EFLAGS_VM) <> 0& Then cpu.v86f = 1& Else cpu.v86f = 0&
    If (merged_flags And EFLAGS_AC) <> 0& Then cpu.acf = 1& Else cpu.acf = 0&
    If (merged_flags And EFLAGS_ID) <> 0& Then cpu.idf = 1& Else cpu.idf = 0&

    If cpu.v86f <> 0& Then
        cpu.cpl = 3&
    Else
        cpu.cpl = CByte(target_cpl And &H3&)
    End If
End Sub

Private Sub cpu_null_seg(ByRef cpu As CPU_t, ByVal reg As Long)
    cpu.segregs(reg) = 0&
    cpu.segcache(reg) = 0&
    cpu.segis32(reg) = 0&
    cpu.seglimit(reg) = 0&
End Sub

Private Function cpu_segment_usable_at_cpl(ByRef cpu As CPU_t, ByVal selector As Long, ByVal target_cpl As Long) As Long
    Dim desc As CPU_SEGDESC_t
    Dim rpl As Long

    selector = (selector And &HFFFF&)
    target_cpl = (target_cpl And &H3&)

    If cpu_selector_offset(selector) = 0& Then
        cpu_segment_usable_at_cpl = 1&
        Exit Function
    End If

    If cpu_read_segdesc(cpu, selector, desc) = 0& Then
        cpu_segment_usable_at_cpl = 0&
        Exit Function
    End If

    If (cpu_desc_is_present(desc) = 0&) Or ((desc.access And &H10&) = 0&) Then
        cpu_segment_usable_at_cpl = 0&
        Exit Function
    End If

    rpl = (selector And &H3&)
    If cpu_desc_is_code(desc) <> 0& Then
        If cpu_desc_is_readable_code(desc) = 0& Then
            cpu_segment_usable_at_cpl = 0&
            Exit Function
        End If
        If (cpu_desc_is_conforming_code(desc) = 0&) And ((desc.dpl < target_cpl) Or (desc.dpl < rpl)) Then
            cpu_segment_usable_at_cpl = 0&
            Exit Function
        End If
        cpu_segment_usable_at_cpl = 1&
        Exit Function
    End If

    If (desc.dpl >= target_cpl) And (desc.dpl >= rpl) Then
        cpu_segment_usable_at_cpl = 1&
    Else
        cpu_segment_usable_at_cpl = 0&
    End If
End Function

Private Sub cpu_cleanup_outer_return_segments(ByRef cpu As CPU_t, ByVal target_cpl As Long)
    If cpu_segment_usable_at_cpl(cpu, cpu.segregs(CPU_REG_ES), target_cpl) = 0& Then cpu_null_seg cpu, CPU_REG_ES
    If cpu_segment_usable_at_cpl(cpu, cpu.segregs(CPU_REG_DS), target_cpl) = 0& Then cpu_null_seg cpu, CPU_REG_DS
    If cpu_segment_usable_at_cpl(cpu, cpu.segregs(CPU_REG_FS), target_cpl) = 0& Then cpu_null_seg cpu, CPU_REG_FS
    If cpu_segment_usable_at_cpl(cpu, cpu.segregs(CPU_REG_GS), target_cpl) = 0& Then cpu_null_seg cpu, CPU_REG_GS
End Sub

Private Sub cpu_retf(ByRef cpu As CPU_t, ByVal adjust As Long)
    Dim retcs As CPU_RETURNCS_t
    Dim new_ip As Long
    Dim new_esp As Long
    Dim stack_ptr As Long
    Dim new_cs As Long
    Dim new_ss As Long

    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        If cpu.isoper32 <> 0& Then
            cpu.ip = popl(cpu)
            putsegreg cpu, CPU_REG_CS, popl(cpu)
        Else
            cpu.ip = popw(cpu)
            putsegreg cpu, CPU_REG_CS, popw(cpu)
        End If

        If cpu.segis32(CPU_REG_SS) <> 0& Then
            cpu.regs_long(CPU_REG_ESP) = U32Add(cpu.regs_long(CPU_REG_ESP), adjust)
        Else
            putreg16 cpu, CPU_REG_ESP, (getreg16(cpu, CPU_REG_ESP) + (adjust And &HFFFF&))
        End If
        Exit Sub
    End If

    stack_ptr = cpu_stack_ptr_value(cpu)

    If cpu.isoper32 <> 0& Then
        new_ip = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), stack_ptr))
        new_cs = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu_stack_ptr_add(cpu, stack_ptr, 4&))) And &HFFFF&)
        If cpu_validate_return_cs(cpu, new_cs, retcs) = 0& Then Exit Sub

        If retcs.outer = 0& Then
            putsegreg cpu, CPU_REG_CS, new_cs
            cpu.ip = new_ip
            If cpu.segis32(CPU_REG_SS) <> 0& Then
                cpu.regs_long(CPU_REG_ESP) = U32Add(cpu.regs_long(CPU_REG_ESP), U32Add(8&, adjust))
            Else
                putreg16 cpu, CPU_REG_ESP, (getreg16(cpu, CPU_REG_ESP) + (8& + (adjust And &HFFFF&)))
            End If
            Exit Sub
        End If

        new_esp = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu_stack_ptr_add(cpu, stack_ptr, 8&)))
        new_ss = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu_stack_ptr_add(cpu, stack_ptr, 12&))) And &HFFFF&)
        If cpu_validate_return_ss(cpu, new_ss, retcs.target_cpl) = 0& Then Exit Sub

        putsegreg cpu, CPU_REG_SS, new_ss
        If cpu.segis32(CPU_REG_SS) <> 0& Then
            cpu.regs_long(CPU_REG_ESP) = U32Add(new_esp, adjust)
        Else
            putreg16 cpu, CPU_REG_ESP, (new_esp + (adjust And &HFFFF&))
        End If
        putsegreg cpu, CPU_REG_CS, new_cs
        cpu.ip = new_ip
        cpu.cpl = CByte(retcs.target_cpl And &H3&)
        cpu_cleanup_outer_return_segments cpu, retcs.target_cpl
    Else
        new_ip = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), stack_ptr)) And &HFFFF&)
        new_cs = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu_stack_ptr_add(cpu, stack_ptr, 2&))) And &HFFFF&)
        If cpu_validate_return_cs(cpu, new_cs, retcs) = 0& Then Exit Sub

        If retcs.outer = 0& Then
            putsegreg cpu, CPU_REG_CS, new_cs
            cpu.ip = new_ip
            If cpu.segis32(CPU_REG_SS) <> 0& Then
                cpu.regs_long(CPU_REG_ESP) = U32Add(cpu.regs_long(CPU_REG_ESP), U32Add(4&, adjust))
            Else
                putreg16 cpu, CPU_REG_ESP, (getreg16(cpu, CPU_REG_ESP) + (4& + (adjust And &HFFFF&)))
            End If
            Exit Sub
        End If

        new_esp = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu_stack_ptr_add(cpu, stack_ptr, 4&))) And &HFFFF&)
        new_ss = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu_stack_ptr_add(cpu, stack_ptr, 6&))) And &HFFFF&)
        If cpu_validate_return_ss(cpu, new_ss, retcs.target_cpl) = 0& Then Exit Sub

        putsegreg cpu, CPU_REG_SS, new_ss
        If cpu.segis32(CPU_REG_SS) <> 0& Then
            cpu.regs_long(CPU_REG_ESP) = U32Add(new_esp, adjust)
        Else
            putreg16 cpu, CPU_REG_ESP, (new_esp + (adjust And &HFFFF&))
        End If
        putsegreg cpu, CPU_REG_CS, new_cs
        cpu.ip = new_ip
        cpu.cpl = CByte(retcs.target_cpl And &H3&)
        cpu_cleanup_outer_return_segments cpu, retcs.target_cpl
    End If
End Sub

Private Sub cpu_iret(ByRef cpu As CPU_t)
    Dim current_cpl As Long
    Dim old_esp As Long
    Dim old_eflags As Long
    Dim merged_flags As Long
    Dim new_esp As Long
    Dim new_cs As Long
    Dim new_eip As Long
    Dim new_eflags As Long
    Dim new_ss As Long
    Dim new_es As Long
    Dim new_ds As Long
    Dim new_fs As Long
    Dim new_gs As Long
    Dim retcs As CPU_RETURNCS_t
    Dim sp16 As Long
    Dim backlink As Long

    current_cpl = cpu.cpl

    If cpu.protected_mode = 0& Then
        old_eflags = makeflagsword(cpu)
        If cpu.isoper32 <> 0& Then
            new_eip = popl(cpu)
            new_cs = popl(cpu)
            new_eflags = popl(cpu)
            putsegreg cpu, CPU_REG_CS, new_cs
            cpu.ip = new_eip
            cpu_restore_iret_flags cpu, new_eflags, current_cpl, current_cpl
        Else
            new_eip = popw(cpu)
            new_cs = popw(cpu)
            new_eflags = ((popw(cpu) And &HFFFF&) Or (old_eflags And &HFFFF0000))
            putsegreg cpu, CPU_REG_CS, new_cs
            cpu.ip = new_eip
            cpu_restore_iret_flags cpu, new_eflags, current_cpl, current_cpl
        End If
        Exit Sub
    End If

    If cpu.v86f <> 0& Then
        old_eflags = makeflagsword(cpu)
        If cpu.iopl < 3& Then
            cpu_raiseException cpu, 13&, 0&
            Exit Sub
        End If

        sp16 = getreg16(cpu, CPU_REG_ESP)

        If cpu.isoper32 <> 0& Then
            new_eip = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), (sp16 And &HFFFF&)))
            new_cs = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((sp16 + 4&) And &HFFFF&)))
            new_eflags = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((sp16 + 8&) And &HFFFF&)))
            putsegreg cpu, CPU_REG_CS, new_cs
            cpu.ip = new_eip
            putreg16 cpu, CPU_REG_ESP, (sp16 + 12&)
            cpu_restore_iret_flags cpu, new_eflags, current_cpl, 3&
            Exit Sub
        End If

        new_eip = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), (sp16 And &HFFFF&))) And &HFFFF&)
        new_cs = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((sp16 + 2&) And &HFFFF&))) And &HFFFF&)
        new_eflags = ((cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((sp16 + 4&) And &HFFFF&))) And &HFFFF&) Or (old_eflags And &HFFFF0000))

        putsegreg cpu, CPU_REG_CS, new_cs
        cpu.ip = new_eip
        putreg16 cpu, CPU_REG_ESP, (sp16 + 6&)
        cpu_restore_iret_flags cpu, new_eflags, current_cpl, 3&
        Exit Sub
    End If

    If cpu.nt <> 0& Then
        backlink = (cpu_readw(cpu, cpu.trbase) And &HFFFF&)
        cpu_task_switch cpu, backlink, TASK_SWITCH_REASON_IRET
        Exit Sub
    End If

    If cpu.isoper32 <> 0& Then
        new_eip = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), cpu.regs_long(CPU_REG_ESP)))
        new_cs = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 4&))) And &HFFFF&)
        new_eflags = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 8&)))
        old_eflags = makeflagsword(cpu)

        If (new_eflags And EFLAGS_VM) <> 0& Then
            If current_cpl <> 0& Then
                cpu_raiseException cpu, 13&, 0&
                Exit Sub
            End If

            new_esp = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 12&)))
            new_ss = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 16&)))
            new_es = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 20&)))
            new_ds = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 24&)))
            new_fs = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 28&)))
            new_gs = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 32&)))

            merged_flags = old_eflags
            merged_flags = ((merged_flags And Not (EFLAGS_CF Or EFLAGS_PF Or EFLAGS_AF Or EFLAGS_ZF Or EFLAGS_SF Or EFLAGS_TF Or EFLAGS_IF Or EFLAGS_DF Or EFLAGS_OF Or EFLAGS_IOPL Or EFLAGS_NT Or EFLAGS_VM Or EFLAGS_AC Or EFLAGS_ID)) Or _
                            (new_eflags And (EFLAGS_CF Or EFLAGS_PF Or EFLAGS_AF Or EFLAGS_ZF Or EFLAGS_SF Or EFLAGS_TF Or EFLAGS_IF Or EFLAGS_DF Or EFLAGS_OF Or EFLAGS_IOPL Or EFLAGS_NT Or EFLAGS_VM Or EFLAGS_AC Or EFLAGS_ID)))
            merged_flags = ((merged_flags And Not EFLAGS_RF) Or (old_eflags And EFLAGS_RF))

            If (merged_flags And EFLAGS_CF) <> 0& Then cpu.cf = 1& Else cpu.cf = 0&
            If (merged_flags And EFLAGS_PF) <> 0& Then cpu.pf = 1& Else cpu.pf = 0&
            If (merged_flags And EFLAGS_AF) <> 0& Then cpu.af = 1& Else cpu.af = 0&
            If (merged_flags And EFLAGS_ZF) <> 0& Then cpu.zf = 1& Else cpu.zf = 0&
            If (merged_flags And EFLAGS_SF) <> 0& Then cpu.sf = 1& Else cpu.sf = 0&
            If (merged_flags And EFLAGS_TF) <> 0& Then cpu.tf = 1& Else cpu.tf = 0&
            If (merged_flags And EFLAGS_IF) <> 0& Then cpu.ifl = 1& Else cpu.ifl = 0&
            If (merged_flags And EFLAGS_DF) <> 0& Then cpu.df = 1& Else cpu.df = 0&
            If (merged_flags And EFLAGS_OF) <> 0& Then cpu.ofl = 1& Else cpu.ofl = 0&
            cpu.iopl = CByte((U32Shr(merged_flags, 12&) And &H3&))
            If (merged_flags And EFLAGS_NT) <> 0& Then cpu.nt = 1& Else cpu.nt = 0&
            If (old_eflags And EFLAGS_RF) <> 0& Then cpu.rf = 1& Else cpu.rf = 0&
            cpu.v86f = 1&
            If (merged_flags And EFLAGS_AC) <> 0& Then cpu.acf = 1& Else cpu.acf = 0&
            If (merged_flags And EFLAGS_ID) <> 0& Then cpu.idf = 1& Else cpu.idf = 0&
            cpu.cpl = 3&

            putsegreg cpu, CPU_REG_SS, new_ss
            cpu.regs_long(CPU_REG_ESP) = new_esp
            putsegreg cpu, CPU_REG_ES, new_es
            putsegreg cpu, CPU_REG_DS, new_ds
            putsegreg cpu, CPU_REG_FS, new_fs
            putsegreg cpu, CPU_REG_GS, new_gs
            putsegreg cpu, CPU_REG_CS, new_cs
            cpu.ip = (new_eip And &HFFFF&)
            Exit Sub
        End If
    Else
        sp16 = getreg16(cpu, CPU_REG_ESP)
        new_eip = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), (sp16 And &HFFFF&))) And &HFFFF&)
        new_cs = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((sp16 + 2&) And &HFFFF&))) And &HFFFF&)
        new_eflags = ((cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((sp16 + 4&) And &HFFFF&))) And &HFFFF&) Or (makeflagsword(cpu) And &HFFFF0000))
    End If

    If cpu_validate_return_cs(cpu, new_cs, retcs) = 0& Then Exit Sub

    If retcs.outer = 0& Then
        putsegreg cpu, CPU_REG_CS, new_cs
        cpu.cpl = CByte(retcs.target_cpl And &H3&)
        cpu.ip = new_eip
        cpu_restore_iret_flags cpu, new_eflags, current_cpl, retcs.target_cpl
        If cpu.isoper32 <> 0& Then
            cpu.regs_long(CPU_REG_ESP) = U32Add(cpu.regs_long(CPU_REG_ESP), 12&)
        Else
            cpu.regs_long(CPU_REG_ESP) = U32Add(cpu.regs_long(CPU_REG_ESP), 6&)
        End If
        Exit Sub
    End If

    If cpu.isoper32 <> 0& Then
        new_esp = cpu_readl(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 12&)))
        new_ss = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), U32Add(cpu.regs_long(CPU_REG_ESP), 16&))) And &HFFFF&)
    Else
        new_esp = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((cpu.regs_long(CPU_REG_ESP) + 6&) And &HFFFF&))) And &HFFFF&)
        new_ss = (cpu_readw(cpu, U32Add(cpu.segcache(CPU_REG_SS), ((cpu.regs_long(CPU_REG_ESP) + 8&) And &HFFFF&))) And &HFFFF&)
    End If

    If cpu_validate_return_ss(cpu, new_ss, retcs.target_cpl) = 0& Then Exit Sub

    putsegreg cpu, CPU_REG_CS, new_cs
    cpu.ip = new_eip
    cpu_restore_iret_flags cpu, new_eflags, current_cpl, retcs.target_cpl
    putsegreg cpu, CPU_REG_SS, new_ss
    cpu.regs_long(CPU_REG_ESP) = new_esp
    cpu.cpl = CByte(retcs.target_cpl And &H3&)
    cpu_cleanup_outer_return_segments cpu, retcs.target_cpl
End Sub

Public Sub op_C2(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim imm As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    imm = (cpu_readw(cpu, codeAddr) And &HFFFF&)
    cpu.ip = pop(cpu)

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        cpu.regs_long(CPU_REG_ESP) = U32Add(cpu.regs_long(CPU_REG_ESP), imm)
    Else
        putreg16 cpu, CPU_REG_ESP, (getreg16(cpu, CPU_REG_ESP) + imm)
    End If
End Sub

Public Sub op_C3(ByRef cpu As CPU_t)
    cpu.ip = pop(cpu)
End Sub

Public Sub op_C8(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim nestIdx As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.stacksize = (cpu_readw(cpu, codeAddr) And &HFFFF&)
    cpu_stepIP cpu, 2&

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.nestlev = (cpu_read(cpu, codeAddr) And &H1F&)
    cpu_stepIP cpu, 1&

    If cpu.isoper32 <> 0& Then
        push cpu, cpu.regs_long(CPU_REG_EBP)
        cpu.frametemp32 = cpu.regs_long(CPU_REG_ESP)

        If cpu.nestlev <> 0& Then
            For nestIdx = 1& To (cpu.nestlev - 1&)
                cpu.regs_long(CPU_REG_EBP) = U32Sub(cpu.regs_long(CPU_REG_EBP), 4&)
                push cpu, cpu.regs_long(CPU_REG_EBP)
            Next nestIdx

            push cpu, cpu.frametemp32
        End If

        cpu.regs_long(CPU_REG_EBP) = cpu.frametemp32
        cpu.regs_long(CPU_REG_ESP) = U32Sub(cpu.regs_long(CPU_REG_EBP), (cpu.stacksize And &HFFFF&))
    Else
        push cpu, getreg16(cpu, CPU_REG_EBP)
        cpu.frametemp = getreg16(cpu, CPU_REG_ESP)

        If cpu.nestlev <> 0& Then
            For nestIdx = 1& To (cpu.nestlev - 1&)
                putreg16 cpu, CPU_REG_EBP, (getreg16(cpu, CPU_REG_EBP) - 2&)
                push cpu, getreg16(cpu, CPU_REG_EBP)
            Next nestIdx

            push cpu, cpu.frametemp
        End If

        putreg16 cpu, CPU_REG_EBP, cpu.frametemp
        putreg16 cpu, CPU_REG_ESP, (cpu.frametemp - (cpu.stacksize And &HFFFF&))
    End If
End Sub

Public Sub op_C9(ByRef cpu As CPU_t)
    If cpu.segis32(CPU_REG_SS) <> 0& Then
        cpu.regs_long(CPU_REG_ESP) = cpu.regs_long(CPU_REG_EBP)
    Else
        putreg16 cpu, CPU_REG_ESP, getreg16(cpu, CPU_REG_EBP)
    End If

    If cpu.isoper32 <> 0& Then
        cpu.regs_long(CPU_REG_EBP) = pop(cpu)
    Else
        putreg16 cpu, CPU_REG_EBP, pop(cpu)
    End If
End Sub

Public Sub op_CA(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.oper1 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
    cpu_retf cpu, cpu.oper1
End Sub

Public Sub op_CB(ByRef cpu As CPU_t)
    cpu_retf cpu, 0&
End Sub

Public Sub op_CE(ByRef cpu As CPU_t)
    If cpu.ofl <> 0& Then
        cpu_intcall cpu, 4&, INT_SOURCE_INTO, 0&
    End If
End Sub

Public Sub op_CF(ByRef cpu As CPU_t)
    cpu_iret cpu
End Sub

Public Sub op_F4(ByRef cpu As CPU_t)
    If (cpu.v86f <> 0&) Or ((cpu.protected_mode <> 0&) And (cpu.cpl <> 0&)) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    cpu.hltstate = 1&
End Sub

Public Sub op_F6(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    op_grp3_8 cpu

    If (cpu.reg > 1&) And (cpu.reg < 4&) Then
        writerm8 cpu, cpu.rm, cpu.res8
    End If
End Sub

Public Sub op_F7(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        op_grp3_32 cpu

        If (cpu.reg > 1&) And (cpu.reg < 4&) Then
            writerm32 cpu, cpu.rm, cpu.res32
        End If
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        op_grp3_16 cpu

        If (cpu.reg > 1&) And (cpu.reg < 4&) Then
            writerm16 cpu, cpu.rm, cpu.res16
        End If
    End If
End Sub

Public Sub op_F1(ByRef cpu As CPU_t)
    cpu_intcall cpu, 1&, INT_SOURCE_SOFTWARE, 0&
End Sub

Public Sub op_F5(ByRef cpu As CPU_t)
    cpu.cf = CByte((cpu.cf Xor 1&) And &H1&)
End Sub

Public Sub op_F8(ByRef cpu As CPU_t)
    cpu.cf = 0&
End Sub

Public Sub op_F9(ByRef cpu As CPU_t)
    cpu.cf = 1&
End Sub

Public Sub op_FA(ByRef cpu As CPU_t)
    If cpu.protected_mode <> 0& Then
        If cpu.iopl >= cpu.cpl Then
            cpu.ifl = 0&
        Else
            cpu_raiseException cpu, 13&, 0&
        End If
    Else
        cpu.ifl = 0&
    End If
End Sub

Public Sub op_FB(ByRef cpu As CPU_t)
    If cpu.protected_mode <> 0& Then
        If cpu.iopl >= cpu.cpl Then
            cpu.ifl = 1&
            cpu_begin_interrupt_shadow cpu
        Else
            cpu_raiseException cpu, 13&, 0&
        End If
    Else
        cpu.ifl = 1&
        cpu_begin_interrupt_shadow cpu
    End If
End Sub

Public Sub op_FC(ByRef cpu As CPU_t)
    cpu.df = 0&
End Sub

Public Sub op_FD(ByRef cpu As CPU_t)
    cpu.df = 1&
End Sub

Public Sub op_FE(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = 1&

    If (cpu.reg And &H7&) = 0& Then
        cpu.tempcf = cpu.cf
        cpu.res8 = ((cpu.oper1b + cpu.oper2b) And &HFF&)
        flag_add8 cpu, cpu.oper1b, cpu.oper2b
        cpu.cf = cpu.tempcf
        writerm8 cpu, cpu.rm, cpu.res8
    Else
        cpu.tempcf = cpu.cf
        cpu.res8 = ((cpu.oper1b - cpu.oper2b) And &HFF&)
        flag_sub8 cpu, cpu.oper1b, cpu.oper2b
        cpu.cf = cpu.tempcf
        writerm8 cpu, cpu.rm, cpu.res8
    End If
End Sub

Public Sub cpu_callf(ByRef cpu As CPU_t, ByVal selector As Long, ByVal newIp As Long)
    Dim desc As CPU_SEGDESC_t
    Dim gate As CPU_GATEDESC_t
    Dim target As CPU_CODETARGET_t
    Dim new_esp As Long
    Dim old_esp As Long
    Dim old_ss_base As Long
    Dim new_ss As Long
    Dim old_ss As Long
    Dim task_selector As Long
    Dim old_stack_is32 As Long
    Dim epl As Long
    Dim gate32 As Long

    selector = (selector And &HFFFF&)

    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        If cpu.isoper32 <> 0& Then
            pushl cpu, cpu.segregs(CPU_REG_CS)
            pushl cpu, cpu.ip
            cpu.ip = newIp
        Else
            pushw cpu, cpu.segregs(CPU_REG_CS)
            pushw cpu, cpu.ip
            cpu.ip = (newIp And &HFFFF&)
        End If
        putsegreg cpu, CPU_REG_CS, selector
        Exit Sub
    End If

    old_ss = (cpu.segregs(CPU_REG_SS) And &HFFFF&)
    old_esp = cpu.regs_long(CPU_REG_ESP)
    old_ss_base = cpu.segcache(CPU_REG_SS)
    old_stack_is32 = cpu.segis32(CPU_REG_SS)

    If cpu_selector_offset(selector) = 0& Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If cpu_read_segdesc(cpu, selector, desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        Exit Sub
    End If

    If cpu_desc_is_code(desc) <> 0& Then
        If cpu_validate_direct_call_target(cpu, selector, target) = 0& Then Exit Sub
        If cpu.isoper32 <> 0& Then
            cpu_push_far_return32 cpu, cpu.segregs(CPU_REG_CS), cpu.ip
        Else
            cpu_push_far_return16 cpu, cpu.segregs(CPU_REG_CS), cpu.ip
        End If
        putsegreg cpu, CPU_REG_CS, target.selector
        If cpu.isoper32 <> 0& Then
            cpu.ip = newIp
        Else
            cpu.ip = (newIp And &HFFFF&)
        End If
        Exit Sub
    End If

    If (cpu_desc_is_task_gate(desc) <> 0&) Or (cpu_desc_is_tss(desc) <> 0&) Then
        If cpu_resolve_task_switch_target(cpu, selector, desc, task_selector) = 0& Then Exit Sub
        cpu_task_switch cpu, task_selector, TASK_SWITCH_REASON_CALL
        Exit Sub
    End If

    If cpu_desc_is_call_gate(desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        Exit Sub
    End If

    cpu_load_gate_desc cpu, desc.addr, desc.access, desc.flags, gate
    If (selector And &H3&) > cpu.cpl Then
        epl = (selector And &H3&)
    Else
        epl = cpu.cpl
    End If
    If epl > gate.dpl Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        Exit Sub
    End If
    If gate.present = 0& Then
        cpu_raiseException cpu, 11&, cpu_selector_error_code(selector)
        Exit Sub
    End If
    If cpu_validate_gate_target_code(cpu, gate.target_selector, target) = 0& Then Exit Sub

    gate32 = cpu_gate_is_32bit(gate)

    If target.outer = 0& Then
        If gate32 <> 0& Then
            cpu_push_far_return32 cpu, cpu.segregs(CPU_REG_CS), cpu.ip
            putsegreg cpu, CPU_REG_CS, target.selector
            cpu.ip = gate.offset
        Else
            cpu_push_far_return16 cpu, cpu.segregs(CPU_REG_CS), cpu.ip
            putsegreg cpu, CPU_REG_CS, target.selector
            cpu.ip = (gate.offset And &HFFFF&)
        End If
        cpu.cpl = CByte(target.target_cpl And &H3&)
        Exit Sub
    End If

    If cpu_fetch_tss_stack(cpu, target.target_cpl, new_ss, new_esp) = 0& Then Exit Sub

    putsegreg cpu, CPU_REG_SS, new_ss
    cpu.regs_long(CPU_REG_ESP) = new_esp
    cpu_copy_call_gate_params_sys cpu, old_ss_base, old_esp, old_stack_is32, gate.param_count, gate32

    If gate32 <> 0& Then
        cpu_stack_pushl_sys cpu, old_ss
        cpu_stack_pushl_sys cpu, old_esp
        cpu_push_far_return32_sys cpu, cpu.segregs(CPU_REG_CS), cpu.ip
        putsegreg cpu, CPU_REG_CS, target.selector
        cpu.ip = gate.offset
    Else
        cpu_stack_pushw_sys cpu, old_ss
        cpu_stack_pushw_sys cpu, (old_esp And &HFFFF&)
        cpu_push_far_return16_sys cpu, cpu.segregs(CPU_REG_CS), cpu.ip
        putsegreg cpu, CPU_REG_CS, target.selector
        cpu.ip = (gate.offset And &HFFFF&)
    End If
    cpu.cpl = CByte(target.target_cpl And &H3&)
End Sub

Private Sub cpu_jmpf(ByRef cpu As CPU_t, ByVal selector As Long, ByVal newIp As Long)
    Dim desc As CPU_SEGDESC_t
    Dim gate As CPU_GATEDESC_t
    Dim target As CPU_CODETARGET_t
    Dim task_selector As Long
    Dim epl As Long

    selector = (selector And &HFFFF&)

    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        If cpu.isoper32 <> 0& Then
            cpu.ip = newIp
        Else
            cpu.ip = (newIp And &HFFFF&)
        End If
        putsegreg cpu, CPU_REG_CS, selector
        Exit Sub
    End If

    If cpu_selector_offset(selector) = 0& Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If cpu_read_segdesc(cpu, selector, desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        Exit Sub
    End If

    If cpu_desc_is_code(desc) <> 0& Then
        If cpu_validate_direct_call_target(cpu, selector, target) = 0& Then Exit Sub
        putsegreg cpu, CPU_REG_CS, target.selector
        If cpu.isoper32 <> 0& Then
            cpu.ip = newIp
        Else
            cpu.ip = (newIp And &HFFFF&)
        End If
        Exit Sub
    End If

    If (cpu_desc_is_task_gate(desc) <> 0&) Or (cpu_desc_is_tss(desc) <> 0&) Then
        If cpu_resolve_task_switch_target(cpu, selector, desc, task_selector) = 0& Then Exit Sub
        cpu_task_switch cpu, task_selector, TASK_SWITCH_REASON_JMP
        Exit Sub
    End If

    If cpu_desc_is_call_gate(desc) = 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        Exit Sub
    End If

    cpu_load_gate_desc cpu, desc.addr, desc.access, desc.flags, gate
    If (selector And &H3&) > cpu.cpl Then
        epl = (selector And &H3&)
    Else
        epl = cpu.cpl
    End If
    If epl > gate.dpl Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(selector)
        Exit Sub
    End If
    If gate.present = 0& Then
        cpu_raiseException cpu, 11&, cpu_selector_error_code(selector)
        Exit Sub
    End If
    If cpu_validate_gate_target_code(cpu, gate.target_selector, target) = 0& Then Exit Sub
    If target.outer <> 0& Then
        cpu_raiseException cpu, 13&, cpu_selector_error_code(gate.target_selector)
        Exit Sub
    End If

    putsegreg cpu, CPU_REG_CS, target.selector
    If cpu_gate_is_32bit(gate) <> 0& Then
        cpu.ip = gate.offset
    Else
        cpu.ip = (gate.offset And &HFFFF&)
    End If
    cpu.cpl = CByte(target.target_cpl And &H3&)
End Sub

Public Sub op_grp5(ByRef cpu As CPU_t)
    Dim new_ip As Long
    Dim new_cs As Long

    Select Case (cpu.reg And &H7&)
        Case 0&  ' INC Ev
            cpu.oper2 = 1&
            cpu.tempcf = cpu.cf
            op_add16 cpu
            cpu.cf = cpu.tempcf
            writerm16 cpu, cpu.rm, cpu.res16

        Case 1&  ' DEC Ev
            cpu.oper2 = 1&
            cpu.tempcf = cpu.cf
            op_sub16 cpu
            cpu.cf = cpu.tempcf
            writerm16 cpu, cpu.rm, cpu.res16

        Case 2&  ' CALL Ev
            push cpu, cpu.ip
            cpu.ip = (cpu.oper1 And &HFFFF&)

        Case 3&  ' CALL Mp
            If cpu.mode = 3& Then
                cpu.ip = cpu_firstip
                cpu_raiseException cpu, 6&, 0&
                Exit Sub
            End If
            getea cpu, cpu.rm
            new_ip = (cpu_readw(cpu, cpu.ea) And &HFFFF&)
            new_cs = (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&)
            cpu_callf cpu, new_cs, new_ip

        Case 4&  ' JMP Ev
            cpu.ip = (cpu.oper1 And &HFFFF&)

        Case 5&  ' JMP Mp
            If cpu.mode = 3& Then
                cpu.ip = cpu_firstip
                cpu_raiseException cpu, 6&, 0&
                Exit Sub
            End If
            getea cpu, cpu.rm
            cpu_jmpf cpu, (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&), (cpu_readw(cpu, cpu.ea) And &HFFFF&)

        Case 6&  ' PUSH Ev
            push cpu, cpu.oper1
    End Select
End Sub

Public Sub op_grp5_32(ByRef cpu As CPU_t)
    Dim new_ip As Long
    Dim new_cs As Long

    Select Case (cpu.reg And &H7&)
        Case 0&  ' INC Ev
            cpu.oper2_32 = 1&
            cpu.tempcf = cpu.cf
            op_add32 cpu
            cpu.cf = cpu.tempcf
            writerm32 cpu, cpu.rm, cpu.res32

        Case 1&  ' DEC Ev
            cpu.oper2_32 = 1&
            cpu.tempcf = cpu.cf
            op_sub32 cpu
            cpu.cf = cpu.tempcf
            writerm32 cpu, cpu.rm, cpu.res32

        Case 2&  ' CALL Ev
            push cpu, cpu.ip
            cpu.ip = cpu.oper1_32

        Case 3&  ' CALL Mp
            If cpu.mode = 3& Then
                cpu.ip = cpu_firstip
                cpu_raiseException cpu, 6&, 0&
                Exit Sub
            End If
            getea cpu, cpu.rm
            new_ip = cpu_readl(cpu, cpu.ea)
            new_cs = (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&)
            cpu_callf cpu, new_cs, new_ip

        Case 4&  ' JMP Ev
            cpu.ip = cpu.oper1_32

        Case 5&  ' JMP Mp
            If cpu.mode = 3& Then
                cpu.ip = cpu_firstip
                cpu_raiseException cpu, 6&, 0&
                Exit Sub
            End If
            getea cpu, cpu.rm
            cpu_jmpf cpu, (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&), cpu_readl(cpu, cpu.ea)

        Case 6&  ' PUSH Ev
            push cpu, cpu.oper1_32
    End Select
End Sub

Public Sub op_FF(ByRef cpu As CPU_t)
    If cpu.isoper32 <> 0& Then
        modregrm cpu
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        op_grp5_32 cpu
    Else
        modregrm cpu
        cpu.oper1 = readrm16(cpu, cpu.rm)
        op_grp5 cpu
    End If
End Sub

Public Sub op_ext_00(ByRef cpu As CPU_t)
    Dim accessByte As Long
    Dim dpl As Long
    Dim present As Long
    Dim typeVal As Long
    Dim cplVal As Long
    Dim newType As Long

    modregrm cpu

    Select Case (cpu.reg And &H7&)
        Case 0&  ' SLDT
            writerm16 cpu, cpu.rm, cpu.ldt_selector

        Case 1&  ' STR
            writerm16 cpu, cpu.rm, cpu.tr_selector

        Case 2&  ' LLDT
            cpu.temp16 = readrm16(cpu, cpu.rm)
            cpu.ldt_selector = (cpu.temp16 And &HFFFF&)
            cpu.tempaddr32 = U32Add(cpu.gdtr, (cpu.temp16 And &HFFF8&))

            cpu.ldtl = ((cpu_readw_sys(cpu, cpu.tempaddr32) And &HFFFF&) Or U32Shl((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 6&)) And &HF&), 16&))
            If (cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 6&)) And &H80&) <> 0& Then
                cpu.ldtl = U32Shl(cpu.ldtl, 12&)
                cpu.ldtl = (cpu.ldtl Or &HFFF&)
            End If

            cpu.ldtr = ((cpu_readw_sys(cpu, U32Add(cpu.tempaddr32, 2&)) And &HFFFF&) _
                Or U32Shl((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 4&)) And &HFF&), 16&) _
                Or U32Shl(U32Shr((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 7&)) And &HFF&), 4&), 24&))

        Case 3&  ' LTR
            cpu.temp16 = readrm16(cpu, cpu.rm)
            cpu.tr_selector = (cpu.temp16 And &HFFFF&)
            cpu.tempaddr32 = U32Add(cpu.gdtr, U32Shl(U32Shr(cpu.temp16, 3&), 3&))

            accessByte = (cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 5&)) And &HFF&)
            typeVal = (accessByte And &HF&)
            dpl = (U32Shr(accessByte, 5&) And &H3&)
            cplVal = (cpu.segregs(CPU_REG_CS) And &H3&)
            present = (U32Shr(accessByte, 7&) And &H1&)

            If (typeVal <> &H9&) And (typeVal <> &HB&) And (typeVal <> &H3&) And (typeVal <> &H7&) Then
                ' Keep C behavior: no exception yet for unsupported type.
            ElseIf present = 0& Then
                cpu_raiseException cpu, 11&, cpu.temp16
                Exit Sub
            ElseIf cplVal > dpl Then
                cpu_raiseException cpu, 13&, cpu.temp16
                Exit Sub
            Else
                cpu.trtype = CByte(typeVal And &HFF&)
                cpu.trlimit = U32Add((cpu_readw_sys(cpu, cpu.tempaddr32) And &HFFFF&), 1&)
                cpu.trbase = ((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 2&)) And &HFF&) _
                    Or U32Shl((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 3&)) And &HFF&), 8&) _
                    Or U32Shl((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 4&)) And &HFF&), 16&) _
                    Or U32Shl((cpu_read_sys(cpu, U32Add(cpu.tempaddr32, 7&)) And &HFF&), 24&))

                If typeVal = &H9& Then
                    newType = &HB&
                Else
                    newType = &H7&
                End If
                cpu_write_sys cpu, U32Add(cpu.tempaddr32, 5&), ((accessByte And &HF0&) Or newType)
            End If

        Case 4&  ' VERR
            debug_log DEBUG_DETAIL, "VERR" & vbCrLf
            cpu.zf = 1&

        Case 5&  ' VERW
            debug_log DEBUG_DETAIL, "VERW" & vbCrLf
            cpu.zf = 1&

        Case Else
            cpu_raiseException cpu, 6&, 0&
    End Select
End Sub

Public Sub op_ext_01(ByRef cpu As CPU_t)
    Dim oldpm As Long
    Dim smswVal As Long

    modregrm cpu

    Select Case (cpu.reg And &H7&)
        Case 0&  ' SGDT
            getea cpu, cpu.rm
            cpu_writew cpu, cpu.ea, cpu.gdtl
            If cpu.isoper32 <> 0& Then
                cpu_writel cpu, U32Add(cpu.ea, 2&), cpu.gdtr
            Else
                cpu_writew cpu, U32Add(cpu.ea, 2&), (cpu.gdtr And &HFFFF&)
                cpu_write cpu, U32Add(cpu.ea, 4&), U32Shr(cpu.gdtr, 16&)
            End If

        Case 1&  ' SIDT
            getea cpu, cpu.rm
            cpu_writew cpu, cpu.ea, cpu.idtl
            If cpu.isoper32 <> 0& Then
                cpu_writel cpu, U32Add(cpu.ea, 2&), cpu.idtr
            Else
                cpu_writew cpu, U32Add(cpu.ea, 2&), (cpu.idtr And &HFFFF&)
                cpu_write cpu, U32Add(cpu.ea, 4&), U32Shr(cpu.idtr, 16&)
            End If

        Case 2&  ' LGDT
            If (cpu.protected_mode <> 0&) And (cpu.cpl > 0&) Then
                debug_log DEBUG_INFO, "Attempted to use LGDT when CPU was already in protected mode and CPL>0!"
                cpu_raiseException cpu, 13&, 0&
                Exit Sub
            End If

            getea cpu, cpu.rm
            cpu.gdtl = (cpu_readw(cpu, cpu.ea) And &HFFFF&)
            If cpu.isoper32 <> 0& Then
                cpu.gdtr = cpu_readl(cpu, U32Add(cpu.ea, 2&))
            Else
                cpu.gdtr = ((cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&) Or U32Shl((cpu_read(cpu, U32Add(cpu.ea, 4&)) And &HFF&), 16&))
            End If

        Case 3&  ' LIDT
            getea cpu, cpu.rm
            cpu.idtl = (cpu_readw(cpu, cpu.ea) And &HFFFF&)
            If cpu.isoper32 <> 0& Then
                cpu.idtr = cpu_readl(cpu, U32Add(cpu.ea, 2&))
            Else
                cpu.idtr = ((cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&) Or U32Shl((cpu_read(cpu, U32Add(cpu.ea, 4&)) And &HFF&), 16&))
            End If

        Case 4&  ' SMSW
            smswVal = (cpu.CR(0&) And &HFFFF&)
            If cpu.have387 = 0& Then
                smswVal = (smswVal Or &H4&)
            End If
            writerm16 cpu, cpu.rm, smswVal

        Case 6&  ' LMSW
            oldpm = (cpu.CR(0&) And &H1&)
            cpu.CR(0&) = ((cpu.CR(0&) And &HFFFFFFE1) Or (readrm16(cpu, cpu.rm) And &H1E&))
            cpu.CR(0&) = (cpu.CR(0&) Or (readrm16(cpu, cpu.rm) And &H11&))

            If ((cpu.CR(0&) And &H1&) <> 0&) And (oldpm = 0&) Then
                cpu.protected_mode = 1&
                cpu.ifl = 0&
            End If

        Case 7&  ' INVLPG
            If cpu.cpl > 0& Then
                cpu_raiseException cpu, 13&, 0&
            ElseIf cpu.mode = 3& Then
                cpu_raiseException cpu, 6&, 0&
            Else
                getea cpu, cpu.rm
                If memory_paging_enabled(cpu) <> 0& Then
                    memory_tlb_invalidate_page cpu, cpu.ea
                End If
            End If
    End Select
End Sub

Public Sub op_ext_02(ByRef cpu As CPU_t)
    Dim desc As CPU_SEGDESC_t
    Dim selector As Long
    Dim epl As Long
    Dim larVal As Long
    Dim typeVal As Long

    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    modregrm cpu
    selector = (readrm16(cpu, cpu.rm) And &HFFFF&)
    If cpu.doexception <> 0& Then Exit Sub

    cpu.zf = 0&
    If cpu_selector_offset(selector) = 0& Then Exit Sub
    If cpu_read_segdesc(cpu, selector, desc) = 0& Then Exit Sub

    If cpu_desc_is_system(desc) <> 0& Then
        typeVal = cpu_desc_type(desc)
        Select Case typeVal
            Case &H1&, &H2&, &H3&, &H4&, &H5&, &H6&, &H7&, &H9&, &HB&, &HC&, &HE&, &HF&
                ' Allowed system descriptor types for LAR
            Case Else
                Exit Sub
        End Select
    End If

    epl = (selector And &H3&)
    If epl < cpu.cpl Then epl = cpu.cpl
    If ((cpu_desc_is_code(desc) = 0&) Or (cpu_desc_is_conforming_code(desc) = 0&)) And (desc.dpl < epl) Then Exit Sub

    larVal = (U32Shl((desc.access And &HFF&), 8&) Or U32Shl((desc.flags And &HF0&), 16&))

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, larVal
    Else
        putreg16 cpu, cpu.reg, larVal
    End If
    cpu.zf = 1&
End Sub

Public Sub op_ext_03(ByRef cpu As CPU_t)
    Dim desc As CPU_SEGDESC_t
    Dim selector As Long
    Dim epl As Long
    Dim limitVal As Long
    Dim typeVal As Long

    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    modregrm cpu
    selector = (readrm16(cpu, cpu.rm) And &HFFFF&)
    If cpu.doexception <> 0& Then Exit Sub

    cpu.zf = 0&
    If cpu_selector_offset(selector) = 0& Then Exit Sub
    If cpu_read_segdesc(cpu, selector, desc) = 0& Then Exit Sub

    If cpu_desc_is_system(desc) <> 0& Then
        typeVal = cpu_desc_type(desc)
        Select Case typeVal
            Case &H1&, &H2&, &H3&, &H9&, &HB&
                ' Allowed system descriptor types for LSL
            Case Else
                Exit Sub
        End Select
    End If

    epl = (selector And &H3&)
    If epl < cpu.cpl Then epl = cpu.cpl
    If ((cpu_desc_is_code(desc) = 0&) Or (cpu_desc_is_conforming_code(desc) = 0&)) And (desc.dpl < epl) Then Exit Sub

    limitVal = ((cpu_readw_sys(cpu, desc.addr) And &HFFFF&) Or U32Shl((desc.flags And &HF&), 16&))
    If (desc.flags And &H80&) <> 0& Then
        limitVal = U32Shl(limitVal, 12&)
        limitVal = (limitVal Or &HFFF&)
    End If

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, limitVal
    Else
        putreg16 cpu, cpu.reg, limitVal
    End If
    cpu.zf = 1&
End Sub

Public Sub op_ext_06(ByRef cpu As CPU_t)
    If cpu.cpl > 0& Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    cpu.CR(0&) = (cpu.CR(0&) And Not &H8&)
End Sub

Public Sub op_ext_08_09(ByRef cpu As CPU_t)
    ' INVD/WBINVD currently no-op in C core.
End Sub
Public Sub op_ext_20(ByRef cpu As CPU_t)
    modregrm cpu

    If (cpu.v86f <> 0&) Or (cpu.cpl > 0&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If (cpu.mode <> 3&) Or ((cpu.reg <> 0&) And (cpu.reg <> 2&) And (cpu.reg <> 3&) And (cpu.reg <> 4&)) Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    cpu.regs_long(cpu.rm) = cpu.CR(cpu.reg)
End Sub

Public Sub op_ext_21(ByRef cpu As CPU_t)
    modregrm cpu

    If (cpu.v86f <> 0&) Or (cpu.cpl > 0&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If cpu.mode <> 3& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    cpu.regs_long(cpu.rm) = cpu.dr(cpu.reg)
End Sub

Public Sub op_ext_22(ByRef cpu As CPU_t)
    Dim i As Long
    Dim new_cr As Long
    Dim old_paging_enabled As Long

    modregrm cpu

    If (cpu.v86f <> 0&) Or (cpu.cpl > 0&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If (cpu.mode <> 3&) Or ((cpu.reg <> 0&) And (cpu.reg <> 2&) And (cpu.reg <> 3&) And (cpu.reg <> 4&)) Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    Select Case cpu.reg
        Case 0&
            old_paging_enabled = memory_paging_enabled(cpu)
            new_cr = cpu.regs_long(cpu.rm)
            If ((new_cr And &H80000000) <> 0&) And ((new_cr And &H1&) = 0&) Then
                cpu_raiseException cpu, 13&, 0&
                Exit Sub
            End If

            cpu.CR(0&) = new_cr
            If (cpu.CR(0&) And &H1&) <> 0& Then
                cpu.protected_mode = 1&
                If memory_paging_enabled(cpu) <> 0& Then
                    cpu.paging = 1&
                Else
                    cpu.paging = 0&
                End If
            Else
                cpu.protected_mode = 0&
                cpu.paging = 0&
                cpu.usegdt = 0&
                cpu.isoper32 = 0&
                cpu.isaddr32 = 0&
                cpu.isCS32 = 0&
                For i = 0& To 5&
                    cpu.segis32(i) = 0&
                Next i
            End If

            If old_paging_enabled <> memory_paging_enabled(cpu) Then
                memory_tlb_flush cpu
            End If

        Case 3&
            cpu.CR(3&) = cpu.regs_long(cpu.rm)
            memory_tlb_flush cpu

        Case 4&
            cpu.CR(4&) = cpu.regs_long(cpu.rm)
            debug_log DEBUG_DETAIL, "CR4: " & Hex$(cpu.CR(4&)) & vbCrLf

        Case Else
            cpu.CR(cpu.reg) = cpu.regs_long(cpu.rm)
    End Select
End Sub

Public Sub op_ext_23(ByRef cpu As CPU_t)
    modregrm cpu

    If (cpu.v86f <> 0&) Or (cpu.cpl > 0&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If cpu.mode <> 3& Then
        cpu.ip = cpu_firstip
        cpu_raiseException cpu, 6&, 0&
        Exit Sub
    End If

    cpu.dr(cpu.reg) = cpu.regs_long(cpu.rm)
End Sub

Public Sub op_ext_24_26(ByRef cpu As CPU_t)
    modregrm cpu
End Sub

Public Sub op_ext_30(ByRef cpu As CPU_t)
    ' WRMSR is currently stubbed in C core.
End Sub

Public Sub op_ext_31(ByRef cpu As CPU_t)
    cpu.regs_long(CPU_REG_EDX) = cpu.totalexec_hi
    cpu.regs_long(CPU_REG_EAX) = cpu.totalexec_lo
End Sub
Public Sub op_ext_40(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.ofl <> 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_41(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.ofl = 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_42(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.cf <> 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_43(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.cf = 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_44(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.zf <> 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_45(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.zf = 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_46(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.cf <> 0&) Or (cpu.zf <> 0&) Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_47(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.cf = 0&) And (cpu.zf = 0&) Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub
Public Sub op_ext_48(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf <> 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_49(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf = 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_4A(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.pf <> 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_4B(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.pf = 0& Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_4C(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf <> cpu.ofl Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_4D(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf = cpu.ofl Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_4E(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.zf <> 0&) Or (cpu.sf <> cpu.ofl) Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub

Public Sub op_ext_4F(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.zf = 0&) And (cpu.sf = cpu.ofl) Then
        If cpu.isoper32 <> 0& Then
            putreg32 cpu, cpu.reg, readrm32(cpu, cpu.rm)
        Else
            putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
        End If
    End If
End Sub
Public Sub op_ext_80(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.ofl <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_81(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.ofl = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_82(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.cf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_83(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.cf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_84(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.zf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_85(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.zf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_86(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If (cpu.cf <> 0&) Or (cpu.zf <> 0&) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_87(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If (cpu.cf = 0&) And (cpu.zf = 0&) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_88(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.sf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_89(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.sf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_8A(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.pf <> 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_8B(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.pf = 0& Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_8C(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.sf <> cpu.ofl Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_8D(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If cpu.sf = cpu.ofl Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_8E(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If (cpu.sf <> cpu.ofl) Or (cpu.zf <> 0&) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_8F(ByRef cpu As CPU_t)
    Dim codeAddr As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    If cpu.isoper32 <> 0& Then
        cpu.temp32 = cpu_readl(cpu, codeAddr)
        cpu_stepIP cpu, 4&
    Else
        cpu.temp32 = cpu_signext16to32(cpu_readw(cpu, codeAddr))
        cpu_stepIP cpu, 2&
    End If

    If (cpu.zf = 0&) And (cpu.sf = cpu.ofl) Then
        cpu_apply_relative_branch cpu, cpu.temp32
    End If
End Sub

Public Sub op_ext_90(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.ofl <> 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_91(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.ofl = 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_92(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.cf <> 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_93(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.cf = 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_94(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.zf <> 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_95(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.zf = 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_96(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.cf <> 0&) Or (cpu.zf <> 0&) Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_97(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.cf = 0&) And (cpu.zf = 0&) Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_98(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf <> 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_99(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf = 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_9A(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.pf <> 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_9B(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.pf = 0& Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_9C(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf <> cpu.ofl Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_9D(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.sf = cpu.ofl Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_9E(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.sf <> cpu.ofl) Or (cpu.zf <> 0&) Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_9F(ByRef cpu As CPU_t)
    modregrm cpu
    If (cpu.zf = 0&) And (cpu.sf = cpu.ofl) Then
        writerm8 cpu, cpu.rm, 1&
    Else
        writerm8 cpu, cpu.rm, 0&
    End If
End Sub

Public Sub op_ext_A0(ByRef cpu As CPU_t)
    push cpu, getsegreg(cpu, CPU_REG_FS)
End Sub

Public Sub op_ext_A1(ByRef cpu As CPU_t)
    putsegreg cpu, CPU_REG_FS, pop(cpu)
End Sub

Public Sub op_ext_A2(ByRef cpu As CPU_t)
    Select Case cpu.regs_long(CPU_REG_EAX)
        Case 0&
            cpu.regs_long(CPU_REG_EAX) = 1&
            cpu.regs_long(CPU_REG_EBX) = &H756E6547
            cpu.regs_long(CPU_REG_EDX) = &H49656E69
            cpu.regs_long(CPU_REG_ECX) = &H6C65746E

        Case 1&
            cpu.regs_long(CPU_REG_EAX) = &H400&
            cpu.regs_long(CPU_REG_EBX) = 0&
            cpu.regs_long(CPU_REG_ECX) = 0&
            cpu.regs_long(CPU_REG_EDX) = ((cpu.have387 And &H1&) Or &H10& Or &H8000&)

        Case Else
            cpu.regs_long(CPU_REG_EAX) = 0&
            cpu.regs_long(CPU_REG_EBX) = 0&
            cpu.regs_long(CPU_REG_ECX) = 0&
            cpu.regs_long(CPU_REG_EDX) = 0&
    End Select
End Sub

Private Function cpu_floor_div_pow2_s32(ByVal value As Long, ByVal shift As Long) As Long
    Dim absVal As Long
    Dim q As Long

    If value >= 0& Then
        If shift = 5& Then
            cpu_floor_div_pow2_s32 = (value \ 32&)
        Else
            cpu_floor_div_pow2_s32 = (value \ 16&)
        End If
        Exit Function
    End If

    absVal = U32Add((Not value), 1&)
    If shift = 5& Then
        q = U32Shr(U32Add(absVal, 31&), 5&)
    Else
        q = U32Shr(U32Add(absVal, 15&), 4&)
    End If

    cpu_floor_div_pow2_s32 = -q
End Function

Private Function cpu_bitstring_ea_adjust(ByVal bitOffset As Long, ByVal operandBits As Long) As Long
    If operandBits = 32& Then
        cpu_bitstring_ea_adjust = U32Shl(cpu_floor_div_pow2_s32(bitOffset, 5&), 2&)
    Else
        cpu_bitstring_ea_adjust = U32Shl(cpu_floor_div_pow2_s32(bitOffset, 4&), 1&)
    End If
End Function

Private Function cpu_bit_access_read32(ByRef cpu As CPU_t, ByVal bitOffset As Long, ByVal adjustMemory As Long) As Long
    If cpu.mode = 3& Then
        cpu_bit_access_read32 = readrm32(cpu, cpu.rm)
        Exit Function
    End If

    getea cpu, cpu.rm
    If adjustMemory <> 0& Then
        cpu.ea = U32Add(cpu.ea, cpu_bitstring_ea_adjust(bitOffset, 32&))
    End If

    cpu_bit_access_read32 = cpu_readl(cpu, cpu.ea)
End Function

Private Function cpu_bit_access_read16(ByRef cpu As CPU_t, ByVal bitOffset As Long, ByVal adjustMemory As Long) As Long
    If cpu.mode = 3& Then
        cpu_bit_access_read16 = (readrm16(cpu, cpu.rm) And &HFFFF&)
        Exit Function
    End If

    getea cpu, cpu.rm
    If adjustMemory <> 0& Then
        cpu.ea = U32Add(cpu.ea, cpu_bitstring_ea_adjust(bitOffset, 16&))
    End If

    cpu_bit_access_read16 = (cpu_readw(cpu, cpu.ea) And &HFFFF&)
End Function

Private Sub cpu_bit_access_write32(ByRef cpu As CPU_t, ByVal value As Long)
    If cpu.mode = 3& Then
        writerm32 cpu, cpu.rm, value
    Else
        cpu_writel cpu, cpu.ea, value
    End If
End Sub

Private Sub cpu_bit_access_write16(ByRef cpu As CPU_t, ByVal value As Long)
    If cpu.mode = 3& Then
        writerm16 cpu, cpu.rm, value
    Else
        cpu_writew cpu, cpu.ea, (value And &HFFFF&)
    End If
End Sub

Public Sub op_ext_A3(ByRef cpu As CPU_t)
    Dim bitOffset As Long
    Dim bitIndex As Long

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        bitOffset = getreg32(cpu, cpu.reg)
        bitIndex = (bitOffset And &H1F&)
        cpu.oper1_32 = cpu_bit_access_read32(cpu, bitOffset, (cpu.mode <> 3&))
        cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
    Else
        bitOffset = cpu_signext16to32(getreg16(cpu, cpu.reg))
        bitIndex = (bitOffset And &HF&)
        cpu.oper1 = cpu_bit_access_read16(cpu, bitOffset, (cpu.mode <> 3&))
        cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
    End If
End Sub

Public Sub op_ext_A4_A5(ByRef cpu As CPU_t)
    Dim count As Long
    Dim codeAddr As Long
    Dim dest As Long
    Dim src As Long
    Dim combo As U64_t
    Dim shifted As U64_t

    modregrm cpu
    If (cpu.opcode And &HFF&) = &HA4& Then
        codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
        count = (cpu_read(cpu, codeAddr) And &HFF&)
        cpu_stepIP cpu, 1&
    Else
        count = cpu_getReg8Low(cpu, CPU_REG_ECX)
    End If

    If cpu.isoper32 <> 0& Then
        count = (count And &H1F&)
        If count <> 0& Then
            dest = readrm32(cpu, cpu.rm)
            src = getreg32(cpu, cpu.reg)
            cpu.cf = CByte(U32Shr(dest, (32& - count)) And &H1&)

            combo = U64_FromParts(src, dest)
            shifted = U64_Shr(combo, (32& - count))
            cpu.res32 = shifted.Lo

            writerm32 cpu, cpu.rm, cpu.res32
            flag_szp32 cpu, cpu.res32

            If count = 1& Then
                cpu.ofl = CByte(((U32Shr(cpu.res32, 31&) And &H1&) Xor (cpu.cf And &H1&)) And &H1&)
            Else
                cpu.ofl = 0&
            End If
        End If
    Else
        count = (count And &H1F&)

        If count <> 0& Then
            dest = readrm16(cpu, cpu.rm)
            src = getreg16(cpu, cpu.reg)
            cpu.cf = CByte(U32Shr((U32Shl((dest And &HFFFF&), 16&) Or (src And &HFFFF&)), (32& - count)) And &H1&)

            cpu.res16 = (U32Shr(U32Shl((U32Shl((dest And &HFFFF&), 16&) Or (src And &HFFFF&)), count), 16&) And &HFFFF&)
            writerm16 cpu, cpu.rm, cpu.res16
            flag_szp16 cpu, cpu.res16

            If count = 1& Then
                cpu.ofl = CByte(((U32Shr(cpu.res16, 15&) And &H1&) Xor (cpu.cf And &H1&)) And &H1&)
            Else
                cpu.ofl = 0&
            End If
        End If
    End If
End Sub

Public Sub op_ext_A6_B0(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = cpu_getReg8Low(cpu, CPU_REG_EAX)
    cpu.oper2b = readrm8(cpu, cpu.rm)
    op_sub8 cpu

    If cpu.zf <> 0& Then
        writerm8 cpu, cpu.rm, getreg8(cpu, cpu.reg)
    Else
        cpu_setReg8Low cpu, CPU_REG_EAX, cpu.oper2b
    End If
End Sub

Public Sub op_ext_A8(ByRef cpu As CPU_t)
    push cpu, getsegreg(cpu, CPU_REG_GS)
End Sub

Public Sub op_ext_A9(ByRef cpu As CPU_t)
    putsegreg cpu, CPU_REG_GS, pop(cpu)
End Sub

Public Sub op_ext_AA(ByRef cpu As CPU_t)
    ' RSM (?)
End Sub

Public Sub op_ext_AB(ByRef cpu As CPU_t)
    Dim bitOffset As Long
    Dim bitIndex As Long
    Dim bitMask As Long

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        bitOffset = getreg32(cpu, cpu.reg)
        bitIndex = (bitOffset And &H1F&)
        bitMask = U32Shl(1&, bitIndex)

        cpu.oper1_32 = cpu_bit_access_read32(cpu, bitOffset, (cpu.mode <> 3&))
        cpu_bit_access_write32 cpu, (cpu.oper1_32 Or bitMask)
        cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
    Else
        bitOffset = cpu_signext16to32(getreg16(cpu, cpu.reg))
        bitIndex = (bitOffset And &HF&)
        bitMask = U32Shl(1&, bitIndex)

        cpu.oper1 = cpu_bit_access_read16(cpu, bitOffset, (cpu.mode <> 3&))
        cpu_bit_access_write16 cpu, (cpu.oper1 Or bitMask)
        cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
    End If
End Sub

Public Sub op_ext_AC_AD(ByRef cpu As CPU_t)
    Dim count As Long
    Dim codeAddr As Long
    Dim signVal As Long
    Dim combo As U64_t
    Dim shifted As U64_t
    Dim shiftedCf As U64_t
    Dim combined As Long

    modregrm cpu

    If (cpu.opcode And &HFF&) = &HAD& Then
        count = cpu_getReg8Low(cpu, CPU_REG_ECX)
    Else
        codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
        count = (cpu_read(cpu, codeAddr) And &HFF&)
        cpu_stepIP cpu, 1&
    End If

    count = (count And &H1F&)
    If count <> 0& Then
        If cpu.isoper32 <> 0& Then
            cpu.oper1_32 = readrm32(cpu, cpu.rm)
            cpu.oper2_32 = getreg32(cpu, cpu.reg)
            signVal = (cpu.oper1_32 And &H80000000)

            combo = U64_FromParts(cpu.oper1_32, cpu.oper2_32)
            shifted = U64_Shr(combo, count)
            shiftedCf = U64_Shr(combo, (count - 1&))
            cpu.cf = CByte(shiftedCf.Lo And &H1&)
            If count = 1& Then
                If (shifted.Lo And &H80000000) <> signVal Then
                    cpu.ofl = 1&
                Else
                    cpu.ofl = 0&
                End If
            End If

            flag_szp32 cpu, shifted.Lo
            writerm32 cpu, cpu.rm, shifted.Lo
        Else
            cpu.oper1 = readrm16(cpu, cpu.rm)
            cpu.oper2 = getreg16(cpu, cpu.reg)
            signVal = (cpu.oper1 And &H8000&)

            combined = ((cpu.oper1 And &HFFFF&) Or U32Shl((cpu.oper2 And &HFFFF&), 16&))
            cpu.res16 = (U32Shr(combined, count) And &HFFFF&)
            cpu.cf = CByte(U32Shr(combined, (count - 1&)) And &H1&)

            If count = 1& Then
                If (cpu.res16 And &H8000&) <> signVal Then
                    cpu.ofl = 1&
                Else
                    cpu.ofl = 0&
                End If
            End If

            flag_szp16 cpu, cpu.res16
            writerm16 cpu, cpu.rm, cpu.res16
        End If
    End If
End Sub

Public Sub op_ext_AF(ByRef cpu As CPU_t)
    Dim src As Long
    Dim dst As Long
    Dim result As Long
    Dim prod64 As U64_t

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        src = readrm32(cpu, cpu.rm)
        dst = getreg32(cpu, cpu.reg)
        prod64 = cpu_signedMul32ToU64(dst, src)

        putreg32 cpu, cpu.reg, prod64.Lo

        If cpu_u64SignExtMatches32(prod64) = 0& Then
            cpu.cf = 1&
            cpu.ofl = 1&
        Else
            cpu.cf = 0&
            cpu.ofl = 0&
        End If
    Else
        src = cpu_signext16to32(readrm16(cpu, cpu.rm))
        dst = cpu_signext16to32(getreg16(cpu, cpu.reg))
        result = (dst * src)

        putreg16 cpu, cpu.reg, result

        If (result < -32768) Or (result > 32767&) Then
            cpu.cf = 1&
            cpu.ofl = 1&
        Else
            cpu.cf = 0&
            cpu.ofl = 0&
        End If
    End If
End Sub

Public Sub op_ext_B1(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu.regs_long(CPU_REG_EAX)
        cpu.oper2_32 = readrm32(cpu, cpu.rm)
        op_sub32 cpu

        If cpu.zf <> 0& Then
            writerm32 cpu, cpu.rm, getreg32(cpu, cpu.reg)
        Else
            cpu.regs_long(CPU_REG_EAX) = cpu.oper2_32
        End If
    Else
        cpu.oper1 = getreg16(cpu, CPU_REG_EAX)
        cpu.oper2 = readrm16(cpu, cpu.rm)
        op_sub16 cpu

        If cpu.zf <> 0& Then
            writerm16 cpu, cpu.rm, getreg16(cpu, cpu.reg)
        Else
            putreg16 cpu, CPU_REG_EAX, cpu.oper2
        End If
    End If
End Sub

Public Sub op_ext_B3(ByRef cpu As CPU_t)
    Dim bitOffset As Long
    Dim bitIndex As Long
    Dim bitMask As Long

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        bitOffset = getreg32(cpu, cpu.reg)
        bitIndex = (bitOffset And &H1F&)
        bitMask = U32Shl(1&, bitIndex)

        cpu.oper1_32 = cpu_bit_access_read32(cpu, bitOffset, (cpu.mode <> 3&))
        cpu_bit_access_write32 cpu, (cpu.oper1_32 And Not bitMask)
        cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
    Else
        bitOffset = cpu_signext16to32(getreg16(cpu, cpu.reg))
        bitIndex = (bitOffset And &HF&)
        bitMask = U32Shl(1&, bitIndex)

        cpu.oper1 = cpu_bit_access_read16(cpu, bitOffset, (cpu.mode <> 3&))
        cpu_bit_access_write16 cpu, (cpu.oper1 And Not bitMask)
        cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
    End If
End Sub

Public Sub op_ext_B2_B4_B5(ByRef cpu As CPU_t)
    modregrm cpu
    getea cpu, cpu.rm

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, cpu_readl(cpu, cpu.ea)
        Select Case (cpu.opcode And &HFF&)
            Case &HB2&
                putsegreg cpu, CPU_REG_SS, (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&)
                cpu_begin_interrupt_shadow cpu
            Case &HB4&
                putsegreg cpu, CPU_REG_FS, (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&)
            Case &HB5&
                putsegreg cpu, CPU_REG_GS, (cpu_readw(cpu, U32Add(cpu.ea, 4&)) And &HFFFF&)
        End Select
    Else
        putreg16 cpu, cpu.reg, (cpu_readw(cpu, cpu.ea) And &HFFFF&)
        Select Case (cpu.opcode And &HFF&)
            Case &HB2&
                putsegreg cpu, CPU_REG_SS, (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&)
                cpu_begin_interrupt_shadow cpu
            Case &HB4&
                putsegreg cpu, CPU_REG_FS, (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&)
            Case &HB5&
                putsegreg cpu, CPU_REG_GS, (cpu_readw(cpu, U32Add(cpu.ea, 2&)) And &HFFFF&)
        End Select
    End If
End Sub

Public Sub op_ext_B6(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, readrm8(cpu, cpu.rm)
    Else
        putreg16 cpu, cpu.reg, readrm8(cpu, cpu.rm)
    End If
End Sub

Public Sub op_ext_B7(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.isoper32 <> 0& Then
        putreg32 cpu, cpu.reg, readrm16(cpu, cpu.rm)
    Else
        putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
    End If
End Sub

Public Sub op_ext_BA(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim bitIndex As Long
    Dim bitMask As Long

    modregrm cpu
    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)

    If cpu.isoper32 <> 0& Then
        cpu.oper2_32 = (cpu_read(cpu, codeAddr) And &HFF&)
        cpu.oper1_32 = cpu_bit_access_read32(cpu, cpu.oper2_32, 0&)
    Else
        cpu.oper2 = (cpu_read(cpu, codeAddr) And &HFF&)
        cpu.oper1 = cpu_bit_access_read16(cpu, cpu.oper2, 0&)
    End If

    cpu_stepIP cpu, 1&

    Select Case (cpu.reg And &H7&)
        Case 4&  ' BT
            If cpu.isoper32 <> 0& Then
                bitIndex = (cpu.oper2_32 And &H1F&)
                cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
            Else
                bitIndex = (cpu.oper2 And &HF&)
                cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
            End If

        Case 5&  ' BTS
            If cpu.isoper32 <> 0& Then
                bitIndex = (cpu.oper2_32 And &H1F&)
                bitMask = U32Shl(1&, bitIndex)
                cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
                cpu_bit_access_write32 cpu, (cpu.oper1_32 Or bitMask)
            Else
                bitIndex = (cpu.oper2 And &HF&)
                bitMask = U32Shl(1&, bitIndex)
                cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
                cpu_bit_access_write16 cpu, (cpu.oper1 Or bitMask)
            End If

        Case 6&  ' BTR
            If cpu.isoper32 <> 0& Then
                bitIndex = (cpu.oper2_32 And &H1F&)
                bitMask = U32Shl(1&, bitIndex)
                cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
                cpu_bit_access_write32 cpu, (cpu.oper1_32 And Not bitMask)
            Else
                bitIndex = (cpu.oper2 And &HF&)
                bitMask = U32Shl(1&, bitIndex)
                cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
                cpu_bit_access_write16 cpu, (cpu.oper1 And Not bitMask)
            End If

        Case 7&  ' BTC
            If cpu.isoper32 <> 0& Then
                bitIndex = (cpu.oper2_32 And &H1F&)
                bitMask = U32Shl(1&, bitIndex)
                cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
                cpu_bit_access_write32 cpu, (cpu.oper1_32 Xor bitMask)
            Else
                bitIndex = (cpu.oper2 And &HF&)
                bitMask = U32Shl(1&, bitIndex)
                cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
                cpu_bit_access_write16 cpu, (cpu.oper1 Xor bitMask)
            End If

        Case Else
            op_ext_illegal cpu
    End Select
End Sub

Public Sub op_ext_BB(ByRef cpu As CPU_t)
    Dim bitOffset As Long
    Dim bitIndex As Long
    Dim bitMask As Long

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        bitOffset = getreg32(cpu, cpu.reg)
        bitIndex = (bitOffset And &H1F&)
        bitMask = U32Shl(1&, bitIndex)

        cpu.oper1_32 = cpu_bit_access_read32(cpu, bitOffset, (cpu.mode <> 3&))
        cpu_bit_access_write32 cpu, (cpu.oper1_32 Xor bitMask)
        cpu.cf = CByte(U32Shr(cpu.oper1_32, bitIndex) And &H1&)
    Else
        bitOffset = cpu_signext16to32(getreg16(cpu, cpu.reg))
        bitIndex = (bitOffset And &HF&)
        bitMask = U32Shl(1&, bitIndex)

        cpu.oper1 = cpu_bit_access_read16(cpu, bitOffset, (cpu.mode <> 3&))
        cpu_bit_access_write16 cpu, (cpu.oper1 Xor bitMask)
        cpu.cf = CByte(U32Shr(cpu.oper1, bitIndex) And &H1&)
    End If
End Sub

Public Sub op_ext_BC(ByRef cpu As CPU_t)
    Dim src As Long
    Dim i As Long

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        src = readrm32(cpu, cpu.rm)
        If src = 0& Then
            cpu.zf = 1&
            Exit Sub
        End If

        cpu.zf = 0&
        For i = 0& To 31&
            If (src And U32Shl(1&, i)) <> 0& Then Exit For
        Next i
        putreg32 cpu, cpu.reg, i
    Else
        src = readrm16(cpu, cpu.rm)
        If src = 0& Then
            cpu.zf = 1&
            Exit Sub
        End If

        cpu.zf = 0&
        For i = 0& To 15&
            If (src And U32Shl(1&, i)) <> 0& Then Exit For
        Next i
        putreg16 cpu, cpu.reg, i
    End If
End Sub

Public Sub op_ext_BD(ByRef cpu As CPU_t)
    Dim src As Long
    Dim i As Long

    modregrm cpu

    If cpu.isoper32 <> 0& Then
        src = readrm32(cpu, cpu.rm)
        If src = 0& Then
            cpu.zf = 1&
            Exit Sub
        End If

        cpu.zf = 0&
        For i = 31& To 0& Step -1&
            If (src And U32Shl(1&, i)) <> 0& Then Exit For
        Next i
        putreg32 cpu, cpu.reg, i
    Else
        src = readrm16(cpu, cpu.rm)
        If src = 0& Then
            cpu.zf = 1&
            Exit Sub
        End If

        cpu.zf = 0&
        For i = 15& To 0& Step -1&
            If (src And U32Shl(1&, i)) <> 0& Then Exit For
        Next i
        putreg16 cpu, cpu.reg, i
    End If
End Sub

Public Sub op_ext_BE(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_signext8to32(readrm8(cpu, cpu.rm))
        putreg32 cpu, cpu.reg, cpu.oper1_32
    Else
        cpu.oper1 = cpu_signext8to16(readrm8(cpu, cpu.rm))
        putreg16 cpu, cpu.reg, cpu.oper1
    End If
End Sub

Public Sub op_ext_BF(ByRef cpu As CPU_t)
    modregrm cpu
    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = cpu_signext16to32(readrm16(cpu, cpu.rm))
        putreg32 cpu, cpu.reg, cpu.oper1_32
    Else
        putreg16 cpu, cpu.reg, readrm16(cpu, cpu.rm)
    End If
End Sub

Public Sub op_ext_C0(ByRef cpu As CPU_t)
    modregrm cpu
    cpu.oper1b = readrm8(cpu, cpu.rm)
    cpu.oper2b = getreg8(cpu, cpu.reg)
    op_add8 cpu
    writerm8 cpu, cpu.rm, cpu.res8
    putreg8 cpu, cpu.reg, cpu.oper1b
End Sub

Public Sub op_ext_C1(ByRef cpu As CPU_t)
    modregrm cpu

    If cpu.isoper32 <> 0& Then
        cpu.oper1_32 = readrm32(cpu, cpu.rm)
        cpu.oper2_32 = getreg32(cpu, cpu.reg)
        op_add32 cpu
        writerm32 cpu, cpu.rm, cpu.res32
        putreg32 cpu, cpu.reg, cpu.oper1_32
    Else
        cpu.oper1 = readrm16(cpu, cpu.rm)
        cpu.oper2 = getreg16(cpu, cpu.reg)
        op_add16 cpu
        writerm16 cpu, cpu.rm, cpu.res16
        putreg16 cpu, cpu.reg, cpu.oper1
    End If
End Sub

Public Sub op_ext_C8_C9_CA_CB_CC_CD_CE_CF(ByRef cpu As CPU_t)
    Dim regIdx As Long
    Dim val As Long
    Dim outVal As Long

    regIdx = (cpu.opcode And &H7&)
    val = cpu.regs_long(regIdx)

    outVal = U32Shl((val And &HFF&), 24&)
    outVal = (outVal Or U32Shl((U32Shr(val, 8&) And &HFF&), 16&))
    outVal = (outVal Or U32Shl((U32Shr(val, 16&) And &HFF&), 8&))
    outVal = (outVal Or (U32Shr(val, 24&) And &HFF&))

    cpu.regs_long(regIdx) = outVal
End Sub

Public Sub op_ext_illegal(ByRef cpu As CPU_t)
    diag_count_illegal
    cpu_raiseException cpu, 6&, 0&

    If showops <> 0& Then
        debug_log DEBUG_ERROR, "[CPU] Invalid opcode exception at " & Hex$(U32Add(cpu.segcache(CPU_REG_CS), cpu_firstip)) & " (op 0F " & Hex$(cpu.opcode And &HFF&) & ")"
    End If
End Sub

Public Sub op_illegal(ByRef cpu As CPU_t)
    diag_count_illegal
    cpu.ip = cpu_firstip
    cpu_raiseException cpu, 6&, 0&

    If showops <> 0& Then
        debug_log DEBUG_ERROR, "[CPU] Invalid opcode exception at " & Hex$(cpu.segregs(CPU_REG_CS)) & ":" & Hex$(cpu_firstip) & " (" & Hex$(cpu.opcode And &HFF&) & ")"
    End If
End Sub

Private Sub cpu_begin_interrupt_shadow(ByRef cpu As CPU_t)
    cpu.interrupt_inhibit = 1&
End Sub

Public Function cpu_interruptCheck(ByRef machine As MACHINE_t, ByVal slave As Long) As Long
    Dim picIdx As Long
    Dim irq As Long
    Dim intnum As Long
    
    If (machine.cpu.doexception) Then
        cpu_interruptCheck = 0&
        Exit Function
    End If

    If slave <> 0& Then
        picIdx = machine.i8259_slave
    Else
        picIdx = machine.i8259
    End If

    If picIdx < 0& Then
        cpu_interruptCheck = 0&
        Exit Function
    End If

    If (machine.cpu.trap_toggle = 0&) And (machine.cpu.ifl <> 0&) Then
        irq = CLng(i8259_nextintr(picIdx))
        If irq <> &HFF& Then
            intnum = (irq + CLng(i8259_getIntOffset(picIdx))) And &HFF&
            machine.cpu.hltstate = 0&
            cpu_intcall machine.cpu, intnum, INT_SOURCE_HARDWARE, 0&
            cpu_interruptCheck = 1&
            Exit Function
        End If
    End If

    cpu_interruptCheck = 0&
End Function

Public Sub cpu_exec(ByRef machine As MACHINE_t, ByVal execloops As Long)
    Dim loopcount As Long
    Dim docontinue As Long
    Dim codeAddr As Long
    Dim opcode As Long
    Dim flagsCopy As Long
    Dim exVal As Long
    Dim exerr As Long
    Dim exIp As Long
    Dim pendingTrap As Byte
    Dim regsBackup(0& To 7&) As Long
    Dim segregsBackup(0& To 5&) As Long
    Dim segcacheBackup(0& To 5&) As Long

    For loopcount = 0& To (execloops - 1&)
        'If timing_pendingDispatch = True Then
        '    timing_loop True
        'End If
        
        CopyMemory regsBackup(0&), machine.cpu.regs_long(0&), 32&
        CopyMemory segregsBackup(0&), machine.cpu.segregs(0&), 24&
        CopyMemory segcacheBackup(0&), machine.cpu.segcache(0&), 24&
        flagsCopy = makeflagsword(machine.cpu)

        machine.cpu.shadow_esp = machine.cpu.regs_long(CPU_REG_ESP)

        machine.cpu.doexception = 0&
        machine.cpu.exceptionerr = 0&
        machine.cpu.nowrite = 0&
        machine.cpu.exceptionip = machine.cpu.ip
        machine.cpu.startcpl = machine.cpu.cpl

        pendingTrap = 0&
        If machine.cpu.trap_toggle <> 0& Then
            cpu_raiseException machine.cpu, 1&, 0&
            machine.cpu.trap_toggle = 0&
            pendingTrap = 1&
        End If

        If pendingTrap = 0& Then
            If machine.cpu.tf <> 0& Then
                machine.cpu.trap_toggle = 1&
            Else
                machine.cpu.trap_toggle = 0&
            End If
        End If

        If machine.cpu.interrupt_inhibit <> 0& Then
            machine.cpu.interrupt_inhibit = CByte(machine.cpu.interrupt_inhibit - 1&)
        ElseIf machine.cpu.doexception = 0& Then
            If ((i8259_devices(0).irr And ((Not i8259_devices(0).IMR) And &HFF&)) <> 0&) Or _
              ((i8259_devices(1).irr And ((Not i8259_devices(1).IMR) And &HFF&)) <> 0&) Then
                Dim didInt As Integer
                
                didInt = 0&
                If cpu_interruptCheck(machine, 0&) <> 0& Then
                    didInt = 1&
                ElseIf cpu_interruptCheck(machine, 1&) <> 0& Then
                    didInt = 1&
                End If
    
                If didInt <> 0& Then
                    CopyMemory regsBackup(0&), machine.cpu.regs_long(0&), 32&
                    CopyMemory segregsBackup(0&), machine.cpu.segregs(0&), 24&
                    CopyMemory segcacheBackup(0&), machine.cpu.segcache(0&), 24&
                    flagsCopy = makeflagsword(machine.cpu)
                    machine.cpu.shadow_esp = machine.cpu.regs_long(CPU_REG_ESP)
                    machine.cpu.exceptionip = machine.cpu.ip
                    machine.cpu.startcpl = machine.cpu.cpl
                End If
            End If
        End If


        If machine.cpu.hltstate <> 0& Then GoTo skipexecution

        machine.cpu.reptype = 0&
        machine.cpu.segoverride = 0&
        machine.cpu.useseg = machine.cpu.segcache(CPU_REG_DS)
        machine.cpu.currentseg = CPU_REG_DS

        If (machine.cpu.segis32(CPU_REG_CS) <> 0&) And (machine.cpu.v86f = 0&) Then
            machine.cpu.isaddr32 = 1&
            machine.cpu.isoper32 = 1&
            cpu_firstip = machine.cpu.ip
        Else
            machine.cpu.isaddr32 = 0&
            machine.cpu.isoper32 = 0&
            cpu_firstip = (machine.cpu.ip And &HFFFF&)
        End If

        docontinue = 0&

        Do While docontinue = 0&
            If (machine.cpu.segis32(CPU_REG_CS) = 0&) Or (machine.cpu.v86f <> 0&) Then
                machine.cpu.ip = (machine.cpu.ip And &HFFFF&)
            End If

            machine.cpu.savecs = machine.cpu.segregs(CPU_REG_CS)
            machine.cpu.saveip = machine.cpu.ip

            codeAddr = U32Add(machine.cpu.segcache(CPU_REG_CS), machine.cpu.ip)
            opcode = (cpu_read(machine.cpu, codeAddr) And &HFF&)
            machine.cpu.opcode = CByte(opcode)
            cpu_stepIP machine.cpu, 1&

            Select Case opcode
                Case &H2E&
                    machine.cpu.useseg = machine.cpu.segcache(CPU_REG_CS)
                    machine.cpu.segoverride = 1&
                    machine.cpu.currentseg = CPU_REG_CS

                Case &H3E&
                    machine.cpu.useseg = machine.cpu.segcache(CPU_REG_DS)
                    machine.cpu.segoverride = 1&
                    machine.cpu.currentseg = CPU_REG_DS

                Case &H26&
                    machine.cpu.useseg = machine.cpu.segcache(CPU_REG_ES)
                    machine.cpu.segoverride = 1&
                    machine.cpu.currentseg = CPU_REG_ES

                Case &H36&
                    machine.cpu.useseg = machine.cpu.segcache(CPU_REG_SS)
                    machine.cpu.segoverride = 1&
                    machine.cpu.currentseg = CPU_REG_SS

                Case &H64&
                    machine.cpu.useseg = machine.cpu.segcache(CPU_REG_FS)
                    machine.cpu.segoverride = 1&
                    machine.cpu.currentseg = CPU_REG_FS

                Case &H65&
                    machine.cpu.useseg = machine.cpu.segcache(CPU_REG_GS)
                    machine.cpu.segoverride = 1&
                    machine.cpu.currentseg = CPU_REG_GS

                Case &H66&
                    machine.cpu.isoper32 = CByte((machine.cpu.isoper32 Xor 1&) And &H1&)

                Case &H67&
                    machine.cpu.isaddr32 = CByte((machine.cpu.isaddr32 Xor 1&) And &H1&)

                Case &HF0&
                    ' LOCK prefix

                Case &HF3&
                    machine.cpu.reptype = 1&

                Case &HF2&
                    machine.cpu.reptype = 2&

                Case Else
                    docontinue = 1&
            End Select
        Loop

        'tmp = tmp + right$("0000" + Hex$(machine.cpu.segregs(CPU_REG_CS)), 4) + ":" + Hex$(cpu_firstip) + " op " + right$("00" + Hex$(machine.cpu.opcode), 2) + vbCrLf
        'debug_log DEBUG_INFO, tmp

        machine.cpu.totalexec_lo = U32Add(machine.cpu.totalexec_lo, 1&)
        If machine.cpu.totalexec_lo = 0& Then
            machine.cpu.totalexec_hi = U32Add(machine.cpu.totalexec_hi, 1&)
        End If

        cpu_dispatchPrimaryOpcode machine.cpu, machine.cpu.opcode

        If machine.cpu.doexception <> 0& Then
            exVal = (machine.cpu.exceptionval And &HFF&)
            exerr = machine.cpu.exceptionerr
            exIp = machine.cpu.exceptionip

            CopyMemory machine.cpu.regs_long(0&), regsBackup(0&), 32&
            CopyMemory machine.cpu.segregs(0&), segregsBackup(0&), 24&
            CopyMemory machine.cpu.segcache(0&), segcacheBackup(0&), 24&
            machine.cpu.exceptionip = exIp
            decodeflagsword machine.cpu, flagsCopy
            wrcache_init
            cpu_intcall machine.cpu, exVal, INT_SOURCE_EXCEPTION, exerr
        End If

skipexecution:
        wrcache_flush
    Next loopcount
End Sub

Private Function cpu_makeflagsword(ByRef cpu As CPU_t) As Long
    Dim flags As Long

    flags = &H2&
    If cpu.cf <> 0& Then flags = (flags Or EFLAGS_CF)
    If cpu.pf <> 0& Then flags = (flags Or EFLAGS_PF)
    If cpu.af <> 0& Then flags = (flags Or EFLAGS_AF)
    If cpu.zf <> 0& Then flags = (flags Or EFLAGS_ZF)
    If cpu.sf <> 0& Then flags = (flags Or EFLAGS_SF)
    If cpu.tf <> 0& Then flags = (flags Or EFLAGS_TF)
    If cpu.ifl <> 0& Then flags = (flags Or EFLAGS_IF)
    If cpu.df <> 0& Then flags = (flags Or EFLAGS_DF)
    If cpu.ofl <> 0& Then flags = (flags Or EFLAGS_OF)

    flags = (flags Or ((cpu.iopl And &H3&) * &H1000&))

    If cpu.nt <> 0& Then flags = (flags Or EFLAGS_NT)
    If cpu.rf <> 0& Then flags = (flags Or EFLAGS_RF)
    If cpu.v86f <> 0& Then flags = (flags Or EFLAGS_VM)
    If cpu.acf <> 0& Then flags = (flags Or EFLAGS_AC)
    If cpu.idf <> 0& Then flags = (flags Or EFLAGS_ID)

    cpu_makeflagsword = flags
End Function

Private Sub cpu_raiseException(ByRef cpu As CPU_t, ByVal exnum As Long, ByVal exerr As Long)
    Dim exAddr As Long
    Dim nextByte As Long

    If cpu.doexception = 0& Then
        cpu.doexception = 1&
        cpu.exceptionval = CByte(exnum And &HFF&)
        cpu.exceptionerr = exerr

        If exnum = 14& Then cpu.nowrite = 1&

        diag_count_exception exnum

        If (exnum And &HFF&) = 0& Then
            exAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.exceptionip)
            nextByte = (cpu_read(cpu, U32Add(exAddr, 1&)) And &HFF&)
            diag_note_divide_fault cpu.opcode, nextByte, cpu.segregs(CPU_REG_CS), cpu.exceptionip
        End If

        If showops <> 0& Then
            debug_log DEBUG_DETAIL, "EX: " & CStr(exnum) & " (" & Hex$(exerr) & ")"
        End If
    End If
End Sub

Public Sub cpu_raiseExceptionFromFirstIP(ByRef cpu As CPU_t, ByVal exnum As Long, ByVal exerr As Long)
    cpu.ip = cpu_firstip
    cpu_raiseException cpu, exnum, exerr
End Sub

Private Function cpu_segtolinear(ByRef cpu As CPU_t, ByVal seg As Long) As Long
    Dim gdtidx As Long
    Dim addr As Long

    If (cpu.protected_mode = 0&) Or (cpu.v86f <> 0&) Then
        cpu_segtolinear = U32Shl((seg And &HFFFF&), 4&)
        Exit Function
    End If

    If (seg And &H4&) <> 0& Then
        gdtidx = U32Add(cpu.ldtr, (seg And &HFFF8&))
    Else
        gdtidx = U32Add(cpu.gdtr, (seg And &HFFF8&))
    End If

    addr = CLng(cpu_read_sys(cpu, U32Add(gdtidx, 2&)))
    addr = addr Or (CLng(cpu_read_sys(cpu, U32Add(gdtidx, 3&))) * &H100&)
    addr = addr Or (CLng(cpu_read_sys(cpu, U32Add(gdtidx, 4&))) * &H10000)
    addr = addr Or (CLng(cpu_read_sys(cpu, U32Add(gdtidx, 7&))) * &H1000000)

    cpu_segtolinear = addr
End Function

Public Function getreg8(ByRef cpu As CPU_t, ByVal reg As Long) As Long
    Select Case (reg And &H7&)
        Case 0&: getreg8 = cpu_getReg8Low(cpu, CPU_REG_EAX)
        Case 1&: getreg8 = cpu_getReg8Low(cpu, CPU_REG_ECX)
        Case 2&: getreg8 = cpu_getReg8Low(cpu, CPU_REG_EDX)
        Case 3&: getreg8 = cpu_getReg8Low(cpu, CPU_REG_EBX)
        Case 4&: getreg8 = cpu_getReg8High(cpu, CPU_REG_EAX)
        Case 5&: getreg8 = cpu_getReg8High(cpu, CPU_REG_ECX)
        Case 6&: getreg8 = cpu_getReg8High(cpu, CPU_REG_EDX)
        Case Else: getreg8 = cpu_getReg8High(cpu, CPU_REG_EBX)
    End Select
End Function

Public Sub putreg8(ByRef cpu As CPU_t, ByVal reg As Long, ByVal value As Long)
    value = (value And &HFF&)

    Select Case (reg And &H7&)
        Case 0&: cpu_setReg8Low cpu, CPU_REG_EAX, value
        Case 1&: cpu_setReg8Low cpu, CPU_REG_ECX, value
        Case 2&: cpu_setReg8Low cpu, CPU_REG_EDX, value
        Case 3&: cpu_setReg8Low cpu, CPU_REG_EBX, value
        Case 4&: cpu_setReg8High cpu, CPU_REG_EAX, value
        Case 5&: cpu_setReg8High cpu, CPU_REG_ECX, value
        Case 6&: cpu_setReg8High cpu, CPU_REG_EDX, value
        Case Else: cpu_setReg8High cpu, CPU_REG_EBX, value
    End Select
End Sub

Public Function getreg16(ByRef cpu As CPU_t, ByVal reg As Long) As Long
    Select Case (reg And &H7&)
        Case 0&: getreg16 = (cpu.regs_long(CPU_REG_EAX) And &HFFFF&)
        Case 1&: getreg16 = (cpu.regs_long(CPU_REG_ECX) And &HFFFF&)
        Case 2&: getreg16 = (cpu.regs_long(CPU_REG_EDX) And &HFFFF&)
        Case 3&: getreg16 = (cpu.regs_long(CPU_REG_EBX) And &HFFFF&)
        Case 4&: getreg16 = (cpu.regs_long(CPU_REG_ESP) And &HFFFF&)
        Case 5&: getreg16 = (cpu.regs_long(CPU_REG_EBP) And &HFFFF&)
        Case 6&: getreg16 = (cpu.regs_long(CPU_REG_ESI) And &HFFFF&)
        Case Else: getreg16 = (cpu.regs_long(CPU_REG_EDI) And &HFFFF&)
    End Select
End Function

Public Sub putreg16(ByRef cpu As CPU_t, ByVal reg As Long, ByVal writeval As Long)
    Dim v As Long

    v = (writeval And &HFFFF&)

    Select Case (reg And &H7&)
        Case 0&: cpu.regs_long(CPU_REG_EAX) = ((cpu.regs_long(CPU_REG_EAX) And &HFFFF0000) Or v)
        Case 1&: cpu.regs_long(CPU_REG_ECX) = ((cpu.regs_long(CPU_REG_ECX) And &HFFFF0000) Or v)
        Case 2&: cpu.regs_long(CPU_REG_EDX) = ((cpu.regs_long(CPU_REG_EDX) And &HFFFF0000) Or v)
        Case 3&: cpu.regs_long(CPU_REG_EBX) = ((cpu.regs_long(CPU_REG_EBX) And &HFFFF0000) Or v)
        Case 4&: cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or v)
        Case 5&: cpu.regs_long(CPU_REG_EBP) = ((cpu.regs_long(CPU_REG_EBP) And &HFFFF0000) Or v)
        Case 6&: cpu.regs_long(CPU_REG_ESI) = ((cpu.regs_long(CPU_REG_ESI) And &HFFFF0000) Or v)
        Case Else: cpu.regs_long(CPU_REG_EDI) = ((cpu.regs_long(CPU_REG_EDI) And &HFFFF0000) Or v)
    End Select
End Sub

Public Function getreg32(ByRef cpu As CPU_t, ByVal reg As Long) As Long
    getreg32 = cpu.regs_long(reg And &H7&)
End Function

Public Sub putreg32(ByRef cpu As CPU_t, ByVal reg As Long, ByVal writeval As Long)
    cpu.regs_long(reg And &H7&) = writeval
End Sub

Public Function cpu_getReg8Low(ByRef cpu As CPU_t, ByVal reg32 As Long) As Long
    cpu_getReg8Low = (cpu.regs_long(reg32 And &H7&) And &HFF&)
End Function

Public Function cpu_getReg8High(ByRef cpu As CPU_t, ByVal reg32 As Long) As Long
    cpu_getReg8High = (U32Shr(cpu.regs_long(reg32 And &H7&), 8&) And &HFF&)
End Function

Public Sub cpu_setReg8Low(ByRef cpu As CPU_t, ByVal reg32 As Long, ByVal value As Long)
    cpu.regs_long(reg32 And &H7&) = ((cpu.regs_long(reg32 And &H7&) And &HFFFFFF00) Or (value And &HFF&))
End Sub

Public Sub cpu_setReg8High(ByRef cpu As CPU_t, ByVal reg32 As Long, ByVal value As Long)
    cpu.regs_long(reg32 And &H7&) = ((cpu.regs_long(reg32 And &H7&) And &HFFFF00FF) Or U32Shl((value And &HFF&), 8&))
End Sub

Public Function getsegreg(ByRef cpu As CPU_t, ByVal reg As Long) As Long
    If (reg < 0&) Or (reg > 5&) Then
        getsegreg = 0&
    Else
        getsegreg = cpu.segregs(reg)
    End If
End Function

Public Sub putsegreg(ByRef cpu As CPU_t, ByVal reg As Long, ByVal writeval As Long)
    Dim seg As Long
    Dim gdtidx As Long
    Dim fault As Byte
    Dim seglimit As Long
    Dim table_limit As Long
    Dim selector As Long
    Dim descFlags As Long

    If (reg < 0&) Or (reg > 5&) Then Exit Sub

    selector = (writeval And &HFFFF&)

    If (cpu.protected_mode <> 0&) And (cpu.v86f = 0&) Then
        fault = 0&
        seg = (selector And &HFFF8&)

        If (selector And &H4&) <> 0& Then
            table_limit = cpu.ldtl
            gdtidx = U32Add(cpu.ldtr, seg)
        Else
            table_limit = cpu.gdtl
            gdtidx = U32Add(cpu.gdtr, seg)
        End If

        If U32Add(seg, 7&) > table_limit Then
            fault = 1&
        Else
            If (cpu_read_sys(cpu, U32Add(gdtidx, 5&)) And &H80&) = 0& Then
                fault = 1&
            End If
        End If

        If (fault <> 0&) And (reg = CPU_REG_CS) Then
            cpu_raiseException cpu, 13&, (selector And &HFFFC&)
            Exit Sub
        End If

        ' Canonical C always updates descriptor-derived segment metadata in protected mode.
        descFlags = (cpu_read_sys(cpu, U32Add(gdtidx, 6&)) And &HFF&)
        cpu.segis32(reg) = CByte((descFlags \ 64&) And 1&)
        seglimit = (cpu_readw_sys(cpu, gdtidx) And &HFFFF&)
        seglimit = seglimit Or ((descFlags And &HF&) * &H10000)
        If (descFlags And &H80&) <> 0& Then
            seglimit = U32Shl(seglimit, 12&)
            seglimit = seglimit Or &HFFF&
        End If
        cpu.seglimit(reg) = seglimit
    Else
        cpu.segis32(reg) = 0&
    End If

    cpu.segregs(reg) = selector
    cpu.segcache(reg) = cpu_segtolinear(cpu, selector)

    If (reg = CPU_REG_CS) And (cpu.protected_mode <> 0&) Then
        If cpu.v86f <> 0& Then
            cpu.cpl = 3&
        Else
            cpu.cpl = (selector And &H3&)
        End If
    End If
End Sub

Public Sub pushw(ByRef cpu As CPU_t, ByVal pushval As Long)
    Dim sp As Long
    Dim addr As Long

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        sp = U32Sub(cpu.regs_long(CPU_REG_ESP), 2&)
        cpu.regs_long(CPU_REG_ESP) = sp
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
    Else
        sp = (cpu.regs_long(CPU_REG_ESP) And &HFFFF&)
        sp = (sp - 2&) And &HFFFF&
        cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or sp)
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
    End If

    cpu_writew cpu, addr, (pushval And &HFFFF&)
End Sub

Public Sub pushl(ByRef cpu As CPU_t, ByVal pushval As Long)
    Dim sp As Long
    Dim addr As Long

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        sp = U32Sub(cpu.regs_long(CPU_REG_ESP), 4&)
        cpu.regs_long(CPU_REG_ESP) = sp
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
    Else
        sp = (cpu.regs_long(CPU_REG_ESP) And &HFFFF&)
        sp = (sp - 4&) And &HFFFF&
        cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or sp)
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
    End If

    cpu_writel cpu, addr, pushval
End Sub

Public Function popw(ByRef cpu As CPU_t) As Long
    Dim sp As Long
    Dim addr As Long

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        sp = cpu.regs_long(CPU_REG_ESP)
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
        popw = (cpu_readw(cpu, addr) And &HFFFF&)
        cpu.regs_long(CPU_REG_ESP) = U32Add(sp, 2&)
    Else
        sp = (cpu.regs_long(CPU_REG_ESP) And &HFFFF&)
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
        popw = (cpu_readw(cpu, addr) And &HFFFF&)
        sp = (sp + 2&) And &HFFFF&
        cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or sp)
    End If
End Function

Public Function popl(ByRef cpu As CPU_t) As Long
    Dim sp As Long
    Dim addr As Long

    If cpu.segis32(CPU_REG_SS) <> 0& Then
        sp = cpu.regs_long(CPU_REG_ESP)
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
        popl = cpu_readl(cpu, addr)
        cpu.regs_long(CPU_REG_ESP) = U32Add(sp, 4&)
    Else
        sp = (cpu.regs_long(CPU_REG_ESP) And &HFFFF&)
        addr = U32Add(cpu.segcache(CPU_REG_SS), sp)
        popl = cpu_readl(cpu, addr)
        sp = (sp + 4&) And &HFFFF&
        cpu.regs_long(CPU_REG_ESP) = ((cpu.regs_long(CPU_REG_ESP) And &HFFFF0000) Or sp)
    End If
End Function

Public Function makeflagsword(ByRef cpu As CPU_t) As Long
    makeflagsword = cpu_makeflagsword(cpu)
End Function

Public Sub decodeflagsword(ByRef cpu As CPU_t, ByVal flags As Long)
    Dim tmp As Long

    tmp = flags

    cpu.cf = CByte(tmp And &H1&)
    cpu.pf = CByte(U32Shr(tmp, 2&) And &H1&)
    cpu.af = CByte(U32Shr(tmp, 4&) And &H1&)
    cpu.zf = CByte(U32Shr(tmp, 6&) And &H1&)
    cpu.sf = CByte(U32Shr(tmp, 7&) And &H1&)
    cpu.tf = CByte(U32Shr(tmp, 8&) And &H1&)

    If (cpu.cpl = 0&) Or (cpu.protected_mode = 0&) Then
        cpu.ifl = CByte(U32Shr(tmp, 9&) And &H1&)
    End If

    cpu.df = CByte(U32Shr(tmp, 10&) And &H1&)
    cpu.ofl = CByte(U32Shr(tmp, 11&) And &H1&)

    If (cpu.cpl = 0&) Or (cpu.protected_mode = 0&) Then
        cpu.iopl = CByte(U32Shr(tmp, 12&) And &H3&)
    End If

    cpu.nt = CByte(U32Shr(tmp, 14&) And &H1&)
    cpu.rf = CByte(U32Shr(tmp, 16&) And &H1&)
    cpu.v86f = CByte(U32Shr(tmp, 17&) And &H1&)
    If cpu.v86f <> 0& Then cpu.cpl = 3&
    cpu.acf = CByte(U32Shr(tmp, 18&) And &H1&)
    cpu.idf = CByte(U32Shr(tmp, 21&) And &H1&)
End Sub

Public Function segtolinear(ByRef cpu As CPU_t, ByVal seg As Long) As Long
    segtolinear = cpu_segtolinear(cpu, seg)
End Function

Private Sub cpu_stepIP(ByRef cpu As CPU_t, ByVal amount As Long)
    If (cpu.segis32(CPU_REG_CS) = 0&) Or (cpu.v86f <> 0&) Then
        cpu.ip = (cpu.ip + amount) And &HFFFF&
        Exit Sub
    End If

    cpu.ip = U32Add(cpu.ip, amount)
End Sub

Private Sub cpu_apply_relative_branch(ByRef cpu As CPU_t, ByVal displacement As Long)
    cpu.ip = U32Add(cpu.ip, displacement)
    If cpu.isoper32 = 0& Then
        cpu.ip = (cpu.ip And &HFFFF&)
    End If
End Sub

Private Function cpu_signext8to16(ByVal value As Long) As Long
    value = (value And &HFF&)
    If (value And &H80&) <> 0& Then
        cpu_signext8to16 = ((value Or &HFF00&) And &HFFFF&)
    Else
        cpu_signext8to16 = value
    End If
End Function

Private Function cpu_signext8to32(ByVal value As Long) As Long
    value = (value And &HFF&)
    If (value And &H80&) <> 0& Then
        cpu_signext8to32 = (value Or &HFFFFFF00)
    Else
        cpu_signext8to32 = value
    End If
End Function

Private Sub cpu_initParityTable()
    Dim i As Long
    Dim bitCount As Long
    Dim v As Long
    Dim b As Long

    If cpu_parityInit <> 0& Then Exit Sub

    For i = 0& To 255&
        v = i
        bitCount = 0&
        For b = 0& To 7&
            bitCount = bitCount + (v And 1&)
            v = (v \ 2&)
        Next b

        If (bitCount And 1&) = 0& Then
            cpu_parityTable(i) = 1&
        Else
            cpu_parityTable(i) = 0&
        End If
    Next i

    cpu_parityInit = 1&
End Sub

Private Function cpu_parityEven(ByVal value As Long) As Long
    If cpu_parityInit = 0& Then
        cpu_initParityTable
    End If

    cpu_parityEven = cpu_parityTable(value And &HFF&)
End Function

Private Function cpu_u32Add3(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByRef carryOut As Long) As Long
    Dim alo As Long
    Dim blo As Long
    Dim Lo As Long
    Dim carry16 As Long
    Dim ahi As Long
    Dim bhi As Long
    Dim Hi As Long
    Dim resultVal As Long

    alo = (a And &HFFFF&)
    blo = (b And &HFFFF&)
    Lo = alo + blo + (c And &H1&)
    carry16 = (Lo \ &H10000)
    Lo = (Lo And &HFFFF&)

    ahi = ((a And &H7FFF0000) \ &H10000)
    If a < 0& Then ahi = (ahi Or &H8000&)
    bhi = ((b And &H7FFF0000) \ &H10000)
    If b < 0& Then bhi = (bhi Or &H8000&)

    Hi = ahi + bhi + carry16
    carryOut = (Hi \ &H10000)
    Hi = (Hi And &HFFFF&)

    resultVal = (((Hi And &H7FFF&) * &H10000) Or Lo)
    If (Hi And &H8000&) <> 0& Then
        resultVal = (resultVal Or &H80000000)
    End If

    cpu_u32Add3 = resultVal
End Function

Private Function cpu_u32Sub3(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByRef borrowOut As Long) As Long
    Dim alo As Long
    Dim blo As Long
    Dim Lo As Long
    Dim borrow16 As Long
    Dim ahi As Long
    Dim bhi As Long
    Dim Hi As Long
    Dim resultVal As Long

    alo = (a And &HFFFF&)
    blo = (b And &HFFFF&)
    Lo = alo - blo - (c And &H1&)
    If Lo < 0& Then
        Lo = Lo + &H10000
        borrow16 = 1&
    Else
        borrow16 = 0&
    End If

    ahi = ((a And &H7FFF0000) \ &H10000)
    If a < 0& Then ahi = (ahi Or &H8000&)
    bhi = ((b And &H7FFF0000) \ &H10000)
    If b < 0& Then bhi = (bhi Or &H8000&)

    Hi = ahi - bhi - borrow16
    If Hi < 0& Then
        Hi = Hi + &H10000
        borrowOut = 1&
    Else
        borrowOut = 0&
    End If
    Hi = (Hi And &HFFFF&)

    resultVal = (((Hi And &H7FFF&) * &H10000) Or Lo)
    If (Hi And &H8000&) <> 0& Then
        resultVal = (resultVal Or &H80000000)
    End If

    cpu_u32Sub3 = resultVal
End Function

Public Sub modregrm(ByRef cpu As CPU_t)
    Dim codeAddr As Long
    Dim indexVal As Long
    Dim baseVal As Long

    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
    cpu.addrbyte = (cpu_read(cpu, codeAddr) And &HFF&)
    cpu_stepIP cpu, 1&

    cpu.mode = CByte((cpu.addrbyte \ &H40&) And &H3&)
    cpu.reg = CByte((cpu.addrbyte \ &H8&) And &H7&)
    cpu.rm = CByte(cpu.addrbyte And &H7&)

    If cpu.isaddr32 = 0& Then
        Select Case cpu.mode
            Case 0&
                If cpu.rm = 6& Then
                    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
                    cpu.disp16 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
                    cpu_stepIP cpu, 2&
                End If

                If (((cpu.rm = 2&) Or (cpu.rm = 3&)) And (cpu.segoverride = 0&)) Then
                    cpu.useseg = cpu.segcache(CPU_REG_SS)
                End If

            Case 1&
                codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
                cpu.disp16 = cpu_signext8to16(cpu_read(cpu, codeAddr))
                cpu_stepIP cpu, 1&

                If (((cpu.rm = 2&) Or (cpu.rm = 3&) Or (cpu.rm = 6&)) And (cpu.segoverride = 0&)) Then
                    cpu.useseg = cpu.segcache(CPU_REG_SS)
                End If

            Case 2&
                codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
                cpu.disp16 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
                cpu_stepIP cpu, 2&

                If (((cpu.rm = 2&) Or (cpu.rm = 3&) Or (cpu.rm = 6&)) And (cpu.segoverride = 0&)) Then
                    cpu.useseg = cpu.segcache(CPU_REG_SS)
                End If

            Case Else
                cpu.disp8 = 0&
                cpu.disp16 = 0&
        End Select
    Else
        cpu.sib_val = 0&

        If (cpu.mode < 3&) And (cpu.rm = 4&) Then
            codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
            cpu.sib = CByte(cpu_read(cpu, codeAddr) And &HFF&)
            cpu_stepIP cpu, 1&
            cpu.sib_scale = CByte((cpu.sib \ &H40&) And &H3&)
            cpu.sib_index = CByte((cpu.sib \ &H8&) And &H7&)
            cpu.sib_base = CByte(cpu.sib And &H7&)
        End If

        If (cpu.segoverride = 0&) And (cpu.mode < 3&) Then
            If cpu.rm = 4& Then
                If (cpu.sib_base = CPU_REG_ESP) Or ((cpu.sib_base = CPU_REG_EBP) And (cpu.mode > 0&)) Then
                    cpu.useseg = cpu.segcache(CPU_REG_SS)
                End If
            ElseIf (cpu.rm = 5&) And (cpu.mode > 0&) Then
                cpu.useseg = cpu.segcache(CPU_REG_SS)
            End If
        End If

        Select Case cpu.mode
            Case 0&
                If (cpu.rm = 5&) Or ((cpu.rm = 4&) And (cpu.sib_base = 5&)) Then
                    codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
                    cpu.disp32 = cpu_readl(cpu, codeAddr)
                    cpu_stepIP cpu, 4&
                Else
                    cpu.disp32 = 0&
                End If

            Case 1&
                codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
                cpu.disp32 = cpu_signext8to32(cpu_read(cpu, codeAddr))
                cpu_stepIP cpu, 1&

            Case 2&
                codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
                cpu.disp32 = cpu_readl(cpu, codeAddr)
                cpu_stepIP cpu, 4&

            Case Else
                cpu.disp32 = 0&
        End Select

        If (cpu.mode < 3&) And (cpu.rm = 4&) Then
            If cpu.sib_index = CPU_REG_ESP Then
                indexVal = 0&
            Else
                indexVal = U32Shl(cpu.regs_long(cpu.sib_index), (cpu.sib_scale And &H3&))
            End If

            If (cpu.mode = 0&) And (cpu.sib_base = CPU_REG_EBP) Then
                baseVal = 0&
            Else
                If cpu.sib_base = CPU_REG_ESP Then
                    baseVal = cpu.shadow_esp
                Else
                    baseVal = cpu.regs_long(cpu.sib_base)
                End If

                If (cpu.segoverride = 0&) And (cpu.sib_base = CPU_REG_ESP) Then
                    cpu.useseg = cpu.segcache(CPU_REG_SS)
                End If
            End If

            cpu.sib_val = U32Add(baseVal, indexVal)
        End If
    End If
End Sub

Private Function cpu_effective_offset(ByRef cpu As CPU_t, ByVal rmval As Long) As Long
    Dim tempea As Long

    rmval = (rmval And &H7&)
    tempea = 0&

    If cpu.isaddr32 = 0& Then
        Select Case cpu.mode
            Case 0&
                Select Case rmval
                    Case 0&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EBX), getreg16(cpu, CPU_REG_ESI))
                    Case 1&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EBX), getreg16(cpu, CPU_REG_EDI))
                    Case 2&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EBP), getreg16(cpu, CPU_REG_ESI))
                    Case 3&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EBP), getreg16(cpu, CPU_REG_EDI))
                    Case 4&
                        tempea = getreg16(cpu, CPU_REG_ESI)
                    Case 5&
                        tempea = getreg16(cpu, CPU_REG_EDI)
                    Case 6&
                        tempea = cpu.disp16
                    Case 7&
                        tempea = getreg16(cpu, CPU_REG_EBX)
                End Select

            Case 1&, 2&
                Select Case rmval
                    Case 0&
                        tempea = U32Add(U32Add(getreg16(cpu, CPU_REG_EBX), getreg16(cpu, CPU_REG_ESI)), cpu.disp16)
                    Case 1&
                        tempea = U32Add(U32Add(getreg16(cpu, CPU_REG_EBX), getreg16(cpu, CPU_REG_EDI)), cpu.disp16)
                    Case 2&
                        tempea = U32Add(U32Add(getreg16(cpu, CPU_REG_EBP), getreg16(cpu, CPU_REG_ESI)), cpu.disp16)
                    Case 3&
                        tempea = U32Add(U32Add(getreg16(cpu, CPU_REG_EBP), getreg16(cpu, CPU_REG_EDI)), cpu.disp16)
                    Case 4&
                        tempea = U32Add(getreg16(cpu, CPU_REG_ESI), cpu.disp16)
                    Case 5&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EDI), cpu.disp16)
                    Case 6&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EBP), cpu.disp16)
                    Case 7&
                        tempea = U32Add(getreg16(cpu, CPU_REG_EBX), cpu.disp16)
                End Select
        End Select
    Else
        Select Case cpu.mode
            Case 0&
                Select Case rmval
                    Case 0&
                        tempea = cpu.regs_long(CPU_REG_EAX)
                    Case 1&
                        tempea = cpu.regs_long(CPU_REG_ECX)
                    Case 2&
                        tempea = cpu.regs_long(CPU_REG_EDX)
                    Case 3&
                        tempea = cpu.regs_long(CPU_REG_EBX)
                    Case 4&
                        tempea = cpu.sib_val
                        If cpu.sib_base = 5& Then
                            tempea = U32Add(tempea, cpu.disp32)
                        End If
                    Case 5&
                        tempea = cpu.disp32
                    Case 6&
                        tempea = cpu.regs_long(CPU_REG_ESI)
                    Case 7&
                        tempea = cpu.regs_long(CPU_REG_EDI)
                End Select

            Case 1&, 2&
                Select Case rmval
                    Case 0&
                        tempea = U32Add(cpu.regs_long(CPU_REG_EAX), cpu.disp32)
                    Case 1&
                        tempea = U32Add(cpu.regs_long(CPU_REG_ECX), cpu.disp32)
                    Case 2&
                        tempea = U32Add(cpu.regs_long(CPU_REG_EDX), cpu.disp32)
                    Case 3&
                        tempea = U32Add(cpu.regs_long(CPU_REG_EBX), cpu.disp32)
                    Case 4&
                        tempea = U32Add(cpu.sib_val, cpu.disp32)
                    Case 5&
                        tempea = U32Add(cpu.regs_long(CPU_REG_EBP), cpu.disp32)
                    Case 6&
                        tempea = U32Add(cpu.regs_long(CPU_REG_ESI), cpu.disp32)
                    Case 7&
                        tempea = U32Add(cpu.regs_long(CPU_REG_EDI), cpu.disp32)
                End Select
        End Select
    End If

    If cpu.isaddr32 <> 0& Then
        cpu_effective_offset = tempea
    Else
        cpu_effective_offset = (tempea And &HFFFF&)
    End If
End Function

Public Sub getea(ByRef cpu As CPU_t, ByVal rmval As Long)
    Dim tempea As Long

    tempea = cpu_effective_offset(cpu, rmval)

    If cpu.isaddr32 <> 0& Then
        cpu.ea = U32Add(tempea, cpu.useseg)
    Else
        cpu.ea = U32Add((tempea And &HFFFF&), cpu.useseg)
    End If
End Sub

Public Function readrm16(ByRef cpu As CPU_t, ByVal rmval As Long) As Long
    If cpu.mode < 3& Then
        getea cpu, rmval
        readrm16 = (cpu_readw(cpu, cpu.ea) And &HFFFF&)
    Else
        readrm16 = getreg16(cpu, rmval)
    End If
End Function

Public Function readrm32(ByRef cpu As CPU_t, ByVal rmval As Long) As Long
    If cpu.mode < 3& Then
        getea cpu, rmval
        readrm32 = cpu_readl(cpu, cpu.ea)
    Else
        readrm32 = getreg32(cpu, rmval)
    End If
End Function

Public Function readrm64(ByRef cpu As CPU_t, ByVal rmval As Long) As U64_t
    If cpu.mode < 3& Then
        getea cpu, rmval
        readrm64.Lo = cpu_readl(cpu, cpu.ea)
        readrm64.Hi = cpu_readl(cpu, U32Add(cpu.ea, 4&))
    Else
        readrm64.Lo = getreg32(cpu, rmval)
        readrm64.Hi = 0&
    End If
End Function

Public Function readrm8(ByRef cpu As CPU_t, ByVal rmval As Long) As Long
    If cpu.mode < 3& Then
        getea cpu, rmval
        readrm8 = (cpu_read(cpu, cpu.ea) And &HFF&)
    Else
        Select Case (rmval And &H7&)
            Case 0&: readrm8 = cpu_getReg8Low(cpu, CPU_REG_EAX)
            Case 1&: readrm8 = cpu_getReg8Low(cpu, CPU_REG_ECX)
            Case 2&: readrm8 = cpu_getReg8Low(cpu, CPU_REG_EDX)
            Case 3&: readrm8 = cpu_getReg8Low(cpu, CPU_REG_EBX)
            Case 4&: readrm8 = cpu_getReg8High(cpu, CPU_REG_EAX)
            Case 5&: readrm8 = cpu_getReg8High(cpu, CPU_REG_ECX)
            Case 6&: readrm8 = cpu_getReg8High(cpu, CPU_REG_EDX)
            Case Else: readrm8 = cpu_getReg8High(cpu, CPU_REG_EBX)
        End Select
    End If
End Function

Public Sub writerm32(ByRef cpu As CPU_t, ByVal rmval As Long, ByVal value As Long)
    If cpu.mode < 3& Then
        getea cpu, rmval
        cpu_writel cpu, cpu.ea, value
    Else
        putreg32 cpu, rmval, value
    End If
End Sub

Public Sub writerm64(ByRef cpu As CPU_t, ByVal rmval As Long, ByRef value As U64_t)
    If cpu.mode < 3& Then
        getea cpu, rmval
        cpu_writel cpu, cpu.ea, value.Lo
        cpu_writel cpu, U32Add(cpu.ea, 4&), value.Hi
    Else
        putreg32 cpu, rmval, value.Lo
    End If
End Sub

Public Sub writerm16(ByRef cpu As CPU_t, ByVal rmval As Long, ByVal value As Long)
    If cpu.mode < 3& Then
        getea cpu, rmval
        cpu_writew cpu, cpu.ea, (value And &HFFFF&)
    Else
        putreg16 cpu, rmval, value
    End If
End Sub

Public Sub writerm8(ByRef cpu As CPU_t, ByVal rmval As Long, ByVal value As Long)
    value = (value And &HFF&)

    If cpu.mode < 3& Then
        getea cpu, rmval
        cpu_write cpu, cpu.ea, value
    Else
        Select Case (rmval And &H7&)
            Case 0&: cpu_setReg8Low cpu, CPU_REG_EAX, value
            Case 1&: cpu_setReg8Low cpu, CPU_REG_ECX, value
            Case 2&: cpu_setReg8Low cpu, CPU_REG_EDX, value
            Case 3&: cpu_setReg8Low cpu, CPU_REG_EBX, value
            Case 4&: cpu_setReg8High cpu, CPU_REG_EAX, value
            Case 5&: cpu_setReg8High cpu, CPU_REG_ECX, value
            Case 6&: cpu_setReg8High cpu, CPU_REG_EDX, value
            Case Else: cpu_setReg8High cpu, CPU_REG_EBX, value
        End Select
    End If
End Sub

Public Sub push(ByRef cpu As CPU_t, ByVal pushval As Long)
    If cpu.isoper32 <> 0& Then
        pushl cpu, pushval
    Else
        pushw cpu, pushval
    End If
End Sub

Public Function pop(ByRef cpu As CPU_t) As Long
    If cpu.isoper32 <> 0& Then
        pop = popl(cpu)
    Else
        pop = popw(cpu)
    End If
End Function

Public Sub flag_szp8(ByRef cpu As CPU_t, ByVal value As Long)
    value = (value And &HFF&)

    If value = 0& Then
        cpu.zf = 1&
    Else
        cpu.zf = 0&
    End If

    If (value And &H80&) <> 0& Then
        cpu.sf = 1&
    Else
        cpu.sf = 0&
    End If

    cpu.pf = CByte(cpu_parityEven(value))
End Sub

Public Sub flag_szp16(ByRef cpu As CPU_t, ByVal value As Long)
    value = (value And &HFFFF&)

    If value = 0& Then
        cpu.zf = 1&
    Else
        cpu.zf = 0&
    End If

    If (value And &H8000&) <> 0& Then
        cpu.sf = 1&
    Else
        cpu.sf = 0&
    End If

    cpu.pf = CByte(cpu_parityEven(value And &HFF&))
End Sub

Public Sub flag_szp32(ByRef cpu As CPU_t, ByVal value As Long)
    If value = 0& Then
        cpu.zf = 1&
    Else
        cpu.zf = 0&
    End If

    If value < 0& Then
        cpu.sf = 1&
    Else
        cpu.sf = 0&
    End If

    cpu.pf = CByte(cpu_parityEven(value And &HFF&))
End Sub

Public Sub flag_log8(ByRef cpu As CPU_t, ByVal value As Long)
    flag_szp8 cpu, value
    cpu.cf = 0&
    cpu.ofl = 0&
End Sub

Public Sub flag_log16(ByRef cpu As CPU_t, ByVal value As Long)
    flag_szp16 cpu, value
    cpu.cf = 0&
    cpu.ofl = 0&
End Sub

Public Sub flag_log32(ByRef cpu As CPU_t, ByVal value As Long)
    flag_szp32 cpu, value
    cpu.cf = 0&
    cpu.ofl = 0&
End Sub

Public Sub flag_adc8(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim dst As Long

    v1 = (v1 And &HFF&)
    v2 = (v2 And &HFF&)
    v3 = (v3 And &H1&)

    dst = v1 + v2 + v3
    flag_szp8 cpu, dst

    If (((dst Xor v1) And (dst Xor v2) And &H80&) = &H80&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If (dst And &H100&) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If (((v1 Xor v2 Xor dst) And &H10&) = &H10&) Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_adc16(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim dst As Long

    v1 = (v1 And &HFFFF&)
    v2 = (v2 And &HFFFF&)
    v3 = (v3 And &H1&)

    dst = v1 + v2 + v3
    flag_szp16 cpu, dst

    If (((dst Xor v1) And (dst Xor v2) And &H8000&) = &H8000&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If (dst And &H10000) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If (((v1 Xor v2 Xor dst) And &H10&) = &H10&) Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_adc32(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim dst As Long
    Dim carryOut As Long

    dst = cpu_u32Add3(v1, v2, v3, carryOut)
    flag_szp32 cpu, dst

    If (((dst Xor v1) And (dst Xor v2)) < 0&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    cpu.cf = CByte(carryOut And &H1&)

    If (((v1 Xor v2 Xor dst) And &H10&) <> 0&) Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_add8(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long)
    Dim dst As Long

    v1 = (v1 And &HFF&)
    v2 = (v2 And &HFF&)

    dst = v1 + v2
    flag_szp8 cpu, dst

    If (dst And &H100&) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If (((dst Xor v1) And (dst Xor v2) And &H80&) = &H80&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If (((v1 Xor v2 Xor dst) And &H10&) = &H10&) Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_add16(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long)
    Dim dst As Long

    v1 = (v1 And &HFFFF&)
    v2 = (v2 And &HFFFF&)

    dst = v1 + v2
    flag_szp16 cpu, dst

    If (dst And &H10000) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If (((dst Xor v1) And (dst Xor v2) And &H8000&) = &H8000&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If (((v1 Xor v2 Xor dst) And &H10&) = &H10&) Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_add32(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long)
    Dim dst As Long
    Dim carryOut As Long

    dst = cpu_u32Add3(v1, v2, 0&, carryOut)
    flag_szp32 cpu, dst

    cpu.cf = CByte(carryOut And &H1&)

    If (((dst Xor v1) And (dst Xor v2)) < 0&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If (((v1 Xor v2 Xor dst) And &H10&) <> 0&) Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_sbb8(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim dst As Long
    Dim srcWithCf As Long

    v1 = (v1 And &HFF&)
    v2 = (v2 And &HFF&)
    v3 = (v3 And &H1&)

    srcWithCf = (v2 + v3)
    dst = v1 - srcWithCf
    flag_szp8 cpu, dst

    If (dst And &H100&) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If ((dst Xor v1) And (v1 Xor v2) And &H80&) <> 0& Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If ((v1 Xor v2 Xor dst) And &H10&) <> 0& Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_sbb16(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim dst As Long
    Dim srcWithCf As Long

    v1 = (v1 And &HFFFF&)
    v2 = (v2 And &HFFFF&)
    v3 = (v3 And &H1&)

    srcWithCf = (v2 + v3)
    dst = v1 - srcWithCf
    flag_szp16 cpu, dst

    If (dst And &H10000) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If ((dst Xor v1) And (v1 Xor v2) And &H8000&) <> 0& Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If ((v1 Xor v2 Xor dst) And &H10&) <> 0& Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_sbb32(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim dst As Long
    Dim borrowOut As Long

    v3 = (v3 And &H1&)
    dst = cpu_u32Sub3(v1, v2, v3, borrowOut)
    flag_szp32 cpu, dst

    cpu.cf = CByte(borrowOut And &H1&)

    If (((dst Xor v1) And (v1 Xor v2)) < 0&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If ((v1 Xor v2 Xor dst) And &H10&) <> 0& Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_sub8(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long)
    Dim dst As Long

    v1 = (v1 And &HFF&)
    v2 = (v2 And &HFF&)

    dst = v1 - v2
    flag_szp8 cpu, dst

    If (dst And &H100&) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If ((dst Xor v1) And (v1 Xor v2) And &H80&) <> 0& Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If ((v1 Xor v2 Xor dst) And &H10&) <> 0& Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_sub16(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long)
    Dim dst As Long

    v1 = (v1 And &HFFFF&)
    v2 = (v2 And &HFFFF&)

    dst = v1 - v2
    flag_szp16 cpu, dst

    If (dst And &H10000) <> 0& Then
        cpu.cf = 1&
    Else
        cpu.cf = 0&
    End If

    If ((dst Xor v1) And (v1 Xor v2) And &H8000&) <> 0& Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If ((v1 Xor v2 Xor dst) And &H10&) <> 0& Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub flag_sub32(ByRef cpu As CPU_t, ByVal v1 As Long, ByVal v2 As Long)
    Dim dst As Long
    Dim borrowOut As Long

    dst = cpu_u32Sub3(v1, v2, 0&, borrowOut)
    flag_szp32 cpu, dst

    cpu.cf = CByte(borrowOut And &H1&)

    If (((dst Xor v1) And (v1 Xor v2)) < 0&) Then
        cpu.ofl = 1&
    Else
        cpu.ofl = 0&
    End If

    If ((v1 Xor v2 Xor dst) And &H10&) <> 0& Then
        cpu.af = 1&
    Else
        cpu.af = 0&
    End If
End Sub

Public Sub op_adc8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) + (cpu.oper2b And &HFF&) + (cpu.cf And &H1&)) And &HFF&
    flag_adc8 cpu, cpu.oper1b, cpu.oper2b, cpu.cf
End Sub

Public Sub op_adc16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) + (cpu.oper2 And &HFFFF&) + (cpu.cf And &H1&)) And &HFFFF&
    flag_adc16 cpu, cpu.oper1, cpu.oper2, cpu.cf
End Sub

Public Sub op_adc32(ByRef cpu As CPU_t)
    cpu.res32 = U32Add(U32Add(cpu.oper1_32, cpu.oper2_32), (cpu.cf And &H1&))
    flag_adc32 cpu, cpu.oper1_32, cpu.oper2_32, cpu.cf
End Sub

Public Sub op_add8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) + (cpu.oper2b And &HFF&)) And &HFF&
    flag_add8 cpu, cpu.oper1b, cpu.oper2b
End Sub

Public Sub op_add16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) + (cpu.oper2 And &HFFFF&)) And &HFFFF&
    flag_add16 cpu, cpu.oper1, cpu.oper2
End Sub

Public Sub op_add32(ByRef cpu As CPU_t)
    cpu.res32 = U32Add(cpu.oper1_32, cpu.oper2_32)
    flag_add32 cpu, cpu.oper1_32, cpu.oper2_32
End Sub

Public Sub op_and8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) And (cpu.oper2b And &HFF&))
    flag_log8 cpu, cpu.res8
End Sub

Public Sub op_and16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) And (cpu.oper2 And &HFFFF&))
    flag_log16 cpu, cpu.res16
End Sub

Public Sub op_and32(ByRef cpu As CPU_t)
    cpu.res32 = (cpu.oper1_32 And cpu.oper2_32)
    flag_log32 cpu, cpu.res32
End Sub

Public Sub op_or8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) Or (cpu.oper2b And &HFF&))
    flag_log8 cpu, cpu.res8
End Sub

Public Sub op_or16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) Or (cpu.oper2 And &HFFFF&))
    flag_log16 cpu, cpu.res16
End Sub

Public Sub op_or32(ByRef cpu As CPU_t)
    cpu.res32 = (cpu.oper1_32 Or cpu.oper2_32)
    flag_log32 cpu, cpu.res32
End Sub

Public Sub op_xor8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) Xor (cpu.oper2b And &HFF&))
    flag_log8 cpu, cpu.res8
End Sub

Public Sub op_xor16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) Xor (cpu.oper2 And &HFFFF&))
    flag_log16 cpu, cpu.res16
End Sub

Public Sub op_xor32(ByRef cpu As CPU_t)
    cpu.res32 = (cpu.oper1_32 Xor cpu.oper2_32)
    flag_log32 cpu, cpu.res32
End Sub

Public Sub op_sub8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) - (cpu.oper2b And &HFF&)) And &HFF&
    flag_sub8 cpu, cpu.oper1b, cpu.oper2b
End Sub

Public Sub op_sub16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) - (cpu.oper2 And &HFFFF&)) And &HFFFF&
    flag_sub16 cpu, cpu.oper1, cpu.oper2
End Sub

Public Sub op_sub32(ByRef cpu As CPU_t)
    cpu.res32 = U32Sub(cpu.oper1_32, cpu.oper2_32)
    flag_sub32 cpu, cpu.oper1_32, cpu.oper2_32
End Sub

Public Sub op_sbb8(ByRef cpu As CPU_t)
    cpu.res8 = ((cpu.oper1b And &HFF&) - ((cpu.oper2b And &HFF&) + (cpu.cf And &H1&))) And &HFF&
    flag_sbb8 cpu, cpu.oper1b, cpu.oper2b, cpu.cf
End Sub

Public Sub op_sbb16(ByRef cpu As CPU_t)
    cpu.res16 = ((cpu.oper1 And &HFFFF&) - ((cpu.oper2 And &HFFFF&) + (cpu.cf And &H1&))) And &HFFFF&
    flag_sbb16 cpu, cpu.oper1, cpu.oper2, cpu.cf
End Sub

Public Sub op_sbb32(ByRef cpu As CPU_t)
    cpu.res32 = U32Sub(cpu.oper1_32, U32Add(cpu.oper2_32, (cpu.cf And &H1&)))
    flag_sbb32 cpu, cpu.oper1_32, cpu.oper2_32, cpu.cf
End Sub

Private Function cpu_u64Bit31(ByRef v As U64_t) As Long
    If (v.Lo And U32Shl(1&, 31&)) <> 0& Then
        cpu_u64Bit31 = 1&
    Else
        cpu_u64Bit31 = 0&
    End If
End Function

Private Function cpu_u64Bit0(ByRef v As U64_t) As Long
    cpu_u64Bit0 = (v.Lo And &H1&)
End Function

Private Function cpu_u64FromBit31(ByVal bitVal As Long) As U64_t
    cpu_u64FromBit31.Hi = 0&
    If (bitVal And &H1&) <> 0& Then
        cpu_u64FromBit31.Lo = U32Shl(1&, 31&)
    Else
        cpu_u64FromBit31.Lo = 0&
    End If
End Function

Public Function op_grp2_8(ByRef cpu As CPU_t, ByVal cnt As Long) As Long
    Dim s As Long
    Dim shift As Long
    Dim oldcf As Long
    Dim msb As Long

    s = (cpu.oper1b And &HFF&)
    oldcf = (cpu.cf And &H1&)
    cnt = (cnt And &H1F&)
    If cnt = 0& Then
        op_grp2_8 = (s And &HFF&)
        Exit Function
    End If

    Select Case (cpu.reg And &H7&)
        Case 0&  ' ROL
            cnt = (cnt And &H7&)
            If cnt = 0& Then
                op_grp2_8 = (s And &HFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                If (s And &H80&) <> 0& Then
                    cpu.cf = 1&
                Else
                    cpu.cf = 0&
                End If

                s = (U32Shl(s, 1&) And &HFFFF&)
                s = (s Or (cpu.cf And &H1&))
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((U32Shr(s, 7&) And &H1&) Xor (cpu.cf And &H1&)) And &H1&)
            End If

        Case 1&  ' ROR
            cnt = (cnt And &H7&)
            If cnt = 0& Then
                op_grp2_8 = (s And &HFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                cpu.cf = CByte(s And &H1&)
                s = (U32Shr(s, 1&) Or U32Shl((cpu.cf And &H1&), 7&))
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((U32Shr(s, 7&) Xor (U32Shr(s, 6&) And &H1&)) And &H1&)
            End If

        Case 2&  ' RCL
            cnt = (cnt Mod 9&)
            If cnt = 0& Then
                op_grp2_8 = (s And &HFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                oldcf = (cpu.cf And &H1&)

                If (s And &H80&) <> 0& Then
                    cpu.cf = 1&
                Else
                    cpu.cf = 0&
                End If

                s = (U32Shl(s, 1&) And &HFFFF&)
                s = (s Or oldcf)
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((cpu.cf And &H1&) Xor (U32Shr(s, 7&) And &H1&)) And &H1&)
            End If

        Case 3&  ' RCR
            cnt = (cnt Mod 9&)
            If cnt = 0& Then
                op_grp2_8 = (s And &HFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                oldcf = (cpu.cf And &H1&)
                cpu.cf = CByte(s And &H1&)
                s = (U32Shr(s, 1&) Or U32Shl(oldcf, 7&))
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((U32Shr(s, 7&) Xor (U32Shr(s, 6&) And &H1&)) And &H1&)
            End If

        Case 4&, 6&  ' SHL
            For shift = 1& To cnt
                If (s And &H80&) <> 0& Then
                    cpu.cf = 1&
                Else
                    cpu.cf = 0&
                End If

                s = (U32Shl(s, 1&) And &HFF&)
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((U32Shr(s, 7&) And &H1&) Xor (cpu.cf And &H1&)) And &H1&)
            End If

            flag_szp8 cpu, s

        Case 5&  ' SHR
            If (cnt = 1&) And ((s And &H80&) <> 0&) Then
                cpu.ofl = 1&
            Else
                cpu.ofl = 0&
            End If

            For shift = 1& To cnt
                cpu.cf = CByte(s And &H1&)
                s = U32Shr(s, 1&)
            Next shift

            flag_szp8 cpu, s

        Case 7&  ' SAR
            For shift = 1& To cnt
                msb = (s And &H80&)
                cpu.cf = CByte(s And &H1&)
                s = (U32Shr(s, 1&) Or msb)
            Next shift

            cpu.ofl = 0&
            flag_szp8 cpu, s
    End Select

    op_grp2_8 = (s And &HFF&)
End Function

Public Function op_grp2_16(ByRef cpu As CPU_t, ByVal cnt As Long) As Long
    Dim s As Long
    Dim shift As Long
    Dim oldcf As Long
    Dim msb As Long

    s = (cpu.oper1 And &HFFFF&)
    oldcf = (cpu.cf And &H1&)
    cnt = (cnt And &H1F&)
    If cnt = 0& Then
        op_grp2_16 = (s And &HFFFF&)
        Exit Function
    End If

    Select Case (cpu.reg And &H7&)
        Case 0&  ' ROL
            cnt = (cnt And &HF&)
            If cnt = 0& Then
                op_grp2_16 = (s And &HFFFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                If (s And &H8000&) <> 0& Then
                    cpu.cf = 1&
                Else
                    cpu.cf = 0&
                End If

                s = U32Shl(s, 1&)
                s = (s Or (cpu.cf And &H1&))
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((cpu.cf And &H1&) Xor (U32Shr(s, 15&) And &H1&)) And &H1&)
            End If

        Case 1&  ' ROR
            cnt = (cnt And &HF&)
            If cnt = 0& Then
                op_grp2_16 = (s And &HFFFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                cpu.cf = CByte(s And &H1&)
                s = (U32Shr(s, 1&) Or U32Shl((cpu.cf And &H1&), 15&))
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((U32Shr(s, 15&) Xor (U32Shr(s, 14&) And &H1&)) And &H1&)
            End If

        Case 2&  ' RCL
            cnt = (cnt Mod 17&)
            If cnt = 0& Then
                op_grp2_16 = (s And &HFFFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                oldcf = (cpu.cf And &H1&)

                If (s And &H8000&) <> 0& Then
                    cpu.cf = 1&
                Else
                    cpu.cf = 0&
                End If

                s = U32Shl(s, 1&)
                s = (s Or oldcf)
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((cpu.cf And &H1&) Xor (U32Shr(s, 15&) And &H1&)) And &H1&)
            End If

        Case 3&  ' RCR
            cnt = (cnt Mod 17&)
            If cnt = 0& Then
                op_grp2_16 = (s And &HFFFF&)
                Exit Function
            End If

            For shift = 1& To cnt
                oldcf = (cpu.cf And &H1&)
                cpu.cf = CByte(s And &H1&)
                s = (U32Shr(s, 1&) Or U32Shl(oldcf, 15&))
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((U32Shr(s, 15&) Xor (U32Shr(s, 14&) And &H1&)) And &H1&)
            End If

        Case 4&, 6&  ' SHL
            For shift = 1& To cnt
                If (s And &H8000&) <> 0& Then
                    cpu.cf = 1&
                Else
                    cpu.cf = 0&
                End If

                s = (U32Shl(s, 1&) And &HFFFF&)
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((U32Shr(s, 15&) And &H1&) Xor (cpu.cf And &H1&)) And &H1&)
            End If

            flag_szp16 cpu, s

        Case 5&  ' SHR
            If (cnt = 1&) And ((s And &H8000&) <> 0&) Then
                cpu.ofl = 1&
            Else
                cpu.ofl = 0&
            End If

            For shift = 1& To cnt
                cpu.cf = CByte(s And &H1&)
                s = U32Shr(s, 1&)
            Next shift

            flag_szp16 cpu, s

        Case 7&  ' SAR
            For shift = 1& To cnt
                msb = (s And &H8000&)
                cpu.cf = CByte(s And &H1&)
                s = (U32Shr(s, 1&) Or msb)
            Next shift

            cpu.ofl = 0&
            flag_szp16 cpu, s
    End Select

    op_grp2_16 = (s And &HFFFF&)
End Function

Public Function op_grp2_32(ByRef cpu As CPU_t, ByVal cnt As Long) As Long
    Dim s As U64_t
    Dim shift As Long
    Dim oldcf As Long
    Dim msb64 As U64_t
    Dim tmp64 As U64_t

    s = U64_FromU32(cpu.oper1_32)
    oldcf = (cpu.cf And &H1&)
    cnt = (cnt And &H1F&)
    If cnt = 0& Then
        op_grp2_32 = s.Lo
        Exit Function
    End If

    Select Case (cpu.reg And &H7&)
        Case 0&  ' ROL
            For shift = 1& To cnt
                cpu.cf = CByte(cpu_u64Bit31(s))
                s = U64_Shl(s, 1&)
                If (cpu.cf And &H1&) <> 0& Then
                    s.Lo = (s.Lo Or &H1&)
                End If
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((cpu.cf And &H1&) Xor cpu_u64Bit31(s)) And &H1&)
            End If

        Case 1&  ' ROR
            For shift = 1& To cnt
                cpu.cf = CByte(cpu_u64Bit0(s))
                s = U64_Shr(s, 1&)
                If (cpu.cf And &H1&) <> 0& Then
                    s.Lo = (s.Lo Or U32Shl(1&, 31&))
                End If
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((cpu_u64Bit31(s) Xor (U32Shr(s.Lo, 30&) And &H1&)) And &H1&)
            End If

        Case 2&  ' RCL
            For shift = 1& To cnt
                oldcf = (cpu.cf And &H1&)
                cpu.cf = CByte(cpu_u64Bit31(s))
                s = U64_Shl(s, 1&)
                If (oldcf And &H1&) <> 0& Then
                    s.Lo = (s.Lo Or &H1&)
                End If
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte(((cpu.cf And &H1&) Xor cpu_u64Bit31(s)) And &H1&)
            End If

        Case 3&  ' RCR
            For shift = 1& To cnt
                oldcf = (cpu.cf And &H1&)
                cpu.cf = CByte(cpu_u64Bit0(s))
                s = U64_Shr(s, 1&)
                If (oldcf And &H1&) <> 0& Then
                    s.Lo = (s.Lo Or U32Shl(1&, 31&))
                End If
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((cpu_u64Bit31(s) Xor (U32Shr(s.Lo, 30&) And &H1&)) And &H1&)
            End If

        Case 4&, 6&  ' SHL
            For shift = 1& To cnt
                cpu.cf = CByte(cpu_u64Bit31(s))
                s = U64_Shl(s, 1&)
                s.Hi = 0&
            Next shift

            If cnt = 1& Then
                cpu.ofl = CByte((cpu_u64Bit31(s) Xor (cpu.cf And &H1&)) And &H1&)
            End If

            flag_szp32 cpu, s.Lo

        Case 5&  ' SHR
            If (cnt = 1&) And (cpu_u64Bit31(s) <> 0&) Then
                cpu.ofl = 1&
            Else
                cpu.ofl = 0&
            End If

            For shift = 1& To cnt
                cpu.cf = CByte(cpu_u64Bit0(s))
                s = U64_Shr(s, 1&)
            Next shift

            flag_szp32 cpu, s.Lo

        Case 7&  ' SAR
            For shift = 1& To cnt
                msb64 = cpu_u64FromBit31(cpu_u64Bit31(s))
                cpu.cf = CByte(cpu_u64Bit0(s))
                tmp64 = U64_Shr(s, 1&)
                s = U64_Or(tmp64, msb64)
            Next shift

            cpu.ofl = 0&
            flag_szp32 cpu, s.Lo
    End Select

    op_grp2_32 = s.Lo
End Function

Private Function cpu_signext16to32(ByVal value As Long) As Long
    value = (value And &HFFFF&)
    If (value And &H8000&) <> 0& Then
        cpu_signext16to32 = (value Or &HFFFF0000)
    Else
        cpu_signext16to32 = value
    End If
End Function

Private Function cpu_u64IsNegative(ByRef v As U64_t) As Long
    If v.Hi < 0& Then
        cpu_u64IsNegative = 1&
    Else
        cpu_u64IsNegative = 0&
    End If
End Function

Private Function cpu_u64Neg(ByRef v As U64_t) As U64_t
    Dim inv As U64_t
    Dim one As U64_t

    inv.Lo = Not v.Lo
    inv.Hi = Not v.Hi

    one.Lo = 1&
    one.Hi = 0&

    cpu_u64Neg = U64_Add(inv, one)
End Function

Private Function cpu_u64Compare(ByRef a As U64_t, ByRef b As U64_t) As Long
    If U32Lt(a.Hi, b.Hi) <> 0& Then
        cpu_u64Compare = -1&
        Exit Function
    End If

    If U32Lt(b.Hi, a.Hi) <> 0& Then
        cpu_u64Compare = 1&
        Exit Function
    End If

    If U32Lt(a.Lo, b.Lo) <> 0& Then
        cpu_u64Compare = -1&
    ElseIf U32Lt(b.Lo, a.Lo) <> 0& Then
        cpu_u64Compare = 1&
    Else
        cpu_u64Compare = 0&
    End If
End Function

Private Function cpu_u64GetBit(ByRef v As U64_t, ByVal bitIdx As Long) As Long
    If bitIdx >= 32& Then
        cpu_u64GetBit = (U32Shr(v.Hi, bitIdx - 32&) And &H1&)
    Else
        cpu_u64GetBit = (U32Shr(v.Lo, bitIdx) And &H1&)
    End If
End Function

Private Sub cpu_u64SetBit(ByRef v As U64_t, ByVal bitIdx As Long)
    If bitIdx >= 32& Then
        v.Hi = (v.Hi Or U32Shl(1&, bitIdx - 32&))
    Else
        v.Lo = (v.Lo Or U32Shl(1&, bitIdx))
    End If
End Sub

Private Sub cpu_u64DivModU32(ByRef dividend As U64_t, ByVal divisor As Long, ByRef quotient As U64_t, ByRef remOut As Long)
    Dim i As Long
    Dim bitVal As Long
    Dim rem64 As U64_t
    Dim div64 As U64_t

    quotient.Lo = 0&
    quotient.Hi = 0&
    rem64.Lo = 0&
    rem64.Hi = 0&

    div64.Lo = divisor
    div64.Hi = 0&

    For i = 63& To 0& Step -1&
        rem64 = U64_Shl(rem64, 1&)
        bitVal = cpu_u64GetBit(dividend, i)
        If bitVal <> 0& Then
            rem64.Lo = (rem64.Lo Or &H1&)
        End If

        If cpu_u64Compare(rem64, div64) >= 0& Then
            rem64 = U64_Sub(rem64, div64)
            cpu_u64SetBit quotient, i
        End If
    Next i

    remOut = rem64.Lo
End Sub

Private Function cpu_u32MulToU64(ByVal a As Long, ByVal b As Long) As U64_t
    Dim a0 As Long
    Dim a1 As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim p0Lo As Long
    Dim p0Hi As Long
    Dim p1Lo As Long
    Dim p1Hi As Long
    Dim p2Lo As Long
    Dim p2Hi As Long
    Dim p3Lo As Long
    Dim p3Hi As Long
    Dim t As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim r3 As Long
    Dim carry As Long

    a0 = (a And &HFFFF&)
    a1 = (U32Shr(a, 16&) And &HFFFF&)
    b0 = (b And &HFFFF&)
    b1 = (U32Shr(b, 16&) And &HFFFF&)

    cpu_u16MulToU32Parts a0, b0, p0Lo, p0Hi
    cpu_u16MulToU32Parts a0, b1, p1Lo, p1Hi
    cpu_u16MulToU32Parts a1, b0, p2Lo, p2Hi
    cpu_u16MulToU32Parts a1, b1, p3Lo, p3Hi

    t = (p0Hi + p1Lo + p2Lo)
    r1 = (t And &HFFFF&)
    carry = (t \ &H10000)

    t = (p1Hi + p2Hi + p3Lo + carry)
    r2 = (t And &HFFFF&)
    carry = (t \ &H10000)

    r3 = ((p3Hi + carry) And &HFFFF&)

    cpu_u32MulToU64.Lo = ((p0Lo And &HFFFF&) Or U32Shl((r1 And &HFFFF&), 16&))
    cpu_u32MulToU64.Hi = ((r2 And &HFFFF&) Or U32Shl((r3 And &HFFFF&), 16&))
End Function

Private Sub cpu_u16MulToU32Parts(ByVal a As Long, ByVal b As Long, ByRef loPart As Long, ByRef hiPart As Long)
    Dim a0 As Long
    Dim a1 As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim p0 As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim temp As Long

    a = (a And &HFFFF&)
    b = (b And &HFFFF&)

    a0 = (a And &HFF&)
    a1 = ((a \ &H100&) And &HFF&)
    b0 = (b And &HFF&)
    b1 = ((b \ &H100&) And &HFF&)

    p0 = (a0 * b0)
    p1 = ((a0 * b1) + (a1 * b0))
    p2 = (a1 * b1)

    temp = (p0 + ((p1 And &HFF&) * &H100&))
    loPart = (temp And &HFFFF&)
    hiPart = ((p2 + (p1 \ &H100&) + (temp \ &H10000)) And &HFFFF&)
End Sub

Private Function cpu_makeU64(ByVal loPart As Long, ByVal hiPart As Long) As U64_t
    cpu_makeU64.Lo = loPart
    cpu_makeU64.Hi = hiPart
End Function

Public Sub op_div8(ByRef cpu As CPU_t, ByVal valdiv As Long, ByVal divisor As Long)
    Dim q As Long

    divisor = (divisor And &HFF&)
    valdiv = (valdiv And &HFFFF&)

    If divisor = 0& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    q = (valdiv \ divisor)
    If q > &HFF& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    cpu_setReg8High cpu, CPU_REG_EAX, (valdiv Mod divisor)
    cpu_setReg8Low cpu, CPU_REG_EAX, q
End Sub

Public Sub op_idiv8(ByRef cpu As CPU_t, ByVal valdiv As Long, ByVal divisor As Long)
    Dim dividend As Long
    Dim signedDivisor As Long
    Dim quotient As Long
    Dim remainderVal As Long

    dividend = cpu_signext16to32(valdiv And &HFFFF&)
    signedDivisor = cpu_signext8to32(divisor And &HFF&)

    If signedDivisor = 0& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    quotient = (dividend \ signedDivisor)
    remainderVal = (dividend Mod signedDivisor)

    If (quotient < -128&) Or (quotient > 127&) Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    cpu_setReg8High cpu, CPU_REG_EAX, (remainderVal And &HFF&)
    cpu_setReg8Low cpu, CPU_REG_EAX, (quotient And &HFF&)
End Sub

Public Sub op_grp3_8(ByRef cpu As CPU_t)
    Dim imm8 As Long
    Dim codeAddr As Long
    Dim imulResult As Long

    Select Case (cpu.reg And &H7&)
        Case 0&, 1&  ' TEST
            codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
            imm8 = (cpu_read(cpu, codeAddr) And &HFF&)
            flag_log8 cpu, ((cpu.oper1b And &HFF&) And imm8)
            cpu_stepIP cpu, 1&

        Case 2&  ' NOT
            cpu.res8 = ((Not cpu.oper1b) And &HFF&)

        Case 3&  ' NEG
            cpu.res8 = (((Not cpu.oper1b) + 1&) And &HFF&)
            flag_sub8 cpu, 0&, cpu.oper1b
            If cpu.res8 = 0& Then
                cpu.cf = 0&
            Else
                cpu.cf = 1&
            End If

        Case 4&  ' MUL
            cpu.temp1 = ((cpu.oper1b And &HFF&) * cpu_getReg8Low(cpu, CPU_REG_EAX))
            putreg16 cpu, CPU_REG_EAX, cpu.temp1

            If cpu_getReg8High(cpu, CPU_REG_EAX) <> 0& Then
                cpu.cf = 1&
                cpu.ofl = 1&
            Else
                cpu.cf = 0&
                cpu.ofl = 0&
            End If

        Case 5&  ' IMUL
            cpu.temp1 = cpu_signext8to32(cpu_getReg8Low(cpu, CPU_REG_EAX))
            cpu.temp2 = cpu_signext8to32(cpu.oper1b)
            imulResult = (cpu.temp1 * cpu.temp2)
            putreg16 cpu, CPU_REG_EAX, (imulResult And &HFFFF&)

            If (imulResult < -128&) Or (imulResult > 127&) Then
                cpu.cf = 1&
                cpu.ofl = 1&
            Else
                cpu.cf = 0&
                cpu.ofl = 0&
            End If

        Case 6&  ' DIV
            op_div8 cpu, getreg16(cpu, CPU_REG_EAX), cpu.oper1b

        Case 7&  ' IDIV
            op_idiv8 cpu, getreg16(cpu, CPU_REG_EAX), cpu.oper1b
    End Select
End Sub

Public Sub op_div16(ByRef cpu As CPU_t, ByVal valdiv As Long, ByVal divisor As Long)
    Dim dividend64 As U64_t
    Dim q As U64_t
    Dim remVal As Long

    divisor = (divisor And &HFFFF&)

    If divisor = 0& Then
        diag_note_divide_math 0&, valdiv, 0&, divisor, 1&
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    dividend64.Lo = valdiv
    dividend64.Hi = 0&
    cpu_u64DivModU32 dividend64, divisor, q, remVal

    If (q.Hi <> 0&) Or ((q.Lo And &HFFFF0000) <> 0&) Then
        diag_note_divide_math 0&, valdiv, 0&, divisor, 2&
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    putreg16 cpu, CPU_REG_EDX, remVal
    putreg16 cpu, CPU_REG_EAX, q.Lo
End Sub

Public Sub op_div32(ByRef cpu As CPU_t, ByRef valdiv As U64_t, ByVal divisor As Long)
    Dim q As U64_t
    Dim remVal As Long

    If divisor = 0& Then
        diag_note_divide_math 1&, valdiv.Lo, valdiv.Hi, divisor, 1&
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    cpu_u64DivModU32 valdiv, divisor, q, remVal

    If q.Hi <> 0& Then
        diag_note_divide_math 1&, valdiv.Lo, valdiv.Hi, divisor, 2&
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    cpu.regs_long(CPU_REG_EDX) = remVal
    cpu.regs_long(CPU_REG_EAX) = q.Lo
End Sub

Public Sub op_idiv16(ByRef cpu As CPU_t, ByVal valdiv As Long, ByVal divisor As Long)
    Dim dividend As Long
    Dim signedDivisor As Long
    Dim quotient As Long
    Dim remainderVal As Long

    signedDivisor = cpu_signext16to32(divisor And &HFFFF&)
    If signedDivisor = 0& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    dividend = valdiv
    If (dividend = &H80000000) And (signedDivisor = -1&) Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    quotient = (dividend \ signedDivisor)
    remainderVal = (dividend Mod signedDivisor)

    If (quotient < -32768) Or (quotient > 32767&) Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    putreg16 cpu, CPU_REG_EAX, (quotient And &HFFFF&)
    putreg16 cpu, CPU_REG_EDX, (remainderVal And &HFFFF&)
End Sub

Public Sub op_idiv32(ByRef cpu As CPU_t, ByRef valdiv As U64_t, ByVal divisor As Long)
    Dim signedDivisor As Long
    Dim absDividend As U64_t
    Dim absDivisor As Long
    Dim q As U64_t
    Dim remVal As Long
    Dim dividendNeg As Long
    Dim divisorNeg As Long
    Dim quotientNeg As Long
    Dim quotientRaw As Long
    Dim remainderRaw As Long

    signedDivisor = divisor
    If signedDivisor = 0& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    If (valdiv.Hi = &H80000000) And (valdiv.Lo = 0&) And (signedDivisor = -1&) Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    dividendNeg = cpu_u64IsNegative(valdiv)
    If dividendNeg <> 0& Then
        absDividend = cpu_u64Neg(valdiv)
    Else
        absDividend = valdiv
    End If

    If signedDivisor < 0& Then
        divisorNeg = 1&
        absDivisor = U32Add(Not signedDivisor, 1&)
    Else
        divisorNeg = 0&
        absDivisor = signedDivisor
    End If

    quotientNeg = (dividendNeg Xor divisorNeg)
    cpu_u64DivModU32 absDividend, absDivisor, q, remVal

    If q.Hi <> 0& Then
        cpu_raiseException cpu, 0&, 0&
        Exit Sub
    End If

    If quotientNeg <> 0& Then
        If U32Lt(&H80000000, q.Lo) <> 0& Then
            cpu_raiseException cpu, 0&, 0&
            Exit Sub
        End If
        quotientRaw = U32Add(Not q.Lo, 1&)
    Else
        If U32Lt(&H7FFFFFFF, q.Lo) <> 0& Then
            cpu_raiseException cpu, 0&, 0&
            Exit Sub
        End If
        quotientRaw = q.Lo
    End If

    If dividendNeg <> 0& Then
        remainderRaw = U32Add(Not remVal, 1&)
    Else
        remainderRaw = remVal
    End If

    cpu.regs_long(CPU_REG_EAX) = quotientRaw
    cpu.regs_long(CPU_REG_EDX) = remainderRaw
End Sub

Public Sub op_grp3_16(ByRef cpu As CPU_t)
    Dim imm16 As Long
    Dim codeAddr As Long
    Dim axv As Long
    Dim dxv As Long
    Dim prod64 As U64_t
    Dim valdiv As Long

    Select Case (cpu.reg And &H7&)
        Case 0&, 1&  ' TEST
            codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
            imm16 = (cpu_readw(cpu, codeAddr) And &HFFFF&)
            flag_log16 cpu, ((cpu.oper1 And &HFFFF&) And imm16)
            cpu_stepIP cpu, 2&

        Case 2&  ' NOT
            cpu.res16 = ((Not cpu.oper1) And &HFFFF&)

        Case 3&  ' NEG
            cpu.res16 = (((Not cpu.oper1) + 1&) And &HFFFF&)
            flag_sub16 cpu, 0&, cpu.oper1
            If cpu.res16 <> 0& Then
                cpu.cf = 1&
            Else
                cpu.cf = 0&
            End If

        Case 4&  ' MUL
            axv = getreg16(cpu, CPU_REG_EAX)
            prod64 = cpu_u32MulToU64((cpu.oper1 And &HFFFF&), (axv And &HFFFF&))
            cpu.temp1 = prod64.Lo

            putreg16 cpu, CPU_REG_EAX, cpu.temp1
            putreg16 cpu, CPU_REG_EDX, (U32Shr(cpu.temp1, 16&) And &HFFFF&)

            If getreg16(cpu, CPU_REG_EDX) <> 0& Then
                cpu.cf = 1&
                cpu.ofl = 1&
            Else
                cpu.cf = 0&
                cpu.ofl = 0&
            End If

        Case 5&  ' IMUL
            cpu.temp1 = cpu_signext16to32(getreg16(cpu, CPU_REG_EAX))
            cpu.temp2 = cpu_signext16to32(cpu.oper1)
            cpu.temp3 = (cpu.temp1 * cpu.temp2)

            putreg16 cpu, CPU_REG_EAX, cpu.temp3
            putreg16 cpu, CPU_REG_EDX, U32Shr(cpu.temp3, 16&)

            If (cpu.temp3 < -32768) Or (cpu.temp3 > 32767&) Then
                cpu.cf = 1&
                cpu.ofl = 1&
            Else
                cpu.cf = 0&
                cpu.ofl = 0&
            End If

        Case 6&  ' DIV
            dxv = getreg16(cpu, CPU_REG_EDX)
            axv = getreg16(cpu, CPU_REG_EAX)
            valdiv = U32Add(U32Shl(dxv, 16&), axv)
            op_div16 cpu, valdiv, cpu.oper1

        Case 7&  ' IDIV
            dxv = getreg16(cpu, CPU_REG_EDX)
            axv = getreg16(cpu, CPU_REG_EAX)
            valdiv = U32Add(U32Shl(dxv, 16&), axv)
            op_idiv16 cpu, valdiv, cpu.oper1
    End Select
End Sub

Public Sub op_grp3_32(ByRef cpu As CPU_t)
    Dim imm32 As Long
    Dim codeAddr As Long
    Dim valdiv As U64_t
    Dim prod64 As U64_t
    Dim lhs As Long
    Dim rhs As Long
    Dim signFlag As Long

    Select Case (cpu.reg And &H7&)
        Case 0&, 1&  ' TEST
            codeAddr = U32Add(cpu.segcache(CPU_REG_CS), cpu.ip)
            imm32 = cpu_readl(cpu, codeAddr)
            flag_log32 cpu, (cpu.oper1_32 And imm32)
            cpu_stepIP cpu, 4&

        Case 2&  ' NOT
            cpu.res32 = (Not cpu.oper1_32)

        Case 3&  ' NEG
            cpu.res32 = U32Add((Not cpu.oper1_32), 1&)
            flag_sub32 cpu, 0&, cpu.oper1_32
            If cpu.res32 <> 0& Then
                cpu.cf = 1&
            Else
                cpu.cf = 0&
            End If

        Case 4&  ' MUL
            prod64 = cpu_u32MulToU64(cpu.oper1_32, cpu.regs_long(CPU_REG_EAX))
            cpu.regs_long(CPU_REG_EAX) = prod64.Lo
            cpu.regs_long(CPU_REG_EDX) = prod64.Hi

            If cpu.regs_long(CPU_REG_EDX) <> 0& Then
                cpu.cf = 1&
                cpu.ofl = 1&
            Else
                cpu.cf = 0&
                cpu.ofl = 0&
            End If

        Case 5&  ' IMUL
            prod64 = cpu_signedMul32ToU64(cpu.regs_long(CPU_REG_EAX), cpu.oper1_32)

            cpu.regs_long(CPU_REG_EAX) = prod64.Lo
            cpu.regs_long(CPU_REG_EDX) = prod64.Hi

            If cpu_u64SignExtMatches32(prod64) = 0& Then
                cpu.cf = 1&
                cpu.ofl = 1&
            Else
                cpu.cf = 0&
                cpu.ofl = 0&
            End If

        Case 6&  ' DIV
            valdiv = cpu_makeU64(cpu.regs_long(CPU_REG_EAX), cpu.regs_long(CPU_REG_EDX))
            op_div32 cpu, valdiv, cpu.oper1_32

        Case 7&  ' IDIV
            valdiv = cpu_makeU64(cpu.regs_long(CPU_REG_EAX), cpu.regs_long(CPU_REG_EDX))
            op_idiv32 cpu, valdiv, cpu.oper1_32
    End Select
End Sub

Private Function cpu_dispatchIntCallback(ByRef cpu As CPU_t, ByVal intnum As Long) As Long
    cpu_dispatchIntCallback = 0&
End Function

Public Sub cpu_registerIntCallback(ByRef cpu As CPU_t, ByVal interrupt As Long, ByVal cbid As Long)
    interrupt = (interrupt And &HFF&)
    cpu.int_callback(interrupt) = cbid
End Sub

Private Sub cpu_task_switch(ByRef cpu As CPU_t, ByVal new_tss_selector As Long, ByVal reason As Long)
    Dim new_desc_addr As Long
    Dim access As Long
    Dim typeVal As Long
    Dim present As Long
    Dim base As Long
    Dim limitVal As Long
    Dim old_tss_base As Long
    Dim old_access As Long
    Dim new_eflags As Long
    Dim is_task_gate As Long
    Dim is_task_return As Long
    Dim is_task_jump As Long
    Dim is_nested_switch As Long

    new_tss_selector = (new_tss_selector And &HFFFF&)
    If reason = TASK_SWITCH_REASON_GATE Then is_task_gate = 1& Else is_task_gate = 0&
    If reason = TASK_SWITCH_REASON_IRET Then is_task_return = 1& Else is_task_return = 0&
    If reason = TASK_SWITCH_REASON_JMP Then is_task_jump = 1& Else is_task_jump = 0&
    If (reason = TASK_SWITCH_REASON_CALL) Or (is_task_gate <> 0&) Then
        is_nested_switch = 1&
    Else
        is_nested_switch = 0&
    End If

    new_desc_addr = U32Add(cpu.gdtr, (new_tss_selector And &HFFF8&))
    access = (cpu_read_sys(cpu, U32Add(new_desc_addr, 5&)) And &HFF&)
    typeVal = (access And &HF&)
    present = (U32Shr(access, 7&) And &H1&)

    If (present = 0&) Or ((typeVal <> &H9&) And (typeVal <> &HB&)) Then
        cpu_raiseException cpu, 10&, new_tss_selector
        Exit Sub
    End If

    base = ((cpu_read_sys(cpu, U32Add(new_desc_addr, 2&)) And &HFF&) _
        Or U32Shl((cpu_read_sys(cpu, U32Add(new_desc_addr, 3&)) And &HFF&), 8&) _
        Or U32Shl((cpu_read_sys(cpu, U32Add(new_desc_addr, 4&)) And &HFF&), 16&) _
        Or U32Shl((cpu_read_sys(cpu, U32Add(new_desc_addr, 7&)) And &HFF&), 24&))

    limitVal = ((cpu_read_sys(cpu, new_desc_addr) And &HFF&) _
        Or U32Shl((cpu_read_sys(cpu, U32Add(new_desc_addr, 1&)) And &HFF&), 8&) _
        Or U32Shl((cpu_read_sys(cpu, U32Add(new_desc_addr, 6&)) And &HF&), 16&))

    If is_task_return = 0& Then
        old_tss_base = cpu.trbase

        cpu_writel_sys cpu, U32Add(old_tss_base, &H20&), cpu.ip
        cpu_writel_sys cpu, U32Add(old_tss_base, &H24&), makeflagsword(cpu)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H28&), cpu.regs_long(CPU_REG_EAX)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H2C&), cpu.regs_long(CPU_REG_ECX)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H30&), cpu.regs_long(CPU_REG_EDX)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H34&), cpu.regs_long(CPU_REG_EBX)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H38&), cpu.regs_long(CPU_REG_ESP)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H3C&), cpu.regs_long(CPU_REG_EBP)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H40&), cpu.regs_long(CPU_REG_ESI)
        cpu_writel_sys cpu, U32Add(old_tss_base, &H44&), cpu.regs_long(CPU_REG_EDI)

        cpu_writew_sys cpu, U32Add(old_tss_base, &H0&), cpu.segregs(CPU_REG_ES)
        cpu_writew_sys cpu, U32Add(old_tss_base, &H2&), cpu.segregs(CPU_REG_CS)
        cpu_writew_sys cpu, U32Add(old_tss_base, &H4&), cpu.segregs(CPU_REG_SS)
        cpu_writew_sys cpu, U32Add(old_tss_base, &H6&), cpu.segregs(CPU_REG_DS)
        cpu_writew_sys cpu, U32Add(old_tss_base, &H8&), cpu.segregs(CPU_REG_FS)
        cpu_writew_sys cpu, U32Add(old_tss_base, &HA&), cpu.segregs(CPU_REG_GS)
        cpu_writew_sys cpu, U32Add(old_tss_base, &H5C&), cpu.ldtr
        cpu_writel_sys cpu, U32Add(old_tss_base, &H1C&), cpu.CR(3&)

        old_access = (cpu_read_sys(cpu, U32Add(U32Add(cpu.gdtr, (cpu.tr_selector And &HFFF8&)), 5&)) And &HFF&)
        If (is_nested_switch <> 0&) And ((old_access And &HF&) = &H9&) Then
            cpu_write_sys cpu, U32Add(U32Add(cpu.gdtr, (cpu.tr_selector And &HFFF8&)), 5&), (old_access Or &H2&)
        ElseIf (is_task_jump <> 0&) And ((old_access And &HF&) = &HB&) Then
            cpu_write_sys cpu, U32Add(U32Add(cpu.gdtr, (cpu.tr_selector And &HFFF8&)), 5&), (old_access And Not &H2&)
        End If
    Else
        old_access = (cpu_read_sys(cpu, U32Add(U32Add(cpu.gdtr, (cpu.tr_selector And &HFFF8&)), 5&)) And &HFF&)
        If (old_access And &HF&) = &HB& Then
            cpu_write_sys cpu, U32Add(U32Add(cpu.gdtr, (cpu.tr_selector And &HFFF8&)), 5&), (old_access And Not &H2&)
        End If
    End If

    If (is_task_return = 0&) And (typeVal = &H9&) Then
        cpu_write_sys cpu, U32Add(new_desc_addr, 5&), (access Or &H2&)
    End If

    If is_nested_switch <> 0& Then
        cpu_writew_sys cpu, base, cpu.tr_selector
    End If

    cpu.ip = cpu_readl_sys(cpu, U32Add(base, &H20&))
    new_eflags = cpu_readl_sys(cpu, U32Add(base, &H24&))
    cpu.regs_long(CPU_REG_EAX) = cpu_readl_sys(cpu, U32Add(base, &H28&))
    cpu.regs_long(CPU_REG_ECX) = cpu_readl_sys(cpu, U32Add(base, &H2C&))
    cpu.regs_long(CPU_REG_EDX) = cpu_readl_sys(cpu, U32Add(base, &H30&))
    cpu.regs_long(CPU_REG_EBX) = cpu_readl_sys(cpu, U32Add(base, &H34&))
    cpu.regs_long(CPU_REG_ESP) = cpu_readl_sys(cpu, U32Add(base, &H38&))
    cpu.regs_long(CPU_REG_EBP) = cpu_readl_sys(cpu, U32Add(base, &H3C&))
    cpu.regs_long(CPU_REG_ESI) = cpu_readl_sys(cpu, U32Add(base, &H40&))
    cpu.regs_long(CPU_REG_EDI) = cpu_readl_sys(cpu, U32Add(base, &H44&))

    putsegreg cpu, CPU_REG_ES, cpu_readw_sys(cpu, U32Add(base, &H0&))
    putsegreg cpu, CPU_REG_CS, cpu_readw_sys(cpu, U32Add(base, &H2&))
    putsegreg cpu, CPU_REG_SS, cpu_readw_sys(cpu, U32Add(base, &H4&))
    putsegreg cpu, CPU_REG_DS, cpu_readw_sys(cpu, U32Add(base, &H6&))
    putsegreg cpu, CPU_REG_FS, cpu_readw_sys(cpu, U32Add(base, &H8&))
    putsegreg cpu, CPU_REG_GS, cpu_readw_sys(cpu, U32Add(base, &HA&))
    cpu.ldtr = (cpu_readw_sys(cpu, U32Add(base, &H5C&)) And &HFFFF&)

    cpu.CR(3&) = cpu_readl_sys(cpu, U32Add(base, &H1C&))
    memory_tlb_flush cpu

    cpu.tr_selector = new_tss_selector
    cpu.trbase = base

    decodeflagsword cpu, new_eflags

    If is_nested_switch <> 0& Then
        cpu.nt = 1&
    ElseIf (is_task_return <> 0&) Or (is_task_jump <> 0&) Then
        cpu.nt = 0&
    End If
End Sub

Public Sub cpu_intcall(ByRef cpu As CPU_t, ByVal intnum As Long, ByVal Source As Long, ByVal errCode As Long)
    Dim vecAddr As Long
    Dim target_cs As Long
    Dim target_ip As Long
    Dim gate As CPU_GATEDESC_t
    Dim target As CPU_CODETARGET_t
    Dim new_esp As Long
    Dim old_esp As Long
    Dim old_flags As Long
    Dim push_eip As Long
    Dim new_ss As Long
    Dim old_ss As Long
    Dim hasErr As Long
    Dim include_vm86 As Long
    Dim gate32 As Long
    Dim real_style_int As Long

    intnum = (intnum And &HFF&)

    real_style_int = 0&
    If cpu.protected_mode = 0& Then
        real_style_int = 1&
    ElseIf (cpu.v86f <> 0&) And _
           ((Source = INT_SOURCE_SOFTWARE) Or (Source = INT_SOURCE_INT3) Or (Source = INT_SOURCE_INTO)) And _
           (cpu.iopl >= 3&) Then
        real_style_int = 1&
    End If

    If real_style_int <> 0& Then
        If cpu.int_callback(intnum) <> CPU_INTCB_NONE Then
            If cpu_dispatchIntCallback(cpu, intnum) <> 0& Then
                Exit Sub
            End If
        End If

        vecAddr = (intnum * 4&)
        target_cs = (cpu_readw(cpu, vecAddr + 2&) And &HFFFF&)
        target_ip = (cpu_readw(cpu, vecAddr) And &HFFFF&)

        pushw cpu, cpu_makeflagsword(cpu)
        pushw cpu, cpu.segregs(CPU_REG_CS)
        pushw cpu, cpu.ip

        putsegreg cpu, CPU_REG_CS, target_cs
        cpu.ip = target_ip
        cpu.ifl = 0&
        cpu.tf = 0&
        Exit Sub
    End If

    If Source = INT_SOURCE_EXCEPTION Then
        push_eip = cpu.exceptionip
        cpu.nowrite = 0&
    Else
        push_eip = cpu.ip
    End If

    old_flags = makeflagsword(cpu)

    If (cpu.v86f <> 0&) And (Source = INT_SOURCE_SOFTWARE) And (cpu.iopl < 3&) Then
        cpu_raiseException cpu, 13&, 0&
        Exit Sub
    End If

    If cpu_read_idt_gate(cpu, intnum, gate) = 0& Then
        cpu_raiseException cpu, 13&, cpu_idt_error_code(intnum, Source)
        Exit Sub
    End If
    If (cpu.cpl > gate.dpl) And ((Source = INT_SOURCE_SOFTWARE) Or (Source = INT_SOURCE_INT3) Or (Source = INT_SOURCE_INTO)) Then
        cpu_raiseException cpu, 13&, cpu_idt_error_code(intnum, Source)
        Exit Sub
    End If
    If (cpu_gate_is_interrupt_or_trap(gate) = 0&) And (gate.typeVal <> &H5&) Then
        cpu_raiseException cpu, 13&, cpu_idt_error_code(intnum, Source)
        Exit Sub
    End If
    If gate.present = 0& Then
        cpu_raiseException cpu, 11&, cpu_idt_error_code(intnum, Source)
        Exit Sub
    End If

    If gate.typeVal = &H5& Then
        cpu_task_switch cpu, gate.target_selector, TASK_SWITCH_REASON_GATE
        Exit Sub
    End If

    If cpu_validate_gate_target_code(cpu, gate.target_selector, target) = 0& Then Exit Sub

    hasErr = 0&
    If (Source = INT_SOURCE_EXCEPTION) And (cpu_exception_has_error_code(intnum) <> 0&) Then
        hasErr = 1&
    End If

    gate32 = cpu_gate_is_32bit(gate)

    If (target.outer <> 0&) Or ((old_flags And EFLAGS_VM) <> 0&) Then
        If cpu_fetch_tss_stack(cpu, target.target_cpl, new_ss, new_esp) = 0& Then Exit Sub

        old_esp = cpu.regs_long(CPU_REG_ESP)
        old_ss = (cpu.segregs(CPU_REG_SS) And &HFFFF&)
        include_vm86 = 0&
        If ((old_flags And EFLAGS_VM) <> 0&) And (gate32 <> 0&) Then include_vm86 = 1&

        If (old_flags And EFLAGS_VM) <> 0& Then
            cpu.v86f = 0&
        End If

        putsegreg cpu, CPU_REG_SS, new_ss
        cpu.regs_long(CPU_REG_ESP) = new_esp

        If gate32 <> 0& Then
            If include_vm86 <> 0& Then
                cpu_stack_pushl_sys cpu, cpu.segregs(CPU_REG_GS)
                cpu_stack_pushl_sys cpu, cpu.segregs(CPU_REG_FS)
                cpu_stack_pushl_sys cpu, cpu.segregs(CPU_REG_DS)
                cpu_stack_pushl_sys cpu, cpu.segregs(CPU_REG_ES)
            End If

            cpu_stack_pushl_sys cpu, old_ss
            cpu_stack_pushl_sys cpu, old_esp
            cpu_stack_pushl_sys cpu, old_flags
            cpu_stack_pushl_sys cpu, cpu.segregs(CPU_REG_CS)
            cpu_stack_pushl_sys cpu, push_eip
            If hasErr <> 0& Then cpu_stack_pushl_sys cpu, errCode

            If cpu_gate_is_interrupt(gate) <> 0& Then cpu.ifl = 0&
            cpu.nt = 0&
            cpu.v86f = 0&
            cpu.tf = 0&

            putsegreg cpu, CPU_REG_CS, target.selector
            cpu.ip = gate.offset
            cpu.cpl = CByte(target.target_cpl And &H3&)
            Exit Sub
        End If

        cpu_push_far_return16_sys cpu, old_ss, old_esp
        cpu_push_far_return16_sys cpu, (old_flags And &HFFFF&), cpu.segregs(CPU_REG_CS)
        cpu_stack_pushw_sys cpu, push_eip
        If hasErr <> 0& Then cpu_stack_pushw_sys cpu, errCode

        If cpu_gate_is_interrupt(gate) <> 0& Then cpu.ifl = 0&
        cpu.nt = 0&
        cpu.iopl = 0&
        cpu.v86f = 0&
        cpu.tf = 0&

        putsegreg cpu, CPU_REG_CS, target.selector
        cpu.ip = (gate.offset And &HFFFF&)
        cpu.cpl = CByte(target.target_cpl And &H3&)
        Exit Sub
    End If

    If gate32 <> 0& Then
        pushl cpu, old_flags
        pushl cpu, (cpu.segregs(CPU_REG_CS) And &HFFFF&)
        pushl cpu, push_eip
        If hasErr <> 0& Then pushl cpu, errCode

        If cpu_gate_is_interrupt(gate) <> 0& Then cpu.ifl = 0&
        cpu.nt = 0&
        cpu.tf = 0&

        putsegreg cpu, CPU_REG_CS, target.selector
        cpu.cpl = CByte(target.target_cpl And &H3&)
        cpu.ip = gate.offset
        Exit Sub
    End If

    cpu_push_far_return16 cpu, (old_flags And &HFFFF&), cpu.segregs(CPU_REG_CS)
    pushw cpu, push_eip
    If hasErr <> 0& Then pushw cpu, errCode

    If cpu_gate_is_interrupt(gate) <> 0& Then cpu.ifl = 0&
    cpu.nt = 0&
    cpu.tf = 0&

    putsegreg cpu, CPU_REG_CS, target.selector
    cpu.ip = (gate.offset And &HFFFF&)
End Sub

Public Sub cpu_reset(ByRef cpu As CPU_t)
    Dim i As Long
    Static firstreset As Byte

    cpu_buildOpcodeMaps

    If firstreset = 0& Then
        For i = 0& To 255&
            cpu.int_callback(i) = 0&
        Next i
        firstreset = 1&
    End If

    cpu.usegdt = 0&
    cpu.protected_mode = 0&
    cpu.paging = 0&

    For i = 0& To 5&
        cpu.segis32(i) = 0&
    Next i

    cpu.a20_gate = 0&
    putsegreg cpu, CPU_REG_CS, &HFFFF&
    cpu.ip = 0&
    cpu.hltstate = 0&
    cpu.trap_toggle = 0&
    cpu.interrupt_inhibit = 0&
    cpu.totalexec_lo = 0&
    cpu.totalexec_hi = 0&
    cpu.have387 = 1&
    cpu.CR(0&) = (&H10& Or ((cpu.have387 Xor 1&) * CR0_EM))
    memory_tlb_flush cpu
    cpu.nowrite = 0&
    cpu.doexception = 0&
    cpu.exceptionerr = 0&

    fpu_reset
End Sub









