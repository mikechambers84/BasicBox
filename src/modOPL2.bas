Attribute VB_Name = "modOPL2"
Option Explicit

Private Const RSM_FRAC As Long = 10&
Private Const OPL_WRITEBUF_SIZE As Long = 1024&
Private Const OPL_WRITEBUF_DELAY As Long = 2&

Private Const CH_2OP As Long = 0&
Private Const CH_4OP As Long = 1&
Private Const CH_4OP2 As Long = 2&
Private Const CH_DRUM As Long = 3&

Private Const EGK_NORM As Long = &H1&
Private Const EGK_DRUM As Long = &H2&

Private Const ENVELOPE_GEN_ATTACK As Long = 0&
Private Const ENVELOPE_GEN_DECAY As Long = 1&
Private Const ENVELOPE_GEN_SUSTAIN As Long = 2&
Private Const ENVELOPE_GEN_RELEASE As Long = 3&

Private Const SRC_ZERO As Long = 0&
Private Const SRC_TREM As Long = 1&
Private Const SRC_SLOT_OUT_BASE As Long = 1000&
Private Const SRC_SLOT_FB_BASE As Long = 2000&

Private Type OPL3WriteBuf_t
    time As Double
    reg As Long
    data As Long
End Type

Private Type OPL3Slot_t
    channelIdx As Long
    outVal As Long
    fbmod As Long
    modSrc As Long
    prout As Long
    eg_rout As Long
    eg_out As Long
    eg_inc As Long
    eg_gen As Long
    eg_rate As Long
    eg_ksl As Long
    tremSrc As Long
    reg_vib As Long
    reg_type As Long
    reg_ksr As Long
    reg_mult As Long
    reg_ksl As Long
    reg_tl As Long
    reg_ar As Long
    reg_dr As Long
    reg_sl As Long
    reg_rr As Long
    reg_wf As Long
    key As Long
    pg_reset As Long
    pg_phase As Long
    pg_phase_out As Long
    slot_num As Long
End Type

Private Type OPL3Channel_t
    slotIdx(0& To 1&) As Long
    pairIdx As Long
    outSrc(0& To 3&) As Long
    chtype As Long
    f_num As Long
    block As Long
    fb As Long
    con As Long
    alg As Long
    ksv As Long
    cha As Long
    chb As Long
    ch_num As Long
End Type

Private Type OPL3Chip_t
    channel(0& To 17&) As OPL3Channel_t
    slot(0& To 35&) As OPL3Slot_t
    timer As Long
    eg_timer As Double
    eg_timerrem As Long
    eg_state As Long
    eg_add As Long
    newm As Long
    nts As Long
    rhy As Long
    vibpos As Long
    vibshift As Long
    tremolo As Long
    tremolopos As Long
    tremoloshift As Long
    noise As Long
    zeromod As Long
    mixbuff(0& To 1&) As Long
    rm_hh_bit2 As Long
    rm_hh_bit3 As Long
    rm_hh_bit7 As Long
    rm_hh_bit8 As Long
    rm_tc_bit3 As Long
    rm_tc_bit5 As Long
    rateratio As Long
    samplecnt As Long
    oldsamples(0& To 1&) As Long
    samples(0& To 1&) As Long
    writebuf_samplecnt As Double
    writebuf_cur As Long
    writebuf_last As Long
    writebuf_lasttime As Double
    writebuf(0& To OPL_WRITEBUF_SIZE - 1&) As OPL3WriteBuf_t
    data4 As Long
End Type

Private opl3_chip As OPL3Chip_t
Private opl3_portLatch As Long

Private opl3_logsinrom(0& To 255&) As Long
Private opl3_exprom(0& To 255&) As Long
Private opl3_mt(0& To 15&) As Long
Private opl3_kslrom(0& To 15&) As Long
Private opl3_kslshift(0& To 3&) As Long
Private opl3_eg_incstep(0& To 3&, 0& To 3&) As Long
Private opl3_ad_slot(0& To &H1F&) As Long
Private opl3_ch_slot(0& To 17&) As Long
Private opl3_tablesReady As Byte

Private Sub OPL3_InitTables()
    If opl3_tablesReady <> 0& Then Exit Sub

    ' logsinrom[256]
    opl3_logsinrom(0&) = &H859&
    opl3_logsinrom(1&) = &H6C3&
    opl3_logsinrom(2&) = &H607&
    opl3_logsinrom(3&) = &H58B&
    opl3_logsinrom(4&) = &H52E&
    opl3_logsinrom(5&) = &H4E4&
    opl3_logsinrom(6&) = &H4A6&
    opl3_logsinrom(7&) = &H471&
    opl3_logsinrom(8&) = &H443&
    opl3_logsinrom(9&) = &H41A&
    opl3_logsinrom(10&) = &H3F5&
    opl3_logsinrom(11&) = &H3D3&
    opl3_logsinrom(12&) = &H3B5&
    opl3_logsinrom(13&) = &H398&
    opl3_logsinrom(14&) = &H37E&
    opl3_logsinrom(15&) = &H365&
    opl3_logsinrom(16&) = &H34E&
    opl3_logsinrom(17&) = &H339&
    opl3_logsinrom(18&) = &H324&
    opl3_logsinrom(19&) = &H311&
    opl3_logsinrom(20&) = &H2FF&
    opl3_logsinrom(21&) = &H2ED&
    opl3_logsinrom(22&) = &H2DC&
    opl3_logsinrom(23&) = &H2CD&
    opl3_logsinrom(24&) = &H2BD&
    opl3_logsinrom(25&) = &H2AF&
    opl3_logsinrom(26&) = &H2A0&
    opl3_logsinrom(27&) = &H293&
    opl3_logsinrom(28&) = &H286&
    opl3_logsinrom(29&) = &H279&
    opl3_logsinrom(30&) = &H26D&
    opl3_logsinrom(31&) = &H261&
    opl3_logsinrom(32&) = &H256&
    opl3_logsinrom(33&) = &H24B&
    opl3_logsinrom(34&) = &H240&
    opl3_logsinrom(35&) = &H236&
    opl3_logsinrom(36&) = &H22C&
    opl3_logsinrom(37&) = &H222&
    opl3_logsinrom(38&) = &H218&
    opl3_logsinrom(39&) = &H20F&
    opl3_logsinrom(40&) = &H206&
    opl3_logsinrom(41&) = &H1FD&
    opl3_logsinrom(42&) = &H1F5&
    opl3_logsinrom(43&) = &H1EC&
    opl3_logsinrom(44&) = &H1E4&
    opl3_logsinrom(45&) = &H1DC&
    opl3_logsinrom(46&) = &H1D4&
    opl3_logsinrom(47&) = &H1CD&
    opl3_logsinrom(48&) = &H1C5&
    opl3_logsinrom(49&) = &H1BE&
    opl3_logsinrom(50&) = &H1B7&
    opl3_logsinrom(51&) = &H1B0&
    opl3_logsinrom(52&) = &H1A9&
    opl3_logsinrom(53&) = &H1A2&
    opl3_logsinrom(54&) = &H19B&
    opl3_logsinrom(55&) = &H195&
    opl3_logsinrom(56&) = &H18F&
    opl3_logsinrom(57&) = &H188&
    opl3_logsinrom(58&) = &H182&
    opl3_logsinrom(59&) = &H17C&
    opl3_logsinrom(60&) = &H177&
    opl3_logsinrom(61&) = &H171&
    opl3_logsinrom(62&) = &H16B&
    opl3_logsinrom(63&) = &H166&
    opl3_logsinrom(64&) = &H160&
    opl3_logsinrom(65&) = &H15B&
    opl3_logsinrom(66&) = &H155&
    opl3_logsinrom(67&) = &H150&
    opl3_logsinrom(68&) = &H14B&
    opl3_logsinrom(69&) = &H146&
    opl3_logsinrom(70&) = &H141&
    opl3_logsinrom(71&) = &H13C&
    opl3_logsinrom(72&) = &H137&
    opl3_logsinrom(73&) = &H133&
    opl3_logsinrom(74&) = &H12E&
    opl3_logsinrom(75&) = &H129&
    opl3_logsinrom(76&) = &H125&
    opl3_logsinrom(77&) = &H121&
    opl3_logsinrom(78&) = &H11C&
    opl3_logsinrom(79&) = &H118&
    opl3_logsinrom(80&) = &H114&
    opl3_logsinrom(81&) = &H10F&
    opl3_logsinrom(82&) = &H10B&
    opl3_logsinrom(83&) = &H107&
    opl3_logsinrom(84&) = &H103&
    opl3_logsinrom(85&) = &HFF&
    opl3_logsinrom(86&) = &HFB&
    opl3_logsinrom(87&) = &HF8&
    opl3_logsinrom(88&) = &HF4&
    opl3_logsinrom(89&) = &HF0&
    opl3_logsinrom(90&) = &HEC&
    opl3_logsinrom(91&) = &HE9&
    opl3_logsinrom(92&) = &HE5&
    opl3_logsinrom(93&) = &HE2&
    opl3_logsinrom(94&) = &HDE&
    opl3_logsinrom(95&) = &HDB&
    opl3_logsinrom(96&) = &HD7&
    opl3_logsinrom(97&) = &HD4&
    opl3_logsinrom(98&) = &HD1&
    opl3_logsinrom(99&) = &HCD&
    opl3_logsinrom(100&) = &HCA&
    opl3_logsinrom(101&) = &HC7&
    opl3_logsinrom(102&) = &HC4&
    opl3_logsinrom(103&) = &HC1&
    opl3_logsinrom(104&) = &HBE&
    opl3_logsinrom(105&) = &HBB&
    opl3_logsinrom(106&) = &HB8&
    opl3_logsinrom(107&) = &HB5&
    opl3_logsinrom(108&) = &HB2&
    opl3_logsinrom(109&) = &HAF&
    opl3_logsinrom(110&) = &HAC&
    opl3_logsinrom(111&) = &HA9&
    opl3_logsinrom(112&) = &HA7&
    opl3_logsinrom(113&) = &HA4&
    opl3_logsinrom(114&) = &HA1&
    opl3_logsinrom(115&) = &H9F&
    opl3_logsinrom(116&) = &H9C&
    opl3_logsinrom(117&) = &H99&
    opl3_logsinrom(118&) = &H97&
    opl3_logsinrom(119&) = &H94&
    opl3_logsinrom(120&) = &H92&
    opl3_logsinrom(121&) = &H8F&
    opl3_logsinrom(122&) = &H8D&
    opl3_logsinrom(123&) = &H8A&
    opl3_logsinrom(124&) = &H88&
    opl3_logsinrom(125&) = &H86&
    opl3_logsinrom(126&) = &H83&
    opl3_logsinrom(127&) = &H81&
    opl3_logsinrom(128&) = &H7F&
    opl3_logsinrom(129&) = &H7D&
    opl3_logsinrom(130&) = &H7A&
    opl3_logsinrom(131&) = &H78&
    opl3_logsinrom(132&) = &H76&
    opl3_logsinrom(133&) = &H74&
    opl3_logsinrom(134&) = &H72&
    opl3_logsinrom(135&) = &H70&
    opl3_logsinrom(136&) = &H6E&
    opl3_logsinrom(137&) = &H6C&
    opl3_logsinrom(138&) = &H6A&
    opl3_logsinrom(139&) = &H68&
    opl3_logsinrom(140&) = &H66&
    opl3_logsinrom(141&) = &H64&
    opl3_logsinrom(142&) = &H62&
    opl3_logsinrom(143&) = &H60&
    opl3_logsinrom(144&) = &H5E&
    opl3_logsinrom(145&) = &H5C&
    opl3_logsinrom(146&) = &H5B&
    opl3_logsinrom(147&) = &H59&
    opl3_logsinrom(148&) = &H57&
    opl3_logsinrom(149&) = &H55&
    opl3_logsinrom(150&) = &H53&
    opl3_logsinrom(151&) = &H52&
    opl3_logsinrom(152&) = &H50&
    opl3_logsinrom(153&) = &H4E&
    opl3_logsinrom(154&) = &H4D&
    opl3_logsinrom(155&) = &H4B&
    opl3_logsinrom(156&) = &H4A&
    opl3_logsinrom(157&) = &H48&
    opl3_logsinrom(158&) = &H46&
    opl3_logsinrom(159&) = &H45&
    opl3_logsinrom(160&) = &H43&
    opl3_logsinrom(161&) = &H42&
    opl3_logsinrom(162&) = &H40&
    opl3_logsinrom(163&) = &H3F&
    opl3_logsinrom(164&) = &H3E&
    opl3_logsinrom(165&) = &H3C&
    opl3_logsinrom(166&) = &H3B&
    opl3_logsinrom(167&) = &H39&
    opl3_logsinrom(168&) = &H38&
    opl3_logsinrom(169&) = &H37&
    opl3_logsinrom(170&) = &H35&
    opl3_logsinrom(171&) = &H34&
    opl3_logsinrom(172&) = &H33&
    opl3_logsinrom(173&) = &H31&
    opl3_logsinrom(174&) = &H30&
    opl3_logsinrom(175&) = &H2F&
    opl3_logsinrom(176&) = &H2E&
    opl3_logsinrom(177&) = &H2D&
    opl3_logsinrom(178&) = &H2B&
    opl3_logsinrom(179&) = &H2A&
    opl3_logsinrom(180&) = &H29&
    opl3_logsinrom(181&) = &H28&
    opl3_logsinrom(182&) = &H27&
    opl3_logsinrom(183&) = &H26&
    opl3_logsinrom(184&) = &H25&
    opl3_logsinrom(185&) = &H24&
    opl3_logsinrom(186&) = &H23&
    opl3_logsinrom(187&) = &H22&
    opl3_logsinrom(188&) = &H21&
    opl3_logsinrom(189&) = &H20&
    opl3_logsinrom(190&) = &H1F&
    opl3_logsinrom(191&) = &H1E&
    opl3_logsinrom(192&) = &H1D&
    opl3_logsinrom(193&) = &H1C&
    opl3_logsinrom(194&) = &H1B&
    opl3_logsinrom(195&) = &H1A&
    opl3_logsinrom(196&) = &H19&
    opl3_logsinrom(197&) = &H18&
    opl3_logsinrom(198&) = &H17&
    opl3_logsinrom(199&) = &H17&
    opl3_logsinrom(200&) = &H16&
    opl3_logsinrom(201&) = &H15&
    opl3_logsinrom(202&) = &H14&
    opl3_logsinrom(203&) = &H14&
    opl3_logsinrom(204&) = &H13&
    opl3_logsinrom(205&) = &H12&
    opl3_logsinrom(206&) = &H11&
    opl3_logsinrom(207&) = &H11&
    opl3_logsinrom(208&) = &H10&
    opl3_logsinrom(209&) = &HF&
    opl3_logsinrom(210&) = &HF&
    opl3_logsinrom(211&) = &HE&
    opl3_logsinrom(212&) = &HD&
    opl3_logsinrom(213&) = &HD&
    opl3_logsinrom(214&) = &HC&
    opl3_logsinrom(215&) = &HC&
    opl3_logsinrom(216&) = &HB&
    opl3_logsinrom(217&) = &HA&
    opl3_logsinrom(218&) = &HA&
    opl3_logsinrom(219&) = &H9&
    opl3_logsinrom(220&) = &H9&
    opl3_logsinrom(221&) = &H8&
    opl3_logsinrom(222&) = &H8&
    opl3_logsinrom(223&) = &H7&
    opl3_logsinrom(224&) = &H7&
    opl3_logsinrom(225&) = &H7&
    opl3_logsinrom(226&) = &H6&
    opl3_logsinrom(227&) = &H6&
    opl3_logsinrom(228&) = &H5&
    opl3_logsinrom(229&) = &H5&
    opl3_logsinrom(230&) = &H5&
    opl3_logsinrom(231&) = &H4&
    opl3_logsinrom(232&) = &H4&
    opl3_logsinrom(233&) = &H4&
    opl3_logsinrom(234&) = &H3&
    opl3_logsinrom(235&) = &H3&
    opl3_logsinrom(236&) = &H3&
    opl3_logsinrom(237&) = &H2&
    opl3_logsinrom(238&) = &H2&
    opl3_logsinrom(239&) = &H2&
    opl3_logsinrom(240&) = &H2&
    opl3_logsinrom(241&) = &H1&
    opl3_logsinrom(242&) = &H1&
    opl3_logsinrom(243&) = &H1&
    opl3_logsinrom(244&) = &H1&
    opl3_logsinrom(245&) = &H1&
    opl3_logsinrom(246&) = &H1&
    opl3_logsinrom(247&) = &H1&
    opl3_logsinrom(248&) = &H0&
    opl3_logsinrom(249&) = &H0&
    opl3_logsinrom(250&) = &H0&
    opl3_logsinrom(251&) = &H0&
    opl3_logsinrom(252&) = &H0&
    opl3_logsinrom(253&) = &H0&
    opl3_logsinrom(254&) = &H0&
    opl3_logsinrom(255&) = &H0&

    ' exprom[256]
    opl3_exprom(0&) = &H7FA&
    opl3_exprom(1&) = &H7F5&
    opl3_exprom(2&) = &H7EF&
    opl3_exprom(3&) = &H7EA&
    opl3_exprom(4&) = &H7E4&
    opl3_exprom(5&) = &H7DF&
    opl3_exprom(6&) = &H7DA&
    opl3_exprom(7&) = &H7D4&
    opl3_exprom(8&) = &H7CF&
    opl3_exprom(9&) = &H7C9&
    opl3_exprom(10&) = &H7C4&
    opl3_exprom(11&) = &H7BF&
    opl3_exprom(12&) = &H7B9&
    opl3_exprom(13&) = &H7B4&
    opl3_exprom(14&) = &H7AE&
    opl3_exprom(15&) = &H7A9&
    opl3_exprom(16&) = &H7A4&
    opl3_exprom(17&) = &H79F&
    opl3_exprom(18&) = &H799&
    opl3_exprom(19&) = &H794&
    opl3_exprom(20&) = &H78F&
    opl3_exprom(21&) = &H78A&
    opl3_exprom(22&) = &H784&
    opl3_exprom(23&) = &H77F&
    opl3_exprom(24&) = &H77A&
    opl3_exprom(25&) = &H775&
    opl3_exprom(26&) = &H770&
    opl3_exprom(27&) = &H76A&
    opl3_exprom(28&) = &H765&
    opl3_exprom(29&) = &H760&
    opl3_exprom(30&) = &H75B&
    opl3_exprom(31&) = &H756&
    opl3_exprom(32&) = &H751&
    opl3_exprom(33&) = &H74C&
    opl3_exprom(34&) = &H747&
    opl3_exprom(35&) = &H742&
    opl3_exprom(36&) = &H73D&
    opl3_exprom(37&) = &H738&
    opl3_exprom(38&) = &H733&
    opl3_exprom(39&) = &H72E&
    opl3_exprom(40&) = &H729&
    opl3_exprom(41&) = &H724&
    opl3_exprom(42&) = &H71F&
    opl3_exprom(43&) = &H71A&
    opl3_exprom(44&) = &H715&
    opl3_exprom(45&) = &H710&
    opl3_exprom(46&) = &H70B&
    opl3_exprom(47&) = &H706&
    opl3_exprom(48&) = &H702&
    opl3_exprom(49&) = &H6FD&
    opl3_exprom(50&) = &H6F8&
    opl3_exprom(51&) = &H6F3&
    opl3_exprom(52&) = &H6EE&
    opl3_exprom(53&) = &H6E9&
    opl3_exprom(54&) = &H6E5&
    opl3_exprom(55&) = &H6E0&
    opl3_exprom(56&) = &H6DB&
    opl3_exprom(57&) = &H6D6&
    opl3_exprom(58&) = &H6D2&
    opl3_exprom(59&) = &H6CD&
    opl3_exprom(60&) = &H6C8&
    opl3_exprom(61&) = &H6C4&
    opl3_exprom(62&) = &H6BF&
    opl3_exprom(63&) = &H6BA&
    opl3_exprom(64&) = &H6B5&
    opl3_exprom(65&) = &H6B1&
    opl3_exprom(66&) = &H6AC&
    opl3_exprom(67&) = &H6A8&
    opl3_exprom(68&) = &H6A3&
    opl3_exprom(69&) = &H69E&
    opl3_exprom(70&) = &H69A&
    opl3_exprom(71&) = &H695&
    opl3_exprom(72&) = &H691&
    opl3_exprom(73&) = &H68C&
    opl3_exprom(74&) = &H688&
    opl3_exprom(75&) = &H683&
    opl3_exprom(76&) = &H67F&
    opl3_exprom(77&) = &H67A&
    opl3_exprom(78&) = &H676&
    opl3_exprom(79&) = &H671&
    opl3_exprom(80&) = &H66D&
    opl3_exprom(81&) = &H668&
    opl3_exprom(82&) = &H664&
    opl3_exprom(83&) = &H65F&
    opl3_exprom(84&) = &H65B&
    opl3_exprom(85&) = &H657&
    opl3_exprom(86&) = &H652&
    opl3_exprom(87&) = &H64E&
    opl3_exprom(88&) = &H649&
    opl3_exprom(89&) = &H645&
    opl3_exprom(90&) = &H641&
    opl3_exprom(91&) = &H63C&
    opl3_exprom(92&) = &H638&
    opl3_exprom(93&) = &H634&
    opl3_exprom(94&) = &H630&
    opl3_exprom(95&) = &H62B&
    opl3_exprom(96&) = &H627&
    opl3_exprom(97&) = &H623&
    opl3_exprom(98&) = &H61E&
    opl3_exprom(99&) = &H61A&
    opl3_exprom(100&) = &H616&
    opl3_exprom(101&) = &H612&
    opl3_exprom(102&) = &H60E&
    opl3_exprom(103&) = &H609&
    opl3_exprom(104&) = &H605&
    opl3_exprom(105&) = &H601&
    opl3_exprom(106&) = &H5FD&
    opl3_exprom(107&) = &H5F9&
    opl3_exprom(108&) = &H5F5&
    opl3_exprom(109&) = &H5F0&
    opl3_exprom(110&) = &H5EC&
    opl3_exprom(111&) = &H5E8&
    opl3_exprom(112&) = &H5E4&
    opl3_exprom(113&) = &H5E0&
    opl3_exprom(114&) = &H5DC&
    opl3_exprom(115&) = &H5D8&
    opl3_exprom(116&) = &H5D4&
    opl3_exprom(117&) = &H5D0&
    opl3_exprom(118&) = &H5CC&
    opl3_exprom(119&) = &H5C8&
    opl3_exprom(120&) = &H5C4&
    opl3_exprom(121&) = &H5C0&
    opl3_exprom(122&) = &H5BC&
    opl3_exprom(123&) = &H5B8&
    opl3_exprom(124&) = &H5B4&
    opl3_exprom(125&) = &H5B0&
    opl3_exprom(126&) = &H5AC&
    opl3_exprom(127&) = &H5A8&
    opl3_exprom(128&) = &H5A4&
    opl3_exprom(129&) = &H5A0&
    opl3_exprom(130&) = &H59C&
    opl3_exprom(131&) = &H599&
    opl3_exprom(132&) = &H595&
    opl3_exprom(133&) = &H591&
    opl3_exprom(134&) = &H58D&
    opl3_exprom(135&) = &H589&
    opl3_exprom(136&) = &H585&
    opl3_exprom(137&) = &H581&
    opl3_exprom(138&) = &H57E&
    opl3_exprom(139&) = &H57A&
    opl3_exprom(140&) = &H576&
    opl3_exprom(141&) = &H572&
    opl3_exprom(142&) = &H56F&
    opl3_exprom(143&) = &H56B&
    opl3_exprom(144&) = &H567&
    opl3_exprom(145&) = &H563&
    opl3_exprom(146&) = &H560&
    opl3_exprom(147&) = &H55C&
    opl3_exprom(148&) = &H558&
    opl3_exprom(149&) = &H554&
    opl3_exprom(150&) = &H551&
    opl3_exprom(151&) = &H54D&
    opl3_exprom(152&) = &H549&
    opl3_exprom(153&) = &H546&
    opl3_exprom(154&) = &H542&
    opl3_exprom(155&) = &H53E&
    opl3_exprom(156&) = &H53B&
    opl3_exprom(157&) = &H537&
    opl3_exprom(158&) = &H534&
    opl3_exprom(159&) = &H530&
    opl3_exprom(160&) = &H52C&
    opl3_exprom(161&) = &H529&
    opl3_exprom(162&) = &H525&
    opl3_exprom(163&) = &H522&
    opl3_exprom(164&) = &H51E&
    opl3_exprom(165&) = &H51B&
    opl3_exprom(166&) = &H517&
    opl3_exprom(167&) = &H514&
    opl3_exprom(168&) = &H510&
    opl3_exprom(169&) = &H50C&
    opl3_exprom(170&) = &H509&
    opl3_exprom(171&) = &H506&
    opl3_exprom(172&) = &H502&
    opl3_exprom(173&) = &H4FF&
    opl3_exprom(174&) = &H4FB&
    opl3_exprom(175&) = &H4F8&
    opl3_exprom(176&) = &H4F4&
    opl3_exprom(177&) = &H4F1&
    opl3_exprom(178&) = &H4ED&
    opl3_exprom(179&) = &H4EA&
    opl3_exprom(180&) = &H4E7&
    opl3_exprom(181&) = &H4E3&
    opl3_exprom(182&) = &H4E0&
    opl3_exprom(183&) = &H4DC&
    opl3_exprom(184&) = &H4D9&
    opl3_exprom(185&) = &H4D6&
    opl3_exprom(186&) = &H4D2&
    opl3_exprom(187&) = &H4CF&
    opl3_exprom(188&) = &H4CC&
    opl3_exprom(189&) = &H4C8&
    opl3_exprom(190&) = &H4C5&
    opl3_exprom(191&) = &H4C2&
    opl3_exprom(192&) = &H4BE&
    opl3_exprom(193&) = &H4BB&
    opl3_exprom(194&) = &H4B8&
    opl3_exprom(195&) = &H4B5&
    opl3_exprom(196&) = &H4B1&
    opl3_exprom(197&) = &H4AE&
    opl3_exprom(198&) = &H4AB&
    opl3_exprom(199&) = &H4A8&
    opl3_exprom(200&) = &H4A4&
    opl3_exprom(201&) = &H4A1&
    opl3_exprom(202&) = &H49E&
    opl3_exprom(203&) = &H49B&
    opl3_exprom(204&) = &H498&
    opl3_exprom(205&) = &H494&
    opl3_exprom(206&) = &H491&
    opl3_exprom(207&) = &H48E&
    opl3_exprom(208&) = &H48B&
    opl3_exprom(209&) = &H488&
    opl3_exprom(210&) = &H485&
    opl3_exprom(211&) = &H482&
    opl3_exprom(212&) = &H47E&
    opl3_exprom(213&) = &H47B&
    opl3_exprom(214&) = &H478&
    opl3_exprom(215&) = &H475&
    opl3_exprom(216&) = &H472&
    opl3_exprom(217&) = &H46F&
    opl3_exprom(218&) = &H46C&
    opl3_exprom(219&) = &H469&
    opl3_exprom(220&) = &H466&
    opl3_exprom(221&) = &H463&
    opl3_exprom(222&) = &H460&
    opl3_exprom(223&) = &H45D&
    opl3_exprom(224&) = &H45A&
    opl3_exprom(225&) = &H457&
    opl3_exprom(226&) = &H454&
    opl3_exprom(227&) = &H451&
    opl3_exprom(228&) = &H44E&
    opl3_exprom(229&) = &H44B&
    opl3_exprom(230&) = &H448&
    opl3_exprom(231&) = &H445&
    opl3_exprom(232&) = &H442&
    opl3_exprom(233&) = &H43F&
    opl3_exprom(234&) = &H43C&
    opl3_exprom(235&) = &H439&
    opl3_exprom(236&) = &H436&
    opl3_exprom(237&) = &H433&
    opl3_exprom(238&) = &H430&
    opl3_exprom(239&) = &H42D&
    opl3_exprom(240&) = &H42A&
    opl3_exprom(241&) = &H428&
    opl3_exprom(242&) = &H425&
    opl3_exprom(243&) = &H422&
    opl3_exprom(244&) = &H41F&
    opl3_exprom(245&) = &H41C&
    opl3_exprom(246&) = &H419&
    opl3_exprom(247&) = &H416&
    opl3_exprom(248&) = &H414&
    opl3_exprom(249&) = &H411&
    opl3_exprom(250&) = &H40E&
    opl3_exprom(251&) = &H40B&
    opl3_exprom(252&) = &H408&
    opl3_exprom(253&) = &H406&
    opl3_exprom(254&) = &H403&
    opl3_exprom(255&) = &H400&

    ' mt[16]
    opl3_mt(0&) = 1&
    opl3_mt(1&) = 2&
    opl3_mt(2&) = 4&
    opl3_mt(3&) = 6&
    opl3_mt(4&) = 8&
    opl3_mt(5&) = 10&
    opl3_mt(6&) = 12&
    opl3_mt(7&) = 14&
    opl3_mt(8&) = 16&
    opl3_mt(9&) = 18&
    opl3_mt(10&) = 20&
    opl3_mt(11&) = 20&
    opl3_mt(12&) = 24&
    opl3_mt(13&) = 24&
    opl3_mt(14&) = 30&
    opl3_mt(15&) = 30&

    ' kslrom[16]
    opl3_kslrom(0&) = 0&
    opl3_kslrom(1&) = 32&
    opl3_kslrom(2&) = 40&
    opl3_kslrom(3&) = 45&
    opl3_kslrom(4&) = 48&
    opl3_kslrom(5&) = 51&
    opl3_kslrom(6&) = 53&
    opl3_kslrom(7&) = 55&
    opl3_kslrom(8&) = 56&
    opl3_kslrom(9&) = 58&
    opl3_kslrom(10&) = 59&
    opl3_kslrom(11&) = 60&
    opl3_kslrom(12&) = 61&
    opl3_kslrom(13&) = 62&
    opl3_kslrom(14&) = 63&
    opl3_kslrom(15&) = 64&

    ' kslshift[4]
    opl3_kslshift(0&) = 8&
    opl3_kslshift(1&) = 1&
    opl3_kslshift(2&) = 2&
    opl3_kslshift(3&) = 0&

    ' eg_incstep[4][4]
    opl3_eg_incstep(0&, 0&) = 0&
    opl3_eg_incstep(0&, 1&) = 0&
    opl3_eg_incstep(0&, 2&) = 0&
    opl3_eg_incstep(0&, 3&) = 0&
    opl3_eg_incstep(1&, 0&) = 1&
    opl3_eg_incstep(1&, 1&) = 0&
    opl3_eg_incstep(1&, 2&) = 0&
    opl3_eg_incstep(1&, 3&) = 0&
    opl3_eg_incstep(2&, 0&) = 1&
    opl3_eg_incstep(2&, 1&) = 0&
    opl3_eg_incstep(2&, 2&) = 1&
    opl3_eg_incstep(2&, 3&) = 0&
    opl3_eg_incstep(3&, 0&) = 1&
    opl3_eg_incstep(3&, 1&) = 1&
    opl3_eg_incstep(3&, 2&) = 1&
    opl3_eg_incstep(3&, 3&) = 0&

    ' ad_slot[32]
    opl3_ad_slot(0&) = 0&
    opl3_ad_slot(1&) = 1&
    opl3_ad_slot(2&) = 2&
    opl3_ad_slot(3&) = 3&
    opl3_ad_slot(4&) = 4&
    opl3_ad_slot(5&) = 5&
    opl3_ad_slot(6&) = -1&
    opl3_ad_slot(7&) = -1&
    opl3_ad_slot(8&) = 6&
    opl3_ad_slot(9&) = 7&
    opl3_ad_slot(10&) = 8&
    opl3_ad_slot(11&) = 9&
    opl3_ad_slot(12&) = 10&
    opl3_ad_slot(13&) = 11&
    opl3_ad_slot(14&) = -1&
    opl3_ad_slot(15&) = -1&
    opl3_ad_slot(16&) = 12&
    opl3_ad_slot(17&) = 13&
    opl3_ad_slot(18&) = 14&
    opl3_ad_slot(19&) = 15&
    opl3_ad_slot(20&) = 16&
    opl3_ad_slot(21&) = 17&
    opl3_ad_slot(22&) = -1&
    opl3_ad_slot(23&) = -1&
    opl3_ad_slot(24&) = -1&
    opl3_ad_slot(25&) = -1&
    opl3_ad_slot(26&) = -1&
    opl3_ad_slot(27&) = -1&
    opl3_ad_slot(28&) = -1&
    opl3_ad_slot(29&) = -1&
    opl3_ad_slot(30&) = -1&
    opl3_ad_slot(31&) = -1&

    ' ch_slot[18]
    opl3_ch_slot(0&) = 0&
    opl3_ch_slot(1&) = 1&
    opl3_ch_slot(2&) = 2&
    opl3_ch_slot(3&) = 6&
    opl3_ch_slot(4&) = 7&
    opl3_ch_slot(5&) = 8&
    opl3_ch_slot(6&) = 12&
    opl3_ch_slot(7&) = 13&
    opl3_ch_slot(8&) = 14&
    opl3_ch_slot(9&) = 18&
    opl3_ch_slot(10&) = 19&
    opl3_ch_slot(11&) = 20&
    opl3_ch_slot(12&) = 24&
    opl3_ch_slot(13&) = 25&
    opl3_ch_slot(14&) = 26&
    opl3_ch_slot(15&) = 30&
    opl3_ch_slot(16&) = 31&
    opl3_ch_slot(17&) = 32&

    opl3_tablesReady = 1&
End Sub

Private Function OPL3_SrcSlotOut(ByVal slotIdx As Long) As Long
    OPL3_SrcSlotOut = SRC_SLOT_OUT_BASE + slotIdx
End Function

Private Function OPL3_SrcSlotFb(ByVal slotIdx As Long) As Long
    OPL3_SrcSlotFb = SRC_SLOT_FB_BASE + slotIdx
End Function

Private Function OPL3_ReadSrc(ByVal src As Long) As Long
    If src = SRC_ZERO Then
        OPL3_ReadSrc = opl3_chip.zeromod
    ElseIf src = SRC_TREM Then
        OPL3_ReadSrc = opl3_chip.tremolo
    ElseIf src >= SRC_SLOT_FB_BASE Then
        OPL3_ReadSrc = opl3_chip.slot(src - SRC_SLOT_FB_BASE).fbmod
    ElseIf src >= SRC_SLOT_OUT_BASE Then
        OPL3_ReadSrc = opl3_chip.slot(src - SRC_SLOT_OUT_BASE).outVal
    Else
        OPL3_ReadSrc = 0&
    End If
End Function

Private Function OPL3_S16(ByVal v As Long) As Long
    v = v And &HFFFF&
    If (v And &H8000&) <> 0& Then
        OPL3_S16 = v - &H10000
    Else
        OPL3_S16 = v
    End If
End Function

Private Function OPL3_TestBit64(ByVal value As Double, ByVal bit As Long) As Long
    Dim divisor As Double
    Dim q As Double

    divisor = 2# ^ CDbl(bit)
    q = Fix(value / divisor)
    OPL3_TestBit64 = CLng(q - (2# * Fix(q / 2#)))
End Function

Private Function OPL3_ClipSample(ByVal sample As Long) As Integer
    If sample > 32767& Then
        sample = 32767&
    ElseIf sample < -32768 Then
        sample = -32768
    End If
    OPL3_ClipSample = CInt(sample)
End Function

Private Function OPL3_EnvelopeCalcExp(ByVal level As Long) As Long
    Dim tmp As Long

    If level > &H1FFF& Then
        level = &H1FFF&
    End If

    tmp = U32Shl(opl3_exprom(level And &HFF&), 1&)
    OPL3_EnvelopeCalcExp = OPL3_S16(U32Shr(tmp, U32Shr(level, 8&)))
End Function

Private Function OPL3_EnvelopeCalcSin(ByVal wf As Long, ByVal phase As Long, ByVal envelope As Long) As Long
    Dim outv As Long
    Dim neg As Long

    outv = 0&
    neg = 0&
    phase = phase And &H3FF&

    Select Case wf And 7&
        Case 0&
            If (phase And &H200&) <> 0& Then neg = &HFFFF&
            If (phase And &H100&) <> 0& Then
                outv = opl3_logsinrom((phase And &HFF&) Xor &HFF&)
            Else
                outv = opl3_logsinrom(phase And &HFF&)
            End If
            OPL3_EnvelopeCalcSin = OPL3_S16(OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&)) Xor neg)

        Case 1&
            If (phase And &H200&) <> 0& Then
                outv = &H1000&
            ElseIf (phase And &H100&) <> 0& Then
                outv = opl3_logsinrom((phase And &HFF&) Xor &HFF&)
            Else
                outv = opl3_logsinrom(phase And &HFF&)
            End If
            OPL3_EnvelopeCalcSin = OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&))

        Case 2&
            If (phase And &H100&) <> 0& Then
                outv = opl3_logsinrom((phase And &HFF&) Xor &HFF&)
            Else
                outv = opl3_logsinrom(phase And &HFF&)
            End If
            OPL3_EnvelopeCalcSin = OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&))

        Case 3&
            If (phase And &H100&) <> 0& Then
                outv = &H1000&
            Else
                outv = opl3_logsinrom(phase And &HFF&)
            End If
            OPL3_EnvelopeCalcSin = OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&))

        Case 4&
            If (phase And &H300&) = &H100& Then neg = &HFFFF&
            If (phase And &H200&) <> 0& Then
                outv = &H1000&
            ElseIf (phase And &H80&) <> 0& Then
                outv = opl3_logsinrom((U32Shl((phase Xor &HFF&), 1&) And &HFF&))
            Else
                outv = opl3_logsinrom((U32Shl(phase, 1&) And &HFF&))
            End If
            OPL3_EnvelopeCalcSin = OPL3_S16(OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&)) Xor neg)

        Case 5&
            If (phase And &H200&) <> 0& Then
                outv = &H1000&
            ElseIf (phase And &H80&) <> 0& Then
                outv = opl3_logsinrom((U32Shl((phase Xor &HFF&), 1&) And &HFF&))
            Else
                outv = opl3_logsinrom((U32Shl(phase, 1&) And &HFF&))
            End If
            OPL3_EnvelopeCalcSin = OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&))

        Case 6&
            If (phase And &H200&) <> 0& Then neg = &HFFFF&
            OPL3_EnvelopeCalcSin = OPL3_S16(OPL3_EnvelopeCalcExp(U32Shl(envelope, 3&)) Xor neg)

        Case Else
            If (phase And &H200&) <> 0& Then
                neg = &HFFFF&
                phase = (phase And &H1FF&) Xor &H1FF&
            End If
            outv = U32Shl(phase, 3&)
            OPL3_EnvelopeCalcSin = OPL3_S16(OPL3_EnvelopeCalcExp(outv + U32Shl(envelope, 3&)) Xor neg)
    End Select
End Function

Private Sub OPL3_EnvelopeUpdateKSL(ByVal slotIdx As Long)
    Dim chIdx As Long
    Dim ksl As Long

    chIdx = opl3_chip.slot(slotIdx).channelIdx
    ksl = U32Shl(opl3_kslrom(U32Shr(opl3_chip.channel(chIdx).f_num, 6&)), 2&) - U32Shl((8& - opl3_chip.channel(chIdx).block), 5&)
    If ksl < 0& Then ksl = 0&
    opl3_chip.slot(slotIdx).eg_ksl = ksl And &HFF&
End Sub

Private Sub OPL3_EnvelopeCalc(ByVal slotIdx As Long)
    Dim nonzero As Long
    Dim rate As Long
    Dim rate_hi As Long
    Dim rate_lo As Long
    Dim reg_rate As Long
    Dim ks As Long
    Dim eg_shift As Long
    Dim shift As Long
    Dim eg_rout As Long
    Dim eg_inc As Long
    Dim eg_off As Long
    Dim reset As Long
    Dim chIdx As Long

    chIdx = opl3_chip.slot(slotIdx).channelIdx

    reg_rate = 0&
    reset = 0&

    opl3_chip.slot(slotIdx).eg_out = (opl3_chip.slot(slotIdx).eg_rout + U32Shl(opl3_chip.slot(slotIdx).reg_tl, 2&) + U32Shr(opl3_chip.slot(slotIdx).eg_ksl, opl3_kslshift(opl3_chip.slot(slotIdx).reg_ksl)) + OPL3_ReadSrc(opl3_chip.slot(slotIdx).tremSrc)) And &HFFFF&

    If (opl3_chip.slot(slotIdx).key <> 0&) And (opl3_chip.slot(slotIdx).eg_gen = ENVELOPE_GEN_RELEASE) Then
        reset = 1&
        reg_rate = opl3_chip.slot(slotIdx).reg_ar
    Else
        Select Case opl3_chip.slot(slotIdx).eg_gen
            Case ENVELOPE_GEN_ATTACK
                reg_rate = opl3_chip.slot(slotIdx).reg_ar
            Case ENVELOPE_GEN_DECAY
                reg_rate = opl3_chip.slot(slotIdx).reg_dr
            Case ENVELOPE_GEN_SUSTAIN
                If opl3_chip.slot(slotIdx).reg_type = 0& Then reg_rate = opl3_chip.slot(slotIdx).reg_rr
            Case ENVELOPE_GEN_RELEASE
                reg_rate = opl3_chip.slot(slotIdx).reg_rr
        End Select
    End If

    opl3_chip.slot(slotIdx).pg_reset = reset

    ks = U32Shr(opl3_chip.channel(chIdx).ksv, U32Shl((opl3_chip.slot(slotIdx).reg_ksr Xor 1&), 1&))
    nonzero = IIf(reg_rate <> 0&, 1&, 0&)
    rate = ks + U32Shl(reg_rate, 2&)
    rate_hi = U32Shr(rate, 2&)
    rate_lo = rate And 3&

    If (rate_hi And &H10&) <> 0& Then rate_hi = &HF&

    eg_shift = rate_hi + opl3_chip.eg_add
    shift = 0&

    If nonzero <> 0& Then
        If rate_hi < 12& Then
            If opl3_chip.eg_state <> 0& Then
                Select Case eg_shift
                    Case 12&
                        shift = 1&
                    Case 13&
                        shift = U32Shr(rate_lo, 1&) And 1&
                    Case 14&
                        shift = rate_lo And 1&
                End Select
            End If
        Else
            shift = (rate_hi And 3&) + opl3_eg_incstep(rate_lo, opl3_chip.timer And 3&)
            If (shift And 4&) <> 0& Then
                shift = 3&
            End If
            If shift = 0& Then
                shift = opl3_chip.eg_state
            End If
        End If
    End If

    eg_rout = opl3_chip.slot(slotIdx).eg_rout
    eg_inc = 0&
    eg_off = 0&

    If (reset <> 0&) And (rate_hi = &HF&) Then
        eg_rout = 0&
    End If

    If (opl3_chip.slot(slotIdx).eg_rout And &H1F8&) = &H1F8& Then
        eg_off = 1&
    End If

    If (opl3_chip.slot(slotIdx).eg_gen <> ENVELOPE_GEN_ATTACK) And (reset = 0&) And (eg_off <> 0&) Then
        eg_rout = &H1FF&
    End If

    Select Case opl3_chip.slot(slotIdx).eg_gen
        Case ENVELOPE_GEN_ATTACK
            If opl3_chip.slot(slotIdx).eg_rout = 0& Then
                opl3_chip.slot(slotIdx).eg_gen = ENVELOPE_GEN_DECAY
            ElseIf (opl3_chip.slot(slotIdx).key <> 0&) And (shift > 0&) And (rate_hi <> &HF&) Then
                eg_inc = U32Shr(U32Shl((Not opl3_chip.slot(slotIdx).eg_rout), shift), 4&)
            End If

        Case ENVELOPE_GEN_DECAY
            If U32Shr(opl3_chip.slot(slotIdx).eg_rout, 4&) = opl3_chip.slot(slotIdx).reg_sl Then
                opl3_chip.slot(slotIdx).eg_gen = ENVELOPE_GEN_SUSTAIN
            ElseIf (eg_off = 0&) And (reset = 0&) And (shift > 0&) Then
                eg_inc = U32Shl(1&, shift - 1&)
            End If

        Case ENVELOPE_GEN_SUSTAIN, ENVELOPE_GEN_RELEASE
            If (eg_off = 0&) And (reset = 0&) And (shift > 0&) Then
                eg_inc = U32Shl(1&, shift - 1&)
            End If
    End Select

    opl3_chip.slot(slotIdx).eg_rout = (eg_rout + eg_inc) And &H1FF&

    If reset <> 0& Then opl3_chip.slot(slotIdx).eg_gen = ENVELOPE_GEN_ATTACK
    If opl3_chip.slot(slotIdx).key = 0& Then opl3_chip.slot(slotIdx).eg_gen = ENVELOPE_GEN_RELEASE
End Sub

Private Sub OPL3_EnvelopeKeyOn(ByVal slotIdx As Long, ByVal keyType As Long)
    opl3_chip.slot(slotIdx).key = opl3_chip.slot(slotIdx).key Or keyType
End Sub

Private Sub OPL3_EnvelopeKeyOff(ByVal slotIdx As Long, ByVal keyType As Long)
    opl3_chip.slot(slotIdx).key = opl3_chip.slot(slotIdx).key And (Not keyType)
End Sub

Private Sub OPL3_PhaseGenerate(ByVal slotIdx As Long)
    Dim chIdx As Long
    Dim f_num As Long
    Dim basefreq As Long
    Dim rm_xor As Long
    Dim n_bit As Long
    Dim noise As Long
    Dim phase As Long
    Dim rng As Long
    Dim vibpos As Long

    chIdx = opl3_chip.slot(slotIdx).channelIdx
    f_num = opl3_chip.channel(chIdx).f_num

    If opl3_chip.slot(slotIdx).reg_vib <> 0& Then
        rng = U32Shr(f_num, 7&) And 7&
        vibpos = opl3_chip.vibpos

        If (vibpos And 3&) = 0& Then
            rng = 0&
        ElseIf (vibpos And 1&) <> 0& Then
            rng = U32Shr(rng, 1&)
        End If

        rng = U32Shr(rng, opl3_chip.vibshift)
        If (vibpos And 4&) <> 0& Then rng = -rng

        f_num = f_num + rng
    End If

    basefreq = U32Shr(U32Shl(f_num, opl3_chip.channel(chIdx).block), 1&)
    phase = U32Shr(opl3_chip.slot(slotIdx).pg_phase, 9&)

    If opl3_chip.slot(slotIdx).pg_reset <> 0& Then
        opl3_chip.slot(slotIdx).pg_phase = 0&
    End If

    opl3_chip.slot(slotIdx).pg_phase = U32Add(opl3_chip.slot(slotIdx).pg_phase, U32Shr(basefreq * opl3_mt(opl3_chip.slot(slotIdx).reg_mult), 1&))

    noise = opl3_chip.noise
    opl3_chip.slot(slotIdx).pg_phase_out = phase And &HFFFF&

    If opl3_chip.slot(slotIdx).slot_num = 13& Then
        opl3_chip.rm_hh_bit2 = U32Shr(phase, 2&) And 1&
        opl3_chip.rm_hh_bit3 = U32Shr(phase, 3&) And 1&
        opl3_chip.rm_hh_bit7 = U32Shr(phase, 7&) And 1&
        opl3_chip.rm_hh_bit8 = U32Shr(phase, 8&) And 1&
    End If

    If (opl3_chip.slot(slotIdx).slot_num = 17&) And ((opl3_chip.rhy And &H20&) <> 0&) Then
        opl3_chip.rm_tc_bit3 = U32Shr(phase, 3&) And 1&
        opl3_chip.rm_tc_bit5 = U32Shr(phase, 5&) And 1&
    End If

    If (opl3_chip.rhy And &H20&) <> 0& Then
        rm_xor = (opl3_chip.rm_hh_bit2 Xor opl3_chip.rm_hh_bit7) Or (opl3_chip.rm_hh_bit3 Xor opl3_chip.rm_tc_bit5) Or (opl3_chip.rm_tc_bit3 Xor opl3_chip.rm_tc_bit5)

        Select Case opl3_chip.slot(slotIdx).slot_num
            Case 13&
                opl3_chip.slot(slotIdx).pg_phase_out = U32Shl(rm_xor, 9&)
                If (rm_xor Xor (noise And 1&)) <> 0& Then
                    opl3_chip.slot(slotIdx).pg_phase_out = opl3_chip.slot(slotIdx).pg_phase_out Or &HD0&
                Else
                    opl3_chip.slot(slotIdx).pg_phase_out = opl3_chip.slot(slotIdx).pg_phase_out Or &H34&
                End If

            Case 16&
                opl3_chip.slot(slotIdx).pg_phase_out = U32Shl(opl3_chip.rm_hh_bit8, 9&) Or U32Shl((opl3_chip.rm_hh_bit8 Xor (noise And 1&)), 8&)

            Case 17&
                opl3_chip.slot(slotIdx).pg_phase_out = U32Shl(rm_xor, 9&) Or &H80&
        End Select
    End If

    n_bit = (U32Shr(noise, 14&) Xor noise) And 1&
    opl3_chip.noise = U32Shr(noise, 1&) Or U32Shl(n_bit, 22&)
End Sub

Private Sub OPL3_SlotWrite20(ByVal slotIdx As Long, ByVal data As Long)
    If (U32Shr(data, 7&) And 1&) <> 0& Then
        opl3_chip.slot(slotIdx).tremSrc = SRC_TREM
    Else
        opl3_chip.slot(slotIdx).tremSrc = SRC_ZERO
    End If

    opl3_chip.slot(slotIdx).reg_vib = U32Shr(data, 6&) And 1&
    opl3_chip.slot(slotIdx).reg_type = U32Shr(data, 5&) And 1&
    opl3_chip.slot(slotIdx).reg_ksr = U32Shr(data, 4&) And 1&
    opl3_chip.slot(slotIdx).reg_mult = data And &HF&
End Sub

Private Sub OPL3_SlotWrite40(ByVal slotIdx As Long, ByVal data As Long)
    opl3_chip.slot(slotIdx).reg_ksl = U32Shr(data, 6&) And 3&
    opl3_chip.slot(slotIdx).reg_tl = data And &H3F&
    OPL3_EnvelopeUpdateKSL slotIdx
End Sub

Private Sub OPL3_SlotWrite60(ByVal slotIdx As Long, ByVal data As Long)
    opl3_chip.slot(slotIdx).reg_ar = U32Shr(data, 4&) And &HF&
    opl3_chip.slot(slotIdx).reg_dr = data And &HF&
End Sub

Private Sub OPL3_SlotWrite80(ByVal slotIdx As Long, ByVal data As Long)
    opl3_chip.slot(slotIdx).reg_sl = U32Shr(data, 4&) And &HF&
    If opl3_chip.slot(slotIdx).reg_sl = &HF& Then
        opl3_chip.slot(slotIdx).reg_sl = &H1F&
    End If
    opl3_chip.slot(slotIdx).reg_rr = data And &HF&
End Sub

Private Sub OPL3_SlotWriteE0(ByVal slotIdx As Long, ByVal data As Long)
    opl3_chip.slot(slotIdx).reg_wf = data And 7&
    If opl3_chip.newm = 0& Then
        opl3_chip.slot(slotIdx).reg_wf = opl3_chip.slot(slotIdx).reg_wf And 3&
    End If
End Sub

Private Sub OPL3_SlotGenerate(ByVal slotIdx As Long)
    opl3_chip.slot(slotIdx).outVal = OPL3_EnvelopeCalcSin(opl3_chip.slot(slotIdx).reg_wf, opl3_chip.slot(slotIdx).pg_phase_out + OPL3_ReadSrc(opl3_chip.slot(slotIdx).modSrc), opl3_chip.slot(slotIdx).eg_out)
End Sub

Private Sub OPL3_SlotCalcFB(ByVal slotIdx As Long)
    Dim chIdx As Long

    chIdx = opl3_chip.slot(slotIdx).channelIdx

    If opl3_chip.channel(chIdx).fb <> 0& Then
        opl3_chip.slot(slotIdx).fbmod = (opl3_chip.slot(slotIdx).prout + opl3_chip.slot(slotIdx).outVal) \ (2 ^ (9& - opl3_chip.channel(chIdx).fb))
    Else
        opl3_chip.slot(slotIdx).fbmod = 0&
    End If

    opl3_chip.slot(slotIdx).prout = opl3_chip.slot(slotIdx).outVal
End Sub

Private Sub OPL3_ChannelUpdateRhythm(ByVal data As Long)
    Dim chnum As Long
    Dim ch6 As Long
    Dim ch7 As Long
    Dim ch8 As Long

    opl3_chip.rhy = data And &H3F&

    If (opl3_chip.rhy And &H20&) <> 0& Then
        ch6 = 6&
        ch7 = 7&
        ch8 = 8&

        opl3_chip.channel(ch6).outSrc(0&) = OPL3_SrcSlotOut(opl3_chip.channel(ch6).slotIdx(1&))
        opl3_chip.channel(ch6).outSrc(1&) = OPL3_SrcSlotOut(opl3_chip.channel(ch6).slotIdx(1&))
        opl3_chip.channel(ch6).outSrc(2&) = SRC_ZERO
        opl3_chip.channel(ch6).outSrc(3&) = SRC_ZERO

        opl3_chip.channel(ch7).outSrc(0&) = OPL3_SrcSlotOut(opl3_chip.channel(ch7).slotIdx(0&))
        opl3_chip.channel(ch7).outSrc(1&) = OPL3_SrcSlotOut(opl3_chip.channel(ch7).slotIdx(0&))
        opl3_chip.channel(ch7).outSrc(2&) = OPL3_SrcSlotOut(opl3_chip.channel(ch7).slotIdx(1&))
        opl3_chip.channel(ch7).outSrc(3&) = OPL3_SrcSlotOut(opl3_chip.channel(ch7).slotIdx(1&))

        opl3_chip.channel(ch8).outSrc(0&) = OPL3_SrcSlotOut(opl3_chip.channel(ch8).slotIdx(0&))
        opl3_chip.channel(ch8).outSrc(1&) = OPL3_SrcSlotOut(opl3_chip.channel(ch8).slotIdx(0&))
        opl3_chip.channel(ch8).outSrc(2&) = OPL3_SrcSlotOut(opl3_chip.channel(ch8).slotIdx(1&))
        opl3_chip.channel(ch8).outSrc(3&) = OPL3_SrcSlotOut(opl3_chip.channel(ch8).slotIdx(1&))

        For chnum = 6& To 8&
            opl3_chip.channel(chnum).chtype = CH_DRUM
        Next chnum

        OPL3_ChannelSetupAlg ch6
        OPL3_ChannelSetupAlg ch7
        OPL3_ChannelSetupAlg ch8

        If (opl3_chip.rhy And 1&) <> 0& Then
            OPL3_EnvelopeKeyOn opl3_chip.channel(ch7).slotIdx(0&), EGK_DRUM
        Else
            OPL3_EnvelopeKeyOff opl3_chip.channel(ch7).slotIdx(0&), EGK_DRUM
        End If

        If (opl3_chip.rhy And 2&) <> 0& Then
            OPL3_EnvelopeKeyOn opl3_chip.channel(ch8).slotIdx(1&), EGK_DRUM
        Else
            OPL3_EnvelopeKeyOff opl3_chip.channel(ch8).slotIdx(1&), EGK_DRUM
        End If

        If (opl3_chip.rhy And 4&) <> 0& Then
            OPL3_EnvelopeKeyOn opl3_chip.channel(ch8).slotIdx(0&), EGK_DRUM
        Else
            OPL3_EnvelopeKeyOff opl3_chip.channel(ch8).slotIdx(0&), EGK_DRUM
        End If

        If (opl3_chip.rhy And 8&) <> 0& Then
            OPL3_EnvelopeKeyOn opl3_chip.channel(ch7).slotIdx(1&), EGK_DRUM
        Else
            OPL3_EnvelopeKeyOff opl3_chip.channel(ch7).slotIdx(1&), EGK_DRUM
        End If

        If (opl3_chip.rhy And &H10&) <> 0& Then
            OPL3_EnvelopeKeyOn opl3_chip.channel(ch6).slotIdx(0&), EGK_DRUM
            OPL3_EnvelopeKeyOn opl3_chip.channel(ch6).slotIdx(1&), EGK_DRUM
        Else
            OPL3_EnvelopeKeyOff opl3_chip.channel(ch6).slotIdx(0&), EGK_DRUM
            OPL3_EnvelopeKeyOff opl3_chip.channel(ch6).slotIdx(1&), EGK_DRUM
        End If
    Else
        For chnum = 6& To 8&
            opl3_chip.channel(chnum).chtype = CH_2OP
            OPL3_ChannelSetupAlg chnum
            OPL3_EnvelopeKeyOff opl3_chip.channel(chnum).slotIdx(0&), EGK_DRUM
            OPL3_EnvelopeKeyOff opl3_chip.channel(chnum).slotIdx(1&), EGK_DRUM
        Next chnum
    End If
End Sub

Private Sub OPL3_ChannelWriteA0(ByVal chIdx As Long, ByVal data As Long)
    Dim pairIdx As Long

    If (opl3_chip.newm <> 0&) And (opl3_chip.channel(chIdx).chtype = CH_4OP2) Then Exit Sub

    opl3_chip.channel(chIdx).f_num = (opl3_chip.channel(chIdx).f_num And &H300&) Or (data And &HFF&)
    opl3_chip.channel(chIdx).ksv = U32Shl(opl3_chip.channel(chIdx).block, 1&) Or (U32Shr(opl3_chip.channel(chIdx).f_num, (9& - opl3_chip.nts)) And 1&)

    OPL3_EnvelopeUpdateKSL opl3_chip.channel(chIdx).slotIdx(0&)
    OPL3_EnvelopeUpdateKSL opl3_chip.channel(chIdx).slotIdx(1&)

    If (opl3_chip.newm <> 0&) And (opl3_chip.channel(chIdx).chtype = CH_4OP) Then
        pairIdx = opl3_chip.channel(chIdx).pairIdx
        opl3_chip.channel(pairIdx).f_num = opl3_chip.channel(chIdx).f_num
        opl3_chip.channel(pairIdx).ksv = opl3_chip.channel(chIdx).ksv
        OPL3_EnvelopeUpdateKSL opl3_chip.channel(pairIdx).slotIdx(0&)
        OPL3_EnvelopeUpdateKSL opl3_chip.channel(pairIdx).slotIdx(1&)
    End If
End Sub

Private Sub OPL3_ChannelWriteB0(ByVal chIdx As Long, ByVal data As Long)
    Dim pairIdx As Long

    If (opl3_chip.newm <> 0&) And (opl3_chip.channel(chIdx).chtype = CH_4OP2) Then Exit Sub

    opl3_chip.channel(chIdx).f_num = (opl3_chip.channel(chIdx).f_num And &HFF&) Or U32Shl((data And 3&), 8&)
    opl3_chip.channel(chIdx).block = U32Shr(data, 2&) And 7&
    opl3_chip.channel(chIdx).ksv = U32Shl(opl3_chip.channel(chIdx).block, 1&) Or (U32Shr(opl3_chip.channel(chIdx).f_num, (9& - opl3_chip.nts)) And 1&)

    OPL3_EnvelopeUpdateKSL opl3_chip.channel(chIdx).slotIdx(0&)
    OPL3_EnvelopeUpdateKSL opl3_chip.channel(chIdx).slotIdx(1&)

    If (opl3_chip.newm <> 0&) And (opl3_chip.channel(chIdx).chtype = CH_4OP) Then
        pairIdx = opl3_chip.channel(chIdx).pairIdx
        opl3_chip.channel(pairIdx).f_num = opl3_chip.channel(chIdx).f_num
        opl3_chip.channel(pairIdx).block = opl3_chip.channel(chIdx).block
        opl3_chip.channel(pairIdx).ksv = opl3_chip.channel(chIdx).ksv
        OPL3_EnvelopeUpdateKSL opl3_chip.channel(pairIdx).slotIdx(0&)
        OPL3_EnvelopeUpdateKSL opl3_chip.channel(pairIdx).slotIdx(1&)
    End If
End Sub

Private Sub OPL3_ChannelSetupAlg(ByVal chIdx As Long)
    Dim pairIdx As Long
    Dim s0 As Long
    Dim s1 As Long
    Dim ps0 As Long
    Dim ps1 As Long

    s0 = opl3_chip.channel(chIdx).slotIdx(0&)
    s1 = opl3_chip.channel(chIdx).slotIdx(1&)

    If opl3_chip.channel(chIdx).chtype = CH_DRUM Then
        If (opl3_chip.channel(chIdx).ch_num = 7&) Or (opl3_chip.channel(chIdx).ch_num = 8&) Then
            opl3_chip.slot(s0).modSrc = SRC_ZERO
            opl3_chip.slot(s1).modSrc = SRC_ZERO
            Exit Sub
        End If

        If (opl3_chip.channel(chIdx).alg And 1&) = 0& Then
            opl3_chip.slot(s0).modSrc = OPL3_SrcSlotFb(s0)
            opl3_chip.slot(s1).modSrc = OPL3_SrcSlotOut(s0)
        Else
            opl3_chip.slot(s0).modSrc = OPL3_SrcSlotFb(s0)
            opl3_chip.slot(s1).modSrc = SRC_ZERO
        End If
        Exit Sub
    End If

    If (opl3_chip.channel(chIdx).alg And 8&) <> 0& Then Exit Sub

    If (opl3_chip.channel(chIdx).alg And 4&) <> 0& Then
        pairIdx = opl3_chip.channel(chIdx).pairIdx
        ps0 = opl3_chip.channel(pairIdx).slotIdx(0&)
        ps1 = opl3_chip.channel(pairIdx).slotIdx(1&)

        opl3_chip.channel(pairIdx).outSrc(0&) = SRC_ZERO
        opl3_chip.channel(pairIdx).outSrc(1&) = SRC_ZERO
        opl3_chip.channel(pairIdx).outSrc(2&) = SRC_ZERO
        opl3_chip.channel(pairIdx).outSrc(3&) = SRC_ZERO

        Select Case (opl3_chip.channel(chIdx).alg And 3&)
            Case 0&
                opl3_chip.slot(ps0).modSrc = OPL3_SrcSlotFb(ps0)
                opl3_chip.slot(ps1).modSrc = OPL3_SrcSlotOut(ps0)
                opl3_chip.slot(s0).modSrc = OPL3_SrcSlotOut(ps1)
                opl3_chip.slot(s1).modSrc = OPL3_SrcSlotOut(s0)
                opl3_chip.channel(chIdx).outSrc(0&) = OPL3_SrcSlotOut(s1)
                opl3_chip.channel(chIdx).outSrc(1&) = SRC_ZERO
                opl3_chip.channel(chIdx).outSrc(2&) = SRC_ZERO
                opl3_chip.channel(chIdx).outSrc(3&) = SRC_ZERO

            Case 1&
                opl3_chip.slot(ps0).modSrc = OPL3_SrcSlotFb(ps0)
                opl3_chip.slot(ps1).modSrc = OPL3_SrcSlotOut(ps0)
                opl3_chip.slot(s0).modSrc = SRC_ZERO
                opl3_chip.slot(s1).modSrc = OPL3_SrcSlotOut(s0)
                opl3_chip.channel(chIdx).outSrc(0&) = OPL3_SrcSlotOut(ps1)
                opl3_chip.channel(chIdx).outSrc(1&) = OPL3_SrcSlotOut(s1)
                opl3_chip.channel(chIdx).outSrc(2&) = SRC_ZERO
                opl3_chip.channel(chIdx).outSrc(3&) = SRC_ZERO

            Case 2&
                opl3_chip.slot(ps0).modSrc = OPL3_SrcSlotFb(ps0)
                opl3_chip.slot(ps1).modSrc = SRC_ZERO
                opl3_chip.slot(s0).modSrc = OPL3_SrcSlotOut(ps1)
                opl3_chip.slot(s1).modSrc = OPL3_SrcSlotOut(s0)
                opl3_chip.channel(chIdx).outSrc(0&) = OPL3_SrcSlotOut(ps0)
                opl3_chip.channel(chIdx).outSrc(1&) = OPL3_SrcSlotOut(s1)
                opl3_chip.channel(chIdx).outSrc(2&) = SRC_ZERO
                opl3_chip.channel(chIdx).outSrc(3&) = SRC_ZERO

            Case Else
                opl3_chip.slot(ps0).modSrc = OPL3_SrcSlotFb(ps0)
                opl3_chip.slot(ps1).modSrc = SRC_ZERO
                opl3_chip.slot(s0).modSrc = OPL3_SrcSlotOut(ps1)
                opl3_chip.slot(s1).modSrc = SRC_ZERO
                opl3_chip.channel(chIdx).outSrc(0&) = OPL3_SrcSlotOut(ps0)
                opl3_chip.channel(chIdx).outSrc(1&) = OPL3_SrcSlotOut(s0)
                opl3_chip.channel(chIdx).outSrc(2&) = OPL3_SrcSlotOut(s1)
                opl3_chip.channel(chIdx).outSrc(3&) = SRC_ZERO
        End Select
    Else
        If (opl3_chip.channel(chIdx).alg And 1&) = 0& Then
            opl3_chip.slot(s0).modSrc = OPL3_SrcSlotFb(s0)
            opl3_chip.slot(s1).modSrc = OPL3_SrcSlotOut(s0)
            opl3_chip.channel(chIdx).outSrc(0&) = OPL3_SrcSlotOut(s1)
            opl3_chip.channel(chIdx).outSrc(1&) = SRC_ZERO
            opl3_chip.channel(chIdx).outSrc(2&) = SRC_ZERO
            opl3_chip.channel(chIdx).outSrc(3&) = SRC_ZERO
        Else
            opl3_chip.slot(s0).modSrc = OPL3_SrcSlotFb(s0)
            opl3_chip.slot(s1).modSrc = SRC_ZERO
            opl3_chip.channel(chIdx).outSrc(0&) = OPL3_SrcSlotOut(s0)
            opl3_chip.channel(chIdx).outSrc(1&) = OPL3_SrcSlotOut(s1)
            opl3_chip.channel(chIdx).outSrc(2&) = SRC_ZERO
            opl3_chip.channel(chIdx).outSrc(3&) = SRC_ZERO
        End If
    End If
End Sub

Private Sub OPL3_ChannelWriteC0(ByVal chIdx As Long, ByVal data As Long)
    Dim pairIdx As Long

    opl3_chip.channel(chIdx).fb = U32Shr((data And &HE&), 1&)
    opl3_chip.channel(chIdx).con = data And 1&
    opl3_chip.channel(chIdx).alg = opl3_chip.channel(chIdx).con

    If opl3_chip.newm <> 0& Then
        If opl3_chip.channel(chIdx).chtype = CH_4OP Then
            pairIdx = opl3_chip.channel(chIdx).pairIdx
            opl3_chip.channel(pairIdx).alg = 4& Or U32Shl(opl3_chip.channel(chIdx).con, 1&) Or opl3_chip.channel(pairIdx).con
            opl3_chip.channel(chIdx).alg = 8&
            OPL3_ChannelSetupAlg pairIdx
        ElseIf opl3_chip.channel(chIdx).chtype = CH_4OP2 Then
            pairIdx = opl3_chip.channel(chIdx).pairIdx
            opl3_chip.channel(chIdx).alg = 4& Or U32Shl(opl3_chip.channel(pairIdx).con, 1&) Or opl3_chip.channel(chIdx).con
            opl3_chip.channel(pairIdx).alg = 8&
            OPL3_ChannelSetupAlg chIdx
        Else
            OPL3_ChannelSetupAlg chIdx
        End If
    Else
        OPL3_ChannelSetupAlg chIdx
    End If

    If opl3_chip.newm <> 0& Then
        opl3_chip.channel(chIdx).cha = IIf((U32Shr(data, 4&) And 1&) <> 0&, -1&, 0&)
        opl3_chip.channel(chIdx).chb = IIf((U32Shr(data, 5&) And 1&) <> 0&, -1&, 0&)
    Else
        opl3_chip.channel(chIdx).cha = -1&
        opl3_chip.channel(chIdx).chb = -1&
    End If
End Sub

Private Sub OPL3_ChannelKeyOn(ByVal chIdx As Long)
    Dim pairIdx As Long

    If opl3_chip.newm <> 0& Then
        If opl3_chip.channel(chIdx).chtype = CH_4OP Then
            pairIdx = opl3_chip.channel(chIdx).pairIdx
            OPL3_EnvelopeKeyOn opl3_chip.channel(chIdx).slotIdx(0&), EGK_NORM
            OPL3_EnvelopeKeyOn opl3_chip.channel(chIdx).slotIdx(1&), EGK_NORM
            OPL3_EnvelopeKeyOn opl3_chip.channel(pairIdx).slotIdx(0&), EGK_NORM
            OPL3_EnvelopeKeyOn opl3_chip.channel(pairIdx).slotIdx(1&), EGK_NORM
        ElseIf (opl3_chip.channel(chIdx).chtype = CH_2OP) Or (opl3_chip.channel(chIdx).chtype = CH_DRUM) Then
            OPL3_EnvelopeKeyOn opl3_chip.channel(chIdx).slotIdx(0&), EGK_NORM
            OPL3_EnvelopeKeyOn opl3_chip.channel(chIdx).slotIdx(1&), EGK_NORM
        End If
    Else
        OPL3_EnvelopeKeyOn opl3_chip.channel(chIdx).slotIdx(0&), EGK_NORM
        OPL3_EnvelopeKeyOn opl3_chip.channel(chIdx).slotIdx(1&), EGK_NORM
    End If
End Sub

Private Sub OPL3_ChannelKeyOff(ByVal chIdx As Long)
    Dim pairIdx As Long

    If opl3_chip.newm <> 0& Then
        If opl3_chip.channel(chIdx).chtype = CH_4OP Then
            pairIdx = opl3_chip.channel(chIdx).pairIdx
            OPL3_EnvelopeKeyOff opl3_chip.channel(chIdx).slotIdx(0&), EGK_NORM
            OPL3_EnvelopeKeyOff opl3_chip.channel(chIdx).slotIdx(1&), EGK_NORM
            OPL3_EnvelopeKeyOff opl3_chip.channel(pairIdx).slotIdx(0&), EGK_NORM
            OPL3_EnvelopeKeyOff opl3_chip.channel(pairIdx).slotIdx(1&), EGK_NORM
        ElseIf (opl3_chip.channel(chIdx).chtype = CH_2OP) Or (opl3_chip.channel(chIdx).chtype = CH_DRUM) Then
            OPL3_EnvelopeKeyOff opl3_chip.channel(chIdx).slotIdx(0&), EGK_NORM
            OPL3_EnvelopeKeyOff opl3_chip.channel(chIdx).slotIdx(1&), EGK_NORM
        End If
    Else
        OPL3_EnvelopeKeyOff opl3_chip.channel(chIdx).slotIdx(0&), EGK_NORM
        OPL3_EnvelopeKeyOff opl3_chip.channel(chIdx).slotIdx(1&), EGK_NORM
    End If
End Sub

Private Sub OPL3_ChannelSet4Op(ByVal data As Long)
    Dim bit As Long
    Dim chnum As Long

    For bit = 0& To 5&
        chnum = bit
        If bit >= 3& Then chnum = chnum + 6&

        If ((U32Shr(data, bit) And 1&) <> 0&) Then
            opl3_chip.channel(chnum).chtype = CH_4OP
            opl3_chip.channel(chnum + 3&).chtype = CH_4OP2
        Else
            opl3_chip.channel(chnum).chtype = CH_2OP
            opl3_chip.channel(chnum + 3&).chtype = CH_2OP
        End If
    Next bit
End Sub

Private Sub OPL3_Generate(ByRef outL As Long, ByRef outR As Long)
    Dim ii As Long
    Dim jj As Long
    Dim accm As Long
    Dim shift As Long

    outR = OPL3_ClipSample(opl3_chip.mixbuff(1&))

    For ii = 0& To 14&
        OPL3_SlotCalcFB ii
        OPL3_EnvelopeCalc ii
        OPL3_PhaseGenerate ii
        OPL3_SlotGenerate ii
    Next ii

    opl3_chip.mixbuff(0&) = 0&
    For ii = 0& To 17&
        accm = 0&
        For jj = 0& To 3&
            accm = accm + OPL3_ReadSrc(opl3_chip.channel(ii).outSrc(jj))
        Next jj
        opl3_chip.mixbuff(0&) = opl3_chip.mixbuff(0&) + OPL3_S16((accm And opl3_chip.channel(ii).cha) And &HFFFF&)
    Next ii

    For ii = 15& To 17&
        OPL3_SlotCalcFB ii
        OPL3_EnvelopeCalc ii
        OPL3_PhaseGenerate ii
        OPL3_SlotGenerate ii
    Next ii

    outL = OPL3_ClipSample(opl3_chip.mixbuff(0&))

    For ii = 18& To 32&
        OPL3_SlotCalcFB ii
        OPL3_EnvelopeCalc ii
        OPL3_PhaseGenerate ii
        OPL3_SlotGenerate ii
    Next ii

    opl3_chip.mixbuff(1&) = 0&
    For ii = 0& To 17&
        accm = 0&
        For jj = 0& To 3&
            accm = accm + OPL3_ReadSrc(opl3_chip.channel(ii).outSrc(jj))
        Next jj
        opl3_chip.mixbuff(1&) = opl3_chip.mixbuff(1&) + OPL3_S16((accm And opl3_chip.channel(ii).chb) And &HFFFF&)
    Next ii

    For ii = 33& To 35&
        OPL3_SlotCalcFB ii
        OPL3_EnvelopeCalc ii
        OPL3_PhaseGenerate ii
        OPL3_SlotGenerate ii
    Next ii

    If (opl3_chip.timer And &H3F&) = &H3F& Then
        opl3_chip.tremolopos = (opl3_chip.tremolopos + 1&) Mod 210&
    End If

    If opl3_chip.tremolopos < 105& Then
        opl3_chip.tremolo = U32Shr(opl3_chip.tremolopos, opl3_chip.tremoloshift)
    Else
        opl3_chip.tremolo = U32Shr((210& - opl3_chip.tremolopos), opl3_chip.tremoloshift)
    End If

    If (opl3_chip.timer And &H3FF&) = &H3FF& Then
        opl3_chip.vibpos = (opl3_chip.vibpos + 1&) And 7&
    End If

    opl3_chip.timer = (opl3_chip.timer + 1&) And &HFFFF&

    opl3_chip.eg_add = 0&
    If opl3_chip.eg_timer <> 0# Then
        shift = 0&
        Do While (shift < 36&) And (OPL3_TestBit64(opl3_chip.eg_timer, shift) = 0&)
            shift = shift + 1&
        Loop

        If shift > 12& Then
            opl3_chip.eg_add = 0&
        Else
            opl3_chip.eg_add = shift + 1&
        End If
    End If

    If (opl3_chip.eg_timerrem <> 0&) Or (opl3_chip.eg_state <> 0&) Then
        If opl3_chip.eg_timer = 68719476735# Then
            opl3_chip.eg_timer = 0#
            opl3_chip.eg_timerrem = 1&
        Else
            opl3_chip.eg_timer = opl3_chip.eg_timer + 1#
            opl3_chip.eg_timerrem = 0&
        End If
    End If

    opl3_chip.eg_state = opl3_chip.eg_state Xor 1&

    Do While opl3_chip.writebuf(opl3_chip.writebuf_cur).time <= opl3_chip.writebuf_samplecnt
        If (opl3_chip.writebuf(opl3_chip.writebuf_cur).reg And &H200&) = 0& Then Exit Do

        opl3_chip.writebuf(opl3_chip.writebuf_cur).reg = opl3_chip.writebuf(opl3_chip.writebuf_cur).reg And &H1FF&
        OPL3_WriteReg opl3_chip.writebuf(opl3_chip.writebuf_cur).reg, opl3_chip.writebuf(opl3_chip.writebuf_cur).data
        opl3_chip.writebuf_cur = (opl3_chip.writebuf_cur + 1&) Mod OPL_WRITEBUF_SIZE
    Loop

    opl3_chip.writebuf_samplecnt = opl3_chip.writebuf_samplecnt + 1#
End Sub

Private Sub OPL3_GenerateResampled(ByRef outL As Long, ByRef outR As Long)
    Dim l As Long
    Dim r As Long

    Do While opl3_chip.samplecnt >= opl3_chip.rateratio
        opl3_chip.oldsamples(0&) = opl3_chip.samples(0&)
        opl3_chip.oldsamples(1&) = opl3_chip.samples(1&)

        OPL3_Generate l, r
        opl3_chip.samples(0&) = l
        opl3_chip.samples(1&) = r

        opl3_chip.samplecnt = opl3_chip.samplecnt - opl3_chip.rateratio
    Loop

    outL = (opl3_chip.oldsamples(0&) * (opl3_chip.rateratio - opl3_chip.samplecnt) + opl3_chip.samples(0&) * opl3_chip.samplecnt) \ opl3_chip.rateratio
    outR = (opl3_chip.oldsamples(1&) * (opl3_chip.rateratio - opl3_chip.samplecnt) + opl3_chip.samples(1&) * opl3_chip.samplecnt) \ opl3_chip.rateratio

    opl3_chip.samplecnt = opl3_chip.samplecnt + U32Shl(1&, RSM_FRAC)
End Sub

Private Sub OPL3_Reset(ByVal samplerate As Long)
    Dim zeroChip As OPL3Chip_t
    Dim slotnum As Long
    Dim channum As Long
    Dim s0 As Long
    Dim s1 As Long

    opl3_chip = zeroChip

    For slotnum = 0& To 35&
        opl3_chip.slot(slotnum).modSrc = SRC_ZERO
        opl3_chip.slot(slotnum).eg_rout = &H1FF&
        opl3_chip.slot(slotnum).eg_out = &H1FF&
        opl3_chip.slot(slotnum).eg_gen = ENVELOPE_GEN_RELEASE
        opl3_chip.slot(slotnum).tremSrc = SRC_ZERO
        opl3_chip.slot(slotnum).slot_num = slotnum
    Next slotnum

    For channum = 0& To 17&
        s0 = opl3_ch_slot(channum)
        s1 = s0 + 3&

        opl3_chip.channel(channum).slotIdx(0&) = s0
        opl3_chip.channel(channum).slotIdx(1&) = s1

        opl3_chip.slot(s0).channelIdx = channum
        opl3_chip.slot(s1).channelIdx = channum

        If (channum Mod 9&) < 3& Then
            opl3_chip.channel(channum).pairIdx = channum + 3&
        ElseIf (channum Mod 9&) < 6& Then
            opl3_chip.channel(channum).pairIdx = channum - 3&
        Else
            opl3_chip.channel(channum).pairIdx = -1&
        End If

        opl3_chip.channel(channum).outSrc(0&) = SRC_ZERO
        opl3_chip.channel(channum).outSrc(1&) = SRC_ZERO
        opl3_chip.channel(channum).outSrc(2&) = SRC_ZERO
        opl3_chip.channel(channum).outSrc(3&) = SRC_ZERO
        opl3_chip.channel(channum).chtype = CH_2OP
        opl3_chip.channel(channum).cha = -1&
        opl3_chip.channel(channum).chb = -1&
        opl3_chip.channel(channum).ch_num = channum

        OPL3_ChannelSetupAlg channum
    Next channum

    opl3_chip.noise = 1&
    opl3_chip.rateratio = U32Shr(U32Shl(samplerate, RSM_FRAC), 0&) \ 49716
    If opl3_chip.rateratio <= 0& Then opl3_chip.rateratio = 1&
    opl3_chip.tremoloshift = 4&
    opl3_chip.vibshift = 1&
End Sub

Private Sub OPL3_WriteReg(ByVal reg As Long, ByVal v As Long)
    Dim high As Long
    Dim regm As Long
    Dim ad As Long
    Dim idx As Long

    high = U32Shr(reg, 8&) And 1&
    regm = reg And &HFF&

    Select Case (regm And &HF0&)
        Case 0&
            If high <> 0& Then
                Select Case (regm And &HF&)
                    Case 4&
                        OPL3_ChannelSet4Op v
                    Case 5&
                        opl3_chip.newm = v And 1&
                End Select
            Else
                If (regm And &HF&) = 8& Then
                    opl3_chip.nts = U32Shr(v, 6&) And 1&
                End If
            End If

        Case &H20&, &H30&
            ad = opl3_ad_slot(regm And &H1F&)
            If ad >= 0& Then
                idx = (18& * high) + ad
                OPL3_SlotWrite20 idx, v
            End If

        Case &H40&, &H50&
            ad = opl3_ad_slot(regm And &H1F&)
            If ad >= 0& Then
                idx = (18& * high) + ad
                OPL3_SlotWrite40 idx, v
            End If

        Case &H60&, &H70&
            ad = opl3_ad_slot(regm And &H1F&)
            If ad >= 0& Then
                idx = (18& * high) + ad
                OPL3_SlotWrite60 idx, v
            End If

        Case &H80&, &H90&
            ad = opl3_ad_slot(regm And &H1F&)
            If ad >= 0& Then
                idx = (18& * high) + ad
                OPL3_SlotWrite80 idx, v
            End If

        Case &HE0&, &HF0&
            ad = opl3_ad_slot(regm And &H1F&)
            If ad >= 0& Then
                idx = (18& * high) + ad
                OPL3_SlotWriteE0 idx, v
            End If

        Case &HA0&
            If (regm And &HF&) < 9& Then
                OPL3_ChannelWriteA0 (9& * high) + (regm And &HF&), v
            End If

        Case &HB0&
            If (regm = &HBD&) And (high = 0&) Then
                opl3_chip.tremoloshift = U32Shl((U32Shr(v, 7&) Xor 1&), 1&) + 2&
                opl3_chip.vibshift = (U32Shr(v, 6&) And 1&) Xor 1&
                OPL3_ChannelUpdateRhythm v
            ElseIf (regm And &HF&) < 9& Then
                idx = (9& * high) + (regm And &HF&)
                OPL3_ChannelWriteB0 idx, v
                If (v And &H20&) <> 0& Then
                    OPL3_ChannelKeyOn idx
                Else
                    OPL3_ChannelKeyOff idx
                End If
            End If

        Case &HC0&
            If (regm And &HF&) < 9& Then
                OPL3_ChannelWriteC0 (9& * high) + (regm And &HF&), v
            End If
    End Select
End Sub

Private Sub OPL3_WriteRegBuffered(ByVal reg As Long, ByVal v As Long)
    Dim time1 As Double
    Dim time2 As Double

    If (opl3_chip.writebuf(opl3_chip.writebuf_last).reg And &H200&) <> 0& Then
        Call OPL3_WriteReg((opl3_chip.writebuf(opl3_chip.writebuf_last).reg And &H1FF&), opl3_chip.writebuf(opl3_chip.writebuf_last).data)
        opl3_chip.writebuf_cur = (opl3_chip.writebuf_last + 1&) Mod OPL_WRITEBUF_SIZE
        opl3_chip.writebuf_samplecnt = opl3_chip.writebuf(opl3_chip.writebuf_last).time
    End If

    opl3_chip.writebuf(opl3_chip.writebuf_last).reg = reg Or &H200&
    opl3_chip.writebuf(opl3_chip.writebuf_last).data = v

    time1 = opl3_chip.writebuf_lasttime + OPL_WRITEBUF_DELAY
    time2 = opl3_chip.writebuf_samplecnt
    If time1 < time2 Then time1 = time2

    opl3_chip.writebuf(opl3_chip.writebuf_last).time = time1
    opl3_chip.writebuf_lasttime = time1
    opl3_chip.writebuf_last = (opl3_chip.writebuf_last + 1&) Mod OPL_WRITEBUF_SIZE
End Sub

Public Function OPL3_getSample(ByVal opl3 As Long) As Integer
    Dim l As Long
    Dim r As Long

    OPL3_Generate l, r
    OPL3_getSample = CInt(l)
End Function

Public Function opl2_read(ByVal dummy As Long, ByVal portnum As Integer) As Byte
    Dim ret As Long

    portnum = portnum And 1&

    If portnum = 0& Then
        ret = IIf((opl3_chip.data4 And 1&) <> 0&, &H40&, 0&)
        ret = ret Or IIf((opl3_chip.data4 And 2&) <> 0&, &H20&, 0&)
        If ret <> 0& Then ret = ret Or &H80&
        opl2_read = CByte(ret And &HFF&)
    Else
        opl2_read = &HFF&
    End If
End Function

Public Sub opl2_write(ByVal dummy As Long, ByVal portnum As Integer, ByVal value As Byte)
    portnum = portnum And 1&

    Select Case portnum
        Case 0&
            opl3_portLatch = value
        Case 1&
            If opl3_portLatch = &H4& Then
                opl3_chip.data4 = value
            End If
            OPL3_WriteRegBuffered opl3_portLatch, value
    End Select
End Sub

Public Sub OPL3_GenerateStream(ByVal opl3 As Long, ByRef samples() As Integer, ByVal count As Long)
    Dim i As Long
    Dim idx As Long
    Dim base As Long
    Dim l As Long
    Dim r As Long

    If count <= 0& Then Exit Sub

    base = LBound(samples)
    For i = 0& To count - 1&
        OPL3_GenerateResampled l, r

        idx = base + (i * 2&)
        If idx <= UBound(samples) Then samples(idx) = OPL3_ClipSample(l)
        If (idx + 1&) <= UBound(samples) Then samples(idx + 1&) = OPL3_ClipSample(r)
    Next i
End Sub

Public Function opl2_generateSample() As Integer
    Dim l As Long
    Dim r As Long

    OPL3_GenerateResampled l, r
    opl2_generateSample = OPL3_ClipSample(l)
End Function

Public Sub opl2_tickOperator(ByVal op As Long)
    ' Not used with NukedOPL path.
End Sub

Public Sub opl2_init()
    OPL3_Reset SAMPLE_RATE
End Sub

Public Sub OPL3_init(ByRef machine As MACHINE_t)
    OPL3_InitTables
    debug_log DEBUG_INFO, "[OPL] Initializing NukedOPL"
    OPL3_Reset SAMPLE_RATE
    ports_cbRegister &H388&, 2&, PORTS_CB_OPL2, PORTS_CB_NONE, PORTS_CB_OPL2, PORTS_CB_NONE, 0&
    ports_cbRegister &H228&, 2&, PORTS_CB_OPL2, PORTS_CB_NONE, PORTS_CB_OPL2, PORTS_CB_NONE, 0&
End Sub




