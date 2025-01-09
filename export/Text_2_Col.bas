Attribute VB_Name = "Text_2_Col"
Sub TextToColsText()
 Call TextToCols(2)
End Sub
Sub TextToColsGeneral()
 Call TextToCols(1)
End Sub

Sub TextToColsPO()
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, _
        2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), Array(11, 2), Array(12 _
        , 2), Array(13, 2), Array(14, 2), Array(15, 2), Array(16, 2), Array(17, 4), Array(18, 4), _
        Array(19, 4), Array(20, 2), Array(21, 2), Array(22, 2), Array(23, 2), Array(24, 2), Array( _
        25, 2), Array(26, 1), Array(27, 2), Array(28, 2), Array(29, 1), Array(30, 1), Array(31, 1), _
        Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 2), Array(36, 2), Array(37, 2), Array( _
        38, 2), Array(39, 2), Array(40, 2), Array(41, 2), Array(42, 2)), TrailingMinusNumbers _
        :=True
Call TextToColsClearSplitChar
freezeFilter
' Call TextToCols(2, strPattern)
End Sub

Sub TextToCols(iDataFormat As Integer, Optional charSeparator As String)
'Perform text to columns for custom character, setting all output to text if iDataFormat = 2 or general if iDataFormat = 1
'input parameters
'iDataFormat
'charSeparator - specify the separator character if used as
Dim TABSep As Boolean

On Error GoTo TextToColsErr
Dim strArrayParameter As String
TABSep = False
If charSeparator = "" Then
    charSeparator = Application.InputBox("Provide the separator character", "Separator", "|")
End If

If UCase(charSeparator) = "TAB" Then
    TABSep = True
    charSeparator = ""
End If
    
    Selection.TextToColumns Destination:=Selection, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=TABSep, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, _
        OtherChar:=charSeparator, _
        FieldInfo:=Array( _
                Array(1, iDataFormat), Array(2, iDataFormat), Array(3, iDataFormat), Array(4, iDataFormat), Array(5, iDataFormat), Array(6, iDataFormat), Array(7, iDataFormat), Array(8, iDataFormat), Array(9, iDataFormat), Array(10, iDataFormat), Array(11, iDataFormat), Array(12, iDataFormat), Array(13, iDataFormat), Array(14, iDataFormat), Array(15, iDataFormat), Array(16, iDataFormat), Array(17, iDataFormat), Array(18, iDataFormat), Array(19, iDataFormat), Array(20, iDataFormat), _
                Array(21, iDataFormat), Array(22, iDataFormat), Array(23, iDataFormat), Array(24, iDataFormat), Array(25, iDataFormat), Array(26, iDataFormat), Array(27, iDataFormat), Array(28, iDataFormat), Array(29, iDataFormat), Array(30, iDataFormat), Array(31, iDataFormat), Array(32, iDataFormat), Array(33, iDataFormat), Array(34, iDataFormat), Array(35, iDataFormat), Array(36, iDataFormat), Array(37, iDataFormat), Array(38, iDataFormat), Array(39, iDataFormat), Array(40, iDataFormat), _
                Array(41, iDataFormat), Array(42, iDataFormat), Array(43, iDataFormat), Array(44, iDataFormat), Array(45, iDataFormat), Array(46, iDataFormat), Array(47, iDataFormat), Array(48, iDataFormat), Array(49, iDataFormat), Array(50, iDataFormat), Array(51, iDataFormat), Array(52, iDataFormat), Array(53, iDataFormat), Array(54, iDataFormat), Array(55, iDataFormat), Array(56, iDataFormat), Array(57, iDataFormat), Array(58, iDataFormat), Array(59, iDataFormat), Array(60, iDataFormat), _
                Array(61, iDataFormat), Array(62, iDataFormat), Array(63, iDataFormat), Array(64, iDataFormat), Array(65, iDataFormat), Array(66, iDataFormat), Array(67, iDataFormat), Array(68, iDataFormat), Array(69, iDataFormat), Array(70, iDataFormat), Array(71, iDataFormat), Array(72, iDataFormat), Array(73, iDataFormat), Array(74, iDataFormat), Array(75, iDataFormat), Array(76, iDataFormat), Array(77, iDataFormat), Array(78, iDataFormat), Array(79, iDataFormat), Array(80, iDataFormat), _
                Array(81, iDataFormat), Array(82, iDataFormat), Array(83, iDataFormat), Array(84, iDataFormat), Array(85, iDataFormat), Array(86, iDataFormat), Array(87, iDataFormat), Array(88, iDataFormat), Array(89, iDataFormat), Array(90, iDataFormat), Array(91, iDataFormat), Array(92, iDataFormat), Array(93, iDataFormat), Array(94, iDataFormat), Array(95, iDataFormat), Array(96, iDataFormat), Array(97, iDataFormat), Array(98, iDataFormat), Array(99, iDataFormat), Array(100, iDataFormat), _
                Array(101, iDataFormat), Array(102, iDataFormat), Array(103, iDataFormat), Array(104, iDataFormat), Array(105, iDataFormat), Array(106, iDataFormat), Array(107, iDataFormat), Array(108, iDataFormat), Array(109, iDataFormat), Array(110, iDataFormat), Array(111, iDataFormat), Array(112, iDataFormat), Array(113, iDataFormat), Array(114, iDataFormat), Array(115, iDataFormat), Array(116, iDataFormat), Array(117, iDataFormat), Array(118, iDataFormat), Array(119, iDataFormat), Array(120, iDataFormat), _
                Array(121, iDataFormat), Array(122, iDataFormat), Array(123, iDataFormat), Array(124, iDataFormat), Array(125, iDataFormat), Array(126, iDataFormat), Array(127, iDataFormat), Array(128, iDataFormat), Array(129, iDataFormat), Array(130, iDataFormat), Array(131, iDataFormat), Array(132, iDataFormat), Array(133, iDataFormat), Array(134, iDataFormat), Array(135, iDataFormat), Array(136, iDataFormat), Array(137, iDataFormat), Array(138, iDataFormat), Array(139, iDataFormat), Array(140, iDataFormat), _
                Array(141, iDataFormat), Array(142, iDataFormat), Array(143, iDataFormat), Array(144, iDataFormat), Array(145, iDataFormat), Array(146, iDataFormat), Array(147, iDataFormat), Array(148, iDataFormat), Array(149, iDataFormat), Array(150, iDataFormat), Array(151, iDataFormat), Array(152, iDataFormat), Array(153, iDataFormat), Array(154, iDataFormat), Array(155, iDataFormat), Array(156, iDataFormat), Array(157, iDataFormat), Array(158, iDataFormat), Array(159, iDataFormat), Array(160, iDataFormat), _
                Array(161, iDataFormat), Array(162, iDataFormat), Array(163, iDataFormat), Array(164, iDataFormat), Array(165, iDataFormat), Array(166, iDataFormat), Array(167, iDataFormat), Array(168, iDataFormat), Array(169, iDataFormat), Array(170, iDataFormat), Array(171, iDataFormat), Array(172, iDataFormat), Array(173, iDataFormat), Array(174, iDataFormat), Array(175, iDataFormat), Array(176, iDataFormat), Array(177, iDataFormat), Array(178, iDataFormat), Array(179, iDataFormat), Array(180, iDataFormat), _
                Array(181, iDataFormat), Array(182, iDataFormat), Array(183, iDataFormat), Array(184, iDataFormat), Array(185, iDataFormat), Array(186, iDataFormat), Array(187, iDataFormat), Array(188, iDataFormat), Array(189, iDataFormat), Array(190, iDataFormat), Array(191, iDataFormat), Array(192, iDataFormat), Array(193, iDataFormat), Array(194, iDataFormat), Array(195, iDataFormat), Array(196, iDataFormat), Array(197, iDataFormat), Array(198, iDataFormat), Array(199, iDataFormat), Array(200, iDataFormat), _
                Array(201, iDataFormat), Array(202, iDataFormat), Array(203, iDataFormat), Array(204, iDataFormat), Array(205, iDataFormat), Array(206, iDataFormat), Array(207, iDataFormat), Array(208, iDataFormat), Array(209, iDataFormat), Array(210, iDataFormat), Array(211, iDataFormat), Array(212, iDataFormat), Array(213, iDataFormat), Array(214, iDataFormat), Array(215, iDataFormat), Array(216, iDataFormat), Array(217, iDataFormat), Array(218, iDataFormat), Array(219, iDataFormat), Array(220, iDataFormat), _
                Array(221, iDataFormat), Array(222, iDataFormat), Array(223, iDataFormat), Array(224, iDataFormat), Array(225, iDataFormat), Array(226, iDataFormat), Array(227, iDataFormat), Array(228, iDataFormat), Array(229, iDataFormat), Array(230, iDataFormat), Array(231, iDataFormat), Array(232, iDataFormat), Array(233, iDataFormat), Array(234, iDataFormat), Array(235, iDataFormat), Array(236, iDataFormat), Array(237, iDataFormat), Array(238, iDataFormat), Array(239, iDataFormat), Array(240, iDataFormat), _
                Array(241, iDataFormat), Array(242, iDataFormat), Array(243, iDataFormat), Array(244, iDataFormat), Array(245, iDataFormat), Array(246, iDataFormat), Array(247, iDataFormat), Array(248, iDataFormat), Array(249, iDataFormat), Array(250, iDataFormat), Array(251, iDataFormat), Array(252, iDataFormat), Array(253, iDataFormat), Array(254, iDataFormat), Array(255, iDataFormat), Array(256, iDataFormat), Array(257, iDataFormat), Array(258, iDataFormat), Array(259, iDataFormat), Array(260, iDataFormat), _
                Array(261, iDataFormat), Array(262, iDataFormat), Array(263, iDataFormat), Array(264, iDataFormat), Array(265, iDataFormat), Array(266, iDataFormat), Array(267, iDataFormat), Array(268, iDataFormat), Array(269, iDataFormat), Array(270, iDataFormat), Array(271, iDataFormat), Array(272, iDataFormat), Array(273, iDataFormat), Array(274, iDataFormat), Array(275, iDataFormat), Array(276, iDataFormat), Array(277, iDataFormat), Array(278, iDataFormat), Array(279, iDataFormat), Array(280, iDataFormat), _
                Array(281, iDataFormat), Array(282, iDataFormat), Array(283, iDataFormat), Array(284, iDataFormat), Array(285, iDataFormat), Array(286, iDataFormat), Array(287, iDataFormat), Array(288, iDataFormat), Array(289, iDataFormat), Array(290, iDataFormat), Array(291, iDataFormat), Array(292, iDataFormat), Array(293, iDataFormat), Array(294, iDataFormat), Array(295, iDataFormat), Array(296, iDataFormat), Array(297, iDataFormat), Array(298, iDataFormat), Array(299, iDataFormat), Array(300, iDataFormat) _
                ), TrailingMinusNumbers:=True

        Call TextToColsClearSplitChar
        freezeFilter
TextToColsErr:
End Sub

Sub TextToColsClearSplitChar()

Dim clearA1 As Boolean
'expression.TextToColumns (Destination
'                       , DataType
'                       , TextQualifier
'                       , ConsecutiveDelimiter
'                       , Tab
'                       , Semicolon
'                       , Comma
'                       , Space
'                       , Other
'                       , OtherChar
'                       , FieldInfo
'                       , DecimalSeparator
'                       , ThousandsSeparator
'                       , TrailingMinusNumbers)

clearA1 = False 'if there is no value in cell A1 then an error will be generated when applying text to columns. so put something in A1 and set a flag to clear it after
If Range("a1").value = "" Then
    Range("a1").value = "24"
    clearA1 = True
End If

'appl text to columns with no delimeters to clear the settings
Range("A1").TextToColumns DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar:="", _
        FieldInfo:=Array(Array(1, 2))

'clear cell A1 if it had been empty prior to execution
If clearA1 Then
    Range("a1").value = ""
    Range("a1").NumberFormat = "General"
End If
End Sub


