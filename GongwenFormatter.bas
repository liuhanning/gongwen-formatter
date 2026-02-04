Attribute VB_Name = "GongwenFormatter"
'==============================================================================
' Gongwen Document Auto-Formatter (VBA Version)
' GB/T 9704 Standard
'==============================================================================

Option Explicit

' Page Setup (Points: 1cm=28.35pt, 1mm=2.835pt)
Private Const PAGE_MARGIN_TOP As Single = 104.88       ' 37mm
Private Const PAGE_MARGIN_BOTTOM As Single = 99.225    ' 35mm
Private Const PAGE_MARGIN_LEFT As Single = 79.38       ' 28mm
Private Const PAGE_MARGIN_RIGHT As Single = 73.71      ' 26mm
Private Const HEADER_DISTANCE As Single = 51.03        ' 1.8cm
Private Const FOOTER_DISTANCE As Single = 51.03        ' 1.8cm

' Font Size (Points)
Private Const FONT_SIZE_ER As Single = 22              ' Er Hao
Private Const FONT_SIZE_SAN As Single = 16             ' San Hao
Private Const FONT_SIZE_XIAOSI As Single = 12          ' Xiao Si
Private Const FONT_SIZE_SI As Single = 14              ' Si Hao

' Line Spacing (Points)
Private Const LINE_SPACING_30 As Single = 30
Private Const LINE_SPACING_28 As Single = 28

' Chinese Characters (Unicode)
Private Const CN_DUNHAO As String = "&H3001"           ' Dun Hao
Private Const CN_FULLSTOP As String = "&HFF0E"         ' Full-width dot
Private Const CN_LBRACKET As String = "&HFF08"         ' Full-width (
Private Const CN_RBRACKET As String = "&HFF09"         ' Full-width )

'------------------------------------------------------------------------------
' Main Entry
'------------------------------------------------------------------------------

Public Sub FormatGongwen()
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Call SetupPage
    Call FormatAllParagraphs
    Call AddPageNumber
    
    Application.ScreenUpdating = True
    MsgBox "Format completed!", vbInformation, "Gongwen Formatter"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub FormatSelectedParagraphs()
    Dim para As Paragraph
    
    Application.ScreenUpdating = False
    
    For Each para In Selection.Paragraphs
        Call FormatSingleParagraph(para)
    Next para
    
    Application.ScreenUpdating = True
    MsgBox "Selection formatted!", vbInformation, "Gongwen Formatter"
End Sub

'------------------------------------------------------------------------------
' Page Setup
'------------------------------------------------------------------------------

Private Sub SetupPage()
    With ActiveDocument.PageSetup
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .TopMargin = PAGE_MARGIN_TOP
        .BottomMargin = PAGE_MARGIN_BOTTOM
        .LeftMargin = PAGE_MARGIN_LEFT
        .RightMargin = PAGE_MARGIN_RIGHT
        .HeaderDistance = HEADER_DISTANCE
        .FooterDistance = FOOTER_DISTANCE
        .Gutter = 0
        .GutterPos = wdGutterPosLeft
    End With
End Sub

'------------------------------------------------------------------------------
' Paragraph Formatting
'------------------------------------------------------------------------------

Private Sub FormatAllParagraphs()
    Dim para As Paragraph
    Dim i As Long
    Dim total As Long
    
    total = ActiveDocument.Paragraphs.Count
    
    For i = 1 To total
        Set para = ActiveDocument.Paragraphs(i)
        Call FormatSingleParagraph(para)
        
        If i Mod 100 = 0 Then
            Application.StatusBar = "Formatting... " & i & "/" & total
        End If
    Next i
    
    Application.StatusBar = ""
End Sub

Private Sub FormatSingleParagraph(para As Paragraph)
    Dim text As String
    Dim level As String
    
    text = Trim(para.Range.text)
    
    If Len(text) <= 1 Then Exit Sub
    
    level = DetectLevel(text)
    
    Select Case level
        Case "level1"
            Call ApplyLevel1Style(para)
        Case "level2"
            Call ApplyLevel2Style(para)
        Case "level3"
            Call ApplyLevel3Style(para)
        Case "level4"
            Call ApplyLevel4Style(para)
        Case "level5"
            Call ApplyLevel5Style(para)
        Case "level6"
            Call ApplyLevel6Style(para)
        Case "table_title"
            Call ApplyTableTitleStyle(para)
        Case "figure_title"
            Call ApplyFigureTitleStyle(para)
        Case Else
            Call ApplyBodyStyle(para)
    End Select
    
    Call FormatMixedText(para)
End Sub

Private Function DetectLevel(text As String) As String
    Dim firstChar As String
    Dim secondChar As String
    Dim cnNumbers As String
    Dim dunHao As String
    Dim fullDot As String
    Dim lBracket As String
    
    ' Chinese number characters
    cnNumbers = ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09) & ChrW(&H56DB) & _
                ChrW(&H4E94) & ChrW(&H516D) & ChrW(&H4E03) & ChrW(&H516B) & _
                ChrW(&H4E5D) & ChrW(&H5341)
    
    dunHao = ChrW(&H3001)      ' Dun hao
    fullDot = ChrW(&HFF0E)     ' Full-width dot
    lBracket = ChrW(&HFF08)    ' Full-width (
    
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    
    If Len(text) = 0 Then
        DetectLevel = "body"
        Exit Function
    End If
    
    firstChar = Left(text, 1)
    If Len(text) > 1 Then secondChar = Mid(text, 2, 1) Else secondChar = ""
    
    ' Table title: starts with "Biao"
    If firstChar = ChrW(&H8868) Then
        DetectLevel = "table_title"
        Exit Function
    End If
    
    ' Figure title: starts with "Tu"
    If firstChar = ChrW(&H56FE) Then
        DetectLevel = "figure_title"
        Exit Function
    End If
    
    ' Level 2: Yi, Er, San... + DunHao
    If InStr(cnNumbers, firstChar) > 0 And InStr(text, dunHao) > 0 And InStr(text, dunHao) <= 3 Then
        DetectLevel = "level2"
        Exit Function
    End If
    
    ' Level 3: (Yi) (Er)... full-width bracket + CN number
    If firstChar = lBracket Or firstChar = "(" Then
        If InStr(cnNumbers, secondChar) > 0 Then
            DetectLevel = "level3"
            Exit Function
        End If
    End If
    
    ' Level 4: 1. 2. 3.... number + full-width dot
    If IsNumeric(firstChar) And InStr(text, fullDot) > 0 And InStr(text, fullDot) <= 3 Then
        DetectLevel = "level4"
        Exit Function
    End If
    
    ' Level 5: (1) (2)... bracket + Arabic number
    If (firstChar = lBracket Or firstChar = "(") And IsNumeric(secondChar) Then
        DetectLevel = "level5"
        Exit Function
    End If
    
    ' Level 6: Circled numbers
    If IsCircledNumber(firstChar) Then
        DetectLevel = "level6"
        Exit Function
    End If
    
    DetectLevel = "body"
End Function

Private Function IsCircledNumber(char As String) As Boolean
    Dim code As Long
    If Len(char) = 0 Then
        IsCircledNumber = False
        Exit Function
    End If
    code = AscW(char)
    IsCircledNumber = (code >= 9312 And code <= 9321)
End Function

'------------------------------------------------------------------------------
' Style Functions
'------------------------------------------------------------------------------

Private Sub ApplyLevel1Style(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("FZXiaoBiaoSong-B05", "STZhongsong", "SimSun")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_ER
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 8
        .SpaceAfter = 8
        .FirstLineIndent = 0
        .LeftIndent = 0
    End With
End Sub

Private Sub ApplyLevel2Style(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("SimHei", "Microsoft YaHei", "SimSun")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphLeft
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 8
        .SpaceAfter = 8
        .FirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0.85) * 2
    End With
End Sub

Private Sub ApplyLevel3Style(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("KaiTi_GB2312", "KaiTi", "SimKai")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphLeft
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 8
        .SpaceAfter = 8
        .FirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0.85) * 2
    End With
End Sub

Private Sub ApplyLevel4Style(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("FangSong_GB2312", "FangSong", "SimFang")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphLeft
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8
        .SpaceAfter = 8
        .FirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0.85) * 2
    End With
End Sub

Private Sub ApplyLevel5Style(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("FangSong_GB2312", "FangSong", "SimFang")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphLeft
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8
        .SpaceAfter = 8
        .FirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0.85) * 2
    End With
End Sub

Private Sub ApplyLevel6Style(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("FangSong_GB2312", "FangSong", "SimFang")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(0.85) * 2
        .LeftIndent = 0
    End With
End Sub

Private Sub ApplyBodyStyle(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("FangSong_GB2312", "FangSong", "SimFang")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(0.85) * 2
        .LeftIndent = 0
    End With
End Sub

Private Sub ApplyTableTitleStyle(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("SimHei", "Microsoft YaHei", "SimSun")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 8
        .SpaceAfter = 8
        .FirstLineIndent = 0
        .LeftIndent = 0
    End With
End Sub

Private Sub ApplyFigureTitleStyle(para As Paragraph)
    With para.Range.Font
        .NameFarEast = GetAvailableFont("SimHei", "Microsoft YaHei", "SimSun")
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
    End With
    
    With para.Format
        .Alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
        .LeftIndent = 0
    End With
End Sub

'------------------------------------------------------------------------------
' Mixed Text Formatting
'------------------------------------------------------------------------------

Private Sub FormatMixedText(para As Paragraph)
    Dim rng As Range
    Dim i As Long
    Dim charCode As Long
    Dim startPos As Long
    Dim endPos As Long
    
    Set rng = para.Range
    startPos = rng.Start
    endPos = rng.End - 1
    
    For i = startPos To endPos
        Dim charRange As Range
        Set charRange = ActiveDocument.Range(i, i + 1)
        
        charCode = AscW(charRange.text)
        
        If (charCode >= 32 And charCode <= 126) Then
            charRange.Font.Name = "Times New Roman"
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Font Compatibility
'------------------------------------------------------------------------------

Private Function GetAvailableFont(ParamArray fonts() As Variant) As String
    Dim fontName As Variant
    
    For Each fontName In fonts
        If IsFontInstalled(CStr(fontName)) Then
            GetAvailableFont = CStr(fontName)
            Exit Function
        End If
    Next fontName
    
    GetAvailableFont = CStr(fonts(0))
End Function

Private Function IsFontInstalled(fontName As String) As Boolean
    Dim testRange As Range
    
    On Error Resume Next
    
    Set testRange = ActiveDocument.Range(0, 0)
    testRange.Font.Name = fontName
    
    IsFontInstalled = (testRange.Font.Name = fontName)
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Page Number
'------------------------------------------------------------------------------

Private Sub AddPageNumber()
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim fld As Field
    Dim rng As Range
    
    For Each sec In ActiveDocument.Sections
        Set ftr = sec.Footers(wdHeaderFooterPrimary)
        
        ' 清空页脚
        ftr.Range.Delete
        
        ' 设置页脚范围
        Set rng = ftr.Range
        
        ' 添加前导破折号 "— "
        rng.InsertAfter ChrW(&H2014) & " "
        
        ' 插入页码域
        Set rng = ftr.Range
        rng.Collapse Direction:=wdCollapseEnd
        Set fld = ftr.Range.Fields.Add(Range:=rng, Type:=wdFieldPage)
        
        ' 添加后置破折号 " —"
        Set rng = ftr.Range
        rng.Collapse Direction:=wdCollapseEnd
        rng.InsertAfter " " & ChrW(&H2014)
        
        ' 格式化整个页脚
        With ftr.Range
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.Name = "SimSun"
            .Font.NameFarEast = "SimSun"
            .Font.Size = FONT_SIZE_SI
        End With
    Next sec
End Sub

'------------------------------------------------------------------------------
' Header Setup
'------------------------------------------------------------------------------

Public Sub AddHeader(headerText As String)
    Dim sec As Section
    Dim hdr As HeaderFooter
    
    Set sec = ActiveDocument.Sections(1)
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    
    hdr.Range.Delete
    
    With hdr.Range
        .text = headerText
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        .Font.NameFarEast = GetAvailableFont("FangSong_GB2312", "FangSong", "SimFang")
        .Font.NameAscii = "Times New Roman"
        .Font.Size = FONT_SIZE_SAN
    End With
End Sub

'------------------------------------------------------------------------------
' Table Formatting
'------------------------------------------------------------------------------

Public Sub FormatAllTables()
    Dim tbl As Table
    Dim Cell As Cell
    Dim para As Paragraph
    
    For Each tbl In ActiveDocument.Tables
        tbl.Rows.Alignment = wdAlignRowCenter
        
        For Each Cell In tbl.Range.Cells
            For Each para In Cell.Range.Paragraphs
                With para.Range.Font
                    .NameFarEast = GetAvailableFont("FangSong_GB2312", "FangSong", "SimFang")
                    .NameAscii = "Times New Roman"
                    .Size = FONT_SIZE_XIAOSI
                End With
                
                para.Format.Alignment = wdAlignParagraphCenter
                para.Format.LineSpacingRule = wdLineSpaceSingle
            Next para
        Next Cell
    Next tbl
    
    MsgBox "Tables formatted!", vbInformation, "Gongwen Formatter"
End Sub

'------------------------------------------------------------------------------
' Quick Style Apply
'------------------------------------------------------------------------------

Public Sub ApplyLevel1ToSelection()
    Dim para As Paragraph
    For Each para In Selection.Paragraphs
        Call ApplyLevel1Style(para)
    Next para
End Sub

Public Sub ApplyLevel2ToSelection()
    Dim para As Paragraph
    For Each para In Selection.Paragraphs
        Call ApplyLevel2Style(para)
    Next para
End Sub

Public Sub ApplyLevel3ToSelection()
    Dim para As Paragraph
    For Each para In Selection.Paragraphs
        Call ApplyLevel3Style(para)
    Next para
End Sub

Public Sub ApplyLevel4ToSelection()
    Dim para As Paragraph
    For Each para In Selection.Paragraphs
        Call ApplyLevel4Style(para)
    Next para
End Sub

Public Sub ApplyLevel5ToSelection()
    Dim para As Paragraph
    For Each para In Selection.Paragraphs
        Call ApplyLevel5Style(para)
    Next para
End Sub

Public Sub ApplyBodyToSelection()
    Dim para As Paragraph
    For Each para In Selection.Paragraphs
        Call ApplyBodyStyle(para)
    Next para
End Sub
