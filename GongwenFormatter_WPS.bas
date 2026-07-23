Attribute VB_Name = "GongwenFormatter_WPS"
'==============================================================================
' 公文格式化工具 (WPS 兼容版本) v1.3
' 符合 GB/T 9704 标准 + 政府交付版格式标准
' 兼容 WPS Office 和 Microsoft Word
'==============================================================================

Option Explicit

' 格式模式全局变量
Private g_FormatMode As String  ' "standard" 或 "government"

' 文档结构标记
Private g_CoverTitleEnd As Long
Private g_CoverOrgIndex As Long
Private g_CoverDateIndex As Long
Private g_TocTitleIndex As Long
Private g_TocEndIndex As Long
Private g_BodyStartIndex As Long

' 页面设置 (单位：磅，1cm=28.35pt, 1mm=2.835pt)
Private Const PAGE_MARGIN_TOP As Single = 104.88      ' 37mm
Private Const PAGE_MARGIN_BOTTOM As Single = 99.225   ' 35mm
Private Const PAGE_MARGIN_LEFT As Single = 79.38      ' 28mm
Private Const PAGE_MARGIN_RIGHT As Single = 73.71     ' 26mm
Private Const HEADER_DISTANCE As Single = 51.03       ' 1.8cm
Private Const FOOTER_DISTANCE As Single = 51.03       ' 1.8cm

' 字号 (单位：磅)
Private Const FONT_SIZE_ER As Single = 22             ' 二号
Private Const FONT_SIZE_SAN As Single = 16            ' 三号
Private Const FONT_SIZE_SI As Single = 14             ' 四号
Private Const FONT_SIZE_XIAOSI As Single = 12         ' 小四
Private Const FONT_SIZE_WU As Single = 10.5           ' 五号
Private Const FONT_COLOR_BLACK As Long = 0             ' 显式黑色

' 行距 (单位：磅)
Private Const LINE_SPACING_24 As Single = 24          ' 封面标题 line=480(auto)
Private Const LINE_SPACING_35 As Single = 35          ' 封面单位/日期固定值35磅
Private Const LINE_SPACING_18 As Single = 18          ' 正文/标题 line=360(auto)

' Word/WPS 行距规则
Private Const LINE_RULE_EXACTLY As Long = 4           ' wdLineSpaceExactly
Private Const LINE_RULE_MULTIPLE As Long = 5          ' wdLineSpaceMultiple

' Word/WPS 内置样式与大纲级别
Private Const STYLE_NORMAL As Long = -1             ' wdStyleNormal
Private Const STYLE_HEADING_2 As Long = -3          ' wdStyleHeading2 / 标题 2
Private Const STYLE_HEADING_3 As Long = -4          ' wdStyleHeading3 / 标题 3
Private Const OUTLINE_LEVEL_2 As Long = 2           ' wdOutlineLevel2
Private Const OUTLINE_LEVEL_3 As Long = 3           ' wdOutlineLevel3
Private Const OUTLINE_LEVEL_BODY As Long = 10       ' wdOutlineLevelBodyText
Private Const PAGE_NUMBER_ALIGN_CENTER As Long = 1    ' wdAlignPageNumberCenter

'==============================================================================
' 格式模式选择（WPS版：直接固定为政府交付版，不弹窗）
'==============================================================================

Private Function SelectFormatMode() As String
    g_FormatMode = "government"
    SelectFormatMode = g_FormatMode
End Function

'==============================================================================
' 符号替换功能
'==============================================================================

Public Sub ReplaceSymbols()
    Dim undoRec As Object

    Application.ScreenUpdating = False

    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "符号替换"
    End If
    On Error GoTo 0

    On Error GoTo ErrorHandler

    ' 1. 替换英文逗号为中文逗号
    Call DoReplace(",", ChrW(&HFF0C))

    ' 2. 替换英文左括号为中文左括号
    Call DoReplace("(", ChrW(&HFF08))

    ' 3. 替换英文右括号为中文右括号
    Call DoReplace(")", ChrW(&HFF09))

    ' 4. 替换英文冒号为中文冒号
    Call DoReplace(":", ChrW(&HFF1A))
    Call NormalizeHeadingNumberDots

    ' 5. 智能替换引号（交替左右引号）
    Call ReplaceQuotesInternal

    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    On Error GoTo 0

    Application.ScreenUpdating = True

    MsgBox "符号替换完成！" & vbCrLf & vbCrLf & _
           "已替换：" & vbCrLf & _
           "1. 英文逗号 , -> 中文逗号 " & ChrW(&HFF0C) & vbCrLf & _
           "2. 英文括号 () -> 中文括号 " & ChrW(&HFF08) & ChrW(&HFF09) & vbCrLf & _
           "3. 英文冒号 : -> 中文冒号 " & ChrW(&HFF1A) & vbCrLf & _
           "4. 英文引号 -> 中文引号 " & ChrW(&H201C) & ChrW(&H201D) & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销", vbInformation, "符号替换"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description, vbCritical, "错误"
End Sub

' 替换函数 (跳过表格)
Private Sub DoReplace(findWhat As String, replaceWith As String)
    Dim rng As Object
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findWhat
        .Forward = True
        .Wrap = 0 ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=0) ' wdReplaceNone
            ' 检查是否在表格中 (12=wdWithInTable)
            If Not rng.Information(12) Then
                rng.Text = replaceWith
            End If
            rng.Collapse 0 ' wdCollapseEnd
        Loop
    End With
End Sub

'==============================================================================
' 智能引号替换（交替左右引号）
'==============================================================================

Public Sub ReplaceQuotesSmart()
    Dim para As Paragraph
    Dim rng As Range
    Dim txt As String
    Dim i As Long
    Dim isLeft As Boolean
    Dim undoRec As Object
    Dim quoteChar As String

    Application.ScreenUpdating = False

    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "智能引号替换"
    End If
    On Error GoTo 0

    On Error GoTo ErrorHandler

    quoteChar = Chr(34)
    isLeft = True

    For Each para In ActiveDocument.Paragraphs
        If Not para.Range.Information(12) Then ' 跳过表格
            Set rng = para.Range
            txt = rng.Text

            If InStr(txt, quoteChar) > 0 Then
                For i = 1 To Len(txt)
                    If Mid(txt, i, 1) = quoteChar Then
                        Set rng = ActiveDocument.Range(para.Range.Start + i - 1, para.Range.Start + i)
                        If isLeft Then
                            rng.Text = ChrW(&H201C)
                        Else
                            rng.Text = ChrW(&H201D)
                        End If
                        isLeft = Not isLeft
                    End If
                Next i
            End If
        End If
    Next para

    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "智能引号替换完成！", vbInformation, "智能引号"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description, vbCritical, "错误"
End Sub

' 内部引号替换函数
Private Sub ReplaceQuotesInternal()
    Dim para As Paragraph
    Dim rng As Range
    Dim txt As String
    Dim i As Long
    Dim isLeft As Boolean
    Dim quoteChar As String

    quoteChar = Chr(34)
    isLeft = True

    For Each para In ActiveDocument.Paragraphs
        If Not para.Range.Information(12) Then ' 跳过表格
            Set rng = para.Range
            txt = rng.Text

            If InStr(txt, quoteChar) > 0 Then
                For i = 1 To Len(txt)
                    If Mid(txt, i, 1) = quoteChar Then
                        Set rng = ActiveDocument.Range(para.Range.Start + i - 1, para.Range.Start + i)
                        If isLeft Then
                            rng.Text = ChrW(&H201C)
                        Else
                            rng.Text = ChrW(&H201D)
                        End If
                        isLeft = Not isLeft
                    End If
                Next i
            End If
        End If
    Next para
End Sub

'==============================================================================
' 全部符号替换（包含智能引号）
'==============================================================================

Public Sub ReplaceAllSymbols()
    Dim undoRec As Object

    Application.ScreenUpdating = False

    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "全部符号替换"
    End If
    On Error GoTo 0

    On Error GoTo ErrorHandler

    Call DoReplace(",", ChrW(&HFF0C))
    Call DoReplace("(", ChrW(&HFF08))
    Call DoReplace(")", ChrW(&HFF09))
    Call DoReplace(":", ChrW(&HFF1A))
    Call NormalizeHeadingNumberDots
    Call ReplaceQuotesInternal

    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    On Error GoTo 0

    Application.ScreenUpdating = True

    MsgBox "全部符号替换完成！", vbInformation, "符号替换"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description, vbCritical, "错误"
End Sub

'==============================================================================
' 正文符号替换（按标准稿：数字/标点/空格规范化）
'==============================================================================

Private Sub ReplaceBodySymbols()
    Dim para As Paragraph
    Dim rng As Range
    Dim txt As String
    Dim charCode As Long
    Dim i As Long

    Application.ScreenUpdating = False

    For Each para In ActiveDocument.Paragraphs
        On Error Resume Next
        If para.Range.Information(12) Then GoTo NextPara ' 跳过表格 (12=wdWithInTable)
        On Error GoTo 0

        Set rng = para.Range
        txt = para.Range.Text
        If Len(Trim(txt)) <= 1 Then GoTo NextPara

        ' 逐字符处理：西文/数字用 Times New Roman
        For i = rng.Start To rng.End - 1
            On Error Resume Next
            Dim charRng As Range
            Set charRng = ActiveDocument.Range(i, i + 1)
            If Not charRng Is Nothing Then
                charCode = AscW(charRng.Text)
                If (charCode >= 32 And charCode <= 126) Or (charCode >= 48 And charCode <= 57) Then
                    charRng.Font.NameAscii = "Times New Roman"
                End If
            End If
            On Error GoTo 0
        Next i

NextPara:
    Next para

    Application.ScreenUpdating = True
End Sub
Private Sub NormalizeHeadingNumberDots()
    Dim para As Paragraph
    Dim rng As Range, dotRng As Range
    Dim txt As String, ch As String, fullDot As String
    Dim i As Long

    On Error Resume Next
    fullDot = ChrW(&HFF0E)

    For Each para In ActiveDocument.Paragraphs
        If para.Range.Information(12) Then GoTo NextPara

        Set rng = para.Range
        If rng.End <= rng.Start Then GoTo NextPara
        rng.End = rng.End - 1
        txt = rng.Text

        For i = 1 To Len(txt) - 1
            ch = Mid(txt, i, 1)
            If ch <> " " And ch <> vbTab And ch <> ChrW(&H3000) Then
                If IsNumeric(ch) And Mid(txt, i + 1, 1) = fullDot Then
                    Set dotRng = ActiveDocument.Range(rng.Start + i, rng.Start + i + 1)
                    dotRng.Text = "."
                End If
                Exit For
            End If
        Next i

NextPara:
    Next para
End Sub
'==============================================================================
' 主格式化功能 (WPS增强版：封面+目录+分节+页码)
'==============================================================================

Public Sub FormatGongwen()
    Dim modeName As String

    g_FormatMode = SelectFormatMode()
    modeName = "政府交付版（按标准稿对齐）"

    MsgBox "开始格式化文档，请稍候...", vbInformation, "公文格式化工具"

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' 1. 页面设置
    Call SetupPage

    ' 2. 正文符号规范化 + 封面/正文统一格式
    Call ReplaceBodySymbols
    Call DoReplace(",", ChrW(&HFF0C))
    Call DoReplace("(", ChrW(&HFF08))
    Call DoReplace(")", ChrW(&HFF09))
    Call DoReplace(":", ChrW(&HFF1A))
    Call NormalizeHeadingNumberDots
    Call ReplaceQuotesInternal
    Call FormatCoverAndBodyParagraphs

    ' 3. 表格格式
    Call FormatAllTables(False)

    ' 4. 分节与页码
    Call EnsureSectionLayout
    Call SetupPage
    Call DetectDocumentStructure
    Call AddPageNumber

    ' 5. 更新目录，再单独格式化目录
    Call UpdateDocumentToc
    Call DetectDocumentStructure
    Call FormatTocParagraphs

    Application.ScreenUpdating = True

    MsgBox "公文格式化完成！" & vbCrLf & vbCrLf & _
           "格式模式：" & modeName & vbCrLf & vbCrLf & _
           "已完成以下操作：" & vbCrLf & _
           "√ 页面设置（A4纸、标准边距）" & vbCrLf & _
           "√ 正文/封面先行定型（标题、正文、标点、空格）" & vbCrLf & _
           "√ 表格样式同步标准稿" & vbCrLf & _
           "√ 分节与页码（封面/目录/正文）" & vbCrLf & _
           "√ 目录更新并单独格式化" & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销所有更改", vbInformation, "公文格式化工具 v1.3 (WPS兼容版)"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description & vbCrLf & "错误代码：" & Err.Number, vbCritical, "错误"
End Sub

'==============================================================================
' 页面设置
'==============================================================================

Private Sub SetupPage()
    On Error Resume Next

    Dim s As Long
    If ActiveDocument.Sections.Count = 0 Then Exit Sub

    For s = 1 To ActiveDocument.Sections.Count
        Call ApplyPageSetupToSection(ActiveDocument.Sections(s))
    Next s
End Sub

Private Sub ApplyPageSetupToSection(sec As Section)
    On Error Resume Next
    With sec.PageSetup
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .TopMargin = PAGE_MARGIN_TOP
        .BottomMargin = PAGE_MARGIN_BOTTOM
        .LeftMargin = PAGE_MARGIN_LEFT
        .RightMargin = PAGE_MARGIN_RIGHT
        .HeaderDistance = HEADER_DISTANCE
        .FooterDistance = FOOTER_DISTANCE
        .DifferentFirstPageHeaderFooter = False
        .OddAndEvenPagesHeaderFooter = False
        .LayoutMode = 1
        .Gutter = 0
        .GutterPos = 0
    End With
End Sub
'==============================================================================
' 文档结构检测（封面/目录/正文分界）
'==============================================================================

Private Sub ResetDocumentMarkers()
    g_CoverTitleEnd = 0
    g_CoverOrgIndex = 0
    g_CoverDateIndex = 0
    g_TocTitleIndex = 0
    g_TocEndIndex = 0
    g_BodyStartIndex = 0
End Sub

Private Sub DetectDocumentStructure()
    Dim i As Long, total As Long
    Dim para As Paragraph
    Dim txt As String

    Call ResetDocumentMarkers

    total = ActiveDocument.Paragraphs.Count
    If total = 0 Then Exit Sub

    ' 扫描封面：前30段
    Call DetectCoverLayout

    ' 扫描目录标题和正文起点
    For i = 1 To total
        Set para = ActiveDocument.Paragraphs(i)
        On Error Resume Next
        If para.Range.Information(12) Then GoTo NextStructPara
        On Error GoTo 0

        txt = GetCleanParaText(para)

        ' 识别目录标题
        If IsTocTitleText(txt) And g_TocTitleIndex = 0 Then
            g_TocTitleIndex = i
        End If

        ' 识别正文起点（章标题、节标题或第一个正文层级标题）
        If g_BodyStartIndex = 0 And i > g_CoverDateIndex Then
            If g_TocTitleIndex > 0 And i > g_TocTitleIndex And IsTocEntryParagraph(para) Then GoTo NextStructPara
            If IsBodyStartText(txt) Then
                g_BodyStartIndex = i
            End If
        End If

NextStructPara:
    Next i

    ' 目录结束：目录标题后到第一个正文标题之间
    If g_TocTitleIndex > 0 And g_BodyStartIndex > 0 Then
        g_TocEndIndex = g_BodyStartIndex - 1
    End If
End Sub

Private Sub DetectCoverLayout()
    Dim i As Long, total As Long
    Dim para As Paragraph
    Dim txt As String, level As String
    Dim candidateIndexes(1 To 10) As Long
    Dim candidateCount As Long, datePos As Long, orgPos As Long
    Dim scanLimit As Long

    g_CoverTitleEnd = 0
    g_CoverOrgIndex = 0
    g_CoverDateIndex = 0

    total = ActiveDocument.Paragraphs.Count
    scanLimit = total
    If scanLimit > 30 Then scanLimit = 30

    ' 收集前30段中非标题、非表格、长度<=100的候选段
    For i = 1 To scanLimit
        Set para = ActiveDocument.Paragraphs(i)
        On Error Resume Next
        If para.Range.Information(12) Then GoTo ContinueCover
        On Error GoTo 0

        txt = GetCleanParaText(para)
        If txt = "" Then GoTo ContinueCover

        level = DetectLevel(txt)
        If level <> "body" Then
            If candidateCount > 0 Then Exit For
            GoTo ContinueCover
        End If
        If Len(txt) > 100 Then
            If candidateCount > 0 Then Exit For
            GoTo ContinueCover
        End If

        candidateCount = candidateCount + 1
        If candidateCount > 10 Then Exit For
        candidateIndexes(candidateCount) = i
ContinueCover:
    Next i

    If candidateCount < 3 Then Exit Sub

    ' 从后向前找日期
    For i = candidateCount To 1 Step -1
        txt = GetCleanParaText(ActiveDocument.Paragraphs(candidateIndexes(i)))
        If IsCoverDateText(txt) Then
            datePos = i
            Exit For
        End If
    Next i
    If datePos = 0 Then datePos = candidateCount

    ' 在日期前找编制单位
    If datePos < 3 Then Exit Sub
    For i = datePos - 1 To 2 Step -1
        txt = GetCleanParaText(ActiveDocument.Paragraphs(candidateIndexes(i)))
        If IsCoverOrgText(txt) Then
            orgPos = i
            Exit For
        End If
    Next i
    If orgPos = 0 Then orgPos = datePos - 1
    If orgPos < 2 Then Exit Sub

    g_CoverDateIndex = candidateIndexes(datePos)
    g_CoverOrgIndex = candidateIndexes(orgPos)
    g_CoverTitleEnd = candidateIndexes(orgPos - 1)
End Sub

Private Function GetCleanParaText(para As Paragraph) As String
    Dim txt As String
    txt = para.Range.Text
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, ChrW(&H3000), " ")
    GetCleanParaText = Trim(txt)
End Function

Private Function IsCoverDateText(txt As String) As Boolean
    Dim normalized As String
    normalized = Replace(Replace(Replace(txt, " ", ""), vbTab, ""), ChrW(&H3000), "")

    If InStr(normalized, ChrW(&H5E74)) > 0 And InStr(normalized, ChrW(&H6708)) > 0 Then
        IsCoverDateText = True: Exit Function
    End If

    If normalized Like "####-#" Or normalized Like "####-##" Or _
       normalized Like "####/#" Or normalized Like "####/##" Or _
       normalized Like "####.#" Or normalized Like "####.##" Or _
       normalized Like "####-#-#" Or normalized Like "####-##-##" Or _
       normalized Like "####/#/#" Or normalized Like "####/##/##" Or _
       normalized Like "####.#.#" Or normalized Like "####.##.##" Then
        IsCoverDateText = True: Exit Function
    End If

    IsCoverDateText = False
End Function

Private Function IsCoverOrgText(txt As String) As Boolean
    If txt = "" Then Exit Function

    If InStr(txt, "公司") > 0 Or InStr(txt, "咨询") > 0 Or _
       InStr(txt, "研究院") > 0 Or InStr(txt, "研究所") > 0 Or _
       InStr(txt, "中心") > 0 Or InStr(txt, "办公室") > 0 Or _
       InStr(txt, "委员会") > 0 Or InStr(txt, "政府") > 0 Or _
       InStr(txt, "厅") > 0 Or InStr(txt, "局") > 0 Or _
       InStr(txt, "集团") > 0 Or InStr(txt, "大学") > 0 Or _
       InStr(txt, "学院") > 0 Then
        IsCoverOrgText = True: Exit Function
    End If

    IsCoverOrgText = False
End Function

Private Function IsTocTitleText(ByVal txt As String) As Boolean
    Dim normalized As String
    normalized = Replace(Replace(Replace(txt, " ", ""), vbTab, ""), ChrW(&H3000), "")
    IsTocTitleText = (normalized = ChrW(&H76EE) & ChrW(&H5F55)) ' 目录
End Function

Private Function IsChineseTopHeadingText(ByVal txt As String) As Boolean
    Dim cnNumbers As String, dunHao As String
    cnNumbers = ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09) & ChrW(&H56DB) & _
                ChrW(&H4E94) & ChrW(&H516D) & ChrW(&H4E03) & ChrW(&H516B) & _
                ChrW(&H4E5D) & ChrW(&H5341)
    dunHao = ChrW(&H3001)
    IsChineseTopHeadingText = (InStr(cnNumbers, Left(txt, 1)) > 0 And InStr(txt, dunHao) > 0 And InStr(txt, dunHao) <= 3)
End Function

Private Function IsBodyStartText(ByVal txt As String) As Boolean
    Dim normalized As String
    normalized = Replace(Replace(Replace(txt, " ", ""), vbTab, ""), ChrW(&H3000), "")
    IsBodyStartText = IsChineseTopHeadingText(normalized)
End Function
Private Function IsTocEntryParagraph(para As Paragraph) As Boolean
    On Error Resume Next
    Dim styleName As String
    styleName = LCase(CStr(para.Style))

    If InStr(styleName, "toc") > 0 Or InStr(styleName, ChrW(&H76EE) & ChrW(&H5F55)) > 0 Then
        IsTocEntryParagraph = True
        Exit Function
    End If

    If InStr(para.Range.Text, vbTab) > 0 Then
        IsTocEntryParagraph = True
        Exit Function
    End If

    IsTocEntryParagraph = False
End Function

'==============================================================================
' 封面/正文段落格式化
'==============================================================================

Private Sub FormatCoverAndBodyParagraphs()
    Dim para As Paragraph
    Dim i As Long, total As Long
    Dim pct As Single, barLen As Integer, j As Integer
    Dim barText As String, prevPct As Integer

    Call DetectDocumentStructure

    total = ActiveDocument.Paragraphs.Count
    prevPct = -1

    For i = 1 To total
        Set para = ActiveDocument.Paragraphs(i)
        On Error Resume Next
        If para.Range.Information(12) Then GoTo NextFmtPara ' 跳过表格
        On Error GoTo 0

        Call FormatSingleParagraph(para, i)

        ' 每1%更新进度条
        pct = i / total * 100
        If Int(pct) <> prevPct Then
            prevPct = Int(pct)
            barLen = Int(pct / 5)
            barText = "["
            For j = 1 To 20
                If j <= barLen Then
                    barText = barText & ChrW(&H2588)
                Else
                    barText = barText & ChrW(&H2591)
                End If
            Next j
            barText = barText & "] " & Int(pct) & "%"
            Application.StatusBar = barText & "  " & i & "/" & total
            DoEvents
        End If

NextFmtPara:
    Next i

    Application.StatusBar = "格式化完成"
End Sub

'==============================================================================
' 单个段落格式化
'==============================================================================

Private Sub FormatSingleParagraph(para As Paragraph, Optional ByVal paraIndex As Long = 0)
    Dim txt As String, level As String

    On Error Resume Next
    txt = Trim(para.Range.Text)
    If Len(txt) <= 1 Then Exit Sub
    On Error GoTo 0

    level = DetectLevel(txt)

    ' 封面段落优先处理
    If paraIndex > 0 Then
        If g_CoverTitleEnd > 0 And paraIndex <= g_CoverTitleEnd Then
            Call ApplyCoverTitleStyle(para): Exit Sub
        End If
        If paraIndex = g_CoverOrgIndex Then
            Call ApplyCoverOrgStyle(para): Exit Sub
        End If
        If paraIndex = g_CoverDateIndex Then
            Call ApplyCoverDateStyle(para): Exit Sub
        End If
    End If

    Select Case level
        Case "level1": Call ApplyLevel1Style(para)
        Case "level1b": Call ApplyLevel1bStyle(para)
        Case "level2": Call ApplyLevel2Style(para)
        Case "level3": Call ApplyLevel3Style(para)
        Case "level4": Call ApplyLevel4Style(para)
        Case "level5": Call ApplyLevel5Style(para)
        Case "level6": Call ApplyLevel6Style(para)
        Case "cover_title": Call ApplyCoverTitleStyle(para)
        Case "cover_org": Call ApplyCoverOrgStyle(para)
        Case "cover_date": Call ApplyCoverDateStyle(para)
        Case "table_title": Call ApplyTableTitleStyle(para)
        Case "figure_title": Call ApplyFigureTitleStyle(para)
        Case "note": Call ApplyNoteStyle(para)
        Case "toc_title": Call ApplyTocTitleStyle(para)
        Case "toc_entry": Call ApplyTocEntryStyle(para)
        Case Else: Call ApplyBodyStyle(para)
    End Select

    Call FormatMixedText(para)
End Sub

'==============================================================================
' 标题级别检测
'==============================================================================

Private Function DetectLevel(txt As String) As String
    Dim firstChar As String, secondChar As String
    Dim cnNumbers As String, dunHao As String, lBracket As String
    Dim normalized As String

    cnNumbers = ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09) & ChrW(&H56DB) & _
                ChrW(&H4E94) & ChrW(&H516D) & ChrW(&H4E03) & ChrW(&H516B) & _
                ChrW(&H4E5D) & ChrW(&H5341)
    dunHao = ChrW(&H3001)      ' 顿号
    lBracket = ChrW(&HFF08)    ' 全角左括号

    txt = Replace(Replace(txt, vbCr, ""), vbLf, "")
    normalized = Replace(Replace(Replace(txt, " ", ""), vbTab, ""), ChrW(&H3000), "")
    If Len(normalized) = 0 Then DetectLevel = "body": Exit Function

    firstChar = Left(normalized, 1)
    If Len(normalized) > 1 Then secondChar = Mid(normalized, 2, 1) Else secondChar = ""

    ' 表格标题
    If firstChar = ChrW(&H8868) Then DetectLevel = "table_title": Exit Function
    ' 图片标题
    If firstChar = ChrW(&H56FE) Then DetectLevel = "figure_title": Exit Function
    ' 数据来源/注释
    If IsNoteText(normalized) Then DetectLevel = "note": Exit Function
    ' 目录标题
    If IsTocTitleText(normalized) Then DetectLevel = "toc_title": Exit Function

    ' 一级标题：一、二、三... + 顿号
    If IsChineseTopHeadingText(normalized) Then
        DetectLevel = "level1": Exit Function
    End If

    ' 二级标题：（一）（二）... 全角/半角括号 + 中文数字
    If (firstChar = lBracket Or firstChar = "(") And InStr(cnNumbers, secondChar) > 0 Then
        DetectLevel = "level2": Exit Function
    End If

    ' 三级标题：1. 2. 3. ... 数字 + 半角点
    If IsNumeric(firstChar) And secondChar = "." Then
        DetectLevel = "level3": Exit Function
    End If

    ' 四级标题：（1）/(1) ... 括号 + 阿拉伯数字
    If (firstChar = lBracket Or firstChar = "(") And IsNumeric(secondChar) Then
        DetectLevel = "level4": Exit Function
    End If

    ' 五级标题：带圈数字
    If IsCircledNumber(firstChar) Then DetectLevel = "level5": Exit Function

    DetectLevel = "body"
End Function

Private Function IsNoteText(ByVal txt As String) As Boolean
    IsNoteText = (Left(txt, 5) = ChrW(&H6570) & ChrW(&H636E) & ChrW(&H6765) & ChrW(&H6E90) & ChrW(&HFF1A)) Or _
                 (Left(txt, 3) = ChrW(&H6765) & ChrW(&H6E90) & ChrW(&HFF1A)) Or _
                 (Left(txt, 2) = ChrW(&H6CE8) & ChrW(&HFF1A)) Or _
                 (Left(txt, 3) = ChrW(&H8BF4) & ChrW(&H660E) & ChrW(&HFF1A))
End Function
Private Function IsCircledNumber(char As String) As Boolean
    Dim code As Long
    If Len(char) = 0 Then IsCircledNumber = False: Exit Function
    code = AscW(char)
    IsCircledNumber = (code >= 9312 And code <= 9321)
End Function

'==============================================================================
' 样式应用函数
'==============================================================================

Private Function StandardFirstLineIndent() As Single
    StandardFirstLineIndent = 32
End Function

Private Sub ApplyBuiltInHeadingLevel(para As Paragraph, ByVal styleId As Long, ByVal outlineLevel As Long)
    On Error Resume Next
    para.Range.Style = styleId
    para.Format.OutlineLevel = outlineLevel
End Sub

Private Sub ApplyBodyOutlineLevel(para As Paragraph)
    On Error Resume Next
    para.Range.Style = STYLE_NORMAL
    para.Format.OutlineLevel = OUTLINE_LEVEL_BODY
End Sub
Private Sub ApplyAutoLineSpacing(fmt As ParagraphFormat, Optional ByVal lineSpacing As Single = LINE_SPACING_18)
    On Error Resume Next
    fmt.LineSpacingRule = LINE_RULE_MULTIPLE
    fmt.LineSpacing = lineSpacing
End Sub

Private Sub ApplyStandardParagraphFormat(para As Paragraph, ByVal alignment As Long, ByVal firstIndent As Single, Optional ByVal lineSpacing As Single = LINE_SPACING_18)
    On Error Resume Next
    With para.Format
        .Alignment = alignment
        Call ApplyAutoLineSpacing(para.Format, lineSpacing)
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = firstIndent: .LeftIndent = 0: .RightIndent = 0
    End With
End Sub

Private Sub ApplyLevel1Style(para As Paragraph)
    On Error Resume Next
    Call ApplyBuiltInHeadingLevel(para, STYLE_HEADING_2, OUTLINE_LEVEL_2)
    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 3, StandardFirstLineIndent())
End Sub

Private Sub ApplyLevel1bStyle(para As Paragraph)
    ' 旧版“第一节”入口保留为兼容项；新标准不再主动识别章/节层级。
    Call ApplyLevel1Style(para)
End Sub

Private Sub ApplyCoverTitleStyle(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("方正小标宋简体", "华文中宋", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_ER
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 1, 0, LINE_SPACING_24)
End Sub

Private Sub ApplyCoverOrgStyle(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = LINE_RULE_EXACTLY: .LineSpacing = LINE_SPACING_35
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
    End With
End Sub

Private Sub ApplyCoverDateStyle(para As Paragraph)
    Call ApplyCoverOrgStyle(para)
End Sub

Private Sub ApplyLevel2Style(para As Paragraph)
    On Error Resume Next
    Call ApplyBuiltInHeadingLevel(para, STYLE_HEADING_3, OUTLINE_LEVEL_3)
    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 3, StandardFirstLineIndent())
End Sub

Private Sub ApplyLevel3Style(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 3, StandardFirstLineIndent())
End Sub

Private Sub ApplyLevel4Style(para As Paragraph)
    Call ApplyBodyStyle(para)
End Sub

Private Sub ApplyLevel5Style(para As Paragraph)
    Call ApplyBodyStyle(para)
End Sub

Private Sub ApplyLevel6Style(para As Paragraph)
    Call ApplyBodyStyle(para)
End Sub

Private Sub ApplyBodyStyle(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 3, StandardFirstLineIndent())
End Sub

Private Sub ApplyTableTitleStyle(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 1, 0)
End Sub

Private Sub ApplyFigureTitleStyle(para As Paragraph)
    Call ApplyTableTitleStyle(para)
End Sub

Private Sub ApplyNoteStyle(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_WU
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With
    Call ApplyStandardParagraphFormat(para, 3, StandardFirstLineIndent())
End Sub

Private Sub EnsureRightDotTab(fmt As ParagraphFormat, ByVal position As Single)
    On Error Resume Next
    fmt.TabStops.ClearAll
    fmt.TabStops.Add Position:=position, Alignment:=2, Leader:=1
End Sub

Private Sub ApplyTocTitleStyle(para As Paragraph)
    On Error Resume Next
    Call ApplyBodyOutlineLevel(para)
    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = 18
        .Bold = True
        .Color = FONT_COLOR_BLACK
    End With
    With para.Format
        .Alignment = 1
        Call ApplyAutoLineSpacing(para.Format, LINE_SPACING_18)
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
        Call EnsureRightDotTab(para.Format, 415)
    End With
End Sub

Private Sub ApplyTocEntryStyle(para As Paragraph, Optional ByVal tocText As String = "")
    Dim normalized As String
    Dim isSecondLevel As Boolean

    On Error Resume Next
    If tocText = "" Then tocText = GetCleanParaText(para)
    normalized = Replace(Replace(Replace(tocText, " ", ""), vbTab, ""), ChrW(&H3000), "")
    isSecondLevel = (Left(normalized, 1) = ChrW(&HFF08) Or Left(normalized, 1) = "(")

    With para.Range.Font
        If isSecondLevel Then
            .NameFarEast = GetFont("楷体", "楷体_GB2312", "华文楷体")
        Else
            .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        End If
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SI
        .Bold = False
        .Color = FONT_COLOR_BLACK
    End With

    With para.Format
        .Alignment = 3
        Call ApplyAutoLineSpacing(para.Format, LINE_SPACING_18)
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = StandardFirstLineIndent()
        If isSecondLevel Then
            .LeftIndent = StandardFirstLineIndent()
        Else
            .LeftIndent = 0
        End If
        .RightIndent = 0
        Call EnsureRightDotTab(para.Format, 442.2)
    End With
End Sub

Private Sub FormatTocParagraphs()
    Dim i As Long
    Dim para As Paragraph

    If g_TocTitleIndex = 0 Or g_TocEndIndex = 0 Then Exit Sub

    For i = g_TocTitleIndex To g_TocEndIndex
        If i > ActiveDocument.Paragraphs.Count Then Exit For
        Set para = ActiveDocument.Paragraphs(i)
        On Error Resume Next
        If para.Range.Information(12) Then GoTo NextTocPara
        On Error GoTo 0

        Dim txt As String
        txt = GetCleanParaText(para)

        If IsTocTitleText(txt) Then
            Call ApplyTocTitleStyle(para)
        Else
            Call ApplyTocEntryStyle(para, txt)
        End If

NextTocPara:
    Next i
End Sub

'==============================================================================
' 更新目录
'==============================================================================

Private Sub UpdateDocumentToc()
    On Error Resume Next
    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        If fld.Type = 13 Then ' wdFieldTOC
            fld.Update
        End If
    Next fld
End Sub

'==============================================================================
' 分节布局（封面/目录/正文分节）
'==============================================================================

Private Sub EnsureSectionLayout()
    On Error Resume Next

    Dim coverEndPara As Long
    Dim tocStartPara As Long
    Dim bodyStartPara As Long
    Dim totalParas As Long
    Dim breakRng As Range

    totalParas = ActiveDocument.Paragraphs.Count
    If totalParas < 3 Then Exit Sub

    ' 确定封面结束位置
    If g_CoverDateIndex > 0 Then
        coverEndPara = g_CoverDateIndex
    Else
        coverEndPara = 1
    End If

    ' 确定目录起点
    If g_TocTitleIndex > 0 Then
        tocStartPara = g_TocTitleIndex
    Else
        tocStartPara = coverEndPara + 1
    End If
    If tocStartPara < 2 Then tocStartPara = 2
    If tocStartPara > totalParas Then tocStartPara = totalParas

    ' 确定正文起点
    If g_BodyStartIndex > 0 Then
        bodyStartPara = g_BodyStartIndex
    ElseIf g_TocEndIndex > 0 Then
        bodyStartPara = g_TocEndIndex + 1
    Else
        bodyStartPara = tocStartPara + 1
    End If
    If bodyStartPara <= tocStartPara Then bodyStartPara = tocStartPara + 1
    If bodyStartPara > totalParas Then bodyStartPara = totalParas

    ' 清除已有的分节符：分节符位于前一节末尾，不能删除后一节第一个字符
    Dim s As Long
    For s = ActiveDocument.Sections.Count - 1 To 1 Step -1
        Dim secBrk As Range
        Set secBrk = ActiveDocument.Sections(s).Range
        secBrk.Characters.Last.Delete
    Next s

    ' 从后向前插入分节符，先正文、后目录，避免段落序号漂移
    Set breakRng = ActiveDocument.Paragraphs(bodyStartPara).Range
    If Not breakRng Is Nothing Then
        breakRng.Collapse Direction:=1 ' wdCollapseStart
        breakRng.InsertBreak Type:=2 ' wdSectionBreakNextPage
    End If

    Set breakRng = ActiveDocument.Paragraphs(tocStartPara).Range
    If Not breakRng Is Nothing Then
        breakRng.Collapse Direction:=1 ' wdCollapseStart
        breakRng.InsertBreak Type:=2 ' wdSectionBreakNextPage
    End If

    ' 断开所有节的页眉页脚链接，并在分节后重新应用页面设置。
    Call BreakHeaderFooterLinks
    Call SetupPage
End Sub

'==============================================================================
' 混合文本格式化（西文/数字用Times New Roman）
'==============================================================================

Private Sub FormatMixedText(para As Paragraph)
    Dim rng As Range
    Dim i As Long
    Dim charCode As Long

    On Error Resume Next
    Set rng = para.Range
    For i = rng.Start To rng.End - 1
        Dim charRange As Range
        Set charRange = ActiveDocument.Range(i, i + 1)
        If Not charRange Is Nothing Then
            charCode = AscW(charRange.Text)
            If (charCode >= 32 And charCode <= 126) Then
                charRange.Font.NameAscii = "Times New Roman"
            End If
        End If
    Next i
End Sub

'==============================================================================
' 字体检测
'==============================================================================

Private Function GetFont(ParamArray fonts() As Variant) As String
    Dim f As Variant
    For Each f In fonts
        If FontExists(CStr(f)) Then GetFont = CStr(f): Exit Function
    Next f
    GetFont = CStr(fonts(0))
End Function

Private Function FontExists(fontName As String) As Boolean
    On Error Resume Next
    Dim r As Range
    Set r = ActiveDocument.Range(0, 0)
    r.Font.Name = fontName
    FontExists = (r.Font.Name = fontName)
    If Err.Number <> 0 Then FontExists = False
    On Error GoTo 0
End Function

'==============================================================================
' 页码
'==============================================================================

Public Sub AddPageNumber()
    Dim secCount As Long
    Dim s As Long

    On Error Resume Next
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    secCount = ActiveDocument.Sections.Count
    If secCount = 0 Then Exit Sub

    If secCount < 3 Then
        Call DetectDocumentStructure
        Call EnsureSectionLayout
        Call SetupPage
        secCount = ActiveDocument.Sections.Count
    End If

    Call BreakHeaderFooterLinks
    For s = 1 To secCount
        Call ClearSectionFooters(ActiveDocument.Sections(s))
    Next s

    ' 第1节：封面，不编页码。
    If secCount >= 2 Then
        ' 第2节：目录，底部居中 upperRoman 页码，从 I 开始。
        Call AddSectionPageNumber(ActiveDocument.Sections(2), 1, 1)
    End If

    If secCount >= 3 Then
        ' 第3节：正文，底部居中阿拉伯数字页码，从 1 开始，不加破折号。
        Call AddSectionPageNumber(ActiveDocument.Sections(3), 0, 1)
    End If

    ActiveDocument.Fields.Update
End Sub

Private Sub BreakHeaderFooterLinks()
    Dim s As Long, idx As Long
    On Error Resume Next

    For s = 1 To ActiveDocument.Sections.Count
        For idx = 1 To 3
            ActiveDocument.Sections(s).Headers(idx).LinkToPrevious = False
            ActiveDocument.Sections(s).Footers(idx).LinkToPrevious = False
        Next idx
    Next s
End Sub

Private Sub ClearSectionFooters(sec As Section)
    Dim idx As Long
    On Error Resume Next

    sec.PageSetup.DifferentFirstPageHeaderFooter = False
    sec.PageSetup.OddAndEvenPagesHeaderFooter = False

    For idx = 1 To 3
        sec.Footers(idx).LinkToPrevious = False
        sec.Footers(idx).Range.Delete
    Next idx
End Sub

Private Sub ApplyFooterPageNumberStyle(ftr As HeaderFooter)
    On Error Resume Next
    With ftr.Range
        .ParagraphFormat.Alignment = PAGE_NUMBER_ALIGN_CENTER
        .Font.Name = "宋体"
        .Font.Size = FONT_SIZE_SI
        .Font.Color = FONT_COLOR_BLACK
    End With
End Sub
Private Sub AddSectionPageNumber(sec As Section, ByVal numberStyle As Long, ByVal startAt As Long)
    Dim ftr As HeaderFooter, rng As Range
    On Error Resume Next

    sec.PageSetup.DifferentFirstPageHeaderFooter = False
    sec.PageSetup.OddAndEvenPagesHeaderFooter = False

    Set ftr = sec.Footers(1) ' wdHeaderFooterPrimary
    ftr.LinkToPrevious = False

    With ftr.PageNumbers
        .RestartNumberingAtSection = True
        .StartingNumber = startAt
        .NumberStyle = numberStyle
    End With

    Err.Clear
    ftr.PageNumbers.Add PAGE_NUMBER_ALIGN_CENTER, True
    If Err.Number <> 0 Then
        Err.Clear
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        ftr.Range.Fields.Add Range:=rng, Type:=33 ' wdFieldPage
    End If

    With ftr.PageNumbers
        .RestartNumberingAtSection = True
        .StartingNumber = startAt
        .NumberStyle = numberStyle
    End With

    Call ApplyFooterPageNumberStyle(ftr)
    ftr.Range.Fields.Update
End Sub
'==============================================================================
' 表格格式化（增强版：边框+线宽+单元格边距）
'==============================================================================

Public Sub FormatAllTables(Optional ByVal showMessage As Boolean = True)
    Dim tbl As Table, cel As Cell, para As Paragraph

    On Error Resume Next

    For Each tbl In ActiveDocument.Tables
        Call ApplyStandardTableLayout(tbl)

        For Each cel In tbl.Range.Cells
            Call ApplyStandardCellLayout(cel)

            For Each para In cel.Range.Paragraphs
                Call ApplyStandardTableCellParagraph(para)
            Next para
        Next cel
    Next tbl

    If showMessage Then
        MsgBox "表格格式化完成！", vbInformation, "表格格式化"
    End If
End Sub

Private Sub ApplyStandardTableLayout(tbl As Table)
    On Error Resume Next

    tbl.Rows.Alignment = 1 ' wdAlignRowCenter
    tbl.LeftPadding = 5.4
    tbl.RightPadding = 5.4
    tbl.TopPadding = 0
    tbl.BottomPadding = 0

    ' 标准稿：仅保留上下外框线和内部横竖线，左右外框线为无。
    With tbl.Borders(-1) ' wdBorderTop
        .LineStyle = 1: .LineWidth = 4: .Color = 0
    End With
    With tbl.Borders(-3) ' wdBorderBottom
        .LineStyle = 1: .LineWidth = 4: .Color = 0
    End With
    With tbl.Borders(-5) ' wdBorderHorizontal
        .LineStyle = 1: .LineWidth = 4: .Color = 0
    End With
    With tbl.Borders(-6) ' wdBorderVertical
        .LineStyle = 1: .LineWidth = 4: .Color = 0
    End With
    With tbl.Borders(-2) ' wdBorderLeft
        .LineStyle = 0
    End With
    With tbl.Borders(-4) ' wdBorderRight
        .LineStyle = 0
    End With
End Sub

Private Sub ApplyStandardCellLayout(cel As Cell)
    On Error Resume Next

    cel.LeftPadding = 5.4
    cel.RightPadding = 5.4
    cel.TopPadding = 0
    cel.BottomPadding = 0
    cel.VerticalAlignment = 1 ' wdCellAlignVerticalCenter
End Sub

Private Sub ApplyStandardTableCellParagraph(para As Paragraph)
    On Error Resume Next

    para.Range.Style = "表格样式"

    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_WU
        .Bold = True
        .Color = FONT_COLOR_BLACK
    End With
    With para.Format
        .Alignment = 1 ' wdAlignParagraphCenter
        .LineSpacingRule = LINE_RULE_MULTIPLE
        .LineSpacing = LINE_SPACING_18
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
    End With
End Sub

'==============================================================================
' 其他功能
'==============================================================================

Public Sub FormatSelectedParagraphs()
    Dim p As Paragraph
    Dim count As Long

    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub

    On Error Resume Next
    For Each p In Selection.Paragraphs
        Call FormatSingleParagraph(p)
        count = count + 1
    Next p
    On Error GoTo 0

    MsgBox "选中段落格式化完成！共处理 " & count & " 个段落。", vbInformation, "段落格式化"
End Sub

Public Sub UpdatePageNumbers()
    On Error Resume Next
    Application.ScreenUpdating = False
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False

    Call DetectDocumentStructure
    Call EnsureSectionLayout
    Call SetupPage
    Call DetectDocumentStructure
    Call AddPageNumber
    ActiveDocument.Fields.Update

    Application.ScreenUpdating = True
    MsgBox "分节页码已重建！", vbInformation, "更新页码"
End Sub
Public Sub AddHeader(headerText As String)
    On Error Resume Next
    Dim hdr As HeaderFooter
    Set hdr = ActiveDocument.Sections(1).Headers(1)
    hdr.Range.Delete
    With hdr.Range
        .Text = headerText
        .ParagraphFormat.Alignment = 1
        .Font.NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .Font.NameAscii = "Times New Roman"
        .Font.Size = FONT_SIZE_SAN
        .Font.Color = FONT_COLOR_BLACK
    End With
End Sub

Public Sub FormatTitle()
    Dim para As Paragraph
    g_FormatMode = SelectFormatMode()
    For Each para In Selection.Paragraphs
        Call ApplyLevel1Style(para)
    Next para
    MsgBox "一级标题格式化完成！", vbInformation, "标题格式化"
End Sub

'==============================================================================
' 快速样式应用（带完成提示）
'==============================================================================

Public Sub ApplyLevel1ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim count As Integer, p As Paragraph
    count = 0: For Each p In Selection.Paragraphs: ApplyLevel1Style p: count = count + 1: Next
    MsgBox "一级标题样式已应用！共处理 " & count & " 个段落。", vbInformation, "一级标题"
End Sub

Public Sub ApplyLevel2ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim count As Integer, p As Paragraph
    count = 0: For Each p In Selection.Paragraphs: ApplyLevel2Style p: count = count + 1: Next
    MsgBox "二级标题样式已应用！共处理 " & count & " 个段落。", vbInformation, "二级标题"
End Sub

Public Sub ApplyLevel3ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim count As Integer, p As Paragraph
    count = 0: For Each p In Selection.Paragraphs: ApplyLevel3Style p: count = count + 1: Next
    MsgBox "三级标题样式已应用！共处理 " & count & " 个段落。", vbInformation, "三级标题"
End Sub

Public Sub ApplyLevel4ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim count As Integer, p As Paragraph
    count = 0: For Each p In Selection.Paragraphs: ApplyLevel4Style p: count = count + 1: Next
    MsgBox "四级标题样式已应用！共处理 " & count & " 个段落。", vbInformation, "四级标题"
End Sub

Public Sub ApplyLevel5ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim count As Integer, p As Paragraph
    count = 0: For Each p In Selection.Paragraphs: ApplyLevel5Style p: count = count + 1: Next
    MsgBox "五级标题样式已应用！共处理 " & count & " 个段落。", vbInformation, "五级标题"
End Sub

Public Sub ApplyBodyToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim count As Integer, p As Paragraph
    count = 0: For Each p In Selection.Paragraphs: ApplyBodyStyle p: count = count + 1: Next
    MsgBox "正文样式已应用！共处理 " & count & " 个段落。", vbInformation, "正文样式"
End Sub

'==============================================================================
' 中文别名入口（方便WPS宏列表调用）
'==============================================================================

Public Sub 全文套用报告格式()
    Call FormatGongwen
End Sub

Public Sub 替换常用的中文标点()
    Call ReplaceSymbols
End Sub

Public Sub 智能替换中文引号()
    Call ReplaceQuotesSmart
End Sub

Public Sub 按规则格式化选中文本()
    Call FormatSelectedParagraphs
End Sub

Public Sub 全文表格套用格式()
    Call FormatAllTables
End Sub

Public Sub 更新分节页码()
    Call UpdatePageNumbers
End Sub

Public Sub 全文替换全部符号()
    Call ReplaceAllSymbols
End Sub

Public Sub 选中文本设为一级标题()
    Call FormatTitle
End Sub

Public Sub 选中文本设为二级标题()
    Call ApplyLevel2ToSelection
End Sub

Public Sub 选中文本设为三级标题()
    Call ApplyLevel3ToSelection
End Sub

Public Sub 选中文本设为四级标题()
    Call ApplyLevel4ToSelection
End Sub

Public Sub 选中文本设为五级标题()
    Call ApplyLevel5ToSelection
End Sub

Public Sub 选中文本设为正文()
    Call ApplyBodyToSelection
End Sub

Public Sub 格式化公文()
    Call FormatGongwen
End Sub

Public Sub 符号替换()
    Call ReplaceSymbols
End Sub

Public Sub 智能引号()
    Call ReplaceQuotesSmart
End Sub

Public Sub 格式化选中段落()
    Call FormatSelectedParagraphs
End Sub

Public Sub 格式化表格()
    Call FormatAllTables
End Sub

Public Sub 更新页码()
    Call UpdatePageNumbers
End Sub
