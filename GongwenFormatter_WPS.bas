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
Private Const FONT_SIZE_XIAOSI As Single = 12         ' 小四
Private Const FONT_SIZE_SI As Single = 14             ' 四号

' 行距 (单位：磅)
Private Const LINE_SPACING_30 As Single = 30
Private Const LINE_SPACING_28 As Single = 28

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

'==============================================================================
' 主格式化功能 (WPS增强版：封面+目录+分节+页码)
'==============================================================================

Private Sub FormatGongwen()
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
    Call ReplaceQuotesInternal
    Call FormatCoverAndBodyParagraphs

    ' 3. 表格格式
    Call FormatAllTables(False)

    ' 4. 分节与页码
    Call EnsureSectionLayout
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

        ' 识别正文起点（第一个二级标题）
        If IsLevel2Text(txt) And g_BodyStartIndex = 0 And i > 5 Then
            g_BodyStartIndex = i
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

Private Function IsLevel2Text(ByVal txt As String) As Boolean
    Dim cnNumbers As String, dunHao As String
    cnNumbers = ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09) & ChrW(&H56DB) & _
                ChrW(&H4E94) & ChrW(&H516D) & ChrW(&H4E03) & ChrW(&H516B) & _
                ChrW(&H4E5D) & ChrW(&H5341)
    dunHao = ChrW(&H3001)
    IsLevel2Text = (InStr(cnNumbers, Left(txt, 1)) > 0 And InStr(txt, dunHao) > 0 And InStr(txt, dunHao) <= 3)
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
    Dim cnNumbers As String, dunHao As String, fullDot As String, lBracket As String

    cnNumbers = ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09) & ChrW(&H56DB) & _
                ChrW(&H4E94) & ChrW(&H516D) & ChrW(&H4E03) & ChrW(&H516B) & _
                ChrW(&H4E5D) & ChrW(&H5341)
    dunHao = ChrW(&H3001)      ' 顿号
    fullDot = ChrW(&HFF0E)     ' 全角点
    lBracket = ChrW(&HFF08)    ' 全角左括号

    txt = Replace(Replace(txt, vbCr, ""), vbLf, "")
    If Len(txt) = 0 Then DetectLevel = "body": Exit Function

    firstChar = Left(txt, 1)
    If Len(txt) > 1 Then secondChar = Mid(txt, 2, 1) Else secondChar = ""

    ' 表格标题
    If firstChar = ChrW(&H8868) Then DetectLevel = "table_title": Exit Function
    ' 图片标题
    If firstChar = ChrW(&H56FE) Then DetectLevel = "figure_title": Exit Function
    ' 目录标题
    If IsTocTitleText(txt) Then DetectLevel = "toc_title": Exit Function

    ' 章标题：如"第一章""第1章"
    If txt Like "*第*章*" Then
        DetectLevel = "level1": Exit Function
    End If

    ' 节标题：如"第一节""第1节"
    If txt Like "*第*节*" Then
        DetectLevel = "level1b": Exit Function
    End If

    ' 二级标题：一、二、三... + 顿号
    If InStr(cnNumbers, firstChar) > 0 And InStr(txt, dunHao) > 0 And InStr(txt, dunHao) <= 3 Then
        DetectLevel = "level2": Exit Function
    End If

    ' 三级标题：（一）（二）... 全角括号 + 中文数字
    If (firstChar = lBracket Or firstChar = "(") And InStr(cnNumbers, secondChar) > 0 Then
        DetectLevel = "level3": Exit Function
    End If

    ' 四级标题：1．2．3．... 数字 + 全角点
    If IsNumeric(firstChar) And InStr(txt, fullDot) > 0 And InStr(txt, fullDot) <= 3 Then
        DetectLevel = "level4": Exit Function
    End If

    ' 五级标题：(1) (2)... 括号 + 阿拉伯数字
    If (firstChar = lBracket Or firstChar = "(") And IsNumeric(secondChar) Then
        DetectLevel = "level5": Exit Function
    End If

    ' 六级标题：带圈数字
    If IsCircledNumber(firstChar) Then DetectLevel = "level6": Exit Function

    ' 封面标题：在封面标题范围内的body段落
    If g_CoverTitleEnd > 0 Then DetectLevel = "cover_title": Exit Function

    DetectLevel = "body"
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

Private Sub ApplyLevel1Style(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("方正小标宋简体", "华文中宋", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_ER
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 8: .SpaceAfter = 8: .FirstLineIndent = 0: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyLevel1bStyle(para As Paragraph)
    ' 节标题（如"第一节"）：楷体_GB2312 三号，居中，固定值30磅
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 18: .SpaceAfter = 18: .FirstLineIndent = 0: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyCoverTitleStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("方正小标宋简体", "华文中宋", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_ER
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 12: .SpaceAfter = 12
        .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
    End With
End Sub

Private Sub ApplyCoverOrgStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 24: .SpaceAfter = 0
        .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
    End With
End Sub

Private Sub ApplyCoverDateStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 24: .SpaceAfter = 0
        .FirstLineIndent = 0: .LeftIndent = 0: .RightIndent = 0
    End With
End Sub

Private Sub ApplyLevel2Style(para As Paragraph)
    Dim indentValue As Single
    On Error Resume Next
    indentValue = CentimetersToPoints(0.85) * 2  ' 2字符

    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 8: .SpaceAfter = 8
        If g_FormatMode = "government" Then
            .FirstLineIndent = indentValue: .LeftIndent = 0
        Else
            .FirstLineIndent = 0: .LeftIndent = indentValue
        End If
    End With
End Sub

Private Sub ApplyLevel3Style(para As Paragraph)
    Dim indentValue As Single
    On Error Resume Next
    indentValue = CentimetersToPoints(0.85) * 2

    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8: .SpaceAfter = 8
        If g_FormatMode = "government" Then
            .FirstLineIndent = indentValue: .LeftIndent = 0
        Else
            .FirstLineIndent = 0: .LeftIndent = indentValue
        End If
    End With
End Sub

Private Sub ApplyLevel4Style(para As Paragraph)
    Dim indentValue As Single
    On Error Resume Next
    indentValue = CentimetersToPoints(0.85) * 2

    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8: .SpaceAfter = 8
        If g_FormatMode = "government" Then
            .FirstLineIndent = indentValue: .LeftIndent = 0
        Else
            .FirstLineIndent = 0: .LeftIndent = indentValue
        End If
    End With
End Sub

Private Sub ApplyLevel5Style(para As Paragraph)
    Dim indentValue As Single
    On Error Resume Next
    indentValue = CentimetersToPoints(0.85) * 2

    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8: .SpaceAfter = 8
        If g_FormatMode = "government" Then
            .FirstLineIndent = indentValue: .LeftIndent = 0
        Else
            .FirstLineIndent = 0: .LeftIndent = indentValue
        End If
    End With
End Sub

Private Sub ApplyLevel6Style(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 3: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(0.85) * 2: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyBodyStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(0.85) * 2: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyTableTitleStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 0
        .SpaceBefore = 8: .SpaceAfter = 8: .FirstLineIndent = 0: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyFigureTitleStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 0
        .SpaceBefore = 0: .SpaceAfter = 0: .FirstLineIndent = 0: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyTocTitleStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
        .NameAscii = "Times New Roman"
        .Size = 20  ' 20pt
        .Bold = False
    End With
    With para.Format
        .Alignment = 1: .LineSpacingRule = 4: .LineSpacing = 28
        .SpaceBefore = 8: .SpaceAfter = 8: .FirstLineIndent = 0: .LeftIndent = 0
    End With
End Sub

Private Sub ApplyTocEntryStyle(para As Paragraph)
    On Error Resume Next
    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = 16  ' 16pt
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = 28
        .SpaceBefore = 0: .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(0.85) * 2: .LeftIndent = 0
    End With
End Sub

'==============================================================================
' 目录段落格式化
'==============================================================================

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
            Call ApplyTocEntryStyle(para)
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
        If fld.Type = 33 Then ' wdFieldTOC = 33
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
    Dim tocBreakPara As Long
    Dim bodyBreakPara As Long
    Dim totalParas As Long
    Dim breakRng As Range

    totalParas = ActiveDocument.Paragraphs.Count

    ' 确定封面结束位置
    If g_CoverDateIndex > 0 Then
        coverEndPara = g_CoverDateIndex
    Else
        coverEndPara = 5
    End If

    ' 确定目录结束位置
    If g_TocEndIndex > 0 Then
        tocBreakPara = g_TocEndIndex
    Else
        tocBreakPara = coverEndPara + 5
    End If
    If tocBreakPara >= totalParas Then tocBreakPara = totalParas - 1
    If tocBreakPara <= coverEndPara Then tocBreakPara = coverEndPara + 1

    ' 确定正文起点
    If g_BodyStartIndex > 0 Then
        bodyBreakPara = g_BodyStartIndex
    Else
        bodyBreakPara = tocBreakPara + 2
    End If
    If bodyBreakPara >= totalParas Then bodyBreakPara = totalParas - 1

    ' 清除已有的分节符
    Dim s As Long
    For s = ActiveDocument.Sections.Count To 2 Step -1
        Dim secBrk As Range
        Set secBrk = ActiveDocument.Sections(s).Range
        secBrk.Characters(1).Delete
    Next s

    ' 从后向前插入分节符（NextPage分节符）
    ' 先插正文分节符
    Set breakRng = ActiveDocument.Paragraphs(bodyBreakPara).Range
    If Not breakRng Is Nothing Then
        breakRng.InsertBreak Type:=2 ' wdSectionBreakNextPage
    End If

    ' 再插目录分节符
    Set breakRng = ActiveDocument.Paragraphs(tocBreakPara).Range
    If Not breakRng Is Nothing Then
        breakRng.InsertBreak Type:=2 ' wdSectionBreakNextPage
    End If

    ' 断开所有节的页眉页脚链接
    For s = 2 To ActiveDocument.Sections.Count
        Dim hdr As HeaderFooter, ftr As HeaderFooter
        Set hdr = ActiveDocument.Sections(s).Headers(1)
        Set ftr = ActiveDocument.Sections(s).Footers(1)
        hdr.LinkToPrevious = False
        ftr.LinkToPrevious = False
    Next s
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

Private Sub AddPageNumber()
    Dim secCount As Long
    Dim ftr As HeaderFooter, rng As Range

    On Error Resume Next
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    secCount = ActiveDocument.Sections.Count

    If secCount >= 1 Then
        ' 第1节：封面 — 不编页码
        Set ftr = ActiveDocument.Sections(1).Footers(1)
        ftr.Range.Delete
        ftr.PageNumbers.StartingNumber = 1
    End If

    If secCount >= 2 Then
        ' 第2节：目录 — 底部居中页码，upperRoman 格式
        Set ftr = ActiveDocument.Sections(2).Footers(1)
        ftr.Range.Delete
        Set rng = ftr.Range
        rng.InsertAfter ChrW(&H2014) & " "
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        ftr.Range.Fields.Add Range:=rng, Type:=33 ' wdFieldPage
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        rng.InsertAfter " " & ChrW(&H2014)
        With ftr.Range
            .ParagraphFormat.Alignment = 1
            .Font.Name = "宋体"
            .Font.Size = FONT_SIZE_SI
        End With
        ftr.PageNumbers.NumberStyle = 2 ' wdPageNumberStyleUpperCaseRoman = 2
        ftr.PageNumbers.StartingNumber = 1
        ftr.Range.Fields.Update
    End If

    If secCount >= 3 Then
        ' 第3节：正文 — 底部居中页码，从1开始，- N - 样式
        Set ftr = ActiveDocument.Sections(3).Footers(1)
        ftr.Range.Delete
        Set rng = ftr.Range
        rng.InsertAfter ChrW(&H2014) & " "
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        ftr.Range.Fields.Add Range:=rng, Type:=33 ' wdFieldPage
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        rng.InsertAfter " " & ChrW(&H2014)
        With ftr.Range
            .ParagraphFormat.Alignment = 1
            .Font.Name = "宋体"
            .Font.Size = FONT_SIZE_SI
        End With
        ftr.PageNumbers.RestartNumberingAtSection = True
        ftr.PageNumbers.StartingNumber = 1
        ftr.PageNumbers.NumberStyle = 0 ' wdPageNumberStyleArabic = 0
        ftr.Range.Fields.Update
    End If

    ActiveDocument.Fields.Update
End Sub

'==============================================================================
' 表格格式化（增强版：边框+线宽+单元格边距）
'==============================================================================

Public Sub FormatAllTables(Optional ByVal showMessage As Boolean = True)
    Dim tbl As Table, cel As Cell, para As Paragraph
    Dim rowIdx As Long

    On Error Resume Next

    For Each tbl In ActiveDocument.Tables
        tbl.Rows.Alignment = 1 ' wdAlignRowCenter

        ' 设置表格外框线、内部横竖线（黑色单实线，线宽4）
        With tbl.Borders
            .OutsideLineStyle = 1  ' wdLineStyleSingle
            .OutsideLineWidth = 4
            .OutsideColor = 0      ' wdColorAutomatic (black)
            .InsideLineStyle = 1
            .InsideLineWidth = 4
            .InsideColor = 0
        End With

        ' 表格前后各空一行
        Dim tblRng As Range
        Set tblRng = tbl.Range
        If tblRng.Paragraphs.Count > 0 Then
            tblRng.Paragraphs(1).Format.SpaceBefore = 8
            Dim lastPara As Paragraph
            Set lastPara = tblRng.Paragraphs(tblRng.Paragraphs.Count)
            lastPara.Format.SpaceAfter = 8
        End If

        For Each cel In tbl.Range.Cells
            ' 设置单元格左右内边距约5.4pt，上下0
            cel.LeftPadding = 5.4
            cel.RightPadding = 5.4
            cel.TopPadding = 0
            cel.BottomPadding = 0

            rowIdx = cel.RowIndex
            For Each para In cel.Range.Paragraphs
                If rowIdx = 1 Then
                    ' 表头：黑体小四
                    para.Range.Font.NameFarEast = GetFont("黑体", "微软雅黑", "宋体")
                Else
                    ' 表格正文：仿宋小四
                    para.Range.Font.NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
                End If
                para.Range.Font.NameAscii = "Times New Roman"
                para.Range.Font.Size = FONT_SIZE_XIAOSI
                para.Format.Alignment = 1 ' wdAlignParagraphCenter
                para.Format.LineSpacingRule = 0 ' wdLineSpaceSingle

                ' 垂直居中
                cel.VerticalAlignment = 1 ' wdCellAlignVerticalCenter
            Next para
        Next cel
    Next tbl

    If showMessage Then
        MsgBox "表格格式化完成！", vbInformation, "表格格式化"
    End If
End Sub

'==============================================================================
' 其他功能
'==============================================================================

Public Sub UpdatePageNumbers()
    On Error Resume Next
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    ActiveDocument.Fields.Update
    MsgBox "页码已更新！", vbInformation, "更新页码"
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
    End With
End Sub

Public Sub FormatTitle()
    Dim para As Paragraph, cw As Single
    cw = CentimetersToPoints(0.85) * 2
    For Each para In Selection.Paragraphs
        para.Format.LeftIndent = 0
        para.Format.RightIndent = 0
        para.Format.FirstLineIndent = cw
        para.Format.SpaceBefore = 6
        para.Format.SpaceAfter = 6
    Next para
    MsgBox "标题格式化完成！", vbInformation, "标题格式化"
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
