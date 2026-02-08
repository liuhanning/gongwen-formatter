Attribute VB_Name = "GongwenFormatter"
'==============================================================================
' 公文格式化工具 (WPS 兼容版本) v1.1
' 符合 GB/T 9704 标准 + 政府交付版格式标准
' 兼容 WPS Office 和 Microsoft Word
'==============================================================================

Option Explicit

' 格式模式全局变量
Private g_FormatMode As String  ' "standard" 或 "government"

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

Private Function IsWPS() As Boolean
    On Error Resume Next
    IsWPS = (InStr(1, Application.Name, "WPS", vbTextCompare) > 0)
    On Error GoTo 0
End Function

'==============================================================================
' 格式模式选择对话框
'==============================================================================

Public Function SelectFormatMode() As String
    Dim result As VbMsgBoxResult

    result = MsgBox("请选择格式模式：" & vbCrLf & vbCrLf & _
                    "【是】= GB/T 9704标准模式" & vbCrLf & _
                    "      （标题左缩进2字符）" & vbCrLf & vbCrLf & _
                    "【否】= 政府交付版模式" & vbCrLf & _
                    "      （标题首行缩进2字符）" & vbCrLf & vbCrLf & _
                    "【取消】= 取消操作", _
                    vbYesNoCancel + vbQuestion, "选择格式模式")

    If result = vbYes Then
        SelectFormatMode = "standard"
    ElseIf result = vbNo Then
        SelectFormatMode = "government"
    Else
        SelectFormatMode = ""  ' 取消
    End If
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
           "1. 英文逗号 , → 中文逗号 " & ChrW(&HFF0C) & vbCrLf & _
           "2. 英文括号 () → 中文括号 " & ChrW(&HFF08) & ChrW(&HFF09) & vbCrLf & _
           "3. 英文冒号 : → 中文冒号 " & ChrW(&HFF1A) & vbCrLf & _
           "4. 英文引号 "" → 中文引号 " & ChrW(&H201C) & ChrW(&H201D) & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销", vbInformation, "符号替换"
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description & vbCrLf & "错误代码：" & Err.Number, vbCritical, "错误"
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
        .Wrap = 0 ' wdFindStop = 0
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Replace:=0) ' wdReplaceNone = 0
            ' 检查是否在表格中，如果不在则替换 (12 是 wdWithInTable)
            If Not rng.Information(12) Then
                rng.Text = replaceWith
            End If
            rng.Collapse 0 ' wdCollapseEnd = 0
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
        ' 跳过表格中的段落 (12 是 wdWithInTable)
        If Not para.Range.Information(12) Then
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
    MsgBox "智能引号替换完成！" & vbCrLf & vbCrLf & _
           "英文引号已替换为中文引号：" & ChrW(&H201C) & " 和 " & ChrW(&H201D) & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销", vbInformation, "智能引号"
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description, vbCritical, "错误"
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
    
    ' 替换逗号
    Call DoReplace(",", ChrW(&HFF0C))
    
    ' 替换括号
    Call DoReplace("(", ChrW(&HFF08))
    Call DoReplace(")", ChrW(&HFF09))
    
    ' 替换冒号
    Call DoReplace(":", ChrW(&HFF1A))
    
    ' 智能替换引号
    Call ReplaceQuotesInternal
    
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    MsgBox "全部符号替换完成！" & vbCrLf & vbCrLf & _
           "已替换：" & vbCrLf & _
           "1. 逗号 , → " & ChrW(&HFF0C) & vbCrLf & _
           "2. 括号 () → " & ChrW(&HFF08) & ChrW(&HFF09) & vbCrLf & _
           "3. 冒号 : → " & ChrW(&HFF1A) & vbCrLf & _
           "4. 引号 "" → " & ChrW(&H201C) & ChrW(&H201D) & " (智能交替)" & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销", vbInformation, "符号替换"
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
        ' 跳过表格中的段落 (12 是 wdWithInTable)
        If Not para.Range.Information(12) Then
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
' 主格式化功能
'==============================================================================

Public Sub FormatGongwen()
    Dim undoRec As Object
    Dim modeName As String

    ' 选择格式模式
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If

    ' 显示模式名称
    If g_FormatMode = "government" Then
        modeName = "政府交付版（标题首行缩进）"
    Else
        modeName = "GB/T 9704标准（标题左缩进）"
    End If

    Application.ScreenUpdating = False

    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "格式化公文"
    End If
    On Error GoTo 0

    On Error GoTo ErrorHandler

    ' 1. 页面设置
    Call SetupPage

    ' 2. 段落格式化
    Call FormatAllParagraphs

    ' 3. 添加页码
    Call AddPageNumber

    ' 4. 符号替换
    Call DoReplace(",", ChrW(&HFF0C))
    Call DoReplace("(", ChrW(&HFF08))
    Call DoReplace(")", ChrW(&HFF09))
    Call DoReplace(":", ChrW(&HFF1A))

    ' 5. 智能引号替换
    Call ReplaceQuotesInternal

    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "公文格式化完成！" & vbCrLf & vbCrLf & _
           "格式模式：" & modeName & vbCrLf & vbCrLf & _
           "已完成以下操作：" & vbCrLf & _
           "√ 页面设置（A4纸、标准边距）" & vbCrLf & _
           "√ 段落格式化（字体、行距、缩进）" & vbCrLf & _
           "√ 添加页码（页脚居中）" & vbCrLf & _
           "√ 符号替换（逗号、括号、冒号、引号）" & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销所有更改", vbInformation, "公文格式化工具 v1.1 (WPS兼容版)"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description, vbCritical, "错误"
End Sub

Public Sub FormatSelectedParagraphs()
    Dim para As Paragraph
    Dim undoRec As Object
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "格式化选中段落"
    End If
    On Error GoTo 0
    
    On Error GoTo ErrorHandler
    
    For Each para In Selection.Paragraphs
        Call FormatSingleParagraph(para)
    Next para
    
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    MsgBox "选中段落格式化完成！" & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+Z 可撤销", vbInformation, "段落格式化"
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description, vbCritical, "错误"
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
' 段落格式化
'==============================================================================

Private Sub FormatAllParagraphs()
    Dim para As Paragraph
    Dim i As Long, total As Long
    
    total = ActiveDocument.Paragraphs.Count
    
    For i = 1 To total
        Set para = ActiveDocument.Paragraphs(i)
        ' 跳过表格中的段落 (12 是 wdWithInTable)
        If Not para.Range.Information(12) Then
            Call FormatSingleParagraph(para)
        End If
        If i Mod 100 = 0 Then Application.StatusBar = "正在格式化... " & i & "/" & total
    Next i
    
    Application.StatusBar = ""
End Sub

Private Sub FormatSingleParagraph(para As Paragraph)
    Dim txt As String, level As String
    
    On Error Resume Next
    txt = Trim(para.Range.Text)
    If Len(txt) <= 1 Then Exit Sub
    
    level = DetectLevel(txt)
    
    Select Case level
        Case "level1": Call ApplyLevel1Style(para)
        Case "level2": Call ApplyLevel2Style(para)
        Case "level3": Call ApplyLevel3Style(para)
        Case "level4": Call ApplyLevel4Style(para)
        Case "level5": Call ApplyLevel5Style(para)
        Case "level6": Call ApplyLevel6Style(para)
        Case "table_title": Call ApplyTableTitleStyle(para)
        Case "figure_title": Call ApplyFigureTitleStyle(para)
        Case Else: Call ApplyBodyStyle(para)
    End Select
End Sub

Private Function DetectLevel(txt As String) As String
    Dim firstChar As String, secondChar As String
    Dim cnNumbers As String, dunHao As String, fullDot As String, lBracket As String
    
    ' 中文数字：一二三四五六七八九十
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
    
    ' 表格标题：以"表"开头
    If firstChar = ChrW(&H8868) Then DetectLevel = "table_title": Exit Function
    ' 图片标题：以"图"开头
    If firstChar = ChrW(&H56FE) Then DetectLevel = "figure_title": Exit Function
    
    ' 二级标题：一、二、三... + 顿号
    If InStr(cnNumbers, firstChar) > 0 And InStr(txt, dunHao) > 0 And InStr(txt, dunHao) <= 3 Then
        DetectLevel = "level2": Exit Function
    End If
    
    ' 三级标题：（一）（二）... 全角括号 + 中文数字
    If (firstChar = lBracket Or firstChar = "(") And InStr(cnNumbers, secondChar) > 0 Then
        DetectLevel = "level3": Exit Function
    End If
    
    ' 四级标题：1. 2. 3.... 数字 + 全角点
    If IsNumeric(firstChar) And InStr(txt, fullDot) > 0 And InStr(txt, fullDot) <= 3 Then
        DetectLevel = "level4": Exit Function
    End If
    
    ' 五级标题：(1) (2)... 括号 + 阿拉伯数字
    If (firstChar = lBracket Or firstChar = "(") And IsNumeric(secondChar) Then
        DetectLevel = "level5": Exit Function
    End If
    
    ' 六级标题：带圈数字
    If IsCircledNumber(firstChar) Then DetectLevel = "level6": Exit Function
    
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
        ' 根据格式模式设置缩进
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
    indentValue = CentimetersToPoints(0.85) * 2  ' 2字符

    With para.Range.Font
        .NameFarEast = GetFont("楷体_GB2312", "楷体", "华文楷体")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_30
        .SpaceBefore = 8: .SpaceAfter = 8
        ' 根据格式模式设置缩进
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
    indentValue = CentimetersToPoints(0.85) * 2  ' 2字符

    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8: .SpaceAfter = 8
        ' 根据格式模式设置缩进
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
    indentValue = CentimetersToPoints(0.85) * 2  ' 2字符

    With para.Range.Font
        .NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
        .NameAscii = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 0: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
        .SpaceBefore = 8: .SpaceAfter = 8
        ' 根据格式模式设置缩进
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
        .Alignment = 3: .LineSpacingRule = 4: .LineSpacing = LINE_SPACING_28
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
    Dim sec As Section, ftr As HeaderFooter, rng As Range
    
    On Error Resume Next
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    
    For Each sec In ActiveDocument.Sections
        Set ftr = sec.Footers(1)
        ftr.Range.Delete
        
        Set rng = ftr.Range
        rng.InsertAfter ChrW(&H2014) & " "
        
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        ftr.Range.Fields.Add Range:=rng, Type:=33
        
        Set rng = ftr.Range
        rng.Collapse Direction:=0
        rng.InsertAfter " " & ChrW(&H2014)
        
        With ftr.Range
            .ParagraphFormat.Alignment = 1
            .Font.Name = "宋体"
            .Font.Size = FONT_SIZE_SI
        End With
        
        ftr.Range.Fields.Update
    Next sec
    
    ActiveDocument.Fields.Update
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

Public Sub FormatAllTables()
    Dim tbl As Table, cel As Cell, para As Paragraph
    On Error Resume Next
    For Each tbl In ActiveDocument.Tables
        tbl.Rows.Alignment = 1
        For Each cel In tbl.Range.Cells
            For Each para In cel.Range.Paragraphs
                para.Range.Font.NameFarEast = GetFont("仿宋_GB2312", "仿宋", "华文仿宋")
                para.Range.Font.NameAscii = "Times New Roman"
                para.Range.Font.Size = FONT_SIZE_XIAOSI
                para.Format.Alignment = 1
            Next para
        Next cel
    Next tbl
    MsgBox "表格格式化完成！", vbInformation, "表格格式化"
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
' 快速样式应用
'==============================================================================

Public Sub ApplyLevel1ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim p As Paragraph: For Each p In Selection.Paragraphs: ApplyLevel1Style p: Next
End Sub

Public Sub ApplyLevel2ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim p As Paragraph: For Each p In Selection.Paragraphs: ApplyLevel2Style p: Next
End Sub

Public Sub ApplyLevel3ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim p As Paragraph: For Each p In Selection.Paragraphs: ApplyLevel3Style p: Next
End Sub

Public Sub ApplyLevel4ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim p As Paragraph: For Each p In Selection.Paragraphs: ApplyLevel4Style p: Next
End Sub

Public Sub ApplyLevel5ToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim p As Paragraph: For Each p In Selection.Paragraphs: ApplyLevel5Style p: Next
End Sub

Public Sub ApplyBodyToSelection()
    g_FormatMode = SelectFormatMode()
    If g_FormatMode = "" Then Exit Sub
    Dim p As Paragraph: For Each p In Selection.Paragraphs: ApplyBodyStyle p: Next
End Sub

'==============================================================================
' 中文别名（方便调用）
'==============================================================================

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
