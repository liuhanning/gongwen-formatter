Attribute VB_Name = "ResearchProjectFormatter"
'==============================================================================
' Research Project Formatter (Word/WPS Compatible)
'==============================================================================

Option Explicit

Private Const PAGE_MARGIN_TOP_BOTTOM As Single = 70.875
Private Const PAGE_MARGIN_LEFT_RIGHT As Single = 76.545
Private Const HEADER_DISTANCE As Single = 51.03
Private Const FOOTER_DISTANCE As Single = 51.03

Private Const FONT_SIZE_ER As Single = 22
Private Const FONT_SIZE_SAN As Single = 16
Private Const FONT_SIZE_SI As Single = 14
Private Const FONT_SIZE_XIAOSI As Single = 12

Private Const LINE_SPACING_35 As Single = 35
Private Const LINE_SPACING_31 As Single = 31

Public Sub FormatResearchProject()
    Dim undoRec As Object
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then undoRec.StartCustomRecord "Format"
    On Error GoTo ErrorHandler
    
    Call SetupPage
    Call ProcessDocumentParts
    Call FormatAllTables
    Call ReplaceSymbolsSkippingTables
    Call AddResearchPageNumber
    
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    
    MsgBox "Format Complete!", vbInformation, "Done"
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub SetupPage()
    With ActiveDocument.PageSetup
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .TopMargin = PAGE_MARGIN_TOP_BOTTOM
        .BottomMargin = PAGE_MARGIN_TOP_BOTTOM
        .LeftMargin = PAGE_MARGIN_LEFT_RIGHT
        .RightMargin = PAGE_MARGIN_LEFT_RIGHT
        .HeaderDistance = HEADER_DISTANCE
        .FooterDistance = FOOTER_DISTANCE
        .OddAndEvenPagesHeaderFooter = True
    End With
End Sub

Private Sub ProcessDocumentParts()
    Dim para As Paragraph
    Dim i As Long, total As Long
    Dim txt As String
    Dim isAfterTitle As Boolean: isAfterTitle = False
    Dim isAfterAbstract As Boolean: isAfterAbstract = False
    Dim isFirstPara As Boolean: isFirstPara = True
    Dim abstractLabel As String
    Dim refLabel As String
    
    abstractLabel = ChrW(&H6458) & ChrW(&H8981) & ChrW(&HFF1A)
    refLabel = ChrW(&H53C2) & ChrW(&H8003) & ChrW(&H6587) & ChrW(&H732E) & ChrW(&HFF1A)
    
    total = ActiveDocument.Paragraphs.Count
    
    For i = 1 To total
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(para.Range.Text)
        txt = Replace(Replace(txt, vbCr, ""), vbLf, "")
        
        If Len(txt) > 0 Then
            If Not para.Range.Information(12) Then
                If isFirstPara Then
                    Call ApplyTitleStyle(para)
                    isFirstPara = False
                    isAfterTitle = True
                ElseIf isAfterTitle And Not isAfterAbstract And para.Format.Alignment = 1 Then
                    Call ApplyUnitNameStyle(para)
                ElseIf InStr(txt, abstractLabel) = 1 Then
                    Call ApplyAbstractStyle(para)
                    isAfterAbstract = True
                ElseIf InStr(txt, refLabel) = 1 Then
                    Call ApplyReferenceLabelStyle(para)
                ElseIf Left(txt, 1) = "[" And Len(txt) > 1 And IsNumeric(Mid(txt, 2, 1)) Then
                    Call ApplyReferenceItemStyle(para)
                Else
                    Call ProcessNormalOrHeading(para)
                End If
            End If
        End If
        
        If i Mod 50 = 0 Then Application.StatusBar = i & "/" & total
    Next i
    
    Application.StatusBar = ""
End Sub

Private Sub ApplyTitleStyle(para As Paragraph)
    Dim fzXbs As String, hwZs As String, st As String
    fzXbs = ChrW(&H65B9) & ChrW(&H6B63) & ChrW(&H5C0F) & ChrW(&H6807) & ChrW(&H5B8B) & ChrW(&H7B80) & ChrW(&H4F53)
    hwZs = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H4E2D) & ChrW(&H5B8B)
    st = ChrW(&H5B8B) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(fzXbs, hwZs, st)
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_ER
        .Bold = False
    End With
    With para.Format
        .Alignment = 1
        .LineSpacingRule = 4
        .LineSpacing = LINE_SPACING_35
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
    End With
End Sub

Private Sub ApplyUnitNameStyle(para As Paragraph)
    Dim kt As String, hwKt As String
    kt = ChrW(&H6977) & ChrW(&H4F53) & "_GB2312"
    hwKt = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H6977) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(kt, ChrW(&H6977) & ChrW(&H4F53), hwKt)
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    para.Format.Alignment = 1
End Sub

Private Sub ApplyAbstractStyle(para As Paragraph)
    Dim rng As Range
    Dim splitPos As Long
    Dim kt As String, hwKt As String, ht As String, st As String
    
    kt = ChrW(&H6977) & ChrW(&H4F53) & "_GB2312"
    hwKt = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H6977) & ChrW(&H4F53)
    ht = ChrW(&H9ED1) & ChrW(&H4F53)
    st = ChrW(&H5B8B) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(kt, ChrW(&H6977) & ChrW(&H4F53), hwKt)
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    
    splitPos = InStr(para.Range.Text, ChrW(&HFF1A))
    If splitPos = 0 Then splitPos = InStr(para.Range.Text, ":")
    If splitPos > 0 Then
        Set rng = ActiveDocument.Range(para.Range.Start, para.Range.Start + splitPos)
        rng.Font.NameFarEast = GetFont(ht, ChrW(&H5FAE) & ChrW(&H8F6F) & ChrW(&H96C5) & ChrW(&H9ED1), st)
    End If
    
    With para.Format
        .LineSpacingRule = 4
        .LineSpacing = LINE_SPACING_31
        .FirstLineIndent = CentimetersToPoints(0.85) * 2
        .Alignment = 3
    End With
End Sub

Private Sub ProcessNormalOrHeading(para As Paragraph)
    Dim txt As String, level As String
    txt = Trim(para.Range.Text)
    level = DetectLevel(txt)
    
    Select Case level
        Case "level1": Call ApplyHeadLevel1(para)
        Case "level2": Call ApplyHeadLevel2(para)
        Case "level3": Call ApplyHeadLevel3(para)
        Case "level4": Call ApplyHeadLevel4(para)
        Case "table_title": Call ApplyTableTitleStyle(para)
        Case "figure_title": Call ApplyFigureTitleStyle(para)
        Case Else: Call ApplyBodyStyle(para)
    End Select
    
    Call FormatMixedText(para)
End Sub

Private Sub ApplyHeadLevel1(para As Paragraph)
    Dim ht As String, st As String
    ht = ChrW(&H9ED1) & ChrW(&H4F53)
    st = ChrW(&H5B8B) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(ht, ChrW(&H5FAE) & ChrW(&H8F6F) & ChrW(&H96C5) & ChrW(&H9ED1), st)
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    CommonHeadingFormat para
End Sub

Private Sub ApplyHeadLevel2(para As Paragraph)
    Dim kt As String, hwKt As String
    kt = ChrW(&H6977) & ChrW(&H4F53) & "_GB2312"
    hwKt = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H6977) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(kt, ChrW(&H6977) & ChrW(&H4F53), hwKt)
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    CommonHeadingFormat para
End Sub

Private Sub ApplyHeadLevel3(para As Paragraph)
    Dim fs As String, hwFs As String
    fs = ChrW(&H4EFF) & ChrW(&H5B8B) & "_GB2312"
    hwFs = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H4EFF) & ChrW(&H5B8B)
    
    With para.Range.Font
        .NameFarEast = GetFont(fs, ChrW(&H4EFF) & ChrW(&H5B8B), hwFs)
        .Size = FONT_SIZE_SAN
        .Bold = True
    End With
    CommonHeadingFormat para
End Sub

Private Sub ApplyHeadLevel4(para As Paragraph)
    Dim fs As String, hwFs As String
    fs = ChrW(&H4EFF) & ChrW(&H5B8B) & "_GB2312"
    hwFs = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H4EFF) & ChrW(&H5B8B)
    
    With para.Range.Font
        .NameFarEast = GetFont(fs, ChrW(&H4EFF) & ChrW(&H5B8B), hwFs)
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    CommonHeadingFormat para
End Sub

Private Sub CommonHeadingFormat(para As Paragraph)
    With para.Format
        .Alignment = 0
        .FirstLineIndent = CentimetersToPoints(0.85) * 2
        .LineSpacingRule = 4
        .LineSpacing = LINE_SPACING_31
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub

Private Sub ApplyBodyStyle(para As Paragraph)
    Dim fs As String, hwFs As String
    fs = ChrW(&H4EFF) & ChrW(&H5B8B) & "_GB2312"
    hwFs = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H4EFF) & ChrW(&H5B8B)
    
    With para.Range.Font
        .NameFarEast = GetFont(fs, ChrW(&H4EFF) & ChrW(&H5B8B), hwFs)
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    With para.Format
        .Alignment = 3
        .FirstLineIndent = CentimetersToPoints(0.85) * 2
        .LineSpacingRule = 4
        .LineSpacing = LINE_SPACING_31
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub

Private Sub ApplyTableTitleStyle(para As Paragraph)
    Dim ht As String, st As String
    ht = ChrW(&H9ED1) & ChrW(&H4F53)
    st = ChrW(&H5B8B) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(ht, ChrW(&H5FAE) & ChrW(&H8F6F) & ChrW(&H96C5) & ChrW(&H9ED1), st)
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
    End With
    With para.Format
        .Alignment = 1
        .SpaceBefore = 6
        .SpaceAfter = 6
    End With
End Sub

Private Sub ApplyFigureTitleStyle(para As Paragraph)
    ApplyTableTitleStyle para
End Sub

Private Sub ApplyReferenceLabelStyle(para As Paragraph)
    Dim ht As String, st As String
    ht = ChrW(&H9ED1) & ChrW(&H4F53)
    st = ChrW(&H5B8B) & ChrW(&H4F53)
    
    With para.Range.Font
        .NameFarEast = GetFont(ht, ChrW(&H5FAE) & ChrW(&H8F6F) & ChrW(&H96C5) & ChrW(&H9ED1), st)
        .Size = FONT_SIZE_SAN
        .Bold = False
    End With
    para.Format.SpaceBefore = 12
End Sub

Private Sub ApplyReferenceItemStyle(para As Paragraph)
    Dim fs As String, hwFs As String
    fs = ChrW(&H4EFF) & ChrW(&H5B8B) & "_GB2312"
    hwFs = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H4EFF) & ChrW(&H5B8B)
    
    With para.Range.Font
        .NameFarEast = GetFont(fs, ChrW(&H4EFF) & ChrW(&H5B8B), hwFs)
        .Size = FONT_SIZE_XIAOSI
        .Bold = False
    End With
    para.Format.FirstLineIndent = 0
End Sub

Private Function DetectLevel(txt As String) As String
    Dim firstChar As String, secondChar As String
    Dim cnNumbers As String, dunHao As String, fullDot As String, lBracket As String
    
    cnNumbers = ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09) & ChrW(&H56DB) & ChrW(&H4E94) & ChrW(&H516D) & ChrW(&H4E03) & ChrW(&H516B) & ChrW(&H4E5D) & ChrW(&H5341)
    dunHao = ChrW(&H3001)
    fullDot = ChrW(&HFF0E)
    lBracket = ChrW(&HFF08)
    
    If Len(txt) = 0 Then DetectLevel = "body": Exit Function
    
    firstChar = Left(txt, 1)
    If Len(txt) > 1 Then secondChar = Mid(txt, 2, 1) Else secondChar = ""
    
    If firstChar = ChrW(&H8868) Then DetectLevel = "table_title": Exit Function
    If firstChar = ChrW(&H56FE) Then DetectLevel = "figure_title": Exit Function
    
    If InStr(cnNumbers, firstChar) > 0 And secondChar = dunHao Then
        DetectLevel = "level1": Exit Function
    End If
    If (firstChar = lBracket Or firstChar = "(") And InStr(cnNumbers, secondChar) > 0 Then
        DetectLevel = "level2": Exit Function
    End If
    If IsNumeric(firstChar) And InStr(txt, fullDot) > 0 And InStr(txt, fullDot) <= 3 Then
        DetectLevel = "level3": Exit Function
    End If
    If (firstChar = lBracket Or firstChar = "(") And IsNumeric(secondChar) Then
        DetectLevel = "level4": Exit Function
    End If
    
    DetectLevel = "body"
End Function

Private Sub FormatMixedText(para As Paragraph)
    Dim rng As Range
    Dim txt As String
    Dim i As Long, segStart As Long
    Dim code As Long
    Dim inAscii As Boolean
    
    txt = para.Range.Text
    If Len(txt) = 0 Then Exit Sub
    
    inAscii = False
    segStart = 1
    
    For i = 1 To Len(txt)
        code = AscW(Mid(txt, i, 1))
        If (code >= 32 And code <= 126) Then
            If Not inAscii Then
                segStart = i
                inAscii = True
            End If
        Else
            If inAscii Then
                Set rng = ActiveDocument.Range(para.Range.Start + segStart - 1, para.Range.Start + i - 1)
                rng.Font.Name = "Times New Roman"
                inAscii = False
            End If
        End If
    Next i
    
    If inAscii Then
        Set rng = ActiveDocument.Range(para.Range.Start + segStart - 1, para.Range.End - 1)
        rng.Font.Name = "Times New Roman"
    End If
End Sub

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
    On Error GoTo 0
End Function

Private Sub AddResearchPageNumber()
    Dim sec As Section
    Dim ftr As HeaderFooter
    
    For Each sec In ActiveDocument.Sections
        Set ftr = sec.Footers(1)
        FormatPageNumberPara ftr, 2
        
        Set ftr = sec.Footers(3)
        FormatPageNumberPara ftr, 0
    Next sec
End Sub

Private Sub FormatPageNumberPara(ftr As HeaderFooter, align As Integer)
    Dim rng As Range
    Dim st As String
    st = ChrW(&H5B8B) & ChrW(&H4F53)
    
    ftr.Range.Delete
    Set rng = ftr.Range
    
    rng.Collapse 1
    rng.InsertAfter ChrW(&H2014) & " "
    
    Set rng = ftr.Range
    rng.Collapse 0
    ActiveDocument.Fields.Add Range:=rng, Type:=33
    
    Set rng = ftr.Range
    rng.Collapse 0
    rng.InsertAfter " " & ChrW(&H2014)
    
    With ftr.Range
        .ParagraphFormat.Alignment = align
        .Font.Name = st
        .Font.Size = FONT_SIZE_SI
    End With
End Sub

Private Sub ReplaceSymbolsSkippingTables()
    Call DoReplaceSub(",", ChrW(&HFF0C))
    Call DoReplaceSub("(", ChrW(&HFF08))
    Call DoReplaceSub(")", ChrW(&HFF09))
    Call DoReplaceSub(":", ChrW(&HFF1A))
End Sub

Private Sub DoReplaceSub(findWhat As String, replaceWith As String)
    Dim rng As Range
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .Text = findWhat
        .Forward = True
        .Wrap = 0
        Do While .Execute
            If Not rng.Information(12) Then rng.Text = replaceWith
            rng.Collapse 0
        Loop
    End With
End Sub

Private Sub FormatAllTables()
    Dim tbl As Table
    Dim r As Long, c As Long
    Dim cel As Cell
    Dim rng As Range
    Dim ht As String, fs As String, hwFs As String
    Dim txt As String
    Dim i As Long
    Dim code As Long
    
    ht = ChrW(&H9ED1) & ChrW(&H4F53)
    fs = ChrW(&H4EFF) & ChrW(&H5B8B) & "_GB2312"
    hwFs = ChrW(&H534E) & ChrW(&H6587) & ChrW(&H4EFF) & ChrW(&H5B8B)
    
    For Each tbl In ActiveDocument.Tables
        tbl.PreferredWidthType = 2
        tbl.PreferredWidth = 100
        
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                On Error Resume Next
                Set cel = tbl.Cell(r, c)
                If Err.Number = 0 Then
                    Set rng = cel.Range
                    rng.End = rng.End - 1
                    
                    If r = 1 Then
                        rng.Font.NameFarEast = GetFont(ht, ChrW(&H5FAE) & ChrW(&H8F6F) & ChrW(&H96C5) & ChrW(&H9ED1), ChrW(&H5B8B) & ChrW(&H4F53))
                        rng.Font.Size = FONT_SIZE_XIAOSI
                        rng.Font.Bold = False
                    Else
                        rng.Font.NameFarEast = GetFont(fs, ChrW(&H4EFF) & ChrW(&H5B8B), hwFs)
                        rng.Font.Size = FONT_SIZE_XIAOSI
                        rng.Font.Bold = False
                        
                        txt = rng.Text
                        Call FormatTableCellNumbers(cel, txt)
                    End If
                    
                    rng.ParagraphFormat.LineSpacingRule = 0
                    rng.ParagraphFormat.Alignment = 1
                End If
                On Error GoTo 0
            Next c
        Next r
    Next tbl
End Sub

Private Sub FormatTableCellNumbers(cel As Cell, txt As String)
    Dim rng As Range
    Dim i As Long, segStart As Long
    Dim code As Long
    Dim inNum As Boolean
    Dim basePos As Long
    
    If Len(txt) = 0 Then Exit Sub
    
    basePos = cel.Range.Start
    inNum = False
    segStart = 1
    
    For i = 1 To Len(txt)
        code = AscW(Mid(txt, i, 1))
        If (code >= 48 And code <= 57) Or code = 46 Or code = 45 Or code = 37 Then
            If Not inNum Then
                segStart = i
                inNum = True
            End If
        Else
            If inNum Then
                On Error Resume Next
                Set rng = ActiveDocument.Range(basePos + segStart - 1, basePos + i - 1)
                rng.Font.Name = "Times New Roman"
                On Error GoTo 0
                inNum = False
            End If
        End If
    Next i
    
    If inNum Then
        On Error Resume Next
        Set rng = ActiveDocument.Range(basePos + segStart - 1, basePos + Len(txt))
        rng.Font.Name = "Times New Roman"
        On Error GoTo 0
    End If
End Sub
