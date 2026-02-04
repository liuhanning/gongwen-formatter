' ============================================================================
' 公文格式化工具 - 使用说明
' GB/T 9704 党政机关公文格式规范
' ============================================================================

Option Explicit

Dim objShell, objFSO, strScriptPath, strWorkDir
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 获取脚本所在目录
strScriptPath = WScript.ScriptFullName
strWorkDir = objFSO.GetParentFolderName(strScriptPath)

' 显示主菜单
ShowMainMenu()

' ============================================================================
' 主菜单
' ============================================================================
Sub ShowMainMenu()
    Dim strMenu, intChoice
    
    strMenu = "═══════════════════════════════════════" & vbCrLf & _
              "    公文格式化工具 - 使用说明" & vbCrLf & _
              "    GB/T 9704 标准" & vbCrLf & _
              "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
              "请选择操作：" & vbCrLf & vbCrLf & _
              "【1】查看 7 项格式要求" & vbCrLf & _
              "【2】Python 使用方法" & vbCrLf & _
              "【3】VBA 宏使用方法" & vbCrLf & _
              "【4】Web 界面使用方法" & vbCrLf & _
              "【5】启动 Web 服务器" & vbCrLf & _
              "【6】生成示例文档" & vbCrLf & _
              "【7】打开项目文件夹" & vbCrLf & _
              "【0】退出" & vbCrLf & vbCrLf & _
              "请输入选项编号 (0-7)："
    
    intChoice = InputBox(strMenu, "公文格式化工具", "1")
    
    If intChoice = "" Then
        Exit Sub
    End If
    
    Select Case CInt(intChoice)
        Case 1
            ShowFormatRequirements()
        Case 2
            ShowPythonUsage()
        Case 3
            ShowVBAUsage()
        Case 4
            ShowWebUsage()
        Case 5
            StartWebServer()
        Case 6
            GenerateDemoDocument()
        Case 7
            OpenProjectFolder()
        Case 0
            Exit Sub
        Case Else
            MsgBox "无效的选项！", vbExclamation, "错误"
            ShowMainMenu()
    End Select
End Sub

' ============================================================================
' 显示 7 项格式要求
' ============================================================================
Sub ShowFormatRequirements()
    Dim strContent
    
    strContent = "═══════════════════════════════════════" & vbCrLf & _
                 "  GB/T 9704 公文格式 - 7 项要求" & vbCrLf & _
                 "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
                 "【1】页面设置" & vbCrLf & _
                 "    • A4 版幅 (210mm × 297mm)" & vbCrLf & _
                 "    • 上边距：3.7cm，下边距：3.5cm" & vbCrLf & _
                 "    • 左边距：2.8cm，右边距：2.6cm" & vbCrLf & _
                 "    • 页眉页脚：均为 1.8cm" & vbCrLf & vbCrLf & _
                 "【2】标题格式" & vbCrLf & _
                 "    • 一级：方正小标宋简体二号，30磅，居中" & vbCrLf & _
                 "    • 二级：黑体三号，30磅，左缩进2字符" & vbCrLf & _
                 "    • 三级：楷体_GB2312三号，30磅" & vbCrLf & _
                 "    • 四级+：仿宋_GB2312三号，28磅" & vbCrLf & vbCrLf & _
                 "【3】标题序号" & vbCrLf & _
                 "    依次使用：一、→ （一）→ 1．→ （1）→ ①" & vbCrLf & _
                 "    ⚠ 注意：1．使用全角点 ．" & vbCrLf & vbCrLf & _
                 "【4】正文格式" & vbCrLf & _
                 "    • 仿宋_GB2312三号，不加粗" & vbCrLf & _
                 "    • 固定28磅行距，段前段后0行" & vbCrLf & _
                 "    • 两端对齐，首行缩进2字符" & vbCrLf & vbCrLf & _
                 "【5】页眉页码" & vbCrLf & _
                 "    • 页眉：仿宋_GB2312" & vbCrLf & _
                 "    • 页码：宋体四号，居中，格式 — 1 —" & vbCrLf & vbCrLf & _
                 "【6】表格格式" & vbCrLf & _
                 "    • 标题：黑体小四，不加粗，居中" & vbCrLf & _
                 "    • 表头：黑体小四，不加粗" & vbCrLf & _
                 "    • 内容：仿宋_GB2312小四" & vbCrLf & _
                 "    • 单倍行距，表格前后各空一行" & vbCrLf & vbCrLf & _
                 "【7】插图格式" & vbCrLf & _
                 "    • 标题：黑体小四，不加粗，居中" & vbCrLf & _
                 "    • 位于图下方，前后各空一行" & vbCrLf & vbCrLf & _
                 "═══════════════════════════════════════"
    
    MsgBox strContent, vbInformation, "7 项格式要求"
    ShowMainMenu()
End Sub

' ============================================================================
' Python 使用方法
' ============================================================================
Sub ShowPythonUsage()
    Dim strContent
    
    strContent = "═══════════════════════════════════════" & vbCrLf & _
                 "  Python 命令行使用方法" & vbCrLf & _
                 "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
                 "【方法一】格式化已有文档" & vbCrLf & _
                 "  python gongwen_formatter.py input.docx output.docx" & vbCrLf & vbCrLf & _
                 "【方法二】生成示例文档" & vbCrLf & _
                 "  python gongwen_formatter.py" & vbCrLf & vbCrLf & _
                 "【方法三】使用 Python API" & vbCrLf & _
                 "  from gongwen_formatter import GongwenFormatter" & vbCrLf & _
                 "  formatter = GongwenFormatter()" & vbCrLf & _
                 "  formatter.add_page_number()" & vbCrLf & _
                 "  formatter.add_heading('标题', 1)" & vbCrLf & _
                 "  formatter.save('output.docx')" & vbCrLf & vbCrLf & _
                 "【依赖安装】" & vbCrLf & _
                 "  pip install python-docx" & vbCrLf & vbCrLf & _
                 "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
                 "是否现在生成示例文档？"
    
    Dim intResult
    intResult = MsgBox(strContent, vbQuestion + vbYesNo, "Python 使用方法")
    
    If intResult = vbYes Then
        GenerateDemoDocument()
    Else
        ShowMainMenu()
    End If
End Sub

' ============================================================================
' VBA 宏使用方法
' ============================================================================
Sub ShowVBAUsage()
    Dim strContent
    
    strContent = "═══════════════════════════════════════" & vbCrLf & _
                 "  VBA 宏使用方法" & vbCrLf & _
                 "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
                 "【步骤 1】打开 Word 文档" & vbCrLf & _
                 "  打开需要格式化的 Word 文档" & vbCrLf & vbCrLf & _
                 "【步骤 2】打开 VBA 编辑器" & vbCrLf & _
                 "  按 Alt + F11 组合键" & vbCrLf & vbCrLf & _
                 "【步骤 3】导入模块" & vbCrLf & _
                 "  文件 → 导入文件 → 选择 GongwenFormatter.bas" & vbCrLf & vbCrLf & _
                 "【步骤 4】运行宏" & vbCrLf & _
                 "  • FormatGongwen - 格式化整个文档" & vbCrLf & _
                 "  • FormatSelectedParagraphs - 格式化选中段落" & vbCrLf & _
                 "  • FormatAllTables - 格式化所有表格" & vbCrLf & _
                 "  • ApplyLevel1ToSelection - 应用一级标题" & vbCrLf & _
                 "  • ApplyLevel2ToSelection - 应用二级标题" & vbCrLf & _
                 "  • ApplyBodyToSelection - 应用正文格式" & vbCrLf & vbCrLf & _
                 "【快捷方式】" & vbCrLf & _
                 "  可以为常用宏添加快捷键或按钮" & vbCrLf & vbCrLf & _
                 "【注意事项】" & vbCrLf & _
                 "  • 建议先备份原文档" & vbCrLf & _
                 "  • 确保已安装必需字体" & vbCrLf & _
                 "  • 页码格式：— 1 —（已修复）" & vbCrLf & vbCrLf & _
                 "═══════════════════════════════════════"
    
    MsgBox strContent, vbInformation, "VBA 宏使用方法"
    ShowMainMenu()
End Sub

' ============================================================================
' Web 界面使用方法
' ============================================================================
Sub ShowWebUsage()
    Dim strContent
    
    strContent = "═══════════════════════════════════════" & vbCrLf & _
                 "  Web 界面使用方法（推荐）" & vbCrLf & _
                 "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
                 "【步骤 1】启动服务器" & vbCrLf & _
                 "  双击运行 '启动服务器.bat'" & vbCrLf & _
                 "  或在主菜单选择【5】启动 Web 服务器" & vbCrLf & vbCrLf & _
                 "【步骤 2】打开浏览器" & vbCrLf & _
                 "  访问 http://localhost:5000" & vbCrLf & vbCrLf & _
                 "【步骤 3】上传文档" & vbCrLf & _
                 "  • 拖拽 .docx 文件到上传区域" & vbCrLf & _
                 "  • 或点击上传区域选择文件" & vbCrLf & vbCrLf & _
                 "【步骤 4】格式化" & vbCrLf & _
                 "  点击 '格式化文档' 按钮" & vbCrLf & vbCrLf & _
                 "【步骤 5】下载" & vbCrLf & _
                 "  自动下载格式化后的文档" & vbCrLf & vbCrLf & _
                 "【界面特色】" & vbCrLf & _
                 "  • 现代化深色主题" & vbCrLf & _
                 "  • 拖拽上传支持" & vbCrLf & _
                 "  • 实时进度显示" & vbCrLf & _
                 "  • 快捷键：Ctrl+O 打开，Ctrl+Enter 格式化" & vbCrLf & vbCrLf & _
                 "【依赖安装】" & vbCrLf & _
                 "  pip install flask flask-cors python-docx" & vbCrLf & vbCrLf & _
                 "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
                 "是否现在启动 Web 服务器？"
    
    Dim intResult
    intResult = MsgBox(strContent, vbQuestion + vbYesNo, "Web 界面使用方法")
    
    If intResult = vbYes Then
        StartWebServer()
    Else
        ShowMainMenu()
    End If
End Sub

' ============================================================================
' 启动 Web 服务器
' ============================================================================
Sub StartWebServer()
    Dim strBatFile, strCmd
    
    ' 检查批处理文件是否存在
    strBatFile = strWorkDir & "\启动服务器.bat"
    
    If objFSO.FileExists(strBatFile) Then
        MsgBox "正在启动 Web 服务器..." & vbCrLf & vbCrLf & _
               "服务器启动后，浏览器将自动打开" & vbCrLf & _
               "访问地址：http://localhost:5000" & vbCrLf & vbCrLf & _
               "按 Ctrl+C 可停止服务器", vbInformation, "启动 Web 服务器"
        
        ' 启动批处理文件
        objShell.Run """" & strBatFile & """", 1, False
        
        ' 等待 3 秒后打开浏览器
        WScript.Sleep 3000
        objShell.Run "http://localhost:5000"
    Else
        ' 直接运行 Python 命令
        MsgBox "正在启动 Web 服务器..." & vbCrLf & vbCrLf & _
               "访问地址：http://localhost:5000" & vbCrLf & _
               "按 Ctrl+C 可停止服务器", vbInformation, "启动 Web 服务器"
        
        objShell.CurrentDirectory = strWorkDir
        objShell.Run "cmd /k python web_server.py", 1, False
        
        ' 等待 3 秒后打开浏览器
        WScript.Sleep 3000
        objShell.Run "http://localhost:5000"
    End If
End Sub

' ============================================================================
' 生成示例文档
' ============================================================================
Sub GenerateDemoDocument()
    Dim strCmd, intResult
    
    MsgBox "正在生成示例文档..." & vbCrLf & vbCrLf & _
           "示例文档将包含：" & vbCrLf & _
           "• 各级标题示例（一至五级）" & vbCrLf & _
           "• 正文格式示例" & vbCrLf & _
           "• 表格格式示例" & vbCrLf & _
           "• 插图标题示例" & vbCrLf & _
           "• 页码格式：— 1 —" & vbCrLf & vbCrLf & _
           "请稍候...", vbInformation, "生成示例文档"
    
    ' 切换到工作目录并运行 Python
    objShell.CurrentDirectory = strWorkDir
    intResult = objShell.Run("python gongwen_formatter.py", 1, True)
    
    If intResult = 0 Then
        Dim strDemoFile
        strDemoFile = strWorkDir & "\demo_gongwen.docx"
        
        If objFSO.FileExists(strDemoFile) Then
            intResult = MsgBox("示例文档生成成功！" & vbCrLf & vbCrLf & _
                              "文件位置：" & vbCrLf & _
                              strDemoFile & vbCrLf & vbCrLf & _
                              "是否现在打开文档？", _
                              vbQuestion + vbYesNo, "生成成功")
            
            If intResult = vbYes Then
                objShell.Run """" & strDemoFile & """"
            End If
        Else
            MsgBox "示例文档生成失败！" & vbCrLf & _
                   "请检查 Python 环境和依赖包。", vbExclamation, "生成失败"
        End If
    Else
        MsgBox "执行失败！" & vbCrLf & vbCrLf & _
               "可能原因：" & vbCrLf & _
               "• Python 未安装或未添加到 PATH" & vbCrLf & _
               "• 缺少 python-docx 依赖包" & vbCrLf & vbCrLf & _
               "请运行：pip install python-docx", vbExclamation, "执行失败"
    End If
    
    ShowMainMenu()
End Sub

' ============================================================================
' 打开项目文件夹
' ============================================================================
Sub OpenProjectFolder()
    objShell.Run "explorer.exe """ & strWorkDir & """", 1, False
    
    MsgBox "已打开项目文件夹：" & vbCrLf & vbCrLf & _
           strWorkDir & vbCrLf & vbCrLf & _
           "主要文件：" & vbCrLf & _
           "• gongwen_formatter.py - Python 引擎" & vbCrLf & _
           "• GongwenFormatter.bas - VBA 宏" & vbCrLf & _
           "• web_server.py - Web 服务器" & vbCrLf & _
           "• index.html - Web 界面" & vbCrLf & _
           "• 启动服务器.bat - 快速启动" & vbCrLf & _
           "• README.md - 详细文档", vbInformation, "项目文件夹"
    
    ShowMainMenu()
End Sub

' ============================================================================
' 清理对象
' ============================================================================
Set objShell = Nothing
Set objFSO = Nothing
