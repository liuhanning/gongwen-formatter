# 公文格式化工具 (Gongwen Document Formatter)

专业的 Word 文档自动化排版工具，符合 **GB/T 9704 党政机关公文格式规范**。

## ✨ 功能特点

- ✅ **完全符合规范**：严格遵循 GB/T 9704 标准的 7 项格式要求
- 🎨 **现代化界面**：精美的 Web 界面，支持拖拽上传
- 🚀 **一键格式化**：自动识别标题级别，智能应用格式
- 📝 **多种使用方式**：Web 界面、Python API、VBA 宏
- 🔧 **字体兼容**：自动检测并退化到可用字体

## 📋 格式规范（7项要求）

### 1. 页面设置
- A4 版幅（210mm × 297mm）
- 上边距：3.7cm，下边距：3.5cm
- 左边距：2.8cm，右边距：2.6cm
- 页眉页脚：均为 1.8cm

### 2. 标题格式
- **一级标题**：方正小标宋简体二号，30磅行距，居中
- **二级标题**：黑体三号，30磅行距，左缩进2字符
- **三级标题**：楷体_GB2312三号，30磅行距，左缩进2字符
- **四级及以下**：仿宋_GB2312三号，28磅行距，左缩进2字符

### 3. 标题序号
依次使用：`一、` → `（一）` → `1．` → `（1）` → `①`
> ⚠️ 注意：`1．` 使用全角点 `．` 而非半角 `.`

### 4. 正文格式
- 仿宋_GB2312三号，不加粗
- 固定28磅行距，段前段后0行
- 两端对齐，首行缩进2字符
- 英文使用 Times New Roman

### 5. 页眉页码
- 页眉：仿宋_GB2312
- 页码：宋体四号，居中，格式为 **— 1 —**（使用破折号包围）

### 6. 表格格式
- 标题：黑体小四，不加粗，居中
- 表头：黑体小四，不加粗
- 内容：仿宋_GB2312小四，数字用 Times New Roman
- 单倍行距，表格前后各空一行

### 7. 插图格式
- 标题：黑体小四，不加粗，居中
- 位于图下方，前后各空一行

## 🚀 快速开始

### 方法一：Web 界面（推荐）

1. **安装依赖**
```bash
pip install flask flask-cors python-docx
```

2. **启动服务器**
```bash
python web_server.py
```

3. **打开浏览器**
访问 `http://localhost:5000`

4. **使用界面**
- 拖拽或点击上传 `.docx` 文件
- 点击"格式化文档"按钮
- 自动下载格式化后的文档

### 方法二：Python 命令行

1. **安装依赖**
```bash
pip install python-docx
```

2. **格式化已有文档**
```bash
python gongwen_formatter.py input.docx output.docx
```

3. **创建示例文档**
```bash
python gongwen_formatter.py
```

### 方法三：VBA 宏（Word 内使用）

1. 打开 Word 文档
2. 按 `Alt + F11` 打开 VBA 编辑器
3. 导入 `GongwenFormatter.bas` 模块
4. 运行 `FormatGongwen` 宏

## 📁 文件说明

```
wordhong/
├── index.html              # Web 界面主页
├── styles.css              # 界面样式（现代化设计）
├── script.js               # 前端交互逻辑
├── web_server.py           # Flask 后端服务器
├── gongwen_formatter.py    # Python 格式化引擎
├── GongwenFormatter.bas    # VBA 宏（已修复语法）
├── demo_gongwen.docx       # 示例文档
└── README.md               # 本文件
```

## 🔧 Python API 使用

```python
from gongwen_formatter import GongwenFormatter

# 创建格式化器
formatter = GongwenFormatter()

# 添加页码和页眉
formatter.add_page_number()
formatter.add_header('文档标题')

# 添加各级标题
formatter.add_heading('一级标题', 1)
formatter.add_heading('一、二级标题', 2)
formatter.add_heading('（一）三级标题', 3)
formatter.add_heading('1．四级标题', 4)

# 添加正文
formatter.add_body_paragraph('这是正文内容...')

# 添加表格
table_data = [
    ['序号', '项目', '数值'],
    ['1', '测试', '100']
]
formatter.add_table('表1 示例表格', 2, 3, table_data)

# 保存文档
formatter.save('output.docx')
```

## 🐛 BAS 语法修复说明

已修复 `GongwenFormatter.bas` 中的语法错误：
- **第 525 行**：`Dim cell As cell` → `Dim Cell As Cell`
- **第 531-543 行**：所有 `cell` 变量引用已更正为 `Cell`

VBA 对类型名称大小写敏感，`Cell` 是正确的类型名称。

## 🎨 Web 界面特色

- 🌙 **深色主题**：现代化渐变配色
- ✨ **流畅动画**：微交互提升体验
- 📱 **响应式设计**：支持各种屏幕尺寸
- 🎯 **拖拽上传**：直观的文件上传方式
- 📊 **实时进度**：格式化进度可视化
- ⌨️ **快捷键支持**：
  - `Ctrl/Cmd + O`：打开文件
  - `Ctrl/Cmd + Enter`：格式化文档

## 📝 注意事项

1. **字体要求**：
   - 方正小标宋简体（一级标题）
   - 黑体（二级标题、表格）
   - 楷体_GB2312（三级标题）
   - 仿宋_GB2312（四级标题、正文）
   - Times New Roman（英文、数字）

2. **字体缺失处理**：
   - 程序会自动检测系统字体
   - 如字体缺失，会自动退化到相似字体
   - 建议安装所有必需字体以获得最佳效果

3. **序号格式**：
   - 使用静态文本，不使用 Word 自动编号
   - `1．` 必须使用全角点 `．`（Unicode FF0E）
   - 双引号与段落字体保持一致

## 🔍 技术栈

- **前端**：HTML5 + CSS3 + Vanilla JavaScript
- **后端**：Python 3.7+ + Flask
- **文档处理**：python-docx
- **VBA**：Microsoft Word VBA

## 📄 许可证

本项目仅供学习和内部使用。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

**开发者**: Antigravity AI  
**版本**: v1.0  
**最后更新**: 2026-02-04
