# 公文格式化工具 (Gongwen Document Formatter)

专业的 Word 文档自动化排版工具，符合 **GB/T 9704 党政机关公文格式规范** 和 **政府交付版格式标准**。

## ✨ 功能特点

- ✅ **双模式支持**：GB/T 9704标准模式 + 政府交付版模式
- ✅ **完全符合规范**：严格遵循标准的 7 项格式要求
- 🎨 **现代化界面**：精美的 Web 界面，支持拖拽上传
- 🚀 **一键格式化**：自动识别标题级别，智能应用格式
- 📝 **多种使用方式**：Web 界面、Python API、VBA 宏
- 🔧 **字体兼容**：自动检测并退化到可用字体
- 💼 **WPS兼容**：完美支持 WPS Office 和 Microsoft Word

## 🎯 格式模式说明

### 模式一：GB/T 9704标准模式（默认）
- **标题缩进**：二、三、四级标题使用**左缩进2字符**
- **适用场景**：传统公文格式，符合GB/T 9704-2012标准

### 模式二：政府交付版模式
- **标题缩进**：二、三、四级标题使用**首行缩进2字符**
- **适用场景**：政府部门交付文档，咨询报告等
- **新增要求**：文档网格对齐

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
# 使用GB/T 9704标准模式（默认）
python gongwen_formatter.py input.docx output.docx

# 使用政府交付版模式
python gongwen_formatter.py input.docx output.docx --government
# 或使用简写
python gongwen_formatter.py input.docx output.docx -g
```

3. **创建示例文档**
```bash
# 标准模式
python gongwen_formatter.py

# 政府交付版模式
python gongwen_formatter.py --government
```

### 方法三：VBA 宏（Word/WPS 内使用）

#### 支持的软件
- ✅ Microsoft Word 2010 或更高版本
- ✅ WPS Office 2019 或更高版本

#### 详细使用步骤

**第一步：准备工作**

**对于 Microsoft Word：**
1. 确认 Word 版本为 2010 或更高版本
2. 确保已启用宏功能（文件 → 选项 → 信任中心 → 宏设置）

**对于 WPS Office：**
1. 确认 WPS 版本为 2019 或更高版本
2. 启用宏功能（文件 → 选项 → 安全性 → 宏安全性）
3. 勾选"启用所有宏"或"禁用无数字签名的宏"

**第二步：选择合适的代码文件**

| 软件 | 推荐文件 | 备选文件 |
|------|---------|---------|
| Microsoft Word | `GongwenFormatter.bas` | `GongwenFormatter_WPS.bas` |
| WPS Office | `GongwenFormatter_WPS.bas` | `GongwenFormatter.bas` |

> 💡 **说明**：`GongwenFormatter_WPS.bas` 是专门优化的WPS兼容版本，在WPS中运行更稳定

**第三步：导入 VBA 代码**
```
方式一：直接复制粘贴（推荐）
1. 按 Alt + F11 打开 VBA 编辑器
2. 插入 → 模块
3. 复制 GongwenFormatter.bas 全部内容
4. 粘贴到代码窗口

方式二：导入文件
1. 按 Alt + F11 打开 VBA 编辑器
2. 文件 → 导入文件
3. 选择 GongwenFormatter.bas
4. 点击打开
```

**第三步：运行宏命令**

| 宏名称 | 功能说明 | 使用场景 |
|--------|----------|----------|
| `FormatGongwen` | 格式化整个文档 | 完整文档格式化 |
| `FormatSelectedParagraphs` | 格式化选中段落 | 局部格式化 |
| `FormatAllTables` | 格式化所有表格 | 表格专项处理 |
| `ApplyLevel1ToSelection` | 应用一级标题样式 | 手动指定标题 |
| `ApplyLevel2ToSelection` | 应用二级标题样式 | 手动指定标题 |
| `ApplyLevel3ToSelection` | 应用三级标题样式 | 手动指定标题 |
| `ApplyLevel4ToSelection` | 应用四级标题样式 | 手动指定标题 |
| `ApplyLevel5ToSelection` | 应用五级标题样式 | 手动指定标题 |
| `ApplyBodyToSelection` | 应用正文样式 | 正文格式化 |
| `AddHeader("文字")` | 添加页眉 | 设置页眉内容 |
| `AddPageNumber` | 添加页码 | 设置页码格式 |

**第四步：快捷操作**
- 按 `Alt + F8` 打开宏对话框
- 选择对应宏名称
- 点击运行按钮

#### 标题识别规则

| 级别 | 格式要求 | 识别示例 |
|------|----------|----------|
| 一级 | 居中，无序号 | `标题文字` |
| 二级 | 左对齐，`一、`开头 | `一、标题内容` |
| 三级 | 左对齐，`（一）`开头 | `（一）标题内容` |
| 四级 | 左对齐，`1．`开头 | `1．标题内容` |
| 五级 | 左对齐，`（1）`开头 | `（1）标题内容` |
| 六级 | 左对齐，`①`开头 | `①标题内容` |

> ⚠️ 注意：四级标题必须使用全角点 `．`(U+FF0E)，不是半角点 `.`

#### 页面设置参数

- **纸张大小**：A4 (210mm × 297mm)
- **页边距**：上37mm，下35mm，左28mm，右26mm
- **页眉页脚**：距离边界1.8cm
- **字体要求**：程序自动检测并退化到可用字体

#### 注意事项

1. **备份文档**：格式化前建议另存一份原始文档
2. **字体兼容**：如缺少特定字体，程序会自动使用替代字体
3. **安全设置**：确保 Word/WPS 宏安全级别允许运行宏
4. **数字签名**：建议为宏添加数字签名提高安全性
5. **WPS用户**：推荐使用 `GongwenFormatter_WPS.bas` 以获得最佳兼容性

> 📘 **WPS详细说明**：查看 `WPS兼容性说明.md` 了解更多WPS使用技巧

#### 故障排除

**常见问题解决：**

Q: 宏不显示在列表中
A: 检查代码是否完整导入，查看是否有编译错误

Q: 运行时报语法错误
A: 确认使用的是最新版本的 `GongwenFormatter.bas` 文件

Q: 格式化结果不符合预期
A: 检查原文档标题格式是否符合识别规则

Q: 字体显示异常
A: 系统缺少对应字体，程序已自动退化处理

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
from gongwen_formatter import GongwenFormatter, FORMAT_MODE_STANDARD, FORMAT_MODE_GOVERNMENT

# 创建格式化器（默认为标准模式）
formatter = GongwenFormatter()

# 或指定政府交付版模式
formatter = GongwenFormatter(format_mode=FORMAT_MODE_GOVERNMENT)

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

## 📊 格式模式对比

| 项目 | GB/T 9704标准模式 | 政府交付版模式 |
|------|------------------|---------------|
| 一级标题 | 居中，无缩进 | 居中，无缩进 |
| 二级标题 | 左对齐，**左缩进2字符** | 左对齐，**首行缩进2字符** |
| 三级标题 | 左对齐，**左缩进2字符** | 左对齐，**首行缩进2字符** |
| 四级标题 | 左对齐，**左缩进2字符** | 左对齐，**首行缩进2字符** |
| 正文 | 首行缩进2字符 | 首行缩进2字符 |
| 文档网格 | 不要求 | 对齐到网格 |

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
**版本**: v1.1
**最后更新**: 2026-02-08
**新增功能**: 政府交付版格式模式支持
