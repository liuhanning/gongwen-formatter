// ============================================================================
// Gongwen Document Formatter - Web Interface
// ============================================================================

class GongwenFormatterUI {
    constructor() {
        this.selectedFile = null;
        this.initializeElements();
        this.attachEventListeners();
    }

    initializeElements() {
        // Upload elements
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.fileInfo = document.getElementById('fileInfo');
        this.fileName = document.getElementById('fileName');
        this.fileSize = document.getElementById('fileSize');
        this.removeFileBtn = document.getElementById('removeFile');

        // Action buttons
        this.formatBtn = document.getElementById('formatBtn');
        this.createDemoBtn = document.getElementById('createDemoBtn');

        // Status panel
        this.statusPanel = document.getElementById('statusPanel');
        this.statusIcon = document.getElementById('statusIcon');
        this.statusTitle = document.getElementById('statusTitle');
        this.statusMessage = document.getElementById('statusMessage');
        this.statusProgress = document.getElementById('statusProgress');
        this.progressFill = document.getElementById('progressFill');
    }

    attachEventListeners() {
        // Upload area events
        this.uploadArea.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        
        // Drag and drop
        this.uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.uploadArea.addEventListener('drop', (e) => this.handleDrop(e));

        // Remove file
        this.removeFileBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.clearFile();
        });

        // Action buttons
        this.formatBtn.addEventListener('click', () => this.formatDocument());
        this.createDemoBtn.addEventListener('click', () => this.createDemo());
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.remove('drag-over');
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    processFile(file) {
        // Validate file type
        if (!file.name.endsWith('.docx')) {
            this.showStatus('error', '文件格式错误', '请上传 .docx 格式的 Word 文档');
            return;
        }

        this.selectedFile = file;
        
        // Display file info
        this.fileName.textContent = file.name;
        this.fileSize.textContent = this.formatFileSize(file.size);
        
        // Show file info, hide upload area
        this.uploadArea.style.display = 'none';
        this.fileInfo.style.display = 'flex';
        
        // Enable format button
        this.formatBtn.disabled = false;

        // Hide status panel
        this.statusPanel.style.display = 'none';
    }

    clearFile() {
        this.selectedFile = null;
        this.fileInput.value = '';
        
        // Show upload area, hide file info
        this.uploadArea.style.display = 'block';
        this.fileInfo.style.display = 'none';
        
        // Disable format button
        this.formatBtn.disabled = true;

        // Hide status panel
        this.statusPanel.style.display = 'none';
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
    }

    async formatDocument() {
        if (!this.selectedFile) {
            this.showStatus('error', '未选择文件', '请先上传要格式化的文档');
            return;
        }

        // 获取选中的格式模式
        const formatMode = document.querySelector('input[name="formatMode"]:checked').value;
        const modeName = formatMode === 'government' ? '政府交付版' : 'GB/T 9704标准';

        this.showStatus('processing', '正在处理...', `正在使用${modeName}格式化您的文档，请稍候`);
        this.showProgress(0);

        try {
            // Simulate processing with progress
            await this.simulateProgress();

            // Call Python backend
            const formData = new FormData();
            formData.append('file', this.selectedFile);
            formData.append('format_mode', formatMode);

            const response = await fetch('/api/format', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error('格式化失败');
            }

            // Download formatted file
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = this.selectedFile.name.replace('.docx', '_formatted.docx');
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            this.showProgress(100);
            this.showStatus('success', '格式化完成！', `文档已按照${modeName}格式化并下载`);

        } catch (error) {
            console.error('Error:', error);
            this.showStatus('error', '处理失败', this.getErrorMessage(error));
        }
    }

    async createDemo() {
        // 获取选中的格式模式
        const formatMode = document.querySelector('input[name="formatMode"]:checked').value;
        const modeName = formatMode === 'government' ? '政府交付版' : 'GB/T 9704标准';

        this.showStatus('processing', '正在创建示例文档...', `正在生成符合${modeName}的示例文档`);
        this.showProgress(0);

        try {
            // Simulate progress
            await this.simulateProgress();

            const response = await fetch('/api/create-demo', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ format_mode: formatMode })
            });

            if (!response.ok) {
                throw new Error('创建示例文档失败');
            }

            // Download demo file
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'demo_gongwen.docx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            this.showProgress(100);
            this.showStatus('success', '示例文档已创建！', `示例文档已按${modeName}生成并下载`);

        } catch (error) {
            console.error('Error:', error);
            this.showStatus('error', '创建失败', this.getErrorMessage(error));
        }
    }

    async simulateProgress() {
        const steps = [0, 20, 40, 60, 80, 95];
        for (const progress of steps) {
            this.showProgress(progress);
            await this.delay(200);
        }
    }

    showProgress(percent) {
        if (this.statusProgress) {
            this.statusProgress.style.display = 'block';
            this.progressFill.style.width = percent + '%';
        }
    }

    showStatus(type, title, message) {
        this.statusPanel.style.display = 'block';
        this.statusTitle.textContent = title;
        this.statusMessage.textContent = message;

        // Remove all status classes
        this.statusIcon.classList.remove('success', 'error', 'processing');
        
        // Add appropriate class
        this.statusIcon.classList.add(type);

        // Hide progress for success/error
        if (type === 'success' || type === 'error') {
            this.statusProgress.style.display = 'none';
        }

        // Scroll to status
        this.statusPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    getErrorMessage(error) {
        // Check if server is running
        if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            return '无法连接到服务器。请确保 Python 后端正在运行。\n\n运行命令: python gongwen_formatter.py';
        }
        return error.message || '发生未知错误，请重试';
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// ============================================================================
// Initialize on page load
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
    const app = new GongwenFormatterUI();
    
    // Add welcome animation
    setTimeout(() => {
        const panels = document.querySelectorAll('.panel');
        panels.forEach((panel, index) => {
            panel.style.animationDelay = `${0.2 + index * 0.1}s`;
        });
    }, 100);

    // Show instructions on first load
    if (!localStorage.getItem('gongwen_visited')) {
        showWelcomeMessage();
        localStorage.setItem('gongwen_visited', 'true');
    }
});

function showWelcomeMessage() {
    const statusPanel = document.getElementById('statusPanel');
    const statusIcon = document.getElementById('statusIcon');
    const statusTitle = document.getElementById('statusTitle');
    const statusMessage = document.getElementById('statusMessage');

    statusPanel.style.display = 'block';
    statusIcon.classList.add('processing');
    statusTitle.textContent = '欢迎使用公文格式化工具';
    statusMessage.textContent = '上传 Word 文档进行格式化，或创建示例文档查看效果。所有格式均符合 GB/T 9704 党政机关公文格式规范。';

    // Auto-hide after 5 seconds
    setTimeout(() => {
        statusPanel.style.display = 'none';
    }, 5000);
}

// ============================================================================
// Keyboard shortcuts
// ============================================================================

document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + O: Open file
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        document.getElementById('fileInput').click();
    }
    
    // Ctrl/Cmd + Enter: Format document
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        const formatBtn = document.getElementById('formatBtn');
        if (!formatBtn.disabled) {
            formatBtn.click();
        }
    }
});
