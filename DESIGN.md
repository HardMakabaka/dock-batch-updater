# DOCX 批量更新器 - 设计与开发思路

## 一、整体架构设计

本软件采用分层架构设计，自上而下分为：GUI 层、处理协调层、文档处理层、格式保真层。

```
┌─────────────────────────────────────────────────────────────┐
│                    用户界面层 (GUI)                       │
│  ┌──────────────────────────────────────────────────────┐  │
│  │   MainWindow (PyQt5)                              │  │
│  │   - 文件列表组件 (FileListWidget)                  │  │
│  │   - 替换规则组件 (ReplacementRulesWidget)            │  │
│  │   - 进度显示组件 (ProgressWidget)                   │  │
│  │   - 日志输出组件 (LogWidget)                       │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│                    处理协调层                           │
│  ┌──────────────────────────────────────────────────────┐  │
│  │   BatchProcessor                                  │  │
│  │   - 线程池管理 (ThreadPoolExecutor)               │  │
│  │   - 文件有效性校验                                 │  │
│  │   - 进度回调 & 结果汇总                             │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│                  文档处理层 (Core)                       │
│  ┌──────────────────────────────────────────────────────┐  │
│  │   DocxProcessor                                  │  │
│  │   - 文档加载/备份/保存                              │  │
│  │   - 段落文本替换（跨 run 处理）                     │  │
│  │   - 表格单元格替换（含嵌套表格）                      │  │
│  │   - 统计信息生成                                   │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│                  格式保真层 (Utils)                       │
│  ┌──────────────────────────────────────────────────────┐  │
│  │   FormatPreserver                                │  │
│  │   - 捕获 run 格式 (字体、颜色、加粗等)               │  │
│  │   - 捕获段落格式 (对齐、缩进、行距)                   │  │
│  │   - 捕获单元格格式 (背景色、边框等)                    │  │
│  │   - 应用格式到新文本                                 │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### 分层职责

| 层级 | 模块 | 职责 |
|------|------|------|
| GUI 层 | `gui/main_window.py`, `gui/widgets.py` | 用户交互、界面展示、操作响应 |
| 协调层 | `core/batch_processor.py` | 批量任务调度、并发控制、结果汇总 |
| 处理层 | `core/docx_processor.py` | 单文档处理、文本替换、表格遍历 |
| 保真层 | `utils/format_preserver.py` | 格式捕获与恢复 |

---

## 二、核心模块设计

### 2.1 main.py - 程序入口

**关键设计点：**

#### (1) 路径处理（开发环境 vs 打包环境）

```python
# 问题：打包后 main.py 在根目录，但 gui 模块在 src/gui
# 解决方案：动态设置 sys.path

if getattr(sys, 'frozen', False):
    # PyInstaller 打包后的临时目录
    _src_path = os.path.join(sys._MEIPASS, 'src')
else:
    # 开发环境
    _src_path = os.path.dirname(os.path.abspath(__file__))

if _src_path not in sys.path:
    sys.path.insert(0, _src_path)
```

#### (2) 高 DPI 支持（解决高分屏模糊问题）

```python
# 启用高 DPI 缩放
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
```

---

### 2.2 main_window.py - 主窗口

**关键设计点：**

#### (1) 线程安全的日志输出

```python
# 问题：后台线程直接操作 UI 会导致崩溃
# 解决方案：使用 Qt 的队列调用机制

def log_async(message: str):
    """线程安全的日志输出"""
    QMetaObject.invokeMethod(
        log_widget, "append_log",
        Qt.QueuedConnection,  # 关键：队列调用，不阻塞
        Q_ARG(str, message)
    )

class LogWidget(QTextBrowser):
    @pyqtSlot(str)
    def append_log(self, message: str):
        """线程安全的日志追加（被主线程调用）"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.append(f"[{timestamp}] {message}")
        
        # 自动滚动到最新
        scrollbar = self.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
```

#### (2) 独立处理线程（避免 UI 卡顿）

```python
class ProcessingThread(QThread):
    """后台处理线程"""
    finished = pyqtSignal(list)
    
    def __init__(self, batch_processor, file_paths, replacements, **kwargs):
        super().__init__()
        self.batch_processor = batch_processor
        self.file_paths = file_paths
        self.replacements = replacements
        self.kwargs = kwargs
    
    def run(self):
        """在线程中执行批处理"""
        try:
            results = self.batch_processor.process_documents(
                self.file_paths,
                self.replacements,
                **self.kwargs
            )
        except Exception as e:
            results = []
        finally:
            self.finished.emit(results)
```

#### (3) 处理完成后的结果汇总

```python
def processing_finished(self, results):
    """处理完成后显示汇总"""
    summary = self.batch_processor.get_summary()
    
    # 创建汇总弹窗
    dialog = SummaryDialog(summary)
    dialog.exec_()
```

---

### 2.3 docx_processor.py - 文档处理器

**关键设计点：**

#### (1) 跨 run 文本匹配（避免文本被拆分导致匹配失败）

```python
# 问题：Word 将 "Hello World" 存储为多个 run
# Run1: "Hello " (字体：宋体，12pt，黑色)
# Run2: "World"  (字体：Times，14pt，红色)
# 直接替换 run.text = "Hello World" 会丢失格式

# 解决方案：重建 run 结构，保留格式
def _replace_in_paragraph(self, paragraph, find_text, replace_text):
    # 合并所有 run 的文本进行匹配
    full_text = ''.join(run.text for run in paragraph.runs)
    if find_text not in full_text:
        return False
    
    # 精确定位匹配位置
    start_pos = full_text.find(find_text)
    
    # 重建 runs，保留格式
    new_runs = []
    pos = 0
    replaced = False
    
    for run in paragraph.runs:
        if replaced:
            new_runs.append(run)
            continue
        
        run_text = run.text
        run_start = pos
        run_end = pos + len(run_text)
        
        # 检查是否包含匹配文本
        if start_pos >= run_start and start_pos < run_end:
            # 需要替换
            offset = start_pos - run_start
            
            # 创建新的 runs
            new_runs.append(self._create_run_with_text(run, run_text[:offset]))
            new_runs.append(self._create_run_with_text(run, replace_text))
            new_runs.append(self._create_run_with_text(run, run_text[offset + len(find_text):]))
            
            replaced = True
        else:
            new_runs.append(run)
        
        pos = run_end
    
    # 替换原有 runs
    for run in paragraph.runs[:]:
        paragraph._element.remove(run._element)
    for run in new_runs:
        paragraph._element.append(run._element)
    
    return True
```

#### (2) 表格处理（含嵌套表格）

```python
def _replace_in_table(self, table, find_text, replace_text):
    """递归处理表格（支持嵌套表格）"""
    for row in table.rows:
        for cell in row.cells:
            # 处理单元格段落
            for paragraph in cell.paragraphs:
                self._replace_in_paragraph(paragraph, find_text, replace_text)
            
            # 处理嵌套表格（递归）
            for nested_table in cell.tables:
                self._replace_in_table(nested_table, find_text, replace_text)
```

#### (3) 备份机制

```python
def create_backup(self):
    """创建备份文件"""
    if self.backup_enabled:
        if self.backup_dir:
            backup_path = os.path.join(
                self.backup_dir,
                f"{os.path.basename(self.file_path)}_backup"
            )
        else:
            backup_path = f"{self.file_path}_backup"
        
        shutil.copy2(self.file_path, backup_path)
        return backup_path
    return None
```

---

### 2.4 format_preserver.py - 格式保持器

**关键设计点：**

#### (1) Run 级别格式捕获

```python
@staticmethod
def capture_run_format(run: docx.text.run.Run) -> dict:
    """捕获 Run 的所有格式属性"""
    return {
        'font_name': run.font.name,
        'font_size': run.font.size,
        'bold': run.font.bold,
        'italic': run.font.italic,
        'underline': run.font.underline,
        'color': run.font.color.rgb,
        'highlight_color': run.font.highlight_color.rgb,
        'strike': run.font.strike,
        'double_strike': run.font.double_strike,
        'subscript': run.font.subscript,
        'superscript': run.font.superscript,
    }
```

#### (2) 格式应用

```python
@staticmethod
def apply_run_format(run: docx.text.run.Run, format_data: dict):
    """应用格式到 Run"""
    if 'font_name' in format_data:
        run.font.name = format_data['font_name']
    if 'font_size' in format_data:
        run.font.size = format_data['font_size']
    if 'bold' in format_data:
        run.font.bold = format_data['bold']
    # ... 应用所有格式
```

---

### 2.5 batch_processor.py - 批处理协调器

**关键设计点：**

#### (1) 并发处理（线程池）

```python
def process_documents(self, file_paths, replacements, **kwargs):
    """批量处理文档（并发）"""
    self.results = []
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = []
        
        # 提交所有任务
        for file_path in file_paths:
            future = executor.submit(
                self._process_single_file,
                file_path, replacements, **kwargs
            )
            futures.append(future)
        
        # 收集结果
        for future in as_completed(futures):
            try:
                result = future.result()
                self.results.append(result)
                
                # 进度回调
                if self.progress_callback:
                    self.progress_callback(len(self.results), len(file_paths))
            except Exception as e:
                # 错误处理
                error_result = ProcessingResult(
                    file_path=file_path,
                    success=False,
                    error=str(e)
                )
                self.results.append(error_result)
    
    return self.results
```

#### (2) 文件有效性校验

```python
@staticmethod
def is_docx_file(file_path: str) -> bool:
    """判定表覆盖的文件校验逻辑"""
    
    # 条件1：检查文件扩展名
    if not file_path.lower().endswith('.docx'):
        return False
    
    # 条件2：检查文件是否可读
    if not os.access(file_path, os.R_OK):
        return False
    
    # 条件3：检查文件大小（排除空文件和过大文件）
    size = os.path.getsize(file_path)
    if size == 0 or size > 100 * 1024 * 1024:  # 100 MB
        return False
    
    # 条件4：检查文件内容（ZIP 格式验证）
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 检查必要文件是否存在
            required_files = ['[Content_Types].xml', 'word/document.xml']
            return all(f in zf.namelist() for f in required_files)
    except:
        return False
```

---

## 三、关键难点解决方案

### 难点 1：文本被拆分到多个 run 导致匹配失败

**问题现象：**
```python
# Word 内部存储
paragraph.runs[0].text = "Hello "  # 字体：宋体，12pt
paragraph.runs[1].text = "World"   # 字体：Times，14pt

# 直接替换
paragraph.text = "Hello New World"

# 结果：格式丢失
paragraph.runs[0].text = "Hello New World"  # 默认字体
```

**解决方案：**
```python
# 重建 run 结构，逐字符处理
def _rebuild_runs_with_replacement(self, runs, find_text, replace_text):
    full_text = ''.join(run.text for run in runs)
    start_idx = full_text.find(find_text)
    
    if start_idx == -1:
        return runs
    
    new_runs = []
    current_pos = 0
    replaced = False
    
    for run in runs:
        if replaced:
            new_runs.append(run)
            continue
        
        run_text = run.text
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # 检查是否包含匹配文本
        if start_idx >= run_start and start_idx < run_end:
            # 需要替换
            offset = start_idx - run_start
            
            # 分割 run，保留原格式
            new_runs.append(self._create_run_with_text(run, run_text[:offset]))
            new_runs.append(self._create_run_with_text(run, replace_text))
            new_runs.append(self._create_run_with_text(run, run_text[offset + len(find_text):]))
            
            replaced = True
        else:
            new_runs.append(run)
        
        current_pos = run_end
    
    return new_runs
```

---

### 难点 2：线程安全的日志输出

**问题现象：**
```python
# 后台线程直接操作 UI
def worker():
    log_widget.append("处理中...")  # 崩溃！

thread = threading.Thread(target=worker)
thread.start()
```

**解决方案：**
```python
# 使用 Qt 的信号槽机制或 QMetaObject.invokeMethod
class MainWindow(QWidget):
    def log_async(self, message: str):
        """线程安全的日志输出"""
        QMetaObject.invokeMethod(
            self.log_widget, "append_log",
            Qt.QueuedConnection,  # 关键
            Q_ARG(str, message)
        )

class LogWidget(QTextBrowser):
    @pyqtSlot(str)
    def append_log(self, message: str):
        self.append(message)
```

---

### 难点 3：打包后模块导入失败

**问题现象：**
```
Traceback (most recent call last):
  File "src\main.py", line 11, in <module>
    from gui.main_window import MainWindow
ModuleNotFoundError: No module named 'gui'
```

**问题原因：**
- 开发环境：main.py 在 src/ 目录，`from gui.main_window import MainWindow` 正常
- 打包环境：main.py 在根目录，但 gui 模块在临时目录的 src/gui/

**解决方案：**
```python
# 动态设置 sys.path
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后的临时目录
    sys.path.insert(0, os.path.join(sys._MEIPASS, 'src'))
else:
    # 开发环境
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
```

---

## 四、测试策略

### 4.1 后端测试（unittest + 判定表/因果图）

#### test_docx_validation_decision.py - 文件校验测试

**判定表覆盖所有条件组合：**

| 扩展名.docx | 文件可读 | 文件大小有效 | ZIP结构有效 | 期望结果 |
|-------------|---------|-----------|----------|---------|
| 是 | 是 | 是 | 是 | True |
| 否 | 是 | 是 | 是 | False |
| 是 | 否 | 是 | 是 | False |
| 是 | 是 | 否 | 是 | False |
| 是 | 是 | 是 | 否 | False |
| 否 | 否 | 是 | 是 | False |
| 是 | 否 | 否 | 是 | False |
| 否 | 是 | 否 | 是 | False |

#### test_core.py - 核心功能测试

```python
class TestDocxProcessor(unittest.TestCase):
    def test_load_document(self):
        """测试文档加载"""
        processor = DocxProcessor("test.docx")
        processor.load()
        self.assertIsNotNone(processor.doc)
    
    def test_replace_text(self):
        """测试文本替换"""
        processor = DocxProcessor("test.docx")
        processor.load()
        results = processor.replace_text("旧文本", "新文本")
        self.assertEqual(results['total'], 1)
        self.assertEqual(results['success'], 1)
    
    def test_create_backup(self):
        """测试备份功能"""
        processor = DocxProcessor("test.docx")
        processor.backup_enabled = True
        backup_path = processor.create_backup()
        self.assertTrue(os.path.exists(backup_path))

class TestBatchProcessor(unittest.TestCase):
    def test_process_documents(self):
        """测试批量处理"""
        batch = BatchProcessor()
        results = batch.process_documents(["test1.docx", "test2.docx"], {})
        self.assertEqual(len(results), 2)
```

#### test_docx_processor_additional.py - 边界测试

```python
class TestDocxProcessorAdditional(unittest.TestCase):
    def test_cross_run_replacement(self):
        """测试跨 run 替换"""
        # 创建测试文档，文本被拆分到多个 run
        # 验证替换后格式保留
    
    def test_nested_tables(self):
        """测试嵌套表格处理"""
        # 创建包含嵌套表格的文档
        # 验证所有层级都被处理
    
    def test_exception_handling(self):
        """测试异常处理"""
        # 传入无效文件路径
        # 验证错误被正确捕获
```

### 4.2 E2E 测试（可选）

使用 Playwright 进行 GUI 自动化测试：

```python
def test_e2e_batch_processing():
    """端到端测试：批量处理流程"""
    
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        
        # 启动应用
        page.goto("file:///path/to/dist/DOCX Batch Updater.exe")
        
        # 添加文件
        page.click("text=Add Files")
        # ... 选择测试文件
        
        # 添加规则
        page.fill("input[placeholder='Find']", "2024")
        page.fill("input[placeholder='Replace']", "2025")
        page.click("text=Add Rule")
        
        # 启动处理
        page.click("text=Start Processing")
        
        # 验证结果
        page.wait_for_selector("text=Processing completed")
        assert page.inner_text(".statistics") == "Total: 1, Success: 1, Failed: 0"
```

---

## 五、打包配置

### 5.1 PyInstaller 参数说明

| 参数 | 说明 | 必要性 |
|------|------|--------|
| `--name="DOCX Batch Updater"` | exe 文件名称 | 必须 |
| `--windowed` | 不显示控制台窗口（GUI 应用） | 必须 |
| `--onefile` | 打包为单个 exe 文件 | 必须 |
| `--add-data="src;src"` | 包含 src 目录（GUI 和核心代码） | 必须 |
| `--hidden-import=PyQt5.sip` | PyQt5 隐式依赖 | 必须 |
| `--hidden-import=docx` | python-docx 隐式依赖 | 必须 |
| `--hidden-import=docx.opc.constants` | python-docx 子模块 | 必须 |
| `--hidden-import=docx.oxml` | python-docx XML 处理 | 必须 |
| `--hidden-import=docx.oxml.text.paragraph` | 段落处理 | 必须 |
| `--hidden-import=docx.oxml.table` | 表格处理 | 必须 |

### 5.2 完整打包命令

```bash
pyinstaller \
  --name="DOCX Batch Updater" \
  --windowed \
  --onefile \
  --add-data="src;src" \
  --hidden-import=PyQt5.sip \
  --hidden-import=docx \
  --hidden-import=docx.opc.constants \
  --hidden-import=docx.oxml \
  --hidden-import=docx.oxml.text.paragraph \
  --hidden-import=docx.oxml.table \
  src/main.py
```

### 5.3 注意事项

1. **Python 版本兼容性**
   - Python 3.13 可能与 PyInstaller 5.13.0 不兼容
   - 推荐使用 Python 3.8-3.11

2. **警告处理**
   - 打包过程中出现的 "Hidden import 'sip' not found!" 警告可忽略
   - 不影响程序运行

3. **打包时间**
   - 约需 30-60 秒，取决于计算机性能

4. **输出文件**
   - 位置：`dist/DOCX Batch Updater.exe`
   - 大小：约 39 MB
   - 可直接在 Windows 系统上运行，无需安装 Python

---

## 六、项目目录结构

```
dock-batch-updater/
├── src/                           # 源代码目录
│   ├── main.py                    # 程序入口（路径处理、高 DPI）
│   ├── gui/                      # GUI 层
│   │   ├── __init__.py
│   │   ├── main_window.py         # 主窗口（线程安全日志）
│   │   └── widgets.py            # 自定义组件
│   ├── core/                     # 处理层
│   │   ├── __init__.py
│   │   ├── docx_processor.py      # 文档处理器（跨 run、嵌套表格）
│   │   └── batch_processor.py    # 批处理协调器（线程池、校验）
│   └── utils/                    # 工具层
│       ├── __init__.py
│       └── format_preserver.py   # 格式保持器
├── tests/                         # 测试目录
│   ├── test_core.py              # 核心功能测试
│   ├── test_docx_processor_additional.py  # 边界测试
│   └── test_docx_validation_decision.py  # 判定表测试
├── requirements.txt                # Python 依赖
├── build.bat                     # 打包脚本
├── README.md                    # 用户文档
├── INSTRUCTIONS.md               # 需求文档
├── DESIGN.md                    # 设计文档（本文档）
├── AGENTS.md                    # 项目知识库
└── ToDo.md                     # 任务清单
```

---

## 七、技术栈总结

| 技术点 | 选型 | 理由 |
|--------|------|------|
| GUI 框架 | PyQt5 | 跨平台、功能丰富、原生外观 |
| 文档处理 | python-docx | 专为 DOCX 设计、格式保真度高 |
| 并发处理 | ThreadPoolExecutor | 标准库、简单高效 |
| 打包工具 | PyInstaller | 单文件打包、依赖完整 |
| 测试框架 | unittest | Python 内置、断言丰富 |

---

## 八、设计原则

1. **格式保真优先**
   - 使用 run 级别操作，避免破坏段落结构
   - 捕获并恢复所有格式属性

2. **用户体验优先**
   - 实时进度反馈
   - 详细的日志输出
   - 友好的错误提示

3. **性能优化**
   - 多线程并发处理
   - 避免阻塞 UI 线程

4. **代码质量**
   - 遵循 PEP 8 规范
   - 完整的类型提示
   - 详细的文档字符串

5. **测试覆盖**
   - 判定表覆盖所有条件组合
   - 边界测试确保健壮性
   - 单元测试保证功能正确性

---

**文档版本**：1.0  
**最后更新**：2026-02-10  
**状态**：已完成开发、测试、打包
