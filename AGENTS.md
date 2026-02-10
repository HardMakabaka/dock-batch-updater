# PROJECT KNOWLEDGE BASE

**Generated:** 2026-02-10 (Asia/Shanghai)
**Commit:** 0764bb7
**Branch:** master

## OVERVIEW
Windows 桌面应用（PyQt5）用于批量更新 DOCX（python-docx），重点是严格保留格式；可用 PyInstaller 打包成单文件 exe。

## STRUCTURE
```
./
├── src/            # 应用源码（以 src 作为运行工作目录）
├── tests/          # unittest 测试
├── build.bat       # PyInstaller 打包脚本
├── requirements.txt
└── README.md
```

## 文件索引（分类）

### 文档
- `README.md`：运行/构建/使用说明
- `INSTRUCTIONS.md`：需求与验收标准（开发指引）
- `ToDo.md`：本地任务清单（可能含新增修复记录）

### 入口与启动
- `src/main.py`：PyQt5 应用入口（创建 QApplication + MainWindow）

### GUI（PyQt5）
- `src/gui/main_window.py`：主窗口；Start/Stop 流程、线程调度、进度与日志联动（含线程安全日志 `log_async`）
- `src/gui/widgets.py`：文件列表/规则输入/进度条/日志组件等自定义控件
- `src/gui/__init__.py`：GUI 包导出

### Core（批处理与 DOCX 处理）
- `src/core/batch_processor.py`：批量校验与并发处理（ThreadPoolExecutor）、结果聚合、stop 取消、统计摘要
- `src/core/docx_processor.py`：单文档处理；加载/备份/替换/保存/统计；段落与表格替换（支持跨 run 文本替换）
- `src/core/__init__.py`：Core 包导出

### Utils（格式保真）
- `src/utils/format_preserver.py`：捕获/恢复 run、段落、单元格格式；文本定位辅助
- `src/utils/__init__.py`：Utils 包导出

### 测试（unittest）
- `tests/test_core.py`：核心回归测试（DocxProcessor/BatchProcessor/FormatPreserver 等）
- `tests/test_docx_processor_additional.py`：补充 DocxProcessor 边界/异常/跨 run 场景测试
- `tests/test_docx_validation_decision.py`：基于判定表/因果图的 is_docx_file 测试
- `tests/__init__.py`

### 构建与依赖
- `requirements.txt`：python-docx / PyQt5 / PyInstaller 版本
- `build.bat`：打包脚本（PyInstaller onefile + windowed）

### 运行产物（不要提交）
- `__pycache__/`、`*.pyc`：Python 字节码缓存
- `.coverage`：coverage 运行产物
- `build/`、`dist/`、`*.spec`：PyInstaller 构建产物

## 模块说明（分类）

### GUI 层
- `MainWindow`（`src/gui/main_window.py`）：负责采集文件与规则、启动后台处理线程、在主线程更新进度/日志/统计、弹窗汇总结果。
- `ProcessingThread`（`src/gui/main_window.py` 内部类）：在后台线程调用 `BatchProcessor.process_documents`；完成后通过 Qt 信号触发 `processing_finished`。
- `log_async`（`src/gui/main_window.py`）：线程安全日志输出（通过 `QMetaObject.invokeMethod` queued 调用回主线程）。
- Widgets（`src/gui/widgets.py`）：
  - `FileListWidget`：文件选择/目录扫描/列表维护
  - `ReplacementRulesWidget`：替换规则维护（Find/Replace）
  - `ProgressWidget`：进度与统计显示
  - `LogWidget`：日志展示（HTML 追加）

### 处理层
- `BatchProcessor`（`src/core/batch_processor.py`）：
  - 入口：`process_documents(file_paths, replacements, ...)`
  - 负责：文件有效性校验、线程池并发处理、进度回调、结果回调、汇总统计
- `DocxProcessor`（`src/core/docx_processor.py`）：
  - 入口：`load()` / `replace_text()` / `save()` / `create_backup()`
  - 替换范围：段落 + 表格（含嵌套表格）
  - 关键点：`_replace_in_paragraph` 支持跨多个 run 的文本匹配与替换（避免“文本被拆分到多个 run 导致匹配失败”）

### 格式保真
- `FormatPreserver`（`src/utils/format_preserver.py`）：捕获并应用 run/段落/单元格格式数据，供替换时重放格式。

### 测试
- `tests/test_core.py`：主回归集，覆盖批处理、替换、统计、结果对象等核心路径。
- `tests/test_docx_processor_additional.py`：补充边界/异常/跨 run/嵌套表格等场景。
- `tests/test_docx_validation_decision.py`：基于判定表/因果图覆盖 `DocxProcessor.is_docx_file` 的关键条件组合。

## WHERE TO LOOK
| 需求/问题 | 位置 | 备注 |
|---|---|---|
| 点击 Start Processing 的执行链路 | `src/gui/main_window.py` | `start_processing` 创建 QThread 并调用 `BatchProcessor.process_documents` |
| “替换不生效/跨 run” | `src/core/docx_processor.py` | 优先看 `_replace_in_paragraph` 与表格遍历 |
| “处理完成后 UI 异常/窗口关闭” | `src/gui/main_window.py` | 后台线程不得直接更新 UI；使用 `log_async`/`invokeMethod` |
| 批量并发与统计 | `src/core/batch_processor.py` | `ThreadPoolExecutor` + `get_summary` |
| 格式丢失 | `src/utils/format_preserver.py` | run/paragraph/cell 的 capture/apply |
| 打包 exe | `build.bat` | PyInstaller 参数、hidden-import、add-data |
| 运行测试 | `tests/*.py` | unittest 为主 |

## COMMANDS
```bash
# 安装依赖
pip install -r requirements.txt

# 从源码运行（推荐按 README：在 src 目录运行）
cd src
python main.py

# 运行全部测试
python -m unittest discover -s tests -p "test*.py" -v

# 打包（Windows）
build.bat
```

## ANTI-PATTERNS (THIS PROJECT)
- 不要从后台线程直接操作 Qt UI（日志/进度/统计必须通过 queued 调用回主线程）。
- 不要把 `__pycache__/`、`*.pyc`、`.coverage` 之类的运行产物提交到仓库（当前工作区可能会出现这些文件）。
- 不要在替换实现里只依赖单个 run 的 `run.text` 进行匹配（跨 run 会导致“替换无效”）。

## NOTES
- 运行目录约定：按 README 建议从 `src/` 运行（`cd src && python main.py`），避免导入路径不一致。
- 依赖版本兼容性：`PyInstaller==5.13.0` 在 Python 3.13 环境可能无法安装/使用（本地安装日志已体现）；若要打包，请使用兼容版本的 Python 或调整 PyInstaller 版本。
