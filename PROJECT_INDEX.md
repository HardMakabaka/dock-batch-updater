# PROJECT_INDEX.md - 项目索引文件（OpenCode 维护）

## 迁移文件

| 文件路径 | 作用说明 |
|---------|---------|
| `./tests/test_core.py` | 核心功能单元测试 |
| `./tests/test_docx_processor_additional.py` | DOCX 处理器附加测试 |
| `./tests/test_docx_validation_decision.py` | DOCX 验证决策测试 |

## 入口与编排文件

| 文件路径 | 职责说明 |
|---------|---------|
| `./src/main.py` | 应用程序主入口，启动 GUI 界面 |
| `./build.bat` | Windows 平台构建脚本，执行 PyInstaller 打包 |

## 核心自建模块

| 文件路径 | 职责说明 |
|---------|---------|
| `./src/core/batch_processor.py` | 批量处理核心逻辑，协调整个处理流程 |
| `./src/core/docx_processor.py` | DOCX 文件处理专用模块，负责文档解析和修改 |
| `./src/gui/main_window.py` | 主窗口界面组件，包含所有 UI 逻辑 |
| `./src/gui/widgets.py` | 自定义 GUI 组件模块，提供可重用界面元素 |
| `./src/utils/format_preserver.py` | 格式保持工具，确保修改后文档格式不丢失 |

## CI/CD 与部署相关文件

| 文件路径 | 职责说明 |
|---------|---------|
| `./.github/workflows/build.yml` | GitHub Actions 构建工作流，自动化测试和打包 |
| `./openclaw.yml` | OpenClaw 项目配置文件，定义开发流程和检查规则 |
| `./requirements.txt` | Python 依赖清单，定义项目运行所需包 |

## 配置与文档文件

| 文件路径 | 职责说明 |
|---------|---------|
| `./AGENTS.md` | Agent 开发指挥流程规范 |
| `./CI.md` | CI 自动化打包配置说明文档 |
| `./DESIGN.md` | 项目设计文档，包含架构和功能说明 |
| `./INSTRUCTIONS.md` | 使用说明文档 |
| `./README.md` | 项目介绍和快速开始指南 |
| `./ToDo.md` | 待办事项列表 |

---

**维护者**: OpenCode (执行者)
**更新频率**: 项目结构变更时更新
**禁止**: 禁止包含绝对路径、环境细节、敏感信息
**建议**: 忽略 `.gitignore` 命中的路径与文件
