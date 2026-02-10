# Git CI 自动化打包配置

## 概述

本项目已配置 GitHub Actions 实现自动化打包，当代码推送到 `master` 或 `main` 分支时，自动触发构建流程。

## 工作流文件

位置：`.github/workflows/build.yml`

## 触发条件

| 触发方式 | 说明 |
|---------|------|
| **Push** | 推送到 `master` 或 `main` 分支 |
| **Pull Request** | 针对 `master` 或 `main` 分支的 PR |
| **Manual** | 在 GitHub Actions 页面手动触发 |

## 自动化流程

```
┌─────────────────────────────────────────────────────────┐
│ 1. 检出代码 (Checkout)                           │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 2. 设置 Python 3.8 环境                           │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 3. 缓存 pip 依赖（加速后续构建）                   │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 4. 安装依赖 (pip install -r requirements.txt)       │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 5. 运行测试 (unittest discover)                   │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 6. 清理旧构建文件                                  │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 7. 执行打包 (PyInstaller)                          │
│    - 打包为单文件 exe                                │
│    - 包含所有依赖                                    │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 8. 上传构建产物 (Artifact)                          │
│    - 保留 30 天                                      │
│    - 可在 Actions 页面下载                              │
└─────────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────────┐
│ 9. (可选) 创建 Release (仅在 Tag 推送时)              │
└─────────────────────────────────────────────────────────┘
```

## 下载构建产物

### 方式 1：从 Actions 页面下载

1. 访问：`https://github.com/你的用户名/dock-batch-updater/actions`
2. 选择一个成功的构建（绿色 ✓）
3. 滚动到底部，找到 "Artifacts" 部分
4. 下载 `DOCX-Batch-Updater-Windows`

### 方式 2：从 Release 下载

1. 推送 tag（如 `v1.0.0`）
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
2. 访问：`https://github.com/你的用户名/dock-batch-updater/releases`
3. 下载对应的 exe 文件

## 手动触发构建

在 GitHub 网页上手动触发打包：

1. 访问：`https://github.com/你的用户名/dock-batch-updater/actions`
2. 选择 "Build Windows Executable" 工作流
3. 点击 "Run workflow" 按钮
4. 选择分支并确认

## 配置参数说明

| 参数 | 值 | 说明 |
|------|-----|------|
| `runs-on` | `windows-latest` | 在 Windows 虚拟机上运行 |
| `python-version` | `3.8` | Python 版本（PyInstaller 兼容性好） |
| `retention-days` | `30` | Artifact 保留 30 天 |

## PyInstaller 参数

```yaml
pyinstaller --name="DOCX Batch Updater" --windowed --onefile `
  --add-data="src;src" `
  --hidden-import=PyQt5.sip `
  --hidden-import=docx `
  --hidden-import=docx.opc.constants `
  --hidden-import=docx.oxml `
  --hidden-import=docx.oxml.text.paragraph `
  --hidden-import=docx.oxml.table `
  src/main.py
```

## 性能优化

### 1. 依赖缓存

```yaml
- name: Cache pip packages
  uses: actions/cache@v4
  with:
    path: ~\AppData\Local\pip\Cache
    key: ${{ runner.os }}-pip-${{ hashFiles('**/requirements.txt') }}
    restore-keys: |
      ${{ runner.os }}-pip-
```

- 首次构建：约 3-5 分钟
- 后续构建（命中缓存）：约 1-2 分钟

### 2. 并行测试

```yaml
- name: Run tests
  run: |
    python -m unittest discover -s tests -p "test*.py" -v
```

测试失败时，构建会立即停止，不会执行打包步骤。

## 失败排查

### 构建失败常见原因

| 错误 | 原因 | 解决方案 |
|------|------|---------|
| `ModuleNotFoundError` | 依赖未安装 | 检查 `requirements.txt` |
| 测试失败 | 代码问题 | 查看测试日志，修复代码 |
| 打包失败 | PyInstaller 配置错误 | 检查隐藏导入参数 |
| Artifact 上传失败 | 文件未生成 | 检查打包步骤日志 |

### 查看构建日志

1. 访问：Actions 页面
2. 点击失败的构建
3. 展开 "Build EXE" 步骤
4. 查看详细错误信息

## 本地测试 CI 流程

在推送前，可以在本地模拟 CI 流程：

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 运行测试
python -m unittest discover -s tests -p "test*.py" -v

# 3. 清理旧构建
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

# 4. 执行打包
pyinstaller --name="DOCX Batch Updater" --windowed --onefile `
  --add-data="src;src" `
  --hidden-import=PyQt5.sip `
  --hidden-import=docx `
  --hidden-import=docx.opc.constants `
  --hidden-import=docx.oxml `
  --hidden-import=docx.oxml.text.paragraph `
  --hidden-import=docx.oxml.table `
  src/main.py

# 5. 检查产物
ls dist
```

## 高级配置

### 发布流程自动化

推送 tag 时自动创建 Release：

```bash
# 创建 tag
git tag -a v1.0.0 -m "版本 1.0.0"
git push origin v1.0.0

# CI 自动创建 Release 并上传 exe
```

### 多平台构建

如需支持多平台，可以修改 `.github/workflows/build.yml`：

```yaml
jobs:
  build:
    strategy:
      matrix:
        os: [windows-latest, ubuntu-latest, macos-latest]
    runs-on: ${{ matrix.os }}
    steps:
      # ... 构建步骤
```

## 费用说明

GitHub Actions 公开仓库免费：
- 每月 2000 分钟免费额度
- 本项目单次构建约 5 分钟
- 可执行约 400 次构建/月

## 维护建议

1. **定期更新依赖**
   - 每月检查 `requirements.txt` 中的依赖版本
   - 更新后测试构建是否成功

2. **监控构建状态**
   - 在 GitHub 设置中启用通知
   - 构建失败时及时处理

3. **清理旧 Artifact**
   - 30 天自动过期
   - 手动清理不必要的历史产物

---

**配置文件**：`.github/workflows/build.yml`  
**状态**：已启用，自动运行中
