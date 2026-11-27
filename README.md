# PicGenTool-Desktop

一个基于 `Python + PyQt6` 的桌面工具，用于将活动照片批量生成 Word 文档。

## 功能简介

- 输入活动标题、活动地点、活动日期
- 选择或拖拽 1 到 30 张 JPG/JPEG、PNG、GIF、BMP、HEIC、HEIF 图片
- 按文件内容 SHA-256 精确去重，不做视觉相似判断
- 按模板生成活动照片 Word 文档
- 可选生成后自动打开文件
- Agent 可通过无界面 JSON CLI、JSONL 批处理与常驻 stdin/stdout 服务生成文档，使用指南见 `docs/agents-usage.md`
- 记住上次表单偏好；有效拖拽提供高亮与文字反馈
- “关于 / 诊断”提供版本、HEIF、模板、日志和右键菜单状态

## 项目结构

```text
PicGenTool-Desktop/
├── app.py
├── main.py
├── config.py
├── models/
├── services/
├── ui/
├── utils/
├── templates/
├── pgt.spec
└── requirements.txt
```

说明：

- `main.py`：当前主入口
- `app.py`：兼容入口
- `ui/`：界面与后台任务调度
- `services/`：图片处理、模板选择、布局整理、Word 导出
- `models/`：生成任务数据模型
- `utils/`：日志、路径等通用工具
- `templates/`：Word 模板资源

## 开发环境要求

- Python 3.11 及以上
- 建议使用虚拟环境

## 性能基准

图片流水线基准支持固定的 1、10、30 张图片场景。未指定源目录时，脚本会生成
4032×3024 的固定测试图片；结果 JSON 只包含图片数量、每轮耗时和中位数，不包含
图片路径。

```bash
python benchmarks/image_pipeline.py --count 1 --runs 3 --output benchmark-1.json
python benchmarks/image_pipeline.py --count 10 --runs 3 --output benchmark-10.json
python benchmarks/image_pipeline.py --count 30 --runs 3 --output benchmark-30.json
```

优化前后必须在同一台设备上使用同一组样本，各运行至少 3 次并比较中位数。
自动化测试只验证结果结构和输出正确性，不对容易受运行环境影响的绝对耗时设置断言。

## 开发运行说明

### 1. 进入项目目录

```bash
cd /Users/xw/Documents/PicGenTool-Desktop
```

### 2. 创建虚拟环境

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 启动程序

推荐使用新入口：

```bash
python main.py
```

兼容方式也可运行：

```bash
python app.py
```

## Agent 自动化

支持单任务 JSON 请求、请求校验、能力发现、JSONL 批处理和常驻标准流服务。无头模式不启动 Qt、不修改用户偏好、不弹窗或打开 Word，默认不覆盖已有文档。

```bash
python main.py --capabilities --json
python main.py --generate request.json --json
python main.py --batch jobs.jsonl --workers 2 --on-error continue --json
python main.py --serve --json
```

完整请求 Schema、错误码、退出码、超时与安全约束见 `docs/agent-integration.md`；自动化开发指引见 `AGENTS.md`。

## 开发调试建议

### 语法检查

```bash
python -m compileall main.py app.py config.py ui services models utils
```

### 日志位置

Windows 运行日志默认输出到当前用户可写目录：

`%LOCALAPPDATA%\PicGenTool\logs\app.log`

日志最大 5 MiB，保留 3 个轮转备份。主线程和后台线程未捕获异常会记录脱敏堆栈；诊断复制内容不包含标题、地点、照片或输出路径。

其他平台使用对应的用户应用数据目录；如果该目录不可写，会回退到系统临时目录，
不会尝试写入 PyInstaller 安装目录。

### 常见调试点

- 如果程序可以启动但生成失败，先检查 `templates/` 中是否存在对应图片数量的模板文件
- 启动时会检查 1～30 模板文件；可运行 `python tools/check_templates.py` 深度检查 ZIP、XML 和占位符
- 如果图片无法添加，检查文件扩展名是否在 `config.py` 的 `SUPPORTED_FORMATS` 中
- `PGT.exe --diagnose` 无窗口检查冻结环境的 HEIC 解码、30 个模板、1 图 Word 生成、核心导入与 offscreen GUI 启动

## 打包说明

项目当前使用 `PyInstaller` 打包，配置文件为：

`pgt.spec`

### 1. 安装打包依赖

如果本机尚未安装 `PyInstaller`：

```bash
pip install pyinstaller
```

### 2. 执行打包

在项目根目录运行：

```bash
pyinstaller pgt.spec
```

### 3. 打包输出位置

打包完成后，输出位于：

`dist/PGT/`

当前配置使用的是“目录分发”模式，而不是单文件模式。这样做的好处是：

- 启动速度通常更快
- 更适合当前这种依赖较多的桌面工具
- 出问题时更容易定位资源文件

当前 `pgt.spec` 还额外做了以下优化：

- 保留 `one-dir` 打包结构
- 关闭 `UPX` 压缩
- 启用 `noarchive=True`

这套配置更适合当前这种包含 Qt、图片处理和模板资源的桌面程序。

### 4. 打包内容说明

打包时会包含：

- `templates/Template-1.docx` 到 `Template-30.docx`
- `favicon.png`
- `pi-heif` 解码器及其元数据
- 一个固定 HEIC 样本供 `--diagnose` 解码烟测使用

如果你后续新增模板或资源文件，请确认 `pgt.spec` 中的 `datas` 配置同步更新。

### 5. Windows 打包建议

- 请在 Windows 环境下打包 Windows `exe`
- 不建议追求 `one-file`，因为启动速度通常明显更慢
- 推荐直接分发整个 `dist/PGT/` 目录
- 使用 `python tools/package_report.py dist/PGT --max-mib 130 --json package-report.json` 检查体积和最大文件；本轮最终实测 113.10 MiB
- HEIC/HEIF 采用官方解码专用 `pi-heif`，应用只解码并转换 JPEG，不引入 HEIF 编码器体积

## 使用说明

1. 启动程序
2. 输入活动标题
3. 输入活动地点
4. 选择活动日期
5. 添加图片
6. 选择是否自动打开文件
7. 点击“生成文档”

“清空内容”在已有输入或已选照片时会先弹出确认，默认保留当前内容。

生成成功会清除标题和照片，保留地点、日期与复选框，方便连续制作；明确点击“清空内容”会恢复标题、地点、日期默认值。程序只持久化表单偏好、复选框和最近图片目录，不保存照片列表或输出路径。

### 图片右键菜单

运行一次注册脚本，即可在图片文件的右键菜单里看到“一键添加进PGT工具”：

```bash
python tools/register_explorer_menu.py
```

菜单把文件通过 `--add-images` 传给主程序。程序只保留一个窗口：多选或重复打开时，所有照片汇总到已运行的窗口，不会重复启动界面。移动或重新打包程序后，打包版下次启动会只修复已经存在但路径过期或不完整的注册；未安装菜单时不会自动添加。也可重新运行脚本更新命令路径。移除菜单项：

```bash
python tools/register_explorer_menu.py --unregister
```

“窗口置顶”复选框默认开启，可随时关闭。

生成期间底部“生成文档”按钮会切换为“取消任务”。取消采用安全的协作式停止：
当前正在编码的图片或正在保存的 Word 文档会先完成当前原子步骤，随后停止任务并清理临时文件；
取消不会显示“生成失败”弹窗，也不会留下不完整的目标文档。

## 注意事项

- 当前模板机制仍依赖 `templates/Template-1.docx` 到 `templates/Template-30.docx`
- 图片数量必须和对应模板数量匹配
- 输出文件名会根据日期和活动标题自动生成默认名称
- 项目已不再使用 `.ui` 设计器文件，当前界面由代码直接构建

## CI 与发布

GitHub Actions 先在 `windows-latest` 的 Python 3.11/3.12 上执行语法检查、测试和模板深检，再构建 one-dir、运行冻结烟测、检查体积、可选签名并发布 ZIP。

签名仅针对 `dist/PGT/PGT.exe`，不重签第三方 DLL。启用时配置：

- `WINDOWS_CERTIFICATE_BASE64`：PFX 的 Base64 secret
- `WINDOWS_CERTIFICATE_PASSWORD`：PFX 密码 secret
- `WINDOWS_TIMESTAMP_URL`：可选 repository variable，默认 `http://timestamp.digicert.com`

`tools/sign_release.ps1` 使用临时 runner 的当前用户证书存储，签名后验签，并在 `finally` 清理证书和临时 PFX。workflow 会定位 Windows SDK 的 SignTool 并先执行 `-WhatIf`。无证书时 workflow 明确标记未签名，不把未签名产物描述为已签名。同一版本 tag 已存在时发布应失败，必须提升版本号，不删除或重写已发布 tag。

## 后续可优化方向

- 进一步减少模板数量，改为统一模板动态布局
- 增加图片预览和排序能力
- 继续弱化模板驱动，提升布局服务的独立性
