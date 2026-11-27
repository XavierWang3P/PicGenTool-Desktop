# 发现与决策

## 需求
- HEIC/HEIF 支持
- 重复图片去重
- 日志轮转
- 右键菜单路径自愈
- CI 跑测试
- 全局异常兜底
- 模板完整性自检
- 打包体积优化
- Word 与图片阶段重叠
- 记住上次输入
- 拖拽高亮反馈
- 关于/诊断入口
- 代码签名

## 研究发现
- 项目是 Python 3.11+ / PyQt6 桌面应用，使用 Pillow、python-docx、docxtpl。
- 已有 pytest 与 pytest-qt 测试体系，测试覆盖图片服务、Word 导出、主窗口、工作线程、日志、IPC 等。
- 已有 1/10/30 张图片性能基准与启动基准，适合验证流水线重叠和打包优化。
- 打包采用 PyInstaller one-dir、关闭 UPX、`noarchive=True`；现有 `dist/` 与 `build/` 可用于测量基线，但属于生成物。
- 右键菜单当前通过 `tools/register_explorer_menu.py` 手动注册；README 明确移动后需要重新注册，尚无自愈。
- 模板依赖 `Template-1.docx` 至 `Template-30.docx`，完整性自检可在启动/诊断和构建测试复用。
- 当前目录不是 Git 仓库，无法取得 diff 或提交历史；编辑时需避免无关改动。
- `main.py` 只在最外层 `try/except` 记录后重新抛出异常，未安装 `sys.excepthook`、`threading.excepthook` 或 Qt 友好的崩溃提示；全局兜底尚不完整。
- `GenerationWorker` 当前严格串行执行“全部图片准备 → 全部 Word 导出”，图片处理内部已有最多 4 线程并发。
- Word 模板与 `DocxTemplate` 当前只在图片全部完成后打开；安全重叠可让模板打开/上下文准备与图片编码并行，但 `InlineImage` 构造、render、save 仍应留在同一工作线程。
- 图片筛选只看扩展名，不检查存在性或文件内容；`SUPPORTED_FORMATS` 目前不含 `.heic/.heif`。
- `_add_images` 直接追加路径，没有去重；单实例 IPC 与拖拽都会经过该入口，适合集中去重。
- 日志使用无限增长的 `logging.FileHandler`，可替换为 `RotatingFileHandler` 并保留现有回退目录语义。
- 模板 1–30 实际齐全，另有 `Template-Original.docx`；当前 `TemplateService` 仅按需检查一个文件，没有批量完整性报告。
- 主窗口由单个 `QWidget` 组成，没有菜单栏；“关于/诊断”更适合增加紧凑的次要按钮或链接，避免为单入口改成 `QMainWindow`。
- 表单当前不会持久化任何输入；可用 `QSettings` 保存标题、地点、复选框与最近图片目录，但不应自动恢复旧照片列表，以避免隐私、失效路径和误生成。
- 拖拽事件只接受合法格式，未实现 `dragLeaveEvent` 或视觉属性切换。
- 现有测试覆盖串行取消、原子保存、路径隐私和 1/10/30 张真实模板集成，适合扩充而无需重写测试体系。
- 模板目录包含完整的 `Template-1.docx` 至 `Template-30.docx`，当前可建立启动时轻量校验和诊断时详细校验。
- GitHub Actions 当前只有“推送 main 即构建并发布”一个 job；它不安装开发依赖、不运行 pytest，并会删除/重建同版本 tag。应把测试作为独立必经 job，并让发布 job 依赖测试。
- 本地基线：42 项测试全部通过（3.57 秒）。
- 现有 `dist/PGT` 为 Python 3.11 构建，901 个文件、119.01 MiB；最大项包括 Qt 软件 OpenGL 20.64 MB、Qt Core/Gui/Widgets、Pillow AVIF 7.90 MB、Qt PDF 4.62 MB。体积优化重点应是未使用 Qt 模块/插件、Pillow 可选格式插件和 PyInstaller 收集策略，而不是切成 one-file。
- 当前虚拟环境 PyInstaller 版本为 6.22.3；CI 未锁定 PyInstaller 版本，构建结果可能漂移。
- pillow-heif 官方文档支持在 Pillow 使用前调用 `register_heif_opener()` 手动注册 HEIF opener；pi-heif 使用同项目同名 API。这是集中初始化 HEIC/HEIF 解码、且易于测试的最小接入方式。
- 抽样模板 1/2/9/10/30 的占位符结构一致：均含 `title`、`location`、`dateForDoc`，并恰好包含 `image1` 到 `imageN`；完整性自检可据此定义确定规则。
- UI/UX 指南要求拖拽不是唯一入口（现有“选择照片”按钮已满足），拖入时应同时提供边框/背景与文字状态，不能只靠颜色；异步计数/状态变化应给出有上下文的完整消息。
- PyPI 于 2026-11-24 发布 pillow-heif 1.9.0，但它要求 Pillow>=13，与项目固定的 Pillow 12.2 冲突。初期使用 1.8.1 验证解码可行；其 Windows wheel 的 libheif 直接静态导入 libx265，完整包使产物达 139.23 MiB。官方解码专用 pi-heif 1.4.0 支持 Python 3.11/3.12、Pillow>=11.1，提供 Windows wheels；最终产物 113.09 MiB。
- GitHub 官方文档确认 secret 不能直接用于 step 的 `if:`；应先映射为 job/step 环境变量，未设置时值为空字符串。官方也建议尽量避免把 secret 放在进程命令行，并要求 Base64 只作为编码而非加密。
- Microsoft SignTool 文档要求现代 SDK 显式指定 `/fd` 和 RFC3161 `/td`，推荐 SHA256；应用签名验收应使用 `verify /pa`，并用 `/tw` 对缺少时间戳发出警告。

## 技术决策
| 决策 | 理由 |
|------|------|
| 以内容哈希定义重复图片 | 精确、确定、不会误删视觉相似但有意义的照片 |
| 代码签名做成凭据外置的可选发布步骤 | 证书属于敏感外部资产，项目应只定义签名接口和验证步骤 |
| 保持 one-dir 并建立体积预算 | 与现有启动性能目标一致，便于定位体积来源 |
| 流水线重叠限定为“模板预加载与图片编码并行” | 避免跨线程共享 docxtpl/python-docx 对象，同时获得确定性的重叠收益 |
| 只记住表单偏好和最近目录，不恢复照片列表 | 满足便利性，同时避免保留个人照片路径和失效输入 |
| 所有入口统一在 `ImageService`/窗口添加逻辑去重 | 文件选择、拖拽、IPC、右键菜单行为一致 |
| CI 分为 test 和 package/release 两层 | 每次变更先验证，发布只在测试通过后运行，职责清晰 |
| 使用 `pi-heif` 并显式调用 `register_heif_opener()` | 官方支持的 Pillow 插件方式，现有 `Image.open` 流程无需分叉；解码专用包满足体积预算 |
| 固定 `pi-heif==1.4.0` | 解码专用发行版兼容 Pillow 12.2；完整 wheel 的 libheif 静态导入 libx265，导致体积超限，不能手工删 DLL |
| 拖拽反馈复用现有状态卡片，通过动态属性切换边框/背景并更新状态文字 | 不引入新布局，提供稳定、非仅颜色的反馈，离开/放下后可恢复 |
| PFX 用 SecureString 导入临时 runner 证书存储，再按 thumbprint 签名 | 避免把证书密码作为 SignTool 命令行参数；结束时删除证书和临时文件 |
| SHA-256 去重在专用导入线程中运行，批次串行排队 | 完整内容哈希可能读取数百 MB，不能冻结 Qt UI；串行批次避免基于旧选择快照竞争 |
| 单个不可读文件按文件名报告，其他文件继续添加 | 对整批回滚会让一个坏文件阻塞所有正常照片；部分成功更可恢复 |
| 全局异常日志脱敏当前任务登记值与用户目录 | 在保留堆栈结构的同时覆盖应用已知隐私；不承诺识别任意第三方文本的语义 |
| 明确清空会保存默认表单值，生成成功保留地点/日期偏好 | “清空”尊重用户主动重置，“连续制作”减少重复输入 |
| Authenticode 只签 `PGT.exe`，时间戳 URL 可配置 | 不重签第三方二进制；默认采用 Microsoft 文档示例，部署方可替换服务 |

## 遇到的问题
| 问题 | 解决方案 |
|------|---------|
| Git 状态检查不可用 | 记录限制，依靠精确文件清单、测试与构建验证 |
| PowerShell 传给 `python -c` 的字面 `\n` 导致 SyntaxError | 改用单行列表推导式；项目代码和模板未受影响 |
| 设计决策大补丁因原句不匹配未应用 | 先定位精确段落，拆分后成功回写；没有产生半应用状态 |

## 资源
- `README.md`：现有功能、基准、日志位置、打包与右键菜单说明
- `requirements.txt` / `requirements-dev.txt`：运行和测试依赖
- pillow-heif 官方文档：Pillow plugin 的 `register_heif_opener()` 注册方式
- PyPI pillow-heif/pi-heif 项目页（2026-11-26 核验）：版本兼容、Python 要求、Windows wheels、解码专用差异与许可证
- GitHub Docs：Actions secrets 在条件表达式中的环境变量模式、Base64 二进制 secret 注意事项
- Microsoft Learn：SignTool `/fd SHA256`、`/tr`、`/td SHA256` 与 `verify /pa /tw`
