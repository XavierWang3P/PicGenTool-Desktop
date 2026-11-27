# 进度日志

## 会话：2026-11-26

### 当前状态
- **阶段：** 7 - 综合验证与交付（已完成；远端 CI/真实签名待外部环境）
- **开始时间：** 2026-11-26

### 已执行操作
- 加载规划、头脑风暴与实施计划技能要求。
- 初始化持久化计划文件。
- 盘点项目顶层结构、README、运行依赖和开发依赖。
- 确认当前目录不是 Git 仓库。
- 阅读主入口、主窗口、图片/Word/模板服务、日志、路径、右键菜单脚本、PyInstaller 配置和现有测试。
- 确认 1–30 号模板文件当前齐全。
- 通过规划文件格式校验（7 个阶段）。
- 运行基线测试：42 passed in 3.57s。
- 量化现有分发：901 个文件，共 119.01 MiB；记录最大依赖项。
- 审阅现有 GitHub Actions：当前直接构建发布，尚未运行测试。
- 验证模板占位符抽样，确认可建立严格的模板完整性规则。
- 查询并应用桌面拖拽反馈与状态可访问性指导。
- 完成阶段 1 需求与代码勘察，进入设计与实施计划阶段。
- 编写质量与发布优化设计文档及 10 个任务的实施计划。
- 完成独立构建者检查并处理 5 个关键歧义：异步哈希、部分导入、异常脱敏、清空语义、签名范围。
- 自查实施计划：13 项用户需求均有对应任务、测试和验收命令；10 个任务接口一致。
- 再次通过持久化计划格式校验（7 个阶段）。
- 完成 Task 1：HEIC/HEIF 解码、内容 SHA-256 去重、不可读文件部分成功、Qt 异步导入队列。
- pillow-heif 初期固定为 1.8.1；后续因完整 wheel 的 libx265 静态依赖导致体积超限，最终采用官方解码专用 pi-heif 1.4.0。
- 完成 Task 2：模板轻量/深度自检、CLI 门禁和启动禁用生成状态。
- 自检发现并修复 `Template-23.docx` 的 `{{%image1}}` 损坏标签；原件备份在批准的临时目录，修复件通过结构和真实渲染验证。
- 完成 Task 3：Word 模板在 GenerationWorker 线程准备，纯图片阶段在 future 中并行；所有 docxtpl 对象保持创建/消费线程一致。
- 图片 30 张中位数 190.063 ms，对比此前 188.833 ms 为约 0.65% 波动，无实质回归。
- 完成 Task 4：日志改为 5 MiB/3 备份轮转，并安装主线程/后台线程全局异常处理器。
- 异常堆栈按最长匹配脱敏当前任务标题、地点、图片/输出路径和用户目录；处理器故障回退原 hook。
- 完成 Task 5：Explorer 注册状态服务、HEIC/HEIF 菜单覆盖、CLI 共享实现和打包态保守自愈。
- 开发机仅做只读检查，当前既有注册状态为 `partial`；未写入注册表。
- 完成 Task 6：QSettings 恢复表单偏好、成功/明确清空的不同语义、最近图片目录和拖拽高亮反馈。
- 所有窗口测试改用临时 INI 注入，测试不写真实用户设置；照片列表和输出路径不落盘。
- 完成 Task 7：非敏感诊断快照、复制入口、日志目录入口和主窗口“关于 / 诊断”按钮。
- UI/UX 检查：模态焦点进入复制按钮，Tab 顺序按视觉顺序，所有新按钮有 accessible name；拖拽有文字和按钮替代，不依赖颜色。
- 完成 Task 8：Windows Python 3.11/3.12 测试矩阵、compileall/pytest/模板门禁和 build.needs=test。
- PyInstaller 固定到 `>=6.17,<7.0`；本地 CI 等价命令均通过。当前目录无 Git，远端 Actions 仍需推送后验证。
- 完成 Task 9：one-dir 产物 944 files / 113.09 MiB，冻结 HEIC 解码/模板/导入烟测通过；签名脚本与受控 WhatIf 已验证，真实 SDK/证书不可用。
- Task 10 用户与发布文档已更新；最终语法检查、30/30 模板、94 项测试通过；真实模板 1/10/23/30 导出通过。
- 独立审查提出 6 个重要问题与 4 个加固项，均核实并修复：导入竞态、Word 路径泄露、清理线程悬挂、结果文案覆盖、tag 重写、SignTool 定位，以及进度、诊断详情、冻结烟测和 IPC 分片。
- 修复后全量 105 项测试通过；冻结 HEIC/Word/GUI 烟测通过；最终 944 files / 113.10 MiB。
- 二次复核发现保存对话框 Qt 重入空隙及清空确认/状态问题，新增 `_generation_starting` 全入口门控、清空确认与最终恢复、诊断清理保护；当前 108 项测试通过。
- 最终独立复核无 Critical/Important；接受保存路径补测通过，忙碌 IPC 返回 `busy`，避免错误确认已接收图片。
- 全部 13 项优化和审查修复完成；最终包 944 files / 113.10 MiB。未提交/推送 Git，远端 CI 与真实签名留待外部环境验证。

### 测试结果
| 测试 | 预期 | 实际 | 状态 |
|------|------|------|------|
| 规划初始化 | 生成三个规划文件 | 已生成并填入首轮信息 | 通过 |
| 基线 pytest | 全部测试通过 | 42 passed in 3.57s | 通过 |
| 体积基线 | 可重复报告文件数和总大小 | 901 files / 119.01 MiB | 通过 |
| Task 1 RED | 新测试因缺失能力失败 | 7 failed, 22 passed | 通过 |
| Task 1 定向 GREEN | HEIF/去重/CLI/UI 测试通过 | 29 passed | 通过 |
| Task 1 全量回归 | 全部测试通过 | 49 passed in 2.01s | 通过 |
| Task 2 RED | 模板模型、自检和启动门禁缺失 | 6 failed | 通过 |
| Task 2 定向 GREEN | 模板、真实导出和窗口测试通过 | 13 passed in 2.32s | 通过 |
| 模板 CLI | 深度校验全部真实模板 | 30/30 templates valid | 通过 |
| Template-23 DOCX | 包结构与段落不变、23 图可生成 | All validations PASSED | 通过 |
| Task 2 全量回归 | 全部测试通过 | 56 passed in 3.12s | 通过 |
| Task 3 RED | PreparedWordExport 模型和拆分接口缺失 | collection error | 通过 |
| Task 3 定向 GREEN | worker/exporter/真实模板 | 17 passed in 1.71s | 通过 |
| Task 3 重复门禁 | GenerationWorker 连续 5 轮 | 6 passed × 5 | 通过 |
| Task 3 图片基准 | 30 张、3 次 | median 190.063 ms | 通过 |
| Task 3 全量回归 | 全部测试通过 | 62 passed in 2.99s | 通过 |
| Task 4 RED | 异常处理模块缺失 | collection error | 通过 |
| Task 4 定向 GREEN | 轮转、异常、worker | 13 passed in 0.47s | 通过 |
| Task 4 全量回归 | 全部测试通过 | 68 passed in 3.24s | 通过 |
| Task 5 RED | Explorer 服务模块缺失 | collection error | 通过 |
| Task 5 定向 GREEN | fake 注册表和 IPC/CLI | 13 passed in 0.42s | 通过 |
| Task 5 全量回归 | 全部测试通过 | 73 passed in 3.11s | 通过 |
| Task 6 RED | 窗口尚不接受设置服务注入 | 12 failed, 4 passed | 通过 |
| Task 6 定向 GREEN | 设置、窗口、IPC/UI | 20 passed in 0.74s | 通过 |
| Task 6 全量回归 | 全部测试通过 | 81 passed in 3.32s | 通过 |
| Task 7 RED | 诊断模型/服务/对话框缺失 | collection error | 通过 |
| Task 7 定向 GREEN | 诊断与窗口 UI | 11 passed in 0.49s | 通过 |
| Task 7 全量回归 | 全部测试通过 | 84 passed in 3.42s | 通过 |
| Task 7 无头布局 | 最小主窗口/诊断对话框 | 440×520 / 520×400，无重叠 | 通过 |
| Task 8 RED | workflow 缺失 test job | 2 failed | 通过 |
| Task 8 契约 GREEN | CI job/矩阵/门禁 | 2 passed | 通过 |
| Task 8 本地等价门禁 | compileall + pytest + templates | 86 passed; 30/30 valid | 通过 |
| Task 9 契约 | 体积、签名、workflow、HEIF | 10 passed | 通过 |
| 最终打包烟测 | 冻结 HEIC 解码、模板与核心导入 | HEIF sample decode passed; 30/30 valid | 通过 |
| 最终体积预算 | one-dir，130 MiB 上限 | 944 files / 113.09 MiB | 通过 |
| 签名 WhatIf | 本机无 SDK 时真实调用；受控 SignTool 替身 | 真 SDK 缺失退出 1；替身计划验证通过 | 部分验证 |
| Task 9 全量回归 | 全部测试 | 94 passed in 3.79s | 通过 |
| Task 10 最终回归 | 全量、语法、模板 CLI | 94 passed in 3.41s; compileall; 30/30 | 通过 |
| 审查后回归 | 竞态、隐私、清理、IPC、CI 等 | 105 passed in 3.86s | 通过 |
| 保存框重入修复回归 | 生成门控、清空确认、状态恢复、诊断 | 108 passed in 4.02s | 通过 |
| 独立最终复核 | 审查全部重要问题与保存框重入 | 无 Critical/Important | 通过 |
| 最终冻结综合烟测 | HEIC + 模板 + 1 图 Word + GUI | 全部 passed | 通过 |
| 最终包预算 | 修复后 one-dir | 944 files / 113.10 MiB | 通过 |
| 真实模板导出 | 1/10/23/30 张 | 4 passed in 1.77s | 通过 |
| 最终图片基准 | 同机三次中位数 | 1=22.083; 10=113.939; 30=211.911 ms | 通过（较旧 after 基线略慢） |

### 错误
| 错误 | 解决方案 |
|------|---------|
| `git diff --stat`：Not a git repository | 停止重试，后续使用局部文件清单与测试验证 |
| 模板探测命令 SyntaxError | 确认为 `python -c` 换行转义问题；下一次使用无换行表达式 |
| 设计决策大补丁校验失败 | 未修改任何文件；改为先定位原文并拆分补丁 |
| TDD 技能无法加载 | 使用实施计划内的显式失败测试、最小实现和任务/全量测试门禁 |
| worktree/SDD 脚本要求 Git 仓库 | 不初始化 Git；建立本计划专属手工 SDD 账本并记录裁定 |
| pillow-heif 1.9.0 安装失败 | 根因是其要求 Pillow>=13.0；初期使用 1.8.1，最终采用官方解码专用 pi-heif 1.4.0 满足体积预算 |
| 完整 HEIF 包体积超限 | 139.23 MiB；PE 导入表证明 libx265 不能直接删除，换用 pi-heif 并过滤未使用 QtPdf DLL，降至 113.09 MiB |
| 本机 `signtool.exe` 缺失 | 真实 WhatIf 明确退出 1 且未修改 EXE；受控替身仅验证计划分支，真实 SDK/证书验签留待 CI |
| 独立审查 6 个重要问题 | 逐项 RED 复现并修复；新增 11 项回归测试，审查定向 35 项通过 |
| IPC 确认读取测试超时 | Qt 探针证明服务端已发 `ok`；测试在 `waitUntil` 消费 readyRead 后误等下一次信号，改为捕获首个确认事件 |
