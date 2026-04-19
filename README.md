# PicGenTool-Desktop

一个基于 `Python + PyQt6` 的桌面工具，用于将活动照片批量生成 Word 文档。

## 功能简介

- 输入活动标题、活动地点、活动日期
- 选择或拖拽 1 到 30 张图片
- 按模板生成活动照片 Word 文档
- 可选生成后自动打开文件

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

## 开发调试建议

### 语法检查

```bash
python -m compileall main.py app.py config.py ui services models utils
```

### 日志位置

运行日志默认输出到：

`logs/app.log`

### 常见调试点

- 如果程序可以启动但生成失败，先检查 `templates/` 中是否存在对应图片数量的模板文件
- 如果图片无法添加，检查文件扩展名是否在 `config.py` 的 `SUPPORTED_FORMATS` 中

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

打包完成后，输出通常位于：

`dist/图片文档生成工具/`

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

- `templates/*.docx`
- `favicon.png`
- `config.py`

如果你后续新增模板或资源文件，请确认 `pgt.spec` 中的 `datas` 配置同步更新。

### 5. Windows 打包建议

- 请在 Windows 环境下打包 Windows `exe`
- 不建议追求 `one-file`，因为启动速度通常明显更慢
- 推荐直接分发整个 `dist/图片文档生成工具/` 目录

## 使用说明

1. 启动程序
2. 输入活动标题
3. 输入活动地点
4. 选择活动日期
5. 添加图片
6. 选择是否自动打开文件
7. 点击“生成文档”

## 注意事项

- 当前模板机制仍依赖 `templates/Template-1.docx` 到 `templates/Template-30.docx`
- 图片数量必须和对应模板数量匹配
- 输出文件名会根据日期和活动标题自动生成默认名称
- 项目已不再使用 `.ui` 设计器文件，当前界面由代码直接构建

## 后续可优化方向

- 进一步减少模板数量，改为统一模板动态布局
- 增加图片预览和排序能力
- 增加导出过程中的取消任务能力
- 继续弱化模板驱动，提升布局服务的独立性
