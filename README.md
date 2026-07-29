# 货源投放周策略转导入 Excel

这是一个本地 Web 工具，用于读取周策略表里的“本期投放策略”，按用户选择的系统导入格式批量生成真正的 Excel 97-2003 `.xls` 导入表。

## 启动

安装运行依赖：

```powershell
python -m pip install -r strategy_excel_tool\requirements.txt
```

### 测试网页版

在项目根目录运行：

```powershell
python strategy_excel_tool\server.py
```

然后打开：

```text
http://127.0.0.1:8010
```

如需换端口：

```powershell
$env:STRATEGY_TOOL_PORT=8020
python strategy_excel_tool\server.py
```

### 桌面程序模式

在项目根目录运行：

```powershell
python strategy_excel_tool\desktop_launcher.py
```

它会自动找一个可用端口并打开浏览器。关闭启动窗口即可退出。

如果默认端口被单位电脑占用或限制，程序会自动尝试后续端口，例如 `8011`、`8012`、`8013`。请以程序自动打开的浏览器地址为准。

## 模板准备

真实模板是运行时私有业务文件，不进入公开 Git。源码运行时默认从项目根目录读取：

```text
private_templates\
```

原始 `.xls` 文件名必须是：

```text
档位+区倒+市场类型.xls
档位+区县.xls
档位+区县+商圈.xls
档位+市场类型.xls
```

配套 `.xlsx` 文件名必须是：

```text
档位+区倒+市场类型.xlsx
档位+区县.xlsx
档位+区县+商圈.xlsx
档位+市场类型.xlsx
```

商品内码映射文件名必须是：

```text
商品内码实例.xls
```

程序按商品编码从该文件读取商品内码。该映射优先于策略表中的商品内码；如果某个商品没有映射，程序会显示黄色警告并列出对应策略行，商品内码留空但仍允许导出。同一商品编码对应多个不同内码时会阻止导出。

公开的 `strategy_excel_tool/templates/` 只保留[模板说明](templates/README.md)。如果需要把私有模板放在其他位置，可设置环境变量 `STRATEGY_TOOL_TEMPLATE_DIR`。

Windows PowerShell：

```powershell
$env:STRATEGY_TOOL_TEMPLATE_DIR="D:\private\strategy_templates"
python strategy_excel_tool\server.py
```

macOS/Linux：

```bash
export STRATEGY_TOOL_TEMPLATE_DIR="$HOME/private/strategy_templates"
python3 strategy_excel_tool/server.py
```

相对路径会以项目根目录为基准解析。网页中的模板状态提示会显示当前实际加载目录和缺失文件。

## 使用流程

1. 上传周策略表 `.xlsx`。
2. 在下拉框选择导出格式。
3. 点击“解析并预览校验”。
4. 检查校验提示、结构化规则 JSON 和导出预览。
5. 确认无误后点击“导出 Excel”。

工具不会在上传后直接导出，必须先完成预览校验。

## 支持规则

- 自动读取包含 `商品编码`、`商品名称`、`本期投放策略` 的工作表。
- 自动按 `商品编码` 从私有 `商品内码实例.xls` 填写商品内码；缺失时警告并留空，一对多冲突时阻止导出。
- 策略列按 `本期投放策略` 表头动态定位，不绑定 O 列或 P 列；新增 `在途` 等中间字段不会影响解析。
- 自动识别模板中的商品字段、区县/市场类型/商圈字段，以及 `1档` 到 `30档` 列。
- 档位列可以倒序，例如 `30档` 到 `1档`。
- 最终文件是旧版 `.xls`，不会用修改扩展名的方式伪装格式。
- 四种导出格式分别绑定原模板隐藏首行中的系统类型标识，导出前后都会校验该行仍处于隐藏状态且标识正确。
- 导出文件第二行的投放周期会与文件名中的 `YYYYMMW?(MM.DD-MM.DD)` 同步。
- `1-30档3条` 表示 1档 到 30档每个档位都填 `3`。
- 支持多个县区并列，例如 `F县、G县：1-30档1条。`
- 支持多个类型并列，例如 `城市、县城：20-25档2条，26-30档3条。`
- 没有县区/市场类型限定的策略会套用模板中的全部维度行。
- `无库存`、空白或未识别到档位规则的商品会保留输出，所有档位填 `0`，并在校验里提示。
- `无库存`、`库存不足`、`停投`、`暂停投放` 等明确 0 投放文本会自动按 0 处理，不再作为异常刷屏。
- 导出时会先把 Excel 保存到 `outputs` 文件夹，再在页面显示保存路径；即使浏览器下载被拦截，也可以直接去该文件夹取文件。

## 自测

```powershell
python strategy_excel_tool\server.py --self-test
```

自测会验证批量展开、规则解析、未命中填 0、真实 `.xls` 文件头，以及隐藏首行类型标识。

## 打包成 exe

打包需要开发依赖。先安装：

```powershell
python -m pip install -r strategy_excel_tool\requirements-dev.txt
```

然后在项目根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File strategy_excel_tool\build_desktop.ps1
```

打包产物会生成到：

```text
dist\卷烟投放策略转excel工具
```

打包脚本会同时生成 `卷烟投放策略转excel工具` 文件夹和 `卷烟投放策略转excel工具.zip`。把整个文件夹复制到其他电脑，或解压 ZIP 后使用；双击里面的 `卷烟投放策略转excel工具.exe` 启动。私有模板放在同级 `private_templates` 文件夹，导出文件会保存在同级 `outputs` 文件夹。

## 在另一台电脑恢复开发环境

1. 克隆公开 Git 仓库。
2. 安装 Python 3。
3. 运行 `python -m pip install -r strategy_excel_tool/requirements.txt`。
4. 在项目根目录创建 `private_templates/`。
5. 通过单位内部渠道把八个真实模板和 `商品内码实例.xls` 复制到该目录。
6. 运行 `python strategy_excel_tool/server.py`。

Windows 和 macOS 使用相同的源码、目录结构和环境变量。Windows 桌面版使用 `build_desktop.ps1` 打包；macOS 开发时可直接运行源码，后续可单独增加 macOS 打包脚本。

## 公开 Git 边界

可以提交 Python 源码、公开说明、依赖清单和构建脚本。不要提交 `private_templates/`、`outputs/`、`dist/`、`build/`、用户策略表或生成的导入文件。
