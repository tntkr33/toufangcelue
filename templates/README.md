# 模板目录说明

此目录只存放公开说明，不存放真实业务模板。

程序源码运行时默认从项目根目录的 `private_templates/` 读取模板。该目录已被 Git 忽略，需要在每台开发电脑上单独恢复。

## 必需文件

`private_templates/` 中必须同时存在以下九个文件：

```text
档位+区倒+市场类型.xls
档位+区倒+市场类型.xlsx
档位+区县.xls
档位+区县.xlsx
档位+区县+商圈.xls
档位+区县+商圈.xlsx
档位+市场类型.xls
档位+市场类型.xlsx
商品内码实例.xls
```

- `.xlsx` 用于读取表头、档位和维度数据。
- `.xls` 用于生成最终系统导入文件，并保留格式、工作表名称、合并单元格和隐藏系统标识。
- `商品内码实例.xls` 用于按商品编码填写商品内码。文件中必须包含 `商品编码` 和 `商品内码` 表头；缺失映射时程序会警告并留空，同一商品编码对应多个不同内码时会阻止导出。

不要把真实模板提交到公开 Git 仓库，也不要上传到公开 Release。

## 自定义路径

可以通过环境变量 `STRATEGY_TOOL_TEMPLATE_DIR` 指定其他私有目录。

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
