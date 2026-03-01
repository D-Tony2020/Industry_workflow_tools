# 项目：Industry Workflow Tools

工厂生产流程辅助工具集合。当前包含一个子工具：`factory_order_tool`。

## factory_order_tool - 工厂订单转换工具

### 功能

将客户（生久科技）的采购单 PDF 自动转换为工厂 ERP 系统可导入的 Excel 文件。

核心流程：PDF 解析 → 编码映射（客户料号 → 工厂编号）→ Excel 输出（19列模板）

### 技术栈

- Python 3 + Tkinter GUI
- pdfplumber：PDF 表格解析
- openpyxl：Excel 读写
- PyInstaller：打包为 Windows .exe

### 模块结构

```
factory_order_tool/
├── main.py          # GUI 主界面 (OrderConverterApp, Tkinter)
├── pdf_parser.py    # PDF 解析 (pdfplumber, 适配生久科技采购单6/7列表格)
├── code_mapper.py   # 编码映射 (读取 mapping_table.xlsx, 客户料号→工厂编号)
├── excel_writer.py  # Excel 输出 (openpyxl, 19列工厂导入模板)
├── config.py        # 配置常量 (列名定义, 路径, 输出模板列)
├── version.py       # 版本信息 (VERSION, APP_NAME, BUILD_DATE)
├── mapping_table.xlsx  # 映射表数据文件
├── requirements.txt    # 依赖: pdfplumber, openpyxl
├── build.bat           # 打包脚本
└── 订单转换工具.spec    # PyInstaller 打包配置
```

### 关键业务逻辑

- **PDF 格式**：生久科技采购单，每个项目占2行（主行+续行），料件编号以 `YY` 开头
- **映射表**：`mapping_table.xlsx`，键=产品规格（客户料号），值=产品编号+名称+工艺路线
- **输出格式**：19列（`config.py:OUTPUT_COLUMNS`），数量和单价自动转数字
- **GUI 预览**：绿色=映射成功，红色=未映射

### 开发约定

- 版本号统一在 `version.py` 中维护
- 配置常量统一在 `config.py` 中定义
- 打包使用 `build.bat`，确保在虚拟环境中构建
- 当前版本：v1.0.0
