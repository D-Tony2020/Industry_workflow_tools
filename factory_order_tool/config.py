"""配置常量"""
import os
import sys

# 获取程序所在目录（兼容PyInstaller打包后的路径）
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

# 映射表文件路径（与exe同目录）
MAPPING_TABLE_PATH = os.path.join(APP_DIR, "mapping_table.xlsx")

# ===== 映射表（产品定义）列名 =====
# 映射表结构: 序号 | 库存数量 | 产品编号 | 产品名称 | 单位 | 产品规格 | 工艺路线 | ...
MAP_COL_PRODUCT_CODE = "产品编号"      # 工厂系统编码 (如 J00010038)
MAP_COL_PRODUCT_NAME = "产品名称"      # 工厂产品名称
MAP_COL_PRODUCT_SPEC = "产品规格"      # 客户料件编号 (如 YY60030058) - 映射键
MAP_COL_PROCESS_ROUTE = "工艺路线"     # 工艺路线名称
MAP_COL_UNIT = "单位"                  # 单位

# ===== PDF采购单中需要提取的字段 =====
PDF_MAPPING_KEY = "料件编号"  # PDF中用于映射的字段名

# ===== 输出Excel列（导入产品明细模板）=====
OUTPUT_COLUMNS = [
    "产品编号",
    "产品名称",
    "产品规格",
    "数量",
    "计划开始时间",
    "计划结束时间",
    "工艺路线名称",
    "工序列表",
    "备注",
    "更新",
    "工单分类",
    "供应商",
    "供应商名称",
    "供应商联系人",
    "供应商联系电话",
    "收货地址",
    "采购单价",
    "客户选择",
    "关联产品",
]
