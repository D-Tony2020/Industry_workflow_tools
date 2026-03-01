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

# ===== 映射表（料号清单）列名 =====
# 新结构: 序号 | 久益料号 | 生久料号 | 品名规格 | 备注（5列，多sheet）
MAP_COL_PRODUCT_CODE = "久益料号"      # 工厂系统编码 (如 J00010001)
MAP_COL_PRODUCT_NAME = "品名规格"      # 品名规格描述
MAP_COL_PRODUCT_SPEC = "生久料号"      # 客户料件编号 (如 YY60030058) - 映射键
# 固定列位置（0-based），用于跨sheet读取（列C表头名因供应商不同而异）
MAP_COL_IDX_JY_CODE = 1    # B列: 久益料号
MAP_COL_IDX_CUSTOMER = 2   # C列: 客户料号（生久料号/甬阅料号/...）
MAP_COL_IDX_DESC = 3       # D列: 品名规格

# ===== PDF采购单中需要提取的字段 =====
PDF_MAPPING_KEY = "料件编号"  # PDF中用于映射的字段名

# ===== 输出相关 =====
OUTPUT_ORDER_TYPE = "生产工单"    # 工单分类固定值
QUANTITY_SAFETY_MARGIN = 5       # 数量安全余量（采购数量 + 5）

# 输出Excel列（导入产品明细模板，19列）
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

# ===== 图纸比对相关 =====
DRAWING_PRINT_FOLDER = "待打印"                          # 待打印文件夹名
DRAWING_REV_PATTERN = r"(?:版本号|版本)\s*(?:REV)?\s*(\S+)"  # 图纸PDF中REV字段正则
