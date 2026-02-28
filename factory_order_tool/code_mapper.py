"""编码映射模块 - 读取产品定义Excel，执行客户编码到工厂编码的映射"""
import os
from openpyxl import load_workbook
from config import (
    MAPPING_TABLE_PATH,
    MAP_COL_PRODUCT_CODE,
    MAP_COL_PRODUCT_NAME,
    MAP_COL_PRODUCT_SPEC,
    MAP_COL_PROCESS_ROUTE,
    MAP_COL_UNIT,
)


def load_mapping_table(path=None):
    """
    读取产品定义Excel，返回映射字典。

    映射键: 产品规格（客户料件编号，如 YY60030058）
    映射值: dict 包含产品编号、产品名称、工艺路线等

    返回:
        mapping: dict - {客户料件编号: {产品编号, 产品名称, 工艺路线, 单位}}
    """
    path = path or MAPPING_TABLE_PATH
    if not os.path.exists(path):
        return {}

    wb = load_workbook(path, read_only=True)
    ws = wb.active

    # 找到列索引
    header = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
    try:
        spec_idx = header.index(MAP_COL_PRODUCT_SPEC)
        code_idx = header.index(MAP_COL_PRODUCT_CODE)
        name_idx = header.index(MAP_COL_PRODUCT_NAME)
    except ValueError:
        wb.close()
        return {}

    # 可选列
    route_idx = header.index(MAP_COL_PROCESS_ROUTE) if MAP_COL_PROCESS_ROUTE in header else None
    unit_idx = header.index(MAP_COL_UNIT) if MAP_COL_UNIT in header else None

    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        spec = str(row[spec_idx]).strip() if row[spec_idx] else ""
        code = str(row[code_idx]).strip() if row[code_idx] else ""
        name = str(row[name_idx]).strip() if row[name_idx] else ""

        if not spec or not code:
            continue

        route = str(row[route_idx]).strip() if route_idx is not None and row[route_idx] else ""
        unit = str(row[unit_idx]).strip() if unit_idx is not None and row[unit_idx] else ""

        mapping[spec] = {
            "产品编号": code,
            "产品名称": name,
            "工艺路线": route,
            "单位": unit,
        }

    wb.close()
    return mapping


def apply_mapping(items, mapping):
    """
    对解析出的订单项目应用编码映射，生成输出行。

    参数:
        items: list[dict] - PDF解析出的项目列表（含料件编号、采购数量等）
        mapping: dict - 映射字典

    返回:
        output_rows: list[dict] - 输出模板格式的行列表
        unmapped: list[str] - 未找到映射的料件编号列表
    """
    output_rows = []
    unmapped = []

    for item in items:
        customer_code = item.get("料件编号", "").strip()
        product_info = mapping.get(customer_code)

        row = {
            "产品编号": "",
            "产品名称": "",
            "产品规格": customer_code,
            "数量": item.get("采购数量", ""),
            "计划开始时间": "",
            "计划结束时间": item.get("出货日期", ""),
            "工艺路线名称": "",
            "工序列表": "",
            "备注": item.get("备注", ""),
            "更新": "",
            "工单分类": "",
            "供应商": "",
            "供应商名称": "",
            "供应商联系人": "",
            "供应商联系电话": "",
            "收货地址": "",
            "采购单价": item.get("单价", ""),
            "客户选择": "",
            "关联产品": "",
            # 额外保留PDF原始字段供预览用
            "_映射状态": "未映射",
            "_品名": item.get("品名", ""),
            "_图号": item.get("图号", ""),
            "_规格": item.get("规格", ""),
            "_含税金额": item.get("含税金额", ""),
        }

        if product_info:
            row["产品编号"] = product_info["产品编号"]
            row["产品名称"] = product_info["产品名称"]
            row["工艺路线名称"] = product_info["工艺路线"]
            row["_映射状态"] = "已映射"
        else:
            if customer_code:
                unmapped.append(customer_code)

        output_rows.append(row)

    return output_rows, unmapped


def get_mapping_stats(output_rows):
    """统计映射结果"""
    total = len(output_rows)
    mapped = sum(1 for r in output_rows if r.get("_映射状态") == "已映射")
    return total, mapped, total - mapped
