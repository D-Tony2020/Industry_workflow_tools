"""PDF采购单解析模块 - 使用pdfplumber提取表格数据

生久科技采购单PDF表格结构（6列，每个项目占2行）：
  主行:   [项次, "YY编号 规格\n品名\n图号", None, "单价\n数量\n单位", "金额\n日期\n税率", 交期]
  续行:   [None, "备注内容", None, None, None, None]
"""
import pdfplumber
import re


def parse_purchase_order(pdf_path):
    """
    解析生久科技采购单PDF。

    返回:
        header_info: dict - 采购单头部信息
        items: list[dict] - 每行项目的字段字典
    """
    header_info = {}
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        header_info = _extract_header(first_page_text)

        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue
            for table in tables:
                page_items = _parse_table(table)
                items.extend(page_items)

    return header_info, items


def _extract_header(text):
    """从第一页文本中提取采购单头部信息"""
    header = {}
    patterns = {
        "编号": r"编号:\s*(\S+)",
        "采购单号": r"采购单号:\s*(\S+)",
        "供应商": r"供应商:\s*(.+?)(?:\s{2,}|$)",
        "采购日期": r"采购日期:\s*(\S+)",
        "到厂时间": r"到厂时间:\s*(\S+)",
        "联系人": r"联系人:\s*(\S+)",
        "付款条件": r"付款条件:\s*(\S+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            header[key] = match.group(1).strip()
    return header


def _parse_table(table):
    """解析一个PDF表格，提取所有订单项目"""
    items = []
    current_item = None

    for row in table:
        if not row or len(row) < 4:
            continue

        # 清洗
        cells = [str(c).strip() if c else "" for c in row]

        # 跳过表头行
        if _is_header_row(cells):
            continue

        # 跳过合计行
        if any("合计" in c for c in cells):
            continue

        first_cell = cells[0]

        # 新项目（第一列是数字项次）
        if first_cell and first_cell.isdigit():
            if current_item:
                items.append(current_item)
            current_item = _parse_main_row(cells)
        elif current_item and cells[1]:
            # 续行 - 补充备注
            _parse_continuation(current_item, cells)

    if current_item:
        items.append(current_item)

    return items


def _is_header_row(cells):
    """判断是否为表头行"""
    text = " ".join(cells)
    return "项次" in text and ("料件编号" in text or "品名" in text or "规格" in text)


def _parse_main_row(cells):
    """
    解析主行。

    6列结构: [项次, "YY编号 规格\n品名\n图号", 空, "单价\n数量\n单位", "金额\n日期\n税率", 交期]
    7列结构: [项次, "YY编号 规格\n品名\n图号", 空, 空, "单价\n数量\n单位", "金额\n日期\n税率", 交期]
    """
    item = {
        "项次": cells[0],
        "料件编号": "",
        "品名": "",
        "图号": "",
        "规格": "",
        "单价": "",
        "采购数量": "",
        "采购单位": "",
        "含税金额": "",
        "出货日期": "",
        "税率": "",
        "备注": "",
    }

    # ===== 解析列2: "YY编号 规格\n品名\n图号" =====
    col1 = cells[1] if len(cells) > 1 else ""
    if col1:
        lines = col1.split("\n")

        # 第一行: "YY60030058 RoHS/UL1007/24AWG/BLACK/L=360/NA【汇川】;"
        if lines:
            first_line = lines[0].strip()
            yy_match = re.match(r"(YY\d+)\s*(.*)", first_line)
            if yy_match:
                item["料件编号"] = yy_match.group(1)
                item["规格"] = yy_match.group(2).strip().rstrip(";").rstrip("\\")

        # 第二行: 品名（如 "24AWG单股多芯"）
        if len(lines) > 1:
            item["品名"] = lines[1].strip()

        # 第三行: 图号（如 "DX-160711-01"）
        if len(lines) > 2:
            item["图号"] = lines[2].strip()

    # ===== 智能定位数值列（适配6列/7列表格）=====
    # 从cells中找包含"单价\n数量\n单位"格式的列和"金额\n日期\n税率"格式的列
    price_col = None
    amount_col = None
    for i in range(2, len(cells)):
        val = cells[i]
        if not val:
            continue
        if "\n" in val and "PCS" in val.upper():
            price_col = val  # "单价\n数量\n单位"
        elif "\n" in val and "%" in val:
            amount_col = val  # "金额\n日期\n税率"

    if price_col:
        parts = price_col.split("\n")
        if len(parts) >= 1:
            item["单价"] = parts[0].strip().replace(",", "")
        if len(parts) >= 2:
            item["采购数量"] = parts[1].strip().replace(",", "")
        if len(parts) >= 3:
            item["采购单位"] = parts[2].strip()

    if amount_col:
        parts = amount_col.split("\n")
        if len(parts) >= 1:
            item["含税金额"] = parts[0].strip().replace(",", "")
        if len(parts) >= 2:
            item["出货日期"] = parts[1].strip()
        if len(parts) >= 3:
            item["税率"] = parts[2].strip()

    return item


def _parse_continuation(item, cells):
    """解析续行，补充备注信息"""
    col1 = cells[1] if len(cells) > 1 else ""
    if col1:
        # 续行通常是备注（SA/SB/SC...编号开头）
        if item["备注"]:
            item["备注"] += " " + col1.strip()
        else:
            item["备注"] = col1.strip()
