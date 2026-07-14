"""PDF采购单解析模块 - 使用pdfplumber提取表格数据

生久科技采购单PDF表格结构（6列，每个项目占2行）：
  主行:   [项次, "YY编号 规格\n品名\n图号", None, "单价\n数量\n单位", "金额\n日期\n税率", 交期回复]
  续行:   [None, "备注内容", None, None, None, None]

注: 交期回复列（最后一列）通常为空，用户可通过PDF编辑器在此列填写最新图纸版本号

v1.2.4 改进:
  部分订单 PDF 的某页表格线异常，pdfplumber 会把整页拆成多个零碎子表，
  导致项次行的 YY 编号被丢出 cell（曾出现单页整段 8 项缺失）。
  当某页通过 _parse_table 解析出 0 个带 YY 料件编号的项目，
  但纯文本中能匹配到「项次 YY编号」锚点行时，
  触发 _parse_text_fallback 基于 page.extract_text() 重新解析该页。

v1.2.5 鲁棒性建设（解析层防御性改造）:
  - 项次连续性自检: 解析完成后检测项次集合在 min..max 范围内是否有缺失，
    若有缺失自动对全 PDF 文本触发兜底解析、按 (项次, YY 编号) 去重补缺。
  - 合计对账: 提取 PDF 末尾「合计数量 / 采购总含税金额」与所有项目实际汇总比对。
  - parse_purchase_order 改为返回三元组 (header, items, warnings)，
    把 "静默失败" 转为 "显式告警", GUI 据此向用户主动提示。
"""
import pdfplumber
import re


# ===== 文本兜底解析所需正则 (v1.2.4) =====
# 锚点行: "<项次> <YY编号> <单价> <含税金额>" 一行四列数字+编号
_FALLBACK_MAIN_RE = re.compile(r"^\s*(\d+)\s+(YY\d{8,})\s+([\d,.]+)\s+([\d,.]+)\s*$")
_FALLBACK_QTY_DATE_RE = re.compile(r"^([\d,.]+)\s+(\d{4}/\d{1,2}/\d{1,2})\s*$")
_FALLBACK_VER_RE = re.compile(r"^[A-Z][/.]?\d*$")
_FALLBACK_UNIT_TAX_RE = re.compile(r"^(PCS|个|条|套)\s+([\d.]+%)\s*$")
_TRAILING_VER_RE = re.compile(r"\s+([A-Z][/.]?\d+)$")
# 用于判定本页文本中是否含订单项行（决定是否需要兜底）
_PAGE_HAS_ITEM_RE = re.compile(r"^\s*\d+\s+YY\d{8,}", re.MULTILINE)

# ===== 合计行抓取（v1.2.5）=====
_TOTAL_QTY_RE = re.compile(r"合计数量[:：]?\s*([\d,]+(?:\.\d+)?)")
_TOTAL_AMOUNT_RE = re.compile(r"采购总含税金额[:：]?\s*([\d,]+(?:\.\d+)?)")
# 数量/金额对账的容忍误差（绝对值 < 此值视作一致，处理浮点）
_AMOUNT_TOLERANCE = 0.05
_QTY_TOLERANCE = 0.5


def parse_purchase_order(pdf_path):
    """
    解析生久科技采购单PDF。

    返回:
        header_info: dict - 采购单头部信息
        items: list[dict] - 每行项目的字段字典
        warnings: dict - 解析层自检结果（v1.2.5），结构:
            {
              "missing_seq": [int],           # 项次连续性自检发现的缺失项次
              "fallback_recovered": [int],    # 已通过全文兜底成功补回的项次
              "fallback_unresolved": [int],   # 兜底也没找回的项次
              "expected_total_qty": float|None,    # PDF 合计行的总数量
              "actual_total_qty": float,           # 实际项目数量汇总
              "expected_total_amount": float|None, # PDF 合计行的总金额
              "actual_total_amount": float,        # 实际项目金额汇总
              "qty_mismatch": bool,               # 数量是否对不上
              "amount_mismatch": bool,            # 金额是否对不上
              "has_warning": bool,                # 任一告警均为 True 时为 True
            }
    """
    header_info = {}
    items = []
    full_text_parts = []

    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        header_info = _extract_header(first_page_text)

        for page in pdf.pages:
            page_items = []

            # 主路径：表格提取
            tables = page.extract_tables() or []
            for table in tables:
                page_items.extend(_parse_table(table))

            # 兜底触发：该页文本里能匹配到订单项锚点行，但表格路径没识别出任何 YY 编号
            page_text = page.extract_text() or ""
            full_text_parts.append(page_text)

            valid_yy = sum(
                1 for it in page_items if it.get("料件编号", "").startswith("YY")
            )
            if valid_yy == 0 and _PAGE_HAS_ITEM_RE.search(page_text):
                page_items = _parse_text_fallback(page_text)

            items.extend(page_items)

    full_text = "\n".join(full_text_parts)

    # v1.2.5 防御性自检: 项次连续性 + 合计对账
    items, warnings = _post_validate(items, full_text)

    return header_info, items, warnings


def _post_validate(items, full_text):
    """v1.2.5 解析后自检与告警生成。

    1. 项次连续性: 若 items 项次集合在 [min..max] 内有缺失,
       对全文触发文本兜底, 按 (项次, YY编号) 去重补缺。
    2. 合计对账: 对比 PDF 合计行 vs 项目实际汇总的数量 / 金额。
    """
    warnings = {
        "missing_seq": [],
        "fallback_recovered": [],
        "fallback_unresolved": [],
        "expected_total_qty": None,
        "actual_total_qty": 0.0,
        "expected_total_amount": None,
        "actual_total_amount": 0.0,
        "qty_mismatch": False,
        "amount_mismatch": False,
        "has_warning": False,
    }

    # 1. 项次连续性自检
    nums = set()
    for it in items:
        try:
            nums.add(int(it.get("项次", "")))
        except (ValueError, TypeError):
            pass

    if nums:
        seq_min, seq_max = min(nums), max(nums)
        missing = sorted(set(range(seq_min, seq_max + 1)) - nums)
        if missing:
            # 发现缺失 -> 全文档兜底重扫一次
            existing_keys = {
                (it.get("项次"), it.get("料件编号")) for it in items
            }
            recovered_items = _parse_text_fallback(full_text)
            recovered_seqs = set()
            for r in recovered_items:
                key = (r.get("项次"), r.get("料件编号"))
                if key in existing_keys:
                    continue
                try:
                    seq_int = int(r.get("项次", ""))
                except (ValueError, TypeError):
                    continue
                if seq_int in missing:
                    items.append(r)
                    recovered_seqs.add(seq_int)
            warnings["missing_seq"] = missing
            warnings["fallback_recovered"] = sorted(recovered_seqs)
            warnings["fallback_unresolved"] = sorted(set(missing) - recovered_seqs)
            # 重新按项次排序
            items.sort(
                key=lambda it: (
                    int(it["项次"]) if str(it.get("项次", "")).isdigit() else 99999
                )
            )

    # 2. 合计对账
    totals = _extract_totals(full_text)
    if "qty" in totals:
        warnings["expected_total_qty"] = totals["qty"]
    if "amount" in totals:
        warnings["expected_total_amount"] = totals["amount"]

    actual_qty = 0.0
    actual_amount = 0.0
    for it in items:
        try:
            actual_qty += float(str(it.get("采购数量", "")).replace(",", "") or 0)
        except ValueError:
            pass
        try:
            actual_amount += float(str(it.get("含税金额", "")).replace(",", "") or 0)
        except ValueError:
            pass
    warnings["actual_total_qty"] = actual_qty
    warnings["actual_total_amount"] = actual_amount

    if warnings["expected_total_qty"] is not None:
        warnings["qty_mismatch"] = (
            abs(warnings["expected_total_qty"] - actual_qty) > _QTY_TOLERANCE
        )
    if warnings["expected_total_amount"] is not None:
        warnings["amount_mismatch"] = (
            abs(warnings["expected_total_amount"] - actual_amount) > _AMOUNT_TOLERANCE
        )

    warnings["has_warning"] = bool(
        warnings["fallback_unresolved"]
        or warnings["qty_mismatch"]
        or warnings["amount_mismatch"]
    )
    return items, warnings


def _extract_totals(text):
    """从全文提取「合计数量 / 采购总含税金额」"""
    out = {}
    m = _TOTAL_QTY_RE.search(text)
    if m:
        try:
            out["qty"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
    m = _TOTAL_AMOUNT_RE.search(text)
    if m:
        try:
            out["amount"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
    return out


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
        "交期回复": "",
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

    # ===== 交期回复列（表格最后一列）=====
    # 用户可能通过PDF编辑器在此列填写最新图纸版本号（如 A02）
    last_cell = cells[-1] if cells else ""
    if last_cell and last_cell != item["项次"]:
        item["交期回复"] = last_cell.strip()

    return item


def _parse_continuation(item, cells):
    """解析续行，补充备注信息 + 版本号（v1.2.9）"""
    col1 = cells[1] if len(cells) > 1 else ""
    if col1:
        # 续行通常是备注（SA/SB/SC...编号开头）
        if item["备注"]:
            item["备注"] += " " + col1.strip()
        else:
            item["备注"] = col1.strip()

    # v1.2.9: 生久新版 PDF 加了"理论交期"列，主行的该 cell 是日期(如 2026-08-01),
    #   续行的对应 cell 是版本号(如 "A03")。这里从续行末列抓版本号 token,
    #   追加到交期回复字段(用换行分隔), 配合 drawing_checker.extract_version_from_reply
    #   从"日期+版本号"混合文本中智能提取。
    #   老 PDF 的续行末列通常为空,或与 col1 相同, 此逻辑不误伤。
    last_cell = cells[-1] if cells else ""
    text = last_cell.strip()
    if text and text != col1.strip():
        # 仅接受纯版本号 token(字母 + 可选分隔符 + 数字, 或纯字母)
        # 严格模式避免误吃"需求4/10"、"6/15"等非版本内容
        if re.match(r"^[A-Z][/.]?\d+$|^[A-Z]$", text):
            if item["交期回复"]:
                item["交期回复"] += "\n" + text
            else:
                item["交期回复"] = text


def _parse_text_fallback(text):
    """基于纯文本扫描的兜底解析器（v1.2.4）。

    场景: 某页表格线异常，pdfplumber 把整页拆成多个零碎子表，
    _parse_table 提不出任何 YY 编号；但 page.extract_text() 的
    纯文本里订单项数据完整存在。

    策略:
      - 用 "项次 YY编号 单价 含税金额" 一行四值作为锚点
      - 锚点上方 1-2 行: 规格首行
      - 锚点下方逐行特征匹配: 数量+日期 / 单位+税率 / 版本号 / 品名 / 备注 / 续行规格
      - 直到撞上下一锚点或扫满 12 行

    输出字段与 _parse_main_row 完全一致，下游 code_mapper / drawing_checker
    无需感知差异。
    """
    items = []
    lines = [ln.rstrip() for ln in text.split("\n")]

    for i, line in enumerate(lines):
        m = _FALLBACK_MAIN_RE.match(line.strip())
        if not m:
            continue

        item = {
            "项次": m.group(1),
            "料件编号": m.group(2),
            "单价": m.group(3).replace(",", ""),
            "含税金额": m.group(4).replace(",", ""),
            "品名": "",
            "图号": "",
            "规格": "",
            "采购数量": "",
            "采购单位": "",
            "出货日期": "",
            "税率": "",
            "备注": "",
            "交期回复": "",
        }

        # 锚点上方 1-2 行查找规格首行
        for k in range(i - 1, max(-1, i - 3), -1):
            prev = lines[k].strip()
            if prev and not _FALLBACK_MAIN_RE.match(prev):
                if "RoHS" in prev or "UL" in prev.upper() or "/" in prev:
                    item["规格"] = prev.rstrip(";").rstrip("\\")
                break

        # 锚点下方扫至下一锚点或满 12 行
        for j in range(1, 12):
            if i + j >= len(lines):
                break
            n = lines[i + j].strip()
            if not n:
                continue
            if _FALLBACK_MAIN_RE.match(n):
                break

            # 数量 + 日期
            qd = _FALLBACK_QTY_DATE_RE.match(n)
            if qd:
                item["采购数量"] = qd.group(1).replace(",", "")
                item["出货日期"] = qd.group(2)
                continue

            # 单位 + 税率
            ut = _FALLBACK_UNIT_TAX_RE.match(n)
            if ut:
                item["采购单位"] = ut.group(1)
                item["税率"] = ut.group(2)
                continue

            # 版本号独占一行（A01 / B/0 / A.1 / A 等）
            if _FALLBACK_VER_RE.match(n) and len(n) <= 5:
                if not item["交期回复"]:
                    item["交期回复"] = n
                continue

            # 品名（含线规关键词）
            if not item["品名"] and ("AWG" in n or "成品线材" in n or "单股" in n):
                # 品名末尾若有版本号 token，拆出来作为交期回复
                tv = _TRAILING_VER_RE.search(n)
                if tv:
                    item["品名"] = n[: tv.start()].strip()
                    if not item["交期回复"]:
                        item["交期回复"] = tv.group(1)
                else:
                    item["品名"] = n
                continue

            # 备注（含 "需求" 且有项目编号前缀或生产/研发/报废类标记）
            if "需求" in n and any(
                tag in n
                for tag in (
                    "SA", "SB", "SC", "SD", "SE", "SF", "SG", "SZ",
                    "生产", "研发", "报废",
                )
            ):
                if item["备注"]:
                    item["备注"] += " " + n
                else:
                    item["备注"] = n
                continue

            # 续行规格（包含 / 或以 ; 结尾，且品名尚未识别）
            if not item["品名"] and ("/" in n or n.endswith(";")):
                tail = n.rstrip(";").rstrip("\\")
                if item["规格"]:
                    item["规格"] += " " + tail
                else:
                    item["规格"] = tail

        items.append(item)

    return items
