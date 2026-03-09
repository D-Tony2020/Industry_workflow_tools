"""图纸版本比对模块 v1.2.0 - 基于文件名索引的快速版本比对

改进:
  - 预扫描图纸库目录，构建 {YY编号: (路径, 版本号)} 索引
  - 从文件名提取版本号（毫秒级），不再打开PDF
  - 解决扫描版PDF无法识别版本号的问题

工作流:
  1. build_drawing_index() 一次性扫描图纸库，构建索引
  2. 从订单数据中获取每个YY产品的「应有版本号」
  3. O(1) 字典查找替代逐个glob
  4. 严格字符串比对
  5. 匹配的自动复制到待打印文件夹

图纸文件名提取策略:
  - YY编号: 直接搜索 YY\\d{8,} 模式
  - 版本号: 先删除文件名中的 J\\d+ 和 YY\\d+ 片段，再搜索 [A-Z][/.]?\\d+ 模式
  - 兼容: A01/A0/A.1/A/01/B/0 等各种版本号格式
  - 兼容各种人工命名差异（顺序、分隔符、有无"版"字、全角括号等）

推荐命名: J00016025 YY60030362-A01.pdf（可选末尾追加产品类型: J00016025 YY60030362-A01导线.pdf）
"""
import os
import re
import shutil

from config import DRAWING_PRINT_FOLDER


# YY编号提取
YY_CODE_PATTERN = re.compile(r"(YY\d{8,})")

# 用于清理文件名中的J编号和YY编号，以便提取版本号
_CLEAN_JY_PATTERN = re.compile(r"(?:J\d+|YY\d+)")

# 版本号模式: 大写字母 + 可选分隔符 + 1位以上数字（兼容 A0/B0/A.1/A/0 等单位数格式）
_VERSION_PATTERN = re.compile(r"([A-Z][/.]?\d+)")


# ========== 版本号提取 ==========

def extract_version_from_name(product_name):
    """
    从品名规格末尾提取版本号。

    规则: 取最后一个空格后的 token，匹配 [A-Z][/.]?\\d+ 模式
    例: "双剥镀 UL1007#24 黑色2+5 L=360mm B/01" -> "B/01"

    参数:
        product_name: str - 品名规格字符串

    返回:
        str | None - 版本号，或 None 如果未提取到
    """
    if not product_name or not product_name.strip():
        return None

    # 取末尾 token
    token = product_name.strip().split()[-1]

    # 匹配版本号模式: 字母开头 + 可选分隔符 + 数字（含纯字母如 "A"）
    if re.match(r"^[A-Z][/.]?\d*$", token):
        return token

    return None


def extract_version_from_reply(reply_text):
    """
    从交期回复列提取版本号。

    用户通过PDF编辑器填写，格式可能为:
      - 纯版本号: "A02", "B/01"
      - 含前缀: "REV A02"

    参数:
        reply_text: str - 交期回复列内容

    返回:
        str | None - 版本号，或 None 如果为空
    """
    if not reply_text or not reply_text.strip():
        return None
    return reply_text.strip()


def extract_version_from_filename(filename):
    """
    从图纸文件名中提取版本号（宽容模式）。

    策略: 先删除文件名中的 J编号 和 YY编号，再在剩余文本中搜索版本号模式。
    兼容各种人工命名差异。

    示例:
      "J00010183 YY60039112-A03导线.pdf"     → "A03"
      "J00016035 YY60030348-A02 导线.pdf"     → "A02"
      "J00016175 YY60039091-导线-A04.pdf"     → "A04"
      "J00016321 YY60020144-导线（A01）.pdf"  → "A01"
      "J00016312 YY60040718导线-A01.pdf"      → "A01"
      "J00016025 YY60030362-A.pdf"           → "A"  (纯字母版本)
      "J00016025 YY60030362-A导线.pdf"       → "A"  (纯字母版本)
      "J00010272 YY60030215.pdf"              → None
      "YY60030192导线.pdf"                    → None

    参数:
        filename: str - 文件名（不含目录路径）

    返回:
        str | None - 版本号，或 None 如果无法提取
    """
    # 删除J编号和YY编号
    cleaned = _CLEAN_JY_PATTERN.sub("", filename)

    # 主模式: 字母 + 可选分隔符 + 至少1位数字（A01, B/01, A.1, A0）
    match = _VERSION_PATTERN.search(cleaned)
    if match:
        return match.group(1)

    # 回退: 分隔符/括号后的纯字母版本号（如 "-A.pdf", "-A导线.pdf", "（A）.pdf"）
    # 前置: 分隔符(-/空格) 或 括号（/(
    # 后置: 终止符(.pdf/空格/中文/括号)
    fallback = re.search(r'[-\s（(]([A-Z])(?=[.\s\u4e00-\u9fff（()）)])', cleaned)
    if fallback:
        return fallback.group(1)

    return None


# ========== 规范命名生成 ==========

def generate_standard_name(factory_code, yy_code, version):
    """
    生成标准命名的图纸文件名。

    格式: {工厂编号} {客户料号}-{版本号}.pdf
    可在末尾追加产品类型: J00016025 YY60030362-A01导线.pdf

    当 factory_code 为空（未映射物料）时，使用 "???" 作为占位符，
    生成 "??? YY60030362-A01.pdf" 格式，提示用户需先更新映射表。

    参数:
        factory_code: str - 工厂编号 (如 J00016025)，为空则用 ??? 占位
        yy_code: str - 客户料号 (如 YY60030362)
        version: str - 版本号 (如 A01)

    返回:
        str - 标准文件名
    """
    if factory_code:
        return f"{factory_code} {yy_code}-{version}.pdf"
    return f"??? {yy_code}-{version}.pdf"


# ========== 图纸索引 ==========

def build_drawing_index(drawing_dir):
    """
    预扫描图纸库目录，构建 {YY编号: (文件路径, 版本号)} 索引。

    一次性遍历目录中所有PDF文件，从文件名提取YY编号和版本号。
    时间复杂度: O(n) 单次遍历，n为PDF文件数量。

    参数:
        drawing_dir: str - 图纸库目录路径

    返回:
        index: dict - {yy_code: (file_path, version)}
            version 为 None 表示无法提取版本号
        bad_names: list[str] - 含YY编号但无法提取版本号的文件名列表
    """
    index = {}
    bad_names = []

    if not drawing_dir or not os.path.isdir(drawing_dir):
        return index, bad_names

    for fname in os.listdir(drawing_dir):
        # 只处理PDF文件
        if not fname.lower().endswith(".pdf"):
            continue

        fpath = os.path.join(drawing_dir, fname)
        if not os.path.isfile(fpath):
            continue

        # 提取YY编号
        yy_match = YY_CODE_PATTERN.search(fname)
        if not yy_match:
            continue  # 不含YY编号的PDF忽略

        yy_code = yy_match.group(1)

        # 提取版本号（宽容模式）
        version = extract_version_from_filename(fname)

        if version:
            # 同一YY编号多个文件时取文件名排序最后的
            if yy_code not in index or fname > os.path.basename(index[yy_code][0]):
                index[yy_code] = (fpath, version)
        else:
            # 有YY编号但无版本号
            if yy_code not in index:
                index[yy_code] = (fpath, None)
            bad_names.append(fname)

    return index, bad_names


# ========== 核心比对逻辑 ==========

def check_drawings(output_rows, drawing_dir, print_folder=None):
    """
    对订单中的YY产品执行图纸版本比对（v1.2.0 文件名索引版）。

    改进:
    - 预扫描构建索引，O(1)查找替代逐个glob
    - 从文件名提取版本号，不再打开PDF
    - 移除pdfplumber依赖

    参数:
        output_rows: list[dict] - apply_mapping 输出的行列表
        drawing_dir: str - 图纸库目录路径
        print_folder: str | None - 待打印文件夹路径（None则在drawing_dir下创建）

    返回:
        results: list[dict] - 每个项目的比对结果
          status: match / mismatch / no_version / no_drawing / bad_name / skipped
        bad_names: list[str] - 无法提取版本号的文件列表
    """
    results = []

    # 确定待打印文件夹路径
    if print_folder is None:
        print_folder = os.path.join(drawing_dir, DRAWING_PRINT_FOLDER)

    # 清空待打印文件夹（避免残留）
    if os.path.exists(print_folder):
        for f in os.listdir(print_folder):
            fp = os.path.join(print_folder, f)
            if os.path.isfile(fp):
                try:
                    os.remove(fp)
                except Exception:
                    pass
    else:
        os.makedirs(print_folder, exist_ok=True)

    # 一次性构建索引
    drawing_index, bad_names = build_drawing_index(drawing_dir)

    # 去重: 同一个YY编号只比对一次
    seen_codes = set()

    for row in output_rows:
        yy_code = row.get("产品规格", "").strip()
        factory_code = row.get("产品编号", "").strip()

        # 跳过非YY产品
        if not yy_code.startswith("YY"):
            results.append({
                "yy_code": yy_code,
                "order_version": "",
                "local_version": "",
                "drawing_path": "",
                "status": "skipped",
                "message": "非YY产品，跳过",
                "suggested_name": "",
            })
            continue

        # 去重
        if yy_code in seen_codes:
            continue
        seen_codes.add(yy_code)

        # 1. 获取「应有版本号」
        order_version = extract_version_from_reply(row.get("_交期回复", ""))
        if not order_version:
            order_version = extract_version_from_name(row.get("_产品名称", ""))

        if not order_version:
            results.append({
                "yy_code": yy_code,
                "order_version": "",
                "local_version": "",
                "drawing_path": "",
                "status": "no_version",
                "message": "未提供版本号",
                "suggested_name": "",
            })
            continue

        # 2. 从索引查找（O(1)）
        entry = drawing_index.get(yy_code)
        if not entry:
            suggested = generate_standard_name(factory_code, yy_code, order_version)
            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": "",
                "drawing_path": "",
                "status": "no_drawing",
                "message": "未找到图纸",
                "suggested_name": suggested,
            })
            continue

        drawing_path, local_version = entry

        # 3. 文件名版本号缺失
        if not local_version:
            suggested = generate_standard_name(factory_code, yy_code, order_version)
            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": "",
                "drawing_path": drawing_path,
                "status": "bad_name",
                "message": "文件名中无法识别版本号",
                "suggested_name": suggested,
            })
            continue

        # 4. 严格字符串比对
        if order_version == local_version:
            try:
                dest = os.path.join(print_folder, os.path.basename(drawing_path))
                shutil.copy2(drawing_path, dest)
            except Exception:
                pass

            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": local_version,
                "drawing_path": drawing_path,
                "status": "match",
                "message": f"版本一致: {local_version}",
                "suggested_name": "",
            })
        else:
            suggested = generate_standard_name(factory_code, yy_code, order_version)
            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": local_version,
                "drawing_path": drawing_path,
                "status": "mismatch",
                "message": f"本地: {local_version} → 最新: {order_version}",
                "suggested_name": suggested,
            })

    return results, bad_names


# ========== 统计 ==========

def get_check_stats(results):
    """
    统计比对结果。

    返回:
        dict - {total, match, mismatch, no_version, no_drawing, bad_name, skipped}
    """
    stats = {
        "total": len(results),
        "match": 0,
        "mismatch": 0,
        "no_version": 0,
        "no_drawing": 0,
        "bad_name": 0,
        "skipped": 0,
    }

    for r in results:
        status = r.get("status", "")
        if status in stats:
            stats[status] += 1

    return stats


# ========== 批量打印 ==========

def merge_and_print(ordered_paths):
    """
    将多个PDF按顺序合并为单个文件并发送打印。

    使用 pypdfium2 合并，合并文件存放在系统临时目录。
    通过 os.startfile 发送单次打印任务，保证物理打印顺序。

    参数:
        ordered_paths: list[str] - 按表格顺序排列的PDF路径列表

    返回:
        (count, merged_path):
            count: int - 合并的图纸数量（0表示失败）
            merged_path: str | None - 合并文件路径（供确认后清理）
    """
    import tempfile

    paths = [p for p in ordered_paths if os.path.isfile(p)]
    if not paths:
        return 0, None

    # 单个文件直接打印，无需合并
    if len(paths) == 1:
        try:
            os.startfile(paths[0], "print")
        except Exception:
            return 0, None
        return 1, None

    # 多个文件: 倒序合并为单个PDF后打印
    # 倒序原因: 打印机出纸面朝上，先打印的页在最底下，
    # 倒序合并后拿到手从上往下翻正好是表格顺序
    try:
        import pypdfium2 as pdfium

        merged = pdfium.PdfDocument.new()
        for path in reversed(paths):
            src = pdfium.PdfDocument(path)
            merged.import_pages(src)
            src.close()

        # 存到系统临时目录（纯ASCII文件名避免编码问题）
        fd, merged_path = tempfile.mkstemp(suffix=".pdf", prefix="print_merged_")
        os.close(fd)
        merged.save(merged_path)
        merged.close()

        os.startfile(merged_path, "print")
        return len(paths), merged_path
    except Exception:
        return 0, None
