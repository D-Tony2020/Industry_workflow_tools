"""图纸版本比对模块 - 检测客户图纸是否已更新但工厂本地图纸库未同步

工作流:
  1. 从订单数据中获取每个YY产品的「应有版本号」（交期回复列优先 → 品名规格末尾回退）
  2. 在本地图纸库中查找对应PDF
  3. 从图纸PDF中提取REV/版本号字段
  4. 严格字符串比对，不匹配则高亮提醒
  5. 匹配的自动复制到待打印文件夹
"""
import os
import re
import glob
import shutil

import pdfplumber

from config import DRAWING_REV_PATTERN, DRAWING_PRINT_FOLDER


# ========== 版本号提取 ==========

def extract_version_from_name(product_name):
    """
    从品名规格末尾提取版本号。

    规则: 取最后一个空格后的 token，匹配 [A-Z][/.]?\d+ 模式
    例: "双剥镀 UL1007#24 黑色2+5 L=360mm B/01" → "B/01"
         "导线-A02版" → None（不匹配，因为含"版"字不是纯版本号）

    参数:
        product_name: str - 品名规格字符串

    返回:
        str | None - 版本号，或 None 如果未提取到
    """
    if not product_name or not product_name.strip():
        return None

    # 取末尾 token
    token = product_name.strip().split()[-1]

    # 匹配版本号模式: 字母开头 + 可选分隔符 + 数字
    if re.match(r"^[A-Z][/.]?\d+$", token):
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


# ========== 图纸文件查找 ==========

def find_drawing_file(drawing_dir, yy_code):
    """
    在图纸库目录中查找指定YY编号的图纸PDF。

    查找规则: 文件名以 yy_code 开头的 .pdf / .PDF 文件
    例: yy_code="YY60030058" → 匹配 "YY60030058导线-A02版.pdf"

    参数:
        drawing_dir: str - 图纸库目录路径
        yy_code: str - 客户料号 (YY开头)

    返回:
        str | None - 找到的PDF文件完整路径，或 None
    """
    if not drawing_dir or not os.path.isdir(drawing_dir):
        return None

    # glob匹配: YY编号开头的pdf文件
    pattern = os.path.join(drawing_dir, f"{yy_code}*.[pP][dD][fF]")
    matches = glob.glob(pattern)

    if matches:
        # 如有多个匹配，取第一个（按名称排序）
        return sorted(matches)[0]

    return None


# ========== 图纸PDF版本提取 ==========

def extract_version_from_pdf(pdf_path):
    """
    从图纸PDF中提取版本号（REV字段）。

    使用pdfplumber提取文本，通过正则搜索版本号/REV字段。

    参数:
        pdf_path: str - 图纸PDF文件路径

    返回:
        str | None - 版本号，或 None 如果无法提取
    """
    if not pdf_path or not os.path.exists(pdf_path):
        return None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                match = re.search(DRAWING_REV_PATTERN, text)
                if match:
                    return match.group(1).strip()
    except Exception:
        return None

    return None


# ========== 核心比对逻辑 ==========

def check_drawings(output_rows, drawing_dir, print_folder=None):
    """
    对订单中的YY产品执行图纸版本比对。

    规则:
    - 只比对YY开头的产品（F开头等跳过，标记"跳过"）
    - 版本来源优先级: 交期回复列 > 品名规格末尾
    - 严格字符串匹配（B/01 ≠ B01）
    - 每次比对前清空待打印文件夹
    - 匹配的自动复制到待打印文件夹

    参数:
        output_rows: list[dict] - apply_mapping 输出的行列表
        drawing_dir: str - 图纸库目录路径
        print_folder: str | None - 待打印文件夹路径（None则在drawing_dir下创建）

    返回:
        results: list[dict] - 每个项目的比对结果
          {yy_code, order_version, local_version, drawing_path,
           status, message}
          status: match / mismatch / no_version / no_drawing / skipped / no_pdf_version
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

    # 去重: 同一个YY编号只比对一次
    seen_codes = set()

    for row in output_rows:
        yy_code = row.get("产品规格", "").strip()

        # 跳过非YY产品
        if not yy_code.startswith("YY"):
            results.append({
                "yy_code": yy_code,
                "order_version": "",
                "local_version": "",
                "drawing_path": "",
                "status": "skipped",
                "message": "非YY产品，跳过",
            })
            continue

        # 去重
        if yy_code in seen_codes:
            continue
        seen_codes.add(yy_code)

        # 1. 获取「应有版本号」
        # 优先: 交期回复列
        order_version = extract_version_from_reply(row.get("_交期回复", ""))

        # 回退: 品名规格末尾
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
            })
            continue

        # 2. 查找本地图纸PDF
        drawing_path = find_drawing_file(drawing_dir, yy_code)
        if not drawing_path:
            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": "",
                "drawing_path": "",
                "status": "no_drawing",
                "message": "未找到图纸",
            })
            continue

        # 3. 提取图纸PDF中的版本号
        local_version = extract_version_from_pdf(drawing_path)
        if not local_version:
            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": "",
                "drawing_path": drawing_path,
                "status": "no_pdf_version",
                "message": "图纸无版本号",
            })
            continue

        # 4. 严格字符串比对
        if order_version == local_version:
            # 匹配 → 复制到待打印文件夹
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
            })
        else:
            results.append({
                "yy_code": yy_code,
                "order_version": order_version,
                "local_version": local_version,
                "drawing_path": drawing_path,
                "status": "mismatch",
                "message": f"本地: {local_version} → 最新: {order_version}",
            })

    return results


# ========== 统计 ==========

def get_check_stats(results):
    """
    统计比对结果。

    返回:
        dict - {total, match, mismatch, no_version, no_drawing, no_pdf_version, skipped}
    """
    stats = {
        "total": len(results),
        "match": 0,
        "mismatch": 0,
        "no_version": 0,
        "no_drawing": 0,
        "no_pdf_version": 0,
        "skipped": 0,
    }

    for r in results:
        status = r.get("status", "")
        if status in stats:
            stats[status] += 1

    return stats


# ========== 批量打印 ==========

def batch_print(print_folder):
    """
    批量打印待打印文件夹中的所有PDF。

    使用 os.startfile(path, "print") 调用系统默认PDF阅读器打印。
    仅在Windows下有效。

    参数:
        print_folder: str - 待打印文件夹路径

    返回:
        int - 发送打印的文件数量
    """
    if not os.path.isdir(print_folder):
        return 0

    count = 0
    for fname in sorted(os.listdir(print_folder)):
        if fname.lower().endswith(".pdf"):
            fpath = os.path.join(print_folder, fname)
            try:
                os.startfile(fpath, "print")
                count += 1
            except Exception:
                pass

    return count
