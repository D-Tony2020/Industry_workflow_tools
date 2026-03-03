"""工厂订单PDF转Excel工具 - 主界面 v1.2.0"""
import os
import sys
import json
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from version import VERSION, APP_NAME, BUILD_DATE
from config import MAPPING_TABLE_PATH, APP_DIR, DRAWING_PRINT_FOLDER
from pdf_parser import parse_purchase_order
from code_mapper import load_mapping_table, apply_mapping, get_mapping_stats
from excel_writer import write_output_excel
from drawing_checker import check_drawings, get_check_stats, batch_print

# 用户设置文件（与exe同目录）
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")


def _load_settings():
    """读取用户设置"""
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_settings(data):
    """保存用户设置"""
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# 预览表格显示的列
PREVIEW_COLUMNS = [
    "产品编号",
    "_产品名称",
    "产品规格",
    "数量",
    "计划开始时间",
    "计划结束时间",
    "工单分类",
    "_映射状态",
]

PREVIEW_HEADERS = {
    "产品编号": "工厂编号",
    "_产品名称": "产品名称",
    "产品规格": "客户料号",
    "数量": "数量",
    "计划开始时间": "开始日期",
    "计划结束时间": "交货日期",
    "工单分类": "工单分类",
    "_映射状态": "状态",
}

# 图纸比对结果表列
DRAWING_COLUMNS = [
    "yy_code",
    "order_version",
    "local_version",
    "status",
    "message",
]

DRAWING_HEADERS = {
    "yy_code": "客户料号",
    "order_version": "订单版本",
    "local_version": "本地版本",
    "status": "状态",
    "message": "说明",
}

# 状态中文显示
STATUS_LABELS = {
    "match": "匹配",
    "mismatch": "不匹配",
    "no_drawing": "无图纸",
    "no_version": "无版本",
    "bad_name": "命名不规范",
    "skipped": "跳过",
}

# 待处理图纸的状态分组配置（显示顺序、标题、操作说明）
_ACTIONABLE_GROUPS = [
    {
        "status": "mismatch",
        "title": "版本不匹配",
        "hint": "请下载最新版本替换",
    },
    {
        "status": "no_drawing",
        "title": "缺失图纸",
        "hint": "请下载并保存到图纸库",
    },
    {
        "status": "bad_name",
        "title": "命名不规范",
        "hint": "请按以下文件名重命名",
    },
]


class OrderConverterApp:
    """主应用程序"""

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{VERSION}")
        self.root.geometry("1150x900")
        self.root.minsize(950, 700)

        # 数据
        self.pdf_path = tk.StringVar()
        self.drawing_dir = tk.StringVar()
        self.header_info = {}
        self.output_rows = []
        self.mapping = {}
        self.drawing_results = []
        self.status_text = tk.StringVar(value="就绪 - 请选择PDF文件")

        # 加载用户设置（图纸库路径等）
        settings = _load_settings()
        saved_dir = settings.get("drawing_dir", "")
        if saved_dir and os.path.isdir(saved_dir):
            self.drawing_dir.set(saved_dir)

        self._build_ui()
        self._load_mapping()

    def _build_ui(self):
        """构建界面"""
        # ===== 顶部 - 文件选择区 =====
        top_frame = ttk.LabelFrame(self.root, text="文件操作", padding=10)
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        ttk.Label(top_frame, text="PDF文件:").pack(side=tk.LEFT)
        ttk.Entry(top_frame, textvariable=self.pdf_path, width=65).pack(
            side=tk.LEFT, padx=5, fill=tk.X, expand=True
        )
        ttk.Button(top_frame, text="选择PDF", command=self._select_pdf).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(top_frame, text="解析并映射", command=self._parse_pdf).pack(
            side=tk.LEFT, padx=2
        )

        # ===== 工具栏 =====
        tool_frame = ttk.Frame(self.root)
        tool_frame.pack(fill=tk.X, padx=10, pady=2)

        self.open_mapping_btn = tk.Button(
            tool_frame, text="打开映射表(Excel)", command=self._open_mapping_table
        )
        self.open_mapping_btn.pack(side=tk.LEFT, padx=2)
        self.reload_mapping_btn = tk.Button(
            tool_frame,
            text="重新加载映射表",
            command=lambda: self._load_mapping(auto_reprocess=True),
        )
        self.reload_mapping_btn.pack(side=tk.LEFT, padx=2)

        self.mapping_label = ttk.Label(tool_frame, text="映射表: 未加载")
        self.mapping_label.pack(side=tk.LEFT, padx=10)

        ttk.Button(
            tool_frame, text="导出工厂Excel", command=self._export_excel
        ).pack(side=tk.RIGHT, padx=2)
        ttk.Button(tool_frame, text="关于", command=self._show_about).pack(
            side=tk.RIGHT, padx=2
        )

        # ===== 中间 - 数据预览表格 =====
        table_frame = ttk.LabelFrame(self.root, text="数据预览", padding=5)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Treeview + 滚动条（序号列 + 数据列）
        preview_tree_cols = ["_seq"] + list(PREVIEW_COLUMNS)
        self.tree = ttk.Treeview(
            table_frame, columns=preview_tree_cols, show="headings", height=12
        )

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # 序号列
        self.tree.heading("_seq", text="序号")
        self.tree.column("_seq", width=45, minwidth=35, anchor=tk.CENTER)

        col_widths = {
            "产品编号": 110,
            "_产品名称": 250,
            "产品规格": 110,
            "数量": 60,
            "计划开始时间": 90,
            "计划结束时间": 90,
            "工单分类": 80,
            "_映射状态": 60,
        }
        for col in PREVIEW_COLUMNS:
            self.tree.heading(col, text=PREVIEW_HEADERS.get(col, col))
            self.tree.column(col, width=col_widths.get(col, 80), minwidth=40)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 配置tag颜色
        self.tree.tag_configure("unmapped", background="#FFC7CE")
        self.tree.tag_configure("mapped", background="#C6EFCE")

        # ===== 图纸比对区 =====
        drawing_frame = ttk.LabelFrame(self.root, text="图纸版本比对", padding=10)
        drawing_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 图纸库路径选择
        dir_row = ttk.Frame(drawing_frame)
        dir_row.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(dir_row, text="图纸库:").pack(side=tk.LEFT)
        ttk.Entry(dir_row, textvariable=self.drawing_dir, width=55).pack(
            side=tk.LEFT, padx=5, fill=tk.X, expand=True
        )
        ttk.Button(
            dir_row, text="选择图纸库", command=self._select_drawing_dir
        ).pack(side=tk.LEFT, padx=2)

        # 按钮行
        btn_row = ttk.Frame(drawing_frame)
        btn_row.pack(fill=tk.X, pady=(0, 5))

        self.check_btn = tk.Button(
            btn_row, text="图纸比对", command=self._check_drawings
        )
        self.check_btn.pack(side=tk.LEFT, padx=2)

        self.print_all_btn = ttk.Button(
            btn_row,
            text="一键全部打印",
            command=self._batch_print,
            state=tk.DISABLED,
        )
        self.print_all_btn.pack(side=tk.LEFT, padx=2)

        ttk.Button(
            btn_row,
            text="打开待打印文件夹",
            command=self._open_print_folder,
        ).pack(side=tk.LEFT, padx=2)

        # 待处理图纸命名按钮（初始隐藏，比对后按需显示）
        self.naming_btn = ttk.Button(
            btn_row,
            text="查看待处理图纸",
            command=self._show_naming_helper,
        )
        # 初始不pack，比对后有待处理项时才显示

        self.drawing_stats_label = ttk.Label(btn_row, text="")
        self.drawing_stats_label.pack(side=tk.LEFT, padx=10)

        # 比对结果 Treeview（序号列 + 数据列）
        tree_container = ttk.Frame(drawing_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        drawing_tree_cols = ["_seq"] + list(DRAWING_COLUMNS)
        drawing_col_widths = {
            "_seq": 45,
            "yy_code": 120,
            "order_version": 80,
            "local_version": 80,
            "status": 80,
            "message": 350,
        }

        self.drawing_tree = ttk.Treeview(
            tree_container,
            columns=drawing_tree_cols,
            show="headings",
            height=8,
        )
        dvsb = ttk.Scrollbar(
            tree_container, orient="vertical", command=self.drawing_tree.yview
        )
        dhsb = ttk.Scrollbar(
            tree_container, orient="horizontal", command=self.drawing_tree.xview
        )
        self.drawing_tree.configure(
            yscrollcommand=dvsb.set, xscrollcommand=dhsb.set
        )

        # 序号列
        self.drawing_tree.heading("_seq", text="序号")
        self.drawing_tree.column("_seq", width=45, minwidth=35, anchor=tk.CENTER)

        for col in DRAWING_COLUMNS:
            self.drawing_tree.heading(col, text=DRAWING_HEADERS.get(col, col))
            self.drawing_tree.column(
                col, width=drawing_col_widths.get(col, 80), minwidth=40
            )

        self.drawing_tree.grid(row=0, column=0, sticky="nsew")
        dvsb.grid(row=0, column=1, sticky="ns")
        dhsb.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # 配置图纸比对tag颜色
        self.drawing_tree.tag_configure("match", background="#C6EFCE")
        self.drawing_tree.tag_configure("mismatch", background="#FFC7CE")
        self.drawing_tree.tag_configure("no_version", background="#FFFFCC")
        self.drawing_tree.tag_configure("no_drawing", background="#D9D9D9")
        self.drawing_tree.tag_configure("bad_name", background="#FFE0B2")
        self.drawing_tree.tag_configure("skipped", background="#E8E8E8")

        # ===== 底部 - 状态栏 =====
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Label(status_frame, textvariable=self.status_text).pack(side=tk.LEFT)
        ttk.Label(status_frame, text=f"v{VERSION}", foreground="gray").pack(
            side=tk.RIGHT
        )

    # ========== 按钮高亮 ==========

    @staticmethod
    def _highlight_btn(btn, text=None):
        """将 tk.Button 设为醒目状态（红底白字 + ⚠ 前缀）"""
        if text:
            btn.config(text=f"⚠ {text}")
        btn.config(bg="#E74C3C", fg="white", activebackground="#C0392B",
                   activeforeground="white")

    @staticmethod
    def _unhighlight_btn(btn, text=None):
        """将 tk.Button 恢复为默认状态"""
        if text:
            btn.config(text=text)
        btn.config(bg="SystemButtonFace", fg="SystemButtonText",
                   activebackground="SystemButtonFace",
                   activeforeground="SystemButtonText")

    # ========== 订单转换功能 ==========

    def _select_pdf(self):
        """选择PDF文件"""
        path = filedialog.askopenfilename(
            title="选择采购单PDF",
            filetypes=[("PDF文件", "*.pdf *.PDF"), ("所有文件", "*.*")],
        )
        if path:
            self.pdf_path.set(path)

    def _parse_pdf(self):
        """解析PDF并应用映射"""
        path = self.pdf_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先选择有效的PDF文件")
            return

        self.status_text.set("正在解析PDF...")
        self.root.update()

        try:
            self.header_info, items = parse_purchase_order(path)
        except Exception as e:
            messagebox.showerror("解析错误", f"PDF解析失败:\n{e}")
            self.status_text.set("解析失败")
            return

        if not items:
            messagebox.showwarning("提示", "未从PDF中解析到任何订单数据")
            self.status_text.set("未解析到数据")
            return

        # 应用映射
        self.output_rows, unmapped = apply_mapping(items, self.mapping)
        total, mapped, failed = get_mapping_stats(self.output_rows)

        # 刷新表格
        self._refresh_table()

        self.status_text.set(
            f"解析完成: 共{total}条 | 映射成功{mapped}条 | 未映射{failed}条"
        )

        if unmapped:
            self._highlight_btn(self.open_mapping_btn, "打开映射表(Excel)")
            unique_unmapped = sorted(set(unmapped))
            unmapped_str = "\n".join(unique_unmapped)
            messagebox.showinfo(
                "映射提醒",
                f"以下{len(unique_unmapped)}个料件编号未找到映射:\n\n{unmapped_str}\n\n"
                f"请在映射表中添加后，点击「重新加载映射表」再重新解析",
            )
        else:
            self._unhighlight_btn(self.open_mapping_btn, "打开映射表(Excel)")

    def _refresh_table(self):
        """刷新预览表格"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        for idx, row_data in enumerate(self.output_rows, start=1):
            values = [idx] + [row_data.get(col, "") for col in PREVIEW_COLUMNS]
            tag = "mapped" if row_data.get("_映射状态") == "已映射" else "unmapped"
            self.tree.insert("", tk.END, values=values, tags=(tag,))

    def _load_mapping(self, auto_reprocess=False):
        """加载映射表

        参数:
            auto_reprocess: bool - 加载成功后是否自动重新解析并比对
                            用户点击"重新加载映射表"时为True，初始化时为False
        """
        if not os.path.exists(MAPPING_TABLE_PATH):
            self.mapping = {}
            self.mapping_label.config(
                text=f"映射表: 未找到 ({MAPPING_TABLE_PATH})"
            )
            self.status_text.set(
                f"映射表文件不存在，请将料号清单Excel放到: {MAPPING_TABLE_PATH}"
            )
            return

        try:
            self.mapping = load_mapping_table()
            count = len(self.mapping)
            self.mapping_label.config(text=f"映射表: 已加载 {count} 条")
            self.status_text.set(f"映射表已加载: {count}条映射规则")
            # 加载成功，取消"重新加载"高亮
            self._unhighlight_btn(self.reload_mapping_btn, "重新加载映射表")
        except Exception as e:
            self.mapping = {}
            self.mapping_label.config(text="映射表: 加载失败")
            messagebox.showerror("错误", f"映射表加载失败:\n{e}")
            return

        # 自动重新处理链: 映射表重载 → 重新解析PDF → 重新比对图纸
        if auto_reprocess and self.pdf_path.get().strip():
            self._parse_pdf()
            # 解析成功且图纸库已配置时，自动触发比对
            if self.output_rows and self.drawing_dir.get().strip():
                self._check_drawings()

    def _open_mapping_table(self):
        """用系统默认程序打开映射表"""
        if not os.path.exists(MAPPING_TABLE_PATH):
            messagebox.showwarning(
                "提示",
                f"映射表文件不存在:\n{MAPPING_TABLE_PATH}\n\n"
                f"请将料号清单Excel文件复制到该位置并命名为 mapping_table.xlsx",
            )
            return

        try:
            os.startfile(MAPPING_TABLE_PATH)
        except Exception:
            try:
                subprocess.Popen(["start", "", MAPPING_TABLE_PATH], shell=True)
            except Exception as e:
                messagebox.showerror(
                    "错误",
                    f"无法打开映射表:\n{e}\n\n文件位置: {MAPPING_TABLE_PATH}",
                )
                return

        # 打开后：取消自身高亮，引导用户点"重新加载映射表"
        self._unhighlight_btn(self.open_mapping_btn, "打开映射表(Excel)")
        self._highlight_btn(self.reload_mapping_btn, "重新加载映射表")

    def _export_excel(self):
        """导出为工厂系统Excel"""
        if not self.output_rows:
            messagebox.showwarning("提示", "没有数据可导出，请先解析PDF")
            return

        # 默认文件名
        order_no = self.header_info.get("采购单号", "订单")
        default_name = f"工厂订单_{order_no}.xlsx"

        path = filedialog.asksaveasfilename(
            title="保存工厂系统Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel文件", "*.xlsx")],
        )
        if not path:
            return

        try:
            write_output_excel(self.output_rows, path)
            self.status_text.set(f"导出成功: {path}")
            messagebox.showinfo("成功", f"Excel已导出:\n{path}")
        except Exception as e:
            messagebox.showerror("导出错误", f"导出失败:\n{e}")

    # ========== 图纸比对功能 ==========

    def _select_drawing_dir(self):
        """选择图纸库目录并持久化保存"""
        path = filedialog.askdirectory(title="选择图纸库文件夹")
        if path:
            self.drawing_dir.set(path)
            # 持久化保存
            settings = _load_settings()
            settings["drawing_dir"] = path
            _save_settings(settings)

    def _check_drawings(self):
        """执行图纸版本比对"""
        if not self.output_rows:
            messagebox.showwarning("提示", "请先解析PDF订单数据")
            return

        drawing_dir = self.drawing_dir.get().strip()
        if not drawing_dir or not os.path.isdir(drawing_dir):
            messagebox.showwarning("提示", "请先选择有效的图纸库文件夹")
            return

        self.status_text.set("正在比对图纸版本...")
        self.root.update()

        # 待打印文件夹在图纸库下
        print_folder = os.path.join(drawing_dir, DRAWING_PRINT_FOLDER)

        try:
            self.drawing_results, bad_names = check_drawings(
                self.output_rows, drawing_dir, print_folder
            )
        except Exception as e:
            messagebox.showerror("比对错误", f"图纸比对失败:\n{e}")
            self.status_text.set("图纸比对失败")
            return

        # 刷新图纸比对表格
        self._refresh_drawing_table()

        # 按钮文案切换为工作流引导
        self.check_btn.config(text="我已完成最新图纸文件下载")

        # 统计
        stats = get_check_stats(self.drawing_results)

        stat_parts = [f"{stats['match']}匹配"]
        if stats["mismatch"]:
            stat_parts.append(f"{stats['mismatch']}不匹配")
        if stats["no_version"]:
            stat_parts.append(f"{stats['no_version']}无版本")
        if stats["no_drawing"]:
            stat_parts.append(f"{stats['no_drawing']}无图纸")
        if stats.get("bad_name"):
            stat_parts.append(f"{stats['bad_name']}命名不规范")

        self.drawing_stats_label.config(
            text="比对: " + " | ".join(stat_parts)
        )

        # 显示/隐藏"查看待处理图纸"按钮（任一非match状态均显示）
        actionable_count = (
            stats["mismatch"] + stats["no_drawing"] + stats.get("bad_name", 0)
        )
        if actionable_count > 0:
            self.naming_btn.pack(side=tk.LEFT, padx=2)
        else:
            self.naming_btn.pack_forget()

        # 判断是否可以启用一键打印：仅当所有YY产品均为 match 或 skipped 时
        non_match = (
            stats["mismatch"]
            + stats["no_drawing"]
            + stats["no_version"]
            + stats.get("bad_name", 0)
        )
        if non_match == 0 and stats["match"] > 0:
            self.print_all_btn.config(state=tk.NORMAL)
            self._unhighlight_btn(self.check_btn, "我已完成最新图纸文件下载")
            self.status_text.set(
                f"图纸比对完成: 全部匹配！共{stats['match']}个图纸已复制到待打印文件夹"
            )
        else:
            self.print_all_btn.config(state=tk.DISABLED)
            # 有待处理项 → 高亮"已完成下载"按钮引导用户操作
            if actionable_count > 0:
                self._highlight_btn(self.check_btn, "我已完成最新图纸文件下载")
            if stats["mismatch"] > 0:
                self.status_text.set(
                    f"图纸比对完成: {stats['mismatch']}个版本不匹配，请更新后重新比对"
                )
            else:
                self.status_text.set("图纸比对完成")

        # 检查是否有未映射物料 → 高亮"打开映射表"
        code_map = {
            row.get("产品规格", "").strip(): row.get("产品编号", "").strip()
            for row in self.output_rows
        }
        has_unmapped = any(
            not code_map.get(r["yy_code"])
            for r in self.drawing_results
            if r.get("status") in ("mismatch", "no_drawing", "bad_name")
        )
        if has_unmapped:
            self._highlight_btn(self.open_mapping_btn, "打开映射表(Excel)")

        # 命名不规范提示
        if bad_names:
            names_str = "\n".join(bad_names[:20])
            suffix = f"\n...等共{len(bad_names)}个文件" if len(bad_names) > 20 else ""
            messagebox.showwarning(
                "命名提醒",
                f"以下图纸文件不符合标准命名规则，无法自动提取版本号：\n\n"
                f"{names_str}{suffix}\n\n"
                f"标准格式: 工厂编号 客户料号-版本号.pdf\n"
                f"示例: J00016025 YY60030362-A01.pdf\n\n"
                f"可选: 末尾可追加产品类型（导线或其他）\n"
                f"示例: J00016025 YY60030362-A01导线.pdf",
            )

        # 比对完成后自动弹出待处理图纸弹窗
        if actionable_count > 0:
            self._show_naming_helper()

    def _refresh_drawing_table(self):
        """刷新图纸比对结果表格"""
        for row in self.drawing_tree.get_children():
            self.drawing_tree.delete(row)

        for idx, result in enumerate(self.drawing_results, start=1):
            values = [idx] + [result.get(col, "") for col in DRAWING_COLUMNS]
            status = result.get("status", "")
            # 将status列替换为中文（序号偏移+1）
            status_idx = DRAWING_COLUMNS.index("status") + 1  # +1 因为序号列
            values[status_idx] = STATUS_LABELS.get(status, status)
            tag = status if status else "skipped"
            self.drawing_tree.insert("", tk.END, values=values, tags=(tag,))

    def _show_naming_helper(self):
        """显示待处理图纸规范命名助手弹窗（单表 + 状态分色 + 映射补全）"""
        # 筛选所有待处理的结果（mismatch / no_drawing / bad_name）
        actionable_statuses = {"mismatch", "no_drawing", "bad_name"}
        status_order = {"mismatch": 0, "no_drawing": 1, "bad_name": 2}
        actionable = sorted(
            [
                r for r in self.drawing_results
                if r.get("status") in actionable_statuses and r.get("suggested_name")
            ],
            key=lambda r: status_order.get(r.get("status"), 99),
        )

        if not actionable:
            messagebox.showinfo("提示", "没有需要处理的图纸")
            return

        # 构建工厂编号映射（从映射表补全）
        code_map = {}
        for row in self.output_rows:
            yy = row.get("产品规格", "").strip()
            jy = row.get("产品编号", "").strip()
            if yy and jy:
                code_map[yy] = jy

        # 检查未映射的物料
        unmapped_yy = [
            r["yy_code"] for r in actionable if r["yy_code"] not in code_map
        ]

        # 处理方式映射
        action_labels = {
            "mismatch": "下载最新版",
            "no_drawing": "下载图纸",
            "bad_name": "重命名",
        }

        # 统计各组数量（用于标题）
        group_counts = {}
        for r in actionable:
            s = r.get("status")
            group_counts[s] = group_counts.get(s, 0) + 1

        group_summary = []
        for g in _ACTIONABLE_GROUPS:
            cnt = group_counts.get(g["status"], 0)
            if cnt:
                group_summary.append(f"{cnt}个{g['title']}")

        # ===== 创建弹窗 =====
        win = tk.Toplevel(self.root)
        win.title("待处理图纸 — 规范命名参考")
        win.geometry("820x480")
        win.minsize(700, 350)
        win.transient(self.root)
        win.grab_set()

        # 顶部说明
        ttk.Label(
            win,
            text=f"以下图纸需要处理（{' | '.join(group_summary)}），"
            f"请按规范命名下载或重命名后保存到图纸库：",
            wraplength=780,
        ).pack(padx=15, pady=(15, 2), anchor=tk.W)

        tk.Label(
            win,
            text="命名格式: 工厂编号 客户料号-版本号.pdf  （可选末尾追加产品类型）",
            fg="#CC0000",
            font=("", 9, "bold"),
        ).pack(padx=15, pady=(0, 5), anchor=tk.W)

        # 未映射物料警告
        if unmapped_yy:
            warn_frame = tk.Frame(win, bg="#FFF3CD")
            warn_frame.pack(fill=tk.X, padx=15, pady=(0, 5))
            tk.Label(
                warn_frame,
                text=f"⚠ 以下 {len(unmapped_yy)} 个物料未在映射表中找到工厂编号，"
                f"请先更新映射表：{', '.join(unmapped_yy)}",
                bg="#FFF3CD",
                fg="#856404",
                wraplength=770,
                justify=tk.LEFT,
            ).pack(padx=10, pady=6, anchor=tk.W)

        # ===== 表格区 =====
        table_frame = ttk.Frame(win)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 5))

        cols = ("seq", "action", "factory_code", "yy_code", "version", "filename")
        headers = {
            "seq": "序号",
            "action": "处理方式",
            "factory_code": "工厂编号",
            "yy_code": "客户料号",
            "version": "版本号",
            "filename": "规范文件名",
        }
        col_widths = {
            "seq": 40,
            "action": 80,
            "factory_code": 95,
            "yy_code": 120,
            "version": 55,
            "filename": 380,
        }
        col_anchors = {
            "seq": tk.CENTER,
            "action": tk.CENTER,
        }

        tree = ttk.Treeview(
            table_frame, columns=cols, show="headings",
            height=min(len(actionable) + 1, 15),
        )
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in cols:
            tree.heading(col, text=headers[col])
            tree.column(
                col,
                width=col_widths.get(col, 80),
                minwidth=30,
                anchor=col_anchors.get(col, tk.W),
            )

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 配置行颜色（与主表一致）
        tree.tag_configure("mismatch", background="#FFC7CE")
        tree.tag_configure("no_drawing", background="#D9D9D9")
        tree.tag_configure("bad_name", background="#FFE0B2")

        # 填充数据
        all_names = []
        for idx, r in enumerate(actionable, start=1):
            status = r.get("status", "")
            yy_code = r["yy_code"]
            version = r.get("order_version", "")
            factory_code = code_map.get(yy_code, "???")
            name = r["suggested_name"]
            action = action_labels.get(status, "")
            all_names.append(name)
            tree.insert(
                "", tk.END,
                values=(idx, action, factory_code, yy_code, version, name),
                tags=(status,),
            )

        # 双击复制规范文件名
        def on_double_click(event):
            sel = tree.selection()
            if sel:
                vals = tree.item(sel[0])["values"]
                name = vals[5]  # filename 在第6列（index 5）
                win.clipboard_clear()
                win.clipboard_append(str(name))
                self.status_text.set(f"已复制: {name}")

        tree.bind("<Double-1>", on_double_click)

        # ===== 底部按钮 =====
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill=tk.X, padx=15, pady=(5, 15))

        def copy_all():
            text = "\n".join(all_names)
            win.clipboard_clear()
            win.clipboard_append(text)
            messagebox.showinfo(
                "已复制",
                f"已将 {len(all_names)} 个规范文件名复制到剪贴板",
                parent=win,
            )

        tk.Label(
            btn_frame,
            text="提示: 双击任意行可复制该行规范文件名",
            fg="#CC0000",
            font=("", 9, "bold"),
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(btn_frame, text="一键复制全部", command=copy_all).pack(
            side=tk.RIGHT, padx=2
        )
        ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(
            side=tk.RIGHT, padx=2
        )

    def _batch_print(self):
        """一键全部打印"""
        drawing_dir = self.drawing_dir.get().strip()
        if not drawing_dir:
            return

        print_folder = os.path.join(drawing_dir, DRAWING_PRINT_FOLDER)
        if not os.path.isdir(print_folder):
            messagebox.showwarning("提示", "待打印文件夹不存在")
            return

        count = batch_print(print_folder)
        if count > 0:
            self.status_text.set(f"已发送 {count} 个文件到打印机")
            messagebox.showinfo("打印", f"已将 {count} 个PDF文件发送到默认打印机")
        else:
            messagebox.showwarning("提示", "待打印文件夹中没有PDF文件")

    def _open_print_folder(self):
        """打开待打印文件夹"""
        drawing_dir = self.drawing_dir.get().strip()
        if not drawing_dir:
            messagebox.showwarning("提示", "请先选择图纸库文件夹")
            return

        print_folder = os.path.join(drawing_dir, DRAWING_PRINT_FOLDER)
        os.makedirs(print_folder, exist_ok=True)

        try:
            os.startfile(print_folder)
        except Exception:
            try:
                subprocess.Popen(["start", "", print_folder], shell=True)
            except Exception as e:
                messagebox.showerror(
                    "错误", f"无法打开文件夹:\n{e}\n\n路径: {print_folder}"
                )

    # ========== 通用 ==========

    def _show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于",
            f"{APP_NAME}\n\n"
            f"版本: v{VERSION}\n"
            f"构建日期: {BUILD_DATE}\n\n"
            f"功能:\n"
            f"  1. 客户采购单PDF → 工厂系统Excel\n"
            f"  2. 图纸版本比对 + 一键打印\n\n"
            f"图纸命名规则: 工厂编号 客户料号-版本号.pdf\n"
            f"映射表: {MAPPING_TABLE_PATH}\n\n"
            f"如遇问题请联系开发人员并提供版本号",
        )


def main():
    root = tk.Tk()

    # Windows高分屏DPI感知
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    # 使用系统主题
    style = ttk.Style()
    available_themes = style.theme_names()
    if "vista" in available_themes:
        style.theme_use("vista")
    elif "clam" in available_themes:
        style.theme_use("clam")

    app = OrderConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
