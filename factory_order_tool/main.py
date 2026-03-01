"""工厂订单PDF转Excel工具 - 主界面 v1.1.0"""
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from version import VERSION, APP_NAME, BUILD_DATE
from config import MAPPING_TABLE_PATH, APP_DIR, DRAWING_PRINT_FOLDER
from pdf_parser import parse_purchase_order
from code_mapper import load_mapping_table, apply_mapping, get_mapping_stats
from excel_writer import write_output_excel
from drawing_checker import check_drawings, get_check_stats, batch_print

# 预览表格显示的列（v1.1.0: 精简，只展示有值的字段）
PREVIEW_COLUMNS = [
    "产品编号",
    "_产品名称",
    "产品规格",
    "数量",
    "计划开始时间",
    "工单分类",
    "_映射状态",
]

PREVIEW_HEADERS = {
    "产品编号": "工厂编号",
    "_产品名称": "产品名称",
    "产品规格": "客户料号",
    "数量": "数量",
    "计划开始时间": "出货日期",
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

        ttk.Button(
            tool_frame, text="打开映射表(Excel)", command=self._open_mapping_table
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            tool_frame, text="重新加载映射表", command=self._load_mapping
        ).pack(side=tk.LEFT, padx=2)

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

        # Treeview + 滚动条
        self.tree = ttk.Treeview(
            table_frame, columns=PREVIEW_COLUMNS, show="headings", height=12
        )

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        col_widths = {
            "产品编号": 110,
            "_产品名称": 280,
            "产品规格": 110,
            "数量": 60,
            "计划开始时间": 100,
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

        ttk.Button(
            btn_row, text="图纸比对", command=self._check_drawings
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_row,
            text="我已完成最新图纸文件下载",
            command=self._check_drawings,
        ).pack(side=tk.LEFT, padx=2)

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

        self.drawing_stats_label = ttk.Label(btn_row, text="")
        self.drawing_stats_label.pack(side=tk.LEFT, padx=10)

        # 比对结果 Treeview（放在子frame中，避免与pack混用grid）
        tree_container = ttk.Frame(drawing_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        drawing_col_widths = {
            "yy_code": 120,
            "order_version": 80,
            "local_version": 80,
            "status": 80,
            "message": 350,
        }

        self.drawing_tree = ttk.Treeview(
            tree_container,
            columns=DRAWING_COLUMNS,
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
        self.drawing_tree.tag_configure("no_pdf_version", background="#FFE0B2")
        self.drawing_tree.tag_configure("skipped", background="#E8E8E8")

        # ===== 底部 - 状态栏 =====
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Label(status_frame, textvariable=self.status_text).pack(side=tk.LEFT)
        ttk.Label(status_frame, text=f"v{VERSION}", foreground="gray").pack(
            side=tk.RIGHT
        )

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
            unique_unmapped = sorted(set(unmapped))
            unmapped_str = "\n".join(unique_unmapped)
            messagebox.showinfo(
                "映射提醒",
                f"以下{len(unique_unmapped)}个料件编号未找到映射:\n\n{unmapped_str}\n\n"
                f"请在映射表中添加后，点击「重新加载映射表」再重新解析",
            )

    def _refresh_table(self):
        """刷新预览表格"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        for row_data in self.output_rows:
            values = [row_data.get(col, "") for col in PREVIEW_COLUMNS]
            tag = "mapped" if row_data.get("_映射状态") == "已映射" else "unmapped"
            self.tree.insert("", tk.END, values=values, tags=(tag,))

    def _load_mapping(self):
        """加载映射表"""
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
        except Exception as e:
            self.mapping = {}
            self.mapping_label.config(text="映射表: 加载失败")
            messagebox.showerror("错误", f"映射表加载失败:\n{e}")

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
        """选择图纸库目录"""
        path = filedialog.askdirectory(title="选择图纸库文件夹")
        if path:
            self.drawing_dir.set(path)

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
            self.drawing_results = check_drawings(
                self.output_rows, drawing_dir, print_folder
            )
        except Exception as e:
            messagebox.showerror("比对错误", f"图纸比对失败:\n{e}")
            self.status_text.set("图纸比对失败")
            return

        # 刷新图纸比对表格
        self._refresh_drawing_table()

        # 统计
        stats = get_check_stats(self.drawing_results)
        self.drawing_stats_label.config(
            text=(
                f"比对: {stats['match']}匹配 | "
                f"{stats['mismatch']}不匹配 | "
                f"{stats['no_version']}无版本 | "
                f"{stats['no_drawing']}无图纸"
            )
        )

        # 判断是否可以启用一键打印
        # 条件: 所有有版本号的项全部匹配（无mismatch）
        if stats["mismatch"] == 0 and stats["match"] > 0:
            self.print_all_btn.config(state=tk.NORMAL)
            self.status_text.set(
                f"图纸比对完成: 全部匹配！共{stats['match']}个图纸已复制到待打印文件夹"
            )
        else:
            self.print_all_btn.config(state=tk.DISABLED)
            if stats["mismatch"] > 0:
                self.status_text.set(
                    f"图纸比对完成: {stats['mismatch']}个版本不匹配，请更新后重新比对"
                )
            else:
                self.status_text.set("图纸比对完成")

    def _refresh_drawing_table(self):
        """刷新图纸比对结果表格"""
        for row in self.drawing_tree.get_children():
            self.drawing_tree.delete(row)

        for result in self.drawing_results:
            values = [result.get(col, "") for col in DRAWING_COLUMNS]
            status = result.get("status", "")
            tag = status if status else "skipped"
            self.drawing_tree.insert("", tk.END, values=values, tags=(tag,))

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
