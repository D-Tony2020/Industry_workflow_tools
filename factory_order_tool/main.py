"""工厂订单PDF转Excel工具 - 主界面"""
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from version import VERSION, APP_NAME, BUILD_DATE
from config import MAPPING_TABLE_PATH, APP_DIR
from pdf_parser import parse_purchase_order
from code_mapper import load_mapping_table, apply_mapping, get_mapping_stats
from excel_writer import write_output_excel

# 预览表格显示的列（精简，便于查看）
PREVIEW_COLUMNS = [
    "产品编号",
    "产品名称",
    "产品规格",
    "数量",
    "计划结束时间",
    "工艺路线名称",
    "采购单价",
    "备注",
    "_映射状态",
]

PREVIEW_HEADERS = {
    "产品编号": "工厂编号",
    "产品名称": "产品名称",
    "产品规格": "客户料号",
    "数量": "数量",
    "计划结束时间": "出货日期",
    "工艺路线名称": "工艺路线",
    "采购单价": "单价",
    "备注": "备注",
    "_映射状态": "状态",
}


class OrderConverterApp:
    """主应用程序"""

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{VERSION}")
        self.root.geometry("1150x680")
        self.root.minsize(950, 550)

        # 数据
        self.pdf_path = tk.StringVar()
        self.header_info = {}
        self.output_rows = []
        self.mapping = {}
        self.status_text = tk.StringVar(value="就绪 - 请选择PDF文件")

        self._build_ui()
        self._load_mapping()

    def _build_ui(self):
        """构建界面"""
        # 顶部 - 文件选择区
        top_frame = ttk.LabelFrame(self.root, text="文件操作", padding=10)
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        ttk.Label(top_frame, text="PDF文件:").pack(side=tk.LEFT)
        ttk.Entry(top_frame, textvariable=self.pdf_path, width=65).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(top_frame, text="选择PDF", command=self._select_pdf).pack(side=tk.LEFT, padx=2)
        ttk.Button(top_frame, text="解析并映射", command=self._parse_pdf).pack(side=tk.LEFT, padx=2)

        # 工具栏
        tool_frame = ttk.Frame(self.root)
        tool_frame.pack(fill=tk.X, padx=10, pady=2)

        ttk.Button(tool_frame, text="打开映射表(Excel)", command=self._open_mapping_table).pack(side=tk.LEFT, padx=2)
        ttk.Button(tool_frame, text="重新加载映射表", command=self._load_mapping).pack(side=tk.LEFT, padx=2)

        self.mapping_label = ttk.Label(tool_frame, text="映射表: 未加载")
        self.mapping_label.pack(side=tk.LEFT, padx=10)

        ttk.Button(tool_frame, text="导出工厂Excel", command=self._export_excel).pack(side=tk.RIGHT, padx=2)
        ttk.Button(tool_frame, text="关于", command=self._show_about).pack(side=tk.RIGHT, padx=2)

        # 中间 - 数据预览表格
        table_frame = ttk.LabelFrame(self.root, text="数据预览", padding=5)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Treeview + 滚动条
        self.tree = ttk.Treeview(table_frame, columns=PREVIEW_COLUMNS, show="headings", height=20)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        col_widths = {
            "产品编号": 100, "产品名称": 250, "产品规格": 100, "数量": 80,
            "计划结束时间": 100, "工艺路线名称": 130, "采购单价": 70,
            "备注": 200, "_映射状态": 60,
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

        # 底部 - 状态栏
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Label(status_frame, textvariable=self.status_text).pack(side=tk.LEFT)
        ttk.Label(status_frame, text=f"v{VERSION}", foreground="gray").pack(side=tk.RIGHT)

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
            self.mapping_label.config(text=f"映射表: 未找到 ({MAPPING_TABLE_PATH})")
            self.status_text.set(f"映射表文件不存在，请将产品定义Excel放到: {MAPPING_TABLE_PATH}")
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
                f"请将产品定义Excel文件复制到该位置并命名为 mapping_table.xlsx",
            )
            return

        try:
            os.startfile(MAPPING_TABLE_PATH)
        except Exception:
            try:
                subprocess.Popen(["start", "", MAPPING_TABLE_PATH], shell=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开映射表:\n{e}\n\n文件位置: {MAPPING_TABLE_PATH}")

    def _show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于",
            f"{APP_NAME}\n\n"
            f"版本: v{VERSION}\n"
            f"构建日期: {BUILD_DATE}\n\n"
            f"功能: 客户采购单PDF → 工厂系统Excel\n"
            f"映射表: {MAPPING_TABLE_PATH}\n\n"
            f"如遇问题请联系开发人员并提供版本号",
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
