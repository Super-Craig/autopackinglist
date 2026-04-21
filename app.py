import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# =========================
# 📁 基础路径
# =========================
BASE_OUTPUT_DIR = "output"

PART_COL = "料号/Part Number"
DESC_COL = "名称/Description"
QTY_COL = "数量/Quantity"
NO_COL = "NO."

INFO_FILE = "info.xlsx"
TEMPLATE_FILE = "template.xlsx"

os.makedirs(BASE_OUTPUT_DIR, exist_ok=True)


# =========================
# 🧼 清洗
# =========================
def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


# =========================
# 📥 读取数据
# =========================
def load_info():
    excel = pd.ExcelFile(INFO_FILE)
    dfs = []

    for sheet in excel.sheet_names:
        df = excel.parse(sheet, dtype=str)

        if PART_COL not in df.columns:
            continue

        for c in df.columns:
            df[c] = df[c].apply(clean)

        dfs.append(df)

    return dfs


# =========================
# 🔍 查找
# =========================
def search(part, dfs):
    part = clean(part)

    for df in dfs:
        df[PART_COL] = df[PART_COL].astype(str).apply(clean)

        res = df[df[PART_COL] == part]

        if not res.empty:
            r = res.iloc[0]
            return {
                "part": clean(r[PART_COL]),
                "desc": clean(r.get(DESC_COL, "")),
                "qty": clean(r.get(QTY_COL, ""))
            }

    return None


# =========================
# 📊 template列
# =========================
def get_cols(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=4, column=c).value
        if v:
            cols[str(v).strip()] = c
    return cols


def next_row(ws, pn_col):
    r = 5
    while ws.cell(row=r, column=pn_col).value:
        r += 1
    return r


# =========================
# ✍️ 写入
# =========================
def write_to_excel(item):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    cols = get_cols(ws)

    row = next_row(ws, cols[PART_COL])

    # 👉 定义居中样式（关键）
    center_align = Alignment(horizontal="center", vertical="center")

    # 写入数据
    ws.cell(row=row, column=cols[NO_COL], value=row - 4)
    ws.cell(row=row, column=cols[PART_COL], value=item["part"])
    ws.cell(row=row, column=cols[DESC_COL], value=item["desc"])
    ws.cell(row=row, column=cols[QTY_COL], value=item["qty"])

    # 👉 设置居中（核心）
    ws.cell(row=row, column=cols[NO_COL]).alignment = center_align
    ws.cell(row=row, column=cols[PART_COL]).alignment = center_align
    ws.cell(row=row, column=cols[DESC_COL]).alignment = center_align
    ws.cell(row=row, column=cols[QTY_COL]).alignment = center_align

    wb.save(TEMPLATE_FILE)


# =========================
# 🧹 清空
# =========================
def clear_template():
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    for r in ws.iter_rows(min_row=5):
        for c in r:
            c.value = None

    wb.save(TEMPLATE_FILE)


# =========================
# 📦 导出
# =========================
def export_file(output_dir, name):
    if not name:
        files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx")]
        name = str(len(files) + 1)

    path = os.path.join(output_dir, f"{name}.xlsx")

    wb = load_workbook(TEMPLATE_FILE)
    wb.save(path)

    clear_template()

    return path


# =========================
# 🖨 打印
# =========================
def print_file(filepath):
    try:
        os.startfile(filepath, "print")
    except Exception as e:
        messagebox.showerror("打印失败", str(e))


# =========================
# 🖥 GUI
# =========================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 装箱系统")
        self.root.geometry("750x520")

        # 👉 启动时输入订单号
        order_no = simpledialog.askstring("订单号", "请输入订单号（可空）")
        self.root.deiconify()              # 如果被最小化，恢复
        self.root.lift()                   # 提到最前
        self.root.attributes('-topmost', True)
        self.root.after(100, lambda: self.root.attributes('-topmost', False))
        self.root.focus_force()

        if order_no:
            self.output_dir = os.path.join(BASE_OUTPUT_DIR, order_no)
        else:
            self.output_dir = BASE_OUTPUT_DIR

        os.makedirs(self.output_dir, exist_ok=True)

        self.data = load_info()
        self.items = []

        # ===== 输入 =====
        self.input = tk.Entry(root, font=("Arial", 16))
        self.input.pack(fill=tk.X, padx=10, pady=10)
        self.input.bind("<Return>", self.scan)
        self.input.focus_set()

        self.status = tk.Label(root, text=f"当前目录：{self.output_dir}", fg="blue")
        self.status.pack()

        # ===== 表格 =====
        self.tree = ttk.Treeview(root, columns=("no", "part", "desc", "qty"), show="headings")
        self.tree.heading("no", text="NO.")
        self.tree.heading("part", text="料号")
        self.tree.heading("desc", text="名称")
        self.tree.heading("qty", text="数量")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ===== 按钮 =====
        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Button(frame, text="📦 生成装箱单", command=self.export).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="🖨 生成并打印", command=self.export_and_print).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="🧹 清空", command=self.reset).pack(side=tk.LEFT, padx=5)

    def scan(self, event=None):
        code = self.input.get().strip()
        self.input.delete(0, tk.END)

        item = search(code, self.data)

        if not item:
            self.status.config(text=f"❌ 未找到：{code}", fg="red")
            return

        self.items.append(item)
        write_to_excel(item)

        self.tree.insert("", "end", values=(
            len(self.items),
            item["part"],
            item["desc"],
            item["qty"]
        ))

        self.status.config(text=f"✅ 已添加：{item['part']}", fg="green")

    def export(self):
        name = simpledialog.askstring("文件名", "输入装箱单名称（可空）")
        path = export_file(self.output_dir, name)

        messagebox.showinfo("完成", f"已生成：\n{path}")
        self.reset()

    def export_and_print(self):
        name = simpledialog.askstring("文件名", "输入装箱单名称（可空）")
        path = export_file(self.output_dir, name)

        print_file(path)

        messagebox.showinfo("完成", f"已打印：\n{path}")
        self.reset()

    def reset(self):
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        clear_template()
        self.input.focus_set()


# =========================
# 🚀 启动
# =========================
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()