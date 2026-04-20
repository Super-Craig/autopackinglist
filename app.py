import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from openpyxl import load_workbook

# =========================
# 📁 文件配置
# =========================
INFO_FILE = "info.xlsx"
TEMPLATE_FILE = "template.xlsx"
OUTPUT_DIR = "output"

PART_COL = "料号/Part Number"
DESC_COL = "名称/Description"
QTY_COL = "数量/Quantity"
NO_COL = "NO."

os.makedirs(OUTPUT_DIR, exist_ok=True)


# =========================
# 🧼 清洗数据
# =========================
def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


# =========================
# 📥 读取 info.xlsx
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
# 🔍 查找料号
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
# 📊 template列解析
# =========================
def get_cols(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=4, column=c).value
        if v:
            cols[str(v).strip()] = c
    return cols


# =========================
# 📍 找下一行
# =========================
def next_row(ws, pn_col):
    r = 5
    while ws.cell(row=r, column=pn_col).value:
        r += 1
    return r


# =========================
# ✍️ 写入 Excel（含NO）
# =========================
def write_to_excel(item):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    cols = get_cols(ws)

    pn_col = cols.get(PART_COL)
    desc_col = cols.get(DESC_COL)
    qty_col = cols.get(QTY_COL)
    no_col = cols.get(NO_COL)

    if not pn_col or not desc_col or not qty_col:
        raise Exception("template.xlsx 缺少必要列（第4行表头）")

    row = next_row(ws, pn_col)

    # 👉 自动 NO 编号
    no_value = row - 4

    ws.cell(row=row, column=no_col, value=no_value)
    ws.cell(row=row, column=pn_col, value=item["part"])
    ws.cell(row=row, column=desc_col, value=item["desc"])
    ws.cell(row=row, column=qty_col, value=item["qty"])

    wb.save(TEMPLATE_FILE)


# =========================
# 🧹 清空模板
# =========================
def clear_template():
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    for r in ws.iter_rows(min_row=5):
        for c in r:
            c.value = None

    wb.save(TEMPLATE_FILE)


# =========================
# 📦 导出装箱单
# =========================
def export_file(name):
    if not name:
        files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".xlsx")]
        name = str(len(files) + 1)

    path = os.path.join(OUTPUT_DIR, f"{name}.xlsx")

    wb = load_workbook(TEMPLATE_FILE)
    wb.save(path)

    clear_template()

    return path


# =========================
# 🖥 GUI 主程序
# =========================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 装箱扫码系统")
        self.root.geometry("750x520")

        self.data = load_info()
        self.items = []

        # ===== 扫码输入 =====
        self.input = tk.Entry(root, font=("Arial", 16))
        self.input.pack(fill=tk.X, padx=10, pady=10)
        self.input.bind("<Return>", self.scan)
        self.input.focus_set()

        # ===== 状态 =====
        self.status = tk.Label(root, text="请扫码...", fg="blue")
        self.status.pack()

        # ===== 表格 =====
        self.tree = ttk.Treeview(
            root,
            columns=("no", "part", "desc", "qty"),
            show="headings"
        )

        self.tree.heading("no", text="NO.")
        self.tree.heading("part", text="料号")
        self.tree.heading("desc", text="名称")
        self.tree.heading("qty", text="数量")

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ===== 按钮 =====
        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Button(frame, text="📦 生成装箱单", command=self.export).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="🧹 清空", command=self.reset).pack(side=tk.LEFT, padx=5)

    # =========================
    # 📡 扫码逻辑
    # =========================
    def scan(self, event=None):
        code = self.input.get().strip()
        self.input.delete(0, tk.END)

        if not code:
            return

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

    # =========================
    # 📦 导出
    # =========================
    def export(self):
        name = simpledialog.askstring("文件名", "输入装箱单名称（可空）")
        path = export_file(name)

        messagebox.showinfo("完成", f"已生成：\n{path}")

        self.tree.delete(*self.tree.get_children())
        self.items.clear()

        self.status.config(text="已生成装箱单", fg="blue")
        self.input.focus_set()

    # =========================
    # 🧹 重置
    # =========================
    def reset(self):
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        clear_template()

        self.status.config(text="已清空", fg="blue")
        self.input.focus_set()


# =========================
# 🚀 启动
# =========================
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()