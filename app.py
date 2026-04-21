import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# =========================
# 配置
# =========================
BASE_OUTPUT_DIR = "output"

PART_COL = "料号/Part Number"
DESC_COL = "名称/Description"
QTY_COL = "数量/Quantity"
BATCH_COL = "批次/Batch"
EDITION_COL = "版本/Edition"
REMARKS_COL = "备注/Remarks"
NO_COL = "NO."

INFO_FILE = "info.xlsx"
TEMPLATE_FILE = "template.xlsx"

os.makedirs(BASE_OUTPUT_DIR, exist_ok=True)


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


# =========================
# 数据加载
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
# 查找（多结果）
# =========================
def search_all(part, dfs):
    part = clean(part)
    results = []

    for df in dfs:
        res = df[df[PART_COL] == part]

        for _, r in res.iterrows():
            results.append({
                "part": clean(r.get(PART_COL)),
                "desc": clean(r.get(DESC_COL)),
                "qty": clean(r.get(QTY_COL)),
                "batch": clean(r.get(BATCH_COL)),
                "edition": clean(r.get(EDITION_COL)),
                "remarks": clean(r.get(REMARKS_COL)),
            })

    return results


# =========================
# 选择窗口（优化UI）
# =========================
def select_item_popup(root, items):
    win = tk.Toplevel(root)
    win.title("选择物料")
    win.geometry("900x350")

    tree = ttk.Treeview(
        win,
        columns=("part", "desc", "qty", "batch", "edition", "remarks"),
        show="headings"
    )

    headers = ["料号", "名称", "数量", "批次", "版本", "备注"]
    widths = [120, 200, 80, 120, 100, 200]

    for col, txt, w in zip(tree["columns"], headers, widths):
        tree.heading(col, text=txt)
        tree.column(col, width=w, anchor="center")

    tree.pack(fill=tk.BOTH, expand=True)

    for item in items:
        tree.insert("", "end", values=(
            item["part"], item["desc"], item["qty"],
            item["batch"], item["edition"], item["remarks"]
        ))

    selected = {"value": None}

    def confirm():
        sel = tree.selection()
        if not sel:
            return
        selected["value"] = tree.item(sel[0])["values"]
        win.destroy()

    tk.Button(win, text="✔ 确定", command=confirm).pack(pady=8)

    win.grab_set()
    root.wait_window(win)

    return selected["value"]


# =========================
# 写入Excel（居中）
# =========================
def write_to_excel(item):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=4, column=c).value
        if v:
            cols[str(v).strip()] = c

    r = 5
    while ws.cell(row=r, column=cols[PART_COL]).value:
        r += 1

    align = Alignment(horizontal="center", vertical="center")

    mapping = {
        NO_COL: r - 4,
        PART_COL: item["part"],
        DESC_COL: item["desc"],
        QTY_COL: item["qty"],
        BATCH_COL: item["batch"],
        EDITION_COL: item["edition"],
        REMARKS_COL: item["remarks"],
    }

    for k, v in mapping.items():
        if k in cols:
            cell = ws.cell(row=r, column=cols[k], value=v)
            cell.alignment = align

    wb.save(TEMPLATE_FILE)


def clear_template():
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    for r in ws.iter_rows(min_row=5):
        for c in r:
            c.value = None

    wb.save(TEMPLATE_FILE)


def export_file(output_dir, name):
    if not name:
        files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx")]
        name = str(len(files) + 1)

    path = os.path.join(output_dir, f"{name}.xlsx")

    wb = load_workbook(TEMPLATE_FILE)
    wb.save(path)
    clear_template()

    return path


def print_file(path):
    os.startfile(path, "print")


# =========================
# GUI 主程序
# =========================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 装箱系统（升级版）")
        self.root.geometry("950x550")

        order_no = simpledialog.askstring("订单号", "请输入订单号（可空）")

        # 修复最小化问题
        self.root.deiconify()
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(100, lambda: self.root.attributes('-topmost', False))
        self.root.focus_force()

        self.output_dir = os.path.join(BASE_OUTPUT_DIR, order_no) if order_no else BASE_OUTPUT_DIR
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
        self.tree = ttk.Treeview(
            root,
            columns=("no","part","desc","qty","batch","edition","remarks"),
            show="headings"
        )

        headers = ["NO.","料号","名称","数量","批次","版本","备注"]
        widths = [60,120,200,80,120,100,200]

        for col, txt, w in zip(self.tree["columns"], headers, widths):
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ===== 按钮 =====
        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Button(frame, text="📦 生成装箱单", width=18, command=self.export).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="🖨 生成并打印", width=18, command=self.export_and_print).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="🧹 清空", width=12, command=self.reset).pack(side=tk.LEFT, padx=5)

    def scan(self, event=None):
        code = self.input.get().strip()
        self.input.delete(0, tk.END)

        results = search_all(code, self.data)

        if not results:
            self.status.config(text=f"❌ 未找到：{code}", fg="red")
            return

        if len(results) == 1:
            item = results[0]
        else:
            selected = select_item_popup(self.root, results)
            if not selected:
                return
            item = {
                "part": selected[0],
                "desc": selected[1],
                "qty": selected[2],
                "batch": selected[3],
                "edition": selected[4],
                "remarks": selected[5],
            }

        self.items.append(item)
        write_to_excel(item)

        self.tree.insert("", "end", values=(
            len(self.items),
            item["part"],
            item["desc"],
            item["qty"],
            item["batch"],
            item["edition"],
            item["remarks"]
        ))

        self.status.config(text=f"✅ 已添加：{item['part']}", fg="green")

    def export(self):
        name = simpledialog.askstring("文件名", "输入名称（可空）")
        path = export_file(self.output_dir, name)
        messagebox.showinfo("完成", f"已生成：\n{path}")
        self.reset()

    def export_and_print(self):
        name = simpledialog.askstring("文件名", "输入名称（可空）")
        path = export_file(self.output_dir, name)
        print_file(path)
        messagebox.showinfo("完成", f"已打印：\n{path}")
        self.reset()

    def reset(self):
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        clear_template()
        self.input.focus_set()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()