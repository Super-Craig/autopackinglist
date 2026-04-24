import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
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


# =========================
# 工具
# =========================
def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def clear_template():
    if not os.path.exists(TEMPLATE_FILE):
        return
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active
    for r in ws.iter_rows(min_row=5):
        for c in r:
            c.value = None
    wb.save(TEMPLATE_FILE)


def safe_cols(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=4, column=c).value
        if v:
            cols[str(v).strip()] = c
    return cols


# =========================
# 数据加载与查询
# =========================
def load_info():
    if not os.path.exists(INFO_FILE):
        return []
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


def select_item(root, items):
    win = tk.Toplevel(root)
    win.title("选择物料")
    win.geometry("900x350")
    tree = ttk.Treeview(win, columns=("part","desc","qty","batch","edition","remarks"), show="headings")
    headers = ["料号","名称","数量","批次","版本","备注"]
    widths = [120,200,80,120,100,200]
    for col, h, w in zip(tree["columns"], headers, widths):
        tree.heading(col, text=h)
        tree.column(col, width=w, anchor="center")
    tree.pack(fill=tk.BOTH, expand=True)
    for i in items:
        tree.insert("", "end", values=(i["part"], i["desc"], i["qty"], i["batch"], i["edition"], i["remarks"]))
    result = {"val": None}
    def ok():
        sel = tree.selection()
        if sel: result["val"] = tree.item(sel[0])["values"]
        win.destroy()
    tk.Button(win, text="确认", command=ok).pack(pady=5)
    win.grab_set()
    root.wait_window(win)
    return result["val"]


# =========================
# Excel 写入逻辑 (支持重写)
# =========================
def write_items_to_excel(items, target_file):
    """根据 items 列表重写 Excel 整个数据区"""
    wb = load_workbook(target_file)
    ws = wb.active
    cols = safe_cols(ws)
    
    # 1. 先清空第 5 行及之后的数据
    for r in ws.iter_rows(min_row=5):
        for c in r:
            c.value = None

    # 2. 重新写入
    align = Alignment(horizontal="center", vertical="center")
    for i, item in enumerate(items):
        row_idx = 5 + i
        mapping = {
            NO_COL: i + 1,
            PART_COL: item["part"],
            DESC_COL: item["desc"],
            QTY_COL: item["qty"],
            BATCH_COL: item["batch"],
            EDITION_COL: item["edition"],
            REMARKS_COL: item["remarks"],
        }
        for k, v in mapping.items():
            c_idx = cols.get(k)
            if c_idx:
                cell = ws.cell(row=row_idx, column=c_idx, value=v)
                cell.alignment = align
    wb.save(target_file)


# =========================
# GUI
# =========================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 凯实售后装箱系统 v2.4")
        self.root.geometry("950x550")

        clear_template()
        order = simpledialog.askstring("订单号", "请输入订单号（可空）")

        self.root.deiconify()
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(100, lambda: self.root.attributes('-topmost', False))
        self.root.focus_force()

        self.output_dir = os.path.join(BASE_OUTPUT_DIR, order) if order else BASE_OUTPUT_DIR
        os.makedirs(self.output_dir, exist_ok=True)

        self.data = load_info()
        self.items = []
        self.current_file = None

        self.input = tk.Entry(root, font=("Arial", 16))
        self.input.pack(fill=tk.X, padx=10, pady=10)
        self.input.bind("<Return>", self.scan)
        self.input.focus_set()

        self.status = tk.Label(root, text=f"📁 新建模式 | {self.output_dir}", fg="blue")
        self.status.pack()

        self.tree = ttk.Treeview(root, columns=("no","part","desc","qty","batch","edition","remarks"), show="headings")
        headers = ["NO.","料号","名称","数量","批次","版本","备注"]
        widths = [60,120,200,80,120,100,200]
        for c, h, w in zip(self.tree["columns"], headers, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True)

        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Button(frame, text="📂 打开已有", width=14, command=self.open_existing).pack(side=tk.LEFT, padx=5)
        self.btn_export = tk.Button(frame, text="📦 生成", width=14, command=self.export)
        self.btn_export.pack(side=tk.LEFT, padx=5)
        self.btn_print = tk.Button(frame, text="🖨 生成并打印", width=16, command=self.export_print)
        self.btn_print.pack(side=tk.LEFT, padx=5)

        # 编辑模式按钮
        self.btn_save_close = tk.Button(frame, text="💾 保存并关闭", width=16, command=self.save_and_close, bg="#e3f2fd")

        # 原“清空”按钮改名为“删除选中”
        self.btn_delete = tk.Button(frame, text="❌ 删除选中", width=14, command=self.delete_selected, fg="red")
        self.btn_delete.pack(side=tk.LEFT, padx=5)

    def scan(self, event=None):
        code = self.input.get().strip()
        self.input.delete(0, tk.END)
        res = search_all(code, self.data)
        if not res:
            self.status.config(text=f"❌ 未找到 {code}", fg="red")
            return

        if len(res) == 1:
            item = res[0]
        else:
            sel = select_item(self.root, res)
            if not sel: return
            item = dict(zip(["part","desc","qty","batch","edition","remarks"], sel))

        self.items.append(item)
        
        # 实时写入 Excel
        target = self.current_file if self.current_file else TEMPLATE_FILE
        write_items_to_excel(self.items, target)

        self.tree.insert("", "end", values=(
            len(self.items), item["part"], item["desc"], item["qty"],
            item["batch"], item["edition"], item["remarks"]
        ))
        self.status.config(text=f"✅ 已录入 {item['part']}", fg="green")

    def delete_selected(self):
        """逐条删除选中的项"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先在表格中选择要删除的行")
            return

        if not messagebox.askyesno("确认", "确定要删除选中的行吗？"):
            return

        # 获取所有选中行的索引
        indices = sorted([self.tree.index(item) for item in selected], reverse=True)
        
        # 从数据列表和表格中移除
        for idx in indices:
            del self.items[idx]
        
        for item in selected:
            self.tree.delete(item)

        # 重新排序 Treeview 的序号列
        for i, child in enumerate(self.tree.get_children()):
            vals = list(self.tree.item(child, 'values'))
            vals[0] = i + 1  # 更新 NO.
            self.tree.item(child, values=vals)

        # 同步更新 Excel 文件
        target = self.current_file if self.current_file else TEMPLATE_FILE
        write_items_to_excel(self.items, target)
        
        self.status.config(text="⚠️ 已删除选中项并同步文件", fg="orange")

    def open_existing(self):
        path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if not path: return
        self.current_file = path
        wb = load_workbook(path)
        ws = wb.active
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        cols = safe_cols(ws)
        pn = cols.get(PART_COL)
        if not pn:
            messagebox.showerror("错误", "缺少料号列")
            return
        r = 5
        while ws.cell(row=r, column=pn).value:
            item = {
                "part": clean(ws.cell(row=r, column=cols.get(PART_COL)).value),
                "desc": clean(ws.cell(row=r, column=cols.get(DESC_COL)).value),
                "qty": clean(ws.cell(row=r, column=cols.get(QTY_COL)).value),
                "batch": clean(ws.cell(row=r, column=cols.get(BATCH_COL)).value) if cols.get(BATCH_COL) else "",
                "edition": clean(ws.cell(row=r, column=cols.get(EDITION_COL)).value) if cols.get(EDITION_COL) else "",
                "remarks": clean(ws.cell(row=r, column=cols.get(REMARKS_COL)).value) if cols.get(REMARKS_COL) else "",
            }
            self.items.append(item)
            self.tree.insert("", "end", values=(len(self.items), item["part"], item["desc"], item["qty"], item["batch"], item["edition"], item["remarks"]))
            r += 1

        self.status.config(text=f"📄 编辑模式：{os.path.basename(path)}", fg="green")
        self.btn_export.pack_forget()
        self.btn_print.pack_forget()
        self.btn_save_close.pack(side=tk.LEFT, padx=5, before=self.btn_delete)

    def save_and_close(self):
        messagebox.showinfo("完成", f"更改已成功保存至：\n{os.path.basename(self.current_file)}")
        self.reset_ui()

    def export(self):
        if not self.items: return
        name = simpledialog.askstring("文件名", "输入名称")
        if not name:
            files = [f for f in os.listdir(self.output_dir) if f.endswith(".xlsx")]
            name = str(len(files) + 1)
        path = os.path.join(self.output_dir, f"{name}.xlsx")
        wb = load_workbook(TEMPLATE_FILE)
        wb.save(path)
        messagebox.showinfo("完成", f"已生成文件：\n{path}")
        self.reset_ui()

    def export_print(self):
        if not self.items: return
        name = simpledialog.askstring("文件名", "输入名称")
        if not name:
            files = [f for f in os.listdir(self.output_dir) if f.endswith(".xlsx")]
            name = str(len(files) + 1)
        path = os.path.join(self.output_dir, f"{name}.xlsx")
        wb = load_workbook(TEMPLATE_FILE)
        wb.save(path)
        os.startfile(path, "print")
        messagebox.showinfo("完成", f"已打印并生成文件：\n{path}")
        self.reset_ui()

    def reset_ui(self):
        """恢复到初始新建状态"""
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        clear_template()
        self.current_file = None
        self.input.focus_set()
        self.status.config(text=f"📁 新建模式 | {self.output_dir}", fg="blue")
        self.btn_save_close.pack_forget()
        self.btn_export.pack(side=tk.LEFT, padx=5, before=self.btn_delete)
        self.btn_print.pack(side=tk.LEFT, padx=5, before=self.btn_delete)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()