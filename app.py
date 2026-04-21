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


# =========================
# 安全列读取
# =========================
def safe_cols(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=4, column=c).value
        if v:
            cols[str(v).strip()] = c
    return cols


# =========================
# 数据加载
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


# =========================
# 多结果查询
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
# 选择窗口
# =========================
def select_item(root, items):
    win = tk.Toplevel(root)
    win.title("选择物料")
    win.geometry("900x350")

    tree = ttk.Treeview(
        win,
        columns=("part","desc","qty","batch","edition","remarks"),
        show="headings"
    )

    headers = ["料号","名称","数量","批次","版本","备注"]
    widths = [120,200,80,120,100,200]

    for col, h, w in zip(tree["columns"], headers, widths):
        tree.heading(col, text=h)
        tree.column(col, width=w, anchor="center")

    tree.pack(fill=tk.BOTH, expand=True)

    for i in items:
        tree.insert("", "end", values=(
            i["part"], i["desc"], i["qty"],
            i["batch"], i["edition"], i["remarks"]
        ))

    result = {"val": None}

    def ok():
        sel = tree.selection()
        if sel:
            result["val"] = tree.item(sel[0])["values"]
        win.destroy()

    tk.Button(win, text="确认", command=ok).pack(pady=5)

    win.grab_set()
    root.wait_window(win)

    return result["val"]


# =========================
# 写入 Excel
# =========================
def write_to_excel(item, target_file):
    wb = load_workbook(target_file)
    ws = wb.active

    cols = safe_cols(ws)

    pn = cols.get(PART_COL)
    if not pn:
        messagebox.showerror("错误", "缺少料号列")
        return

    r = 5
    while ws.cell(row=r, column=pn).value:
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
        col = cols.get(k)
        if col:
            cell = ws.cell(row=r, column=col, value=v)
            cell.alignment = align

    wb.save(target_file)


# =========================
# 导出 / 打印
# =========================
def export_file(dir_path, name):
    if not name:
        files = [f for f in os.listdir(dir_path) if f.endswith(".xlsx")]
        name = str(len(files) + 1)

    path = os.path.join(dir_path, f"{name}.xlsx")

    wb = load_workbook(TEMPLATE_FILE)
    wb.save(path)

    clear_template()
    return path


def print_file(path):
    os.startfile(path, "print")


# =========================
# GUI
# =========================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 装箱系统 v2.2")
        self.root.geometry("950x550")

        clear_template()

        # 保持原始弹窗逻辑
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

        # 输入
        self.input = tk.Entry(root, font=("Arial", 16))
        self.input.pack(fill=tk.X, padx=10, pady=10)
        self.input.bind("<Return>", self.scan)
        self.input.focus_set()

        self.status = tk.Label(
            root,
            text=f"📁 新建模式 | {self.output_dir}",
            fg="blue"
        )
        self.status.pack()

        # 表格
        self.tree = ttk.Treeview(
            root,
            columns=("no","part","desc","qty","batch","edition","remarks"),
            show="headings"
        )

        headers = ["NO.","料号","名称","数量","批次","版本","备注"]
        widths = [60,120,200,80,120,100,200]

        for c, h, w in zip(self.tree["columns"], headers, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True)

        # 按钮容器
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(pady=10)

        tk.Button(self.btn_frame, text="📂 打开已有", width=14, command=self.open_existing).pack(side=tk.LEFT, padx=5)
        
        # 定义动态显示的按钮
        self.btn_export = tk.Button(self.btn_frame, text="📦 生成", width=14, command=self.export)
        self.btn_export.pack(side=tk.LEFT, padx=5)
        
        self.btn_print = tk.Button(self.btn_frame, text="🖨 生成并打印", width=16, command=self.export_print)
        self.btn_print.pack(side=tk.LEFT, padx=5)

        # 编辑模式独占：保存按钮（默认隐藏）
        self.btn_save_close = tk.Button(self.btn_frame, text="💾 保存并关闭", width=16, command=self.save_and_close, bg="#e3f2fd")

        # 公共按钮
        self.btn_clear = tk.Button(self.btn_frame, text="🧹 清空", width=10, command=self.reset)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

    # =========================
    # 扫码
    # =========================
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
            if not sel:
                return
            item = dict(zip(["part","desc","qty","batch","edition","remarks"], sel))

        self.items.append(item)

        target = self.current_file if self.current_file else TEMPLATE_FILE
        write_to_excel(item, target)

        self.tree.insert("", "end", values=(
            len(self.items),
            item["part"], item["desc"], item["qty"],
            item["batch"], item["edition"], item["remarks"]
        ))

        self.status.config(text=f"✅ 已录入 {item['part']}", fg="green")

    # =========================
    # 打开已有
    # =========================
    def open_existing(self):
        path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if not path:
            return

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
                "part": ws.cell(row=r, column=cols.get(PART_COL)).value,
                "desc": ws.cell(row=r, column=cols.get(DESC_COL)).value,
                "qty": ws.cell(row=r, column=cols.get(QTY_COL)).value,
                "batch": ws.cell(row=r, column=cols.get(BATCH_COL)).value if cols.get(BATCH_COL) else "",
                "edition": ws.cell(row=r, column=cols.get(EDITION_COL)).value if cols.get(EDITION_COL) else "",
                "remarks": ws.cell(row=r, column=cols.get(REMARKS_COL)).value if cols.get(REMARKS_COL) else "",
            }

            self.items.append(item)

            self.tree.insert("", "end", values=(
                len(self.items),
                item["part"], item["desc"], item["qty"],
                item["batch"], item["edition"], item["remarks"]
            ))

            r += 1

        self.status.config(text=f"📄 编辑模式：{os.path.basename(path)}", fg="green")
        
        # 切换按钮：隐藏生成，显示保存
        self.btn_export.pack_forget()
        self.btn_print.pack_forget()
        self.btn_save_close.pack(side=tk.LEFT, padx=5, before=self.btn_clear)

    # =========================
    # 动作逻辑
    # =========================
    def save_and_close(self):
        # 扫码时已实时保存，此处仅作提示并退出
        messagebox.showinfo("完成", f"更改已成功保存至：\n{os.path.basename(self.current_file)}")
        self.reset()

    def export(self):
        if not self.items:
            messagebox.showwarning("警告", "列表为空，无法生成")
            return
        name = simpledialog.askstring("文件名", "输入名称")
        path = export_file(self.output_dir, name)
        messagebox.showinfo("完成", f"已生成文件：\n{path}")
        self.reset()

    def export_print(self):
        if not self.items:
            messagebox.showwarning("警告", "列表为空，无法打印")
            return
        name = simpledialog.askstring("文件名", "输入名称")
        path = export_file(self.output_dir, name)
        print_file(path)
        messagebox.showinfo("完成", f"已打印并生成文件：\n{path}")
        self.reset()

    def reset(self):
        # 彻底清空显示条目
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        clear_template()
        self.current_file = None
        self.input.focus_set()
        self.status.config(text=f"📁 新建模式 | {self.output_dir}", fg="blue")
        
        # 恢复默认按钮布局
        self.btn_save_close.pack_forget()
        self.btn_export.pack(side=tk.LEFT, padx=5, before=self.btn_clear)
        self.btn_print.pack(side=tk.LEFT, padx=5, before=self.btn_clear)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()