import pandas as pd
import os
from openpyxl import load_workbook

INFO_FILE = "info.xlsx"
TEMPLATE_FILE = "template.xlsx"
OUTPUT_DIR = "output"

PART_COL = "料号/Part Number"
DESC_COL = "名称/Description"
QTY_COL = "数量/Quantity"
NO_COL = "NO."

os.makedirs(OUTPUT_DIR, exist_ok=True)


# =========================
# 🧼 清洗数据（统一 string）
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

        for col in df.columns:
            df[col] = df[col].apply(clean)

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
            row = res.iloc[0]
            return {
                "part_number": clean(row[PART_COL]),
                "description": clean(row.get(DESC_COL, "")),
                "quantity": clean(row.get(QTY_COL, ""))
            }

    return None


# =========================
# 📊 获取 template 列
# =========================
def get_cols(ws):
    header_row = 4
    cols = {}

    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
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
# ✍️ 写入一条 + NO自动编号
# =========================
def write_item(item):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    cols = get_cols(ws)

    pn_col = cols.get(PART_COL)
    desc_col = cols.get(DESC_COL)
    qty_col = cols.get(QTY_COL)
    no_col = cols.get(NO_COL)

    if not pn_col or not desc_col or not qty_col:
        raise Exception("template 缺少必要列")

    row = next_row(ws, pn_col)

    # 👉 NO 自动递增（核心）
    no_value = row - 4  # 因为数据从第5行开始

    ws.cell(row=row, column=no_col, value=no_value)
    ws.cell(row=row, column=pn_col, value=item["part_number"])
    ws.cell(row=row, column=desc_col, value=item["description"])
    ws.cell(row=row, column=qty_col, value=item["quantity"])

    wb.save(TEMPLATE_FILE)


# =========================
# 🧹 清空模板数据（保留表头）
# =========================
def clear():
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    for r in ws.iter_rows(min_row=5):
        for c in r:
            c.value = None

    wb.save(TEMPLATE_FILE)


# =========================
# 📦 导出装箱单
# =========================
def export(name=None):
    if not name:
        files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".xlsx")]
        name = str(len(files) + 1)

    path = os.path.join(OUTPUT_DIR, f"{name}.xlsx")

    wb = load_workbook(TEMPLATE_FILE)
    wb.save(path)

    clear()

    print(f"\n📦 装箱单生成成功：{path}")


# =========================
# 🚀 主流程
# =========================
def main():
    print("🚀 装箱系统启动\n")

    dfs = load_info()

    if not dfs:
        print("❌ info.xlsx 无数据")
        return

    while True:
        code = input("📡 扫码（done结束）：").strip()

        if code.lower() == "done":
            break

        item = search(code, dfs)

        if item:
            write_item(item)
            print(f"✅ 已写入：NO自动生成 → {item}")
        else:
            print(f"❌ 未找到料号：{code}")

    name = input("\n📝 装箱单名称（回车自动编号）：").strip()

    export(name if name else None)


if __name__ == "__main__":
    main()