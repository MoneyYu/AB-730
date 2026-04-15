# -*- coding: utf-8 -*-
"""
AB-730 英業達課程 — 建立補充 Excel 範例檔案
新增：PO 採購訂單追蹤、ERP 原始出貨記錄
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
from datetime import datetime, timedelta
import random

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

HEADER_FONT = Font(name="Microsoft JhengHei", bold=True, size=11, color="FFFFFF")
DATA_FONT = Font(name="Microsoft JhengHei", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

def style_header_row(ws, row, max_col, fill_color="2F5496"):
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = CENTER
        cell.border = THIN_BORDER

def style_data_rows(ws, start_row, end_row, max_col):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = DATA_FONT
            cell.alignment = CENTER
            cell.border = THIN_BORDER

def auto_width(ws, min_width=10, max_width=35):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        width = min(max(max_len + 4, min_width), max_width)
        ws.column_dimensions[col_letter].width = width

def add_table(ws, ref, name):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tab)


# ═══════════════════════════════════════════════════════════
# 檔案: 英業達_PO採購訂單追蹤.xlsx
# ═══════════════════════════════════════════════════════════
def create_po_tracker():
    wb = openpyxl.Workbook()

    # ── Sheet 1: GPU Module PO 追蹤 ──
    ws1 = wb.active
    ws1.title = "GPU Module PO追蹤"
    ws1.sheet_properties.tabColor = "2F5496"

    headers = ["PO編號", "供應商", "零件型號", "產品線", "下單日期",
               "預計到貨日", "數量(pcs)", "單價(USD)", "金額(USD)",
               "PO狀態", "實際到貨日", "備註"]
    data = [
        ["PO-2025-03-0921", "NovaTech", "NT-500", "ProServer X200", "2025/3/10", "2025/5/5", 4000, 285, 1140000, "已到貨", "2025/5/3", "準時"],
        ["PO-2025-03-0935", "NovaTech", "NT-500", "ProServer X100", "2025/3/18", "2025/5/13", 8000, 285, 2280000, "已到貨", "2025/5/15", "延遲2天"],
        ["PO-2025-04-0988", "NovaTech", "NT-500", "ProServer X200", "2025/4/5", "2025/5/31", 6000, 285, 1710000, "已到貨", "2025/6/2", "延遲2天"],
        ["PO-2025-04-1015", "NovaTech", "NT-500", "ProServer X100", "2025/4/15", "2025/6/10", 10000, 285, 2850000, "已到貨", "2025/6/12", "延遲2天"],
        ["PO-2025-05-1102", "NovaTech", "NT-500", "ProServer X200", "2025/5/8", "2025/7/3", 8000, 285, 2280000, "已到貨", "2025/7/5", "延遲2天"],
        ["PO-2025-05-1130", "NovaTech", "NT-500", "ProServer X100", "2025/5/20", "2025/7/15", 12000, 285, 3420000, "已到貨", "2025/7/18", "延遲3天"],
        ["PO-2025-06-1205", "NovaTech", "NT-500", "CloudBook L15", "2025/6/3", "2025/7/29", 5000, 285, 1425000, "已到貨", "2025/7/30", "延遲1天"],
        ["PO-2025-06-1248", "NovaTech", "NT-500", "ProServer X200", "2025/6/18", "2025/8/13", 8000, 285, 2280000, "已到貨", "2025/8/20", "延遲7天"],
        # 在途 PO
        ["PO-2025-07-1310", "NovaTech", "NT-500", "ProServer X100", "2025/7/2", "2025/8/27", 15000, 285, 4275000, "在途", "", "已出貨，海運中"],
        ["PO-2025-07-1345", "NovaTech", "NT-500", "ProServer X200", "2025/7/10", "2025/9/4", 10000, 285, 2850000, "在途", "", "已出貨，海運中"],
        ["PO-2025-07-1380", "NovaTech", "NT-500", "CloudBook L15", "2025/7/18", "2025/9/12", 6000, 285, 1710000, "在途", "", "NovaTech 確認已備料"],
        # 未出貨 PO（會受斷料影響）
        ["PO-2025-08-1420", "NovaTech", "NT-500", "ProServer X100", "2025/8/1", "2025/9/26", 18000, 285, 5130000, "已確認", "", "⚠️ 可能受斷料影響"],
        ["PO-2025-08-1455", "NovaTech", "NT-500", "ProServer X200", "2025/8/8", "2025/10/3", 12000, 285, 3420000, "已確認", "", "⚠️ 可能受斷料影響"],
        ["PO-2025-08-1490", "NovaTech", "NT-500", "CloudBook L15", "2025/8/15", "2025/10/10", 8000, 285, 2280000, "待確認", "", "⚠️ 尚未確認，可能受影響"],
        ["PO-2025-09-1520", "NovaTech", "NT-500", "ProServer X100", "2025/9/1", "2025/10/27", 20000, 285, 5700000, "規劃中", "", "⚠️ 預計下單，可能受影響"],
        ["PO-2025-09-1555", "NovaTech", "NT-500", "ProServer X200", "2025/9/5", "2025/10/31", 10000, 285, 2850000, "規劃中", "", "⚠️ 預計下單，可能受影響"],
    ]

    for c, h in enumerate(headers, 1):
        ws1.cell(row=1, column=c, value=h)
    style_header_row(ws1, 1, len(headers))

    for r, row_data in enumerate(data, 2):
        for c, val in enumerate(row_data, 1):
            ws1.cell(row=r, column=c, value=val)
    style_data_rows(ws1, 2, len(data) + 1, len(headers))

    # Color-code PO status
    status_colors = {
        "已到貨": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "在途": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
        "已確認": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        "待確認": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        "規劃中": PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
    }
    for r in range(2, len(data) + 2):
        status = ws1.cell(row=r, column=10).value
        if status in status_colors:
            ws1.cell(row=r, column=10).fill = status_colors[status]

    for r in range(2, len(data) + 2):
        ws1.cell(row=r, column=9).number_format = '#,##0'
        ws1.cell(row=r, column=7).number_format = '#,##0'
        ws1.cell(row=r, column=12).alignment = LEFT_WRAP

    auto_width(ws1, max_width=25)
    add_table(ws1, f"A1:L{len(data)+1}", "GPU_PO追蹤")

    # ── Sheet 2: 庫存狀況 ──
    ws2 = wb.create_sheet("安全庫存狀況")
    ws2.sheet_properties.tabColor = "548235"

    headers2 = ["零件型號", "供應商", "產品線", "月均用量(pcs)", "安全庫存目標(週)",
                "目前庫存(pcs)", "可用週數", "庫存狀態", "下次到貨PO", "預計到貨日"]
    data2 = [
        ["NT-500", "NovaTech", "ProServer X100", 20000, 2, 12500, 2.5, "🟡 偏低", "PO-2025-07-1310", "2025/8/27"],
        ["NT-500", "NovaTech", "ProServer X200", 10000, 2, 8200, 3.3, "🟢 正常", "PO-2025-07-1345", "2025/9/4"],
        ["NT-500", "NovaTech", "CloudBook L15", 5000, 2, 3800, 3.0, "🟢 正常", "PO-2025-07-1380", "2025/9/12"],
        ["NT-D5-32G", "NovaTech", "ProServer X100", 80000, 4, 95000, 4.8, "🟢 正常", "PO-2025-07-D210", "2025/8/20"],
        ["NT-D5-32G", "NovaTech", "ProServer X200", 40000, 4, 52000, 5.2, "🟢 正常", "PO-2025-07-D225", "2025/9/1"],
        ["MP-DDR5-32", "MemoryPlus", "ProServer X100", 30000, 2, 18000, 2.4, "🟡 偏低", "PO-2025-08-M108", "2025/9/15"],
        ["PC-2000W", "PowerCore", "ProServer X100", 8000, 3, 15000, 7.5, "🟢 充足", "PO-2025-08-P055", "2025/9/20"],
        ["PC-2000W", "PowerCore", "ProServer X200", 5000, 3, 12000, 9.6, "🟢 充足", "PO-2025-09-P062", "2025/10/5"],
    ]

    for c, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header_row(ws2, 1, len(headers2), "548235")

    for r, row_data in enumerate(data2, 2):
        for c, val in enumerate(row_data, 1):
            ws2.cell(row=r, column=c, value=val)
    style_data_rows(ws2, 2, len(data2) + 1, len(headers2))

    for r in range(2, len(data2) + 2):
        ws2.cell(row=r, column=4).number_format = '#,##0'
        ws2.cell(row=r, column=6).number_format = '#,##0'

    auto_width(ws2)
    add_table(ws2, f"A1:J{len(data2)+1}", "安全庫存")

    # ── Sheet 3: DRAM PO 追蹤 ──
    ws3 = wb.create_sheet("DRAM PO追蹤")
    ws3.sheet_properties.tabColor = "C55A11"

    headers3 = ["PO編號", "供應商", "零件型號", "產品線", "下單日期",
                "預計到貨日", "數量(pcs)", "單價(USD)", "金額(USD)",
                "PO狀態", "實際到貨日", "備註"]
    data3 = [
        ["PO-2025-04-D180", "NovaTech", "NT-D5-32G", "ProServer X100", "2025/4/10", "2025/5/8", 40000, 42, 1680000, "已到貨", "2025/5/8", "準時"],
        ["PO-2025-04-D195", "MemoryPlus", "MP-DDR5-32", "ProServer X100", "2025/4/15", "2025/5/27", 30000, 38, 1140000, "已到貨", "2025/5/30", "延遲3天"],
        ["PO-2025-05-D210", "NovaTech", "NT-D5-32G", "ProServer X200", "2025/5/5", "2025/6/2", 20000, 42, 840000, "已到貨", "2025/6/3", "延遲1天"],
        ["PO-2025-06-D230", "MemoryPlus", "MP-DDR5-32", "ProServer X100", "2025/6/10", "2025/7/22", 30000, 38, 1140000, "已到貨", "2025/7/25", "延遲3天"],
        ["PO-2025-07-D250", "NovaTech", "NT-D5-32G", "ProServer X100", "2025/7/8", "2025/8/5", 50000, 42, 2100000, "在途", "", "已出貨"],
        ["PO-2025-07-D265", "MemoryPlus", "MP-DDR5-32", "CloudBook L15", "2025/7/20", "2025/8/31", 40000, 38, 1520000, "已確認", "", ""],
        ["PO-2025-08-M108", "MemoryPlus", "MP-DDR5-32", "ProServer X100", "2025/8/5", "2025/9/16", 30000, 38, 1140000, "規劃中", "", "⚠️ 考慮品質問題是否改用 NovaTech"],
    ]

    for c, h in enumerate(headers3, 1):
        ws3.cell(row=1, column=c, value=h)
    style_header_row(ws3, 1, len(headers3), "C55A11")

    for r, row_data in enumerate(data3, 2):
        for c, val in enumerate(row_data, 1):
            ws3.cell(row=r, column=c, value=val)
    style_data_rows(ws3, 2, len(data3) + 1, len(headers3))

    for r in range(2, len(data3) + 2):
        ws3.cell(row=r, column=9).number_format = '#,##0'
        ws3.cell(row=r, column=7).number_format = '#,##0'

    auto_width(ws3, max_width=25)
    add_table(ws3, f"A1:L{len(data3)+1}", "DRAM_PO追蹤")

    path = os.path.join(OUTPUT_DIR, "英業達_PO採購訂單追蹤.xlsx")
    wb.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案: 英業達_ERP原始出貨記錄.xlsx
# ═══════════════════════════════════════════════════════════
def create_erp_raw_data():
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "ERP原始出貨記錄"
    ws.sheet_properties.tabColor = "7030A0"

    headers = ["出貨單號", "出貨日期", "客戶代碼", "客戶名稱", "產品料號",
               "產品名稱", "出貨數量(台)", "單價(USD)", "金額(USD)",
               "出貨倉庫", "運送方式", "目的地"]

    # Generate realistic shipping records
    products = [
        ("PS-X100-A01", "ProServer X100 Standard", 3200, "PS"),
        ("PS-X100-B02", "ProServer X100 High-Mem", 3800, "PS"),
        ("PS-X100-C03", "ProServer X100 GPU-Accel", 4200, "PS"),
        ("CB-L15-A01", "CloudBook L15 i5-8G", 680, "CB"),
        ("CB-L15-B02", "CloudBook L15 i7-16G", 820, "CB"),
        ("CB-L15-C03", "CloudBook L15 i7-32G", 1050, "CB"),
        ("IG-G3-A01", "IoT Gateway G3 Basic", 285, "IG"),
        ("IG-G3-B02", "IoT Gateway G3 Pro", 380, "IG"),
    ]

    customers = [
        ("DELL-US", "Dell Technologies (US)"),
        ("DELL-EU", "Dell Technologies (EU)"),
        ("DELL-AP", "Dell Technologies (APAC)"),
    ]

    warehouses = ["桃園廠 WH-A", "桃園廠 WH-B", "上海廠 WH-C"]
    shipping = ["海運", "空運", "海運"]
    destinations = ["Austin, TX", "Limerick, Ireland", "Penang, Malaysia"]

    random.seed(42)
    data = []
    ship_num = 30001

    for month in range(1, 13):
        for _ in range(random.randint(25, 40)):
            prod = random.choice(products)
            cust_idx = random.randint(0, 2)
            cust = customers[cust_idx]

            day = random.randint(1, 28)
            date_str = f"2025/{month:02d}/{day:02d}"

            # Server: 50-500 units; Notebook: 500-5000; IoT: 100-1000
            if prod[3] == "PS":
                qty = random.randint(50, 500)
            elif prod[3] == "CB":
                qty = random.randint(500, 5000)
            else:
                qty = random.randint(100, 1000)

            amount = qty * prod[2]

            data.append([
                f"SO-2025-{ship_num}",
                date_str,
                cust[0],
                cust[1],
                prod[0],
                prod[1],
                qty,
                prod[2],
                amount,
                warehouses[cust_idx],
                shipping[cust_idx],
                destinations[cust_idx],
            ])
            ship_num += 1

    # Sort by date
    data.sort(key=lambda x: x[1])

    # Add some duplicates (for data cleaning exercise)
    dupes = [data[5].copy(), data[20].copy(), data[45].copy()]
    for d in dupes:
        data.append(d)

    # Add some with inconsistent date formats (for cleaning exercise)
    data.append(["SO-2025-39001", "2025-03-15", "DELL-US", "Dell Technologies (US)", "PS-X100-A01", "ProServer X100 Standard", 200, 3200, 640000, "桃園廠 WH-A", "海運", "Austin, TX"])
    data.append(["SO-2025-39002", "Mar 22, 2025", "DELL-EU", "Dell Technologies (EU)", "CB-L15-B02", "CloudBook L15 i7-16G", 1500, 820, 1230000, "桃園廠 WH-B", "空運", "Limerick, Ireland"])
    data.append(["SO-2025-39003", "15/06/2025", "DELL-AP", "Dell Technologies (APAC)", "IG-G3-A01", "IoT Gateway G3 Basic", 300, 285, 85500, "上海廠 WH-C", "海運", "Penang, Malaysia"])

    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header_row(ws, 1, len(headers), "7030A0")

    for r, row_data in enumerate(data, 2):
        for c, val in enumerate(row_data, 1):
            ws.cell(row=r, column=c, value=val)
    style_data_rows(ws, 2, len(data) + 1, len(headers))

    for r in range(2, len(data) + 2):
        ws.cell(row=r, column=7).number_format = '#,##0'
        ws.cell(row=r, column=9).number_format = '#,##0'

    auto_width(ws, max_width=30)
    add_table(ws, f"A1:L{len(data)+1}", "ERP出貨記錄")

    path = os.path.join(OUTPUT_DIR, "英業達_ERP原始出貨記錄.xlsx")
    wb.save(path)
    print(f"✅ Created: {path} ({len(data)} records)")


if __name__ == "__main__":
    print("=" * 60)
    print("  AB-730 英業達 — 建立補充 Excel 範例檔案")
    print("=" * 60)
    create_po_tracker()
    create_erp_raw_data()
    print("=" * 60)
    print("  All supplementary Excel files created!")
    print("=" * 60)
