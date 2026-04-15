# -*- coding: utf-8 -*-
"""
AB-730 英業達課程 — 建立所有 Excel 範例檔案
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
import os

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Shared styles ──
HEADER_FONT = Font(name="Microsoft JhengHei", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FILL_GREEN = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
HEADER_FILL_ORANGE = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")
HEADER_FILL_PURPLE = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
DATA_FONT = Font(name="Microsoft JhengHei", size=10)
TITLE_FONT = Font(name="Microsoft JhengHei", bold=True, size=14, color="2F5496")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

def style_header_row(ws, row, max_col, fill=None):
    if fill is None:
        fill = HEADER_FILL
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
    from openpyxl.worksheet.table import Table, TableStyleInfo
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tab)


# ═══════════════════════════════════════════════════════════
# 檔案 1: 英業達_供應商評估表.xlsx
# ═══════════════════════════════════════════════════════════
def create_supplier_evaluation():
    wb = openpyxl.Workbook()

    # ── Sheet 1: GPU 供應商評估 ──
    ws1 = wb.active
    ws1.title = "GPU供應商評估"
    ws1.sheet_properties.tabColor = "2F5496"

    headers = ["評估項目", "NovaTech (韓國)", "ChipMax (中國)", "SiliconEdge (日本)"]
    data = [
        ["零件型號", "NT-500", "CM-G5", "SE-Pro"],
        ["產品類別", "GPU Module (PCIe Gen5)", "GPU Module (PCIe Gen5)", "GPU Module (PCIe Gen5)"],
        ["單價 (USD)", 285, 262, 298],
        ["MOQ (最小訂購量)", 1000, 2000, 500],
        ["Lead Time (週)", 8, 12, 6],
        ["急單能力", "可縮短至6週(+5%)", "無法急單", "可縮短至4週(+8%)"],
        ["品質評分 (1-5)", 4.2, 3.8, 4.5],
        ["財務穩定性 (A-D)", "A", "B", "A"],
        ["過去12個月不良率 (PPM)", 150, 320, 80],
        ["ISO 認證", "ISO 9001/14001", "ISO 9001", "ISO 9001/14001/IATF 16949"],
        ["出廠檢驗標準 (AQL)", 0.4, 1.0, 0.25],
        ["備用產能 (%)", 15, 30, 10],
        ["技術支援回應時間 (小時)", 4, 24, 2],
        ["在地 FAE 支援", "台灣有駐點", "無", "台灣有駐點"],
        ["合約彈性", "中", "高", "低"],
        ["地緣政治風險", "低", "高 (美中關稅)", "低"],
        ["年度價格調整趨勢", "+3~5%", "+5~8%", "+2~3%"],
        ["退換貨處理時效 (天)", 5, 15, 3],
        ["MTBF (小時)", ">200,000", ">100,000", ">250,000"],
        ["主要客戶參考", "Dell, Lenovo", "Inspur, Sugon", "HPE, Fujitsu"],
        ["綜合加權評分", "", "", ""],
    ]

    for c, h in enumerate(headers, 1):
        ws1.cell(row=1, column=c, value=h)
    style_header_row(ws1, 1, len(headers))

    for r, row_data in enumerate(data, 2):
        for c, val in enumerate(row_data, 1):
            ws1.cell(row=r, column=c, value=val)
    style_data_rows(ws1, 2, len(data) + 1, len(headers))

    # Left-align the first column
    for r in range(2, len(data) + 2):
        ws1.cell(row=r, column=1).alignment = LEFT_WRAP
        ws1.cell(row=r, column=1).font = Font(name="Microsoft JhengHei", bold=True, size=10)

    auto_width(ws1)
    ref = f"A1:D{len(data)+1}"
    add_table(ws1, ref, "GPU供應商評估")

    # ── Sheet 2: DRAM 供應商評估 ──
    ws2 = wb.create_sheet("DRAM供應商評估")
    ws2.sheet_properties.tabColor = "548235"

    headers2 = ["評估項目", "NovaTech (韓國)", "MemoryPlus (中國)", "SiliconEdge (日本)"]
    data2 = [
        ["零件型號", "NT-D5-32G", "MP-DDR5-32", "SE-DDR5-32"],
        ["產品類別", "DDR5-32GB ECC RDIMM", "DDR5-32GB ECC RDIMM", "DDR5-32GB ECC RDIMM"],
        ["速度規格", "4800 MHz", "4800 MHz", "4800 MHz"],
        ["單價 (USD)", 42, 38, 45],
        ["MOQ", 5000, 10000, 2000],
        ["Lead Time (週)", 4, 6, 3],
        ["品質評分 (1-5)", 4.0, 3.5, 4.6],
        ["財務穩定性", "A", "B", "A"],
        ["不良率 (PPM)", 200, 450, 60],
        ["ISO 認證", "ISO 9001/14001", "ISO 9001", "ISO 9001/14001"],
        ["備用產能 (%)", 20, 40, 8],
        ["技術支援回應 (小時)", 4, 48, 2],
        ["合約彈性", "中", "高", "低"],
        ["地緣政治風險", "低", "高", "低"],
        ["JEDEC 合規", "是", "是", "是"],
        ["主要客戶參考", "Samsung OEM, Dell", "Inspur, H3C", "HPE, NEC"],
        ["綜合加權評分", "", "", ""],
    ]

    for c, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header_row(ws2, 1, len(headers2), HEADER_FILL_GREEN)

    for r, row_data in enumerate(data2, 2):
        for c, val in enumerate(row_data, 1):
            ws2.cell(row=r, column=c, value=val)
    style_data_rows(ws2, 2, len(data2) + 1, len(headers2))
    for r in range(2, len(data2) + 2):
        ws2.cell(row=r, column=1).alignment = LEFT_WRAP
        ws2.cell(row=r, column=1).font = Font(name="Microsoft JhengHei", bold=True, size=10)
    auto_width(ws2)
    add_table(ws2, f"A1:D{len(data2)+1}", "DRAM供應商評估")

    # ── Sheet 3: 電源供應器供應商評估 ──
    ws3 = wb.create_sheet("電源供應器供應商評估")
    ws3.sheet_properties.tabColor = "C55A11"

    headers3 = ["評估項目", "PowerCore (台灣)", "VoltMax (中國)", "EnergyPro (日本)"]
    data3 = [
        ["零件型號", "PC-2000W", "VM-2000S", "EP-2000X"],
        ["產品類別", "2000W 冗餘電源", "2000W 冗餘電源", "2000W 冗餘電源"],
        ["80+ 認證等級", "Titanium", "Gold", "Titanium"],
        ["單價 (USD)", 165, 142, 178],
        ["MOQ", 500, 1000, 300],
        ["Lead Time (週)", 5, 8, 4],
        ["品質評分 (1-5)", 4.3, 3.6, 4.7],
        ["財務穩定性", "A", "B", "A"],
        ["不良率 (PPM)", 120, 380, 50],
        ["MTBF (小時)", ">300,000", ">150,000", ">400,000"],
        ["效率 (滿載)", "96.5%", "93.8%", "97.1%"],
        ["備用產能 (%)", 25, 35, 12],
        ["技術支援回應 (小時)", 2, 24, 4],
        ["地緣政治風險", "低", "高", "低"],
        ["安規認證", "UL/CE/BSMI/CCC", "UL/CE/CCC", "UL/CE/PSE/BSMI"],
        ["綜合加權評分", "", "", ""],
    ]

    for c, h in enumerate(headers3, 1):
        ws3.cell(row=1, column=c, value=h)
    style_header_row(ws3, 1, len(headers3), HEADER_FILL_ORANGE)

    for r, row_data in enumerate(data3, 2):
        for c, val in enumerate(row_data, 1):
            ws3.cell(row=r, column=c, value=val)
    style_data_rows(ws3, 2, len(data3) + 1, len(headers3))
    for r in range(2, len(data3) + 2):
        ws3.cell(row=r, column=1).alignment = LEFT_WRAP
        ws3.cell(row=r, column=1).font = Font(name="Microsoft JhengHei", bold=True, size=10)
    auto_width(ws3)
    add_table(ws3, f"A1:D{len(data3)+1}", "PSU供應商評估")

    # ── Sheet 4: 評估權重設定 ──
    ws4 = wb.create_sheet("評估權重設定")
    ws4.sheet_properties.tabColor = "7030A0"

    headers4 = ["評估維度", "權重 (%)", "說明"]
    data4 = [
        ["價格競爭力", 25, "單價、量級折扣、年度價格穩定性"],
        ["交期可靠性", 20, "標準交期、急單能力、歷史準時交貨率"],
        ["品質表現", 20, "不良率(PPM)、品質評分、ISO認證完整度"],
        ["財務穩定性", 10, "財務評級、營收規模、營運年數"],
        ["技術支援", 10, "回應時間、在地FAE支援、技術文件完整度"],
        ["地緣政治風險", 10, "所在國家政策風險、關稅影響、出口管制"],
        ["產能彈性", 5, "備用產能比例、MOQ靈活度、合約彈性"],
    ]

    for c, h in enumerate(headers4, 1):
        ws4.cell(row=1, column=c, value=h)
    style_header_row(ws4, 1, len(headers4), HEADER_FILL_PURPLE)

    for r, row_data in enumerate(data4, 2):
        for c, val in enumerate(row_data, 1):
            ws4.cell(row=r, column=c, value=val)
    style_data_rows(ws4, 2, len(data4) + 1, len(headers4))
    for r in range(2, len(data4) + 2):
        ws4.cell(row=r, column=1).alignment = LEFT_WRAP
        ws4.cell(row=r, column=3).alignment = LEFT_WRAP
    auto_width(ws4, max_width=50)
    add_table(ws4, f"A1:C{len(data4)+1}", "評估權重")

    path = os.path.join(OUTPUT_DIR, "英業達_供應商評估表.xlsx")
    wb.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 2: 英業達_QBR_KPI_Dashboard.xlsx
# ═══════════════════════════════════════════════════════════
def create_qbr_kpi():
    wb = openpyxl.Workbook()

    # ── Sheet 1: 季度KPI總覽 ──
    ws1 = wb.active
    ws1.title = "季度KPI總覽"
    ws1.sheet_properties.tabColor = "2F5496"

    headers = ["季度", "產品線", "出貨量(台)", "交期達成率(%)", "DPPM",
               "客訴件數", "客訴關閉率(%)", "實際成本節省(USD)",
               "目標成本節省(USD)", "成本節省達成率(%)"]
    data = [
        ["Q1 2025", "ProServer X100", 12500, 96.2, 180, 3, 100, 125000, 100000, 125.0],
        ["Q1 2025", "CloudBook L15", 45000, 98.1, 95, 1, 100, 280000, 250000, 112.0],
        ["Q1 2025", "IoT Gateway G3", 8200, 99.0, 45, 0, None, 35000, 30000, 116.7],
        ["Q2 2025", "ProServer X100", 14200, 97.5, 150, 2, 100, 142000, 150000, 94.7],
        ["Q2 2025", "CloudBook L15", 52000, 97.8, 110, 2, 100, 310000, 300000, 103.3],
        ["Q2 2025", "IoT Gateway G3", 9500, 98.5, 55, 1, 100, 42000, 40000, 105.0],
        ["Q3 2025", "ProServer X100", 13800, 94.1, 220, 5, 80, 98000, 150000, 65.3],
        ["Q3 2025", "CloudBook L15", 48000, 96.5, 130, 3, 67, 265000, 300000, 88.3],
        ["Q3 2025", "IoT Gateway G3", 10200, 97.0, 70, 1, 100, 38000, 45000, 84.4],
        ["Q4 2025", "ProServer X100", 15000, 91.3, 280, 7, 57, 85000, 150000, 56.7],
        ["Q4 2025", "CloudBook L15", 55000, 95.2, 145, 4, 75, 290000, 350000, 82.9],
        ["Q4 2025", "IoT Gateway G3", 11000, 96.8, 65, 2, 50, 45000, 50000, 90.0],
    ]

    for c, h in enumerate(headers, 1):
        ws1.cell(row=1, column=c, value=h)
    style_header_row(ws1, 1, len(headers))

    for r, row_data in enumerate(data, 2):
        for c, val in enumerate(row_data, 1):
            cell = ws1.cell(row=r, column=c, value=val)
    style_data_rows(ws1, 2, len(data) + 1, len(headers))

    # Format numbers
    for r in range(2, len(data) + 2):
        ws1.cell(row=r, column=3).number_format = '#,##0'
        ws1.cell(row=r, column=8).number_format = '#,##0'
        ws1.cell(row=r, column=9).number_format = '#,##0'

    auto_width(ws1)
    add_table(ws1, f"A1:J{len(data)+1}", "季度KPI")

    # ── Sheet 2: 品質異常明細 ──
    ws2 = wb.create_sheet("品質異常明細")
    ws2.sheet_properties.tabColor = "C00000"

    headers2 = ["編號", "發生日期", "產品線", "異常描述", "根因分類",
                "影響數量(台)", "狀態", "負責人", "預計關閉日", "實際關閉日"]
    data2 = [
        ["QA-2025-031", "2025/7/15", "ProServer X100", "電源模組過熱導致系統當機，環境溫度 35°C 時觸發 OTP", "來料不良", 350, "已關閉", "張偉", "2025/8/10", "2025/8/8"],
        ["QA-2025-033", "2025/7/28", "CloudBook L15", "LCD 面板亮點超過規格(>3點)", "來料不良", 180, "已關閉", "王美玲", "2025/8/25", "2025/8/20"],
        ["QA-2025-038", "2025/8/22", "ProServer X100", "DRAM 相容性問題造成 BSOD (特定批次 DDR5-4800)", "設計問題", 520, "已關閉", "林明哲", "2025/9/15", "2025/9/12"],
        ["QA-2025-042", "2025/9/10", "CloudBook L15", "鍵盤背光不均勻(左上角偏暗 15%)", "製程問題", 1200, "已關閉", "王美玲", "2025/10/5", "2025/10/3"],
        ["QA-2025-045", "2025/10/3", "ProServer X100", "GPU 散熱片組裝偏移 0.5mm，導致接觸不良溫度升高", "製程問題", 680, "進行中", "張偉", "2026/1/15", ""],
        ["QA-2025-048", "2025/10/18", "ProServer X100", "BMC 韌體更新後 IPMI 遠端管理異常(v2.3.1)", "設計問題", 420, "進行中", "陳志豪", "2026/1/30", ""],
        ["QA-2025-051", "2025/11/5", "CloudBook L15", "電池續航力低於規格 8%(實測 7.2hr vs 規格 8hr)", "來料不良", 2100, "進行中", "黃淑芬", "2026/2/15", ""],
        ["QA-2025-053", "2025/11/12", "IoT Gateway G3", "Wi-Fi 6E 訊號不穩定(距離>15m 斷線)", "設計問題", 320, "進行中", "陳志豪", "2026/2/28", ""],
        ["QA-2025-055", "2025/11/20", "ProServer X100", "網路介面卡 Broadcom NIC 與特定 switch 相容性問題", "設計問題", 310, "待處理", "李建宏", "2026/2/28", ""],
    ]

    for c, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header_row(ws2, 1, len(headers2), PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"))

    for r, row_data in enumerate(data2, 2):
        for c, val in enumerate(row_data, 1):
            ws2.cell(row=r, column=c, value=val)
    style_data_rows(ws2, 2, len(data2) + 1, len(headers2))

    # Color-code status
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    for r in range(2, len(data2) + 2):
        status_cell = ws2.cell(row=r, column=7)
        if status_cell.value == "待處理":
            status_cell.fill = red_fill
        elif status_cell.value == "進行中":
            status_cell.fill = yellow_fill
        elif status_cell.value == "已關閉":
            status_cell.fill = green_fill

    for r in range(2, len(data2) + 2):
        ws2.cell(row=r, column=4).alignment = LEFT_WRAP
    auto_width(ws2, max_width=40)
    add_table(ws2, f"A1:J{len(data2)+1}", "品質異常")

    # ── Sheet 3: 月度交期追蹤 ──
    ws3 = wb.create_sheet("月度交期追蹤")
    ws3.sheet_properties.tabColor = "548235"

    headers3 = ["月份", "產品線", "訂單數(台)", "準時出貨(台)", "延遲1-3天(台)",
                "延遲3天以上(台)", "交期達成率(%)"]
    data3 = [
        ["2025/07", "ProServer X100", 4800, 4420, 250, 130, 92.1],
        ["2025/08", "ProServer X100", 4500, 4250, 180, 70, 94.4],
        ["2025/09", "ProServer X100", 4500, 4300, 140, 60, 95.6],
        ["2025/10", "ProServer X100", 5200, 4600, 380, 220, 88.5],
        ["2025/11", "ProServer X100", 5000, 4550, 300, 150, 91.0],
        ["2025/12", "ProServer X100", 4800, 4500, 200, 100, 93.8],
        ["2025/07", "CloudBook L15", 16500, 15800, 500, 200, 95.8],
        ["2025/08", "CloudBook L15", 16000, 15500, 350, 150, 96.9],
        ["2025/09", "CloudBook L15", 15500, 15100, 280, 120, 97.4],
        ["2025/10", "CloudBook L15", 18000, 17100, 600, 300, 95.0],
        ["2025/11", "CloudBook L15", 19000, 18000, 700, 300, 94.7],
        ["2025/12", "CloudBook L15", 18000, 17200, 550, 250, 95.6],
        ["2025/07", "IoT Gateway G3", 2800, 2770, 20, 10, 98.9],
        ["2025/08", "IoT Gateway G3", 3000, 2950, 35, 15, 98.3],
        ["2025/09", "IoT Gateway G3", 3200, 3120, 50, 30, 97.5],
        ["2025/10", "IoT Gateway G3", 3500, 3400, 70, 30, 97.1],
        ["2025/11", "IoT Gateway G3", 3800, 3680, 80, 40, 96.8],
        ["2025/12", "IoT Gateway G3", 3700, 3600, 65, 35, 97.3],
    ]

    for c, h in enumerate(headers3, 1):
        ws3.cell(row=1, column=c, value=h)
    style_header_row(ws3, 1, len(headers3), HEADER_FILL_GREEN)

    for r, row_data in enumerate(data3, 2):
        for c, val in enumerate(row_data, 1):
            ws3.cell(row=r, column=c, value=val)
    style_data_rows(ws3, 2, len(data3) + 1, len(headers3))

    for r in range(2, len(data3) + 2):
        for c in [3, 4, 5, 6]:
            ws3.cell(row=r, column=c).number_format = '#,##0'

    auto_width(ws3)
    add_table(ws3, f"A1:G{len(data3)+1}", "月度交期")

    path = os.path.join(OUTPUT_DIR, "英業達_QBR_KPI_Dashboard.xlsx")
    wb.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 3: 英業達_BOM成本分析.xlsx
# ═══════════════════════════════════════════════════════════
def create_bom_analysis():
    wb = openpyxl.Workbook()

    # ── Sheet 1: ProServer X200 BOM ──
    ws1 = wb.active
    ws1.title = "ProServer X200 BOM"
    ws1.sheet_properties.tabColor = "2F5496"

    headers = ["項次", "零件類別", "零件名稱", "規格", "供應商",
               "單價(USD)", "數量", "小計(USD)", "佔BOM比例(%)"]
    data = [
        [1, "CPU", "Intel Xeon W5-3435X", "16C/32T, 4.7GHz", "Intel", 1250, 2, 2500, 32.5],
        [2, "GPU Module", "NT-500", "PCIe Gen5, 400W TDP", "NovaTech", 285, 4, 1140, 14.8],
        [3, "DRAM", "DDR5-32G ECC RDIMM", "4800MHz ECC Registered", "NovaTech", 42, 16, 672, 8.7],
        [4, "SSD", "NVMe 1.6TB", "PCIe Gen4, 7000MB/s", "Samsung", 185, 4, 740, 9.6],
        [5, "主機板", "自製 MB-X200", "Dual Socket, ATX-E", "英業達(自製)", 320, 1, 320, 4.2],
        [6, "電源供應器", "PC-2000W", "80+ Titanium 冗餘", "PowerCore", 165, 2, 330, 4.3],
        [7, "散熱模組", "液冷散熱 Cold Plate", "400W TDP Direct Contact", "CoolTech", 95, 2, 190, 2.5],
        [8, "CDU", "Coolant Distribution Unit", "機架式 CDU", "CoolTech", 280, 0.25, 70, 0.9],
        [9, "機殼", "2U Rackmount Chassis", "標準19吋 720mm深", "英業達(自製)", 85, 1, 85, 1.1],
        [10, "風扇模組", "80mm Hot-swap Fan", "4-pin PWM, 12000RPM", "英業達(自製)", 12, 8, 96, 1.2],
        [11, "網路介面卡", "25GbE Dual Port NIC", "SFP28, PCIe Gen4", "Mellanox/NVIDIA", 180, 2, 360, 4.7],
        [12, "RAID 控制器", "HW RAID Controller", "SAS 12Gb/s, 4GB Cache", "Broadcom", 210, 1, 210, 2.7],
        [13, "背板", "SAS/SATA Backplane", "支援8顆3.5吋/2.5吋", "英業達(自製)", 45, 1, 45, 0.6],
        [14, "Riser Card", "PCIe Riser Card", "x16 Gen5, Dual Slot", "英業達(自製)", 35, 2, 70, 0.9],
        [15, "線材/連接器", "各式線材", "電源線/SATA線/訊號線", "多家", 55, 1, 55, 0.7],
        [16, "包裝材料", "伺服器包裝", "防震泡棉/紙箱/棧板", "英業達(自製)", 35, 1, 35, 0.5],
        [17, "其他", "標籤/螺絲/附件", "導軌/理線架/說明書", "多家", 15, 1, 15, 0.2],
    ]

    for c, h in enumerate(headers, 1):
        ws1.cell(row=1, column=c, value=h)
    style_header_row(ws1, 1, len(headers))

    for r, row_data in enumerate(data, 2):
        for c, val in enumerate(row_data, 1):
            ws1.cell(row=r, column=c, value=val)
    style_data_rows(ws1, 2, len(data) + 1, len(headers))

    # Total row
    total_row = len(data) + 2
    ws1.cell(row=total_row, column=1, value="")
    ws1.cell(row=total_row, column=2, value="合計")
    ws1.cell(row=total_row, column=8, value=6933)
    ws1.cell(row=total_row, column=9, value=100)
    for c in range(1, len(headers) + 1):
        cell = ws1.cell(row=total_row, column=c)
        cell.font = Font(name="Microsoft JhengHei", bold=True, size=11)
        cell.border = THIN_BORDER
        cell.alignment = CENTER

    for r in range(2, total_row + 1):
        ws1.cell(row=r, column=6).number_format = '#,##0'
        ws1.cell(row=r, column=8).number_format = '#,##0'

    auto_width(ws1)
    add_table(ws1, f"A1:I{len(data)+1}", "BOM明細")

    # ── Sheet 2: GPU 替代方案成本比較 ──
    ws2 = wb.create_sheet("GPU替代方案成本比較")
    ws2.sheet_properties.tabColor = "C55A11"

    headers2 = ["比較項目", "現況\n(NovaTech)", "方案A\n(ChipMax)", "方案B\n(SiliconEdge)", "方案C\n(雙供應商策略)"]
    data2 = [
        ["GPU 供應商", "NovaTech (韓國)", "ChipMax (中國)", "SiliconEdge (日本)", "NovaTech + ChipMax"],
        ["GPU 型號", "NT-500", "CM-G5", "SE-Pro", "NT-500 + CM-G5"],
        ["GPU 單價 (USD)", 285, 262, 298, "285 / 262 混合"],
        ["每台使用數量", 4, 4, 4, "4 (2+2)"],
        ["每台 GPU 成本 (USD)", 1140, 1048, 1192, 1094],
        ["與現況差異 (每台, USD)", "基準", -92, 52, -46],
        ["年產量預估 (台)", 60000, 60000, 60000, 60000],
        ["年度成本差異 (USD)", "基準", -5520000, 3120000, -2760000],
        ["換線驗證成本 (一次性, USD)", "N/A", 350000, 280000, 450000],
        ["驗證週期 (週)", "N/A", 8, 6, 10],
        ["換線對產出的影響", "無", "停線2天", "停線1天", "停線3天"],
        ["品質風險評估", "低", "中高 (PPM 320)", "低 (PPM 80)", "中"],
        ["供應穩定性風險", "低", "高 (地緣政治)", "低", "低 (分散風險)"],
        ["技術支援品質", "良好", "一般", "優秀", "良好"],
        ["12個月 TCO (USD)", "基準", "-5,170,000", "+2,840,000", "-2,310,000"],
        ["建議優先順序", "維持現狀", "第三選擇", "第二選擇", "推薦方案 ★"],
    ]

    for c, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header_row(ws2, 1, len(headers2), HEADER_FILL_ORANGE)

    for r, row_data in enumerate(data2, 2):
        for c, val in enumerate(row_data, 1):
            ws2.cell(row=r, column=c, value=val)
    style_data_rows(ws2, 2, len(data2) + 1, len(headers2))

    for r in range(2, len(data2) + 2):
        ws2.cell(row=r, column=1).alignment = LEFT_WRAP
        ws2.cell(row=r, column=1).font = Font(name="Microsoft JhengHei", bold=True, size=10)

    auto_width(ws2, max_width=25)
    add_table(ws2, f"A1:E{len(data2)+1}", "替代方案比較")

    path = os.path.join(OUTPUT_DIR, "英業達_BOM成本分析.xlsx")
    wb.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 4: 英業達_NPI_ProServerX200_專案追蹤.xlsx
# ═══════════════════════════════════════════════════════════
def create_npi_tracker():
    wb = openpyxl.Workbook()

    # ── Sheet 1: 專案里程碑 ──
    ws1 = wb.active
    ws1.title = "專案里程碑"
    ws1.sheet_properties.tabColor = "2F5496"

    headers = ["階段", "里程碑", "計畫日期", "實際日期", "狀態", "負責人", "備註"]
    data = [
        ["Planning", "客戶需求確認 (Requirements Freeze)", "2025/3/15", "2025/3/15", "✅ 完成", "James Chen (PM)", "需求書 v2.1 已由 HP 簽核"],
        ["Planning", "NPI Kick-off Meeting", "2025/4/1", "2025/4/1", "✅ 完成", "James Chen (PM)", "全員參加，會議紀錄已發布"],
        ["Design", "散熱方案確認 (液冷)", "2025/4/11", "2025/4/10", "✅ 完成", "Kevin Liu (R&D)", "選定直接接觸液冷+CDU方案"],
        ["Design", "機構設計完成", "2025/4/30", "2025/5/5", "✅ 完成", "Kevin Liu (R&D)", "延遲5天，散熱方案變更"],
        ["Design", "電路設計完成", "2025/5/15", "2025/5/15", "✅ 完成", "Kevin Liu (R&D)", "含液冷介面整合"],
        ["Procurement", "供應商評估報告", "2025/4/4", "2025/4/4", "✅ 完成", "Lisa Wang (SCM)", "GPU: NovaTech優先"],
        ["Procurement", "BOM 鎖定", "2025/5/1", "2025/5/8", "✅ 完成", "Lisa Wang (SCM)", "GPU供應商延遲確認"],
        ["Procurement", "關鍵料備齊", "2025/6/15", "2025/6/28", "✅ 完成", "Lisa Wang (SCM)", "DRAM短缺延遲13天"],
        ["EVT", "EVT 樣品組裝 (10台)", "2025/6/30", "2025/7/5", "✅ 完成", "Kevin Liu (R&D)", ""],
        ["EVT", "EVT 測試完成", "2025/7/31", "2025/8/10", "✅ 完成", "Amy Hsu (QA)", "散熱不過需ECO，延遲10天"],
        ["DVT", "DVT 樣品組裝 (50台)", "2025/8/15", "", "🔄 進行中", "Kevin Liu (R&D)", "組裝中"],
        ["DVT", "DVT 測試完成", "2025/9/30", "", "⬜ 待開始", "Amy Hsu (QA)", "含漏液/震動/高溫高濕測試"],
        ["PVT", "PVT 驗證 (200台)", "2025/10/15", "", "⬜ 待開始", "Amy Hsu (QA)", ""],
        ["MP", "量產準備完成 (MP Readiness)", "2025/11/1", "", "⬜ 待開始", "James Chen (PM)", "含 MES 設定/SOP 完成"],
        ["MP", "正式量產 (Mass Production)", "2025/11/15", "", "⬜ 待開始", "All", "客戶要求 Q4 2025 量產"],
    ]

    for c, h in enumerate(headers, 1):
        ws1.cell(row=1, column=c, value=h)
    style_header_row(ws1, 1, len(headers))

    for r, row_data in enumerate(data, 2):
        for c, val in enumerate(row_data, 1):
            ws1.cell(row=r, column=c, value=val)
    style_data_rows(ws1, 2, len(data) + 1, len(headers))

    for r in range(2, len(data) + 2):
        ws1.cell(row=r, column=2).alignment = LEFT_WRAP
        ws1.cell(row=r, column=7).alignment = LEFT_WRAP

    auto_width(ws1, max_width=40)
    add_table(ws1, f"A1:G{len(data)+1}", "里程碑")

    # ── Sheet 2: 每週進度 ──
    ws2 = wb.create_sheet("每週進度")
    ws2.sheet_properties.tabColor = "548235"

    headers2 = ["週次", "日期區間", "完成項目", "進行中項目", "風險/議題", "下週計畫"]
    data2 = [
        ["W14", "4/1-4/5", "• Kick-off 完成\n• 專案章程核准", "• 機構設計\n• BOM 初版", "散熱設計 TDP 超出預期，需評估液冷方案", "散熱方案評估"],
        ["W15", "4/7-4/11", "• 散熱方案選定(液冷)\n• 熱模擬完成 72°C", "• 機構修改圖面\n• 供應商洽談", "GPU 供應商報價中\nChipMax 交期 12 週偏長", "完成機構修改圖面"],
        ["W16", "4/14-4/18", "• GPU 供應商評估報告完成\n• 初版 BOM 完成\n• Teams 頻道建立", "• 電路設計整合液冷介面\n• CoolTech 供應商導入\n• ERP 料號建檔\n• DVT 測試計畫制定", "Lead time 長(ChipMax 12週)\nNovaTech Q3 產能預訂 85%\nDRAM 價格波動 ±15%", "BOM 鎖定\nERP 建檔完成\n供應商合約簽訂"],
        ["W17", "4/21-4/25", "", "• 電路設計\n• CoolTech 導入\n• DVT 測試計畫", "NovaTech 報價等總部核准", "完成 CoolTech 供應商導入"],
        ["W18", "4/28-5/2", "", "• 電路設計收尾\n• BOM 鎖定", "客戶詢問 EVT 樣品時程", "電路設計完成\nBOM 正式鎖定"],
    ]

    for c, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header_row(ws2, 1, len(headers2), HEADER_FILL_GREEN)

    for r, row_data in enumerate(data2, 2):
        for c, val in enumerate(row_data, 1):
            ws2.cell(row=r, column=c, value=val)
    style_data_rows(ws2, 2, len(data2) + 1, len(headers2))

    for r in range(2, len(data2) + 2):
        for c in range(3, 7):
            ws2.cell(row=r, column=c).alignment = LEFT_WRAP

    auto_width(ws2, max_width=45)
    ws2.row_dimensions[2].height = 40
    ws2.row_dimensions[3].height = 40
    ws2.row_dimensions[4].height = 60
    ws2.row_dimensions[5].height = 40
    ws2.row_dimensions[6].height = 40
    add_table(ws2, f"A1:F{len(data2)+1}", "每週進度")

    # ── Sheet 3: 風險登記簿 ──
    ws3 = wb.create_sheet("風險登記簿")
    ws3.sheet_properties.tabColor = "C00000"

    headers3 = ["風險編號", "風險描述", "影響層面", "發生機率", "影響程度",
                "風險等級", "緩解措施", "負責人", "狀態", "更新日期"]
    data3 = [
        ["R-001", "GPU供應商(NovaTech)交期延遲\nQ3產能預訂85%，可能無法按時供貨", "時程", "高", "高", "🔴 嚴重", "1. 備選供應商SiliconEdge已評估\n2. 提前下單鎖定產能\n3. 簽訂產能保證書", "Lisa Wang", "監控中", "2025/4/18"],
        ["R-002", "散熱方案未通過驗證\n液冷方案技術複雜度較高", "時程/品質", "中", "高", "🟡 高", "1. 備用方案(氣冷+增加風扇)\n2. 已完成熱模擬驗證(72°C通過)", "Kevin Liu", "✅ 已緩解", "2025/4/15"],
        ["R-003", "DRAM價格上漲20%\n市場供需緊張", "成本", "中", "中", "🟡 中", "1. 提前鎖定半年用量\n2. 與SiliconEdge洽談長約折扣\n3. 評估替代規格", "Lisa Wang", "監控中", "2025/4/18"],
        ["R-004", "客戶(HP)變更規格\n設計凍結後仍有 CR 風險", "全面", "低", "高", "🟡 中", "1. 定期確認凍結規格\n2. CR 流程需客戶簽核\n3. 影響評估 SOP", "James Chen", "監控中", "2025/4/1"],
        ["R-005", "測試設備產能不足\nSGS實驗室檔期緊張", "時程", "中", "中", "🟡 中", "1. 已預約SGS 8月第二週\n2. 備選UL實驗室\n3. 評估購置部分設備", "Amy Hsu", "進行中", "2025/4/15"],
        ["R-006", "美中關稅政策變化\n影響中國供應商料件成本", "成本/供應", "中", "高", "🟡 高", "1. ChipMax 已列為第三備選\n2. 優先使用韓國/日本供應商\n3. 持續監控政策動態", "Lisa Wang", "監控中", "2025/4/18"],
        ["R-007", "量產時程壓縮\n客戶要求Q4量產，僅7個月", "時程", "高", "中", "🟡 高", "1. 平行作業(設計/採購同步)\n2. 加速EVT/DVT時程\n3. 每兩週進度追蹤", "James Chen", "監控中", "2025/4/1"],
    ]

    for c, h in enumerate(headers3, 1):
        ws3.cell(row=1, column=c, value=h)
    style_header_row(ws3, 1, len(headers3), PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"))

    for r, row_data in enumerate(data3, 2):
        for c, val in enumerate(row_data, 1):
            ws3.cell(row=r, column=c, value=val)
    style_data_rows(ws3, 2, len(data3) + 1, len(headers3))

    for r in range(2, len(data3) + 2):
        ws3.cell(row=r, column=2).alignment = LEFT_WRAP
        ws3.cell(row=r, column=7).alignment = LEFT_WRAP
        ws3.row_dimensions[r].height = 55

    auto_width(ws3, max_width=40)
    add_table(ws3, f"A1:J{len(data3)+1}", "風險登記")

    path = os.path.join(OUTPUT_DIR, "英業達_NPI_ProServerX200_專案追蹤.xlsx")
    wb.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# Run all
# ═══════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  AB-730 英業達 — 建立 Excel 範例檔案")
    print("=" * 60)
    create_supplier_evaluation()
    create_qbr_kpi()
    create_bom_analysis()
    create_npi_tracker()
    print("=" * 60)
    print("  All Excel files created successfully!")
    print("=" * 60)
