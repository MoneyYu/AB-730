# -*- coding: utf-8 -*-
"""
AB-730 英業達課程 — 建立會議總結範本 Word 檔
"""
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import os

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

def set_cell_shading(cell, color):
    shading_elm = cell._tc.get_or_add_tcPr()
    shading = shading_elm.makeelement(qn('w:shd'), {
        qn('w:fill'): color, qn('w:val'): 'clear',
    })
    shading_elm.append(shading)

def add_styled_table(doc, headers, rows, header_color="2F5496"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_shading(cell, header_color)
    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(9)
    return table

def create_meeting_summary_template():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    # ═══ Title ═══
    title = doc.add_heading('', level=0)
    run = title.add_run('英業達股份有限公司')
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(47, 84, 150)

    doc.add_heading('會議總結範本 — 多角度摘要指南', level=1)

    p = doc.add_paragraph()
    run = p.add_run('本文件提供同一場會議的四種不同角度摘要範本。')
    run.font.size = Pt(11)
    p = doc.add_paragraph()
    run = p.add_run('會議來源：ProServer X200 設計審查會議 (2025/4/15)')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(128, 128, 128)

    doc.add_paragraph()

    # ═══════════════════════════════════════
    # Version 1: CEO / VP
    # ═══════════════════════════════════════
    doc.add_page_break()
    h = doc.add_heading('版本一：CEO / VP 角度', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(192, 0, 0)

    p = doc.add_paragraph()
    run = p.add_run('📌 設計原則：30 秒看完、只有結論和決策、不含技術細節')
    run.bold = True
    run.font.color.rgb = RGBColor(192, 0, 0)

    doc.add_paragraph()
    doc.add_heading('ProServer X200 設計審查 — 執行摘要', level=2)

    info = [
        ("日期", "2025/4/15"),
        ("狀態", "🟡 黃燈 — 進度可控但有供應鏈風險"),
    ]
    for label, val in info:
        p = doc.add_paragraph()
        run1 = p.add_run(f"{label}：")
        run1.bold = True
        run2 = p.add_run(val)

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('一句話結論：')
    run.bold = True
    run.font.size = Pt(12)
    doc.add_paragraph('散熱方案（液冷）通過設計審查，技術風險已排除。但 GPU 主供應商 Q3 產能緊張，建議啟動備援供應商驗證。')

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('⚡ 需要您決策的事項：')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(192, 0, 0)

    decisions = [
        "1. 核可 SiliconEdge（日本）備援供應商驗證費用 USD 45,000",
        "2. 核可液冷模組供應商選用 CoolTech（台灣，$95/個）而非 ThermalPro（中國，$72/個）",
        "   → 年化成本差異約 USD 1.38M，但品質風險和地緣政治風險顯著降低",
    ]
    for d in decisions:
        doc.add_paragraph(d)

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('💰 成本影響：')
    run.bold = True
    doc.add_paragraph('若採用雙供應商策略（NovaTech 60% + SiliconEdge 40%），年度 GPU 成本增加約 USD 780K，但供應穩定性從「高風險」降至「低風險」。')

    p = doc.add_paragraph()
    run = p.add_run('⏰ 下一個關鍵里程碑：')
    run.bold = True
    doc.add_paragraph('2025/4/30 — DVT 測試計畫完成')

    # ═══════════════════════════════════════
    # Version 2: PM
    # ═══════════════════════════════════════
    doc.add_page_break()
    h = doc.add_heading('版本二：PM（專案經理）角度', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(47, 84, 150)

    p = doc.add_paragraph()
    run = p.add_run('📌 設計原則：完整追蹤、每個細節都記錄、含相依關係和風險')
    run.bold = True
    run.font.color.rgb = RGBColor(47, 84, 150)

    doc.add_paragraph()
    doc.add_heading('一、決議事項', level=2)

    decisions_pm = [
        ("TD-001", "散熱方案確認", "採用直接接觸液冷板 + CDU 架構", "熱模擬結果 72°C 通過（目標 <85°C）", "R&D 後續設計基於液冷"),
        ("TD-002", "液冷供應商", "選用 CoolTech（台灣）", "品質優、交期短、避免地緣政治風險集中", "SCM 啟動供應商導入"),
        ("TD-003", "GPU 策略", "NovaTech 為主 + SiliconEdge 備援", "分散風險、增加議價籌碼", "需管理層核可驗證費用"),
        ("TD-004", "IO 面板", "後面板微調配合液冷管路", "改動不大，不影響電路設計", "Kevin 直接處理"),
    ]
    add_styled_table(doc,
        ["編號", "決議項目", "決議內容", "決議原因", "後續影響"],
        decisions_pm, "2F5496"
    )

    doc.add_paragraph()
    doc.add_heading('二、行動項目追蹤', level=2)

    actions_pm = [
        ("A-010", "Kevin Liu + David Wu", "電路設計整合液冷介面", "2025/4/30", "A-001 已完成", "🟡 中"),
        ("A-011", "Lisa Wang", "CoolTech 供應商導入流程", "2025/4/25", "無", "🟡 中"),
        ("A-012", "Lisa Wang", "NovaTech GPU 合約最終條款", "2025/4/25", "等 NovaTech 總部核價", "🟡 中"),
        ("A-013", "Amy Hsu", "DVT 測試計畫（含液冷測試）", "2025/4/30", "需 A-010 完成", "🔴 高"),
        ("A-014", "Tom Lin", "ERP 料號建檔完成", "2025/4/17", "需初版 BOM（已完成）", "🟢 低"),
        ("A-015", "Sarah Chang", "專案追蹤表更新", "2025/4/16", "無", "🟢 低"),
    ]
    add_styled_table(doc,
        ["編號", "負責人", "行動項目", "截止日", "相依項目", "風險"],
        actions_pm, "2F5496"
    )

    doc.add_paragraph()
    doc.add_heading('三、風險更新', level=2)

    risks_pm = [
        ("R-001", "GPU 供應商交期", "🔴→🔴", "維持嚴重", "NovaTech Q3 產能 85%，交期可能延長"),
        ("R-002", "散熱方案", "🔴→🟢", "已緩解 ✅", "液冷方案通過設計審查"),
        ("R-006", "美中關稅", "🟡→🟡", "維持中等", "ChipMax 列為第三備選"),
        ("新增", "液冷漏液風險", "新增 🟡", "中等", "DVT 需增加漏液/震動/高溫高濕測試"),
    ]
    add_styled_table(doc,
        ["編號", "風險描述", "等級變化", "目前狀態", "說明"],
        risks_pm, "C00000"
    )

    doc.add_paragraph()
    doc.add_heading('四、時程影響評估', level=2)
    doc.add_paragraph('本次設計審查未產生額外延遲。整體時程仍落後 5 天（機構設計延遲已吸收），量產目標 2025/11/15 維持不變。')
    doc.add_paragraph('⚠️ 關鍵路徑：電路設計(4/30) → BOM 鎖定(5/8) → 關鍵料備齊(6/15) → EVT(7/31) → DVT(9/30) → MP(11/15)')

    doc.add_paragraph()
    doc.add_heading('五、下次會議', level=2)
    doc.add_paragraph('日期：2025/4/29 (二) 10:00-11:30')
    doc.add_paragraph('重點議程：電路設計進度、CoolTech 導入狀態、DVT 測試計畫審查')

    # ═══════════════════════════════════════
    # Version 3: Individual (Amy Hsu)
    # ═══════════════════════════════════════
    doc.add_page_break()
    h = doc.add_heading('版本三：個人角度（以 QA Amy Hsu 為例）', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(84, 130, 53)

    p = doc.add_paragraph()
    run = p.add_run('📌 設計原則：只保留「跟我有關的」、清楚告訴我「我要做什麼」')
    run.bold = True
    run.font.color.rgb = RGBColor(84, 130, 53)

    doc.add_paragraph()
    doc.add_heading('📋 ProServer X200 設計審查 — Amy 的會議摘要', level=2)

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('跟我有關的決議：')
    run.bold = True
    run.font.size = Pt(12)
    items = [
        "散熱方案改為液冷 → 我的 DVT 測試計畫需要增加液冷相關測試項目",
        "液冷供應商選定 CoolTech → 需確認 CoolTech 的品質認證文件是否符合我們的 IQC 標準",
    ]
    for item in items:
        doc.add_paragraph(f"• {item}")

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('✅ 我需要做的事：')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(192, 0, 0)

    my_actions = [
        ("A-013", "完成 DVT 測試計畫（含液冷測試項目）", "2025/4/30", "🔴 高優先"),
        ("新增", "新增 5 項液冷相關測試：漏液(96hr)、高溫高濕(500hr)、震動、耐壓、溫度循環", "納入 A-013", ""),
        ("維持", "外部實驗室預約確認（SGS 8月第二週）", "已完成 ✅", ""),
        ("新增", "確認 UL 實驗室作為備案", "2025/4/25", ""),
    ]
    add_styled_table(doc,
        ["編號", "待辦事項", "截止日", "優先順序"],
        my_actions, "548235"
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('📎 我需要別人給我的東西：')
    run.bold = True
    run.font.size = Pt(12)
    needs = [
        "Kevin Liu → 液冷模組的完整設計規格書（含管路壓力、流量參數）→ 我才能制定測試標準",
        "Lisa Wang → CoolTech 的品質認證文件和可靠度報告",
        "Kevin Liu → 電路設計完成後的整機 BOM（最終版）→ 我才能完成 DVT 測試項目清單",
    ]
    for n in needs:
        doc.add_paragraph(f"• {n}")

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('⚠️ 影響我的風險：')
    run.bold = True
    doc.add_paragraph('• EVT 如果一次不過 → ECO + 重新驗證至少 2 週 → 壓縮我的 DVT 時間')
    doc.add_paragraph('• 液冷漏液測試如果不過 → 修改模具可能要 4 週 → 嚴重影響時程')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('📅 我的重要日期：')
    run.bold = True
    dates = [
        "4/17 — Tom 完成 ERP 建檔（跟我無直接關係，但 BOM 料號要對）",
        "4/25 — 確認 UL 實驗室備案",
        "4/30 — DVT 測試計畫完成 ⭐（我的最重要截止日）",
        "8月第二週 — SGS 實驗室開始 DVT 測試",
    ]
    for d in dates:
        doc.add_paragraph(f"• {d}")

    # ═══════════════════════════════════════
    # Version 4: Client (HP)
    # ═══════════════════════════════════════
    doc.add_page_break()
    h = doc.add_heading('版本四：客戶報告角度（HP Enterprise）', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(197, 90, 17)

    p = doc.add_paragraph()
    run = p.add_run('📌 設計原則：過濾內部資訊、不含成本/供應商名稱/內部爭論、正面專業')
    run.bold = True
    run.font.color.rgb = RGBColor(197, 90, 17)

    p = doc.add_paragraph()
    run = p.add_run('⚠️ 使用前必須人工審核：確認沒有洩漏供應商報價、BOM 成本、內部風險評估等敏感資訊')
    run.bold = True
    run.font.color.rgb = RGBColor(192, 0, 0)

    doc.add_paragraph()
    doc.add_heading('ProServer X200 — Development Progress Update', level=2)

    p = doc.add_paragraph()
    run = p.add_run('Date: ')
    run.bold = True
    p.add_run('April 15, 2025')
    p = doc.add_paragraph()
    run = p.add_run('From: ')
    run.bold = True
    p.add_run('Inventec Corporation — ProServer X200 Project Team')
    p = doc.add_paragraph()
    run = p.add_run('To: ')
    run.bold = True
    p.add_run('HP Enterprise — Server & Networking BU')

    doc.add_paragraph()
    doc.add_heading('Executive Summary', level=3)
    doc.add_paragraph('The ProServer X200 development is progressing well. We have completed the design review for the thermal solution, confirming a liquid cooling architecture that meets the 400W TDP requirement with significant thermal margin. All key design milestones are on track for the Q4 2025 mass production target.')

    doc.add_heading('Design Review Highlights', level=3)
    items_en = [
        "Thermal Solution: Liquid cooling design validated through thermal simulation. GPU junction temperature measured at 72°C under full load, well within the 85°C specification limit.",
        "Mechanical Design: Chassis layout finalized. Liquid cooling integration does not increase the 2U form factor dimensions.",
        "Electrical Design: On track for completion by May 15, 2025.",
        "Quality Planning: DVT test plan under development, incorporating comprehensive reliability tests including thermal cycling, vibration, and humidity exposure.",
    ]
    for item in items_en:
        doc.add_paragraph(f"• {item}")

    doc.add_heading('Milestone Status', level=3)
    milestones = [
        ("Thermal Design Review", "April 15", "✅ Completed"),
        ("Electrical Design Complete", "May 15", "On Track"),
        ("BOM Freeze", "May 8", "On Track"),
        ("EVT Build (10 units)", "July 5", "On Track"),
        ("DVT Complete", "September 30", "On Track"),
        ("Mass Production", "November 15", "On Track"),
    ]
    add_styled_table(doc,
        ["Milestone", "Target Date", "Status"],
        milestones, "C55A11"
    )

    doc.add_paragraph()
    doc.add_heading('Next Steps', level=3)
    next_steps = [
        "Complete electrical design integration with liquid cooling interface (April 30)",
        "Finalize DVT test plan including liquid cooling reliability tests (April 30)",
        "Begin EVT sample build upon BOM freeze (target: early July)",
        "Next progress update: May 1, 2025",
    ]
    for ns in next_steps:
        doc.add_paragraph(f"• {ns}")

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('Note: ')
    run.bold = True
    run.font.color.rgb = RGBColor(128, 128, 128)
    p.add_run('This update is for HP Enterprise internal reference. Please contact James Chen (james.chen@inventec.com) for any questions or additional details.')

    # ═══════════════════════════════════════
    # Copilot Prompt Guide
    # ═══════════════════════════════════════
    doc.add_page_break()
    h = doc.add_heading('附錄：Copilot 提示範例（產生各版本摘要）', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(47, 84, 150)

    doc.add_paragraph()
    doc.add_heading('CEO/VP 版本提示', level=2)
    doc.add_paragraph("""請用「製造副總」的角度摘要這場會議。他只有 30 秒讀這份摘要。

格式：
🚦 整體狀態：🟢🟡🔴（一個燈號）
📌 一句話結論：（20 字以內）
⚡ 需要您決策的事項：（最多 2 項，每項一句話）
💰 成本影響：（一個數字）
⏰ 下一個關鍵里程碑：（日期 + 內容）

不要細節，不要技術術語，只要結論和行動。""")

    doc.add_heading('PM 版本提示', level=2)
    doc.add_paragraph("""請用「專案經理」的角度摘要這場會議。我需要追蹤每一個細節。

格式：
一、決議事項（每項含：決議內容 / 決議原因 / 影響範圍）
二、行動項目表格：編號 | 負責人 | 行動項目 | 截止日 | 相依項目 | 風險等級
三、風險更新（新增/升級/已緩解的風險）
四、時程影響評估（是否影響量產日期？）
五、下次會議議程預覽""")

    doc.add_heading('個人版本提示（以 QA Amy 為例）', level=2)
    doc.add_paragraph("""請用「QA 工程師 Amy Hsu」的角度摘要這場會議。只保留跟 Amy 有關的部分。

格式：
📋 跟我有關的決議：
✅ 我需要做的事：（行動項目 + 截止日）
📎 我需要的資源/input：（誰要給我什麼、什麼時候）
⚠️ 需要注意的風險：
📅 重要日期：""")

    doc.add_heading('客戶版本提示', level=2)
    doc.add_paragraph("""請產生一份可以分享給客戶 HP Enterprise 的會議進度更新。

注意以下原則：
❌ 不可包含：內部成本數字、供應商報價、BOM 細節、內部爭論
✅ 可以包含：設計進度、測試計畫、里程碑達成狀態、下一步

格式：
ProServer X200 — Development Progress Update
Date: [日期]
Summary: [1 段英文摘要]
Key Milestones: [表格]
Next Steps: [列點]

語氣：專業自信，展現進度順利。英文撰寫。""")

    doc.add_heading('自訂角度提示模板', level=2)
    doc.add_paragraph("""請用「[角色名稱]」的角度摘要這場會議。

[角色]關心的重點是：[列出 2-3 個重點]
[角色]不需要知道的是：[列出要過濾的內容]

格式：[指定格式]
語氣：[指定語氣]
語言：[中文/英文]
長度：[字數限制]""")

    path = os.path.join(OUTPUT_DIR, "英業達_會議總結範本.docx")
    doc.save(path)
    print(f"✅ Created: {path}")

if __name__ == "__main__":
    create_meeting_summary_template()
