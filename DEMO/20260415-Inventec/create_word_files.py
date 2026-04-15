# -*- coding: utf-8 -*-
"""
AB-730 英業達課程 — 建立所有 Word 範例檔案
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import os

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

def set_cell_shading(cell, color):
    shading_elm = cell._tc.get_or_add_tcPr()
    shading = shading_elm.makeelement(qn('w:shd'), {
        qn('w:fill'): color,
        qn('w:val'): 'clear',
    })
    shading_elm.append(shading)

def add_styled_table(doc, headers, rows, header_color="2F5496"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
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

    # Data
    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(9)

    return table


# ═══════════════════════════════════════════════════════════
# 檔案 5: 英業達_NPI_Kickoff_會議逐字稿.docx
# ═══════════════════════════════════════════════════════════
def create_npi_kickoff_transcript():
    doc = Document()

    # Title
    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('英業達股份有限公司')
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(47, 84, 150)

    sub = doc.add_heading('ProServer X200 NPI Kick-off 會議紀錄（逐字稿）', level=1)
    for run in sub.runs:
        run.font.color.rgb = RGBColor(47, 84, 150)

    # Meeting info
    info_lines = [
        ("日期", "2025年4月1日（星期二）10:00-11:30"),
        ("地點", "桃園總部 3F 會議室 A / Microsoft Teams 線上同步"),
        ("會議類型", "NPI Kick-off Meeting"),
        ("專案代號", "PSX200"),
        ("客戶", "HP Enterprise"),
        ("主席", "James Chen (資深專案經理)"),
    ]

    for label, val in info_lines:
        p = doc.add_paragraph()
        run1 = p.add_run(f"{label}：")
        run1.bold = True
        run1.font.size = Pt(11)
        run2 = p.add_run(val)
        run2.font.size = Pt(11)

    doc.add_paragraph()
    doc.add_heading('出席者', level=2)

    attendees = [
        ("James Chen", "PM", "專案管理部", "資深專案經理"),
        ("Lisa Wang", "SCM", "供應鏈管理部", "資深供應鏈經理"),
        ("Kevin Liu", "R&D", "研發設計部", "硬體設計主管"),
        ("Amy Hsu", "QA", "品質保證部", "品質工程經理"),
        ("Tom Lin", "IT", "資訊系統部", "IT 系統工程師"),
        ("Sarah Chang", "BPM", "流程管理部", "流程改善專員"),
    ]

    add_styled_table(doc,
        ["姓名", "部門代碼", "部門", "職稱"],
        attendees
    )

    doc.add_paragraph()
    doc.add_heading('會議逐字紀錄', level=2)

    transcript = [
        ("10:02", "James Chen", """各位早安，感謝大家撥冗參加今天的 ProServer X200 NPI Kick-off。先說明一下背景，我們接到 HP Enterprise 的新一代 AI 推論伺服器訂單，型號 ProServer X200。客戶要求在 Q4 2025 進入量產，也就是說我們只有大約 7 個月的時間。時程非常緊，所以今天要確認各部門分工和關鍵里程碑。"""),

        ("10:05", "James Chen", """先看一下客戶的核心規格需求。這是一台 2U AI 推論伺服器，搭載雙 Intel Xeon W5-3435X 處理器、4 張 GPU 加速卡、512GB DDR5 ECC 記憶體、2000W 鈦金級冗餘電源。客戶特別要求散熱必須支援 400W TDP 的 GPU，這個是技術挑戰。Kevin，你先說一下研發的狀況？"""),

        ("10:08", "Kevin Liu", """好的。機構設計目前完成大約 80%，主要是 2U 機殼和主機板佈局已經定稿。但散熱是最大的問題——400W TDP 的 GPU 在 2U 空間裡用傳統氣冷方案很勉強。我們上週做了熱模擬，氣冷方案在滿載時 GPU junction temperature 會到 95 度，超過客戶要求的 85 度上限。所以我建議改用液冷散熱方案，但這需要重新設計散熱模組的管路佈局，預計需要額外 3 週的時間。"""),

        ("10:12", "James Chen", """3 週的延遲我們承受得起嗎？本來機構設計預計 4 月底完成。"""),

        ("10:13", "Kevin Liu", """如果 4 月 7 號開始改，最快 4 月 25 號可以出修改圖面，5 月 5 號之前完成驗證。電路設計不受影響，5 月 15 號可以完成。"""),

        ("10:14", "James Chen", """OK，那機構延 5 天還算可控。Lisa，供應鏈那邊的狀況呢？"""),

        ("10:15", "Lisa Wang", """報告一下幾個關鍵零件的供應商狀況。GPU 模組是這次最關鍵的零件，目前有三家候選供應商在評估：

第一家是韓國的 NovaTech，型號 NT-500，單價 285 美元，MOQ 1000，交期 8 週。品質評分 4.2 分，財務穩定性 A 級。

第二家是中國的 ChipMax，型號 CM-G5，單價 262 美元比較便宜，但 MOQ 要 2000，交期 12 週比較長。品質評分只有 3.8，財務穩定性 B 級。另外考量到目前的美中關稅局勢，有地緣政治風險。

第三家是日本的 SiliconEdge，型號 SE-Pro，單價 298 美元最貴，但 MOQ 只要 500，交期最短只有 6 週。品質評分 4.5 是最高的，財務穩定性 A 級，不良率也是三家最低的只有 80 PPM。"""),

        ("10:20", "Lisa Wang", """從時程來看，ChipMax 交期 12 週太長了，會影響 EVT 時程。我建議優先鎖定 NovaTech 或 SiliconEdge。DRAM 部分也有類似的三家在評估，細節我整理在供應商評估表裡了。"""),

        ("10:22", "James Chen", """價差不小，NovaTech 和 SiliconEdge 一台差 52 美元，量產 6 萬台的話就是 312 萬美元。我們需要更仔細評估。Lisa 妳本週可以完成供應商評估報告嗎？"""),

        ("10:23", "Lisa Wang", """可以，我預計週五完成評估表，包含加權評分。不過 BOM 鎖定需要等 Kevin 確認散熱方案，因為液冷模組的供應商也要納入。"""),

        ("10:25", "Amy Hsu", """品保這邊，我需要在 5 月底之前完成 DVT 測試計畫。但是前提是 R&D 要先提供完整的 BOM 和設計規格書，我才能規劃測試項目。另外散熱改成液冷的話，測試項目會增加——要加上漏液測試、長時間運轉穩定性測試、震動測試等等。這些測試設備我們部分需要外借，要提前預約。"""),

        ("10:28", "Amy Hsu", """預計 EVT 測試需要 4 週，DVT 測試需要 6 週。如果 EVT 不過，ECO 至少要再 2 週。時程真的很緊。"""),

        ("10:30", "Tom Lin", """IT 這邊，ERP 系統需要新增 ProServer X200 的料號體系。這包含成品料號、半成品料號、所有新料件的料號。請 Lisa 提供初版 BOM 之後我來建檔，大概需要 3 個工作天。另外 MES 產線系統也要設定新機種的 SOP 和檢測站流程，這個等 DVT 通過後再處理。"""),

        ("10:33", "Sarah Chang", """流程管理這邊，我會負責追蹤整個 NPI 流程的合規性，確保每個 Gate Review 的文件都齊備。我建議我們用 Teams 建立一個 #NPI-ProServerX200 的專用頻道，所有跨部門溝通都在那邊進行，方便追蹤和紀錄。每兩週一次進度會議，我會負責會議紀錄和行動追蹤。"""),

        ("10:36", "James Chen", """好的，Sarah 的建議很好。那我做個總結。

第一，Kevin 下週五 4 月 11 號之前完成散熱方案評估和修改圖面，5 月 5 號完成機構設計驗證。

第二，Lisa 這週五 4 月 4 號完成 GPU 供應商評估報告和初版 BOM。BOM 鎖定目標 5 月 8 號。

第三，Amy 在拿到完整 BOM 和設計規格後兩週內產出 DVT 測試計畫。同時這週開始預約外部測試設備。

第四，Tom 收到 BOM 後 3 天內完成 ERP 建檔。

第五，Sarah 今天建立 Teams 頻道 #NPI-ProServerX200，設定兩週一次的進度會議。

還有沒有其他議題？"""),

        ("10:39", "Lisa Wang", """補充一點，如果美中關稅在 Q3 有新的變化，ChipMax 的方案可能需要完全排除。我建議我們同步啟動 Plan B——NovaTech 為主供應商、SiliconEdge 為備援。"""),

        ("10:41", "James Chen", """同意，這個列為風險項目追蹤。好的，那今天的會議就到這裡。下次會議 4 月 15 號，同一時間，到時候各部門報告進度。散會，謝謝大家！"""),
    ]

    for time, speaker, content in transcript:
        p = doc.add_paragraph()
        run_time = p.add_run(f"[{time}] ")
        run_time.bold = True
        run_time.font.color.rgb = RGBColor(128, 128, 128)
        run_time.font.size = Pt(10)

        run_speaker = p.add_run(f"{speaker}：")
        run_speaker.bold = True
        run_speaker.font.size = Pt(11)
        run_speaker.font.color.rgb = RGBColor(47, 84, 150)

        doc.add_paragraph(content.strip())

    # Action Items Summary
    doc.add_page_break()
    doc.add_heading('行動項目摘要', level=2)

    action_items = [
        ("A-001", "Kevin Liu (R&D)", "散熱方案評估與機構修改圖面", "2025/4/11"),
        ("A-002", "Kevin Liu (R&D)", "機構設計驗證完成", "2025/5/5"),
        ("A-003", "Lisa Wang (SCM)", "GPU 供應商評估報告 + 初版 BOM", "2025/4/4"),
        ("A-004", "Lisa Wang (SCM)", "BOM 鎖定", "2025/5/8"),
        ("A-005", "Amy Hsu (QA)", "DVT 測試計畫（BOM完成後2週內）", "依BOM完成日"),
        ("A-006", "Amy Hsu (QA)", "預約外部測試設備", "2025/4/11"),
        ("A-007", "Tom Lin (IT)", "ERP 料號建檔（BOM完成後3天內）", "依BOM完成日"),
        ("A-008", "Sarah Chang (BPM)", "建立 Teams 頻道 #NPI-ProServerX200", "2025/4/1"),
        ("A-009", "Sarah Chang (BPM)", "設定兩週一次進度會議", "2025/4/1"),
    ]

    add_styled_table(doc,
        ["編號", "負責人", "行動項目", "到期日"],
        action_items
    )

    path = os.path.join(OUTPUT_DIR, "英業達_NPI_Kickoff_會議逐字稿.docx")
    doc.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 6: 英業達_NPI_設計審查會議逐字稿.docx
# ═══════════════════════════════════════════════════════════
def create_design_review_transcript():
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('英業達股份有限公司')
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(47, 84, 150)

    sub = doc.add_heading('ProServer X200 設計審查會議（Design Review Meeting）', level=1)
    for run in sub.runs:
        run.font.color.rgb = RGBColor(47, 84, 150)

    info_lines = [
        ("日期", "2025年4月15日（星期二）14:00-15:30"),
        ("地點", "桃園總部 3F 會議室 B / Microsoft Teams 線上同步"),
        ("會議類型", "Design Review + Progress Meeting"),
        ("專案代號", "PSX200"),
        ("主席", "James Chen (資深專案經理)"),
    ]

    for label, val in info_lines:
        p = doc.add_paragraph()
        run1 = p.add_run(f"{label}：")
        run1.bold = True
        run2 = p.add_run(val)

    doc.add_heading('出席者', level=2)

    attendees = [
        ("James Chen", "PM", "專案管理部", "資深專案經理"),
        ("Lisa Wang", "SCM", "供應鏈管理部", "資深供應鏈經理"),
        ("Kevin Liu", "R&D", "研發設計部", "硬體設計主管"),
        ("David Wu", "R&D", "研發設計部", "散熱工程師"),
        ("Amy Hsu", "QA", "品質保證部", "品質工程經理"),
        ("Tom Lin", "IT", "資訊系統部", "IT 系統工程師"),
        ("Sarah Chang", "BPM", "流程管理部", "流程改善專員"),
    ]

    add_styled_table(doc, ["姓名", "部門代碼", "部門", "職稱"], attendees)

    doc.add_paragraph()
    doc.add_heading('會議逐字紀錄', level=2)

    transcript = [
        ("14:02", "James Chen", """各位下午好，今天是 ProServer X200 的第二次進度會議，同時也是散熱方案的設計審查。先請 Kevin 報告研發進度。"""),

        ("14:04", "Kevin Liu", """上次會議後，我跟 David 確認了液冷方案的可行性。好消息是熱模擬結果出來了——液冷方案在 GPU 滿載 400W 時，junction temperature 可以控制在 72 度，比客戶要求的 85 度上限低了 13 度，散熱餘裕充足。"""),

        ("14:06", "David Wu", """補充一下技術細節。我們採用的是直接接觸式液冷板 (cold plate) 加上外部 CDU (Coolant Distribution Unit) 的架構。管路佈局已經完成，不會增加 2U 機殼的外部尺寸。但有一個變更——液冷管路需要從機箱後方出，所以後面板的 IO 佈局要微調。這部分已經跟 Kevin 確認過，改動不大。

液冷迴路的工作參數如下：
- 冷卻液：50% Propylene Glycol 水溶液
- 流量：每 GPU 約 0.8 L/min
- 進液溫度：35°C（CDU 出口）
- 出液溫度：約 45°C
- 壓損：< 30 kPa

漏液防護方面，我們在 cold plate 接頭處加了雙重 O-ring 密封，並在機殼底部加了漏液偵測感測器，觸發時會自動關機保護。"""),

        ("14:09", "Kevin Liu", """機構設計修改圖面在 4 月 10 號已經完成，比預期早一天。電路設計沒有受到影響，預計 5 月 15 號完成。目前最大的議題是液冷散熱模組的供應商。"""),

        ("14:11", "Lisa Wang", """液冷模組我評估了兩家：

第一家是台灣的 CoolTech，單價 95 美元，交期 5 週，品質不錯，過去合作過散熱風扇模組，配合度好。他們有伺服器液冷的量產經驗，客戶包含 Supermicro 和 QCT。

第二家是中國的 ThermalPro，單價 72 美元，便宜 24%，但交期要 9 週。品質紀錄一般，過去不良率大概在 500 PPM 左右。而且考量地緣政治風險，加上我們 GPU 已經在考慮非中國供應商了，液冷也用中國的話風險太集中。

我建議選 CoolTech。

另外 GPU 供應商評估報告已經在上週五發出了，結論是建議 NovaTech 為主、SiliconEdge 為備援。加權評分 NovaTech 得 82 分、SiliconEdge 85 分、ChipMax 只有 68 分。雖然 SiliconEdge 分數最高，但價差的影響在量產規模下很大，所以還是建議 NovaTech 為主供應商。BOM 初版也完成了，已經給 Tom 建 ERP。"""),

        ("14:15", "Tom Lin", """收到 BOM 了，我已經開始建料號，預計週三 4 月 17 號完成。有一個問題——液冷模組的 CoolTech 在我們的供應商資料庫裡沒有，需要先跑供應商導入流程。這個流程通常需要 5-7 個工作天，包含財務審查、品質審查和合約簽署。Lisa 那邊可以啟動嗎？"""),

        ("14:17", "Lisa Wang", """已經在跑了，預計下週完成供應商導入的文件審查。CoolTech 的財務報告和品質認證文件我已經收到了，看起來沒有問題。"""),

        ("14:19", "Amy Hsu", """品保更新。我已經拿到初版 BOM 和設計規格了，DVT 測試計畫正在制定中。因為改成液冷，我新增了以下測試項目：

1. 液冷迴路漏液測試：96 小時持續運轉，每 24 小時檢查一次
2. 高溫高濕耐久測試：85°C / 85% RH 條件下 500 小時
3. 運輸振動測試：模擬陸運 + 海運的複合振動
4. 液冷管路耐壓測試：2 倍工作壓力持續 30 分鐘
5. 快速溫度循環測試：-10°C 到 55°C，200 cycles

外部測試設備已經預約到 SGS 台灣的環境可靠度實驗室，8 月第二週可用。UL 那邊也有在問，作為備案。"""),

        ("14:22", "Amy Hsu", """另外補充一個風險——EVT 如果一次不過，ECO 加上重新驗證至少需要 2 週。特別是液冷的部分，如果漏液測試不過，修改模具可能要 4 週。我建議 DVT 目標不要太樂觀，預留 2 週 buffer。"""),

        ("14:24", "Sarah Chang", """流程追蹤方面，#NPI-ProServerX200 頻道已經建立，目前所有人都已加入。我也設定好了每兩週的定期會議。上次的行動項目追蹤：

✅ A-001 Kevin：散熱方案評估與機構修改圖面 → 已完成（4/10，提前1天）
✅ A-003 Lisa：GPU 供應商評估報告 + 初版 BOM → 已完成（4/4，準時）
🔄 A-005 Amy：DVT 測試計畫 → 進行中，預計 4 月底完成
🔄 A-007 Tom：ERP 料號建檔 → 進行中，預計 4/17 完成
⬜ A-007b Tom：MES 產線設定 → 等 DVT 通過後啟動
✅ A-008 Sarah：Teams 頻道建立 → 已完成
✅ A-009 Sarah：定期會議排程 → 已完成

整體進度落後 5 天（機構設計延遲），但在可控範圍。"""),

        ("14:27", "James Chen", """進度整體來說還不錯，散熱問題已經解決，這是最大的技術風險排除。那我總結今天的行動項目：

第一，Kevin 加上 David，4 月底完成電路設計整合液冷介面。

第二，Lisa 下週完成 CoolTech 供應商導入，同時確認 NovaTech GPU 供應商合約條款。

第三，Amy 4 月 30 號前完成 DVT 測試計畫，含液冷相關新增測試項目。

第四，Tom 4 月 17 號完成 ERP 建檔。

第五，Sarah 更新專案追蹤表，下次會議 4 月 29 號。

大家辛苦了，散會！"""),
    ]

    for time, speaker, content in transcript:
        p = doc.add_paragraph()
        run_time = p.add_run(f"[{time}] ")
        run_time.bold = True
        run_time.font.color.rgb = RGBColor(128, 128, 128)
        run_time.font.size = Pt(10)

        run_speaker = p.add_run(f"{speaker}：")
        run_speaker.bold = True
        run_speaker.font.size = Pt(11)
        run_speaker.font.color.rgb = RGBColor(47, 84, 150)

        doc.add_paragraph(content.strip())

    # Technical Decision Summary
    doc.add_page_break()
    doc.add_heading('技術決策摘要', level=2)

    decisions = [
        ("TD-001", "散熱方案", "採用直接接觸液冷板 + CDU 架構", "Kevin Liu", "2025/4/15"),
        ("TD-002", "液冷供應商", "選用 CoolTech (台灣)，棄用 ThermalPro (中國)", "Lisa Wang", "2025/4/15"),
        ("TD-003", "GPU 供應商策略", "NovaTech 為主供應商，SiliconEdge 為備援", "Lisa Wang", "2025/4/4"),
        ("TD-004", "IO 面板佈局", "後面板微調以配合液冷管路出口", "Kevin Liu", "2025/4/15"),
    ]

    add_styled_table(doc,
        ["編號", "決策項目", "決議內容", "負責人", "決議日期"],
        decisions,
        "548235"
    )

    doc.add_heading('新增行動項目', level=2)
    new_actions = [
        ("A-010", "Kevin Liu + David Wu", "電路設計整合液冷介面", "2025/4/30"),
        ("A-011", "Lisa Wang", "CoolTech 供應商導入完成", "2025/4/25"),
        ("A-012", "Lisa Wang", "NovaTech GPU 合約條款確認", "2025/4/25"),
        ("A-013", "Amy Hsu", "DVT 測試計畫完成（含液冷測試）", "2025/4/30"),
        ("A-014", "Tom Lin", "ERP 料號建檔完成", "2025/4/17"),
        ("A-015", "Sarah Chang", "專案追蹤表更新", "2025/4/16"),
    ]

    add_styled_table(doc, ["編號", "負責人", "行動項目", "到期日"], new_actions)

    path = os.path.join(OUTPUT_DIR, "英業達_NPI_設計審查會議逐字稿.docx")
    doc.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 7: 英業達_供應商往來郵件串.docx
# ═══════════════════════════════════════════════════════════
def create_supplier_emails():
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('英業達 ↔ NovaTech 供應商往來郵件串')
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(47, 84, 150)

    p = doc.add_paragraph()
    run = p.add_run('主旨：Inventec - GPU Module 合作洽談 | 共 4 封郵件 | 2025/3/25 - 2025/4/3')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(128, 128, 128)

    emails = [
        {
            "num": "郵件 1/4",
            "from": "Lisa Wang <lisa.wang@inventec.com>",
            "to": "David Park <david.park@novatech.co.kr>",
            "cc": "",
            "date": "2025/3/25 09:15 (GMT+8)",
            "subject": "Inventec - GPU Module 合作洽談",
            "body": """Dear David,

感謝您上週四在我們桃園總部的產品簡報，團隊對 NovaTech 的 NT-500 GPU Module 印象深刻，尤其是 PCIe Gen5 的效能表現和您們在伺服器市場的成功案例。

我們目前正在進行新一代 AI 推論伺服器（ProServer X200）的 NPI 專案，需要一款支援 400W TDP 的高效能 GPU Module。NT-500 的規格基本上符合我們的需求，我想進一步確認以下商務條件：

1. MOQ 及單價階梯（1K / 5K / 10K 級距）
2. 標準 Lead Time 及急單交付能力
3. 長約（6個月/12個月）是否有額外折扣
4. 品質認證文件（ISO、可靠度測試報告、MTBF 數據）
5. 技術支援服務條款（回應時間 SLA、台灣在地 FAE 支援）

我們的 BOM 鎖定時程在 5 月初，所以希望能在 4 月中之前收到正式報價。

另外想請教，NT-500 在 85°C ambient 條件下的降頻行為（Power Limit Throttling）如何？我們的散熱設計目標是 junction temp 控制在 85°C 以下。

期待您的回覆。

Best regards,
Lisa Wang
Senior Supply Chain Manager
Inventec Corporation
Tel: +886-3-XXX-XXXX ext. 2580
"""
        },
        {
            "num": "郵件 2/4",
            "from": "David Park <david.park@novatech.co.kr>",
            "to": "Lisa Wang <lisa.wang@inventec.com>",
            "cc": "James Lee <james.lee@novatech.co.kr> (FAE Manager)",
            "date": "2025/3/27 14:30 (GMT+9)",
            "subject": "Re: Inventec - GPU Module 合作洽談",
            "body": """Dear Lisa,

Thank you for your interest in our NT-500 GPU Module and the warm hospitality during my visit. 以下是您詢問的商務條件：

1. 價格階梯：
   • 1K pcs: USD 285/unit
   • 5K pcs: USD 275/unit
   • 10K pcs: USD 268/unit
   (以上為 CIF 桃園價格，含運費及保險)

2. 標準 Lead Time: 8 週 (from PO confirmation)
   急單（加價 5%）: 可縮短至 6 週
   目前 Q3 產能利用率約 85%，建議提早確認數量

3. 長約折扣：
   • 6 個月合約: 額外 2% off
   • 12 個月合約: 額外 4% off
   • 合約需包含最低取貨量承諾

4. 品質認證文件：附件中包含：
   • ISO 9001:2015 證書
   • ISO 14001:2015 證書
   • NT-500 可靠度測試報告（MTBF > 200,000 小時）
   • HALT/HASS 測試結果
   • RoHS / REACH 合規聲明

5. 技術支援 SLA:
   • 標準回應時間: 4 小時（韓國工作時間 KST 9:00-18:00）
   • 台灣在地 FAE: James Lee 可提供技術支援，每季至少一次 on-site visit
   • 緊急技術問題: 24 小時內提供 remote debug support

關於 NT-500 的降頻行為：
   • Ambient 85°C 條件下，如果 junction temp 超過 T_j_max (105°C)，會觸發 Power Limit Throttling，約降頻 5%
   • 但如果散熱設計能將 junction temp 控制在 85°C 以下，NT-500 可維持 full boost clock，不會降頻
   • 建議安排我們的 FAE James 與貴司 R&D 做技術討論，確認散熱設計能滿足需求

4 月 8 號下午 2 點的 Teams 會議沒問題，我會寄出邀請。

Best regards,
David Park
Sales Director, NovaTech Co., Ltd.
Tel: +82-2-XXX-XXXX
"""
        },
        {
            "num": "郵件 3/4",
            "from": "Lisa Wang <lisa.wang@inventec.com>",
            "to": "David Park <david.park@novatech.co.kr>",
            "cc": "James Chen <james.chen@inventec.com> (PM)",
            "date": "2025/4/2 10:45 (GMT+8)",
            "subject": "Re: Re: Inventec - GPU Module 合作洽談",
            "body": """Dear David,

感謝您的詳細報價。報價內容我已經跟內部團隊（PM James Chen 和 R&D Kevin Liu）討論過，有以下幾點想進一步確認：

1. 數量級距重新議價：我們預估第一年用量約 240K pcs（每月約 20K），以這個量級是否有更好的價格？我們的目標價位在 USD 265 以下。

2. 品質標準確認：
   • 我們的 incoming QC 標準是 AQL 0.65 Level II，請問你們出廠檢驗標準是多少？
   • 不良品退換貨流程及時效？
   • 是否接受 PPM 條款？我們的目標是 <200 PPM

3. 產能保證：Q3 產能吃緊的問題我有些擔心。以我們 240K/年的量，能否提供書面的產能保證承諾（Capacity Reservation Agreement）？

4. 技術問題：NT-500 在使用液冷散熱方案（cold plate 直接接觸）時，cold plate 的安裝介面規格是什麼？是否有推薦的 TIM (Thermal Interface Material)？

5. 付款條件：我們標準付款條件是 Net 60 days，是否可接受？

另外，我們也在評估其他供應商（包含日本的 SiliconEdge）。希望 NovaTech 能給出最有競爭力的方案，我們很希望與 NovaTech 建立長期合作關係。

線上會議確認 4 月 8 號下午 2 點 (GMT+8)。

Best regards,
Lisa Wang
"""
        },
        {
            "num": "郵件 4/4",
            "from": "David Park <david.park@novatech.co.kr>",
            "to": "Lisa Wang <lisa.wang@inventec.com>",
            "cc": "James Lee <james.lee@novatech.co.kr>; S.K. Kim <sk.kim@novatech.co.kr> (VP Sales)",
            "date": "2025/4/3 16:20 (GMT+9)",
            "subject": "Re: Re: Re: Inventec - GPU Module 合作洽談",
            "body": """Dear Lisa,

感謝您提供的詳細需求，我已經副本給我們的銷售副總 S.K. Kim，以加速核價流程。以下逐一回覆：

1. 重新核價：年用量 240K pcs 是很有吸引力的數字。初步估計應該可以到 USD 262-265 的區間（12個月合約+含 4% 折扣的基礎上）。但需要等總部正式核准，預計下週一 (4/7) 前可以給您 official quotation。

2. 品質標準：
   • 我們的出廠檢驗是 AQL 0.4 Level II，比您的 0.65 更嚴格
   • 不良品退換貨：確認不良後 5 個工作天內空運寄出替換品，運費由我方承擔
   • PPM 條款可以接受，我們承諾 <150 PPM（我們目前實績是 150 PPM）
   • 我們也提供免費的 FA (Failure Analysis) 服務

3. 產能保證：以 240K/年的量，我們會為 Inventec 預留專屬產線產能。書面的 CRA 可以提供，承諾交期達成率 > 95%。如有產能異常，提前 8 週通知。

4. 技術資訊：
   • Cold plate 安裝介面：標準 4-point mounting，孔距 76mm x 76mm (Intel LGA 4677 相容)
   • 推薦 TIM：Shin-Etsu X-23-7921-5 (thermal conductivity 6.0 W/mK) 或 Honeywell PTM7950 (phase change material)
   • 詳細 mechanical drawing 和 thermal design guide 會在 4/8 會議上由 James Lee 提供

5. 付款條件：Net 60 days 可以接受。如果是 12 個月長約 + 每月固定取貨量，我們也可以考慮月結。

我理解貴司也在評估其他供應商，我相信 NovaTech 在品質和技術支援方面的表現會讓您滿意。我們非常重視與英業達的合作機會。

4/8 會議我會跟 James Lee (FAE) 一起出席。期待與您進一步討論！

Best regards,
David Park
"""
        },
    ]

    for email in emails:
        doc.add_paragraph()
        doc.add_heading(f"═══ {email['num']} ═══", level=2)

        fields = [
            ("寄件者", email["from"]),
            ("收件者", email["to"]),
        ]
        if email["cc"]:
            fields.append(("副本", email["cc"]))
        fields.extend([
            ("日期", email["date"]),
            ("主旨", email["subject"]),
        ])

        for label, val in fields:
            p = doc.add_paragraph()
            run1 = p.add_run(f"{label}：")
            run1.bold = True
            run1.font.size = Pt(10)
            run1.font.color.rgb = RGBColor(100, 100, 100)
            run2 = p.add_run(val)
            run2.font.size = Pt(10)

        doc.add_paragraph("─" * 50)
        for line in email["body"].strip().split("\n"):
            p = doc.add_paragraph(line)
            p.paragraph_format.space_after = Pt(2)

    path = os.path.join(OUTPUT_DIR, "英業達_供應商往來郵件串.docx")
    doc.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 8: 英業達_QBR準備郵件串.docx
# ═══════════════════════════════════════════════════════════
def create_qbr_emails():
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('2025 Q4 Dell QBR 準備工作郵件串')
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(47, 84, 150)

    p = doc.add_paragraph()
    run = p.add_run('主旨：【重要】2025 Q4 Dell QBR 準備工作分配 | 共 5 封郵件 | 2025/12/15 - 2025/12/19')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(128, 128, 128)

    emails = [
        {
            "num": "郵件 1/5 — PM 發起工作分配",
            "from": "James Chen <james.chen@inventec.com>",
            "to": "Lisa Wang; Amy Hsu; Tom Lin; Sarah Chang",
            "cc": "VP Manufacturing — Robert Tsai",
            "date": "2025/12/15 (一) 09:00",
            "subject": "【重要】2025 Q4 Dell QBR 準備工作分配",
            "body": """各位同仁，

Dell 的 2025 年度 QBR (Quarterly Business Review) 會議已確認在 2026/1/15 舉行，由 Dell Server & Networking BU 的 VP — Michael Thompson 主持。出席的還有 Dell 的 Global Procurement Director — Jennifer Walsh 和 Quality Director — Robert Kim。

這次 QBR 非常重要，因為 Q3-Q4 我們在 ProServer X100 的交期和品質表現都有下滑（交期從 97.5% 降到 91.3%，DPPM 從 150 升到 280），客戶已經在上個月的月會中明確表達了 concern，並提到如果沒有明顯改善，可能考慮增加第二供應商。

以下是準備工作分配：

1. Lisa (SCM) — 交期分析：
   • 準備 Q3-Q4 交期達成率的月度明細及 root cause 分析
   • 特別說明 10 月份交期下降到 88.5% 的原因
   • 已實施和計畫中的改善措施
   • 截止日：12/30

2. Amy (QA) — 品質報告：
   • 品質 KPI (DPPM, 客訴件數) 的季度趨勢報告
   • 9 件品質異常的 8D 報告狀態更新
   • 每件的預計關閉時程
   • Q1 2026 品質改善計畫
   • 截止日：12/30

3. Tom (IT) — 數據彙整：
   • 從 ERP 彙整全年度出貨量、營收概算、成本節省數據
   • Dashboard 所需的月度/季度原始數據（Excel 格式）
   • 截止日：12/25（比較早，因為其他人需要用你的數據）

4. Sarah (BPM) — 流程改善：
   • 彙整 2025 年度流程改善成果報告
   • 量化改善效益（效率提升%、成本節省金額等）
   • 2026 年度改善提案（含時程和預算）
   • 截止日：12/30

5. 我 (James) — 簡報整合：
   • 彙整所有資料做成 QBR 簡報（PowerPoint）
   • 撰寫 Executive Summary
   • 與 Dell Michael 確認議程和重點關注項目
   • 準備 Q&A 預演

⚠️ 1/5 (一) 下午 2 點，我們先做一次內部預演 (dry run)，請各位務必在截止日前完成。

有任何問題隨時在 Teams #Dell-QBR 頻道討論。

Thanks,
James Chen
Senior Project Manager
"""
        },
        {
            "num": "郵件 2/5 — SCM 回報交期分析進度",
            "from": "Lisa Wang <lisa.wang@inventec.com>",
            "to": "James Chen",
            "cc": "全員",
            "date": "2025/12/16 (二) 11:30",
            "subject": "Re: 【重要】2025 Q4 Dell QBR 準備工作分配",
            "body": """James,

收到，交期部分初步分析如下：

Q3-Q4 ProServer X100 交期下滑的三大根因：

1. 【來料問題】10 月有一批 DRAM（MemoryPlus 供應）來料品質問題，incoming QC reject rate 高達 8.5%（正常 <1%），需要整批退換。影響約 380 台的組裝時程，是 10 月份交期降到 88.5% 的主因。

2. 【需求激增】11 月份 Dell 臨時追加 15% 訂單量（從原本 4,500 台增加到 5,200 台），我們的產線排程已滿，無法在現有產能下 100% 消化，導致部分訂單延遲 3-5 天。

3. 【供應商交期延長】GPU 供應商 NovaTech 從 Q4 開始將交期從 8 週延長到 10 週，原因是他們自己的晶圓代工產能調整。我們有 2 批 PO 的到料晚了 2 週。

改善措施（已實施 + 計畫中）：
✅ 已將 DRAM 安全庫存從 2 週提高到 4 週（12月已生效）
✅ 已要求 MemoryPlus 加強出貨檢驗並提供 COC (Certificate of Conformance)
🔄 正在與 NovaTech 協商 Q1 2026 的產能保證合約
🔄 正在建立第二 DRAM 供應商 SiliconEdge 的備援機制（預計 Q1 完成驗證）
📋 計畫增加一條 ProServer 組裝線（需投資 USD 1.2M，預計 Q2 可用）

12/30 前提供完整報告（含月度交期數據圖表）。

Lisa
"""
        },
        {
            "num": "郵件 3/5 — QA 回報品質分析進度",
            "from": "Amy Hsu <amy.hsu@inventec.com>",
            "to": "James Chen",
            "cc": "全員",
            "date": "2025/12/17 (三) 14:15",
            "subject": "Re: 【重要】2025 Q4 Dell QBR 準備工作分配",
            "body": """James,

品質 KPI 我已經開始整理了。目前的狀況：

DPPM 趨勢（ProServer X100）：
• Q1: 180 → Q2: 150 → Q3: 220 → Q4: 280
明顯在 Q3-Q4 惡化，主要因為幾個新的品質異常集中爆發。

9 件客訴目前狀態：
• 已關閉：4 件（QA-031, 033, 038, 042）
• 進行中：3 件（QA-045 散熱片偏移、QA-048 BMC韌體、QA-051 電池續航）
• 待處理：1 件（QA-055 NIC 相容性）
• 新增：1 件（QA-053 IoT Gateway Wi-Fi 不穩）

比較擔心的幾個：
1. QA-045（GPU 散熱片組裝偏移）→ 已導入定位治具，12 月新產出已驗證改善，不良率從 4.5% 降至 0.3%
2. QA-048（BMC 韌體 IPMI 異常）→ R&D 已釋出 v2.3.2 修正版，正在客戶端驗證中
3. QA-055（Broadcom NIC 相容性問題）→ 這個最頭痛，Broadcom 那邊回應慢，firmware fix 的 ETA 是 1 月中。不確定能不能在 QBR 前有進展。

Q4 DPPM 上升根因分析：
1. GPU 散熱片組裝偏移（佔 35%）→ 治具改善已完成
2. BMC 韌體問題（佔 25%）→ 韌體更新中
3. 新進作業員訓練不足（佔 20%）→ 11月已重新排訓，增加40小時實作訓練
4. DRAM 來料不良波及（佔 12%）→ 供應商已改善
5. 其他零星問題（佔 8%）

Q1 2026 品質改善計畫初步構想：
• 導入 AI 視覺檢測（AOI 升級）在 GPU 散熱片組裝站
• 建立即時 SPC 監控系統
• 供應商品質審核頻率從每半年加密到每季
• 新進人員訓練時數從 80hr 提高到 120hr

8D 報告會在 12/30 之前全部更新完畢。

Amy
"""
        },
        {
            "num": "郵件 4/5 — IT 回報數據彙整進度",
            "from": "Tom Lin <tom.lin@inventec.com>",
            "to": "James Chen",
            "cc": "全員",
            "date": "2025/12/18 (四) 09:45",
            "subject": "Re: 【重要】2025 Q4 Dell QBR 準備工作分配",
            "body": """James,

ERP 數據我這邊動作比較快，初步數字如下：

═══ 2025 全年度 Dell 專案總覽 ═══

出貨量：
• ProServer X100: 55,500 台（YoY +18%）
• CloudBook L15: 200,000 台（YoY +8%）
• IoT Gateway G3: 38,900 台（YoY +25%，新專案成長快）
• 總計: 294,400 台

營收概算：
• 總營收：約 USD 485M（YoY +12%）
• ProServer X100 佔比：58%
• CloudBook L15 佔比：32%
• IoT Gateway G3 佔比：10%

成本節省實績：
• 全年度實際節省：USD 1,050,000
• 年度目標：USD 1,200,000
• 達成率：87.5%
• 主要差距來源：Q3-Q4 ProServer X100 的品質異常處理成本（rework + scrap）吃掉約 USD 180,000 的預計節省

Dashboard 原始數據我在整理格式中（季度 KPI + 月度交期 + 品質異常明細），12/25 前提供完整 Excel 檔。需要我做什麼格式的圖表可以先說，我可以一起準備。

Tom
"""
        },
        {
            "num": "郵件 5/5 — BPM 回報流程改善進度",
            "from": "Sarah Chang <sarah.chang@inventec.com>",
            "to": "James Chen",
            "cc": "全員",
            "date": "2025/12/19 (五) 16:00",
            "subject": "Re: 【重要】2025 Q4 Dell QBR 準備工作分配",
            "body": """James,

流程改善的部分，2025 年度成果彙整如下：

═══ 2025 年度四大改善成果 ═══

1. 導入自動化 AOI 檢測（Q1 完成）
   • 投資：USD 350,000
   • 效益：人工檢測誤判率降低 35%
   • 影響產品：ProServer X100, CloudBook L15
   • ROI：預估 8 個月回收

2. 優化 SMT 換線流程（Q2 完成）
   • 投資：USD 80,000（含 training）
   • 效益：換線時間從 45 分鐘縮短至 36 分鐘（-20%）
   • 年化節省：約 USD 120,000（產能提升）

3. 建立供應商績效即時 Dashboard（Q2 完成）
   • 投資：USD 45,000（IT 開發）
   • 效益：SCM 異常反應時間從平均 48 小時縮短至 12 小時
   • 供應商交期異常預警準確率 82%

4. 引入 AI 預測性維護系統（Q3 完成）
   • 投資：USD 520,000
   • 效益：設備異常停機次數減少 15%，平均修復時間 (MTTR) 縮短 30%
   • 影響產線：SMT Line 1-3, Assembly Line A-B

═══ 2026 年度改善提案（初稿）═══

1. 擴大 AI 品檢覆蓋率到整機組裝段
   • 預估投資：USD 600,000
   • 目標效益：DPPM 降低 40%
   • 時程：Q1-Q2 2026

2. 導入 Digital Twin 做產線模擬
   • 預估投資：USD 400,000
   • 目標效益：新機種量產準備時間縮短 25%
   • 時程：Q2-Q3 2026

3. 建立即時品質 SPC 監控系統
   • 預估投資：USD 200,000
   • 目標效益：品質異常偵測時間從 4 小時縮短至即時
   • 時程：Q1 2026

12/30 前提供完整報告含投資效益分析表。

Sarah
"""
        },
    ]

    for email in emails:
        doc.add_paragraph()
        doc.add_heading(f"═══ {email['num']} ═══", level=2)

        fields = [
            ("寄件者", email["from"]),
            ("收件者", email["to"]),
        ]
        if email["cc"]:
            fields.append(("副本", email["cc"]))
        fields.extend([
            ("日期", email["date"]),
            ("主旨", email["subject"]),
        ])

        for label, val in fields:
            p = doc.add_paragraph()
            run1 = p.add_run(f"{label}：")
            run1.bold = True
            run1.font.size = Pt(10)
            run1.font.color.rgb = RGBColor(100, 100, 100)
            run2 = p.add_run(val)
            run2.font.size = Pt(10)

        doc.add_paragraph("─" * 50)
        for line in email["body"].strip().split("\n"):
            p = doc.add_paragraph(line)
            p.paragraph_format.space_after = Pt(2)

    path = os.path.join(OUTPUT_DIR, "英業達_QBR準備郵件串.docx")
    doc.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 9: 英業達_供應鏈風險分析報告.docx
# ═══════════════════════════════════════════════════════════
def create_risk_report():
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('英業達股份有限公司')
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(47, 84, 150)

    doc.add_heading('供應鏈風險分析報告', level=1)
    doc.add_heading('ProServer X200 專案 — GPU Module 供應風險評估', level=2)

    info = [
        ("文件編號", "SCM-2025-RA-003"),
        ("版本", "1.0"),
        ("日期", "2025/4/18"),
        ("撰寫單位", "供應鏈管理部"),
        ("撰寫人", "Lisa Wang, Senior SCM Manager"),
        ("核准人", "__________ (供應鏈副總)"),
        ("機密等級", "內部機密 — Confidential"),
    ]
    for label, val in info:
        p = doc.add_paragraph()
        run1 = p.add_run(f"{label}：")
        run1.bold = True
        run2 = p.add_run(val)

    doc.add_paragraph()
    doc.add_heading('目錄', level=1)
    toc = [
        "1. 執行摘要",
        "2. 專案背景",
        "3. 風險識別與評估",
        "   3.1 地緣政治風險",
        "   3.2 單一供應商依賴風險",
        "   3.3 價格波動風險",
        "   3.4 產能瓶頸風險",
        "4. 替代方案分析",
        "5. 成本影響評估",
        "6. 建議與行動計畫",
        "7. 附件",
    ]
    for item in toc:
        p = doc.add_paragraph(item)
        p.paragraph_format.space_after = Pt(2)

    doc.add_page_break()

    # Section 1
    doc.add_heading('1. 執行摘要', level=1)
    p = doc.add_paragraph()
    run = p.add_run('[此章節請使用 Copilot 草擬]')
    run.font.color.rgb = RGBColor(192, 0, 0)
    run.bold = True
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('💡 Copilot 提示範例：')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    doc.add_paragraph('「根據本報告的第 2-6 章內容，撰寫一段約 200 字的執行摘要。重點涵蓋：主要風險、建議的供應商策略、預估成本影響。語氣專業，適合高階管理層閱讀。」')

    # Section 2
    doc.add_heading('2. 專案背景', level=1)
    doc.add_paragraph("""英業達 (Inventec Corporation) 為全球前五大伺服器 ODM (Original Design Manufacturer) 廠商，年營收超過新台幣 4,000 億元。公司主要產品包含伺服器、筆記型電腦及 IoT 裝置，客戶涵蓋 HP Enterprise、Dell Technologies、Lenovo 等全球知名品牌。

ProServer X200 為英業達因應 HP Enterprise AI 推論伺服器需求所開發的新一代產品。該產品搭載高效能 GPU 加速卡，為 BOM 中單一最高成本零組件（CPU 除外），佔 BOM 比例約 14.8%，年度採購金額預估超過 USD 68M。

目前 GPU Module 主要供應商為韓國 NovaTech Co., Ltd.，為 100% 單一來源供應。鑑於近期全球地緣政治局勢變化及供應鏈不確定性，本報告旨在全面評估 GPU Module 的供應風險，並提出風險緩解建議。""")

    # Section 3
    doc.add_heading('3. 風險識別與評估', level=1)

    doc.add_heading('3.1 地緣政治風險', level=2)
    doc.add_paragraph("""風險等級：🟡 中高

主要風險因素：
• 美中科技制裁持續擴大：中國供應商 ChipMax 生產的 GPU Module 使用部分美國技術授權的製程，面臨出口管制風險。若未來制裁範圍擴大，ChipMax 可能無法繼續供貨。
• 美韓 CHIPS 法案影響：NovaTech 位於韓國，受美韓半導體合作框架約束。部分產能需優先供應美國客戶，可能壓縮對亞洲 ODM 客戶的供應量。
• 關稅風險：中國進口零件面臨 25% 額外關稅，直接影響 ChipMax 方案的成本競爭力。""")

    p = doc.add_paragraph()
    run = p.add_run('[請使用 Researcher Agent 搜尋補充最新地緣政治情勢分析]')
    run.font.color.rgb = RGBColor(192, 0, 0)
    run.bold = True
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('💡 Researcher Agent 提示：')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    doc.add_paragraph('「搜尋 2025 年最新的美中半導體出口管制政策更新，以及對 GPU 模組供應鏈的影響分析。」')

    doc.add_heading('3.2 單一供應商依賴風險', level=2)
    doc.add_paragraph("""風險等級：🔴 嚴重

目前 GPU Module 100% 依賴 NovaTech 供應，存在以下風險：
• NovaTech Q3/Q4 2025 產能利用率已達 85%，交期從標準 8 週延長至 10 週
• 若 NovaTech 發生天災、罷工、設備故障等不可抗力事件，將直接中斷 ProServer X200 的生產
• NovaTech 同時供貨給 Dell、Lenovo 等競爭客戶，產能分配可能不利於英業達
• 單一供應商在議價能力上對英業達不利，過去一年已漲價 3-5%""")

    doc.add_heading('3.3 價格波動風險', level=2)
    doc.add_paragraph("""風險等級：🟡 中等

• GPU Module 過去 12 個月漲價約 12%，主要受晶圓代工成本上升及 AI 需求激增推動
• DRAM 現貨市場價格波動 ±15%，影響 BOM 成本穩定性
• 美元對韓元/日幣匯率波動可能影響 NovaTech 和 SiliconEdge 的報價""")

    doc.add_heading('3.4 產能瓶頸風險', level=2)
    doc.add_paragraph("""風險等級：🟡 中高

• 全球 AI 伺服器需求爆發，GPU Module 供不應求
• NovaTech Q3 產能預訂率 85%，Q4 預估將達 90%
• SiliconEdge 備用產能僅 10%，大量接單能力受限
• ChipMax 雖有 30% 備用產能，但品質和地緣政治風險限制其可用性""")

    # Section 4
    doc.add_heading('4. 替代方案分析', level=1)
    p = doc.add_paragraph()
    run = p.add_run('[此章節請使用 Copilot 引用 英業達_供應商評估表.xlsx 和 英業達_BOM成本分析.xlsx 的數據來生成分析]')
    run.font.color.rgb = RGBColor(192, 0, 0)
    run.bold = True
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('💡 Copilot 提示範例：')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    doc.add_paragraph('「參考 /英業達_供應商評估表.xlsx 和 /英業達_BOM成本分析.xlsx，針對三家 GPU 供應商（NovaTech、ChipMax、SiliconEdge）進行比較分析。從價格、品質、交期、風險四個維度展開，並給出推薦排序。」')

    doc.add_paragraph("""初步方案概述：

方案一（維持現狀）：100% NovaTech
方案二（切換）：100% SiliconEdge
方案三（切換）：100% ChipMax
方案四（分散風險 — 推薦）：NovaTech 60% + SiliconEdge 40%

各方案詳細比較請參見 英業達_BOM成本分析.xlsx「GPU替代方案成本比較」工作表。""")

    # Section 5
    doc.add_heading('5. 成本影響評估', level=1)
    p = doc.add_paragraph()
    run = p.add_run('[此章節請使用 Copilot + Analyst Agent 根據 Excel 數據生成定量分析]')
    run.font.color.rgb = RGBColor(192, 0, 0)
    run.bold = True
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('💡 Analyst Agent 提示範例：')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    doc.add_paragraph('「分析 BOM成本分析.xlsx 中四個替代方案的年度成本差異，考慮單價差異、換線成本、驗證成本、品質風險成本（預估不良率×處理成本）。計算 12 個月的 TCO (Total Cost of Ownership) 並產出比較圖表。」')

    # Section 6
    doc.add_heading('6. 建議與行動計畫', level=1)
    p = doc.add_paragraph()
    run = p.add_run('[此章節請使用 Copilot 草擬]')
    run.font.color.rgb = RGBColor(192, 0, 0)
    run.bold = True
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('💡 Copilot 提示範例：')
    run.bold = True
    run.font.color.rgb = RGBColor(0, 112, 192)
    doc.add_paragraph('「基於本報告的風險評估和替代方案分析，撰寫建議與行動計畫章節。建議採用雙供應商策略（NovaTech + SiliconEdge），列出具體行動項目、負責人、時程，以及需要管理層核可的事項。格式使用編號清單，語氣果斷專業。」')

    doc.add_paragraph("""初步建議方向：

建議採用「方案四：雙供應商策略」(NovaTech 60% + SiliconEdge 40%)

理由：
1. 分散地緣政治及單一供應商風險
2. 雖然年化成本略增，但供應穩定性大幅提升
3. 增加議價籌碼
4. SiliconEdge 品質最優（PPM 80），可拉高整體品質水準""")

    # Section 7
    doc.add_heading('7. 附件', level=1)
    doc.add_paragraph("""附件清單：
1. 英業達_供應商評估表.xlsx — 三家供應商完整評估數據
2. 英業達_BOM成本分析.xlsx — ProServer X200 BOM 明細及替代方案成本比較
3. NovaTech ISO 9001 / ISO 14001 認證書（影本）
4. SiliconEdge ISO 9001 / IATF 16949 認證書（影本）
5. 供應商財務評級報告（Dun & Bradstreet）
""")

    path = os.path.join(OUTPUT_DIR, "英業達_供應鏈風險分析報告.docx")
    doc.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# 檔案 10: 英業達_NPI週報範本.docx
# ═══════════════════════════════════════════════════════════
def create_npi_weekly_report():
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('英業達股份有限公司')
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(47, 84, 150)

    doc.add_heading('NPI 週報 — ProServer X200', level=1)

    info = [
        ("專案代號", "PSX200"),
        ("客戶", "HP Enterprise"),
        ("報告週次", "W16 (2025/4/14 - 4/18)"),
        ("報告人", "James Chen (PM)"),
        ("發送對象", "NPI 核心團隊 + 製造副總 Robert Tsai"),
        ("下次會議", "2025/4/29 (二) 10:00"),
    ]
    for label, val in info:
        p = doc.add_paragraph()
        run1 = p.add_run(f"{label}：")
        run1.bold = True
        run2 = p.add_run(val)

    # Section 1
    doc.add_heading('一、整體狀態', level=1)

    status_table = [
        ["整體進度", "🟡 黃燈 — 落後5天（機構設計延遲已回復，目前可控）"],
        ["時程風險", "🟡 中等 — EVT 時程緊張，液冷測試項目增加"],
        ["成本狀態", "🟢 正常 — BOM 成本在目標範圍內"],
        ["品質風險", "🟢 低 — 散熱方案已通過熱模擬驗證"],
        ["供應鏈風險", "🟡 中等 — NovaTech Q3 產能緊張需提前鎖定"],
    ]
    add_styled_table(doc,
        ["項目", "狀態"],
        status_table,
        "2F5496"
    )

    # Section 2
    doc.add_heading('二、本週進度摘要', level=1)

    progress = [
        ("研發設計 (R&D)", [
            "✅ 液冷散熱方案圖面修改完成 (4/10)，提前 1 天",
            "✅ 熱模擬驗證通過：GPU junction temp 72°C (目標 <85°C)",
            "🔄 電路設計進行中，整合液冷介面設計",
            "🔄 後面板 IO 佈局微調（配合液冷管路出口）",
        ]),
        ("供應鏈 (SCM)", [
            "✅ GPU 供應商評估報告完成 — 推薦 NovaTech 為主、SiliconEdge 備援",
            "✅ 初版 BOM 完成並交付 IT 建檔",
            "🔄 CoolTech (液冷模組) 供應商導入流程進行中",
            "🔄 NovaTech 年度合約條款協商中，等待總部核價",
        ]),
        ("品質保證 (QA)", [
            "🔄 DVT 測試計畫制定中，新增 5 項液冷相關測試",
            "✅ 外部測試設備已預約 SGS (8月第二週)",
            "📋 UL 實驗室列為備案",
        ]),
        ("資訊系統 (IT)", [
            "🔄 ERP 料號建檔進行中，預計 4/17 完成",
            "⚠️ CoolTech 尚未在供應商資料庫中，需配合 SCM 完成導入",
        ]),
        ("流程管理 (BPM)", [
            "✅ Teams 頻道 #NPI-ProServerX200 已建立",
            "✅ 兩週一次定期會議已排定",
            "✅ 上次行動項目追蹤已更新",
        ]),
    ]

    for dept, items in progress:
        p = doc.add_paragraph()
        run = p.add_run(f"▎{dept}")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(47, 84, 150)
        for item in items:
            doc.add_paragraph(item, style='List Bullet')

    # Section 3
    doc.add_heading('三、風險與議題', level=1)

    risks = [
        ["🔴 嚴重", "R-001", "GPU 供應商 NovaTech Q3 產能預訂 85%，可能影響交期", "Lisa Wang", "提前下單 + SiliconEdge 備援"],
        ["🟡 高", "R-006", "美中關稅政策變化可能排除 ChipMax 方案", "Lisa Wang", "優先使用韓/日供應商"],
        ["🟡 高", "R-007", "量產時程壓縮，僅 7 個月", "James Chen", "平行作業 + 加速驗證"],
        ["🟡 中", "R-003", "DRAM 價格波動 ±15%", "Lisa Wang", "提前鎖量鎖價"],
        ["🟡 中", "R-005", "SGS 測試設備檔期緊張", "Amy Hsu", "已預約 + UL 備案"],
        ["🟢 已緩解", "R-002", "散熱方案驗證 — 液冷方案已通過熱模擬", "Kevin Liu", "維持監控"],
    ]

    add_styled_table(doc,
        ["風險等級", "編號", "描述", "負責人", "緩解措施"],
        risks,
        "C00000"
    )

    # Section 4
    doc.add_heading('四、行動項目追蹤', level=1)

    actions = [
        ["A-001", "散熱方案修改圖面", "Kevin Liu", "4/11", "✅ 完成 (4/10)"],
        ["A-003", "GPU 供應商評估報告", "Lisa Wang", "4/4", "✅ 完成"],
        ["A-003b", "初版 BOM", "Lisa Wang", "4/4", "✅ 完成"],
        ["A-008", "Teams 頻道建立", "Sarah Chang", "4/1", "✅ 完成"],
        ["A-009", "定期會議排程", "Sarah Chang", "4/1", "✅ 完成"],
        ["A-010", "電路設計整合液冷介面", "Kevin Liu", "4/30", "🔄 進行中 (50%)"],
        ["A-011", "CoolTech 供應商導入", "Lisa Wang", "4/25", "🔄 進行中"],
        ["A-012", "NovaTech GPU 合約確認", "Lisa Wang", "4/25", "🔄 進行中"],
        ["A-013", "DVT 測試計畫", "Amy Hsu", "4/30", "🔄 進行中 (30%)"],
        ["A-014", "ERP 料號建檔", "Tom Lin", "4/17", "🔄 進行中 (80%)"],
        ["A-015", "專案追蹤表更新", "Sarah Chang", "4/16", "🔄 進行中"],
    ]

    add_styled_table(doc,
        ["編號", "行動項目", "負責人", "到期日", "狀態"],
        actions,
        "548235"
    )

    # Section 5
    doc.add_heading('五、下週計畫 (W17: 4/21-4/25)', level=1)

    next_week = [
        ("Kevin Liu (R&D)", "持續推進電路設計，完成液冷介面佈線"),
        ("Lisa Wang (SCM)", "完成 CoolTech 供應商導入 + NovaTech 合約最終條款"),
        ("Amy Hsu (QA)", "持續 DVT 測試計畫制定，確認液冷測試標準"),
        ("Tom Lin (IT)", "完成 ERP 建檔 (4/17)，開始規劃 MES 設定"),
        ("James Chen (PM)", "與 HP 客戶確認 EVT 樣品數量和交付時程"),
    ]

    for person, plan in next_week:
        p = doc.add_paragraph()
        run1 = p.add_run(f"• {person}：")
        run1.bold = True
        run2 = p.add_run(plan)

    # Section 6
    doc.add_heading('六、需要管理層決策事項', level=1)

    p = doc.add_paragraph()
    run = p.add_run('決策事項 1：GPU 供應商策略')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(192, 0, 0)

    doc.add_paragraph("""建議採用 NovaTech 為主供應商 + SiliconEdge 為備援的雙供應商策略。

• NovaTech 單價 USD 285（量產價可能降至 USD 262-265）
• SiliconEdge 單價 USD 298（單價高 +$13，但品質最優 PPM 80、交期最短 6 週）
• 雙供應商策略初期需額外驗證費用約 USD 45,000
• 長期效益：降低供應中斷風險，增加議價籌碼

👉 請管理層核可：
1. 啟動 SiliconEdge 的供應商驗證（費用 USD 45,000）
2. 確認雙供應商的分配比例（建議 NovaTech 60% / SiliconEdge 40%）""")

    p = doc.add_paragraph()
    run = p.add_run('決策事項 2：液冷散熱模組供應商選擇')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(192, 0, 0)

    doc.add_paragraph("""建議選用台灣 CoolTech（單價 USD 95），而非中國 ThermalPro（單價 USD 72）。

• 年化成本差異：約 USD 1.38M（60,000台 × 2個/台 × $23 差價）
• 但 CoolTech 品質風險顯著較低、交期較短（5週 vs 9週）、且避免地緣政治風險集中

👉 請管理層核可選用 CoolTech 作為液冷散熱模組供應商。""")

    doc.add_paragraph()
    doc.add_paragraph("═" * 40)
    p = doc.add_paragraph()
    run = p.add_run('報告結束 | Next Meeting: 2025/4/29 (二) 10:00')
    run.bold = True
    run.font.color.rgb = RGBColor(128, 128, 128)

    path = os.path.join(OUTPUT_DIR, "英業達_NPI週報範本.docx")
    doc.save(path)
    print(f"✅ Created: {path}")


# ═══════════════════════════════════════════════════════════
# Run all
# ═══════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  AB-730 英業達 — 建立 Word 範例檔案")
    print("=" * 60)
    create_npi_kickoff_transcript()
    create_design_review_transcript()
    create_supplier_emails()
    create_qbr_emails()
    create_risk_report()
    create_npi_weekly_report()
    print("=" * 60)
    print("  All Word files created successfully!")
    print("=" * 60)
