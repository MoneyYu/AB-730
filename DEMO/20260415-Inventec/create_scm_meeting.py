# -*- coding: utf-8 -*-
"""
AB-730 英業達課程 — 建立供應鏈風險緊急會議逐字稿
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

def create_scm_meeting_transcript():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

    title = doc.add_heading('', level=0)
    run = title.add_run('英業達股份有限公司')
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(47, 84, 150)

    sub = doc.add_heading('GPU 供應鏈風險評估暨替代方案討論會議（逐字稿）', level=1)
    for run in sub.runs:
        run.font.color.rgb = RGBColor(47, 84, 150)

    info_lines = [
        ("日期", "2025年4月18日（星期五）10:00-11:30"),
        ("地點", "桃園總部 5F 會議室 C / Microsoft Teams 線上同步"),
        ("會議類型", "供應鏈風險評估專題會議"),
        ("會議目的", "評估 GPU Module 供應風險、審查替代方案、決定供應商策略"),
        ("主席", "Lisa Wang (資深供應鏈經理)"),
        ("會議背景", "NovaTech Q3 產能預訂 85%，交期延長至 10 週，加上美中關稅不確定性，需評估供應鏈風險並提出因應方案"),
    ]
    for label, val in info_lines:
        p = doc.add_paragraph()
        run1 = p.add_run(f"{label}：")
        run1.bold = True
        run2 = p.add_run(val)

    doc.add_paragraph()
    doc.add_heading('出席者', level=2)
    attendees = [
        ("Lisa Wang", "SCM", "供應鏈管理部", "資深供應鏈經理（主席）"),
        ("James Chen", "PM", "專案管理部", "資深專案經理"),
        ("Kevin Liu", "R&D", "研發設計部", "硬體設計主管"),
        ("Amy Hsu", "QA", "品質保證部", "品質工程經理"),
        ("Tom Lin", "IT", "資訊系統部", "IT 系統工程師"),
        ("Sarah Chang", "BPM", "流程管理部", "流程改善專員"),
        ("Robert Tsai", "Mfg", "製造部", "製造副總（列席）"),
    ]
    add_styled_table(doc, ["姓名", "部門代碼", "部門", "職稱"], attendees)

    doc.add_paragraph()
    doc.add_heading('會議議程', level=2)
    agenda = [
        ("1", "供應現況與風險概述", "Lisa Wang", "15 min"),
        ("2", "替代供應商評估結果", "Lisa Wang", "20 min"),
        ("3", "技術相容性評估", "Kevin Liu", "10 min"),
        ("4", "品質驗證需求", "Amy Hsu", "10 min"),
        ("5", "BOM 成本影響分析", "Lisa Wang", "15 min"),
        ("6", "IT 系統變更影響", "Tom Lin", "5 min"),
        ("7", "決議與行動計畫", "All", "15 min"),
    ]
    add_styled_table(doc, ["#", "議程項目", "報告人", "時間"], agenda, "C55A11")

    doc.add_paragraph()
    doc.add_heading('會議逐字紀錄', level=2)

    transcript = [
        ("10:02", "Lisa Wang", """各位早安，感謝大家撥冗參加這場專題會議，特別感謝 Robert 副總的列席。今天要討論的主題非常緊急——我們的 GPU Module 供應鏈面臨嚴重風險，需要在今天做出策略決定。

先快速回顧現況。我們目前 GPU Module NT-500 100% 依賴韓國的 NovaTech 供應。過去三個月發生了幾件讓我非常擔心的事：

第一，NovaTech 的 Q3 產能預訂率已經到 85%，他們在上週的電話中告訴我，如果我們不在 4 月底之前確認 Q3 的訂單數量，就無法保證產能分配。

第二，他們的標準交期從 8 週延長到 10 週。原因是他們的晶圓代工合作夥伴調整了產能配置，優先供應美國客戶——這跟美韓 CHIPS 法案有關。

第三，美中關稅的不確定性。我們原本考慮的中國供應商 ChipMax，面臨 25% 的額外關稅風險，而且出口管制的範圍一直在擴大。

簡單來說：我們現在是把所有雞蛋放在一個籃子裡，而這個籃子開始搖晃了。"""),

        ("10:08", "Robert Tsai", """Lisa，謝謝你的報告。我直接問——如果 NovaTech 突然無法供貨，我們的產線可以撐多久？"""),

        ("10:09", "Lisa Wang", """根據目前的安全庫存，ProServer X100 可以撐大約 2.5 週，ProServer X200 大約 3.3 週，CloudBook L15 大約 3 週。也就是說，如果 NovaTech 完全斷供，最快 2.5 週後 ProServer X100 的產線就會停。

以 ProServer X100 月產 2 萬台、每台售價約 USD 8,000 來算，停線一個月的營收損失大約是 USD 160M。這還不算客戶可能的罰款和信譽損失。"""),

        ("10:12", "Robert Tsai", """數字很驚人。那替代方案呢？"""),

        ("10:13", "Lisa Wang", """好的，進入第二個議程。我過去兩週密集評估了三家供應商，完整的評估報告已經放在 SharePoint 上的供應商評估表。我簡要報告結論。

現有供應商 NovaTech（韓國）：單價 285 美元，交期 8 週（目前延長到 10 週），品質評分 4.2，不良率 150 PPM。整體不錯，但單一供應商風險太高。

替代供應商一，ChipMax（中國）：單價最低 262 美元，但交期 12 週太長，品質只有 3.8，不良率 320 PPM 偏高。最大的問題是地緣政治風險——25% 關稅加上出口管制，我不建議把它作為主要替代方案。

替代供應商二，SiliconEdge（日本）：單價 298 美元最貴，但交期最短只要 6 週，品質最好 4.5，不良率最低 80 PPM。財務穩定性 A 級，有 IATF 16949 認證。主要客戶包含 HPE 和 Fujitsu，技術支援回應 2 小時。

我的加權評分結果：SiliconEdge 85 分最高，NovaTech 82 分，ChipMax 68 分。"""),

        ("10:18", "James Chen", """Lisa，如果選 SiliconEdge 當備援，每台成本增加 13 美元，一年 6 萬台就是 78 萬美元。這個數字客戶 HP 能接受嗎？我們的報價有這個 margin 嗎？"""),

        ("10:20", "Lisa Wang", """好問題。如果 100% 切到 SiliconEdge，年度成本增加 312 萬美元，這太高了。但我建議的方案不是 100% 切換。

我提出四個方案：

方案一，維持現狀：100% NovaTech。成本最低，但風險最高。
方案二，100% 切到 SiliconEdge。品質最好但成本增加 312 萬。
方案三，100% 切到 ChipMax。成本最低但品質和地緣政治風險太高。
方案四，我推薦的——雙供應商策略：NovaTech 60% + SiliconEdge 40%。

方案四的數字是這樣的：混合單價大約 290 美元，每台 GPU 成本 1,160 美元，比現在多 20 美元。年度增加成本 120 萬美元。但是——我們可以跟 NovaTech 用 SiliconEdge 做議價槓桿，他們已經報了 240K 量級的特殊價格 262-265 美元。如果我們拿到 265 美元，混合成本反而會下降。

另外，方案四有一個一次性的換線驗證成本 45 萬美元，包含 SiliconEdge 的樣品驗證、產線 ECN、和 ERP 系統變更。"""),

        ("10:25", "Kevin Liu", """技術面我來補充。我已經拿到 SiliconEdge SE-Pro 的 datasheet，跟 NovaTech NT-500 比較過了。

好消息：
1. 兩者的 PCIe Gen5 介面完全相容，pin-out 一致
2. 電氣規格都符合我們的設計需求
3. 功耗差異在 5% 以內，散熱方案不需要改

需要注意的：
1. SE-Pro 的 cold plate 安裝孔距是 78mm x 78mm，我們的 NT-500 是 76mm x 76mm。差 2mm，散熱模組的安裝板需要做一個 adapter bracket，估計開模費 1 萬 5 千美元，兩週可以完成。
2. BIOS 需要更新以支援 SE-Pro 的 device ID。R&D 評估大約 3 天的工作量。
3. BMC 韌體也要小改，大約 2 天。

總結：技術上可行，不是大改。主要是 adapter bracket 和韌體更新。整個技術驗證大概需要 4 週。"""),

        ("10:30", "Amy Hsu", """品質驗證的部分。如果要導入 SiliconEdge 作為第二供應商，我需要做以下驗證：

1. 進料檢驗標準制定：根據 SE-Pro 的規格建立 IQC 標準，大約 1 週
2. 功能驗證測試：在 ProServer X200 上做 72 小時 burn-in test，確認與 NT-500 的效能一致性
3. 可靠度驗證：HALT/HASS 測試、溫度循環測試、震動測試。這些需要 3 週
4. 相容性測試：確認 SE-Pro 和 NT-500 可以混用在同一台機器中（如果採用方案四的話）
5. 產線 First Article Inspection (FAI)：首批 50 台全檢

整個品質驗證流程需要 6 週。如果跟 Kevin 的技術驗證並行，最快 6 週可以完成。也就是說，如果今天決定啟動，最快 6 月初可以開始用 SiliconEdge 的料。

不良率方面，SiliconEdge 的歷史不良率是 80 PPM，比 NovaTech 的 150 PPM 好很多。品質上我反而支持導入 SiliconEdge。"""),

        ("10:35", "Tom Lin", """IT 這邊的影響比較小但不能忽略：

1. ERP 系統：需要新增 SiliconEdge 的供應商主檔和 SE-Pro 的料號。包含 BOM alternate 的設定——同一個 BOM 位置允許 NT-500 或 SE-Pro 兩種料。大約 2 天。

2. MES 系統：產線系統要加入 SE-Pro 的辨識邏輯，確保組裝和測試站能正確識別用的是哪家的 GPU。大約 3 天。

3. 報表系統：供應商績效報表和進料追蹤報表要新增 SiliconEdge 的數據。大約 1 天。

4. 一個重要提醒：如果採用混用策略，出貨給 HP 的時候要記錄每台用的是哪家 GPU，因為 HP 的維修備品管理需要這個資訊。

總工作量大約 5 個工作天，可以跟技術驗證並行。"""),

        ("10:38", "Sarah Chang", """流程管理的角度，我有幾點補充：

1. 供應商導入流程：CoolTech 的經驗告訴我們，新供應商導入大概需要 5-7 個工作天走完文件審查。SiliconEdge 的 ISO 和 IATF 認證都很齊全，應該不會有卡關。

2. 雙供應商管理 SOP：我們目前沒有「同一零件雙供應商」的標準作業流程。需要建立一套 SOP，包含：
   - 採購分配邏輯（60/40 怎麼分？按 PO 分還是按月分？）
   - 進料管理（IQC 是否對兩家用同一標準？）
   - 問題處理流程（如果一家出問題，怎麼快速把量轉給另一家？）

3. Gate Review 文件：這個供應商變更需要走 ECN 流程，包含設計變更、品質驗證報告、成本分析。我可以負責彙整所有文件。

這些流程建立大約需要 2 週。"""),

        ("10:42", "Robert Tsai", """聽完大家的報告，我的判斷是這樣的。

第一，單一供應商的風險確實不能再忽視了。160M 的營收損失，加上 HP 可能轉單的風險，遠超過每年多花 120 萬美元的成本。

第二，SiliconEdge 的品質比 NovaTech 更好，這是一個加分項目。如果我們跟 HP 說「我們導入了一個品質更好的備援供應商」，客戶應該會正面看待。

第三，ChipMax 我同意暫時不考慮，地緣政治風險太大。

我的決定：同意啟動方案四——雙供應商策略（NovaTech 60% + SiliconEdge 40%）。

Lisa，你說需要管理層核可的 45,000 美元驗證費用，我今天就核准。"""),

        ("10:45", "Lisa Wang", """謝謝 Robert 副總。那我統整一下今天的決議和行動計畫。"""),

        ("10:46", "James Chen", """等一下，我想確認一個問題。方案四實施後，如果 NovaTech 知道我們有第二供應商，會不會影響我們跟他們的關係？"""),

        ("10:47", "Lisa Wang", """好問題。我跟 David Park 私下聊過，NovaTech 其實理解客戶做 dual source 是業界常態。重點是我們怎麼跟他們溝通。我的策略是：

1. 不要讓 NovaTech 覺得我們要「拋棄」他們——強調他們仍然是主供應商（60%）
2. 把 dual source 定位為「分散我們自己的供應風險」，而不是「對他們不滿意」
3. 用這個來跟 NovaTech 談更好的價格——如果他們給 265 美元的年度合約價，我們可以維持 70/30 甚至 80/20 的分配

我計畫在下週跟 David Park 的會議中用這個角度來溝通。"""),

        ("10:50", "Robert Tsai", """Lisa 的溝通策略很好。那我們來總結行動項目吧。"""),

        ("10:51", "Lisa Wang", """好的。今天的決議和行動項目如下：

決議一：核准啟動雙供應商策略（NovaTech 60% + SiliconEdge 40%）
決議二：核准 SiliconEdge 驗證費用 USD 45,000
決議三：暫不考慮 ChipMax，待地緣政治局勢明朗後再評估

行動項目：

1. Lisa Wang — 本週內啟動 SiliconEdge 供應商導入流程 — 截止 4/25
2. Lisa Wang — 下週與 NovaTech David Park 溝通 dual source 策略 — 截止 4/25
3. Lisa Wang — 向 SiliconEdge 下 50 台驗證用料的 sample PO — 截止 4/22
4. Kevin Liu — 完成 SE-Pro 技術相容性評估報告 — 截止 4/25
5. Kevin Liu — 設計 adapter bracket 並開模 — 截止 5/9
6. Kevin Liu — BIOS 和 BMC 韌體更新 — 截止 5/2
7. Amy Hsu — 制定 SE-Pro IQC 進料標準 — 截止 4/25
8. Amy Hsu — 完成 SE-Pro 品質驗證計畫書 — 截止 4/30
9. Tom Lin — ERP 新增 SiliconEdge 供應商主檔和料號 — 截止 4/25
10. Tom Lin — MES 系統新增 SE-Pro 辨識邏輯 — 截止 5/2
11. Sarah Chang — 建立雙供應商管理 SOP — 截止 5/9
12. Sarah Chang — 準備 ECN 文件包 — 截止 5/16
13. Lisa Wang — 完成供應鏈風險分析報告（含方案四詳細分析）— 截止 4/25

目標時程：6 月初完成所有驗證，6 月中開始使用 SiliconEdge 料件量產。

下次會議：5 月 2 號，同一時間，追蹤行動項目進度。

大家有沒有遺漏的？"""),

        ("10:56", "Amy Hsu", """補充一點——如果 sample 料到貨後發現品質問題需要來回，驗證時程可能會拉長。建議 Lisa 在跟 SiliconEdge 下 sample PO 的時候，同時要求他們提供完整的出廠檢驗報告和 FA 樣品，這樣我這邊可以平行作業。"""),

        ("10:57", "Lisa Wang", """好的，我把這點加入行動項目。還有其他問題嗎？"""),

        ("10:58", "Robert Tsai", """沒有了。Lisa 你辛苦了，這份評估做得很完整。大家下去執行吧。散會。"""),
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

    # Decision & Action Summary
    doc.add_page_break()
    doc.add_heading('決議摘要', level=2)
    decisions = [
        ("D-001", "雙供應商策略", "核准啟動 NovaTech 60% + SiliconEdge 40%", "Robert Tsai 核准"),
        ("D-002", "驗證費用", "核准 SiliconEdge 驗證費用 USD 45,000", "Robert Tsai 核准"),
        ("D-003", "ChipMax 暫緩", "暫不考慮 ChipMax，待地緣政治明朗後再評估", "全員同意"),
    ]
    add_styled_table(doc, ["編號", "決議項目", "決議內容", "核准者"], decisions, "548235")

    doc.add_paragraph()
    doc.add_heading('行動項目清單', level=2)
    actions = [
        ("A-101", "Lisa Wang", "啟動 SiliconEdge 供應商導入流程", "4/25"),
        ("A-102", "Lisa Wang", "與 NovaTech 溝通 dual source 策略", "4/25"),
        ("A-103", "Lisa Wang", "向 SiliconEdge 下 sample PO (50台驗證料)", "4/22"),
        ("A-104", "Kevin Liu", "SE-Pro 技術相容性評估報告", "4/25"),
        ("A-105", "Kevin Liu", "設計 adapter bracket 並開模", "5/9"),
        ("A-106", "Kevin Liu", "BIOS + BMC 韌體更新", "5/2"),
        ("A-107", "Amy Hsu", "制定 SE-Pro IQC 進料標準", "4/25"),
        ("A-108", "Amy Hsu", "完成 SE-Pro 品質驗證計畫書", "4/30"),
        ("A-109", "Tom Lin", "ERP 新增供應商主檔和料號", "4/25"),
        ("A-110", "Tom Lin", "MES 系統新增 SE-Pro 辨識邏輯", "5/2"),
        ("A-111", "Sarah Chang", "建立雙供應商管理 SOP", "5/9"),
        ("A-112", "Sarah Chang", "準備 ECN 文件包", "5/16"),
        ("A-113", "Lisa Wang", "完成供應鏈風險分析報告", "4/25"),
    ]
    add_styled_table(doc, ["編號", "負責人", "行動項目", "截止日"], actions)

    path = os.path.join(OUTPUT_DIR, "英業達_供應鏈風險會議逐字稿.docx")
    doc.save(path)
    print(f"✅ Created: {path}")

if __name__ == "__main__":
    create_scm_meeting_transcript()
