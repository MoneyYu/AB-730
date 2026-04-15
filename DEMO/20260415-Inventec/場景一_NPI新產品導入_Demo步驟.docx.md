# 場景一：新產品導入 NPI (New Product Introduction)
## ProServer X200 AI 推論伺服器開發專案

---

## 場景背景（開場白）

> 各位學員，歡迎來到場景一。英業達接到 HP Enterprise 的新一代 AI 推論伺服器訂單——ProServer X200。客戶要求 Q4 2025 量產，專案經理 James Chen 需要統籌 R&D、SCM、QA、IT、BPM 五個部門，在 7 個月內完成從 Kick-off 到量產的全部流程。
>
> 在這個場景中，我們將體驗 PM 如何利用 Copilot 的各種功能，快速產出會議紀錄、追蹤進度、彙整週報——所有 PM 每天都在做的「產出交付物」工作。

**涉及角色**：PM（主導）、SCM、R&D、QA、IT、BPM
**使用的範例檔案**：
- 英業達_NPI_Kickoff_會議逐字稿.docx
- 英業達_NPI_設計審查會議逐字稿.docx
- 英業達_NPI_ProServerX200_專案追蹤.xlsx
- 英業達_NPI週報範本.docx
- 英業達_會議總結範本.docx

> 📎 **Teams 會議完整生命週期 Demo**（會前→會中→會後多角度摘要）已獨立為專用檔案：
> **Teams會議全流程_會前會中會後_Demo步驟.docx.md**
> 建議在 Demo 1 和 Demo 2 之間穿插使用，或作為 Teams 專題教學的獨立模組（45-60 分鐘）。

---

## Demo 1：Teams Facilitator — NPI Kick-off 即時會議協作
**交付物**：即時共編筆記 + 議程追蹤 + 行動項目清單
**預計時間**：20-25 分鐘
**對應模組**：模組 5（管理會議與協作）

### 教學重點先說明（2 分鐘）

在開始 Demo 前，先向學員說明 **Copilot in Teams** 和 **Facilitator** 的核心差異：

| | Copilot in Teams | Facilitator |
|---|---|---|
| 可見性 | 🔒 私密 — 只有自己看到 | 👁️ 公開 — 全員都看到 |
| 行為模式 | 像個人助理，私下幫你 | 像一個 AI 與會者，參與討論 |
| 適用場景 | 個人筆記、私下提問 | 團隊協作、共同決策 |
| 儲存方式 | 個人聊天紀錄 | .loop 檔案存在 OneDrive/Meetings |
| 互動方式 | 在 Copilot 側邊面板提問 | 在會議聊天中 @Facilitator |

### 步驟 1：開啟會議並啟用 Facilitator（3 分鐘）

1. 開啟 Microsoft Teams，進入一個已排程的會議（或開啟一個即時會議）
2. 加入會議後，在會議工具列找到 **「Notes」** 按鈕，點擊開啟
3. 確認 Facilitator 已自動加入會議（會看到 Facilitator 出現在參與者清單中）
4. 如果沒有看到 Facilitator，在會議聊天中輸入 `@Facilitator` 並送出任意訊息來觸發

> 💡 **教學 Tip**：告訴學員 Facilitator 需要管理員在 Teams Admin Center 中啟用。如果學員的租戶沒有啟用，可以用截圖或錄影展示。

### 步驟 2：使用 Facilitator 追蹤議程（5 分鐘）

1. 在會議聊天中輸入以下提示：
   ```
   @Facilitator 請把以下議程加入今天的追蹤清單：
   1. 各部門分工確認
   2. 散熱方案討論
   3. GPU 供應商評估
   4. 測試計畫時程
   5. IT 系統準備
   ```
2. Facilitator 會在會議聊天中回應，並在 Notes 面板中建立議程追蹤
3. 觀察 **Visual Timeline Markers**：Facilitator 會標記目前討論到哪個議題
4. 示範 Facilitator 的 **midway reminder**：當會議進行到一半時，Facilitator 會提醒剩餘時間

> 💡 **教學 Tip**：強調這是「公開的」——所有與會者都能看到議程追蹤，不需要有人手動記錄「現在討論到第幾個議題了」。

### 步驟 3：即時記錄討論重點（5 分鐘）

1. 模擬會議討論（可以請學員扮演不同角色，或由老師口述會議內容）
2. 參考 **英業達_NPI_Kickoff_會議逐字稿.docx** 的內容進行模擬
3. 觀察 **Notes Tab** 的即時筆記：
   - 點擊會議上方的 **Notes** 分頁
   - 看到 Facilitator 自動記錄的討論重點
   - 所有參與者都可以即時編輯這份筆記（Loop 格式）
4. 在聊天中輸入：
   ```
   @Facilitator 剛才 Kevin 提到散熱方案要改成液冷，預計延遲 3 週。請把這個決議記錄下來。
   ```
5. Facilitator 會將此決議加入 Notes

> 💡 **教學 Tip**：強調 Notes 是共編的——任何人都可以修改、補充。這跟 Copilot in Teams 的私人筆記完全不同。

### 步驟 4：請 Facilitator 彙整行動項目（5 分鐘）

1. 在會議聊天中輸入：
   ```
   @Facilitator 請列出目前為止的行動項目，包含負責人和截止日期
   ```
2. Facilitator 會在聊天中回應一份結構化的行動清單（全員可見），例如：
   - Kevin Liu — 散熱方案評估與修改圖面 — 4/11
   - Lisa Wang — GPU 供應商評估報告 + 初版 BOM — 4/4
   - Amy Hsu — DVT 測試計畫 — BOM 完成後 2 週
   - Tom Lin — ERP 料號建檔 — BOM 完成後 3 天
   - Sarah Chang — 建立 Teams 頻道 — 今天

3. **(Preview 功能)** 示範 Facilitator 與 Planner 的整合：
   ```
   @Facilitator 請把這些行動項目建立到 Planner 中
   ```
   - Facilitator 會將行動項目自動建立為 Planner 任務
   - 每個任務包含負責人、截止日期

> 💡 **教學 Tip**：Planner 整合目前是 Preview 功能。如果租戶沒有啟用 Preview，可以跳過此步驟，改用截圖展示。

### 步驟 5：會後檢視 Facilitator 產出（3 分鐘）

1. 會議結束後，展示 Facilitator 產生的 .loop 檔案：
   - 路徑：OneDrive > Meetings 資料夾
   - 包含完整的會議筆記、議程追蹤結果、行動項目
2. 這份 .loop 檔案可以直接分享給團隊
3. 也可以在 Teams 頻道中貼上這份 Loop 元件，讓團隊持續編輯

> 🎯 **本 Demo 學員 Takeaway**：
> - Facilitator = 公開的 AI 會議助理，適合團隊協作
> - 即時議程追蹤 + AI 筆記 + 行動項目一次搞定
> - 會後自動產出 .loop 文件，團隊可持續協作

---

## Demo 2：Teams Notes + AI Summary — NPI 設計審查會後摘要
**交付物**：AI 會議摘要 + 結構化筆記 + Copilot Pages 協作空間
**預計時間**：20-25 分鐘
**對應模組**：模組 5（管理會議與協作）

### 教學重點先說明（2 分鐘）

這個 Demo 展示的是 **會議結束後** 的 Copilot 使用情境，跟 Demo 1 的「即時」不同。

| | 即時 (Facilitator/Notes) | 會後 (AI Summary/Copilot) |
|---|---|---|
| 時間點 | 會議進行中 | 會議結束後 |
| 資料來源 | 即時聽到的內容 | 完整的會議轉錄稿 + 聊天紀錄 |
| 精準度 | 即時記錄，可能遺漏 | 回顧整場會議，更完整 |
| 用途 | 當下協作 | 事後整理、報告、追蹤 |

### 前置準備

- **方式 A**：將 **英業達_NPI_設計審查會議逐字稿.docx** 的內容貼入 Copilot Chat，使用以下提示讓 Copilot 分析這份逐字稿（不需要實際 Teams 會議錄製）
- **方式 B**：若學員環境有實際會議紀錄，可直接使用 Teams Recap 功能
- 確認 **英業達_NPI_設計審查會議逐字稿.docx** 已上傳至 OneDrive

### 步驟 1：開啟會後 Recap（3 分鐘）

1. 在 Teams 行事曆中找到已結束的會議
2. 點擊會議，進入 **Recap** 頁面
3. 觀察 Copilot 自動產生的內容：
   - 📝 **AI Summary**：會議摘要
   - 🎯 **Action Items**：行動項目
   - 📌 **Mentions**：被提到的人名
   - 🕐 **Topics**：依時間軸標記的討論主題

> 💡 **教學 Tip**：AI Summary 結合了 transcript（語音轉錄）和 meeting chat（文字聊天）的內容，比單純的語音轉錄更準確。

### 步驟 2：使用 Copilot in Teams 深入提問（8 分鐘）

1. 在 Recap 頁面右側開啟 **Copilot** 面板
2. 依序輸入以下提示（每個都展示結果後再進下一個）：

   **提示 A — 完整摘要**：
   ```
   請產生這次設計審查會議的完整摘要，格式包含：
   1. 會議目的
   2. 關鍵討論點
   3. 技術決策
   4. 行動項目
   ```
   ➡️ 展示 Copilot 產生的結構化摘要

   **提示 B — 技術決策提取**：
   ```
   這次會議做了哪些技術決策？請列出決策內容、決策原因和負責人
   ```
   ➡️ 展示 Copilot 精準提取出：液冷方案確認、CoolTech 供應商選擇、GPU 供應商策略

   **提示 C — 風險識別**：
   ```
   這次會議中提到了哪些風險或 concern？每個風險的嚴重程度如何？
   ```
   ➡️ 展示 Copilot 從會議內容中識別出的風險項目

   **提示 D — 特定人物追蹤**：
   ```
   Lisa Wang 在會議中報告了哪些內容？她有哪些待辦事項？
   ```
   ➡️ 展示依人員篩選的追蹤功能

3. 注意 Copilot 在每次回答後會**建議 Follow-up Prompts**（後續追問），示範點擊其中一個

> 💡 **教學 Tip**：強調 Copilot in Teams 的回答是**私人的**——只有你自己看到。適合在會後私下整理筆記，不用擔心問了「笨問題」被其他人看到。

### 步驟 3：將摘要轉到 Copilot Pages（5 分鐘）

1. 在 Copilot 的回答中，找到 **「Edit in Pages」** 或 **「Open in Copilot Pages」** 按鈕
2. 點擊後，Copilot 會將摘要內容轉移到一個新的 Copilot Pages 頁面
3. 在 Copilot Pages 中：
   - 展示可以邀請團隊成員一起編輯
   - 點擊 **Share** → 輸入團隊成員名稱
   - 示範即時共編：你修改一段，其他人可以即時看到
4. 在 Pages 中進一步使用 Copilot：
   ```
   請根據上面的行動項目，產生一份格式化的行動追蹤表，包含：編號、負責人、項目、截止日、狀態
   ```
5. Copilot Pages 會產生一個互動式表格

> 💡 **教學 Tip**：Copilot Pages 是團隊協作的「中繼站」——從 Copilot 產出的內容開始，再由團隊共同編輯、補充、精煉。

### 步驟 4：展示 Notes 功能（3 分鐘）

1. 回到 Teams 會議紀錄
2. 點擊 **Notes** 分頁
3. 展示 AI 在會議期間自動產生的筆記（Loop 格式）：
   - 自動分段落
   - 標記重要決策
   - 列出行動項目
4. 說明 Notes 和 AI Summary 的差異：
   - **Notes**：會議進行中即時產生，可共編（用 Facilitator 增強）
   - **AI Summary**：會議結束後產生，基於完整轉錄稿，更完整

> 🎯 **本 Demo 學員 Takeaway**：
> - 會後用 Copilot in Teams 做個人摘要（私密）
> - 用 Copilot Pages 把好的摘要分享給團隊（協作）
> - Notes vs AI Summary：即時 vs 事後，各有適用場景
> - Follow-up Prompts：Copilot 會引導你問更深入的問題

---

## Demo 3：Teams Copilot Chat — 跨頻道進度彙整
**交付物**：每週狀態報告 + Notebook 持續追蹤
**預計時間**：15-20 分鐘
**對應模組**：模組 2（Copilot Chat）+ 模組 5

### 前置準備

- 在 Teams 中建立 **#NPI-ProServerX200** 頻道（若無法建立，使用任一現有專案頻道代替，Demo 時將頻道名稱替換即可）
- 將 **英業達_NPI_ProServerX200_專案追蹤.xlsx** 上傳至 OneDrive 或 SharePoint
- 事先在頻道中貼入 3-5 則模擬訊息（可從 **英業達_NPI_Kickoff_會議逐字稿.docx** 中擷取各角色的發言作為頻道訊息素材）

### 步驟 1：開啟 Copilot Chat（Work 模式）（2 分鐘）

1. 開啟 Microsoft 365 Copilot（m365.cloud.microsoft 或 Teams 側邊欄的 Copilot）
2. 確認切換到 **Work** 模式（不是 Web 模式）
3. 說明 Work 模式的差異：**可以存取你的 Teams 頻道、郵件、檔案等組織內部資料**

### 步驟 2：彙整頻道討論（5 分鐘）

1. 在 Copilot Chat 中輸入：
   ```
   彙整過去一週 Teams 頻道 #NPI-ProServerX200 中的所有討論，按以下分類整理：
   1. 進度更新
   2. 風險與議題
   3. 決議事項
   4. 待辦行動
   ```
2. 展示 Copilot 的回應——它會從頻道訊息中提取並分類資訊
3. 如果頻道資料不足，可以改用以下替代提示：
   ```
   根據我最近參加的 ProServer X200 相關會議和郵件，彙整過去一週的專案進度更新
   ```

> 💡 **教學 Tip**：Work 模式的 Copilot 會搜尋你有權限存取的所有 M365 資料——Teams 訊息、郵件、OneDrive 檔案、SharePoint 文件。它是跨 App 的。

### 步驟 3：產生格式化週報（5 分鐘）

1. 根據上一步的彙整，繼續輸入：
   ```
   請根據以上彙整，產生一份 NPI 週報，格式如下：
   
   一、整體狀態（用紅黃綠燈表示）
   二、本週進度摘要（按部門分列）
   三、風險與議題（標註風險等級）
   四、行動項目追蹤（表格：編號/負責人/項目/到期日/狀態）
   五、下週計畫
   六、需要管理層決策的事項
   
   語氣專業，適合發給管理層。
   ```
2. 展示 Copilot 產出的格式化週報
3. 對照 **英業達_NPI週報範本.docx**，比較 Copilot 產出的格式是否符合需求
4. 如果有需要調整的地方，示範追問：
   ```
   請把風險等級用 emoji 標示（🔴嚴重 🟡中等 🟢低），並在每個風險後面加上緩解措施
   ```

### 步驟 4：加入 Notebook（3 分鐘）

1. 在 Copilot Chat 的回答中，找到 **「Add to Notebook」** 按鈕
2. 點擊後，這份週報會被保存到 Copilot Notebook
3. 說明 Notebook 的用途：
   - 可以持續累積每週的進度彙整
   - 下次寫週報時可以引用上週的 Notebook 內容
   - Notebook 中的內容可以進一步編輯和分享

### 步驟 5：設定排程提示（3 分鐘）

1. 展示如何設定排程提示（Scheduled Prompts）：
   ```
   每週五下午 4 點，自動彙整 #NPI-ProServerX200 頻道本週的進度更新，格式化為週報
   ```
2. 在 Copilot Chat 中找到 **排程提示** 的設定介面
3. 設定頻率：每週五 16:00
4. 說明排程提示的效果：每週五自動執行，結果會出現在你的 Copilot Chat 中

> 💡 **教學 Tip**：排程提示是把重複性的彙整工作自動化。PM 不用每週五手動去翻頻道訊息，Copilot 會自動幫你做好。

> 🎯 **本 Demo 學員 Takeaway**：
> - Copilot Chat (Work) 可以跨頻道、跨 App 搜尋組織內部資料
> - 用追問來迭代優化格式和內容
> - Notebook 保存重要的 Copilot 產出
> - 排程提示自動化每週彙整

---

## Demo 5：Teams Channel Agent — 頻道智慧管理
**交付物**：頻道重點摘要 + deadline 追蹤
**預計時間**：10 分鐘
**對應模組**：模組 2（Copilot Chat）

### 步驟 1：在頻道中啟用 Channel Agent（2 分鐘）

1. 進入 Teams 頻道 #NPI-ProServerX200
2. 在頻道中找到 Copilot 圖示或在訊息框中輸入 `@Copilot`
3. 說明 Channel Agent 的作用：專門幫你管理和追蹤頻道中的對話

### 步驟 2：使用 Channel Agent 追蹤重點（5 分鐘）

1. 輸入以下提示：
   ```
   @Copilot 摘要這個頻道過去一週的重要討論，特別是有提到 deadline 或截止日期的訊息
   ```
2. 展示 Channel Agent 的回應——它會標記出重要的截止日期和行動項目

3. 進一步提問：
   ```
   @Copilot 這個頻道中有哪些尚未回覆或需要跟進的問題？
   ```
4. Channel Agent 會找出沒有被回覆的問題或需要 follow-up 的討論串

### 步驟 3：產出頻道週報摘要（3 分鐘）

1. 輸入：
   ```
   @Copilot 產生這個頻道的每週摘要，包含：
   - 本週重要決策
   - 待辦事項和負責人
   - 需要注意的風險
   ```
2. 展示回應，說明這可以直接作為團隊更新的素材

> 🎯 **本 Demo 學員 Takeaway**：
> - Channel Agent 專注於單一頻道的智慧管理
> - 自動追蹤 deadline、找出未回覆的問題
> - 適合 BPM/PM 角色用於頻道管理

---

## Demo 13：Word — NPI 週報自動產生
**交付物**：格式化 NPI 週報文件
**預計時間**：15-20 分鐘
**對應模組**：模組 6（定義 Copilot 角色）

### 教學重點先說明（2 分鐘）

這個 Demo 是場景一的「收尾」——展示如何把 Teams 和 Outlook 彙整的素材，透過 Word Copilot 產出正式的週報文件。核心概念是 **跨 App 資料流**：

```
Teams 頻道彙整 → Outlook 郵件補充 → Word 產出正式文件
```

### 前置準備

- 開啟 **英業達_NPI週報範本.docx** 作為參考
- 確認 OneDrive 中有 **英業達_NPI_ProServerX200_專案追蹤.xlsx**

### 步驟 1：在 Word 中啟動 Copilot 草擬（5 分鐘）

1. 開啟 Microsoft Word（新文件）
2. 點擊 **Copilot** 圖示（或在空白處按下 Draft with Copilot）
3. 輸入以下提示（注意用 `/` 引用檔案）：
   ```
   請根據 /英業達_NPI_ProServerX200_專案追蹤.xlsx 中 W16 的進度資料，
   撰寫一份 NPI 週報，格式包含：
   
   一、整體狀態（紅黃綠燈）
   二、本週進度摘要（按部門：R&D、SCM、QA、IT、BPM）
   三、風險與議題（標註風險等級和緩解措施）
   四、行動項目追蹤（表格格式）
   五、下週計畫
   六、需要管理層決策的事項
   
   語氣專業，適合發給製造副總。
   ```
4. 等待 Copilot 生成草稿
5. 展示生成結果，與 **英業達_NPI週報範本.docx** 對照

> 💡 **教學 Tip**：用 `/` 可以引用 OneDrive/SharePoint 中的檔案作為 Copilot 的上下文。這是讓 Copilot 產出更精準內容的關鍵技巧。

### 步驟 2：使用 Copilot 精煉內容（5 分鐘）

1. 選取「風險與議題」的段落
2. 點擊 Copilot 的 **Rewrite** 功能
3. 選擇不同選項展示效果：
   - **Make it more concise**：讓內容更精簡
   - **Make it more professional**：讓語氣更專業
   - **Change tone to formal**：正式語氣
4. 展示前後對比

5. 選取管理層決策事項的段落，使用 Copilot：
   ```
   請把這段改寫得更有說服力，加入具體數字來支持建議。
   特別強調雙供應商策略的長期效益和風險降低。
   ```
6. 展示 Copilot 如何加入量化數據讓建議更有力

### 步驟 3：產生執行摘要（3 分鐘）

1. 將游標移到文件最前面
2. 使用 Copilot：
   ```
   請在文件開頭加入一段 150 字的執行摘要 (Executive Summary)，
   涵蓋本週整體狀態、最關鍵的風險、需要管理層立即決策的事項。
   ```
3. 展示 Copilot 產生的執行摘要

### 步驟 4：展示跨 App 資料流（3 分鐘）

1. 總結這個 Demo 的跨 App 工作流程：
   ```
   Teams 頻道 (#NPI-ProServerX200) 
       ↓ Copilot Chat 彙整
   Outlook 郵件 (供應商溝通、客戶更新) 
       ↓ Copilot 摘要
   Excel 專案追蹤表
       ↓ Word Copilot 引用 (/)
   Word 正式週報文件
       ↓ 完成！
   ```
2. 強調：整個流程從「資料散落在各處」到「正式文件產出」，Copilot 就是那個把所有碎片串起來的工具

> 🎯 **本 Demo 學員 Takeaway**：
> - Word Copilot 用 `/` 引用檔案，讓產出基於真實資料
> - Rewrite 功能快速調整語氣、長度、專業度
> - 跨 App 資料流是 Copilot 最大的價值——不是單一 App 的功能，而是整個工作流程的串接

---

## 加碼實戰情境：PM 的一天（貼近日常的 Copilot 應用）

以下是 NPI PM 在英業達實際會遇到的日常情境，每個都有完整的 Copilot 操作步驟。老師可以穿插在 Demo 之間，或作為學員的進階練習。

---

### 實戰情境 A：客戶臨時變更規格（Spec Change Request）

> **背景**：星期三下午 4 點，HP 的 PM David 在 Teams 私訊你：「Hi James，我們的 AI team 做了最新的 benchmark，需要把 GPU 從 4 張改成 6 張，同時記憶體從 512GB 升到 768GB。請評估對時程和成本的影響，我需要在明天 COB 前收到 Impact Assessment。」
>
> **痛點**：你只有不到 24 小時，需要快速計算 BOM 影響、確認供應商產能、通知所有相關部門、寫出正式的影響評估報告。

**步驟 1 — 用 Copilot Chat 快速評估影響**

在 Copilot Chat (Work) 中輸入：
```
我收到客戶的規格變更請求：ProServer X200 需要從 4 張 GPU 改成 6 張，記憶體從 512GB (16x32GB) 改成 768GB (24x32GB)。

請根據 /英業達_BOM成本分析.xlsx 的現有 BOM 資料，幫我計算：
1. BOM 成本增加多少？（GPU 單價 $285/張，DRAM 單價 $42/條）
2. 這個變更對機殼、散熱、電源有什麼連鎖影響？
3. 2U 機殼能容納 6 張 GPU 嗎？可能需要改成幾 U？

用表格呈現，並列出需要立即確認的技術問題。
```

**步驟 2 — 用 Copilot 在 Teams 通知相關部門**

在 #NPI-ProServerX200 頻道中輸入：
```
@Kevin Liu @Lisa Wang @Amy Hsu

⚠️ 緊急：客戶規格變更通知

HP 要求以下變更：
- GPU: 4 張 → 6 張
- DRAM: 512GB → 768GB (16條 → 24條)

需要各部門在明天中午前回覆影響評估：
- R&D: 機殼尺寸是否需要從 2U 改成 4U？散熱方案是否需要重新設計？
- SCM: GPU 和 DRAM 的額外採購量能否在現有 lead time 內取得？
- QA: DVT 測試計畫是否需要重新制定？

請在本則訊息下方回覆，我明天下午要交 Impact Assessment 給客戶。
```

**步驟 3 — 用 Word Copilot 產出 Impact Assessment**

等各部門在 Teams 回覆後，在 Word 中輸入：
```
請撰寫一份客戶規格變更影響評估報告 (Engineering Change Impact Assessment)，格式如下：

1. 變更摘要：GPU 4→6, DRAM 512→768GB
2. BOM 成本影響：每台增加 USD ___（GPU $570 + DRAM $336 = $906/台）
3. 時程影響：機構是否需重新設計？散熱方案？
4. 供應鏈影響：額外料件的交期和產能
5. 測試影響：測試計畫是否需更新
6. 風險評估
7. 建議與替代方案

語氣專業，適合發給客戶的 PM。英文撰寫。
引用 Teams 頻道 #NPI-ProServerX200 中各部門的回覆作為評估依據。
```

> 💡 **教學 Tip**：這就是 PM 的真實一天——突發的客戶需求、跨部門協調、限時交付。Copilot 幫你把原本需要半天的作業壓縮到 2 小時。

---

### 實戰情境 B：EVT 測試失敗的緊急應對

> **背景**：QA Amy 在 Teams 發訊息：「James，EVT 散熱測試不過。GPU junction temp 在滿載時到了 88°C，超過客戶的 85°C 上限。需要開 ECO (Engineering Change Order)。」時間是星期五下午 3 點，下週一客戶要來看 EVT 結果。
>
> **痛點**：你要在週末前搞定：原因分析、ECO 文件、通知客戶時程延遲、重新安排會議。

**步驟 1 — 用 Copilot 快速產生 ECO 文件**

在 Word 中使用 Copilot：
```
請撰寫一份 ECO (Engineering Change Order) 文件，格式如下：

ECO 編號：ECO-PSX200-003
變更名稱：散熱模組液冷管路流量優化
變更原因：EVT 散熱測試中，GPU junction temp 於滿載時達 88°C，超過客戶要求上限 85°C
根因分析：Cold plate 進液流量不足，每 GPU 實測 0.6 L/min（設計值 0.8 L/min），
          推測為管路佈局轉彎處流阻過大
變更內容：
- 增加管路直徑從 8mm 至 10mm
- 減少管路轉彎處從 6 處降為 4 處
- CDU 泵浦壓力提升 15%
影響評估：
- 時程影響：延遲約 2 週（修改模具 + 重新驗證）
- 成本影響：Cold plate 模具修改費 USD 15,000
- 品質影響：預期 junction temp 可降至 72-75°C

需要核准：R&D 主管 / PM / 客戶端 PM
```

**步驟 2 — 用 Outlook Copilot 通知客戶延遲**

在 Outlook 中草擬郵件：
```
請草擬一封郵件給 HP 的 PM David Chen，通知 EVT 散熱測試需要做 ECO，預計延遲 2 週。
內容包含：
1. 坦誠說明測試結果（88°C vs 85°C 上限）
2. 已識別的根因（管路流量不足）
3. 修改方案和預期效果（72-75°C）
4. 修改時程：模具修改 1 週 + 重新驗證 1 週
5. 請求將下週一的 EVT review 會議延後到兩週後

語氣坦誠但有信心，強調我們快速反應且方案明確。英文。
```

**步驟 3 — 用 Copilot 在 Teams 安排緊急會議**

在 Copilot Chat 中：
```
幫我安排一場緊急會議：
- 主題：ProServer X200 ECO-003 散熱修改方案討論
- 參與者：Kevin Liu, David Wu, Lisa Wang, Amy Hsu
- 時間：今天下午 5 點（30 分鐘）
- 議程：1) 測試數據回顧 2) 修改方案確認 3) 時程影響 4) 客戶溝通策略
```

> 💡 **教學 Tip**：測試失敗是 NPI 最常見的「突發狀況」。PM 的價值不在於技術解決，而在於**快速協調、產出文件、管理客戶期望**。Copilot 幫你把這些行政工作加速 3-5 倍。

---

### 實戰情境 C：主管問你「專案整體狀況如何？」

> **背景**：星期一早上 9 點，製造副總 Robert Tsai 在走廊上遇到你：「James，ProServer X200 目前狀況如何？下午有個 director meeting 我要 update 一下。給我一頁摘要就好。」
>
> **痛點**：你的專案資訊散落在 Teams 頻道、Outlook 郵件、Excel 追蹤表。你需要在 30 分鐘內產出一頁 executive summary。

**用 Copilot Chat (Work) 一分鐘產出摘要**

```
幫我產生 ProServer X200 專案的 one-page executive summary，
基於以下來源：
- Teams 頻道 #NPI-ProServerX200 最近的討論
- 我最近跟 HP 客戶和供應商的郵件
- /英業達_NPI_ProServerX200_專案追蹤.xlsx

格式：
- 整體狀態：🟡/🟢/🔴
- 進度摘要（3 句話）
- Top 3 風險
- 需要副總協助的事項
- 下一個里程碑和日期

控制在一頁 A4 以內。語氣精簡果斷。
```

> 💡 **教學 Tip**：高階主管只需要「一頁」和「三句話」。Copilot 幫你從散亂的資料中萃取出管理層需要的精華。學員可以把這個當成每天的例行動作——早上花 2 分鐘用 Copilot 產出「今天的專案狀態摘要」。

---

### 實戰情境 D：你剛接手同事離職留下的專案

> **背景**：原本負責 CloudBook L15 的 PM 上個月離職了，你被指派接手。但你對這個專案一無所知——他留下了一堆 Teams 頻道訊息、半年的郵件、十幾份 Excel 追蹤表。
>
> **痛點**：你需要在一週內搞清楚這個專案的全貌、所有待辦事項、以及潛在的地雷。

**步驟 1 — 用 Copilot Chat (Work) 做專案摘要**

```
我剛接手 CloudBook L15 專案，之前的 PM 已經離職。
請幫我彙整以下資料，產生一份「專案交接摘要」：

1. 搜尋 Teams 頻道 #CloudBook-L15 過去 3 個月的討論，摘要重點
2. 搜尋我 Outlook 中與 CloudBook L15 相關的郵件
3. 列出所有進行中的行動項目、負責人、截止日
4. 識別任何「看起來像是問題但沒有被解決」的議題
5. 列出所有跟這個專案相關的供應商和客戶聯繫人

格式化為「專案交接清單」。
```

**步驟 2 — 用 Copilot 在 Teams 向團隊自我介紹**

```
請幫我在 #CloudBook-L15 頻道發一則訊息：
- 自我介紹：我是新的 PM James Chen
- 感謝前任 PM 的工作
- 說明我已經閱讀了過去三個月的討論（用 Copilot 彙整）
- 列出我理解的目前 Top 3 待辦
- 請團隊成員確認這些理解是否正確
- 邀請大家在本週五跟我做一次 30 分鐘的 sync up

語氣謙虛但積極，展現我有做功課。
```

> 💡 **教學 Tip**：「接手別人的專案」是每個人都會遇到的情境。Copilot 的跨 App 搜尋能力在這裡發揮最大價值——3 個月的資訊碎片，10 分鐘就能掌握全貌。

---

### 實戰情境 E：多專案週會的高效準備

> **背景**：你同時管理 3 個 NPI 專案（ProServer X200、CloudBook L16、IoT Gateway G4）。每週一早上 10 點有「NPI 週會」，需要向 Director 報告所有專案的狀態。你有 30 分鐘準備。
>
> **痛點**：三個專案的資訊分散在不同的 Teams 頻道、不同的 Excel、不同的郵件串。

**用 Copilot Chat 一次彙整三個專案**

```
請分別搜尋以下三個 Teams 頻道過去一週的討論，
並產生一份統一的「多專案狀態更新」：

1. #NPI-ProServerX200
2. #NPI-CloudBookL16
3. #NPI-IoTGatewayG4

每個專案用以下格式：
---
專案名稱 | 🟢🟡🔴 狀態
本週重點（3 句話以內）
風險 / 議題
下週關鍵里程碑
---

最後加一個「需要 Director 關注的事項」彙總區塊。
控制在 2 頁 A4 以內。
```

> 💡 **教學 Tip**：管多個專案的 PM 每天都在做「資訊彙整」。Copilot 最大的價值不是幫你寫漂亮的文字，而是幫你**從噪音中提取信號**。

---

### 實戰情境 F：週五下午的「本週回顧+下週預告」Teams 訊息

> **背景**：很多 PM 有個習慣——每週五下午在 Teams 頻道發一則「本週回顧+下週預告」，讓團隊所有人都知道大方向。但寫這則訊息通常要花 30-45 分鐘回顧整週的事情。
>
> **痛點**：你需要快速回顧整週的會議、郵件、頻道討論，然後寫一則簡潔但完整的更新。

**用排程提示自動化**

在 Copilot Chat 中設定排程提示：
```
每週五下午 3:30，請自動執行以下動作：

1. 搜尋 Teams 頻道 #NPI-ProServerX200 本週的所有訊息
2. 搜尋我本週所有標記為「ProServer X200」的郵件
3. 搜尋我本週參加的所有含「ProServer」或「PSX200」的會議

產生以下格式的訊息草稿：

📋 ProServer X200 週回顧 (W16)

✅ 本週完成：
- [列出 3-5 項]

🔄 進行中：
- [列出 2-3 項]

⚠️ 需要注意：
- [列出風險或問題]

📅 下週重點：
- [列出 2-3 項]

🙏 感謝本週辛苦的：[提及有特殊貢獻的人]

語氣輕鬆但專業，適合 Teams 頻道發布。
```

> 💡 **教學 Tip**：這個排程提示每週五會自動產出草稿，你只需要花 5 分鐘檢查修改後發布。原本 30-45 分鐘的工作變成 5 分鐘。而且因為 Copilot 搜尋的是完整的資料源，不會像人一樣遺漏。

---

## 場景一學員練習建議

### 練習 A：會議摘要練習（10 分鐘）
1. 兩人一組，開啟一場 5 分鐘的 Teams 會議
2. 討論任意工作議題
3. 會後使用 Copilot in Teams 產生會議摘要
4. 嘗試至少 3 種不同的提示來提取不同角度的資訊

### 練習 B：週報產出練習（15 分鐘）
1. 使用 **英業達_NPI_ProServerX200_專案追蹤.xlsx** 作為資料來源
2. 在 Word 中用 Copilot 產生一份 NPI 週報
3. 使用 Rewrite 調整至少一段的語氣和長度
4. 在開頭加入 Executive Summary

### 練習 C：接手專案練習（10 分鐘）
1. 想像你剛接手一個同事離職留下的專案
2. 使用 Copilot Chat (Work) 搜尋 Teams 頻道和郵件
3. 產出一份「專案交接摘要」
4. 識別需要立即處理的待辦事項

### 練習 D：排程提示練習（5 分鐘）
1. 在 Copilot Chat 中設定一個排程提示
2. 例如：每天早上 9 點摘要昨天的重要 Teams 訊息

---

## 場景一 Tips 總整理

| # | Tip | 適用 Demo |
|---|-----|----------|
| 1 | Facilitator 是公開的 AI 與會者，Copilot in Teams 是私人助理 | Demo 1, 2 |
| 2 | 用 @Facilitator 在會議聊天中互動，全員可見 | Demo 1 |
| 3 | 會後 AI Summary 比即時筆記更完整（基於完整轉錄稿） | Demo 2 |
| 4 | Copilot Pages 是團隊協作的中繼站 | Demo 2 |
| 5 | Work 模式的 Copilot Chat 可跨 App 搜尋組織資料 | Demo 3 |
| 6 | 排程提示自動化重複性彙整工作 | Demo 3 |
| 7 | Channel Agent 專注單一頻道，追蹤 deadline 和未回覆問題 | Demo 5 |
| 8 | Word 用 `/` 引用檔案作為 Copilot 上下文 | Demo 13 |
| 9 | Rewrite 功能快速調整語氣、長度、專業度 | Demo 13 |
| 10 | 跨 App 資料流才是 Copilot 最大的價值 | 全部 |
