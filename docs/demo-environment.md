# AB-730 demo 環境與唯讀盤點界線

## 交付狀態

本文件記錄 **2026-10-14 文件實作交接時的本機盤點快照**，不是租戶驗證報告。教材與產業教學設計見 [teaching-guide.md](teaching-guide.md)。

| 項目 | 狀態 | 可以／不可以推論 |
| --- | --- | --- |
| 本機 `seed-data\` | 既有未追蹤的 MS-4018／MS-4022 相關副本 | 可讀結構，不能當成 AB-730 已發布、已測試或 fresh clone 會有的交付物 |
| 有效實際設定 | 交接盤點找到的 `config.json` 數量為 **0** | 不具備已確認的 tenant／app／操作帳號目標；不是認證成功 |
| 設定範本 | 四份 `config.json.example`；tenant／secret 欄位為 placeholder | 只是欄位範例，template domain、角色與 resource 宣告不是實際登入或資源存在證據 |
| 真實租戶 GET | **blocked：缺有效、已授權的實際設定** | 未取得 token／claims 核對，未完成角色、owner、Teams、OneDrive 或 SharePoint 盤點 |
| Copilot 介面／功能 | 未驗證 | 官方功能存在不等於操作帳號看得到或本租戶可用 |
| 汽車產業資料 | 尚未設計／產生／灌入 | 舊 Ford／Inventec 情境不可改名後冒充本課成果 |
| 正式 runner／雲端寫入 | 本次不執行 | 不寄信、不建立 Teams 訊息、檔案、會議、agent、Page、Notebook 或 Cowork 任務 |

此快照來自本次 session plan 的本機盤點與 `git --no-pager status --short`；後續主 agent 若取得新證據，應更新狀態並保留查證日期，不能靜默把本機檢查改寫成線上完成。公共文件不列帳號、domain、tenant／app／resource IDs、secret、token、原始訊息或 receipt。

## 本機可確認的程式與素材

下列路徑是**本機未追蹤副本的證據定位**，不是公開 repo 保證包含的檔案；fresh clone 缺少這些檔案屬預期，不代表依賴已安裝。

| 證據 | 已確認內容 | 對 AB-730 的影響 |
| --- | --- | --- |
| `seed-data\README.md:1-39` | 描述其他課程 seeding 與其使用方式 | 原課程成功紀錄不能替代本課驗證 |
| `seed-data\generator\pack.schema.json:22` | `^[A-Z]{2}-[0-9]{4}$`，只接受四位數課號 | `AB-730` 不符合；需後續 issue 修正並回歸 MS-4018／MS-4022 |
| `seed-data\generator\build_scenario.py:164` | workbook creator 固定 `"MS-4018 demo"` | 不能宣稱輸出已有 AB-730 branding |
| `seed-data\generator\build_scenario.py:200-202` | 說明頁名稱與提示文字固定簡體中文 | 本課繁中資料生成仍須獨立處理與驗收 |
| `seed-data\engine\Seed-GraphRead.ps1:23-59` | 共用集合 reader 跟隨 `@odata.nextLink`，預設 `MaxPages = 100`，對上限與自循環拋錯 | 可作後續限定 GET reader 的入口；不能只用第一頁 |
| `seed-data\engine\Invoke-SeedPreflight.ps1:109-123` | display name／alias 兩個 group 查詢只取 `$r.value`，未跟隨 nextLink | preflight 不是完整 group inventory；分頁是後續 issue 的前置修正／測試範圍 |
| `seed-data\scenarios\ms4018-ford-auto-workshop-b-20260929\run.ps1:1-18` | 有 `-PreflightOnly`；先連線與 preflight，未帶 switch 時繼續檔案／email／Teams 寫入 | 不可執行無參數 runner；有 switch 也需先讀實際呼叫路徑、授權與設定 |

本次不修改 generator、schema、preflight、runner、tests 或 `.gitignore`。把這些缺口留給後續資料 issue，比直接複製既有 scenario 安全，也避免錯課號／錯語系與共享租戶寫入。

## 能力與資料缺口矩陣

「未驗證」表示尚無實際證據；「blocked」表示已知缺前置；403／429 不可轉寫為資源不存在。此文件不認證任何 template surface 或帳號。

| Surface／功能 | 後續需確認的能力與資料 | 目前結果 | 課堂備案 |
| --- | --- | --- | --- |
| 操作帳號 | 有效設定的 `adminUpn`、account enabled、與 token tenant／client claims 一致 | blocked；未確認 | 投影片；不登入其他角色 |
| Microsoft 365／Copilot license | 既有 assignment、產品 label、app 支援、政策 | 未驗證 | 講解官方條件，不宣稱 live 可用 |
| Copilot Chat | 工作入口、EDP、web／app context、Work IQ／來源權限 | 未驗證 | Prompt 評析與來源辨識 |
| Outlook | 僅 scenario 宣告且獲授權的內容、操作帳號的存取 | blocked；未讀信件 | 通知文字設計，不開收件者或寄送 |
| Teams | 既有 Team／channel、owner／member、訊息完整性與角色可讀 | blocked；未查集合／訊息 | 口頭協作流程；不發訊息 |
| OneDrive | 已核准 folder／檔案、manifest 對照、操作帳號能讀 | blocked；未查檔案 | 不上傳來源、不生成替代 Office 檔案 |
| SharePoint | 已核准 site、library、list、items、操作帳號可見性 | blocked；未查站台／文件／清單 | Agent 知識來源設計 |
| Excel／Analyst | BOM／交期／批次檢驗合成資料、欄位、數字一致性、功能權限 | 未驗證；資料尚未產生 | 分母／join／單位檢查練習 |
| Word／PowerPoint／Researcher | 已核准摘要、引用來源、app／agent 可用性 | 未驗證 | 報告結構與研究問題 |
| Facilitator／Teams recap | 會議類型、license、policy、同意／逐字稿／語言 | 未驗證 | 已核准文字紀錄或只談議程 |
| Pages／Notebooks | storage／service plan、enablement、source／分享控制 | 未驗證 | 畫布與參考來源設計；不建立或分享 |
| Cowork | assigned license、enablement、usage-based billing、model／subprocessor policy | 未驗證；本次禁止提交任務 | 官方影片與紙上任務拆解 |
| Calendar／Meeting Chats 寫入 | 目前沒有新增這些 surfaces 的教學／寫入核准 | 不在本次範圍 | 不為課程完整性擴張範圍 |

官方前置條件與實際測試保持分欄。Researcher／Analyst 的 add-on、Facilitator 的啟動條件與 Cowork 計費不從 template 或 seeder app-role 推論。

## 取得有效設定後的唯讀查證規則

以下是**待執行檢查流程**，不是已完成紀錄，也不是現在取得 write 授權：

1. 由 operator 在本機提供既有、已授權的實際設定位置；先查 tracked／ignore 狀態，不能把秘密貼進對話、公共 issue 或 repo。沒有有效設定就維持 blocked。
2. 只使用既有 app 與 permissions，不建立帳號、不新增 consent、不補 owner。Token 可向既有 OAuth endpoint 取得，但 Graph 資料盤點**只用 GET**；確認 claims 與設定匹配，只保存脫敏結論。
3. 按既有 scenario 的明確 allowlist 查角色與資源，先查操作帳號可見性，再核對 Team／channel／檔案／list 等 surfaces。不要為找資料掃全租戶或無關信箱。
4. 所有集合跟隨 `@odata.nextLink`，設定頁數上限；到達上限、解析失敗、重名、權限不足或部分內容要保留限制。未分頁的 preflight group 查詢不足以證明完整 inventory。
5. 只有 runner 實際路徑仍為唯讀、設定與授權齊備，才可評估 `-PreflightOnly`；不為通過 write-role 檢查新增權限。不能執行無參數 runner，不重跑原 Ford／Inventec 灌入。
6. 403 是存取失敗、429 是節流；都不是「沒有資料」。可在核准範圍尊重 Retry-After 重試，不能修補、刪除、改名、收養或重跑碰撞資源。
7. 原始帳號／IDs／claims／資源回應留在本機 session；公共文件只更新「可用、缺資料、缺權限、未驗證、blocked」與查證日期。必要的實際介面確認由主 agent 完成，不以 app GET 取代操作帳號的 UI 存取證明。

## 後續 demo data issue 的必要前置

本文件列待辦條件，不自行建立 issue，不代表資料產生／灌入已獲授權。

- **相容性先做：** `AB-730` 三位數課號 schema、workbook creator／branding、繁中說明頁、group pagination；保留舊課程回歸與 fail-closed 測試。
- **資料先設計：** 虛構整車／零組件、供應商交期、BOM 成本、品質異常與改善會議；明確指標、單位、唯一鍵、來源版本與數字驗算，跨 email／Teams／文件不互相矛盾。本次不生成資料。
- **surface 不膨脹：** 只選教學需要的 Excel／文件、Outlook threads、Teams 內容、SharePoint knowledge；不順便啟用 Calendar／Meeting Chats 或新增使用者／修改職稱。
- **保護本機：** 精確忽略 secrets／receipt／runtime／下載檔，保留使用者既有修改；本次不修改 `.gitignore`，不將整個未追蹤 `seed-data\` 納入提交。
- **離線測試再核准：** 修正 engine 前做 RED／GREEN；驗證分頁、碰撞、半成品、不可覆寫、數字一致性與資料邊界。不得把現成 script 或 template 當成測試通過。
- **線上寫入另開範圍：** 精確 tenant／app／操作帳號／目標／資料量、獨立高風險計畫與實作審查、明確 operator 授權；只向全新核准目標做加法建立。碰撞、半成品、403 即停止。
- **結果另驗收：** 原始 receipts 留本機，線上狀態唯讀核對，主 agent 再用實際介面確認可演示；沒有這些證據不能關閉後續 issue 或宣稱備課全部完成。

## 官方條件參考

- Chat grounding 與 permissions：https://learn.microsoft.com/en-us/copilot/overview
- License options：https://learn.microsoft.com/en-us/microsoft-365/copilot/microsoft-365-copilot-licensing
- EDP 與 web-query 例外：https://learn.microsoft.com/en-us/microsoft-365/copilot/enterprise-data-protection
- Pages storage／授權條件：https://support.microsoft.com/en-us/microsoft-365-copilot/how-microsoft-365-copilot-pages-works
- Notebooks service plan／限定來源：https://support.microsoft.com/en-us/microsoft-365-copilot/how-microsoft-365-copilot-notebooks-works
- Facilitator license／meeting eligibility：https://support.microsoft.com/en-us/teams/copilot/facilitator-in-microsoft-teams-meetings
- Cowork license／enablement／billing：https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/get-started

以上僅證明官方條件，不證明本租戶符合。未完成項目：有效設定、token／角色／owner／資料盤點、license／功能 UI、產業資料與正式 demo；狀態仍未驗證或 blocked。
