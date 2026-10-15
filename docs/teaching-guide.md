# AB-730 汽車產業講師備課指南

## 定位與使用界線

本指南依 **Transform business workflows with generative AI** 的八個 module 編排，面向已熟悉 Microsoft 365 的商務使用者，不教授 Azure 部署、模型開發或 Graph 工程。產業改編是**講師自訂活動**，不是 Microsoft 官方 lab，也不是已完成的汽車產業資料集。

本階段只完成教學設計與文件；**真實租戶盤點 blocked、功能實際可用性未驗證**。環境證據與後續前置條件見 [demo-environment.md](demo-environment.md)。不得將「官方產品支援」說成「本租戶可演示」。

- 使用虛構企業「遠景汽車製造」，串連整車組裝、零組件交期、供應鏈風險與品質改善；不借用真實客戶的業績、召回、BOM 或內部通訊。
- 本文件的 prompt 是設計稿；來源占位符須在後續資料核准後換成實際可開啟的來源。缺來源時只討論方法，不要求 AI 補出事實或生成替代資料。
- 真實租戶演示只能使用有效設定中既有、已授權的 `adminUpn` 操作帳號；不登入其他商務角色、不建立帳號、不改 profile、不補 owner 或權限。
- 本次不建立或上傳 Excel／Word／簡報資料、不寄信、不發 Teams 訊息、不建立會議、Page、Notebook、agent、排程或 Cowork 任務。Live 操作須另有精確授權；本指南優先採口頭推演、投影片與已驗證官方影片。
- 引用 PPT 檔名與**該檔案的實際 slide number**，不沿用 notes 中各 learning path 重起算的 module 編號。Speaker notes 用來理解教學意圖，不逐字轉載。

課程研究入口沿用 README 的官方教材頁：https://learn.microsoft.com/en-us/training/courses/ab-730t00 。學員參考資料在 [README](../README.md)，不要把租戶盤點、登入資訊或講師內部路徑放回學員頁。

## 一日議程建議

以下是講師安排，不是官方保證時數；共 390 分鐘教學，另安排午餐及兩次休息。若沒有已授權環境，將 app 操作時段改為 prompt 評析與人工驗證練習，不假裝完成 live lab。

| 時段 | 內容 | 分鐘 |
| --- | --- | ---: |
| 開場 | 帳號／來源／授權界線與案例說明 | 15 |
| M01 | 負責任使用與結果核對 | 35 |
| M02 | Chat、grounding 與功能差異 | 35 |
| M03 | 提示詞與反覆修正 | 45 |
| M04 | 摘要、Pages 與 agents | 50 |
| M05 | 內容工作流程與 Researcher | 55 |
| M06 | BOM 成本、品質指標與 Analyst | 60 |
| M07 | 改善會議與協作 | 50 |
| M08 | Cowork 委派、核准與用量 | 35 |
| 收束 | 課後應用與考試版本提醒 | 10 |

## 官方練習與自訂活動對照

只查閱練習網頁，不代表已登入、下載資料或完成練習。官方 lab 在 lab provider 提供的環境依其指示執行，**不得照搬到共享 demo 租戶**。

| Module | 官方 exercise 查證結果 | 本指南的汽車產業改編 |
| --- | --- | --- |
| M01 | `_01.pptx` slides 18–19 的知識檢核是可選路線，不是必做 lab | 敏感來源與品質異常摘要的人工檢核 |
| M02 | `_02.pptx` slide 20 為可選知識檢核 | 區分公開交期資訊、提供內容與組織搜尋 |
| M03 | `_03.pptx` slide 2 notes 明示無 exercise unit；自加練習應標示選用 | 將模糊缺料請求改成可驗證的 prompt |
| M04 | [Plan a client summit with Copilot Chat](https://microsoftlearning.github.io/AB-730-AI-business-professional/Instructions/Labs/01-plan-a-client-summit-with-copilot-chat.html)；`_04.pptx` slide 18 | 供應商風險簡報與品質知識 agent 設計 |
| M05 | [Draft and refine business content with Microsoft Copilot](https://microsoftlearning.github.io/AB-730-AI-business-professional/Instructions/Labs/02-draft-and-refine-business-content.html)；`_05.pptx` slide 18 | 品質改善主管摘要、簡報大綱與未寄送信件文字 |
| M06 | [Analyze and visualize data with Microsoft Copilot](https://microsoftlearning.github.io/AB-730-AI-business-professional/Instructions/Labs/03-analyze-and-visualize-data.html)；`_06.pptx` slide 13 | BOM 成本、交期與不良率的分析設計 |
| M07 | [Manage meetings and collaboration with Microsoft Copilot](https://microsoftlearning.github.io/AB-730-AI-business-professional/Instructions/Labs/04-manage-meetings-and-collaboration.html)；`_07.pptx` slide 18 | 採購／生管／品質改善會議與待確認行動 |
| M08 | [Get Started with Copilot Cowork](https://microsoftlearning.github.io/AB-6008_delegate_work_with_AI/Instructions/Self-Paced/01-cowork-orientation.html)；`_08.pptx` slide 29 的 relationship 指向 AB-6008 | 多步驟改善會前準備的紙上委派；不執行 Cowork |

**教材校正提醒：** `_05.pptx` slide 18 的顯示標題是內容撰寫，但 notes／relationship 指到會議 lab；以上採已核對標題與內容的 **02-draft-and-refine-business-content** 頁面。這是指南中的勘誤，不修改原簡報。M08 的來源跨課程也保留標示，不宣稱存在另一個已驗證的 AB-730 Cowork lab。

## 授權事實不等於租戶可用性

查證基準為 2026-10-14 的官方產品文件；名稱正從 Microsoft 365 Copilot／Copilot Chat 轉為 Microsoft Copilot／Copilot Chat。保留官方 module 原名，並說明 UI 可能不同。

| 能力 | 官方文件所述條件／邊界 | 本環境狀態 |
| --- | --- | --- |
| Copilot Chat | 符合資格的 Microsoft 365 工作／學校帳號可使用 Chat；包含 web、使用者提供內容及受支援 app context。沒有 add-on 不代表完全不能讀工作內容，也不代表具備完整 Work IQ | 未驗證資格、入口、保護標示、app context |
| Microsoft 365 Copilot | Add-on 加入更廣的組織 grounding、優先存取與 app／agent 能力；既有資料權限仍然有效 | 未驗證操作帳號的 license assignment 與實際來源存取 |
| Researcher | 商務／企業帳號需 Microsoft 365 Copilot add-on；Support 另列個人 Premium／Pro，但本指南不使用個人帳號。多來源研究仍須人工開啟引用 | 未驗證 agent 顯示、政策與研究用量 |
| Analyst | 官方商務 Chat 比較頁將 Analyst／Researcher 列為有 Copilot license 的進階 agents；用於深入調查，不是僅依檔案數量選擇 | 未驗證可用性、資料讀取或計算結果 |
| Facilitator | 加入／啟動須 Microsoft Copilot license；組織內參與者可看共用更新，外部參與者例外。需管理員／會議政策允許；不是每種會議都支援 | 未驗證 license、policy、會議類型、語言與逐字稿 |
| Pages | 工作／學校帳號有 SharePoint 或 OneDrive storage 者可用，包含沒有 Copilot add-on 的使用者；共編不等於取得原 Chat 或來源權限 | 未驗證 storage、建立／分享控制與收件者存取 |
| Notebooks | 目前 Support 列 Microsoft Copilot 或 Copilot Chat 使用者；建立須 SharePoint 或 OneDrive service plan。限定參考來源，不是全租戶搜尋 | 未驗證入口、source limits、分享與來源權限 |
| Cowork | 目前 Learn 說工作／學校帳號 GA、個人帳號 Preview；開始使用須 assigned Microsoft 365 Copilot license、Cowork enabled、usage-based／Cowork billing enabled。App skill 另須 Frontier | 未驗證 enablement、計費、模型／subprocessor 政策；不發送測試任務 |

官方依據：

- Chat 與進階 agents：https://support.microsoft.com/en-us/microsoft-365-copilot/how-copilot-chat-works-with-and-without-a-microsoft-365-copilot-license
- 來源／入口差異：https://learn.microsoft.com/en-us/copilot/overview
- License 識別：https://support.microsoft.com/en-us/microsoft-365-copilot/what-copilot-license-do-i-have
- Add-on 選項：https://learn.microsoft.com/en-us/microsoft-365/copilot/microsoft-365-copilot-licensing
- Researcher：https://support.microsoft.com/en-us/microsoft-365-copilot/get-started-with-researcher-in-microsoft-365-copilot
- Facilitator：https://support.microsoft.com/en-us/teams/copilot/facilitator-in-microsoft-teams-meetings
- Pages：https://support.microsoft.com/en-us/microsoft-365-copilot/how-microsoft-365-copilot-pages-works
- Notebooks：https://support.microsoft.com/en-us/microsoft-365-copilot/how-microsoft-365-copilot-notebooks-works
- Cowork 發布狀態：https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/
- Cowork prerequisites：https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/get-started

不要把 PPT 的初期 Frontier 說法、產品舊名或「含某功能」當成永久授權承諾；實際 assignment、管理員政策、版本與地區須由 operator 再核對。不要為 demo 新增付費項目或啟用外部模型。

## 貫穿全課的教學原則

1. **先問成果，再選工具。** Chat 適合逐輪討論；Excel 適合 workbook 內分析；Researcher 調查證據；Analyst 深入數據；Pages 共編；Notebooks 限定來源；Cowork 委派多步驟成果。
2. **來源與期待分開。** 指定提供的資料不代表 AI 能自動搜尋所有信件或工廠系統；Work IQ／connector 必須已配置且帳號可讀。
3. **讀取、生成、核准、發佈是不同動作。** 提示詞中的「不要寄送」不是技術權限限制；本階段靠不啟動寫入流程維持界線。
4. **AI 不能核准停線、放行或供應商更換。** 生成內容須由生管、品質與採購專業人員確認；品質因果分析不能把相關性當根因。
5. **資料保護有條件。** EDP 對 prompts、responses、Graph 資料有不訓練 foundation models 的承諾，但 Bing web queries 與部分模型的資料處理條款／區域承諾有例外。不可說「所有 AI 資料永不離開本地區」。

依據：https://learn.microsoft.com/en-us/microsoft-365/copilot/enterprise-data-protection 。

## M01 - Prepare to use generative AI at work

**學習目標：** 說明生成與分類的差異、負責任 AI、常見風險、內容核對及依任務選 mode。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_01.pptx` slides 6–7 用生成與分類建立非技術心智模型；slides 10–14 說明六原則、敏感資料與核對順序；slides 16–18 比較四種工作方式。不要把流暢答案解釋成人類理解，也不承諾固定效率提升。

**演示條件：** 只需投影片與講師口頭情境。實際 Chat 若另行核准，先核對工作帳號、EDP 狀態與允許提供的合成文字；不用真實客戶異常單。

**汽車產業 prompt（自訂、討論稿）：**

> 我是遠景汽車製造的品質協調人員。只根據 /06-Quality-Brief.docx，改寫一段給跨部門主管的簡短說明。把已確認事實、可能原因、待查證事項分開；不要推定召回、停線或零件放行。先指出哪些資訊不應放進提示詞；不輸出個資、帳密或真實客戶資料。

**預期輸出：** 一份帶有證據／推測區分的文字草稿及資料最小化建議，不是正式品質處置命令。

**人工核對：** 開啟來源確認零件／批次敘述；逐項看是否新增原因或承諾；指出錯誤、不公平標籤、敏感資料及 overreliance 風險。不能查證的專業主張交由品質工程師。

**失敗備案：** 無來源／無 Chat 就請學員評析 prompt 的資料邊界；不編造一份「已驗證」品質摘要。知識檢核可選原 assessment 或講師檢查過的生成題，兩者不必都做。

**重要資料：** https://support.microsoft.com/en-us/microsoft-365-copilot/diagnose-copilot-output 。

## M02 - Explore how Microsoft Copilot Chat works

**學習目標：** 找到入口、辨識保護狀態與 grounding，開始／追問對話並解釋 app／license 差異。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_02.pptx` slides 7、10、12–16 比較提供內容、app context、Work IQ 與 response mode；slide 19 notes 要對區域與 web-query 資料處理作限定。不要將 Auto／Think deeper 固定對應特定模型版本。

**演示條件：** 有效工作帳號與可見 Chat；完整組織搜尋另須證明 license 與來源權限。本階段未驗證，不能假裝已查到 Teams／Outlook。

**汽車產業 prompt（自訂、討論稿）：**

> 我在準備電動車零組件供應鏈風險說明。只用公開、可開啟的來源，整理最近三個月與交期風險有關的資訊；列來源發布日期、適用地區與限制。不要聲稱讀取本公司的採購信件、Teams 或 ERP。若缺證據就標示未知。

**預期輸出：** 附可開啟來源與日期的公開資訊摘要；不是遠景汽車的實際缺料清單。

**人工核對：** 檢查發布日期不是僅擷取日期；確認引用支持敘述；對照目前 grounding／EDP indicator。追問「哪些結論不能套用到本公司？」測試是否保留界線。

**失敗備案：** 無 web grounding 就討論提示詞的來源範圍，不改接個人 Copilot。介面不同時以 slide 說明；不清除 chat history、不變更租戶政策。

**重要資料：** https://learn.microsoft.com/en-us/copilot/privacy-and-protections 。

## M03 - Write effective prompts to achieve optimal results

**學習目標：** 使用 Goal／Context／Source／Expectations，避免矛盾指示，反覆修正並區分 Memory、custom instructions 與 Prompt Lab。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_03.pptx` slides 5–8 提示詞要素與驗證；slides 15–17 個人化與 temporary chat；slides 19–21 prompt 重用／分享／排程。短 prompt 不強制四要素齊備；temporary chat 不等於無保存或免於治理。

**演示條件：** 可在投影片上比較兩個 prompt；不儲存 Memory、不分享 prompt、不建立排程。未授權來源不能靠 `/`、`#` 或 `@` 取得權限，快捷語法依介面而異。

**汽車產業 prompt（自訂、討論稿）：**

> 情境：我是生管協調人，要向採購與品質主管報告缺料風險。目標：只根據 /02-Supplier-Delivery.xlsx，列本週需人工確認的前三項風險。來源：只限我指定的文件，不搜尋其他信箱。期待：繁體中文表格，欄位為來源位置、風險、影響假設、需確認事項；不補造日期、數量或負責人。

**預期輸出：** 可逐列回查的風險草稿。追問「改成給採購主管的三點摘要，保留所有未知」後，內容範圍與限定應一致。

**人工核對：** 四要素各圈一處；看風險排序的依據；對照第一輪與追問是否遺失未知／假設；確認分享 prompt 文字不等於授權其引用文件。

**失敗備案：** Prompt Lab 不可用就評析 prompt 文字，不宣稱已收藏或已分享。notes 說 revised prompt 需另存新版；當前 Support 的可編輯選項與 UI 若不同，依實際已驗證介面說明，不把 notes 當永久規則。

**重要資料：** https://support.microsoft.com/en-us/microsoft-365-copilot/save-your-favorite-prompts-to-use-again ；https://support.microsoft.com/en-us/microsoft-365-copilot/share-your-best-prompts-with-others 。

## M04 - Work smarter with Microsoft Copilot Chat

**學習目標：** 從內容發想與摘要到 Pages 共編，辨識 agent 的專用指示、知識、權限與用量。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_04.pptx` slides 6–7 分開生成圖片、資料圖表與檔案分析；slides 11–12 說明 Page 與 Chat 的分享邊界；slides 15–17 提醒 agent 的 cost／來源與 Viewer／Editor 差異。

**演示條件：** 後續若核准，使用可讀的合成改善摘要；Pages 要有 storage／enablement。Agent Builder 只設計設定，不建立、連線、分享或啟用計費。

**汽車產業 prompt（自訂、討論稿）：**

> 根據 /06-Quality-Brief.docx，提出三個跨部門改善討論方向。每項列支持證據、反例與待確認事項；只回覆文字，不建立 Page、檔案、agent 或任務。另描述一個「品質文件問答 agent」應有的指示、允許的知識範圍與拒答條件，不實際建立。

**預期輸出：** 三個可討論方向及 agent 設計草稿；Pages 的目標是可持續編輯畫布，不是自動對外公開。Agent 回答來源不足時應拒絕補造放行規範。

**人工核對：** 證據與推測分離；AI 圖片不被當工廠現況；若未來分享 Page，另核對誰能看 Page／來源、誰能 edit，不能認定 `@mention` 自動授權。

**失敗備案：** 無 Pages／agents 時只比較 Chat 與專用助理的工作方式；不以新增 Copilot Studio／付費 meter 解決。沒有可用檔案，不上傳原授權教材作為替代知識。

**重要資料：** https://support.microsoft.com/en-us/microsoft-365-copilot/share-a-microsoft-365-copilot-page ；https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/agent-builder-add-knowledge 。

## M05 - Draft and refine business content using Microsoft Copilot

**學習目標：** 將相同已核對事實改成主管摘要、簡報與信件，理解 Researcher 的多步驟研究與來源核對。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_05.pptx` slides 7–9 選擇 Chat only／Allow editing；slides 10–14 跨 app 維持事實、受眾與商務目標一致；slides 16–17 是研究→深化→重組，不是研究報告自動成為決策。

**演示條件：** 本次只談內容輪廓。未來 app demo 需實際 assignment、app version 與已核准來源；Researcher 另核對可見性與政策。**不要開新郵件填收件者、不要按 Send，也不產生 Office 檔案。**

**汽車產業 prompt（自訂、討論稿）：**

> 只依［已核准的品質改善會議摘要與供應商回覆］，在聊天中提出給廠長的一頁報告結構、五段簡報大綱與未寄送的供應商溝通文字。三種形式使用同一組事實與待確認事項，不新增承諾、交期或費用；不建立 Word／PowerPoint／Outlook 草稿、不寄送或分享。

**Researcher 選用 prompt：**

> 針對「汽車零件供應商變更前應查證哪些風險」設計研究範圍、證據類型與報告架構。實際研究時只用核准且可讀的來源，標明地區、日期與資料缺口；不把公開一般資訊說成本公司的品質核准規範。

**預期輸出：** 三種格式的文字設計稿與研究檢查框架；不是已完成主管報告、已寄出郵件或實際供應商替換決定。

**人工核對：** 各版本的料號、日期、成本與風險一致；coaching 不會把「待確認」改成承諾；研究引用與地區法規適用性由專業人員確認。跨 app 改變形式，不改變事實。

**失敗備案：** 未有 Researcher 時用 Chat 討論研究問題，不宣稱取得同等深度；沒有 Word／PowerPoint 時保留文字大綱，不新建資料集。提示詞內容符合「不寄送」仍不能證明工具權限受限。

**重要資料：** https://support.microsoft.com/en-us/word/copilot/draft-and-add-content-with-copilot-in-word ；https://support.microsoft.com/en-us/outlook/copilot-pages/get-email-coaching-with-copilot-in-outlook 。

## M06 - Analyze and visualize data using Microsoft Copilot

**學習目標：** 分析趨勢、異常、文字主題、公式與圖表；依任務／分析深度／輸出選 Excel、Analyst 或 Chat。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_06.pptx` slides 5–8 檢查模式、來源列與零值；slides 10–12 解釋 Analyst 即使只有一個 workbook 仍可能合適，換成信件摘要不等於新的分析證據。

**演示條件：** 汽車 BOM／交期／批次檢驗資料**尚未產生**；本次不生成 xlsx／csv，不把官方銷售 lab 資料冒充汽車情境。未來資料需有欄位定義、幣別、計量單位、期間、唯一鍵與可手算的預期結果。

**汽車產業 prompt（自訂、討論稿）：**

> 使用［已核准的 BOM 成本、供應商交期與批次檢驗資料］，先檢查料號／批次鍵、單位、期間、缺值與重複列。說明如何比較 BOM 成本變動、逾期風險與批次不良率，指出適合的圖表與計算式；無資料時只列必要欄位與驗證步驟。不要建立或修改 workbook，不生成數字或判定品質根因。

**Analyst 選用 prompt：**

> 設計「成本增加是否與品質異常同時出現」的調查計畫，列需要的資料、join 鍵、分母、反例與不確定性。實際分析前先核對欄位與樣本數；相關性不能作為供應商歸責證據。

**預期輸出：** 可執行的分析計畫、指標定義與驗算清單，不是憑空生成的洞察或已完成圖表。

**人工核對：** 不良率為不良件數／檢驗件數，不能直接平均各批次百分比；成本先核對幣別與數量基礎；逾期比較承諾日與實到日，缺實到日不可當準時；檢查重複 join、零分母與軸標示。未來至少手算一組總計、一個比率與一個異常列。

**失敗備案：** 沒有資料／Excel／Analyst 就由學員選工具並解釋驗證策略，不把「沒有數據」當成生成數據的授權。未來 live demo 用 practice copy，先查模式，不能拿原檔測試直接編輯。

**重要資料：** https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel ；https://support.microsoft.com/en-us/microsoft-365-copilot/get-started-with-analyst-in-microsoft-365-copilot 。

## M07 - Manage meetings and collaboration using Microsoft Copilot

**學習目標：** 串連會前準備、會中私人／共享協助、recap、後續文字與 Pages／Notebooks。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_07.pptx` slides 6–11 區分準備與 calendar action，從證據回查負責人／期限；slides 13–17 比較共編畫布、參考來源、分享與音訊來源限制。

**產品文件優先的校正：** notes 的逐字稿要求較強；目前 Support 說會中 Copilot 可依政策無 transcription 使用，但會後查詢／history 需要 live transcription。Facilitator 只能加在支援的排定會議，不是 channel／instant meeting／Teams call 都可用。Notebook 的 notes「前 100 references」已與目前 Support「最多 300 files grounding」不同；各頁採 references／files 不同單位，課前核對當前限制，不混為一談。Cowork／Notebook 新功能也不要只照舊 notes 下定論。

**演示條件：** 已授權會議、參與者同意與 transcription policy 才能做 live；本次不建立／錄製會議、不啟動 Facilitator、不分享 Page／Notebook，先用會議設計稿。

**汽車產業 prompt（自訂、討論稿）：**

> 只根據［已核准的採購、生管與品質改善會議紀錄］，整理供應商交期、缺料與品質阻擋事項。用表格分開已決議、建議與待確認；只在來源明確時列 owner／due date，否則標「待會議確認」。在聊天中提供下一次會議議程與後續通知文字，不安排會議、不寄信、不發訊息或建立任務。

**Notebook 設計 prompt：**

> 規劃一個限定來源的品質改善 Notebook，列必要參考文件、版本／存取核對與需問的問題；另列 Page 上可共同編輯的目標、里程碑與風險區塊。不實際建立或分享。

**預期輸出：** 有來源位置的決議／待辦草稿、會議議程與來源設計；「建議 owner」不能變成已承諾 owner。

**人工核對：** 回到逐字稿／紀錄確認是提議還是同意；檢查來源版本與誰能讀；私人 Copilot 與共享 Facilitator 的可見性不同。分享 Notebook 對其 references 的影響須逐項核對分享對話框與既有權限，不因只送連結就認定安全。

**失敗備案：** 缺 transcription、Facilitator 或 license 就使用已核准的文字紀錄流程；不偷偷錄音、不掃描其他信箱。沒有紀錄只演練議程設計，不生成假的會議決議。

**重要資料：** https://support.microsoft.com/en-us/teams/copilot/catch-up-on-meetings-with-microsoft-365-copilot-in-teams ；https://support.microsoft.com/en-us/microsoft-365-copilot/share-a-microsoft-365-copilot-notebook 。

## M08 - Get more done with Copilot Cowork

**學習目標：** 依工作形狀選 Chat／agents／Cowork，寫成果導向 request，理解用量、skills、automation、核准與驗收。

**PPT／講解重點：** `AB-730T00-ENU-PowerPoint_08.pptx` slides 5–8 選工具；slides 10–15 看用量與核准；slides 18–19 寫成果、情境與交付物；slides 23–28 權限、限制與何時／如何重複。

**演示條件：** Cowork actual enablement／assigned license／billing／model policy 均未驗證。**本次不提交任何 Cowork prompt**，不按 Approve／Always allow／Approve All、不建立 custom skill、不排程，避免任務即使未寄送仍產生檔案、存取與用量副作用。

**汽車產業委派 prompt（紙上設計，不送入 Cowork）：**

> 目標：準備本週跨部門缺料與品質改善檢討。範圍：只使用［明確核准的本週交期摘要、BOM 成本說明與品質會議紀錄］。交付設計：來源清單、風險摘要、主管 briefing 結構與待確認事項；不要搜尋其他信箱、寄送、分享、建立檔案、修改日曆或任務。先提出計畫與可能產生副作用的步驟，由人類評估是否可另行授權。

**預期輸出：** 在課堂紙上產生任務拆解與核准決策表，不是 Cowork 執行結果。若未來要真的執行，即使文字限定不寫入，也須先查工具與會話既有授權，不把 prompt 當安全閘門。

**人工核對：** 每一步標為 read／draft／write／share／schedule；原始來源可追溯、用量可看、未知事項未消失；逐項檢查收件者、時間、輸出位置與 permissions。Automation 決定何時，custom skill 決定如何；本次兩者都不建立。

**失敗備案：** 使用 README 已驗證 Microsoft 官方影片或投影片講解。PPT 的 `/cost`／Copilot Credits 是教材操作線索，未在本租戶驗證，不保證固定費率或 UI；以當前 billing prerequisite 與 operator 可見資訊為準。notes「無本機檔案存取」不應概括為不能提供本機檔案：目前 Learn 已提供 upload from device，但本次不執行上傳。

**重要資料：** https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/get-started ；https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/use-cowork 。

## 常見問題與收束

- **有 license 就一定能查到文件？** 不一定。還有帳號權限、來源可見性、索引、版本、政策與功能 rollout；打不開來源就停止，不以擴權解決。
- **引用可以當真？** 引用只是回查入口。核對是否存在、日期、適用範圍與原文是否支持結論；不能只看摘要說得有信心。
- **Pages／Notebooks 一定要 add-on？** 不可一概而論。依上表查 storage／service plan 與當前 Support，再核對租戶實際能力。
- **Intelligent recap 等於完整 Copilot license？** 官方 recap 頁列 Teams Premium 或 Microsoft Copilot 可提供 intelligent recap；不代表 Premium 單獨解鎖所有 Copilot／agents。依據：https://support.microsoft.com/en-us/teams/meetings/recap-in-microsoft-teams 。
- **可以先跑現有 Ford／Inventec seeder 嗎？** 不可以。本階段沒有 write 授權；舊課程成功記錄不等於 AB-730 驗證。詳見環境文件與後續 issue 前置條件。

### 考試大綱版本提醒

官方 study guide 目前顯示 **Skills measured as of October 20, 2026** 與 change log；在本指南查證日 **2026-10-14** 尚未生效。不能把該日之後的新權重／agents 區塊描述為查證當日已生效，也不從 change log 反推未展示的舊版完整權重。授課／應試前依實際日期、考試語言與官方頁重新核對；README 的原有 Exam 入口保持不動。

依據：https://learn.microsoft.com/en-us/credentials/certifications/resources/study-guides/ab-730 。課程完成不等於通過考試，也沒有本次已查證的額外 Applied Skills 或 Skillable key 可以補填。

### 課前與時間救援

- 課前 1–2 天由 operator 確認有效 config、既有授權與來源；缺設定繼續標 blocked，不把 example domain 當目標。
- 先區分可 live、可影片、只可討論的 module；驗證當前 license／UI／Notebook limits／Cowork pricing policy，不用購買或改政策補洞。
- 預先挑一個 prompt 改善點與一個人工核對點；生成知識檢核須講師先看，不把 AI answer key 視為官方答案。
- 每 module 留下「我會查什麼證據才採用」的回答；時間不足先刪裝飾／生成圖片演示，不刪人工驗證與資料邊界。
- 本次不下載官方 lab dataset、不生成汽車資料、不執行雲端寫入。後續正式資料設計、測試、核准與灌入是獨立 issue，尚未完成。
