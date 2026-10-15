# AB-730 demo 環境與唯讀盤點界線

## 交付狀態

本文件保留 **2026-10-14 文件實作交接時的本機盤點快照**，並補充 **2026-10-15 限定範圍的租戶唯讀盤點**。後者不是全租戶、內容完整性或講師介面驗收。教材與產業教學設計見 [teaching-guide.md](teaching-guide.md)。

## 最新执行状态：AB-730 专用盘点

完整 A→B→C（盘点、离线生成、精确清单内的实际租户创建及只读验收）已获用户批准。以下是批准后的实际结果；后方原有章节是历史快照，不再表示当前缺少有效配置或讲师尚未确认。

**最新验收范围调整：**用户明确表示「可以跳過瀏覽器驗證，我看 Email 跟 Teams 訊息都有」。浏览器验收现为 `SkippedByUser`，不再阻挡本次资料新增收尾；用户确认Email／Teams消息可见，但这不等于主agent已完成浏览器或全部15条Teams内容的逐项验收。Teams API回读403仍如实保留，不再为此重复尝试、切beta或增加权限。Copilot功能测试未执行。原目录engine／generator／tests已完成本机整合及离线回归；用户现已要求提交、推送并合并全部已完成的公开文件，发布状态以对应 GitHub PR 为准，私有配置与运行记录不发布。

**原目录整合与修正证据：**20个受审文件按显式allowlist落地，12个现存文件在修改前有私有备份；整合时原60个非allowlist文件及学员README字节不变。原目录测试发现一个隔离目录布局依赖：测试helper从session父目录寻找canonical spec，导致`581/17 RED`；改为repo内公开`generator\tests\ab730-data-spec.json`并缺失即明确失败，未改engine或批准素材。相同原目录命令`pwsh -NoProfile -File .\seed-data\tests\Test-SeedEngine.ps1 -FixtureScenarioDir .\seed-data\scenarios\ms4018-ford-auto-workshop-b-20260929`实际输出`Passed: 598 / Failed: 0`、exit0。`uv run --default-index https://packagefeedproxy.microsoft.io/pypi/simple/ --with pytest --with pyyaml --with jsonschema --with openpyxl --with python-docx pytest .\seed-data\generator\tests -q`最终实际输出`37 passed in 9.54s`、exit0。两者均为离线验证，不是再次Live。该限定修正计划及最终整合实作均已由原Duck批准，不以历史审查替代。

**本机整合最终复核：**原独立Duck实际核对19文件与受审源字节相同、唯一portable helper差异及AST，并检查父流程RED／GREEN与generator日志、整合／隐私报告及文档／经验行，结论`Approved / Blocking 0`。父流程最终检查12原备份、58保护文件未变＋2项显式README文档更新、主README完全未变、10代码／12素材批准pins不变、私有输入仍忽略且未追踪、25候选无新私密值，均通过。旧测试中原有地址常量仅保留本机，不因此批准公开发布。请求`gpt-6.1-sol / max`，runtime映射无法独立确认；reviewer未重跑套件或读取真实配置，未执行网络／浏览器／Live。原高风险写入边界Council批准范围保持冻结，本轮没有产品代码新差异或新增reviewer；本次完成资料新增与本机交付，不代表完整UI／消息API内容／AI功能已验收，也没有commit／push。

| 检查 | 最新证据 | 状态／边界 |
| --- | --- | --- |
| 认证与讲师 | token 与私有批准清单的 tenant／client、Graph audience、expiry 相符；讲师及七个业务账号各自 enabled、身份不同 | `Verified`；Admin 不承担业务角色 |
| License | 讲师的 6 项产品含 `Microsoft Copilot for Microsoft 365`、`Microsoft 365 E5 (no Teams)`、`Microsoft Teams Enterprise`；8 项名称含 Copilot 的 service plans 映射为 Enabled，Unknown plan 为 0 | SKU 与 plan 状态有证据，不等于 UI、policy、rollout 或功能实测 |
| 业务账号 | 七个角色均有 E5 (no Teams) 与 Teams Enterprise；未创建账号、修改 profile 或增加许可 | `Verified`；不要求讲师登录业务账号 |
| OneDrive | 已实际新增一个批准文件夹及5 XLSX＋3 DOCX；一次Live执行exit0，最终精确GET核对八项ID／名称／大小／parent，私有receipt为Verified，八个批准hash相符 | API新增及结果 `Verified`；讲师网站UI仍未验收，不得重跑／收养／修补 |
| 写入前碰撞 | 选中三面的真实只读preflight通过后执行一次Live；该次在Teams成员步骤附近停止，未进入SharePoint；随后仅对仍全新的SharePoint执行一次选中phase Live | 没有重发邮件、重跑Team／OneDrive、收养、修补、删除或增加权限 |
| 浏览器 | 用户明确允许跳过，并确认Email／Teams消息可见；先前主流程Computer Use仅取得原窗口的可信URL、六频道及部分消息文本 | `SkippedByUser`，不是Playwright／主agent完整UI验收通过；不再操作浏览器，不另开profile或绕过身份核对 |
| 新资产 | v3 修正版有5 XLSX＋3 DOCX；实际内容与 canonical spec 一致，原目录generator 37测试通过；Excel原生重算67公式、0错误；Word原生PDF六页已由主 agent逐页检查 | 本机资产／runbook及写入边界独立审查均已通过、范围冻结；原目录engine已整合、598项离线测试通过；公开文件发布由对应 GitHub PR 追踪；云端验收见分项，不等于网站UI／AI通过 |
| Outlook | 已实际发送五串共15封；父流程GET逐封核对唯一internetMessageId、发件人、To／CC、主旨及uniqueBody中的当前内容，全相符 | `Verified`；15封均CC讲师；没有重发或读取其他信箱正文 |
| Teams | 已新增Team、默认及五个业务频道，完成15条消息写入调用与migration；GET确认六频道、八个精确成员、Admin唯一owner及backing group集合 | 成员／频道 `Verified`；消息v1.0 GET实际403，内容回读 `Forbidden`，不能宣称消息内容完整验收 |
| SharePoint | 一次独立选中phase Live exit0；GET确认新private group／site、DemoKnowledge三份Word、Admin唯一owner及七个业务member | 正文与版式parts、原document relationships逐项相同；ZIP大小／整体hash不同，新增内部customXml及包元数据，不伪称字节相等 |

**当前审查门与执行证据：**原B7对象边界已解决；sanitized error不保留native response/request/content/headers/inner/raw引用，只保留typed HTTP status及0..3600秒的Retry-After scalar。原Duck最终Approved／Blocking0；同一Council独立22个helper及2个完整native OAuth probes通过，最终Approved／Blocking0，范围冻结，不再增加reviewer或同范围授权门。请求`gpt-6.1-sol / max`，底层模型／effort映射仍不能独立证明。父流程终版mandatory harness为`598/0`（历史实际运行），本次再次核对十项实际代码hash及十二项input pins；root认证值不变，八角色绑定与私有批准源一致。直接调用已审查隔离`Invoke-AB730.ps1 -Phases @('Emails','Teams','SharePoint')`，真实输出`Preflight-only verification succeeded (0 writes executed)`、exit0；之后按既有授权执行一次同选项`-Live`，不使用standalone phase或默认All。以下旧审查段落保留为历史过程，不再表示当前仍有代码审查Blocking。

**实际停点与只读验收：**三面Live exit1，私有receipt保留五个完成的邮件conversation IDs；首个父wrapper只保留失败布尔值，未保留原始脱敏错误码，因此**不能确认Teams当时停止的根因**，也不推断为最终一致性。之后只读GET确认Team成员／owner／backing group均完整，没有执行修补；仍全新的SharePoint以同一已审查入口`-Phases @('SharePoint') -Live`单独建立，exit0。父流程邮件验证输出`Outlook GET verification:5threads/15messages;exact sender/To/CC/subject/current uniqueBody;15uniqueInternetMessageIds;PASS;GET=18;0datawrites`。

Teams消息回读的实际403明确要求`ChannelMessage.Read.All`或`ChannelMessage.Read.Group`；已有`Group.ReadWrite.All`在本次请求中未满足服务端检查，官方表格不能替代实测。先前询问只读beta时未取得批准，**没有批准即不执行beta、不新增consent**。用户之后已允许跳过浏览器验证并确认Email／Teams可见；保留API内容回读限制及AI未测试状态，不把用户观察记为主agent实测。原仓库本机整合与公开发布分开：前者已有实际离线通过证据，后者未请求、未执行，不再作为本次交付阻断。

**后续浏览器观测（不是验收通过）：**原窗口的供应链频道可访问性文本包含PO-005准时状态Unknown、PO-002对应A-02于2026-10-16到期、A-03仍为ProposalOnly等批准情境内容，但没有逐项核对全部15消息及作者。一次原生Outlook刷新遭用户输入中断，按工具要求重读原窗口后仅重试该动作一次；刷新后观察到Outlook官方cloud域名及已加载页面，仍未取得账号菜单身份。Chrome DevTools默认Chrome连接不可用，未启动其他profile；未增加权限、执行beta、提交AI任务或重跑租户写入。详细去敏观测持久保存在本机session的`ab730-browser-observation.json`，不是fresh clone交付物。以下原文档审查结论不涵盖此后新增的观测记录。

**结果文档／经验复核：**原独立Duck已完成当前进度、runbook／场次禁止重跑说明及skill新增reference rows的限定审查，`Approved / Blocking 0`；重读最新状态后确认未知停因、Teams 403／未批准beta、SharePoint包字节差异及UI／AI／发布限制均如实保留。请求`gpt-6.1-sol / max`，runtime映射未独立确认。未重开已冻结engine／资产审查，也未新增审查者；此批准不解除任何线上验收限制。

SharePoint三份下载的`word/document.xml`与所有原Word正文／版式parts均与冻结来源一致，原document relationships保留；新增document relationships仅指向内部customXml。包整体hash／大小不相同，实测差异限于包声明／relationships／customXml／core/custom properties及新增trash parts，未追查触发者。保留现状，不重传以强行追求ZIP相同。父流程输出`SharePoint:exact3documents;all original Word content/layout parts identical;original relationships preserved;added document relationships internal customXml only;Adminowner+7members exact:PASS`，验证0datawrites。

## 分阶段历史记录（不代表当前阻断）

以下保留当时状态及修正经过；当前结果以上方最新执行状态及原目录整合证据为准，旧文中的「待完成／Live禁止／零写入」不是当前执行状态或再次写入的授权。

**实际命令与结果：**session-only `pwsh -NoProfile -File ab730-inventory-v2.ps1 -ConfigPath <ignored-local-config> -PlanningInputPath <private-input> -DataSpecPath <canonical-spec> -OutputPath <session-result>` 输出 `Inventory: roles=8; targets=4; GET=17; blockers=0; UI=NotTested`，exit 0；`pwsh -NoProfile -File Test-AB730Inventory.ps1` 输出 `PASS: 28 inventory boundary assertions; no network calls.`。这些脚本／结果留在本机 session，不是 fresh clone 交付物。GET 集合上限 20，检测 malformed response、循环、外域、用户／folder／query 改变；OAuth POST 仅认证，没有租户数据写入。

盘点读取器独立 Rubber Duck 审查及定向修复复查已完成，最终 `Approved / Blocking 0`。请求 reviewer 为 `gemini-4-argon / high`；runtime 无法独立确认实际映射。Graph 使用 `mailFolders('inbox')/messages` 的合法 nextLink 已通过限定 Inbox 等价语法修复，不放宽到其他信箱／folder。许可名称来自 Microsoft 官方 [product／service-plan mapping](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference)，本地 CSV SHA256 为 `A7A0FACF663BE9D1CE880FF570188C8F186B15D192A79482D661822D9CE9B9E5`；不使用不支持 app-only 的 `/licenseDetails`，不扩权查 `/subscribedSkus`。

**离线验证证据：**隔离目录运行 `uv run --default-index https://packagefeedproxy.microsoft.io/pypi/simple/ --with pytest --with pyyaml --with jsonschema --with openpyxl --with python-docx pytest seed-data\generator\tests -q` → `37 passed in 8.17s`。session `ab730-office-validation.ps1 -ScenarioDir <corrected-v3> -OutputDir <new-session-directory>` → 8文件成功打开、67公式0错误；Word三份各2页，六页均已检查，ID／日期／表格行可读、无跨页拆行或孤立尾页。主 agent另用安装的 Excel UI 打开五份相同字节的工作簿，核对表头、日期、百分比与数值，未编辑／保存，已关闭自有窗口；所有来源文件 hash 仍与 manifest 相符。Excel打印／PDF版面未验收，长说明可由单元格的完整值读取；此本机 Office 验证不代替讲师租户网站 UI 或 Copilot 功能实测。

**公开资产实作审查：**独立 Rubber Duck 已完成 canonical 内容、八文件 hash、消息数量与不确定性、导出隐私／拒绝覆写、runbook，以及 Word 版面经验的 reference/application 检查，结论 `Approved / Blocking 0`。请求模型／effort 为 `gpt-6.1-sol / max`，与资产作者跨供应商；reviewer 无法独立证明实际模型／effort 映射。既有37测试与原生 Office 证据被审查使用而非由 reviewer 重跑。此批准不包含 runner／原目录整合、live engine／数据边界、发布、租户／网站 UI／Copilot 功能或 Excel打印版面。

**引擎离线证据：**主 agent在隔离目录运行 `pwsh -NoProfile -File .\seed-data\tests\Test-SeedEngine.ps1 -FixtureScenarioDir .\seed-data\scenarios\ms4018-ford-auto-workshop-b-20260929` → `Passed: 547 / Failed: 0`，exit 0。这是 mock harness，不是旧场次 live runner。整合候选的8个 Office 文件与冻结 v3 来源逐个字节 hash 相等。实际 runner 已实现必填 `ConfigPath,ManifestPath`、默认只读 preflight、显式 `-Live`，无需额外同范围交互授权；独立写入边界审查与原目录整合仍待完成。

**整合隐私检查的定向修正：**新增 reference-fixture helper 原沿用了旧测试账号常量；作者现已改为 `example.invalid` 合成身份，调整直接耦合的断言并增加2项source/output隐私测试，报告 `537/12 RED → 549/0 GREEN`。主 agent已独立调用实际 helper生成到全新session目录，检查源码和2个JSON输出的7个合成角色，结果 `Parent reference-fixture privacy: PASS`；29个新公开候选文件（含 Office XML）真实私密值扫描 `PASS`。完整suite将在最新 engine修正后由主 agent实际重跑，作者报告不替代它。旧suite的原27行domain断言原本均为既有内容；只修改该新fixture直接依赖，不顺手重构剩余旧ownership mock，也不批准旧私有fixture公开发布。没有触及真实账号 profile、旧 packs 或租户资料。

**实际写入边界审查未批准：**同一最小 Council 通过 mock 重现确认7项 Blocking：批准spec／完整inputs／code未固定；Teams owning碰撞漏检及按名称取得外来ID；nextLink查询范围扩张；读错误／malformed回应用作absence；SharePoint来源逃逸／名称PUT覆写；owner冲突／null响应仍继续成功；实际错误路径未去敏。保持零租户写入、禁止 `-Live`；修正计划按这7项最小措施及针对性RED／GREEN执行，由原作者修正、同一reviewer仅复查直接影响。请求 reviewer `gpt-6.1-sol / max`，实际effort映射无法独立确认。离线历史547/0不能覆盖这些已证实缺口。

先前只读 Computer Use 检查重新发现并定位既有 Edge「AB-730」窗口时返回 `url: missing — fail-closed`；当时未尝试截图／导航绕过。后续已取得可信URL，见最前方最新浏览器观测；讲师登录身份及完整网站内容仍未验收。此 UI 缺口不阻挡通过安全门后的既定 Graph 新增，但不能据此宣称总体完成。

**最新定向修正证据：**原作者先捕获23项针对性失败（`550/23 RED`），实施后报告`573/0 GREEN`；主 agent实际运行相同mandatory命令，输出 `Passed: 573 / Failed: 0`，exit 0。批准 spec 实际 hash 与冻结文件相符，4份叙事／预期 JSON 字节未变、8个 Office hash 未变；父流程已固定 spec 与12项实际input hashes，未自行改变批准内容。必须等待同一 Council清除7项Blocking，再固定通过的code hashes、精确整合及fresh preflight；测试通过不是Live授权安全门完成。

**限定复查更新：**同一 Council已清除B4／B5原重现与fixture隐私，并确认mandatory pins／10档code closure／矩阵负例、标准202headers、Fresh owner409立即停止、PUT409不重试；这些范围冻结。剩余B1 helpers先于hash gate执行、B2 operation／GET改绑resource anchor、B3 top100→100000仍发第2GET、B6只检查成员数量而非精确identity／role集合、B7邮件／preflight／upload错误去敏，原作者仅补这5项及直接影响。

实际安装SDK的进程级request context默认`MaxRetry=3`，engine原无设置；父流程无认证／网络测试`Set-MgRequestContext -MaxRetry 0`再GET查证0，通过后finally恢复原3，未修改machine/global profile。原禁止自动write retry也必须在B7实际SDK层落实，不能仅靠mock409负例。支持依据：[Set-MgRequestContext](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/set-mgrequestcontext?view=graph-powershell-1.0)。当前仍零租户写入，Live禁止。

**06:07 用户要求后的 OneDrive 实际结果：**先完成既定八份 Office 新增，不等待通用engine或浏览器。session原生REST小执行器无SDK／外部helper、零自动重试、内存字节hash固定、create-only、返回ID锚定、逐笔私有receipt及最终精确GET。receipt唯一Blocking经两行删除及实际RED→GREEN回归修正后，原Duck及Council均Approved／Blocking0，冻结SHA256为`4FE627D1E66AAB5C0BB37F02566D5438FF7AB52A0478B3110FCFD17C10914700`。主 agent真实fresh preflight通过，再执行一次 `Invoke-AB730OneDriveNow.ps1 <explicit-approved-paths-and-hash> -Live`，实际输出 `OneDrive: 8 approved Office files created and exact inventory verified`、exit 0。随后主 agent独立检查私有receipt：Verified、5 XLSX＋3 DOCX、八项批准hash相符、没有provider URL字段；另一次只读GET确认新folder身份及普通登录保护链接，仅提供给用户、不保存或公开到issue。文件确已建立，但网站UI／AI未验收。

**剩余通用engine修正：**主 agent以相同mandatory命令实际输出 `Passed: 593 / Failed: 0`、exit 0；原Duck及Council已通过选中phase路由增量，Council清除B1／B3。仍有四项具体局部缺口：GET Team空／缺displayName仍接受、exact成员检查误伤legacy presence-only、去敏重建exception丢typed503导致错误absence、orchestrator实际OAuth异常未去敏。原作者只修这四项及直接影响，原Council定向复查，不新增Council或重审冻结范围。后续邮件、Teams、SharePoint仅选各自FreshOnly；已经完成的OneDrive只允许GET，不收养、修补或重新调用完整四surface流程。此API安排不解除UI／AI未验收限制。

**后续局部核对：**四项修正后主 agent实际mandatory suite为 `597/0`、exit 0；原Council独立25个helper／直接影响检查与4个完整gate／OAuth probes通过，原四项清除。Duck进一步使用native `HttpResponseException` 检查实际error object，确认Message虽去敏，原Response.RequestMessage.Content／Response.Content仍可携带凭据或能力URL；全为synthetic离线重现，没有真实Graph或私人输入。当前只修这项B7直接对象边界：不保留原response/request/content/headers/引用，仅保留safe typed HTTP status；必要Retry-After只提取validated scalar，并保留503不变absence／有界GET回归。其他identity／legacy／routing范围冻结，两位原reviewer只做局部核对。原仓库新scenario16份公开文件已create-only落地、12素材hash相符，16文件含Office XML隐私检查通过；既有原engine未覆盖，后续使用已审查隔离代码与原repo私有输入执行选中phase。

### M01–M08 readiness

| Module | 已确认的资料／许可前置 | 尚未验收／备案 |
| --- | --- | --- |
| M01 | 讲师 Copilot 产品／plan、质量摘要与检验数据规格 | Copilot Chat／Word UI、真实来源文件及回答未实测；可用质量事实／推测辨识练习 |
| M02 | 讲师产品／plan 与既有 app-only 身份核对 | UI、policy、EDP 指示未查证；讲解官方边界，不推论许可足够 |
| M03 | 五邮件thread共15封已建立并核对当前内容、收件者；八份Office已建立 | Outlook／Office UI未验证；可先使用已提供文件夹及离线来源 |
| M04 | 讲师 Graph-grounded chat／SharePoint plan 为 Enabled；三份知识 DOCX 规格 | Researcher／agent UI 与站点来源未验收；不创建 agent，不执行研究任务 |
| M05 | 五频道、三份文档、明确事实／未知与跨部门来源矩阵 | 线上来源与操作账号 UI 未验收；Pages／Notebooks 不创建 |
| M06 | 本机 Excel已实算：BOM 2582→2682 CNY、交期缺口 150 件、加权不良率 30／1100＝2.73%、零分母 N/A、生产差异 70 个模拟总成 | 租户 Office／Analyst UI 和 Copilot回答未验收；本机计算不能代替功能实测，指标不能混单位或因果归因 |
| M07 | 六项行动，A-03 owner／deadline 未确认；明确模拟会议记录 | 不录音、不伪造真实 recap／逐字稿，不启动 Facilitator |
| M08 | 新数据契约与纸上任务拆解 | Cowork UI／billing／policy 未验证，禁止提交任务或排程 |

## 历史快照：初次文件交接

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

## 2026-10-15 租戶唯讀盤點結果

使用 operator 指定的本機 `seed-data\config.json` 完成 client credentials 認證，token 的 tenant、client、Graph audience 與有效期限檢查通過。設定已確認被 Git 忽略且未追蹤。既有 token 包含 user、group、file、site 等可用角色；沒有新增 consent 或權限，也沒有使用其中的寫入能力。

盤點只依四份既有 scenario 範本宣告的角色 UPN、精確 Team 名稱、site alias 與指定 OneDrive folder 查詢；Microsoft Graph 共 **31 次 GET**，所有集合跟隨 `@odata.nextLink`，頁數上限為 20。僅保留角色標籤、數量與查證狀態，不保存原始 Graph 回應、信件、訊息、文件內容或下載網址。OAuth POST 僅用於取得 token。

| 宣告情境 | 角色帳號 | Teams | OneDrive | SharePoint |
| --- | --- | --- | --- | --- |
| 原始 Ford | 9 個宣告角色均找到且啟用 | 精確 group 名稱查詢為 0 筆；不推論其他名稱的 Team 是否存在 | 5 個預期檔名皆找到；folder 直接子項目共 6 筆 | 未宣告，未查 |
| Ford WorkshopB | 9 個宣告角色均找到且啟用 | 精確匹配 1 個已 provision 的 Team；6 個 channel，涵蓋 5 個宣告 business channel；1 位 group owner，包含範本候選 Admin | 5 個預期檔名皆找到；folder 直接子項目共 6 筆 | 未宣告，未查 |
| 舊 Product Support | 2 個宣告角色查詢均回傳 HTTP 404／`Request_ResourceNotFound`；保留失敗，不採用此角色映射 | 未宣告，未查 | 未宣告，未查 | 精確 group alias 查詢為 0 筆；未查後代 site／list |
| 日期版 Product Support | 2 個宣告角色均找到且啟用 | 精確 group 名稱查詢為 0 筆 | 未宣告，未查 | 精確 group alias 查詢為 0 筆；未查後代 site／list |

成功找到的範本候選 Admin 回應含 6 筆 license assignment，其餘成功找到的角色各有 2 筆。**assignment 數量不能辨識 Microsoft 365 Copilot 產品、有效 service plan 或實際功能**；此盤點尚未解析 license SKU／service plan，不能據此宣稱具備 Copilot 授權。

以下限制仍存在：

- 本機認證缺口已解除；但 root 設定只有認證欄位，沒有 operator 已確認的 `adminUpn`。範本候選 Admin 的 enabled／owner 查證不代表講師已確認登入帳號。
- OneDrive 只核對檔名，不代表內容、版本、數字一致性或講師帳號的存取已驗證。多出的直接子項目不刪除、不收養、不解讀為故障。
- Teams 僅核對 group／owner／channel metadata。Token 未包含 `ChannelMessage.Read.All`，本次沒有讀取訊息或嘗試 beta；這是既有 token 角色的觀察，不是已執行訊息 GET 後得到的 403。
- Outlook 未讀：operator 尚未確認，且本次不以全 Inbox 掃描尋找 scenario。SharePoint alias 查詢無匹配，也不代表全租戶沒有 SharePoint site。
- 既有 Ford 資源不是 AB-730 汽車課程專用資料；未生成或灌入本課資料，未執行正式 runner，也未驗證 Copilot UI。

**證據與重現：**本次 session-only `ab730-inventory.ps1` 執行結果為 exit 0，輸出 `graphRequestCount: 31`；脫敏結果保存在本機 session 的 `ab730-tenant-inventory-result.json`，認證結果另存 `ab730-tenant-auth-result.json`。這些不是 repo 交付檔案，fresh clone 不含它們。重新盤點需由 operator 在本機建立受忽略的設定，沿用下方唯讀查證規則，不能把舊結果當成當日證據。

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

本次不修改 generator、schema、preflight、runner 或 tests。2026-10-15 依 operator 要求精確增補 `.gitignore`，保護 root 與 scenario 的本機 `config.json`，保留 `config.json.example` 可追蹤。其他缺口留給後續資料 issue，比直接複製既有 scenario 安全，也避免錯課號／錯語系與共享租戶寫入。

## 能力與資料缺口矩陣

「未驗證」表示尚無該項實際證據；「blocked」表示已知缺前置；403／429 不可轉寫為資源不存在。下表以 2026-10-15 結果更新；範本角色與資源查證不等於 operator 確認或 AB-730 demo 驗收。

| Surface／功能 | 後續需確認的能力與資料 | 目前結果 | 課堂備案 |
| --- | --- | --- | --- |
| 操作帳號 | 有效設定的 `adminUpn`、account enabled、與 token tenant／client claims 一致 | 認證與候選角色查證通過；operator 尚未確認，blocked | 投影片；不登入其他角色 |
| Microsoft 365／Copilot license | 既有 assignment、產品 label、app 支援、政策 | 已讀 assignment 數量；SKU／service plan 與功能未驗證 | 講解官方條件，不宣稱 live 可用 |
| Copilot Chat | 工作入口、EDP、web／app context、Work IQ／來源權限 | 未驗證 | Prompt 評析與來源辨識 |
| Outlook | 僅 scenario 宣告且獲授權的內容、操作帳號的存取 | blocked；未讀信件 | 通知文字設計，不開收件者或寄送 |
| Teams | 既有 Team／channel、owner／member、訊息完整性與角色可讀 | WorkshopB metadata 有證據；訊息內容、member 與講師 UI 未驗證 | 口頭協作流程；不發訊息 |
| OneDrive | 已核准 folder／檔案、manifest 對照、操作帳號能讀 | 兩個 Ford folder 各找到 5／5 預期檔名；內容與講師存取未驗證 | 不上傳來源、不生成替代 Office 檔案 |
| SharePoint | 已核准 site、library、list、items、操作帳號可見性 | 兩個 Product Support 精確 alias 查詢為 0 筆；未驗證後代內容 | Agent 知識來源設計 |
| Excel／Analyst | BOM／交期／批次檢驗合成資料、欄位、數字一致性、功能權限 | 未驗證；資料尚未產生 | 分母／join／單位檢查練習 |
| Word／PowerPoint／Researcher | 已核准摘要、引用來源、app／agent 可用性 | 未驗證 | 報告結構與研究問題 |
| Facilitator／Teams recap | 會議類型、license、policy、同意／逐字稿／語言 | 未驗證 | 已核准文字紀錄或只談議程 |
| Pages／Notebooks | storage／service plan、enablement、source／分享控制 | 未驗證 | 畫布與參考來源設計；不建立或分享 |
| Cowork | assigned license、enablement、usage-based billing、model／subprocessor policy | 未驗證；本次禁止提交任務 | 官方影片與紙上任務拆解 |
| Calendar／Meeting Chats 寫入 | 目前沒有新增這些 surfaces 的教學／寫入核准 | 不在本次範圍 | 不為課程完整性擴張範圍 |

官方前置條件與實際測試保持分欄。Researcher／Analyst 的 add-on、Facilitator 的啟動條件與 Cowork 計費不從 template 或 seeder app-role 推論。

## 取得有效設定後的唯讀查證規則

以下是每次盤點的查證規則；已執行範圍以上方日期紀錄為準，不代表取得 write 授權：

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
- **保護本機：** 精確忽略 secrets／receipt／runtime／下載檔，保留使用者既有修改；本次只增補本機 config 排除規則，後續 runtime／receipt 仍須另查保護，不將整個未追蹤 `seed-data\` 納入提交。
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

以上連結僅證明官方條件，不證明本租戶符合。2026-10-15 已完成本機認證與限定 metadata 盤點；未完成項目為 operator 確認、訊息／文件內容完整性、license SKU／service plan、講師功能 UI、本課產業資料與正式 demo。整體備課驗收仍有未驗證或 blocked 項目。
