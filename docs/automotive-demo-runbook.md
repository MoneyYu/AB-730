# AB-730 汽车产业演示手册

## 当前状态与使用边界

本手册对应已批准的 `20261015` 新场次。全部故事为虚构「远景汽车制造」，语言为简体中文、币种为 CNY；场次标识不更改学员 README 的日期、Course ID 或教材名称／链接。

**资料已经实际新增，但不宣称完整 live ready：**OneDrive八份Office、Outlook五串15封、Teams／六频道／15条消息写入，以及SharePoint知识库三份Word均已建立。邮件作者／收件者／当前内容、Team精确成员／owner、SharePoint正文／版式与成员已用GET核对。Teams消息v1.0回读实际403，未获批准不切beta；指定网站UI及Copilot功能仍未验收。SharePoint存储后ZIP包元数据／整体hash不同，但原Word正文／版式parts与document relationships相同。写入边界独立审查已通过，代码／公开资产范围冻结；原目录engine／generator／tests已整合，原目录598项engine及37项generator测试通过。用户现已要求提交、推送与合并公开文件，发布状态以对应 GitHub PR 为准。精确证据及首次三面Live exit1的停点限制见 [demo-environment.md](demo-environment.md)。**不得重跑本场次Live或为验收而重发／重传／补owner。** 以下是讲师演示步骤与验算契约，不是 agent 已执行 Copilot 任务的证据。

**用户确认的收尾范围：**浏览器验证已按用户明确指示跳过（`SkippedByUser`）；用户已看到Email及Teams消息。无需再为浏览器验收重跑写入或新增权限。Teams API内容回读403及Copilot未测试仍作为已知限制保留，不宣称上述项目已验证通过；本机整合已经完成，公开文件发布不改变这些租户验收边界。

只使用私有配置中已确认的讲师账号。业务角色是数据作者，不要求讲师登录业务账号。不得为了演示新增用户、改 profile、增加 app consent／订阅，或创建 Calendar、Meeting Chats、agent、Pages、Notebooks、Cowork 任务。

## AB-730 日常资料建立（仅限未来全新场次）

**本节不是重跑 `20261015` 的指令。** 已完成场次只读使用，禁止再次 Live。未来新场次沿用 [seeding skill 的默认入口](../.github/skills/m365-demo-data-seeding/SKILL.md#quick-start)，正常流程只有三步：

1. AI复用现有 profile／schema、generator 和 Office 模板生成虚构资料，输出到全新目录，检查格式、数量、数字／来源一致性及文件可用性。
2. AI一次准备本机 config 和本次准确清单对应的私有 deployment manifest、素材 hash／路径绑定；代码 pins 来自独立受审固定基线。凭据与私有输入留在受 git 忽略的位置；不复用旧 manifest／receipt 作为新场次授权。同一清单已获授权，不逐阶段重复询问。
3. 从仓库根目录执行既有 engine；完成后做必要 API 数量／目标／成员／内容核对，报告建立结果、入口及真实限制后结束。三个变量必须先由 AI 绑定到该全新场次及私有输入：

```powershell
pwsh -NoProfile -File .\seed-data\engine\Invoke-AB730.ps1 `
    -ScenarioDir $ScenarioDir `
    -ConfigPath $ConfigPath `
    -ManifestPath $ManifestPath `
    -Live
```

`-Live` 已含 fresh GET preflight，不另跑相同检查，不直接调用单阶段 writer。generator 不产出 runner／deployment manifest；engine 已有 runner 模板 fallback，不另建 wrapper 或猜 `-Industry`／`-PreflightOnly` 参数。只在原请求选择部分 surface 时使用 `-Phases`；现有全量权限要求保持不变，缺项一次报清，不主动增加权限。

工具源码未变化时复用既有验证与内建安全检查，不重做引擎研发、全套审查／故障注入、全租户盘点或 Office 出版验收。只新增、冲突即停、未知写入不重送；保留脱敏错误、阶段及 exit code，部分成功如实报告。浏览器／Copilot验证不作为日常资料建立的额外完成门；下面的完整教学验收契约属于独立演示范围，不要求为此重传或重发。

## 新资料清单

| Surface | 批准的新目标与内容 |
| --- | --- |
| OneDrive | `AB730-Automotive-20261015`；5 XLSX＋3 DOCX，另最多1份必要 engine proof metadata |
| Outlook | 5 个 `【AB-730 20261015 Automotive】` 主旨串，每串首封＋2次回复；共15封，每封 CC 讲师 |
| Teams | `20261015-AB-730`；默认频道不写故事；5个业务频道各2主贴＋1回复，共15条消息 |
| SharePoint | private site alias `ab730-automotive-20261015`；`DemoKnowledge` 仅3份与 OneDrive 同源的 DOCX；0 custom lists／news／pages |

频道为 `供应链交期`、`BOM与成本`、`零件质量`、`生产影响`、`改善行动`。讲师为新 Team／Site owner，七个已确认业务角色为 members；真实身份只保留于受忽略的私有配置，不写入公开故事源。

## 文件与必须保留的事实

| 文件 | 必须验算／保留的内容 |
| --- | --- |
| `01-BOM-Cost.xlsx` | 12项零件；基准2582、当前2682、增加100 CNY／模拟子总成，不是整车 BOM |
| `02-Supplier-Delivery.xlsx` | 5笔 PO：计划1200件、实收1050件、缺口150件；PO-002／004／005各缺50件；PO-005实际到货日 Unknown |
| `03-Quality-Inspection.xlsx` | Q-001/P-101：1000检20不良；Q-002/P-103：100检10不良；Q-003/P-108：0/0为 N/A；独立生产／库存抽检，不是五笔 PO 普查 |
| `04-Production-Impact.xlsx` | 两线合计计划900、实际830，差异70个模拟总成；不能归因成150件零件缺口 |
| `05-Action-Tracker.xlsx` | 六行动；A-03只有建议负责角色，owner 未确认、due date为空，不得补猜 |
| `06-Quality-Brief.docx` | 上述三个批次、加权分母、不确定事项及人工核对边界；不得改写零件／PO对应关系 |
| `07-Supplier-Response.docx` | 同源五笔 PO、各自缺口与未知日期；模拟回复不是已发送的正式供应商承诺 |
| `08-Improvement-Meeting.docx` | 明确标注模拟会议记录；行动来源、角色、日期、状态与 tracker 完全一致，不是假录音或真实逐字稿 |

加权不良率为 `30 / 1100 = 2.73%`，不是 `(2% + 10%) / 2 = 6%`；0/0不能当作0%或合格。交期、抽检与产出有不同范围／单位，不做未经确认的 join、因果归因、放行、召回或停线决策。

| 行动 | 来源 | 负责角色 | 确认／日期 | 状态 |
| --- | --- | --- | --- | --- |
| A-01 | Q-002 | Quality | 已确认；2026-10-17 | PendingInvestigation |
| A-02 | PO-002 | Supply | 已确认；2026-10-16 | PendingSupplierConfirmation |
| A-03 | PO-005 | Supply | 未确认；日期为空 | ProposalOnly |
| A-04 | P-101 | Engineering | 已确认；2026-10-18 | PendingCostVerification |
| A-05 | LINE-01 | Operations | 已确认；2026-10-16 | PendingImpactVerification |
| A-06 | QualityWeightedRate | Analytics | 已确认；2026-10-16 | PendingReport |

## Module 演示顺序

| Module | 讲师操作／提示词 | 核对与限制 |
| --- | --- | --- |
| M01 | 打开质量文件：「我正在协助质量团队。仅依据当前来源区分已知事实、未知与建议，逐项给出来源键。」 | 不良率、批次和样本范围正确；不用讲师假扮 Quality |
| M02 | 讲解账号／数据边界，核对功能入口与权限指示 | 产品／service plan Enabled 不等于 policy、rollout 或任务实测 |
| M03 | 打开交期邮件与工作簿：「汇总五笔 PO 的缺口，单独列出到货日未知的项；不要推论准时交货。」 | 合计150件，PO-005 Unknown；若引用缺失先核对来源 |
| M04 | 打开知识来源三份 Word，准备问题：「哪些结论已由资料支持，哪些仍需向供应商确认？」 | 必须能打开同源知识文件；不依赖 Custom List rows，不由 agent 创建新 agent |
| M05 | 结合 Teams／Word：「为管理层草拟事实与待确认事项分开的摘要，不添加来源未提供的原因。」 | 来源包含对应文件／消息，避免把模拟回复当正式承诺 |
| M06 | 使用 Excel／Analyst（仅在入口与功能经讲师确认后）：「分别计算 BOM差异、交期缺口与加权不良率；说明分母、单位和零分母处理。」 | 100 CNY、150件、2.73%、N/A；不把70个总成差异与150件直接相减 |
| M07 | 打开模拟会议记录：「列出已确认行动与待确认提案，保留原来源、角色、截止日和状态。」 | A-03单独列为提案；不补日期、不伪造真实 recap／逐字稿 |
| M08 | 纸上拆解：「将资料核对任务拆成输入、人工确认点、产物和停止条件。」 | 不提交 Cowork、不排程、不让自动化做质量放行／召回／停线决策 |

上述提示词尚未由本次 agent 提交／执行；静态数据与公式正确不能冒充 Copilot 回答已验收。未验证功能使用离线资料／提示词评析备案。

## 建立与验收停止条件

执行授权是同一私有 deployment manifest 的**仅新增**清单。先完成 canonical Office／JSON 验证、独立实作审查，再重新做 fresh GET；新 folder、Team、site alias／displayName 或邮件主旨有任何匹配就停止，不收养、不覆盖、不追加后缀，不重跑旧 Ford。

创建过程只使用本 run 在进程内取得的新资源 IDs。退出 run 后只做 GET 验证；saved receipt、name hit 或旧 IDs 不是修补／补 owner 的授权。写入结果不确定时停止、不自动重试、不回滚或删除，保留脱敏 partial 结果。

写后逐项核对8份 OneDrive Office 文件、15封邮件的作者／收件者／讲师 CC、5业务频道15条消息、private知识站点3份 Word 和讲师存取，再由主 agent 在用户指定的既有 Edge「AB-730」窗口验收。只剩 `about:blank`、连接未确认或工具不可用时，UI维持 blocked，不能宣称总体完成。
