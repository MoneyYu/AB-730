# AB-730 ab730-automotive-20261015

本场次为 AB-730 教学模拟，使用简体中文（`zh-CN`）和 CNY；故事中的企业与供应商均为虚构。仅绑定已批准、现有且启用的账号。执行器严格使用 `additive-fresh-targets-only` 模式，不创建用户、修改 profile，或覆盖／删除／收养既有资源。

## 当前场次状态

四个surface均已实际新增：OneDrive五份Excel＋三份Word；Outlook五串15封；Teams及六频道、15条消息写入；SharePoint private site的DemoKnowledge三份Word。邮件内容与收件者、Team精确成员／owner、SharePoint正文／版式及owner／members已用GET核对。Teams消息v1.0回读实际403，尚未获准切beta；网站UI／AI未验收。SharePoint ZIP整体hash／大小不等于来源，原Word正文／版式parts及document relationships相同，包元数据差异已记录。

用户之后明确允许跳过浏览器验证，并确认Email及Teams消息可见。浏览器状态为`SkippedByUser`而非主agent验收通过；不再为此操作浏览器、切beta、增加权限或重跑写入。Teams API内容回读及AI功能的既有未验收限制不变。

**本场次禁止再次执行任何Live／上传／发信／migration／membership写入。** 首次三面Live在Teams成员步骤附近exit1，后续GET确认成员完整；没有重跑，SharePoint由同一已审查入口单独建立一次。原错误详情未保留，不能推断原因；saved IDs／receipt不授权修补。原仓库engine／generator／tests已完成本机整合，原目录离线测试实际为598项engine、37项generator通过；未请求commit／push，不宣称已经发布或fresh clone包含本机未追踪文件。最新证据及限制见仓库`docs\demo-environment.md`。

## 创建前接口参考（本场次已建立，不得再次执行）

- 私有 `config.json` 只保留在已被 Git 忽略且未追踪的 `seed-data\config.json`；**不得**复制到本场次目录。`config.json.example` 是中立模板，不含实际身份或凭据。
- root config 的 `adminUpn` 与 `demoUserUpn` 均绑定批准清单的 `roles.Admin.upn`；八个 `roles` 各含 `upn`，与私有清单完全一致且互不重复。`tenantId`／`clientId` 与清单 `target` 一致，凭据只在 root config 中。
- root config 使用 `"filesSourceDir": "scenarios\\ab730-automotive-20261015\\DEMO-FILE"` 与 `"teamDisplayName": "20261015-AB-730"`。
- 私有批准清单必须包含已批准 canonical spec 的 `dataSpecPath`／`dataSpecSha256`、完整 `artifactHashes` 和执行依赖的 `codeHashes`。这些 hash 由父流程在内容核对／实作审查完成后固定；runtime 不得自行重算、补齐或更新批准值。
- 默认只执行fresh preflight：检查config忽略／未追踪状态、清单／身份／目标绑定、素材SHA-256、JWT claims和碰撞；当前资源已存在，因此碰撞停止是预期，不是重新写入的理由。Graph集合最多20页，未穷尽或读失败不视为不存在。

```powershell
pwsh -NoProfile -File .\seed-data\scenarios\ab730-automotive-20261015\run.ps1 `
    -ConfigPath .\seed-data\config.json `
    -ManifestPath <path-to-ab730-deployment-manifest.local.json>
```

- `-Live`只用于尚未创建、已获准确批准且独立审查／fresh preflight通过的目标。**本场次不再有可执行的Live步骤**，后续仅GET验收；不要把创建接口当作恢复命令。

## 写入与验收边界

批准清单为 OneDrive 8份 Office 文件（另最多1份必要 proof metadata）、5个邮件串共15封、5个业务频道共15条消息，以及独立 private SharePoint 站点中的3份同源 Word 文档。讲师为新 Team／Site 唯一 owner，七个业务角色为 members；默认频道不写故事，不创建 custom lists／news／pages。

碰撞、未知、权限或不完整响应必须停止。结果不确定的 write 不自动重送；不按旧 receipt／名字接管资源，不删除或修补部分创建结果。退出后只做 GET 验收，不补授予权限。

本机资产检查、mock 测试与实作审查不代表真实云端、讲师网站 UI 或 Copilot 功能验收。实际状态以仓库 `docs\demo-environment.md` 为准；未发布前不宣称 fresh clone 已包含本机新增执行器与素材。
