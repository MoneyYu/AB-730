function Initialize-PL7008ReferenceFixture {
    param(
        [Parameter(Mandatory)][string]$FixtureRoot
    )

    if (Test-Path -LiteralPath $FixtureRoot) {
        Remove-Item -LiteralPath $FixtureRoot -Recurse -Force
    }
    $scenarioDir = Join-Path $FixtureRoot 'seed-data\scenarios\pl7008-it-helpdesk-20260831'
    $demoDir = Join-Path $scenarioDir 'DEMO-FILE'
    $docsDir = Join-Path $FixtureRoot 'docs'
    $ms4022Dir = Join-Path $FixtureRoot 'seed-data\scenarios\ms4022-productsupport'
    New-Item -ItemType Directory -Force -Path $demoDir | Out-Null
    New-Item -ItemType Directory -Force -Path $docsDir | Out-Null
    New-Item -ItemType Directory -Force -Path $ms4022Dir | Out-Null

    # 1. Legacy MS4022 archived runner
    $ms4022Run = @'
param(
    [ValidatePattern("^\d{8}$")][string]$Date = (Get-Date -Format "yyyyMMdd"),
    [switch]$WhatIf
)
$ErrorActionPreference = "Stop"
if ($WhatIf) {
    Write-Output "Legacy MS-4022 runner is archived; it does not execute against the hardened engine."
    Write-Output "Use scenarios/ms4022-productsupport-20260929/run.ps1 -PreflightOnly instead."
    return
}
throw "Legacy MS-4022 runner is archived; live execution is disabled. Use scenarios/ms4022-productsupport-20260929/run.ps1 -PreflightOnly instead."
$configPath = Join-Path $PSScriptRoot "config.json"
'@
    Set-Content -LiteralPath (Join-Path $ms4022Dir 'run.ps1') -Value $ms4022Run -Encoding UTF8

    # 2. PL-7008 run.ps1
    $pl7008Run = @'
param(
    [ValidateSet("technology", "finance")][string]$Industry = "technology",
    [switch]$PreflightOnly
)
$ErrorActionPreference = "Stop"
$scenarioRoot = $PSScriptRoot
$configPath = Join-Path $scenarioRoot "config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Missing config.json in $scenarioRoot. Copy config.json.example to config.json first."
    exit 1
}
$engine = Join-Path $scenarioRoot "..\..\engine" | Resolve-Path
function Invoke-SeedPhase {
    param([string]$Name, [scriptblock]$Action)
    try {
        & $Action
    } catch {
        Write-Error "Phase '$Name' failed: $($_.Exception.Message)"
        exit 1
    }
}
Invoke-SeedPhase "Connect" { & (Join-Path $engine "Connect-GraphApp.ps1") -ConfigPath $configPath | Out-Null }
Invoke-SeedPhase "Preflight" { & (Join-Path $engine "Invoke-SeedPreflight.ps1") -ConfigPath $configPath -ScenarioDir $scenarioRoot -Surfaces @('Users','Files','Emails','Teams','SharePoint') }
if ($PreflightOnly) { return }
Invoke-SeedPhase "Profiles" { & (Join-Path $engine "Invoke-SeedUserProfiles.ps1") -ConfigPath $configPath -ProfilesPath (Join-Path $scenarioRoot "user-profiles.json") -Industry $Industry }
Invoke-SeedPhase "Upload" { & (Join-Path $engine "Invoke-UploadFiles.ps1") -ConfigPath $configPath -ManifestPath (Join-Path $scenarioRoot "files-manifest.json") }
Invoke-SeedPhase "Emails" { & (Join-Path $engine "Invoke-SeedEmails.ps1") -ConfigPath $configPath -EmailsPath (Join-Path $scenarioRoot "emails.json") }
Invoke-SeedPhase "Teams" { & (Join-Path $engine "Invoke-SeedTeamsChannel.ps1") -ConfigPath $configPath -TeamsMessagesPath (Join-Path $scenarioRoot "teams-messages.json") }
Invoke-SeedPhase "SharePoint" { & (Join-Path $engine "Invoke-SeedSharePoint.ps1") -ConfigPath $configPath -SharePointPath (Join-Path $scenarioRoot "sharepoint-sites.json") }
Write-Host "All phases completed"
'@
    Set-Content -LiteralPath (Join-Path $scenarioDir 'run.ps1') -Value $pl7008Run -Encoding UTF8

    # 3. config.json.example
    $exampleObj = [ordered]@{
        _comment        = "Graph permissions: User.ReadWrite.All, Mail.Send, Mail.ReadWrite, Teamwork.Migrate.All, TeamMember.ReadWrite.All, Group.ReadWrite.All, Files.ReadWrite.All, Sites.Manage.All, Sites.ReadWrite.All, Sites.FullControl.All"
        tenantId        = "<your-tenant-id>"
        clientId        = "<your-client-id>"
        clientSecret    = "<your-client-secret>"
        demoUserUpn     = "admin@example.invalid"
        adminUpn        = "admin@example.invalid"
        teamDisplayName = "PL-7008 IT Helpdesk — 2026-08-31"
        targetFolder    = "PL7008-IT-Helpdesk-20260831"
        filesSourceDir  = "DEMO-FILE"
        roles           = [ordered]@{
            Admin       = @{ upn = "admin@example.invalid" }
            ITLead      = @{ upn = "itlead@example.invalid" }
            ITEngineer1 = @{ upn = "itengineer1@example.invalid" }
            ITEngineer2 = @{ upn = "itengineer2@example.invalid" }
            Employee1   = @{ upn = "employee1@example.invalid" }
            Employee2   = @{ upn = "employee2@example.invalid" }
            Employee3   = @{ upn = "employee3@example.invalid" }
        }
    }
    $exampleObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $scenarioDir 'config.json.example') -Encoding UTF8

    # 4. user-profiles.json
    $profilesMap = [ordered]@{
        ITLead      = @{ upn = 'itlead@example.invalid';      jobTitle = '數位轉型專案經理'; department = '金融事業處 / 數位金融部' }
        ITEngineer1 = @{ upn = 'itengineer1@example.invalid'; jobTitle = '資深系統架構師';   department = '金融事業處 / 資訊科技部' }
        ITEngineer2 = @{ upn = 'itengineer2@example.invalid'; jobTitle = '資料工程師';       department = '金融事業處 / 資訊科技部' }
        Employee1   = @{ upn = 'employee1@example.invalid';   jobTitle = '風險管理副理';     department = '金融事業處 / 風險管理部' }
        Employee2   = @{ upn = 'employee2@example.invalid';   jobTitle = '法遵稽核經理';     department = '金融事業處 / 法令遵循部' }
        Employee3   = @{ upn = 'employee3@example.invalid';   jobTitle = '金融事業處 副總';  department = '金融事業處' }
    }
    $userProfilesObj = [ordered]@{
        industries = [ordered]@{
            technology = @{ profiles = $profilesMap }
            finance    = @{ profiles = $profilesMap }
        }
    }
    $userProfilesObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $scenarioDir 'user-profiles.json') -Encoding UTF8

    # 5. DEMO-FILE binaries (3 PDFs + 2 DOCXs >= 1024 bytes) & files-manifest.json
    $fileNames = @(
        'IT-Helpdesk-SOP.pdf'
        'VPN-Troubleshooting-Guide.pdf'
        'Hardware-Loan-Policy.pdf'
        'Onboarding-Checklist.docx'
        'Incident-Escalation-Runbook.docx'
    )
    $padding = 'A' * 1200
    foreach ($fn in $fileNames) {
        $fullPath = Join-Path $demoDir $fn
        if ($fn.EndsWith('.pdf')) {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes("%PDF-1.4`n$padding")
        } else {
            $bytes = [byte[]]@(0x50, 0x4B, 0x03, 0x04) + [System.Text.Encoding]::ASCII.GetBytes($padding)
        }
        [System.IO.File]::WriteAllBytes($fullPath, $bytes)
    }

    $manifestObj = [ordered]@{
        targetFolder = "PL7008-IT-Helpdesk-20260831"
        uploadToRole = "Admin"
        files        = @(
            $fileNames | ForEach-Object {
                [ordered]@{
                    localName   = $_
                    description = "Demo file $_"
                }
            }
        )
    }
    $manifestObj | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $scenarioDir 'files-manifest.json') -Encoding UTF8

    # 6. emails.json (2 threads: 4 messages and 3 messages, referencing TKT-100011 and TKT-100008, Admin CC on all)
    $emailsObj = [ordered]@{
        emailThreads = @(
            [ordered]@{
                threadTopic = "筆電黑屏報修"
                emails      = @(
                    [ordered]@{ order = 1; fromRole = "Employee1";   toRoles = @("ITLead");      ccRoles = @("Admin"); subject = "【報修】我的公司筆電開機後黑屏，急需處理";       body = "<p>我的筆電黑屏，已建單 TKT-100011。</p>"; importance = "high" }
                    [ordered]@{ order = 2; fromRole = "ITLead";      toRoles = @("ITEngineer1"); ccRoles = @("Admin"); subject = "Re: 【報修】我的公司筆電開機後黑屏，急需處理";   body = "<p>請跟進 TKT-100011。</p>";               importance = "normal" }
                    [ordered]@{ order = 3; fromRole = "ITEngineer1"; toRoles = @("Employee1");   ccRoles = @("Admin"); subject = "Re: 【報修】我的公司筆電開機後黑屏，急需處理";   body = "<p>TKT-100011 已安排備機。</p>";           importance = "normal" }
                    [ordered]@{ order = 4; fromRole = "Employee1";   toRoles = @("ITEngineer1"); ccRoles = @("Admin"); subject = "Re: 【報修】我的公司筆電開機後黑屏，急需處理";   body = "<p>收到 TKT-100011 備機，謝謝。</p>";       importance = "normal" }
                )
            }
            [ordered]@{
                threadTopic = "設備借用申請"
                emails      = @(
                    [ordered]@{ order = 1; fromRole = "Employee2";   toRoles = @("ITLead");      ccRoles = @("Admin"); subject = "【設備借用】下週外部會議室需借便攜投影機 + HDMI 線";     body = "<p>申請單號 TKT-100008。</p>";       importance = "normal" }
                    [ordered]@{ order = 2; fromRole = "ITLead";      toRoles = @("ITEngineer2"); ccRoles = @("Admin"); subject = "Re: 【設備借用】下週外部會議室需借便攜投影機 + HDMI 線"; body = "<p>請準備 TKT-100008 設備。</p>";   importance = "normal" }
                    [ordered]@{ order = 3; fromRole = "ITEngineer2"; toRoles = @("Employee2");   ccRoles = @("Admin"); subject = "Re: 【設備借用】下週外部會議室需借便攜投影機 + HDMI 線"; body = "<p>TKT-100008 設備已備妥。</p>";     importance = "normal" }
                )
            }
        )
    }
    $emailsObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $scenarioDir 'emails.json') -Encoding UTF8

    # 7. teams-messages.json (1 channel '工單協作', 6 top-level messages, 9 replies total; msg 0 has 5 replies so DropReplyAt 4 works)
    $teamsObj = [ordered]@{
        teamDisplayName = "PL-7008 IT Helpdesk — 2026-08-31"
        teamDescription = "PL-7008 IT Helpdesk Collaboration"
        channels        = @(
            [ordered]@{
                channelName = "工單協作"
                description = "IT 工單協作頻道"
                messages    = @(
                    [ordered]@{
                        order           = 1
                        fromRole        = "ITLead"
                        createdDateTime = "2026-08-25T09:00:00Z"
                        bodyHtml        = "<p>早安，今日高優工單 TKT-100001 與 TKT-100011 請優先處理。</p>"
                        content         = "<p>早安，今日高優工單 TKT-100001 與 TKT-100011 請優先處理。</p>"
                        mentions        = @("ITEngineer1")
                        replies         = @(
                            [ordered]@{ order = 1; fromRole = "ITEngineer1"; createdDateTime = "2026-08-25T09:05:00Z"; bodyHtml = "<p>收到 TKT-100001。</p>"; content = "<p>收到 TKT-100001。</p>"; mentions = @() }
                            [ordered]@{ order = 2; fromRole = "ITEngineer2"; createdDateTime = "2026-08-25T09:10:00Z"; bodyHtml = "<p>協助排查 TKT-100002。</p>"; content = "<p>協助排查 TKT-100002。</p>"; mentions = @() }
                            [ordered]@{ order = 3; fromRole = "Employee1";   createdDateTime = "2026-08-25T09:15:00Z"; bodyHtml = "<p>補充 TKT-100011 狀況。</p>"; content = "<p>補充 TKT-100011 狀況。</p>"; mentions = @() }
                            [ordered]@{ order = 4; fromRole = "ITEngineer1"; createdDateTime = "2026-08-25T09:20:00Z"; bodyHtml = "<p>TKT-100011 備機已發出。</p>"; content = "<p>TKT-100011 備機已發出。</p>"; mentions = @() }
                            [ordered]@{ order = 5; fromRole = "ITLead";      createdDateTime = "2026-08-25T09:25:00Z"; bodyHtml = "<p>感謝更新 TKT-100011。</p>"; content = "<p>感謝更新 TKT-100011。</p>"; mentions = @() }
                        )
                    }
                    [ordered]@{
                        order           = 2
                        fromRole        = "ITEngineer1"
                        createdDateTime = "2026-08-25T10:00:00Z"
                        bodyHtml        = "<p>VPN 連線工單 TKT-100003 進度更新。</p>"
                        content         = "<p>VPN 連線工單 TKT-100003 進度更新。</p>"
                        mentions        = @()
                        replies         = @(
                            [ordered]@{ order = 1; fromRole = "ITLead"; createdDateTime = "2026-08-25T10:05:00Z"; bodyHtml = "<p>請同步更新知識庫。</p>"; content = "<p>請同步更新知識庫。</p>"; mentions = @() }
                        )
                    }
                    [ordered]@{
                        order           = 3
                        fromRole        = "ITEngineer2"
                        createdDateTime = "2026-08-25T11:00:00Z"
                        bodyHtml        = "<p>信箱容量工單 TKT-100004 已完成。</p>"
                        content         = "<p>信箱容量工單 TKT-100004 已完成。</p>"
                        mentions        = @()
                        replies         = @(
                            [ordered]@{ order = 1; fromRole = "Employee2"; createdDateTime = "2026-08-25T11:05:00Z"; bodyHtml = "<p>已確認 TKT-100004 恢復正常。</p>"; content = "<p>已確認 TKT-100004 恢復正常。</p>"; mentions = @() }
                        )
                    }
                    [ordered]@{
                        order           = 4
                        fromRole        = "Employee2"
                        createdDateTime = "2026-08-25T13:00:00Z"
                        bodyHtml        = "<p>詢問投影機借用工單 TKT-100008。</p>"
                        content         = "<p>詢問投影機借用工單 TKT-100008。</p>"
                        mentions        = @()
                        replies         = @(
                            [ordered]@{ order = 1; fromRole = "ITEngineer2"; createdDateTime = "2026-08-25T13:05:00Z"; bodyHtml = "<p>TKT-100008 已完成備料。</p>"; content = "<p>TKT-100008 已完成備料。</p>"; mentions = @() }
                        )
                    }
                    [ordered]@{
                        order           = 5
                        fromRole        = "Employee3"
                        createdDateTime = "2026-08-25T14:00:00Z"
                        bodyHtml        = "<p>印表機驅動工單 TKT-100005 請協助。</p>"
                        content         = "<p>印表機驅動工單 TKT-100005 請協助。</p>"
                        mentions        = @()
                        replies         = @(
                            [ordered]@{ order = 1; fromRole = "ITEngineer1"; createdDateTime = "2026-08-25T14:05:00Z"; bodyHtml = "<p>TKT-100005 已遠端安裝完成。</p>"; content = "<p>TKT-100005 已遠端安裝完成。</p>"; mentions = @() }
                        )
                    }
                    [ordered]@{
                        order           = 6
                        fromRole        = "ITLead"
                        createdDateTime = "2026-08-25T16:00:00Z"
                        bodyHtml        = "<p>今日工單 TKT-100006 至 TKT-100010 皆已盤點。</p>"
                        content         = "<p>今日工單 TKT-100006 至 TKT-100010 皆已盤點。</p>"
                        mentions        = @()
                        replies         = @()
                    }
                )
            }
        )
    }
    $teamsObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $scenarioDir 'teams-messages.json') -Encoding UTF8

    # 8. sharepoint-sites.json (1 site, 5 documents, 1 list 'IT Tickets' with TKT-100001..TKT-100011)
    $ticketItems = @(
        1..10 | ForEach-Object {
            $id = 'TKT-{0:D6}' -f (100000 + $_)
            [ordered]@{
                Title       = $id
                Subject     = "User's ticket $id issue"
                RequestedBy = "Johanna Lorenz"
                AssignedTo  = "Isaiah Langer"
                Urgency     = "Medium"
                Status      = "In Progress"
            }
        }
    )
    $ticketItems += [ordered]@{
        Title       = "TKT-100011"
        Subject     = "公司筆電開機後黑屏急需處理"
        RequestedBy = "Irvin Sayers"
        AssignedTo  = "Isaiah Langer"
        Urgency     = "High"
        Status      = "In Progress"
    }
    $spObj = [ordered]@{
        _notes = "IT Tickets range: TKT-100001 ~ TKT-100011"
        sites  = @(
            [ordered]@{
                alias       = "pl7008-it-helpdesk-20260831"
                displayName = "PL-7008 — IT Helpdesk — 2026-08-31"
                description = "PL-7008 IT Helpdesk site"
                owners      = @("Admin")
                members     = @("ITLead", "ITEngineer1", "ITEngineer2", "Employee1", "Employee2", "Employee3")
                documents   = @(
                    $fileNames | ForEach-Object {
                        [ordered]@{
                            sourceFilename = $_
                            targetFilename = $_
                        }
                    }
                )
                lists       = @(
                    [ordered]@{
                        displayName = "IT Tickets"
                        description = "IT Helpdesk Tickets"
                        keyField    = "Title"
                        columns     = @(
                            @{ name = "Subject";     type = "text" }
                            @{ name = "RequestedBy"; type = "text" }
                            @{ name = "AssignedTo";  type = "text" }
                            @{ name = "Urgency";     type = "text" }
                            @{ name = "Status";      type = "text" }
                        )
                        items       = $ticketItems
                    }
                )
            }
        )
    }
    $spObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $scenarioDir 'sharepoint-sites.json') -Encoding UTF8

    # 9. Scenario README.md and docs/demo-environment*.md
    $permsText = "Permissions: User.ReadWrite.All, Mail.Send, Mail.ReadWrite, Teamwork.Migrate.All, TeamMember.ReadWrite.All, Group.ReadWrite.All, Files.ReadWrite.All, Sites.Manage.All, Sites.ReadWrite.All, Sites.FullControl.All"
    $scenarioReadme = @"
# PL-7008 IT Helpdesk — 2026-08-31
Target folder: PL7008-IT-Helpdesk-20260831
Site alias: pl7008-it-helpdesk-20260831
Site displayName: PL-7008 — IT Helpdesk — 2026-08-31
IT Tickets: 11 筆工單 (TKT-100001 ~ TKT-100011)
Preflight: Invoke-SeedPreflight.ps1
$permsText
"@
    Set-Content -LiteralPath (Join-Path $scenarioDir 'README.md') -Value $scenarioReadme -Encoding UTF8

    $liveProofEn = @"
# Demo Environment
$permsText
IT Tickets (**11** items)

Live-verified 2026-08-31: one corrective run.ps1 added TKT-100011 (1 added / 10 skipped), and the 196-leaf snapshots taken before and after it differ only in the IT Tickets leaves; the two runs after it both exited 0 and produced no new resources, no duplicates, no list items, no messages and no files, each followed by a read-only verification of 72/72 checks, and a 196-leaf snapshot comparison showed 0 differences. Phase 1 still sends a user-profile PATCH per role account and Phase 4 still re-issues completeMigration plus one POST per team member. The offline suite seed-data/tests/Test-SeedEngine.ps1 passes 459 / 459 passed.
"@
    $liveProofZh = @"
# Demo 環境
$permsText
IT Tickets（**11** 筆）

2026-08-31 實機驗證：修正版 run.ps1 補上 TKT-100011（新增 1 筆／略過 10 筆），前後 196 節點快照僅差該工單；其後兩次執行皆以 0 結束，沒有新增任何資源、沒有重複項目、沒有新增 list item、訊息或檔案，每次皆通過 72/72 唯讀驗證，且兩次 196 節點快照比對為零差異（0 differences）。Phase 1 仍會對每個角色送出 PATCH，Phase 4 仍會重送 completeMigration 與 team member POST。離線測試通過 459 / 459。
"@
    Set-Content -LiteralPath (Join-Path $docsDir 'demo-environment.md') -Value $liveProofEn -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $docsDir 'demo-environment.zh-TW.md') -Value $liveProofZh -Encoding UTF8
    $skillDir = Join-Path $FixtureRoot '.github\skills\m365-demo-data-seeding'
    New-Item -ItemType Directory -Force -Path $skillDir | Out-Null
    Set-Content -LiteralPath (Join-Path $skillDir 'SKILL.md') -Value $liveProofEn -Encoding UTF8
    [System.IO.File]::WriteAllText((Join-Path $FixtureRoot '.gitignore'), "seed-data/scenarios/*/config.json")

    return [pscustomobject]@{
        FixtureRoot = $FixtureRoot
        ScenarioDir = $scenarioDir
        DocsDir     = $docsDir
        MS4022Run   = (Join-Path $ms4022Dir 'run.ps1')
    }
}
