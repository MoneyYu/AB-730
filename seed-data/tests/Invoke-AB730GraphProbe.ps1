#Requires -Version 7.0
<#
.SYNOPSIS
    Self-contained offline stub Graph probes for AB-730 fresh-target-only engine and Invoke-AB730.ps1.
    Invoked from Test-AB730Engine.ps1.
#>

function New-AB730TestJwt {
    param(
        [string]$TenantId = '11111111-2222-3333-4444-555555555555',
        [string]$ClientId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
        [string]$Audience = 'https://graph.microsoft.com',
        [long]$Exp = ([DateTimeOffset]::UtcNow.AddHours(1).ToUnixTimeSeconds()),
        [string[]]$Roles = @(
            'User.Read.All',
            'Group.ReadWrite.All',
            'Files.ReadWrite.All',
            'Sites.Manage.All',
            'Sites.ReadWrite.All',
            'Mail.Send',
            'Mail.ReadWrite',
            'Teamwork.Migrate.All',
            'TeamMember.ReadWrite.All'
        )
    )
    function ConvertTo-Base64Url([string]$Text) {
        $b = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return ([Convert]::ToBase64String($b)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
    $header = ConvertTo-Base64Url '{"alg":"RS256","typ":"JWT"}'
    $payloadObj = [ordered]@{
        tid   = $TenantId
        appid = $ClientId
        aud   = $Audience
        exp   = $Exp
        roles = $Roles
    }
    $payload = ConvertTo-Base64Url ($payloadObj | ConvertTo-Json -Compress)
    return "$header.$payload.stubsignature"
}

function Initialize-AB730ProbeEnvironment {
    param([Parameter(Mandatory)][string]$RootDir)

    if (Test-Path -LiteralPath $RootDir) {
        Remove-Item -LiteralPath $RootDir -Recurse -Force
    }
    $seedRoot = Join-Path $RootDir 'seed-data'
    $engineDest = Join-Path $seedRoot 'engine'
    $scenarioDir = Join-Path $seedRoot 'scenarios\ab730-automotive-20261015'
    $demoDir = Join-Path $scenarioDir 'DEMO-FILE'
    New-Item -ItemType Directory -Force -Path $engineDest | Out-Null
    New-Item -ItemType Directory -Force -Path $demoDir | Out-Null

    Copy-Item (Join-Path $EngineDir '*.ps1') $engineDest -Force
    if (Test-Path (Join-Path $EngineDir 'SharePointSeedHelpers.psm1')) {
        Copy-Item (Join-Path $EngineDir 'SharePointSeedHelpers.psm1') $engineDest -Force
    }
    $templateRun = Join-Path $EngineDir 'Invoke-AB730ScenarioRun.template.ps1'
    if (Test-Path -LiteralPath $templateRun) {
        Copy-Item -LiteralPath $templateRun -Destination (Join-Path $scenarioDir 'run.ps1') -Force
    }

    # Use the repository-owned public fixture, not a private session layout.
    $realScenarioDir = Join-Path (Split-Path -Parent $EngineDir) 'scenarios\ab730-automotive-20261015'
    $canonicalSpecSrc = Join-Path (Split-Path -Parent $EngineDir) 'generator\tests\ab730-data-spec.json'
    $dataSpecPath = Join-Path $RootDir 'ab730-data-spec.json'
    if (-not (Test-Path -LiteralPath $canonicalSpecSrc -PathType Leaf)) {
        throw 'Required public canonical spec fixture is missing.'
    }
    Copy-Item -LiteralPath $canonicalSpecSrc -Destination $dataSpecPath -Force
    Copy-Item -LiteralPath $canonicalSpecSrc -Destination (Join-Path $scenarioDir 'ab730-data-spec.json') -Force

    $officeFiles = @(
        '01-BOM-Cost.xlsx'
        '02-Supplier-Delivery.xlsx'
        '03-Quality-Inspection.xlsx'
        '04-Production-Impact.xlsx'
        '05-Action-Tracker.xlsx'
        '06-Quality-Brief.docx'
        '07-Supplier-Response.docx'
        '08-Improvement-Meeting.docx'
    )
    $spDocs = @(
        '06-Quality-Brief.docx'
        '07-Supplier-Response.docx'
        '08-Improvement-Meeting.docx'
    )
    $pad = 'B' * 1200
    $assetShaMap = [ordered]@{}
    foreach ($f in $officeFiles) {
        $bytes = [byte[]]@(0x50, 0x4B, 0x03, 0x04) + [System.Text.Encoding]::UTF8.GetBytes("$f|$pad")
        $fPath = Join-Path $demoDir $f
        [System.IO.File]::WriteAllBytes($fPath, $bytes)
        $assetShaMap[$f] = (Get-FileHash -LiteralPath $fPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }

    $filesManifestPath = Join-Path $scenarioDir 'files-manifest.json'
    $filesManifest = [ordered]@{
        targetFolder = 'AB730-Automotive-20261015'
        uploadToRole = 'Admin'
        files        = @(
            $officeFiles | ForEach-Object {
                [ordered]@{
                    localName = $_
                    sha256    = $assetShaMap[$_]
                }
            }
        )
        assetSHA256  = $assetShaMap
    }
    $filesManifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $filesManifestPath -Encoding UTF8

    foreach ($scenJson in @('emails.json', 'teams-messages.json', 'sharepoint-sites.json')) {
        $srcJson = Join-Path $realScenarioDir $scenJson
        if (Test-Path -LiteralPath $srcJson -PathType Leaf) {
            Copy-Item -LiteralPath $srcJson -Destination (Join-Path $scenarioDir $scenJson) -Force
        }
    }

    $rolesMap = [ordered]@{
        Admin       = @{ upn = 'admin@contoso.com' }
        Sponsor     = @{ upn = 'sponsor@contoso.com' }
        Planner     = @{ upn = 'planner@contoso.com' }
        Supply      = @{ upn = 'supply@contoso.com' }
        Quality     = @{ upn = 'quality@contoso.com' }
        Operations  = @{ upn = 'operations@contoso.com' }
        Engineering = @{ upn = 'engineering@contoso.com' }
        Analytics   = @{ upn = 'analytics@contoso.com' }
    }
    $configPath = Join-Path $seedRoot 'config.json'
    $configObj = [ordered]@{
        tenantId        = '11111111-2222-3333-4444-555555555555'
        clientId        = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        clientSecret    = 'stub-secret-value-not-real'
        adminUpn        = 'admin@contoso.com'
        demoUserUpn     = 'admin@contoso.com'
        teamDisplayName = '20261015-AB-730'
        targetFolder    = 'AB730-Automotive-20261015'
        filesSourceDir  = 'scenarios\ab730-automotive-20261015\DEMO-FILE'
        roles           = $rolesMap
    }
    $configObj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $configPath -Encoding UTF8

    $artifactHashes = [ordered]@{}
    foreach ($jsonRel in @('files-manifest.json', 'emails.json', 'teams-messages.json', 'sharepoint-sites.json')) {
        $artifactHashes[$jsonRel] = (Get-FileHash -LiteralPath (Join-Path $scenarioDir $jsonRel) -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    foreach ($f in $officeFiles) {
        $artifactHashes[$f] = $assetShaMap[$f]
    }

    $codeHashes = [ordered]@{}
    $engineScripts = @(
        'Invoke-AB730.ps1',
        'Invoke-SeedPreflight.ps1',
        'Invoke-UploadFiles.ps1',
        'Invoke-SeedEmails.ps1',
        'Invoke-SeedTeamsChannel.ps1',
        'Invoke-SeedSharePoint.ps1',
        'Seed-Idempotency.ps1',
        'Seed-GraphRead.ps1',
        'Seed-OneDriveOwnership.ps1'
    )
    foreach ($es in $engineScripts) {
        $esPath = Join-Path $engineDest $es
        $codeHashes["seed-data/engine/$es"] = (Get-FileHash -LiteralPath $esPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $runScriptPath = Join-Path $scenarioDir 'run.ps1'
    if (Test-Path -LiteralPath $runScriptPath -PathType Leaf) {
        $codeHashes['seed-data/scenarios/ab730-automotive-20261015/run.ps1'] = (Get-FileHash -LiteralPath $runScriptPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }

    $dataSpecHash = if (Test-Path -LiteralPath $dataSpecPath -PathType Leaf) {
        (Get-FileHash -LiteralPath $dataSpecPath -Algorithm SHA256).Hash.ToLowerInvariant()
    } else { '' }

    $manifestPath = Join-Path $RootDir 'ab730-deployment-manifest.local.json'
    $manifestObj = [ordered]@{
        version              = 1
        state                = 'ApprovedForLiveExecution'
        credentialConfigPath = $configPath
        dataSpecPath         = $dataSpecPath
        dataSpecSha256       = $dataSpecHash
        artifactHashes       = $artifactHashes
        codeHashes           = $codeHashes
        target               = [ordered]@{
            tenantId = '11111111-2222-3333-4444-555555555555'
            clientId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        }
        roles                = $rolesMap
        locale               = 'zh-CN'
        currency             = 'CNY'
        resources            = [ordered]@{
            oneDrive   = [ordered]@{
                ownerRole                     = 'Admin'
                folder                        = 'AB730-Automotive-20261015'
                officeFileCount               = 8
                engineProofMetadataFileCount  = 1
                existingItemOverwriteAllowed  = $false
            }
            teams      = [ordered]@{
                displayName          = '20261015-AB-730'
                ownerRoles           = @('Admin')
                memberRoles          = @('Sponsor','Planner','Supply','Quality','Operations','Engineering','Analytics')
                businessChannelCount = 5
                defaultChannelCount  = 1
                storyMessageCount    = 15
            }
            sharePoint = [ordered]@{
                alias           = 'ab730-automotive-20261015'
                displayName     = 'AB-730 - Automotive Knowledge - 20261015'
                visibility      = 'Private'
                ownerRoles      = @('Admin')
                memberRoles     = @('Sponsor','Planner','Supply','Quality','Operations','Engineering','Analytics')
                documentLibrary = 'DemoKnowledge'
                documentCount   = 3
                customListCount = 0
            }
            outlook    = [ordered]@{
                threadCount                               = 5
                sendAndReplyCount                         = 15
                ccEveryMessage                            = @('Admin')
                newlySentMessageLookupInDeclaredRoleInboxes = $true
                unrelatedBodyOrAttachmentReadAllowed      = $false
                externalRecipientsAllowed                 = $false
            }
        }
        writeMode                                         = 'additive-fresh-targets-only'
        preflightMustFindNewTargets                       = $true
        newAccountOrProfileChangeAllowed                  = $false
        newAppConsentOrSubscriptionAllowed                = $false
        oldScenarioRerunAllowed                           = $false
        deletionRenameAdoptionRepairAllowed               = $false
        automaticWriteRetryAllowed                        = $false
        partialCreationCanRemainOnFailure                 = $true
        requiresFreshAccountPermissionAndCollisionChecks  = $true
        requiresOfflineValidationAndCollisionChecks       = $true
        requiresOfflineValidationAndRiskReview            = $true
        requiresExplicitUserPlanApproval                  = $true
        tenantWriteApproved                               = $true
    }
    $manifestObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    return [pscustomobject]@{
        RootDir      = $RootDir
        SeedRoot     = $seedRoot
        EngineDir    = $engineDest
        ScenarioDir  = $scenarioDir
        DemoDir      = $demoDir
        ConfigPath   = $configPath
        ManifestPath = $manifestPath
        DataSpecPath = $dataSpecPath
        OfficeFiles  = $officeFiles
        SpDocs       = $spDocs
    }
}

function Invoke-AB730Probe {
    param(
        [Parameter(Mandatory)][ValidateSet('Preflight','Upload','Emails','Teams','SharePoint','RunnerPreflight','RunnerLive')][string]$Target,
        [Parameter(Mandatory)][string]$Mode,
        [hashtable]$Overrides = @{}
    )

    $probeRoot = Join-Path $TestsDir '.tmp-ab730-probe'
    $envInfo = Initialize-AB730ProbeEnvironment -RootDir $probeRoot
    $requests = [System.Collections.Generic.List[object]]::new()
    $tokenRequests = [System.Collections.Generic.List[object]]::new()
    $outputLines = [System.Collections.Generic.List[string]]::new()
    $threw = $false
    $errorMsg = ''

    # Apply any file/config/manifest overrides before running
    if ($Overrides.ContainsKey('ModifyConfig')) {
        $cfg = Get-Content -LiteralPath $envInfo.ConfigPath -Raw | ConvertFrom-Json
        & $Overrides['ModifyConfig'] $cfg $envInfo
        $cfg | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $envInfo.ConfigPath -Encoding UTF8
    }
    if ($Overrides.ContainsKey('ModifyTeams')) {
        $tmPath = Join-Path $envInfo.ScenarioDir 'teams-messages.json'
        $tm = Get-Content -LiteralPath $tmPath -Raw | ConvertFrom-Json
        & $Overrides['ModifyTeams'] $tm $envInfo
        $tm | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $tmPath -Encoding UTF8
    }
    if ($Overrides.ContainsKey('ModifyEmails')) {
        $emPath = Join-Path $envInfo.ScenarioDir 'emails.json'
        $em = Get-Content -LiteralPath $emPath -Raw | ConvertFrom-Json
        & $Overrides['ModifyEmails'] $em $envInfo
        $em | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $emPath -Encoding UTF8
    }
    if ($Overrides.ContainsKey('ModifySharePoint')) {
        $spPath = Join-Path $envInfo.ScenarioDir 'sharepoint-sites.json'
        $sp = Get-Content -LiteralPath $spPath -Raw | ConvertFrom-Json
        & $Overrides['ModifySharePoint'] $sp $envInfo
        $sp | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $spPath -Encoding UTF8
    }
    if ($Overrides.ContainsKey('TamperDemoFile')) {
        $targetFile = Join-Path $envInfo.DemoDir $envInfo.OfficeFiles[0]
        [System.IO.File]::AppendAllText($targetFile, 'tampered')
    }
    if ($Overrides.ContainsKey('CustomSetup')) {
        & $Overrides['CustomSetup'] $envInfo
    }
    if ($Overrides.ContainsKey('ModifyManifest')) {
        $mf = Get-Content -LiteralPath $envInfo.ManifestPath -Raw | ConvertFrom-Json
        & $Overrides['ModifyManifest'] $mf $envInfo
        $mf | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $envInfo.ManifestPath -Encoding UTF8
    }
    if ($Overrides.ContainsKey('PrecreateReceipt')) {
        $receiptDir = Join-Path $envInfo.ScenarioDir '.seed-Execution'
        New-Item -ItemType Directory -Force -Path $receiptDir | Out-Null
    }

    $jwt = if ($Overrides.ContainsKey('Jwt')) { $Overrides['Jwt'] } else { New-AB730TestJwt }
    $gitTracked = [bool]($Overrides['GitTracked'])
    $gitNotIgnored = [bool]($Overrides['GitNotIgnored'])

    $hasPrevToken = Test-Path Variable:\global:AccessToken
    $prevToken = if ($hasPrevToken) { $global:AccessToken } else { $null }
    $prevIrm = Get-Item Function:\Invoke-RestMethod -ErrorAction SilentlyContinue
    $prevIwr = Get-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
    $prevWh  = Get-Item Function:\Write-Host -ErrorAction SilentlyContinue
    $prevGit = Get-Item Function:\git -ErrorAction SilentlyContinue
    $prevSleep = Get-Item Function:\Start-Sleep -ErrorAction SilentlyContinue
    $prevSetMgCtx = Get-Item Function:\Set-MgRequestContext -ErrorAction SilentlyContinue
    $prevGetMgCtx = Get-Item Function:\Get-MgRequestContext -ErrorAction SilentlyContinue

    # Stateful stub state
    $state = @{
        GroupQueryCount          = 0
        FolderQueryCount         = 0
        InboxQueryCount          = 0
        TeamCreated              = $false
        SiteGroupCreated         = $false
        DemoLibCreated           = $false
        OneDriveFolderId         = $null
        StrictDriveItems         = [System.Collections.Generic.List[object]]::new()
        StrictProof              = $null
        CreatedChannels          = [System.Collections.Generic.List[object]]::new()
        SentMailSeq              = 0
        LastSentInternetId       = ''
        LastThreadConversationId = ''
        AdminInboxMessages       = [System.Collections.Generic.List[object]]::new()
        RoleInboxes              = @{}
        UploadedSpDocs           = [System.Collections.Generic.List[object]]::new()
        AddedTeamMembers         = [System.Collections.Generic.List[object]]::new()
        MgMaxRetry               = 3
        MgMaxRetrySetCalls       = 0
        MgMaxRetryAtFirstWrite   = $null
        MailRead429Injected      = $false
        SleepDelays              = [System.Collections.Generic.List[int]]::new()
    }

    $errorStatusCode = $null
    $caughtEr = $null
    try {
        Set-StrictMode -Off
        $global:AccessToken = $jwt
        Set-Item Function:\global:Start-Sleep -Value {
            param($Seconds, $Milliseconds)
            if ($null -ne $Seconds) { $state.SleepDelays.Add([int]$Seconds) }
        }.GetNewClosure()
        Set-Item Function:\global:Set-MgRequestContext -Value {
            param([int]$MaxRetry, [int]$RetryDelay)
            $state.MgMaxRetrySetCalls++
            if ($Mode -ne 'MgContextEnforceFail' -and $PSBoundParameters.ContainsKey('MaxRetry')) {
                $state.MgMaxRetry = $MaxRetry
            }
            return [pscustomobject]@{ MaxRetry = $state.MgMaxRetry; RetryDelay = 3 }
        }.GetNewClosure()
        Set-Item Function:\global:Get-MgRequestContext -Value {
            return [pscustomobject]@{ MaxRetry = $state.MgMaxRetry; RetryDelay = 3 }
        }.GetNewClosure()
        Set-Item Function:\global:Write-Host -Value {
            param([Parameter(ValueFromRemainingArguments)]$RemainingArgs)
            $text = ($RemainingArgs | Where-Object { $_ -isnot [System.ConsoleColor] }) -join ' '
            $outputLines.Add($text)
        }.GetNewClosure()

        Set-Item Function:\global:git -Value {
            param([Parameter(ValueFromRemainingArguments)]$GitArgs)
            $joined = $GitArgs -join ' '
            if ($joined -match 'ls-files') {
                if ($gitTracked) {
                    $global:LASTEXITCODE = 0
                    return 'seed-data/config.json'
                }
                $global:LASTEXITCODE = 1
                return ''
            }
            if ($joined -match 'check-ignore') {
                if ($gitNotIgnored) {
                    $global:LASTEXITCODE = 1
                    return ''
                }
                $global:LASTEXITCODE = 0
                return 'seed-data/config.json'
            }
            $global:LASTEXITCODE = 0
            return ''
        }.GetNewClosure()

        $graphHandler = {
            param([string]$Method, [string]$Uri, $Headers, $Body, [string]$ContentType, $InFile, [string]$ResponseHeadersVariable, [string]$StatusCodeVariable)
            if ($Uri -match '^https://login\.microsoftonline\.com/') {
                $tokenRequests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
                if ($Mode -eq 'OAuthErrorLeakTest') {
                    $oauthReq = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $Uri)
                    $oauthReq.Content = [System.Net.Http.StringContent]::new('grant_type=client_credentials&client_id=00000000-0000-0000-0000-000000000002&client_secret=stub-secret-value-not-real')
                    $oauthResp = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::Unauthorized)
                    $oauthResp.RequestMessage = $oauthReq
                    $oauthResp.Content = [System.Net.Http.StringContent]::new('{"error":"invalid_client","error_description":"AADSTS7000215 tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}')
                    $ex = [Microsoft.PowerShell.Commands.HttpResponseException]::new(
                        'Response status code does not indicate success: 401 (Unauthorized) Authorization: ****** client_secret=stub-secret-value-not-real',
                        $oauthResp
                    )
                    $ex.Data['RawSecret'] = 'stub-secret-value-not-real'
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Unauthorized', [System.Management.Automation.ErrorCategory]::AuthenticationError, $oauthResp)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":"invalid_client","error_description":"AADSTS7000215 tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}')
                    throw $er
                }
                return [pscustomobject]@{ access_token = $jwt }
            }
            if ($Method.ToUpperInvariant() -ne 'GET' -and $null -eq $state.MgMaxRetryAtFirstWrite) {
                $state.MgMaxRetryAtFirstWrite = $state.MgMaxRetry
            }
            $requests.Add([pscustomobject]@{
                Method = $Method.ToUpperInvariant()
                Uri    = $Uri
                Body   = $Body
            })
            $decUri = [System.Uri]::UnescapeDataString($Uri)

            # 1. User existence checks
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/users/([^/?]+)(\?|$)') {
                $userSeg = $Matches[1]
                if ($decUri -notmatch '/drive' -and $decUri -notmatch '/mailFolders' -and $decUri -notmatch '/messages') {
                    return [pscustomobject]@{
                        id                = "uid-$userSeg"
                        displayName       = $userSeg
                        userPrincipalName = $userSeg
                        accountEnabled    = $true
                    }
                }
            }

            # 2. OneDrive root & folder checks
            if ($Method -eq 'GET' -and $decUri -match '/drive/root\?') {
                return [pscustomobject]@{ id = 'root-1'; webUrl = 'https://contoso-my.sharepoint.com/personal/admin/Documents' }
            }
            if ($Method -eq 'GET' -and $decUri -match '/drive/root:/AB730-Automotive-20261015(\?|$)') {
                $state.FolderQueryCount++
                if ($Mode -eq 'OneDrive503WithItemNotFoundText') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 503 (Service Unavailable).')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'ServiceUnavailable', [System.Management.Automation.ErrorCategory]::ConnectionError, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"serviceNotAvailable","message":"Transient upstream proxy cache miss (itemNotFound in backend replica)."}}')
                    throw $er
                }
                if ($Mode -eq 'ExistingOneDriveFolder' -or ($Mode -eq 'OwningRaceOneDrive' -and $state.FolderQueryCount -ge 1)) {
                    return [pscustomobject]@{
                        id     = 'fld-existing'
                        name   = 'AB730-Automotive-20261015'
                        eTag   = 'etag-existing'
                        folder = [pscustomobject]@{ childCount = 1 }
                    }
                }
                if ($null -ne $state.OneDriveFolderId) {
                    return [pscustomobject]@{
                        id     = $state.OneDriveFolderId
                        name   = 'AB730-Automotive-20261015'
                        eTag   = 'etag-created'
                        folder = [pscustomobject]@{ childCount = 0 }
                    }
                }
                $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 404 (Not Found).')
                $er = [System.Management.Automation.ErrorRecord]::new($ex, 'NotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $null)
                $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"itemNotFound","message":"The resource could not be found."}}')
                throw $er
            }
            if ($Method -eq 'GET' -and $decUri -match '/drive/items/[^/]+/children') {
                return [pscustomobject]@{ value = @($state.StrictDriveItems) }
            }
            if ($Method -eq 'GET' -and $decUri -match '/drive/items/strict-[^/]+/content$') {
                return $state.StrictProof
            }
            if ($Method -eq 'POST' -and $decUri -match '/drive/root/children$') {
                if ($Mode -eq 'UploadErrorLeakTest') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 409 (Conflict) Authorization: Bearer eyJSECRET_JWT.payload.sig')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Conflict', [System.Management.Automation.ErrorCategory]::ResourceExists, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"nameAlreadyExists","message":"Folder exists tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}}')
                    throw $er
                }
                if ($Mode -eq 'PostCreateRaceOneDrive') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 409 (Conflict).')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Conflict', [System.Management.Automation.ErrorCategory]::ResourceExists, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"nameAlreadyExists","message":"Name already exists"}}')
                    throw $er
                }
                $state.OneDriveFolderId = 'fld-new-1'
                return [pscustomobject]@{ id = 'fld-new-1'; name = 'AB730-Automotive-20261015'; eTag = 'etag-new-1'; folder = [pscustomobject]@{ childCount = 0 } }
            }
            if ($Method -eq 'POST' -and $decUri -match '/drive/items/[^:]+:/([^:]+):/createUploadSession$') {
                $fn = $Matches[1]
                return [pscustomobject]@{ uploadUrl = "https://stub-upload.invalid/$([System.Uri]::EscapeDataString($fn))" }
            }
            if ($Method -eq 'PUT' -and $Uri -match '^https://stub-upload\.invalid/(.+)$') {
                $fn = [System.Uri]::UnescapeDataString($Matches[1])
                $len = if ($Body -is [byte[]]) { $Body.Length } else { 1250 }
                $idx = $state.StrictDriveItems.Count + 1
                $state.StrictDriveItems.Add([pscustomobject]@{
                    id   = "strict-$idx"
                    name = $fn
                    eTag = "etag-$idx"
                    size = $len
                    file = [pscustomobject]@{}
                })
                if ($fn -eq '.ms4018-seed-proof.json' -and $Body -is [byte[]]) {
                    $state.StrictProof = ([System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json)
                }
                return [pscustomobject]@{
                    id                             = "strict-$idx"
                    name                           = $fn
                    size                           = $len
                    eTag                           = "etag-$idx"
                    '@microsoft.graph.downloadUrl' = 'https://contoso.sharepoint.com/download?tempauth=STUBSECRET_TOKEN'
                }
            }
            if ($Method -eq 'PUT' -and $decUri -match '/drive/root:/AB730-Automotive-20261015/([^:?]+):/content') {
                $fn = $Matches[1]
                return [pscustomobject]@{
                    id   = "file-$fn"
                    name = $fn
                    size = 1250
                    eTag = "etag-$fn"
                    '@microsoft.graph.downloadUrl' = 'https://contoso.sharepoint.com/download?tempauth=STUBSECRET_TOKEN'
                }
            }

            # 3. Groups queries (Teams displayName, SharePoint mailNickname, SharePoint displayName, owners, members)
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/groups/([^/]+)/owners') {
                return [pscustomobject]@{
                    value = @([pscustomobject]@{ id = 'uid-admin@contoso.com'; userPrincipalName = 'admin@contoso.com' })
                }
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/groups/([^/]+)/members') {
                $grpId = $Matches[1]
                if ($Mode -eq 'SharePointVerifyMismatch') {
                    return [pscustomobject]@{
                        value = @([pscustomobject]@{ id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' })
                    }
                }
                if ($Mode -eq 'SharePointExtraMember' -and $grpId -notmatch '^team-') {
                    return [pscustomobject]@{
                        value = @(
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-planner@contoso.com'; userPrincipalName = 'planner@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-supply@contoso.com'; userPrincipalName = 'supply@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-quality@contoso.com'; userPrincipalName = 'quality@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-operations@contoso.com'; userPrincipalName = 'operations@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-engineering@contoso.com'; userPrincipalName = 'engineering@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-analytics@contoso.com'; userPrincipalName = 'analytics@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-extra@contoso.com'; userPrincipalName = 'extra@contoso.com' }
                        )
                    }
                }
                if ($grpId -match '^team-') {
                    if ($Mode -eq 'TeamsBackingGroupForeignMember') {
                        return [pscustomobject]@{
                            value = @(
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-admin@contoso.com'; userPrincipalName = 'admin@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-planner@contoso.com'; userPrincipalName = 'planner@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-supply@contoso.com'; userPrincipalName = 'supply@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-quality@contoso.com'; userPrincipalName = 'quality@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-operations@contoso.com'; userPrincipalName = 'operations@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-engineering@contoso.com'; userPrincipalName = 'engineering@contoso.com' }
                                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-foreign@contoso.com'; userPrincipalName = 'foreign@contoso.com' }
                            )
                        }
                    }
                    return [pscustomobject]@{
                        value = @(
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-admin@contoso.com'; userPrincipalName = 'admin@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-planner@contoso.com'; userPrincipalName = 'planner@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-supply@contoso.com'; userPrincipalName = 'supply@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-quality@contoso.com'; userPrincipalName = 'quality@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-operations@contoso.com'; userPrincipalName = 'operations@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-engineering@contoso.com'; userPrincipalName = 'engineering@contoso.com' }
                            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-analytics@contoso.com'; userPrincipalName = 'analytics@contoso.com' }
                        )
                    }
                }
                return [pscustomobject]@{
                    value = @(
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-planner@contoso.com'; userPrincipalName = 'planner@contoso.com' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-supply@contoso.com'; userPrincipalName = 'supply@contoso.com' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-quality@contoso.com'; userPrincipalName = 'quality@contoso.com' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-operations@contoso.com'; userPrincipalName = 'operations@contoso.com' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-engineering@contoso.com'; userPrincipalName = 'engineering@contoso.com' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-analytics@contoso.com'; userPrincipalName = 'analytics@contoso.com' }
                    )
                }
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/groups\?') {
                if ($Mode -eq 'PreflightErrorLeakTest') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 403 (Forbidden) Authorization: Bearer eyJSECRET_JWT.payload.sig')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Forbidden', [System.Management.Automation.ErrorCategory]::PermissionDenied, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"Authorization_RequestDenied","message":"Denied tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}}')
                    throw $er
                }
                $state.GroupQueryCount++
                $isPage2 = ($decUri -match 'skiptoken=page2')

                if ($decUri -match "displayName eq '20261015-AB-730'") {
                    if ($Mode -eq 'OwningRaceNonTeamGroup') {
                        if (-not $isPage2) {
                            return [pscustomobject]@{
                                value             = @()
                                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups?$filter=displayName%20eq%20''20261015-AB-730''&$select=id,displayName,resourceProvisioningOptions&$skiptoken=page2'
                            }
                        }
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'g-nonteam-collide'; displayName = '20261015-AB-730'; mailNickname = '20261015-ab-730'; resourceProvisioningOptions = @() })
                        }
                    }
                    if ($Mode -eq 'Page2TeamCollision' -or ($Mode -eq 'OwningRaceTeams' -and $Target -eq 'Teams')) {
                        if (-not $isPage2) {
                            return [pscustomobject]@{
                                value             = @([pscustomobject]@{ id = 'g-other'; displayName = 'Other'; resourceProvisioningOptions = @('Team') })
                                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups?$filter=displayName%20eq%20''20261015-AB-730''&$select=id,displayName,mailNickname,resourceProvisioningOptions&$skiptoken=page2'
                            }
                        }
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'g-team-collide'; displayName = '20261015-AB-730'; mailNickname = '20261015-ab-730'; resourceProvisioningOptions = @('Team') })
                        }
                    }
                    if ($state.TeamCreated) {
                        if ($Mode -eq 'Teams202EmptyBodyWithForeignRace') {
                            # Simulate a foreign team with the same displayName appearing in group search after POST /teams
                            return [pscustomobject]@{
                                value = @(
                                    [pscustomobject]@{ id = 'team-foreign-race'; displayName = '20261015-AB-730'; resourceProvisioningOptions = @('Team') }
                                )
                            }
                        }
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'team-new-1'; displayName = '20261015-AB-730'; resourceProvisioningOptions = @('Team') })
                        }
                    }
                    return [pscustomobject]@{ value = @() }
                }

                if ($decUri -match "mailNickname eq 'ab730-automotive-20261015'") {
                    if ($Mode -eq 'Page2SiteAliasCollision' -or ($Mode -eq 'OwningRaceSiteAlias' -and $Target -eq 'SharePoint')) {
                        if (-not $isPage2) {
                            return [pscustomobject]@{
                                value             = @()
                                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups?$filter=mailNickname%20eq%20''ab730-automotive-20261015''&$select=id,displayName,mailNickname,groupTypes&$skiptoken=page2'
                            }
                        }
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'g-site-collide'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; mailNickname = 'ab730-automotive-20261015'; groupTypes = @('Unified') })
                        }
                    }
                    if ($state.SiteGroupCreated) {
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'g-site-new-1'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; mailNickname = 'ab730-automotive-20261015'; groupTypes = @('Unified') })
                        }
                    }
                    return [pscustomobject]@{ value = @() }
                }

                if ($decUri -match "displayName eq 'AB-730 - Automotive Knowledge - 20261015'") {
                    if ($Mode -eq 'Page2SiteDisplayNameCollision' -or ($Mode -eq 'OwningRaceSiteDisplayName' -and $Target -eq 'SharePoint')) {
                        if (-not $isPage2) {
                            return [pscustomobject]@{
                                value             = @()
                                '@odata.nextLink' = "$Uri&`$skiptoken=page2"
                            }
                        }
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'g-sitedisp-collide'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; mailNickname = 'other-alias'; groupTypes = @('Unified') })
                        }
                    }
                    if ($state.SiteGroupCreated) {
                        return [pscustomobject]@{
                            value = @([pscustomobject]@{ id = 'g-site-new-1'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; mailNickname = 'ab730-automotive-20261015'; groupTypes = @('Unified') })
                        }
                    }
                    return [pscustomobject]@{ value = @() }
                }

                return [pscustomobject]@{ value = @() }
            }

            # 4. Mail Inbox queries and send/reply
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/users/([^/]+)/mailFolders(?:/inbox|\(''inbox''\))/messages') {
                $mailboxUpn = $Matches[1]
                $state.InboxQueryCount++
                $isPage2 = ($decUri -match '\$skip=100')

                if ($mailboxUpn -ieq 'admin@contoso.com') {
                    if ($Mode -eq 'Page2MailCollision' -or ($Mode -eq 'OwningRaceEmails' -and $Target -eq 'Emails')) {
                        if (-not $isPage2) {
                            return [pscustomobject]@{
                                value             = @([pscustomobject]@{ id = 'm-unrelated'; subject = 'Unrelated'; conversationId = 'c-0'; receivedDateTime = '2026-09-01T00:00:00Z' })
                                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders(''inbox'')/messages?$select=id,subject,conversationId,receivedDateTime&$top=100&$skip=100'
                            }
                        }
                        return [pscustomobject]@{
                            value = @(
                                [pscustomobject]@{ id = 'm-collide-1'; subject = '【AB-730 20261015 Automotive】供应商交期与缺料风险'; conversationId = 'c-1'; receivedDateTime = '2026-09-10T00:00:00Z' }
                            )
                        }
                    }
                    return [pscustomobject]@{ value = @($state.AdminInboxMessages) }
                }

                # Sender/recipient Inbox lookup for newly sent message
                if ($Mode -eq 'MailReplyLookupFails') {
                    return [pscustomobject]@{ value = @() }
                }
                if ($decUri -match "internetMessageId eq '([^']+)'") {
                    $wantedImid = $Matches[1]
                    if ($wantedImid -eq $state.LastSentInternetId) {
                        return [pscustomobject]@{
                            value = @(
                                [pscustomobject]@{
                                    id                = "msg-$($mailboxUpn)-$($state.SentMailSeq)"
                                    subject           = 'Matched by internetMessageId'
                                    conversationId    = "conv-$($state.SentMailSeq)"
                                    internetMessageId = $wantedImid
                                    receivedDateTime  = [DateTimeOffset]::UtcNow.ToString('o')
                                }
                            )
                        }
                    }
                    return [pscustomobject]@{ value = @() }
                }

                # Sender's own Inbox snapshot before/after sendMail to capture internetMessageId
                if ($Mode -eq 'MailRead429RetryAfter' -and $state.SentMailSeq -eq 1 -and -not $state.MailRead429Injected) {
                    $state.MailRead429Injected = $true
                    $resp429 = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::TooManyRequests)
                    $resp429.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([TimeSpan]::FromSeconds(7))
                    $resp429.Content = [System.Net.Http.StringContent]::new('{"error":{"code":"activityLimitReached","message":"Throttled tempauth=SECRET_CAP_TOKEN"}}')
                    $ex429 = [Microsoft.PowerShell.Commands.HttpResponseException]::new('Response status code does not indicate success: 429 (Too Many Requests).', $resp429)
                    $er429 = [System.Management.Automation.ErrorRecord]::new($ex429, 'TooManyRequests', [System.Management.Automation.ErrorCategory]::LimitsExceeded, $resp429)
                    throw $er429
                }
                if ($state.RoleInboxes.ContainsKey($mailboxUpn)) {
                    return [pscustomobject]@{ value = @($state.RoleInboxes[$mailboxUpn]) }
                }
                return [pscustomobject]@{ value = @() }
            }

            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/users/([^/]+)/sendMail$') {
                if ($Mode -eq 'EmailsErrorLeakTest') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 403 (Forbidden) Authorization: Bearer eyJSECRET_JWT.payload.sig')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Forbidden', [System.Management.Automation.ErrorCategory]::PermissionDenied, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"ErrorAccessDenied","message":"SendMail denied tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}}')
                    throw $er
                }
                $senderUpn = $Matches[1]
                $state.SentMailSeq++
                $imid = "<ab730-msg-$($state.SentMailSeq)@contoso.com>"
                $convId = "conv-$($state.SentMailSeq)"
                $state.LastSentInternetId = $imid
                $state.LastThreadConversationId = $convId
                $bodyText = if ($Body -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($Body) } else { "$Body" }
                $bodyObj = if ($bodyText) { $bodyText | ConvertFrom-Json } else { $null }
                $subj = if ($bodyObj) { "$($bodyObj.message.subject)" } else { 'Subject' }
                $msgObj = [pscustomobject]@{
                    id                = "admin-msg-$($state.SentMailSeq)"
                    subject           = $subj
                    conversationId    = $convId
                    internetMessageId = $imid
                    receivedDateTime  = [DateTimeOffset]::UtcNow.ToString('o')
                }
                $state.AdminInboxMessages.Add($msgObj)
                foreach ($roleUpn in @('admin@contoso.com','sponsor@contoso.com','planner@contoso.com','supply@contoso.com','quality@contoso.com','operations@contoso.com','engineering@contoso.com','analytics@contoso.com')) {
                    $existingInb = if ($state.RoleInboxes.ContainsKey($roleUpn)) { @($state.RoleInboxes[$roleUpn]) } else { @() }
                    $state.RoleInboxes[$roleUpn] = @($msgObj) + $existingInb
                }
                return $null
            }

            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/users/([^/]+)/messages/([^/]+)/reply$') {
                $senderUpn = $Matches[1]
                $state.SentMailSeq++
                $imid = "<ab730-msg-$($state.SentMailSeq)@contoso.com>"
                $state.LastSentInternetId = $imid
                $convId = if ($state.LastThreadConversationId) { $state.LastThreadConversationId } else { 'conv-1' }
                $msgObj = [pscustomobject]@{
                    id                = "admin-reply-$($state.SentMailSeq)"
                    subject           = 'Re: thread'
                    conversationId    = $convId
                    internetMessageId = $imid
                    receivedDateTime  = [DateTimeOffset]::UtcNow.ToString('o')
                }
                $state.AdminInboxMessages.Add($msgObj)
                foreach ($roleUpn in @('admin@contoso.com','sponsor@contoso.com','planner@contoso.com','supply@contoso.com','quality@contoso.com','operations@contoso.com','engineering@contoso.com','analytics@contoso.com')) {
                    $existingInb = if ($state.RoleInboxes.ContainsKey($roleUpn)) { @($state.RoleInboxes[$roleUpn]) } else { @() }
                    $state.RoleInboxes[$roleUpn] = @($msgObj) + $existingInb
                }
                return $null
            }

            # 5. Teams creation, operations, channels, messages, completeMigration, members
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/teams$') {
                if ($Mode -eq 'TeamsErrorLeakTest') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 403 (Forbidden) Authorization: Bearer eyJSECRET_JWT.payload.sig.')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Forbidden', [System.Management.Automation.ErrorCategory]::PermissionDenied, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}}')
                    throw $er
                }
                $state.TeamCreated = $true
                $state.CreatedChannels.Add([pscustomobject]@{ id = 'ch-general'; displayName = 'General' })
                if ($Mode -in @('Teams202EmptyBodyWithForeignRace', 'TeamsOperationAnchorMismatch', 'TeamsGetIdAnchorMismatch', 'TeamsGetBlankDisplayName', 'TeamsGetMissingDisplayName')) {
                    $global:__AB730LastHeaders = @{
                        Location           = @("/teams('team-202-hdr')/operations('op-202-hdr')")
                        'Content-Location' = @("/teams('team-202-hdr')")
                    }
                    $global:__AB730LastStatus = 202
                    return $null
                }
                return [pscustomobject]@{ id = 'team-new-1' }
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/teams(?:\(''([^'']+)''\)|/([^/]+))/operations(?:\(''([^'']+)''\)|/([^/?]+))') {
                $tId = if ($Matches[1]) { $Matches[1] } else { $Matches[2] }
                $opId = if ($Matches[3]) { $Matches[3] } else { $Matches[4] }
                $opTargetId = if ($Mode -eq 'TeamsOperationAnchorMismatch') { 'team-mutated-op' } else { $tId }
                return [pscustomobject]@{
                    id                     = $opId
                    operationType          = 'createTeam'
                    status                 = 'succeeded'
                    targetResourceId       = $opTargetId
                    targetResourceLocation = "/teams('$opTargetId')"
                }
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/teams/([^/]+)/channels(\?|$)') {
                return [pscustomobject]@{ value = @($state.CreatedChannels) }
            }
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/teams/([^/]+)/channels$') {
                $b = if ($Body -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json } elseif ($Body -is [string]) { $Body | ConvertFrom-Json } else { $Body }
                $chId = "ch-$($state.CreatedChannels.Count + 1)"
                $chObj = [pscustomobject]@{ id = $chId; displayName = $b.displayName }
                $state.CreatedChannels.Add($chObj)
                return $chObj
            }
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/teams/([^/]+)/channels/([^/]+)/messages$') {
                return [pscustomobject]@{ id = "tmsg-$($requests.Count)" }
            }
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/teams/([^/]+)/channels/([^/]+)/messages/([^/]+)/replies$') {
                return [pscustomobject]@{ id = "treply-$($requests.Count)" }
            }
            if ($Method -eq 'POST' -and $decUri -match '/completeMigration$') {
                return $null
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/teams(?:\(''([^'']+)''\)|/([^/?]+))(\?|$)') {
                $tId = if ($Matches[1]) { $Matches[1] } else { $Matches[2] }
                $returnedTeamId = if ($Mode -eq 'TeamsGetIdAnchorMismatch') { 'team-mutated-get' } else { $tId }
                if ($Mode -eq 'TeamsGetBlankDisplayName') {
                    return [pscustomobject]@{ id = $returnedTeamId; displayName = '   ' }
                }
                if ($Mode -eq 'TeamsGetMissingDisplayName') {
                    return [pscustomobject]@{ id = $returnedTeamId }
                }
                return [pscustomobject]@{ id = $returnedTeamId; displayName = '20261015-AB-730' }
            }
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/teams/([^/]+)/members$') {
                if ($Mode -eq 'TeamsMemberConflictInFreshOnly') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 409 (Conflict).')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Conflict', [System.Management.Automation.ErrorCategory]::ResourceExists, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"Request_BadRequest","message":"One or more added object references already exist for the following modified properties: ''members''."}}')
                    throw $er
                }
                $b = if ($Body -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json } elseif ($Body -is [string]) { $Body | ConvertFrom-Json } else { $Body }
                $bindStr = "$($b.'user@odata.bind')"
                $upnMatch = if ($bindStr -match "uid-([^']+)") { $Matches[1] } else { 'user@contoso.com' }
                $memObj = [pscustomobject]@{
                    id     = "tm-$($requests.Count)"
                    userId = "uid-$upnMatch"
                    email  = $upnMatch
                    roles  = @($b.roles)
                }
                $state.AddedTeamMembers.Add($memObj)
                return $memObj
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/teams/([^/]+)/members(\?|$)') {
                if ($Mode -eq 'TeamsMemberVerifyMismatch') {
                    return [pscustomobject]@{
                        value = @(
                            [pscustomobject]@{ id = 'tm-1'; userId = 'uid-admin@contoso.com'; email = 'admin@contoso.com'; roles = @('owner') }
                            [pscustomobject]@{ id = 'tm-2'; userId = 'uid-sponsor@contoso.com'; email = 'sponsor@contoso.com'; roles = @('owner') }
                        )
                    }
                }
                if ($Mode -eq 'TeamsMemberForeignReplacement') {
                    return [pscustomobject]@{
                        value = @(
                            [pscustomobject]@{ id = 'tm-1'; userId = 'uid-admin@contoso.com'; email = 'admin@contoso.com'; roles = @('owner') }
                            [pscustomobject]@{ id = 'tm-2'; userId = 'uid-sponsor@contoso.com'; email = 'sponsor@contoso.com'; roles = @() }
                            [pscustomobject]@{ id = 'tm-3'; userId = 'uid-planner@contoso.com'; email = 'planner@contoso.com'; roles = @() }
                            [pscustomobject]@{ id = 'tm-4'; userId = 'uid-supply@contoso.com'; email = 'supply@contoso.com'; roles = @() }
                            [pscustomobject]@{ id = 'tm-5'; userId = 'uid-quality@contoso.com'; email = 'quality@contoso.com'; roles = @() }
                            [pscustomobject]@{ id = 'tm-6'; userId = 'uid-operations@contoso.com'; email = 'operations@contoso.com'; roles = @() }
                            [pscustomobject]@{ id = 'tm-7'; userId = 'uid-engineering@contoso.com'; email = 'engineering@contoso.com'; roles = @() }
                            [pscustomobject]@{ id = 'tm-8'; userId = 'uid-foreign@contoso.com'; email = 'foreign@contoso.com'; roles = @() }
                        )
                    }
                }
                return [pscustomobject]@{ value = @($state.AddedTeamMembers) }
            }

            # 6. SharePoint group, site, DemoKnowledge library, document upload
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/groups$') {
                $state.SiteGroupCreated = $true
                return [pscustomobject]@{ id = 'g-site-new-1'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; mailNickname = 'ab730-automotive-20261015' }
            }
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/groups/([^/]+)/members/\$ref$') {
                return $null
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/groups/([^/]+)/sites/root$') {
                return [pscustomobject]@{ id = 'site-new-1'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; webUrl = 'https://contoso.sharepoint.com/sites/ab730-automotive-20261015' }
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/sites/([^/]+)/drives') {
                $drives = @([pscustomobject]@{ id = 'drv-default'; name = 'Documents'; webUrl = 'https://contoso.sharepoint.com/sites/ab730/Shared%20Documents' })
                if ($state.DemoLibCreated -or $Mode -eq 'ExistingDemoKnowledgeFiles') {
                    $drives += [pscustomobject]@{ id = 'drv-demoknowledge'; name = 'DemoKnowledge'; webUrl = 'https://contoso.sharepoint.com/sites/ab730/DemoKnowledge' }
                }
                return [pscustomobject]@{ value = $drives }
            }
            if ($Method -eq 'POST' -and $decUri -match '/v1\.0/sites/([^/]+)/lists$') {
                $state.DemoLibCreated = $true
                return [pscustomobject]@{ id = 'list-demoknowledge'; displayName = 'DemoKnowledge' }
            }
            if ($Method -eq 'GET' -and $decUri -match '/v1\.0/drives/drv-demoknowledge/root/children') {
                if ($Mode -eq 'ExistingDemoKnowledgeFiles') {
                    return [pscustomobject]@{
                        value = @([pscustomobject]@{ id = 'sp-existing-1'; name = '06-Quality-Brief.docx'; file = [pscustomobject]@{} })
                    }
                }
                if ($Mode -eq 'SharePointVerifyMismatch' -and $state.UploadedSpDocs.Count -gt 0) {
                    return [pscustomobject]@{ value = @($state.UploadedSpDocs | Select-Object -First 2) }
                }
                return [pscustomobject]@{ value = @($state.UploadedSpDocs) }
            }
            if ($Method -eq 'PUT' -and $decUri -match '/v1\.0/drives/drv-demoknowledge/root:/([^:?]+):/content') {
                $fn = $Matches[1]
                if ($Mode -eq 'SharePointErrorLeakTest') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 409 (Conflict) Authorization: Bearer eyJSECRET_JWT.payload.sig.')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Conflict', [System.Management.Automation.ErrorCategory]::ResourceExists, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"nameAlreadyExists","message":"File exists url=https://contoso.sharepoint.com/sp?tempauth=SECRET_CAP_TOKEN \"@microsoft.graph.downloadUrl\":\"https://contoso.sharepoint.com/dl?sig=SECRET_DL\" client_secret=stub-secret-value-not-real"}}')
                    throw $er
                }
                if ($Mode -eq 'SharePointRaceConflictOnPut' -and $decUri -match 'conflictBehavior=fail') {
                    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 409 (Conflict).')
                    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'Conflict', [System.Management.Automation.ErrorCategory]::ResourceExists, $null)
                    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"nameAlreadyExists","message":"The specified item name already exists."}}')
                    throw $er
                }
                if ($Mode -eq 'SharePointPutNullResponse') {
                    return $null
                }
                $len = if ($Body -is [byte[]]) { $Body.Length } else { 1250 }
                $docItem = [pscustomobject]@{
                    id                             = "spdoc-$fn"
                    name                           = $fn
                    size                           = $len
                    file                           = [pscustomobject]@{}
                    '@microsoft.graph.downloadUrl' = 'https://contoso.sharepoint.com/spdownload?tempauth=STUBSECRET_SP'
                }
                $state.UploadedSpDocs.Add($docItem)
                return $docItem
            }

            throw "Unexpected Graph request in stub probe ($Mode): $Method $Uri"
        }.GetNewClosure()

        $global:__AB730GraphHandler = $graphHandler
        Set-Item Function:\global:Invoke-RestMethod -Value {
            param([string]$Method, [string]$Uri, $Headers, $Body, [string]$ContentType, $InFile, [string]$ResponseHeadersVariable, [string]$StatusCodeVariable)
            $global:__AB730LastHeaders = $null
            $global:__AB730LastStatus = $null
            $res = & $global:__AB730GraphHandler -Method $Method -Uri $Uri -Headers $Headers -Body $Body -ContentType $ContentType -InFile $InFile -ResponseHeadersVariable $ResponseHeadersVariable -StatusCodeVariable $StatusCodeVariable
            if ($ResponseHeadersVariable -and $null -ne $global:__AB730LastHeaders) {
                Set-Variable -Name $ResponseHeadersVariable -Value $global:__AB730LastHeaders -Scope 1
            }
            if ($StatusCodeVariable -and $null -ne $global:__AB730LastStatus) {
                Set-Variable -Name $StatusCodeVariable -Value $global:__AB730LastStatus -Scope 1
            }
            return $res
        }
        Set-Item Function:\global:Invoke-WebRequest -Value {
            param([string]$Uri, [string]$Method = 'GET', $Headers, $Body, [string]$ContentType)
            $requests.Add([pscustomobject]@{ Method = $Method.ToUpperInvariant(); Uri = $Uri; Body = $Body })
            if ($Method -eq 'POST' -and $Uri -match '/v1\.0/teams$') {
                $state.TeamCreated = $true
                $state.CreatedChannels.Add([pscustomobject]@{ id = 'ch-general'; displayName = 'General' })
                $teamHdrId = if ($Mode -eq 'Teams202EmptyBodyWithForeignRace') { 'team-202-hdr' } else { 'team-new-1' }
                $opHdrId   = if ($Mode -eq 'Teams202EmptyBodyWithForeignRace') { 'op-202-hdr' } else { 'op-1' }
                return [pscustomobject]@{
                    StatusCode = 202
                    Content    = ''
                    Headers    = @{
                        Location           = "/teams('$teamHdrId')/operations('$opHdrId')"
                        'Content-Location' = "/teams('$teamHdrId')"
                    }
                }
            }
            return [pscustomobject]@{ StatusCode = 202; Content = ''; Headers = @{ Location = 'https://graph.microsoft.com/v1.0/teams/team-new-1/operations/op-1' } }
        }.GetNewClosure()

        switch ($Target) {
            'Preflight' {
                & (Join-Path $envInfo.EngineDir 'Invoke-SeedPreflight.ps1') `
                    -ConfigPath $envInfo.ConfigPath `
                    -EmailsPath (Join-Path $envInfo.ScenarioDir 'emails.json') `
                    -TeamsMessagesPath (Join-Path $envInfo.ScenarioDir 'teams-messages.json') `
                    -SharePointPath (Join-Path $envInfo.ScenarioDir 'sharepoint-sites.json') `
                    -FilesManifestPath (Join-Path $envInfo.ScenarioDir 'files-manifest.json') `
                    -FreshOnly
            }
            'Upload' {
                & (Join-Path $envInfo.EngineDir 'Invoke-UploadFiles.ps1') `
                    -ConfigPath $envInfo.ConfigPath `
                    -ManifestPath (Join-Path $envInfo.ScenarioDir 'files-manifest.json') `
                    -StrictScenario `
                    -FreshOnly
            }
            'Emails' {
                & (Join-Path $envInfo.EngineDir 'Invoke-SeedEmails.ps1') `
                    -ConfigPath $envInfo.ConfigPath `
                    -EmailsPath (Join-Path $envInfo.ScenarioDir 'emails.json') `
                    -FreshOnly
            }
            'Teams' {
                & (Join-Path $envInfo.EngineDir 'Invoke-SeedTeamsChannel.ps1') `
                    -ConfigPath $envInfo.ConfigPath `
                    -TeamsMessagesPath (Join-Path $envInfo.ScenarioDir 'teams-messages.json') `
                    -FreshOnly
            }
            'SharePoint' {
                & (Join-Path $envInfo.EngineDir 'Invoke-SeedSharePoint.ps1') `
                    -ConfigPath $envInfo.ConfigPath `
                    -SharePointPath (Join-Path $envInfo.ScenarioDir 'sharepoint-sites.json') `
                    -FreshOnly
            }
            'RunnerPreflight' {
                $cfgToUse = if ($Overrides.ContainsKey('CustomConfigPath')) { $Overrides['CustomConfigPath'] } else { $envInfo.ConfigPath }
                $mfToUse  = if ($Overrides.ContainsKey('CustomManifestPath')) { $Overrides['CustomManifestPath'] } else { $envInfo.ManifestPath }
                $engToUse = if ($Overrides.ContainsKey('CustomEngineDir')) { $Overrides['CustomEngineDir'] } else { $envInfo.EngineDir }
                $runnerParams = @{
                    ConfigPath   = $cfgToUse
                    ManifestPath = $mfToUse
                    ScenarioDir  = $envInfo.ScenarioDir
                }
                if ($Overrides.ContainsKey('Phases')) { $runnerParams['Phases'] = $Overrides['Phases'] }
                & (Join-Path $engToUse 'Invoke-AB730.ps1') @runnerParams
            }
            'RunnerLive' {
                $cfgToUse = if ($Overrides.ContainsKey('CustomConfigPath')) { $Overrides['CustomConfigPath'] } else { $envInfo.ConfigPath }
                $mfToUse  = if ($Overrides.ContainsKey('CustomManifestPath')) { $Overrides['CustomManifestPath'] } else { $envInfo.ManifestPath }
                $engToUse = if ($Overrides.ContainsKey('CustomEngineDir')) { $Overrides['CustomEngineDir'] } else { $envInfo.EngineDir }
                $runnerParams = @{
                    ConfigPath   = $cfgToUse
                    ManifestPath = $mfToUse
                    ScenarioDir  = $envInfo.ScenarioDir
                    Live         = $true
                }
                if ($Overrides.ContainsKey('Phases')) { $runnerParams['Phases'] = $Overrides['Phases'] }
                & (Join-Path $engToUse 'Invoke-AB730.ps1') @runnerParams
            }
        }
    } catch {
        $threw = $true
        $caughtEr = $_
        $errorMsg = "$($caughtEr.Exception.Message)"
        $errorStatusCode = if ($caughtEr.Exception -and ($caughtEr.Exception.PSObject.Properties.Name -contains 'StatusCode')) { $caughtEr.Exception.StatusCode } else { $null }
    } finally {
        if ($hasPrevToken) { $global:AccessToken = $prevToken } else { Remove-Item Variable:\global:AccessToken -ErrorAction SilentlyContinue }
        if ($prevIrm) { Set-Item Function:\global:Invoke-RestMethod -Value $prevIrm.ScriptBlock } else { Remove-Item Function:\Invoke-RestMethod -ErrorAction SilentlyContinue }
        if ($prevIwr) { Set-Item Function:\global:Invoke-WebRequest -Value $prevIwr.ScriptBlock } else { Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue }
        if ($prevWh)  { Set-Item Function:\global:Write-Host -Value $prevWh.ScriptBlock } else { Remove-Item Function:\Write-Host -ErrorAction SilentlyContinue }
        if ($prevGit) { Set-Item Function:\global:git -Value $prevGit.ScriptBlock } else { Remove-Item Function:\git -ErrorAction SilentlyContinue }
        if ($prevSleep) { Set-Item Function:\global:Start-Sleep -Value $prevSleep.ScriptBlock } else { Remove-Item Function:\Start-Sleep -ErrorAction SilentlyContinue }
        if ($prevSetMgCtx) { Set-Item Function:\global:Set-MgRequestContext -Value $prevSetMgCtx.ScriptBlock } else { Remove-Item Function:\Set-MgRequestContext -ErrorAction SilentlyContinue }
        if ($prevGetMgCtx) { Set-Item Function:\global:Get-MgRequestContext -Value $prevGetMgCtx.ScriptBlock } else { Remove-Item Function:\Get-MgRequestContext -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $probeRoot) {
            Remove-Item -LiteralPath $probeRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Set-StrictMode -Version Latest
    }

    return [pscustomobject]@{
        Threw                  = $threw
        ErrorMessage           = $errorMsg
        ErrorStatusCode        = $errorStatusCode
        ErrorRecord            = $caughtEr
        Requests               = @($requests)
        TokenRequests          = @($tokenRequests)
        Stdout                 = ($outputLines -join "`n")
        MgMaxRetry             = $state.MgMaxRetry
        MgMaxRetrySetCalls     = $state.MgMaxRetrySetCalls
        MgMaxRetryAtFirstWrite = $state.MgMaxRetryAtFirstWrite
        SleepDelays            = @($state.SleepDelays)
    }
}

# 1. Preflight -FreshOnly clean vs later-page collisions
$pfClean = Invoke-AB730Probe -Target 'Preflight' -Mode 'Clean'
Assert-True 'AB-730 Preflight (-FreshOnly, clean tenant): succeeds' {
    -not $pfClean.Threw
} "expected clean preflight to succeed, got: $($pfClean.ErrorMessage)"
Assert-Value 'AB-730 Preflight (-FreshOnly, clean tenant): issues 0 writes' 0 {
    @($pfClean.Requests | Where-Object { $_.Method -ne 'GET' }).Count
}

$pfTeamPage2 = Invoke-AB730Probe -Target 'Preflight' -Mode 'Page2TeamCollision'
Assert-True 'AB-730 Preflight (-FreshOnly): rejects later-page (page 2) Team displayName collision' {
    $pfTeamPage2.Threw -and ($pfTeamPage2.ErrorMessage -match '20261015-AB-730|collision|already exists')
} "expected page-2 Team collision abort, got Threw=$($pfTeamPage2.Threw) '$($pfTeamPage2.ErrorMessage)'"
Assert-Value 'AB-730 Preflight (-FreshOnly, page-2 Team collision): issues 0 writes' 0 {
    @($pfTeamPage2.Requests | Where-Object { $_.Method -ne 'GET' }).Count
}

$pfSiteAliasPage2 = Invoke-AB730Probe -Target 'Preflight' -Mode 'Page2SiteAliasCollision'
Assert-True 'AB-730 Preflight (-FreshOnly): rejects later-page (page 2) SharePoint alias collision' {
    $pfSiteAliasPage2.Threw -and ($pfSiteAliasPage2.ErrorMessage -match 'ab730-automotive-20261015|collision|already exists')
} "expected page-2 Site alias collision abort, got Threw=$($pfSiteAliasPage2.Threw) '$($pfSiteAliasPage2.ErrorMessage)'"

$pfSiteDispPage2 = Invoke-AB730Probe -Target 'Preflight' -Mode 'Page2SiteDisplayNameCollision'
Assert-True 'AB-730 Preflight (-FreshOnly): rejects later-page (page 2) SharePoint displayName collision' {
    $pfSiteDispPage2.Threw -and ($pfSiteDispPage2.ErrorMessage -match 'AB-730 - Automotive Knowledge - 20261015|collision|already exists')
} "expected page-2 Site displayName collision abort, got Threw=$($pfSiteDispPage2.Threw) '$($pfSiteDispPage2.ErrorMessage)'"

$pfMailPage2 = Invoke-AB730Probe -Target 'Preflight' -Mode 'Page2MailCollision'
Assert-True 'AB-730 Preflight (-FreshOnly): rejects later-page (page 2) Admin Inbox subject collision via mailFolders(''inbox'')' {
    $pfMailPage2.Threw -and ($pfMailPage2.ErrorMessage -match '供应商交期|專案啟動|collision|already')
} "expected page-2 Mail subject collision abort, got Threw=$($pfMailPage2.Threw) '$($pfMailPage2.ErrorMessage)'"

$pfFolderExists = Invoke-AB730Probe -Target 'Preflight' -Mode 'ExistingOneDriveFolder'
Assert-True 'AB-730 Preflight (-FreshOnly): rejects existing OneDrive target folder' {
    $pfFolderExists.Threw -and ($pfFolderExists.ErrorMessage -match 'AB730-Automotive-20261015|already exists|collision')
} "expected existing OneDrive folder abort, got Threw=$($pfFolderExists.Threw) '$($pfFolderExists.ErrorMessage)'"

# 2. Owning-phase race guards (-FreshOnly): zero write before guard
$notMissingCmdOrParam = 'A parameter cannot be found|is not recognized as a name'

$upRace = Invoke-AB730Probe -Target 'Upload' -Mode 'OwningRaceOneDrive'
Assert-True 'AB-730 Upload (-FreshOnly): owning-phase guard rejects existing OneDrive folder before any write' {
    $upRace.Threw -and ($upRace.ErrorMessage -notmatch $notMissingCmdOrParam)
} "expected Upload owning-phase race to abort, got: $($upRace.ErrorMessage)"
Assert-Value 'AB-730 Upload (-FreshOnly, owning-phase race): 0 POST/PUT writes' 0 {
    @($upRace.Requests | Where-Object { $_.Method -ne 'GET' }).Count
}

$upPostRace = Invoke-AB730Probe -Target 'Upload' -Mode 'PostCreateRaceOneDrive'
Assert-True 'AB-730 Upload (-FreshOnly): POST /drive/root/children nameAlreadyExists aborts instead of adopting folder' {
    $upPostRace.Threw -and ($upPostRace.ErrorMessage -notmatch $notMissingCmdOrParam)
} 'must not adopt folder on nameAlreadyExists in FreshOnly mode'
Assert-Value 'AB-730 Upload (-FreshOnly, POST conflict): 0 PUT file uploads after conflict' 0 {
    @($upPostRace.Requests | Where-Object { $_.Method -eq 'PUT' }).Count
}

$emRace = Invoke-AB730Probe -Target 'Emails' -Mode 'OwningRaceEmails'
Assert-True 'AB-730 Emails (-FreshOnly): owning-phase guard rejects existing thread on page 2 before any sendMail' {
    $emRace.Threw -and ($emRace.ErrorMessage -notmatch $notMissingCmdOrParam)
} "expected Emails owning-phase race to abort, got: $($emRace.ErrorMessage)"
Assert-Value 'AB-730 Emails (-FreshOnly, owning-phase race): 0 POST sendMail/reply writes' 0 {
    @($emRace.Requests | Where-Object { $_.Method -ne 'GET' }).Count
}

$emNoAdminCc = Invoke-AB730Probe -Target 'Emails' -Mode 'Clean' -Overrides @{
    ModifyEmails = {
        param($em)
        $em.emailThreads[0].emails[0].ccRoles = @()
    }
}
Assert-True 'AB-730 Emails (-FreshOnly): rejects message without Admin in ccRoles before any write' {
    $emNoAdminCc.Threw -and ($emNoAdminCc.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($emNoAdminCc.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} 'every AB-730 message must CC Admin'

$emLookupFail = Invoke-AB730Probe -Target 'Emails' -Mode 'MailReplyLookupFails'
Assert-True 'AB-730 Emails (-FreshOnly): aborts immediately without retry when sender Inbox internetMessageId lookup fails' {
    $emLookupFail.Threw -and ($emLookupFail.ErrorMessage -notmatch $notMissingCmdOrParam)
} 'must abort when newly sent message cannot be resolved by current-run internetMessageId'
Assert-Value 'AB-730 Emails (-FreshOnly, reply lookup failure): sends only the 1 first message with 0 retries' 1 {
    @($emLookupFail.Requests | Where-Object { $_.Method -eq 'POST' }).Count
}

$tmRace = Invoke-AB730Probe -Target 'Teams' -Mode 'OwningRaceTeams'
Assert-True 'AB-730 Teams (-FreshOnly): owning-phase guard rejects page-2 existing group before POST /teams' {
    $tmRace.Threw -and ($tmRace.ErrorMessage -notmatch $notMissingCmdOrParam)
} "expected Teams owning-phase race to abort, got: $($tmRace.ErrorMessage)"
Assert-Value 'AB-730 Teams (-FreshOnly, owning-phase race): 0 POST/PUT/PATCH writes' 0 {
    @($tmRace.Requests | Where-Object { $_.Method -ne 'GET' }).Count
}

$tmFuture = Invoke-AB730Probe -Target 'Teams' -Mode 'Clean' -Overrides @{
    ModifyTeams = {
        param($tm)
        $tm.channels[0].messages[0].createdDateTime = [DateTimeOffset]::UtcNow.AddDays(5).ToString('o')
    }
}
Assert-True 'AB-730 Teams (-FreshOnly): rejects future createdDateTime before any write' {
    $tmFuture.Threw -and ($tmFuture.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($tmFuture.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} 'migration messages must have timestamps strictly in the past'

$spRaceAlias = Invoke-AB730Probe -Target 'SharePoint' -Mode 'OwningRaceSiteAlias'
Assert-True 'AB-730 SharePoint (-FreshOnly): owning-phase guard rejects page-2 alias collision before POST /groups' {
    $spRaceAlias.Threw -and ($spRaceAlias.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($spRaceAlias.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} "expected SharePoint alias race to abort with 0 writes, got Threw=$($spRaceAlias.Threw)"

$spRaceDisp = Invoke-AB730Probe -Target 'SharePoint' -Mode 'OwningRaceSiteDisplayName'
Assert-True 'AB-730 SharePoint (-FreshOnly): owning-phase guard rejects page-2 displayName collision before POST /groups' {
    $spRaceDisp.Threw -and ($spRaceDisp.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($spRaceDisp.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} "expected SharePoint displayName race to abort with 0 writes, got Threw=$($spRaceDisp.Threw)"

$spExistingDocs = Invoke-AB730Probe -Target 'SharePoint' -Mode 'ExistingDemoKnowledgeFiles'
Assert-True 'AB-730 SharePoint (-FreshOnly): aborts before PUT when DemoKnowledge already contains a file' {
    $spExistingDocs.Threw -and ($spExistingDocs.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($spExistingDocs.Requests | Where-Object { $_.Method -eq 'PUT' }).Count -eq 0)
} 'DemoKnowledge must be empty in FreshOnly mode'

# 3. Invoke-AB730.ps1 Runner contract & fail-closed guards
$runDefaultPreflight = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean'
Assert-True 'Invoke-AB730.ps1 (default preflight-only): succeeds on clean tenant' {
    -not $runDefaultPreflight.Threw
} "expected default preflight-only run to succeed, got: $($runDefaultPreflight.ErrorMessage)"
Assert-Value 'Invoke-AB730.ps1 (default preflight-only): issues 0 write requests' 0 {
    @($runDefaultPreflight.Requests | Where-Object { $_.Method -ne 'GET' }).Count
}

$runTrackedCfg = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{ GitTracked = $true }
Assert-True 'Invoke-AB730.ps1: aborts before any Graph call when config/manifest is git-tracked' {
    $runTrackedCfg.Threw -and ($runTrackedCfg.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runTrackedCfg.Requests.Count -eq 0)
} 'tracked config or manifest must fail closed before Graph calls'

$runUnignoredCfg = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{ GitNotIgnored = $true }
Assert-True 'Invoke-AB730.ps1: aborts before any Graph call when config/manifest is not git-ignored' {
    $runUnignoredCfg.Threw -and ($runUnignoredCfg.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runUnignoredCfg.Requests.Count -eq 0)
} 'non-ignored config or manifest must fail closed before Graph calls'

$runCopiedCfg = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    ModifyConfig = {
        param($cfg, $envInfo)
        $copied = Join-Path $envInfo.ScenarioDir 'config.json'
        $cfg | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $copied -Encoding UTF8
    }
    CustomConfigPath = (Join-Path $TestsDir '.tmp-ab730-probe\seed-data\scenarios\ab730-automotive-20261015\config.json')
}
Assert-True 'Invoke-AB730.ps1: rejects ConfigPath inside scenario directory before any Graph call' {
    $runCopiedCfg.Threw -and ($runCopiedCfg.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runCopiedCfg.Requests.Count -eq 0)
} 'root config must never be copied into scenario directory'

$runEscapingSource = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    ModifyConfig = {
        param($cfg)
        $cfg.filesSourceDir = '..\..\outside-folder'
    }
}
Assert-True 'Invoke-AB730.ps1: rejects filesSourceDir escaping scenario DEMO-FILE before any Graph call' {
    $runEscapingSource.Threw -and ($runEscapingSource.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runEscapingSource.Requests.Count -eq 0)
} 'filesSourceDir must resolve to scenario DEMO-FILE'

$runBadHash = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{ TamperDemoFile = $true }
Assert-True 'Invoke-AB730.ps1: rejects modified artifact SHA-256 hash before any Graph call' {
    $runBadHash.Threw -and ($runBadHash.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runBadHash.Requests.Count -eq 0)
} 'modified DEMO-FILE or JSON hash must fail closed before Graph calls'

$runManifestMismatch = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    ModifyManifest = {
        param($mf)
        $mf.roles.Supply.upn = 'wrong-supply@contoso.com'
    }
}
Assert-True 'Invoke-AB730.ps1: rejects config vs manifest role UPN mismatch before any Graph call' {
    $runManifestMismatch.Threw -and ($runManifestMismatch.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runManifestMismatch.Requests.Count -eq 0)
} 'manifest and config role bindings must match exactly'

$runDupRole = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    ModifyConfig = { param($cfg) $cfg.roles.Sponsor.upn = 'admin@contoso.com' }
    ModifyManifest = { param($mf) $cfg.roles.Sponsor.upn = 'admin@contoso.com' }
}
Assert-True 'Invoke-AB730.ps1: rejects Admin reused as a business role before any Graph call' {
    $runDupRole.Threw -and ($runDupRole.ErrorMessage -notmatch $notMissingCmdOrParam) -and ($runDupRole.Requests.Count -eq 0)
} 'Admin must not be reused as any business role'

$runBadJwtTenant = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    Jwt = (New-AB730TestJwt -TenantId '99999999-9999-9999-9999-999999999999')
}
Assert-True 'Invoke-AB730.ps1: rejects runtime token with mismatched tid claim' {
    $runBadJwtTenant.Threw -and ($runBadJwtTenant.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($runBadJwtTenant.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} 'token tid mismatch must abort'

$runExpiredJwt = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    Jwt = (New-AB730TestJwt -Exp ([DateTimeOffset]::UtcNow.AddMinutes(-10).ToUnixTimeSeconds()))
}
Assert-True 'Invoke-AB730.ps1: rejects expired runtime token' {
    $runExpiredJwt.Threw -and ($runExpiredJwt.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($runExpiredJwt.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} 'expired token must abort'

$runMissingRoleJwt = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    Jwt = (New-AB730TestJwt -Roles @('User.Read.All','Group.ReadWrite.All'))
}
Assert-True 'Invoke-AB730.ps1: rejects runtime token missing required application permissions' {
    $runMissingRoleJwt.Threw -and ($runMissingRoleJwt.ErrorMessage -notmatch $notMissingCmdOrParam) -and (@($runMissingRoleJwt.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} 'token missing required permissions must abort'

# 4. Full Invoke-AB730.ps1 -Live offline probe on clean stubbed tenant
$runLive = Invoke-AB730Probe -Target 'RunnerLive' -Mode 'Clean'
Assert-True 'Invoke-AB730.ps1 (-Live, clean tenant): completes all 4 write surfaces' {
    -not $runLive.Threw
} "expected -Live run to succeed, got: $($runLive.ErrorMessage)"

Assert-Value 'Invoke-AB730.ps1 (-Live): uploads exactly 8 OneDrive Office files (plus at most 1 engine proof metadata file)' 8 {
    @($runLive.Requests | Where-Object {
        $_.Method -eq 'PUT' -and (
            $_.Uri -match '/drive/root:/AB730-Automotive-20261015/' -or
            ($_.Uri -match '^https://stub-upload\.invalid/' -and $_.Uri -notmatch '\.ms4018-seed-proof\.json')
        )
    }).Count
}
Assert-Value 'Invoke-AB730.ps1 (-Live): sends 5 root emails and 10 replies (15 total)' 15 {
    @($runLive.Requests | Where-Object { $_.Method -eq 'POST' -and ($_.Uri -match '/sendMail$' -or $_.Uri -match '/reply$') }).Count
}
Assert-Value 'Invoke-AB730.ps1 (-Live): creates 5 Teams business channels and posts 15 messages (10 top-level + 5 replies)' 15 {
    @($runLive.Requests | Where-Object { $_.Method -eq 'POST' -and ($_.Uri -match '/channels/[^/]+/messages$' -or $_.Uri -match '/replies$') }).Count
}
Assert-True 'Invoke-AB730.ps1 (-Live): adds Team members only AFTER team completeMigration, with Admin as sole owner and 7 business members as non-owners' {
    $reqs = @($runLive.Requests)
    $teamCompleteIdx = -1
    $memberReqs = @()
    for ($i = 0; $i -lt $reqs.Count; $i++) {
        if ($reqs[$i].Method -eq 'POST' -and $reqs[$i].Uri -match '/teams/[^/]+/completeMigration$') {
            $teamCompleteIdx = $i
        }
        if ($reqs[$i].Method -eq 'POST' -and $reqs[$i].Uri -match '/teams/[^/]+/members$') {
            $bodyStr = if ($reqs[$i].Body -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($reqs[$i].Body) } else { "$($reqs[$i].Body)" }
            $memberReqs += [pscustomobject]@{ Index = $i; Body = ($bodyStr | ConvertFrom-Json) }
        }
    }
    if ($teamCompleteIdx -lt 0 -or $memberReqs.Count -ne 8) { return $false }
    $allAfterComplete = (@($memberReqs | Where-Object { $_.Index -le $teamCompleteIdx }).Count -eq 0)
    $owners = @($memberReqs | Where-Object { @($_.Body.roles) -contains 'owner' })
    $nonOwners = @($memberReqs | Where-Object { @($_.Body.roles).Count -eq 0 })
    $allAfterComplete -and ($owners.Count -eq 1) -and ($nonOwners.Count -eq 7) -and
    ($owners[0].Body.'user@odata.bind' -match 'admin%40contoso\.com|admin@contoso\.com')
} 'Team members must be added after team completeMigration with Admin sole owner and 7 business roles as plain members'

Assert-Value 'Invoke-AB730.ps1 (-Live): uploads 3 .docx files to SharePoint DemoKnowledge library' 3 {
    @($runLive.Requests | Where-Object { $_.Method -eq 'PUT' -and $_.Uri -match '/drives/drv-demoknowledge/root:/' }).Count
}
Assert-Value 'Invoke-AB730.ps1 (-Live): creates 1 SharePoint document library (DemoKnowledge) and 0 custom lists' 1 {
    @($runLive.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/sites/[^/]+/lists$' }).Count
}
Assert-Value 'Invoke-AB730.ps1 (-Live): issues 0 DELETE, 0 PATCH, 0 /beta/, and 0 /v1.0/users collection listing requests' 0 {
    @($runLive.Requests | Where-Object {
        $_.Method -in @('DELETE','PATCH') -or
        $_.Uri -match '/beta/' -or
        $_.Uri -match '/v1\.0/users(\?|$)'
    }).Count
}
Assert-True 'Invoke-AB730.ps1 (-Live): stdout never leaks tempauth, downloadUrl, or clientSecret' {
    ($runLive.Stdout -notmatch 'tempauth=') -and
    ($runLive.Stdout -notmatch 'downloadUrl') -and
    ($runLive.Stdout -notmatch 'STUBSECRET') -and
    ($runLive.Stdout -notmatch 'stub-secret-value-not-real')
} 'stdout must never leak pre-authenticated URLs or secrets'

# 5. Targeted Council Blocking Findings (RED-B1 through RED-B7)
$redB1MissingHashes = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    ModifyManifest = {
        param($mf)
        $mf.PSObject.Properties.Remove('artifactHashes')
        $mf.PSObject.Properties.Remove('dataSpecSha256')
        $mf.PSObject.Properties.Remove('codeHashes')
    }
}
Assert-True 'RED-B1-1: Invoke-AB730.ps1 rejects manifest missing mandatory artifactHashes/dataSpecSha256/codeHashes with 0 token/Graph requests' {
    $redB1MissingHashes.Threw -and
    ($redB1MissingHashes.ErrorMessage -match 'artifactHashes|dataSpecSha256|codeHashes') -and
    ($redB1MissingHashes.TokenRequests.Count -eq 0) -and
    ($redB1MissingHashes.Requests.Count -eq 0)
} "expected missing pinned hashes to fail closed with 0 token/Graph requests, got Threw=$($redB1MissingHashes.Threw) Err='$($redB1MissingHashes.ErrorMessage)'"

$redB1SelfRecomputedNoExt = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $docPath = Join-Path $envInfo.DemoDir '06-Quality-Brief.docx'
        [System.IO.File]::WriteAllBytes($docPath, [System.Text.Encoding]::UTF8.GetBytes(('MUTATED-DOC-BYTES-' * 80)))
        $newHash = (Get-FileHash -LiteralPath $docPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $fmPath = Join-Path $envInfo.ScenarioDir 'files-manifest.json'
        $fmObj = Get-Content -LiteralPath $fmPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($fEntry in @($fmObj.files)) {
            if ($fEntry.localName -eq '06-Quality-Brief.docx') { $fEntry.sha256 = $newHash }
        }
        $fmObj.assetSHA256.'06-Quality-Brief.docx' = $newHash
        $fmObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $fmPath -Encoding UTF8
        $mfObj = Get-Content -LiteralPath $envInfo.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $mfObj.PSObject.Properties.Remove('artifactHashes')
        $mfObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $envInfo.ManifestPath -Encoding UTF8
    }
}
Assert-True 'RED-B1-2a: Invoke-AB730.ps1 rejects mutated DEMO-FILE with recomputed public files-manifest.json self-hashes when artifactHashes is omitted (0 token/Graph requests)' {
    $redB1SelfRecomputedNoExt.Threw -and
    ($redB1SelfRecomputedNoExt.ErrorMessage -match 'hash|mismatch|artifactHashes|06-Quality-Brief\.docx') -and
    ($redB1SelfRecomputedNoExt.TokenRequests.Count -eq 0) -and
    ($redB1SelfRecomputedNoExt.Requests.Count -eq 0)
} "expected self-recomputed public files-manifest.json without artifactHashes to fail closed, got Threw=$($redB1SelfRecomputedNoExt.Threw) Err='$($redB1SelfRecomputedNoExt.ErrorMessage)'"

$redB1SelfRecomputedPinned = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $docPath = Join-Path $envInfo.DemoDir '06-Quality-Brief.docx'
        [System.IO.File]::WriteAllBytes($docPath, [System.Text.Encoding]::UTF8.GetBytes(('MUTATED-DOC-BYTES-' * 80)))
        $newHash = (Get-FileHash -LiteralPath $docPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $fmPath = Join-Path $envInfo.ScenarioDir 'files-manifest.json'
        $fmObj = Get-Content -LiteralPath $fmPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($fEntry in @($fmObj.files)) {
            if ($fEntry.localName -eq '06-Quality-Brief.docx') { $fEntry.sha256 = $newHash }
        }
        $fmObj.assetSHA256.'06-Quality-Brief.docx' = $newHash
        $fmObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $fmPath -Encoding UTF8
    }
}
Assert-True 'RED-B1-2b: Invoke-AB730.ps1 rejects mutated DEMO-FILE with recomputed public files-manifest.json self-hashes when trusted manifest pins original hashes (0 token/Graph requests)' {
    $redB1SelfRecomputedPinned.Threw -and
    ($redB1SelfRecomputedPinned.ErrorMessage -match 'hash|mismatch|artifactHashes|06-Quality-Brief\.docx') -and
    ($redB1SelfRecomputedPinned.TokenRequests.Count -eq 0) -and
    ($redB1SelfRecomputedPinned.Requests.Count -eq 0)
} "expected pinned manifest hashes to reject self-recomputed public files-manifest.json, got Threw=$($redB1SelfRecomputedPinned.Threw) Err='$($redB1SelfRecomputedPinned.ErrorMessage)'"

$redB1MutatedEmailActor = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $emPath = Join-Path $envInfo.ScenarioDir 'emails.json'
        $emObj = Get-Content -LiteralPath $emPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $emObj.emailThreads[0].emails[0].fromRole = 'Sponsor'
        $emObj.emailThreads[0].emails[0].ccRoles = @('Admin', 'Analytics')
        $emObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $emPath -Encoding UTF8
        $mfObj = Get-Content -LiteralPath $envInfo.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $mfObj.artifactHashes.'emails.json' = (Get-FileHash -LiteralPath $emPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $mfObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $envInfo.ManifestPath -Encoding UTF8
    }
}
Assert-True 'RED-B1-3: Invoke-AB730.ps1 rejects mutated email sender/CC/body against pinned ab730-data-spec.json matrix with 0 token/Graph requests' {
    $redB1MutatedEmailActor.Threw -and
    ($redB1MutatedEmailActor.ErrorMessage -notmatch $notMissingCmdOrParam) -and
    ($redB1MutatedEmailActor.TokenRequests.Count -eq 0) -and
    ($redB1MutatedEmailActor.Requests.Count -eq 0)
} "expected mutated emails.json to fail canonical spec matrix check with 0 token/Graph requests, got Threw=$($redB1MutatedEmailActor.Threw) Err='$($redB1MutatedEmailActor.ErrorMessage)'"

$redB1MutatedTeamsActor = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $tmPath = Join-Path $envInfo.ScenarioDir 'teams-messages.json'
        $tmObj = Get-Content -LiteralPath $tmPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $tmObj.channels[0].messages[0].fromRole = 'Analytics'
        $tmObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $tmPath -Encoding UTF8
        $mfObj = Get-Content -LiteralPath $envInfo.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $mfObj.artifactHashes.'teams-messages.json' = (Get-FileHash -LiteralPath $tmPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $mfObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $envInfo.ManifestPath -Encoding UTF8
    }
}
Assert-True 'RED-B1-4: Invoke-AB730.ps1 rejects mutated Teams message author/content against pinned ab730-data-spec.json matrix with 0 token/Graph requests' {
    $redB1MutatedTeamsActor.Threw -and
    ($redB1MutatedTeamsActor.ErrorMessage -notmatch $notMissingCmdOrParam) -and
    ($redB1MutatedTeamsActor.TokenRequests.Count -eq 0) -and
    ($redB1MutatedTeamsActor.Requests.Count -eq 0)
} "expected mutated teams-messages.json to fail canonical spec matrix check with 0 token/Graph requests, got Threw=$($redB1MutatedTeamsActor.Threw) Err='$($redB1MutatedTeamsActor.ErrorMessage)'"

$redB1MutatedSpMembers = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $spPath = Join-Path $envInfo.ScenarioDir 'sharepoint-sites.json'
        $spObj = Get-Content -LiteralPath $spPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $spObj.sites[0].members = @('Sponsor', 'Sponsor', 'Planner', 'Supply', 'Quality', 'Operations', 'Engineering')
        $spObj.sites[0].documents[0].sourceFilename = '01-BOM-Cost.xlsx'
        $spObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $spPath -Encoding UTF8
        $mfObj = Get-Content -LiteralPath $envInfo.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $mfObj.artifactHashes.'sharepoint-sites.json' = (Get-FileHash -LiteralPath $spPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $mfObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $envInfo.ManifestPath -Encoding UTF8
    }
}
Assert-True 'RED-B1-5: Invoke-AB730.ps1 rejects mutated SharePoint members/sourceFilename against canonical matrix with 0 token/Graph requests' {
    $redB1MutatedSpMembers.Threw -and
    ($redB1MutatedSpMembers.ErrorMessage -notmatch $notMissingCmdOrParam) -and
    ($redB1MutatedSpMembers.TokenRequests.Count -eq 0) -and
    ($redB1MutatedSpMembers.Requests.Count -eq 0)
} "expected mutated sharepoint-sites.json members/documents to fail closed with 0 token/Graph requests, got Threw=$($redB1MutatedSpMembers.Threw) Err='$($redB1MutatedSpMembers.ErrorMessage)'"

$redB1MutatedCodeHash = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    ModifyManifest = {
        param($mf)
        $mf.codeHashes.'seed-data/engine/Invoke-SeedSharePoint.ps1' = ('0' * 64)
    }
}
Assert-True 'RED-B1-6: Invoke-AB730.ps1 rejects mismatched codeHashes before any token/Graph request' {
    $redB1MutatedCodeHash.Threw -and
    ($redB1MutatedCodeHash.ErrorMessage -match 'codeHashes|Invoke-SeedSharePoint\.ps1|hash') -and
    ($redB1MutatedCodeHash.TokenRequests.Count -eq 0) -and
    ($redB1MutatedCodeHash.Requests.Count -eq 0)
} "expected mismatched codeHashes to fail closed with 0 token/Graph requests, got Threw=$($redB1MutatedCodeHash.Threw) Err='$($redB1MutatedCodeHash.ErrorMessage)'"

$redB2NonTeamGroup = Invoke-AB730Probe -Target 'Teams' -Mode 'OwningRaceNonTeamGroup'
Assert-True 'RED-B2-1: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) aborts when a non-Team group with the same displayName exists on page 2 (0 writes)' {
    $redB2NonTeamGroup.Threw -and
    ($redB2NonTeamGroup.ErrorMessage -match '20261015-AB-730|collision|already exists') -and
    (@($redB2NonTeamGroup.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} "expected non-Team group displayName collision to abort before POST /teams, got Threw=$($redB2NonTeamGroup.Threw) Writes=$(@($redB2NonTeamGroup.Requests | Where-Object { $_.Method -ne 'GET' }).Count)"

$redB2EmptyBody202 = Invoke-AB730Probe -Target 'Teams' -Mode 'Teams202EmptyBodyWithForeignRace'
Assert-True 'RED-B2-2: Invoke-SeedTeamsChannel.ps1 resolves Team ID from HTTP 202 Location/Content-Location headers and never adopts foreign displayName match' {
    (-not $redB2EmptyBody202.Threw) -and
    (@($redB2EmptyBody202.Requests | Where-Object { $_.Uri -match 'team-202-hdr' }).Count -gt 0) -and
    (@($redB2EmptyBody202.Requests | Where-Object { $_.Uri -match 'team-foreign-race' }).Count -eq 0)
} "expected HTTP 202 header-based Team ID resolution (team-202-hdr) and 0 calls to team-foreign-race, got Threw=$($redB2EmptyBody202.Threw) Err='$($redB2EmptyBody202.ErrorMessage)'"

$redB4Service503 = Invoke-AB730Probe -Target 'Upload' -Mode 'OneDrive503WithItemNotFoundText'
Assert-True 'RED-B4-2: Invoke-UploadFiles.ps1 (-FreshOnly) aborts on HTTP 503 even when error body contains itemNotFound text (0 writes)' {
    $redB4Service503.Threw -and
    (@($redB4Service503.Requests | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
} "expected HTTP 503 with itemNotFound text to fail closed with 0 writes, got Threw=$($redB4Service503.Threw) Writes=$(@($redB4Service503.Requests | Where-Object { $_.Method -ne 'GET' }).Count)"

$redB5EscapedSource = Invoke-AB730Probe -Target 'SharePoint' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $outsideFile = Join-Path $envInfo.RootDir 'outside-secret.docx'
        [System.IO.File]::WriteAllBytes($outsideFile, [System.Text.Encoding]::UTF8.GetBytes(('OUTSIDE-DOC-' * 100)))
        $spPath = Join-Path $envInfo.ScenarioDir 'sharepoint-sites.json'
        $spObj = Get-Content -LiteralPath $spPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $spObj.sites[0].documents[0].sourceFilename = '..\..\..\..\outside-secret.docx'
        $spObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $spPath -Encoding UTF8
    }
}
Assert-True 'RED-B5-1: Invoke-SeedSharePoint.ps1 (-FreshOnly) rejects filePath escaping DEMO-FILE before any PUT upload' {
    $redB5EscapedSource.Threw -and
    (@($redB5EscapedSource.Requests | Where-Object { $_.Method -eq 'PUT' }).Count -eq 0)
} "expected escaped SharePoint filePath to abort with 0 PUT uploads, got Threw=$($redB5EscapedSource.Threw) PUTs=$(@($redB5EscapedSource.Requests | Where-Object { $_.Method -eq 'PUT' }).Count)"

$redB5ConflictPut = Invoke-AB730Probe -Target 'SharePoint' -Mode 'SharePointRaceConflictOnPut'
Assert-True 'RED-B5-2: Invoke-SeedSharePoint.ps1 (-FreshOnly) passes @microsoft.graph.conflictBehavior=fail on PUT and aborts on 409 conflict' {
    $putReqs = @($redB5ConflictPut.Requests | Where-Object { $_.Method -eq 'PUT' })
    $redB5ConflictPut.Threw -and
    ($putReqs.Count -eq 1) -and
    ($putReqs[0].Uri -match 'conflictBehavior=fail')
} "expected PUT with @microsoft.graph.conflictBehavior=fail to abort on 1st conflict, got Threw=$($redB5ConflictPut.Threw) PUTs=$(@($redB5ConflictPut.Requests | Where-Object { $_.Method -eq 'PUT' }).Count)"

$redB6TeamMemberConflict = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsMemberConflictInFreshOnly'
Assert-True 'RED-B6-1: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) aborts when POST /teams/{id}/members returns 409 already exists' {
    $redB6TeamMemberConflict.Threw -and
    ($redB6TeamMemberConflict.ErrorMessage -match '409|already exist|member|owner|Conflict')
} "expected Teams member conflict in FreshOnly mode to abort, got Threw=$($redB6TeamMemberConflict.Threw) Err='$($redB6TeamMemberConflict.ErrorMessage)'"

$redB6TeamVerifyMismatch = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsMemberVerifyMismatch'
Assert-True 'RED-B6-2: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) verifies final owners/members via GET and aborts on unexpected owner/member count' {
    $redB6TeamVerifyMismatch.Threw -and
    ($redB6TeamVerifyMismatch.ErrorMessage -match 'owner|member|verification|mismatch')
} "expected post-write Team member verification mismatch to abort, got Threw=$($redB6TeamVerifyMismatch.Threw) Err='$($redB6TeamVerifyMismatch.ErrorMessage)'"

$redB6SpPutNull = Invoke-AB730Probe -Target 'SharePoint' -Mode 'SharePointPutNullResponse'
Assert-True 'RED-B6-3: Invoke-SeedSharePoint.ps1 (-FreshOnly) aborts when PUT returns null instead of a valid DriveItem' {
    $redB6SpPutNull.Threw -and
    (@($redB6SpPutNull.Requests | Where-Object { $_.Method -eq 'PUT' }).Count -eq 1)
} "expected null PUT response in SharePoint upload to abort immediately, got Threw=$($redB6SpPutNull.Threw) PUTs=$(@($redB6SpPutNull.Requests | Where-Object { $_.Method -eq 'PUT' }).Count)"

$redB6SpVerifyMismatch = Invoke-AB730Probe -Target 'SharePoint' -Mode 'SharePointVerifyMismatch'
Assert-True 'RED-B6-4: Invoke-SeedSharePoint.ps1 (-FreshOnly) verifies final DemoKnowledge files and group owners/members via GET and aborts on mismatch' {
    $redB6SpVerifyMismatch.Threw -and
    ($redB6SpVerifyMismatch.ErrorMessage -match 'verification|member|DemoKnowledge|mismatch|count')
} "expected post-write SharePoint verification mismatch to abort, got Threw=$($redB6SpVerifyMismatch.Threw) Err='$($redB6SpVerifyMismatch.ErrorMessage)'"

$redB7SpErrorLeak = Invoke-AB730Probe -Target 'SharePoint' -Mode 'SharePointErrorLeakTest'
Assert-True 'RED-B7-1: Invoke-SeedSharePoint.ps1 transport catch sanitizes Bearer/tempauth/downloadUrl/client_secret while preserving stage & code' {
    $combined = "$($redB7SpErrorLeak.ErrorMessage)`n$($redB7SpErrorLeak.Stdout)"
    $redB7SpErrorLeak.Threw -and
    ($combined -notmatch 'eyJSECRET_JWT') -and
    ($combined -notmatch 'SECRET_CAP_TOKEN') -and
    ($combined -notmatch 'SECRET_DL') -and
    ($combined -notmatch 'stub-secret-value-not-real') -and
    ($combined -match '409|nameAlreadyExists')
} "expected SharePoint error path to redact secrets while keeping status/code, got Err='$($redB7SpErrorLeak.ErrorMessage)' Stdout='$($redB7SpErrorLeak.Stdout)'"

$redB7TeamsErrorLeak = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsErrorLeakTest'
Assert-True 'RED-B7-2: Invoke-SeedTeamsChannel.ps1 transport catch sanitizes Bearer/tempauth/downloadUrl/client_secret while preserving stage & code' {
    $combined = "$($redB7TeamsErrorLeak.ErrorMessage)`n$($redB7TeamsErrorLeak.Stdout)"
    $redB7TeamsErrorLeak.Threw -and
    ($combined -notmatch 'eyJSECRET_JWT') -and
    ($combined -notmatch 'SECRET_CAP_TOKEN') -and
    ($combined -notmatch 'SECRET_DL') -and
    ($combined -notmatch 'stub-secret-value-not-real') -and
    ($combined -match '403|Authorization_RequestDenied')
} "expected Teams error path to redact secrets while keeping status/code, got Err='$($redB7TeamsErrorLeak.ErrorMessage)' Stdout='$($redB7TeamsErrorLeak.Stdout)'"

# 10. Council targeted recheck RED/GREEN tests (B1, B2, B6, B7 + SDK MaxRetry=0 safeguard)
$global:__AB730DependencyLoadedBeforeGate = 0
$redB1DepBeforeGate = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{
    CustomSetup = {
        param($envInfo)
        $isolatedEngineDir = Join-Path $envInfo.RootDir 'seed-data\engine-isolated'
        Copy-Item -LiteralPath $envInfo.EngineDir -Destination $isolatedEngineDir -Recurse -Force
        foreach ($helperFile in @('Seed-Idempotency.ps1', 'Seed-GraphRead.ps1', 'Seed-OneDriveOwnership.ps1')) {
            $hp = Join-Path $isolatedEngineDir $helperFile
            $origContent = Get-Content -LiteralPath $hp -Raw -Encoding UTF8
            Set-Content -LiteralPath $hp -Value ("`$global:__AB730DependencyLoadedBeforeGate++`r`n" + $origContent) -Encoding UTF8
        }
        # Point Invoke-AB730.ps1 repoRoot resolution inside isolated tree by placing engine at seed-data\engine of a sub-root
        $subRoot = Join-Path $envInfo.RootDir 'b1-subroot'
        $subSeed = Join-Path $subRoot 'seed-data'
        $subEngine = Join-Path $subSeed 'engine'
        New-Item -ItemType Directory -Path $subEngine -Force | Out-Null
        Copy-Item -Path (Join-Path $isolatedEngineDir '*') -Destination $subEngine -Recurse -Force
        $Overrides['CustomEngineDir'] = $subEngine
    }
}
Assert-True 'RED-B1-7: Invoke-AB730.ps1 does not dot-source or execute any helper script before codeHashes verification completes (dependency-before-gate=0, token/Graph=0)' {
    $redB1DepBeforeGate.Threw -and
    ($global:__AB730DependencyLoadedBeforeGate -eq 0) -and
    ($redB1DepBeforeGate.ErrorMessage -match 'codeHashes|Seed-Idempotency|Seed-GraphRead|Seed-OneDriveOwnership|hash') -and
    ($redB1DepBeforeGate.TokenRequests.Count -eq 0) -and
    ($redB1DepBeforeGate.Requests.Count -eq 0)
} "expected dependency-before-gate=0 and hash rejection, got DepLoaded=$global:__AB730DependencyLoadedBeforeGate Threw=$($redB1DepBeforeGate.Threw) Err='$($redB1DepBeforeGate.ErrorMessage)'"
Remove-Item Variable:\global:__AB730DependencyLoadedBeforeGate -ErrorAction SilentlyContinue

$redB2OpAnchorMismatch = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsOperationAnchorMismatch'
$postTeamWritesOp = @($redB2OpAnchorMismatch.Requests | Where-Object { $_.Method -ne 'GET' -and $_.Uri -notmatch '/v1\.0/teams$' })
Assert-True 'RED-B2-3: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) aborts when operation targetResourceId disagrees with immutable POST 202 Team ID anchor (0 channel/message/member writes)' {
    $redB2OpAnchorMismatch.Threw -and
    ($redB2OpAnchorMismatch.ErrorMessage -match 'anchor|targetResourceId|mismatch|team-202-hdr|team-mutated-op') -and
    ($postTeamWritesOp.Count -eq 0)
} "expected operation targetResourceId mismatch to abort before channel/message/member writes, got Threw=$($redB2OpAnchorMismatch.Threw) PostTeamWrites=$($postTeamWritesOp.Count) Err='$($redB2OpAnchorMismatch.ErrorMessage)'"

$redB2GetIdAnchorMismatch = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsGetIdAnchorMismatch'
$postTeamWritesGet = @($redB2GetIdAnchorMismatch.Requests | Where-Object { $_.Method -ne 'GET' -and $_.Uri -notmatch '/v1\.0/teams$' })
Assert-True 'RED-B2-4: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) aborts when GET /teams/{id} body id disagrees with immutable Team ID anchor (0 channel/message/member writes)' {
    $redB2GetIdAnchorMismatch.Threw -and
    ($redB2GetIdAnchorMismatch.ErrorMessage -match 'anchor|mismatch|team-202-hdr|team-mutated-get') -and
    ($postTeamWritesGet.Count -eq 0)
} "expected GET /teams/{id} body id mismatch to abort before channel/message/member writes, got Threw=$($redB2GetIdAnchorMismatch.Threw) PostTeamWrites=$($postTeamWritesGet.Count) Err='$($redB2GetIdAnchorMismatch.ErrorMessage)'"

$redB6TeamsForeignMember = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsMemberForeignReplacement'
Assert-True 'RED-B6-5: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) aborts when GET /teams/{id}/members has equal count (8) and valid Admin owner but replaces a business member with a foreign identity' {
    $redB6TeamsForeignMember.Threw -and
    ($redB6TeamsForeignMember.ErrorMessage -match 'foreign|unexpected|member|identity|mismatch|verification')
} "expected foreign replacement in Teams conversation members to fail closed, got Threw=$($redB6TeamsForeignMember.Threw) Err='$($redB6TeamsForeignMember.ErrorMessage)'"

$redB6TeamsBackingGroup = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsBackingGroupForeignMember'
Assert-True 'RED-B6-6: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) verifies Teams backing group owners (1 Admin) and members (8 exact roles) and aborts on foreign/extra identity' {
    $redB6TeamsBackingGroup.Threw -and
    ($redB6TeamsBackingGroup.ErrorMessage -match 'group|member|foreign|unexpected|identity|mismatch|verification')
} "expected foreign replacement in Teams backing M365 group members to fail closed, got Threw=$($redB6TeamsBackingGroup.Threw) Err='$($redB6TeamsBackingGroup.ErrorMessage)'"

$redB6SpExtraMember = Invoke-AB730Probe -Target 'SharePoint' -Mode 'SharePointExtraMember'
Assert-True 'RED-B6-7: Invoke-SeedSharePoint.ps1 (-FreshOnly) aborts when GET /groups/{id}/members contains an extra undeclared member' {
    $redB6SpExtraMember.Threw -and
    ($redB6SpExtraMember.ErrorMessage -match 'extra|unexpected|member|count|mismatch|verification')
} "expected extra member in SharePoint group to fail closed, got Threw=$($redB6SpExtraMember.Threw) Err='$($redB6SpExtraMember.ErrorMessage)'"

$redB7EmailsErrorLeak = Invoke-AB730Probe -Target 'Emails' -Mode 'EmailsErrorLeakTest'
Assert-True 'RED-B7-3: Invoke-SeedEmails.ps1 transport catch sanitizes Bearer/tempauth/downloadUrl/client_secret while preserving stage & code' {
    $combined = "$($redB7EmailsErrorLeak.ErrorMessage)`n$($redB7EmailsErrorLeak.Stdout)"
    $redB7EmailsErrorLeak.Threw -and
    ($combined -notmatch 'eyJSECRET_JWT') -and
    ($combined -notmatch 'SECRET_CAP_TOKEN') -and
    ($combined -notmatch 'SECRET_DL') -and
    ($combined -notmatch 'stub-secret-value-not-real') -and
    ($combined -match '403|ErrorAccessDenied')
} "expected Emails error path to redact secrets while keeping status/code, got Err='$($redB7EmailsErrorLeak.ErrorMessage)' Stdout='$($redB7EmailsErrorLeak.Stdout)'"

$redB7PreflightErrorLeak = Invoke-AB730Probe -Target 'Preflight' -Mode 'PreflightErrorLeakTest'
Assert-True 'RED-B7-4: Invoke-SeedPreflight.ps1 transport catch sanitizes Bearer/tempauth/downloadUrl/client_secret while preserving stage & code' {
    $combined = "$($redB7PreflightErrorLeak.ErrorMessage)`n$($redB7PreflightErrorLeak.Stdout)"
    $redB7PreflightErrorLeak.Threw -and
    ($combined -notmatch 'eyJSECRET_JWT') -and
    ($combined -notmatch 'SECRET_CAP_TOKEN') -and
    ($combined -notmatch 'SECRET_DL') -and
    ($combined -notmatch 'stub-secret-value-not-real') -and
    ($combined -match '403|Authorization_RequestDenied')
} "expected Preflight error path to redact secrets while keeping status/code, got Err='$($redB7PreflightErrorLeak.ErrorMessage)' Stdout='$($redB7PreflightErrorLeak.Stdout)'"

$redB7UploadErrorLeak = Invoke-AB730Probe -Target 'Upload' -Mode 'UploadErrorLeakTest'
Assert-True 'RED-B7-5: Invoke-UploadFiles.ps1 transport catch sanitizes Bearer/tempauth/downloadUrl/client_secret while preserving stage & code' {
    $combined = "$($redB7UploadErrorLeak.ErrorMessage)`n$($redB7UploadErrorLeak.Stdout)"
    $redB7UploadErrorLeak.Threw -and
    ($combined -notmatch 'eyJSECRET_JWT') -and
    ($combined -notmatch 'SECRET_CAP_TOKEN') -and
    ($combined -notmatch 'SECRET_DL') -and
    ($combined -notmatch 'stub-secret-value-not-real') -and
    ($combined -match '409|nameAlreadyExists')
} "expected Upload error path to redact secrets while keeping status/code, got Err='$($redB7UploadErrorLeak.ErrorMessage)' Stdout='$($redB7UploadErrorLeak.Stdout)'"

Assert-True 'RED-B7-7a: FreshOnly transport enforces SDK MaxRetry=0 before first POST/PUT write when context starts at MaxRetry=3' {
    (-not $runLive.Threw) -and
    ($runLive.MgMaxRetrySetCalls -gt 0) -and
    ($runLive.MgMaxRetryAtFirstWrite -eq 0) -and
    ($runLive.MgMaxRetry -eq 0)
} "expected MgMaxRetryAtFirstWrite=0 and MgMaxRetrySetCalls>0 in RunnerLive, got AtFirstWrite=$($runLive.MgMaxRetryAtFirstWrite) SetCalls=$($runLive.MgMaxRetrySetCalls)"

$redB7MgCtxEnforceFail = Invoke-AB730Probe -Target 'Upload' -Mode 'MgContextEnforceFail'
$mgEnforceWrites = @($redB7MgCtxEnforceFail.Requests | Where-Object { $_.Method -ne 'GET' })
Assert-True 'RED-B7-7b: FreshOnly transport aborts before any POST/PUT write (0 writes) when Set-MgRequestContext/Get-MgRequestContext fails to enforce MaxRetry=0' {
    $redB7MgCtxEnforceFail.Threw -and
    ($redB7MgCtxEnforceFail.ErrorMessage -match 'MaxRetry|write-retry|Set-MgRequestContext|Get-MgRequestContext') -and
    ($mgEnforceWrites.Count -eq 0)
} "expected MgContextEnforceFail to abort with 0 POST/PUT writes, got Threw=$($redB7MgCtxEnforceFail.Threw) Writes=$($mgEnforceWrites.Count) Err='$($redB7MgCtxEnforceFail.ErrorMessage)'"

# ───────────────────────────────────────────────────────
# 9. Bounded -Phases routing TDD (RED-PHASES-1..3)
# ───────────────────────────────────────────────────────
Write-Host "`n── AB-730 Bounded -Phases Routing Tests ──" -ForegroundColor Cyan

$redPhasesEmpty = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{ Phases = @() }
Assert-True 'RED-PHASES-1a: Invoke-AB730.ps1 rejects empty -Phases before any helper load or OAuth/Graph request' {
    $redPhasesEmpty.Threw -and
    ($redPhasesEmpty.ErrorMessage -notmatch 'parameter cannot be found') -and
    ($redPhasesEmpty.ErrorMessage -match 'Phases|Cannot validate argument') -and
    ($redPhasesEmpty.TokenRequests.Count -eq 0) -and
    ($redPhasesEmpty.Requests.Count -eq 0)
} "expected empty -Phases validation error with 0 token/Graph requests, got Threw=$($redPhasesEmpty.Threw) Token=$($redPhasesEmpty.TokenRequests.Count) Err='$($redPhasesEmpty.ErrorMessage)'"

$redPhasesInvalid = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{ Phases = @('UnknownSurface') }
Assert-True 'RED-PHASES-1b: Invoke-AB730.ps1 rejects unknown -Phases value before any helper load or OAuth/Graph request' {
    $redPhasesInvalid.Threw -and
    ($redPhasesInvalid.ErrorMessage -notmatch 'parameter cannot be found') -and
    ($redPhasesInvalid.ErrorMessage -match 'Phases|UnknownSurface|ValidateSet') -and
    ($redPhasesInvalid.TokenRequests.Count -eq 0) -and
    ($redPhasesInvalid.Requests.Count -eq 0)
} "expected unknown -Phases validation error with 0 token/Graph requests, got Threw=$($redPhasesInvalid.Threw) Token=$($redPhasesInvalid.TokenRequests.Count) Err='$($redPhasesInvalid.ErrorMessage)'"

$redPhasesMixedAll = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'Clean' -Overrides @{ Phases = @('All', 'Emails') }
Assert-True 'RED-PHASES-1c: Invoke-AB730.ps1 rejects mixing All with specific -Phases values before any helper load or OAuth/Graph request' {
    $redPhasesMixedAll.Threw -and
    ($redPhasesMixedAll.ErrorMessage -notmatch 'parameter cannot be found') -and
    ($redPhasesMixedAll.ErrorMessage -match 'cannot mix|All') -and
    ($redPhasesMixedAll.TokenRequests.Count -eq 0) -and
    ($redPhasesMixedAll.Requests.Count -eq 0)
} "expected mixed All+Emails -Phases validation error with 0 token/Graph requests, got Threw=$($redPhasesMixedAll.Threw) Token=$($redPhasesMixedAll.TokenRequests.Count) Err='$($redPhasesMixedAll.ErrorMessage)'"

$redPhasesRemaining3 = Invoke-AB730Probe -Target 'RunnerLive' -Mode 'ExistingOneDriveFolder' -Overrides @{ Phases = @('Emails', 'Teams', 'SharePoint') }
$odReqsRemaining3 = @($redPhasesRemaining3.Requests | Where-Object { $_.Uri -match '/users/[^/]+/drive/' })
$mailWritesRemaining3 = @($redPhasesRemaining3.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/sendMail$|/reply$' })
$teamWritesRemaining3 = @($redPhasesRemaining3.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/v1\.0/teams' })
$spWritesRemaining3   = @($redPhasesRemaining3.Requests | Where-Object { $_.Method -in @('POST', 'PUT') -and $_.Uri -match '/v1\.0/groups|/v1\.0/sites/' })
Assert-True 'RED-PHASES-2: Invoke-AB730.ps1 -Phases Emails,Teams,SharePoint -Live skips completed OneDrive (0 drive requests) and seeds remaining 3 fresh phases' {
    (-not $redPhasesRemaining3.Threw) -and
    ($odReqsRemaining3.Count -eq 0) -and
    ($mailWritesRemaining3.Count -eq 15) -and
    ($teamWritesRemaining3.Count -gt 0) -and
    ($spWritesRemaining3.Count -gt 0)
} "expected -Phases Emails,Teams,SharePoint -Live to succeed with 0 OneDrive requests and seed Emails/Teams/SharePoint, got Threw=$($redPhasesRemaining3.Threw) ODReqs=$($odReqsRemaining3.Count) MailWrites=$($mailWritesRemaining3.Count) Err='$($redPhasesRemaining3.ErrorMessage)'"

$redPhasesSingleEmails = Invoke-AB730Probe -Target 'RunnerLive' -Mode 'ExistingOneDriveFolder' -Overrides @{ Phases = @('Emails') }
$nonMailReqsSingle = @($redPhasesSingleEmails.Requests | Where-Object { $_.Uri -match '/drive/|/teams|/groups|/sites/' })
$mailWritesSingle  = @($redPhasesSingleEmails.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/sendMail$|/reply$' })
Assert-True 'RED-PHASES-3a: Invoke-AB730.ps1 -Phases Emails -Live executes only Emails preflight+writes with 0 OneDrive/Teams/SharePoint requests' {
    (-not $redPhasesSingleEmails.Threw) -and
    ($nonMailReqsSingle.Count -eq 0) -and
    ($mailWritesSingle.Count -eq 15)
} "expected -Phases Emails -Live to make 0 non-mail resource requests and 15 mail writes, got Threw=$($redPhasesSingleEmails.Threw) NonMailReqs=$($nonMailReqsSingle.Count) MailWrites=$($mailWritesSingle.Count) Err='$($redPhasesSingleEmails.ErrorMessage)'"

$redPhasesUnselectedTamper = Invoke-AB730Probe -Target 'RunnerLive' -Mode 'Clean' -Overrides @{
    Phases      = @('Emails')
    CustomSetup = {
        param($envInfo)
        $docPath = Join-Path $envInfo.DemoDir '06-Quality-Brief.docx'
        [System.IO.File]::WriteAllBytes($docPath, [System.Text.Encoding]::UTF8.GetBytes(('MUTATED-DOC-BYTES-' * 80)))
        $newHash = (Get-FileHash -LiteralPath $docPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $fmPath = Join-Path $envInfo.ScenarioDir 'files-manifest.json'
        $fmObj = Get-Content -LiteralPath $fmPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($fEntry in @($fmObj.files)) {
            if ($fEntry.localName -eq '06-Quality-Brief.docx') { $fEntry.sha256 = $newHash }
        }
        $fmObj.assetSHA256.'06-Quality-Brief.docx' = $newHash
        $fmObj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $fmPath -Encoding UTF8
    }
}
Assert-True 'RED-PHASES-3b: Invoke-AB730.ps1 -Phases Emails still enforces full 12-artifact closure on unselected OneDrive file before OAuth/Graph' {
    $redPhasesUnselectedTamper.Threw -and
    ($redPhasesUnselectedTamper.ErrorMessage -match 'hash|mismatch|artifactHashes|06-Quality-Brief\.docx') -and
    ($redPhasesUnselectedTamper.TokenRequests.Count -eq 0) -and
    ($redPhasesUnselectedTamper.Requests.Count -eq 0)
} "expected unselected artifact tamper to abort before OAuth/Graph even under -Phases Emails, got Threw=$($redPhasesUnselectedTamper.Threw) Token=$($redPhasesUnselectedTamper.TokenRequests.Count) Err='$($redPhasesUnselectedTamper.ErrorMessage)'"

# ───────────────────────────────────────────────────────
# 10. Final 4 Council Recheck Gaps (B2 blank/missing displayName, B7 OAuth sanitization)
# ───────────────────────────────────────────────────────
$redB2BlankDisp = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsGetBlankDisplayName'
$postTeamWritesBlankDisp = @($redB2BlankDisp.Requests | Where-Object { $_.Method -ne 'GET' -and $_.Uri -notmatch '/v1\.0/teams$' })
$redB2MissingDisp = Invoke-AB730Probe -Target 'Teams' -Mode 'TeamsGetMissingDisplayName'
$postTeamWritesMissingDisp = @($redB2MissingDisp.Requests | Where-Object { $_.Method -ne 'GET' -and $_.Uri -notmatch '/v1\.0/teams$' })
Assert-True 'RED-B2-5: Invoke-SeedTeamsChannel.ps1 (-FreshOnly) aborts when GET /teams/{id} returns blank or missing displayName (0 channel/message/member writes)' {
    $redB2BlankDisp.Threw -and
    ($redB2BlankDisp.ErrorMessage -match 'displayName|mismatch') -and
    ($postTeamWritesBlankDisp.Count -eq 0) -and
    $redB2MissingDisp.Threw -and
    ($redB2MissingDisp.ErrorMessage -match 'displayName|mismatch') -and
    ($postTeamWritesMissingDisp.Count -eq 0)
} "expected blank/missing displayName on GET /teams/{id} to abort before channel/message/member writes, got BlankThrew=$($redB2BlankDisp.Threw) BlankWrites=$($postTeamWritesBlankDisp.Count) MissingThrew=$($redB2MissingDisp.Threw) MissingWrites=$($postTeamWritesMissingDisp.Count)"

$redB7OAuthLeak = Invoke-AB730Probe -Target 'RunnerPreflight' -Mode 'OAuthErrorLeakTest'
Assert-True 'RED-B7-9: Invoke-AB730.ps1 OAuth POST catch sanitizes Bearer/tempauth/downloadUrl/client_secret while preserving typed StatusCode (401), OAuth stage, and detaching Response/RequestMessage/Content/Headers/TargetObject/Data' {
    $combinedOAuth = "$($redB7OAuthLeak.ErrorMessage)`n$($redB7OAuthLeak.Stdout)"
    $erObj = $redB7OAuthLeak.ErrorRecord
    $exObj = if ($erObj) { $erObj.Exception } else { $null }
    $hasRespProp = ($null -ne $exObj) -and ($exObj.PSObject.Properties.Name -contains 'Response') -and ($null -ne $exObj.Response)
    $hasReqProp  = ($null -ne $exObj) -and ($exObj.PSObject.Properties.Name -contains 'RequestMessage') -and ($null -ne $exObj.RequestMessage)
    $hasCntProp  = ($null -ne $exObj) -and ($exObj.PSObject.Properties.Name -contains 'Content') -and ($null -ne $exObj.Content)
    $hasHdrProp  = ($null -ne $exObj) -and ($exObj.PSObject.Properties.Name -contains 'Headers') -and ($null -ne $exObj.Headers)
    $redB7OAuthLeak.Threw -and
    ($redB7OAuthLeak.TokenRequests.Count -eq 1) -and
    ($redB7OAuthLeak.Requests.Count -eq 0) -and
    ($redB7OAuthLeak.ErrorStatusCode -eq [System.Net.HttpStatusCode]::Unauthorized) -and
    ($null -ne $erObj) -and
    ($null -eq $erObj.TargetObject) -and
    ($null -eq $exObj.InnerException) -and
    (-not $hasRespProp) -and
    (-not $hasReqProp) -and
    (-not $hasCntProp) -and
    (-not $hasHdrProp) -and
    ($exObj.Data.Count -eq 0) -and
    ($combinedOAuth -notmatch 'eyJSECRET_JWT') -and
    ($combinedOAuth -notmatch 'SECRET_CAP_TOKEN') -and
    ($combinedOAuth -notmatch 'SECRET_DL') -and
    ($combinedOAuth -notmatch 'stub-secret-value-not-real') -and
    ($combinedOAuth -match 'OAuth|401|invalid_client')
} "expected OAuth POST error path to sanitize secrets, preserve StatusCode=401, and detach Response/RequestMessage/Content/Headers/TargetObject/Data, got Status=$($redB7OAuthLeak.ErrorStatusCode) HasResp=$(if ($redB7OAuthLeak.ErrorRecord -and $redB7OAuthLeak.ErrorRecord.Exception.PSObject.Properties.Name -contains 'Response' -and $null -ne $redB7OAuthLeak.ErrorRecord.Exception.Response) { $true } else { $false }) Err='$($redB7OAuthLeak.ErrorMessage)'"

$redB7Mail429RetryAfter = Invoke-AB730Probe -Target 'Emails' -Mode 'MailRead429RetryAfter'
Assert-True 'RED-B7-10: Invoke-SeedEmails.ps1 bounded GET honors Retry-After delay scalar (7s) extracted into Data without retaining Response/Headers' {
    (-not $redB7Mail429RetryAfter.Threw) -and
    ($redB7Mail429RetryAfter.SleepDelays -contains 7)
} "expected MailRead429RetryAfter to succeed and sleep for Retry-After=7s, got Threw=$($redB7Mail429RetryAfter.Threw) SleepDelays=[$($redB7Mail429RetryAfter.SleepDelays -join ',')] Err='$($redB7Mail429RetryAfter.ErrorMessage)'"



