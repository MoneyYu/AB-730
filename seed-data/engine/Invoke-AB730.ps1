<#
.SYNOPSIS
    Fail-closed, fresh-target-only orchestrator for the AB-730 M365 demo scenario.

.DESCRIPTION
    Enforces the AB-730 execution contract before and during seeding:
      - Mandatory -ConfigPath, -ManifestPath, and -ScenarioDir (never discovers or copies root config.json).
      - Default mode is preflight-only (read-only). Passing -Live is required to execute writes.
      - Fails closed if ConfigPath or ManifestPath is tracked by git or not git-ignored.
      - Validates exact tenantId, clientId, operator Admin UPN, 8 role bindings (1 Admin + 7 business roles),
        target resource names, counts, and SHA-256 hashes of all local DEMO-FILE artifacts before any Graph call.
      - Validates filesSourceDir resolves relative to ConfigPath and stays strictly inside ScenarioDir.
      - Acquires an app-only Graph token and verifies JWT claims (tid, appid/azp, aud, exp, required app roles)
        before any resource query.
      - Runs Invoke-SeedPreflight.ps1 with -FreshOnly (rejecting any existing OneDrive folder, Team, SharePoint
        group alias/displayName, or Admin Inbox thread subject across up to 20 pages).
      - When -Live is specified, executes OneDrive, Outlook, Teams, and SharePoint seeding phases in -FreshOnly
        mode with zero automatic write retries and zero DELETE/PATCH recovery calls.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [Parameter(Mandatory)][string]$ManifestPath,
    [Parameter(Mandatory)][string]$ScenarioDir,
    [switch]$Live,
    [ValidateSet('All', 'OneDrive', 'Emails', 'Teams', 'SharePoint')]
    [string[]]$Phases = @('All')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$rawPhases = @($Phases | Where-Object { -not [string]::IsNullOrWhiteSpace("$_") } | ForEach-Object { "$_".Trim() })
if ($rawPhases.Count -eq 0 -or $rawPhases.Count -ne @($Phases).Count) {
    throw "Invalid -Phases parameter: must specify at least one non-empty phase ('All', 'OneDrive', 'Emails', 'Teams', 'SharePoint')."
}
$hasAllPhase = [bool]($rawPhases | Where-Object { [string]::Equals($_, 'All', [System.StringComparison]::OrdinalIgnoreCase) })
if ($hasAllPhase -and $rawPhases.Count -gt 1) {
    throw "Invalid -Phases parameter: cannot mix 'All' with specific phase names ($($rawPhases -join ', '))."
}
$canonicalPhaseOrder = @('OneDrive', 'Emails', 'Teams', 'SharePoint')
$selectedPhases = if ($hasAllPhase) {
    $canonicalPhaseOrder
} else {
    @($canonicalPhaseOrder | Where-Object {
        $canon = $_
        [bool]($rawPhases | Where-Object { [string]::Equals($_, $canon, [System.StringComparison]::OrdinalIgnoreCase) })
    })
}

function Test-AB730GitIgnoreAndUntracked {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$Label
    )
    $parentDir = Split-Path -Parent $FilePath
    $fileName = Split-Path -Leaf $FilePath

    $trackedOut = & git -C $parentDir ls-files --error-unmatch -- $fileName 2>$null
    if ($LASTEXITCODE -eq 0) {
        throw "Security boundary violation: $Label '$FilePath' is tracked by git. Refusing to run."
    }

    $ignoredOut = & git -C $parentDir check-ignore -q -- $fileName 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "Security boundary violation: $Label '$FilePath' is not git-ignored. Refusing to run."
    }
}

function ConvertFrom-AB730JwtPayload {
    param([Parameter(Mandatory)][string]$Token)
    $parts = $Token -split '\.'
    if ($parts.Count -lt 2 -or [string]::IsNullOrWhiteSpace($parts[1])) {
        throw "Acquired access token is not a valid JWT."
    }
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
        1 { throw "Acquired access token has a malformed base64url payload." }
    }
    $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
    return ($json | ConvertFrom-Json)
}

function Invoke-AB730GraphGet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Token
    )
    $headers = @{
        Authorization  = "Bearer $Token"
        'Content-Type' = 'application/json'
    }
    try {
        return (Invoke-RestMethod -Method GET -Uri $Uri -Headers $headers)
    } catch {
        $safe = Format-SeedSafeError -ErrorRecord $_ -Operation "GET $Uri"
        throw $safe
    }
}

# ───────────────────────────────────────────────────────
# 1. File existence & scenario directory hygiene
# ───────────────────────────────────────────────────────
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "ConfigPath not found: '$ConfigPath'."
}
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "ManifestPath not found: '$ManifestPath'."
}
if (-not (Test-Path -LiteralPath $ScenarioDir -PathType Container)) {
    throw "ScenarioDir not found: '$ScenarioDir'."
}

$resolvedConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$resolvedManifestPath = (Resolve-Path -LiteralPath $ManifestPath).Path
$resolvedScenarioDir = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $ScenarioDir).Path).TrimEnd('\', '/')

$scenarioLocalConfig = Join-Path $resolvedScenarioDir 'config.json'
if ([string]::Equals($resolvedConfigPath, [System.IO.Path]::GetFullPath($scenarioLocalConfig), [System.StringComparison]::OrdinalIgnoreCase) -or
    (Test-Path -LiteralPath $scenarioLocalConfig)) {
    throw "Security boundary violation: config.json must not exist inside ScenarioDir '$resolvedScenarioDir'. Keep private config at repo root only."
}

Test-AB730GitIgnoreAndUntracked -FilePath $resolvedConfigPath -Label 'private config'
Test-AB730GitIgnoreAndUntracked -FilePath $resolvedManifestPath -Label 'deployment manifest'

# ───────────────────────────────────────────────────────
# 2. Parse & bind ConfigPath + ManifestPath + Scenario JSONs
# ───────────────────────────────────────────────────────
$config = Get-Content -LiteralPath $resolvedConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$manifest = Get-Content -LiteralPath $resolvedManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not ($manifest.PSObject.Properties.Name -contains 'writeMode') -or
    $manifest.writeMode -ne 'additive-fresh-targets-only' -or
    -not [bool]$manifest.preflightMustFindNewTargets -or
    [bool]$manifest.deletionRenameAdoptionRepairAllowed -or
    [bool]$manifest.automaticWriteRetryAllowed) {
    throw "Manifest policy violation: manifest must enforce writeMode='additive-fresh-targets-only', preflightMustFindNewTargets=true, deletionRenameAdoptionRepairAllowed=false, automaticWriteRetryAllowed=false."
}
if ($Live -and (-not ($manifest.PSObject.Properties.Name -contains 'tenantWriteApproved') -or -not [bool]$manifest.tenantWriteApproved)) {
    throw "Manifest policy violation: -Live requires manifest.tenantWriteApproved=true."
}

$expectedTenantId = "$($manifest.target.tenantId)"
$expectedClientId = "$($manifest.target.clientId)"
if ([string]::IsNullOrWhiteSpace($expectedTenantId) -or [string]::IsNullOrWhiteSpace($expectedClientId)) {
    throw "Manifest is missing target.tenantId or target.clientId."
}
if (-not [string]::Equals("$($config.tenantId)", $expectedTenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Manifest binding mismatch: config.tenantId '$($config.tenantId)' does not match manifest target.tenantId '$expectedTenantId'."
}
if (-not [string]::Equals("$($config.clientId)", $expectedClientId, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Manifest binding mismatch: config.clientId '$($config.clientId)' does not match manifest target.clientId '$expectedClientId'."
}

$expectedRoles = @('Admin', 'Sponsor', 'Planner', 'Supply', 'Quality', 'Operations', 'Engineering', 'Analytics')
$manifestRoleNames = @($manifest.roles.PSObject.Properties | ForEach-Object { $_.Name })
$configRoleNames = @($config.roles.PSObject.Properties | ForEach-Object { $_.Name })

if (@(Compare-Object -ReferenceObject ($expectedRoles | Sort-Object) -DifferenceObject ($manifestRoleNames | Sort-Object)).Count -ne 0) {
    throw "Manifest role set mismatch: expected exactly [$($expectedRoles -join ', ')], found [$($manifestRoleNames -join ', ')]."
}
if (@(Compare-Object -ReferenceObject ($expectedRoles | Sort-Object) -DifferenceObject ($configRoleNames | Sort-Object)).Count -ne 0) {
    throw "Config role set mismatch: expected exactly [$($expectedRoles -join ', ')], found [$($configRoleNames -join ', ')]."
}

$manifestAdminUpn = "$($manifest.roles.Admin.upn)"
if (-not [string]::Equals("$($config.adminUpn)", $manifestAdminUpn, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Manifest binding mismatch: config.adminUpn '$($config.adminUpn)' does not match manifest Admin UPN '$manifestAdminUpn'."
}
if (($config.PSObject.Properties.Name -contains 'demoUserUpn') -and $config.demoUserUpn -and
    -not [string]::Equals("$($config.demoUserUpn)", $manifestAdminUpn, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Manifest binding mismatch: config.demoUserUpn '$($config.demoUserUpn)' does not match manifest Admin UPN '$manifestAdminUpn'."
}

$seenUpns = @{}
foreach ($roleKey in $expectedRoles) {
    $mUpn = "$($manifest.roles.$roleKey.upn)"
    $cUpn = "$($config.roles.$roleKey.upn)"
    if ([string]::IsNullOrWhiteSpace($mUpn) -or [string]::IsNullOrWhiteSpace($cUpn)) {
        throw "Role '$roleKey' has an empty UPN in config or manifest."
    }
    if (-not [string]::Equals($cUpn, $mUpn, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Manifest binding mismatch for role '$roleKey': config UPN '$cUpn' != manifest UPN '$mUpn'."
    }
    $normUpn = $cUpn.ToLowerInvariant()
    if ($seenUpns.ContainsKey($normUpn)) {
        throw "Role UPN collision: role '$roleKey' shares UPN '$cUpn' with role '$($seenUpns[$normUpn])'."
    }
    $seenUpns[$normUpn] = $roleKey
}

# Validate filesSourceDir relative to ConfigPath stays inside ScenarioDir
$configDir = [System.IO.Path]::GetFullPath((Split-Path -Parent $resolvedConfigPath)).TrimEnd('\', '/')
if ([string]::IsNullOrWhiteSpace("$($config.filesSourceDir)") -or [System.IO.Path]::IsPathRooted("$($config.filesSourceDir)")) {
    throw "Invalid filesSourceDir '$($config.filesSourceDir)': must be a relative path resolving inside ScenarioDir."
}
$resolvedFilesSourceDir = [System.IO.Path]::GetFullPath((Join-Path $configDir $config.filesSourceDir)).TrimEnd('\', '/')
$expectedScenarioPrefix = "$resolvedScenarioDir$([System.IO.Path]::DirectorySeparatorChar)"
if (-not $resolvedFilesSourceDir.StartsWith($expectedScenarioPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Security boundary violation: filesSourceDir '$($config.filesSourceDir)' resolved to '$resolvedFilesSourceDir', outside ScenarioDir '$resolvedScenarioDir'."
}

# Validate scenario JSONs and DEMO-FILE SHA-256 hashes before any Graph call
$filesManifestPath = Join-Path $resolvedScenarioDir 'files-manifest.json'
$emailsPath = Join-Path $resolvedScenarioDir 'emails.json'
$teamsPath = Join-Path $resolvedScenarioDir 'teams-messages.json'
$spPath = Join-Path $resolvedScenarioDir 'sharepoint-sites.json'

foreach ($reqPath in @($filesManifestPath, $emailsPath, $teamsPath, $spPath)) {
    if (-not (Test-Path -LiteralPath $reqPath -PathType Leaf)) {
        throw "Required scenario file missing: '$reqPath'."
    }
}

$filesManifest = Get-Content -LiteralPath $filesManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$emailsData = Get-Content -LiteralPath $emailsPath -Raw -Encoding UTF8 | ConvertFrom-Json
$teamsData = Get-Content -LiteralPath $teamsPath -Raw -Encoding UTF8 | ConvertFrom-Json
$spData = Get-Content -LiteralPath $spPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Mandatory canonical dataSpec (ab730-data-spec.json) hash & structure binding
if (-not ($manifest.PSObject.Properties.Name -contains 'dataSpecSha256') -or
    [string]::IsNullOrWhiteSpace("$($manifest.dataSpecSha256)") -or
    "$($manifest.dataSpecSha256)".Trim() -notmatch '^[0-9a-fA-F]{64}$') {
    throw "Manifest contract violation: deployment manifest must declare a valid 64-hex 'dataSpecSha256'."
}
$manifestDir = Split-Path -Parent $resolvedManifestPath
$candidateDataSpecPaths = @()
if (($manifest.PSObject.Properties.Name -contains 'dataSpecPath') -and -not [string]::IsNullOrWhiteSpace("$($manifest.dataSpecPath)")) {
    $rawSpecPath = "$($manifest.dataSpecPath)".Trim()
    if ([System.IO.Path]::IsPathRooted($rawSpecPath)) {
        $candidateDataSpecPaths += $rawSpecPath
    } else {
        $candidateDataSpecPaths += (Join-Path $manifestDir $rawSpecPath)
        $candidateDataSpecPaths += (Join-Path $configDir $rawSpecPath)
    }
}
$candidateDataSpecPaths += (Join-Path $manifestDir 'files\ab730-data-spec.json')
$candidateDataSpecPaths += (Join-Path $manifestDir 'ab730-data-spec.json')
$candidateDataSpecPaths += (Join-Path $resolvedScenarioDir 'ab730-data-spec.json')

$resolvedDataSpecPath = $null
foreach ($cand in $candidateDataSpecPaths) {
    if (Test-Path -LiteralPath $cand -PathType Leaf) {
        $resolvedDataSpecPath = (Resolve-Path -LiteralPath $cand).Path
        break
    }
}
if (-not $resolvedDataSpecPath) {
    throw "Required canonical data spec 'ab730-data-spec.json' not found."
}
$actualDataSpecHash = (Get-FileHash -LiteralPath $resolvedDataSpecPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualDataSpecHash -ne "$($manifest.dataSpecSha256)".Trim().ToLowerInvariant()) {
    throw "Manifest binding mismatch: ab730-data-spec.json SHA-256 '$actualDataSpecHash' != manifest dataSpecSha256 '$($manifest.dataSpecSha256)'."
}
$dataSpec = Get-Content -LiteralPath $resolvedDataSpecPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Mandatory external artifactHashes (8 Office files + 4 scenario JSON files = 12 keys)
if (-not ($manifest.PSObject.Properties.Name -contains 'artifactHashes') -or $null -eq $manifest.artifactHashes) {
    throw "Manifest contract violation: deployment manifest must declare mandatory 'artifactHashes'."
}
$requiredScenarioJsons = @('emails.json', 'teams-messages.json', 'files-manifest.json', 'sharepoint-sites.json')
$requiredOfficeFiles = @($dataSpec.officeFiles | ForEach-Object { "$($_.name)" })
$artifactHashProps = @($manifest.artifactHashes.PSObject.Properties)
if ($artifactHashProps.Count -ne ($requiredScenarioJsons.Count + $requiredOfficeFiles.Count)) {
    throw "Manifest contract violation: artifactHashes must contain exactly $($requiredScenarioJsons.Count + $requiredOfficeFiles.Count) entries (8 Office files + 4 scenario JSONs), found $($artifactHashProps.Count)."
}
foreach ($jsonRel in $requiredScenarioJsons) {
    $jProp = $manifest.artifactHashes.PSObject.Properties[$jsonRel]
    if (-not $jProp -or [string]::IsNullOrWhiteSpace("$($jProp.Value)") -or "$($jProp.Value)".Trim() -notmatch '^[0-9a-fA-F]{64}$') {
        throw "Manifest contract violation: artifactHashes is missing a valid SHA-256 entry for '$jsonRel'."
    }
    $jFull = Join-Path $resolvedScenarioDir $jsonRel
    $jHash = (Get-FileHash -LiteralPath $jFull -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($jHash -ne "$($jProp.Value)".Trim().ToLowerInvariant()) {
        throw "Scenario JSON SHA-256 mismatch against deployment manifest for '$jsonRel': expected '$($jProp.Value)', actual '$jHash'."
    }
}

# Mandatory codeHashes (10 executed engine + runner scripts)
if (-not ($manifest.PSObject.Properties.Name -contains 'codeHashes') -or $null -eq $manifest.codeHashes) {
    throw "Manifest contract violation: deployment manifest must declare mandatory 'codeHashes' for all executed engine and runner scripts."
}
$requiredCodeKeys = @(
    'engine/Invoke-AB730.ps1',
    'engine/Invoke-SeedPreflight.ps1',
    'engine/Invoke-UploadFiles.ps1',
    'engine/Invoke-SeedEmails.ps1',
    'engine/Invoke-SeedTeamsChannel.ps1',
    'engine/Invoke-SeedSharePoint.ps1',
    'engine/Seed-Idempotency.ps1',
    'engine/Seed-GraphRead.ps1',
    'engine/Seed-OneDriveOwnership.ps1',
    'scenarios/ab730-automotive-20261015/run.ps1'
)
$codeHashProps = @($manifest.codeHashes.PSObject.Properties)
if ($codeHashProps.Count -ne $requiredCodeKeys.Count) {
    throw "Manifest contract violation: codeHashes must contain exactly $($requiredCodeKeys.Count) entries, found $($codeHashProps.Count)."
}
foreach ($codeKey in $requiredCodeKeys) {
    $cProp = $manifest.codeHashes.PSObject.Properties[$codeKey]
    if (-not $cProp) {
        $cProp = $manifest.codeHashes.PSObject.Properties["seed-data/$codeKey"]
    }
    if (-not $cProp -or [string]::IsNullOrWhiteSpace("$($cProp.Value)") -or "$($cProp.Value)".Trim() -notmatch '^[0-9a-fA-F]{64}$') {
        throw "Manifest contract violation: codeHashes is missing a valid SHA-256 entry for '$codeKey'."
    }
    $codeFilePath = if ($codeKey -like 'engine/*') {
        Join-Path $PSScriptRoot ($codeKey.Substring('engine/'.Length))
    } else {
        $scenRun = Join-Path $resolvedScenarioDir 'run.ps1'
        if (Test-Path -LiteralPath $scenRun -PathType Leaf) { $scenRun } else { Join-Path $PSScriptRoot 'Invoke-AB730ScenarioRun.template.ps1' }
    }
    if (-not (Test-Path -LiteralPath $codeFilePath -PathType Leaf)) {
        throw "Required code dependency missing for hash verification: '$codeFilePath' ($codeKey)."
    }
    $actualCodeHash = (Get-FileHash -LiteralPath $codeFilePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualCodeHash -ne "$($cProp.Value)".Trim().ToLowerInvariant()) {
        throw "Code integrity violation: SHA-256 mismatch for '$codeKey' ('$codeFilePath')."
    }
}

# OneDrive manifest & artifact hash checks
if ("$($filesManifest.targetFolder)" -cne "$($manifest.resources.oneDrive.folder)" -or
    "$($filesManifest.targetFolder)" -cne "$($dataSpec.resources.oneDriveFolder)") {
    throw "Target mismatch: files-manifest targetFolder '$($filesManifest.targetFolder)' != manifest/dataSpec '$($manifest.resources.oneDrive.folder)'."
}
if ("$($filesManifest.uploadToRole)" -cne "$($manifest.resources.oneDrive.ownerRole)" -or "$($filesManifest.uploadToRole)" -cne 'Admin') {
    throw "Target mismatch: files-manifest uploadToRole '$($filesManifest.uploadToRole)' must be 'Admin'."
}
$declaredFiles = @($filesManifest.files)
if ($declaredFiles.Count -ne [int]$manifest.resources.oneDrive.officeFileCount -or
    $declaredFiles.Count -ne $requiredOfficeFiles.Count) {
    throw "Count mismatch: files-manifest declares $($declaredFiles.Count) files, expected $($manifest.resources.oneDrive.officeFileCount)."
}

if (-not ($filesManifest.PSObject.Properties.Name -contains 'assetSHA256') -or $null -eq $filesManifest.assetSHA256) {
    throw "Manifest contract violation: files-manifest.json must declare 'assetSHA256' filename-to-SHA256 map."
}
$assetShaProps = @($filesManifest.assetSHA256.PSObject.Properties)
if ($assetShaProps.Count -ne [int]$manifest.resources.oneDrive.officeFileCount) {
    throw "Count mismatch: files-manifest.json assetSHA256 has $($assetShaProps.Count) entries, expected $($manifest.resources.oneDrive.officeFileCount)."
}

for ($fi = 0; $fi -lt $requiredOfficeFiles.Count; $fi++) {
    $expectedFileName = $requiredOfficeFiles[$fi]
    $f = $declaredFiles[$fi]
    if ("$($f.localName)" -cne $expectedFileName) {
        throw "OneDrive file list mismatch at index ${fi}: expected '$expectedFileName', found '$($f.localName)'."
    }
    $localFile = Join-Path $resolvedFilesSourceDir $f.localName
    if (-not (Test-Path -LiteralPath $localFile -PathType Leaf)) {
        throw "Missing DEMO-FILE artifact: '$localFile'."
    }
    $actualHash = (Get-FileHash -LiteralPath $localFile -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not ($f.PSObject.Properties.Name -contains 'sha256') -or [string]::IsNullOrWhiteSpace("$($f.sha256)")) {
        throw "Manifest contract violation: files-manifest.json entry '$($f.localName)' is missing 'sha256'."
    }
    if ($actualHash -ne "$($f.sha256)".Trim().ToLowerInvariant()) {
        throw "Artifact SHA-256 mismatch for '$($f.localName)': expected '$($f.sha256)', actual '$actualHash'."
    }
    $prop = $filesManifest.assetSHA256.PSObject.Properties[$f.localName]
    if (-not $prop -or $actualHash -ne "$($prop.Value)".Trim().ToLowerInvariant()) {
        throw "Artifact SHA-256 mismatch in files-manifest assetSHA256 for '$($f.localName)'."
    }
    $mProp = $manifest.artifactHashes.PSObject.Properties[$f.localName]
    if (-not $mProp) { $mProp = $manifest.artifactHashes.PSObject.Properties["DEMO-FILE/$($f.localName)"] }
    if (-not $mProp -or [string]::IsNullOrWhiteSpace("$($mProp.Value)") -or $actualHash -ne "$($mProp.Value)".Trim().ToLowerInvariant()) {
        throw "Artifact SHA-256 mismatch against deployment manifest for '$($f.localName)'."
    }
}

# Teams manifest & canonical dataSpec checks
if ("$($config.teamDisplayName)" -cne "$($manifest.resources.teams.displayName)" -or
    "$($config.teamDisplayName)" -cne "$($dataSpec.resources.teamDisplayName)") {
    throw "Target mismatch: config.teamDisplayName '$($config.teamDisplayName)' != manifest/dataSpec '$($manifest.resources.teams.displayName)'."
}
if (($teamsData.PSObject.Properties.Name -contains 'teamDisplayName') -and $teamsData.teamDisplayName -and
    "$($teamsData.teamDisplayName)" -cne "$($manifest.resources.teams.displayName)") {
    throw "Target mismatch: teams-messages.json teamDisplayName '$($teamsData.teamDisplayName)' != manifest '$($manifest.resources.teams.displayName)'."
}
$declaredChannels = @($teamsData.channels)
$specChannels = @($dataSpec.teamChannels)
if ($declaredChannels.Count -ne [int]$manifest.resources.teams.businessChannelCount -or
    $declaredChannels.Count -ne $specChannels.Count) {
    throw "Count mismatch: teams-messages.json declares $($declaredChannels.Count) channels, expected $($manifest.resources.teams.businessChannelCount)."
}
$totalTeamsMessages = 0
$nowUtc = [datetime]::UtcNow
for ($ci = 0; $ci -lt $specChannels.Count; $ci++) {
    $ch = $declaredChannels[$ci]
    $specCh = $specChannels[$ci]
    if ("$($ch.channelName)" -cne "$($specCh.name)") {
        throw "Teams channel name mismatch at index ${ci}: expected '$($specCh.name)', found '$($ch.channelName)'."
    }
    $chMsgs = @($ch.messages)
    $expectedTopAuthors = @($specCh.topAuthors | ForEach-Object { "$_" })
    if ($chMsgs.Count -ne $expectedTopAuthors.Count) {
        throw "Teams channel '$($ch.channelName)' top-level message count mismatch: expected $($expectedTopAuthors.Count), found $($chMsgs.Count)."
    }
    for ($mi = 0; $mi -lt $chMsgs.Count; $mi++) {
        $m = $chMsgs[$mi]
        $totalTeamsMessages++
        if ("$($m.fromRole)" -cne $expectedTopAuthors[$mi]) {
            throw "Teams message author mismatch in channel '$($ch.channelName)' post $($mi + 1): expected '$($expectedTopAuthors[$mi])', found '$($m.fromRole)'."
        }
        if ([string]::IsNullOrWhiteSpace("$($m.bodyHtml)")) {
            throw "Teams message bodyHtml cannot be empty in channel '$($ch.channelName)'."
        }
        $dt = [datetime]::Parse("$($m.createdDateTime)", [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
        if ($dt -ge $nowUtc) {
            throw "Teams contract violation: message createdDateTime '$($m.createdDateTime)' is not in the past."
        }
        $mReplies = @(if (($m.PSObject.Properties.Name -contains 'replies') -and $null -ne $m.replies) { $m.replies } else { @() })
        if ($mi -eq 0) {
            if (@($mReplies).Count -ne 1 -or "$($mReplies[0].fromRole)" -cne "$($specCh.firstPostReplyAuthor)") {
                throw "Teams reply author/count mismatch in channel '$($ch.channelName)': expected 1 reply from '$($specCh.firstPostReplyAuthor)'."
            }
            $totalTeamsMessages++
            if ([string]::IsNullOrWhiteSpace("$($mReplies[0].bodyHtml)")) {
                throw "Teams reply bodyHtml cannot be empty in channel '$($ch.channelName)'."
            }
            $rdt = [datetime]::Parse("$($mReplies[0].createdDateTime)", [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
            if ($rdt -ge $nowUtc) {
                throw "Teams contract violation: reply createdDateTime '$($mReplies[0].createdDateTime)' is not in the past."
            }
        } else {
            if (@($mReplies).Count -ne 0) {
                throw "Teams reply count mismatch in channel '$($ch.channelName)' post $($mi + 1): expected 0 replies, found $(@($mReplies).Count)."
            }
        }
    }
}
if ($totalTeamsMessages -ne [int]$manifest.resources.teams.storyMessageCount -or
    $totalTeamsMessages -ne [int]$dataSpec.expectedFacts.teamMessageCount) {
    throw "Count mismatch: teams-messages.json has $totalTeamsMessages total messages, expected $($manifest.resources.teams.storyMessageCount)."
}

# SharePoint manifest & canonical dataSpec checks
$declaredSites = @($spData.sites)
if ($declaredSites.Count -ne 1) {
    throw "Count mismatch: sharepoint-sites.json declares $($declaredSites.Count) sites, expected 1."
}
$spSite = $declaredSites[0]
$spSiteLib = if ($spSite.PSObject.Properties.Name -contains 'documentLibrary' -and $spSite.documentLibrary) {
    "$($spSite.documentLibrary)"
} elseif ($spSite.documents -and @($spSite.documents).Count -gt 0 -and ($spSite.documents[0].PSObject.Properties.Name -contains 'libraryName')) {
    "$($spSite.documents[0].libraryName)"
} else {
    ''
}
if ("$($spSite.alias)" -cne "$($manifest.resources.sharePoint.alias)" -or
    "$($spSite.alias)" -cne "$($dataSpec.resources.sharePointAlias)" -or
    "$($spSite.displayName)" -cne "$($manifest.resources.sharePoint.displayName)" -or
    "$($spSite.displayName)" -cne "$($dataSpec.resources.sharePointDisplayName)" -or
    $spSiteLib -cne "$($manifest.resources.sharePoint.documentLibrary)" -or
    $spSiteLib -cne "$($dataSpec.resources.sharePointLibrary)") {
    throw "Target mismatch: sharepoint-sites.json site identity/library does not match manifest and dataSpec."
}
if (($spSite.PSObject.Properties.Name -contains 'visibility') -and "$($spSite.visibility)" -ne 'Private') {
    throw "SharePoint contract violation: site visibility must be 'Private'."
}
if (@($spSite.owners).Count -ne 1 -or "$($spSite.owners[0])" -ne 'Admin') {
    throw "SharePoint contract violation: site owners must be exclusively ['Admin']."
}
$expectedBusinessRoles = @($expectedRoles | Where-Object { $_ -ne 'Admin' })
$declaredSiteMembers = @($spSite.members | ForEach-Object { "$_" })
if ($declaredSiteMembers -contains 'Admin' -or
    $declaredSiteMembers.Count -ne 7 -or
    @(Compare-Object -ReferenceObject ($expectedBusinessRoles | Sort-Object) -DifferenceObject ($declaredSiteMembers | Sort-Object)).Count -ne 0) {
    throw "SharePoint contract violation: site members must be the exact 7 business roles without Admin."
}
$expectedSpDocs = @($dataSpec.sharePointDocuments | ForEach-Object { "$_" })
$declaredSpDocs = @($spSite.documents)
if ($declaredSpDocs.Count -ne [int]$manifest.resources.sharePoint.documentCount -or
    $declaredSpDocs.Count -ne $expectedSpDocs.Count) {
    throw "Count mismatch: sharepoint-sites.json declares $($declaredSpDocs.Count) documents, expected $($manifest.resources.sharePoint.documentCount)."
}
for ($di = 0; $di -lt $expectedSpDocs.Count; $di++) {
    $docFile = "$($declaredSpDocs[$di].sourceFilename)"
    if ($docFile -cne $expectedSpDocs[$di]) {
        throw "SharePoint document mismatch at index ${di}: expected '$($expectedSpDocs[$di])', found '$docFile'."
    }
}
$spLists = if (($spSite.PSObject.Properties.Name -contains 'lists') -and $null -ne $spSite.lists) { @($spSite.lists) } else { @() }
$spPages = if (($spSite.PSObject.Properties.Name -contains 'pages') -and $null -ne $spSite.pages) { @($spSite.pages) } else { @() }
if (@($spLists).Count -gt 0 -or @($spPages).Count -gt 0) {
    throw "SharePoint contract violation: custom lists ($(@($spLists).Count)) and pages ($(@($spPages).Count)) are forbidden."
}

# Outlook manifest & canonical dataSpec checks
$declaredThreads = @($emailsData.emailThreads)
$specThreads = @($dataSpec.emailThreads)
if ($declaredThreads.Count -ne [int]$manifest.resources.outlook.threadCount -or
    $declaredThreads.Count -ne $specThreads.Count) {
    throw "Count mismatch: emails.json declares $($declaredThreads.Count) threads, expected $($manifest.resources.outlook.threadCount)."
}
$totalEmails = 0
for ($ti = 0; $ti -lt $specThreads.Count; $ti++) {
    $th = $declaredThreads[$ti]
    $specTh = $specThreads[$ti]
    $thEmails = @($th.emails)
    $specAuthors = @($specTh.authors | ForEach-Object { "$_" })
    if ($thEmails.Count -ne 3 -or $specAuthors.Count -ne 3) {
        throw "Count mismatch: thread '$($th.threadName)' has $($thEmails.Count) messages, expected 3."
    }
    for ($ei = 0; $ei -lt 3; $ei++) {
        $em = $thEmails[$ei]
        $totalEmails++
        $expectedAction = if ($ei -eq 0) { 'send' } else { 'reply' }
        if ("$($em.action)" -cne $expectedAction) {
            throw "Email action mismatch in thread '$($th.threadName)' message $($ei + 1): expected '$expectedAction', found '$($em.action)'."
        }
        $expectedSubject = if ($ei -eq 0) { "$($specTh.subject)" } else { "Re: $($specTh.subject)" }
        if ("$($em.subject)" -cne $expectedSubject) {
            throw "Email subject mismatch in thread '$($th.threadName)' message $($ei + 1): expected '$expectedSubject', found '$($em.subject)'."
        }
        if ("$($em.fromRole)" -cne $specAuthors[$ei]) {
            throw "Email fromRole mismatch in thread '$($th.threadName)' message $($ei + 1): expected '$($specAuthors[$ei])', found '$($em.fromRole)'."
        }
        $expectedTo = @(if ($ei -eq 0) {
            $specTh.initialToRoles | ForEach-Object { "$_" }
        } else {
            $specTh.replyToRoles[$ei - 1] | ForEach-Object { "$_" }
        })
        $actualTo = @($em.toRoles | ForEach-Object { "$_" })
        if (@($actualTo).Count -ne @($expectedTo).Count -or ($actualTo -join ',') -cne ($expectedTo -join ',')) {
            throw "Email toRoles mismatch in thread '$($th.threadName)' message $($ei + 1): expected [$($expectedTo -join ', ')], found [$($actualTo -join ', ')]."
        }
        $ccRoles = @($em.ccRoles | ForEach-Object { "$_" })
        $expectedCc = @($dataSpec.emailRules.ccEveryMessage | ForEach-Object { "$_" })
        if (@($ccRoles).Count -ne @($expectedCc).Count -or ($ccRoles -join ',') -cne ($expectedCc -join ',')) {
            throw "Email contract violation: ccRoles in thread '$($th.threadName)' message $($ei + 1) must be exactly [$($expectedCc -join ', ')]."
        }
        if ([string]::IsNullOrWhiteSpace("$($em.bodyHtml)")) {
            throw "Email bodyHtml cannot be empty in thread '$($th.threadName)' message $($ei + 1)."
        }
        foreach ($srcKey in @($specTh.sources | ForEach-Object { "$_" })) {
            if ("$($em.bodyHtml)" -notmatch [regex]::Escape($srcKey)) {
                throw "Email bodyHtml in thread '$($th.threadName)' message $($ei + 1) is missing canonical source marker '$srcKey'."
            }
        }
    }
}
if ($totalEmails -ne [int]$manifest.resources.outlook.sendAndReplyCount -or
    $totalEmails -ne [int]$dataSpec.expectedFacts.emailMessageCount) {
    throw "Count mismatch: emails.json has $totalEmails total messages, expected $($manifest.resources.outlook.sendAndReplyCount)."
}

# Only load engine helper scripts AFTER the complete codeHashes and canonical matrix gate succeeds.
. "$PSScriptRoot\Seed-Idempotency.ps1"
. "$PSScriptRoot\Seed-GraphRead.ps1"
. "$PSScriptRoot\Seed-OneDriveOwnership.ps1"

# ───────────────────────────────────────────────────────
# 3. Acquire token & verify JWT claims before any Graph resource call
# ───────────────────────────────────────────────────────
$tokenUri = "https://login.microsoftonline.com/$($config.tenantId)/oauth2/v2.0/token"
$tokenBody = @{
    grant_type    = 'client_credentials'
    client_id     = $config.clientId
    client_secret = $config.clientSecret
    scope         = 'https://graph.microsoft.com/.default'
}
$tokenResp = try {
    Invoke-RestMethod -Method POST -Uri $tokenUri -Body $tokenBody -ContentType 'application/x-www-form-urlencoded'
} catch {
    $secretVal = if ($config -and ($config.PSObject.Properties.Name -contains 'clientSecret')) { "$($config.clientSecret)" } else { '' }
    throw (New-SeedSanitizedErrorRecord -ErrorRecord $_ -Operation "OAuth POST $tokenUri failed" -Secrets @($secretVal))
}
if (-not $tokenResp -or -not ($tokenResp.PSObject.Properties.Name -contains 'access_token') -or [string]::IsNullOrWhiteSpace("$($tokenResp.access_token)")) {
    throw "Failed to acquire Graph access token."
}
$accessToken = "$($tokenResp.access_token)"
$claims = ConvertFrom-AB730JwtPayload -Token $accessToken

if (-not ($claims.PSObject.Properties.Name -contains 'tid') -or
    -not [string]::Equals("$($claims.tid)", $expectedTenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "JWT claim mismatch: token tid '$($claims.tid)' != manifest tenantId '$expectedTenantId'."
}
$tokenAppId = if ($claims.PSObject.Properties.Name -contains 'appid') { "$($claims.appid)" } elseif ($claims.PSObject.Properties.Name -contains 'azp') { "$($claims.azp)" } else { '' }
if (-not [string]::Equals($tokenAppId, $expectedClientId, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "JWT claim mismatch: token appid '$tokenAppId' != manifest clientId '$expectedClientId'."
}
$validAuds = @('https://graph.microsoft.com', 'https://graph.microsoft.com/', '00000003-0000-0000-c000-000000000000')
if (-not ($claims.PSObject.Properties.Name -contains 'aud') -or ($validAuds -notcontains "$($claims.aud)")) {
    throw "JWT claim mismatch: invalid token audience '$($claims.aud)'."
}
if (-not ($claims.PSObject.Properties.Name -contains 'exp') -or $null -eq $claims.exp) {
    throw "JWT claim mismatch: token is missing 'exp' claim."
}
$expUtc = [DateTimeOffset]::FromUnixTimeSeconds([long]$claims.exp).UtcDateTime
if ($expUtc -le [DateTime]::UtcNow.AddMinutes(5)) {
    throw "JWT claim mismatch: token is expired or expires within 5 minutes ($($expUtc.ToString('o')))."
}

$requiredAppRoles = @(
    'User.Read.All',
    'Mail.ReadWrite',
    'Mail.Send',
    'Teamwork.Migrate.All',
    'TeamMember.ReadWrite.All',
    'Group.ReadWrite.All',
    'Files.ReadWrite.All',
    'Sites.Manage.All',
    'Sites.ReadWrite.All'
)
$grantedRoles = if ($claims.PSObject.Properties.Name -contains 'roles' -and $null -ne $claims.roles) { @($claims.roles | ForEach-Object { "$_" }) } else { @() }
$missingRoles = @($requiredAppRoles | Where-Object { @($grantedRoles) -notcontains $_ })
if (@($missingRoles).Count -gt 0) {
    throw "JWT missing required application permissions: $($missingRoles -join ', ')."
}

$global:AccessToken = $accessToken

# Disable StrictMode before invoking the shared engine phase scripts, which rely on standard
# PowerShell scalar/collection .Count semantics under $ErrorActionPreference = 'Stop'.
Set-StrictMode -Off

# ───────────────────────────────────────────────────────
# 4. Run FreshOnly Preflight (strictly read-only)
# ───────────────────────────────────────────────────────
$preflightSplat = @{
    ConfigPath = $resolvedConfigPath
    FreshOnly  = $true
}
if ($selectedPhases -contains 'OneDrive')   { $preflightSplat['FilesManifestPath'] = $filesManifestPath }
if ($selectedPhases -contains 'Emails')     { $preflightSplat['EmailsPath']        = $emailsPath }
if ($selectedPhases -contains 'Teams')      { $preflightSplat['TeamsMessagesPath'] = $teamsPath }
if ($selectedPhases -contains 'SharePoint') { $preflightSplat['SharePointPath']    = $spPath }

& "$PSScriptRoot\Invoke-SeedPreflight.ps1" @preflightSplat

if (-not $Live) {
    Write-Host "`n[AB-730] Preflight-only verification succeeded (0 writes executed). Pass -Live to seed fresh targets." -ForegroundColor Green
    return
}

# ───────────────────────────────────────────────────────
# 5. Live FreshOnly Seeding Phases (no retries, no DELETE/PATCH recovery)
# ───────────────────────────────────────────────────────
Assert-SeedNoWriteRetryContext -FreshOnly
Write-Host "`n[AB-730] Starting -Live fresh-target-only seeding ($($selectedPhases -join ', '))..." -ForegroundColor Cyan

if ($selectedPhases -contains 'OneDrive') {
    & "$PSScriptRoot\Invoke-UploadFiles.ps1" `
        -ConfigPath $resolvedConfigPath `
        -FilesManifestPath $filesManifestPath `
        -StrictScenario `
        -FreshOnly
}

if ($selectedPhases -contains 'Emails') {
    & "$PSScriptRoot\Invoke-SeedEmails.ps1" `
        -ConfigPath $resolvedConfigPath `
        -EmailsPath $emailsPath `
        -FreshOnly
}

if ($selectedPhases -contains 'Teams') {
    & "$PSScriptRoot\Invoke-SeedTeamsChannel.ps1" `
        -ConfigPath $resolvedConfigPath `
        -TeamsMessagesPath $teamsPath `
        -FreshOnly
}

if ($selectedPhases -contains 'SharePoint') {
    & "$PSScriptRoot\Invoke-SeedSharePoint.ps1" `
        -ConfigPath $resolvedConfigPath `
        -SharePointPath $spPath `
        -FreshOnly
}

Write-Host "`n[AB-730] Selected fresh-target-only seeding phases completed ($($selectedPhases -join ', '))." -ForegroundColor Green
