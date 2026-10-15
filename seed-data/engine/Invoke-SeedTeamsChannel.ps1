<#
.SYNOPSIS
    Seed Teams channels and messages via Microsoft Graph (Application Permissions).
.DESCRIPTION
    Creates a Team (migration mode), channels, and injects messages with historical dates
    using the Teams Migration API, completes migration, then adds members.

    Idempotency (see engine/Seed-Idempotency.ps1):
      * The team is keyed by the exact dated displayName.
          0 matching team groups -> create
          1 matching team group  -> reuse
          >1                     -> throw, listing resource ids/names only
      * A reused team must first PROVE that the demo operator admin already owns it: the
        Microsoft 365 group owner collection must contain the admin, and the team membership
        must carry the admin with a 'roles' entry of 'owner'. Either one missing -> ABORT
        before any write. The seeder never promotes a role on a team it did not create.
      * Channels are keyed by exact channel displayName (reuse one, throw on duplicates), and
        the channel set must match exactly: one General, one instance of every declared
        channel, no extras.
      * Message identity is a SHA-256 hash of whitespace-normalized HTML content.
        Replies are fetched separately per parent via /messages/{id}/replies because
        list-channel-messages does not return replies.
      * A reused team whose expected top-level/reply set is complete -> skip messages.
      * A reused team missing any expected message or reply -> ABORT.
        Controller ruling (SDD ledger, Task 3): no automatic recovery. Migration may
        already be complete, in which case historical timestamps can no longer be
        written and a "resume" would silently corrupt the demo dataset.
      * Both completeMigration levels run on every path, including a reused complete team:
        a run that wrote all messages but died before completion would otherwise leave the
        team in migration mode forever, and membership cannot be added in that state.
        "Already completed" / "not in migration" is success; anything else throws.
      * This script never deletes anything and never matches resources by wildcard.
.PARAMETER ConfigPath
    Path to the scenario config.json.
.PARAMETER TeamsMessagesPath
    Path to the teams-messages.json data file.
#>
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [Parameter(Mandatory)][string]$TeamsMessagesPath,
    [switch]$FreshOnly
)

$ErrorActionPreference = "Stop"

. "$PSScriptRoot\Seed-Idempotency.ps1"
. "$PSScriptRoot\Seed-GraphRead.ps1"

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$teamsData = Get-Content $TeamsMessagesPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Build role -> user info lookup
$roleMap = @{}
$userIdCache = @{}
foreach ($prop in $config.roles.PSObject.Properties) {
    $roleMap[$prop.Name] = $prop.Value
}

function Invoke-Graph {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [string]$ApiVersion = "v1.0",
        [switch]$IncludeHeaders
    )
    $headers = @{
        Authorization  = "Bearer $($global:AccessToken)"
        "Content-Type" = "application/json"
    }
    $bodyJson = if ($Body) { $Body | ConvertTo-Json -Depth 20 -Compress } else { $null }
    $params = @{ Method = $Method; Uri = $Uri; Headers = $headers }
    if ($bodyJson) { $params.Body = [System.Text.Encoding]::UTF8.GetBytes($bodyJson) }
    if ($FreshOnly -and -not [string]::Equals($Method, 'GET', [System.StringComparison]::OrdinalIgnoreCase)) {
        Assert-SeedNoWriteRetryContext -FreshOnly
    }
    try {
        if ($IncludeHeaders) {
            $respHeaders = $null
            $respStatus = $null
            $irmCmd = Get-Command Invoke-RestMethod -ErrorAction SilentlyContinue
            if ($null -ne $irmCmd -and $irmCmd.Parameters -and $irmCmd.Parameters.ContainsKey('ResponseHeadersVariable')) {
                $params['ResponseHeadersVariable'] = 'respHeaders'
            }
            if ($null -ne $irmCmd -and $irmCmd.Parameters -and $irmCmd.Parameters.ContainsKey('StatusCodeVariable')) {
                $params['StatusCodeVariable'] = 'respStatus'
            }
            $respBody = Invoke-RestMethod @params
            return [pscustomobject]@{
                Body       = $respBody
                Headers    = $respHeaders
                StatusCode = $respStatus
            }
        }
        return Invoke-RestMethod @params
    } catch {
        $secretVal = if ($config -and ($config.PSObject.Properties.Name -contains 'clientSecret')) { "$($config.clientSecret)" } else { '' }
        throw (New-SeedSanitizedErrorRecord -ErrorRecord $_ -Operation "Graph $Method $Uri failed" -Secrets @($secretVal))
    }
}

$userDisplayNameCache = @{}
function Get-UserId([string]$upn) {
    $trimmedUpn = "$upn".Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedUpn)) {
        throw "Cannot resolve empty user UPN."
    }
    if (-not $userIdCache.ContainsKey($trimmedUpn)) {
        $user = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$trimmedUpn"
        if ($null -eq $user -or -not ($user.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($user.id)")) {
            throw "User resolution failed for '$trimmedUpn': Graph returned null or missing id."
        }
        if (($user.PSObject.Properties.Name -contains 'userPrincipalName') -and -not [string]::IsNullOrWhiteSpace("$($user.userPrincipalName)") -and
            -not [string]::Equals("$($user.userPrincipalName)".Trim(), $trimmedUpn, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "User resolution mismatch for '$trimmedUpn': Graph returned userPrincipalName '$($user.userPrincipalName)'."
        }
        $userIdCache[$trimmedUpn] = "$($user.id)".Trim()
        if ($user -and ($user.PSObject.Properties.Name -contains 'displayName') -and $user.displayName) {
            $userDisplayNameCache[$trimmedUpn] = "$($user.displayName)"
        } else {
            $userDisplayNameCache[$trimmedUpn] = $trimmedUpn
        }
    }
    return $userIdCache[$trimmedUpn]
}

function Get-UserForMessage([string]$roleName) {
    if (-not $roleMap.ContainsKey($roleName)) { throw "Unknown role in teams-messages.json: $roleName" }
    $role = $roleMap[$roleName]
    $userId = Get-UserId $role.upn
    $disp = if (($role.PSObject.Properties.Name -contains 'displayName') -and $role.displayName) {
        "$($role.displayName)"
    } elseif ($userDisplayNameCache.ContainsKey($role.upn)) {
        $userDisplayNameCache[$role.upn]
    } else {
        "$($role.upn)"
    }
    return @{
        id          = $userId
        upn         = $role.upn
        displayName = $disp
    }
}

function Get-SeedGroupsByDisplayName {
    param([Parameter(Mandatory)][string]$DisplayName)
    $escaped = $DisplayName -replace "'", "''"
    $filter = [System.Uri]::EscapeDataString("displayName eq '$escaped'")
    $groups = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName,resourceProvisioningOptions"
    $matched = @($groups | Where-Object {
        $null -ne $_ -and (
            (-not ($_.PSObject.Properties.Name -contains 'displayName')) -or
            [string]::Equals("$($_.displayName)".Trim(), $DisplayName.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
        )
    })
    return , @($matched)
}

# Exact displayName lookup, restricted to groups that are actually Teams.
function Get-SeedTeamGroup {
    param([Parameter(Mandatory)][string]$DisplayName)
    $groups = Get-SeedGroupsByDisplayName -DisplayName $DisplayName
    $matched = @($groups | Where-Object {
        $null -ne $_ -and $_.resourceProvisioningOptions -contains "Team"
    })
    return , @($matched)
}

function Get-SeedHeaderValue {
    param($Headers, [string]$HeaderName)
    if ($null -eq $Headers) { return '' }
    if ($Headers -is [System.Collections.IDictionary]) {
        foreach ($k in $Headers.Keys) {
            if ([string]::Equals("$k", $HeaderName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $val = $Headers[$k]
                if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                    return (@($val)[0]).ToString().Trim()
                }
                return "$val".Trim()
            }
        }
    } elseif ($Headers.PSObject -and ($Headers.PSObject.Properties.Name -contains $HeaderName)) {
        $val = $Headers.$HeaderName
        if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
            return (@($val)[0]).ToString().Trim()
        }
        return "$val".Trim()
    }
    return ''
}

function Resolve-SeedTeamPostIdentity {
    param($PostResponse)
    $teamId = ''
    $operationUri = ''
    if ($null -eq $PostResponse) {
        return [pscustomobject]@{ TeamId = ''; OperationUri = '' }
    }
    $bodyId = ''
    $contentLocId = ''
    $locationId = ''
    $bodyObj = if ($PostResponse.PSObject.Properties.Name -contains 'Body') { $PostResponse.Body } else { $PostResponse }
    if ($bodyObj -and ($bodyObj.PSObject.Properties.Name -contains 'id') -and -not [string]::IsNullOrWhiteSpace("$($bodyObj.id)")) {
        $bodyId = "$($bodyObj.id)".Trim()
    }
    $headers = if ($PostResponse.PSObject.Properties.Name -contains 'Headers') { $PostResponse.Headers } else { $null }
    $contentLoc = Get-SeedHeaderValue -Headers $headers -HeaderName 'Content-Location'
    $location = Get-SeedHeaderValue -Headers $headers -HeaderName 'Location'

    if (-not [string]::IsNullOrWhiteSpace($contentLoc)) {
        $m = [regex]::Match($contentLoc, '(?i)/teams(?:\(\x27([^\x27]+)\x27\)|/([^/\x27?]+))')
        if ($m.Success) {
            $contentLocId = $(if ($m.Groups[1].Success) { $m.Groups[1].Value } else { $m.Groups[2].Value }).Trim()
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        $m = [regex]::Match($location, '(?i)/teams(?:\(\x27([^\x27]+)\x27\)|/([^/\x27?]+))')
        if ($m.Success) {
            $locationId = $(if ($m.Groups[1].Success) { $m.Groups[1].Value } else { $m.Groups[2].Value }).Trim()
        }
    }
    $candidateIds = @(@($bodyId, $contentLocId, $locationId) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($candidateIds.Count -gt 0) {
        $teamId = $candidateIds[0]
        foreach ($cid in $candidateIds) {
            if (-not [string]::Equals($cid, $teamId, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Team provisioning identity mismatch: POST /teams response body/headers returned conflicting team IDs ($($candidateIds -join ', '))."
            }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        if ($location -match '^https://graph\.microsoft\.com/v1\.0/') {
            $operationUri = $location
        } elseif ($location.StartsWith('/v1.0/', [System.StringComparison]::OrdinalIgnoreCase)) {
            $operationUri = "https://graph.microsoft.com$location"
        } elseif ($location.StartsWith('/teams', [System.StringComparison]::OrdinalIgnoreCase)) {
            $operationUri = "https://graph.microsoft.com/v1.0$location"
        }
    }
    return [pscustomobject]@{
        TeamId       = $teamId
        OperationUri = $operationUri
    }
}

# Bounded polling after an async team creation. Returns $null when the budget is spent;
# the caller turns that into a terminating error.
function Wait-SeedForTeamProvisioning {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [string]$TeamId = '',
        [string]$OperationUri = '',
        [switch]$FreshOnly,
        [int]$MaxRetries = 30,
        [int]$DelaySec = 5
    )
    $anchorId = "$TeamId".Trim()
    if (-not [string]::IsNullOrWhiteSpace($OperationUri) -and $OperationUri -match '(?i)/teams(?:\(\x27([^\x27]+)\x27\)|/([^/\x27?]+))') {
        $opUriTeamId = $(if ($Matches[1]) { $Matches[1] } else { $Matches[2] }).Trim()
        if (-not [string]::IsNullOrWhiteSpace($anchorId) -and -not [string]::Equals($opUriTeamId, $anchorId, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Team provisioning identity mismatch: OperationUri team ID '$opUriTeamId' disagrees with immutable team anchor '$anchorId'."
        }
        if ([string]::IsNullOrWhiteSpace($anchorId)) {
            $anchorId = $opUriTeamId
        }
    }
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        Start-Sleep -Seconds $DelaySec
        try {
            if (-not [string]::IsNullOrWhiteSpace($OperationUri) -or -not [string]::IsNullOrWhiteSpace($anchorId)) {
                $opDone = [string]::IsNullOrWhiteSpace($OperationUri)
                if (-not [string]::IsNullOrWhiteSpace($OperationUri)) {
                    $op = Invoke-Graph -Method GET -Uri $OperationUri
                    if ($op -and ($op.PSObject.Properties.Name -contains 'targetResourceId') -and -not [string]::IsNullOrWhiteSpace("$($op.targetResourceId)")) {
                        $opTargetId = "$($op.targetResourceId)".Trim()
                        if (-not [string]::IsNullOrWhiteSpace($anchorId) -and -not [string]::Equals($opTargetId, $anchorId, [System.StringComparison]::OrdinalIgnoreCase)) {
                            throw "Team provisioning identity mismatch: operation targetResourceId '$opTargetId' disagrees with immutable team anchor '$anchorId'."
                        }
                        $anchorId = $opTargetId
                    }
                    if ($op -and ($op.PSObject.Properties.Name -contains 'targetResourceLocation') -and -not [string]::IsNullOrWhiteSpace("$($op.targetResourceLocation)")) {
                        if ("$($op.targetResourceLocation)" -match '(?i)/teams(?:\(\x27([^\x27]+)\x27\)|/([^/\x27?]+))') {
                            $opLocId = $(if ($Matches[1]) { $Matches[1] } else { $Matches[2] }).Trim()
                            if (-not [string]::IsNullOrWhiteSpace($anchorId) -and -not [string]::Equals($opLocId, $anchorId, [System.StringComparison]::OrdinalIgnoreCase)) {
                                throw "Team provisioning identity mismatch: operation targetResourceLocation '$($op.targetResourceLocation)' disagrees with immutable team anchor '$anchorId'."
                            }
                        }
                    }
                    $opStatus = if ($op -and ($op.PSObject.Properties.Name -contains 'status')) { "$($op.status)".Trim().ToLowerInvariant() } else { '' }
                    if ($opStatus -in @('failed', 'invalid')) { return $null }
                    if ($opStatus -eq 'succeeded' -or [string]::IsNullOrWhiteSpace($opStatus)) {
                        $opDone = $true
                    }
                }
                if ($opDone -and -not [string]::IsNullOrWhiteSpace($anchorId)) {
                    $teamObj = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$anchorId"
                    if ($null -eq $teamObj -or -not ($teamObj.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($teamObj.id)")) {
                        throw "Team provisioning identity mismatch: GET /teams/$anchorId returned null or missing id."
                    }
                    $returnedId = "$($teamObj.id)".Trim()
                    if (-not [string]::Equals($returnedId, $anchorId, [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw "Team provisioning identity mismatch: GET /teams/$anchorId body id '$returnedId' disagrees with immutable team anchor '$anchorId'."
                    }
                    if (-not ($teamObj.PSObject.Properties.Name -contains 'displayName') -or [string]::IsNullOrWhiteSpace("$($teamObj.displayName)") -or -not [string]::Equals("$($teamObj.displayName)".Trim(), $DisplayName.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
                        $actualDisp = if ($teamObj.PSObject.Properties.Name -contains 'displayName') { "$($teamObj.displayName)" } else { '<missing>' }
                        throw "Team provisioning identity mismatch: GET /teams/$anchorId displayName '$actualDisp' is missing, blank, or disagrees with expected '$DisplayName'."
                    }
                    return [pscustomobject]@{ id = $anchorId; displayName = "$($teamObj.displayName)" }
                }
                continue
            }
            if ($FreshOnly) {
                return $null
            }
            $found = Get-SeedTeamGroup -DisplayName $DisplayName
            if ($found.Count -eq 1) { return $found[0] }
            if ($found.Count -gt 1) { return $null }
        } catch {
            if ((Get-SeedGraphErrorText -ErrorObject $_) -match 'Team provisioning identity mismatch') {
                throw
            }
            # Bounded, deliberate retry while the group is still provisioning.
            # No state is changed here and the caller throws once the budget is spent.
        }
    }
    return $null
}

function New-SeedMessageBody {
    param([Parameter(Mandatory)]$Message)

    $userInfo = Get-UserForMessage $Message.fromRole
    $body = @{
        createdDateTime = $Message.createdDateTime
        from            = @{
            user = @{
                id               = $userInfo.id
                displayName      = $userInfo.displayName
                userIdentityType = "aadUser"
            }
        }
        body            = @{
            contentType = "html"
            content     = $Message.bodyHtml
        }
    }

    if ($Message.mentions -and $Message.mentions.Count -gt 0) {
        $mentionsList = @()
        for ($mi = 0; $mi -lt $Message.mentions.Count; $mi++) {
            $mentionUser = Get-UserForMessage $Message.mentions[$mi]
            $mentionsList += @{
                id          = $mi
                mentionText = $mentionUser.displayName
                mentioned   = @{
                    user = @{
                        id               = $mentionUser.id
                        displayName      = $mentionUser.displayName
                        userIdentityType = "aadUser"
                    }
                }
            }
        }
        $body.mentions = $mentionsList
    }

    return $body
}

# ---------------------------------------------------------------
Write-Host "`n===== Seeding Teams Channels =====" -ForegroundColor Cyan

$teamDisplayName = $config.teamDisplayName
if (-not $teamDisplayName) { throw "config.json is missing 'teamDisplayName' - the team is keyed on the exact dated display name." }
$adminUpn = $config.adminUpn
if (-not $adminUpn) { throw "config.json is missing 'adminUpn' - the demo operator must be provable as team owner." }

$plan = Get-SeedExpectedTeamsPlan -Channels $teamsData.channels
Write-Host "  Declared: $($plan.TopLevelCount) top-level messages / $($plan.ReplyCount) replies in $($plan.Channels.Count) channel(s)" -ForegroundColor Gray

# Team + channel creation dates must precede the earliest declared message.
$allDates = @()
$nowUtc = [datetime]::UtcNow
foreach ($ch in @($teamsData.channels)) {
    foreach ($m in @($ch.messages)) {
        $dt = [datetime]::Parse($m.createdDateTime, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
        if ($FreshOnly -and $dt -ge $nowUtc) {
            throw "FreshOnly Teams contract violation: message createdDateTime '$($m.createdDateTime)' in channel '$($ch.channelName)' is not in the past. No writes were made."
        }
        $allDates += $dt
        if (($m.PSObject.Properties.Name -contains 'replies') -and $m.replies) {
            foreach ($r in @($m.replies)) {
                $rdt = [datetime]::Parse($r.createdDateTime, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
                if ($FreshOnly -and $rdt -ge $nowUtc) {
                    throw "FreshOnly Teams contract violation: reply createdDateTime '$($r.createdDateTime)' in channel '$($ch.channelName)' is not in the past. No writes were made."
                }
                $allDates += $rdt
            }
        }
    }
}
if ($allDates.Count -eq 0) { throw "teams-messages.json declares no messages." }
$teamCreatedDate = (($allDates | Sort-Object | Select-Object -First 1).ToUniversalTime().AddDays(-1)).ToString("yyyy-MM-ddT00:00:00Z")

# --- Step 1: resolve the team (create or reuse, never guess) ---
Write-Host "`n[Step 1] Resolving team: $teamDisplayName" -ForegroundColor Yellow

$allMatchingGroups = Get-SeedGroupsByDisplayName -DisplayName $teamDisplayName
$existingTeams = @($allMatchingGroups | Where-Object { $null -ne $_ -and $_.resourceProvisioningOptions -contains "Team" })
if ($FreshOnly -and $allMatchingGroups.Count -gt 0) {
    $ids = ($allMatchingGroups | ForEach-Object { $_.id }) -join '; '
    throw "Collision (FreshOnly): group with displayName '$teamDisplayName' already exists ($ids). Refusing to create or adopt existing resource."
}
if ($allMatchingGroups.Count -gt $existingTeams.Count -and $existingTeams.Count -eq 0) {
    $ids = ($allMatchingGroups | ForEach-Object { $_.id }) -join '; '
    throw "Collision: a non-Team group already uses displayName '$teamDisplayName' ($ids). Refusing to create a duplicate group."
}
$teamAction = Get-SeedResourceAction -Existing $existingTeams -ResourceKind 'team group' -ResourceKey $teamDisplayName -FreshOnly:$FreshOnly

$ownerUpn = if ($config.PSObject.Properties.Name -contains 'demoUserUpn') { $config.demoUserUpn } else { $adminUpn }
$teamId = $null
$isNewTeam = $false

if ($teamAction -eq 'Reuse') {
    $teamId = $existingTeams[0].id
    Write-Host "  REUSE: existing team group $teamId" -ForegroundColor Yellow

    # Same ownership proof the preflight already ran, executed again against live state right
    # before the write window (defense in depth). The demo operator admin must ALREADY be a
    # group owner and carry the Teams 'owner' role; a reused team that left admin as a plain
    # member aborts here, before migration completion and before any membership call. The
    # seeder never promotes a role on a team it did not create.
    $ownership = Get-SeedTeamOwnershipVerdict -GroupOwners (Get-SeedGroupOwnerUpns -GroupId $teamId) `
                                              -TeamMembers (Get-SeedTeamMembers -TeamId $teamId) `
                                              -AdminUpn $adminUpn `
                                              -TeamKey $teamDisplayName
    Write-Host "  OWNERSHIP: '$adminUpn' verified as group owner and team owner [$($ownership.Action)]" -ForegroundColor Yellow
} else {
    Write-Host "  CREATE: new team in migration mode (createdDateTime $teamCreatedDate)" -ForegroundColor White
    $isNewTeam = $true

    $teamBody = @{
        "template@odata.bind"               = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        displayName                         = $teamDisplayName
        description                         = $config.teamDescription
        createdDateTime                     = $teamCreatedDate
        "@microsoft.graph.teamCreationMode" = "migration"
    }

    $postResponse = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams" -Body $teamBody -IncludeHeaders
    $postIdentity = Resolve-SeedTeamPostIdentity -PostResponse $postResponse
    if ($postResponse -and $postResponse.Body -and ($postResponse.Body.PSObject.Properties.Name -contains 'id') -and -not [string]::IsNullOrWhiteSpace("$($postResponse.Body.id)")) {
        $teamId = "$($postResponse.Body.id)".Trim()
    } else {
        if ($FreshOnly -and [string]::IsNullOrWhiteSpace($postIdentity.TeamId) -and [string]::IsNullOrWhiteSpace($postIdentity.OperationUri)) {
            throw "POST /v1.0/teams for '$teamDisplayName' returned neither a resource id nor a Location/Content-Location async operation header in FreshOnly mode; refusing to poll by displayName."
        }
        Write-Host "  Waiting for team provisioning..." -ForegroundColor Gray
        $provisioned = Wait-SeedForTeamProvisioning -DisplayName $teamDisplayName -TeamId $postIdentity.TeamId -OperationUri $postIdentity.OperationUri -FreshOnly:$FreshOnly
        if (-not $provisioned) {
            throw ("Team '$teamDisplayName' did not provision as the expected team resource within the polling budget. " +
                   "Check Teamwork.Migrate.All consent and the tenant state, then re-run.")
        }
        $teamId = $provisioned.id
    }
    Write-Host "  Team created: $teamId" -ForegroundColor Green
}

# --- Step 2: resolve channels + decide message action ---
Write-Host "`n[Step 2] Resolving channels..." -ForegroundColor Yellow

$channelPlans = @()

if (-not $isNewTeam) {
    # Same verification the preflight already ran, executed again against live state right
    # before the write window: channel cardinality plus top-level AND reply content hashes.
    # Complete -> skip messages; partial, extra, missing or ambiguous -> abort.
    $existingChannels = Get-SeedTeamChannels -TeamId $teamId
    $verdict = Get-SeedTeamContentVerdict -PlanChannels $plan.Channels `
                                          -ExistingChannels $existingChannels `
                                          -ChannelStateProvider { param($Channel) Get-SeedChannelMessageState -TeamId $teamId -ChannelId $Channel.id } `
                                          -TeamKey $teamDisplayName

    foreach ($planChannel in $plan.Channels) {
        $cv = @($verdict.Channels | Where-Object { $_.ChannelName -eq $planChannel.ChannelName })[0]
        Write-Host "  REUSE channel '$($cv.ChannelName)': $($cv.ChannelId)" -ForegroundColor Yellow
        Write-Host "    Messages: $($cv.MessageAction) ($($cv.ExistingTopLevel) top-level found)" -ForegroundColor Gray
        $channelPlans += [pscustomobject]@{
            Plan          = $planChannel
            ChannelId     = $cv.ChannelId
            MessageAction = $cv.MessageAction
        }
    }
} else {
    foreach ($planChannel in $plan.Channels) {
        $channelName = $planChannel.ChannelName
        Write-Host "  CREATE channel '$channelName'" -ForegroundColor White
        $channelBody = @{
            "@microsoft.graph.channelCreationMode" = "migration"
            displayName                            = $channelName
            description                            = $planChannel.Source.channelDescription
            membershipType                         = "standard"
            createdDateTime                        = $teamCreatedDate
        }
        $created = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels" -Body $channelBody
        $channelId = $created.id
        if (-not $channelId) { throw "Channel creation for '$channelName' returned no channel id." }
        Write-Host "    Channel ID: $channelId" -ForegroundColor Gray
        Start-Sleep -Seconds 2

        $channelPlans += [pscustomobject]@{
            Plan          = $planChannel
            ChannelId     = $channelId
            MessageAction = 'Create'
        }
    }
}

# --- Step 3: inject messages with historical dates ---
Write-Host "`n[Step 3] Injecting messages..." -ForegroundColor Yellow

foreach ($cp in $channelPlans) {
    $channelName = $cp.Plan.ChannelName
    Write-Host "`n  --- Channel: $channelName ---" -ForegroundColor Yellow

    if ($cp.MessageAction -eq 'Skip') {
        Write-Host "    SKIP: all $($cp.Plan.Messages.Count) top-level messages and their replies already present." -ForegroundColor Yellow
        continue
    }

    foreach ($msg in $cp.Plan.Messages) {
        $source = $msg.Source
        Write-Host "    [$($source.fromRole)] $($source.createdDateTime)" -ForegroundColor White

        $msgBody = New-SeedMessageBody -Message $source
        $msgResult = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$($cp.ChannelId)/messages" -Body $msgBody
        $messageId = $msgResult.id
        if (-not $messageId) { throw "Teams message creation for channel '$channelName' returned no message id." }

        $sourceReplies = if (($source.PSObject.Properties.Name -contains 'replies') -and $source.replies) { @($source.replies) } else { @() }
        foreach ($reply in $sourceReplies) {
            Write-Host "      -> [$($reply.fromRole)] reply" -ForegroundColor Gray
            $replyBody = New-SeedMessageBody -Message $reply
            Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$($cp.ChannelId)/messages/$messageId/replies" -Body $replyBody | Out-Null
            Start-Sleep -Seconds 1
        }
        Start-Sleep -Seconds 2
    }
}

# --- Step 4: complete migration (channels first, then team) ---
# This runs on EVERY path, not only after a fresh creation. A previous run that wrote all
# messages but died before completing migration leaves the team in migration mode; on the
# re-run its content hashes look complete, message creation is skipped, and membership would
# then fail because a team in migration mode accepts no members. Both levels are idempotent:
# "already completed" / "not in migration" is success, anything else throws.
Write-Host "`n[Step 4] Completing migration (idempotent, channel level then team level)..." -ForegroundColor Yellow

# Every channel, including General, must be completed before the team.
$allChannels = Get-SeedTeamChannels -TeamId $teamId
foreach ($ch in @($allChannels)) {
    Write-Host "  Completing channel: $($ch.displayName)" -ForegroundColor White
    try {
        Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$($ch.id)/completeMigration" | Out-Null
        Start-Sleep -Seconds 2
    } catch {
        if (Test-SeedMigrationAlreadyComplete -ErrorRecord $_) {
            Write-Host "    (already completed)" -ForegroundColor Gray
        } else {
            $secretVal = if ($config -and ($config.PSObject.Properties.Name -contains 'clientSecret')) { "$($config.clientSecret)" } else { '' }
            throw "Failed to complete migration for channel '$($ch.displayName)' ($($ch.id)): $(Format-SeedSafeError -Message $_ -Secrets @($secretVal))"
        }
    }
}

Write-Host "  Completing team migration..." -ForegroundColor White
try {
    Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/completeMigration" | Out-Null
    Start-Sleep -Seconds 5
} catch {
    if (Test-SeedMigrationAlreadyComplete -ErrorRecord $_) {
        Write-Host "    (already completed)" -ForegroundColor Gray
    } else {
        $secretVal = if ($config -and ($config.PSObject.Properties.Name -contains 'clientSecret')) { "$($config.clientSecret)" } else { '' }
        throw "Failed to complete migration for team '$teamDisplayName' ($teamId): $(Format-SeedSafeError -Message $_ -Secrets @($secretVal))"
    }
}

# --- Step 5: add team members (idempotent) ---
# A duplicate-member error is absorbed here ONLY because the reuse path above already proved
# that the demo operator admin holds the expected owner role on this team (and the create path
# built the team itself). Without that proof, "already exists" could be hiding a member entry
# whose role is not the declared one.
Write-Host "`n[Step 5] Adding team members..." -ForegroundColor Yellow

$expectedTeamMembersByUpn = @{}
$seenPlannedUserIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($prop in $config.roles.PSObject.Properties) {
    $role = $prop.Value
    $upn = "$($role.upn)".Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) {
        throw "Teams membership error: role '$($prop.Name)' has an empty UPN."
    }

    $userId = Get-UserId $upn
    # `$x = if (...) { @('owner') } else { @() }` assigns the OUTPUT of a statement, which
    # PowerShell unwraps: a one-element array collapses to a scalar and an empty array to
    # $null. ConvertTo-Json then emits "roles":"owner" / "roles":null, and Graph rejects the
    # membership with 400 "Could not cast or convert from System.String to
    # System.Collections.Generic.IEnumerable`1[System.String]". aadUserConversationMember.roles
    # is a String collection, so the array is built by direct typed assignment instead.
    [string[]]$memberRoles = @()
    if ($FreshOnly) {
        if ($expectedTeamMembersByUpn.ContainsKey($upn.ToLowerInvariant()) -or -not $seenPlannedUserIds.Add("$userId")) {
            throw "FreshOnly Teams contract violation: duplicate UPN '$upn' or userId '$userId' across configured roles."
        }
        if ($prop.Name -eq 'Admin') {
            [string[]]$memberRoles = @('owner')
        } elseif ([string]::Equals("$upn", "$adminUpn", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "FreshOnly Teams contract violation: business role '$($prop.Name)' cannot share UPN with Admin '$adminUpn' or receive owner role."
        }
        $expectedTeamMembersByUpn[$upn.ToLowerInvariant()] = [pscustomobject]@{
            RoleName      = $prop.Name
            Upn           = $upn
            UserId        = "$userId".Trim()
            ExpectedRoles = @($memberRoles)
        }
    } else {
        if ($upn -eq $ownerUpn -or $upn -eq $adminUpn) { [string[]]$memberRoles = @('owner') }
    }
    $roleDisplay = if (($role.PSObject.Properties.Name -contains 'displayName') -and $role.displayName) { "$($role.displayName)" } else { "$upn" }
    Write-Host "  Adding: $roleDisplay ($upn) [$(if ($memberRoles.Count -gt 0) { 'owner' } else { 'member' })]" -ForegroundColor White

    $memberBody = @{
        "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
        "roles"           = $memberRoles
        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$userId')"
    }

    try {
        Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Body $memberBody | Out-Null
        Start-Sleep -Seconds 1
    } catch {
        if ((Test-SeedMemberAlreadyExists -ErrorRecord $_) -and (-not $FreshOnly)) {
            Write-Host "    (already a member)" -ForegroundColor Gray
        } else {
            $secretVal = if ($config -and ($config.PSObject.Properties.Name -contains 'clientSecret')) { "$($config.clientSecret)" } else { '' }
            throw "Failed to add '$upn' to team '$teamDisplayName' ($teamId): $(Format-SeedSafeError -Message $_ -Secrets @($secretVal))"
        }
    }
}

if ($FreshOnly) {
    $verifiedTeamMembers = Get-SeedTeamMembers -TeamId $teamId
    $expectedMemberCount = $expectedTeamMembersByUpn.Count
    if (@($verifiedTeamMembers).Count -ne $expectedMemberCount) {
        throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): expected $expectedMemberCount members, found $(@($verifiedTeamMembers).Count)."
    }
    $seenTeamUpns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $seenTeamUids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($tm in @($verifiedTeamMembers)) {
        if ($null -eq $tm -or
            -not ($tm.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($tm.id)") -or
            -not ($tm.PSObject.Properties.Name -contains 'userId') -or [string]::IsNullOrWhiteSpace("$($tm.userId)") -or
            -not ($tm.PSObject.Properties.Name -contains 'email') -or [string]::IsNullOrWhiteSpace("$($tm.email)")) {
            throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): member entry is missing id, userId, or email."
        }
        $tmEmail = "$($tm.email)".Trim()
        $tmUid = "$($tm.userId)".Trim()
        $tmKey = $tmEmail.ToLowerInvariant()
        if (-not $expectedTeamMembersByUpn.ContainsKey($tmKey)) {
            throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): unexpected/foreign member '$tmEmail' (userId '$tmUid')."
        }
        $expEntry = $expectedTeamMembersByUpn[$tmKey]
        if (-not [string]::Equals($tmUid, $expEntry.UserId, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): member '$tmEmail' userId '$tmUid' does not match expected '$($expEntry.UserId)'."
        }
        if (-not $seenTeamUpns.Add($tmEmail) -or -not $seenTeamUids.Add($tmUid)) {
            throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): duplicate member entry for '$tmEmail' ('$tmUid')."
        }
        $actualRoles = @($tm.roles | Where-Object { -not [string]::IsNullOrWhiteSpace("$_") } | ForEach-Object { "$_".Trim().ToLowerInvariant() })
        if ($expEntry.RoleName -eq 'Admin') {
            if ($actualRoles.Count -ne 1 -or $actualRoles[0] -ne 'owner') {
                throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): Admin '$tmEmail' must have exact role 'owner' (found: $($actualRoles -join ', '))."
            }
        } else {
            if ($actualRoles.Count -ne 0) {
                throw "FreshOnly Teams membership verification failed for '$teamDisplayName' ($teamId): business member '$tmEmail' ($($expEntry.RoleName)) must have 0 owner/elevated roles (found: $($actualRoles -join ', '))."
            }
        }
    }

    # Verify Teams backing M365 group owners (exact 1 Admin owner) and members (exact 8 configured roles)
    $backingOwners = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$teamId/owners?`$select=id,userPrincipalName"
    if (@($backingOwners).Count -ne 1 -or $null -eq $backingOwners[0] -or
        -not ($backingOwners[0].PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($backingOwners[0].id)") -or
        -not ($backingOwners[0].PSObject.Properties.Name -contains 'userPrincipalName') -or [string]::IsNullOrWhiteSpace("$($backingOwners[0].userPrincipalName)")) {
        throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): expected exactly 1 valid group owner, found $(@($backingOwners).Count)."
    }
    $bo = $backingOwners[0]
    $expectedAdminEntry = $expectedTeamMembersByUpn[$adminUpn.Trim().ToLowerInvariant()]
    if (-not [string]::Equals("$($bo.userPrincipalName)".Trim(), $expectedAdminEntry.Upn, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals("$($bo.id)".Trim(), $expectedAdminEntry.UserId, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): group owner '$($bo.userPrincipalName)' ('$($bo.id)') does not match Admin '$($expectedAdminEntry.Upn)' ('$($expectedAdminEntry.UserId)')."
    }

    $backingMembers = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$teamId/members"
    if (@($backingMembers).Count -ne $expectedMemberCount) {
        throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): expected $expectedMemberCount group members, found $(@($backingMembers).Count)."
    }
    $seenBackingUpns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $seenBackingUids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($gm in @($backingMembers)) {
        if ($null -eq $gm -or
            (($gm.PSObject.Properties.Name -contains '@odata.type') -and $gm.'@odata.type' -ne '#microsoft.graph.user') -or
            -not ($gm.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($gm.id)") -or
            -not ($gm.PSObject.Properties.Name -contains 'userPrincipalName') -or [string]::IsNullOrWhiteSpace("$($gm.userPrincipalName)")) {
            throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): group member entry is invalid or missing id/userPrincipalName."
        }
        $gmUpn = "$($gm.userPrincipalName)".Trim()
        $gmUid = "$($gm.id)".Trim()
        $gmKey = $gmUpn.ToLowerInvariant()
        if (-not $expectedTeamMembersByUpn.ContainsKey($gmKey)) {
            throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): unexpected/foreign group member '$gmUpn' (id '$gmUid')."
        }
        $expGm = $expectedTeamMembersByUpn[$gmKey]
        if (-not [string]::Equals($gmUid, $expGm.UserId, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): group member '$gmUpn' id '$gmUid' does not match expected '$($expGm.UserId)'."
        }
        if (-not $seenBackingUpns.Add($gmUpn) -or -not $seenBackingUids.Add($gmUid)) {
            throw "FreshOnly Teams backing group verification failed for '$teamDisplayName' ($teamId): duplicate group member '$gmUpn' ('$gmUid')."
        }
    }
}

$global:SeedTeamsResult = [pscustomobject]@{
    TeamDisplayName = $teamDisplayName
    TeamId          = $teamId
    TeamAction      = $teamAction
    Channels        = @($channelPlans | ForEach-Object {
        [pscustomobject]@{
            ChannelName   = $_.Plan.ChannelName
            ChannelId     = $_.ChannelId
            MessageAction = $_.MessageAction
        }
    })
}

Write-Host "`n===== Teams Channel Seeding Complete =====" -ForegroundColor Cyan
Write-Host "Team: $teamDisplayName [$teamAction] -> $teamId" -ForegroundColor Green
foreach ($c in $global:SeedTeamsResult.Channels) {
    Write-Host "  Channel '$($c.ChannelName)' [$($c.MessageAction)] -> $($c.ChannelId)" -ForegroundColor Green
}
