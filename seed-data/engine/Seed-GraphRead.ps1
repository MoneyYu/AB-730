<#
.SYNOPSIS
    Shared read-only Microsoft Graph collection readers for the PL-7008 M365 seed engine.
.DESCRIPTION
    GET requests only. Nothing here creates, modifies or deletes a tenant resource, and no
    resource is ever matched by wildcard.

    Dot-source this file from a script that already defines its own

        Invoke-Graph -Method GET -Uri <uri>

    wrapper (that is where the bearer token lives). Every collection read pages through
    `@odata.nextLink`: a partial first page would otherwise make a complete team or a
    complete mail thread look incomplete and abort a healthy re-run.

    Preflight (engine/Invoke-SeedPreflight.ps1) and the Teams phase
    (engine/Invoke-SeedTeamsChannel.ps1) both consume these readers so that the state they
    verify is produced by exactly one implementation.
#>

. "$PSScriptRoot\Seed-Idempotency.ps1"

function ConvertTo-SeedCanonicalGraphCollectionPath {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][System.Uri]$UriObj)

    $rawPath = [System.Uri]::UnescapeDataString($UriObj.AbsolutePath).TrimEnd('/')
    $rawPath = $rawPath -replace '(?i)&(?:\$|%24)skiptoken=[^/]+$', ''
    # Exchange Online returns @odata.nextLink for /users/<UPN>/mailFolders/inbox/messages using
    # the OData key-predicate form /users/<UPN>/mailFolders('inbox')/messages on the same mailbox.
    # Canonicalize ONLY that exact known Inbox path for the same user; never widen to other folders
    # or different users.
    if ($rawPath -match '^(?i)/(v1\.0|beta)/users/([^/]+)/mailFolders\(''inbox''\)/messages$') {
        $rawPath = "/$($Matches[1])/users/$($Matches[2])/mailFolders/inbox/messages"
    }
    return $rawPath
}

function ConvertTo-SeedGraphQueryMap {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][System.Uri]$UriObj)

    $map = @{}
    $q = $UriObj.Query
    if ([string]::IsNullOrWhiteSpace($q)) { return $map }
    $trimmed = $q.TrimStart('?')
    foreach ($part in ($trimmed -split '&')) {
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        $idx = $part.IndexOf('=')
        if ($idx -lt 0) {
            $k = [System.Uri]::UnescapeDataString($part).Trim().ToLowerInvariant()
            $v = ''
        } else {
            $k = [System.Uri]::UnescapeDataString($part.Substring(0, $idx)).Trim().ToLowerInvariant()
            $v = [System.Uri]::UnescapeDataString($part.Substring($idx + 1)).Trim()
        }
        $map[$k] = $v
    }
    return $map
}

# Follow @odata.nextLink until the collection is exhausted.
function Get-SeedGraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$MaxPages = 20,
        [switch]$AllowInitialNotFound
    )

    $initialUri = $null
    if (-not [System.Uri]::TryCreate($Uri, [System.UriKind]::Absolute, [ref]$initialUri) -or
        $initialUri.Scheme -ne 'https' -or
        -not [string]::Equals($initialUri.Host, 'graph.microsoft.com', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Graph collection '$Uri' must be an absolute https://graph.microsoft.com URI."
    }
    $expectedPath = ConvertTo-SeedCanonicalGraphCollectionPath -UriObj $initialUri
    $initialQuery = ConvertTo-SeedGraphQueryMap -UriObj $initialUri
    $allowedCursorKeys = @('$skip', '$skiptoken')
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

    $items = @()
    $next = $Uri
    $page = 0

    while ($next) {
        if (-not $visited.Add($next)) {
            throw "Graph collection '$Uri' returned a cyclic or repeated @odata.nextLink ('$next'); refusing to loop."
        }
        $page++
        if ($page -gt $MaxPages) {
            throw "Graph collection '$Uri' returned more than $MaxPages pages; refusing to keep paging."
        }

        $response = $null
        try {
            $response = Invoke-Graph -Method GET -Uri $next
        } catch {
            if ($AllowInitialNotFound -and $page -eq 1 -and (Test-SeedNotFoundError -ErrorRecord $_)) {
                return , @()
            }
            throw
        }
        if ($null -eq $response) {
            throw "Graph collection '$Uri' returned a null response on page $page."
        }

        $props = @($response.PSObject.Properties.Name)
        if ($props -notcontains 'value') {
            throw "Graph collection '$Uri' returned a malformed response missing 'value' on page $page."
        }
        $val = $response.value
        if ($null -eq $val -or -not ($val -is [System.Array] -or $val -is [System.Collections.IList])) {
            throw "Graph collection '$Uri' returned a non-array 'value' property on page $page."
        }
        $items += @($val)

        $following = $null
        if ($props -contains '@odata.nextLink') {
            $rawNext = $response.'@odata.nextLink'
            if ($null -eq $rawNext -or [string]::IsNullOrWhiteSpace("$rawNext")) {
                throw "Graph collection '$Uri' returned a null or empty @odata.nextLink on page $page."
            }
            $following = "$rawNext"
            if ($following -eq $next) {
                throw "Graph collection '$Uri' returned a self-referencing @odata.nextLink; refusing to loop."
            }
            if ($visited.Contains($following)) {
                throw "Graph collection '$Uri' returned a cyclic @odata.nextLink ('$following') to an already-visited page; refusing to loop."
            }
            $nextUri = $null
            if (-not [System.Uri]::TryCreate($following, [System.UriKind]::Absolute, [ref]$nextUri) -or
                $nextUri.Scheme -ne 'https' -or
                -not [string]::Equals($nextUri.Host, 'graph.microsoft.com', [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Graph collection '$Uri' returned a foreign or non-HTTPS host in @odata.nextLink ('$following'); refusing to follow."
            }
            $nextPath = ConvertTo-SeedCanonicalGraphCollectionPath -UriObj $nextUri
            if (-not [string]::Equals($nextPath, $expectedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Graph collection '$Uri' returned an @odata.nextLink outside the collection path boundary ('$nextPath' vs '$expectedPath'); refusing to follow."
            }
            $nextQuery = ConvertTo-SeedGraphQueryMap -UriObj $nextUri
            foreach ($initKey in $initialQuery.Keys) {
                if ($allowedCursorKeys -contains $initKey) { continue }
                if (-not $nextQuery.ContainsKey($initKey) -or $nextQuery[$initKey] -cne $initialQuery[$initKey]) {
                    throw "Graph collection '$Uri' returned an @odata.nextLink that altered or dropped query invariant '$initKey' ('$following'); refusing to follow."
                }
            }
            foreach ($nextKey in $nextQuery.Keys) {
                if ($allowedCursorKeys -contains $nextKey) { continue }
                if (-not $initialQuery.ContainsKey($nextKey)) {
                    throw "Graph collection '$Uri' returned an @odata.nextLink that injected unexpected query parameter '$nextKey' ('$following'); refusing to follow."
                }
            }
        }
        $next = $following
    }

    return , @($items)
}

# ───────────────────────────────────────────────────────
# Drive folders (OneDrive and SharePoint document libraries)
# ───────────────────────────────────────────────────────

# The names of the items directly under one drive folder, following @odata.nextLink so a
# truncated first page can never make a stored file look absent.
#
# A folder that does not exist yet answers 404 itemNotFound on page 1; that is reported as an
# empty snapshot, because the caller creates the folder and uploads everything. Any failure on
# page 2+ or any non-404 failure is rethrown so partial page-1 collisions are never dropped.
function Get-SeedDriveChildNames {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ChildrenUri)

    $items = Get-SeedGraphCollection -Uri $ChildrenUri -AllowInitialNotFound

    $names = @()
    foreach ($item in @($items)) {
        if ($null -eq $item) { continue }
        $props = @($item.PSObject.Properties.Name)
        if (($props -contains 'name') -and $item.name) { $names += "$($item.name)" }
    }
    return , @($names)
}

# ───────────────────────────────────────────────────────
# Mailbox (Inbox only)
# ───────────────────────────────────────────────────────

# Every seeded message reaches the admin mailbox as an Inbox item (admin is CC on all of
# them). The unscoped /users/{upn}/messages collection also returns Deleted Items and Sent
# Items, so a thread deleted by the operator would still be counted: a partial thread would
# not clear, and a "complete" thread could be skipped although no usable Inbox mail exists.
function Get-SeedInboxMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Upn,
        [int]$PageSize = 100
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$Upn/mailFolders/inbox/messages?" +
           "`$select=id,subject,conversationId,receivedDateTime&`$top=$PageSize"
    $messages = Get-SeedGraphCollection -Uri $uri
    return , @($messages)
}

# Most recent Inbox page only - used while polling for a message that was just sent.
function Get-SeedRecentInboxMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Upn,
        [int]$Top = 50
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$Upn/mailFolders/inbox/messages?" +
           "`$select=id,subject,conversationId,internetMessageId,receivedDateTime&`$top=$Top&`$orderby=receivedDateTime%20desc"
    $response = Invoke-Graph -Method GET -Uri $uri
    if ($null -eq $response) {
        throw "Recent Inbox query for '$Upn' returned a null response."
    }
    $props = @($response.PSObject.Properties.Name)
    if ($props -notcontains 'value' -or $null -eq $response.value -or -not ($response.value -is [System.Array] -or $response.value -is [System.Collections.IList])) {
        throw "Recent Inbox query for '$Upn' returned a missing or non-array 'value' property."
    }
    return , @($response.value)
}

# Thread state for one declared thread, matched manually against the Inbox collection.
# Pass -Messages to reuse a mailbox snapshot across several threads.
function Get-SeedInboxThreadState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Upn,
        [Parameter(Mandatory)][string]$FirstSubject,
        [Parameter()][AllowNull()]$Messages
    )

    $pool = $Messages
    if ($null -eq $pool) { $pool = Get-SeedInboxMessages -Upn $Upn }
    return Select-SeedThreadMessages -Messages $pool -FirstSubject $FirstSubject
}

# ───────────────────────────────────────────────────────
# Group ownership and Teams membership (identity proofs)
# ───────────────────────────────────────────────────────

# UPNs of the owners of a Microsoft 365 group, paged. Used to prove that the demo operator
# admin already owns a reused site group or the group behind a reused team.
# GET /groups/{id}/owners is served by the documented Group.ReadWrite.All application role:
# https://learn.microsoft.com/graph/api/group-list-owners
function Get-SeedGroupOwnerUpns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GroupId)

    try {
        $owners = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/owners?`$select=id,userPrincipalName"
    } catch {
        # Fail closed, but say what is missing: an unreadable owner collection is indistinguishable
        # from "nobody owns it", and treating it as the latter would let a write land in a resource
        # whose ownership was never proven.
        throw ("Cannot read the owners of group $GroupId, so the demo operator's ownership cannot be proven: " +
               "$(Format-SeedSafeError -Message $_). GET /groups/{id}/owners is served by the " +
               "Group.ReadWrite.All application permission (GroupMember.Read.All is the least-privileged " +
               "alternative); grant and consent it, then re-run. Nothing was written.")
    }
    # Assigned first on purpose: `@(Get-SeedUpnList ...)` would wrap the array the helper emits
    # as ONE pipeline object into a second array, and the nested value compares as
    # "System.Object[]" against every UPN.
    $upns = Get-SeedUpnList -Items $owners
    return , @($upns)
}

function Get-SeedGroupMemberUpns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GroupId)

    try {
        $members = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members"
    } catch {
        throw ("Cannot read members of group ${GroupId}: $(Format-SeedSafeError -Message $_). " +
               "Verify group membership read permission before seeding; no access is granted automatically.")
    }
    $users = @($members | Where-Object { $_ -and $_.'@odata.type' -eq '#microsoft.graph.user' })
    $upns = Get-SeedUpnList -Items $users
    return , @($upns)
}

function Assert-SeedSiteGroupRoles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter()][string]$Alias = '',
        [string[]]$OwnerRoles = @(),
        [string[]]$MemberRoles = @(),
        [Parameter(Mandatory)][Alias('RolesConfig')][object]$RoleMap,
        [string[]]$OwnerUpns = @(),
        [hashtable]$ExpectedUserIdsByUpn = $null,
        [Alias('ExactSet')][switch]$FreshOnly
    )

    $roleLookup = @{}
    if ($RoleMap -is [hashtable]) {
        foreach ($k in $RoleMap.Keys) {
            $v = $RoleMap[$k]
            $u = if ($v -is [string]) { "$v".Trim() } elseif ($null -ne $v -and ($v.PSObject.Properties.Name -contains 'upn')) { "$($v.upn)".Trim() } else { '' }
            $roleLookup["$k"] = $u
        }
    } elseif ($null -ne $RoleMap -and $null -ne $RoleMap.PSObject) {
        foreach ($p in $RoleMap.PSObject.Properties) {
            $v = $p.Value
            $u = if ($v -is [string]) { "$v".Trim() } elseif ($null -ne $v -and ($v.PSObject.Properties.Name -contains 'upn')) { "$($v.upn)".Trim() } else { '' }
            $roleLookup["$($p.Name)"] = $u
        }
    }

    $rawMembers = @()
    $memberUpns = @()
    if (@($MemberRoles).Count -gt 0 -or $null -ne $ExpectedUserIdsByUpn) {
        try {
            $rawMembers = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members"
        } catch {
            throw ("Cannot read members of group ${GroupId}: $(Format-SeedSafeError -Message $_). " +
                   "Verify group membership read permission before seeding; no access is granted automatically.")
        }
        $users = @($rawMembers | Where-Object { $_ -and $_.'@odata.type' -eq '#microsoft.graph.user' })
        $memberUpns = Get-SeedUpnList -Items $users
    }

    foreach ($entry in @(
        @{ Kind = 'owner'; Roles = @($OwnerRoles); Upns = @($OwnerUpns) }
        @{ Kind = 'member'; Roles = @($MemberRoles); Upns = @($memberUpns) }
    )) {
        foreach ($role in $entry.Roles) {
            if (-not $roleLookup.ContainsKey($role) -or [string]::IsNullOrWhiteSpace($roleLookup[$role])) {
                throw "Site '$Alias' declares undefined $($entry.Kind) role '$role'. No writes were made."
            }
            $upn = "$($roleLookup[$role])"
            if (-not (Test-SeedUpnInList -Upns $entry.Upns -Upn $upn)) {
                throw "Site '$Alias' group $GroupId is missing declared $($entry.Kind) role '$role' ($upn). No access was granted; aborting before writes."
            }
        }
    }

    if ($FreshOnly -or $null -ne $ExpectedUserIdsByUpn) {
        $expectedOwnerUpnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($r in @($OwnerRoles)) {
            if (-not $expectedOwnerUpnSet.Add($roleLookup[$r])) {
                throw "Site '$Alias' group $GroupId declares duplicate owner role/UPN '$($roleLookup[$r])'."
            }
        }
        $seenOwnerUpnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($ou in @($OwnerUpns)) {
            $trimmedOu = "$ou".Trim()
            if ([string]::IsNullOrWhiteSpace($trimmedOu) -or -not $expectedOwnerUpnSet.Contains($trimmedOu)) {
                throw "Site '$Alias' group $GroupId has unexpected or extra owner '$trimmedOu' (expected exact owners: $(@($expectedOwnerUpnSet) -join ', '))."
            }
            if (-not $seenOwnerUpnSet.Add($trimmedOu)) {
                throw "Site '$Alias' group $GroupId has duplicate owner '$trimmedOu'."
            }
        }
        if ($seenOwnerUpnSet.Count -ne $expectedOwnerUpnSet.Count) {
            throw "Site '$Alias' group $GroupId owner count mismatch: expected $($expectedOwnerUpnSet.Count), found $($seenOwnerUpnSet.Count)."
        }

        if (@($MemberRoles).Count -gt 0 -or $null -ne $ExpectedUserIdsByUpn) {
            $expectedMemberUpnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($r in @($MemberRoles)) {
                if (-not $expectedMemberUpnSet.Add($roleLookup[$r])) {
                    throw "Site '$Alias' group $GroupId declares duplicate member role/UPN '$($roleLookup[$r])'."
                }
            }
            if (@($rawMembers).Count -ne $expectedMemberUpnSet.Count -or @($memberUpns).Count -ne $expectedMemberUpnSet.Count) {
                throw "Site '$Alias' group $GroupId member count/identity mismatch: expected $($expectedMemberUpnSet.Count) exact user members, found $(@($rawMembers).Count) raw items ($(@($memberUpns).Count) user UPNs)."
            }
            $seenMemberUpnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($mu in @($memberUpns)) {
                $trimmedMu = "$mu".Trim()
                if ([string]::IsNullOrWhiteSpace($trimmedMu) -or -not $expectedMemberUpnSet.Contains($trimmedMu)) {
                    throw "Site '$Alias' group $GroupId has unexpected or extra member '$trimmedMu' (expected exact members: $(@($expectedMemberUpnSet) -join ', '))."
                }
                if (-not $seenMemberUpnSet.Add($trimmedMu)) {
                    throw "Site '$Alias' group $GroupId has duplicate member '$trimmedMu'."
                }
            }
            if ($null -ne $ExpectedUserIdsByUpn) {
                foreach ($rm in @($rawMembers)) {
                    if ($null -eq $rm -or -not ($rm.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($rm.id)") -or
                        -not ($rm.PSObject.Properties.Name -contains 'userPrincipalName') -or [string]::IsNullOrWhiteSpace("$($rm.userPrincipalName)")) {
                        throw "Site '$Alias' group $GroupId member verification failed: member entry is missing id or userPrincipalName."
                    }
                    $rmUpnKey = "$($rm.userPrincipalName)".Trim().ToLowerInvariant()
                    $rmId = "$($rm.id)".Trim()
                    if (-not $ExpectedUserIdsByUpn.ContainsKey($rmUpnKey) -or
                        -not [string]::Equals($rmId, "$($ExpectedUserIdsByUpn[$rmUpnKey])".Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw "Site '$Alias' group $GroupId member verification failed: member '$($rm.userPrincipalName)' id '$rmId' does not match resolved user ID."
                    }
                }
            }
        }
    }
}

# The team membership collection is the only place the Teams 'owner' ROLE exists — the group
# owner collection does not carry it. GET /teams/{id}/members is served by the
# TeamMember.ReadWrite.All application role that Phase 4's owner-membership POST already requires
# (documented in config.json.example):
# https://learn.microsoft.com/graph/api/team-list-members
function Get-SeedTeamMembers {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TeamId)

    try {
        $members = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/members"
    } catch {
        throw ("Cannot read the membership of team $TeamId, so the demo operator's Teams owner role cannot be " +
               "proven: $(Format-SeedSafeError -Message $_). GET /teams/{id}/members requires the " +
               "TeamMember.ReadWrite.All application permission (see the permission list in config.json.example; " +
               "the same role is what lets Phase 4 add the admin with the 'owner' role); grant and consent it, " +
               "then re-run. The seeder fails closed rather than writing into a team whose ownership is " +
               "unverified.")
    }
    return , @($members)
}

# ───────────────────────────────────────────────────────
# Teams
# ───────────────────────────────────────────────────────

function Get-SeedTeamChannels {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TeamId)

    $channels = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/channels"
    return , @($channels)
}

# Channel message reads are the one collection the seeder cannot always reach on v1.0: that
# endpoint requires ChannelMessage.Read.All / ChannelMessage.Read.Group, which the documented
# application permission set does not include, while the identical beta collection is served
# under Group.ReadWrite.All. The version is probed once per script and then reused, so a run
# issues exactly one 403 instead of one per request. Both variants are GET-only.
$script:SeedChannelApiVersion = $null

function Get-SeedChannelGraphCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PathAndQuery)

    if ($script:SeedChannelApiVersion) {
        $cached = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/$($script:SeedChannelApiVersion)$PathAndQuery"
        return , @($cached)
    }

    try {
        $items = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0$PathAndQuery"
        $script:SeedChannelApiVersion = 'v1.0'
        return , @($items)
    } catch {
        if (-not (Test-SeedChannelReadForbidden -ErrorRecord $_)) { throw }
        Write-Host ("      (channel read: v1.0 needs ChannelMessage.Read.*, which this app registration " +
                    "does not hold - falling back to the read-only beta collection)") -ForegroundColor DarkGray
        $items = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/beta$PathAndQuery"
        $script:SeedChannelApiVersion = 'beta'
        return , @($items)
    }
}

# Content fingerprint of one channel: the hashes of every top-level message plus, per parent,
# the hashes of its replies. Replies must be fetched separately for every parent because
# list-channel-messages never returns them.
function Get-SeedChannelMessageState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelId
    )

    $topHashes = @()
    $replyMap = @{}

    $messages = Get-SeedChannelGraphCollection -PathAndQuery "/teams/$TeamId/channels/$ChannelId/messages?`$top=50"
    foreach ($m in @($messages)) {
        if ($m.messageType -and $m.messageType -ne 'message') { continue }
        if (-not $m.body -or -not $m.body.content) { continue }

        $hash = Get-SeedContentHash -Content $m.body.content
        $topHashes += $hash

        $replies = Get-SeedChannelGraphCollection -PathAndQuery "/teams/$TeamId/channels/$ChannelId/messages/$($m.id)/replies?`$top=50"
        $replyHashes = @()
        foreach ($r in @($replies)) {
            if ($r.messageType -and $r.messageType -ne 'message') { continue }
            if (-not $r.body -or -not $r.body.content) { continue }
            $replyHashes += (Get-SeedContentHash -Content $r.body.content)
        }

        if ($replyMap.ContainsKey($hash)) {
            $replyMap[$hash] = @($replyMap[$hash]) + $replyHashes
        } else {
            $replyMap[$hash] = $replyHashes
        }
    }

    return [pscustomobject]@{ TopLevelHashes = $topHashes; ReplyHashes = $replyMap }
}
