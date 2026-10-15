#Requires -Version 7.0
<#
.SYNOPSIS
    Offline TDD suite for AB-730 fresh-target-only engine hardening and Invoke-AB730.ps1 runner.
    Invoked by Test-SeedEngine.ps1; uses its Pass/Fail/Assert-* harness counters.
#>

Write-Section 'AB-730: Seed-GraphRead collection normalization, MaxPages=20, and Inbox nextLink canonicalization'

. (Join-Path $EngineDir 'Seed-GraphRead.ps1')
$ab730GraphReadCmd = Get-Command Get-SeedGraphCollection -ErrorAction SilentlyContinue
Assert-True 'Get-SeedGraphCollection defaults MaxPages to 20' {
    $null -ne $ab730GraphReadCmd -and
    $ab730GraphReadCmd.Parameters.ContainsKey('MaxPages') -and
    ($ab730GraphReadCmd -is [System.Management.Automation.FunctionInfo]) -and
    ($ab730GraphReadCmd.Definition -match '\[int\]\$MaxPages\s*=\s*20')
} 'Get-SeedGraphCollection must default MaxPages to 20'

function Invoke-WithStubGraph {
    param(
        [Parameter(Mandatory)][scriptblock]$Handler,
        [Parameter(Mandatory)][scriptblock]$Body
    )
    $prev = Get-Item Function:\Invoke-Graph -ErrorAction SilentlyContinue
    try {
        Set-Item Function:\script:Invoke-Graph -Value $Handler
        & $Body
    } finally {
        if ($prev) {
            Set-Item Function:\script:Invoke-Graph -Value $prev.ScriptBlock
        } else {
            Remove-Item Function:\Invoke-Graph -ErrorAction SilentlyContinue
        }
    }
}

Assert-True 'Get-SeedGraphCollection: 0-element value array returns flat System.Array with Count=0 under StrictMode' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{ value = @() }
    } -Body {
        $res = Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$select=id'
        ($res -is [System.Array]) -and ($res.Count -eq 0)
    }
} '0-element response must return an empty array object with .Count = 0'

Assert-True 'Get-SeedGraphCollection: 1-element value array returns flat System.Array with Count=1 and non-array element' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'g-1' }) }
    } -Body {
        $res = Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$select=id'
        ($res -is [System.Array]) -and ($res.Count -eq 1) -and (-not ($res[0] -is [System.Array])) -and ($res[0].id -eq 'g-1')
    }
} '1-element response must return a 1-element array without nesting'

Assert-True 'Get-SeedGraphCollection: N-element value array across pages returns flat System.Array' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        if ($Uri -notmatch 'skiptoken') {
            return [pscustomobject]@{
                value             = @([pscustomobject]@{ id = 'g-1' })
                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups?$select=id&$skiptoken=p2'
            }
        }
        return [pscustomobject]@{
            value = @([pscustomobject]@{ id = 'g-2' }, [pscustomobject]@{ id = 'g-3' })
        }
    } -Body {
        $res = Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$select=id'
        ($res -is [System.Array]) -and ($res.Count -eq 3) -and ($res[0].id -eq 'g-1') -and ($res[2].id -eq 'g-3')
    }
} 'multi-page response must return flat 3-element array'

Assert-Throws 'Get-SeedGraphCollection: null response throws' {
    Invoke-WithStubGraph -Handler { param($Method, $Uri) return $null } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'null'

Assert-Throws 'Get-SeedGraphCollection: missing value property throws' {
    Invoke-WithStubGraph -Handler { param($Method, $Uri) return [pscustomobject]@{ id = 'not-a-collection' } } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'value'

Assert-Throws 'Get-SeedGraphCollection: null value property throws' {
    Invoke-WithStubGraph -Handler { param($Method, $Uri) return [pscustomobject]@{ value = $null } } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'value'

Assert-Throws 'Get-SeedGraphCollection: non-array value property throws' {
    Invoke-WithStubGraph -Handler { param($Method, $Uri) return [pscustomobject]@{ value = 'not-an-array' } } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'value'

Assert-Throws 'Get-SeedGraphCollection: explicit null @odata.nextLink throws' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{ value = @(); '@odata.nextLink' = $null }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'nextLink'

Assert-Throws 'Get-SeedGraphCollection: self-referencing @odata.nextLink throws' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'g-1' })
            '@odata.nextLink' = $Uri
        }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$select=id'
    }
} 'loop|self-referencing|visited'

Assert-Throws 'Get-SeedGraphCollection: multi-page cycle (A -> B -> A) throws' {
    $uriA = 'https://graph.microsoft.com/v1.0/groups?$select=id'
    $uriB = 'https://graph.microsoft.com/v1.0/groups?$select=id&$skiptoken=b'
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $next = if ($Uri -eq $uriA) { $uriB } else { $uriA }
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'g-1' })
            '@odata.nextLink' = $next
        }
    } -Body {
        Get-SeedGraphCollection -Uri $uriA
    }
} 'loop|cycle|visited'

Assert-Throws 'Get-SeedGraphCollection: foreign host in @odata.nextLink throws' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'g-1' })
            '@odata.nextLink' = 'https://evil.example.com/v1.0/groups?$skiptoken=p2'
        }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'host|boundary|graph\.microsoft\.com'

Assert-Throws 'Get-SeedGraphCollection: cross-resource path in @odata.nextLink throws' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'g-1' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users?$skiptoken=p2'
        }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'path|boundary'

Assert-Throws 'Get-SeedGraphCollection: 20-page cutoff throws when page 20 still returns @odata.nextLink' {
    $script:pageCounter = 0
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $script:pageCounter++
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = "g-$($script:pageCounter)" })
            '@odata.nextLink' = "https://graph.microsoft.com/v1.0/groups?`$skiptoken=page$($script:pageCounter + 1)"
        }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups'
    }
} 'more than 20 pages|MaxPages'

Assert-True 'Get-SeedGraphCollection: accepts live Exchange Inbox nextLink with /mailFolders(''inbox'')/messages for same user' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        if ($Uri -notmatch '\$skip=100') {
            return [pscustomobject]@{
                value             = @([pscustomobject]@{ id = 'm-1'; subject = 'Subj 1' })
                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders(''inbox'')/messages?$select=id,subject,conversationId,receivedDateTime&$top=100&$skip=100'
            }
        }
        return [pscustomobject]@{
            value = @([pscustomobject]@{ id = 'm-2'; subject = 'Subj 2' })
        }
    } -Body {
        $msgs = Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders/inbox/messages?$select=id,subject,conversationId,receivedDateTime&$top=100'
        ($msgs -is [System.Array]) -and ($msgs.Count -eq 2) -and ($msgs[0].id -eq 'm-1') -and ($msgs[1].id -eq 'm-2')
    }
} 'slash and key-predicate mailFolders(inbox) for the same user must be accepted as the same Inbox collection'

Assert-Throws 'Get-SeedGraphCollection: rejects mailFolders(''inbox'') nextLink when user UPN changes' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'm-1' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/other%40contoso.com/mailFolders(''inbox'')/messages?$skip=100'
        }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders/inbox/messages?$top=100'
    }
} 'path|boundary'

Assert-Throws 'Get-SeedGraphCollection: rejects mailFolders key-predicate nextLink for non-inbox folder' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'm-1' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders(''sentitems'')/messages?$skip=100'
        }
    } -Body {
        Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders/inbox/messages?$top=100'
    }
} 'path|boundary'

# ═══════════════════════════════════════════════════════
Write-Section 'AB-730: Fresh-target-only idempotency helpers and token/error guards'
# ═══════════════════════════════════════════════════════

Assert-Throws 'Get-SeedResourceAction -FreshOnly: rejects 1 existing match instead of reusing' {
    Get-SeedResourceAction -ResourceKind 'team' -ResourceKey '20261015-AB-730' -Existing @([pscustomobject]@{ id = 't-1' }) -FreshOnly
} 'already exists|collision|fresh'

Assert-Value 'Get-SeedResourceAction -FreshOnly: returns Create when 0 existing matches' 'Create' {
    Get-SeedResourceAction -ResourceKind 'team' -ResourceKey '20261015-AB-730' -Existing @() -FreshOnly
}

Assert-Throws 'Get-SeedEmailThreadAction -FreshOnly: rejects existing thread even when message count matches expected' {
    Get-SeedEmailThreadAction -ThreadKey '【專案啟動】2026 Q4 華東車型線束增產與 OTD 98% 目標對齊' -ExistingMessageCount 3 -ExpectedMessageCount 3 -FreshOnly
} 'already|collision|fresh'

Assert-Value 'Get-SeedEmailThreadAction -FreshOnly: returns Create when 0 existing messages' 'Create' {
    Get-SeedEmailThreadAction -ThreadKey '【專案啟動】2026 Q4 華東車型線束增產與 OTD 98% 目標對齊' -ExistingMessageCount 0 -ExpectedMessageCount 3 -FreshOnly
}

Assert-Throws 'Get-SeedSiteGroupVerdict -FreshOnly: rejects existing Unified group owned by admin' {
    $grp = [pscustomobject]@{ id = 'g-1'; displayName = 'AB-730 - Automotive Knowledge - 20261015'; groupTypes = @('Unified') }
    $verdict = Get-SeedSiteGroupVerdict -Alias 'ab730-automotive-20261015' -DeclaredDisplayName 'AB-730 - Automotive Knowledge - 20261015' -ExistingGroups @($grp) -OwnerUpns @('admin@contoso.com') -AdminUpn 'admin@contoso.com' -FreshOnly
    if ($verdict.Action -ne 'Abort') { throw "Expected Abort, got $($verdict.Action)" }
    throw $verdict.Reason
} 'already exists|collision|fresh'

Assert-True 'Format-SeedSafeError redacts Bearer tokens, tempauth, downloadUrl, and clientSecret values' {
    $raw = 'Failed request: Authorization: Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0aWQiOiIxIn0.sig url=https://contoso.sharepoint.com/_layouts/15/download.aspx?tempauth=SECRET_TEMPAUTH_123&e=1 "@microsoft.graph.downloadUrl":"https://contoso/dl?sig=abc" client_secret=SUPER_SECRET_VAL'
    $sanitized = Format-SeedSafeError -Message $raw -Secrets @('SUPER_SECRET_VAL')
    ($sanitized -notmatch 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9') -and
    ($sanitized -notmatch 'SECRET_TEMPAUTH_123') -and
    ($sanitized -notmatch 'https://contoso/dl\?sig=abc') -and
    ($sanitized -notmatch 'SUPER_SECRET_VAL')
} 'Format-SeedSafeError must scrub tokens, tempauth, downloadUrl, and secrets'

Assert-True 'RED-B3-1: Get-SeedGraphCollection rejects @odata.nextLink that widens $select to body before 2nd GET' {
    $calls = [System.Collections.Generic.List[string]]::new()
    $script:b3Threw = $false
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $calls.Add($Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'm-1'; subject = 'S1' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders(''inbox'')/messages?$select=id,subject,body&$top=100&$skip=100'
        }
    } -Body {
        try {
            [void](Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders/inbox/messages?$select=id,subject,conversationId,receivedDateTime&$top=100')
        } catch {
            $script:b3Threw = $true
        }
    }
    $script:b3Threw -and ($calls.Count -eq 1)
} 'must reject @odata.nextLink that alters $select before 2nd GET'

Assert-True 'RED-B3-2: Get-SeedGraphCollection rejects @odata.nextLink that injects $expand=attachments before 2nd GET' {
    $calls = [System.Collections.Generic.List[string]]::new()
    $script:b3Threw = $false
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $calls.Add($Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'm-1'; subject = 'S1' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders(''inbox'')/messages?$select=id,subject,conversationId,receivedDateTime&$expand=attachments&$top=100&$skip=100'
        }
    } -Body {
        try {
            [void](Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders/inbox/messages?$select=id,subject,conversationId,receivedDateTime&$top=100')
        } catch {
            $script:b3Threw = $true
        }
    }
    $script:b3Threw -and ($calls.Count -eq 1)
} 'must reject @odata.nextLink that adds $expand before 2nd GET'

Assert-True 'RED-B3-3: Get-SeedGraphCollection rejects @odata.nextLink that drops or alters $filter before 2nd GET' {
    $calls = [System.Collections.Generic.List[string]]::new()
    $script:b3Threw = $false
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $calls.Add($Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'g-1'; displayName = 'Other' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName&$skiptoken=page2'
        }
    } -Body {
        try {
            [void](Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName%20eq%20'20261015-AB-730'&`$select=id,displayName")
        } catch {
            $script:b3Threw = $true
        }
    }
    $script:b3Threw -and ($calls.Count -eq 1)
} 'must reject @odata.nextLink that drops or alters $filter before 2nd GET'

Assert-True 'RED-B3-4: Get-SeedGraphCollection rejects @odata.nextLink that mutates $top=100 to $top=100000 before 2nd GET' {
    $calls = [System.Collections.Generic.List[string]]::new()
    $script:b3TopThrew = $false
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $calls.Add($Uri)
        return [pscustomobject]@{
            value             = @([pscustomobject]@{ id = 'm-1'; subject = 'S1' })
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders(''inbox'')/messages?$select=id,subject,conversationId,receivedDateTime&$top=100000&$skip=100'
        }
    } -Body {
        try {
            [void](Get-SeedGraphCollection -Uri 'https://graph.microsoft.com/v1.0/users/admin%40contoso.com/mailFolders/inbox/messages?$select=id,subject,conversationId,receivedDateTime&$top=100')
        } catch {
            $script:b3TopThrew = $true
        }
    }
    $script:b3TopThrew -and ($calls.Count -eq 1)
} 'must reject @odata.nextLink that changes $top before 2nd GET'

Assert-Throws 'RED-B6-8: Assert-SeedSiteGroupRoles (-FreshOnly) rejects extra/unexpected group members even when declared roles are present' {
    $rolesObj = [pscustomobject]@{
        Admin   = 'admin@contoso.com'
        Sponsor = 'sponsor@contoso.com'
        Planner = 'planner@contoso.com'
    }
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value = @(
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' }
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-planner@contoso.com'; userPrincipalName = 'planner@contoso.com' }
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-foreign@contoso.com'; userPrincipalName = 'foreign@contoso.com' }
            )
        }
    } -Body {
        Assert-SeedSiteGroupRoles `
            -GroupId 'g-test-extra' `
            -RolesConfig $rolesObj `
            -OwnerRoles @('Admin') `
            -MemberRoles @('Sponsor', 'Planner') `
            -OwnerUpns @('admin@contoso.com') `
            -FreshOnly
    }
} 'member|owner|unexpected|extra|count|mismatch'

Assert-True 'RED-B6-9: Assert-SeedSiteGroupRoles legacy presence-only reuse allows extra group owners/members when declared roles are present' {
    $rolesObj = [pscustomobject]@{
        Admin   = 'admin@contoso.com'
        Sponsor = 'sponsor@contoso.com'
        Planner = 'planner@contoso.com'
    }
    $script:legacyReusePassed = $false
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{
            value = @(
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-sponsor@contoso.com'; userPrincipalName = 'sponsor@contoso.com' }
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-planner@contoso.com'; userPrincipalName = 'planner@contoso.com' }
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'uid-foreign@contoso.com'; userPrincipalName = 'foreign@contoso.com' }
            )
        }
    } -Body {
        Assert-SeedSiteGroupRoles `
            -GroupId 'g-test-legacy-reuse' `
            -RolesConfig $rolesObj `
            -OwnerRoles @('Admin') `
            -MemberRoles @('Sponsor', 'Planner') `
            -OwnerUpns @('admin@contoso.com', 'extra-owner@contoso.com')
        $script:legacyReusePassed = $true
    }
    $script:legacyReusePassed
} 'legacy presence-only Assert-SeedSiteGroupRoles must allow extra owners/members when declared roles are present'

Assert-True 'RED-B7-8: New-SeedSanitizedErrorRecord preserves typed HttpRequestException.StatusCode (503) so 503 with itemNotFound text never becomes 404 absence' {
    $req503 = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, 'https://graph.microsoft.com/v1.0/users/admin@contoso.com/drive/root/children')
    $req503.Content = [System.Net.Http.StringContent]::new('client_secret=stub-secret-value-not-real')
    $resp503 = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::ServiceUnavailable)
    $resp503.RequestMessage = $req503
    $resp503.Content = [System.Net.Http.StringContent]::new('{"error":{"code":"serviceNotAvailable","message":"Backend replica itemNotFound tempauth=SECRET_CAP_TOKEN"}}')
    $ex = [Microsoft.PowerShell.Commands.HttpResponseException]::new(
        'Response status code does not indicate success: 503 (Service Unavailable) itemNotFound Authorization: ******',
        $resp503
    )
    $ex.Data['RawSecret'] = 'stub-secret-value-not-real'
    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'ServiceUnavailable', [System.Management.Automation.ErrorCategory]::ConnectionError, $resp503)
    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"serviceNotAvailable","message":"Backend replica itemNotFound tempauth=SECRET_CAP_TOKEN"}}')

    $beforeIsNotFound = Test-SeedNotFoundError -ErrorMessage $er
    $sanEr = New-SeedSanitizedErrorRecord -ErrorRecord $er -Operation 'Graph GET https://graph.microsoft.com/v1.0/users/admin@contoso.com/drive/root/children failed' -Secrets @('stub-secret-value-not-real')
    $afterIsNotFound = Test-SeedNotFoundError -ErrorMessage $sanEr
    $typedStatus = $sanEr.Exception.StatusCode
    $hasResponseRef = ($sanEr.Exception.PSObject.Properties.Name -contains 'Response') -and ($null -ne $sanEr.Exception.Response)

    # Also verify 429 Retry-After scalar extraction into Data without retaining Response/Headers reference
    $resp429 = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::TooManyRequests)
    $resp429.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([TimeSpan]::FromSeconds(17))
    $resp429.Content = [System.Net.Http.StringContent]::new('{"error":{"code":"activityLimitReached","message":"Throttled tempauth=SECRET_CAP_TOKEN"}}')
    $ex429 = [Microsoft.PowerShell.Commands.HttpResponseException]::new('Response status code does not indicate success: 429 (Too Many Requests).', $resp429)
    $er429 = [System.Management.Automation.ErrorRecord]::new($ex429, 'TooManyRequests', [System.Management.Automation.ErrorCategory]::LimitsExceeded, $resp429)
    $san429 = New-SeedSanitizedErrorRecord -ErrorRecord $er429 -Operation 'Graph GET mailbox failed'
    $has429ResponseRef = ($san429.Exception.PSObject.Properties.Name -contains 'Response') -and ($null -ne $san429.Exception.Response)
    $retryScalar = if ($san429.Exception.Data -and $san429.Exception.Data.Contains('SeedRetryAfterSeconds')) { [int]$san429.Exception.Data['SeedRetryAfterSeconds'] } else { -1 }

    (-not $beforeIsNotFound) -and
    (-not $afterIsNotFound) -and
    ($typedStatus -eq [System.Net.HttpStatusCode]::ServiceUnavailable) -and
    (-not $hasResponseRef) -and
    ($null -eq $sanEr.TargetObject) -and
    ($null -eq $sanEr.Exception.InnerException) -and
    ($sanEr.Exception.Data.Count -eq 0) -and
    (-not $has429ResponseRef) -and
    ($retryScalar -eq 17) -and
    ("$($sanEr.Exception.Message)" -notmatch 'eyJSECRET_JWT|SECRET_CAP_TOKEN|stub-secret-value-not-real')
} 'New-SeedSanitizedErrorRecord must preserve typed StatusCode=503, extract bounded RetryAfter scalar without retaining Response/RequestMessage/Content/Headers/TargetObject/Data, and keep Test-SeedNotFoundError=false'

Assert-True 'RED-B7-6: Assert-SeedNoWriteRetryContext enforces process-local Set-MgRequestContext -MaxRetry 0 and verifies Get-MgRequestContext (restored in finally)' {
    if (-not (Get-Command -Name 'Assert-SeedNoWriteRetryContext' -ErrorAction SilentlyContinue)) {
        return $false
    }
    $origCtx = Get-MgRequestContext
    $origMaxRetry = if ($null -ne $origCtx -and ($origCtx.PSObject.Properties.Name -contains 'MaxRetry')) { [int]$origCtx.MaxRetry } else { 3 }
    try {
        Set-MgRequestContext -MaxRetry 3 | Out-Null
        $beforeVerify = [int](Get-MgRequestContext).MaxRetry
        Assert-SeedNoWriteRetryContext -FreshOnly
        $afterVerify = [int](Get-MgRequestContext).MaxRetry
        return ($beforeVerify -eq 3) -and ($afterVerify -eq 0)
    } finally {
        Set-MgRequestContext -MaxRetry $origMaxRetry | Out-Null
    }
} 'Assert-SeedNoWriteRetryContext must set process-local SDK MaxRetry=0 from 3 and verify via Get-MgRequestContext'

Assert-Throws 'RED-B4-1: Get-SeedDriveChildNames fails closed when page 2 throws 404 instead of dropping page 1 items' {
    $page = 0
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        $script:page++
        if ($script:page -eq 1) {
            return [pscustomobject]@{
                value             = @([pscustomobject]@{ name = '01-BOM-Cost.xlsx' })
                '@odata.nextLink' = "$Uri&`$skiptoken=page2"
            }
        }
        $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 404 (Not Found).')
        $er = [System.Management.Automation.ErrorRecord]::new($ex, 'NotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $null)
        $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"itemNotFound","message":"Not found on page 2"}}')
        throw $er
    } -Body {
        $script:page = 0
        [void](Get-SeedDriveChildNames -ChildrenUri 'https://graph.microsoft.com/v1.0/drives/drv-1/root/children?$select=id,name&$top=200')
    }
} '404|itemNotFound|Not Found|page'

Assert-Throws 'RED-B4-3: Get-SeedRecentInboxMessages rejects null response or non-array value property' {
    Invoke-WithStubGraph -Handler {
        param($Method, $Uri)
        return [pscustomobject]@{ value = 'not-an-array' }
    } -Body {
        [void](Get-SeedRecentInboxMessages -Upn 'admin@contoso.com' -Top 10)
    }
} 'array|value|Invalid'

Assert-True 'RED-B4-4: Test-SeedNotFoundError returns false when HTTP 503 response body mentions itemNotFound' {
    $ex = [System.Net.Http.HttpRequestException]::new('Response status code does not indicate success: 503 (Service Unavailable).')
    $er = [System.Management.Automation.ErrorRecord]::new($ex, 'ServiceUnavailable', [System.Management.Automation.ErrorCategory]::ConnectionError, $null)
    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":{"code":"serviceNotAvailable","message":"Upstream replica returned itemNotFound during failover"}}')
    -not (Test-SeedNotFoundError -ErrorRecord $er)
} 'Test-SeedNotFoundError must return false for HTTP 503 even if message contains itemNotFound'

# ═══════════════════════════════════════════════════════
Write-Section 'AB-730: Offline Graph probes (later-page collisions, owning-phase race, surface contracts, Invoke-AB730 runner)'
# ═══════════════════════════════════════════════════════

$ab730RunnerPath = Join-Path $EngineDir 'Invoke-AB730.ps1'
$ab730TemplatePath = Join-Path $EngineDir 'Invoke-AB730ScenarioRun.template.ps1'

$allowedSyntheticUpns = @(
    'admin@example.invalid',
    'itlead@example.invalid',
    'itengineer1@example.invalid',
    'itengineer2@example.invalid',
    'employee1@example.invalid',
    'employee2@example.invalid',
    'employee3@example.invalid'
)
Assert-True 'Privacy guard: New-PL7008ReferenceFixture.ps1 source contains only @example.invalid synthetic UPNs' {
    $helperSrc = Get-Content -LiteralPath (Join-Path $TestsDir 'New-PL7008ReferenceFixture.ps1') -Raw -Encoding UTF8
    $emailMatches = [regex]::Matches($helperSrc, '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')
    if ($emailMatches.Count -eq 0) { return $false }
    $unexpected = @($emailMatches | ForEach-Object { $_.Value.ToLowerInvariant() } | Where-Object { $allowedSyntheticUpns -notcontains $_ })
    $unexpected.Count -eq 0
} 'New-PL7008ReferenceFixture.ps1 must use only neutral @example.invalid role UPNs'

Assert-True 'Privacy guard: synthesized PL-7008 fixture outputs contain only @example.invalid synthetic UPNs' {
    $unexpected = @()
    foreach ($relFile in @('config.json.example', 'user-profiles.json')) {
        $text = Get-Content -LiteralPath (Join-Path $ScenarioDir $relFile) -Raw -Encoding UTF8
        $matchesFound = [regex]::Matches($text, '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')
        foreach ($m in $matchesFound) {
            $val = $m.Value.ToLowerInvariant()
            if ($allowedSyntheticUpns -notcontains $val) { $unexpected += $relFile }
        }
    }
    $unexpected.Count -eq 0
} 'synthesized fixture config.json.example and user-profiles.json must use only neutral @example.invalid role UPNs'

Assert-True 'engine/Invoke-AB730.ps1 exists' {
    Test-Path -LiteralPath $ab730RunnerPath
} 'Invoke-AB730.ps1 must exist in seed-data/engine'

Assert-True 'engine/Invoke-AB730ScenarioRun.template.ps1 exists' {
    Test-Path -LiteralPath $ab730TemplatePath
} 'Invoke-AB730ScenarioRun.template.ps1 must exist in seed-data/engine'

$ab730ProbeScript = Join-Path $TestsDir 'Invoke-AB730GraphProbe.ps1'
Assert-True 'tests/Invoke-AB730GraphProbe.ps1 exists' {
    Test-Path -LiteralPath $ab730ProbeScript
} 'Invoke-AB730GraphProbe.ps1 must exist in seed-data/tests'

if (Test-Path -LiteralPath $ab730ProbeScript) {
    . $ab730ProbeScript
}
