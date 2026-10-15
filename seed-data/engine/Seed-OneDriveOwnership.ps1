<#
.SYNOPSIS
    Fail-closed ownership proof for generated, flat OneDrive demo folders.
.DESCRIPTION
    A local receipt is issued only after this process creates a new folder and uploads every
    declared file. Its random challenge is also stored in a remote proof file. A rerun compares
    both proofs, the folder ID, and every file ID, eTag and size before skipping all uploads.
    An unmarked, incomplete or changed folder is never adopted or repaired.
#>

. "$PSScriptRoot\Seed-GraphRead.ps1"

$script:SeedOneDriveProofName = '.ms4018-seed-proof.json'

function Get-SeedStrictOneDrivePlan {
    param([string]$ConfigPath, [string]$FilesManifestPath, [switch]$FreshOnly)

    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $manifest = Get-Content $FilesManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $folderName = "$($manifest.targetFolder)"
    if (-not $folderName -or $folderName -match '[\\/]' -or
        $manifest.uploadToRole -ne 'Admin' -or
        -not [string]::Equals("$($config.roles.Admin.upn)", "$($config.adminUpn)", [StringComparison]::OrdinalIgnoreCase)) {
        throw "Strict OneDrive seeding requires one neutral top-level folder and the verified Admin role."
    }
    $configDir = [IO.Path]::GetFullPath((Split-Path $ConfigPath -Parent))
    $scenarioDir = [IO.Path]::GetFullPath((Split-Path $FilesManifestPath -Parent))
    if (-not $config.filesSourceDir -or [IO.Path]::IsPathRooted("$($config.filesSourceDir)")) {
        throw "Strict OneDrive filesSourceDir must stay inside the scenario root."
    }
    $sourceDir = [IO.Path]::GetFullPath((Join-Path $configDir $config.filesSourceDir))
    if (-not $sourceDir.StartsWith($scenarioDir + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Strict OneDrive filesSourceDir must stay inside the scenario root."
    }
    $files = @()
    foreach ($entry in @($manifest.files)) {
        $name = if ($entry.PSObject.Properties.Name -contains 'localName' -and $entry.localName) {
            "$($entry.localName)"
        } elseif ($entry.PSObject.Properties.Name -contains 'fileName' -and $entry.fileName) {
            "$($entry.fileName)"
        } else {
            ''
        }
        $hasSubfolder = ($entry.PSObject.Properties.Name -contains 'subfolder') -and [bool]$entry.subfolder
        if (-not $name -or $name -in @('.', '..', $script:SeedOneDriveProofName) -or
            $name -match '[\\/]' -or $hasSubfolder) {
            throw "Strict OneDrive seeding requires flat, uniquely named files."
        }
        $path = Join-Path $sourceDir $name
        if (-not (Test-Path $path -PathType Leaf)) { throw "Declared demo file is missing: $path" }
        $files += [pscustomobject]@{
            Name = $name
            Path = $path
            Hash = (Get-FileHash $path -Algorithm SHA256).Hash
        }
    }
    if (-not $files.Count -or @($files | Group-Object { $_.Name.ToLowerInvariant() } | Where-Object Count -gt 1).Count) {
        throw "Strict OneDrive seeding requires a nonempty set of uniquely named files."
    }
    $identity = [ordered]@{
        Tenant = "$($config.tenantId)"
        Admin = "$($config.adminUpn)"
        Folder = $folderName
        Files = @($files | Sort-Object Name | ForEach-Object { [ordered]@{ Name = $_.Name; Hash = $_.Hash } })
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($identity | ConvertTo-Json -Depth 8 -Compress))
    return [pscustomobject]@{
        Admin = "$($config.adminUpn)"
        Folder = $folderName
        Files = $files
        Fingerprint = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
        ReceiptPath = Join-Path $scenarioDir 'onedrive-receipt.json'
        FreshOnly = [bool]$FreshOnly
    }
}

function Get-SeedStrictOneDriveChildren {
    param($Plan, [string]$FolderId)
    $admin = [Uri]::EscapeDataString($Plan.Admin)
    $folder = [Uri]::EscapeDataString($FolderId)
    $items = Get-SeedGraphCollection -Uri "https://graph.microsoft.com/v1.0/users/$admin/drive/items/$folder/children?`$select=id,name,eTag,size,file,folder&`$top=200"
    return $items
}

function Get-SeedStrictOneDriveState {
    param($Plan, [switch]$FreshOnly)

    $isFreshOnly = [bool]$FreshOnly -or ($Plan.PSObject.Properties.Name -contains 'FreshOnly' -and [bool]$Plan.FreshOnly)
    if ($isFreshOnly -and (Test-Path $Plan.ReceiptPath)) {
        throw "Collision (FreshOnly): local OneDrive receipt '$($Plan.ReceiptPath)' already exists; refusing to reuse."
    }

    $admin = [Uri]::EscapeDataString($Plan.Admin)
    $folderName = [Uri]::EscapeDataString($Plan.Folder)
    $uri = "https://graph.microsoft.com/v1.0/users/$admin/drive/root:/$folderName`?`$select=id,name,folder"
    try {
        $folder = Invoke-Graph -Method GET -Uri $uri
    } catch {
        if (-not (Test-SeedNotFoundError -ErrorRecord $_)) { throw }
        if (Test-Path $Plan.ReceiptPath) { throw "OneDrive receipt exists but its folder is missing; stop without writing." }
        return [pscustomobject]@{ Action = 'Create' }
    }
    if (-not $folder.id -or -not $folder.folder -or $folder.name -cne $Plan.Folder) {
        throw "OneDrive target is not the expected folder; stop without writing."
    }
    if ($isFreshOnly) {
        throw "Collision (FreshOnly): OneDrive target folder '$($Plan.Folder)' already exists (id=$($folder.id)). Fresh-target-only mode never reuses or adopts an existing folder."
    }
    if (-not (Test-Path $Plan.ReceiptPath -PathType Leaf)) {
        throw "Existing OneDrive folder has no local scenario receipt; never adopt or add files to it."
    }
    $receipt = Get-Content $Plan.ReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($receipt.version -ne 1 -or -not $receipt.nonce -or
        $receipt.folderId -cne $folder.id -or $receipt.fingerprint -cne $Plan.Fingerprint) {
        throw "OneDrive folder receipt does not match this scenario; stop without writing."
    }
    $children = @(Get-SeedStrictOneDriveChildren -Plan $Plan -FolderId $folder.id)
    if ($children.Count -ne ($Plan.Files.Count + 1) -or @($receipt.files).Count -ne $Plan.Files.Count) {
        throw "OneDrive folder content differs from the scenario receipt; stop without writing."
    }
    $proof = @($children | Where-Object { $_.name -ceq $script:SeedOneDriveProofName })
    if ($proof.Count -ne 1 -or -not $proof[0].id -or -not $proof[0].file) {
        throw "OneDrive scenario proof is missing or ambiguous; stop without writing."
    }
    $proofId = [Uri]::EscapeDataString("$($proof[0].id)")
    $remoteProof = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$admin/drive/items/$proofId/content"
    if ($remoteProof -is [array] -and $remoteProof.Count -gt 0 -and $remoteProof[0] -is [byte]) {
        $remoteProof = [byte[]]$remoteProof
    }
    if ($remoteProof -is [byte[]]) { $remoteProof = [Text.Encoding]::UTF8.GetString($remoteProof) }
    if ($remoteProof -is [string]) { $remoteProof = $remoteProof | ConvertFrom-Json }
    if ($remoteProof.nonce -cne $receipt.nonce -or
        $remoteProof.folderId -cne $folder.id -or
        $remoteProof.fingerprint -cne $Plan.Fingerprint) {
        throw "OneDrive scenario proof does not match the local receipt; stop without writing."
    }
    foreach ($expected in @($receipt.files)) {
        $item = @($children | Where-Object { $_.name -ceq $expected.name })
        if ($item.Count -ne 1 -or -not $item[0].file -or -not $item[0].id -or -not $item[0].eTag -or
            $item[0].id -cne $expected.id -or $item[0].eTag -cne $expected.eTag -or
            $item[0].size -ne $expected.size) {
            throw "OneDrive file '$($expected.name)' differs from the scenario receipt; stop without writing."
        }
    }
    return [pscustomobject]@{ Action = 'Skip'; FolderId = $folder.id }
}

function Send-SeedStrictOneDriveFile {
    param($Plan, [string]$FolderId, [string]$Name, [byte[]]$Bytes)

    $admin = [Uri]::EscapeDataString($Plan.Admin)
    $folder = [Uri]::EscapeDataString($FolderId)
    $fileName = [Uri]::EscapeDataString($Name)
    $uri = "https://graph.microsoft.com/v1.0/users/$admin/drive/items/$folder`:/$fileName`:/createUploadSession"
    $session = Invoke-Graph -Method POST -Uri $uri -Body @{ item = @{ '@microsoft.graph.conflictBehavior' = 'fail' } }
    if (-not $session.uploadUrl) { throw "No OneDrive upload session was returned for '$Name'." }
    $chunkSize = 5 * 327680
    for ($start = 0; $start -lt $Bytes.Length; $start += $chunkSize) {
        $length = [Math]::Min($chunkSize, $Bytes.Length - $start)
        $chunk = New-Object byte[] $length
        [Array]::Copy($Bytes, $start, $chunk, 0, $length)
        $last = $start + $length - 1
        try {
            if ($null -ne $Plan -and ($Plan.PSObject.Properties.Name -contains 'FreshOnly') -and $Plan.FreshOnly) {
                Assert-SeedNoWriteRetryContext -FreshOnly
            }
            Invoke-RestMethod -Method PUT -Uri $session.uploadUrl `
                -Headers @{ 'Content-Range' = "bytes $start-$last/$($Bytes.Length)" } `
                -ContentType 'application/octet-stream' -Body $chunk | Out-Null
        } catch {
            $safeUploadErr = Format-SeedSafeError -Message $_
            throw "OneDrive upload session failed for '$Name' ($safeUploadErr); stop and inspect the partial folder. No existing file was replaced."
        }
    }
}

function Invoke-SeedStrictOneDriveUpload {
    param($Plan, [switch]$FreshOnly)

    $state = Get-SeedStrictOneDriveState -Plan $Plan -FreshOnly:$FreshOnly
    if ($state.Action -eq 'Skip') {
        Write-Host "  SKIP: verified OneDrive scenario folder and all files; no uploads performed."
        return
    }
    $admin = [Uri]::EscapeDataString($Plan.Admin)
    $uri = "https://graph.microsoft.com/v1.0/users/$admin/drive/root/children"
    $folder = Invoke-Graph -Method POST -Uri $uri -Body @{
        name = $Plan.Folder
        folder = @{}
        '@microsoft.graph.conflictBehavior' = 'fail'
    }
    if (-not $folder.id -or $folder.name -cne $Plan.Folder -or -not $folder.folder) {
        throw "OneDrive did not return the exact new scenario folder; stop without uploading."
    }
    foreach ($file in $Plan.Files) {
        Send-SeedStrictOneDriveFile -Plan $Plan -FolderId $folder.id -Name $file.Name `
            -Bytes ([IO.File]::ReadAllBytes($file.Path))
    }
    $children = @(Get-SeedStrictOneDriveChildren -Plan $Plan -FolderId $folder.id)
    if ($children.Count -ne $Plan.Files.Count) {
        throw "OneDrive upload result has $($children.Count) files; expected $($Plan.Files.Count). Stop without claiming completion."
    }
    $items = @()
    foreach ($file in $Plan.Files) {
        $match = @($children | Where-Object { $_.name -ceq $file.Name })
        if ($match.Count -ne 1 -or -not $match[0].id -or -not $match[0].eTag -or -not $match[0].file) {
            throw "OneDrive uploaded file '$($file.Name)' could not be verified; stop without claiming completion."
        }
        $items += [ordered]@{ name = $file.Name; id = $match[0].id; eTag = $match[0].eTag; size = $match[0].size }
    }
    $nonce = [Guid]::NewGuid().ToString('N')
    $proof = [ordered]@{ nonce = $nonce; folderId = $folder.id; fingerprint = $Plan.Fingerprint }
    $proofBytes = [Text.Encoding]::UTF8.GetBytes(($proof | ConvertTo-Json -Compress))
    Send-SeedStrictOneDriveFile -Plan $Plan -FolderId $folder.id -Name $script:SeedOneDriveProofName -Bytes $proofBytes
    $receipt = [ordered]@{ version = 1; nonce = $nonce; folderId = $folder.id; fingerprint = $Plan.Fingerprint; files = $items }
    $receipt | ConvertTo-Json -Depth 8 | Set-Content $Plan.ReceiptPath -Encoding UTF8
    $verifyPlan = [pscustomobject]@{
        Admin       = $Plan.Admin
        Folder      = $Plan.Folder
        Files       = $Plan.Files
        Fingerprint = $Plan.Fingerprint
        ReceiptPath = $Plan.ReceiptPath
        FreshOnly   = $false
    }
    $verified = Get-SeedStrictOneDriveState -Plan $verifyPlan
    if ($FreshOnly -and (Test-Path -LiteralPath $Plan.ReceiptPath)) {
        Remove-Item -LiteralPath $Plan.ReceiptPath -Force -ErrorAction SilentlyContinue
    }
    if ($verified.Action -ne 'Skip') { throw "OneDrive scenario proof could not be verified." }
    Write-Host "  Created and verified OneDrive scenario folder with $($items.Count) files."
}
