<#
.SYNOPSIS
    Scenario runner for ab730-automotive-20261015.
    Delegates to seed-data\engine\Invoke-AB730.ps1 in fresh-target-only mode.
    Default mode is preflight-only (read-only); pass -Live to execute additive fresh writes.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [Parameter(Mandatory)][string]$ManifestPath,
    [switch]$Live
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$engineRunner = Join-Path $PSScriptRoot '..\..\engine\Invoke-AB730.ps1'
if (-not (Test-Path -LiteralPath $engineRunner -PathType Leaf)) {
    throw "Engine runner not found at '$engineRunner'."
}

& $engineRunner `
    -ConfigPath $ConfigPath `
    -ManifestPath $ManifestPath `
    -ScenarioDir $PSScriptRoot `
    -Live:$Live
