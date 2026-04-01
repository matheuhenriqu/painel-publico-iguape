[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

$env:PUBLIC_WORKFLOW_MODE = '1'

& (Join-Path $scriptRoot 'sync.ps1')

& (Join-Path $scriptRoot 'build-public-contracts-dashboard.ps1') `
    -SourcePath (Join-Path $scriptRoot '..\storage\contracts.json') `
    -OutputPath (Join-Path $scriptRoot '..\data\contracts-dashboard.json')
