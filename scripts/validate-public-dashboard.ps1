[CmdletBinding()]
param(
    [string]$DashboardDataPath = '',
    [string]$IndexPath = '',
    [string]$ScriptPath = '',
    [string]$StylesPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

if ([string]::IsNullOrWhiteSpace($DashboardDataPath)) {
    $DashboardDataPath = Join-Path $scriptRoot '..\data\contracts-dashboard.json'
}
if ([string]::IsNullOrWhiteSpace($IndexPath)) {
    $IndexPath = Join-Path $scriptRoot '..\index.html'
}
if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    $ScriptPath = Join-Path $scriptRoot '..\app.js'
}
if ([string]::IsNullOrWhiteSpace($StylesPath)) {
    $StylesPath = Join-Path $scriptRoot '..\styles.css'
}

$allowedHosts = @(
    'iguape.sp.gov.br',
    'www.iguape.sp.gov.br'
)

function Assert-Condition {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Condition,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Assert-Condition -Condition (Test-Path -LiteralPath $Path -PathType Leaf) -Message "Required file missing: $Path"
}

function Get-JsonValue {
    param(
        [AllowNull()]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($null -eq $Item) {
        return $null
    }

    if ($Item.PSObject -and $Item.PSObject.Properties[$PropertyName]) {
        return $Item.$PropertyName
    }

    return $null
}

function Get-ExternalUrls {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Records
    )

    foreach ($record in $Records) {
        $links = Get-JsonValue -Item $record -PropertyName 'links'
        if ($null -eq $links) {
            continue
        }

        foreach ($name in @('diary', 'portal')) {
            $value = Get-JsonValue -Item $links -PropertyName $name
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                [pscustomobject]@{
                    LinkName = $name
                    Value = [string]$value
                }
            }
        }
    }
}

function Assert-NoBrokenStaticText {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Paths
    )

    $brokenMarkers = @(
        [string][char]0x00C3,
        [string][char]0x00C2,
        [string][char]0x252C,
        [string][char]0xFFFD
    )

    foreach ($path in $Paths) {
        $content = Get-Content -LiteralPath $path -Raw -Encoding utf8
        foreach ($marker in $brokenMarkers) {
            Assert-Condition -Condition (-not $content.Contains($marker)) -Message "Broken text marker found in: $path"
        }
    }
}

Assert-FileExists -Path $DashboardDataPath
Assert-FileExists -Path $IndexPath
Assert-FileExists -Path $ScriptPath
Assert-FileExists -Path $StylesPath

$dashboard = Get-Content -LiteralPath $DashboardDataPath -Raw -Encoding utf8 | ConvertFrom-Json
$records = @($dashboard.records)
$summary = $dashboard.summary
$filters = $dashboard.filters

Assert-Condition -Condition ($null -ne $summary) -Message 'Dashboard summary missing.'
Assert-Condition -Condition ($null -ne $filters) -Message 'Dashboard filters missing.'
Assert-Condition -Condition ($records.Count -gt 0) -Message 'No contracts were generated for the public dashboard.'
Assert-Condition -Condition ($null -ne (Get-JsonValue -Item $dashboard -PropertyName 'generatedAt')) -Message 'generatedAt is missing.'

$generatedAt = [DateTime]::MinValue
Assert-Condition -Condition ([DateTime]::TryParse([string]$dashboard.generatedAt, [ref]$generatedAt)) -Message 'generatedAt is invalid.'

$currentRecords = @(
    $records | Where-Object {
        Get-JsonValue -Item (Get-JsonValue -Item $_ -PropertyName 'vigency') -PropertyName 'isCurrent'
    }
)

Assert-Condition -Condition ($currentRecords.Count -gt 0) -Message 'No current contracts were generated for the public dashboard.'
Assert-Condition -Condition ([int]$summary.contratosAtuais -eq $currentRecords.Count) -Message 'Summary mismatch: contratosAtuais.'
Assert-Condition -Condition ([int]$summary.semGestor -eq @($currentRecords | Where-Object { $_.managementState -eq 'sem_gestor' -or $_.managementState -eq 'sem_gestor_e_fiscal' }).Count) -Message 'Summary mismatch: semGestor.'
Assert-Condition -Condition ([int]$summary.semFiscal -eq @($currentRecords | Where-Object { $_.managementState -eq 'sem_fiscal' -or $_.managementState -eq 'sem_gestor_e_fiscal' }).Count) -Message 'Summary mismatch: semFiscal.'
Assert-Condition -Condition ([int]$summary.semGestorEFiscal -eq @($currentRecords | Where-Object { $_.managementState -eq 'sem_gestor_e_fiscal' }).Count) -Message 'Summary mismatch: semGestorEFiscal.'
Assert-Condition -Condition ([int]$summary.comResponsaveisCompletos -eq @($currentRecords | Where-Object { $_.managementState -eq 'completos' }).Count) -Message 'Summary mismatch: comResponsaveisCompletos.'
Assert-Condition -Condition ([int]$summary.somenteDiario -eq @($currentRecords | Where-Object { $_.sourceStatus -eq 'somente_diario' }).Count) -Message 'Summary mismatch: somenteDiario.'
Assert-Condition -Condition ([int]$summary.somentePortal -eq @($currentRecords | Where-Object { $_.sourceStatus -eq 'somente_portal' }).Count) -Message 'Summary mismatch: somentePortal.'
Assert-Condition -Condition ([int]$summary.cruzados -eq @($currentRecords | Where-Object { $_.sourceStatus -eq 'cruzado' }).Count) -Message 'Summary mismatch: cruzados.'

$externalUrls = @(Get-ExternalUrls -Records $records)
foreach ($entry in $externalUrls) {
    $uri = [Uri]$entry.Value
    Assert-Condition -Condition ($uri.Scheme -eq 'https') -Message "Insecure external link found: $($entry.Value)"
    Assert-Condition -Condition ($allowedHosts -contains $uri.Host.ToLowerInvariant()) -Message "Unexpected external host found: $($entry.Value)"
}

$indexContent = Get-Content -LiteralPath $IndexPath -Raw -Encoding utf8
$scriptContent = Get-Content -LiteralPath $ScriptPath -Raw -Encoding utf8
$stylesContent = Get-Content -LiteralPath $StylesPath -Raw -Encoding utf8

Assert-Condition -Condition ($indexContent -match 'Content-Security-Policy') -Message 'index.html is missing Content-Security-Policy.'
Assert-Condition -Condition ($indexContent -match 'trusted-types contracts-dashboard') -Message 'index.html is missing Trusted Types.'
Assert-Condition -Condition ($indexContent -match 'app\.js') -Message 'index.html does not reference app.js.'
Assert-Condition -Condition ($indexContent -match 'styles\.css') -Message 'index.html does not reference styles.css.'
Assert-Condition -Condition ($scriptContent -match 'fetchDashboardPayload') -Message 'app.js is missing protected dashboard loading.'
Assert-Condition -Condition ($scriptContent -match 'sanitizeExternalUrl') -Message 'app.js is missing external link sanitization.'
Assert-Condition -Condition ($stylesContent -match '--shell-width') -Message 'styles.css is missing the responsive layout baseline.'

Assert-NoBrokenStaticText -Paths @($IndexPath, $StylesPath)

Write-Host 'Validation completed successfully.'
