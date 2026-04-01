[CmdletBinding()]
param(
    [string]$SourcePath = '',
    [string]$OutputPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Join-Path $scriptRoot '..\storage\contracts.json'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $scriptRoot '..\data\contracts-dashboard.json'
}

. (Join-Path $scriptRoot 'common.ps1')

function Get-CleanText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return (Collapse-Whitespace -Text ([string]$Value)).Trim()
}

function Convert-ToIsoString {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [DateTime]) {
        return ([DateTime]$Value).ToString('s')
    }

    $text = Get-CleanText -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $brazilian = Convert-BrazilianDateToDateTime -Text $text
    if ($brazilian) {
        return $brazilian.ToString('s')
    }

    try {
        return ([DateTime]::Parse($text)).ToString('s')
    }
    catch {
        return $null
    }
}

function Convert-ToDateTimeSafe {
    param(
        [AllowNull()]
        [object]$Value
    )

    $iso = Convert-ToIsoString -Value $Value
    if ([string]::IsNullOrWhiteSpace($iso)) {
        return $null
    }

    return [DateTime]::Parse($iso)
}

function Get-ObjectValue {
    param(
        [AllowNull()]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [object]$Default = $null
    )

    if ($null -eq $Item) {
        return $Default
    }

    if ($Item -is [System.Collections.IDictionary] -and $Item.Contains($Name)) {
        return $Item[$Name]
    }

    if ($Item.PSObject -and $Item.PSObject.Properties[$Name]) {
        return $Item.$Name
    }

    return $Default
}

function Get-NormalizedContractKey {
    param(
        [AllowNull()]
        [object]$Value
    )

    $text = Get-CleanText -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    if ($text -match '(?<number>\d{1,9})\s*[/-]\s*(?<year>\d{2,4})') {
        $number = [int]$matches['number']
        $year = [string]$matches['year']
        if ($year.Length -eq 2) {
            $year = "20$year"
        }
        return ('{0}/{1}' -f $number, $year)
    }

    return (Normalize-IndexText -Text $text).ToUpperInvariant()
}

function Get-ContractYear {
    param(
        [AllowNull()]
        [string]$NormalizedKey,

        [AllowNull()]
        [object[]]$Movements = @(),

        [AllowNull()]
        [object]$OfficialContract = $null
    )

    if (-not [string]::IsNullOrWhiteSpace($NormalizedKey) -and $NormalizedKey -match '/(?<year>\d{4})$') {
        return [int]$matches['year']
    }

    $candidateDates = New-Object System.Collections.ArrayList
    if ($OfficialContract) {
        foreach ($value in @((Get-ObjectValue -Item $OfficialContract -Name 'signatureDate'), (Get-ObjectValue -Item $OfficialContract -Name 'publishedAt'))) {
            $date = Convert-ToDateTimeSafe -Value $value
            if ($date) {
                [void]$candidateDates.Add($date)
            }
        }
    }

    foreach ($movement in @($Movements)) {
        foreach ($value in @((Get-ObjectValue -Item $movement -Name 'signatureDate'), (Get-ObjectValue -Item $movement -Name 'publishedAt'))) {
            $date = Convert-ToDateTimeSafe -Value $value
            if ($date) {
                [void]$candidateDates.Add($date)
            }
        }
    }

    if ($candidateDates.Count -gt 0) {
        return ($candidateDates | Sort-Object | Select-Object -First 1).Year
    }

    return $null
}

function Get-AdministrationLabel {
    param(
        [AllowNull()]
        [int]$Year
    )

    if ($null -eq $Year) {
        return 'Período não identificado'
    }

    if ($Year -ge 2025) {
        return 'Gestão 2025-2028'
    }
    if ($Year -ge 2021) {
        return 'Gestão 2021-2024'
    }
    if ($Year -ge 2017) {
        return 'Gestão 2017-2020'
    }

    return 'Até 2016'
}

function Get-PersonSnapshot {
    param(
        [AllowNull()]
        [string]$Name,

        [AllowNull()]
        [string]$Role,

        [AllowNull()]
        [object]$AssignedAt
    )

    $cleanName = Get-CleanText -Value $Name
    $cleanRole = Get-CleanText -Value $Role
    $needsReview = $false

    if (-not [string]::IsNullOrWhiteSpace($cleanName) -and (Test-SuspiciousManagementAssignment -Name $cleanName -Role $cleanRole)) {
        $cleanName = ''
        $cleanRole = ''
        $needsReview = $true
    }

    return [ordered]@{
        name = $cleanName
        role = $cleanRole
        assignedAt = Convert-ToIsoString -Value $AssignedAt
        needsReview = [bool]$needsReview
    }
}

function Test-UsablePersonName {
    param(
        [AllowNull()]
        [string]$Name
    )

    $clean = Get-CleanText -Value $Name
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $false
    }

    $normalized = Normalize-IndexText -Text $clean
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $false
    }

    if ($normalized -match '\b(PORTARIA|DECRETO|RESOLUCAO|RESOLUCAO|ATO|EXTRATO|CONTRATO|PROCESSO|PREFEITO|SECRETARIA|GESTOR|FISCAL|ART)\b') {
        return $false
    }

    if ($normalized -match '\d') {
        return $false
    }

    $tokens = @($clean -split '\s+' | Where-Object { $_ })
    if ($tokens.Count -lt 2 -or $tokens.Count -gt 8) {
        return $false
    }

    $lastToken = (Normalize-IndexText -Text ([string]$tokens[-1])).ToUpperInvariant()
    if ($lastToken -in @('DE', 'DA', 'DO', 'DOS', 'DAS', 'E')) {
        return $false
    }

    return $tokens[-1].Trim().Length -ge 2
}

function Convert-ToPublicPdfAbsolutePath {
    param(
        [AllowNull()]
        [string]$RelativePath
    )

    $clean = Get-CleanText -Value $RelativePath
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ''
    }

    if ([System.IO.Path]::IsPathRooted($clean)) {
        return $clean
    }

    return (Join-Path $script:AppRoot ($clean -replace '/', '\'))
}

function Get-PublicPdfPages {
    param(
        [AllowNull()]
        [string]$PdfPath
    )

    $cleanPath = Get-CleanText -Value $PdfPath
    if ([string]::IsNullOrWhiteSpace($cleanPath) -or -not $script:PublicPdfToTextTool -or -not (Test-Path -LiteralPath $cleanPath)) {
        return @()
    }

    if ($script:PublicPdfPageCache.ContainsKey($cleanPath)) {
        return @($script:PublicPdfPageCache[$cleanPath])
    }

    $rawText = & $script:PublicPdfToTextTool -enc UTF-8 -layout $cleanPath - 2>$null
    $joined = ($rawText | Out-String)
    if ([string]::IsNullOrWhiteSpace($joined)) {
        $script:PublicPdfPageCache[$cleanPath] = @()
        return @()
    }

    $pages = New-Object System.Collections.ArrayList
    $pageNumber = 1

    foreach ($page in @($joined -split [char]12)) {
        $text = [string]$page
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            [void]$pages.Add([pscustomobject]@{
                pageNumber = $pageNumber
                text = $text
            })
        }
        $pageNumber++
    }

    $script:PublicPdfPageCache[$cleanPath] = @($pages)
    return @($script:PublicPdfPageCache[$cleanPath])
}

function Get-ManagementTextWindowFromMovement {
    param(
        [AllowNull()]
        [object]$Movement
    )

    if ($null -eq $Movement) {
        return ''
    }

    $cacheKey = '{0}|{1}|{2}' -f `
        (Get-CleanText -Value (Get-ObjectValue -Item $Movement -Name 'localPdfRelative')), `
        (Get-CleanText -Value (Get-ObjectValue -Item $Movement -Name 'pageNumber')), `
        (Get-CleanText -Value (Get-ObjectValue -Item $Movement -Name 'contractNumber'))

    if ($script:PublicManagementTextCache.ContainsKey($cacheKey)) {
        return [string]$script:PublicManagementTextCache[$cacheKey]
    }

    $pdfPath = Convert-ToPublicPdfAbsolutePath -RelativePath (Get-ObjectValue -Item $Movement -Name 'localPdfRelative')
    $pageNumber = [int]$(if ($null -ne (Get-ObjectValue -Item $Movement -Name 'pageNumber')) { (Get-ObjectValue -Item $Movement -Name 'pageNumber') } else { 0 })
    $pageText = Get-CleanText -Value $(
        Get-ObjectValue -Item `
            (@(Get-PublicPdfPages -PdfPath $pdfPath | Where-Object { [int]$_.pageNumber -eq $pageNumber }) | Select-Object -First 1) `
            -Name 'text'
    )

    if ([string]::IsNullOrWhiteSpace($pageText)) {
        $script:PublicManagementTextCache[$cacheKey] = ''
        return ''
    }

    $window = ''
    $actTitle = Get-CleanText -Value (Get-ObjectValue -Item $Movement -Name 'actTitle')
    if (-not [string]::IsNullOrWhiteSpace($actTitle) -and $actTitle -match '^(PORTARIA|DECRETO|RESOLU)') {
        $pattern = '(?is).{0,80}' + [regex]::Escape($actTitle) + '.{0,1600}'
        $match = [regex]::Match($pageText, $pattern)
        if ($match.Success) {
            $window = [string]$match.Value
        }
    }

    if ([string]::IsNullOrWhiteSpace($window)) {
        $normalizedKey = Get-NormalizedContractKey -Value (Get-ObjectValue -Item $Movement -Name 'contractNumber')
        if (-not [string]::IsNullOrWhiteSpace($normalizedKey) -and $normalizedKey -match '^(?<number>\d+)/(?<year>\d{4})$') {
            $number = [int]$matches['number']
            $year = [string]$matches['year']
            $pattern = '(?is).{0,500}(?:CONTRATO\s*(?:N\S{0,2}\s*)?)?0*' + $number + '\s*/\s*' + $year + '.{0,1200}'
            $match = [regex]::Match($pageText, $pattern)
            if ($match.Success) {
                $window = [string]$match.Value
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($window)) {
        $window = $pageText
    }

    $script:PublicManagementTextCache[$cacheKey] = $window
    return $window
}

function New-ManagementPersonFromMatch {
    param(
        [AllowNull()]
        [System.Text.RegularExpressions.Match]$Match
    )

    if ($null -eq $Match -or -not $Match.Success) {
        return $null
    }

    $name = Get-CleanText -Value $Match.Groups['name'].Value
    $role = Get-CleanText -Value $Match.Groups['role'].Value
    if ($role) {
        $role = $role -replace ',?\s*(?:inscrit[oa].*|titular da.*|portador.*|cpf.*|rg.*)$', ''
        $role = (Get-CleanText -Value $role).Trim(" -,:;.")
    }

    if (-not (Test-UsablePersonName -Name $name)) {
        return $null
    }

    return [ordered]@{
        name = $name
        role = $role
    }
}

function Get-ManagementPersonFromText {
    param(
        [AllowNull()]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [ValidateSet('manager', 'inspector')]
        [string]$RoleType
    )

    $compact = Get-CleanText -Value $Text
    if ([string]::IsNullOrWhiteSpace($compact)) {
        return $null
    }

    $patterns = switch ($RoleType) {
        'manager' {
            @(
                '(?is)(?:FICAM?\s+DESIGNADOS?\s+|FICA\s+DESIGNADO\s+|DESIGNA,\s*)(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|,\s*PORTADOR|\s+CPF|\s+RG)).{0,260}?\bPARA\s+EXERCER\s+A?\s*FUN\S{0,8}\s+DE\s+GESTOR\b'
            )
        }
        default {
            @(
                '(?is)\bGESTOR\b.{0,160}?\bCONTRATO\b.{0,80}?,\s*E\s+(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|,\s*PORTADOR|\s+CPF|\s+RG)).{0,260}?\bPARA\s+(?:ATUAR\s+COMO|EXERCER\s+A?\s*FUN\S{0,8}\s+DE)\s+FISCAL\b',
                '(?is)(?:FICA\s+DESIGNADO\s+|DESIGNA,\s*)(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|,\s*PORTADOR|\s+CPF|\s+RG)).{0,180}?\bPARA\s+EXERCER\s+A?\s*FUN\S{0,8}\s+DE\s+FISCAL\b'
            )
        }
    }

    $candidateTexts = @($compact)
    if ($RoleType -eq 'inspector') {
        $fiscalIndex = $compact.IndexOf('Fiscal', [System.StringComparison]::OrdinalIgnoreCase)
        if ($fiscalIndex -ge 0) {
            $start = [Math]::Max(0, $fiscalIndex - 520)
            $length = [Math]::Min($compact.Length - $start, 920)
            $localized = $compact.Substring($start, $length)
            if (-not [string]::IsNullOrWhiteSpace($localized)) {
                $candidateTexts = @($localized, $compact)
            }
        }
    }

    foreach ($candidateText in @($candidateTexts)) {
        foreach ($pattern in @($patterns)) {
            $person = New-ManagementPersonFromMatch -Match ([regex]::Match(
                    $candidateText,
                    $pattern,
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
                    [System.Text.RegularExpressions.RegexOptions]::Singleline
                ))

            if ($person) {
                return $person
            }
        }
    }

    return $null
}

function Get-RepairedAssignmentsFromMovement {
    param(
        [AllowNull()]
        [object]$Movement
    )

    if ($null -eq $Movement) {
        return [ordered]@{
            manager = $null
            inspector = $null
        }
    }

    $scopeText = (
        @(
            (Get-ObjectValue -Item $Movement -Name 'actTitle'),
            (Get-ObjectValue -Item $Movement -Name 'excerpt')
        ) |
        ForEach-Object { Get-CleanText -Value $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    ) -join ' '
    $normalizedScope = Normalize-IndexText -Text $scopeText
    $allowsManager = $true
    $allowsInspector = $true
    $hasManagerScope = $normalizedScope -match 'GESTOR E FISCAL DO CONTRATO|GESTOR DO CONTRATO|FUNCAO DE GESTOR'
    $hasInspectorScope = $normalizedScope -match 'GESTOR E FISCAL DO CONTRATO|FISCAL DO CONTRATO|FUNCAO DE FISCAL|ATUAR COMO FISCAL'
    if ($hasManagerScope -or $hasInspectorScope) {
        $allowsManager = [bool]$hasManagerScope
        $allowsInspector = [bool]$hasInspectorScope
    }

    $texts = New-Object System.Collections.ArrayList
    foreach ($candidate in @(
            (Get-ManagementTextWindowFromMovement -Movement $Movement),
            (Get-ObjectValue -Item $Movement -Name 'excerpt')
        )) {
        $clean = Get-CleanText -Value $candidate
        if (-not [string]::IsNullOrWhiteSpace($clean)) {
            [void]$texts.Add($clean)
        }
    }

    $manager = $null
    $inspector = $null
    foreach ($text in @($texts)) {
        if ($allowsManager -and -not $manager) {
            $manager = Get-ManagementPersonFromText -Text $text -RoleType 'manager'
        }
        if ($allowsInspector -and -not $inspector) {
            $inspector = Get-ManagementPersonFromText -Text $text -RoleType 'inspector'
        }
        if ($manager -and $inspector) {
            break
        }
    }

    return [ordered]@{
        manager = $manager
        inspector = $inspector
    }
}

function Get-ResolvedManagementPeople {
    param(
        [AllowNull()]
        [object]$Profile,

        [AllowNull()]
        [object[]]$Movements = @()
    )

    $managementMovements = @(
        $Movements |
        Where-Object { [string]$_.recordClass -eq 'gestao_contratual' } |
        Sort-Object @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true }
    )

    $manager = $null
    $inspector = $null

    foreach ($movement in @($managementMovements)) {
        $repaired = Get-RepairedAssignmentsFromMovement -Movement $movement
        if (-not $manager -and $repaired.manager) {
            $manager = Get-PersonSnapshot `
                -Name ([string]$repaired.manager.name) `
                -Role ([string]$repaired.manager.role) `
                -AssignedAt (Get-ObjectValue -Item $movement -Name 'publishedAt')
        }
        if (-not $inspector -and $repaired.inspector) {
            $inspector = Get-PersonSnapshot `
                -Name ([string]$repaired.inspector.name) `
                -Role ([string]$repaired.inspector.role) `
                -AssignedAt (Get-ObjectValue -Item $movement -Name 'publishedAt')
        }
        if ($manager -and $inspector) {
            break
        }
    }

    if (-not $manager) {
        $manager = Get-PersonSnapshot -Name (Get-ObjectValue -Item $Profile -Name 'managerName') -Role (Get-ObjectValue -Item $Profile -Name 'managerRole') -AssignedAt (Get-ObjectValue -Item $Profile -Name 'managerAssignedAt')
    }
    if (-not $inspector) {
        $inspector = Get-PersonSnapshot -Name (Get-ObjectValue -Item $Profile -Name 'inspectorName') -Role (Get-ObjectValue -Item $Profile -Name 'inspectorRole') -AssignedAt (Get-ObjectValue -Item $Profile -Name 'inspectorAssignedAt')
    }

    return [ordered]@{
        manager = $manager
        inspector = $inspector
    }
}

function Get-ManagementStateSummary {
    param(
        [AllowNull()]
        [string]$State,

        [bool]$ManagerChanged = $false,

        [bool]$InspectorChanged = $false,

        [bool]$ManagerExonerationSignal = $false,

        [bool]$InspectorExonerationSignal = $false
    )

    switch ($State) {
        'completos' {
            if ($ManagerChanged -or $InspectorChanged) {
                return 'Gestor e fiscal atuais identificados com troca recente.'
            }
            return 'Gestor e fiscal atuais identificados.'
        }
        'sem_gestor' {
            if ($ManagerExonerationSignal) {
                return 'Sem gestor atual; ultimo designado teve exoneracao posterior.'
            }
            return 'Sem gestor atual identificado.'
        }
        'sem_fiscal' {
            if ($InspectorExonerationSignal) {
                return 'Sem fiscal atual; ultimo designado teve exoneracao posterior.'
            }
            return 'Sem fiscal atual identificado.'
        }
        'sem_gestor_e_fiscal' {
            if ($ManagerExonerationSignal -or $InspectorExonerationSignal) {
                return 'Sem gestor e fiscal atuais; ha exoneracao posterior de responsavel designado.'
            }
            return 'Sem gestor e fiscal atuais identificados.'
        }
        'revisao' { return 'Responsavel atual com revisao pendente.' }
        'exoneracao' { return 'Responsavel com sinal de exoneracao.' }
        default { return 'Situacao de gestao em acompanhamento.' }
    }
}

$script:PublicPdfToTextTool = Get-PdfToTextToolPath
$script:PublicPdfPageCache = @{}
$script:PublicManagementTextCache = @{}

function Get-PreferredMovement {
    param(
        [AllowNull()]
        [object[]]$Movements = @()
    )

    $preferred = @(
        $Movements |
        Sort-Object `
            @{ Expression = { if ([string]$_.recordClass -eq 'execucao_contratual') { 0 } else { 1 } } }, `
            @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true }
    )

    if (@($preferred).Count -gt 0) {
        return $preferred[0]
    }

    return $null
}

function Get-PreferredTextValue {
    param(
        [string[]]$Values
    )

    foreach ($value in @($Values)) {
        $clean = Get-CleanText -Value $value
        if (-not [string]::IsNullOrWhiteSpace($clean)) {
            return $clean
        }
    }

    return ''
}

function Get-PreferredNumericValue {
    param(
        [AllowNull()]
        [object[]]$Values = @()
    )

    foreach ($value in @($Values)) {
        if ($null -eq $value) {
            continue
        }
        try {
            $number = [double]$value
            if ($number -gt 0) {
                return $number
            }
        }
        catch {
        }
    }

    return 0
}

function Get-LifecycleSourceLabel {
    param(
        [AllowNull()]
        [string[]]$Sources = @()
    )

    $normalizedSources = @(
        @($Sources) |
        ForEach-Object { (Get-CleanText -Value $_).ToLowerInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if (($normalizedSources -contains 'portal') -and ($normalizedSources -contains 'diario')) {
        return 'Portal da Transparência e Diário Oficial'
    }
    if ($normalizedSources -contains 'portal') {
        return 'Portal da Transparência'
    }
    if ($normalizedSources -contains 'diario') {
        return 'Diário Oficial'
    }
    if ($normalizedSources -contains 'inferencia') {
        return 'Inferência documental'
    }

    return 'Ciclo contratual'
}

function Get-LifecycleTextBundle {
    param(
        [AllowNull()]
        [object[]]$Values = @()
    )

    $parts = @(
        @($Values) |
        ForEach-Object { Get-CleanText -Value $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    $raw = Collapse-Whitespace -Text ($parts -join ' ')
    return [ordered]@{
        raw = $raw
        normalized = Get-ContractCrossSimpleText -Text $raw
    }
}

function Get-LifecycleSignatureProfile {
    param(
        [AllowNull()]
        [object[]]$TextValues = @(),

        [AllowNull()]
        [object]$FallbackSignatureDate = $null,

        [AllowNull()]
        [object]$FallbackPublishedAt = $null
    )

    $bundle = Get-LifecycleTextBundle -Values $TextValues
    $raw = [string]$bundle.raw
    foreach ($pattern in @(
            'DATA(?:\s+DA)?\s+ASSINATURA\s*[:\-]?\s*(?<date>\d{2}/\d{2}/\d{4})',
            'ASSINADO\s+EM\s*[:\-]?\s*(?<date>\d{2}/\d{2}/\d{4})',
            'DATA\s*[:\-]?\s*(?<date>\d{2}/\d{2}/\d{4})'
        )) {
        if ($raw -match $pattern) {
            return [ordered]@{
                text = [string]$matches['date']
                iso = Convert-ToIsoString -Value ([string]$matches['date'])
            }
        }
    }

    $fallbackSignatureText = Get-CleanText -Value $FallbackSignatureDate
    if (-not [string]::IsNullOrWhiteSpace($fallbackSignatureText)) {
        return [ordered]@{
            text = $fallbackSignatureText
            iso = Convert-ToIsoString -Value $fallbackSignatureText
        }
    }

    $fallbackPublishedText = Get-CleanText -Value $FallbackPublishedAt
    return [ordered]@{
        text = $fallbackPublishedText
        iso = Convert-ToIsoString -Value $fallbackPublishedText
    }
}

function Get-LifecycleDateWindow {
    param(
        [AllowNull()]
        [object[]]$TextValues = @(),

        [AllowNull()]
        [object]$FallbackSignatureDate = $null,

        [AllowNull()]
        [object]$FallbackPublishedAt = $null,

        [AllowNull()]
        [object]$FallbackEndDate = $null
    )

    $bundle = Get-LifecycleTextBundle -Values $TextValues
    $signature = Get-LifecycleSignatureProfile -TextValues $TextValues -FallbackSignatureDate $FallbackSignatureDate -FallbackPublishedAt $FallbackPublishedAt
    $raw = [string]$bundle.raw

    foreach ($pattern in @(
            '(?<start>\d{2}/\d{2}/\d{4})\s*(?:A|ATE|ATÉ|-)\s*(?<end>\d{2}/\d{2}/\d{4})',
            'PERIODO(?:\s+DE)?\s*(?<start>\d{2}/\d{2}/\d{4})\s*(?:A|ATE|ATÉ|-)\s*(?<end>\d{2}/\d{2}/\d{4})',
            'VIGENCIA(?:\s+DO\s+AJUSTE)?[^\d]{0,20}(?<start>\d{2}/\d{2}/\d{4})\s*(?:A|ATE|ATÉ|-)\s*(?<end>\d{2}/\d{2}/\d{4})'
        )) {
        if ($raw -match $pattern) {
            return [ordered]@{
                startDate = Convert-ToIsoString -Value ([string]$matches['start'])
                endDate = Convert-ToIsoString -Value ([string]$matches['end'])
                resolution = 'document_range'
            }
        }
    }

    $fallbackEndIso = Convert-ToIsoString -Value $FallbackEndDate
    if (-not [string]::IsNullOrWhiteSpace($fallbackEndIso)) {
        return [ordered]@{
            startDate = $signature.iso
            endDate = $fallbackEndIso
            resolution = 'explicit_end'
        }
    }

    $vigencyProbe = Get-OfficialContractVigencyInfo -Item ([pscustomobject]@{
            portalStatus = ''
            signatureDate = [string]$signature.text
            object = $raw
            excerpt = $raw
            term = $raw
            actTitle = $raw
        })
    $probeEndDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $vigencyProbe -Name 'endDate')
    if (-not [string]::IsNullOrWhiteSpace($probeEndDate)) {
        return [ordered]@{
            startDate = $signature.iso
            endDate = $probeEndDate
            resolution = Get-CleanText -Value (Get-ObjectValue -Item $vigencyProbe -Name 'source')
        }
    }

    return [ordered]@{
        startDate = $signature.iso
        endDate = $null
        resolution = ''
    }
}

function Get-LifecycleEffectiveDate {
    param(
        [AllowNull()]
        [object[]]$TextValues = @(),

        [AllowNull()]
        [object]$FallbackSignatureDate = $null,

        [AllowNull()]
        [object]$FallbackPublishedAt = $null
    )

    $bundle = Get-LifecycleTextBundle -Values $TextValues
    $normalized = [string]$bundle.normalized
    foreach ($pattern in @(
            'A PARTIR DE (?<date>\d{2}/\d{2}/\d{4})',
            'RETROAG\w*.*?(?<date>\d{2}/\d{2}/\d{4})',
            'EFEITOS?.{0,40}?(?<date>\d{2}/\d{2}/\d{4})'
        )) {
        if ($normalized -match $pattern) {
            return Convert-ToIsoString -Value ([string]$matches['date'])
        }
    }

    $signature = Get-LifecycleSignatureProfile -TextValues $TextValues -FallbackSignatureDate $FallbackSignatureDate -FallbackPublishedAt $FallbackPublishedAt
    if (-not [string]::IsNullOrWhiteSpace([string]$signature.iso)) {
        return [string]$signature.iso
    }

    return Convert-ToIsoString -Value $FallbackPublishedAt
}

function Get-LifecycleEventOrdinal {
    param(
        [AllowNull()]
        [string]$Text
    )

    $cleanText = Get-CleanText -Value $Text
    if ($cleanText -match '(?<!\d)(?<value>\d{1,2})\s*[ºo°]') {
        return [int]$matches['value']
    }
    if ($cleanText -match 'N[ºo°]?\s*(?<value>\d{1,2})(?!\d)') {
        return [int]$matches['value']
    }

    return $null
}

function Get-LifecycleEventClass {
    param(
        [AllowNull()]
        [string]$Type
    )

    $normalizedType = Get-ContractCrossSimpleText -Text $Type
    switch -Regex ($normalizedType) {
        '^TERMO ADITIVO$' { return 'termo_aditivo' }
        '^APOSTILAMENTO$' { return 'apostilamento' }
        '^RESCISAO$' { return 'rescisao' }
        '^(CONTRATO|CONVENIO|CG - CONTRATOS GERAIS/OUTROS|ATO CONTRATUAL|TC - TERMO DE COLABORACAO|AC - ACORDO DE COOPERACAO|ARP - ATA DE REGISTRO DE PRECO|CL - CONTRATO DE LOCACAO|TERMO DE PARCERIA)$' { return 'contrato' }
        default { return '' }
    }
}

function Get-LifecycleEventClassification {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EventClass,

        [AllowNull()]
        [object[]]$TextValues = @()
    )

    $normalized = [string](Get-LifecycleTextBundle -Values $TextValues).normalized

    switch ($EventClass) {
        'rescisao' {
            if ($normalized -match 'ANULAC\w*\s+DA\s+RESCISAO|ANULAD\w*.*RESCISAO|RESCISAO\s+CONTRATUAL\s+ANTERIORMENTE\s+FORMALIZADA') {
                return [ordered]@{
                    kind = 'anulacao_rescisao'
                    label = 'Anulação de rescisão'
                    affectsTerm = $false
                    affectsTermination = $false
                    reversesTermination = $true
                }
            }

            return [ordered]@{
                kind = 'rescisao'
                label = 'Rescisão'
                affectsTerm = $false
                affectsTermination = $true
                reversesTermination = $false
            }
        }
        'termo_aditivo' {
            if ($normalized -match 'PRORROG|PRAZO|VIGENCIA') {
                return [ordered]@{
                    kind = 'prorrogacao_prazo'
                    label = 'Prorrogação de prazo'
                    affectsTerm = $true
                    affectsTermination = $false
                    reversesTermination = $false
                }
            }
            if ($normalized -match 'ACRESCIMO.*VALOR|REAJUSTE|REEQUILIBRIO|VALOR') {
                return [ordered]@{
                    kind = 'alteracao_valor'
                    label = 'Alteração de valor'
                    affectsTerm = $false
                    affectsTermination = $false
                    reversesTermination = $false
                }
            }
            if ($normalized -match 'SUPRESS') {
                return [ordered]@{
                    kind = 'supressao'
                    label = 'Supressão'
                    affectsTerm = $false
                    affectsTermination = $false
                    reversesTermination = $false
                }
            }
            if ($normalized -match 'INCORPORACAO|CESSAO|SUBROG|SUCESSAO|ALTERACAO\s+DE\s+EMPRESA') {
                return [ordered]@{
                    kind = 'alteracao_partes'
                    label = 'Alteração de partes'
                    affectsTerm = $false
                    affectsTermination = $false
                    reversesTermination = $false
                }
            }

            return [ordered]@{
                kind = 'termo_aditivo'
                label = 'Termo aditivo'
                affectsTerm = $false
                affectsTermination = $false
                reversesTermination = $false
            }
        }
        'apostilamento' {
            return [ordered]@{
                kind = 'apostilamento'
                label = 'Apostilamento'
                affectsTerm = [bool]($normalized -match 'PRORROG|PRAZO|VIGENCIA')
                affectsTermination = $false
                reversesTermination = $false
            }
        }
        default {
            return [ordered]@{
                kind = 'contrato_inicial'
                label = 'Contrato inicial'
                affectsTerm = $true
                affectsTermination = $false
                reversesTermination = $false
            }
        }
    }
}

function New-LifecycleEventModel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$EventClass,

        [Parameter(Mandatory = $true)]
        [hashtable]$Classification,

        [AllowNull()]
        [string]$Title = '',

        [AllowNull()]
        [object]$PublishedAt = $null,

        [AllowNull()]
        [object]$EffectiveDate = $null,

        [AllowNull()]
        [object]$StartDate = $null,

        [AllowNull()]
        [object]$EndDate = $null,

        [AllowNull()]
        [object]$Ordinal = $null,

        [AllowNull()]
        [string]$ViewUrl = '',

        [AllowNull()]
        [string]$Reference = ''
    )

    return [pscustomobject][ordered]@{
        source = $Source
        sourceLabel = Get-LifecycleSourceLabel -Sources @($Source)
        eventClass = $EventClass
        kind = Get-CleanText -Value (Get-ObjectValue -Item $Classification -Name 'kind')
        label = Get-CleanText -Value (Get-ObjectValue -Item $Classification -Name 'label')
        ordinal = $Ordinal
        title = Get-CleanText -Value $Title
        publishedAt = Convert-ToIsoString -Value $PublishedAt
        effectiveDate = Convert-ToIsoString -Value $EffectiveDate
        startDate = Convert-ToIsoString -Value $StartDate
        endDate = Convert-ToIsoString -Value $EndDate
        affectsTerm = [bool](Get-ObjectValue -Item $Classification -Name 'affectsTerm')
        affectsTermination = [bool](Get-ObjectValue -Item $Classification -Name 'affectsTermination')
        reversesTermination = [bool](Get-ObjectValue -Item $Classification -Name 'reversesTermination')
        viewUrl = Get-CleanText -Value $ViewUrl
        reference = Get-CleanText -Value $Reference
    }
}

function New-LifecycleEventFromMovement {
    param(
        [AllowNull()]
        [object]$Movement
    )

    if ($null -eq $Movement) {
        return $null
    }

    $eventClass = Get-LifecycleEventClass -Type (Get-ObjectValue -Item $Movement -Name 'type')
    if ([string]::IsNullOrWhiteSpace($eventClass)) {
        return $null
    }

    $title = Get-CleanText -Value (Get-ObjectValue -Item $Movement -Name 'actTitle')
    $textValues = @(
        $title,
        (Get-ObjectValue -Item $Movement -Name 'term'),
        (Get-ObjectValue -Item $Movement -Name 'excerpt'),
        (Get-ObjectValue -Item $Movement -Name 'object')
    )
    $classification = Get-LifecycleEventClassification -EventClass $eventClass -TextValues $textValues
    $dateWindow = Get-LifecycleDateWindow -TextValues $textValues -FallbackSignatureDate (Get-ObjectValue -Item $Movement -Name 'signatureDate') -FallbackPublishedAt (Get-ObjectValue -Item $Movement -Name 'publishedAt')
    $effectiveDate = if ($eventClass -eq 'rescisao') {
        Get-LifecycleEffectiveDate -TextValues $textValues -FallbackSignatureDate (Get-ObjectValue -Item $Movement -Name 'signatureDate') -FallbackPublishedAt (Get-ObjectValue -Item $Movement -Name 'publishedAt')
    }
    else {
        Get-PreferredTextValue -Values @((Get-ObjectValue -Item $dateWindow -Name 'startDate'), (Get-ObjectValue -Item $Movement -Name 'signatureDate'), (Get-ObjectValue -Item $Movement -Name 'publishedAt'))
    }

    return (New-LifecycleEventModel `
            -Source 'diario' `
            -EventClass $eventClass `
            -Classification $classification `
            -Title $title `
            -PublishedAt (Get-ObjectValue -Item $Movement -Name 'publishedAt') `
            -EffectiveDate $effectiveDate `
            -StartDate (Get-ObjectValue -Item $dateWindow -Name 'startDate') `
            -EndDate (Get-ObjectValue -Item $dateWindow -Name 'endDate') `
            -Ordinal $(if ($eventClass -eq 'termo_aditivo') { Get-LifecycleEventOrdinal -Text $title } else { $null }) `
            -ViewUrl (Get-ObjectValue -Item $Movement -Name 'viewUrl') `
            -Reference ('diario:' + [string](Get-ObjectValue -Item $Movement -Name 'diaryId'))
        )
}

function Get-ContractLifecycle {
    param(
        [AllowNull()]
        [object]$OfficialContract,

        [AllowNull()]
        [object[]]$Movements = @()
    )

    $events = New-Object System.Collections.ArrayList

    if ($OfficialContract) {
        $contractTextValues = @(
            (Get-ObjectValue -Item $OfficialContract -Name 'actTitle'),
            (Get-ObjectValue -Item $OfficialContract -Name 'term'),
            (Get-ObjectValue -Item $OfficialContract -Name 'object'),
            (Get-ObjectValue -Item $OfficialContract -Name 'excerpt')
        )
        $contractWindow = Get-LifecycleDateWindow `
            -TextValues $contractTextValues `
            -FallbackSignatureDate (Get-ObjectValue -Item $OfficialContract -Name 'signatureDate') `
            -FallbackPublishedAt (Get-ObjectValue -Item $OfficialContract -Name 'publishedAt') `
            -FallbackEndDate (Get-ObjectValue -Item (Get-ObjectValue -Item $OfficialContract -Name 'vigency') -Name 'endDate')

        [void]$events.Add((New-LifecycleEventModel `
                -Source 'portal' `
                -EventClass 'contrato' `
                -Classification (Get-LifecycleEventClassification -EventClass 'contrato' -TextValues $contractTextValues) `
                -Title (Get-PreferredTextValue -Values @((Get-ObjectValue -Item $OfficialContract -Name 'actTitle'), ('Contrato ' + [string](Get-ObjectValue -Item $OfficialContract -Name 'contractNumber')))) `
                -PublishedAt (Get-ObjectValue -Item $OfficialContract -Name 'publishedAt') `
                -EffectiveDate (Get-PreferredTextValue -Values @((Get-ObjectValue -Item $contractWindow -Name 'startDate'), (Get-ObjectValue -Item $OfficialContract -Name 'signatureDate'), (Get-ObjectValue -Item $OfficialContract -Name 'publishedAt'))) `
                -StartDate (Get-ObjectValue -Item $contractWindow -Name 'startDate') `
                -EndDate (Get-ObjectValue -Item $contractWindow -Name 'endDate') `
                -ViewUrl (Get-ObjectValue -Item $OfficialContract -Name 'viewUrl') `
                -Reference ('portal:' + [string](Get-ObjectValue -Item $OfficialContract -Name 'portalContractId'))
            ))

        foreach ($portalAdditive in @($(Get-ObjectValue -Item $OfficialContract -Name 'additives'))) {
            $portalTextValues = @(
                (Get-ObjectValue -Item $portalAdditive -Name 'title'),
                (Get-ObjectValue -Item $portalAdditive -Name 'term'),
                (Get-ObjectValue -Item $portalAdditive -Name 'observations'),
                (Get-ObjectValue -Item $portalAdditive -Name 'documentType')
            )
            $portalWindow = Get-LifecycleDateWindow `
                -TextValues $portalTextValues `
                -FallbackSignatureDate (Get-ObjectValue -Item $portalAdditive -Name 'signatureDate') `
                -FallbackPublishedAt (Get-ObjectValue -Item $portalAdditive -Name 'signatureDateIso') `
                -FallbackEndDate (Get-ObjectValue -Item $portalAdditive -Name 'termEndDate')

            [void]$events.Add((New-LifecycleEventModel `
                    -Source 'portal' `
                    -EventClass 'termo_aditivo' `
                    -Classification (Get-LifecycleEventClassification -EventClass 'termo_aditivo' -TextValues $portalTextValues) `
                    -Title (Get-ObjectValue -Item $portalAdditive -Name 'title') `
                    -PublishedAt (Get-ObjectValue -Item $portalAdditive -Name 'signatureDateIso') `
                    -EffectiveDate (Get-PreferredTextValue -Values @((Get-ObjectValue -Item $portalWindow -Name 'startDate'), (Get-ObjectValue -Item $portalAdditive -Name 'signatureDateIso'))) `
                    -StartDate (Get-ObjectValue -Item $portalWindow -Name 'startDate') `
                    -EndDate (Get-ObjectValue -Item $portalWindow -Name 'endDate') `
                    -Ordinal (Get-LifecycleEventOrdinal -Text (Get-ObjectValue -Item $portalAdditive -Name 'title')) `
                    -ViewUrl (Get-ObjectValue -Item $portalAdditive -Name 'webPdfPath') `
                    -Reference ('portal-additivo:' + [string](Get-ObjectValue -Item $portalAdditive -Name 'downloadTokenPath'))
                ))
        }
    }

    foreach ($movement in @($Movements)) {
        $event = New-LifecycleEventFromMovement -Movement $movement
        if ($event) {
            [void]$events.Add($event)
        }
    }

    $sortedEvents = @(
        @($events) |
        Sort-Object `
            @{ Expression = { Convert-ToDateTimeSafe -Value $(if ([string]$_.effectiveDate) { $_.effectiveDate } else { $_.publishedAt }) }; Descending = $false }, `
            @{ Expression = { switch ([string]$_.eventClass) { 'contrato' { 0 } 'termo_aditivo' { 1 } 'apostilamento' { 2 } 'rescisao' { 3 } default { 9 } } }; Descending = $false }, `
            @{ Expression = { [string]$_.title }; Descending = $false }
    )

    $currentStartDate = $null
    $currentStartTitle = ''
    $currentStartSource = ''
    $currentStartSourceLabel = ''
    $currentEndDate = $null
    $currentEndTitle = ''
    $currentEndSource = ''
    $currentEndSourceLabel = ''
    $terminationDate = $null
    $terminationTitle = ''
    $hasActiveTermination = $false
    $latestEventAt = $null
    $latestEventTitle = ''

    foreach ($event in @($sortedEvents)) {
        $eventMoment = Convert-ToDateTimeSafe -Value $(if ([string]$event.effectiveDate) { $event.effectiveDate } else { $event.publishedAt })
        if ($eventMoment -and ($null -eq $latestEventAt -or $eventMoment -gt $latestEventAt)) {
            $latestEventAt = $eventMoment
            $latestEventTitle = [string]$event.title
        }

        if ([bool]$event.reversesTermination) {
            $hasActiveTermination = $false
            $terminationDate = $null
            $terminationTitle = ''
            continue
        }

        if ([bool]$event.affectsTermination) {
            $hasActiveTermination = $true
            $terminationDate = Convert-ToDateTimeSafe -Value $(if ([string]$event.effectiveDate) { $event.effectiveDate } else { $event.publishedAt })
            $terminationTitle = [string]$event.title
            continue
        }

        if ([bool]$event.affectsTerm) {
            $eventStartDate = Convert-ToDateTimeSafe -Value $event.startDate
            if ($eventStartDate) {
                $currentStartDate = $eventStartDate
                $currentStartTitle = [string]$event.title
                $currentStartSource = [string]$event.source
                $currentStartSourceLabel = Get-LifecycleSourceLabel -Sources @([string]$event.source)
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$event.endDate)) {
                $currentEndDate = Convert-ToDateTimeSafe -Value $event.endDate
                $currentEndTitle = [string]$event.title
                $currentEndSource = [string]$event.source
                $currentEndSourceLabel = Get-LifecycleSourceLabel -Sources @([string]$event.source)
                if ($hasActiveTermination -and $terminationDate -and $eventMoment -and $eventMoment -ge $terminationDate) {
                    $hasActiveTermination = $false
                    $terminationDate = $null
                    $terminationTitle = ''
                }
            }
        }
    }

    $uniqueKeySet = New-Object 'System.Collections.Generic.HashSet[string]'
    $additiveEvents = New-Object System.Collections.ArrayList
    $apostilleEvents = New-Object System.Collections.ArrayList
    $terminationEvents = New-Object System.Collections.ArrayList
    foreach ($event in @($sortedEvents)) {
        $key = [string]::Join('|', @(
                [string]$event.eventClass,
                [string]$(if ($null -ne $event.ordinal) { $event.ordinal } else { '' }),
                [string]$event.kind,
                [string]$event.startDate,
                [string]$event.endDate,
                [string]$event.effectiveDate,
                [string]$event.title
            ))
        if (-not $uniqueKeySet.Add($key)) {
            continue
        }

        switch ([string]$event.eventClass) {
            'termo_aditivo' { [void]$additiveEvents.Add($event) }
            'apostilamento' { [void]$apostilleEvents.Add($event) }
            'rescisao' { [void]$terminationEvents.Add($event) }
        }
    }

    $summary = if ($hasActiveTermination -and $terminationDate) {
        'Encerrado por rescisão em ' + $terminationDate.ToString('dd/MM/yyyy')
    }
    elseif ($currentEndDate) {
        'Prazo consolidado até ' + $currentEndDate.ToString('dd/MM/yyyy')
    }
    else {
        'Sem prazo consolidado'
    }

    return [ordered]@{
        summary = $summary
        eventCount = [int]@($sortedEvents).Count
        latestEventAt = Convert-ToIsoString -Value $latestEventAt
        latestEventTitle = $latestEventTitle
        currentStartDate = Convert-ToIsoString -Value $currentStartDate
        currentStartTitle = $currentStartTitle
        currentStartSource = $currentStartSource
        currentStartSourceLabel = $currentStartSourceLabel
        currentEndDate = Convert-ToIsoString -Value $currentEndDate
        currentEndTitle = $currentEndTitle
        currentEndSource = $currentEndSource
        currentEndSourceLabel = $currentEndSourceLabel
        additiveCount = [int]@($additiveEvents).Count
        apostilleCount = [int]@($apostilleEvents).Count
        terminationCount = [int]@($terminationEvents).Count
        hasActiveTermination = [bool]$hasActiveTermination
        terminationDate = Convert-ToIsoString -Value $terminationDate
        terminationTitle = $terminationTitle
        isAdditivado = [bool](@($additiveEvents).Count -gt 0)
        events = @(
            @($sortedEvents) |
            ForEach-Object {
                [pscustomobject][ordered]@{
                    source = [string]$_.source
                    sourceLabel = [string]$_.sourceLabel
                    eventClass = [string]$_.eventClass
                    kind = [string]$_.kind
                    label = [string]$_.label
                    ordinal = $_.ordinal
                    title = [string]$_.title
                    publishedAt = [string]$_.publishedAt
                    effectiveDate = [string]$_.effectiveDate
                    startDate = [string]$_.startDate
                    endDate = [string]$_.endDate
                    affectsTerm = [bool]$_.affectsTerm
                    affectsTermination = [bool]$_.affectsTermination
                    reversesTermination = [bool]$_.reversesTermination
                    viewUrl = [string]$_.viewUrl
                    reference = [string]$_.reference
                }
            }
        )
    }
}

function Get-RecordVigency {
    param(
        [AllowNull()]
        [object]$OfficialContract,

        [AllowNull()]
        [object[]]$Movements = @(),

        [AllowNull()]
        [hashtable]$Lifecycle = @{}
    )

    $today = (Get-Date).Date
    $lifecycleEndDate = Convert-ToDateTimeSafe -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndDate')
    $lifecycleEndSourceLabel = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndSourceLabel')
    $terminationDate = Convert-ToDateTimeSafe -Value (Get-ObjectValue -Item $Lifecycle -Name 'terminationDate')
    $hasActiveTermination = [bool](Get-ObjectValue -Item $Lifecycle -Name 'hasActiveTermination')
    $officialVigency = $null
    if ($OfficialContract -and $OfficialContract.PSObject.Properties['vigency'] -and $OfficialContract.vigency) {
        $officialVigency = $OfficialContract.vigency
    }

    if ($hasActiveTermination) {
        return [ordered]@{
            state = 'encerrado'
            label = if ($terminationDate) { 'Encerrado por rescisão em ' + $terminationDate.ToString('dd/MM/yyyy') } else { 'Encerrado por rescisão contratual' }
            sourceLabel = Get-PreferredTextValue -Values @('Cadeia de vida contratual', (Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'terminationTitle')))
            endDate = Convert-ToIsoString -Value $(if ($terminationDate) { $terminationDate } elseif ($lifecycleEndDate) { $lifecycleEndDate } else { $null })
            daysUntilEnd = if ($terminationDate) { [int][Math]::Floor(($terminationDate.Date - $today).TotalDays) } else { $null }
            isCurrent = $false
            isConfirmed = $false
        }
    }

    if ($officialVigency) {
        if ([bool]$officialVigency.isActive) {
            $confirmedEndDate = if ($lifecycleEndDate -and $lifecycleEndDate.Date -ge $today) {
                $lifecycleEndDate
            }
            else {
                Convert-ToDateTimeSafe -Value $officialVigency.endDate
            }
            return [ordered]@{
                state = 'vigente_confirmado'
                label = if ($confirmedEndDate) { 'Vigente até ' + $confirmedEndDate.ToString('dd/MM/yyyy') } else { Get-CleanText -Value $officialVigency.summaryLabel }
                sourceLabel = Get-PreferredTextValue -Values @($lifecycleEndSourceLabel, (Get-CleanText -Value $officialVigency.sourceLabel))
                endDate = Convert-ToIsoString -Value $confirmedEndDate
                daysUntilEnd = if ($confirmedEndDate) { [int][Math]::Floor(($confirmedEndDate.Date - $today).TotalDays) } elseif ($null -ne $officialVigency.daysUntilEnd) { [int]$officialVigency.daysUntilEnd } else { $null }
                isCurrent = $true
                isConfirmed = $true
            }
        }
    }

    if ($lifecycleEndDate -and $lifecycleEndDate.Date -ge $today) {
        return [ordered]@{
            state = 'vigente_inferido'
            label = 'Prazo consolidado até ' + $lifecycleEndDate.ToString('dd/MM/yyyy')
            sourceLabel = Get-PreferredTextValue -Values @($lifecycleEndSourceLabel, 'Cadeia de vida contratual')
            endDate = Convert-ToIsoString -Value $lifecycleEndDate
            daysUntilEnd = [int][Math]::Floor(($lifecycleEndDate.Date - $today).TotalDays)
            isCurrent = $true
            isConfirmed = $false
        }
    }

    $candidate = Get-PreferredMovement -Movements $Movements
    if ($candidate) {
        $inferred = Get-OfficialContractVigencyInfo -Item ([pscustomobject]@{
            portalStatus = ''
            signatureDate = Get-ObjectValue -Item $candidate -Name 'signatureDate'
            object = Get-ObjectValue -Item $candidate -Name 'object'
            excerpt = Get-ObjectValue -Item $candidate -Name 'excerpt'
            term = Get-ObjectValue -Item $candidate -Name 'term'
            actTitle = Get-ObjectValue -Item $candidate -Name 'actTitle'
        })
        if ($inferred -and [bool]$inferred.isActive) {
            return [ordered]@{
                state = 'vigente_inferido'
                label = Get-CleanText -Value $inferred.summaryLabel
                sourceLabel = Get-CleanText -Value $inferred.sourceLabel
                endDate = Convert-ToIsoString -Value $inferred.endDate
                daysUntilEnd = if ($null -ne $inferred.daysUntilEnd) { [int]$inferred.daysUntilEnd } else { $null }
                isCurrent = $true
                isConfirmed = $false
            }
        }

        $latestMovementDate = Convert-ToDateTimeSafe -Value (Get-ObjectValue -Item $candidate -Name 'publishedAt')
        if ($latestMovementDate -and $latestMovementDate.Date -ge $today.AddDays(-400)) {
            return [ordered]@{
                state = 'em_acompanhamento'
                label = 'Em acompanhamento por movimentação recente no Diário Oficial'
                sourceLabel = 'Movimentação recente no Diário Oficial'
                endDate = $null
                daysUntilEnd = $null
                isCurrent = $true
                isConfirmed = $false
            }
        }
    }

    if ($officialVigency) {
        return [ordered]@{
            state = 'encerrado'
            label = if ($lifecycleEndDate) { 'Encerrado em ' + $lifecycleEndDate.ToString('dd/MM/yyyy') } else { Get-CleanText -Value $officialVigency.summaryLabel }
            sourceLabel = Get-PreferredTextValue -Values @($lifecycleEndSourceLabel, (Get-CleanText -Value $officialVigency.sourceLabel))
            endDate = Convert-ToIsoString -Value $(if ($lifecycleEndDate) { $lifecycleEndDate } else { $officialVigency.endDate })
            daysUntilEnd = if ($lifecycleEndDate) { [int][Math]::Floor(($lifecycleEndDate.Date - $today).TotalDays) } elseif ($null -ne $officialVigency.daysUntilEnd) { [int]$officialVigency.daysUntilEnd } else { $null }
            isCurrent = $false
            isConfirmed = [bool]$officialVigency.activeByPortal
        }
    }

    if ($lifecycleEndDate) {
        return [ordered]@{
            state = 'encerrado'
            label = 'Encerrado em ' + $lifecycleEndDate.ToString('dd/MM/yyyy')
            sourceLabel = Get-PreferredTextValue -Values @($lifecycleEndSourceLabel, 'Cadeia de vida contratual')
            endDate = Convert-ToIsoString -Value $lifecycleEndDate
            daysUntilEnd = [int][Math]::Floor(($lifecycleEndDate.Date - $today).TotalDays)
            isCurrent = $false
            isConfirmed = $false
        }
    }

    return [ordered]@{
        state = 'sem_sinal_atual'
        label = 'Sem sinal suficiente de vigência atual'
        sourceLabel = 'Sem vigência identificada'
        endDate = $null
        daysUntilEnd = $null
        isCurrent = $false
        isConfirmed = $false
    }
}

function Get-PersonnelStatusForResponsible {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Person,

        [hashtable]$PersonnelEventIndex = @{}
    )

    if ([string]::IsNullOrWhiteSpace([string]$Person.name)) {
        return [ordered]@{
            status = 'sem_evento_pessoal'
            latestEventType = ''
            latestEventAt = $null
            latestEventExcerpt = ''
        }
    }

    return (Get-PersonnelStatusAfterAssignment `
            -NormalizedName (Normalize-IndexText -Text ([string]$Person.name)) `
            -AssignedAt ([string]$Person.assignedAt) `
            -PersonnelEventIndex $PersonnelEventIndex
        )
}

function Get-ManagementState {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Manager,

        [Parameter(Mandatory = $true)]
        [hashtable]$Inspector,

        [bool]$ManagerExonerationSignal,

        [bool]$InspectorExonerationSignal
    )

    $managerPresent = -not [string]::IsNullOrWhiteSpace([string]$Manager.name) -and -not $ManagerExonerationSignal
    $inspectorPresent = -not [string]::IsNullOrWhiteSpace([string]$Inspector.name) -and -not $InspectorExonerationSignal

    if (-not $managerPresent -and -not $inspectorPresent) {
        return 'sem_gestor_e_fiscal'
    }
    if (-not $managerPresent) {
        return 'sem_gestor'
    }
    if (-not $inspectorPresent) {
        return 'sem_fiscal'
    }
    if ([bool]$Manager.needsReview -or [bool]$Inspector.needsReview) {
        return 'revisao'
    }

    return 'completos'
}

function Get-SeverityWeight {
    param(
        [AllowNull()]
        [string]$Severity
    )

    switch ((Get-CleanText -Value $Severity).ToLowerInvariant()) {
        'critical' { return 3 }
        'warning' { return 2 }
        'info' { return 1 }
        default { return 0 }
    }
}

function Add-Alert {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Alerts,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [string]$Severity = 'warning',

        [string]$Description = ''
    )

    [void]$Alerts.Add([pscustomobject][ordered]@{
        title = $Title
        severity = $Severity
        description = $Description
        weight = Get-SeverityWeight -Severity $Severity
    })
}

function Test-FieldValuePresent {
    param(
        [AllowNull()]
        [object]$Value,

        [string]$Kind = 'text'
    )

    switch ($Kind) {
        'number' {
            if ($null -eq $Value) {
                return $false
            }

            try {
                return ([double]$Value) -gt 0
            }
            catch {
                return $false
            }
        }
        'date' {
            return -not [string]::IsNullOrWhiteSpace([string](Convert-ToIsoString -Value $Value))
        }
        default {
            return -not [string]::IsNullOrWhiteSpace([string](Get-CleanText -Value $Value))
        }
    }
}

function Convert-FieldValue {
    param(
        [AllowNull()]
        [object]$Value,

        [string]$Kind = 'text'
    )

    switch ($Kind) {
        'number' {
            if (Test-FieldValuePresent -Value $Value -Kind $Kind) {
                return [double]$Value
            }
            return $null
        }
        'date' {
            return Convert-ToIsoString -Value $Value
        }
        default {
            return Get-CleanText -Value $Value
        }
    }
}

function Get-SelectedFieldCandidate {
    param(
        [AllowNull()]
        [object[]]$Candidates = @(),

        [string]$Kind = 'text'
    )

    foreach ($candidate in @($Candidates)) {
        if ($null -eq $candidate) {
            continue
        }

        $value = Get-ObjectValue -Item $candidate -Name 'value'
        if (-not (Test-FieldValuePresent -Value $value -Kind $Kind)) {
            continue
        }

        return [ordered]@{
            source = Get-CleanText -Value (Get-ObjectValue -Item $candidate -Name 'source')
            value = Convert-FieldValue -Value $value -Kind $Kind
        }
    }

    return $null
}

function Get-MovementFieldCandidates {
    param(
        [AllowNull()]
        [object[]]$Movements = @(),

        [Parameter(Mandatory = $true)]
        [string]$FieldName,

        [string]$Kind = 'text'
    )

    return @(
        @($Movements) |
        Sort-Object -Property `
            @{ Expression = {
                    switch ((Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'confidenceLabel')).ToLowerInvariant()) {
                        'alta' { 0 }
                        'media' { 1 }
                        default { 2 }
                    }
                }; Descending = $false }, `
            @{ Expression = {
                    switch ((Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'completeness')).ToLowerInvariant()) {
                        'alta' { 0 }
                        'media' { 1 }
                        default { 2 }
                    }
                }; Descending = $false }, `
            @{ Expression = {
                    $value = Get-ObjectValue -Item $_ -Name $FieldName
                    if ($Kind -eq 'text') {
                        return (Get-CleanText -Value $value).Length
                    }

                    if (Test-FieldValuePresent -Value $value -Kind $Kind) {
                        return 1
                    }

                    return 0
                }; Descending = $true }, `
            @{ Expression = { Convert-ToDateTimeSafe -Value (Get-ObjectValue -Item $_ -Name 'publishedAt') }; Descending = $true } |
        ForEach-Object {
            @{
                source = 'diario'
                value = (Get-ObjectValue -Item $_ -Name $FieldName)
            }
        }
    )
}

function Get-FieldSourceList {
    param(
        [AllowNull()]
        [object[]]$Candidates = @(),

        [string]$Kind = 'text'
    )

    $sources = New-Object System.Collections.ArrayList

    foreach ($candidate in @($Candidates)) {
        if ($null -eq $candidate) {
            continue
        }

        $value = Get-ObjectValue -Item $candidate -Name 'value'
        if (-not (Test-FieldValuePresent -Value $value -Kind $Kind)) {
            continue
        }

        $source = Get-CleanText -Value (Get-ObjectValue -Item $candidate -Name 'source')
        if (-not [string]::IsNullOrWhiteSpace($source) -and $source -notin $sources) {
            [void]$sources.Add($source)
        }
    }

    return @($sources)
}

function Get-FieldStatusProfile {
    param(
        [AllowNull()]
        [object]$SelectedCandidate,

        [string[]]$Sources = @(),

        [bool]$NeedsReview = $false,

        [bool]$IsInferred = $false
    )

    if ($null -eq $SelectedCandidate) {
        return [ordered]@{
            status = 'nao_localizado'
            confidence = 'baixa'
        }
    }

    if ($NeedsReview) {
        return [ordered]@{
            status = 'revisao'
            confidence = 'baixa'
        }
    }

    if ($IsInferred -or [string]$SelectedCandidate.source -eq 'inferencia') {
        return [ordered]@{
            status = 'inferido'
            confidence = 'media'
        }
    }

    if (($Sources -contains 'portal') -and ($Sources -contains 'diario')) {
        return [ordered]@{
            status = 'confirmado'
            confidence = 'alta'
        }
    }

    if ([string]$SelectedCandidate.source -eq 'portal') {
        return [ordered]@{
            status = 'confirmado'
            confidence = 'alta'
        }
    }

    if ([string]$SelectedCandidate.source -in @('diario', 'perfil')) {
        return [ordered]@{
            status = 'confirmado'
            confidence = 'media'
        }
    }

    return [ordered]@{
        status = 'confirmado'
        confidence = 'media'
    }
}

function New-FieldProfile {
    param(
        [AllowNull()]
        [object[]]$Candidates = @(),

        [string]$Kind = 'text',

        [bool]$NeedsReview = $false,

        [bool]$IsInferred = $false
    )

    $selected = Get-SelectedFieldCandidate -Candidates $Candidates -Kind $Kind
    $sources = @(Get-FieldSourceList -Candidates $Candidates -Kind $Kind)
    $statusProfile = Get-FieldStatusProfile -SelectedCandidate $selected -Sources $sources -NeedsReview $NeedsReview -IsInferred $IsInferred

    $defaultValue = if ($Kind -in @('number', 'date')) { $null } else { '' }

    return [ordered]@{
        value = if ($selected) { $selected.value } else { $defaultValue }
        selectedSource = if ($selected) { [string]$selected.source } else { '' }
        sources = $sources
        status = [string]$statusProfile.status
        confidence = [string]$statusProfile.confidence
    }
}

function Get-ConfidenceWeight {
    param(
        [AllowNull()]
        [string]$Level
    )

    switch ((Get-CleanText -Value $Level).ToLowerInvariant()) {
        'alta' { return 3 }
        'media' { return 2 }
        default { return 1 }
    }
}

function Get-OverallConfidence {
    param(
        [string[]]$Levels = @()
    )

    if (@($Levels).Count -eq 0) {
        return 'baixa'
    }

    $weights = @($Levels | ForEach-Object { Get-ConfidenceWeight -Level $_ })
    $average = ($weights | Measure-Object -Average).Average

    if ($average -ge 2.6) {
        return 'alta'
    }
    if ($average -ge 1.8) {
        return 'media'
    }

    return 'baixa'
}

function Get-ReviewPriorityWeight {
    param(
        [AllowNull()]
        [string]$Level
    )

    switch ((Get-CleanText -Value $Level).ToLowerInvariant()) {
        'alta' { return 3 }
        'media' { return 2 }
        default { return 1 }
    }
}

function Get-CollectionCount {
    param(
        [AllowNull()]
        [object]$Items
    )

    if ($null -eq $Items) {
        return 0
    }

    if ($Items -is [System.Array]) {
        return [int]$Items.Length
    }

    if ($Items -is [System.Collections.ICollection]) {
        return [int]$Items.Count
    }

    return [int](@($Items).Length)
}

function Add-ReviewReason {
    param(
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Reasons,

        [string]$Code,

        [string]$Title,

        [string]$Detail,

        [string]$Priority = 'media'
    )

    if ($null -eq $Reasons -or [string]::IsNullOrWhiteSpace($Code)) {
        return
    }

    foreach ($existing in @($Reasons)) {
        if ([string](Get-ObjectValue -Item $existing -Name 'code') -eq $Code) {
            return
        }
    }

    $priorityLevel = (Get-CleanText -Value $Priority).ToLowerInvariant()
    if ($priorityLevel -notin @('alta', 'media', 'baixa')) {
        $priorityLevel = 'media'
    }

    [void]$Reasons.Add([pscustomobject][ordered]@{
        code = Get-CleanText -Value $Code
        title = Get-CleanText -Value $Title
        detail = Get-CleanText -Value $Detail
        priority = $priorityLevel
        weight = Get-ReviewPriorityWeight -Level $priorityLevel
    })
}

function Get-ReviewReasonSummary {
    param(
        [AllowNull()]
        [object[]]$Reasons = @()
    )

    $reasonList = @($Reasons)
    if ((Get-CollectionCount -Items $reasonList) -eq 0) {
        return 'Sem revisao dirigida.'
    }

    $firstTitle = Get-CleanText -Value (Get-ObjectValue -Item $reasonList[0] -Name 'title')
    if ((Get-CollectionCount -Items $reasonList) -eq 1) {
        return $firstTitle
    }

    return ('{0} e mais {1} ponto(s).' -f $firstTitle, ((Get-CollectionCount -Items $reasonList) - 1))
}

function Get-ReviewSourceAlignment {
    param(
        [AllowNull()]
        [string]$SourceStatus,

        [AllowNull()]
        [object[]]$ReviewItems = @(),

        [AllowNull()]
        [object[]]$Divergences = @()
    )

    $divergenceTypes = @(
        @($Divergences) |
        ForEach-Object { (Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'type')).ToLowerInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if (($divergenceTypes -contains 'organization_mismatch') -or ($divergenceTypes -contains 'official_without_diary')) {
        return 'divergente'
    }

    if (($divergenceTypes -contains 'pending_review') -or (Get-CollectionCount -Items $ReviewItems) -gt 0) {
        return 'revisao'
    }

    if ((Get-CleanText -Value $SourceStatus).ToLowerInvariant() -eq 'cruzado') {
        return 'alinhado'
    }

    return 'parcial'
}

function New-ReviewProfile {
    param(
        [bool]$IsCurrent = $false,

        [string]$ManagementState = '',

        [string]$SourceStatus = '',

        [string]$OperationalConfidence = 'baixa',

        [string]$DocumentalConfidence = 'baixa',

        [string]$OverallConfidence = 'baixa',

        [string[]]$MissingFields = @(),

        [AllowNull()]
        [object]$ManagerModel = $null,

        [AllowNull()]
        [object]$InspectorModel = $null,

        [AllowNull()]
        [object[]]$ReviewItems = @(),

        [AllowNull()]
        [object[]]$Divergences = @(),

        [bool]$HasEndDate = $false,

        [string]$VigencyState = ''
    )

    $reviewReasons = New-Object System.Collections.ArrayList
    $criticalMissingFields = @(
        @($MissingFields) |
        Where-Object { $_ -in @('object', 'supplier', 'processNumber', 'startDate', 'endDate') } |
        Select-Object -Unique
    )
    $sortedReviewItems = @(
        @($ReviewItems) |
        Sort-Object `
            @{ Expression = { [int](Get-ObjectValue -Item $_ -Name 'recommendedScore' -Default 0) }; Descending = $true }, `
            @{ Expression = { Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'recommendedConfidence') }; Descending = $true }
    )
    $topReviewItem = $sortedReviewItems | Select-Object -First 1
    $recommendedConfidence = (Get-CleanText -Value (Get-ObjectValue -Item $topReviewItem -Name 'recommendedConfidence')).ToLowerInvariant()
    $recommendedScore = [int](Get-ObjectValue -Item $topReviewItem -Name 'recommendedScore' -Default 0)
    $divergenceTypes = @(
        @($Divergences) |
        ForEach-Object { (Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'type')).ToLowerInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if ((Get-CollectionCount -Items $sortedReviewItems) -gt 0 -or ($divergenceTypes -contains 'pending_review')) {
        Add-ReviewReason `
            -Reasons $reviewReasons `
            -Code 'cruzamento_pendente' `
            -Title 'Cruzamento pendente com o portal' `
            -Detail ('{0} candidato(s) exigem confirmacao manual.' -f (Get-CollectionCount -Items $sortedReviewItems)) `
            -Priority $(if ($IsCurrent) { 'alta' } else { 'media' })
    }

    if ($divergenceTypes -contains 'organization_mismatch') {
        Add-ReviewReason `
            -Reasons $reviewReasons `
            -Code 'divergencia_orgao' `
            -Title 'Divergencia entre orgaos' `
            -Detail 'Diario Oficial e portal apresentam orgaos diferentes para a mesma referencia.' `
            -Priority 'alta'
    }

    if ($divergenceTypes -contains 'official_without_diary') {
        Add-ReviewReason `
            -Reasons $reviewReasons `
            -Code 'portal_sem_diario' `
            -Title 'Contrato oficial sem ato correspondente' `
            -Detail 'O portal apresenta contrato sem ato consolidado correspondente no Diario Oficial.' `
            -Priority $(if ($IsCurrent) { 'alta' } else { 'media' })
    }

    if ([bool](Get-ObjectValue -Item $ManagerModel -Name 'needsReview') -or [bool](Get-ObjectValue -Item $InspectorModel -Name 'needsReview')) {
        Add-ReviewReason `
            -Reasons $reviewReasons `
            -Code 'responsavel_impreciso' `
            -Title 'Responsavel com leitura imprecisa' `
            -Detail 'A leitura automatica do nome ou cargo do responsavel precisa confirmacao manual.' `
            -Priority $(if ($IsCurrent) { 'alta' } else { 'media' })
    }

    if ($IsCurrent -and $DocumentalConfidence -eq 'baixa' -and (Get-CollectionCount -Items $criticalMissingFields) -ge 3 -and $SourceStatus -ne 'cruzado' -and $ManagementState -eq 'completos') {
        Add-ReviewReason `
            -Reasons $reviewReasons `
            -Code 'documentacao_incompleta' `
            -Title 'Documentacao incompleta' `
            -Detail 'Objeto, fornecedor, processo ou prazo ainda nao estao completos com confianca suficiente.' `
            -Priority 'media'
    }

    $priorityWeight = 0
    foreach ($reason in @($reviewReasons)) {
        $weight = [int](Get-ObjectValue -Item $reason -Name 'weight' -Default 0)
        if ($weight -gt $priorityWeight) {
            $priorityWeight = $weight
        }
    }

    $priority = switch ($priorityWeight) {
        3 { 'alta' }
        2 { 'media' }
        default { 'baixa' }
    }

    $candidates = @(
        $sortedReviewItems |
        Select-Object -First 3 |
        ForEach-Object {
            [pscustomobject][ordered]@{
                portalContractId = Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'portalContractId')
                contractNumber = Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'contractNumber')
                organization = Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'organization')
                score = [int](Get-ObjectValue -Item $_ -Name 'recommendedScore' -Default (Get-ObjectValue -Item $_ -Name 'score' -Default 0))
                confidence = (Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'recommendedConfidence')).ToLowerInvariant()
                reason = Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'reason')
            }
        }
    )

    $divergenceSnapshot = @(
        @($Divergences) |
        Select-Object -First 3 |
        ForEach-Object {
            [pscustomobject][ordered]@{
                type = (Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'type')).ToLowerInvariant()
                title = Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'title')
                reason = Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'reason')
                severity = (Get-CleanText -Value (Get-ObjectValue -Item $_ -Name 'severity')).ToLowerInvariant()
            }
        }
    )

    return [ordered]@{
        required = [bool]((Get-CollectionCount -Items $reviewReasons) -gt 0)
        priority = $priority
        priorityWeight = $priorityWeight
        sourceAlignment = Get-ReviewSourceAlignment -SourceStatus $SourceStatus -ReviewItems $ReviewItems -Divergences $Divergences
        reasonSummary = Get-ReviewReasonSummary -Reasons $reviewReasons
        reasonCount = Get-CollectionCount -Items $reviewReasons
        reasons = @($reviewReasons | Sort-Object weight -Descending)
        candidateCount = Get-CollectionCount -Items $sortedReviewItems
        recommendedConfidence = $recommendedConfidence
        recommendedScore = $recommendedScore
        candidates = $candidates
        divergenceTypes = $divergenceTypes
        divergenceCount = Get-CollectionCount -Items $Divergences
        divergences = $divergenceSnapshot
        criticalMissingFields = $criticalMissingFields
        operationalConfidence = $OperationalConfidence
        documentalConfidence = $DocumentalConfidence
        overallConfidence = $OverallConfidence
    }
}

function Get-AdditiveSummary {
    param(
        [AllowNull()]
        [object]$OfficialContract,

        [AllowNull()]
        [object[]]$Movements = @(),

        [AllowNull()]
        [hashtable]$Lifecycle = @{}
    )

    $lifecycleEvents = @($(Get-ObjectValue -Item $Lifecycle -Name 'events'))
    $additiveEvents = @($lifecycleEvents | Where-Object { [string]$_.eventClass -eq 'termo_aditivo' })
    $apostilles = @($lifecycleEvents | Where-Object { [string]$_.eventClass -eq 'apostilamento' })
    $terminations = @($lifecycleEvents | Where-Object { [string]$_.eventClass -eq 'rescisao' })
    $latestAdditive = $additiveEvents | Sort-Object @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true } | Select-Object -First 1
    $portalCount = [int]@($additiveEvents | Where-Object { [string]$_.source -eq 'portal' }).Count
    $diaryCount = [int]@($additiveEvents | Where-Object { [string]$_.source -eq 'diario' }).Count
    $knownCount = [int]@($additiveEvents).Count

    return [ordered]@{
        isAdditivado = [bool]($knownCount -gt 0)
        totalKnown = [int]$knownCount
        portalCount = [int]$portalCount
        diaryCount = $diaryCount
        apostilleCount = [int](@($apostilles).Count)
        terminationCount = [int](@($terminations).Count)
        latestAdditiveAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $latestAdditive -Name 'publishedAt')
        currentEndDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndDate')
        hasActiveTermination = [bool](Get-ObjectValue -Item $Lifecycle -Name 'hasActiveTermination')
        terminationDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $Lifecycle -Name 'terminationDate')
    }
}

function Get-LifecycleSnapshot {
    param(
        [AllowNull()]
        [hashtable]$Lifecycle = @{},

        [switch]$IncludeEvents
    )

    $snapshot = [ordered]@{
        summary = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'summary')
        eventCount = [int](Get-ObjectValue -Item $Lifecycle -Name 'eventCount' -Default 0)
        latestEventAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $Lifecycle -Name 'latestEventAt')
        latestEventTitle = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'latestEventTitle')
        currentStartDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentStartDate')
        currentStartTitle = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentStartTitle')
        currentStartSource = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentStartSource')
        currentStartSourceLabel = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentStartSourceLabel')
        currentEndDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndDate')
        currentEndTitle = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndTitle')
        currentEndSource = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndSource')
        currentEndSourceLabel = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndSourceLabel')
        additiveCount = [int](Get-ObjectValue -Item $Lifecycle -Name 'additiveCount' -Default 0)
        apostilleCount = [int](Get-ObjectValue -Item $Lifecycle -Name 'apostilleCount' -Default 0)
        terminationCount = [int](Get-ObjectValue -Item $Lifecycle -Name 'terminationCount' -Default 0)
        isAdditivado = [bool](Get-ObjectValue -Item $Lifecycle -Name 'isAdditivado')
        hasActiveTermination = [bool](Get-ObjectValue -Item $Lifecycle -Name 'hasActiveTermination')
        terminationDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $Lifecycle -Name 'terminationDate')
        terminationTitle = Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'terminationTitle')
    }

    if ($IncludeEvents) {
        $snapshot['events'] = @($(Get-ObjectValue -Item $Lifecycle -Name 'events'))
    }

    return $snapshot
}

function Get-SourceEvidence {
    param(
        [AllowNull()]
        [object]$OfficialContract,

        [AllowNull()]
        [object[]]$Movements = @(),

        [AllowNull()]
        [object]$LatestMovement,

        [string]$SourceStatus = ''
    )

    $movementList = @($Movements)
    $movementTypeCount = @{}
    $diaryIds = New-Object System.Collections.ArrayList

    foreach ($movement in $movementList) {
        $type = Get-CleanText -Value (Get-ObjectValue -Item $movement -Name 'type')
        if (-not [string]::IsNullOrWhiteSpace($type)) {
            if (-not $movementTypeCount.ContainsKey($type)) {
                $movementTypeCount[$type] = 0
            }
            $movementTypeCount[$type] = [int]$movementTypeCount[$type] + 1
        }

        $diaryId = Get-CleanText -Value (Get-ObjectValue -Item $movement -Name 'diaryId')
        if (-not [string]::IsNullOrWhiteSpace($diaryId) -and $diaryId -notin $diaryIds) {
            [void]$diaryIds.Add($diaryId)
        }
    }

    $movementTypes = @(
        $movementTypeCount.GetEnumerator() |
        Sort-Object @{ Expression = { $_.Value }; Descending = $true }, @{ Expression = { $_.Name }; Descending = $false } |
        ForEach-Object {
            [pscustomobject][ordered]@{
                type = [string]$_.Name
                count = [int]$_.Value
            }
        }
    )

    $movementCount = [int](($movementList | Measure-Object).Count)

    return [ordered]@{
        sourceStatus = $SourceStatus
        diary = [ordered]@{
            hasDiary = [bool]($movementCount -gt 0)
            movementCount = $movementCount
            diaryIds = @($diaryIds)
            latestDiaryId = Get-CleanText -Value (Get-ObjectValue -Item $LatestMovement -Name 'diaryId')
            latestEdition = Get-CleanText -Value (Get-ObjectValue -Item $LatestMovement -Name 'edition')
            latestPublishedAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $LatestMovement -Name 'publishedAt')
            latestActTitle = Get-CleanText -Value (Get-ObjectValue -Item $LatestMovement -Name 'actTitle')
            movementTypes = $movementTypes
            latestViewUrl = Get-CleanText -Value (Get-ObjectValue -Item $LatestMovement -Name 'viewUrl')
        }
        portal = [ordered]@{
            hasPortal = [bool]$OfficialContract
            portalContractId = Get-CleanText -Value (Get-ObjectValue -Item $OfficialContract -Name 'portalContractId')
            updatedAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $OfficialContract -Name 'updatedAt')
            publishedAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $OfficialContract -Name 'publishedAt')
            portalStatus = Get-CleanText -Value (Get-ObjectValue -Item $OfficialContract -Name 'portalStatus')
            aditiveCount = [int]$(if ($OfficialContract) { Get-ObjectValue -Item $OfficialContract -Name 'aditiveCount' -Default 0 } else { 0 })
            viewUrl = Get-CleanText -Value (Get-ObjectValue -Item $OfficialContract -Name 'viewUrl')
        }
    }
}

function New-ResponsibilityRoleModel {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Person,

        [bool]$ExonerationSignal = $false
    )

    $needsReview = [bool]$Person.needsReview
    $nameField = New-FieldProfile -Candidates @(@{ source = 'diario'; value = (Get-ObjectValue -Item $Person -Name 'name') }) -NeedsReview $needsReview
    $roleField = New-FieldProfile -Candidates @(@{ source = 'diario'; value = (Get-ObjectValue -Item $Person -Name 'role') }) -NeedsReview $needsReview
    $assignedAtField = New-FieldProfile -Candidates @(@{ source = 'diario'; value = (Get-ObjectValue -Item $Person -Name 'assignedAt') }) -Kind 'date' -NeedsReview $needsReview

    $status = if ($ExonerationSignal) {
        'exoneracao_sinalizada'
    }
    elseif ($needsReview) {
        'revisao'
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$nameField.value)) {
        'confirmado'
    }
    else {
        'nao_localizado'
    }

    return [ordered]@{
        name = $nameField.value
        role = $roleField.value
        assignedAt = $assignedAtField.value
        status = $status
        confidence = Get-OverallConfidence -Levels @($nameField.confidence, $roleField.confidence, $assignedAtField.confidence)
        needsReview = $needsReview
        exonerationSignal = [bool]$ExonerationSignal
        fields = [ordered]@{
            name = $nameField
            role = $roleField
            assignedAt = $assignedAtField
        }
    }
}

function New-MasterContractModel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,

        [Parameter(Mandatory = $true)]
        [string]$NormalizedKey,

        [AllowNull()]
        [object]$Profile,

        [AllowNull()]
        [object]$OfficialContract,

        [AllowNull()]
        [object[]]$Movements = @(),

        [AllowNull()]
        [object]$PreferredMovement,

        [AllowNull()]
        [object]$LatestMovement,

        [AllowNull()]
        [hashtable]$Vigency = @{},

        [AllowNull()]
        [hashtable]$Lifecycle = @{},

        [Parameter(Mandatory = $true)]
        [hashtable]$Manager,

        [Parameter(Mandatory = $true)]
        [hashtable]$Inspector,

        [string]$ManagementState = '',

        [string]$ManagementSummary = '',

        [bool]$ManagerExonerationSignal = $false,

        [bool]$InspectorExonerationSignal = $false,

        [string]$Administration = '',

        [AllowNull()]
        [object]$Year = $null,

        [string]$Organization = '',

        [string]$SourceStatus = '',

        [AllowNull()]
        [object[]]$ReviewItems = @(),

        [AllowNull()]
        [object[]]$Divergences = @()
    )

    $lifecycleStartSource = switch ((Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentStartSource')).ToLowerInvariant()) {
        'portal' { 'portal' }
        'diario' { 'diario' }
        default { 'inferencia' }
    }
    $lifecycleEndSource = switch ((Get-CleanText -Value (Get-ObjectValue -Item $Lifecycle -Name 'currentEndSource')).ToLowerInvariant()) {
        'portal' { 'portal' }
        'diario' { 'diario' }
        default { 'inferencia' }
    }

    $contractNumberField = New-FieldProfile -Candidates @(
        @{ source = 'perfil'; value = (Get-ObjectValue -Item $Profile -Name 'contractNumber') },
        @{ source = 'diario'; value = (Get-ObjectValue -Item $PreferredMovement -Name 'contractNumber') },
        @{ source = 'diario'; value = (Get-ObjectValue -Item $LatestMovement -Name 'contractNumber') },
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'contractNumber') },
        @{ source = 'inferencia'; value = $NormalizedKey }
    )

    $processNumberField = New-FieldProfile -Candidates @(
        @{ source = 'perfil'; value = (Get-ObjectValue -Item $Profile -Name 'processNumber') }
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'processNumber'
        )
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'processNumber') }
    )

    $organizationField = New-FieldProfile -Candidates @(
        @{ source = 'diario'; value = (Get-ObjectValue -Item $PreferredMovement -Name 'primaryOrganizationName') },
        @{ source = 'diario'; value = (Get-ObjectValue -Item $LatestMovement -Name 'primaryOrganizationName') },
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'primaryOrganizationName') }
    )

    $supplierField = New-FieldProfile -Candidates @(
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'contractor'
        )
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'contractor') }
    )

    $supplierDocumentField = New-FieldProfile -Candidates @(
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'cnpj'
        )
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'cnpj') }
    )

    $objectField = New-FieldProfile -Candidates @(
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'object'
        )
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'object') }
    )

    $valueLabelField = New-FieldProfile -Candidates @(
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'value'
        )
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'value') }
    )

    $valueAmountField = New-FieldProfile -Candidates @(
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'valueNumber' -Kind 'number'
        )
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'valueNumber') }
    ) -Kind 'number'

    $startDateField = New-FieldProfile -Candidates @(
        @{ source = $lifecycleStartSource; value = (Get-ObjectValue -Item $Lifecycle -Name 'currentStartDate') },
        @{ source = 'portal'; value = (Get-ObjectValue -Item $OfficialContract -Name 'signatureDate') },
        @{ source = 'portal'; value = (Get-ObjectValue -Item (Get-ObjectValue -Item $OfficialContract -Name 'vigency') -Name 'signatureDate') },
        @(
            Get-MovementFieldCandidates -Movements $Movements -FieldName 'signatureDate' -Kind 'date'
        )
    ) -Kind 'date'

    $endDateField = New-FieldProfile -Candidates @(
        @{ source = $lifecycleEndSource; value = (Get-ObjectValue -Item $Lifecycle -Name 'currentEndDate') },
        @{ source = 'portal'; value = (Get-ObjectValue -Item (Get-ObjectValue -Item $OfficialContract -Name 'vigency') -Name 'endDate') },
        @{ source = 'inferencia'; value = (Get-ObjectValue -Item $Vigency -Name 'endDate') }
    ) -Kind 'date'

    $additives = Get-AdditiveSummary -OfficialContract $OfficialContract -Movements $Movements -Lifecycle $Lifecycle
    $lifecycleSnapshot = Get-LifecycleSnapshot -Lifecycle $Lifecycle -IncludeEvents
    $sources = Get-SourceEvidence -OfficialContract $OfficialContract -Movements $Movements -LatestMovement $LatestMovement -SourceStatus $SourceStatus
    $managerModel = New-ResponsibilityRoleModel -Person $Manager -ExonerationSignal $ManagerExonerationSignal
    $inspectorModel = New-ResponsibilityRoleModel -Person $Inspector -ExonerationSignal $InspectorExonerationSignal
    $valueConfidence = Get-OverallConfidence -Levels @($valueLabelField.confidence, $valueAmountField.confidence)
    $termConfidence = Get-OverallConfidence -Levels @($startDateField.confidence, $endDateField.confidence)
    $additivesConfidence = if ([bool]$additives.isAdditivado) {
        if ([int]$additives.portalCount -gt 0 -and [int]$additives.diaryCount -gt 0) { 'alta' } else { 'media' }
    }
    elseif ($OfficialContract) {
        'media'
    }
    else {
        'baixa'
    }

    $operationalConfidence = Get-OverallConfidence -Levels @(
        $contractNumberField.confidence,
        $termConfidence,
        $managerModel.confidence,
        $inspectorModel.confidence
    )

    $documentalConfidence = Get-OverallConfidence -Levels @(
        $processNumberField.confidence,
        $objectField.confidence,
        $supplierField.confidence,
        $valueConfidence,
        $additivesConfidence
    )

    $overallConfidence = Get-OverallConfidence -Levels @(
        $contractNumberField.confidence,
        $objectField.confidence,
        $supplierField.confidence,
        $termConfidence,
        $managerModel.confidence,
        $inspectorModel.confidence
    )

    $missingFields = New-Object System.Collections.ArrayList
    foreach ($item in @(
        @{ name = 'processNumber'; present = -not [string]::IsNullOrWhiteSpace([string]$processNumberField.value) },
        @{ name = 'object'; present = -not [string]::IsNullOrWhiteSpace([string]$objectField.value) },
        @{ name = 'supplier'; present = -not [string]::IsNullOrWhiteSpace([string]$supplierField.value) },
        @{ name = 'manager'; present = -not [string]::IsNullOrWhiteSpace([string]$managerModel.name) },
        @{ name = 'inspector'; present = -not [string]::IsNullOrWhiteSpace([string]$inspectorModel.name) },
        @{ name = 'startDate'; present = $null -ne $startDateField.value },
        @{ name = 'endDate'; present = $null -ne $endDateField.value }
    )) {
        if (-not [bool]$item.present) {
            [void]$missingFields.Add([string]$item.name)
        }
    }

    $review = New-ReviewProfile `
        -IsCurrent ([bool](Get-ObjectValue -Item $Vigency -Name 'isCurrent')) `
        -ManagementState $ManagementState `
        -SourceStatus $SourceStatus `
        -OperationalConfidence $operationalConfidence `
        -DocumentalConfidence $documentalConfidence `
        -OverallConfidence $overallConfidence `
        -MissingFields @($missingFields) `
        -ManagerModel $managerModel `
        -InspectorModel $inspectorModel `
        -ReviewItems $ReviewItems `
        -Divergences $Divergences `
        -HasEndDate ($null -ne $endDateField.value) `
        -VigencyState (Get-CleanText -Value (Get-ObjectValue -Item $Vigency -Name 'state'))

    return [pscustomobject][ordered]@{
        id = $Id
        normalizedKey = $NormalizedKey
        administration = Get-CleanText -Value $Administration
        year = if ($null -ne $Year -and "$Year" -ne '') { [int]$Year } else { $null }
        contractNumber = $contractNumberField.value
        processNumber = $processNumberField.value
        organization = if (-not [string]::IsNullOrWhiteSpace([string]$organizationField.value)) { $organizationField.value } else { Get-CleanText -Value $Organization }
        supplier = [ordered]@{
            name = $supplierField.value
            document = $supplierDocumentField.value
            confidence = Get-OverallConfidence -Levels @($supplierField.confidence, $supplierDocumentField.confidence)
            fields = [ordered]@{
                name = $supplierField
                document = $supplierDocumentField
            }
        }
        object = $objectField.value
        value = [ordered]@{
            amount = $valueAmountField.value
            label = $valueLabelField.value
            confidence = $valueConfidence
            fields = [ordered]@{
                amount = $valueAmountField
                label = $valueLabelField
            }
        }
        term = [ordered]@{
            startDate = $startDateField.value
            endDate = $endDateField.value
            daysUntilEnd = Get-ObjectValue -Item $Vigency -Name 'daysUntilEnd'
            state = Get-CleanText -Value (Get-ObjectValue -Item $Vigency -Name 'state')
            label = Get-CleanText -Value (Get-ObjectValue -Item $Vigency -Name 'label')
            sourceLabel = Get-CleanText -Value (Get-ObjectValue -Item $Vigency -Name 'sourceLabel')
            isCurrent = [bool](Get-ObjectValue -Item $Vigency -Name 'isCurrent')
            isConfirmed = [bool](Get-ObjectValue -Item $Vigency -Name 'isConfirmed')
            confidence = $termConfidence
            fields = [ordered]@{
                startDate = $startDateField
                endDate = $endDateField
            }
        }
        lifecycle = $lifecycleSnapshot
        additives = $additives
        responsibilities = [ordered]@{
            state = $ManagementState
            summary = $ManagementSummary
            manager = $managerModel
            inspector = $inspectorModel
        }
        sources = $sources
        confidence = [ordered]@{
            overall = $overallConfidence
            operational = $operationalConfidence
            documental = $documentalConfidence
            contractNumber = $contractNumberField.confidence
            processNumber = $processNumberField.confidence
            organization = $organizationField.confidence
            supplier = $supplierField.confidence
            object = $objectField.confidence
            value = $valueConfidence
            term = $termConfidence
            manager = $managerModel.confidence
            inspector = $inspectorModel.confidence
            additives = $additivesConfidence
        }
        review = $review
        missingFields = @($missingFields)
    }
}

$source = Get-Content -LiteralPath $SourcePath -Raw | ConvertFrom-Json
$diariesById = Get-DiariesById
$personnelEventIndex = Get-PersonnelEventIndex -DiariesById $diariesById

$movementGroups = @{}
foreach ($movement in @($source.contractMovements)) {
    $groupKeys = @()
    foreach ($candidateKey in @($movement.managementProfileKey, $movement.referenceKey, $movement.contractNumber)) {
        $normalized = Get-NormalizedContractKey -Value $candidateKey
        if (-not [string]::IsNullOrWhiteSpace($normalized) -and $normalized -notin $groupKeys) {
            $groupKeys += $normalized
        }
    }

    foreach ($normalized in @($groupKeys)) {
        if (-not $movementGroups.ContainsKey($normalized)) {
            $movementGroups[$normalized] = @()
        }
        $movementGroups[$normalized] = @($movementGroups[$normalized]) + @($movement)
    }
}

$officialByKey = @{}
foreach ($official in @($source.officialContracts)) {
    $normalized = Get-NormalizedContractKey -Value $(if ($official.referenceKey) { $official.referenceKey } else { $official.contractNumber })
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        continue
    }

    if (-not $officialByKey.ContainsKey($normalized)) {
        $officialByKey[$normalized] = @()
    }
    $officialByKey[$normalized] = @($officialByKey[$normalized]) + @($official)
}

$alertByKey = @{}
foreach ($item in @($source.crossSourceAlerts)) {
    $normalized = Get-NormalizedContractKey -Value $(if ($item.crossKey) { $item.crossKey } else { $item.contractNumber })
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        continue
    }

    if (-not $alertByKey.ContainsKey($normalized)) {
        $alertByKey[$normalized] = @()
    }
    $alertByKey[$normalized] = @($alertByKey[$normalized]) + @($item)
}

$divergenceByKey = @{}
foreach ($item in @($source.crossSourceDivergences)) {
    $normalized = Get-NormalizedContractKey -Value $(if ($item.crossKey) { $item.crossKey } else { $item.contractNumber })
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        continue
    }

    if (-not $divergenceByKey.ContainsKey($normalized)) {
        $divergenceByKey[$normalized] = @()
    }
    $divergenceByKey[$normalized] = @($divergenceByKey[$normalized]) + @($item)
}

$reviewByKey = @{}
foreach ($item in @($source.crossReviewQueue)) {
    $normalized = Get-NormalizedContractKey -Value $(if ($item.crossKey) { $item.crossKey } else { $item.movementReference })
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        continue
    }

    if (-not $reviewByKey.ContainsKey($normalized)) {
        $reviewByKey[$normalized] = @()
    }
    $reviewByKey[$normalized] = @($reviewByKey[$normalized]) + @($item)
}

$usedPortalIds = @{}
$records = New-Object System.Collections.ArrayList
$masterContracts = New-Object System.Collections.ArrayList

foreach ($profile in @($source.managementProfiles)) {
    $normalizedKey = Get-NormalizedContractKey -Value $profile.contractKey
    if ([string]::IsNullOrWhiteSpace($normalizedKey)) {
        continue
    }

    $movements = if ($movementGroups.ContainsKey($normalizedKey)) { @($movementGroups[$normalizedKey]) } else { @() }
    $officialCandidates = if ($officialByKey.ContainsKey($normalizedKey)) { @($officialByKey[$normalizedKey]) } else { @() }
    $officialContract = if (@($officialCandidates).Count -eq 1) { $officialCandidates[0] } else { $null }
    if ($officialContract -and (Get-ObjectValue -Item $officialContract -Name 'portalContractId')) {
        $usedPortalIds[[string](Get-ObjectValue -Item $officialContract -Name 'portalContractId')] = $true
    }

    $latestMovement = @($movements | Sort-Object @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true }) | Select-Object -First 1
    $preferredMovement = Get-PreferredMovement -Movements $movements
    $lifecycle = Get-ContractLifecycle -OfficialContract $officialContract -Movements $movements
    $lifecycleSnapshot = Get-LifecycleSnapshot -Lifecycle $lifecycle
    $resolvedPeople = Get-ResolvedManagementPeople -Profile $profile -Movements $movements
    $manager = $resolvedPeople.manager
    $inspector = $resolvedPeople.inspector
    $vigency = Get-RecordVigency -OfficialContract $officialContract -Movements $movements -Lifecycle $lifecycle
    $managerPersonnelStatus = Get-PersonnelStatusForResponsible -Person $manager -PersonnelEventIndex $personnelEventIndex
    $inspectorPersonnelStatus = Get-PersonnelStatusForResponsible -Person $inspector -PersonnelEventIndex $personnelEventIndex
    $managerExonerationSignal = [bool]$profile.managerExonerationSignal -or ([string](Get-ObjectValue -Item $managerPersonnelStatus -Name 'status') -eq 'exonerado')
    $inspectorExonerationSignal = [bool]$profile.inspectorExonerationSignal -or ([string](Get-ObjectValue -Item $inspectorPersonnelStatus -Name 'status') -eq 'exonerado')
    $managementState = Get-ManagementState -Manager $manager -Inspector $inspector -ManagerExonerationSignal $managerExonerationSignal -InspectorExonerationSignal $inspectorExonerationSignal
    $managementSummary = Get-ManagementStateSummary `
        -State $managementState `
        -ManagerChanged ([bool](Get-ObjectValue -Item $profile -Name 'managerChanged')) `
        -InspectorChanged ([bool](Get-ObjectValue -Item $profile -Name 'inspectorChanged')) `
        -ManagerExonerationSignal $managerExonerationSignal `
        -InspectorExonerationSignal $inspectorExonerationSignal
    $year = Get-ContractYear -NormalizedKey $normalizedKey -Movements $movements -OfficialContract $officialContract
    $organization = Get-PreferredTextValue -Values @(
        (Get-ObjectValue -Item $preferredMovement -Name 'primaryOrganizationName'),
        (Get-ObjectValue -Item $latestMovement -Name 'primaryOrganizationName'),
        (Get-ObjectValue -Item $officialContract -Name 'primaryOrganizationName')
    )
    $supplier = Get-PreferredTextValue -Values @(
        (Get-ObjectValue -Item $preferredMovement -Name 'contractor'),
        (Get-ObjectValue -Item $latestMovement -Name 'contractor'),
        (Get-ObjectValue -Item $officialContract -Name 'contractor')
    )
    $description = Get-PreferredTextValue -Values @(
        (Get-ObjectValue -Item $preferredMovement -Name 'object'),
        (Get-ObjectValue -Item $latestMovement -Name 'object'),
        (Get-ObjectValue -Item $officialContract -Name 'object')
    )
    $processNumber = Get-PreferredTextValue -Values @(
        (Get-ObjectValue -Item $profile -Name 'processNumber'),
        (Get-ObjectValue -Item $preferredMovement -Name 'processNumber'),
        (Get-ObjectValue -Item $latestMovement -Name 'processNumber'),
        (Get-ObjectValue -Item $officialContract -Name 'processNumber')
    )
    $valueNumber = Get-PreferredNumericValue -Values @(
        (Get-ObjectValue -Item $preferredMovement -Name 'valueNumber'),
        (Get-ObjectValue -Item $latestMovement -Name 'valueNumber'),
        (Get-ObjectValue -Item $officialContract -Name 'valueNumber')
    )
    $valueLabel = Get-PreferredTextValue -Values @(
        (Get-ObjectValue -Item $preferredMovement -Name 'value'),
        (Get-ObjectValue -Item $latestMovement -Name 'value'),
        (Get-ObjectValue -Item $officialContract -Name 'value')
    )
    $reviewItems = if ($reviewByKey.ContainsKey($normalizedKey)) { @($reviewByKey[$normalizedKey]) } else { @() }
    $divergences = if ($divergenceByKey.ContainsKey($normalizedKey)) { @($divergenceByKey[$normalizedKey]) } else { @() }
    $alerts = New-Object System.Collections.ArrayList

    if ($vigency.isCurrent -and $managementState -eq 'sem_gestor_e_fiscal') {
        Add-Alert -Alerts $alerts -Title 'Sem gestor e fiscal atuais' -Severity 'critical' -Description 'Não foi possível confirmar gestor nem fiscal atuais para este contrato.'
    }
    elseif ($vigency.isCurrent -and $managementState -eq 'sem_gestor') {
        Add-Alert -Alerts $alerts -Title 'Sem gestor atual' -Severity 'critical' -Description 'O contrato está em acompanhamento, mas sem gestor atual confirmado.'
    }
    elseif ($vigency.isCurrent -and $managementState -eq 'sem_fiscal') {
        Add-Alert -Alerts $alerts -Title 'Sem fiscal atual' -Severity 'critical' -Description 'O contrato está em acompanhamento, mas sem fiscal atual confirmado.'
    }

    if ($vigency.isCurrent -and ($managerExonerationSignal -or $inspectorExonerationSignal)) {
        Add-Alert -Alerts $alerts -Title 'Sinal de exoneração de responsável' -Severity 'critical' -Description 'Há indício de exoneração envolvendo gestor ou fiscal vinculados a este contrato.'
    }

    if ($vigency.isCurrent -and -not $officialContract) {
        Add-Alert -Alerts $alerts -Title 'Sem cadastro correspondente no portal' -Severity 'warning' -Description 'Há movimentação no Diário Oficial, mas não houve correspondência automática com contrato oficial.'
    }

    if ($vigency.isCurrent -and $vigency.state -eq 'em_acompanhamento') {
        Add-Alert -Alerts $alerts -Title 'Vigência em acompanhamento' -Severity 'warning' -Description 'A situação atual decorre de movimentação recente no Diário e ainda não de vigência confirmada.'
    }

    if ($vigency.isCurrent -and $null -ne $vigency.daysUntilEnd -and [int]$vigency.daysUntilEnd -le 30) {
        Add-Alert -Alerts $alerts -Title 'Prazo final próximo' -Severity $(if ([int]$vigency.daysUntilEnd -le 15) { 'critical' } else { 'warning' }) -Description ('Prazo final estimado em {0} dia(s).' -f [int]$vigency.daysUntilEnd)
    }

    if ([bool]$manager.needsReview -or [bool]$inspector.needsReview) {
        Add-Alert -Alerts $alerts -Title 'Responsável com extração imprecisa' -Severity 'warning' -Description 'O nome extraído do Diário Oficial precisa revisão manual antes de ser tratado como confirmado.'
    }

    foreach ($item in @($(if ($alertByKey.ContainsKey($normalizedKey)) { $alertByKey[$normalizedKey] } else { @() }) | Select-Object -First 3)) {
        Add-Alert -Alerts $alerts -Title (Get-CleanText -Value $item.title) -Severity (Get-CleanText -Value $item.severity) -Description (Get-CleanText -Value $item.reason)
    }

    foreach ($item in @($(if ($divergenceByKey.ContainsKey($normalizedKey)) { $divergenceByKey[$normalizedKey] } else { @() }) | Select-Object -First 2)) {
        Add-Alert -Alerts $alerts -Title (Get-CleanText -Value $item.title) -Severity (Get-CleanText -Value $item.severity) -Description (Get-CleanText -Value $item.reason)
    }

    $highestAlertWeight = 0
    foreach ($alert in @($alerts)) {
        if ([int]$alert.weight -gt $highestAlertWeight) {
            $highestAlertWeight = [int]$alert.weight
        }
    }

    $recordId = 'profile:' + $normalizedKey
    $masterId = 'master:' + $normalizedKey
    $additives = Get-AdditiveSummary -OfficialContract $officialContract -Movements $movements -Lifecycle $lifecycle

    $masterContract = New-MasterContractModel `
        -Id $masterId `
        -NormalizedKey $normalizedKey `
        -Profile $profile `
        -OfficialContract $officialContract `
        -Movements $movements `
        -PreferredMovement $preferredMovement `
        -LatestMovement $latestMovement `
        -Vigency $vigency `
        -Lifecycle $lifecycle `
        -Manager $manager `
        -Inspector $inspector `
        -ManagementState $managementState `
        -ManagementSummary $managementSummary `
        -ManagerExonerationSignal $managerExonerationSignal `
        -InspectorExonerationSignal $inspectorExonerationSignal `
        -Administration (Get-AdministrationLabel -Year $year) `
        -Year $year `
        -Organization $organization `
        -SourceStatus $(if ($officialContract) { 'cruzado' } else { 'somente_diario' }) `
        -ReviewItems $reviewItems `
        -Divergences $divergences
    [void]$masterContracts.Add($masterContract)

    [void]$records.Add([pscustomobject][ordered]@{
        id = $recordId
        masterContractId = $masterId
        recordType = 'diario_monitorado'
        normalizedKey = $normalizedKey
        contractNumber = Get-PreferredTextValue -Values @($profile.contractNumber, $preferredMovement.contractNumber, $latestMovement.contractNumber, $normalizedKey)
        processNumber = $processNumber
        administration = Get-AdministrationLabel -Year $year
        year = $year
        organization = $organization
        supplier = $supplier
        object = $description
        valueLabel = $valueLabel
        valueNumber = $valueNumber
        vigency = $vigency
        lifecycle = $lifecycleSnapshot
        additives = $additives
        managementState = $managementState
        managementSummary = $managementSummary
        manager = $manager
        inspector = $inspector
        managerPersonnelStatus = Get-CleanText -Value (Get-ObjectValue -Item $managerPersonnelStatus -Name 'status')
        inspectorPersonnelStatus = Get-CleanText -Value (Get-ObjectValue -Item $inspectorPersonnelStatus -Name 'status')
        managerExonerationSignal = $managerExonerationSignal
        inspectorExonerationSignal = $inspectorExonerationSignal
        hasDiary = $true
        hasOfficialPortal = [bool]$officialContract
        sourceStatus = if ($officialContract) { 'cruzado' } else { 'somente_diario' }
        confidence = $masterContract.confidence
        review = $masterContract.review
        publishedAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $latestMovement -Name 'publishedAt')
        managementActAt = Convert-ToIsoString -Value $profile.lastManagementActAt
        lastMovementTitle = Get-CleanText -Value (Get-ObjectValue -Item $latestMovement -Name 'actTitle')
        movementCount = @($movements).Count
        alertWeight = $highestAlertWeight
        alertCount = @($alerts).Count
        alerts = @($alerts | Sort-Object weight -Descending)
        links = [ordered]@{
            diary = Get-CleanText -Value (Get-ObjectValue -Item $latestMovement -Name 'viewUrl')
            portal = Get-CleanText -Value (Get-ObjectValue -Item $officialContract -Name 'viewUrl')
        }
    })
}

foreach ($official in @($source.officialContracts | Sort-Object @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true })) {
    $portalId = Get-CleanText -Value $official.portalContractId
    if (-not [string]::IsNullOrWhiteSpace($portalId) -and $usedPortalIds.ContainsKey($portalId)) {
        continue
    }

    $normalizedKey = Get-NormalizedContractKey -Value $(if ($official.referenceKey) { $official.referenceKey } else { $official.contractNumber })
    $movements = if ($movementGroups.ContainsKey($normalizedKey)) { @($movementGroups[$normalizedKey]) } else { @() }
    $lifecycle = Get-ContractLifecycle -OfficialContract $official -Movements $movements
    $lifecycleSnapshot = Get-LifecycleSnapshot -Lifecycle $lifecycle
    $vigency = Get-RecordVigency -OfficialContract $official -Movements $movements -Lifecycle $lifecycle
    $year = Get-ContractYear -NormalizedKey $normalizedKey -Movements $movements -OfficialContract $official
    $latestMovement = $movements | Sort-Object @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true } | Select-Object -First 1
    $preferredMovement = Get-PreferredMovement -Movements $movements
    $reviewItems = if ($reviewByKey.ContainsKey($normalizedKey)) { @($reviewByKey[$normalizedKey]) } else { @() }
    $divergences = if ($divergenceByKey.ContainsKey($normalizedKey)) { @($divergenceByKey[$normalizedKey]) } else { @() }
    $alerts = New-Object System.Collections.ArrayList

    if ($vigency.isCurrent -and -not @($movements).Count) {
        Add-Alert -Alerts $alerts -Title 'Sem ato correspondente no Diário' -Severity 'critical' -Description 'O contrato aparece como vigente, mas sem movimentação correspondente consolidada no Diário Oficial.'
    }

    foreach ($item in @($(if ($alertByKey.ContainsKey($normalizedKey)) { $alertByKey[$normalizedKey] } else { @() }) | Select-Object -First 3)) {
        Add-Alert -Alerts $alerts -Title (Get-CleanText -Value $item.title) -Severity (Get-CleanText -Value $item.severity) -Description (Get-CleanText -Value $item.reason)
    }

    foreach ($item in @($(if ($divergenceByKey.ContainsKey($normalizedKey)) { $divergenceByKey[$normalizedKey] } else { @() }) | Select-Object -First 2)) {
        Add-Alert -Alerts $alerts -Title (Get-CleanText -Value $item.title) -Severity (Get-CleanText -Value $item.severity) -Description (Get-CleanText -Value $item.reason)
    }

    $highestAlertWeight = 0
    foreach ($alert in @($alerts)) {
        if ([int]$alert.weight -gt $highestAlertWeight) {
            $highestAlertWeight = [int]$alert.weight
        }
    }

    $recordId = 'official:' + $(if ($portalId) { $portalId } else { $normalizedKey })
    $masterId = 'master:' + $(if ($portalId) { $portalId } else { $normalizedKey })
    $emptyManager = [ordered]@{ name = ''; role = ''; assignedAt = $null; needsReview = $false }
    $emptyInspector = [ordered]@{ name = ''; role = ''; assignedAt = $null; needsReview = $false }
    $additives = Get-AdditiveSummary -OfficialContract $official -Movements $movements -Lifecycle $lifecycle

    $masterContract = New-MasterContractModel `
        -Id $masterId `
        -NormalizedKey $normalizedKey `
        -Profile $null `
        -OfficialContract $official `
        -Movements $movements `
        -PreferredMovement $preferredMovement `
        -LatestMovement $latestMovement `
        -Vigency $vigency `
        -Lifecycle $lifecycle `
        -Manager $emptyManager `
        -Inspector $emptyInspector `
        -ManagementState 'sem_gestor_e_fiscal' `
        -ManagementSummary (Get-CleanText -Value $official.managementSummary) `
        -Administration (Get-AdministrationLabel -Year $year) `
        -Year $year `
        -Organization (Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'primaryOrganizationName')) `
        -SourceStatus $(if (@($movements).Count) { 'cruzado' } else { 'somente_portal' }) `
        -ReviewItems $reviewItems `
        -Divergences $divergences
    [void]$masterContracts.Add($masterContract)

    [void]$records.Add([pscustomobject][ordered]@{
        id = $recordId
        masterContractId = $masterId
        recordType = 'portal_oficial'
        normalizedKey = $normalizedKey
        contractNumber = Get-PreferredTextValue -Values @($official.contractNumber, $normalizedKey)
        processNumber = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'processNumber')
        administration = Get-AdministrationLabel -Year $year
        year = $year
        organization = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'primaryOrganizationName')
        supplier = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'contractor')
        object = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'object')
        valueLabel = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'value')
        valueNumber = Get-PreferredNumericValue -Values @((Get-ObjectValue -Item $official -Name 'valueNumber'))
        vigency = $vigency
        lifecycle = $lifecycleSnapshot
        additives = $additives
        managementState = 'sem_gestor_e_fiscal'
        managementSummary = Get-CleanText -Value $official.managementSummary
        manager = $emptyManager
        inspector = $emptyInspector
        managerPersonnelStatus = 'sem_evento_pessoal'
        inspectorPersonnelStatus = 'sem_evento_pessoal'
        managerExonerationSignal = $false
        inspectorExonerationSignal = $false
        hasDiary = [bool](@($movements).Count)
        hasOfficialPortal = $true
        sourceStatus = if (@($movements).Count) { 'cruzado' } else { 'somente_portal' }
        confidence = $masterContract.confidence
        review = $masterContract.review
        publishedAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $official -Name 'publishedAt')
        managementActAt = $null
        lastMovementTitle = ''
        movementCount = @($movements).Count
        alertWeight = $highestAlertWeight
        alertCount = @($alerts).Count
        alerts = @($alerts | Sort-Object weight -Descending)
        links = [ordered]@{
            diary = Get-CleanText -Value (Get-ObjectValue -Item $latestMovement -Name 'viewUrl')
            portal = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'viewUrl')
        }
    })
}

$sortedRecords = @(
    $records |
    Sort-Object `
        @{ Expression = { if ([bool]$_.vigency.isCurrent) { 0 } else { 1 } } }, `
        @{ Expression = { -1 * [int]$_.alertWeight } }, `
        @{ Expression = { Convert-ToDateTimeSafe -Value $_.managementActAt }; Descending = $true }, `
        @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true }
)

$masterById = @{}
foreach ($master in @($masterContracts)) {
    $masterById[[string]$master.id] = $master
}

$sortedMasterContracts = @(
    $sortedRecords |
    ForEach-Object {
        $masterId = [string]$_.masterContractId
        if (-not [string]::IsNullOrWhiteSpace($masterId) -and $masterById.ContainsKey($masterId)) {
            $masterById[$masterId]
        }
    }
)

$currentRecords = @($sortedRecords | Where-Object { [bool]$_.vigency.isCurrent })
$reviewQueue = @(
    $sortedRecords |
    Where-Object { [bool](Get-ObjectValue -Item (Get-ObjectValue -Item $_ -Name 'review') -Name 'required') } |
    Sort-Object `
        @{ Expression = { -1 * [int](Get-ObjectValue -Item (Get-ObjectValue -Item $_ -Name 'review') -Name 'priorityWeight' -Default 0) } }, `
        @{ Expression = { if ([bool]$_.vigency.isCurrent) { 0 } else { 1 } } }, `
        @{ Expression = { if ($null -ne $_.vigency.daysUntilEnd) { [int]$_.vigency.daysUntilEnd } else { 999999 } } }, `
        @{ Expression = { Convert-ToDateTimeSafe -Value $_.managementActAt }; Descending = $true }, `
        @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true } |
    ForEach-Object {
        $review = Get-ObjectValue -Item $_ -Name 'review'
        $confidence = Get-ObjectValue -Item $_ -Name 'confidence'

        [pscustomobject][ordered]@{
            id = 'review:' + [string]$_.masterContractId
            masterContractId = [string]$_.masterContractId
            normalizedKey = Get-CleanText -Value $_.normalizedKey
            contractNumber = Get-CleanText -Value $_.contractNumber
            administration = Get-CleanText -Value $_.administration
            organization = Get-CleanText -Value $_.organization
            sourceStatus = Get-CleanText -Value $_.sourceStatus
            managementState = Get-CleanText -Value $_.managementState
            isCurrent = [bool]$_.vigency.isCurrent
            priority = Get-CleanText -Value (Get-ObjectValue -Item $review -Name 'priority')
            priorityWeight = [int](Get-ObjectValue -Item $review -Name 'priorityWeight' -Default 0)
            sourceAlignment = Get-CleanText -Value (Get-ObjectValue -Item $review -Name 'sourceAlignment')
            reasonSummary = Get-CleanText -Value (Get-ObjectValue -Item $review -Name 'reasonSummary')
            reasonCount = [int](Get-ObjectValue -Item $review -Name 'reasonCount' -Default 0)
            reasons = @($(Get-ObjectValue -Item $review -Name 'reasons'))
            candidateCount = [int](Get-ObjectValue -Item $review -Name 'candidateCount' -Default 0)
            recommendedConfidence = Get-CleanText -Value (Get-ObjectValue -Item $review -Name 'recommendedConfidence')
            recommendedScore = [int](Get-ObjectValue -Item $review -Name 'recommendedScore' -Default 0)
            divergenceCount = [int](Get-ObjectValue -Item $review -Name 'divergenceCount' -Default 0)
            divergenceTypes = @($(Get-ObjectValue -Item $review -Name 'divergenceTypes'))
            criticalMissingFields = @($(Get-ObjectValue -Item $review -Name 'criticalMissingFields'))
            overallConfidence = Get-CleanText -Value (Get-ObjectValue -Item $confidence -Name 'overall')
            operationalConfidence = Get-CleanText -Value (Get-ObjectValue -Item $confidence -Name 'operational')
            documentalConfidence = Get-CleanText -Value (Get-ObjectValue -Item $confidence -Name 'documental')
            publishedAt = Convert-ToIsoString -Value $_.publishedAt
            managementActAt = Convert-ToIsoString -Value $_.managementActAt
            endDate = Convert-ToIsoString -Value (Get-ObjectValue -Item $_.vigency -Name 'endDate')
            daysUntilEnd = Get-ObjectValue -Item $_.vigency -Name 'daysUntilEnd'
            links = [ordered]@{
                diary = Get-CleanText -Value (Get-ObjectValue -Item (Get-ObjectValue -Item $_ -Name 'links') -Name 'diary')
                portal = Get-CleanText -Value (Get-ObjectValue -Item (Get-ObjectValue -Item $_ -Name 'links') -Name 'portal')
            }
        }
    }
)
$organizationSummary = @(
    $currentRecords |
    Group-Object organization |
    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Name) } |
    Sort-Object Count -Descending |
    Select-Object -First 8 |
    ForEach-Object {
        [pscustomobject][ordered]@{
            organization = [string]$_.Name
            count = [int]$_.Count
        }
    }
)

$masterSummary = [ordered]@{
    totalContracts = [int](@($sortedMasterContracts).Count)
    withContractNumber = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.contractNumber) }).Count)
    withProcessNumber = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.processNumber) }).Count)
    withObject = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.object) }).Count)
    withSupplier = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.supplier.name) }).Count)
    withValue = [int](@($sortedMasterContracts | Where-Object { $null -ne $_.value.amount -or -not [string]::IsNullOrWhiteSpace([string]$_.value.label) }).Count)
    withStartDate = [int](@($sortedMasterContracts | Where-Object { $null -ne $_.term.startDate }).Count)
    withEndDate = [int](@($sortedMasterContracts | Where-Object { $null -ne $_.term.endDate }).Count)
    withManager = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.responsibilities.manager.name) }).Count)
    withInspector = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.responsibilities.inspector.name) }).Count)
    withCompleteResponsibilities = [int](@($sortedMasterContracts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.responsibilities.manager.name) -and -not [string]::IsNullOrWhiteSpace([string]$_.responsibilities.inspector.name) }).Count)
    aditivados = [int](@($sortedMasterContracts | Where-Object { [bool]$_.additives.isAdditivado }).Count)
    apostilados = [int](@($sortedMasterContracts | Where-Object { [int]$_.lifecycle.apostilleCount -gt 0 }).Count)
    rescindidos = [int](@($sortedMasterContracts | Where-Object { [bool]$_.lifecycle.hasActiveTermination }).Count)
    highOverallConfidence = [int](@($sortedMasterContracts | Where-Object { [string]$_.confidence.overall -eq 'alta' }).Count)
    mediumOverallConfidence = [int](@($sortedMasterContracts | Where-Object { [string]$_.confidence.overall -eq 'media' }).Count)
    lowOverallConfidence = [int](@($sortedMasterContracts | Where-Object { [string]$_.confidence.overall -eq 'baixa' }).Count)
    highOperationalConfidence = [int](@($sortedMasterContracts | Where-Object { [string]$_.confidence.operational -eq 'alta' }).Count)
    mediumOperationalConfidence = [int](@($sortedMasterContracts | Where-Object { [string]$_.confidence.operational -eq 'media' }).Count)
    lowOperationalConfidence = [int](@($sortedMasterContracts | Where-Object { [string]$_.confidence.operational -eq 'baixa' }).Count)
    reviewRequired = [int](@($sortedMasterContracts | Where-Object { [bool]$_.review.required }).Count)
    highReviewPriority = [int](@($sortedMasterContracts | Where-Object { [string]$_.review.priority -eq 'alta' }).Count)
}

$reviewSummary = [ordered]@{
    total = [int]@($reviewQueue).Count
    current = [int]@($reviewQueue | Where-Object { [bool]$_.isCurrent }).Count
    high = [int]@($reviewQueue | Where-Object { [string]$_.priority -eq 'alta' }).Count
    medium = [int]@($reviewQueue | Where-Object { [string]$_.priority -eq 'media' }).Count
    low = [int]@($reviewQueue | Where-Object { [string]$_.priority -eq 'baixa' }).Count
    divergent = [int]@($reviewQueue | Where-Object { [string]$_.sourceAlignment -eq 'divergente' }).Count
    crossPending = [int]@($reviewQueue | Where-Object { [int]$_.candidateCount -gt 0 }).Count
    operationalLow = [int]@($reviewQueue | Where-Object { [string]$_.operationalConfidence -eq 'baixa' }).Count
    documentalLow = [int]@($reviewQueue | Where-Object { [string]$_.documentalConfidence -eq 'baixa' }).Count
}

$payload = [ordered]@{
    generatedAt = Convert-ToIsoString -Value $source.generatedAt
    masterSchemaVersion = '2026-04-01.4'
    methodology = [ordered]@{
        title = 'Leitura cruzada de contratos vigentes e responsáveis'
        summary = 'O painel cruza Diário Oficial, contratos do portal e eventos de gestão contratual para destacar vigência, gestor, fiscal e alertas de vacância.'
        notes = @(
            'Quando o portal não confirma a vigência, o painel mostra contratos em acompanhamento com base em movimentação recente no Diário Oficial.',
            'Nomes extraídos como portaria, decreto ou texto corrido não são tratados como responsáveis confirmados e ficam marcados para revisão.',
            'A ausência de vínculo automático com o portal é exibida como alerta, sem esconder o contrato identificado no Diário Oficial.'
        )
    }
    summary = [ordered]@{
        totalMonitorados = [int]@($sortedRecords).Count
        contratosAtuais = [int]@($currentRecords).Count
        vigentesConfirmados = [int]@($currentRecords | Where-Object { [string]$_.vigency.state -eq 'vigente_confirmado' }).Count
        vigentesInferidos = [int]@($currentRecords | Where-Object { [string]$_.vigency.state -eq 'vigente_inferido' }).Count
        emAcompanhamento = [int]@($currentRecords | Where-Object { [string]$_.vigency.state -eq 'em_acompanhamento' }).Count
        semGestor = [int]@($currentRecords | Where-Object { [string]$_.managementState -in @('sem_gestor', 'sem_gestor_e_fiscal') }).Count
        semFiscal = [int]@($currentRecords | Where-Object { [string]$_.managementState -in @('sem_fiscal', 'sem_gestor_e_fiscal') }).Count
        semGestorEFiscal = [int]@($currentRecords | Where-Object { [string]$_.managementState -eq 'sem_gestor_e_fiscal' }).Count
        comResponsaveisCompletos = [int]@($currentRecords | Where-Object { [string]$_.managementState -eq 'completos' }).Count
        sinaisExoneracao = [int]@($currentRecords | Where-Object { [bool]$_.managerExonerationSignal -or [bool]$_.inspectorExonerationSignal }).Count
        alertasCriticos = [int]@($currentRecords | Where-Object { [int]$_.alertWeight -ge 3 }).Count
        aditivadosAtuais = [int]@($currentRecords | Where-Object { [bool]$_.additives.isAdditivado }).Count
        comPrazoConsolidado = [int]@($currentRecords | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.lifecycle.currentEndDate) }).Count
        somenteDiario = [int]@($currentRecords | Where-Object { [string]$_.sourceStatus -eq 'somente_diario' }).Count
        somentePortal = [int]@($currentRecords | Where-Object { [string]$_.sourceStatus -eq 'somente_portal' }).Count
        cruzados = [int]@($currentRecords | Where-Object { [string]$_.sourceStatus -eq 'cruzado' }).Count
        revisaoDirigida = [int]@($reviewQueue | Where-Object { [bool]$_.isCurrent }).Count
        analisados = [int]$source.analyzedDiaryCount
        contratosPortal = [int]$source.officialPortalContracts
    }
    filters = [ordered]@{
        organizations = @($sortedRecords | Group-Object organization | Where-Object { $_.Name } | Sort-Object Name | ForEach-Object { [string]$_.Name })
        administrations = @($sortedRecords | Group-Object administration | Where-Object { $_.Name } | Sort-Object Name | ForEach-Object { [string]$_.Name })
        vigencyStates = @('todos', 'vigente_confirmado', 'vigente_inferido', 'em_acompanhamento', 'encerrado', 'sem_sinal_atual')
        managementStates = @('todos', 'completos', 'sem_gestor', 'sem_fiscal', 'sem_gestor_e_fiscal', 'revisao', 'exoneracao')
        sourceStates = @('todos', 'cruzado', 'somente_diario', 'somente_portal')
        scopeStates = @('atuais', 'todos')
    }
    organizationSummary = $organizationSummary
    masterSummary = $masterSummary
    reviewSummary = $reviewSummary
    reviewQueue = $reviewQueue
    masterContracts = $sortedMasterContracts
    records = $sortedRecords
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$payload | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Output ('Arquivo público gerado em ' + $OutputPath)
