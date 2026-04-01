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
    $OutputPath = Join-Path $scriptRoot '..\docs\data\contracts-dashboard.json'
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

    if ($Item -is [hashtable] -and $Item.ContainsKey($Name)) {
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

        [bool]$InspectorChanged = $false
    )

    switch ($State) {
        'completos' {
            if ($ManagerChanged -or $InspectorChanged) {
                return 'Gestor e fiscal atuais identificados com troca recente.'
            }
            return 'Gestor e fiscal atuais identificados.'
        }
        'sem_gestor' { return 'Sem gestor atual identificado.' }
        'sem_fiscal' { return 'Sem fiscal atual identificado.' }
        'sem_gestor_e_fiscal' { return 'Sem gestor e fiscal atuais identificados.' }
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

function Get-RecordVigency {
    param(
        [AllowNull()]
        [object]$OfficialContract,

        [AllowNull()]
        [object[]]$Movements = @()
    )

    $today = (Get-Date).Date

    if ($OfficialContract -and $OfficialContract.PSObject.Properties['vigency'] -and $OfficialContract.vigency) {
        $officialVigency = $OfficialContract.vigency
        if ([bool]$officialVigency.isActive) {
            return [ordered]@{
                state = if ([bool]$officialVigency.activeByPortal) { 'vigente_confirmado' } else { 'vigente_inferido' }
                label = Get-CleanText -Value $officialVigency.summaryLabel
                sourceLabel = Get-CleanText -Value $officialVigency.sourceLabel
                endDate = Convert-ToIsoString -Value $officialVigency.endDate
                daysUntilEnd = if ($null -ne $officialVigency.daysUntilEnd) { [int]$officialVigency.daysUntilEnd } else { $null }
                isCurrent = $true
                isConfirmed = [bool]$officialVigency.activeByPortal
            }
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

    if ($OfficialContract -and $OfficialContract.PSObject.Properties['vigency'] -and $OfficialContract.vigency) {
        return [ordered]@{
            state = 'encerrado'
            label = Get-CleanText -Value $OfficialContract.vigency.summaryLabel
            sourceLabel = Get-CleanText -Value $OfficialContract.vigency.sourceLabel
            endDate = Convert-ToIsoString -Value $OfficialContract.vigency.endDate
            daysUntilEnd = if ($null -ne $OfficialContract.vigency.daysUntilEnd) { [int]$OfficialContract.vigency.daysUntilEnd } else { $null }
            isCurrent = $false
            isConfirmed = [bool]$OfficialContract.vigency.activeByPortal
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

function Get-ManagementState {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Manager,

        [Parameter(Mandatory = $true)]
        [hashtable]$Inspector,

        [bool]$ManagerExonerationSignal,

        [bool]$InspectorExonerationSignal
    )

    if ($ManagerExonerationSignal -or $InspectorExonerationSignal) {
        return 'exoneracao'
    }
    if ([string]::IsNullOrWhiteSpace([string]$Manager.name) -and [string]::IsNullOrWhiteSpace([string]$Inspector.name)) {
        return 'sem_gestor_e_fiscal'
    }
    if ([string]::IsNullOrWhiteSpace([string]$Manager.name)) {
        return 'sem_gestor'
    }
    if ([string]::IsNullOrWhiteSpace([string]$Inspector.name)) {
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

$source = Get-Content -LiteralPath $SourcePath -Raw | ConvertFrom-Json

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

$usedPortalIds = @{}
$records = New-Object System.Collections.ArrayList

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
    $resolvedPeople = Get-ResolvedManagementPeople -Profile $profile -Movements $movements
    $manager = $resolvedPeople.manager
    $inspector = $resolvedPeople.inspector
    $vigency = Get-RecordVigency -OfficialContract $officialContract -Movements $movements
    $managerExonerationSignal = [bool]$profile.managerExonerationSignal
    $inspectorExonerationSignal = [bool]$profile.inspectorExonerationSignal
    $managementState = Get-ManagementState -Manager $manager -Inspector $inspector -ManagerExonerationSignal $managerExonerationSignal -InspectorExonerationSignal $inspectorExonerationSignal
    $managementSummary = Get-ManagementStateSummary `
        -State $managementState `
        -ManagerChanged ([bool](Get-ObjectValue -Item $profile -Name 'managerChanged')) `
        -InspectorChanged ([bool](Get-ObjectValue -Item $profile -Name 'inspectorChanged'))
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

    [void]$records.Add([pscustomobject][ordered]@{
        id = 'profile:' + $normalizedKey
        recordType = 'diario_monitorado'
        normalizedKey = $normalizedKey
        contractNumber = Get-PreferredTextValue -Values @($profile.contractNumber, $preferredMovement.contractNumber, $latestMovement.contractNumber, $normalizedKey)
        administration = Get-AdministrationLabel -Year $year
        year = $year
        organization = $organization
        supplier = $supplier
        object = $description
        valueLabel = $valueLabel
        valueNumber = $valueNumber
        vigency = $vigency
        managementState = $managementState
        managementSummary = $managementSummary
        manager = $manager
        inspector = $inspector
        hasDiary = $true
        hasOfficialPortal = [bool]$officialContract
        sourceStatus = if ($officialContract) { 'cruzado' } else { 'somente_diario' }
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
    $vigency = Get-RecordVigency -OfficialContract $official -Movements $movements
    $year = Get-ContractYear -NormalizedKey $normalizedKey -Movements $movements -OfficialContract $official
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

    [void]$records.Add([pscustomobject][ordered]@{
        id = 'official:' + $(if ($portalId) { $portalId } else { $normalizedKey })
        recordType = 'portal_oficial'
        normalizedKey = $normalizedKey
        contractNumber = Get-PreferredTextValue -Values @($official.contractNumber, $normalizedKey)
        administration = Get-AdministrationLabel -Year $year
        year = $year
        organization = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'primaryOrganizationName')
        supplier = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'contractor')
        object = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'object')
        valueLabel = Get-CleanText -Value (Get-ObjectValue -Item $official -Name 'value')
        valueNumber = Get-PreferredNumericValue -Values @((Get-ObjectValue -Item $official -Name 'valueNumber'))
        vigency = $vigency
        managementState = 'sem_gestor_e_fiscal'
        managementSummary = Get-CleanText -Value $official.managementSummary
        manager = [ordered]@{ name = ''; role = ''; assignedAt = $null; needsReview = $false }
        inspector = [ordered]@{ name = ''; role = ''; assignedAt = $null; needsReview = $false }
        hasDiary = [bool]@($movements).Count
        hasOfficialPortal = $true
        sourceStatus = if (@($movements).Count) { 'cruzado' } else { 'somente_portal' }
        publishedAt = Convert-ToIsoString -Value (Get-ObjectValue -Item $official -Name 'publishedAt')
        managementActAt = $null
        lastMovementTitle = ''
        movementCount = @($movements).Count
        alertWeight = $highestAlertWeight
        alertCount = @($alerts).Count
        alerts = @($alerts | Sort-Object weight -Descending)
        links = [ordered]@{
            diary = Get-CleanText -Value $(if (@($movements).Count) { (Get-ObjectValue -Item (@($movements | Sort-Object @{ Expression = { Convert-ToDateTimeSafe -Value $_.publishedAt }; Descending = $true } | Select-Object -First 1)) -Name 'viewUrl') } else { '' })
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

$currentRecords = @($sortedRecords | Where-Object { [bool]$_.vigency.isCurrent })
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

$payload = [ordered]@{
    generatedAt = Convert-ToIsoString -Value $source.generatedAt
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
        sinaisExoneracao = [int]@($currentRecords | Where-Object { [string]$_.managementState -eq 'exoneracao' }).Count
        alertasCriticos = [int]@($currentRecords | Where-Object { [int]$_.alertWeight -ge 3 }).Count
        somenteDiario = [int]@($currentRecords | Where-Object { [string]$_.sourceStatus -eq 'somente_diario' }).Count
        somentePortal = [int]@($currentRecords | Where-Object { [string]$_.sourceStatus -eq 'somente_portal' }).Count
        cruzados = [int]@($currentRecords | Where-Object { [string]$_.sourceStatus -eq 'cruzado' }).Count
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
    records = $sortedRecords
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Output ('Arquivo público gerado em ' + $OutputPath)
