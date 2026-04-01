Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Web

$script:AppRoot = Split-Path -Parent $PSScriptRoot
$script:BasePortalUri = [Uri]'https://www.iguape.sp.gov.br'
$script:PortalDiarioPath = '/portal/diario-oficial'
$script:LegacyTransparencyPortalUri = [Uri]'http://pmiguape.ddns.net:81/PortaldaTransparencia/'
$script:LegacyTransparencyContractPath = '/PortaldaTransparencia/Pages/Geral/wfContrato.aspx'
$script:LegacyTransparencyExpensePath = '/PortaldaTransparencia/Pages/Geral/wfDespesa.aspx'
$script:LegacyTransparencyExpenseDetailPath = '/PortaldaTransparencia/Pages/Geral/wfDespesaExibicao.aspx'
$script:DadosIguapeUri = [Uri]'https://smfiguape.vercel.app/'
$script:UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Codex-Iguape-Contracts/1.0'
$script:ParserVersion = '2026.03.18.10'

$script:StorageRoot = Join-Path $script:AppRoot 'storage'
$script:DataRoot = Join-Path $script:AppRoot 'data'
$script:PdfRoot = Join-Path $script:StorageRoot 'pdfs'
$script:AnalysisRoot = Join-Path $script:StorageRoot 'analysis'
$script:PersonnelAnalysisRoot = Join-Path $script:StorageRoot 'personnel-analysis'
$script:StateRoot = Join-Path $script:StorageRoot 'state'

$script:DiariesPath = Join-Path $script:StorageRoot 'diaries.json'
$script:ContractsPath = Join-Path $script:StorageRoot 'contracts.json'
$script:PortalContractsPath = Join-Path $script:StorageRoot 'portal-contracts.json'
$script:StatusPath = Join-Path $script:StateRoot 'status.json'
$script:SyncLockPath = Join-Path $script:StateRoot 'sync.lock.json'
$script:UsersPath = Join-Path $script:StateRoot 'users.json'
$script:SupportPath = Join-Path $script:StateRoot 'support.json'
$script:ContractCrossReviewPath = Join-Path $script:StateRoot 'contract-cross-review.json'
$script:WorkspaceStatePath = Join-Path $script:StateRoot 'workspace.json'
$script:ObservabilityPath = Join-Path $script:StateRoot 'observability.json'
$script:SearchIndexPath = Join-Path $script:StateRoot 'search-index.json'
$script:OrganizationCatalogPath = Join-Path $script:DataRoot 'organization-catalog.json'
$script:UsersSchemaVersion = '2026.03.26.02'
$script:SupportSchemaVersion = '2026.03.26.02'
$script:ContractCrossReviewSchemaVersion = '2026.03.26.02'
$script:WorkspaceSchemaVersion = '2026.03.30.02'
$script:ObservabilitySchemaVersion = '2026.03.26.03'
$script:SearchIndexSchemaVersion = '2026.03.30.01'
$script:DashboardPayloadContractVersion = '2026.03.30.02'
$script:WorkspaceSessionPayloadContractVersion = '2026.03.27.01'
$script:ContractAuditPayloadContractVersion = '2026.03.27.01'
$script:ContractDetailPayloadContractVersion = '2026.03.27.02'
$script:ContractCollectionPayloadContractVersion = '2026.03.30.01'
$script:StatusPayloadContractVersion = '2026.03.27.01'
$script:PublicStatusPayloadContractVersion = '2026.03.27.02'
$script:SearchPayloadContractVersion = '2026.03.27.01'
$script:PersonnelParserVersion = '2026.03.23.01'
$script:JsonFileCache = @{}
$script:DashboardPayloadCache = @{}
$script:WorkspaceSessionPayloadCache = @{}
$script:SearchEntriesCache = @{}
$script:FinancialPortalMetadataCache = @{}
$script:RuntimeCacheMetrics = [ordered]@{
    json = [ordered]@{
        hits = 0
        misses = 0
        evictions = 0
    }
    dashboard = [ordered]@{
        hits = 0
        misses = 0
        evictions = 0
    }
    workspace = [ordered]@{
        hits = 0
        misses = 0
        evictions = 0
    }
    financialPortal = [ordered]@{
        hits = 0
        misses = 0
        evictions = 0
    }
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-IsoNow {
    (Get-Date).ToString('s')
}

function Get-TextFingerprint {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text = ''
    )

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-ApiContractCatalog {
    return [ordered]@{
        workspace = $script:WorkspaceSessionPayloadContractVersion
        dashboard = $script:DashboardPayloadContractVersion
        contractAudit = $script:ContractAuditPayloadContractVersion
        contractDetail = $script:ContractDetailPayloadContractVersion
        contractCollection = $script:ContractCollectionPayloadContractVersion
        status = $script:StatusPayloadContractVersion
        publicStatus = $script:PublicStatusPayloadContractVersion
        search = $script:SearchPayloadContractVersion
    }
}

function Get-ApiContractVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    switch (([string]$Name).Trim()) {
        'workspace' { return $script:WorkspaceSessionPayloadContractVersion }
        'dashboard' { return $script:DashboardPayloadContractVersion }
        'contractAudit' { return $script:ContractAuditPayloadContractVersion }
        'contractDetail' { return $script:ContractDetailPayloadContractVersion }
        'contractCollection' { return $script:ContractCollectionPayloadContractVersion }
        'status' { return $script:StatusPayloadContractVersion }
        'publicStatus' { return $script:PublicStatusPayloadContractVersion }
        'search' { return $script:SearchPayloadContractVersion }
        default { return '' }
    }
}

function New-ApiContractDescriptor {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return [ordered]@{
        name = Collapse-Whitespace -Text $Name
        version = Get-ApiContractVersion -Name $Name
    }
}

function Get-DefaultStatus {
    [ordered]@{
        isSyncRunning = $false
        syncStage = 'idle'
        message = 'Aguardando sincronizacao.'
        syncStartedAt = $null
        syncFinishedAt = $null
        scannedPages = 0
        totalPages = 0
        totalDiaries = 0
        newDiaries = 0
        updatedDiaries = 0
        downloadedPdfCount = 0
        candidateDiaries = 0
        analyzedDiaries = 0
        pendingAnalysis = 0
        lastError = $null
        lastSuccessfulSyncAt = $null
        syncHistory = @()
    }
}

function Get-PublicStatusPayload {
    $status = Get-StatusHashtable
    $contractsAggregate = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
    $pendingTasksPreview = @(Get-PendingAnalysisTasks -Limit 3)
    $dashboardGeneratedAt = [string]$contractsAggregate.generatedAt
    $lastSuccessfulSyncAt = [string]$status.lastSuccessfulSyncAt
    $trackedContractsCount = [int]$(if ($null -ne $contractsAggregate.totalItems) { $contractsAggregate.totalItems } else { 0 })
    $pendingAnalysisCount = [int]$(if ($null -ne $status.pendingAnalysis) { $status.pendingAnalysis } else { 0 })
    $formatStatusDate = {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$Value
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return ''
        }

        try {
            return ([DateTime]::Parse($Value)).ToString('dd/MM/yyyy HH:mm')
        }
        catch {
            return [string]$Value
        }
    }

    $availabilityKey = 'server_only'
    $statusTone = 'neutral'
    $statusLabel = 'Servidor iniciado'
    $headline = 'Servidor local online'
    $summary = 'Entre com seu login institucional para acessar o painel.'
    $nextStep = 'Use o login de 4 digitos. No primeiro acesso, a troca de senha sera solicitada logo apos a autenticacao.'
    $environmentLabel = 'Ambiente local'
    $dataFreshnessLabel = 'Sem base consolidada'
    $securityLabel = 'Bloqueio e sessao local'
    $securityNote = 'Tentativas seguidas geram bloqueio temporario, e o painel libera sessao individual por perfil depois da autenticacao.'

    if ([bool]$status.isSyncRunning) {
        $availabilityKey = 'syncing'
        $statusTone = 'progress'
        $statusLabel = 'Atualizando base'
        $headline = 'Sincronizacao em andamento'
        $summary = if ([string]::IsNullOrWhiteSpace([string]$status.message)) {
            'O servidor esta online e a base local esta sendo atualizada agora.'
        }
        else {
            [string]$status.message
        }
        $nextStep = 'O acesso ja pode ser feito. A atualizacao seguira em segundo plano ate a base consolidada ficar pronta.'
        $environmentLabel = 'Atualizacao em andamento'
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$status.lastError)) {
        $availabilityKey = 'degraded'
        $statusTone = 'attention'
        $statusLabel = 'Atencao operacional'
        $headline = 'Servidor online com alerta de sincronizacao'
        $summary = [string]$status.lastError
        $nextStep = if (-not [string]::IsNullOrWhiteSpace($dashboardGeneratedAt)) {
            'O acesso continua disponivel, mas vale conferir a ultima atualizacao antes do uso diario.'
        }
        else {
            'Antes de liberar o uso diario, vale conferir a sincronizacao inicial da base local.'
        }
        $environmentLabel = 'Servidor com alerta'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($dashboardGeneratedAt)) {
        $availabilityKey = 'ready'
        $statusTone = 'healthy'
        $statusLabel = 'Base pronta'
        $headline = 'Base local pronta para autenticacao'
        $summary = 'Servidor e base consolidada disponiveis para o login institucional.'
        $environmentLabel = 'Servidor e base prontos'
    }
    else {
        $availabilityKey = 'warming'
        $statusTone = 'neutral'
        $statusLabel = 'Aguardando carga'
        $headline = 'Servidor online sem base consolidada'
        $summary = 'O painel ja esta no ar, mas ainda nao ha uma leitura consolidada pronta para consulta.'
        $nextStep = 'Se este estado persistir, rode a sincronizacao antes de liberar o uso diario do painel.'
        $environmentLabel = 'Servidor sem base pronta'
    }

    if (-not [string]::IsNullOrWhiteSpace($dashboardGeneratedAt)) {
        $dataFreshnessLabel = "Base consolidada em $(& $formatStatusDate $dashboardGeneratedAt)"
    }
    elseif (-not [string]::IsNullOrWhiteSpace($lastSuccessfulSyncAt)) {
        $dataFreshnessLabel = "Ultimo sync em $(& $formatStatusDate $lastSuccessfulSyncAt)"
    }

    if ($availabilityKey -eq 'degraded') {
        $securityNote = 'O acesso continua protegido por bloqueio temporario de tentativas e por sessao individual, mas a ultima sincronizacao precisa de conferencia.'
    }
    elseif ($availabilityKey -eq 'warming') {
        $securityNote = 'O login institucional ja esta protegido, mas a base ainda nao terminou a carga inicial para uso diario.'
    }

    return [ordered]@{
        available = $true
        availabilityKey = $availabilityKey
        statusTone = $statusTone
        statusLabel = $statusLabel
        headline = $headline
        summary = $summary
        nextStep = $nextStep
        environmentLabel = $environmentLabel
        dataFreshnessLabel = $dataFreshnessLabel
        securityLabel = $securityLabel
        securityNote = $securityNote
        isSyncRunning = [bool]$status.isSyncRunning
        syncStage = [string]$status.syncStage
        syncMessage = [string]$status.message
        dashboardGeneratedAt = $dashboardGeneratedAt
        lastSuccessfulSyncAt = $lastSuccessfulSyncAt
        trackedContractsCount = $trackedContractsCount
        pendingAnalysisCount = $pendingAnalysisCount
        pendingTasksPreview = $pendingTasksPreview
        apiContract = New-ApiContractDescriptor -Name 'publicStatus'
        apiContracts = Get-ApiContractCatalog
    }
}

function Get-EmptyDiariesPayload {
    [ordered]@{
        generatedAt = $null
        source = ([Uri]::new($script:BasePortalUri, $script:PortalDiarioPath)).AbsoluteUri
        diaries = @()
    }
}

function Get-EmptyContractsPayload {
    [ordered]@{
        generatedAt = $null
        parserVersion = $script:ParserVersion
        totalItems = 0
        totalValue = 0
        analyzedDiaryCount = 0
        officialPortalContracts = 0
        uniqueSuppliers = 0
        qualitySummary = [ordered]@{
            confirmedItems = 0
            procurementItems = 0
            managementItems = 0
            flaggedItems = 0
            highConfidenceItems = 0
        }
        typeSummary = @()
        organizationSummary = @()
        managementSummary = [ordered]@{
            trackedContracts = 0
            withManager = 0
            withInspector = 0
            withoutManager = 0
            withoutInspector = 0
            managerChanged = 0
            inspectorChanged = 0
            exonerationSignals = 0
            atRisk = 0
        }
        crosswalkSummary = [ordered]@{
            officialMatched = 0
            officialUnmatched = 0
            officialPendingReview = 0
            movementMatched = 0
            movementUnmatched = 0
            movementPendingReview = 0
            automaticMatches = 0
            reviewedMatches = 0
            divergences = 0
            suppressedDivergences = 0
            operationalAlerts = 0
            latestMatchedAt = $null
        }
        analyses = @()
        managementProfiles = @()
        crossReviewQueue = @()
        crossSourceDivergences = @()
        crossSourceAlerts = @()
        crossSourceSuppressionSummary = [ordered]@{
            total = 0
            reasons = @()
        }
        financialMonitoring = [ordered]@{
            mode = 'partial'
            modeLabel = 'Integracao parcial'
            note = ''
            monitoredContracts = 0
            searchableContracts = 0
            queryReadyContracts = 0
            withContractValue = 0
            expenseContracts = 0
            revenueContracts = 0
            automatedContracts = 0
            assistedContracts = 0
            limitedContracts = 0
            unmappedContracts = 0
            averageCoverageScore = 0
            automationReadySources = 0
            sourceCount = 0
            executionStageCount = 0
            detailSectionCount = 0
            expensePortal = $null
            coverageBreakdown = @()
            stageSummary = @()
            sources = @()
        }
        officialContracts = @()
        contractMovements = @()
        items = @()
    }
}

function Get-EmptyPortalContractsPayload {
    [ordered]@{
        generatedAt = $null
        source = ([Uri]::new($script:BasePortalUri, '/portal/contratos')).AbsoluteUri
        totalItems = 0
        downloadedDocumentCount = 0
        items = @()
    }
}

function Get-EmptyContractCrossReviewPayload {
    [ordered]@{
        generatedAt = $null
        version = $script:ContractCrossReviewSchemaVersion
        decisions = @()
    }
}

function Get-EmptyWorkspacePayload {
    [ordered]@{
        generatedAt = $null
        version = $script:WorkspaceSchemaVersion
        favorites = @()
        savedViews = @()
        contractNotes = @()
        workflowItems = @()
        alertStates = @()
        notificationStates = @()
        activityLog = @()
        aggregateSnapshots = @()
        recentChanges = (Get-EmptyWorkspaceAggregateChangeSummary)
    }
}

function Get-EmptyWorkspaceAggregateChangeSummary {
    return [ordered]@{
        currentGeneratedAt = $null
        previousGeneratedAt = $null
        headline = 'Sem historico de base registrado.'
        windowLabel = ''
        summaryText = 'As comparacoes entre sincronizacoes apareceram aqui conforme o painel registrar novas recomposicoes.'
        changedContractsCount = 0
        items = @()
        contractVersionChanges = @()
        supplierMovers = @()
        detailGroups = @()
    }
}

function Set-WorkspacePayloadRecentChanges {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$RecentChanges = $null
    )

    $resolvedRecentChanges = if ($null -eq $RecentChanges) {
        Get-EmptyWorkspaceAggregateChangeSummary
    }
    else {
        $RecentChanges
    }

    if (-not (Test-ObjectProperty -Item $Payload -Name 'recentChanges')) {
        $Payload | Add-Member -NotePropertyName recentChanges -NotePropertyValue $resolvedRecentChanges
    }
    else {
        $Payload.recentChanges = $resolvedRecentChanges
    }
}

function Get-EmptyObservabilityPayload {
    [ordered]@{
        generatedAt = $null
        version = $script:ObservabilitySchemaVersion
        events = @()
        counters = [ordered]@{}
    }
}

function Get-EmptyOrganizationCatalog {
    [ordered]@{
        version = $null
        checkedAt = $null
        notes = @()
        sources = @()
        areas = @()
        organizations = @()
    }
}

function Get-FileCacheStamp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return 'missing'
    }

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return 'missing'
    }

    return '{0}|{1}' -f [int64]$item.Length, [int64]$item.LastWriteTimeUtc.Ticks
}

function Add-ScriptCacheEntry {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Cache,

        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [object]$Value,

        [Parameter(Mandatory = $false)]
        [int]$MaxEntries = 12,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CacheName = ''
    )

    if ($Cache.Count -ge $MaxEntries) {
        if (-not [string]::IsNullOrWhiteSpace($CacheName)) {
            Add-CacheMetric -CacheName $CacheName -Metric 'evictions'
        }
        $Cache.Clear()
    }

    $Cache[$Key] = $Value
}

function Add-CacheMetric {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('json', 'dashboard', 'workspace', 'financialPortal')]
        [string]$CacheName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('hits', 'misses', 'evictions')]
        [string]$Metric,

        [Parameter(Mandatory = $false)]
        [int]$Amount = 1
    )

    if (-not $script:RuntimeCacheMetrics.Contains($CacheName)) {
        return
    }

    $bucket = $script:RuntimeCacheMetrics[$CacheName]
    $currentValue = 0
    if ($bucket.Contains($Metric)) {
        $currentValue = [int]$bucket[$Metric]
    }
    $bucket[$Metric] = [int]($currentValue + [Math]::Max($Amount, 0))
}

function New-CacheMetricSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [hashtable]$Cache
    )

    $bucket = if ($script:RuntimeCacheMetrics.Contains($Key)) { $script:RuntimeCacheMetrics[$Key] } else { [ordered]@{} }
    $hits = if ($bucket.Contains('hits')) { [int]$bucket['hits'] } else { 0 }
    $misses = if ($bucket.Contains('misses')) { [int]$bucket['misses'] } else { 0 }
    $evictions = if ($bucket.Contains('evictions')) { [int]$bucket['evictions'] } else { 0 }
    $requests = $hits + $misses
    $hitRate = if ($requests -gt 0) { [math]::Round(($hits / $requests) * 100, 0) } else { 0 }

    return [pscustomobject][ordered]@{
        key = $Key
        label = $Label
        hits = $hits
        misses = $misses
        requests = $requests
        hitRate = [int]$hitRate
        evictions = $evictions
        entries = [int]$Cache.Count
    }
}

function Get-CacheObservabilitySnapshot {
    $cacheItems = @(
        (New-CacheMetricSnapshot -Key 'json' -Label 'Arquivos JSON' -Cache $script:JsonFileCache)
        (New-CacheMetricSnapshot -Key 'dashboard' -Label 'Payload do dashboard' -Cache $script:DashboardPayloadCache)
        (New-CacheMetricSnapshot -Key 'workspace' -Label 'Sessao do workspace' -Cache $script:WorkspaceSessionPayloadCache)
        (New-CacheMetricSnapshot -Key 'financialPortal' -Label 'Portal financeiro legado' -Cache $script:FinancialPortalMetadataCache)
    )
    $totalHits = @($cacheItems | Measure-Object -Property hits -Sum).Sum
    $totalMisses = @($cacheItems | Measure-Object -Property misses -Sum).Sum
    $totalRequests = [int]($totalHits + $totalMisses)
    $totalEntries = [int](@($cacheItems | Measure-Object -Property entries -Sum).Sum)
    $overallHitRate = if ($totalRequests -gt 0) { [math]::Round(($totalHits / $totalRequests) * 100, 0) } else { 0 }

    return [ordered]@{
        caches = @($cacheItems)
        summary = [ordered]@{
            hits = [int]$totalHits
            misses = [int]$totalMisses
            requests = $totalRequests
            hitRate = [int]$overallHitRate
            activeEntries = $totalEntries
        }
    }
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [object]$Data
    )

    $parent = Split-Path -Parent $Path
    if ($parent) {
        Ensure-Directory -Path $parent
    }

    $tempPath = "$Path.tmp"
    $json = $Data | ConvertTo-Json -Depth 100
    Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8
    Move-Item -LiteralPath $tempPath -Destination $Path -Force
    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $script:JsonFileCache.Remove($resolvedPath) | Out-Null
}

function Read-JsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [object]$Default = $null
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        $script:JsonFileCache.Remove($resolvedPath) | Out-Null
        return $Default
    }

    $stamp = Get-FileCacheStamp -Path $resolvedPath
    $raw = $null
    $cached = $script:JsonFileCache[$resolvedPath]
    if ($cached -and [string]$cached.stamp -eq $stamp) {
        Add-CacheMetric -CacheName 'json' -Metric 'hits'
        $raw = [string]$cached.raw
    }
    else {
        Add-CacheMetric -CacheName 'json' -Metric 'misses'
        $raw = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8
        Add-ScriptCacheEntry -Cache $script:JsonFileCache -Key $resolvedPath -Value ([ordered]@{
            stamp = $stamp
            raw = $raw
        }) -MaxEntries 20 -CacheName 'json'
    }

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $Default
    }

    try {
        return $raw | ConvertFrom-Json
    }
    catch {
        return $Default
    }
}

function Get-OrganizationCatalog {
    return Read-JsonFile -Path $script:OrganizationCatalogPath -Default (Get-EmptyOrganizationCatalog)
}

function Get-StatusHashtable {
    $raw = Read-JsonFile -Path $script:StatusPath -Default $null
    $status = Get-DefaultStatus

    if ($null -eq $raw) {
        return $status
    }

    foreach ($property in $raw.PSObject.Properties) {
        $status[$property.Name] = $property.Value
    }

    return $status
}

function Update-Status {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Updates
    )

    $status = Get-StatusHashtable
    foreach ($key in $Updates.Keys) {
        $status[$key] = $Updates[$key]
    }

    Write-JsonFile -Path $script:StatusPath -Data $status
}

function Register-SyncHistoryEntry {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('success', 'error')]
        [string]$Outcome,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Metrics = $null
    )

    $status = Get-StatusHashtable
    $history = @($status.syncHistory)
    $entry = [pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        outcome = $Outcome
        message = Collapse-Whitespace -Text $Message
        createdAt = Get-IsoNow
        metrics = $Metrics
    }

    $status.syncHistory = @($history) + @($entry) | Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true } | Select-Object -First 25
    Write-JsonFile -Path $script:StatusPath -Data $status
    Register-ObservabilityEvent -Type 'sync' -Status $Outcome -Message $Message -Metadata $Metrics | Out-Null
    return $entry
}

function Remove-DiacriticsCommon {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Collapse-Whitespace {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    return (($Text -replace "[\r\n\t]+", ' ') -replace '\s{2,}', ' ').Trim()
}

function Normalize-IndexText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $normalized = Remove-DiacriticsCommon -Text $Text
    $normalized = $normalized.ToUpperInvariant()
    $normalized = $normalized -replace '[^A-Z0-9/\s]', ' '
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Clean-PersonDisplayName {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $clean = Collapse-Whitespace -Text $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ''
    }

    $clean = $clean -replace '^(?:O|A)\s+', ''
    $clean = $clean -replace ',?\s*(?:inscrit[oa].*|titular da.*|portador.*|cpf.*|rg.*)$', ''
    $clean = $clean.Trim(" -,:;.")
    $wordCount = @($clean -split '\s+' | Where-Object { $_ }).Count
    if ($wordCount -lt 2 -or $wordCount -gt 8) {
        return ''
    }

    return $clean
}

function Clean-RoleDisplayText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $clean = Collapse-Whitespace -Text $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ''
    }

    $clean = $clean -replace ',?\s*(?:inscrit[oa].*|titular da.*|cpf.*|rg.*)$', ''
    return $clean.Trim(" -,:;.")
}

function Get-ContractReferenceTokens {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ContractNumber,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ProcessNumber
    )

    $tokenSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $values = @($ContractNumber, $ProcessNumber)

    foreach ($value in $values) {
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            continue
        }

        $compact = ([string]$value -replace '\s+', '').ToUpperInvariant()
        if (-not [string]::IsNullOrWhiteSpace($compact)) {
            $null = $tokenSet.Add($compact)
        }

        $matches = [regex]::Matches($compact, '(?<number>\d{1,10})/(?<year>\d{4})')
        foreach ($match in $matches) {
            $number = [int]$match.Groups['number'].Value
            $year = [string]$match.Groups['year'].Value
            $null = $tokenSet.Add("$number/$year")
        }
    }

    return @($tokenSet)
}

function Get-PrimaryContractReferenceToken {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ContractNumber,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ProcessNumber
    )

    $tokens = @(Get-ContractReferenceTokens -ContractNumber $ContractNumber -ProcessNumber $ProcessNumber)
    foreach ($token in $tokens) {
        if ($token -match '^\d+/\d{4}$') {
            return $token
        }
    }

    if (@($tokens).Count -gt 0) {
        return [string]$tokens[0]
    }

    return ''
}

function Get-ManagementTextSource {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    return Collapse-Whitespace -Text ([string]$Item.excerpt)
}

function Get-ManagementActionType {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $normalized = Normalize-IndexText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return 'designacao'
    }

    if ($normalized -match '\bEXONER') {
        return 'exoneracao'
    }

    if ($normalized -match '\bSUBSTITU' -or
        $normalized -match '\bALTERACAO DO ACOMPANHAMENTO E FISCALIZACAO\b' -or
        $normalized -match '\bONDE SE LE\b' -or
        $normalized -match '\bLEIA SE\b' -or
        $normalized -match '\bALTERA(?:R|CAO)\b.{0,40}\b(?:GESTOR|FISCAL)\b') {
        return 'alteracao'
    }

    return 'designacao'
}

function Get-ManagementRoleMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        $match = [regex]::Match(
            $Text,
            $pattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        if (-not $match.Success) {
            continue
        }

        $name = Clean-PersonDisplayName -Text ([string]$match.Groups['name'].Value)
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $role = Clean-RoleDisplayText -Text ([string]$match.Groups['role'].Value)
        return [pscustomobject]@{
            name = $name
            normalizedName = (Normalize-IndexText -Text $name)
            role = $role
        }
    }

    return $null
}

function Get-ManagementAssignmentsFromText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $compact = Collapse-Whitespace -Text $Text
    if ([string]::IsNullOrWhiteSpace($compact)) {
        return [pscustomobject]@{
            manager = $null
            inspector = $null
            actionType = 'designacao'
        }
    }

    $managerPatterns = @(
        '(?is)(?:FICAM?\s+DESIGNADOS?\s+|FICA\s+DESIGNADO\s+|DESIGNA,\s*|DESIGNA\s+)(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|\s+CPF|\s+RG|\s+PARA\s+EXERCER)).{0,160}?\s+PARA\s+EXERCER\s+A?\s*FUNCAO\s+DE\s+GESTOR',
        '(?is)(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|\s+CPF|\s+RG|\s+PARA\s+EXERCER)).{0,160}?\s+PARA\s+EXERCER\s+A?\s*FUNCAO\s+DE\s+GESTOR'
    )
    $inspectorPatterns = @(
        '(?is)(?:,\s*E\s+|FICAM?\s+DESIGNADOS?\s+|FICA\s+DESIGNADO\s+|DESIGNA,\s*|DESIGNA\s+)(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|\s+CPF|\s+RG|\s+PARA\s+(?:ATUAR|EXERCER))).{0,180}?\s+PARA\s+(?:ATUAR\s+COMO|EXERCER\s+A?\s*FUNCAO\s+DE)\s+FISCAL',
        '(?is)(?<name>[^,\n]{4,120}?)\s*,\s*(?<role>.*?)(?=(?:,\s*INSCRIT|\s+INSCRIT|,\s*TITULAR|\s+TITULAR|\s+CPF|\s+RG|\s+PARA\s+(?:ATUAR|EXERCER))).{0,180}?\s+PARA\s+(?:ATUAR\s+COMO|EXERCER\s+A?\s*FUNCAO\s+DE)\s+FISCAL'
    )

    return [pscustomobject]@{
        manager = (Get-ManagementRoleMatch -Text $compact -Patterns $managerPatterns)
        inspector = (Get-ManagementRoleMatch -Text $compact -Patterns $inspectorPatterns)
        actionType = (Get-ManagementActionType -Text $compact)
    }
}

function Get-PersonnelExonerationIndex {
    $eventIndex = @{}
    $files = @(Get-ChildItem -LiteralPath $script:PersonnelAnalysisRoot -Filter '*.json' -File -ErrorAction SilentlyContinue)

    foreach ($file in $files) {
        $analysis = Read-JsonFile -Path $file.FullName -Default $null
        if ($null -eq $analysis) {
            continue
        }

        if ([string]$analysis.parserVersion -ne $script:PersonnelParserVersion) {
            continue
        }

        foreach ($event in @($analysis.events)) {
            if ([string]$event.type -ne 'exoneracao') {
                continue
            }

            $normalizedName = Normalize-IndexText -Text ([string]$event.normalizedName)
            if ([string]::IsNullOrWhiteSpace($normalizedName)) {
                $normalizedName = Normalize-IndexText -Text ([string]$event.personName)
            }
            if ([string]::IsNullOrWhiteSpace($normalizedName)) {
                continue
            }

            if (-not $eventIndex.ContainsKey($normalizedName)) {
                $eventIndex[$normalizedName] = New-Object System.Collections.Generic.List[object]
            }

            $eventIndex[$normalizedName].Add([pscustomobject]@{
                type = [string]$event.type
                personName = [string]$event.personName
                normalizedName = $normalizedName
                publishedAt = [string]$event.publishedAt
                diaryId = [string]$event.diaryId
                edition = [string]$event.edition
                pageNumber = [int]$(if ($null -ne $event.pageNumber) { $event.pageNumber } else { 0 })
                excerpt = [string]$event.excerpt
            })
        }
    }

    return $eventIndex
}

function Get-ManagementStatusTone {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$HasManager,

        [Parameter(Mandatory = $true)]
        [bool]$HasInspector,

        [Parameter(Mandatory = $true)]
        [bool]$ManagerChanged,

        [Parameter(Mandatory = $true)]
        [bool]$InspectorChanged,

        [Parameter(Mandatory = $true)]
        [bool]$ManagerExonerationSignal,

        [Parameter(Mandatory = $true)]
        [bool]$InspectorExonerationSignal
    )

    if ((-not $HasManager) -or (-not $HasInspector) -or $ManagerExonerationSignal -or $InspectorExonerationSignal) {
        return 'critical'
    }

    if ($ManagerChanged -or $InspectorChanged) {
        return 'warning'
    }

    return 'stable'
}

function Build-ContractManagementProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items,

        [Parameter(Mandatory = $false)]
        [hashtable]$PersonnelExonerationIndex = @{}
    )

    $groups = @{}

    foreach ($item in @($Items)) {
        if ([string]$item.recordClass -ne 'gestao_contratual') {
            continue
        }

        $primaryToken = Get-PrimaryContractReferenceToken -ContractNumber ([string]$item.contractNumber) -ProcessNumber ([string]$item.processNumber)
        if ([string]::IsNullOrWhiteSpace($primaryToken)) {
            continue
        }

        if (-not $groups.ContainsKey($primaryToken)) {
            $groups[$primaryToken] = [ordered]@{
                contractKey = $primaryToken
                referenceTokens = New-Object 'System.Collections.Generic.HashSet[string]'
                events = New-Object System.Collections.Generic.List[object]
            }
        }

        foreach ($token in @(Get-ContractReferenceTokens -ContractNumber ([string]$item.contractNumber) -ProcessNumber ([string]$item.processNumber))) {
            if (-not [string]::IsNullOrWhiteSpace($token)) {
                $null = $groups[$primaryToken].referenceTokens.Add([string]$token)
            }
        }

        $assignments = Get-ManagementAssignmentsFromText -Text (Get-ManagementTextSource -Item $item)
        $publishedAt = [string]$item.publishedAt
        $publishedTimestamp = 0
        try {
            if (-not [string]::IsNullOrWhiteSpace($publishedAt)) {
                $publishedTimestamp = [DateTime]::Parse($publishedAt).Ticks
            }
        }
        catch {
            $publishedTimestamp = 0
        }

        $groups[$primaryToken].events.Add([pscustomobject]@{
            publishedAt = $publishedAt
            publishedTimestamp = $publishedTimestamp
            diaryId = [string]$item.diaryId
            edition = [string]$item.edition
            actTitle = [string]$item.actTitle
            excerpt = [string]$item.excerpt
            contractNumber = [string]$item.contractNumber
            processNumber = [string]$item.processNumber
            manager = $assignments.manager
            inspector = $assignments.inspector
            actionType = [string]$assignments.actionType
        })
    }

    $profiles = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $groups.GetEnumerator()) {
        $sortedEvents = @($entry.Value.events | Sort-Object -Property publishedTimestamp)
        $latestManagerEvent = $sortedEvents |
            Where-Object { $null -ne $_.manager -and -not [string]::IsNullOrWhiteSpace([string]$_.manager.name) } |
            Sort-Object -Property publishedTimestamp -Descending |
            Select-Object -First 1
        $latestInspectorEvent = $sortedEvents |
            Where-Object { $null -ne $_.inspector -and -not [string]::IsNullOrWhiteSpace([string]$_.inspector.name) } |
            Sort-Object -Property publishedTimestamp -Descending |
            Select-Object -First 1

        $managerHistory = @(
            $sortedEvents |
            Where-Object { $null -ne $_.manager -and -not [string]::IsNullOrWhiteSpace([string]$_.manager.name) } |
            Sort-Object -Property publishedTimestamp -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    name = [string]$_.manager.name
                    normalizedName = [string]$_.manager.normalizedName
                    role = [string]$_.manager.role
                    publishedAt = [string]$_.publishedAt
                    actTitle = [string]$_.actTitle
                    actionType = [string]$_.actionType
                    diaryId = [string]$_.diaryId
                    edition = [string]$_.edition
                }
            }
        )
        $inspectorHistory = @(
            $sortedEvents |
            Where-Object { $null -ne $_.inspector -and -not [string]::IsNullOrWhiteSpace([string]$_.inspector.name) } |
            Sort-Object -Property publishedTimestamp -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    name = [string]$_.inspector.name
                    normalizedName = [string]$_.inspector.normalizedName
                    role = [string]$_.inspector.role
                    publishedAt = [string]$_.publishedAt
                    actTitle = [string]$_.actTitle
                    actionType = [string]$_.actionType
                    diaryId = [string]$_.diaryId
                    edition = [string]$_.edition
                }
            }
        )

        $managerNames = @($managerHistory | ForEach-Object { [string]$_.normalizedName } | Where-Object { $_ } | Select-Object -Unique)
        $inspectorNames = @($inspectorHistory | ForEach-Object { [string]$_.normalizedName } | Where-Object { $_ } | Select-Object -Unique)

        $managerChanged = @($managerNames).Count -gt 1
        $inspectorChanged = @($inspectorNames).Count -gt 1

        $managerExonerationEvents = @()
        if ($latestManagerEvent -and $latestManagerEvent.manager -and $PersonnelExonerationIndex.ContainsKey([string]$latestManagerEvent.manager.normalizedName)) {
            $managerExonerationEvents = @(
                $PersonnelExonerationIndex[[string]$latestManagerEvent.manager.normalizedName] |
                Where-Object {
                    if ([string]::IsNullOrWhiteSpace([string]$_.publishedAt) -or [string]::IsNullOrWhiteSpace([string]$latestManagerEvent.publishedAt)) {
                        return $false
                    }

                    try {
                        return [DateTime]::Parse([string]$_.publishedAt) -ge [DateTime]::Parse([string]$latestManagerEvent.publishedAt)
                    }
                    catch {
                        return $false
                    }
                } |
                Sort-Object -Property publishedAt -Descending
            )
        }

        $inspectorExonerationEvents = @()
        if ($latestInspectorEvent -and $latestInspectorEvent.inspector -and $PersonnelExonerationIndex.ContainsKey([string]$latestInspectorEvent.inspector.normalizedName)) {
            $inspectorExonerationEvents = @(
                $PersonnelExonerationIndex[[string]$latestInspectorEvent.inspector.normalizedName] |
                Where-Object {
                    if ([string]::IsNullOrWhiteSpace([string]$_.publishedAt) -or [string]::IsNullOrWhiteSpace([string]$latestInspectorEvent.publishedAt)) {
                        return $false
                    }

                    try {
                        return [DateTime]::Parse([string]$_.publishedAt) -ge [DateTime]::Parse([string]$latestInspectorEvent.publishedAt)
                    }
                    catch {
                        return $false
                    }
                } |
                Sort-Object -Property publishedAt -Descending
            )
        }

        $hasManager = $null -ne $latestManagerEvent
        $hasInspector = $null -ne $latestInspectorEvent
        $managerExonerationSignal = @($managerExonerationEvents).Count -gt 0
        $inspectorExonerationSignal = @($inspectorExonerationEvents).Count -gt 0
        $tone = Get-ManagementStatusTone `
            -HasManager:$hasManager `
            -HasInspector:$hasInspector `
            -ManagerChanged:$managerChanged `
            -InspectorChanged:$inspectorChanged `
            -ManagerExonerationSignal:$managerExonerationSignal `
            -InspectorExonerationSignal:$inspectorExonerationSignal

        $summaryParts = New-Object System.Collections.Generic.List[string]
        if (-not $hasManager) { $summaryParts.Add('Sem gestor atual identificado') }
        if (-not $hasInspector) { $summaryParts.Add('Sem fiscal atual identificado') }
        if ($managerChanged) { $summaryParts.Add('Gestor ja foi alterado') }
        if ($inspectorChanged) { $summaryParts.Add('Fiscal ja foi alterado') }
        if ($managerExonerationSignal) { $summaryParts.Add('Há sinal de exoneração do gestor atual') }
        if ($inspectorExonerationSignal) { $summaryParts.Add('Há sinal de exoneração do fiscal atual') }
        if (@($summaryParts).Count -eq 0) { $summaryParts.Add('Gestor e fiscal atuais identificados no Diário') }

        $profiles.Add([pscustomobject]@{
            contractKey = [string]$entry.Value.contractKey
            referenceTokens = @($entry.Value.referenceTokens)
            contractNumber = if ($latestManagerEvent -and [string]$latestManagerEvent.contractNumber) { [string]$latestManagerEvent.contractNumber } elseif ($latestInspectorEvent -and [string]$latestInspectorEvent.contractNumber) { [string]$latestInspectorEvent.contractNumber } else { [string]$entry.Value.contractKey }
            processNumber = if ($latestManagerEvent -and [string]$latestManagerEvent.processNumber) { [string]$latestManagerEvent.processNumber } elseif ($latestInspectorEvent -and [string]$latestInspectorEvent.processNumber) { [string]$latestInspectorEvent.processNumber } else { '' }
            hasManager = [bool]$hasManager
            hasInspector = [bool]$hasInspector
            managerName = if ($hasManager) { [string]$latestManagerEvent.manager.name } else { '' }
            managerRole = if ($hasManager) { [string]$latestManagerEvent.manager.role } else { '' }
            managerAssignedAt = if ($hasManager) { [string]$latestManagerEvent.publishedAt } else { $null }
            inspectorName = if ($hasInspector) { [string]$latestInspectorEvent.inspector.name } else { '' }
            inspectorRole = if ($hasInspector) { [string]$latestInspectorEvent.inspector.role } else { '' }
            inspectorAssignedAt = if ($hasInspector) { [string]$latestInspectorEvent.publishedAt } else { $null }
            managerChanged = [bool]$managerChanged
            inspectorChanged = [bool]$inspectorChanged
            managerChangeCount = [Math]::Max(@($managerNames).Count - 1, 0)
            inspectorChangeCount = [Math]::Max(@($inspectorNames).Count - 1, 0)
            managerExonerationSignal = [bool]$managerExonerationSignal
            inspectorExonerationSignal = [bool]$inspectorExonerationSignal
            managerExonerationAt = if ($managerExonerationSignal) { [string]$managerExonerationEvents[0].publishedAt } else { $null }
            inspectorExonerationAt = if ($inspectorExonerationSignal) { [string]$inspectorExonerationEvents[0].publishedAt } else { $null }
            managerExonerationExcerpt = if ($managerExonerationSignal) { [string]$managerExonerationEvents[0].excerpt } else { '' }
            inspectorExonerationExcerpt = if ($inspectorExonerationSignal) { [string]$inspectorExonerationEvents[0].excerpt } else { '' }
            lastManagementActAt = if (@($sortedEvents).Count -gt 0) { [string]$sortedEvents[-1].publishedAt } else { $null }
            lastManagementActTitle = if (@($sortedEvents).Count -gt 0) { [string]$sortedEvents[-1].actTitle } else { '' }
            linkedMovementCount = @($sortedEvents).Count
            statusTone = $tone
            summary = [string]($summaryParts -join ' | ')
            managerExonerationEvents = @($managerExonerationEvents)
            inspectorExonerationEvents = @($inspectorExonerationEvents)
            managementEvents = @($sortedEvents)
            managerHistory = @($managerHistory)
            inspectorHistory = @($inspectorHistory)
        })
    }

    return @(
        $profiles |
        Sort-Object -Property `
            @{ Expression = { $_.lastManagementActAt }; Descending = $true }, `
            @{ Expression = { $_.contractKey }; Descending = $false }
    )
}

function Find-ManagementProfileForItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [hashtable]$ProfileIndex
    )

    foreach ($token in @(Get-ContractReferenceTokens -ContractNumber ([string]$Item.contractNumber) -ProcessNumber ([string]$Item.processNumber))) {
        if ($ProfileIndex.ContainsKey([string]$token)) {
            return $ProfileIndex[[string]$token]
        }
    }

    return $null
}

function Add-ManagementFieldsToItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [hashtable]$ProfileIndex
    )

    $profile = Find-ManagementProfileForItem -Item $Item -ProfileIndex $ProfileIndex
    $tokens = @(Get-ContractReferenceTokens -ContractNumber ([string]$Item.contractNumber) -ProcessNumber ([string]$Item.processNumber))
    $data = [ordered]@{}

    foreach ($property in $Item.PSObject.Properties) {
        $data[$property.Name] = $property.Value
    }

    $data.managementTracked = @($tokens).Count -gt 0
    $data.managementProfileKey = if ($profile) { [string]$profile.contractKey } else { $null }
    $data.referenceKey = if ($profile) {
        [string]$profile.contractKey
    }
    elseif (@($tokens).Count -gt 0) {
        [string]$tokens[0]
    }
    elseif ($data.Contains('portalContractId') -and -not [string]::IsNullOrWhiteSpace([string]$data.portalContractId)) {
        [string]$data.portalContractId
    }
    else {
        $null
    }
    $data.hasManager = if ($profile) { [bool]$profile.hasManager } else { $false }
    $data.hasInspector = if ($profile) { [bool]$profile.hasInspector } else { $false }
    $data.managerName = if ($profile) { [string]$profile.managerName } else { '' }
    $data.managerRole = if ($profile) { [string]$profile.managerRole } else { '' }
    $data.managerAssignedAt = if ($profile) { [string]$profile.managerAssignedAt } else { $null }
    $data.inspectorName = if ($profile) { [string]$profile.inspectorName } else { '' }
    $data.inspectorRole = if ($profile) { [string]$profile.inspectorRole } else { '' }
    $data.inspectorAssignedAt = if ($profile) { [string]$profile.inspectorAssignedAt } else { $null }
    $data.managerChanged = if ($profile) { [bool]$profile.managerChanged } else { $false }
    $data.inspectorChanged = if ($profile) { [bool]$profile.inspectorChanged } else { $false }
    $data.managerChangeCount = if ($profile) { [int]$profile.managerChangeCount } else { 0 }
    $data.inspectorChangeCount = if ($profile) { [int]$profile.inspectorChangeCount } else { 0 }
    $data.managerExonerationSignal = if ($profile) { [bool]$profile.managerExonerationSignal } else { $false }
    $data.inspectorExonerationSignal = if ($profile) { [bool]$profile.inspectorExonerationSignal } else { $false }
    $data.managerExonerationAt = if ($profile) { [string]$profile.managerExonerationAt } else { $null }
    $data.inspectorExonerationAt = if ($profile) { [string]$profile.inspectorExonerationAt } else { $null }
    $data.managementStatusTone = if ($profile) { [string]$profile.statusTone } else { 'critical' }
    $data.managementSummary = if ($profile) { [string]$profile.summary } else { 'Sem ato de gestor e fiscal identificado no Diário.' }
    $data.managementLastActAt = if ($profile) { [string]$profile.lastManagementActAt } else { $null }
    $data.managementLinkedMovementCount = if ($profile) { [int]$profile.linkedMovementCount } else { 0 }

    return [pscustomobject]$data
}

function Get-ManagementSummaryFromProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Profiles
    )

    $trackedContracts = @($Profiles).Count
    $withManager = @($Profiles | Where-Object { [bool]$_.hasManager }).Count
    $withInspector = @($Profiles | Where-Object { [bool]$_.hasInspector }).Count
    $managerChanged = @($Profiles | Where-Object { [bool]$_.managerChanged }).Count
    $inspectorChanged = @($Profiles | Where-Object { [bool]$_.inspectorChanged }).Count
    $exonerationSignals = @($Profiles | Where-Object { [bool]$_.managerExonerationSignal -or [bool]$_.inspectorExonerationSignal }).Count
    $atRisk = @($Profiles | Where-Object {
        -not [bool]$_.hasManager -or
        -not [bool]$_.hasInspector -or
        [bool]$_.managerExonerationSignal -or
        [bool]$_.inspectorExonerationSignal
    }).Count

    return [ordered]@{
        trackedContracts = [int]$trackedContracts
        withManager = [int]$withManager
        withInspector = [int]$withInspector
        withoutManager = [int]([Math]::Max($trackedContracts - $withManager, 0))
        withoutInspector = [int]([Math]::Max($trackedContracts - $withInspector, 0))
        managerChanged = [int]$managerChanged
        inspectorChanged = [int]$inspectorChanged
        exonerationSignals = [int]$exonerationSignals
        atRisk = [int]$atRisk
    }
}

function Get-RoleLabel {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Role
    )

    switch (([string]$Role).Trim().ToLowerInvariant()) {
        'admin' { return 'Administrador' }
        'auditor' { return 'Auditoria' }
        'reviewer' { return 'Revisao' }
        default { return 'Consulta' }
    }
}

function Get-RoleCapabilities {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Role
    )

    $normalizedRole = ([string]$Role).Trim().ToLowerInvariant()
    switch ($normalizedRole) {
        'admin' {
            return [ordered]@{
                canRefresh = $true
                canManageUsers = $true
                canManageSupport = $true
                canManageAlerts = $true
                canReviewContracts = $true
                canManageWorkflow = $true
                canCommentContracts = $true
                canSaveViews = $true
                canFavoriteContracts = $true
                canSeeAllSupport = $true
                canSeeActivityLog = $true
                canAcknowledgeNotifications = $true
                canArchiveNotifications = $true
                canViewGuide = $true
                canSeeObservability = $true
                canExportContracts = $true
                canManageFinancialAutomation = $true
            }
        }
        'auditor' {
            return [ordered]@{
                canRefresh = $false
                canManageUsers = $false
                canManageSupport = $true
                canManageAlerts = $true
                canReviewContracts = $true
                canManageWorkflow = $true
                canCommentContracts = $true
                canSaveViews = $true
                canFavoriteContracts = $true
                canSeeAllSupport = $true
                canSeeActivityLog = $true
                canAcknowledgeNotifications = $true
                canArchiveNotifications = $true
                canViewGuide = $true
                canSeeObservability = $true
                canExportContracts = $true
                canManageFinancialAutomation = $false
            }
        }
        'reviewer' {
            return [ordered]@{
                canRefresh = $false
                canManageUsers = $false
                canManageSupport = $false
                canManageAlerts = $false
                canReviewContracts = $true
                canManageWorkflow = $false
                canCommentContracts = $true
                canSaveViews = $true
                canFavoriteContracts = $true
                canSeeAllSupport = $false
                canSeeActivityLog = $false
                canAcknowledgeNotifications = $true
                canArchiveNotifications = $true
                canViewGuide = $true
                canSeeObservability = $true
                canExportContracts = $true
                canManageFinancialAutomation = $false
            }
        }
        default {
            return [ordered]@{
                canRefresh = $false
                canManageUsers = $false
                canManageSupport = $false
                canManageAlerts = $false
                canReviewContracts = $false
                canManageWorkflow = $false
                canCommentContracts = $true
                canSaveViews = $true
                canFavoriteContracts = $true
                canSeeAllSupport = $false
                canSeeActivityLog = $false
                canAcknowledgeNotifications = $true
                canArchiveNotifications = $true
                canViewGuide = $true
                canSeeObservability = $false
                canExportContracts = $true
                canManageFinancialAutomation = $false
            }
        }
    }
}

function Get-RoleCatalog {
    @(
        [pscustomobject][ordered]@{
            value = 'viewer'
            label = 'Consulta'
            description = 'Acesso de leitura ao painel, favoritos, busca global e visoes salvas.'
            capabilities = (Get-RoleCapabilities -Role 'viewer')
        }
        [pscustomobject][ordered]@{
            value = 'reviewer'
            label = 'Revisao'
            description = 'Leitura operacional com revisao de vinculos, comentarios internos e apoio ao dossie.'
            capabilities = (Get-RoleCapabilities -Role 'reviewer')
        }
        [pscustomobject][ordered]@{
            value = 'auditor'
            label = 'Auditoria'
            description = 'Acompanha fila geral, pendencias, suporte interno e workflow dos contratos.'
            capabilities = (Get-RoleCapabilities -Role 'auditor')
        }
        [pscustomobject][ordered]@{
            value = 'admin'
            label = 'Administrador'
            description = 'Controle total do painel, usuarios, sincronizacao e governanca operacional.'
            capabilities = (Get-RoleCapabilities -Role 'admin')
        }
    )
}

function Get-WorkspacePayload {
    $payload = Read-JsonFile -Path $script:WorkspaceStatePath -Default (Get-EmptyWorkspacePayload)
    if ($null -eq $payload) {
        return (Get-EmptyWorkspacePayload)
    }

    if ($payload -is [System.Collections.IDictionary]) {
        $convertedPayload = [ordered]@{}
        foreach ($entry in $payload.GetEnumerator()) {
            $convertedPayload[[string]$entry.Key] = $entry.Value
        }
        $payload = [pscustomobject]$convertedPayload
    }

    foreach ($propertyName in @('favorites', 'savedViews', 'contractNotes', 'workflowItems', 'alertStates', 'notificationStates', 'activityLog', 'aggregateSnapshots')) {
        if (-not (Test-ObjectProperty -Item $payload -Name $propertyName)) {
            $payload | Add-Member -NotePropertyName $propertyName -NotePropertyValue @()
        }
        elseif ($null -eq $payload.$propertyName) {
            $payload.$propertyName = @()
        }
        elseif ($payload.$propertyName -is [string] -or $payload.$propertyName -isnot [System.Collections.IEnumerable] -or $payload.$propertyName -is [System.Collections.IDictionary]) {
            $payload.$propertyName = @($payload.$propertyName)
        }
        else {
            $payload.$propertyName = @($payload.$propertyName)
        }
    }

    foreach ($workflowItem in @($payload.workflowItems)) {
        if ($null -eq $workflowItem) {
            continue
        }
        if (-not (Test-ObjectProperty -Item $workflowItem -Name 'history')) {
            $workflowItem | Add-Member -NotePropertyName history -NotePropertyValue @()
        }
    }

    foreach ($notificationState in @($payload.notificationStates)) {
        if ($null -eq $notificationState) {
            continue
        }
        foreach ($propertyName in @('userId', 'userLogin', 'status', 'updatedAt', 'updatedBy', 'history')) {
            if (-not (Test-ObjectProperty -Item $notificationState -Name $propertyName)) {
                $defaultValue = switch ($propertyName) {
                    'status' { 'novo' }
                    'history' { @() }
                    default { '' }
                }
                $notificationState | Add-Member -NotePropertyName $propertyName -NotePropertyValue $defaultValue
            }
        }
    }

    if (-not (Test-ObjectProperty -Item $payload -Name 'version')) {
        $payload | Add-Member -NotePropertyName version -NotePropertyValue $script:WorkspaceSchemaVersion
    }

    if (-not (Test-ObjectProperty -Item $payload -Name 'recentChanges') -or $null -eq $payload.recentChanges) {
        Set-WorkspacePayloadRecentChanges -Payload $payload
    }

    return $payload
}

function Save-WorkspacePayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $Payload.generatedAt = Get-IsoNow
    $Payload.version = $script:WorkspaceSchemaVersion
    $Payload.favorites = @(
        @($Payload.favorites) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true }, @{ Expression = { [string]$_.reference }; Descending = $false }
    )
    $Payload.savedViews = @(
        @($Payload.savedViews) |
        Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true }, @{ Expression = { [string]$_.name }; Descending = $false }
    )
    $Payload.contractNotes = @(
        @($Payload.contractNotes) |
        Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true }, @{ Expression = { [string]$_.reference }; Descending = $false }
    )
    $Payload.workflowItems = @(
        @($Payload.workflowItems) |
        Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true }, @{ Expression = { [string]$_.reference }; Descending = $false }
    )
    $Payload.alertStates = @(
        @($Payload.alertStates) |
        Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true }, @{ Expression = { [string]$_.alertKey }; Descending = $false }
    ) | Select-Object -First 500
    $Payload.notificationStates = @(
        @($Payload.notificationStates) |
        Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true }, @{ Expression = { [string]$_.notificationKey }; Descending = $false }
    ) | Select-Object -First 800
    $Payload.activityLog = @(
        @($Payload.activityLog) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true }
    ) | Select-Object -First 250
    $Payload.aggregateSnapshots = @(
        @($Payload.aggregateSnapshots) |
        Sort-Object @{ Expression = { [string]$_.generatedAt }; Descending = $true }
    ) | Select-Object -First 60
    Set-WorkspacePayloadRecentChanges -Payload $Payload -RecentChanges $Payload.recentChanges
    Write-JsonFile -Path $script:WorkspaceStatePath -Data $Payload
}

function Get-ObservabilityPayload {
    $payload = Read-JsonFile -Path $script:ObservabilityPath -Default (Get-EmptyObservabilityPayload)
    if ($null -eq $payload) {
        return (Get-EmptyObservabilityPayload)
    }

    if ($payload -is [System.Collections.IDictionary]) {
        $convertedPayload = [ordered]@{}
        foreach ($entry in $payload.GetEnumerator()) {
            $convertedPayload[[string]$entry.Key] = $entry.Value
        }
        $payload = [pscustomobject]$convertedPayload
    }

    if (-not $payload.PSObject.Properties['events']) {
        $payload | Add-Member -NotePropertyName events -NotePropertyValue @()
    }

    if (-not $payload.PSObject.Properties['counters']) {
        $payload | Add-Member -NotePropertyName counters -NotePropertyValue ([ordered]@{})
    }
    elseif ($payload.counters -isnot [System.Collections.IDictionary]) {
        $counterMap = [ordered]@{}
        foreach ($property in @($payload.counters.PSObject.Properties)) {
            $counterMap[[string]$property.Name] = $property.Value
        }
        $payload.counters = $counterMap
    }

    if (-not $payload.PSObject.Properties['version']) {
        $payload | Add-Member -NotePropertyName version -NotePropertyValue $script:ObservabilitySchemaVersion
    }

    return $payload
}

function Save-ObservabilityPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $Payload.generatedAt = Get-IsoNow
    $Payload.version = $script:ObservabilitySchemaVersion
    $Payload.events = @(
        @($Payload.events) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true }
    ) | Select-Object -First 400
    Write-JsonFile -Path $script:ObservabilityPath -Data $Payload
}

function Register-ObservabilityEvent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Status = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Message = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$UserLogin = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Metadata = $null
    )

    $payload = Get-ObservabilityPayload
    $eventType = Collapse-Whitespace -Text $Type
    $eventStatus = Collapse-Whitespace -Text $Status
    $counterKey = if ([string]::IsNullOrWhiteSpace($eventStatus)) { $eventType } else { "$eventType.$eventStatus" }
    $currentCounter = 0
    if ($payload.counters -isnot [System.Collections.IDictionary]) {
        $counterMap = [ordered]@{}
        foreach ($property in @($payload.counters.PSObject.Properties)) {
            $counterMap[[string]$property.Name] = $property.Value
        }
        $payload.counters = $counterMap
    }
    if ($payload.counters.Contains($counterKey)) {
        $currentCounter = [int]$payload.counters[$counterKey]
    }
    $payload.counters[$counterKey] = ($currentCounter + 1)
    $event = [pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        type = $eventType
        status = $eventStatus
        message = Collapse-Whitespace -Text $Message
        userLogin = Collapse-Whitespace -Text $UserLogin
        createdAt = Get-IsoNow
        metadata = $Metadata
    }
    $payload.events = @($payload.events) + @($event)
    Save-ObservabilityPayload -Payload $payload
    return $event
}

function Get-ObservabilityDashboardData {
    $payload = Get-ObservabilityPayload
    $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
    $financialMonitoring = $contractsPayload.financialMonitoring
    $crossReviewQueue = @($contractsPayload.crossReviewQueue)
    $suppressedSummary = $contractsPayload.crossSourceSuppressionSummary
    $events = @($payload.events | Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true })
    $now = Get-Date
    $searchWindow = $now.AddHours(-24)
    $operationalWindow = $now.AddDays(-7)
    $searches24h = @($events | Where-Object {
        [string]$_.type -eq 'search' -and [DateTime]::Parse([string]$_.createdAt) -ge $searchWindow
    }).Count
    $frontendLoads24h = @($events | Where-Object {
        [string]$_.type -eq 'frontend_perf' -and [DateTime]::Parse([string]$_.createdAt) -ge $searchWindow
    }).Count
    $alertActions7d = @($events | Where-Object {
        [string]$_.type -eq 'alert_action' -and [DateTime]::Parse([string]$_.createdAt) -ge $operationalWindow
    }).Count
    $workflowActions7d = @($events | Where-Object {
        [string]$_.type -eq 'workflow_action' -and [DateTime]::Parse([string]$_.createdAt) -ge $operationalWindow
    }).Count
    $notificationActions7d = @($events | Where-Object {
        [string]$_.type -eq 'notification_action' -and [DateTime]::Parse([string]$_.createdAt) -ge $operationalWindow
    }).Count
    $syncSuccess7d = @($events | Where-Object {
        [string]$_.type -eq 'sync' -and [string]$_.status -eq 'success' -and [DateTime]::Parse([string]$_.createdAt) -ge $operationalWindow
    }).Count
    $syncError7d = @($events | Where-Object {
        [string]$_.type -eq 'sync' -and [string]$_.status -eq 'error' -and [DateTime]::Parse([string]$_.createdAt) -ge $operationalWindow
    }).Count
    $frontendWarnings7d = @($events | Where-Object {
        [string]$_.type -eq 'frontend_contract' -and [string]$_.status -eq 'warning' -and [DateTime]::Parse([string]$_.createdAt) -ge $operationalWindow
    }).Count
    $latestSync = @($events | Where-Object { [string]$_.type -eq 'sync' } | Select-Object -First 1) | Select-Object -First 1
    $latestError = @($events | Where-Object { [string]$_.type -eq 'sync' -and [string]$_.status -eq 'error' } | Select-Object -First 1) | Select-Object -First 1

    return [ordered]@{
        metrics = @(
            [pscustomobject][ordered]@{
                key = 'searches_24h'
                label = 'Buscas nas ultimas 24h'
                value = [int]$searches24h
                meta = 'Uso recente da busca global no painel.'
            }
            [pscustomobject][ordered]@{
                key = 'frontend_loads_24h'
                label = 'Cargas validadas no frontend'
                value = [int]$frontendLoads24h
                meta = 'Leituras de workspace, dashboard e dossie registradas pelo cliente nas ultimas 24h.'
            }
            [pscustomobject][ordered]@{
                key = 'alert_actions_7d'
                label = 'Acoes em alertas'
                value = [int]$alertActions7d
                meta = 'Reconhecimentos, atribuicoes, adiamentos e reaberturas na ultima semana.'
            }
            [pscustomobject][ordered]@{
                key = 'workflow_actions_7d'
                label = 'Movimentacoes de workflow'
                value = [int]$workflowActions7d
                meta = 'Atualizacoes operacionais do workflow contratual na ultima semana.'
            }
            [pscustomobject][ordered]@{
                key = 'notification_actions_7d'
                label = 'Acoes em notificacoes'
                value = [int]$notificationActions7d
                meta = 'Leituras e arquivamentos de notificacoes pessoais na ultima semana.'
            }
            [pscustomobject][ordered]@{
                key = 'financial_query_ready'
                label = 'Busca financeira pronta'
                value = "$(if ($financialMonitoring) { [int]$financialMonitoring.queryReadyContracts } else { 0 })/$(if ($financialMonitoring) { [int]$financialMonitoring.monitoredContracts } else { 0 })"
                meta = if ($financialMonitoring) { "$([string]$financialMonitoring.modeLabel) | score medio $([int]$financialMonitoring.averageCoverageScore)." } else { 'Sem leitura financeira consolidada.' }
            }
            [pscustomobject][ordered]@{
                key = 'cross_review_backlog'
                label = 'Fila de revisao entre fontes'
                value = [int]@($crossReviewQueue).Count
                meta = "Ruido suprimido: $(if ($suppressedSummary) { [int]$suppressedSummary.total } else { 0 }) ocorrencia(s)."
            }
            [pscustomobject][ordered]@{
                key = 'frontend_contract_warnings_7d'
                label = 'Alertas de contrato do frontend'
                value = [int]$frontendWarnings7d
                meta = 'Avisos de payload inconsistente detectados pelo cliente na ultima semana.'
            }
            [pscustomobject][ordered]@{
                key = 'sync_health'
                label = 'Sincronizacoes na semana'
                value = "$syncSuccess7d/$syncError7d"
                meta = if ($latestSync) { "Ultimo sync em $([string]$latestSync.createdAt)." } else { 'Ainda nao ha sincronizacao registrada.' }
            }
        )
        latestSync = $latestSync
        latestError = $latestError
        recentEvents = @($events | Select-Object -First 12)
    }
}

function Add-WorkspaceActivity {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload,

        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Summary = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$CreatedBy = 'system',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Metadata = $null
    )

    $entry = [pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        type = Collapse-Whitespace -Text $Type
        title = Collapse-Whitespace -Text $Title
        summary = Collapse-Whitespace -Text $Summary
        reference = Collapse-Whitespace -Text $Reference
        createdAt = Get-IsoNow
        createdBy = Collapse-Whitespace -Text $CreatedBy
        metadata = $Metadata
    }

    $Payload.activityLog = @($Payload.activityLog) + @($entry)
    return $entry
}

function New-ReferenceTokenSet {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $referenceValue = Collapse-Whitespace -Text $Reference
    if ([string]::IsNullOrWhiteSpace($referenceValue)) {
        return $set
    }

    foreach ($token in @(Get-ContractReferenceTokens -ContractNumber $referenceValue -ProcessNumber $referenceValue)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
            $null = $set.Add([string]$token)
        }
    }

    $normalizedReference = Normalize-IndexText -Text $referenceValue
    if (-not [string]::IsNullOrWhiteSpace($normalizedReference)) {
        $null = $set.Add(($normalizedReference -replace '\s+', ''))
        $null = $set.Add($normalizedReference)
    }

    return $set
}

function Test-ReferenceMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reference,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$CandidateValues = @()
    )

    $referenceTokens = New-ReferenceTokenSet -Reference $Reference
    if ($referenceTokens.Count -eq 0) {
        return $false
    }

    foreach ($candidate in @($CandidateValues)) {
        foreach ($token in @(Get-ContractReferenceTokens -ContractNumber ([string]$candidate) -ProcessNumber ([string]$candidate))) {
            if (-not [string]::IsNullOrWhiteSpace([string]$token) -and $referenceTokens.Contains([string]$token)) {
                return $true
            }
        }

        $normalizedCandidate = Normalize-IndexText -Text ([string]$candidate)
        if (-not [string]::IsNullOrWhiteSpace($normalizedCandidate)) {
            if ($referenceTokens.Contains(($normalizedCandidate -replace '\s+', '')) -or $referenceTokens.Contains($normalizedCandidate)) {
                return $true
            }
        }
    }

    return $false
}

function Get-EmptyUsersPayload {
    [ordered]@{
        generatedAt = $null
        version = $script:UsersSchemaVersion
        users = @()
    }
}

function Get-EmptySupportPayload {
    [ordered]@{
        generatedAt = $null
        version = $script:SupportSchemaVersion
        tickets = @()
    }
}

function Get-PasswordPolicy {
    return [ordered]@{
        minLength = 6
        maxLength = 10
        requireLetter = $true
        requireDigit = $true
        forbidden = @('admin', '123456', '1234567', '12345678', '000000', '0000', 'senha', 'iguape')
    }
}

function Test-PasswordPolicy {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Password,

        [Parameter(Mandatory = $false)]
        [string]$Login = ''
    )

    $policy = Get-PasswordPolicy
    $reasons = [System.Collections.Generic.List[string]]::new()
    $passwordValue = [string]$Password
    $loginValue = ([string]$Login -replace '\D', '')

    if ([string]::IsNullOrWhiteSpace($passwordValue)) {
        $reasons.Add('Informe uma senha.')
    }
    else {
        if ($passwordValue.Length -lt [int]$policy.minLength -or $passwordValue.Length -gt [int]$policy.maxLength) {
            $reasons.Add("A senha deve ter entre $($policy.minLength) e $($policy.maxLength) caracteres.")
        }

        if ($passwordValue -match '\s') {
            $reasons.Add('A senha nao pode conter espacos.')
        }

        if ($policy.requireLetter -and $passwordValue -notmatch '[A-Za-z]') {
            $reasons.Add('A senha deve conter pelo menos uma letra.')
        }

        if ($policy.requireDigit -and $passwordValue -notmatch '\d') {
            $reasons.Add('A senha deve conter pelo menos um numero.')
        }

        if ($loginValue -and $passwordValue -eq $loginValue) {
            $reasons.Add('A senha nao pode ser igual ao login.')
        }

        if ($passwordValue.ToLowerInvariant() -in @($policy.forbidden)) {
            $reasons.Add('A senha informada e muito fraca. Escolha outra combinacao.')
        }
    }

    return [pscustomobject][ordered]@{
        valid = ($reasons.Count -eq 0)
        message = if ($reasons.Count -gt 0) { $reasons[0] } else { '' }
        reasons = @($reasons)
        policy = $policy
    }
}

function New-PasswordHashRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Password
    )

    $saltBytes = New-Object byte[] 16
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($saltBytes)
    }
    finally {
        $rng.Dispose()
    }

    $iterations = 120000
    $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $saltBytes, $iterations)
    try {
        $hashBytes = $derive.GetBytes(32)
    }
    finally {
        $derive.Dispose()
    }

    return [ordered]@{
        salt = [Convert]::ToBase64String($saltBytes)
        hash = [Convert]::ToBase64String($hashBytes)
        iterations = $iterations
    }
}

function Test-PasswordHashRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $true)]
        [string]$Salt,

        [Parameter(Mandatory = $true)]
        [string]$Hash,

        [Parameter(Mandatory = $true)]
        [int]$Iterations
    )

    $saltBytes = [Convert]::FromBase64String($Salt)
    $expectedHashBytes = [Convert]::FromBase64String($Hash)
    $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $saltBytes, $Iterations)
    try {
        $actualHashBytes = $derive.GetBytes($expectedHashBytes.Length)
    }
    finally {
        $derive.Dispose()
    }

    if ($actualHashBytes.Length -ne $expectedHashBytes.Length) {
        return $false
    }

    $difference = 0
    for ($index = 0; $index -lt $expectedHashBytes.Length; $index++) {
        $difference = $difference -bor ($actualHashBytes[$index] -bxor $expectedHashBytes[$index])
    }

    return ($difference -eq 0)
}

function New-UserRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Login,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $true)]
        [ValidateSet('admin', 'viewer')]
        [string]$Role,

        [Parameter(Mandatory = $false)]
        [string]$CreatedBy = 'bootstrap',

        [Parameter(Mandatory = $false)]
        [bool]$MustChangePassword = $true
    )

    $hashRecord = New-PasswordHashRecord -Password $Password
    return [ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        login = $Login
        name = $Name
        role = $Role
        isActive = $true
        passwordSalt = $hashRecord.salt
        passwordHash = $hashRecord.hash
        passwordIterations = [int]$hashRecord.iterations
        mustChangePassword = $MustChangePassword
        passwordChangedAt = $null
        createdAt = (Get-IsoNow)
        createdBy = $CreatedBy
        lastLoginAt = $null
    }
}

function Get-UsersPayload {
    $payload = Read-JsonFile -Path $script:UsersPath -Default (Get-EmptyUsersPayload)
    if ($null -eq $payload) {
        return (Get-EmptyUsersPayload)
    }

    if (-not $payload.PSObject.Properties.Match('users')) {
        $payload | Add-Member -NotePropertyName users -NotePropertyValue @()
    }

    if (-not $payload.PSObject.Properties.Match('version')) {
        $payload | Add-Member -NotePropertyName version -NotePropertyValue $script:UsersSchemaVersion
    }

    return $payload
}

function Ensure-UsersSecurityState {
    $payload = Get-UsersPayload
    $changed = $false

    foreach ($user in @($payload.users)) {
        $mustChangeProperty = $user.PSObject.Properties['mustChangePassword']
        if ($null -eq $mustChangeProperty) {
            $user | Add-Member -NotePropertyName mustChangePassword -NotePropertyValue $true
            $changed = $true
        }

        $passwordChangedProperty = $user.PSObject.Properties['passwordChangedAt']
        if ($null -eq $passwordChangedProperty) {
            $user | Add-Member -NotePropertyName passwordChangedAt -NotePropertyValue $null
            $changed = $true
        }

        $mustChangePassword = [bool]$user.PSObject.Properties['mustChangePassword'].Value
        $passwordChangedAt = [string]$user.PSObject.Properties['passwordChangedAt'].Value

        if ($passwordChangedAt -and -not $mustChangePassword) {
            continue
        }

        if (-not $passwordChangedAt -and -not $mustChangePassword) {
            $user.passwordChangedAt = [string]$user.createdAt
            $changed = $true
        }
    }

    if ($changed) {
        Save-UsersPayload -Payload $payload
    }
}

function Save-UsersPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $Payload.generatedAt = Get-IsoNow
    $Payload.version = $script:UsersSchemaVersion
    $Payload.users = @($Payload.users | Sort-Object login)
    Write-JsonFile -Path $script:UsersPath -Data $Payload
}

function Get-UserPublicProjection {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )

    return [ordered]@{
        id = [string]$User.id
        login = [string]$User.login
        name = [string]$User.name
        role = [string]$User.role
        roleLabel = (Get-RoleLabel -Role ([string]$User.role))
        capabilities = (Get-RoleCapabilities -Role ([string]$User.role))
        isActive = [bool]$User.isActive
        mustChangePassword = [bool]$User.mustChangePassword
        createdAt = [string]$User.createdAt
        createdBy = [string]$User.createdBy
        lastLoginAt = [string]$User.lastLoginAt
        passwordChangedAt = [string]$User.passwordChangedAt
    }
}

function Ensure-DefaultAdminUser {
    $payload = Get-UsersPayload
    $activeAdmins = @($payload.users | Where-Object { $_.role -eq 'admin' -and $_.isActive })
    if ($activeAdmins.Count -gt 0) {
        return
    }

    $payload.users = @($payload.users) + @([pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        login = '0001'
        name = 'Administrador Gerencial'
        role = 'admin'
        isActive = $true
        passwordSalt = 'l8XBlmWee8bn4B+P6AwdRw=='
        passwordHash = 'iXqEO1hgEjX84tn9nf+8ZeqAKmiyBH6yxavvTc4WGCI='
        passwordIterations = 120000
        mustChangePassword = $true
        passwordChangedAt = $null
        createdAt = (Get-IsoNow)
        createdBy = 'bootstrap'
        lastLoginAt = $null
    })

    Save-UsersPayload -Payload $payload
}

function Find-UserByLogin {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Login
    )

    $normalizedLogin = ($Login -replace '\D', '')
    return @((Get-UsersPayload).users | Where-Object { $_.login -eq $normalizedLogin -and $_.isActive }) | Select-Object -First 1
}

function Update-UserLastLogin {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    $payload = Get-UsersPayload
    $updated = $false
    foreach ($user in @($payload.users)) {
        if ([string]$user.id -eq $UserId) {
            $user.lastLoginAt = Get-IsoNow
            $updated = $true
            break
        }
    }

    if ($updated) {
        Save-UsersPayload -Payload $payload
    }
}

function Get-ManagedUsers {
    @((Get-UsersPayload).users | ForEach-Object { [pscustomobject](Get-UserPublicProjection -User $_) })
}

function Update-UserPassword {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $false)]
        [bool]$MustChangePassword = $false
    )

    $payload = Get-UsersPayload
    $targetUser = @($payload.users | Where-Object { [string]$_.id -eq $UserId -and $_.isActive }) | Select-Object -First 1
    if ($null -eq $targetUser) {
    throw 'Usuário não encontrado.'
    }

    $validation = Test-PasswordPolicy -Password $Password -Login ([string]$targetUser.login)
    if (-not $validation.valid) {
        throw $validation.message
    }

    $hashRecord = New-PasswordHashRecord -Password $Password
    $targetUser.passwordSalt = $hashRecord.salt
    $targetUser.passwordHash = $hashRecord.hash
    $targetUser.passwordIterations = [int]$hashRecord.iterations
    $targetUser.mustChangePassword = $MustChangePassword
    $targetUser.passwordChangedAt = if ($MustChangePassword) { $null } else { Get-IsoNow }

    Save-UsersPayload -Payload $payload
    return [pscustomobject](Get-UserPublicProjection -User $targetUser)
}

function Reset-ViewerPassword {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Login,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $false)]
        [bool]$MustChangePassword = $true
    )

    $user = Find-UserByLogin -Login $Login
    if ($null -eq $user) {
    throw 'Usuário não encontrado.'
    }

    if ([string]$user.role -ne 'viewer') {
        throw 'Somente usuarios de consulta podem ter a senha redefinida nesta tela.'
    }

    return Update-UserPassword -UserId ([string]$user.id) -Password $Password -MustChangePassword:$MustChangePassword
}

function Add-ViewerUser {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Login,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $false)]
        [string]$CreatedBy = '0001'
    )

    $normalizedLogin = ($Login -replace '\D', '')
    if ($normalizedLogin -notmatch '^\d{4}$') {
        throw 'O login deve conter exatamente 4 numeros.'
    }

    $cleanName = ($Name -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($cleanName)) {
        throw 'Informe o nome do usuario.'
    }

    $passwordValidation = Test-PasswordPolicy -Password $Password -Login $normalizedLogin
    if (-not $passwordValidation.valid) {
        throw $passwordValidation.message
    }

    $payload = Get-UsersPayload
    if (@($payload.users | Where-Object { $_.login -eq $normalizedLogin }).Count -gt 0) {
        throw 'Ja existe um usuario com esse login.'
    }

    $user = [pscustomobject](New-UserRecord -Login $normalizedLogin -Name $cleanName -Password $Password -Role 'viewer' -CreatedBy $CreatedBy -MustChangePassword $true)
    $payload.users = @($payload.users) + @($user)
    Save-UsersPayload -Payload $payload
    return [pscustomobject](Get-UserPublicProjection -User $user)
}

function Reset-ManagedUserPassword {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Login,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $false)]
        [bool]$MustChangePassword = $true
    )

    $user = Find-UserByLogin -Login $Login
    if ($null -eq $user) {
        throw 'Usuario nao encontrado.'
    }

    if ([string]$user.role -eq 'admin') {
        throw 'A senha do perfil administrador deve ser trocada pelo proprio usuario.'
    }

    return Update-UserPassword -UserId ([string]$user.id) -Password $Password -MustChangePassword:$MustChangePassword
}

function Add-ManagedUser {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Login,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $false)]
        [ValidateSet('viewer', 'reviewer', 'auditor')]
        [string]$Role = 'viewer',

        [Parameter(Mandatory = $false)]
        [string]$CreatedBy = '0001'
    )

    if ($Role -eq 'viewer') {
        return Add-ViewerUser -Login $Login -Name $Name -Password $Password -CreatedBy $CreatedBy
    }

    $normalizedLogin = ($Login -replace '\D', '')
    if ($normalizedLogin -notmatch '^\d{4}$') {
        throw 'O login deve conter exatamente 4 numeros.'
    }

    $cleanName = ($Name -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($cleanName)) {
        throw 'Informe o nome do usuario.'
    }

    $passwordValidation = Test-PasswordPolicy -Password $Password -Login $normalizedLogin
    if (-not $passwordValidation.valid) {
        throw $passwordValidation.message
    }

    $payload = Get-UsersPayload
    if (@($payload.users | Where-Object { $_.login -eq $normalizedLogin }).Count -gt 0) {
        throw 'Ja existe um usuario com esse login.'
    }

    $user = [pscustomobject](New-UserRecord -Login $normalizedLogin -Name $cleanName -Password $Password -Role $Role -CreatedBy $CreatedBy -MustChangePassword $true)
    $payload.users = @($payload.users) + @($user)
    Save-UsersPayload -Payload $payload
    return [pscustomobject](Get-UserPublicProjection -User $user)
}

function Normalize-SupportText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Value,

        [Parameter(Mandatory = $false)]
        [int]$MaxLength = 2000
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $normalized = ([string]$Value) -replace '\r\n?', "`n"
    $normalized = ($normalized -split "`n" | ForEach-Object { ($_ -replace '\s+', ' ').TrimEnd() }) -join "`n"
    $normalized = $normalized.Trim()

    if ($normalized.Length -gt $MaxLength) {
        $normalized = $normalized.Substring(0, $MaxLength).Trim()
    }

    return $normalized
}

function New-SupportHistoryEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Actor,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Summary = ''
    )

    return [pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        action = Collapse-Whitespace -Text $Action
        actor = Collapse-Whitespace -Text $Actor
        summary = Collapse-Whitespace -Text $Summary
        createdAt = Get-IsoNow
    }
}

function Get-SupportPayload {
    $payload = Read-JsonFile -Path $script:SupportPath -Default (Get-EmptySupportPayload)
    if ($null -eq $payload) {
        return (Get-EmptySupportPayload)
    }

    if (-not $payload.PSObject.Properties.Match('tickets')) {
        $payload | Add-Member -NotePropertyName tickets -NotePropertyValue @()
    }

    if (-not $payload.PSObject.Properties.Match('version')) {
        $payload | Add-Member -NotePropertyName version -NotePropertyValue $script:SupportSchemaVersion
    }

    $changed = $false
    foreach ($ticket in @($payload.tickets)) {
        foreach ($propertyName in @('assigneeUserId', 'assigneeLogin', 'assigneeName', 'dueDate', 'history')) {
            if (-not $ticket.PSObject.Properties.Match($propertyName)) {
                $ticket | Add-Member -NotePropertyName $propertyName -NotePropertyValue $(if ($propertyName -eq 'history') { @() } else { '' })
                $changed = $true
            }
        }

        if ($ticket.PSObject.Properties.Match('dueDate') -and [string]$ticket.dueDate -eq '') {
            $ticket.dueDate = $null
        }
    }

    if ($changed) {
        Save-SupportPayload -Payload $payload
    }

    return $payload
}

function Save-SupportPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $Payload.generatedAt = Get-IsoNow
    $Payload.version = $script:SupportSchemaVersion
    $Payload.tickets = @(
        $Payload.tickets | Sort-Object `
            @{ Expression = { [string]$_.updatedAt }; Descending = $true }, `
            @{ Expression = { [string]$_.createdAt }; Descending = $true }
    )
    Write-JsonFile -Path $script:SupportPath -Data $Payload
}

function New-SupportTicketRecord {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [ValidateSet('duvida', 'ajuste', 'autorizacao', 'cadastro', 'outro')]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [ValidateSet('baixa', 'normal', 'alta')]
        [string]$Priority
    )

    $now = Get-IsoNow
    return [ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        category = $Category
        subject = $Subject
        message = $Message
        priority = $Priority
        status = 'aberto'
        requesterUserId = [string]$User.id
        requesterLogin = [string]$User.login
        requesterName = [string]$User.name
        adminResponse = ''
        internalNote = ''
        assigneeUserId = ''
        assigneeLogin = ''
        assigneeName = ''
        dueDate = $null
        createdAt = $now
        updatedAt = $now
        lastUpdatedBy = [string]$User.login
        history = @(
            New-SupportHistoryEntry -Action 'created' -Actor ([string]$User.login) -Summary 'Solicitacao registrada no painel.'
        )
    }
}

function Get-SupportTicketProjection {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Ticket,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeInternal
    )

    $projection = [ordered]@{
        id = [string]$Ticket.id
        category = [string]$Ticket.category
        subject = [string]$Ticket.subject
        message = [string]$Ticket.message
        priority = [string]$Ticket.priority
        status = [string]$Ticket.status
        requester = [ordered]@{
            id = [string]$Ticket.requesterUserId
            login = [string]$Ticket.requesterLogin
            name = [string]$Ticket.requesterName
        }
        adminResponse = [string]$Ticket.adminResponse
        assignee = if (-not [string]::IsNullOrWhiteSpace([string]$Ticket.assigneeLogin)) {
            [ordered]@{
                id = [string]$Ticket.assigneeUserId
                login = [string]$Ticket.assigneeLogin
                name = [string]$Ticket.assigneeName
            }
        }
        else {
            $null
        }
        dueDate = if ([string]::IsNullOrWhiteSpace([string]$Ticket.dueDate)) { $null } else { [string]$Ticket.dueDate }
        createdAt = [string]$Ticket.createdAt
        updatedAt = [string]$Ticket.updatedAt
        lastUpdatedBy = [string]$Ticket.lastUpdatedBy
        history = @($Ticket.history)
    }

    if ($IncludeInternal) {
        $projection['internalNote'] = [string]$Ticket.internalNote
    }

    return $projection
}

function Get-SupportTickets {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeAll
    )

    $payload = Get-SupportPayload
    $tickets = if ($IncludeAll) {
        @($payload.tickets)
    }
    else {
        @($payload.tickets | Where-Object { [string]$_.requesterUserId -eq [string]$User.id })
    }

    $sortedTickets = @(
        $tickets | Sort-Object `
            @{ Expression = { [string]$_.updatedAt }; Descending = $true }, `
            @{ Expression = { [string]$_.createdAt }; Descending = $true }
    )
    return @($sortedTickets | ForEach-Object {
        [pscustomobject](Get-SupportTicketProjection -Ticket $_ -IncludeInternal:$IncludeAll)
    })
}

function Add-SupportTicket {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [string]$Priority
    )

    if ($Category -notin @('duvida', 'ajuste', 'autorizacao', 'cadastro', 'outro')) {
        throw 'Categoria de suporte invalida.'
    }

    if ($Priority -notin @('baixa', 'normal', 'alta')) {
        throw 'Prioridade de suporte invalida.'
    }

    $cleanSubject = Normalize-SupportText -Value $Subject -MaxLength 120
    if ([string]::IsNullOrWhiteSpace($cleanSubject) -or $cleanSubject.Length -lt 4) {
        throw 'Informe um assunto com pelo menos 4 caracteres.'
    }

    $cleanMessage = Normalize-SupportText -Value $Message -MaxLength 2000
    if ([string]::IsNullOrWhiteSpace($cleanMessage) -or $cleanMessage.Length -lt 12) {
        throw 'Descreva a solicitacao com pelo menos 12 caracteres.'
    }

    $payload = Get-SupportPayload
    $ticket = [pscustomobject](New-SupportTicketRecord -User $User -Category $Category -Subject $cleanSubject -Message $cleanMessage -Priority $Priority)
    $payload.tickets = @($payload.tickets) + @($ticket)
    Save-SupportPayload -Payload $payload
    return [pscustomobject](Get-SupportTicketProjection -Ticket $ticket)
}

function Update-SupportTicket {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TicketId,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [string]$Priority,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AdminResponse = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$InternalNote = '',

        [Parameter(Mandatory = $true)]
        [string]$UpdatedBy
    )

    if ($Status -notin @('aberto', 'em_analise', 'aguardando_retorno', 'autorizado', 'concluido')) {
        throw 'Status de suporte invalido.'
    }

    if ($Priority -notin @('baixa', 'normal', 'alta')) {
        throw 'Prioridade de suporte invalida.'
    }

    $payload = Get-SupportPayload
    $ticket = @($payload.tickets | Where-Object { [string]$_.id -eq $TicketId }) | Select-Object -First 1
    if ($null -eq $ticket) {
    throw 'Solicitação de suporte não encontrada.'
    }

    $ticket.status = $Status
    $ticket.priority = $Priority
    $ticket.adminResponse = Normalize-SupportText -Value $AdminResponse -MaxLength 2000
    $ticket.internalNote = Normalize-SupportText -Value $InternalNote -MaxLength 2000
    $ticket.updatedAt = Get-IsoNow
    $ticket.lastUpdatedBy = $UpdatedBy

    Save-SupportPayload -Payload $payload
    return [pscustomobject](Get-SupportTicketProjection -Ticket $ticket -IncludeInternal)
}

function Update-SupportTicketRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TicketId,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [string]$Priority,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AdminResponse = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$InternalNote = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AssigneeLogin = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$DueDate = '',

        [Parameter(Mandatory = $true)]
        [string]$UpdatedBy
    )

    if ($Status -notin @('aberto', 'em_analise', 'aguardando_retorno', 'autorizado', 'concluido')) {
        throw 'Status de suporte invalido.'
    }

    if ($Priority -notin @('baixa', 'normal', 'alta')) {
        throw 'Prioridade de suporte invalida.'
    }

    $payload = Get-SupportPayload
    $ticket = @($payload.tickets | Where-Object { [string]$_.id -eq $TicketId }) | Select-Object -First 1
    if ($null -eq $ticket) {
        throw 'Solicitacao de suporte nao encontrada.'
    }

    $cleanAssigneeLogin = ([string]$AssigneeLogin -replace '\D', '')
    $assignee = $null
    if (-not [string]::IsNullOrWhiteSpace($cleanAssigneeLogin)) {
        $assignee = Find-UserByLogin -Login $cleanAssigneeLogin
        if ($null -eq $assignee) {
            throw 'Usuario responsavel nao encontrado.'
        }
    }

    $normalizedDueDate = if ([string]::IsNullOrWhiteSpace([string]$DueDate)) {
        $null
    }
    else {
        $parsedDueDate = [DateTime]::MinValue
        if (-not [DateTime]::TryParse([string]$DueDate, [ref]$parsedDueDate)) {
            throw 'Prazo informado para o suporte e invalido.'
        }
        $parsedDueDate.ToString('yyyy-MM-dd')
    }

    $cleanAdminResponse = Normalize-SupportText -Value $AdminResponse -MaxLength 2000
    $cleanInternalNote = Normalize-SupportText -Value $InternalNote -MaxLength 2000
    $history = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($ticket.history)) {
        $history.Add($entry)
    }

    if ([string]$ticket.status -ne [string]$Status) {
        $history.Add((New-SupportHistoryEntry -Action 'status_changed' -Actor $UpdatedBy -Summary "Status alterado para $Status."))
    }
    if ([string]$ticket.priority -ne [string]$Priority) {
        $history.Add((New-SupportHistoryEntry -Action 'priority_changed' -Actor $UpdatedBy -Summary "Prioridade alterada para $Priority."))
    }
    if (([string]$ticket.assigneeLogin) -ne $(if ($assignee) { [string]$assignee.login } else { '' })) {
        $history.Add((New-SupportHistoryEntry -Action 'assignment_changed' -Actor $UpdatedBy -Summary $(if ($assignee) { "Responsavel definido para $([string]$assignee.name)." } else { 'Responsavel removido.' })))
    }
    if (([string]$ticket.dueDate) -ne $(if ($normalizedDueDate) { [string]$normalizedDueDate } else { '' })) {
        $history.Add((New-SupportHistoryEntry -Action 'due_date_changed' -Actor $UpdatedBy -Summary $(if ($normalizedDueDate) { "Prazo ajustado para $normalizedDueDate." } else { 'Prazo removido.' })))
    }
    if (([string]$ticket.adminResponse) -ne $cleanAdminResponse) {
        $history.Add((New-SupportHistoryEntry -Action 'response_updated' -Actor $UpdatedBy -Summary 'Retorno administrativo atualizado.'))
    }
    if (([string]$ticket.internalNote) -ne $cleanInternalNote) {
        $history.Add((New-SupportHistoryEntry -Action 'internal_note_updated' -Actor $UpdatedBy -Summary 'Anotacao interna atualizada.'))
    }

    $ticket.status = $Status
    $ticket.priority = $Priority
    $ticket.adminResponse = $cleanAdminResponse
    $ticket.internalNote = $cleanInternalNote
    $ticket.assigneeUserId = if ($assignee) { [string]$assignee.id } else { '' }
    $ticket.assigneeLogin = if ($assignee) { [string]$assignee.login } else { '' }
    $ticket.assigneeName = if ($assignee) { [string]$assignee.name } else { '' }
    $ticket.dueDate = $normalizedDueDate
    $ticket.updatedAt = Get-IsoNow
    $ticket.lastUpdatedBy = $UpdatedBy
    $ticket.history = @(
        @($history) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true }
    )

    Save-SupportPayload -Payload $payload
    Register-ObservabilityEvent -Type 'support_action' -Status $Status -Message "Suporte $TicketId atualizado." -UserLogin $UpdatedBy -Metadata ([ordered]@{
        ticketId = [string]$ticket.id
        priority = [string]$Priority
        assigneeLogin = [string]$ticket.assigneeLogin
        dueDate = $ticket.dueDate
    }) | Out-Null
    return [pscustomobject](Get-SupportTicketProjection -Ticket $ticket -IncludeInternal)
}

function Get-WorkspaceFavoriteEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    @((Get-WorkspacePayload).favorites | Where-Object { [string]$_.userId -eq $UserId })
}

function Set-WorkspaceFavorite {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Reference,

        [Parameter(Mandatory = $false)]
        [bool]$IsFavorite = $true
    )

    $cleanReference = Collapse-Whitespace -Text $Reference
    if ([string]::IsNullOrWhiteSpace($cleanReference)) {
        throw 'Referencia do contrato nao informada para favorito.'
    }

    $payload = Get-WorkspacePayload
    $payload.favorites = @($payload.favorites | Where-Object {
        -not (
            [string]$_.userId -eq [string]$User.id -and
            [string]$_.reference -eq $cleanReference
        )
    })

    if ($IsFavorite) {
        $payload.favorites = @($payload.favorites) + @([pscustomobject][ordered]@{
            id = [Guid]::NewGuid().ToString('N')
            userId = [string]$User.id
            userLogin = [string]$User.login
            reference = $cleanReference
            createdAt = Get-IsoNow
        })
        Add-WorkspaceActivity -Payload $payload -Type 'favorite' -Title 'Contrato favoritado' -Summary "Favorito salvo para $cleanReference." -Reference $cleanReference -CreatedBy ([string]$User.login) | Out-Null
    }
    else {
        Add-WorkspaceActivity -Payload $payload -Type 'favorite_removed' -Title 'Favorito removido' -Summary "Favorito removido de $cleanReference." -Reference $cleanReference -CreatedBy ([string]$User.login) | Out-Null
    }

    Save-WorkspacePayload -Payload $payload
    return [ordered]@{
        reference = $cleanReference
        isFavorite = $IsFavorite
    }
}

function Get-WorkspaceSavedViews {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    @((Get-WorkspacePayload).savedViews | Where-Object { [string]$_.userId -eq $UserId })
}

function Save-WorkspaceView {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Page,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [object]$Definition,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ViewId = ''
    )

    $cleanPage = Collapse-Whitespace -Text $Page
    $cleanName = Collapse-Whitespace -Text $Name
    if ([string]::IsNullOrWhiteSpace($cleanPage) -or [string]::IsNullOrWhiteSpace($cleanName)) {
        throw 'Nome da visao ou pagina nao informados.'
    }

    $payload = Get-WorkspacePayload
    $existing = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$ViewId)) {
        $existing = @($payload.savedViews | Where-Object { [string]$_.id -eq [string]$ViewId -and [string]$_.userId -eq [string]$User.id }) | Select-Object -First 1
    }

    if ($null -eq $existing) {
        $payload.savedViews = @($payload.savedViews) + @([pscustomobject][ordered]@{
            id = [Guid]::NewGuid().ToString('N')
            userId = [string]$User.id
            userLogin = [string]$User.login
            page = $cleanPage
            name = $cleanName
            definition = $Definition
            createdAt = Get-IsoNow
            updatedAt = Get-IsoNow
            updatedBy = [string]$User.login
        })
    }
    else {
        $existing.page = $cleanPage
        $existing.name = $cleanName
        $existing.definition = $Definition
        $existing.updatedAt = Get-IsoNow
        $existing.updatedBy = [string]$User.login
    }

    Add-WorkspaceActivity -Payload $payload -Type 'saved_view' -Title 'Visao salva' -Summary "$cleanName atualizada para a pagina $cleanPage." -CreatedBy ([string]$User.login) | Out-Null
    Save-WorkspacePayload -Payload $payload
    return @((Get-WorkspacePayload).savedViews | Where-Object { [string]$_.userId -eq [string]$User.id -and [string]$_.page -eq $cleanPage -and [string]$_.name -eq $cleanName } | Select-Object -First 1)[0]
}

function Remove-WorkspaceView {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$ViewId
    )

    $payload = Get-WorkspacePayload
    $view = @($payload.savedViews | Where-Object { [string]$_.id -eq [string]$ViewId -and [string]$_.userId -eq [string]$User.id }) | Select-Object -First 1
    if ($null -eq $view) {
        throw 'Visao salva nao encontrada.'
    }

    $payload.savedViews = @($payload.savedViews | Where-Object { [string]$_.id -ne [string]$ViewId })
    Add-WorkspaceActivity -Payload $payload -Type 'saved_view_removed' -Title 'Visao removida' -Summary "Visao $([string]$view.name) removida." -CreatedBy ([string]$User.login) | Out-Null
    Save-WorkspacePayload -Payload $payload
    return [ordered]@{
        removed = $true
        viewId = [string]$ViewId
    }
}

function Get-WorkspaceContractNotes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reference
    )

    @((Get-WorkspacePayload).contractNotes | Where-Object {
        Test-ReferenceMatch -Reference $Reference -CandidateValues @([string]$_.reference)
    })
}

function Add-WorkspaceContractNote {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Reference,

        [Parameter(Mandatory = $true)]
        [string]$Body
    )

    $cleanReference = Collapse-Whitespace -Text $Reference
    $cleanBody = Normalize-SupportText -Value $Body -MaxLength 2000
    if ([string]::IsNullOrWhiteSpace($cleanReference) -or [string]::IsNullOrWhiteSpace($cleanBody)) {
        throw 'Referencia e comentario do contrato sao obrigatorios.'
    }

    $payload = Get-WorkspacePayload
    $note = [pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        reference = $cleanReference
        body = $cleanBody
        createdAt = Get-IsoNow
        createdBy = [string]$User.login
        createdByName = [string]$User.name
        updatedAt = Get-IsoNow
        updatedBy = [string]$User.login
    }
    $payload.contractNotes = @($payload.contractNotes) + @($note)
    Add-WorkspaceActivity -Payload $payload -Type 'contract_note' -Title 'Comentario interno' -Summary "Comentario registrado para $cleanReference." -Reference $cleanReference -CreatedBy ([string]$User.login) | Out-Null
    Save-WorkspacePayload -Payload $payload
    Register-ObservabilityEvent -Type 'contract_note' -Status 'created' -Message "Comentario registrado para $cleanReference." -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        reference = $cleanReference
    }) | Out-Null
    return $note
}

function Get-WorkspaceContractWorkflowItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reference
    )

    @((Get-WorkspacePayload).workflowItems | Where-Object {
        Test-ReferenceMatch -Reference $Reference -CandidateValues @([string]$_.reference)
    } | Select-Object -First 1) | Select-Object -First 1
}

function Set-WorkspaceContractWorkflow {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Reference,

        [Parameter(Mandatory = $true)]
        [ValidateSet('novo', 'em_analise', 'aguardando_documento', 'regularizado', 'encerrado')]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AssigneeLogin = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$DueDate = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Note = ''
    )

    $cleanReference = Collapse-Whitespace -Text $Reference
    if ([string]::IsNullOrWhiteSpace($cleanReference)) {
        throw 'Referencia do workflow contratual nao informada.'
    }

    $assignee = $null
    $cleanAssigneeLogin = ([string]$AssigneeLogin -replace '\D', '')
    if (-not [string]::IsNullOrWhiteSpace($cleanAssigneeLogin)) {
        $assignee = Find-UserByLogin -Login $cleanAssigneeLogin
        if ($null -eq $assignee) {
            throw 'Usuario responsavel nao encontrado.'
        }
    }

    $normalizedDueDate = if ([string]::IsNullOrWhiteSpace([string]$DueDate)) {
        $null
    }
    else {
        $parsedDueDate = [DateTime]::MinValue
        if (-not [DateTime]::TryParse([string]$DueDate, [ref]$parsedDueDate)) {
            throw 'Prazo do workflow contratual invalido.'
        }
        $parsedDueDate.ToString('yyyy-MM-dd')
    }

    $payload = Get-WorkspacePayload
    $workflowItem = @($payload.workflowItems | Where-Object { Test-ReferenceMatch -Reference $cleanReference -CandidateValues @([string]$_.reference) } | Select-Object -First 1) | Select-Object -First 1
    if ($null -eq $workflowItem) {
        $workflowItem = [pscustomobject][ordered]@{
            id = [Guid]::NewGuid().ToString('N')
            reference = $cleanReference
            status = 'novo'
            assigneeUserId = ''
            assigneeLogin = ''
            assigneeName = ''
            dueDate = $null
            note = ''
            createdAt = Get-IsoNow
            createdBy = [string]$User.login
            updatedAt = ''
            updatedBy = ''
            history = @()
        }
        $payload.workflowItems = @($payload.workflowItems) + @($workflowItem)
    }

    foreach ($propertyName in @('status', 'assigneeUserId', 'assigneeLogin', 'assigneeName', 'dueDate', 'note', 'createdAt', 'createdBy', 'updatedAt', 'updatedBy', 'history')) {
        if (-not (Test-ObjectProperty -Item $workflowItem -Name $propertyName)) {
            $defaultValue = switch ($propertyName) {
                'status' { 'novo' }
                'dueDate' { $null }
                'history' { @() }
                default { '' }
            }
            $workflowItem | Add-Member -NotePropertyName $propertyName -NotePropertyValue $defaultValue
        }
    }

    $workflowItem.status = $Status
    $workflowItem.assigneeUserId = if ($assignee) { [string]$assignee.id } else { '' }
    $workflowItem.assigneeLogin = if ($assignee) { [string]$assignee.login } else { '' }
    $workflowItem.assigneeName = if ($assignee) { [string]$assignee.name } else { '' }
    $workflowItem.dueDate = $normalizedDueDate
    $workflowItem.note = Normalize-SupportText -Value $Note -MaxLength 2000
    $workflowItem.updatedAt = Get-IsoNow
    $workflowItem.updatedBy = [string]$User.login
    $workflowHistory = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($workflowItem.history)) {
        $workflowHistory.Add($entry)
    }
    $workflowHistory.Add([pscustomobject][ordered]@{
        action = 'workflow_update'
        actor = [string]$User.login
        status = $Status
        assigneeLogin = if ($assignee) { [string]$assignee.login } else { '' }
        assigneeName = if ($assignee) { [string]$assignee.name } else { '' }
        dueDate = $normalizedDueDate
        note = $workflowItem.note
        createdAt = Get-IsoNow
    })
    $workflowItem.history = @(
        @($workflowHistory.ToArray()) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true } |
        Select-Object -First 30
    )

    Add-WorkspaceActivity -Payload $payload -Type 'contract_workflow' -Title 'Workflow contratual' -Summary "Workflow de $cleanReference atualizado para $Status." -Reference $cleanReference -CreatedBy ([string]$User.login) | Out-Null
    Save-WorkspacePayload -Payload $payload
    Register-ObservabilityEvent -Type 'workflow_action' -Status $Status -Message "Workflow atualizado para $cleanReference." -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        reference = $cleanReference
        assigneeLogin = if ($assignee) { [string]$assignee.login } else { '' }
        dueDate = $normalizedDueDate
    }) | Out-Null
    return $workflowItem
}

function Get-WorkspaceAlertKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Title = ''
    )

    $parts = @(
        (Normalize-IndexText -Text $Type)
        (Normalize-IndexText -Text $Reference)
        (Normalize-IndexText -Text $Title)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    return ([string]::Join('|', $parts) -replace '\s+', '')
}

function Get-WorkspaceAlertStateItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AlertKey
    )

    @((Get-WorkspacePayload).alertStates | Where-Object { [string]$_.alertKey -eq [string]$AlertKey } | Select-Object -First 1) | Select-Object -First 1
}

function Set-WorkspaceAlertState {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$AlertKey,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AlertType = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Title = '',

        [Parameter(Mandatory = $true)]
        [ValidateSet('acknowledge', 'assign', 'snooze', 'resolve', 'reopen')]
        [string]$Action,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AssigneeLogin = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$DueDate = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$SnoozeUntil = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Justification = ''
    )

    $cleanAlertKey = Collapse-Whitespace -Text $AlertKey
    if ([string]::IsNullOrWhiteSpace($cleanAlertKey)) {
        throw 'Chave do alerta nao informada.'
    }

    $assignee = $null
    $cleanAssigneeLogin = ([string]$AssigneeLogin -replace '\D', '')
    if (-not [string]::IsNullOrWhiteSpace($cleanAssigneeLogin)) {
        $assignee = Find-UserByLogin -Login $cleanAssigneeLogin
        if ($null -eq $assignee) {
            throw 'Usuario responsavel do alerta nao encontrado.'
        }
    }

    $normalizeDateValue = {
        param([string]$Value, [string]$Label)
        if ([string]::IsNullOrWhiteSpace([string]$Value)) {
            return $null
        }
        $parsedDate = [DateTime]::MinValue
        if (-not [DateTime]::TryParse([string]$Value, [ref]$parsedDate)) {
            throw "$Label invalido."
        }
        return $parsedDate.ToString('yyyy-MM-dd')
    }

    $normalizedDueDate = & $normalizeDateValue ([string]$DueDate) 'Prazo do alerta'
    $normalizedSnoozeUntil = & $normalizeDateValue ([string]$SnoozeUntil) 'Data de adiamento'
    $cleanJustification = Normalize-SupportText -Value $Justification -MaxLength 1200
    if ($Action -eq 'snooze' -and -not $normalizedSnoozeUntil) {
        throw 'Informe ate quando o alerta deve ficar adiado.'
    }
    if (($Action -eq 'snooze' -or $Action -eq 'resolve') -and [string]::IsNullOrWhiteSpace($cleanJustification)) {
        throw 'Informe uma justificativa para registrar esta acao no alerta.'
    }

    $payload = Get-WorkspacePayload
    $alertState = @($payload.alertStates | Where-Object { [string]$_.alertKey -eq [string]$cleanAlertKey } | Select-Object -First 1) | Select-Object -First 1
    if ($null -eq $alertState) {
        $alertState = [pscustomobject][ordered]@{
            id = [Guid]::NewGuid().ToString('N')
            alertKey = $cleanAlertKey
            alertType = Collapse-Whitespace -Text $AlertType
            reference = Collapse-Whitespace -Text $Reference
            title = Collapse-Whitespace -Text $Title
            status = 'novo'
            assigneeUserId = ''
            assigneeLogin = ''
            assigneeName = ''
            dueDate = $null
            snoozeUntil = $null
            justification = ''
            createdAt = Get-IsoNow
            createdBy = [string]$User.login
            updatedAt = ''
            updatedBy = ''
            history = @()
        }
        $payload.alertStates = @($payload.alertStates) + @($alertState)
    }

    foreach ($propertyName in @('alertType', 'reference', 'title', 'status', 'assigneeUserId', 'assigneeLogin', 'assigneeName', 'dueDate', 'snoozeUntil', 'justification', 'createdAt', 'createdBy', 'updatedAt', 'updatedBy', 'history')) {
        if (-not (Test-ObjectProperty -Item $alertState -Name $propertyName)) {
            $defaultValue = switch ($propertyName) {
                'dueDate' { $null }
                'snoozeUntil' { $null }
                'history' { @() }
                'status' { 'novo' }
                default { '' }
            }
            $alertState | Add-Member -NotePropertyName $propertyName -NotePropertyValue $defaultValue
        }
    }

    $history = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($alertState.history)) {
        $history.Add($entry)
    }

    switch ($Action) {
        'acknowledge' {
            $alertState.status = 'reconhecido'
            if (-not [string]::IsNullOrWhiteSpace($cleanJustification)) {
                $alertState.justification = $cleanJustification
            }
        }
        'assign' {
            $alertState.status = if ($assignee) { 'atribuido' } else { 'reconhecido' }
            $alertState.assigneeUserId = if ($assignee) { [string]$assignee.id } else { '' }
            $alertState.assigneeLogin = if ($assignee) { [string]$assignee.login } else { '' }
            $alertState.assigneeName = if ($assignee) { [string]$assignee.name } else { '' }
            $alertState.dueDate = $normalizedDueDate
            if (-not [string]::IsNullOrWhiteSpace($cleanJustification)) {
                $alertState.justification = $cleanJustification
            }
        }
        'snooze' {
            $alertState.status = 'adiado'
            $alertState.snoozeUntil = $normalizedSnoozeUntil
            $alertState.justification = $cleanJustification
        }
        'resolve' {
            $alertState.status = 'resolvido'
            $alertState.justification = $cleanJustification
        }
        'reopen' {
            $alertState.status = 'novo'
            $alertState.snoozeUntil = $null
            if (-not [string]::IsNullOrWhiteSpace($cleanJustification)) {
                $alertState.justification = $cleanJustification
            }
        }
    }

    $alertState.alertType = if ([string]::IsNullOrWhiteSpace([string]$alertState.alertType)) { Collapse-Whitespace -Text $AlertType } else { [string]$alertState.alertType }
    $alertState.reference = if ([string]::IsNullOrWhiteSpace([string]$alertState.reference)) { Collapse-Whitespace -Text $Reference } else { [string]$alertState.reference }
    $alertState.title = if ([string]::IsNullOrWhiteSpace([string]$alertState.title)) { Collapse-Whitespace -Text $Title } else { [string]$alertState.title }
    $alertState.updatedAt = Get-IsoNow
    $alertState.updatedBy = [string]$User.login

    $actionLabel = switch ($Action) {
        'acknowledge' { 'Alerta reconhecido.' }
        'assign' { if ($assignee) { "Alerta atribuido a $([string]$assignee.name)." } else { 'Alerta mantido sem responsavel.' } }
        'snooze' { "Alerta adiado ate $normalizedSnoozeUntil." }
        'resolve' { 'Alerta marcado como resolvido.' }
        'reopen' { 'Alerta reaberto para tratamento.' }
    }
    $history.Add([pscustomobject][ordered]@{
        action = $Action
        actor = [string]$User.login
        summary = [string]$actionLabel
        createdAt = Get-IsoNow
        assigneeLogin = if ($assignee) { [string]$assignee.login } else { [string]$alertState.assigneeLogin }
        dueDate = $alertState.dueDate
        snoozeUntil = $alertState.snoozeUntil
    })
    $alertState.history = @(
        @($history.ToArray()) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true } |
        Select-Object -First 20
    )

    Add-WorkspaceActivity -Payload $payload -Type 'alert_action' -Title 'Alerta operacional atualizado' -Summary "$cleanAlertKey tratado com a acao $Action." -Reference ([string]$alertState.reference) -CreatedBy ([string]$User.login) | Out-Null
    Save-WorkspacePayload -Payload $payload
    Register-ObservabilityEvent -Type 'alert_action' -Status $Action -Message $actionLabel -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        alertKey = $cleanAlertKey
        reference = [string]$alertState.reference
        assigneeLogin = [string]$alertState.assigneeLogin
        dueDate = $alertState.dueDate
        snoozeUntil = $alertState.snoozeUntil
    }) | Out-Null
    return $alertState
}

function Get-WorkspaceNotificationKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Title = ''
    )

    $parts = @(
        (Normalize-IndexText -Text $Type)
        (Normalize-IndexText -Text $Reference)
        (Normalize-IndexText -Text $Title)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    return ([string]::Join('|', $parts) -replace '\s+', '')
}

function Get-WorkspaceNotificationStateItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$NotificationKey
    )

    @((Get-WorkspacePayload).notificationStates | Where-Object {
        [string]$_.notificationKey -eq [string]$NotificationKey -and
        [string]$_.userId -eq [string]$User.id
    } | Select-Object -First 1) | Select-Object -First 1
}

function Set-WorkspaceNotificationState {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$NotificationKey,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$NotificationType = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Title = '',

        [Parameter(Mandatory = $true)]
        [ValidateSet('read', 'archive', 'reopen')]
        [string]$Action
    )

    $cleanNotificationKey = Collapse-Whitespace -Text $NotificationKey
    if ([string]::IsNullOrWhiteSpace($cleanNotificationKey)) {
        throw 'Chave da notificacao nao informada.'
    }

    $payload = Get-WorkspacePayload
    $notificationState = @($payload.notificationStates | Where-Object {
        [string]$_.notificationKey -eq [string]$cleanNotificationKey -and
        [string]$_.userId -eq [string]$User.id
    } | Select-Object -First 1) | Select-Object -First 1

    if ($null -eq $notificationState) {
        $notificationState = [pscustomobject][ordered]@{
            id = [Guid]::NewGuid().ToString('N')
            notificationKey = $cleanNotificationKey
            notificationType = Collapse-Whitespace -Text $NotificationType
            reference = Collapse-Whitespace -Text $Reference
            title = Collapse-Whitespace -Text $Title
            userId = [string]$User.id
            userLogin = [string]$User.login
            status = 'novo'
            createdAt = Get-IsoNow
            createdBy = [string]$User.login
            updatedAt = ''
            updatedBy = ''
            history = @()
        }
        $payload.notificationStates = @($payload.notificationStates) + @($notificationState)
    }

    foreach ($propertyName in @('notificationType', 'reference', 'title', 'userId', 'userLogin', 'status', 'createdAt', 'createdBy', 'updatedAt', 'updatedBy', 'history')) {
        if (-not (Test-ObjectProperty -Item $notificationState -Name $propertyName)) {
            $defaultValue = switch ($propertyName) {
                'status' { 'novo' }
                'history' { @() }
                default { '' }
            }
            $notificationState | Add-Member -NotePropertyName $propertyName -NotePropertyValue $defaultValue
        }
    }

    switch ($Action) {
        'read' { $notificationState.status = 'lido' }
        'archive' { $notificationState.status = 'arquivado' }
        'reopen' { $notificationState.status = 'novo' }
    }

    $notificationState.notificationType = if ([string]::IsNullOrWhiteSpace([string]$notificationState.notificationType)) { Collapse-Whitespace -Text $NotificationType } else { [string]$notificationState.notificationType }
    $notificationState.reference = if ([string]::IsNullOrWhiteSpace([string]$notificationState.reference)) { Collapse-Whitespace -Text $Reference } else { [string]$notificationState.reference }
    $notificationState.title = if ([string]::IsNullOrWhiteSpace([string]$notificationState.title)) { Collapse-Whitespace -Text $Title } else { [string]$notificationState.title }
    $notificationState.updatedAt = Get-IsoNow
    $notificationState.updatedBy = [string]$User.login

    $notificationHistory = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($notificationState.history)) {
        $notificationHistory.Add($entry)
    }
    $notificationHistory.Add([pscustomobject][ordered]@{
        action = $Action
        actor = [string]$User.login
        createdAt = Get-IsoNow
        status = [string]$notificationState.status
    })
    $notificationState.history = @(
        @($notificationHistory.ToArray()) |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true } |
        Select-Object -First 12
    )

    $actionLabel = switch ($Action) {
        'read' { 'Notificacao marcada como lida.' }
        'archive' { 'Notificacao arquivada.' }
        'reopen' { 'Notificacao reaberta.' }
    }

    Add-WorkspaceActivity -Payload $payload -Type 'notification_action' -Title 'Notificacao atualizada' -Summary $actionLabel -Reference ([string]$notificationState.reference) -CreatedBy ([string]$User.login) | Out-Null
    Save-WorkspacePayload -Payload $payload
    Register-ObservabilityEvent -Type 'notification_action' -Status $Action -Message $actionLabel -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        notificationKey = $cleanNotificationKey
        reference = [string]$notificationState.reference
    }) | Out-Null
    return $notificationState
}

function Invoke-WorkspaceBatchAction {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [ValidateSet('workflow', 'alert', 'notification', 'favorite')]
        [string]$Target,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Items = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Status = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$AssigneeLogin = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$DueDate = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Note = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Justification = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$SnoozeUntil = '',

        [Parameter(Mandatory = $false)]
        [bool]$IsFavorite = $true
    )

    $normalizedTarget = ([string]$Target).Trim().ToLowerInvariant()
    $normalizedAction = ([string]$Action).Trim().ToLowerInvariant()
    $entries = @($Items)
    if (@($entries).Count -eq 0) {
        throw 'Nenhum item foi informado para a operacao em lote.'
    }

    $processed = New-Object System.Collections.Generic.List[object]
    $processedKeyLookup = New-Object 'System.Collections.Generic.HashSet[string]'

    switch ($normalizedTarget) {
        'workflow' {
            if ([string]::IsNullOrWhiteSpace([string]$Status)) {
                throw 'Informe o status que deve ser aplicado ao workflow em lote.'
            }

            foreach ($entry in @($entries)) {
                $reference = if ($entry -is [string]) {
                    Collapse-Whitespace -Text ([string]$entry)
                }
                else {
                    Collapse-Whitespace -Text (Get-ObjectStringValue -Item $entry -Name 'reference')
                }
                if ([string]::IsNullOrWhiteSpace($reference)) {
                    continue
                }
                if (-not $processedKeyLookup.Add($reference)) {
                    continue
                }

                $workflow = Set-WorkspaceContractWorkflow `
                    -User $User `
                    -Reference $reference `
                    -Status $Status `
                    -AssigneeLogin $AssigneeLogin `
                    -DueDate $DueDate `
                    -Note $Note

                $processed.Add([pscustomobject][ordered]@{
                    target = 'workflow'
                    reference = $reference
                    status = [string]$workflow.status
                    assigneeLogin = [string]$workflow.assigneeLogin
                    dueDate = [string]$workflow.dueDate
                })
            }
        }
        'alert' {
            foreach ($entry in @($entries)) {
                $alertKey = if ($entry -is [string]) {
                    Collapse-Whitespace -Text ([string]$entry)
                }
                else {
                    Collapse-Whitespace -Text (Get-ObjectStringValue -Item $entry -Name 'alertKey')
                }
                if ([string]::IsNullOrWhiteSpace($alertKey)) {
                    continue
                }
                if (-not $processedKeyLookup.Add($alertKey)) {
                    continue
                }

                $alertType = if ($entry -is [string]) { '' } else { Get-ObjectStringValue -Item $entry -Name 'alertType' }
                $reference = if ($entry -is [string]) { '' } else { Get-ObjectStringValue -Item $entry -Name 'reference' }
                $title = if ($entry -is [string]) { '' } else { Get-ObjectStringValue -Item $entry -Name 'title' }

                $alert = Set-WorkspaceAlertState `
                    -User $User `
                    -AlertKey $alertKey `
                    -AlertType $alertType `
                    -Reference $reference `
                    -Title $title `
                    -Action $normalizedAction `
                    -AssigneeLogin $AssigneeLogin `
                    -DueDate $DueDate `
                    -SnoozeUntil $SnoozeUntil `
                    -Justification $Justification

                $processed.Add([pscustomobject][ordered]@{
                    target = 'alert'
                    alertKey = $alertKey
                    reference = [string]$alert.reference
                    status = [string]$alert.status
                    assigneeLogin = [string]$alert.assigneeLogin
                    dueDate = [string]$alert.dueDate
                })
            }
        }
        'notification' {
            foreach ($entry in @($entries)) {
                $notificationKey = if ($entry -is [string]) {
                    Collapse-Whitespace -Text ([string]$entry)
                }
                else {
                    Collapse-Whitespace -Text (Get-ObjectStringValue -Item $entry -Name 'notificationKey')
                }
                if ([string]::IsNullOrWhiteSpace($notificationKey)) {
                    continue
                }
                if (-not $processedKeyLookup.Add($notificationKey)) {
                    continue
                }

                $notificationType = if ($entry -is [string]) { '' } else { Get-ObjectStringValue -Item $entry -Name 'notificationType' }
                $reference = if ($entry -is [string]) { '' } else { Get-ObjectStringValue -Item $entry -Name 'reference' }
                $title = if ($entry -is [string]) { '' } else { Get-ObjectStringValue -Item $entry -Name 'title' }

                $notification = Set-WorkspaceNotificationState `
                    -User $User `
                    -NotificationKey $notificationKey `
                    -NotificationType $notificationType `
                    -Reference $reference `
                    -Title $title `
                    -Action $normalizedAction

                $processed.Add([pscustomobject][ordered]@{
                    target = 'notification'
                    notificationKey = $notificationKey
                    reference = [string]$notification.reference
                    status = [string]$notification.status
                })
            }
        }
        'favorite' {
            foreach ($entry in @($entries)) {
                $reference = if ($entry -is [string]) {
                    Collapse-Whitespace -Text ([string]$entry)
                }
                else {
                    Collapse-Whitespace -Text (Get-ObjectStringValue -Item $entry -Name 'reference')
                }
                if ([string]::IsNullOrWhiteSpace($reference)) {
                    continue
                }
                if (-not $processedKeyLookup.Add($reference)) {
                    continue
                }

                $favorite = Set-WorkspaceFavorite -User $User -Reference $reference -IsFavorite:$IsFavorite
                $processed.Add([pscustomobject][ordered]@{
                    target = 'favorite'
                    reference = $reference
                    isFavorite = [bool]$favorite.isFavorite
                })
            }
        }
    }

    $processedItems = @($processed.ToArray())
    if (@($processedItems).Count -eq 0) {
        throw 'Nenhum item valido foi encontrado para a operacao em lote.'
    }

    $summaryLabel = switch ($normalizedTarget) {
        'workflow' { 'workflow contratual' }
        'alert' { 'alerta operacional' }
        'notification' { 'notificacao' }
        'favorite' { 'favorito' }
    }

    $workspacePayload = Get-WorkspacePayload
    Add-WorkspaceActivity -Payload $workspacePayload -Type 'batch_action' -Title 'Operacao em lote' -Summary "$([int]@($processedItems).Count) item(ns) de $summaryLabel tratado(s) em lote." -CreatedBy ([string]$User.login) -Metadata ([ordered]@{
        target = $normalizedTarget
        action = $normalizedAction
        count = [int]@($processedItems).Count
    }) | Out-Null
    Save-WorkspacePayload -Payload $workspacePayload
    Register-ObservabilityEvent -Type 'batch_action' -Status $normalizedAction -Message "Operacao em lote concluida para $summaryLabel." -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        target = $normalizedTarget
        count = [int]@($processedItems).Count
    }) | Out-Null

    return [ordered]@{
        target = $normalizedTarget
        action = $normalizedAction
        count = [int]@($processedItems).Count
        items = @($processedItems)
    }
}

function Merge-WorkspaceNotificationEntry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$BaseNotification,

        [Parameter(Mandatory = $true)]
        [object]$WorkspacePayload,

        [Parameter(Mandatory = $true)]
        [object]$User
    )

    $stringifyValue = {
        param([AllowNull()][object]$Value)
        if ($null -eq $Value) {
            return ''
        }
        if ($Value -is [string]) {
            return [string]$Value
        }
        if ($Value -is [System.Collections.IEnumerable]) {
            return [string]::Join(' | ', @($Value | ForEach-Object { [string]$_ }))
        }
        return [string]$Value
    }

    $baseType = if (Test-ObjectProperty -Item $BaseNotification -Name 'type') { & $stringifyValue (Get-ObjectStringValue -Item $BaseNotification -Name 'type') } else { '' }
    $baseReference = if (Test-ObjectProperty -Item $BaseNotification -Name 'reference') { & $stringifyValue (Get-ObjectStringValue -Item $BaseNotification -Name 'reference') } else { '' }
    $baseTitle = if (Test-ObjectProperty -Item $BaseNotification -Name 'title') { & $stringifyValue (Get-ObjectStringValue -Item $BaseNotification -Name 'title') } else { '' }
    $existingNotificationKey = if (Test-ObjectProperty -Item $BaseNotification -Name 'notificationKey') { Get-ObjectStringValue -Item $BaseNotification -Name 'notificationKey' } else { '' }
    $notificationKey = if ([string]::IsNullOrWhiteSpace($existingNotificationKey)) {
        Get-WorkspaceNotificationKey -Type $baseType -Reference $baseReference -Title $baseTitle
    }
    else {
        $existingNotificationKey
    }

    $workspaceNotificationStates = if (Test-ObjectProperty -Item $WorkspacePayload -Name 'notificationStates') {
        @($WorkspacePayload.notificationStates)
    }
    else {
        @()
    }
    $state = @($workspaceNotificationStates | Where-Object {
        [string]$_.notificationKey -eq [string]$notificationKey -and
        [string]$_.userId -eq [string]$User.id
    } | Select-Object -First 1) | Select-Object -First 1

    $merged = [ordered]@{}
    foreach ($property in @(Get-ObjectPropertyEntries -Item $BaseNotification)) {
        $merged[$property.Name] = $property.Value
    }
    $merged['notificationKey'] = $notificationKey
    $merged['notificationType'] = $baseType
    $merged['stateStatus'] = if ($state) { [string]$state.status } else { 'novo' }
    $merged['stateUpdatedAt'] = if ($state) { [string]$state.updatedAt } else { '' }
    $merged['stateUpdatedBy'] = if ($state) { [string]$state.updatedBy } else { '' }
    $merged['stateHistory'] = if ($state) { @($state.history) } else { @() }
    $merged['isArchived'] = [bool]($state -and [string]$state.status -eq 'arquivado')
    $merged['isRead'] = [bool]($state -and [string]$state.status -in @('lido', 'arquivado'))
    $effectiveDueDate = if ($state -and [string]$state.dueDate) { [string]$state.dueDate } elseif (Test-ObjectProperty -Item $BaseNotification -Name 'dueDate') { [string]$BaseNotification.dueDate } else { '' }
    $createdAtValue = if (Test-ObjectProperty -Item $BaseNotification -Name 'createdAt') { [string]$BaseNotification.createdAt } else { '' }
    $ageDays = $null
    if (-not [string]::IsNullOrWhiteSpace($createdAtValue)) {
        try {
            $ageDays = [int][math]::Floor(((Get-Date) - [DateTime]::Parse($createdAtValue)).TotalDays)
        }
        catch {
            $ageDays = $null
        }
    }
    $dueToday = $false
    $overdue = $false
    $daysUntilDue = $null
    if (-not [string]::IsNullOrWhiteSpace($effectiveDueDate)) {
        try {
            $parsedDueDate = [DateTime]::Parse($effectiveDueDate).Date
            $daysUntilDue = [int](($parsedDueDate - (Get-Date).Date).TotalDays)
            $dueToday = ($daysUntilDue -eq 0)
            $overdue = ($daysUntilDue -lt 0)
        }
        catch {
            $daysUntilDue = $null
        }
    }

    $baseCategoryLabel = if (Test-ObjectProperty -Item $BaseNotification -Name 'categoryLabel') {
        Get-ObjectStringValue -Item $BaseNotification -Name 'categoryLabel'
    }
    else {
        ''
    }
    $categoryLabel = if (-not [string]::IsNullOrWhiteSpace($baseCategoryLabel)) {
        [string]$baseCategoryLabel
    }
    else {
        switch ($baseType) {
            'alert_notification' { 'Central operacional' }
            'workflow_notification' { 'Workflow contratual' }
            'support_notification' { 'Suporte interno' }
            'review_notification' { 'Revisao manual' }
            'expiring_contract' { 'Vigencia contratual' }
            'sync_notification' { 'Observabilidade' }
            'financial_automation' { 'Financeiro oficial' }
            'change_notification' { 'Mudanca entre sincronizacoes' }
            default { 'Fila do painel' }
        }
    }
    $baseReason = if (Test-ObjectProperty -Item $BaseNotification -Name 'reason') {
        Get-ObjectStringValue -Item $BaseNotification -Name 'reason'
    }
    else {
        ''
    }
    $reason = if (-not [string]::IsNullOrWhiteSpace($baseReason)) {
        [string]$baseReason
    }
    elseif (Test-ObjectProperty -Item $BaseNotification -Name 'summary') {
        [string]$BaseNotification.summary
    }
    else {
        ''
    }
    $baseNextStep = if (Test-ObjectProperty -Item $BaseNotification -Name 'nextStep') {
        Get-ObjectStringValue -Item $BaseNotification -Name 'nextStep'
    }
    else {
        ''
    }
    $nextStep = if (-not [string]::IsNullOrWhiteSpace($baseNextStep)) {
        [string]$baseNextStep
    }
    else {
        switch ($baseType) {
            'alert_notification' { 'Abrir o alerta e registrar responsavel, prazo e justificativa.' }
            'workflow_notification' { 'Atualizar o workflow do contrato e confirmar o encaminhamento.' }
            'support_notification' { 'Responder ou concluir o chamado de suporte em aberto.' }
            'review_notification' { 'Revisar o vinculo entre as fontes e salvar a decisao manual.' }
            'expiring_contract' { 'Conferir vigencia, renovacao e providencias do contrato.' }
            'sync_notification' { 'Abrir observabilidade e validar a falha da ultima sincronizacao.' }
            'financial_automation' { 'Checar as fontes financeiras e a prontidao da automacao oficial.' }
            'change_notification' { 'Abrir o dossie e avaliar o impacto da mudanca entre snapshots.' }
            default { 'Abrir o item e registrar a proxima decisao operacional.' }
        }
    }

    $priorityScore = if (Test-ObjectProperty -Item $BaseNotification -Name 'priorityScore') {
        try { [int]$BaseNotification.priorityScore } catch { 0 }
    }
    else {
        switch ([string]$merged['tone']) {
            'critical' { 80 }
            'warning' { 55 }
            default { 30 }
        }
    }
    if ($overdue) {
        $priorityScore += 35
    }
    elseif ($dueToday) {
        $priorityScore += 20
    }
    elseif ($null -ne $daysUntilDue -and $daysUntilDue -le 3) {
        $priorityScore += 10
    }
    if (-not [bool]$merged['isRead']) {
        $priorityScore += 5
    }
    if ($state -and -not [string]::IsNullOrWhiteSpace([string]$state.assigneeLogin) -and [string]$state.assigneeLogin -eq [string]$User.login) {
        $priorityScore += 5
    }

    $baseSlaLabel = if (Test-ObjectProperty -Item $BaseNotification -Name 'slaLabel') {
        Get-ObjectStringValue -Item $BaseNotification -Name 'slaLabel'
    }
    else {
        ''
    }
    $slaLabel = if (-not [string]::IsNullOrWhiteSpace($baseSlaLabel)) {
        [string]$baseSlaLabel
    }
    elseif ($overdue) {
        "Prazo vencido ha $([math]::Abs([int]$daysUntilDue)) dia(s)"
    }
    elseif ($dueToday) {
        'Prazo vence hoje'
    }
    elseif ($null -ne $daysUntilDue) {
        "Prazo em $([int]$daysUntilDue) dia(s)"
    }
    elseif (-not [string]::IsNullOrWhiteSpace($effectiveDueDate)) {
        "Prazo $effectiveDueDate"
    }
    else {
        'Sem prazo operacional'
    }
    $baseManagerialLabel = if (Test-ObjectProperty -Item $BaseNotification -Name 'managerialLabel') {
        Get-ObjectStringValue -Item $BaseNotification -Name 'managerialLabel'
    }
    else {
        ''
    }
    $managerialLabel = if (-not [string]::IsNullOrWhiteSpace($baseManagerialLabel)) {
        [string]$baseManagerialLabel
    }
    elseif ($overdue) {
        'Em atraso'
    }
    elseif ($dueToday) {
        'Agir hoje'
    }
    elseif ($priorityScore -ge 95) {
        'Prioridade alta'
    }
    elseif ([string]$merged['tone'] -eq 'critical') {
        'Monitorar com urgencia'
    }
    elseif ([string]$merged['tone'] -eq 'warning') {
        'Tratar nesta rodada'
    }
    else {
        'Acompanhar'
    }
    $merged['effectiveDueDate'] = $effectiveDueDate
    $merged['ageDays'] = $ageDays
    $merged['isStale'] = [bool]($null -ne $ageDays -and $ageDays -ge 7 -and -not [bool]$merged['isRead'])
    $merged['priorityScore'] = [int]$priorityScore
    $merged['categoryLabel'] = $categoryLabel
    $merged['reason'] = Collapse-Whitespace -Text $reason
    $merged['nextStep'] = Collapse-Whitespace -Text $nextStep
    $merged['managerialLabel'] = $managerialLabel
    $merged['slaLabel'] = $slaLabel
    return [pscustomobject]$merged
}

function Get-WorkspaceGuideData {
    $sources = @(
        [pscustomobject][ordered]@{
            key = 'diario_oficial'
            label = 'Diario Oficial de Iguape'
            href = ([Uri]::new($script:BasePortalUri, $script:PortalDiarioPath)).AbsoluteUri
            role = 'publicacao_oficial'
            note = 'Fonte primaria dos atos publicados, incluindo extratos, designacoes, aditivos e rescisões.'
            automationStatus = 'ready'
        }
        [pscustomobject][ordered]@{
            key = 'portal_contratos'
            label = 'Portal de Contratos do Municipio'
            href = ([Uri]::new($script:BasePortalUri, '/portal/contratos')).AbsoluteUri
            role = 'cadastro_oficial'
            note = 'Base oficial de contratos e documentos associados usada como fonte principal do cadastro contratual.'
            automationStatus = 'ready'
        }
    ) + @(Get-ContractFinancialMonitoringSources)
    $shortcutCatalog = @(Get-SearchShortcutCatalog)

    return [ordered]@{
        highlights = @(
            [pscustomobject][ordered]@{
                key = 'vinculo'
                title = 'Como ler um vinculo'
                summary = 'Use o dossie para comparar portal, Diario e cobertura financeira antes de confirmar uma correspondencia ambigua.'
                href = '/guia.html#glossario'
            }
            [pscustomobject][ordered]@{
                key = 'triagem'
                title = 'Roteiro de triagem'
                summary = 'Priorize itens criticos, filas com prazo vencido e contratos sem responsavel operacional definido.'
                href = '/guia.html#triagem'
            }
            [pscustomobject][ordered]@{
                key = 'busca'
                title = 'Atalhos de busca'
                summary = 'A busca global aceita prefixos por contrato, processo, fornecedor, alerta, suporte e usuario.'
                href = '/guia.html#atalhos'
            }
            [pscustomobject][ordered]@{
                key = 'fontes'
                title = 'Fontes oficiais'
                summary = 'O painel cruza Diario Oficial, portal contratual e monitoramento financeiro oficial com niveis diferentes de automacao.'
                href = '/guia.html#fontes'
            }
        )
        glossary = @(
            [pscustomobject][ordered]@{ term = 'Vinculo'; description = 'Relacao entre um cadastro oficial e um ato do Diario confirmada automaticamente ou por revisao manual.' }
            [pscustomobject][ordered]@{ term = 'Confianca'; description = 'Indicador do quanto o cruzamento depende de numero, processo, orgao, fornecedor, datas e outros sinais combinados.' }
            [pscustomobject][ordered]@{ term = 'Divergencia'; description = 'Diferenca material entre fontes, como status, valor, vigencia, orgao ou ausencia de registro esperado.' }
            [pscustomobject][ordered]@{ term = 'Vigencia'; description = 'Janela em que o contrato permanece ativo segundo o portal ou segundo a leitura do documento publicado.' }
            [pscustomobject][ordered]@{ term = 'Cobertura financeira'; description = 'Quanto do contrato ja tem fonte oficial de execucao vinculada, seja por API, painel oficial ou consulta assistida.' }
        )
        triage = @(
            [pscustomobject][ordered]@{ title = '1. Filtre o que exige decisao hoje'; description = 'Comece pela fila de decisao, alertas criticos, revisoes manuais pendentes e itens com prazo vencido.' }
            [pscustomobject][ordered]@{ title = '2. Abra o comparador do contrato'; description = 'Confirme referencia, orgao, fornecedor, valor e situacao lado a lado antes de decidir.' }
            [pscustomobject][ordered]@{ title = '3. Registre a decisao'; description = 'Use workflow, comentario interno e acao de alerta para deixar a trilha auditavel.' }
            [pscustomobject][ordered]@{ title = '4. Feche o ciclo'; description = 'Arquive a notificacao pessoal e mantenha o alerta operacional resolvido ou reaberto conforme o caso.' }
        )
        shortcuts = @($shortcutCatalog)
        sources = @($sources)
        accessibility = @(
            [pscustomobject][ordered]@{ title = 'Busca global'; description = 'A busca fica sempre visivel no topo e aceita refinamento por prefixo.' }
            [pscustomobject][ordered]@{ title = 'Foco visivel'; description = 'Campos e acoes operacionais destacam foco para navegação por teclado.' }
            [pscustomobject][ordered]@{ title = 'Leitura por blocos'; description = 'As telas foram separadas em paineis curtos para evitar sobrecarga em mobile e tela reduzida.' }
        )
    }
}

function Get-AggregateSnapshotStringList {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Values = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 40
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($Values)) {
        $text = Collapse-Whitespace -Text ([string]$value)
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }
        if ($set.Add($text)) {
            $list.Add($text)
        }
        if ($list.Count -ge $Limit) {
            break
        }
    }

    return @($list.ToArray())
}

function Get-AggregateSnapshotOrganizationRows {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$OrganizationSummary = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 12
    )

    return @(
        @($OrganizationSummary) |
        Sort-Object @{ Expression = { [int]$_.count }; Descending = $true }, @{ Expression = { [double]$_.totalValue }; Descending = $true }, @{ Expression = { [string]$_.name }; Descending = $false } |
        Select-Object -First $Limit |
        ForEach-Object {
            [pscustomobject][ordered]@{
                organizationId = [string]$_.organizationId
                name = [string]$_.name
                count = [int]$_.count
                totalValue = [double]$_.totalValue
            }
        }
    )
}

function Add-AggregateSnapshotLookupCount {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Lookup,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Key,

        [Parameter(Mandatory = $false)]
        [int]$Increment = 1
    )

    $cleanKey = Collapse-Whitespace -Text $Key
    if ([string]::IsNullOrWhiteSpace($cleanKey)) {
        return
    }

    $Lookup[$cleanKey] = [int]$(if ($Lookup.ContainsKey($cleanKey)) { $Lookup[$cleanKey] } else { 0 }) + $Increment
}

function Get-AggregateSnapshotSupplierRows {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Contracts = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 12
    )

    $supplierCounter = @{}
    foreach ($contract in @($Contracts)) {
        $supplierName = Collapse-Whitespace -Text ([string]$contract.contractor)
        if ([string]::IsNullOrWhiteSpace($supplierName)) {
            continue
        }

        if (-not $supplierCounter.ContainsKey($supplierName)) {
            $supplierCounter[$supplierName] = [ordered]@{
                name = $supplierName
                count = 0
                totalValue = 0.0
            }
        }

        $supplierCounter[$supplierName].count = [int]$supplierCounter[$supplierName].count + 1
        $valueNumber = if ($null -ne $contract.valueNumber) { [double]$contract.valueNumber } else { (Convert-BrazilianCurrencyToNumber -Text ([string]$contract.value)) }
        $supplierCounter[$supplierName].totalValue = [math]::Round(([double]$supplierCounter[$supplierName].totalValue + $valueNumber), 2)
    }

    return @(
        $supplierCounter.Values |
        Sort-Object @{ Expression = { [int]$_.count }; Descending = $true }, @{ Expression = { [double]$_.totalValue }; Descending = $true }, @{ Expression = { [string]$_.name }; Descending = $false } |
        Select-Object -First $Limit |
        ForEach-Object {
            [pscustomobject][ordered]@{
                name = [string]$_.name
                count = [int]$_.count
                totalValue = [double]$_.totalValue
            }
        }
    )
}

function Get-AggregateSnapshotContractRows {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$OfficialContracts = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$CrossSourceAlerts = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$CrossSourceDivergences = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$CrossReviewQueue = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 400
    )

    $alertLookup = @{}
    foreach ($alert in @($CrossSourceAlerts)) {
        Add-AggregateSnapshotLookupCount -Lookup $alertLookup -Key ([string]$alert.crossKey)
        Add-AggregateSnapshotLookupCount -Lookup $alertLookup -Key ([string]$alert.portalContractId)
        Add-AggregateSnapshotLookupCount -Lookup $alertLookup -Key ([string]$alert.contractNumber)
    }

    $divergenceLookup = @{}
    foreach ($entry in @($CrossSourceDivergences)) {
        Add-AggregateSnapshotLookupCount -Lookup $divergenceLookup -Key ([string]$entry.crossKey)
        Add-AggregateSnapshotLookupCount -Lookup $divergenceLookup -Key ([string]$entry.portalContractId)
        Add-AggregateSnapshotLookupCount -Lookup $divergenceLookup -Key ([string]$entry.contractNumber)
    }

    $reviewLookup = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($review in @($CrossReviewQueue)) {
        foreach ($candidateKey in @([string]$review.crossKey, [string]$review.movementReference, [string]$review.movementKey)) {
            $cleanCandidate = Collapse-Whitespace -Text $candidateKey
            if (-not [string]::IsNullOrWhiteSpace($cleanCandidate)) {
                $null = $reviewLookup.Add($cleanCandidate)
            }
        }
    }

    return @(
        @($OfficialContracts) |
        Sort-Object @{ Expression = { [string]$_.publishedAt }; Descending = $true }, @{ Expression = { [string]$(if ([string]$_.referenceKey) { $_.referenceKey } else { $_.contractNumber }) }; Descending = $false } |
        Select-Object -First $Limit |
        ForEach-Object {
            $reference = [string]$(if ([string]$_.referenceKey) { $_.referenceKey } elseif ([string]$_.contractNumber) { $_.contractNumber } elseif ([string]$_.processNumber) { $_.processNumber } else { $_.portalContractId })
            $crossKey = [string]$(if ($_.PSObject.Properties['crossSource'] -and [string]$_.crossSource.crossKey) { $_.crossSource.crossKey } else { $reference })
            $managementStatus = if (-not [bool]$_.managementTracked) {
                'nao_monitorada'
            }
            elseif ([bool]$_.hasManager -and [bool]$_.hasInspector) {
                'completa'
            }
            elseif (-not [bool]$_.hasManager -and -not [bool]$_.hasInspector) {
                'sem_gestor_e_fiscal'
            }
            elseif (-not [bool]$_.hasManager) {
                'sem_gestor'
            }
            else {
                'sem_fiscal'
            }

            [pscustomobject][ordered]@{
                reference = $reference
                contractNumber = [string]$_.contractNumber
                processNumber = [string]$_.processNumber
                portalContractId = [string]$_.portalContractId
                crossKey = $crossKey
                title = [string]$(if ([string]$_.actTitle) { $_.actTitle } elseif ([string]$_.contractNumber) { $_.contractNumber } else { $reference })
                organization = [string]$_.primaryOrganizationName
                contractor = [string]$_.contractor
                value = [string]$_.value
                valueNumber = if ($null -ne $_.valueNumber) { [double]$_.valueNumber } else { (Convert-BrazilianCurrencyToNumber -Text ([string]$_.value)) }
                portalStatus = [string]$_.portalStatus
                publishedAt = [string]$_.publishedAt
                isActive = [bool]$_.vigency.isActive
                vigencyLabel = [string]$(if ($_.vigency -and [string]$_.vigency.summaryLabel) { $_.vigency.summaryLabel } elseif ([bool]$_.vigency.isActive) { 'Contrato vigente' } else { 'Vigencia indefinida' })
                daysUntilEnd = if ($null -ne $_.vigency.daysUntilEnd) { [int]$_.vigency.daysUntilEnd } else { $null }
                managementStatus = $managementStatus
                hasManager = [bool]$_.hasManager
                hasInspector = [bool]$_.hasInspector
                hasExonerationSignal = [bool]$_.managerExonerationSignal -or [bool]$_.inspectorExonerationSignal
                reviewPending = [bool]([string]$_.crossSource.status -eq 'pending_review' -or $reviewLookup.Contains($crossKey) -or $reviewLookup.Contains($reference))
                crossStatus = [string]$(if ($_.PSObject.Properties['crossSource'] -and [string]$_.crossSource.status) { $_.crossSource.status } else { 'unmatched' })
                alertCount = [int]$(
                    if ($alertLookup.ContainsKey($crossKey)) { $alertLookup[$crossKey] }
                    elseif ($alertLookup.ContainsKey([string]$_.portalContractId)) { $alertLookup[[string]$_.portalContractId] }
                    elseif ($alertLookup.ContainsKey([string]$_.contractNumber)) { $alertLookup[[string]$_.contractNumber] }
                    else { 0 }
                )
                divergenceCount = [int]$(
                    if ($divergenceLookup.ContainsKey($crossKey)) { $divergenceLookup[$crossKey] }
                    elseif ($divergenceLookup.ContainsKey([string]$_.portalContractId)) { $divergenceLookup[[string]$_.portalContractId] }
                    elseif ($divergenceLookup.ContainsKey([string]$_.contractNumber)) { $divergenceLookup[[string]$_.contractNumber] }
                    else { 0 }
                )
                localDocument = [bool](-not [string]::IsNullOrWhiteSpace([string]$_.localPdfRelative))
            }
        }
    )
}

function Get-SnapshotContractStatusLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    switch ($Key) {
        'isActive' { return $(if ([bool]$Value) { 'Vigente' } else { 'Nao vigente' }) }
        'reviewPending' { return $(if ([bool]$Value) { 'Pendente' } else { 'Sem pendencia' }) }
        'hasExonerationSignal' { return $(if ([bool]$Value) { 'Com sinal' } else { 'Sem sinal' }) }
        'localDocument' { return $(if ([bool]$Value) { 'Com arquivo local' } else { 'Sem arquivo local' }) }
        'managementStatus' {
            switch ([string]$Value) {
                'completa' { return 'Gestao completa' }
                'sem_gestor_e_fiscal' { return 'Sem gestor e fiscal' }
                'sem_gestor' { return 'Sem gestor' }
                'sem_fiscal' { return 'Sem fiscal' }
                default { return 'Nao monitorada' }
            }
        }
        'crossStatus' {
            switch ([string]$Value) {
                'reviewed' { return 'Vinculo revisado' }
                'pending_review' { return 'Pendente de revisao' }
                'matched' { return 'Vinculo automatico' }
                default { return 'Sem vinculo' }
            }
        }
        'valueNumber' {
            $number = 0.0
            try { $number = [double]$Value } catch { $number = 0.0 }
            return ("R$ {0:N2}" -f $number)
        }
        'daysUntilEnd' {
            if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
                return 'Sem prazo final'
            }
            return "$([int]$Value) dia(s)"
        }
        default {
            if ($null -eq $Value) {
                return ''
            }
            return Collapse-Whitespace -Text ([string]$Value)
        }
    }
}

function Get-SnapshotContractVersionChanges {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$CurrentRow,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$PreviousRow
    )

    if ($null -eq $CurrentRow -or $null -eq $PreviousRow) {
        return @()
    }

    $fieldDefinitions = @(
        [ordered]@{ key = 'organization'; label = 'Orgao principal'; severity = 15 }
        [ordered]@{ key = 'contractor'; label = 'Fornecedor'; severity = 15 }
        [ordered]@{ key = 'valueNumber'; label = 'Valor'; severity = 25 }
        [ordered]@{ key = 'portalStatus'; label = 'Status no portal'; severity = 25 }
        [ordered]@{ key = 'isActive'; label = 'Vigencia'; severity = 40 }
        [ordered]@{ key = 'daysUntilEnd'; label = 'Prazo final'; severity = 25 }
        [ordered]@{ key = 'managementStatus'; label = 'Gestao'; severity = 35 }
        [ordered]@{ key = 'crossStatus'; label = 'Vinculo'; severity = 30 }
        [ordered]@{ key = 'reviewPending'; label = 'Revisao manual'; severity = 40 }
        [ordered]@{ key = 'divergenceCount'; label = 'Divergencias'; severity = 35 }
        [ordered]@{ key = 'alertCount'; label = 'Alertas'; severity = 35 }
        [ordered]@{ key = 'hasExonerationSignal'; label = 'Sinal de exoneracao'; severity = 45 }
        [ordered]@{ key = 'localDocument'; label = 'Documento local'; severity = 20 }
    )

    $changes = New-Object System.Collections.Generic.List[object]
    foreach ($definition in @($fieldDefinitions)) {
        $key = [string]$definition.key
        $currentValue = if (Test-ObjectProperty -Item $CurrentRow -Name $key) { $CurrentRow.$key } else { $null }
        $previousValue = if (Test-ObjectProperty -Item $PreviousRow -Name $key) { $PreviousRow.$key } else { $null }
        $isDifferent = if ($key -eq 'valueNumber') {
            [math]::Round([double]$(if ($null -ne $currentValue) { $currentValue } else { 0 }), 2) -ne [math]::Round([double]$(if ($null -ne $previousValue) { $previousValue } else { 0 }), 2)
        }
        else {
            [string]$currentValue -ne [string]$previousValue
        }

        if (-not $isDifferent) {
            continue
        }

        $changes.Add([pscustomobject][ordered]@{
            key = $key
            label = [string]$definition.label
            previous = Get-SnapshotContractStatusLabel -Key $key -Value $previousValue
            current = Get-SnapshotContractStatusLabel -Key $key -Value $currentValue
            severity = [int]$definition.severity
        }) | Out-Null
    }

    return @($changes.ToArray())
}

function Find-SnapshotContractRow {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Rows = @(),

        [Parameter(Mandatory = $true)]
        [string]$Reference
    )

    @(
        @($Rows) |
        Where-Object {
            Test-ReferenceMatch -Reference $Reference -CandidateValues @(
                [string]$_.reference,
                [string]$_.contractNumber,
                [string]$_.processNumber,
                [string]$_.portalContractId,
                [string]$_.crossKey
            )
        } |
        Select-Object -First 1
    ) | Select-Object -First 1
}

function Get-SnapshotArrayProperty {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Snapshot,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Snapshot -or -not (Test-ObjectProperty -Item $Snapshot -Name $Name)) {
        return @()
    }

    return @($Snapshot.$Name)
}

function Get-SnapshotCountDeltaItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$CurrentSnapshot,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$PreviousSnapshot,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $currentValue = if ($CurrentSnapshot -and (Test-ObjectProperty -Item $CurrentSnapshot -Name $PropertyName)) { [int]$CurrentSnapshot.$PropertyName } else { 0 }
    $previousValue = if ($PreviousSnapshot -and (Test-ObjectProperty -Item $PreviousSnapshot -Name $PropertyName)) { [int]$PreviousSnapshot.$PropertyName } else { 0 }
    return [pscustomobject][ordered]@{
        key = $Key
        label = $Label
        current = $currentValue
        previous = $previousValue
        delta = $currentValue - $previousValue
    }
}

function Get-SnapshotStringDeltaGroup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$CurrentValues = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$PreviousValues = @()
    )

    $currentSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($value in @(Get-AggregateSnapshotStringList -Values $CurrentValues -Limit 80)) {
        $null = $currentSet.Add([string]$value)
    }

    $previousSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($value in @(Get-AggregateSnapshotStringList -Values $PreviousValues -Limit 80)) {
        $null = $previousSet.Add([string]$value)
    }

    $added = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($currentSet)) {
        if (-not $previousSet.Contains([string]$value)) {
            $added.Add([string]$value)
        }
    }

    $removed = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($previousSet)) {
        if (-not $currentSet.Contains([string]$value)) {
            $removed.Add([string]$value)
        }
    }

    return [pscustomobject][ordered]@{
        key = $Key
        title = $Title
        description = $Description
        added = @($added.ToArray() | Sort-Object | Select-Object -First 8)
        removed = @($removed.ToArray() | Sort-Object | Select-Object -First 8)
    }
}

function Update-WorkspaceAggregateSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Aggregate
    )

    $payload = Get-WorkspacePayload
    $officialContracts = @($Aggregate.officialContracts)
    $criticalAlerts = @($Aggregate.crossSourceAlerts | Where-Object { [string]$_.severity -eq 'critical' })
    $withoutDocument = @($officialContracts | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.localPdfRelative) })
    $incompleteManagement = @($officialContracts | Where-Object { [bool]$_.managementTracked -and (-not [bool]$_.hasManager -or -not [bool]$_.hasInspector) })
    $activeContracts = @($officialContracts | Where-Object { [bool]$_.vigency.isActive })
    $supplierTop = @(Get-AggregateSnapshotSupplierRows -Contracts $officialContracts)
    $contractVersionRows = @(Get-AggregateSnapshotContractRows -OfficialContracts $officialContracts -CrossSourceAlerts @($Aggregate.crossSourceAlerts) -CrossSourceDivergences @($Aggregate.crossSourceDivergences) -CrossReviewQueue @($Aggregate.crossReviewQueue))
    $contractFingerprintSource = [string]::Join('||', @(
        $contractVersionRows |
        ForEach-Object {
            [string]::Join('|', @(
                [string]$_.reference,
                [string]$_.portalStatus,
                [string]$_.organization,
                [string]$_.contractor,
                [string]$_.valueNumber,
                [string]$_.isActive,
                [string]$_.daysUntilEnd,
                [string]$_.managementStatus,
                [string]$_.crossStatus,
                [string]$_.reviewPending,
                [string]$_.divergenceCount,
                [string]$_.alertCount,
                [string]$_.hasExonerationSignal,
                [string]$_.localDocument
            ))
        }
    ))
    $supplierFingerprintSource = [string]::Join('||', @(
        $supplierTop |
        ForEach-Object {
            [string]::Join('|', @([string]$_.name, [string]$_.count, [string]$_.totalValue))
        }
    ))

    $snapshot = [pscustomobject][ordered]@{
        generatedAt = [string]$Aggregate.generatedAt
        totalItems = [int]$Aggregate.totalItems
        officialContracts = [int]@($officialContracts).Count
        contractMovements = [int]@($Aggregate.contractMovements).Count
        divergences = [int]$Aggregate.crosswalkSummary.divergences
        suppressedDivergences = [int]$Aggregate.crosswalkSummary.suppressedDivergences
        crossReviewQueue = [int]@($Aggregate.crossReviewQueue).Count
        alerts = [int]@($Aggregate.crossSourceAlerts).Count
        criticalAlerts = [int]@($criticalAlerts).Count
        searchableFinancialContracts = [int]$Aggregate.financialMonitoring.searchableContracts
        totalValue = [double]$Aggregate.totalValue
        activeContracts = [int]@($activeContracts).Count
        withoutDocument = [int]@($withoutDocument).Count
        incompleteManagement = [int]@($incompleteManagement).Count
        reviewReferences = @(Get-AggregateSnapshotStringList -Values @($Aggregate.crossReviewQueue | ForEach-Object { if ([string]$_.crossKey) { [string]$_.crossKey } elseif ([string]$_.movementReference) { [string]$_.movementReference } else { [string]$_.movementKey } }))
        criticalAlertReferences = @(Get-AggregateSnapshotStringList -Values @($criticalAlerts | ForEach-Object { if ([string]$_.crossKey) { [string]$_.crossKey } elseif ([string]$_.contractNumber) { [string]$_.contractNumber } else { [string]$_.portalContractId } }))
        activeReferences = @(Get-AggregateSnapshotStringList -Values @($activeContracts | ForEach-Object { if ([string]$_.referenceKey) { [string]$_.referenceKey } elseif ([string]$_.contractNumber) { [string]$_.contractNumber } else { [string]$_.portalContractId } }))
        organizationTop = @(Get-AggregateSnapshotOrganizationRows -OrganizationSummary @($Aggregate.organizationSummary))
        supplierTop = @($supplierTop)
        contractVersionRows = @($contractVersionRows)
        contractFingerprint = Get-TextFingerprint -Text $contractFingerprintSource
        supplierFingerprint = Get-TextFingerprint -Text $supplierFingerprintSource
    }

    $lastSnapshot = @($payload.aggregateSnapshots | Select-Object -First 1) | Select-Object -First 1
    $isDifferent = $true
    if ($lastSnapshot) {
        $isDifferent = (
            [int]$lastSnapshot.totalItems -ne [int]$snapshot.totalItems -or
            [int]$lastSnapshot.officialContracts -ne [int]$snapshot.officialContracts -or
            [int]$lastSnapshot.contractMovements -ne [int]$snapshot.contractMovements -or
            [int]$lastSnapshot.divergences -ne [int]$snapshot.divergences -or
            [int]$lastSnapshot.crossReviewQueue -ne [int]$snapshot.crossReviewQueue -or
            [int]$lastSnapshot.alerts -ne [int]$snapshot.alerts -or
            [int]$(if (Test-ObjectProperty -Item $lastSnapshot -Name 'criticalAlerts') { $lastSnapshot.criticalAlerts } else { 0 }) -ne [int]$snapshot.criticalAlerts -or
            [int]$(if (Test-ObjectProperty -Item $lastSnapshot -Name 'activeContracts') { $lastSnapshot.activeContracts } else { 0 }) -ne [int]$snapshot.activeContracts -or
            [int]$(if (Test-ObjectProperty -Item $lastSnapshot -Name 'withoutDocument') { $lastSnapshot.withoutDocument } else { 0 }) -ne [int]$snapshot.withoutDocument -or
            [int]$(if (Test-ObjectProperty -Item $lastSnapshot -Name 'incompleteManagement') { $lastSnapshot.incompleteManagement } else { 0 }) -ne [int]$snapshot.incompleteManagement -or
            [int]$lastSnapshot.searchableFinancialContracts -ne [int]$snapshot.searchableFinancialContracts -or
            [double]$lastSnapshot.totalValue -ne [double]$snapshot.totalValue -or
            [string]$(if (Test-ObjectProperty -Item $lastSnapshot -Name 'contractFingerprint') { $lastSnapshot.contractFingerprint } else { '' }) -ne [string]$snapshot.contractFingerprint -or
            [string]$(if (Test-ObjectProperty -Item $lastSnapshot -Name 'supplierFingerprint') { $lastSnapshot.supplierFingerprint } else { '' }) -ne [string]$snapshot.supplierFingerprint
        )
    }

    $persistedSnapshots = if ($isDifferent) {
        @(
            (@($payload.aggregateSnapshots) + @($snapshot)) |
            Sort-Object @{ Expression = { [string]$_.generatedAt }; Descending = $true }
        )
    }
    else {
        @(
            @($payload.aggregateSnapshots) |
            Sort-Object @{ Expression = { [string]$_.generatedAt }; Descending = $true }
        )
    }
    $persistedCurrentSnapshot = @($persistedSnapshots | Select-Object -First 1) | Select-Object -First 1
    $persistedPreviousSnapshot = @($persistedSnapshots | Select-Object -Skip 1 -First 1) | Select-Object -First 1
    $existingRecentChanges = if (Test-ObjectProperty -Item $payload -Name 'recentChanges') { $payload.recentChanges } else { $null }
    $recentChangesDirty = -not (Test-WorkspaceAggregateChangeSummaryCurrent -Summary $existingRecentChanges -CurrentSnapshot $persistedCurrentSnapshot -PreviousSnapshot $persistedPreviousSnapshot)
    if ($recentChangesDirty) {
        Set-WorkspacePayloadRecentChanges -Payload $payload -RecentChanges (Get-WorkspaceAggregateChangeSummary -Snapshots $persistedSnapshots)
    }

    if ($isDifferent) {
        $payload.aggregateSnapshots = @($persistedSnapshots)
        Add-WorkspaceActivity -Payload $payload -Type 'snapshot' -Title 'Base recomposta' -Summary "Painel recomposto com $([int]$snapshot.totalItems) registro(s)." -CreatedBy 'system' | Out-Null
    }

    if ($isDifferent -or $recentChangesDirty) {
        Save-WorkspacePayload -Payload $payload
    }

    return $snapshot
}

function Test-WorkspaceAggregateChangeSummaryCurrent {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Summary,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$CurrentSnapshot,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$PreviousSnapshot
    )

    if ($null -eq $Summary) {
        return $false
    }

    if (-not (Test-ObjectProperty -Item $Summary -Name 'detailGroups')) {
        return $false
    }

    $summaryCurrentGeneratedAt = if (Test-ObjectProperty -Item $Summary -Name 'currentGeneratedAt') {
        [string]$Summary.currentGeneratedAt
    }
    else {
        ''
    }
    $summaryPreviousGeneratedAt = if (Test-ObjectProperty -Item $Summary -Name 'previousGeneratedAt') {
        [string]$Summary.previousGeneratedAt
    }
    else {
        ''
    }
    $expectedCurrentGeneratedAt = if ($CurrentSnapshot) { [string]$CurrentSnapshot.generatedAt } else { '' }
    $expectedPreviousGeneratedAt = if ($PreviousSnapshot) { [string]$PreviousSnapshot.generatedAt } else { '' }

    return (
        $summaryCurrentGeneratedAt -eq $expectedCurrentGeneratedAt -and
        $summaryPreviousGeneratedAt -eq $expectedPreviousGeneratedAt
    )
}

function Get-WorkspaceAggregateChangeSummary {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Snapshots = $null
    )

    $workspacePayload = $null
    if ($PSBoundParameters.ContainsKey('Snapshots')) {
        $snapshots = @(@($Snapshots) | Sort-Object @{ Expression = { [string]$_.generatedAt }; Descending = $true })
    }
    else {
        $workspacePayload = Get-WorkspacePayload
        $snapshots = @(@($workspacePayload.aggregateSnapshots) | Sort-Object @{ Expression = { [string]$_.generatedAt }; Descending = $true })
    }
    $currentSnapshot = @($snapshots | Select-Object -First 1) | Select-Object -First 1
    $previousSnapshot = @($snapshots | Select-Object -Skip 1 -First 1) | Select-Object -First 1

    if (-not $PSBoundParameters.ContainsKey('Snapshots')) {
        $cachedSummary = if (Test-ObjectProperty -Item $workspacePayload -Name 'recentChanges') { $workspacePayload.recentChanges } else { $null }
        if (Test-WorkspaceAggregateChangeSummaryCurrent -Summary $cachedSummary -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot) {
            return $cachedSummary
        }
    }

    if ($null -eq $currentSnapshot) {
        return (Get-EmptyWorkspaceAggregateChangeSummary)
    }

    $items = @(
        (Get-SnapshotCountDeltaItem -Key 'total_items' -Label 'Registros revisados' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'totalItems')
        (Get-SnapshotCountDeltaItem -Key 'divergences' -Label 'Divergencias materiais' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'divergences')
        (Get-SnapshotCountDeltaItem -Key 'review_queue' -Label 'Pendencias de revisao' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'crossReviewQueue')
        (Get-SnapshotCountDeltaItem -Key 'alerts' -Label 'Alertas operacionais' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'alerts')
        (Get-SnapshotCountDeltaItem -Key 'critical_alerts' -Label 'Alertas criticos' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'criticalAlerts')
        (Get-SnapshotCountDeltaItem -Key 'active_contracts' -Label 'Contratos vigentes' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'activeContracts')
        (Get-SnapshotCountDeltaItem -Key 'without_document' -Label 'Sem documento local' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'withoutDocument')
        (Get-SnapshotCountDeltaItem -Key 'management_gaps' -Label 'Gestao incompleta' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'incompleteManagement')
        (Get-SnapshotCountDeltaItem -Key 'financial' -Label 'Cobertura financeira pesquisavel' -CurrentSnapshot $currentSnapshot -PreviousSnapshot $previousSnapshot -PropertyName 'searchableFinancialContracts')
    )

    $organizationMovers = New-Object System.Collections.Generic.List[object]
    $currentOrganizationRows = @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'organizationTop')
    $previousOrganizationRows = @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'organizationTop')
    $organizationKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in @($currentOrganizationRows)) {
        $null = $organizationKeys.Add([string]$(if ([string]$row.organizationId) { $row.organizationId } else { [string]$row.name }))
    }
    foreach ($row in @($previousOrganizationRows)) {
        $null = $organizationKeys.Add([string]$(if ([string]$row.organizationId) { $row.organizationId } else { [string]$row.name }))
    }

    foreach ($key in @($organizationKeys)) {
        $currentRow = @($currentOrganizationRows | Where-Object { [string]$(if ([string]$_.organizationId) { $_.organizationId } else { $_.name }) -eq $key } | Select-Object -First 1) | Select-Object -First 1
        $previousRow = @($previousOrganizationRows | Where-Object { [string]$(if ([string]$_.organizationId) { $_.organizationId } else { $_.name }) -eq $key } | Select-Object -First 1) | Select-Object -First 1
        $currentCount = if ($currentRow) { [int]$currentRow.count } else { 0 }
        $previousCount = if ($previousRow) { [int]$previousRow.count } else { 0 }
        $delta = $currentCount - $previousCount
        if ($delta -eq 0) {
            continue
        }

        $organizationMovers.Add([pscustomobject][ordered]@{
            organizationId = if ($currentRow) { [string]$currentRow.organizationId } elseif ($previousRow) { [string]$previousRow.organizationId } else { '' }
            name = if ($currentRow) { [string]$currentRow.name } elseif ($previousRow) { [string]$previousRow.name } else { $key }
            current = $currentCount
            previous = $previousCount
            delta = $delta
            totalValue = if ($currentRow) { [double]$currentRow.totalValue } elseif ($previousRow) { [double]$previousRow.totalValue } else { 0.0 }
        })
    }

    $supplierMovers = New-Object System.Collections.Generic.List[object]
    $currentSupplierRows = @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'supplierTop')
    $previousSupplierRows = @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'supplierTop')
    $supplierKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in @($currentSupplierRows)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.name)) {
            $null = $supplierKeys.Add([string]$row.name)
        }
    }
    foreach ($row in @($previousSupplierRows)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.name)) {
            $null = $supplierKeys.Add([string]$row.name)
        }
    }

    foreach ($key in @($supplierKeys)) {
        $currentRow = @($currentSupplierRows | Where-Object { [string]$_.name -eq $key } | Select-Object -First 1) | Select-Object -First 1
        $previousRow = @($previousSupplierRows | Where-Object { [string]$_.name -eq $key } | Select-Object -First 1) | Select-Object -First 1
        $currentCount = if ($currentRow) { [int]$currentRow.count } else { 0 }
        $previousCount = if ($previousRow) { [int]$previousRow.count } else { 0 }
        $delta = $currentCount - $previousCount
        if ($delta -eq 0) {
            continue
        }

        $supplierMovers.Add([pscustomobject][ordered]@{
            name = if ($currentRow) { [string]$currentRow.name } elseif ($previousRow) { [string]$previousRow.name } else { $key }
            current = $currentCount
            previous = $previousCount
            delta = $delta
            totalValue = if ($currentRow) { [double]$currentRow.totalValue } elseif ($previousRow) { [double]$previousRow.totalValue } else { 0.0 }
        }) | Out-Null
    }

    $contractChangeEntries = New-Object System.Collections.Generic.List[object]
    $currentContractRows = @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'contractVersionRows')
    $previousContractRows = @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'contractVersionRows')
    $contractReferences = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in @($currentContractRows)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.reference)) {
            $null = $contractReferences.Add([string]$row.reference)
        }
    }
    foreach ($row in @($previousContractRows)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.reference)) {
            $null = $contractReferences.Add([string]$row.reference)
        }
    }

    foreach ($reference in @($contractReferences)) {
        $currentRow = Find-SnapshotContractRow -Rows $currentContractRows -Reference $reference
        $previousRow = Find-SnapshotContractRow -Rows $previousContractRows -Reference $reference
        if ($null -eq $currentRow -and $null -eq $previousRow) {
            continue
        }

        $changeKind = 'stable'
        $changedFields = @()
        $summary = ''
        $reason = ''
        $nextStep = ''
        $severityScore = 0
        $hostRow = if ($currentRow) { $currentRow } else { $previousRow }
        $hrefReference = [string]$(if ($hostRow -and [string]$hostRow.reference) { $hostRow.reference } else { $reference })

        if ($currentRow -and $previousRow) {
            $changedFields = @(
                @(Get-SnapshotContractVersionChanges -CurrentRow $currentRow -PreviousRow $previousRow) |
                Sort-Object @{ Expression = { [int]$_.severity }; Descending = $true }, @{ Expression = { [string]$_.label }; Descending = $false }
            )
            if (@($changedFields).Count -eq 0) {
                continue
            }

            $changeKind = 'updated'
            $severityScore = [int](@($changedFields | Measure-Object -Property severity -Sum).Sum)
            $topChanges = @($changedFields | Select-Object -First 3)
            $summary = [string]::Join(' | ', @(
                $topChanges |
                ForEach-Object { "$([string]$_.label): $([string]$_.previous) -> $([string]$_.current)" }
            ))
            $reason = "O contrato mudou em $([int]@($changedFields).Count) campo(s) desde a base anterior."
            $nextStep = if (@($changedFields | Where-Object { [int]$_.severity -ge 35 }).Count -gt 0) {
                'Revisar o dossie e confirmar se a mudanca exige acao operacional.'
            }
            else {
                'Registrar a alteracao no acompanhamento e manter o monitoramento.'
            }
        }
        elseif ($currentRow) {
            $changeKind = 'new'
            $severityScore = [int](
                20 +
                ([int]$currentRow.alertCount * 6) +
                ([int]$currentRow.divergenceCount * 5) +
                $(if ([bool]$currentRow.reviewPending) { 25 } else { 0 })
            )
            $summary = [string]$(if ([string]$currentRow.portalStatus) {
                "Contrato entrou na carteira atual com status $([string]$currentRow.portalStatus)."
            }
            else {
                'Contrato entrou na carteira atual da base consolidada.'
            })
            $reason = 'O contrato nao existia no snapshot anterior.'
            $nextStep = 'Abrir o dossie e classificar o novo contrato no acompanhamento gerencial.'
        }
        else {
            $changeKind = 'removed'
            $severityScore = [int](
                15 +
                ([int]$previousRow.alertCount * 6) +
                ([int]$previousRow.divergenceCount * 5) +
                $(if ([bool]$previousRow.reviewPending) { 20 } else { 0 })
            )
            $summary = 'Contrato deixou de aparecer na carteira atual da base consolidada.'
            $reason = 'O contrato constava no snapshot anterior e nao apareceu no recorte atual.'
            $nextStep = 'Conferir se a saida decorre de encerramento, reclassificacao ou falha de carga.'
        }

        $contractChangeEntries.Add([pscustomobject][ordered]@{
            reference = $hrefReference
            title = [string]$(if ($hostRow -and [string]$hostRow.title) { $hostRow.title } elseif ($hostRow -and [string]$hostRow.contractNumber) { $hostRow.contractNumber } else { $reference })
            organization = [string]$(if ($hostRow) { $hostRow.organization } else { '' })
            contractor = [string]$(if ($hostRow) { $hostRow.contractor } else { '' })
            changeKind = $changeKind
            changeCount = [int]@($changedFields).Count
            severityScore = [int]$severityScore
            summary = Collapse-Whitespace -Text $summary
            reason = Collapse-Whitespace -Text $reason
            nextStep = Collapse-Whitespace -Text $nextStep
            changedFields = @($changedFields)
            changedLabels = @($changedFields | ForEach-Object { [string]$_.label })
            publishedAt = [string]$(if ($hostRow) { $hostRow.publishedAt } else { '' })
            href = if ([string]::IsNullOrWhiteSpace($hrefReference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($hrefReference))" }
        }) | Out-Null
    }

    $topContractChanges = @(
        @($contractChangeEntries.ToArray()) |
        Sort-Object @{ Expression = { [int]$_.severityScore }; Descending = $true }, @{ Expression = { [int]$_.changeCount }; Descending = $true }, @{ Expression = { [string]$_.publishedAt }; Descending = $true }, @{ Expression = { [string]$_.reference }; Descending = $false } |
        Select-Object -First 8
    )
    $topSupplierMovers = @(
        @($supplierMovers.ToArray()) |
        Sort-Object @{ Expression = { [Math]::Abs([int]$_.delta) }; Descending = $true }, @{ Expression = { [double]$_.totalValue }; Descending = $true }, @{ Expression = { [string]$_.name }; Descending = $false } |
        Select-Object -First 8
    )

    $detailGroups = @(
        (Get-SnapshotStringDeltaGroup -Key 'review_queue' -Title 'Mudancas na fila de revisao' -Description 'Referencias que entraram ou sairam da fila manual desde a sincronizacao anterior.' -CurrentValues @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'reviewReferences') -PreviousValues @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'reviewReferences'))
        (Get-SnapshotStringDeltaGroup -Key 'critical_alerts' -Title 'Mudancas nos alertas criticos' -Description 'Referencias que passaram a concentrar alertas criticos ou deixaram de concentrar esse risco.' -CurrentValues @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'criticalAlertReferences') -PreviousValues @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'criticalAlertReferences'))
        (Get-SnapshotStringDeltaGroup -Key 'active_contracts' -Title 'Mudancas na carteira vigente' -Description 'Referencias que passaram a constar ou deixaram de constar como vigentes.' -CurrentValues @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'activeReferences') -PreviousValues @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'activeReferences'))
        [pscustomobject][ordered]@{
            key = 'organization_load'
            title = 'Orgaos com maior variacao'
            description = 'Mudancas de volume por orgao entre as duas ultimas recomposicoes.'
            movers = @(
                @($organizationMovers.ToArray()) |
                Sort-Object @{ Expression = { [Math]::Abs([int]$_.delta) }; Descending = $true }, @{ Expression = { [double]$_.totalValue }; Descending = $true }, @{ Expression = { [string]$_.name }; Descending = $false } |
                Select-Object -First 8
            )
        }
        [pscustomobject][ordered]@{
            key = 'supplier_load'
            title = 'Fornecedores com maior variacao'
            description = 'Mudancas de volume entre os fornecedores mais recorrentes da carteira oficial.'
            movers = @($topSupplierMovers)
        }
        [pscustomobject][ordered]@{
            key = 'contract_versions'
            title = 'Contratos com mudanca relevante'
            description = 'Contratos que mudaram entre os dois ultimos snapshots, incluindo entrada, saida ou alteracao de campos-chave.'
            changes = @($topContractChanges)
        }
    )

    $changedContractsCount = [int]@($contractChangeEntries.ToArray()).Count
    $deltaSignalsCount = [int]@($items | Where-Object { [int]$_.delta -ne 0 }).Count
    $headline = if ($previousSnapshot -eq $null) {
        'Primeira comparacao historica registrada para a base consolidada.'
    }
    elseif ($changedContractsCount -gt 0) {
        "$changedContractsCount contrato(s) mudaram desde a ultima sincronizacao consolidada."
    }
    elseif ($deltaSignalsCount -gt 0) {
        'Os totais da base mudaram, mas sem alteracoes contratuais prioritarias no topo da carteira.'
    }
    else {
        'Sem mudancas estruturais relevantes desde a ultima sincronizacao.'
    }
    $windowLabel = if ($previousSnapshot) {
        "$([string]$previousSnapshot.generatedAt) -> $([string]$currentSnapshot.generatedAt)"
    }
    else {
        [string]$currentSnapshot.generatedAt
    }
    $summaryText = if ($previousSnapshot -eq $null) {
        'Use a proxima sincronizacao para comecar a comparar deltas, entradas e saidas de contratos.'
    }
    elseif ($changedContractsCount -gt 0) {
        'O diff gerencial agora destaca contratos que entraram, sairam ou mudaram em campos-chave.'
    }
    else {
        'A carteira permaneceu estavel entre os dois ultimos snapshots registrados.'
    }

    $summary = [ordered]@{
        currentGeneratedAt = [string]$currentSnapshot.generatedAt
        previousGeneratedAt = if ($previousSnapshot) { [string]$previousSnapshot.generatedAt } else { $null }
        headline = $headline
        windowLabel = $windowLabel
        summaryText = $summaryText
        changedContractsCount = $changedContractsCount
        items = @($items)
        contractVersionChanges = @($topContractChanges)
        supplierMovers = @($topSupplierMovers)
        detailGroups = @($detailGroups)
    }

    if (-not $PSBoundParameters.ContainsKey('Snapshots')) {
        Set-WorkspacePayloadRecentChanges -Payload $workspacePayload -RecentChanges $summary
        Save-WorkspacePayload -Payload $workspacePayload
    }

    return $summary
}

function Get-ContractCrossReviewPayload {
    $payload = Read-JsonFile -Path $script:ContractCrossReviewPath -Default (Get-EmptyContractCrossReviewPayload)
    if ($null -eq $payload) {
        return (Get-EmptyContractCrossReviewPayload)
    }

    if (-not $payload.PSObject.Properties.Match('decisions')) {
        $payload | Add-Member -NotePropertyName decisions -NotePropertyValue @()
    }

    if (-not $payload.PSObject.Properties.Match('version')) {
        $payload | Add-Member -NotePropertyName version -NotePropertyValue $script:ContractCrossReviewSchemaVersion
    }

    $changed = $false
    foreach ($decision in @($payload.decisions)) {
        if (-not $decision.PSObject.Properties.Match('status')) {
            $decision | Add-Member -NotePropertyName status -NotePropertyValue $(if (-not [string]::IsNullOrWhiteSpace([string]$decision.officialPortalContractId)) { 'confirmed' } else { 'no_link' })
            $changed = $true
        }
        if (-not $decision.PSObject.Properties.Match('history')) {
            $decision | Add-Member -NotePropertyName history -NotePropertyValue @()
            $changed = $true
        }
    }

    if ($changed) {
        Save-ContractCrossReviewPayload -Payload $payload
    }

    return $payload
}

function Save-ContractCrossReviewPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $Payload.generatedAt = Get-IsoNow
    $Payload.version = $script:ContractCrossReviewSchemaVersion
    $Payload.decisions = @(
        @($Payload.decisions) |
        Sort-Object `
            @{ Expression = { [string]$_.updatedAt }; Descending = $true }, `
            @{ Expression = { [string]$_.movementKey }; Descending = $false }
    )

    Write-JsonFile -Path $script:ContractCrossReviewPath -Data $Payload
}

function Set-ContractCrossReviewDecision {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MovementKey,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$OfficialPortalContractId,

        [Parameter(Mandatory = $true)]
        [string]$UpdatedBy,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$CrossKey = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Note = ''
    )

    $cleanMovementKey = Collapse-Whitespace -Text $MovementKey
    if ([string]::IsNullOrWhiteSpace($cleanMovementKey)) {
        throw 'Chave da movimentacao nao informada.'
    }

    if ($cleanMovementKey.Length -gt 240) {
        throw 'Chave da movimentacao invalida.'
    }

    $cleanOfficialPortalContractId = Collapse-Whitespace -Text $OfficialPortalContractId
    $cleanCrossKey = Collapse-Whitespace -Text $CrossKey
    $cleanNote = Collapse-Whitespace -Text $Note
    if ($cleanNote.Length -gt 240) {
        $cleanNote = $cleanNote.Substring(0, 240).Trim()
    }

    $payload = Get-ContractCrossReviewPayload
    $existing = @($payload.decisions | Where-Object { [string]$_.movementKey -eq $cleanMovementKey })
    $payload.decisions = @($payload.decisions | Where-Object { [string]$_.movementKey -ne $cleanMovementKey })

    if ([string]::IsNullOrWhiteSpace($cleanOfficialPortalContractId)) {
        Save-ContractCrossReviewPayload -Payload $payload
        return [ordered]@{
            ok = $true
            removed = ($existing.Count -gt 0)
            movementKey = $cleanMovementKey
        }
    }

    $decision = [pscustomobject][ordered]@{
        id = if ($existing.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$existing[0].id)) { [string]$existing[0].id } else { [Guid]::NewGuid().ToString('N') }
        movementKey = $cleanMovementKey
        crossKey = $cleanCrossKey
        officialPortalContractId = $cleanOfficialPortalContractId
        status = 'confirmed'
        note = $cleanNote
        updatedAt = Get-IsoNow
        updatedBy = $UpdatedBy
    }

    $payload.decisions = @($payload.decisions) + @($decision)
    Save-ContractCrossReviewPayload -Payload $payload
    return $decision
}

function Save-ContractCrossReviewDecision {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MovementKey,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$OfficialPortalContractId,

        [Parameter(Mandatory = $true)]
        [string]$UpdatedBy,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$CrossKey = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Note = '',

        [Parameter(Mandatory = $false)]
        [ValidateSet('confirm', 'reject', 'no_link', 'reopen')]
        [string]$Action = 'confirm'
    )

    $cleanMovementKey = Collapse-Whitespace -Text $MovementKey
    if ([string]::IsNullOrWhiteSpace($cleanMovementKey)) {
        throw 'Chave da movimentacao nao informada.'
    }

    $payload = Get-ContractCrossReviewPayload
    $existing = @($payload.decisions | Where-Object { [string]$_.movementKey -eq $cleanMovementKey } | Select-Object -First 1) | Select-Object -First 1
    $history = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($(if ($existing) { $existing.history } else { @() }))) {
        $history.Add($entry)
    }

    if ($Action -eq 'reopen') {
        $workspacePayload = Get-WorkspacePayload
        $payload.decisions = @($payload.decisions | Where-Object { [string]$_.movementKey -ne $cleanMovementKey })
        Add-WorkspaceActivity -Payload $workspacePayload -Type 'cross_review_reopen' -Title 'Revisao reaberta' -Summary "Vinculo $cleanMovementKey voltou para a fila." -Reference (Collapse-Whitespace -Text $CrossKey) -CreatedBy $UpdatedBy | Out-Null
        Save-ContractCrossReviewPayload -Payload $payload
        Save-WorkspacePayload -Payload $workspacePayload
        Register-ObservabilityEvent -Type 'cross_review' -Status $Action -Message "Revisao reaberta para $cleanMovementKey." -UserLogin $UpdatedBy -Metadata ([ordered]@{
            movementKey = $cleanMovementKey
            crossKey = (Collapse-Whitespace -Text $CrossKey)
        }) | Out-Null
        return [ordered]@{
            movementKey = $cleanMovementKey
            action = $Action
            removed = $true
        }
    }

    $decisionStatus = switch ($Action) {
        'reject' { 'rejected' }
        'no_link' { 'no_link' }
        default { 'confirmed' }
    }

    if ($decisionStatus -eq 'confirmed' -and [string]::IsNullOrWhiteSpace([string]$OfficialPortalContractId)) {
        throw 'Informe o cadastro oficial para confirmar o vinculo.'
    }

    $cleanCrossKey = Collapse-Whitespace -Text $CrossKey
    $cleanNote = Collapse-Whitespace -Text $Note
    if ($cleanNote.Length -gt 240) {
        $cleanNote = $cleanNote.Substring(0, 240).Trim()
    }

    $history.Add([pscustomobject][ordered]@{
        id = [Guid]::NewGuid().ToString('N')
        action = $Action
        status = $decisionStatus
        updatedAt = Get-IsoNow
        updatedBy = $UpdatedBy
        note = $cleanNote
        officialPortalContractId = Collapse-Whitespace -Text $OfficialPortalContractId
    })

    $decision = [pscustomobject][ordered]@{
        id = if ($existing -and -not [string]::IsNullOrWhiteSpace([string]$existing.id)) { [string]$existing.id } else { [Guid]::NewGuid().ToString('N') }
        movementKey = $cleanMovementKey
        crossKey = $cleanCrossKey
        officialPortalContractId = if ($decisionStatus -eq 'confirmed') { (Collapse-Whitespace -Text $OfficialPortalContractId) } else { '' }
        status = $decisionStatus
        note = $cleanNote
        updatedAt = Get-IsoNow
        updatedBy = $UpdatedBy
        history = @($history | Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true })
    }

    $payload.decisions = @($payload.decisions | Where-Object { [string]$_.movementKey -ne $cleanMovementKey }) + @($decision)
    Save-ContractCrossReviewPayload -Payload $payload
    $workspacePayload = Get-WorkspacePayload
    Add-WorkspaceActivity -Payload $workspacePayload -Type 'cross_review' -Title 'Revisao de vinculo' -Summary "Acao $Action registrada para $cleanMovementKey." -Reference $cleanCrossKey -CreatedBy $UpdatedBy | Out-Null
    Save-WorkspacePayload -Payload $workspacePayload
    Register-ObservabilityEvent -Type 'cross_review' -Status $Action -Message "Acao $Action registrada para $cleanMovementKey." -UserLogin $UpdatedBy -Metadata ([ordered]@{
        movementKey = $cleanMovementKey
        crossKey = $cleanCrossKey
        officialPortalContractId = (Collapse-Whitespace -Text $OfficialPortalContractId)
    }) | Out-Null
    return $decision
}

function Initialize-AppStorage {
    Ensure-Directory -Path $script:StorageRoot
    Ensure-Directory -Path $script:PdfRoot
    Ensure-Directory -Path $script:AnalysisRoot
    Ensure-Directory -Path $script:PersonnelAnalysisRoot
    Ensure-Directory -Path $script:StateRoot

    if (-not (Test-Path -LiteralPath $script:DiariesPath)) {
        Write-JsonFile -Path $script:DiariesPath -Data (Get-EmptyDiariesPayload)
    }

    if (-not (Test-Path -LiteralPath $script:ContractsPath)) {
        Write-JsonFile -Path $script:ContractsPath -Data (Get-EmptyContractsPayload)
    }

    if (-not (Test-Path -LiteralPath $script:PortalContractsPath)) {
        Write-JsonFile -Path $script:PortalContractsPath -Data (Get-EmptyPortalContractsPayload)
    }

    if (-not (Test-Path -LiteralPath $script:StatusPath)) {
        Write-JsonFile -Path $script:StatusPath -Data (Get-DefaultStatus)
    }

    if (-not (Test-Path -LiteralPath $script:UsersPath)) {
        Write-JsonFile -Path $script:UsersPath -Data (Get-EmptyUsersPayload)
    }

    if (-not (Test-Path -LiteralPath $script:SupportPath)) {
        Write-JsonFile -Path $script:SupportPath -Data (Get-EmptySupportPayload)
    }

    if (-not (Test-Path -LiteralPath $script:ContractCrossReviewPath)) {
        Write-JsonFile -Path $script:ContractCrossReviewPath -Data (Get-EmptyContractCrossReviewPayload)
    }

    if (-not (Test-Path -LiteralPath $script:WorkspaceStatePath)) {
        Write-JsonFile -Path $script:WorkspaceStatePath -Data (Get-EmptyWorkspacePayload)
    }

    if (-not (Test-Path -LiteralPath $script:ObservabilityPath)) {
        Write-JsonFile -Path $script:ObservabilityPath -Data (Get-EmptyObservabilityPayload)
    }

    Ensure-DefaultAdminUser
    Ensure-UsersSecurityState
}

function Get-PortalContractsPayload {
    Read-JsonFile -Path $script:PortalContractsPath -Default (Get-EmptyPortalContractsPayload)
}

function Save-PortalContractsPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    Write-JsonFile -Path $script:PortalContractsPath -Data $Payload
}

function Get-AbsolutePortalUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathOrUrl
    )

    if ($PathOrUrl -match '^https?://') {
        return $PathOrUrl
    }

    return ([Uri]::new($script:BasePortalUri, $PathOrUrl)).AbsoluteUri
}

function Invoke-PortalRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathOrUrl,

        [Parameter(Mandatory = $false)]
        [string]$OutFile = $null
    )

    $headers = @{
        'User-Agent' = $script:UserAgent
        'Accept-Language' = 'pt-BR,pt;q=0.9,en;q=0.8'
        'Cache-Control' = 'no-cache'
    }

    $uri = Get-AbsolutePortalUrl -PathOrUrl $PathOrUrl

    if ($OutFile) {
        Invoke-WebRequest -UseBasicParsing -Uri $uri -Headers $headers -TimeoutSec 120 -OutFile $OutFile | Out-Null
        return $OutFile
    }

    return (Invoke-WebRequest -UseBasicParsing -Uri $uri -Headers $headers -TimeoutSec 120).Content
}

function HtmlDecode-Safe {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return $null
    }

    return [System.Web.HttpUtility]::HtmlDecode($Text).Trim()
}

function Get-FirstRegexValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $match = [regex]::Match(
        $Text,
        $Pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if (-not $match.Success) {
        return $null
    }

    if ($match.Groups['value'].Success) {
        return HtmlDecode-Safe -Text $match.Groups['value'].Value
    }

    if ($match.Groups.Count -gt 1) {
        return HtmlDecode-Safe -Text $match.Groups[1].Value
    }

    return HtmlDecode-Safe -Text $match.Value
}

function Get-RegexBlock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $match = [regex]::Match(
        $Text,
        $Pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if (-not $match.Success) {
        return ''
    }

    if ($match.Groups['value'].Success) {
        return [string]$match.Groups['value'].Value
    }

    if ($match.Groups.Count -gt 1) {
        return [string]$match.Groups[1].Value
    }

    return [string]$match.Value
}

function Get-AllRegexValues {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $matches = [regex]::Matches(
        $Text,
        $Pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    $values = @()
    foreach ($match in @($matches)) {
        $rawValue = if ($match.Groups['value'].Success) {
            [string]$match.Groups['value'].Value
        }
        elseif ($match.Groups.Count -gt 1) {
            [string]$match.Groups[1].Value
        }
        else {
            [string]$match.Value
        }

        $cleanValue = HtmlDecode-Safe -Text $rawValue
        if (-not [string]::IsNullOrWhiteSpace($cleanValue)) {
            $values += $cleanValue
        }
    }

    return @($values)
}

function Get-UniqueTextList {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$Values = @()
    )

    $items = New-Object 'System.Collections.Generic.List[string]'
    $seen = @{}
    foreach ($value in @($Values)) {
        $cleanValue = Collapse-Whitespace -Text ([string]$value)
        if ([string]::IsNullOrWhiteSpace($cleanValue)) {
            continue
        }

        $key = Normalize-IndexText -Text $cleanValue
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        $items.Add($cleanValue)
    }

    return @($items)
}

function Get-PortalComboMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html,

        [Parameter(Mandatory = $true)]
        [string]$BaseControlId,

        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $escapedBaseId = [regex]::Escape($BaseControlId)
    $optionsBlock = Get-RegexBlock `
        -Text $Html `
        -Pattern ('id="' + $escapedBaseId + '_DDD_L_LBT"[^>]*>(?<value>.*?)</table><div id="' + $escapedBaseId + '_DDD_L_BS"')

    return [pscustomobject][ordered]@{
        key = $Key
        label = $Label
        controlId = $BaseControlId
        selected = (Collapse-Whitespace -Text (Get-FirstRegexValue -Text $Html -Pattern ('id="' + $escapedBaseId + '_I"[^>]*value="(?<value>[^"]*)"')))
        selectedValue = (Collapse-Whitespace -Text (Get-FirstRegexValue -Text $Html -Pattern ('id="' + $escapedBaseId + '_VI"[^>]*value="(?<value>[^"]*)"')))
        options = @(
            Get-UniqueTextList -Values (
                Get-AllRegexValues -Text $optionsBlock -Pattern '<td class="dxeListBoxItem">(?<value>.*?)</td>'
            )
        )
    }
}

function Get-PortalGridHeaders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html,

        [Parameter(Mandatory = $true)]
        [string]$GridId
    )

    $escapedGridId = [regex]::Escape($GridId)
    return @(
        Get-UniqueTextList -Values (
            Get-AllRegexValues -Text $Html -Pattern ('id="' + $escapedGridId + '_col\d+"[^>]*>.*?<td class="dx-wrap">(?<value>.*?)</td>')
        )
    )
}

function Get-LegacyExpensePortalCapabilityProfile {
    $gridId = 'cphConteudo_ASPxRoundPanel1_gvwDespesa'

    return [pscustomobject][ordered]@{
        comboDefinitions = @(
            [pscustomobject][ordered]@{
                key = 'exercise'
                label = 'Exercicio'
                controlId = 'cphConteudo_rpnCadastro_cbxExercicio'
                defaultSelected = ''
                defaultOptions = @()
                note = 'O exercicio define o recorte anual da despesa consultada.'
            },
            [pscustomobject][ordered]@{
                key = 'budgetType'
                label = 'Tipo'
                controlId = 'cphConteudo_rpnCadastro_cbxTipo'
                defaultSelected = 'ORCAMENTARIO'
                defaultOptions = @('ORCAMENTARIO', 'EXTRAORCAMENTARIA', 'RESTOS A PAGAR')
                note = 'A leitura contratual costuma comecar por ORCAMENTARIO.'
            },
            [pscustomobject][ordered]@{
                key = 'expenseType'
                label = 'Tipo de Despesa'
                controlId = 'cphConteudo_rpnCadastro_cbxDespesa'
                defaultSelected = 'TODOS'
                defaultOptions = @('EMPENHADA', 'LIQUIDADA', 'PAGA', 'TODOS')
                note = 'TODOS permite comparar empenho, liquidacao e pagamento na mesma trilha.'
            },
            [pscustomobject][ordered]@{
                key = 'filterBy'
                label = 'Filtrar por'
                controlId = 'cphConteudo_rpnCadastro_cbxFiltro'
                defaultSelected = 'TODOS'
                defaultOptions = @('TODOS', 'ORGAO', 'SUB-FUNCOES', 'PROGRAMA', 'ACAO', 'CATEGORIA', 'FONTE DE RECURSOS', 'NUMERO DA FICHA DE DESPESA')
                note = 'O filtro inicial mais seguro e TODOS; depois a grade pode ser refinada.'
            }
        )
        queryFields = @(
            [pscustomobject][ordered]@{
                key = 'entity_document'
                label = 'Entidade (CNPJ)'
                controlId = 'cphConteudo_rpnCadastro_lksEntidade_txtValueField'
                controlName = 'ctl00$cphConteudo$rpnCadastro$lksEntidade$txtValueField'
                note = 'Confere a entidade pagadora antes da pesquisa.'
            },
            [pscustomobject][ordered]@{
                key = 'entity_name'
                label = 'Entidade (nome)'
                controlId = 'cphConteudo_rpnCadastro_lksEntidade_txtTextField'
                controlName = 'ctl00$cphConteudo$rpnCadastro$lksEntidade$txtTextField'
                note = 'Campo de apoio para validar que a consulta esta na entidade correta.'
            },
            [pscustomobject][ordered]@{
                key = 'period_start'
                label = 'Periodo inicial'
                controlId = 'cphConteudo_rpnCadastro_detInicial'
                controlName = 'ctl00$cphConteudo$rpnCadastro$detInicial'
                note = 'A janela temporal deve cobrir a vigencia e os pagamentos esperados.'
            },
            [pscustomobject][ordered]@{
                key = 'period_end'
                label = 'Periodo final'
                controlId = 'cphConteudo_rpnCadastro_detFinal'
                controlName = 'ctl00$cphConteudo$rpnCadastro$detFinal'
                note = 'Use o final do exercicio ou a ultima movimentacao conhecida.'
            },
            [pscustomobject][ordered]@{
                key = 'supplier_document'
                label = 'Credor / Fornecedor (CPF-CNPJ)'
                controlId = 'cphConteudo_rpnCadastro_lksFornecedor_txtValueField'
                controlName = 'ctl00$cphConteudo$rpnCadastro$lksFornecedor$txtValueField'
                note = 'Melhor chave inicial quando o contrato traz CNPJ consolidado.'
            },
            [pscustomobject][ordered]@{
                key = 'supplier_name'
                label = 'Credor / Fornecedor (nome)'
                controlId = 'cphConteudo_rpnCadastro_lksFornecedor_txtTextField'
                controlName = 'ctl00$cphConteudo$rpnCadastro$lksFornecedor$txtTextField'
                note = 'Serve como plano B quando o CNPJ nao estiver estruturado.'
            }
        )
        grid = [pscustomobject][ordered]@{
            id = $gridId
            label = 'Resultado de despesa'
            columns = @('Empenho', 'CPF/CNPJ', 'Credor/Fornecedor', 'Mod. Lic.', 'Licitacao', 'Empenhado (R$)', 'Liquidado (R$)', 'Pago (R$)', 'Anulado (R$)')
            filterableColumns = @(
                [pscustomobject][ordered]@{
                    key = 'empenho'
                    label = 'Empenho'
                    controlId = 'cphConteudo_ASPxRoundPanel1_gvwDespesa_DXFREditorcol3'
                    controlName = 'ctl00$cphConteudo$ASPxRoundPanel1$gvwDespesa$DXFREditorcol3'
                    note = 'Refina a grade quando o numero do empenho ja aparece em ato ou documento complementar.'
                },
                [pscustomobject][ordered]@{
                    key = 'cpf_cnpj'
                    label = 'CPF/CNPJ'
                    controlId = 'cphConteudo_ASPxRoundPanel1_gvwDespesa_DXFREditorcol8'
                    controlName = 'ctl00$cphConteudo$ASPxRoundPanel1$gvwDespesa$DXFREditorcol8'
                    note = 'Permite confirmar o favorecido financeiro na propria grade.'
                },
                [pscustomobject][ordered]@{
                    key = 'creditor'
                    label = 'Credor/Fornecedor'
                    controlId = 'cphConteudo_ASPxRoundPanel1_gvwDespesa_DXFREditorcol9'
                    controlName = 'ctl00$cphConteudo$ASPxRoundPanel1$gvwDespesa$DXFREditorcol9'
                    note = 'Filtro textual para aproximar a pesquisa quando o CNPJ nao fecha sozinho.'
                },
                [pscustomobject][ordered]@{
                    key = 'modality'
                    label = 'Mod. Lic.'
                    controlId = 'cphConteudo_ASPxRoundPanel1_gvwDespesa_DXFREditorcol10'
                    controlName = 'ctl00$cphConteudo$ASPxRoundPanel1$gvwDespesa$DXFREditorcol10'
                    note = 'Ajuda a confirmar a modalidade licitatoria vinculada ao empenho.'
                },
                [pscustomobject][ordered]@{
                    key = 'bidding'
                    label = 'Licitacao'
                    controlId = 'cphConteudo_ASPxRoundPanel1_gvwDespesa_DXFREditorcol11'
                    controlName = 'ctl00$cphConteudo$ASPxRoundPanel1$gvwDespesa$DXFREditorcol11'
                    note = 'Campo util quando o cadastro do contrato cita o numero da licitacao.'
                }
            )
        }
        detailSections = @(
            [pscustomobject][ordered]@{
                key = 'identificacao'
                label = 'Identificacao do empenho'
                controlId = 'cphConteudo_pcDetalhes_gvwLinha1'
                fields = @('Empenho/Ano', 'Credor/Fornecedor', 'C.N.P.J./C.P.F.')
                note = 'Confirma o empenho, o favorecido e o documento fiscal.'
                callbackRequired = $true
            },
            [pscustomobject][ordered]@{
                key = 'processo'
                label = 'Processo e emissao'
                controlId = 'cphConteudo_pcDetalhes_gvwLinha2'
                fields = @('Processo', 'Unidade Executora', 'Data de Emissao')
                note = 'Conecta o empenho ao processo administrativo e a unidade executora.'
                callbackRequired = $true
            },
            [pscustomobject][ordered]@{
                key = 'dotacao'
                label = 'Dotacao e contrato'
                controlId = 'cphConteudo_pcDetalhes_gvwLinha12'
                fields = @('Categoria Economica', 'Destinacao de Recursos', 'Contrato')
                note = 'O campo Contrato aparece no detalhe oficial e e o principal ponto de confirmacao com o dossie.'
                callbackRequired = $true
            },
            [pscustomobject][ordered]@{
                key = 'valores'
                label = 'Totais oficiais'
                controlId = 'cphConteudo_pcDetalhes_gvwValores'
                fields = @('Empenhado', 'Em Liquidacao', 'Liquidado', 'Pago', 'Anulado')
                note = 'Resume a execucao financeira acumulada do empenho.'
                callbackRequired = $true
            },
            [pscustomobject][ordered]@{
                key = 'historico'
                label = 'Historico do empenho'
                controlId = 'cphConteudo_pcDetalhes_gvwHistoricoEmpenho'
                fields = @('Item', 'Quantidade', 'Unidade', 'Valor Unitario', 'Total')
                note = 'Mostra itens e composicao do empenho quando o portal devolve linhas.'
                callbackRequired = $true
            },
            [pscustomobject][ordered]@{
                key = 'liquidacoes'
                label = 'Liquidacoes'
                controlId = 'cphConteudo_pcDetalhes_gvwLiquidacoes'
                fields = @('Data da Liquidacao', 'Data do Vencimento', 'Numero da Liquidacao', 'Complemento Historico', 'Valor Liquidado', 'Valor Estornado')
                note = 'Permite auditar a etapa de liquidacao do empenho.'
                callbackRequired = $true
            },
            [pscustomobject][ordered]@{
                key = 'pagamentos'
                label = 'Pagamentos'
                controlId = 'cphConteudo_pcDetalhes_gvwPagamentos'
                fields = @('Data do Pagamento', 'Numero do Pagamento', 'Numero de Liquidacao', 'Complemento Historico', 'Valor Pago', 'Valor Estornado')
                note = 'Permite auditar a etapa de pagamento do empenho.'
                callbackRequired = $true
            }
        )
        executionStages = @(
            [pscustomobject][ordered]@{
                key = 'empenho'
                label = 'Empenho'
                status = 'supported'
                statusLabel = 'Confirmavel no portal'
                note = 'A grade principal e o popup oficial permitem localizar o empenho e validar processo, contrato e totais.'
                evidenceFields = @('Empenho/Ano', 'Processo', 'Contrato', 'Empenhado')
            },
            [pscustomobject][ordered]@{
                key = 'liquidacao'
                label = 'Liquidacao'
                status = 'supported'
                statusLabel = 'Confirmavel no portal'
                note = 'A grade de liquidacoes do popup mostra data, numero, historico e valor liquidado.'
                evidenceFields = @('Data da Liquidacao', 'Numero da Liquidacao', 'Valor Liquidado', 'Valor Estornado')
            },
            [pscustomobject][ordered]@{
                key = 'pagamento'
                label = 'Pagamento'
                status = 'supported'
                statusLabel = 'Confirmavel no portal'
                note = 'A grade de pagamentos do popup mostra data, numero do pagamento, numero da liquidacao e valor pago.'
                evidenceFields = @('Data do Pagamento', 'Numero do Pagamento', 'Numero de Liquidacao', 'Valor Pago')
            }
        )
        limitations = @(
            'A consulta abre em formulario legado ASP.NET com postback e estado de viewstate.',
            'O portal apresenta CAPTCHA na abertura da tela, o que impede uma extracao cega estavel.',
            'Liquidações e pagamentos aparecem dentro do popup de detalhes do empenho e dependem de callbacks do componente legado.',
            'Nao ha filtro principal direto por numero do contrato na tela inicial; a confirmacao do contrato ocorre no detalhe oficial do empenho.'
        )
    }
}

function Get-LegacyExpensePortalMetadata {
    $cacheKey = 'legacy-expense-portal'
    $cached = $script:FinancialPortalMetadataCache[$cacheKey]
    if ($cached -and $cached.expiresAtTicks -gt (Get-Date).Ticks) {
        Add-CacheMetric -CacheName 'financialPortal' -Metric 'hits'
        return $cached.value
    }

    Add-CacheMetric -CacheName 'financialPortal' -Metric 'misses'

    $profile = Get-LegacyExpensePortalCapabilityProfile
    $portalUrl = ([Uri]::new($script:LegacyTransparencyPortalUri, $script:LegacyTransparencyExpensePath)).AbsoluteUri
    $accessedAt = $null
    $html = ''
    $status = 'reference'
    $statusLabel = 'Perfil conhecido'
    $accessNote = 'Usando o perfil conhecido do portal legado de despesa.'
    $accessError = ''
    $requiresCaptcha = $true
    $requiresPostback = $true
    $requiresCallback = $true
    $popupDetails = $true
    $searchButtonUniqueId = 'ctl00$cphConteudo$btnPesquisar'
    $formAction = $script:LegacyTransparencyExpensePath

    try {
        $html = Invoke-PortalRequest -PathOrUrl $portalUrl
        $accessedAt = Get-IsoNow
        $status = 'available'
        $statusLabel = 'Portal acessivel'
        $accessNote = 'A tela oficial de despesa foi acessada e o portal segue expondo empenho, liquidacao e pagamento no detalhe do empenho.'
        $requiresCaptcha = ($html -match 'popupcaptcha' -or $html -match 'VALIDA..O CAPTCHA')
        $requiresPostback = ($html -match '__VIEWSTATE' -or $html -match '__doPostBack')
        $requiresCallback = ($html -match 'WebForm_DoCallback')
        $popupDetails = ($html -match 'pcDetalhes')
        $searchButtonUniqueId = Collapse-Whitespace -Text (Get-FirstRegexValue -Text $html -Pattern 'name="(?<value>ctl00\$cphConteudo\$btnPesquisar)"')
        if ([string]::IsNullOrWhiteSpace($searchButtonUniqueId)) {
            $searchButtonUniqueId = 'ctl00$cphConteudo$btnPesquisar'
        }

        $resolvedFormAction = Collapse-Whitespace -Text (Get-FirstRegexValue -Text $html -Pattern '<form[^>]*action="(?<value>[^"]*wfDespesa\.aspx[^"]*)"')
        if (-not [string]::IsNullOrWhiteSpace($resolvedFormAction)) {
            $formAction = $resolvedFormAction
        }
    }
    catch {
        $status = 'unavailable'
        $statusLabel = 'Portal indisponivel'
        $accessError = Collapse-Whitespace -Text ($_.Exception.Message)
        $accessNote = if ([string]::IsNullOrWhiteSpace($accessError)) {
            'Nao foi possivel acessar a tela oficial de despesa nesta execucao.'
        }
        else {
            "Nao foi possivel acessar a tela oficial de despesa nesta execucao: $accessError"
        }
    }

    $comboLookup = @{}
    foreach ($definition in @($profile.comboDefinitions)) {
        $parsedCombo = if (-not [string]::IsNullOrWhiteSpace($html)) {
            Get-PortalComboMetadata -Html $html -BaseControlId ([string]$definition.controlId) -Key ([string]$definition.key) -Label ([string]$definition.label)
        }
        else {
            $null
        }

        $comboLookup[[string]$definition.key] = [pscustomobject][ordered]@{
            key = [string]$definition.key
            label = [string]$definition.label
            controlId = [string]$definition.controlId
            selected = if (
                $parsedCombo -and
                -not [string]::IsNullOrWhiteSpace([string]$parsedCombo.selected) -and
                [string]$parsedCombo.selected -notmatch 'Ã'
            ) {
                [string]$parsedCombo.selected
            }
            else {
                [string]$definition.defaultSelected
            }
            options = if ($parsedCombo -and @($parsedCombo.options).Count -gt 0) { @($parsedCombo.options) } else { @($definition.defaultOptions) }
            note = [string]$definition.note
        }
    }

    $defaultFilters = [ordered]@{
        exercise = [string]$comboLookup['exercise'].selected
        entityDocument = if (-not [string]::IsNullOrWhiteSpace($html)) {
            Collapse-Whitespace -Text (Get-FirstRegexValue -Text $html -Pattern 'name="cphConteudo_rpnCadastro_lksEntidade_txtValueField_Raw"[^>]*value="(?<value>[^"]*)"')
        }
        else {
            ''
        }
        entityName = if (-not [string]::IsNullOrWhiteSpace($html)) {
            Collapse-Whitespace -Text (Get-FirstRegexValue -Text $html -Pattern 'name="ctl00\$cphConteudo\$rpnCadastro\$lksEntidade\$txtTextField"[^>]*value="(?<value>[^"]*)"')
        }
        else {
            ''
        }
        periodStart = if (-not [string]::IsNullOrWhiteSpace($html)) {
            Collapse-Whitespace -Text (Get-FirstRegexValue -Text $html -Pattern 'id="cphConteudo_rpnCadastro_detInicial_I"[^>]*value="(?<value>[^"]*)"')
        }
        else {
            ''
        }
        periodEnd = if (-not [string]::IsNullOrWhiteSpace($html)) {
            Collapse-Whitespace -Text (Get-FirstRegexValue -Text $html -Pattern 'id="cphConteudo_rpnCadastro_detFinal_I"[^>]*value="(?<value>[^"]*)"')
        }
        else {
            ''
        }
        budgetType = [string]$comboLookup['budgetType'].selected
        expenseType = [string]$comboLookup['expenseType'].selected
        filterBy = [string]$comboLookup['filterBy'].selected
    }

    $queryFields = @(
        $profile.queryFields |
        ForEach-Object {
            [pscustomobject][ordered]@{
                key = [string]$_.key
                label = [string]$_.label
                controlId = [string]$_.controlId
                controlName = [string]$_.controlName
                available = if (-not [string]::IsNullOrWhiteSpace($html)) {
                    ($html -match [regex]::Escape([string]$_.controlId) -or $html -match [regex]::Escape([string]$_.controlName))
                }
                else {
                    $true
                }
                note = [string]$_.note
            }
        }
    )

    $gridColumns = if (-not [string]::IsNullOrWhiteSpace($html)) {
        @(Get-PortalGridHeaders -Html $html -GridId ([string]$profile.grid.id))
    }
    else {
        @()
    }
    if (@($gridColumns).Count -eq 0) {
        $gridColumns = @($profile.grid.columns)
    }

    $filterableColumns = @(
        $profile.grid.filterableColumns |
        ForEach-Object {
            [pscustomobject][ordered]@{
                key = [string]$_.key
                label = [string]$_.label
                controlId = [string]$_.controlId
                controlName = [string]$_.controlName
                available = if (-not [string]::IsNullOrWhiteSpace($html)) {
                    ($html -match [regex]::Escape([string]$_.controlId) -or $html -match [regex]::Escape([string]$_.controlName))
                }
                else {
                    $true
                }
                note = [string]$_.note
            }
        }
    )

    $detailSections = @(
        $profile.detailSections |
        ForEach-Object {
            $fields = if (-not [string]::IsNullOrWhiteSpace($html)) {
                @(Get-PortalGridHeaders -Html $html -GridId ([string]$_.controlId))
            }
            else {
                @()
            }

            if (@($fields).Count -eq 0) {
                $fields = @($_.fields)
            }

            [pscustomobject][ordered]@{
                key = [string]$_.key
                label = [string]$_.label
                controlId = [string]$_.controlId
                available = if (-not [string]::IsNullOrWhiteSpace($html)) {
                    ($html -match [regex]::Escape([string]$_.controlId))
                }
                else {
                    $true
                }
                callbackRequired = [bool]$_.callbackRequired
                fields = @($fields)
                note = [string]$_.note
            }
        }
    )

    $executionStages = @(
        $profile.executionStages |
        ForEach-Object {
            [pscustomobject][ordered]@{
                key = [string]$_.key
                label = [string]$_.label
                status = [string]$_.status
                statusLabel = [string]$_.statusLabel
                available = $true
                note = [string]$_.note
                evidenceFields = @($_.evidenceFields)
            }
        }
    )

    $limitations = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in @($profile.limitations)) {
        $limitations.Add([string]$item)
    }
    if (-not [string]::IsNullOrWhiteSpace($accessError)) {
        $limitations.Add($accessNote)
    }

    $observedSignals = New-Object 'System.Collections.Generic.List[string]'
    if ([string]$status -eq 'available') {
        if ($requiresCaptcha) {
            $observedSignals.Add('CAPTCHA identificado na abertura da tela oficial de despesa.')
        }
        if ($popupDetails) {
            $observedSignals.Add('Popup oficial de detalhes do empenho reconhecido na pagina.')
        }
        if ($requiresCallback) {
            $observedSignals.Add('Liquidacoes e pagamentos continuam dependentes de callback do grid legado.')
        }
        if ($detailSections | Where-Object { [string]$_.key -eq 'dotacao' -and @($_.fields).Count -gt 0 }) {
            $observedSignals.Add('Campo Contrato continua exposto no detalhe oficial do empenho.')
        }
        if ($searchButtonUniqueId) {
            $observedSignals.Add('Botao de pesquisa reconhecido para orientar consulta assistida repetivel.')
        }
    }

    $value = [pscustomobject][ordered]@{
        key = 'portal_transparencia_despesa'
        label = 'Portal Transparencia - Despesa'
        href = $portalUrl
        status = $status
        statusLabel = $statusLabel
        accessNote = $accessNote
        accessedAt = $accessedAt
        formAction = $formAction
        searchButtonUniqueId = $searchButtonUniqueId
        requiresCaptcha = [bool]$requiresCaptcha
        requiresPostback = [bool]$requiresPostback
        requiresCallback = [bool]$requiresCallback
        popupDetails = [bool]$popupDetails
        hasSupplierSearch = $true
        hasCnpjSearch = $true
        hasContractFieldInDetail = $true
        defaultFilters = [pscustomobject]$defaultFilters
        combos = @($comboLookup.Values)
        queryFields = @($queryFields)
        grid = [pscustomobject][ordered]@{
            id = [string]$profile.grid.id
            label = [string]$profile.grid.label
            columns = @($gridColumns)
            filterableColumns = @($filterableColumns)
        }
        detailSections = @($detailSections)
        executionStages = @($executionStages)
        limitations = @($limitations | Select-Object -Unique)
        observedSignals = @($observedSignals | Select-Object -Unique)
        summaryLabel = 'O portal oficial de despesa expõe empenho, liquidacao e pagamento no detalhe do empenho, com confirmacao de contrato e processo dentro do popup oficial.'
    }

    Add-ScriptCacheEntry -Cache $script:FinancialPortalMetadataCache -Key $cacheKey -Value ([ordered]@{
        expiresAtTicks = (Get-Date).AddMinutes(45).Ticks
        value = $value
    }) -MaxEntries 8 -CacheName 'financialPortal'

    return $value
}

function Convert-PortalDateTime {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $normalized = HtmlDecode-Safe -Text $Text
    $normalized = $normalized.Replace('às', 'as').Replace('�s', 'as')
    $normalized = ($normalized -replace '\s+', ' ').Trim()

    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    $styles = [System.Globalization.DateTimeStyles]::AssumeLocal
    $formats = @(
        "dd/MM/yyyy 'as' HH'h'mm",
        "dd/MM/yyyy 'as' HH:mm",
        "dd/MM/yyyy HH'h'mm",
        "dd/MM/yyyy HH:mm",
        "dd/MM/yyyy"
    )

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($normalized, 'dd/MM/yyyy', $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.ToString('s')
    }

    foreach ($format in $formats) {
        if ([DateTime]::TryParseExact($normalized, $format, $culture, $styles, [ref]$parsed)) {
            return $parsed.ToString('s')
        }
    }

    return $normalized
}

function Parse-IntegerLike {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return [Int64]0
    }

    $digits = $Text -replace '\D', ''
    if ([string]::IsNullOrWhiteSpace($digits)) {
        return 0
    }

    return [int]$digits
}

function Convert-BrazilianCurrencyToNumber {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 0
    }

    $normalized = $Text -replace '[^\d,.-]', ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return 0
    }

    $normalized = $normalized.Replace('.', '').Replace(',', '.')

    $value = 0.0
    if ([double]::TryParse($normalized, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
        return [math]::Round($value, 2)
    }

    return 0
}

function Get-ContractKeywords {
    @(
        'contrato',
        'extrato de contrato',
        'aditivo',
        'apostilamento',
        'dispensa',
        'inexigibilidade',
        'registro de preços',
        'rescisão',
        'rescisao',
        'homologação',
        'homologacao'
    )
}

function Get-DiarioPageUrl {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Page,

        [Parameter(Mandatory = $false)]
        [string]$Keyword = '0'
    )

    $keywordSegment = if ([string]::IsNullOrWhiteSpace($Keyword)) { '0' } else { [Uri]::EscapeDataString($Keyword) }
    return "$($script:PortalDiarioPath)/$Page/0/0/$keywordSegment/0"
}

function Parse-DiarioListingPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html
    )

    $totalPagesText = Get-FirstRegexValue -Text $Html -Pattern 'data-total-paginas="(?<value>\d+)"'
    $totalResultsText = Get-FirstRegexValue -Text $Html -Pattern '<span class="sw_qtde_resultados">(?<value>[^<]+)</span>'
    $blocks = [regex]::Matches(
        $Html,
        '<div class="dof_publicacao_diario sw_item_listagem">(?<block>.*?)(?=<div class="dof_publicacao_diario sw_item_listagem">|<div class="sw_area_paginacao"|</div>\s*</div>\s*</div>\s*</div>)',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    $items = New-Object System.Collections.Generic.List[object]

    foreach ($match in $blocks) {
        $block = $match.Groups['block'].Value
        $id = Get-FirstRegexValue -Text $block -Pattern 'href="/portal/diario-oficial/ver/(?<value>\d+)'
        $edition = Get-FirstRegexValue -Text $block -Pattern 'Edi.{0,15}n.? (?<value>\d+)'
        $downloadTokenPath = Get-FirstRegexValue -Text $block -Pattern 'data-href="(?<value>[^"]+)"'

        if ([string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($edition)) {
            continue
        }

        $postedAtRaw = Get-FirstRegexValue -Text $block -Pattern '<strong>Postagem:</strong>\s*<span>(?<value>[^<]+)</span>'
        $sizeLabel = Get-FirstRegexValue -Text $block -Pattern '<strong>Tamanho:</strong>\s*<span>(?<value>[^<]+)</span>'
        $pageCount = Parse-IntegerLike -Text (Get-FirstRegexValue -Text $sizeLabel -Pattern '(?<value>\d+)\s*p')
        $fileSize = Get-FirstRegexValue -Text $sizeLabel -Pattern '(?<value>[\d\.,]+\s*(?:KB|MB|GB))'

        $item = [ordered]@{
            id = [string]$id
            edition = [string]$edition
            isExtra = [bool]($block -match 'dof_edicao_extra')
            viewPath = "/portal/diario-oficial/ver/$id"
            viewUrl = (Get-AbsolutePortalUrl -PathOrUrl "/portal/diario-oficial/ver/$id")
            downloadTokenPath = $downloadTokenPath
            downloadTokenUrl = if ($downloadTokenPath) { Get-AbsolutePortalUrl -PathOrUrl $downloadTokenPath } else { $null }
            postedAtRaw = $postedAtRaw
            publishedAt = (Convert-PortalDateTime -Text $postedAtRaw)
            pageCount = $pageCount
            fileSize = $fileSize
            sizeLabel = $sizeLabel
        }

        $items.Add([pscustomobject]$item)
    }

    return [pscustomobject]@{
        totalPages = [Math]::Max((Parse-IntegerLike -Text $totalPagesText), 1)
        totalResults = (Parse-IntegerLike -Text $totalResultsText)
        items = $items.ToArray()
    }
}

function Get-PdfRedirectUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DownloadTokenPath
    )

    $html = Invoke-PortalRequest -PathOrUrl $DownloadTokenPath
    $relativePdfPath = Get-FirstRegexValue -Text $html -Pattern 'url=(?<value>[^">]+\.pdf)'

    if ([string]::IsNullOrWhiteSpace($relativePdfPath)) {
        throw "Nao foi possivel descobrir o PDF real para $DownloadTokenPath."
    }

    return Get-AbsolutePortalUrl -PathOrUrl $relativePdfPath
}

function Get-LocalPdfInfo {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Diary,

        [Parameter(Mandatory = $true)]
        [string]$PdfUrl
    )

    $publishedYear = if ($Diary.publishedAt -and $Diary.publishedAt.Length -ge 4) { $Diary.publishedAt.Substring(0, 4) } else { 'sem-ano' }
    $fileName = [System.IO.Path]::GetFileName(([Uri]$PdfUrl).LocalPath)

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "diario-$($Diary.id)-edicao-$($Diary.edition).pdf"
    }

    $relativePath = Join-Path (Join-Path 'storage\pdfs' $publishedYear) $fileName
    $absolutePath = Join-Path $script:AppRoot $relativePath

    return [pscustomobject]@{
        relativePath = $relativePath.Replace('\', '/')
        absolutePath = $absolutePath
        webPath = ('/pdfs/' + ($publishedYear + '/' + $fileName).Replace('\', '/'))
    }
}

function Get-LocalPortalPdfInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PortalYear,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $safeYear = if ([string]::IsNullOrWhiteSpace($PortalYear)) { 'sem-ano' } else { ($PortalYear -replace '[^\d]', '') }
    if ([string]::IsNullOrWhiteSpace($safeYear)) {
        $safeYear = 'sem-ano'
    }

    $relativePath = Join-Path (Join-Path 'storage\pdfs\portal-contratos' $safeYear) $FileName
    $absolutePath = Join-Path $script:AppRoot $relativePath

    return [pscustomobject]@{
        relativePath = $relativePath.Replace('\', '/')
        absolutePath = $absolutePath
        webPath = ('/pdfs/' + ('portal-contratos/' + $safeYear + '/' + $FileName).Replace('\', '/'))
    }
}

function Get-DiariesPayload {
    Read-JsonFile -Path $script:DiariesPath -Default (Get-EmptyDiariesPayload)
}

function Save-DiariesPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    Write-JsonFile -Path $script:DiariesPath -Data $Payload
}

function Get-DiariesById {
    $payload = Get-DiariesPayload
    $index = @{}

    foreach ($diary in @($payload.diaries)) {
        $index[[string]$diary.id] = $diary
    }

    return $index
}

function Get-AnalysisPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DiaryId
    )

    return Join-Path $script:AnalysisRoot "$DiaryId.json"
}

function Get-ExistingAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DiaryId
    )

    return Read-JsonFile -Path (Get-AnalysisPath -DiaryId $DiaryId) -Default $null
}

function Test-AnalysisCurrent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Diary,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Analysis
    )

    if ($null -eq $Analysis) {
        return $false
    }

    if ([string]$Analysis.parserVersion -ne $script:ParserVersion) {
        return $false
    }

    if ([string]$Analysis.sourcePdfUrl -ne [string]$Diary.pdfUrl) {
        return $false
    }

    return $true
}

function Get-DiaryAnalysisReason {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Diary
    )

    $keywords = @($Diary.candidateKeywords)
    $labels = @()

    if ($keywords.Count -gt 0) {
        $labels += 'keyword_map'
    }
    else {
        $labels += 'full_catalog'
    }

    return [pscustomobject]@{
        primary = [string]$labels[0]
        labels = @($labels)
    }
}

function Get-PendingAnalysisTasks {
    param(
        [Parameter(Mandatory = $false)]
        [int]$Limit = 3
    )

    $payload = Get-DiariesPayload
    $tasks = New-Object System.Collections.Generic.List[object]

    foreach ($diary in @($payload.diaries | Sort-Object publishedAt -Descending)) {
        $keywords = @($diary.candidateKeywords)

        if ([string]::IsNullOrWhiteSpace([string]$diary.localPdfPath) -or -not (Test-Path -LiteralPath $diary.localPdfPath)) {
            continue
        }

        $analysis = Get-ExistingAnalysis -DiaryId ([string]$diary.id)
        if (Test-AnalysisCurrent -Diary $diary -Analysis $analysis) {
            continue
        }

        $analysisReason = Get-DiaryAnalysisReason -Diary $diary
        $tasks.Add([pscustomobject]@{
            diaryId = [string]$diary.id
            edition = [string]$diary.edition
            publishedAt = [string]$diary.publishedAt
            isExtra = [bool]$diary.isExtra
            candidateKeywords = @($keywords)
            analysisReason = @($analysisReason.labels)
            pdfUrl = [string]$diary.webPdfPath
            sourcePdfUrl = [string]$diary.pdfUrl
            localPdfRelative = [string]$diary.localPdfRelative
            pageCount = [int]$diary.pageCount
        })

        if ($tasks.Count -ge $Limit) {
            break
        }
    }

    return $tasks.ToArray()
}

function Get-DateTimestampValue {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 0
    }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse([string]$Text, [ref]$parsed)) {
        return [Int64]$parsed.Ticks
    }

    return [Int64]0
}

function Get-ObjectPropertyEntries {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Item
    )

    if ($null -eq $Item) {
        return @()
    }

    if ($Item -is [System.Collections.IDictionary]) {
        return @(
            $Item.GetEnumerator() |
            ForEach-Object {
                [pscustomobject][ordered]@{
                    Name = [string]$_.Key
                    Value = $_.Value
                }
            }
        )
    }

    return @(
        $Item.PSObject.Properties |
        ForEach-Object {
            [pscustomobject][ordered]@{
                Name = [string]$_.Name
                Value = $_.Value
            }
        }
    )
}

function Test-ObjectProperty {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return [bool](@(Get-ObjectPropertyEntries -Item $Item | Where-Object { [string]$_.Name -eq $Name } | Select-Object -First 1).Count -gt 0)
}

function Get-ObjectStringValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $property = @(Get-ObjectPropertyEntries -Item $Item | Where-Object { [string]$_.Name -eq $Name } | Select-Object -First 1) | Select-Object -First 1
    if ($null -eq $property -or $null -eq $property.Value) {
        return ''
    }

    return [string]$property.Value
}

function Get-ContractCrossKey {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $values = @(
        (Get-ObjectStringValue -Item $Item -Name 'referenceKey'),
        (Get-ObjectStringValue -Item $Item -Name 'managementProfileKey'),
        (Get-ObjectStringValue -Item $Item -Name 'contractNumber'),
        (Get-ObjectStringValue -Item $Item -Name 'processNumber')
    )

    foreach ($value in $values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $compact = ($value -replace '\s+', '').ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($compact)) {
            continue
        }

        $match = [regex]::Match($compact, '(?<number>\d{1,10})/(?<year>\d{4})')
        if ($match.Success) {
            return ('{0}/{1}' -f ([int]$match.Groups['number'].Value), [string]$match.Groups['year'].Value)
        }

        return $compact
    }

    return ''
}

function Get-ContractCrossItemKey {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $portalContractId = Get-ObjectStringValue -Item $Item -Name 'portalContractId'
    $itemType = if (-not [string]::IsNullOrWhiteSpace($portalContractId)) { 'official' } else { 'movement' }
    $parts = @(
        $itemType,
        $portalContractId,
        (Get-ObjectStringValue -Item $Item -Name 'diaryId'),
        (Get-ContractCrossKey -Item $Item),
        (Get-ObjectStringValue -Item $Item -Name 'publishedAt'),
        (Get-ObjectStringValue -Item $Item -Name 'pageNumber'),
        (Get-ObjectStringValue -Item $Item -Name 'contractNumber'),
        (Get-ObjectStringValue -Item $Item -Name 'processNumber'),
        (Get-ObjectStringValue -Item $Item -Name 'actTitle')
    )

    return ($parts -join '::')
}

function Get-ContractCrossSimpleText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    return Normalize-IndexText -Text (Collapse-Whitespace -Text $Text)
}

function Get-ContractCrossNormalizedCnpj {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $digits = [regex]::Replace(([string]$Text), '\D', '')
    if ($digits.Length -eq 14) {
        return $digits
    }

    return ''
}

function Get-ContractCrossTextVariants {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $normalized = Get-ContractCrossSimpleText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $variantSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $candidates = New-Object 'System.Collections.Generic.List[string]'
    $candidates.Add($normalized)

    $organizationPatterns = @(
        '\bPREFEITURA MUNICIPAL DE\b',
        '\bPREFEITURA DE\b',
        '\bMUNICIPIO DE\b',
        '\bMUNICIPAL DE\b',
        '\bSECRETARIA MUNICIPAL DE\b',
        '\bSECRETARIA DE\b',
        '\bDEPARTAMENTO DE\b',
        '\bDIRETORIA DE\b',
        '\bDIVISAO DE\b',
        '\bSERVICO MUNICIPAL DE\b',
        '\bFUNDO MUNICIPAL DE\b',
        '\bAUTARQUIA MUNICIPAL DE\b',
        '\bGABINETE DO\b'
    )
    $supplierPatterns = @(
        '\bLTDA\b',
        '\bLIMITADA\b',
        '\bEIRELI\b',
        '\bMEI\b',
        '\bME\b',
        '\bEPP\b',
        '\bSPE\b',
        '\bSA\b',
        '\bS/A\b',
        '\bSOCIEDADE ANONIMA\b',
        '\bEMPRESA INDIVIDUAL DE RESPONSABILIDADE LIMITADA\b'
    )

    foreach ($pattern in @($organizationPatterns + $supplierPatterns)) {
        $candidate = Collapse-Whitespace -Text (($normalized -replace $pattern, ' ') -replace '\s+', ' ')
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $candidates.Add((Normalize-IndexText -Text $candidate))
        }
    }

    foreach ($candidate in @($candidates)) {
        $cleanValue = Collapse-Whitespace -Text $candidate
        if ([string]::IsNullOrWhiteSpace($cleanValue)) {
            continue
        }

        $null = $variantSet.Add([string]$cleanValue)
        $reduced = Collapse-Whitespace -Text (($cleanValue -replace '\b(DE|DA|DO|DOS|DAS|E)\b', ' ') -replace '\s+', ' ')
        if (-not [string]::IsNullOrWhiteSpace($reduced)) {
            $null = $variantSet.Add([string]$reduced)
        }
    }

    return @($variantSet)
}

function Get-ContractCrossYear {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    if ($Text -match '^(?<year>\d{4})-') {
        return [string]$matches['year']
    }

    if ($Text -match '(?<year>20\d{2})') {
        return [string]$matches['year']
    }

    return ''
}

function Test-ContractCrossCompatibleText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Left,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Right
    )

    $leftVariants = @(Get-ContractCrossTextVariants -Text $Left)
    $rightVariants = @(Get-ContractCrossTextVariants -Text $Right)

    if (@($leftVariants).Count -eq 0 -or @($rightVariants).Count -eq 0) {
        return $false
    }

    foreach ($leftValue in @($leftVariants)) {
        foreach ($rightValue in @($rightVariants)) {
            if ($leftValue -eq $rightValue) {
                return $true
            }

            if ($leftValue.Contains($rightValue) -or $rightValue.Contains($leftValue)) {
                return $true
            }

            $leftTokens = @($leftValue -split '\s+' | Where-Object { $_ -and $_.Length -ge 4 })
            $rightTokenSet = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($token in @($rightValue -split '\s+' | Where-Object { $_ -and $_.Length -ge 4 })) {
                $null = $rightTokenSet.Add([string]$token)
            }

            $overlap = 0
            foreach ($token in $leftTokens) {
                if ($rightTokenSet.Contains([string]$token)) {
                    $overlap++
                }
            }

            if ($overlap -ge 2) {
                return $true
            }
        }
    }

    return $false
}

function Get-ContractCrossKeywordSet {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $normalized = Get-ContractCrossSimpleText -Text $Text
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @($set)
    }

    $stopWords = @('PARA', 'COM', 'DOS', 'DAS', 'QUE', 'POR', 'SEM', 'NOS', 'NAS', 'UMA', 'UM', 'ENTRE', 'PELA', 'PELO', 'DE', 'DA', 'DO', 'E', 'O', 'A')
    foreach ($token in @($normalized -split '\s+')) {
        if (-not $token -or $token.Length -lt 4) {
            continue
        }

        if ($token -in $stopWords) {
            continue
        }

        $null = $set.Add([string]$token)
    }

    return @($set)
}

function Get-ContractCrossOverlapCount {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$Left,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$Right
    )

    if ($null -eq $Left) { $Left = @() }
    if ($null -eq $Right) { $Right = @() }

    $rightSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($token in @($Right)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
            $null = $rightSet.Add([string]$token)
        }
    }

    $count = 0
    foreach ($token in @($Left)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$token) -and $rightSet.Contains([string]$token)) {
            $count++
        }
    }

    return $count
}

function Get-ContractCrossCandidateEvaluation {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Movement,

        [Parameter(Mandatory = $true)]
        [object]$Official
    )

    $score = 0
    $reasons = New-Object 'System.Collections.Generic.List[string]'
    $movementCrossKey = Get-ContractCrossKey -Item $Movement
    $officialCrossKey = Get-ContractCrossKey -Item $Official
    $sameOrganization = $false
    $sameContractor = $false
    $sameValueBand = $false
    $movementCnpj = Get-ContractCrossNormalizedCnpj -Text ([string]$Movement.cnpj)
    $officialCnpj = Get-ContractCrossNormalizedCnpj -Text ([string]$Official.cnpj)

    if (-not [string]::IsNullOrWhiteSpace($movementCrossKey) -and $movementCrossKey -eq $officialCrossKey) {
        $score += 50
        $reasons.Add("mesma referencia $movementCrossKey")
    }

    if (Test-ContractCrossCompatibleText -Left ([string]$Movement.primaryOrganizationName) -Right ([string]$Official.primaryOrganizationName)) {
        $score += 20
        $sameOrganization = $true
        $reasons.Add('mesmo orgao')
    }

    if (-not [string]::IsNullOrWhiteSpace($movementCnpj) -and $movementCnpj -eq $officialCnpj) {
        $score += 25
        $sameContractor = $true
        $reasons.Add('mesmo cnpj')
    }
    elseif (Test-ContractCrossCompatibleText -Left ([string]$Movement.contractor) -Right ([string]$Official.contractor)) {
        $score += 20
        $sameContractor = $true
        $reasons.Add('mesmo fornecedor')
    }

    $movementValue = if ($null -ne $Movement.valueNumber) { [double]$Movement.valueNumber } else { 0.0 }
    $officialValue = if ($null -ne $Official.valueNumber) { [double]$Official.valueNumber } else { 0.0 }
    if ($movementValue -gt 0 -and $officialValue -gt 0) {
        $tolerance = [Math]::Max(($officialValue * 0.05), 100.0)
        if ([Math]::Abs($movementValue - $officialValue) -le $tolerance) {
            $score += 10
            $sameValueBand = $true
            $reasons.Add('valor compativel')
        }
    }

    $movementYear = Get-ContractCrossYear -Text ([string]$Movement.publishedAt)
    $officialYear = Get-ContractCrossYear -Text ([string]$Official.publishedAt)
    if (-not [string]::IsNullOrWhiteSpace($movementYear) -and $movementYear -eq $officialYear) {
        $score += 5
        $reasons.Add("mesmo ano $movementYear")
    }

    $movementKeywords = @(Get-ContractCrossKeywordSet -Text ([string]$Movement.object))
    $officialKeywords = @(Get-ContractCrossKeywordSet -Text ([string]$Official.object))
    $overlap = Get-ContractCrossOverlapCount -Left $movementKeywords -Right $officialKeywords
    if ($overlap -ge 4) {
        $score += 10
        $reasons.Add('objeto semelhante')
    }
    elseif ($overlap -ge 2) {
        $score += 6
        $reasons.Add('objeto compativel')
    }

    $confidence = if ($score -ge 80) {
        'alta'
    }
    elseif ($score -ge 55) {
        'media'
    }
    else {
        'baixa'
    }

    return [ordered]@{
        score = [int]$score
        confidence = $confidence
        sameOrganization = [bool]$sameOrganization
        sameContractor = [bool]$sameContractor
        sameValueBand = [bool]$sameValueBand
        reason = if ($reasons.Count -gt 0) { ($reasons -join ' | ') } else { 'sem evidencia complementar forte' }
    }
}

function Get-ContractIssueSeverityRank {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Severity
    )

    switch ((Collapse-Whitespace -Text $Severity).ToLowerInvariant()) {
        'critical' { return 3 }
        'warning' { return 2 }
        'info' { return 1 }
        default { return 0 }
    }
}

function Add-ContractCounterValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Counter,

        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    if (-not $Counter.ContainsKey($Key)) {
        $Counter[$Key] = 0
    }

    $Counter[$Key] = [int]$Counter[$Key] + 1
}

function Convert-ContractCounterToSummary {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Counter
    )

    $items = @(
        $Counter.GetEnumerator() |
        Sort-Object -Property @{ Expression = { $_.Value }; Descending = $true }, @{ Expression = { $_.Key }; Descending = $false } |
        ForEach-Object {
            [pscustomobject][ordered]@{
                reason = [string]$_.Key
                count = [int]$_.Value
            }
        }
    )

    return [ordered]@{
        total = [int]@($items | ForEach-Object { [int]$_.count } | Measure-Object -Sum).Sum
        reasons = @($items)
    }
}

function Add-ContractCrossIssue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Store,

        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [string]$Severity,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$CrossKey,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ContractNumber,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$PortalContractId,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Reason,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$PublishedAt,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$SourceView,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$BucketKey = $null,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$Tags = @()
    )

    $resolvedBucketKey = if (-not [string]::IsNullOrWhiteSpace($BucketKey)) {
        $BucketKey
    }
    else {
        (@($Type, $CrossKey, $PortalContractId, $SourceView) -join '|')
    }

    if (-not $Store.ContainsKey($resolvedBucketKey)) {
        $Store[$resolvedBucketKey] = [pscustomobject][ordered]@{
            type = $Type
            severity = $Severity
            crossKey = $CrossKey
            contractNumber = $ContractNumber
            portalContractId = $PortalContractId
            title = $Title
            reason = $Reason
            publishedAt = $PublishedAt
            sourceView = $SourceView
            occurrenceCount = 0
            tags = @()
        }
    }

    $entry = $Store[$resolvedBucketKey]
    if ((Get-ContractIssueSeverityRank -Severity $Severity) -gt (Get-ContractIssueSeverityRank -Severity ([string]$entry.severity))) {
        $entry.severity = $Severity
    }

    $incomingPublishedAt = Get-DateTimestampValue -Text $PublishedAt
    $currentPublishedAt = Get-DateTimestampValue -Text ([string]$entry.publishedAt)
    if ($incomingPublishedAt -ge $currentPublishedAt) {
        $entry.reason = $Reason
        $entry.title = $Title
        $entry.publishedAt = $PublishedAt
        if (-not [string]::IsNullOrWhiteSpace($ContractNumber)) {
            $entry.contractNumber = $ContractNumber
        }
        if (-not [string]::IsNullOrWhiteSpace($PortalContractId)) {
            $entry.portalContractId = $PortalContractId
        }
        if (-not [string]::IsNullOrWhiteSpace($SourceView)) {
            $entry.sourceView = $SourceView
        }
        if (-not [string]::IsNullOrWhiteSpace($CrossKey)) {
            $entry.crossKey = $CrossKey
        }
    }

    $entry.occurrenceCount = [int]$entry.occurrenceCount + 1
    $existingTags = @([string[]]$entry.tags)
    foreach ($tag in @($Tags)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$tag) -and -not ($existingTags -contains [string]$tag)) {
            $existingTags += [string]$tag
        }
    }
    $entry.tags = @($existingTags | Select-Object -First 6)
}

function Convert-ContractCrossIssueStoreToItems {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Store
    )

    return @(
        $Store.GetEnumerator() |
        ForEach-Object {
            $entry = $_.Value
            $reason = [string]$entry.reason
            if ([int]$entry.occurrenceCount -gt 1) {
                $reason = "$reason Ocorrencias consolidadas: $([int]$entry.occurrenceCount)."
            }

            [pscustomobject][ordered]@{
                type = [string]$entry.type
                severity = [string]$entry.severity
                crossKey = [string]$entry.crossKey
                contractNumber = [string]$entry.contractNumber
                portalContractId = [string]$entry.portalContractId
                title = [string]$entry.title
                reason = $reason
                publishedAt = [string]$entry.publishedAt
                sourceView = [string]$entry.sourceView
                occurrenceCount = [int]$entry.occurrenceCount
                tags = @($entry.tags)
            }
        }
    )
}

function Test-ContractCrossHasExecutionEvidence {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Movements
    )

    return @(
        @($Movements) |
        Where-Object { [string]$_.recordClass -eq 'execucao_contratual' }
    ).Count -gt 0
}

function Test-ShouldMaterializeOfficialWithoutDiary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Official,

        [Parameter(Mandatory = $true)]
        [object]$Vigency
    )

    if ([bool]$Vigency.isActive) {
        return $true
    }

    $recentCutoffTicks = (Get-Date).AddDays(-365).Ticks
    $publishedAtTicks = Get-DateTimestampValue -Text ([string]$Official.publishedAt)
    $signatureAt = Convert-BrazilianDateToDateTime -Text ([string]$Official.signatureDate)
    $signatureTicks = if ($signatureAt) { $signatureAt.Ticks } else { 0 }
    return ($publishedAtTicks -ge $recentCutoffTicks -or $signatureTicks -ge $recentCutoffTicks)
}

function Get-ContractFinancialMonitoringSources {
    $expensePortal = Get-LegacyExpensePortalMetadata
    $expenseQueryModeLabel = if ([bool]$expensePortal.requiresCaptcha -and [bool]$expensePortal.requiresPostback) {
        'Formulario legado com CAPTCHA e postback'
    }
    elseif ([bool]$expensePortal.requiresPostback) {
        'Formulario legado com postback'
    }
    else {
        'Consulta web oficial'
    }

    $expenseNoteParts = @(
        [string]$expensePortal.summaryLabel,
        $(if (@($expensePortal.observedSignals).Count -gt 0) { [string]$expensePortal.observedSignals[0] } else { '' })
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

    return @(
        [pscustomobject][ordered]@{
            key = 'portal_transparencia_despesa'
            label = 'Portal Transparencia - Despesa'
            href = ([Uri]::new($script:LegacyTransparencyPortalUri, $script:LegacyTransparencyExpensePath)).AbsoluteUri
            role = 'empenho_liquidacao_pagamento'
            note = if (@($expenseNoteParts).Count -gt 0) { $expenseNoteParts -join ' ' } else { 'Fonte oficial de empenho, liquidacao e pagamento no portal legado do municipio.' }
            automationStatus = 'assistida'
            automationStatusLabel = 'Consulta guiada'
            queryMode = 'formulario_legado'
            queryModeLabel = $expenseQueryModeLabel
            coverage = 'assisted'
            coverageLabel = 'Consulta guiada oficial'
        },
        [pscustomobject][ordered]@{
            key = 'portal_transparencia_contratos'
            label = 'Portal Transparencia - Contratos'
            href = ([Uri]::new($script:LegacyTransparencyPortalUri, $script:LegacyTransparencyContractPath)).AbsoluteUri
            role = 'cadastro_contratual_complementar'
            note = 'Cadastro contratual complementar do portal legado para conferencia manual.'
            automationStatus = 'assistida'
            automationStatusLabel = 'Consulta assistida'
            queryMode = 'formulario_legado'
            queryModeLabel = 'Formulario legado com postback'
            coverage = 'assisted'
            coverageLabel = 'Consulta assistida'
        },
        [pscustomobject][ordered]@{
            key = 'dados_iguape'
            label = 'Dados Iguape'
            href = $script:DadosIguapeUri.AbsoluteUri
            role = 'api_publica_de_receitas'
            note = 'Painel oficial complementar com API publica identificada em /api/receitas/analise; a execucao contratual especifica de despesa ainda nao foi automatizada.'
            automationStatus = 'ready'
            automationStatusLabel = 'API reconhecida'
            queryMode = 'api_publica'
            queryModeLabel = 'API publica complementar'
            coverage = 'limited'
            coverageLabel = 'Cobertura complementar'
        }
    )
}

function Add-ContractFinancialSearchHint {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Collection,

        [Parameter(Mandatory = $true)]
        [hashtable]$Seen,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    $cleanValue = Collapse-Whitespace -Text $Value
    if ([string]::IsNullOrWhiteSpace($cleanValue)) {
        return
    }

    $key = ((Normalize-IndexText -Text $Label) + '|' + (Normalize-IndexText -Text $cleanValue))
    if ($Seen.ContainsKey($key)) {
        return
    }

    $Seen[$key] = $true
    $Collection.Add([pscustomobject][ordered]@{
        label = $Label
        value = $cleanValue
    })
}

function Get-ContractFinancialSearchHints {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$OfficialContract,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ManagementProfile,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$RelatedMovements = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = ''
    )

    $hints = @()
    $seen = @{}

    $officialPublishedAt = if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'publishedAt' } else { '' }
    $candidateHints = @(
        [pscustomobject]@{ label = 'Referencia principal'; value = $Reference },
        [pscustomobject]@{ label = 'Contrato'; value = $(if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'contractNumber' } else { '' }) },
        [pscustomobject]@{ label = 'Processo'; value = $(if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'processNumber' } else { '' }) },
        [pscustomobject]@{ label = 'Fornecedor'; value = $(if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'contractor' } else { '' }) },
        [pscustomobject]@{ label = 'CNPJ'; value = $(if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'cnpj' } else { '' }) },
        [pscustomobject]@{ label = 'Orgao'; value = $(if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'primaryOrganizationName' } else { '' }) },
        [pscustomobject]@{ label = 'Modalidade'; value = $(if ($OfficialContract) { Get-ObjectStringValue -Item $OfficialContract -Name 'modality' } else { '' }) },
        [pscustomobject]@{ label = 'Ano base'; value = $(if (-not [string]::IsNullOrWhiteSpace($officialPublishedAt) -and $officialPublishedAt.Length -ge 4) { $officialPublishedAt.Substring(0, 4) } else { '' }) },
        [pscustomobject]@{ label = 'Gestor atual'; value = $(if ($ManagementProfile) { Get-ObjectStringValue -Item $ManagementProfile -Name 'managerName' } else { '' }) },
        [pscustomobject]@{ label = 'Fiscal atual'; value = $(if ($ManagementProfile) { Get-ObjectStringValue -Item $ManagementProfile -Name 'inspectorName' } else { '' }) }
    )

    foreach ($movement in @($RelatedMovements | Select-Object -First 4)) {
        $candidateHints += [pscustomobject]@{
            label = 'Ato relacionado'
            value = (Get-ObjectStringValue -Item $movement -Name 'actTitle')
        }
    }

    foreach ($candidate in @($candidateHints)) {
        $cleanValue = Collapse-Whitespace -Text ([string]$candidate.value)
        if ([string]::IsNullOrWhiteSpace($cleanValue)) {
            continue
        }

        $key = ((Normalize-IndexText -Text ([string]$candidate.label)) + '|' + (Normalize-IndexText -Text $cleanValue))
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        $hints += [pscustomobject][ordered]@{
            label = [string]$candidate.label
            value = $cleanValue
        }
    }

    return @($hints)
}

function Get-ContractFinancialPreferredHints {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$SearchHints = @()
    )

    $orderedLabels = @('Referencia principal', 'Contrato', 'Processo', 'Fornecedor', 'CNPJ', 'Orgao', 'Ano base', 'Modalidade')
    $preferred = @()
    foreach ($label in $orderedLabels) {
        $hint = @($SearchHints | Where-Object { [string]$_.label -eq $label } | Select-Object -First 1) | Select-Object -First 1
        if ($null -ne $hint) {
            $preferred += $hint
        }
    }

    return @($preferred)
}

function Get-ContractFinancialHintValue {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$SearchHints = @(),

        [Parameter(Mandatory = $true)]
        [string[]]$Labels
    )

    foreach ($label in @($Labels)) {
        $entry = @($SearchHints | Where-Object { [string]$_.label -eq $label } | Select-Object -First 1) | Select-Object -First 1
        if ($null -ne $entry -and -not [string]::IsNullOrWhiteSpace([string]$entry.value)) {
            return (Collapse-Whitespace -Text ([string]$entry.value))
        }
    }

    return ''
}

function Get-ContractFinancialQueryPlan {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$OfficialContract,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$SearchHints = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$PortalMetadata = $null,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = ''
    )

    $portal = if ($PortalMetadata) { $PortalMetadata } else { Get-LegacyExpensePortalMetadata }
    $cnpjHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('CNPJ')
    $supplierHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('Fornecedor')
    $processHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('Processo')
    $contractHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('Contrato', 'Referencia principal')
    $organizationHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('Orgao')
    $yearHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('Ano base')
    $modalityHint = Get-ContractFinancialHintValue -SearchHints $SearchHints -Labels @('Modalidade')
    $valueNumber = if ($OfficialContract -and $null -ne $OfficialContract.valueNumber) { [double]$OfficialContract.valueNumber } else { 0.0 }

    $primaryLookup = if (-not [string]::IsNullOrWhiteSpace($cnpjHint) -and [bool]$portal.hasCnpjSearch) {
        [pscustomobject][ordered]@{
            label = 'Consulta principal'
            field = 'Credor / Fornecedor (CPF-CNPJ)'
            value = $cnpjHint
            note = 'Comece pelo documento do favorecido no formulario legado; depois abra o detalhe do empenho para confirmar contrato e processo.'
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($supplierHint) -and [bool]$portal.hasSupplierSearch) {
        [pscustomobject][ordered]@{
            label = 'Consulta principal'
            field = 'Credor / Fornecedor (nome)'
            value = $supplierHint
            note = 'Use o nome do fornecedor como aproximacao inicial e confirme o CNPJ no detalhe do empenho.'
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($processHint)) {
        [pscustomobject][ordered]@{
            label = 'Consulta principal'
            field = 'Grade de despesa + conferencia de processo'
            value = $processHint
            note = 'Sem CNPJ estruturado, a leitura passa pela grade e pela confirmacao do processo no popup oficial.'
        }
    }
    else {
        [pscustomobject][ordered]@{
            label = 'Consulta principal'
            field = 'Revisao assistida'
            value = if (-not [string]::IsNullOrWhiteSpace($Reference)) { $Reference } else { 'Contrato sem chave forte' }
            note = 'A pesquisa precisa combinar grade, popup do empenho e confrontacao manual com o dossie.'
        }
    }

    $fallbackLookups = @()
    if (-not [string]::IsNullOrWhiteSpace($supplierHint) -and [string]$primaryLookup.value -ne $supplierHint) {
        $fallbackLookups += [pscustomobject][ordered]@{
            label = 'Plano B'
            field = 'Credor / Fornecedor (nome)'
            value = $supplierHint
            note = 'Aplicar quando o CNPJ nao devolver linha suficiente na grade.'
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($cnpjHint) -and [string]$primaryLookup.value -ne $cnpjHint) {
        $fallbackLookups += [pscustomobject][ordered]@{
            label = 'Plano B'
            field = 'Credor / Fornecedor (CPF-CNPJ)'
            value = $cnpjHint
            note = 'Aplicar quando a razao social vier abreviada ou variar entre portal e cadastro.'
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($modalityHint)) {
        $fallbackLookups += [pscustomobject][ordered]@{
            label = 'Refino'
            field = 'Mod. Lic.'
            value = $modalityHint
            note = 'Ajuda a reduzir falsos positivos quando ha muitos empenhos para o mesmo fornecedor.'
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($contractHint)) {
        $fallbackLookups += [pscustomobject][ordered]@{
            label = 'Confirmacao'
            field = 'Campo Contrato no detalhe do empenho'
            value = $contractHint
            note = 'Nao e filtro principal da tela inicial; serve para confirmar o vinculo dentro do popup oficial.'
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($processHint)) {
        $fallbackLookups += [pscustomobject][ordered]@{
            label = 'Confirmacao'
            field = 'Campo Processo no detalhe do empenho'
            value = $processHint
            note = 'Confirma que a despesa localizada pertence ao processo administrativo do dossie.'
        }
    }

    $steps = @()
    if ([bool]$portal.requiresCaptcha) {
        $steps += 'Abrir o portal oficial de despesa e validar o CAPTCHA antes de acionar a pesquisa.'
    }
    else {
        $steps += 'Abrir o portal oficial de despesa e iniciar a consulta no formulario legado.'
    }

    $budgetType = Collapse-Whitespace -Text ([string]$portal.defaultFilters.budgetType)
    $expenseType = Collapse-Whitespace -Text ([string]$portal.defaultFilters.expenseType)
    $periodStart = Collapse-Whitespace -Text ([string]$portal.defaultFilters.periodStart)
    $periodEnd = Collapse-Whitespace -Text ([string]$portal.defaultFilters.periodEnd)
    $exerciseValue = if (-not [string]::IsNullOrWhiteSpace($yearHint)) { $yearHint } else { Collapse-Whitespace -Text ([string]$portal.defaultFilters.exercise) }
    $steps += (
        ('Conferir exercicio {0}{1}{2} e manter Tipo = {3} com Tipo de Despesa = {4} na primeira leitura.' -f
            $(if (-not [string]::IsNullOrWhiteSpace($exerciseValue)) { $exerciseValue } else { 'corrente' }),
            $(if (-not [string]::IsNullOrWhiteSpace($periodStart)) { ", periodo inicial $periodStart" } else { '' }),
            $(if (-not [string]::IsNullOrWhiteSpace($periodEnd)) { " e final $periodEnd" } else { '' }),
            $(if (-not [string]::IsNullOrWhiteSpace($budgetType)) { $budgetType } else { 'ORCAMENTARIO' }),
            $(if (-not [string]::IsNullOrWhiteSpace($expenseType)) { $expenseType } else { 'TODOS' })
        )
    )
    $steps += ("Aplicar a consulta principal em {0}: {1}." -f ([string]$primaryLookup.field), ([string]$primaryLookup.value))
    $steps += 'Se a grade nao fechar sozinha, usar os filtros laterais de Empenho, CPF-CNPJ, Credor, Mod. Lic. e Licitacao para reduzir o conjunto.'
    $steps += 'Abrir o popup oficial do empenho e confirmar os campos Contrato, Processo, fornecedor e documento fiscal antes de assumir o vinculo financeiro.'
    $steps += 'Consolidar os totais Empenhado, Liquidado e Pago a partir do bloco de valores e das grades de liquidacoes e pagamentos.'
    if ($valueNumber -gt 0) {
        $steps += ("Comparar os totais financeiros do portal com o valor contratado registrado no cadastro oficial ({0})." -f (Format-BrazilianCurrency -Value $valueNumber))
    }

    $confirmationTargets = @()
    foreach ($target in @(
        [pscustomobject]@{ label = 'Contrato'; value = $contractHint; note = 'Campo exibido no detalhe oficial do empenho.' },
        [pscustomobject]@{ label = 'Processo'; value = $processHint; note = 'Campo exibido no detalhe oficial do empenho.' },
        [pscustomobject]@{ label = 'Fornecedor'; value = $supplierHint; note = 'Conferir com o credor exibido no empenho.' },
        [pscustomobject]@{ label = 'CNPJ'; value = $cnpjHint; note = 'Conferir com o documento do favorecido financeiro.' },
        [pscustomobject]@{ label = 'Orgao'; value = $organizationHint; note = 'Conferir entidade e unidade executora antes de fechar a leitura.' },
        [pscustomobject]@{ label = 'Valor contratado'; value = $(if ($valueNumber -gt 0) { Format-BrazilianCurrency -Value $valueNumber } else { '' }); note = 'Usar como referencia para comparar empenhado, liquidado e pago.' }
    )) {
        if ([string]::IsNullOrWhiteSpace([string]$target.value)) {
            continue
        }

        $confirmationTargets += [pscustomobject][ordered]@{
            label = [string]$target.label
            value = [string]$target.value
            note = [string]$target.note
        }
    }

    return [pscustomobject][ordered]@{
        summary = if (-not [string]::IsNullOrWhiteSpace($cnpjHint)) {
            'Pesquisar primeiro pelo CNPJ do fornecedor e confirmar contrato e processo no detalhe oficial do empenho.'
        }
        elseif (-not [string]::IsNullOrWhiteSpace($supplierHint)) {
            'Pesquisar primeiro pelo fornecedor e fechar o vinculo financeiro no popup oficial do empenho.'
        }
        else {
            'A leitura financeira depende de busca guiada na grade de despesa e confirmacao do contrato no detalhe oficial.'
        }
        primaryLookup = $primaryLookup
        fallbackLookups = @($fallbackLookups)
        steps = @($steps | Select-Object -Unique)
        confirmationTargets = @($confirmationTargets)
        stageChecklist = @($portal.executionStages)
    }
}

function Get-ContractFinancialCoverageInfo {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$OfficialContract,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ManagementProfile,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$RelatedMovements = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$SearchHints = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = ''
    )

    $portalMetadata = Get-LegacyExpensePortalMetadata
    $score = 0
    $reasons = New-Object 'System.Collections.Generic.List[string]'
    $queryChecklist = New-Object 'System.Collections.Generic.List[string]'
    $valueNumber = if ($OfficialContract -and $null -ne $OfficialContract.valueNumber) { [double]$OfficialContract.valueNumber } else { 0.0 }
    $revenueExpense = if ($OfficialContract) { Normalize-IndexText -Text ([string]$OfficialContract.revenueExpense) } else { '' }
    $preferredHints = @(Get-ContractFinancialPreferredHints -SearchHints $SearchHints)
    $hintLabels = @($preferredHints | ForEach-Object { [string]$_.label })
    $sourceItems = @(Get-ContractFinancialMonitoringSources)

    if ([string]$portalMetadata.status -eq 'available') {
        $score += 8
        $reasons.Add('portal oficial de despesa acessivel')
    }
    elseif (@($portalMetadata.executionStages).Count -gt 0) {
        $score += 5
        $reasons.Add('roteiro financeiro oficial mapeado')
    }

    if ($OfficialContract) {
        $score += 22
        $reasons.Add('cadastro oficial consolidado')
    }

    if ($Reference -match '\d{1,10}/\d{4}' -or $hintLabels -contains 'Contrato') {
        $score += 20
        $reasons.Add('referencia contratual aproveitavel para confirmacao')
        $queryChecklist.Add('Ao abrir o detalhe do empenho, conferir o campo Contrato com a referencia principal do dossie.')
    }

    if ($hintLabels -contains 'Processo') {
        $score += 14
        $reasons.Add('processo disponivel para consulta')
        $queryChecklist.Add('Conferir o campo Processo no detalhe do empenho antes de fechar o vinculo financeiro.')
    }

    if ($hintLabels -contains 'Fornecedor') {
        $score += 12
        $reasons.Add('fornecedor identificado')
        $queryChecklist.Add('Usar o campo Credor / Fornecedor para aproximar a pesquisa quando o CNPJ nao fechar sozinho.')
    }

    if ($hintLabels -contains 'CNPJ') {
        $score += 12
        $reasons.Add('cnpj disponivel para confirmacao')
        $queryChecklist.Add('Preencher o CPF-CNPJ do fornecedor no formulario legado e conferir se o favorecido financeiro corresponde ao cadastro contratual.')
    }

    if ($hintLabels -contains 'Orgao') {
        $score += 9
        $reasons.Add('orgao principal identificado')
        $queryChecklist.Add('Conferir a entidade ou orgao pagador antes de abrir os detalhes do empenho.')
    }

    if ($hintLabels -contains 'Ano base') {
        $score += 6
        $reasons.Add('ano base definido')
    }

    if ($valueNumber -gt 0) {
        $score += 8
        $reasons.Add('valor contratado estruturado')
        $queryChecklist.Add('Comparar o valor contratado com os totais Empenhado, Liquidado e Pago exibidos no detalhe oficial.')
    }

    if ($revenueExpense -match 'DESPESA') {
        $score += 6
        $reasons.Add('classificado como despesa')
    }
    elseif ($revenueExpense -match 'RECEITA') {
        $score += 3
        $reasons.Add('classificado como receita')
    }

    if (@($RelatedMovements).Count -gt 0) {
        $score += 8
        $reasons.Add('atos correlatos no diario')
    }

    if ($ManagementProfile -and ([bool]$ManagementProfile.hasManager -or [bool]$ManagementProfile.hasInspector)) {
        $score += 4
        $reasons.Add('gestao contratual identificada')
    }

    if (@($SearchHints).Count -ge 5) {
        $score += 8
        $reasons.Add('conjunto de chaves forte para consulta')
    }
    elseif (@($SearchHints).Count -ge 3) {
        $score += 4
        $reasons.Add('chaves minimas para consulta manual')
    }

    if ([bool]$portalMetadata.requiresCaptcha) {
        $queryChecklist.Add('Validar o CAPTCHA na abertura da tela oficial antes de usar a pesquisa assistida.')
    }
    if ([bool]$portalMetadata.requiresPostback) {
        $queryChecklist.Add('Tratar a consulta como guiada: o portal depende de postback ASP.NET e nao aceita automacao cega estavel.')
    }
    if ([bool]$portalMetadata.requiresCallback) {
        $queryChecklist.Add('Abrir o popup do empenho para acessar os blocos de liquidacoes e pagamentos.')
    }

    $score = [Math]::Min(100, [int]$score)
    $coverageLevel = if ($score -ge 78) {
        'assisted'
    }
    elseif ($score -ge 45) {
        'limited'
    }
    else {
        'none'
    }

    $coverageLabel = switch ($coverageLevel) {
        'assisted' { 'Consulta assistida forte' }
        'limited' { 'Cobertura limitada' }
        default { 'Cobertura insuficiente' }
    }

    $queryReadiness = if ($score -ge 78) {
        'alta'
    }
    elseif ($score -ge 45) {
        'media'
    }
    else {
        'baixa'
    }

    $queryReadinessLabel = switch ($queryReadiness) {
        'alta' { 'Pronto para pesquisa' }
        'media' { 'Pesquisa com apoio' }
        default { 'Requer complemento manual' }
    }

    $queryLinks = @(
        $sourceItems |
        ForEach-Object {
            $preferredHint = if ([string]$_.key -eq 'portal_transparencia_despesa') {
                @($preferredHints | Where-Object { [string]$_.label -in @('Referencia principal', 'Contrato', 'Processo', 'Fornecedor', 'CNPJ') } | Select-Object -First 1) | Select-Object -First 1
            }
            elseif ([string]$_.key -eq 'portal_transparencia_contratos') {
                @($preferredHints | Where-Object { [string]$_.label -in @('Contrato', 'Processo', 'Fornecedor', 'CNPJ') } | Select-Object -First 1) | Select-Object -First 1
            }
            else {
                @($preferredHints | Where-Object { [string]$_.label -in @('Ano base', 'Orgao', 'Fornecedor', 'CNPJ') } | Select-Object -First 1) | Select-Object -First 1
            }

            [pscustomobject][ordered]@{
                key = [string]$_.key
                label = [string]$_.label
                href = [string]$_.href
                queryMode = [string]$_.queryMode
                queryModeLabel = [string]$_.queryModeLabel
                suggestedQuery = if ($preferredHint) { "$([string]$preferredHint.label): $([string]$preferredHint.value)" } else { '' }
                note = [string]$_.note
            }
        }
    )

    $summaryLabel = if ($OfficialContract) {
        if ($coverageLevel -eq 'assisted') {
            'O dossie ja traz chaves suficientes para consulta assistida consistente nas bases oficiais financeiras.'
        }
        elseif ($coverageLevel -eq 'limited') {
            'O dossie aponta as fontes oficiais e algumas chaves, mas a conferencia financeira ainda exige leitura manual complementar.'
        }
        else {
            'Ainda faltam chaves estruturadas para uma consulta financeira reproduzivel com baixa friccao.'
        }
    }
    else {
        'Sem cadastro oficial consolidado para sustentar uma leitura financeira confiavel.'
    }

    $automationStatus = switch ($coverageLevel) {
        'assisted' { 'assistida' }
        'limited' { 'parcial' }
        default { 'pendente' }
    }

    $automationNote = if ($coverageLevel -eq 'assisted') {
        'O portal legado de despesa continua exigindo CAPTCHA, postback ASP.NET e callbacks no detalhe, mas as chaves deste contrato ja permitem consulta guiada previsivel e repetivel.'
    }
    elseif ($coverageLevel -eq 'limited') {
        'A pesquisa financeira ainda depende do portal legado ASP.NET e de conferencia manual no popup oficial para fechar o vinculo entre contrato e despesa.'
    }
    else {
        'Ainda nao ha chaves suficientes para uma rotina financeira assistida confiavel no portal legado.'
    }

    return [pscustomobject][ordered]@{
        coverageLevel = $coverageLevel
        coverageLabel = $coverageLabel
        coverageScore = [int]$score
        queryReadiness = $queryReadiness
        queryReadinessLabel = $queryReadinessLabel
        automationStatus = $automationStatus
        summaryLabel = $summaryLabel
        automationNote = $automationNote
        portalStatus = [string]$portalMetadata.status
        portalStatusLabel = [string]$portalMetadata.statusLabel
        portalRequiresCaptcha = [bool]$portalMetadata.requiresCaptcha
        executionStageCount = [int]@($portalMetadata.executionStages).Count
        detailSectionCount = [int]@($portalMetadata.detailSections).Count
        reasons = @($reasons | Select-Object -Unique)
        queryChecklist = @($queryChecklist | Select-Object -Unique)
        queryLinks = @($queryLinks)
    }
}

function Get-ContractSearchIndex {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $crossSource = if (Test-ObjectProperty -Item $Item -Name 'crossSource') { $Item.crossSource } else { $null }
    $operationalAlerts = if (Test-ObjectProperty -Item $Item -Name 'operationalAlerts') { @($Item.operationalAlerts) } else { @() }
    $values = @(
        (Get-ObjectStringValue -Item $Item -Name 'actTitle'),
        (Get-ObjectStringValue -Item $Item -Name 'type'),
        (Get-ObjectStringValue -Item $Item -Name 'recordClassLabel'),
        (Get-ObjectStringValue -Item $Item -Name 'contractor'),
        (Get-ObjectStringValue -Item $Item -Name 'object'),
        (Get-ObjectStringValue -Item $Item -Name 'contractNumber'),
        (Get-ObjectStringValue -Item $Item -Name 'processNumber'),
        (Get-ObjectStringValue -Item $Item -Name 'modality'),
        (Get-ObjectStringValue -Item $Item -Name 'portalStatus'),
        (Get-ObjectStringValue -Item $Item -Name 'managerName'),
        (Get-ObjectStringValue -Item $Item -Name 'managerRole'),
        (Get-ObjectStringValue -Item $Item -Name 'inspectorName'),
        (Get-ObjectStringValue -Item $Item -Name 'inspectorRole'),
        (Get-ObjectStringValue -Item $Item -Name 'managementSummary'),
        (Get-ObjectStringValue -Item $Item -Name 'referenceKey'),
        (Get-ObjectStringValue -Item $Item -Name 'primaryOrganizationName'),
        (Get-ObjectStringValue -Item $Item -Name 'excerpt'),
        (Get-ObjectStringValue -Item $Item -Name 'cnpj')
    )

    if (Test-ObjectProperty -Item $Item -Name 'mentionedOrganizationNames') {
        $values += @($Item.mentionedOrganizationNames)
    }
    if ($crossSource) {
        $values += @(
            [string]$crossSource.reason,
            @($crossSource.officialNumbers),
            @($crossSource.officialStatuses),
            @($crossSource.divergenceTypes)
        )
    }
    foreach ($alert in @($operationalAlerts)) {
        $values += ("{0} {1}" -f ([string]$alert.title), ([string]$alert.reason))
    }

    return (Normalize-IndexText -Text ($values -join ' ')).ToLowerInvariant()
}

function Get-ContractsFinancialMonitoringSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$OfficialContracts
    )

    $expensePortal = Get-LegacyExpensePortalMetadata
    $sources = @(Get-ContractFinancialMonitoringSources)
    $searchableContracts = @(
        $OfficialContracts |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_.contractNumber) -or
            -not [string]::IsNullOrWhiteSpace([string]$_.processNumber) -or
            -not [string]::IsNullOrWhiteSpace([string]$_.contractor)
        }
    ).Count
    $withContractValue = @(
        $OfficialContracts |
        Where-Object { $null -ne $_.valueNumber -and [double]$_.valueNumber -gt 0 }
    ).Count
    $expenseContracts = @(
        $OfficialContracts |
        Where-Object { (Normalize-IndexText -Text ([string]$_.revenueExpense)) -match 'DESPESA' }
    ).Count
    $revenueContracts = @(
        $OfficialContracts |
        Where-Object { (Normalize-IndexText -Text ([string]$_.revenueExpense)) -match 'RECEITA' }
    ).Count
    $coverageItems = @(
        $OfficialContracts |
        ForEach-Object {
            $reference = Get-ContractCrossKey -Item $_
            Get-ContractFinancialCoverageInfo `
                -OfficialContract $_ `
                -ManagementProfile $null `
                -RelatedMovements @() `
                -SearchHints (Get-ContractFinancialSearchHints -OfficialContract $_ -ManagementProfile $null -RelatedMovements @() -Reference $reference) `
                -Reference $reference
        }
    )
    $queryReadyContracts = @(
        $coverageItems |
        Where-Object { [string]$_.coverageLevel -eq 'assisted' }
    ).Count
    $limitedContracts = @(
        $coverageItems |
        Where-Object { [string]$_.coverageLevel -eq 'limited' }
    ).Count
    $unmappedContracts = [Math]::Max(0, @($OfficialContracts).Count - $queryReadyContracts - $limitedContracts)
    $averageCoverageScore = if (@($coverageItems).Count -gt 0) {
        [Math]::Round((@($coverageItems | Measure-Object -Property coverageScore -Average).Average), 0)
    }
    else {
        0
    }
    $automationReadySources = @(
        $sources |
        Where-Object { [string]$_.automationStatus -eq 'ready' }
    ).Count

    return [ordered]@{
        mode = if ($queryReadyContracts -gt 0) { 'assisted' } else { 'partial' }
        modeLabel = if ($queryReadyContracts -gt 0) { 'Cobertura guiada' } else { 'Integracao parcial' }
        note = if ($queryReadyContracts -gt 0) {
            'O painel ja separa contratos com chaves suficientes para pesquisa guiada nas fontes oficiais. O espelho automatico completo ainda depende do portal legado com CAPTCHA, postback e callback.'
        }
        else {
            'A API oficial de receitas do Dados Iguape ja foi reconhecida pelo painel. A leitura de empenho, liquidacao e pagamento continua dependente do portal legado com CAPTCHA, postback e callback.'
        }
        monitoredContracts = [int]@($OfficialContracts).Count
        searchableContracts = [int]$searchableContracts
        queryReadyContracts = [int]$queryReadyContracts
        withContractValue = [int]$withContractValue
        expenseContracts = [int]$expenseContracts
        revenueContracts = [int]$revenueContracts
        automatedContracts = 0
        assistedContracts = [int]$queryReadyContracts
        limitedContracts = [int]$limitedContracts
        unmappedContracts = [int]$unmappedContracts
        averageCoverageScore = [int]$averageCoverageScore
        automationReadySources = [int]$automationReadySources
        sourceCount = [int]@($sources).Count
        executionStageCount = [int]@($expensePortal.executionStages).Count
        detailSectionCount = [int]@($expensePortal.detailSections).Count
        expensePortal = [pscustomobject][ordered]@{
            status = [string]$expensePortal.status
            statusLabel = [string]$expensePortal.statusLabel
            requiresCaptcha = [bool]$expensePortal.requiresCaptcha
            requiresCallback = [bool]$expensePortal.requiresCallback
            executionStageCount = [int]@($expensePortal.executionStages).Count
            detailSectionCount = [int]@($expensePortal.detailSections).Count
            accessNote = [string]$expensePortal.accessNote
        }
        coverageBreakdown = @(
            [pscustomobject][ordered]@{
                key = 'assisted'
                label = 'Consulta assistida forte'
                count = [int]$queryReadyContracts
            },
            [pscustomobject][ordered]@{
                key = 'limited'
                label = 'Cobertura limitada'
                count = [int]$limitedContracts
            },
            [pscustomobject][ordered]@{
                key = 'none'
                label = 'Sem chave suficiente'
                count = [int]$unmappedContracts
            }
        )
        sources = @($sources)
    }
}

function Get-ContractFinancialExecutionInfo {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$OfficialContract,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ManagementProfile,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$RelatedMovements = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = ''
    )

    $portalMetadata = Get-LegacyExpensePortalMetadata
    $sources = @(Get-ContractFinancialMonitoringSources)
    $searchHints = @(Get-ContractFinancialSearchHints -OfficialContract $OfficialContract -ManagementProfile $ManagementProfile -RelatedMovements $RelatedMovements -Reference $Reference)
    $latestMovementAt = @($RelatedMovements | ForEach-Object { [string]$_.publishedAt } | Sort-Object -Descending | Select-Object -First 1) | Select-Object -First 1
    $valueNumber = if ($OfficialContract -and $null -ne $OfficialContract.valueNumber) { [double]$OfficialContract.valueNumber } else { 0.0 }
    $coverage = Get-ContractFinancialCoverageInfo -OfficialContract $OfficialContract -ManagementProfile $ManagementProfile -RelatedMovements $RelatedMovements -SearchHints $searchHints -Reference $Reference
    $queryPlan = Get-ContractFinancialQueryPlan -OfficialContract $OfficialContract -SearchHints $searchHints -PortalMetadata $portalMetadata -Reference $Reference
    $executionStages = @($portalMetadata.executionStages)
    $detailSections = @($portalMetadata.detailSections)
    $queryFields = @($portalMetadata.queryFields)
    $limitations = @($portalMetadata.limitations)

    return [ordered]@{
        status = [string]$coverage.coverageLevel
        statusLabel = [string]$coverage.coverageLabel
        automationStatus = [string]$coverage.automationStatus
        automationNote = [string]$coverage.automationNote
        summaryLabel = [string]$coverage.summaryLabel
        executionSummaryLabel = [string]$portalMetadata.summaryLabel
        portalAccessLabel = [string]$portalMetadata.statusLabel
        portalAccessNote = [string]$portalMetadata.accessNote
        coverageLevel = [string]$coverage.coverageLevel
        coverageLabel = [string]$coverage.coverageLabel
        coverageScore = [int]$coverage.coverageScore
        queryReadiness = [string]$coverage.queryReadiness
        queryReadinessLabel = [string]$coverage.queryReadinessLabel
        contractedValue = $valueNumber
        contractedValueLabel = if ($valueNumber -gt 0) { (Format-BrazilianCurrency -Value $valueNumber) } else { 'Nao informado no cadastro oficial' }
        revenueExpense = if ($OfficialContract) { [string]$OfficialContract.revenueExpense } else { '' }
        lastMovementAt = [string]$latestMovementAt
        relatedMovements = [int]@($RelatedMovements).Count
        searchableHintCount = [int]@($searchHints).Count
        coverageReasons = @($coverage.reasons)
        queryChecklist = @($coverage.queryChecklist)
        queryLinks = @($coverage.queryLinks)
        executionStageCount = [int]@($executionStages).Count
        detailSectionCount = [int]@($detailSections).Count
        limitations = @($limitations)
        queryPlan = $queryPlan
        executionStages = @($executionStages)
        detailSections = @($detailSections)
        queryFields = @($queryFields)
        portalMetadata = $portalMetadata
        sources = @($sources)
        searchHints = @($searchHints)
    }
}

function Convert-BrazilianDateToDateTime {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $normalized = Collapse-Whitespace -Text $Text
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($normalized, 'dd/MM/yyyy', $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function Add-DateDurationValue {
    param(
        [Parameter(Mandatory = $true)]
        [DateTime]$Date,

        [Parameter(Mandatory = $true)]
        [int]$Amount,

        [Parameter(Mandatory = $true)]
        [string]$Unit
    )

    $normalizedUnit = Normalize-IndexText -Text $Unit
    if ($normalizedUnit -match '^ANO') {
        return $Date.AddYears($Amount)
    }
    if ($normalizedUnit -match '^MES') {
        return $Date.AddMonths($Amount)
    }
    return $Date.AddDays($Amount)
}

function Get-OfficialContractVigencyInfo {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $portalStatus = Collapse-Whitespace -Text ([string]$Item.portalStatus)
    $signatureDate = Convert-BrazilianDateToDateTime -Text ([string]$Item.signatureDate)
    $combinedText = Collapse-Whitespace -Text (
        @(
            [string]$Item.object,
            [string]$Item.excerpt,
            [string]$Item.term,
            [string]$Item.actTitle
        ) -join ' '
    )

    $endDate = $null
    $source = ''

    if ($combinedText -match '(?:VIGENCIA|PRAZO(?:\s+DE\s+VIGENCIA)?|ENCERRAMENTO)[^\d]{0,24}(?<date>\d{2}/\d{2}/\d{4})') {
        $endDate = Convert-BrazilianDateToDateTime -Text ([string]$matches['date'])
        $source = 'document_explicit_date'
    }

    if ($null -eq $endDate -and $signatureDate) {
        $durationMatch = [regex]::Match(
            (Get-ContractCrossSimpleText -Text $combinedText),
            '(?:VIGENCIA|PRAZO DE VIGENCIA|EXECUCAO E VIGENCIA)[^\d]{0,20}(?<amount>\d{1,4})\s*(?:\((?:[^)]*)\))?\s*(?<unit>ANOS?|MESES?|DIAS?)'
        )
        if ($durationMatch.Success) {
            $endDate = Add-DateDurationValue -Date $signatureDate -Amount ([int]$durationMatch.Groups['amount'].Value) -Unit ([string]$durationMatch.Groups['unit'].Value)
            $source = 'document_duration'
        }
    }

    $today = (Get-Date).Date
    $activeByPortal = ($portalStatus.ToUpperInvariant() -eq 'VIGENTE')
    $activeByDocument = ($null -ne $endDate -and $endDate.Date -ge $today)
    $isActive = ($activeByPortal -or $activeByDocument)

    return [ordered]@{
        isActive = [bool]$isActive
        portalStatus = $portalStatus
        signatureDate = if ($signatureDate) { $signatureDate.ToString('s') } else { $null }
        endDate = if ($endDate) { $endDate.ToString('s') } else { $null }
        source = $source
        activeByPortal = [bool]$activeByPortal
        activeByDocument = [bool]$activeByDocument
        daysUntilEnd = if ($endDate) { [int][Math]::Floor(($endDate.Date - $today).TotalDays) } else { $null }
        sourceLabel = if ($activeByPortal) {
            'Status oficial do portal'
        }
        elseif ($activeByDocument) {
            'Vigencia inferida do documento'
        }
        elseif (-not [string]::IsNullOrWhiteSpace($portalStatus)) {
            "Situacao oficial: $portalStatus"
        }
        else {
            'Sem vigencia identificada'
        }
        summaryLabel = if ($activeByPortal -and $endDate) {
            ('Vigente ate {0}' -f $endDate.ToString('dd/MM/yyyy'))
        }
        elseif ($activeByPortal) {
            'Vigente no portal'
        }
        elseif ($activeByDocument -and $endDate) {
            ('Vigencia inferida ate {0}' -f $endDate.ToString('dd/MM/yyyy'))
        }
        elseif ($endDate) {
            ('Encerrado em {0}' -f $endDate.ToString('dd/MM/yyyy'))
        }
        elseif (-not [string]::IsNullOrWhiteSpace($portalStatus)) {
            "Situacao oficial: $portalStatus"
        }
        else {
            'Sem vigencia identificada'
        }
    }
}

function Resolve-ContractCrosswalk {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$OfficialContracts,

        [Parameter(Mandatory = $true)]
        [object[]]$ContractMovements,

        [Parameter(Mandatory = $true)]
        [object]$ReviewPayload
    )

    $officialByCrossKey = @{}
    $movementByCrossKey = @{}
    $officialMatchesByItemKey = @{}
    $pendingReviewByOfficialItemKey = @{}
    $reviewDecisionIndex = @{}
    $reviewQueue = @()
    $divergenceStore = @{}
    $suppressedDivergenceCounter = @{}
    $alerts = @()
    $latestMatchedAt = [Int64]0
    $automaticMatches = 0
    $reviewedMatches = 0
    $materialMovementWithoutOfficialKeys = @{}

    foreach ($decision in @($ReviewPayload.decisions)) {
        $movementKey = [string]$decision.movementKey
        if (-not [string]::IsNullOrWhiteSpace($movementKey)) {
            $reviewDecisionIndex[$movementKey] = $decision
        }
    }

    foreach ($item in @($OfficialContracts)) {
        $crossKey = Get-ContractCrossKey -Item $item
        if (-not [string]::IsNullOrWhiteSpace($crossKey)) {
            if (-not $officialByCrossKey.ContainsKey($crossKey)) {
                $officialByCrossKey[$crossKey] = @()
            }
            $officialByCrossKey[$crossKey] += $item
        }
    }

    foreach ($movement in @($ContractMovements)) {
        $crossKey = Get-ContractCrossKey -Item $movement
        if (-not [string]::IsNullOrWhiteSpace($crossKey)) {
            if (-not $movementByCrossKey.ContainsKey($crossKey)) {
                $movementByCrossKey[$crossKey] = @()
            }
            $movementByCrossKey[$crossKey] += $movement
        }
    }

    foreach ($movement in @($ContractMovements)) {
        try {
        $movementItemKey = Get-ContractCrossItemKey -Item $movement
        $crossKey = Get-ContractCrossKey -Item $movement
        $movementGroup = if ($crossKey -and $movementByCrossKey.ContainsKey($crossKey)) { @($movementByCrossKey[$crossKey]) } else { @($movement) }
        $officialCandidates = if ($crossKey -and $officialByCrossKey.ContainsKey($crossKey)) { @($officialByCrossKey[$crossKey]) } else { @() }
        $selectedOfficial = $null
        $status = 'unmatched'
        $confidence = 'baixa'
        $confidenceScore = 20
        $reason = if ($crossKey) { "Nenhum cadastro oficial foi encontrado para a referencia $crossKey." } else { 'Movimentacao sem referencia principal para cruzamento.' }
        $reviewCandidates = @()
        $reviewDecision = if ($reviewDecisionIndex.ContainsKey($movementItemKey)) { $reviewDecisionIndex[$movementItemKey] } else { $null }
        $reviewDecisionStatus = if ($reviewDecision) { [string]$reviewDecision.status } else { '' }

        if (@($officialCandidates).Count -eq 1) {
            $selectedOfficial = $officialCandidates[0]
            $evaluation = Get-ContractCrossCandidateEvaluation -Movement $movement -Official $selectedOfficial
            $status = 'automatic'
            $confidence = 'alta'
            $confidenceScore = [Math]::Max([int]$evaluation.score, 90)
            $reason = "Vinculo automatico pela referencia $crossKey exclusiva no portal."
            if ($evaluation.reason -and $evaluation.reason -ne 'sem evidencia complementar forte') {
                $reason = "$reason Evidencias: $([string]$evaluation.reason)."
            }
            $automaticMatches++
        }
        elseif (@($officialCandidates).Count -gt 1) {
            $reviewCandidates = @(
                $officialCandidates |
                ForEach-Object {
                    $evaluation = Get-ContractCrossCandidateEvaluation -Movement $movement -Official $_
                    [pscustomobject][ordered]@{
                        portalContractId = [string]$_.portalContractId
                        contractNumber = [string]$_.contractNumber
                        organization = [string]$_.primaryOrganizationName
                        contractor = [string]$_.contractor
                        publishedAt = [string]$_.publishedAt
                        portalStatus = [string]$_.portalStatus
                        score = [int]$evaluation.score
                        confidence = [string]$evaluation.confidence
                        reason = [string]$evaluation.reason
                    }
                } |
                Sort-Object @{ Expression = { [int]$_.score }; Descending = $true }, @{ Expression = { [string]$_.contractNumber }; Descending = $false }
            )

            if ($reviewDecisionStatus -in @('rejected', 'no_link')) {
                $status = 'reviewed_no_match'
                $confidence = 'media'
                $confidenceScore = 75
                $reason = if ($crossKey) { "A referencia $crossKey foi revisada manualmente e marcada como sem vinculo." } else { 'Movimentacao revisada manualmente como sem vinculo oficial.' }
                if (-not [string]::IsNullOrWhiteSpace([string]$reviewDecision.note)) {
                    $reason = "$reason Observacao: $([string]$reviewDecision.note)."
                }
            }
            elseif ($reviewDecision -and -not [string]::IsNullOrWhiteSpace([string]$reviewDecision.officialPortalContractId)) {
                $selectedOfficial = @($officialCandidates | Where-Object { [string]$_.portalContractId -eq [string]$reviewDecision.officialPortalContractId }) | Select-Object -First 1
                if ($selectedOfficial) {
                    $evaluation = Get-ContractCrossCandidateEvaluation -Movement $movement -Official $selectedOfficial
                    $status = 'reviewed'
                    $confidence = 'alta'
                    $confidenceScore = [Math]::Max([int]$evaluation.score, 92)
                    $reason = "Vinculo confirmado manualmente para a referencia $crossKey."
                    if (-not [string]::IsNullOrWhiteSpace([string]$reviewDecision.note)) {
                        $reason = "$reason Observacao: $([string]$reviewDecision.note)."
                    }
                    $reviewedMatches++
                }
            }

            if (-not $selectedOfficial -and $status -ne 'reviewed_no_match') {
                $status = 'pending_review'
                $confidence = if (@($reviewCandidates).Count -gt 0) { [string]$reviewCandidates[0].confidence } else { 'baixa' }
                $confidenceScore = if (@($reviewCandidates).Count -gt 0) { [int]$reviewCandidates[0].score } else { 35 }
                $reason = "A referencia $crossKey possui $(@($officialCandidates).Count) cadastros oficiais e exige revisao manual."
                $reviewQueue += [pscustomobject][ordered]@{
                    movementKey = $movementItemKey
                    crossKey = $crossKey
                    publishedAt = [string]$movement.publishedAt
                    movementTitle = [string]$(if ([string]$movement.actTitle) { $movement.actTitle } else { $movement.contractNumber })
                    movementReference = [string]$(if ([string]$movement.contractNumber) { $movement.contractNumber } else { $movement.processNumber })
                    movementOrganization = [string]$movement.primaryOrganizationName
                    movementContractor = [string]$movement.contractor
                    candidateCount = [int]@($officialCandidates).Count
                    recommendedConfidence = if (@($reviewCandidates).Count -gt 0) { [string]$reviewCandidates[0].confidence } else { 'baixa' }
                    recommendedScore = if (@($reviewCandidates).Count -gt 0) { [int]$reviewCandidates[0].score } else { 0 }
                    candidates = @($reviewCandidates)
                }

                foreach ($candidate in @($officialCandidates)) {
                    $officialItemKey = Get-ContractCrossItemKey -Item $candidate
                    if (-not $pendingReviewByOfficialItemKey.ContainsKey($officialItemKey)) {
                        $pendingReviewByOfficialItemKey[$officialItemKey] = 0
                    }
                    $pendingReviewByOfficialItemKey[$officialItemKey]++
                }
            }
        }

        if ($selectedOfficial) {
            $officialItemKey = Get-ContractCrossItemKey -Item $selectedOfficial
            if (-not $officialMatchesByItemKey.ContainsKey($officialItemKey)) {
                $officialMatchesByItemKey[$officialItemKey] = @()
            }
            $officialMatchesByItemKey[$officialItemKey] += $movement
            $latestMatchedAt = [Math]::Max($latestMatchedAt, (Get-DateTimestampValue -Text ([string]$movement.publishedAt)))
            $latestMatchedAt = [Math]::Max($latestMatchedAt, (Get-DateTimestampValue -Text ([string]$selectedOfficial.publishedAt)))
        }
        elseif ($status -eq 'pending_review') {
            Add-ContractCrossIssue `
                -Store $divergenceStore `
                -Type 'pending_review' `
                -Severity 'warning' `
                -CrossKey $crossKey `
                -ContractNumber ([string]$movement.contractNumber) `
                -PortalContractId $null `
                -Title 'Vinculo pendente de revisao manual' `
                -Reason $reason `
                -PublishedAt ([string]$movement.publishedAt) `
                -SourceView 'movements' `
                -BucketKey ("pending_review|$movementItemKey")
        }
        elseif ($crossKey -and [string]$movement.contractNumber -and [string]$movement.recordClass -ne 'licitacao_ou_contratacao') {
            $hasExecutionEvidence = Test-ContractCrossHasExecutionEvidence -Movements $movementGroup
            if ($hasExecutionEvidence) {
                $materialMovementWithoutOfficialKeys[$crossKey] = $true
                $classTags = @(
                    $movementGroup |
                    ForEach-Object { [string]$_.recordClass } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Unique
                )
                $groupReason = "Foram localizados $(@($movementGroup).Count) ato(s) do Diario para a referencia $crossKey sem cadastro correspondente no portal."
                if (@($classTags).Count -gt 0) {
                    $groupReason = "$groupReason Classes: $($classTags -join ' | ')."
                }

                Add-ContractCrossIssue `
                    -Store $divergenceStore `
                    -Type 'movement_without_official' `
                    -Severity 'warning' `
                    -CrossKey $crossKey `
                    -ContractNumber ([string]$movement.contractNumber) `
                    -PortalContractId $null `
                    -Title 'Referencia do Diario sem cadastro correspondente no portal' `
                    -Reason $groupReason `
                    -PublishedAt ([string]$movement.publishedAt) `
                    -SourceView 'movements' `
                    -Tags $classTags
            }
            else {
                Add-ContractCounterValue -Counter $suppressedDivergenceCounter -Key 'gestao_sem_execucao_sem_cadastro_oficial'
            }
        }

        $movement | Add-Member -NotePropertyName crossSource -NotePropertyValue ([pscustomobject][ordered]@{
            crossKey = $crossKey
            itemKey = $movementItemKey
            matched = [bool]$selectedOfficial
            status = $status
            confidence = $confidence
            confidenceScore = [int]$confidenceScore
            reason = $reason
            candidateCount = [int]@($officialCandidates).Count
            matchedOfficialCount = if ($selectedOfficial) { 1 } else { 0 }
            matchedMovementCount = 0
            officialPortalContractIds = if ($selectedOfficial) { @([string]$selectedOfficial.portalContractId) } else { @() }
            officialNumbers = if ($selectedOfficial) { @([string]$selectedOfficial.contractNumber) } else { @($reviewCandidates | Select-Object -ExpandProperty contractNumber -First 3) }
            officialStatuses = if ($selectedOfficial) { @([string]$selectedOfficial.portalStatus) } else { @($reviewCandidates | Select-Object -ExpandProperty portalStatus -First 3) }
            latestCounterpartAt = if ($selectedOfficial) { [string]$selectedOfficial.publishedAt } else { $null }
            reviewRequired = ($status -eq 'pending_review')
            reviewedBy = if ($reviewDecision) { [string]$reviewDecision.updatedBy } else { $null }
            reviewedAt = if ($reviewDecision) { [string]$reviewDecision.updatedAt } else { $null }
            divergenceCount = 0
            divergenceTypes = @()
        }) -Force
        }
        catch {
            throw "Falha ao cruzar a movimentacao ${movementItemKey}: $($_.Exception.Message)"
        }
    }

    foreach ($official in @($OfficialContracts)) {
        try {
        $officialStage = 'init'
        $officialItemKey = Get-ContractCrossItemKey -Item $official
        $crossKey = Get-ContractCrossKey -Item $official
        $matchedMovements = if ($officialMatchesByItemKey.ContainsKey($officialItemKey)) { @($officialMatchesByItemKey[$officialItemKey]) } else { @() }
        $allOfficialCandidates = if ($crossKey -and $officialByCrossKey.ContainsKey($crossKey)) { @($officialByCrossKey[$crossKey]) } else { @() }
        $duplicateCount = @($allOfficialCandidates).Count
        $pendingReviewCount = if ($pendingReviewByOfficialItemKey.ContainsKey($officialItemKey)) { [int]$pendingReviewByOfficialItemKey[$officialItemKey] } else { 0 }
        $latestMovementAt = @($matchedMovements | ForEach-Object { [string]$_.publishedAt } | Sort-Object -Descending | Select-Object -First 1) | Select-Object -First 1
        $vigency = Get-OfficialContractVigencyInfo -Item $official
        $status = 'unmatched'
        $reason = if ($crossKey) { "Nenhum ato do Diario foi encontrado para a referencia $crossKey." } else { 'Contrato oficial sem referencia principal para cruzamento.' }
        $confidence = 'baixa'
        $confidenceScore = 25
        $officialDivergenceTypes = @()

        $officialStage = 'status'
        if (@($matchedMovements).Count -gt 0 -and $duplicateCount -le 1) {
            $status = 'automatic'
            $confidence = 'alta'
            $confidenceScore = 95
            $reason = "Vinculo automatico com $(@($matchedMovements).Count) ato(s) do Diario pela referencia $crossKey."
        }
        elseif (@($matchedMovements).Count -gt 0) {
            $status = if ($pendingReviewCount -gt 0) { 'pending_review' } else { 'reviewed' }
            $confidence = 'alta'
            $confidenceScore = if ($status -eq 'reviewed') { 92 } else { 85 }
            $reason = if ($status -eq 'reviewed') {
                "Vinculo confirmado manualmente para a referencia $crossKey."
            }
            else {
                "Parte do vinculo para a referencia $crossKey ainda depende de revisao manual."
            }
        }
        elseif ($pendingReviewCount -gt 0) {
            $status = 'pending_review'
            $confidence = 'media'
            $confidenceScore = 55
            $reason = "A referencia $crossKey possui candidaturas pendentes de revisao manual."
        }

        $officialStage = 'divergences'
        if (@($matchedMovements).Count -gt 0) {
            $movementOrganizations = @($matchedMovements | ForEach-Object { [string]$_.primaryOrganizationName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $movementContractors = @($matchedMovements | ForEach-Object { [string]$_.contractor } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $movementCnpjs = @($matchedMovements | ForEach-Object { Get-ContractCrossNormalizedCnpj -Text ([string]$_.cnpj) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $movementValues = @($matchedMovements | Where-Object { $null -ne $_.valueNumber -and [double]$_.valueNumber -gt 0 } | ForEach-Object { [double]$_.valueNumber })
            $officialNormalizedCnpj = Get-ContractCrossNormalizedCnpj -Text ([string]$official.cnpj)

            if ([string]$official.primaryOrganizationName -and @($movementOrganizations).Count -gt 0 -and @($movementOrganizations | Where-Object { -not (Test-ContractCrossCompatibleText -Left $_ -Right ([string]$official.primaryOrganizationName)) }).Count -gt 0) {
                $officialDivergenceTypes += 'organization_mismatch'
                Add-ContractCrossIssue `
                    -Store $divergenceStore `
                    -Type 'organization_mismatch' `
                    -Severity 'warning' `
                    -CrossKey $crossKey `
                    -ContractNumber ([string]$official.contractNumber) `
                    -PortalContractId ([string]$official.portalContractId) `
                    -Title 'Orgao divergente entre portal e Diario' `
                    -Reason ("Portal: $([string]$official.primaryOrganizationName) | Diario: $($movementOrganizations -join ' | ')") `
                    -PublishedAt ([string]$(if ($latestMovementAt) { $latestMovementAt } else { $official.publishedAt })) `
                    -SourceView 'official'
            }

            $hasSupplierAliasMismatch = [string]$official.contractor -and @($movementContractors).Count -gt 0 -and @($movementContractors | Where-Object { -not (Test-ContractCrossCompatibleText -Left $_ -Right ([string]$official.contractor)) }).Count -gt 0
            $hasSupplierCnpjMatch = -not [string]::IsNullOrWhiteSpace($officialNormalizedCnpj) -and (@($movementCnpjs | Where-Object { $_ -eq $officialNormalizedCnpj }).Count -gt 0)
            if ($hasSupplierAliasMismatch -and -not $hasSupplierCnpjMatch) {
                $officialDivergenceTypes += 'supplier_mismatch'
                Add-ContractCrossIssue `
                    -Store $divergenceStore `
                    -Type 'supplier_mismatch' `
                    -Severity 'warning' `
                    -CrossKey $crossKey `
                    -ContractNumber ([string]$official.contractNumber) `
                    -PortalContractId ([string]$official.portalContractId) `
                    -Title 'Fornecedor divergente entre portal e Diario' `
                    -Reason ("Portal: $([string]$official.contractor) | Diario: $($movementContractors -join ' | ')") `
                    -PublishedAt ([string]$(if ($latestMovementAt) { $latestMovementAt } else { $official.publishedAt })) `
                    -SourceView 'official'
            }
            elseif ($hasSupplierAliasMismatch -and $hasSupplierCnpjMatch) {
                Add-ContractCounterValue -Counter $suppressedDivergenceCounter -Key 'fornecedor_compativel_por_cnpj'
            }

            $officialValue = if ($null -ne $official.valueNumber) { [double]$official.valueNumber } else { 0.0 }
            if ($officialValue -gt 0 -and @($movementValues).Count -gt 0) {
                $tolerance = [Math]::Max(($officialValue * 0.05), 100.0)
                $compatibleValue = @($movementValues | Where-Object { [Math]::Abs([double]$_ - $officialValue) -le $tolerance }).Count -gt 0
                if (-not $compatibleValue) {
                    $officialDivergenceTypes += 'value_mismatch'
                    Add-ContractCrossIssue `
                        -Store $divergenceStore `
                        -Type 'value_mismatch' `
                        -Severity 'warning' `
                        -CrossKey $crossKey `
                        -ContractNumber ([string]$official.contractNumber) `
                        -PortalContractId ([string]$official.portalContractId) `
                        -Title 'Valor divergente entre portal e Diario' `
                        -Reason ("Portal: $officialValue | Diario: $(@($movementValues | Select-Object -First 3) -join ' | ')") `
                        -PublishedAt ([string]$(if ($latestMovementAt) { $latestMovementAt } else { $official.publishedAt })) `
                        -SourceView 'official'
                }
            }
        }
        elseif ($crossKey) {
            if (Test-ShouldMaterializeOfficialWithoutDiary -Official $official -Vigency $vigency) {
                $officialDivergenceTypes += 'official_without_diary'
                Add-ContractCrossIssue `
                    -Store $divergenceStore `
                    -Type 'official_without_diary' `
                    -Severity $(if ([bool]$vigency.isActive) { 'critical' } else { 'warning' }) `
                    -CrossKey $crossKey `
                    -ContractNumber ([string]$official.contractNumber) `
                    -PortalContractId ([string]$official.portalContractId) `
                    -Title 'Contrato oficial sem ato correspondente no Diario' `
                    -Reason $reason `
                    -PublishedAt ([string]$official.publishedAt) `
                    -SourceView 'official'
            }
            else {
                Add-ContractCounterValue -Counter $suppressedDivergenceCounter -Key 'contrato_historico_sem_ato_correspondente'
            }
        }

        if ($pendingReviewCount -gt 0 -and -not ($officialDivergenceTypes -contains 'pending_review')) {
            $officialDivergenceTypes += 'pending_review'
        }

        $officialStage = 'alerts'
        $officialAlerts = @()
        if ($status -eq 'pending_review') {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'pending_review'
                severity = 'warning'
                title = 'Vinculo depende de revisao manual'
                reason = $reason
            }
        }
        if ([bool]$vigency.isActive -and $null -ne $vigency.daysUntilEnd -and [int]$vigency.daysUntilEnd -le 60) {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'expiring_soon'
                severity = if ([int]$vigency.daysUntilEnd -le 15) { 'critical' } else { 'warning' }
                title = 'Contrato proximo do encerramento'
                reason = [string]$vigency.summaryLabel
            }
        }
        if ([bool]$vigency.isActive -and @($matchedMovements).Count -eq 0) {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'active_without_diary'
                severity = 'warning'
                title = 'Contrato vigente sem ato correspondente no Diario'
                reason = 'Nao ha ato vinculado do Diario para acompanhar a execucao atual.'
            }
        }
        elseif ([bool]$vigency.isActive -and (Get-DateTimestampValue -Text ([string]$latestMovementAt)) -lt (Get-Date).AddDays(-180).Ticks) {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'active_without_recent_diary'
                severity = 'warning'
                title = 'Contrato vigente sem movimentacao recente no Diario'
                reason = if ($latestMovementAt) { "Ultimo ato vinculado em $latestMovementAt." } else { 'Sem data de ato vinculada.' }
            }
        }
        if ([bool]$official.managementTracked -and (-not [bool]$official.hasManager -or -not [bool]$official.hasInspector)) {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'missing_management'
                severity = 'critical'
                title = 'Sem gestor ou fiscal atual identificado'
                reason = [string]$official.managementSummary
            }
        }
        if ([bool]$official.managerExonerationSignal -or [bool]$official.inspectorExonerationSignal) {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'exoneration_signal'
                severity = 'critical'
                title = 'Ha sinal de exoneracao nos responsaveis atuais'
                reason = [string]$official.managementSummary
            }
        }
        if ([string]::IsNullOrWhiteSpace([string]$official.localPdfRelative)) {
            $officialAlerts += [pscustomobject][ordered]@{
                type = 'missing_document'
                severity = 'warning'
                title = 'Cadastro oficial sem documento local'
                reason = 'Nao ha PDF local vinculado para conferencia.'
            }
        }

        foreach ($alert in @($officialAlerts)) {
            $alerts += [pscustomobject][ordered]@{
                type = [string]$alert.type
                severity = [string]$alert.severity
                crossKey = $crossKey
                contractNumber = [string]$official.contractNumber
                portalContractId = [string]$official.portalContractId
                title = [string]$alert.title
                reason = [string]$alert.reason
                publishedAt = [string]$(if ($latestMovementAt) { $latestMovementAt } else { $official.publishedAt })
            }
        }

        $officialStage = 'persist'
        $official | Add-Member -NotePropertyName crossSource -NotePropertyValue ([pscustomobject][ordered]@{
            crossKey = $crossKey
            itemKey = $officialItemKey
            matched = (@($matchedMovements).Count -gt 0)
            status = $status
            confidence = $confidence
            confidenceScore = [int]$confidenceScore
            reason = $reason
            candidateCount = [int]$duplicateCount
            matchedOfficialCount = 0
            matchedMovementCount = [int]@($matchedMovements).Count
            officialPortalContractIds = @([string]$official.portalContractId)
            officialNumbers = @([string]$official.contractNumber)
            officialStatuses = @([string]$official.portalStatus)
            latestCounterpartAt = [string]$latestMovementAt
            reviewRequired = ($status -eq 'pending_review')
            reviewedBy = $null
            reviewedAt = $null
            divergenceCount = [int]@($officialDivergenceTypes).Count
            divergenceTypes = @($officialDivergenceTypes)
            topTypes = @($matchedMovements | ForEach-Object { [string]$_.type } | Where-Object { $_ } | Select-Object -Unique -First 3)
        }) -Force
        $official | Add-Member -NotePropertyName vigency -NotePropertyValue ([pscustomobject]$vigency) -Force
        $official | Add-Member -NotePropertyName operationalAlerts -NotePropertyValue @($officialAlerts) -Force

        $officialStage = 'movement_sync'
        foreach ($movement in @($matchedMovements)) {
            $types = @()
            foreach ($type in @($officialDivergenceTypes)) {
                if (-not ($types -contains [string]$type)) {
                    $types += [string]$type
                }
            }
            if ($movement.crossSource.status -eq 'pending_review' -and -not ($types -contains 'pending_review')) {
                $types += 'pending_review'
            }
            $movement.crossSource.divergenceCount = [int]@($types).Count
            $movement.crossSource.divergenceTypes = @($types)
        }
        }
        catch {
            throw "Falha ao consolidar o contrato oficial ${officialItemKey} na etapa ${officialStage}: $($_.Exception.Message)"
        }
    }

    foreach ($movement in @($ContractMovements | Where-Object { $_.crossSource -and -not [bool]$_.crossSource.matched })) {
        if ($movement.crossSource.status -eq 'pending_review') {
            $movement.crossSource.divergenceCount = [int]@('pending_review').Count
            $movement.crossSource.divergenceTypes = @('pending_review')
        }
        elseif ($movement.crossSource.status -eq 'unmatched') {
            $movementCrossKey = [string]$movement.crossSource.crossKey
            if ($movementCrossKey -and $materialMovementWithoutOfficialKeys.ContainsKey($movementCrossKey)) {
                $movement.crossSource.divergenceCount = [int]@('movement_without_official').Count
                $movement.crossSource.divergenceTypes = @('movement_without_official')
            }
            else {
                $movement.crossSource.divergenceCount = 0
                $movement.crossSource.divergenceTypes = @()
            }
        }
    }

    $reviewQueueItems = @($reviewQueue | Sort-Object @{ Expression = { [int]$_.recommendedScore }; Descending = $true }, @{ Expression = { [string]$_.publishedAt }; Descending = $true })
    $divergenceItems = @(
        Convert-ContractCrossIssueStoreToItems -Store $divergenceStore |
        Sort-Object @{ Expression = { [string]$_.severity }; Descending = $false }, @{ Expression = { [string]$_.publishedAt }; Descending = $true }
    )
    $alertItems = @($alerts | Sort-Object @{ Expression = { [string]$_.severity }; Descending = $false }, @{ Expression = { [string]$_.publishedAt }; Descending = $true })
    $suppressionSummary = Convert-ContractCounterToSummary -Counter $suppressedDivergenceCounter
    $officialMatched = @($OfficialContracts | Where-Object { $_.crossSource -and [bool]$_.crossSource.matched }).Count
    $officialPendingReview = @($OfficialContracts | Where-Object { $_.crossSource -and [string]$_.crossSource.status -eq 'pending_review' }).Count
    $movementMatched = @($ContractMovements | Where-Object { $_.crossSource -and [bool]$_.crossSource.matched }).Count
    $movementPendingReview = @($ContractMovements | Where-Object { $_.crossSource -and [string]$_.crossSource.status -eq 'pending_review' }).Count

    return [ordered]@{
        officialContracts = @($OfficialContracts)
        contractMovements = @($ContractMovements)
        reviewQueue = @($reviewQueueItems)
        divergences = @($divergenceItems)
        alerts = @($alertItems)
        suppressedDivergences = $suppressionSummary
        summary = [ordered]@{
            officialMatched = [int]$officialMatched
            officialUnmatched = [int][Math]::Max(@($OfficialContracts).Count - $officialMatched - $officialPendingReview, 0)
            officialPendingReview = [int]$officialPendingReview
            movementMatched = [int]$movementMatched
            movementUnmatched = [int][Math]::Max(@($ContractMovements).Count - $movementMatched - $movementPendingReview, 0)
            movementPendingReview = [int]$movementPendingReview
            automaticMatches = [int]$automaticMatches
            reviewedMatches = [int]$reviewedMatches
            divergences = [int]@($divergenceItems).Count
            suppressedDivergences = [int]$suppressionSummary.total
            operationalAlerts = [int]@($alertItems).Count
            latestMatchedAt = if ($latestMatchedAt -gt 0) { ([DateTime]::new([Int64]$latestMatchedAt)).ToString('s') } else { $null }
        }
    }
}

function Refresh-ContractsAggregate {
    $diariesById = Get-DiariesById
    $analysisFiles = @(Get-ChildItem -LiteralPath $script:AnalysisRoot -Filter '*.json' -File -ErrorAction SilentlyContinue)
    $portalContractsPayload = Get-PortalContractsPayload
    $items = New-Object System.Collections.Generic.List[object]
    $movementItems = New-Object System.Collections.Generic.List[object]
    $officialItems = New-Object System.Collections.Generic.List[object]
    $analyses = New-Object System.Collections.Generic.List[object]
    $supplierSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $typeCounter = @{}
    $organizationCounter = @{}
    $personnelExonerationIndex = Get-PersonnelExonerationIndex
    $totalValue = 0.0
    $confirmedItems = 0
    $procurementItems = 0
    $managementItems = 0
    $flaggedItems = 0
    $highConfidenceItems = 0
    $portalOfficialItems = 0

    foreach ($file in $analysisFiles) {
        $analysis = Read-JsonFile -Path $file.FullName -Default $null
        if ($null -eq $analysis) {
            continue
        }

        if ([string]$analysis.parserVersion -ne $script:ParserVersion) {
            continue
        }

        $diaryId = [string]$analysis.diaryId
        $diary = if ($diariesById.ContainsKey($diaryId)) { $diariesById[$diaryId] } else { $null }
        $analysisItems = @($analysis.items)

        $analyses.Add([pscustomobject]@{
            diaryId = $diaryId
            edition = if ($diary) { [string]$diary.edition } else { $null }
            publishedAt = if ($diary) { [string]$diary.publishedAt } else { $null }
            analyzedAt = [string]$analysis.analyzedAt
            parserVersion = [string]$analysis.parserVersion
            itemCount = $analysisItems.Count
            totalValue = (Convert-BrazilianCurrencyToNumber -Text ([string]$analysis.summary.totalValue))
            summary = $analysis.summary
        })

        foreach ($item in $analysisItems) {
            $valueNumber = Convert-BrazilianCurrencyToNumber -Text ([string]$item.value)
            $totalValue += $valueNumber

            $contractor = [string]$item.contractor
            if (-not [string]::IsNullOrWhiteSpace($contractor)) {
                $null = $supplierSet.Add($contractor.Trim())
            }

            $recordClass = [string]$item.recordClass
            switch ($recordClass) {
                'execucao_contratual' { $confirmedItems++ }
                'licitacao_ou_contratacao' { $procurementItems++ }
                'gestao_contratual' { $managementItems++ }
            }

            if (@($item.flags).Count -gt 0) {
                $flaggedItems++
            }

            if ([string]$item.confidenceLabel -eq 'alta') {
                $highConfidenceItems++
            }

            $itemType = [string]$item.type
            if (-not [string]::IsNullOrWhiteSpace($itemType)) {
                if (-not $typeCounter.ContainsKey($itemType)) {
                    $typeCounter[$itemType] = 0
                }
                $typeCounter[$itemType]++
            }

            $organizationId = [string]$item.primaryOrganizationId
            if (-not [string]::IsNullOrWhiteSpace($organizationId)) {
                if (-not $organizationCounter.ContainsKey($organizationId)) {
                    $organizationCounter[$organizationId] = [ordered]@{
                        organizationId = $organizationId
                        name = [string]$item.primaryOrganizationName
                        sphere = [string]$item.primaryOrganizationSphere
                        areaId = [string]$item.primaryOrganizationAreaId
                        count = 0
                        totalValue = 0.0
                    }
                }

                $organizationCounter[$organizationId].count++
                $organizationCounter[$organizationId].totalValue += $valueNumber
            }

            $movementItem = [pscustomobject]@{
                diaryId = $diaryId
                edition = if ($diary) { [string]$diary.edition } else { $null }
                publishedAt = if ($diary) { [string]$diary.publishedAt } else { $null }
                isExtra = if ($diary) { [bool]$diary.isExtra } else { $false }
                pageCount = if ($diary) { [int]$diary.pageCount } else { 0 }
                localPdfRelative = if ($diary) { [string]$diary.localPdfRelative } else { $null }
                webPdfPath = if ($diary) { [string]$diary.webPdfPath } else { $null }
                viewUrl = if ($diary) { [string]$diary.viewUrl } else { $null }
                candidateKeywords = if ($diary) { @($diary.candidateKeywords) } else { @($analysis.keywords) }
                type = [string]$item.type
                pageNumber = [int]$item.pageNumber
                contractNumber = [string]$item.contractNumber
                processNumber = [string]$item.processNumber
                modality = [string]$item.modality
                contractor = $contractor
                cnpj = [string]$item.cnpj
                object = [string]$item.object
                value = [string]$item.value
                valueNumber = $valueNumber
                actTitle = [string]$item.actTitle
                recordClass = $recordClass
                recordClassLabel = [string]$item.recordClassLabel
                confidenceLabel = [string]$item.confidenceLabel
                confidenceScore = [int]$item.confidenceScore
                term = [string]$item.term
                signatureDate = [string]$item.signatureDate
                legalBasis = [string]$item.legalBasis
                primaryOrganizationId = [string]$item.primaryOrganizationId
                primaryOrganizationName = [string]$item.primaryOrganizationName
                primaryOrganizationSphere = [string]$item.primaryOrganizationSphere
                primaryOrganizationAreaId = [string]$item.primaryOrganizationAreaId
                mentionedOrganizationIds = @($item.mentionedOrganizationIds)
                mentionedOrganizationNames = @($item.mentionedOrganizationNames)
                excerpt = [string]$item.excerpt
                completeness = [string]$item.completeness
                flags = @($item.flags)
            }

            $movementItems.Add($movementItem)
            $items.Add($movementItem)
        }
    }

    foreach ($item in @($portalContractsPayload.items)) {
        $valueNumber = if ($null -ne $item.valueNumber) { [double]$item.valueNumber } else { (Convert-BrazilianCurrencyToNumber -Text ([string]$item.value)) }
        $totalValue += $valueNumber
        $portalOfficialItems++

        $contractor = [string]$item.contractor
        if (-not [string]::IsNullOrWhiteSpace($contractor)) {
            $null = $supplierSet.Add($contractor.Trim())
        }

        $confirmedItems++

        if (@($item.flags).Count -gt 0) {
            $flaggedItems++
        }

        if ([string]$item.confidenceLabel -eq 'alta') {
            $highConfidenceItems++
        }

        $itemType = [string]$item.type
        if (-not [string]::IsNullOrWhiteSpace($itemType)) {
            if (-not $typeCounter.ContainsKey($itemType)) {
                $typeCounter[$itemType] = 0
            }
            $typeCounter[$itemType]++
        }

        $organizationId = [string]$item.primaryOrganizationId
        if (-not [string]::IsNullOrWhiteSpace($organizationId)) {
            if (-not $organizationCounter.ContainsKey($organizationId)) {
                $organizationCounter[$organizationId] = [ordered]@{
                    organizationId = $organizationId
                    name = [string]$item.primaryOrganizationName
                    sphere = [string]$item.primaryOrganizationSphere
                    areaId = [string]$item.primaryOrganizationAreaId
                    count = 0
                    totalValue = 0.0
                }
            }

            $organizationCounter[$organizationId].count++
            $organizationCounter[$organizationId].totalValue += $valueNumber
        }

        $officialItem = [pscustomobject]@{
            sourceType = [string]$item.sourceType
            sourceLabel = [string]$item.sourceLabel
            portalContractId = [string]$item.portalContractId
            diaryId = $null
            edition = [string]$item.edition
            publishedAt = [string]$item.publishedAt
            isExtra = $false
            pageCount = 0
            localPdfRelative = [string]$item.localPdfRelative
            webPdfPath = [string]$item.webPdfPath
            viewUrl = [string]$item.viewUrl
            candidateKeywords = @($item.candidateKeywords)
            type = [string]$item.type
            pageNumber = 1
            contractNumber = [string]$item.contractNumber
            processNumber = [string]$item.processNumber
            modality = [string]$item.modality
            contractor = $contractor
            cnpj = [string]$item.cnpj
            object = [string]$item.object
            value = [string]$item.value
            valueNumber = $valueNumber
            actTitle = [string]$item.actTitle
            recordClass = [string]$item.recordClass
            recordClassLabel = [string]$item.recordClassLabel
            confidenceLabel = [string]$item.confidenceLabel
            confidenceScore = [int]$item.confidenceScore
            term = [string]$item.term
            signatureDate = [string]$item.signatureDate
            legalBasis = [string]$item.legalBasis
            primaryOrganizationId = [string]$item.primaryOrganizationId
            primaryOrganizationName = [string]$item.primaryOrganizationName
            primaryOrganizationSphere = [string]$item.primaryOrganizationSphere
            primaryOrganizationAreaId = [string]$item.primaryOrganizationAreaId
            mentionedOrganizationIds = @($item.mentionedOrganizationIds)
            mentionedOrganizationNames = @($item.mentionedOrganizationNames)
            excerpt = [string]$item.excerpt
            completeness = [string]$item.completeness
            flags = @($item.flags)
            portalStatus = [string]$item.portalStatus
            portalOrigin = [string]$item.portalOrigin
            revenueExpense = [string]$item.revenueExpense
            updatedAt = [string]$item.updatedAt
            aditiveCount = [int]$(if ($null -ne $item.aditiveCount) { $item.aditiveCount } else { 0 })
        }

        $officialItems.Add($officialItem)
        $items.Add($officialItem)
    }

    $typeSummary = @(
        $typeCounter.GetEnumerator() |
        Sort-Object -Property Value -Descending |
        ForEach-Object {
            [pscustomobject]@{
                type = [string]$_.Key
                count = [int]$_.Value
            }
        }
    )

    $organizationSummary = @(
        $organizationCounter.GetEnumerator() |
        ForEach-Object { [pscustomobject]$_.Value } |
        Sort-Object -Property @{ Expression = { $_.count }; Descending = $true }, @{ Expression = { $_.name }; Descending = $false }
    )

    $managementProfiles = @(Build-ContractManagementProfiles -Items $movementItems.ToArray() -PersonnelExonerationIndex $personnelExonerationIndex)
    $managementProfileIndex = @{}
    foreach ($profile in $managementProfiles) {
        foreach ($token in @($profile.referenceTokens)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
                $managementProfileIndex[[string]$token] = $profile
            }
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$profile.contractKey)) {
            $managementProfileIndex[[string]$profile.contractKey] = $profile
        }
    }

    $enrichedOfficialContracts = @(
        $officialItems.ToArray() |
        ForEach-Object { Add-ManagementFieldsToItem -Item $_ -ProfileIndex $managementProfileIndex } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.contractNumber }; Descending = $false }
    )
    $enrichedContractMovements = @(
        $movementItems.ToArray() |
        ForEach-Object { Add-ManagementFieldsToItem -Item $_ -ProfileIndex $managementProfileIndex } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.pageNumber }; Descending = $false }
    )
    $crossReviewPayload = Get-ContractCrossReviewPayload
    $crosswalk = Resolve-ContractCrosswalk -OfficialContracts $enrichedOfficialContracts -ContractMovements $enrichedContractMovements -ReviewPayload $crossReviewPayload
    $enrichedOfficialContracts = @($crosswalk.officialContracts)
    $enrichedContractMovements = @($crosswalk.contractMovements)
    $enrichedItems = @(
        @($enrichedOfficialContracts) + @($enrichedContractMovements) |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.pageNumber }; Descending = $false }
    )
    $managementSummary = Get-ManagementSummaryFromProfiles -Profiles $managementProfiles
    $financialMonitoring = Get-ContractsFinancialMonitoringSummary -OfficialContracts $enrichedOfficialContracts
    foreach ($item in @($enrichedItems)) {
        $item | Add-Member -NotePropertyName searchIndex -NotePropertyValue (Get-ContractSearchIndex -Item $item) -Force
    }

    $aggregate = [ordered]@{
        generatedAt = (Get-IsoNow)
        parserVersion = $script:ParserVersion
        totalItems = $enrichedItems.Count
        totalValue = [math]::Round($totalValue, 2)
        analyzedDiaryCount = $analyses.Count
        officialPortalContracts = $portalOfficialItems
        uniqueSuppliers = $supplierSet.Count
        qualitySummary = [ordered]@{
            confirmedItems = $confirmedItems
            procurementItems = $procurementItems
            managementItems = $managementItems
            flaggedItems = $flaggedItems
            highConfidenceItems = $highConfidenceItems
        }
        typeSummary = $typeSummary
        organizationSummary = $organizationSummary
        managementSummary = $managementSummary
        crosswalkSummary = $crosswalk.summary
        analyses = @($analyses | Sort-Object publishedAt -Descending)
        managementProfiles = @($managementProfiles)
        crossReviewQueue = @($crosswalk.reviewQueue)
        crossSourceDivergences = @($crosswalk.divergences)
        crossSourceAlerts = @($crosswalk.alerts)
        crossSourceSuppressionSummary = $crosswalk.suppressedDivergences
        financialMonitoring = $financialMonitoring
        officialContracts = @($enrichedOfficialContracts)
        contractMovements = @($enrichedContractMovements)
        items = @($enrichedItems)
    }

    Write-JsonFile -Path $script:ContractsPath -Data $aggregate
    Update-WorkspaceAggregateSnapshot -Aggregate $aggregate | Out-Null

    $pendingCount = @(Get-PendingAnalysisTasks -Limit 5000).Count
    Update-Status -Updates @{
        analyzedDiaries = $analyses.Count
        pendingAnalysis = $pendingCount
    }

    return $aggregate
}

function Get-SyncLock {
    $lock = Read-JsonFile -Path $script:SyncLockPath -Default $null
    if ($null -eq $lock) {
        return $null
    }

    $startedAt = $null
    try {
        if ($lock.startedAt) {
            $startedAt = [DateTime]::Parse([string]$lock.startedAt)
        }
    }
    catch {
        $startedAt = $null
    }

    if ($startedAt -and $startedAt -lt (Get-Date).AddHours(-12)) {
        Remove-Item -LiteralPath $script:SyncLockPath -Force -ErrorAction SilentlyContinue
        return $null
    }

    return $lock
}

function Set-SyncLock {
    Write-JsonFile -Path $script:SyncLockPath -Data ([ordered]@{
        startedAt = (Get-IsoNow)
        processId = $PID
    })
}

function Clear-SyncLock {
    Remove-Item -LiteralPath $script:SyncLockPath -Force -ErrorAction SilentlyContinue
}

function Start-DetachedSyncProcess {
    $lock = Get-SyncLock
    if ($lock) {
        return $false
    }

    $syncScript = Join-Path $PSScriptRoot 'sync.ps1'
    $argumentList = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        "`"$syncScript`""
    )

    Start-Process -FilePath 'powershell.exe' -ArgumentList $argumentList -WorkingDirectory $script:AppRoot -WindowStyle Hidden | Out-Null
    return $true
}

function Get-DashboardOverviewSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DiariesPayload,

        [Parameter(Mandatory = $true)]
        [object]$ContractsPayload
    )

    $contracts = @($ContractsPayload.items)
    $diaries = @($DiariesPayload.diaries)
    $analyses = @($ContractsPayload.analyses)
    $qualitySummary = if ($null -ne $ContractsPayload.qualitySummary) { $ContractsPayload.qualitySummary } else { [ordered]@{} }
    $supplierStats = @{}
    $organizationStats = @{}
    $recentCutoff = (Get-Date).AddDays(-30)

    $withValue = 0
    $withOrganization = 0
    $withMainNumber = 0
    $withoutValue = 0
    $withoutOrganization = 0
    $withoutObject = 0
    $withoutMainNumber = 0
    $recentContracts = 0
    $latestPublishedAt = $null

    foreach ($item in $contracts) {
        $valueNumber = 0.0
        try {
            $valueNumber = [double]$item.valueNumber
        }
        catch {
            $valueNumber = 0.0
        }

        if ($valueNumber -gt 0) {
            $withValue += 1
        }
        else {
            $withoutValue += 1
        }

        $organizationName = [string]$item.primaryOrganizationName
        if ([string]::IsNullOrWhiteSpace($organizationName)) {
            $withoutOrganization += 1
        }
        else {
            $withOrganization += 1
            $organizationStats[$organizationName] = [int]$(if ($organizationStats.ContainsKey($organizationName)) { $organizationStats[$organizationName] + 1 } else { 1 })
        }

        $mainNumber = if (-not [string]::IsNullOrWhiteSpace([string]$item.contractNumber)) {
            [string]$item.contractNumber
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$item.processNumber)) {
            [string]$item.processNumber
        }
        else {
            ''
        }

        if ([string]::IsNullOrWhiteSpace($mainNumber)) {
            $withoutMainNumber += 1
        }
        else {
            $withMainNumber += 1
        }

        if ([string]::IsNullOrWhiteSpace([string]$item.object)) {
            $withoutObject += 1
        }

        $publishedAtRaw = [string]$item.publishedAt
        if (-not [string]::IsNullOrWhiteSpace($publishedAtRaw)) {
            try {
                $publishedAt = [DateTime]::Parse($publishedAtRaw)
                if ($publishedAt -ge $recentCutoff) {
                    $recentContracts += 1
                }
                if ($null -eq $latestPublishedAt -or $publishedAt -gt $latestPublishedAt) {
                    $latestPublishedAt = $publishedAt
                }
            }
            catch {
                # Mantem o resumo resiliente mesmo com datas fora do padrao.
            }
        }

        $supplierName = [string]$item.contractor
        if (-not [string]::IsNullOrWhiteSpace($supplierName)) {
            if (-not $supplierStats.ContainsKey($supplierName)) {
                $supplierStats[$supplierName] = [ordered]@{
                    name = $supplierName
                    count = 0
                    total = 0.0
                }
            }

            $supplierEntry = $supplierStats[$supplierName]
            $supplierEntry.count = [int]$supplierEntry.count + 1
            $supplierEntry.total = [math]::Round(([double]$supplierEntry.total + $valueNumber), 2)
        }
    }

    $topOrganization = $null
    if ($organizationStats.Count -gt 0) {
        $topOrganization = @(
            $organizationStats.GetEnumerator() |
            Sort-Object -Property @{ Expression = { $_.Value }; Descending = $true }, @{ Expression = { $_.Key }; Descending = $false } |
            Select-Object -First 1 |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Key
                    count = [int]$_.Value
                }
            }
        )[0]
    }

    $topSupplier = $null
    if ($supplierStats.Count -gt 0) {
        $topSupplier = @(
            $supplierStats.Values |
            Sort-Object -Property @{ Expression = { $_.count }; Descending = $true }, @{ Expression = { $_.total }; Descending = $true }, @{ Expression = { $_.name }; Descending = $false } |
            Select-Object -First 1 |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.name
                    count = [int]$_.count
                    total = [double]$_.total
                }
            }
        )[0]
    }

    $contractsCount = [Math]::Max($contracts.Count, 1)
    $candidateDiaryCount = @($diaries | Where-Object { @($_.candidateKeywords).Count -gt 0 }).Count
    $extractedDiaryCount = @($analyses | Where-Object { [int]$_.itemCount -gt 0 }).Count

    return [ordered]@{
        generatedAt = if ($ContractsPayload.generatedAt) { [string]$ContractsPayload.generatedAt } elseif ($DiariesPayload.generatedAt) { [string]$DiariesPayload.generatedAt } else { $null }
        parserVersion = if ($ContractsPayload.parserVersion) { [string]$ContractsPayload.parserVersion } else { $script:ParserVersion }
        totalItems = [int]$ContractsPayload.totalItems
        totalValue = [double]$ContractsPayload.totalValue
        analyzedDiaryCount = [int]$ContractsPayload.analyzedDiaryCount
        uniqueSuppliers = [int]$ContractsPayload.uniqueSuppliers
        confirmedItems = [int]$qualitySummary.confirmedItems
        procurementItems = [int]$qualitySummary.procurementItems
        managementItems = [int]$qualitySummary.managementItems
        flaggedItems = [int]$qualitySummary.flaggedItems
        highConfidenceItems = [int]$qualitySummary.highConfidenceItems
        diariesCount = [int]$diaries.Count
        analysesCount = [int]$analyses.Count
        candidateDiaryCount = [int]$candidateDiaryCount
        extractedDiaryCount = [int]$extractedDiaryCount
        withValue = [int]$withValue
        withoutValue = [int]$withoutValue
        withOrganization = [int]$withOrganization
        withoutOrganization = [int]$withoutOrganization
        withMainNumber = [int]$withMainNumber
        withoutMainNumber = [int]$withoutMainNumber
        withoutObject = [int]$withoutObject
        recentContracts = [int]$recentContracts
        latestPublishedAt = if ($latestPublishedAt) { $latestPublishedAt.ToString('o') } else { $null }
        topOrganization = $topOrganization
        topOrganizationShare = if ($topOrganization) { [math]::Round(($topOrganization.count / $contractsCount), 4) } else { 0 }
        topSupplier = $topSupplier
        topSupplierShare = if ($topSupplier) { [math]::Round(($topSupplier.count / $contractsCount), 4) } else { 0 }
    }
}

function Get-DashboardAreaSummary {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Contracts = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$OrganizationCatalog = $null,

        [Parameter(Mandatory = $false)]
        [int]$Limit = 4
    )

    $areaLabels = @{}
    foreach ($area in @($OrganizationCatalog.areas)) {
        $areaId = Collapse-Whitespace -Text ([string]$area.id)
        if ([string]::IsNullOrWhiteSpace($areaId)) {
            continue
        }
        $areaLabels[$areaId] = [string]$(if ([string]$area.label) { $area.label } else { $areaId })
    }

    $areaCounter = @{}
    foreach ($item in @($Contracts)) {
        $areaId = Collapse-Whitespace -Text ([string]$item.primaryOrganizationAreaId)
        if ([string]::IsNullOrWhiteSpace($areaId)) {
            continue
        }

        if (-not $areaCounter.ContainsKey($areaId)) {
            $areaCounter[$areaId] = [ordered]@{
                areaId = $areaId
                label = if ($areaLabels.ContainsKey($areaId)) { [string]$areaLabels[$areaId] } else { $areaId }
                count = 0
            }
        }

        $areaCounter[$areaId].count = [int]$areaCounter[$areaId].count + 1
    }

    return @(
        $areaCounter.Values |
        Sort-Object @{ Expression = { [int]$_.count }; Descending = $true }, @{ Expression = { [string]$_.label }; Descending = $false } |
        Select-Object -First $Limit |
        ForEach-Object {
            [pscustomobject][ordered]@{
                areaId = [string]$_.areaId
                label = [string]$_.label
                count = [int]$_.count
            }
        }
    )
}

function Get-DashboardSupplierSummary {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Contracts = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 8
    )

    return @(
        @(Get-AggregateSnapshotSupplierRows -Contracts $Contracts -Limit $Limit) |
        ForEach-Object {
            [pscustomobject][ordered]@{
                name = [string]$_.name
                count = [int]$_.count
                totalValue = [double]$_.totalValue
            }
        }
    )
}

function Get-DashboardDiaryPreview {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Diaries = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Analyses = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 6
    )

    $analysisLookup = @{}
    foreach ($analysis in @($Analyses)) {
        $analysisLookup[[string]$analysis.diaryId] = $analysis
    }

    return @(
        @($Diaries) |
        Sort-Object @{ Expression = { [string]$_.publishedAt }; Descending = $true }, @{ Expression = { [int]$(if ($_.edition) { $_.edition } else { 0 }) }; Descending = $true } |
        Select-Object -First $Limit |
        ForEach-Object {
            $analysis = if ($analysisLookup.ContainsKey([string]$_.id)) { $analysisLookup[[string]$_.id] } else { $null }
            [pscustomobject][ordered]@{
                id = [string]$_.id
                edition = [string]$_.edition
                isExtra = [bool]$_.isExtra
                publishedAt = [string]$_.publishedAt
                pageCount = [int]$_.pageCount
                sizeLabel = [string]$_.sizeLabel
                candidateKeywords = @($_.candidateKeywords)
                pdfDownloadedAt = [string]$_.pdfDownloadedAt
                webPdfPath = [string]$_.webPdfPath
                viewUrl = [string]$_.viewUrl
                previewItemCount = if ($analysis) { [int]$analysis.itemCount } else { $null }
            }
        }
    )
}

function Get-DashboardAnalysisAlertPreview {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Diaries = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Analyses = @(),

        [Parameter(Mandatory = $false)]
        [int]$Limit = 8
    )

    $diaryLookup = @{}
    foreach ($diary in @($Diaries)) {
        $diaryLookup[[string]$diary.id] = $diary
    }

    $alerts = New-Object System.Collections.Generic.List[object]
    foreach ($analysis in @($Analyses)) {
        $diaryId = [string]$analysis.diaryId
        if (-not $diaryLookup.ContainsKey($diaryId)) {
            continue
        }

        $diary = $diaryLookup[$diaryId]
        $hasCandidateKeywords = @($diary.candidateKeywords).Count -gt 0
        $itemCount = if ($null -ne $analysis.itemCount) { [int]$analysis.itemCount } else { 0 }
        $errorMessage = ''
        if ($null -ne $analysis -and (Test-ObjectProperty -Item $analysis -Name 'summary') -and $null -ne $analysis.summary -and (Test-ObjectProperty -Item $analysis.summary -Name 'error')) {
            $errorMessage = Collapse-Whitespace -Text ([string]$analysis.summary.error)
        }
        if ([string]::IsNullOrWhiteSpace($errorMessage) -and (-not $hasCandidateKeywords -or $itemCount -gt 0)) {
            continue
        }

        $alerts.Add([pscustomobject][ordered]@{
            diaryId = $diaryId
            edition = [string]$diary.edition
            isExtra = [bool]$diary.isExtra
            publishedAt = [string]$diary.publishedAt
            message = if (-not [string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage } else { 'Sem atos extraidos apesar de marcador contratual.' }
        }) | Out-Null
    }

    return @(
        @($alerts.ToArray()) |
        Sort-Object @{ Expression = { [string]$_.publishedAt }; Descending = $true }, @{ Expression = { [int]$(if ($_.edition) { $_.edition } else { 0 }) }; Descending = $true } |
        Select-Object -First $Limit
    )
}

function Find-ContractEntryByReference {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reference,

        [Parameter(Mandatory = $true)]
        [object]$ContractsPayload
    )

    @(
        @($ContractsPayload.officialContracts) + @($ContractsPayload.contractMovements) |
        Where-Object {
            Test-ReferenceMatch -Reference $Reference -CandidateValues @(
                [string]$_.referenceKey,
                [string]$_.managementProfileKey,
                [string]$_.contractNumber,
                [string]$_.processNumber,
                [string]$_.portalContractId
            )
        } |
        Sort-Object -Property @{ Expression = { [string]$_.publishedAt }; Descending = $true }
    ) | Select-Object -First 1
}

function Convert-ContractItemToWorkspaceSummary {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [string]$Reference
    )

    if ($null -eq $Item) {
        return [ordered]@{
            reference = $Reference
            title = $Reference
            organization = ''
            contractor = ''
            publishedAt = $null
            portalStatus = ''
            value = ''
            detailUrl = "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($Reference))"
            sourceType = ''
        }
    }

    return [ordered]@{
        reference = if ([string]::IsNullOrWhiteSpace([string]$Item.referenceKey)) { $Reference } else { [string]$Item.referenceKey }
        title = if ([string]::IsNullOrWhiteSpace([string]$Item.actTitle)) { [string]$(if ([string]$Item.contractNumber) { $Item.contractNumber } else { $Reference }) } else { [string]$Item.actTitle }
        organization = [string]$Item.primaryOrganizationName
        contractor = [string]$Item.contractor
        publishedAt = [string]$Item.publishedAt
        portalStatus = [string]$Item.portalStatus
        value = [string]$Item.value
        detailUrl = "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($(if ([string]::IsNullOrWhiteSpace([string]$Item.referenceKey)) { $Reference } else { [string]$Item.referenceKey })) )"
        sourceType = [string]$Item.recordClass
    }
}

function Merge-WorkspaceAlertEntry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$BaseAlert,

        [Parameter(Mandatory = $true)]
        [object]$WorkspacePayload
    )

    $stringifyValue = {
        param([AllowNull()][object]$Value)
        if ($null -eq $Value) {
            return ''
        }
        if ($Value -is [string]) {
            return [string]$Value
        }
        if ($Value -is [System.Collections.IEnumerable]) {
            return [string]::Join(' | ', @($Value | ForEach-Object { [string]$_ }))
        }
        return [string]$Value
    }

    $baseType = if (Test-ObjectProperty -Item $BaseAlert -Name 'type') { & $stringifyValue (Get-ObjectStringValue -Item $BaseAlert -Name 'type') } else { '' }
    $baseReference = if (Test-ObjectProperty -Item $BaseAlert -Name 'reference') { & $stringifyValue (Get-ObjectStringValue -Item $BaseAlert -Name 'reference') } else { '' }
    $baseTitle = if (Test-ObjectProperty -Item $BaseAlert -Name 'title') { & $stringifyValue (Get-ObjectStringValue -Item $BaseAlert -Name 'title') } else { '' }
    $existingAlertKey = if (Test-ObjectProperty -Item $BaseAlert -Name 'alertKey') { Get-ObjectStringValue -Item $BaseAlert -Name 'alertKey' } else { '' }
    $alertKey = if ([string]::IsNullOrWhiteSpace($existingAlertKey)) {
        $normalizedKeyParts = @(
            Normalize-IndexText -Text ([string]$baseType)
            Normalize-IndexText -Text ([string]$baseReference)
            Normalize-IndexText -Text ([string]$baseTitle)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
        ([string]::Join('|', $normalizedKeyParts) -replace '\s+', '')
    }
    else {
        $existingAlertKey
    }
    $workspaceAlertStates = if (Test-ObjectProperty -Item $WorkspacePayload -Name 'alertStates') {
        @($WorkspacePayload.alertStates)
    }
    else {
        @()
    }
    $state = @($workspaceAlertStates | Where-Object { [string]$_.alertKey -eq [string]$alertKey } | Select-Object -First 1) | Select-Object -First 1
    $merged = [ordered]@{}
    foreach ($property in @(Get-ObjectPropertyEntries -Item $BaseAlert)) {
        $merged[$property.Name] = $property.Value
    }
    $merged['alertKey'] = $alertKey
    $merged['stateStatus'] = if ($state) { [string]$state.status } else { 'novo' }
    $merged['stateAssigneeLogin'] = if ($state) { [string]$state.assigneeLogin } else { '' }
    $merged['stateAssigneeName'] = if ($state) { [string]$state.assigneeName } else { '' }
    $merged['stateDueDate'] = if ($state) { [string]$state.dueDate } else { $null }
    $merged['stateSnoozeUntil'] = if ($state) { [string]$state.snoozeUntil } else { $null }
    $merged['stateJustification'] = if ($state) { [string]$state.justification } else { '' }
    $merged['stateUpdatedAt'] = if ($state) { [string]$state.updatedAt } else { '' }
    $merged['stateUpdatedBy'] = if ($state) { [string]$state.updatedBy } else { '' }
    $merged['stateHistory'] = if ($state) { @($state.history) } else { @() }
    $merged['isResolved'] = [bool]($state -and [string]$state.status -eq 'resolvido')
    $merged['isSnoozed'] = [bool]($state -and [string]$state.status -eq 'adiado' -and -not [string]::IsNullOrWhiteSpace([string]$state.snoozeUntil) -and [DateTime]::Parse([string]$state.snoozeUntil) -ge (Get-Date).Date)
    return [pscustomobject]$merged
}

function Get-WorkspaceDashboardData {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [object]$ContractsPayload
    )

    $workspacePayload = Get-WorkspacePayload
    $capabilities = Get-RoleCapabilities -Role ([string]$User.role)
    $roleKey = ([string]$User.role).Trim().ToLowerInvariant()
    $favoriteEntries = @($workspacePayload.favorites | Where-Object { [string]$_.userId -eq [string]$User.id })
    $savedViews = @($workspacePayload.savedViews | Where-Object { [string]$_.userId -eq [string]$User.id })
    $workflowItems = if ([bool]$capabilities.canManageWorkflow) {
        @($workspacePayload.workflowItems)
    }
    else {
        @($workspacePayload.workflowItems | Where-Object {
            [string]$_.assigneeUserId -eq [string]$User.id -or
            [string]$_.createdBy -eq [string]$User.login
        })
    }

    $favorites = @(
        $favoriteEntries |
        ForEach-Object {
            $reference = [string]$_.reference
            [pscustomobject](Convert-ContractItemToWorkspaceSummary -Item (Find-ContractEntryByReference -Reference $reference -ContractsPayload $ContractsPayload) -Reference $reference)
        }
    )
    $favoriteReferenceLookup = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($favorite in @($favorites)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$favorite.reference)) {
            $null = $favoriteReferenceLookup.Add((Normalize-IndexText -Text ([string]$favorite.reference)))
        }
    }

    $activityLog = if ([bool]$capabilities.canSeeActivityLog) {
        @($workspacePayload.activityLog)
    }
    else {
        @($workspacePayload.activityLog | Where-Object { [string]$_.createdBy -eq [string]$User.login })
    }

    $supportPayload = Get-SupportPayload
    $supportTickets = if ([bool]$capabilities.canSeeAllSupport) {
        @($supportPayload.tickets)
    }
    else {
        @($supportPayload.tickets | Where-Object {
            [string]$_.requesterUserId -eq [string]$User.id -or
            [string]$_.assigneeUserId -eq [string]$User.id
        })
    }

    $alertEntries = New-Object System.Collections.Generic.List[object]
    foreach ($alert in @($ContractsPayload.crossSourceAlerts | Sort-Object @{ Expression = { [string]$_.severity }; Descending = $false }, @{ Expression = { [string]$_.publishedAt }; Descending = $true } | Select-Object -First 10)) {
        $reference = [string]$(if ([string]$alert.crossKey) { $alert.crossKey } else { $alert.contractNumber })
        $alertEntries.Add([pscustomobject][ordered]@{
            type = 'contract_alert'
            severity = [string]$alert.severity
            title = [string]$alert.title
            summary = [string]$alert.reason
            reference = $reference
            href = if ([string]::IsNullOrWhiteSpace($reference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($reference))" }
            sourceLabel = 'Cruzamento contratual'
            dueDate = $null
            createdAt = [string]$alert.publishedAt
        })
    }

    foreach ($workflowItem in @($workflowItems | Where-Object { [string]$_.status -notin @('regularizado', 'encerrado') } | Sort-Object @{ Expression = { [string]$_.dueDate }; Descending = $false }, @{ Expression = { [string]$_.updatedAt }; Descending = $true } | Select-Object -First 8)) {
        $alertEntries.Add([pscustomobject][ordered]@{
            type = 'workflow'
            severity = if ([string]$workflowItem.dueDate -and [DateTime]::Parse([string]$workflowItem.dueDate) -lt (Get-Date)) { 'critical' } else { 'warning' }
            title = "Workflow $([string]$workflowItem.status -replace '_', ' ')"
            summary = [string]$(if ([string]$workflowItem.assigneeName) { "Responsavel: $([string]$workflowItem.assigneeName)." } else { 'Sem responsavel definido.' })
            reference = [string]$workflowItem.reference
            href = if ([string]::IsNullOrWhiteSpace([string]$workflowItem.reference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode([string]$workflowItem.reference))" }
            sourceLabel = 'Workflow interno'
            dueDate = [string]$workflowItem.dueDate
            createdAt = [string]$workflowItem.updatedAt
        })
    }

    foreach ($ticket in @($supportTickets | Where-Object { [string]$_.status -notin @('autorizado', 'concluido') } | Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true } | Select-Object -First 8)) {
        $alertEntries.Add([pscustomobject][ordered]@{
            type = 'support'
            severity = if ([string]$ticket.priority -eq 'alta') { 'critical' } else { 'warning' }
            title = [string]$ticket.subject
            summary = "Suporte em $([string]$ticket.status -replace '_', ' ')."
            reference = [string]$ticket.id
            href = '/suporte.html'
            sourceLabel = 'Suporte interno'
            dueDate = [string]$ticket.dueDate
            createdAt = [string]$ticket.updatedAt
        })
    }

    $alertsCenter = @(
        $alertEntries |
        ForEach-Object { [pscustomobject](Merge-WorkspaceAlertEntry -BaseAlert $_ -WorkspacePayload $workspacePayload) } |
        Where-Object { -not [bool]$_.isResolved -and -not [bool]$_.isSnoozed } |
        Sort-Object @{ Expression = { if ([string]$_.severity -eq 'critical') { 0 } else { 1 } }; Descending = $false }, @{ Expression = { if ([string]$_.stateDueDate) { [string]$_.stateDueDate } elseif ([string]$_.dueDate) { [string]$_.dueDate } else { '9999-12-31' } }; Descending = $false }, @{ Expression = { [string]$_.createdAt }; Descending = $true }
    )

    $expiringContracts = @(
        @($ContractsPayload.officialContracts) |
        Where-Object {
            [bool]$_.vigency.isActive -and $null -ne $_.vigency.daysUntilEnd -and [int]$_.vigency.daysUntilEnd -le 30
        } |
        Sort-Object @{ Expression = { [int]$_.vigency.daysUntilEnd }; Descending = $false }, @{ Expression = { [double]$_.valueNumber }; Descending = $true } |
        Select-Object -First 8 |
        ForEach-Object {
            [pscustomobject][ordered]@{
                reference = [string]$(if ([string]$_.referenceKey) { $_.referenceKey } else { $_.contractNumber })
                title = [string]$(if ([string]$_.contractNumber) { $_.contractNumber } else { $_.referenceKey })
                organization = [string]$_.primaryOrganizationName
                contractor = [string]$_.contractor
                value = [string]$_.value
                daysUntilEnd = [int]$_.vigency.daysUntilEnd
                summary = [string]$_.vigency.summaryLabel
                href = "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($(if ([string]$_.referenceKey) { [string]$_.referenceKey } else { [string]$_.contractNumber })) )"
            }
        }
    )

    $unassignedAlertsCount = @($alertsCenter | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.stateAssigneeLogin) }).Count
    $assignedAlertsCount = @($alertsCenter | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.stateAssigneeLogin) }).Count
    $assignedToMeAlertsCount = @($alertsCenter | Where-Object { [string]$_.stateAssigneeLogin -eq [string]$User.login }).Count
    $overdueAlertsCount = @($alertsCenter | Where-Object {
        $effectiveDueDate = if ([string]$_.stateDueDate) { [string]$_.stateDueDate } else { [string]$_.dueDate }
        -not [string]::IsNullOrWhiteSpace($effectiveDueDate) -and [DateTime]::Parse($effectiveDueDate) -lt (Get-Date).Date
    }).Count
    $alertsDueTodayCount = @($alertsCenter | Where-Object {
        $effectiveDueDate = if ([string]$_.stateDueDate) { [string]$_.stateDueDate } else { [string]$_.dueDate }
        -not [string]::IsNullOrWhiteSpace($effectiveDueDate) -and [DateTime]::Parse($effectiveDueDate).Date -eq (Get-Date).Date
    }).Count
    $alertsDueThisWeekCount = @($alertsCenter | Where-Object {
        $effectiveDueDate = if ([string]$_.stateDueDate) { [string]$_.stateDueDate } else { [string]$_.dueDate }
        if ([string]::IsNullOrWhiteSpace($effectiveDueDate)) {
            $false
        }
        else {
            $dueDate = [DateTime]::Parse($effectiveDueDate).Date
            ($dueDate -ge (Get-Date).Date -and $dueDate -le (Get-Date).Date.AddDays(7))
        }
    }).Count
    $criticalAlertsCount = @($alertsCenter | Where-Object { [string]$_.severity -eq 'critical' }).Count
    $activeWorkflowItems = @($workflowItems | Where-Object { [string]$_.status -notin @('regularizado', 'encerrado') })
    $assignedWorkflowCount = @($activeWorkflowItems | Where-Object { [string]$_.assigneeLogin -eq [string]$User.login }).Count
    $overdueWorkflowCount = @($activeWorkflowItems | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_.dueDate) -and [DateTime]::Parse([string]$_.dueDate).Date -lt (Get-Date).Date
    }).Count
    $favoriteExpiringCount = @($expiringContracts | Where-Object { $favoriteReferenceLookup.Contains((Normalize-IndexText -Text ([string]$_.reference))) }).Count
    $reviewQueueCount = [int]@($ContractsPayload.crossReviewQueue).Count
    $crossDivergenceCount = [int]@($ContractsPayload.crossSourceDivergences).Count
    $openSupportCount = [int]@($supportTickets | Where-Object { [string]$_.status -notin @('autorizado', 'concluido') }).Count
    $recentChanges = Get-WorkspaceAggregateChangeSummary

    $notificationsBase = New-Object System.Collections.Generic.List[object]
    foreach ($alert in @($alertsCenter | Select-Object -First 10)) {
        if (
            [string]$alert.stateAssigneeLogin -eq [string]$User.login -or
            [string]$alert.severity -eq 'critical' -or
            ([bool]$capabilities.canManageAlerts -and [string]::IsNullOrWhiteSpace([string]$alert.stateAssigneeLogin))
        ) {
            $notificationsBase.Add([pscustomobject][ordered]@{
                type = 'alert_notification'
                tone = if ([string]$alert.severity -eq 'critical') { 'critical' } else { 'warning' }
                title = [string]$alert.title
                summary = [string]$alert.summary
                reference = [string]$alert.reference
                href = [string]$alert.href
                sourceLabel = 'Central operacional'
                createdAt = [string]$(if ([string]$alert.stateUpdatedAt) { $alert.stateUpdatedAt } else { $alert.createdAt })
            })
        }
    }
    foreach ($workflowItem in @($activeWorkflowItems | Sort-Object @{ Expression = { [string]$_.dueDate }; Descending = $false }, @{ Expression = { [string]$_.updatedAt }; Descending = $true } | Select-Object -First 8)) {
        if ([string]$workflowItem.assigneeLogin -eq [string]$User.login -or [bool]$capabilities.canManageWorkflow) {
            $notificationsBase.Add([pscustomobject][ordered]@{
                type = 'workflow_notification'
                tone = if ([string]$workflowItem.dueDate -and [DateTime]::Parse([string]$workflowItem.dueDate).Date -lt (Get-Date).Date) { 'critical' } else { 'warning' }
                title = "Workflow $([string]$workflowItem.status -replace '_', ' ')"
                summary = [string]$(if ([string]$workflowItem.note) { $workflowItem.note } elseif ([string]$workflowItem.assigneeName) { "Responsavel: $([string]$workflowItem.assigneeName)." } else { 'Sem responsavel definido.' })
                reference = [string]$workflowItem.reference
                href = if ([string]::IsNullOrWhiteSpace([string]$workflowItem.reference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode([string]$workflowItem.reference))" }
                sourceLabel = 'Workflow interno'
                createdAt = [string]$workflowItem.updatedAt
            })
        }
    }
    foreach ($ticket in @($supportTickets | Where-Object { [string]$_.status -notin @('autorizado', 'concluido') } | Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true } | Select-Object -First 6)) {
        if (
            [string]$ticket.assigneeLogin -eq [string]$User.login -or
            [string]$ticket.requesterUserId -eq [string]$User.id -or
            [bool]$capabilities.canSeeAllSupport
        ) {
            $notificationsBase.Add([pscustomobject][ordered]@{
                type = 'support_notification'
                tone = if ([string]$ticket.priority -eq 'alta') { 'critical' } else { 'info' }
                title = [string]$ticket.subject
                summary = "Suporte em $([string]$ticket.status -replace '_', ' ')."
                reference = [string]$ticket.id
                href = '/suporte.html'
                sourceLabel = 'Suporte interno'
                createdAt = [string]$ticket.updatedAt
            })
        }
    }
    if ([bool]$capabilities.canReviewContracts) {
        foreach ($review in @($ContractsPayload.crossReviewQueue | Select-Object -First 6)) {
            $notificationsBase.Add([pscustomobject][ordered]@{
                type = 'review_notification'
                tone = 'critical'
                title = [string]$(if ([string]$review.movementReference) { $review.movementReference } else { $review.movementTitle })
                summary = "Revisao manual pendente com $([int]$review.candidateCount) candidato(s)."
                reference = [string]$review.crossKey
                href = if ([string]::IsNullOrWhiteSpace([string]$review.crossKey)) { '/contratos.html?quick=pendingReview' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode([string]$review.crossKey))" }
                sourceLabel = 'Revisao de vinculos'
                createdAt = [string]$review.publishedAt
            })
        }
    }
    foreach ($contract in @($expiringContracts | Select-Object -First 5)) {
        if ($roleKey -eq 'viewer' -or $favoriteReferenceLookup.Contains((Normalize-IndexText -Text ([string]$contract.reference)))) {
            $notificationsBase.Add([pscustomobject][ordered]@{
                type = 'expiring_contract'
                tone = if ([int]$contract.daysUntilEnd -le 15) { 'critical' } else { 'warning' }
                title = [string]$contract.title
                summary = [string]$contract.summary
                reference = [string]$contract.reference
                href = [string]$contract.href
                sourceLabel = 'Vigencia contratual'
                createdAt = Get-IsoNow
            })
        }
    }

    $baseObservability = Get-ObservabilityDashboardData
    if ([bool]$capabilities.canSeeObservability -and $baseObservability.latestError) {
        $notificationsBase.Add([pscustomobject][ordered]@{
            type = 'sync_notification'
            tone = 'critical'
            title = 'Falha recente de sincronizacao'
            summary = [string]$(if ([string]$baseObservability.latestError.message) { $baseObservability.latestError.message } else { 'A ultima sincronizacao registrou erro.' })
            reference = ''
            href = '/index.html'
            sourceLabel = 'Observabilidade'
            createdAt = [string]$baseObservability.latestError.createdAt
        })
    }
    if ([bool]$capabilities.canManageFinancialAutomation) {
        $notificationsBase.Add([pscustomobject][ordered]@{
            type = 'financial_automation'
            tone = if ([int]$ContractsPayload.financialMonitoring.automationReadySources -gt 0) { 'info' } else { 'warning' }
            title = 'Cobertura financeira oficial'
            summary = [string]$ContractsPayload.financialMonitoring.note
            reference = ''
            href = '/contratos.html'
            sourceLabel = 'Financeiro oficial'
            createdAt = [string]$ContractsPayload.generatedAt
        })
    }
    foreach ($change in @($recentChanges.contractVersionChanges | Select-Object -First 5)) {
        $tone = if ([string]$change.changeKind -eq 'removed' -or [int]$change.severityScore -ge 95) {
            'critical'
        }
        elseif ([string]$change.changeKind -eq 'updated' -or [int]$change.severityScore -ge 55) {
            'warning'
        }
        else {
            'info'
        }
        $notificationsBase.Add([pscustomobject][ordered]@{
            type = 'change_notification'
            tone = $tone
            title = [string]$(if ([string]$change.title) { $change.title } else { $change.reference })
            summary = [string]$change.summary
            reason = [string]$change.reason
            nextStep = [string]$change.nextStep
            categoryLabel = 'Mudanca entre sincronizacoes'
            managerialLabel = [string]$(if ([string]$change.changeKind -eq 'new') { 'Novo contrato' } elseif ([string]$change.changeKind -eq 'removed') { 'Saiu da carteira' } else { 'Contrato alterado' })
            priorityScore = [int]$change.severityScore
            reference = [string]$change.reference
            href = [string]$(if ([string]$change.href) { $change.href } else { '/contratos.html' })
            sourceLabel = 'Historico da base'
            createdAt = [string]$(if ([string]$recentChanges.currentGeneratedAt) { $recentChanges.currentGeneratedAt } else { $ContractsPayload.generatedAt })
        })
    }

    $notificationsCenter = @(
        $notificationsBase |
        ForEach-Object { [pscustomobject](Merge-WorkspaceNotificationEntry -BaseNotification $_ -WorkspacePayload $workspacePayload -User $User) } |
        Where-Object { -not [bool]$_.isArchived } |
        Sort-Object `
            @{ Expression = { if (-not [bool]$_.isRead) { 0 } else { 1 } }; Descending = $false }, `
            @{ Expression = { [int]$_.priorityScore }; Descending = $true }, `
            @{ Expression = { if ([string]$_.effectiveDueDate) { [string]$_.effectiveDueDate } else { '9999-12-31' } }; Descending = $false }, `
            @{ Expression = { [string]$_.createdAt }; Descending = $true }
    )
    $unreadNotificationsCount = @($notificationsCenter | Where-Object { -not [bool]$_.isRead }).Count
    $criticalNotificationsCount = @($notificationsCenter | Where-Object { [string]$_.tone -eq 'critical' }).Count
    $notificationsDueTodayCount = @($notificationsCenter | Where-Object { [string]$_.slaLabel -eq 'Prazo vence hoje' }).Count
    $staleNotificationsCount = @($notificationsCenter | Where-Object { [bool]$_.isStale }).Count
    $managerialNotificationsCount = @($notificationsCenter | Where-Object { [int]$_.priorityScore -ge 85 }).Count
    $changeNotificationsCount = @($notificationsCenter | Where-Object { [string]$_.type -eq 'change_notification' }).Count

    $decisionItems = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($notificationsCenter | Select-Object -First 8)) {
        $decisionItems.Add([pscustomobject][ordered]@{
            type = [string]$entry.type
            tone = if ([string]$entry.tone -eq 'critical') { 'critical' } else { 'warning' }
            title = [string]$entry.title
            summary = [string]$entry.summary
            meta = [string]$entry.sourceLabel
            href = [string]$entry.href
            reference = [string]$entry.reference
        })
    }

    $profileTitle = switch ($roleKey) {
        'admin' { 'Governanca da operacao' }
        'auditor' { 'Radar de auditoria' }
        'reviewer' { 'Fila de revisao do dia' }
        default { 'Leitura orientada do painel' }
    }
    $profileSummary = switch ($roleKey) {
        'admin' { 'Priorize filas criticas, falhas de sincronizacao e itens sem responsavel definido.' }
        'auditor' { 'Concentre revisoes manuais, divergencias materiais e contratos com risco operacional.' }
        'reviewer' { 'Trate os vinculos pendentes, os workflows atribuidos e as notificacoes pessoais antes de abrir novos casos.' }
        default { 'Use favoritos, visoes salvas e notificacoes para acompanhar os contratos que mais afetam sua rotina.' }
    }

    $focusCards = switch ($roleKey) {
        'admin' {
            @(
                [pscustomobject][ordered]@{ key = 'critical_alerts'; label = 'Alertas criticos'; value = [int]$criticalAlertsCount; meta = 'Ocorrencias que pedem acao imediata.' }
                [pscustomobject][ordered]@{ key = 'overdue_queue'; label = 'Fila em atraso'; value = [int]($overdueAlertsCount + $overdueWorkflowCount); meta = 'Alertas e workflow ja vencidos.' }
                [pscustomobject][ordered]@{ key = 'manual_reviews'; label = 'Revisoes pendentes'; value = [int]$reviewQueueCount; meta = 'Vinculos ambíguos aguardando decisao.' }
                [pscustomobject][ordered]@{ key = 'unread_notifications'; label = 'Notificacoes novas'; value = [int]$unreadNotificationsCount; meta = 'Pendencias pessoais ainda nao tratadas.' }
            )
        }
        'auditor' {
            @(
                [pscustomobject][ordered]@{ key = 'cross_reviews'; label = 'Revisoes manuais'; value = [int]$reviewQueueCount; meta = 'Casos com correspondencia ambigua.' }
                [pscustomobject][ordered]@{ key = 'divergences'; label = 'Divergencias'; value = [int]$crossDivergenceCount; meta = 'Diferencas materiais entre fontes oficiais.' }
                [pscustomobject][ordered]@{ key = 'overdue_alerts'; label = 'Alertas vencidos'; value = [int]$overdueAlertsCount; meta = 'Itens com prazo operacional ultrapassado.' }
                [pscustomobject][ordered]@{ key = 'open_support'; label = 'Suporte aberto'; value = [int]$openSupportCount; meta = 'Demandas internas ainda em andamento.' }
            )
        }
        'reviewer' {
            @(
                [pscustomobject][ordered]@{ key = 'assigned_workflows'; label = 'Workflow atribuido'; value = [int]$assignedWorkflowCount; meta = 'Contratos com voce como responsavel.' }
                [pscustomobject][ordered]@{ key = 'assigned_alerts'; label = 'Alertas comigo'; value = [int]$assignedToMeAlertsCount; meta = 'Central operacional atribuida ao seu login.' }
                [pscustomobject][ordered]@{ key = 'pending_reviews'; label = 'Vinculos para revisar'; value = [int]$reviewQueueCount; meta = 'Casos com fila manual de revisao.' }
                [pscustomobject][ordered]@{ key = 'expiring_soon'; label = 'Vigencias proximas'; value = [int]@($expiringContracts).Count; meta = 'Contratos ativos com prazo curto.' }
            )
        }
        default {
            @(
                [pscustomobject][ordered]@{ key = 'favorite_contracts'; label = 'Favoritos'; value = [int]@($favorites).Count; meta = 'Contratos marcados para acompanhamento.' }
                [pscustomobject][ordered]@{ key = 'saved_views'; label = 'Visoes salvas'; value = [int]@($savedViews).Count; meta = 'Consultas recorrentes prontas para uso.' }
                [pscustomobject][ordered]@{ key = 'favorite_expiring'; label = 'Favoritos vencendo'; value = [int]$favoriteExpiringCount; meta = 'Favoritos com vigencia proxima.' }
                [pscustomobject][ordered]@{ key = 'new_notifications'; label = 'Notificacoes novas'; value = [int]$unreadNotificationsCount; meta = 'Pendencias pessoais ainda nao lidas.' }
            )
        }
    }

    $quickActions = switch ($roleKey) {
        'admin' {
            @(
                [pscustomobject][ordered]@{ label = 'Abrir central de alertas'; href = '/index.html#alerts-center'; tone = 'primary' }
                [pscustomobject][ordered]@{ label = 'Ir para contratos'; href = '/contratos.html?quick=pendingReview'; tone = 'secondary' }
                [pscustomobject][ordered]@{ label = 'Ler o guia'; href = '/guia.html'; tone = 'ghost' }
            )
        }
        'auditor' {
            @(
                [pscustomobject][ordered]@{ label = 'Fila de revisao'; href = '/contratos.html?quick=pendingReview'; tone = 'primary' }
                [pscustomobject][ordered]@{ label = 'Painel de auditoria'; href = '/auditoria.html'; tone = 'secondary' }
                [pscustomobject][ordered]@{ label = 'Ler o guia'; href = '/guia.html'; tone = 'ghost' }
            )
        }
        'reviewer' {
            @(
                [pscustomobject][ordered]@{ label = 'Meus contratos'; href = '/contratos.html?quick=myAssignments'; tone = 'primary' }
                [pscustomobject][ordered]@{ label = 'Pendencias do dia'; href = '/index.html#notifications-panel'; tone = 'secondary' }
                [pscustomobject][ordered]@{ label = 'Ler o guia'; href = '/guia.html'; tone = 'ghost' }
            )
        }
        default {
            @(
                [pscustomobject][ordered]@{ label = 'Abrir favoritos'; href = '/index.html#favorite-contracts'; tone = 'primary' }
                [pscustomobject][ordered]@{ label = 'Ir para contratos'; href = '/contratos.html'; tone = 'secondary' }
                [pscustomobject][ordered]@{ label = 'Ler o guia'; href = '/guia.html'; tone = 'ghost' }
            )
        }
    }

    $executiveHome = [ordered]@{
        roleKey = $roleKey
        roleLabel = Get-RoleLabel -Role ([string]$User.role)
        profileTitle = $profileTitle
        profileSummary = $profileSummary
        focusCards = @($focusCards)
        decisionItems = @($decisionItems | Select-Object -First 10)
        expiringContracts = @($expiringContracts)
        quickActions = @($quickActions)
    }

    $notificationsSummary = [ordered]@{
        unread = [int]$unreadNotificationsCount
        critical = [int]$criticalNotificationsCount
        read = [int](@($notificationsCenter | Where-Object { [bool]$_.isRead }).Count)
        assignedToMe = [int]($assignedToMeAlertsCount + $assignedWorkflowCount)
        dueToday = [int]$notificationsDueTodayCount
        stale = [int]$staleNotificationsCount
        managerial = [int]$managerialNotificationsCount
        changedContracts = [int]$changeNotificationsCount
        digest = [string]$recentChanges.headline
    }

    $alertsSummary = [ordered]@{
        critical = [int]$criticalAlertsCount
        unassigned = [int]$unassignedAlertsCount
        assigned = [int]$assignedAlertsCount
        dueToday = [int]$alertsDueTodayCount
        dueThisWeek = [int]$alertsDueThisWeekCount
        overdue = [int]$overdueAlertsCount
    }

    $observability = [ordered]@{
        metrics = @(
            @($baseObservability.metrics) +
            @(
                [pscustomobject][ordered]@{
                    key = 'unread_notifications'
                    label = 'Notificacoes nao lidas'
                    value = [int]$unreadNotificationsCount
                    meta = 'Pendencias pessoais abertas no painel.'
                }
                [pscustomobject][ordered]@{
                    key = 'operational_overdue'
                    label = 'Pendencias vencidas'
                    value = [int]($overdueAlertsCount + $overdueWorkflowCount)
                    meta = 'Itens operacionais com prazo ja ultrapassado.'
                }
                [pscustomobject][ordered]@{
                    key = 'financial_coverage'
                    label = 'Fontes financeiras prontas'
                    value = "$([int]$ContractsPayload.financialMonitoring.automationReadySources)/$([int]$ContractsPayload.financialMonitoring.sourceCount)"
                    meta = 'Cobertura automatizada entre as fontes financeiras oficiais conhecidas.'
                }
            )
        )
        latestSync = $baseObservability.latestSync
        latestError = $baseObservability.latestError
        recentEvents = @($baseObservability.recentEvents | Select-Object -First 12)
    }
    $guide = Get-WorkspaceGuideData

    return [ordered]@{
        favorites = @($favorites)
        savedViews = @($savedViews)
        workflowItems = @(
            $workflowItems |
            Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true } |
            Select-Object -First 20
        )
        activityLog = @($activityLog | Select-Object -First 20)
        recentChanges = $recentChanges
        alertsCenter = @($alertsCenter | Select-Object -First 16)
        alertsSummary = $alertsSummary
        notificationsCenter = @($notificationsCenter | Select-Object -First 16)
        notificationsSummary = $notificationsSummary
        executiveHome = $executiveHome
        guide = $guide
        observability = $observability
        syncHistory = @((Get-StatusHashtable).syncHistory | Select-Object -First 12)
    }
}

function Get-WorkspaceSessionPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )

    $cacheKey = Get-WorkspaceSessionCacheKey -User $User
    if ($script:WorkspaceSessionPayloadCache.ContainsKey($cacheKey)) {
        Add-CacheMetric -CacheName 'workspace' -Metric 'hits'
        return $script:WorkspaceSessionPayloadCache[$cacheKey]
    }

    Add-CacheMetric -CacheName 'workspace' -Metric 'misses'

    $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
    $workspace = Get-WorkspaceDashboardData -User $User -ContractsPayload $contractsPayload
    $notes = if ([bool](Get-RoleCapabilities -Role ([string]$User.role)).canCommentContracts) {
        @((Get-WorkspacePayload).contractNotes | Sort-Object @{ Expression = { [string]$_.updatedAt }; Descending = $true } | Select-Object -First 60)
    }
    else {
        @()
    }
    $userCapabilities = Get-RoleCapabilities -Role ([string]$User.role)
    $assignableUsers = if ([bool]$userCapabilities.canManageSupport -or [bool]$userCapabilities.canManageWorkflow -or [bool]$userCapabilities.canManageAlerts) {
        @(Get-ManagedUsers | Where-Object { [string]$_.role -ne 'admin' })
    }
    else {
        @()
    }

    $payload = [ordered]@{
        generatedAt = Get-IsoNow
        apiContract = New-ApiContractDescriptor -Name 'workspace'
        apiContracts = Get-ApiContractCatalog
        roleCatalog = @(Get-RoleCatalog)
        assignableUsers = @($assignableUsers)
        favorites = @($workspace.favorites)
        savedViews = @($workspace.savedViews)
        workflowItems = @($workspace.workflowItems)
        activityLog = @($workspace.activityLog)
        recentChanges = $workspace.recentChanges
        alertsCenter = @($workspace.alertsCenter)
        alertsSummary = $workspace.alertsSummary
        notificationsCenter = @($workspace.notificationsCenter)
        notificationsSummary = $workspace.notificationsSummary
        executiveHome = $workspace.executiveHome
        guide = $workspace.guide
        observability = $workspace.observability
        syncHistory = @($workspace.syncHistory)
        contractNotes = @($notes)
    }

    Add-ScriptCacheEntry -Cache $script:WorkspaceSessionPayloadCache -Key $cacheKey -Value $payload -MaxEntries 24 -CacheName 'workspace'
    return $payload
}

function Get-SearchShortcutCatalog {
    @(
        [pscustomobject][ordered]@{ prefix = 'contrato:'; label = 'Contrato'; description = 'Foca em numero de contrato e dossie.' }
        [pscustomobject][ordered]@{ prefix = 'processo:'; label = 'Processo'; description = 'Prioriza numero de processo.' }
        [pscustomobject][ordered]@{ prefix = 'fornecedor:'; label = 'Fornecedor'; description = 'Busca por empresa ou contratado.' }
        [pscustomobject][ordered]@{ prefix = 'gestor:'; label = 'Gestor'; description = 'Procura gestor do contrato.' }
        [pscustomobject][ordered]@{ prefix = 'fiscal:'; label = 'Fiscal'; description = 'Procura fiscal do contrato.' }
        [pscustomobject][ordered]@{ prefix = 'alerta:'; label = 'Alerta'; description = 'Filtra a central operacional.' }
        [pscustomobject][ordered]@{ prefix = 'suporte:'; label = 'Suporte'; description = 'Foca chamados e fila interna.' }
        [pscustomobject][ordered]@{ prefix = 'usuario:'; label = 'Usuario'; description = 'Busca logins e perfis.' }
    )
}

function Get-SearchAliasVariants {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    $normalized = Normalize-IndexText -Text $Text
    $variants = New-Object 'System.Collections.Generic.HashSet[string]'
    if (-not [string]::IsNullOrWhiteSpace($normalized)) {
        $null = $variants.Add($normalized)
    }

    $withoutLegalSuffix = Collapse-Whitespace -Text (
        $normalized `
            -replace '\b(LTDA|EIRELI|MEI|ME|EPP|S/A|SA|SOCIEDADE ANONIMA|EMPRESA INDIVIDUAL)\b', ' ' `
            -replace '\bCOMERCIO\b', ' ' `
            -replace '\bSERVICOS\b', ' '
    )
    if (-not [string]::IsNullOrWhiteSpace($withoutLegalSuffix)) {
        $null = $variants.Add($withoutLegalSuffix)
    }

    $aliasPairs = [ordered]@{
        'sms' = 'secretaria municipal de saude'
        'saude' = 'secretaria municipal de saude'
        'sme' = 'secretaria municipal de educacao'
        'educacao' = 'secretaria municipal de educacao'
        'smf' = 'secretaria municipal da fazenda'
        'fazenda' = 'secretaria municipal da fazenda'
        'pref' = 'prefeitura municipal'
        'pm' = 'prefeitura municipal'
        'sec' = 'secretaria'
        'mun' = 'municipal'
    }

    foreach ($pair in $aliasPairs.GetEnumerator()) {
        $source = Normalize-IndexText -Text ([string]$pair.Key)
        $target = Normalize-IndexText -Text ([string]$pair.Value)
        if ($normalized -like "*$source*") {
            $null = $variants.Add(($normalized -replace [regex]::Escape($source), $target))
            $null = $variants.Add($target)
        }
        if ($normalized -like "*$target*") {
            $null = $variants.Add(($normalized -replace [regex]::Escape($target), $source))
            $null = $variants.Add($source)
        }
    }

    return @($variants)
}

function Get-EmptySearchIndexPayload {
    [ordered]@{
        generatedAt = $null
        version = $script:SearchIndexSchemaVersion
        sharedEntries = @()
        baseAlertEntries = @()
    }
}

function New-SearchDatasetEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$Aliases = @(),

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Title = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Subtitle = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Meta = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Reference = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Href = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Scope = '',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ExtraSearchText = ''
    )

    $aliasSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($alias in @($Aliases)) {
        $cleanAlias = Collapse-Whitespace -Text ([string]$alias)
        if (-not [string]::IsNullOrWhiteSpace($cleanAlias)) {
            $null = $aliasSet.Add($cleanAlias.ToLowerInvariant())
        }
    }

    $normalizedReference = Normalize-IndexText -Text $Reference
    $normalizedTitle = Normalize-IndexText -Text $Title
    $normalizedSubtitle = Normalize-IndexText -Text $Subtitle
    $normalizedMeta = Normalize-IndexText -Text $Meta
    $normalizedExtra = Normalize-IndexText -Text $ExtraSearchText

    $variantSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($sourceText in @($Reference, $Title, $Subtitle, $Meta, $ExtraSearchText)) {
        foreach ($variant in @(Get-SearchAliasVariants -Text ([string]$sourceText))) {
            if (-not [string]::IsNullOrWhiteSpace([string]$variant)) {
                $null = $variantSet.Add([string]$variant)
            }
        }
    }

    $haystackParts = @(
        $normalizedReference
        $normalizedTitle
        $normalizedSubtitle
        $normalizedMeta
        $normalizedExtra
        @($variantSet)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

    return [pscustomobject][ordered]@{
        type = Collapse-Whitespace -Text $Type
        title = Collapse-Whitespace -Text $Title
        subtitle = Collapse-Whitespace -Text $Subtitle
        meta = Collapse-Whitespace -Text $Meta
        reference = Collapse-Whitespace -Text $Reference
        href = Collapse-Whitespace -Text $Href
        scope = Collapse-Whitespace -Text $Scope
        aliases = @(@($aliasSet) | Sort-Object)
        normalizedReference = $normalizedReference
        normalizedTitle = $normalizedTitle
        normalizedSubtitle = $normalizedSubtitle
        normalizedMeta = $normalizedMeta
        haystack = Normalize-IndexText -Text ([string]::Join(' ', @($haystackParts)))
    }
}

function Get-SearchSavedViewHref {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Definition = $null
    )

    $queryParts = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Definition) {
        return '/contratos.html'
    }

    if ([string]$Definition.view -and [string]$Definition.view -ne 'official') { $queryParts.Add("view=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.view))") }
    if ([string]$Definition.quick -and [string]$Definition.quick -ne 'all') { $queryParts.Add("quick=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.quick))") }
    if ([string]$Definition.sort -and [string]$Definition.sort -ne 'recent') { $queryParts.Add("sort=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.sort))") }
    if ([string]$Definition.layout -and [string]$Definition.layout -ne 'cards') { $queryParts.Add("layout=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.layout))") }
    if ([string]$Definition.search) { $queryParts.Add("search=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.search))") }
    if ([string]$Definition.filterClass) { $queryParts.Add("class=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterClass))") }
    if ([string]$Definition.filterType) { $queryParts.Add("type=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterType))") }
    if ([string]$Definition.filterYear) { $queryParts.Add("year=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterYear))") }
    if ([string]$Definition.filterArea) { $queryParts.Add("area=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterArea))") }
    if ([string]$Definition.filterManagement) { $queryParts.Add("management=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterManagement))") }
    if ([string]$Definition.filterDocument) { $queryParts.Add("document=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterDocument))") }
    if ([string]$Definition.filterOrganization) { $queryParts.Add("organization=$([System.Web.HttpUtility]::UrlEncode([string]$Definition.filterOrganization))") }

    if ($queryParts.Count -gt 0) {
        return "/contratos.html?$([string]::Join('&', $queryParts))"
    }

    return '/contratos.html'
}

function Get-SharedContractSearchIndexEntries {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ContractsPayload
    )

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($ContractsPayload.officialContracts) + @($ContractsPayload.contractMovements)) {
        if ($null -eq $item) {
            continue
        }

        $aliases = New-Object System.Collections.Generic.List[string]
        $aliases.Add('contract') | Out-Null
        if (-not [string]::IsNullOrWhiteSpace([string]$item.processNumber)) { $aliases.Add('process') | Out-Null }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.contractor)) { $aliases.Add('supplier') | Out-Null }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.managerName)) { $aliases.Add('manager') | Out-Null }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.inspectorName)) { $aliases.Add('inspector') | Out-Null }

        $reference = [string]$(if ([string]$item.referenceKey) { $item.referenceKey } elseif ([string]$item.contractNumber) { $item.contractNumber } else { [string]$item.processNumber })
        $href = if ([string]::IsNullOrWhiteSpace($reference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($reference))" }
        $title = [string]$(if ([string]$item.actTitle) { $item.actTitle } elseif ([string]$item.contractNumber) { $item.contractNumber } else { [string]$item.processNumber })
        $subtitle = [string]$(if ([string]$item.contractor) { $item.contractor } else { $item.primaryOrganizationName })
        $scopeLabel = if ([bool]($item.PSObject.Properties.Match('portalContractId').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.portalContractId))) {
            'Contrato oficial'
        }
        else {
            'Movimentacao do Diario'
        }

        $searchEntry = New-SearchDatasetEntry `
            -Type 'contract' `
            -Aliases @($aliases.ToArray()) `
            -Title $title `
            -Subtitle $subtitle `
            -Meta ([string]$item.primaryOrganizationName) `
            -Reference $reference `
            -Href $href `
            -Scope $scopeLabel `
            -ExtraSearchText ([string]::Join(' ', @(
                [string]$item.object,
                [string]$item.primaryOrganizationName,
                [string]$item.managerName,
                [string]$item.inspectorName,
                [string]$item.processNumber,
                [string]$item.contractNumber,
                [string]$item.referenceKey
            )))
        $entries.Add($searchEntry) | Out-Null
    }

    return @($entries.ToArray())
}

function Get-SearchBaseAlertEntriesFromContractsPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ContractsPayload
    )

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($alert in @($ContractsPayload.crossSourceAlerts)) {
        if ($null -eq $alert) {
            continue
        }

        $reference = [string]$(if ([string]$alert.crossKey) { $alert.crossKey } elseif ([string]$alert.contractNumber) { $alert.contractNumber } else { [string]$alert.portalContractId })
        $entries.Add([pscustomobject][ordered]@{
            type = 'contract_alert'
            severity = [string]$alert.severity
            title = [string]$alert.title
            summary = [string]$alert.reason
            reference = $reference
            href = if ([string]::IsNullOrWhiteSpace($reference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($reference))" }
            sourceLabel = 'Cruzamento contratual'
            dueDate = $null
            createdAt = [string]$alert.publishedAt
        }) | Out-Null
    }

    return @($entries.ToArray())
}

function Save-SearchIndexPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $Payload.version = $script:SearchIndexSchemaVersion
    $Payload.sharedEntries = @($Payload.sharedEntries)
    $Payload.baseAlertEntries = @($Payload.baseAlertEntries)
    Write-JsonFile -Path $script:SearchIndexPath -Data $Payload
}

function Update-SearchIndexPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ContractsPayload
    )

    $payload = [ordered]@{
        generatedAt = [string]$ContractsPayload.generatedAt
        version = $script:SearchIndexSchemaVersion
        sharedEntries = @(Get-SharedContractSearchIndexEntries -ContractsPayload $ContractsPayload)
        baseAlertEntries = @(Get-SearchBaseAlertEntriesFromContractsPayload -ContractsPayload $ContractsPayload)
    }

    Save-SearchIndexPayload -Payload $payload
    return [pscustomobject]$payload
}

function Get-SearchIndexPayload {
    $payload = Read-JsonFile -Path $script:SearchIndexPath -Default $null
    $needsRebuild = $false

    if ($null -eq $payload) {
        $needsRebuild = $true
    }
    else {
        if ($payload -is [System.Collections.IDictionary]) {
            $convertedPayload = [ordered]@{}
            foreach ($entry in $payload.GetEnumerator()) {
                $convertedPayload[[string]$entry.Key] = $entry.Value
            }
            $payload = [pscustomobject]$convertedPayload
        }

        if (-not (Test-ObjectProperty -Item $payload -Name 'version') -or [string]$payload.version -ne $script:SearchIndexSchemaVersion) {
            $needsRebuild = $true
        }
        foreach ($propertyName in @('sharedEntries', 'baseAlertEntries')) {
            if (-not (Test-ObjectProperty -Item $payload -Name $propertyName) -or $null -eq $payload.$propertyName) {
                $needsRebuild = $true
                break
            }
        }
    }

    if ($needsRebuild) {
        $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
        if (
            @($contractsPayload.officialContracts).Count -gt 0 -or
            @($contractsPayload.contractMovements).Count -gt 0 -or
            @($contractsPayload.crossSourceAlerts).Count -gt 0
        ) {
            return (Update-SearchIndexPayload -ContractsPayload $contractsPayload)
        }

        return (Get-EmptySearchIndexPayload)
    }

    $payload.sharedEntries = @($payload.sharedEntries)
    $payload.baseAlertEntries = @($payload.baseAlertEntries)
    return $payload
}

function Get-SearchEntriesCacheKey {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )

    $parts = @(
        [string]$User.id
        [string]$User.role
        (Get-FileCacheStamp -Path $script:SearchIndexPath)
        (Get-FileCacheStamp -Path $script:WorkspaceStatePath)
        (Get-FileCacheStamp -Path $script:SupportPath)
        (Get-FileCacheStamp -Path $script:UsersPath)
    )

    return [string]::Join('|', $parts)
}

function Get-SearchEntriesForUser {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )

    if (-not (Test-Path -LiteralPath $script:SearchIndexPath)) {
        $null = Get-SearchIndexPayload
    }

    $cacheKey = Get-SearchEntriesCacheKey -User $User
    if ($script:SearchEntriesCache.ContainsKey($cacheKey)) {
        return $script:SearchEntriesCache[$cacheKey]
    }

    $searchIndexPayload = Get-SearchIndexPayload
    $workspacePayload = Get-WorkspacePayload
    $supportPayload = Get-SupportPayload
    $capabilities = Get-RoleCapabilities -Role ([string]$User.role)
    $entries = New-Object System.Collections.Generic.List[object]
    $contractLookup = @{}

    foreach ($entry in @($searchIndexPayload.sharedEntries)) {
        if ($null -eq $entry) {
            continue
        }

        $entries.Add($entry) | Out-Null

        $normalizedReference = if (Test-ObjectProperty -Item $entry -Name 'normalizedReference') {
            [string]$entry.normalizedReference
        }
        else {
            Normalize-IndexText -Text ([string]$entry.reference)
        }
        if ([string]$entry.type -eq 'contract' -and -not [string]::IsNullOrWhiteSpace($normalizedReference) -and -not $contractLookup.ContainsKey($normalizedReference)) {
            $contractLookup[$normalizedReference] = $entry
        }
    }

    foreach ($view in @($workspacePayload.savedViews | Where-Object { [string]$_.userId -eq [string]$User.id })) {
        if ($null -eq $view) {
            continue
        }

        $searchEntry = New-SearchDatasetEntry `
            -Type 'saved_view' `
            -Aliases @('saved_view') `
            -Title ([string]$view.name) `
            -Subtitle ("Visao salva da pagina $([string]$view.page)") `
            -Meta '' `
            -Reference ([string]$view.id) `
            -Href (Get-SearchSavedViewHref -Definition $view.definition) `
            -Scope 'Visao salva' `
            -ExtraSearchText ([string]::Join(' ', @([string]$view.page, [string]$view.userLogin)))
        $entries.Add($searchEntry) | Out-Null
    }

    foreach ($favorite in @($workspacePayload.favorites | Where-Object { [string]$_.userId -eq [string]$User.id })) {
        if ($null -eq $favorite) {
            continue
        }

        $reference = Collapse-Whitespace -Text ([string]$favorite.reference)
        $lookupKey = Normalize-IndexText -Text $reference
        $matchedEntry = if ($contractLookup.ContainsKey($lookupKey)) { $contractLookup[$lookupKey] } else { $null }
        $title = if ($matchedEntry) { [string]$matchedEntry.title } else { $reference }
        $subtitle = if ($matchedEntry) { [string]$matchedEntry.meta } else { '' }
        $meta = if ($matchedEntry) { [string]$matchedEntry.subtitle } else { '' }
        $href = if ($matchedEntry -and [string]$matchedEntry.href) { [string]$matchedEntry.href } elseif ([string]::IsNullOrWhiteSpace($reference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($reference))" }

        $searchEntry = New-SearchDatasetEntry `
            -Type 'favorite' `
            -Aliases @('favorite', 'contract') `
            -Title $title `
            -Subtitle $subtitle `
            -Meta $meta `
            -Reference $reference `
            -Href $href `
            -Scope 'Favorito' `
            -ExtraSearchText ([string]$matchedEntry.scope)
        $entries.Add($searchEntry) | Out-Null
    }

    foreach ($baseAlert in @($searchIndexPayload.baseAlertEntries)) {
        if ($null -eq $baseAlert) {
            continue
        }

        $alert = Merge-WorkspaceAlertEntry -BaseAlert $baseAlert -WorkspacePayload $workspacePayload
        if ([bool]$alert.isResolved -or [bool]$alert.isSnoozed) {
            continue
        }

        $alertHref = if (-not [string]::IsNullOrWhiteSpace([string]$alert.href)) { [string]$alert.href } elseif ([string]::IsNullOrWhiteSpace([string]$alert.reference)) { '/index.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode([string]$alert.reference))" }
        $searchEntry = New-SearchDatasetEntry `
            -Type 'alert' `
            -Aliases @('alert') `
            -Title ([string]$alert.title) `
            -Subtitle ([string]$alert.summary) `
            -Meta ([string]$alert.severity) `
            -Reference ([string]$alert.reference) `
            -Href $alertHref `
            -Scope 'Alerta' `
            -ExtraSearchText ([string]::Join(' ', @(
                [string]$alert.sourceLabel,
                [string]$alert.stateAssigneeName,
                [string]$alert.stateAssigneeLogin,
                [string]$alert.stateStatus
            )))
        $entries.Add($searchEntry) | Out-Null
    }

    $workflowItems = if ([bool]$capabilities.canManageWorkflow) {
        @($workspacePayload.workflowItems)
    }
    else {
        @($workspacePayload.workflowItems | Where-Object {
            [string]$_.assigneeUserId -eq [string]$User.id -or
            [string]$_.createdBy -eq [string]$User.login
        })
    }
    foreach ($workflowItem in @($workflowItems | Where-Object { [string]$_.status -notin @('regularizado', 'encerrado') })) {
        if ($null -eq $workflowItem) {
            continue
        }

        $workflowSummary = [string]$(if ([string]$workflowItem.note) { $workflowItem.note } elseif ([string]$workflowItem.assigneeName) { "Responsavel: $([string]$workflowItem.assigneeName)." } else { 'Workflow contratual atualizado.' })
        $workflowReference = [string]$workflowItem.reference
        $searchEntry = New-SearchDatasetEntry `
            -Type 'workflow' `
            -Aliases @('alert') `
            -Title ("Workflow $([string]$workflowItem.status -replace '_', ' ')") `
            -Subtitle $workflowSummary `
            -Meta ([string]$workflowItem.status) `
            -Reference $workflowReference `
            -Href $(if ([string]::IsNullOrWhiteSpace($workflowReference)) { '/contratos.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($workflowReference))" }) `
            -Scope 'Workflow interno' `
            -ExtraSearchText ([string]::Join(' ', @(
                [string]$workflowItem.assigneeName,
                [string]$workflowItem.assigneeLogin,
                [string]$workflowItem.createdBy
            )))
        $entries.Add($searchEntry) | Out-Null
    }

    $visibleTickets = if ([bool]$capabilities.canSeeAllSupport) {
        @($supportPayload.tickets)
    }
    else {
        @($supportPayload.tickets | Where-Object { [string]$_.requesterUserId -eq [string]$User.id -or [string]$_.assigneeUserId -eq [string]$User.id })
    }
    foreach ($ticket in @($visibleTickets)) {
        if ($null -eq $ticket) {
            continue
        }

        $searchEntry = New-SearchDatasetEntry `
            -Type 'support' `
            -Aliases @('support') `
            -Title ([string]$ticket.subject) `
            -Subtitle ([string]$ticket.requesterName) `
            -Meta ([string]$ticket.status) `
            -Reference ([string]$ticket.id) `
            -Href '/suporte.html' `
            -Scope 'Suporte' `
            -ExtraSearchText ([string]::Join(' ', @(
                [string]$ticket.message,
                [string]$ticket.adminResponse,
                [string]$ticket.assigneeName,
                [string]$ticket.priority
            )))
        $entries.Add($searchEntry) | Out-Null
    }

    if ([bool]$capabilities.canManageUsers) {
        foreach ($managedUser in @(Get-ManagedUsers)) {
            if ($null -eq $managedUser) {
                continue
            }

            $searchEntry = New-SearchDatasetEntry `
                -Type 'user' `
                -Aliases @('user') `
                -Title ([string]$managedUser.name) `
                -Subtitle ([string]$managedUser.roleLabel) `
                -Meta ("Login $([string]$managedUser.login)") `
                -Reference ([string]$managedUser.login) `
                -Href '/usuarios.html' `
                -Scope 'Usuario' `
                -ExtraSearchText ([string]$managedUser.role)
            $entries.Add($searchEntry) | Out-Null
        }
    }

    $resolvedEntries = @($entries.ToArray())
    Add-ScriptCacheEntry -Cache $script:SearchEntriesCache -Key $cacheKey -Value $resolvedEntries -MaxEntries 24
    return $resolvedEntries
}

function Get-SearchEntryScore {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$ScopeFilter,

        [Parameter(Mandatory = $true)]
        [string]$NormalizedQuery,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$NormalizedQueryVariants = @()
    )

    $aliases = @($Entry.aliases | ForEach-Object { [string]$_ })
    if ($ScopeFilter -ne 'all' -and $aliases -notcontains $ScopeFilter) {
        return -1
    }

    $haystack = [string]$Entry.haystack
    $matchedQuery = @(
        @($NormalizedQueryVariants) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and $haystack -like "*$_*" } |
        Select-Object -First 1
    ) | Select-Object -First 1

    if ([string]::IsNullOrWhiteSpace([string]$matchedQuery)) {
        return -1
    }

    $referenceValue = [string]$Entry.normalizedReference
    $titleValue = [string]$Entry.normalizedTitle
    $subtitleValue = [string]$Entry.normalizedSubtitle
    $metaValue = [string]$Entry.normalizedMeta
    $score = 20
    if ($referenceValue -eq $NormalizedQuery) { $score += 220 }
    elseif ($referenceValue -like "$NormalizedQuery*") { $score += 170 }
    elseif ($referenceValue -like "*$NormalizedQuery*") { $score += 120 }

    if ($titleValue -eq $NormalizedQuery) { $score += 180 }
    elseif ($titleValue -like "$NormalizedQuery*") { $score += 140 }
    elseif ($titleValue -like "*$NormalizedQuery*") { $score += 90 }

    if ($subtitleValue -like "$NormalizedQuery*") { $score += 80 }
    elseif ($subtitleValue -like "*$NormalizedQuery*") { $score += 55 }

    if ($metaValue -like "$NormalizedQuery*") { $score += 50 }
    elseif ($metaValue -like "*$NormalizedQuery*") { $score += 30 }

    if ($ScopeFilter -ne 'all') { $score += 25 }
    if ($NormalizedQuery.Length -le 6 -and $referenceValue -like "*$NormalizedQuery*") { $score += 20 }
    if ([string]$matchedQuery -ne $NormalizedQuery) { $score += 35 }
    return $score
}

function Get-PanelSearchPayload {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Query,

        [Parameter(Mandatory = $false)]
        [int]$Limit = 24
    )

    $cleanQuery = Collapse-Whitespace -Text $Query
    $shortcuts = @(Get-SearchShortcutCatalog)
    if ([string]::IsNullOrWhiteSpace($cleanQuery) -or $cleanQuery.Length -lt 2) {
        return [ordered]@{
            query = $cleanQuery
            scope = 'all'
            total = 0
            items = @()
            shortcuts = @($shortcuts)
            apiContract = New-ApiContractDescriptor -Name 'search'
            apiContracts = Get-ApiContractCatalog
        }
    }

    $scopeMap = [ordered]@{
        contrato = 'contract'
        processo = 'process'
        fornecedor = 'supplier'
        gestor = 'manager'
        fiscal = 'inspector'
        alerta = 'alert'
        suporte = 'support'
        usuario = 'user'
        favorito = 'favorite'
        visao = 'saved_view'
    }
    $scopeFilter = 'all'
    $searchText = $cleanQuery
    $scopeMatch = [regex]::Match($cleanQuery, '^(?<scope>contrato|processo|fornecedor|gestor|fiscal|alerta|suporte|usuario|favorito|visao)\s*:\s*(?<term>.+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($scopeMatch.Success) {
        $scopeFilter = [string]$scopeMap[$scopeMatch.Groups['scope'].Value.ToLowerInvariant()]
        $searchText = Collapse-Whitespace -Text ([string]$scopeMatch.Groups['term'].Value)
    }

    $normalizedQuery = Normalize-IndexText -Text $searchText
    $normalizedQueryVariants = @(Get-SearchAliasVariants -Text $searchText)
    if ([string]::IsNullOrWhiteSpace($normalizedQuery) -or $normalizedQuery.Length -lt 2) {
        return [ordered]@{
            query = $cleanQuery
            scope = $scopeFilter
            total = 0
            items = @()
            shortcuts = @($shortcuts)
            apiContract = New-ApiContractDescriptor -Name 'search'
            apiContracts = Get-ApiContractCatalog
        }
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @(Get-SearchEntriesForUser -User $User)) {
        if ($null -eq $entry) {
            continue
        }

        $score = Get-SearchEntryScore -Entry $entry -ScopeFilter $scopeFilter -NormalizedQuery $normalizedQuery -NormalizedQueryVariants $normalizedQueryVariants
        if ($score -lt 0) {
            continue
        }

        $results.Add([pscustomobject][ordered]@{
            type = [string]$entry.type
            title = [string]$entry.title
            subtitle = [string]$entry.subtitle
            meta = [string]$entry.meta
            reference = [string]$entry.reference
            href = [string]$entry.href
            scope = [string]$entry.scope
            score = [int]$score
        }) | Out-Null
    }

    $items = @(
        $results |
        Sort-Object @{ Expression = { [int]$_.score }; Descending = $true }, @{ Expression = { [string]$_.title }; Descending = $false } |
        Select-Object -First $Limit
    )

    Register-ObservabilityEvent -Type 'search' -Status $scopeFilter -Message "Busca executada para '$searchText'." -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        query = $cleanQuery
        normalizedQuery = $normalizedQuery
        scope = $scopeFilter
        total = [int]$results.Count
        indexedEntries = [int]$results.Count
    }) | Out-Null

    return [ordered]@{
        query = $cleanQuery
        scope = $scopeFilter
        total = [int]$results.Count
        items = @($items)
        shortcuts = @($shortcuts)
        apiContract = New-ApiContractDescriptor -Name 'search'
        apiContracts = Get-ApiContractCatalog
    }
}

function Search-PanelData {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [string]$Query,

        [Parameter(Mandatory = $false)]
        [int]$Limit = 24
    )

    $cleanQuery = Collapse-Whitespace -Text $Query
    $shortcuts = @(Get-SearchShortcutCatalog)
    if ([string]::IsNullOrWhiteSpace($cleanQuery) -or $cleanQuery.Length -lt 2) {
        return [ordered]@{
            query = $cleanQuery
            scope = 'all'
            total = 0
            items = @()
            shortcuts = @($shortcuts)
            apiContract = New-ApiContractDescriptor -Name 'search'
            apiContracts = Get-ApiContractCatalog
        }
    }

    $scopeMap = [ordered]@{
        contrato = 'contract'
        processo = 'process'
        fornecedor = 'supplier'
        gestor = 'manager'
        fiscal = 'inspector'
        alerta = 'alert'
        suporte = 'support'
        usuario = 'user'
        favorito = 'favorite'
        visao = 'saved_view'
    }
    $scopeFilter = 'all'
    $searchText = $cleanQuery
    $scopeMatch = [regex]::Match($cleanQuery, '^(?<scope>contrato|processo|fornecedor|gestor|fiscal|alerta|suporte|usuario|favorito|visao)\s*:\s*(?<term>.+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($scopeMatch.Success) {
        $scopeFilter = [string]$scopeMap[$scopeMatch.Groups['scope'].Value.ToLowerInvariant()]
        $searchText = Collapse-Whitespace -Text ([string]$scopeMatch.Groups['term'].Value)
    }

    $normalizedQuery = Normalize-IndexText -Text $searchText
    $normalizedQueryVariants = @(Get-SearchAliasVariants -Text $searchText)
    if ([string]::IsNullOrWhiteSpace($normalizedQuery) -or $normalizedQuery.Length -lt 2) {
        return [ordered]@{
            query = $cleanQuery
            scope = $scopeFilter
            total = 0
            items = @()
            shortcuts = @($shortcuts)
            apiContract = New-ApiContractDescriptor -Name 'search'
            apiContracts = Get-ApiContractCatalog
        }
    }

    $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
    $workspacePayload = Get-WorkspaceSessionPayload -User $User
    $capabilities = Get-RoleCapabilities -Role ([string]$User.role)
    $results = New-Object System.Collections.Generic.List[object]
    $computeScore = {
        param(
            [string[]]$Aliases,
            [string]$ReferenceText,
            [string]$TitleText,
            [string]$SubtitleText,
            [string]$MetaText
        )

        if ($scopeFilter -ne 'all' -and @($Aliases) -notcontains $scopeFilter) {
            return -1
        }

        $referenceValue = Normalize-IndexText -Text $ReferenceText
        $titleValue = Normalize-IndexText -Text $TitleText
        $subtitleValue = Normalize-IndexText -Text $SubtitleText
        $metaValue = Normalize-IndexText -Text $MetaText
        $aliasHaystack = [string]::Join(' ', @(
            @(Get-SearchAliasVariants -Text $ReferenceText)
            @(Get-SearchAliasVariants -Text $TitleText)
            @(Get-SearchAliasVariants -Text $SubtitleText)
            @(Get-SearchAliasVariants -Text $MetaText)
        ))
        $haystack = Normalize-IndexText -Text ([string]::Join(' ', @($referenceValue, $titleValue, $subtitleValue, $metaValue, $aliasHaystack)))
        $matchedQuery = @($normalizedQueryVariants | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and $haystack -like "*$_*" } | Select-Object -First 1) | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace([string]$matchedQuery)) {
            return -1
        }

        $score = 20
        if ($referenceValue -eq $normalizedQuery) { $score += 220 }
        elseif ($referenceValue -like "$normalizedQuery*") { $score += 170 }
        elseif ($referenceValue -like "*$normalizedQuery*") { $score += 120 }

        if ($titleValue -eq $normalizedQuery) { $score += 180 }
        elseif ($titleValue -like "$normalizedQuery*") { $score += 140 }
        elseif ($titleValue -like "*$normalizedQuery*") { $score += 90 }

        if ($subtitleValue -like "$normalizedQuery*") { $score += 80 }
        elseif ($subtitleValue -like "*$normalizedQuery*") { $score += 55 }

        if ($metaValue -like "$normalizedQuery*") { $score += 50 }
        elseif ($metaValue -like "*$normalizedQuery*") { $score += 30 }

        if ($scopeFilter -ne 'all') { $score += 25 }
        if ($normalizedQuery.Length -le 6 -and $referenceValue -like "*$normalizedQuery*") { $score += 20 }
        if ([string]$matchedQuery -ne $normalizedQuery) { $score += 35 }
        return $score
    }

    foreach ($item in @($contractsPayload.officialContracts) + @($contractsPayload.contractMovements)) {
        $aliases = @('contract')
        if (-not [string]::IsNullOrWhiteSpace([string]$item.processNumber)) { $aliases += 'process' }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.contractor)) { $aliases += 'supplier' }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.managerName)) { $aliases += 'manager' }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.inspectorName)) { $aliases += 'inspector' }
        $score = & $computeScore `
            -Aliases $aliases `
            -ReferenceText ([string]::Join(' ', @([string]$item.referenceKey, [string]$item.contractNumber, [string]$item.processNumber))) `
            -TitleText ([string]$(if ([string]$item.actTitle) { $item.actTitle } elseif ([string]$item.contractNumber) { $item.contractNumber } else { $item.processNumber })) `
            -SubtitleText ([string]$(if ([string]$item.contractor) { $item.contractor } else { $item.primaryOrganizationName })) `
            -MetaText ([string]::Join(' ', @([string]$item.object, [string]$item.primaryOrganizationName, [string]$item.managerName, [string]$item.inspectorName)))
        if ($score -lt 0) {
            continue
        }

        $isOfficialContract = [bool]($item.PSObject.Properties.Match('portalContractId').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.portalContractId))

        $results.Add([pscustomobject][ordered]@{
            type = 'contract'
            title = [string]$(if ([string]$item.actTitle) { $item.actTitle } elseif ([string]$item.contractNumber) { $item.contractNumber } else { [string]$item.processNumber })
            subtitle = [string]$(if ([string]$item.contractor) { $item.contractor } else { $item.primaryOrganizationName })
            meta = [string]$item.primaryOrganizationName
            reference = [string]$(if ([string]$item.referenceKey) { $item.referenceKey } elseif ([string]$item.contractNumber) { $item.contractNumber } else { $item.processNumber })
            href = "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode($(if ([string]$item.referenceKey) { [string]$item.referenceKey } elseif ([string]$item.contractNumber) { [string]$item.contractNumber } else { [string]$item.processNumber })) )"
            scope = if ($isOfficialContract) { 'Contrato oficial' } else { 'Movimentacao do Diario' }
            score = [int]$score
        })
    }

    foreach ($view in @($workspacePayload.savedViews)) {
        $score = & $computeScore -Aliases @('saved_view') -ReferenceText ([string]$view.id) -TitleText ([string]$view.name) -SubtitleText ([string]$view.page) -MetaText ''
        if ($score -lt 0) {
            continue
        }
        $definition = $view.definition
        $queryParts = New-Object System.Collections.Generic.List[string]
        if ([string]$definition.view -and [string]$definition.view -ne 'official') { $queryParts.Add("view=$([System.Web.HttpUtility]::UrlEncode([string]$definition.view))") }
        if ([string]$definition.quick -and [string]$definition.quick -ne 'all') { $queryParts.Add("quick=$([System.Web.HttpUtility]::UrlEncode([string]$definition.quick))") }
        if ([string]$definition.sort -and [string]$definition.sort -ne 'recent') { $queryParts.Add("sort=$([System.Web.HttpUtility]::UrlEncode([string]$definition.sort))") }
        if ([string]$definition.layout -and [string]$definition.layout -ne 'cards') { $queryParts.Add("layout=$([System.Web.HttpUtility]::UrlEncode([string]$definition.layout))") }
        if ([string]$definition.search) { $queryParts.Add("search=$([System.Web.HttpUtility]::UrlEncode([string]$definition.search))") }
        if ([string]$definition.filterClass) { $queryParts.Add("class=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterClass))") }
        if ([string]$definition.filterType) { $queryParts.Add("type=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterType))") }
        if ([string]$definition.filterYear) { $queryParts.Add("year=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterYear))") }
        if ([string]$definition.filterArea) { $queryParts.Add("area=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterArea))") }
        if ([string]$definition.filterManagement) { $queryParts.Add("management=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterManagement))") }
        if ([string]$definition.filterDocument) { $queryParts.Add("document=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterDocument))") }
        if ([string]$definition.filterOrganization) { $queryParts.Add("organization=$([System.Web.HttpUtility]::UrlEncode([string]$definition.filterOrganization))") }
        $savedViewHref = if ($queryParts.Count -gt 0) { "/contratos.html?$([string]::Join('&', $queryParts))" } else { '/contratos.html' }
        $results.Add([pscustomobject][ordered]@{
            type = 'saved_view'
            title = [string]$view.name
            subtitle = "Visao salva da pagina $([string]$view.page)"
            meta = ''
            reference = [string]$view.id
            href = $savedViewHref
            scope = 'Visao salva'
            score = [int]$score
        })
    }

    foreach ($favorite in @($workspacePayload.favorites)) {
        $score = & $computeScore -Aliases @('favorite', 'contract') -ReferenceText ([string]$favorite.reference) -TitleText ([string]$favorite.title) -SubtitleText ([string]$favorite.organization) -MetaText ([string]$favorite.contractor)
        if ($score -lt 0) {
            continue
        }
        $results.Add([pscustomobject][ordered]@{
            type = 'favorite'
            title = [string]$favorite.title
            subtitle = [string]$favorite.organization
            meta = [string]$favorite.contractor
            reference = [string]$favorite.reference
            href = [string]$favorite.detailUrl
            scope = 'Favorito'
            score = [int]$score
        })
    }

    foreach ($alert in @($workspacePayload.alertsCenter)) {
        $score = & $computeScore -Aliases @('alert') -ReferenceText ([string]$alert.reference) -TitleText ([string]$alert.title) -SubtitleText ([string]$alert.sourceLabel) -MetaText ([string]::Join(' ', @([string]$alert.summary, [string]$alert.stateAssigneeName)))
        if ($score -lt 0) {
            continue
        }
        $results.Add([pscustomobject][ordered]@{
            type = 'alert'
            title = [string]$alert.title
            subtitle = [string]$alert.summary
            meta = [string]$alert.severity
            reference = [string]$alert.reference
            href = if (-not [string]::IsNullOrWhiteSpace([string]$alert.href)) { [string]$alert.href } elseif ([string]::IsNullOrWhiteSpace([string]$alert.reference)) { '/index.html' } else { "/contrato.html?ref=$([System.Web.HttpUtility]::UrlEncode([string]$alert.reference))" }
            scope = 'Alerta'
            score = [int]$score
        })
    }

    $supportPayload = Get-SupportPayload
    $visibleTickets = if ([bool]$capabilities.canSeeAllSupport) {
        @($supportPayload.tickets)
    }
    else {
        @($supportPayload.tickets | Where-Object { [string]$_.requesterUserId -eq [string]$User.id -or [string]$_.assigneeUserId -eq [string]$User.id })
    }

    foreach ($ticket in @($visibleTickets)) {
        $score = & $computeScore -Aliases @('support') -ReferenceText ([string]$ticket.id) -TitleText ([string]$ticket.subject) -SubtitleText ([string]$ticket.requesterName) -MetaText ([string]::Join(' ', @([string]$ticket.message, [string]$ticket.assigneeName)))
        if ($score -lt 0) {
            continue
        }
        $results.Add([pscustomobject][ordered]@{
            type = 'support'
            title = [string]$ticket.subject
            subtitle = [string]$ticket.requesterName
            meta = [string]$ticket.status
            reference = [string]$ticket.id
            href = '/suporte.html'
            scope = 'Suporte'
            score = [int]$score
        })
    }

    if ([bool]$capabilities.canManageUsers) {
        foreach ($managedUser in @(Get-ManagedUsers)) {
            $score = & $computeScore -Aliases @('user') -ReferenceText ([string]$managedUser.login) -TitleText ([string]$managedUser.name) -SubtitleText ([string]$managedUser.roleLabel) -MetaText ''
            if ($score -lt 0) {
                continue
            }
            $results.Add([pscustomobject][ordered]@{
                type = 'user'
                title = [string]$managedUser.name
                subtitle = [string]$managedUser.roleLabel
                meta = "Login $([string]$managedUser.login)"
                reference = [string]$managedUser.login
                href = '/usuarios.html'
                scope = 'Usuario'
                score = [int]$score
            })
        }
    }

    $items = @(
        $results |
        Sort-Object @{ Expression = { [int]$_.score }; Descending = $true }, @{ Expression = { [string]$_.title }; Descending = $false } |
        Select-Object -First $Limit
    )

    Register-ObservabilityEvent -Type 'search' -Status $scopeFilter -Message "Busca executada para '$searchText'." -UserLogin ([string]$User.login) -Metadata ([ordered]@{
        query = $cleanQuery
        normalizedQuery = $normalizedQuery
        scope = $scopeFilter
        total = [int]$results.Count
    }) | Out-Null

    return [ordered]@{
        query = $cleanQuery
        scope = $scopeFilter
        total = [int]$results.Count
        items = @($items)
        shortcuts = @($shortcuts)
        apiContract = New-ApiContractDescriptor -Name 'search'
        apiContracts = Get-ApiContractCatalog
    }
}

function Get-ScopedFinancialMonitoringPayload {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$FinancialMonitoring,

        [Parameter(Mandatory = $false)]
        [bool]$IncludeDetails = $false
    )

    $source = if ($null -ne $FinancialMonitoring) { $FinancialMonitoring } else { (Get-EmptyContractsPayload).financialMonitoring }

    return [ordered]@{
        mode = [string]$source.mode
        modeLabel = [string]$source.modeLabel
        note = [string]$source.note
        monitoredContracts = [int]$source.monitoredContracts
        searchableContracts = [int]$source.searchableContracts
        queryReadyContracts = [int]$source.queryReadyContracts
        withContractValue = [int]$source.withContractValue
        expenseContracts = [int]$source.expenseContracts
        revenueContracts = [int]$source.revenueContracts
        automatedContracts = [int]$source.automatedContracts
        assistedContracts = [int]$source.assistedContracts
        limitedContracts = [int]$source.limitedContracts
        unmappedContracts = [int]$source.unmappedContracts
        averageCoverageScore = [int]$source.averageCoverageScore
        automationReadySources = [int]$source.automationReadySources
        sourceCount = [int]$source.sourceCount
        executionStageCount = [int]$source.executionStageCount
        detailSectionCount = [int]$source.detailSectionCount
        expensePortal = if ($null -ne $source.expensePortal) {
            [pscustomobject][ordered]@{
                status = [string]$source.expensePortal.status
                statusLabel = [string]$source.expensePortal.statusLabel
                requiresCaptcha = [bool]$source.expensePortal.requiresCaptcha
                requiresCallback = [bool]$source.expensePortal.requiresCallback
                executionStageCount = [int]$source.expensePortal.executionStageCount
                detailSectionCount = [int]$source.expensePortal.detailSectionCount
                accessNote = [string]$source.expensePortal.accessNote
            }
        }
        else {
            $null
        }
        coverageBreakdown = if ($IncludeDetails) { @($source.coverageBreakdown) } else { @() }
        sources = if ($IncludeDetails) { @($source.sources) } else { @() }
    }
}

function Get-DashboardPayloadCacheKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageScope
    )

    $parts = @(
        $PageScope
        (Get-FileCacheStamp -Path $script:StatusPath)
        (Get-FileCacheStamp -Path $script:DiariesPath)
        (Get-FileCacheStamp -Path $script:ContractsPath)
        (Get-FileCacheStamp -Path $script:OrganizationCatalogPath)
    )

    return [string]::Join('|', $parts)
}

function Get-WorkspaceSessionCacheKey {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )

    $parts = @(
        [string]$User.id
        [string]$User.role
        (Get-FileCacheStamp -Path $script:ContractsPath)
        (Get-FileCacheStamp -Path $script:WorkspaceStatePath)
        (Get-FileCacheStamp -Path $script:SupportPath)
        (Get-FileCacheStamp -Path $script:UsersPath)
        (Get-FileCacheStamp -Path $script:ObservabilityPath)
    )

    return [string]::Join('|', $parts)
}

function Get-DashboardPayload {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Page = 'resumo',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$User = $null
    )

    $pageScope = ([string]$Page).Trim().ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($pageScope)) {
        $pageScope = 'resumo'
    }

    $cacheKey = Get-DashboardPayloadCacheKey -PageScope $pageScope
    $basePayload = $null
    if ($script:DashboardPayloadCache.ContainsKey($cacheKey)) {
        $basePayload = $script:DashboardPayloadCache[$cacheKey]
    }
    else {
        $status = Get-StatusHashtable
        $diariesPayload = Get-DiariesPayload
        $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
        $organizationCatalog = Get-OrganizationCatalog
        $overviewSummary = Get-DashboardOverviewSummary -DiariesPayload $diariesPayload -ContractsPayload $contractsPayload

        $includeOfficialContracts = $false
        $includeContractMovements = $false
        $includeDiaries = $false
        $includeAnalyses = $false
        $includeDiaryPreview = $false
        $includeCatalog = $false
        $includeTypeSummary = $false
        $includeOrganizationSummary = $false
        $includeAreaSummary = $false
        $includeSupplierSummary = $false
        $includeContractDiagnostics = $false

        switch ($pageScope) {
            'contratos' {
                $includeOfficialContracts = $true
                $includeContractMovements = $true
                $includeCatalog = $true
                $includeContractDiagnostics = $true
            }
            'organizacao' {
                $includeOfficialContracts = $true
                $includeCatalog = $true
                $includeTypeSummary = $true
                $includeOrganizationSummary = $true
                $includeAreaSummary = $true
                $includeSupplierSummary = $true
            }
            'edicoes' {
                $includeDiaries = $true
                $includeAnalyses = $true
            }
            default {
                $includeOfficialContracts = $true
                $includeDiaryPreview = $true
                $includeCatalog = $true
                $includeTypeSummary = $true
                $includeOrganizationSummary = $true
                $includeAreaSummary = $true
                $includeSupplierSummary = $true
            }
        }

        $basePayload = [ordered]@{
            generatedAt = (Get-IsoNow)
            apiContract = New-ApiContractDescriptor -Name 'dashboard'
            apiContracts = Get-ApiContractCatalog
            parserVersion = $script:ParserVersion
            sourcePortal = ([Uri]::new($script:BasePortalUri, $script:PortalDiarioPath)).AbsoluteUri
            keywords = (Get-ContractKeywords)
            roleCatalog = @(Get-RoleCatalog)
            status = $status
            overviewSummary = $overviewSummary
            diaries = if ($includeDiaries) { @($diariesPayload.diaries) } else { @() }
            analyses = if ($includeAnalyses) { @($contractsPayload.analyses) } else { @() }
            diaryPreview = if ($includeDiaryPreview) { @(Get-DashboardDiaryPreview -Diaries @($diariesPayload.diaries) -Analyses @($contractsPayload.analyses)) } else { @() }
            analysisAlertPreview = if ($includeDiaryPreview) { @(Get-DashboardAnalysisAlertPreview -Diaries @($diariesPayload.diaries) -Analyses @($contractsPayload.analyses)) } else { @() }
            contracts = @()
            officialContracts = if ($includeOfficialContracts) { @($contractsPayload.officialContracts) } else { @() }
            contractMovements = if ($includeContractMovements) { @($contractsPayload.contractMovements) } else { @() }
            managementProfiles = @()
            crossReviewQueue = if ($includeContractDiagnostics) { @($contractsPayload.crossReviewQueue) } else { @() }
            crossSourceDivergences = if ($includeContractDiagnostics) { @($contractsPayload.crossSourceDivergences) } else { @() }
            crossSourceAlerts = if ($includeContractDiagnostics) { @($contractsPayload.crossSourceAlerts) } else { @() }
            crossSourceSuppressionSummary = if ($includeContractDiagnostics) { $contractsPayload.crossSourceSuppressionSummary } else { [ordered]@{ total = 0; reasons = @() } }
            financialMonitoring = if ($includeOfficialContracts -or $includeContractMovements) { (Get-ScopedFinancialMonitoringPayload -FinancialMonitoring $contractsPayload.financialMonitoring -IncludeDetails:$includeContractDiagnostics) } else { (Get-ScopedFinancialMonitoringPayload -FinancialMonitoring $null -IncludeDetails:$false) }
            organizationCatalog = if ($includeCatalog) { $organizationCatalog } else { (Get-EmptyOrganizationCatalog) }
            qualitySummary = $contractsPayload.qualitySummary
            managementSummary = $contractsPayload.managementSummary
            crosswalkSummary = $contractsPayload.crosswalkSummary
            typeSummary = if ($includeTypeSummary) { @($contractsPayload.typeSummary) } else { @() }
            organizationSummary = if ($includeOrganizationSummary) { @($contractsPayload.organizationSummary) } else { @() }
            areaSummary = if ($includeAreaSummary) { @(Get-DashboardAreaSummary -Contracts @($contractsPayload.items) -OrganizationCatalog $organizationCatalog) } else { @() }
            supplierSummary = if ($includeSupplierSummary) { @(Get-DashboardSupplierSummary -Contracts @($contractsPayload.items)) } else { @() }
            contractSummary = [ordered]@{
                totalItems = [int]$contractsPayload.totalItems
                totalValue = [double]$contractsPayload.totalValue
                analyzedDiaryCount = [int]$contractsPayload.analyzedDiaryCount
                uniqueSuppliers = [int]$contractsPayload.uniqueSuppliers
                officialPortalContracts = [int]$contractsPayload.officialPortalContracts
                officialMatched = [int]$contractsPayload.crosswalkSummary.officialMatched
                officialPendingReview = [int]$contractsPayload.crosswalkSummary.officialPendingReview
                movementMatched = [int]$contractsPayload.crosswalkSummary.movementMatched
                movementPendingReview = [int]$contractsPayload.crosswalkSummary.movementPendingReview
                divergences = [int]$contractsPayload.crosswalkSummary.divergences
                suppressedDivergences = [int]$contractsPayload.crosswalkSummary.suppressedDivergences
                operationalAlerts = [int]$contractsPayload.crosswalkSummary.operationalAlerts
                searchableFinancialContracts = [int]$contractsPayload.financialMonitoring.searchableContracts
                queryReadyFinancialContracts = [int]$contractsPayload.financialMonitoring.queryReadyContracts
                averageFinancialCoverageScore = [int]$contractsPayload.financialMonitoring.averageCoverageScore
            }
        }

        Add-ScriptCacheEntry -Cache $script:DashboardPayloadCache -Key $cacheKey -Value $basePayload -MaxEntries 12
    }

    return [ordered]@{
        generatedAt = $basePayload.generatedAt
        apiContract = $basePayload.apiContract
        apiContracts = $basePayload.apiContracts
        parserVersion = $basePayload.parserVersion
        sourcePortal = $basePayload.sourcePortal
        keywords = $basePayload.keywords
        roleCatalog = $basePayload.roleCatalog
        status = $basePayload.status
        overviewSummary = $basePayload.overviewSummary
        diaries = $basePayload.diaries
        analyses = $basePayload.analyses
        diaryPreview = $basePayload.diaryPreview
        analysisAlertPreview = $basePayload.analysisAlertPreview
        contracts = $basePayload.contracts
        officialContracts = $basePayload.officialContracts
        contractMovements = $basePayload.contractMovements
        managementProfiles = $basePayload.managementProfiles
        crossReviewQueue = $basePayload.crossReviewQueue
        crossSourceDivergences = $basePayload.crossSourceDivergences
        crossSourceAlerts = $basePayload.crossSourceAlerts
        crossSourceSuppressionSummary = $basePayload.crossSourceSuppressionSummary
        financialMonitoring = $basePayload.financialMonitoring
        organizationCatalog = $basePayload.organizationCatalog
        qualitySummary = $basePayload.qualitySummary
        managementSummary = $basePayload.managementSummary
        crosswalkSummary = $basePayload.crosswalkSummary
        typeSummary = $basePayload.typeSummary
        organizationSummary = $basePayload.organizationSummary
        areaSummary = $basePayload.areaSummary
        supplierSummary = $basePayload.supplierSummary
        contractSummary = $basePayload.contractSummary
    }
}

function Test-SuspiciousManagementAssignment {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Role
    )

    $normalizedName = Normalize-IndexText -Text $Name
    $normalizedRole = Normalize-IndexText -Text $Role

    if ([string]::IsNullOrWhiteSpace($normalizedName)) {
        return $false
    }

    if ($normalizedName -match '^(PORTARIA|DECRETO|LEI|ART|EXTRATO|CONTRATO|DESIGNA|NOMEIA)\b') {
        return $true
    }

    if ($normalizedRole.Length -ge 170) {
        return $true
    }

    if ($normalizedRole -match '\bDESIGNA GESTOR E FISCAL\b') {
        return $true
    }

    if ($normalizedRole -match '\bPREFEITO MUNICIPAL\b' -and $normalizedRole.Length -ge 120) {
        return $true
    }

    return $false
}

function Get-ContractAuditPayload {
    $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
    $officialContracts = @($contractsPayload.officialContracts)
    $contractMovements = @($contractsPayload.contractMovements)
    $managementProfiles = @($contractsPayload.managementProfiles)
    $crossReviewQueue = @($contractsPayload.crossReviewQueue)
    $crossSourceDivergences = @($contractsPayload.crossSourceDivergences)
    $crossSourceAlerts = @($contractsPayload.crossSourceAlerts)
    $crossSourceSuppressionSummary = $contractsPayload.crossSourceSuppressionSummary
    $financialMonitoring = $contractsPayload.financialMonitoring

    $atRiskContracts = @(
        $officialContracts |
        Where-Object {
            [bool]$_.managementTracked -and (
                -not [bool]$_.hasManager -or
                -not [bool]$_.hasInspector -or
                [bool]$_.managerExonerationSignal -or
                [bool]$_.inspectorExonerationSignal
            )
        } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.contractNumber }; Descending = $false }
    )
    $changedAssignments = @(
        $officialContracts |
        Where-Object { [bool]$_.managerChanged -or [bool]$_.inspectorChanged } |
        Sort-Object -Property `
            @{ Expression = { $_.managementLastActAt }; Descending = $true }, `
            @{ Expression = { $_.contractNumber }; Descending = $false }
    )
    $withoutDocument = @(
        $officialContracts |
        Where-Object { [string]::IsNullOrWhiteSpace([string]$_.localPdfRelative) } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.contractNumber }; Descending = $false }
    )
    $suspiciousAssignments = @(
        $officialContracts |
        Where-Object {
            (Test-SuspiciousManagementAssignment -Name ([string]$_.managerName) -Role ([string]$_.managerRole)) -or
            (Test-SuspiciousManagementAssignment -Name ([string]$_.inspectorName) -Role ([string]$_.inspectorRole))
        } |
        Sort-Object -Property `
            @{ Expression = { $_.managementLastActAt }; Descending = $true }, `
            @{ Expression = { $_.publishedAt }; Descending = $true }
    )
    $movementsWithoutReference = @(
        $contractMovements |
        Where-Object { [string]::IsNullOrWhiteSpace([string]$_.referenceKey) } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.pageNumber }; Descending = $false }
    )

    return [ordered]@{
        generatedAt = (Get-IsoNow)
        apiContract = New-ApiContractDescriptor -Name 'contractAudit'
        apiContracts = Get-ApiContractCatalog
        parserVersion = $script:ParserVersion
        summary = [ordered]@{
            officialContracts = [int]@($officialContracts).Count
            trackedContracts = [int]@($managementProfiles).Count
            atRiskContracts = [int]@($atRiskContracts).Count
            changedAssignments = [int]@($changedAssignments).Count
            exonerationSignals = [int]@($officialContracts | Where-Object { [bool]$_.managerExonerationSignal -or [bool]$_.inspectorExonerationSignal }).Count
            withoutManager = [int]@($officialContracts | Where-Object { [bool]$_.managementTracked -and -not [bool]$_.hasManager }).Count
            withoutInspector = [int]@($officialContracts | Where-Object { [bool]$_.managementTracked -and -not [bool]$_.hasInspector }).Count
            withoutDocument = [int]@($withoutDocument).Count
            suspiciousAssignments = [int]@($suspiciousAssignments).Count
            movementsWithoutReference = [int]@($movementsWithoutReference).Count
            crossReviewQueue = [int]@($crossReviewQueue).Count
            crossSourceDivergences = [int]@($crossSourceDivergences).Count
            crossSourceSuppressed = [int]$crossSourceSuppressionSummary.total
            crossSourceAlerts = [int]@($crossSourceAlerts).Count
            searchableFinancialContracts = [int]$financialMonitoring.searchableContracts
            queryReadyFinancialContracts = [int]$financialMonitoring.queryReadyContracts
            averageFinancialCoverageScore = [int]$financialMonitoring.averageCoverageScore
        }
        atRiskContracts = @($atRiskContracts)
        changedAssignments = @($changedAssignments)
        withoutDocument = @($withoutDocument)
        suspiciousAssignments = @($suspiciousAssignments)
        movementsWithoutReference = @($movementsWithoutReference)
        crossReviewQueue = @($crossReviewQueue)
        crossSourceDivergences = @($crossSourceDivergences)
        crossSourceAlerts = @($crossSourceAlerts)
        crossSourceSuppressionSummary = $crossSourceSuppressionSummary
        financialMonitoring = $financialMonitoring
        crosswalkSummary = $contractsPayload.crosswalkSummary
    }
}

function Get-ContractWorkspaceHistory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reference
    )

    $workspacePayload = Get-WorkspacePayload
    $supportPayload = Get-SupportPayload
    $history = New-Object System.Collections.Generic.List[object]

    foreach ($note in @($workspacePayload.contractNotes | Where-Object {
        Test-ReferenceMatch -Reference $Reference -CandidateValues @([string]$_.reference)
    })) {
        $history.Add([pscustomobject][ordered]@{
            type = 'note'
            createdAt = [string]$note.createdAt
            title = 'Comentario interno'
            detail = [string]$note.body
            actor = [string]$(if ([string]$note.createdByName) { $note.createdByName } else { $note.createdBy })
            status = ''
            sourceLabel = 'Equipe interna'
        })
    }

    foreach ($workflowItem in @($workspacePayload.workflowItems | Where-Object {
        Test-ReferenceMatch -Reference $Reference -CandidateValues @([string]$_.reference)
    } | Select-Object -First 1)) {
        foreach ($entry in @($workflowItem.history)) {
            $history.Add([pscustomobject][ordered]@{
                type = 'workflow'
                createdAt = [string]$entry.createdAt
                title = "Workflow $([string]$entry.status -replace '_', ' ')"
                detail = [string]$(if ([string]$entry.note) { $entry.note } elseif ([string]$entry.assigneeName) { "Responsavel: $([string]$entry.assigneeName)." } else { 'Workflow contratual atualizado.' })
                actor = [string]$entry.actor
                status = [string]$entry.status
                sourceLabel = 'Workflow interno'
            })
        }
    }

    foreach ($alertState in @($workspacePayload.alertStates | Where-Object {
        Test-ReferenceMatch -Reference $Reference -CandidateValues @([string]$_.reference)
    })) {
        foreach ($entry in @($alertState.history)) {
            $history.Add([pscustomobject][ordered]@{
                type = 'alert'
                createdAt = [string]$entry.createdAt
                title = [string]$(if ([string]$alertState.title) { $alertState.title } else { 'Alerta operacional' })
                detail = [string]$(if ([string]$entry.summary) { $entry.summary } elseif ([string]$alertState.justification) { $alertState.justification } else { 'Acao registrada na central operacional.' })
                actor = [string]$entry.actor
                status = [string]$alertState.status
                sourceLabel = 'Central operacional'
            })
        }
    }

    foreach ($activity in @($workspacePayload.activityLog | Where-Object {
        Test-ReferenceMatch -Reference $Reference -CandidateValues @([string]$_.reference)
    })) {
        $history.Add([pscustomobject][ordered]@{
            type = [string]$activity.type
            createdAt = [string]$activity.createdAt
            title = [string]$activity.title
            detail = [string]$activity.summary
            actor = [string]$activity.createdBy
            status = ''
            sourceLabel = 'Trilha de atividade'
        })
    }

    foreach ($ticket in @($supportPayload.tickets | Where-Object {
        $searchFields = @([string]$_.subject, [string]$_.message, [string]$_.adminResponse)
        @($searchFields | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and (Normalize-IndexText -Text ([string]$_)) -like "*$(Normalize-IndexText -Text $Reference)*" }).Count -gt 0
    })) {
        $history.Add([pscustomobject][ordered]@{
            type = 'support'
            createdAt = [string]$(if ([string]$ticket.updatedAt) { $ticket.updatedAt } else { $ticket.createdAt })
            title = [string]$ticket.subject
            detail = [string]$(if ([string]$ticket.adminResponse) { $ticket.adminResponse } else { $ticket.message })
            actor = [string]$(if ([string]$ticket.assigneeName) { $ticket.assigneeName } else { $ticket.requesterName })
            status = [string]$ticket.status
            sourceLabel = 'Suporte interno'
        })
    }

    return @(
        $history |
        Sort-Object @{ Expression = { [string]$_.createdAt }; Descending = $true }, @{ Expression = { [string]$_.title }; Descending = $false } |
        Select-Object -First 40
    )
}

function Get-ContractVersionComparisonData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reference,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$OfficialContract = $null,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ManagementProfile = $null
    )

    $referenceCandidates = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @(
        [string]$Reference,
        [string]$(if ($OfficialContract) { $OfficialContract.referenceKey } else { '' }),
        [string]$(if ($OfficialContract) { $OfficialContract.contractNumber } else { '' }),
        [string]$(if ($OfficialContract) { $OfficialContract.processNumber } else { '' }),
        [string]$(if ($OfficialContract) { $OfficialContract.portalContractId } else { '' }),
        [string]$(if ($ManagementProfile) { $ManagementProfile.contractKey } else { '' }),
        [string]$(if ($ManagementProfile) { $ManagementProfile.contractNumber } else { '' }),
        [string]$(if ($ManagementProfile) { $ManagementProfile.processNumber } else { '' })
    )) {
        $cleanCandidate = Collapse-Whitespace -Text $candidate
        if (-not [string]::IsNullOrWhiteSpace($cleanCandidate) -and -not @($referenceCandidates).Contains($cleanCandidate)) {
            $referenceCandidates.Add($cleanCandidate) | Out-Null
        }
    }

    $snapshots = @((Get-WorkspacePayload).aggregateSnapshots | Sort-Object @{ Expression = { [string]$_.generatedAt }; Descending = $true })
    $currentSnapshot = @($snapshots | Select-Object -First 1) | Select-Object -First 1
    $previousSnapshot = @($snapshots | Select-Object -Skip 1 -First 1) | Select-Object -First 1

    if ($null -eq $currentSnapshot) {
        return [ordered]@{
            available = $false
            reference = $Reference
            status = 'unavailable'
            headline = 'Sem snapshots historicos para comparar.'
            summary = 'A comparacao entre versoes ficara disponivel apos a proxima recomposicao registrada no workspace.'
            nextStep = 'Execute nova sincronizacao para formar a primeira comparacao entre versões.'
            currentGeneratedAt = $null
            previousGeneratedAt = $null
            changeCount = 0
            severityScore = 0
            changedFields = @()
            history = @()
            currentRow = $null
            previousRow = $null
        }
    }

    $currentRows = @(Get-SnapshotArrayProperty -Snapshot $currentSnapshot -Name 'contractVersionRows')
    $previousRows = @(Get-SnapshotArrayProperty -Snapshot $previousSnapshot -Name 'contractVersionRows')
    $currentRow = $null
    $previousRow = $null

    foreach ($candidate in @($referenceCandidates)) {
        if ($null -eq $currentRow) {
            $currentRow = Find-SnapshotContractRow -Rows $currentRows -Reference $candidate
        }
        if ($null -eq $previousRow) {
            $previousRow = Find-SnapshotContractRow -Rows $previousRows -Reference $candidate
        }
        if ($currentRow -and $previousRow) {
            break
        }
    }

    if ($null -eq $currentRow) {
        return [ordered]@{
            available = $false
            reference = [string]$(if (@($referenceCandidates).Count -gt 0) { $referenceCandidates[0] } else { $Reference })
            status = 'unavailable'
            headline = 'Sem linha contratual consolidada no snapshot atual.'
            summary = 'O dossie ainda nao encontrou esse contrato dentro dos snapshots historicos usados para diff.'
            nextStep = 'Mantenha a comparacao historica em acompanhamento ate o contrato aparecer na carteira snapshot.'
            currentGeneratedAt = [string]$currentSnapshot.generatedAt
            previousGeneratedAt = if ($previousSnapshot) { [string]$previousSnapshot.generatedAt } else { $null }
            changeCount = 0
            severityScore = 0
            changedFields = @()
            history = @()
            currentRow = $null
            previousRow = $previousRow
        }
    }

    $changedFields = if ($previousRow) {
        @(
            @(Get-SnapshotContractVersionChanges -CurrentRow $currentRow -PreviousRow $previousRow) |
            Sort-Object @{ Expression = { [int]$_.severity }; Descending = $true }, @{ Expression = { [string]$_.label }; Descending = $false }
        )
    }
    else {
        @()
    }
    $status = if (-not $previousRow) {
        'new'
    }
    elseif (@($changedFields).Count -gt 0) {
        'changed'
    }
    else {
        'stable'
    }
    $severityScore = if (@($changedFields).Count -gt 0) {
        [int](@($changedFields | Measure-Object -Property severity -Sum).Sum)
    }
    else {
        0
    }
    $headline = switch ($status) {
        'new' { 'Contrato apareceu pela primeira vez no historico comparavel.' }
        'changed' { "Contrato mudou em $([int]@($changedFields).Count) campo(s) desde a ultima base." }
        default { 'Sem alteracao de campos-chave no ultimo intervalo comparado.' }
    }
    $summary = switch ($status) {
        'new' {
            [string]$(if ([string]$currentRow.portalStatus) {
                "Entrada no historico com status $([string]$currentRow.portalStatus)."
            }
            else {
                'Entrada no historico consolidado.'
            })
        }
        'changed' {
            [string]::Join(' | ', @(
                @($changedFields | Select-Object -First 3) |
                ForEach-Object { "$([string]$_.label): $([string]$_.previous) -> $([string]$_.current)" }
            ))
        }
        default {
            'Cadastro, vigencia, cruzamento e sinais operacionais permaneceram estaveis neste recorte.'
        }
    }
    $nextStep = switch ($status) {
        'new' { 'Classificar a entrada do contrato e validar o enquadramento operacional inicial.' }
        'changed' {
            if (@($changedFields | Where-Object { [int]$_.severity -ge 35 }).Count -gt 0) {
                'Conferir o dossie e registrar se a mudanca altera prioridade, risco ou workflow.'
            }
            else {
                'Anotar a mudanca e manter o acompanhamento regular do contrato.'
            }
        }
        default { 'Usar essa comparacao como base historica e acompanhar o proximo ciclo.' }
    }

    $history = New-Object System.Collections.Generic.List[object]
    foreach ($snapshot in @($snapshots | Select-Object -First 6)) {
        $snapshotRows = @(Get-SnapshotArrayProperty -Snapshot $snapshot -Name 'contractVersionRows')
        $snapshotRow = $null
        foreach ($candidate in @($referenceCandidates)) {
            $snapshotRow = Find-SnapshotContractRow -Rows $snapshotRows -Reference $candidate
            if ($snapshotRow) {
                break
            }
        }
        if (-not $snapshotRow) {
            continue
        }

        $history.Add([pscustomobject][ordered]@{
            generatedAt = [string]$snapshot.generatedAt
            portalStatus = [string]$snapshotRow.portalStatus
            vigencyLabel = Get-SnapshotContractStatusLabel -Key 'isActive' -Value $snapshotRow.isActive
            valueLabel = Get-SnapshotContractStatusLabel -Key 'valueNumber' -Value $snapshotRow.valueNumber
            managementLabel = Get-SnapshotContractStatusLabel -Key 'managementStatus' -Value $snapshotRow.managementStatus
            crossLabel = Get-SnapshotContractStatusLabel -Key 'crossStatus' -Value $snapshotRow.crossStatus
            alertsLabel = Get-SnapshotContractStatusLabel -Key 'alertCount' -Value $snapshotRow.alertCount
            divergencesLabel = Get-SnapshotContractStatusLabel -Key 'divergenceCount' -Value $snapshotRow.divergenceCount
            reviewPending = [bool]$snapshotRow.reviewPending
            summary = [string]::Join(' | ', @(
                @(
                    [string]$snapshotRow.portalStatus,
                    (Get-SnapshotContractStatusLabel -Key 'isActive' -Value $snapshotRow.isActive),
                    (Get-SnapshotContractStatusLabel -Key 'managementStatus' -Value $snapshotRow.managementStatus),
                    (Get-SnapshotContractStatusLabel -Key 'crossStatus' -Value $snapshotRow.crossStatus)
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
            ))
        }) | Out-Null
    }

    return [ordered]@{
        available = $true
        reference = [string]$(if ([string]$currentRow.reference) { $currentRow.reference } elseif (@($referenceCandidates).Count -gt 0) { $referenceCandidates[0] } else { $Reference })
        status = $status
        headline = $headline
        summary = Collapse-Whitespace -Text $summary
        nextStep = Collapse-Whitespace -Text $nextStep
        currentGeneratedAt = [string]$currentSnapshot.generatedAt
        previousGeneratedAt = if ($previousSnapshot) { [string]$previousSnapshot.generatedAt } else { $null }
        changeCount = [int]@($changedFields).Count
        severityScore = [int]$severityScore
        changedFields = @($changedFields)
        history = @($history.ToArray())
        currentRow = $currentRow
        previousRow = $previousRow
    }
}

function Get-ContractDetailPayload {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Reference
    )

    $contractsPayload = Read-JsonFile -Path $script:ContractsPath -Default (Get-EmptyContractsPayload)
    $officialContracts = @($contractsPayload.officialContracts)
    $contractMovements = @($contractsPayload.contractMovements)
    $managementProfiles = @($contractsPayload.managementProfiles)
    $crossReviewQueue = @($contractsPayload.crossReviewQueue)
    $crossSourceDivergences = @($contractsPayload.crossSourceDivergences)
    $crossSourceAlerts = @($contractsPayload.crossSourceAlerts)
    $ref = [string]$Reference

    if ([string]::IsNullOrWhiteSpace($ref)) {
        throw 'Referência do contrato não informada.'
    }

    if ($ref.Length -gt 96) {
        throw 'Referencia do contrato invalida.'
    }

    $lookupTokens = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($token in @(Get-ContractReferenceTokens -ContractNumber $ref -ProcessNumber $ref)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
            $null = $lookupTokens.Add([string]$token)
        }
    }

    $normalizedRef = Normalize-IndexText -Text $ref
    if (-not [string]::IsNullOrWhiteSpace($normalizedRef)) {
        $null = $lookupTokens.Add(($normalizedRef -replace '\s+', ''))
    }

    $officialContract = $officialContracts |
        Where-Object {
            $itemTokens = @(
                Get-ContractReferenceTokens -ContractNumber ([string]$_.contractNumber) -ProcessNumber ([string]$_.processNumber)
            ) + @([string]$_.referenceKey, [string]$_.portalContractId)

            @($itemTokens | Where-Object { $_ -and $lookupTokens.Contains([string]$_) }).Count -gt 0
        } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.updatedAt }; Descending = $true } |
        Select-Object -First 1

    $managementProfile = $managementProfiles |
        Where-Object {
            @(
                @($_.referenceTokens) + @([string]$_.contractKey, [string]$_.contractNumber, [string]$_.processNumber) |
                Where-Object { $_ -and $lookupTokens.Contains([string]$_) }
            ).Count -gt 0
        } |
        Select-Object -First 1

    if (-not $officialContract -and $managementProfile) {
        $officialContract = $officialContracts |
            Where-Object {
                @(
                    Get-ContractReferenceTokens -ContractNumber ([string]$_.contractNumber) -ProcessNumber ([string]$_.processNumber) |
                    Where-Object { $_ -and (@($managementProfile.referenceTokens) -contains [string]$_) }
                ).Count -gt 0
            } |
            Sort-Object -Property @{ Expression = { $_.publishedAt }; Descending = $true } |
            Select-Object -First 1
    }

    if (-not $managementProfile -and $officialContract) {
        $managementProfile = $managementProfiles |
            Where-Object {
                @(
                    @($_.referenceTokens) |
                    Where-Object { $_ -and (@(Get-ContractReferenceTokens -ContractNumber ([string]$officialContract.contractNumber) -ProcessNumber ([string]$officialContract.processNumber)) -contains [string]$_) }
                ).Count -gt 0
            } |
            Select-Object -First 1
    }

    $effectiveTokens = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($token in @($lookupTokens)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
            $null = $effectiveTokens.Add([string]$token)
        }
    }
    if ($managementProfile) {
        foreach ($token in @($managementProfile.referenceTokens)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
                $null = $effectiveTokens.Add([string]$token)
            }
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$managementProfile.contractKey)) {
            $null = $effectiveTokens.Add([string]$managementProfile.contractKey)
        }
    }

    $detailCrossKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($token in @($effectiveTokens)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$token) -and $token -match '^\d+/\d{4}$') {
            $null = $detailCrossKeys.Add([string]$token)
        }
    }

    $relatedMovements = @(
        $contractMovements |
        Where-Object {
            $itemTokens = @(
                Get-ContractReferenceTokens -ContractNumber ([string]$_.contractNumber) -ProcessNumber ([string]$_.processNumber)
            ) + @([string]$_.referenceKey, [string]$_.managementProfileKey)

            @($itemTokens | Where-Object { $_ -and $effectiveTokens.Contains([string]$_) }).Count -gt 0
        } |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { if ($_.PSObject.Properties['pageNumber']) { [int]$_.pageNumber } else { 0 } }; Descending = $false }
    )

    if (-not $officialContract -and -not $managementProfile -and @($relatedMovements).Count -eq 0) {
        throw 'Contrato não encontrado.'
    }

    $detailCrossKey = if ($officialContract -and $officialContract.PSObject.Properties['crossSource']) {
        [string]$officialContract.crossSource.crossKey
    }
    elseif ($managementProfile -and -not [string]::IsNullOrWhiteSpace([string]$managementProfile.contractKey)) {
        [string]$managementProfile.contractKey
    }
    else {
        @($detailCrossKeys) | Select-Object -First 1
    }

    if (-not [string]::IsNullOrWhiteSpace($detailCrossKey)) {
        $null = $detailCrossKeys.Add($detailCrossKey)
    }

    $detailDivergences = @(
        $crossSourceDivergences |
        Where-Object {
            (
                -not [string]::IsNullOrWhiteSpace([string]$_.crossKey) -and $detailCrossKeys.Contains([string]$_.crossKey)
            ) -or (
                $officialContract -and -not [string]::IsNullOrWhiteSpace([string]$_.portalContractId) -and [string]$_.portalContractId -eq [string]$officialContract.portalContractId
            )
        } |
        Sort-Object -Property @{ Expression = { [string]$_.publishedAt }; Descending = $true }
    )
    $detailAlerts = @(
        $crossSourceAlerts |
        Where-Object {
            (
                -not [string]::IsNullOrWhiteSpace([string]$_.crossKey) -and $detailCrossKeys.Contains([string]$_.crossKey)
            ) -or (
                $officialContract -and -not [string]::IsNullOrWhiteSpace([string]$_.portalContractId) -and [string]$_.portalContractId -eq [string]$officialContract.portalContractId
            )
        } |
        Sort-Object -Property @{ Expression = { [string]$_.publishedAt }; Descending = $true }
    )
    $reviewQueueEntry = @(
        $crossReviewQueue |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.crossKey) -and $detailCrossKeys.Contains([string]$_.crossKey) } |
        Select-Object -First 1
    ) | Select-Object -First 1

    $timeline = New-Object System.Collections.Generic.List[object]

    if ($officialContract) {
        $timeline.Add([pscustomobject]@{
            eventType = 'contrato_oficial'
            publishedAt = [string]$officialContract.publishedAt
            title = [string]$officialContract.actTitle
            detail = [string]$officialContract.portalStatus
            sourceLabel = 'Portal de contratos'
            sourceUrl = [string]$officialContract.viewUrl
            webPdfPath = [string]$officialContract.webPdfPath
            pageNumber = 0
        })
    }

    foreach ($movement in @($relatedMovements)) {
        $timeline.Add([pscustomobject]@{
            eventType = [string]$movement.recordClass
            publishedAt = [string]$movement.publishedAt
            title = [string]$movement.actTitle
            detail = [string]$movement.recordClassLabel
            sourceLabel = if ($movement.edition) { "Diário Oficial - edição $([string]$movement.edition)" } else { 'Diário Oficial' }
            sourceUrl = [string]$movement.viewUrl
            webPdfPath = [string]$movement.webPdfPath
            pageNumber = [int]$movement.pageNumber
        })
    }

    if ($officialContract -and $officialContract.PSObject.Properties['crossSource']) {
        $crossStatus = [string]$officialContract.crossSource.status
        if (-not [string]::IsNullOrWhiteSpace($crossStatus) -and $crossStatus -ne 'unmatched') {
            $timeline.Add([pscustomobject]@{
                eventType = 'cruzamento_fontes'
                publishedAt = if ([string]$officialContract.crossSource.reviewedAt) { [string]$officialContract.crossSource.reviewedAt } elseif ([string]$officialContract.updatedAt) { [string]$officialContract.updatedAt } else { [string]$officialContract.publishedAt }
                title = if ($crossStatus -eq 'reviewed') { 'Vinculo entre portal e Diario confirmado manualmente' } elseif ($crossStatus -eq 'pending_review') { 'Vinculo entre fontes pendente de revisao' } else { 'Vinculo entre portal e Diario consolidado' }
                detail = [string]$officialContract.crossSource.reason
                sourceLabel = 'Cruzamento de fontes'
                sourceUrl = $null
                webPdfPath = $null
                pageNumber = 0
            })
        }
    }

    if ($managementProfile) {
        foreach ($event in @($managementProfile.managerExonerationEvents)) {
            $timeline.Add([pscustomobject]@{
                eventType = 'exoneracao_gestor'
                publishedAt = [string]$event.publishedAt
                title = 'Sinal de exoneração do gestor'
                detail = [string]$event.personName
                sourceLabel = if ($event.edition) { "Diário Oficial - edição $([string]$event.edition)" } else { 'Diário Oficial' }
                sourceUrl = $null
                webPdfPath = $null
                pageNumber = [int]$event.pageNumber
            })
        }

        foreach ($event in @($managementProfile.inspectorExonerationEvents)) {
            $timeline.Add([pscustomobject]@{
                eventType = 'exoneracao_fiscal'
                publishedAt = [string]$event.publishedAt
                title = 'Sinal de exoneração do fiscal'
                detail = [string]$event.personName
                sourceLabel = if ($event.edition) { "Diário Oficial - edição $([string]$event.edition)" } else { 'Diário Oficial' }
                sourceUrl = $null
                webPdfPath = $null
                pageNumber = [int]$event.pageNumber
            })
        }
    }

    $timelineItems = @(
        $timeline |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { if ($_.PSObject.Properties['pageNumber']) { [int]$_.pageNumber } else { 0 } }; Descending = $false }
    )

    $managerNeedsReview = if ($managementProfile) {
        Test-SuspiciousManagementAssignment -Name ([string]$managementProfile.managerName) -Role ([string]$managementProfile.managerRole)
    }
    else {
        Test-SuspiciousManagementAssignment -Name ([string]$officialContract.managerName) -Role ([string]$officialContract.managerRole)
    }

    $inspectorNeedsReview = if ($managementProfile) {
        Test-SuspiciousManagementAssignment -Name ([string]$managementProfile.inspectorName) -Role ([string]$managementProfile.inspectorRole)
    }
    else {
        Test-SuspiciousManagementAssignment -Name ([string]$officialContract.inspectorName) -Role ([string]$officialContract.inspectorRole)
    }

    $alerts = New-Object System.Collections.Generic.List[string]
    if (-not $officialContract) { $alerts.Add('Sem cadastro oficial consolidado para este contrato.') }
    if ($officialContract -and [string]::IsNullOrWhiteSpace([string]$officialContract.localPdfRelative)) { $alerts.Add('Sem documento local vinculado ao cadastro oficial.') }
    if ($managementProfile -and -not [bool]$managementProfile.hasManager) { $alerts.Add('Sem gestor atual identificado na leitura contratual.') }
    if ($managementProfile -and -not [bool]$managementProfile.hasInspector) { $alerts.Add('Sem fiscal atual identificado na leitura contratual.') }
    if ($managementProfile -and [bool]$managementProfile.managerExonerationSignal) { $alerts.Add('Há sinal de exoneração associado ao gestor atual.') }
    if ($managementProfile -and [bool]$managementProfile.inspectorExonerationSignal) { $alerts.Add('Há sinal de exoneração associado ao fiscal atual.') }
    if ($managerNeedsReview) { $alerts.Add('A leitura automática do gestor requer conferência manual.') }
    if ($inspectorNeedsReview) { $alerts.Add('A leitura automática do fiscal requer conferência manual.') }

    $detailReference = if ($managementProfile) { [string]$managementProfile.contractKey } elseif ($officialContract) { [string]$officialContract.referenceKey } else { $ref }
    $workspaceHistory = @(Get-ContractWorkspaceHistory -Reference $detailReference)
    $versionComparison = Get-ContractVersionComparisonData `
        -Reference $detailReference `
        -OfficialContract $officialContract `
        -ManagementProfile $managementProfile

    $financialExecution = Get-ContractFinancialExecutionInfo `
        -OfficialContract $officialContract `
        -ManagementProfile $managementProfile `
        -RelatedMovements $relatedMovements `
        -Reference $detailReference

    $integrity = [ordered]@{
        hasOfficialRecord = [bool]$officialContract
        hasLocalDocument = [bool]($officialContract -and -not [string]::IsNullOrWhiteSpace([string]$officialContract.localPdfRelative))
        hasManagementProfile = [bool]$managementProfile
        hasRelatedMovements = ([int]@($relatedMovements).Count -gt 0)
        relatedMovements = [int]@($relatedMovements).Count
        crossMatched = [bool]($officialContract -and $officialContract.PSObject.Properties['crossSource'] -and [bool]$officialContract.crossSource.matched)
        crossStatus = if ($officialContract -and $officialContract.PSObject.Properties['crossSource']) { [string]$officialContract.crossSource.status } else { '' }
        crossDivergences = [int]@($detailDivergences).Count
        operationalAlerts = [int]@($detailAlerts).Count
        managementEvents = if ($managementProfile) { [int]@($managementProfile.managementEvents).Count } else { 0 }
        managerNeedsReview = [bool]$managerNeedsReview
        inspectorNeedsReview = [bool]$inspectorNeedsReview
        needsManagementReview = [bool]($managerNeedsReview -or $inspectorNeedsReview)
        financialSearchHints = [int]$financialExecution.searchableHintCount
        financialSources = [int]@($financialExecution.sources).Count
        parserVersion = $script:ParserVersion
        generatedAt = (Get-IsoNow)
    }

    return [ordered]@{
        generatedAt = (Get-IsoNow)
        apiContract = New-ApiContractDescriptor -Name 'contractDetail'
        apiContracts = Get-ApiContractCatalog
        parserVersion = $script:ParserVersion
        reference = $detailReference
        officialContract = $officialContract
        managementProfile = $managementProfile
        crossSource = if ($officialContract -and $officialContract.PSObject.Properties['crossSource']) { $officialContract.crossSource } else { $null }
        relatedMovements = @($relatedMovements)
        operationalAlerts = @($detailAlerts)
        crossSourceDivergences = @($detailDivergences)
        reviewQueueEntry = $reviewQueueEntry
        workspaceHistory = @($workspaceHistory)
        timeline = @($timelineItems)
        integrity = $integrity
        versionComparison = $versionComparison
        financialExecution = $financialExecution
        alerts = @($alerts)
    }
}

function Convert-PortalDateTime {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $normalized = HtmlDecode-Safe -Text $Text
    $normalized = $normalized -replace '(?<date>\d{2}/\d{2}/\d{4})\s+\S+\s+(?<time>\d{2}(?:h|:)\d{2})', '${date} as ${time}'
    $normalized = ($normalized -replace '\s+', ' ').Trim()

    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    $styles = [System.Globalization.DateTimeStyles]::AssumeLocal
    $formats = @(
        "dd/MM/yyyy 'as' HH'h'mm",
        "dd/MM/yyyy 'as' HH:mm",
        "dd/MM/yyyy HH'h'mm",
        "dd/MM/yyyy HH:mm",
        "dd/MM/yyyy"
    )

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($normalized, 'dd/MM/yyyy', $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.ToString('s')
    }

    foreach ($format in $formats) {
        if ([DateTime]::TryParseExact($normalized, $format, $culture, $styles, [ref]$parsed)) {
            return $parsed.ToString('s')
        }
    }

    return $normalized
}

function Get-LocalPdfInfo {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Diary,

        [Parameter(Mandatory = $true)]
        [string]$PdfUrl
    )

    if ([string]$Diary.publishedAt -match '^(?<year>\d{4})-') {
        $publishedYear = $matches['year']
    }
    elseif ([string]$Diary.publishedAt -match '^\d{2}/\d{2}/(?<year>20\d{2})$') {
        $publishedYear = $matches['year']
    }
    elseif ([string]$Diary.postedAtRaw -match '\b(?<year>20\d{2})\b') {
        $publishedYear = $matches['year']
    }
    else {
        $publishedYear = 'sem-ano'
    }

    $fileName = [System.IO.Path]::GetFileName(([Uri]$PdfUrl).LocalPath)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "diario-$($Diary.id)-edicao-$($Diary.edition).pdf"
    }

    $relativePath = Join-Path (Join-Path 'storage\pdfs' $publishedYear) $fileName
    $absolutePath = Join-Path $script:AppRoot $relativePath

    return [pscustomobject]@{
        relativePath = $relativePath.Replace('\', '/')
        absolutePath = $absolutePath
        webPath = ('/pdfs/' + ($publishedYear + '/' + $fileName).Replace('\', '/'))
    }
}

function Get-PdfToTextToolPath {
    $candidates = @(
        (Join-Path $script:AppRoot 'tools\xpdf-tools-win-4.06\xpdf-tools-win-4.06\bin64\pdftotext.exe'),
        (Join-Path $script:AppRoot 'tools\xpdf-tools-win-4.06\xpdf-tools-win-4.06\bin32\pdftotext.exe')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Format-BrazilianCurrency {
    param(
        [Parameter(Mandatory = $false)]
        [double]$Value = 0
    )

    return ('R$ ' + $Value.ToString('N2', [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')))
}

function Get-AppRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $appUri = [Uri]((Resolve-Path -LiteralPath $script:AppRoot).Path.TrimEnd('\') + '\')
    $targetUri = [Uri](Resolve-Path -LiteralPath $Path).Path
    $relativeUri = $appUri.MakeRelativeUri($targetUri)
    return [Uri]::UnescapeDataString($relativeUri.ToString()).Replace('\', '/')
}

$script:BackendDomainScripts = @(
    (Join-Path $PSScriptRoot 'domains\public-status-domain.ps1'),
    (Join-Path $PSScriptRoot 'domains\workspace-domain.ps1'),
    (Join-Path $PSScriptRoot 'domains\observability-domain.ps1'),
    (Join-Path $PSScriptRoot 'domains\auth-domain.ps1'),
    (Join-Path $PSScriptRoot 'domains\cross-source-domain.ps1'),
    (Join-Path $PSScriptRoot 'domains\financial-domain.ps1'),
    (Join-Path $PSScriptRoot 'domains\contracts-domain.ps1')
)

foreach ($domainScript in $script:BackendDomainScripts) {
    if (Test-Path -LiteralPath $domainScript) {
        . $domainScript
    }
}
