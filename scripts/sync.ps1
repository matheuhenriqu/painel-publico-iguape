Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'common.ps1')

$script:OrganizationIndexCache = $null
$script:PublicWorkflowMode = $env:PUBLIC_WORKFLOW_MODE -eq '1'

function Test-PublicWorkflowMode {
    return [bool]$script:PublicWorkflowMode
}

function Test-ExistingDiaryAnalysisCurrent {
    param(
        [AllowNull()]
        [object]$ExistingDiary,

        [AllowNull()]
        [object]$Analysis,

        [AllowNull()]
        [object]$NewDiaryRecord
    )

    if ($null -eq $ExistingDiary -or $null -eq $Analysis -or $null -eq $NewDiaryRecord) {
        return $false
    }

    if ([string]$Analysis.parserVersion -ne $script:ParserVersion) {
        return $false
    }

    if ([string]$Analysis.sourcePdfUrl -ne [string]$ExistingDiary.pdfUrl) {
        return $false
    }

    if ([string]$ExistingDiary.downloadTokenPath -ne [string]$NewDiaryRecord.downloadTokenPath) {
        return $false
    }

    if ([string]$ExistingDiary.publishedAt -ne [string]$NewDiaryRecord.publishedAt) {
        return $false
    }

    if ([int]$ExistingDiary.pageCount -ne [int]$NewDiaryRecord.pageCount) {
        return $false
    }

    if ([string]$ExistingDiary.fileSize -ne [string]$NewDiaryRecord.fileSize) {
        return $false
    }

    return $true
}

function Remove-Diacritics {
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

function Get-ExistingPdfLookup {
    $lookup = @{}
    $files = @(Get-ChildItem -LiteralPath $script:PdfRoot -Recurse -Filter '*.pdf' -File -ErrorAction SilentlyContinue)

    foreach ($file in $files) {
        if ($file.Length -le 0) {
            continue
        }

        if (-not $lookup.ContainsKey($file.Name)) {
            $lookup[$file.Name] = $file.FullName
        }
    }

    return $lookup
}

function Normalize-TextBlock {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $normalized = $Text -replace "`r", ''
    $normalized = Remove-Diacritics -Text $normalized
    $normalized = $normalized -replace '[ \t]+\n', "`n"
    $normalized = $normalized -replace '\n{3,}', "`n`n"
    $normalized = $normalized -replace '[ \t]{2,}', ' '
    return $normalized.Trim()
}

function Get-ActType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block
    )

    $map = [ordered]@{
        'Extrato de Contrato' = 'EXTRATO\s+DE\s+CONTRATO'
        'Termo Aditivo' = 'TERMO\s+ADITIVO|ADITIVO\s+AO\s+CONTRATO'
        'Apostilamento' = 'APOSTILAMENTO|TERMO\s+DE\s+APOSTILAMENTO'
        'Dispensa' = 'DISPENSA\s+DE\s+LICITACAO|RATIFICACAO\s+DE\s+DISPENSA'
        'Inexigibilidade' = 'INEXIGIBILIDADE\s+DE\s+LICITACAO|RATIFICACAO\s+DE\s+INEXIGIBILIDADE'
        'Registro de Precos' = 'ATA\s+DE\s+REGISTRO\s+DE\s+PRECOS|REGISTRO\s+DE\s+PRECOS'
        'Rescisao' = 'RESCISAO\s+CONTRATUAL|TERMO\s+DE\s+RESCISAO'
        'Homologacao' = 'HOMOLOGACAO(?:\s+E\s+ADJUDICACAO)?|ADJUDICACAO(?:\s+E\s+HOMOLOGACAO)?'
        'Contrato' = 'CONTRATO\s+N[O0]'
    }

    foreach ($label in $map.Keys) {
        if ($Block -match $map[$label]) {
            return $label
        }
    }

    return 'Ato Contratual'
}

function Get-FirstPatternValue {
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

        if ($match.Success) {
            $value = if ($match.Groups['value'].Success) { $match.Groups['value'].Value } else { $match.Groups[1].Value }
            return (($value -replace '\s+', ' ').Trim())
        }
    }

    return ''
}

function Get-TextBlocksFromPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageText
    )

    $headerPattern = '(?m)(^|\n)(EXTRATO\s+DE\s+CONTRATO|TERMO\s+ADITIVO|ADITIVO\s+AO\s+CONTRATO|APOSTILAMENTO|TERMO\s+DE\s+APOSTILAMENTO|DISPENSA\s+DE\s+LICITACAO|RATIFICACAO\s+DE\s+DISPENSA|INEXIGIBILIDADE\s+DE\s+LICITACAO|RATIFICACAO\s+DE\s+INEXIGIBILIDADE|ATA\s+DE\s+REGISTRO\s+DE\s+PRECOS|RESCISAO\s+CONTRATUAL|TERMO\s+DE\s+RESCISAO|HOMOLOGACAO(?:\s+E\s+ADJUDICACAO)?|ADJUDICACAO(?:\s+E\s+HOMOLOGACAO)?|CONTRATO\s+N[O0])'
    $matches = [regex]::Matches($PageText, $headerPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    if ($matches.Count -eq 0) {
        if ($PageText -match 'contrato|dispensa|inexigibilidade|aditivo|apostilamento|homologacao|registro de precos|rescisao') {
            return @($PageText)
        }

        return @()
    }

    $blocks = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $matches.Count; $index++) {
        $start = $matches[$index].Index
        $end = if ($index + 1 -lt $matches.Count) { $matches[$index + 1].Index } else { $PageText.Length }
        $length = $end - $start
        if ($length -le 0) {
            continue
        }

        $block = $PageText.Substring($start, $length).Trim()
        if ($block.Length -gt 30) {
            $blocks.Add($block)
        }
    }

    return $blocks.ToArray()
}

function New-ContractItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block,

        [Parameter(Mandatory = $true)]
        [int]$PageNumber
    )

    $contractNumber = Get-FirstPatternValue -Text $Block -Patterns @(
        'CONTRATO\s+N[O0]?\s*(?<value>[A-Z0-9\-\/\.]+)',
        'CT\s+N[O0]?\s*(?<value>[A-Z0-9\-\/\.]+)'
    )

    $processNumber = Get-FirstPatternValue -Text $Block -Patterns @(
        'PROCESSO(?:\s+ADMINISTRATIVO)?\s+N[O0]?\s*(?<value>[A-Z0-9\-\/\.]+)',
        'PROC\.\s*N[O0]?\s*(?<value>[A-Z0-9\-\/\.]+)'
    )

    $modality = Get-FirstPatternValue -Text $Block -Patterns @(
        'MODALIDADE\s*:?\s*(?<value>[^\n\.;]+)',
        '(?<value>(?:PREGAO(?:\s+ELETRONICO|\s+PRESENCIAL)?|DISPENSA(?:\s+DE\s+LICITACAO)?|INEXIGIBILIDADE(?:\s+DE\s+LICITACAO)?|CONCORRENCIA(?:\s+ELETRONICA)?|TOMADA\s+DE\s+PRECOS|CREDENCIAMENTO)[^\n\.;]*)'
    )

    $contractor = Get-FirstPatternValue -Text $Block -Patterns @(
        'CONTRATAD[AO]\s*:?\s*(?<value>[^\n]+?)(?=\s*(?:CNPJ|CPF|OBJETO|VALOR|VIGENCIA|DATA|ASSINATURA|PROCESSO|MODALIDADE)\s*:|\n{2,}|$)',
        'FAVORECID[AO]\s*:?\s*(?<value>[^\n]+?)(?=\s*(?:CNPJ|CPF|OBJETO|VALOR|VIGENCIA|DATA|ASSINATURA|PROCESSO|MODALIDADE)\s*:|\n{2,}|$)',
        'DETENTOR[AO]\s*:?\s*(?<value>[^\n]+?)(?=\s*(?:CNPJ|CPF|OBJETO|VALOR|VIGENCIA|DATA|ASSINATURA|PROCESSO|MODALIDADE)\s*:|\n{2,}|$)',
        'EMPRESA\s*:?\s*(?<value>[^\n]+?)(?=\s*(?:CNPJ|CPF|OBJETO|VALOR|VIGENCIA|DATA|ASSINATURA|PROCESSO|MODALIDADE)\s*:|\n{2,}|$)'
    )

    $cnpj = Get-FirstPatternValue -Text $Block -Patterns @(
        '(?<value>\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2})'
    )

    $objectValue = Get-FirstPatternValue -Text $Block -Patterns @(
        'OBJETO\s*:?\s*(?<value>[\s\S]{20,400}?)(?=\n[A-Z ]{3,}\s*:|\n{2,}|VALOR|VIGENCIA|ASSINATURA|PRAZO|MODALIDADE|PROCESSO|CONTRATAD[AO]|CNPJ|CPF|$)'
    )

    $term = Get-FirstPatternValue -Text $Block -Patterns @(
        'VIGENCIA\s*:?\s*(?<value>[^\n\.;]+)',
        'PRAZO(?:\s+DE\s+VIGENCIA)?\s*:?\s*(?<value>[^\n\.;]+)'
    )

    $signatureDate = Get-FirstPatternValue -Text $Block -Patterns @(
        'ASSINATURA\s*:?\s*(?<value>[^\n\.;]+)',
        'DATA\s+DA\s+ASSINATURA\s*:?\s*(?<value>[^\n\.;]+)'
    )

    $legalBasis = Get-FirstPatternValue -Text $Block -Patterns @(
        'FUNDAMENTO\s+LEGAL\s*:?\s*(?<value>[^\n\.;]+)',
        '(?<value>Lei\s+Federal\s+n[o0]?\s*14\.133\/2021[^\n\.;]*)',
        '(?<value>Lei\s+n[o0]?\s*8\.666\/93[^\n\.;]*)'
    )

    $valueMatch = [regex]::Match($Block, 'R\$\s?[\d\.\,]+')
    $value = if ($valueMatch.Success) { $valueMatch.Value } else { '' }

    $flags = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($contractNumber)) { $flags.Add('Sem numero do contrato') }
    if ([string]::IsNullOrWhiteSpace($contractor)) { $flags.Add('Sem contratada identificada') }
    if ([string]::IsNullOrWhiteSpace($objectValue)) { $flags.Add('Sem objeto identificado') }
    if ([string]::IsNullOrWhiteSpace($value)) { $flags.Add('Sem valor identificado') }

    $completenessScore = 0
    foreach ($field in @($contractNumber, $contractor, $objectValue, $value, $signatureDate, $term)) {
        if (-not [string]::IsNullOrWhiteSpace($field)) {
            $completenessScore++
        }
    }

    if ($completenessScore -ge 5) {
        $completeness = 'alta'
    }
    elseif ($completenessScore -ge 3) {
        $completeness = 'media'
    }
    else {
        $completeness = 'baixa'
    }

    return [pscustomobject]@{
        type = (Get-ActType -Block $Block)
        pageNumber = $PageNumber
        contractNumber = $contractNumber
        processNumber = $processNumber
        modality = $modality
        contractor = $contractor
        cnpj = $cnpj
        object = $objectValue
        value = $value
        term = $term
        signatureDate = $signatureDate
        legalBasis = $legalBasis
        excerpt = $Block.Substring(0, [Math]::Min(460, $Block.Length))
        completeness = $completeness
        flags = $flags.ToArray()
    }
}

function Convert-PdfToPageTexts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PdfPath,

        [Parameter(Mandatory = $true)]
        [string]$PdfToTextTool
    )

    $rawText = & $PdfToTextTool -enc UTF-8 -layout $PdfPath - 2>$null
    $joined = ($rawText | Out-String)
    if ([string]::IsNullOrWhiteSpace($joined)) {
        return @()
    }

    $pages = $joined -split [char]12

    $pageTexts = New-Object System.Collections.Generic.List[object]
    $pageNumber = 1

    foreach ($page in $pages) {
        if ([string]::IsNullOrWhiteSpace([string]$page)) {
            $pageNumber++
            continue
        }

        $normalized = Normalize-TextBlock -Text ([string]$page)
        if ([string]::IsNullOrWhiteSpace($normalized)) {
            $pageNumber++
            continue
        }

        $pageTexts.Add([pscustomobject]@{
            pageNumber = $pageNumber
            text = $normalized
        })

        $pageNumber++
    }

    return $pageTexts.ToArray()
}

function Get-UniqueContractItems {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items
    )

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $deduped = New-Object System.Collections.Generic.List[object]

    foreach ($item in $Items) {
        $objectText = [string]$item.object
        $objectKey = $objectText.Substring(0, [Math]::Min(80, $objectText.Length))
        $key = '{0}|{1}|{2}|{3}|{4}' -f $item.type, $item.contractNumber, $item.contractor, $objectKey, $item.pageNumber

        if ($seen.Add($key)) {
            $deduped.Add($item)
        }
    }

    return $deduped.ToArray()
}

function Invoke-ServerSideAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Diaries
    )

    $pdfToTextTool = Get-PdfToTextToolPath
    if (-not $pdfToTextTool) {
        Update-Status -Updates @{
            message = 'Ferramenta pdftotext nao encontrada; analise textual foi ignorada.'
        }
        return
    }

    $candidates = @($Diaries | Where-Object { @($_.candidateKeywords).Count -gt 0 -and $_.localPdfPath -and (Test-Path -LiteralPath $_.localPdfPath) })
    $totalCandidates = @($candidates).Count
    $currentIndex = 0

    foreach ($diary in $candidates) {
        $currentIndex++
        $analysis = Get-ExistingAnalysis -DiaryId ([string]$diary.id)
        if (Test-AnalysisCurrent -Diary $diary -Analysis $analysis) {
            continue
        }

        Update-Status -Updates @{
            syncStage = 'analisando'
            message = "Analisando contratos: edicao $($diary.edition) ($currentIndex de $totalCandidates)."
        }

        try {
            $pageTexts = @(Convert-PdfToPageTexts -PdfPath $diary.localPdfPath -PdfToTextTool $pdfToTextTool)
            $items = New-Object System.Collections.Generic.List[object]

            foreach ($page in $pageTexts) {
                $blocks = Get-TextBlocksFromPage -PageText $page.text
                foreach ($block in $blocks) {
                    if ($block -notmatch 'contrato|dispensa|inexigibilidade|aditivo|apostilamento|homologacao|registro de precos|rescisao') {
                        continue
                    }

                    $items.Add((New-ContractItem -Block $block -PageNumber $page.pageNumber))
                }
            }

            $dedupedItems = @(Get-UniqueContractItems -Items $items.ToArray())
            $totalValue = 0.0
            foreach ($item in $dedupedItems) {
                $totalValue += (Convert-BrazilianCurrencyToNumber -Text ([string]$item.value))
            }

            $analysisPayload = [ordered]@{
                diaryId = [string]$diary.id
                parserVersion = $script:ParserVersion
                analyzedAt = Get-IsoNow
                sourcePdfUrl = [string]$diary.pdfUrl
                sourceLocalPdfRelative = [string]$diary.localPdfRelative
                keywords = @($diary.candidateKeywords)
                summary = [ordered]@{
                    pageCount = @($pageTexts).Count
                    itemCount = @($dedupedItems).Count
                    totalValue = (Format-BrazilianCurrency -Value $totalValue)
                }
                items = @($dedupedItems)
            }
        }
        catch {
            $analysisPayload = [ordered]@{
                diaryId = [string]$diary.id
                parserVersion = $script:ParserVersion
                analyzedAt = Get-IsoNow
                sourcePdfUrl = [string]$diary.pdfUrl
                sourceLocalPdfRelative = [string]$diary.localPdfRelative
                keywords = @($diary.candidateKeywords)
                summary = [ordered]@{
                    pageCount = 0
                    itemCount = 0
                    totalValue = (Format-BrazilianCurrency -Value 0)
                    error = $_.Exception.Message
                }
                items = @()
            }
        }

        Write-JsonFile -Path (Get-AnalysisPath -DiaryId ([string]$diary.id)) -Data $analysisPayload
    }

    $analysisCount = @(Get-ChildItem -LiteralPath $script:AnalysisRoot -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    Update-Status -Updates @{
        analyzedDiaries = $analysisCount
        pendingAnalysis = [Math]::Max($totalCandidates - $analysisCount, 0)
    }
}

function Merge-Diaries {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$CollectedDiaries,

        [Parameter(Mandatory = $true)]
        [hashtable]$CandidateMap
    )

    $existingById = Get-DiariesById
    $existingPdfLookup = Get-ExistingPdfLookup
    $merged = New-Object System.Collections.Generic.List[object]
    $newCount = 0
    $updatedCount = 0
    $downloadedCount = 0
    $currentIndex = 0
    $totalDiaries = @($CollectedDiaries).Count

    foreach ($diary in ($CollectedDiaries | Sort-Object publishedAt -Descending)) {
        $currentIndex++
        $id = [string]$diary.id
        $existing = if ($existingById.ContainsKey($id)) { $existingById[$id] } else { $null }
        $candidateKeywords = if ($CandidateMap.ContainsKey($id)) { @($CandidateMap[$id]) } else { @() }

        Update-Status -Updates @{
            syncStage = 'baixando-pdfs'
            message = "Processando PDF $currentIndex de $totalDiaries."
        }

        $record = [ordered]@{
            id = $id
            edition = [string]$diary.edition
            isExtra = [bool]$diary.isExtra
            viewPath = [string]$diary.viewPath
            viewUrl = [string]$diary.viewUrl
            downloadTokenPath = [string]$diary.downloadTokenPath
            downloadTokenUrl = [string]$diary.downloadTokenUrl
            postedAtRaw = [string]$diary.postedAtRaw
            publishedAt = [string]$diary.publishedAt
            pageCount = [int]$diary.pageCount
            fileSize = [string]$diary.fileSize
            sizeLabel = [string]$diary.sizeLabel
            candidateKeywords = @($candidateKeywords)
            syncedAt = (Get-IsoNow)
            pdfUrl = if ($existing) { [string]$existing.pdfUrl } else { $null }
            localPdfRelative = if ($existing) { [string]$existing.localPdfRelative } else { $null }
            localPdfPath = if ($existing) { [string]$existing.localPdfPath } else { $null }
            webPdfPath = if ($existing) { [string]$existing.webPdfPath } else { $null }
            pdfDownloadedAt = if ($existing) { [string]$existing.pdfDownloadedAt } else { $null }
        }

        if ($null -eq $existing) {
            $newCount++
        }
        elseif (
            [string]$existing.downloadTokenPath -ne [string]$record.downloadTokenPath -or
            [string]$existing.publishedAt -ne [string]$record.publishedAt -or
            [int]$existing.pageCount -ne [int]$record.pageCount -or
            [string]$existing.fileSize -ne [string]$record.fileSize
        ) {
            $updatedCount++
        }

        if (
            -not [string]::IsNullOrWhiteSpace($record.localPdfPath) -and
            (Test-Path -LiteralPath $record.localPdfPath) -and
            ((Get-Item -LiteralPath $record.localPdfPath).Length -gt 0) -and
            -not [string]::IsNullOrWhiteSpace($record.pdfUrl)
        ) {
            $canonicalPdfInfo = Get-LocalPdfInfo -Diary ([pscustomobject]$record) -PdfUrl ([string]$record.pdfUrl)
            if ($record.localPdfPath -ne $canonicalPdfInfo.absolutePath) {
                Ensure-Directory -Path (Split-Path -Parent $canonicalPdfInfo.absolutePath)
                if (-not (Test-Path -LiteralPath $canonicalPdfInfo.absolutePath)) {
                    Move-Item -LiteralPath $record.localPdfPath -Destination $canonicalPdfInfo.absolutePath -Force
                }

                if (Test-Path -LiteralPath $canonicalPdfInfo.absolutePath) {
                    $record.localPdfPath = $canonicalPdfInfo.absolutePath
                    $record.localPdfRelative = $canonicalPdfInfo.relativePath
                    $record.webPdfPath = $canonicalPdfInfo.webPath
                }
            }

            $merged.Add([pscustomobject]$record)
            continue
        }

        if (Test-PublicWorkflowMode) {
            $existingAnalysis = if (@($candidateKeywords).Count -gt 0) { Get-ExistingAnalysis -DiaryId $id } else { $null }
            if (Test-ExistingDiaryAnalysisCurrent -ExistingDiary $existing -Analysis $existingAnalysis -NewDiaryRecord ([pscustomobject]$record)) {
                $record.pdfUrl = [string]$existing.pdfUrl
                $record.localPdfRelative = [string]$existing.localPdfRelative
                $record.localPdfPath = [string]$existing.localPdfPath
                $record.webPdfPath = [string]$existing.webPdfPath
                $record.pdfDownloadedAt = [string]$existing.pdfDownloadedAt
                $merged.Add([pscustomobject]$record)
                continue
            }
        }

        $pdfUrl = Get-PdfRedirectUrl -DownloadTokenPath $record.downloadTokenPath
        $pdfInfo = Get-LocalPdfInfo -Diary ([pscustomobject]$record) -PdfUrl $pdfUrl
        $pdfFileName = [System.IO.Path]::GetFileName($pdfInfo.absolutePath)

        if ($existingPdfLookup.ContainsKey($pdfFileName)) {
            $existingPdfPath = $existingPdfLookup[$pdfFileName]
            $record.pdfUrl = $pdfUrl
            $record.localPdfPath = $existingPdfPath
            $record.localPdfRelative = Get-AppRelativePath -Path $existingPdfPath
            $record.webPdfPath = ('/' + $record.localPdfRelative.Replace('storage/pdfs/', 'pdfs/'))
            $record.pdfDownloadedAt = (Get-IsoNow)
            $merged.Add([pscustomobject]$record)
            continue
        }

        Ensure-Directory -Path (Split-Path -Parent $pdfInfo.absolutePath)
        Invoke-PortalRequest -PathOrUrl $pdfUrl -OutFile $pdfInfo.absolutePath | Out-Null

        $record.pdfUrl = $pdfUrl
        $record.localPdfRelative = $pdfInfo.relativePath
        $record.localPdfPath = $pdfInfo.absolutePath
        $record.webPdfPath = $pdfInfo.webPath
        $record.pdfDownloadedAt = (Get-IsoNow)
        $downloadedCount++

        $merged.Add([pscustomobject]$record)
    }

    return [pscustomobject]@{
        diaries = $merged.ToArray()
        newCount = $newCount
        updatedCount = $updatedCount
        downloadedCount = $downloadedCount
    }
}

function Get-AllDiariesFromPortal {
    $all = New-Object System.Collections.Generic.List[object]
    $firstHtml = Invoke-PortalRequest -PathOrUrl (Get-DiarioPageUrl -Page 1)
    $firstPage = Parse-DiarioListingPage -Html $firstHtml
    $totalPages = [Math]::Max([int]$firstPage.totalPages, 1)

    Update-Status -Updates @{
        syncStage = 'catalogando'
        message = 'Catalogando todas as edicoes do Diario Oficial.'
        totalPages = $totalPages
        scannedPages = 1
        totalDiaries = $firstPage.totalResults
    }

    foreach ($item in @($firstPage.items)) {
        $all.Add($item)
    }

    for ($page = 2; $page -le $totalPages; $page++) {
        $html = Invoke-PortalRequest -PathOrUrl (Get-DiarioPageUrl -Page $page)
        $parsed = Parse-DiarioListingPage -Html $html

        foreach ($item in @($parsed.items)) {
            $all.Add($item)
        }

        Update-Status -Updates @{
            scannedPages = $page
            message = "Catalogando edicoes: pagina $page de $totalPages."
        }

        Start-Sleep -Milliseconds 120
    }

    return $all.ToArray()
}

function Get-CandidateDiaryMap {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Keywords
    )

    $map = @{}

    foreach ($keyword in $Keywords) {
        Update-Status -Updates @{
            syncStage = 'mapeando-contratos'
            message = "Buscando edicoes com o termo '$keyword'."
        }

        $firstHtml = Invoke-PortalRequest -PathOrUrl (Get-DiarioPageUrl -Page 1 -Keyword $keyword)
        $firstPage = Parse-DiarioListingPage -Html $firstHtml
        $totalPages = [Math]::Max([int]$firstPage.totalPages, 1)

        foreach ($item in @($firstPage.items)) {
            if (-not $map.ContainsKey([string]$item.id)) {
                $map[[string]$item.id] = New-Object System.Collections.Generic.List[string]
            }

            if (-not $map[[string]$item.id].Contains($keyword)) {
                $null = $map[[string]$item.id].Add($keyword)
            }
        }

        for ($page = 2; $page -le $totalPages; $page++) {
            $html = Invoke-PortalRequest -PathOrUrl (Get-DiarioPageUrl -Page $page -Keyword $keyword)
            $parsed = Parse-DiarioListingPage -Html $html

            foreach ($item in @($parsed.items)) {
                if (-not $map.ContainsKey([string]$item.id)) {
                    $map[[string]$item.id] = New-Object System.Collections.Generic.List[string]
                }

                if (-not $map[[string]$item.id].Contains($keyword)) {
                    $null = $map[[string]$item.id].Add($keyword)
                }
            }

            Update-Status -Updates @{
                message = "Termo '$keyword': pagina $page de $totalPages."
            }

            Start-Sleep -Milliseconds 120
        }
    }

    $normalized = @{}
    foreach ($key in $map.Keys) {
        $normalized[$key] = @($map[$key] | Sort-Object -Unique)
    }

    return $normalized
}

function Normalize-SearchText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $normalized = Remove-Diacritics -Text $Text
    $normalized = $normalized.ToUpperInvariant()
    $normalized = $normalized -replace '[^A-Z0-9/\s]', ' '
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Expand-OrganizationAliasSet {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Aliases,

        [Parameter(Mandatory = $true)]
        [string]$Sphere,

        [Parameter(Mandatory = $true)]
        [string]$Kind
    )

    $expanded = New-Object System.Collections.Generic.List[string]

    foreach ($alias in @($Aliases)) {
        if ([string]::IsNullOrWhiteSpace([string]$alias)) {
            continue
        }

        $seed = ([string]$alias).Trim()
        if (-not $expanded.Contains($seed)) {
            $expanded.Add($seed)
        }

        if ($Sphere -ne 'municipal' -or $Kind -ne 'secretaria') {
            continue
        }

        $normalizedSeed = Normalize-SearchText -Text $seed
        if ($normalizedSeed -match '^SECRETARIA(?: MUNICIPAL)? (DE|DA|DO) (.+)$') {
            $preposition = $matches[1]
            $tail = $matches[2]
            foreach ($variation in @(
                "SECRETARIA $preposition $tail",
                "SECRETARIA MUNICIPAL $preposition $tail",
                "SECRETARIO MUNICIPAL $preposition $tail",
                "SECRETARIA MUNIC $preposition $tail",
                "SECRETARIO MUNIC $preposition $tail",
                "SECRETARIA ADJUNTA $preposition $tail",
                "SECRETARIO ADJUNTO $preposition $tail",
                "SECRETARIA MUNICIPAL ADJUNTA $preposition $tail",
                "SECRETARIO MUNICIPAL ADJUNTO $preposition $tail",
                "SECRETARIA MUNICIPAL ADJUNTO $preposition $tail",
                "SECRETARIO MUNICIPAL ADJUNTA $preposition $tail"
            )) {
                if (-not $expanded.Contains($variation)) {
                    $expanded.Add($variation)
                }
            }
        }
    }

    return $expanded.ToArray()
}

function Get-OrganizationIndex {
    if ($script:OrganizationIndexCache) {
        return $script:OrganizationIndexCache
    }

    $catalog = Get-OrganizationCatalog
    $entries = New-Object System.Collections.Generic.List[object]

    foreach ($organization in @($catalog.organizations)) {
        $aliases = New-Object System.Collections.Generic.List[string]
        $sourceAliases = Expand-OrganizationAliasSet `
            -Aliases (@([string]$organization.name) + @($organization.aliases)) `
            -Sphere ([string]$organization.sphere) `
            -Kind ([string]$organization.kind)

        foreach ($alias in @($sourceAliases)) {
            $normalizedAlias = Normalize-SearchText -Text ([string]$alias)
            if ([string]::IsNullOrWhiteSpace($normalizedAlias)) {
                continue
            }

            if ($normalizedAlias.Length -lt 10 -and $normalizedAlias -notin @('MEC', 'FNDE', 'AGU', 'PGM', 'PGE', 'SMIURB', 'DIVTRAN')) {
                continue
            }

            if (-not $aliases.Contains($normalizedAlias)) {
                $aliases.Add($normalizedAlias)
            }
        }

        $entries.Add([pscustomobject]@{
            id = [string]$organization.id
            sphere = [string]$organization.sphere
            kind = [string]$organization.kind
            areaId = [string]$organization.areaId
            name = [string]$organization.name
            aliases = @($aliases | Sort-Object -Unique | Sort-Object Length -Descending)
            maxAliasLength = if ($aliases.Count -gt 0) { ($aliases | Measure-Object -Property Length -Maximum).Maximum } else { 0 }
        })
    }

    $orderedEntries = @(
        $entries.ToArray() |
        Sort-Object -Property @{ Expression = { $_.maxAliasLength }; Descending = $true }, @{ Expression = { $_.name }; Descending = $false }
    )

    $script:OrganizationIndexCache = [pscustomobject]@{
        catalog = $catalog
        organizations = $orderedEntries
    }

    return $script:OrganizationIndexCache
}

function Find-OrganizationsInText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $normalizedText = Normalize-SearchText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalizedText)) {
        return @()
    }

    $index = Get-OrganizationIndex
    $matches = New-Object System.Collections.Generic.List[object]

    foreach ($organization in @($index.organizations)) {
        $bestIndex = $null
        $bestAlias = ''

        foreach ($alias in @($organization.aliases)) {
            $matchIndex = $normalizedText.IndexOf($alias)
            if ($matchIndex -ge 0) {
                if ($null -eq $bestIndex -or $matchIndex -lt $bestIndex -or ($matchIndex -eq $bestIndex -and $alias.Length -gt $bestAlias.Length)) {
                    $bestIndex = $matchIndex
                    $bestAlias = $alias
                }
            }
        }

        if ($null -ne $bestIndex) {
            $matches.Add([pscustomobject]@{
                id = [string]$organization.id
                sphere = [string]$organization.sphere
                kind = [string]$organization.kind
                areaId = [string]$organization.areaId
                name = [string]$organization.name
                matchedAlias = [string]$bestAlias
                matchIndex = [int]$bestIndex
            })
        }
    }

    return @(
        $matches.ToArray() |
        Sort-Object -Property `
            @{ Expression = { [int]$_.matchIndex }; Descending = $false }, `
            @{ Expression = { if ([string]$_.sphere -eq 'municipal') { 0 } else { 1 } }; Descending = $false }, `
            @{ Expression = { if ([string]$_.id -eq 'municipal-prefeitura') { 1 } else { 0 } }; Descending = $false }, `
            @{ Expression = { [string]$_.name }; Descending = $false }
    )
}

function Merge-OrganizationMatches {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$Matches
    )

    $merged = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($match in @($Matches | Sort-Object -Property @{ Expression = { [int]$_.matchIndex }; Descending = $false })) {
        if ($seen.Add([string]$match.id)) {
            $merged.Add($match)
        }
    }

    return $merged.ToArray()
}

function Get-PreferredPrimaryOrganization {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$BlockMatches,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$PageMatches
    )

    $safeBlockMatches = @($BlockMatches | Where-Object { $null -ne $_ -and $_.PSObject.Properties['id'] -and $_.PSObject.Properties['sphere'] })
    $safePageMatches = @($PageMatches | Where-Object { $null -ne $_ -and $_.PSObject.Properties['id'] -and $_.PSObject.Properties['sphere'] })

    $preferenceGroups = @(
        @($safeBlockMatches | Where-Object { $_.sphere -eq 'municipal' -and $_.id -ne 'municipal-prefeitura' } | Sort-Object -Property @{ Expression = { if ([string]$_.kind -eq 'secretaria') { 0 } else { 1 } }; Descending = $false }, @{ Expression = { [int]$_.matchIndex }; Descending = $false }),
        @($safePageMatches | Where-Object { $_.sphere -eq 'municipal' -and $_.id -ne 'municipal-prefeitura' } | Sort-Object -Property @{ Expression = { if ([string]$_.kind -eq 'secretaria') { 0 } else { 1 } }; Descending = $false }, @{ Expression = { [int]$_.matchIndex }; Descending = $false }),
        @($safeBlockMatches | Where-Object { $_.sphere -eq 'municipal' } | Sort-Object -Property @{ Expression = { if ([string]$_.id -eq 'municipal-prefeitura') { 1 } else { 0 } }; Descending = $false }, @{ Expression = { [int]$_.matchIndex }; Descending = $false }),
        @($safePageMatches | Where-Object { $_.sphere -eq 'municipal' } | Sort-Object -Property @{ Expression = { if ([string]$_.id -eq 'municipal-prefeitura') { 1 } else { 0 } }; Descending = $false }, @{ Expression = { [int]$_.matchIndex }; Descending = $false }),
        @(Merge-OrganizationMatches -Matches (@($safeBlockMatches) + @($safePageMatches)))
    )

    foreach ($group in $preferenceGroups) {
        if (@($group).Count -gt 0) {
            return @($group)[0]
        }
    }

    return $null
}

function Clean-ObjectText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $clean = ($Text -replace '\s+', ' ').Trim()
    $clean = $clean -replace '^(OBJETO\s*:?\s*)', ''
    $clean = $clean -replace '^(O OBJETO(?: DESTE CONTRATO)? E\s+)', ''
    $clean = $clean -replace '^(DESTE CONTRATO E\s+)', ''
    $clean = $clean -replace '^(DO CONTRATO\s+\d{1,5}/\d{4},\s*)', ''
    $clean = $clean -replace '^(CONTRATO\s+\d{1,5}/\d{4},\s*)', ''
    $clean = $clean -replace '\bArt\.?\s*2.*$', ''
    $clean = $clean -replace '\bESTA PORTARIA\b.*$', ''
    $clean = $clean -replace '\bGABINETE DO PREFEITO\b.*$', ''
    return $clean.Trim(' -,:;.')
}

function Refine-ActTitle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [string]$Block,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$ContractNumber = '',

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$ProcessNumber = ''
    )

    $clean = (($Title -replace '\s+', ' ').Trim(' -'))
    $portariaTitle = Get-FirstPatternValue -Text $Block -Patterns @(
        '(?<value>PORTARIA\s+N\S*\s*:?\s*\d+[^\n]*)'
    )

    if ($clean -match '^Art\.?' -and -not [string]::IsNullOrWhiteSpace($portariaTitle)) {
        $clean = (($portariaTitle -replace '\s+', ' ').Trim(' -'))
    }

    if ($Type -eq 'Termo Aditivo') {
        $trimmedAditivo = Get-FirstPatternValue -Text $clean -Patterns @(
            '(?<value>.*?CONTRATO(?:\s+N\S*\s*:?)?\s*\d{1,5}/\d{4})'
        )
        if (-not [string]::IsNullOrWhiteSpace($trimmedAditivo)) {
            $clean = $trimmedAditivo
        }
    }

    if ($Type -eq 'Contrato' -and ([string]::IsNullOrWhiteSpace($clean) -or $clean -match '^CONTRATANTE')) {
        if (-not [string]::IsNullOrWhiteSpace($ContractNumber)) {
            $clean = "Contrato $ContractNumber"
        }
        else {
            $clean = 'Contrato publicado'
        }
    }

    if ($Type -eq 'Gestao Contratual' -and $clean -notmatch '^PORTARIA' -and -not [string]::IsNullOrWhiteSpace($ContractNumber)) {
        $clean = "Portaria de gestao do contrato $ContractNumber"
    }

    if ([string]::IsNullOrWhiteSpace($clean) -and -not [string]::IsNullOrWhiteSpace($ProcessNumber)) {
        $clean = "Ato do processo $ProcessNumber"
    }

    return (($clean -replace '\s+', ' ').Trim(' -'))
}

function Get-PageContextText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageText
    )

    $markerMatch = [regex]::Match($PageText, 'CONTRATOS E LICITACAO', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($markerMatch.Success) {
        return $PageText.Substring(0, $markerMatch.Index)
    }

    $lines = @($PageText -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return (@($lines | Select-Object -First 24) -join "`n")
}

function Test-StructuredBlockStart {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line
    )

    $normalized = Normalize-SearchText -Text $Line
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $false
    }

    return $normalized -match '^(EXTRATO(?: DE CONTRATO)?|[0-9]+ TERMO ADITIVO|TERMO ADITIVO|APOSTILAMENTO|ATA DE REGISTRO DE PRECOS|DISPENSA DE LICITACAO|INEXIGIBILIDADE|ADJUDICACAO HOMOLOGACAO|AVISO DE LICITACAO|AVISO DE RETI|PREGAO ELETRONICO|CONCORRENCIA|CHAMADA PUBLICA|RATIFICACAO)'
}

function Split-StructuredBlocks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $lines = @($Text -split "`n")
    $blocks = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines) {
        $trimmed = [string]$line
        if (Test-StructuredBlockStart -Line $trimmed) {
            if (@($current).Count -gt 0) {
                $block = (@($current) -join "`n").Trim()
                if ($block.Length -gt 40) {
                    $blocks.Add($block)
                }
                $current.Clear()
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($trimmed) -or @($current).Count -gt 0) {
            $current.Add($trimmed)
        }
    }

    if (@($current).Count -gt 0) {
        $block = (@($current) -join "`n").Trim()
        if ($block.Length -gt 40) {
            $blocks.Add($block)
        }
    }

    return @($blocks)
}

function Get-ManagementBlocksFromPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageText
    )

    $parts = $PageText -split '(?m)(?=^\s*PORTARIA\s+N)'
    $blocks = New-Object System.Collections.Generic.List[string]

    foreach ($part in @($parts)) {
        $normalized = Normalize-SearchText -Text $part
        if ($normalized -match 'GESTOR(?: E FISCAL)? DO CONTRATO|FISCAL DO CONTRATO') {
            $block = $part.Trim()
            if ($block.Length -gt 40) {
                $blocks.Add($block)
            }
        }
    }

    return @($blocks)
}

function Test-RelevantContractBlock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block
    )

    $normalized = Normalize-SearchText -Text $Block
    $hasManagement = $normalized -match 'GESTOR(?: E FISCAL)? DO CONTRATO|FISCAL DO CONTRATO'
    $hasStrongMarker = $normalized -match 'CONTRATANTE|CONTRATADA|FORNECEDOR|DETENTORA|OBJETO|VALOR|PREGAO ELETRONICO|CONCORRENCIA|DISPENSA DE LICITACAO|INEXIGIBILIDADE|ATA DE REGISTRO DE PRECOS|TERMO ADITIVO|EXTRATO DE CONTRATO|ADJUDICACAO|HOMOLOGACAO|AVISO DE LICITACAO|RATIFICACAO|CHAMADA PUBLICA|CONTRATO N'
    $hasForbiddenOnly = $normalized -match 'PROCESSO SELETIVO|CONVOCACAO|CLASSIFICACAO FINAL|RESULTADO DE DESEMPATE|EXONERACAO|NOMEACAO|RESOLUCAO|DECRETO'

    if ($hasManagement) {
        return $true
    }

    if ($hasStrongMarker) {
        return $true
    }

    if ($hasForbiddenOnly) {
        return $false
    }

    return $false
}

function Get-TextBlocksFromPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageText
    )

    $normalizedPage = Normalize-SearchText -Text $PageText
    if ($normalizedPage.Contains('CONTRATOS E LICITACAO')) {
        $markerMatch = [regex]::Match($PageText, 'CONTRATOS E LICITACAO', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $sectionText = if ($markerMatch.Success) { $PageText.Substring($markerMatch.Index + $markerMatch.Length) } else { $PageText }
        return @(
            Split-StructuredBlocks -Text $sectionText |
            Where-Object { Test-RelevantContractBlock -Block $_ }
        )
    }

    return @(
        Get-ManagementBlocksFromPage -PageText $PageText |
        Where-Object { Test-RelevantContractBlock -Block $_ }
    )
}

function Get-FirstMeaningfulLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block
    )

    $lines = @($Block -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $usable = @()

    foreach ($line in $lines) {
        $normalized = Normalize-SearchText -Text $line
        if ($normalized -in @('EXTRATO', 'CONTRATOS E LICITACAO', 'PORTARIA', 'DIARIO OFICIAL', 'IGUAPE SP')) {
            continue
        }

        if ($normalized -match '^ANO [0-9]' -or $normalized -match '^(CADERNO EXECUTIVO|ADMINISTRACAO GERAL)$') {
            continue
        }

        $usable += $line
        if ($usable.Count -ge 2) {
            break
        }
    }

    if ($usable.Count -eq 0) {
        return 'Ato contratual'
    }

    if ($usable.Count -ge 2 -and (Normalize-SearchText -Text $usable[0]) -in @('AVISO DE LICITACAO', 'ADJUDICACAO HOMOLOGACAO', 'AVISO DE RETI RATIFICACAO', 'EXTRATO')) {
        return (($usable[0] + ' - ' + $usable[1]) -replace '\s+', ' ').Trim(' -')
    }

    return (($usable[0]) -replace '\s+', ' ').Trim()
}

function Get-ActType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block
    )

    $normalized = Normalize-SearchText -Text $Block

    if ($normalized -match 'GESTOR(?: E FISCAL)? DO CONTRATO|FISCAL DO CONTRATO') { return 'Gestao Contratual' }
    if ($normalized -match 'TERMO ADITIVO') { return 'Termo Aditivo' }
    if ($normalized -match 'APOSTILAMENTO') { return 'Apostilamento' }
    if ($normalized -match 'ATA DE REGISTRO DE PRECOS') { return 'Registro de Precos' }
    if ($normalized -match 'DISPENSA DE LICITACAO') { return 'Dispensa' }
    if ($normalized -match 'INEXIGIBILIDADE') { return 'Inexigibilidade' }
    if ($normalized -match 'RESCISAO CONTRATUAL|TERMO DE RESCISAO') { return 'Rescisao' }
    if ($normalized -match 'RESULTADO DA CHAMADA PUBLICA|RESULTADO DO CREDENCIAMENTO') { return 'Homologacao' }
    if ($normalized -match 'ADJUDICACAO|HOMOLOGACAO') { return 'Homologacao' }
    if ($normalized -match 'AVISO DE LICITACAO|PREGAO ELETRONICO|CONCORRENCIA|CHAMADA PUBLICA') { return 'Aviso de Licitacao' }
    if ($normalized -match 'EXTRATO DE CONTRATO') { return 'Extrato de Contrato' }
    if ($normalized -match 'CONTRATO N') { return 'Contrato' }

    return 'Ato Contratual'
}

function Get-RecordClass {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Type
    )

    switch ($Type) {
        'Gestao Contratual' { return 'gestao_contratual' }
        'Homologacao' { return 'licitacao_ou_contratacao' }
        'Aviso de Licitacao' { return 'licitacao_ou_contratacao' }
        'Dispensa' { return 'licitacao_ou_contratacao' }
        'Inexigibilidade' { return 'licitacao_ou_contratacao' }
        default { return 'execucao_contratual' }
    }
}

function Get-RecordClassLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RecordClass
    )

    switch ($RecordClass) {
        'gestao_contratual' { return 'Gestao do contrato' }
        'licitacao_ou_contratacao' { return 'Licitacao e contratacao' }
        default { return 'Execucao contratual' }
    }
}

function Clean-EntityName {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $clean = ($Text -replace '\s+', ' ').Trim()
    $clean = $clean -replace '\s*-\s*CNPJ.*$', ''
    $clean = $clean -replace '\s+CNPJ.*$', ''
    $clean = $clean -replace '\s*,?\s*inscrit[ao].*$', ''
    $clean = $clean.Trim(" -,:;.")
    $wordCount = @($clean -split '\s+' | Where-Object { $_ }).Count

    if ($wordCount -gt 18) {
        return ''
    }

    if ((Normalize-SearchText -Text $clean) -match '^(COM BASE NA LEI|ART|INSCRITA NO|INSCRITO NO|O MUNICIPIO|A PREFEITURA|DE R|POIS QUE|AO ANALISAR O CASO|EIS QUE)') {
        return ''
    }

    return $clean
}

function Get-PrimaryCurrencyValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block
    )

    $patterns = @(
        'VALOR(?: TOTAL| GLOBAL| FIXADO[^:]*)?\s*:?\s*(?<value>R\$\s?[\d\.\,]+)',
        'IMPORTANCIA DE\s*(?<value>R\$\s?[\d\.\,]+)',
        'ACRESCIMO(?: DE SERVICOS)? NO VALOR DE\s*(?<value>R\$\s?[\d\.\,]+)',
        '(?<value>R\$\s?[\d\.\,]+)'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($Block, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($match.Success) {
            return ($match.Groups['value'].Value -replace '\s+', ' ').Trim()
        }
    }

    return ''
}

function Get-ConfidenceProfile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block,

        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [string]$RecordClass,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ContractNumber,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ProcessNumber,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Contractor,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ObjectValue,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    $normalized = Normalize-SearchText -Text $Block
    $score = 0

    if ($Type -ne 'Ato Contratual') { $score += 3 }
    if ($normalized -match 'CONTRATANTE|CONTRATADA|FORNECEDOR|DETENTORA') { $score += 2 }
    if ($normalized -match 'OBJETO') { $score += 2 }
    if ($normalized -match 'VALOR') { $score += 1 }
    if (-not [string]::IsNullOrWhiteSpace($ContractNumber)) { $score += 2 }
    if (-not [string]::IsNullOrWhiteSpace($ProcessNumber)) { $score += 1 }
    if (-not [string]::IsNullOrWhiteSpace($Contractor)) { $score += 2 }
    if (-not [string]::IsNullOrWhiteSpace($ObjectValue)) { $score += 2 }
    if (-not [string]::IsNullOrWhiteSpace($Value)) { $score += 1 }
    if ($RecordClass -eq 'gestao_contratual') { $score += 1 }

    if ($score -ge 9) {
        $label = 'alta'
    }
    elseif ($score -ge 6) {
        $label = 'media'
    }
    else {
        $label = 'baixa'
    }

    return [pscustomobject]@{
        score = $score
        label = $label
    }
}

function New-ContractItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Block,

        [Parameter(Mandatory = $true)]
        [int]$PageNumber,

        [Parameter(Mandatory = $false)]
        [string]$PageContext = ''
    )

    $type = Get-ActType -Block $Block
    $recordClass = Get-RecordClass -Type $type
    $actTitle = Get-FirstMeaningfulLine -Block $Block

    $contractNumber = Get-FirstPatternValue -Text $Block -Patterns @(
        'CONTRATO(?:\s+ADMINISTRATIVO)?\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
        'DO\s+CONTRATO\s*(?<value>\d{1,5}\/\d{4})',
        'AO\s+CONTRATO\s*(?<value>\d{1,5}\/\d{4})',
        'CONTRATO\s*(?<value>\d{1,5}\/\d{4})',
        'ATA\s+DE\s+REGISTRO\s+DE\s+PRECOS(?:\s+N\S*\s*:?)?\s*(?<value>\d{1,5}\/\d{4})'
    )

    $processNumber = Get-FirstPatternValue -Text $Block -Patterns @(
        'PROCESSO(?:\s+ADMINISTRATIVO)?\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
        'PROCESSO\s+N\S*\s*(?<value>\d{1,5}\/\d{4})',
        'EDITAL\s+CHAMADA\s+PUBLICA\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
        'CHAMADA\s+PUBLICA\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
        'PR-[A-Z]\s*-\s*PREGAO(?:\s+ELETRONICO|\s+PRESENCIAL)?\s*-\s*(?<value>\d{1,5}\/\d{4})',
        'PREGAO(?:\s+ELETRONICO|\s+PRESENCIAL)?\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
        'CP-[A-Z]\s*-\s*CONCORRENCIA(?:\s+PUBLICA)?(?:\s+ELETRONICA)?\s*-\s*(?<value>\d{1,5}\/\d{4})',
        'CONCORRENCIA(?:\s+PUBLICA)?(?:\s+ELETRONICA)?\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
        'INEXIGIBILIDADE(?:\s+DE\s+LICITACAO)?\s*-\s*N\S*\s*(?<value>\d{1,5}\/\d{4})',
        'DISPENSA(?:\s+DE\s+LICITACAO)?\s*-\s*N\S*\s*(?<value>\d{1,5}\/\d{4})'
    )

    $modality = Get-FirstPatternValue -Text $Block -Patterns @(
        'MODALIDADE\s*:?\s*(?<value>[^\n\.;]+)',
        '(?<value>PR-[A-Z]\s*-\s*PREGAO(?:\s+ELETRONICO|\s+PRESENCIAL)?\s*-\s*\d{1,5}\/\d{4}(?:\s*-\s*EDITAL\s+N\S*\s*:?\s*\d{1,5}\/\d{4})?)',
        '(?<value>CP-[A-Z]\s*-\s*CONCORRENCIA(?:\s+PUBLICA)?(?:\s+ELETRONICA)?\s*-\s*\d{1,5}\/\d{4}(?:\s*-\s*EDITAL\s+N\S*\s*:?\s*\d{1,5}\/\d{4})?)',
        '(?<value>EDITAL\s+CHAMADA\s+PUBLICA\s+N\S*\s*:?\s*\d{1,5}\/\d{4})',
        '(?<value>PREGAO\s+ELETRONICO\s+N\S*\s*:?\s*\d{1,5}\/\d{4})',
        '(?<value>PREGAO\s+PRESENCIAL\s+N\S*\s*:?\s*\d{1,5}\/\d{4})',
        '(?<value>CONCORRENCIA(?:\s+PUBLICA)?(?:\s+ELETRONICA)?\s+N\S*\s*:?\s*\d{1,5}\/\d{4})',
        '(?<value>CHAMADA\s+PUBLICA\s+N\S*\s*:?\s*\d{1,5}\/\d{4})'
    )

    $contractor = Get-FirstPatternValue -Text $Block -Patterns @(
        'CONTRATAD[AO]\s*:?\s*(?<value>[^\n]+)',
        'FORNECEDOR\s*:?\s*(?<value>[^\n]+)',
        'CREDENCIAD[AO]\s*:?\s*(?<value>[^\n]+)',
        'DETENTOR[AO]\s*:?\s*(?<value>[^\n]+)'
    )

    $cnpj = Get-FirstPatternValue -Text $Block -Patterns @(
        '(?<value>\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2})'
    )

    $objectValue = Get-FirstPatternValue -Text $Block -Patterns @(
        'OBJETO\s*:?\s*(?<value>[\s\S]{20,650}?)(?=\n(?:VALOR|VIGENCIA|PRAZO|DATA|ASSINATURA|INTERESSAD[OA]S?|CONTRATANTE|CONTRATAD[AO]|FORNECEDOR|CNPJ|CPF|FUNDAMENTO|O RECEBIMENTO|A ABERTURA|INICIO DA SESSAO|ADJUDICO|HOMOLOGO|IGUAPE,)|$)',
        'PROCESSO(?:\s+ADMINISTRATIVO)?\s+N\S*\s*:?\s*\d{1,5}\/\d{4}\s*(?<value>[\s\S]{20,650}?)(?=\n(?:PR-[A-Z]|CP-[A-Z]|PREGAO|CONCORRENCIA|EDITAL\s+CHAMADA\s+PUBLICA|INTERESSAD[OA]S?|ADJUDICO|HOMOLOGO|A PREFEITURA MUNICIPAL|IGUAPE,)|$)',
        'EDITAL\s+CHAMADA\s+PUBLICA\s+N\S*\s*:?\s*\d{1,5}\/\d{4}\s*(?<value>[\s\S]{20,650}?)(?=\n(?:INTERESSAD[OA]S?|ADJUDICO|HOMOLOGO|IGUAPE,)|$)',
        '(?:DO\s+)?CONTRATO\s+\d{1,5}\/\d{4},\s*(?<value>[\s\S]{20,420}?)(?=\.\s*Art\.?\s*2|\n\s*Art\.?\s*2|\n\s*Esta\s+Portaria|$)',
        'O OBJETO DESTE CONTRATO E\s*(?<value>[\s\S]{20,320}?)(?=\n(?:ART\.|ART |ESTA PORTARIA|GABINETE|$))'
    )

    $term = Get-FirstPatternValue -Text $Block -Patterns @(
        'VIGENCIA\s*:?\s*(?<value>[^\n]+)',
        'PRAZO(?: DE VIGENCIA)?\s*:?\s*(?<value>[^\n]+)',
        'VALIDADE POR\s*(?<value>[^\n]+)'
    )

    $signatureDate = Get-FirstPatternValue -Text $Block -Patterns @(
        'DATA(?: DA ASSINATURA)?\s*:?\s*(?<value>[^\n\.;]+)',
        'ASSINATURA\s*:?\s*(?<value>[^\n\.;]+)'
    )

    $legalBasis = Get-FirstPatternValue -Text $Block -Patterns @(
        'FUNDAMENTO LEGAL\s*:?\s*(?<value>[^\n]+)',
        '(?<value>Lei Federal\s+n\S*\s*14\.133\/2021[^\n\.;]*)',
        '(?<value>Lei\s+n\S*\s*8\.666\/93[^\n\.;]*)'
    )

    $value = Get-PrimaryCurrencyValue -Block $Block
    $contractor = Clean-EntityName -Text $contractor
    $objectValue = Clean-ObjectText -Text $objectValue
    $modality = (($modality -replace '\s+', ' ').Trim())

    if ([string]::IsNullOrWhiteSpace($modality) -and $type -eq 'Aviso de Licitacao') {
        $modality = $actTitle
    }

    if ($processNumber -eq 'OBJETO') {
        $processNumber = ''
    }

    if ($recordClass -eq 'gestao_contratual' -and $processNumber -match '^\d{1,3}$') {
        $processNumber = ''
    }

    if ([string]::IsNullOrWhiteSpace($processNumber) -and $recordClass -ne 'gestao_contratual') {
        $processNumber = Get-FirstPatternValue -Text $actTitle -Patterns @(
            'PROCESSO\s+N\S*\s*:?\s*(?<value>\d{1,5}\/\d{4})',
            '(?:PREGAO|CONCORRENCIA|CHAMADA\s+PUBLICA|INEXIGIBILIDADE|DISPENSA)[^\n]*?(?<value>\d{1,5}\/\d{4})'
        )
    }

    $blockOrganizationMatches = @(Find-OrganizationsInText -Text $Block)
    $pageOrganizationMatches = @(Find-OrganizationsInText -Text $PageContext)
    $organizationMatches = @(Merge-OrganizationMatches -Matches (@($blockOrganizationMatches) + @($pageOrganizationMatches)))
    $primaryOrganization = Get-PreferredPrimaryOrganization -BlockMatches $blockOrganizationMatches -PageMatches $pageOrganizationMatches

    if ($recordClass -eq 'gestao_contratual' -and [string]::IsNullOrWhiteSpace($contractNumber) -and [string]::IsNullOrWhiteSpace($objectValue)) {
        return $null
    }

    $actTitle = Refine-ActTitle -Title $actTitle -Type $type -Block $Block -ContractNumber $contractNumber -ProcessNumber $processNumber

    $flags = New-Object System.Collections.Generic.List[string]
    switch ($recordClass) {
        'execucao_contratual' {
            if ([string]::IsNullOrWhiteSpace($contractNumber)) { $flags.Add('Sem numero do contrato') }
            if ([string]::IsNullOrWhiteSpace($contractor)) { $flags.Add('Sem contratada identificada') }
            if ([string]::IsNullOrWhiteSpace($objectValue)) { $flags.Add('Sem objeto identificado') }
            if ([string]::IsNullOrWhiteSpace($value)) { $flags.Add('Sem valor identificado') }
        }
        'licitacao_ou_contratacao' {
            if ([string]::IsNullOrWhiteSpace($objectValue)) { $flags.Add('Sem objeto identificado') }
            if ([string]::IsNullOrWhiteSpace($modality)) { $flags.Add('Sem modalidade identificada') }
            if ([string]::IsNullOrWhiteSpace($processNumber) -and [string]::IsNullOrWhiteSpace($contractNumber)) { $flags.Add('Sem numero principal identificado') }
        }
        'gestao_contratual' {
            if ([string]::IsNullOrWhiteSpace($contractNumber)) { $flags.Add('Sem numero do contrato') }
            if ([string]::IsNullOrWhiteSpace($objectValue)) { $flags.Add('Sem objeto identificado') }
            if ($null -eq $primaryOrganization) { $flags.Add('Sem orgao responsavel identificado') }
        }
    }

    $filledFields = 0
    foreach ($field in @($contractNumber, $processNumber, $contractor, $objectValue, $value, $modality, $signatureDate, $term)) {
        if (-not [string]::IsNullOrWhiteSpace($field)) {
            $filledFields++
        }
    }

    if ($filledFields -ge 6) {
        $completeness = 'alta'
    }
    elseif ($filledFields -ge 3) {
        $completeness = 'media'
    }
    else {
        $completeness = 'baixa'
    }

    $confidence = Get-ConfidenceProfile -Block $Block -Type $type -RecordClass $recordClass -ContractNumber $contractNumber -ProcessNumber $processNumber -Contractor $contractor -ObjectValue $objectValue -Value $value

    return [pscustomobject]@{
        type = $type
        actTitle = $actTitle
        recordClass = $recordClass
        recordClassLabel = (Get-RecordClassLabel -RecordClass $recordClass)
        confidenceScore = [int]$confidence.score
        confidenceLabel = [string]$confidence.label
        pageNumber = $PageNumber
        contractNumber = $contractNumber
        processNumber = $processNumber
        modality = $modality
        contractor = $contractor
        cnpj = $cnpj
        object = $objectValue
        value = $value
        term = $term
        signatureDate = $signatureDate
        legalBasis = $legalBasis
        primaryOrganizationId = if ($primaryOrganization) { [string]$primaryOrganization.id } else { $null }
        primaryOrganizationName = if ($primaryOrganization) { [string]$primaryOrganization.name } else { $null }
        primaryOrganizationSphere = if ($primaryOrganization) { [string]$primaryOrganization.sphere } else { $null }
        primaryOrganizationAreaId = if ($primaryOrganization) { [string]$primaryOrganization.areaId } else { $null }
        mentionedOrganizationIds = @($organizationMatches | ForEach-Object { [string]$_.id })
        mentionedOrganizationNames = @($organizationMatches | ForEach-Object { [string]$_.name })
        excerpt = $Block.Substring(0, [Math]::Min(520, $Block.Length))
        completeness = $completeness
        flags = $flags.ToArray()
    }
}

function Get-UniqueContractItems {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items
    )

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $deduped = New-Object System.Collections.Generic.List[object]

    foreach ($item in $Items) {
        $objectText = [string]$item.object
        $objectKey = $objectText.Substring(0, [Math]::Min(120, $objectText.Length))
        $primaryNumber = if (-not [string]::IsNullOrWhiteSpace([string]$item.contractNumber)) {
            [string]$item.contractNumber
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$item.processNumber)) {
            [string]$item.processNumber
        }
        else {
            ''
        }

        $key = '{0}|{1}|{2}|{3}' -f $item.recordClass, $item.type, $primaryNumber, $objectKey

        if ($seen.Add($key)) {
            $deduped.Add($item)
        }
    }

    return $deduped.ToArray()
}

function Invoke-ServerSideAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Diaries
    )

    $pdfToTextTool = Get-PdfToTextToolPath
    if (-not $pdfToTextTool) {
        Update-Status -Updates @{
            message = 'Ferramenta pdftotext nao encontrada; analise textual foi ignorada.'
        }
        return
    }

    $candidates = @($Diaries | Where-Object { @($_.candidateKeywords).Count -gt 0 -and $_.localPdfPath -and (Test-Path -LiteralPath $_.localPdfPath) })
    $totalCandidates = @($candidates).Count
    $currentIndex = 0

    foreach ($diary in $candidates) {
        $currentIndex++
        $analysis = Get-ExistingAnalysis -DiaryId ([string]$diary.id)
        if (Test-AnalysisCurrent -Diary $diary -Analysis $analysis) {
            continue
        }

        Update-Status -Updates @{
            syncStage = 'analisando'
            message = "Analisando contratos: edicao $($diary.edition) ($currentIndex de $totalCandidates)."
        }

        try {
            $pageTexts = @(Convert-PdfToPageTexts -PdfPath $diary.localPdfPath -PdfToTextTool $pdfToTextTool)
            $items = New-Object System.Collections.Generic.List[object]

            foreach ($page in $pageTexts) {
                $pageContext = Get-PageContextText -PageText $page.text
                $blocks = Get-TextBlocksFromPage -PageText $page.text
                foreach ($block in $blocks) {
                    if (-not (Test-RelevantContractBlock -Block $block)) {
                        continue
                    }

                    $candidateItem = New-ContractItem -Block $block -PageNumber $page.pageNumber -PageContext $pageContext
                    if ($null -eq $candidateItem) {
                        continue
                    }
                    if ([string]$candidateItem.confidenceLabel -eq 'baixa' -and [string]$candidateItem.recordClass -ne 'gestao_contratual') {
                        continue
                    }

                    $items.Add($candidateItem)
                }
            }

            $dedupedItems = @(Get-UniqueContractItems -Items $items.ToArray())
            $totalValue = 0.0
            foreach ($item in $dedupedItems) {
                $totalValue += (Convert-BrazilianCurrencyToNumber -Text ([string]$item.value))
            }

            $analysisPayload = [ordered]@{
                diaryId = [string]$diary.id
                parserVersion = $script:ParserVersion
                analyzedAt = Get-IsoNow
                sourcePdfUrl = [string]$diary.pdfUrl
                sourceLocalPdfRelative = [string]$diary.localPdfRelative
                keywords = @($diary.candidateKeywords)
                summary = [ordered]@{
                    pageCount = @($pageTexts).Count
                    itemCount = @($dedupedItems).Count
                    totalValue = (Format-BrazilianCurrency -Value $totalValue)
                }
                items = @($dedupedItems)
            }
        }
        catch {
            $analysisPayload = [ordered]@{
                diaryId = [string]$diary.id
                parserVersion = $script:ParserVersion
                analyzedAt = Get-IsoNow
                sourcePdfUrl = [string]$diary.pdfUrl
                sourceLocalPdfRelative = [string]$diary.localPdfRelative
                keywords = @($diary.candidateKeywords)
                summary = [ordered]@{
                    pageCount = 0
                    itemCount = 0
                    totalValue = (Format-BrazilianCurrency -Value 0)
                    error = $_.Exception.Message
                }
                items = @()
            }
        }

        Write-JsonFile -Path (Get-AnalysisPath -DiaryId ([string]$diary.id)) -Data $analysisPayload
    }

    $analysisCount = @(Get-ChildItem -LiteralPath $script:AnalysisRoot -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    Update-Status -Updates @{
        analyzedDiaries = $analysisCount
        pendingAnalysis = [Math]::Max($totalCandidates - $analysisCount, 0)
    }
}

function Get-PersonnelAnalysisPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DiaryId
    )

    return Join-Path $script:PersonnelAnalysisRoot "$DiaryId.json"
}

function Get-ExistingPersonnelAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DiaryId
    )

    return Read-JsonFile -Path (Get-PersonnelAnalysisPath -DiaryId $DiaryId) -Default $null
}

function Test-PersonnelAnalysisCurrent {
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

    if ([string]$Analysis.parserVersion -ne $script:PersonnelParserVersion) {
        return $false
    }

    if ([string]$Analysis.sourcePdfUrl -ne [string]$Diary.pdfUrl) {
        return $false
    }

    return $true
}

function Get-PersonnelEventsFromPageText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageText,

        [Parameter(Mandatory = $true)]
        [int]$PageNumber
    )

    $compact = ($PageText -replace "[\r\n]+", ' ' -replace '\s{2,}', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($compact)) {
        return @()
    }

    $normalized = Normalize-SearchText -Text $compact
    if ($normalized -notmatch '\bEXONERAD[OA]\b') {
        return @()
    }

    $events = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $matches = [regex]::Matches(
        $compact,
        '(?is)Fica\s+exonerad[oa](?:,\s*|\s+)(?:a\s+partir\s+de\s+[^,\.]+,\s*)?(?:o|a)\s+(?:servidor(?:a)?\s+public[ao]\s+)?(?<name>[\p{L}''\-]+(?:\s+[\p{L}''\-]+){1,7})(?=\s*(?:\(|,\s*ocupante|\s+ocupante|\s+do\s+cargo|,?\s*matr[iÃ­]cula|\.))',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    foreach ($match in $matches) {
        $personName = Clean-PersonDisplayName -Text ([string]$match.Groups['name'].Value)
        if ([string]::IsNullOrWhiteSpace($personName)) {
            continue
        }

        $normalizedName = Normalize-IndexText -Text $personName
        if ([string]::IsNullOrWhiteSpace($normalizedName) -or -not $seen.Add($normalizedName)) {
            continue
        }

        $startIndex = [Math]::Max($match.Index - 80, 0)
        $excerptLength = [Math]::Min(280, $compact.Length - $startIndex)
        $excerpt = $compact.Substring($startIndex, $excerptLength).Trim()

        $events.Add([pscustomobject]@{
            type = 'exoneracao'
            personName = $personName
            normalizedName = $normalizedName
            pageNumber = $PageNumber
            excerpt = $excerpt
        })
    }

    return @($events.ToArray())
}

function Invoke-PersonnelEventAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Diaries
    )

    $pdfToTextTool = Get-PdfToTextToolPath
    if (-not $pdfToTextTool) {
        Update-Status -Updates @{
            message = 'Ferramenta pdftotext nao encontrada; a leitura de exoneracoes nao foi executada.'
        }
        return
    }

    $candidates = @($Diaries | Where-Object { $_.localPdfPath -and (Test-Path -LiteralPath $_.localPdfPath) })
    $totalCandidates = @($candidates).Count
    $currentIndex = 0

    foreach ($diary in $candidates) {
        $currentIndex++
        $analysis = Get-ExistingPersonnelAnalysis -DiaryId ([string]$diary.id)
        if (Test-PersonnelAnalysisCurrent -Diary $diary -Analysis $analysis) {
            continue
        }

        Update-Status -Updates @{
            syncStage = 'analisando-pessoal'
            message = "Verificando exonerações e trocas de pessoal ($currentIndex de $totalCandidates)."
        }

        try {
            $pageTexts = @(Convert-PdfToPageTexts -PdfPath $diary.localPdfPath -PdfToTextTool $pdfToTextTool)
            $events = New-Object System.Collections.Generic.List[object]

            foreach ($page in $pageTexts) {
                foreach ($event in @(Get-PersonnelEventsFromPageText -PageText $page.text -PageNumber ([int]$page.pageNumber))) {
                    $events.Add([pscustomobject]@{
                        type = [string]$event.type
                        personName = [string]$event.personName
                        normalizedName = [string]$event.normalizedName
                        pageNumber = [int]$event.pageNumber
                        excerpt = [string]$event.excerpt
                        publishedAt = [string]$diary.publishedAt
                        diaryId = [string]$diary.id
                        edition = [string]$diary.edition
                    })
                }
            }

            $analysisPayload = [ordered]@{
                diaryId = [string]$diary.id
                parserVersion = $script:PersonnelParserVersion
                analyzedAt = Get-IsoNow
                sourcePdfUrl = [string]$diary.pdfUrl
                sourceLocalPdfRelative = [string]$diary.localPdfRelative
                eventCount = @($events).Count
                events = @($events)
            }
        }
        catch {
            $analysisPayload = [ordered]@{
                diaryId = [string]$diary.id
                parserVersion = $script:PersonnelParserVersion
                analyzedAt = Get-IsoNow
                sourcePdfUrl = [string]$diary.pdfUrl
                sourceLocalPdfRelative = [string]$diary.localPdfRelative
                eventCount = 0
                error = $_.Exception.Message
                events = @()
            }
        }

        Write-JsonFile -Path (Get-PersonnelAnalysisPath -DiaryId ([string]$diary.id)) -Data $analysisPayload
    }
}

function Convert-HtmlFragmentToText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ''
    }

    $normalized = [regex]::Replace($Html, '<br\s*/?>', "`n", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $normalized = [regex]::Replace($normalized, '<[^>]+>', ' ')
    $normalized = HtmlDecode-Safe -Text $normalized
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $normalized = $normalized -replace "[\r\t]+", ' '
    $normalized = $normalized -replace '\s{2,}', ' '
    return ($normalized -replace ' ?\n ?', "`n").Trim()
}

function Get-PortalContractListPagePaths {
    $firstPageHtml = Invoke-PortalRequest -PathOrUrl '/portal/contratos'
    $templatePath = Get-FirstRegexValue -Text $firstPageHtml -Pattern '(?<value>/portal/contratos/\d+(?:/0)+)'
    $firstPageIds = [regex]::Matches($firstPageHtml, '/portal/contrato/\d+') | ForEach-Object { $_.Value } | Select-Object -Unique
    $pageNumbers = @(
        [regex]::Matches($firstPageHtml, '/portal/contratos/(?<page>\d+)(?:/0)+') |
        ForEach-Object { Parse-IntegerLike -Text $_.Groups['page'].Value } |
        Where-Object { $_ -gt 0 } |
        Select-Object -Unique
    )
    $pageSize = [Math]::Max(@($firstPageIds).Count, 1)
    $parsedTotalResults = Parse-IntegerLike -Text (Get-FirstRegexValue -Text $firstPageHtml -Pattern '(?<value>\d+)\s+registros encontrados')
    $totalResults = [Math]::Max($parsedTotalResults, @($firstPageIds).Count)
    $totalPages = if (@($pageNumbers).Count -gt 0) {
        [int]((@($pageNumbers) | Measure-Object -Maximum).Maximum)
    }
    else {
        [Math]::Max([int][Math]::Ceiling($totalResults / [double]$pageSize), 1)
    }
    $paths = New-Object System.Collections.Generic.List[string]
    $paths.Add('/portal/contratos')

    if ([string]::IsNullOrWhiteSpace($templatePath)) {
        $templatePath = '/portal/contratos/2/0/0/0/0/0/0/0/0/0/0/0/0/0'
    }

    for ($page = 2; $page -le $totalPages; $page++) {
        $paths.Add(([regex]::Replace($templatePath, '(?<=/portal/contratos/)\d+', [string]$page)))
    }

    return [pscustomobject]@{
        pagePaths = @($paths | Select-Object -Unique)
        totalPages = $totalPages
        totalResults = $totalResults
        firstPageHtml = $firstPageHtml
    }
}

function Get-PortalContractSummaryFromCard {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ContractId,

        [Parameter(Mandatory = $true)]
        [string]$CardHtml
    )

    $cardText = Convert-HtmlFragmentToText -Html $CardHtml
    $objectText = Get-FirstRegexValue -Text $cardText -Pattern "^(?<value>.+?)(?=\s+(?:Contratada\(s\):|N(?:u|ú)mero:|N(?:º|°|o)\s*processo:|Vig(?:e|ê)ncia:|Assinado em:|Origem:|Tipo:)|$)"
    $contractor = Get-FirstRegexValue -Text $cardText -Pattern "Contratada\(s\):\s*(?<value>.+?)(?=\s+(?:N(?:u|ú)mero:|N(?:º|°|o)\s*processo:|Vig(?:e|ê)ncia:|Assinado em:|Origem:|Tipo:)|$)"

    return [pscustomobject]@{
        portalContractId = [string]$ContractId
        viewPath = "/portal/contrato/$ContractId"
        viewUrl = (Get-AbsolutePortalUrl -PathOrUrl "/portal/contrato/$ContractId")
        listObject = (Clean-ObjectText -Text ([string]$objectText))
        listContractor = (Clean-EntityName -Text ([string]$contractor))
        contractNumber = (Get-FirstRegexValue -Text $cardText -Pattern "N(?:u|ú)mero:\s*(?<value>[^\s]+)")
        processNumber = (Get-FirstRegexValue -Text $cardText -Pattern "N(?:º|°|o)\s*processo:\s*(?<value>[^\s]+)")
        listTerm = (Get-FirstRegexValue -Text $cardText -Pattern "Vig(?:e|ê)ncia:\s*(?<value>[^\n]+?)(?=\s+(?:Assinado em:|Origem:|Tipo:)|$)")
        signatureDate = (Get-FirstRegexValue -Text $cardText -Pattern 'Assinado em:\s*(?<value>[^\s]+)')
        portalOrigin = (Get-FirstRegexValue -Text $cardText -Pattern 'Origem:\s*(?<value>.+?)(?=\s+Tipo:|$)')
        type = (Get-FirstRegexValue -Text $cardText -Pattern 'Tipo:\s*(?<value>.+?)(?=\s+\d+\s+aditivos?|\s+\d+\s+anexos?|$)')
        listExcerpt = $cardText
    }
}

function Get-PortalContractSummaries {
    $pageInfo = Get-PortalContractListPagePaths
    $summaries = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'

    for ($index = 0; $index -lt @($pageInfo.pagePaths).Count; $index++) {
        $pagePath = [string]$pageInfo.pagePaths[$index]
        $html = if ($index -eq 0) { $pageInfo.firstPageHtml } else { Invoke-PortalRequest -PathOrUrl $pagePath }

        Update-Status -Updates @{
            syncStage = 'sincronizando-contratos'
            message = "Lendo a listagem oficial de contratos ($($index + 1) de $(@($pageInfo.pagePaths).Count))."
        }

        $matches = [regex]::Matches(
            $html,
            '<a[^>]+href="/portal/contrato/(?<id>\d+)"[^>]*>(?<body>[\s\S]*?)</a>',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )

        foreach ($match in $matches) {
            $contractId = [string]$match.Groups['id'].Value
            if ([string]::IsNullOrWhiteSpace($contractId) -or -not $seen.Add($contractId)) {
                continue
            }

            $summaries.Add((Get-PortalContractSummaryFromCard -ContractId $contractId -CardHtml ([string]$match.Groups['body'].Value)))
        }
    }

    $summaryItems = @($summaries.ToArray())
    return [ordered]@{
        totalResults = [int]$pageInfo.totalResults
        items = $summaryItems
    }
}

function Get-PortalContractDetailField {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $pattern = '(?s)<div class="sw_titulo_detalhe[^"]*">\s*' + [regex]::Escape($Label) + '\s*</div>\s*<div class="sw_descricao_detalhe[^"]*">(?<value>.*?)</div>'
    return Convert-HtmlFragmentToText -Html (Get-FirstRegexValue -Text $Html -Pattern $pattern)
}

function Save-PortalContractDocument {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$DownloadTokenPath,

        [Parameter(Mandatory = $true)]
        [string]$PortalYear,

        [Parameter(Mandatory = $true)]
        [string]$PortalContractId,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ExistingItem
    )

    if ([string]::IsNullOrWhiteSpace($DownloadTokenPath)) {
        return [pscustomobject]@{
            localPdfRelative = $null
            webPdfPath = $null
            pdfUrl = $null
            downloaded = $false
        }
    }

    $pdfUrl = Get-PdfRedirectUrl -DownloadTokenPath $DownloadTokenPath
    $fileName = [System.IO.Path]::GetFileName(([Uri]$pdfUrl).LocalPath)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "contrato-portal-$PortalContractId.pdf"
    }

    $localInfo = Get-LocalPortalPdfInfo -PortalYear $PortalYear -FileName $fileName
    Ensure-Directory -Path (Split-Path -Parent $localInfo.absolutePath)
    $shouldDownload = -not (Test-Path -LiteralPath $localInfo.absolutePath)

    if (-not $shouldDownload -and $ExistingItem) {
        $existingDownloadPath = [string]$ExistingItem.downloadTokenPath
        if ($existingDownloadPath -ne $DownloadTokenPath) {
            $shouldDownload = $true
        }
    }

    if (Test-PublicWorkflowMode) {
        $shouldDownload = $false
    }

    if ($shouldDownload) {
        Invoke-WebRequest -UseBasicParsing -Uri $pdfUrl -Headers @{ 'User-Agent' = $script:UserAgent } -TimeoutSec 180 -OutFile $localInfo.absolutePath | Out-Null
    }

    return [pscustomobject]@{
        localPdfRelative = $localInfo.relativePath
        webPdfPath = $localInfo.webPath
        pdfUrl = $pdfUrl
        downloaded = $shouldDownload
    }
}

function Get-PortalContractItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ExistingItem
    )

    $detailHtml = Invoke-PortalRequest -PathOrUrl ([string]$Summary.viewPath)
    $updatedAt = Convert-PortalDateTime -Text (Get-FirstRegexValue -Text $detailHtml -Pattern 'Atualizado em:\s*(?<value>[^<]+)')
    $detailNumber = Convert-HtmlFragmentToText -Html (Get-FirstRegexValue -Text $detailHtml -Pattern '(?s)<div class="cnt_titulo_contrato[^"]*">(?<value>.*?)</div>')
    $detailNumber = ($detailNumber -replace "^\s*N(?:º|°|o|\.)*\s*", "").Trim()
    $status = Convert-HtmlFragmentToText -Html (Get-FirstRegexValue -Text $detailHtml -Pattern '(?s)<!--\s*Situa.*?<span>(?<value>[^<]+)</span>')
    $contractor = Convert-HtmlFragmentToText -Html (Get-FirstRegexValue -Text $detailHtml -Pattern '(?s)Contratada\(s\):</strong>\s*(?<value>.*?)</div>')
    $objectText = Clean-ObjectText -Text (Convert-HtmlFragmentToText -Html (Get-FirstRegexValue -Text $detailHtml -Pattern '(?s)<div class="cnt_titulo">Objeto</div>\s*<div class="cnt_descricao">(?<value>.*?)</div>'))
    $vigencia = Get-PortalContractDetailField -Html $detailHtml -Label "Vigência"
    $signatureDate = Get-PortalContractDetailField -Html $detailHtml -Label 'Data da Assinatura'
    $origin = Get-PortalContractDetailField -Html $detailHtml -Label 'Origem'
    $type = Get-PortalContractDetailField -Html $detailHtml -Label 'Tipo'
    $revenueExpense = Get-PortalContractDetailField -Html $detailHtml -Label 'Receita ou Despesa'
    $value = Get-PortalContractDetailField -Html $detailHtml -Label 'Valor'
    $downloadTokenPath = Get-FirstRegexValue -Text $detailHtml -Pattern '(?<value>/portal/download/contratos/[^"]+)'
    $aditiveCount = Parse-IntegerLike -Text (Get-FirstRegexValue -Text $detailHtml -Pattern 'cnt_info_total[^>]*>\s*Total:\s*(?<value>\d+)')
    $contractNumber = if (-not [string]::IsNullOrWhiteSpace($detailNumber)) { $detailNumber } else { [string]$Summary.contractNumber }
    $contractor = if (-not [string]::IsNullOrWhiteSpace($contractor)) { Clean-EntityName -Text $contractor } else { [string]$Summary.listContractor }
    $objectText = if (-not [string]::IsNullOrWhiteSpace($objectText)) { $objectText } else { [string]$Summary.listObject }
    $type = if (-not [string]::IsNullOrWhiteSpace($type)) { $type } else { [string]$Summary.type }
    $origin = if (-not [string]::IsNullOrWhiteSpace($origin)) { $origin } else { [string]$Summary.portalOrigin }
    $vigencia = if (-not [string]::IsNullOrWhiteSpace($vigencia)) { $vigencia } else { [string]$Summary.listTerm }
    $signatureDate = if (-not [string]::IsNullOrWhiteSpace($signatureDate)) { $signatureDate } else { [string]$Summary.signatureDate }
    $publishedAt = Convert-PortalDateTime -Text $signatureDate
    if ([string]::IsNullOrWhiteSpace($publishedAt)) {
        $publishedAt = $updatedAt
    }

    $contextText = @(
        $contractNumber,
        $contractor,
        $objectText,
        $type,
        $origin,
        $revenueExpense,
        [string]$Summary.processNumber
    ) -join "`n"

    $organizationMatches = Find-OrganizationsInText -Text $contextText
    $preferredOrganization = Get-PreferredPrimaryOrganization -BlockMatches $organizationMatches -PageMatches @()
    if ($null -eq $preferredOrganization) {
        $preferredOrganization = [pscustomobject]@{
            id = ''
            name = ''
            sphere = ''
            areaId = ''
        }
    }
    $mentionedIds = @($organizationMatches | ForEach-Object { [string]$_.id } | Select-Object -Unique)
    $mentionedNames = @($organizationMatches | ForEach-Object { [string]$_.name } | Select-Object -Unique)
    $flags = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($contractNumber)) { $flags.Add('Sem numero do contrato') }
    if ([string]::IsNullOrWhiteSpace($contractor)) { $flags.Add('Sem contratada identificada') }
    if ([string]::IsNullOrWhiteSpace($objectText)) { $flags.Add('Sem objeto identificado') }
    if ([string]::IsNullOrWhiteSpace($value)) { $flags.Add('Sem valor identificado') }
    if ([string]::IsNullOrWhiteSpace($preferredOrganization.name)) { $flags.Add('Sem orgao principal identificado') }

    $completenessScore = 0
    foreach ($field in @($contractNumber, $contractor, $objectText, $vigencia, $signatureDate, $type, $origin)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$field)) {
            $completenessScore++
        }
    }

    $completeness = if ($completenessScore -ge 6) {
        'alta'
    }
    elseif ($completenessScore -ge 4) {
        'media'
    }
    else {
        'baixa'
    }

    $portalYear = if ($publishedAt -and $publishedAt.Length -ge 4) { $publishedAt.Substring(0, 4) } elseif ($updatedAt -and $updatedAt.Length -ge 4) { $updatedAt.Substring(0, 4) } else { 'sem-ano' }
    $documentInfo = Save-PortalContractDocument -DownloadTokenPath $downloadTokenPath -PortalYear $portalYear -PortalContractId ([string]$Summary.portalContractId) -ExistingItem $ExistingItem

    return [pscustomobject]@{
        sourceType = 'portal_contratos'
        sourceLabel = 'Portal de Contratos da Prefeitura'
        portalContractId = [string]$Summary.portalContractId
        edition = 'Portal'
        publishedAt = $publishedAt
        updatedAt = $updatedAt
        localPdfRelative = [string]$documentInfo.localPdfRelative
        webPdfPath = [string]$documentInfo.webPdfPath
        pdfUrl = [string]$documentInfo.pdfUrl
        viewUrl = [string]$Summary.viewUrl
        downloadTokenPath = [string]$downloadTokenPath
        candidateKeywords = @('portal-contratos', 'site-oficial')
        type = if (-not [string]::IsNullOrWhiteSpace($type)) { $type } else { 'Contrato' }
        contractNumber = $contractNumber
        processNumber = [string]$Summary.processNumber
        modality = $origin
        contractor = $contractor
        cnpj = (Get-FirstRegexValue -Text ($contextText + "`n" + [string]$Summary.listExcerpt) -Pattern '(?<value>\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2})')
        object = $objectText
        value = $value
        valueNumber = (Convert-BrazilianCurrencyToNumber -Text $value)
        actTitle = if (-not [string]::IsNullOrWhiteSpace($contractNumber)) { "Contrato $contractNumber" } else { 'Contrato oficial' }
        recordClass = 'execucao_contratual'
        recordClassLabel = 'Contrato oficial'
        confidenceLabel = 'alta'
        confidenceScore = 18
        term = $vigencia
        signatureDate = $signatureDate
        legalBasis = ''
        primaryOrganizationId = [string]$preferredOrganization.id
        primaryOrganizationName = [string]$preferredOrganization.name
        primaryOrganizationSphere = [string]$preferredOrganization.sphere
        primaryOrganizationAreaId = [string]$preferredOrganization.areaId
        mentionedOrganizationIds = @($mentionedIds)
        mentionedOrganizationNames = @($mentionedNames)
        excerpt = if (-not [string]::IsNullOrWhiteSpace($objectText)) { $objectText.Substring(0, [Math]::Min(460, $objectText.Length)) } else { [string]$Summary.listExcerpt }
        completeness = $completeness
        flags = @($flags)
        portalStatus = $status
        portalOrigin = $origin
        revenueExpense = $revenueExpense
        aditiveCount = $aditiveCount
        downloadedDocument = [bool]$documentInfo.downloaded
    }
}

function Sync-PortalContracts {
    $existingPayload = Get-PortalContractsPayload
    $existingById = @{}
    foreach ($item in @($existingPayload.items)) {
        $existingById[[string]$item.portalContractId] = $item
    }

    $summaries = Get-PortalContractSummaries
    $items = New-Object System.Collections.Generic.List[object]
    $downloadedCount = 0
    $newCount = 0
    $updatedCount = 0
    $total = @($summaries.items).Count

    for ($index = 0; $index -lt $total; $index++) {
        $summary = $summaries.items[$index]
        $existingItem = if ($existingById.ContainsKey([string]$summary.portalContractId)) { $existingById[[string]$summary.portalContractId] } else { $null }

        Update-Status -Updates @{
            syncStage = 'sincronizando-contratos'
            message = "Atualizando contratos oficiais ($($index + 1) de $total)."
        }

        $item = Get-PortalContractItem -Summary $summary -ExistingItem $existingItem
        if ([bool]$item.downloadedDocument) {
            $downloadedCount++
        }

        if ($null -eq $existingItem) {
            $newCount++
        }
        elseif (
            [string]$existingItem.updatedAt -ne [string]$item.updatedAt -or
            [string]$existingItem.pdfUrl -ne [string]$item.pdfUrl -or
            [string]$existingItem.localPdfRelative -ne [string]$item.localPdfRelative
        ) {
            $updatedCount++
        }

        $items.Add($item)
    }

    $payload = Get-EmptyPortalContractsPayload
    $payload.generatedAt = Get-IsoNow
    $payload.totalItems = $items.Count
    $payload.downloadedDocumentCount = $downloadedCount
    $payload.items = @(
        $items |
        Sort-Object -Property `
            @{ Expression = { $_.publishedAt }; Descending = $true }, `
            @{ Expression = { $_.contractNumber }; Descending = $false }
    )
    Save-PortalContractsPayload -Payload $payload

    return [pscustomobject]@{
        totalItems = $payload.totalItems
        downloadedCount = $downloadedCount
        newCount = $newCount
        updatedCount = $updatedCount
    }
}

Initialize-AppStorage

$existingLock = Get-SyncLock
if ($existingLock) {
    Update-Status -Updates @{
        isSyncRunning = $true
        syncStage = 'aguardando'
        message = 'Ja existe uma sincronizacao em andamento.'
    }
    exit 0
}

Set-SyncLock

try {
    Update-Status -Updates @{
        isSyncRunning = $true
        syncStartedAt = (Get-IsoNow)
        syncFinishedAt = $null
        syncStage = 'iniciando'
        message = 'Iniciando varredura do Diario Oficial e do portal de contratos de Iguape.'
        scannedPages = 0
        totalPages = 0
        newDiaries = 0
        updatedDiaries = 0
        downloadedPdfCount = 0
        lastError = $null
    }

    $diaries = Get-AllDiariesFromPortal
    $candidateMap = Get-CandidateDiaryMap -Keywords (Get-ContractKeywords)

    Update-Status -Updates @{
        syncStage = 'baixando-pdfs'
        message = 'Baixando e atualizando os PDFs locais.'
        candidateDiaries = $candidateMap.Count
    }

    $mergeResult = Merge-Diaries -CollectedDiaries $diaries -CandidateMap $candidateMap
    $payload = Get-EmptyDiariesPayload
    $payload.generatedAt = Get-IsoNow
    $payload.diaries = @($mergeResult.diaries)
    Save-DiariesPayload -Payload $payload

    Invoke-ServerSideAnalysis -Diaries $mergeResult.diaries
    Invoke-PersonnelEventAnalysis -Diaries $mergeResult.diaries
    $portalContractsSync = Sync-PortalContracts
    Update-Status -Updates @{
        syncStage = 'recompondo-base'
        message = 'Recompondo a base consolidada e o diff entre sincronizacoes.'
    }
    Refresh-ContractsAggregate | Out-Null

    $finishedAt = Get-IsoNow
    Update-Status -Updates @{
        isSyncRunning = $false
        syncStage = 'concluido'
        message = 'Sincronizacao do Diario Oficial e do portal de contratos concluida.'
        syncFinishedAt = $finishedAt
        lastSuccessfulSyncAt = $finishedAt
        totalDiaries = @($mergeResult.diaries).Count
        candidateDiaries = $candidateMap.Count
        newDiaries = $mergeResult.newCount
        updatedDiaries = $mergeResult.updatedCount
        downloadedPdfCount = ($mergeResult.downloadedCount + $portalContractsSync.downloadedCount)
    }
    Register-SyncHistoryEntry -Outcome 'success' -Message 'Sincronizacao concluida com sucesso.' -Metrics ([ordered]@{
        totalDiaries = @($mergeResult.diaries).Count
        candidateDiaries = $candidateMap.Count
        newDiaries = $mergeResult.newCount
        updatedDiaries = $mergeResult.updatedCount
        downloadedPdfCount = ($mergeResult.downloadedCount + $portalContractsSync.downloadedCount)
    }) | Out-Null
}
catch {
    Update-Status -Updates @{
        isSyncRunning = $false
        syncStage = 'erro'
        message = 'Falha durante a sincronizacao.'
        syncFinishedAt = (Get-IsoNow)
        lastError = $_.Exception.Message
    }
    Register-SyncHistoryEntry -Outcome 'error' -Message $_.Exception.Message -Metrics ([ordered]@{
        stage = 'erro'
    }) | Out-Null
    throw
}
finally {
    Clear-SyncLock
}
