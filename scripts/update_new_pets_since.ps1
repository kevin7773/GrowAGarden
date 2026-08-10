param(
    [string]$CutoffDate = '2026-05-25',
    [string]$SourceJson = 'pets_full.json',
    [string]$WikiApi = 'https://growagarden.fandom.com/api.php',
    [switch]$DryRun,
    [switch]$SkipGenerate
)

$ErrorActionPreference = 'Stop'

function Normalize-TitleKey {
    param([string]$Title)
    return (($Title ?? '') -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function ConvertTo-WikiTextPlain {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $value = $Text
    $value = $value -replace '(?i)<br\s*/?>', '; '
    $value = $value -replace '\[\[File:[^\]]+\]\]', ''
    $value = $value -replace '\[\[[^|\]]+\|([^\]]+)\]\]', '$1'
    $value = $value -replace '\[\[([^\]]+)\]\]', '$1'
    $value = $value -replace "'''?", ''
    $value = $value -replace '\{\{[^{}]*\}\}', ''
    $value = $value -replace '<[^>]+>', ''
    $value = [System.Net.WebUtility]::HtmlDecode($value)
    $value = $value -replace '\s+', ' '
    return $value.Trim()
}

function ConvertFrom-WikiDate {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $clean = ConvertTo-WikiTextPlain $Text
    $clean = $clean -replace '(?<=\d)(st|nd|rd|th)\b', ''
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($clean, [Globalization.CultureInfo]::GetCultureInfo('en-US'), [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        return $parsed.Date
    }
    return $null
}

function Invoke-WikiApi {
    param([hashtable]$Query)
    $pairs = foreach ($key in $Query.Keys) {
        '{0}={1}' -f [uri]::EscapeDataString($key), [uri]::EscapeDataString([string]$Query[$key])
    }
    $uri = "$($WikiApi)?$(($pairs -join '&'))"
    return (Invoke-WebRequest -Uri $uri -UseBasicParsing).Content | ConvertFrom-Json
}

function Get-PetTemplateFields {
    param([string]$WikiText)
    $fields = [ordered]@{}
    $inTemplate = $false
    $currentKey = $null
    $currentValue = New-Object System.Text.StringBuilder

    foreach ($line in ($WikiText -split "`r?`n")) {
        if (-not $inTemplate) {
            if ($line -match '^\s*\{\{\s*Pets\b') { $inTemplate = $true }
            continue
        }
        if ($line -match '^\s*\}\}') { break }
        if ($line -match '^\s*\|\s*([^=]+?)\s*=\s*(.*)$') {
            if ($currentKey) { $fields[$currentKey] = $currentValue.ToString().Trim() }
            $currentKey = $matches[1].Trim().ToLowerInvariant()
            $currentValue = New-Object System.Text.StringBuilder
            [void]$currentValue.Append($matches[2].Trim())
        } elseif ($currentKey) {
            [void]$currentValue.Append("`n$line")
        }
    }
    if ($currentKey) { $fields[$currentKey] = $currentValue.ToString().Trim() }
    return $fields
}

function Get-OverviewText {
    param([string]$WikiText)
    $match = [regex]::Match($WikiText, '(?s)==\s*Overview\s*==\s*(.+?)(?:\n==|\z)')
    if (-not $match.Success) { return '' }
    $paragraphs = $match.Groups[1].Value -split "(`r?`n){2,}"
    foreach ($paragraph in $paragraphs) {
        $plain = ConvertTo-WikiTextPlain $paragraph
        if ($plain -ne '') { return $plain }
    }
    return ''
}

function Get-Tier {
    param([string]$RawTier, [string[]]$Categories)
    $rarities = @('Common', 'Uncommon', 'Rare', 'Legendary', 'Mythical', 'Divine', 'Prismatic')
    foreach ($rarity in $rarities) {
        if ($Categories -contains $rarity) { return $rarity }
    }
    foreach ($rarity in $rarities) {
        if ($RawTier -match "(?i)(?:File:)?$([regex]::Escape($rarity))Icon\.(?:png|gif|webp)") { return $rarity }
    }
    $plain = ConvertTo-WikiTextPlain $RawTier
    foreach ($rarity in $rarities) {
        if ($plain -match "(?i)\b$rarity\b") { return $rarity }
    }
    return $plain
}

function Get-PetImageFile {
    param([string]$RawImageField, [object]$Page)
    $raw = ($RawImageField ?? '').Trim()
    $candidates = @()

    if ($raw -match '(?i)File:([^|\]\r\n<>]+\.(?:png|gif|webp|jpg|jpeg))') {
        $candidates += $matches[1].Trim()
    } elseif ($raw -match '(?im)^\s*([A-Za-z0-9][^|\r\n<>]+\.(?:png|gif|webp|jpg|jpeg))\s*(?:\||$)') {
        $candidates += $matches[1].Trim()
    } elseif ($raw -match '(?i)([A-Za-z0-9][^|\r\n<>]+\.(?:png|gif|webp|jpg|jpeg))') {
        $candidates += $matches[1].Trim()
    }

    if ($Page -and $Page.images) {
        foreach ($image in @($Page.images)) {
            $filename = ($image.title -replace '^File:', '').Trim()
            if ($filename -match '(?i)Icon\.(png|gif|webp)$') { continue }
            if ($filename -match '(?i)^Rainbow') { continue }
            $candidates += $filename
        }
    }

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate }
    }
    return ''
}

function ConvertTo-PetRecord {
    param(
        [string]$Title,
        [string]$WikiText,
        [string[]]$Categories,
        [string]$ImageUrl
    )
    $fields = Get-PetTemplateFields $WikiText
    $tier = Get-Tier $fields['tier'] $Categories
    $wikiTitle = $Title -replace ' ', '_'
    $wikiUrl = 'https://growagarden.fandom.com/wiki/' + [uri]::EscapeDataString($wikiTitle).Replace('%2F', '/')

    return [pscustomobject]@{
        Title       = $Title
        Tier        = @($tier)
        Obtaining   = @(ConvertTo-WikiTextPlain $fields['obtaining method'])
        HatchChance = @(ConvertTo-WikiTextPlain $fields['hatch_chance'])
        Passive     = @(ConvertTo-WikiTextPlain $fields['passive_ability'])
        DateAdded   = @(ConvertTo-WikiTextPlain $fields['date_added'])
        Image       = $ImageUrl
        WikiPage    = $wikiUrl
        Appearance  = Get-OverviewText $WikiText
    }
}

function Get-AllPetCategoryMembers {
    $members = @()
    $continue = $null
    do {
        $query = @{
            action = 'query'
            format = 'json'
            list = 'categorymembers'
            cmtitle = 'Category:Pets'
            cmnamespace = '0'
            cmlimit = '500'
            cmprop = 'ids|title|timestamp'
        }
        if ($continue) { $query['cmcontinue'] = $continue }
        $response = Invoke-WikiApi $query
        $members += @($response.query.categorymembers)
        $continue = $response.continue.cmcontinue
    } while ($continue)
    return $members
}

function Get-PetPageBatch {
    param([object[]]$Members)
    if ($Members.Count -eq 0) { return @() }
    $titles = ($Members | ForEach-Object { $_.title }) -join '|'
    $response = Invoke-WikiApi @{
        action = 'query'
        format = 'json'
        prop = 'revisions|categories|images'
        rvprop = 'content'
        rvslots = 'main'
        cllimit = 'max'
        imlimit = 'max'
        titles = $titles
    }
    return @($response.query.pages.PSObject.Properties.Value)
}

function Get-ImageUrlMap {
    param([object[]]$Pages)
    $fileTitles = @()
    foreach ($page in $Pages) {
        $content = $page.revisions[0].slots.main.'*'
        $fields = Get-PetTemplateFields $content
        $image = Get-PetImageFile $fields['image1'] $page
        if ($image -ne '') { $fileTitles += "File:$image" }
    }
    $map = @{}
    foreach ($chunkStart in 0..([Math]::Max([Math]::Ceiling($fileTitles.Count / 50) - 1, 0))) {
        $chunk = @($fileTitles | Select-Object -Skip ($chunkStart * 50) -First 50)
        if ($chunk.Count -eq 0) { continue }
        $response = Invoke-WikiApi @{
            action = 'query'
            format = 'json'
            prop = 'imageinfo'
            iiprop = 'url'
            titles = ($chunk -join '|')
        }
        foreach ($imagePage in @($response.query.pages.PSObject.Properties.Value)) {
            if ($imagePage.imageinfo -and $imagePage.imageinfo.Count -gt 0) {
                $map[$imagePage.title] = $imagePage.imageinfo[0].url
            }
        }
    }
    return $map
}

$cutoff = (ConvertFrom-WikiDate $CutoffDate)
if (-not $cutoff) { throw "Could not parse cutoff date: $CutoffDate" }
if (-not (Test-Path $SourceJson)) { throw "Missing source JSON: $SourceJson" }

$existingPets = @(Get-Content $SourceJson -Raw | ConvertFrom-Json)
$existingKeys = @{}
for ($existingIndex = 0; $existingIndex -lt $existingPets.Count; $existingIndex++) {
    $existingKeys[(Normalize-TitleKey $existingPets[$existingIndex].Title)] = $existingIndex
}

Write-Host "Checking Grow a Garden wiki pets added after $($cutoff.ToString('yyyy-MM-dd'))..."
$members = @(Get-AllPetCategoryMembers)
$newPets = @()
$refreshedPets = @()
$recentPets = @()

for ($i = 0; $i -lt $members.Count; $i += 50) {
    $memberChunk = @($members | Select-Object -Skip $i -First 50)
    $pages = @(Get-PetPageBatch $memberChunk)
    $imageMap = Get-ImageUrlMap $pages
    foreach ($page in $pages) {
        if (-not $page.revisions) { continue }
        $title = $page.title
        $content = $page.revisions[0].slots.main.'*'
        $fields = Get-PetTemplateFields $content
        $dateAdded = ConvertFrom-WikiDate $fields['date_added']
        $categoryMember = $memberChunk | Where-Object { $_.title -eq $title } | Select-Object -First 1
        $categoryTimestamp = if ($categoryMember) { ([datetime]$categoryMember.timestamp).Date } else { $null }
        $addedAfterCutoff = ($dateAdded -and $dateAdded -gt $cutoff) -or (-not $dateAdded -and $categoryTimestamp -and $categoryTimestamp -gt $cutoff)
        if (-not $addedAfterCutoff) { continue }

        $key = Normalize-TitleKey $title

        $categories = @()
        if ($page.categories) {
            $categories = @($page.categories | ForEach-Object { ($_.title -replace '^Category:', '') })
        }
        $imageFile = Get-PetImageFile $fields['image1'] $page
        $imageUrl = ''
        if ($imageFile -ne '' -and $imageMap.ContainsKey("File:$imageFile")) { $imageUrl = $imageMap["File:$imageFile"] }
        $record = ConvertTo-PetRecord $title $content $categories $imageUrl
        $recentPets += $record
        if ($existingKeys.ContainsKey($key)) {
            $existingPets[$existingKeys[$key]] = $record
            $refreshedPets += $title
            continue
        }

        $newPets += $record
        $existingKeys[$key] = $existingPets.Count + $newPets.Count - 1
    }
}

$reportPath = "new_pets_since_$($cutoff.ToString('yyyy-MM-dd')).json"
$recentPets | Sort-Object Title | ConvertTo-Json -Depth 10 | Set-Content -Path $reportPath -Encoding UTF8

if ($newPets.Count -eq 0) {
    Write-Host "No missing pets found after $($cutoff.ToString('yyyy-MM-dd'))."
} else {
    Write-Host "Found $($newPets.Count) missing pet(s): $((@($newPets | ForEach-Object Title)) -join ', ')"
}

if ($refreshedPets.Count -gt 0) {
    Write-Host "Refreshed $($refreshedPets.Count) existing recent pet(s): $((@($refreshedPets | Sort-Object -Unique)) -join ', ')"
}

if (-not $DryRun -and ($newPets.Count -gt 0 -or $refreshedPets.Count -gt 0)) {
    @($existingPets + $newPets) |
        Sort-Object Title |
        ConvertTo-Json -Depth 10 |
        Set-Content -Path $SourceJson -Encoding UTF8
    Write-Host "Updated $SourceJson"
}

if (-not $DryRun -and -not $SkipGenerate) {
    $python = 'python'
    $bundledPython = 'C:\Users\klsma\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path $bundledPython) { $python = $bundledPython }

    & .\regen_by_passive.ps1
    if ($LASTEXITCODE -ne 0) { throw 'regen_by_passive.ps1 failed' }
    & $python .\scripts\build_alphabetical_guide.py
    if ($LASTEXITCODE -ne 0) { throw 'scripts/build_alphabetical_guide.py failed' }
    & pandoc grow_a_garden_pet_guide_alphabetical.md -o grow_a_garden_ultimate_pet_guide_alphabetical.docx --resource-path=.,media,assets --extract-media=./media
    if ($LASTEXITCODE -ne 0) { throw 'pandoc failed for alphabetical DOCX' }
    & $python .\scripts\make_two_column_indexes.py grow_a_garden_ultimate_pet_guide_alphabetical.docx
    if ($LASTEXITCODE -ne 0) { throw 'scripts/make_two_column_indexes.py failed' }
    Write-Host 'Generated grow_a_garden_pet_guide_by_passive.md'
    Write-Host 'Generated grow_a_garden_pet_guide_alphabetical.md'
    Write-Host 'Generated grow_a_garden_ultimate_pet_guide_alphabetical.docx'
}
