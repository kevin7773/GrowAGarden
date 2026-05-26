param()

$inJson = 'pets_full.json'
$outMd = 'grow_a_garden_pet_guide_by_passive.md'
$outDocx = 'grow_a_garden_pet_guide_by_passive.docx'
$mediaDir = 'media'

if (-not (Test-Path $inJson)) { Write-Error "Missing $inJson"; exit 1 }
if (-not (Test-Path $mediaDir)) { New-Item -ItemType Directory -Path $mediaDir | Out-Null }

$blacklist = @(
    'Christmas Harvest Event', 'Common Egg', 'Cooking Event', 'Garden Ascension', 'Garden Coins',
    'Pet Eggs', 'Prehistoric Event', 'Premium Fall Egg', 'Sheckles', 'Small Toy',
    'Small Treat', 'Golden Acorn', 'Garden Guide', 'Pet Pouch'
)

$json = Get-Content $inJson -Raw | ConvertFrom-Json

# Normalize entries and filter
$pets = @()
$seen = @{}
foreach ($p in $json) {
    if ($blacklist -contains $p.Title) { continue }
    $title = $p.Title.Trim()
    if ($title -eq '') { continue }
    $tier = $null
    if ($p.Tier -is [System.Array]) { $tier = ($p.Tier -join ' ' ).Trim() } else { $tier = ($p.Tier ?? '').Trim() }
    $hatch = ''
    if ($p.HatchChance -is [System.Array]) { $hatch = ($p.HatchChance -join ' ' ).Trim() } else { $hatch = ($p.HatchChance ?? '').Trim() }
    if ($hatch -ne '') { $hatch = [regex]::Replace($hatch,'(%)(?=[A-Za-z])','$1`n') }
    $obt = ''
    if ($p.Obtaining -is [System.Array]) {
        $replacement = '$1' + [Environment]::NewLine + '$2'
        $obt = ($p.Obtaining | ForEach-Object {
            $item = ($_ ?? '').Trim()
            if ($item -ne '') { [regex]::Replace($item, '([a-z])([A-Z])', $replacement) } else { $item }
        }) -join [Environment]::NewLine
        $obt = $obt.Trim()
    } else {
        $obt = ($p.Obtaining ?? '').Trim()
        if ($obt -ne '') { $replacement = '$1' + [Environment]::NewLine + '$2'; $obt = [regex]::Replace($obt, '([a-z])([A-Z])', $replacement).Trim() }
    }
    $passive = ''
    if ($p.Passive -is [System.Array]) { $passive = ($p.Passive -join ' ' ).Trim() } else { $passive = ($p.Passive ?? '').Trim() }
    $date = ''
    if ($p.DateAdded -is [System.Array]) { $date = ($p.DateAdded -join ' ' ).Trim() } else { $date = ($p.DateAdded ?? '').Trim() }
    $imgUrl = ($p.Image ?? '').Trim()
    $wiki = ($p.WikiPage ?? '').Trim()
    $appearance = ($p.Appearance ?? '').Trim()

    # download image to media dir
    $imgLocal = ''
    if ($imgUrl -ne '') {
        try {
            $ext = [System.IO.Path]::GetExtension($imgUrl.Split('?')[0])
            if ([string]::IsNullOrWhiteSpace($ext)) { $ext = '.png' }
            $safe = ($title -replace '[^A-Za-z0-9\- ]','') -replace '\s+','_' 
            $imgLocal = Join-Path $mediaDir ("$safe$ext")
            if (-not (Test-Path $imgLocal)) {
                Invoke-WebRequest -Uri $imgUrl -OutFile $imgLocal -UseBasicParsing -ErrorAction SilentlyContinue
            }
        } catch { $imgLocal = '' }
    }

    # Deduplicate by a normalized title (ignore spaces/punctuation) to treat "Bumble Bee" and "Bumblebee" as the same
    $normTitle = ($title -replace '[^A-Za-z0-9]','').ToLower()
    # Normalize display name for known duplicates (prefer the concatenated form)
    if ($normTitle -eq 'bumblebee') { $title = 'Bumblebee' }
    if ($seen.ContainsKey($normTitle)) { continue }
    $seen[$normTitle] = $true

    $pets += [pscustomobject]@{
        Title=$title; Tier=$tier; Hatch=$hatch; Obtaining=$obt; Passive=$passive; Date=$date; ImageLocal=$imgLocal; ImageUrl=$imgUrl; Wiki=$wiki; Appearance=$appearance
    }
}

# Slug helper
$makeSlug = { param($s) $slug = $s.ToLower() -replace '[^a-z0-9\s-]',''; $slug = $slug -replace '\s+','-'; return $slug.Trim('-') }

# Define passive categories and keywords (expanded per user request)
$categories = @{
    'Cosmetics' = @('cosmetic','cosmetics')
    'Crafting' = @('craft','crafting','crafted')
    'Egg Helpers' = @('chest','chests','egg','eggs','hatch','hatched')
    'Gear' = @('gear','shop','reward','refund')
    'Harvest or Seed Helpers' = @('harvest','harvested','seed','collecting')
    'Mutators' = @('mutate','mutation','mutations','mutates','variant','mutating','chakra','turn a nearby fruit','turns a nearby crop','pollinates','nearby fruit','random fruit','nearby crop')
    'Pet Helpers' = @('xp','experience','cooldown','copies','ability','hunger','base weight')
    'Player Interactive' = @('player','increased movement speed')
    'Plant Growth' = @('grow','growth','size','nap')
    'Levellers' = @('level','levels')
}

$manualCategoryOverrides = @{
    'Woodpecker' = 'Harvest or Seed Helpers'
    'Lyrebird' = 'Harvest or Seed Helpers'
}

# Assign pets to categories (first matching category) or Special
$catMap = @{}
foreach ($k in $categories.Keys) { $catMap[$k] = [System.Collections.ArrayList]@() }
$catMap['Special'] = [System.Collections.ArrayList]@()

# Index map collects every pet that matches each category (allows multi-category index entries)
$indexMap = @{}
foreach ($k in $categories.Keys) { $indexMap[$k] = [System.Collections.ArrayList]@() }
$indexMap['Special'] = [System.Collections.ArrayList]@()

foreach ($pp in $pets) {
    $matchedCategories = @()
    $pl = ($pp.Passive ?? '').ToLower()

    if ($manualCategoryOverrides.ContainsKey($pp.Title)) {
        $matchedCategories = @($manualCategoryOverrides[$pp.Title])
    } else {
        foreach ($k in $categories.Keys) {
            foreach ($kw in $categories[$k]) {
                # match whole words to avoid accidental substring matches
                $pattern = '\b' + [regex]::Escape($kw) + '\b'
                if ($pl -match $pattern) { $matchedCategories += $k; break }
            }
        }
    }

    if ($matchedCategories.Count -gt 1 -and $matchedCategories -contains 'Player Interactive') {
        $matchedCategories = @($matchedCategories | Where-Object { $_ -ne 'Player Interactive' })
    }
    if ($matchedCategories.Count -gt 0) {
        # primary assignment: first matching category
        $primary = $matchedCategories[0]
        $catMap[$primary].Add($pp) | Out-Null
        # record in index map for every matched category (avoid duplicates)
        foreach ($cat in $matchedCategories) {
            if (-not $indexMap[$cat].Contains($pp)) { $indexMap[$cat].Add($pp) | Out-Null }
        }
    } else {
        $catMap['Special'].Add($pp) | Out-Null
        if (-not $indexMap['Special'].Contains($pp)) { $indexMap['Special'].Add($pp) | Out-Null }
    }
}

# Fixed category order and top navigation rows
$categoryOrder = @(
    'Cosmetics', 'Crafting', 'Egg Helpers', 'Gear',
    'Harvest or Seed Helpers', 'Levellers', 'Mutators', 'Pet Helpers',
    'Plant Growth', 'Player Interactive', 'Special'
)
$sortedCats = $categoryOrder

# Build markdown
$sb = New-Object System.Text.StringBuilder
$sb.AppendLine('# Grow a Garden Pet Guide — Sorted by Passive Ability') | Out-Null
$sb.AppendLine('') | Out-Null
$sb.AppendLine('Source: https://growagarden.fandom.com/wiki/Grow_a_Garden_Wiki') | Out-Null
$sb.AppendLine('') | Out-Null
# Top links in three fixed rows
$topRows = @(
    $categoryOrder[0..3],
    $categoryOrder[4..7],
    $categoryOrder[8..10]
)
foreach ($row in $topRows) {
    $sb.AppendLine("<div align='center'>" + (($row | ForEach-Object { "[$_]" + "(#" + (& $makeSlug $_) + ")" }) -join ' | ') + "</div>") | Out-Null
    $sb.AppendLine('') | Out-Null
}

foreach ($c in $sortedCats) {
    $cSlug = & $makeSlug $c
    $sb.AppendLine('\pagebreak') | Out-Null
    $sb.AppendLine('') | Out-Null
    $sb.AppendLine("## $c {#$cSlug}") | Out-Null
    $sb.AppendLine('') | Out-Null
    $items = $catMap[$c] | Sort-Object Title
    foreach ($pp in $items) {
        $slug = & $makeSlug $pp.Title
        $sb.AppendLine("### $($pp.Title) {#$slug}") | Out-Null
        $sb.AppendLine('') | Out-Null
        if ($pp.ImageLocal -ne '' -and (Test-Path $pp.ImageLocal)) {
            $rel = $pp.ImageLocal -replace '\\','/'
            $sb.AppendLine("![$($pp.Title)]($rel){width=80px}") | Out-Null
        } elseif ($pp.ImageUrl -ne '') {
            $sb.AppendLine("![$($pp.Title)]($($pp.ImageUrl)){width=80px}") | Out-Null
        }
        $sb.AppendLine('') | Out-Null
        $sb.AppendLine("- **Tier:** $($pp.Tier)") | Out-Null
        $sb.AppendLine('') | Out-Null
        if ($pp.Hatch -ne '') { $sb.AppendLine("- **Hatch Chance:**"); $pp.Hatch -split "`n" | ForEach-Object { if ($_ -ne '') { $sb.AppendLine("  - $_") | Out-Null } }; $sb.AppendLine('') | Out-Null } else { $sb.AppendLine("- **Hatch Chance:** N/A") | Out-Null; $sb.AppendLine('') | Out-Null }
        if ($pp.Obtaining -ne '') {
            $sb.AppendLine('- **Obtaining Method:**') | Out-Null
            $pp.Obtaining -split "`n" | ForEach-Object { if ($_ -ne '') { $sb.AppendLine("  $_") | Out-Null } }
            $sb.AppendLine('') | Out-Null
        } else {
            $sb.AppendLine('- **Obtaining Method:** N/A') | Out-Null
            $sb.AppendLine('') | Out-Null
        }
        if ($pp.Passive -ne '') {
            $sb.AppendLine("- **Passive Ability:** $($pp.Passive)") | Out-Null
        } else {
            $sb.AppendLine('- **Passive Ability:** N/A') | Out-Null
        }
        $sb.AppendLine('') | Out-Null
        $sb.AppendLine("- **Appearance:** $($pp.Appearance)") | Out-Null
        $sb.AppendLine('') | Out-Null
        $sb.AppendLine("- **Date Added:** $($pp.Date)") | Out-Null
        $sb.AppendLine('') | Out-Null
        if ($pp.Wiki -ne '') { $sb.AppendLine("- **Wiki Page:** <$($pp.Wiki)>") | Out-Null; $sb.AppendLine('') | Out-Null }
        $sb.AppendLine('') | Out-Null
    }
}

# Pet index grouped by passive category
$sb.AppendLine('\pagebreak') | Out-Null
$sb.AppendLine('') | Out-Null
$sb.AppendLine('## Pet Index (by Passive Ability Category)') | Out-Null
$sb.AppendLine('') | Out-Null
foreach ($c in $sortedCats) {
    $cSlug = & $makeSlug $c
    $sb.AppendLine("### $c {#$($cSlug)-index}") | Out-Null
    # Use indexMap so pets appear under every category they match
    $items = $indexMap[$c] | Sort-Object Title
    foreach ($pp in $items) { $sb.AppendLine("- [" + $pp.Title + "](#" + (& $makeSlug $pp.Title) + ")" ) | Out-Null }
    $sb.AppendLine('') | Out-Null
}

Set-Content -Path $outMd -Value $sb.ToString() -Encoding UTF8
Write-Host "WROTE $outMd"

# Convert with pandoc
$pandoc = Get-Command pandoc -ErrorAction SilentlyContinue
if ($pandoc) {
    & pandoc $outMd -o $outDocx --resource-path=.,$mediaDir --extract-media=./$mediaDir
    if ($LASTEXITCODE -eq 0 -and (Test-Path $outDocx)) { Write-Host 'CONVERSION_OK' } else { Write-Host 'CONVERSION_FAILED' }
} else { Write-Host 'PANDOC_NOT_FOUND' }
