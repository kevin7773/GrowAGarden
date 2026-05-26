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

Write-Host "Total pets after normalization: $($pets.Count)"

# Slug helper
$makeSlug = { param($s) $slug = $s.ToLower() -replace '[^a-z0-9\s-]',''; $slug = $slug -replace '\s+','-'; return $slug.Trim('-') }

# Define passive categories and keywords (expanded per user request)
$categories = @{
    'Cosmetics' = @('cosmetic','cosmetics')
    'Crafting' = @('craft','crafting','crafted')
    'Egg Helpers' = @('chest','chests','egg','eggs','hatch','hatched')
    'Gear' = @('gear','shop','reward','refund')
    'Harvest or Seed Helpers' = @('harvest','harvested','seed','collecting','duplicate')
    'Mutators' = @('mutate','mutation','mutations','mutates','variant','mutating','chakra','turn a nearby fruit','turns a nearby crop','pollinates','nearby fruit','random fruit','nearby crop')
    'Pet Helpers' = @('xp','experience','cooldown','copies','ability','hunger','base weight')
    'Player Interactive' = @('player','increased movement speed')
    'Plant Growth' = @('grow','growth','size','nap')
    'Levellers' = @('level','levels')
}

# Assign pets to categories (first matching category) or Special
$catMap = @{}
foreach ($k in $categories.Keys) { 
    $catMap[$k] = [System.Collections.ArrayList]@()
    Write-Host "Init category: $k"
}
$catMap['Special'] = [System.Collections.ArrayList]@()

Write-Host "Processing pets..."
$errorCount = 0
foreach ($pp in $pets) {
    try {
        $matchedCategories = @()
        $pl = ($pp.Passive ?? '').ToLower()
        foreach ($k in $categories.Keys) {
            foreach ($kw in $categories[$k]) {
                # match whole words to avoid accidental substring matches
                $pattern = '\b' + [regex]::Escape($kw) + '\b'
                if ($pl -match $pattern) { $matchedCategories += $k; break }
            }
        }
        if ($matchedCategories.Count -gt 1 -and $matchedCategories -contains 'Player Interactive') {
            $matchedCategories = $matchedCategories | Where-Object { $_ -ne 'Player Interactive' }
        }
        if ($matchedCategories.Count -gt 0) {
            $catKey = $matchedCategories[0]
            if ($null -eq $catMap[$catKey]) {
                Write-Host "ERROR: catMap[$catKey] is null for pet '$($pp.Title)'"
                $errorCount++
            } else {
                $catMap[$catKey].Add($pp) | Out-Null
            }
        } else {
            $catMap['Special'].Add($pp) | Out-Null
        }
    } catch {
        Write-Host "ERROR processing $($pp.Title): $_"
        $errorCount++
    }
}

Write-Host "Total errors during categorization: $errorCount"
Write-Host "Category map keys: $($catMap.Keys -join ', ')"
foreach ($k in $catMap.Keys) {
    Write-Host "  $($k): $($catMap[$k].Count) pets"
}
