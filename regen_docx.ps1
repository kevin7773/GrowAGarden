param()

$inJson = 'pets_full.json'
$outMd = 'grow_a_garden_pet_guide_for_docx.md'
$outDocx = 'grow_a_garden_pet_guide_fixed.docx'
$mediaDir = 'media'

if (-not (Test-Path $inJson)) { Write-Error "Missing $inJson"; exit 1 }
if (-not (Test-Path $mediaDir)) { New-Item -ItemType Directory -Path $mediaDir | Out-Null }

$blacklist = @(
    'Christmas Harvest Event', 'Common Egg', 'Cooking Event', 'Garden Ascension', 'Garden Coins',
    'Pet Eggs', 'Prehistoric Event', 'Premium Fall Egg', 'Sheckles', 'Small Toy',
    'Small Treat', 'Golden Acorn', 'Garden Guide'
)

$json = Get-Content $inJson -Raw | ConvertFrom-Json

# Normalize entries and filter
$pets = @()
foreach ($p in $json) {
    if ($blacklist -contains $p.Title) { continue }
    $title = $p.Title.Trim()
    if ($title -eq '') { continue }
    $tier = $null
    if ($p.Tier -is [System.Array]) { $tier = ($p.Tier -join ' ' ).Trim() } else { $tier = ($p.Tier ?? '').Trim() }
    if ($title -eq 'Wolf') { $tier = 'Legendary' }
    $hatch = ''
    if ($p.HatchChance -is [System.Array]) { $hatch = ($p.HatchChance -join ' ' ).Trim() } else { $hatch = ($p.HatchChance ?? '').Trim() }
    # split hatch chances where percent is followed immediately by a letter
    if ($hatch -ne '') { $hatch = [regex]::Replace($hatch,'(%)(?=[A-Za-z])','$1`n') }
    $obt = ''
    if ($p.Obtaining -is [System.Array]) { $obt = ($p.Obtaining -join ' ' ).Trim() } else { $obt = ($p.Obtaining ?? '').Trim() }
    $passive = ''
    if ($p.Passive -is [System.Array]) { $passive = ($p.Passive -join ' ' ).Trim() } else { $passive = ($p.Passive ?? '').Trim() }
    $date = ''
    if ($p.DateAdded -is [System.Array]) { $date = ($p.DateAdded -join ' ' ).Trim() } else { $date = ($p.DateAdded ?? '').Trim() }
    $imgUrl = ($p.Image ?? '').Trim()
    $wiki = ($p.WikiPage ?? '').Trim()
    $appearance = ($p.Appearance ?? '').Trim()

    # download image to media dir, filename based on title
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

    $pets += [pscustomobject]@{
        Title=$title; Tier=$tier; Hatch=$hatch; Obtaining=$obt; Passive=$passive; Date=$date; ImageLocal=$imgLocal; ImageUrl=$imgUrl; Wiki=$wiki; Appearance=$appearance
    }
}

# Group by Tier preserving first-seen order
$tierOrderRaw = @()
foreach ($p in $pets) { if (-not ($tierOrderRaw -contains $p.Tier)) { $tierOrderRaw += $p.Tier } }

# Desired tier order
$desiredOrder = @('Common','Uncommon','Rare','Legendary','Mythical','Divine','Prismatic')

# Build slug helper
$makeSlug = { param($s) $slug = $s.ToLower() -replace '[^a-z0-9\s-]',''; $slug = $slug -replace '\s+','-'; return $slug.Trim('-') }

# Ensure tiers follow desired order but include any missing tiers at end
$tierOrder = @()
foreach ($t in $desiredOrder) { if ($tierOrderRaw -contains $t) { $tierOrder += $t } }
foreach ($t in $tierOrderRaw) { if (-not ($tierOrder -contains $t)) { $tierOrder += $t } }

# Top-of-document tier links (should target headings with explicit IDs)
$tierLinks = ($tierOrder | ForEach-Object { $t = $_; "[$t](#" + (& $makeSlug $t) + ")" }) -join ' | '

# Build content markdown
$sbContent = New-Object System.Text.StringBuilder
$sbContent.AppendLine('# Grow a Garden Pet Guide') | Out-Null
$sbContent.AppendLine('') | Out-Null
$sbContent.AppendLine('Source: https://growagarden.fandom.com/wiki/Grow_a_Garden_Wiki') | Out-Null
$sbContent.AppendLine('') | Out-Null
$sbContent.AppendLine($tierLinks) | Out-Null
$sbContent.AppendLine('') | Out-Null

# Build the main content (tiers and pets)
$sbRight = New-Object System.Text.StringBuilder
$tierIndex = 0
foreach ($t in $tierOrder) {
    if ($tierIndex -gt 0) {
        $sbRight.AppendLine('\pagebreak') | Out-Null
        $sbRight.AppendLine('') | Out-Null
    }
    $tierSlug = & $makeSlug $t
    $sbRight.AppendLine("## $t {#$tierSlug}") | Out-Null
    $sbRight.AppendLine('') | Out-Null
    $ps = $pets | Where-Object { $_.Tier -eq $t } | Sort-Object Title
    foreach ($pp in $ps) {
        $slug = & $makeSlug $pp.Title
        $sbRight.AppendLine("### $($pp.Title) {#$slug}") | Out-Null
        $sbRight.AppendLine('') | Out-Null

        # Image only, 80px width
        if ($pp.ImageLocal -ne '' -and (Test-Path $pp.ImageLocal)) {
            $rel = $pp.ImageLocal -replace '\\','/'
            $sbRight.AppendLine("![$($pp.Title)]($rel){width=80px}") | Out-Null
        } elseif ($pp.ImageUrl -ne '') {
            $sbRight.AppendLine("![$($pp.Title)]($($pp.ImageUrl)){width=80px}") | Out-Null
        }
        $sbRight.AppendLine('') | Out-Null

        # Tier
        $sbRight.AppendLine("- **Tier:** $($pp.Tier)") | Out-Null
        $sbRight.AppendLine('') | Out-Null

        # Hatch chances: each on its own line
        if ($pp.Hatch -ne '') {
            $sbRight.AppendLine("- **Hatch Chance:**") | Out-Null
            $pp.Hatch -split "`n" | ForEach-Object { if ($_ -ne '') { $sbRight.AppendLine("  - $_") | Out-Null } }
            $sbRight.AppendLine('') | Out-Null
        } else {
            $sbRight.AppendLine("- **Hatch Chance:** N/A") | Out-Null
            $sbRight.AppendLine('') | Out-Null
        }

        # Obtaining
        $sbRight.AppendLine("- **Obtaining Method:** $($pp.Obtaining)") | Out-Null
        $sbRight.AppendLine('') | Out-Null

        # Passive
        $sbRight.AppendLine("- **Passive Ability:** $($pp.Passive)") | Out-Null
        $sbRight.AppendLine('') | Out-Null

        # Appearance then Date
        $sbRight.AppendLine("- **Appearance:** $($pp.Appearance)") | Out-Null
        $sbRight.AppendLine('') | Out-Null
        $sbRight.AppendLine("- **Date Added:** $($pp.Date)") | Out-Null
        $sbRight.AppendLine('') | Out-Null

        if ($pp.Wiki -ne '') { $sbRight.AppendLine("- **Wiki Page:** <$($pp.Wiki)>") | Out-Null; $sbRight.AppendLine('') | Out-Null }

        # spacer between pets
        $sbRight.AppendLine('') | Out-Null
    }
    $tierIndex++
}

# Pet index at end grouped by desired tier order
$sbRight.AppendLine('') | Out-Null
$sbRight.AppendLine('## Pet Index') | Out-Null
$sbRight.AppendLine('') | Out-Null
foreach ($t in $desiredOrder) {
    $tierSlug = & $makeSlug $t
    $sbRight.AppendLine("### $t {#$tierSlug-index}") | Out-Null
    $ps = $pets | Where-Object { $_.Tier -eq $t } | Sort-Object Title
    foreach ($pp in $ps) { $sbRight.AppendLine("- [" + $pp.Title + "](#" + (& $makeSlug $pp.Title) + ")" ) | Out-Null }
    $sbRight.AppendLine('') | Out-Null
}

# Save markdown
Set-Content -Path $outMd -Value ($sbContent.ToString() + "`n" + $sbRight.ToString()) -Encoding UTF8
Write-Host "WROTE $outMd"

# Run pandoc with media extraction to ensure images included
$pandoc = Get-Command pandoc -ErrorAction SilentlyContinue
if ($pandoc) {
    & pandoc $outMd -o $outDocx --resource-path=.,$mediaDir --extract-media=./$mediaDir
    if ($LASTEXITCODE -eq 0 -and (Test-Path $outDocx)) {
        Write-Host 'CONVERSION_OK'
    } else {
        Write-Host 'CONVERSION_FAILED'
    }
} else {
    Write-Host 'PANDOC_NOT_FOUND'
}
