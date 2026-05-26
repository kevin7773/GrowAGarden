$in = 'grow_a_garden_pet_guide.md'
$outMd = 'grow_a_garden_pet_guide_for_docx.md'
$outDocx = 'grow_a_garden_pet_guide_updated.docx'

$text = Get-Content $in -Raw -Encoding UTF8

# Insert tier links after Source line
$sourceMatch = [regex]::Match($text, '(?m)^Source:.*$')
if ($sourceMatch.Success) {
    $tiers = [regex]::Matches($text, '(?m)^##\s+(.*)$') | ForEach-Object { $_.Groups[1].Value }
    if ($tiers.Count -gt 0) {
        $makeSlug = { param($t) $s = $t.ToLower() -replace "[^a-z0-9\s-]", '' ; $s = $s -replace '\\s+', '-' ; return $s }
        $links = ($tiers | ForEach-Object { "[$_]($( & $makeSlug $_ | ForEach-Object { $_ }) )" }) -join ' | '
        # The above creates wrong link syntax due to scriptblock; build manually
        $linksArr = @()
        foreach ($t in $tiers) {
            $s = $t.ToLower() -replace "[^a-z0-9\\s-]", '' -replace '\\s+', '-'
            $linksArr += "[$t](#$s)"
        }
        $links = $linksArr -join ' | '
        $insertPos = $sourceMatch.Index + $sourceMatch.Length
        $text = $text.Substring(0,$insertPos) + "`n`n" + $links + "`n`n" + $text.Substring($insertPos)
    }
}

# Replace pet blocks
$pattern = [regex]::new('(?ms)^(###\s+(.+?)\r?\n)((?:(?!^###\s|^##\s).*(?:\r?\n))*)')
$result = $pattern.Replace($text, [System.Text.RegularExpressions.MatchEvaluator]{ param($m)
    $heading = $m.Groups[2].Value.Trim()
    $body = $m.Groups[3].Value
    $imgMatch = [regex]::Match($body, '- Image:\s*!\[.*?\]\((.*?)\)')
    if (-not $imgMatch.Success) { return $m.Value }
    $img = $imgMatch.Groups[1].Value.Trim()
    $bodyClean = [regex]::Replace($body, '(?m)^- Image:.*\r?\n', '')
    $bodyHtml = $bodyClean.TrimEnd()
    $html = @"
<table>
<tr>
<td valign="top"><img src="$img" width="120" height="120" /></td>
<td valign="top">
<h3>$heading</h3>
$bodyHtml
</td>
</tr>
</table>
"@
    return $html
})

Set-Content -Path $outMd -Value $result -Encoding UTF8
Write-Host "WROTE $outMd"

# Run pandoc
$pandoc = Get-Command pandoc -ErrorAction SilentlyContinue
if ($pandoc) {
    & pandoc $outMd -o $outDocx --resource-path=.
    if (Test-Path $outDocx) { Write-Host 'CONVERSION_OK' } else { Write-Host 'CONVERSION_FAILED' }
} else {
    Write-Host 'PANDOC_NOT_FOUND'
}
