# insert_index_table.ps1
param(
    [string]$docxPath = 'grow_a_garden_pet_guide_by_passive.docx',
    [string]$outDocx = 'grow_a_garden_pet_guide_by_passive_indexed.docx'
)
$base = Split-Path -Parent $docxPath
if ([string]::IsNullOrEmpty($base)) { $base = (Get-Location).Path }
$work = Join-Path $base '_docx_edit'
if (Test-Path $work) { Remove-Item -Recurse -Force $work }
New-Item -ItemType Directory -Path $work | Out-Null
Expand-Archive -Path $docxPath -DestinationPath $work -Force
$docXml = Join-Path $work 'word\document.xml'
[xml]$doc = Get-Content $docXml -Raw
# Find the heading paragraph node that contains the exact heading text
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$paras = $doc.SelectNodes('//w:body/w:p', $ns)
$startIndex = -1
for ($i=0; $i -lt $paras.Count; $i++) {
    $p = $paras[$i]
    if ($p.InnerText -eq 'Pet Index (by Passive Ability Category)') { $startIndex = $i; break }
}
if ($startIndex -lt 0) { Write-Host 'HEADING_NOT_FOUND'; exit 1 }

# center the first few top-link paragraphs that contain category hyperlinks
$topLinkCount = 3
$found = 0
foreach ($p in $paras) {
    if ($p.SelectSingleNode('.//w:hyperlink', $ns)) {
        $pPr = $p.SelectSingleNode('w:pPr', $ns)
        if (-not $pPr) { $pPr = $doc.CreateElement('w','pPr',$w); $p.PrependChild($pPr) | Out-Null }
        $jc = $pPr.SelectSingleNode('w:jc', $ns)
        if (-not $jc) {
            $jcEl = $doc.CreateElement('w','jc',$w)
            $attr = $doc.CreateAttribute('w','val',$w); $attr.Value = 'center'; $jcEl.Attributes.Append($attr) | Out-Null
            $pPr.AppendChild($jcEl) | Out-Null
        }
        $found++
        if ($found -ge $topLinkCount) { break }
    }
}
# Collect following paragraph nodes up to next Heading2 or end
$collected = @()
for ($j=$startIndex+1; $j -lt $paras.Count; $j++) {
    $txt = $paras[$j].InnerText
    # stop if next Heading2 found
    $pPr = $paras[$j].SelectSingleNode('w:pPr/w:pStyle', $ns)
    if ($pPr -and $pPr.GetAttribute('w:val') -eq 'Heading2') { break }
    $collected += $paras[$j]
}
# split into two halves
$half = [math]::Ceiling($collected.Count/2)
$left = $collected[0..($half-1)]
$right = if ($half -lt $collected.Count) { $collected[$half..($collected.Count-1)] } else { @() }
# create table element under Word namespace
$w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$tbl = $doc.CreateElement('w','tbl',$w)
$tblPr = $doc.CreateElement('w','tblPr',$w); $tbl.AppendChild($tblPr) | Out-Null
$tblGrid = $doc.CreateElement('w','tblGrid',$w); $tbl.AppendChild($tblGrid) | Out-Null
for ($c=0;$c -lt 2;$c++) {
    $col = $doc.CreateElement('w','gridCol',$w)
    $attr = $doc.CreateAttribute('w','w',$w)
    $attr.Value = '5000'
    $col.Attributes.Append($attr) | Out-Null
    $tblGrid.AppendChild($col) | Out-Null
}
$tr = $doc.CreateElement('w','tr',$w); $tbl.AppendChild($tr) | Out-Null
$tcLeft = $doc.CreateElement('w','tc',$w); $tr.AppendChild($tcLeft) | Out-Null
$tcPrL = $doc.CreateElement('w','tcPr',$w); $tcLeft.AppendChild($tcPrL) | Out-Null
$tcWl = $doc.CreateElement('w','tcW',$w)
$attr1 = $doc.CreateAttribute('w','w',$w); $attr1.Value = '5000'; $tcWl.Attributes.Append($attr1) | Out-Null
$attr2 = $doc.CreateAttribute('w','type',$w); $attr2.Value = 'dxa'; $tcWl.Attributes.Append($attr2) | Out-Null
$tcPrL.AppendChild($tcWl) | Out-Null
# move left paras into tcLeft (preserve hyperlinks/bookmarks)
foreach ($node in $left) {
    $node.ParentNode.RemoveChild($node) | Out-Null
    $tcLeft.AppendChild($node) | Out-Null
}
$tcRight = $doc.CreateElement('w','tc',$w); $tr.AppendChild($tcRight) | Out-Null
$tcPrR = $doc.CreateElement('w','tcPr',$w); $tcRight.AppendChild($tcPrR) | Out-Null
$tcWr = $doc.CreateElement('w','tcW',$w)
$attr3 = $doc.CreateAttribute('w','w',$w); $attr3.Value = '5000'; $tcWr.Attributes.Append($attr3) | Out-Null
$attr4 = $doc.CreateAttribute('w','type',$w); $attr4.Value = 'dxa'; $tcWr.Attributes.Append($attr4) | Out-Null
$tcPrR.AppendChild($tcWr) | Out-Null
# move right paras into tcRight (preserve hyperlinks/bookmarks)
foreach ($node in $right) {
    $node.ParentNode.RemoveChild($node) | Out-Null
    $tcRight.AppendChild($node) | Out-Null
}
# original paragraphs were moved into the table cells above
# Insert table after heading paragraph
$headingNode = $paras[$startIndex]
$headingNode.ParentNode.InsertAfter($tbl, $headingNode) | Out-Null
# Save modified document.xml
$doc.Save($docXml)
# Repack into new docx
if ([string]::IsNullOrEmpty($base)) { $outPath = Join-Path (Get-Location).Path $outDocx } else { $outPath = Join-Path $base $outDocx }
if (Test-Path $outPath) { Remove-Item $outPath -Force }
Push-Location $work
Compress-Archive -Path * -DestinationPath $outPath -Force
Pop-Location
Write-Host 'REPACKED' 
