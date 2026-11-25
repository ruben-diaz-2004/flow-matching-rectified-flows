# Exporta cada diapositiva a PNG usando PowerPoint (para revisión visual).
param([string]$Deck = "FlowMatching_RectifiedFlows.pptx", [string]$Out = "render")
$deckPath = (Resolve-Path $Deck).Path
$outDir = Join-Path (Get-Location) $Out
if (Test-Path $outDir) { Remove-Item -Recurse -Force $outDir -Confirm:$false }
New-Item -ItemType Directory -Force $outDir | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($deckPath, $true, $false, $false)
$i = 0
foreach ($s in $pres.Slides) {
    $i++
    $s.Export((Join-Path $outDir ("slide-{0:D2}.png" -f $i)), "PNG", 1600, 900)
}
$pres.Close()
$app.Quit()
"exported $i slides to $outDir"
