$find = Read-Host "Find"
$replace = Read-Host "Replace"

while ($p = Read-Host "Path") {
    if (!$p) { break }
    $p = $p.Trim('"')
    if (!(Test-Path $p)) { continue }
    
    $newName = (Split-Path $p -Leaf) -replace [regex]::Escape($find), $replace
    if ($newName -eq (Split-Path $p -Leaf)) { continue }
    
    Copy-Item $p (Join-Path (Split-Path $p) $newName) -Force -ErrorAction SilentlyContinue
}