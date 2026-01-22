$find = Read-Host "Find"
$repls = (Read-Host "Replace (comma sep)").Split(',') | % {$_.Trim()} | ? {$_}

while ($p = Read-Host "Path") {
    if (!$p) {break}
    $p = $p.Trim('"')
    if (!(Test-Path $p)) {continue}
    
    $dir  = Split-Path $p
    $name = Split-Path $p -Leaf
    $newbase = $name -replace [regex]::Escape($find),''
    
    if ($newbase -eq $name) {continue}
    
    foreach ($r in $repls) {
        $newname = $newbase -replace '(\.[^.]+)?$',$r+'$1'
        Copy-Item $p (Join-Path $dir $newname) -Force -EA SilentlyContinue
    }
}
