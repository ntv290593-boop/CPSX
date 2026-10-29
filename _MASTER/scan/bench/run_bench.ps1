param([string]$bench,[string[]]$models,[int]$par=4)
$agy=(Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Filter agy.exe -Recurse | Select-Object -First 1).FullName
$models=@($models -split ","); $crops=Get-ChildItem "$bench\crops" -Filter *.png
$jobsList=@(); foreach($m in $models){ foreach($c in $crops){ $out="$bench\out\$m\$($c.BaseName).txt"; if(-not (Test-Path $out) -or (Get-Item $out).Length -lt 2){ $jobsList+=,@($m,$c.FullName,$out) } } }
"to run: $($jobsList.Count)"
$running=@()
foreach($j in $jobsList){
  while(($running | Where-Object State -eq Running).Count -ge $par){ Start-Sleep 3 }
  New-Item -ItemType Directory -Force (Split-Path $j[2]) | Out-Null
  $running+=Start-Job -ScriptBlock { param($agy,$m,$img,$out,$cwd)
    Set-Location $cwd
    $prompt="Do NOT run shell commands; use your file-viewing tool directly on this absolute path: $img . It is a crop of a hotel joinery detail drawing (dimensions in millimetres, drawn as vector lines). List EVERY dimension number printed along dimension lines in the crop, including numbers on vertical dimension lines, one entry per printed number, in reading order. Do NOT include detail numbers inside circles, scale text, level markers (FFL/FCL) or note text. If a number is unreadable write ""?"". Never guess. Output ONLY a JSON array of strings, nothing else."
    & $agy --model $m --print $prompt --print-timeout 280s 2>$null | Out-File $out -Encoding utf8
  } -ArgumentList $agy,$j[0],$j[1],$j[2],$bench
}
$running | Wait-Job | Out-Null
"done"

