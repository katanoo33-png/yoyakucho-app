$src = "H:\マイドライブ\じ‗自動化PJ\4.歯科衛生指導記録簿_など_システム化\厚生局提出スケジュール\yoyakucho-app"
$dst = "C:\Users\grace\Projects\yoyakucho-app"

$files = @(
  "GAS_backend.gs",
  "src\App.tsx",
  "src\App.css",
  "src\api.ts",
  "src\types.ts",
  "src\utils.ts"
)

foreach ($f in $files) {
  $s = Join-Path $src $f
  $d = Join-Path $dst $f
  if (Test-Path $s) {
    Copy-Item -Path $s -Destination $d -Force
    Write-Host "Copied: $f"
  } else {
    Write-Host "Skip (not found): $f"
  }
}

Write-Host "`nSync complete."
