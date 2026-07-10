Set-Location $PSScriptRoot

Write-Host "Regenerating speaker data..." -ForegroundColor Cyan
python generate_data.py
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: speaker data generation failed." -ForegroundColor Red; exit 1 }

Write-Host "Regenerating sponsorship data..." -ForegroundColor Cyan
python generate_sponsorship.py
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: sponsorship data generation failed." -ForegroundColor Red; exit 1 }

Write-Host "`nCommitting and pushing..." -ForegroundColor Cyan
git add speakers.json sponsorship.json
git commit -m "Refresh speaker and sponsorship data from updated Excel files"
git push

Write-Host "`nDone. Live at https://accesscd.github.io/ICISA_Dashboard/" -ForegroundColor Green
