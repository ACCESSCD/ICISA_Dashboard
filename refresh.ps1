Set-Location $PSScriptRoot

Write-Host "Regenerating speaker data..." -ForegroundColor Cyan
C:\Users\carol\PycharmProjects\EA26_AnaesthCareEurope\venv\Scripts\python.exe generate_data.py
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: data generation failed." -ForegroundColor Red; exit 1 }

Write-Host "`nCommitting and pushing..." -ForegroundColor Cyan
git add speakers.json
git commit -m "Refresh speaker data from updated Excel files"
git push

Write-Host "`nDone. Live at https://accesscd.github.io/ICISA_Dashboard/" -ForegroundColor Green
