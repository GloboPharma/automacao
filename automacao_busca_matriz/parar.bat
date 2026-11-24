@echo off
rem Para o app Streamlit deste projeto iniciado pelo iniciar.bat
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$py = '%~dp0.venv\Scripts\python.exe';" ^
  "$procs = Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -like ('*' + $py + '*streamlit run app.py*') };" ^
  "if ($procs) { $procs | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; Write-Host 'App parado.' } else { Write-Host 'O app nao estava rodando.' }"
timeout /t 3 >nul
