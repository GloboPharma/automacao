@echo off
rem Inicia o app Streamlit em segundo plano (sem janela) em http://localhost:8501
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$py = '%~dp0.venv\Scripts\python.exe';" ^
  "$rodando = Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -like ('*' + $py + '*streamlit run app.py*') };" ^
  "if ($rodando) { Write-Host 'O app ja esta rodando em http://localhost:8501'; exit };" ^
  "Start-Process -FilePath $py -ArgumentList '-m streamlit run app.py --server.address localhost --server.port 8501 --server.headless true' -WorkingDirectory '%~dp0.' -WindowStyle Hidden -RedirectStandardOutput '%~dp0streamlit.log' -RedirectStandardError '%~dp0streamlit.err.log';" ^
  "Write-Host 'App iniciado em http://localhost:8501'"
timeout /t 3 >nul
