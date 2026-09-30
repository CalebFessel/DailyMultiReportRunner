@echo off
rem Dispatch 24-hour snapshot: refresh the call store, build both boards,
rem open the results. Uses the C:\Reports venv directly so the Windows
rem Store "python" placeholder is never involved.
setlocal
set PY=C:\Reports\.venv\Scripts\python.exe
cd /d C:\Reports

echo Refreshing call history...
"%PY%" jivetel_cdr_store.py incremental
if errorlevel 1 echo (call store refresh failed; hourly chart may be stale)

echo Building reports...
"%PY%" dispatch_24h_report.py --scope both %*
if errorlevel 1 (
  echo Report failed. See messages above.
  pause
  exit /b 1
)

rem Open the two newest snapshot files.
for /f "delims=" %%F in ('dir /b /o-d "C:\Reports\Reports\dispatch_snapshot_main_*.html"') do (start "" "C:\Reports\Reports\%%F" & goto :wv)
:wv
for /f "delims=" %%F in ('dir /b /o-d "C:\Reports\Reports\dispatch_snapshot_wv_*.html"') do (start "" "C:\Reports\Reports\%%F" & goto :done)
:done
echo Done. Reports are in C:\Reports\Reports
