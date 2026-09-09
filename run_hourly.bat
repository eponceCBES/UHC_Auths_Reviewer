@echo off
REM Hourly UHC job for the scheduler machine. Pulls the latest code first, so
REM deploying a change = git push from the dev machine.
REM
REM One-time setup on the scheduler machine (see TRANSFER.md section 8):
REM   git clone <repo-url> C:\AiHub\uhc_reviewer
REM   setx REPORT_SUBSCRIPTIONS_DIR "C:\Users\<user>\OneDrive - Central Boston Elder Services, Inc\Report Subscriptions"
REM   Task Scheduler action: C:\AiHub\uhc_reviewer\run_hourly.bat   (Start in: C:\AiHub\uhc_reviewer)

setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUNBUFFERED=1

echo [%date% %time%] git pull
git pull --ff-only
if errorlevel 1 echo [warn] git pull failed - running the code already on disk

echo [%date% %time%] pipeline
python uhc_pipeline.py
if errorlevel 1 echo [warn] pipeline exited with %errorlevel%

echo [%date% %time%] note review (legacy rows / retries)
python review_notes.py

echo [%date% %time%] done
endlocal
