@echo off
rem Oak: push the code in this folder to Google and update the live app (same URL).
rem First time only:  npm install -g @google/clasp   then   clasp login
cd /d "%~dp0"
if not exist .clasp.json (
  echo Missing .clasp.json. Run:  clasp clone YOUR_SCRIPT_ID   (see SETUP.md^)
  exit /b 1
)
call clasp push -f || exit /b 1
if not exist deployment.txt (
  echo Pushed. Paste your Web app deployment ID into deployment.txt to also update the live URL.
  exit /b 0
)
set /p DEP=<deployment.txt
call clasp deploy -i %DEP% -d "Update %date% %time%" && echo Oak is updated. Same link, new version.
