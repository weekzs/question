@echo off
setlocal
cd /d "%~dp0"

where conda >nul 2>nul
if errorlevel 1 (
  echo Conda was not found on PATH.
  echo Please install or initialize conda, then run this file again.
  pause
  exit /b 1
)

call conda activate red-movie-xlsx >nul 2>nul
if errorlevel 1 (
  echo Creating conda environment red-movie-xlsx...
  call conda env create -f environment.yml
  if errorlevel 1 (
    pause
    exit /b 1
  )
  call conda activate red-movie-xlsx
)

python -m pip install -q fastapi httpx jinja2 openpyxl pytest uvicorn
if errorlevel 1 (
  pause
  exit /b 1
)

start "" http://127.0.0.1:8765
python app.py --host 127.0.0.1 --port 8765
pause
