@echo off
setlocal EnableDelayedExpansion
title Arabic PDF Tools (PowerPoint + Word) - by Beedo El-Omda
echo.
echo  Installing Arabic PDF Tools 1.3: "PDF to PPTX" and "PDF to Word"
echo  by Beedo El-Omda
echo.

set "APPDIR=%LOCALAPPDATA%\PDF2PPTX-Arabic"
set "SENDTO=%APPDATA%\Microsoft\Windows\SendTo"

rem ---- 1. Find or install Python ----------------------------------------
set "PY="
for %%P in ("%LOCALAPPDATA%\Programs\Python\Python313\python.exe" "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%LOCALAPPDATA%\Programs\Python\Python311\python.exe") do (
  if not defined PY if exist %%P set "PY=%%~P"
)
if not defined PY (
  for /f "delims=" %%i in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%i"
)
if not defined PY (
  echo  [1/3] Python not found - installing Python 3.12 with winget...
  winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PY (
  echo.
  echo  Could not install Python automatically.
  echo  Install "Python 3.12" from the Microsoft Store, then run Install.cmd again.
  pause & exit /b 1
)
echo  [1/3] Using Python: !PY!

rem ---- 2. Libraries --------------------------------------------------------
echo  [2/3] Installing libraries (PyMuPDF, python-pptx, python-docx, pdf2docx, pikepdf, Pillow, fontTools)...
rem exact versions this tool was tested with (core libraries only, no AI extras)
"!PY!" -m pip install --user --quiet --no-deps pymupdf==2.0 python-pptx==1.0.2 python-docx==1.2.0 pdf2docx==0.5.13 pikepdf==10.16.0 pillow==12.3.0 fonttools==4.66.1 numpy==2.5.3 opencv-python-headless==5.0.0.93 lxml==6.1.3 typing_extensions==4.16.0 XlsxWriter==3.2.9 Deprecated==3.0.1 wrapt==2.5.2 packaging==26.3 fire==0.7.1 termcolor==3.3.0
if errorlevel 1 ( echo  Library install failed - check internet connection. & pause & exit /b 1 )

rem ---- 3. Copy scripts and add Send To entry -------------------------------
echo  [3/3] Adding to right-click "Send to" menu...
if not exist "%APPDIR%" mkdir "%APPDIR%"
copy /y "%~dp0pdf2pptx.py" "%APPDIR%\pdf2pptx.py" >nul
copy /y "%~dp0fix_arabic_pdf.py" "%APPDIR%\fix_arabic_pdf.py" >nul
copy /y "%~dp0pdf2docx_arabic.py" "%APPDIR%\pdf2docx_arabic.py" >nul
copy /y "%~dp0docx_flow.py" "%APPDIR%\docx_flow.py" >nul
copy /y "%~dp0Uninstall.cmd" "%APPDIR%\Uninstall.cmd" >nul
del "%SENDTO%\PDF to PPTX (Arabic-safe).lnk" 2>nul

set "PYW=!PY:python.exe=pythonw.exe!"
if not exist "!PYW!" set "PYW=!PY!"
"!PY!" "%APPDIR%\pdf2pptx.py" --install-sendto "!PYW!"
"!PY!" "%APPDIR%\pdf2docx_arabic.py" --install-sendto "!PYW!"

if exist "%SENDTO%\PDF to Word (Arabic, editable).lnk" (
  echo.
  echo  Done. Right-click any PDF ^> Send to ^>
  echo     "PDF to PPTX (Arabic, editable)"  - PowerPoint
  echo     "PDF to Word (Arabic, editable)"  - Word
  echo  The new file is saved next to the PDF.
  echo.
  echo  Enjoy , Beedo El-Omda
) else (
  echo  Could not create the Send To shortcuts.
)
echo.
pause
