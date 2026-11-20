@echo off
del "%APPDATA%\Microsoft\Windows\SendTo\PDF to PPTX (Arabic, editable).lnk" 2>nul
del "%APPDATA%\Microsoft\Windows\SendTo\PDF to PPTX (Arabic-safe).lnk" 2>nul
del "%APPDATA%\Microsoft\Windows\SendTo\PDF to Word (Arabic, editable).lnk" 2>nul
del "%LOCALAPPDATA%\PDF2PPTX-Arabic\pdf2pptx.py" 2>nul
del "%LOCALAPPDATA%\PDF2PPTX-Arabic\pdf2docx_arabic.py" 2>nul
del "%LOCALAPPDATA%\PDF2PPTX-Arabic\docx_flow.py" 2>nul
del "%LOCALAPPDATA%\PDF2PPTX-Arabic\fix_arabic_pdf.py" 2>nul
del "%LOCALAPPDATA%\PDF2PPTX-Arabic\arabic_common_words.py" 2>nul
echo Removed "PDF to PPTX" and "PDF to Word" (Arabic) from the Send to menu.
echo (Python and its libraries were left installed.)
pause
