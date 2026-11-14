; Arabic PDF Tools - portable launcher (no installation, no admin)
; Starts the bundled Python app that sits next to this exe.
; Drag PDFs onto the exe to open them in the app.
Unicode true
SilentInstall silent
RequestExecutionLevel user
SetCompressor /SOLID lzma

!define APPNAME    "Arabic PDF Tools"
!ifndef APPVERSION
  !define APPVERSION "0.0.0"   ; build.sh passes the real version from VERSION
!endif
!define AUTHOR     "Beedo El-Omda"

Name "${APPNAME}"
Caption "${APPNAME}"
OutFile "ArabicPDFTools.exe"
Icon "stage\app\icon.ico"

VIProductVersion "${APPVERSION}.0"
VIAddVersionKey "ProductName" "${APPNAME} (Portable)"
VIAddVersionKey "FileDescription" "${APPNAME} - by ${AUTHOR}"
VIAddVersionKey "FileVersion" "${APPVERSION}"
VIAddVersionKey "ProductVersion" "${APPVERSION}"
VIAddVersionKey "CompanyName" "${AUTHOR}"
VIAddVersionKey "LegalCopyright" "${AUTHOR}"

!include "FileFunc.nsh"

Section
  ${GetParameters} $0
  IfFileExists "$EXEDIR\data\python\pythonw.exe" +3 0
    MessageBox MB_ICONSTOP "The 'data' folder is missing.$\r$\n$\r$\nKeep ArabicPDFTools.exe together with its 'data' folder (unzip the whole package)."
    Quit
  Exec '"$EXEDIR\data\python\pythonw.exe" "$EXEDIR\data\app\ArabicPDFTools.pyw" $0'
SectionEnd
