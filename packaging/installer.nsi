; Arabic PDF Tools - Windows installer (per-user, no admin needed)
Unicode true
SetCompressor /SOLID lzma
SetCompressorDictSize 64

!define APPNAME     "Arabic PDF Tools"
!define APPKEY      "ArabicPDFTools"
!ifndef APPVERSION
  !define APPVERSION "0.0.0"   ; build.sh passes the real version from VERSION
!endif
!define PUBLISHER   "Beedo El-Omda"
!define UNINSTKEY   "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPKEY}"
!define CTXKEY      "Software\Classes\SystemFileAssociations\.pdf\shell"
!define PYW         "$INSTDIR\python\pythonw.exe"
!define APPPYW      "$INSTDIR\app\ArabicPDFTools.pyw"
!define ICON        "$INSTDIR\app\icon.ico"

Name "${APPNAME}"
OutFile "ArabicPDFTools-Setup-${APPVERSION}.exe"
InstallDir "$LOCALAPPDATA\Programs\${APPKEY}"
InstallDirRegKey HKCU "Software\${APPKEY}" "InstallDir"
RequestExecutionLevel user
BrandingText "${APPNAME} ${APPVERSION} - by ${PUBLISHER}"

VIProductVersion "${APPVERSION}.0"
VIAddVersionKey "ProductName" "${APPNAME}"
VIAddVersionKey "FileDescription" "${APPNAME} Setup"
VIAddVersionKey "FileVersion" "${APPVERSION}"
VIAddVersionKey "ProductVersion" "${APPVERSION}"
VIAddVersionKey "CompanyName" "${PUBLISHER}"
VIAddVersionKey "LegalCopyright" "${PUBLISHER}"

!include "MUI2.nsh"
!include "FileFunc.nsh"

!define MUI_ICON   "stage\app\icon.ico"
!define MUI_UNICON "stage\app\icon.ico"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "Install ${APPNAME}$\r$\nby ${PUBLISHER}"
!define MUI_WELCOMEPAGE_TEXT "${APPNAME} converts PDFs into editable PowerPoint and Word files with correct Arabic text (no more boxes instead of letters).$\r$\n$\r$\nEverything it needs is included - no Python or internet connection required.$\r$\n$\r$\nAfter installing, right-click any PDF to convert it, or open the app from the Start menu."
!define MUI_FINISHPAGE_TEXT "${APPNAME} is installed.$\r$\n$\r$\nRight-click any PDF to convert it, or open the app from the Start menu.$\r$\n$\r$\nEnjoy , Beedo El-Omda"
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_FUNCTION LaunchApp
!define MUI_FINISHPAGE_RUN_TEXT "Open ${APPNAME} now"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Function LaunchApp
  Exec '"${PYW}" "${APPPYW}"'
FunctionEnd

!macro AddVerb VERB LABEL FLAG
  WriteRegStr HKCU "${CTXKEY}\${VERB}" "" "${LABEL}"
  WriteRegStr HKCU "${CTXKEY}\${VERB}" "Icon" "${ICON}"
  WriteRegStr HKCU "${CTXKEY}\${VERB}\command" "" '"${PYW}" "${APPPYW}" ${FLAG} "%1"'
!macroend

Section "Install"
  SetShellVarContext current
  ; replace an older version cleanly (keeps nothing of the user's)
  RMDir /r "$INSTDIR\python"
  RMDir /r "$INSTDIR\app"

  SetOutPath "$INSTDIR\python"
  File /r "stage\python\*.*"
  SetOutPath "$INSTDIR\app"
  File /r "stage\app\*.*"
  SetOutPath "$INSTDIR"

  ; remove entries left by the earlier script-based version
  Delete "$SENDTO\PDF to PPTX (Arabic-safe).lnk"
  Delete "$SENDTO\PDF to PPTX (Arabic, editable).lnk"
  Delete "$SENDTO\PDF to Word (Arabic, editable).lnk"

  ; Start menu + desktop
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "${PYW}" '"${APPPYW}"' "${ICON}" 0 SW_SHOWNORMAL "" "Convert PDFs to editable PowerPoint / Word with correct Arabic"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Uninstall ${APPNAME}.lnk" "$INSTDIR\Uninstall.exe"
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "${PYW}" '"${APPPYW}"' "${ICON}" 0

  ; Send To (works with many files at once)
  CreateShortcut "$SENDTO\Arabic PDF - to PowerPoint.lnk" "${PYW}" '"${APPPYW}" --pptx' "${ICON}" 0
  CreateShortcut "$SENDTO\Arabic PDF - to Word.lnk" "${PYW}" '"${APPPYW}" --docx' "${ICON}" 0

  ; Right-click menu on every PDF (Windows 11: under "Show more options")
  !insertmacro AddVerb "ArabicPDF_1pptx" "Convert to PowerPoint (Arabic)" "--pptx"
  !insertmacro AddVerb "ArabicPDF_2docx" "Convert to Word (Arabic)" "--docx"
  !insertmacro AddVerb "ArabicPDF_3fix"  "Repair Arabic text in PDF" "--fix"

  ; Apps & features entry
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\${APPKEY}" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTKEY}" "DisplayName" "${APPNAME}"
  WriteRegStr HKCU "${UNINSTKEY}" "DisplayVersion" "${APPVERSION}"
  WriteRegStr HKCU "${UNINSTKEY}" "Publisher" "${PUBLISHER}"
  WriteRegStr HKCU "${UNINSTKEY}" "DisplayIcon" "${ICON}"
  WriteRegStr HKCU "${UNINSTKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTKEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "${UNINSTKEY}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteRegDWORD HKCU "${UNINSTKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTKEY}" "NoRepair" 1
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  IntFmt $0 "0x%08X" $0
  WriteRegDWORD HKCU "${UNINSTKEY}" "EstimatedSize" "$0"

  ; tell Explorer the menus changed
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, i 0, i 0)'
SectionEnd

Section "Uninstall"
  SetShellVarContext current
  DeleteRegKey HKCU "${CTXKEY}\ArabicPDF_1pptx"
  DeleteRegKey HKCU "${CTXKEY}\ArabicPDF_2docx"
  DeleteRegKey HKCU "${CTXKEY}\ArabicPDF_3fix"
  Delete "$SENDTO\Arabic PDF - to PowerPoint.lnk"
  Delete "$SENDTO\Arabic PDF - to Word.lnk"
  Delete "$DESKTOP\${APPNAME}.lnk"
  RMDir /r "$SMPROGRAMS\${APPNAME}"
  RMDir /r "$INSTDIR\python"
  RMDir /r "$INSTDIR\app"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINSTKEY}"
  DeleteRegKey HKCU "Software\${APPKEY}"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, i 0, i 0)'
SectionEnd
