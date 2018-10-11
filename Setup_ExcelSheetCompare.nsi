; example1.nsi
;
; This script is perhaps one of the simplest NSIs you can make. All of the
; optional settings are left to their default settings. The installer simply 
; prompts the user asking them where to install, and drops a copy of example1.nsi
; there. 

;--------------------------------

Unicode true
XPStyle on
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Japanese-MeiryoUI.nlf"

!define APP "ExcelSheetCompare"
!finalize 'MySign "%1"'

; The name of the installer
Name "${APP}"

; The file to write
OutFile "Setup_${APP}.exe"

; The default installation directory
InstallDir "$APPDATA\Microsoft\AddIns"

; Request application privileges for Windows Vista
RequestExecutionLevel user

;--------------------------------

; Pages

Page directory
Page components
Page instfiles

;--------------------------------

; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File "エクセルファイルを比較.xlam"
  
SectionEnd ; end the section

Section "エクセルファイルを比較.xlam 活性化"
  SetOutPath "$APPDATA\${APP}"
  
  File "InstAddin.vbs"
  ExecWait 'cscript.exe InstAddin.vbs' $0
  DetailPrint "結果: $0"
SectionEnd
