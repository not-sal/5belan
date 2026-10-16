Unicode true
!define APPNAME "5belan"
!define APPVER "2.16.0"
!define UNINSTKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\5belan"
!define PYURL "https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"

Name "${APPNAME}"
OutFile "5belan-Setup.exe"
InstallDir "$LOCALAPPDATA\Programs\5belan"
InstallDirRegKey HKCU "Software\5belan" "InstallDir"
RequestExecutionLevel user
SetCompressor /SOLID lzma
BrandingText "5belan installer"

!include "MUI2.nsh"
!include "LogicLib.nsh"

!define MUI_ICON "5belan.ico"
!define MUI_UNICON "5belan.ico"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "Welcome to 5belan setup"
!define MUI_WELCOMEPAGE_TEXT "This wizard installs 5belan, your AI assistant, on this computer.$\r$\n$\r$\nIt needs an internet connection and about 1 GB of free disk space. The first install downloads the required components (Python and Qt) and takes around 5 to 10 minutes.$\r$\n$\r$\nClick Next to continue."
!define MUI_FINISHPAGE_TITLE "5belan is installed"
!define MUI_FINISHPAGE_TEXT "Setup finished.$\r$\n$\r$\nOn first launch open Settings and paste your DeepSeek API key.$\r$\nA shortcut was added to your Desktop and Start menu."
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_FUNCTION LaunchApp
!define MUI_FINISHPAGE_RUN_TEXT "Start 5belan now"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Var PYEXE
Function LaunchApp
  SetOutPath "$INSTDIR"
  Exec '"$INSTDIR\.venv\Scripts\pythonw.exe" "$INSTDIR\guard.py" deepseek_chat.py'
FunctionEnd
Var TMP1

; Finds a usable Python (3.10 to 3.14). Result in $PYEXE, empty when none.
Function FindPython
  StrCpy $PYEXE ""
  Delete "$TEMP\5b_py.txt"
  nsExec::ExecToStack 'cmd /c py -3 -c "import sys;ok=(3,10)<=sys.version_info[:2]<=(3,14);print(sys.executable) if ok else None;sys.exit(0 if ok else 1)" > "$TEMP\5b_py.txt" 2>nul'
  Pop $0
  ${If} $0 != "0"
    Delete "$TEMP\5b_py.txt"
    nsExec::ExecToStack 'cmd /c python -c "import sys;ok=(3,10)<=sys.version_info[:2]<=(3,14);print(sys.executable) if ok else None;sys.exit(0 if ok else 1)" > "$TEMP\5b_py.txt" 2>nul'
    Pop $0
  ${EndIf}
  ${If} $0 == "0"
    FileOpen $1 "$TEMP\5b_py.txt" r
    ${IfNot} ${Errors}
      FileRead $1 $PYEXE
      FileClose $1
    ${EndIf}
    trim:
      StrCpy $TMP1 $PYEXE 1 -1
      ${If} $TMP1 == "$\r"
        StrCpy $PYEXE $PYEXE -1
        Goto trim
      ${ElseIf} $TMP1 == "$\n"
        StrCpy $PYEXE $PYEXE -1
        Goto trim
      ${EndIf}
    ${IfNot} ${FileExists} "$PYEXE"
      StrCpy $PYEXE ""
    ${EndIf}
  ${EndIf}
  Delete "$TEMP\5b_py.txt"
FunctionEnd

Section "5belan" SecMain
  SetOutPath "$INSTDIR"
  DetailPrint "Installing to $INSTDIR"

  ; ---- 1. program files (your chats and settings are never overwritten)
  File "deepseek_chat.py"
  File "guard.py"
  File "5belan.ico"
  File "5belan.png"
  File "requirements.txt"

  ; ---- 2. Python
  DetailPrint "Looking for Python..."
  Call FindPython
  ${If} $PYEXE == ""
    DetailPrint "Python was not found. Downloading Python 3.12 (about 25 MB)..."
    nsExec::ExecToLog 'powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri ${PYURL} -OutFile \"$TEMP\5b_python.exe\""'
    Pop $0
    ${IfNot} ${FileExists} "$TEMP\5b_python.exe"
      MessageBox MB_ICONSTOP "Could not download Python. Check your internet connection and run the installer again."
      Abort
    ${EndIf}
    DetailPrint "Installing Python (this takes a minute)..."
    ExecWait '"$TEMP\5b_python.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_test=0 Include_doc=0 Shortcuts=0' $0
    Delete "$TEMP\5b_python.exe"
    StrCpy $PYEXE "$LOCALAPPDATA\Programs\Python\Python312\python.exe"
    ${IfNot} ${FileExists} "$PYEXE"
      MessageBox MB_ICONSTOP "Python installation failed (code $0). Install Python 3.12 from python.org and run this installer again."
      Abort
    ${EndIf}
  ${EndIf}
  DetailPrint "Using Python: $PYEXE"

  ; ---- 3. private environment inside the install folder
  ${IfNot} ${FileExists} "$INSTDIR\.venv\Scripts\pythonw.exe"
    DetailPrint "Creating the environment..."
    nsExec::ExecToLog '"$PYEXE" -m venv "$INSTDIR\.venv"'
    Pop $0
    ${IfNot} ${FileExists} "$INSTDIR\.venv\Scripts\pythonw.exe"
      MessageBox MB_ICONSTOP "Could not create the Python environment (code $0)."
      Abort
    ${EndIf}
  ${EndIf}

  ; ---- 4. packages
  DetailPrint "Installing packages. This is the slow step, please wait..."
  nsExec::ExecToLog '"$INSTDIR\.venv\Scripts\python.exe" -m pip install --disable-pip-version-check --progress-bar off --upgrade pip'
  Pop $0
  nsExec::ExecToLog '"$INSTDIR\.venv\Scripts\python.exe" -m pip install --disable-pip-version-check --progress-bar off PySide6-Essentials PySide6-Addons pypdf python-docx openpyxl matplotlib arabic-reshaper python-bidi'
  Pop $0
  ${If} $0 != "0"
    MessageBox MB_ICONSTOP "Package installation failed (code $0). Check your internet connection and run the installer again."
    Abort
  ${EndIf}
  DetailPrint "Installing the built-in browser engine (optional)..."
  nsExec::ExecToLog '"$INSTDIR\.venv\Scripts\python.exe" -m pip install --disable-pip-version-check --progress-bar off PySide6-WebEngine'
  Pop $0
  ${If} $0 != "0"
    DetailPrint "Note: the built-in browser could not be installed. The app works without it."
  ${EndIf}

  ; ---- 5. check
  DetailPrint "Checking the installation..."
  nsExec::ExecToLog '"$INSTDIR\.venv\Scripts\python.exe" -c "import PySide6.QtWidgets, pypdf, docx, openpyxl, matplotlib, arabic_reshaper, bidi"'
  Pop $0
  ${If} $0 != "0"
    MessageBox MB_ICONSTOP "The components were installed but could not be loaded (code $0). If the log mentions a DLL, install the Microsoft Visual C++ Redistributable from https://aka.ms/vs/17/release/vc_redist.x64.exe and run this installer again."
    Abort
  ${EndIf}

  ; ---- 6. shortcuts (no console window) and uninstaller
  SetOutPath "$INSTDIR"
  CreateShortCut "$DESKTOP\5belan.lnk" "$INSTDIR\.venv\Scripts\pythonw.exe" '"$INSTDIR\guard.py" deepseek_chat.py' "$INSTDIR\5belan.ico" 0
  CreateDirectory "$SMPROGRAMS\5belan"
  CreateShortCut "$SMPROGRAMS\5belan\5belan.lnk" "$INSTDIR\.venv\Scripts\pythonw.exe" '"$INSTDIR\guard.py" deepseek_chat.py' "$INSTDIR\5belan.ico" 0
  CreateShortCut "$SMPROGRAMS\5belan\Uninstall 5belan.lnk" "$INSTDIR\Uninstall.exe"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\5belan" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTKEY}" "DisplayName" "5belan"
  WriteRegStr HKCU "${UNINSTKEY}" "DisplayVersion" "${APPVER}"
  WriteRegStr HKCU "${UNINSTKEY}" "DisplayIcon" "$INSTDIR\5belan.ico"
  WriteRegStr HKCU "${UNINSTKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTKEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegDWORD HKCU "${UNINSTKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTKEY}" "NoRepair" 1
  DetailPrint "Done."
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\5belan.lnk"
  Delete "$SMPROGRAMS\5belan\5belan.lnk"
  Delete "$SMPROGRAMS\5belan\Uninstall 5belan.lnk"
  RMDir "$SMPROGRAMS\5belan"
  RMDir /r "$INSTDIR\.venv"
  RMDir /r "$INSTDIR\_last_good"
  RMDir /r "$INSTDIR\_last_try"
  RMDir /r "$INSTDIR\__pycache__"
  Delete "$INSTDIR\deepseek_chat.py"
  Delete "$INSTDIR\deepseek_chat.py.*"
  Delete "$INSTDIR\guard.py"
  Delete "$INSTDIR\guard.py.*"
  Delete "$INSTDIR\5belan.ico"
  Delete "$INSTDIR\5belan.png"
  Delete "$INSTDIR\error.log"
  Delete "$INSTDIR\launch.log"
  Delete "$INSTDIR\crash.log*"
  Delete "$INSTDIR\events.log*"
  Delete "$INSTDIR\run_state.json"
  Delete "$INSTDIR\resume.json"
  Delete "$INSTDIR\requirements.txt"
  Delete "$INSTDIR\update_source.txt"
  Delete "$INSTDIR\update.json"
  RMDir /r "$INSTDIR\_update_backup"
  MessageBox MB_YESNO|MB_ICONQUESTION "Also delete your chats and settings (including the saved API key)?" IDNO keepdata
    Delete "$INSTDIR\config.json"
    Delete "$INSTDIR\chats.json"
    RMDir /r "$APPDATA\5belan"
  keepdata:
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINSTKEY}"
  DeleteRegKey HKCU "Software\5belan"
SectionEnd
