Unicode True
InstallDir "$PROGRAMFILES\pyOle"
RequestExecutionLevel admin

Name "pyOle installer"
OutFile "installer.exe"

Page directory
Page instfiles

LoadLanguageFile "${NSISDIR}\Contrib\Language files\English.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Spanish.nlf"

Section "Instalar"

    SetOutPath "$INSTDIR"

    File /r dist\*.*

    ExecWait '"$INSTDIR\pyOle.exe" --register --silent'
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "DisplayName" "pyOle"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "DisplayIcon" "$INSTDIR\pyOle.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "Publisher" "TuNombre"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "DisplayVersion" "1.0"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp" "NoRepair" 1    

SectionEnd

Section "Uninstall"
    MessageBox MB_YESNO "Are you sure you want to uninstall?" IDNO NoUninstall

    ExecWait '"$INSTDIR\pyOle.exe" --unregister --silent'

    Delete "$INSTDIR\*.*"
    RMDir "$INSTDIR"

    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyApp"
    
    MessageBox MB_OK "The program has been uninstalled."

    NoUninstall:
SectionEnd