Function ExecPSScript(strPSFile, strLogFile)
    ' Build powershell command
    strCmd = "powershell -NonInteractive -Command ""& { & '" & strPSFile & "'; exit $LastExitCode }"""

    ' Build command to run
    strScriptCmd = "cmd.exe /c """ & strCmd & """"

    ' Output to log file
    strScriptCmd = strScriptCmd & " > " & strLogFile

    ' Must call powershell.exe through cmd.exe in order to capture Write-Host output
    ExecPSScript = CreateObject("WScript.Shell").Run(strScriptCmd, 0, True)

End Function

Function TimeStamp()
    Dim t 
    t = Now
    timeStamp = Year(t) & "-" & _
    Right("0" & Month(t),2)  & "-" & _
    Right("0" & Day(t),2)  & "_" & _  
    Right("0" & Hour(t),2) & _
    Right("0" & Minute(t),2) & _ 
    Right("0" & Second(t),2) 
End Function

Function GetScriptFolder()
    Set objShell = CreateObject("Wscript.Shell")

    strPath = Wscript.ScriptFullName

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Set objFile = objFSO.GetFile(strPath)

    GetScriptFolder = objFSO.GetParentFolderName(objFile) 
End Function