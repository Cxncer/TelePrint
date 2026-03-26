' Start Noble Printer Bot silently
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Set working directory to your bot folder
strPath = "F:\OneDrive\OneDrive - Noble\M. Sopxnha's files - Personal\AI - Workflow\VSCode\NoblePrinter_Bot"
objShell.CurrentDirectory = strPath

' Use pythonw.exe from the virtual environment (no console window)
' Note: pythonw.exe is in the same folder as python.exe
strPython = strPath & "\venv\Scripts\pythonw.exe"
strScript = "bot_tray_launcher.py"

' Check if pythonw.exe exists (optional - remove if you don't want the message)
If objFSO.FileExists(strPython) Then
    ' Run the bot - 0 means hidden window
    objShell.Run """" & strPython & """ """ & strScript & """", 0, False
Else
    MsgBox "Error: pythonw.exe not found at: " & strPython
End If

Set objShell = Nothing
Set objFSO = Nothing