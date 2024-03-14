Option Explicit
On Error Resume Next
Const HKCU = &H80000001
Dim strComputer, objReg, strOrigPath, strNewPath, arrKeys, strKey, strPrinter

strComputer = "."
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
 strComputer & "\root\default:StdRegProv")
strOrigPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows\SessionDefaultDevices"
strNewPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows"

objReg.EnumKey HKCU, strOrigPath, arrKeys
For Each strKey In arrKeys
    objReg.GetStringValue HKCU, strOrigPath & "\" & strKey, "Device", strPrinter
    If strPrinter <> vbNull Then
     objReg.SetStringValue HKCU, strNewPath, "Device", strPrinter
    End If
Next

Set strComputer = Nothing
Set objReg = Nothing
Set strOrigPath = Nothing
Set strNewPath = Nothing
Set arrKeys = Nothing
Set strKey = Nothing
Set strPrinter = Nothing

