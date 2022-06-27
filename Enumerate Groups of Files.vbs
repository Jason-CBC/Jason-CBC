'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 3.0
'
' NAME: 
'
' AUTHOR: CBC , CBC
' DATE  : 02/10/2004
'
' COMMENT: 
'
'==========================================================================
'On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colFiles = objWMIService. _
ExecQuery("Select * from CIM_DataFile where FileType = (Microsoft Access Application)")
For Each objFile in colFiles
Wscript.Echo objFile.Name & " -- " & objFile.FileSize & " -- " & objFile.FileType
Next


