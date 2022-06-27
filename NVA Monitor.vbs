'**************************************************
'*
'* Monitor Folders for Creation of Files
'*
'* NAME: NVA Monitor.vbs
'*
'* AUTHOR: Jason C. Sowers
'* DATE  : 02/19/2003
'*
'* COMMENT: Used to monitor for new NVA Files
'*
'****************************************************

Option Explicit

'****************************************************
'Variables Section
'****************************************************
Dim CurDate, dtmTargetDate, dtmConvertedDate
Dim TargetDate, strComputer, objFSO
Dim objWMIService, colFolders, objFolder

'****************************************************
'Set Objects
'****************************************************
Set objFSO = CreateObject ("Scripting.FileSystemObject")
Set dtmTargetDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

'****************************************************
'Initialize Variables Section
'****************************************************
Const ForReading = 1
Const LOCAL_TIME = True
CurDate = Date()
TargetDate = CurDate - 1

'****************************************************
'Process Section
'****************************************************
WScript.Echo(CurDate)
WScript.Echo(TargetDate)

dtmTargetDate.SetVarDate TargetDate, LOCAL_TIME
strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory Where " _
        & "CreationDate > '" & dtmTargetDate & "'")
For each objFolder in colFolders
    dtmConvertedDate.Value = objFolder.CreationDate 
    Wscript.Echo objFolder.Name & VbTab & _
        dtmConvertedDate.GetVarDate(LOCAL_TIME)
Next

'****************************************************
'Subroutine Section
'****************************************************