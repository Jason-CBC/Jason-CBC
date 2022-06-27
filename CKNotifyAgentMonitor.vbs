'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: CCA Notify Agent Service Monitor
'
' AUTHOR: Jason C. Sowers , Capital Blue Cross
' DATE  : 9/11/2006
'
' COMMENT: 
'
'==========================================================================

'--------------------------------------------------------------------------
' Variables
'--------------------------------------------------------------------------
Dim objShell, strComputer, cmdline, objWMIService, colListOfServices
Dim objService

'--------------------------------------------------------------------------
' Set Variables
'--------------------------------------------------------------------------
Set objShell = CreateObject("Wscript.Shell")
cmdline = "cmd /c blat.exe c:\scripts\blat.txt -to cbcccaitdailysupp@capbluecross.com -f vmcca19@capbluecross.com -s CK_Notify_Service_Restarted"
strComputer = "."

'--------------------------------------------------------------------------
' Main Program
'--------------------------------------------------------------------------
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colListOfServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where State = 'Stopped' and Name = " _
     & "'Alerter'")
     'CKNotifyAgent

For Each objService in colListOfServices
    objService.StartService()
    objShell.Run cmdline
Next

WScript.Quit
