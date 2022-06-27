'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: Report Service Offline - Multiple Failures
'
' AUTHOR: Jason C. Sowers , Capital Blue Cross
' DATE  : 9/15/2006
'
' COMMENT: 
'
'==========================================================================

'--------------------------------------------------------------------------
' Variables
'--------------------------------------------------------------------------
Dim objShell, cmdline, objService

'--------------------------------------------------------------------------
' Set Variables
'--------------------------------------------------------------------------
Set objShell = CreateObject("Wscript.Shell")
cmdline = "cmd /c blat.exe c:\scripts\blat2.txt -to cbcccaitsupport@capbluecross.com -f eacks05@capbluecross.com -s CK_Notify_Service_Failed"

'--------------------------------------------------------------------------
' Main Program
'--------------------------------------------------------------------------
objShell.Run cmdline

WScript.Quit