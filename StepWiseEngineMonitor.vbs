'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME: StepWise Service AutoStart
'
' AUTHOR: Jason C. Sowers , CBC
' DATE  : 5/19/2008
'
' COMMENT: 
'
'==========================================================================

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery _
 ("Select * from Win32_Service Where DisplayName = 'Print Spooler' and State = 'Stopped' and StartMode = " _
     & "'Auto'")
For Each objService in colListOfServices
    objService.StartService()
Next
