'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: Uptime Report
'
' AUTHOR: Jason Sowers
' DATE  : 04/08/2003
'
' COMMENT: <comment>
'
'==========================================================================

strComputer = "EARAS01"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOS in colOperatingSystems
    dtmBootup = objOS.LastBootUpTime
    dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
    dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
    Wscript.Echo dtmSystemUptime 
Next

Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & _
         Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
         & " " & Mid (dtmBootup, 9, 2) & ":" & _
         Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, _
         13, 2))
End Function

