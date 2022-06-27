Dim sFile_App
Dim sFile_Sec
Dim sFile_Sys
Dim sDate
Dim Machine
Dim Hr
Dim Minu
Dim MM
Dim DD
Dim YYYY


Set Machine = GetObject("winmgmts:").InstancesOf ("Win32_ComputerSystem")

'
' The files will be in the format <log>_<MachineName>_<date>.EVT, where:
'		<log> is the log type
'		<MachineName> is the Systemname of the target system
'		<date> is the date that the log was Archived (in mddyyyy format)
'
     Hr = Hour(Now)
     Minu = Minute(Now)
     MM = Month(Now)
     DD = Day(Now)
     YYYY= Year(Now)
'    sDate = FormatDateTime(Now(),2) & "_" & FormatDateTime(Now(),3) & ".evt"
'    sdate = FormatDateTime(Now(),2)
'    sdate = left(sdate,2) & "-" & mid(sdate,4,2) & "-" & right(sdate,4)
     sDate = MM & DD & YYYY & "_" & Hr & Minu & ".evt"

    For each System in Machine
    sFile_App = "APP_" & System.Name & "_" & sDate
    sFile_Sec = "SEC_" & System.Name & "_" & sDate
    sFile_Sys = "SYS_" & System.Name & "_" & sDate
    Next
' 
' WMI Script - Backup event log (VBScript)
'
' NOTE:  This script only applies to NT-based systems since Win9x does support event logs
'
Set LogFileSet = GetObject("winmgmts:{(Backup,Security)}").ExecQuery("select * from Win32_NTEventLogFile where LogfileName='Application'")

for each Logfile in LogFileSet
	RetVal = LogFile.BackupEventlog("\\easrv05\w2kinstalls\" & sFile_App)' & ".evt")
	RetVal = LogFile.ClearEventlog()
Next

Set LogFileSet = GetObject("winmgmts:{(Backup,Security)}").ExecQuery("select * from Win32_NTEventLogFile where LogfileName='Security'")

for each Logfile in LogFileSet
	RetVal = LogFile.BackupEventlog("\\easrv05\w2kinstalls\" & sFile_Sec)' & ".evt")
	RetVal = LogFile.ClearEventlog()
Next

Set LogFileSet = GetObject("winmgmts:{(Backup,Security)}").ExecQuery("select * from Win32_NTEventLogFile where LogfileName='System'")

for each Logfile in LogFileSet
	RetVal = LogFile.BackupEventlog("\\easrv05\w2kinstalls\" & sFile_Sys)' & ".evt")
	RetVal = LogFile.ClearEventlog()
Next