'*** LastLogoff.vbs Script Begins
'Alan Kaplan for VA Salisbury
'www.akaplan.com/tools.html
'Last revision 2/21/01
'This program gets the last logoff and other information for users of an Accounts Domain
'whose names contain name of resource domain.  Gets information from
'Accounts Domain PDC and local BDC, and dumps information into a spreadsheet

option explicit
Dim usr, user, userobj,wshshell
dim key,key2,major,minor,ver
Dim regEx, retVal,result,server,serverlist,slist,slist2
Dim cont,dnameobj, domain, searchfor, quote, command
dim logfile,comma,newline,message,cdname,title,pwage,lldate
dim fso,appendout,ofile

'message format helpers..
newline = ""
comma = ","
quote = chr(34)

Set WshShell = WScript.CreateObject("WScript.Shell")

'Test for ADSI and current WSH
SysTest

'get resource domain name as default search string within logon domain
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\Currentversion\Winlogon\CachePrimaryDomain"
CDname = WshShell.RegRead (key)
'on error resume next

'*********** EDIT HERE ***************************
'*** List DCs to query, each in quotes
'*** Limit list to those your users are most likely to
'*** authenticate.
'*** If you have more than 5, must add s# to serverlist array
'*** following list.
'*** Default is the resource domain of current PC and Accounts Domain PDC
'*** Not the most elegant way of doing this, but most clear....

dim s1,s2,s3,s4,s5
s1 = "EADSA10"		'Domain Controller
s2 = "EADSA11"		'Domain Controller
s3 = "WSDSA10"		'Domain Controller
s4 = "LVDSA10"		'Domain Controller
s5 = "VWDSA10"		'Domain Controller

serverlist = Array(s1,s2,s3,s4,s5)

'*** Name of log file. UNC path is okay. Leave extension as CSV
'*** to have it opened by Excel.  Change to TXT
'*** if Excel not loaded on workstation running script
logfile = "c:\LastServerLogon.csv"	
'*********** EDIT ENDS ***************************

' Define dialog box variables.
message = "Please enter search pattern for user name"
title = "Limit Search"

'get domain name, domain default
searchfor = InputBox(message, title, CDName)
' Evaluate the user input.
If searchfor = "" Then    ' Canceled by the user
    WScript.quit
End If


'Make sure running Cscript. Mostly G. Born code...
If (Not IsCScript()) Then 	'If not CScript, re-run with cscript...
	WshShell.Run "CScript.exe " & WScript.ScriptFullName
    WScript.Quit            	'...and stop running as WScript
End If

'setup log
'logfile is set to append.
Const ForAppend = 8
set fso = CreateObject("Scripting.FileSystemObject")

'delete old log
	if (fso.FileExists(logfile))then
	set oFile = fso.GetFile (logfile)
	oFile.Delete
end if

'create header for log
set AppendOut = fso.OpenTextFile(logfile, ForAppend, True)
appendout.writeline "User Name,Last Login,PWAgeDays,Full Name,Disabled,Description,Server"


'Main Routine
on error goto 0
For Each Server in Serverlist 
	If server <> "" then
		Notify
		dnameobj = "WinNT://" & Server		'attach to server
		Set cont = GetObject(dnameobj)
		cont.Filter = Array("user")		'search for users matching searchfor
			For Each usr In cont
			MoreTests SearchFor,usr.name	'Run MoreTests function
			Next
	End If
Next


'Done.  open file with associated app, Excel if installed.
command = "cmd.exe /c start " & logfile  
WshShell.Run command
wscript.quit				'The end

'Functions and Subroutines


Sub Notify
title = "Getting User Information"
message = "Querying " & Server & ", please wait ...."
WshShell.Popup message,3,title,64
end sub

Function MoreTests(patrn, strng)
'checks for pattern in username
on error resume next
  Set regEx = New RegExp        ' Create regular expression.
  regEx.Pattern = patrn         ' Set pattern.
  regEx.IgnoreCase = True	' Set case sensitivity.
  retVal = regEx.Test(strng)	' Execute the search test.
  If retVal Then		' if match, setup to write
	'formatting routines
	Set UserObj = GetObject(dnameobj&"/"& usr.name)
	'Convert PW Age from seconds to days
	PWAge = Round((UserObj.PasswordAge/86400),0)
	'Strip Time from Last Login
	LLDate = Split(UserObj.LastLogin,chr(32),-1,1)
	'setup output
	message = UserObj.Name 
	message = message & ","
	message = message & LLDate(0)
	message = message & ","	
	message = message & PWAge
	message = message & ","	
	message = message & quote & UserObj.FullName & quote
	message = message & ","
	message = message & UserObj.AccountDisabled
	message = message & ","
	message = message & quote & UserObj.Description & quote
	message = message & ","
	message = message & server 
	echoandlog message	'write to log
	End If
'fail on errors
on error goto 0
'clear values
message = ""
PWAge = ""
LLDate = ""
End Function

Function IsCScript()
'Function to test whether host CScript.exe.
    If (InStr(UCase(WScript.FullName), "CSCRIPT") <> 0) Then
        IsCScript = True
    Else
        IsCScript = False
    End If
End Function


Sub EchoAndLog (message)
'Subroutine to echo to screen and write to log
	Wscript.Echo message
'appendout is defined as part of log setup
	AppendOut.WriteLine message
End Sub  

Sub SysTest()
on error resume next
' WSH version tested
Major = (ScriptEngineMinorVersion())
Minor = (ScriptEngineMinorVersion())/10
Ver = major + minor
'Need version 5.5
	If err.number or ver < 5.5 then
	message = "You have must load Version 5.5 (or later) of Windows Script Host" &vbCrLf
End If

'Test for ADSI
err.clear
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{E92B03AB-B707-11d2-9CBD-0000F87A369E}\version"
key2 = WshShell.RegRead (key)
if err <> 0 then
  	message = message & "ADSI must be installed on local workstation to continue" & vbCrLf
WshShell.Popup message,0,"Workstation Setup Error",vbInformation
	WScript.Quit
	End if
End Sub
 
'******Script Ends