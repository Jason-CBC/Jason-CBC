On Error Resume Next

Const ForAppending = 8
Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\scripts\servers.txt", ForReading)
    
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objCSVFile = objFSO.OpenTextFile _
    ("c:\scripts\groups.csv", ForAppending, True)

Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrServiceList = Split(strNextLine , ",")
    Wscript.Echo "Server name: " & arrServiceList(0)
    	For i = 1 To Ubound(arrServiceList)
        	Wscript.Echo "Service: " & arrServiceList(i)
		Next
	strComputer = arrServiceList(0)
	'WScript.Echo strComputer
	'Set objWMIService = GetObject("winmgmts:" _
    '& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
 	' Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
'     Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount Where Domain='" & strComputer & "'")
'     
'     'objCSVFile.WriteLine "USERS"
'     'objCSVFile.WriteLine
'     
'    	'objCSVFile.WriteLine "Account,Description,Disabled,Full Name,Local Account,Lockout,Name,Password Changeable,Password Expires,Password Required"
' 
' 		'For Each objItem In colItems
'     		'objCSVFile.WriteLine objItem.Caption & "," & objItem.Description & "," & objItem.Disabled & "," & objItem.FullName & "," & objItem.LocalAccount & "," & objItem.Lockout & "," & objItem.Name & "," & objItem.PasswordChangeable & "," & objItem.PasswordExpires & "," & objItem.PasswordRequired
'     	'Next
' 		
' 		objWMIService.Close
		
		'objCSVFile.WriteLine
		'objCSVFile.WriteLine "GROUPS"
		'objCSVFile.WriteLine

		Set colGroups = GetObject("WinNT://" & strComputer & "")
		colGroups.Filter = Array("group")
		For Each objGroup In colGroups
    		'Wscript.Echo objGroup.Name 
    		For Each objUser in objGroup.Members
        		objCSVFile.WriteLine strComputer & "," & objGroup.Name & "," & objUser.Name
    		Next
		Next
		'objCSVFile.WriteLine
Loop    
objTextFile.Close
objCSVFile.Close