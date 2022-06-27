'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: Folder Creation and Security Settings
'
' AUTHOR: Jason C. Sowers , Capital Blue Cross
' DATE  : 7/18/2006
'
' COMMENT: Requires XCACLS.VBS in C:\Windows, WSH 5.6 or higher.
'
'==========================================================================
Dim FolderName, objFSO, objFolder, strDirectory, Initials, Username
Dim MyDate, MyTime

MyDate = Date
MyTime = Time
WScript.Echo "Begin Process " & MyDate & " " & MyTime

' Asks for input from the user 
FolderName = InputBox("FolderName (Initials):" , "Folder Creation")
UserName = InputBox("UserName (Domain\Username):" , "UserName Insertion")

' Sets the path for the folder and the folder name 
strDirectory = "P:\Shared\PilotGroups\Pilot_1_May06\" & FolderName 'EDIT AS NEEDED

' Create FileSystemObject. So we can apply .createFolder method 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

' Checks if folder already exists 
If objFSO.FolderExists(strDirectory) Then 
	Set objFolder = objFSO.GetFolder(strDirectory) 
	WScript.Echo FolderName & " already exist, no permissions altered " 
Else 
	' Folder Creation
	WScript.Echo "Creating Folder" 
	Set objFolder = objFSO.CreateFolder(strDirectory) 
	
	'Set Permissions 
	Set Shell = createobject("wscript.shell")
	 
	Call Inherit (strDirectory)
	WScript.Sleep 5000
	Call Permit (strDirectory,Username)
	WScript.Sleep 5000
	Call Revoke (strDirectory)
	WScript.Sleep 5000
	Call Owner (strDirectory)
End If

MyDate = Date
MyTime = Time
WScript.Echo "End Process " & MyDate & " " & MyTime

Function Inherit (strDirectory)
	WScript.Echo "Removing Inheritance"
	Shell.Run "cscript c:\windows\xcacls.vbs " & strDirectory &" /i copy",WINDOWNORMAL,WaitOnReturn
	WScript.Echo strDirectory
End Function

Function Permit (strDirectory, Username)
	WScript.Echo "Applying Username and Permissions"
	Shell.Run "cscript c:\windows\xcacls.vbs " & strDirectory &" /e /g "& Username &":M " & "/f /s /t",WINDOWNORMAL,WaitOnReturn
End Function

Function Revoke (strDirectory)
	WScript.Echo "Revoking BCNET\Domain Users"
	Shell.Run "cscript c:\windows\xcacls.vbs " & strDirectory &" /e /r " & CHR(34) & "BCNET\Domain Users" & Chr(34),WINDOWNORMAL,WaitOnReturn
End Function

Function Owner (strDirectory)
	WScript.Echo "Applying Ownership Rights"
	Shell.Run "cscript c:\windows\xcacls.vbs " & strDirectory &" /e /o " & "Administrators",WINDOWNORMAL,WaitOnReturn
End Function