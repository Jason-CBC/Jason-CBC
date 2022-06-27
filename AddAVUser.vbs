	' ACCESS_MODE
Const DENY_ACCESS = 		3
Const GRANT_ACCESS = 		1
Const REVOKE_ACCESS = 		4
Const SET_ACCESS = 			2
Const SET_AUDIT_FAILURE = 	6
Const SET_AUDIT_SUCCESS = 	5

' ACTIONS
Const ACTN_ADDACE = 		1
Const ACTN_CLEARDACL = 		16
Const ACTN_CLEARSACL = 		32
Const ACTN_COPYDOMAIN = 	1024
Const ACTN_COPYTRUSTEE = 	1024
Const ACTN_DOMAIN = 		8192
Const ACTN_LIST = 			2
Const ACTN_REMOVEDOMAIN = 	512
Const ACTN_REMOVETRUSTEE = 	512
Const ACTN_REPLACEDOMAIN = 	256
Const ACTN_REPLACETRUSTEE = 256
Const ACTN_RESETCHILDPERMS = 128
Const ACTN_RESTORE = 		2048
Const ACTN_SETGROUP = 		8
Const ACTN_SETINHFROMPAR = 	64
Const ACTN_SETOWNER = 		4
Const ACTN_TRUSTEE = 		4096

' INHERITANCE
Const INHPARCOPY = 			2	' p_c:	Protected, ACEs from parent are copied
Const INHPARNOCHANGE = 		0	' nc:	No Change
Const INHPARNOCOPY = 		4	' p_nc:	Protected, ACEs from parent are NOT copied
Const INHPARYES = 			1	' np:	Not Protected, inherits from parent

' LISTFORMATS
Const LIST_CSV = 			1
Const LIST_SDDL = 			0
Const LIST_TAB = 			2

' LISTNAMES
Const LIST_NAME = 			1
Const LIST_NAME_SID = 		3
Const LIST_SID = 			2

' RECURSION
Const RECURSE_CONT = 		2
Const RECURSE_CONT_OBJ = 	6
Const RECURSE_NO = 			1
Const RECURSE_OBJ = 		4

' RETCODES
Const RTN_ERR_ADD_ACE = 	32
Const RTN_ERR_CONVERT_SD = 	27
Const RTN_ERR_COPY_ACL = 	31
Const RTN_ERR_CREATE_SD = 	45
Const RTN_ERR_DEL_ACE = 	30
Const RTN_ERR_DIS_PRIV = 	13
Const RTN_ERR_EN_PRIV = 	12
Const RTN_ERR_FINDFILE = 	16
Const RTN_ERR_GENERAL =		2
Const RTN_ERR_GET_SD_CONTROL = 17
Const RTN_ERR_GETSECINFO = 	5
Const RTN_ERR_IGNORED = 	44
Const RTN_ERR_INTERNAL = 	18
Const RTN_ERR_INV_DIR_PERMS = 7
Const RTN_ERR_INV_DOMAIN = 	43
Const RTN_ERR_INV_PRN_PERMS = 8
Const RTN_ERR_INV_REG_PERMS = 9
Const RTN_ERR_INV_SHR_PERMS = 11
Const RTN_ERR_INV_SVC_PERMS = 10
Const RTN_ERR_INVALID_SD = 	38
Const RTN_ERR_LIST_ACL = 	28
Const RTN_ERR_LIST_FAIL = 	15
Const RTN_ERR_LIST_OPTIONS = 26
Const RTN_ERR_LOOKUP_SID = 	6
Const RTN_ERR_LOOP_ACL = 	29
Const RTN_ERR_NO_LOGFILE =	33
Const RTN_ERR_NO_NOTIFY = 	14
Const RTN_ERR_OBJECT_NOT_SET = 4
Const RTN_ERR_OPEN_LOGFILE = 34
Const RTN_ERR_OS_NOT_SUPPORTED = 37
Const RTN_ERR_OUT_OF_MEMORY = 46
Const RTN_ERR_PARAMS = 		3
Const RTN_ERR_PREPARE = 	24
Const RTN_ERR_READ_LOGFILE = 35
Const RTN_ERR_REG_CONNECT = 21
Const RTN_ERR_REG_ENUM = 	23
Const RTN_ERR_REG_OPEN = 	22
Const RTN_ERR_REG_PATH = 	20
Const RTN_ERR_SET_SD_DACL = 39
Const RTN_ERR_SET_SD_GROUP = 42
Const RTN_ERR_SET_SD_OWNER = 41
Const RTN_ERR_SET_SD_SACL = 40
Const RTN_ERR_SETENTRIESINACL = 19
Const RTN_ERR_SETSECINFO = 	25
Const RTN_ERR_WRITE_LOGFILE = 36
Const RTN_OK = 0
Const RTN_USAGE = 1

' SDINFO
Const ACL_DACL = 			1
Const ACL_SACL = 			2
Const SD_GROUP = 			8
Const SD_OWNER = 			4

' SE_OBJECT_TYPE
Const SE_FILE_OBJECT = 		1
Const SE_LMSHARE = 			5
Const SE_PRINTER = 			3
Const SE_REGISTRY_KEY = 	4
Const SE_SERVICE = 			2

' userAccountControl Flags
'
Const ADS_UF_SCRIPT = &h00000001					' The logon script is executed.
Const ADS_UF_ACCOUNTDISABLE = &h2					' The user account is disabled.
Const ADS_UF_HOMEDIR_REQUIRED = &h00000008			' The home directory is required.
Const ADS_UF_LOCKOUT = &h00000010					' The account is currently locked out.
Const ADS_UF_PASSWD_NOTREQD = &h00000020			' No password is required.
Const ADS_UF_PASSWD_CANT_CHANGE = &h40				' The user cannot change the password.
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &h00000080	' The user can send an encrypted password.
Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = &h00000100	' This is an account for users whose primary account is in another domain. This account provides user access to this domain, but not to any domain that trusts this domain. Also known as a local user account.
Const ADS_UF_NORMAL_ACCOUNT = &h00000200			' This is a default account type that represents a typical user.
Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &h00000800	' This is a permit to trust account for a system domain that trusts other domains.
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &h00001000	' This is a computer account for a computer that is a member of this domain.
Const ADS_UF_SERVER_TRUST_ACCOUNT = &h00002000		' This is a computer account for a system backup domain controller that is a member of this domain.
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000			' The password for this account will never expire.
Const ADS_UF_MNS_LOGON_ACCOUNT = &h00020000			' This is an MNS logon account.
Const ADS_UF_SMARTCARD_REQUIRED = &h00040000		' The user must log on using a smart card.
Const ADS_UF_TRUSTED_FOR_DELEGATION = &h00080000	' The service account (user or computer account), under which a service runs, is trusted for Kerberos delegation. Any such service can impersonate a client requesting the service.
Const ADS_UF_NOT_DELEGATED = &h00100000				' The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation.
Const ADS_UF_USE_DES_KEY_ONLY = &h00200000			' Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys.
Const ADS_UF_DONT_REQUIRE_PREAUTH = &h00400000		' This account does not require Kerberos pre-authentication for logon.
Const ADS_UF_PASSWORD_EXPIRED = &h00800000			' The user password has expired. This flag is created by the system using data from the Pwd-Last-Set attribute and the domain policy.
Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = &h01000000


Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000
Const ADS_PROPERTY_APPEND = 3 

Const RootDrv = "D:"

Public TPID
Public password

set WshShell = WScript.CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
Set ScriptFile = objFS.GetFile(WScript.ScriptFullname)
LogFile = WScript.ScriptName
LogName = Split(LogFile,".")
ScriptPath = ScriptFile.ParentFolder
Set objLogFile = objFS.OpenTextFile(ScriptPath & "\" & LogName(0) & ".log", 8, True)
If AltLog = 1 Then
	Set objAltLogFile = objFS.OpenTextFile(ScriptPath & "\" & AltLogFile, 2, True)
End If

main

Sub Main()

	If objArgs.Count = 0 Then
		Wscript.Echo "Adds Trading Partner ID specified by their four digit ID to"
		WScript.Echo "EXTNET, creates an associated group and adds the user to this"
		WScript.Echo "group and the EXTNET_FTP.Group.  Creates the folder and permissions"
		WScript.Echo
		WScript.Echo "Usage: CScript CreateTPUser.vbs [TPID] /p[password]"
		WScript.Echo
		WScript.echo "TPID		The Four (4) Digit Trading Partner ID"
		WScript.Echo "password	The password for the account"
		WScript.Echo
	Else

		OutputText True, "Script: CreateTPUser.vbs v1.0 was run by " & WshShell.ExpandEnvironmentStrings("%username%") & " at " & Time & " on " & Date & Chr(13) & Chr(10)

		ProcessArgs(objArgs)

		CreateTPUser
		CreateTPGroup
		AddUserToGroups

		CreateTPFolders
		
		AddACL RootDrv & "\ftproot\AvInbound\" & TPID & "\", "EXTNET\AV00" & TPID & "_FTP.Group", True
		RemoveTrustee RootDrv & "\ftproot\AvInbound\" & TPID, "EXTNET_FTP.Group"
		AddACL RootDrv & "\ftproot\AvOutbound\" & TPID & "\", "EXTNET\AV00" & TPID & "_FTP.Group", True
		RemoveTrustee RootDrv & "\ftproot\AvOutbound\" & TPID, "EXTNET_FTP.Group"

		OutputText True, Chr(13) & Chr(10) & "Script: Completed at " & Time & " on " & Date
		OutputText True, "-------------------------------------------------------" & Chr(13) & Chr(10)

	End If

End Sub	

Function CreateTPUser()
	' Create New Trading Parner User
	'
	On Error Resume Next
	
	Set objOU = GetObject("LDAP://OU=Users,OU=Avalon,dc=EXTNET,dc=CapBlueCross,dc=com")
	Set objUser = objOU.Create("User", "cn=AV00" & TPID & " User")
	objUser.Put "sAMAccountName", "AV00" & TPID
	objUser.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Setting sAMAccountName"
	End If
	objUser.Put "givenName", "Trading Partner"
	objUser.Put "sn", TPID
	objUser.Put "displayName", "AV00" & TPID & " User"
	objUser.AccountDisabled = FALSE
	objUser.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Setting givenName, sn, displayName and enabling the account"
	End If
	objUser.SetPassword password
	objUser.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Setting user password"
	End If
	userActCtrl = objUser.Get("userAccountControl")	
	userActCtrl = userActCtrl or ADS_UF_DONT_EXPIRE_PASSWD
	objUser.Put "userAccountControl", userActCtrl
	objUser.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Setting password not to expire"
	End If
	SetUserCannotChangePassword "LDAP://cn=AV00" & TPID & " User,OU=Users,OU=Avalon,dc=EXTNET,dc=CapBlueCross,dc=com","" ,"" ,1
'''''''''''''''''''''''''''''''''''''''
'Turn off the change password at next logon - added 11/14/06 jtw
	objuser.Put "pwdLastSet", -1
'''''''''''''''''''''''''''''''''''''''
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Setting 'User Cannot Change Password' option"
	End If
	OutputText True, "* AV00" & TPID & " Created"
	'WScript.Quit
End Function

Function CreateTPGroup()	 
	' Create associated Trading Partner Group
	'
	Set objOU = GetObject("LDAP://OU=Groups,OU=Avalon,dc=EXTNET,dc=CapBlueCross,dc=com")
	Set objGroup = objOU.Create("Group", "cn=AV00" & TPID & "_FTP.Group")
	objGroup.Put "sAMAccountName", "AV00" & TPID & "_FTP.Group"
	objGroup.Put "groupType", ADS_GROUP_TYPE_GLOBAL_GROUP Or _
	    ADS_GROUP_TYPE_SECURITY_ENABLED
	objGroup.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Creating AV00" & TPID & "_FTP.Group"
	End If
	OutputText True, "* AV00" & TPID & "_FTP.Group Created"
End Function

Function AddUserToGroups()	 

	' Add User to newely created Group
	'
	Set objGroup = GetObject("LDAP://cn=AV00" & TPID & "_FTP.Group,OU=Groups,OU=Avalon,dc=EXTNET,dc=CapBlueCross,dc=com")
	objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array("cn= AV00" & TPID & " User,OU=Users,OU=Avalon,dc=EXTNET,dc=CapBlueCross,dc=com")
	objGroup.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Adding user AV00" & TPID & " to AV00" & TPID & "_FTP.Group"
	End If
	OutputText True, "* AV00" & TPID & " added to AV00" & TPID & "_FTP.Group"

	' Add user to existing global EXTNET_FTP.Group Group
	'
	Set objGroup = GetObject("LDAP://CN=EXTNET_FTP.Group,OU=Trading Partners,dc=EXTNET,dc=CapBlueCross,dc=com")  
	objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array("cn= AV00" & TPID & " User,OU=Users,OU=Avalon,dc=EXTNET,dc=CapBlueCross,dc=com")
	objGroup.SetInfo
	If (Err.Number <> 0) Then
		BailOnFailure Err.Number, "Adding user AV00" & TPID & " to EXTNET_FTP.Group"
	End If
	OutputText True, "* AV00" & TPID & " added to EXTNET_FTP.Group"

End Function

Function CreateTPFolders()

	On Error Resume Next
	
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	If TPID > 2999 And TPID < 5000 Then

		If Not objFS.FolderExists (RootDrv & "\ftproot\AvOutbound\" & TPID) Then
			WshShell.Run scriptpath & "\Robocopy " & RootDrv & "\templates\FTPMailboxOutbound3xxx-4xxx " & RootDrv & "\ftproot\AvOutbound\" & TPID & " /SEC /E", 7, True
			'objFS.CopyFolder RootDrv & "\templates\FTPMailboxOutbound3xxx-4xxx", RootDrv & "\ftproot\AvOutbound\" & TPID
			If (Err.Number <> 0) Then
				BailOnFailure Err.Number, "Copying contents of template folder .\templates\FTPMailboxOutbound3xxx-4xxx To \ftproot\AvOutbound\" & TPID
			End If
			OutputText True, "* Created " & RootDrv & "\ftproot\AvOutbound\" & TPID & " from .\templates\FTPMailboxOutbound3xxx-4xxx"
		Else
			BailOnFailure 1, RootDrv & "\ftproot\AvOutbound\" & TPID & " aready exists!"
		End If
	
		If Not objFS.FolderExists (RootDrv & "\ftproot\AvInbound\" & TPID) Then
			WshShell.Run scriptpath & "\Robocopy " & RootDrv & "\templates\FTPMailboxInbound3xxx-4xxx " & RootDrv & "\ftproot\AvInbound\" & TPID & " /SEC /E", 7, True
			'objFS.CopyFolder RootDrv & "\templates\FTPMailboxInbound3xxx-4xxx", RootDrv & "\ftproot\AvInbound\" & TPID
			If (Err.Number <> 0) Then
				BailOnFailure Err.Number, "Copying contents of template folder .\templates\FTPMailboxInbound3xxx-4xxx To " & RootDrv & "\ftproot\AvInbound\" & TPID
			End If
			OutputText True, "* Created " & RootDrv & "\ftproot\AvInbound\" & TPID & " from .\templates\FTPMailboxInbound3xxx-4xxx"
		Else
			BailOnFailure 1, RootDrv & "\ftproot\AvInbound\" & TPID & " aready exists!"
		End If
		
	ElseIf TPID >= 5000 Or TPID < 3000 Then

		If Not objFS.FolderExists (RootDrv & "\ftproot\AvOutbound\" & TPID) Then
			WshShell.Run scriptpath & "\Robocopy " & RootDrv & "\templates\FTPMailboxOutbound6xxx " & RootDrv & "\ftproot\AvOutbound\" & TPID & " /SEC /E", 7, True
			'objFS.CopyFolder RootDrv & "\templates\FTPMailboxOutbound6xxx", RootDrv & "\ftproot\AvOutbound\" & TPID
			If (Err.Number <> 0) Then
				BailOnFailure Err.Number, "Copying contents of template folder .\templates\FTPMailboxOutbound6xxx To \ftproot\AvOutbound\" & TPID
			End If
			OutputText True, "* Created " & RootDrv & "\ftproot\AvOutbound\" & TPID & " from .\templates\FTPMailboxOutbound6xxx"
		Else
			BailOnFailure 1, RootDrv & "\ftproot\AvOutbound\" & TPID & " aready exists!"
		End If
	
		If Not objFS.FolderExists (RootDrv & "\ftproot\AVInbound\" & TPID) Then
			WshShell.Run scriptpath & "\Robocopy " & RootDrv & "\templates\FTPMailboxInbound6xxx " & RootDrv & "\ftproot\AvInbound\" & TPID & " /SEC /E", 7, True
			'objFS.CopyFolder RootDrv & "\templates\FTPMailboxInbound6xxx", RootDrv & "\ftproot\AvInbound\" & TPID
			If (Err.Number <> 0) Then
				BailOnFailure Err.Number, "Copying contents of template folder .\templates\FTPMailboxInbound6xxx To " & RootDrv & "\ftproot\AvInbound\" & TPID
			End If
			OutputText True, "* Created " & RootDrv & "\ftproot\AvInbound\" & TPID & " from .\templates\FTPMailboxInbound6xxx"
		Else
			BailOnFailure 1, RootDrv & "\ftproot\AvInbound\" & TPID & " aready exists!"
		End If

	Else
	
		BailOnFailure 1, "Unrecognized folder ID!"

	End If	

End Function

Function AddACL (strobjectPath, strTrustee, BlockInheritance)

	Set setAcl = CreateObject("SETACL.SetACLCtrl.1") 

   With SetAcl
      nError = .SetObject(strobjectPath, SE_FILE_OBJECT)	' Sets the file string and the object type
      If nError <> RTN_OK Then
         BailOnFailure 1, "SetObject failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If
      
      nError = .SetAction(ACTN_ADDACE)	' Sets action to Add and ACE
      If nError <> RTN_OK Then
         BailOnFailure 1, "SetAction failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If
      
      nError = .AddACE(strTrustee, False, "Change", INHPARCOPY, False, GRANT_ACCESS, ACL_DACL)	' Adds Trustee and permissions to object
      If nError <> RTN_OK Then
         BailOnFailure 1, "SetACEOptions failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If
      
      If BlockInheritance = True Then
	      nError = .AddAction(ACTN_SETINHFROMPAR)	' Needed to block inheritance
	      If nError <> RTN_OK Then
	         BailOnFailure 1, "SetAction failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
	         Exit Function
	      End If
	
	      nError = .SetObjectFlags(INHPARCOPY, INHPARCOPY, False, False)	' Stets the block inheritance flag
	      If nError <> RTN_OK Then
	         BailOnFailure 1, "SetObjectFlags failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
	         Exit Function
	      End If
	  End If

      nError = .Run	' Run the command
      If nError <> RTN_OK Then
         BailOnFailure 1, "Run failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If

   End With
	OutputText True, "* " & strTrustee & " added to " & strobjectPath & " with CHANGE access "
	OutputText True, "* Inheritance broken - " & BlockInheritance 
	
End Function

Function RemoveTrustee (strobjectPath, strTrustee)

	Set setAcl = CreateObject("SETACL.SetACLCtrl.1") 

   With SetAcl
      nError = .SetObject(strobjectPath, SE_FILE_OBJECT)	' Sets the file string and the object type
      If nError <> RTN_OK Then
         BailOnFailure 1, "SetObject failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If
      
        nError = .SetAction(ACTN_TRUSTEE)	' Sets action to remove a trustee
      If nError <> RTN_OK Then
         BailOnFailure 1, "SetAction failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If
      
      nError = .AddTrustee(strTrustee, "", False, False, ACTN_REMOVETRUSTEE, True, True)	' Removes Trustee permissions from object
      If nError <> RTN_OK Then
         BailOnFailure 1, "SetTrusteeOptions failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If

      nError = .Run	' Run the command
      If nError <> RTN_OK Then
         BailOnFailure 1, "Run failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Function
      End If

   End With

	OutputText True, "* " & strTrustee & " removed from " & strobjectPath

End Function


Function ProcessArgs(Arguments())

	For Each Arg in Arguments
		
		If Len(Arg) = 4 and IsNumeric(Arg) Then
			TPID = Arg
			OutputText True, "  - The Trading Partner ID is " & TPID
			OutputText True, ""
		End If

		If Left(Arg,1) = "/" Then
			Select Case Left(lcase(Arg),2)
				Case "/p"
					WorkProvided = 1
					password = Right(Arg,(Len(Arg) - 2))
					OutputText True, "  - Set user password to  " & password
					OutputText True, ""
			End Select
		End If
	Next			

End Function

'-------------------------------------------------------------------------
' OutputText
'   This subroutine display test to the screen and optionally write the
'   same test to a log file.
'
' Change log:
'   04/19/2002 (HBCXUPD) - Originaly created
'	07/08/2009 (SSCJSOW) - Modified for Avalon
'
' Parameters:
'   OutputText	String - The actual string to be echoed
'   LogWrite	True - The string will be echoed to the log file.
'		False - The string will NOT be echoed to the log file.
'
' Return value:
'   None
'-------------------------------------------------------------------------
Sub OutputText(LogWrite, OutText)	
	If LogWrite = True Then
		objLogFile.WriteLine(OutText)
	End If
	WScript.Echo OutText
End Sub

'''''''''''''''''''''''''''''''''''''''
'Failure Display subroutines
'''''''''''''''''''''''''''''''''''''''
Function BailOnFailure(ErrNum, ErrText)
	strText = "Error 0x" & Hex(ErrNum) & " (" & ErrNum & ") " & ErrText
	
	OutputText True, (strText)
	WScript.Quit
End Function

'******************************************************************************
'
'    SetUserCannotChangePassword
'
'    Sets the "User Cannot Change Password" permission using the LDAP provider
'    to the specified setting. This is accomplished by finding the existing
'    ACEs and modifying the AceType. The ACEs should always be present, but
'    it is possible that the default DACL excludes them. This function will not
'    work correctly if both ACEs are not present.
'
'    strUserDN - A string that contains the LDAP ADsPath of the user object to
'    modify.
'
'    strUsername - A string that contains the user name to use for
'    authorization. If this is an empty string, the credentials of the current
'    user are used.
'
'    strPassword - A string that contains the password to use for authorization.
'    This is ignored if strUsername is empty.
'
'    fCannotChangePassword - Contains the new setting for the privilege.
'    Contains True if the user cannot change their password or False if
'    the can change their password.
'
'******************************************************************************
Sub SetUserCannotChangePassword(strUserDN, strUsername, strPassword, fUserCannotChangePassword)
    Dim oUser
    Dim oSecDesc
    Dim oACL
    Dim oACE
    
    fEveryone = False
    fSelf = False
    
    If "" <> strUsername Then
        Dim dso
        
        ' Bind to the group with the specified username and password.
        Set dso = GetObject("LDAP:")
        Set oUser = dso.OpenDSObject(strUserDN, strUsername, strPassword, 1)
    Else
        ' Bind to the group with the current credentials.
        Set oUser = GetObject(strUserDN)
    End If
    
    Set oSecDesc = oUser.Get("ntSecurityDescriptor")
    Set oACL = oSecDesc.DiscretionaryAcl
    
    ' Modify the existing entries.
    For Each oACE In oACL
        If UCase(oACE.ObjectType) = UCase(CHANGE_PASSWORD_GUID) Then
            If oACE.Trustee = "Everyone" Then
                ' Modify the ace type of the entry.
                If fUserCannotChangePassword Then
                    oACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
                Else
                    oACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
                End If
            End If
        
            If oACE.Trustee = "NT AUTHORITY\SELF" Then
                ' Modify the ace type of the entry.
                If fUserCannotChangePassword Then
                    oACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
                Else
                    oACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
                End If
            End If
        End If
    Next
    
    ' Update the ntSecurityDescriptor property.
    oUser.Put "ntSecurityDescriptor", oSecDesc
    
    ' Commit the changes to the server.
    oUser.SetInfo
End Sub