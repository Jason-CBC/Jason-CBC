'**********************************************************************************
' SYSTEM:       LifeMasters
' APPLICATION:  LifeMasters Monthly Billing and Reconciliation Process 
' AUTHOR:       Kevin Rill
' DATE:         05/21/2003
' PRODUCTION:   05/21/2003
'
' DESCRIPTION:  Visual Basic Script to complete the following tasks.
'		- Delete the prior months CMBS and Facets Eligibility text files.
'		- Access a .dll file to obtain the dedicated/nonexpiring
'		    userid and password.
'               - Create a text file containg the necessary FTP parameters to 
'		    download the mainframe based CMBS and Facets Eligibility files.
'               - Execute a shell command to invoke the FTP process.
'		- Delete the text file containing the FTP parameters.
'
' CHANGE LOG:	ACMS 2002000923 K. Rill	   05/21/2003	New VB Script
' Edited On:  07/15/2003 JCS and CEP
'**********************************************************************************

sub ftp_lm_elig()

on error resume next

Dim oShell
Dim oFSO
Dim oCMBSfso
Dim oFACETSfso
Dim sCmd
Dim sFtpCmds
Dim sCMBSfile
Dim sFACETSfile
Dim PWObject, Password, UserID
Dim iDebugsw

Const OverwriteIfExist = -1
Const OpenAsASCII =  0



'**********************************************************************************
'  Set Debug Switch to "Y" or "N".
'**********************************************************************************
iDebugsw = "N"
'iDebugsw = "N"

'**********************************************************************************
'  Delete and Create a text file for writing debugging statements.
'**********************************************************************************
sDebugFile = "E:\LifeMasters\Debug.txt"
Set oDebugFSO = CreateObject("Scripting.FileSystemObject")

oDebugFSO.DeleteFile(sDebugfile)

If Err.Number = 53 Or Err.Number = 0 Then
    Err.Clear           ' Clear Err object fields
Else
    msgbox Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Debug Error Message"
End If

If iDebugsw = "Y" Then
    Set fDebugMsgs =  oDebugFSO.CreateTextFile(sDebugFile, OverwriteIfExist, OpenAsASCII)
    fDebugMsgs.WriteLine "Debug file created"
End If



'**********************************************************************************
'  Delete the prior months CMBS and Facets Eligibility text files.
'**********************************************************************************
sCMBSfile = "E:\LifeMasters\Eligibility_CMBS.txt"
sFACETSfile = "E:\LifeMasters\Eligibility_FACETS.txt"

set oCMBSfso = Createobject("Scripting.FileSystemObject")
set oFACETSfso = Createobject("Scripting.FileSystemObject")

oCMBSfso.DeleteFile(sCMBSfile)

If Err.Number = 53 Or Err.Number = 0 Then
    Err.Clear           ' Clear Err object fields
Else
    msgbox Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Error Message"
End If

oFACETSfso.DeleteFile(sFACETSfile)
If Err.Number = 53 Or Err.Number = 0 Then
    Err.Clear           ' Clear Err object fields
Else
    msgbox Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Error Message"
End If


If iDebugsw = "Y" Then
    fDebugMsgs.WriteLine "Prior Month Eligibility files deleted"
End If


'**********************************************************************************
'  Access a .dll file to obtain the dedicated/nonexpiring userid and password.
'**********************************************************************************
' Create a server object (dll) that holds the userid and password
Set PWObject = CreateObject("LM_PW.LM_PW_ID")

UserID = PWObject.Get_UserID()
Password = PWObject.Get_Password()

' Release the object resources
Set PWObject = Nothing


If iDebugsw = "Y" Then
    fDebugMsgs.WriteLine "DLL object created"
End If


'**********************************************************************************
'  Create a text file containg the FTP parameters to download the mainframe files.
'**********************************************************************************
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

sFtpCmds = "E:\LifeMasters\ftp_main.inp"

Set fFtpCmds =  oFSO.CreateTextFile(sFtpCmds, OverwriteIfExist, OpenAsASCII)

' server to open
fFtpCmds.WriteLine "open cbcmvs1.internal.capbluecross.com"
fFtpCmds.WriteLine UserID
fFtpCmds.WriteLine Password
fFtpCmds.WriteLine "lcd " & """E:\LifeMasters"""
fFtpCmds.WriteLine "get 'HCMP.HCMPM95X.LFMST.CNT.EXT(0)' Eligibility_CMBS.txt"
fFtpCmds.WriteLine "get 'HCMP.HFMSI03.LMSTELIG.FULLFILE(0)' Eligibility_FACETS.txt"
fFtpCmds.WriteLine "quit"

fFtpCmds.Close

If iDebugsw = "Y" Then
    fDebugMsgs.WriteLine "FTP statements created"
End If


'**********************************************************************************
'  Execute a shell command to invoke the FTP process.
'**********************************************************************************
sCmd = "ftp -i -s:" & """E:\LifeMasters\ftp_main.inp"""
                   
oShell.Run sCmd, 0, True


If iDebugsw = "Y" Then
    fDebugMsgs.WriteLine "Shell command executed"
End If


'**********************************************************************************
'  Delete the text file containing the FTP parameters.
'**********************************************************************************
oFSO.DeleteFile(sFtpCmds)


'**********************************************************************************
'  Close Debug text file containing the Debug messages.
'**********************************************************************************
fDebugMsgs.Close


End Sub

ftp_lm_elig
