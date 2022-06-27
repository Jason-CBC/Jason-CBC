'**********************************************************************************
' SYSTEM:       Encrypted Password FTP Process
' APPLICATION:  Tidal Trigger File Security 
' AUTHOR:       Jason C. Sowers
' DATE:         02/24/2006
'
' Portions of code provided by Greg Page and Kevin Rill
'**********************************************************************************

sub ftp_lm_elig()

'on error resume next

Dim oShell
Dim oFSO
Dim sCmd
Dim sFtpCmds
Dim iDebugsw, oDebugFSO, sDebugFile
Dim val, enUser, enPass, deUser, dePass
Dim encryptor

Const OverwriteIfExist = -1
Const OpenAsASCII =  0



'**********************************************************************************
'  Set Debug Switch to "Y" or "N".
'**********************************************************************************
iDebugsw = "N"
'iDebugsw = "Y"

'**********************************************************************************
'  Encryption to obtain the dedicated/nonexpiring userid and password.
'**********************************************************************************
' Create a server object (encryptor) that holds the UserID & Password
'val = "Provider=MSDAORA.1;Password=pwd;User ID=userid;Data Source=D001"

Set encryptor = CreateObject("Com.CapBlueCross.Core.Encryption.RijndaelEncryptor")
encryptor.setParameters "D09Dkd045nas03/gk20Ma2==", "J58DmjLm380Ma1mqi3kpZu=="

enUser = "KXd1k8IncNchWEgfnLhhIg=="
enPass = "IYCP8ZekjSdCFN1wciAB7w=="

' Release the object resources
Set encryptor = Nothing


If iDebugsw = "Y" Then
    fDebugMsgs.WriteLine "Encryption object created"
End If


'**********************************************************************************
'  Create a text file containg the FTP parameters to download the mainframe files.
'**********************************************************************************
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

sFtpCmds = "D:\ftp_main.inp"

Set fFtpCmds =  oFSO.CreateTextFile(sFtpCmds, OverwriteIfExist, OpenAsASCII)

' server to open
fFtpCmds.WriteLine "open 10.10.4.37"
deUser = encryptor.decrypt(enUser)
fFtpCmds.WriteLine deUser
dePass = encryptor.decrypt(enPass)
fFtpCmds.WriteLine dePass
fFtpCmds.WriteLine "asc"
fFtpCmds.WriteLine "cd /apps/hsa/transfer"
'fFtpCmds.WriteLine "lcd " & """E:\LifeMasters"""
fFtpCmds.WriteLine "LCD " & """D:\Inetpub\ftproot\Inbox\UnixTransfer"""
fFtpCmds.WriteLine "put triggerfile.txt"
fFtpCmds.WriteLine "put trigger_hra.txt"
fFtpCmds.WriteLine "quit"

fFtpCmds.Close

If iDebugsw = "Y" Then
    fDebugMsgs.WriteLine "FTP statements created"
End If


'**********************************************************************************
'  Execute a shell command to invoke the FTP process.
'**********************************************************************************
sCmd = "ftp -i -s:" & """D:\Scripts\ftp_main.inp"""
                   
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