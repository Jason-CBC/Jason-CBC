'****************************************************
'*
'*  Monitor Folders for Creation of New Files
'*
'*  Created By:  Jason C. Sowers
'*  Created On:  02/19/2003
'*
'****************************************************

Option Explicit

'****************************************************
'Variables Section
'****************************************************
Dim CurDate, objFolder, colSubfolders
Dim TargetDate, sInterval, objFSO, WSH
Dim objSubfolder, sNewTime, sSubFolTime
Dim dCounter

'****************************************************
'Initialize Variables Section
'****************************************************
CurDate = Date()
TargetDate = CurDate - 1
sInterval = "d"
dCounter = 0

'****************************************************
'Set Objects
'****************************************************
Set objFSO = CreateObject ("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("E:\NVA")
Set colSubfolders = objFolder.Subfolders
Set WSH = Wscript.CreateObject("Wscript.Shell")

'****************************************************
'Process Section
'****************************************************

For Each objSubfolder in colSubfolders
	sSubFolTime = objSubfolder.DateLastModified
	sNewTime = DateDiff(sInterval, sSubFolTime, TargetDate)
	If sNewTime < 0 Then
		If Not objSubfolder.Name = "LogFiles" Then
			dCounter = dCounter + 1
		End If
	'WScript.Echo objSubfolder.Name, objSubfolder.Size, objSubfolder.DateLastModified
    'WScript.Echo TargetDate
    End If
Next

If dCounter > 0 Then
	WSH.Run "blat c:\blat.txt -to cbcvisionreportusers@capbluecross.com" & _
	" -f NVA_Vision_Mailer@capbluecross.com -s NVA-Files"
End If

dCounter = 0

WScript.Quit