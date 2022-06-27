'************************************************
'*
'*  Parse "Search and Replace" Log and File Copy
'*
'*  Created By:  Jason C. Sowers
'*  Created On:  12/06/2002
'*	Modified On:  06/30/2003
'*
'************************************************

Option Explicit

'************************************************
'*  Variables Section
'************************************************
Dim sPath, oFSO, FileName, strLine
Dim strWord, strLen, strTest, strText
Dim strPath, NewFile, WordLength
Dim CopyFile, f1, strCopy, LogFile
Dim MyDate, MyTime, mVar

strWord = "Processing file : "

Const ForReading = 1
Const ForWriting = 2
Const Destination = "G:\P_FIN\Working Aged\" 'Edit as needed
Const LogFileLoc = "G:\P_FIN\Working Aged\CopiedFiles.txt"  'Edit as needed
Const OverwriteExisting = True

'************************************************
'*  Object Creation Section
'************************************************

Set oFSO = CreateObject ("Scripting.FileSystemObject")
Set LogFile=oFSO.OpenTextFile(LogFileLoc, 8, TRUE)

'************************************************
'*  Main Program
'************************************************

Parse  'Execute Subroutine
MyDate = Date
MyTime = Time
Question  'Execute Subroutine
LogFile.WriteLine "Begin Process " & " " & MyDate & " " & MyTime
Copy  'Execute Subroutine
MyDate = Date
MyTime = Time
LogFile.WriteLine "End Process " & " " & MyDate & " " & MyTime
LogFile.WriteLine " "

'************************************************
'*  Functions and Subroutines
'************************************************
Sub Parse

Set FileName = oFSO.OpenTextFile ("H:\Medicare Scans\P FIN Working Aged results.txt", ForReading)
Set NewFile = oFSO.CreateTextFile ("G:\CopyList.txt", ForWriting)

WordLength = Len(strWord)

Do 
	On Error Resume Next
	strText = FileName.ReadLine
	strLen = Len(strText)
	strTest = Instr(1, strText, strWord)
	If strTest = 1 Then
		strPath = Right(strText,(strLen - WordLength))
		NewFile.WriteLine(strPath)
	End If
Loop While Not FileName.AtEndOfStream

FileName.Close

End Sub
'*************************************************
Sub Copy

Set CopyFile = oFSO.OpenTextFile("G:\test.txt", ForReading)

Do 
	On Error Resume Next
	strCopy = CopyFile.ReadLine
	Set f1 = oFSO.GetFile(strCopy)
	WScript.Echo(Destination)
	f1.Copy Destination,TRUE
	LogFile.WriteLine f1.Path & vbTab & f1.Size & " kb" & " " &_
	vbTab & vbTab & vbTab & "was copied" & vbTab & Now
	
	'  error handling
	If Err.Number Then
		LogFile.WriteLine "Error Message" &_
		vbNewLine & "Error Number: " &_
		Err.Number & vbNewLine &_
		"Error Description: " & Err.Description
		Err.Clear
	End If
Loop While Not CopyFile.AtEndOfStream

CopyFile.Close

End Sub 
'**************************************************
Sub Question

mVar = MsgBox ("Parsing Completed, continue with File Copy?", 36, "Parsing Completed")
If mVar = 6 Then
	Exit Sub
Else
	WScript.Quit(0)
End If

End Sub