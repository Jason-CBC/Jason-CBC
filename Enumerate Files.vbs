'===================================================================
'
' NAME: filemap.vbs
'
' AUTHOR: Jason C. Sowers
' DATE : 02/10/2004
'
' COMMENT: Uses WMI to produce a list of files that match the criteria specified.
'               The script also calculates the amount of space the files take up.
'
'====================================================================

Set FileSet = GetObject("winmgmts:").ExecQuery("select * from CIM_DataFile where Drive = 'C:' AND Extension= 'mdb'")
Set WshShell = WScript.CreateObject("WScript.Shell")

'The following returns the Computername so I can use it in the name of the file. 
Set WshSysEnv = WshShell.Environment("PROCESS")
strMachine = WshSysEnv("COMPUTERNAME")

Dim intTotalSpace, tCount
intTotalSpace = 0
tCount = 0

'Get a handle on the output file. Change the filename/location here.

fname = WshShell.SpecialFolders("Desktop")& "\" & strMachine & "-mdbfile.csv"

set fso = CreateObject ("Scripting.FileSystemObject")

If fso.FileExists (fname) THEN
     set objFile = fso.GetFile (fname)
     objFile.Delete
End If

'Create the output file.
set objFile = fso.CreateTextFile (fname, True)


for each File in FileSet
		  fYear = Left(File.LastModified,4)
		  fMonth = Mid(File.LastModified,5,2)
		  fDay = Mid(File.LastModified,7,2)
          objFile.WriteLine File.Name & "," & File.FileSize & "," & fMonth & "/" & fDay & "/" & fYear
          intFileSize = File.FileSize
          intTotalSpace = intTotalSpace + intFileSize
          tCount = tCount + 1
          'Wscript.Echo File.Name & "," & intFileSize & "," & fMonth & "/" & fDay & "/" & fYear
          'NB: Other info you can grab is File.Encrypted, File.FileType, File.Extension and more.
Next

objFile.WriteLine "Complete..." & vbcrlf & "Used Space =" & intTotalMegs & " MB. " & "Total Files: " & tCount

Set FileSet = nothing
Set WshShell = nothing
Set fname = nothing

intTotalMegs = (intTotalSpace/1024)/1024
Wscript.Echo "Complete..." & vbcrlf & "Used Space =" & intTotalMegs & " MB. " & "Total Files: " & tCount