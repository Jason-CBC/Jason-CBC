'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 3.0
'
' NAME: 
'
' AUTHOR: CBC , CBC
' DATE  : 12/07/2004
'
' COMMENT: 
'
'==========================================================================

Set FSO = CreateObject("Scripting.FileSystemObject")
ShowSubfolders FSO.GetFolder("C:\Scripts")

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        Wscript.Echo Subfolder.Path
        WScript.Echo Subfolder.Size
        ShowSubFolders Subfolder
    Next
End Sub