'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 3.0
'
' NAME: VPN Drive Mapper
'
' AUTHOR: Jason C. Sowers
' DATE  : 08/04/2004
'
' COMMENT: Checking for mappings before making new connections, 
'          disconnecting unwanted mappings.
'
'==========================================================================
Option Explicit

On Error Resume Next
Dim WshNetwork, oDrives, i, sUserName, WshShell, objEnv

'-- Setting objects --
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

'Set WshShell = WScript.CreateObject("WScript.Shell")
'Set objEnv = WshShell.Environment("PROCESS")
'sUserName = objEnv("USERNAME")
sUserName = InputBox("Please enter your BCNET Userid", "BCNET Userid Request for VPN Access")

'Example:
DriveMapper "h:", "\\easrv02\" & sUserName & "$"
DriveMapper "p:", "\\easrv03\public"

Sub DriveMapper(Drive, Share)
	For i = 0 to oDrives.Count -1 Step 2
	If LCase(Drive) = LCase(oDrives.Item(i)) Then
		If not LCase(Share) = LCase(oDrives.Item(i+1)) Then
			WshNetwork.RemoveNetworkDrive Drive, true, True
		Else
			Exit Sub
		End If
	End If
	Next
	WshNetwork.MapNetworkDrive Drive, Share
End Sub