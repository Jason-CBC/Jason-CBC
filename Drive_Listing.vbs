'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.1
'
' NAME:  Count Drive Mappings
'
' AUTHOR: Jason C. Sowers , Capital Blue Cross
' DATE  : 9/4/2007
'
' COMMENT:  Scan a Merged File - Logon Scripts 
'
'==========================================================================
Dim strLinetoParse, objFSO, objTextFile, dtmDrive

Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\final.bat", _
    ForReading)

Do While objTextFile.AtEndOfStream <> True
    strLinetoParse = objTextFile.ReadLine
    dtmDrive = Mid(strLinetoParse, 9, 1)
    ldtmDrive = LCase(dtmDrive)
    Select Case ldtmDrive
    	Case "a"
    		Counta = Counta + 1
    	Case "b"
    		Countb = Countb + 1
    	Case "c"
    		Countc = Countc + 1
    	Case "d"
    		Countd = Countd + 1
    	Case "e"
    		Counte = Counte + 1
    	Case "f"
    		Countf = Countf + 1
    	Case "g"
    		Countg = Countg + 1
    	Case "h"
    		Counth = Counth + 1
    	Case "i"
    		Counti = Counti + 1
    	Case "j"
    		Countj = Countj + 1
    	Case "k" 
    		Countk = Countk + 1
    	Case "l"
    		Countl = Countl + 1
    	Case "m"
    		Countm = Countm + 1
    	Case "n"
    		Countn = Countn + 1
    	Case "o"
    		Counto = Counto + 1
    	Case "p"
    		Countp = Countp + 1
    	Case "q"
    		Countq = Countq + 1
    	Case "r"
    		Countr = Countr + 1
    	Case "s"
    		Counts = Counts + 1
    	Case "t"
    		Countt = Countt + 1
    	Case "u"
    		Countu = Countu + 1
    	Case "v"
    		Countv = Countv + 1
    	Case "w"
    		Countw = Countw + 1
    	Case "x"
    		Countx = Countx + 1
    	Case "y"
    		County = County + 1
    	Case "z"
    		Countz = Countz + 1
    End Select
Loop
WScript.Echo "Drive a: =" & Counta
WScript.Echo "Drive b: =" & Countb
WScript.Echo "Drive c: =" & Countc
WScript.Echo "Drive d: =" & Countd
WScript.Echo "Drive e: =" & Counte
WScript.Echo "Drive f: =" & Countf
WScript.Echo "Drive g: =" & Countg
WScript.Echo "Drive h: =" & Counth
WScript.Echo "Drive i: =" & Counti
WScript.Echo "Drive j: =" & Countj
WScript.Echo "Drive k: =" & Countk
WScript.Echo "Drive l: =" & Countl
WScript.Echo "Drive m: =" & Countm
WScript.Echo "Drive n: =" & Countn
WScript.Echo "Drive o: =" & Counto
WScript.Echo "Drive p: =" & Countp
WScript.Echo "Drive q: =" & Countq
WScript.Echo "Drive r: =" & Countr
WScript.Echo "Drive s: =" & Counts
WScript.Echo "Drive t: =" & Countt
WScript.Echo "Drive u: =" & Countu
WScript.Echo "Drive v: =" & Countv
WScript.Echo "Drive w: =" & Countw
WScript.Echo "Drive x: =" & Countx
WScript.Echo "Drive y: =" & County
WScript.Echo "Drive z: =" & Countz
objTextFile.Close

