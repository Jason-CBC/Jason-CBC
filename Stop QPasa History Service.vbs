strComputer  = "EAAPP12"
strService = "HDService"
strWMIMoniker = "winmgmts:!//" & strComputer
strQuery = "select * from Win32_Service where Name='" & strService & "'"

For Each wmiWin32Service in GetObject(strWMIMoniker).ExecQuery(strQuery)
	wmiWin32Service.StopService()
Next
