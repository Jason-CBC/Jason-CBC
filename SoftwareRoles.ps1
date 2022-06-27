$GroupName = "ITDesktopServices.SCCM*"
Get-ADGroup -filter {name -like $GroupName} -properties * |select SAMAccountName, Description|Export-Csv adGroupList.csv