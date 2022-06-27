$GroupName = "Azure.License.PowerBIPro"
Get-ADgroupmember -identity $GroupName | get-aduser -property displayname | select name, displayname|Export-Csv adGroupList-PowerBI.csv