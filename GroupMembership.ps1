# Name:  Group Membership
# Created By:  Jason Sowers
# Date:  June 23, 2022
#
$GroupName = "App.Corp.FEPDirect"
Get-ADGroupMember -Identity $GroupName -Recursive | select SAMAccountName, name | Export-Csv GroupUserList.csv