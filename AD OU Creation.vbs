'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 3.0
'
' NAME: 
'
' AUTHOR: CBC , CBC
' DATE  : 04/16/2004
'
' COMMENT: 
'
'==========================================================================

'Script to create users and ous for dev environment


CRLF = Chr(13) & Chr(10)

'This statement binds the Global catalog
Set root = GetObject("GC://RootDSE") 

'Default Domain capture
Set domname = GetObject( "GC://" & root.Get("rootDomainNamingContext"))

'Ask user for input
dspath=inputbox("Please specify the directory path."+CRLF+"For example: DC=e-learn,DC=com","Sample NT 

5.0 Script", domname.distinguishedname)
If dspath="" then
WScript.Quit(1)
end if


ounum=inputbox("How many OU's would you like to create?","Sample NT 5.0 Script",1)
If ounum="" then
WScript.Quit(1)
end if

usernum=inputbox("How many user's would you like to create per OU?","Sample NT 5.0 Script",10)
If usernum="" then
WScript.Quit(1)
end if


msgbox "OU's and user accounts will now be created. This may take a while, depending on how many objects 

you specified."

'Set the starting LDAP path
set ns=getobject("LDAP://"+dspath)

'Create OUs and Users via for for next loop

for x=1 to ounum
oname="NT5OU-"+Cstr(x)

' the command to create an OU
set orgu=ns.Create("organizationalunit", "OU="+oname)

' this writes it to the DS
orgu.setinfo

for y= 1 to usernum
'specify context for the creation of the user
Set u = ns.Create("user", "CN="+"NT5User"+cstr(x)+"-"+cstr(y)+",OU="+oname)

'specify downstream (NT 4.0) compatible name
u.Put "sAMAccountName", "NT5User"+cstr(x)+"-"+cstr(y)

'specify UPN
u.PUT "userPrincipalName","NT5User"+cstr(x)+"-"+cstr(y)+"@nowhere.com"

'set user account properties 
u.Put "userAccountControl",66048
u.PUT "userPassword","password"

'write user account to DS
u.SetInfo
next
Next

msgbox "Users and OU's Created!"