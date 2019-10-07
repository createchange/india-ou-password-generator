get-aduser -filter * -Properties DistinguishedName,SamAccountName,EmailAddress |
where {$_.emailaddress} |
select @{N='Distinguished Name';E={$_.DistinguishedName}},@{N='UserName';E={$_.SamAccountName}},@{N='Email Address';E={$_.EmailAddress}} |
convertto-json | 
Set-Content C:\Users\jhweaver\Desktop\AD_User_info.json
