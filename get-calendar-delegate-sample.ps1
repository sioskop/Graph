# Define AppId, secret and scope, your tenant name and endpoint URL
$AppId= "" #appid
$AppSecret= "" #appsecret
$Scope = "https://graph.microsoft.com/.default"
$TenantName = "" #.onmicrosoft.com of your domain
$Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"

# Add System.Web for urlencode
Add-Type -AssemblyName System.Web


# Create body
$Body = @{
    client_id = $AppId
      client_secret = $AppSecret
      scope = $Scope
      grant_type = 'client_credentials'
}

# Splat the parameters for Invoke-Restmethod for cleaner code
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method = 'POST'
    # Create string by joining bodylist with '&'
   Body = $Body
    Uri = $Url
}

# Request the token!
$Request = Invoke-RestMethod @PostSplat 


$Header = @{
    Authorization = "$($Request.token_type) $($Request.access_token)"
}


$user="" #user email address to list Calendar folder permission
$Uri = "https://graph.microsoft.com/v1.0/users/$user/calendar/calendarPermissions/"


$GraphRequest = Invoke-RestMethod -Uri $Uri -Headers $Header -Method Get -ContentType "application/json" 
 write-host -ForegroundColor DarkCyan "User Calendar permission list "$user" "
 write-host -ForegroundColor DarkCyan "-------------------------------------"
foreach ($p in $GraphRequest.value){
    
    write-host "Delegated user: "$p.emailaddress.Name" "
    $roles=$p.allowedRoles -join ", "
    write-host "Delegation permissions: "$roles
    write-host "Role access: "$p.role" "
    write-host -ForegroundColor DarkCyan "-------------------------------------"

}
