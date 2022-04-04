#region Initialize default properties
$config = ConvertFrom-Json $configuration
$p = $person | ConvertFrom-Json -AsHashTable

$success = $false
$auditLogs = [Collections.Generic.List[PSCustomObject]]::new();

# AzureAD Application Parameters #
$config = ConvertFrom-Json $configuration

$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

#endregion Initialize default properties

#region Change mapping here
# Change mapping here

#region Execute
try{
    #Find Azure AD ACcount by UserPrincipalName
    Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
    $baseAuthUri = "https://login.microsoftonline.com/"
    $authUri = $baseAuthUri + "$($AADTenantID)/oauth2/token"

    $body = @{
        grant_type      = "client_credentials"
        client_id       = "$($AADAppId)"
        client_secret   = "$($AADAppSecret)"
        resource        = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken"
        'Content-Type' = "application/json"
        Accept = "application/json"
    }

    $baseGraphUri = "https://graph.microsoft.com/"
    $searchUri = $baseGraphUri + "v1.0/users?`$filter=employeeID eq '$($p.ExternalId)'"
    #Write-Information $searchUri
    $response = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    
    $account = $response.value

    if ($account.count -gt 1) { throw "Found multiple Users $($account.userPrincipalName)" }
    if ($null -eq $account.id) { throw "Could not find Azure user $($account.userPrincipalName)" }

    Write-Information "Account correlated to $($account.userPrincipalName) ($($account.id))"
    $aRef = $account.id

	$auditLogs.Add([PSCustomObject]@{
                Action = "CreateAccount"
                Message = "Account correlated to $($account.userPrincipalName) ($($account.id))"
                IsError = $false
            })

    # Find Manager
    Write-Information "Manager ID: $($p.PrimaryContract.Manager.ExternalId)" 
    
    $baseGraphUri = "https://graph.microsoft.com/"
    $searchUri = $baseGraphUri + "v1.0/users?`$filter=employeeID eq '$($p.PrimaryContract.Manager.ExternalId)'"
    #Write-Information $searchUri
    
    $response = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $azureManager = $response.value
    
    Write-Information "Manager: $($azureManager.userPrincipalName)"
    
    #Set Manager
    if($null -ne $azureManager.Id)
    {
        if(-Not($dryRun -eq $True)) {
             $baseGraphUri = "https://graph.microsoft.com/"
             $Uri = $baseGraphUri + "v1.0/users/$($account.id)/manager/`$ref"

             $body = @{ "@odata.id"= "https://graph.microsoft.com/v1.0/users/$($azureManager.id)" }
             
             $response = Invoke-RestMethod -Uri $Uri -Method PUT -Headers $authorization -Body ($body | ConvertTo-Json)

             $auditLogs.Add([PSCustomObject]@{
                Action = "UpdateAccount"
                Message = "Updated manager to $($azureManager.UserPrincipalName) ($($azureManager.id))"
                IsError = $false;
            });
        }

    }
    

    $success = $true
} catch {
    $auditLogs.Add([PSCustomObject]@{
                Action = "CreateAccount"
                Message = "Account failed to correlate to $($account.userPrincipalName): $_"
                IsError = $True
            })
	Write-Verbose -Verbose "$_"
}
#endregion Execute

#region build up result
$result = [PSCustomObject]@{
    Success= $success
    AccountReference= $aRef
    AuditLogs = $auditLogs
    Account = @{ UserPrincipalName = $account.userPrincipalName}

}

Write-Output $result | ConvertTo-Json -Depth 10
#endregion build up result
