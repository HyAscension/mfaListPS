Write-Host "Hello World!"

# Ignore non-terminating error. Displays the error message and stops executing
$ErrorActionPreference = "Stop"

# Output message delayed by one second
function Write-DelayedMessage($message) {
    Write-Host $message
    Start-Sleep -Seconds 1
}

Write-DelayedMessage -message "Starting Section 1..."

function Get-UserInfo($clientID, $tenantID, $cert) {
    # Use Command: Install-Module Microsoft.Graph -Scope CurrentUser
    Connect-MgGraph -ClientID $clientID -TenantID $tenantID -CertificateThumbprint $cert
}

function Get-GraphUser($searchLimit, [String[]]$parameters) {
    # Read properties and relationships of the user object
    $mgAccList = Get-MgUser -Top $searchLimit -Select $parameters | Where-Object { $_.UserType -ne "Guest" }
    Write-DelayedMessage -message "Total $($usersAll.Count) accounts"

    # New generic list instance
    $userLogData = [System.Collections.Generic.List[Object]]::new()

    # Loop and populate object with specific data
    foreach($acc in $mgAccList)
    {
        # Every field here match selected column data in Get-MgUser above
        $eachAccObj = [PSCustomObject] @{
            UserPrincipalName = $acc.UserPrincipalName
            Name = $acc.DisplayName
            # Apply ternary statement
            IsLicensed = ($acc.AssignedLicenses.Count -ne 0) ? $true : $false
            IsGuestUser = ($acc.UserType -eq 'Guest') ? $true : $false
            AppInteractLastLogin = $acc.SignInActivity.LastSignInDateTime
            TokenInteractLastLogin = $acc.SignInActivity.LastNonInteractiveSignInDateTime
        }

        # Add obj to generic list
        $userLogData.Add($eachAccObj)
    }

    # Create a grid view of object list
    $userLogData | Select-Object UserPrincipalName, Name, IsLicensed, IsGuestUser, AppInteractLastLogin, TokenInteractLastLogin | Out-GridView
    $userLogData | Export-CSV -NoTypeInformation -Encoding UTF8 # "(Directory link)"
    python # "(Directory link)"
}

Get-UserInfo -clientID "" -tenantID "" -cert ""
Get-GraphUser -searchLimit 0 -parameters # Parameters