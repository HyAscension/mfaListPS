 <#
                      __,,,,_
       _ __..-;''`--/'/ /.',-`-.
   (`/' ` |  \ \ \\ / / / / .-'/`,_
  /'`\ \   |  \ | \| // // / -.,/_,'-,
 /<7' ;  \ \  | ; ||/ /| | \/    |`-/,/-.,_,/')
/  _.-, `,-\,__|  _-| / \ \/|_/  |    '-/.;.\'
`-`  f/ ;      / __/ \__ `/ |__/ |
     `-'      |  -| =|\_  \  |-' |
           __/   /_..-' `  ),'  //
       RC ((__.-'((___..-'' \__.'

#>

# Ignore non-terminating error. Displays the error message and stops executing
$ErrorActionPreference = "Stop"

# Output message delayed by one second
function Write-DelayedMessage($message) {
    Write-Host $message
    Start-Sleep -Seconds 1
}

function Get-UserInfo($clientID, $tenantID, $cert) {
    # Use Command: Install-Module Microsoft.Graph -Scope CurrentUser
    Connect-MgGraph -ClientID $clientID -TenantID $tenantID -CertificateThumbprint $cert
}

function Get-GraphUser($searchLimit, [String[]]$parameters) {
    # Read properties and relationships of the user object
    $mgAccList = Get-MgUser -Top $searchLimit -Select $parameters | Where-Object { $_.UserType -ne "Guest" }
    Write-DelayedMessage -message "Total $($mgAccList.Count) accounts"

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
            IsLicensed = if ($acc.AssignedLicenses.Count -ne 0) { $true } else { $false }
            IsGuestUser = if ($acc.UserType -eq 'Guest') { $true } else { $false }
            AppInteractLastLogin = $acc.SignInActivity.LastSignInDateTime
            TokenInteractLastLogin = $acc.SignInActivity.LastNonInteractiveSignInDateTime
        }

        # Add obj to generic list
        $userLogData.Add($eachAccObj)
    }

    # Create a grid view of object list
    # $userLogData | Select-Object UserPrincipalName, Name, IsLicensed, IsGuestUser, AppInteractLastLogin, TokenInteractLastLogin | Out-GridView
    $userLogData | Export-CSV -NoTypeInformation -Encoding UTF8 Path\output\recentLoginList.csv
    python Path\python\concat_two_tables.py
}

# Get password file and convert to secure password. Use createHiddenPassword.ps1 to create
$username = ""
$securedPass = Get-Content "file" | ConvertTo-SecureString
$msCred = New-Object System.Management.Automation.PSCredential($username, $securedPass)

function Get-MSUser {
    # Use command: Install-Module MSOnline
    Connect-MsolService -Credential $msCred

    # Search all users
    $users = Get-MsolUser -All | Where-Object { $_.UserType -ne "Guest" }

    # Create output file
    $genericList = [System.Collections.Generic.List[Object]]::new()
    Write-DelayedMessage -message "Read $($Users.Count) accounts..."
    
    # Loop through each users to create a custom data line for csv extract
    foreach ($person in $users) {
        $MFAEnforced = $person.StrongAuthenticationRequirements.State
        $MFAPhone = $person.StrongAuthenticationUserDetails
        $DefaultMFAMethod = ($person.StrongAuthenticationMethods | Where-Object { $_.IsDefault -eq "True" }).MethodType

        #Categorize default mfa methods
        if (($MFAEnforced -eq "Enforced") -or ($MFAEnforced -eq "Enabled")) {
            switch ($DefaultMFAMethod) {
                "PhoneAppOTP" { $MethodUsed = "Authenticator app" }
                "PhoneAppNotification" { $MethodUsed = "Authenticator app" }
                "OneWaySMS" { $MethodUsed = "One-way SMS" }
                "TwoWayVoiceMobile" { $MethodUsed = "Phone call" }
                "TwoWayVoiceOffice" { $MethodUsed = "Phone call" }
                "PhoneApp" { $MethodUsed = "Authenticator app" }
                "HardwareToken" { $MethodUsed = "Hardware token" }
                "Email" { $MethodUsed = "Email" }
                "Unknown" { $MethodUsed = "Unknown" }
            }
            if ($MethodUsed -eq "MFA Not Used") { $MethodUsed = "Uncertain" }
        }
        elseif ($MFAEnforced -eq "") {
            $MFAEnforced = "Not Enabled"
            $MethodUsed = "MFA Not Used"
            $MFAPhone = ""
        }
        else {
            $MFAEnforced = "Not Enabled"
            $MethodUsed = "MFA Not Used"
            $MFAPhone = ""
        }

        $activeStatus = $person.BlockCredential
        $accStatsOutput = ""

        if ($activeStatus -eq $false) {
            $accStatsOutput = "Active"
        } 
        elseif ($activeStatus -eq $true) {
            $accStatsOutput = "Disabled" 
        }

        $userObj = [PSCustomObject] @{
            User        = $person.UserPrincipalName
            Name        = $person.DisplayName
            Department  = $person.Department
            JobTitle    = $person.JobTitle
            MFAUsed     = $MFAEnforced
            MFAMethod   = $MethodUsed
            PhoneNumber = $MFAPhone
            ActiveAccounts  = $accStatsOutput
        }

        $genericList.Add($userObj)
    }

    # $genericList | Select-Object User, Name, Department, JobTitle, MFAUsed, MFAMethod, PhoneNumber, ActiveAccounts | Sort Name | Out-GridView
    $genericList | Sort-Object Name | Export-CSV -NoTypeInformation -Encoding UTF8 Path\output\recentLoginList.csv
}

function Get-MailInfo {
    Connect-ExchangeOnline -Credential $msCred

    $sharedMBList = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Select-Object Name,WindowsEmailAddress
    $roomMBList = Get-Mailbox -RecipientTypeDetails RoomMailbox | Select-Object Identity, PrimarySmtpAddress
    $equipmentMBList = Get-Mailbox -RecipientTypeDetails EquipmentMailbox | Select-Object Identity, PrimarySmtpAddress

    $sharedMBList | Sort-Object Name | Export-CSV -NoTypeInformation -Encoding UTF8 Path\output\sharedMBList.csv
    $roomMBList | Sort-Object Identity | Export-CSV -NoTypeInformation -Encoding UTF8 Path\output\roomMBList.csv
    $equipmentMBList | Sort-Object Identity | Export-CSV -NoTypeInformation -Encoding UTF8 Path\output\equipmentMBList.csv

    python Path\custom_row_rm.py
    Write-DelayedMessage -Message "Both reports are in Path\output folder"
}

Write-DelayedMessage -message "Starting Section 1..."
Get-UserInfo -clientID "" -tenantID "" -cert ""
Get-GraphUser -searchLimit 3000 -parameters UserPrincipalName, DisplayName, AssignedLicenses, UserType, SignInActivity

Write-DelayedMessage -message "Starting Section 2..."
Get-MSUser
Get-MailInfo
