<#
.SYNOPSIS
    Different Exchange Online Modules
.DESCRIPTION
    - Generate new MoveRequests from AD Group Membership or OU
    - Connect to Office 365 and Exchange Online
    - Complete MoveRequest
    - Setting Mailbox Settings for Mailboxes
.OUTPUTS
    Nothing but magic
.EXAMPLE
    New-O365MoveRequest
    Complete-O365MoveRequest
    Set-O365MailboxSettings
    Connect-O365 
    Add-O365License 


.NOTES
    Author:         Sebastian Wild	
    Email:          sebastian.wild@dynabcs.at
    Company:        DynaBCS Informatik
	Date :          06.04.2021

    Changelog:
    1.1             Updated Functions
	1.0             Initial Release
#>

function New-O365MoveRequest {
    [CmdletBinding()]
    param (
        [string]$Group,
        [string]$OU, 
        [Parameter(Mandatory)]
        [string]$TargetDeliveryDomain,
        [switch]$Suspendwhenreadytocomplete

    )
	
    begin {
        Import-Module ActiveDirectory, MSOnline, CredentialManager

        $namefilter = @("Mailbox1", "Discovery*", "Administrator", "Health*" )

        try {
            $cred = Get-StoredCredential -target O365
            $opcred = Get-StoredCredential -target AD
	
            if ($null -eq (Get-PSSession)) {
                $cred = Get-StoredCredential -target O365	
                Connect-MsolService -Credential $cred
                $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
                Import-PSSession $s -AllowClobber
            }


        }
        catch {
            Throw "Could not connect to Office 365"
            Write-Host $_
        }
    }
	
    process {
        try {
            $endpoint = Get-MigrationEndPoint
            $mailboxes = Get-Mailbox
            $moverequests = Get-Moverequest 
        }
        catch {
            throw "Couldnt get relevant Information, e.g Mailboxes / MoveRequests or the Hybrid Migrationendpoint !"
            $_
        }

        if ($Group) {
            $users = Get-ADGroupMember -Identity $group | Where-Object ( { $_.Name -notin $namefilter }) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName, emailaddress
        }
        else {
            $users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName
        }
        $i = 0
        Foreach ($user in $users) {
            $i = $i + 1
            Write-Progress -Activity "Generating MoveRequests" -Id 1 -Status "Processing $i/$($users.count) User" -PercentComplete ($i / $users.count * 100)
            try {
                if (($user.Name -notin $moverequests.DisplayName) -or ($user.Name -notin $mailboxes.Name)) {
                    New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain $TargetDeliveryDomain -RemoteCredential $opcred -SuspendWhenReadyToComplete:$($Suspendwhenreadytocomplete) | Out-Null
                    Write-Host 'MoveRequest für' $user.Name' erstellt' -ForeGroundColor Green
                }
            }
            catch {
                Write-Host 'Fehler bei '$user.Name -ForeGroundColor Red
                Write-Host $_
            }
        }
        
    }
    end {
        Get-PSSession | Remove-PSSession
    }
}

function Complete-O365MoveRequest {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        Import-Module ActiveDirectory, MSOnline, CredentialManager

        try {
            if ($null -eq (Get-PSSession)) {
                $cred = Get-StoredCredential -target O365	
                Connect-MsolService -Credential $cred
                $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
                Import-PSSession $s -AllowClobber
            }
            $moverequests = Get-Moverequest | Where-Object { $_.Status -eq "Autosuspended" }
        }

        catch {
            Throw "Could not connect to Office 365"
            Write-Host $_
        }
    }
    
    process {
        $chosen = $moverequests | Out-GridView -PassThru

        foreach ($c in $chosen) {
            Get-Moverequest $c.DisplayName | Set-Moverequest -SuspendWhenReadyToComplete:$False -CompleteAfter (Get-Date)
            Get-Moverequest $c.DisplayName | Resume-Moverequest
        }
    }
    end {
        Get-PSSession | Remove-PSSession
    }
}

function Set-O365MailboxSettings {
    [CmdletBinding()]
    param ()
    
    begin {
        Import-Module ActiveDirectory, MSOnline, CredentialManager
        try {
            if ($null -eq (Get-PSSession)) {
                $cred = Get-StoredCredential -target O365	
                Connect-MsolService -Credential $cred
                $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
                Import-PSSession $s -AllowClobber
            }
        }

        catch {
            Throw "Could not connect to Office 365"
            Write-Host $_
        }
    }
    
    process {
        $mailboxes = Get-Mailbox
        $i = 0
        foreach ($mailbox in $mailboxes) {
            $i = $i + 1
            Write-Progress -Activity "Adding Regionalsettings on Mailboxes" -Id 1 -Status "Processing $i/$($mailboxes.count) mailboxes" -PercentComplete ($i / $mailboxes.count * 100)
            Set-MailboxRegionalConfiguration -Identity $mailbox.UserPrincipalName -Language De-de -DateFormat "dd.MM.yyyy" -TimeZone "W. Europe Standard Time" | Out-Null
        }
    }
    
    end {
        Get-PSSession | Remove-PSSession
    }
}	

function Connect-O365 {
    [CmdletBinding()]
    param (
        [string]$Target = "O365",
        [switch]$Authenticate = $False
    )
        
    begin {}
    
    process {
        try {
            
                Write-Progress -Activity "Initializing Connection" -Status "Fetching Credentials"
                Start-Sleep 3
                if ($Authenticate) {
                    $cred = Get-Credential
                }
                else{
                    $cred = Get-StoredCredential -target $Target	
                }
                Write-Progress -Activity "Initializing Connection" -Status "Connecting to Office 365" 
                Connect-MsolService -Credential $cred
                $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
                Import-PSSession $s -AllowClobber
                Write-Progress -Activity "Initializing Connection" -Status "Connecting to Exchange Online" 
                Connect-ExchangeOnline -Credential $cred -ShowBanner:$false
                Clear-Host
                Write-Host "Successfully connected to Office 365 and Exchange Online" -ForegroundColor Green
            }

        catch {
            Throw "Could not connect to Office 365" 
            Write-Host $_
        }
    }
    
    end {}
}

function Add-O365License {
    [CmdletBinding()]
    param (
        [switch]$Connect
    )
    
    begin {
        if ($Connect) {
            Connect-O365 
        }
    }
    
    process {
        try {
            $users = Get-MsolUser
            $licenses = Get-MsolAccountSku

            $chosenusers = $users | Sort-Object UserPrincipalName | Out-GridView -PassThru -Title "Choose User(s) you want to assign a license"
            $chosenlicenses = $licenses | Sort-Object AccountSkuId | Out-GridView -PassThru -Title "Choose the license(s) you want to assign" 

            $UsersExistingLicenses = @()
            
            foreach ($user in $chosenusers) {
                $u = get-msoluser -UserPrincipalName $user.UserPrincipalName | Select-Object UserPrincipalName, DisplayName, Licenses, FirstName, LastName 
                
                $UsersExistingLicense = [PSCustomObject]@{
                    UserName = $u.DisplayName
                }

                for ($z = 0; $z -lt $u.Licenses.AccountSkuId.Length; $z++ ) {
                    $UsersExistingLicense | Add-Member -type NoteProperty -Name "Lizenz$($z)" -Value $u.Licenses.AccountSkuId[$z]
                }

                $UsersExistingLicenses += $UsersExistingLicense
            }

            $UsersExistingLicenses | Out-GridView -Title "These License are currently assigned to the Users"

            $i = 0
            foreach ($user in $chosenusers) {
                Write-Progress -Activity "Processing User" -Id 1 -Status "Processing $i/$($chosenusers.count) User(s)" -PercentComplete ($i / $chosenusers.count * 100)
                foreach ($license in $chosenlicenses) {
                    $t = $t + 1
                    Write-Progress -Activity "Assigning License(s)" -Id 2 -Status "Processing $t/$($chosenlicenses.count) license(s)" -PercentComplete ($t / $chosenlicenses.count * 100)
                    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $license.AccountSkuId
                }
            }
        }
        catch {
            Write-Host "Could assign license" $license.AccountSkuId "to" $user.UserPrincipalName
            Write-Host $_
        }
    }
    
    end {
        Get-PSSession | Remove-PSSession
    }
}

