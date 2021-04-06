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


.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   30.03.2021

    Changelog:
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
	
            Connect-MsolService -Credential $cred
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
            Import-PSSession $s -AllowClobber


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
            $users = Get-ADGroupMember -Identity $group | Where-Object ( { $_.Name -notin $namefilter }) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName
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
            $cred = Get-StoredCredential -target O365	
            Connect-MsolService -Credential $cred
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
            Import-PSSession $s -AllowClobber

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
    param (
        
    )
    
    begin {
        Import-Module ActiveDirectory, MSOnline, CredentialManager

        try {
            $cred = Get-StoredCredential -target O365	
            Connect-MsolService -Credential $cred
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
            Import-PSSession $s -AllowClobber
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
        
    }
}	

function Connect-O365 {
    [CmdletBinding()]
    param (
        [string]$Target = "O365"
    )
        
    begin {}
    
    process {
        try {
            Write-Progress -Activity "Initializing Connection" -Status "Fetching Credentials"
            Start-Sleep 3
            $cred = Get-StoredCredential -target $Target	
            Write-Progress -Activity "Initializing Connection" -Status "Connecting to Office 365" 
            Connect-MsolService -Credential $cred
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
            Import-PSSession $s -AllowClobber
            Write-Progress -Activity "Initializing Connection" -Status "Connecting to Exchange Online" 
            Connect-ExchangeOnline -Credential (Get-StoredCredential -Target $Target) -ShowBanner:$false
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
