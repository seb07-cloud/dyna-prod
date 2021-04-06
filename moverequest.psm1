<#
.SYNOPSIS
    Generate new MoveRequests from AD Group Membership or OU
.DESCRIPTION
    Generate new MoveRequests from AD Group Membership or OU
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\New-O365MoveRequest -Group "gr_O365-Sync"
	.\New-O365MoveRequest -OU "OU=Contoso-Groups,DC=Contoso,DC=local"

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
            $importresults = Import-PSSession $s -AllowClobber

            $moverequests = Get-Moverequest 
        }
        catch {
            Throw "Could not connect to Office 365"
            Write-Host $_
        }
    }
	
    process {
        try {
            $endpoint = Get-MigrationEndPoint
        }
        catch {
            throw "No Migrationendpoint found !"
            $_
        }

        if ($Group) {
            $users = Get-ADGroupMember -Identity $group | Where-Object ( { $_.Name -notin $namefilter }) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName

            Foreach ($user in $users) {

                try {
                    Foreach ($r in $moverequests) {

                        if ($r.DisplayName -ne $user.Name) {
	
                            New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain $TargetDeliveryDomain -RemoteCredential $opcred -SuspendWhenReadyToComplete:$($Suspendwhenreadytocomplete) | Out-Null
                            Write-Host 'MoveRequest für ' $user.Name ' erstellt' -ForeGroundColor Green
                        }
                    }
                }
                catch {
                    Write-Host 'Fehler bei '$user.Name -ForeGroundColor Red
                    Write-Host $_
                }
            }
        }
        else {
            $users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName

            try {
                Foreach ($r in $moverequests) {

                    if ($r.DisplayName -ne $user.Name) {

                        New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain $TargetDeliveryDomain -RemoteCredential $opcred -SuspendWhenReadyToComplete:$($Suspendwhenreadytocomplete) | Out-Null
                        Write-Host 'MoveRequest für ' $user.Name ' erstellt' -ForeGroundColor Green
                    }
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
        foreach ($mailbox in (Get-Mailbox)) {
            Set-MailboxRegionalConfiguration -Identity $mailbox.UserPrincipalName -LocalizeDefaultFolderName:$true -Language De-de -DateFormat "dd.MM.yyyy"
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

            Write-Host "Connecting to Office 365, please wait ...." -ForegroundColor Green
            $cred = Get-StoredCredential -target $Target	
            Connect-MsolService -Credential $cred
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
            Import-PSSession $s -AllowClobber
            Clear-Host
            Write-Host "Connecting to Exchange Online, please wait ...." -ForegroundColor Green
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
	
