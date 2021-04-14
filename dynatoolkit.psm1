<#
.SYNOPSIS
    The Swiss Army Knive of Hybrid Migration 
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

function Install-DynaProfile {
    [CmdletBinding()]
    param(
        [ValidateSet('AllUsersAllHosts', 'AllUsersCurrentHost', 'CurrentUserAllHosts', 'CurrentUserCurrentHost')]
        [string]$Scope = "CurrentUserAllHosts",
        [Uri]$URL = 'https://raw.githubusercontent.com/seboo30/Productive/main/Microsoft.PowerShell_profile.ps1'
    )
    
    begin {}
    
    process {
        $profile_dir = Split-Path $PROFILE.$Scope
        $profile_file = $profile.$Scope
        $request = Invoke-WebRequest $URL -UseBasicParsing  -ContentType "text/plain; charset=utf-8"

        if (-not (Test-Path $profile_dir)) {
            New-Item -Path $profile_dir -ItemType Directory | Out-Null
            Write-Verbose "Created new profile directory: $profile_dir"
        }

        [IO.File]::WriteAllLines($profile_file, $request.Content)
        Write-Verbose "Wrote profile file: $profile_file with content from: $URL"
        & $profile_file -Verbose
    }
    
    end {}
}

function Install-DynaModule {
    [CmdletBinding()]
    param (
        [Uri]$URL = 'https://raw.githubusercontent.com/seboo30/Productive/main/dynatoolkit.psm1',
        [String]$Dynamodulepath = "C:\pccfg\Scripts\Modules\DynaToolKit"
    )
    
    begin {
        $module_file = "$Dynamodulepath\dynatoolkit.psm1"
        $request = Invoke-WebRequest $URL -UseBasicParsing  -ContentType "text/plain; charset=utf-8"
    }
    
    process {
        if (!(Test-Path $dynamodulepath )) {
            New-Item -Type Directory -Path $dynamodulepath -Force 
            New-Item -Path $dynamodulepath -ItemType Directory -Name DynaToolKit -Force
        }

        [IO.File]::WriteAllLines($module_file, $request.Content)
        Write-Verbose "Wrote Module file: $module_file with content from: $URL"
        & $module_file -Verbose
    }
    
    end {
        
    }
}

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
        [switch]$MFA = $False
    )
        
    begin {}
    
    process {
        try {
            if (!$MFA) {
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
            else {
                Write-Progress -Activity "Initializing Connection" -Status "Fetching Credentials"
                $cred = Get-Credential
                Write-Progress -Activity "Initializing Connection" -Status "Connecting to Office 365" 
                Connect-MsolService -Credential $cred
                $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
                Import-PSSession $s -AllowClobber
                Write-Progress -Activity "Initializing Connection" -Status "Connecting to Exchange Online" 
                Connect-ExchangeOnline -Credential $cred -ShowBanner:$false
                Clear-Host
                Write-Host "Successfully connected to Office 365 and Exchange Online" -ForegroundColor Green
            }
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

function New-O365SingleMoveRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$User,
        [Parameter(Mandatory = $true)]
        [string]$RoutingDomain,
        [switch]$Connect
    )
    begin {

        if ($Connect) {
            Connect-O365 
        }
        Write-Host "Getting Office 365 Migration Endpoint ......" -ForegroundColor Green
        $endpoint = Get-MigrationEndPoint

    }
    
    process {
        {
            Try {
                New-MoveRequest -Erroraction Stop -Identity $User -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain $RoutingDomain -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true
                Write-Host 'MoveRequest für ' $User ' erstellt' -ForeGroundColor Green
            }
            Catch {
                Write-Host 'Fehler bei ' $User -ForeGroundColor Red
            }
        }	
                
    }
    
    end {
        Remove-PSSession $s -Confirm:$False
    }
}
function Add-O365Routingaddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Tenant
    )
	
    begin {}
	
    process {
        try {
            foreach ($missing in (Get-Mailbox -Filter { emailaddresses -notlike "*microsoft.com" })) { 	
                $upn = $missing.Userprincipalname.Split("@")
                $mail = $upn[0] + "@" + $Tenant
                Set-Mailbox $missing -EmailAddresses @{add = $mail } -WarningAction SilentlyContinue
                Write-Host "Added Mailaddress $mail to $missing" -ForegroundColor Green
                $i = $i + 1 
            }
        }
        catch {
            Write-Host "Couldnt add" $mail "to" $missing $_
        }

    }
	
    end {
        if ($i -gt 0) { Write-Host "Added Routingaddresses on $i Mailboxes" -ForegroundColor Green }
    }
}
function Add-OnPremRoutingaddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Tenant
    )
	
    begin {}
	
    process {
        try {
            foreach ($missing in (Get-Mailbox -Filter { emailaddresses -notlike "*microsoft.com" })) { 	
                $upn = $missing.Userprincipalname.Split("@")
                $mail = $upn[0] + "@" + $Tenant
                Set-Mailbox $missing -EmailAddresses @{add = $mail } -WarningAction SilentlyContinue
                Write-Host "Added Mailaddress $mail to $missing" -ForegroundColor Green
                $i = $i + 1 
            }
        }
        catch {
            Write-Host "Couldnt add" $mail "to" $missing $_
        }

    }
	
    end {
        if ($i -gt 0) { Write-Host "Added Routingaddresses on $i Mailboxes" -ForegroundColor Green }
    }
}

function Remove-OnPremMailDomain {
    param (
        [Parameter(Mandatory)]
        [string]$Domain
    )
    begin {}
    
    process {
        try {
            $users = get-mailbox | Where-Object { $_.emailaddresses -like $Domain }
            foreach ($user in $users) {
                $addresses = (get-mailbox $user.alias).emailaddresses
                $fixedaddresses = $addresses | Where-Object { $_.proxyaddressstring -notlike $Domain }
                set-mailbox $user.alias -emailaddresses $fixedaddresses
                Write-Host "Removed Maildomain from" $user.Name -ForeGroundColor Green
            }

        }
        catch {
            Write-Host $_
        }
    }
    end {}
}

function New-OnPremMoveRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet("Group", "OU", "All")]
        [string]$Scope,
        [Parameter(Mandatory)]
        [string]$TargetDeliveryDomain,
        [switch]$Suspendwhenreadytocomplete
    )
	
    begin {
        $namefilter = @("Mailbox1", "Discovery*", "Administrator", "Health*" )
    }
	
    process {
        try {
            $mailboxes = Get-Mailbox
            $moverequests = Get-Moverequest 
        }
        catch {
            throw "Couldnt get relevant Information, e.g Mailboxes / MoveRequests !"
            $_
        }

        switch ($Scope) {
            "Group" { $users = Get-ADGroupMember -Identity $group | Where-Object ( { $_.Name -notin $namefilter }) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName }
            "OU" { $users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName }
            "All" { $users = Get-Mailbox }
            Default {}
        }

        if ($Scope -eq "Group") {
            $users = Get-ADGroupMember -Identity $group | Where-Object ( { $_.Name -notin $namefilter }) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName
        }
        if ($Scope -eq "OU") {
            $users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName
        }
        if ($Scope -eq "All") {
            $users = Get-Mailbox 
        }
        if ($null -eq $Scope) {
            throw "Scope no defined"
        }
        $i = 0
        Foreach ($user in $users) {
            $i = $i + 1
            Write-Progress -Activity "Generating MoveRequests" -Id 1 -Status "Processing $i/$($users.count) User" -PercentComplete ($i / $users.count * 100)
            try {
                if (($user.Name -notin $moverequests.DisplayName) -or ($user.Name -notin $mailboxes.Name)) {
                    New-MoveRequest -Identity $user.userPrincipalName -TargetDatabase $targetdatabase -SuspendWhenReadyToComplete:$Suspendwhenreadytocomplete | Out-Null
                    Write-Host 'MoveRequest für' $user.Name' erstellt' -ForeGroundColor Green
                }
            }
            catch {
                Write-Host 'Fehler bei '$user.Name -ForeGroundColor Red
                Write-Host $_
            }
        }
    }
    end {}
}

function Complete-OnPremMoveRequest {
    [CmdletBinding()]
    param ()
    
    begin {
        $moverequests = Get-Moverequest | Where-Object { $_.Status -eq "Autosuspended" }
        
    }
    process {
        $chosen = $moverequests | Out-GridView -PassThru

        foreach ($c in $chosen) {
            Get-Moverequest $c.DisplayName | Set-Moverequest -SuspendWhenReadyToComplete:$False -CompleteAfter (Get-Date)
            Get-Moverequest $c.DisplayName | Resume-Moverequest
        }
    }
    end {}
}