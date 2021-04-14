<#
.SYNOPSIS
    On Prem Exchange / Hybrid Exchange Toolkit
.DESCRIPTION
    On Prem Exchange / Hybrid Exchange Toolkit
.OUTPUTS
    Nothing but magic
.EXAMPLE
    Add-OnPremRoutingaddress -Tenant "customer.mail.onmicrosoft.com"
    Remove-OnPremMailDomain -Domain "Test.com"

.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   30.03.2021
       
    Changelog:
		1.0             Initial Release
#>

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
        [ValidateSet("Group","OU","All")]
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
            "OU" {$users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName}
            "All" {$users = Get-Mailbox }
            Default {}
        }

        if ($Scope -eq "Group") {
            $users = Get-ADGroupMember -Identity $group | Where-Object ( { $_.Name -notin $namefilter }) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName
        }
        if ($Scope -eq "OU"){
            $users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName
        }
        if ($Scope -eq "All") {
            $users = Get-Mailbox 
        }
        if ($null -eq $Scope){
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