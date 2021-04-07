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