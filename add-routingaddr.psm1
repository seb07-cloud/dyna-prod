<#
.SYNOPSIS
    Add Microsoft Routing Addresses
.DESCRIPTION
    Script was designed to provision a Active Directory User from Ally CSV Export 
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\Add-Routingaddress -Tenant "customer.mail.onmicrosoft.com"

.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   30.03.2021
       
    Changelog:
		1.0             Initial Release
#>

function Add-Routingaddress {
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

	