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
			$users = Get-ADGroupMember -Identity $group | Where(( { $_.Name -notin $namefilter })) | ForEach-Object { (Get-ADUser $_.SamAccountName -Properties * | where-object { $null -ne $_.msExchRecipientTypeDetails }) } | Select-Object Name, userPrincipalName

			Foreach ($user in $users) {
				try {
					if ($user.Name -notin $moverequests.DisplayName) {
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
		else {
			$users = Get-ADUser -SearchBase $ou -Properties * | Where-Object { $_.SamAccountName -notin $namefilter } -and { $null -ne $_.msExchRecipientTypeDetails } | Select-Object Name, UserPrincipalName
			
			Foreach ($user in $users) {
				try {
					if ($user.Name -notin $moverequests.DisplayName) {
						New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain $TargetDeliveryDomain -RemoteCredential $opcred -SuspendWhenReadyToComplete:$($Suspendwhenreadytocomplete) | Out-Null
						Write-Host 'MoveRequest für '$user.Name' erstellt' -ForeGroundColor Green
					}
				}
				catch {
					Write-Host 'Fehler bei '$user.Name -ForeGroundColor Red
					Write-Host $_
				}
			}
		}
	}
}
end {
	Get-PSSession | Remove-PSSession
}


		
	
