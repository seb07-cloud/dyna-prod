<#
.SYNOPSIS
    Create Policies for Conditional Access
.DESCRIPTION
    Create Policies for Conditional Access 
.OUTPUTS
    Nothing but magic
.EXAMPLE
    Module required :
        AzureAD 

    .\New-ConditionalAccess.ps1 

.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   24.11.21
       
    Changelog:
		1.0            Initial Release
#>


function IsValidEmail { 
    param([string]$Email)
    $Regex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'

   try {
        $obj = [mailaddress]$Email
        if($obj.Address -match $Regex){
            return $True
        }
        return $False
    }
    catch {
        return $False
    } 
}

# Read Adminuser

Write-Host "Type in the UserPrincipalName of the Adminuser which you want to protect:" -ForegroundColor Green
$Admin = Read-Host -Prompt "Admin Email Address"

while (!(IsValidEmail -Email $Admin)){
    Write-Host "Type in a valid Email Address !" -ForegroundColor Red
    $Admin = Read-Host -Prompt "Admin Email Address"
}

# New Trusted Subnet

$Dyna = New-Object -TypeName Microsoft.Open.MSGraph.Model.IpRange
$Dyna.cidrAddress = '194.183.133.226/32'

$Location = New-AzureADMSNamedLocationPolicy -OdataType "#microsoft.graph.ipNamedLocation" -DisplayName 'Dyna' -IsTrusted $True -IpRanges $Dyna

# Create Policy Settings

$CAConditions = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessConditionSet
$CAConditions.Applications = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessApplicationCondition
$CAConditions.Applications.IncludeApplications = 'All'

# Create Conditions

$CAConditions.Users = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessUserCondition
$CAConditions.Users.IncludeUsers = $Admin
$CAConditions.Locations = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessLocationCondition
$CAConditions.Locations.IncludeLocations = $Location.Id

$CAControls = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessGrantControls
$CAControls._Operator = "OR"
$CAControls.BuiltInControls = "Mfa"

# New CA Policy

New-AzureADMSConditionalAccessPolicy -DisplayName "MFA_Admin" -State "Enabled" -Conditions $CAConditions -GrantControls $CAControls