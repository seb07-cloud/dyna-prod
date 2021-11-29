<#
.SYNOPSIS
    Create Policies for Conditional Access
.DESCRIPTION
    Create Policies for Conditional Access 
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\New-ConditionalAccess.ps1 

.NOTES
    Module Required:    AzureAD 

    Author:             Sebastian Wild	
    Email:              sebastian.wild@dynabcs.at
    Company:            DynaBCS Informatik
    Date:               24.11.21
       
    Changelog:
    1.0                 Initial Release
#>

Import-Module PSWriteColor
Import-Module AzureAD

$error.clear()

if (!(Get-PSSession)){
    connect-azuread
}
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

Write-Host "Type in the UserPrincipalName of the adminuser which you want to protect:" -ForegroundColor Green
$Admin = Read-Host -Prompt "admin emailaddress"

while (!(IsValidEmail -Email $Admin)){
    Write-Host "Type in a valid emailaddress !" -ForegroundColor Red
    $Admin = Read-Host -Prompt "admin emailaddress"
}

while (!(Get-AzureAdUser -SearchString $Admin)) {
    Write-Host "User not found ! Type in a valid user/emailaddress" -ForegroundColor Red
    $Admin = Read-Host -Prompt "admin emailaddress"
}

$AdminUser = Get-AzureAdUser -SearchString $Admin

# New Trusted Subnet

$Dyna = New-Object -TypeName Microsoft.Open.MSGraph.Model.IpRange
$Dyna.cidrAddress = '194.183.133.224/27'

$Location = New-AzureADMSNamedLocationPolicy -OdataType "#microsoft.graph.ipNamedLocation" -DisplayName 'Dyna' -IsTrusted $True -IpRanges $Dyna

# Create Policy Settings

$CAConditions = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessConditionSet
$CAConditions.Applications = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessApplicationCondition
$CAConditions.Applications.IncludeApplications = 'All'

# Create Conditions

$CAConditions.Users = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessUserCondition
$CAConditions.Users.IncludeUsers = $AdminUser.ObjectId
$CAConditions.Locations = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessLocationCondition
$CAConditions.Locations.IncludeLocations = 'All'
$CAConditions.Locations.ExcludeLocations = $Location.Id

$CAControls = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessGrantControls
$CAControls._Operator = "OR"
$CAControls.BuiltInControls = "Mfa"

# Create CA Policy

New-AzureADMSConditionalAccessPolicy -DisplayName "MFA_Admin" -State "Enabled" -Conditions $CAConditions -GrantControls $CAControls

if (!($error)){
    Write-Color -Text "Conditional Access Policy for",
        " $($Admin)",
        " and an exclude for the following subnet", 
        " $($Dyna.cidrAddress)",
        " is created !" -Color Green,DarkYellow,Green,DarkYellow,Green
}