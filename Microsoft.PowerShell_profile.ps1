[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#$Env:PSModulePath = $Env:PSModulePath + ";" + $dynamodulepath

try {
    Import-Module C:\pccfg\Scripts\Modules\DynaToolKit\dynatoolkit.psm1
}
catch {
    Write-Host "Module not found, try <Install-DynaModule>"
}

$modules = @(
    "ExchangeOnlineManagement"
    "MSOnline"
    "CredentialManager"
    "Orca"
)

foreach ($module in $modules) {
    if ($Null -eq (Get-InstalledModule -Name $module)) {
        Write-Host "Installing $module" -ForegroundColor Green
        Install-Module -Name $module -Confirm:$False -Force #-Scope CurrentUser
    }
}

if (!(Test-Path $profile)){
    New-Item –Path $Profile –Type File –Force
}

