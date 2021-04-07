[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (!(Test-Path $profile)) { 
    try {
        New-Item -ItemType File -Path $PROFILE -Force
        if (!(Test-Path $HOME\Documents\PowerShell\Modules )) {
            New-Item -Type Directory -Path $HOME\Documents\PowerShell\Modules -Force
            '$Env:PSModulePath = $Env:PSModulePath + ";$($HOME)\Documents\PowerShell\Modules"' | Add-Content $profile
            New-Item -Path "$($HOME)\Documents\PowerShell\Modules" -ItemType Directory -Name DynaPowershell -Force
        }
    }
    catch {
        Write-Host $_
    }
}

try {
    Invoke-WebRequest -Method Get -Uri "https://github.com/seboo30/Productive/archive/refs/heads/main.zip" -OutFile .\main.zip
    if (!(Test-Path "$($HOME)\Documents\PowerShell\Modules\DynaPowershell")) {
        New-Item -Path "$($HOME)\Documents\PowerShell\Modules" -ItemType Directory -Name DynaPowershell -Force
        Expand-Archive .\main.zip -DestinationPath "$($HOME)\Documents\PowerShell\Modules\DynaPowershell" 
        Remove-Item .\main.zip -Force
        Add-Content -Path $profile -Value "Import-Module $($HOME)\Documents\PowerShell\Modules\DynaPowershell\moverequest.psm1" 
    }
}
catch {
    Write-Host $_
}

$ModuleEXO = Get-Module -Name ExchangeOnlineManagement
$ModuleMSOL = Get-Module -Name MSOnline
$ModuleCredMan = Get-Module -Name CredentialManager
$ModuleOrca = Get-Module -Name Orca

If ($Null -eq $ModuleEXO){
    Install-Module -Name ExchangeOnlineManagement -Confirm:$False
}
if ($null -eq $ModuleMSOL){
    Install-Module -Name MSOnline -Confirm:$False
}
if ($null -eq $ModuleCredMan){
    Install-Module -Name CredentialManager -Confirm:$False
}
if ($null -eq $ModuleOrca){
    Install-Module -Name Orca -Confirm:$False
}
