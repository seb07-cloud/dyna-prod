[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (!(Test-Path $profile)) { 
    try {
        New-Item -ItemType File -Path $PROFILE -Force
        if (!(Test-Path $HOME\Documents\PowerShell\Modules )) {
            New-Item -Type Directory -Path $HOME\Documents\PowerShell\Modules -Force
            'Import-Module $Env:PSModulePath = $Env:PSModulePath + ";$($HOME)\Documents\PowerShell\Modules"' | Add-Content $profile
            '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12' | Add-Content $profile
            New-Item -Path "$($HOME)\Documents\PowerShell\Modules" -ItemType Directory -Name DynaPowershell -Force
        }
    }
    catch {
        Write-Host $_
    }
}

try {
    Invoke-WebRequest -Method Get -Uri "https://github.com/seboo30/Productive/archive/refs/heads/main.zip" -OutFile .\main.zip
    $mymodulepath = "$($HOME)\Documents\PowerShell\Modules"
    try {
        New-Item -Path $mymodulepath -ItemType Directory -Name DynaPowershell -Force -ErrorAction SilentlyContinue
        New-Item -Path "C:\" -ItemType Directory -Name Temp -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "Couldnt create Folder or Folder already exists, proceeding ...."
    }
    Expand-Archive .\main.zip -DestinationPath "C:\Temp" 
    Remove-Item .\main.zip -Force
    Get-ChildItem "C:\Temp\Productive-main" | Copy-Item -Destination (Join-Path -path $mymodulepath -childpath "DynaPowershell")
    Add-Content -Path $profile -Value (Join-Path -path $mymodulepath -childpath "msoltoolkit.psm1")
    Import-Module (Join-Path -path $mymodulepath -childpath "DynaPowershell\msoltoolkit.psm1")

    if(Test-Path -path (Join-Path -path $mymodulepath -childpath "DynaPowershell\msoltoolkit.psm1")){
        Remove-Item "C:\Temp\Productive-main" -Force -Recurse
    }
}
catch {
    Write-Host $_
}

If ($Null -eq (Get-Module -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Confirm:$False -Force
}
if ($null -eq (Get-Module -Name MSOnline)) {
    Install-Module -Name MSOnline -Confirm:$False -Force
}
if ($null -eq (Get-Module -Name CredentialManager)) {
    Install-Module -Name CredentialManager -Confirm:$False -Force
}
if ($null -eq (Get-Module -Name Orca)) {
    Install-Module -Name Orca -Confirm:$False -Force
}
