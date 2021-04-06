if (!(Test-Path $profile)) { 
    New-Item -ItemType File -Path $PROFILE -Force
    if (!(Test-Path $HOME\Documents\PowerShell\Modules )) {
        New-Item -Type Directory -Path $HOME\Documents\PowerShell\Modules
        $Env:PSModulePath = $Env:PSModulePath + ";$($HOME)\Documents\PowerShell\Modules" | Add-Content -Path $profile 
        Copy-Item .\DynaPowershell -Destination $HOME\Documents\PowerShell\Modules\ -Recurse
    }
}
