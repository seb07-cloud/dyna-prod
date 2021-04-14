[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Env:PSModulePath = $Env:PSModulePath + ";$($HOME)\Documents\PowerShell\Modules"