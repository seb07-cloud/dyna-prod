[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$modules = @(
    "ExchangeOnlineManagement"
    "MSOnline"
    "CredentialManager"
    "Orca"
)

foreach ($module in $modules) {
    if ($Null -eq (Get-Module -Name $module)) {
        Install-Module -Name $module -Confirm:$False -Force -Scope CurrentUser
    }
}