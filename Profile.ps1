
. "$PSSCRIPTROOT\Functions\Functions.ps1"
. "$PSSCRIPTROOT\Functions\Git-Functions.ps1"
. "$PSSCRIPTROOT\PS Personal\Functions\Personal-Functions.ps1"
    
# Domain Admin Check
If ((hostname) -ne "SCARTER3-1170") {
    If ($(Get-ADGroupMember -Identity "Domain Admins" -Recursive | Select -ExpandProperty SAMAccountName) -contains (([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8)) {
            . "$PSSCRIPTROOT\Functions\DA-Functions.ps1"
    }
}
#Elevated Shell Check
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    . "$PSSCRIPTROOT\Functions\Admin-Functions.ps1"
}

SC-LoadModule PSReadline
Set-PSReadLineOption -PredictionSource History
Set-PSReadLineOption -PredictionViewStyle ListView
Set-PSReadLineOption -EditMode Windows
Set-PSReadLineKeyHandler -key Enter -Function ValidateAndAcceptLine

SC-LoadModule Terminal-Icons

Oh-My-Posh init pwsh --config $PSSCRIPTROOT\OMP\my.omp.json | Invoke-Expression
Enable-Poshtooltips
Write-Host