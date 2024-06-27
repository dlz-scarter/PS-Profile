##############################################################################
## ACTIVE DIRECTORY FUNCTIONS
##############################################################################

##############################################################################
## SCCM FUNCTIONS
##############################################################################

##  Connect to SCCM
Function SC-ConnectSCCM () {
    # Site configuration
    $SiteCode = "DLZ" # Site code 
    $ProviderMachineName = "SCCM.DLZCORP.COM" # SMS Provider machine name

    # Customizations
    $initParams = @{}
    #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
    #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

    # Do not change anything below this line
    
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }

    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }
    
    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
}

##  Update Content of Single DP
##  IMPROVED VERSION IN DA-Functions (SC-UpdateSCCMDP)
#Function SC-UpdateContentForSingleDP{
#    param (
#    [Parameter(Mandatory=$true)]$SiteCode,
#    [Parameter(Mandatory=$true)]$DPFQDN
#    )
#    $Failures = Get-WmiObject -Namespace root\sms\site_$SiteCode -Class sms_packagestatusDistPointsSummarizer | Where-Object State -GT 1 | Where-Object SourceNALPath -Match $DPFQDN 
# 
#    foreach ($Failure in $Failures) { 
#        $s = New-PSSession -ComputerName $DPFQDN
#        $PackageID = $Failure.PackageID 
#        Write-Output "Failed PackageID: $PackageID"
#        Invoke-Command -Session $s -Scriptblock {Get-SmbOpenFile | Where-Object -Property sharerelativepath -match "$PackageID" | Close-SmbOpenFile -force}
#        Remove-PsSession $s
# 
#        $DP = Get-WmiObject -Namespace root\sms\site_$SiteCode -Class sms_distributionpoint | Where-Object ServerNalPath -match $DPFQDN | Where-Object PackageID -EQ $PackageID 
#        $DP.RefreshNow = $true 
#        $DP.put() 
#    }
#}

##############################################################################
## MISC FUNCTIONS
##############################################################################

## Search for silent uninstall strings
Function SC-GetUninstallString{
    param (
        [Parameter(Mandatory = $true)]$searchterm
    )
    $results = (Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -match $searchterm } | Select-Object -Property DisplayName, UninstallString)
    return $results
}

## Clean up space on C:
Function SC-ReclaimDiskSpace {
    $diskInfo = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = $diskInfo.FreeSpace / 1GB
    $totalSpaceGB = $diskInfo.Size / 1GB
    $freeSpacePercentage = ($diskInfo.FreeSpace / $diskInfo.Size) * 100
    
    $diskSpaceReport = "Before: {0:N2} GB Free ({1:N2}%) " -f $freeSpaceGB, $freeSpacePercentage
    
    # Stop WUAUSERV and BITS services
    Stop-Service -Name WUAUSERV, BITS -Force
    
    $tempFolders = @(
        "C:\Windows\Temp\*",
        "C:\Users\*\Appdata\Local\Temp\*",
        "C:\Windows\CCMCache\*",
        "C:\Windows\SoftwareDistribution\*"
    )
    Remove-Item $tempFolders -Force -Recurse -ErrorAction SilentlyContinue
    
    # Start WUAUSERV and BITS services
    Start-Service -Name WUAUSERV, BITS
    
    # Refresh disk information
    $diskInfo = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = $diskInfo.FreeSpace / 1GB
    $freeSpacePercentage = ($diskInfo.FreeSpace / $diskInfo.Size) * 100
    
    $diskSpaceReport += "`nAfter:  {0:N2} GB Free ({1:N2}%) " -f ($diskInfo.FreeSpace / 1GB), ($diskInfo.FreeSpace / $diskInfo.Size * 100)
    
    Write-Output $diskSpaceReport
}

Function Reclaim-DiskSpace {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$ComputerName
    )

    if (-not $ComputerName) {
		write-output "Setting computername to $env:computername"
        $ComputerName = $env:COMPUTERNAME
    }

    $diskInfo = Get-WmiObject -ComputerName $ComputerName -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = $diskInfo.FreeSpace / 1GB
    $totalSpaceGB = $diskInfo.Size / 1GB
    $freeSpacePercentage = ($diskInfo.FreeSpace / $diskInfo.Size) * 100
    
    $diskSpaceReport = "Before: {0:N2} GB Free ({1:N2}%) " -f $freeSpaceGB, $freeSpacePercentage
    
    Stop-Service -ComputerName $ComputerName -Name WUAUSERV, BITS -Force
    
    $tempFolders = @(
        "C:\Windows\Temp\*",
        "C:\Users\*\Appdata\Local\Temp\*",
        "C:\Windows\CCMCache\*",
        "C:\Windows\SoftwareDistribution\*"
    )
    
    Remove-Item -ComputerName $ComputerName -Path $tempFolders -Force -Recurse -ErrorAction SilentlyContinue
    
    Start-Service -ComputerName $ComputerName -Name WUAUSERV, BITS
    
    $diskInfo = Get-WmiObject -ComputerName $ComputerName -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = $diskInfo.FreeSpace / 1GB
    $freeSpacePercentage = ($diskInfo.FreeSpace / $diskInfo.Size) * 100
    
    $diskSpaceReport += "`nAfter:  {0:N2} GB Free ({1:N2}%) " -f ($diskInfo.FreeSpace / 1GB), ($diskInfo.FreeSpace / $diskInfo.Size * 100)
    
    Write-Output $diskSpaceReport
}