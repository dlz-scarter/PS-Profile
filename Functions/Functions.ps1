Function SC-AddDLMember {
    Param ([Parameter(Mandatory = $true)]
        [String]
        $DLName,
        $Username)
    Add-DistributionGroupMember -Identity $DLName -Member $Username
}

Function ssh-copy-id([string]$sshHost) {
    ##########################################
    ##  Copy public SSH key to Linux system for passwordless SSH connection
    cat ~/.ssh/id_rsa.pub | ssh "$sshHost" "mkdir -p ~/.ssh && touch ~/.ssh/authorized_keys && chmod -R go= ~/.ssh && cat >> ~/.ssh/authorized_keys"
}

Function SC-Init () {
    If (-not (Get-Module -Name oh-my-posh-core)) {
        Set-ExecutionPolicy Bypass -Scope Process -Force; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://ohmyposh.dev/install.ps1'))
    }
    #Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope currentuser
    $Modules = @(
        'ThreadJob',
        'PSReadLine',
        'Oh-My-Posh',
        'Posh-Git',
        'Terminal-Icons',
        'MicrosoftTeams',
        'ExchangeOnlineManagement',
        'MSOnline',
        'Microsoft.Online.SharePoint.PowerShell',
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Users.Actions')
    ForEach ($Module In $Modules) {
        If (!(Get-Module -ListAvailable -Name $Module)) {
            Write-Output "Attempting to install $Module..."
            Try {
                SC-LoadModule $Module -Scope CurrentUser -AllowClobber -Confirm:$False -Force
            }
            Catch [Exception] {
                $_.message
                Exit
            }
            Clear-Host
        }
    }
}

Function SC-Update() {
    ##########################################
    ##  Download and install modules used in this profile
    $CD = $PWD
    Set-Location -Path $PSSCRIPTROOT
    $UpdateCheck = (Git Status)
    If ($UpdateCheck -notlike "*nothing to commit*") {
        Write-Output "Script out of date, would you like to update?"
        $Response = Read-Host "`nPlease select Y or N"
        While (($Response -match "[YyNn]") -eq $false) {
            $Response = Read-Host "This is a binary situation. Y or N please."
        }

        If ($Response -match "[Yy]") {
            git pull
        }
    }
    Set-Location -Path $PWD
}

Function SC-LoadModule {
    param(
        [parameter(Mandatory=$true)] [string]$Module
    )
    Try {
        Import-Module $Module -ErrorAction 'Stop' -WarningAction 'SilentlyContinue'
    }
    Catch [System.IO.FileNotFoundException] {
        Try {
            Write-Output "$Module not found, attempting to install..."
            Install-Module -Name $Module -Scope CurrentUser -AllowClobber -Confirm:$False -Force -ErrorAction 'Stop'
            Import-Module $Module
        }
        Catch [Exception] {
            $_.message
        }
    }
}

Function Run-Step([string]$Description, [ScriptBlock]$script) {
    ##########################################
    ##  Visual Feedback for loading module
    Write-Host -NoNewline "Loading " $Description.PadRight(20)
    & $script
    Write-Host
}

Function Is-Elevated {
    ##########################################
    ##  Set Titlebar
    $prp = new-object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
    $prp.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

Function SC-GetPubIP {
    ##########################################
    ## Get Public IP Address
    (Invoke-WebRequest http://ifconfig.me/ip).Content
}

Function SC-GetUTC {
    ##########################################
    ##  Get UTC Time
    (Get-Date).ToUniversalTime()
}

Function SC-GenerateStrongPassword ([Parameter(Mandatory = $true)][int]$PasswordLength) {
    ##########################################
    ##  Password Generator
    Add-Type -AssemblyName System.Web
    $PassComplexCheck = $false
    Do
    {
        $newPassword = [System.Web.Security.Membership]::GeneratePassword($PasswordLength, 2)
        If (($newPassword -cmatch "[A-Z\p{Lu}\s]") `
            -and ($newPassword -cmatch "[a-z\p{Ll}\s]") `
            -and ($newPassword -match "[\d]") `
            -and ($newPassword -match "[^\w]")
        )
        {
            $PassComplexCheck = $True
        }
    }
    While ($PassComplexCheck -eq $false)
    Return $newPassword
}

Function SC-GetCSPUsers() {
    ##########################################
    ##  Get members of LIC-CSP O365-E3 group
    Try {
        Write-Output "`nThe requested information is being collected and sorted, please wait...`n"
        $Base = Get-ADGroupMember "Lic-CSP O365-E3" | Get-ADUser -properties WhenCreated, Title, Enabled, LastLogonDate
        $CSPUsers = $Base | Where-Object { $_.Enabled -eq $true }
        $CSPDisabledUsers = $Base | Where-Object { $_.Enabled -eq $false }
        $CSPUsers | Sort-Object Title | Select-Object Name, SAMAccountName, WhenCreated, Title | Format-Table -autosize
        Write-Host -NoNewLine $CSPUsers.count "CSP licenses assigned to active accounts via group membership`n`n"
        If (($Base.count - $CSPUsers.count) -gt 0) {
            Write-Host -NoNewLine ($Base.count - $CSPUsers.count) "CSP license(s) assigned to the following disabled account(s) via group membership:`n"
            $CSPDisabledUsers | Sort-Object LastLogonDate | Select-Object Name, SAMAccountName, LastLogonDate, Title | Format-Table -autosize
        }
    }
    Catch {
        Write-Host "`nAn error occurred.`n"
    }
}

Function SC-GetEAUsers() {
    ##########################################
    ##  Get members of LIC-EA M365-E3 group
    Try {
        Write-Output "`nThe requested information is being collected and sorted, please wait...`n"
        $Base = Get-ADGroupMember "Lic-EA M365-E3" | Get-ADUser -properties WhenCreated,Title,Enabled,LastLogonDate
        $EAUsers = $Base | ? {$_.Enabled -eq $true}
        $EADisabledUsers = $Base | ? {$_.Enabled -eq $false}
        $EAUsers | sort Title| select Name, SAMAccountName, WhenCreated, Title | ft -autosize
        Write-Host -NoNewLine $EAUsers.count "EA licenses assigned to active accounts via group membership`n`n"
        If (($Base.count - $EAUsers.count) -gt 0) {
            Write-Host -NoNewLine ($Base.count - $EAUsers.count) "EA license(s) assigned to the following disabled account(s) via group membership:`n"
            $EADisabledUsers | Sort-Object LastLogonDate | Select-Object Name, SAMAccountName, LastLogonDate, Title | Format-Table -autosize
        }
    }
    Catch {
        write-host "`nAn error occurred.`n"
    }
}

Function SC-GetLAPS() {
    ##########################################
    ##  Get LAPS password of a computer
    param (
    [Parameter(Mandatory = $true)] [String]$ComputerName
    )

    SC-LoadModule ActiveDirectory
    try {
        Get-ADComputer $ComputerName | Out-Null

        $passwordInfo = Get-LapsAdPassword -Identity $computerName -AsPlainText | Select-Object ComputerName,Password,PasswordUpdateTime,ExpirationTimestamp
        If ($passwordInfo -ne $null) {
            If ($passwordInfo.PasswordUpdateTime -ne $null) {
                $passwordInfo.Password | set-clipboard
                Write-Host "`n**********************************************"
                Write-Host "**********************************************"
                Write-Host "**                                          **"
                Write-Host "**  LAPS Username:     .\lapsadmin          **"
                Write-Host "**  LAPS Password:     $($passwordinfo.Password)           **"
                Write-Host "**                                          **"
                Write-Host "**  Last Updated:      $($passwordInfo.PasswordUpdateTime)  **"
                Write-Host "**  Expiration Time:   $($passwordinfo.ExpirationTimestamp)  **"
                Write-Host "**                                          **"
                Write-Host "**********************************************"
                Write-Host "**********************************************`n"
                Write-Host "The password has been copied to the clipboard.`n"
            }
            Else {
                SC-GetOldLAPS $ComputerName
            }
        }
        Else {
            write-host "`nNo LAPS Password set for $($Computer.'CN')`n"
        }
    }
    catch {
        write-host "`nComputer not found.`n"
    }
}

Function SC-GetOLDLAPS() {
    param (
    [Parameter(Mandatory = $true)] [String]$ComputerName
    )

    SC-LoadModule ActiveDirectory
    try {
        $Computer = Get-ADComputer $ComputerName -property *

        if ($Computer.'ms-Mcs-AdmPwd'){
            $strComputerExpTime = $Computer.'ms-MCS-AdmPwdExpirationTime'
            $strComputerPassword | set-clipboard
            if ($strComputerExpTime -ge 0) {$strComputerExpTime = $([datetime]::FromFileTime([convert]::ToInt64($strComputerExpTime)))}
            $strComputerExpTime = "{0:yyyy-MM-dd HH:mm:ss}" -f [datetime]$strComputerExpTime
            Write-Host "`n**********************************************"
            Write-Host "**********************************************"
            Write-Host "**                                          **"
            Write-Host "**  LAPS Username      .\dlzadmin           **"
            Write-Host "**  LAPS Password:     $($Computer.'ms-Mcs-AdmPwd')             **"
            Write-Host "**                                          **"
            Write-Host "**  Last Updated:      N/A                  **"
            Write-Host "**  Expiration Time:   $strComputerExpTime  **"
            Write-Host "**                                          **"
            Write-Host "**********************************************"
            Write-Host "**********************************************`n"
            Write-Host "The password has been copied to the clipboard.`n"
            }
        else{
            write-host "`nNo LAPS Password set for $($Computer.'CN')`n"
        }
    }
    catch {
        write-host "`nAn error occured.`n"
    }
}

Function SC-AddOrgToO365() {
    ##########################################
    ##  Add Org Groups to O365 Group
    ### (ProjectCenter temp workaround)
    param (
        [Parameter(Mandatory = $true)]
        [String]$OrgGroup,
        [Parameter(Mandatory = $true)]
        [String]$O365Group
    )

    #import the Active Directory module if not already up and loaded
    $Module = Get-Module | Where-Object {
        $_.Name -eq 'ActiveDirectory'
    }
    If ($Module -eq $null)
    {
        Write-Host "Loading Active Directory PowerShell Module"
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    }
    Try
    {
        $var = Get-AzureADTenantDetail
        $var = Get-EmailAddressPolicy
    }

    Catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        Write-Host "Establishing connection to Azure..."
        SC-ConnectO365Services
    }
    Catch [System.Management.Automation.CommandNotFoundException]
    {
        Write-Host "Establishing connection to Azure..."
        SC-ConnectO365Services
    }
    try
    {
        Write-Host "`nAdding $OrgGroup members to $(Get-UnifiedGroup $O365Group):`n"

        Get-ADGroupMember  $OrgGroup | Get-ADUser | ForEach {
            Write-Host "   Adding $((Get-ADUser ($_)).name)"
            Add-UnifiedGroupLinks -Identity $O365Group -LinkType Members -Links $_.UserPrincipalName
        }
        Write-Host ""
    }
    catch
    {
        write-host "`nAn error occurred.`n"
    }
}

Function SC-ConnectSCCM () {
    ##########################################
    ##  Connect to SCCM
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
        #New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName -Description "SCCM Site" @initParams | Out-Null
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
}

Function SC-ConnectMicrosoftGraph() {
    $AppID = "22f763d8-9019-477e-84d3-a2ed18db716e"
    $TenantID = "41a4547c-4314-46c1-8ca5-46e010cf3108"
    $ClientSecret = get-content $psscriptroot\..\Creds\MsGraph-$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8))@$(hostname).txt | convertto-securestring

    SC-LoadModule MSAL.PS
    SC-LoadModule Microsoft.Graph.Authentication
    SC-LoadModule Microsoft.Graph.Users
    $Token = Get-MsalToken -TenantId $TenantId -ClientId $AppId -ClientSecret $ClientSecret
    Write-Host "`nConnecting to Microsoft Graph"
    Connect-MgGraph -NoWelcome -AccessToken (ConvertTo-SecureString $Token.AccessToken -AsPlainText -Force)
}

Function SC-ConnectO365Services() {
    ##########################################
    ##  Connect to all O365 Services
    $Username = "scarter@dlz.com"
    $Password = get-content $psscriptroot\..\Creds\$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8))@$(hostname).txt | convertto-securestring
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password
    
    #Microsoft Graph
    Try {
        $var = Get-MGOrganization -ErrorAction Stop
    }
    Catch [System.Management.Automation.RuntimeException]
    {
        SC-ConnectMicrosoftGraph
    }

    #ExchangeOnlineManagement
    Write-Host "Connecting to Exchange Online"
    SC-LoadModule ExchangeOnlineManagement
    Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false -WarningAction 'SilentlyContinue' | Out-Null

    #Security and Compliance Ceter
    Write-Host "Connecting to Security and Compliance Center"
    Connect-IPPSSession -Credential $Credential -ShowBanner:$false -WarningAction 'SilentlyContinue' | Out-Null

    #Teams
    If ($PSVersionTable.PSVersion -like ("5.*")) {
        SC-LoadModule MicrosoftTeams
    }
    Else {
        Import-Module MicrosoftTeams -UseWindowsPowerShell | Out-Null
    }
    Write-Host "Connecting to Microsoft Teams"
    Connect-MicrosoftTeams -Credential $Credential -WarningAction 'SilentlyContinue' | Out-Null

    #Sharepoint
    $orgName="DLZ807"
    Write-Host "Connecting to SharePoint`n"
    SC-LoadModule Microsoft.Online.SharePoint.PowerShell
    Connect-SPOService -Url "https://$orgName-admin.sharepoint.com" -Credential $Credential | Out-Null
}

Function SC-UpdateCred () {
    ##########################################
    ##  Update saved credential (for O365 connections)
    $hostname = hostname
    (Get-Credential).Password | ConvertFrom-SecureString | Out-File $psscriptroot\..\Creds\$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8).ToUpper())@$((hostname).ToUpper()).txt
}

Function SC-UpdateMsGraphClientSecret () {
    ##########################################
    ##  Update saved client secret (for MS Graph connections)
    $hostname = hostname
    (Get-Credential).Password | ConvertFrom-SecureString | Out-File $psscriptroot\..\Creds\MsGraph-$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8).ToUpper())@$((hostname).ToUpper()).txt
}

Function SC-AllowExternal () {
    ##########################################
    ##  Allow external senders to email O365 group
    param (
    [Parameter(Mandatory = $true)] [String]$Project
    )
    Set-UnifiedGroup $Project -RequireSenderAuthenticationEnabled $false
    Get-UnifiedGroup $Project | Select PrimarySMTPAddress, RequireSenderAuthenticationEnabled
}

function SC-DomainAdminCheck{
    ##########################################
    ##  Check if PS is being run as a domain administrator
    If ($(Get-ADGroupMember -Identity "Domain Admins" -Recursive | Select -ExpandProperty SAMAccountName) -contains (([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8)) {
        Return 1
    }
    Else {
        Return 0
    }
}

Function SC-GetDirSize {
    ##########################################
    ##  Get directory size
    Param ($directory,
        $recurse)

    Write-Progress -Activity "SC-GetDirSize" -Status "Reading '$($directory)'"
    $files = $directory | Get-ChildItem -Force -Recurse:$recurse | Where-Object {-not $_.PSIsContainer}
    If ($files) {
        Write-Progress -Activity "SC-GetDirSize" -Status "Calculating '$($directory)'"
        Return ($files | Measure-Object -Sum -Property Length | Select-Object @{Name = "Path"; Expression = {$directory}}, @{Name = "Files"; Expression = {$_.Count; $script:totalcount += $_.Count}}, @{Name = "Size"; Expression = {$_.Sum; $script:totalbytes += $_.Sum}})
    }
    Else {
        Return ("" | Select-Object @{Name= "Path"; Expression = {$directory}},@{Name = "Files"; Expression = {0}},@{Name = "Size"; Expression = {0}})
    }
}

Function SC-ConnectVC {
    ##########################################
    ##  Connect to VCenter Servers

    Write-Host "`nInitializing connection to vSphere Servers.  Please wait..."

    $Username = "scarter@dlz.com"
    $Password = get-content $psscriptroot\..\Creds\$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8))@$(hostname).txt | convertto-securestring
    $Credential = New-Object System.Management.Automation.PSCredential $Username, $Password
    $randomNumber = Get-Random -Minimum 0 -Maximum 2

    #Import-Module VMware.VimAutomation.Core | Out-Null
    Set-PowerCLIConfiguration -Scope User -ParticipateInCeip $true -Confirm:$false | Out-Null
    Set-PowerCLIConfiguration -InvalidCertificateAction ignore -Confirm:$false | Out-Null
    Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null

    Try { Disconnect-ViServer * -Confirm:$false -ErrorAction SilentlyContinue }
    Catch {    }
    Write-Host "`nConnecting to DUBLIN-VCENTER01..."
    Connect-VIServer -Server DUBLIN-VCENTER01.DLZCORP.COM -Force -Protocol https -Credential $Credential | Out-Null

    Write-Host "`nConnecting to INDY-VCENTER01..."
    Connect-VIServer -Server INDY-VCENTER01.DLZCORP.COM -Force -Protocol https -Credential $Credential | Out-Null

    Write-Host "`nConnected to PowerCLI.  For cmdlet reference, visit https://developer.vmware.com/docs/powercli/latest/products/"
}

Function SC-ConnectHV {
    ##########################################
    ##  Connect to Horizon Connection Servers
    Param (
        [Parameter(Mandatory = $true)]
        [String]
        $ConnServer
    )
    $Username = "dlzcorp.com\scarter"
    $Password = get-content $psscriptroot\..\Creds\$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8))@$(hostname).txt | convertto-securestring
    $Credential = New-Object System.Management.Automation.PSCredential $Username, $Password

    #Import-Module VMware.VimAutomation.Core | Out-Null

    Set-PowerCLIConfiguration -Scope User -ParticipateInCeip $false -Confirm:$false | Out-Null
    Set-PowerCLIConfiguration -InvalidCertificateAction ignore -Confirm:$false | Out-Null
    Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null

    Write-Host "`nConnecting to $($ConnServer.ToUpper())..."
    $hvServer = Connect-HVServer -Server $ConnServer -Credential $Credential
    $hvServices = $hvServer.ExtensionData
    $csService = New-Object VMware.Hv.ConnectionServerService
    #$csList = $csService.ConnectionServer_List($hvServices)
    #ForEach ($info In $csList) {
    #    Write-Host $info.id
    #}
}

Function SC-CreateZoomRoom {
    ##########################################
    ##  Convert resource mailbox to Zoom room
    Param (
        [Parameter(Mandatory = $true)]
        [String]
        $RoomAlias
    )
    Write-Host "Granting 'Send As' rights to SVC_ZoomRoomAdmin for the $RoomAlias mailbox..."
    Add-RecipientPermission -Identity $RoomAlias -Trustee "SVC_ZoomRoomAdmin" -AccessRights SendAs
    Write-Host "Granting 'Full Access' rights to SVC_ZoomRoomAdmin for the $RoomAlias mailbox..."
    Add-MailboxPermission -Identity $RoomAlias -User "SVC_ZoomRoomAdmin" -AccessRights FullAccess
    Write-Host "Granting Editor rights to SVC_ZoomRoomAdmin for the $RoomAlias calendar..."
    Add-MailboxFolderPermission $RoomAlias":\Calendar" -User "SVC_ZoomRoomAdmin" -AccessRights editor
    Write-Host "Modifying Calendar Processing settings for the $RoomAlias calendar..."
    Set-CalendarProcessing -Identity $RoomAlias -AllowConflicts $true -BookingWindowInDays 1080 -MaximumDurationInMinutes 1440 -AddOrganizerToSubject $false -OrganizerInfo $true -DeleteAttachments $true -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false
    Write-Host "Updating ResourceCustom field for $RoomAlias..."
    Get-Mailbox $RoomAlias | Set-Mailbox -ResourceCustom @{ add = "Zoom" }
    Get-Mailbox $RoomAlias | Set-Mailbox -ResourceCustom @{ remove = "PolyCom" }
    Write-Host "$RoomAlias is now a Zoom Room.`n"
}

Function SC-GetForwards {
    ##########################################
    ##  Find mailboxes configured to forward to a certain user
    param ([Parameter(Mandatory = $true)]
        [String]
        $ForwardTarget)
    $Recipient = (Get-Mailbox $ForwardTarget).Name
    $Results = Get-Mailbox -resultsize unlimited -Filter { ForwardingAddress -like "*" } | Where-Object{ $_.ForwardingAddress -eq $Recipient } | Select-Object Name
    Write-Host "`nThe following mailboxes forward to $($Recipient):"
}


##########################################
##  Horizon Functions

Function SC-CheckPoolAvailability {
    Param (
        [Parameter(Mandatory = $true)]
        [string]
        $CsvPath
    )

    $data = Import-Csv -Path $CsvPath

    $groupedData = $data | Group-Object -Property { $_.MachineName -replace '^([^\\-]+-[^\\-]+).*$', '$1' }

    $output = ForEach ($group In $groupedData) {
        $groupName = $group.Name
        $rowCount = 1
        $unassignedMachines = $group.Group | Where-Object { [string]::IsNullOrWhiteSpace($_.AssignedUser) }
        $rowCount += $unassignedMachines.Count
        $totalRowCount = $group.Group.Count

        [PSCustomObject]@{
            "Machine Prefix"        = $groupName
            "Unassigned"    = $rowCount
            "Total"         = $totalRowCount
        }
    }

    $output | Format-Table -AutoSize
}

Function SC-GetClusterVGPUInfo {
    Param (
        [Parameter(Mandatory = $false)]
        [string]
        $ListMachines = $false
    )

    # Get all clusters
    $clusters = Get-Cluster

    # Variables
    $output = @()

    ForEach ($cluster In $clusters) {
        $clusterName = $cluster.Name
        $clusterVideoRAM = 0
        $clusterTotalVGPUCount = @{ }
        $clusterTotalVGPUVRAM = @{ }
        $clusterTotal4GBSupport = 0
        $totalVRAMUsed = 0

        Write-Host "`nCluster: $clusterName"

        # Loop through hosts in the cluster
        ForEach ($vmhost In $cluster | Get-VMHost) {
            $nvidiaDevices = $vmhost.ExtensionData.Config.Hardware.Device | Where-Object { $_.DeviceInfo.Label -match "NVIDIA" }
            $hostName = $vmhost.Name
            $totalVRAMOnHost = 0

            ForEach ($nvidiaDevice In $nvidiaDevices) {
                $totalVRAMOnHost += [int]($nvidiaDevice.DeviceInfo.Summary.Split(" ")[-2])
            }

            # Get VMs on the host and their assigned vGPU profiles
            $VMs = $vmhost | Get-VM
            $VMProfiles = @{ }

            ForEach ($VM In $VMs) {
                $vGPUDevice = $VM.ExtensionData.Config.Hardware.Device | Where-Object { $_.Backing.Vgpu }
                $ProfileType = $vGPUDevice.Backing.Vgpu
                If ($ProfileType -ne $null) {

                    # Count vGPU profiles for this host
                    If ($clusterTotalVGPUCount.ContainsKey($ProfileType)) {
                        $clusterTotalVGPUCount[$ProfileType]++
                    }
                    Else {
                        $clusterTotalVGPUCount[$ProfileType] = 1
                    }

                    # Calculate vGPU RAM usage for this host and profile

                    $profileVRAM = $totalVRAMOnHost * $ProfileSizeInGB
                    If ($clusterTotalVGPUVRAM.ContainsKey($ProfileType)) {
                        $clusterTotalVGPUVRAM[$ProfileType] += $profileVRAM
                    }
                    Else {
                        $clusterTotalVGPUVRAM[$ProfileType] = $profileVRAM
                    }
                }
            }

            # Move this line outside the host loop to accumulate total cluster video RAM
            $clusterVideoRAM += $totalVRAMOnHost
        }

        # Output how many VMs are assigned to each GRID profile
        Write-Host "Assigned VMs to GRID Profiles in $clusterName"
        ForEach ($profileType In $clusterTotalVGPUCount.Keys) {
            # Extract the digits between "-" and the last alpha character in the profile name
            $ProfileSizeInGB = [regex]::Match($ProfileType, '-(\d+)\w').Groups[1].Value
            $ProfileTotalVRAM = $clusterTotalVGPUCount[$profileType] * $ProfileSizeinGB
            Write-Host "$profileType $($clusterTotalVGPUCount[$profileType]) VM(s), using $ProfileTotalVRAM GB VRAM"
            $TotalVRAMUsed += $ProfileTotalVRAM
        }

        # Display the total video RAM in the cluster (in GB)
        Write-Host "Total Video RAM in "$clusterName": $totalVRAMUsed GB"

        # Calculate how many additional VMs can be supported with a 4GB vGPU profile
        $total4GBSupport = [math]::Floor($clusterVideoRAMInGB / 4) - $clusterTotalVGPUCount['4GB']
        Write-Host "`nAdditional VMs Supported with 4GB vGPU Profile in "$clusterName": $total4GBSupport"

    }
}

Function SC-GetHVAssignments {
    # VARIABLE DECLARATIONS
    $username = "scarter"
    $domain = "dlzcorp"
    $password = Get-Content "$PSScriptRoot\..\Creds\$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).Substring(8))@$(hostname).txt" | ConvertTo-SecureString

    # Convert the secure password to plain text
    $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

    # Convert the username and password to Base64 encoding
    $base64Auth = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($username):$($plainPassword)"))

    # Create the authorization header
    $authorization = "Basic $base64Auth"

    # Set the REST API URL
    $apiUrl = "https://dublin-con-01.dlzcorp.com"

    # Define the REST API endpoint
    $endpoint = "/rest/login"

    # Create the headers for the API request
    $headers = @{
        "Authorization" = $authorization
        "Content-Type"  = "application/json"
    }

    # Create the request body
    $body = @{
        "username" = $username
        "password" = $plainPassword
        "domain"   = $domain
    } | ConvertTo-Json

    # Make the API request to get the session information
    $sessionResponse = Invoke-RestMethod -Uri "$apiUrl$endpoint" -Headers $headers -Method Post -Body $body

    # Extract the access key from the response
    $access_token = $sessionResponse.access_token.ToString()
    $headers["Authorization"] = "Bearer $access_token"

    # Create an empty array to store machine and user information
    $machineInfo = @()

    # Retrieve a list of VMs
    $machineEndpoint = "/rest/inventory/v1/machines"
    $machineList = Invoke-RestMethod -Uri "$apiUrl$machineEndpoint" -Method GET -Headers $headers
    Write-Output $machinelist

    # Loop through each VM and retrieve the assigned user
    ForEach ($machine In $machineList) {
        If ($machine.type -ne "UNMANAGED_MACHINE") {
            $machineId = $machine.id
            $assignmentsEndpoint = "/rest/inventory/v1/machines/$machineId/"
            $assignments = Invoke-RestMethod -Uri "$apiUrl$assignmentsEndpoint" -Method GET -Headers $headers
            If ($assignments[0].user_ids -ne $null) {
                $user = (Get-ADUser -Identity $($assignments[0].user_ids.Trim("{}")))
            }
            Else {
                $user = $null
            }
            # Create a custom object and add machine and user information
            $machineInfoObject = [PSCustomObject]@{
                MachineName     = $machine.name
                AssignedUser    = $user.samaccountname
                AssignedUserSID = $user.sid
            }

            # Add the custom object to the array
            $machineInfo += $machineInfoObject
        }
    }

    # Export the machine information to a CSV file
    $exportPath = "C:\users\scarter\output.csv"
    $machineInfo | Sort-Object MachineName | Export-Csv -Path $exportPath -NoTypeInformation
}

Function SC-GetHVDisabledAssignments {
    $CsvPathDublin = "\\colo-pan1\Files\Restricted\IT\Documentation\Network Information\VDI\VM Assignments\DUBLIN-CON-01.csv"
    $CsvPathIndy = "\\colo-pan1\Files\Restricted\IT\Documentation\Network Information\VDI\VM Assignments\INDY-CON-01.csv"
    # Column names in the CSV file
    $userSIDColumnName = "user_ids"
    $nameColumnName = "name"
    # Read and process the CSV file
    $csvData = Import-Csv -Path $csvPathDublin
    $csvData += Import-Csv -Path $CsvPathIndy
    Write-Output "`nListing Horizon VMs with disabled users assigned`n"
    ForEach ($row In $csvData) {
        $userSID = $row.$userSIDColumnName
        $name = $row.$nameColumnName
        If ($userSID -like "S-1-5-21*") {
            # Query Active Directory for the user account using the SID
            $user = Get-ADUser -Filter { objectSID -eq $userSID } -Properties * -ErrorAction SilentlyContinue
            # Check if the account is disabled
            If (-not $user.Enabled) {
                # Pad $name to a fixed width of 15 characters
                $formattedName = "{0,-15}" -f $name
                Write-Output "$formattedName assigned to $($user.UserPrincipalName)"
            }
        }
    }
    Write-Output ""
}

Function SC-GetvGPUSummary {
    $Clusters = Get-Cluster
    ForEach ($Cluster In $Clusters) {
        # Variables
        $vmhosts = Get-VMHost -Location $Cluster
        $profileCounts = @{ }
        $UsedVRAMTotal = 0
        $InstalledVRAMTotal = 0

        ForEach ($vmhost In $vmhosts) {
            $VMhost = Get-VMHost $vmhost
            $VMs = Get-VMHost -Name $vmhost | Get-VM

            # Calculate the total video RAM size of all NVIDIA cards on the host
            $totalHostVRAMKB = 0
            $nvidiaDevices = $VMhost.ExtensionData.Config.GraphicsInfo
            ForEach ($nvidiaDevice In $nvidiaDevices) {
                $totalHostVRAMKB += $nvidiaDevice.MemorySizeInKB
            }
            $totalHostVRAMGB = [math]::Round($totalHostVRAMKB / 1MB, 2)
            $InstalledVRAMTotal += $totalHostVRAMGB
            ForEach ($VM In $VMs) {
                $vGPUDevice = $VM.ExtensionData.Config.Hardware.Device | Where-Object { $_.Backing.Vgpu }
                $ProfileType = $vGPUDevice.Backing.Vgpu

                # Calculate video RAM usage based on the profile name
                If ($ProfileType -ne $null) {
                    $profileCounts[$ProfileType] += 1
                }
            }
        }

        # Display the vGPU profile summary for the cluster
        Write-Host "`nvGPU Profile Summary for Cluster: $Cluster`n"
        $ProfileSizeInGB = 0
        $profileCounts.GetEnumerator() | ForEach-Object {
            $ProfileType = $_.Key
            $Count = $_.Value
            $ProfileSizeInGB = [int]($ProfileType -replace '.*\D(\d+)\D.*', '$1')
            $ProfileVRAMTotal = $Count * $ProfileSizeInGB
            $UsedVRAMTotal += $ProfileVRAMTotal
            Write-Host ("{0,11}: {1,3} VM(s) using {2,4} GB VRAM" -f $ProfileType, $Count, $ProfileVRAMTotal)
        }
        $string = "`n   Total vGPU VRAM installed: $InstalledVRAMTotal GB"
        $AvailableVRAMTotal = $InstalledVRAMTotal - $UsedVRAMTotal
        Write-Host $String
        Write-Host "                      in use: $UsedVRAMTotal GB"
        Write-Host $('-' * $($string.Length - 1))
        Write-Host "                   available: $AvailableVRAMTotal GB"
        Write-Host ("`nCapacity for new GenProd VMs: {0,3}`n" -f $($AvailableVRAMTotal / 4))
    }
}
