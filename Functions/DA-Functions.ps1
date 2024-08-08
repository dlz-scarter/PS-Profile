##########################################
##  Grant modify rights to a single org group on a single project folder
Function SC-AddOrgToProject {
    Param ($Directory,
        $Org)
    $ACL = Get-Acl $Directory
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("DLZCORP\ORG_00$Org", "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
    $ACL.SetAccessRule($AccessRule)
    $ACL | Set-Acl $directory
}

##########################################
##  Move User from CSP License to EA License
Function SC-CSPtoEA() {
    Param (
        [Parameter(Mandatory = $true)]
        [String]
        $SAMAccountName
    )
    $User = Get-ADUser $SAMAccountName
    $M365 = Get-ADGroupMember -Identity "Lic-EA M365-E3" -Recursive | Select-Object -ExpandProperty Name
    $O365 = Get-ADGroupMember -Identity "Lic-CSP O365-E3" -Recursive | Select-Object -ExpandProperty Name
    $EMS = Get-ADGroupMember -Identity "Lic-CSP EMS-E3" -Recursive | Select-Object -ExpandProperty Name
    If ($M365 -contains $User) {
        Write-Output $User.Name "is already licensed via EA..`n"
    }
    Else {
        If ($O365 -contains $User.Name) {
            Write-Output "`n$($User.Name) was found in Lic-CSP O365-E3, removing..."
            Remove-ADGroupMember -Identity "Lic-CSP O365-E3" -Members $user -Confirm:$false
        }
        If ($EMS -contains $User.Name) {
            Write-Output "$($User.Name) was found in Lic-CSP EMS-E3, removing..."
            Remove-ADGroupMember -Identity "Lic-CSP EMS-E3" -Members $user -Confirm:$false
        }
        Write-Output "$($User.Name) is now being added to Lic-EA M365-E3...`n"
        Add-ADGroupMember -Identity "Lic-EA M365-E3" -Members $User
    }
}

##########################################
##  Check time drift on domain controllers
Function SC-GetDCTimeDrift () {
    Param (
        [Parameter(Mandatory = $true)]
        $DC
    )
    $Servers = $(Get-ADDomainController -Filter * | Select-Object name | Sort-Object name)
    #Invoke-Command -ComputerName $Servers -ArgumentList $Servers -Scriptblock {
    Invoke-Command -ComputerName $DC -ArgumentList $Servers -Scriptblock {
        Param ($Servers)
        ForEach ($Server In $Servers) {
            $Check = w32tm /monitor /computers:$Server /nowarn
            #$ICMP = (($Check | Select-String "ICMP") -Replace "ICMP: ", "").Trim()
            $ICMP = (($Check | Select-String "ICMP") -Replace "ICMP: ", "")[0].Trim()
            $ICMPVal = [int]($ICMP -split "ms")[0]
            $Source = w32tm /query /source
            $Name = Hostname
            If ($name.length -lt 8) {
                $name = $name + "`t"
            }
            Switch ($ICMPVal) {
                { $ICMPVal -le 0 } { $Status = "Optimal time synchronisation" }
                #you probably need another value here since you'll get no status if it is between 0 and 2m
                { $ICMPVal -lt 100000 } { $Status = "0-2 Minute time difference" }
                { $ICMPVal -ge 100000 } { $Status = "Warning, 2 minutes time difference" }
                { $ICMPVal -ge 300000 } { $Status = "Critical. Over 5 minutes time difference!" }
            }
            $String = $Name.ToUpper() + "`t$Status" + "`t$ICMP" + "`tSource: $Source"
            Write-Output $String
        }
    }
}

##########################################
##  Force SCCM DP to retry failed distributions
Function SC-UpdateSCCMDP {
    Param (
        [Parameter(Mandatory = $false)]
        $DPFQDN
    )
    $CD = Get-Location
    $Primary = "SCCM.DLZCORP.COM"
    $SiteCode = "DLZ"
    SC-ConnectSCCM | Out-Null
    If ($DPFQDN -eq $NULL) {
        Write-Output "`nNo DP was specified, enumerating through all of them...  This may take a while."
    }
    Write-Output "`nQuerying Distribution Points..."
    $DPList = Get-CMDistributionPoint -SiteCode $SiteCode
    $session = New-PSSession -ComputerName $Primary
    ForEach ($DP In $DPList) {
        $CurrentDPFQDN = $DP.NetworkOSPath -replace '\\', ''
        If (($DPFQDN -eq $null) -or ($CurrentDPFQDN -like  "$DPFQDN*")) {
            Write-Output "`nSearching $($CurrentDPFQDN.ToUpper()) for failed distributions.  Please wait...`n"
            $Failures = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName sms_packagestatusDistPointsSummarizer -ComputerName $Primary `
            | Where-Object { $_.State -GT 1 -and $_.SourceNALPath -match $CurrentDPFQDN }
            ForEach ($Failure In $Failures) {
                $PackageID = $Failure.PackageID
                Write-Output "Failed PackageID: $PackageID"
                Invoke-Command -Session $session -ScriptBlock {
                    Get-SmbOpenFile | Where-Object -Property sharerelativepath -match "$using:PackageID" | Close-SmbOpenFile -Force -ErrorAction SilentlyContinue
                }
                $DPInstance = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName sms_distributionpoint -ComputerName $Primary `
                | Where-Object { $_.ServerNalPath -match $CurrentDPFQDN -and $_.PackageID -eq $PackageID }
                $DPInstance.RefreshNow = $true
                Set-CimInstance -InputObject $DPInstance | Out-Null
            }
        }
    }
    Remove-PSSession $session
    Set-Location $CD
    If (Test-Path -Path "($SiteCode):") {
        Remove-PSDrive $SiteCode
    }
}

##########################################
##  Find BitLocker Recovery Passord based on Key ID
Function SC-FindBitLockerPasswordByKey {
    $key = (read-host -Prompt "Enter at least part of the recovery key ID").ToUpper()
    $computers = get-adobject -Filter * | Where-Object { $_.ObjectClass -eq "msFVE-RecoveryInformation" }
    $records = $computers | Where-Object { $_.DistinguishedName -like "*$key*" }
    ForEach ($rec In $records) {
        $computer = get-adcomputer -identity ($records.DistinguishedName.Split(",")[1]).split("=")[1]
        $recoveryPass = Get-ADObject -Filter { objectclass -eq 'msFVE-RecoveryInformation' } -SearchBase $computer.DistinguishedName -Properties 'msFVE-RecoveryPassword' | Where-Object { $_.DistinguishedName -like "*$key*" }
        [pscustomobject][ordered]@{
            Computer          = $computer
            'Recovery Key ID' = $rec.Name.Split("{")[1].split("}")[0]
            'Recovery Password' = $recoveryPass.'msFVE-RecoveryPassword'
        } | Format-List
    }
}

##########################################
##  Connect to remote session if computer is online
Function SC-Remote {
    Param (
        [Parameter(Mandatory = $true)]
        [String]
        $ComputerName
    )
    If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $ComputerName -Quiet) {
        Enter-PSSession $ComputerName
    }
    Else {
        Write-Output "$ComputerName is currently unreachable.."
    }
}

##########################################
##  WIP - Silently Remove Software
#Function SC-RemoveSoftware {
#    param (
#        [Parameter(Mandatory = $true)] [String]$Name
#    )
#    $Products=Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object {$_.DisplayName -match "$name"}
#    $Products | Select DisplayName, Publisher, UninstallString
##    foreach ($product in $products) {            
#        msiexec /x $Product.pschildname /qn REBOOT=REALLYSUPPRESS
#    }
#}


##########################################
##  Account Disablement Function
Function SC-DisableAccount {
    Param (
        [Parameter(Mandatory = $true)]
        [string]
        $SAMAccountName,
        [Parameter(Mandatory = $false)]
        [bool]
        $Confirm = $true,
        [Parameter(Mandatory = $false)]
        [string]
        $ForwardEmailTo,
        [Parameter(Mandatory = $false)]
        [string]
        $GrantMailboxAccessTo
    )
    
    # The following four lines only need to be declared once in your script.
    $promptyes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Name matches, disable account."
    $promptno = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Name doesn't match, abort."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($promptyes, $promptno)
    
    #Import the Active Directory module if not already up and loaded
    $Module = Get-Module | Where-Object {
        $_.Name -eq 'ActiveDirectory'
    }
    If ($Module -eq $null) {
        Write-Output "Loading Active Directory PowerShell Module"
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    }
    Try {
        $var = Get-MGOrganization -ErrorAction Stop
    }
    Catch [System.Management.Automation.RuntimeException]
    {
        SC-ConnectMicrosoftGraph
    }
    Try {
        $var = Get-EmailAddressPolicy -ErrorAction Stop
    }
    Catch [System.Management.Automation.CommandNotFoundException]
    {
        SC-ConnectO365Services
    }
    
    Try {
        $User = Get-ADUser -Properties * $SAMAccountName
    }
    Catch {
        Write-Output "$SAMAccountName is not found in Active Directory."
        Break
    }
    
    Try {
        If ($ForwardEmailTo) {
            $Forward = Get-ADUser -Properties * $ForwardEmailTo
        }
    }
    Catch {
        Write-Output "$ForwardEmailTo is not found in Active Directory."
        Break
    }
    
    Try {
        If ($GrantMailboxAccessTo) {
            $MBAccess = Get-ADUser -Properties * $GrantMailboxAccessTo
        }
    }
    Catch {
        Write-Output "$GrantMailboxAccessTo is not found in Active Directory."
        Break
    }
    
    If ($Confirm) {
        $title = "##### Account Disablement Confirmation #####"
        If ($forward) {
            $fwdMessage = "`n  Forward Email To :   " + $forward.Name
        }
        If ($mbAccess) {
            $mbAccessMessage = "`n  Mailbox Access   :   " + $mbAccess.Name
        }
        $message = "`n  User Name        :   " + $User.name + "`n  Last Logon Date  :   " + $User.LastLogonDate + $fwdMessage + $mbAccessMessage + "`n`n###########################################`n`n"
        
        $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    }
    Else {
        $result = 0
        Write-Output $User.name
    }
    
    If ($result -ne 0) {
        Write-Output "`nCancelling...`n"
    }
    Else {
        $Timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $FileName = $User.SAMAccountName + "-" + $Timestamp + ".txt"
        $FilePath = "\\dlzcorp.com\pan-files$\Files\Restricted\IT\Users\SCarter\DisabledUsers"
        $Groups = Get-ADPrincipalGroupMembership -Identity $User.SAMAccountName | Where-Object { $_.Name -ne "Domain Users" }
        
        $AADGroups = Get-MgUserMemberOfAsGroup -UserId $User.UserPrincipalName | Where-Object { $_.OnPremisesSyncEnabled -eq $null -and $_.DisplayName -ne "VBO365 - Disabled User Accounts" }
        
        $Password = SC-GenerateStrongPassword(16)
        
        # Remove manager from User account
        Write-Output "MANAGER NAME:`n-------------" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        If ($User.Manager -ne $null) {
            Write-Output "`nClearing manager attribute"
            Write-Output "$((Get-ADUser ($User.Manager)).name)" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
            Set-ADUser $User -Manager $null
        }
        Else {
            Write-Output "BLANK" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        
        # Convert mailbox to shared
        If ((Get-Mailbox $User.SAMAccountName).RecipientTypeDetails -ne "SharedMailbox") {
            Write-Output "`nConverting user mailbox to shared mailbox"
            Set-Mailbox -Identity $User.SAMAccountName -Type Shared -WarningAction SilentlyContinue
        }
        
        # Setup email forward
        Write-Output "`n`nEMAIL FORWARD RECIPIENT:`n------------------------" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        If ($Forward -ne $null) {
            Write-Output "`nSetting up email forward to $($Forward.Name) ($($Forward.UserPrincipalName))"
            Write-Output $Forward.Name | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
            Set-Mailbox -Identity $User.SAMAccountName -ForwardingAddress $Forward.UserPrincipalName -DeliverToMailboxAndForward $true | Out-Null
        }
        Else {
            Write-Output "NOT REQUESTED" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        
        # Grant mailbox access
        Write-Output "`n`nMAILBOX ACCESS GRANTED TO:`n-------------------------" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        If ($MBAccess -ne $null) {
            Write-Output "`nGranting mailbox rights to $($MBaccess.Name) ($($MBAccess.UserPrincipalName))"
            Write-Output $MBAccess.Name | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
            Add-MailboxPermission $User.SAMAccountName -User $MBAccess.SAMAccountName -AccessRights "FullAccess" -WarningAction 'SilentlyContinue' | Out-Null
        }
        Else {
            Write-Output "NOT REQUESTED" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        
        # Disable account
        If ($User.Enabled -eq $true) {
            Write-Output "`nDisabling account"
            Disable-ADAccount $User.SAMAccountName
            
        }
        # Reset password
        Write-Output "`nChanging password to $Password"
        Set-ADAccountPassword -identity $User.SAMAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
        
        # Add NoSearch to description
        If (!($User.Description -match "NoSearch*")) {
            $Desc = "NoSearch " + $User.Description
            Write-Output "`nPrepending description field with '$Desc'"
            Set-ADUser $User -Description $Desc
        }
        
        # Remove account from AD groups
        Write-Output "`n`nLOCAL AD GROUPS:`n----------------" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        If ($Groups -ne $null) {
            Start-Sleep -Milliseconds 250
            Write-Output "`nRemoving user account from AD Groups`n"
            ForEach ($Group In $Groups) {
                Write-Output "   $($Group.Name)"
                Remove-ADPrincipalGroupMembership -Identity $User.SAMAccountName -MemberOf $Group -Confirm:$false
            }
            Start-Sleep -Milliseconds 250
            $Groups | Select-Object -ExpandProperty Name | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        Else {
            Write-Output "NONE FOUND" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        # Remove account from Unified groups
        Write-Output "`n`nUNIFIED GROUPS (O365, TEAMS, ETC):`n----------------------------------" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        If ($AADGroups -ne $null) {
            Start-Sleep -Milliseconds 250
            Write-Output "`nRemoving user account from Unified (O365, Teams, etc) Groups`n"
            If ($AADGroups -ne $null) {
                Start-Sleep -Milliseconds 250
                ForEach ($AADGroup In $AADGroups) {
                    $filter = "DisplayName eq '$($AADGroup.DisplayName)'"
                    $group = Get-MgGroup -Filter $filter
                    If ($group.GroupTypes -notmatch "DynamicMembership") {
                        Write-Output "   $($AADGroup.DisplayName)"
                        Remove-UnifiedGroupLinks -identity $AADGroup.Id -linktype Owners -Links $User.UserPrincipalName -Confirm:$false -ErrorAction SilentlyContinue
                        Remove-UnifiedGroupLinks -identity $AADGroup.Id -linktype Members -Links $User.UserPrincipalName -Confirm:$false -ErrorAction SilentlyContinue
                        $AADGroup | Select-Object -ExpandProperty DisplayName | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
                    }
                }
            }
        }
        Else {
            Write-Output "NONE FOUND" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        
        # Remove directly-assigned licenses
        Write-Output "`n`nDIRECTLY-ASSIGNED LICENSES REMOVED:`n-----------------------------------" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        $UserLicenses = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName
        If ($UserLicenses.count -gt 0) {
            $count = 0
            ForEach ($License In $UserLicenses) {
                Start-Sleep -Milliseconds 250
                Set-MgUserLicense -userid $User.UserPrincipalName -RemoveLicenses @($License.skuId) -AddLicenses @() -ErrorAction SilentlyContinue | Out-Null
                If ($Error[0].exception -ne $null) {
                    $Error[0] = $null
                }
                Else {
                    If ($count -eq 0) {
                        Write-Output "`nRemoving directly-assigned licenses from user account`n"
                        $count = 1
                    }
                    Write-Output "   $($License.SkuPartNumber)"
                    $License.SkuPartNumber | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
                }
            }
        }
        Else {
            Write-Output "NONE FOUND" | Out-File -Encoding ascii -FilePath $FilePath\$FileName -Append
        }
        
        Write-Output "`n$($User.Displayname)'s account has been disabled.`n"
        
        If ($confirm -eq $true) {
            Disconnect-MgGraph | Out-Null
        }
    }
}

Function SC-FreeDiskSpace {
    Param (
        [string]
        $ComputerName
    )
    
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        # Folders to be processed
        $tempFolders = @(
            "C:\Windows\Temp\",
            "C:\Users\*\Appdata\Local\Temp\",
            "C:\Windows\CCMCache\",
            "C:\Windows\SoftwareDistribution\"
        )
        
        # Record initial drive utilization
        $diskInfo = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
        $freeSpaceGB = $diskInfo.FreeSpace / 1GB
        $totalSpaceGB = $diskInfo.Size / 1GB
        $freeSpacePercentage = ($diskInfo.FreeSpace / $diskInfo.Size) * 100
        $diskSpaceReport = "Before: {0:N2} GB Free ({1:N2}%) " -f $freeSpaceGB, $freeSpacePercentage
        
        # Stop WUAUSERV and BITS (necessary for proper handling of C:\Windows\SoftwareDistribution)
        Write-Output "`nStopping WUAUSERV and BITS services"
        Stop-Service -Name WUAUSERV, BITS -Force
        
        # Iterate through each folder, deleting everything, silently continuing if files are still open
        ForEach ($folder In $tempFolders) {
            Write-Output "   Processing $folder"
            Stop-Service -Name WUAUSERV, BITS -Force
            Get-ChildItem -Path $folder -Recurse -Force | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        }
        
        # Restart WUAUSERV and BITS
        Write-Output "Starting WUAUSERV and BITS services`n"
        Start-Service -Name WUAUSERV, BITS
        
        # Re-check drive utilization and output results to the console
        $diskInfo = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
        $freeSpaceGB = $diskInfo.FreeSpace / 1GB
        $freeSpacePercentage = ($diskInfo.FreeSpace / $diskInfo.Size) * 100
        $diskSpaceReport += "`nAfter:  {0:N2} GB Free ({1:N2}%)`n" -f ($diskInfo.FreeSpace / 1GB), ($diskInfo.FreeSpace / $diskInfo.Size * 100)
        Write-Output $diskSpaceReport
    }
}