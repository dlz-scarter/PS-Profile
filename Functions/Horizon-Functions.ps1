Function SC-ConnectHorizon {
    Param (
        $ConnServer
    )
    $HVServerUser = "dlzcorp.com\scarter"
    $HVServerPassword = get-content $psscriptroot\Creds\$((([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name).substring(8))@$(hostname).txt | convertto-securestring
    Import-Module -Name VMware.VimAutomation.HorizonView
    Connect-HVServer $ConnServer -Username $HVServerUser -Password $HVServerPassword
}


Function SC-GetHVDesktop {
    <#  
    .SYNOPSIS  
        This cmdlet retrieves the virtual desktops on a horizon view Server.
    .DESCRIPTION 
        This cmdlet retrieves the virtual desktops on a horizon view Server.
    .NOTES  
        Author:  Alan Renouf, @alanrenouf,virtu-al.net
    .PARAMETER State
        Hash table containing states to filter on
    .PARAMETTER ConnServer
        Name of Connection Server to query
    .EXAMPLE
        List All Desktops
        Get-HVDesktop

    .EXAMPLE
        List All Problem Desktops
        Get-HVDesktop -state @('PROVISIONING_ERROR', 
                        'ERROR', 
                        'AGENT_UNREACHABLE', 
                        'AGENT_ERR_STARTUP_IN_PROGRESS',
                        'AGENT_ERR_DISABLED', 
                        'AGENT_ERR_INVALID_IP', 
                        'AGENT_ERR_NEED_REBOOT', 
                        'AGENT_ERR_PROTOCOL_FAILURE', 
                        'AGENT_ERR_DOMAIN_FAILURE', 
                        'AGENT_CONFIG_ERROR', 
                        'UNKNOWN')
    #>
    Param (
        $State,$ConnServer
    )
    If (! ($global:DefaultHVServers | Where-Object { $_.name -eq $ConnServer })) {
        SC-ConnectHorizon $ConnServer | Out-Null
    }
    $Server = ($global:DefaultHVServers | Where-Object { $_.name -eq $ConnServer})
    $ViewAPI = $Server.ExtensionData
    $query_service = New-Object "Vmware.Hv.QueryServiceService"
    $query = New-Object "Vmware.Hv.QueryDefinition"
    $query.queryEntityType = 'MachineSummaryView'
    If ($State) {
        [VMware.Hv.QueryFilter []]$filters = @()
        ForEach ($filterstate In $State) {
            $filters += new-object VMware.Hv.QueryFilterEquals -property @{ 'memberName' = 'base.basicState'; 'value' = $filterstate }
        }
        $orFilter = new-object VMware.Hv.QueryFilterOr -property @{ 'filters' = $filters }
        $query.Filter = $orFilter
    }
    $Desktops = $query_service.QueryService_Query($ViewAPI, $query)
    $Desktops.Results.Base | Sort-Object Name
}


