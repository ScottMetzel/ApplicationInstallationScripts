### Specify the FQDN, IP, and port number for the Log Analytics Gateway
[System.String]$LogAnalyticsGatewayFQDN = "mygateway.myorganization.com"
[System.String]$LogAnalyticsGatewayIPAddress = "123.456.789.012"
[System.Int32]$LogAnalyticsGatewayPortNumber = 8080

### Supply the Log Analytics Workspace ID to use
[System.String]$LogAnalyticsWorkspaceID = ""

### Supply the product code for the Microsoft Monitoring Agent
[System.String]$MMAProductCode = "{88EE688B-31C6-4B90-90DF-FBB345223F94}"

### These shouldn't be edited
$InformationPreference = "Continue"
[System.Boolean]$UseLogAnalyticsGateway = $false
[System.Boolean]$UseLogAnalyticsGatewayFQDN = $false
[System.Boolean]$MMAInstalled = $false
[System.Boolean]$LogAnalyticsWorkspaceConfigured = $false
[System.Boolean]$LogAnalyticsGatewayConfigured = $false
[System.Boolean]$InstallationConfigurationSuccessful = $false
[System.Int32]$ExitCode = 0
###

### Check for an existing installation of the MMA using the product code
Write-Information -MessageData "START: Microsoft Monitoring Agent installation check."
Write-Information -MessageData "Checking for an existing installation of the Microsoft Monitoring Agent"
if (Get-InstalledApplication -ProductCode $MMAProductCode -ErrorAction SilentlyContinue) {
    Write-Information -MessageData "Microsoft Monitoring Agent is already installed."
    [System.Boolean]$MMAInstalled = $true
}
else {
    Write-Information -MessageData "Microsoft Monitoring Agent is not installed."
    [System.Boolean]$MMAInstalled = $false
}
Write-Information -MessageData "END: Microsoft Monitoring Agent installation check."
###

### Check resolution of the Gateway
Write-Information -MessageData "START: Log Analytics Gateway connectivity test."
Write-Information -MessageData "Attempting to resolve the Fully Qualified Domain Name of the Log Analytics Gateway, which is: '$LogAnalyticsGatewayFQDN'."
if (Resolve-DnsName -Name $LogAnalyticsGatewayFQDN -Type A -ErrorAction SilentlyContinue) {

    ### If the FQDN resolved, now try a connection to the it.
    Write-Information -MessageData "The Gateway FQDN was resolved. Attempting a TCP Connection."
    $TestLAGatewayConnection = Test-NetConnection -ComputerName $LogAnalyticsGatewayFQDN -Port $LogAnalyticsGatewayPortNumber -ErrorAction SilentlyContinue

    if ($TestLAGatewayConnection.TcpTestSucceeded) {
        ### If the connection succeeded, indicate we'll use it to connect the MMA to the workspace and use the FQDN of the Gateway
        Write-Information -MessageData "Log Analytics Gateway connection test succeeded. Gateway FQDN will be used."
        [System.Boolean]$UseLogAnalyticsGateway = $true
        [System.Boolean]$UseLogAnalyticsGatewayFQDN = $true
    }
    else {
        ### If the connection test did not succeed, set the MMA agent to go direct
        Write-Information -MessageData "Log Analytics Gateway connection test failed. Gateway will not be used."
        [System.Boolean]$UseLogAnalyticsGateway = $false
        [System.Boolean]$UseLogAnalyticsGatewayFQDN = $false
    }
}
else {
    ### Since the FQDN could not be resolved, try the specified IP address
    Write-Information -MessageData "The FQDN could not be resolved. Trying a TCP connection test its specified IP Address: '$LogAnalyticsGatewayIPAddress'."
    $TestLAGatewayConnection = Test-NetConnection -ComputerName $LogAnalyticsGatewayIPAddress -Port $LogAnalyticsGatewayPortNumber -ErrorAction SilentlyContinue

    if ($TestLAGatewayConnection.TcpTestSucceeded) {
        ### If the connection succeeded, indicate we'll use it to connect the MMA to the workspace and use the IP Address of the Gateway
        Write-Information -MessageData "Log Analytics Gateway connection test succeeded. Gateway IP Address will be used."
        [System.Boolean]$UseLogAnalyticsGateway = $true
        [System.Boolean]$UseLogAnalyticsGatewayFQDN = $false
    }
    else {
        ### If the connection test did not succeed, set the MMA agent to go direct
        Write-Information -MessageData "Log Analytics Gateway connection test failed. Gateway will not be used."
        [System.Boolean]$UseLogAnalyticsGateway = $false
        [System.Boolean]$UseLogAnalyticsGatewayFQDN = $false
    }
}
Write-Information -MessageData "END: Log Analytics Gateway connectivity test."
###

### Create a new COM Object for the MMA.
Write-Information -MessageData "Creating MMA COM Object."
[System.__ComObject]$NewMMACOMObject = New-Object -ComObject "AgentConfigManager.MgmtSvcCfg"

<#
    Get workspaces and add them to a new array list (even if there's only 1)
    An Array List is used here to keep results consistant; if there's only 1 workspace returned and no array list is used,
    then the object returned isn't an array object, which results in code below having to account for that.
#>
Write-Information -MessageData "Getting workspaces."
[System.Collections.ArrayList]$CloudWorkspaces = @()
$NewMMACOMObject.GetCloudWorkspaces() | ForEach-Object -Process {
    $CloudWorkspaces.Add($_) | Out-Null
}

<#
    Using the array of workspaces returned from the above code, check if the workspace ID defined up top in the variables
    is already present.
#>
if ($CloudWorkspaces | Where-Object -FilterScript { $_.workspaceId -eq $LogAnalyticsWorkspaceID }) {
    ### If the workspace is found, check its connectivity.
    Write-Information -MessageData "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' is already present."

    [System.__ComObject]$CloudWorkspace = $CloudWorkspaces | Where-Object -FilterScript { $_.workspaceId -eq $LogAnalyticsWorkspaceID }

    ### Check if the workspace is connected
    if ($CloudWorkspace.ConnectionStatus -eq 0) {
        ### If the workspace is connected, there's no need to connect or disconnect
        Write-Information -MessageData "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' is already connected."
        [System.Boolean]$LogAnalyticsWorkspaceConfigured = $true
    }
    else {
        ### If the workspace is not connected, then set the workspace to be removed and added back in an attempt to get it healthy
        Write-Information -MessageData "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' is not connected."
        [System.Boolean]$LogAnalyticsWorkspaceConfigured = $false
    }
}
else {
    ### If the workspace is not found at all, set it to be added
    Write-Information -MessageData "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' not present."
    [System.Boolean]$LogAnalyticsWorkspaceConfigured = $false
}

<#
    Since the MMA is installed and the service is running and using the gateway check from above,
    create the Gateway URL and then see if the gateway has already been added to the MMA service config
#>
if ($true -eq $UseLogAnalyticsGateway) {
    ### Since the Gateway is supposed to be used, determine if the FQDN or the IP Address is supposed to be used
    if ($true -eq $UseLogAnalyticsGatewayFQDN) {
        ### If the FQDN is supposed to be used, create the URL for the gateway using it
        Write-Information -MessageData "Creating the Log Analytics Gateway URL String using the FQDN of the Gateway."
        [System.String]$LogAnalyticsGatewayURL = [System.String]::Concat("https://", $LogAnalyticsGatewayFQDN, ":", $LogAnalyticsGatewayPortNumber)
    }
    else {
        ### If the FQDN cannot be used, create the URL for the gateway using its (manuall-supplied) IP Address instead
        Write-Information -MessageData "Creating the Log Analytics Gateway URL String using the IP Address of the Gateway."
        [System.String]$LogAnalyticsGatewayURL = [System.String]::Concat("https://", $LogAnalyticsGatewayIPAddress, ":", $LogAnalyticsGatewayPortNumber)
    }

    ### Log the value of the URL.
    Write-Information -MessageData "The URL of the Log Analytics Gateway is now: '$LogAnalyticsGatewayURL'."

    ### With the URL of the gateway created, now check for its presence in the existing MMA settings.

    if (!($NewMMACOMObject | Get-Member -Name "SetProxyInfo" -ErrorAction SilentlyContinue)) {
        ### If the SetProxyInfo method doesn't exist on the object, then the version of the MMA is too old to work with a Gateway
        Write-Information -MessageData """SetProxyInfo"" property doesn't exist on object, so not taking any proxy/gateway-related actions."
        [System.Boolean]$LogAnalyticsGatewayConfigured = $false
    }
    else {
        ### Check the value of the proxy and if it doesn't match what's expected, then set a new one.
        [System.String]$CurrentProxyURL = $NewMMACOMObject.proxyUrl
        if ($CurrentProxyURL -eq $LogAnalyticsGatewayURL) {
            ### The current value matches the desired value, don't do anything
            Write-Information -MessageData "The current proxy value of: '$CurrentProxyURL' matches the desired value of: '$LogAnalyticsGatewayURL'."
            [System.Boolean]$LogAnalyticsGatewayConfigured = $true
        }
        else {
            ### The current value does not match the desired value, set the proxy. The configuration doesn't have to be reloaded for the changes to take effect
            Write-Information -MessageData "The current proxy value of: '$CurrentProxyURL' does not match the desired value of: '$LogAnalyticsGatewayURL'. Setting a proxy."
            [System.Boolean]$LogAnalyticsGatewayConfigured = $false
        }
    }
}
else {
    Write-Information -MessageData "Installation was set to go direct, so not checking the gateay configuration."
    [System.Boolean]$LogAnalyticsGatewayConfigured = $true
}

<#
    Now that presence of the MMA agent has been checked, the workspace ID has been checked,
    and the gateway configuration have been checked, determine if there's a valid installation.

    Only one of the checks below has to flip the value to false for the installation to come back
    as failed.
#>
### Flip the final boolean value based on whether the MMA is installed or not.
if ($true -eq $MMAInstalled) {
    Write-Information -MessageData "The MMA agent is installed. Proceeding to next check."
    [System.Boolean]$InstallationConfigurationSuccessful = $true
}
else {
    Write-Information -MessageData "The MMA agent is not installed. Proceeding to next check."
    [System.Boolean]$InstallationConfigurationSuccessful = $false
}

### Flip the final boolean value based on whether the workspace ID is present or not.
if ($true -eq $LogAnalyticsWorkspaceConfigured) {
    Write-Information -MessageData "The desired Log Analytics Workspace is configured and connected. Proceeding to final check."
    [System.Boolean]$InstallationConfigurationSuccessful = $true
}
else {
    Write-Information -MessageData "The desired Log Analytics Workspace is not configured or may not be connected. Proceeding to next check."
    [System.Boolean]$InstallationConfigurationSuccessful = $false
}

### Flip the final boolean value based on whether the gateway configuration has been set or not.
if ($true -eq $LogAnalyticsGatewayConfigured) {
    Write-Information -MessageData "The desired Log Analytics Gateway state has been set. Proceeding to write out overall installation state."
    [System.Boolean]$InstallationConfigurationSuccessful = $true
}
else {
    Write-Information -MessageData "The desired Log Analytics Gateway state has not been set. Proceeding to write out overall installation state."
    [System.Boolean]$InstallationConfigurationSuccessful = $true
}

<#
    Now with everything checked, write out the installation state.
    This is based on: https://docs.microsoft.com/en-us/previous-versions/system-center/system-center-2012-R2/gg682159(v=technet.10)#to-use-a-custom-script-to-determine-the-presence-of-a-deployment-type
#>
if ($true -eq $InstallationConfigurationSuccessful) {
    [System.String]$Message = "Installation and/or configuration of the Microsoft Monitoring Agent detected successfully. Writing to output stream and exiting 0."
    [System.Int32]$ExitCode = 0
    Write-Information -MessageData $Message
    Write-Output -InputObject $Message
}
else {
    [System.String]$Message = "Installation and/or configuration of the Microsoft Monitoring Agent NOT detected. Not writing to output stream, but still exiting 0."
    [System.Int32]$ExitCode = 0
    Write-Information -MessageData $Message
}
exit $ExitCode