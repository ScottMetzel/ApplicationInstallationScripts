<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory = $false)]
	[ValidateSet('Install', 'Uninstall', 'Repair')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory = $false)]
	[ValidateSet('Interactive', 'Silent', 'NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory = $false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory = $false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory = $false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try {
		Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
 }
 Catch {
 }

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Microsoft Corporation'
	[string]$appName = 'Monitoring Agent'
	[string]$appVersion = '10.20.18053.0'
	[string]$appArch = 'amd64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '10/24/2021'
	[string]$appScriptAuthor = 'Scott Metzel'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''

	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.8.4'
	[string]$deployAppScriptDate = '26/01/2021'
	[hashtable]$deployAppScriptParameters = $psBoundParameters

	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') {
		$InvocationInfo = $HostInvocation
 }
 Else {
		$InvocationInfo = $MyInvocation
 }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) {
			Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]."
  }
		If ($DisableLogging) {
			. $moduleAppDeployToolkitMain -DisableLogging
  }
		Else {
			. $moduleAppDeployToolkitMain
  }
	}
	Catch {
		If ($mainExitCode -eq 0) {
			[int32]$mainExitCode = 60008
  }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') {
			$script:ExitCode = $mainExitCode; Exit
  }
		Else {
			Exit $mainExitCode
  }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		Show-InstallationWelcome -CloseApps 'AgentControlPanel' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Installation tasks here>
		### Define variables related to the LA Gateway and Workspace, and make it so we have to prove we can use it by specifying a boolean value
		Show-InstallationProgress -StatusMessage "Performing Microsoft Monitoring Agent pre-installation tasks. Please wait..."
		Write-Log -Message "Specifying installation-specific variables."

		### Specify the FQDN, IP, and port number for the Log Analytics Gateway
		[System.String]$LogAnalyticsGatewayFQDN = "mygateway.myorganization.com"
		[System.String]$LogAnalyticsGatewayIPAddress = "123.456.789.012"
		[System.Int32]$LogAnalyticsGatewayPortNumber = 8080

		### Specify the Azure cloud type: 0 = Azure Commercial, 1 = Azure Government
		[System.Int32]$AzureCloudType = 1

		### Supply the Log Analytics Workspace ID and Key to use
		[System.String]$LogAnalyticsWorkspaceID = ""
		[System.String]$LogAnalyticsWorkspaceKey = ""

		### These shouldn't be edited
		[System.String]$LogAnalyticsGatewayURI = [System.String]::Concat($LogAnalyticsGatewayFQDN, ":", $LogAnalyticsGatewayPortNumber)
		[System.Boolean]$UseLogAnalyticsGateway = $false
		[System.Boolean]$AddLogAnalyticsWorkspace = $false
		[System.Boolean]$RemoveLogAnalyticsWorkspace = $false
		[System.Boolean]$InstallMMA = $false
		[System.String]$SetupPath = [System.String]::Concat($dirFiles, "\", "Setup.exe")
		###

		### Check resolution of the Gateway
		Write-Log -Message "Attempting to resolve the Fully Qualified Domain Name of the Log Analytics Gateway, which is: '$LogAnalyticsGatewayFQDN'."
		if (Resolve-DnsName -Name $LogAnalyticsGatewayFQDN -Type A -ErrorAction SilentlyContinue) {

			### If the FQDN resolved, now try a connection to the it.
			Write-Log -Message "The Gateway FQDN was resolved. Attempting a TCP Connection."
			$TestLAGatewayConnection = Test-NetConnection -ComputerName $LogAnalyticsGatewayFQDN -Port $LogAnalyticsGatewayPortNumber -ErrorAction SilentlyContinue
			if ($TestLAGatewayConnection.TcpTestSucceeded) {
				### If the connection succeeded, indicate we'll use it to connect the MMA to it, and on to the 'Workspace
				Write-Log -Message "Log Analytics Gateway connection test succeeded. Gateway will be used."
				[System.Boolean]$UseLogAnalyticsGateway = $true
			}
			else {
				Write-Log -Message "Log Analytics Gateway connection test failed. Gateway will not be used."
				[System.Boolean]$UseLogAnalyticsGateway = $false
			}
		}
		else {
			### Since the FQDN could not be resolved, try the specified IP address
			Write-Log -Message "The FQDN could not be resolved. Trying a TCP connection test its specified IP Address: '$LogAnalyticsGatewayIPAddress'."

			$TestLAGatewayConnection = Test-NetConnection -ComputerName $LogAnalyticsGatewayIPAddress -Port $LogAnalyticsGatewayPortNumber -ErrorAction SilentlyContinue
			if ($TestLAGatewayConnection.TcpTestSucceeded) {
				### If the connection succeeded, indicate we'll use it to connect the MMA to it, and on to the 'Workspace
				Write-Log -Message "Log Analytics Gateway connection test succeeded. Gateway will be used."
				[System.Boolean]$UseLogAnalyticsGateway = $true
			}
			else {
				Write-Log -Message "Log Analytics Gateway connection test failed. Gateway will not be used."
				[System.Boolean]$UseLogAnalyticsGateway = $false
			}
		}

		### Check for an existing installation
		Write-Log -Message "Checking for an existing installation of the Microsoft Monitoring Agent"
		[System.ServiceProcess.ServiceController]$GetMMAService = Get-Service -Name "HealthService" -ErrorAction SilentlyContinue

		if ($GetMMAService) {
			### If the service is present, check if it's running
			Write-Log -Message "MMA service is present."
			if ($GetMMAService.Status -eq "Running") {

				### If it's running check workspaces, and set installation of the MMA to false
				Write-Log -Message "MMA service is also running."
				[System.Boolean]$InstallMMA = $false

				Write-Log -Message "Creating MMA COM Object."
				[System.__ComObject]$NewMMACOMObject = New-Object -ComObject "AgentConfigManager.MgmtSvcCfg"

				### Get workspaces
				Write-Log -Message "Getting workspaces."
				[System.Collections.ArrayList]$CloudWorkspaces = @()
				$NewMMACOMObject.GetCloudWorkspaces() | ForEach-Object -Process {
					$CloudWorkspaces.Add($_) | Out-Null
				}


				if ($CloudWorkspaces | Where-Object -FilterScript { $_.workspaceId -eq $LogAnalyticsWorkspaceID }) {
					### If the workspace is found, check its connectivity.
					Write-Log -Message "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' is already present."

					[System.__ComObject]$CloudWorkspace = $CloudWorkspaces | Where-Object -FilterScript { $_.workspaceId -eq $LogAnalyticsWorkspaceID }

					### Check if the workspace is connected
					if ($CloudWorkspace.ConnectionStatus -eq 0) {
						Write-Log -Message "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' is already connected."
						[System.Boolean]$AddLogAnalyticsWorkspace = $false
						[System.Boolean]$RemoveLogAnalyticsWorkspace = $false
					}
					else {
						Write-Log -Message "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' is not connected. Will remove and re-add."
						[System.Boolean]$AddLogAnalyticsWorkspace = $true
						[System.Boolean]$RemoveLogAnalyticsWorkspace = $true
					}
				}
				else {
					### If the workspace is not found, set it to be added
					Write-Log -Message "Log Analytics Workspace ID: '$LogAnalyticsWorkspaceID' not present."
					[System.Boolean]$AddLogAnalyticsWorkspace = $true
					[System.Boolean]$RemoveLogAnalyticsWorkspace = $false
				}
			}
			else {
				Write-Log -Message "MMA service is not running. Setting MMA to be installed."
				[System.Boolean]$AddLogAnalyticsWorkspace = $false
				[System.Boolean]$RemoveLogAnalyticsWorkspace = $false
				[System.Boolean]$InstallMMA = $true
			}
		}
		else {
			Write-Log -Message "MMA service not found. Setting MMA to be installed."
			[System.Boolean]$InstallMMA = $true
		}

		### Now create the string for arguments for setup.exe for the MMA based on the connection tests.
		if ($UseLogAnalyticsGateway -and ($true -eq $InstallMMA)) {
			Write-Log -Message "Creating parameter string for setup.exe specifying Log Analytics Gateway."
			[System.String]$LogAnalyticsGatewayParameters = "/qn NOAPM=1 ADD_OPINSIGHTS_WORKSPACE=1 OPINSIGHTS_WORKSPACE_AZURE_CLOUD_TYPE=$AzureCloudType OPINSIGHTS_WORKSPACE_ID=""$LogAnalyticsWorkspaceID"" OPINSIGHTS_WORKSPACE_KEY=""$LogAnalyticsWorkspaceKey"" OPINSIGHTS_PROXY_URL=""$LogAnalyticsGatewayURI"" AcceptEndUserLicenseAgreement=1"
		}
		elseif ($true -eq $InstallMMA) {
			Write-Log -Message "Creating parameter string for setup.exe not specifying Log Analytics Gateway."
			[System.String]$LogAnalyticsGatewayParameters = "/qn NOAPM=1 ADD_OPINSIGHTS_WORKSPACE=1 OPINSIGHTS_WORKSPACE_AZURE_CLOUD_TYPE=$AzureCloudType OPINSIGHTS_WORKSPACE_ID=""$LogAnalyticsWorkspaceID"" OPINSIGHTS_WORKSPACE_KEY=""$LogAnalyticsWorkspaceKey"" AcceptEndUserLicenseAgreement=1"
		}
		else {
			Write-Log -Message "'InstallMMA' is set to: '$InstallMMA' so not creating parameter string."
		}

		### If the Log Analytics workspace needs to be removed, remove it.
		if ($true -eq $RemoveLogAnalyticsWorkspace) {
			Write-Log -Message "Removing Log Analytics Workspace from Microsoft Monitoring Agent already present."
			$NewMMACOMObject.RemoveCloudWorkspace($LogAnalyticsWorkspaceID)

			Write-Log -Message "Reloading MMA configuration."
			$NewMMACOMObject.ReloadConfiguration()
		}

		### If the Log Analytics workspace needs to be added, add it.
		if ($true -eq $AddLogAnalyticsWorkspace) {
			Write-Log -Message "Adding Log Analytics Workspace to Microsoft Monitoring Agent already present."
			$NewMMACOMObject.AddCloudWorkspace($LogAnalyticsWorkspaceID, $LogAnalyticsWorkspaceKey, $AzureCloudType)

			Write-Log -Message "Reloading MMA configuration."
			$NewMMACOMObject.ReloadConfiguration()
		}

		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		## Handle Zero-Config MSI Installations
		<#
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) {
				$ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
   }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) {
				$defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ }
   }
		}
		#>
		## <Perform Installation tasks here>

		## Install Microsoft Monitoring Agent by calling setup.exe

		if ($true -eq $InstallMMA) {
			Show-InstallationProgress -StatusMessage "Installing the Microsoft Monitoring Agent. Please wait..."
			Execute-Process -Path $SetupPath -Parameters $LogAnalyticsGatewayParameters -WindowStyle Hidden -IgnoreExitCodes "3010"
		}
		else {
			Show-InstallationProgress -StatusMessage "Not installing the Microsoft Monitoring Agent."
		}

		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>

		## Display a message at the end of the install
		Show-InstallationPrompt -Message "Microsoft Monitoring Agent installation finished." -ButtonRightText 'OK' -Icon Information -Timeout 5 -ExitOnTimeout $true
	}
	ElseIf ($deploymentType -ieq 'Uninstall') {
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Uninstallation tasks here>


		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		## Handle Zero-Config MSI Uninstallations
		<#If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) {
				$ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
   }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		#>
		# <Perform Uninstallation tasks here>
		Show-InstallationProgress -StatusMessage "Uninstalling the Microsoft Monitoring Agent. Please wait..."
		Execute-MSI -Action 'Uninstall' -Path 'MOMAgent.msi' -Parameters '/qn'

		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'

		## <Perform Post-Uninstallation tasks here>
		Show-InstallationPrompt -Message "Microsoft Monitoring Agent uninstallation finished." -ButtonRightText 'OK' -Icon Information -Timeout 5 -ExitOnTimeout $true

	}
	ElseIf ($deploymentType -ieq 'Repair') {
		##*===============================================
		##* PRE-REPAIR
		##*===============================================
		[string]$installPhase = 'Pre-Repair'

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Repair tasks here>

		##*===============================================
		##* REPAIR
		##*===============================================
		[string]$installPhase = 'Repair'

		## Handle Zero-Config MSI Repairs
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) {
				$ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
   }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		# <Perform Repair tasks here>

		##*===============================================
		##* POST-REPAIR
		##*===============================================
		[string]$installPhase = 'Post-Repair'

		## <Perform Post-Repair tasks here>


	}
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
