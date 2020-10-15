
#################################################################################################################
# Dung K Hoang 
#														#
#    v1.0		Initial									6/1/2020	#
#    v2.0 		Add health check                        9/1/2020
#  														#
#################################################################################################################
param (
    [string]$SourceXLS 			= $(throw "Excel Inventory Sheet filename is required."),
	[string]$OV_Appliance_IP 	= $(throw "OV Appliance IP is required."),
	[string]$serverFilter 		= ''	
)
$Host.UI.RawUI.WindowTitle = "Server Vitals Info"

#CONSTANTS#######################################################################################################

$CRLF 	= "`r`n"
# LOAD REFERENCES ###############################################################################################
. .\config.ps1



#################################################################################################################
function CreateLogDir ($log_dir) {
										### Create Dirs & Log File
	if (!(Test-Path -path "$log_dirname\")) {
		New-Item "$log_dirname" -type directory
		Write-Output "Log Directory $log_dirname Created Successfull - Continuing " 
	} else {
		Write-Output "Log Directory $log_dirname already exists." 
	}
}



#################################################################################################################
function createLogFile ($log_filename) {
	if(Test-Path -path "$log_filename") {
		Write-Output "LOG FILE $log_filename ALREADY EXIST"   
	} else {
		New-Item "$log_filename" -ItemType File 
		Write-Output "Creating Log file "
	}

	if(Test-Path -path "$log_filename") {
		Write-Output "LOG FILE $log_filename EXIST"   
	} else {
		Write-Output "*** FATAL ERROR *** LOG FILE could not be created "
		Read-Host ("Script Execution TERMINATED - Please Hit Enter to dismiss the console window")
		Exit 1
	} 
}



#################################################################################################################
function CreateDestinationDir ($name_dir) {
										### Create Dir
	if (!(Test-Path -path "$name_dir\")) {
		New-Item "$name_dir" -type directory
		Write-Output "Directory $name_dir Created Successfull - Continuing " 
	} else {
		Write-Output "Directory $name_dir already exists." 
	}
}



#################################################################################################################
filter log($log_filename) {
	write-host $_ -foregroundcolor "magenta" 
	(get-date).toString('dd-MMM-yyyy hh:mm:ss') + $delim + $_ | Add-Content $log_filename
}



#################################################################################################################
function LogInfo($message, $log_filename){
    $log_date = get-date -format "M/d/yyyy HH:mm:ss"
    $OutContent = "$log_date - $message"
	Write-Output $OutContent
    Add-Content $log_filename "$OutContent"
}



#################################################################################################################
# MAIN														#
#################################################################################################################
##### HKD

class server
{
	[string]$Model
	[string]$Grid
	[string]$Elevation
	[string]$Serial
	[string]$HostName
	[string]$iloIP
	[string]$ProfileTemplate

	[string]$status_fqdn			= 'OK' 

	[string]$status_system 			= 'OK'

	[string]$status_firmware		= 'OK'
	[string]$iLOfirmware
	[string]$BIOSfirmware	
	[string]$NVMeBackplaneFirmware
	
	[string]$status_Memory 			= 'OK'
	[string]$totalMemory
	[string]$memoryState 	

	[string]$status_network 		= 'OK'
	[string]$nicCount
	[string]$nicName
	[string]$nicFirmware
	[string]$nicHealth
	[string]$nicState


	[string]$status_controller		= 'OK'
	[string]$controllerSlot 	
	[string]$status_encryption	 	= 'N/A'
	[string]$encryption 
	[string]$status_logicalDrive	
	[string]$logicalDriveCount	
	[string]$physicalDriveCount	
	[string]$logicalDrive
	[string]$physicalDrive

	[string]$status_NVMe 			= 'OK'
	[string]$NVMeCount		
	[string]$NVMeDrive

	[string]$status_fan				= 'OK'
	[string]$fanState 

	[string]$status_powerSupply		= 'OK'
	[string]$powerSupplyState 				

}



class ethernet
{
	[string]$nicName
	[string]$nicHealth
	[string]$nicState
	[string]$nicFirmware	
}

class logicaldrive  
{
	[string]$ldRaid
	[string]$ldHealth
	[string]$ldState
	[string]$dataDrive
	[string]$dataHealth
	[string]$dataState
	
}

class datadrive
{
	[string]$dataLocation
	[string]$dataHealth
	[string]$dataState
}







###  Generate Output file
$stamp 				= (get-Date).ToString("MMM-dd_@_HH-mm-ss")

if (test-path $SourceXLS)
{
	$site 			= (dir $SourceXLS).BaseName.Split('-')[0]
}


$output_file 		= "Server-Vitals-$Site-$stamp.xlsx"

$ServerOK 			= @()
$ServerNotinOV		= @()

$fwRepair 			= @()
$nicRepair			= @()
$ldRepair			= @()
$NVMeRepair			= @()
$memoryRepair     	= @()
$fanRepair			= @()
$powerRepair 		= @()
$fqdnFix 			= @()
$sysRepair			= @()

$okSheet 			= "$Site-OK"
$notInOvSheet 		= "Server_not_in_OV"
$fwRepairSheet 		= "fw_Repair"	
$nicRepairSheet		= "nic_Repair"
$ldRepairSheet		= "ld_Repair"
$NVMeRepairSheet    = "NVMe_Repair"
$memoryRepairSheet 	= "memory_Repair"
$fanRepairSheet 	= "fan_Repair"
$powerRepairSheet 	= "powerSupply_Repair"
$fqdnFixSheet 		= "FQDN_fix"
$sysRepairSheet 	= "System_Repair"		




### set logging parameters
$run_stamp 			= get-date -format "MM-dd-yyyy_@_HH-mm-ss"
$log_dirname 		= ".\LOGS\Get-Server-Vitals-"+$run_stamp
CreateLogDir "$log_dirname"
$log_filename 		= "$log_dirname\Get-Server-Vitals-"+$run_stamp+".log"	
$failed_hosts 		= "$log_dirname\FailedHosts-MissingData-"+$run_stamp+".csv"
$server_info 		= "$log_dirname\Server-Vitals.csv"


createLogFile "$log_filename"



LogInfo "Get Server Vitals from OneView Script Started"  "$log_filename"   
$log_user 			= $(whoami)
LogInfo "Script Initiated by User $log_user"  "$log_filename" 

Write-Host "Server Vitals Info OneView Script `n`n`n"

								### check for csv file
if(Test-Path -path $SourceXLS) {
	LogInfo "Checking for XLS file for Servers ot be added"   "$log_filename" 
} else {
	LogInfo "*** FATAL ERROR *** SOURCE DATA CSV FILE $SourceXLS NOT FOUND - TERMINATING.. " "$log_filename" 
	Exit 1
} 


if ([string]::IsNullOrEmpty($iLO_Username)) { 
	Write-Host "****** ERRROR *****" -ForegroundColor Red
	LogInfo "****** ERRROR *****"   "$log_filename" 	
	LogInfo "MISSING iLO Username in CONFIG.PS1 File"   "$log_filename" 	
	Exit 1
}
if ([string]::IsNullOrEmpty($iLO_Password)) { 
	Write-Host "****** ERRROR *****" -ForegroundColor Red
	LogInfo "****** ERRROR *****"   "$log_filename" 	
	LogInfo "MISSING iLO Password in CONFIG.PS1 File"   "$log_filename" 	
	Exit 1
}



### Import HPOVMgmt
LogInfo "Importing HPOV POSH Modules" "$log_filename" 
Import-Module HPOneView.500
Import-Module ImportExcel


### Load data
LogInfo "Importing Source and Target Server Data" "$log_filename" 
$SourceXLS_Data     =  Import-Excel $SourceXLS -WorksheetName IPs -DataOnly | Where-Object {$_.Port -eq 'MGT'}

if ($serverFilter)
{
	$serverFilter ="*$serverFilter*"
	LogInfo "---------------  You select filter for server ---> $serverFilter" "$log_filename" 
	$SourceXLS_Data = $SourceXLS_Data	| where HostName -like $serverFilter
}
								### Connect to OV Appliance
$SecurePassword 	= ConvertTo-SecureString $OV_password -AsPlainText -Force

$cred 				= New-Object System.Management.Automation.PSCredential -ArgumentList ($OV_username,$SecurePassword)

								
LogInfo "Issuing a Disconnect HPOV Call for any exisitng connections" "$log_filename" 
$errorOrigSetting 		= $errorActionPreference
$errorActionPreference 	= "SilentlyContinue"
if ($global:connectedSessions) 
{
	disconnect-hpOVMgmt
}
$errorActionPreference 	= $errorOrigSetting

LogInfo "Connecting to HPOV Management Appliance: $OV_Appliance_IP as $OV_username" "$log_filename" 

Connect-HPOVMgmt -hostname $OV_Appliance_IP -Credential $cred | out-host


$SourceXLS_Data			= $SourceXLS_Data | where HostName -notlike "*mp*"		# exlcude Simplivity nodes
foreach ($entry in $SourceXLS_Data) 
{	
	$iLO_IPv4_Address 	= $entry.IPv4

	$dnsName 			= $entry.Hostname
	$iLoName 			= $entry.Hostname
	
	$assignedRole 		= $entry.Role
	$assignedEnv 		= $entry.Environment
	$serverProfileName 	= $entry.Profile

	$serverModel 		= $entry.model
	$grid 				= $entry.Grid
	$ru 				= $entry.Elevation
	



	
	if ([string]::IsNullOrEmpty($iLO_IPv4_Address)) 
	{ 
		Write-Host "****** ERRROR *****" -ForegroundColor Red
		LogInfo "****** ERRROR *****"   "$log_filename" 	
		LogInfo "MISSING iLO IP ADDRESS EMPTY for Hostname: $iloName "    "$log_filename" 	
		Add-Content $failed_hosts "$hostRecord"
		continue
	}



	
	if ([string]::IsNullOrEmpty($serverModel)) 
	{ 
		Write-Host "****** ERRROR *****" -ForegroundColor Red
		LogInfo "****** ERRROR *****"   "$log_filename" 	
		LogInfo "MISSING Server Model for Hostname: $iloName"    "$log_filename" 	
		Add-Content $failed_hosts "$hostRecord"
		continue
	}


    $hpOVServer 	= Get-HPOVServer | where-object { $_.mpHostInfo.mpIpAddresses[0].address -eq $iLO_IPv4_Address }
	$isILOConnect = $True
	#$isILOConnect	= (Test-NetConnection -ComputerName $iLO_IPv4_Address -Port 443 -informationLevel Quiet).TcpTestSucceeded
    
	if ($isILOConnect)
	{
		if ($NULL -ne $hpOVServer)
		{        

			$iloIP          = $hpOVServer.mpHostInfo.mpIpAddresses[0].address


			LogInfo  "Retrieving Info from Server: $iLoname  iLOIP: $iloIP"   "$log_filename" 


			$iloIP          = $hpOVServer.mpHostInfo.mpIpAddresses[0].address
			$HostName 		= $hpOVServer.name
			$sn  			= $hpOVServer.serialNumber

			$sp 			= Send-HPOVRequest -uri $hpOVServer.serverProfileUri
			if ($sp.serverProfileTemplateUri)
			{
				$spt 			= (Send-HPOVRequest -uri $sp.serverProfileTemplateUri).name
			}

			$iLOfirmware 	= $hpOVServer.mpFirmwareVersion
			$BIOSfirmware 	= $hpOVServer.romVersion

			$serverState 	= $True				# Initial value State == OK
			
			$s 					= new-object server
			$s.Model 			= $ServerModel
			$s.Grid 			= $grid
			$s.Elevation 		= $ru
			$s.Serial 			= $sn
			$s.iloIP			= $iloIP 
			$s.HostName 		= $HostName
			$s.ProfileTemplate  = $spt


			# #####################################
			# 
			# 		hostName  check
			#
			# #####################################
			if ($HostName -notlike "*$dnssuffix")
			{
				$fqdnFix 		+= $s
				$s.status_fqdn  = 'Fix FQDN'
			}
			

			# #####################################
			# 
			# 		Firmware Check 
			#
			# #####################################



			$s.iLOfirmware		= $iLOfirmware 
			$s.BIOSfirmware 	= $BIOSfirmware

			$ISiloFWcompliant 	= ($iLOfirmware -like "*$iLOBaseline*") 	-or ($iLOfirmware -like "*$iLOBaselineDate*")
			$ISbiosFWcompliant 	= ($BIOSfirmware -like "*$BIOSBaseline*") 	-or ($BIOSfirmware -like "*$BIOSBaselineDate*")

			if (  $ISiloFWcompliant  -and $ISbiosFWcompliant )
			{
				$serverState 	= $True -and $serverState 	
			}
			else
			{
				$s.status_firmware 	= "Check iLO firmware or BIOS firmware"
				$fwRepair			+= $s	
				$serverState 		= $False	
			}



			# #####################################
			# 
			# 		Network Check 
			#
			# #####################################


			#---- GET list of subresources
			$resources 					= (send-HPOVRequest -uri $hpOVServer.subResources.devices.uri).data

			$All_Nic_Array 				= @()

			$NIClist 					= $resources | where DeviceType -like '*NIC*'

			foreach ($NIC in $NIClist)
			{
				## Log NIC problem here
				$n 						= new-object ethernet

				$n.nicName 				= $NIC.name
				$n.nicHealth			= $health = $NIC.status.health
				$n.nicState				= $state  = $NIC.status.State
				$n.nicFirmware 			= $NIC.FirmwareVersion.Current.versionString

				$All_Nic_Array			+= $n

			}

			$all_health		= $all_state 	= $True
			$nicName 		= ""
			$nicFirmware 	= ""

			foreach ($NIC in $All_Nic_Array)
			{
				$NIC.nicHealth 	= if ($NIC.nicHealth)	{ $NIC.nicHealth}	else {'Unkmown'} 
				$NIC.nicState 	= if ($NIC.nicState) 	{ $NIC.nicState}	else {'Unkmown'} 

				$nicFirmware 	+= $NIC.nicFirmware + '|'
				$nicName 		+= $NIC.nicName + '|'

				$all_health 	= $all_health -and ($NIC.nichealth -eq 'OK') 
				$all_state  	= $all_health -and ($NIC.nicState  -eq 'Enabled') 
			}
			$all_health 		= $all_health -and $all_state
			$s.nicName 			= $nicName
			$s.nicFirmware		= $nicFirmware
			$s.nicCount			= $All_Nic_array.count

			
			if ($all_health)
			{
				$s.nicHealth 	= 'OK'
				$serverState 	= $True -and $serverState 	
			}
			else
			{
				$h = $st  = ""
				$All_Nic_Array | % { $h += $_.nicHealth + '|'; $st += $_.nicState + '|'}
				$s.nicHealth		= $h
				$s.NicState 		= $st
				$s.status_network 	= "Check health and state of NICs"

				$nicRepair		+= $s
				
				$serverState 	= $False
			}


			# #####################################
			# 
			# 		Logical Disk Check 
			#
			# #####################################


			#---- GET list of subresources
			$localStorageCount 			= $hpOVServer.subResources.localStorage.Count
			$resources 					= (send-HPOVRequest -uri $hpOVServer.subResources.localStorage.uri).data

			if ( ($serverModel -like '*Apollo*') -and ($localStorageCount -ge 1) ) # Check Apollo servers and data disks only
			{
				$ldArray 				= @()
				$pdArray 				= @()
				$DataController			= $resources[1] 

				#------- Check encryption and drives
				$slot 					= $DataController.Location
				$encryption 			= $DataController.EncryptionEnabled
				$logicalDriveCount		= $DataController.logicalDrives.Count
				$physicalDriveCount		= $DataController.physicalDrives.Count

				$s.controllerSlot 		= $slot
				$s.logicalDriveCount	= $logicalDriveCount
				$s.physicalDriveCount	= $physicalDriveCount
				$s.encryption 			= $encryption

				$LD_eq_PD  				= $logicalDriveCount -eq $physicalDriveCount

				$driveHealth 			= ($DataController.status.Health -eq 'OK') -and ($DataController.status.State -eq 'Enabled')

				if ($driveHealth -and $encryption -and $LD_eq_PD)
				{
					$s.status_encryption		= 'OK'
					$s.status_logicalDrive 		= 'OK'
					$serverState 		+= $True -and $serverState
				}

				else
				{
					$serverState 		= $False
					$s.status_encryption		= if ($encryption)  {'OK'} else {'Check Encryption'}
					$s.status_logicalDrive 		= if ($LD_eq_PD) 	{'OK'} else {'Check logical disk'}
					if (-not $DriveHealth)
					{
						foreach ($ld in $DataController.LogicalDrives)
						{
							$lDrive 	= "Drive:{0}-Health:{1}-State:{2}" 	-f $ld.LogicalDriveNumber, $ld.status.Health, $ld.status.State
							$ldArray		+= $lDrive
						}

						foreach ($pd  in $DataController.physicalDrives)   #physical disks
						{
								$pDrive 	= "Location:{0}-Health:{1}-State:{2}" 	-f $pd.location, $pd.status.Health, $pd.status.State
								$pdArray 	+= $pDrive
						}

						$s.logicalDrive 	= if ($ldArray) { $ldArray -join "|$CRLF" } else {''}
						$s.physicalDrive 	= if ($pdArray) { $pdArray -join "|$CRLF"} else {''}

						$s.status_logicalDrive 	= 'Check status logical disk/physical disk'
						LogInfo "Check Logical disk/Physical disk " "$log_filename" 
					}

					$ldRepair 				+= $s
				}
			}

			#----------------- Need access to iLO now
			$iLOConnection = Connect-HPEiLO -Address $iLOIP -Username $iLO_Username -Password $iLO_Password -DisableCertificateAuthentication

			# #####################################
			# 
			# 		NVMe disk Check and Firmware Report
			#
			# #####################################


			if ($serverModel -like '*ProLiant*')    # Check on proLiant only
			{
				$pciDevice      =  (Get-HPEiLOPCIDeviceInventory -Connection $iLOConnection ).PCIDevice      | where devicelocation -like '*NVME*'

				$_miss 			= ''
				$MissingNVMe	= $pciDevice | where name -like '*Empty*'
				foreach ($m in $MissingNVMe)
				{
					$_miss 		+= $m.locationString + '|'
				}


				if ($Null -eq $MissingNVMe)
				{
					$s.NVMeCount 	= $pciDevice.count
					
					$serverState 	= $True -and $serverState
				}
				else
				{
					$s.NVMeCount	= $MissingNVMe.Count
					$s.NVMeDrive	= $_miss -replace ".$"
					$s.status_NVMe	= 'Check NVMe missing drive'

					$serverState 	= $False
					$NVMeRepair 	+= $s
					write-host -foregroundcolor Yellow "Missing $_miss "
					
				}

				# --------------   NVMe Firmware Report
				$nv 						= $HPOVServer |  Show-HPOVFirmwareReport | where component -like 'NVMe*Plane*'
				$s.NVMeBackplaneFirmware 	= "{0}" -f $nv.Installed

			}

			if ($HPOVserver.PowerState -eq 'On')			# Server is ON - Skipp memory/fan
			{
				# #####################################
				# 
				# 		Memory Check 
				#
				# #####################################	
				
				#---- GET list of subresources
				$badmArray 					= @()
				$memorySize 				= 0
				$resources 					= (send-HPOVRequest -uri $hpOVServer.subResources.MemoryList.uri).data


				foreach ($m in $resources)
				{
					$memorySize 			+= $m.Boardtotalmemorysize / 1KB
				}

				$s.totalMemory 				= "{0} GB" -f $memorySize


				$memoryListperCPU 			=  (Get-HPEiLOMemoryInfo -Connection $iloConnection).memoryDetails
				foreach ($m in $memoryListperCPU)
				{
					foreach ($mData in $m.memoryData)
					{
						$size 					= $mData.CapacityMiB/1KB
						if ($size -ne 0)
						{
							$memStatus			= "Location:{0} - Size:{1:00} GB - Health:{2} - State:{3}" -f $mData.DeviceLocator, $size , $mData.Status.health, $mData.status.state
							if ( $memStatus -notlike "*Health:OK - State:Enabled*")
							{
								$serverState 	= $False 
								$badmArray		+= $memStatus  
							}
						}
					}
				}

				if ($badmArray)
				{
					$memoryStatus 		= $badmArray -join "|$CRLF"
					$s.memoryState		= $memoryStatus
					$s.status_Memory	= 'Check memory state'
					$memoryRepair 		+= $s

					write-host -foregroundcolor Yellow $memoryStatus
				}


				# #####################################
				# 
				# 		Fan check
				#
				# #####################################

				$badFan						= @() 
				$fanList 					= (Get-HPEiLOFan -Connection $iloConnection).fans
				foreach ($f in $fanList)
				{
					$fanStatus 				= "Location:{0} - status:{1}" -f $f.name, $f.Status.health
					if ($fanStatus -notlike '*status:OK*')
					{
						$serverState		= $False
						$badFan 			+= $fanStatus
					}
				}

				if ($badFan)
				{
					$fanStatus	 			= $badFan -join "|$CRLF"
					$s.fanState 			= $fanStatus
					$s.status_fan 			= 'Check fan state'
					$fanRepair 				+= $s

					write-host -foregroundcolor Yellow $fanStatus
				}
			}
			else
			{
				write-host -foreground Yellow "Server $hostName is OFF. Skip memory and fan check "
			}


			# #####################################
			# 
			# 		Power Supply check
			#
			# #####################################

			$badPower					= @()
			$psList 					= (Get-HPEiLOPowerSupply -Connection $iloConnection).powerSupplies | where serialNumber -ne $NULL

			foreach ($p in $psList)
			{
				$powerStatus 			= "Bay:{0} - status:{1}" -f $p.bayNumber, $p.status.health 
				if ($powerStatus -notlike '*status:OK*')
				{
					$serverState		= $False
					$badPower 			+= $powerStatus
				}
			}

			if ($badPower)
			{
				$powerStatus 			= $badPower -join "|$CRLF"
				$s.powerSupplyState		= $powerStatus
				$s.status_powerSupply	= 'Check Power Supply'

				$powerRepair			+= $s

			}



			# #####################################
			# 
			# 		System Health Check 
			#
			# #####################################
			$sys		= Get-HPEiLOSystemInfo -Connection $iloConnection
			$Healthy 	= ($sys.SystemStatus.HealthRollUp -eq 'OK') -and ($sys.SystemStatus.Health -eq 'OK')

			if (-not $Healthy)
			{
				$serverState 		= $False
				$s.status_system	= 'Check System Health'
				$sysRepair			+= $s
			}



			# --------------------- Final
			if ($serverState)
			{
				$serverOK 		+= $s
			}

		}
		else
		{
			write-host -ForegroundColor Yellow "This server $iLOName is not in OneView......" 
			$ServerNotinOV		+= $s
		}
	}
	else 
	{
			write-host -ForegroundColor Yellow "This server: $iLoname  has no connection to iLO: $iloIP......" 
			$ServerNotinOV		+= $s
	}
}




# Generate Excel 
$fixCondition 		= new-conditionaltext -text 'Fix' 		-ConditionalType ContainsText -conditionalTextColor Black -backGroundColor Yellow
$checkCondition 	= new-conditionaltext -text 'Check' 	-ConditionalType ContainsText -conditionalTextColor Black -backGroundColor Yellow
$criticalCondition 	= new-conditionaltext -text 'Critical' 	-ConditionalType ContainsText -conditionalTextColor Black -backGroundColor Yellow 
$warningCondition 	= new-conditionaltext -text 'Warning' 	-ConditionalType ContainsText -conditionalTextColor Black -backGroundColor Yellow
$disableCondition   = new-conditionaltext -text 'Disabled' 	-ConditionalType ContainsText -conditionalTextColor Black -backGroundColor Yellow 
$NVMeFwCondition 	= new-conditionaltext -text '1.24' 		-ConditionalType ContainsText -conditionalTextColor Black -backGroundColor Cyan
$allConditions 		= @($fixCondition, $checkCondition,$criticalCondition,$warningCondition,$disableCondition,$NVMeFwCondition)

if ($serverOK)
{ $ServerOK 		| export-Excel -AutoSize -path $output_file -WorksheetName $okSheet }

if ($ServerNotinOV)
{ 	$ServerNotinOV	| export-Excel -AutoSize -path $output_file -WorksheetName $notInOvSheet -conditionaltext $allConditions }

if ($sysRepair)
{ 	$sysRepair		| export-Excel -AutoSize -path $output_file -WorksheetName $sysRepairSheet -conditionaltext $allConditions }

if ($fqdnFix)
{ 	$fqdnFix		| export-Excel -AutoSize -path $output_file -WorksheetName $fqdnFixSheet -conditionaltext $allConditions }

if ($fwRepair)
{ 	$fwRepair 		| export-Excel -AutoSize -path $output_file -WorksheetName $fwRepairSheet -conditionaltext $allConditions }

if ($nicRepair)
{ 	$nicRepair 		| export-Excel -AutoSize -path $output_file -WorksheetName $nicRepairSheet -conditionaltext $allConditions  }

if ($ldRepair)
{ 	$ldRepair 		| export-Excel  -path $output_file -WorksheetName $ldRepairSheet -conditionaltext $allConditions  }

if ($NVMeRepair)
{ 	$NVMeRepair 	| export-Excel -AutoSize -path $output_file -WorksheetName $NVMeRepairSheet -conditionaltext $allConditions  }

if ($memoryRepair)
{ 	$memoryRepair 	| export-Excel -AutoSize -path $output_file -WorksheetName $memoryRepairSheet -conditionaltext $allConditions  }
 
if ($fanRepair)
{ 	$fanRepair 		| export-Excel -AutoSize -path $output_file -WorksheetName $fanRepairSheet -conditionaltext $allConditions  }

if ($powerRepair)
{ 	$powerRepair 	| export-Excel -AutoSize -path $output_file -WorksheetName $powerRepairSheet -conditionaltext $allConditions  } 



LogInfo "Script Execution Completed" "$log_filename" 

### disconnect from hpov
disconnect-hpOVMgmt  | out-null

$a = @"
-------------------------------------------------------------

You can review the result here --> $output_file
while AHS log is being generated for you

-------------------------------------------------------------
"@

write-host -foregroundcolor Cyan $a

#################################################################################################################
function generateAHSLog ($repair, [string]$repairText) 
{
	if ($repair)
	{
		$repairText 	= $repairText -replace '_', ' '
		write-host -foregroundcolor Cyan "### Generating AHS Log for $repairText"
		foreach ($s in $repair)
		{
			$iLOConnection = Connect-HPEiLO -Address $s.iLOIP -Username $iLO_Username -Password $iLO_Password -DisableCertificateAuthentication
			Save-HPEiLOAHSLog -filelocation C:\AHSLog -Duration All -CompanyName HPE -connection $iLOConnection
			
			Disconnect-HPEiLO -connection $iLOConnection
		}
	}
}

#################################################################################################################

if ($nicRepair) 	{ generateAHSLog -repair $nicRepair -repairText 'nic_Repair'}
if ($ldRepair) 		{ generateAHSLog -repair $ldRepair -repairText 'ld_Repair'}
# if ($NVMeRepair) 	{ generateAHSLog -repair $NVMeRepair -repairText 'NVMe_Repair'}
if ($memoryRepair) 	{ generateAHSLog -repair $memoryRepair -repairText 'Memory_Repair'}
if ($fanRepair) 	{ generateAHSLog -repair $fanRepair -repairText 'fan_Repair'}
if ($powerRepair)	{ generateAHSLog -repair $powerRepair -repairText 'power_Repair'}

#################################################################################################################
# END														#
#################################################################################################################

