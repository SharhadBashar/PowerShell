###------------------------------------------------------------------------
###------------------------------------------------------------------------
### Montrium RMW Verification Script: To automate verification currently done manually (for example in IQs)
### Author: Kevin Chambers, Michael Ignoto, Simon Chen, Sharhad Bashar, Shovan Acharjee
### Script version: 1.0
### Date: 25-JUL-2018
### Reference No of related document: MTM-RMW-DSP-XX
### The script is intended to check what is listed in the accompanying XML and comparing any values to those in the XML. 
### The result is put into an HTML file, uploaded to the chosen landing site and converted to PDF.
###
###------------------------------------------------------------------------

#This is needed for function Get-FileName
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
#This is needed for function screenshot
[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
function Get-FileName($initialDirectory, $fileType)
{
# Get-FileName: Get the file name from a open file dialog starting at an initial directory and if applicable filtering on a file type.
# Parameters:
# -> initialDirectory: The initial directory to start at
# -> fileType: The file types to filter on

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	if ($fileType -eq "CSV")
	{
		$OpenFileDialog.filter = "CSV files (*.csv)| *.csv"
	}
	elseif ($fileType -eq "XML")
	{
		$OpenFileDialog.filter = "XML files (*.xml)| *.xml"
	}
	else
	{
		$OpenFileDialog.filter = "All files (*.*)| *.*"
	}
	$OpenFileDialog.ShowHelp = $true
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

#Function that takes screen shots
function screenshot($path) #Sharhad
{
	$Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
	$width = $Screen.width
	$Height = $Screen.Height
	$bounds = [Drawing.Rectangle]::FromLTRB(0, 0, $width, $height)
	$bmp = New-Object Drawing.Bitmap $bounds.width, $bounds.height
	$graphics = [Drawing.Graphics]::FromImage($bmp)
	$graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)
	$bmp.Save($path)
	$graphics.Dispose()
	$bmp.Dispose()
}

#maximizes IE Screen
function maxIE #Sharhad
{
	param($ie)
	$asm = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null	
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $ie.Width = $screen.width
    $ie.Height =$screen.height
    $ie.Top =  0
    $ie.Left = 0
}

function CheckForSiteAccessAndIfAdmin([string]$URLToCheck, [string]$UserName)
{
#Checks for access to site and if site admin
# Parameters:
# -> URLToCheck: The URL to check or ALL for general shell admin
# -> UserName: The user name to check

	$siteToCheck = get-spsite $URLToCheck
	
	if ($siteToCheck -ne $null)
	{
		$siteToCheckRootweb = $siteToCheck.Rootweb
		
		if ($siteToCheckRootweb -ne $null)
		{
			$siteToCheckURL = $siteToCheckRootweb.URL
			
			if ($siteToCheckURL -ne $null)
			{
				$userIsSiteAdmin = $siteToCheckRootweb.CurrentUser.IsSiteAdmin
				
				if ($userisSiteAdmin -eq $true)
				{
					$resultOfAccessAndAdminCheck = "PASS"
				}
				else
				{
					$resultOfAccessAndAdminCheck = "$UserName is not a site admin"
				}
			}
			else
			{
				$resultOfAccessAndAdminCheck = "Unable to access the rootweb URL"
			}
		}
		else
		{
			$resultOfAccessAndAdminCheck = "Unable to access the rootweb"
		}
	}
	else
	{
		$resultOfAccessAndAdminCheck = "$URLToCheck does not correspond to a valid site collection URL"
	}
	
	if ($siteToCheckRootweb -ne $null)
	{
		$siteToCheckRootweb.dispose()
	}
	
	if ($siteToCheck -ne $null)
	{
		$siteToCheck.dispose()
	}
	
	return $resultOfAccessAndAdminCheck
}

function GenerateGeneralInfoTable
{
	Write-Output "`n--------------------Generating General Info Table------------------"
	
	$Date = Get-Date -format "dd-MMM-yyyy"
	
	#General Information -------------------------------------------------------------------------------------------
	
	#$UpdateOrVerification By ---------------------------------------------------------------------------------------------------
	$spSite = get-spsite $AdminLandingSiteCollectionURL
	$rootweb = $spSite.rootweb
	
	$user = $rootweb.EnsureUser($FullLoginName)
	if ($user -eq $Null)
	{
		Write-Output "Could not find user named $FullLoginName, exiting..."
		stop-transcript
		Exit
	}
	else 
	{
		$VerifiedBy = $user.DisplayName
	}
	
	$rootweb.dispose()
	$spSite.dispose()
	
	##add to table
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "$UpdateOrVerification Performed By:"
	$GeneralInfoRow["Configuration Value"] = $VerifiedBy
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	#$UpdateOrVerification User Name	---------------------------------------------------------------------------------------------
	#got this already at the beginning of the main part of the script.
	
	##add to table
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "$UpdateOrVerification User Name:"
	$GeneralInfoRow["Configuration Value"] = $VerificationUserName 
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	#Server name and type ----------------------------------------------------------------------------------------------
	$ComputerName = [system.environment]::MachineName
	
	$farm = get-spfarm
	$servers = $farm.Servers

	$AppServerArray = @()
	$DatabaseServerArray = @()
	$WebFrontEndArray = @()
	$ArrayOfAPPAndWFEToCheck = @()

	foreach ($srv in $servers)
	{
		$ServerName = $srv.Name
		#get the current server only
		if ($ServerName -eq $ComputerName)
		{
			$ServerType = ""
			
			if ($srv.Role -eq "Application")
			{
				$ServerType = "Application Server"
			}
			else
			{
				#To check for WFE and Database need to check instances
				$instances = $srv.ServiceInstances
				
				foreach ($inst in $instances)
				{
					if ($inst.TypeName -eq "Microsoft SharePoint Foundation Database")
					{
						$ServerType = "Database Server"
					}
					elseif ($inst.TypeName -eq "Microsoft SharePoint Foundation Web Application")
					{
						$ServerName = $srv.Name
						$ServerType = "Web Front End Server"
					}
				}
			}
			
			#if servertype is still blank, put in whatever the role says
			if ($ServerType -eq "")
			{
				$ServerType = $srv.Role
			}
			
			##add to table
			$GeneralInfoRow = $GeneralInfoTable.NewRow()
			$GeneralInfoRow["Info"] = "SharePoint Server Name:"
			$GeneralInfoRow["Configuration Value"] = $ServerName
			$GeneralInfoRow["Is Current Server"] = "Yes"
			$GeneralInfoRow["Type"] = $ServerType
			$GeneralInfoRow["Date"] = $Date 
			$GeneralInfoTable.Rows.Add($GeneralInfoRow)
			
			#break, don't need to check other servers
			Break
		}
	}
	
	#Get SharePoint version
	$farm = get-spfarm
	$SPBuildVersion = $farm.BuildVersion
	$MajorVersionString = [string]$SPBuildVersion.Major
	$MinorVersionString = [string]$SPBuildVersion.Minor
	$BuildVersionString = [string]$SPBuildVersion.Build
	$RevisionVersionString = [string]$SPBuildVersion.Revision
		
	$InstalledProgramVersion = $MajorVersionString + "." + $MinorVersionString + "." + $BuildVersionString + "." + $RevisionVersionString
	
	if ($MajorVersionString -eq "15")
	{
		$InstalledProgrameName = "Microsoft SharePoint Server 2013"
	}
				
	##add SharePoint version and build number to General Info table (for now, think it makes more sense there)
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "SharePoint Build Version"
	$GeneralInfoRow["Configuration Value"] = $InstalledProgramVersion
	$GeneralInfoRow["Type"] = $InstalledProgrameName
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	##Add the PowerShell version to the general information table
	$PowershellMajorVersionNumber = $PSVersionTable.PSVersion.Major
	
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "PowerShell Version Number:"
	$GeneralInfoRow["Configuration Value"] = $PowershellMajorVersionNumber
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	##add Admin Landing site URL to General Info table (for now, think it makes more sense there)
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "Selected Admin Landing Site URL"
	$GeneralInfoRow["Configuration Value"] = $AdminLandingSiteCollectionURL
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	##add Site Collection chosen to General Info table (for now, think it makes more sense there)
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "Selected target Site Collection URL"
	$GeneralInfoRow["Configuration Value"] = $TargetSiteCollectionURL
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	Write-Output "`n`n Generating General Info Table Complete!"
}

function GeneratePrerequisitesTable
{
	$Date = Get-Date -format "dd-MMM-yyyy"
	
	#$SharepointApplicationServerName = Read-Host "`nPlease enter the Sharepoint Application Server Name"
	$SharepointApplicationServerName = [system.environment]::MachineName
	$SPServerIP = Read-Host "`nPlease enter the Sharepoint Application Server IP Address"
	#$LandingSite = Read-Host "`nPlease enter the Landing Site URL"
	$LandingSite = $AdminLandingSiteCollectionURL
	$DCWFileName = Read-Host "`nPlease enter the DCW filename"

	
	##Adding the Sharepoint Application Server to the Prerequisite Table
	$PrerequisitesRow = $PrerequisitesTable.NewRow()
	$PrerequisitesRow["Prerequisite"] = "Sharepoint Application Server Name"
	$PrerequisitesRow["Value"] = $SharepointApplicationServerName
	$PrerequisitesRow["Date"] = $Date
	$PrerequisitesTable.Rows.Add($PrerequisitesRow)
	
	##Adding the Sharepoint Application Server IP Address to the Prerequisite Table
	$PrerequisitesRow = $PrerequisitesTable.NewRow()
	$PrerequisitesRow["Prerequisite"] = "Sharepoint Application Server IP Address"
	$PrerequisitesRow["Value"] = $SPServerIP
	$PrerequisitesRow["Date"] = $Date
	$PrerequisitesTable.Rows.Add($PrerequisitesRow)
	
	##Adding the Landing Site URL to the Prerequisite Table
	$PrerequisitesRow = $PrerequisitesTable.NewRow()
	$PrerequisitesRow["Prerequisite"] = "Landing Site URL"
	$PrerequisitesRow["Value"] = $LandingSite
	$PrerequisitesRow["Date"] = $Date
	$PrerequisitesTable.Rows.Add($PrerequisitesRow)
	
	##Adding the DCWFileName to the Prerequisite Table
	$PrerequisitesRow = $PrerequisitesTable.NewRow()
	$PrerequisitesRow["Prerequisite"] = "DCW filename to be used"
	$PrerequisitesRow["Value"] = $DCWFileName
	$PrerequisitesRow["Date"] = $Date
	$PrerequisitesTable.Rows.Add($PrerequisitesRow)
}

function GenerateDeploymentTable
{
	$Date = Get-Date -format "dd-MMM-yyyy"
	$ActualDCWFileName = "MTM-RMW44-DCW-11.xlsm" #To replace with real code
	$RequiredDCWFileName = $PrerequisitesTable.rows[3]["Value"]
	
	##Verifying if the correct DCW was used for the workspace creation
	$DeploymentRow = $DeploymentTable.NewRow()
	$DeploymentRow["Parameters"] = "DCW that was used for the workspace creation"
	$DeploymentRow["Value"] = $ActualDCWFileName
	$DeploymentRow["Value Required"] = $RequiredDCWFileName
	
	if ($ActualDCWFileName -eq $RequiredDCWFileName)
	{
		$DeploymentRow["Pass or Fail"] = "PASS"
	}
	else
	{
		$DeploymentRow["Pass or Fail"] = "FAIL"
	}
	$DeploymentRow["Date"] = $Date
	$DeploymentTable.Rows.Add($DeploymentRow)
}

function RMWCheckSharePointFeatures([string]$WAOrRC, [string]$RMWSiteURL) #Shovan
{
	#RMWCheckSharePointFeatures: Check Active Sharepoint Features  - Shovan
	$SharePointFeaturesArray = $XMLConfig.XML.SharePointFeaturesArray.SPFeatureInfo
	
	#If are entries in the XML
	if (($SharePointFeaturesArray -ne $null) -or ($SharePointFeaturesArray.Count -gt 0))
	{
		if ($WAorRC -eq "WA")
		{
			Write-Output "`n--- Checking Work Area SharePoint Features ---"
		}
		elseif ($WAorRC -eq "RC")
		{
			Write-Output "`n--- Checking Record Center SharePoint Features ---"
		}
		
		foreach ($featureToCheck in $SharePointFeaturesArray)
		{
			$SiteCollection = $featureToCheck.SiteCollection
			$SubSiteURLCode = $featureToCheck.SubSiteURLCode
			$featureToCheckFeatureName = $featureToCheck.FeatureName
			#$featureToCheckWSPFile = $featureToCheck.WSPFile
			$featureToCheckScope = $featureToCheck.Scope
			#$featureToCheckLocation = $featureToCheck.Location
			$featureToCheckActiveOrInactive = $featureToCheck.ActiveOrInactive
			$featureToCheckStringToLookFor = $featureToCheck.StringToLookFor
			#$featureToCheckElementID = $featureToCheck.ElementID
			$featureToCheckPassFail = "FAIL"
			$featureToCheckValueFound = ""
			
			if($SiteCollection -eq $WAOrRC)
			{
				if ($SubSiteURLCode -ne "_NOT_APPLICABLE_")
				{
					$targetURL = $RMWSiteURL + $SubSiteURLCode
				}
				else
				{
					$targetURL = $RMWSiteURL
				}
				if($featureToCheckScope -eq "Site")
				{
					$featureName =(Get-SPFeature -Site $targetURL).GetTitle(1033)
					$featureDescription = (Get-SPFeature -Site $targetURL).GetDescription(1033)
				}
				else
				{
					$featureName =(Get-SPFeature -Web $targetURL).GetTitle(1033)
					$featureDescription = (Get-SPFeature -Web $targetURL).GetDescription(1033)
				}
				
				if($featureName -contains $featureToCheckFeatureName)
				{
					
					if($featureDescription -contains $featureToCheckStringToLookFor)
					{
						foreach($spfeature in $featureDescription)
						{
							if ($spfeature -eq $featureToCheckStringToLookFor)
							{
								$spfeatureCharLocation = $spfeature.LastIndexOf("v")
								$begin = $spfeatureCharLocation +1 
								$end = $spfeature.LastIndexOf("]")
								$num = $end - $begin
								$featureVersion = $spfeature.substring($begin,$num)
								$featureToCheckValueFound = "Active with version $featureVersion"
								$featureToCheckPassFail = "PASS"
								break
							}
							else
							{
								$featureToCheckValueFound = "Active with wrong description"
								$featureToCheckPassFail = "FAIL"
							}
						}
					}
					else
					{
						$featureToCheckValueFound = "Active"
						$featureToCheckPassFail = "PASS"
					}
				}
				else
				{
					$featureToCheckValueFound = "Couldn't find name"
					$featureToCheckPassFail = "FAIL"
				}
				
				$CheckSharepointFeaturesRow = $CheckSharepointFeaturesTable.NewRow()
				$CheckSharepointFeaturesRow["WA or RC"] = $WAOrRC
				$CheckSharepointFeaturesRow["Feature Name"] = $featureToCheckFeatureName
				$CheckSharepointFeaturesRow["Value Required"] = $featureToCheckActiveOrInactive
				$CheckSharepointFeaturesRow["Pass or Fail"] = $featureToCheckPassFail
				$CheckSharepointFeaturesRow["Date"] = $Date 
				$CheckSharepointFeaturesRow["Value"] = $featureToCheckValueFound
				$CheckSharepointFeaturesRow["Pass or Fail"] = $featureToCheckPassFail
				$CheckSharepointFeaturesTable.Rows.Add($CheckSharepointFeaturesRow)
				
			}
			
		}
	}
	else
	{
		Write-Output "`n--- No SharePoint Features to check ---"
	}
}

function RMWDataConnection ([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
	Write-Output "`n--- Checking Data Connections ---"
	$dataConnections =  $XMLConfig.XML.DataConnectionsArray.List
	$totalFiles = $dataConnections.Items
	$choices = $dataConnections.Choices
	$approvalOptions = @{}
	if ($dataConnections.SubSite -ne $null){
		foreach ($choice in $choices)
		{
			$key = $choice.Status
			$value = $choice.Choice
			$approvalOptions.add($key, $value)
		}
		
		if ($SubSiteURLCode -ne "_NOT_APPLICABLE_")
		{
			$targetURL = $RMWSiteURL + $SubSiteURLCode
		}
		else
		{
			$targetURL = $RMWSiteURL
		}
		$targetWeb = Get-SPWeb $targetURL
		
		$targetList = $targetWeb.Lists["Data Connection Library"]
		$itemsInList = $targetList.Items.Count
		$items = $targetList.Items
		$names = $items.DisplayName
		
		$row = $DataConnectionTable.NewRow()
		$row["In List?"] = "Totals items in list:"
		$row["Expected Value"] = $totalFiles
		$row["Pass or Fail"] = "Fail"
		$row["Date"] = $Date
		
		if ($itemsInList -eq $totalFiles)
		{
			$row["Actual Value"] = $itemsInList
			$row["Pass or Fail"] = "Pass"
			$DataConnectionTable.Rows.Add($row)
			
			foreach ($dataConnectionFile in $dataConnections.DataConnection)
			{
				$row = $DataConnectionTable.NewRow()
				$row["Data Connection"] = $dataConnectionFile.Name
				$row["Expected Value"] = $dataConnectionFile.ApprovalStatus
				$row["Pass or Fail"] = "Fail"
				$row["Date"] = $Date
				
				if ($names -contains $dataConnectionFile.Name){
					$row["In List?"] = "Yes"
				}
				foreach ($item in $items)
				{
					if ($item.DisplayName -eq $dataConnectionFile.Name)
					{
						$status = $item["Approval Status"]
						$statusRequired = $dataConnectionFile.ApprovalStatus
						$statusRequiredInt = $approvalOptions[$statusRequired]
						if($status -eq $statusRequiredInt)
						{
							$row["Actual Value"] = $statusRequired
							$row["Pass or Fail"] = "Pass"
						}
					}
				}
				$DataConnectionTable.Rows.Add($row)
			}
			
		}
		else
		{
			$row["In List?"] = "Totals items in list:"
			$row["Expected Value"] = $totalFiles
			$row["Actual Value"] = $itemsInList
			$row["Pass or Fail"] = "Fail"
			$row["Date"] = "Date"
			$DataConnectionTable.Rows.Add($row)
			$row = $DataConnectionTable.NewRow()
			$row["In List?"] = "Not Enough items in list"
			$DataConnectionTable.Rows.Add($row)
		}
		#$DataConnectionTable | format-table -AutoSize
	}
}

function RMWCheckInfoPathFormServices #Sharhad
{
#RMWCheckInfoPathFormServices: Check InfoPath Form Services in Central Admin  - Michael

	#Get array from XML
	$InfoPathFormsServicesArray = $XMLConfig.XML.InfoPathFormServicesArray.SettingInfo
	
	#Get InfoPath Forms Service
	$SPInfoPathFormsServices = Get-SPInfoPathFormsService

	#If are entries in the XML
	if (($InfoPathFormsServicesArray -ne $null) -or ($InfoPathFormsServicesArray.Count -gt 0))
	{
		Write-Output "`n--- Checking InfoPath Form Services Settings ---"

		foreach ($settingToCheck in $InfoPathFormsServicesArray)
		{
			#get the value of the setting
			$settingToCheckSettingName = $settingToCheck.SettingName
			$settingToCheckSettingNameforTable = $settingToCheck.TableName
			$settingToCheckRequiredValue = $settingToCheck.RequiredValue
			$SettingTOCheckResult = $SPInfoPathFormsServices.$settingToCheckSettingName
			
			##add to table
			if ($SettingTOCheckResult -eq $settingToCheckRequiredValue)
			{
				$SettingToCheckPassFail = "PASS"
			}
			else 
			{
				$SettingToCheckPassFail = "FAIL"
			}
			
			#Add result to InfoPath Form Services table
			$InfoPathFormServicesInfoRow = $InfoPathFormServicesTable.NewRow()
			$InfoPathFormServicesInfoRow["Property"] = $settingToCheckSettingNameforTable
			$InfoPathFormServicesInfoRow["Value"] = $SettingTOCheckResult
			$InfoPathFormServicesInfoRow["Value Required"] = $SettingToCheckRequiredValue
			$InfoPathFormServicesInfoRow["Pass or Fail"] = $SettingToCheckPassFail
			$InfoPathFormServicesInfoRow["Date"] = $Date 
			$InfoPathFormServicesTable.Rows.Add($InfoPathFormServicesInfoRow)
		}
	}
}
 
function RMWCheckInfoPathForms ([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
#RMWCheckInfoPathForms: Check Form Templates library in RMW to see what forms have been deployed  - Michael

	#Get array from XML
	$InfoPathFormsArray = $XMLConfig.XML.InfoPathFormsArray.FormInfo
	
	#If are entries in the XML
	if (($InfoPathFormsArray -ne $null) -or ($InfoPathFormsArray.Count -gt 0))
	{
		Write-Output "`n--- Checking Deployed Form Templates ---"
		
		#TODO: set global variable for InfoPathForms table to true
		$SPSiteRMW = Get-SPSite $RMWSiteURL
		$RMWrootweb = $SPSiteRMW.rootweb
		$InfoPathFormTemplateLibrary = $RMWrootweb.Lists["Form Templates"]
		$NameField = $InfoPathFormTemplateLibrary.Fields["Name"]
		$NameFieldInternalName = $NameField.InternalName
		
		foreach ($formToCheck in $InfoPathFormsArray)
		{
			$formToCheckFormName = $formToCheck.FormName
			$formToCheckPassFail = "FAIL"
			$formToCheckValueFound = ""
			
			#Can do a simple CAML query to see if the form templates library has value
			$camlQuery = "<Where><Contains><FieldRef Name=""$NameFieldInternalName""/><Value Type=""text"">$formToCheckFormName</Value></Contains></Where>"
			$spQuery = New-Object Microsoft.SharePoint.SPQuery
			$spQuery.Query = $camlQuery
			$spQuery.ViewAttributes = "Scope='Recursive'"

			$Forms = $InfoPathFormTemplateLibrary.GetItems($spQuery)
			if ($Forms.Count -eq 1)
			{
				$formToCheckValueFound = "True"
				$formToCheckPassFail = "PASS"
			}
			else 
			{
				$formToCheckValueFound = "False"
				$formToCheckPassFail = "FAIL"
			}
			
			#Add result to InfoPath Form Templates table
			$InfoPathFormTemplatesInfoRow = $InfoPathFormTemplatesTable.NewRow()
			$InfoPathFormTemplatesInfoRow["Form Name"] = $formToCheckFormName
			$InfoPathFormTemplatesInfoRow["Uploaded"] = $formToCheckValueFound
			$InfoPathFormTemplatesInfoRow["Pass or Fail"] = $formToCheckPassFail
			$InfoPathFormTemplatesInfoRow["Date"] = $Date 
			$InfoPathFormTemplatesTable.Rows.Add($InfoPathFormTemplatesInfoRow)
		}
	}
	else
	{
		Write-Output "`n--- No Form Templates to check ---"
	}
}

function RMWCheckSiteSettings([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
#RMWCheckSiteSettings: Check RMW Site Settings - Simon
# Parameters:
# 	-> WAOrRC: Whether it is the WA (Work Area) or RC (Records Center) that is to be verified
# 	-> RMWSiteURL: URL of the RMW site collection you want to verify
		
	#Get array from XML
	$SiteSettingsArray = $XMLConfig.XML.SiteSettingsArray.SettingInfo
	
	#If are entries in the XML
	if (($SiteSettingsArray -ne $null) -or ($SiteSettingsArray.Count -gt 0))
	{
		if ($WAorRC -eq "WA")
		{
			Write-Output "`n--- Checking Work Area Site Settings ---"
		}
		elseif ($WAorRC -eq "RC")
		{
			Write-Output "`n--- Checking Record Center Site Settings ---"
		}
		
		#Write-Output "`n--- Checking Site Settings ---"
		
		#TODO: set global variable for Site Settings table to true
		
		$SPSiteRMW = Get-SPSite $RMWSiteURL
		$RMWrootweb = $SPSiteRMW.rootweb
		
		#Loop through each setting in the array
		foreach ($SettingToCheck in $SiteSettingsArray)
		{
			#Get the relevant information from the XML file
			$SiteCollection = $SettingToCheck.SiteCollection
			$SettingToCheckSettingType = $SettingToCheck.SettingType
			$SettingToCheckSettingName = $SettingToCheck.SettingName
			$SettingToCheckRequiredValue = $SettingToCheck.RequiredValue
			$SettingToCheckAllPropertiesValue = $SettingToCheck.AllPropertiesValue
			$SettingToCheckAllPropertiesType = $SettingToCheck.AllPropertiesType
			$SettingToCheckPassFail = "FAIL"
			$SettingToCheckValueFound = ""
			$SettingToCheckValueType = ""
			$SettingTOCheckResult = ""
			
			#If the setting is for the current site collection
			if ($SiteCollection -eq $WAOrRC)
			{
				#Check the Setting type
				if ($SettingToCheckSettingType -eq "Content Organizer Settings")
				{
					$SettingToCheckValueFound = $RMWrootweb.AllProperties[$SettingToCheckAllPropertiesValue]
					
					#Check the value type
					if ($SettingToCheckAllPropertiesType -eq "ENABLED_DISABLED")
					{
						if (($SettingToCheckRequiredValue -eq "ENABLED") -or ($SettingToCheckRequiredValue -eq "DISABLED"))
						{
							if ($SettingToCheckValueFound -eq $true)
							{
								$SettingTOCheckResult = "ENABLED"
							}
							else
							{
								$SettingTOCheckResult = "DISABLED"
							}
							
							if ($SettingToCheckRequiredValue -eq $SettingTOCheckResult)
							{
								$SettingToCheckPassFail = "PASS"
							}
						}
						else
						{
							Write-Output "`nError: Unable to handle required value of $SettingToCheckRequiredValue"
							$SettingTOCheckResult = "Unable to handle required value of $SettingToCheckRequiredValue"
						}
					}
					elseif ($SettingToCheckAllPropertiesType -eq "String")
					{
						$SettingTOCheckResult = $SettingToCheckValueFound
						
						if ($SettingToCheckRequiredValue -eq $SettingToCheckValueFound)
						{
							$SettingToCheckPassFail = "PASS"
						}
					}
					else
					{
						Write-Output "`nError: Unable to handle setting name of $SettingToCheckSettingName for setting type of $SettingToCheckSettingType"
						$SettingTOCheckResult = "Unable to handle setting name of $SettingToCheckSettingName for setting type of $SettingToCheckSettingType"
					}
				}
				elseif($SettingToCheckSettingType -eq "Record Declaration Settings")
				{
					$SettingToCheckValueFound = $RMWrootweb.AllProperties[$SettingToCheckAllPropertiesValue]
					
					#Check the value type
					if ($SettingToCheckAllPropertiesType -eq "AVAILABLE_NOTAVAILABLE")
					{
						if (($SettingToCheckRequiredValue -eq "Available") -or ($SettingToCheckRequiredValue -eq "Not Available"))
						{
							if ($SettingToCheckValueFound -eq "False")
							{
								$SettingTOCheckResult = "Not Available"
							}
							else
							{
								$SettingTOCheckResult = "Available"
							}
							
							if ($SettingToCheckRequiredValue -eq $SettingTOCheckResult)
							{
								$SettingToCheckPassFail = "PASS"
							}
						}
						else
						{
							Write-Output "`nError: Unable to handle required value of $SettingToCheckRequiredValue"
							$SettingTOCheckResult = "Unable to handle required value of $SettingToCheckRequiredValue"
						}
					}
					elseif ($SettingToCheckAllPropertiesType -eq "String")
					{
						$SettingTOCheckResult = $SettingToCheckValueFound
						
						if ($SettingToCheckRequiredValue -eq $SettingToCheckValueFound)
						{
							$SettingToCheckPassFail = "PASS"
						}
					}
					else
					{
						Write-Output "`nError: Unable to handle setting name of $SettingToCheckSettingName for setting type of $SettingToCheckSettingType"
						$SettingTOCheckResult = "Unable to handle setting name of $SettingToCheckSettingName for setting type of $SettingToCheckSettingType"
					}
				}
				elseif($SettingToCheckSettingType -eq "Site Collection Audit Settings")
				{
					#TODO: Here is how you can set audit settings, turn it around to be able to verify
					$SettingToCheckSettingName = "The following parameters are Enabled:
					Opening or downloading documents, viewing items in lists, or viewing item properties, 
					Editing Items, 
					Checking out or checking in items, 
					Moving or copying items to another location in the site, 
					Deleting or restoring items, 
					Editing content types and columns, 
					Editing users and permissions"
					
					if($SPSiteRMW.Audit.AuditFlags -eq $SettingToCheckRequiredValue)
					{
						$SettingTOCheckResult = "YES"
						$SettingToCheckPassFail = "PASS"
					}
					else
					{
						$SettingTOCheckResult = "NO"
						$SettingToCheckPassFail = "FAIL"
					}
					
					$SettingToCheckRequiredValue = "YES"
					
				}
				elseif($SettingToCheckSettingType -eq "Content Organizer Rule")
				{
					$RoutingRuleArray = $SettingToCheck.RoutingRule
					$ContentTypeToCheck = $SettingToCheck.SettingName #New value for content type
					
					foreach ($RoutingRuleHistoryXMLElement in $RoutingRuleArray)
					{
						$List = $RMWrootweb.lists["Content Organizer Rules"]
						$SettingToCheckSettingName = $RoutingRuleHistoryXMLElement.SettingName
						$rulename = $RoutingRuleHistoryXMLElement.RuleName
						
						$TitleFieldInternalName = $list.Fields["Title"].InternalName
						$SubmissionContentTypeInternalName = $list.Fields["Submission Content Type"].InternalName

						$camlQuery = "<Where><And><Eq><FieldRef Name=""$TitleFieldInternalName"" /><Value Type=""text"">$rulename</Value></Eq><Eq><FieldRef Name=""$SubmissionContentTypeInternalName"" /><Value Type=""text"">$ContentTypeToCheck</Value></Eq></And></Where>"
						$spQuery = New-Object Microsoft.SharePoint.SPQuery
						$spQuery.Query = $camlQuery
						$spQuery.ViewAttributes = "Scope='Recursive'"
						$spQuery.RowLimit = 1

						$RoutingRuleItem = $list.GetItems($spQuery)
						
						$SettingToCheckRequiredValue = $RoutingRuleHistoryXMLElement.RequiredValue
						
						$SettingTOCheckResult = $RoutingRuleItem[0][$SettingToCheckSettingName]
						#$SettingTOCheckResult = $RoutingRuleItem[$SettingToCheckSettingName]
						
						#If the value to compare is a boolean, perform this block of code
						if ($SettingToCheckSettingName -eq "Active")
						{
							if ($SettingTOCheckResult -eq $true)
							{
								$SettingTOCheckResult = "Yes"
							}
							else
							{
								$SettingTOCheckResult = "No"
							}
						}

						if ($SettingTOCheckResult -eq $SettingToCheckRequiredValue)
						{
							$SettingToCheckPassFail = "PASS"
						}
						else
						{
							$SettingToCheckPassFail = "FAIL"
						}
						
						#Add result to SiteSettings table
						$SiteSettingsRow = $SiteSettingsTable.NewRow()
						$SiteSettingsRow["WA or RC"] = $WAOrRC
						$SiteSettingsRow["Setting Type"] = $SettingToCheckSettingType
						$SiteSettingsRow["Setting Name"] = $SettingToCheckSettingName
						$SiteSettingsRow["Value"] = $SettingTOCheckResult
						$SiteSettingsRow["Value Required"] = $SettingToCheckRequiredValue
						$SiteSettingsRow["Pass or Fail"] = $SettingToCheckPassFail
						$SiteSettingsRow["Date"] = $Date 
						$SiteSettingsTable.Rows.Add($SiteSettingsRow)
					}
				}
				else
				{
					Write-Output "`nError: Unable to handle setting type of $SettingToCheckSettingType"
					$SettingTOCheckResult = "Unable to handle setting type of $SettingToCheckSettingType"
				}
				
				if ($SettingToCheckSettingType -ne "Content Organizer Rule")
				{
				#Add result to SiteSettings table for all setting types that are not part of Content Organizer Rule
				$SiteSettingsRow = $SiteSettingsTable.NewRow()
				$SiteSettingsRow["WA or RC"] = $WAOrRC
				$SiteSettingsRow["Setting Type"] = $SettingToCheckSettingType
				$SiteSettingsRow["Setting Name"] = $SettingToCheckSettingName
				$SiteSettingsRow["Value"] = $SettingTOCheckResult
				$SiteSettingsRow["Value Required"] = $SettingToCheckRequiredValue
				$SiteSettingsRow["Pass or Fail"] = $SettingToCheckPassFail
				$SiteSettingsRow["Date"] = $Date 
				$SiteSettingsTable.Rows.Add($SiteSettingsRow)
				}
			}
		}
		
		$RMWrootweb.dispose()
		$SPSiteRMW.dispose()
	}
	else
	{
		Write-Output "`n--- No Site Settings to check ---"
	}
}

function RMWCheckLibrarySettings([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
#RMWCheckLibrarySettings: Check RMW Library Settings  - Simon
# Parameters:
# 	-> WAOrRC: Whether it is the WA (Work Area) or RC (Records Center) that is to be verified
# 	-> RMWSiteURL: URL of the RMW site collection you want to verify

	#Get array from XML
	$LibrarySettingsArray = $XMLConfig.XML.LibrarySettingsArray.SettingInfo
	
	#If are entries in the XML
	if (($LibrarySettingsArray -ne $null) -or ($LibrarySettingsArray.Count -gt 0))
	{
		if ($WAorRC -eq "WA")
		{
			Write-Output "`n--- Checking Work Area Library Settings ---"
		}
		elseif ($WAorRC -eq "RC")
		{
			Write-Output "`n--- Checking Record Center Library Settings ---"
		}
		
		#TODO: set global variable for LibrarySettings table to true
		
		
		#Loop through each setting in the array
		foreach ($SettingToCheck in $LibrarySettingsArray)
		{
			#Get the relevant information from the XML file
			$SiteCollection = $SettingToCheck.SiteCollection
			$SubSiteURLCode = $SettingToCheck.SubSiteURLCode
			$SettingToCheckLibraryName = $SettingToCheck.LibraryName
			$VerificationStepName = $SettingToCheck.VerificationStep
			$SettingToCheckSettingName = $SettingToCheck.SettingName
			$SettingToCheckRequiredValue = $SettingToCheck.RequiredValue
			$SettingToCheckPassFail = "FAIL"
			$SettingTOCheckResult = ""
			
			#If the setting is for the current site collection
			if ($SiteCollection -eq $WAOrRC)
			{
				if ($SubSiteURLCode -ne "_NOT_APPLICABLE_")
				{
					$targetURL = $RMWSiteURL + $SubSiteURLCode
				}
				else
				{
					$targetURL = $RMWSiteURL
				}
				
				$targetWeb = Get-SPWeb $targetURL
				
				#Get the list
				$targetList = $targetWeb.Lists[$SettingToCheckLibraryName]
				
				if ($targetList -ne $null)
				{
					#Checking Library Version Setting
					if ($SettingToCheckSettingName -eq "Version Setting")
					{
						$versionHistoryXMLElement = $SettingToCheck.VersionHistory #Retrieving the sub array
						$SettingToCheckRequiredValue = $versionHistoryXMLElement.RequiredValue #Retrieving sub-array required value
						$SettingToCheckSettingName = $versionHistoryXMLElement.SettingName #Retrieving sub array settingname
						$VerificationStepName = $versionHistoryXMLElement.VerificationStep #Retrieving sub array verification step
						
							if ($SettingToCheckSettingName -eq "Create major versions") #Look for Document Version History
							{
								if ($targetList.EnableVersioning -eq $true -and $targetList.EnableMinorVersions -eq $false)
								{
									$SettingTOCheckResult = "Major"
								}
								elseif ($targetList.EnableVersioning -eq $false -and $targetList.EnableMinorVersions -eq $true)
								{
									$SettingTOCheckResult = "Minor"
								}
								else
								{
									$SettingTOCheckResult = "None"
								}
							}
							elseif ($SettingToCheckSettingName -eq "Keep the following number of major versions")
							{
								if ($targetList.majorversionlimit -eq 0)
								{
									$SettingTOCheckResult = "NO"
								}
								else
								{
									$SettingTOCheckResult = "YES"
								}
							}
							
					}
					#Verifying if content type is associated with the library
					Elseif ($SettingToCheckSettingName -eq "Content type")
					{
						foreach ($ContentType in $targetList.ContentTypes) #Verify that the required content type is associated to the library by looping through all content types in that library
						{ 
							if ($ContentType.Name -eq $SettingToCheckRequiredValue) 
							{ 
								$SettingTOCheckResult = "YES" #Required content type is associated to the library. Exit loop.
								break
							}
							else
							{
								$SettingTOCheckResult = "NO"
							}
						}
						$SettingToCheckRequiredValue = "YES"
					}
					Elseif ($SettingToCheckSettingName -eq "Content Type Advanced Setting")
					{
						$ContentTypeXMLElement = $SettingToCheck.ContentType #Retrieving the sub array
						$SettingToCheckRequiredValue = $ContentTypeXMLElement.RequiredValue #Retrieving sub-array required value
						$SettingToCheckSettingName = $ContentTypeXMLElement.SettingName #Retrieving sub array settingname
						$VerificationStepName = $ContentTypeXMLElement.VerificationStep #Retrieving sub array verification step
						$ContentTypeName = $ContentTypeXMLElement.ContentTypeName
						$ContentType = $targetList.ContentTypes | Where {$_.Name -Match $ContentTypeName}

						if ($SettingToCheckSettingName -eq "Document Template") #Looking for associated document template
						{
							if ($SettingToCheckRequiredValue -eq "_NOT_EMPTY_")
							{
								$CurrentTemplateurl = $ContentType.documenttemplateurl
								if(($CurrentTemplateurl -ne "") -and ($CurrentTemplateurl -ne $Null)) 
								{
									$SettingTOCheckResult = "YES"
								}
								else
								{
									$SettingTOCheckResult = "NO"
								}
							}
							elseif ($ContentType.DocumentTemplate -eq $SettingToCheckRequiredValue)
							{
								$SettingTOCheckResult = "YES"
							}
							else
							{
								$SettingTOCheckResult = "NO"
							}
							$SettingToCheckRequiredValue = "YES"
						}
						else #Verifying read only status
						{
							if ($ContentType.ReadOnly -eq $SettingToCheckRequiredValue)
							{
								$SettingTOCheckResult = "YES"
							}
							else
							{
								$SettingTOCheckResult = "NO"
							}
						}
					}
					ElseIf ($SettingToCheckSettingName -eq "Library Advanced Setting")
					{
						$AdvancedSettingXMLElement = $SettingToCheck.Advancedsetting #Retrieving the sub array
						$SettingToCheckRequiredValue = $AdvancedSettingXMLElement.RequiredValue #Retrieving sub-array required value
						$SettingToCheckSettingName = $AdvancedSettingXMLElement.SettingName #Retrieving sub array settingname
						$VerificationStepName = $AdvancedSettingXMLElement.VerificationStep #Retrieving sub array verification step
						
						if ($SettingToCheckSettingName -eq "Allow management of content types")
						{
							if ($targetlist.ContentTypesEnabled -eq $true) #Allow management of content types is set to Yes
							{
								$SettingTOCheckResult = "YES"
							}
							else
							{
								$SettingTOCheckResult = "NO"
							}
						}
						elseif($SettingToCheckSettingName -eq "Use the server default")
						{
							if ($targetlist.DefaultItemOpenUseListSetting -eq $false) #Verify that Default open behavior for browser-enabled documents is: Use the server default
							{
								$SettingTOCheckResult = "Use the server default"
							}
							elseif ($targetlist.DefaultItemOpenUseListSetting -eq "PreferClient")
							{
								$SettingTOCheckResult = "Open in the client application"
							}
							else
							{
								$SettingTOCheckResult = "Open in the browser"
							}
						}
					}
				}
					
				
				else
				{
					Write-Output "`nError: Unable to find list named $SettingToCheckLibraryName"
					$SettingTOCheckResult = "Unable to find list named $SettingToCheckLibraryName"
				}
				
				#Determine if step is pass or fail	
				if ($SettingTOCheckResult -eq $SettingToCheckRequiredValue)
				{
					$SettingToCheckPassFail = "PASS"
				}
				else
				{
					$SettingToCheckPassFail = "FAIL"
				}	
				
				#TODO: Add result to LibrarySettings table
				$LibrarySettingsRow = $LibrarySettingsTable.NewRow()
				$LibrarySettingsRow["WA or RC"] = $WAOrRC
				$LibrarySettingsRow["Library Name"] = $SettingToCheckLibraryName
				$LibrarySettingsRow["Verification Step"] = $VerificationStepName
				$LibrarySettingsRow["Value"] = $SettingTOCheckResult
				$LibrarySettingsRow["Value Required"] = $SettingToCheckRequiredValue
				$LibrarySettingsRow["Pass or Fail"] = $SettingToCheckPassFail
				$LibrarySettingsRow["Date"] = $Date 
				$LibrarySettingsTable.Rows.Add($LibrarySettingsRow)
			}
		}
		
		if ($targetWeb -ne $null)
		{
			$targetWeb.dispose()
		}
	}
	else
	{
		Write-Output "`n--- No Library Settings to check ---"
	}
}

function RMWCheckWorkflowConstants([string]$WAOrRC, [string]$RMWSiteURL) #Shovan
{
	#RMWCheckWorkflowConstants: Check Workflow Constants - Shovan

	#This is needed to update the Nintex Workflow Constants
	[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow") | Out-Null

	$workflowConstantSettings = $XMLConfig.XML.WorkflowConstantsArray.WorkflowConstantInfo
	
	if (($workflowConstantSettings -ne $null) -or ($workflowConstantSettings.Count -gt 0))
	{
		Write-Output "`n--- Checking Workflow Constants ---"
				
		foreach ($settingToCheck in $workflowConstantSettings)
		{
			$settingToCheckName = $settingToCheck.Name
			$settingToCheckScope = $settingToCheck.Scope
			$settingToCheckConstantType = $settingToCheck.ConstantType
			$settingToCheckValue = $settingToCheck.Value
			$settingToCheckPassFail = "PASS"
			$settingToCheckValueFound = ""
			
			$targetURL = $RMWSiteURL 
			$targetSite = Get-SPSite $targetURL
			$targetRootWeb = $targetSite.rootweb
			
			$WFConstantCollection = [Nintex.Workflow.WorkflowConstantCollection]::GetWorkflowConstants($targetRootWeb.id,$targetSite.ID)
			$workflowConstant = $WFConstantCollection | Where {$_.Title -eq $settingToCheckName}
			$workflowConstantType = $workflowConstant.Type
			$workflowConstantValue = $workflowConstant.Value
			

			
			
			if ($settingToCheckConstantType -eq $workflowConstantType)
			{
				if ($settingToCheckConstantType -eq "Credential")
				{
					if ($settingToCheckValue -eq "_NOT_EMPTY_")
					{
						if (($workflowConstantValue -match "Username") -and ($workflowConstantValue -match "Password"))
						{
							$settingToCheckPassFail = "PASS"
							$settingToCheckValueFound = "Credential has username and password"
						}
						else
						{
							if ($workflowConstantValue -notmatch "Username")
							{
								Write-Output "Workflow constant named $WorkflowConstantName is missing a Username"
								$settingToCheckValueFound = "Credential missing username"
							}
							
							if ($workflowConstantValue -notmatch "Password")
							{
								Write-Output "Workflow constant named $WorkflowConstantName is missing a Password"
								$settingToCheckValueFound = $settingToCheckValueFound + " Credential missing password"
							}
							$settingToCheckPassFail = "FAIL"
						}
					}
				}
				
				elseif ($settingToCheckConstantType -eq "String")
				{
					$settingToCheckValueFound = $workflowConstantValue
					
					if ($settingToCheckValueFound -eq $settingToCheckValue)
					{
							$settingToCheckPassFail = "PASS"
					}
					else
					{
							$settingToCheckPassFail = "FAIL"
					}
					
				}
				else
				{
					Write-Output "`nError unable to handle settingToCheckConstantType of $settingToCheckConstantType"
					$settingToCheckPassFail = "FAIL"
					$settingToCheckValueFound = "Error unable to check"
				}
			}
			else
			{
				Write-Output "`nError $workflowConstantType does not match $settingToCheckConstantType" 
				$settingToCheckValueFound = "Error unable to check"
				$settingToCheckPassFail = "FAIL"
			}
			
			
		    $WorkflowConstantRow = $WorkflowConstantTable.NewRow()
			$WorkflowConstantRow["Workflow Constant Name"] = $settingToCheckName
			$WorkflowConstantRow["Value"] = $settingToCheckValueFound
			$WorkflowConstantRow["Value Required"] = $settingToCheckValue
			$WorkflowConstantRow["Pass or Fail"] = $settingToCheckPassFail
			$WorkflowConstantRow["Date"] = $Date 
			$WorkflowConstantTable.Rows.Add($WorkflowConstantRow)
		}
	}
	else
	{
		Write-Output "`n--- No Workflow Constants to check ---"
	}
	
	$targetRootWeb.dispose()
	$targetSite.dispose()
}

function RMWCheckNintexWorkflows ([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
	Write-Output "`n--- Checking Nintex Workflows ---"
	
	if ($SubSiteURLCode -ne "_NOT_APPLICABLE_")
	{
		$targetURL = $RMWSiteURL + $SubSiteURLCode
	}
	else
	{
		$targetURL = $RMWSiteURL
	}
	
	$rmwWeb = Get-SPWeb ($targetURL)
	$targetWeb = $rmwWeb
	$workflowsInfo =  $XMLConfig.XML.NintexWorkFlowArray.WorkflowInfo
	$workflowArray = @()
	foreach($item in $workflowsInfo)
	{
		$SubSiteURLCode = $item.SubSiteURLCode		
		
		if ($item.WorkflowType -eq "List" -And $item.SubSiteURLCode -eq "RMW")
		{
			$list = $targetWeb.lists[$item.LibraryName]
			$workflowInList = $list.WorkflowAssociations.Name
			$workflowInList = $workflowInList | select -uniq 
			$workflowName = $item.WorkflowName
			foreach ($wfname in $workflowInList)
			{
				if ($wfname -notmatch "Previous"){
					$workflowArray += $wfname
				}
			}
			$workflowArray = $workflowArray | select -uniq
		}
		elseif ($item.WorkflowType -eq "Site" -And $item.SubSiteURLCode -eq "RMW")
		{
			$workflowInList = $targetWeb.WorkflowTemplates.Name
			$workflowInList = $workflowInList | select -uniq 
			$workflowName = $item.WorkflowName
			foreach ($wfname in $workflowInList)
			{
				if ($wfname -notmatch "Previous"){
					$workflowArray += $wfname
				}
			}
			$workflowArray = $workflowArray | select -uniq
		}
		elseif ($item.WorkflowType -eq "List" -And $item.SubSiteURLCode -eq "Product")
		{
			$targetWeb = Get-SPWeb ($RMWSiteURL + $product)
			[String]$productListName = $product + " " + $item.LibraryName
			$list = $targetWeb.lists[$productListName]
			$workflowInList = $list.WorkflowAssociations.Name
			$workflowInList = $workflowInList | select -uniq 
			$workflowName = $item.WorkflowName
			foreach ($wfname in $workflowInList)
			{
				if ($wfname -notmatch "Previous"){
					$workflowArray += $wfname
				}
			}
			$workflowArray = $workflowArray | select -uniq
		}
		elseif ($item.WorkflowType -eq "List" -And $item.SubSiteURLCode -eq "Study")
		{
			$targetWeb = Get-SPWeb ($RMWSiteURL + $study)
			[String]$productListName = $study + " " + $item.LibraryName	
			$list = $targetWeb.lists[$productListName]
			$workflowInList = $list.WorkflowAssociations.Name
			$workflowInList = $workflowInList | select -uniq 
			$workflowName = $item.WorkflowName
			foreach ($wfname in $workflowInList)
			{
				if ($wfname -notmatch "Previous"){
					$workflowArray += $wfname
				}
			}
			$workflowArray = $workflowArray | select -uniq
		}
		elseif ($item.SubSiteURLCode -eq "RMW-RC")
		{
			$targetWeb = Get-SPWeb ($RMWSiteURL.Substring(0,$RMWSiteURL.Length-1) + "-RC/")	
			$list = $targetWeb.lists[$item.LibraryName]
			$workflowInList = $list.WorkflowAssociations.Name
			$workflowInList = $workflowInList | select -uniq 
			$workflowName = $item.WorkflowName
			foreach ($wfname in $workflowInList)
			{
				if ($wfname -notmatch "Previous"){
					$workflowArray += $wfname
				}
			}
			$workflowArray = $workflowArray | select -uniq
		}
	}
	foreach ($workflow in $workflowsInfo)
	{
		$row = $WorkflowTable.NewRow()
		$row["Workflow Name"] = $workflow.WorkflowName
		$row["Library"] = $workflow.LibraryName
		if ($workflow.SubSiteURLCode -eq "Product"){
			$row["Library"] = $product + " " + $workflow.LibraryName
		}
		if ($workflow.SubSiteURLCode -eq "Study"){
			$row["Library"] = $study + " " + $workflow.LibraryName
		}
		$row["Value"] = "Does not Exist"
		$row["Value Required"] = "Exists"
		$row["Pass or Fail"] = "FAIL" 
		$row["Date"] = Get-Date -Format g
		if ($workflow.WorkflowType -eq "Site"){
			$row["Library"] = $workflow.LibraryName + " Site"
		}
		if ($workflowArray.Contains($workflow.WorkflowName))
		{
			$row["Value"] = "Exists"
			$row["Pass or Fail"] = "PASS"
		}
		$WorkflowTable.Rows.Add($row)
	} 
	
	
	if ($targetWeb -ne $null)
	{
		$targetWeb.dispose()
	}

}

function RMWCheckCoSignSettings ([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
	Write-Output "`n--- Checking CoSign Settings ---"
	$settings = $XMLConfig.XML.CoSignSettingsArray
	if ($settings.List.SubSiteURLCode -ne "_NOT_APPLICABLE_")
	{
		$targetURL = $RMWSiteURL + $SubSiteURLCode
	}
	else
	{
		$targetURL = $RMWSiteURL
	}

	
	$rmwWeb = Get-SPWeb ($targetURL)
	$targetWeb = $rmwWeb
	
	foreach ($item in $settings.List)
	{
		$row = $CoSignSettingsTable.NewRow()
		$row["CoSign Setting"] = $item.SubSiteURLCode + " CoSign Settings in " + $item.ListName
		$CoSignSettingsTable.Rows.Add($row)
		
		if ($item.SubSiteURLCode -eq "Study"){$targetWeb = Get-SPWeb ($RMWSiteURL + $study)}
		elseif ($item.SubSiteURLCode -eq "Product"){$targetWeb = Get-SPWeb ($RMWSiteURL + $product)}
		
		$targetList = $targetWeb.Lists[$item.ListName]
		
		$listID = $targetList.ID.Guid.ToLower()
		$key = "arx_" + $listID
		#check if key exists
		if ($targetWeb.AllProperties.ContainsKey($key))
		{
			$CoSignSettings = $item.AttributeArray.AttributeInfo
			
			#Get the xml for the coSign properties
			$xmlDoc = New-Object System.XML.XMLDocument
			$xmlDoc.LoadXML($targetWeb.AllProperties[$key].ToString())
			
			$parentNode = $xmlDoc.SelectSingleNode("/ARXDocLibSetting")
			foreach($attribute in $CoSignSettings)
			{
				#List settings
				if ($attribute.AttributeLevel -eq "Library Settings")
				{
					$predefReasonExists = $false
					$row = $CoSignSettingsTable.NewRow()
					$row["CoSign Setting"] = $attribute.AttributeDefinition
					$row["Setting Location"] = $attribute.AttributeLevel
					$row["Expected Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
					$row["Actual Value"] = ""
					$row["Pass or Fail"] = "FAIL"
					$row["Date"] = $Date
					$tempAttribute = $attribute.AttributeValue
					if ($tempAttribute -match ",")
					{
						$tempAttribute = $attribute.AttributeValue.Replace(", ", ",").Split(",")
						foreach ($reason in $tempAttribute)
						{
							if ($parentNode.($attribute.AtributeName) -contains $reason){$predefReasonExists = $true}
							else {$predefReasonExists = $false}
						}
					}
					if ($parentNode.($attribute.AtributeName) -eq $tempAttribute -or $predefReasonExists)
					{
						
						$row["Actual Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} 
											   elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
						$row["Pass or Fail"] = "PASS"
					}
					$CoSignSettingsTable.Rows.Add($row)
				}
			}
			$CoSignSettingsContentType = $parentNode.ContentType
			$contentTypes = $item.ContentType.Replace(", ", ",").Split(",")
			foreach ($contentType in $contentTypes)
			{
				$row = $CoSignSettingsTable.NewRow()
				$row["CoSign Setting"] = $contentType
				$CoSignSettingsTable.Rows.Add($row)
				foreach($attribute in $CoSignSettings)
				{
					if ($attribute.AttributeLevel -eq "Content Type Settings"){
						$row = $CoSignSettingsTable.NewRow()
						$row["CoSign Setting"] = $attribute.AttributeDefinition
						$row["Setting Location"] = $attribute.AttributeLevel
						$row["Expected Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
						$row["Actual Value"] = ""
						$row["Pass or Fail"] = "FAIL"
						$row["Date"] = $Date
						if ($CoSignSettingsContentType.($attribute.AtributeName) -eq $attribute.AttributeValue)
						{
							$row["Actual Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
							$row["Pass or Fail"] = "PASS"
						}
						$CoSignSettingsTable.Rows.Add($row)
					}
				}
				$CoSignSettingsContentTypeParameters = $parentNode.ContentType.NewSigFieldSet			
				foreach($attribute in $CoSignSettings)
				{
					if ($attribute.AttributeLevel -eq "Content Type Check Settings"){
						$row = $CoSignSettingsTable.NewRow()
						$row["CoSign Setting"] = $attribute.AttributeDefinition
						$row["Setting Location"] = if($attribute.AttributeLevel -eq "Content Type Check Settings"){"Content Type Settings";} else {$attribute.AttributeLevel;}
						$row["Expected Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
						$row["Actual Value"] = ""
						$row["Pass or Fail"] = "FAIL"
						$row["Date"] = $Date
						if ($CoSignSettingsContentTypeParameters.($attribute.AtributeName) -eq $attribute.AttributeValue)
						{
							$row["Actual Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
							$row["Pass or Fail"] = "PASS"
						}
						$CoSignSettingsTable.Rows.Add($row)
					}
				}
				if ($item.SignatureProfile -eq "Yes")
				{
					$contentTypeToUpdate = $targetList.ContentTypes[$item.ContentType]
					$contentTypeToUpdateParentIDString = $contentTypeToUpdate.parent.Id.ToString().ToLower()
					$contentTypeToUpdateNode = $xmlDoc.SelectSingleNode("//ContentType[@ContentTypeID=""$contentTypeToUpdateParentIDString""]")
					
					if ($contentTypeToUpdateNode -eq $null)
					{
						Write-Output "`n`n### ERROR ### Could not find arx settings for content type named $contentTypeToUpdateName"
					}
					else{
						$signatureProfile = $contentTypeToUpdateNode.Field
						foreach($attribute in $CoSignSettings)
						{
							if ($attribute.AttributeLevel -eq "Signature Profile Settings"){
								$row = $CoSignSettingsTable.NewRow()
								$row["CoSign Setting"] = $attribute.AttributeDefinition
								$row["Setting Location"] = $attribute.AttributeLevel
								$row["Expected Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
								$row["Actual Value"] = ""
								$row["Pass or Fail"] = "FAIL"
								$row["Date"] = $Date
								if ($signatureProfile.($attribute.AtributeName) -eq $attribute.AttributeValue)
								{
									$row["Actual Value"] = if ($attribute.AttributeValue -eq "true"){"Checked";} elseif($attribute.AttributeValue -eq "false") {"Unchecked";} else {$attribute.AttributeValue;}
									$row["Pass or Fail"] = "PASS"
								}
								$CoSignSettingsTable.Rows.Add($row)
							}
						}
					}
				}
			}	
		}
		$row = $CoSignSettingsTable.NewRow()
		$CoSignSettingsTable.Rows.Add($row)
	}	
	$targetWeb.dispose()
}

function RMWContentTypeAssociation ([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad
{
	$items = $XMLConfig.XML.ContentTypeArray.ContentTypeInfo
	Write-Output "`n--- Checking Content Types Association ---"
	if ($items.SubSiteURLCode -ne "_NOT_APPLICABLE_")
	{
		$targetURL = $RMWSiteURL + $SubSiteURLCode
	}
	else
	{
		$targetURL = $RMWSiteURL
	}
	
	$rmwWeb = Get-SPWeb ($targetURL)
	$targetWeb = $rmwWeb
	foreach ($item in $items)
	{		
		$targetList = $targetWeb.Lists[$item.List]
		$contentTypes = $targetList.ContentTypes.Name
		$row = $ContentTypeTable.NewRow()
		$row["Content Type"] = $item.ContentTypeName
		$row["List Name"] = $item.List
		$row["Expected Value"] = "Exists"
		$row["Actual Value"] = "Does not exist"
		$row["Pass or Fail"] = "Fail"
		$row["Date"] = $Date

		if ($contentTypes -contains $item.ContentTypeName)
		{
			$row["Actual Value"] = "Exists"
			$row["Pass or Fail"] = "PASS"
		}
		
		$ContentTypeTable.Rows.Add($row)	
	}
	
	$rmwWeb.dispose()
}

function RMWCheckSharePointSolutions
{
	Write-Output "`n--- Checking Sharepoint Solutions ---"
	$solutions = $XMLConfig.XML.SharePointSolutionsArray.SPSolutionInfo
	$solutionsInFarm = Get-SPSolution
	foreach ($solution in $solutions)
	{
		$row = $SPSolutionsTable.NewRow()
		$row["Solution Name"] = $solution.WSPFile
		$row["Status"] = "Not Deployed"
		$row["Status Required"] = "Deployed"
		$row["Pass or Fail"] = "FAIL"
		$row["Date"] = $Date
		if ($solutionsInFarm.Name -contains $solution.WSPFile)
		{
			$farmsolution = Get-SPSolution -Identity $solution.WSPFile
			$deloyStatus = $farmsolution.Deployed
			if ($deloyStatus -eq $solution.Deployed)
			{
				$row["Status"] = "Deployed"
				$row["Pass or Fail"] = "PASS"
			}
		}
		$SPSolutionsTable.Rows.Add($row)
	}
	
}

function RMWCheckHelpCenter([string]$WAOrRC, [string]$RMWSiteURL) #Shovan
{
	#RMWCheckHelpCenter: Check RMW Help Centers - Shovan

	$helpCenterSettings = $XMLConfig.XML.HelpCenterSettingsArray.SettingInfo
		
		#If are entries in the XML
	if (($helpCenterSettings -ne $null) -or ($helpCenterSettings.Count -gt 0))
	{
		Write-Output "`n--- Checking HelpCenter ---"
		
		
		
		foreach ($settingToCheck in $helpCenterSettings)
		{
			$SubSiteURLCode = $settingToCheck.SubSiteURLCode
			$settingToCheckListName = $settingToCheck.ListName
			$settingToCheckSettingType = $settingToCheck.SettingType
			$settingToCheckSettingValue = $settingToCheck.SettingValue
			$settingToCheckPassFail = "FAIL"
			$settingToCheckValueFound = ""
			
			if ($SubSiteURLCode -ne "_NOT_APPLICABLE_")
			{
				$targetURL = $RMWSiteURL + $SubSiteURLCode
			}
			else
			{
				$targetURL = $RMWSiteURL
			}
			
			$targetWeb = Get-SPWeb($targetURL)
			$targetList = $targetWeb.lists[$settingToCheckListName]
			
			#Check for error names
			if ($settingToCheckSettingValue -eq "_ALL_")
			{
				foreach ($item in $targetList.items)
				{
					$LinkLocation = $item["Link Location"]
					$url, $urlDescription = $LinkLocation.split(',')
					
					
					try
					{
						$webReq = Invoke-WebRequest -URI $url -UseDefaultCredentials
						$statusCode = $webReq.StatusCode
					}
					catch
					{
						$statusCode = $_.Exception.Response.StatusCode.Value__
					}
					
					
					if ($statusCode -eq 200)
					{
						$settingToCheckValueFound = "Yes"
						$settingToCheckPassFail = "PASS"
						
					}
					else 
					{
						$settingToCheckValueFound = "No"
						$settingToCheckPassFail = "FAIL"
					}
					
					$HelpCenterSettingsRow = $HelpCenterTable.NewRow()
					$HelpCenterSettingsRow["Tile Title"] = $item.name
					$HelpCenterSettingsRow["Value"] = $settingToCheckValueFound
					$HelpCenterSettingsRow["Value Required"] = "Is this item accessible?"
					$HelpCenterSettingsRow["Pass or Fail"] = $settingToCheckPassFail
					$HelpCenterSettingsRow["Date"] = $Date 
					$HelpCenterTable.Rows.Add($HelpCenterSettingsRow)
				}
			}
			else
			{
				$item = $targetList.Items | ? {$_.Title -eq $settingToCheckSettingValue}
				$LinkLocation = $item["Link Location"]
				$url, $urlDescription = $LinkLocation.split(',')
				
				
				try
				{
					$webReq = Invoke-WebRequest -URI $url -UseDefaultCredentials
					$statusCode = $webReq.StatusCode
				}
				catch
				{
					$statusCode = $_.Exception.Response.StatusCode.Value__
				}
				
				
				if ($statusCode -eq 200)
				{
					$settingToCheckValueFound = "Yes"
					$settingToCheckPassFail = "PASS"
				}
				else 
				{
					$settingToCheckValueFound = "No"
					$settingToCheckPassFail = "FAIL"
				}
				
				$HelpCenterSettingsRow = $HelpCenterTable.NewRow()
				$HelpCenterSettingsRow["Tile Title"] = $item.name
				$HelpCenterSettingsRow["Value"] = $settingToCheckValueFound
				$HelpCenterSettingsRow["Value Required"] = "Is this item accessible?"
				$HelpCenterSettingsRow["Pass or Fail"] = $settingToCheckPassFail
				$HelpCenterSettingsRow["Date"] = $Date 
				$HelpCenterTable.Rows.Add($HelpCenterSettingsRow)
			}
		}
	}
	else
	{
		Write-Output "`n--- No SharePoint Help Center Settings to check ---"
	}
		
	if ($targetWeb -ne $null)
	{
		$targetWeb.dispose()
	}

}

function RMWCustomButtonCheck ([string]$WAOrRC, [string]$RMWSiteURL) #Sharhad 
{
	$buttonSettings = $XMLConfig.XML.CustomRibbonButtonArray.ButtonInfo
	if ($items.SubSiteURLCode -ne "_NOT_APPLICABLE_")
	{
		$targetURL = $RMWSiteURL + $SubSiteURLCode
	}
	else
	{
		$targetURL = $RMWSiteURL
	}
	Write-Output "`n--- Checking Custom Buttons ---"
	$targetWeb = Get-SPWeb ($RMWSiteURL + "Migration/")
	foreach ($button in $buttonSettings)
	{
		$mappingLibrary = $targetWeb.Lists[$button.LibraryName]
		$buttonXML = $mappingLibrary.UserCustomActions.CommandUIExtension
		
		$CustomButtonRow = $CustomButtonTable.NewRow()
		$CustomButtonRow["Library Name"] = $button.LibraryName
		$CustomButtonRow["Button Name"] = $button.ButtonName
		$CustomButtonRow["Value"] = "Inactive"
		$CustomButtonRow["Value Required"] = "Active"
		$CustomButtonRow["Pass or Fail"] = "FAIL"
		$CustomButtonRow["Date"] = $Date
		
		if ($buttonXML.Contains($button.ButtonName))
		{
			$CustomButtonRow["Value"] = "Active"
			$CustomButtonRow["Pass or Fail"] = "PASS"
		}		 
		$CustomButtonTable.Rows.Add($CustomButtonRow)
	}
}

function RMWVerification
{
#RMWVerification: Main RMW verification function that calls all the verification sub functions

	$Date = Get-Date -format "dd-MMM-yyyy"
	
	#Type of server being checked -------------------------------------------------------------------------------------------
	$result = ""
	$waitingForInput = $true
	While ($waitingForInput)
	{
		Write-Output "`n`n--------------------------Environment Type Options------------------"		
		Write-Output "`n"
		Write-Output "            1.  Production or Montrium Connect environment"
		Write-Output "            2.  Test or Montrium Connect-Training environment"
		Write-Output "            3.  QA environment"
		Write-Output "            4.  DEV environment"
		Write-Output "            5.  Quit"


		$result = Read-Host "`nPlease select an option"
		
		switch ($result)
		{
			1 {$TestOrProd = "Prod"; $waitingForInput = $false}
			2 {$TestOrProd = "Test"; $waitingForInput = $false}
			3 {$TestOrProd = "QA"; $waitingForInput = $false}
			4 {$TestOrProd = "DEV"; $waitingForInput = $false}
			5 {stop-transcript; Exit}
		}
	}
	
	Write-Output "`nYou chose a $TestOrProd type environment"	
	
	##add to table
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "SharePoint Environment Type:"
	$GeneralInfoRow["Configuration Value"] = $TestOrProd
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	#Site Collection URL
	$RMWURL = $TargetSiteCollectionURL
	
	Write-Output "`nYou chose a URL of $rMWURL"
	
	#Check RMW RC URL
	$waitingForRealSite = $true
	#Update URL now ends in / so need to remove that first
	#$RMWRCURL = $RMWURL + "-RC"
	$RMWRCURL = $RMWURL.Substring(0,$RMWURL.Length-1) + "-RC"
	do 
	{
		#test to see if the site exists
		$testURL = get-spsite $RMWRCURL -ErrorAction SilentlyContinue
		
		#if it is equal to null then the site doesn't exist
		if ($testURL -eq $null)
		{
			Write-Output "`nThe URL ($RMWRCURL) does not correspond to a valid site collection URL" 
			$RMWRCURL = Read-Host "`nPlease enter the site collection URL where RMW Records Center is located"
		}
		#otherwise it does exist and you don't have to keep asking
		else
		{
			$testURL.dispose()
			$waitingForRealSite = $false
		}
	} while ($waitingForRealSite)
	
	#Add to General Info Table
	$GeneralInfoRow = $GeneralInfoTable.NewRow()
	$GeneralInfoRow["Info"] = "Target RC Site Collection URL:"
	$GeneralInfoRow["Configuration Value"] = $RMWRCURL
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	#We are assuming that the $RMWRCURL ends with /, double check this and ensure it does
	#Check the last character of the site collection URL string
	$lastChar = $RMWRCURL.substring($RMWRCURL.length - 1, 1)
		
	#Check if it is a /, if not, then need to add it in or it will mess things up
	if ($lastChar -ne "/")
	{
		$RMWRCURL = $RMWRCURL + "/"
	}
	
	$product = Read-Host "`nPlease enter the name of the Product in RegDocs Connect"
	
	$study = Read-Host "`nPlease enter the name of the Study in eTMF Connect"
	
	#Call the various functions that can only apply to one location, either Work Area or Central Admin depending
	GeneratePrerequisitesTable
	#RMWCheckSharePointSolutions
	#RMWCheckSecureStore
	GenerateDeploymentTable
	
	#Call the functions that can apply to either the Work Area or RC for the Work Area
	$WAOrRC = "WA"
	RMWCheckSiteSettings $WAOrRC $RMWURL
	RMWCheckLibrarySettings $WAOrRC $RMWURL
	RMWDataConnection $WAOrRC $RMWURL 
	RMWCheckInfoPathForms $WAOrRC $RMWURL
	#RMWCheckWorkflowConstants $WAOrRC $RMWURL
	RMWCheckSharePointFeatures $WAOrRC $RMWURL
	RMWCheckNintexWorkflows $WAOrRC $RMWURL 
	RMWCheckInfoPathFormServices
	RMWCheckCoSignSettings $WAOrRC $RMWURL 
	RMWContentTypeAssociation $WAOrRC $RMWURL
	RMWCheckHelpCenter $WAOrRC $RMWURL
	RMWCustomButtonCheck $WAOrRC $RMWURL	
	
	#Call the functions that can apply to either the Work Area or RC for the RC
	$WAOrRC = "RC"
	RMWCheckSiteSettings $WAOrRC $RMWRCURL 
	RMWCheckLibrarySettings $WAOrRC $RMWRCURL 
	RMWCheckSharePointFeatures $WAOrRC $RMWRCURL
	
	Write-Output "`n`n RMW Verification Complete!"
}

Function CreateHTMLFile([string]$Path,[string[]]$HeadData,[string]$BodyData)
# CreateHTMLFile: Functions to output HTML file as record of deployment
# Parameters:
# 	-> Path: Where to save the HTML File
#	-> HeadData: Data for the Head of the HTML File
#   -> BodyData: Data for the Body of the HTML File
{
	$head = $HeadData
	
	$body = $BodyData
		
	$null | ConvertTo-HTML -head $head -body $body | Set-Content $Path
}

Function GenerateReport
# GenerateReport: The main function (choice number 2) which Generates the report as a proof of what the script did
{
	$ComputerName = [system.environment]::MachineName
	
	Write-Host "`n--------------------Generating Report------------------"
	$IQCode = $XMLConfig.XML.IQCode
	#HTML Formatting
	##################################################
	#General Information Fragment formatting
	$GeneralInfoFragment = $GeneralInfoTable | ConvertTo-HTML "Info","Configuration Value","Type","Is Current Server","Date" -fragment
	$GeneralInfoHTML = "<br> General Information about the SharePoint servers:<br>$GeneralInfoFragment<br><hr>"
	
	#Prerequisites Table Fragment formatting
	$PrerequisitesFragment = $PrerequisitesTable | ConvertTo-HTML "Prerequisite","Value","Date" -fragment
	$PrerequisitesHTML = "<br> Prerequisites:<br>$PrerequisitesFragment<br><hr>"
	
	#Prerequisites Table Fragment formatting
	#$OtherPrerequisitesFragment = $OtherPrerequisitesTable | ConvertTo-HTML "Prerequisite","Value","Date" -fragment
	#$OtherPrerequisitesHTML = "<br> Other Prerequisites:<br>$OtherPrerequisitesFragment<br><hr>"
	
	#TODO: SharePoint Features Fragment formatting
	$SharepointFeaturesFragment = $CheckSharepointFeaturesTable | ConvertTo-HTML "WA or RC","Feature Name","Value","Value Required","Pass or Fail","Date" -fragment
	$SharepointFeaturesHTML =  "<br> Check Sharepoint Features Information:<br>$SharepointFeaturesFragment<br><hr>"
	
	#Data Connections Fragment Formatting
	$DataConnectionsFragment = $dataConnectionTable | ConvertTo-HTML "Data Connection","In List?","Expected Value","Actual Value","Pass or Fail","Date" -fragment
	$DataConnectionsHTML = "<br> Data Connections:<br>$DataConnectionsFragment<br><hr>"
	
	#InfoPath Forms Services Settings Fragment Formatting
	$InfoPathFormServicesSettingsFragment = $InfoPathFormServicesTable | ConvertTo-HTML "Property","Value","Value Required","Pass or Fail","Date" -fragment
	$InfoPathFormServicesSettingsHTML = "<br> InfoPath Form Services Settings:<br>$InfoPathFormServicesSettingsFragment<br><hr>"

	#InfoPath Forms Templates Fragment Formatting
	$InfoPathFormTemplatesFragment = $InfoPathFormTemplatesTable | ConvertTo-HTML "Form Name","Uploaded","Pass or Fail","Date" -fragment
	$InfoPathFormTemplatesHTML = "<br> InfoPath Form Templates:<br>$InfoPathFormTemplatesFragment<br><hr>"	
	
	#Site Settings Fragment formatting
	$SiteSettingsFragment = $SiteSettingsTable | ConvertTo-HTML "WA or RC","Setting Type","Setting Name","Value","Value Required","Pass or Fail","Date" -fragment
	$SiteSettingsHTML = "<br> Site Setting Information:<br>$SiteSettingsFragment<br><hr>"
	
	#Library Settings Fragment formatting
	$LibrarySettingsFragment = $LibrarySettingsTable | ConvertTo-HTML "WA or RC","Library Name","Verification Step","Value","Value Required","Pass or Fail","Date" -fragment
	$LibrarySettingsHTML = "<br> Library Setting Information:<br>$LibrarySettingsFragment<br><hr>"
	
	#Workflow Constants Fragment Formatting
	$WorkflowConstantFragment = $WorkflowConstantTable | ConvertTo-HTML "Workflow Constant Name","Value","Value Required","Pass or Fail","Date" -fragment
	$WorkflowConstantHTML = "<br> Workflow Constants:<br>$WorkflowConstantFragment<br><hr>"
	
	#Deployment Logs Fragment formatting
	$DeploymentFragment = $DeploymentTable | ConvertTo-HTML "Parameters","Value","Value Required","Pass or Fail","Date" -fragment
	$DeploymentHTML = "<br> Deployment Logs:<br>$DeploymentFragment<br><hr>"
	
	#Deployment Log Files Fragment formatting
	$DeploymentLogFilesFragment = $DeploymentLogFilesTable | ConvertTo-HTML "Manual Steps","Result","Date" -fragment
	$DeploymentLogFilesFragmentHTML = "<br> Deployment Log Files:<br>$DeploymentLogFilesFragment<br><hr>"
	
	#Record Center Settings Fragment Formatting
	$RecordCenterSettingsFragment = $RecordCenterSettingsTable | ConvertTo-HTML "Verification Steps","Value","Value Required","Pass or Fail","Date" -fragment
	$RecordCenterSettingsHTML = "<br> Record Center Settings:<br>$RecordCenterSettingsFragement<br><hr>"
	
	#Solution Fragment Formatting
	$SPSolutionsFragment = $SPSolutionsTable | ConvertTo-HTML "Solution Name","Status","Status Required","Pass or Fail","Date" -fragment
	$SPSolutionsTableHTML = "<br> Solutions:<br>$SPSolutionsFragment<br><hr>"
	
	#Workflows Fragment Formatting
	$WorkflowsFragment = $WorkflowTable | ConvertTo-HTML "Workflow Name", "Library","Value","Value Required","Pass or Fail","Date" -fragment
	$WorkflowTableHTML = "<br> Workflows:<br>$WorkflowsFragment<br><hr>"
	
	
	#CoSign Settings Fragment Formatting
	$CoSignSettingsFragment = $CoSignSettingsTable | ConvertTo-HTML "CoSign Setting","Setting Location","Expected Value","Actual Value","Pass or Fail","Date" -fragment
	$CoSigSettingsHTML = "<br> CoSign Features:<br>$CoSignSettingsFragment<br><hr>"
	
	#Content Type Fragment Formatting
	$ContentTypeFragment = $ContentTypeTable | ConvertTo-HTML "Content Type","List Name","Expected Value","Actual Value","Pass or Fail","Date" -fragment
	$ContentTypeHTML = "<br>Content Type Association:<br>$ContentTypeFragment<br><hr>"
	
	#HelpCenter Fragment Formatting
	$HelpCenterFragment = $HelpCenterTable | ConvertTo-HTML "Tile Title","Value","Value Required","Pass or Fail","Date" -fragment
	$HelpCenterHTML = "<br> Help Center:<br>$HelpCenterFragment<br><hr>"
	
	#CustomButton Fragment Formatting
	$CustomButtonFragment = $CustomButtonTable | ConvertTo-HTML "Library Name","Button Name","Value","Value Required","Pass or Fail","Date" -fragment
	$CustomButtonHTML = "<br> Custom Buttons:<br>$CustomButtonFragment<br><hr>"
	
	#Transcript Fragment formatting
	$TranscriptContent = get-content -Path $TranscriptPath
	$TranscriptFragment = ""
	foreach ($line in $TranscriptContent)
	{
		$TranscriptFragment = $TranscriptFragment + "<p>$line</p>"
	}
	$TranscriptHTML = "<br> Transcript of script:<br>$TranscriptFragment<br><hr>"
		
	###Create HTML File
	###################################################
	$Date = "{0:yyyy-MM-dd HH.mm.ss.fff}" -f (get-date)
	$Currentdir = [string](Get-location) 
	$FileName = "$ComputerName $ProductCodeFromXML $ProductVersionFromXML $ScriptType $Date.html"
	$Path =  "$Currentdir\$Filename"
	#note when putting an array on multiple lines, the final closing one ("@) cannot be indented at all, must be right at the beginning of the line
	$headTag = @"
	<style>
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #CAE8EA;}
	TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
	</style>
	<title>
	$Title
	</title>
"@
	
	$BodyStart = "<h1 style=""color:#9fb11e;font-size:30px"">$ProductCodeFromXML $ProductVersionFromXML $ScriptType</h1><br><h3 style=""color:#9fb11e;margin-left:30px;"">IQ Script for $IQCode</h3><hr>"
	
	#TODO: update the below to handle all the tables from all the sections
	
	#$BodyTag = "$Bodystart $GeneralInfoHTML $SiteSettingsHTML $CoSignHTML $OfficeWebAppsHTML $NintexHTML $ThirdPartyHTML $MontriumWSPHTML $TranscriptHTML"
	#$BodyTag = "$Bodystart $GeneralInfoHTML $PrerequisitesHTML $OtherPrerequisitesHTML $SiteSettingsHTML $LibrarySettingsHTML $DeploymentHTML $DeploymentLogFilesFragmentHTML $RecordCenterSettingsHTML $TranscriptHTML"
	#$BodyTag = "$Bodystart $GeneralInfoHTML" #this one worked
	$BodyTag = "$Bodystart $GeneralInfoHTML $PrerequisitesHTML $DeploymentHTML $DeploymentLogFilesFragmentHTML $SharepointFeaturesHTML $DataConnectionsHTML $InfoPathFormServicesSettingsHTML $InfoPathFormTemplatesHTML $SiteSettingsHTML $LibrarySettingsHTML $SPSolutionsTableHTML $WorkflowConstantHTML $RecordCenterSettingsHTML $WorkflowTableHTML $CoSignFeaturesHTML $CoSigSettingsHTML $ContentTypeHTML $AttachmentSeperatorHTML $HelpCenterHTML $CustomButtonHTML $TranscriptHTML"
	
	$BodyTagPass =  $BodyTag | foreach {if ($_ -match "PASS") {$_ -replace "PASS", "<font color=green><b>PASS</b></font>"}}
	$BodyTagPassFail =  $BodyTagPass | foreach {if ($_ -match "FAIL") {$_ -replace "FAIL", "<font color=red><b>FAIL</b></font>"}}
	#$BodyTagPassFailNewLine =  $BodyTagPassFail | foreach {if ($_ -match "_New_Line_") {$_ -replace "_New_Line_", "<br>"}}
	
	#Write-Output "Path:"
	#$Path
	
	#Write-Output "headTag:"
	#$headTag
	
	#Write-Output "BodyStart:"
	#$BodyStart
	
	#Write-Output "BodyTag:"
	#$BodyTag
	
	#Write-Output "BodyTagPassFail:"
	#$BodyTagPassFail
	
	#Write-Output "BodyTagPassFailNewLine:"
	#$BodyTagPassFailNewLine	
	
	#CreateHTMLFile $Path $headTag $BodyTagPassFailNewLine
	CreateHTMLFile $Path $headTag $BodyTagPassFail
	
	#CreateHTMLFile $Path $headTag $BodyTag

#TODO: figure out why body tag passfail is empty, same for new line.	
	
	
	#Can stay write host, not part of the transcript
	Write-Host "`n`n Report Generated!" -fore green
	
	UploadReport $FileName $Path
}

Function UploadReport ([string]$ReportName, [string]$ReportPath)
#UploadReport: Uploads the report to a library named Montrium Logs in the landing site
{
	Write-Host "`n--------------------Uploading Report------------------"
	
	# Root Lists and folder collection
	$spSite = get-spsite $AdminLandingSiteCollectionURL
	$rootweb = $spSite.rootweb
	$EntityList = $rootweb.Lists[$destinationLibraryName]
	$spFolder = $rootweb.GetFolder($destinationLibraryName)
	$spFileCollection = $spFolder.Files
	
	$currentFile = get-childitem $ReportPath
	
	$destPath = $destinationLibraryName + "/" + $ReportName
	
	# Uploads the file, and gets a SPFile obj
	[Microsoft.SharePoint.SPFile] $fileRef = $spFileCollection.Add($destPath, $currentFile.OpenRead(),$true)
	
	#update metadata
	[Microsoft.SharePoint.SPListItem] $listItem = $fileRef.Item
	
	#Product Metadata - single line of text
	$listItem["Product"] = $ProductNameFromXML
	
	#Product Version Metadata - single line of text
	$listItem["Product Version"] = $ProductVersionFromXML
	
	#Script Version Metadata - single line of text
	$listItem["Script Version"] = $ScriptVersion
	
	#Web App - Single line of Text
	if ($AppliesToWebApps -eq "ALL")
	{
		$WebAppString = ""
		$AllContentWebApps = get-spwebapplication
		foreach ($curentWebApp in $AllContentWebApps)
		{
			$CurrentWebAppDisplayName = $curentWebApp.DisplayName
			if ($WebAppString -eq "")
			{
				$WebAppString = $CurrentWebAppDisplayName
			}
			else
			{
				$WebAppString = $WebAppString + ", " + $CurrentWebAppDisplayName
			}
		}
		$listItem["Web App"] = $WebAppString
	}
	else
	{
		$targetSite = get-spsite $TargetSiteCollectionURL
		$curentWebApp = $targetSite.WebApplication
		$CurrentWebAppDisplayName = $curentWebApp.DisplayName
		$listItem["Web App"] = $CurrentWebAppDisplayName
		$targetSite.dispose()
	}
	
	#Entity Code - Single line of text
	if ($AppliesToEntities -eq "ALL")
	{
		$listItem["Entity Code"] = "All Entities"
	}
	else
	{
		#has to be target site collection URL here
		$siteURLArray = $TargetSiteCollectionURL.split("/")
		$entityCode = $siteURLArray[4]
		$listItem["Entity Code"] = $entityCode
	}
	
	#Product Metadata - type person or group - people only
	$user = $rootweb.EnsureUser($FullLoginName)
	if ($user -eq $Null)
	{
		#Not part of transcript can stay write-host
		Write-Host "Could not find user named $FullLoginName" -fore yellow
	}
	else 
	{
		$listItem["Script Initiator"] = $user
	}
	
	#set to $false then it will not change the version number
	$listItem.SystemUpdate($false)
	
	$IDOfHTML = $listItem.ID
	
	#Get the file name with and without extension and get the extension
	$SPFileName = $fileRef.Name
	$IndexOfLastDot = $SPFileName.LastIndexOf(".")
	$IndexOfExtension = $IndexOfLastDot + 1
	$LenghtOfExtension = $SPFileName.Length - $IndexOfExtension
	
	$SPFileExtension = $SPFileName.substring($IndexOfExtension, $LenghtOfExtension)
	$SPFileNameNoExtension = $SPFileName.substring(0, $IndexOfLastDot)
	
	#convert the uploaded file to PDF - see above AdminLandingSiteCollectionURL will end with /
	$URI = $AdminLandingSiteCollectionURL + "_vti_bin/MTM_methods.asmx?WSDL"  ##URL of workflow.asmx on the subsite you are working on, need the "?WSDL" at the end or it doesn't work
	
	$proxy = New-WebServiceProxy -Uri $URI -UseDefaultCredential
	
	$libraryURL = $AdminLandingSiteCollectionURL + $destinationLibraryName
	
	$inFileUrl = $libraryURL + "/" + $SPFileName
	$outFileUrl = $libraryURL + "/" + $SPFileNameNoExtension + ".pdf"
	$xslFileUrl = ""
	$xmlJobTemplateUri = ""
	$contentTypeName = "Document"
	$overwrite = "false"
	$useAdlib = "NO"
	
	$IDOfNewPDF = $proxy.ConvertFile($inFileUrl, $outFileUrl, $xslFileUrl, $xmlJobTemplateUri, $contentTypeName, $overwrite, $useAdlib)
	
	$intIDOfHTML = [int]$IDOfHTML
	$intIDOfNewPDF = [int]$IDOfNewPDF
	
    #Can stay write host, not part of the transcript
	if (($intIDOfHTML + 1) -eq $intIDOfNewPDF )
	{
		Write-Host "`n`n Report Uploaded and converted to PDF!" -fore green
	}
	else
	{
		Write-Host "`n`n Report Uploaded but some problem with conversion to pdf:" -fore yellow
		$IDOfNewPDF
	}
	
	$rootweb.dispose()
	$spSite.dispose()
}


################################ MAIN ##########################################



##Build tables:
	#Build General Information table
	$GeneralInfoTable = New-Object system.Data.DataTable "General Information"
	$BlankCol = New-Object system.Data.DataColumn Blank,([string])
	$BlankCol.ColumnName = "Info"
	$ConfigCol = New-Object system.Data.DataColumn Config,([string])
	$ConfigCol.ColumnName = "Configuration Value"
	$TypeCol = New-Object system.Data.DataColumn Type,([string])
	$TypeCol.ColumnName = "Type"
	$IsCurrentCol = New-Object system.Data.DataColumn IsCurrent,([string])
	$IsCurrentCol.ColumnName = "Is Current Server"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$GeneralInfoTable.columns.add($BlankCol)
	$GeneralInfoTable.columns.add($ConfigCol)
	$GeneralInfoTable.columns.add($TypeCol)
	$GeneralInfoTable.columns.add($IsCurrentCol)
	$GeneralInfoTable.columns.add($DateCol)

	#Build Prerequisites table
	$PrerequisitesTable = New-Object system.Data.DataTable "Prerequisites"
	$PrerequisiteCol = New-Object system.Data.DataColumn Blank,([string])
	$PrerequisiteCol.ColumnName = "Prerequisite"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$PrerequisitesTable.columns.add($PrerequisiteCol)
	$PrerequisitesTable.columns.add($ValueCol)
	$PrerequisitesTable.columns.add($DateCol)

	#Build Other Prerequisites table
	$OtherPrerequisitesTable = New-Object system.Data.DataTable "Other Prerequisites"
	$PrerequisiteCol = New-Object system.Data.DataColumn Blank,([string])
	$PrerequisiteCol.ColumnName = "Prerequisite"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$OtherPrerequisitesTable.columns.add($PrerequisiteCol)
	$OtherPrerequisitesTable.columns.add($ValueCol)
	$OtherPrerequisitesTable.columns.add($DateCol)
	
	#Build Sharepoint Features Table
	$CheckSharepointFeaturesTable = New-Object system.Data.DataTable "Check Sharepoint Features"
	$SiteCollectionCol = New-Object system.Data.DataColumn SiteCollection,([string])
	$SiteCollectionCol.ColumnName = "WA or RC"
	$FeatureNameCol = New-Object system.Data.DataColumn Property,([string])
	$FeatureNameCol.ColumnName = "Feature Name"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$CheckSharepointFeaturesTable.columns.add($SiteCollectionCol)
	$CheckSharepointFeaturesTable.columns.add($FeatureNameCol)
	$CheckSharepointFeaturesTable.columns.add($ValueCol)
	$CheckSharepointFeaturesTable.columns.add($ValueRequiredCol)
	$CheckSharepointFeaturesTable.columns.add($PassFailCol)
	$CheckSharepointFeaturesTable.columns.add($DateCol)

	
	#Data Connections Table
	$DataConnectionTable = New-Object system.Data.DataTable "Data Connections"
	$dataConnectionCol = New-Object system.Data.DataColumn colName_1,([string])
	$dataConnectionCol.ColumnName = "Data Connection"
	$inListCol = New-Object system.Data.DataColumn colName_1,([string])
	$inListCol.ColumnName = "In List?"
	$expectedApprovalCol = New-Object system.Data.DataColumn colName_3,([string])
	$expectedApprovalCol.ColumnName = "Expected Value"
	$actualApprovalCol = New-Object system.Data.DataColumn colName_4,([string])
	$actualApprovalCol.ColumnName = "Actual Value"
	$passFailCol = New-Object system.Data.DataColumn colName_5,([string])
	$passFailCol.ColumnName = "Pass or Fail" 
	$dateCol = New-Object system.Data.DataColumn colName_6,([string])
	$dateCol.ColumnName = "Date"
	$DataConnectionTable.columns.add($dataConnectionCol)
	$DataConnectionTable.columns.add($inListCol)
	$DataConnectionTable.columns.add($expectedApprovalCol)
	$DataConnectionTable.columns.add($actualApprovalCol)
	$DataConnectionTable.columns.add($passFailCol)
	$DataConnectionTable.columns.add($dateCol)
	
	#Build InfoPath Form Services table
	$InfoPathFormServicesTable = New-Object system.Data.DataTable "InfoPath Form Services"
	$PropertyCol = New-Object system.Data.DataColumn Property,([string])
	$PropertyCol.ColumnName = "Property"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$InfoPathFormServicesTable.columns.add($PropertyCol)
	$InfoPathFormServicesTable.columns.add($ValueCol)
	$InfoPathFormServicesTable.columns.add($ValueRequiredCol)
	$InfoPathFormServicesTable.columns.add($PassFailCol)
	$InfoPathFormServicesTable.columns.add($DateCol)

	#Build InfoPath Form Templates table
	$InfoPathFormTemplatesTable = New-Object system.Data.DataTable "InfoPath Form Templates"
	$FormNameCol = New-Object system.Data.DataColumn FormName,([string])
	$FormNameCol.ColumnName = "Form Name"
	$UploadedCol = New-Object system.Data.DataColumn Uploaded,([string])
	$UploadedCol.ColumnName = "Uploaded"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$InfoPathFormTemplatesTable.columns.add($FormNameCol)
	$InfoPathFormTemplatesTable.columns.add($UploadedCol)
	$InfoPathFormTemplatesTable.columns.add($PassFailCol)
	$InfoPathFormTemplatesTable.columns.add($DateCol)

	#Build Site Settings table
	$SiteSettingsTable = New-Object system.Data.DataTable "Site Settings"
	$SiteCollectionCol = New-Object system.Data.DataColumn SiteCollection,([string])
	$SiteCollectionCol.ColumnName = "WA or RC"
	$SettingTypeCol = New-Object system.Data.DataColumn SettingType,([string])
	$SettingTypeCol.ColumnName = "Setting Type"
	$SettingNameCol = New-Object system.Data.DataColumn SettingName,([string])
	$SettingNameCol.ColumnName = "Setting Name"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$SiteSettingsTable.columns.add($SiteCollectionCol)
	$SiteSettingsTable.columns.add($SettingTypeCol)
	$SiteSettingsTable.columns.add($SettingNameCol)
	$SiteSettingsTable.columns.add($ValueCol)
	$SiteSettingsTable.columns.add($ValueRequiredCol)
	$SiteSettingsTable.columns.add($PassFailCol)
	$SiteSettingsTable.columns.add($DateCol)

	#Build Library Settings table
	$LibrarySettingsTable = New-Object system.Data.DataTable "Library Settings"
	$SiteCollectionCol = New-Object system.Data.DataColumn SiteCollection,([string])
	$SiteCollectionCol.ColumnName = "WA or RC"
	$LibraryNameCol = New-Object system.Data.DataColumn LibraryName,([string])
	$LibraryNameCol.ColumnName = "Library Name"
	$SettingNameCol = New-Object system.Data.DataColumn SettingName,([string])
	$SettingNameCol.ColumnName = "Verification Step"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$LibrarySettingsTable.columns.add($SiteCollectionCol)
	$LibrarySettingsTable.columns.add($LibraryNameCol)
	$LibrarySettingsTable.columns.add($SettingNameCol)
	$LibrarySettingsTable.columns.add($ValueCol)
	$LibrarySettingsTable.columns.add($ValueRequiredCol)
	$LibrarySettingsTable.columns.add($PassFailCol)
	$LibrarySettingsTable.columns.add($DateCol)

	#Build Solutions table
	$SPSolutionsTable = New-Object system.Data.DataTable "Sharepoint Solutions Table"
	$SolutionNameCol = New-Object system.Data.DataColumn WorkflowConstantName, ([string])
	$SolutionNameCol.ColumnName = "Solution Name"
	$StatusCol = New-Object system.Data.DataColumn Value, ([string])
	$StatusCol.ColumnName = "Status"
	$StatusRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$StatusRequiredCol.ColumnName = "Status Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	
	$SPSolutionsTable.columns.add($SolutionNameCol)
	$SPSolutionsTable.columns.add($StatusCol)
	$SPSolutionsTable.columns.add($StatusRequiredCol)
	$SPSolutionsTable.columns.add($PassFailCol)
	$SPSolutionsTable.columns.add($DateCol)
	
	#Build Workflow Constants table
	$WorkflowConstantTable = New-Object system.Data.DataTable "Workflow Constants Table"
	$WorkflowConstantNameCol = New-Object system.Data.DataColumn WorkflowConstantName, ([string])
	$WorkflowConstantNameCol.ColumnName = "Workflow Constant Name"
	$ValueCol = New-Object system.Data.DataColumn Value, ([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$WorkflowConstantTable.columns.add($WorkflowConstantNameCol)
	$WorkflowConstantTable.columns.add($ValueCol)
	$WorkflowConstantTable.columns.add($ValueRequiredCol)
	$WorkflowConstantTable.columns.add($PassFailCol)
	$WorkflowConstantTable.columns.add($DateCol)

	#Build Deployment Logs table
	$DeploymentTable = New-Object system.Data.DataTable "Deployment Logs"
	$ParameterCol = New-Object system.Data.DataColumn Parameter,([string])
	$ParameterCol.ColumnName = "Parameters"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$DeploymentTable.columns.add($ParameterCol)
	$DeploymentTable.columns.add($ValueCol)
	$DeploymentTable.columns.add($ValueRequiredCol)
	$DeploymentTable.columns.add($PassFailCol)
	$DeploymentTable.columns.add($DateCol)

	#Build Deployment Log Files table
	$DeploymentLogFilesTable = New-Object system.Data.DataTable "Deployment Log Files"
	$ManualStepsCol = New-Object system.Data.DataColumn ManualSteps,([string])
	$ManualStepsCol.ColumnName = "Manual Steps"
	$ResultsCol = New-Object system.Data.DataColumn Results,([string])
	$ResultsCol.ColumnName = "Result"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$DeploymentLogFilesTable.columns.add($ManualStepsCol)
	$DeploymentLogFilesTable.columns.add($ResultsCol)
	$DeploymentLogFilesTable.columns.add($DateCol)

	#Build Record Center Settings table
	$RecordCenterSettingsTable = New-Object system.Data.DataTable "Record Center Settings"
	$VerificationStepsCol = New-Object system.Data.DataColumn VerificationSteps,([string])
	$VerificationStepsCol.ColumnName = "Verification Steps"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$RecordCenterSettingsTable.columns.add($VerificationStepsCol)
	$RecordCenterSettingsTable.columns.add($ValueCol)
	$RecordCenterSettingsTable.columns.add($ValueRequiredCol)
	$RecordCenterSettingsTable.columns.add($PassFailCol)
	$RecordCenterSettingsTable.columns.add($DateCol)

	#Build Workflow table
	$WorkflowTable = New-Object system.Data.DataTable "Workflows"
	$WorkflowNameCol = New-Object system.Data.DataColumn WorkflowName,([string])
	$WorkflowNameCol.ColumnName = "Workflow Name"
	$LibraryCol = New-Object system.Data.DataColumn Library,([string])
	$LibraryCol.ColumnName = "Library"
	$ValueCol = New-Object system.Data.DataColumn Value,([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"

	$WorkflowTable.columns.add($WorkflowNameCol)
	$WorkflowTable.columns.add($LibraryCol)
	$WorkflowTable.columns.add($ValueCol)
	$WorkflowTable.columns.add($ValueRequiredCol)
	$WorkflowTable.columns.add($PassFailCol)
	$WorkflowTable.columns.add($DateCol)

	#Build CoSign Features Table
	$CoSignFeaturesTable = New-Object system.Data.DataTable "CoSign Features"
	$FeatureNameCol = New-Object system.Data.DataColumn colName_1,([string])
	$FeatureNameCol.ColumnName = "CoSign Feature"
	$ExpectedCol = New-Object system.Data.DataColumn colName_3,([string])
	$ExpectedCol.ColumnName = "Expected Value"
	$ValueCol = New-Object system.Data.DataColumn colName_4,([string])
	$ValueCol.ColumnName = "Actual Value"
	$PassFailCol = New-Object system.Data.DataColumn colName_5,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn colName_6,([string])
	$DateCol.ColumnName = "Date"

	$CoSignFeaturesTable.columns.add($FeatureNameCol)
	$CoSignFeaturesTable.columns.add($ExpectedCol)
	$CoSignFeaturesTable.columns.add($ValueCol)
	$CoSignFeaturesTable.columns.add($PassFailCol)
	$CoSignFeaturesTable.columns.add($DateCol)
	
	#Build CoSign Settings Table
	$CoSignSettingsTable = New-Object system.Data.DataTable "CoSign Settings"
	$CoSignSettingCol = New-Object system.Data.DataColumn colName_1,([string])
	$CoSignSettingCol.ColumnName = "CoSign Setting"
	$SettingLocCol = New-Object system.Data.DataColumn colName_1,([string])
	$SettingLocCol.ColumnName = "Setting Location"
	$ExpectedCol = New-Object system.Data.DataColumn colName_3,([string])
	$ExpectedCol.ColumnName = "Expected Value"
	$ValueCol = New-Object system.Data.DataColumn colName_4,([string])
	$ValueCol.ColumnName = "Actual Value"
	$PassFailCol = New-Object system.Data.DataColumn colName_5,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn colName_6,([string])
	$DateCol.ColumnName = "Date"

	$CoSignSettingsTable.columns.add($CoSignSettingCol)
	$CoSignSettingsTable.columns.add($SettingLocCol)
	$CoSignSettingsTable.columns.add($ExpectedCol)
	$CoSignSettingsTable.columns.add($ValueCol)
	$CoSignSettingsTable.columns.add($PassFailCol)
	$CoSignSettingsTable.columns.add($DateCol)
	
	#Build Content Type Association Table
	$ContentTypeTable = New-Object system.Data.DataTable "Content Type"
	#define Coloums
	$contentTypeCol = New-Object system.Data.DataColumn colName_1,([string])
	$contentTypeCol.ColumnName = "Content Type"
	$listCol = New-Object system.Data.DataColumn colName_1,([string])
	$listCol.ColumnName = "List Name"
	$expectedValueCol = New-Object system.Data.DataColumn colName_3,([string])
	$expectedValueCol.ColumnName = "Expected Value"
	$actualValueCol = New-Object system.Data.DataColumn colName_4,([string])
	$actualValueCol.ColumnName = "Actual Value"
	$passFailCol = New-Object system.Data.DataColumn colName_5,([string])
	$passFailCol.ColumnName = "Pass or Fail" 
	$dateCol = New-Object system.Data.DataColumn colName_6,([string])
	$dateCol.ColumnName = "Date"
	
	$ContentTypeTable.columns.add($contentTypeCol)
	$ContentTypeTable.columns.add($listCol)
	$ContentTypeTable.columns.add($expectedValueCol)
	$ContentTypeTable.columns.add($actualValueCol)
	$ContentTypeTable.columns.add($passFailCol)
	$ContentTypeTable.columns.add($dateCol)
	
	#build attachment seperator table
	$AttachmentSeperatorTable = New-Object system.Data.DataTable "Attachment Seperator"
	#define Coloums
	$fileNameCol = New-Object system.Data.DataColumn colName_1,([string])
	$fileNameCol.ColumnName = "File Name"
	$listCol = New-Object system.Data.DataColumn colName_1,([string])
	$listCol.ColumnName = "List Name"
	$expectedValueCol = New-Object system.Data.DataColumn colName_3,([string])
	$expectedValueCol.ColumnName = "Expected Value"
	$actualValueCol = New-Object system.Data.DataColumn colName_4,([string])
	$actualValueCol.ColumnName = "Actual Value"
	$passFailCol = New-Object system.Data.DataColumn colName_5,([string])
	$passFailCol.ColumnName = "Pass or Fail" 
	$dateCol = New-Object system.Data.DataColumn colName_6,([string])
	$dateCol.ColumnName = "Date"
	
	$AttachmentSeperatorTable.columns.add($fileNameCol)
	$AttachmentSeperatorTable.columns.add($listCol)
	$AttachmentSeperatorTable.columns.add($expectedValueCol)
	$AttachmentSeperatorTable.columns.add($actualValueCol)
	$AttachmentSeperatorTable.columns.add($passFailCol)
	$AttachmentSeperatorTable.columns.add($dateCol)
	
	#Build Help Center Table
	$HelpCenterTable = New-Object system.Data.DataTable "Help Center Table"
	$TileTitleCol = New-Object system.Data.DataColumn TileTitle, ([string])
	$TileTitleCol.ColumnName = "Tile Title"
	$ValueCol = New-Object system.Data.DataColumn Value, ([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn RequiredValue,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$HelpCenterTable.columns.add($TileTitleCol)
	$HelpCenterTable.columns.add($ValueCol)
	$HelpCenterTable.columns.add($ValueRequiredCol)
	$HelpCenterTable.columns.add($PassFailCol)
	$HelpCenterTable.columns.add($DateCol)

	#Build Custom Button Table
	$CustomButtonTable = New-Object system.Data.DataTable "Custom Button Table"
	$LibraryCol = New-Object system.Data.DataColumn Library, ([string])
	$LibraryCol.ColumnName = "Library Name"
	$ButtonCol = New-Object system.Data.DataColumn Button, ([string])
	$ButtonCol.ColumnName = "Button Name"
	$ValueCol = New-Object system.Data.DataColumn Value, ([string])
	$ValueCol.ColumnName = "Value"
	$ValueRequiredCol = New-Object system.Data.DataColumn RequiredValue,([string])
	$ValueRequiredCol.ColumnName = "Value Required"
	$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
	$PassFailCol.ColumnName = "Pass or Fail"
	$DateCol = New-Object system.Data.DataColumn Date,([string])
	$DateCol.ColumnName = "Date"
	$CustomButtonTable.columns.add($LibraryCol)
	$CustomButtonTable.columns.add($ButtonCol)
	$CustomButtonTable.columns.add($ValueCol)
	$CustomButtonTable.columns.add($ValueRequiredCol)
	$CustomButtonTable.columns.add($PassFailCol)
	$CustomButtonTable.columns.add($DateCol)


###########Get Info from XML and provide User Input Choices###########

$ScriptVersion = "1.0"

$ScriptType = "Testing Script"

#Where the logs will be stored
$destinationLibraryName = "Montrium Logs"

#Whether the script applies to "ALL" or "Single" web app. If single then it will apply to the web app of the site collection URL you are saying.
$AppliesToWebApps = "Single"

#Whether the script applies to "ALL" or "Single" URL/Client/Entity...
$AppliesToEntities = "Single"

#Whether it is an update type script or verification type script
$UpdateOrVerification = "Verification"

$TranscriptFileName = "RMW Verification Script"

############## END OF THINGS TO UPDATE

#Get the user running the script
$VerificationUserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

#Check if the user is local admin
$isLocalAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if ($isLocalAdmin)
{
	#Check if farm is not null and if not then check if is farm admin
	$farm = get-spfarm
	
	if ($farm -ne $null)
	{
		if($farm.CurrentUserIsAdministrator())
		{
			$CurrentDir = Get-Location

			Read-Host "Press enter to choose the accompanying XML file"
			$AccompanyingXML = Get-FileName -initialDirectory $Currentdir -fileType "XML"
			
			$ConfigPathExists = Test-Path $AccompanyingXML
			if ($ConfigPathExists -eq $True)
			{
				Write-Output "`nConfiguration file found at $AccompanyingXML"
				[xml]$XMLConfig = Get-Content $AccompanyingXML
				Write-Output "Getting content..."

				$ProductNameFromXML = $XMLConfig.XML.Product
				$ProductCodeFromXML = $XMLConfig.XML.ProductCode
				$ProductVersionFromXML =  $XMLConfig.XML.ProductVersion
				$ProductSharePointVersionFromXML =  $XMLConfig.XML.SharePointVersionForProduct
			}
			else
			{
				Write-Output "`nConfiguration file not found, aborting!"
				stop-transcript
				Exit
			}
			
			$Date = "{0:yyyy-MM-dd HH.mm.ss.fff}" -f (get-date)
			$TranscriptFileNameWithDateAndExtension = "$TranscriptFileName $Date.txt"
			$TranscriptPath =  "$Currentdir\$TranscriptFileNameWithDateAndExtension"

			start-transcript -path $TranscriptPath
			
			######################## Inform User of what script will do ############################
			Write-Output "`n`n-----------------------------------$ProductCodeFromXML $ScriptType Script---------------"

			Write-Output "`nThis Script will check the Foundation components, comparing them to an XML file which has the versions to look for."
			Write-Output "`nThis Script must be run as local admin, farm admin and site admin."
			Write-Output "You will be asked to perform the following steps:"
			Write-Output "`n"
			Write-Output "      1. Select the accompanying XML file."
			Write-Output "      2. Enter the site collection URL where the Admin landing site is located"
			Write-Output "      3. Enter your login name without domain."
			Write-Output "      4. Select whether CoSign should be checked or not."
			Write-Output "      5. Select whether Office Web Apps should be checked or not."
			Write-Output "      6. Select the environment type (DEV, QA, TEST or PROD)"
			
			Write-Output "`n"
			Write-Output "The script will then perform the following actions:"
			Write-Output "`n"
			Write-Output "      1.  Log the user's display and the account the script is being run under."
			Write-Output "      2.  Check each server that is part of the farm seeing what type it is."
			Write-Output "      3.  Look for CoSign Settings if applicable, as listed in the XML."
			Write-Output "      4.  Look for Office Web Apps Settings if applicable, as listed in the XML."
			Write-Output "      5.  Check the Nintex Settings, as listed in the XML."
			Write-Output "      6.  Open up Internet Explorer."
			Write-Output "      7.  Check the third party programs installed on the server listed in the XML."
			Write-Output "      8.  Open up Internet Explorer."
			Write-Output "      9.  Check the third party programs listed in Central admin listed in the XML."
			Write-Output "      10. Check Montrium's WSPs listed in the XML."
			Write-Output "      11. Close Internet Explorer."
			Write-Output "      12. Create HTML report based on the information gathered."
			Write-Output "      13. Upload HTML report to the $destinationLibraryName library in the landing site."
			Write-Output "      14. Convert the HTML report to PDF."
			
			Write-Output "`n`n"	

			#Need to get the site collection
			$waitingForRealSite = $true

			do 
			{
				$AdminLandingSiteCollectionURL = Read-Host "`nPlease enter the site collection URL where the Admin Landing site is located"
				
				#test to see if the site exists
				$testURL = get-spsite $AdminLandingSiteCollectionURL -ErrorAction SilentlyContinue
				
				#if it is equal to null then the site doesn't exist
				if ($testURL -eq $null)
				{
					Write-Output "`nThe URL you entered ($AdminLandingSiteCollectionURL) does not correspond to a valid site collection URL"
				}
				#if it isn't null but has -RC in it then it isn't an admin site
				elseif ($AdminLandingSiteCollectionURL -match "-RC")
				{
					Write-Output "`nThe URL you entered ($AdminLandingSiteCollectionURL) is a valid URL but matches -RC so is a Record Center"
				}
				#if it isn't null but has QMW in it then it isn't an admin site
				elseif ($AdminLandingSiteCollectionURL -match "QMW")
				{
					Write-Output "`nThe URL you entered ($AdminLandingSiteCollectionURL) is a valid URL but matches QMW so is a Quality Management site"
				}
				#if it isn't null but has RMW in it then it isn't an admin site
				elseif ($AdminLandingSiteCollectionURL -match "RMW")
				{
					Write-Output "`nThe URL you entered ($AdminLandingSiteCollectionURL) is a valid URL but matches RMW so is a Records Management"
				}
				#if it isn't null but has Home in it then it isn't an admin site
				elseif ($AdminLandingSiteCollectionURL -match "Home")
				{
					Write-Output "`nThe URL you entered ($AdminLandingSiteCollectionURL) is a valid URL but matches Home so is a Home Site"
				}
				#otherwise it does exist but now you have to check for a montrium logs library
				else
				{
					#check for Montrium logs library
					$testRootweb = $testURL.rootweb
					$testMontriumLogsList = $testRootweb.Lists["Montrium Logs"]
					
					if ($testMontriumLogsList -ne $null)
					{
						$testRootweb.dispose()
						$testURL.dispose()
						$waitingForRealSite = $false
					}
					else
					{
						Write-Output "`nThe URL you entered ($AdminLandingSiteCollectionURL) is a valid URL but does not contain a Montrium Logs library"
					}
				}
			} while ($waitingForRealSite)

			Write-Output "The URL you entered for the Admin Landing Site is: $AdminLandingSiteCollectionURL"
			
			#We are assuming that the $AdminLandingSiteCollectionURL ends with /, double check this and ensure it does
			#Check the last character of the site collection URL string
			$lastChar = $AdminLandingSiteCollectionURL.substring($AdminLandingSiteCollectionURL.length - 1, 1)
				
			#Check if it is a /, if not, then need to add it in or it will mess things up
			if ($lastChar -ne "/")
			{
				$AdminLandingSiteCollectionURL = $AdminLandingSiteCollectionURL + "/"
			}
			
			$resultOfCheck = CheckForSiteAccessAndIfAdmin -URLToCheck $AdminLandingSiteCollectionURL -UserName $VerificationUserName

			if ($resultOfCheck -ne "PASS")
			{
				Write-Output "`nERROR: $resultOfCheck, so cancelling"
				stop-transcript
				Exit
			}

			$VerificationUserNameArray = $VerificationUserName.Split("\")
			$domain = $VerificationUserNameArray[0]
			$LoginName = Read-Host "`nPlease enter your login name without domain"
			Write-Output "You entered $LoginName"
			$FullLoginName = $domain + "\" + $LoginName
			
			#Ask for site collection URL
			$waitingForRealSite = $true

			do 
			{
				$TargetSiteCollectionURL = Read-Host "`nPlease enter the RMW site collection URL you wish to update"
				
				#test to see if the site exists
				$testURL = get-spsite $TargetSiteCollectionURL -ErrorAction SilentlyContinue
				
				#if it is equal to null then the site doesn't exist
				if ($testURL -eq $null)
				{
					Write-Output "`nThe URL you entered ($TargetSiteCollectionURL) does not correspond to a valid site collection URL"
				}
				#if it isn't null but has -RC in it then it isn't a work area
				elseif ($TargetSiteCollectionURL -match "-RC")
				{
					Write-Output "`nThe URL you entered ($TargetSiteCollectionURL) is a valid URL but matches -RC so is a Record Center"
				}
				#if it isn't null but doesn't have RMW in it then it isn't a RMW site
				elseif ($TargetSiteCollectionURL -notmatch "RMW")
				{
					Write-Output "`nThe URL you entered ($TargetSiteCollectionURL) is a valid URL but does not match RMW so is not a RMW site"
				}
				#otherwise it does exist and you don't have to keep asking
				else
				{
					$testURL.dispose()
					$waitingForRealSite = $false
				}
			} while ($waitingForRealSite)

			Write-Output "`nThe URL you entered for the RMW site collection is: $TargetSiteCollectionURL"
			
			#We are assuming that the $TargetSiteCollectionURL ends with /, double check this and ensure it does
			#Check the last character of the site collection URL string
			$lastChar = $TargetSiteCollectionURL.substring($TargetSiteCollectionURL.length - 1, 1)
				
			#Check if it is a /, if not, then need to add it in or it will mess things up
			if ($lastChar -ne "/")
			{
				$TargetSiteCollectionURL = $TargetSiteCollectionURL + "/"
			}
			
			$resultOfCheck = CheckForSiteAccessAndIfAdmin -URLToCheck $TargetSiteCollectionURL -UserName $VerificationUserName

			if ($resultOfCheck -ne "PASS")
			{
				Write-Output "`nERROR: $resultOfCheck, so cancelling"
				stop-transcript
				Exit
			}

			Write-Output "`n`n--------------------------$ProductCodeFromXML $ProductVersionFromXML $ScriptType on $ProductSharePointVersionFromXML------------------"	
			Write-Output "`n"	
			Write-Output "`n"
			
			#Call the function that will create the General Info Table
			GenerateGeneralInfoTable
			
			#Call the function that will do the checking
			RMWVerification
			
			stop-transcript
			
			#generate report, which will call the upload report function so that it is uploaded
			GenerateReport
		}
		else
		{
			Write-Output "User $VerificationUserName is not a farm admin so cancelling"
			Exit
		}
	}
	else
	{
		Write-Output "No access to the farm so cancelling"
		Exit
	}
}
else
{
	Write-Output "User $VerificationUserName is not a local admin so cancelling"
	Exit
}

#############################################################################
Exit