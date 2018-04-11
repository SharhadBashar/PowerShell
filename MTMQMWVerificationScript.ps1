###------------------------------------------------------------------------
### Montrium QMW Verification Script: To automate verification currently done manually (for example in IQs)
### Author: Kevin Chambers
### Script version: 1.0
### Date: 3-MAR-2018
### Reference No of related document: MTM-QMW-DSP-XX
### The script is intended to check what is listed in the accompanying XML and comparing any values to those in the XML. 
### The result is put into an HTML file, uploaded to the chosen landing site and converted to PDF.
###
###------------------------------------------------------------------------

#This is needed for function Get-FileName
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

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

function QMWCheckSharePointFeatures
{
#QMWCheckSharePointFeatures: Check SharePoint Features

	#Get array from XML
	$SharePointFeaturesArray = $XMLConfig.XML.SharePointFeaturesArray.SPFeatureInfo
	
	#If are entries in the XML
	if (($SharePointFeaturesArray -ne $null) -or ($SharePointFeaturesArray.Count -gt 0))
	{
		Write-Output "`n--- Checking SharePoint Features ---"
		
		#TODO: set global variable for SharePointFeatures table to true
		
		foreach ($featureToCheck in $SharePointFeaturesArray)
		{
			$SubSiteURLCode = $featureToCheck.SubSiteURLCode
			$featureToCheckWSPName = $featureToCheck.WSPName
			$featureToCheckWSPFile = $featureToCheck.WSPFile
			$featureToCheckWSPVersion = $featureToCheck.WSPVersion
			$featureToCheckLocation = $featureToCheck.Location
			$featureToCheckActiveOrInactive = $featureToCheck.ActiveOrInactive
			$featureToCheckStringToLookFor = $featureToCheck.StringToLookFor
			$featureToCheckElementID = $featureToCheck.ElementID
			$featureToCheckPassFail = "FAIL"
			$featureToCheckValueFound = ""
			
			#TODO: Check for featureToCheck
			
			#TODO: Add result to features table
		}
	}
	else
	{
		Write-Output "`n--- No SharePoint Features to check ---"
	}
}

function QMWCheckInfoPathFormServices
{
#QMWCheckInfoPathFormServices: Check InfoPath Form Services in Central Admin

}

function QMWCheckInfoPathForms
{
#QMWCheckInfoPathForms: Check Form Templates library in QMW to see what forms have been deployed

	#Get array from XML
	$InfoPathFormsArray = $XMLConfig.XML.InfoPathFormsArray.FormInfo
	
	#If are entries in the XML
	if (($InfoPathFormsArray -ne $null) -or ($InfoPathFormsArray.Count -gt 0))
	{
		Write-Output "`n--- Checking Deployed Form Templates ---"
		
		#TODO: set global variable for InfoPathForms table to true
		
		foreach ($formToCheck in $InfoPathFormsArray)
		{
			$formToCheckFormName = $formToCheck.FormName
			$formToCheckPassFail = "FAIL"
			$formToCheckValueFound = ""
			
			#TODO: Check for formToCheck
			#Can do a simple CAML query to see if the form templates library has value
			
			#TODO: Add result to InfoPathForms table
		}
	}
	else
	{
		Write-Output "`n--- No Form Templates to check ---"
	}
}


function QMWCheckSiteSettings([string]$WAOrRC, [string]$QMWSiteURL)
{
#QMWCheckSiteSettings: Check QMW Site Settings
# Parameters:
# 	-> WAOrRC: Whether it is the WA (Work Area) or RC (Records Center) that is to be verified
# 	-> QMWSiteURL: URL of the QMW site collection you want to verify

	#Get array from XML
	$SiteSettingsArray = $XMLConfig.XML.SiteSettingsArray.SettingInfo
	
	#If are entries in the XML
	if (($SiteSettingsArray -ne $null) -or ($SiteSettingsArray.Count -gt 0))
	{
		Write-Output "`n--- Checking Site Settings ---"
		
		#TODO: set global variable for Site Settings table to true
		
		$SPSiteQMW = Get-SPSite $QMWSiteURL
		$QMWrootweb = $SPSiteQMW.rootweb
		
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
					$SettingToCheckValueFound = $QMWrootweb.AllProperties[$SettingToCheckAllPropertiesValue]
					
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
					$SettingToCheckValueFound = $QMWrootweb.AllProperties[$SettingToCheckAllPropertiesValue]
					
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
					
					<#
					#Navigate to site settings, and click on ‘Site Collection Audit Settings’.  For the setting ‘Automatically trim the audit log for this site’, specify ‘No’, and audit all auditable events except ‘Searching Documents’.
					$auditmask = [Microsoft.SharePoint.SPAuditMaskType]::CheckOut -bxor [Microsoft.SharePoint.SPAuditMaskType]::CheckIn -bxor [Microsoft.SharePoint.SPAuditMaskType]::View -bxor 
								 [Microsoft.SharePoint.SPAuditMaskType]::Delete -bxor [Microsoft.SharePoint.SPAuditMaskType]::Update -bxor [Microsoft.SharePoint.SPAuditMaskType]::ProfileChange -bxor
								 [Microsoft.SharePoint.SPAuditMaskType]::ChildDelete -bxor [Microsoft.SharePoint.SPAuditMaskType]::SchemaChange -bxor [Microsoft.SharePoint.SPAuditMaskType]::SecurityChange -bxor
								 [Microsoft.SharePoint.SPAuditMaskType]::Undelete -bxor [Microsoft.SharePoint.SPAuditMaskType]::Workflow -bxor [Microsoft.SharePoint.SPAuditMaskType]::Copy -bxor
								 [Microsoft.SharePoint.SPAuditMaskType]::Move
					
					$SPSiteQMW.TrimAuditLog = $false
					$SPSiteQMW.Audit.AuditFlags = $auditmask
					$SPSiteQMW.Audit.Update()
					Write-Output "`nUpdated Site Collection Audit Settings"
					#>
				}
				else
				{
					Write-Output "`nError: Unable to handle setting type of $SettingToCheckSettingType"
					$SettingTOCheckResult = "Unable to handle setting type of $SettingToCheckSettingType"
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
		
		$QMWrootweb.dispose()
		$SPSiteQMW.dispose()
	}
	else
	{
		Write-Output "`n--- No Site Settings to check ---"
	}
}

function QMWCheckLibrarySettings([string]$WAOrRC, [string]$QMWSiteURL)
{
#QMWCheckLibrarySettings: Check QMW Library Settings
# Parameters:
# 	-> WAOrRC: Whether it is the WA (Work Area) or RC (Records Center) that is to be verified
# 	-> QMWSiteURL: URL of the QMW site collection you want to verify

	#Get array from XML
	$LibrarySettingsArray = $XMLConfig.XML.LibrarySettingsArray.SettingInfo
	
	#If are entries in the XML
	if (($LibrarySettingsArray -ne $null) -or ($LibrarySettingsArray.Count -gt 0))
	{
		Write-Output "`n--- Checking Library Settings ---"
		
		#TODO: set global variable for LibrarySettings table to true
		
		$SPSiteQMW = Get-SPSite $QMWSiteURL
		$QMWrootweb = $SPSiteQMW.rootweb
		
		#Loop through each setting in the array
		foreach ($SettingToCheck in $LibrarySettingsArray)
		{
			#Get the relevant information from the XML file
			$SiteCollection = $SettingToCheck.SiteCollection
			$SettingToCheckLibraryName = $SettingToCheck.LibraryName
			$SettingToCheckSettingName = $SettingToCheck.SettingName
			$SettingToCheckRequiredValue = $SettingToCheck.RequiredValue
			$SettingToCheckPassFail = "FAIL"
			$SettingTOCheckResult = ""
			
			#If the setting is for the current site collection
			if ($SiteCollection -eq $WAOrRC)
			{
				#Get the list
				$targetList = $QMWrootweb.Lists[$SettingToCheckLibraryName]
				
				if ($targetList -ne $null)
				{
					#Check the Setting type
					if ($SettingToCheckSettingName -eq "Version History")
					{
						$versioningEnabled = $targetList.EnableVersioning
						$minorVersionsEnabled = $targetList.EnableMinorVersions
						
						if ($versioningEnabled)
						{
							if ($minorVersionsEnabled)
							{
								$SettingTOCheckResult = "Minor"
							}
							else
							{
								$SettingTOCheckResult = "Major"
							}
						}
						else
						{
							$SettingTOCheckResult = "None"
						}
						
						if ($SettingToCheckRequiredValue -eq $SettingTOCheckResult)
						{
							$SettingToCheckPassFail = "PASS"
						}
					}
					elseif ($SettingToCheckSettingName -eq "Optionally limit the number of versions to retain")
					{
						#TODO:
					}
					
					#TODO put in more options and error handling
				}
				else
				{
					Write-Output "`nError: Unable to find list named $SettingToCheckLibraryName"
					$SettingTOCheckResult = "Unable to find list named $SettingToCheckLibraryName"
				}
				
				
				#TODO: Add result to LibrarySettings table
			}
		}
		
		$QMWrootweb.dispose()
		$SPSiteQMW.dispose()
	}
	else
	{
		Write-Output "`n--- No Library Settings to check ---"
	}
}

function QMWCheckWorkflowConstants
{
#QMWCheckWorkflowConstants: Check Workflow Constants

#TODO:

}


function QMWCheckNintexWorkflows
{
#QMWCheckNintexWorkflows: Check to see what Nintex Workflows were deployed

#TODO:

}


function QMWCheckCoSignSettings
{
#QMWCheckCoSignSettings: Check CoSign library settings in QMW

#TODO:

}


function QMWCheckSecureStore
{
#QMWCheckSecureStore: Check Secure Store in Central Admin

#TODO:

}

function QMWCheckHelpCenter
{
#QMWCheckHelpCenter: Check QMW Help Centers

#TODO:

}
		
function QMWVerification
{
#QMWVerification: Main QMW verification function that calls all the verification sub functions

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
	$QMWURL = $TargetSiteCollectionURL
	
	Write-Output "`nYou chose a URL of $QMWURL"
	
	#Check QMW RC URL
	$waitingForRealSite = $true
	#Update URL now ends in / so need to remove that first
	#$QMWRCURL = $QMWURL + "-RC"
	$QMWRCURL = $QMWURL.Substring(0,$QMWURL.Length-1) + "-RC"
	do 
	{
		#test to see if the site exists
		$testURL = get-spsite $QMWRCURL -ErrorAction SilentlyContinue
		
		#if it is equal to null then the site doesn't exist
		if ($testURL -eq $null)
		{
			Write-Output "`nThe URL ($QMWRCURL) does not correspond to a valid site collection URL" 
			$QMWRCURL = Read-Host "`nPlease enter the site collection URL where QMW Records Center is located"
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
	$GeneralInfoRow["Configuration Value"] = $QMWRCURL
	$GeneralInfoRow["Type"] = "N/A"
	$GeneralInfoRow["Is Current Server"] = "N/A"
	$GeneralInfoRow["Date"] = $Date 
	$GeneralInfoTable.Rows.Add($GeneralInfoRow)
	
	#Call the various functions that can only apply to one location, either Work Area or Central Admin depending
	QMWCheckSharePointFeatures
	QMWCheckInfoPathFormServices
	QMWCheckInfoPathForms
	QMWCheckWorkflowConstants
	QMWCheckNintexWorkflows
	QMWCheckCoSignSettings
	QMWCheckSecureStore
	QMWCheckHelpCenter
	
	#Call the functions that can apply to either the Work Area or RC for the Work Area
	$WAOrRC = "WA"
	QMWCheckSiteSettings $WAOrRC $QMWURL
	QMWCheckLibrarySettings $WAOrRC $QMWURL
	
	#Call the functions that can apply to either the Work Area or RC for the RC
	$WAOrRC = "RC"
	QMWCheckSiteSettings $WAOrRC $QMWRCURL
	QMWCheckLibrarySettings $WAOrRC $QMWRCURL
	
	Write-Output "`n`n QMW Verification Complete!"
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

	#HTML Formatting
	##################################################
	#General Information Fragment formatting
	$GeneralInfoFragment = $GeneralInfoTable | ConvertTo-HTML "Info","Configuration Value","Type","Is Current Server","Date" -fragment
	$GeneralInfoHTML = "<br> General Information about the SharePoint servers:<br>$GeneralInfoFragment<br><hr>"
	
	#Site Settings Fragment formatting
	$SiteSettingsFragment = $SiteSettingsTable | ConvertTo-HTML "WA or RC","Setting Type","Setting Name","Value","Value Required","Pass or Fail","Date" -fragment
	$SiteSettingsHTML = "<br> Site Setting Information:<br>$SiteSettingsFragment<br><hr>"
	
	#TODO: update the below to handle all the tables from all the sections
	
	#CoSign Fragment formatting
	$CoSignFragment = $CoSignTable | ConvertTo-HTML "Parameter","Value","Value Required","Pass or Fail","Date" -fragment
	$CoSignHTML = "<br> CoSign Information:<br>$CoSignFragment<br><hr>"
	
	#Office Web Apps Fragment formatting
	$OfficeWebAppsFragment = $OfficeWebAppsTable | ConvertTo-HTML "Parameter","Value","Value Required","Pass or Fail","Date" -fragment
	$OfficeWebAppsHTML = "<br> Office Web Apps Information:<br>$OfficeWebAppsFragment<br><hr>"
	
	#Nintex Fragment formatting
	$NintexFragment = $NintexTable | ConvertTo-HTML "Parameter","Value","Value Required","Pass or Fail","Date" -fragment
	$NintexHTML = "<br> Nintex Information:<br>$NintexFragment<br><hr>"
	
	#Third Party Program Fragment Formatting
	$ThirdPartyFragement = $ThirdPartyTable | ConvertTo-HTML "Third Party Name","Version","Version Required","License","Pass or Fail","Date" -fragment
	$ThirdPartyHTML = "<br> Third party program information is listed below:<br>$ThirdPartyFragement<br><hr>"
	
	#Montrium WSP Fragment formatting
	$MontriumWSPFragment = $MontriumWSPTable | ConvertTo-HTML "Montrium WSP","Version","Version Required","Pass or Fail","Date" -fragment
	$MontriumWSPHTML = "<br> Montrium's WSP information is listed below:<br>$MontriumWSPFragment<br><hr>"
	
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
	
	$BodyStart = "<h1 style=""color:#9fb11e;font-size:30px"">$ProductCodeFromXML $ProductVersionFromXML $ScriptType</h1><br><h3 style=""color:#9fb11e;margin-left:30px;"">Quick Summary</h3><hr>"
	
	#TODO: update the below to handle all the tables from all the sections
	
	$BodyTag = "$Bodystart $GeneralInfoHTML $SiteSettingsHTML $CoSignHTML $OfficeWebAppsHTML $NintexHTML $ThirdPartyHTML $MontriumWSPHTML $TranscriptHTML"
	$BodyTagPass =  $BodyTag | foreach {if ($_ -match "PASS") {$_ -replace "PASS", "<font color=green><b>PASS</b></font>"}}
	$BodyTagPassFail =  $BodyTagPass | foreach {if ($_ -match "FAIL") {$_ -replace "FAIL", "<font color=red><b>FAIL</b></font>"}}
	$BodyTagPassFailNewLine =  $BodyTagPassFail | foreach {if ($_ -match "_New_Line_") {$_ -replace "_New_Line_", "<br>"}}
	CreateHTMLFile $Path $headTag $BodyTagPassFailNewLine
    
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


################################MAIN##########################################



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

#TODO: Update tables below and/or add tables

#Build CoSign table (formerly Other server table)
$CoSignTable = New-Object system.Data.DataTable "CoSign Information"
$ParameterCol = New-Object system.Data.DataColumn Parameter,([string])
$ParameterCol.ColumnName = "Parameter"
$ValueCol = New-Object system.Data.DataColumn Value,([string])
$ValueCol.ColumnName = "Value"
$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
$ValueRequiredCol.ColumnName = "Value Required"
$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
$PassFailCol.ColumnName = "Pass or Fail"
$DateCol = New-Object system.Data.DataColumn Date,([string])
$DateCol.ColumnName = "Date"
$CoSignTable.columns.add($ParameterCol)
$CoSignTable.columns.add($ValueCol)
$CoSignTable.columns.add($ValueRequiredCol)
$CoSignTable.columns.add($PassFailCol)
$CoSignTable.columns.add($DateCol)

#Build Office Web Apps table
$OfficeWebAppsTable = New-Object system.Data.DataTable "Office Web Apps Information"
$ParameterCol = New-Object system.Data.DataColumn Parameter,([string])
$ParameterCol.ColumnName = "Parameter"
$ValueCol = New-Object system.Data.DataColumn Value,([string])
$ValueCol.ColumnName = "Value"
$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
$ValueRequiredCol.ColumnName = "Value Required"
$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
$PassFailCol.ColumnName = "Pass or Fail"
$DateCol = New-Object system.Data.DataColumn Date,([string])
$DateCol.ColumnName = "Date"
$OfficeWebAppsTable.columns.add($ParameterCol)
$OfficeWebAppsTable.columns.add($ValueCol)
$OfficeWebAppsTable.columns.add($ValueRequiredCol)
$OfficeWebAppsTable.columns.add($PassFailCol)
$OfficeWebAppsTable.columns.add($DateCol)

#Build Nintex table
$NintexTable = New-Object system.Data.DataTable "Nintex Information"
$ParameterCol = New-Object system.Data.DataColumn Parameter,([string])
$ParameterCol.ColumnName = "Parameter"
$ValueCol = New-Object system.Data.DataColumn Value,([string])
$ValueCol.ColumnName = "Value"
$ValueRequiredCol = New-Object system.Data.DataColumn ValueRequired,([string])
$ValueRequiredCol.ColumnName = "Value Required"
$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
$PassFailCol.ColumnName = "Pass or Fail"
$DateCol = New-Object system.Data.DataColumn Date,([string])
$DateCol.ColumnName = "Date"
$NintexTable.columns.add($ParameterCol)
$NintexTable.columns.add($ValueCol)
$NintexTable.columns.add($ValueRequiredCol)
$NintexTable.columns.add($PassFailCol)
$NintexTable.columns.add($DateCol)

#3rd Party Programs Table
$ThirdPartyTable = New-Object system.Data.DataTable "Third Party"
$ThirdPartyCol = New-Object system.Data.DataColumn ThirdParty,([string])
$ThirdPartyCol.ColumnName = "Third Party Name"
$VersionCol = New-Object system.Data.DataColumn Version,([string])
$VersionCol.ColumnName = "Version"
$VersionRequiredCol = New-Object system.Data.DataColumn VersionRequired,([string])
$VersionRequiredCol.ColumnName = "Version Required"
$LicenseCol = New-Object system.Data.DataColumn License,([string])
$LicenseCol.ColumnName = "License"
$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
$PassFailCol.ColumnName = "Pass or Fail"
$DateCol = New-Object system.Data.DataColumn Date,([string])
$DateCol.ColumnName = "Date"
$ThirdPartyTable.columns.add($ThirdPartyCol)
$ThirdPartyTable.columns.add($VersionCol)
$ThirdPartyTable.columns.add($VersionRequiredCol)
$ThirdPartyTable.columns.add($LicenseCol)
$ThirdPartyTable.columns.add($PassFailCol)
$ThirdPartyTable.columns.add($DateCol)

#Montrium WSP Table
$MontriumWSPTable = New-Object system.Data.DataTable "Montrium WSPs"
$MontriumWSPCol = New-Object system.Data.DataColumn MontriumWSP,([string])
$MontriumWSPCol.ColumnName = "Montrium WSP"
$VersionCol = New-Object system.Data.DataColumn Version,([string])
$VersionCol.ColumnName = "Version"
$VersionRequiredCol = New-Object system.Data.DataColumn VersionRequired,([string])
$VersionRequiredCol.ColumnName = "Version Required"
$PassFailCol = New-Object system.Data.DataColumn PassFail,([string])
$PassFailCol.ColumnName = "Pass or Fail"
$DateCol = New-Object system.Data.DataColumn Date,([string])
$DateCol.ColumnName = "Date"
$MontriumWSPTable.columns.add($MontriumWSPCol)
$MontriumWSPTable.columns.add($VersionCol)
$MontriumWSPTable.columns.add($VersionRequiredCol)
$MontriumWSPTable.columns.add($PassFailCol)
$MontriumWSPTable.columns.add($DateCol)

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

$TranscriptFileName = "QMW Verification Script"

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
				$TargetSiteCollectionURL = Read-Host "`nPlease enter the QMW site collection URL you wish to update"
				
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
				#if it isn't null but doesn't have QMW in it then it isn't a QMW site
				elseif ($TargetSiteCollectionURL -notmatch "QMW")
				{
					Write-Output "`nThe URL you entered ($TargetSiteCollectionURL) is a valid URL but does not match QMW so is not a QMW site"
				}
				#otherwise it does exist and you don't have to keep asking
				else
				{
					$testURL.dispose()
					$waitingForRealSite = $false
				}
			} while ($waitingForRealSite)

			Write-Output "`nThe URL you entered for the QMW site collection is: $TargetSiteCollectionURL"
			
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
			QMWVerification
			
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