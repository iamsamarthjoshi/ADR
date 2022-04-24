#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '9/24/2021 9:28:40 PM'.

# Uncomment the line below if running in an environment where script signing is 
# required.
#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Site configuration
$SiteCode = "CM1" # Site code 
$ProviderMachineName = "CM1.corp.contoso.com" # SMS Provider machine name

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

##########################################################################################################################################################
#Optional section - Depends on the name of your SUG. Current SUG name assumption - <year - month - OS>

$date = Get-Date

#Getting the value for Current Month

$year = $date.Year
$monthnum = $date.Month

$month = switch ( $monthnum )
{
    1 { 'Jan'    }
    2 { 'Feb'   }
    3 { 'Mar' }
    4 { 'Apr'  }
    5 { 'May'    }
    6 { 'Jun'  }
    7 { 'Jul'  }
    8 { 'Aug'  }
    9 { 'Sep'  }
    10 { 'Oct'  }
    11 { 'Nov'  }
    12 { 'Dec'  }
}

##########################################################################################################################################################

#Get the value for Patch Tuesday
#Source - https://github.com/tsrob50/Get-PatchTuesday

Function Get-PatchTuesday {
  [CmdletBinding()]
  Param
  (
    [Parameter(position = 0)]
    [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
    [String]$weekDay = 'Tuesday',
    [ValidateRange(0, 5)]
    [Parameter(position = 1)]
    [int]$findNthDay = 2
  )
  # Get the date and find the first day of the month
  # Find the first instance of the given weekday
  [datetime]$today = [datetime]::NOW
  $todayM = $today.Month.ToString()
  $todayY = $today.Year.ToString()
  [datetime]$strtMonth = $todayM + '/1/' + $todayY
  while ($strtMonth.DayofWeek -ine $weekDay ) { $strtMonth = $StrtMonth.AddDays(1) }
  $firstWeekDay = $strtMonth

  # Identify and calculate the day offset
  if ($findNthDay -eq 1) {
    $dayOffset = 0
  }
  else {
    $dayOffset = ($findNthDay - 1) * 7
  }
  
  # Return date of the day/instance specified
  $patchTuesday = $firstWeekDay.AddDays($dayOffset) 
  return $patchTuesday
}

$patchTuesdayValue = Get-PatchTuesday

##########################################################################################################################################################

#text to look for inside the Excel file. May vary as per your requirements.
#MQR stands for Monthly Quality Rollup

$MQRNaming1 = New-ConditionalText 'Cumulative Update for Windows Server' -ConditionalTextColor 'Red'
$MQRNaming2 = New-ConditionalText 'Cumulative Security Update' -ConditionalTextColor 'Blue'
$MQRNaming3 = New-ConditionalText 'Security Monthly Quality Rollup' -ConditionalTextColor 'Red'
$MQRNaming4 = New-ConditionalText 'Cumulative Update for Microsoft server' -ConditionalTextColor 'Red'

$SecNaming = New-ConditionalText 'Security Only' -ConditionalTextColor 'Yellow'

#Storing the names of the monthly SUGs
$MQRSUGName = "$year-$month - OS - MQR"
$SecOnlySUGName = "$year-$month - OS - Sec Only"
$MQRNETSUGName = "$year-$month - NET - MQR"
$SecOnlyNETSUGName = "$year-$month - NET - Sec Only"

#Path to store the output file
$resultExcelPath = "$env:SystemDrive\temp\$month-$year.xlsx"

#Delete if the file exists previously
Remove-Item -Path $resultExcelPath -ErrorAction Ignore

##########################################################################################################################################################

#Getting the applicable software update category for subscribed Products

$cat = Get-CMSoftwareUpdateCategory -Fast -TypeName "Product" | Where-Object { $_.IsSubscribed }

#Getting the software updates released after patch Tuesday and Exporting the result to xlsx
Get-CMSoftwareUpdate -Fast -Category $cat -DatePostedMin $patchTuesdayValue -IsSuperseded $false | 
Select-Object ArticleID,LocalizedDisplayName,DatePosted,DateRevised,IsDeployed,IsSuperseded,NumTotal,NumMissing,NumPresent,NumUnknown,PercentCompliant,MaxExecutionTime,LocalizedInformativeURL |
Export-Excel -Path $resultExcelPath -AutoSize -AutoFilter -WorksheetName 'MEMCM Data' -ConditionalText $MQRNaming1, $MQRNaming2, $MQRNaming3, $SecNaming, $MQRNaming4

#Getting the software updates inside the Monthly Quality Rollup SUG Exporting the result to xlsx
Get-CMSoftwareUpdate -Fast -UpdateGroupName $MQRSUGName | 
Select-Object ArticleID,LocalizedDisplayName,DatePosted,DateRevised,IsDeployed,IsSuperseded,NumTotal,NumMissing,NumPresent,NumUnknown,PercentCompliant,MaxExecutionTime,LocalizedInformativeURL |
Export-Excel -Path $resultExcelPath -AutoSize -AutoFilter -WorksheetName 'ADR - OS - MQR' -Append -ConditionalText $MQRNaming1, $MQRNaming2, $MQRNaming3, $SecNaming, $MQRNaming4

#Getting the software updates inside the Security-only Updates SUG Exporting the result to xlsx
Get-CMSoftwareUpdate -Fast -UpdateGroupName $SecOnlySUGName | 
Select-Object ArticleID,LocalizedDisplayName,DatePosted,DateRevised,IsDeployed,IsSuperseded,NumTotal,NumMissing,NumPresent,NumUnknown,PercentCompliant,MaxExecutionTime,LocalizedInformativeURL |
Export-Excel -Path $resultExcelPath -AutoSize -AutoFilter -WorksheetName 'ADR - OS - Sec' -Append -ConditionalText $MQRNaming1, $MQRNaming2, $MQRNaming3, $SecNaming, $MQRNaming4

#Getting the software updates inside the .NET Monthly Quality Rollup SUG Exporting the result to xlsx
Get-CMSoftwareUpdate -Fast -UpdateGroupName $MQRNETSUGName | 
Select-Object ArticleID,LocalizedDisplayName,DatePosted,DateRevised,IsDeployed,IsSuperseded,NumTotal,NumMissing,NumPresent,NumUnknown,PercentCompliant,MaxExecutionTime,LocalizedInformativeURL |
Export-Excel -Path $resultExcelPath -AutoSize -AutoFilter -WorksheetName 'ADR - NET - MQR' -Append -ConditionalText $MQRNaming1, $MQRNaming2, $MQRNaming3, $SecNaming, $MQRNaming4

#Getting the software updates inside the .NET Security Only Updates SUG Exporting the result to xlsx
Get-CMSoftwareUpdate -Fast -UpdateGroupName $SecOnlyNETSUGName | 
Select-Object ArticleID,LocalizedDisplayName,DatePosted,DateRevised,IsDeployed,IsSuperseded,NumTotal,NumMissing,NumPresent,NumUnknown,PercentCompliant,MaxExecutionTime,LocalizedInformativeURL |
Export-Excel -Path $resultExcelPath -AutoSize -AutoFilter -WorksheetName 'ADR - NET - Sec' -Append -ConditionalText $MQRNaming1, $MQRNaming2, $MQRNaming3, $SecNaming, $MQRNaming4

#Sending the mail to the team with the file attached
Send-MailMessage -From 'Senders email address' -To 'Receivers email address' -Subject "Monthly Patches for $month - $year" -Body "Please find the file attached." -Attachments $resultExcelPath -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer 'smtp server address'
