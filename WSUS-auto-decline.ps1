####################################################
#
#   WSUS-auto-decline.ps1      
#
####################################################
#
# Decline un-needed WSUS Updates to save disk space on server and reduce clutter in WSUS Server manager
# Original Author: "nadnerB"
# https://community.spiceworks.com/scripts/show/2795-decline-wsus-updates
# 
# Run powershell script from task scheduler:  https://community.spiceworks.com/how_to/17736-run-powershell-scripts-from-task-scheduler
#
##########################
#       CHANGELOG
##########################
# 2016-04-21
# By John Lucas ==> Started modifying Script
#
# 2016-04-22
# By John Lucas ==> Updated to work with 2012R2
#   - corrected script for Server 2012 R2 based on comments at:
#     https://gallery.technet.microsoft.com/scriptcenter/Automatically-Declining-a4fec7be/view/Discussions#content
#
#   - Invoke the WSUS server cleanup wizard functionality direclty with powershell
#     Documentation at: https://technet.microsoft.com/en-us/library/hh826162.aspx
#
#     Invoke-WsusServerCleanup [-CleanupObsoleteComputers] [-CleanupObsoleteUpdates] [-CleanupUnneededContentFiles]
#       [-CompressUpdates] [-DeclineExpiredUpdates] [-DeclineSupersededUpdates] [-InformationAction
#       <System.Management.Automation.ActionPreference> {SilentlyContinue | Stop | Continue | Inquire | Ignore | Suspend} ]
#       [-InformationVariable <System.String> ] [-UpdateServer <IUpdateServer> ] [-Confirm] [-WhatIf] [ <CommonParameters>]
#
#   - Have this script self-elevate to administrator level permissions
#     See: https://blogs.msdn.microsoft.com/virtual_pc_guy/2010/09/23/a-self-elevating-powershell-script/
#
# 2017-03-15
# By John Lucas ==> added auto-decline for "Technical Preview" and "Insider Preview"
#
# 2017-07-25
# By John Lucas ==> added auto-decline for 64bit Project & Access (in code, added to "Office" exclusions list)
#
# 2017-08-14
# By John Lucas ==> added auto-decline for 64bit OneNote (in code, added to "Office" exclusions list)
#
# 2018-05-10  
# By John Lucas ==> added auto-decline for "Preview Of"
#
# 2019-01-24
# By John Lucas ==> added auto-decline for "ARM64-based" items.
#
##########################



#############################################################################
##    Begin script elevation code
# Get the ID and security principal of the current user account
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
 
# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
 
$MyOrigBackColor = $Host.UI.RawUI.BackgroundColor
# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))
   {
   # We are running "as Administrator" - so change the title and background color to indicate this
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
   # For colors, see: http://windowsitpro.com/powershell/take-control-powershell-consoles-colors
   $Host.UI.RawUI.BackgroundColor = "DarkRed"
   clear-host
   }
else
   {
   # We are not running "as Administrator" - so relaunch as administrator
   
   # Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   
   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;
   
   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";
   
   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess);
   
   # Exit from the current, unelevated, process
   exit
   }
 
# Run your code that needs to be elevated here
#Write-Host -NoNewLine "Press any key to continue..."
#$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

##    End script elevation code
#############################################################################


#############################################################################
##    Begin WSUS Auto Decline stuff

#$WsusServer = "<Insert WSUS server FQDN>"
#$PortNumber = <Insert default WSUS Port>

$WsusServer = "mywsus1.example.local"
$UseSSL = $false
$PortNumber = 8530
$TrialRun = 0
# 1 = Yes
# 0 = No

$ServerCleanupOutput = "Sorry, No Data Returned"

# Deliver the results by email
Function Mailer
    {
     $emailTo = "myemail@example.com"
     $emailFrom = "WSUS-PS-autodecline.support@example.com" 
     $subject="WSUS - Declined Updates" 
     $smtpserver="internalemailrelay.example.com" 
     $smtp=new-object System.Net.Mail.SmtpClient($smtpServer)
     $Message = @" 
    
Package Declined Status:	
    $IA64_counted Itanium updates have been declined.
    $Office64_count MS Office 64-Bit updates have been declined.
    $sharepoint_counted SharePoint updates have been declined.
    $technicalpreview_counted Technical Preview updates have been declined.
    $insiderpreview_counted Insider Preview updates have been declined.
    $previewof_counted Preview of -- updates have been declined.
    $arm64based_counted AMD64-based updates have been declined.

WSUS server cleanup wizard output:	
$ServerCleanupOutput
 
Thank you, 
Your trusty WSUS server's powershell cleanup script

"@ 
     If ($TrialRun -eq 1)
        {
            $Subject += " Trial Run"
        }

     $smtp.Send($emailFrom, $emailTo, $subject, $message)
    }

# Connect to the WSUS 3.0 interface.
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null

#Next line is not compatible with 2012R2
#$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer,$UseSSL,$PortNumber);

#This next line works with 2012R2
$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]:: getUpdateServer();

# Searching in just the title of the update
# Itanium/IA64
$itanium = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "ia64|itanium"}
$IA64_counted = $itanium.count
    If ($itanium.count -lt 1)
        {
            $IA64_counted = 0
        }
    If ($TrialRun -eq 0 -and $itanium.count -gt 0)
        {
            $itanium  | %{$_.Decline()}
        }
# MS Office 64-Bit
$Office64 = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "Excel|Lync|Office|Outlook|Powerpoint|Visio|word|Project|Access|OneNote" -and $_.Title -match "64-bit"}
$Office64_count = $Office64.count
    If ($Office64.count -lt 1)
        {
            $Office64_count = 0
        }
    If ($TrialRun -eq 0 -and $Office64.count -gt 0)
        {
            $Office64 | %{$_.Decline()}
        }
# SharePoint
$sharepoint = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "SharePoint"}
$sharepoint_counted = $sharepoint.count
    If ($sharepoint.count -lt 1)
        {
            $sharepoint_counted = 0
        }
    If ($TrialRun -eq 0 -and $sharepoint.count -gt 0)
        {
            $sharepoint | %{$_.Decline()}
        }
# Technical Preview
$technicalpreview = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "Technical Preview"}
$technicalpreview_counted = $technicalpreview.count
    If ($technicalpreview.count -lt 1)
        {
            $technicalpreview_counted = 0
        }
    If ($TrialRun -eq 0 -and $technicalpreview.count -gt 0)
        {
            $technicalpreview | %{$_.Decline()}
        }
# Insider Preview
$insiderpreview = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "Insider Preview"}
$insiderpreview_counted = $insiderpreview.count
    If ($insiderpreview.count -lt 1)
        {
            $insiderpreview_counted = 0
        }
    If ($TrialRun -eq 0 -and $insiderpreview.count -gt 0)
        {
            $insiderpreview | %{$_.Decline()}
        }

# Preview Of
$previewof = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "Preview Of"}
$previewof_counted = $previewof.count
    If ($previewof.count -lt 1)
        {
            $previewof_counted = 0
        }
    If ($TrialRun -eq 0 -and $previewof.count -gt 0)
        {
            $previewof | %{$_.Decline()}
        }

# ARM64-based
$arm64based = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "ARM64-based"}
$arm64based_counted = $arm64based.count
    If ($arm64based.count -lt 1)
        {
            $arm64based_counted = 0
        }
    If ($TrialRun -eq 0 -and $arm64based.count -gt 0)
        {
            $arm64based | %{$_.Decline()}
        }


##    End WSUS Auto Decline stuff
#############################################################################


#############################################################################
##    Begin the WSUS Server cleanup wizard
If ($TrialRun -eq 1)
	{
	  $ServerCleanupOutput = "This is the Trial output"
	  Get-WsusServer | Invoke-WsusServerCleanup -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates -DeclineExpiredUpdates -DeclineSupersededUpdates -WhatIf -outvariable ServerCleanupOutput
	  #$ServerCleanupOutput = Get-WsusServer | Invoke-WsusServerCleanup -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates -DeclineExpiredUpdates -DeclineSupersededUpdates -WhatIf -verbose
	  
	  #Note: Whatif output can't be logged to variable.  
	  #  See: https://www.reddit.com/r/PowerShell/comments/2rmt7z/how_can_i_log_whatif_output/
	  $ServerCleanupOutput += " (As this script was set to run in trial mode, please check the console for verbose -WhatIf output)"
	}
else
	{
	  $ServerCleanupOutput = "This is the real output"
      Get-WsusServer | Invoke-WsusServerCleanup -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates -DeclineExpiredUpdates -DeclineSupersededUpdates -outvariable ServerCleanupOutput
	}

# call email send function
Mailer

If ($TrialRun -eq 1)
	{
    # Return colors and WindowTitle to normal
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition
    $Host.UI.RawUI.BackgroundColor = $MyOrigBackColor
    Write-Host -NoNewLine "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# cmd /c pause | out-null
