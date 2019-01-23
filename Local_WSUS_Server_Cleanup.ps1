# WSUS Server Cleanup Script - Local Only
# Written By: Steve Lunn (s.m.lunn@wolftech.f9.co.uk)
# Downloaded From: https://github.com/Gilgamoth

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see https://www.gnu.org/licenses/

Clear-Host
Set-PSDebug -strict
$ErrorActionPreference = "SilentlyContinue"

# **************************** CONFIG SECTION ****************************

# Declare Variables

# Only needed if not localhost or not default port, leave blank otherwise.
    $Cfg_WSUSServer = "WSUSServer.FQDN.local" # WSUS Server Name
    $Cfg_WSUSSSL = $true # WSUS Server Using SSL
    $Cfg_WSUSPort = 8531 # WSUS Port Number

# E-Mail Report Details
	$Cfg_Email_To_Address = "recipient@domain.local"
	$Cfg_Email_From_Address = "WSUS-Report@domain.local"
	$Cfg_Email_Subject = "WSUS: Cleanup Results From " + $env:computername
	$Cfg_Email_Server = "mail.domain.local"
    $Cfg_Email_Send = $false
    $Cfg_Email_Send = $true # Comment out if no e-mail required

# E-Mail Server Credentials to send report (Leave Blank if not Required)
	$Cfg_Smtp_User = ""
	$Cfg_Smtp_Password = ""

    $Cleanup_Results = $null

# *************************** FUNCTION SECTION ***************************

# ****************************** CODE START ******************************

$StartTime = get-date
Write-Host "Starting at $StartTime"

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
If($Cfg_WSUSServer) {
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($Cfg_WSUSServer, $Cfg_WSUSSSL, $Cfg_WSUSPort)
} else {
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer()
}

$cleanupScope = new-object Microsoft.UpdateServices.Administration.CleanupScope;
$cleanupScope.DeclineSupersededUpdates = $true
$cleanupScope.DeclineExpiredUpdates = $true
$cleanupScope.CleanupObsoleteUpdates = $true
$cleanupScope.CompressUpdates = $true
$cleanupScope.CleanupObsoleteComputers = $false
$cleanupScope.CleanupUnneededContentFiles = $true
$cleanupManager = $wsus.GetCleanupManager()
$Cleanup_Results = $cleanupManager.PerformCleanup($cleanupScope)

if($Cleanup_Results -ne $null) {
    if($cleanupScope.DeclineSupersededUpdates) {
        $SupersededUpdatesDeclined = "{0:N0}" -f $Cleanup_Results.SupersededUpdatesDeclined
        $message_body = "SupersededUpdatesDeclined : " + $SupersededUpdatesDeclined + "<br>"
        write-host "SupersededUpdatesDeclined : $SupersededUpdatesDeclined"
    } else {
        $message_body = "SupersededUpdatesDeclined : Skipped<br>"
        write-host "SupersededUpdatesDeclined : Skipped"
    }

    If($cleanupScope.DeclineExpiredUpdates) {
        $ExpiredUpdatesDeclined = "{0:N0}" -f $Cleanup_Results.ExpiredUpdatesDeclined
        $message_body += "ExpiredUpdatesDeclined    : " + $ExpiredUpdatesDeclined + "<br>"
        Write-Host "ExpiredUpdatesDeclined    : $ExpiredUpdatesDeclined"
    } else {
        $message_body += "ExpiredUpdatesDeclined    : Skipped<br>"
        Write-Host "ExpiredUpdatesDeclined    : Skipped"
    }

    If($cleanupScope.CleanupObsoleteUpdates) {
        $ObsoleteUpdatesDeleted = "{0:N0}" -f $Cleanup_Results.ObsoleteUpdatesDeleted
        $message_body += "ObsoleteUpdatesDeleted    : " + $ObsoleteUpdatesDeleted + "<br>"
        write-host "ObsoleteUpdatesDeleted    : $ObsoleteUpdatesDeleted"
    } else {
        $message_body += "ObsoleteUpdatesDeleted    : Skipped<br>"
        write-host "ObsoleteUpdatesDeleted    : Skipped"
    }

    If($cleanupScope.CompressUpdates) {
        $UpdatesCompressed = "{0:N0}" -f $Cleanup_Results.UpdatesCompressed
        $message_body += "UpdatesCompressed         : " + $UpdatesCompressed + "<br>"
        write-host "UpdatesCompressed         : $UpdatesCompressed"
    } else {
        $message_body += "UpdatesCompressed         : Skipped<br>"
        write-host "UpdatesCompressed         : Skipped"
    }

    If($cleanupScope.CleanupObsoleteComputers) {
        $ObsoleteComputersDeleted = "{0:N0}" -f $Cleanup_Results.ObsoleteComputersDeleted
        $message_body += "ObsoleteComputersDeleted  : " + $ObsoleteComputersDeleted + "<br>"
        write-host "ObsoleteComputersDeleted  : $ObsoleteComputersDeleted"
    } else {
        $message_body += "ObsoleteComputersDeleted  : Skipped<br>"
        write-host "ObsoleteComputersDeleted  : Skipped"
    }

    If($cleanupScope.CleanupUnneededContentFiles) {
        $DiskSpaceFreed = "{0:N0}" -f ($Cleanup_Results.DiskSpaceFreed/1MB)
        $message_body += "DiskSpaceFreed            : " + $DiskSpaceFreed + " MB<br>"
        write-host "DiskSpaceFreed            : $DiskSpaceFreed MB"
    } else {
        $message_body += "DiskSpaceFreed            : Skipped<br>"
        write-host "DiskSpaceFreed            : Skipped"
    }
	
	$Cfg_Email_Subject += " Success"
}
else {
    $message_body += "Warning: Clean Up Timed Out!"
    write-host "Warning: Clean Up Timed Out!"
	$Cfg_Email_Subject += " FAILED"
}
$EndTime = get-date
$RunTime = $EndTime - $StartTime
$FormatTime = "{0:N2}" -f $RunTime.TotalMinutes

Write-Host "Finished at $EndTime"
Write-Host "Job took $FormatTime minutes to run"

if ($Cfg_Email_Send) {
    $message_body += "<br>Started at $StartTime<br>"
    $message_body += "Finished at $EndTime<br>"
    $message_body += "Job took $FormatTime minutes to run<br>"

    Write-Host "`nSending Report E-Mail to" $Cfg_Email_To_Address
    $smtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
    if ($Cfg_Smtp_User) {
		$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
	}
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
    $message.To.Add($Cfg_Email_To_Address)
    $message.Subject = $Cfg_Email_Subject + " - " + (get-date).ToString("dd/MM/yyyy")
    $message.isBodyHtml = $true
    $message.Body = $Message_Body
    $smtp.Send($message)
}