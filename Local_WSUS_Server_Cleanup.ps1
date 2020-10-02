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
#$ErrorActionPreference = "SilentlyContinue"

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
	$Message_Body = "WSUS: Cleanup Results From " + $env:computername +"<br>"

# E-Mail Server Credentials to send report (Leave Blank if not Required)
	$Cfg_Smtp_User = ""
	$Cfg_Smtp_Password = ""

# Misc Variables
    $Cleanup_Results = $null

    $LogFile = "C:\PS_Scripts\Cleanup-Log.txt"
    #$SQLInstance = "\\.\pipe\MSSQL$MICROSOFT##SSEE\sql\query" # Use with Pre Windows 2012, or SQL Express
    $SQLInstance = "\\.\pipe\Microsoft##WID\tsql\query" # Use with Windows 2012 and later Windows Internal Database
    $SQLDB = "SUSDB"



# *************************** FUNCTION SECTION ***************************

function Invoke-SQL {
    param(
        [string] $dataSource = ".\SQLEXPRESS",
        [string] $database = "MasterData",
        [string] $sqlCommand = $(throw "Please specify a query.")
      )

    $connectionString = "Data Source=$dataSource; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$database"
    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $command.CommandTimeout = 3600;
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables
}

# ****************************** CODE START ******************************

$StartTime = get-date
Write-Host "Starting at $StartTime"

do {

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
    $cleanupScope.CleanupObsoleteComputers = $false # Removes Computers not checking in in 30 days, better to know what's not checking in
    $cleanupScope.CleanupUnneededContentFiles = $true
    $cleanupManager = $wsus.GetCleanupManager()
    $Cleanup_Results = $cleanupManager.PerformCleanup($cleanupScope)

    if($Cleanup_Results -eq $null) { # If the cleanup function fails, cleanup updates directly in SQL as per https://social.technet.microsoft.com/Forums/ie/en-US/7b12f8b2-d0e6-4f63-a98a-019356183c29/getting-past-wsus-cleanup-wizard-time-out-removing-unnecessary-updates?forum=winserverwsus
        $FailedStartTime = get-date
        Write-Host "Starting at $FailedStartTime"
        Write-Output "$FailedStartTime - Starting" | Out-File -FilePath $LogFile -Force

        $SQLCmd = "exec spGetObsoleteUpdatesToCleanup"
        $UpdatesToCleanup = Invoke-SQL -dataSource $SQLInstance -database $SQLDB -sqlCommand $SQLCmd
        $UpdatesDone = 1
        $UpdatesTotal = ($UpdatesToCleanup.LocalUpdateID).count
        $Timestamp = (get-date).ToString("HH:mm:ss")
        Write-Host "$Timestamp - Processing $UpdatesTotal Updates"
        $message_body += "Updates Processed through SQL: $UpdatesTotal <br>"
        Write-Output "$Timestamp - Processing $UpdatesTotal Updates" | Out-File -FilePath $LogFile -Append
        foreach ($Update in $UpdatesToCleanup) {
            $UStartTime = get-date
            $Timestamp = $UStartTime.ToString("HH:mm:ss")
            $UpdateKB = $Update.ItemArray
            Write-Host "$Timestamp - Processing KB $UpdateKB($UpdatesDone of $UpdatesTotal)"
            Write-Output "$Timestamp - Processing KB $UpdateKB ($UpdatesDone of $UpdatesTotal)" | Out-File -FilePath $LogFile -Append
            $SQLCmd = "exec spDeleteUpdate @localUpdateID="+$Update.ItemArray
            Invoke-SQL -dataSource $SQLInstance -database $SQLDB -sqlCommand $SQLCmd
            $UpdatesDone++
            $UEndTime = get-date
            $Timestamp = $UEndTime.ToString("HH:mm:ss")
            $URunTime = $UEndTime - $UStartTime
            $UFormatRunTime = "{0:N2}" -f $URunTime.TotalMinutes
            Write-Output "$Timestamp - Processed $UpdateKB in $UFormatRunTime mins" | Out-File -FilePath $LogFile -Append
        }
        $FailedEndTime = get-date
        $FailedRunTime = $FailedEndTime - $FailedStartTime
        $FormatRunTime = "{0:N2}" -f $FailedRunTime.TotalMinutes
        $FormatEndTime = ($FailedEndTime).ToString("HH:mm:ss")

        Write-Host "Finished at $EndTime"
        Write-Host "Job took $FormatRunTime minutes to run"
        Write-Output "$FormatEndTime - Finished. Job took $FormatRunTime minutes to run" | Out-File -FilePath $LogFile -Append
    }

} until ($Cleanup_Results -ne $null)

if($cleanupScope.DeclineSupersededUpdates) {
    $SupersededUpdatesDeclined = "{0:N0}" -f $Cleanup_Results.SupersededUpdatesDeclined
    $message_body += "SupersededUpdatesDeclined : " + $SupersededUpdatesDeclined + "<br>"
    write-host "SupersededUpdatesDeclined : $SupersededUpdatesDeclined"
} else {
    $message_body += "SupersededUpdatesDeclined : Skipped<br>"
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