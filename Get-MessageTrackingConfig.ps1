<#
.SYNOPSIS
Get-MessageTrackingConfig.ps1

.DESCRIPTION 
Generate a CSV report of the message tracking configuration of your Exchange 2013 servers.

.EXAMPLE
.\Get-MessageTrackingConfig.ps1

.LINK

.NOTES
Written by: Paul Cunningham

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Change Log:
V1.00, 14/04/2015 - Initial version
#>

#requires -version 2

[CmdletBinding()]

param ()

#----------------------------------------
# Initialize
#----------------------------------------

if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
    . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
else
{
    Write-Host "Exchange management tools are not installed on this computer."
    EXIT
}


#----------------------------------------
# Variables
#----------------------------------------

$Exchange2013Report = @()
$Exchange2010Report = @() # Exchange 2010 servers not supported by this script yet

$now = Get-Date


#----------------------------------------
# Main Script
#----------------------------------------

$exchangeservers = @(Get-ExchangeServer | Where {$_.IsHubTransportServer -or $_.IsMailboxServer} | Sort AdminDisplayVersion,Name)

# Loop through the servers
foreach ($srv in $exchangeservers)
{
   
    Write-Host "------ Processing $($srv.Name)"

    $uncpath = $null
    $oldest = $null
    $ActualTotalSize = $null

    $version = $null

    $TransportServiceLogFiles = $null
    $ModeratedTransportServiceLogFiles = $null
    $MailboxServerLogFiles = $null
    $MailboxTransportDeliveryLogFiles = $null
    $mailboxTransportSubmissionLogFiles = $null

    $E14TransportRole = $null
    $E14MailboxRole = $null
    $E15TransportRole = $null


    # Get friendly version name and server roles
    if ($srv.AdminDisplayVersion -like "Version 15*")
    {
        $version = "Exchange Server 2013"
        
        # Get Transport service settings
        try
        {
            $E15TransportRole = Get-TransportService $srv.Identity -ErrorAction STOP
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }
    }

    if ($srv.AdminDisplayVersion -like "Version 14*")
    {
        $version = "Exchange Server 2010"
        
        # Get Hub Transport role settings
        try
        {
            $E14TransportRole = Get-TransportServer $srv.Identity -ErrorAction STOP
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }

        # Get Mailbox role settings
        try
        {
            $E14MailboxRole = Get-MailboxServer $srv.Identity -ErrorAction STOP
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }
    }

   
    # Process Exchange 2013 Mailbox Server
    if ($version -eq "Exchange Server 2013")
    {
 
        # Create custom object to store results
        $serverObj = New-Object PSObject
	    $serverObj | Add-Member NoteProperty -Name "Server Name" -Value $srv.Name
	    $serverObj | Add-Member NoteProperty -Name "Server Version" -Value $version
        $serverObj | Add-Member NoteProperty -Name "Message Tracking Enabled" -Value $E15TransportRole.MessageTrackingLogEnabled
        $serverObj | Add-Member NoteProperty -Name "Max Age (Days)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Oldest Log File (Days)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Configured Max Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Estimated Max Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Total Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Transport Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Moderated Transport Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Mailbox Transport Delivery Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Mailbox Transport Submission Size (MB)" -Value "n/a"
        $serverObj | Add-Member NoteProperty -Name "Subject Logging" -Value "n/a"    
        $serverObj | Add-Member NoteProperty -Name "Log Path" -Value "n/a"

        if ($E15TransportRole.MessageTrackingLogEnabled -eq $true)
        {
            
            # Calculate UNC path to message tracking log files
            $uncpath = "\\$($srv.Name)\" + ($($E15TransportRole.MessageTrackingLogPath) -replace(":","$"))

            # Collect the log files
            $LogFiles = @(Get-ChildItem $uncpath\MSGTRK*.log)

            # Break out log files into different types
            $TransportServiceLogFiles = @($Logfiles | Where {$_.Name -match "MSGTRK\d+"})
            $ModeratedTransportServiceLogFiles = @($Logfiles | Where {$_.Name -match "MSGTRKMA\d+"})
            $MailboxTransportDeliveryLogFiles = @($Logfiles | Where {$_.Name -match "MSGTRKMD\d+"})
            $MailboxTransportSubmissionLogFiles = @($Logfiles | Where {$_.Name -match "MSGTRKMS\d+"})

            # Calculate the oldest of all the log files
            $oldest = $LogFiles | Sort LastWriteTime | Select -First 1
      
            #Calculate the total size of each log file type
            [int]$TransportServiceActualSize = ($TransportServiceLogFiles | Measure-Object Length -Sum).Sum/1MB
            [int]$ModeratedTransportServiceActualSize = ($ModeratedTransportServiceLogFiles | Measure-Object Length -Sum).Sum/1MB
            [int]$MailboxTransportDeliveryActualSize = ($MailboxTransportDeliveryLogFiles | Measure-Object Length -Sum).Sum/1MB
            [int]$MailboxTransportSubmissionActualSize = ($MailboxTransportSubmissionLogFiles | Measure-Object Length -Sum).Sum/1MB

            #Calculate the total size
            [int]$ActualTotalSize = $TransportServiceActualSize + $ModeratedTransportServiceActualSize + $MailboxTransportDeliveryActualSize + $MailboxTransportSubmissionActualSize

            # Add age values to the custom object
            $serverObj | Add-Member NoteProperty -Name "Max Age (Days)" -Value $($E15TransportRole.MessageTrackingLogMaxAge.Days) -Force
            $serverObj | Add-Member NoteProperty -Name "Oldest Log File (Days)" -Value $(($now - $oldest.LastWriteTime).Days) -Force

            # Add size values to the custom object
            $serverObj | Add-Member NoteProperty -Name "Configured Max Size (MB)" -Value $($E15TransportRole.MessageTrackingLogMaxDirectorySize.Value.ToMb()) -Force
            $serverObj | Add-Member NoteProperty -Name "Estimated Max Size (MB)" -Value $($E15TransportRole.MessageTrackingLogMaxDirectorySize.Value.ToMb()*3) -Force
            $serverObj | Add-Member NoteProperty -Name "Total Size (MB)" -Value $ActualTotalSize -Force
            $serverObj | Add-Member NoteProperty -Name "Transport Size (MB)" -Value $TransportServiceActualSize -Force
            $serverObj | Add-Member NoteProperty -Name "Moderated Transport Size (MB)" -Value $ModeratedTransportServiceActualSize -Force
            $serverObj | Add-Member NoteProperty -Name "Mailbox Transport Delivery Size (MB)" -Value $MailboxTransportDeliveryActualSize -Force
            $serverObj | Add-Member NoteProperty -Name "Mailbox Transport Submission Size (MB)" -Value $MailboxTransportSubmissionActualSize -Force

            # Add other settings to the custom object
            $serverObj | Add-Member NoteProperty -Name "Subject Logging" -Value $E15TransportRole.MessageTrackingLogSubjectLoggingEnabled -Force    
            $serverObj | Add-Member NoteProperty -Name "Log Path" -Value $E15TransportRole.MessageTrackingLogPath -Force
        }
        else
        {
            # Message tracking is not enabled for this server
        }

    $Exchange2013Report += $serverObj
    
    }

    # Process Exchange 2010 Server
    if ($version -eq "Exchange Server 2010")
    {
        # Still to come
    }
}

Write-Verbose "------ All servers completed"


# Output the reports to CSV
if ($Exchange2013Report)
{
}

if ($Exchange2010Report)
{
}


#----------------------------------------
# Finished
#----------------------------------------
