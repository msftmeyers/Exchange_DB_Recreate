<#
.SYNOPSIS
    Script to verify and recreate an empty Exchange database file to free up space in the assigned volume.
    Script will check homemdb attributes, will save some important DB settings, will remove all passive and
    lagged copies and finally, it will re-create EDB and LOG files without re-creating the Database AD object.
        
.PARAMETER Database
    <required> The DBName, which should be re-created

.EXAMPLE
    .\exchange_DBrecreate.ps1 [-Database <DBName>]

.VERSIONS
    V1.0  03.11.2025 - Initial Version
    V1.1  07.11.2025 - Minor changes
    V1.2  10.11.2025 - Minor changes how to add copies and changed the way, isexcludedfromprovisioning will be handled if lagged copies are detected
    
.AUTHOR/COPYRIGHT:
    Steffen Meyer
    Cloud Solution Architect
    Microsoft Deutschland GmbH
#>

[CmdletBinding()]
Param(
     [Parameter(Mandatory=$true,Position=0,HelpMessage='Insert single Database Name')]
     [ValidateNotNullOrEmpty()]
     [String]$Database
     )

$version = "V1.2_10.11.2025"

$now = Get-Date

#Start script
try
{
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path -ErrorAction Stop
}
catch
{
    Write-Host "`nDo not forget to save the script!" -ForegroundColor Red
}

Write-Host "`nScript version: $version"
Write-Host   "Script started: $($now.tostring("dd.MM.yyyy HH:mm:ss"))"

Write-Host "`n---------------------------------------------------------------------------------------" -Foregroundcolor Green
Write-Host   "Script to verify and re-create empty Exchange database files (EDB & LOGs) to free up   " -Foregroundcolor Green
Write-Host   "to free up space in the assigned volume. Script will check homemdb attributes, will    " -Foregroundcolor Green
Write-Host   "save important DB settings, will remove all passive and lagged copies and finally,     " -Foregroundcolor Green
Write-Host   "it will re-create EDB and LOG files WITHOUT re-creating the Database AD object.        " -Foregroundcolor Green
Write-Host   "---------------------------------------------------------------------------------------" -Foregroundcolor Green

#Check if Exchange SnapIn is available and load it
if (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange")
{
    if ((Get-PSSnapin -Registered).name -contains "Microsoft.Exchange.Management.PowerShell.SnapIn")
    {
        Write-Host "`nLoading the Exchange Powershell SnapIn..." -ForegroundColor Yellow
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-ExchangeServer -auto -AllowClobber
    }
    else
    {
        Write-Host "`nExchange Management Tools are not installed. Run the script on a different machine." -ForegroundColor Red
        Return
    }
}

#Detect, where the script is executed
if (!(Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue))
{
    Write-Host "`nATTENTION: Script is executed on a non-Exchangeserver..." -ForegroundColor Cyan
}

Write-Host "`n--------------------"
Write-Host   "PREREQUISITES CHECK:"
Write-Host   "--------------------"
Write-Host "`nWe will now start checking ALL prerequisites before asking you for safely re-creating all files of Database ""$Database""..."

Set-ADServerSettings -ViewEntireForest $true

#Checking Database name
Write-Host "`nCHECK 1: Is database ""$database"" available in this Exchange Organization..." -ForegroundColor Cyan
$DB = Get-MailboxDatabase -Identity $database -Status -ErrorAction SilentlyContinue
if (!($DB))
{
    Write-Host "`nATTENTION: Database ""$database"" cannot be found in this Exchange Organization." -ForegroundColor Red
    Return
}
else
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
}

#AD Lookup for Objects pointing still to homemdb of $database
Write-Host "`nCHECK 2: Are there any enabled mailbox(es) left pointing to ""$database"" (this may take a while)..." -ForegroundColor Cyan
try
{
    Import-Module ActiveDirectory
    $Mailboxes = Get-ADUser -Filter * -Properties homeMDB,msExchArchiveDatabaseLink -ErrorAction Stop | Where-Object {($_.homemdb -like "*$database*" -or $_.msExchArchiveDatabaseLink -like "*$database*") -and $_.samaccountname -notlike "HealthMailbox*"}
}
catch
{
    Write-Host "`nATTENTION: We couldn't get a list of enabled mailboxes/archives still pointing to ""$database"" in ActiveDirectory, please verify and restart script." -ForegroundColor Red
}

if ($Mailboxes)
{
    Write-Host "`nATTENTION: We found still $(($Mailboxes).count) active mailbox(es)/archive(s) (except HealthMailboxes) in ""$database"", please move them first using ""exchange_DBredistribute.ps1"" Script before re-creating EDB and LOG files using this script." -ForegroundColor Red
    Return
}
else
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
}

#Check, if last mailbox removal time is older than mailbox retention on database
Write-Host "`nCHECK 3: We will now check important statistics of ""$database"" (this may take a while)..." -ForegroundColor Cyan
if ($DB.Mounted -eq $True)
{
    try
    {
        $DBStats = Get-MailboxStatistics -Database $database -ErrorAction Stop | where-object disconnectdate
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 5
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't get any statistics of ""$database""." -ForegroundColor Red
        Return
    }
}
else
{
    Write-Host "NOTICE: ""$database"" isn't mounted, we were not able to get all statistics, ARE YOU SURE YOU WANT TO CONTINUE? ( Y / N ): " -ForegroundColor Yellow -NoNewline
    $Dismounted = Read-Host

    if ($Dismounted -ne "Y")
    {
        Write-Host "`nNOTICE: Verify DB state and/or mount ""$database"" manually to get DB statistics and run this script again." -ForegroundColor Yellow
        Return
    }
}

#Detect RecoveryDB
Write-Host "`nCHECK 4: Is database ""$Database"" a RECOVERY database..." -ForegroundColor Cyan

if (!($DB.Recovery -eq $True))
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
    Start-Sleep 5
}
else
{
    Write-Host "`nATTENTION: Database ""$database"" is a RECOVERY database, we do not support re-creating Recovery Database files with this script." -ForegroundColor Red
    Return
}

#Detect and save all copies and copysettings
Write-Host "`nCHECK 5: Is database ""$Database"" a DAG-replicated database with passive and/or lagged copies..." -ForegroundColor Cyan
try
{
    $DBCopies = Get-MailboxDatabaseCopyStatus $database -ErrorAction Stop | Where-Object ActiveCopy -ne $True
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green

    if ($DBCopies)
    {
        Write-Host "`nFor documentation, a table of all configured Database Copies for Database ""$database"":" -ForegroundColor Green

        Get-MailboxDatabaseCopyStatus $database | format-table Name,Status,ActivationPreference,@{n="ReplayLagTime (days)";e={$_.replaylagstatus.configuredlagtime.totaldays}}
        Start-Sleep 10
    }
}
catch
{
    Write-Host "`nATTENTION: We couldn't get a list of additional database copies of database ""$database""." -ForegroundColor Red
    Return
}

#Is there any lagged copy? If yes, what is maximum replay lag time?
Write-Host "`nCHECK 6: Are there any lagged copies configured and what is the maximum lag time of database ""$Database""..." -ForegroundColor Cyan
if ($DBCopies)
{
    $lagtime=@()

    foreach ($DBCopy in $DBCopies)
    {
        if ($DBCopy.replaylagstatus.enabled -eq $True)
        {
            $lagtime += New-Object -type PSObject -Prop @{MailboxServer=$DBCopy.MailboxServer;Lagtime=$DBCopy.ReplayLagStatus.ConfiguredLagTime.Days}
        }
    }
    if ($lagtime)
    {
        $maxlag = $lagtime | Sort-Object -Descending Lagtime | Select-Object -First 1
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 5
    }
    else
    {
        Write-Host "...SUCCESSFUL! (There are NO LAGGED copies configured for database ""$database"")." -ForegroundColor Green
        Start-Sleep 5
    }
}
else
{
    Write-Host "...SUCCESSFUL! (There are NO DATABASE COPIES configured for database ""$database"")." -ForegroundColor Green
    Start-Sleep 5
}

#Find out, what is higher, mailboxretention or maximum lagtime
if ($maxlag.Lagtime -ge $DB.MailboxRetention.TotalDays)
{
    $timetowait = $maxlag.Lagtime
}
else
{
    $timetowait = $DB.MailboxRetention.TotalDays
}

#If youngest disconnectdate is greater then mailboxretention days back from today, it is not safe to delete edb file because of mailbox reconnect/recovery purposes
Write-Host "`nCHECK 7: Is the mailbox retention or maximum lag time passed after the last mailbox was moved out of this database..." -ForegroundColor Cyan

if (($DBStats | sort -Descending disconnectdate | select -first 1).disconnectdate -ge $now.AddDays(-$timetowait))
{
    Write-Host "`nATTENTION: We couldn't find any ""classic"" backup of database ""$database"" and we found the last mailbox" -ForegroundColor Red
    Write-Host   "disconnectdate $((($DBStats | sort -Descending disconnectdate | select -first 1).disconnectdate).tostring("dd.MM.yyyy")), which is not older than the minimum DB retention time of $timetowait days back from today." -ForegroundColor Red
    Write-Host "`nTo have still possibilities to restore or reconnect mailboxes, we recommend you to wait with re-creation  " -ForegroundColor Red
    Write-Host   "of EDB and LOG files of database ""$database"" at least until $($([datetime]($DBStats | sort -Descending disconnectdate | select -first 1).disconnectdate.AddDays($timetowait + 1 )).ToString("dd.MMMM yyyy"))." -ForegroundColor Red
   
    #But you can bypass this check if you want to and continue with the re-creation of EDB and LOG files
    Write-Host "`nDo you want me to continue with the prerequisites check? ( Y / N ): " -ForegroundColor Yellow -NoNewline
    $ForceRet = Read-Host

    if ($ForceRet -ne "Y")
    {
        Write-Host "`nNOTICE: The script hasn't changed anything. EDB and LOG files of Database ""$database"" were not re-created." -ForegroundColor Yellow
        Return
    }
    else
    {
        Write-Host "...CONTINUING..." -ForegroundColor Green
        Start-Sleep 5
    }
}
else
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
    Start-Sleep 5
}

#End of checking prerequisites
Write-Host "`nALL PREREQUISITES ARE FULLFILLED, we can now continue with re-creating tasks..." -ForegroundColor Green

#Start with Re-Creation Tasks?
Write-Host "`nDo you want to RE-CREATE EDB and LOG files for Database ""$database"" now? ( Y / N ): " -ForegroundColor Yellow -NoNewline
$Cont = Read-Host

If ($Cont -eq "Y")
{
    Write-Host "`n------------------"
    Write-Host   "RE-CREATION TASKS:"
    Write-Host   "------------------"

    #Starting Re-Creation Tasks
    Write-Host "`nWe will now start creating empty EDB and LOG files and starting a new transaction log file sequence.                     "
    Write-Host   "The AD object of the Database will not be re-created, but all DB copies and copy parameters will be re-established again."

    #Disable CircularLogging before removing copies
    Write-Host "`nTASK 1: DISABLE Circular Logging for Database ""$database""..." -ForegroundColor Cyan
    if ($DB.CircularLoggingEnabled -eq $True)
    {
        try
        {
            Get-MailboxDatabase $Database | Set-MailboxDatabase -CircularLoggingEnabled $false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Host "...SUCCESSFUL, waiting 30 seconds for AD replication..." -ForegroundColor Green
            Start-Sleep 30
        }
        catch
        {
            Write-Host "`nATTENTION: We couldn't disable Circular Logging for ""$database"", re-run the script." -ForegroundColor Red
            Return
        }
    }
    else
    {
        Write-Host "Circular Logging for database ""$database"" is NOT ENABLED." -ForegroundColor Green
        Start-Sleep 5
    }

    #Remove Copies
    Write-Host "`nTASK 2: REMOVING all passive and lagged copies..." -ForegroundColor Cyan
    if ($DBCopies)
    {
        foreach ($DBCopy in $DBCopies)
        {
            try
            {
                Get-MailboxDatabaseCopyStatus $DBCopy.Name | Remove-MailboxDatabaseCopy -Confirm:$False -ErrorAction Stop -WarningAction SilentlyContinue
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't remove all passive or lagged copies of ""$database"", remove copies manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
                Return
            }
        }
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 5
    }
    else
    {
        Write-Host "...There are no copies for database ""$database"" configured." -ForegroundColor Green
        Start-Sleep 5
    }

    #Dismount of Database
    Write-Host "`nTASK 3: DISMOUNTING Database ""$database""..." -ForegroundColor Cyan
    try
    {
        Get-MailboxDatabase $Database | Dismount-Database -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 5
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't DISMOUNT database ""$database"", dismount manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
        Return
    }

    #Deleting EDB and LOG folder content on server with active copy
    Write-Host "`nTASK 4: REMOVING old EDB and LOG folder content on server ""$(($DB).MountedOnServer)""..." -ForegroundColor Cyan
    try
    {
        $Server = ($DB).MountedOnServer
        $EDBFolder = Split-Path (Get-MailboxDatabase $DB).EdbFilePath -Parent
        $LOGFolder = (Get-MailboxDatabase $DB).LogFolderPath.Pathname
    
        Invoke-Command -ComputerName $Server -ScriptBlock {
        param ($EDBFolder, $LOGFolder)
        Remove-Item "$EDBFolder\*" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$LOGFolder\*" -Recurse -Force -ErrorAction SilentlyContinue
        } -ArgumentList $EDBFolder, $LOGFolder -ErrorAction Stop -WarningAction SilentlyContinue

        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 5
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't DELETE all EDB and LOG files of database ""$database"" on server ""$(($DB).MountedOnServer)"", delete all files manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
        Return
    }

    #Forcibly Mount Database and create new EDB and LOG files
    Write-Host "`nTASK 5: MOUNTING Database ""$database"" and create new EDB and LOG files (this may take a while)..." -ForegroundColor Cyan
    try
    {
        Get-MailboxDatabase $Database | Mount-Database -Force -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 5
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't MOUNT database ""$database"", use MOUNT-DATABASE -FORCE manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
        Return
    }

    #Creating copies
    Write-Host "`nTASK 6: ADDING, SUSPENDING and SEEDING Database copies for database ""$database"" on the same servers like before, with same activation preference. If it was a lagged one, the REPLAYLAGTIME will be added back again..." -ForegroundColor Cyan
    if ($DBCopies)
    {
        $DBCopies = $DBCopies | Sort-Object activationpreference
    
        $CopyCount = 1

        foreach ($DBCopy in $DBCopies)
        {
            $CopyCount++

            try
            {
                Write-host "`nADDING DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                Get-MailboxDatabase $database | Add-MailboxDatabaseCopy -MailboxServer $DBCopy.MailboxServer -ActivationPreference $DBCopy.ActivationPreference -ReplayLagTime $DBCopy.ReplayLagStatus.ConfiguredLagTime -WarningAction SilentlyContinue -ErrorAction Stop
                Write-Host "...SUCCESSFUL, DBCopy #$CopyCount was added, waiting 30 seconds for AD replication." -ForegroundColor Green
                Start-Sleep 30
            
                Write-host "SUSPENDING DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                Suspend-MailboxDatabaseCopy $DBCopy.Name -WarningAction SilentlyContinue -ErrorAction Stop
                Write-Host "...SUCCESSFUL, DBCopy #$CopyCount was suspended." -ForegroundColor Green
            
                Write-host "SEEDING DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                Get-MailboxDatabaseCopyStatus $DBCopy.Name | Update-MailboxDatabaseCopy -DeleteExistingFiles -Confirm:$false -Force -WarningAction SilentlyContinue -ErrorAction Stop
                Write-Host "...SUCCESSFUL, DBCopy #$CopyCount was seeded." -ForegroundColor Green

                #Suspend  LaggedCopy with -ActivationOnly
                if ($DBCopy.ReplayLagStatus.Enabled -eq $True)
                {
                    Write-host "SUSPEND (-ActivationOnly) LAGGED DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                    Get-MailboxDatabaseCopyStatus $DBCopy.Name | Suspend-MailboxDatabaseCopy -ActivationOnly -WarningAction SilentlyContinue -ErrorAction Stop
                    Write-Host "...SUCCESSFUL, Lagged DBCopy #$CopyCount was suspended (-ActivationOnly)." -ForegroundColor Green    
                }
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't ADD, SUSPEND or SEED DBCopy #$CopyCount of database ""$database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
            }
        }
        Write-Host "`nNOTICE: List of configured Database Copies for database ""$database"":" -ForegroundColor Green
                
        #Output DBCopy table
        Get-MailboxDatabaseCopyStatus $database | format-table Name,Status,ActivationPreference,@{n="ReplayLagTime (days)";e={$_.replaylagstatus.configuredlagtime.totaldays}}
        Start-Sleep 10
    }
    else
    {
        Write-Host "There were NO ADDITIONAL DATABASE COPIES configured before, so we didn't add anything back here." -ForegroundColor Green
        Start-Sleep 5
    }

    #Enabling CircularLogging
    Write-Host "`nTASK 7: IF Circular Logging was active before, it will be enabled again..." -ForegroundColor Cyan
    if ($DB.CircularLoggingEnabled -eq $True)
    {
        try
        {
            Get-MailboxDatabase $database | Set-MailboxDatabase -CircularLoggingEnabled $True -WarningAction SilentlyContinue -ErrorAction Stop
            Write-Host "...SUCCESSFUL!" -ForegroundColor Green
            Start-Sleep 5
        }
        catch
        {
            Write-Host "`nATTENTION: We couldn't ENABLE Circular Logging for database ""$database"", enable it manually by using ""Set-Mailboxdatabase $database -CircularLoggingEnabled $True""" -ForegroundColor Red
        }
    }
    else
    {
        Write-Host "...Circular Logging for database ""$database"" was NOT CONFIGURED before." -ForegroundColor Green
        Start-Sleep 5
    }

   #Disable IsExcludeFromProvisioning, if it was active before 
    Write-Host "`nTASK 8: IF Database ""$database"" was excluded from Mailbox provisioning, it will be included again..." -ForegroundColor Cyan
    if ($DB.IsExcludedFromProvisioning -eq "True")
    {
        if (!($DBCopies.ReplayLagStatus.Enabled -contains "True"))
        {
            try
            {
                $IsExcluded = (Get-MailboxDatabase $database).distinguishedname | Set-ADObject -Replace @{msExchProvisioningFlags=1} -WarningAction SilentlyContinue -ErrorAction Stop
                Write-Host "...SUCCESSFUL!" -ForegroundColor Green
                Start-Sleep 5
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't ENABLE Mailbox Provisioning for database ""$database""." -ForegroundColor Red
            }
        }
        else
        {
            Write-Host "`nNOTICE: We detected at least ONE LAGGED COPY with a REPLAYLAGTIME of $($maxlag.lagtime) days. You" -ForegroundColor Yellow
            Write-Host   "should wait at least $($maxlag.lagtime) days before moving Mailboxes to this database to ensure SLA compliance." -ForegroundColor Yellow
            Write-Host "`nWe didn't include Database ""$database"" back into Exchange Mailbox provisioning, you need to do this MANUALLY in $($maxlag.lagtime) days!" -ForegroundColor Red
            Write-Host   "Use: ""Set-MailboxDatabase $($database) -IsExcludedFromProvisioning `$false"""
        }
    }
        
    #Final statement
    Write-Host "`nRESULT: You successfully created new and empty EDB and LOG files for Database ""$database"", including new copies and all settings like before, but without a new Database AD object, well done!" -ForegroundColor Green
    
}
else
{
    Write-Host "`nNOTICE: The script hasn't changed anything. EDB and LOG files of Database ""$database"" were not re-created." -ForegroundColor Yellow
}
#END