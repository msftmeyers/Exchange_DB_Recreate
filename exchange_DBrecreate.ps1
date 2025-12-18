<#
.SYNOPSIS
    Script to verify and re-create EDB file and a new transaction LOG stream of an Exchange Mailbox Database
    with no mailboxes left. This will free up space in the assigned volume immediately.
    
.DESCRIPTION
    The script consists of seven checks and eight tasks. It will start checking all prerequisites which needs to 
    be fulfilled to re-create EDB and LOG files and only if ALL checks are successful, it will ask you to continue
    with the tasks of re-creating EDB and LOG files.

    This script is NOT USING Remove-MailboxDatabase cmdlet, it will just re-create files, so all important DB settings will
    be preserved and the databases will not be moved to the end of the Get-MailboxDatabase list.
    
    Script will check homemdb attributes, will save important DB settings, will remove all passive and
    lagged copies (if DB is part of a DAG) and finally, it will re-create EDB and LOG files without re-creating
    the Mailbox Database AD object.
        
.PARAMETER Database
    <required> The DBName, for which the EDB and LOG files should be re-created

.EXAMPLE
    .\exchange_DBrecreate.ps1 [-Database <DBName>]

.NOTES
    Steffen Meyer
    Cloud Solution Architect
    Microsoft Deutschland GmbH

    V1.0  03.11.2025 - Initial Version
    V1.1  07.11.2025 - Minor changes
    V1.2  10.11.2025 - Minor changes how to add copies and changed the way, isexcludedfromprovisioning will be handled if lagged copies are detected
    V1.3  03.12.2025 - minor changes, description added
    V1.4  10.12.2025 - Changed DB Filter
    V2.0  18.12.2025 - Changed the order to add, suspend and seed DB copies. Due to different AD replication delays, I've added a function to wait for AD replication
#>

[CmdletBinding()]
Param(
     [Parameter(Mandatory=$true,Position=0,HelpMessage='Insert single Database Name')]
     [ValidateNotNullOrEmpty()]
     [String]$Database
     )

$version = "V2.0_18.12.2025"

$now = Get-Date

#Function for a delay until AD replication is successful
function Wait-Until
{
    param(
        [scriptblock]$Condition,
        [int]$TimeoutSeconds = 300,
        [int]$SleepSeconds = 15
    )

    $End = (Get-Date).AddSeconds($TimeoutSeconds)

    while ((Get-Date) -lt $End)
    {
        if (& $Condition) {return $true }
        Write-Host "...waiting 15 seconds for AD replication..."
        Start-Sleep -Seconds $SleepSeconds
    }

    return $false
}

#START SCRIPT
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
Write-Host   "space in the assigned volume. Script will check homemdb attributes, it will save       " -Foregroundcolor Green
Write-Host   "all important DB settings, will remove all passive and lagged copies and finally,      " -Foregroundcolor Green
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

Write-Host "`n------------------------"
Write-Host   "PREREQUISITE CHECKS (7):"
Write-Host   "------------------------"
Write-Host "`nWe will now start checking ALL prerequisites before asking you for safely re-creating all files of Database ""$Database""..."

Set-ADServerSettings -ViewEntireForest $true

#Checking Database name
Write-Host "`nCHECK 1 of 7: Is database ""$Database"" available in this Exchange Organization..." -ForegroundColor Cyan
$DB = Get-MailboxDatabase -Identity $Database -Status -ErrorAction SilentlyContinue
if (!($DB))
{
    Write-Host "`nATTENTION: Database ""$Database"" cannot be found in this Exchange Organization." -ForegroundColor Red
    Return
}
else
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
}

#AD Lookup for Objects pointing still to homemdb of $Database
Write-Host "`nCHECK 2 of 7: Are there any types of enabled mailboxes left pointing to ""$Database"" (this may take a while)..." -ForegroundColor Cyan
try
{
    Import-Module ActiveDirectory
    $Mailboxes = Get-ADUser -Filter * -Properties homeMDB,msExchArchiveDatabaseLink -ErrorAction Stop | Where-Object {($_.homemdb -eq $DB.distinguishedname -or $_.msExchArchiveDatabaseLink -like "*$Database*") -and $_.samaccountname -notlike "HealthMailbox*"}
}
catch
{
    Write-Host "`nATTENTION: We couldn't get a list of enabled mailboxes/archives still pointing to ""$Database"" in ActiveDirectory, please verify and restart script." -ForegroundColor Red
}

if ($Mailboxes)
{
    Write-Host "`nATTENTION: We found still $(($Mailboxes | Measure-Object).count) active mailbox(es)/archive(s) (except HealthMailboxes) in ""$Database"", please move them first using ""exchange_DBredistribute.ps1"" Script before re-creating EDB and LOG files using this script." -ForegroundColor Red
    Return
}
else
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
}

#Check, if last mailbox removal time is older than mailbox retention on database
Write-Host "`nCHECK 3 of 7: We will now check important statistics of ""$Database"" (this may take a while)..." -ForegroundColor Cyan
if ($DB.Mounted -eq $True)
{
    try
    {
        $DBStats = Get-MailboxStatistics -Database $Database -ErrorAction Stop | where-object disconnectdate
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 2
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't get any statistics of ""$Database""." -ForegroundColor Red
        Return
    }
}
else
{
    Write-Host "NOTICE: ""$Database"" isn't mounted, we were not able to get all statistics, ARE YOU SURE YOU WANT TO CONTINUE? ( Y / N ): " -ForegroundColor Yellow -NoNewline
    $Dismounted = Read-Host

    if ($Dismounted -ne "Y")
    {
        Write-Host "`nNOTICE: Verify DB state and/or mount ""$Database"" manually to get DB statistics and run this script again." -ForegroundColor Yellow
        Return
    }
}

#Detect RecoveryDB
Write-Host "`nCHECK 4 of 7: Database ""$Database"" is NOT a RECOVERY database..." -ForegroundColor Cyan

if (!($DB.Recovery -eq $True))
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
    Start-Sleep 2
}
else
{
    Write-Host "`nATTENTION: Database ""$Database"" is a RECOVERY database, we do not support re-creating Recovery Database files with this script." -ForegroundColor Red
    Return
}

#Detect and save all DB copies and DB copy settings
Write-Host "`nCHECK 5 of 7: Is database ""$Database"" a DAG-replicated database with passive and/or lagged copies..." -ForegroundColor Cyan
try
{
    $DBCopies = Get-MailboxDatabaseCopyStatus $Database -ErrorAction Stop | Where-Object ActiveCopy -ne $True
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green

    if ($DBCopies)
    {
        Write-Host "`nFor documentation, a table of all configured Database Copies for Database ""$Database"":"
        
        $DBCopyLayout = Get-MailboxDatabaseCopyStatus $Database
        #Console output
        $DBCopyLayout | Format-Table Name,Status,ActivationPreference,@{n="ReplayLagTime (days)";e={$_.replaylagstatus.configuredlagtime.totaldays}},ActivationSuspended
        #Export to CSV
        try
        {
            $DBCopyLayout | Select-Object Name,Status,ActivationPreference,@{n="ReplayLagTime (days)";e={$_.replaylagstatus.configuredlagtime.totaldays}},ActivationSuspended | Export-Csv "$ScriptPath\DBCopyLayout_$($Database)_before.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
            Write-Host "NOTICE: You will also find all configured copies and parameters in ""$ScriptPath\DBCopyLayout_$($Database)_before.csv."""
        }
        catch
        {
            Write-Host "`nATTENTION: We couldn't export all configured copies and parameters to ""$ScriptPath\DBCopyLayout_$($Database)_before.csv.""" -ForegroundColor Red
        }
        Start-Sleep 5
    }
}
catch
{
    Write-Host "`nATTENTION: We couldn't get a list of additional database copies of database ""$Database""." -ForegroundColor Red
    Return
}

#Is there any lagged copy? If yes, what is maximum replay lag time?
Write-Host "`nCHECK 6 of 7: Are there any lagged copies configured and what is the maximum lag time of database ""$Database""..." -ForegroundColor Cyan
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
        Start-Sleep 2
    }
    else
    {
        Write-Host "...SUCCESSFUL! (There are NO LAGGED copies configured for database ""$Database"")." -ForegroundColor Green
        Start-Sleep 2
    }
}
else
{
    Write-Host "...SUCCESSFUL! (There are NO DATABASE COPIES configured for database ""$Database"")." -ForegroundColor Green
    Start-Sleep 2
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

#If youngest disconnectdate is greater then mailboxretention days back from today, it is not safe to delete EDB file because of mailbox reconnect/recovery purposes
Write-Host "`nCHECK 7 of 7: Is the mailbox retention or maximum lag time passed after the last mailbox was moved out of this database..." -ForegroundColor Cyan

if (($DBStats | sort -Descending disconnectdate | select -first 1).disconnectdate -ge $now.AddDays(-$timetowait) -and (!($DB.LastFullBackup) -or !($DB.LastCopyBackup)))
{
    Write-Host "`nATTENTION: We couldn't find any ""classic"" backup of database ""$Database"" and we found the last mailbox" -ForegroundColor Red
    Write-Host   "disconnectdate $((($DBStats | sort -Descending disconnectdate | select -first 1).disconnectdate).tostring("dd.MM.yyyy")), which is not older than the minimum DB retention time of $timetowait days back from today." -ForegroundColor Red
    Write-Host "`nTo have still possibilities to restore or reconnect mailboxes, we recommend you to wait with re-creation  " -ForegroundColor Red
    Write-Host   "of EDB and LOG files of database ""$Database"" at least until $($([datetime]($DBStats | sort -Descending disconnectdate | select -first 1).disconnectdate.AddDays($timetowait + 1 )).ToString("dd.MMMM yyyy"))." -ForegroundColor Red
   
    #But you can bypass this check if you want to and continue with the re-creation of EDB and LOG files
    Write-Host "`nDo you want me to continue with the prerequisites check? ( Y / N ): " -ForegroundColor Yellow -NoNewline
    $ForceRet = Read-Host

    if ($ForceRet -ne "Y")
    {
        Write-Host "`nNOTICE: The script hasn't changed anything. EDB and LOG files of Database ""$Database"" were not re-created." -ForegroundColor Yellow
        Return
    }
    else
    {
        Write-Host "...CONTINUING..." -ForegroundColor Green
        Start-Sleep 2
    }
}
else
{
    Write-Host "...SUCCESSFUL!" -ForegroundColor Green
    Start-Sleep 2
}

#End of checking prerequisites
Write-Host "`nALL PREREQUISITES ARE FULFILLED, we can now continue with re-creation tasks..." -ForegroundColor Green

#Start with Re-Creation Tasks?
Write-Host "`nDo you want to RE-CREATE EDB and LOG files for Database ""$Database"" now? ( Y / N ): " -ForegroundColor Yellow -NoNewline
$Cont = Read-Host

If ($Cont -eq "Y")
{
    Write-Host "`n-----------------------"
    Write-Host   "RE-CREATION TASKS (11):"
    Write-Host   "-----------------------"

    #Starting Re-Creation Tasks
    Write-Host "`nWe will now start creating empty EDB and LOG files and starting a new transaction log file sequence.                     "
    Write-Host   "The AD object of the Database will NOT be re-created, but all DB copies and parameters will be re-established again."

    #Disable CircularLogging before removing copies
    Write-Host "`nTASK 1 of 11: DISABLE Circular Logging for Database ""$Database""..." -ForegroundColor Cyan
    if ($DB.CircularLoggingEnabled -eq $True)
    {
        try
        {
            $circ = Get-MailboxDatabase $Database | Set-MailboxDatabase -CircularLoggingEnabled $false -ErrorAction Stop -WarningAction SilentlyContinue

            $Result = Wait-Until {
            (Get-MailboxDatabase $Database).circularloggingenabled -eq $False}

            if (!($Result))
            {
                Write-Host "`nATTENTION: Within the last 5 minutes, we couldn't find a replicated, disabled CircularLogging setting of database ""$Database"", please verify." -ForegroundColor Red
                Return
            }
            else
            {
                Write-Host "...SUCCESSFUL!" -ForegroundColor Green
                Start-Sleep 2
            }
        }
        catch
        {
            Write-Host "`nATTENTION: We couldn't disable CircularLogging for ""$Database"", re-run the script." -ForegroundColor Red
            Return
        }
    }
    else
    {
        Write-Host "...Circular Logging for database ""$Database"" is NOT ENABLED." -ForegroundColor Green
        Start-Sleep 2
    }

    #Remove Copies
    Write-Host "`nTASK 2 of 11: REMOVING all passive and lagged copies..." -ForegroundColor Cyan
    if ($DBCopies)
    {
        foreach ($DBCopy in $DBCopies)
        {
            try
            {
                $Remove = Get-MailboxDatabaseCopyStatus $DBCopy.Name | Remove-MailboxDatabaseCopy -Confirm:$False -ErrorAction Stop -WarningAction SilentlyContinue
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't remove all passive or lagged copies of ""$Database"", remove copies manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
                Return
            }
        }
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 2
    }
    else
    {
        Write-Host "...There are NO DATABASE COPIES for database ""$Database"" configured." -ForegroundColor Green
        Start-Sleep 2
    }

    #Dismount of Database
    Write-Host "`nTASK 3 of 11: DISMOUNTING Database ""$Database""..." -ForegroundColor Cyan
    try
    {
        $Dismount = Get-MailboxDatabase $Database | Dismount-Database -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 2
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't DISMOUNT database ""$Database"", dismount manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
        Return
    }

    #Deleting EDB and LOG folder content on server with active copy
    Write-Host "`nTASK 4 of 11: REMOVING old EDB and LOG folder content on server ""$(($DB).MountedOnServer)""..." -ForegroundColor Cyan
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
        Start-Sleep 2
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't DELETE all EDB and LOG files of database ""$Database"" on server ""$(($DB).MountedOnServer)"", delete all files manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
        Return
    }

    #Forcibly Mount Database and create new EDB and LOG files
    Write-Host "`nTASK 5 of 11: MOUNTING Database ""$Database"" and create new EDB and LOG files (this may take a while)..." -ForegroundColor Cyan
    try
    {
        $Mount = Get-MailboxDatabase $Database | Mount-Database -Force -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "...SUCCESSFUL!" -ForegroundColor Green
        Start-Sleep 2
    }
    catch
    {
        Write-Host "`nATTENTION: We couldn't MOUNT database ""$Database"", use MOUNT-DATABASE -FORCE manually, re-run the script and take care of all copies to be re-created afterwards." -ForegroundColor Red
        Return
    }

    #Creating copies
    Write-Host "`nTASK 6 of 11: ADDING Database copies for database ""$Database""..." -ForegroundColor Cyan
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
                $add = Get-MailboxDatabase $Database | Add-MailboxDatabaseCopy -MailboxServer $DBCopy.MailboxServer -ActivationPreference $DBCopy.ActivationPreference -ReplayLagTime $DBCopy.ReplayLagStatus.ConfiguredLagTime -SeedingPostponed -WarningAction SilentlyContinue -ErrorAction Stop
                
                $Result = Wait-Until {
                (Get-MailboxDatabaseCopyStatus $($DBCopy).Name) -and ((Get-MailboxDatabaseCopyStatus $($DBCopy).Name ).status -notin "Unknown" )}

                if (!($Result))
                {
                    Write-Host "`nATTENTION: Within the last 5 minutes, we couldn't add and/or replicate DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
                }
                else
                {
                    Write-Host "...SUCCESSFUL, DBCopy #$CopyCount was added." -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't ADD DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
            }
            Start-Sleep 2
        }
    }
    else
    {
        Write-Host "...There were NO ADDITIONAL DATABASE COPIES configured before, so we didn't need to add anything here." -ForegroundColor Green
        Start-Sleep 2
    }
        
    #Suspending copies
    Write-Host "`nTASK 7 of 11: SUSPENDING Database copies for database ""$Database""..." -ForegroundColor Cyan
    if ($DBCopies)
    {
        $DBCopies = $DBCopies | Sort-Object activationpreference
    
        $CopyCount = 1

        foreach ($DBCopy in $DBCopies)
        {
            $CopyCount++

            try
            {
                Write-host "`nSUSPENDING DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                $suspend = Suspend-MailboxDatabaseCopy $DBCopy.Name -WarningAction SilentlyContinue -ErrorAction Stop
                
                $Result = Wait-Until {
                (Get-MailboxDatabaseCopyStatus $($DBCopy).Name).status -notin "Failed","Unknown"}

                if (!($Result))
                {
                    Write-Host "`nATTENTION: Within the last 5 minutes, we couldn't add and/or replicate DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
                }
                else
                {
                    Write-Host "...SUCCESSFUL, DBCopy #$CopyCount was suspended." -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't SUSPEND DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
            }
            Start-Sleep 2
        }
    }
    else
    {
        Write-Host "...There were NO ADDITIONAL DATABASE COPIES configured before, so we didn't need to suspend anything here." -ForegroundColor Green
        Start-Sleep 2
    }

    #Seeding copies
    Write-Host "`nTASK 8 of 11: SEEDING Database copies for database ""$Database""..." -ForegroundColor Cyan
    if ($DBCopies)
    {
        $DBCopies = $DBCopies | Sort-Object activationpreference
    
        $CopyCount = 1

        foreach ($DBCopy in $DBCopies)
        {
            $CopyCount++

            try
            {
                Write-host "`nSEEDING DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                $seed = Get-MailboxDatabaseCopyStatus $DBCopy.Name | Update-MailboxDatabaseCopy -DeleteExistingFiles -Confirm:$false -Force -WarningAction SilentlyContinue -ErrorAction Stop
                
                $Result = Wait-Until {
                (Get-MailboxDatabaseCopyStatus $($DBCopy).Name).status -in "Healthy"}

                if (!($Result))
                {
                    Write-Host "`nATTENTION: Within the last 5 minutes, we couldn't add and/or replicate DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
                }
                else
                {
                    Write-Host "...SUCCESSFUL, DBCopy #$CopyCount was seeded." -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't SEED DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
            }
            Start-Sleep 2
        }
    }
    else
    {
        Write-Host "...There were NO ADDITIONAL DATABASE COPIES configured before, so we didn't need to seed anything here." -ForegroundColor Green
        Start-Sleep 2
    }

    #Suspend LaggedCopy with -ActivationOnly
    Write-Host "`nTASK 9 of 11: SUSPENDING lagged Database copies (using -ActivationOnly) for database ""$Database""..." -ForegroundColor Cyan
    if ($DBCopies)
    {
        $DBCopies = $DBCopies | Sort-Object activationpreference
    
        $CopyCount = 1

        foreach ($DBCopy in $DBCopies)
        {
            $CopyCount++

            try
            {
                if ($DBCopy.ReplayLagStatus.Enabled -eq $True)
                {
                    Write-host "`nSUSPEND (-ActivationOnly) LAGGED DBCopy #$CopyCount (""$($DBCopy.Name)"")..."
                    $suspend = Suspend-MailboxDatabaseCopy $DBCopy.Name -ActivationOnly -WarningAction SilentlyContinue -ErrorAction Stop
                    Write-Host "...SUCCESSFUL, Lagged DBCopy #$CopyCount was suspended (-ActivationOnly)." -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't SUSPEND (-ActivationOnly) DBCopy #$CopyCount of database ""$Database"" on Mailboxserver ""$($DBCopy.Mailboxserver)"", please verify." -ForegroundColor Red
            }
            Start-Sleep 2
        }
        Write-Host "`nNOTICE: Final table of all configured Database Copies for database ""$Database"":"
        
        $DBCopyLayout = Get-MailboxDatabaseCopyStatus $Database
        #Console output
        $DBCopyLayout | Format-Table Name,Status,ActivationPreference,@{n="ReplayLagTime (days)";e={$_.replaylagstatus.configuredlagtime.totaldays}},ActivationSuspended
        #Export to CSV
        try
        {        
            $DBCopyLayout | Select-Object Name,Status,ActivationPreference,@{n="ReplayLagTime (days)";e={$_.replaylagstatus.configuredlagtime.totaldays}},ActivationSuspended | Export-Csv "$ScriptPath\DBCopyLayout_$($Database)_after.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation
            Write-Host "NOTICE: You will also find all configured copies and parameters after re-creation in ""$ScriptPath\DBCopyLayout_$($Database)_after.csv."""
        }
        catch
        {
            Write-Host "`nATTENTION: We couldn't export all configured copies and parameters after re-creation to ""$ScriptPath\DBCopyLayout_$($Database)_after.csv.""" -ForegroundColor Red 
        }
        Start-Sleep 5
    }
    else
    {
        Write-Host "...There were NO ADDITIONAL DATABASE COPIES configured before, so we didn't need to suspend any lagged copies here." -ForegroundColor Green
        Start-Sleep 2
    }

    #Enabling CircularLogging
    Write-Host "`nTASK 10 of 11: IF Circular Logging was active before, it will be enabled again..." -ForegroundColor Cyan
    if ($DB.CircularLoggingEnabled -eq $True)
    {
        try
        {
            $circ = Get-MailboxDatabase $Database | Set-MailboxDatabase -CircularLoggingEnabled $True -WarningAction SilentlyContinue -ErrorAction Stop
            Write-Host "...SUCCESSFUL!" -ForegroundColor Green
            Start-Sleep 2
        }
        catch
        {
            Write-Host "`nATTENTION: We couldn't ENABLE Circular Logging for database ""$Database"", enable it manually by using ""Set-Mailboxdatabase $Database -CircularLoggingEnabled $True""" -ForegroundColor Red
        }
    }
    else
    {
        Write-Host "...Circular Logging for database ""$Database"" was NOT ACTIVE before re-creation." -ForegroundColor Green
        Start-Sleep 2
    }

   #Disable IsExcludedFromProvisioning, if it was active before 
    Write-Host "`nTASK 11 of 11: IF Database ""$Database"" was excluded from Mailbox provisioning, it will be included again..." -ForegroundColor Cyan
    if ($DB.IsExcludedFromProvisioning -eq "True")
    {
        if (!($DBCopies.ReplayLagStatus.Enabled -contains "True"))
        {
            try
            {
                $IsExcluded = (Get-MailboxDatabase $Database).distinguishedname | Set-ADObject -Replace @{msExchProvisioningFlags=1} -WarningAction SilentlyContinue -ErrorAction Stop
                Write-Host "...SUCCESSFUL!" -ForegroundColor Green
                Start-Sleep 2
            }
            catch
            {
                Write-Host "`nATTENTION: We couldn't ENABLE Mailbox Provisioning for database ""$Database""." -ForegroundColor Red
                Start-Sleep 2
            }
        }
        else
        {
            Write-Host "`nNOTICE: We detected at least ONE LAGGED COPY with a REPLAYLAGTIME of $($maxlag.lagtime) days. You" -ForegroundColor Yellow
            Write-Host   "should wait at least $($maxlag.lagtime) days before moving Mailboxes to this database to ensure SLA compliance." -ForegroundColor Yellow
            Write-Host "`nWe didn't include Database ""$Database"" back into Exchange Mailbox provisioning, you need to do this MANUALLY in $($maxlag.lagtime) days!" -ForegroundColor Red
            Write-Host   "Use: ""Set-MailboxDatabase $($Database) -IsExcludedFromProvisioning `$false"""
        }
    }
    else
    {
        Write-Host "...Setting ""IsExcludedFromProvisioning"" of ""$Database"" wasn't TRUE before re-creation." -ForegroundColor Green
        Start-Sleep 2
    }
        
    #Final statement
    Write-Host "`nRESULT: Re-creation of an empty EDB and new LOG files for Database ""$Database"" successfully done, including new DB copies and all settings like before, but without re-creating the Database AD object." -ForegroundColor Green
}
else
{
    Write-Host "`nNOTICE: The script hasn't changed anything. EDB and LOG files of Database ""$Database"" were not re-created." -ForegroundColor Yellow
}
#END SCRIPT