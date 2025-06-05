
## NUT server ENUM : https://github.com/networkupstools/nut/blob/master/clients/status.h

function Send-NUTCommand {
    param (
        [string]$Command
    )

    $NutResponse = @()
    $NutMoreData = $true
    $NutMultiLine = $false

    $client = New-Object System.Net.Sockets.TcpClient

    try
    {
        $client.Connect($config.NUT.NUTserver, $config.NUT.NUTPort)
        $stream = $client.GetStream()
        $writer = New-Object System.IO.StreamWriter($stream)
        $reader = New-Object System.IO.StreamReader($stream)

        $writer.WriteLine($Command)
        $writer.Flush()

        while ($NutMoreData)
        {
            $NutResponseLine = $reader.ReadLine()

            ## Multil row
            if ($NutResponseLine -match "^BEGIN")
            {
                $NutMultiLine = $true
                $NutResponse += $NutResponseLine

            }
            elseif (($NutResponseLine -notmatch "^END") -and ($NutMultiLine ))
            {
                $NutResponse += $NutResponseLine
            }
            else
            {
                #OK, END, ERR are expected here.
                # anything else is also cauth just in case to make sure we don't keep going
                $NutResponse += $NutResponseLine
                $NutMoreData = $false
            }
        }

        return $NutResponse 
    }
    catch 
    {
        Write-Output "Error: $_"
        return "TCPERR"
    }
    finally {
        # Close the connection
        if ($client) { $client.Close() }
    }

}


function Init-NUTConnection
{
    param (
        [string]$TargetUPS
    )
    $UPSFound = $false

    $NUTResponse = Send-NUTCommand "LIST UPS"

    if ($NUTResponse[0] -eq "BEGIN LIST UPS")
    {
        #Connected cleanly, and got back a list of UPSes (good)
        for ($i = 1; $i -lt ($NUTResponse.Count -1) ; $i++) 
        {
            if ($NUTResponse -match "^UPS $TargetUPS ")
            {
                #Write-Host "Found: $($NUTResponse[$i])"
                $UPSFound = $true
            }
        }


    }
    else
    {
        #failed to get a list of UPSes, error 
        Write-Warning "Unexpected NUT Reponse: $($NUTResponse[0])"
    }
    return $UPSFound
}


function Get-NUTVAR
{
    param (
        [string]$TargetUPS,
        [string]$NUTVAR
    )

    try
    {
        $Respose = Send-NUTCommand "GET VAR $TargetUPS $NUTVAR"

        if ($Respose -eq "ERR VAR-NOT-SUPPORTED")
        {

        }
        elseif ($Respose -match "^VAR $TargetUPS $NUTVAR `"(?<NutVAR>\S+)`"")
        {
            return $Matches['NutVAR']
        }
        else
        {
            #wtf
        }
    }
    catch
    {

    }
}


## Syslog (RFC 5424): Emergency → Alert → Critical → Error → Warning → Notice → Informational → Debug
## Windows Log levels and mapping to Syslog levels:
##  Critical (Emergency) → Error (Error)  → Warning (Warning) → Information (Informational)  → Verbose (Debug)
function Write-LogEntry
{
    param (
        [ValidateSet('Emergency','Alert','Critical','Error','Warning','Notice','Informational','Debug')]
        [string]
        $Loglevel, 
        [string]
        $UPSName, 
        [string]
        $UPSStatus, 
        [string]
        $UPSLogType, 
        [string]
        $UPSLogData, 
        [Nullable[DateTime]]
        $LastOL, 
        [Nullable[TimeSpan]]
        $TimeOnBattery, 
        [string]
        $BatteryCharge, 
        [string]
        $BatteryRuntime, 
        [string]
        $MessageString
    )
    $date = (Get-Date).ToUniversalTime()
    $EventTime = $date.ToString("u")
    #$LogFileName = "NUT-$($date.ToUniversalTime().ToString("yyyy-MM-dd")).log"

    $LogFilePath = Join-Path $Config.LogPath "NUT-$($date.ToString("yyyy-MM-dd")).log"

    if (Test-Path -Path $LogFilePath -PathType leaf)
    {
        ## File exists, no need to write header
    }
    else
    {
        "UTC`tLogLevel`tHostname`tSource`tUPSName`tUPSStatus`tUPSLogType`tUPSLogData`tLastOL`tTimeOnBattery`tBatteryPercentage`tBatteryRuntime`tMessageString" | Out-File -Append -FilePath $LogFilePath 
    }
    
    if ($Loglevel -in @('Notice','Informational','Debug'))
    {
        Write-Host "$UPSName`t$UPSStatus`t$UPSLogType`t$BatteryCharge`t$UPSLogData`t$MessageString"
    }
    else
    {
        Write-Warning "$UPSName`t$UPSStatus`t$UPSLogType`t$BatteryCharge`t$UPSLogData`t$MessageString"
    }

    if ($LastOL)
    {
        $LastOLStr = ($LastOL.ToString('u'))
    }

    if ($TimeOnBattery)
    {
        $TimeOnBatteryStr = [int]$TimeOnBattery.Minutes
    }

    "$EventTime `t$Loglevel`t$($env:computername)`tWinNUTPS`t$UPSName`t$UPSStatus`t$UPSLogType`t$UPSLogData`t$LastOLStr`t$TimeOnBatteryStr`t$BatteryCharge`t$BatteryRuntime`t$MessageString" | Out-File -Append -FilePath $LogFilePath 
}
 

## Load local config
if (Test-Path -Path ./WindowNUTPS.config -PathType leaf)
{
    $Config = gc ./WindowNUTPS.config | ConvertFrom-Json
}
else
{
    throw "Startup Failure - Unable to load config"
}

## Basic verfiy correct config type 
if ($Config.ScriptConfig -ne 'WindowNUTPS.ps1')
{
     throw "Startup Failure - Wrong config - expecting WindowNUTPS.ps1 config file"
}

## Logging, check folder exists
if (!(Test-Path -Path $Config.LogPath -PathType Container))
{
    throw "Startup Failure - Log Path does not exist:$($Config.LogPath)"
}

if ($Config.RunMode -ne 'Production')
{
    Write-Warning "Config run mode is not `Production`, system will NOT trigger shutdowns."
}


################
##  Main runtime start and loop

$StartupConnectFailureLogInterval = 60  ## Number of mintus to wait between log entries while still failling to connect to UPS
$OnBatteryLogInterval = 5  ## Number of mintus to wait between log entries while still failling to connect to UPS

$InitTime = (Get-Date).ToUniversalTime()
$LastStatusTime = $InitTime
$StatusRefreshTime = 60  ## Default to 60, drop to 15 if battery not at 100%
$NUTPool_CurrentRate = $Config.NUT.NUTPoll_ErrorRate

$LastConnectTime = $null
$LastOLTime = $null
$LastBattery_charge = $null
$CurrentUPSStatus = $null
$LastFullBatteryTime = $null
$TriggerShutdown = $false
$ShutdownDelay = 0
$LastLogCleanup = $null

while ($true)
{
    $CurrentTime = (Get-Date).ToUniversalTime()

    if (!$LastConnectTime)
    {
        $UPSConnected = Init-NUTConnection $config.NUT.NUTUPSName

        if ($UPSConnected)
        {
            ## Inital connection to UPS successfull 
            $LastConnectTime = $CurrentTime

            $LastUPSStatus = Get-NUTVAR $config.NUT.NUTUPSName "ups.status"

            if ($LastUPSStatus -eq "OL")
            {
                $LastOLTime = $CurrentTime
                #$UPS_battery_charge = Get-NUTVAR $NUTServerUPSName "battery.charge"
                $NUTPool_CurrentRate = $Config.NUT.NUTPoll_OLRate

                #Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $LastUPSStatus -UPSLogType "Startup-Success" -MessageString "Shutdown Thresholds - TimeOnBattery:$($Config.NUT.ShutdownThresholds.TimeOnBattery) BatteryPercentRemaing:$([int]$Config.NUT.ShutdownThresholds.BatteryPercentRemaing) BatteryRuntimeRemaing:$([int]$Config.NUT.ShutdownThresholds.BatteryRuntimeRemaing)"
            }
            else
            {
                ## TODO - Startup not found UPS in OL state
                Write-LogEntry -Loglevel Warning -UPSName $config.NUT.NUTUPSName -UPSLogType "Startup-Warning" -UPSStatus $LastUPSStatus -MessageString "Startup not in OL state"
                #Write-LogEntry -Loglevel Error -UPSName $config.NUT.NUTUPSName -UPSStatus $LastUPSStatus -UPSLogType "Startup-Success" -MessageString "Shutdown Thresholds - TimeOnBattery:$($Config.NUT.ShutdownThresholds.TimeOnBattery) BatteryPercentRemaing:$([int]$Config.NUT.ShutdownThresholds.BatteryPercentRemaing) BatteryRuntimeRemaing:$([int]$Config.NUT.ShutdownThresholds.BatteryRuntimeRemaing)"
            }

            Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $LastUPSStatus -UPSLogType "Startup-Success" -MessageString "Shutdown Thresholds - TimeOnBattery:$($Config.NUT.ShutdownThresholds.TimeOnBattery) BatteryPercentRemaing:$([int]$Config.NUT.ShutdownThresholds.BatteryPercentRemaing) BatteryRuntimeRemaing:$([int]$Config.NUT.ShutdownThresholds.BatteryRuntimeRemaing)"

        }
        else
        {
            ## IF UPS is on battery and NUT is offline, no current way to check 'old' status and trigger a shutdown.
            if ($LastStatusTime.AddMinutes($RetryStatus_Time).CompareTo($CurrentTime) -lt $StartupConnectFailureLogInterval)
            {
                $LastStatusTime = $CurrentTime
                Write-LogEntry -Loglevel Error -UPSName $config.NUT.NUTUPSName -UPSLogType "Startup-UnableToConnect" 
            }
        }
    }
    else
    {
        ## Connected at least once

        ## Main run state
        $CurrentUPSStatus = Get-NUTVAR $config.NUT.NUTUPSName "ups.status"

        if ($CurrentUPSStatus)
        {
            $LastConnectTime = $CurrentTime

            ## Only update values if connected
            [double]$UPSbattery_charge = Get-NUTVAR $config.NUT.NUTUPSName "battery.charge"    ## % of charge
            [int]$UPSbattery_runtime = Get-NUTVAR $config.NUT.NUTUPSName "battery.runtime"  ## Seconds
            #            "battery.runtime.low" ## UPS low runtime value in seconds
            #            "ups.load"            ## % of capacity
        }
        else
        {
            ## UPS connection to NUT lost (or client to NUT), Carry forward last status. 
            $CurrentUPSStatus = $LastUPSStatus 
        }

        if ($UPSbattery_charge -eq 100)
        {
            if ($LastBattery_charge -ne $UPSbattery_charge)
            {
                ## If we just got back to full charge (or on startup, detected full charge)
                $LastStatusTime = $CurrentTime
                $StatusRefreshTime = 60
                Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-Battery" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS Battery at full charge"
            }
            $LastFullBatteryTime = $LastConnectTime
            $TimeOnBattery = $LastConnectTime - $LastFullBatteryTime
        }
        elseif ($LastFullBatteryTime)
        {
            ## IF UPS not full track time on battery, even if connection lost
            $TimeOnBattery = $CurrentTime - $LastFullBatteryTime
        }
        elseif (!($LastBattery_charge))
        {
            ## on startup, report current partal charge status
            $LastStatusTime = $CurrentTime
            $StatusRefreshTime = 15
            Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-Battery" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "Startup UPS Battery charge percentage: $UPSbattery_charge"
        }
        $LastBattery_charge = $UPSbattery_charge

        switch ($CurrentUPSStatus)
        {
                "OL"
                {
                    ## ONLINE                
                    if ($LastUPSStatus -eq "OL")
                    {
                        #UPS normal, and was normal
                        if ($UPSbattery_charge -eq 100)
                        {
                            $NUTPool_CurrentRate = $Config.NUT.NUTPoll_OLRate
                        }
                        if ($LastStatusTime.AddMinutes($StatusRefreshTime).CompareTo($CurrentTime) -lt 1 )
                        {
                            $LastStatusTime = $CurrentTime
                            Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-OL" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS still in OL state"
                        }
                    }
                    else
                    {
                        $LastStatusTime = $CurrentTime
                        Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-OL" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS retuned to OL state from $LastUPSStatus"
                        $LastUPSStatus = $CurrentUPSStatus
                        
                        if ($UPSbattery_charge -ne 100)
                        {
                            ## If moving back to OL from any state, and the battery is not full, set status to be more frequent. 
                            $StatusRefreshTime = 15
                        }
                        ########################################
                        ## If a shutdown was triggered, abort it 
                        if ($TriggerShutdown)
                        {
                            if ($Config.RunMode -eq 'Production')
                            {
                                $ShutRes = cmd /c "C:\Windows\System32\shutdown.exe /a" 2>&1
                                Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-Abort" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS retuned to OL aborting shutdown"
                            }
                            else 
                            {
                                Write-LogEntry -Loglevel Informational -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-Abort" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS retuned to OL aborting shutdown - TEST"
                            }
                        }

                    }
                    $LastOLTime = $CurrentTime
                    $TriggerShutdown = $false
                }
                "OB"
                {
                    ## ON BATTERY"
                    if ($LastUPSStatus -eq "OL")
                    {
                        #UPS no longer normal, and was normal
                        $LastStatusTime = $CurrentTime
                        Write-LogEntry -Loglevel Error -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-OB" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS just went to OB state from $LastUPSStatus"

                        $NUTPool_CurrentRate = $Config.NUT.NUTPoll_ErrorRate
                        #$LastOLTime = $CurrentTime

                    }
                    else
                    {
                        if ($LastStatusTime.AddMinutes($OnBatteryLogInterval).CompareTo($CurrentTime) -lt 1 )
                        {
                            $LastStatusTime = $CurrentTime
                            Write-LogEntry -Loglevel Error -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-OB" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS still OB"
                        }
                    }


                    #################################################
                    #### On battery - function to process threasholds


                    if ($LastFullBatteryTime) 
                    {
                         ## Battery was at 100% at one point, 
                        if ($LastFullBatteryTime.AddMinutes([int]$Config.NUT.ShutdownThresholds.TimeOnBattery).CompareTo($CurrentTime) -lt 1)
                        {
                            $TriggerShutdown = $true
                            $ShutdownDelay = 128
                            ## If longer than $Config.NUT.ShutdownThresholds.TimeOnBattery - Trigger shutdown
                            Write-LogEntry -Loglevel Emergency -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-TimeOnBattery" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS On Battery time of $TimeOnBattery over threshold of $([int]$Config.NUT.ShutdownThresholds.TimeOnBattery)" 
                        }
                    }

                    if ($UPSbattery_charge -le [int]$Config.NUT.ShutdownThresholds.BatteryPercentRemaing)
                    {
                        $TriggerShutdown = $true
                        $ShutdownDelay = 127
                        Write-LogEntry -Loglevel Emergency -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-BatteryPercentRemaing" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS Battery charge of $UPSbattery_charge under threshold of $([int]$Config.NUT.ShutdownThresholds.BatteryPercentRemaing)" 
                    }


                    if ($UPSbattery_runtime -le ([int]$Config.NUT.ShutdownThresholds.BatteryRuntimeRemaing * 60))
                    {
                        $TriggerShutdown = $true
                        $ShutdownDelay = 126
                        Write-LogEntry -Loglevel Emergency -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-BatteryRuntimeRemaing" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS runtime of $UPSbattery_runtime under threshold of $([int]$Config.NUT.ShutdownThresholds.BatteryRuntimeRemaing * 60)" 
                    }

                    $LastUPSStatus = $CurrentUPSStatus
                }
                 "LB"
                {
                    ## LOW BATTERY" - Everything goes off irrelevent of threasholds 
                    #if ($LastUPSStatus -eq "OB")
                    #{
                        $TriggerShutdown = $true
                        $ShutdownDelay = 32
                        #UPS just went from OB to LB
                        $NUTPool_CurrentRate = $Config.NUT.NUTPoll_ErrorRate
                        Write-LogEntry -Loglevel Emergency -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-LB" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS just went to LB state from $LastUPSStatus"
                    #}

                    $LastUPSStatus = $CurrentUPSStatus
                }
                ("OFF", "RB", "OVER", "TRIM", "BOOST", "CAL", "BYPASS")
                {
                    ## No action for these current, keep 'last' status to OL, but don't update the time
                    ## OFF, REPLACE BATTERY, OVERLOAD, VOLTAGE TRIM, VOLTAGE BOOST, CALIBRATION, BYPASS
                    Write-LogEntry -Loglevel Notice -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-Other" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS state not handled yet prior handled state $LastUPSStatus"
                }
                default
                {
                    Write-LogEntry -Loglevel Notice -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "UPS-Other" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "UPS state unexpected state $LastUPSStatus"
                    ## Unknown UPS state
                }
            }

        if ($TriggerShutdown)
        {
            if ($Config.RunMode -eq 'Production')
            {
                $ShutRes = cmd /c "C:\Windows\System32\shutdown.exe /d u:6:12 /t $ShutdownDelay /s /f /c ""NUT Shutdown in $ShutdownDelay""" 2>&1
                #C:\Windows\System32\shutdown.exe "/d" "u:6:12" "/t" $ShutdownDelay "/s" "/f"
                ##U      6       12      Power Failure: Environment
                    
                ## If this is the first triggered shutdown, Log it, if there is one already set, no log.
                if (!($ShutRes -match "A system shutdown has already been scheduled."))
                {
                    Write-LogEntry -Loglevel Emergency -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-Started" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "Shutdown Triggered - Delay: $ShutdownDelay" 
                }
            }
            else
            {
                Write-LogEntry -Loglevel Emergency -UPSName $config.NUT.NUTUPSName -UPSStatus $CurrentUPSStatus -UPSLogType "Shutdown-Started" -LastOL $LastOLTime -TimeOnBattery $TimeOnBattery -BatteryCharge $UPSbattery_charge -BatteryRuntime $UPSbattery_runtime -MessageString "Shutdown Triggered - Delay: $ShutdownDelay - TEST" 
            }
        }
    }

    ## Daily log cleanup check - try/catch to risk failures for just log cleanup problems
    try
    {
        if ($CurrentTime.AddDays(-1) -gt $LastLogCleanup)
        {
            $LastLogCleanup = $CurrentTime
            for ($DayOffset = -([int]($Config.MaxLogAge)); $DayOffset -gt -([int]($Config.MaxLogAge) * 3 ); $DayOffset--)
            { 
                $LogFilePathClean = Join-Path $Config.LogPath "NUT-$($CurrentTime.AddDays($DayOffset).ToString("yyyy-MM-dd")).log"

                if (Test-Path -Path $LogFilePathClean -PathType leaf)
                {
                    Write-host "Removing: $LogFilePathClean"
                    Remove-Item -Path $RemoveItem -ErrorAction SilentlyContinue
                }
            }
        }
    }
    catch
    {

    }

    Start-Sleep -Seconds $NUTPool_CurrentRate 
}