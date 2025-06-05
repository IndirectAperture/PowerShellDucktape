
param
(
    [Parameter(Mandatory=$true)]
    [string]
    $TargetFolder,
    [Parameter(Mandatory=$true)]
    [string]
    $BaseScript,
    [switch]
    $Uninstall
)

## .\Deploy-PSTask.ps1 -BaseScript "WindowNUTPS" -TargetFolder "NUT"

######################
function Deploy-PSTask
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $TargetFolder,
        [Parameter(Mandatory=$true)]
        [string]
        $BaseScript,
        [switch]
        $Uninstall
    )

    $InstallPath = (Join-Path $env:programfiles $TargetFolder)
    $LogsPath = (Join-Path $InstallPath "Logs")

    ## Ensure Admin creds
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = new-object System.Security.Principal.WindowsPrincipal($identity)
    $RunAsadmin = $identity.Groups -contains 'S-1-5-32-544'
    if (!($RunAsadmin))
    {
        write-error "This setup script must be run as local administrator"
        return
    }

    if (!$Uninstall)
    {
        ### Check for ps1
        if (!(Test-Path -Path ".\$BaseScript.ps1" -PathType leaf))
        {
            Write-Warning "Unable to find script file: "
            return
        }

        ### Check for config file
        if (Test-Path -Path ".\$BaseScript.config" -PathType leaf)
        {
            $ConfigFile = $true
        }
        else
        {
            $ConfigFile = $false
            Write-host "Unable to find Config file: $BaseScript.config"
            #return
        }

        ### Deploy log folder(s)
        if (!(Test-Path -PathType Container $InstallPath))
        {
            New-Item -Path $InstallPath -ItemType Directory | out-null
        }
        if (!(Test-Path -PathType Container $LogsPath))
        {
            New-Item -Path $LogsPath -ItemType Directory | out-null
        }

        Write-Host "Copying Files..."
        ## Copy Files
        Copy-Item -Path ".\$BaseScript.ps1" -Destination $InstallPath

        if ($ConfigFile)
        {
            Copy-Item -Path ".\$BaseScript.config" -Destination $InstallPath
        }

        Write-Host "Creating Scheduled Task..."
        ## Create and Register SchTask
        $SchTask_Argument = '-NonInteractive -File "' + $InstallPath + '\' + $BaseScript + '.ps1"'
        $SchTask_Action = New-ScheduledTaskAction -Execute "%windir%\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument $SchTask_Argument -WorkingDirectory $InstallPath
        $SchTask_Trigger = @()
        $SchTask_Trigger += New-ScheduledTaskTrigger -AtStartup
        $SchTask_Trigger += New-ScheduledTaskTrigger -Daily -At 2am
        $SchTask_Principal = New-ScheduledTaskPrincipal -RunLevel Highest -UserId "SYSTEM" -LogonType ServiceAccount
        $SchTask_Settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
        $SchTask_Task = New-ScheduledTask -Action $SchTask_Action -Principal $SchTask_Principal -Trigger $SchTask_Trigger -Settings $SchTask_Settings
        Register-ScheduledTask -TaskName $BaseScript -InputObject $SchTask_Task | Out-Null

        Write-Host "Done"
    }
    else
    {
        #Uninstall
        Write-Host "Removing Files..."
        If (Test-Path "$InstallPath\$BaseScript.ps1" -PathType Container)
        {
            Get-ChildItem "$InstallPath\$BaseScript.ps1" | Remove-Item -Force
        }

        ## Config file and Log folder left inplace
        Write-Host "Removing Scheduled Task..."
        ### Remove SchTask
        Unregister-ScheduledTask -TaskName $BaseScript -Confirm:$false -ErrorAction SilentlyContinue

        Write-Host "Done - Uninstall"
    }
}

if ((!([String]::IsNullOrEmpty($MyInvocation.InvocationName))) -or ($MyInvocation.InvocationName -eq '.'))
{
    if ($PSBoundParameters)
    {
        Deploy-PSTask @PSBoundParameters
    }
    else
    {
        Deploy-PSTask
    }
}