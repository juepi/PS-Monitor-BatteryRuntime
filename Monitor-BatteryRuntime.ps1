###############################################
## PS script to monitor battery runtime
## between charges
## Author: jpichlbauer
## https://github.com/juepi
###############################################
Param ([switch]$verbose)
if ($verbose) {
    # use "-verbose" to get some console output
    $VerbosePreference = "continue" 
}

# Scheduled Task Settings
# Task will be created automagically if necessary
# Task will run with SYSTEM account and start on System startup
$TaskName = "Monitor-BatteryRuntime"

# Dirs and Files
$LogDir = ($PSScriptRoot + "\BatteryRuntime-Log")
$CurrRunTimeFile = ($LogDir + "\CurrBatRuntime.log")
$DiscStartCapFile = ($LogDir + "\DischargeStartCap.log")
$SchedTaskCreated = ($LogDir + "\SchedTaskCreated.txt")
$ResultFile = ($LogDir + "\Results.csv")
#WMI data
$BatStatOnBat = "1" #Any other status means on AC (but not neccessarily charging)
$BatStatWMI = "win32_battery"
$BatStatProp = "BatteryStatus"
$BatCapacityProp = "EstimatedChargeRemaining"

function WaitUntilFullMinute () {
    # Function will sleep until the next full minute (00:00,00:01,00:02,...)
    $gt = Get-Date -Second 0
    do { Start-Sleep -Seconds 1 } until ((Get-Date) -ge ($gt.addminutes(1 - ($gt.minute % 1))))
    return $true
}

###########################################
## Main
###########################################

if ( -not (Test-Path -Path $LogDir)) {
    New-Item -ItemType Directory -Path $PSScriptRoot -Name "BatteryRuntime-Log" | Out-Null
}
if ( -not (Test-Path -Path $ResultFile)) {
    # Create CSV header for result file
    Write-Output '"sep=;"' | Out-File -FilePath $ResultFile
    Write-Output "Date;Battery Runtime [h:m];Discharged Capacity [%];Estimated Full Bat. Runtime [h:m]" | Out-File -FilePath $ResultFile -Append
}

# Create Scheduled Task if neccessary, start it and exit
# BUG: Running get-scheduledtask requires elevation to list our task, so we rely on the helper-file $SchedTaskCreated
if (-not (Test-Path -Path $SchedTaskCreated)) {
    Write-Host "Elevating this script for scheduled task creation.." -ForegroundColor Yellow
    Start-Sleep 2
    if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
        if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
            $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
            Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
            Exit
        }
    }
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-executionpolicy unrestricted -file $($PSScriptRoot + "\" + $MyInvocation.MyCommand.Name)"
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings
    try { Register-ScheduledTask -TaskName $TaskName -User "System" -InputObject $task -ErrorAction Stop }
    catch { Write-Host "FAILED to create scheduled task!" -ForegroundColor Red ; Start-Sleep 5 ; Exit 1 }
    Get-Date | Out-File -FilePath $SchedTaskCreated -Force
    Write-Host "Scheduled Task has been created. Task will be started now and this script instance will be stopped." -ForegroundColor Green
    Write-Host "Results will be logged to: $($ResultFile)" -ForegroundColor Green
    Start-Sleep 5
    Start-ScheduledTask -TaskName $TaskName
    Exit 0
}

#Helpers
[int]$PrevBatStat = (Get-WmiObject $BatStatWMI).$BatStatProp

if (Test-Path -Path $CurrRunTimeFile) {
    [int]$CurrRuntime = (Get-Content $CurrRunTimeFile).Trim()
    Write-Verbose "Read Current Battery Runtime from file: $($CurrRuntime)"
}
else {
    [int]$CurrRuntime = 0
}

if (Test-Path -Path $DiscStartCapFile) {
    [int]$DischargeStartCap = (Get-Content $DiscStartCapFile).Trim()
    Write-Verbose "Read Discharge Start SoC Value from file: $($DischargeStartCap)"
}
else {
    [int]$DischargeStartCap = (Get-WmiObject $BatStatWMI).$BatCapacityProp
}

Write-Verbose "Starting main loop.."

while (WaitUntilFullMinute) {
    $CurrentBatStat = (Get-WmiObject $BatStatWMI).$BatStatProp
    Write-Verbose "Current Battery Status: $($CurrentBatStat)"
    Write-Verbose "Previous Battery Status: $($PrevBatStat)"
    Write-Verbose "Discharge start SoC: $($DischargeStartCap)"

    if ( $CurrentBatStat -ne $BatStatOnBat ) {
        # Currently On AC
        if ($CurrentBatStat -eq $PrevBatStat) {
            Write-Verbose "No status change, still on AC."
            continue
        }
        else {
            Write-Verbose "Now on AC, previously on Battery, generating results."
            [int]$CurrentCap = (Get-WmiObject $BatStatWMI).$BatCapacityProp
            [int]$DischargedCap = $DischargeStartCap - $CurrentCap
            [int]$CalcFullRuntime = $CurrRuntime * 100 / $DischargedCap
            $RuntimeTS = New-TimeSpan -Minutes $CurrRuntime
            $CalcTS = New-TimeSpan -Minutes $CalcFullRuntime            
            # Write to output file
            Write-Output "$($(Get-Date).ToString().Trim());$($RuntimeTS.Hours):$($RuntimeTS.Minutes);$($DischargedCap);$($CalcTS.Hours):$($CalcTS.Minutes)" | Out-File -FilePath $ResultFile -Append
        }
    }
    else {
        # On Battery
        if ($CurrentBatStat -eq $PrevBatStat) {
            # still on battery, update counters
            $CurrRuntime++
            Write-Output "$CurrRuntime" | Out-File -FilePath $CurrRunTimeFile -Force
            Write-Verbose "Current Runtime on Battery: $($CurrRuntime) minutes."
            Write-Verbose "Discharge start SoC: $($DischargeStartCap)"
        }
        else {
            Write-Verbose "Now on battery, start battery runtime measurement."
            $CurrRuntime = 1
            Write-Output "$($CurrRuntime)" | Out-File -FilePath $CurrRunTimeFile -Force
            $DischargeStartCap = (Get-WmiObject $BatStatWMI).$BatCapacityProp
            Write-Output "$($DischargeStartCap)" | Out-File -FilePath $DiscStartCapFile -Force
        }
    }
    # Prepare for next loop
    $PrevBatStat = $CurrentBatStat
}
