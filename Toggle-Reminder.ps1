# Toggles the Outlook calendar reminder app on/off using a flag file.
# OFF = disabled.flag exists (the login-startup and watchdog launchers then do nothing).
param([switch]$Silent)
Add-Type -AssemblyName System.Windows.Forms

$dir   = $PSScriptRoot
$flag  = Join-Path $dir 'disabled.flag'
$vbs   = Join-Path $dir 'run_reminder.vbs'
$title = 'Calendar Reminder'

function Show($msg, $icon = 'Information') {
    if ($Silent) { Write-Output $msg; return }
    [void][System.Windows.Forms.MessageBox]::Show($msg, $title, 'OK', $icon)
}
function Get-App {
    @(Get-CimInstance Win32_Process -Filter "Name='pythonw.exe'" |
        Where-Object { $_.CommandLine -like '*reminder.py*' })
}

try {
    if (-not (Test-Path $vbs)) { throw "Launcher not found: $vbs" }

    if (-not (Test-Path $flag)) {
        # ---- turn OFF ----
        New-Item $flag -ItemType File -Force | Out-Null
        Get-App | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        Show "Calendar reminders are now OFF.`n`nRun this again to turn them back on." 'Warning'
    }
    else {
        # ---- turn ON ----
        [System.IO.File]::Delete($flag)
        Start-Process wscript.exe -ArgumentList "`"$vbs`""
        for ($i = 0; $i -lt 25 -and (Get-App).Count -eq 0; $i++) { Start-Sleep -Seconds 1 }
        if ((Get-App).Count -gt 0) { Show "Calendar reminders are now ON." }
        else { Show "Turned on, but the app did not start. Check reminder.log." 'Error' }
    }
}
catch {
    Show "Toggle failed:`n$($_.Exception.Message)" 'Error'
}
