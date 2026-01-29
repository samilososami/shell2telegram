# ============================================
# Telegram Bot Cleanup Script
# ============================================

Write-Host "=== TELEGRAM BOT CLEANUP ==="
Write-Host "Removing bot persistence and processes"
Write-Host ""

# 1. Stop bot processes
Write-Host "[+] Stopping bot processes..."
$processesStopped = 0
$allPowershellProcesses = Get-Process -Name "powershell*" -ErrorAction SilentlyContinue

foreach ($proc in $allPowershellProcesses) {
    if ($proc.Id -eq $PID) { continue }
    
    try {
        $wmiProcess = Get-WmiObject Win32_Process -Filter "ProcessId = $($proc.Id)" -ErrorAction SilentlyContinue
        $commandLine = $wmiProcess.CommandLine
        
        if ($commandLine -like "*TelegramBot*" -or
            $commandLine -like "*WindowsUpdate*" -or
            $commandLine -like "*$env:APPDATA\Microsoft\Windows\TelegramBot.ps1*" -or
            $commandLine -like "*$env:APPDATA\WindowsUpdate.ps1*") {
            
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            $processesStopped++
        }
    } catch { }
}

Write-Host "   Stopped processes: $processesStopped"

# 2. Remove registry entries
Write-Host "[+] Removing registry entries..."
$registryRemoved = 0
$registryPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
)
$registryNames = @("TelegramBotService", "WindowsUpdate", "WinUpdate", "UpdateService")

foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        foreach ($name in $registryNames) {
            try {
                $property = Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue
                if ($property) {
                    Remove-ItemProperty -Path $path -Name $name -Force -ErrorAction SilentlyContinue
                    $registryRemoved++
                }
            } catch { }
        }
    }
}

Write-Host "   Registry entries removed: $registryRemoved"

# 3. Remove scheduled tasks
Write-Host "[+] Removing scheduled tasks..."
$tasksRemoved = 0
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if ($isAdmin) {
    try {
        $allTasks = Get-ScheduledTask -ErrorAction SilentlyContinue
        foreach ($task in $allTasks) {
            $shouldRemove = $false
            
            if ($task.TaskName -like "*TelegramBot*" -or
                $task.TaskName -like "*WinUpdate*" -or
                $task.TaskName -like "*WindowsUpdate*") {
                $shouldRemove = $true
            }
            
            if (-not $shouldRemove -and $task.Actions) {
                foreach ($action in $task.Actions) {
                    if ($action.Execute -like "*TelegramBot.ps1*" -or
                        $action.Execute -like "*WindowsUpdate.ps1*") {
                        $shouldRemove = $true
                        break
                    }
                }
            }
            
            if ($shouldRemove) {
                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                $tasksRemoved++
            }
        }
    } catch { }
} else {
    Write-Host "   (Admin required for tasks)"
}

Write-Host "   Scheduled tasks removed: $tasksRemoved"

# 4. Delete script files
Write-Host "[+] Deleting script files..."
$filesDeleted = 0
$filesToDelete = @(
    "$env:APPDATA\Microsoft\Windows\TelegramBot.ps1",
    "$env:APPDATA\WindowsUpdate.ps1",
    "$env:APPDATA\Microsoft\Windows\WindowsUpdate.ps1",
    "$env:USERPROFILE\Desktop\TelegramBot.ps1",
    "$env:USERPROFILE\Downloads\TelegramBot.ps1",
    "C:\TelegramBot.ps1"
)

foreach ($file in $filesToDelete) {
    if (Test-Path $file) {
        try {
            Remove-Item $file -Force -ErrorAction SilentlyContinue
            $filesDeleted++
        } catch { }
    }
}

Write-Host "   Files deleted: $filesDeleted"

# 5. Clean environment variables
Write-Host "[+] Cleaning environment variables..."
try {
    Remove-ItemProperty -Path "HKCU:\Environment" -Name "TELEGRAM_BOT_TOKEN" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\Environment" -Name "TELEGRAM_CHAT_ID" -ErrorAction SilentlyContinue
} catch { }

# Summary
Write-Host ""
Write-Host "=== CLEANUP COMPLETE ==="
Write-Host "Processes stopped: $processesStopped"
Write-Host "Registry entries: $registryRemoved"
Write-Host "Scheduled tasks: $tasksRemoved"
Write-Host "Files deleted: $filesDeleted"
Write-Host ""
Write-Host "Bot removed from system."
