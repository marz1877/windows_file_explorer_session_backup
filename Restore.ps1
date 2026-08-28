# Restore-FolderTabs v2026-04-02_2 – Open via Shell COM

# Script parameters
param([switch]$DebugLog)

# Configuration Constants
$script:Config = @{
    PollDelaysMs          = @(50, 100, 150, 200, 200, 200, 500, 500, 1000, 1000)  # Adaptive polling (#1)
    SessionTimeoutMinutes = 30
    CleanupInterval       = 10  # GC every N items (#5)
    MaxRetries            = 3
    RetryDelayMs          = 500
}

# Debug logging
$script:LogFile = $null
if ($DebugLog) {
    $scriptDir = $PSScriptRoot
    $script:LogFile = Join-Path $scriptDir "restore_debug_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    "=== Restore Session Started $(Get-Date) ===" | Out-File -FilePath $script:LogFile -Encoding UTF8
}

function Write-DebugLog {
    param([string]$message)
    if ($script:LogFile) {
        "$(Get-Date -Format 'HH:mm:ss.fff') - $message" | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    }
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create singleton Shell COM object with retry logic
$script:ShellApp = $null
for ($i = 0; $i -lt $script:Config.MaxRetries; $i++) {
    try {
        $script:ShellApp = New-Object -ComObject Shell.Application
        Write-DebugLog "Successfully created Shell.Application COM object"
        break
    }
    catch {
        if ($i -eq $script:Config.MaxRetries - 1) {
            $errMsg = "Failed to create Shell.Application COM object after $($script:Config.MaxRetries) attempts: $_"
            Write-DebugLog $errMsg
            Write-Error $errMsg
            exit 1
        }
        Write-DebugLog "COM creation attempt $($i+1) failed, retrying..."
        Start-Sleep -Milliseconds $script:Config.RetryDelayMs
    }
}

function Get-ExplorerPaths {
    # Returns array of all open Explorer folder paths
    $script:ShellApp.Windows() | ForEach-Object {
        try {
            $p = $_.Document.Folder.Self.Path
            if ($p) { $p }
        }
        catch {}
    }
}

function Get-DesktopInventory {
    # Returns a hashtable of { Path = Count } for currently open windows
    $inventory = @{}
    Get-ExplorerPaths | ForEach-Object {
        $p = $_
        if ($inventory.ContainsKey($p)) {
            $inventory[$p]++
        }
        else {
            $inventory[$p] = 1
        }
    }
    return $inventory
}

function Test-SafePath {
    param ([string]$path)
    
    # Security validation to prevent path traversal attacks
    Write-DebugLog "Validating path: $path"
    
    # Reject paths with suspicious patterns
    if ($path -match '\.\.' -or $path -match '[<>"|?*]') {
        Write-DebugLog "Path rejected: contains suspicious characters"
        return $false
    }
    
    # Only allow absolute paths with valid drive letters or UNC paths
    if ($path -notmatch '^[A-Z]:\\' -and $path -notmatch '^\\\\') {
        Write-DebugLog "Path rejected: not an absolute path"
        return $false
    }
    
    Write-DebugLog "Path validated successfully"
    return $true
}

function Set-WindowMaximized {
    param ([string]$path)
    
    # Maximize the window for the given path (Full Screen Feature)
    Write-DebugLog "Attempting to maximize window: $path"
    
    try {
        if (-not ([System.Management.Automation.PSTypeName]'Win32Window').Type) {
            Add-Type @"
                using System;
                using System.Runtime.InteropServices;
                public class Win32Window {
                    [DllImport("user32.dll")]
                    [return: MarshalAs(UnmanagedType.Bool)]
                    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
                    public const int SW_MAXIMIZE = 3;
                }
"@
        }
        
        # Find the window by path
        foreach ($w in $script:ShellApp.Windows()) {
            try {
                $windowPath = $w.Document.Folder.Self.Path
                if ($windowPath -eq $path) {
                    # Get window handle
                    $hwnd = $w.HWND
                    
                    # Maximize using Windows API
                    $result = [Win32Window]::ShowWindow([IntPtr]$hwnd, 3)  # SW_MAXIMIZE = 3
                    
                    if ($result -or $true) {
                        # ShowWindow can return false but still work
                        Write-DebugLog "Window maximized successfully: $path (HWND: $hwnd)"
                        return $true
                    }
                    else {
                        Write-DebugLog "ShowWindow returned false for: $path"
                    }
                }
            }
            catch {
                Write-DebugLog "Error processing window: $_"
                continue
            }
        }
    }
    catch {
        Write-DebugLog "Failed to maximize window: $_"
    }
    
    return $false
}

function Invoke-OpenPathWithCOM {
    param (
        [string]$originalPath,
        [ref]$inventoryRef
    )

    $orig = $originalPath.Trim()
    if ($orig -eq "" -or $orig.StartsWith("#")) { 
        return $null 
    }

    # Security: Validate path before processing (#7)
    if (-not (Test-SafePath -path $orig)) {
        Write-DebugLog "SECURITY: Rejected unsafe path: $orig"
        return @{ Status = "SECURITY REJECTED"; Details = "SECURITY REJECTED: $orig"; Success = $false }
    }

    $candidate = $orig
    $didFallback = $false

    while ($true) {
        if (Test-Path -LiteralPath $candidate) {
            # Feature #1, #4, #5, #6: Advanced Duplication Logic
            $isDirect = ($candidate -ieq $orig)
            
            # For DIRECT matches, try to use an existing window from our snapshot (F1, F4, F6)
            if ($isDirect -and $inventoryRef.Value.ContainsKey($candidate) -and $inventoryRef.Value[$candidate] -gt 0) {
                Write-DebugLog "Path already open (Direct match available), using existing: $candidate"
                $inventoryRef.Value[$candidate]--
                
                # Maximize existing window (F1 requirement - still ensure it's full screen)
                Set-WindowMaximized -path $candidate | Out-Null

                return @{
                    Status  = "ALREADY OPEN"
                    Details = "ALREADY OPEN (Direct match used): $candidate"
                    Success = $true
                }
            }
            
            # NOTE: Fallbacks (Feature #3, #6) ALWAYS open a new window.
            # ALSO: If a direct match is in the file twice but only 1 window existed, 
            # the 2nd one falls through to here and opens (Satisfying F4).

            # get namespace folder
            $ns = $script:ShellApp.NameSpace($candidate)
            if (-not $ns) { break }

            # snapshot before
            $before = @(Get-ExplorerPaths)

            # open via COM verb
            Write-DebugLog "Opening path: $candidate (Fallback: $didFallback)"
            $ns.Self.InvokeVerb("open")

            # Adaptive polling with exponential backoff
            $found = $false
            foreach ($delay in $script:Config.PollDelaysMs) {
                Start-Sleep -Milliseconds $delay
                $after = @(Get-ExplorerPaths)
                $new = $after | Where-Object { $before -notcontains $_ }
                if ($new -icontains $candidate) {
                    $found = $true
                    Write-DebugLog "Window opened successfully after polling: $candidate"
                    break
                }
            }

            if ($found) {
                # Maximize the window (Full Screen Feature)
                Start-Sleep -Milliseconds 200  # Brief delay to ensure window is ready
                Set-WindowMaximized -path $candidate | Out-Null

                if ($didFallback) {
                    return @{
                        Status  = "FALLBACK"
                        Details = "FALLBACK: $orig -> $candidate"
                        Success = $true
                    }
                }
                else {
                    return @{
                        Status  = "RESTORED"
                        Details = "RESTORED: $candidate"
                        Success = $true
                    }
                }
            }
            else {
                return @{
                    Status  = "UNRESOLVED"
                    Details = "UNRESOLVED (no window): $orig (tried $candidate)"
                    Success = $false
                }
            }
        }

        # climb up
        $parent = Split-Path -Path $candidate -Parent
        if (-not $parent -or $parent -eq $candidate) { break }
        $candidate = $parent
        $didFallback = $true
    }

    return @{
        Status  = "UNRESOLVED"
        Details = "UNRESOLVED (no path): $orig"
        Success = $false
    }
}

function Show-ResultPopup {
    param (
        [int]     $restored,
        [int]     $fallbacks,
        [int]     $skipped,
        [int]     $unresolved,
        [int]     $total,
        [string[]]$details
    )

    $form = New-Object Windows.Forms.Form
    $form.Text = "Restore Summary"
    $form.Size = New-Object Drawing.Size(500, 280)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $lbl = New-Object Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Location = New-Object Drawing.Point(20, 20)
    $lbl.Font = New-Object Drawing.Font("Segoe UI", 10)
    $lbl.Text = @"
Total Rows in File:  $total
Restored (Direct):   $restored
Fallbacks (Parent):  $fallbacks
Already Open:        $skipped
Unresolved:          $unresolved
"@
    $form.Controls.Add($lbl)

    $view = New-Object Windows.Forms.Button
    $view.Text = "View Details"
    $view.Size = New-Object Drawing.Size(100, 30)
    $view.Location = New-Object Drawing.Point(20, 180)
    $view.Add_Click({
            $tmp = [IO.Path]::GetTempFileName().Replace('.tmp', '.txt')
            Set-Content -Path $tmp -Value ($details -join "`r`n") -Encoding UTF8
            Start-Process notepad.exe -ArgumentList $tmp
        })
    $form.Controls.Add($view)

    $close = New-Object Windows.Forms.Button
    $close.Text = "Close"
    $close.Size = New-Object Drawing.Size(100, 30)
    $close.Location = New-Object Drawing.Point(140, 180)
    $close.Add_Click({ $form.Close() })
    $form.Controls.Add($close)

    $form.Topmost = $true
    $form.ShowDialog()
}

# MAIN LOOP with session timeout
$script:SessionStart = Get-Date
Write-DebugLog "Session started at $($script:SessionStart)"

while ($true) {
    # Check for session timeout
    if (((Get-Date) - $script:SessionStart).TotalMinutes -gt $script:Config.SessionTimeoutMinutes) {
        Write-Host "`n⏳ Session timeout reached ($($script:Config.SessionTimeoutMinutes) minutes). Exiting..."
        Write-DebugLog "Session timeout reached"
        break
    }
    
    Clear-Host
    $scriptDir = $PSScriptRoot
    $backupDir = Join-Path $scriptDir 'Backups'
    $files = Get-ChildItem -Path $backupDir -Filter "Backup_*.txt" | Sort-Object CreationTime -Descending

    if (-not $files) {
        Write-Host "No backup files found in 'Backups'."
        break
    }

    Write-Host "`nAvailable backup files:`n"
    $files | ForEach-Object { Write-Host $_.Name }

    $choice = Read-Host "`nEnter the backup filename to restore (or 'exit')"
    if ($choice -eq 'exit') { 
        Write-DebugLog "User chose to exit"
        break 
    }

    # Security: Validate filename input (#6)
    if ($choice -match '[<>:"/\\|?*]' -or $choice -match '\.\.') {
        Write-Host "❌ Invalid filename contains forbidden characters."
        Write-DebugLog "SECURITY: Rejected filename with invalid characters: $choice"
        Start-Sleep -Seconds 1.5
        continue
    }

    # Use -LiteralPath to prevent wildcard expansion
    $pathFile = Join-Path $backupDir $choice
    if (-not (Test-Path -LiteralPath $pathFile)) {
        Write-Host "❌ File not found: $choice"
        Write-DebugLog "File not found: $pathFile"
        Start-Sleep -Seconds 1.5
        continue
    }

    $lines = Get-Content -LiteralPath $pathFile -Encoding UTF8
    $valid = $lines | Where-Object { ($_.Trim() -ne '') -and (-not $_.Trim().StartsWith('#')) }

    Write-Host "`n🔄 Restoring $($valid.Count) folders..."
    Write-DebugLog "Starting restoration of $($valid.Count) folders"
    
    # Snapshot of open windows BEFORE this restoration run
    $inventory = Get-DesktopInventory
    
    $stats = @{ Total = 0; Restored = 0; Fallbacks = 0; Skipped = 0; Unresolved = 0; Details = @() }

    $i = 0
    foreach ($ln in $valid) {
        $i++
        $stats.Total++
        
        # Progress indicator
        $percentComplete = [math]::Round(($i / $valid.Count) * 100)
        Write-Progress -Activity "Restoring Folders" -Status "$i of $($valid.Count) - $percentComplete%" -PercentComplete $percentComplete
        
        $res = Invoke-OpenPathWithCOM -originalPath $ln -inventoryRef ([ref]$inventory)
        if ($res) {
            switch ($res.Status) {
                "ALREADY OPEN" { $stats.Skipped++ }
                "RESTORED" { $stats.Restored++ }
                "FALLBACK" { $stats.Fallbacks++ }
                "UNRESOLVED" { $stats.Unresolved++ }
            }
            $stats.Details += $res.Details
        }
        
        # Periodic garbage collection (#5)
        if ($i % $script:Config.CleanupInterval -eq 0) {
            Write-DebugLog "Running garbage collection (iteration $i)"
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    
    Write-Progress -Activity "Restoring Folders" -Completed
    Write-DebugLog "Restoration complete"

    Show-ResultPopup `
        -restored   $stats.Restored `
        -fallbacks  $stats.Fallbacks `
        -skipped    $stats.Skipped `
        -unresolved $stats.Unresolved `
        -total      $stats.Total `
        -details    $stats.Details
}
