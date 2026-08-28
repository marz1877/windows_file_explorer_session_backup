# Backup v2026-04-02

# Configuration Constants
$script:Config = @{
    MaxBackupFiles = 50
    DateTimeFormat = "yyyy-MM-dd_HH-mm-ss"
    FileEncoding   = "UTF8"
}

$shellApplication = New-Object -ComObject Shell.Application
$windows = $shellApplication.Windows()
$openFolders = @()
$scriptDirectory = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
$backupDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDirectory "Backups")).TrimEnd('\')

function Replace-GuidWithPathLabel {
    param ([string]$folderPath)
    # More precise GUID pattern with case-insensitive flag
    if ($folderPath -match '(?i)\{[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}\}') {
        return "Home"
    }
    return $folderPath
}

# Preserve window order as provided by Shell
foreach ($window in $windows) {
    try {
        $rawPath = $window.Document.Folder.Self.Path
        if ($rawPath) {
            try {
                # Normalize for comparison (add trailing slash for recursive check)
                $normPath = [System.IO.Path]::GetFullPath($rawPath).TrimEnd('\') + '\'
                $normScript = $scriptDirectory + '\'
                $normBackup = $backupDir + '\'

                # Feature: Ignore the directory the script is in and the Backups folder (Recursive) (#37)
                if ($normPath.StartsWith($normScript, [System.StringComparison]::OrdinalIgnoreCase) -or 
                    $normPath.StartsWith($normBackup, [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }
            }
            catch {
                # If path isn't a valid filesystem path (e.g. ::{GUID}), just keep processing
                Write-Debug "Skipping normalization for special path: $rawPath"
            }

            $folderPath = Replace-GuidWithPathLabel $rawPath
            $openFolders += $folderPath
        }
    }
    catch {}
}

$tabCount = $openFolders.Count
$dateTime = Get-Date -Format $script:Config.DateTimeFormat
$outputFileName = "Backup_${dateTime}-${tabCount}.txt"

# Ensure Backups folder exists
if (-not (Test-Path $backupDir)) {
    New-Item -Path $backupDir -ItemType Directory | Out-Null
}

$tempFile = Join-Path $backupDir "Temp_Backup.txt"
$openFolders | Out-File -FilePath $tempFile -Encoding $script:Config.FileEncoding

$backupFiles = Get-ChildItem -Path $backupDir -Filter "Backup_*.txt" |
Sort-Object { $_.Name -replace "Backup_|\.txt", "" } -Descending

if ($backupFiles.Count -gt 0) {
    $lastBackupFile = $backupFiles[0]
    
    # Use hash comparison for better performance
    $lastHash = (Get-FileHash -Path $lastBackupFile.FullName -Algorithm MD5).Hash
    $newHash = (Get-FileHash -Path $tempFile -Algorithm MD5).Hash

    if ($lastHash -eq $newHash) {
        # Files are identical, update timestamp of existing backup
        (Get-Item $lastBackupFile.FullName).LastWriteTime = Get-Date
        Remove-Item $tempFile -Force
    }
    else {
        # Content changed, create new backup
        Rename-Item -Path $tempFile -NewName (Join-Path $backupDir $outputFileName)
    }
}
else {
    # First backup
    Rename-Item -Path $tempFile -NewName (Join-Path $backupDir $outputFileName)
}

# Prune old backups
$backupFiles = Get-ChildItem -Path $backupDir -Filter "Backup_*.txt" |
Sort-Object { $_.Name -replace "Backup_|\.txt", "" } -Descending

if ($backupFiles.Count -gt $script:Config.MaxBackupFiles) {
    $backupFiles | Select-Object -Skip $script:Config.MaxBackupFiles | Remove-Item -Force
}

exit
