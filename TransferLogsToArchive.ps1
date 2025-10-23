<#
.SYNOPSIS
    Moves log files older than 6 months to an archive folder.

.DESCRIPTION
    This script scans a specified source folder for *.log files older than 6 months (180 days)
    and moves them to a specified archive folder. Supports logging, -WhatIf, and -Verbose.

.PARAMETER SourcePath
    The folder containing the log files to archive.

.PARAMETER ArchivePath
    The destination folder where old log files will be moved.

.PARAMETER LogFile
    (Optional) A log file path to record actions and errors.

.EXAMPLE
    .\Move-OldLogsToArchive.ps1 -SourcePath "C:\Logs" -ArchivePath "D:\Archive\Logs" -LogFile "C:\Logs\transfer.log" -Verbose -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ArchivePath,

    [string]$LogFile
)

function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"

    if ($LogFile) {
        Add-Content -Path $LogFile -Value $logEntry
    }

    Write-Verbose $logEntry
}

try {
    # Validate paths
    if (-not (Test-Path -Path $SourcePath)) {
        throw "Source path '$SourcePath' does not exist."
    }

    if (-not (Test-Path -Path $ArchivePath)) {
        Write-Log "Archive path '$ArchivePath' does not exist. Creating it..."
        New-Item -Path $ArchivePath -ItemType Directory -Force | Out-Null
        Write-Log "Archive path created: $ArchivePath"
    }

    # Define cutoff date (6 months = approx. 180 days)
    $cutoffDate = (Get-Date).AddMonths(-6)
    Write-Log "Starting transfer of log files older than $cutoffDate from '$SourcePath' to '$ArchivePath'"

    # Get .log files older than 6 months
    $oldLogs = Get-ChildItem -Path $SourcePath -Filter *.log -File -Recurse | Where-Object {
        $_.LastWriteTime -lt $cutoffDate
    }

    foreach ($file in $oldLogs) {
        $destination = Join-Path -Path $ArchivePath -ChildPath $file.Name

        if ($PSCmdlet.ShouldProcess($file.FullName, "Move to $destination")) {
            Move-Item -Path $file.FullName -Destination $destination -Force
            Write-Log "Moved: $($file.FullName) -> $destination"
        }
    }

    Write-Log "Transfer complete. Total files moved: $($oldLogs.Count)"
} catch {
    Write-Error "Error: $_"
    Write-Log "ERROR: $_"
}
