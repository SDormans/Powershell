<#
.SYNOPSIS
    Finds and prints newly created report files (e.g., PDFs) in a folder.

.DESCRIPTION
    This script scans a folder for newly created report files (default: .pdf)
    within the last 24 hours and sends them to the default printer.

.PARAMETER Folder
    The folder where reports are stored.

.PARAMETER Extension
    (Optional) File extension to search for. Default is '.pdf'.

.PARAMETER Hours
    (Optional) Look back window in hours. Default is 24.

.EXAMPLE
    .\Print-NewReports.ps1 -Folder "C:\Reports"

.EXAMPLE
    .\Print-NewReports.ps1 -Folder "C:\Reports" -Extension ".docx" -Hours 12
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Folder,

    [string]$Extension = ".pdf",

    [int]$Hours = 24
)

function Print-File {
    param (
        [string]$FilePath
    )
    try {
        Start-Process -FilePath $FilePath -Verb Print -PassThru | Out-Null
        Write-Output "Sent to printer: $FilePath"
    } catch {
        Write-Warning "Failed to print: $FilePath - $_"
    }
}

try {
    if (-Not (Test-Path $Folder)) {
        throw "Folder '$Folder' does not exist."
    }

    $cutoff = (Get-Date).AddHours(-$Hours)
    Write-Output "Searching for *$Extension files created after $cutoff in $Folder"

    $newReports = Get-ChildItem -Path $Folder -Filter "*$Extension" -File -Recurse | Where-Object {
        $_.CreationTime -gt $cutoff
    }

    if ($newReports.Count -eq 0) {
        Write-Output "No new reports found."
        return
    }

    foreach ($report in $newReports) {
        Print-File -FilePath $report.FullName
    }

    Write-Output "Printing complete. Total printed: $($newReports.Count)"
} catch {
    Write-Error "Error: $_"
}
