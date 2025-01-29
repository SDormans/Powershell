# Define source & destination, and log paths
# Replace 'YourPhoneName' with the actual name of your phone's MTP device as it appears in Windows Explorer

$sourcePath = "This PC\\YourPhoneName\\Internal Storage\\"
$destinationPath = "C:\\YourDestinationFolder\\"
$LogFile = "C:\Logs\Transferfileslog.txt" 

# Ensure the log directory exists
$LogDir = Split-Path $LogFile
if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Check if the destination folder exists, create if it doesn't
if (!(Test-Path -Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath
}

# Redirect output streams to the log file
Start-Transcript -Path $LogFile -Append

# Function to copy files recursively from MTP device
function Copy-MTPFiles {
    param (
        [string]$source,
        [string]$destination
    )

    write-host (Get-Date) "script started" -ForegroundColor Cyan
    # Get the list of files and directories in the source path
    $items = Get-ChildItem -Path $source -Force -ErrorAction SilentlyContinue

    foreach ($item in $items) {
        $sourceItemPath = Join-Path -Path $source -ChildPath $item.Name
        $destinationItemPath = Join-Path -Path $destination -ChildPath $item.Name
        try { 
            if ($item.PSIsContainer) {
                # If it's a directory, create it in the destination and recurse
                if (!(Test-Path -Path $destinationItemPath)) {
                    New-Item -ItemType Directory -Path $destinationItemPath
                    Write-Host "Folder copied successfully."
                }
                Copy-MTPFiles -source $sourceItemPath -destination $destinationItemPath
            } else {
                # If it's a file, copy it
                Copy-Item -Path $sourceItemPath -Destination $destinationItemPath -Force
                Write-Host "Item copied successfully."
            }    
        }
        catch {
             # Log errors
            Write-Host (get-date) "An error occurred: $_" -ForegroundColor Red
        }       
    }
}
try {
    #Start copying files
    Copy-MTPFiles -source $sourcePath -destination $destinationPath
    Write-Host "File transfer complete."
}
catch {
     # Log errors
     Write-Host (get-date) "An error occurred: $_" -ForegroundColor Red
}
finally {
    # Completion message
    Write-Host (Get-Date)"Operation complete. Check the log at $LogFile for details." -ForegroundColor Cyan
    # Always close the transcript
    Stop-Transcript
}