#TODO try-catch it
#TODO Log it

# Define source and destination paths
# Replace 'YourPhoneName' with the actual name of your phone's MTP device as it appears in Windows Explorer

$sourcePath = "This PC\\YourPhoneName\\Internal Storage\\"
$destinationPath = "C:\\YourDestinationFolder\\"

# Check if the destination folder exists, create if it doesn't
if (!(Test-Path -Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath
}

# Function to copy files recursively from MTP device
function Copy-MTPFiles {
    param (
        [string]$source,
        [string]$destination
    )

    # Get the list of files and directories in the source path
    $items = Get-ChildItem -Path $source -Force -ErrorAction SilentlyContinue

    foreach ($item in $items) {
        $sourceItemPath = Join-Path -Path $source -ChildPath $item.Name
        $destinationItemPath = Join-Path -Path $destination -ChildPath $item.Name

        if ($item.PSIsContainer) {
            # If it's a directory, create it in the destination and recurse
            if (!(Test-Path -Path $destinationItemPath)) {
                New-Item -ItemType Directory -Path $destinationItemPath
            }
            Copy-MTPFiles -source $sourceItemPath -destination $destinationItemPath
        } else {
            # If it's a file, copy it
            Copy-Item -Path $sourceItemPath -Destination $destinationItemPath -Force
        }
    }
}

# Start copying files
Copy-MTPFiles -source $sourcePath -destination $destinationPath

Write-Host "File transfer complete."
