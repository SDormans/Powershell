# Define the source directory where the files are located
$sourceDirectory = "C:\Users\Local-Admin\Downloads"

# Define the destination directories for each file type
$musicDirectory = "F:\Musiek"
$photoDirectory = "F:\Foto's"
$pdfDirectory = "F:\Programmeren"

# Create the destination directories if they don't exist
if (-not (Test-Path $musicDirectory)) {
    New-Item -ItemType Directory -Path $musicDirectory | Out-Null
}
if (-not (Test-Path $photoDirectory)) {
    New-Item -ItemType Directory -Path $photoDirectory | Out-Null
}
if (-not (Test-Path $pdfDirectory)) {
    New-Item -ItemType Directory -Path $pdfDirectory | Out-Null
}

# List all files in the source directory
$files = Get-ChildItem -Path $sourceDirectory

# Loop through each file and move them to their respective directories based on file extension
foreach ($file in $files) {
    $extension = $file.Extension.ToLower()

    if ($extension -eq ".mp3" -or $extension -eq ".wav" -or $extension -eq ".flac") {
        Move-Item -Path $file.FullName -Destination $musicDirectory -Force
        Write-Host "Moved $($file.Name) to $musicDirectory"
    }
    elseif ($extension -eq ".jpg" -or $extension -eq ".png" -or $extension -eq ".gif") {
        Move-Item -Path $file.FullName -Destination $photoDirectory -Force
        Write-Host "Moved $($file.Name) to $photoDirectory"
    }
    elseif ($extension -eq ".pdf") {
        Move-Item -Path $file.FullName -Destination $pdfDirectory -Force
        Write-Host "Moved $($file.Name) to $pdfDirectory"
    }
    else {
        Write-Host "File $($file.Name) has an unsupported extension and was not moved."
    }
}
