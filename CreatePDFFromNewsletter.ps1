<#
.SYNOPSIS
    Saves the latest Outlook email matching a subject keyword as a PDF.

.DESCRIPTION
    This script searches the default Inbox in Outlook for the latest email matching
    a given subject keyword, opens it, and prints it to PDF using "Microsoft Print to PDF".

.PARAMETER SubjectKeyword
    Keyword to match in the subject line of the newsletter.

.PARAMETER OutputFolder
    Folder where the PDF should be saved.

.EXAMPLE
    .\Save-NewsletterToPDF.ps1 -SubjectKeyword "Weekly Newsletter" -OutputFolder "C:\Newsletters"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SubjectKeyword,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder
)

function Save-EmailToPdf {
    param (
        [Microsoft.Office.Interop.Outlook.MailItem]$MailItem,
        [string]$PdfPath
    )

    $tempPath = "$env:TEMP\NewsletterTemp.msg"
    $MailItem.SaveAs($tempPath)

    # Open in Outlook Inspector (GUI window)
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItemFromTemplate($tempPath)

    # Save as PDF using Microsoft Print to PDF
    $shell = New-Object -ComObject Shell.Application
    $mail.Display()
    Start-Sleep -Seconds 2  # Let the window fully render

    $wshell = New-Object -ComObject WScript.Shell
    $wshell.AppActivate($mail.Subject)
    Start-Sleep -Milliseconds 500
    $wshell.SendKeys("^p")  # Ctrl+P to open print dialog
    Start-Sleep -Seconds 1
    $wshell.SendKeys("{TAB}{TAB}{TAB}{TAB}{TAB}") # Navigate to printer selection
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys("Microsoft Print to PDF")
    Start-Sleep -Milliseconds 500
    $wshell.SendKeys("{ENTER}")
    Start-Sleep -Seconds 2
    $wshell.SendKeys("$PdfPath")
    Start-Sleep -Milliseconds 500
    $wshell.SendKeys("{ENTER}")
    Start-Sleep -Seconds 3

    # Cleanup
    Remove-Item $tempPath -Force
    $mail.Close(0)
}

try {
    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)  # 6 = Inbox

    $items = $inbox.Items | Where-Object {
        $_.Subject -like "*$SubjectKeyword*" -and $_ -is [Microsoft.Office.Interop.Outlook.MailItem]
    } | Sort-Object ReceivedTime -Descending

    if ($items.Count -eq 0) {
        Write-Output "No matching email found with subject containing '$SubjectKeyword'"
        return
    }

    $latest = $items[0]
    $dateStr = $latest.ReceivedTime.ToString("yyyy-MM-dd_HH-mm")
    $safeSubject = ($latest.Subject -replace '[^\w\d-]', '_') -replace '_+', '_'
    $pdfFile = Join-Path $OutputFolder "$safeSubject`_$dateStr.pdf"

    Save-EmailToPdf -MailItem $latest -PdfPath $pdfFile

    Write-Output "Saved newsletter as PDF: $pdfFile"
} catch {
    Write-Error "Failed to save newsletter: $_"
}
