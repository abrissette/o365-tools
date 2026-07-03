<#
.SYNOPSIS
    Extracts out-of-office (OOO) auto-reply messages from the sales@caresimple.com
    shared mailbox to a CSV file for analysis.

.DESCRIPTION
    Connects to Microsoft Graph, queries the sales@caresimple.com mailbox for messages
    with subjects matching common OOO patterns received in the past 36 months, and
    exports key fields to a CSV. The CSV is then fed into a downstream Claude/Cowork
    pipeline that parses, classifies by market segment, dedupes against HubSpot, and
    produces a reviewable XLSX.

.PREREQUISITES
    - PowerShell 7.x recommended (Windows PowerShell 5.1 also works)
    - Microsoft Graph PowerShell SDK installed:
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    - Andre's account must have one of:
        a) Delegated FullAccess (or ReadOnly) on sales@caresimple.com, OR
        b) App-level Mail.Read permission with admin consent

.NOTES
    Author:  Drafted by Claude/Cowork for Michel Nadeau / CareSimple Inc.
    Date:    2026-05-26
    Target:  sales@caresimple.com
    Scope:   2023-05-26 onward (36 months)
    Output:  CSV in user's Desktop folder

.EXAMPLE
    pwsh -File .\Extract-CareSimple-OOO-Replies.ps1
#>

# -------- Configuration --------
$MailboxOwner = "sales@caresimple.com"
$StartDate    = (Get-Date "2023-05-26").ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndDate      = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$OutputCsv    = Join-Path ([Environment]::GetFolderPath("Desktop")) "caresimple_ooo_replies_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " CareSimple OOO Reply Extractor" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "Mailbox:      $MailboxOwner"
Write-Host "Date range:   $StartDate  to  $EndDate"
Write-Host "Output:       $OutputCsv"
Write-Host ""

# -------- Connect to Graph --------
# Mail.Read.Shared = delegated access to mailboxes Andre has explicit permission on.
# If Andre has app-level Mail.Read with admin consent, this still works.
try {
    Connect-MgGraph -Scopes "Mail.Read.Shared" -NoWelcome
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

# Verify the connection
$context = Get-MgContext
Write-Host "Connected as: $($context.Account)" -ForegroundColor Green
Write-Host ""

# -------- Build the filter --------
# We capture multiple subject patterns because OOO replies aren't standardized:
#   - "Automatic reply:" (Microsoft Outlook default)
#   - "Automatic Reply:" (case variation)
#   - "Out of Office:" / "Out of the Office:"
#   - "Auto-Reply:" / "AutoReply:"
# Server-side OData filter narrows results dramatically before they hit the network.
$Filter = @"
receivedDateTime ge $StartDate and receivedDateTime le $EndDate and (
    startswith(subject,'Automatic reply') or
    startswith(subject,'Automatic Reply') or
    startswith(subject,'AutoReply') or
    startswith(subject,'Auto-Reply') or
    startswith(subject,'Auto Reply') or
    startswith(subject,'Out of Office') or
    startswith(subject,'Out of the Office')
)
"@ -replace "`r?`n", " "

# Properties we need for downstream parsing
$Select = "id,subject,from,sender,receivedDateTime,sentDateTime,bodyPreview,internetMessageId,hasAttachments"

Write-Host "Querying Graph for OOO messages (this may take 30-90 seconds)..." -ForegroundColor Yellow

try {
    $messages = Get-MgUserMessage `
        -UserId $MailboxOwner `
        -Filter $Filter `
        -Property $Select `
        -PageSize 100 `
        -All `
        -ErrorAction Stop
} catch {
    Write-Error "Graph query failed: $_"
    Write-Host ""
    Write-Host "Common causes:" -ForegroundColor Yellow
    Write-Host "  - Your account doesn't have permission on $MailboxOwner"
    Write-Host "    Fix: have a Global Admin grant FullAccess or ReadOnly on the shared mailbox"
    Write-Host "  - The Mail.Read.Shared scope was not consented"
    Write-Host "    Fix: Disconnect-MgGraph, then run again and approve the consent prompt"
    Write-Host "  - The mailbox owner identifier is wrong"
    Write-Host "    Fix: confirm sales@caresimple.com is the exact UPN"
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit 1
}

Write-Host "Found $($messages.Count) OOO messages." -ForegroundColor Green
Write-Host ""

# -------- Transform and export --------
Write-Host "Building CSV..." -ForegroundColor Yellow

$rows = $messages | ForEach-Object {
    # Sender vs From: Microsoft Graph distinguishes the technical sender from the
    # "From" header. For OOO replies the From header is the actual person.
    $fromAddr = $null
    $fromName = $null
    if ($_.From -and $_.From.EmailAddress) {
        $fromAddr = $_.From.EmailAddress.Address
        $fromName = $_.From.EmailAddress.Name
    } elseif ($_.Sender -and $_.Sender.EmailAddress) {
        $fromAddr = $_.Sender.EmailAddress.Address
        $fromName = $_.Sender.EmailAddress.Name
    }

    # Clean preview: collapse all whitespace runs (incl. newlines) to single spaces.
    $preview = $_.BodyPreview
    if ($preview) {
        $preview = ($preview -replace "[\r\n\t]+", " " -replace "\s{2,}", " ").Trim()
    }

    [PSCustomObject]@{
        sender_email        = $fromAddr
        sender_name         = $fromName
        received            = if ($_.ReceivedDateTime) { $_.ReceivedDateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } else { "" }
        subject             = $_.Subject
        preview             = $preview
        internet_message_id = $_.InternetMessageId
        has_attachments     = $_.HasAttachments
    }
}

# Drop rows with no sender_email (extremely rare; usually system-generated bounces).
$rowsWithEmail = $rows | Where-Object { $_.sender_email -and $_.sender_email -ne "" }
$droppedCount  = $rows.Count - $rowsWithEmail.Count
if ($droppedCount -gt 0) {
    Write-Host "Dropped $droppedCount rows with no sender email." -ForegroundColor DarkYellow
}

# Export
$rowsWithEmail | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host " Export complete." -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host "Records exported: $($rowsWithEmail.Count)" -ForegroundColor Green
Write-Host "File:             $OutputCsv" -ForegroundColor Green
Write-Host ""
Write-Host "Next step: share $OutputCsv with Michel for downstream processing."
Write-Host ""

# Quick stats by year and month so we can spot gaps
Write-Host "Distribution by month:" -ForegroundColor Cyan
$rowsWithEmail |
    Group-Object { ([datetime]$_.received).ToString("yyyy-MM") } |
    Sort-Object Name -Descending |
    Select-Object @{N="Month";E={$_.Name}}, @{N="Count";E={$_.Count}} |
    Format-Table -AutoSize

Disconnect-MgGraph -ErrorAction SilentlyContinue