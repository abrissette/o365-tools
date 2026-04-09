## Import the Microsoft Graph module
Import-Module Microsoft.Graph

## MS Graph connection
$applicationId = "a97957e3-f190-4b60-b21c-aee9c39f05da" # Your Application (client) ID
$tenantID = "656a82e9-959e-499b-84f6-2357caca4966" # Your Tenant ID
$clientSecret = Read-Host -Prompt "Enter the client secret" -AsSecureString

# Define the user, the archive folder, and the year to archive
#$userId = "abrissette@tactiohealth.com"
$userId = "mnadeau@caresimple.com"
$year = Read-Host -Prompt "Enter year to archive" 
$yearFolderName = "Inbox $year"

# Connect to Microsoft Graph
$clientSecretCredential = New-Object `
    -TypeName System.Management.Automation.PSCredential `
    -ArgumentList $applicationId, $clientSecret
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $clientSecretCredential 

# Check if the folder for the year already exists under 
try {
    $yearFolder = Get-MgUserMailFolder -UserId $userId -Filter "DisplayName eq '$yearFolderName'" -ErrorAction Stop | Select-Object DisplayName, Id
} 
catch {
    Write-Host "Error retrieving folders for user: $($userId)" -ForegroundColor Red
    Exit
}

if (-not $yearFolder) {
    # If the folder does not exist, create it
    $body = @{
        DisplayName = $yearFolderName
    }
    $yearFolder = New-MgUserMailFolder -UserId $userId -BodyParameter $body

    Write-Host "Folder '$yearFolderName' has been added for user: $($userId)" -ForegroundColor Green
}
else {
    Write-Host "Folder '$yearFolderName' already exists for user: $($userId)" -ForegroundColor Yellow
}

# Connect to Exchange Online - Favor Exo commandlet when available for performance / bulk support reasons  
$userPrincipal = "mso365@tactiohealth.com"
try {
    Connect-ExchangeOnline -UserPrincipalName $userPrincipal -Device
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_"
    exit
}

# Replicate a "move to folder" rule manually for existing messages
$messages = Get-MgUserMailFolderMessage -UserId $userId `
    -MailFolderId "inbox" -All -PageSize 100 `
    -Filter "receivedDateTime ge $year-01-01T00:00:00Z and receivedDateTime le $year-12-31T23:59:59Z"

$total = $messages.Count
$count = 0
Write-Host "Moving $total messages from Inbox to '$yearFolderName'..." -ForegroundColor Cyan

foreach ($message in $messages) {
    try {
        Move-MgUserMessage -UserId $userId `
            -MessageId $message.Id `
            -BodyParameter @{ DestinationId = $yearFolder.Id }
        $count++
        if ($count % 100 -eq 0) {
            Write-Host "  Moved $count / $total messages..." -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "  Failed to move message $($message.Id): $_" -ForegroundColor Red
    }
}

Write-Host "Done. Sorted $count / $total messages to '$yearFolderName'." -ForegroundColor Green

# Apply "Move to Archive Immediately" retention tag to the year folder
$tagId = "bdc79f42-c30c-4ed6-9635-d450567e1d1a"
try {
    $folder = Get-MgUserMailFolder -UserId $userId -Filter "DisplayName eq '$yearFolderName'" -ErrorAction Stop
    Update-MgUserMailFolder -UserId $userId -MailFolderId $folder.Id `
        -BodyParameter @{ retentionTag = @{ isExplicit = $true; tagId = $tagId } }
    Write-Host "Retention tag applied to '$yearFolderName'" -ForegroundColor Green
}
catch {
    Write-Host "Failed to apply retention tag to '$yearFolderName': $_" -ForegroundColor Red
}

# Trigger Managed Folder Assistant to process the tag immediately
try {
    Start-ManagedFolderAssistant -Identity $userId
    Write-Host "Managed Folder Assistant triggered for $userId" -ForegroundColor Green
}
catch {
    Write-Host "Failed to trigger Managed Folder Assistant: $_" -ForegroundColor Red
}

