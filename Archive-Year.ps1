## Import the Microsoft Graph module
Import-Module Microsoft.Graph

## MS Graph connection
$applicationId = "a97957e3-f190-4b60-b21c-aee9c39f05da" # Your Application (client) ID
$tenantID = "656a82e9-959e-499b-84f6-2357caca4966" # Your Tenant ID
$clientSecret = Read-Host -Prompt "Enter the client secret" -AsSecureString

# Define the user, the archive folder, and the year to archive
$userId = "abrissette@tactiohealth.com"
#$userId = "mnadeau@caresimple.com"
$year = "2023"
$yearFolderName = "Inbox $year"

# Connect to Microsoft Graph
$clientSecretCredential = New-Object `
    -TypeName System.Management.Automation.PSCredential `
    -ArgumentList $applicationId, $clientSecret
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $clientSecretCredential 

# Check if the folder for the year already exists
# Thanks to @alitarjan.bsky.social who unblocked me for the folder verification and creation
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
    Connect-ExchangeOnline -UserPrincipalName $userPrincipal -ShowProgress $true
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_"
    exit
}

# Create a rule to move inbox email for the period into the folder for the year
try {
    New-InboxRule -Name "Archive $year inbox email" -Mailbox $userId -ReceivedAfterDate "$year-01-01T00:00:00Z" -ReceivedBeforeDate "$year-12-31T23:59:59Z" -MoveToFolder "$userId`:\$yearFolderName" -StopProcessingRules $true
    Write-Host "Inbox rule created successfully for $userId" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create inbox rule: $_"
}

# ...then enable auto archive of all message in year folder into In Place Archive coresponding folder


