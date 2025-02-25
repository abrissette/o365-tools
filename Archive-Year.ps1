## Import the Microsoft Graph module
Import-Module Microsoft.Graph

## MS Graph connection
$applicationId = "a97957e3-f190-4b60-b21c-aee9c39f05da" # Your Application (client) ID
$tenantID = "656a82e9-959e-499b-84f6-2357caca4966" # Your Tenant ID
$clientSecret = Read-Host -Prompt "Enter the client secret" -AsSecureString

# Define the user, the archive folder, and the year to archive
$userId = "abrissette@tactiohealth.com"
$year = "2019"
$yearFolderName = "Inbox $year"

# Connect to Microsoft Graph
$clientSecretCredential = New-Object `
    -TypeName System.Management.Automation.PSCredential `
    -ArgumentList $applicationId, $clientSecret
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $clientSecretCredential 

# Check if the folder for the year already exists
# Thanks to @alitarjan.bsky.social who unblocked me for the folder verification and creation
try {
    $yearFolder = Get-MgUserMailFolder -UserId $userId -Filter "DisplayName eq '$yearFolderName'" -ErrorAction Stop | Select-Object DisplayName
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
    $null = New-MgUserMailFolder -UserId $userId -BodyParameter $body

    Write-Host "Folder '$yearFolderName' has been added for user: $($userId)" -ForegroundColor Green
}
else {
    Write-Host "Folder '$yearFolderName' already exists for user: $($userId)" -ForegroundColor Yellow
}

<#
 # {# Connect to Exchange Online - Favor Exo commandlet when available for performance / bulk support reasons  
$userPrincipal = "mso365@tactiohealth.com"
try {
    Connect-ExchangeOnline -UserPrincipalName $userPrincipal -ShowProgress $true
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_"
    exit
}:Enter a comment or description}
#>



<#
Foreach($user in $userList){ 

New-InboxRule -Name " Move Emails to Unifier Folder " -Mailbox $user -From "test@test.com" -MoveToFolder ($Mailbox.alias+': \ testfolder ') 

}
#>

<#


 # Move all email of the year, in Inbox, to the archive folder for the year
foreach ($email in $oldEmails) {
    Move-Mailbox -Identity $UserId -SourceFolder "Inbox" -DestinationFolder "Archive" -Subject $email.Subject
}:Enter a comment or description}
Move-MgUserMailFolderMessage
    -MailFolderId <String>
    -MessageId <String>
    -UserId <String>
#>
