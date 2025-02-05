## Import the Microsoft Graph module
Import-Module Microsoft.Graph

## Define information here
$ApplicationId = "a97957e3-f190-4b60-b21c-aee9c39f05da" # Your Application (client) ID
$tenantID = "656a82e9-959e-499b-84f6-2357caca4966" # Your Tenant ID
$filepath = "$HOME/Downloads/it/" 
$EmailUserId = "mnadeau@caresimple.com"

# Prompt user for the client secret
$ClientSecret = Read-Host -Prompt "Enter the client secret" -AsSecureString

# Connect to Microsoft Graph
$ClientSecretCredential = New-Object `
    -TypeName System.Management.Automation.PSCredential `
    -ArgumentList $ApplicationId, $ClientSecret
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential 

# Ensure the application has Mail.Read or Mail.ReadBasic permissions
# Store email sent in a specific period 
$endDate = "2019-01-01T00:00:00Z"
$startDate = "2018-12-01T00:00:00Z"

Write-Host "Gettings emails for $EmailUserId sent between $startDate and $endDate"  

try {
    # Get received messages
    $receivedMessages = Get-MgUserMessage -UserId $EmailUserId -Filter "receivedDateTime ge $startDate and receivedDateTime le $endDate" -All

    # Get sent messages
    $sentMessages = Get-MgUserMessage -UserId $EmailUserId -Filter "receivedDateTime ge $startDate and receivedDateTime le $endDate and from/emailAddress/address eq '$EmailUserId'" -All 

    # Combine both received and sent messages
    $messages = $receivedMessages + $sentMessages
} catch {
    Write-Host "Error retrieving messages: $_"
    exit
}

# Check if messages were retrieved successfully
if ($messages -eq $null) {
    Write-Host "No messages retrieved. Please check permissions and try again."
    exit
}
else {
    $messageCount = $messages.Count
    Write-Host $messages. "$messageCount messages were retrieved"
}

# Download all emails (example placeholder for further processing)
foreach ($message in $messages) {
    # Create a filename based on the received date and email subject
    $receivedDate = $message.ReceivedDateTime.ToString("yyyyMMdd_HHmmss")
    $subject = $message.Subject -replace '[\\/:*?"<>|]', '' # Remove invalid filename characters
    $filename = "$filepath$receivedDate`_$subject.txt"

    # Save the email content to a file
    $messageBody = $message.Body.Content
    $messageBody | Out-File -FilePath $filename -Encoding UTF8   
    
    Write-Host "Saving email to $filename"

    # for testing, only process the first email
    break
}