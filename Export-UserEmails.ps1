## Import the Microsoft Graph module
Import-Module Microsoft.Graph

## Define information here
$ApplicationId = "a97957e3-f190-4b60-b21c-aee9c39f05da" # Your Application (client) ID
$tenantID = "656a82e9-959e-499b-84f6-2357caca4966" # Your Tenant ID
$filepath = "$HOME/Downloads/it/" # Example: /Users/username/Downloads/
$EmailUserId = "mnadeau@caresimple.com"

# Prompt user for the client secret
$ClientSecret = Read-Host -Prompt "Enter the client secret" -AsSecureString

# Connect to Microsoft Graph
$ClientSecretCredential = New-Object `
    -TypeName System.Management.Automation.PSCredential `
    -ArgumentList $ApplicationId, $ClientSecret
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential

# Ensure the application has Mail.Read or Mail.ReadBasic permissions
# Store all emails before January 1st, 2019
$filterDate = "2019-01-01T00:00:00Z"
$messages = Get-MgUserMessage -UserId $EmailUserId -Filter "receivedDateTime lt $filterDate" -All

# Download all emails (example placeholder for further processing)
foreach ($message in $messages) {
    # Add your code to download or process each message here
}