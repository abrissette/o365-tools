## Import the Microsoft Graph module
Import-Module Microsoft.Graph

## Define information here
$applicationId = "a97957e3-f190-4b60-b21c-aee9c39f05da" # Your Application (client) ID
$tenantID = "656a82e9-959e-499b-84f6-2357caca4966" # Your Tenant ID
$filePath = "$HOME/Downloads/it/" 
$emailUserId = "mnadeau@caresimple.com"

# Prompt user for the client secret
$clientSecret = Read-Host -Prompt "Enter the client secret" -AsSecureString

# Connect to Microsoft Graph
$clientSecretCredential = New-Object `
    -TypeName System.Management.Automation.PSCredential `
    -ArgumentList $applicationId, $clientSecret
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $clientSecretCredential 

# Retrieve the token
$context = Get-MgContext
if ($null -eq $context) {
    Write-Host "Failed to retrieve Microsoft Graph context."
    exit
}

$token = $context.AccessToken
if ($null -eq $token) {
    Write-Host "Failed to retrieve access token."
    exit
}

Write-Host "Connected to Microsoft Graph with token: $token"

# Ensure the application has Mail.Read or Mail.ReadBasic permissions 
# Make sure its for all mailboxes cause there is also a Mail.Read for current user only

# Store email sent in a specific period 
$startDate = "2018-12-01T00:00:00Z"
$endDate = "2018-12-02T00:00:00Z"


Write-Host "Gettings emails for $emailUserId sent between $startDate and $endDate"  

try {
    # Get received messages
    $receivedMessages = Get-MgUserMessage -UserId $emailUserId -Filter "receivedDateTime ge $startDate and receivedDateTime le $endDate" -All

    # Get sent messages
    $sentMessages = Get-MgUserMessage -UserId $emailUserId -Filter "receivedDateTime ge $startDate and receivedDateTime le $endDate and from/emailAddress/address eq '$emailUserId'" -All 

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

# Download all emails 
foreach ($message in $messages) {
    # Create a filename with email subject and received date
    $fileName = ($File = "$($message.subject) $($message.ReceivedDateTime).eml").Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
    $outFile = $filePath + $fileName
    $parentFolder = $message.parentFolderId

    # Get the name of the parent folder
    try {
        $parentFolderName = (Get-MgUserMailFolder -UserId $emailUserid -MailFolderId $parentFolder).DisplayName
    }
    Catch {
        $parentFolderName = "Unknown Folder"
}

    # Save the email content to a file
    try {
        Get-MgUserMessageContent -UserId $emailUserid -MessageId $message.id -OutFile $outfile
        Write-Host "Exported email $outfile in folder $parentFolderName"

        # If attachments save it to a folder 
        $attachments = Get-MgUserMessageAttachment -UserId $emailUserid -MessageId $message.id 
        $attachmentsCount = $attachments.Count
        Write-Host "email has $attachmentsCount attachments"

        if ($attachmentsCount -gt 0) {
        
            $attachmentFolder = $filePath + $fileName + "_attachments"
            if (!(Test-Path $attachmentFolder)) {
                New-Item -ItemType Directory -Path $attachmentFolder
            }
            
            foreach ($attachment in $attachments) {
                $attachmentFileName = $attachmentFolder + "/" + ($attachment.Name -replace '[<>:"/\\|?*]', '_')
                      
                # Retrieve and save the attachment content
                try {
                    $attachmentContent = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$emailUserid/messages/$($message.id)/attachments/$($attachment.Id)/$value" -Headers @{Authorization = "Bearer $($token)"}
                    [System.IO.File]::WriteAllBytes($attachmentFileName, [System.Convert]::FromBase64String($attachmentContent.ContentBytes))
        
                    Write-Host "Exported attachment $attachmentFileName"
                }
                Catch {
                    Write-Host "Failed to export attachment $attachmentFileName : $_"
                }        
            }
        }    

        # Delete the email after exporting
        #Remove-MgUserMessage -UserId $emailUserid -MessageId $message.id
        #Write-Host "Deleted email $message.id"

    }
    Catch {
        Write-Host "Unable to export email $fileName in folder $parentFolderName : $_"
    }    
    
    # for testing, only process the first email
    #break
}