# various helpful pwsh command to manage O365

# connect to O365 to manage Exchange.  
# Overlap with mgConnect-Graph but provide access to more Exchange cmdlets

#Connect-ExchangeOnline -UserPrincipalName mso365@tactiohealth.com -ShowProgress $true

Connect-ExchangeOnline -UserPrincipalName mso365@tactiohealth.com -Device


# Remove all calendar events for someone
Remove-CalendarEvents -Identity reboustani@tactiohealth.com -CancelOrganizedMeetings -QueryWindowInDays 120

# Verify and set permissions on a mailbox (ex: for adding rules with powershel)
Get-EXOMailboxPermission -Identity abrissette@tactiohealth.com | Format-List

Add-MailboxPermission -Identity abrissette -User mso365 -AccessRights FullAccess

# Force MRM / Retention Policy to be apply immediately
Start-ManagedFolderAssistant -Identity "mnadeau@tactiohealth.com"   

Get-Mailbox mnadeau@caresimple.com | FL RetentionPolicy

# Troubleshoot index  (With new Outlook Desktop, indexing is on server side)
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com -Device
Get-MailboxStatistics user@domain.com | FL BigFunnel*
