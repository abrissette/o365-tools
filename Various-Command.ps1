# various helpful pwsh command to manage O365

# connect to O365 to manage Exchange.  
# Overlap with mgConnect-Graph but provide access to more Exchange cmdlets
Connect-ExchangeOnline -UserPrincipalName mso365@tactiohealth.com -ShowProgress $true

# Remove all calendar events for someone
Remove-CalendarEvents -Identity reboustani@tactiohealth.com -CancelOrganizedMeetings -QueryWindowInDays 120

# Verify and set permissions on a mailbox (ex: for adding rules with powershel)
Get-EXOMailboxPermission -Identity abrissette@tactiohealth.com | Format-List

Add-MailboxPermission -Identity abrissette -User mso365 -AccessRights FullAccess