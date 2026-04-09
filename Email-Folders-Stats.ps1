

# Get the list of both primary and archive folders for a user, tag each with its source
#
$userPrincipal = "mso365@tactiohealth.com"
try {
    Connect-ExchangeOnline -UserPrincipalName $userPrincipal -Device
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_"
    exit
}

$primary = Get-EXOMailboxFolderStatistics -Identity mnadeau@caresimple.com `
    | Select-Object *, @{Name="Location"; Expression={"Primary"}}

$archive = Get-EXOMailboxFolderStatistics -Identity mnadeau@caresimple.com -Archive `
    | Select-Object *, @{Name="Location"; Expression={"In-Place Archive"}}

# Combine, sort, and display as a tree
$primary + $archive | Sort-Object Location, FolderPath | ForEach-Object {
    $depth  = ($_.FolderPath -split "/").Count - 2
    $indent = "  " * $depth
    $label  = if ($_.Location -eq "In-Place Archive") { " [ARCHIVE]" } else { "" }
    Write-Host "$indent$($_.Name)$label  ($($_.ItemsInFolder) items)"
}