# Install Echange module
Install-Module ExchangeOnlineManagement

# Import Exchange module
Import-Module ExchangeOnlineManagement

# Connect using o365 tenant creds
Connect-IPPSSession

# Read search name
$searchName = (Read-Host "Searchname (e.g Search for spam 01-02-2003)")

# Read subject
$subject = (Read-Host "Subject")

# Create search object
$Search = New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery Subject:$subject

# Start search
Start-ComplianceSearch -Identity $Search.Identity

# Check status of search
# Get-ComplianceSearch -Identity $searchName

# Loop until the search status is 'Completed'
Write-Host "Waiting for the compliance search to complete..." -ForegroundColor Yellow

do {
    # Get the current status of the search
    $status = (Get-ComplianceSearch -Identity $searchName).Status

    # Display current status
    Write-Host "Current status: $status" -ForegroundColor Cyan

    # Wait for a few seconds before checking again
    Start-Sleep -Seconds 10

} while ($status -ne 'Completed')

Write-Host "Compliance search completed!" -ForegroundColor Green

# Delete messages
New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete
