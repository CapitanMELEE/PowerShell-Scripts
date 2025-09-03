@"
# PowerShell Scripts

A collection of helpful PowerShell scripts for Microsoft 365 administration.

## Scripts

### Get-DirectLicenseAssignments.ps1
Lists users with a specific M365 license *directly assigned* (not via group).

### Remove-LicensesFromCsv.ps1
Consumes the CSV from the above script and removes the specified license from each user.

## Usage
See comments inside each script for setup and examples.
"@ | Out-File README.md -Encoding utf8

