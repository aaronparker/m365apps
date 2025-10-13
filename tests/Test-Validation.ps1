#Requires -PSEdition Desktop
# Test script for validation functions

# Import the module
Import-Module "$PSScriptRoot\Microsoft365AppsPackage.psm1" -Force

# Test validation with ValidateOnly parameter
Write-Host "Testing validation functionality..." -ForegroundColor Green

# Test parameters that should work
$testParams = @{
    Path = $PSScriptRoot
    Destination = "$PSScriptRoot\package-test"
    ConfigurationFile = "$PSScriptRoot\configs\O365ProPlus.xml"
    Channel = "Current"
    CompanyName = "TestCompany"
    TenantId = "12345678-1234-1234-1234-123456789012"
    ClientId = "87654321-4321-4321-4321-210987654321"
    UsePsadt = $false
}

# Run validation
Write-Host "`nRunning package validation..." -ForegroundColor Cyan
$results = Invoke-PackageValidation @testParams

Write-Host "`nValidation completed. Results:" -ForegroundColor Green
$results | ForEach-Object {
    Write-Host "- $($_.Test): $($_.Status) - $($_.Message)" -ForegroundColor $(if ($_.Status -eq "PASS") { "Green" } else { "Red" })
}