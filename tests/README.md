# Microsoft 365 Apps Package Validation Guide

This document explains how to use the new validation features in the Microsoft 365 Apps package script.

## New Parameters

### `-ValidateOnly`
Runs comprehensive validation without executing any actual operations. Perfect for testing configurations, CI/CD pipelines, and development.

### `-SkipImport`
Creates the package but skips the Intune import step. Useful when you want to create packages but import them manually or via different automation.

### `-WhatIf` (Built-in PowerShell)
Shows what operations would be performed without actually executing them.

## Usage Examples

### 1. Full Validation Only
```powershell
.\New-Microsoft365AppsPackage.ps1 -ValidateOnly `
    -ConfigurationFile ".\configs\O365ProPlus.xml" `
    -Channel "Current" `
    -CompanyName "MyCompany" `
    -TenantId "12345678-1234-1234-1234-123456789012"
```

### 2. Create Package Without Intune Import
```powershell
.\New-Microsoft365AppsPackage.ps1 -SkipImport `
    -ConfigurationFile ".\configs\O365ProPlus.xml" `
    -Channel "Current" `
    -CompanyName "MyCompany" `
    -TenantId "12345678-1234-1234-1234-123456789012" `
    -Import
```

### 3. WhatIf Mode (Shows what would happen)
```powershell
.\New-Microsoft365AppsPackage.ps1 -WhatIf `
    -ConfigurationFile ".\configs\O365ProPlus.xml" `
    -Channel "Current" `
    -CompanyName "MyCompany" `
    -TenantId "12345678-1234-1234-1234-123456789012"
```

### 4. Validation with PSADT
```powershell
.\New-Microsoft365AppsPackage.ps1 -ValidateOnly -UsePsadt `
    -ConfigurationFile ".\configs\O365ProPlus.xml" `
    -Channel "MonthlyEnterprise" `
    -CompanyName "MyCompany" `
    -TenantId "12345678-1234-1234-1234-123456789012"
```

## Validation Categories

### 📋 Parameter Validation
- Verifies all required paths exist
- Validates GUID formats for TenantId
- Checks XML configuration file structure
- Confirms required files are present

### 📝 XML Configuration Validation
- Simulates XML updates without modifying original files
- Validates channel settings
- Tests tenant ID and company name updates
- Ensures XML structure remains valid

### 📁 File Operations Validation
- Simulates directory structure creation
- Validates file copy operations
- Checks for missing dependencies
- Tests PSADT vs standard package structures

### 📦 Package Manifest Validation
- Validates manifest creation logic
- Tests display name generation
- Verifies detection rules
- Checks install/uninstall commands

## Validation Output

The validation provides detailed feedback:

```
🔍 Starting Microsoft 365 Apps Package Validation...
============================================================

📋 Parameter Validation
  ✅ Path Validation: Repository path: C:\path\to\m365apps
  ✅ Configuration File Validation: Config file: C:\path\to\config.xml
  ✅ XML Structure Validation: Valid Microsoft 365 Apps configuration XML
  ✅ Required File Check: File: C:\path\to\m365\setup.exe
  ✅ TenantId GUID Validation: TenantId: 12345678-1234-1234-1234-123456789012

📝 XML Configuration Validation
  ✅ XML Update Simulation: Successfully simulated XML updates

📁 File Operations Validation
  ✅ Directory Structure Validation: Would create directories at: C:\path\to\package
  ✅ File Copy Validation: Available: 3, Missing: 0

📦 Package Manifest Validation
  ✅ Package Manifest Validation: Successfully created manifest structure

============================================================
📊 Validation Summary:
   Total Tests: 8
   Passed: 8
   Failed: 0

✅ All validations passed! Package is ready for creation.
```

## CI/CD Integration

### GitHub Actions Example
```yaml
- name: Validate Microsoft 365 Apps Package
  shell: powershell
  run: |
    $params = @{
        ValidateOnly = $true
        ConfigurationFile = "${{ github.workspace }}\configs\O365ProPlus.xml"
        Channel = "Current"
        CompanyName = "MyCompany"
        TenantId = "${{ secrets.TENANT_ID }}"
    }
    $results = & "${{ github.workspace }}\New-Microsoft365AppsPackage.ps1" @params
    
    $failed = ($results | Where-Object Status -eq "FAIL").Count
    if ($failed -gt 0) {
        throw "Validation failed with $failed errors"
    }
```

### Azure DevOps Example
```yaml
- task: PowerShell@2
  displayName: 'Validate Package Configuration'
  inputs:
    targetType: 'inline'
    script: |
      $results = .\New-Microsoft365AppsPackage.ps1 -ValidateOnly `
        -ConfigurationFile "configs\O365ProPlus.xml" `
        -Channel "Current" `
        -CompanyName "MyCompany" `
        -TenantId "$(TenantId)"
      
      $failed = ($results | Where-Object Status -eq "FAIL").Count
      Write-Host "##vso[task.setvariable variable=ValidationFailed]$failed"
```

## Benefits

✅ **Fast Feedback** - Catch issues early without waiting for full execution  
✅ **No Authentication Required** - Test configurations without Intune credentials  
✅ **CI/CD Friendly** - Perfect for automated pipelines  
✅ **Development Safe** - Test changes without side effects  
✅ **Comprehensive Coverage** - Validates all major operations  
✅ **Clear Reporting** - Easy to understand pass/fail results