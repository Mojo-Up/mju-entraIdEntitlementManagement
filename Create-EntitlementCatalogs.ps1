<#	
    .NOTES
        ===========================================================================
         Created on:   	03/06/2025
         Author:        Jean-Pierre Simonis
         Version:   	1.0.0
         Organisation: 	Mojo Up
         Filename:      Create-EntitlementCatalogs.ps1
        ===========================================================================
    .DESCRIPTION
        This script creates entitlement catalogs in Microsoft Entra ID Entitlement Management using data from an Excel file.
        It requires the Microsoft.Graph.Identity.Governance and ImportExcel PowerShell modules.

        Excel column definitions:
        DisplayName – Name of the catalog
        Description – Description of the catalog

    .PARAMETER ExcelFile
        Specify the path to the Excel file containing Entra ID Entitlement Management catalog definitions.
        The file should contain the required columns as described in the .DESCRIPTION section.
    .EXAMPLE
        This example shows how to run the script with an Excel file containing catalog definitions.

        powershell.exe -exectuionpolicy bypass -file .\Create-EntitlementCatalogs.ps1 -ExcelFile ".\catalogs-example.xlsx"
    .EXAMPLE
        This example shows how to run the script with an Excel file containing catalog definitions with logging enabled.
        
        powershell.exe -exectuionpolicy bypass -file .\Create-EntitlementCatalogs.ps1 -ExcelFile ".\catalogs-example.xlsx" -LogtoFile
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFile,
    [switch]$LogtoFile = $false)

##########################
#        Variables       #
##########################

# Required columns in the Excel file
$requiredColumns = @("DisplayName", "Description")

#Logging variables
$logFile = "Create-EntitlementCatalogs.log"


##########################
#        Functions       #
##########################

function Check-Module {
    param(
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module $ModuleName..." -ForegroundColor Yellow
        Install-Module $ModuleName -Scope CurrentUser -Force
    }
    Write-Host "Importing module $ModuleName..." -ForegroundColor Cyan
    # Import the module, force re-import if already loaded
    Import-Module $ModuleName -Force
}

##########################
#        Execution       #
##########################

# Start logging if LogtoFile switch is set
if ($LogtoFile) {
    Start-Transcript -Path $logFile -Append
    Write-Host "Logging output to $logFile" -ForegroundColor Green
}


Write-Host "Loading required Modules..." -ForegroundColor Green
# Ensure required modules are installed and imported
Check-Module -ModuleName "Microsoft.Graph.Identity.Governance"
Check-Module -ModuleName "ImportExcel"

# Connect to Microsoft Graph if not already connected
if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
    Connect-MgGraph -Scopes "EntitlementManagement.ReadWrite.All" -NoWelcome
}

# Import Excel file as PowerShell object and validate required columns data
if (-not (Test-Path $ExcelFile)) {
    Write-Error "Excel file not found at path: $ExcelFile"
    exit 1
} else {
    
    # Get the Excel file and import it as a PowerShell object
    $catalogs = Import-Excel -Path $ExcelFile

    # Validate required columns
    foreach ($col in $requiredColumns) {
        if (-not $catalogs[0].PSObject.Properties.Name -contains $col) {
            Write-Error "Missing required column: $col"
            exit 1
        }
    }

}

# Create catalogs
Write-Host "Processing Excel file: $ExcelFile" -ForegroundColor Cyan

#Counter variables for progress tracking
$recordCount = $catalogs.Count
$recordIndex = 1

foreach ($catalog in $catalogs) {
    
    # Calculate percent complete for progress bar
    $percentComplete = [int](($recordIndex / $recordCount) * 100)
    Write-Progress -Activity "Processing Catalogs" -Status "Processing Record ($recordIndex/$recordCount)" -PercentComplete $percentComplete

    # Display current record being processed
    Write-Host "========================================" -ForegroundColor DarkGray
    Write-Host "Processing Record ($recordIndex/$recordCount)" -ForegroundColor Cyan

    # Check if catalog already exists
    $catalogexists = Get-MgEntitlementManagementCatalog -Filter "displayName eq '$($catalog.DisplayName)'"
    if (-not $catalogexists) {
        try {
            $newCatalog = New-MgEntitlementManagementCatalog -DisplayName $catalog.DisplayName `
                                               -Description $catalog.Description
            Write-Host "✅ Created catalog: $($catalog.DisplayName)"
        } catch {
            Write-Warning "❌ Failed to create catalog '$($catalog.DisplayName)': $($_.Exception.Message)"
        }
        
    } else {
        Write-Host "Catalog '$($catalog.DisplayName)' already exists. Skipping creation." -ForegroundColor Yellow
        continue
    }
$recordIndex++
}

# Disconnect from Microsoft Graph
Write-Host "All Entitlement Catalogs have been created... Disconnecting from MS Graph" -ForegroundColor Green
$LogoffGraph = Disconnect-Graph
if ($LogoffGraph) {
    Write-Host "Disconnected from Microsoft Graph successfully." -ForegroundColor Green
} else {
    Write-Warning "Failed to disconnect from Microsoft Graph."
}

# Stop logging if LogtoFile switch is set
if ($LogtoFile) {
    Stop-Transcript 
}