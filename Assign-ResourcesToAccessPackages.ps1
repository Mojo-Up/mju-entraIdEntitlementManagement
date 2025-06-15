
<#    
    .NOTES
        ===========================================================================
         Created on:    03/06/2025
         Author:        Jean-Pierre Simonis
         Version:       1.0.0
         Organisation:  Mojo Up
         Filename:      Assign-ResourcesToAccessPackages.ps1
        ===========================================================================
    .DESCRIPTION
        This script assigns resources to access packages in Microsoft Entra ID Entitlement Management using data from an Excel file.
        It requires the Microsoft.Graph and ImportExcel PowerShell modules.
        
        Excel columns definitions:
        AccessPackageName – Name of the access package
        ResourceName – Name of the resource (e.g., Group Display Name, Enterprise Application Display Name, SharePoint site URL)
        ResourceType – Type of resource (Group, Application, SharePoint)
        PermissionLevel – Role or permission level (e.g., Owner, Member, Reader)

    .PARAMETER ExcelFile
        Specify the path to the Excel file containing resource assignments.
        The file should have columns for AccessPackageName, ResourceName, ResourceType, and PermissionLevel.
    .EXAMPLE
        This example runs the script to assign resources to access packages using the specified Excel file without Logging to file
        
        powershell.exe -executionpolicy bypass -file .\Assign-ResourcesToAccessPackages.ps1 -ExcelFile ".\assignresourcestoaccesspackage-example.xlsx"
    .EXAMPLE
        This example runs the script to assign resources to access packages using the specified Excel file with Logging to file

        powershell.exe -executionpolicy bypass -file .\Assign-ResourcesToAccessPackages.ps1 -ExcelFile ".\assignresourcestoaccesspackage-example.xlsx" -LogtoFile
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFile,
    [switch]$LogtoFile = $false)

##########################
#        Variables       #
##########################

# Required columns in the Excel file
$requiredColumns = @("AccessPackageName", "ResourceName", "ResourceType", "PermissionLevel")

#Logging variables
$logFile = "Assign-ResourcesToAccessPackages.log"

##########################
#        Functions       #
##########################

# Function to check if a module is installed and import it
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

# Function to resolves resouces by name and type to a resource object
# Supported types: AccessPackage, Catalog, Group, Application, SharePoint
function Resolve-Resource {
    param(
        [string]$Name,
        [string]$Type
    )
    $result = $null
    switch ($Type) {
        "AccessPackage" {
            $result = Get-MgBetaEntitlementManagementAccessPackage -Filter "displayName eq '$Name'"
        }
        "Catalog" {
            $result = Get-MgEntitlementManagementCatalog -AccessPackageCatalogId $Name
        }
         "Group" {
            $result = Get-MgGroup -Filter "displayName eq '$Name'"
        }
        "Application" {
            $result = Get-MgServicePrincipal -Filter "displayName eq '$Name'"
        }
        "SharePoint" {
            # Assuming SharePoint site URL is provided in ResourceName
            $result = Get-MgSite -Search "$Name"
        }
        default {
            Write-Error "Unsupported resource type: $Type"
            return $null
        }
    }
    
return $result

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
# required because standard Microsoft.Graph module does not collect the catalogID from the access package
Check-Module -ModuleName "Microsoft.Graph.Beta.Identity.Governance"
Check-Module -ModuleName "ImportExcel"


# Connect to Microsoft Graph if not already connected
if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
    Connect-MgGraph -Scopes "EntitlementManagement.ReadWrite.All", "Group.Read.All", "User.Read.All, Application.Read.All, Sites.Read.All" -NoWelcome
}

# Import Excel file as PowerShell object and validate required columns data
if (-not (Test-Path $ExcelFile)) {
    Write-Error "Excel file not found at path: $ExcelFile"
    exit 1
} else {
    
    # Get the Excel file and import it as a PowerShell object
    $assignments = Import-Excel -Path $ExcelFile

    # Validate required columns
    foreach ($col in $requiredColumns) {
        if (-not $assignments[0].PSObject.Properties.Name -contains $col) {
            Write-Error "Missing required column: $col"
            exit 1
        }
    }

}

Write-Host "Processing Excel file: $ExcelFile" -ForegroundColor Cyan

#Counter variables for progress tracking
$recordCount = $assignments.Count
$recordIndex = 1

# Assign resources to access packages
foreach ($assignment in $assignments) {
    try {
        
        # Calculate percent complete for progress bar
        $percentComplete = [int](($recordIndex / $recordCount) * 100)
        Write-Progress -Activity "Processing Assignments" -Status "Processing Record ($recordIndex/$recordCount)" -PercentComplete $percentComplete
        
        # Display current record being processed
        Write-Host "========================================" -ForegroundColor DarkGray
        Write-Host "Processing Record ($recordIndex/$recordCount)" -ForegroundColor Cyan

        #Get the Access Package by name
        $accessPackage = Resolve-Resource -Name $assignment.AccessPackageName -Type "AccessPackage"
        if (-not $accessPackage) {
            Write-Warning "⚠️ Access package '$($assignment.AccessPackageName)' not found. Skipping entry."
            continue
        }

        # Get Catalog ID from target access package
        $CatalogId = $accessPackage.CatalogId

        # Get the Resource ID based on the ResourceName and ResourceType
        $resource = Resolve-Resource -Name $assignment.ResourceName -Type $assignment.ResourceType
        if (-not $resource) {
            Write-Warning "⚠️ Resource '$($assignment.ResourceName)' of type '$($assignment.ResourceType)' not found. Skipping entry."
            continue
        }
        # Get all resources assigned to the catalog
        $assignedResources = Get-MgEntitlementManagementCatalogResource -AccessPackageCatalogId $CatalogId -ExpandProperty "scopes" -All
        $CatalogName = (Resolve-Resource -Name $CatalogId -Type "Catalog").DisplayName
        Write-Host "Checking Catalog: $CatalogName for Resource: $($assignment.ResourceName) ($($assignment.ResourceType))" -ForegroundColor Green
        # Check if the resource is already assigned by display name
        $existingResource = $assignedResources | Where-Object { $_.DisplayName -eq $assignment.ResourceName -or $_.Description -eq $assignment.ResourceName }
        # If the resource is already assigned to catalog, skip the assignment otherwise add it
        if ($existingResource) {
            Write-Host " - ⚠️ Resource '$($assignment.ResourceName)' is already assigned to catalog '$($CatalogName)'. Skipping assignment." -ForegroundColor Yellow
        } else {
            # Add the Group as a resource to the catalog 
            If ($assignment.ResourceType -eq "Group") {

                # Parameters for adding group as a resource to the Catalog
                $GroupResourceAddParameters = @{
                    requestType = "adminAdd"
                    resource = @{
                        originId = $resource.Id
                        originSystem = "AadGroup"
                    }
                catalog = @{ id = $CatalogId }
                }
                # Create the Group resource in the entitlement management catalog
                $AssignGroupToCatalog = New-MgEntitlementManagementResourceRequest -BodyParameter $GroupResourceAddParameters
                
                Write-Host " - ✅ Assigned group '$($assignment.ResourceName)' to Catalog '$($CatalogName)'"        
            }

            # Add the Application as a resource to the catalog 
            If ($assignment.ResourceType -eq "Application") {

                # Parameters for adding application as a resource to the Catalog
                $ApplicationResourceAddParameters = @{
                    requestType = "adminAdd"
                    resource = @{
                        originId = $resource.Id
                        originSystem = "aadApplication"
                    }
                catalog = @{ id = $CatalogId }
                }
                # Create the Application in the entitlement management catalog
                $AssignApplicationToCatalog = New-MgEntitlementManagementResourceRequest -BodyParameter $ApplicationResourceAddParameters
                
                Write-Host " - ✅ Assigned application '$($assignment.ResourceName)' to Catalog '$($CatalogName)'"        
            }

            # Add the SharePoint SIte as a resource to the catalog 
            If ($assignment.ResourceType -eq "SharePoint") {

                # Parameters for adding the SharePoint site as a resource to the Catalog
                $SharePointResourceAddParameters = @{
                requestType = "adminAdd"
                resource = @{
                    originId = $assignment.ResourceName
                    originSystem = "SharePointOnline"
                }
                catalog = @{ id = $CatalogId }
                }
                # Create the SharePoint Site in the entitlement management catalog
                $AssignSPSiteToCatalog = New-MgEntitlementManagementResourceRequest -BodyParameter $SharePointResourceAddParameters
                
                Write-Host " - ✅ Assigned SharePoint Site '$($assignment.ResourceName)' to Catalog '$($CatalogName)'"        
            }
        
        }
        Write-Host "Assign Resource: $($assignment.ResourceName) ($($assignment.ResourceType)) to Access Package: $($assignment.AccessPackageName)" -ForegroundColor Magenta
        # Assign the resource to the access package with the specified permission level
        
        ## Assigning Group to Access Package
        If ($assignment.ResourceType -eq "Group") {
            # Get the Group as a resource from the Catalog
            $CatalogResources = Get-MgEntitlementManagementCatalogResource -AccessPackageCatalogId $CatalogId -ExpandProperty "scopes" -All
            $GroupResource = $CatalogResources | Where-Object OriginId -eq $resource.id
            $GroupResourceId = $GroupResource.id
            $GroupResourceScope = $GroupResource.Scopes[0]

            # Add the Group as a resource role to the Access Package
            $GroupResourceFilter = "(originSystem eq 'AadGroup' and resource/id eq '" + $GroupResourceId + "')"
            $GroupResourceRoles = Get-MgEntitlementManagementCatalogResourceRole -AccessPackageCatalogId $CatalogId -Filter $GroupResourceFilter -ExpandProperty "resource"
            $GroupMemberRole = $GroupResourceRoles | Where-Object DisplayName -eq "$($assignment.PermissionLevel)"
            
            # Parameters for adding the group as a resource to the Access Package
            $GroupResourceRoleScopeParameters = @{
                role = @{
                    displayName =  "$($assignment.PermissionLevel)"
                    description =  ""
                    originSystem =  $GroupMemberRole.OriginSystem
                    originId =  $GroupMemberRole.OriginId
                    resource = @{
                        id = $GroupResource.Id
                        originId = $GroupResource.OriginId
                        originSystem = $GroupResource.OriginSystem
                    }
                }
                scope = @{
                    id = $GroupResourceScope.Id
                    originId = $GroupResourceScope.OriginId
                    originSystem = $GroupResourceScope.OriginSystem
                }
            }
        
            $AssignGroupToAccessPackage = New-MgEntitlementManagementAccessPackageResourceRoleScope -AccessPackageId $AccessPackage.Id -BodyParameter $GroupResourceRoleScopeParameters
            Write-Host " - ✅ Assigned group '$($assignment.ResourceName)' to access package '$($assignment.AccessPackageName)' with permission level '$($assignment.PermissionLevel)'"
        }

        ## Assigning Application to Access Package
        If ($assignment.ResourceType -eq "Application") {
            # Get the Application as a resource from the Catalog
            $CatalogResources = Get-MgEntitlementManagementCatalogResource -AccessPackageCatalogId $CatalogId -ExpandProperty "scopes" -All
            $ApplicationResource = $CatalogResources | Where-Object { $_.OriginId -eq $resource.id }
            $ApplicationResourceId = $ApplicationResource.id
            $ApplicationResourceScope = $ApplicationResource.Scopes[0]

            # Add the Application as a resource role to the Access Package
            $ApplicationResourceFilter = "(originSystem eq 'AadApplication' and resource/id eq '" + $ApplicationResourceId + "')"
            $ApplicationResourceRoles = Get-MgEntitlementManagementCatalogResourceRole -AccessPackageCatalogId $CatalogId -Filter $ApplicationResourceFilter -All -ExpandProperty "resource"
            $ApplicationMemberRole = $ApplicationResourceRoles | Where-Object { $_.DisplayName -eq "$($assignment.PermissionLevel)" }
            
            # Parameters for adding the application as a resource to the Access Package
            $ApplicationResourceRoleScopeParameters = @{
            role = @{
                displayName = $ApplicationResource.DisplayName
                description = ""
                originSystem = $ApplicationMemberRole.OriginSystem
                originId = $ApplicationMemberRole.OriginId
                resource = @{
                id = $ApplicationResource.Id
                originId = $ApplicationResource.OriginId
                originSystem = $ApplicationResource.OriginSystem
                }
            }
            scope = @{
                id = $ApplicationResourceScope.Id
                originId = $ApplicationResourceScope.OriginId
                originSystem = $ApplicationResourceScope.OriginSystem
            }
            }
                    
            $AssignApplicationToAccessPackage = New-MgEntitlementManagementAccessPackageResourceRoleScope -AccessPackageId $AccessPackage.Id -BodyParameter $ApplicationResourceRoleScopeParameters
            Write-Host " - ✅ Assigned Application '$($assignment.ResourceName)' to access package '$($assignment.AccessPackageName)' with permission level '$($assignment.PermissionLevel)'"
        }

        ## Assigning SharePoint Site to Access Package
        If ($assignment.ResourceType -eq "SharePoint") {
            # Get the SharePoint Site as a resource from the Catalog
            $CatalogResources = Get-MgEntitlementManagementCatalogResource -AccessPackageCatalogId $CatalogId -ExpandProperty "scopes" -All
            $SPSiteResource = $CatalogResources | Where-Object { $_.OriginId -eq $assignment.ResourceName }
            $SPSiteResourceId = $SPSiteResource.id

            #Set the SharePoint permission level mapping based requested permission level
            # Note: SharePoint permission levels are typically "Visitors", "Members", "Owners"
            $SPSitePermissionValue = switch ($assignment.PermissionLevel) {
                "Vistors" { "4" }
                "Members" { "5" }
                "Owners"  { "3" }
                default   { "4" } # or handle as needed
            }

            # Parameters for adding the SharePoint Site as a resource to the Access Package
            $SPSiteResourceRoleScopeParameters = @{
                role = @{
                    displayName = $SPSiteResource.DisplayName
                    originSystem = "SharePointOnline"
                    originId = $SPSitePermissionValue
                    resource = @{
                        id = $SPSiteResourceId
                    }
                }
                scope = @{
                    displayName = "Root"
                    description = "Root Scope"
                    originId = $assignment.ResourceName
                    originSystem = "SharePointOnline"
                    isRootScope = $true
                }
            }

            $AssignSPSiteToAccessPackage = New-MgEntitlementManagementAccessPackageResourceRoleScope -AccessPackageId $AccessPackage.Id -BodyParameter $SPSiteResourceRoleScopeParameters
            Write-Host " - ✅ Assigned SharePoint Site '$($assignment.ResourceName)' to access package '$($assignment.AccessPackageName)' with permission level '$($assignment.PermissionLevel)'"
        }

    #increment the record index for progress tracking
    $recordIndex++
    } catch {
        Write-Warning "❌ Failed to assign resource '$($assignment.ResourceName)' to access package '$($assignment.AccessPackageName)': $($_.Exception.Message)"
    }
}
Write-Progress -Activity "Processing Assignments" -Completed

# Disconnect from Microsoft Graph
Write-Host "All resources have been processed and assigned to access packages... Disconnecting from MS Graph" -ForegroundColor Green
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