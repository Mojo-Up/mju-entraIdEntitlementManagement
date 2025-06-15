<#
    .NOTES
        ===========================================================================
         Created on:    03/06/2025
         Author:        Jean-Pierre Simonis
         Version:       1.0.0
         Organisation:  Mojo Up
         Filename:      Create-EntitlementAccessPackages.ps1
        ===========================================================================
    .DESCRIPTION
        This script creates Access Packages, Access Package Polices and assigns them to the desired Catalog
        in Microsoft Entra ID Entitlement Management using data from an Excel file.
        It requires the Microsoft.Graph.Identity.Governance and ImportExcel PowerShell modules.

        Excel column definitions:
        AccessPackageName – Name of the access package
        Description – Description of the access package
        CatalogName – Display name of the catalog (resolved to ID in script)
        PolicyName – Name of the assignment policy
        PolicyDescription – Description of the policy
        ApprovalEnabled – True or False, whether approval policy is required for access requests
        Approver - UPN of the user or Display Name of the group that will approve requests for this access package
        ApproverType - "user" or "group"
        EscalationApprover - UPN of the user or Display Name of the group that will approve escalation requests for this access package
        EscalationApproverType - "user" or "group" 
        DurationInDays – Leave blank for no expiration
        AutoAssignmentEnabled – True or False, whether auto-assignment policy is required for access requests
        DynamicMembershipRule – Dynamic membership rule for the group if AutoAssignmenEnabled (optional) eg (user.department -eq "Department X") 
        TargetGroupName – Group display name (resolved to ID in script)
        IsHidden – True or False (whether the access package is hidden from users)
        AccessReviews - True or False, if set to true, access reviews will be created for the access package.

    .PARAMETER ExcelFile
        Specify the path to the Excel file containing Entra ID Entitlement Management catalog definitions.
        The file should contain the required columns as described in the .DESCRIPTION section.
    .EXAMPLE
        This example shows how to run the script with an Excel file containing access package definitions.

        powershell.exe -executionpolicy bypass -file .\Create-EntitlementAccessPackages.ps1 -ExcelFile ".\accesspackages-example.xlsx"
    .EXAMPLE
        This example shows how to run the script with an Excel file containing access package definitions with logging enabled.
         
        powershell.exe -executionpolicy bypass -file .\Create-EntitlementAccessPackages.ps1 -ExcelFile ".\accesspackages-example.xlsx" -LogtoFile
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFile,
    [switch]$LogtoFile = $false)

##########################
#        Variables       #
##########################

# Required columns in the Excel file
$requiredColumns = @("AccessPackageName", "Description", "CatalogName", "PolicyName", "PolicyDescription", "ApprovalEnabled", "Approver", "ApproverType", "EscalationApprover", "EscalationApproverType", "DurationInDays", "AutoAssignmentEnabled", "DynamicMembershipRule", "TargetGroupName", "IsHidden", "AccessReviews")

#Logging variables
$logFile = "Create-EntitlementAccessPackages.log"

# Prefix for auto assignment policy names
$AutoPolicyNamePrefix = "Automatic Assignment Policy for" # Prefix for auto assignment policy names

# Access Review Settings
$accessReviewExpiration = "P14D" # Set to 14 days, adjust as needed
$recurrenceType = "absoluteMonthly" # or "monthly", "weekly", "daily"
[Int32]$recurrenceTypeInterval = 3    # every 1 week or month, adjust as needed
[Int32]$recurrenceTypeMonth = 0 # For monthly recurrence, specify the month (1-12)
[Int32]$recurrenceDayOfMonth = 0 # For absolute monthly recurrence, specify the day of the month (1-31)
$recurranceDaysOfWeek = @() # For weekly recurrence, specify the days of the week
$accessReviewExpirationBehavior = "keepAccess" # or "removeAccess" or "acceptAccessRecommendation"
$accessReviewRangeNumberOfOccurrences = 0 # Set to 0 occurrences, adjust as needed

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
    Connect-MgGraph -Scopes "EntitlementManagement.ReadWrite.All", "Group.Read.All", "User.Read.All" -NoWelcome
}

# Import Excel file as PowerShell object and validate required columns data
if (-not (Test-Path $ExcelFile)) {
    Write-Error "Excel file not found at path: $ExcelFile"
    exit 1
} else {
    
    # Get the Excel file and import it as a PowerShell object
    $policies = Import-Excel -Path $ExcelFile

    # Validate required columns
    foreach ($col in $requiredColumns) {
        if (-not $policies[0].PSObject.Properties.Name -contains $col) {
            Write-Error "Missing required column: $col"
            exit 1
        }
    }

}

Write-Host "Processing Excel file: $ExcelFile" -ForegroundColor Cyan

#Counter variables for progress tracking
$recordCount = $policies.Count
$recordIndex = 1


foreach ($entry in $policies) {

    # Calculate percent complete for progress bar
    $percentComplete = [int](($recordIndex / $recordCount) * 100)
    Write-Progress -Activity "Processing Access Packages" -Status "Processing Record ($recordIndex/$recordCount)" -PercentComplete $percentComplete

    # Display current record being processed
    Write-Host "========================================" -ForegroundColor DarkGray
    Write-Host "Processing Record ($recordIndex/$recordCount)" -ForegroundColor Cyan

    # Resolve group name to group ID
    $group = Get-MgGroup -Filter "displayName eq '$($entry.TargetGroupName)'" -ConsistencyLevel eventual
    if (-not $group) {
        Write-Warning "Group '$($entry.TargetGroupName)' not found. Skipping entry."
        continue
    }

    # Resolve catalog if it does not exist skip record
    $catalog = Get-MgEntitlementManagementCatalog -Filter "displayName eq '$($entry.CatalogName)'"
    if (-not $catalog) {
        Write-Warning "Catalog '$($entry.CatalogName)' not found. Skipping entry."
        continue
    }

    try {
            try {
                # Check if the access package already exists
                $existingAccessPackage = Get-MgEntitlementManagementAccessPackage -Filter "displayName eq '$($entry.AccessPackageName)'"

                if ($existingAccessPackage) {
                    Write-Host "⚠️ Access package '$($entry.AccessPackageName)' already exists in catalog '$($entry.CatalogName)'. Skipping creation." -ForegroundColor Yellow
                    $accessPackage = $existingAccessPackage
                } else {
                    try {
                        # Define the parameters for the new access package
                        $AccessPackageParameters = @{
                            displayName = $entry.AccessPackageName
                            description = $entry.Description
                            isHidden = $entry.IsHidden
                            catalog = @{
                                id = $catalog.Id
                            }
                        }

                        # Create the access package
                        Write-Host "✅ Creating new Access Package: $($entry.AccessPackageName) in Catalog: $($entry.CatalogName)" -ForegroundColor Green
                        $accessPackage = New-MgEntitlementManagementAccessPackage -BodyParameter $AccessPackageParameters
                    } catch {
                        Write-Error "❌ Failed to create access package: $($_.Exception.Message)"
                        continue
                    }
                } 
            

            If($entry.ApprovalEnabled -eq $true) {
                # Request policy parameters
                $AccessPackageId = $accessPackage.Id
                $RequestPolicyName = $entry.PolicyName
                $PolicyDescription = $entry.PolicyDescription
                $isAccessReviewRequired = If($entry.AccessReviews -eq "true"){$true} else {$false}
                $membershipRule = "allMemberUsers" # "allMemberUsers", "specificAllowedTargets", "allConfiguredConnectedOrganizationUsers", "notSpecified"
                ## Set Expiration block based on DurationInDays
                $expirationBlock = if ([string]::IsNullOrWhiteSpace($entry.DurationInDays)) {
                    @{
                        type = "noExpiration"
                    }
                } else {
                    @{
                        type = "afterDuration"
                        duration = "P$($entry.DurationInDays)D"
                    }
                }
                ## Check if ApproverType is valid and resolve Approver to Object ID
                If($entry.ApproverType -eq "User") {
                    $Approver = (Get-MgUser -Filter "userPrincipalName eq '$($entry.Approver)'").Id # Object ID of the user in Entra ID
                } elseif ($entry.ApproverType -eq "Group") {
                    $Approver = (Get-MgGroup -Filter "displayName eq '$($entry.Approver)'").Id # Object ID of the group in Entra ID
                } else {
                    Write-Error "❌ Invalid ApproverType specified. Use 'user' or 'group'."
                    continue
                }
                
                ## Check if Approver was found
                if (-not $Approver) {
                    Write-Warning "Approver '$($entry.Approver)' not found. Skipping entry."
                    continue
                }
                
                ## Set Approver block based on ApproverType and Approver
                $ApproverBlock = @(
                                    @{
                                        "@odata.type" = if ($entry.ApproverType -eq "group") { "#microsoft.graph.groupMembers" } else { "#microsoft.graph.singleUser" }
                                        userId = if ($entry.ApproverType -eq "user") { $Approver } else { $null }
                                        groupId = if ($entry.ApproverType -eq "group") { $Approver } else { $null }
                                    }
                                )
                
                ## Check if EscalationApproverType is valid and resolve EscalationApprover to Object ID
                if ($entry.EscalationApproverType -eq "User") {
                    $EscalationApprover = (Get-MgUser -Filter "userPrincipalName eq '$($entry.EscalationApprover)'").Id
                } elseif ($entry.EscalationApproverType -eq "Group") {
                    $EscalationApprover = (Get-MgGroup -Filter "displayName eq '$($entry.EscalationApprover)'").Id
                } elseif (![string]::IsNullOrWhiteSpace($entry.EscalationApprover)) {
                    Write-Error "❌ Invalid EscalationApproverType specified. Use 'user' or 'group'."
                    continue
                } else {
                    $EscalationApprover = $null
                }

                ## Check if EscalationApprover was found (if specified)
                if ($entry.EscalationApprover -and -not $EscalationApprover) {
                    Write-Warning "EscalationApprover '$($entry.EscalationApprover)' not found. Skipping entry."
                    continue
                }

                ## Set EscalationApprover block based on EscalationApproverType and EscalationApprover
                $EscalationApproverBlock = @()
                if ($EscalationApprover) {
                    $EscalationApproverBlock = @(
                        @{
                            "@odata.type" = if ($entry.EscalationApproverType -eq "Group") { "#microsoft.graph.groupMembers" } else { "#microsoft.graph.singleUser" }
                            userId = if ($entry.EscalationApproverType -eq "User") { $EscalationApprover } else { $null }
                            groupId = if ($entry.EscalationApproverType -eq "Group") { $EscalationApprover } else { $null }
                        }
                    )
                }

                $PrimaryReviewersBlock = @(
                            @{
                                "@odata.type" = if ($entry.ApproverType -eq "group") { "#microsoft.graph.groupMembers" } else { "#microsoft.graph.singleUser" }
                                userId = if ($entry.ApproverType -eq "user") { $Approver } else { $null }
                                groupId = if ($entry.ApproverType -eq "group") { $Approver } else { $null }
                            }
                        )
                
                # Request policy parameters for access package approval requests
                
                $RequestPolicyNameParameters = @{
                    displayName = $RequestPolicyName
                    description = $PolicyDescription
                    allowedTargetScope = $membershipRule
                
                    expiration = $expirationBlock

                    requestorSettings = @{
                        enableTargetsToSelfAddAccess = $true
                        enableTargetsToSelfUpdateAccess = $true
                        enableTargetsToSelfRemoveAccess = $true
                        allowCustomAssignmentSchedule = $true
                        enableOnBehalfRequestorsToAddAccess = $false
                        enableOnBehalfRequestorsToUpdateAccess = $false
                        enableOnBehalfRequestorsToRemoveAccess = $false
                        onBehalfRequestors = @(
                        )
                    }
                    requestApprovalSettings = @{
                        isApprovalRequiredForAdd = "true"
                        isApprovalRequiredForUpdate = "true"
                        stages = @(
                            @{
                                durationBeforeAutomaticDenial = "P7D" # 7 days
                                isApproverJustificationRequired = "false"
                                isEscalationEnabled = if ($EscalationApproverBlock.Count -gt 0) { "true" } else { "false" } 
                                fallbackPrimaryApprovers = @(
                                )
                                escalationApprovers = $EscalationApproverBlock
                                fallbackEscalationApprovers = @(
                                )
                                primaryApprovers = $ApproverBlock
                            }
                        )
                    }
                    reviewSettings = @{
                        isEnabled = $isAccessReviewRequired
                        expirationBehavior = $accessReviewExpirationBehavior
                        isRecommendationEnabled = $true
                        isReviewerJustificationRequired = $true
                        isSelfReview = $false
                        schedule = @{
                            startDateTime = [System.DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") # Set to current UTC date/time
                            #startDateTime = [System.DateTime]::Parse("2022-07-02T06:59:59.998Z")
                            expiration = @{
                                duration = $accessReviewExpiration
                                type = "afterDuration"
                            }
                            recurrence = @{
                                pattern = @{
                                    type = $recurrenceType
                                    interval = $recurrenceTypeInterval
                                    month = $recurrenceTypeMonth
                                    dayOfMonth = $recurrenceDayOfMonth
                                    daysOfWeek = $recurranceDaysOfWeek
                                }
                                range = @{
                                    type = "noEnd"
                                    numberOfOccurrences = $accessReviewRangeNumberOfOccurrences
                                }
                            }
                        }
                        primaryReviewers = $PrimaryReviewersBlock
                        fallbackReviewers = @(
                            )
                    }
                    accessPackage = @{
                        id = $AccessPackageId
                    }
                }
                # Check if the access package assignment policy name already exists)
                $existingApprovalPolicy = Get-MgEntitlementManagementAssignmentPolicy | Where-Object {
                    $_.DisplayName -eq $RequestPolicyName
                }


                if ($existingApprovalPolicy) {
                    # Skip policy creation if it already exists
                    Write-Host " - ⚠️ Assignment policy '$RequestPolicyName' already exists in catalog '$($catalog.DisplayName)'. Skipping creation." -ForegroundColor Yellow
                } else {
                    # Create Access Package Assignment Policy for Approval Requests
                    If($isAccessReviewRequired) { Write-Host " - Access Reviews enabled for this access package" -ForegroundColor Magenta} else { Write-Host " - Access Reviews not enabled for this access package" -ForegroundColor Magenta }
                    $NewApprovalPolicy = New-MgEntitlementManagementAssignmentPolicy -BodyParameter $RequestPolicyNameParameters
                    Write-Host " - ✅ Created access package policy: $($entry.PolicyName) associated with Access Package: $($entry.AccessPackageName)" -ForegroundColor Green
                }
            } else {
                Write-Host " - Approval policy not enabled for this access package" -ForegroundColor Magenta
            }

            If($entry.AutoAssignmentEnabled -eq $true) {
                
                # Auto assignment policy parameters
                $AccessPackageId = $accessPackage.Id
                $AutoPolicyName = "$($AutoPolicyNamePrefix) $($entry.AccessPackageName)"
                $AutoPolicyDescription = $entry.PolicyDescription 
                $AutoAssignmentPolicyFilter = $entry.DynamicMembershipRule 
                
                # Creating the auto assignment policy
                
                $AutoPolicyParameters = @{
                    DisplayName = $AutoPolicyName
                    Description = $AutoPolicyDescription
                    AllowedTargetScope = "specificDirectoryUsers"
                    SpecificAllowedTargets = @(
                        @{
                            "@odata.type" = "#microsoft.graph.attributeRuleMembers"
                            description = $AutoPolicyDescription
                            membershipRule = $AutoAssignmentPolicyFilter
                        }
                    )
                    AutomaticRequestSettings = @{
                        RequestAccessForAllowedTargets = $true
                        RemoveAccessWhenTargetLeavesAllowedTargets = $true
                    }
                    AccessPackage = @{
                        Id = $AccessPackageId
                    }
                }
                

                # Check if the access package assignment policy name already exists)
                $existingAutoAssignmentPolicy = Get-MgEntitlementManagementAssignmentPolicy | Where-Object {
                    $_.DisplayName -eq $AutoPolicyName
                }


                if ($existingAutoAssignmentPolicy) {
                    # Skip policy creation if it already exists
                    Write-Host " - ⚠️ Assignment policy '$AutoPolicyName' already exists in catalog '$($catalog.DisplayName)'. Skipping creation." -ForegroundColor Yellow
                } else {
                    # Create Access Package Auto Assignment Policy
                    $NewAutoPolicy = New-MgEntitlementManagementAssignmentPolicy -BodyParameter $AutoPolicyParameters
                    Write-Host " - ✅ Created auto assignment policy: $($AutoPolicyName) for Access Package: $($entry.AccessPackageName)" -ForegroundColor Green
                }


                
            } else {
                Write-Host " - Auto-Assignment policy not enabled for this access package" -ForegroundColor Magenta
            }
            
            Write-Host "In catalog: $($catalog.DisplayName) with ID: $($catalog.Id)" -ForegroundColor Cyan


    } catch {
        Write-Error "❌ Failed to create access package or policy: $($_.Exception.Message)"
        continue
    }
} catch {
        Write-Error "❌ An error occurred while processing entry: $($_.Exception.Message)"
        continue
    }
$recordIndex++
}

# Disconnect from Microsoft Graph
Write-Host "All access packages and policies have been processed... Disconnecting from MS Graph" -ForegroundColor Green
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