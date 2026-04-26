<#
.SYNOPSIS
    Azure Resilience & Disaster Recovery Assessment Script
    Extracts all Azure resources with zone redundancy information across all subscriptions.

.DESCRIPTION
    This enterprise-class script:
    - Loops through all accessible Azure subscriptions
    - Extracts all resources with their properties, tags, and zone information
    - Identifies logical and physical availability zone mappings
    - Determines zone redundancy status based on SKU and configuration
    - Exports data to individual CSV files per resource type AND a combined master file

.PARAMETER TenantId
    Optional. Specify a tenant ID to scope the assessment.

.PARAMETER OutputPath
    Optional. Specify output folder path. Defaults to current directory with timestamp.

.PARAMETER SubscriptionIds
    Optional. Array of specific subscription IDs to assess. If not provided, all accessible subscriptions are assessed.

.EXAMPLE
    .\Get-AzureZoneRedundancyAssessment.ps1
    
.EXAMPLE
    .\Get-AzureZoneRedundancyAssessment.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OutputPath "C:\Reports"

.NOTES
    Author: Azure Assessment Tool
    Version: 1.0
    Date: 2026-03-25
    Requires: Az.Accounts, Az.ResourceGraph modules
#>

#Requires -Modules Az.Accounts, Az.ResourceGraph

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionIds
)

#region Configuration
$ErrorActionPreference = "Continue"
$ProgressPreference = "Continue"

# Load System.Web for HTML encoding
Add-Type -AssemblyName System.Web

# Function to check and install required modules
function Install-RequiredModule {
    param(
        [string]$ModuleName,
        [switch]$Required
    )
    
    Write-Host "Checking for $ModuleName module..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
            Write-Host "$ModuleName module installed successfully." -ForegroundColor Green
            return $true
        }
        catch {
            if ($Required) {
                Write-Error "Failed to install required module $ModuleName. Error: $_"
                throw "Cannot continue without $ModuleName module."
            }
            else {
                Write-Warning "Failed to install $ModuleName module: $_"
                Write-Warning "Related functionality will be skipped."
                return $false
            }
        }
    }
    else {
        Write-Host "$ModuleName module found." -ForegroundColor Green
        return $true
    }
}

# Check and install all required modules
Write-Host "`n=== Checking Required Modules ===" -ForegroundColor Magenta

# Required modules (script cannot run without these)
$azAccountsOk = Install-RequiredModule -ModuleName "Az.Accounts" -Required
$azResourceGraphOk = Install-RequiredModule -ModuleName "Az.ResourceGraph" -Required

# Optional modules (enhance functionality)
$excelModuleOk = Install-RequiredModule -ModuleName "ImportExcel"

# Import modules
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.ResourceGraph -ErrorAction Stop
Import-Module ImportExcel -ErrorAction SilentlyContinue
$ExcelAvailable = $null -ne (Get-Module -Name ImportExcel)

Write-Host "=== Module Check Complete ===`n" -ForegroundColor Magenta

# Define resource types that support Availability Zones with their zone detection logic
# Based on Microsoft Documentation: https://learn.microsoft.com/en-us/azure/reliability/availability-zones-service-support
$ZoneAwareResourceTypes = @{
    #region Compute
    # VMs can only be Zonal (single zone) - Zone Redundancy requires VMSS or multiple VMs
    'Microsoft.Compute/virtualMachines' = @{ ZoneProperty = 'zones'; RedundancyType = 'Zonal'; SingleZoneOnly = $true }
    # VMSS can spread across zones
    'Microsoft.Compute/virtualMachineScaleSets' = @{ ZoneProperty = 'zones'; RedundancyType = 'ZoneSpread' }
    # Managed Disks - ZRS replicates across 3 zones
    'Microsoft.Compute/disks' = @{ SkuPattern = 'ZRS$'; RedundancyType = 'SKU-Based' }
    # Availability Sets - Fault domain isolation only, not zone-aware
    'Microsoft.Compute/availabilitySets' = @{ Default = 'LocallyRedundant'; RedundancyType = 'Default' }
    #endregion
    
    #region Storage
    # Storage Accounts - ZRS, GZRS, RA-GZRS provide zone redundancy
    # LRS = Locally Redundant, GRS = Geo Redundant (single zone in each region)
    'Microsoft.Storage/storageAccounts' = @{ SkuPattern = 'ZRS|GZRS|RAGZRS'; GeoPattern = 'GRS|RAGRS|GZRS|RAGZRS'; RedundancyType = 'SKU-Based' }
    #endregion
    
    #region Networking
    # Public IPs - Can be zonal (single zone) or zone-redundant (all 3 zones)
    'Microsoft.Network/publicIPAddresses' = @{ ZoneProperty = 'zones'; RedundancyType = 'Zonal' }
    # Load Balancers - Standard SKU supports zone redundancy
    'Microsoft.Network/loadBalancers' = @{ SkuName = 'Standard'; ZoneProperty = 'zones'; RedundancyType = 'SKU+Zone' }
    # Application Gateway v2 - Zone redundant when zones specified
    'Microsoft.Network/applicationGateways' = @{ ZoneProperty = 'zones'; SkuTier = 'Standard_v2|WAF_v2'; RedundancyType = 'Zonal' }
    # Azure Firewall - Zone redundant when zones specified
    'Microsoft.Network/azureFirewalls' = @{ ZoneProperty = 'zones'; RedundancyType = 'Zonal' }
    # VPN/ExpressRoute Gateway - AZ SKUs are zone redundant
    'Microsoft.Network/virtualNetworkGateways' = @{ SkuPattern = 'AZ$'; RedundancyType = 'SKU-Based' }
    # NAT Gateway - Zonal resource
    'Microsoft.Network/natGateways' = @{ ZoneProperty = 'zones'; RedundancyType = 'Zonal'; SingleZoneOnly = $true }
    # Bastion - Zone redundant when zones specified (Standard SKU)
    'Microsoft.Network/bastionHosts' = @{ ZoneProperty = 'zones'; RedundancyType = 'Zonal' }
    #endregion
    
    #region Databases
    # Azure SQL Database - Zone redundant is a configuration option
    'Microsoft.Sql/servers/databases' = @{ PropertyPath = 'properties.zoneRedundant'; RedundancyType = 'Property-Based' }
    # SQL Managed Instance - Zone redundant in Business Critical tier
    'Microsoft.Sql/managedInstances' = @{ PropertyPath = 'properties.zoneRedundant'; RedundancyType = 'Property-Based' }
    # PostgreSQL Flexible - HA mode ZoneRedundant
    'Microsoft.DBforPostgreSQL/flexibleServers' = @{ PropertyPath = 'properties.highAvailability.mode'; ValueMatch = 'ZoneRedundant'; RedundancyType = 'Property-Based' }
    # MySQL Flexible - HA mode ZoneRedundant
    'Microsoft.DBforMySQL/flexibleServers' = @{ PropertyPath = 'properties.highAvailability.mode'; ValueMatch = 'ZoneRedundant'; RedundancyType = 'Property-Based' }
    # Cosmos DB - isZoneRedundant per location
    'Microsoft.DocumentDB/databaseAccounts' = @{ PropertyPath = 'properties.locations[*].isZoneRedundant'; RedundancyType = 'Property-Based' }
    # Redis Cache - Premium tier with zones
    'Microsoft.Cache/redis' = @{ ZoneProperty = 'zones'; SkuFamily = 'P'; RedundancyType = 'Zonal' }
    'Microsoft.Cache/redisEnterprise' = @{ ZoneProperty = 'zones'; RedundancyType = 'Zonal' }
    #endregion
    
    #region App Services
    # App Service Plan - Zone redundant in Premium v2/v3 plans
    'Microsoft.Web/serverFarms' = @{ PropertyPath = 'properties.zoneRedundant'; RedundancyType = 'Property-Based' }
    # App Service/Function App - Inherits from hosting plan
    'Microsoft.Web/sites' = @{ InheritFrom = 'serverFarms'; RedundancyType = 'Inherited' }
    #endregion
    
    #region Containers
    # AKS - Zone redundancy via agent pool zones
    'Microsoft.ContainerService/managedClusters' = @{ PropertyPath = 'properties.agentPoolProfiles[*].availabilityZones'; RedundancyType = 'Property-Based' }
    # Container Registry - Premium tier with zone redundancy
    'Microsoft.ContainerRegistry/registries' = @{ PropertyPath = 'properties.zoneRedundancy'; ValueMatch = 'Enabled'; SkuName = 'Premium'; RedundancyType = 'Property-Based' }
    # Container Apps - Inherits from managed environment
    'Microsoft.App/containerApps' = @{ InheritFrom = 'managedEnvironments'; RedundancyType = 'Inherited' }
    'Microsoft.App/managedEnvironments' = @{ PropertyPath = 'properties.zoneRedundant'; RedundancyType = 'Property-Based' }
    #endregion
    
    #region Messaging
    # Service Bus Premium - Zone redundant by premium tier
    'Microsoft.ServiceBus/namespaces' = @{ PropertyPath = 'properties.zoneRedundant'; SkuName = 'Premium'; RedundancyType = 'Property-Based' }
    # Event Hubs Premium/Dedicated - Zone redundant
    'Microsoft.EventHub/namespaces' = @{ PropertyPath = 'properties.zoneRedundant'; RedundancyType = 'Property-Based' }
    # Event Grid - Zone redundant by default for domains/topics
    'Microsoft.EventGrid/domains' = @{ Default = 'ZoneRedundant'; RedundancyType = 'Default' }
    'Microsoft.EventGrid/topics' = @{ Default = 'ZoneRedundant'; RedundancyType = 'Default' }
    'Microsoft.EventGrid/systemTopics' = @{ Default = 'RedundantByDefault'; RedundancyType = 'Default' }
    #endregion
    
    #region Recovery Services
    # Recovery Services Vault - Storage redundancy setting
    'Microsoft.RecoveryServices/vaults' = @{ PropertyPath = 'properties.redundancySettings.standardTierStorageRedundancy'; RedundancyType = 'Property-Based' }
    #endregion
    
    #region Search & AI
    # Azure AI Search - Zone redundant with 2+ replicas in supported regions
    'Microsoft.Search/searchServices' = @{ PropertyPath = 'properties.replicaCount'; MinValue = 2; RedundancyType = 'Property-Based' }
    # Cognitive Services - Platform managed, zone redundant by default
    'Microsoft.CognitiveServices/accounts' = @{ Default = 'RedundantByDefault'; RedundancyType = 'Default' }
    #endregion
    
    #region SignalR
    # SignalR/WebPubSub - Premium tier for zone redundancy
    'Microsoft.SignalRService/signalR' = @{ SkuTier = 'Premium'; RedundancyType = 'SKU-Based' }
    'Microsoft.SignalRService/WebPubSub' = @{ SkuTier = 'Premium'; RedundancyType = 'SKU-Based' }
    #endregion
    
    #region API Management
    # APIM - Premium tier with zones or multiple units
    'Microsoft.ApiManagement/service' = @{ ZoneProperty = 'zones'; SkuName = 'Premium'; RedundancyType = 'SKU+Zone' }
    #endregion
}

# Resources that are platform-managed (zone redundant by default - no user configuration needed)
# These are control plane or metadata resources where Microsoft handles redundancy
$ZoneRedundantByDefault = @(
    # Networking - Control plane resources
    'Microsoft.Network/virtualNetworks',
    'Microsoft.Network/networkSecurityGroups',
    'Microsoft.Network/routeTables',
    'Microsoft.Network/privateEndpoints',
    'Microsoft.Network/privateLinkServices',
    'Microsoft.Network/networkInterfaces',
    'Microsoft.Network/networkWatchers',
    'Microsoft.Network/networkIntentPolicies',
    'Microsoft.Network/serviceEndpointPolicies',
    
    # DNS - Globally distributed
    'Microsoft.Network/dnsZones',
    'Microsoft.Network/privateDnsZones',
    'Microsoft.Network/privateDnsZones/virtualNetworkLinks',
    
    # Global Load Balancing - Globally distributed
    'Microsoft.Network/trafficManagerProfiles',
    'Microsoft.Network/frontDoors',
    'Microsoft.Cdn/profiles',
    
    # Logic Apps & Automation - Platform managed
    'Microsoft.Logic/workflows',
    'Microsoft.Automation/automationAccounts',
    
    # Identity - Globally replicated
    'Microsoft.ManagedIdentity/userAssignedIdentities',
    'Microsoft.Authorization/*',
    
    # Monitoring & Diagnostics - Platform managed
    'Microsoft.Insights/components',                       # Application Insights
    'Microsoft.Insights/dataCollectionRules',             # DCR
    'Microsoft.Insights/dataCollectionEndpoints',         # DCE
    'Microsoft.Insights/actionGroups',
    'Microsoft.OperationalInsights/workspaces',           # Log Analytics
    'Microsoft.OperationsManagement/solutions',
    'Microsoft.Monitor/accounts',                          # Azure Monitor workspace
    
    # Key Vault - Automatically replicated within region
    'Microsoft.KeyVault/vaults',
    
    # DevTest Lab schedules
    'Microsoft.DevTestLab/schedules',
    
    # Web connections (Logic Apps connectors)
    'Microsoft.Web/connections',
    
    # Certificate Orders - Global
    'Microsoft.CertificateRegistration/certificateOrders',
    
    # Hybrid resources - Agent-based, managed by Azure
    'Microsoft.HybridCompute/machines',
    'Microsoft.HybridCompute/machines/extensions',
    'Microsoft.HybridCompute/machines/licenseProfiles',
    'Microsoft.HybridCompute/licenses',
    
    # Azure Arc Data
    'Microsoft.AzureArcData/sqlServerInstances',
    'Microsoft.AzureArcData/sqlServerInstances/databases',
    'Microsoft.AzureArcData/sqlServerInstances/availabilityGroups',
    'Microsoft.AzureArcData/sqlServerLicenses',
    'Microsoft.AzureArcData/sqlServerEsuLicenses',
    
    # Azure Stack HCI - On-premises managed
    'Microsoft.AzureStackHCI/virtualHardDisks',
    'Microsoft.AzureStackHCI/networkInterfaces',
    
    # VM Extensions - Metadata
    'Microsoft.Compute/virtualMachines/extensions'
)

# Azure Paired Regions for Cross-Region DR Planning
# Source: https://learn.microsoft.com/en-us/azure/reliability/cross-region-replication-azure
$AzurePairedRegions = @{
    # Americas
    'eastus' = @{ PairedRegion = 'westus'; Geography = 'United States' }
    'eastus2' = @{ PairedRegion = 'centralus'; Geography = 'United States' }
    'westus' = @{ PairedRegion = 'eastus'; Geography = 'United States' }
    'westus2' = @{ PairedRegion = 'westcentralus'; Geography = 'United States' }
    'westus3' = @{ PairedRegion = 'eastus'; Geography = 'United States' }
    'centralus' = @{ PairedRegion = 'eastus2'; Geography = 'United States' }
    'northcentralus' = @{ PairedRegion = 'southcentralus'; Geography = 'United States' }
    'southcentralus' = @{ PairedRegion = 'northcentralus'; Geography = 'United States' }
    'westcentralus' = @{ PairedRegion = 'westus2'; Geography = 'United States' }
    'canadacentral' = @{ PairedRegion = 'canadaeast'; Geography = 'Canada' }
    'canadaeast' = @{ PairedRegion = 'canadacentral'; Geography = 'Canada' }
    'brazilsouth' = @{ PairedRegion = 'southcentralus'; Geography = 'Brazil' }
    'brazilsoutheast' = @{ PairedRegion = 'brazilsouth'; Geography = 'Brazil' }
    
    # Europe
    'northeurope' = @{ PairedRegion = 'westeurope'; Geography = 'Europe' }
    'westeurope' = @{ PairedRegion = 'northeurope'; Geography = 'Europe' }
    'uksouth' = @{ PairedRegion = 'ukwest'; Geography = 'United Kingdom' }
    'ukwest' = @{ PairedRegion = 'uksouth'; Geography = 'United Kingdom' }
    'francecentral' = @{ PairedRegion = 'francesouth'; Geography = 'France' }
    'francesouth' = @{ PairedRegion = 'francecentral'; Geography = 'France' }
    'germanywestcentral' = @{ PairedRegion = 'germanynorth'; Geography = 'Germany' }
    'germanynorth' = @{ PairedRegion = 'germanywestcentral'; Geography = 'Germany' }
    'switzerlandnorth' = @{ PairedRegion = 'switzerlandwest'; Geography = 'Switzerland' }
    'switzerlandwest' = @{ PairedRegion = 'switzerlandnorth'; Geography = 'Switzerland' }
    'norwayeast' = @{ PairedRegion = 'norwaywest'; Geography = 'Norway' }
    'norwaywest' = @{ PairedRegion = 'norwayeast'; Geography = 'Norway' }
    'swedencentral' = @{ PairedRegion = 'swedensouth'; Geography = 'Sweden' }
    'swedensouth' = @{ PairedRegion = 'swedencentral'; Geography = 'Sweden' }
    'polandcentral' = @{ PairedRegion = 'norwayeast'; Geography = 'Poland' }
    'italynorth' = @{ PairedRegion = 'francecentral'; Geography = 'Italy' }
    'spaincentral' = @{ PairedRegion = 'francecentral'; Geography = 'Spain' }
    
    # Middle East
    'uaenorth' = @{ PairedRegion = 'uaecentral'; Geography = 'UAE' }
    'uaecentral' = @{ PairedRegion = 'uaenorth'; Geography = 'UAE' }
    'qatarcentral' = @{ PairedRegion = $null; Geography = 'Qatar'; Note = 'No paired region - data residency region. Manual DR planning required to UAE or other regions.' }
    'israelcentral' = @{ PairedRegion = 'italynorth'; Geography = 'Israel' }
    
    # Asia Pacific
    'eastasia' = @{ PairedRegion = 'southeastasia'; Geography = 'Asia Pacific' }
    'southeastasia' = @{ PairedRegion = 'eastasia'; Geography = 'Asia Pacific' }
    'australiaeast' = @{ PairedRegion = 'australiasoutheast'; Geography = 'Australia' }
    'australiasoutheast' = @{ PairedRegion = 'australiaeast'; Geography = 'Australia' }
    'australiacentral' = @{ PairedRegion = 'australiacentral2'; Geography = 'Australia' }
    'australiacentral2' = @{ PairedRegion = 'australiacentral'; Geography = 'Australia' }
    'japaneast' = @{ PairedRegion = 'japanwest'; Geography = 'Japan' }
    'japanwest' = @{ PairedRegion = 'japaneast'; Geography = 'Japan' }
    'koreacentral' = @{ PairedRegion = 'koreasouth'; Geography = 'Korea' }
    'koreasouth' = @{ PairedRegion = 'koreacentral'; Geography = 'Korea' }
    'centralindia' = @{ PairedRegion = 'southindia'; Geography = 'India' }
    'southindia' = @{ PairedRegion = 'centralindia'; Geography = 'India' }
    'westindia' = @{ PairedRegion = 'southindia'; Geography = 'India' }
    'jioindiawest' = @{ PairedRegion = 'jioindiacentral'; Geography = 'India' }
    'jioindiacentral' = @{ PairedRegion = 'jioindiawest'; Geography = 'India' }
    
    # Africa
    'southafricanorth' = @{ PairedRegion = 'southafricawest'; Geography = 'South Africa' }
    'southafricawest' = @{ PairedRegion = 'southafricanorth'; Geography = 'South Africa' }
    
    # China (21Vianet)
    'chinaeast' = @{ PairedRegion = 'chinanorth'; Geography = 'China' }
    'chinanorth' = @{ PairedRegion = 'chinaeast'; Geography = 'China' }
    'chinaeast2' = @{ PairedRegion = 'chinanorth2'; Geography = 'China' }
    'chinanorth2' = @{ PairedRegion = 'chinaeast2'; Geography = 'China' }
    'chinaeast3' = @{ PairedRegion = 'chinanorth3'; Geography = 'China' }
    'chinanorth3' = @{ PairedRegion = 'chinaeast3'; Geography = 'China' }
}

# Resources that support cross-region replication/geo-redundancy
$CrossRegionCapableResources = @{
    'Microsoft.Storage/storageAccounts' = @{
        GeoRedundantSkus = @('Standard_GRS', 'Standard_RAGRS', 'Standard_GZRS', 'Standard_RAGZRS')
        CheckProperty = 'sku.name'
        ReplicationType = 'Automatic'
        Description = 'Geo-redundant storage replicates data to paired region'
    }
    'Microsoft.DocumentDB/databaseAccounts' = @{
        CheckProperty = 'properties.locations'
        ReplicationType = 'Multi-Region Write'
        Description = 'Cosmos DB supports multi-region replication'
    }
    'Microsoft.Sql/servers/databases' = @{
        CheckProperty = 'properties.secondaryType'
        ReplicationType = 'Geo-Replication / Failover Groups'
        Description = 'Azure SQL supports active geo-replication'
    }
    'Microsoft.RecoveryServices/vaults' = @{
        GeoRedundantSkus = @('GeoRedundant')
        CheckProperty = 'properties.redundancySettings.crossRegionRestore'
        ReplicationType = 'Cross-Region Restore'
        Description = 'Recovery Services Vault with cross-region restore'
    }
    'Microsoft.Cache/redis' = @{
        CheckProperty = 'properties.linkedServers'
        ReplicationType = 'Geo-Replication'
        Description = 'Redis Cache Premium supports geo-replication'
    }
    'Microsoft.ContainerRegistry/registries' = @{
        CheckProperty = 'properties.replicationLocations'
        ReplicationType = 'Geo-Replication'
        Description = 'Premium ACR supports geo-replication'
    }
}

#endregion

#region Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info' { 'White' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
    
    # Also log to file
    if ($script:LogFile) {
        "[$timestamp] [$Level] $Message" | Out-File -FilePath $script:LogFile -Append
    }
}

function Get-AllAzGraphResources {
    param (
        [string[]]$SubscriptionId,
        [string]$Query
    )
    
    try {
        if ($SubscriptionId) {
            $result = Search-AzGraph -Query $Query -First 1000 -Subscription $SubscriptionId -ErrorAction Stop
        } else {
            $result = Search-AzGraph -Query $Query -First 1000 -UseTenantScope -ErrorAction Stop
        }
        
        $allResources = @($result)
        
        # Paginate through results
        while ($result.SkipToken) {
            if ($SubscriptionId) {
                $result = Search-AzGraph -Query $Query -SkipToken $result.SkipToken -First 1000 -Subscription $SubscriptionId -ErrorAction Stop
            } else {
                $result = Search-AzGraph -Query $Query -SkipToken $result.SkipToken -First 1000 -UseTenantScope -ErrorAction Stop
            }
            $allResources += $result
        }
        
        return $allResources
    }
    catch {
        Write-Log "Error querying Azure Resource Graph: $_" -Level Error
        return @()
    }
}

function Get-ZoneMappings {
    param (
        [string]$SubscriptionId,
        [string]$SubscriptionName
    )
    
    $zoneMappings = @()
    
    try {
        $response = Invoke-AzRestMethod -Method GET -Path "/subscriptions/$SubscriptionId/locations?api-version=2022-12-01" -ErrorAction Stop
        
        if ($response.StatusCode -eq 200) {
            $locations = ($response.Content | ConvertFrom-Json).value
            
            foreach ($location in $locations) {
                if ($location.availabilityZoneMappings) {
                    foreach ($azMapping in $location.availabilityZoneMappings) {
                        $zoneMappings += [PSCustomObject]@{
                            SubscriptionName = $SubscriptionName
                            SubscriptionId = $SubscriptionId
                            Location = $location.name
                            DisplayName = $location.displayName
                            LogicalZone = $azMapping.logicalZone
                            PhysicalZone = $azMapping.physicalZone
                            SupportsAZ = $true
                        }
                    }
                } else {
                    $zoneMappings += [PSCustomObject]@{
                        SubscriptionName = $SubscriptionName
                        SubscriptionId = $SubscriptionId
                        Location = $location.name
                        DisplayName = $location.displayName
                        LogicalZone = 'N/A'
                        PhysicalZone = 'N/A'
                        SupportsAZ = $false
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error getting zone mappings for subscription $SubscriptionName : $_" -Level Warning
    }
    
    return $zoneMappings
}

function Get-ResourceZoneRedundancy {
    param (
        [PSObject]$Resource,
        [hashtable]$ZoneMappingsLookup
    )
    
    $resourceType = $Resource.type
    $zoneStatus = 'Unknown'
    $zoneDetail = ''
    $zones = @()
    $physicalZones = @()
    
    # Get zones array if exists
    if ($Resource.zones) {
        $zones = @($Resource.zones)
    }
    
    # Map logical zones to physical zones
    $locationKey = "$($Resource.subscriptionId)_$($Resource.location)"
    if ($zones.Count -gt 0 -and $ZoneMappingsLookup.ContainsKey($locationKey)) {
        foreach ($logicalZone in $zones) {
            $mapping = $ZoneMappingsLookup[$locationKey] | Where-Object { $_.LogicalZone -eq $logicalZone }
            if ($mapping) {
                $physicalZones += $mapping.PhysicalZone
            }
        }
    }
    
    # Check if resource type is zone redundant by default
    $isDefaultZR = $false
    foreach ($defaultType in $ZoneRedundantByDefault) {
        if ($resourceType -like $defaultType) {
            $isDefaultZR = $true
            break
        }
    }
    
    if ($isDefaultZR) {
        $zoneStatus = 'RedundantByDefault'
        $zoneDetail = 'Platform-managed redundancy'
    }
    # Check specific resource type configurations
    elseif ($ZoneAwareResourceTypes.ContainsKey($resourceType)) {
        $config = $ZoneAwareResourceTypes[$resourceType]
        
        # Zone property based detection
        if ($config.ZoneProperty) {
            if ($zones.Count -ge 2) {
                $zoneStatus = 'ZoneRedundant'
                $zoneDetail = "Deployed across zones: $($zones -join ', ')"
            } elseif ($zones.Count -eq 1) {
                $zoneStatus = 'Zonal'
                $zoneDetail = "Single zone: $($zones[0])"
            } else {
                # Resource supports zones but none configured - Regional/Non-Zonal deployment
                $zoneStatus = 'NonZonal'
                $zoneDetail = 'No availability zone configured (Regional deployment)'
            }
        }
        # SKU pattern based detection
        elseif ($config.SkuPattern) {
            $skuName = if ($Resource.sku.name) { $Resource.sku.name } else { $Resource.sku }
            if ($skuName -match $config.SkuPattern) {
                $zoneStatus = 'ZoneRedundant'
                $zoneDetail = "SKU: $skuName"
            } else {
                $zoneStatus = 'LocallyRedundant'
                $zoneDetail = "SKU: $skuName"
            }
        }
        # SKU tier based detection
        elseif ($config.SkuTier) {
            $skuTier = $Resource.sku.tier
            if ($skuTier -eq $config.SkuTier) {
                $zoneStatus = 'ZoneRedundant'
                $zoneDetail = "Tier: $skuTier"
            } else {
                $zoneStatus = 'NonZonal'
                $zoneDetail = "Tier: $skuTier"
            }
        }
        # Property path based detection
        elseif ($config.PropertyPath) {
            $propValue = $null
            try {
                $propPaths = $config.PropertyPath -split '\.'
                $current = $Resource
                foreach ($path in $propPaths) {
                    if ($path -match '\[\*\]$') {
                        $path = $path -replace '\[\*\]$', ''
                        $current = $current.$path
                        if ($current -is [array]) {
                            $propValue = $current | ForEach-Object { $_.$($propPaths[-1]) }
                        }
                        break
                    } else {
                        $current = $current.$path
                    }
                }
                if ($null -eq $propValue) { $propValue = $current }
            }
            catch {
                $propValue = $null
            }
            
            if ($config.ValueMatch) {
                if ($propValue -match $config.ValueMatch) {
                    $zoneStatus = 'ZoneRedundant'
                    $zoneDetail = "$($config.PropertyPath): $propValue"
                } else {
                    $zoneStatus = 'NonZonal'
                    $zoneDetail = "$($config.PropertyPath): $propValue"
                }
            }
            elseif ($config.MinValue) {
                if ($propValue -ge $config.MinValue) {
                    $zoneStatus = 'ZoneRedundant'
                    $zoneDetail = "$($config.PropertyPath): $propValue (>= $($config.MinValue))"
                } else {
                    $zoneStatus = 'NonZonal'
                    $zoneDetail = "$($config.PropertyPath): $propValue"
                }
            }
            elseif ($propValue -eq $true) {
                $zoneStatus = 'ZoneRedundant'
                $zoneDetail = "$($config.PropertyPath): Enabled"
            }
            elseif ($propValue -eq $false) {
                $zoneStatus = 'NonZonal'
                $zoneDetail = "$($config.PropertyPath): Disabled"
            }
            else {
                $zoneStatus = 'NonZonal'
                $zoneDetail = "$($config.PropertyPath): Not configured or null"
            }
        }
        # Inherited from parent resource (e.g., Web Sites inherit from App Service Plans)
        elseif ($config.InheritFrom) {
            $zoneStatus = 'InheritedFromParent'
            $zoneDetail = "Zone redundancy inherited from $($config.InheritFrom) - check parent resource"
        }
        # Default value
        elseif ($config.Default) {
            $zoneStatus = $config.Default
            $zoneDetail = 'Default configuration'
        }
        # Fallback for known zone-aware types with no matching config
        else {
            $zoneStatus = 'NonZonal'
            $zoneDetail = 'Zone-aware resource with no specific zone configuration'
        }
    }
    # No zones and not a known type
    elseif ($zones.Count -eq 0) {
        $zoneStatus = 'NonZonal'
        $zoneDetail = 'No zone configuration detected'
    }
    # Unclassified resource type that has a zones property set - use zones to determine status
    else {
        if ($zones.Count -ge 2) {
            $zoneStatus = 'ZoneRedundant'
            $zoneDetail = "Deployed across zones: $($zones -join ', ')"
        } elseif ($zones.Count -eq 1) {
            $zoneStatus = 'Zonal'
            $zoneDetail = "Single zone: $($zones[0])"
        } else {
            $zoneStatus = 'NonZonal'
            $zoneDetail = 'No zone configuration detected'
        }
    }

    # ── Physical Zone Placement Note for VMs ─────────────────────────────────
    # For Zonal VMs: physical zone is determined from zone mappings above.
    # For NonZonal VMs: zone placement is chosen by Azure scheduler — not deterministic.
    # The closest available indicators are Fault Domain + Update Domain (instanceView),
    # but instanceView is not available in Resource Graph. We surface this as guidance.
    $physicalZonePlacementNote = ''
    $vmTypes = @('microsoft.compute/virtualmachines', 'microsoft.compute/virtualmachinescalesets')
    if ($resourceType.ToLower() -in $vmTypes) {
        switch ($zoneStatus) {
            'Zonal' {
                $pz = if ($physicalZones.Count -gt 0) { $physicalZones -join ', ' } else { "Physical zone not in zone-mapping table for this subscription/region" }
                $physicalZonePlacementNote = "Logical Zone: $($zones -join ', ') → Physical Zone(s): $pz. VM is pinned to this specific datacenter."
            }
            'ZoneRedundant' {
                $physicalZonePlacementNote = "Zone-redundant deployment across logical zones: $($zones -join ', '). Physical datacenter mapping: $($physicalZones -join ', ')."
            }
            'NonZonal' {
                $physicalZonePlacementNote = "⚠️ Undetermined — VM deployed without specifying an Availability Zone. Azure places it in any available datacenter. Physical placement is not exposed via Resource Graph. To identify: (1) Azure Portal → VM → Properties → Fault Domain / Update Domain, or (2) run: az vm get-instance-view --ids <resourceId> --query 'instanceView.{FaultDomain:platformFaultDomain,UpdateDomain:platformUpdateDomain}'. This information is critical for planning zone-transition: VMs in the same resource group may share a physical zone, so a single zone outage could affect multiple 'non-zonal' VMs simultaneously."
            }
            default {
                $physicalZonePlacementNote = ''
            }
        }
    }
    
    return @{
        ZoneStatus               = $zoneStatus
        ZoneDetail               = $zoneDetail
        LogicalZones             = ($zones -join ', ')
        PhysicalZones            = ($physicalZones -join ', ')
        PhysicalZonePlacementNote = $physicalZonePlacementNote
    }
}

function Get-StorageAccountRedundancy {
    param([PSObject]$Resource)
    
    $skuName = if ($Resource.sku.name) { $Resource.sku.name } else { $Resource.sku }
    
    $redundancy = switch -Regex ($skuName) {
        'Standard_LRS|Premium_LRS' { 'LocallyRedundant' }
        'Standard_ZRS|Premium_ZRS' { 'ZoneRedundant' }
        'Standard_GRS|Standard_RAGRS' { 'GeoRedundant' }
        'Standard_GZRS|Standard_RAGZRS' { 'GeoZoneRedundant' }
        default { 'LocallyRedundant' }  # Assume LRS if SKU cannot be determined
    }
    
    return @{
        Redundancy = $redundancy
        SKU = $skuName
    }
}

function Get-DiskRedundancy {
    param([PSObject]$Resource)
    
    $skuName = if ($Resource.sku.name) { $Resource.sku.name } else { $Resource.sku }
    
    $redundancy = switch -Regex ($skuName) {
        '_LRS$' { 'LocallyRedundant' }
        '_ZRS$' { 'ZoneRedundant' }
        default { 'LocallyRedundant' }  # Assume LRS if SKU cannot be determined
    }
    
    return @{
        Redundancy = $redundancy
        SKU = $skuName
    }
}

# Cross-Region Assessment Function
function Get-CrossRegionStatus {
    param (
        [PSObject]$Resource,
        [hashtable]$PairedRegions
    )
    
    $location = $Resource.location
    $resourceType = $Resource.type
    $crossRegionStatus = 'Single-Region'
    $crossRegionDetail = ''
    $pairedRegion = ''
    $geoRedundant = $false
    $drReady = $false
    $isNoPairRegion = $false
    $isGlobalResource = $false
    
    # Handle global resources first
    if ($location -eq 'global' -or [string]::IsNullOrEmpty($location)) {
        $isGlobalResource = $true
        $crossRegionStatus = 'Global'
        $crossRegionDetail = 'Global resource - inherently multi-region'
        $drReady = $true
        $geoRedundant = $true
        $pairedRegion = 'N/A (Global)'
    }
    # Get paired region info for regional resources
    elseif ($PairedRegions.ContainsKey($location)) {
        $regionInfo = $PairedRegions[$location]
        $pairedRegion = $regionInfo.PairedRegion
        
        # Check if this is a no-pair region (PairedRegion is null)
        if ($null -eq $pairedRegion -or $pairedRegion -eq '') {
            $isNoPairRegion = $true
            $crossRegionDetail = "NO PAIRED REGION - Data residency region. Manual DR planning required."
            if ($regionInfo.Note) {
                $crossRegionDetail = $regionInfo.Note
            }
        } else {
            $crossRegionDetail = "Paired region: $pairedRegion"
            if ($regionInfo.Note) {
                $crossRegionDetail += " ($($regionInfo.Note))"
            }
        }
    } else {
        $crossRegionDetail = "Region '$location' - paired region info not available"
        $pairedRegion = 'Unknown'
    }
    
    # Skip resource type specific checks for global resources
    if ($isGlobalResource) {
        return @{
            CrossRegionStatus = $crossRegionStatus
            CrossRegionDetail = $crossRegionDetail
            PairedRegion = $pairedRegion
            GeoRedundant = $geoRedundant
            DRReady = $drReady
        }
    }
    
    # Check for geo-redundancy based on resource type
    switch -Wildcard ($resourceType) {
        'Microsoft.Storage/storageAccounts' {
            $skuName = if ($Resource.sku.name) { $Resource.sku.name } else { $Resource.sku }
            if ($skuName -match 'GRS|RAGRS|GZRS|RAGZRS') {
                if ($isNoPairRegion) {
                    $geoRedundant = $false
                    $crossRegionStatus = 'Single-Region'
                    $crossRegionDetail = "Storage SKU is $skuName but region has NO paired region. Manual DR required."
                } else {
                    $geoRedundant = $true
                    $crossRegionStatus = 'Geo-Redundant'
                    $crossRegionDetail = "Storage replicates to paired region ($pairedRegion) via $skuName"
                    $drReady = $true
                }
            } else {
                $crossRegionStatus = 'Single-Region'
                if ($isNoPairRegion) {
                    $crossRegionDetail = "Storage is LRS/ZRS in NO-PAIR region. Manual DR to another region required."
                } else {
                    $crossRegionDetail = "Storage is not geo-redundant ($skuName). Consider GRS/GZRS for DR to $pairedRegion."
                }
            }
        }
        'Microsoft.DocumentDB/databaseAccounts' {
            $locations = $Resource.properties.locations
            if ($locations -and $locations.Count -gt 1) {
                $regionNames = ($locations | ForEach-Object { $_.locationName }) -join ', '
                $crossRegionStatus = 'Multi-Region'
                $crossRegionDetail = "Cosmos DB deployed in: $regionNames"
                $drReady = $true
                $geoRedundant = $true
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Single region Cosmos DB. Consider multi-region for DR."
            }
        }
        'Microsoft.Sql/servers/databases' {
            # Check for geo-replication (would need additional API call to fully verify)
            $crossRegionStatus = 'Unknown'
            $crossRegionDetail = "Check Azure Portal for active geo-replication / failover groups"
        }
        'Microsoft.RecoveryServices/vaults' {
            $redundancy = $Resource.properties.redundancySettings.standardTierStorageRedundancy
            $crossRegionRestore = $Resource.properties.redundancySettings.crossRegionRestore
            if ($redundancy -eq 'GeoRedundant' -or $crossRegionRestore -eq 'Enabled') {
                $geoRedundant = $true
                $crossRegionStatus = 'Geo-Redundant'
                $crossRegionDetail = "Recovery Services Vault with geo-redundancy"
                $drReady = $true
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Enable GRS and Cross-Region Restore for DR"
            }
        }
        'Microsoft.ContainerRegistry/registries' {
            $replications = $Resource.properties.replicationLocations
            if ($replications -and $replications.Count -gt 1) {
                $crossRegionStatus = 'Geo-Redundant'
                $crossRegionDetail = "ACR replicated to multiple regions"
                $geoRedundant = $true
                $drReady = $true
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Single region ACR. Consider geo-replication for DR."
            }
        }
        'Microsoft.Cache/redis' {
            if ($Resource.properties.linkedServers -and $Resource.properties.linkedServers.Count -gt 0) {
                $crossRegionStatus = 'Geo-Redundant'
                $crossRegionDetail = "Redis with geo-replication configured"
                $geoRedundant = $true
                $drReady = $true
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Redis without geo-replication. Consider Premium tier with geo-replication for DR."
            }
        }
        'Microsoft.ServiceBus/namespaces' {
            $sku = $Resource.sku.tier
            if ($sku -eq 'Premium') {
                $crossRegionStatus = 'Unknown'
                $crossRegionDetail = "Service Bus Premium - check for Geo-DR configuration in portal"
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Service Bus $sku tier. Upgrade to Premium for Geo-DR capability."
            }
        }
        'Microsoft.EventHub/namespaces' {
            $sku = $Resource.sku.tier
            if ($sku -eq 'Premium' -or $sku -eq 'Standard') {
                $crossRegionStatus = 'Unknown'
                $crossRegionDetail = "Event Hub $sku - check for Geo-DR configuration in portal"
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Event Hub $sku tier does not support Geo-DR."
            }
        }
        'Microsoft.ApiManagement/service' {
            $locations = $Resource.properties.additionalLocations
            if ($locations -and $locations.Count -gt 0) {
                $regionNames = ($locations | ForEach-Object { $_.location }) -join ', '
                $crossRegionStatus = 'Multi-Region'
                $crossRegionDetail = "API Management deployed in multiple regions: $($Resource.location), $regionNames"
                $drReady = $true
                $geoRedundant = $true
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "API Management single region. Consider Premium tier multi-region."
            }
        }
        'Microsoft.Search/searchServices' {
            $replicaCount = $Resource.properties.replicaCount
            if ($replicaCount -gt 1) {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Azure Search with $replicaCount replicas (high availability within region). Deploy to paired region for DR."
            } else {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Azure Search single replica. Add replicas for HA, deploy to paired region for DR."
            }
        }
        'Microsoft.Compute/virtualMachines' {
            $crossRegionStatus = 'Single-Region'
            if ($isNoPairRegion) {
                $crossRegionDetail = "VM requires Azure Site Recovery (ASR) - NO PAIRED REGION. Select DR target manually (e.g., UAE North, West Europe)."
            } else {
                $crossRegionDetail = "VM requires Azure Site Recovery (ASR) for cross-region DR to $pairedRegion"
            }
        }
        'Microsoft.KeyVault/vaults' {
            $crossRegionStatus = 'Single-Region'
            if ($isNoPairRegion) {
                $crossRegionDetail = "Key Vault - NO PAIRED REGION. Use Azure Backup or manual replication to chosen DR region."
            } else {
                $crossRegionDetail = "Key Vault - use Azure Backup or manual replication to paired region ($pairedRegion)"
            }
        }
        'Microsoft.Web/sites' {
            $crossRegionStatus = 'Single-Region'
            if ($isNoPairRegion) {
                $crossRegionDetail = "App Service - NO PAIRED REGION. Deploy to multiple regions with Traffic Manager/Front Door."
            } else {
                $crossRegionDetail = "App Service - deploy to $pairedRegion with Traffic Manager/Front Door for DR"
            }
        }
        'Microsoft.Web/serverFarms' {
            $crossRegionStatus = 'Single-Region'
            if ($isNoPairRegion) {
                $crossRegionDetail = "App Service Plan - NO PAIRED REGION. Create matching plan in chosen DR region."
            } else {
                $crossRegionDetail = "App Service Plan - create matching plan in $pairedRegion for DR"
            }
        }
        # Global services (inherently multi-region)
        'Microsoft.Network/trafficManagerProfiles' {
            $crossRegionStatus = 'Global'
            $crossRegionDetail = 'Traffic Manager provides global load balancing across regions'
            $drReady = $true
            $geoRedundant = $true
        }
        'Microsoft.Cdn/profiles' {
            $crossRegionStatus = 'Global'
            $crossRegionDetail = 'Azure CDN provides global content delivery'
            $drReady = $true
            $geoRedundant = $true
        }
        'Microsoft.Network/frontDoors' {
            $crossRegionStatus = 'Global'
            $crossRegionDetail = 'Azure Front Door provides global load balancing'
            $drReady = $true
            $geoRedundant = $true
        }
        default {
            # For other resources, provide appropriate cross-region guidance
            if ($isNoPairRegion) {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Single region in NO-PAIR region. Manual DR to chosen region (e.g., UAE North, West Europe) required."
            }
            elseif ($pairedRegion -and $pairedRegion -ne '' -and $pairedRegion -ne 'Unknown' -and $pairedRegion -ne 'N/A (Global)') {
                $crossRegionStatus = 'Single-Region'
                $crossRegionDetail = "Single region deployment. For DR, deploy to paired region: $pairedRegion"
            }
            else {
                $crossRegionStatus = 'Single-Region'
                # Keep existing detail if set, otherwise provide generic message
                if ([string]::IsNullOrEmpty($crossRegionDetail)) {
                    $crossRegionDetail = "Single region deployment. Check Azure documentation for DR options."
                }
            }
        }
    }
    
    return @{
        CrossRegionStatus = $crossRegionStatus
        CrossRegionDetail = $crossRegionDetail
        PairedRegion = $pairedRegion
        GeoRedundant = $geoRedundant
        DRReady = $drReady
    }
}

function Export-ResourceTypeCSV {
    param(
        [string]$ResourceType,
        [array]$Resources,
        [string]$OutputFolder
    )
    
    if ($Resources.Count -eq 0) { return }
    
    try {
        # Sanitize resource type for filename
        $fileName = $ResourceType -replace '/', '_' -replace '\.', '_'
        $filePath = Join-Path $OutputFolder "$fileName.csv"
        
        # Get all unique properties across all resources of this type
        $allProps = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($resource in $Resources) {
            foreach ($prop in $resource.PSObject.Properties.Name) {
                [void]$allProps.Add($prop)
            }
        }
        
        $Resources | Select-Object -Property @($allProps) | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        Write-Log "Exported $($Resources.Count) resources to $fileName.csv" -Level Info
    }
    catch {
        Write-Log "Error exporting $ResourceType : $_" -Level Error
    }
}

function Flatten-Object {
    param(
        [PSObject]$Object,
        [string]$Prefix = ''
    )
    
    $result = @{}
    
    foreach ($prop in $Object.PSObject.Properties) {
        $key = if ($Prefix) { "${Prefix}_$($prop.Name)" } else { $prop.Name }
        
        if ($prop.Value -is [PSObject] -and $prop.Value.PSObject.Properties.Count -gt 0 -and $prop.Name -notin @('tags', 'properties')) {
            $nested = Flatten-Object -Object $prop.Value -Prefix $key
            foreach ($nestedKey in $nested.Keys) {
                $result[$nestedKey] = $nested[$nestedKey]
            }
        }
        elseif ($prop.Value -is [array]) {
            $result[$key] = ($prop.Value | ConvertTo-Json -Compress -Depth 3)
        }
        else {
            $result[$key] = $prop.Value
        }
    }
    
    return $result
}

#endregion

#region Main Script

# Initialize output folder
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) "AzureZoneAssessment_$timestamp"
}

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$script:LogFile = Join-Path $OutputPath "Assessment_Log.txt"
$resourceTypesFolder = Join-Path $OutputPath "ResourceTypes"
New-Item -ItemType Directory -Path $resourceTypesFolder -Force | Out-Null

Write-Log "=" * 60 -Level Info
Write-Log "Azure Resilience & DR Assessment Started" -Level Success
Write-Log "Output Path: $OutputPath" -Level Info
Write-Log "=" * 60 -Level Info

# Verify Azure connection
if (-not (Get-AzContext)) {
    Write-Log "Not connected to Azure. Please run Connect-AzAccount first." -Level Error
    exit 1
}

# Get subscriptions
Write-Log "Retrieving subscriptions..." -Level Info
try {
    if ($SubscriptionIds) {
        $subscriptions = $SubscriptionIds | ForEach-Object { Get-AzSubscription -SubscriptionId $_ -ErrorAction SilentlyContinue }
    }
    elseif ($TenantId) {
        $subscriptions = Get-AzSubscription -TenantId $TenantId | Where-Object { $_.State -eq 'Enabled' }
    }
    else {
        $subscriptions = Get-AzSubscription | Where-Object { $_.State -eq 'Enabled' }
    }
}
catch {
    Write-Log "Error retrieving subscriptions: $_" -Level Error
    exit 1
}

$totalSubs = $subscriptions.Count
Write-Log "Found $totalSubs enabled subscription(s)" -Level Success

if ($totalSubs -eq 0) {
    Write-Log "No subscriptions found. Exiting." -Level Error
    exit 1
}

# Initialize collections
$allResources = @()
$allZoneMappings = @()
$resourcesByType = @{}
$subCounter = 0

# Process each subscription
foreach ($subscription in $subscriptions) {
    $subCounter++
    Write-Log "-" * 40 -Level Info
    Write-Log "Processing [$subCounter/$totalSubs]: $($subscription.Name) ($($subscription.Id))" -Level Info
    
    try {
        # Set context
        Set-AzContext -SubscriptionId $subscription.Id -ErrorAction Stop | Out-Null
        
        # Get zone mappings for this subscription
        Write-Log "  Getting availability zone mappings..." -Level Info
        $zoneMappings = Get-ZoneMappings -SubscriptionId $subscription.Id -SubscriptionName $subscription.Name
        $allZoneMappings += $zoneMappings
        
        # Create lookup table for zone mappings
        $zoneMappingsLookup = @{}
        foreach ($zm in $zoneMappings) {
            $key = "$($zm.SubscriptionId)_$($zm.Location)"
            if (-not $zoneMappingsLookup.ContainsKey($key)) {
                $zoneMappingsLookup[$key] = @()
            }
            $zoneMappingsLookup[$key] += $zm
        }
        
        # Query all resources using Azure Resource Graph
        Write-Log "  Querying resources via Azure Resource Graph..." -Level Info
        
        $query = @"
Resources
| project id, name, type, kind, location, resourceGroup, subscriptionId, 
          tenantId, sku, plan, zones, identity, tags, properties, 
          managedBy, extendedLocation
| extend ResourceId = id
"@
        
        $resources = Get-AllAzGraphResources -SubscriptionId $subscription.Id -Query $query
        Write-Log "  Found $($resources.Count) resources" -Level Info
        
        # Process each resource
        foreach ($resource in $resources) {
            # Add subscription name
            $resource | Add-Member -NotePropertyName 'SubscriptionName' -NotePropertyValue $subscription.Name -Force
            
            # Extract resource type components
            $typeParts = $resource.type -split '/'
            $resource | Add-Member -NotePropertyName 'ResourceProvider' -NotePropertyValue $typeParts[0] -Force
            $resource | Add-Member -NotePropertyName 'ResourceTypeName' -NotePropertyValue ($typeParts[1..($typeParts.Length-1)] -join '/') -Force
            
            # Get zone redundancy status
            $zoneInfo = Get-ResourceZoneRedundancy -Resource $resource -ZoneMappingsLookup $zoneMappingsLookup
            $resource | Add-Member -NotePropertyName 'ZoneRedundancyStatus' -NotePropertyValue $zoneInfo.ZoneStatus -Force
            $resource | Add-Member -NotePropertyName 'ZoneRedundancyDetail' -NotePropertyValue $zoneInfo.ZoneDetail -Force
            $resource | Add-Member -NotePropertyName 'LogicalZones' -NotePropertyValue $zoneInfo.LogicalZones -Force
            $resource | Add-Member -NotePropertyName 'PhysicalZones' -NotePropertyValue $zoneInfo.PhysicalZones -Force
            $resource | Add-Member -NotePropertyName 'PhysicalZonePlacementNote' -NotePropertyValue $zoneInfo.PhysicalZonePlacementNote -Force
            
            # Get cross-region status
            $crossRegionInfo = Get-CrossRegionStatus -Resource $resource -PairedRegions $AzurePairedRegions
            $resource | Add-Member -NotePropertyName 'CrossRegionStatus' -NotePropertyValue $crossRegionInfo.CrossRegionStatus -Force
            $resource | Add-Member -NotePropertyName 'CrossRegionDetail' -NotePropertyValue $crossRegionInfo.CrossRegionDetail -Force
            $resource | Add-Member -NotePropertyName 'PairedRegion' -NotePropertyValue $crossRegionInfo.PairedRegion -Force
            $resource | Add-Member -NotePropertyName 'GeoRedundant' -NotePropertyValue $crossRegionInfo.GeoRedundant -Force
            $resource | Add-Member -NotePropertyName 'DRReady' -NotePropertyValue $crossRegionInfo.DRReady -Force
            
            # Special handling for Storage Accounts
            if ($resource.type -eq 'Microsoft.Storage/storageAccounts') {
                $storageInfo = Get-StorageAccountRedundancy -Resource $resource
                $resource | Add-Member -NotePropertyName 'StorageRedundancy' -NotePropertyValue $storageInfo.Redundancy -Force
                $resource | Add-Member -NotePropertyName 'StorageSKU' -NotePropertyValue $storageInfo.SKU -Force
                $resource.ZoneRedundancyStatus = $storageInfo.Redundancy
                $resource.ZoneRedundancyDetail = "SKU: $($storageInfo.SKU)"
            }
            
            # Special handling for Managed Disks
            if ($resource.type -eq 'Microsoft.Compute/disks') {
                $diskInfo = Get-DiskRedundancy -Resource $resource
                $resource | Add-Member -NotePropertyName 'DiskRedundancy' -NotePropertyValue $diskInfo.Redundancy -Force
                $resource | Add-Member -NotePropertyName 'DiskSKU' -NotePropertyValue $diskInfo.SKU -Force
                $resource.ZoneRedundancyStatus = $diskInfo.Redundancy
                $resource.ZoneRedundancyDetail = "SKU: $($diskInfo.SKU)"
            }
            
            # Extract all tags as separate columns
            if ($resource.tags) {
                foreach ($tagProp in $resource.tags.PSObject.Properties) {
                    $tagName = "Tag_$($tagProp.Name)" -replace '[^a-zA-Z0-9_]', '_'
                    $resource | Add-Member -NotePropertyName $tagName -NotePropertyValue $tagProp.Value -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Extract key properties
            if ($resource.properties) {
                # Common properties to extract
                $propsToExtract = @(
                    'provisioningState', 'creationTime', 'timeCreated', 'status', 'state',
                    'vmSize', 'osType', 'version', 'tier', 'capacity', 'size',
                    'zoneRedundant', 'highAvailability', 'replicaCount', 'redundancyMode',
                    'publicNetworkAccess', 'privateEndpointConnections', 'encryption',
                    'backup', 'geoReplication', 'failoverPolicy', 'replicationPolicy'
                )
                
                foreach ($propName in $propsToExtract) {
                    if ($resource.properties.$propName) {
                        $resource | Add-Member -NotePropertyName "Prop_$propName" -NotePropertyValue $resource.properties.$propName -Force -ErrorAction SilentlyContinue
                    }
                }
                
                # Extract dependency information
                $dependencies = @()
                
                # Network dependencies
                if ($resource.properties.networkProfile) {
                    foreach ($nic in $resource.properties.networkProfile.networkInterfaces) {
                        if ($nic.id) { $dependencies += "NIC:$($nic.id)" }
                    }
                }
                if ($resource.properties.ipConfigurations) {
                    foreach ($ipConfig in $resource.properties.ipConfigurations) {
                        if ($ipConfig.properties.subnet.id) { $dependencies += "SUBNET:$($ipConfig.properties.subnet.id)" }
                    }
                }
                if ($resource.properties.subnet) {
                    if ($resource.properties.subnet.id) { $dependencies += "SUBNET:$($resource.properties.subnet.id)" }
                }
                if ($resource.properties.virtualNetwork) {
                    if ($resource.properties.virtualNetwork.id) { $dependencies += "VNET:$($resource.properties.virtualNetwork.id)" }
                }
                
                # Storage dependencies
                if ($resource.properties.storageProfile) {
                    if ($resource.properties.storageProfile.osDisk.managedDisk.id) {
                        $dependencies += "DISK:$($resource.properties.storageProfile.osDisk.managedDisk.id)"
                    }
                    foreach ($disk in $resource.properties.storageProfile.dataDisks) {
                        if ($disk.managedDisk.id) { $dependencies += "DISK:$($disk.managedDisk.id)" }
                    }
                }
                if ($resource.properties.storageAccount) {
                    if ($resource.properties.storageAccount.id) { $dependencies += "STORAGE:$($resource.properties.storageAccount.id)" }
                }
                
                # Key Vault dependencies
                if ($resource.properties.vaultUri) {
                    $dependencies += "KEYVAULT:$($resource.properties.vaultUri)"
                }
                if ($resource.properties.keyVaultProperties) {
                    if ($resource.properties.keyVaultProperties.keyVaultUri) {
                        $dependencies += "KEYVAULT:$($resource.properties.keyVaultProperties.keyVaultUri)"
                    }
                }
                
                # Managed Identity dependencies
                if ($resource.identity) {
                    if ($resource.identity.principalId) {
                        $dependencies += "IDENTITY:$($resource.identity.principalId)"
                    }
                    if ($resource.identity.userAssignedIdentities) {
                        foreach ($uai in $resource.identity.userAssignedIdentities.PSObject.Properties) {
                            $dependencies += "IDENTITY:$($uai.Name)"
                        }
                    }
                }
                
                # Parent resource dependencies
                if ($resource.managedBy) {
                    $dependencies += "MANAGED_BY:$($resource.managedBy)"
                }
                
                # Database/SQL dependencies
                if ($resource.properties.server) {
                    $dependencies += "SERVER:$($resource.properties.server)"
                }
                
                # App Service dependencies
                if ($resource.properties.serverFarmId) {
                    $dependencies += "APP_PLAN:$($resource.properties.serverFarmId)"
                }
                
                # Container dependencies
                if ($resource.properties.managedEnvironmentId) {
                    $dependencies += "CONTAINER_ENV:$($resource.properties.managedEnvironmentId)"
                }
                if ($resource.properties.containerRegistryServer) {
                    $dependencies += "ACR:$($resource.properties.containerRegistryServer)"
                }
                
                # Load balancer dependencies
                if ($resource.properties.loadBalancerBackendAddressPools) {
                    foreach ($pool in $resource.properties.loadBalancerBackendAddressPools) {
                        if ($pool.id) { $dependencies += "LB_POOL:$($pool.id)" }
                    }
                }
                
                # Add dependencies as a column
                if ($dependencies.Count -gt 0) {
                    $resource | Add-Member -NotePropertyName 'Dependencies' -NotePropertyValue ($dependencies -join '; ') -Force
                } else {
                    $resource | Add-Member -NotePropertyName 'Dependencies' -NotePropertyValue '' -Force
                }
            }
            
            # Extract SKU details
            if ($resource.sku) {
                if ($resource.sku -is [PSObject]) {
                    $resource | Add-Member -NotePropertyName 'SkuName' -NotePropertyValue $resource.sku.name -Force -ErrorAction SilentlyContinue
                    $resource | Add-Member -NotePropertyName 'SkuTier' -NotePropertyValue $resource.sku.tier -Force -ErrorAction SilentlyContinue
                    $resource | Add-Member -NotePropertyName 'SkuSize' -NotePropertyValue $resource.sku.size -Force -ErrorAction SilentlyContinue
                    $resource | Add-Member -NotePropertyName 'SkuCapacity' -NotePropertyValue $resource.sku.capacity -Force -ErrorAction SilentlyContinue
                } else {
                    $resource | Add-Member -NotePropertyName 'SkuName' -NotePropertyValue $resource.sku -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Group by resource type for separate exports
            $resourceType = $resource.type
            if (-not $resourcesByType.ContainsKey($resourceType)) {
                $resourcesByType[$resourceType] = @()
            }
            $resourcesByType[$resourceType] += $resource
            
            # Add to master collection
            $allResources += $resource
        }
    }
    catch {
        Write-Log "Error processing subscription $($subscription.Name): $_" -Level Error
    }
}

Write-Log "=" * 60 -Level Info
Write-Log "Processing complete. Exporting data..." -Level Info

# Export zone mappings
$zoneMappingFile = Join-Path $OutputPath "ZoneMappings_AllSubscriptions.csv"
$allZoneMappings | Export-Csv -Path $zoneMappingFile -NoTypeInformation -Encoding UTF8
Write-Log "Exported zone mappings to ZoneMappings_AllSubscriptions.csv" -Level Success

# Export individual resource type files
Write-Log "Exporting individual resource type files..." -Level Info
foreach ($resourceType in $resourcesByType.Keys) {
    Export-ResourceTypeCSV -ResourceType $resourceType -Resources $resourcesByType[$resourceType] -OutputFolder $resourceTypesFolder
}

# Export master report with selected columns
Write-Log "Exporting master report..." -Level Info

$masterColumns = @(
    'SubscriptionName', 'subscriptionId', 'resourceGroup', 'name', 'type', 
    'ResourceProvider', 'ResourceTypeName', 'location', 'kind',
    'ZoneRedundancyStatus', 'ZoneRedundancyDetail', 'LogicalZones', 'PhysicalZones',
    'CrossRegionStatus', 'CrossRegionDetail', 'PairedRegion', 'GeoRedundant', 'DRReady',
    'SkuName', 'SkuTier', 'SkuSize', 'SkuCapacity',
    'Prop_provisioningState', 'Prop_status', 'Prop_zoneRedundant',
    'StorageRedundancy', 'StorageSKU', 'DiskRedundancy', 'DiskSKU',
    'ResourceId'
)

# Get all tag columns
$tagColumns = $allResources | ForEach-Object { $_.PSObject.Properties.Name } | Where-Object { $_ -like 'Tag_*' } | Sort-Object -Unique
$masterColumns += $tagColumns

$masterReportFile = Join-Path $OutputPath "MasterReport_AllResources.csv"
$allResources | Select-Object -Property $masterColumns -ErrorAction SilentlyContinue | Export-Csv -Path $masterReportFile -NoTypeInformation -Encoding UTF8
Write-Log "Exported master report to MasterReport_AllResources.csv" -Level Success

# Export summary report
Write-Log "Generating summary report..." -Level Info
$summaryData = $allResources | Group-Object -Property type, ZoneRedundancyStatus | ForEach-Object {
    [PSCustomObject]@{
        ResourceType = ($_.Name -split ', ')[0]
        ZoneStatus = ($_.Name -split ', ')[1]
        Count = $_.Count
    }
}

$summaryFile = Join-Path $OutputPath "Summary_ZoneRedundancy.csv"
$summaryData | Sort-Object ResourceType, ZoneStatus | Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8
Write-Log "Exported summary to Summary_ZoneRedundancy.csv" -Level Success

# Export subscription summary
$subSummary = $allResources | Group-Object SubscriptionName | ForEach-Object {
    $subResources = $_.Group
    [PSCustomObject]@{
        SubscriptionName = $_.Name
        TotalResources = $_.Count
        ZoneRedundant = ($subResources | Where-Object { $_.ZoneRedundancyStatus -eq 'ZoneRedundant' }).Count
        Zonal = ($subResources | Where-Object { $_.ZoneRedundancyStatus -eq 'Zonal' }).Count
        LocallyRedundant = ($subResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' }).Count
        NonZonal = ($subResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' }).Count
        GeoRedundant = ($subResources | Where-Object { $_.ZoneRedundancyStatus -like '*Geo*' }).Count
        RedundantByDefault = ($subResources | Where-Object { $_.ZoneRedundancyStatus -eq 'RedundantByDefault' }).Count
        Unknown = ($subResources | Where-Object { $_.ZoneRedundancyStatus -eq 'Unknown' }).Count
    }
}

$subSummaryFile = Join-Path $OutputPath "Summary_BySubscription.csv"
$subSummary | Export-Csv -Path $subSummaryFile -NoTypeInformation -Encoding UTF8
Write-Log "Exported subscription summary to Summary_BySubscription.csv" -Level Success

# Export Excel reports with formatting (if ImportExcel module is available)
if ($ExcelAvailable) {
    Write-Log "Generating formatted Excel reports..." -Level Info
    
    # Define conditional formatting colors
    $statusColors = @{
        'ZoneRedundant'     = '#28a745' # Green
        'Zonal'             = '#007bff' # Blue
        'LocallyRedundant'  = '#ffc107' # Yellow
        'NonZonal'          = '#dc3545' # Red
        'GeoRedundant'      = '#17a2b8' # Cyan
        'GeoZoneRedundant'  = '#20c997' # Teal
        'RedundantByDefault' = '#6c757d' # Gray
        'Unknown'           = '#fd7e14' # Orange
    }
    
    # Master Report Excel
    $masterExcelFile = Join-Path $OutputPath "MasterReport_AllResources.xlsx"
    
    # Export with formatting
    $excelParams = @{
        Path = $masterExcelFile
        AutoSize = $true
        AutoFilter = $true
        FreezeTopRow = $true
        BoldTopRow = $true
        TableStyle = 'Medium2'
        WorksheetName = 'All Resources'
    }
    
    $excel = $allResources | Select-Object -Property $masterColumns -ErrorAction SilentlyContinue | 
        Export-Excel @excelParams -PassThru
    
    $ws = $excel.Workbook.Worksheets['All Resources']
    
    # Find the ZoneRedundancyStatus column
    $statusCol = $null
    for ($col = 1; $col -le $ws.Dimension.Columns; $col++) {
        if ($ws.Cells[1, $col].Text -eq 'ZoneRedundancyStatus') {
            $statusCol = $col
            break
        }
    }
    
    # Apply cell colors based on status value
    if ($statusCol -and $ws.Dimension.Rows -gt 1) {
        for ($row = 2; $row -le $ws.Dimension.Rows; $row++) {
            $cellValue = $ws.Cells[$row, $statusCol].Text
            if ($statusColors.ContainsKey($cellValue)) {
                $color = [System.Drawing.ColorTranslator]::FromHtml($statusColors[$cellValue])
                $ws.Cells[$row, $statusCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[$row, $statusCol].Style.Fill.BackgroundColor.SetColor($color)
                
                # White text for dark backgrounds
                if ($cellValue -in @('ZoneRedundant', 'NonZonal', 'Zonal', 'GeoRedundant', 'GeoZoneRedundant')) {
                    $ws.Cells[$row, $statusCol].Style.Font.Color.SetColor([System.Drawing.Color]::White)
                }
            }
        }
    }
    
    # Set column widths for better readability
    $ws.Column(1).Width = 25  # SubscriptionName
    $ws.Column(2).Width = 38  # subscriptionId
    $ws.Column(3).Width = 30  # resourceGroup
    $ws.Column(4).Width = 40  # name
    $ws.Column(5).Width = 50  # type
    
    Close-ExcelPackage $excel -Show:$false
    
    Write-Log "Exported formatted master report to MasterReport_AllResources.xlsx" -Level Success
    
    # Summary Excel with multiple worksheets
    $summaryExcelFile = Join-Path $OutputPath "Summary_ZoneRedundancy.xlsx"
    
    # Summary by Resource Type
    $summaryData | Sort-Object ResourceType, ZoneStatus | 
        Export-Excel -Path $summaryExcelFile -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle 'Medium6' -WorksheetName 'By Resource Type'
    
    # Summary by Subscription
    $summaryExcel = $subSummary | 
        Export-Excel -Path $summaryExcelFile -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle 'Medium9' -WorksheetName 'By Subscription' -PassThru
    
    $ws = $summaryExcel.Workbook.Worksheets['By Subscription']
    
    # Color code the count columns (cell by cell for compatibility)
    if ($ws.Dimension.Rows -gt 1) {
        $greenColor = [System.Drawing.ColorTranslator]::FromHtml('#d4edda')
        $redColor = [System.Drawing.ColorTranslator]::FromHtml('#f8d7da')
        
        for ($row = 2; $row -le $ws.Dimension.Rows; $row++) {
            # ZoneRedundant column (green) - column 3
            $ws.Cells[$row, 3].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $ws.Cells[$row, 3].Style.Fill.BackgroundColor.SetColor($greenColor)
            
            # NonZonal column (red) - column 6
            $ws.Cells[$row, 6].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $ws.Cells[$row, 6].Style.Fill.BackgroundColor.SetColor($redColor)
        }
    }
    
    Close-ExcelPackage $summaryExcel -Show:$false
    
    # Pivot-style summary
    $pivotData = $allResources | Group-Object SubscriptionName, ZoneRedundancyStatus | ForEach-Object {
        [PSCustomObject]@{
            SubscriptionName = ($_.Name -split ', ')[0]
            ZoneStatus = ($_.Name -split ', ')[1]
            Count = $_.Count
        }
    }
    
    $pivotData | Sort-Object SubscriptionName, ZoneStatus | 
        Export-Excel -Path $summaryExcelFile -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle 'Medium3' -WorksheetName 'Pivot Data'
    
    Write-Log "Exported formatted summary to Summary_ZoneRedundancy.xlsx" -Level Success
}
else {
    Write-Log "ImportExcel module not available. Skipping Excel export." -Level Warning
}

# Generate HTML Dashboard Report
Write-Log "Generating HTML Dashboard Report..." -Level Info

# Get unique subscriptions for filter dropdown
$uniqueSubscriptions = $allResources | Select-Object -ExpandProperty SubscriptionName -Unique | Sort-Object

# Get unique locations for filter dropdown
$uniqueLocations = $allResources | Select-Object -ExpandProperty location -Unique | Sort-Object

# Build dynamic column list from all resource properties
$allPropertyNames = $allResources | ForEach-Object { $_.PSObject.Properties.Name } | Sort-Object -Unique

# Define column display configuration
# Priority columns shown by default, others available but hidden
$priorityColumns = @(
    @{ Name = 'SubscriptionName'; DisplayName = 'Subscription'; DefaultVisible = $true }
    @{ Name = 'name'; DisplayName = 'Resource Name'; DefaultVisible = $true }
    @{ Name = 'type'; DisplayName = 'Type'; DefaultVisible = $true }
    @{ Name = 'location'; DisplayName = 'Location'; DefaultVisible = $true }
    @{ Name = 'resourceGroup'; DisplayName = 'Resource Group'; DefaultVisible = $true }
    @{ Name = 'ZoneRedundancyStatus'; DisplayName = 'Zone Status'; DefaultVisible = $true }
    @{ Name = 'ZoneRedundancyDetail'; DisplayName = 'Zone Detail'; DefaultVisible = $true }
    @{ Name = 'LogicalZones'; DisplayName = 'Logical Zones'; DefaultVisible = $false }
    @{ Name = 'PhysicalZones'; DisplayName = 'Physical Zones'; DefaultVisible = $false }
    @{ Name = 'CrossRegionStatus'; DisplayName = 'Cross-Region Status'; DefaultVisible = $true }
    @{ Name = 'CrossRegionDetail'; DisplayName = 'Cross-Region Detail'; DefaultVisible = $true }
    @{ Name = 'PairedRegion'; DisplayName = 'Paired Region'; DefaultVisible = $true }
    @{ Name = 'GeoRedundant'; DisplayName = 'Geo Redundant'; DefaultVisible = $false }
    @{ Name = 'DRReady'; DisplayName = 'DR Ready'; DefaultVisible = $false }
    @{ Name = 'SkuName'; DisplayName = 'SKU Name'; DefaultVisible = $true }
    @{ Name = 'SkuTier'; DisplayName = 'SKU Tier'; DefaultVisible = $true }
    @{ Name = 'kind'; DisplayName = 'Kind'; DefaultVisible = $false }
    @{ Name = 'SkuSize'; DisplayName = 'SKU Size'; DefaultVisible = $false }
    @{ Name = 'SkuCapacity'; DisplayName = 'SKU Capacity'; DefaultVisible = $false }
    @{ Name = 'ResourceId'; DisplayName = 'Resource ID'; DefaultVisible = $false }
    @{ Name = 'subscriptionId'; DisplayName = 'Subscription ID'; DefaultVisible = $false }
    @{ Name = 'ResourceProvider'; DisplayName = 'Resource Provider'; DefaultVisible = $false }
    @{ Name = 'ResourceTypeName'; DisplayName = 'Resource Type Name'; DefaultVisible = $false }
    @{ Name = 'StorageRedundancy'; DisplayName = 'Storage Redundancy'; DefaultVisible = $false }
    @{ Name = 'StorageSKU'; DisplayName = 'Storage SKU'; DefaultVisible = $false }
    @{ Name = 'DiskRedundancy'; DisplayName = 'Disk Redundancy'; DefaultVisible = $false }
    @{ Name = 'DiskSKU'; DisplayName = 'Disk SKU'; DefaultVisible = $false }
)

# Get property columns (Prop_*)
$propColumns = $allPropertyNames | Where-Object { $_ -like 'Prop_*' } | ForEach-Object {
    @{ Name = $_; DisplayName = ($_ -replace '^Prop_', ''); DefaultVisible = $false }
}

# Get tag columns (Tag_*)
$tagColumnNames = $allPropertyNames | Where-Object { $_ -like 'Tag_*' }
$tagColumnsConfig = $tagColumnNames | ForEach-Object {
    @{ Name = $_; DisplayName = ($_ -replace '^Tag_', 'Tag: '); DefaultVisible = $false }
}

# Combine all columns in order: priority, properties, tags
$allColumns = @()
$allColumns += $priorityColumns
$allColumns += $propColumns
$allColumns += $tagColumnsConfig

# Filter to only columns that exist in the data
$allColumns = $allColumns | Where-Object { $allPropertyNames -contains $_.Name }

# Generate column selector HTML
$columnSelectorHtml = ""
$colIndex = 0
foreach ($col in $allColumns) {
    $checked = if ($col.DefaultVisible) { "checked" } else { "" }
    $displayStyle = if ($col.DefaultVisible) { "" } else { "display:none;" }
    $columnSelectorHtml += "            <label><input type=`"checkbox`" $checked onchange=`"toggleColumn($colIndex)`"> $($col.DisplayName)</label>`n"
    $colIndex++
}

# Generate table header HTML
$tableHeaderHtml = ""
$colIndex = 0
foreach ($col in $allColumns) {
    $displayStyle = if ($col.DefaultVisible) { "" } else { " style=`"display:none;`"" }
    $tableHeaderHtml += "                            <th data-col=`"$colIndex`"$displayStyle>$($col.DisplayName)</th>`n"
    $colIndex++
}

$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Azure Resilience & Disaster Recovery Assessment</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        /* ═══════════════════════════════════════════════════════════
           MICROSOFT AZURE — BCDR ASSESSMENT DASHBOARD
           Design: Single professional palette, portal-style
           ═══════════════════════════════════════════════════════════ */
        :root {
            --ms-blue:       #0078d4;
            --ms-blue-dark:  #003f87;
            --ms-blue-light: #cce4f6;
            --ms-navy:       #002050;
            --ms-gray-bg:    #f3f5f7;
            --ms-gray-border:#d0d5dd;
            --ms-green:      #107c10;
            --ms-green-bg:   #dff6dd;
            --ms-red:        #c50f1f;
            --ms-red-bg:     #fde7e9;
            --ms-amber:      #835b00;
            --ms-amber-bg:   #fff4ce;
            --ms-teal:       #006058;
            --ms-teal-bg:    #e0f3f1;
            --card-radius:   10px;
            --shadow-sm:     0 1px 4px rgba(0,0,0,0.10);
            --shadow-md:     0 3px 12px rgba(0,0,0,0.12);
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--ms-gray-bg); color: #1b1b1b; line-height: 1.5; }

        /* ── Microsoft Top-Bar ─────────────────────────────────────── */
        .ms-topbar {
            background: var(--ms-navy);
            color: white;
            padding: 0 28px;
            height: 44px;
            display: flex;
            align-items: center;
            gap: 16px;
            font-size: 0.82em;
            letter-spacing: 0.2px;
            border-bottom: 3px solid var(--ms-blue);
        }
        .ms-topbar .ms-logo { display: flex; align-items: center; gap: 8px; text-decoration: none; color: white; font-weight: 600; font-size: 1.1em; }
        .ms-topbar .ms-logo svg { flex-shrink: 0; }
        .ms-topbar .ms-logo-sep { width: 1px; height: 22px; background: rgba(255,255,255,0.3); }
        .ms-topbar .ms-product { opacity: 0.85; font-size: 0.95em; }
        .ms-topbar .ms-spacer { flex: 1; }
        .ms-topbar .ms-topbar-pill {
            background: rgba(255,255,255,0.12);
            border: 1px solid rgba(255,255,255,0.2);
            border-radius: 14px;
            padding: 3px 12px;
            font-size: 0.82em;
            color: rgba(255,255,255,0.9);
        }

        /* ── Page Header ───────────────────────────────────────────── */
        .header {
            background: linear-gradient(135deg, var(--ms-blue-dark) 0%, var(--ms-blue) 60%, #0091ea 100%);
            color: white;
            padding: 28px 32px 22px;
            border-bottom: 1px solid rgba(0,0,0,0.15);
        }
        .header-inner { max-width: 100%; display: flex; align-items: flex-start; justify-content: space-between; flex-wrap: wrap; gap: 16px; }
        .header-title h1 { font-size: 1.65em; font-weight: 600; margin-bottom: 4px; letter-spacing: -0.3px; }
        .header-title p { font-size: 0.88em; opacity: 0.82; }
        .header-stats { display: flex; gap: 24px; }
        .header-stat { text-align: center; background: rgba(255,255,255,0.12); border-radius: 8px; padding: 8px 18px; border: 1px solid rgba(255,255,255,0.18); min-width: 90px; }
        .header-stat .stat-val { font-size: 1.6em; font-weight: 700; line-height: 1.1; }
        .header-stat .stat-lbl { font-size: 0.72em; opacity: 0.78; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }

        /* ── Power BI Banner ─────────────────────────────────────────── */
        .powerbi-banner { background: #fffbeb; border-bottom: 1px solid #e8d77a; padding: 10px 28px; font-size: 0.84em; color: #5a4a00; display: flex; align-items: center; gap: 8px; }
        .powerbi-banner a { color: var(--ms-blue); font-weight: 600; }

        /* ── Container ───────────────────────────────────────────────── */
        .container { max-width: 100%; margin: 0 auto; padding: 20px 24px; }

        /* ── Export Info strip ──────────────────────────────────────── */
        .export-info { background: white; border: 1px solid var(--ms-gray-border); border-left: 4px solid var(--ms-green); padding: 10px 16px; margin: 0 0 18px; border-radius: 6px; font-size: 0.83em; color: #333; }

        /* ── Summary Cards ───────────────────────────────────────────── */
        .summary-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(195px, 1fr)); gap: 14px; margin: 0 0 16px; }
        .card {
            background: white;
            border-radius: var(--card-radius);
            padding: 18px 16px 12px;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--ms-gray-border);
            border-top: 4px solid var(--ms-gray-border);
            transition: box-shadow 0.18s, transform 0.18s;
            text-align: center;
            position: relative;
        }
        .card:hover { box-shadow: var(--shadow-md); transform: translateY(-2px); }
        .card.total     { border-top-color: var(--ms-blue); }
        .card.zone-redundant { border-top-color: var(--ms-green); }
        .card.zonal     { border-top-color: #0091ea; }
        .card.non-zonal { border-top-color: var(--ms-red); }
        .card.locally-redundant { border-top-color: #ca5010; }
        .card.rbd-card  { border-top-color: #00bcf2; }

        /* Card text elements */
        .card-label-business { font-size: 0.95em; font-weight: 600; color: #1b1b1b; margin-bottom: 2px; line-height: 1.25; }
        .card-label-tech { display: inline-block; font-size: 0.72em; font-weight: 600; letter-spacing: 0.4px; padding: 1px 8px; border-radius: 10px; margin-bottom: 10px; font-family: 'Segoe UI', monospace; }
        .card.total .card-label-tech       { background: var(--ms-blue-light); color: var(--ms-blue-dark); }
        .card.zone-redundant .card-label-tech { background: var(--ms-green-bg); color: #0a5e0a; }
        .card.zonal .card-label-tech       { background: #d6f0ff; color: #004d8a; }
        .card.non-zonal .card-label-tech   { background: var(--ms-red-bg); color: #8a0010; }
        .card.locally-redundant .card-label-tech { background: var(--ms-amber-bg); color: var(--ms-amber); }
        .card.rbd-card .card-label-tech    { background: #d8f4fd; color: #005578; }

        .card .value { font-size: 2.4em; font-weight: 700; line-height: 1.0; margin: 4px 0 6px; }
        .card.total .value          { color: var(--ms-blue); }
        .card.zone-redundant .value { color: var(--ms-green); }
        .card.zonal .value          { color: #0091ea; }
        .card.non-zonal .value      { color: var(--ms-red); }
        .card.locally-redundant .value { color: #ca5010; }
        .card.rbd-card .value       { color: #005578; }

        .card-subtitle { font-size: 0.71em; color: #6e6e6e; line-height: 1.35; margin-bottom: 2px; }
        .card-scope-pill { display: inline-block; font-size: 0.68em; padding: 1px 7px; border-radius: 8px; margin-top: 4px; font-weight: 600; }
        .scope-iaas  { background: #ffe8d9; color: #7a2e00; }
        .scope-paas  { background: #d6eeff; color: #003f87; }
        .scope-both  { background: #f0e8f8; color: #3a006b; }
        .scope-all   { background: #e8e8e8; color: #3a3a3a; }
        .scope-plat  { background: #d8f4fd; color: #005578; }

        .card-expand-btn { display: block; width: 100%; margin-top: 10px; padding: 5px 0; font-size: 0.74em; color: var(--ms-blue); background: none; border: none; border-top: 1px solid #eaeaea; cursor: pointer; text-align: center; letter-spacing: 0.2px; }
        .card-expand-btn:hover { color: var(--ms-blue-dark); background: #f0f7ff; border-radius: 0 0 8px 8px; }
        .card-detail-panel { display: none; text-align: left; margin-top: 8px; padding: 10px 12px; background: #fafcff; border-radius: 7px; font-size: 0.78em; line-height: 1.55; color: #444; border: 1px solid #d6e9f8; }
        .card-detail-panel.open { display: block; }
        .card-detail-panel p { margin: 0 0 6px 0; }
        .card-detail-panel p:last-child { margin-bottom: 0; }
        .card-detail-panel strong { color: #1b1b1b; }
        .card-detail-panel .example-types { margin-top: 6px; }
        .card-detail-panel .example-types span { display: inline-block; background: #e8f0f8; border-radius: 4px; padding: 1px 6px; margin: 2px 3px 2px 0; font-size: 0.92em; font-family: monospace; color: #1b4f8a; }
        .card-header { display: flex; align-items: flex-start; justify-content: center; gap: 5px; }
        .tooltip-icon { cursor: help; font-size: 0.82em; opacity: 0.5; flex-shrink: 0; margin-top: 2px; }
        .tooltip-icon:hover { opacity: 1; }

        /* ── Charts ──────────────────────────────────────────────────── */
        .charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 16px; margin: 16px 0; }
        .chart-container { background: white; border-radius: var(--card-radius); padding: 20px; box-shadow: var(--shadow-sm); border: 1px solid var(--ms-gray-border); }
        .chart-container h2 { font-size: 0.95em; font-weight: 600; margin-bottom: 14px; color: #1b1b1b; border-bottom: 2px solid var(--ms-blue); padding-bottom: 8px; letter-spacing: 0.1px; }
        .chart-container .chart-subtitle { font-size: 0.76em; color: #666; margin: -10px 0 12px; }

        /* ── Tables ──────────────────────────────────────────────────── */
        .table-container { background: white; border-radius: var(--card-radius); padding: 20px; box-shadow: var(--shadow-sm); border: 1px solid var(--ms-gray-border); margin: 16px 0; }
        .table-container h2 { font-size: 1.05em; font-weight: 600; margin-bottom: 12px; color: #1b1b1b; border-bottom: 2px solid var(--ms-blue); padding-bottom: 8px; }
        .table-wrapper { overflow-x: auto; max-height: 600px; overflow-y: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 0.84em; min-width: 1200px; table-layout: auto; }
        th { background: var(--ms-blue); color: white; padding: 10px 8px; text-align: left; font-weight: 600; position: sticky; top: 0; z-index: 10; white-space: nowrap; cursor: col-resize; border-right: 1px solid rgba(255,255,255,0.2); min-width: 80px; resize: horizontal; overflow: auto; }
        th:last-child { border-right: none; }
        td { padding: 7px 8px; border-bottom: 1px solid #eef0f3; max-width: 300px; overflow: hidden; text-overflow: ellipsis; word-wrap: break-word; }
        td.expanded { white-space: normal; max-width: none; overflow: visible; }
        td:hover { background: #f0f7ff; cursor: pointer; }
        tr:hover { background: #f0f7ff; }
        tr:nth-child(even) { background: #fafbfc; }
        tr:nth-child(even):hover { background: #f0f7ff; }

        /* ── Status Badges ───────────────────────────────────────────── */
        .status-badge { padding: 2px 9px; border-radius: 12px; font-size: 0.8em; font-weight: 600; display: inline-block; white-space: nowrap; }
        .status-ZoneRedundant { background: var(--ms-green-bg); color: #0a5e0a; }
        .status-Zonal { background: #d6f0ff; color: #004d8a; }
        .status-NonZonal { background: var(--ms-red-bg); color: #8a0010; }
        .status-LocallyRedundant { background: var(--ms-amber-bg); color: var(--ms-amber); }
        .status-RedundantByDefault { background: #d8f4fd; color: #005578; }
        .status-GeoRedundant, .status-GeoZoneRedundant { background: var(--ms-teal-bg); color: var(--ms-teal); }
        .status-Unknown, .status-InheritedFromParent { background: #f3f2f1; color: #605e5c; }
        .risk-high { background: #fff0f1 !important; }
        .risk-medium { background: #fffdf0 !important; }

        /* ── Filter Bar ──────────────────────────────────────────────── */
        .filter-bar { background: white; border-radius: var(--card-radius); padding: 13px 18px; margin: 16px 0; box-shadow: var(--shadow-sm); border: 1px solid var(--ms-gray-border); display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
        .filter-bar label { font-weight: 600; font-size: 0.84em; color: #444; }
        .filter-bar input, .filter-bar select { padding: 7px 11px; border: 1px solid #d0d5dd; border-radius: 5px; font-size: 0.86em; }
        .filter-bar input[type="text"] { width: 240px; }
        .filter-group { display: flex; align-items: center; gap: 5px; }
        .btn { padding: 7px 15px; border: none; border-radius: 5px; cursor: pointer; font-size: 0.86em; font-weight: 600; transition: background 0.15s; }
        .btn-primary { background: var(--ms-blue); color: white; }
        .btn-primary:hover { background: var(--ms-blue-dark); }
        .btn-secondary { background: white; color: #333; border: 1px solid #ccc; }
        .btn-secondary:hover { background: #f0f0f0; border-color: var(--ms-blue); }

        /* ── Column Selector ─────────────────────────────────────────── */
        .column-selector { background: white; border: 1px solid var(--ms-gray-border); border-radius: 8px; padding: 15px; margin: 10px 0; display: none; max-height: 400px; overflow-y: auto; }
        .column-selector.show { display: block; }
        .column-selector h4 { margin-bottom: 10px; color: #333; position: sticky; top: 0; background: white; padding: 5px 0; font-size: 0.9em; }
        .column-selector label { display: inline-flex; align-items: center; margin: 4px 14px 4px 0; font-size: 0.87em; cursor: pointer; min-width: 180px; }
        .column-selector input[type="checkbox"] { margin-right: 5px; }
        .column-selector-actions { margin-bottom: 10px; padding-bottom: 10px; border-bottom: 1px solid #eee; }
        .column-selector-actions button { margin-right: 10px; padding: 4px 10px; font-size: 0.83em; }

        /* ── Recommendations ─────────────────────────────────────────── */
        .recommendations { background: #fffef5; border-left: 4px solid #ca5010; padding: 14px 18px; margin: 14px 0; border-radius: 0 8px 8px 0; border: 1px solid #f0e4cc; border-left-width: 4px; }
        .recommendations h4 { color: #333; margin-bottom: 8px; font-size: 0.95em; }
        .recommendations ul { margin-left: 18px; }
        .recommendations li { margin: 5px 0; color: #555; font-size: 0.9em; }

        /* ── BCDR Note Box ────────────────────────────────────────────── */
        .bcdr-note-box { background: white; border: 1px solid #b3d8f5; border-left: 4px solid var(--ms-blue); border-radius: 8px; padding: 16px 22px; margin: 16px 0; }
        .bcdr-note-box h4 { color: var(--ms-blue-dark); margin-bottom: 10px; font-size: 0.97em; }
        .bcdr-note-box p { color: #444; font-size: 0.88em; line-height: 1.6; margin-bottom: 6px; }
        .bcdr-note-box p:last-child { margin-bottom: 0; }

        /* ── Category Legend ─────────────────────────────────────────── */
        .category-legend { background: white; border: 1px solid var(--ms-gray-border); border-radius: var(--card-radius); padding: 18px 22px; margin: 16px 0; box-shadow: var(--shadow-sm); }
        .category-legend h4 { margin: 0 0 14px; color: #1b1b1b; font-size: 0.97em; border-bottom: 2px solid var(--ms-blue); padding-bottom: 8px; }
        .legend-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 12px; }
        .legend-item { padding: 11px 13px; border-radius: 7px; border-left: 4px solid; font-size: 0.85em; }
        .legend-item .legend-title { font-weight: 700; font-size: 0.94em; margin-bottom: 3px; }
        .legend-item .legend-scope { font-size: 0.83em; font-style: italic; color: #666; margin-bottom: 5px; }
        .legend-item .legend-examples span { display: inline-block; background: rgba(0,0,0,0.07); border-radius: 3px; padding: 1px 5px; margin: 1px 2px 1px 0; font-family: monospace; font-size: 0.93em; }
        .legend-zr   { background: #f0fbf0; border-color: var(--ms-green); }
        .legend-zonal { background: #f0f7ff; border-color: #0091ea; }
        .legend-nz   { background: #fff3f3; border-color: var(--ms-red); }
        .legend-lr   { background: #fffef3; border-color: #ca5010; }
        .legend-rbd  { background: #f0fbff; border-color: #00bcf2; }
        .legend-geo  { background: #f0f5f5; border-color: var(--ms-teal); }
        .legend-rbd-resources { background: #f0fbff; border: 1px solid #b3e8fb; border-radius: 7px; padding: 10px 14px; margin-top: 12px; font-size: 0.83em; }
        .legend-rbd-resources .rbd-title { font-weight: 700; color: #005578; margin-bottom: 6px; }
        .legend-rbd-resources .rbd-list { display: flex; flex-wrap: wrap; gap: 4px; }
        .legend-rbd-resources .rbd-list span { background: #cdf0fb; border-radius: 4px; padding: 2px 8px; font-family: monospace; font-size: 0.9em; color: #004c54; }

        /* ── Multi-Select Dropdowns ──────────────────────────────────── */
        .mdd-wrapper { position: relative; min-width: 160px; }
        .mdd-btn { width: 100%; padding: 7px 10px; border: 1px solid #d0d5dd; border-radius: 5px; background: white; cursor: pointer; text-align: left; font-size: 0.86em; display: flex; justify-content: space-between; align-items: center; gap: 6px; white-space: nowrap; }
        .mdd-btn:hover { border-color: var(--ms-blue); background: #f0f7ff; }
        .mdd-btn .mdd-label { overflow: hidden; text-overflow: ellipsis; max-width: 160px; }
        .mdd-btn .mdd-arrow { font-size: 0.72em; flex-shrink: 0; }
        .mdd-panel { position: absolute; top: calc(100% + 3px); left: 0; min-width: 260px; max-width: 420px; background: white; border: 1px solid #bbb; border-radius: 7px; box-shadow: 0 6px 20px rgba(0,0,0,0.14); z-index: 2000; display: none; }
        .mdd-panel.open { display: block; }
        .mdd-actions { padding: 7px 10px; border-bottom: 1px solid #eee; display: flex; gap: 8px; background: #f7f9fb; border-radius: 7px 7px 0 0; }
        .mdd-actions button { padding: 3px 10px; font-size: 0.81em; border: 1px solid #ccc; border-radius: 4px; cursor: pointer; background: white; }
        .mdd-actions button:hover { background: var(--ms-blue); color: white; border-color: var(--ms-blue); }
        .mdd-search { width: calc(100% - 16px); margin: 6px 8px 2px; padding: 5px 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 0.84em; display: block; }
        .mdd-list { padding: 4px 0; max-height: 240px; overflow-y: auto; }
        .mdd-item { display: flex; align-items: center; padding: 5px 12px; cursor: pointer; font-size: 0.86em; }
        .mdd-item:hover { background: #f0f7ff; }
        .mdd-item input[type="checkbox"] { margin-right: 8px; cursor: pointer; flex-shrink: 0; }
        .mdd-item label { cursor: pointer; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }

        /* ── Footer ──────────────────────────────────────────────────── */
        .footer { background: var(--ms-navy); color: rgba(255,255,255,0.65); text-align: center; padding: 16px 24px; font-size: 0.8em; margin-top: 32px; border-top: 3px solid var(--ms-blue); }
        .footer a { color: rgba(255,255,255,0.7); }
        .resource-count { background: var(--ms-blue); color: white; padding: 2px 9px; border-radius: 10px; font-size: 0.83em; margin-left: 8px; font-weight: 600; }

        /* ── Section headers ─────────────────────────────────────────── */
        .section-header { font-size: 1.0em; font-weight: 700; color: #1b1b1b; margin: 22px 0 10px; padding-left: 10px; border-left: 4px solid var(--ms-blue); }

        @media (max-width: 768px) {
            .charts-row { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .filter-bar input[type="text"] { width: 100%; }
            .header-stats { display: none; }
            .ms-topbar .ms-topbar-pill { display: none; }
        }
    </style>
</head>
<body>
    <!-- Microsoft Top Bar -->
    <div class="ms-topbar">
        <a class="ms-logo" href="https://azure.microsoft.com" target="_blank">
            <!-- Microsoft logo (4-square) SVG -->
            <svg width="20" height="20" viewBox="0 0 21 21" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
                <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
                <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
                <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>
            Microsoft Azure
        </a>
        <div class="ms-logo-sep"></div>
        <span class="ms-product">BCDR &amp; Zone Resilience Assessment</span>
        <div class="ms-spacer"></div>
        <span class="ms-topbar-pill">🏢 Solution Engineering · Infrastructure</span>
        <span class="ms-topbar-pill">Generated: $(Get-Date -Format "dd MMM yyyy HH:mm")</span>
    </div>

    <!-- Power BI Upgrade Notice -->
    <div class="powerbi-banner">
        📊 <strong>Interactive Dashboards:</strong> For executive-ready, filterable Power BI reports, import the CSV files from the output folder into 
        <a href="https://powerbi.microsoft.com" target="_blank">Microsoft Power BI ↗</a>
        &nbsp;|&nbsp; This HTML view is for quick field review — all data is also available in the generated Excel workbook.
    </div>

    <!-- Page Header -->
    <div class="header">
        <div class="header-inner">
            <div class="header-title">
                <h1>Azure Resilience &amp; Disaster Recovery Assessment</h1>
                <p>Subscriptions: $totalSubs &nbsp;·&nbsp; Total Resources: $($allResources.Count) &nbsp;·&nbsp; $(Get-Date -Format "MMMM dd, yyyy")</p>
            </div>
            <div class="header-stats">
                <div class="header-stat">
                    <div class="stat-val" style="color:#6fcf6f;">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'ZoneRedundant' }).Count)</div>
                    <div class="stat-lbl">Zone Redundant</div>
                </div>
                <div class="header-stat">
                    <div class="stat-val" style="color:#ff9494;">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' }).Count)</div>
                    <div class="stat-lbl">Non-Zonal</div>
                </div>
                <div class="header-stat">
                    <div class="stat-val" style="color:#ffd166;">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' }).Count)</div>
                    <div class="stat-lbl">Locally Redundant</div>
                </div>
                <div class="header-stat">
                    <div class="stat-val">$totalSubs</div>
                    <div class="stat-lbl">Subscriptions</div>
                </div>
            </div>
        </div>
    </div>
    
    <div class="container">
        <!-- Export Info -->
        <div class="export-info">
            <strong>Output Location:</strong> $OutputPath &nbsp;·&nbsp;
            <strong>Files:</strong> MasterReport_AllResources.csv, Summary_ZoneRedundancy.csv, Summary_BySubscription.csv, ZoneMappings_AllSubscriptions.csv, MasterReport_AllResources.xlsx, Summary_ZoneRedundancy.xlsx, and per-type CSVs in /ResourceTypes/.
        </div>

        <!-- ═══ SUMMARY CARDS ═══════════════════════════════════════════════ -->
        <div class="section-header">📊 Zone Resilience At a Glance</div>
        <div class="summary-cards">
            <!-- CARD 1: Total -->
            <div class="card total">
                <div class="card-label-business">Total Azure Resources</div>
                <div class="card-label-tech">All Types · All Regions</div>
                <div class="value" id="card-total">$($allResources.Count)</div>
                <div class="card-subtitle">Across all $totalSubs subscriptions in scope</div>
                <span class="card-scope-pill scope-all">IaaS + PaaS + Platform</span>
                <button class="card-expand-btn" onclick="toggleCardDetail('cd-total',this)" data-open-label="▾ What's included?">▾ What's included?</button>
                <div class="card-detail-panel" id="cd-total">
                    <p>Every Azure resource discovered across all subscriptions in scope — VMs, databases, storage, networking, app services, and platform-managed resources.</p>
                    <p><strong>Note:</strong> Includes infrastructure resources (NICs, NSGs, Route Tables) which are not directly zone-configurable but are counted for completeness.</p>
                </div>
            </div>

            <!-- CARD 2: Zone Redundant -->
            <div class="card zone-redundant">
                <div class="card-header">
                    <div>
                        <div class="card-label-business">Resilient Across Datacenters</div>
                        <div class="card-label-tech">Zone Redundant (ZR)</div>
                    </div>
                    <span class="tooltip-icon" title="Technical term: Zone Redundant — service runs simultaneously across 2-3 Availability Zones (physical datacenters) within the same region">ℹ️</span>
                </div>
                <div class="value" id="card-zr">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'ZoneRedundant' }).Count)</div>
                <div class="card-subtitle">PaaS &amp; data services spanning multiple physical datacenters</div>
                <span class="card-scope-pill scope-paas">PaaS / Data Services only</span>
                <button class="card-expand-btn" onclick="toggleCardDetail('cd-zr',this)" data-open-label="▾ Learn more">▾ Learn more</button>
                <div class="card-detail-panel" id="cd-zr">
                    <p><strong>What this means (Business):</strong> These services are simultaneously active across 2–3 separate Microsoft data buildings. If one building goes dark, service continues automatically — <strong>zero downtime</strong>.</p>
                    <p><strong>What this means (IT):</strong> Zone Redundancy (ZR) is an explicit configuration on PaaS services — a ZR SKU, ZR flag, or ZR HA mode was enabled at deployment time. This does <em>not</em> apply to VMs; see the "Zonal" and "Non-Zonal" cards for VM-specific placement.</p>
                    <p><strong>✅ No action needed for zone resilience.</strong> Consider cross-region DR if full-region failures are in scope.</p>
                    <p class="example-types"><strong>Example services:</strong>
                        <span>Azure SQL (ZR=true)</span><span>AKS (ZR node pools)</span><span>Redis Premium+ZR</span><span>Event Hub Premium</span><span>Service Bus Premium</span><span>App Service P2v3 ZR</span><span>API Management Premium</span><span>PostgreSQL Flexible ZR-HA</span>
                    </p>
                </div>
            </div>

            <!-- CARD 3: Zonal (IaaS pinned) -->
            <div class="card zonal">
                <div class="card-header">
                    <div>
                        <div class="card-label-business">Assigned to One Datacenter</div>
                        <div class="card-label-tech">Zonal (Zone-Pinned IaaS)</div>
                    </div>
                    <span class="tooltip-icon" title="Technical term: Zonal — IaaS resource explicitly placed in a specific Availability Zone (Zone 1, 2 or 3). Your team chose the zone at deployment.">ℹ️</span>
                </div>
                <div class="value" id="card-zonal">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'Zonal' }).Count)</div>
                <div class="card-subtitle">IaaS resources explicitly placed in Zone 1, 2, or 3</div>
                <span class="card-scope-pill scope-iaas">IaaS Only — VMs, Disks, PIPs</span>
                <button class="card-expand-btn" onclick="toggleCardDetail('cd-zonal',this)" data-open-label="▾ Learn more">▾ Learn more</button>
                <div class="card-detail-panel" id="cd-zonal">
                    <p><strong>What this means (Business):</strong> These resources live in a known, specific building (Zone 1, 2, or 3). You have control and visibility. If that building has an outage, <em>this specific resource</em> is affected.</p>
                    <p><strong>What this means (IT — IaaS focus):</strong> This category is specifically about IaaS resources (Virtual Machines, Managed Disks, Public IP Addresses) where your deployment team set <code>zones: [1]</code> or <code>[2]</code> or <code>[3]</code>. This is NOT the same as Zone Redundant — a single-zone VM is still vulnerable to that zone's failure.</p>
                    <p><strong>⚠️ Action:</strong> For resilience, deploy matching VMs in other zones and front them with a Zone-Redundant Load Balancer or Traffic Manager.</p>
                    <p class="example-types"><strong>Typical resources:</strong>
                        <span>VM (zones=[1])</span><span>Managed Disk (zone-attached)</span><span>Public IP (Zone SKU)</span><span>Load Balancer (Zone frontend IP)</span>
                    </p>
                </div>
            </div>

            <!-- CARD 4: Non-Zonal — IaaS with no zone set -->
            <div class="card non-zonal">
                <div class="card-header">
                    <div>
                        <div class="card-label-business">No Datacenter Assignment</div>
                        <div class="card-label-tech">Non-Zonal — IaaS &amp; Zone-Capable PaaS</div>
                    </div>
                    <span class="tooltip-icon" title="Technical term: Non-Zonal — Resource was deployed without specifying an Availability Zone. Azure placed it in any available datacenter. Highest risk category for IaaS workloads.">ℹ️</span>
                </div>
                <div class="value" id="card-nonzonal">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' }).Count)</div>
                <div class="card-subtitle">⚠️ Resources without zone placement — Azure chose the datacenter</div>
                <span class="card-scope-pill scope-both">IaaS VMs · Zone-capable PaaS</span>
                <button class="card-expand-btn" onclick="toggleCardDetail('cd-nz',this)" data-open-label="▾ Why does this matter?">▾ Why does this matter?</button>
                <div class="card-detail-panel" id="cd-nz">
                    <p><strong>⚠️ IMPORTANT — This category is primarily about IaaS resources (Virtual Machines) and zone-capable PaaS services where zone redundancy was NOT enabled at deployment.</strong></p>
                    <p><strong>What this means (Business):</strong> These resources were placed by Azure wherever capacity was available. You don't know which building they're in. If a data-center outage occurs, <em>you won't know which resources are affected</em> until it happens.</p>
                    <p><strong>What this means (IT — VMs specifically):</strong> Non-Zonal VMs have no <code>zones</code> property set. Azure assigned them to an undisclosed physical zone. Multiple non-zonal VMs could be in the same zone — a single zone failure could take down several at once.</p>
                    <p><strong>What this means (IT — PaaS services):</strong> Services like AKS, PostgreSQL Flexible, App Gateway v2, and Azure Container Apps support Zone Redundancy but must have it configured at creation time. If created without it, they become non-zonal and cannot be converted in-place — a new resource must be created.</p>
                    <p>📌 <strong>Note:</strong> Not all non-zonal resources are critical. Prioritize based on workload tier — see the BCDR Excel report for per-resource guidance.</p>
                    <p class="example-types"><strong>Typical resources in this category:</strong>
                        <span>VM (no zones set)</span><span>AKS (no zone config)</span><span>PostgreSQL Flexible (no ZR HA)</span><span>App Gateway v2 (no zones)</span><span>Container Apps Env (not ZR)</span>
                    </p>
                </div>
            </div>

            <!-- CARD 5: Locally Redundant -->
            <div class="card locally-redundant">
                <div class="card-header">
                    <div>
                        <div class="card-label-business">Within One Datacenter Only</div>
                        <div class="card-label-tech">Locally Redundant (LRS)</div>
                    </div>
                    <span class="tooltip-icon" title="Technical term: Locally Redundant Storage (LRS) — 3 copies within a single datacenter. No zone protection.">ℹ️</span>
                </div>
                <div class="value" id="card-lr">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' }).Count)</div>
                <div class="card-subtitle">Storage &amp; Disks using LRS — single-datacenter replication only</div>
                <span class="card-scope-pill scope-iaas">Storage Accounts · Managed Disks</span>
                <button class="card-expand-btn" onclick="toggleCardDetail('cd-lr',this)" data-open-label="▾ Learn more">▾ Learn more</button>
                <div class="card-detail-panel" id="cd-lr">
                    <p><strong>Applies to: Storage Accounts (LRS) and Managed Disks (Premium_LRS / Standard_LRS) — NOT VMs or other compute resources.</strong></p>
                    <p><strong>What this means (Business):</strong> Your data is copied 3 times, but all copies are inside the same data building. A building-level failure means this data is temporarily unavailable.</p>
                    <p><strong>What this means (IT):</strong> LRS protects against disk/rack failures within a datacenter, but does NOT protect against an Availability Zone outage. Upgrade to ZRS (Zone-Redundant Storage) for zone protection — this is a live migration for storage accounts and requires no downtime (1-72 hr process).</p>
                    <p>✅ <strong>Quick Win:</strong> Storage Account LRS → ZRS via the Azure Portal "Redundancy" settings. No data loss, no downtime.</p>
                    <p class="example-types"><strong>Typical resources:</strong>
                        <span>Storage Account (Standard_LRS)</span><span>Storage Account (Premium_LRS)</span><span>Managed Disk (Standard_LRS)</span><span>Managed Disk (Premium_LRS)</span>
                    </p>
                </div>
            </div>

            <!-- CARD 6: Redundant By Default -->
            <div class="card rbd-card">
                <div class="card-header">
                    <div>
                        <div class="card-label-business">Microsoft Manages Resilience</div>
                        <div class="card-label-tech">Redundant by Default (Platform)</div>
                    </div>
                    <span class="tooltip-icon" title="Technical term: Redundant By Default — Azure platform infrastructure where Microsoft guarantees resilience as part of the service SLA. No zone configuration needed.">ℹ️</span>
                </div>
                <div class="value" id="card-rbd">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'RedundantByDefault' }).Count)</div>
                <div class="card-subtitle">Azure platform fabric — SLA-backed, no customer zone config needed</div>
                <span class="card-scope-pill scope-plat">Platform Infrastructure</span>
                <button class="card-expand-btn" onclick="toggleCardDetail('cd-rbd',this)" data-open-label="▾ What's included?">▾ What's included?</button>
                <div class="card-detail-panel" id="cd-rbd">
                    <p><strong>Applies to: Azure Platform Infrastructure and Networking Primitives — NOT business workloads you deploy.</strong></p>
                    <p><strong>What this means (Business):</strong> Microsoft automatically manages the resilience of these services as part of the Azure platform SLA. Your team does not need to configure any zone settings. These are the "plumbing" of Azure.</p>
                    <p><strong>What this means (IT):</strong> Resources like VNets, NSGs, NICs, Route Tables, Private DNS Zones, Key Vaults, Log Analytics Workspaces, and Recovery Services Vaults are managed by the Azure control plane. Microsoft distributes them across the region's infrastructure. Your team's responsibility is only to include them in cross-region DR plans in case of a full-region failure.</p>
                    <p>✅ <strong>No zone-resilience action required.</strong> Focus BCDR effort on NonZonal and LocallyRedundant categories.</p>
                    <p class="example-types"><strong>Includes:</strong>
                        <span>VNet</span><span>NSG</span><span>NIC</span><span>Route Table</span><span>Private DNS Zone</span><span>Private Endpoint</span><span>Key Vault</span><span>Log Analytics</span><span>Recovery Services Vault</span><span>Automation Account</span><span>App Insights</span><span>Logic Apps</span>
                    </p>
                </div>
            </div>
        </div>

        <!-- Resource Category Legend (Self-Explanatory Reference Panel) -->
        <div class="category-legend">
            <h4>📚 Quick Reference: Which Category Applies to Which Resource Type?</h4>
            <div class="legend-grid">
                <div class="legend-item legend-zr">
                    <div class="legend-title" style="color:#107c10">✅ Zone Redundant — PaaS / Data Services</div>
                    <div class="legend-scope">PaaS services where you have explicitly enabled zone-redundancy (ZR flag / ZR SKU / ZR HA mode)</div>
                    <div class="legend-examples"><span>Azure SQL (ZR)</span><span>AKS (ZR nodepool)</span><span>Redis (Premium+ZR)</span><span>Event Hub Premium</span><span>Service Bus Premium</span><span>App Service P2v3+ZR</span><span>API Management Premium</span><span>PostgreSQL Flexible (ZR HA)</span><span>Azure Container Apps Env (ZR)</span></div>
                </div>
                <div class="legend-item legend-zonal">
                    <div class="legend-title" style="color:#0078d4">📍 Zonal — IaaS Pinned Resources</div>
                    <div class="legend-scope">IaaS resources pinned to a specific zone by your team during deployment</div>
                    <div class="legend-examples"><span>VM (zone=1/2/3)</span><span>Managed Disk (zone-attached)</span><span>Public IP (Zone SKU)</span><span>Load Balancer (Zone frontend)</span><span>VMSS (zone-specific)</span></div>
                </div>
                <div class="legend-item legend-nz">
                    <div class="legend-title" style="color:#d83b01">⚠️ Non-Zonal — Needs Attention</div>
                    <div class="legend-scope">Resources that support zones but have none configured — Azure decides placement</div>
                    <div class="legend-examples"><span>VM (no zone)</span><span>AKS (no zone config)</span><span>PostgreSQL (no HA)</span><span>App Gateway v2 (no zone)</span><span>Container Apps (no ZR)</span><span>SQL MI (no failover)</span></div>
                </div>
                <div class="legend-item legend-lr">
                    <div class="legend-title" style="color:#ca5010">🗄️ Locally Redundant — Storage &amp; Disks Only</div>
                    <div class="legend-scope">Storage accounts and managed disks replicating only within a single datacenter</div>
                    <div class="legend-examples"><span>Storage (Standard_LRS)</span><span>Storage (Premium_LRS)</span><span>Managed Disk (Standard_LRS)</span><span>Managed Disk (Premium_LRS)</span><span>Managed Disk (UltraSSD_LRS)</span></div>
                </div>
                <div class="legend-item legend-rbd">
                    <div class="legend-title" style="color:#005578">🛡️ Redundant By Default — Platform Infrastructure</div>
                    <div class="legend-scope">Azure platform services Microsoft operates with built-in redundancy — no customer action required for zone resilience</div>
                    <div class="legend-examples"><span>VNet</span><span>NSG</span><span>NIC</span><span>Route Table</span><span>Private DNS Zone</span><span>Private Endpoint</span><span>Key Vault</span><span>Log Analytics</span><span>Recovery Services Vault</span><span>Automation Account</span><span>App Insights</span><span>Logic Apps</span></div>
                </div>
                <div class="legend-item legend-geo">
                    <div class="legend-title" style="color:#004c54">🌍 Geo-Redundant — Cross-Region Storage</div>
                    <div class="legend-scope">Storage accounts using GRS/GZRS — data replicated to a secondary Azure region automatically</div>
                    <div class="legend-examples"><span>Storage (Standard_GRS)</span><span>Storage (Standard_RAGRS)</span><span>Storage (Standard_GZRS)</span><span>Storage (Standard_RAGZRS)</span></div>
                </div>
            </div>
            <div class="legend-rbd-resources">
                <div class="rbd-title">🛡️ Complete List: "Microsoft Manages This" (RedundantByDefault) Resource Types</div>
                <div class="rbd-list">
                    <span>microsoft.network/virtualnetworks</span><span>microsoft.network/networkinterfaces</span><span>microsoft.network/networksecuritygroups</span><span>microsoft.network/routetables</span><span>microsoft.network/privatednszone</span><span>microsoft.network/privateendpoints</span><span>microsoft.network/publicipaddresses (Basic)</span><span>microsoft.network/natgateways</span><span>microsoft.keyvault/vaults</span><span>microsoft.operationalinsights/workspaces</span><span>microsoft.insights/components</span><span>microsoft.insights/activitylogalerts</span><span>microsoft.automation/automationaccounts</span><span>microsoft.logic/workflows (Consumption)</span><span>microsoft.recoveryservices/vaults</span><span>microsoft.network/dnszones</span><span>microsoft.network/firewallpolicies</span><span>microsoft.network/virtualnetworklinks</span><span>microsoft.network/restorepointcollections</span><span>microsoft.compute/snapshots</span>
                </div>
                <p style="margin-top:6px; font-size:0.82em; color:#555;">These services are managed by the Azure platform fabric. Zone resilience is handled by Microsoft. Your responsibility is limited to ensuring they are included in your cross-region DR plan if a full-region failure is a concern.</p>
            </div>
        </div>

        <!-- Cross-Region DR Reference Note (replaces misleading "DR Ready" cards) -->
        <div class="bcdr-note-box" style="margin-top: 20px;">
            <h4>🌍 Cross-Region Disaster Recovery — What Does This Mean For You?</h4>
            <p>DR readiness is not a simple yes/no metric — it depends on your specific RPO/RTO targets, workload criticality, budget, and compliance requirements. A resource being "single-region" is not necessarily wrong; it depends on business context.</p>
            <p>For a detailed, resource-by-resource BCDR analysis with specific Microsoft recommendations, action plans, RPO/RTO targets, and effort estimates, please open the <strong>BCDR Assessment Report Excel file (SA_Recommendations sheet)</strong> generated by Phase 2 of this assessment.</p>
            <p style="font-size: 0.88em; color: #666;">⚠️ <strong>Qatar Central note:</strong> Qatar Central has <strong>no Azure paired region</strong> due to data residency requirements. Cross-region DR must be manually planned — recommended target regions are <strong>Sweden Central</strong> (via Azure Backup Region of Choice) or <strong>UAE North</strong> for workload DR via Azure Site Recovery.</p>
        </div>

        <!-- Zone Redundancy Recommendations -->
        <div class="recommendations">
            <h4>🔷 Zone Redundancy Recommendations</h4>
            <ul>
                <li><strong>$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' -and $_.type -eq 'microsoft.compute/virtualmachines' }).Count)</strong> Virtual Machines are not deployed in availability zones - consider redeploying for zone redundancy</li>
                <li><strong>$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' -and $_.type -eq 'Microsoft.Storage/storageAccounts' }).Count)</strong> Storage Accounts use LRS - consider upgrading to ZRS for zone redundancy</li>
                <li><strong>$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' -and $_.type -like '*loadBalancers' }).Count)</strong> Load Balancers are non-zonal - ensure Standard SKU with zone configuration</li>
                <li><strong>$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' -and $_.type -eq 'Microsoft.Compute/disks' }).Count)</strong> Managed Disks use LRS - consider ZRS for critical workloads</li>
            </ul>
        </div>

        <!-- Cross-Region DR Recommendations -->
        <div class="recommendations" style="background: #e8f6f7; border-left-color: #004c54;">
            <h4>🌍 Cross-Region Disaster Recovery Analysis</h4>
            <p style="margin-bottom: 10px; font-size: 0.9em; color: #666;"><strong>Important for Qatar Central:</strong> Qatar Central has NO paired region due to data residency requirements. Manual cross-region DR planning to UAE North or other regions is essential.</p>
            
            <!-- BCDR_REPORT_PLACEHOLDER_START -->
            <div style="background: #fff4ce; padding: 12px; margin-bottom: 15px; border-radius: 8px; border-left: 4px solid #ca5010;">
                <strong>📋 Note:</strong> The recommendations below provide a high-level overview. For detailed, service-specific BCDR guidance tailored to Qatar constraints, including implementation steps, effort estimates, and compliance notes, please refer to the comprehensive <strong>BCDR Assessment Report</strong> generated by Phase 2 of this assessment.
            </div>
            <!-- BCDR_REPORT_PLACEHOLDER_END -->
            
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; margin-bottom: 15px;">
                <div style="background: #fff8e5; padding: 12px; border-radius: 8px; border-left: 4px solid #ca5010;">
                    <h5 style="margin: 0 0 8px 0; color: #ca5010;">🏴 Qatar Central - No Paired Region</h5>
                    <p style="margin: 0; font-size: 0.9em;"><strong>$(($allResources | Where-Object { $_.location -eq 'qatarcentral' }).Count)</strong> resources require manual DR planning</p>
                    <ul style="margin: 8px 0 0 0; padding-left: 20px; font-size: 0.85em;">
                        <li><strong>Data Residency:</strong> No automatic geo-replication available</li>
                        <li><strong>Recommended DR Target:</strong> UAE North, West Europe, or North Europe</li>
                        <li><strong>Action:</strong> Configure Azure Site Recovery, manual backup replication</li>
                    </ul>
                </div>
                <div style="background: white; padding: 14px; border-radius: 8px; border-left: 4px solid #0078d4;">
                    <h5 style="margin: 0 0 10px 0; color: #0078d4;">📊 Cross-Region Status Breakdown
                        <span style="font-size:0.78em; font-weight:400; color:#666; display:block; margin-top:2px;">All <strong>$(($allResources).Count)</strong> resources classified — categories are mutually exclusive and sum to total.</span>
                    </h5>
                    <table style="width:100%; border-collapse:collapse; font-size:0.84em;">
                        <thead>
                            <tr style="background:#f0f7ff; border-bottom:2px solid #cce4f6;">
                                <th style="padding:6px 8px; text-align:left; font-weight:600;">Category</th>
                                <th style="padding:6px 8px; text-align:center; font-weight:600;">Count</th>
                                <th style="padding:6px 8px; text-align:left; font-weight:600;">How determined</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr style="border-bottom:1px solid #e8f0f8;">
                                <td style="padding:5px 8px;">🌍 <strong style="color:#107c10;">Geo-Redundant</strong></td>
                                <td style="padding:5px 8px; text-align:center; font-weight:700; color:#107c10;">$(($allResources | Where-Object { $_.CrossRegionStatus -eq 'Geo-Redundant' }).Count)</td>
                                <td style="padding:5px 8px; color:#555; font-size:0.9em;">Storage with GRS/GZRS/RA-GRS SKU, or Recovery Services Vault with GRS + Cross-Region Restore enabled, or ACR with geo-replication, or Redis with linked servers</td>
                            </tr>
                            <tr style="border-bottom:1px solid #e8f0f8; background:#fafcff;">
                                <td style="padding:5px 8px;">🌐 <strong style="color:#0078d4;">Global / Multi-Region</strong></td>
                                <td style="padding:5px 8px; text-align:center; font-weight:700; color:#0078d4;">$(($allResources | Where-Object { $_.CrossRegionStatus -in @('Global','Multi-Region') }).Count)</td>
                                <td style="padding:5px 8px; color:#555; font-size:0.9em;">Resources at location=<em>global</em> (Traffic Manager, CDN, Front Door, DNS Zones), or Cosmos DB with multiple write regions, or API Management with additional locations</td>
                            </tr>
                            <tr style="border-bottom:1px solid #e8f0f8;">
                                <td style="padding:5px 8px;">⚠️ <strong style="color:#c50f1f;">Single-Region</strong></td>
                                <td style="padding:5px 8px; text-align:center; font-weight:700; color:#c50f1f;">$(($allResources | Where-Object { $_.CrossRegionStatus -eq 'Single-Region' }).Count)</td>
                                <td style="padding:5px 8px; color:#555; font-size:0.9em;">Resource type was evaluated and found to have no cross-region replication configured (e.g., VMs without ASR, Storage with LRS/ZRS, single-region Cosmos DB)</td>
                            </tr>
                            <tr style="border-bottom:1px solid #e8f0f8; background:#fafcff;">
                                <td style="padding:5px 8px;">🔍 <strong style="color:#835b00;">Needs Manual Check</strong></td>
                                <td style="padding:5px 8px; text-align:center; font-weight:700; color:#835b00;">$(($allResources | Where-Object { $_.CrossRegionStatus -eq 'Unknown' }).Count)</td>
                                <td style="padding:5px 8px; color:#555; font-size:0.9em;">Service Bus Premium and Event Hub (Geo-DR config not readable via Resource Graph API — requires portal or ARM API check). Also SQL DB failover groups require separate API call.</td>
                            </tr>
                            <tr style="background:#f3f5f7; font-weight:700; border-top:2px solid #d0d5dd;">
                                <td style="padding:6px 8px;">∑ Total</td>
                                <td style="padding:6px 8px; text-align:center; color:#002050;">$(($allResources).Count)</td>
                                <td style="padding:6px 8px; color:#666; font-weight:400; font-size:0.88em;">Sum of all 4 categories above = total resource count</td>
                            </tr>
                        </tbody>
                    </table>
                    <p style="margin-top:8px; font-size:0.78em; color:#777; font-style:italic;">Note: Cross-region status is assessed per resource type using properties available in Azure Resource Graph. Resources whose DR config cannot be read via Graph API are marked "Needs Manual Check".</p>
                </div>
            </div>
            
            <h5 style="margin: 15px 0 10px 0;">🔧 Action Items by Resource Type:</h5>
            <table style="width: 100%; border-collapse: collapse; font-size: 0.85em; background: white;">
                <thead>
                    <tr style="background: #e8f6f7;">
                        <th style="padding: 8px; text-align: left; border: 1px solid #cce5e8;">Resource Type</th>
                        <th style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">Total</th>
                        <th style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">Single Region</th>
                        <th style="padding: 8px; text-align: left; border: 1px solid #cce5e8;">DR Recommendation</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">💾 Storage Accounts</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Storage/storageAccounts' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Storage/storageAccounts' -and $_.GeoRedundant -ne $true }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Enable GRS/GZRS/RA-GRS</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">🖥️ Virtual Machines</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Compute/virtualMachines' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Compute/virtualMachines' -and $_.CrossRegionStatus -eq 'Single-Region' }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Configure Azure Site Recovery (ASR)</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">🗄️ SQL Databases</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Sql/servers/databases' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Sql/servers/databases' -and $_.CrossRegionStatus -ne 'Geo-Redundant' }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Enable geo-replication / failover groups</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">🌐 Cosmos DB</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.DocumentDB/databaseAccounts' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.DocumentDB/databaseAccounts' -and $_.CrossRegionStatus -eq 'Single-Region' }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Enable multi-region writes</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">🔑 Key Vaults</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.KeyVault/vaults' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.KeyVault/vaults' }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Azure Backup or manual replication</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">🌍 App Services</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Web/sites' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.Web/sites' -and $_.CrossRegionStatus -eq 'Single-Region' }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Deploy to multiple regions + Traffic Manager</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">🔄 Recovery Services Vaults</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.RecoveryServices/vaults' }).Count)</td>
                        <td style="padding: 8px; text-align: center; border: 1px solid #cce5e8; color: #d83b01;">$(($allResources | Where-Object { $_.type -eq 'Microsoft.RecoveryServices/vaults' -and $_.GeoRedundant -ne $true }).Count)</td>
                        <td style="padding: 8px; border: 1px solid #cce5e8;">Enable GRS + Cross-Region Restore</td>
                    </tr>
                </tbody>
            </table>
            
            <div style="margin-top: 15px; padding: 10px; background: #fff4ce; border-radius: 5px; font-size: 0.85em;">
                <strong>⚡ Quick Reference:</strong> <a href="https://learn.microsoft.com/en-us/azure/reliability/cross-region-replication-azure" target="_blank" style="color: #0078d4;">Azure Paired Regions Documentation</a> | 
                <a href="https://learn.microsoft.com/en-us/azure/site-recovery/azure-to-azure-tutorial-enable-replication" target="_blank" style="color: #0078d4;">Azure Site Recovery Setup</a> | 
                <a href="https://learn.microsoft.com/en-us/azure/storage/common/storage-redundancy" target="_blank" style="color: #0078d4;">Storage Redundancy Options</a>
            </div>
        </div>

        <!-- ═══ CHARTS ROW 1 — Zone Distribution + Subscription Breakdown ═══ -->
        <div class="section-header">📈 Charts &amp; Analysis</div>
        <div class="charts-row">
            <div class="chart-container">
                <h2>Zone Resilience Distribution</h2>
                <p class="chart-subtitle">Proportion of resources in each resilience category</p>
                <canvas id="pieChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>Zone Status by Subscription</h2>
                <p class="chart-subtitle">Stacked view — green = protected, red = needs attention</p>
                <canvas id="barChart"></canvas>
            </div>
        </div>

        <!-- ═══ CHARTS ROW 2 — Risk + Resource Types ═══════════════════════ -->
        <div class="charts-row">
            <div class="chart-container">
                <h2>Top Resource Types by Count</h2>
                <p class="chart-subtitle">Most common Azure resource types across all subscriptions</p>
                <canvas id="resourceTypeChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>Risk Assessment — Zone Exposure</h2>
                <p class="chart-subtitle">High risk = critical resources with no zone protection</p>
                <canvas id="riskChart"></canvas>
            </div>
        </div>

        <!-- ═══ CHARTS ROW 3 — Cross-Region + Regional Distribution ═════════ -->
        <div class="charts-row">
            <div class="chart-container">
                <h2>Cross-Region Replication Status</h2>
                <p class="chart-subtitle">Geo-redundant vs single-region vs global resources</p>
                <canvas id="crossRegionChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>Regional Distribution</h2>
                <p class="chart-subtitle">Resource count per Azure region</p>
                <canvas id="regionChart"></canvas>
            </div>
        </div>

        <!-- ═══ CHARTS ROW 4 (NEW) — NonZonal Breakdown + IaaS vs PaaS ═════ -->
        <div class="charts-row">
            <div class="chart-container">
                <h2>Non-Zonal Resources — by Type</h2>
                <p class="chart-subtitle">Which resource types make up the Non-Zonal category (top 12)</p>
                <canvas id="nonZonalBreakdownChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>IaaS vs PaaS vs Platform Distribution</h2>
                <p class="chart-subtitle">Breakdown of resource scope across the estate</p>
                <canvas id="iaasPaasChart"></canvas>
            </div>
        </div>

        <!-- ═══ CHARTS ROW 5 (NEW) — Subscription Risk Scores ══════════════ -->
        <div class="charts-row">
            <div class="chart-container" style="grid-column: 1 / -1;">
                <h2>Subscription Risk Score — Zone Exposure</h2>
                <p class="chart-subtitle">Normalized risk score per subscription (0–100). Score = weighted sum of NonZonal×3 + LocallyRedundant×2 + Zonal×1 relative to total resources. Higher = more exposure.</p>
                <canvas id="subRiskChart" style="max-height: 300px;"></canvas>
            </div>
        </div>

        <!-- Filter Bar — Multi-Select Checkbox Dropdowns -->
        <div class="filter-bar">
            <div class="filter-group">
                <label>Search:</label>
                <input type="text" id="searchInput" placeholder="Search resources..." oninput="applyFilters()" style="width:220px;">
            </div>
            <div class="filter-group">
                <label>Subscription:</label>
                <div class="mdd-wrapper" id="sub-mdd">
                    <button class="mdd-btn" onclick="toggleMDD(event,'sub-mdd')">
                        <span class="mdd-label">All Subscriptions</span><span class="mdd-arrow">▾</span>
                    </button>
                    <div class="mdd-panel">
                        <div class="mdd-actions">
                            <button onclick="selectAllMDD('sub-mdd')">Select All</button>
                            <button onclick="clearMDD('sub-mdd')">Clear All</button>
                        </div>
                        <input type="text" class="mdd-search" placeholder="Search subscriptions..." oninput="searchMDD(this)">
                        <div class="mdd-list" id="sub-mdd-list"></div>
                    </div>
                </div>
            </div>
            <div class="filter-group">
                <label>Zone Status:</label>
                <div class="mdd-wrapper" id="status-mdd">
                    <button class="mdd-btn" onclick="toggleMDD(event,'status-mdd')">
                        <span class="mdd-label">All Statuses</span><span class="mdd-arrow">▾</span>
                    </button>
                    <div class="mdd-panel">
                        <div class="mdd-actions">
                            <button onclick="selectAllMDD('status-mdd')">Select All</button>
                            <button onclick="clearMDD('status-mdd')">Clear All</button>
                        </div>
                        <div class="mdd-list" id="status-mdd-list"></div>
                    </div>
                </div>
            </div>
            <div class="filter-group">
                <label>Resource Type:</label>
                <div class="mdd-wrapper" id="type-mdd">
                    <button class="mdd-btn" onclick="toggleMDD(event,'type-mdd')">
                        <span class="mdd-label">All Types</span><span class="mdd-arrow">▾</span>
                    </button>
                    <div class="mdd-panel">
                        <div class="mdd-actions">
                            <button onclick="selectAllMDD('type-mdd')">Select All</button>
                            <button onclick="clearMDD('type-mdd')">Clear All</button>
                        </div>
                        <input type="text" class="mdd-search" placeholder="Search resource types..." oninput="searchMDD(this)">
                        <div class="mdd-list" id="type-mdd-list"></div>
                    </div>
                </div>
            </div>
            <div class="filter-group">
                <label>Location:</label>
                <div class="mdd-wrapper" id="loc-mdd">
                    <button class="mdd-btn" onclick="toggleMDD(event,'loc-mdd')">
                        <span class="mdd-label">All Locations</span><span class="mdd-arrow">▾</span>
                    </button>
                    <div class="mdd-panel">
                        <div class="mdd-actions">
                            <button onclick="selectAllMDD('loc-mdd')">Select All</button>
                            <button onclick="clearMDD('loc-mdd')">Clear All</button>
                        </div>
                        <div class="mdd-list" id="loc-mdd-list"></div>
                    </div>
                </div>
            </div>
            <div class="filter-group">
                <label>Resource Group:</label>
                <div class="mdd-wrapper" id="rg-mdd">
                    <button class="mdd-btn" onclick="toggleMDD(event,'rg-mdd')">
                        <span class="mdd-label">All Resource Groups</span><span class="mdd-arrow">▾</span>
                    </button>
                    <div class="mdd-panel">
                        <div class="mdd-actions">
                            <button onclick="selectAllMDD('rg-mdd')">Select All</button>
                            <button onclick="clearMDD('rg-mdd')">Clear All</button>
                        </div>
                        <input type="text" class="mdd-search" placeholder="Search resource groups..." oninput="searchMDD(this)">
                        <div class="mdd-list" id="rg-mdd-list"></div>
                    </div>
                </div>
            </div>
            <button class="btn btn-secondary" onclick="toggleColumnSelector()">Columns</button>
            <button class="btn btn-secondary" onclick="resetFilters()">Reset All</button>
            <button class="btn btn-primary" onclick="exportTableToCSV()">Export View</button>
        </div>

        <!-- Column Selector -->
        <div id="columnSelector" class="column-selector">
            <h4>Select Visible Columns: <span style="font-weight:normal; font-size:0.85em; color:#666;">($($allColumns.Count) columns available)</span></h4>
            <div class="column-selector-actions">
                <button class="btn btn-secondary" onclick="selectAllColumns()">Select All</button>
                <button class="btn btn-secondary" onclick="deselectAllColumns()">Deselect All</button>
                <button class="btn btn-secondary" onclick="resetColumns()">Reset to Default</button>
            </div>
$columnSelectorHtml
        </div>

        <!-- Resources Table -->
        <div class="table-container">
            <h2>All Resources <span class="resource-count" id="visibleCount">$($allResources.Count)</span></h2>
            <p style="font-size: 0.85em; color: #666; margin: -10px 0 10px 0;"><strong>Tip:</strong> Click a cell to expand/collapse text. Double-click column header to expand all cells in that column. Drag column header edges to resize.</p>
            <div class="table-wrapper">
                <table id="resourceTable">
                    <thead>
                        <tr>
$tableHeaderHtml
                        </tr>
                    </thead>
                    <tbody>
"@

# Add table rows dynamically
foreach ($resource in ($allResources | Sort-Object SubscriptionName, ZoneRedundancyStatus, type, name)) {
    $statusClass = "status-$($resource.ZoneRedundancyStatus -replace '\s','')"
    $riskClass = switch ($resource.ZoneRedundancyStatus) {
        'NonZonal' { 'risk-high' }
        'LocallyRedundant' { 'risk-medium' }
        default { '' }
    }
    
    # Build row with all columns dynamically
    $rowHtml = "                        <tr class=`"$riskClass`" data-subscription=`"$($resource.SubscriptionName)`" data-status=`"$($resource.ZoneRedundancyStatus)`" data-type=`"$($resource.type)`" data-location=`"$($resource.location)`" data-resourcegroup=`"$($resource.resourceGroup)`" data-crossregion=`"$($resource.CrossRegionStatus)`">`n"
    
    $colIndex = 0
    foreach ($col in $allColumns) {
        $colName = $col.Name
        $value = $resource.$colName
        $escapedValue = [System.Web.HttpUtility]::HtmlEncode($value)
        $displayStyle = if ($col.DefaultVisible) { "" } else { " style=`"display:none;`"" }
        
        # Special formatting for certain columns
        if ($colName -eq 'name') {
            $rowHtml += "                            <td data-col=`"$colIndex`"$displayStyle><strong>$escapedValue</strong></td>`n"
        }
        elseif ($colName -eq 'ZoneRedundancyStatus') {
            $rowHtml += "                            <td data-col=`"$colIndex`"$displayStyle><span class=`"status-badge $statusClass`">$escapedValue</span></td>`n"
        }
        elseif ($colName -eq 'CrossRegionStatus') {
            $crossRegionClass = switch ($value) {
                'Geo-Redundant' { 'status-zr' }
                'Multi-Region' { 'status-zonal' }
                'Global' { 'status-default' }
                'Single-Region' { 'status-nonzonal' }
                default { 'status-unknown' }
            }
            $rowHtml += "                            <td data-col=`"$colIndex`"$displayStyle><span class=`"status-badge $crossRegionClass`">$escapedValue</span></td>`n"
        }
        elseif ($colName -eq 'ZoneRedundancyDetail' -or $colName -eq 'CrossRegionDetail') {
            $rowHtml += "                            <td data-col=`"$colIndex`"$displayStyle title=`"$escapedValue`">$escapedValue</td>`n"
        }
        else {
            $rowHtml += "                            <td data-col=`"$colIndex`"$displayStyle>$escapedValue</td>`n"
        }
        $colIndex++
    }
    
    $rowHtml += "                        </tr>"
    $htmlReport += $rowHtml
}

# Calculate chart data
$zoneRedundantCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'ZoneRedundant' }).Count
$zonalCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'Zonal' }).Count
$nonZonalCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' }).Count
$locallyRedundantCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' }).Count
$redundantByDefaultCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'RedundantByDefault' }).Count
$geoRedundantCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -like '*Geo*' }).Count
$unknownCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'Unknown' -or $_.ZoneRedundancyStatus -eq 'InheritedFromParent' }).Count

# Subscription data for bar chart
$subLabels = ($subSummary | ForEach-Object { "`"$($_.SubscriptionName.Substring(0, [Math]::Min(20, $_.SubscriptionName.Length)))`"" }) -join ','
$subZR = ($subSummary | ForEach-Object { $_.ZoneRedundant }) -join ','
$subNonZonal = ($subSummary | ForEach-Object { $_.NonZonal }) -join ','
$subLR = ($subSummary | ForEach-Object { $_.LocallyRedundant }) -join ','

# Resource type data (top 10)
$topTypes = $allResources | Group-Object type | Sort-Object Count -Descending | Select-Object -First 10
$typeLabels = ($topTypes | ForEach-Object { "`"$(($_.Name -split '/')[1])`"" }) -join ','
$typeCounts = ($topTypes | ForEach-Object { $_.Count }) -join ','

# Risk data
$criticalResources = @('microsoft.compute/virtualmachines', 'Microsoft.Sql/servers/databases', 'Microsoft.Storage/storageAccounts')
$highRiskCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -in @('NonZonal','LocallyRedundant') -and $_.type -in $criticalResources }).Count
$mediumRiskCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -in @('NonZonal','LocallyRedundant') -and $_.type -notin $criticalResources }).Count
$lowRiskCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -in @('ZoneRedundant','RedundantByDefault','Zonal') }).Count

# Cross-region data — 4 mutually exclusive categories that always sum to total resource count
# Geo-Redundant : Storage GRS/GZRS, Recovery Vault GRS+CRR, ACR geo-rep, Redis linked servers
# Global/Multi   : location='global' (TM,CDN,FrontDoor,DNS) + Cosmos DB multi-write + APIM multi-region
# Single-Region  : evaluated and confirmed no cross-region config
# Unknown        : config not readable via Resource Graph (SQL failover groups, Service Bus Geo-DR, Event Hub Geo-DR)
$crossRegionGeoRedundantCount    = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Geo-Redundant' }).Count
$crossRegionGlobalMultiCount     = ($allResources | Where-Object { $_.CrossRegionStatus -in @('Global','Multi-Region') }).Count
$crossRegionSingleCount          = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Single-Region' }).Count
$crossRegionUnknownCount         = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Unknown' }).Count
# Sanity: these four should equal total
$crossRegionTotal = $crossRegionGeoRedundantCount + $crossRegionGlobalMultiCount + $crossRegionSingleCount + $crossRegionUnknownCount

# Regional distribution data (top 10 regions)
$topRegions = $allResources | Group-Object location | Sort-Object Count -Descending | Select-Object -First 10
$regionLabels = ($topRegions | ForEach-Object { "`"$($_.Name)`"" }) -join ','
$regionCounts = ($topRegions | ForEach-Object { $_.Count }) -join ','

# NEW: Non-Zonal breakdown by resource type (top 12)
$topNonZonalTypes = $allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' } |
    Group-Object type | Sort-Object Count -Descending | Select-Object -First 12
$nonZonalTypeLabels = ($topNonZonalTypes | ForEach-Object { "`"$(($_.Name -split '/')[1])`"" }) -join ','
$nonZonalTypeCounts  = ($topNonZonalTypes | ForEach-Object { $_.Count }) -join ','

# NEW: IaaS vs PaaS vs Platform category counts
$iaasTypes  = @('microsoft.compute/virtualmachines','microsoft.compute/disks','microsoft.compute/virtualmachinescalesets','microsoft.network/loadbalancers','microsoft.network/applicationgateways','microsoft.network/publicipaddresses')
$platTypes  = @('microsoft.network/virtualnetworks','microsoft.network/networkinterfaces','microsoft.network/networksecuritygroups','microsoft.network/routetables','microsoft.keyvault/vaults','microsoft.operationalinsights/workspaces','microsoft.recoveryservices/vaults','microsoft.automation/automationaccounts','microsoft.insights/components','microsoft.logic/workflows')
$iaasCount  = ($allResources | Where-Object { $_.type.ToLower() -in $iaasTypes }).Count
$platCount  = ($allResources | Where-Object { $_.type.ToLower() -in $platTypes }).Count
$paasCount  = $allResources.Count - $iaasCount - $platCount
if ($paasCount -lt 0) { $paasCount = 0 }

# NEW: Subscription risk scores (weighted: NonZonal×3 + LocallyRedundant×2 + Zonal×1 / total×3 × 100)
$subRiskLabels = ($subSummary | ForEach-Object {
    "`"$($_.SubscriptionName.Substring(0, [Math]::Min(25, $_.SubscriptionName.Length)))`""
}) -join ','
$subRiskScores = ($subSummary | ForEach-Object {
    $total = $_.TotalResources
    if ($total -eq 0) { "0" }
    else { [Math]::Round((($_.NonZonal * 3 + $_.LocallyRedundant * 2 + $_.Zonal * 1) / ($total * 3)) * 100, 1) }
}) -join ','

# Prepare zone mappings data for the table
$zoneMappingsWithAZ = $allZoneMappings | Where-Object { $_.SupportsAZ -eq $true } | Sort-Object SubscriptionName, Location, LogicalZone

$htmlReport += @"
                </tbody>
            </table>
        </div>

        <!-- Zone Mappings Table -->
        <div class="table-container">
            <h2>Availability Zone Mappings (Logical to Physical) <span class="resource-count" id="zmCount">Loading...</span></h2>
            <p style="font-size: 0.85em; color: #666; margin: -10px 0 10px 0;"><strong>Note:</strong> Zone mappings differ per subscription. Logical Zone 1 in Sub A may map to a different physical zone than Logical Zone 1 in Sub B.</p>
            <div style="margin-bottom: 10px; display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">
                <label>Filter: <input type="text" id="zmSearch" placeholder="Search location..." onkeyup="filterZoneMappings()" style="padding: 6px; border: 1px solid #ddd; border-radius: 4px;"></label>
                <label>Subscription: <select id="zmSubFilter" onchange="filterZoneMappings()" style="padding: 6px; border: 1px solid #ddd; border-radius: 4px;">
                    <option value="">All Subscriptions</option>
                    $($uniqueSubscriptions | ForEach-Object { "<option value=`"$_`">$_</option>" })
                </select></label>
                <label>Region: <select id="zmRegionFilter" onchange="filterZoneMappings()" style="padding: 6px; border: 1px solid #ddd; border-radius: 4px;">
                    <option value="">All Regions</option>
                    $($zoneMappingsWithAZ | Select-Object -ExpandProperty Location -Unique | Sort-Object | ForEach-Object { "<option value=`"$_`">$_</option>" })
                </select></label>
            </div>
            <div class="table-wrapper" style="max-height: 400px;">
            <table id="zoneMappingTable">
                <thead>
                    <tr>
                        <th>Subscription</th>
                        <th>Location</th>
                        <th>Display Name</th>
                        <th>Logical Zone</th>
                        <th>Physical Zone</th>
                    </tr>
                </thead>
                <tbody>
"@

# Add ALL zone mapping rows that support AZ
foreach ($zm in $zoneMappingsWithAZ) {
    $htmlReport += @"
                    <tr data-subscription="$($zm.SubscriptionName)" data-location="$($zm.Location)">
                        <td>$($zm.SubscriptionName)</td>
                        <td>$($zm.Location)</td>
                        <td>$($zm.DisplayName)</td>
                        <td>Zone $($zm.LogicalZone)</td>
                        <td>$($zm.PhysicalZone)</td>
                    </tr>
"@
}

$htmlReport += @"
                </tbody>
            </table>
            </div>
        </div>
    </div>

    <div class="footer">
        <div style="margin-bottom: 6px;">
            <svg width="14" height="14" viewBox="0 0 21 21" fill="none" style="vertical-align: middle; margin-right: 5px;" xmlns="http://www.w3.org/2000/svg">
                <rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
                <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>
            <strong style="color:white;">Microsoft Azure — BCDR &amp; Zone Resilience Assessment</strong>
        </div>
        <div>
            Solution Engineering &nbsp;·&nbsp; Infrastructure &nbsp;·&nbsp; Dashboard v2.0 &nbsp;·&nbsp;
            <a href="https://learn.microsoft.com/azure/reliability/availability-zones-overview" target="_blank">Azure Availability Zones docs ↗</a>
        </div>
        <div style="margin-top: 6px; font-size: 0.75em; opacity: 0.55;">
            This report is for internal assessment purposes only. Data is sourced via Azure Resource Graph at report generation time.
            Re-run the assessment to capture environment changes.
        </div>
    </div>

    <script>
        // ── Global chart references (updated on filter change) ──────────────
        var gPieChart, gBarChart, gResourceTypeChart, gRiskChart, gCrossRegionChart, gRegionChart;
        var gNonZonalBreakdownChart, gIaasPaasChart, gSubRiskChart;

        // ── Percentage label plugin for doughnut charts ──────────────────────
        const percentagePlugin = {
            id: 'percentageLabels',
            afterDatasetsDraw(chart) {
                const ctx = chart.ctx;
                const datasets = chart.data.datasets;
                const total = datasets[0].data.reduce((a, b) => a + b, 0);
                if (total === 0) return;
                chart.getDatasetMeta(0).data.forEach((arc, index) => {
                    const value = datasets[0].data[index];
                    if (value === 0) return;
                    const percentage = ((value / total) * 100).toFixed(1);
                    const midAngle = (arc.startAngle + arc.endAngle) / 2;
                    const radius = (arc.innerRadius + arc.outerRadius) / 2;
                    const x = arc.x + Math.cos(midAngle) * radius;
                    const y = arc.y + Math.sin(midAngle) * radius;
                    ctx.save();
                    ctx.fillStyle = '#fff';
                    ctx.font = 'bold 11px Arial';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.shadowColor = 'rgba(0,0,0,0.5)';
                    ctx.shadowBlur = 3;
                    ctx.fillText(percentage + '%', x, y);
                    ctx.restore();
                });
            }
        };

        // ── Chart initialisation ─────────────────────────────────────────────
        gPieChart = new Chart(document.getElementById('pieChart'), {
            type: 'doughnut',
            data: {
                labels: ['Resilient Across Datacenters (ZR)', 'Assigned to One Datacenter (Zonal)', 'No Datacenter Assignment (NonZonal)', 'Within One Datacenter Only (LRS)', 'Microsoft Manages Resilience (RBD)', 'Geo Redundant (Cross-Region)', 'Other / Unknown'],
                datasets: [{ data: [$zoneRedundantCount, $zonalCount, $nonZonalCount, $locallyRedundantCount, $redundantByDefaultCount, $geoRedundantCount, $unknownCount], backgroundColor: ['#107c10','#0091ea','#c50f1f','#ca5010','#00bcf2','#006058','#605e5c'] }]
            },
            options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } }, tooltip: { callbacks: { label: function(ctx) { const total = ctx.dataset.data.reduce((a,b)=>a+b,0); return ctx.label+': '+ctx.raw+' ('+((ctx.raw/total)*100).toFixed(1)+'%)'; } } } } },
            plugins: [percentagePlugin]
        });

        gBarChart = new Chart(document.getElementById('barChart'), {
            type: 'bar',
            data: {
                labels: [$subLabels],
                datasets: [
                    { label: 'Zone Redundant (Protected)', data: [$subZR], backgroundColor: '#107c10' },
                    { label: 'Non-Zonal (Needs Attention)', data: [$subNonZonal], backgroundColor: '#c50f1f' },
                    { label: 'Locally Redundant (Storage)', data: [$subLR], backgroundColor: '#ca5010' }
                ]
            },
            options: { responsive: true, scales: { x: { stacked: true }, y: { stacked: true } }, plugins: { legend: { position: 'bottom' } } }
        });

        gResourceTypeChart = new Chart(document.getElementById('resourceTypeChart'), {
            type: 'bar',
            data: { labels: [$typeLabels], datasets: [{ label: 'Count', data: [$typeCounts], backgroundColor: '#0078d4' }] },
            options: { responsive: true, indexAxis: 'y' }
        });

        gRiskChart = new Chart(document.getElementById('riskChart'), {
            type: 'doughnut',
            data: {
                labels: ['High Risk (Critical Non-Zonal)', 'Medium Risk (Other Non-Zonal)', 'Low Risk (Protected)'],
                datasets: [{ data: [$highRiskCount, $mediumRiskCount, $lowRiskCount], backgroundColor: ['#d83b01','#ffb900','#107c10'] }]
            },
            options: { responsive: true, plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: function(ctx) { const total = ctx.dataset.data.reduce((a,b)=>a+b,0); return ctx.label+': '+ctx.raw+' ('+((ctx.raw/total)*100).toFixed(1)+'%)'; } } } } },
            plugins: [percentagePlugin]
        });

        gCrossRegionChart = new Chart(document.getElementById('crossRegionChart'), {
            type: 'doughnut',
            data: {
                labels: [
                    'Geo-Redundant (GRS/GZRS/geo-rep)',
                    'Global / Multi-Region (CDN, TM, Cosmos multi-write)',
                    'Single-Region (no cross-region config)',
                    'Needs Manual Check (SQL failover, SB/EH Geo-DR)'
                ],
                datasets: [{ data: [$crossRegionGeoRedundantCount, $crossRegionGlobalMultiCount, $crossRegionSingleCount, $crossRegionUnknownCount], backgroundColor: ['#107c10','#0091ea','#c50f1f','#835b00'] }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 11 } } },
                    tooltip: { callbacks: { label: function(ctx) {
                        const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
                        return ctx.label+': '+ctx.raw+' ('+(total>0?((ctx.raw/total)*100).toFixed(1):0)+'%)  [sums to '+total+' total]';
                    }}}
                }
            },
            plugins: [percentagePlugin]
        });

        gRegionChart = new Chart(document.getElementById('regionChart'), {
            type: 'bar',
            data: { labels: [$regionLabels], datasets: [{ label: 'Resources', data: [$regionCounts], backgroundColor: '#0078d4' }] },
            options: { responsive: true, indexAxis: 'y', plugins: { legend: { display: false } } }
        });

        // ── NEW CHARTS ────────────────────────────────────────────────────────

        // NonZonal Breakdown by Resource Type
        gNonZonalBreakdownChart = new Chart(document.getElementById('nonZonalBreakdownChart'), {
            type: 'bar',
            data: {
                labels: [$nonZonalTypeLabels],
                datasets: [{ label: 'Non-Zonal Count', data: [$nonZonalTypeCounts], backgroundColor: '#c50f1f' }]
            },
            options: { responsive: true, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true } } }
        });

        // IaaS vs PaaS vs Platform donut
        gIaasPaasChart = new Chart(document.getElementById('iaasPaasChart'), {
            type: 'doughnut',
            data: {
                labels: ['IaaS (VMs, Disks, LB, AppGW)', 'PaaS & App Services', 'Platform Infrastructure (RBD)'],
                datasets: [{ data: [$iaasCount, $paasCount, $platCount], backgroundColor: ['#0091ea','#107c10','#00bcf2'] }]
            },
            options: { responsive: true, plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: function(ctx) { const total = ctx.dataset.data.reduce((a,b)=>a+b,0); return ctx.label+': '+ctx.raw+' ('+((ctx.raw/total)*100).toFixed(1)+'%)'; } } } } },
            plugins: [percentagePlugin]
        });

        // Subscription Risk Score horizontal bar
        gSubRiskChart = new Chart(document.getElementById('subRiskChart'), {
            type: 'bar',
            data: {
                labels: [$subRiskLabels],
                datasets: [{
                    label: 'Risk Score (0–100)',
                    data: [$subRiskScores],
                    backgroundColor: function(ctx) {
                        var v = ctx.raw;
                        if (v >= 60) return '#c50f1f';
                        if (v >= 30) return '#ca5010';
                        return '#107c10';
                    }
                }]
            },
            options: {
                responsive: true,
                indexAxis: 'y',
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: { label: function(ctx) { return 'Risk Score: ' + ctx.raw + '/100'; } } }
                },
                scales: { x: { min: 0, max: 100, title: { display: true, text: 'Risk Score (0=fully protected, 100=fully exposed)' } } }
            }
        });

        // ── Build multi-select dropdown lists from table row data ────────────
        function buildDropdownOptions() {
            const rows = document.querySelectorAll('#resourceTable tbody tr');
            const subs = new Set(), statuses = new Set(), types = new Set(), locs = new Set(), rgs = new Set();
            rows.forEach(function(row) {
                subs.add(row.getAttribute('data-subscription') || '');
                statuses.add(row.getAttribute('data-status') || '');
                types.add(row.getAttribute('data-type') || '');
                locs.add(row.getAttribute('data-location') || '');
                rgs.add(row.getAttribute('data-resourcegroup') || '');
            });
            function buildList(listId, values) {
                const container = document.getElementById(listId);
                if (!container) return;
                const sorted = Array.from(values).filter(Boolean).sort();
                container.innerHTML = sorted.map(function(v) {
                    const escaped = v.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
                    return '<div class="mdd-item"><input type="checkbox" id="cb-'+listId+'-'+escaped+'" value="'+escaped+'" checked onchange="applyFilters()"><label for="cb-'+listId+'-'+escaped+'">'+escaped+'</label></div>';
                }).join('');
            }
            buildList('sub-mdd-list', subs);
            buildList('status-mdd-list', statuses);
            buildList('type-mdd-list', types);
            buildList('loc-mdd-list', locs);
            buildList('rg-mdd-list', rgs);
        }

        // ── Multi-Select Dropdown helpers ────────────────────────────────────
        function toggleMDD(event, id) {
            event.stopPropagation();
            var panel = document.querySelector('#'+id+' .mdd-panel');
            var isOpen = panel.classList.contains('open');
            document.querySelectorAll('.mdd-panel.open').forEach(function(p) { p.classList.remove('open'); });
            if (!isOpen) panel.classList.add('open');
        }
        document.addEventListener('click', function() {
            document.querySelectorAll('.mdd-panel.open').forEach(function(p) { p.classList.remove('open'); });
        });
        function selectAllMDD(id) {
            document.querySelectorAll('#'+id+' .mdd-list input[type="checkbox"]').forEach(function(cb) { cb.checked = true; });
            updateMDDLabel(id); applyFilters();
        }
        function clearMDD(id) {
            document.querySelectorAll('#'+id+' .mdd-list input[type="checkbox"]').forEach(function(cb) { cb.checked = false; });
            updateMDDLabel(id); applyFilters();
        }
        function searchMDD(input) {
            var search = input.value.toLowerCase();
            var panel = input.closest('.mdd-panel');
            panel.querySelectorAll('.mdd-item').forEach(function(item) {
                item.style.display = item.textContent.toLowerCase().includes(search) ? '' : 'none';
            });
        }
        function updateMDDLabel(id) {
            var cbs = document.querySelectorAll('#'+id+' .mdd-list input[type="checkbox"]');
            var checked = Array.from(cbs).filter(function(cb) { return cb.checked; });
            var label = document.querySelector('#'+id+' .mdd-label');
            if (!label) return;
            var defaults = {'sub-mdd':'All Subscriptions','status-mdd':'All Statuses','type-mdd':'All Types','loc-mdd':'All Locations','rg-mdd':'All Resource Groups'};
            if (checked.length === 0 || checked.length === cbs.length) {
                label.textContent = defaults[id] || 'All';
            } else {
                label.textContent = checked.length + ' selected';
            }
        }
        function getCheckedValues(id) {
            var cbs = document.querySelectorAll('#'+id+' .mdd-list input[type="checkbox"]');
            if (cbs.length === 0) return null;
            var checked = Array.from(cbs).filter(function(cb) { return cb.checked; }).map(function(cb) { return cb.value; });
            if (checked.length === 0 || checked.length === cbs.length) return null;
            return checked;
        }

        // ── Main filter + chart/card update function ─────────────────────────
        function applyFilters() {
            var search = (document.getElementById('searchInput').value || '').toLowerCase();
            var subs    = getCheckedValues('sub-mdd');
            var statuses = getCheckedValues('status-mdd');
            var types   = getCheckedValues('type-mdd');
            var locs    = getCheckedValues('loc-mdd');
            var rgs     = getCheckedValues('rg-mdd');

            ['sub-mdd','status-mdd','type-mdd','loc-mdd','rg-mdd'].forEach(updateMDDLabel);

            var rows = document.querySelectorAll('#resourceTable tbody tr');
            var visibleCount = 0;
            var statusCounts = {}, subCounts = {}, typeCounts = {}, locCounts = {}, crossCounts = {};
            var criticalTypes = ['microsoft.compute/virtualmachines','microsoft.sql/servers/databases','microsoft.storage/storageaccounts'];
            var highRisk = 0, medRisk = 0, lowRisk = 0;

            rows.forEach(function(row) {
                var rowSub  = row.getAttribute('data-subscription') || '';
                var rowStat = row.getAttribute('data-status') || '';
                var rowType = row.getAttribute('data-type') || '';
                var rowLoc  = row.getAttribute('data-location') || '';
                var rowRG   = row.getAttribute('data-resourcegroup') || '';
                var rowCross = row.getAttribute('data-crossregion') || '';
                var rowText  = row.textContent.toLowerCase();

                var ok = (!search || rowText.includes(search))
                      && (!subs    || subs.indexOf(rowSub) >= 0)
                      && (!statuses|| statuses.indexOf(rowStat) >= 0)
                      && (!types   || types.indexOf(rowType) >= 0)
                      && (!locs    || locs.indexOf(rowLoc) >= 0)
                      && (!rgs     || rgs.indexOf(rowRG) >= 0);

                row.style.display = ok ? '' : 'none';
                if (ok) {
                    visibleCount++;
                    statusCounts[rowStat] = (statusCounts[rowStat] || 0) + 1;
                    crossCounts[rowCross] = (crossCounts[rowCross] || 0) + 1;
                    typeCounts[rowType]   = (typeCounts[rowType]   || 0) + 1;
                    locCounts[rowLoc]     = (locCounts[rowLoc]     || 0) + 1;
                    if (!subCounts[rowSub]) subCounts[rowSub] = {zr:0,nz:0,lr:0};
                    if (rowStat === 'ZoneRedundant')    subCounts[rowSub].zr++;
                    else if (rowStat === 'NonZonal')    subCounts[rowSub].nz++;
                    else if (rowStat === 'LocallyRedundant') subCounts[rowSub].lr++;
                    // Risk calc
                    var t = rowType.toLowerCase();
                    if (rowStat === 'NonZonal' || rowStat === 'LocallyRedundant') {
                        if (criticalTypes.indexOf(t) >= 0) highRisk++; else medRisk++;
                    } else if (rowStat === 'ZoneRedundant' || rowStat === 'RedundantByDefault' || rowStat === 'Zonal') {
                        lowRisk++;
                    }
                }
            });

            document.getElementById('visibleCount').textContent = visibleCount;

            // Update summary cards
            document.getElementById('card-total').textContent   = visibleCount;
            document.getElementById('card-zr').textContent      = statusCounts['ZoneRedundant'] || 0;
            document.getElementById('card-zonal').textContent   = statusCounts['Zonal'] || 0;
            document.getElementById('card-nonzonal').textContent= statusCounts['NonZonal'] || 0;
            document.getElementById('card-lr').textContent      = statusCounts['LocallyRedundant'] || 0;
            document.getElementById('card-rbd').textContent     = statusCounts['RedundantByDefault'] || 0;

            var geoCount = (statusCounts['GeoRedundant'] || 0) + (statusCounts['GeoZoneRedundant'] || 0);
            var unknownCount = (statusCounts['Unknown'] || 0) + (statusCounts['InheritedFromParent'] || 0);

            // Update Pie chart
            if (gPieChart) {
                gPieChart.data.datasets[0].data = [
                    statusCounts['ZoneRedundant']||0, statusCounts['Zonal']||0,
                    statusCounts['NonZonal']||0, statusCounts['LocallyRedundant']||0,
                    statusCounts['RedundantByDefault']||0, geoCount, unknownCount
                ];
                gPieChart.update();
            }
            if (gBarChart) {
                var subNames = Object.keys(subCounts);
                gBarChart.data.labels = subNames.map(function(s) { return s.substring(0,20); });
                gBarChart.data.datasets[0].data = subNames.map(function(s) { return subCounts[s].zr; });
                gBarChart.data.datasets[1].data = subNames.map(function(s) { return subCounts[s].nz; });
                gBarChart.data.datasets[2].data = subNames.map(function(s) { return subCounts[s].lr; });
                gBarChart.update();
            }

            // Update Resource Type chart (top 10)
            if (gResourceTypeChart) {
                var topTypes = Object.entries(typeCounts).sort(function(a,b){return b[1]-a[1];}).slice(0,10);
                gResourceTypeChart.data.labels = topTypes.map(function(e) { return e[0].split('/').slice(-1)[0]; });
                gResourceTypeChart.data.datasets[0].data = topTypes.map(function(e) { return e[1]; });
                gResourceTypeChart.update();
            }

            // Update Risk chart
            if (gRiskChart) {
                gRiskChart.data.datasets[0].data = [highRisk, medRisk, lowRisk];
                gRiskChart.update();
            }

            // Update Cross-Region chart (4 mutually exclusive categories — must sum to total)
            if (gCrossRegionChart) {
                gCrossRegionChart.data.datasets[0].data = [
                    crossCounts['Geo-Redundant']||0,
                    (crossCounts['Multi-Region']||0)+(crossCounts['Global']||0),
                    crossCounts['Single-Region']||0,
                    crossCounts['Unknown']||0
                ];
                gCrossRegionChart.update();
            }

            // Update Region chart (top 10)
            if (gRegionChart) {
                var topLocs = Object.entries(locCounts).sort(function(a,b){return b[1]-a[1];}).slice(0,10);
                gRegionChart.data.labels = topLocs.map(function(e) { return e[0]; });
                gRegionChart.data.datasets[0].data = topLocs.map(function(e) { return e[1]; });
                gRegionChart.update();
            }

            // Update NonZonal Breakdown chart (top 12 non-zonal types)
            if (gNonZonalBreakdownChart) {
                var nzCounts = {};
                rows.forEach(function(row) {
                    if (row.style.display === 'none') return;
                    if ((row.getAttribute('data-status')||'') === 'NonZonal') {
                        var t = (row.getAttribute('data-type')||'').split('/').pop();
                        nzCounts[t] = (nzCounts[t]||0) + 1;
                    }
                });
                var topNZ = Object.entries(nzCounts).sort(function(a,b){return b[1]-a[1];}).slice(0,12);
                gNonZonalBreakdownChart.data.labels = topNZ.map(function(e){return e[0];});
                gNonZonalBreakdownChart.data.datasets[0].data = topNZ.map(function(e){return e[1];});
                gNonZonalBreakdownChart.update();
            }

            // Update Sub Risk chart
            if (gSubRiskChart) {
                var subRisk = {};
                var subTotal = {};
                rows.forEach(function(row) {
                    if (row.style.display === 'none') return;
                    var sub = (row.getAttribute('data-subscription')||'Unknown').substring(0,25);
                    var st  = row.getAttribute('data-status')||'';
                    if (!subRisk[sub]) { subRisk[sub]=0; subTotal[sub]=0; }
                    subTotal[sub]++;
                    if (st==='NonZonal') subRisk[sub]+=3;
                    else if (st==='LocallyRedundant') subRisk[sub]+=2;
                    else if (st==='Zonal') subRisk[sub]+=1;
                });
                var subNames = Object.keys(subRisk);
                gSubRiskChart.data.labels = subNames;
                gSubRiskChart.data.datasets[0].data = subNames.map(function(s) {
                    return subTotal[s]===0 ? 0 : Math.round((subRisk[s]/(subTotal[s]*3))*1000)/10;
                });
                gSubRiskChart.update();
            }
        }

        // Keep legacy name as alias for zone-mapping table compatibility
        function filterTable() { applyFilters(); }

        // ── Column visibility ─────────────────────────────────────────────────
        function toggleColumn(colIndex) {
            document.querySelectorAll('[data-col="'+colIndex+'"]').forEach(function(cell) {
                cell.style.display = cell.style.display === 'none' ? '' : 'none';
            });
        }
        function selectAllColumns() {
            document.querySelectorAll('#columnSelector input[type="checkbox"]').forEach(function(cb, i) {
                cb.checked = true;
                document.querySelectorAll('[data-col="'+i+'"]').forEach(function(c) { c.style.display = ''; });
            });
        }
        function deselectAllColumns() {
            document.querySelectorAll('#columnSelector input[type="checkbox"]').forEach(function(cb, i) {
                cb.checked = false;
                document.querySelectorAll('[data-col="'+i+'"]').forEach(function(c) { c.style.display = 'none'; });
            });
        }
        function resetColumns() {
            var defaultVisible = [0, 1, 2, 3, 4, 5, 6, 9, 10, 11, 14, 15];
            document.querySelectorAll('#columnSelector input[type="checkbox"]').forEach(function(cb, i) {
                var show = defaultVisible.indexOf(i) >= 0;
                cb.checked = show;
                document.querySelectorAll('[data-col="'+i+'"]').forEach(function(c) { c.style.display = show ? '' : 'none'; });
            });
        }
        function toggleColumnSelector() {
            document.getElementById('columnSelector').classList.toggle('show');
        }

        // ── Reset all filters ─────────────────────────────────────────────────
        function resetFilters() {
            document.getElementById('searchInput').value = '';
            ['sub-mdd','status-mdd','type-mdd','loc-mdd','rg-mdd'].forEach(function(id) {
                document.querySelectorAll('#'+id+' .mdd-list input[type="checkbox"]').forEach(function(cb) { cb.checked = true; });
                updateMDDLabel(id);
            });
            applyFilters();
        }

        // ── Export visible rows to CSV ────────────────────────────────────────
        function exportTableToCSV() {
            var table = document.getElementById('resourceTable');
            var rows = table.querySelectorAll('tr');
            var csv = [];
            var headers = [];
            rows[0].querySelectorAll('th').forEach(function(cell) {
                if (cell.style.display !== 'none') headers.push('"'+cell.textContent.trim()+'"');
            });
            csv.push(headers.join(','));
            for (var i = 1; i < rows.length; i++) {
                var row = rows[i];
                if (row.style.display !== 'none') {
                    var rowData = [];
                    row.querySelectorAll('td').forEach(function(col) {
                        if (col.style.display !== 'none') rowData.push('"'+col.textContent.trim().replace(/"/g,'""')+'"');
                    });
                    csv.push(rowData.join(','));
                }
            }
            var blob = new Blob([csv.join('\n')], { type: 'text/csv;charset=utf-8;' });
            var link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'ZoneRedundancy_Export_'+new Date().toISOString().slice(0,10)+'.csv';
            link.click();
        }

        // ── Cell expand on click ──────────────────────────────────────────────
        document.querySelectorAll('#resourceTable td').forEach(function(cell) {
            cell.addEventListener('click', function(e) {
                if (e.target.tagName === 'A' || e.target.classList.contains('status-badge')) return;
                this.classList.toggle('expanded');
            });
        });
        document.querySelectorAll('#resourceTable th').forEach(function(header, colIndex) {
            header.addEventListener('dblclick', function() {
                var cells = document.querySelectorAll('#resourceTable td[data-col="'+colIndex+'"]');
                var allExpanded = Array.from(cells).every(function(c) { return c.classList.contains('expanded'); });
                cells.forEach(function(cell) {
                    if (allExpanded) cell.classList.remove('expanded'); else cell.classList.add('expanded');
                });
            });
            header.title = 'Drag edge to resize. Double-click to expand/collapse column content.';
        });

        // ── Zone Mapping table filter ─────────────────────────────────────────
        function filterZoneMappings() {
            var search = (document.getElementById('zmSearch').value || '').toLowerCase();
            var subscription = (document.getElementById('zmSubFilter') || {}).value || '';
            var region = (document.getElementById('zmRegionFilter') || {}).value || '';
            var rows = document.querySelectorAll('#zoneMappingTable tbody tr');
            var visibleCount = 0;
            rows.forEach(function(row) {
                var rowSub = row.getAttribute('data-subscription') || '';
                var rowLoc = row.getAttribute('data-location') || '';
                var text = row.textContent.toLowerCase();
                var ok = (!search || text.includes(search)) && (!subscription || rowSub === subscription) && (!region || rowLoc === region);
                row.style.display = ok ? '' : 'none';
                if (ok) visibleCount++;
            });
            var zmCount = document.getElementById('zmCount');
            if (zmCount) zmCount.textContent = visibleCount + ' mappings';
        }

        // ── Initialise on page load ───────────────────────────────────────────
        buildDropdownOptions();
        applyFilters();
        filterZoneMappings();

        // ── Expandable card detail panels ─────────────────────────────────────
        function toggleCardDetail(id, btn) {
            var panel = document.getElementById(id);
            if (!panel) return;
            var isOpen = panel.classList.contains('open');
            panel.classList.toggle('open');
            btn.textContent = isOpen ? '▾ Learn more' : '▲ Collapse';
            if (btn.textContent === '▾ What\'s included?' || btn.textContent === '▾ Why does this matter?' || btn.textContent === '▾ What resources are included?') {
                btn.textContent = isOpen ? btn.getAttribute('data-open-label') : '▲ Collapse';
            }
        }
        // Set open labels on the buttons
        document.querySelectorAll('.card-expand-btn').forEach(function(btn) {
            btn.setAttribute('data-open-label', btn.textContent);
        });
    </script>
</body>
</html>
"@

$htmlReportFile = Join-Path $OutputPath "Dashboard_ZoneRedundancy.html"
$htmlReport | Out-File -FilePath $htmlReportFile -Encoding UTF8
Write-Log "Generated HTML Dashboard: Dashboard_ZoneRedundancy.html" -Level Success

# Note: HTML dashboard will be opened by Phase 2 after Excel report is generated

# Final summary
Write-Log "=" * 60 -Level Info
Write-Log "Assessment Complete!" -Level Success
Write-Log "Total Subscriptions Processed: $totalSubs" -Level Info
Write-Log "Total Resources Found: $($allResources.Count)" -Level Info
Write-Log "Output Location: $OutputPath" -Level Info
Write-Log "" -Level Info
Write-Log "Files Generated:" -Level Info
Write-Log "  - Dashboard_ZoneRedundancy.html (Interactive HTML Dashboard)" -Level Success
Write-Log "  - MasterReport_AllResources.csv (Combined report)" -Level Info
Write-Log "  - ZoneMappings_AllSubscriptions.csv (Logical to Physical zone mapping)" -Level Info
Write-Log "  - Summary_ZoneRedundancy.csv (By resource type)" -Level Info
Write-Log "  - Summary_BySubscription.csv (By subscription)" -Level Info
Write-Log "  - ResourceTypes/*.csv (Individual resource type files)" -Level Info
Write-Log "  - Assessment_Log.txt (Execution log)" -Level Info
Write-Log "=" * 60 -Level Info

# Return summary object
[PSCustomObject]@{
    TotalSubscriptions = $totalSubs
    TotalResources = $allResources.Count
    OutputPath = $OutputPath
    ZoneRedundantCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'ZoneRedundant' }).Count
    LocallyRedundantCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' }).Count
    NonZonalCount = ($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' }).Count
}

#endregion

# Exit with success code
exit 0
