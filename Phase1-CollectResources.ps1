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
    
    return @{
        ZoneStatus = $zoneStatus
        ZoneDetail = $zoneDetail
        LogicalZones = ($zones -join ', ')
        PhysicalZones = ($physicalZones -join ', ')
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
        default { 'Unknown' }
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
        default { 'Unknown' }
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
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; color: #333; }
        .header { background: linear-gradient(135deg, #0078d4, #005a9e); color: white; padding: 30px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .container { max-width: 100%; margin: 0 auto; padding: 20px; }
        .summary-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin: 20px 0; }
        .card { background: white; border-radius: 10px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); transition: transform 0.2s; text-align: center; }
        .card:hover { transform: translateY(-5px); }
        .card h3 { font-size: 0.85em; color: #666; text-transform: uppercase; margin-bottom: 8px; }
        .card .value { font-size: 2.2em; font-weight: bold; }
        .card.zone-redundant .value { color: #107c10; }
        .card.zonal .value { color: #0078d4; }
        .card.non-zonal .value { color: #d83b01; }
        .card.locally-redundant .value { color: #ffb900; }
        .card.total .value { color: #0078d4; }
        .charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 20px; margin: 20px 0; }
        .chart-container { background: white; border-radius: 10px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .chart-container h2 { font-size: 1.1em; margin-bottom: 15px; color: #333; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        .table-container { background: white; border-radius: 10px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin: 20px 0; }
        .table-container h2 { font-size: 1.2em; margin-bottom: 15px; color: #333; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        .table-wrapper { overflow-x: auto; max-height: 600px; overflow-y: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 0.85em; min-width: 1200px; table-layout: auto; }
        th { background: #0078d4; color: white; padding: 12px 8px; text-align: left; font-weight: 600; position: sticky; top: 0; z-index: 10; white-space: nowrap; cursor: col-resize; border-right: 2px solid #005a9e; min-width: 80px; resize: horizontal; overflow: auto; }
        th:last-child { border-right: none; }
        td { padding: 8px; border-bottom: 1px solid #eee; max-width: 300px; overflow: hidden; text-overflow: ellipsis; word-wrap: break-word; }
        td.expanded { white-space: normal; max-width: none; overflow: visible; }
        td:hover { background: #e8f4fc; cursor: pointer; }
        tr:hover { background: #f0f7ff; }
        tr:nth-child(even) { background: #fafafa; }
        tr:nth-child(even):hover { background: #f0f7ff; }
        .status-badge { padding: 3px 10px; border-radius: 15px; font-size: 0.8em; font-weight: 500; display: inline-block; white-space: nowrap; }
        .status-ZoneRedundant { background: #dff6dd; color: #107c10; }
        .status-Zonal { background: #deecf9; color: #0078d4; }
        .status-NonZonal { background: #fed9cc; color: #d83b01; }
        .status-LocallyRedundant { background: #fff4ce; color: #797673; }
        .status-RedundantByDefault { background: #e8f4e8; color: #107c10; }
        .status-GeoRedundant, .status-GeoZoneRedundant { background: #e0f4f7; color: #004c54; }
        .status-Unknown, .status-InheritedFromParent { background: #f3f2f1; color: #605e5c; }
        .risk-high { background: #fde7e9 !important; }
        .risk-medium { background: #fff4ce !important; }
        .filter-bar { background: white; border-radius: 10px; padding: 15px 20px; margin: 20px 0; box-shadow: 0 2px 10px rgba(0,0,0,0.1); display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
        .filter-bar label { font-weight: 600; font-size: 0.9em; color: #333; }
        .filter-bar input, .filter-bar select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 5px; font-size: 0.9em; }
        .filter-bar input[type="text"] { width: 250px; }
        .filter-bar select { min-width: 150px; }
        .filter-group { display: flex; align-items: center; gap: 5px; }
        .btn { padding: 8px 16px; border: none; border-radius: 5px; cursor: pointer; font-size: 0.9em; transition: background 0.2s; }
        .btn-primary { background: #0078d4; color: white; }
        .btn-primary:hover { background: #005a9e; }
        .btn-secondary { background: #f3f2f1; color: #333; border: 1px solid #ddd; }
        .btn-secondary:hover { background: #e1e1e1; }
        .column-selector { background: white; border: 1px solid #ddd; border-radius: 8px; padding: 15px; margin: 10px 0; display: none; max-height: 400px; overflow-y: auto; }
        .column-selector.show { display: block; }
        .column-selector h4 { margin-bottom: 10px; color: #333; position: sticky; top: 0; background: white; padding: 5px 0; }
        .column-selector label { display: inline-flex; align-items: center; margin: 5px 15px 5px 0; font-size: 0.9em; cursor: pointer; min-width: 180px; }
        .column-selector input[type="checkbox"] { margin-right: 5px; }
        .column-selector-actions { margin-bottom: 10px; padding-bottom: 10px; border-bottom: 1px solid #ddd; }
        .column-selector-actions button { margin-right: 10px; padding: 5px 10px; font-size: 0.85em; }
        .recommendations { background: #fff8e5; border-left: 4px solid #ffb900; padding: 15px 20px; margin: 15px 0; border-radius: 0 10px 10px 0; }
        .recommendations h4 { color: #333; margin-bottom: 8px; }
        .recommendations ul { margin-left: 20px; }
        .recommendations li { margin: 5px 0; color: #555; font-size: 0.95em; }
        .footer { text-align: center; padding: 20px; color: #666; font-size: 0.9em; border-top: 1px solid #ddd; margin-top: 20px; }
        .export-info { background: #e8f4e8; border-left: 4px solid #107c10; padding: 10px 15px; margin: 10px 0; border-radius: 0 8px 8px 0; font-size: 0.9em; }
        .resource-count { background: #0078d4; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.85em; margin-left: 10px; }
        @media (max-width: 768px) { 
            .charts-row { grid-template-columns: 1fr; } 
            .filter-bar { flex-direction: column; align-items: stretch; }
            .filter-bar input[type="text"] { width: 100%; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Azure Resilience & Disaster Recovery Assessment</h1>
        <p>Generated: $(Get-Date -Format "MMMM dd, yyyy HH:mm:ss") | Region Focus: Qatar Central | Subscriptions: $totalSubs | Total Resources: $($allResources.Count)</p>
    </div>
    
    <div class="container">
        <!-- Export Info -->
        <div class="export-info">
            <strong>Output Location:</strong> $OutputPath<br>
            <strong>CSV Files Exported:</strong> MasterReport_AllResources.csv, Summary_ZoneRedundancy.csv, Summary_BySubscription.csv, ZoneMappings_AllSubscriptions.csv, and individual files per resource type in /ResourceTypes folder.<br>
            <strong>Excel Reports:</strong> MasterReport_AllResources.xlsx, Summary_ZoneRedundancy.xlsx (with color-coded formatting)
        </div>

        <!-- Summary Cards -->
        <div class="summary-cards">
            <div class="card total">
                <h3>Total Resources</h3>
                <div class="value">$($allResources.Count)</div>
            </div>
            <div class="card zone-redundant">
                <h3>Zone Redundant</h3>
                <div class="value">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'ZoneRedundant' }).Count)</div>
            </div>
            <div class="card zonal">
                <h3>Zonal</h3>
                <div class="value">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'Zonal' }).Count)</div>
            </div>
            <div class="card non-zonal">
                <h3>Non-Zonal</h3>
                <div class="value">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'NonZonal' }).Count)</div>
            </div>
            <div class="card locally-redundant">
                <h3>Locally Redundant</h3>
                <div class="value">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'LocallyRedundant' }).Count)</div>
            </div>
            <div class="card" style="background: #e8f4e8;">
                <h3>Redundant By Default</h3>
                <div class="value" style="color: #107c10;">$(($allResources | Where-Object { $_.ZoneRedundancyStatus -eq 'RedundantByDefault' }).Count)</div>
            </div>
        </div>

        <!-- Cross-Region Availability Summary -->
        <div class="summary-cards" style="margin-top: 20px;">
            <div class="card" style="background: linear-gradient(135deg, #004c54, #006d77);">
                <h3 style="color: white;">Cross-Region</h3>
                <div class="value" style="color: white;">DR Analysis</div>
            </div>
            <div class="card" style="background: #e0f4f7;">
                <h3>Geo-Redundant</h3>
                <div class="value" style="color: #004c54;">$(($allResources | Where-Object { $_.GeoRedundant -eq $true }).Count)</div>
            </div>
            <div class="card" style="background: #d4e6ff;">
                <h3>DR Ready</h3>
                <div class="value" style="color: #0078d4;">$(($allResources | Where-Object { $_.DRReady -eq $true }).Count)</div>
            </div>
            <div class="card" style="background: #fff4ce;">
                <h3>Single Region Only</h3>
                <div class="value" style="color: #ca5010;">$(($allResources | Where-Object { $_.CrossRegionStatus -eq 'Single-Region' }).Count)</div>
            </div>
            <div class="card" style="background: #dff6dd;">
                <h3>Global Services</h3>
                <div class="value" style="color: #107c10;">$(($allResources | Where-Object { $_.CrossRegionStatus -eq 'Global' }).Count)</div>
            </div>
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
                <div style="background: white; padding: 12px; border-radius: 8px; border-left: 4px solid #0078d4;">
                    <h5 style="margin: 0 0 8px 0; color: #0078d4;">📊 Cross-Region Replication Status</h5>
                    <ul style="margin: 0; padding-left: 20px; font-size: 0.85em;">
                        <li><strong>$(($allResources | Where-Object { $_.GeoRedundant -eq $true }).Count)</strong> resources with geo-redundancy configured</li>
                        <li><strong>$(($allResources | Where-Object { $_.DRReady -eq $true }).Count)</strong> resources DR-ready (multi-region/global)</li>
                        <li><strong>$(($allResources | Where-Object { $_.CrossRegionStatus -eq 'Single-Region' }).Count)</strong> resources in single region only</li>
                    </ul>
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

        <!-- Charts Row -->
        <div class="charts-row">
            <div class="chart-container">
                <h2>Zone Redundancy Distribution</h2>
                <canvas id="pieChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>By Subscription</h2>
                <canvas id="barChart"></canvas>
            </div>
        </div>

        <div class="charts-row">
            <div class="chart-container">
                <h2>Top Resource Types by Count</h2>
                <canvas id="resourceTypeChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>Risk Assessment</h2>
                <canvas id="riskChart"></canvas>
            </div>
        </div>

        <div class="charts-row">
            <div class="chart-container">
                <h2>Cross-Region Status</h2>
                <canvas id="crossRegionChart"></canvas>
            </div>
            <div class="chart-container">
                <h2>Regional Distribution</h2>
                <canvas id="regionChart"></canvas>
            </div>
        </div>

        <!-- Filter Bar -->
        <div class="filter-bar">
            <div class="filter-group">
                <label>Search:</label>
                <input type="text" id="searchInput" placeholder="Search resources..." onkeyup="filterTable()">
            </div>
            <div class="filter-group">
                <label>Subscription:</label>
                <select id="subscriptionFilter" onchange="filterTable()">
                    <option value="">All Subscriptions</option>
                    $($uniqueSubscriptions | ForEach-Object { "<option value=`"$_`">$_</option>" })
                </select>
            </div>
            <div class="filter-group">
                <label>Status:</label>
                <select id="statusFilter" onchange="filterTable()">
                    <option value="">All Statuses</option>
                    <option value="ZoneRedundant">Zone Redundant</option>
                    <option value="Zonal">Zonal</option>
                    <option value="NonZonal">Non-Zonal</option>
                    <option value="LocallyRedundant">Locally Redundant</option>
                    <option value="RedundantByDefault">Redundant By Default</option>
                    <option value="GeoRedundant">Geo Redundant</option>
                    <option value="Unknown">Unknown</option>
                </select>
            </div>
            <div class="filter-group">
                <label>Type:</label>
                <select id="typeFilter" onchange="filterTable()">
                    <option value="">All Resource Types</option>
                    $($resourcesByType.Keys | Sort-Object | ForEach-Object { "<option value=`"$_`">$_</option>" })
                </select>
            </div>
            <div class="filter-group">
                <label>Location:</label>
                <select id="locationFilter" onchange="filterTable()">
                    <option value="">All Locations</option>
                    $($allResources | Select-Object -ExpandProperty location -Unique | Sort-Object | ForEach-Object { "<option value=`"$_`">$_</option>" })
                </select>
            </div>
            <div class="filter-group">
                <label>Resource Group:</label>
                <select id="resourceGroupFilter" onchange="filterTable()">
                    <option value="">All Resource Groups</option>
                    $($allResources | Select-Object -ExpandProperty resourceGroup -Unique | Sort-Object | ForEach-Object { "<option value=`"$_`">$_</option>" })
                </select>
            </div>
            <div class="filter-group">
                <label>Cross-Region:</label>
                <select id="crossRegionFilter" onchange="filterTable()">
                    <option value="">All</option>
                    <option value="Geo-Redundant">Geo-Redundant</option>
                    <option value="Multi-Region">Multi-Region</option>
                    <option value="Global">Global</option>
                    <option value="Single-Region">Single-Region</option>
                    <option value="Unknown">Unknown</option>
                </select>
            </div>
            <button class="btn btn-secondary" onclick="toggleColumnSelector()">Columns</button>
            <button class="btn btn-secondary" onclick="resetFilters()">Reset</button>
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

# Cross-region data
$crossRegionGeoRedundantCount = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Geo-Redundant' }).Count
$crossRegionDrReadyCount = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Multi-Region' -or $_.CrossRegionStatus -eq 'Global' }).Count
$crossRegionSingleCount = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Single-Region' }).Count
$crossRegionGlobalCount = ($allResources | Where-Object { $_.CrossRegionStatus -eq 'Global' }).Count

# Regional distribution data (top 10 regions)
$topRegions = $allResources | Group-Object location | Sort-Object Count -Descending | Select-Object -First 10
$regionLabels = ($topRegions | ForEach-Object { "`"$($_.Name)`"" }) -join ','
$regionCounts = ($topRegions | ForEach-Object { $_.Count }) -join ','

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
        <p>Azure Resilience & DR Assessment Tool | Developed by Zahir Hussain Shah - Dashboard Version 2.0</p>
    </div>

    <script>
        // Percentage label plugin for doughnut charts
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
                    
                    // Add shadow for better visibility
                    ctx.shadowColor = 'rgba(0,0,0,0.5)';
                    ctx.shadowBlur = 3;
                    ctx.fillText(percentage + '%', x, y);
                    ctx.restore();
                });
            }
        };

        // Pie Chart - Zone Distribution
        new Chart(document.getElementById('pieChart'), {
            type: 'doughnut',
            data: {
                labels: ['Zone Redundant', 'Zonal', 'Non-Zonal', 'Locally Redundant', 'Redundant By Default', 'Geo Redundant', 'Other'],
                datasets: [{
                    data: [$zoneRedundantCount, $zonalCount, $nonZonalCount, $locallyRedundantCount, $redundantByDefaultCount, $geoRedundantCount, $unknownCount],
                    backgroundColor: ['#107c10', '#0078d4', '#d83b01', '#ffb900', '#00bcf2', '#004c54', '#605e5c']
                }]
            },
            options: { 
                responsive: true, 
                plugins: { 
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((context.raw / total) * 100).toFixed(1);
                                return context.label + ': ' + context.raw + ' (' + percentage + '%)';
                            }
                        }
                    }
                } 
            },
            plugins: [percentagePlugin]
        });

        // Bar Chart - By Subscription
        new Chart(document.getElementById('barChart'), {
            type: 'bar',
            data: {
                labels: [$subLabels],
                datasets: [
                    { label: 'Zone Redundant', data: [$subZR], backgroundColor: '#107c10' },
                    { label: 'Non-Zonal', data: [$subNonZonal], backgroundColor: '#d83b01' },
                    { label: 'Locally Redundant', data: [$subLR], backgroundColor: '#ffb900' }
                ]
            },
            options: { responsive: true, scales: { x: { stacked: true }, y: { stacked: true } } }
        });

        // Resource Type Chart
        new Chart(document.getElementById('resourceTypeChart'), {
            type: 'bar',
            data: {
                labels: [$typeLabels],
                datasets: [{ label: 'Count', data: [$typeCounts], backgroundColor: '#0078d4' }]
            },
            options: { responsive: true, indexAxis: 'y' }
        });

        // Risk Chart
        new Chart(document.getElementById('riskChart'), {
            type: 'doughnut',
            data: {
                labels: ['High Risk (Critical Non-Zonal)', 'Medium Risk (Other Non-Zonal)', 'Low Risk (Protected)'],
                datasets: [{
                    data: [$highRiskCount, $mediumRiskCount, $lowRiskCount],
                    backgroundColor: ['#d83b01', '#ffb900', '#107c10']
                }]
            },
            options: { 
                responsive: true, 
                plugins: { 
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((context.raw / total) * 100).toFixed(1);
                                return context.label + ': ' + context.raw + ' (' + percentage + '%)';
                            }
                        }
                    }
                } 
            },
            plugins: [percentagePlugin]
        });

        // Cross-Region Chart
        new Chart(document.getElementById('crossRegionChart'), {
            type: 'doughnut',
            data: {
                labels: ['Geo-Redundant', 'Multi-Region/Global', 'Single Region'],
                datasets: [{
                    data: [$crossRegionGeoRedundantCount, $crossRegionDrReadyCount, $crossRegionSingleCount],
                    backgroundColor: ['#107c10', '#0078d4', '#ffb900']
                }]
            },
            options: { 
                responsive: true, 
                plugins: { 
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((context.raw / total) * 100).toFixed(1);
                                return context.label + ': ' + context.raw + ' (' + percentage + '%)';
                            }
                        }
                    }
                } 
            },
            plugins: [percentagePlugin]
        });

        // Region Chart
        new Chart(document.getElementById('regionChart'), {
            type: 'bar',
            data: {
                labels: [$regionLabels],
                datasets: [{ label: 'Resources', data: [$regionCounts], backgroundColor: '#8764b8' }]
            },
            options: { 
                responsive: true, 
                indexAxis: 'y',
                plugins: { legend: { display: false } }
            }
        });

        // Filter function with all filters including subscription
        function filterTable() {
            const search = document.getElementById('searchInput').value.toLowerCase();
            const subscription = document.getElementById('subscriptionFilter').value;
            const status = document.getElementById('statusFilter').value;
            const type = document.getElementById('typeFilter').value;
            const location = document.getElementById('locationFilter').value;
            const resourceGroup = document.getElementById('resourceGroupFilter').value;
            const crossRegion = document.getElementById('crossRegionFilter').value;
            const rows = document.querySelectorAll('#resourceTable tbody tr');
            
            let visibleCount = 0;
            
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                const rowSubscription = row.getAttribute('data-subscription');
                const rowStatus = row.getAttribute('data-status');
                const rowType = row.getAttribute('data-type');
                const rowLocation = row.getAttribute('data-location');
                const rowResourceGroup = row.getAttribute('data-resourcegroup');
                const rowCrossRegion = row.getAttribute('data-crossregion');
                
                const matchSearch = !search || text.includes(search);
                const matchSubscription = !subscription || rowSubscription === subscription;
                const matchStatus = !status || rowStatus === status;
                const matchType = !type || rowType === type;
                const matchLocation = !location || rowLocation === location;
                const matchResourceGroup = !resourceGroup || rowResourceGroup === resourceGroup;
                const matchCrossRegion = !crossRegion || rowCrossRegion === crossRegion;
                
                const isVisible = matchSearch && matchSubscription && matchStatus && matchType && matchLocation && matchResourceGroup && matchCrossRegion;
                row.style.display = isVisible ? '' : 'none';
                if (isVisible) visibleCount++;
            });
            
            document.getElementById('visibleCount').textContent = visibleCount;
        }

        // Toggle column visibility
        function toggleColumn(colIndex) {
            const cells = document.querySelectorAll('[data-col="' + colIndex + '"]');
            cells.forEach(cell => {
                cell.style.display = cell.style.display === 'none' ? '' : 'none';
            });
        }

        // Select all columns
        function selectAllColumns() {
            const checkboxes = document.querySelectorAll('#columnSelector input[type="checkbox"]');
            checkboxes.forEach((checkbox, index) => {
                checkbox.checked = true;
                const cells = document.querySelectorAll('[data-col="' + index + '"]');
                cells.forEach(cell => cell.style.display = '');
            });
        }

        // Deselect all columns
        function deselectAllColumns() {
            const checkboxes = document.querySelectorAll('#columnSelector input[type="checkbox"]');
            checkboxes.forEach((checkbox, index) => {
                checkbox.checked = false;
                const cells = document.querySelectorAll('[data-col="' + index + '"]');
                cells.forEach(cell => cell.style.display = 'none');
            });
        }

        // Reset columns to default visibility
        function resetColumns() {
            const checkboxes = document.querySelectorAll('#columnSelector input[type="checkbox"]');
            // Default visible columns: Subscription(0), Name(1), Type(2), Location(3), ResourceGroup(4),
            // ZoneStatus(5), ZoneDetail(6), CrossRegionStatus(9), CrossRegionDetail(10), PairedRegion(11), SkuName(14), SkuTier(15)
            const defaultVisible = [0, 1, 2, 3, 4, 5, 6, 9, 10, 11, 14, 15];
            checkboxes.forEach((checkbox, index) => {
                const shouldBeVisible = defaultVisible.includes(index);
                checkbox.checked = shouldBeVisible;
                const cells = document.querySelectorAll('[data-col="' + index + '"]');
                cells.forEach(cell => cell.style.display = shouldBeVisible ? '' : 'none');
            });
        }

        // Toggle column selector panel
        function toggleColumnSelector() {
            const selector = document.getElementById('columnSelector');
            selector.classList.toggle('show');
        }

        // Reset all filters
        function resetFilters() {
            document.getElementById('searchInput').value = '';
            document.getElementById('subscriptionFilter').value = '';
            document.getElementById('statusFilter').value = '';
            document.getElementById('typeFilter').value = '';
            document.getElementById('locationFilter').value = '';
            document.getElementById('resourceGroupFilter').value = '';
            document.getElementById('crossRegionFilter').value = '';
            filterTable();
        }

        // Export visible table data to CSV
        function exportTableToCSV() {
            const table = document.getElementById('resourceTable');
            const rows = table.querySelectorAll('tr');
            let csv = [];
            
            // Get visible columns
            const headers = [];
            const headerRow = rows[0];
            const headerCells = headerRow.querySelectorAll('th');
            headerCells.forEach((cell, index) => {
                if (cell.style.display !== 'none') {
                    headers.push('"' + cell.textContent.trim() + '"');
                }
            });
            csv.push(headers.join(','));
            
            // Get data rows
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (row.style.display !== 'none') {
                    const cols = row.querySelectorAll('td');
                    const rowData = [];
                    cols.forEach((col, index) => {
                        if (col.style.display !== 'none') {
                            let text = col.textContent.trim().replace(/"/g, '""');
                            rowData.push('"' + text + '"');
                        }
                    });
                    csv.push(rowData.join(','));
                }
            }
            
            // Download
            const csvContent = csv.join('\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'ZoneRedundancy_FilteredExport_' + new Date().toISOString().slice(0,10) + '.csv';
            link.click();
        }

        // Toggle cell expansion on click for long text
        document.querySelectorAll('#resourceTable td').forEach(cell => {
            cell.addEventListener('click', function(e) {
                // Don't interfere with links or badges
                if (e.target.tagName === 'A' || e.target.classList.contains('status-badge')) return;
                this.classList.toggle('expanded');
            });
        });

        // Add double-click to expand all cells in a column
        document.querySelectorAll('#resourceTable th').forEach((header, colIndex) => {
            header.addEventListener('dblclick', function() {
                const cells = document.querySelectorAll('#resourceTable td[data-col=\"' + colIndex + '\"]');
                const allExpanded = Array.from(cells).every(c => c.classList.contains('expanded'));
                cells.forEach(cell => {
                    if (allExpanded) {
                        cell.classList.remove('expanded');
                    } else {
                        cell.classList.add('expanded');
                    }
                });
            });
            header.title = 'Drag edge to resize. Double-click to expand/collapse column content.';
        });

        // Filter zone mappings table
        function filterZoneMappings() {
            const search = document.getElementById('zmSearch').value.toLowerCase();
            const subscription = document.getElementById('zmSubFilter').value;
            const region = document.getElementById('zmRegionFilter').value;
            const rows = document.querySelectorAll('#zoneMappingTable tbody tr');
            
            let visibleCount = 0;
            rows.forEach(row => {
                const rowSub = row.getAttribute('data-subscription');
                const rowLocation = row.getAttribute('data-location');
                const text = row.textContent.toLowerCase();
                
                const matchesSearch = !search || text.includes(search);
                const matchesSub = !subscription || rowSub === subscription;
                const matchesRegion = !region || rowLocation === region;
                
                if (matchesSearch && matchesSub && matchesRegion) {
                    row.style.display = '';
                    visibleCount++;
                } else {
                    row.style.display = 'none';
                }
            });
            
            document.getElementById('zmCount').textContent = visibleCount + ' mappings';
        }

        // Initialize - update visible count
        filterTable();
        filterZoneMappings();
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
