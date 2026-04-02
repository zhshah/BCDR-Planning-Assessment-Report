<#
.SYNOPSIS
    Complete Azure BCDR Assessment - Runs Phase 1 + Phase 2 Sequentially
    
.DESCRIPTION
    This wrapper script runs both assessment phases in sequence:
    - Phase 1: Zone Redundancy Assessment (Phase1-CollectResources.ps1)
    - Phase 2: SA BCDR Recommendations (Phase2-AddRecommendations.ps1)
    
    All outputs are placed in a single timestamped folder.
    
.PARAMETER TenantId
    Optional. Specify a tenant ID to scope the assessment.

.PARAMETER SubscriptionIds
    Optional. Array of specific subscription IDs to assess.

.PARAMETER CustomerName
    Required. Customer name for the executive summary in the Excel report.

.EXAMPLE
    .\Run-CompleteBCDRAssessment.ps1 -CustomerName "Contoso Corporation"
    
.EXAMPLE
    .\Run-CompleteBCDRAssessment.ps1 -CustomerName "Fabrikam Inc" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    .\Run-CompleteBCDRAssessment.ps1 -CustomerName "Qatar National Library" -SubscriptionIds @("sub-id-1", "sub-id-2")

.NOTES
    Author  : Microsoft Solution Engineering Team
    Version : 3.0 Combined Wrapper
    Date    : 2026-04-01
    
    Output Files Generated:
    PHASE 1:
      - Dashboard_ZoneRedundancy.html (Interactive HTML)
      - MasterReport_AllResources.csv
      - Summary_ZoneRedundancy.csv
      - Summary_BySubscription.csv
      - ZoneMappings_AllSubscriptions.csv
      - ResourceTypes/*.csv
      - MasterReport_AllResources.xlsx
      - Summary_ZoneRedundancy.xlsx
      - Assessment_Log.txt
    PHASE 2:
      - [CustomerName] - BCDR Assessment Report.xlsx (Final comprehensive BCDR report)
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionIds,
    
    [Parameter(Mandatory = $true)]
    [string]$CustomerName
)

$ErrorActionPreference = "Stop"

Clear-Host
Write-Host "`n" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  AZURE BCDR ASSESSMENT - COMPLETE AUTOMATED WORKFLOW" -ForegroundColor Yellow
Write-Host "  Phase 1: Zone Redundancy Assessment → Phase 2: SA BCDR Recommendations" -ForegroundColor Yellow
Write-Host "  Single execution - End to end automation" -ForegroundColor Yellow
Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "`n"
Write-Host "  This script will automatically:" -ForegroundColor White
Write-Host "  1. Run Phase 1 assessment (collect all Azure resources, analyze zones)" -ForegroundColor Gray
Write-Host "  2. Note Phase 1 completion" -ForegroundColor Gray
Write-Host "  3. Automatically start Phase 2 (enrich with SA DR recommendations)" -ForegroundColor Gray
Write-Host "  4. Generate all outputs in a single timestamped folder" -ForegroundColor Gray
Write-Host "`n"
Write-Host "  Assessment for: " -NoNewline -ForegroundColor White
Write-Host "$CustomerName" -ForegroundColor Cyan
Write-Host "`n"

# Script paths
$scriptPath = $PSScriptRoot
$phase1Script = Join-Path $scriptPath "Phase1-CollectResources.ps1"
$phase2Script = Join-Path $scriptPath "Phase2-AddRecommendations.ps1"

# Verify both scripts exist
if (-not (Test-Path $phase1Script)) {
    Write-Host "ERROR: Phase 1 script not found at: $phase1Script" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $phase2Script)) {
    Write-Host "ERROR: Phase 2 script not found at: $phase2Script" -ForegroundColor Red
    exit 1
}

# Create output directory with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = Join-Path $scriptPath "CompleteBCDRAssessment_$timestamp"
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

Write-Host "Output directory: $outputPath" -ForegroundColor Green
Write-Host "`n"

################################################################################
# PHASE 1: ZONE REDUNDANCY ASSESSMENT
################################################################################

Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "  PHASE 1: Zone Redundancy Assessment" -ForegroundColor Yellow
Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "`n"

try {
    # Build parameters for Phase 1
    $phase1Params = @{
        OutputPath = $outputPath
    }
    
    if ($TenantId) {
        $phase1Params.Add('TenantId', $TenantId)
    }
    
    if ($SubscriptionIds) {
        $phase1Params.Add('SubscriptionIds', $SubscriptionIds)
    }
    
    # Execute Phase 1
    Write-Host "▶ Starting Phase 1: Zone Redundancy Assessment..." -ForegroundColor Cyan
    Write-Host "  Collecting Azure resources, analyzing zone redundancy, generating reports..." -ForegroundColor Gray
    Write-Host "`n"
    
    & $phase1Script @phase1Params
    
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null) {
        throw "Phase 1 script failed with exit code $LASTEXITCODE"
    }
    
    Write-Host "`n"
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  ✓✓✓ PHASE 1 COMPLETE ✓✓✓" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Zone assessment data collected. Proceeding to Phase 2..." -ForegroundColor White
    Write-Host "`n"
    
    Start-Sleep -Seconds 2  # Brief pause for user visibility
}
catch {
    Write-Host "`n"
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ✗✗✗ PHASE 1 FAILED ✗✗✗" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ERROR: $_" -ForegroundColor Red
    Write-Host "  Assessment halted. Phase 2 will not run." -ForegroundColor Yellow
    Write-Host "`n"
    exit 1
}

################################################################################
# PHASE 2: SA BCDR RECOMMENDATIONS
################################################################################

Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "  PHASE 2: SA BCDR Recommendations" -ForegroundColor Yellow
Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "`n"

try {
    # Path to the master report CSV generated by Phase 1
    $masterCSV = Join-Path $outputPath "MasterReport_AllResources.csv"
    
    if (-not (Test-Path $masterCSV)) {
        throw "Phase 1 output not found: $masterCSV - Phase 2 cannot continue"
    }
    
    # Build parameters for Phase 2
    $phase2Params = @{
        CsvFilePath = $masterCSV
        OutputPath = $outputPath
        CustomerName = $CustomerName
    }
    
    # Execute Phase 2
    Write-Host "▶ Starting Phase 2: SA BCDR Recommendations..." -ForegroundColor Cyan
    Write-Host "  Enriching data with Senior Architect DR guidance, RPO/RTO, action steps..." -ForegroundColor Gray
    Write-Host "`n"
    
    & $phase2Script @phase2Params
    
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null) {
        throw "Phase 2 script failed with exit code $LASTEXITCODE"
    }
    
    Write-Host "`n"
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  ✓✓✓ PHASE 2 COMPLETE ✓✓✓" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  BCDR recommendations applied. Comprehensive assessment report generated." -ForegroundColor White
    Write-Host "`n"
}
catch {
    Write-Host "`n"
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ✗✗✗ PHASE 2 FAILED ✗✗✗" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ERROR: $_" -ForegroundColor Red
    Write-Host "  Phase 1 completed successfully, but Phase 2 encountered an error." -ForegroundColor Yellow
    Write-Host "  You can manually run Phase2-AddRecommendations.ps1 pointing to:" -ForegroundColor Yellow
    Write-Host "  $masterCSV" -ForegroundColor Cyan
    Write-Host "`n"
    exit 1
}

################################################################################
# FINAL SUMMARY
################################################################################

Write-Host "`n" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  COMPLETE BCDR ASSESSMENT FINISHED SUCCESSFULLY!" -ForegroundColor Yellow
Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "`n"

Write-Host "Customer: " -NoNewline -ForegroundColor White
Write-Host "$CustomerName" -ForegroundColor Cyan
Write-Host "Output Location: " -NoNewline -ForegroundColor White
Write-Host "$outputPath" -ForegroundColor Green
Write-Host "`n"

Write-Host "Files Generated:" -ForegroundColor White
Write-Host "  PHASE 1 OUTPUT:" -ForegroundColor Cyan
Write-Host "    ✓ Dashboard_ZoneRedundancy.html (Interactive HTML Dashboard)" -ForegroundColor White
Write-Host "    ✓ MasterReport_AllResources.csv" -ForegroundColor White
Write-Host "    ✓ MasterReport_AllResources.xlsx" -ForegroundColor White
Write-Host "    ✓ Summary_ZoneRedundancy.csv" -ForegroundColor White
Write-Host "    ✓ Summary_ZoneRedundancy.xlsx" -ForegroundColor White
Write-Host "    ✓ Summary_BySubscription.csv" -ForegroundColor White
Write-Host "    ✓ ZoneMappings_AllSubscriptions.csv" -ForegroundColor White
Write-Host "    ✓ ResourceTypes/*.csv (Individual resource type files)" -ForegroundColor White
Write-Host "`n"

Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "`n"

# Note: HTML Dashboard will be opened by Phase 2 after updating with Excel report link

# Display final report location
$saReport = Get-ChildItem -Path $outputPath -Filter "*BCDR Assessment Report.xlsx" | Select-Object -First 1 -ExpandProperty FullName
if ($saReport -and (Test-Path $saReport)) {
    Write-Host "🎯 FINAL BCDR ASSESSMENT REPORT GENERATED:" -ForegroundColor Green
    Write-Host ""
    Write-Host "   $saReport" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   This comprehensive Excel workbook contains:" -ForegroundColor White
    Write-Host "   • 12 worksheets with detailed BCDR recommendations" -ForegroundColor Gray
    Write-Host "   • Qatar-specific compliance guidance (PDPPL, NIA/NCSA)" -ForegroundColor Gray
    Write-Host "   • Timeline, Testing Plans, Dependencies Matrix" -ForegroundColor Gray
    Write-Host "   • BCDR Strategy Reference (design patterns & rationale)" -ForegroundColor Gray
    Write-Host ""
}

Write-Host "`nAssessment complete! Review the outputs in: $outputPath" -ForegroundColor Green
