# Azure BCDR Assessment Automation Framework

![Azure](https://img.shields.io/badge/Azure-0078D4?style=flat&logo=microsoft-azure&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=flat&logo=powershell&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-1.0-blue.svg)

**Comprehensive Business Continuity and Disaster Recovery (BCDR) Assessment Tool for Azure Environments**

Developed by **Zahir Hussain Shah**, Sr. Solution Engineer — Cloud & AI Infrastructure, Microsoft Qatar

---

## 🎯 Overview

This automated framework provides **end-to-end BCDR assessment** for Azure subscriptions with specialized capabilities for:
- **Qatar Central region** constraints (no paired region)
- **NIA/NCSA compliance** validation for Qatar government and regulated sectors
- **Azure Backup Region of Choice (RoC)** preview support
- **Cross-region disaster recovery** strategy design
- **Smart workload tier classification** for cost-optimized DR planning

**⚠️ CRITICAL GUIDANCE:** This framework **emphasizes cross-region disaster recovery** over reliance on availability zones alone. While availability zones provide in-region high availability, they do NOT protect against region-wide failures, natural disasters, or geo-political events. **Cross-region DR is mandatory** for business-critical workloads, especially in regions with constraints like Qatar Central (no paired region) or where zone redundancy is restricted.

---

## 📋 Table of Contents

- [Key Features](#-key-features)
- [Architecture](#-architecture)
- [Prerequisites](#-prerequisites)
- [Quick Start](#-quick-start)
- [What Each Phase Does](#-what-each-phase-does)
- [Output Reports](#-output-reports)
- [Assessment Methodology](#-assessment-methodology)
- [Qatar-Specific Compliance](#-qatar-specific-compliance)
- [BCDR Strategies Recommended](#-bcdr-strategies-recommended)
- [Supported Azure Services](#-supported-azure-services)
- [Limitations & Disclaimers](#%EF%B8%8F-limitations--disclaimers)
- [Version History](#-version-history)

---

## 🚀 Key Features

### ✅ **2-Phase Automated Assessment Workflow**
- **Single command execution** orchestrates both phases automatically
- **Phase 1:** Resource collection, zone redundancy analysis, cross-region readiness
- **Phase 2:** BCDR recommendation enrichment with SA guidance, RPO/RTO targets, cost analysis

### ✅ **Smart Workload Tier Classification** 🆕
Automatically detects environment type to optimize DR costs:
- **Production** → Active-Passive (Warm Standby) or Active-Active strategies
- **Non-Production** → Backup & Restore
- **Dev/Test** → Backup & Restore (Cost-Optimized) with 24h RPO acceptable
- **Sandbox** → No DR Required (Recreate from IaC) — Zero DR cost

**Detection Method:** 4-priority hierarchy
1. Subscription name analysis (High confidence)
2. Resource tags: Environment, Tier, Stage, Workload (Medium confidence)
3. Resource group naming patterns (Low confidence)
4. Resource name patterns (Lowest confidence)

**Cost Impact:** Potential **96% DR cost reduction** for non-production workloads by avoiding expensive ASR, SQL Failover Groups, or AKS warm standby for dev/test environments.

### ✅ **Comprehensive 12-Sheet Excel Report**
1. **Introduction** — Assessment methodology, tier detection logic, priority assignment criteria
2. **Executive Summary** — Key findings, Qatar constraints, top recommendations
3. **SA_Recommendations** — Full resource inventory (339+ resources) with per-service BCDR guidance
4. **QuickWins** — Low-effort, high-value actions (e.g., enable ACR geo-replication, configure object replication)
5. **P1_Critical_Actions** — Suggested highest-priority resources (customer validation required)
6. **Summary_ByResourceType** — Aggregated DR gap analysis by Azure service type
7. **Summary_BySubscription** — Subscription-level DR readiness overview
8. **Timeline_ActionPlan** — Phased implementation roadmap (Foundation → Quick Wins → P1 → P2/P3)
9. **DR_Testing_Plan** — Quarterly DR testing template with success criteria and rollback procedures
10. **Dependencies_Matrix** — Service dependency mapping extracted from real environment data
11. **Compliance_Checklist** — Qatar PDPPL & NIA/NCSA compliance tracker with DPO sign-off
12. **BCDR_Strategy_Reference** — Complete DR strategy design patterns guide with Qatar constraints

### ✅ **Interactive HTML Dashboard**
- **Zone redundancy heatmap** by subscription
- **Cross-region DR status** visualization
- **Clickable BCDR report card** with direct Excel report link
- **Subscription-level statistics** and DR gap summary

### ✅ **Priority & Criticality Methodology** 🆕
Transparent, documented classification logic:
- **2-Step Process:** Criticality Assessment → Priority Assignment
- **Input Factors:** Subscription name, resource group, resource name, resource type, zone redundancy status
- **Output:** P1/P2/P3/P4 priority labels with **(Suggested — Confirm with Customer)** disclaimer
- **Full Documentation:** Introduction sheet explains exact keyword patterns, resource type criteria, and decision tree

### ✅ **Cross-Region DR Emphasis**
**This framework prioritizes cross-region disaster recovery** over availability zones because:
- **Availability zones protect against datacenter failures** (single AZ outage)
- **Availability zones do NOT protect against:**
  - Region-wide failures (entire region outage)
  - Natural disasters (earthquakes, floods affecting entire region)
  - Geo-political events or compliance changes
  - Catastrophic infrastructure failures
- **Qatar Central specific:** Zone redundancy is RESTRICTED (one AZ at full capacity)
- **Best Practice:** Use zones for in-region HA + cross-region for DR (defense in depth)

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│  Start-AzureBCDRAssessment.ps1 (Main Orchestrator)                 │
└────────────────┬────────────────────────────────────────────────────┘
                 │
                 ├─► PHASE 1: Phase1-CollectResources.ps1
                 │   ├─ Connect to Azure subscriptions
                 │   ├─ Query Azure Resource Graph (all resources)
                 │   ├─ Analyze zone redundancy status per resource
                 │   ├─ Map logical → physical availability zones
                 │   ├─ Assess cross-region DR readiness
                 │   ├─ Extract resource dependencies (VNet, Subnet, KeyVault, Storage)
                 │   ├─ Export 54 resource type CSV files
                 │   ├─ Generate master inventory CSV
                 │   └─ Create HTML dashboard with placeholder for Phase 2
                 │
                 └─► PHASE 2: Phase2-AddRecommendations.ps1
                     ├─ Load Phase 1 CSV output
                     ├─ Smart workload tier classification
                     │   ├─ Detect Production/Non-Prod/Dev-Test/Sandbox
                     │   ├─ Adjust BCDR strategy based on tier
                     │   └─ Reduce DR costs for non-production workloads
                     ├─ Apply service-specific DR knowledge base (54 resource types)
                     │   ├─ Map to BCDR strategies (Active-Active, Active-Passive, Backup & Restore, IaC DR)
                     │   ├─ Assign RPO/RTO targets per service
                     │   ├─ Provide implementation effort & cost estimates
                     │   └─ Generate action-required steps
                     ├─ Criticality & Priority classification
                     │   ├─ Critical/High/Medium/Low based on naming + resource type
                     │   └─ P1/P2/P3/P4 priority based on criticality + zone redundancy
                     ├─ Generate 12-sheet comprehensive Excel report
                     ├─ Update HTML dashboard with clickable Excel report card
                     └─ Open HTML dashboard in browser
```

---

## 📦 Prerequisites

### **Required PowerShell Modules**
```powershell
# Azure PowerShell modules
Install-Module -Name Az.Accounts -Scope CurrentUser -Force
Install-Module -Name Az.ResourceGraph -Scope CurrentUser -Force

# Excel export module
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

### **Azure Permissions**
- **Minimum:** `Reader` role on target subscriptions
- **Recommended:** `Reader` + `Tag Contributor` (for tagging recommendations)

### **PowerShell Version**
- PowerShell 7.0 or higher

---

## ⚡ Quick Start

### **Single-Command Execution (Recommended)**

```powershell
.\Start-AzureBCDRAssessment.ps1 -CustomerName "Your Organization" -TenantId "your-tenant-id"
```

This will:
1. ✅ Authenticate to Azure
2. ✅ Run Phase 1 (resource collection & zone analysis)
3. ✅ Automatically run Phase 2 (BCDR enrichment)
4. ✅ Generate comprehensive Excel report (12 sheets)
5. ✅ Update HTML dashboard with clickable report card
6. ✅ Open HTML dashboard in browser

### **Output Location**
All reports generated in timestamped folder:
```
CompleteBCDRAssessment_YYYYMMDD_HHMMSS/
├── [Customer Name] - BCDR Assessment Report.xlsx (12 sheets)
├── Dashboard_ZoneRedundancy.html
├── MasterReport_AllResources.csv
├── Summary_ZoneRedundancy.csv
├── Summary_BySubscription.csv
├── ZoneMappings_AllSubscriptions.csv
├── ResourceTypes/ (54 CSV files)
└── Assessment_Log.txt
```

---

## 📊 What Each Phase Does

### **Phase 1: Resource Collection & Zone Redundancy Analysis**

**Script:** `Phase1-CollectResources.ps1`

#### **What It Does:**
1. **Azure Authentication**
   - Connects to specified Azure tenant
   - Retrieves all enabled subscriptions

2. **Availability Zone Mapping**
   - Queries physical zone IDs per subscription
   - Maps logical zones (Zone 1/2/3) → physical zones (actual datacenter IDs)
   - **Why Important:** Logical Zone 1 in Subscription A ≠ Logical Zone 1 in Subscription B
   - Enables accurate zone-wide failure impact assessment

3. **Resource Graph Query**
   - Collects **ALL resources** across subscriptions (VMs, databases, storage, networking, etc.)
   - Extracts: Name, Type, Location, SKU, Tags, Properties, Resource ID

4. **Zone Redundancy Classification**
   - Analyzes each resource's zone configuration
   - **Categories:**
     - ✅ **Zone Redundant** — Service automatically distributes across zones (e.g., Standard Public IP with ZRS)
     - ⚠️ **Zonal** — Pinned to single zone (zone failure = service outage)
     - ❌ **Non-Zonal** — No zone redundancy (regional service or locally redundant)
     - ℹ️ **Not Applicable** — Service not zone-aware

5. **Cross-Region DR Readiness Assessment**
   - Checks for geo-redundant configurations:
     - Storage: GRS, RA-GRS, GZRS, RA-GZRS
     - SQL Database: Auto-Failover Groups, Geo-Replication
     - Cosmos DB: Multi-region writes, read replicas
   - **Key Insight:** Identifies resources relying ONLY on zones without cross-region protection

6. **Dependency Extraction**
   - Parses resource properties to identify dependencies:
     - VNet/Subnet references
     - Key Vault secret/certificate URIs
     - Storage account endpoints
     - Managed Identity assignments
   - **Use Case:** Understand blast radius of DR failover

7. **Report Generation**
   - **54 CSV files** (one per resource type) → `ResourceTypes/microsoft_compute_virtualmachines.csv`
   - **Master CSV** (all resources) → `MasterReport_AllResources.csv`
   - **Zone mapping CSV** → `ZoneMappings_AllSubscriptions.csv`
   - **Summary CSVs** → By resource type and by subscription
   - **Interactive HTML Dashboard** → Visual heatmap and statistics

#### **Phase 1 Output Example:**
```
Total Subscriptions: 3
Total Resources: 339
Zone Redundant: 1 (0.3%)
Locally Redundant: 29 (8.6%)
Non-Zonal: 51 (15.0%)
```

**⚠️ Key Finding:** Only 0.3% zone redundant → **Cross-region DR is critical**

---

### **Phase 2: BCDR Recommendation Enrichment**

**Script:** `Phase2-AddRecommendations.ps1`

#### **What It Does:**

1. **Load Phase 1 Data**
   - Imports `MasterReport_AllResources.csv` from Phase 1
   - Validates data integrity and row count

2. **Smart Workload Tier Classification** 🆕
   - **Purpose:** Automatically detect environment type to optimize DR costs
   - **Detection Logic:**
     ```
     Priority 1: Subscription Name (High Confidence)
       Keywords: prod, production, prd, live → Production
       Keywords: dev, test, tst, uat, qa, stage → Dev/Test
       Keywords: sandbox, poc, demo, trial, lab → Sandbox
       Keywords: nonprod, non-prod → Non-Production
     
     Priority 2: Resource Tags (Medium Confidence)
       Tags: Environment, Env, Tier, Stage, Workload, AppTier
       Values: Production, Prod, Live → Production
       Values: Development, Dev, Test, UAT, QA → Dev/Test
       Values: Sandbox, POC, Demo → Sandbox
     
     Priority 3: Resource Group Name (Low Confidence)
       Pattern matching on RG name (same keywords as subscription)
     
     Priority 4: Resource Name (Lowest Confidence)
       Last resort pattern matching on resource name
     
     Default: Production (Assumed) — Low confidence
       ⚠️ Warning added to SA_ActionRequired: "Classification uncertain, customer must confirm"
     ```
   
   - **Strategy Adjustment:**
     - **Sandbox:** `No DR Required (Recreate from IaC)` — Zero cost
     - **Dev/Test:** `Backup & Restore (Cost-Optimized)` — 24h RPO, daily backup acceptable
     - **Non-Production:** `Backup & Restore` — Standard backup strategy
     - **Production:** Original Active-Passive or Active-Active strategies

3. **Criticality Assessment**
   - **Input Factors:**
     - Subscription name patterns
     - Resource group naming
     - Resource name patterns
     - Resource type (VMs, AKS, SQL = critical infrastructure)
   
   - **Classification:**
     - **Critical (Suggested):** Production + critical infrastructure → VMs, AKS, SQL, Storage, Key Vault
     - **High (Suggested):** Production + other resources
     - **Medium (Suggested):** Unclear production indicators or critical infrastructure without "prod" naming
     - **Low (Suggested):** Non-critical resources
     - **Not Assessed (Dev/Test):** Subscription/RG/Resource contains dev/test/sandbox keywords

4. **Priority Assignment**
   - **Mapping:**
     - Critical → **P1 (Suggested — Confirm with Customer)**
     - High → **P2 (Suggested — Confirm with Customer)**
     - Medium + No Zone Redundancy → **P2** (infrastructure gap detected)
     - Medium + Zone Redundancy → **P3** (some HA configured)
     - Low → **P4 (Suggested — Confirm with Customer)**
     - Not Assessed (Dev/Test) → **Not Prioritised** (Customer to assess DR need)

5. **BCDR Strategy Knowledge Base (54 Resource Types)**
   - Service-specific DR strategy recommendations:
     - **Virtual Machines:** Azure Site Recovery (ASR) to customer-chosen DR region
     - **SQL Database:** Auto-Failover Groups with async geo-replication
     - **AKS:** GitOps + IaC redeployment or warm standby cluster in DR region
     - **Storage Accounts:** Object Replication (Qatar Central) or GRS (West Europe)
     - **Cosmos DB:** Multi-region writes or read replicas
     - **Key Vault:** Custom sync via Azure Functions (no native geo-replication)
     - **And 48+ more services...**

6. **Qatar Central Specific Constraints Applied**
   - **No Paired Region:** Manual DR region selection (West Europe, North Europe recommended)
   - **GRS Not Available:** Use Object Replication for Blob Storage, AzCopy/ADF for Azure Files
   - **Azure Backup CRR Not Available:** Engage Microsoft Engineering for Region of Choice (RoC) preview
   - **Zone Redundancy Restricted:** One AZ at capacity, ZRS blocked for most services
   - **IaC Parity Critical:** Without IaC, DR provisioning can take days (per Microsoft Qatar BCDR Plan)

7. **RPO/RTO Target Assignment**
   - Based on BCDR strategy:
     - **Active-Active:** < 5 min RTO, near-zero RPO
     - **Active-Passive (Warm):** < 30 min RTO, < 1 hour RPO
     - **Active-Passive (Cold):** < 4 hours RTO, < 24 hours RPO
     - **Backup & Restore:** 4-12 hours RTO, 24 hours RPO
     - **IaC DR:** 1-2 hours RTO (pipeline execution time), near-zero RPO (if state externalized)

8. **Implementation Effort & Cost Analysis**
   - **Effort Estimates:**
     - Low: < 2 hours (e.g., enable ACR geo-replication)
     - Medium: 1-2 days (e.g., configure Storage Object Replication)
     - High: 1-2 weeks (e.g., design and deploy VNet topology in DR region)
   
   - **Cost Impact:**
     - Low: < $100/month (backup storage only)
     - Medium: $500-2000/month (warm standby, read replicas)
     - High: $5000+/month (Active-Active, full duplicate infrastructure)

9. **Action-Required Steps Generation**
   - Per-resource action items with specific Azure CLI/PowerShell commands
   - Example (SQL Database):
     ```
     1. Create SQL Server in West Europe (DR region)
     2. Configure Auto-Failover Group: az sql failover-group create ...
     3. Test failover quarterly
     4. Document failover procedure in runbook
     5. Obtain DPO approval for cross-border data replication (Qatar PDPPL)
     ```

10. **Comprehensive Excel Report (12 Sheets)**
    - **Introduction:** Methodology, tier detection logic, priority criteria, filtering tips
    - **Executive_Summary:** Key findings, Qatar constraints, top 10 recommendations
    - **SA_Recommendations:** Full enriched inventory with all SA_* columns
    - **QuickWins:** Low-effort, high-value subset
    - **P1_Critical_Actions:** Suggested P1 resources (customer validation required)
    - **Summary_ByResourceType:** Aggregated gap analysis
    - **Summary_BySubscription:** Subscription-level overview
    - **Timeline_ActionPlan:** Phased implementation (Foundation → Quick Wins → P1 → P2/P3 → Testing)
    - **DR_Testing_Plan:** Quarterly testing template with rollback procedures
    - **Dependencies_Matrix:** Service dependencies from environment data
    - **Compliance_Checklist:** Qatar PDPPL, NIA/NCSA, DPO sign-off tracking
    - **BCDR_Strategy_Reference:** Strategy design patterns, Qatar constraints, decision matrix

11. **HTML Dashboard Update**
    - Replaces placeholder with clickable BCDR report card
    - Shows actual Excel file path and 12-sheet overview
    - Opens dashboard in browser automatically

#### **Phase 2 Output Example:**
```
Total Resources Enriched: 339
P1 Critical (Suggested): 45
P2 High (Suggested): 78
Quick Win Actions: 12
Production Tier: 245
Dev/Test Tier: 67
Sandbox Tier: 27
```

---

## 📈 Output Reports

### **1. Comprehensive Excel Report (12 Sheets)**

#### **Sheet 1: Introduction**
- Report overview and purpose
- **Assessment Automation Framework** section with developer attribution
- Worksheet guide (12 sheets explained)
- SA_ column definitions
- **ASSESSMENT METHODOLOGY** 🆕
  - Step 1: Criticality Assessment (detailed table with conditions and results)
  - Step 2: Priority Assignment (criticality + zone redundancy → P1/P2/P3/P4)
  - Classification examples with 5 real-world scenarios
  - **Smart Workload Tier Auto-Detection** explanation
- Filtering & navigation tips (25+ filter combinations)
- Qatar PDPPL and NIA/NCSA compliance notes

#### **Sheet 2: Executive Summary**
- Customer name and assessment date
- **Automation Developed By:** Zahir Hussain Shah, Microsoft Qatar | Build Version 1.0
- Assessment disclaimer (priorities are suggested, not final)
- **Methodology Reference** pointing to Introduction sheet
- Key findings: Zone redundancy status, GRS availability, Azure Backup RoC status
- Resource statistics (total, P1, P2, Quick Wins, DR gaps, by region)
- Top 10 recommendations (Azure Backup RoC, Object Replication, ASR, Cosmos DB multi-region, etc.)
- Compliance notes (Qatar PDPPL, NIA/NCSA Certificate ID: 10018)

#### **Sheet 3: SA_Recommendations (Main Working Sheet)**
Columns include:
- **Identity:** SubscriptionName, ResourceGroup, ResourceName, ResourceType, Location, Kind, SKU
- **Assessment:** ZoneRedundancyStatus, CrossRegionStatus, PairedRegion, GeoRedundant, DRReady
- **Workload Tier:** 🆕 SA_WorkloadTier, SA_TierConfidence, SA_TierSource
- **SA Recommendations:** SA_Criticality, SA_Priority, SA_BCDRStrategy, SA_DRRegionChoice, SA_DRMethod, SA_RPO, SA_RTO, SA_ActionRequired, SA_ImplementationEffort, SA_CostImpact, SA_QuickWin, SA_Dependencies, SA_ComplianceNote
- **All Resource Tags:** Tag_Environment, Tag_Owner, Tag_CostCenter, etc.

**Row Color-Coding:**
- Red background: Critical resources
- Orange background: High priority
- Yellow background: Medium priority
- Gray background: Dev/Test (Not Assessed)

**SA_WorkloadTier Cell Color-Coding:** 🆕
- Green: Production (confirmed)
- Yellow: Non-Production
- Blue: Dev/Test
- Gray: Sandbox
- Orange: Production (Assumed) — Low confidence, needs review

#### **Sheet 4: QuickWins**
Subset filtered by `SA_QuickWin = "Yes"`
- Low implementation effort (< 2 hours to 2 days)
- Immediate value (cost savings, compliance, or RTO improvement)
- Examples:
  - Enable ACR Premium Geo-Replication to West Europe
  - Configure Automation Account + Key Vault sync
  - Add West Europe endpoints to Traffic Manager
  - Enable geo-redundant backup on MySQL/PostgreSQL

#### **Sheet 5: P1_Critical_Actions**
Subset filtered by `SA_Priority LIKE "P1*"`
- **Includes workload tier columns** 🆕 for validation
- Suggested highest-priority resources (customer must confirm actual criticality)
- Action-required steps with effort and cost estimates
- Example P1 resources:
  - VMs named "vm-prod-*" in production subscriptions
  - SQL Databases in critical infrastructure resource groups
  - AKS clusters without cross-region backup
  - Key Vaults storing production secrets

#### **Sheet 6: Summary_ByResourceType**
Aggregated statistics per Azure service type:
- Total count, Critical count, High count, Medium count, Low count, Dev/Test count
- Quick Win count, P1 Actions count
- Common BCDR strategy, DR method, RPO, RTO for that resource type

#### **Sheet 7: Summary_BySubscription**
Subscription-level rollup:
- Total resources per subscription
- Zone redundant count, Locally redundant count, Non-zonal count
- P1 count, Quick Win count
- DR ready resources vs. DR gap resources
- Subscription-level recommendations

#### **Sheet 8: Timeline_ActionPlan**
Phased implementation roadmap:
- **Phase 0: Foundation (Week 1-2)** — VNet topology in DR region, IaC repository setup
- **Phase 1: Quick Wins (Week 3-4)** — ACR geo-replication, Storage Object Replication, Automation Account sync
- **Phase 2: P1 Critical (Month 2-3)** — ASR for VMs, SQL Failover Groups, AKS DR strategy, Cosmos DB multi-region
- **Phase 3: P2/P3 (Month 4-6)** — Remaining production resources, non-critical workloads
- **Phase 4: Testing & Validation (Ongoing)** — Quarterly DR tests, runbook updates

**Disclaimer:** Timeline is TENTATIVE and GENERAL GUIDANCE. Customer to adjust based on budget, resources, and business priorities.

#### **Sheet 9: DR_Testing_Plan**
Quarterly DR testing template:
- **Test Objectives:** Validate RTO/RPO, verify runbooks, train staff
- **Pre-Test Checklist:** Backup verification, change freeze, stakeholder notification
- **Test Scenarios:** Planned failover, unplanned failover, failback
- **Success Criteria:** RTO achieved, data loss within RPO, applications functional
- **Rollback Procedure:** Step-by-step failback instructions
- **Post-Test Review:** Lessons learned, runbook updates, improvement actions

#### **Sheet 10: Dependencies_Matrix**
Service dependencies extracted from real environment:
- Resource → Dependent Service mappings
- Example: VM "vm-prod-web-01" → Depends on: VNet (vnet-prod-001), Subnet (snet-web), Key Vault (kv-prod-secrets), Storage (stgprodboot)
- **Use Case:** Understand which resources must failover together
- **Blast Radius Analysis:** If Key Vault fails, which VMs/Apps/Functions are impacted?

#### **Sheet 11: Compliance_Checklist**
Qatar BCDR compliance tracker:
- **Data Classification:** Public, Internal, Confidential, Restricted
- **Qatar PDPPL:** DPO approval for cross-border transfer, data residency documentation
- **NIA/NCSA Certification:** Verify DR region is NIA-certified (West Europe, North Europe)
  - Certificate ID: 10018 | Valid Until: 28 August 2026
- **Technical Readiness:** IaC parity, runbook documentation, access controls
- **Testing & Validation:** DR test results, RTO/RPO validation
- **Sign-Off Tracking:** Customer stakeholder approval checkboxes

#### **Sheet 12: BCDR_Strategy_Reference** 🆕
Complete BCDR strategy guide:
- **Strategy Patterns:** Active-Active, Active-Passive (Warm/Cold), Backup & Restore, Geo-Replication, IaC DR, Hybrid
  - Each pattern includes: Definition, When to Use, RPO/RTO, Cost, Example Services
- **Qatar Central Specific Constraints:**
  - No Paired Region → Impact and recommended approach
  - Zone Redundancy Restricted → Use cross-region DR, not zones alone
  - Azure Backup RoC Preview → Supported workloads, target regions (Sweden Central, Switzerland North only)
  - IaC Parity Critical → Most critical factor for RTO (per Microsoft Azure Qatar BCDR Plan)
  - NIA/NCSA Certification → West Europe and North Europe approved for Qatar government/regulated sectors
- **DR Strategy Decision Matrix:**
  - Tier 1 (Mission-Critical): < 5 min RTO → Active-Active or Hot Standby
  - Tier 2 (Production): < 30 min RTO → Active-Passive (Warm)
  - Tier 3 (Non-Critical): < 4 hours RTO → Backup & Restore or Cold Standby
- **Smart Workload Tier Auto-Detection:** 🆕
  - 4-priority detection logic explained
  - Tier-based strategy adjustment table
  - Cost impact comparison (Production vs. Dev/Test vs. Sandbox)
  - How to review and override tier classification
- **References:** Points to Microsoft Azure Qatar BCDR Plan and Azure Backup RoC documentation
- **Key Takeaways:** IaC priority, zone redundancy limitations, NIA certification requirements

---

### **2. Interactive HTML Dashboard**

**File:** `Dashboard_ZoneRedundancy.html`

**Features:**
- **Subscription-level statistics** (resource count, zone redundant %, DR ready %)
- **Zone redundancy heatmap** (color-coded by subscription)
- **Cross-region DR status** (geo-redundant vs. single-region resources)
- **Resource type breakdown** (top 10 resource types by count)
- **Clickable BCDR Report Card** 🆕
  - Blue gradient card with "OPEN FULL REPORT" button
  - Shows actual Excel file path
  - Lists 12 comprehensive worksheets, timeline plans, dependencies matrix, compliance checklist
  - One-click to open Excel report
  - Quick Tip section explaining report contents

---

### **3. CSV Data Files**

- **MasterReport_AllResources.csv** — Combined inventory (all subscriptions)
- **ZoneMappings_AllSubscriptions.csv** — Logical → Physical zone mapping
- **Summary_ZoneRedundancy.csv** — Aggregated by resource type
- **Summary_BySubscription.csv** — Aggregated by subscription
- **ResourceTypes/*.csv** — 54 individual files (one per resource type)

---

## 🔍 Assessment Methodology

### **Criticality Assessment (Step 1)**

**Input Data:**
1. Subscription name
2. Resource group name
3. Resource name
4. Resource type

**Classification Logic:**

| Condition | Result |
|-----------|--------|
| Subscription contains: `dev`, `test`, `tst`, `poc`, `sandbox`, `uat`, `qa`, `stage`, `stg` | **Not Assessed (Dev/Test)** |
| Name contains: `prod`, `prd`, `production`, `live` **AND** Critical Infrastructure Resource Type* | **Critical (Suggested)** |
| Name contains: `prod`, `prd`, `production`, `live` **AND** Other resource types | **High (Suggested)** |
| Name contains: `test`, `dev`, `poc`, `stage`, `sandbox`, `uat`, `qa` (in RG or resource name) | **Not Assessed (Dev/Test)** |
| No clear indicators **BUT** Critical Infrastructure Resource Type* | **Medium (Suggested)** |
| No clear indicators **AND** Non-critical resource type | **Low (Suggested)** |

**Critical Infrastructure Resource Types (*):**
- VirtualMachines, ManagedClusters (AKS), SQL/MySQL/PostgreSQL Databases, Cosmos DB, Storage Accounts, Key Vaults, NetApp Volumes, Recovery Services Vaults, API Management, Event Hubs, Service Bus, Azure VMware Solution, Redis Enterprise, Data Factory

---

### **Priority Assignment (Step 2)**

**Input:**
- Criticality from Step 1
- Zone Redundancy Status from Phase 1

**Priority Mapping:**

| Criticality | Zone Redundancy Status | Final Priority |
|-------------|------------------------|----------------|
| **Critical** | Any | **P1 (Suggested — Confirm with Customer)** |
| **High** | Any | **P2 (Suggested — Confirm with Customer)** |
| **Medium** | Non-Zonal or Locally Redundant | **P2 (Suggested — Confirm with Customer)** — Infrastructure gap detected; elevated priority |
| **Medium** | Zone Redundant or Zonal | **P3 (Suggested — Confirm with Customer)** — Some HA configured; lower priority |
| **Low** | Any | **P4 (Suggested — Confirm with Customer)** |
| **Not Assessed (Dev/Test)** | Any | **Not Prioritised (Dev/Test — Customer to Assess DR Need)** |

**⚠️ Important:** All priorities are **SUGGESTED** based on automated naming analysis. Customer must validate with business stakeholders to confirm actual criticality and DR requirements.

---

### **Workload Tier Auto-Detection (Cost Optimization)** 🆕

**Purpose:** Automatically detect environment type to assign cost-optimized DR strategies for non-production workloads.

**Detection Priority:**

1. **Subscription Name (High Confidence)**
   - Production keywords: `\bprod\b`, `production`, `\bprd\b`, `\blive\b`
   - Sandbox keywords: `sandbox`, `\bsbx\b`, `\bpoc\b`, `demo`, `trial`, `lab`
   - Dev/Test keywords: `\bdev\b`, `\btest\b`, `\btst\b`, `\buat\b`, `\bqa\b`, `\bstage\b`, `\bstaging\b`
   - Non-Prod keywords: `nonprod`, `non-prod`, `\bnp-`

2. **Resource Tags (Medium Confidence)**
   - Tags checked: `Environment`, `Env`, `Tier`, `Stage`, `Workload`, `AppTier`
   - Tag values: `Production`, `Prod`, `Live`, `Development`, `Dev`, `Test`, `UAT`, `QA`, `Sandbox`, `POC`

3. **Resource Group Name (Low Confidence)**
   - Pattern matching on RG name (same keywords as subscription)

4. **Resource Name (Lowest Confidence)**
   - Last resort pattern matching on resource name

5. **Default: Production (Assumed)**
   - If no keywords found → Default to Production with **Low confidence**
   - ⚠️ Warning added to SA_ActionRequired: "Classification uncertain, customer must confirm workload tier"

**Strategy Adjustment Based on Tier:**

| Tier | Adjusted Strategy | Adjusted Priority | Adjusted RPO | Adjusted RTO | Cost Impact | Example |
|------|-------------------|-------------------|--------------|--------------|-------------|---------|
| **Production** | Original (Active-Passive/Active-Active) | Original (P1/P2/P3) | < 1 hour | < 30 min | High | VM in "Prod-Subscription" → Active-Passive (ASR) |
| **Non-Production** | Backup & Restore | P2 (Non-Prod — Adjusted) | < 24 hours | < 8 hours | Medium | DB in "NonProd-Subscription" → Backup only |
| **Dev/Test** | Backup & Restore (Cost-Optimized) | P3 (Dev/Test — Reduced Priority) | 24 hours (daily backup) | 4-8 hours | Low | VM in "Dev-Subscription" → Daily backup |
| **Sandbox** | No DR Required (Recreate from IaC) | Not Prioritised (Sandbox) | N/A | Recreate from IaC | **Zero** | AKS in "Sandbox-Subscription" → No DR |
| **Production (Assumed)** — Low Confidence | Original + ⚠️ Warning | Original | Original | Original | High until confirmed | Resource with no clear indicators |

**Cost Optimization Example:**

**Before (treating all as Production):**
- Dev VM → Active-Passive (ASR) = ~$150/month
- Test SQL DB → Failover Group = ~$300/month
- Sandbox AKS → Warm standby = ~$500/month
- **Total: ~$950/month per environment**

**After (smart classification):**
- Dev VM → Backup & Restore = ~$20/month
- Test SQL DB → Backup only = ~$15/month
- Sandbox AKS → No DR = $0
- **Total: ~$35/month — 96% cost reduction!**

---

## 🇶🇦 Qatar-Specific Compliance

### **Based on Microsoft Azure Qatar BCDR Plan**

This framework implements recommendations from the **Microsoft Azure Qatar BCDR Plan** document, which provides service-specific DR strategies, prerequisites, backup approaches, and checklists for Azure PaaS and IaaS services in Qatar Central.

**Key Guidance from the Plan:**
> "IaC parity is the MOST CRITICAL FACTOR for RTO in Qatar Central due to lack of paired region automation. Without IaC, DR provisioning can take days."

All recommendations in this framework emphasize Infrastructure-as-Code (Bicep, ARM, Terraform) as mandatory for production workloads.

---

### **Azure Backup Region of Choice (RoC) — Preview**

**Background:**
- Qatar Central has **no paired Azure region** due to data residency requirements
- Standard **Cross-Region Restore (CRR)** for Azure Backup is NOT available
- Microsoft Engineering provisioned **"Region of Choice (RoC)"** preview feature

**Supported Workloads:**
- ✅ IaaS VMs (Azure Virtual Machines)
- ✅ SQL Server in Azure VM
- ✅ SAP HANA in Azure VM
- ✅ Azure File Share (AFS)

**NOT Supported Yet:**
- ❌ PostgreSQL Flexible Server Backup
- ❌ AKS (Azure Kubernetes Service) Backup cross-region replication

**Target Regions:**
- **Sweden Central (SDC)** ✅
- **Switzerland North (SZN)** ✅
- ❌ West Europe, North Europe, UAE North (NOT available as RoC targets)

**Action Required:**
1. Engage Microsoft account team to whitelist subscription for RoC preview
2. Create Recovery Services Vault in Sweden Central or Switzerland North
3. Configure backup policies for supported workloads
4. Enable Soft Delete and Multi-User Authorization (MUA) on vault
5. Document RoC as preview feature in compliance checklist

**For West Europe Resources:**
- Standard CRR to North Europe (paired region) **IS available** — no Engineering engagement required

---

### **NIA/NCSA Compliance**

**Certification Details:**
- **Certificate Holder:** Microsoft Azure Qatar Program
- **Standard:** NIA V2.0 (National Information Assurance, Qatar)
- **Certificate ID:** 10018
- **Issued:** 29 August 2023
- **Valid Until:** 28 August 2026
- **Certification Body:** Qatar National Cyber Security Agency (NCSA)
- **Accredited Auditor:** Forvis Mazars LLC

**Scope — Certified Regions:**
1. ✅ **Qatar Central (Doha)** — Primary region
2. ✅ **West Europe (Amsterdam)** — NIA-certified DR region
3. ✅ **North Europe (Dublin)** — NIA-certified DR region

**NOT Certified:**
- ❌ Sweden Central (used for Azure Backup RoC)
- ❌ Switzerland North (used for Azure Backup RoC)
- ❌ UAE North, East US, other regions

**Compliance Statement:**
> "This single, consolidated certificate confirms all three regions (Qatar Central, West Europe, North Europe) meet Qatar Information Security Management System (ISMS) requirements per NIA V2.0 standards."

**Implications for DR Region Selection:**

| Workload Type | DR Region Recommendation | Rationale |
|---------------|--------------------------|-----------|
| **Qatar Government** | West Europe or North Europe | NIA/NCSA certification mandatory |
| **Qatar Public Sector** | West Europe or North Europe | NIA/NCSA certification required |
| **Regulated Industries** (Finance, Healthcare) | West Europe or North Europe | NIA/NCSA certification strongly recommended |
| **Private Sector (Non-Regulated)** | Customer choice: West Europe, North Europe, UAE North, or other | NIA certification not mandatory, consider latency and cost |
| **Azure Backup Vault** (RoC) | Sweden Central or Switzerland North (preview) | RoC technical constraint — NOT NIA-certified. Document in compliance checklist and obtain customer security/DPO approval. |

**Action Required:**
- Customer to conduct **data classification workshop**
- Categorize resources by sensitivity: Public, Internal, Confidential, Restricted
- For **Confidential/Restricted data** → Prefer NIA-certified regions (West Europe, North Europe)
- For **Azure Backup RoC** to Sweden Central/Switzerland North → Document that backup vaults are NOT in NIA-certified regions
- Submit DR replication plan to **Data Protection Officer (DPO)** for approval per Qatar PDPPL

---

### **Qatar PDPPL (Personal Data Protection Privacy Law)**

**Background:**
- Qatar PDPPL governs processing and **cross-border transfer** of personal data
- Applies to: Government, public sector, private sector organizations in Qatar

**Key Requirement:**
- **Cross-border data transfer** to West Europe, North Europe, or any secondary region requires **DPO review and approval**

**Assessment Action Items:**
1. **Data Classification Workshop**
   - Classify all Azure resources by data sensitivity
   - Identify resources containing personal data (PII, customer data)

2. **Cross-Border Transfer Impact Assessment**
   - Document which resources will replicate to DR region
   - Assess data residency and sovereignty implications

3. **DPO Approval**
   - Submit DR replication plan to Data Protection Officer
   - Obtain written approval for cross-border data movement
   - Document approval in Compliance_Checklist sheet

4. **PDPPL Compliance Checklist Items:**
   - [ ] Data classification completed (Public/Internal/Confidential/Restricted)
   - [ ] Cross-border data transfer documented
   - [ ] DPO approval obtained
   - [ ] Data Processing Agreement (DPA) with Microsoft reviewed
   - [ ] Employee training on data handling during DR events
   - [ ] Incident response plan includes PDPPL notification requirements

---

### **⚠️ Cross-Region DR Emphasis**

**Why This Framework Prioritizes Cross-Region DR Over Availability Zones:**

1. **Availability Zones Protect Against:**
   - ✅ Single datacenter failure (one AZ outage)
   - ✅ Infrastructure failures (power, cooling, networking within one zone)
   - ✅ Planned maintenance in one zone

2. **Availability Zones DO NOT Protect Against:**
   - ❌ **Region-wide failures** (entire Qatar Central region outage)
   - ❌ **Natural disasters** (earthquakes, floods affecting entire region)
   - ❌ **Geo-political events** (regional conflicts, compliance changes)
   - ❌ **Catastrophic infrastructure failures** (regional network outage, control plane failure)
   - ❌ **Regulatory changes** requiring immediate data residency moves

3. **Qatar Central Specific Risks:**
   - **Zone Redundancy RESTRICTED:** One availability zone is at full capacity — ZRS blocked for most services
   - **No Paired Region:** Cannot leverage automatic Azure paired region DR (GRS, CRR, Service Bus Geo-DR)
   - **Relying on zones alone = single point of failure** at region level

**Best Practice — Defense in Depth:**
```
Layer 1: Availability Zones (In-Region HA)
   ↓ Protects against: Single AZ failure
   
Layer 2: Cross-Region DR (Geo-Redundancy)
   ↓ Protects against: Entire region failure
   
Layer 3: IaC + Automation (Rapid Redeployment)
   ↓ Protects against: DR provisioning delays
```

**Microsoft Recommendation:**
> "Use availability zones for **in-region high availability** where available. Always implement **cross-region disaster recovery** for business-critical workloads. Zones are NOT a substitute for geo-redundancy."

---

## 🛡️ BCDR Strategies Recommended

### **1. Active-Active (Multi-Region)**

**Definition:** Both regions are LIVE and actively serving production traffic simultaneously. Traffic is load-balanced across regions using Azure Front Door or Traffic Manager.

**When to Use:**
- Mission-critical workloads requiring **zero downtime**
- Tier-1 applications with strict SLA requirements (e.g., e-government portals, financial trading platforms, healthcare critical systems)

**RPO:** Near-zero (synchronous or near-synchronous replication)  
**RTO:** Near-zero (traffic automatically re-routes to healthy region)  
**Cost:** High (2x compute cost — both regions fully provisioned and running)

**Example Services:**
- Azure Cosmos DB (multi-region writes) — Automatic failover with <1 min RTO
- Azure Container Registry Premium (Geo-Replication) — Single endpoint, multiple replicas
- Azure Static Web Apps — Global CDN distribution built-in
- Azure Traffic Manager / Azure Front Door — Global load balancers

---

### **2. Active-Passive (Warm Standby)**

**Definition:** Primary region is LIVE. Secondary region is **pre-provisioned** but idle/standby. On DR event, traffic is manually or automatically re-routed to secondary. Secondary can be "warm" (pre-deployed but scaled down) or "hot" (pre-deployed at full scale).

**When to Use:**
- **Most common pattern** for production Azure workloads
- Production workloads with moderate RTO requirements (5-30 minutes)
- Cost-optimized alternative to Active-Active

**RPO:** Depends on replication method (typically <5 min for databases, near-real-time for VMs via ASR)  
**RTO:** <30 minutes with automation; <5 minutes with health probe-based traffic re-routing (Azure Front Door)  
**Cost:** Medium (secondary infrastructure pre-provisioned but possibly scaled down)

**Example Services:**
- **Virtual Machines:** Azure Site Recovery (ASR) to West Europe — Continuous replication, manual failover
- **Azure SQL Database:** Auto-Failover Groups — Async replication, automatic or manual failover
- **Azure App Service:** Standby instance in West Europe — Pre-deployed app, traffic switched via Front Door
- **AKS:** Standby cluster in West Europe — IaC-provisioned, scaled to minimal size, scaled up on DR
- **Azure MySQL/PostgreSQL:** Cross-Region Read Replica — Async replication, manual promotion to writable

---

### **3. Active-Passive (Cold Standby)**

**Definition:** Primary region is LIVE. Secondary region has **NO pre-provisioned infrastructure**. DR resources are deployed **ON-DEMAND** during a DR event using Infrastructure-as-Code (IaC).

**When to Use:**
- Non-production environments (dev/test) with relaxed RTO/RPO
- Cost-sensitive scenarios where pre-provisioning secondary is not justified
- Workloads fully recreatable from IaC

**RPO:** Dependent on data backup frequency (hours to days)  
**RTO:** Hours (infrastructure provisioning + data restore)  
**Cost:** Low (no secondary compute cost — pay only for backup storage)

**Example Services:**
- **Virtual Machines:** IaC Redeployment + Backup Restore from Recovery Services Vault — Redeploy VMs via Bicep/ARM during DR
- **Azure App Service:** IaC Redeployment — Redeploy app via ARM template or CI/CD pipeline
- **Azure Logic Apps:** Git-based redeployment — Recreate from source-controlled definitions
- **Azure Key Vault:** Manual creation + export/import — Rebuild vault and restore secrets

---

### **4. Backup & Restore**

**Definition:** Periodic backups of data are stored in geo-redundant storage. On DR event, data is restored from backup to a new or existing service instance in the DR region. No live secondary.

**When to Use:**
- Non-critical workloads with longer acceptable RTO/RPO (hours)
- Data protection scenarios (corruption, accidental deletion, ransomware)
- Compliance-driven retention (7 years data retention)

**RPO:** Backup frequency (daily = 24h RPO; 4-hourly enhanced policy = 4h RPO)  
**RTO:** Restore time (4-12 hours depending on data size)  
**Cost:** Low (backup storage only)

**Example Services:**
- **Azure Backup (Recovery Services Vault):** VM backup, SQL backup, Azure Files backup — Restore to West Europe or other regions
- **Azure SQL Database:** Geo-Redundant Backup + Geo-Restore — Automated daily backups, restore to any Azure region
- **Azure Blob Storage:** Soft Delete + Versioning + Immutable Blobs — Protect against accidental deletion and ransomware
- **Azure NetApp Files:** ANF Backup — On-demand snapshots + cross-region backup

---

### **5. Geo-Replication (Platform-Managed)**

**Definition:** Azure service automatically replicates data to secondary region. Failover is manual or automatic depending on service.

**When to Use:**
- Leveraging built-in Azure service resilience without custom configuration
- Services that support native geo-replication

**RPO:** Near-zero (<15 seconds)  
**RTO:** Automatic (service-managed) or manual (customer-initiated)  
**Cost:** Typically included in Premium/Standard tiers

**Example Services:**
- **Azure Storage:** GRS, RA-GRS, GZRS, RA-GZRS — Async replication to paired region (NOT available for Qatar Central — use Object Replication instead)
- **Azure Cosmos DB:** Multi-Region Replication — Async replication with automatic failover priorities
- **Azure Event Hubs Premium:** Geo-Replication — Continuous replication of events AND metadata
- **Azure Service Bus Premium:** Geo-Replication — Continuous replication of messages AND configuration

---

### **6. Infrastructure-as-Code (IaC) DR**

**Definition:** All infrastructure is defined in code (Bicep, Terraform, ARM). DR region infrastructure is provisioned via automated deployment pipelines. Critical for stateless services and serverless architectures.

**When to Use:**
- Cloud-native applications, microservices, serverless (Functions, Logic Apps), container orchestration (AKS)
- **Mandatory for Qatar Central workloads** due to lack of paired region automation

**RPO:** Near-zero (if state is externalized to geo-redundant storage)  
**RTO:** Pipeline execution time (15 minutes to 2 hours)  
**Cost:** Low (no standby infrastructure — deploy only on DR trigger)

**⚠️ Critical Success Factor (per Microsoft Azure Qatar BCDR Plan):**
> "IaC parity is the MOST CRITICAL FACTOR for RTO in non-paired regions. Without IaC, manual DR provisioning can take days."

**Example Services:**
- **AKS:** GitOps + IaC redeployment to West Europe
- **Azure Functions / App Service:** CI/CD pipeline redeploys to West Europe on DR trigger
- **Azure Virtual Networks:** VNet topology defined in Bicep; redeployed to West Europe
- **Azure API Management:** ARM export + redeployment (Premium Multi-Region is preferred but costly)
- **Azure Data Factory:** Git-integrated pipelines redeployed to West Europe ADF instance

---

### **7. No DR Required (Sandbox / POC)**

**Definition:** Sandbox, POC, demo, or trial environments with no business value. DR is not justified. If data is valuable, implement backup only.

**When to Use:**
- **Sandbox/POC environments** detected via smart workload tier classification
- Resources used for testing, demos, labs, training
- No production dependencies

**RPO:** N/A  
**RTO:** Recreate from IaC if needed (hours to days)  
**Cost:** **Zero** (no DR strategy = zero DR cost)

**⚠️ Warning:** Automatically assigned to resources in subscriptions/RGs with "sandbox", "poc", "demo", "trial", "lab" keywords. Customer to review SA_WorkloadTier column and confirm if DR is actually needed.

---

## 🔧 Supported Azure Services

This framework provides service-specific BCDR recommendations for **54 Azure resource types**:

### **Compute**
- Virtual Machines
- Azure Kubernetes Service (AKS / Managed Clusters)
- App Service (Web Apps)
- Azure Functions
- Container Apps
- Virtual Machine Scale Sets
- Azure Virtual Desktop
- Azure VMware Solution (AVS)

### **Databases**
- Azure SQL Database / Managed Instance
- MySQL Flexible Server
- PostgreSQL Flexible Server
- Cosmos DB
- Azure Database for MariaDB

### **Storage**
- Storage Accounts (Blob, Files, Queue, Table)
- Azure NetApp Files
- Managed Disks

### **Networking**
- Virtual Networks (VNets)
- Network Security Groups (NSGs)
- Public IP Addresses
- Azure Bastion
- NAT Gateway
- Azure Firewall
- Application Gateway
- Load Balancer
- Azure Front Door
- Traffic Manager
- Private Endpoints
- Virtual Network Gateways (VPN/ExpressRoute)
- Route Tables
- Private DNS Zones

### **Security & Identity**
- Key Vault
- Managed Identities
- Azure AD Domain Services

### **Management & Monitoring**
- Recovery Services Vault
- Log Analytics Workspace
- Application Insights
- Azure Monitor
- Automation Account

### **Integration & Messaging**
- Logic Apps
- Event Hubs
- Service Bus
- Event Grid
- API Management

### **AI & Machine Learning**
- Cognitive Services (AI Services)
- Machine Learning Workspace
- AI Search (Cognitive Search)

### **Container & Registry**
- Azure Container Registry (ACR)

### **And 21+ additional resource types...**

---

## ⚠️ Limitations & Disclaimers

### **Assessment Limitations**

1. **Suggested Priorities, Not Final:**
   - All P1/P2/P3/P4 priorities are **SUGGESTED** based on naming analysis
   - Customer must validate actual business criticality with stakeholders
   - Resource naming does not always reflect actual production status

2. **Workload Tier Auto-Detection:**
   - Detection is based on naming patterns and tags, not actual workload behavior
   - **Low confidence classifications** (Production (Assumed)) require manual review
   - Filter by `SA_TierConfidence = "Low"` to identify uncertain classifications

3. **Not a Definitive DR Mandate:**
   - This is **technical guidance** based on Azure best practices
   - Customer to decide: which workloads require DR, acceptable RTO/RPO, DR region choice, implementation timeline, budget

4. **Dev/Test Resources:**
   - Resources in dev/test subscriptions marked as "Not Assessed"
   - Customer to confirm whether DR is needed for these environments

### **Qatar Central Constraints**

1. **No Paired Region:**
   - Qatar Central has NO Azure-designated paired region
   - Native GRS/GZRS for Storage is NOT available
   - Cross-Region Restore (CRR) for Recovery Services Vault is NOT available
   - Customer must manually select DR region (West Europe, North Europe, UAE North, etc.)

2. **Zone Redundancy Restricted:**
   - One availability zone is at full capacity in Qatar Central
   - Zone-redundant SKUs (ZRS, GZRS) are blocked or unavailable for many services
   - **Do NOT rely solely on zone redundancy** — cross-region DR is mandatory

3. **Azure Backup Region of Choice (RoC) — Preview:**
   - RoC is a **PREVIEW feature** — confirm current availability with Microsoft account team
   - Supported workloads: IaaS VM, SQL in VM, SAP HANA in VM, Azure Files
   - NOT supported: PostgreSQL, AKS
   - Target regions: Sweden Central or Switzerland North ONLY (NOT NIA-certified)

4. **IaC Parity is Critical:**
   - Per Microsoft Qatar BCDR Plan: "IaC parity is the MOST CRITICAL FACTOR for RTO"
   - Without IaC, manual DR provisioning can take **days**
   - ALL production infrastructure must be in Bicep, ARM, or Terraform

### **Opting for DR**

**Implementing DR is a customer decision.** Microsoft provides best practice guidance and technical support. The choice to implement, scope, or defer DR for any workload remains entirely with the customer.

Factors to consider:
- Business criticality and revenue impact
- Acceptable downtime (RTO) and data loss (RPO)
- Regulatory compliance requirements (Qatar PDPPL, NIA/NCSA)
- Budget and resource availability
- Operational complexity and support capability

---

## 📚 Version History

### **Version 1.0** — April 2026

**Major Features:**
- ✅ 2-phase automated assessment (Zone Redundancy + SA BCDR Recommendations)
- ✅ Single-command execution wrapper (`Start-AzureBCDRAssessment.ps1`)
- ✅ 12-sheet comprehensive Excel report with Qatar NIA/NCSA compliance
- ✅ Smart workload tier classification (Production/Non-Prod/Dev-Test/Sandbox)
- ✅ Cost-optimized DR strategy assignment per environment tier (96% cost reduction potential)
- ✅ Priority & criticality methodology documentation (2-step process fully explained)
- ✅ Interactive HTML dashboard with clickable BCDR report card
- ✅ Dependency mapping from real environment data
- ✅ Compliance checklist with DPO sign-off tracking
- ✅ BCDR_Strategy_Reference sheet with Qatar constraints and decision matrix
- ✅ Azure Backup Region of Choice (RoC) preview support
- ✅ Cross-region DR emphasis over availability zones
- ✅ NIA/NCSA Certificate ID: 10018 validation for DR region selection

**Framework Capabilities:**
- Automated zone redundancy analysis across 54 resource types
- Service-specific BCDR recommendations per resource type
- Qatar Central constraints handling (no paired region, zone redundancy restrictions)
- 4-priority workload tier detection hierarchy (subscription → tags → RG → resource name)
- Timeline & phased implementation roadmap (Foundation → Quick Wins → P1 → P2/P3 → Testing)
- Quarterly DR testing plan template with rollback procedures
- All Azure resource tags included for advanced filtering

**Developed By:**
- Zahir Hussain Shah, Sr. Solution Engineer — Cloud & AI Infrastructure, Microsoft Qatar

---

## 📧 Support & Contact

For questions, technical clarifications, or DR implementation support:

- **Microsoft Account Team** — Engage your Microsoft Solution Architect or Customer Success Manager
- **Automation Developer** — Zahir Hussain Shah, Microsoft Qatar

---

## 📄 License

© Microsoft Corporation. Internal use only.

---

## 🙏 Acknowledgments

This framework is based on:
- **Microsoft Azure Qatar BCDR Plan** — Service-specific DR strategies and Qatar Central constraints
- **Microsoft Azure Backup Region of Choice (RoC)** — Engineering Team Content
- **Azure Well-Architected Framework** — BCDR best practices
- **Microsoft NIA/NCSA Certification** — Qatar compliance requirements (Certificate ID: 10018)

---

**Build Version:** 1.0  
**Last Updated:** April 2026  
**Framework:** PowerShell 7 + Azure PowerShell + ImportExcel  
**Tested With:** Azure PowerShell 11.x, ImportExcel 7.x

---

🚀 **Ready to assess your Azure environment? Run the assessment now:**

```powershell
.\Start-AzureBCDRAssessment.ps1 -CustomerName "Your Organization" -TenantId "your-tenant-id"
```
