# Azure BCDR Assessment Automation Framework

![Azure](https://img.shields.io/badge/Azure-0078D4?style=flat&logo=microsoft-azure&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=flat&logo=powershell&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-2.0-blue.svg)

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

> ⚠️ **CRITICAL GUIDANCE:** This framework **emphasizes cross-region disaster recovery** over reliance on availability zones alone. Availability zones do NOT protect against region-wide failures. **Cross-region DR is essential** for business-critical workloads — especially in Qatar Central where zone redundancy is restricted and there is no Azure-designated paired region.

> 📌 **ADVISORY NOTE:** All recommendations, priorities, criticality labels, workload tier classifications, RPO/RTO targets, effort estimates, and cost estimates produced by this tool are **suggestions only**. Microsoft does not prescribe or mandate any BCDR approach. The customer decides which workloads require DR, what strategies to implement, acceptable RTO/RPO, DR region of choice, implementation timeline, and budget. Always validate suggestions against your own business requirements, compliance obligations, and risk tolerance.

---

## 📋 Table of Contents

- [What's New in v2.0](#-whats-new-in-v20)
- [Key Features](#-key-features)
- [Architecture](#-architecture)
- [Prerequisites](#-prerequisites)
- [Quick Start](#-quick-start)
- [Output Reports](#-output-reports)
- [Assessment Methodology](#-assessment-methodology)
- [Qatar-Specific Compliance](#-qatar-specific-compliance)
- [BCDR Strategies Recommended](#-bcdr-strategies-recommended)
- [Supported Azure Services](#-supported-azure-services)
- [Limitations & Disclaimers](#%EF%B8%8F-limitations--disclaimers)
- [Version History](#-version-history)

---

## 🆕 What's New in v2.0

### **Dashboard — Microsoft Design System Overhaul**
- Full Microsoft-branded UI with CSS custom properties (`--ms-blue`, `--ms-navy`)
- Top bar with Microsoft SVG logo and generation timestamp
- Redesigned header with 4 live stat tiles (Zone Redundant / Non-Zonal / Locally Redundant / Subscriptions)
- 6 summary cards with dual-label system: business-friendly name + technical badge + scope pill
- 3 new charts: **NonZonal Breakdown** (top 12 resource types), **IaaS/PaaS/Platform Split**, **Subscription Risk Score** (0–100 normalized)
- Cross-Region Replication card replaced with a 4-row mutually exclusive breakdown table (Geo-Redundant / Global-Multi-Region / Single-Region / Needs Manual Check) with a ∑ Total row for verification
- Professional navy footer with Azure docs link and disclaimer

### **Excel Report — 16 Sheets (was 12)**
- **P2_Actions** sheet added — suggested P2 (High) priority resources
- **P3_Actions** sheet added — suggested P3 (Medium/Dev-Test reduced) priority resources
- **VM_Zone_Planner** now includes a yellow legend row (row 2) explaining why `LogicalZones` / `PhysicalZones` are empty (expected for Qatar Central where customer-accessible AZs are blocked)
- `FreezePanes` updated to row 3 in VM_Zone_Planner to keep both header and legend visible

### **"Suggested" Language Throughout — No Dictating**
All SA_ recommendation columns are now clearly marked as suggestions:

| Column | Example Value |
|--------|--------------|
| `SA_WorkloadTier` | `(Suggested) Production` |
| `SA_Criticality` | `Critical (Suggested)` |
| `SA_Priority` | `P1 (Suggested — Confirm with Customer)` |
| `SA_BCDRStrategy` | `(Suggested) Active-Passive (Warm Standby)` |
| `SA_DRRegionChoice` | `(Customer to confirm) West Europe ...` |
| `SA_DRMethod` | `(Suggested) Azure Site Recovery (ASR)` |
| `SA_RPO` | `(Suggested target) < 1 hour` |
| `SA_RTO` | `(Suggested target) < 30 min` |
| `SA_ImplementationEffort` | `(Estimated) High` |
| `SA_CostImpact` | `(Estimated) Medium` |
| `SA_ActionRequired` | `⚠️ ADVISORY — Suggested actions only. Customer must review and confirm...` |
| `SA_ZoneTransitionPath` | `(Suggested) Step 1: ...` |

### **Color-Coded SA_WorkloadTier Cells**
- 🟢 Green: `(Suggested) Production`
- 🟠 Orange: `(Suggested) Production (Assumed)` — uncertain, needs customer review
- 🟡 Yellow: `(Suggested) Non-Production`
- 🔵 Blue: `(Suggested) Dev/Test`
- ⚫ Gray: `(Suggested) Sandbox`

### **Softened Advisory Language**
- All `MANDATORY` and `Must Be` language replaced with recommended/suggested phrasing throughout action guidance
- Console output now shows P2 and P3 counts alongside P1 at completion

---

## 🚀 Key Features

### ✅ **2-Phase Automated Assessment Workflow**
- **Single command execution** orchestrates both phases automatically
- **Phase 1:** Resource collection, zone redundancy analysis, cross-region readiness, HTML dashboard
- **Phase 2:** BCDR recommendation enrichment — SA guidance, RPO/RTO, cost analysis, 16-sheet Excel report

### ✅ **Smart Workload Tier Classification**
Automatically detects environment type to optimize DR costs:
- **Production** → Active-Passive (Warm Standby) or Active-Active strategies
- **Non-Production** → Backup & Restore
- **Dev/Test** → Backup & Restore (Cost-Optimized) — 24h RPO acceptable
- **Sandbox** → No DR Required (Recreate from IaC) — Zero DR cost
- **Production (Assumed)** → Production strategy + ⚠️ warning to confirm

**Detection Priority (4-level hierarchy):**
1. Subscription name (High confidence)
2. Resource tags: `Environment`, `Tier`, `Stage`, `Workload`, `AppTier` (Medium confidence)
3. Resource group name patterns (Low confidence)
4. Resource name patterns (Lowest confidence)

**Cost Impact:** Potential **96% DR cost reduction** for non-production workloads by avoiding expensive ASR replications and warm standby clusters.

### ✅ **Comprehensive 16-Sheet Excel Report**
1. **Introduction** — Methodology, tier detection logic, priority criteria, filtering tips
2. **Executive_Summary** — Key findings, Qatar constraints, top recommendations
3. **SA_Recommendations** — Full resource inventory with all SA_ guidance (main working sheet)
4. **QuickWins** — Low-effort, high-value actions
5. **P1_Critical_Actions** — Suggested P1 (Critical) resources — customer to confirm
6. **P2_Actions** — Suggested P2 (High) resources — customer to confirm 🆕
7. **P3_Actions** — Suggested P3 (Medium/Dev-Test) resources — customer to confirm 🆕
8. **Summary_ByResourceType** — Aggregated DR gap analysis by Azure service type
9. **Summary_BySubscription** — Subscription-level DR readiness overview
10. **Timeline_ActionPlan** — Phased implementation roadmap (tentative guidance)
11. **DR_Testing_Plan** — Quarterly DR testing template with rollback procedures
12. **Dependencies_Matrix** — Service dependency mapping from real environment data
13. **Compliance_Checklist** — Qatar PDPPL & NIA/NCSA compliance tracker with DPO sign-off
14. **BCDR_Strategy_Reference** — Strategy design patterns, Qatar constraints, decision matrix
15. **VM_Zone_Planner** — Zone transition planner for VMs/AKS/DBs with legend row and color coding
16. **Risk_Heatmap** — Subscription × Zone Status heatmap with normalized risk scores (0–100)

### ✅ **Microsoft-Branded Interactive HTML Dashboard**
- Zone redundancy distribution by subscription
- Cross-region DR status (4-category breakdown table with ∑ Total)
- 9 charts: IaaS/PaaS split, NonZonal breakdown by resource type, Subscription Risk Scores, and more
- Clickable BCDR report card linking to Excel report

### ✅ **Transparent Priority & Criticality Methodology**
- **2-Step Process:** Criticality Assessment → Priority Assignment
- All labels include **(Suggested — Confirm with Customer)** disclaimer
- Full decision tree documented in the Introduction sheet of the Excel report

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│  Start-AzureBCDRAssessment.ps1 (Main Orchestrator)                 │
└────────────────┬────────────────────────────────────────────────────┘
                 │
                 ├─► PHASE 1: Phase1-CollectResources.ps1
                 │   ├─ Connect to Azure subscriptions (via Az module)
                 │   ├─ Query Azure Resource Graph (all resource types)
                 │   ├─ Analyse zone redundancy status per resource
                 │   ├─ Map logical → physical availability zones
                 │   ├─ Assess cross-region DR readiness (4 categories)
                 │   ├─ Extract resource dependencies (VNet, KeyVault, Storage)
                 │   ├─ Export per-resource-type CSV files
                 │   ├─ Generate master inventory CSV
                 │   └─ Build Microsoft-branded HTML dashboard (9 charts)
                 │
                 └─► PHASE 2: Phase2-AddRecommendations.ps1
                     ├─ Load Phase 1 CSV output
                     ├─ Smart workload tier classification (4-level hierarchy)
                     ├─ Apply service-specific DR knowledge base (54 resource types)
                     ├─ Criticality & priority classification (2-step process)
                     ├─ Generate advisory SA_ columns (all prefixed with Suggested/Estimated)
                     ├─ Build 16-sheet Excel workbook
                     ├─ Update HTML dashboard with clickable Excel report card
                     └─ Open HTML dashboard in browser
```

---

## 📦 Prerequisites

### **Required PowerShell Modules**
```powershell
# Azure PowerShell modules
Install-Module -Name Az.Accounts      -Scope CurrentUser -Force
Install-Module -Name Az.ResourceGraph -Scope CurrentUser -Force

# Excel export module (required for Phase 2)
Install-Module -Name ImportExcel      -Scope CurrentUser -Force
```

### **Azure Permissions**
- **Minimum:** `Reader` role on all target subscriptions
- **Recommended:** `Reader` + `Tag Contributor`

### **PowerShell Version**
- PowerShell 7.0 or higher (`pwsh`)

---

## ⚡ Quick Start

### **Single-Command Execution (Recommended)**
```powershell
.\Start-AzureBCDRAssessment.ps1 -CustomerName "Your Organization" -TenantId "your-tenant-id"
```

This orchestrates both phases automatically:
1. ✅ Authenticate to Azure
2. ✅ Run Phase 1 — resource collection, zone analysis, HTML dashboard
3. ✅ Run Phase 2 — BCDR enrichment, 16-sheet Excel report
4. ✅ Update HTML dashboard with clickable report card
5. ✅ Open HTML dashboard in browser

### **Run Phases Individually**
```powershell
# Phase 1 only
.\Phase1-CollectResources.ps1 -CustomerName "Your Organization" -TenantId "your-tenant-id"

# Phase 2 only (requires Phase 1 output folder path)
.\Phase2-AddRecommendations.ps1 -OutputPath ".\CompleteBCDRAssessment_YYYYMMDD_HHMMSS"
```

### **Output Folder Structure**
```
CompleteBCDRAssessment_YYYYMMDD_HHMMSS\
├── [Customer Name] - BCDR Assessment Report.xlsx   ← 16-sheet Excel
├── Dashboard_ZoneRedundancy.html                   ← HTML dashboard
├── MasterReport_AllResources.csv
├── Summary_ZoneRedundancy.csv
├── Summary_BySubscription.csv
├── ZoneMappings_AllSubscriptions.csv
├── ResourceTypes\                                  ← Per-resource-type CSVs
└── Assessment_Log.txt
```

---

## 📊 Output Reports

### **SA_Recommendations — Main Working Sheet**

Full resource inventory enriched with advisory SA_ columns:

| Column Group | Columns |
|---|---|
| **Identity** | SubscriptionName, ResourceGroup, ResourceName, ResourceType, Location, SKU |
| **Assessment** | ZoneRedundancyStatus, CrossRegionStatus, GeoRedundant, DRReady, LogicalZones, PhysicalZones |
| **Workload Tier** | SA_WorkloadTier *(Suggested)*, SA_TierConfidence, SA_TierSource |
| **SA Recommendations** | SA_Criticality, SA_Priority, SA_BCDRStrategy, SA_DRRegionChoice, SA_DRMethod, SA_RPO, SA_RTO |
| **SA Actions** | SA_ActionRequired *(⚠️ ADVISORY preamble)*, SA_ImplementationEffort *(Estimated)*, SA_CostImpact *(Estimated)* |
| **SA Details** | SA_QuickWin, SA_Dependencies, SA_ComplianceNote, SA_PhysicalZonePlacement, SA_ZoneTransitionPath |
| **Resource Tags** | All Azure tags as Tag_* columns |

**Row Color-Coding by SA_Criticality:**
- 🔴 Red: Critical  |  🟡 Orange-Yellow: High  |  🟡 Light Yellow: Medium  |  ⚫ Gray: Not Assessed (Dev/Test)

### **VM_Zone_Planner Sheet**
- Filtered to: VMs, AKS, Flexible Servers, Application Gateways, Managed Instances
- **Row 2:** Yellow legend row explaining why `LogicalZones`/`PhysicalZones` are empty in Qatar Central (expected — AZs are not customer-accessible)
- Color-coded rows: 🔴 Red = NonZonal, 🟡 Yellow = Zonal, 🟢 Green = ZoneRedundant
- Includes: SA_PhysicalZonePlacement and SA_ZoneTransitionPath with step-by-step guidance

### **Risk_Heatmap Sheet**
- Matrix: Subscriptions (rows) × Zone Redundancy Status (columns)
- Normalized risk score per subscription (0–100)
- Color gradient: 🟢 Green (low risk) → 🔴 Red (high risk)

---

## 🔍 Assessment Methodology

### **Step 1 — Criticality Assessment**

| Condition | Result |
|---|---|
| Subscription contains: `dev`, `test`, `tst`, `poc`, `sandbox`, `uat`, `qa`, `stage` | **Not Assessed (Dev/Test)** |
| Name contains: `prod`, `prd`, `production`, `live` + Critical resource type* | **Critical (Suggested)** |
| Name contains: `prod`, `prd`, `production`, `live` + other types | **High (Suggested)** |
| No prod indicator + Critical resource type* | **Medium (Suggested)** |
| No indicators + non-critical type | **Low (Suggested)** |

*Critical resource types: VMs, AKS, SQL, MySQL, PostgreSQL, Cosmos DB, Storage, Key Vault, NetApp, Recovery Services Vaults, API Management, Event Hubs, Service Bus, Redis Enterprise, Data Factory*

### **Step 2 — Priority Assignment**

| Criticality | Zone Redundancy | Final Priority |
|---|---|---|
| Critical | Any | **P1 (Suggested — Confirm with Customer)** |
| High | Any | **P2 (Suggested — Confirm with Customer)** |
| Medium | Non-Zonal or Locally Redundant | **P2 (Suggested — Confirm with Customer)** — gap detected |
| Medium | Zone Redundant or Zonal | **P3 (Suggested — Confirm with Customer)** |
| Low | Any | **P4 (Suggested — Confirm with Customer)** |
| Not Assessed | Any | **Not Prioritised (Dev/Test — Customer to Assess)** |

> ⚠️ All priorities are **suggestions** based on automated naming analysis only. Customer must validate with business stakeholders before taking any action.

### **Workload Tier Auto-Detection**

| Tier | DR Strategy Applied | Cost Impact |
|---|---|---|
| (Suggested) Production | Active-Passive / Active-Active | High |
| (Suggested) Non-Production | Backup & Restore | Medium |
| (Suggested) Dev/Test | Backup & Restore (Cost-Optimized), 24h RPO | Low |
| (Suggested) Sandbox | No DR Required (Recreate from IaC) | **Zero** |
| (Suggested) Production (Assumed) | Production strategy + ⚠️ warning to confirm | High until confirmed |

---

## 🇶🇦 Qatar-Specific Compliance

### **Key Qatar Central Constraints**
| Constraint | Detail |
|---|---|
| No paired region | GRS storage unavailable; Azure Backup CRR unavailable; manual DR region selection required |
| Zone redundancy restricted | One AZ at full capacity; ZRS/GZRS blocked for most services |
| Azure Backup RoC (preview) | Supported workloads: IaaS VM, SQL in VM, SAP HANA, Azure Files only; Target regions: Sweden Central or Switzerland North ONLY |
| IaC parity critical | *"IaC parity is the MOST CRITICAL FACTOR for RTO in non-paired regions"* — Microsoft Azure Qatar BCDR Plan |

### **NIA/NCSA Certification**
- **Certificate ID:** 10018 | **Valid Until:** 28 August 2026
- **Certified regions:** Qatar Central ✅, West Europe ✅, North Europe ✅
- Sweden Central and Switzerland North (Azure Backup RoC targets) are **NOT NIA-certified** — document in compliance checklist

### **Qatar PDPPL**
- Cross-border data replication requires **DPO approval**
- Conduct data classification workshop before enabling geo-replication
- Compliance checklist sheet tracks DPO sign-off and PDPPL requirements

---

## 🛡️ BCDR Strategies Recommended

| Strategy | RTO | RPO | Cost | Use Case |
|---|---|---|---|---|
| **Active-Active** | Near-zero | Near-zero | High | Mission-critical (Tier 1) |
| **Active-Passive (Warm)** | < 30 min | < 1 hour | Medium | Production workloads |
| **Active-Passive (Cold / IaC)** | Hours | Hours–Days | Low | Cost-sensitive production |
| **Backup & Restore** | 4–12 hours | 24 hours | Low | Non-critical / Dev/Test |
| **Geo-Replication** | Automatic | < 15 sec | Premium tier | Storage, Cosmos DB, Service Bus Premium |
| **No DR (IaC Recreate)** | Hours | N/A | Zero | Sandbox / POC |

---

## 🔧 Supported Azure Services (54 Resource Types)

**Compute:** VMs, AKS, App Service, Functions, Container Apps, VMSS, AVS  
**Databases:** Azure SQL, Managed Instance, MySQL Flexible, PostgreSQL Flexible, Cosmos DB  
**Storage:** Storage Accounts, Azure NetApp Files, Managed Disks  
**Networking:** VNets, NSGs, Public IPs, App Gateway, Load Balancer, Front Door, Traffic Manager, Firewall, Bastion, VPN/ExpressRoute Gateways, Private Endpoints, Private DNS  
**Security:** Key Vault, Managed Identities  
**Integration:** Logic Apps, Event Hubs, Service Bus, Event Grid, API Management  
**Management:** Recovery Services Vaults, Log Analytics, Application Insights, Automation Accounts  
**Containers:** Azure Container Registry  
**AI/ML:** Cognitive Services, Machine Learning, AI Search  

---

## ⚠️ Limitations & Disclaimers

1. **All SA_ outputs are SUGGESTIONS.** Customer decides — tool advises.
2. **Workload tier detection is based on naming patterns only.** Filter `SA_TierConfidence = "Low"` to review uncertain classifications.
3. **LogicalZones / PhysicalZones empty for Qatar Central** — expected behavior. Qatar Central's AZs are not customer-accessible (one AZ at full capacity). See the VM_Zone_Planner legend row.
4. **Azure Backup RoC is a preview feature.** Confirm current availability with your Microsoft account team before designing DR around it.
5. **Zone redundancy alone is insufficient.** Zones protect against single-AZ failures only. Cross-region DR is always required for business-critical workloads.
6. **Qatar Central has no paired region.** DR region selection is the customer's decision. The tool references West Europe as a common pattern but the customer must confirm.

---

## 📚 Version History

### **Version 2.0** — April 2026

**New Features:**
- ✅ Microsoft-branded HTML dashboard (full design system overhaul)
- ✅ 9 charts in dashboard (3 new: NonZonal breakdown, IaaS/PaaS/Platform, Subscription Risk Score)
- ✅ Cross-region card redesigned as 4-category mutually exclusive breakdown table
- ✅ P2_Actions and P3_Actions sheets added (16 sheets total, was 12)
- ✅ VM_Zone_Planner yellow legend row explaining empty LogicalZones/PhysicalZones
- ✅ Risk_Heatmap sheet (Subscription × Zone Status, 0–100 normalized risk score)
- ✅ All SA_ columns now prefixed with `(Suggested)` / `(Estimated)` / `(Customer to confirm)`
- ✅ `SA_ActionRequired` includes `⚠️ ADVISORY` preamble on every row
- ✅ `SA_WorkloadTier` prefixed with `(Suggested)` — color-coded with 5 distinct tier colors
- ✅ Advisory language throughout — no dictating, no mandatory statements
- ✅ P2 and P3 counts shown in console output at completion

### **Version 1.0** — April 2026

**Initial Release:**
- ✅ 2-phase automated assessment framework
- ✅ Single-command execution wrapper
- ✅ 12-sheet Excel report with Qatar NIA/NCSA compliance
- ✅ Smart workload tier classification (Production / Non-Prod / Dev-Test / Sandbox)
- ✅ Priority & criticality methodology (2-step process)
- ✅ Interactive HTML dashboard
- ✅ Dependency mapping from live environment data
- ✅ Compliance checklist with DPO sign-off tracking
- ✅ Azure Backup Region of Choice (RoC) preview support

---

## 📧 Support & Contact

- **Microsoft Account Team** — Engage your Microsoft Solution Architect or Customer Success Manager
- **Automation Developer** — Zahir Hussain Shah, Sr. Solution Engineer, Microsoft Qatar

---

## 📄 License

© Microsoft Corporation. Internal use only.

---

## 🙏 Acknowledgments

- **Microsoft Azure Qatar BCDR Plan** — Service-specific DR strategies and Qatar Central constraints
- **Microsoft Azure Backup Region of Choice (RoC)** — Engineering Team Content
- **Azure Well-Architected Framework** — BCDR best practices
- **NIA/NCSA Certification** — Qatar compliance requirements (Certificate ID: 10018)

---

**Build Version:** 2.0 | **Last Updated:** April 2026 | **Framework:** PowerShell 7 + Azure PowerShell + ImportExcel 7.x
