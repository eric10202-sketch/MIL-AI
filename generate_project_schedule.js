/**
 * Project Trinity - IT Carve-Out Project Schedule Generator
 * Generates a CSV file importable into Microsoft Project.
 *
 * Usage:  node generate_project_schedule.js
 * Output: Trinity_Project_Schedule.csv
 *
 * Context:
 *   Buyer  : Bosch (acquiring part of Johnson Controls International)
 *   Seller : JCI (performing full IT carve-out)
 *   Period : 01 Jul 2026 – 31 Dec 2029
 *   TSA    : 18 months (JCI operates IT for Bosch after Day 1)
 *   Interim: Merger Zone (temporary Bosch environment for all JCI IT)
 *   Scope  : 180 sites (AP largest, AM largest), ~8,000 employees,
 *            ~6,000 client devices, ~500 applications, FRAME-comparable ERP
 *
 * MS Project import columns:
 *   ID | Outline Level | Name | Duration | Start | Finish |
 *   Predecessors | Resource Names | Notes | Milestone
 */

const fs = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// Task definition helper
// ---------------------------------------------------------------------------
function t(id, level, name, duration, start, finish, predecessors, resources, notes, milestone) {
  return { id, level, name, duration, start, finish, predecessors, resources, notes, milestone };
}

// ---------------------------------------------------------------------------
// Task list
// ---------------------------------------------------------------------------
const tasks = [

  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 1 – INITIALIZATION  (01 Jul 2026 → 30 Sep 2026)
  // ═══════════════════════════════════════════════════════════════════════════
  t(1,  1, 'Phase 1 - Initialization',                                   '65 days',  '07/01/26', '09/30/26', '',          'Bosch IT + JCI IT + KPMG',         '', 'No'),
  t(2,  2, '1.1 Project Setup and Governance',                           '20 days',  '07/01/26', '07/28/26', '',          'Bosch IT PM + JCI IT PM',          '', 'No'),
  t(3,  3, 'Appoint IT Project Leads (Bosch and JCI)',                   '10 days',  '07/01/26', '07/14/26', '',          'Bosch CIO + JCI CIO',              '', 'No'),
  t(4,  3, 'Establish Steering Committee',                               '10 days',  '07/07/26', '07/20/26', '3',         'Bosch + JCI Senior Mgmt',          '', 'No'),
  t(5,  3, 'Define RACI and Governance Model',                           '10 days',  '07/07/26', '07/20/26', '3',         'KPMG + IT PM',                     '', 'No'),
  t(6,  3, 'Set Up PMO Tools and SharePoint Collaboration Site',         '15 days',  '07/14/26', '08/03/26', '3',         'PMO + KPMG',                       '', 'No'),
  t(7,  3, 'Onboard External Consulting Partner',                        '20 days',  '07/14/26', '08/10/26', '3',         'Procurement',                      '', 'No'),

  t(8,  2, '1.2 Initial IT Landscape Assessment',                        '35 days',  '07/21/26', '09/08/26', '3',         'JCI IT + Bosch IT',                '180 sites across AP/AM/EMEA; AP and AM are the largest regions', 'No'),
  t(9,  3, 'High-Level IT Inventory Request to JCI',                     '10 days',  '07/21/26', '08/03/26', '5',         'Bosch IT PM',                      '', 'No'),
  t(10, 3, 'Location List - 180 Sites (AP / AM / EMEA)',                 '20 days',  '07/28/26', '08/24/26', '9',         'JCI IT + Facilities',              'AP and AM sites prioritised; collect full addresses and site types', 'No'),
  t(11, 3, 'Legal Entity Identification (All Regions)',                  '20 days',  '08/03/26', '08/28/26', '9',         'Legal + Finance',                  '', 'No'),
  t(12, 3, 'Identify Key IT Dependencies and JCI Shared Services',       '25 days',  '08/10/26', '09/11/26', '10',        'JCI IT Architects',                '', 'No'),

  t(13, 2, '1.3 Project Charter and Scope Definition',                   '25 days',  '08/17/26', '09/18/26', '8',         'Bosch IT PM + KPMG',               '', 'No'),
  t(14, 3, 'Draft Project Charter',                                      '15 days',  '08/17/26', '09/04/26', '8',         'KPMG + Bosch IT PM',               '', 'No'),
  t(15, 3, 'Define Initial Scope Boundaries',                            '15 days',  '08/24/26', '09/11/26', '10',        'KPMG + JCI IT',                    '~8000 employees; ~6000 IT client devices; FRAME-comparable IT/ERP landscape', 'No'),
  t(16, 3, 'Guiding Principles Workshop (Bosch and JCI)',                '5 days',   '09/01/26', '09/05/26', '14',        'All WS Leads',                     '', 'No'),
  t(17, 3, 'TSA Framework Preliminary Definition (18-Month TSA)',        '10 days',  '09/01/26', '09/12/26', '14',        'Legal + IT PM',                    'JCI operates all IT for Bosch during 18-month TSA period', 'No'),
  t(18, 3, 'Project Charter Approval',                                   '5 days',   '09/15/26', '09/19/26', '14,15,16,17', 'Steering Committee',             '', 'Yes'),
  t(19, 2, 'QG1 - Initialization Quality Gate',                         '0 days',   '09/30/26', '09/30/26', '18',        'Steering Committee',               'All deliverables signed off; project charter approved; proceed to Concept', 'Yes'),

  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 2 – CONCEPT  (01 Oct 2026 → 31 Mar 2027)
  // ═══════════════════════════════════════════════════════════════════════════
  t(20, 1, 'Phase 2 - Concept',                                          '130 days', '10/01/26', '03/31/27', '19',        'All Workstreams + KPMG',           '', 'No'),

  t(21, 2, '2.1 Detailed As-Is Analysis',                                '60 days',  '10/01/26', '12/23/26', '19',        'JCI IT + KPMG',                    '', 'No'),
  t(22, 3, 'Application Landscape Inventory (LeanIX / CMDB) ~500 Apps',  '40 days',  '10/01/26', '11/26/26', '19',        'JCI IT Architects + KPMG',         'Comparable to FRAME; use LeanIX as master landscape tool', 'No'),
  t(23, 3, 'IT Infrastructure As-Is Mapping (180 Sites)',                '40 days',  '10/01/26', '11/26/26', '19',        'JCI Infra Team',                   'Site visits for top 20 locations; AP and AM hubs prioritised', 'No'),
  t(24, 3, 'ERP Landscape Analysis (SAP and Non-SAP)',                   '40 days',  '10/01/26', '11/26/26', '19',        'JCI ERP Team + KPMG',              'FRAME-comparable SAP landscape', 'No'),
  t(25, 3, 'Data Ownership and Separation Rules',                        '30 days',  '11/02/26', '12/11/26', '22,24',     'KPMG + Legal',                     'Non-selective migration approach preferred (ref. FRAME)', 'No'),
  t(26, 3, 'Contract and License Inventory',                             '40 days',  '10/01/26', '11/26/26', '19',        'JCI Procurement + IT',             'Change of control clauses; ~6000 device licenses', 'No'),
  t(27, 3, 'HR IT Systems Mapping (8000 Employees)',                     '30 days',  '10/01/26', '11/11/26', '19',        'JCI HR IT',                        'Country-by-country HR rules required', 'No'),
  t(28, 3, 'OT and Production IT Assessment',                            '30 days',  '10/15/26', '11/25/26', '19',        'JCI OT + Production IT',           '', 'No'),
  t(29, 3, 'Country-Specific IT Requirements (Brazil China India Mexico)','25 days', '10/19/26', '11/20/26', '19',        'Regional IT Leads',                'Brazil ERP tax complexity; China local FTS; India RBIN', 'No'),

  t(30, 2, '2.2 Merger Zone Architecture Design',                        '60 days',  '12/01/26', '02/19/27', '22,23',     'Bosch IT Architects + KPMG',       '', 'No'),
  t(31, 3, 'Merger Zone Concept and Design Workshop',                    '10 days',  '12/01/26', '12/14/26', '22',        'Bosch + JCI Architects',           'Temporary environment to receive all JCI IT services and applications', 'No'),
  t(32, 3, 'WAN and LAN Architecture Design (180 Sites)',                '20 days',  '12/14/26', '01/11/27', '31',        'Bosch Infra + KPMG',               'Regional hub model: AP hub + AM hub + EMEA hub', 'No'),
  t(33, 3, 'Active Directory Design (Merger Zone)',                      '20 days',  '12/14/26', '01/11/27', '31',        'Bosch AD Team',                    '', 'No'),
  t(34, 3, 'M365 and Azure Tenant Design',                               '20 days',  '12/14/26', '01/11/27', '31',        'Bosch Azure Team',                 '6000+ mailboxes and OneDrive migration scope', 'No'),
  t(35, 3, 'Co-Locator and Data Center Selection (per Region)',          '30 days',  '12/01/26', '01/13/27', '22',        'Bosch Procurement + IT',           'Recommend 3 regional co-locators: AP / AM / EMEA', 'No'),
  t(36, 3, 'Security Architecture Design',                               '25 days',  '01/11/27', '02/13/27', '32,33',     'Bosch CISO',                       '', 'No'),
  t(37, 3, 'IAM and IdM Architecture Design',                            '20 days',  '01/11/27', '02/05/27', '33',        'Bosch IAM Team',                   '', 'No'),

  t(38, 2, '2.3 Migration Strategy',                                     '45 days',  '12/14/26', '02/12/27', '22,24',     'KPMG + All WS Leads',              '', 'No'),
  t(39, 3, 'Carve-Out Model Selection (Stand Alone)',                    '10 days',  '12/14/26', '12/25/26', '22',        'Bosch + JCI Management',           'JCI full carve-out; Merger Zone as interim state before Bosch integration', 'No'),
  t(40, 3, 'ERP Migration Strategy (Relocation vs Split vs Greenfield)', '20 days',  '12/28/26', '01/22/27', '39',        'ERP Architects + KPMG',            '', 'No'),
  t(41, 3, 'Application Categorization and Wave Planning (~500 Apps)',   '25 days',  '01/04/27', '02/05/27', '22,39',     'KPMG + WS Leads',                  'Sort: keep/replace/no-action; assign to migration waves', 'No'),
  t(42, 3, 'Data Separation Rules Finalized',                            '15 days',  '01/11/27', '01/29/27', '25',        'KPMG + Legal',                     '', 'No'),
  t(43, 3, 'TSA Service Catalog Definition (18-Month JCI Services)',     '25 days',  '01/11/27', '02/12/27', '39',        'JCI IT + Bosch IT',                'Catalog all JCI services; define exit criteria per service', 'No'),
  t(44, 3, 'Merger Zone to Bosch Migration Strategy',                    '20 days',  '01/25/27', '02/19/27', '40,41',     'Bosch IT Architects',              'Post-TSA migration roadmap', 'No'),

  t(45, 2, '2.4 Client Device and Workplace Strategy',                   '20 days',  '01/11/27', '02/05/27', '22',        'Bosch CWP Team',                   '', 'No'),
  t(46, 3, 'Client Device Inventory (~6000 Devices)',                    '10 days',  '01/11/27', '01/22/27', '22',        'JCI IT + Asset Mgmt',              'Identify device age; out-of-support hardware replacement plan', 'No'),
  t(47, 3, 'Client Migration Approach (Reimaging vs In-Place)',          '10 days',  '01/23/27', '02/05/27', '46',        'Bosch CWP Team',                   '~6000 clients across 180 sites; wave-based by region', 'No'),

  t(48, 2, '2.5 Planning and Baseline',                                  '25 days',  '02/22/27', '03/26/27', '38,30',     'KPMG + IT PM',                     '', 'No'),
  t(49, 3, 'Detailed Project Plan Development',                          '20 days',  '02/22/27', '03/19/27', '38,30',     'KPMG + IT PM',                     '', 'No'),
  t(50, 3, 'Risk Register Baseline',                                     '15 days',  '02/22/27', '03/12/27', '38',        'KPMG + All WS Leads',              '', 'No'),
  t(51, 3, 'OPL Establishment',                                          '10 days',  '03/01/27', '03/12/27', '49',        'PMO',                              '', 'No'),
  t(52, 3, 'Resource Plan and Budget',                                   '15 days',  '03/02/27', '03/20/27', '49',        'Finance + IT PM',                  '', 'No'),
  t(53, 2, 'QG2 - Concept Quality Gate',                                 '0 days',   '03/31/27', '03/31/27', '48,44,43',  'Steering Committee',               'Architecture approved; migration strategy confirmed; TSA catalog accepted; proceed to Development', 'Yes'),

  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 3 – DEVELOPMENT  (01 Apr 2027 → 28 Nov 2027)
  // ═══════════════════════════════════════════════════════════════════════════
  t(54, 1, 'Phase 3 - Development',                                      '175 days', '04/01/27', '11/28/27', '53',        'All Workstreams',                  '', 'No'),

  t(55, 2, '3.1 Merger Zone Infrastructure Build',                       '120 days', '04/01/27', '09/16/27', '53',        'Bosch Infra + Partners',           '', 'No'),
  t(56, 3, 'WAN Provider Selection and Ordering (180 Sites)',            '20 days',  '04/01/27', '04/28/27', '53',        'Bosch Procurement',                'Order immediately after QG2 - 4-6 month lead time for all sites', 'No'),
  t(57, 3, 'Co-Locator Data Center Procurement (3 Regions)',             '40 days',  '04/01/27', '05/28/27', '53',        'Bosch Infra',                      'AP + AM + EMEA co-locators', 'No'),
  t(58, 3, 'Active Directory Build (Merger Zone)',                       '30 days',  '04/29/27', '06/10/27', '56',        'Bosch AD Team',                    'Foundation for 6000+ client and application migrations', 'No'),
  t(59, 3, 'WAN Implementation - EMEA Sites',                            '40 days',  '05/17/27', '07/09/27', '56',        'WAN Provider + Bosch Infra',       '', 'No'),
  t(60, 3, 'WAN Implementation - AP Sites',                              '50 days',  '05/17/27', '07/23/27', '56',        'WAN Provider + Bosch Infra',       'Largest region; stagger by country', 'No'),
  t(61, 3, 'WAN Implementation - AM Sites',                              '50 days',  '05/17/27', '07/23/27', '56',        'WAN Provider + Bosch Infra',       'Largest region; US hubs first', 'No'),
  t(62, 3, 'LAN and WiFi Setup (Top 30 Sites)',                          '40 days',  '06/14/27', '08/06/27', '58',        'Bosch Infra',                      'Remaining 150 sites use existing JCI LAN under TSA', 'No'),
  t(63, 3, 'Server Infrastructure Build (Merger Zone - 3 Regional DCs)','40 days',  '05/17/27', '07/09/27', '57',        'Bosch Infra',                      '', 'No'),
  t(64, 3, 'M365 Tenant Provisioning (6000+ Mailboxes)',                 '30 days',  '05/17/27', '06/27/27', '53',        'Bosch Azure Team',                 '', 'No'),
  t(65, 3, 'Azure and Cloud Environment Setup',                          '30 days',  '06/28/27', '08/08/27', '64',        'Bosch Cloud Team',                 '', 'No'),
  t(66, 3, 'Telephony Setup Merger Zone (180 Sites)',                    '20 days',  '06/14/27', '07/10/27', '59,60,61',  'Bosch Telecom',                    'Phone number porting 4-8 weeks per site per operator', 'No'),
  t(67, 3, 'File Share Infrastructure (NAS / Cloud)',                    '20 days',  '06/14/27', '07/10/27', '63',        'Bosch Infra',                      'User documents for ~8000 employees', 'No'),
  t(68, 3, 'Backup and Archiving Setup',                                 '15 days',  '07/12/27', '07/30/27', '63',        'Bosch Infra',                      '', 'No'),

  t(69, 2, '3.2 ERP and SAP Development',                                '110 days', '04/01/27', '09/02/27', '53',        'ERP Team + Migration Partner',     '', 'No'),
  t(70, 3, 'SAP System Landscape Analysis',                              '20 days',  '04/01/27', '04/28/27', '53',        'ERP Architects',                   '', 'No'),
  t(71, 3, 'SAP Shell Copy Preparation (Merger Zone)',                   '40 days',  '04/29/27', '06/25/27', '70',        'Migration Partner + Bosch SAP',    '', 'No'),
  t(72, 3, 'SAP Interface Identification and Adaptation Plan',           '30 days',  '04/29/27', '06/11/27', '70',        'ERP Team',                         '', 'No'),
  t(73, 3, 'SAP Authorization Concept (Disconnect from JCI IdM)',        '25 days',  '06/12/27', '07/16/27', '71',        'Bosch SAP Security',               '', 'No'),
  t(74, 3, 'BW and Reporting Setup',                                     '30 days',  '06/26/27', '08/07/27', '71',        'Bosch BW Team',                    '', 'No'),
  t(75, 3, 'EDI and Integration Setup',                                  '30 days',  '06/26/27', '08/07/27', '72',        'Integration Team',                 '', 'No'),
  t(76, 3, 'ERP Test Environment Setup (SIT and UAT)',                   '20 days',  '07/17/27', '08/13/27', '71',        'Migration Partner + Bosch SAP',    '', 'No'),

  t(77, 2, '3.3 Application Migration Preparation (~500 Apps)',          '90 days',  '04/01/27', '08/06/27', '53',        'SWS3 + KPMG',                      '', 'No'),
  t(78, 3, 'Application Inventory Finalization',                         '20 days',  '04/01/27', '04/28/27', '53',        'KPMG + JCI IT',                    'Validate LeanIX entries; add local/regional apps', 'No'),
  t(79, 3, 'Application Grouping into Migration Waves',                  '15 days',  '04/29/27', '05/19/27', '78',        'KPMG + WS Leads',                  'Wave 1: critical; Wave 2: standard; Wave 3: remaining', 'No'),
  t(80, 3, 'Migration Package Development per Wave',                     '40 days',  '05/20/27', '07/14/27', '79',        'SWS3 + App Owners',                '', 'No'),
  t(81, 3, 'Custom Application Adaptation for Merger Zone',              '30 days',  '06/14/27', '07/25/27', '79',        'Dev Teams',                        '', 'No'),
  t(82, 3, 'Integration Development (APIs and Interfaces)',               '30 days',  '06/14/27', '07/25/27', '79',        'Integration Team',                 '', 'No'),

  t(83, 2, '3.4 Client Workplace - 6000 Device Migration Prep',          '60 days',  '04/01/27', '06/25/27', '53',        'Bosch CWP Team',                   '', 'No'),
  t(84, 3, 'Client Software Management Setup (SCCM / Intune)',           '20 days',  '04/01/27', '04/28/27', '53',        'Bosch CWP',                        '', 'No'),
  t(85, 3, 'Device Reimaging Standard and Packaging (~6000 Devices)',    '30 days',  '04/29/27', '06/11/27', '84',        'Bosch CWP',                        '', 'No'),
  t(86, 3, 'Regional Wave Plan - Client Migration (AP AM EMEA)',         '15 days',  '05/01/27', '05/21/27', '84',        'Bosch CWP + IT PM',                'AP Wave A+B; AM Wave A+B; EMEA Wave A', 'No'),
  t(87, 3, 'User Profile and OneDrive Migration Tooling Setup',          '20 days',  '05/22/27', '06/18/27', '85',        'Bosch Azure Team',                 '', 'No'),

  t(88, 2, '3.5 Security and IAM',                                       '80 days',  '04/01/27', '07/19/27', '53',        'Bosch CISO + IAM Team',            '', 'No'),
  t(89, 3, 'IAM Solution Deployment (Identity Provider)',                 '40 days',  '04/01/27', '05/28/27', '53',        'Bosch IAM',                        '', 'No'),
  t(90, 3, 'IT Security Concept Finalization (incl. External Assessment)','20 days', '05/29/27', '06/25/27', '89',        'Bosch CISO',                       '', 'No'),
  t(91, 3, 'OT Security Assessment and Remediation',                     '30 days',  '05/01/27', '06/11/27', '53',        'OT Security Team',                 '', 'No'),
  t(92, 3, 'CMDB Setup (ServiceNow) - 6000 Assets',                     '30 days',  '05/29/27', '07/09/27', '89',        'Bosch ITSM Team',                  '', 'No'),

  t(93,  2, '3.6 IT Organization and Operations Setup',                  '80 days',  '04/01/27', '07/19/27', '53',        'Bosch IT Org Team',                '', 'No'),
  t(94,  3, 'Bosch IT Org Structure (8000 Employees / 180 Sites)',       '20 days',  '04/01/27', '04/28/27', '53',        'Bosch HR + IT',                    'Regional IT leads for AP / AM / EMEA', 'No'),
  t(95,  3, 'MSP and ITO Contracting (Field Support 180 Sites)',         '40 days',  '04/29/27', '06/25/27', '94',        'Bosch Procurement',                'On-site support for AP and AM sites', 'No'),
  t(96,  3, 'ITSM and ServiceNow Configuration',                         '30 days',  '05/29/27', '07/09/27', '92',        'Bosch ITSM',                       '', 'No'),
  t(97,  3, 'Help Desk Setup and Staffing (Multi-Region 24x5)',          '20 days',  '06/26/27', '07/23/27', '95',        'Bosch IT Ops',                     'Multilingual support for AP / AM / EMEA time zones', 'No'),

  t(98,  2, '3.7 TSA Framework Finalization',                            '60 days',  '04/01/27', '06/25/27', '53',        'JCI IT + Legal',                   '', 'No'),
  t(99,  3, 'TSA Service Descriptions Finalized (All 180 Sites)',        '30 days',  '04/01/27', '05/13/27', '53',        'JCI IT + Bosch IT',                '', 'No'),
  t(100, 3, 'TSA SLA and KPI Framework',                                 '20 days',  '05/14/27', '06/10/27', '99',        'IT PM + Legal',                    '', 'No'),
  t(101, 3, 'TSA Governance Model (Monthly Reviews and Reporting)',      '15 days',  '06/11/27', '07/01/27', '100',       'IT PM',                            '', 'No'),
  t(102, 3, 'TSA Contracts Legal Review and Finalization',               '20 days',  '06/11/27', '07/08/27', '100',       'Legal',                            '', 'No'),

  t(103, 2, 'QG3 - Development Quality Gate',                            '0 days',   '11/28/27', '11/28/27', '55,69,77,83,88,93,98', 'Steering Committee', 'Merger Zone ready; ERP built; app packages ready; 6000 devices packaged; TSA contracts signed', 'Yes'),

  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 4 – IMPLEMENTATION  (01 Dec 2027 → 30 Jun 2028)
  // ═══════════════════════════════════════════════════════════════════════════
  t(104, 1, 'Phase 4 - Implementation',                                  '152 days', '12/01/27', '06/30/28', '103',       'All Workstreams',                  '', 'No'),

  t(105, 2, '4.1 Infrastructure Migration to Merger Zone (180 Sites)',   '90 days',  '12/01/27', '04/04/28', '103',       'Bosch Infra + Partners',           '', 'No'),
  t(106, 3, 'WAN Cutover - EMEA Sites Batch 1',                         '20 days',  '12/01/27', '12/28/27', '103',       'WAN Team + JCI Infra',             '', 'No'),
  t(107, 3, 'WAN Cutover - AP Sites Batch 1 (Major Hubs)',               '20 days',  '12/01/27', '12/28/27', '103',       'WAN Team + JCI Infra',             'AP largest region; hub-first approach', 'No'),
  t(108, 3, 'WAN Cutover - AM Sites Batch 1 (Major Hubs)',               '20 days',  '12/01/27', '12/28/27', '103',       'WAN Team + JCI Infra',             '', 'No'),
  t(109, 3, 'WAN Cutover - AP Sites Batch 2 (Remaining)',                '20 days',  '01/04/28', '01/31/28', '107',       'WAN Team + JCI Infra',             '', 'No'),
  t(110, 3, 'WAN Cutover - AM Sites Batch 2 (Remaining)',                '20 days',  '01/04/28', '01/31/28', '108',       'WAN Team + JCI Infra',             '', 'No'),
  t(111, 3, 'WAN Cutover - EMEA Sites Batch 2 (Remaining)',              '20 days',  '01/04/28', '01/31/28', '106',       'WAN Team + JCI Infra',             '', 'No'),
  t(112, 3, 'File Share Migration to Merger Zone (~8000 Users)',         '30 days',  '01/04/28', '02/14/28', '106',       'Bosch Infra',                      'User documents and shared drives', 'No'),
  t(113, 3, 'Telephony Cutover to Merger Zone (180 Sites)',               '40 days',  '01/18/28', '03/14/28', '109,110,111', 'Bosch Telecom',                 'Staggered by site; 4-8 weeks per operator', 'No'),

  t(114, 2, '4.2 Client Device Migration - 6000 Devices',                '100 days', '12/01/27', '04/19/28', '103',       'Bosch CWP + Regional IT',          '', 'No'),
  t(115, 3, 'Client Migration Wave 1 - AP Region Hub Sites',             '30 days',  '12/01/27', '01/12/28', '103',       'AP IT Team',                       '~900 devices (AP hubs)', 'No'),
  t(116, 3, 'Client Migration Wave 2 - AM Region Hub Sites',             '30 days',  '12/01/27', '01/12/28', '103',       'AM IT Team',                       '~900 devices (AM hubs)', 'No'),
  t(117, 3, 'Client Migration Wave 3 - EMEA Sites',                      '25 days',  '01/13/28', '02/16/28', '115',       'EMEA IT Team',                     '~700 devices', 'No'),
  t(118, 3, 'Client Migration Wave 4 - AP Region Remaining Sites',       '30 days',  '01/13/28', '02/23/28', '115',       'AP IT Team',                       '~1500 devices', 'No'),
  t(119, 3, 'Client Migration Wave 5 - AM Region Remaining Sites',       '30 days',  '01/13/28', '02/23/28', '116',       'AM IT Team',                       '~1500 devices', 'No'),
  t(120, 3, 'Client Migration Wave 6 - Stragglers and Exceptions',       '20 days',  '02/26/28', '03/22/28', '117,118,119', 'CWP Team',                      '~500 devices; special handling', 'No'),
  t(121, 3, 'Email and M365 Mailbox Migration (6000+ Mailboxes)',        '45 days',  '12/29/27', '02/29/28', '106',       'Bosch Azure Team',                 'Migrate in batches aligned with client waves', 'No'),
  t(122, 3, 'OneDrive and User Document Migration',                      '30 days',  '01/18/28', '02/29/28', '121',       'Bosch Azure Team',                 '', 'No'),

  t(123, 2, '4.3 ERP Migration to Merger Zone',                          '90 days',  '12/01/27', '04/04/28', '103',       'ERP Team + Migration Partner',     '', 'No'),
  t(124, 3, 'SAP Shell Copy Execution - Test Run',                       '20 days',  '12/01/27', '12/28/27', '103',       'Migration Partner + Bosch SAP',    '', 'No'),
  t(125, 3, 'SAP System Integration Test 1 (SIT1)',                      '25 days',  '01/04/28', '02/07/28', '124',       'ERP Team + Business',              '', 'No'),
  t(126, 3, 'SAP Shell Copy Execution - Production Prep',                '15 days',  '02/08/28', '02/28/28', '125',       'Migration Partner + Bosch SAP',    '', 'No'),
  t(127, 3, 'SAP System Integration Test 2 (SIT2)',                      '25 days',  '03/01/28', '04/04/28', '126',       'ERP Team + Business',              '', 'No'),
  t(128, 3, 'Business User Acceptance Testing (UAT)',                    '20 days',  '03/11/28', '04/05/28', '127',       'Business Key Users',               '', 'No'),
  t(129, 3, 'UAT Sign-Off',                                              '0 days',   '04/05/28', '04/05/28', '128',       'Business Leads + IT PM',           '', 'Yes'),

  t(130, 2, '4.4 Application Migration to Merger Zone',                  '90 days',  '12/01/27', '04/04/28', '103',       'SWS3 + App Teams',                 '', 'No'),
  t(131, 3, 'Wave 1 Application Migration - Critical Business Apps',     '25 days',  '12/01/27', '01/05/28', '103',       'App Teams + KPMG',                 '', 'No'),
  t(132, 3, 'Wave 2 Application Migration - Standard Apps',              '30 days',  '01/08/28', '02/16/28', '131',       'App Teams',                        '', 'No'),
  t(133, 3, 'Wave 3 Application Migration - Remaining and Local Apps',   '30 days',  '02/19/28', '03/29/28', '132',       'App Teams',                        'Regional apps (AP / AM / EMEA) included', 'No'),
  t(134, 3, 'End-to-End Integration Testing',                            '20 days',  '03/01/28', '03/28/28', '132',       'Test Team + Business',             '', 'No'),
  t(135, 3, 'Regression Testing Sign-Off',                               '0 days',   '04/04/28', '04/04/28', '133,134',   'Test Manager',                     '', 'Yes'),

  t(136, 2, '4.5 Signing and Legal Closeout',                            '50 days',  '02/12/28', '04/19/28', '105',       'Legal + Finance',                  '', 'No'),
  t(137, 3, 'Legal Entity Structure Finalization (All Regions)',         '20 days',  '02/12/28', '03/11/28', '105',       'Legal',                            '', 'No'),
  t(138, 3, 'TSA Contracts Finalized and Signed',                        '15 days',  '03/12/28', '04/01/28', '137',       'Legal + JCI',                      '', 'No'),
  t(139, 3, 'License Transfer Agreements Executed (~6000 Devices)',      '15 days',  '03/12/28', '04/01/28', '137',       'Procurement + Legal',              '', 'No'),
  t(140, 3, 'Change of Control Notifications to Vendors',                '10 days',  '04/02/28', '04/15/28', '139',       'Procurement',                      '', 'No'),
  t(141, 3, 'Signing - SPA and APA Execution',                           '0 days',   '04/19/28', '04/19/28', '138,140',   'Executive Leadership',             '', 'Yes'),
  t(142, 3, 'Frozen Zone Begins (Minimize Production Changes)',          '0 days',   '04/19/28', '04/19/28', '141',       'IT PM',                            'No production changes without SteerCo approval from this point', 'Yes'),

  t(143, 2, '4.6 Day 1 Cutover Preparation',                             '40 days',  '04/22/28', '06/14/28', '129,135,141', 'IT PM + All WS',                '', 'No'),
  t(144, 3, 'Cutover Plan Finalization',                                  '15 days',  '04/22/28', '05/10/28', '142',       'IT PM + WS Leads',                 '', 'No'),
  t(145, 3, 'End-User Communication Plan (8000 Users / 180 Sites)',      '20 days',  '05/13/28', '06/07/28', '144',       'Comms + Regional IT',              'Multi-language; AP + AM + EMEA', 'No'),
  t(146, 3, 'Dress Rehearsal and Cutover Simulation',                    '5 days',   '06/03/28', '06/07/28', '144',       'All WS Leads',                     '', 'No'),
  t(147, 3, 'Help Desk 24x7 Activation and Training',                   '10 days',  '06/03/28', '06/14/28', '144',       'IT Ops',                           'Multi-region time zone coverage', 'No'),
  t(148, 3, 'Go No-Go Readiness Assessment',                             '0 days',   '06/14/28', '06/14/28', '145,146,147', 'Steering Committee',             '', 'Yes'),
  t(149, 2, 'QG4 - Implementation Quality Gate',                         '0 days',   '06/28/28', '06/28/28', '123,130,143', 'Steering Committee',             '6000 clients migrated; ERP live in MZ; apps live in MZ; UAT passed; cutover plan approved', 'Yes'),

  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 5 – GOLIVE AND HYPERCARE  (01 Jul 2028 → 30 Sep 2028)
  // ═══════════════════════════════════════════════════════════════════════════
  t(150, 1, 'Phase 5 - GoLive and Hypercare (90 Days)',                  '65 days',  '07/01/28', '09/30/28', '149',       'All Workstreams',                  '', 'No'),

  t(151, 2, '5.1 Day 1 Closing Activities',                              '5 days',   '07/01/28', '07/05/28', '149',       'IT PM + Business',                 '', 'No'),
  t(152, 3, 'Day 1 - Closing (Legal Transfer of JCI Business to Bosch)', '0 days',   '07/01/28', '07/01/28', '149',       'Executive Leadership',             '', 'Yes'),
  t(153, 3, 'Network Domain Cutover to Merger Zone',                     '2 days',   '07/01/28', '07/02/28', '152',       'Bosch Infra',                      '', 'No'),
  t(154, 3, 'ERP Day 1 Go-Live Activation',                              '2 days',   '07/01/28', '07/02/28', '152',       'ERP Team',                         '', 'No'),
  t(155, 3, 'Application Day 1 Activation (All Regions)',                '2 days',   '07/01/28', '07/02/28', '152',       'App Teams',                        '', 'No'),
  t(156, 3, 'TSA Services Activation - JCI Begins Operating IT for Bosch','1 days',  '07/01/28', '07/01/28', '152',       'JCI IT + Bosch IT',                '18-month TSA clock starts 01 Jul 2028; expires 31 Dec 2029', 'No'),
  t(157, 3, 'User Announcement and Go-Live Communication (8000 Users)',  '1 days',   '07/01/28', '07/01/28', '152',       'Comms + Regional IT',              '', 'No'),
  t(158, 3, 'Help Desk 24x7 Live (Multi-Region)',                        '1 days',   '07/01/28', '07/01/28', '152',       'IT Ops',                           '', 'No'),

  t(159, 2, '5.2 Hypercare - 90 Calendar Days',                          '65 days',  '07/01/28', '09/30/28', '151',       'All WS Leads + IT Ops',            '', 'No'),
  t(160, 3, 'Daily Operations Review Meetings (Regional Standups)',      '65 days',  '07/01/28', '09/30/28', '151',       'IT PM + Regional WS Leads',        'AP + AM + EMEA standup cadence', 'No'),
  t(161, 3, 'Incident Management and P1/P2 Issue Resolution',            '65 days',  '07/01/28', '09/30/28', '151',       'IT Ops + Help Desk',               '', 'No'),
  t(162, 3, 'Performance Monitoring and TSA SLA Tracking',               '65 days',  '07/01/28', '09/30/28', '151',       'IT Ops',                           'JCI SLA compliance tracked from Day 1', 'No'),
  t(163, 3, 'Issue Backlog Management and Prioritization',               '65 days',  '07/01/28', '09/30/28', '151',       'IT PM',                            '', 'No'),
  t(164, 3, 'Hypercare Close - Formal Sign-Off',                         '0 days',   '09/30/28', '09/30/28', '160,161,162,163', 'Steering Committee',        '', 'Yes'),
  t(165, 2, 'QG5 - GoLive Quality Gate',                                 '0 days',   '09/30/28', '09/30/28', '164',       'Steering Committee',               'Hypercare complete; P1/P2 resolved; SLA baseline established; stabilization can begin', 'Yes'),

  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 6 – STABILIZATION AND TSA EXIT  (01 Oct 2028 → 31 Dec 2029)
  // ═══════════════════════════════════════════════════════════════════════════
  t(166, 1, 'Phase 6 - Stabilization and TSA Exit',                      '330 days', '10/01/28', '12/31/29', '165',       'All Workstreams + JCI',            'JCI TSA expires 31 Dec 2029 (18 months from Day 1 on 01 Jul 2028)', 'No'),

  t(167, 2, '6.1 Stabilization and Knowledge Transfer',                  '130 days', '10/01/28', '03/28/29', '165',       'IT PM + JCI TSA Team',             '', 'No'),
  t(168, 3, 'TSA Governance - Monthly Service Reviews (JCI)',             '130 days', '10/01/28', '03/28/29', '165',       'IT PM + JCI IT',                   'Monthly reviews; track against SLA and service exit criteria', 'No'),
  t(169, 3, 'Performance Optimization (Merger Zone Operations)',          '60 days',  '10/01/28', '12/24/28', '165',       'IT Ops',                           '', 'No'),
  t(170, 3, 'Knowledge Transfer - JCI to Bosch (Runbooks Processes Docs)','90 days', '10/01/28', '02/07/29', '165',       'JCI IT + Bosch IT',                'All 180 sites; AP and AM priority', 'No'),
  t(171, 3, 'Bosch IT Environment Preparation for Integration',          '60 days',  '10/01/28', '12/24/28', '165',       'Bosch IT Architects',              '', 'No'),
  t(172, 3, 'Bosch Network Integration Planning (180 Sites)',             '30 days',  '12/02/28', '01/14/29', '171',       'Bosch Infra',                      '', 'No'),

  t(173, 2, '6.2 Migration - Merger Zone to Bosch Environment',          '150 days', '01/15/29', '08/14/29', '167,172',   'Bosch IT + App Teams',             '', 'No'),
  t(174, 3, 'Bosch AD Integration and Identity Federation',              '30 days',  '01/15/29', '02/25/29', '172',       'Bosch IAM',                        'Merge Merger Zone AD into Bosch AD', 'No'),
  t(175, 3, 'ERP Integration with Bosch Systems',                        '40 days',  '02/26/29', '04/22/29', '174',       'ERP Team',                         '', 'No'),
  t(176, 3, 'Client Re-Migration to Bosch Domain (6000 Devices)',        '45 days',  '02/26/29', '04/29/29', '174',       'Bosch CWP',                        'Wave-based by region; AP + AM largest effort', 'No'),
  t(177, 3, 'Application Migration Wave 1 - Merger Zone to Bosch',       '30 days',  '02/26/29', '04/08/29', '174',       'App Teams',                        '', 'No'),
  t(178, 3, 'Application Migration Wave 2 - Merger Zone to Bosch',       '30 days',  '04/09/29', '05/20/29', '177',       'App Teams',                        '', 'No'),
  t(179, 3, 'Application Migration Wave 3 - Merger Zone to Bosch',       '30 days',  '05/21/29', '07/01/29', '178',       'App Teams',                        '', 'No'),
  t(180, 3, 'M365 Tenant Migration to Bosch M365 (6000 Mailboxes)',      '30 days',  '04/09/29', '05/20/29', '174',       'Bosch Azure Team',                 '', 'No'),
  t(181, 3, 'Network Integration to Bosch Backbone (180 Sites)',          '30 days',  '06/02/29', '07/11/29', '178',       'Bosch Infra',                      '', 'No'),

  t(182, 2, '6.3 TSA Exit - Wave-Based Service Termination',             '100 days', '07/02/29', '11/14/29', '173,181',   'JCI IT + Bosch IT',                '', 'No'),
  t(183, 3, 'TSA Exit Planning per Service (All JCI Services)',           '15 days',  '07/02/29', '07/22/29', '181',       'IT PM + JCI',                      'Define exit criteria and acceptance tests per service', 'No'),
  t(184, 3, 'TSA Service Exit Wave 1 - Infrastructure Services',         '25 days',  '07/23/29', '08/26/29', '183',       'JCI IT + Bosch IT',                '', 'No'),
  t(185, 3, 'TSA Service Exit Wave 2 - Application Services',            '25 days',  '08/27/29', '09/30/29', '184',       'JCI IT + Bosch IT',                '', 'No'),
  t(186, 3, 'TSA Service Exit Wave 3 - Final Remaining Services',        '25 days',  '10/01/29', '11/04/29', '185',       'JCI IT + Bosch IT',                '', 'No'),
  t(187, 3, 'TSA Exit Confirmation (All JCI Services Terminated)',       '0 days',   '11/14/29', '11/14/29', '186',       'JCI + Bosch Executive',            '', 'Yes'),

  t(188, 2, '6.4 Project Closure',                                       '30 days',  '11/17/29', '12/29/29', '187',       'IT PM + KPMG',                     '', 'No'),
  t(189, 3, 'Lessons Learned Workshop (All Regions)',                    '5 days',   '11/17/29', '11/21/29', '187',       'All WS Leads',                     '', 'No'),
  t(190, 3, 'Project Closure Report',                                    '15 days',  '11/24/29', '12/12/29', '189',       'IT PM + KPMG',                     '', 'No'),
  t(191, 3, 'Final Project Sign-Off',                                    '0 days',   '12/29/29', '12/29/29', '190',       'Executive Leadership',             '', 'Yes'),
  t(192, 2, 'QG6 - Stabilization Quality Gate',                          '0 days',   '12/31/29', '12/31/29', '191',       'Steering Committee',               'TSA fully exited; all 180 sites on Bosch environment; project formally closed', 'Yes'),
];

// ---------------------------------------------------------------------------
// CSV writer
// ---------------------------------------------------------------------------
function escapeCsv(value) {
  const str = String(value ?? '');
  return str.includes(',') || str.includes('"') || str.includes('\n')
    ? `"${str.replace(/"/g, '""')}"`
    : str;
}

const HEADERS = [
  'ID', 'Outline Level', 'Name', 'Duration', 'Start', 'Finish',
  'Predecessors', 'Resource Names', 'Notes', 'Milestone',
];

const lines = [HEADERS.map(escapeCsv).join(',')];

for (const task of tasks) {
  const row = [
    task.id, task.level, task.name, task.duration,
    task.start, task.finish, task.predecessors,
    task.resources, task.notes, task.milestone,
  ];
  lines.push(row.map(escapeCsv).join(','));
}

const outputPath = path.join(__dirname, 'Trinity_Project_Schedule.csv');
fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------
const totalTasks  = tasks.length;
const milestones  = tasks.filter(t => t.milestone === 'Yes').length;
const phases      = tasks.filter(t => t.level === 1).length;
const summaries   = tasks.filter(t => t.level === 2).length;
const leafTasks   = tasks.filter(t => t.level === 3).length;

console.log('─'.repeat(55));
console.log(' Project Trinity – Schedule Generator');
console.log('─'.repeat(55));
console.log(` Output      : ${outputPath}`);
console.log(` Total rows  : ${totalTasks}`);
console.log(`   Phases    : ${phases}`);
console.log(`   Sub-groups: ${summaries}`);
console.log(`   Tasks     : ${leafTasks}`);
console.log(`   Milestones: ${milestones}`);
console.log('─'.repeat(55));
console.log(' MS Project import: File → Open → select CSV,');
console.log(' map Outline Level for WBS hierarchy.');
console.log('─'.repeat(55));
