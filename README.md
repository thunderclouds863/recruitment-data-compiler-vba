# Recruitment Data Compiler (Excel VBA)

An Excel VBA system that compiles weekly recruitment data from many recruiter files into one workbook, then builds the monitoring reports and the weekly Board of Directors (BOD) presentation.

Built while working in Corporate Talent Acquisition and People Analytics at Cimory Group (2026). It was later rebuilt as a web application: [fptk-database](https://github.com/thunderclouds863/fptk-database).

> **No real data.** The original workbook holds candidate data, so only its structure is published: a clean template without data, macros or dashboards, plus the VBA source. Colleague names were replaced with placeholders (`Recruiter01`, `Recruiter02`, …). Business-unit codes (CORP, MP, MS, CMD, JESS) are kept.

## Problem

Each recruiter and business unit kept their own Excel file for manpower requests (FPTK) and candidate sourcing. Every week about 17 files with different headers, column orders and date formats had to be combined by hand before the BOD report could be prepared.

## What the system does

| Area | What it does | Main modules |
|---|---|---|
| Compile | Reads every Excel file in the configured input folders and compiles manpower requests (FPTK), candidate sourcing, position codes and the candidate blacklist into master sheets | `m02_MainCompiler`, `m11_DBSourcingCompiler`, `m06_CompileDBKodePosisi_FPTK`, `m10_CompileBlacklist`, `m09_helpercompilefptk` |
| Schema-driven headers | Maps differently named or ordered source headers to one target layout using a schema configuration sheet, with a form to adjust mappings and a log of schema changes | `modSchemaAdjustment`, `modSchemaFormSupport`, `frmHeaderAdjustment` |
| Data quality | Normalizes names, dates and codes, validates rows, detects duplicates, and keeps a trace of corrupted rows | `modLookup`, `m13_UtilityFunctions`, `modFPTKCompatibilityCleanup`, `Module4` |
| Monitoring reports | Builds sourcing monitoring, recruitment funneling, recruiter performance and the FPTK-to-recruiter mapping report | `modMonitoring`, `modRecruiterMapping` |
| BOD presentation | Updates the weekly PowerPoint deck from a mapping sheet (late binding, no PowerPoint reference needed) | `UpdatetoPPT` |
| Run control | Progress bar, run log, settings sheet, backups before compiling, archive of outputs, save to OneDrive | `m15_ProgressManagement`, `m14_SettingsAndLog`, `m17_ExportAndBackup`, `modSaveToOneDrive` |
| Input forms | Installs data-entry forms for FPTK and sourcing | `modinstallerdbfptk`, `modSourcingSupport`, `Module11` |

## Result

- About 17 heterogeneous recruiter files compiled automatically into one standardized dataset.
- About 60–70% of the recurring weekly reporting automated, including the Excel-to-PowerPoint update. Qualitative updates and PIC-level progress stay manual.

## Compile modes

| Entry point | Runs |
|---|---|
| `CompileFPTK_All` | Compile data only (fastest) |
| `CompileFPTK_All_Full` | Compile, monitoring, funneling, export and backup |
| `CompileDBKodePosisiOnly` | Position-code database only |
| `RefreshReportsAndExportOnly` | Reports and export without recompiling |

## Files

- `FPTK_Template.xlsx`: empty workbook structure (no macros, no dashboards):
  - header rows of the data sheets: FPTK, DB Sourcing, DB Kode Posisi, Blacklist Candidate, Recruiter Master, Direktorat, Kode Dropdown
  - `Settings` with setting names only
  - `Schema_Config` in full: 341 field definitions with aliases, data types, form controls and dropdown values that drive the header mapping
  - `Flow Map Model` in full: recruitment steps per recruitment model
  - header rows of the log sheets
- `src/`: 32 VBA modules and 1 user form, about 49,000 lines. To run them, import the modules into a macro-enabled copy of the template (VBA editor → File → Import File).
- `frmHeaderAdjustment.frm` contains the form code only; its layout file (`.frx`) is not included.

## Tools

Excel VBA, Excel tables, PowerPoint automation, schema-driven data mapping.
