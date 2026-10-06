# WAHT Trauma & Orthopaedics Digital Handover – UAT Prototype

## Run locally

```bash
npm install
npm run dev
```

Open the local Vite URL shown in the terminal (normally http://localhost:5173).

## NHFD v16 export

The NOF Dashboard now allows users to:

1. Select individual NOF patients or select all NOF patients.
2. Export the selected cohort as an Excel workbook (`.xlsx`).
3. Export the same selected cohort as an NHFD v16 CSV (`.csv`).
4. Review a Validation sheet in the workbook showing missing NHFD essential fields.
5. Review a Local source data sheet showing the fields currently held by this UAT prototype.

The export uses the NHFD v16 (2026) import column structure. NHFD currently accepts CSV import data; the Excel workbook is intended for review/completion before producing the final CSV for upload.

The default operating hospital code is RWP for Worcestershire Acute Hospitals NHS Trust. Confirm this against the Trust/NHFD configuration before a real submission.

## Important

This remains a synthetic UAT prototype. Do not enter or export real patient information until the application has been formally approved, secured, tested, and integrated with the Trust's information-governance and clinical-system requirements.

## NHFD v16 NOF patient form

The NOF Dashboard now includes a dedicated **Add NOF patient** form covering the NHFD v16 (2026) import dataset fields, including conditional surgery, ward-care, delirium, discharge and follow-up fields. Existing NOF patients can be opened with **Edit** in the NHFD data column.

The form stores the NHFD fields with each NOF record. The existing **Export selected for NHFD** action uses those stored fields to create the NHFD v16 CSV and Excel workbook, with a Validation sheet. The two calculated 4AT score columns are intentionally excluded from the import because NHFD calculates them from their component fields.

The field list/options were aligned to the NHFD v16 import specification (dataset v1.02, dated 16/02/2026). See the official NHFD dataset documentation for the current specification before any real submission.

## NOF performance dashboard

The NOF Dashboard includes a reporting-period selector for **Last 24 hours, Last 1 week, Last 30 days, and Last 12 months**. Performance percentages are calculated from NOF records with an NHFD presentation date in the selected period. The dashboard displays percentage bar charts for NHFD-aligned assessment, perioperative and outcome/best-practice measures. Records without an NHFD presentation date are excluded from period calculations and are flagged.

The performance display is a local operational implementation using the fields recorded in the NOF NHFD form. NHFD's own calculations, exclusions, follow-up windows and submitted-data results remain authoritative for formal NHFD reporting.

## NOF dashboard sample data

The NOF Dashboard contains 10 fully synthetic sample NOF patients (UAT-001, UAT-002, UAT-006, UAT-007, UAT-010, UAT-013, UAT-015, UAT-017, UAT-019 and UAT-021) with populated NHFD v16 fields. Their dates are distributed across 2026 so the reporting-period percentages and monthly run chart display meaningful demonstration data. These records are fictional and must not be treated as real patient data or submitted to NHFD.
