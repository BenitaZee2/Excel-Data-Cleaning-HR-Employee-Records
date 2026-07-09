# HR Employee Records , Excel Data Cleaning

Cleaning a messy HR dataset into an analysis-ready format using Excel formulas and text functions.

## Overview
Raw employee records contained inconsistent formatting across nearly every field. This project walks through identifying and systematically fixing each issue to produce a clean, analysis-ready dataset.

## Issues Identified & Fixed
- Extra whitespace in names (`"  Sarah Taylor  "` → `Sarah Taylor`)
- Malformed emails (`john at company.com` → `john@company.com`)
- Inconsistent status casing/values (`actve`, `On leave` → standardized `Active`, `On Leave`)
- Currency strings converted to numeric values (`"$117872"` → `117872`)
- Inconsistent location casing (`Sao PAULO` → `Sao Paulo`)
- Region codes standardized against a lookup (`ET` → `Eastern Region`)
- Bonus percentage applied to compute Full Package (Salary + Bonus)
- Salary tiering logic applied to band employees (`High`, etc.)

## Tools
Microsoft Excel , TRIM, SUBSTITUTE, PROPER, VLOOKUP/XLOOKUP, formula-based validation

## Repository Structure
```
before/   , untouched class template (raw dirty data, no cleaning applied yet)
final/    , completed workbook (DirtyData, Rough Work, My Cleaned Data, CleanedData)
```
Within the `final/` workbook:
- `DirtyData` , original raw dataset
- `Rough Work` , intermediate cleaning steps and formulas
- `My Cleaned Data` / `CleanedData` , final cleaned, analysis-ready dataset

## Skills Demonstrated
Data cleaning · text normalization · formula-based validation · HR data structuring
