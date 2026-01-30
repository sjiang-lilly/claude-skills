---
name: ccsp-scraper
description: Process CCSP (Cancer Cell Screening Project) assay data - combines IC50 and Max % Inhibition data from multiple cell line batches with metadata
allowed-tools: Read, Write, Edit, Bash, Glob, Grep
---

# CCSP Scraper Skill

Process Cancer Cell Screening Project (CCSP) cell line assay data into a consolidated output file.

## Author

**Author:** Mandy Jiang, shan.jiang2@lilly.com
**Organization:** Eli Lilly and Company - Bioinformatics
**Version:** 1.0.0
**Created:** January 2026

## What This Skill Does

1. Reads assay data from batch folders (folders ending with `_cells`)
2. Extracts IC50 and Max % Inhibition for each compound
3. Maps TA IDs to compound names using `CompoundList.xlsx`
4. Merges with cell line metadata from `samplemeta_TAA_DepMap25Q2.RData`
5. Outputs consolidated Excel file with merged header row

## Required Input Files

- **Batch folders** (`*_cells/`): Excel files with "Summary" sheet containing assay data
- **CompoundList.xlsx**: TA ID to compound name mapping
- **samplemeta_TAA_DepMap25Q2.RData**: Cell line metadata

## Usage

When user invokes `/ccsp-scraper`, do the following:

1. Check if required files exist in the working directory:
   - At least one folder ending with `cells`
   - `CompoundList.xlsx`
   - `samplemeta_TAA_DepMap25Q2.RData`

2. If `combine_batches.py` doesn't exist, create it with the standard pipeline code

3. Ask user which gene markers to include (default: TACSTD2, CEACAM5)

4. Update `SELECTED_GENE_MARKERS` in `combine_batches.py`

5. Run the script:
   ```bash
   /opt/anaconda3/bin/python3 combine_batches.py
   ```

6. Report results including:
   - Number of batches processed
   - Number of cell lines
   - Output file location

## Dependencies

| Package | Version |
|---------|---------|
| pandas | >= 2.1.4 |
| pyreadr | >= 0.5.4 |
| openpyxl | >= 3.0.10 |

Install with:
```bash
pip install pandas pyreadr openpyxl
```

## Output

**File:** `combined_all_batches.xlsx`

### Output Structure

```
┌─────────┬───────┬───────────────────────────────────────────────────┬───────────────┬─────────────────────────────────────────────────────────┬─────────────────────────────────────────────────────────┐
│         │       │                     Metadata                      │     Gene      │                          IC50                           │                       Max%Inhib                         │
├─────────┼───────┼─────────┼──────────┼──────────┼─────┬──────┬──────┼───────┬───────┼────────────┬─────────────────┬─────────────────┬──────┼────────────┬─────────────────┬─────────────────┬──────┤
│CellLine │ Batch │ ModelID │ Oncotree │ Oncotree │ ... │ Code │in_Bio│TACSTD2│CEACAM5│Staurosporine│Trop2-NMTi-DAR8│CEACAM5-ecys-...│ ... │Staurosporine│Trop2-NMTi-DAR8│CEACAM5-ecys-...│ ... │
├─────────┼───────┼─────────┼──────────┼──────────┼─────┼──────┼──────┼───────┼───────┼────────────┼─────────────────┼─────────────────┼──────┼────────────┼─────────────────┼─────────────────┼──────┤
│ BXPC3   │1st_10 │ACH-000  │ Pancreas │   ...    │ ... │  ... │ Yes  │  9.57 │  8.25 │       0.47 │            0.54 │            1.15 │ ...  │      99.98 │           97.99 │           94.21 │ ...  │
└─────────┴───────┴─────────┴──────────┴──────────┴─────┴──────┴──────┴───────┴───────┴────────────┴─────────────────┴─────────────────┴──────┴────────────┴─────────────────┴─────────────────┴──────┘
```

- **Row 1:** Merged header indicators (Metadata, Gene, IC50, Max%Inhib)
- **Row 2:** Column names
- **Row 3+:** Data (one row per cell line)

### Example Output

| CellLine | Batch | ModelID | OncotreeLineage | ... | TACSTD2 | CEACAM5 | Staurosporine | Trop2-NMTi-DAR8 | ... | Staurosporine | Trop2-NMTi-DAR8 | ... |
|----------|-------|---------|-----------------|-----|---------|---------|---------------|-----------------|-----|---------------|-----------------|-----|
| BXPC3 | 1st_10 cells | ACH-000535 | Pancreas | ... | 9.57 | 8.25 | 0.47 | 0.54 | ... | 99.98 | 97.99 | ... |
| C2BBE1 | 1st_10 cells | ACH-000009 | Bowel | ... | 4.93 | 3.11 | 6.28 | 300 | ... | 97.37 | 39.72 | ... |
| BT20 | 2nd_24 cells | ACH-000578 | Breast | ... | 8.12 | 0.45 | 1.23 | 0.89 | ... | 99.12 | 95.45 | ... |

### Column Breakdown

| Category | Columns | Count |
|----------|---------|-------|
| CellLine, Batch | Cell identifier and batch source | 2 |
| Metadata | ModelID, OncotreeLineage, OncotreePrimaryDisease, OncotreeSubtype, OncotreeCode, in_BioMetas | 6 |
| Gene | User-selected gene markers (default: TACSTD2, CEACAM5) | 2+ |
| IC50 | IC50 values for each compound (nM) | 7 |
| Max%Inhib | Maximum inhibition percentage for each compound | 7 |

## Available Gene Markers

TACSTD2, CEACAM5, ERBB2, ERBB3, MET, EGFR, CD276, F3, MUC1, PTK7, ITGB6, FOLR1, DLL3, ADAM9, LRRC15, FAP, ITGAV, AXL, CDCP1, TPBG, CEACAM6, CLDN18, MSLN, MUC16, CDH17, SLFN11, TOP1, ABCB1, ABCC3, ABCG2, PAF1

## Example Invocation

User: `/ccsp-scraper`

Claude should:
1. Verify input files exist
2. Ask about gene marker selection
3. Run the processing pipeline
4. Report summary of results
