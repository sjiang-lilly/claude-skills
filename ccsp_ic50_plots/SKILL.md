---
name: ccsp-ic50-plots
description: Extract IC50 dose-response curve plots from CCSP (Cancer Cell Screening Project) Excel files and generate an HTML summary table. Use when user wants to extract plots from "Data analysis for IC50" sheets, create visual comparisons of IC50 curves across cell lines, or generate plot tables from CCSP screening data files (*_paste.xlsx).
---

# CCSP IC50 Plot Extractor

Extract IC50 % Inhibition dose-response curve plots from CCSP Excel files and generate an HTML summary table with cell lines as rows and compounds as columns.

## Author

**Author:** Mandy Jiang (shan.jiang2@lilly.com)
**Organization:** Eli Lilly and Company - Bioinformatics
**Version:** 1.3.0
**Created:** February 2026

## When to Use

- User uploads a zip file containing CCSP Excel files
- User provides a path to folder containing CCSP Excel files (format: `*_paste.xlsx`)
- User wants to extract % Inhibition plots (not Response plots)
- User wants to create a comparison table of IC50 curves across cell lines

## Input Formats

The skill accepts two input formats:

1. **Zip file**: Upload a `.zip` file containing CCSP Excel files. The script will automatically extract and find the Excel files (handles nested folders, ignores `__MACOSX`).

2. **Folder path**: Provide a path to a folder containing CCSP Excel files directly or in a subfolder.

## Workflow

### 1. Handle Input (Zip or Folder)

- If input is a `.zip` file: extract to temporary directory
- If input is a folder: use directly or search subfolders for Excel files
- Find all `*_paste.xlsx` files (exclude Summary files)

### 2. Extract Cell Line Name

Parse filename to get cell line: `YYYYMMDD_CELLLINE_NTA_NNH_paste.xlsx` → `CELLLINE`

### 3. Extract Compound IDs from Excel (Dynamic)

Read compound IDs from the **XLFit Chart section** in "Data analysis for IC50" sheet:
1. Open Excel file with openpyxl
2. Find "XLFit Chart" marker row in column B
3. Read compound names from subsequent rows in column B
4. Extract compound ID from format `COMPOUND_CELLLINE` → `COMPOUND` (first part before underscore)
5. **Exclude Staurosporine** (control compound) from the list
6. Stop when reaching empty row

**Important**: 
- The XLFit Chart section defines the actual plot order, which matches the embedded EMF images
- Staurosporine is excluded from extraction and output
- The number of compounds is determined dynamically from the Excel file

### 4. Extract and Convert % Inhibition Images Only

For each Excel file:
1. Unzip the .xlsx file (it's a ZIP archive)
2. Find EMF images in `xl/media/` folder
3. Filter to actual plots only (file size > 3000 bytes, excludes placeholder images)
4. Sort by image number
5. **Take first N images only** where N = number of test compounds (excluding Staurosporine)
   - These are the % Inhibition plots for test compounds
   - Staurosporine plot and all Response plots are excluded
6. Convert EMF to PNG using inkscape:
   ```bash
   inkscape input.emf --export-filename output.png
   ```

**Inkscape path**: `/Users/L052239/.local/bin/homebrew/bin/inkscape`

### 5. Generate HTML Table

Create table with:
- **Rows**: Cell lines (sorted alphabetically)
- **Columns**: Test compounds only (Staurosporine excluded), in plot order
- **Cells**: Embedded PNG images (base64 encoded)

## Script Usage

```bash
# Set Inkscape path (Homebrew installation)
export PATH="/Users/L052239/.local/bin/homebrew/bin:$PATH"

# From zip file
python scripts/extract_ic50_plots.py data.zip output.html

# From folder path
python scripts/extract_ic50_plots.py /path/to/ccsp_folder output.html

# With optional configuration
python scripts/extract_ic50_plots.py data.zip output.html \
    --compound-map compound_map.json \
    --cell-colors cell_colors.json
```

### Optional JSON Configuration Files

**compound_map.json** - Map compound IDs to display names (optional):
```json
{
    "TA145": "ADC-001",
    "TA146": "ADC-002",
    "TA147": "ADC-003"
}
```

**cell_colors.json** - Color-code cell lines by group:
```json
{
    "BXPC3": "#90EE90",
    "ASPC1": "#87CEEB",
    "LS513": "#FFD700"
}
```

## Dependencies

- Python 3.x
- openpyxl (for reading Excel files)
- inkscape (for EMF to PNG conversion)

Install dependencies:
```bash
pip install openpyxl
apt-get install inkscape  # or brew install inkscape on macOS
```

## Output

HTML file with:
- Responsive table layout
- Column headers showing test compound IDs only (Staurosporine excluded)
- Compound headers match the actual plots in the same column
- Optional compound display names from mapping file
- Cell line column with optional color coding
- Base64-embedded images (self-contained, no external dependencies)

## Key Features

1. **Dynamic compound detection**: Compounds read from XLFit Chart section to match plot order
2. **% Inhibition only**: Extracts only % Inhibition plots, excludes Response plots
3. **Staurosporine excluded**: Control compound not included in extraction or output
4. **Correct plot-to-header mapping**: Compound headers match the actual plots
5. **Flexible input**: Accepts zip files or folder paths
6. **Self-contained output**: All images embedded as base64 in HTML
