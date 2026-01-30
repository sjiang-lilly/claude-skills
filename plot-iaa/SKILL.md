# Plot IAA/Receptor Bar Chart

## Description
Generate a horizontal bar plot with a broken x-axis showing IAA/Receptor values by gene (TAA). The plot uses a blue-to-red color gradient based on values.

## Input Requirements
- Excel file (`.xlsx`) with two columns:
  - `TAA`: Gene names
  - `IAA/Receptor`: Numeric values

## Usage
```
/plot-iaa
```
Then provide:
1. Path to your Excel file
2. Language preference (Python or R)
3. Output directory (optional, defaults to same as input file)

## Instructions

When this skill is invoked:

1. **Ask the user for**:
   - Path to the Excel file containing TAA and IAA/Receptor data
   - Preferred language: Python or R
   - Output location (optional)

2. **Generate the plotting script** based on language choice:

### Python Version
- Use `pandas` to read Excel
- Use `matplotlib` with broken axis (two subplots side-by-side)
- Apply `coolwarm` colormap based on values
- Sort data by IAA/Receptor (ascending for bottom-to-top display)
- Add diagonal break marks between axis segments
- Save as both PNG (300 dpi) and PDF

### R Version
- Use `readxl` to read Excel
- Use `ggplot2` with `ggbreak` for broken axis
- Apply blue-to-red gradient fill
- Sort and convert TAA to ordered factor
- Use `theme_minimal()` with centered bold title
- Save as both PNG (300 dpi) and PDF

3. **Key plot features**:
   - Horizontal bars (genes on y-axis, values on x-axis)
   - Broken x-axis to handle outliers (break at 15-55 range)
   - Color gradient: blue (low) to red (high)
   - Title: "IAA/Receptor Values by Gene"
   - Figure size: 10 x 7 inches

4. **Run the script** and confirm output files were created.

## Example

User: "I want to create a bar plot for my receptor data"

Response:
1. Ask for Excel file path
2. Ask for Python or R
3. Generate script with user's file paths
4. Execute and show results

## Dependencies

### Python (version 3.11.7)
```
pandas==2.1.4
matplotlib==3.8.0
openpyxl==3.0.10
numpy==1.26.4
```

### R (version 4.4.1)
```
readxl==1.4.3
ggplot2==3.5.1
ggbreak==0.1.6
scales==1.3.0
```
