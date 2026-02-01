# PyLaTeX Beam Force Analysis Report Generator

**Engineering Internship Evaluation Project**

A Python script that generates professional engineering PDF reports from beam force data using PyLaTeX with TikZ/pgfplots vector diagrams.

## Overview

This script reads beam structural analysis data from an Excel file and automatically generates a comprehensive PDF report with:
- LaTeX Tabular force data table (selectable text)
- Shear Force Diagram (SFD) using TikZ/pgfplots
- Bending Moment Diagram (BMD) using TikZ/pgfplots
- Simply supported beam diagram
- Professional engineering report structure

## Key Features

- **Pure PyLaTeX** - No matplotlib, all plots generated with TikZ/pgfplots
- **Programmatic coordinates** - Plot data generated from Python arrays
- **Vector graphics** - Fully scalable diagrams
- **Professional layout** - Industry-standard engineering report format
- **Command-line interface** - Flexible argparse-based execution

## Requirements

### System Requirements
- Python 3.7 or higher
- LaTeX distribution with `pdflatex`:
  - **macOS**: MacTeX or BasicTeX
  - **Linux**: texlive-full
  - **Windows**: MiKTeX or TeX Live

### Python Packages
Install via: `pip install -r requirements.txt`
- pandas (Excel reading)
- openpyxl (Excel engine)
- numpy (numerical operations)
- pylatex (PDF generation)

## Installation

### 1. Install LaTeX

**macOS:**
```bash
brew install --cask basictex
sudo tlmgr update --self
sudo tlmgr install tikz pgfplots float hyperref
```

**Linux:**
```bash
sudo apt-get install texlive-full
```

**Windows:**
Install [MiKTeX](https://miktex.org/download) or [TeX Live](https://www.tug.org/texlive/)

### 2. Install Python Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Basic Command

```bash
python generate_beam_report.py --excel <your_data.xlsx> --output <report_name>
```

### With Beam Image

```bash
python generate_beam_report.py \
    --excel actual_beam_data.xlsx \
    --image beam_diagram.png \
    --output beam_report
```

### Command-Line Arguments

- `--excel` or `-e`: Excel file with beam data (**required**)
- `--image` or `-i`: Beam diagram image (optional)
- `--output` or `-o`: Output PDF name (default: beam_analysis_report)

### Example

```bash
python generate_beam_report.py -e actual_beam_data.xlsx -i beam_diagram.png -o final_report
```
### Excel File Structure

Your Excel file must contain the following columns:

| Column Name | Description | Unit |
|-------------|-------------|------|
| `x` | Position along beam | meters (m) |
| `Shear force` | Shear force at position | kilonewtons (kN) |
| `Bending Moment` | Bending moment at position | kilonewton-meters (kN·m) |

**Example (`actual_beam_data.xlsx`):**

```
x       Shear force     Bending Moment
0.0     45.0           0.0
1.5     36.0           60.75
3.0     27.0           108.0
...
```

### Beam Image (Optional)

Supported formats: PNG, JPG, JPEG. Place in project directory and reference with `--image` flag.

## Output

Generates a professional PDF report containing:
1. Title page with metadata
2. Table of contents (clickable)
3. Introduction (with beam image if provided)
4. Beam description and specifications
5. Force data table (LaTeX Tabular format)
6. Theoretical background
7. Shear Force Diagram (TikZ vector graphic)
8. Bending Moment Diagram (TikZ vector graphic)
9. Summary with statistics
10. Conclusion

## Troubleshooting

**pdflatex not found:** Install LaTeX distribution (see Installation)

**Missing LaTeX packages:** Run `sudo tlmgr install tikz pgfplots float hyperref`

**Excel read error:** Ensure openpyxl is installed: `pip install openpyxl`

**PDF generation fails:** Check that Excel file has correct column names# test contribution
