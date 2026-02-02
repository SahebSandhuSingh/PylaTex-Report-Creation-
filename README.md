# Beam Force Analysis Report Generator

**Automated PDF Report Generation for Structural Engineering Analysis**

A professional Python-based tool that transforms beam structural analysis data into publication-ready PDF reports with high-quality vector graphics, leveraging PyLaTeX and TikZ/pgfplots for superior visualization.

---

## Overview

This automated report generator processes beam analysis data from Excel spreadsheets and produces comprehensive engineering documentation featuring:

- **Data Tables**: LaTeX-formatted tabular displays with full text selectability
- **Shear Force Diagrams (SFD)**: Color-coded contour visualizations with gradient fills
- **Bending Moment Diagrams (BMD)**: Multi-color spectrum representations
- **Structural Diagrams**: Integrated beam configuration imagery
- **Professional Formatting**: Industry-standard report structure with automatic table of contents

All visualizations are rendered as **pure vector graphics** using TikZ/pgfplots, ensuring scalability and print-quality output.

---

## Key Features

### Technical Excellence
- **Native LaTeX Rendering**: All diagrams generated through TikZ/pgfplots—no matplotlib dependency
- **High-Resolution Gradients**: 200-point interpolation for smooth color transitions
- **Programmatic Coordinates**: Direct array-to-plot conversion from numerical data
- **Vector Graphics**: Infinite scalability without quality degradation
- **Intelligent Color Mapping**: Spectrum-based value encoding (blue→cyan→green→yellow→orange→red)

### User Experience
- **Command-Line Interface**: Streamlined argparse-based workflow
- **Flexible Input Handling**: Automatic column detection with multiple naming conventions
- **Error Resilience**: Comprehensive validation and informative error messages
- **Cross-Platform Compatibility**: Works on macOS, Linux, and Windows

---

## System Requirements

### Software Dependencies

| Component | Minimum Version | Purpose |
|-----------|----------------|---------|
| Python | 3.7+ | Core runtime environment |
| LaTeX Distribution | 2020+ | PDF compilation engine |
| pdflatex | Included in LaTeX | Document rendering |

### LaTeX Distribution Options

**macOS:**
- [MacTeX](https://www.tug.org/mactex/) (Full: 4GB) or [BasicTeX](https://www.tug.org/mactex/morepackages.html) (Minimal: 100MB)

**Linux:**
- `texlive-full` (Debian/Ubuntu) or `texlive-scheme-full` (Fedora/RHEL)

**Windows:**
- [MiKTeX](https://miktex.org/) (Auto-installs packages) or [TeX Live](https://www.tug.org/texlive/)

### Required LaTeX Packages

These packages are automatically included in full distributions or can be installed individually:

```
tikz, pgfplots, float, hyperref, geometry, 
fontenc, inputenc, lastpage, graphicx, xcolor
```

### Python Dependencies

All Python packages are specified in `requirements.txt`:

```plaintext
pandas>=1.3.0         # Excel data manipulation
openpyxl>=3.0.0       # Excel file reader
numpy>=1.20.0         # Numerical array operations
scipy>=1.7.0          # Scientific computing (interpolation)
pylatex>=1.4.0        # LaTeX document generation
```

---

## Installation Guide

### Step 1: LaTeX Installation

**macOS (Homebrew):**
```bash
# Install BasicTeX (recommended for minimal footprint)
brew install --cask basictex

# Update package manager and install required packages
sudo tlmgr update --self
sudo tlmgr install tikz pgfplots float hyperref geometry
```

**Linux (Debian/Ubuntu):**
```bash
sudo apt-get update
sudo apt-get install texlive-latex-extra texlive-pictures
```

**Linux (Fedora/RHEL):**
```bash
sudo dnf install texlive-scheme-medium
```

**Windows:**
1. Download and install [MiKTeX](https://miktex.org/download)
2. During installation, select "Install missing packages on-the-fly: Yes"

### Step 2: Python Environment Setup

**Create Virtual Environment (Recommended):**
```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
# macOS/Linux:
source venv/bin/activate
# Windows:
venv\Scripts\activate
```

**Install Python Dependencies:**
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### Step 3: Verify Installation

```bash
# Check LaTeX installation
pdflatex --version

# Check Python packages
pip list | grep -E 'pandas|numpy|pylatex'

# Test run (with sample data)
python3 generate_beam_report.py -e actual_beam_data.xlsx -o test_report
```

---

## Usage

### Quick Start

```bash
python3 generate_beam_report.py --excel <data_file.xlsx> --output <report_name>
```

### Complete Example

```bash
python3 generate_beam_report.py \
    --excel actual_beam_data.xlsx \
    --image beam_diagram.png \
    --output beam_analysis_report
```

This generates: `beam_analysis_report.pdf`

### Command-Line Arguments

| Argument | Short | Required | Description | Default |
|----------|-------|----------|-------------|---------|
| `--excel` | `-e` | **Yes** | Path to Excel file with beam data | None |
| `--image` | `-i` | No | Path to beam diagram image (PNG/JPG) | None |
| `--output` | `-o` | No | Output PDF filename (without .pdf) | `beam_analysis_report` |

### Usage Examples

**Minimal (Excel only):**
```bash
python3 generate_beam_report.py -e beam_data.xlsx
```

**With beam diagram:**
```bash
python3 generate_beam_report.py -e beam_data.xlsx -i beam_structure.png
```

**Custom output name:**
```bash
python3 generate_beam_report.py -e data.xlsx -o project_final_report
```

**Full command:**
```bash
python3 generate_beam_report.py \
    --excel experiment_beam_analysis.xlsx \
    --image simply_supported_beam.png \
    --output mechanical_analysis_report_2026
```

---

## Input Data Format

### Excel File Structure

The Excel file **must** contain three columns with force analysis data. The script automatically detects column names with flexible matching:

#### Column Requirements

| Column (Variants) | Data Type | Unit | Description |
|-------------------|-----------|------|-------------|
| `x`, `Position` | Float | meters (m) | Position along beam length |
| `Shear force`, `Shear Force` | Float | kilonewtons (kN) | Shear force at each position |
| `Bending Moment`, `Bending moment` | Float | kN·m | Bending moment at each position |

#### Sample Excel File (`actual_beam_data.xlsx`)

| x   | Shear force | Bending Moment |
|-----|-------------|----------------|
| 0.0 | 45.0        | 0.0            |
| 1.5 | 36.0        | 60.75          |
| 3.0 | 27.0        | 108.0          |
| 4.5 | 18.0        | 141.75         |
| 6.0 | 9.0         | 162.0          |
| 7.5 | 0.0         | 168.75         |
| 9.0 | -9.0        | 162.0          |
| 10.5| -18.0       | 141.75         |
| 12.0| -27.0       | 108.0          |
| 13.5| -36.0       | 60.75          |
| 15.0| -45.0       | 0.0            |

**Data Requirements:**
- Minimum 3 data points (recommended: 10+ for smooth curves)
- Consistent units throughout
- No missing or null values
- Sorted by position (ascending order preferred)

### Beam Diagram Image (Optional)

**Supported Formats:** PNG, JPG, JPEG  
**Recommended Resolution:** 1200x400 pixels or higher  
**Purpose:** Visual representation of beam configuration included in report introduction

Place the image in the project directory and reference it with the `--image` flag.

---

## Output

### Generated PDF Report Structure

The tool generates a comprehensive, professionally formatted PDF document containing:

#### 1. **Title Page**
   - Report title
   - Generation date and timestamp
   - Project metadata

#### 2. **Table of Contents**
   - Hyperlinked section navigation
   - Automatic page numbering

#### 3. **Introduction**
   - Project overview
   - Beam diagram (if provided)
   - Analysis objectives

#### 4. **Beam Description**
   - Structural configuration
   - Support conditions
   - Material properties (if applicable)

#### 5. **Force Data Table**
   - LaTeX tabular format
   - Position, shear force, and bending moment columns
   - Fully selectable text
   - Centered alignment with borders

#### 6. **Theoretical Background**
   - Beam theory fundamentals
   - Sign conventions
   - Calculation methodologies

#### 7. **Shear Force Diagram (SFD)**
   - **Visualization Type:** Contour-style gradient fill
   - **Color Encoding:** Blue (negative) → Cyan → Green → Yellow → Orange → Red (positive)
   - **Resolution:** 200 interpolated points for smooth gradients
   - **Features:**
     - Black trajectory curve with markers
     - Zero reference baseline (dashed)
     - Axis labels with units
     - Grid lines for reading values

#### 8. **Bending Moment Diagram (BMD)**
   - **Visualization Type:** Multi-color spectrum representation
   - **Color Encoding:** Deep Blue (maximum magnitude) → Cyan → Yellow → Red (near zero)
   - **Curve Type:** Smooth interpolation (quadratic)
   - **Features:**
     - Color-coded magnitude representation
     - Continuous gradient fill below curve
     - Axis annotations

#### 9. **Analysis Summary**
   - Maximum/minimum values
   - Critical positions
   - Statistical data

#### 10. **Conclusion**
   - Key findings
   - Engineering recommendations

### Technical Specifications

- **File Format:** PDF (Portable Document Format)
- **Graphics:** Vector-based (TikZ/pgfplots)
- **Typography:** Latin Modern font family
- **Page Size:** A4 (210mm × 297mm)
- **Margins:** 2.5cm all sides
- **Hyperlinks:** Fully functional internal navigation

---

## Technical Implementation

### Architecture

```
┌─────────────────────────────────────────────────┐
│         generate_beam_report.py                 │
├─────────────────────────────────────────────────┤
│  1. Data Ingestion (pandas/openpyxl)          │
│  2. High-Resolution Interpolation (scipy)      │
│  3. Color Gradient Generation (numpy)          │
│  4. TikZ Code Synthesis (200 \fill commands)   │
│  5. LaTeX Document Assembly (pylatex)          │
│  6. PDF Compilation (pdflatex)                 │
└─────────────────────────────────────────────────┘
```

### Visualization Technology Stack

**Color Gradient System:**
- **Method:** 200 narrow vertical rectangular strips per diagram
- **Color Specification:** xcolor package mixing notation (`blue!50!cyan`)
- **Interpolation:** scipy.interpolate.interp1d (linear for SFD, quadratic for BMD)
- **Rendering:** Pure TikZ `\fill` commands—no external graphics

**Data Flow:**
```
Excel → pandas DataFrame → numpy arrays → scipy interpolation 
→ normalized color values → xcolor mix strings → TikZ \fill commands 
→ pgfplots axis environment → pdflatex → PDF output
```

### Why TikZ/pgfplots?

1. **Vector Graphics:** Infinite scalability without pixelation
2. **Text Integration:** Equations and labels render with document fonts
3. **File Size:** Smaller than raster image alternatives
4. **Consistency:** Matches LaTeX document typography
5. **Publication-Ready:** Accepted by academic journals and technical publications

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `pdflatex: command not found` | Install LaTeX distribution (see Installation section) and add to PATH: `export PATH="/Library/TeX/texbin:$PATH"` |
| PDF generation fails with LaTeX errors | Check Excel column names (must be: x, Shear force, Bending Moment). View detailed errors: `cat Beam_Report.log` |
| `ValueError: x must be strictly increasing` | Sort Excel data by position column in ascending order before running |
| Plots appear black or incorrectly colored | Ensure complete LaTeX installation with all packages: `sudo tlmgr install tikz pgfplots xcolor` |

---

## Project Structure

```
PylaTex-Report-Creation-/
│
├── generate_beam_report.py    # Main script
├── requirements.txt            # Python dependencies
├── README.md                   # Documentation (this file)
│
├── actual_beam_data.xlsx       # Sample input data
├── beam_diagram.png            # Sample beam image
│
├── Beam_Report.tex             # Generated LaTeX source (temporary)
├── Beam_Report.pdf             # Generated output document
└── Beam_Report.{aux,log,toc}   # LaTeX compilation artifacts
```

---

## Advanced Configuration

### Customizing Color Schemes

Edit the color gradient mapping in `generate_beam_report.py`:

**Shear Force Diagram (lines 175-195):**
```python
# Modify color transitions
if norm_val < 0.5:
    color_def = f"purple!{blend}!blue"  # Change from blue
else:
    color_def = f"red!{blend}!orange"    # Customize high end
```

**Bending Moment Diagram (lines 285-305):**
```python
# Adjust magnitude-to-color mapping
color_def = f"green!{blend}!yellow"  # Use different color scheme
```

### Adjusting Interpolation Resolution

**Higher quality (slower, larger file):**
```python
x_high_res = np.linspace(positions.min(), positions.max(), 500)  # 500 points
```

**Lower quality (faster, smaller file):**
```python
x_high_res = np.linspace(positions.min(), positions.max(), 50)   # 50 points
```

### Modifying Plot Dimensions

Edit axis parameters in the pgfplots code (lines 215-230):

```python
width=\textwidth,      # Change to width=0.8\textwidth for smaller
height=7cm,            # Adjust height
```

---

## Performance Benchmarks

**System:** MacBook Pro M1, 16GB RAM  
**LaTeX:** BasicTeX 2024  
**Test Data:** 11 data points

| Operation | Time | Notes |
|-----------|------|-------|
| Data Loading | 0.15s | Excel read + validation |
| Interpolation | 0.08s | 200 points × 2 diagrams |
| LaTeX Generation | 0.22s | TikZ code synthesis |
| PDF Compilation | 2.3s | pdflatex processing |
| **Total Runtime** | **~3s** | End-to-end execution |

**File Sizes:**
- Input Excel: 9 KB
- Generated LaTeX: 45 KB
- Output PDF: 165 KB (with embedded vector graphics)

---

## Frequently Asked Questions

**Q: Can I use this with data from FEA software?**  
A: Yes, export your results to Excel with the required column format (x, Shear force, Bending Moment).

**Q: Does it support multiple beams in one report?**  
A: Currently supports single beam analysis. Modify the script for multi-beam reports.

**Q: Can I change units (e.g., pounds instead of kilonewtons)?**  
A: Yes, update axis labels in the `generate_sfd_plot()` and `generate_bmd_plot()` functions.

**Q: Is Windows Subsystem for Linux (WSL) supported?**  
A: Yes, install LaTeX within WSL environment: `sudo apt-get install texlive-full`

**Q: Can I embed this in a web application?**  
A: Yes, but ensure server has Python + LaTeX installed and appropriate file permissions.

---

## Contributing

Contributions are welcome! Areas for improvement:

- [ ] Add support for distributed loads visualization
- [ ] Implement moment diagram reactions at supports
- [ ] Create web interface with Flask/Django
- [ ] Add unit testing suite
- [ ] Support for cantilever and continuous beams
- [ ] Export to different page sizes (Letter, Legal)
- [ ] Multi-language support for report text

---

## License

This project is provided as-is for educational and professional use. Modify and distribute freely with attribution.

---

## Acknowledgments

- **PyLaTeX**: Python interface for LaTeX document generation
- **TikZ/pgfplots**: Superior LaTeX graphics packages
- **pandas**: Powerful data manipulation library
- **scipy**: Scientific computing tools

---

## Contact & Support

For issues, feature requests, or questions:
- Open an issue in the repository
- Check existing documentation in README
- Review LaTeX compilation logs for detailed error messages

**Version:** 1.0  
**Last Updated:** February 2026  
**Python Compatibility:** 3.7+  
**LaTeX Compatibility:** TeXLive 2020+, MiKTeX 2020+
