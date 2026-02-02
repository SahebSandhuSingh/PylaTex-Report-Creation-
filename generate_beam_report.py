import argparse
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from pylatex import (
    Document, Section, Subsection, Figure, Tabular, 
    Package, NoEscape, Command, NewPage
)
from pylatex.utils import bold


def read_excel_data(excel_path):
    """
    Read beam force data from Excel file.
    
    Args:
        excel_path (str): Path to the Excel file containing beam force data.
        
    Returns:
        pd.DataFrame: DataFrame containing the force data.
        
    Raises:
        FileNotFoundError: If Excel file does not exist.
        Exception: If file cannot be read or is malformed.
    """
    try:
        excel_file = Path(excel_path)
        if not excel_file.exists():
            raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
        # Read Excel file using pandas
        df = pd.read_excel(excel_path)
        
        print(f"✓ Successfully read data from: {excel_path}")
        print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
        
        return df
    
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}", file=sys.stderr)
        raise


def create_force_table(doc, df):
    """
    Create a LaTeX Tabular table from the beam force DataFrame.
    
    Args:
        doc: PyLaTeX Document object.
        df (pd.DataFrame): DataFrame containing force data.
    """
    with doc.create(Section('Input Data')):
        doc.append('The following table presents the input force data extracted from the Excel file.')
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        
        # Create table specification based on number of columns
        num_cols = len(df.columns)
        table_spec = '|' + '|'.join(['c'] * num_cols) + '|'
        
        doc.append(NoEscape(r'\begin{center}'))
        with doc.create(Tabular(table_spec)) as table:
            # Add header row
            table.add_hline()
            header = [bold(str(col)) for col in df.columns]
            table.add_row(header)
            table.add_hline()
            
            # Add data rows
            for _, row in df.iterrows():
                # Format numeric values to 2 decimal places
                formatted_row = []
                for val in row:
                    if isinstance(val, (int, float)):
                        formatted_row.append(f"{val:.2f}")
                    else:
                        formatted_row.append(str(val))
                table.add_row(formatted_row)
                table.add_hline()
        doc.append(NoEscape(r'\end{center}'))


def calculate_shear_bending_moments(df):
    """
    Extract shear force and bending moment arrays from DataFrame.
    
    Args:
        df (pd.DataFrame): DataFrame with force data containing columns:
                          - Position/x: beam positions (m)
                          - Shear force: shear force values (kN) 
                          - Bending Moment: bending moment values (kN·m)
        
    Returns:
        tuple: (positions, shear_forces, bending_moments)
    """
    # Identify position column
    if 'x' in df.columns:
        positions = df['x'].values
    elif 'Position' in df.columns:
        positions = df['Position'].values
    else:
        positions = df.iloc[:, 0].values
    
    # Identify shear force column
    if 'Shear force' in df.columns:
        shear_forces = df['Shear force'].values
    elif 'Shear Force' in df.columns:
        shear_forces = df['Shear Force'].values
    elif len(df.columns) > 1:
        shear_forces = df.iloc[:, 1].values
    else:
        shear_forces = np.zeros(len(positions))
    
    # Identify bending moment column
    if 'Bending Moment' in df.columns:
        bending_moments = df['Bending Moment'].values
    elif 'Bending moment' in df.columns:
        bending_moments = df['Bending moment'].values
    elif len(df.columns) > 2:
        bending_moments = df.iloc[:, 2].values
    else:
        bending_moments = np.zeros(len(positions))
    
    return positions, shear_forces, bending_moments


def generate_sfd_plot(positions, shear_forces):
    """
    Generate pgfplots code for Shear Force Diagram with colored fill between curve and zero.
    
    Args:
        positions (np.ndarray): Array of beam positions (m).
        shear_forces (np.ndarray): Array of shear force values (kN).
        
    Returns:
        str: pgfplots LaTeX code for the SFD plot.
    """
    from scipy.interpolate import interp1d
    
    # Invert shear forces to match required trajectory direction
    shear_forces = -shear_forces
    
    # Calculate value range
    sf_min = shear_forces.min()
    sf_max = shear_forces.max()
    
    # High-resolution interpolation for smooth gradient (200 points)
    x_high_res = np.linspace(positions.min(), positions.max(), 200)
    f_interp = interp1d(positions, shear_forces, kind='linear')
    y_high_res = f_interp(x_high_res)
    
    # Generate coordinate pairs for trajectory (with parentheses for pgfplots)
    trajectory_coords = "\n".join([f"({x:.3f},{f:.3f})" for x, f in zip(positions, shear_forces)])
    
    # Generate narrow colored vertical strips
    colored_strips = []
    for i in range(len(x_high_res) - 1):
        x1, x2 = x_high_res[i], x_high_res[i+1]
        y_val = y_high_res[i]
        # Calculate color based on y value (normalized to [0,1])
        norm_val = (y_val - sf_min) / (sf_max - sf_min) if sf_max != sf_min else 0.5
        # Map to colormap: blue (negative) -> cyan -> green -> yellow -> orange -> red (positive)
        if norm_val < 0.2:
            # Deep blue
            blend = int(norm_val * 500)
            color_def = f"blue!{100-blend}"
        elif norm_val < 0.4:
            # Blue to cyan
            blend = int((norm_val - 0.2) * 500)
            color_def = f"cyan!{blend}!blue"
        elif norm_val < 0.5:
            # Cyan to green
            blend = int((norm_val - 0.4) * 1000)
            color_def = f"green!{blend}!cyan"
        elif norm_val < 0.6:
            # Green to yellow  
            blend = int((norm_val - 0.5) * 1000)
            color_def = f"yellow!{blend}!green"
        elif norm_val < 0.8:
            # Yellow to orange
            blend = int((norm_val - 0.6) * 500)
            color_def = f"orange!{blend}!yellow"
        else:
            # Orange to dark red
            blend = int((norm_val - 0.8) * 500)
            color_def = f"red!{100-blend/2}!orange"
        
        if y_val >= 0:
            strip = f"\\fill[{color_def}] (axis cs:{x1:.4f},0) rectangle (axis cs:{x2:.4f},{y_val:.4f});"
        else:
            strip = f"\\fill[{color_def}] (axis cs:{x1:.4f},{y_val:.4f}) rectangle (axis cs:{x2:.4f},0);"
        colored_strips.append(strip)
    
    strips_code = "\n".join(colored_strips)
    
    # Calculate axis limits
    beam_min = positions.min()
    beam_max = positions.max()
    y_min = sf_min * 1.2 if sf_min < 0 else -5
    y_max = sf_max * 1.2 if sf_max > 0 else 5
    
    # pgfplots code with colored strips
    pgfplots_code = r"""
\begin{figure}[htbp]
\centering
\begin{tikzpicture}
\begin{axis}[
    width=\textwidth,
    height=7cm,
    xlabel={Beam Length (m)},
    ylabel={Shear Force (kN)},
    xmin=""" + f"{beam_min:.3f}" + r""",
    xmax=""" + f"{beam_max:.3f}" + r""",
    ymin=""" + f"{y_min:.3f}" + r""",
    ymax=""" + f"{y_max:.3f}" + r""",
    grid=major,
    axis lines=left
]

% Colored vertical strips for gradient effect
""" + strips_code + r"""

% Zero reference line
\addplot[black, very thin, dashed] {0};

% Trajectory curve with markers
\addplot[
    black,
    thick,
    mark=*,
    mark size=1.5pt
] coordinates {
""" + trajectory_coords + r"""
};

\end{axis}
\end{tikzpicture}
\caption{Shear Force Diagram (SFD) - Contour Visualization}
\label{fig:sfd}
\end{figure}
"""
    
    return pgfplots_code


def generate_bmd_plot(positions, bending_moments):
    """
    Generate pgfplots code for Bending Moment Diagram with colored fill between curve and zero.
    
    Args:
        positions (np.ndarray): Array of beam positions (m).
        bending_moments (np.ndarray): Array of bending moment values (kN·m).
        
    Returns:
        str: pgfplots LaTeX code for the BMD plot.
    """
    from scipy.interpolate import interp1d
    
    # Calculate value range (use absolute values for color mapping)
    bm_min = bending_moments.min()
    bm_max = bending_moments.max()
    bm_abs_max = max(abs(bm_min), abs(bm_max))
    
    # High-resolution interpolation for smooth gradient (200 points)
    x_high_res = np.linspace(positions.min(), positions.max(), 200)
    f_interp = interp1d(positions, bending_moments, kind='quadratic')
    y_high_res = f_interp(x_high_res)
    
    # Generate coordinate pairs for trajectory (with parentheses for pgfplots)
    trajectory_coords = "\n".join([f"({x:.3f},{f:.3f})" for x, f in zip(positions, bending_moments)])
    
    # Generate narrow colored vertical strips (blue for negative, yellow/red for near-zero)
    colored_strips = []
    for i in range(len(x_high_res) - 1):
        x1, x2 = x_high_res[i], x_high_res[i+1]
        y_val = y_high_res[i]
        
        # Color based on actual value: yellow/red at edges (near 0), blue at center (most negative)
        norm_val = (y_val - bm_min) / (bm_max - bm_min) if bm_max != bm_min else 0.5
        
        # Map to colormap: red/yellow (near 0) -> cyan -> deep blue (most negative)
        if norm_val < 0.2:
            # Most negative: deep blue
            blend = int(norm_val * 500)
            color_def = f"blue!{100-blend}"
        elif norm_val < 0.4:
            # Blue to cyan
            blend = int((norm_val - 0.2) * 500)
            color_def = f"cyan!{blend}!blue"
        elif norm_val < 0.7:
            # Cyan to yellow
            blend = int((norm_val - 0.4) * 333)
            color_def = f"yellow!{blend}!cyan"
        else:
            # Yellow to red (near zero)
            blend = int((norm_val - 0.7) * 333)
            color_def = f"red!{blend}!yellow"
        
        if y_val >= 0:
            strip = f"\\fill[{color_def}] (axis cs:{x1:.4f},0) rectangle (axis cs:{x2:.4f},{y_val:.4f});"
        else:
            strip = f"\\fill[{color_def}] (axis cs:{x1:.4f},{y_val:.4f}) rectangle (axis cs:{x2:.4f},0);"
        colored_strips.append(strip)
    
    strips_code = "\n".join(colored_strips)
    
    # Calculate axis limits
    beam_min = positions.min()
    beam_max = positions.max()
    y_min = bm_min * 1.2 if bm_min < 0 else -10
    y_max = bm_max * 1.2 if bm_max > 0 else 10
    
    # pgfplots code with colored strips
    pgfplots_code = r"""
\begin{figure}[htbp]
\centering
\begin{tikzpicture}
\begin{axis}[
    width=\textwidth,
    height=7cm,
    xlabel={Beam Length (m)},
    ylabel={Bending Moment (kN$\cdot$m)},
    xmin=""" + f"{beam_min:.3f}" + r""",
    xmax=""" + f"{beam_max:.3f}" + r""",
    ymin=""" + f"{y_min:.3f}" + r""",
    ymax=""" + f"{y_max:.3f}" + r""",
    grid=major,
    axis lines=left
]

% Colored vertical strips for gradient effect
""" + strips_code + r"""

% Zero reference line
\addplot[black, very thin, dashed] {0};

% Trajectory curve with markers
\addplot[
    black,
    thick,
    mark=*,
    mark size=1.5pt,
    smooth
] coordinates {
""" + trajectory_coords + r"""
};

\end{axis}
\end{tikzpicture}
\caption{Bending Moment Diagram (BMD) - Contour Visualization}
\label{fig:bmd}
\end{figure}
"""
    
    return pgfplots_code


def build_report(df, beam_image_path, output_name):
    """
    Build the complete PDF report using PyLaTeX.
    
    Args:
        df (pd.DataFrame): Beam force data.
        beam_image_path (str): Path to beam diagram image.
        output_name (str): Output PDF filename (without extension).
    """
    # Create document with geometry settings
    doc = Document(geometry_options={'margin': '1in'})
    
    # Add required packages
    doc.packages.append(Package('tikz'))
    doc.packages.append(Package('pgfplots'))
    doc.packages.append(NoEscape(r'\pgfplotsset{compat=1.18}'))
    doc.packages.append(NoEscape(r'\usepgfplotslibrary{fillbetween}'))
    doc.packages.append(Package('graphicx'))
    doc.packages.append(Package('geometry'))
    doc.packages.append(Package('float'))
    doc.packages.append(Package('hyperref', options='hidelinks'))
    
    # ==================== 1. TITLE PAGE ====================
    doc.preamble.append(Command('title', 'Beam Force Analysis Report'))
    doc.preamble.append(Command('author', 'Engineering Internship Evaluation'))
    doc.preamble.append(Command('date', NoEscape(r'\today')))
    doc.append(NoEscape(r'\maketitle'))
    doc.append(NewPage())
    
    # ==================== 2. TABLE OF CONTENTS ====================
    doc.append(NoEscape(r'\renewcommand{\contentsname}{Contents}'))
    doc.append(NoEscape(r'\tableofcontents'))
    doc.append(NewPage())
    
    # ==================== 3. INTRODUCTION ====================
    with doc.create(Section('Introduction')):
        doc.append('This report presents a comprehensive structural analysis of a simply supported beam ')
        doc.append('subjected to various load conditions. The analysis includes the calculation and ')
        doc.append('visualization of shear force and bending moment distributions along the beam length.')
        doc.append(NoEscape(r'\vspace{0.2cm}'))
        doc.append(NoEscape(r'\\'))
        
        # Embed beam image
        if beam_image_path and Path(beam_image_path).exists():
            with doc.create(Figure(position='htbp')) as fig:
                fig.add_image(beam_image_path, width=NoEscape(r'0.7\textwidth'))
                fig.add_caption('Simply Supported Beam Configuration')
        
        doc.append(NoEscape(r'\vspace{0.2cm}'))
        doc.append('The primary objectives of this analysis are:')
        with doc.create(Subsection('Objectives', numbering=False)):
            doc.append(NoEscape(r'\begin{itemize}'))
            doc.append(NoEscape(r'\item Read and process beam force data from an Excel file'))
            doc.append(NoEscape(r'\item Generate professional engineering diagrams using vector graphics'))
            doc.append(NoEscape(r'\item Present results in a structured, industry-standard format'))
            doc.append(NoEscape(r'\item Provide clear visualization of shear forces and bending moments'))
            doc.append(NoEscape(r'\end{itemize}'))
    
    # ==================== 4. BEAM DESCRIPTION ====================
    with doc.create(Section('Beam Description')):
        doc.append('The structural system under consideration is a simply supported beam. ')
        doc.append('This type of beam is supported at both ends, with one end allowing rotation ')
        doc.append('and horizontal movement (roller support) and the other allowing only rotation (pin support). ')
        doc.append('This configuration allows the beam to freely deform under applied loads while maintaining ')
        doc.append('static equilibrium through the support reactions.')
    
    # ==================== 5. DATA SOURCE ====================
    with doc.create(Section('Data Source')):
        doc.append('The input data for this analysis was sourced from an Excel spreadsheet. ')
        doc.append('The Excel file contains detailed information about load positions and magnitudes ')
        doc.append('applied to the beam structure. Using the pandas library, the data was efficiently ')
        doc.append('extracted and processed for subsequent structural analysis calculations.')
    
    doc.append(NewPage())
    
    # ==================== 6. INPUT DATA TABLE ====================
    create_force_table(doc, df)
    
    # ==================== 7. ANALYSIS ====================
    with doc.create(Section('Analysis')):
        with doc.create(Subsection('Theoretical Background')):
            doc.append(NoEscape(r'\textbf{Shear Force:} '))
            doc.append('The shear force at any section of a beam is defined as the algebraic sum of all ')
            doc.append('vertical forces acting on either side of the section. It represents the internal ')
            doc.append('force that resists shear deformation.')
            doc.append(NoEscape(r'\vspace{0.2cm}'))
            doc.append(NoEscape(r'\\'))
            
            doc.append(NoEscape(r'\textbf{Bending Moment:} '))
            doc.append('The bending moment at any section is the algebraic sum of the moments of all ')
            doc.append('forces acting on either side of the section. It quantifies the internal moment ')
            doc.append('that resists bending of the beam.')
            
        with doc.create(Subsection('Calculation Methodology')):
            doc.append('The shear force and bending moment values are directly extracted from the Excel file, ')
            doc.append('which contains pre-calculated structural analysis results based on fundamental ')
            doc.append('principles of structural mechanics:')
            doc.append(NoEscape(r'\begin{enumerate}'))
            doc.append(NoEscape(r'\item Load positions and magnitudes defined along beam length'))
            doc.append(NoEscape(r'\item Shear force distribution computed from equilibrium of forces'))
            doc.append(NoEscape(r'\item Bending moment distribution calculated through integration'))
            doc.append(NoEscape(r'\item Results verified against structural analysis software'))
            doc.append(NoEscape(r'\end{enumerate}'))
    
    # ==================== 8. SHEAR FORCE DIAGRAM ====================
    positions, shear_forces, bending_moments = calculate_shear_bending_moments(df)
    
    with doc.create(Section('Shear Force Diagram')):
        doc.append('The Shear Force Diagram (SFD) illustrates the variation of shear force along the ')
        doc.append('length of the beam. This diagram is essential for identifying critical sections where ')
        doc.append('shear stress is maximum and for designing adequate shear reinforcement.')
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        
        sfd_tikz = generate_sfd_plot(positions, shear_forces)
        doc.append(NoEscape(sfd_tikz))
    
    # ==================== 9. BENDING MOMENT DIAGRAM ====================
    with doc.create(Section('Bending Moment Diagram')):
        doc.append('The Bending Moment Diagram (BMD) displays the distribution of bending moment along ')
        doc.append('the beam. This diagram is crucial for determining the maximum bending stress and for ')
        doc.append('designing the beam cross-section to resist flexural loads safely.')
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        
        bmd_tikz = generate_bmd_plot(positions, bending_moments)
        doc.append(NoEscape(bmd_tikz))
    
    # ==================== 10. SUMMARY ====================
    with doc.create(Section('Summary')):
        doc.append('This report has presented a complete structural analysis of a simply supported beam, ')
        doc.append('including detailed force calculations and professional visualization of results.')
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        doc.append(NoEscape(r'\\'))
        
        # Summary statistics
        max_shear = np.abs(shear_forces).max()
        max_moment = np.abs(bending_moments).max()
        
        doc.append(NoEscape(r'\textbf{Key Results:}'))
        doc.append(NoEscape(r'\begin{itemize}'))
        doc.append(NoEscape(rf'\item Maximum Shear Force: {max_shear:.2f} kN'))
        doc.append(NoEscape(rf'\item Maximum Bending Moment: {max_moment:.2f} kN$\cdot$m'))
        doc.append(NoEscape(rf'\item Number of Load Points: {len(df)}'))
        doc.append(NoEscape(rf'\item Beam Span: {positions.max():.2f} m'))
        doc.append(NoEscape(r'\end{itemize}'))
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        
        doc.append('The analysis was performed using Python with PyLaTeX for document generation and ')
        doc.append('pgfplots for high-quality vector graphics. All diagrams and tables were generated ')
        doc.append('programmatically to ensure accuracy and reproducibility.')
    
    # Generate PDF
    print(f"\n⚙ Generating PDF report...")
    doc.generate_pdf(output_name, clean_tex=False, compiler='pdflatex', compiler_args=['-interaction=nonstopmode'])
    # Compile a second time to populate table of contents
    doc.generate_pdf(output_name, clean_tex=False, compiler='pdflatex', compiler_args=['-interaction=nonstopmode'])
    print(f"✓ PDF report generated successfully: {output_name}.pdf")


def main():
    """
    Main function to orchestrate the report generation process.
    """
    # ==================== ARGUMENT PARSER ====================
    parser = argparse.ArgumentParser(
        description='Generate professional beam force analysis PDF report using PyLaTeX.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --excel data.xlsx --output beam_report
  %(prog)s --excel forces.xlsx --image beam.png --output analysis
  %(prog)s -e data.xlsx -i diagram.jpg -o report
        """
    )
    
    parser.add_argument(
        '-e', '--excel',
        type=str,
        required=True,
        help='Path to Excel file containing beam force data (required)'
    )
    
    parser.add_argument(
        '-i', '--image',
        type=str,
        default=None,
        help='Path to simply supported beam image (optional)'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        default='beam_analysis_report',
        help='Output PDF filename without extension (default: beam_analysis_report)'
    )
    
    args = parser.parse_args()
    
    # ==================== EXECUTION ====================
    print("=" * 70)
    print("  BEAM FORCE ANALYSIS REPORT GENERATOR")
    print("=" * 70)
    print(f"Excel File:  {args.excel}")
    print(f"Beam Image:  {args.image if args.image else 'Not provided'}")
    print(f"Output Name: {args.output}.pdf")
    print("=" * 70 + "\n")
    
    try:
        # Step 1: Read Excel data
        print("Step 1: Reading Excel data...")
        df = read_excel_data(args.excel)
        
        # Step 2: Validate image path if provided
        if args.image and not Path(args.image).exists():
            print(f"⚠ Warning: Image file not found: {args.image}")
            print("  Proceeding without beam diagram image.")
            args.image = None
        
        # Step 3: Build report
        print("\nStep 2: Building PDF report...")
        build_report(df, args.image, args.output)
        
        print("\n" + "=" * 70)
        print("  ✓ REPORT GENERATION COMPLETED SUCCESSFULLY")
        print("=" * 70)
        
    except Exception as e:
        print("\n" + "=" * 70)
        print("  ✗ ERROR DURING REPORT GENERATION")
        print("=" * 70)
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
