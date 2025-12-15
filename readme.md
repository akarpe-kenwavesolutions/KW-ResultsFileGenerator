# KW Results Generator

A Python utility to automate the generation of "Results Files" for pipeline inspection data. This tool takes a master Excel template (containing formatted charts and styles) and populates it with site-specific segment data, automatically resizing charts and preserving print formatting.

## Features
- **Template-Based Generation**: Preserves all formatting, print settings, and chart styles from a master Excel file.
- **Dynamic Chart Resizing**: Automatically adjusts chart data ranges to fit the exact number of data points for each site.
- **Unit Conversion**: Optional command-line prompt to convert thickness values from Millimeters (mm) to Inches (in).
- **Batch Processing**: Generates a multi-tab workbook with one sheet per site.
- **OOP Structure**: Modular, object-oriented design for easy maintenance and extension.

## Prerequisites
- Python 3.8+
- [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel) (for viewing output)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/KW-Results-Generator.git
   cd KW-Results-Generator
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Prepare Inputs**:
   - Ensure your master template file is named `KW-Results-Gresham-SL-Team-Draft.xlsx` and placed in the `input/` directory.

2. **Run the Script**:
   ```bash
   python src/main.py
   ```

3. **Follow Prompts**:
   - The script will ask: `Do you want to convert input Thickness values from mm to inches? (y/n)`
   - Enter `y` to divide all thickness values by 25.4.
   - Enter `n` to keep values as they are.

4. **View Output**:
   - The generated file will be saved to `output/KW-Results-Final_Generated.xlsx`.

## Project Structure
- `src/`: Source code directory.
  - `main.py`: Entry point for the application.
  - `generator.py`: Contains the `ResultsGenerator` class with core logic.
  - `utils.py`: Helper functions for user input and file handling.
- `input/`: Directory for the master template.
- `output/`: Directory where the final results file is saved.
