# KW Results Generator

A Python utility to automate the generation of Results Files for pipeline inspection data. The tool takes a master Excel template and populates it with site-specific segment data, automatically resizing charts and preserving print formatting.

## Features
- **Template-Based Generation**: Uses a single master template (`src/Master_Results_Skeleton_Template.xlsx`) shared across all projects — no per-project copy needed.
- **Dynamic Chart Resizing**: Automatically adjusts chart data ranges to fit the exact number of data points for each site.
- **Unit Conversion**: Optional prompt to convert thickness values from mm to inches.
- **Batch Processing**: Generates a multi-tab workbook with one sheet per group.
- **Consistent Group Length**: Group length grid truncates to the last complete 2m step (matching summary file behaviour).
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

1. **Prepare Inputs**: Place the following files in your project's `Dataframes/` folder:
   - `*_seg_df.csv` — segment definitions
   - `*_pipe_asset_ids_df.csv` — pipe asset IDs (optional)
   - `*_segGroups.xlsx` — segment group assignments

2. **Run the Script**:
   ```bash
   python src/main.py
   ```

3. **Follow Prompts**:
   - Enter the project name (must match a folder under `projects/`).
   - Choose whether to include Pipe Asset IDs.
   - Choose unit system (Imperial ft/in or Metric m/mm).

4. **View Output**:
   - The generated file is saved to `projects/<project name>/output/`.

## Project Structure

```
ResultsFile/
├── src/
│   ├── main.py                            # Entry point
│   ├── config.py                          # Path and settings configuration
│   ├── data_loader.py                     # Loads and processes segment data
│   ├── generator.py                       # Builds the Excel results file
│   ├── chart_manager.py                   # Handles chart resizing
│   └── Master_Results_Skeleton_Template.xlsx  # Single shared template (all projects)
└── projects/
    └── <Project Name>/
        ├── Dataframes/                    # Input dataframes for this project
        │   ├── *_seg_df.csv
        │   ├── *_pipe_asset_ids_df.csv    # Optional
        │   └── *_segGroups.xlsx
        └── output/                        # Generated results files
```

## Notes
- The master template lives in `src/` and is referenced directly — no need to copy it into each project folder.
- New projects only need a `Dataframes/` folder with the required CSVs and the segGroups Excel file.
- The `output/` folder is created automatically if it does not exist.
