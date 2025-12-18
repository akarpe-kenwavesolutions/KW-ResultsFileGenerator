import os
import glob


class Config:
    # ---------------------------------------------------------
    # DIRECTORIES
    # ---------------------------------------------------------
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    PROJECTS_ROOT = os.path.join(BASE_DIR, 'projects')

    # ---------------------------------------------------------
    # DYNAMIC PATHS
    # ---------------------------------------------------------
    PROJECT_DIR = None
    INPUT_DIR = None
    OUTPUT_DIR = None

    # ---------------------------------------------------------
    # FILE PATHS
    # ---------------------------------------------------------
    FILE_SEG_DF = None
    FILE_PIPE_ASSETS = None
    FILE_PIPE_SPECS = None
    FILE_FIELD_NOTES = None
    FILE_SUMMARY = None
    FILE_SEG_GROUPS = None
    TEMPLATE_PATH = None

    # ---------------------------------------------------------
    # EXCEL REPORT HEADER SETTINGS
    # ---------------------------------------------------------
    # Column indices (1-based) for writing header info (e.g., cell B5)
    COL_SITE_NAME = 2
    COL_PIPE_TYPE = 2
    COL_AP_NAMES = 2
    COL_START_AP = 2
    COL_END_AP = 2
    COL_DATE = 2
    COL_RESOLUTION = 2

    # Row indices for where to write header info
    ROW_SITE_NAME = 3
    ROW_PIPE_TYPE = 5
    ROW_AP_NAMES = 4
    ROW_DATE = 7
    ROW_RESOLUTION = 6

    # ---------------------------------------------------------
    # EXCEL REPORT DATA TABLE SETTINGS
    # ---------------------------------------------------------
    # Rows
    TABLE_START_ROW = 10  # Standard name
    DATA_START_ROW = 10  # Alias for generator.py compatibility

    # Columns (1-based index)
    # A=1, B=2, C=3, D=4, E=5, F=6, G=7
    COL_ACCESS_POINT_LBL = 1  # Column A
    COL_START_FT = 2  # Column B
    COL_END_FT = 3  # Column C
    # Column 4 (D) is usually skipped/spacer or Avg Thickness
    COL_DRI_THICKNESS = 5  # Column E
    COL_NOM_THICKNESS = 6  # Column F
    COL_ASSET_ID = 7  # Column G

    # ---------------------------------------------------------
    # PLOT SETTINGS
    # ---------------------------------------------------------
    PLOT_HEIGHT = 9.5
    PLOT_WIDTH = 14
    DPI = 100

    # ---------------------------------------------------------
    # CSV KEYS & LOGIC SETTINGS
    # ---------------------------------------------------------
    REQUIRE_ASSET_IDS = True

    # Metadata Keys (seg_df)
    KEY_META_AP1 = 'Access Point 1'
    KEY_META_AP2 = 'Access Point 2'
    KEY_META_MAT = 'Material'
    KEY_META_DIA = 'Diameter'

    # Asset Keys (pipe_asset_ids_df)
    KEY_GROUP = 'seg_group'
    KEY_START = 'start_loc'
    KEY_END = 'end_loc'
    KEY_ASSET_ID = 'pipe_asset_id'

    # Measurement Keys (summary file)
    KEY_MEASURE_POS = 'ap_ex_pos'
    KEY_MEASURE_VAL = 'avg_wall_thickness'

    # ---------------------------------------------------------
    # SETUP METHODS
    # ---------------------------------------------------------
    @classmethod
    def set_project(cls, project_identifier):
        # 1. Locate Project
        if os.path.isdir(project_identifier):
            cls.PROJECT_DIR = project_identifier
            project_name = os.path.basename(project_identifier)
        elif os.path.exists(os.path.join(cls.PROJECTS_ROOT, project_identifier)):
            cls.PROJECT_DIR = os.path.join(cls.PROJECTS_ROOT, project_identifier)
            project_name = project_identifier
        else:
            if not os.path.exists(cls.PROJECTS_ROOT):
                cls.PROJECTS_ROOT = os.path.join(cls.BASE_DIR, 'input')

            candidates = [d for d in os.listdir(cls.PROJECTS_ROOT)
                          if os.path.isdir(os.path.join(cls.PROJECTS_ROOT, d))
                          and project_identifier.lower() in d.lower()]

            if len(candidates) == 1:
                print(f"Auto-detected project folder: '{candidates[0]}'")
                cls.PROJECT_DIR = os.path.join(cls.PROJECTS_ROOT, candidates[0])
                project_name = candidates[0]
            elif len(candidates) > 1:
                raise ValueError(f"Ambiguous project name. Found: {candidates}")
            else:
                raise FileNotFoundError(f"Project '{project_identifier}' not found in {cls.PROJECTS_ROOT}")

        # 2. Set Paths
        cls.INPUT_DIR = os.path.join(cls.PROJECT_DIR, 'input')
        cls.OUTPUT_DIR = os.path.join(cls.PROJECT_DIR, 'output')

        if not os.path.exists(cls.INPUT_DIR):
            raise FileNotFoundError(f"Input directory not found: {cls.INPUT_DIR}")
        if not os.path.exists(cls.OUTPUT_DIR):
            os.makedirs(cls.OUTPUT_DIR)

        # 3. Find Files
        cls.FILE_SEG_DF = cls.find_input_file(cls.INPUT_DIR, 'seg_df')
        cls.FILE_PIPE_ASSETS = cls.find_input_file(cls.INPUT_DIR, 'pipe_asset_ids_df')
        cls.FILE_PIPE_SPECS = cls.find_input_file(cls.INPUT_DIR, 'pipe_spec_df')
        cls.FILE_FIELD_NOTES = cls.find_input_file(cls.INPUT_DIR, 'fieldNotes_df')
        cls.FILE_SUMMARY = cls.find_input_file(cls.INPUT_DIR, 'summary', extensions=['.xlsx'])
        cls.FILE_SEG_GROUPS = cls.find_input_file(cls.INPUT_DIR, 'segGroups', extensions=['.xlsx'])
        cls.TEMPLATE_PATH = cls.find_input_file(cls.INPUT_DIR, 'template', extensions=['.xlsx', '.xlsm'])

        print(f"--- Configuration for '{project_name}' ---")
        print(f"Input Dir:     {cls.INPUT_DIR}")
        print(
            f"Assets File:   {os.path.basename(cls.FILE_PIPE_ASSETS) if cls.FILE_PIPE_ASSETS else 'MISSING (Method 1)'}")
        print(f"Summary File:  {os.path.basename(cls.FILE_SUMMARY) if cls.FILE_SUMMARY else 'MISSING (Method 2)'}")
        print(
            f"SegGroups:     {os.path.basename(cls.FILE_SEG_GROUPS) if cls.FILE_SEG_GROUPS else 'MISSING (Method 2)'}")
        print(f"Template:      {os.path.basename(cls.TEMPLATE_PATH) if cls.TEMPLATE_PATH else 'None'}")
        print("------------------------------------------\n")

    @classmethod
    def find_input_file(cls, directory, keyword, extensions=['.csv']):
        try:
            files = os.listdir(directory)
            for f in files:
                if not any(f.endswith(ext) for ext in extensions): continue
                if keyword.lower() in f.lower():
                    return os.path.join(directory, f)
            return None
        except OSError:
            return None
