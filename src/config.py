import os


class Config:
    # ---------------------------------------------------------
    # DIRECTORIES
    # ---------------------------------------------------------
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    PROJECTS_ROOT = os.path.join(BASE_DIR, 'projects')
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
    # EXCEL REPORT HEADER SETTINGS (from KW-Results-RJN-Lake-Forest-Draft.xlsx)
    # ---------------------------------------------------------
    # --- Cell coordinates for WRITING VALUES ---
    ROW_SITE_NAME = 1  # Example: 'LakeForest-Waukegan' in A1 or merged cell
    COL_SITE_NAME = 1

    ROW_START_AP_VAL = 6  # Value for Start AP in B6
    COL_START_AP_VAL = 2

    ROW_END_AP_VAL = 6  # Value for End AP in C6
    COL_END_AP_VAL = 3

    ROW_PIPE_TYPE_VAL = 5  # Value for Pipe Type in G5
    COL_PIPE_TYPE_VAL = 7

    ROW_RESOLUTION_VAL = 6  # Value for Resolution in G6
    COL_RESOLUTION_VAL = 7

    ROW_DATE = 1  # Placeholder
    COL_DATE = 1

    # ---------------------------------------------------------
    # EXCEL REPORT DATA TABLE SETTINGS
    # ---------------------------------------------------------
    TABLE_HEADER_ROW = 8  # "Access Points", "Start (ft)" headers are on Row 8
    DATA_START_ROW = 9  # Actual data begins on Row 9

    # Columns (1-based index)
    COL_ACCESS_POINT_LBL = 1  # A: Access Points
    COL_START_FT = 2  # B: Start (ft)
    COL_END_FT = 3  # C: End (ft)
    COL_DRI_THICKNESS = 5  # E: DRI Thickness
    COL_NOM_THICKNESS = 6  # F: Nominal Thickness
    COL_ASSET_ID = 7  # G: Pipe Asset ID

    # ---------------------------------------------------------
    # LOGIC & CSV KEYS
    # ---------------------------------------------------------
    REQUIRE_ASSET_IDS = True
    KEY_META_AP1 = 'Access Point 1'
    KEY_META_AP2 = 'Access Point 2'
    KEY_META_MAT = 'Material'
    KEY_META_DIA = 'Diameter'
    KEY_GROUP = 'seg_group'
    KEY_START = 'start_loc'
    KEY_END = 'end_loc'
    KEY_ASSET_ID = 'pipe_asset_id'
    KEY_MEASURE_POS = 'ap_ex_pos'
    KEY_MEASURE_VAL = 'avg_wall_thickness'

    @classmethod
    def set_project(cls, project_identifier):
        if os.path.isdir(project_identifier):
            cls.PROJECT_DIR = project_identifier
        elif os.path.exists(os.path.join(cls.PROJECTS_ROOT, project_identifier)):
            cls.PROJECT_DIR = os.path.join(cls.PROJECTS_ROOT, project_identifier)
        else:
            if not os.path.exists(cls.PROJECTS_ROOT): cls.PROJECTS_ROOT = os.path.join(cls.BASE_DIR, 'input')
            candidates = [d for d in os.listdir(cls.PROJECTS_ROOT) if project_identifier.lower() in d.lower()]
            if len(candidates) == 1:
                cls.PROJECT_DIR = os.path.join(cls.PROJECTS_ROOT, candidates[0])
            else:
                raise FileNotFoundError("Project not found")

        cls.INPUT_DIR = os.path.join(cls.PROJECT_DIR, 'input')
        cls.OUTPUT_DIR = os.path.join(cls.PROJECT_DIR, 'output')
        if not os.path.exists(cls.OUTPUT_DIR): os.makedirs(cls.OUTPUT_DIR)

        cls.FILE_SEG_DF = cls.find_input_file(cls.INPUT_DIR, 'seg_df')
        cls.FILE_PIPE_ASSETS = cls.find_input_file(cls.INPUT_DIR, 'pipe_asset_ids_df')
        cls.FILE_SEG_GROUPS = cls.find_input_file(cls.INPUT_DIR, 'segGroups', extensions=['.xlsx'])
        cls.TEMPLATE_PATH = cls.find_input_file(cls.INPUT_DIR, 'template', extensions=['.xlsx', '.xlsm'])

    @classmethod
    def find_input_file(cls, directory, keyword, extensions=['.csv']):
        try:
            for f in os.listdir(directory):
                if any(f.endswith(ext) for ext in extensions) and keyword.lower() in f.lower(): return os.path.join(
                    directory, f)
        except:
            return None
