import os
import glob


class Config:
    # Base Paths
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    PROJECTS_ROOT = os.path.join(BASE_DIR, 'projects')

    # Defaults (will be updated by set_project)
    INPUT_DIR = os.path.join(BASE_DIR, 'input')
    OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

    # Files (Placeholder)
    FILE_SEG_DF = None
    FILE_PIPE_ASSETS = None
    FILE_PIPE_SPECS = None
    TEMPLATE_PATH = None

    # Column Mappings
    KEY_GROUP = 'seg_group'
    KEY_START = 'start_loc'
    KEY_END = 'end_loc'
    KEY_ASSET_ID = 'pipe_asset_id'

    # Meta Keys
    KEY_META_AP1 = 'Access Point ID 1'
    KEY_META_AP2 = 'Access Point ID 2'
    KEY_META_DIA = 'Pipe Diameter'
    KEY_META_MAT = 'Material'

    # Output Columns
    COL_ACCESS_POINT_LBL = 1
    COL_START_FT = 2
    COL_END_FT = 3
    COL_X_AXIS = 4  # Hidden D column
    COL_DRI_THICKNESS = 5
    COL_NOM_THICKNESS = 6
    COL_ASSET_ID = 7
    COL_RESOLUTION = 8
    COL_PIPE_TYPE = 5
    DATA_START_ROW = 9

    @classmethod
    def find_input_file(cls, directory, substring, extensions=['.csv', '.xlsx']):
        """Helper to find a file containing a substring in a directory."""
        if not os.path.exists(directory):
            return None
        for file in os.listdir(directory):
            if substring in file and any(file.endswith(ext) for ext in extensions):
                return os.path.join(directory, file)
        return None

    @classmethod
    def set_project(cls, project_name):
        """Updates Input/Output paths based on Project Name."""
        # 1. Set Directories
        project_path = os.path.join(cls.PROJECTS_ROOT, project_name)
        cls.INPUT_DIR = os.path.join(project_path, 'input')
        cls.OUTPUT_DIR = os.path.join(project_path, 'output')

        # 2. Create Structure if missing
        if not os.path.exists(cls.INPUT_DIR):
            try:
                os.makedirs(cls.INPUT_DIR)
                os.makedirs(cls.OUTPUT_DIR)
                print(f"\n[SETUP] Created new project folders at:\n  {project_path}")
                print("  Please place your Input Files (CSVs and Template) in the 'input' folder and run again.")
                exit(0)  # Stop here so user can add files
            except OSError as e:
                print(f"[ERROR] Could not create project directories: {e}")
                exit(1)

        # 3. Find Files in new Input Dir
        cls.FILE_SEG_DF = cls.find_input_file(cls.INPUT_DIR, 'seg_df')
        cls.FILE_PIPE_ASSETS = cls.find_input_file(cls.INPUT_DIR, 'pipe_asset_ids_df')
        cls.FILE_PIPE_SPECS = cls.find_input_file(cls.INPUT_DIR, 'pipe_spec_df')
        # Template can be 'Draft' or 'Skeleton'
        cls.TEMPLATE_PATH = cls.find_input_file(cls.INPUT_DIR, 'Draft', ['.xlsx'])
        if not cls.TEMPLATE_PATH:
            cls.TEMPLATE_PATH = cls.find_input_file(cls.INPUT_DIR, 'Skeleton', ['.xlsx'])
