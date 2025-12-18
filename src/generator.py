import openpyxl
import os
import copy
from config import Config
from chart_manager import ChartManager


class ResultsGenerator:
    def __init__(self, project_name="KW-Results"):
        self.output_dir = Config.OUTPUT_DIR
        self.template_path = Config.TEMPLATE_PATH
        self.project_name = f"KW-Results-{project_name}"

        # Settings
        self.convert_units = False  # Determines if we label it Imperial/Metric
        self.include_asset_ids = True

        self.wb = None
        self.ws_master = None
        self.master_sheet_name = None

    def get_user_preference(self):
        """Prompts user for options."""
        print("-" * 50)
        print("GENERATION SETTINGS")
        print("-" * 50)

        print("Input Data is assumed to be in METRIC (Meters/mm).")
        while True:
            response = input("Select Output Units [1=Imperial, 2=Metric]: ").strip()
            if response == '1':
                self.convert_units = True
                print(">> IMPERIAL (Feet/Inches)")
                break
            elif response == '2':
                self.convert_units = False
                print(">> METRIC (Meters/mm)")
                break

        print("-" * 20)
        while True:
            resp_id = input("Include Pipe Asset IDs in Output? [y/n]: ").strip().lower()
            if resp_id in ['y', 'yes']:
                self.include_asset_ids = True
                print(">> IDs Included")
                break
            elif resp_id in ['n', 'no']:
                self.include_asset_ids = False
                print(">> IDs Excluded (Column G will be empty)")
                break

    def load_template(self):
        if not self.template_path:
            raise FileNotFoundError(f"Template file not found in {Config.INPUT_DIR}")

        print(f"\nLoading template: {os.path.basename(self.template_path)}...")
        self.wb = openpyxl.load_workbook(self.template_path)

        if 'Template' in self.wb.sheetnames:
            self.master_sheet_name = 'Template'
        else:
            self.master_sheet_name = self.wb.sheetnames[0]

        self.ws_master = self.wb[self.master_sheet_name]

    def _sanitize_sheet_name(self, name):
        return name[:31].replace("/", "-").replace("?", "").replace(":", "")

    def _convert_distance(self, val):
        """
        Handles distance values (Start/End).
        DataLoader now returns these ALREADY converted to feet if Imperial.
        So we just round them.
        """
        if val is None: return None
        try:
            val = float(val)
            # DataLoader provides feet, so no conversion needed here if units match
            # If user selected Metric (2), DataLoader might need adjustment or
            # we assume DataLoader provides Feet and we convert BACK to meters?
            # Based on current setup: DataLoader provides FEET.

            if not self.convert_units:
                # User wants Metric, but Data is in Feet from DataLoader
                # Convert Feet -> Meters
                result = val * 0.3048
            else:
                # User wants Imperial, Data is in Feet
                result = val

            return round(result, 3)
        except (ValueError, TypeError):
            return val

    def _convert_thickness(self, val):
        """
        Handles thickness values (Input assumed mm).
        """
        if val is None: return None
        try:
            val = float(val)
            if self.convert_units:
                # mm to inches
                result = val / 25.4
            else:
                result = val
            return round(result, 3)
        except (ValueError, TypeError):
            return val

    def process_site(self, site_data):
        site_name = self._sanitize_sheet_name(site_data.get('site_name', 'Unknown'))
        print(f"Processing Sheet: {site_name}")

        ws_new = self.wb.copy_worksheet(self.ws_master)
        ws_new.title = site_name

        # Manual Chart Copy (OpenPyXL workaround)
        if hasattr(self.ws_master, '_charts') and self.ws_master._charts:
            try:
                original_chart = self.ws_master._charts[0]
                new_chart = copy.deepcopy(original_chart)
                ws_new.add_chart(new_chart)
            except Exception as e:
                print(f"  [WARNING] Could not copy chart: {e}")

        # Write Metadata
        ws_new.cell(row=Config.ROW_AP_NAMES, column=2).value = site_data.get('ap_id_1')
        ws_new.cell(row=Config.ROW_AP_NAMES, column=3).value = site_data.get('ap_id_2')
        ws_new.cell(row=Config.ROW_PIPE_TYPE, column=Config.COL_PIPE_TYPE).value = site_data.get('pipe_type')
        ws_new.cell(row=Config.ROW_RESOLUTION, column=Config.COL_RESOLUTION).value = site_data.get('resolution')

        # Write Segments
        segments = site_data.get('segments', [])

        # Clear data area
        # Using 1000 as a safe upper limit for clearing
        for row in ws_new.iter_rows(min_row=Config.DATA_START_ROW, max_row=Config.DATA_START_ROW + 1000, min_col=1,
                                    max_col=8):
            for cell in row:
                cell.value = None

        current_row = Config.DATA_START_ROW
        for seg in segments:
            ws_new.cell(row=current_row, column=Config.COL_ACCESS_POINT_LBL).value = seg.get('access_point_label')
            ws_new.cell(row=current_row, column=Config.COL_START_FT).value = self._convert_distance(seg.get('start_ft'))
            ws_new.cell(row=current_row, column=Config.COL_END_FT).value = self._convert_distance(seg.get('end_ft'))
            ws_new.cell(row=current_row, column=Config.COL_DRI_THICKNESS).value = self._convert_thickness(
                seg.get('dri_thickness'))
            ws_new.cell(row=current_row, column=Config.COL_NOM_THICKNESS).value = self._convert_thickness(
                seg.get('nom_thickness'))

            if self.include_asset_ids:
                ws_new.cell(row=current_row, column=Config.COL_ASSET_ID).value = seg.get('pipe_asset_id')

            current_row += 1

        ChartManager.update_chart_range(ws_new, len(segments))

    def save(self, input_site_names):
        sanitized_names = [self._sanitize_sheet_name(n) for n in input_site_names]
        if self.master_sheet_name not in sanitized_names:
            del self.wb[self.master_sheet_name]

        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        filename = f"{self.project_name}.xlsx"
        full_path = os.path.join(self.output_dir, filename)

        print(f"\nSaving results to: {full_path}")
        self.wb.save(full_path)
        print("Generation Complete.")

    def run(self, data):
        self.get_user_preference()
        self.load_template()
        for site in data:
            self.process_site(site)
        self.save([s['site_name'] for s in data])
