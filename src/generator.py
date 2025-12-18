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

        # Defaults
        self.convert_units = True  # Default to Imperial
        self.include_asset_ids = True

        self.wb = None
        self.ws_master = None

    def get_user_preference(self):
        """Prompts user for output options."""
        print("-" * 50)
        print("GENERATION SETTINGS")
        print("-" * 50)

        print("Input Data contains both Metric and Imperial values.")
        while True:
            response = input("Select Output Units [1=Imperial (Feet), 2=Metric (Meters)]: ").strip()
            if response == '1':
                self.convert_units = True  # Use _ft keys
                print(">> IMPERIAL (Feet/Inches)")
                break
            elif response == '2':
                self.convert_units = False  # Use _m keys
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
                print(">> IDs Excluded")
                break

    def load_template(self):
        if not self.template_path:
            raise FileNotFoundError(f"Template file not found in {Config.INPUT_DIR}")

        print(f"\nLoading template: {os.path.basename(self.template_path)}...")
        self.wb = openpyxl.load_workbook(self.template_path)
        self.ws_master = self.wb[self.wb.sheetnames[0]]

    def _sanitize_sheet_name(self, name):
        return name[:31].replace("/", "-").replace("?", "").replace(":", "")

    def _convert_thickness(self, val):
        """Handles thickness values (Input assumed mm, convert if Imperial selected)."""
        if val is None: return None
        try:
            val = float(val)
            # Assuming Input is always MM or Inches?
            # If standard is MM -> Inches for Imperial
            if self.convert_units:
                result = val / 25.4  # mm to inches
            else:
                result = val  # Keep mm
            return round(result, 3)
        except (ValueError, TypeError):
            return val

    def process_site(self, site_data):
        site_name = self._sanitize_sheet_name(site_data.get('site_name', 'Unknown'))
        print(f"Processing Sheet: {site_name}")

        ws_new = self.wb.copy_worksheet(self.ws_master)
        ws_new.title = site_name

        if hasattr(self.ws_master, '_charts') and self.ws_master._charts:
            try:
                ws_new.add_chart(copy.deepcopy(self.ws_master._charts[0]))
            except:
                pass

        # --- METADATA ---
        ws_new.cell(row=Config.ROW_SITE_NAME, column=Config.COL_SITE_NAME).value = site_name
        ws_new.cell(row=Config.ROW_START_AP_VAL, column=Config.COL_START_AP_VAL).value = site_data.get('ap_id_1')
        ws_new.cell(row=Config.ROW_END_AP_VAL, column=Config.COL_END_AP_VAL).value = site_data.get('ap_id_2')
        ws_new.cell(row=Config.ROW_PIPE_TYPE_VAL, column=Config.COL_PIPE_TYPE_VAL).value = site_data.get('pipe_type')
        ws_new.cell(row=Config.ROW_RESOLUTION_VAL, column=Config.COL_RESOLUTION_VAL).value = site_data.get('resolution')

        # --- DATA TABLE ---
        segments = site_data.get('segments', [])

        # Clean up template placeholder data
        for r in range(Config.DATA_START_ROW, Config.DATA_START_ROW + 500):
            for c in range(1, 9): ws_new.cell(row=r, column=c).value = None

        current_row = Config.DATA_START_ROW
        for seg in segments:
            ws_new.cell(row=current_row, column=Config.COL_ACCESS_POINT_LBL).value = seg.get('access_point_label')

            # --- UNIT SELECTION LOGIC ---
            if self.convert_units:
                start_val = seg.get('start_ft')
                end_val = seg.get('end_ft')
            else:
                start_val = seg.get('start_m')
                end_val = seg.get('end_m')

            ws_new.cell(row=current_row, column=Config.COL_START_FT).value = round(float(start_val),
                                                                                   3) if start_val is not None else None
            ws_new.cell(row=current_row, column=Config.COL_END_FT).value = round(float(end_val),
                                                                                 3) if end_val is not None else None

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
        if self.ws_master.title not in sanitized_names:
            del self.wb[self.ws_master.title]

        if not os.path.exists(self.output_dir): os.makedirs(self.output_dir)
        full_path = os.path.join(self.output_dir, f"{self.project_name}.xlsx")
        print(f"\nSaving results to: {full_path}")
        self.wb.save(full_path)
        print("Generation Complete.")

    def run(self, data):
        self.get_user_preference()
        self.load_template()
        for site in data: self.process_site(site)
        self.save([s['site_name'] for s in data])
