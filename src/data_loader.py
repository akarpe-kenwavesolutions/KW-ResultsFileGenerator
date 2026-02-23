import pandas as pd
import os
import re
import numpy as np
from config import Config


class DataLoader:
    def __init__(self):
        self.seg_df_path = Config.FILE_SEG_DF
        self.asset_path = Config.FILE_PIPE_ASSETS
        self.seg_groups_path = Config.FILE_SEG_GROUPS
        self.INCREMENT_METERS = 2.0
        self.FEET_PER_METER = 3.281

    def _get_val_robust(self, row, keys_list):
        for key in keys_list:
            if key in row.index:
                val = row[key]
                if pd.notna(val) and str(val).strip().lower() != 'nan':
                    return val
        return None

    def _get_column_name(self, df, candidates):
        for col in df.columns:
            if col in candidates: return col
            for cand in candidates:
                if col.lower() == cand.lower(): return col
        return None

    def _format_ap_name(self, val):
        """Parse access point names as strings, removing decimal points from numeric values."""
        if val is None:
            return None
        try:
            f = float(val)
            if f == int(f):
                return str(int(f))
            return str(f)
        except (ValueError, TypeError):
            return str(val)

    def _order_segments_for_group(self, df_seg_meta, segment_codes):
        """
        Orders segments based on start and end positions.
        Returns ordered segment string like "S01 - S29 - S02 - S07..."
        """
        # Filter for segments in this group
        group_segs = df_seg_meta[df_seg_meta['Unnamed: 0'].isin(segment_codes)].copy()

        if group_segs.empty:
            return ""

        # Extract segment IDs and positions
        df_working = group_segs[['Unnamed: 0', 'ap_1_loc', 'ap_2_loc']].copy()
        df_working.columns = ['Segment', 'Start_Pos', 'End_Pos']

        # Remove rows with missing positions
        df_working = df_working.dropna(subset=['Start_Pos', 'End_Pos'])

        if df_working.empty:
            return ""

        # Sort by Start_Pos (ascending), then by End_Pos (ascending)
        df_sorted = df_working.sort_values(['Start_Pos', 'End_Pos'], ascending=[True, True])

        # Create the ordered string
        ordered_string = ' - '.join(df_sorted['Segment'].tolist())

        return ordered_string

    def _extract_all_pipe_specs(self, df_seg_meta, segment_codes):
        """
        Extract pipe specifications following the ORDERED segment sequence.
        Uses ps_change_locs column to determine actual spec change positions.
        Only reports specs when there's a genuine transition (not continuous pipe).
        Returns list of (diameter, material) tuples representing actual spec changes.
        """
        # Get segments in correct order
        group_segs = df_seg_meta[df_seg_meta['Unnamed: 0'].isin(segment_codes)].copy()
        if group_segs.empty:
            return []

        # Sort by start position to get correct sequence
        df_working = group_segs[['Unnamed: 0', 'ap_1_loc', 'ap_2_loc']].copy()
        df_working.columns = ['Segment', 'Start_Pos', 'End_Pos']
        df_working = df_working.dropna(subset=['Start_Pos', 'End_Pos'])

        if df_working.empty:
            return []

        df_sorted = df_working.sort_values(['Start_Pos', 'End_Pos'], ascending=[True, True])
        ordered_segment_ids = df_sorted['Segment'].tolist()

        # Build list of (absolute_position, diameter, material) tuples
        all_transitions = []

        for seg_id in ordered_segment_ids:
            seg_row = df_seg_meta[df_seg_meta['Unnamed: 0'] == seg_id]
            if seg_row.empty:
                continue

            row = seg_row.iloc[0]

            # Get segment start position
            seg_start_pos = row.get('ap_1_loc')
            if pd.isna(seg_start_pos):
                continue

            seg_start_pos = float(seg_start_pos)

            # Get diameter and material
            diameter_raw = self._get_val_robust(row, [Config.KEY_META_DIA, 'diameter', 'Diameter'])
            material_raw = self._get_val_robust(row, [Config.KEY_META_MAT, 'material', 'Material'])

            # Get ps_change_locs (pipe spec change locations)
            ps_change_locs_raw = row.get('ps_change_locs')

            # Split specs by '/'
            diameters = []
            materials = []

            if diameter_raw:
                dia_str = str(diameter_raw)
                diameters = [d.strip() for d in dia_str.split('/') if d.strip()]

            if material_raw:
                mat_str = str(material_raw)
                materials = [m.strip() for m in mat_str.split('/') if m.strip()]

            # Format diameters (remove .0)
            formatted_diameters = []
            for dia in diameters:
                try:
                    dia_float = float(dia)
                    if dia_float == int(dia_float):
                        formatted_diameters.append(str(int(dia_float)))
                    else:
                        formatted_diameters.append(dia)
                except:
                    formatted_diameters.append(dia)

            # Parse ps_change_locs
            change_positions = []
            if pd.notna(ps_change_locs_raw) and str(ps_change_locs_raw).strip():
                ps_str = str(ps_change_locs_raw)
                try:
                    # Handle both single values and comma-separated lists
                    if ',' in ps_str:
                        change_positions = [float(x.strip()) for x in ps_str.split(',') if x.strip()]
                    else:
                        change_positions = [float(ps_str)]
                except:
                    pass

            # Build transitions for this segment
            if not change_positions:
                # No mid-segment changes - segment has single spec throughout
                if formatted_diameters and materials:
                    dia = formatted_diameters[0]
                    mat = materials[0]
                    all_transitions.append((seg_start_pos, dia, mat))
            else:
                # Has mid-segment changes
                # First spec starts at segment start
                if formatted_diameters and materials:
                    all_transitions.append((seg_start_pos, formatted_diameters[0], materials[0]))

                # Add transitions at each change location
                for i, relative_pos in enumerate(change_positions):
                    absolute_pos = seg_start_pos + relative_pos
                    spec_index = i + 1  # Next spec after the change

                    if spec_index < len(formatted_diameters) and spec_index < len(materials):
                        dia = formatted_diameters[spec_index]
                        mat = materials[spec_index]
                        all_transitions.append((absolute_pos, dia, mat))

        # Sort by position
        all_transitions.sort(key=lambda x: x[0])

        # Remove consecutive duplicates (same spec at different positions = continuous pipe)
        deduplicated = []
        prev_spec = None

        for pos, dia, mat in all_transitions:
            current_spec = (dia, mat)
            if current_spec != prev_spec:
                deduplicated.append(current_spec)
                prev_spec = current_spec

        return deduplicated

    def load_seg_groups(self):
        if not self.seg_groups_path: return None
        try:
            xls = pd.ExcelFile(self.seg_groups_path)
            sheet_name = xls.sheet_names[0]
            df_groups = pd.read_excel(self.seg_groups_path, sheet_name=sheet_name, index_col=0)
            seg_groups_dict = {}
            for group_name, row in df_groups.iterrows():
                segments_str = row.get('segments', '[]')
                if isinstance(segments_str, str):
                    import ast
                    try:
                        seg_list = ast.literal_eval(segments_str)
                    except:
                        seg_list = [segments_str] if segments_str else []
                else:
                    seg_list = segments_str if isinstance(segments_str, list) else []
                seg_groups_dict[group_name] = seg_list
            return seg_groups_dict
        except Exception as e:
            print(f"WARNING: Could not load segGroups: {e}")
            return None

    def load_access_points_for_group(self, group_name):
        if not self.seg_groups_path: return {}
        try:
            xls = pd.ExcelFile(self.seg_groups_path)
            if group_name in xls.sheet_names:
                df_aps = pd.read_excel(self.seg_groups_path, sheet_name=group_name)
                ap_dict = {}
                for _, row in df_aps.iterrows():
                    ap_name = row.get('ap_name')
                    position = row.get('position')
                    if ap_name and pd.notna(position):
                        ap_dict[str(ap_name).strip()] = float(position)
                return ap_dict
        except:
            pass
        return {}

    def load_data(self):
        # Validate required files
        missing_files = []

        if not self.seg_df_path:
            missing_files.append("Segment DF")

        # Only require asset file if user requested pipe asset IDs
        if Config.REQUIRE_ASSET_IDS and not self.asset_path:
            missing_files.append("Asset IDs DF")

        if missing_files:
            raise FileNotFoundError(f"Missing Input Files:\n- " + "\n- ".join(missing_files))

        # Load files
        print(f"Loading Metadata: {os.path.basename(self.seg_df_path)}")

        # Only load asset file if it exists AND is required
        if self.asset_path and Config.REQUIRE_ASSET_IDS:
            print(f"Loading Assets: {os.path.basename(self.asset_path)}")
            df_assets = pd.read_csv(self.asset_path)
        else:
            if not Config.REQUIRE_ASSET_IDS:
                print("Skipping Pipe Asset IDs (not required for this project)")
            df_assets = pd.DataFrame()  # Empty dataframe

        if self.seg_groups_path:
            print(f"Loading SegGroups: {os.path.basename(self.seg_groups_path)}")

        df_seg_meta = pd.read_csv(self.seg_df_path)
        seg_groups_dict = self.load_seg_groups()

        if not seg_groups_dict:
            raise FileNotFoundError("segGroups file is required.")

        grouped_data = []
        print(f"Processing {len(seg_groups_dict)} sites...")

        for group_id, segment_codes in seg_groups_dict.items():
            if not segment_codes or (len(segment_codes) == 1 and segment_codes[0] == ''):
                continue

            # --- A. Filter Metadata ---
            group_meta_rows = df_seg_meta[df_seg_meta['Unnamed: 0'].isin(segment_codes)].copy()
            if group_meta_rows.empty:
                continue

            # --- NEW: Calculate ordered segment string ---
            ordered_segments = self._order_segments_for_group(df_seg_meta, segment_codes)

            # --- NEW: Extract ALL pipe specs (not just first one) ---
            pipe_specs_list = self._extract_all_pipe_specs(df_seg_meta, segment_codes)

            # Pipe Type Extraction (for compatibility/fallback)
            pipe_type = "Unknown"
            best_dia = ""
            best_mat = ""
            for _, row in group_meta_rows.iterrows():
                d = self._get_val_robust(row, [Config.KEY_META_DIA, 'diameter', 'Diameter'])
                m = self._get_val_robust(row, [Config.KEY_META_MAT, 'material', 'Material'])
                if d and not best_dia: best_dia = str(d)
                if m and not best_mat: best_mat = str(m)
                if best_dia and best_mat: break
            if best_dia or best_mat:
                pipe_type = f"{best_dia} {best_mat}".strip()

            # --- B. Load Access Points (Source of Truth for Distances) ---
            ap_positions_dict = self.load_access_points_for_group(group_id)
            unique_aps = sorted([(pos, name) for name, pos in ap_positions_dict.items()], key=lambda x: x[0])

            if not unique_aps:
                print(f"  Skipping {group_id} (No Access Points found)")
                continue

            # --- C. Load Assets (only if required) ---
            asset_group_col = None
            group_asset_rows = pd.DataFrame()

            if not df_assets.empty:
                asset_group_col = self._get_column_name(df_assets, [Config.KEY_GROUP, 'seg_group', 'Group'])
                if asset_group_col:
                    group_asset_rows = df_assets[df_assets[asset_group_col] == group_id].copy()
                    start_col = self._get_column_name(df_assets, [Config.KEY_START, 'start_loc', 'Start'])
                    if not group_asset_rows.empty and start_col:
                        group_asset_rows = group_asset_rows.sort_values(by=start_col)

            # --- D. Determine Exact End Point (Last AP Position) ---
            last_ap_pos = unique_aps[-1][0]
            total_end_meters = last_ap_pos

            # --- E. Grid Generation ---
            site_dict = {
                'site_name': str(group_id),
                'ap_id_1': self._format_ap_name(unique_aps[0][1]) if unique_aps else '',
                'ap_id_2': self._format_ap_name(unique_aps[-1][1]) if unique_aps else '',

                'pipe_type': pipe_type,
                'pipe_specs_list': pipe_specs_list,  # NEW: List of all pipe specs
                'resolution': 'Standard',
                'ordered_segments': ordered_segments,  # NEW: Add ordered segment string
                'segments': []
            }

            num_steps = int(np.ceil(total_end_meters / self.INCREMENT_METERS))

            # Find Asset Columns (only if we have asset data)
            asset_id_col = None
            start_col_asset = None
            end_col_asset = None

            if not group_asset_rows.empty:
                asset_id_col = self._get_column_name(df_assets, [Config.KEY_ASSET_ID, 'pipe_asset_id', 'Asset ID'])
                start_col_asset = self._get_column_name(df_assets, [Config.KEY_START, 'start_loc', 'Start'])
                end_col_asset = self._get_column_name(df_assets, [Config.KEY_END, 'end_loc', 'End'])

            for i in range(num_steps):
                start_m = i * self.INCREMENT_METERS
                end_m = (i + 1) * self.INCREMENT_METERS

                # --- STORE BOTH UNITS ---
                start_ft = start_m * self.FEET_PER_METER
                end_ft = end_m * self.FEET_PER_METER

                # 1. Labels
                slice_labels = []
                for ap_pos, ap_name in unique_aps:
                    if start_m <= ap_pos < end_m:
                        slice_labels.append(ap_name)
                    elif abs(ap_pos - end_m) < 0.05 and abs(ap_pos - total_end_meters) < 0.05:
                        slice_labels.append(ap_name)

                # 2. Pipe Asset ID (only if required and data available)
                pipe_asset_id = None
                if not group_asset_rows.empty and start_col_asset and end_col_asset and asset_id_col:
                    midpoint_m = (start_m + end_m) / 2
                    match = group_asset_rows[
                        (group_asset_rows[start_col_asset] <= midpoint_m) &
                        (group_asset_rows[end_col_asset] > midpoint_m)
                        ]
                    if match.empty:
                        match = group_asset_rows[
                            (group_asset_rows[start_col_asset] <= start_m) &
                            (group_asset_rows[end_col_asset] >= end_m)
                            ]
                    if not match.empty:
                        pipe_asset_id = match.iloc[0][asset_id_col]

                ap_col_val = " / ".join(self._format_ap_name(name) for name in slice_labels) if slice_labels else None


                segment_slice = {
                    'start_m': start_m,  # Metric
                    'end_m': end_m,  # Metric
                    'start_ft': start_ft,  # Imperial
                    'end_ft': end_ft,  # Imperial
                    'dri_thickness': None,
                    'nom_thickness': None,
                    'pipe_asset_id': pipe_asset_id,
                    'access_point_label': ap_col_val
                }

                site_dict['segments'].append(segment_slice)

            # Ensure the very last label appears
            if site_dict['segments'] and unique_aps:
                last_ap_name = unique_aps[-1][1]
                found = False
                for s in site_dict['segments'][-2:]:
                    if s['access_point_label'] and last_ap_name in s['access_point_label']:
                        found = True
                if not found:
                    site_dict['segments'][-1]['access_point_label'] = last_ap_name

            grouped_data.append(site_dict)

        return grouped_data
