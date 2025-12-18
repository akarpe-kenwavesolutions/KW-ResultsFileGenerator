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
        missing_files = []
        if not self.seg_df_path: missing_files.append("Segment DF")
        if not self.asset_path and Config.REQUIRE_ASSET_IDS: missing_files.append("Asset IDs DF")
        if missing_files: raise FileNotFoundError(f"Missing Input Files:\n- " + "\n- ".join(missing_files))

        print(f"Loading Metadata: {os.path.basename(self.seg_df_path)}")
        if self.asset_path: print(f"Loading Assets:   {os.path.basename(self.asset_path)}")
        if self.seg_groups_path: print(f"Loading SegGroups: {os.path.basename(self.seg_groups_path)}")

        df_assets = pd.read_csv(self.asset_path) if self.asset_path else pd.DataFrame()
        df_seg_meta = pd.read_csv(self.seg_df_path)
        seg_groups_dict = self.load_seg_groups()

        if not seg_groups_dict: raise FileNotFoundError("segGroups file is required.")

        grouped_data = []
        print(f"Processing {len(seg_groups_dict)} sites...")

        for group_id, segment_codes in seg_groups_dict.items():
            if not segment_codes or (len(segment_codes) == 1 and segment_codes[0] == ''): continue

            # --- A. Filter Metadata ---
            group_meta_rows = df_seg_meta[df_seg_meta['Unnamed: 0'].isin(segment_codes)].copy()
            if group_meta_rows.empty: continue

            # Pipe Type Extraction
            pipe_type = "Unknown"
            best_dia = ""
            best_mat = ""
            for _, row in group_meta_rows.iterrows():
                d = self._get_val_robust(row, [Config.KEY_META_DIA, 'diameter', 'Diameter'])
                m = self._get_val_robust(row, [Config.KEY_META_MAT, 'material', 'Material'])
                if d and not best_dia: best_dia = str(d)
                if m and not best_mat: best_mat = str(m)
                if best_dia and best_mat: break
            if best_dia or best_mat: pipe_type = f"{best_dia} {best_mat}".strip()

            # --- B. Load Access Points (Source of Truth for Distances) ---
            ap_positions_dict = self.load_access_points_for_group(group_id)
            unique_aps = sorted([(pos, name) for name, pos in ap_positions_dict.items()], key=lambda x: x[0])

            if not unique_aps:
                print(f"  Skipping {group_id} (No Access Points found)")
                continue

            # --- C. Load Assets ---
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
                'ap_id_1': unique_aps[0][1] if unique_aps else '',
                'ap_id_2': unique_aps[-1][1] if unique_aps else '',
                'pipe_type': pipe_type,
                'resolution': 'Standard',
                'segments': []
            }

            num_steps = int(np.ceil(total_end_meters / self.INCREMENT_METERS))

            # Find Asset Columns
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

                # 2. Pipe Asset ID
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

                ap_col_val = " / ".join(slice_labels) if slice_labels else None

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
