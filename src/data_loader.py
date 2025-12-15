import pandas as pd
import os
import re
import numpy as np
from config import Config


class DataLoader:
    def __init__(self):
        self.seg_df_path = Config.FILE_SEG_DF
        self.asset_path = Config.FILE_PIPE_ASSETS
        self.spec_path = Config.FILE_PIPE_SPECS
        self.INCREMENT_METERS = 2.0  # Fixed increment size

    def _extract_group_from_seg_code(self, row_val):
        """Extracts text inside parentheses, e.g., 'S01 (GroupA)' -> 'GroupA'."""
        if pd.isna(row_val): return None
        match = re.search(r'\((.*?)\)', str(row_val))
        return match.group(1) if match else None

    def load_data(self):
        # 1. Validation
        missing_files = []
        if not self.seg_df_path: missing_files.append("Segment DF (filename must contain 'seg_df')")
        if not self.asset_path: missing_files.append("Asset IDs DF (filename must contain 'pipe_asset_ids_df')")
        if missing_files:
            raise FileNotFoundError(f"Missing Input Files:\n- " + "\n- ".join(missing_files))

        print(f"Loading Metadata: {os.path.basename(self.seg_df_path)}")
        print(f"Loading Assets:   {os.path.basename(self.asset_path)}")

        df_assets = pd.read_csv(self.asset_path)
        df_seg_meta = pd.read_csv(self.seg_df_path)

        # 2. Metadata Lookup & Processing Order
        # Extract groups to determine the source order
        df_seg_meta['extracted_group'] = df_seg_meta['Unnamed: 0'].apply(self._extract_group_from_seg_code)

        # Get unique groups strictly in the order they appear in seg_df
        ordered_groups_source = df_seg_meta['extracted_group'].dropna().unique()

        # Get set of available assets
        available_asset_groups = set(df_assets[Config.KEY_GROUP].unique())

        # Final List: Intersection of Ordered Source & Available Assets
        final_processing_list = [g for g in ordered_groups_source if g in available_asset_groups]

        # Build Metadata Dictionary
        meta_lookup = {}
        for _, row in df_seg_meta.dropna(subset=['extracted_group']).iterrows():
            group = row['extracted_group']
            pipe_type_str = f"{row.get(Config.KEY_META_DIA, '')} {row.get(Config.KEY_META_MAT, '')}"

            meta_lookup[group] = {
                'ap_id_1': row.get(Config.KEY_META_AP1),
                'ap_id_2': row.get(Config.KEY_META_AP2),
                'ap_1_loc': row.get('ap_1_loc', 0),  # Default to 0 if missing
                'ap_2_loc': row.get('ap_2_loc', 0),
                'pipe_type': pipe_type_str,
                'resolution': 'Standard'
            }

        # 3. Process Groups
        grouped_data = []
        print(f"Processing {len(final_processing_list)} sites (Ordered by Source)...")

        for group_id in final_processing_list:
            meta = meta_lookup.get(group_id, {})
            site_dict = {
                'site_name': str(group_id),
                'ap_id_1': meta.get('ap_id_1', ''),
                'ap_id_2': meta.get('ap_id_2', ''),
                'pipe_type': meta.get('pipe_type', ''),
                'resolution': meta.get('resolution', ''),
                'segments': []
            }

            # Get Access Point Locations (Float conversion with safety default -1)
            try:
                ap1_loc = float(meta.get('ap_1_loc')) if pd.notna(meta.get('ap_1_loc')) else -1.0
            except:
                ap1_loc = -1.0

            try:
                ap2_loc = float(meta.get('ap_2_loc')) if pd.notna(meta.get('ap_2_loc')) else -1.0
            except:
                ap2_loc = -1.0

            # Get Assets for this Group
            group_rows = df_assets[df_assets[Config.KEY_GROUP] == group_id].copy()
            group_rows = group_rows.sort_values(by=Config.KEY_START)

            if group_rows.empty:
                grouped_data.append(site_dict)
                continue

            # Determine Total Length
            total_end = group_rows[Config.KEY_END].max()

            # Generate Continuous Grid
            current_pos = 0.0

            while current_pos < total_end:
                next_pos = current_pos + self.INCREMENT_METERS

                # Handle Last Step Overhang
                if next_pos >= (total_end - 0.001):
                    next_pos = total_end

                # 1. Lookup Asset ID (Midpoint Method)
                midpoint = (current_pos + next_pos) / 2
                matching_asset = group_rows[
                    (group_rows[Config.KEY_START] <= midpoint) &
                    (group_rows[Config.KEY_END] > midpoint)
                    ]

                # Fallback if midpoint hits a gap/boundary
                if matching_asset.empty:
                    matching_asset = group_rows[
                        (group_rows[Config.KEY_START] <= current_pos) &
                        (group_rows[Config.KEY_END] >= next_pos)
                        ]

                asset_id = matching_asset.iloc[0][Config.KEY_ASSET_ID] if not matching_asset.empty else None

                # 2. Check for Access Point Match (SNAP LOGIC)
                ap_label = None

                # --- AP 1 (Start) ---
                if ap1_loc != -1:
                    # Normal Match: Inside this slice
                    if current_pos <= ap1_loc < (next_pos - 0.001):
                        ap_label = str(meta.get('ap_id_1'))
                    # Snap Match: If this is the FIRST row (pos=0) and AP1 is <= 0
                    elif current_pos == 0.0 and ap1_loc <= 0.001:
                        ap_label = str(meta.get('ap_id_1'))

                # --- AP 2 (End) ---
                if ap2_loc != -1:
                    # Normal Match: Inside this slice
                    if current_pos <= ap2_loc < (next_pos - 0.001):
                        ap_label = str(meta.get('ap_id_2'))
                    # Snap Match: If this is the LAST row and AP2 is >= current pos
                    elif next_pos >= (total_end - 0.001) and ap2_loc >= current_pos:
                        ap_label = str(meta.get('ap_id_2'))

                segment_slice = {
                    'start_ft': current_pos,
                    'end_ft': next_pos,
                    'dri_thickness': None,
                    'nom_thickness': None,
                    'pipe_asset_id': asset_id,
                    'access_point_label': ap_label
                }
                site_dict['segments'].append(segment_slice)

                current_pos = next_pos

            grouped_data.append(site_dict)

        return grouped_data
