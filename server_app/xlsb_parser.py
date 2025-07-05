import pandas as pd
from collections import defaultdict

class XLSBParser:
    """Parse XLSB files according to mappings defined in the VBA modules."""

    def __init__(self, path):
        self.path = path

    def _read_range(self, sheet_name, cell_range):
        df = pd.read_excel(self.path, sheet_name=sheet_name, engine='pyxlsb', header=None)
        # Convert Excel-style range (e.g., "F18:F28") to zero-based index slicing
        # This is simplified: we read full sheet and slice with pandas' iloc.
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        return df.iloc[min_row-1:max_row, min_col-1:max_col]

    def parse_id(self):
        tables = {}
        try:
            id_df = self._read_range('Identif', 'C7:D21')
            id_df.columns = ['Key', 'Value']
            tables['ID'] = id_df
        except Exception:
            pass
        return tables

    def parse_all(self):
        data = defaultdict(lambda: pd.DataFrame(columns=['Key', 'Value']))
        data.update(self.parse_id())
        synth = self.parse_synth()
        if not synth.empty:
            data['SYNTH'] = synth
        return data
    def get_fille_ids(self):
        try:
            df = pd.read_excel(self.path, sheet_name='Identif', engine='pyxlsb', header=None)
            row = df.iloc[56]  # row 57 in Excel
            ids = [c for c in row[5:10] if pd.notna(c)]
            return ids
        except Exception:
            return []
    def parse_synth(self):
        data = []
        ids = self.get_fille_ids()
        for idx, fid in enumerate(ids):
            try:
                df = self._read_range('Fiche_Synthse', f'F18:F28')
                df.columns = ['Key']
                values = self._read_range('Fiche_Synthse', f'F18:F28').iloc[:, idx]
                for k, v in zip(df['Key'], values):
                    data.append({'Id2': fid, 'Key': k, 'Value': v})
            except Exception:
                continue
        return pd.DataFrame(data)
