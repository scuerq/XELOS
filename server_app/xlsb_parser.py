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
        """Parse identification information from the 'Identif' sheet."""
        tables = {}
        try:
            ranges = [
                ('C7:D21'),
                ('M19:O19'),
                ('M22:O22'),
                ('G85:M85'),
                ('C86:E86'),
            ]
            frames = []
            for r in ranges:
                df = self._read_range('Identif', r)
                if df.shape[1] < 2:
                    continue
                df = df.iloc[:, :2]
                frames.append(df)
            if frames:
                id_df = pd.concat(frames, ignore_index=True)
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
        loc = self.parse_locatif()
        if not loc.empty:
            data['LOCATIF'] = loc
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

    def parse_locatif(self):
        """Parse locatif data from the 'LoyersEtCharges' sheet."""
        ids = self.get_fille_ids()
        if not ids:
            return pd.DataFrame()

        try:
            df = pd.read_excel(
                self.path,
                sheet_name='LoyersEtCharges',
                engine='pyxlsb',
                header=14,
                usecols='C:T',
                nrows=41,
            )
        except Exception:
            return pd.DataFrame()

        df = df.dropna(how='all')
        data = []
        for _, row in df.iterrows():
            row_id = str(row.iloc[0])
            for fid in ids:
                if str(row_id) in str(fid):
                    for col in df.columns[1:]:
                        data.append({'Id2': fid, 'Key': col, 'Value': row[col]})
        return pd.DataFrame(data)
