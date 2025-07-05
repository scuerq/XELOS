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

    def _offset_range(self, cell_range, col_offset=0, row_offset=0):
        """Return a new range string offset from the original range."""
        from openpyxl.utils import range_boundaries, get_column_letter
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        min_col += col_offset
        max_col += col_offset
        min_row += row_offset
        max_row += row_offset
        from openpyxl.utils import get_column_letter
        return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"

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
            data['SYNTHESE'] = synth
        loc = self.parse_locatif()
        if not loc.empty:
            data['ANALYSE_LOYERS'] = loc
        prp = self.parse_prp()
        if not prp.empty:
            data['ANALYSE_PRP'] = prp
        finan = self.parse_financement()
        if not finan.empty:
            data['ANALYSE_FINANCEMENT'] = finan
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

    def parse_prp(self):
        """Parse PRP values from the 'PRP IKOS' sheet."""
        ids = self.get_fille_ids()
        if not ids:
            return pd.DataFrame()

        base_cells = ['D52', 'D80', 'D99', 'D108', 'D119', 'D124', 'D136', 'D138']
        rows = []
        for idx, fid in enumerate(ids):
            for cell in base_cells:
                try:
                    key = self._read_range('PRP IKOS', cell).iloc[0, 0]
                    val_cell = self._offset_range(cell, col_offset=10 + idx)
                    val = self._read_range('PRP IKOS', val_cell).iloc[0, 0]
                    rows.append({'Id2': fid, 'Key': key, 'Value': val})
                except Exception:
                    continue
        return pd.DataFrame(rows)

    def parse_financement(self):
        """Parse financing data from the 'Financement' sheet."""
        ids = self.get_fille_ids()
        if not ids:
            return pd.DataFrame()

        sections = [
            ('E13:E33', 4, 'Subvention'),
            ('F35:F45', 3, 'Fonds propres'),
            ('F46:F71', 3, 'Prets'),
        ]

        rows = []
        for idx, fid in enumerate(ids):
            for base_range, offset, label in sections:
                try:
                    keys = self._read_range('Financement', base_range).iloc[:, 0]
                    val_range = self._offset_range(base_range, col_offset=offset + idx)
                    values = self._read_range('Financement', val_range).iloc[:, 0]
                    for k, v in zip(keys, values):
                        rows.append({'Id2': fid, 'Key': f'{label}:{k}', 'Value': v})
                except Exception:
                    continue
        return pd.DataFrame(rows)
