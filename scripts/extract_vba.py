import zipfile
import os
import sys


def extract_vba(xlsm_path, out_dir='vba_src'):
    """Extract the vbaProject.bin from the XLSM file."""
    if not os.path.isfile(xlsm_path):
        raise FileNotFoundError(xlsm_path)

    with zipfile.ZipFile(xlsm_path) as zf:
        members = zf.namelist()
        try:
            member = next(m for m in members if m.endswith('vbaProject.bin'))
        except StopIteration:
            raise KeyError('vbaProject.bin not found in the XLSM')

        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, 'vbaProject.bin')
        with zf.open(member) as src, open(out_path, 'wb') as dst:
            dst.write(src.read())
        return out_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python extract_vba.py <workbook.xlsm> [out_dir]')
        sys.exit(1)
    out = extract_vba(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else 'vba_src')
    print('Saved:', out)
