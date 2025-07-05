# XELOS

This project contains VBA class modules extracted from an Excel workbook. A small
Flask application has been added in `server_app/` that demonstrates how the
workbook logic could be exposed on a server.

Run the application with:

```bash
python -m server_app
```

Upload `.xlsb` files through the web form to parse identification, synthesis,
locatif, PRP and financing data. Once uploaded, visit `/dashboard` to view the
records. The page accepts optional `table`, `file` and `id2` parameters to
filter the displayed rows. A list of uploaded files is available at `/files`.

## Extract VBA macros

The original workbook `VBA_python_BENCHMARK_3.10.06.25.xlsm` contains its VBA
modules inside `xl/vbaProject.bin`. Use the helper script below to extract this
binary to the `vba_src/` directory for inspection with external tools:

```bash
python scripts/extract_vba.py VBA_python_BENCHMARK_3.10.06.25.xlsm
```

The script simply copies `vbaProject.bin` from the workbook into `vba_src/`.

