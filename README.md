# XELOS

This project contains VBA class modules extracted from an Excel workbook. A small
Flask application has been added in `server_app/` that demonstrates how the
workbook logic could be exposed on a server.

Run the application with:

```bash
python -m server_app
```

Upload `.xlsb` files through the web form to parse identification data along
with four analysis tables that mirror the original workbook sheets:

* `SYNTHESE`
* `ANALYSE_PRP`
* `ANALYSE_FINANCEMENT`
* `ANALYSE_LOYERS`

Once uploaded, visit `/dashboard` to view these records. The page accepts
optional `table`, `file` and `id2` parameters to filter the displayed rows. A
list of uploaded files is available at `/files`, where you can also delete
unwanted uploads. The dashboard groups rows by table name so you can quickly
review the different analyses.

The interface now uses simple HTML templates located in `server_app/templates`
to provide a cleaner layout for the upload page, dashboard and file manager.

## Extract VBA macros

The original workbook `VBA_python_BENCHMARK_3.10.06.25.xlsm` contains its VBA
modules inside `xl/vbaProject.bin`. Use the helper script below to extract this
binary to the `vba_src/` directory for inspection with external tools:

```bash
python scripts/extract_vba.py VBA_python_BENCHMARK_3.10.06.25.xlsm
```

The script simply copies `vbaProject.bin` from the workbook into `vba_src/`.

