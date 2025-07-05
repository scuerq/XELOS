# XELOS

This project contains VBA class modules extracted from an Excel workbook. A small
Flask application has been added in `server_app/` that demonstrates how the
workbook logic could be exposed on a server.

Run the application with:

```bash
python -m server_app
```

Upload `.xlsb` files through the web form to parse identification, synthesis and
locatif data. Once uploaded, visit `/dashboard` to view the records. The page
accepts optional `table` and `file` parameters to filter the displayed rows.

