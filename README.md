# XELOS

This project contains VBA class modules extracted from an Excel workbook. A small
Flask application has been added in `server_app/` that demonstrates how the
workbook logic could be exposed on a server.

Run the application with:

```bash
python -m server_app
```

Upload `.xlsb` files through the web form to parse basic identification data and
display it in a simple dashboard. The parser is incomplete but shows how the VBA
logic could be migrated to Python.

