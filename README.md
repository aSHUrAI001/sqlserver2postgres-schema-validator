# SQL Server to PostgreSQL Schema Validator

![Main UI Screenshot](PLACEHOLDER_FOR_UI_IMAGE)
*Replace this line with your application's main UI screenshot.*

## Overview

**SQL Server to PostgreSQL Schema Validator** is a modern desktop tool for comparing and validating database schemas between SQL Server and PostgreSQL. It features a user-friendly GUI, robust schema comparison logic, and generates detailed Excel reports for analysis.

---

## Features

- **Intuitive Tkinter GUI**: Configure connections, run validations, and view recent reports.
- **Windows & SQL Authentication**: Supports both authentication modes for SQL Server.
- **Configurable Database Lists**: Validate multiple databases in one go.
- **Customizable Mappings**: Easily adjust schema mappings for your environment.
- **Detailed Excel Reports**: Automatically generated and accessible from the UI.
- **Recent Reports Management**: View, open, or delete recent reports directly from the app.

---

## Getting Started

### 1. Clone the Repository

```sh
git clone https://github.com/<your-username>/sqlserver2postgres-schema-validator.git
cd sqlserver2postgres-schema-validator/v2
```

### 2. Install Requirements

Ensure you have Python 3.8+ installed.  
Install dependencies using:

```sh
pip install -r requirements.txt
```

### 3. Configure Database Connections

Edit `config.py` to set your SQL Server and PostgreSQL connection details.  
- For **Windows Authentication**, set `windows_auth = True` in `SQL_SERVER_CONFIG`.
- For **SQL Authentication**, set `windows_auth = False` and provide `username` and `password`.

Example:
```python
SQL_SERVER_CONFIG = {
    'server': 'YOUR_SQL_SERVER',
    'database': 'YOUR_DB',
    'username': 'YOUR_USER',
    'password': 'YOUR_PASS',
    'driver': 'ODBC Driver 17 for SQL Server',
    'windows_auth': True  # or False
}
POSTGRES_CONFIG = {
    'host': 'localhost',
    'database': 'YOUR_PG_DB',
    'user': 'YOUR_PG_USER',
    'password': 'YOUR_PG_PASS',
    'port': '5432'
}
DB_LIST = ['db1', 'db2']
```

You can also use the **Config UI** in the application to update these settings interactively.

### 4. Customize Mappings

Edit `mappings.py` to adjust column, type, or table mappings as needed for your schema comparison.

---

## Running the Application

```sh
python SchemaValidatorUI.py
```

- The main window will open, allowing you to configure connections, select databases, and start validation.
- Use the **Validate & Generate Report** button to run schema comparison.
- Access recent reports from the right panel; open or delete them as needed.

---

## UI Guide

- **Config Panel (Left)**: Set SQL Server and PostgreSQL connection details, choose authentication mode, and specify databases.
- **Validation Panel (Right, Top)**: Start validation and view status messages.
- **Recent Reports (Right, Bottom)**: Manage generated Excel reports.
- **About & Help**: Click the About button for version info and support details.

---

## Modifying the Tool

- **Configuration**: Use `config.py` or the Config UI to change connection settings and database lists.
- **Mappings**: Update `mappings.py` for custom schema mapping logic.
- **Schema Comparison Logic**: See `SchemaValidatior.py` for backend features, including:
  - Column, trigger, and constraint matching
  - Authentication handling
  - Report generation

---

## Reports

- Reports are saved as `.xlsx` files in the `SchemaValidationReports` folder.
- Each report details schema differences, missing columns, mismatches, and more.
- Use the UI to view or delete recent reports, or open the folder directly.

---

## Packaging as an Executable

To create a standalone `.exe` (Windows):

```sh
pyinstaller --onefile --windowed SchemaValidatorUI.py
```

See the README for troubleshooting tips if you encounter packaging issues.

---

## Support

For issues or feature requests, please open a GitHub issue or contact the development team.

---

## License

This project is licensed under the MIT License.

---

**Replace the UI image placeholder above with your application's screenshot for best presentation.**
