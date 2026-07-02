# Repository Guidelines

## Project Structure & Module Organization
This repository is a small Python toolkit for Oracle SQL and CTE metadata workflows. Source files live at the repository root:

- `oracle_query_tool.py`: Tkinter GUI for loading `.sql` files, executing Oracle queries, and exporting results to Excel.
- `export_cte_metadata.py`: Tkinter GUI for inspecting CTE column metadata through Oracle cursor descriptions.
- `config.example.json`: Template for local Oracle credentials and file paths. Copy it to `config.json` for local use.
- `.gitignore`: Excludes local secrets, virtual environments, generated Excel files, logs, and Oracle client folders.

There is currently no dedicated `tests/` directory or packaged asset folder. Keep new modules close to these scripts unless the project grows enough to justify subpackages.

## Build, Test, and Development Commands
Create and activate a virtual environment before development:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install oracledb pandas openpyxl chardet
```

Prepare local configuration:

```powershell
Copy-Item config.example.json config.json
```

Run the GUI tools:

```powershell
python oracle_query_tool.py
python export_cte_metadata.py
```

Use a syntax smoke test before committing:

```powershell
python -m py_compile oracle_query_tool.py export_cte_metadata.py
```

## Coding Style & Naming Conventions
Use Python 3.9+ and follow PEP 8 where practical. Use four-space indentation, `snake_case` for functions and variables, and `PascalCase` for Tkinter application classes. Keep UI text, configuration loading, Oracle access, and export logic separated into focused methods. Prefer `pathlib` or `os.path` consistently when adding file handling.

## Testing Guidelines
No automated test suite is present yet. For changes that touch parsing, configuration, or export behavior, add focused tests under `tests/` using `pytest` or `unittest`. Name test files `test_<feature>.py`. When database access is required, mock `oracledb.connect` rather than using real credentials.

## Commit & Pull Request Guidelines
The current Git history only contains `Initial commit`, so no strict convention is established. Use concise imperative commit subjects, for example `Add metadata export validation`. Pull requests should describe the user-facing change, list manual verification commands, and mention any Oracle version or client-mode assumptions.

## Security & Configuration Tips
Never commit `config.json`, `.env`, Oracle passwords, generated logs, or exported workbooks. Keep `config.example.json` sanitized and update it only with non-sensitive placeholder values.
