# Company Run Instructions

Use `run.bat` on Windows. Keep this file, `run.bat`, `requirements.txt`, and `2-6-2026_Final_Working.py` in the main project folder.

Expected folder structure:

```text
Project Main/
  2-6-2026_Final_Working.py
  run.bat
  requirements.txt
  System Files/
    Master Sheet Updated.xlsx
  User Input/
    ITD Projection/
      Input File.xlsx
    Missing SKU File/
    Print Review File/
      Print Review V10 May, 26 Update.xlsm
    PPB RM/
      Main.xlsx
    PSPD RM/
      Bhadrachalam.xlsx
      Guntur.xlsx
      PSPD Sheet Stocks.xlsx
  User Output/
```

Logs are written to `logs/`.

Backups of existing overwritten Excel files are written to `backups/`.

If missing MSKUs are found, the first run creates:

```text
User Input/Missing SKU File/Missing SKU.xlsx
```

Fill the required columns in that file, then run `run.bat` again. The script will not overwrite a partially filled missing-SKU file.
