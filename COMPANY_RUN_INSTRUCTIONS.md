# Company Run Instructions

Use the two step `.bat` files on Windows. Keep this file, the two `.bat` files, `requirements.txt`, and `2-6-2026_Final_Working.py` in the main project folder.

Expected folder structure:

```text
Project Main/
  2-6-2026_Final_Working.py
  Step1_Generate_Missing_SKU.bat
  Step2_Run_Full_Pipeline.bat
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

Run this first:

```text
Step1_Generate_Missing_SKU.bat
```

If missing MSKUs are found, Step 1 creates or checks:

```text
User Input/Missing SKU File/Missing SKU.xlsx
```

Fill the required columns in that file and save it.

Then run:

```text
Step2_Run_Full_Pipeline.bat
```

Step 2 updates the master file and generates the output reports. The script will not overwrite a partially filled missing-SKU file.
