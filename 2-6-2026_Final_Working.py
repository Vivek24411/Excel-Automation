import importlib
import logging
import shutil
import sys
from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
BACKUP_DIR = BASE_DIR / "backups"
RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = LOG_DIR / f"run_{RUN_ID}.log"
_BACKED_UP_THIS_RUN = set()

LOG_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


class Tee:
    def __init__(self, stream, log_path):
        self.stream = stream
        self.log_path = log_path

    def write(self, message):
        self.stream.write(message)
        with open(self.log_path, "a", encoding="utf-8") as handle:
            handle.write(message)

    def flush(self):
        self.stream.flush()


sys.stdout = Tee(sys.__stdout__, LOG_FILE)
sys.stderr = Tee(sys.__stderr__, LOG_FILE)


def log_uncaught_exception(exc_type, exc_value, exc_traceback):
    logging.critical(
        "Unhandled exception",
        exc_info=(exc_type, exc_value, exc_traceback),
    )
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = log_uncaught_exception


def require_dependencies():
    missing = []
    for package in ["pandas", "openpyxl", "xlsxwriter"]:
        try:
            importlib.import_module(package)
        except ModuleNotFoundError:
            missing.append(package)

    if missing:
        raise RuntimeError(
            "Missing Python packages: "
            + ", ".join(missing)
            + ". Run install first: pip install -r requirements.txt"
        )


require_dependencies()

import pandas as pd
import xlsxwriter
import warnings
warnings.filterwarnings("ignore")
print("Libraries ready")
print("Log file:", LOG_FILE)
logging.info("Process started from %s", BASE_DIR)


def rel_path(*parts):
    return str(BASE_DIR.joinpath(*parts))


def ensure_parent(path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)


def require_file(path, label):
    if not Path(path).exists():
        raise FileNotFoundError(f"{label} not found: {path}")


def backup_existing(path, reason="overwrite"):
    src = Path(path)
    if not src.exists():
        return
    resolved = src.resolve()
    if resolved in _BACKED_UP_THIS_RUN:
        return
    backup_name = f"{src.stem}.backup_{RUN_ID}{src.suffix}"
    backup_path = BACKUP_DIR / backup_name
    shutil.copy2(src, backup_path)
    _BACKED_UP_THIS_RUN.add(resolved)
    print(f"Backup created before {reason}: {backup_path}")
    logging.info("Backup created for %s at %s", src, backup_path)


def prepare_output(path, backup=True):
    ensure_parent(path)
    if backup:
        backup_existing(path)


def open_workbook(path, backup=True):
    prepare_output(path, backup=backup)
    return xlsxwriter.Workbook(path)


def write_excel(df, path, **kwargs):
    prepare_output(path, backup=True)
    df.to_excel(path, **kwargs)


def missing_file_ready(path, required_mskus, required_columns):
    path_obj = Path(path)
    if not path_obj.exists():
        return False
    try:
        df_check = pd.read_excel(path, dtype=str, engine="openpyxl").fillna("")
    except Exception as exc:
        print(f"Existing missing SKU file could not be read: {exc}")
        return False

    if "MSKU" not in df_check.columns:
        return False

    df_check["MSKU"] = df_check["MSKU"].str.strip()
    present_mskus = set(df_check["MSKU"])
    needed_mskus = set(str(msku).strip() for msku in required_mskus)
    if not needed_mskus.issubset(present_mskus):
        return False

    rows = df_check[df_check["MSKU"].isin(needed_mskus)].copy()
    missing_columns = [col for col in required_columns if col not in rows.columns]
    if missing_columns:
        print("Missing SKU file is missing required columns:", missing_columns)
        return False

    blank_columns = [
        col for col in required_columns
        if rows[col].astype(str).str.strip().eq("").any()
    ]
    if blank_columns:
        print("Missing SKU file needs these columns filled:", blank_columns)
        return False

    return True

# ---- CONFIGURE PATHS HERE --------------------------------------------------
SYSTEM_FILE  = rel_path("System Files", "Master Sheet Updated.xlsx")   # master reference file
INPUT_FILE   = rel_path("User Input", "ITD Projection", "Input File.xlsx")    # file to check MSKUs from
MISSING_FILE = rel_path("User Input", "Missing SKU File", "Missing SKU.xlsx")  # will be overwritten in place
MSKU_COL     = "MSKU"                # column name in both files

require_file(SYSTEM_FILE, "Master/system file")
require_file(INPUT_FILE, "Input projection file")

# ---- LOAD FILES & FIND MISSING MSKUs --------------------------------------
df_system = pd.read_excel(SYSTEM_FILE, dtype=str, engine="openpyxl").fillna("")
df_input  = pd.read_excel(INPUT_FILE,  dtype=str, engine="openpyxl").fillna("")

system_mskus  = set(df_system[MSKU_COL].str.strip())
missing_mskus = (df_input
    .loc[~df_input[MSKU_COL].str.strip().isin(system_mskus), MSKU_COL]
    .str.strip().unique().tolist())

print("System MSKUs  :", len(system_mskus))
print("Input rows    :", len(df_input))
print("Missing MSKUs :", len(missing_mskus))
print(missing_mskus[:10])

# ---- BUILD MISSING FILE WITH DROPDOWNS ------------------------------------
DROPDOWN_COLS  = ["Pack Style", "Board Type", "Route"]
OUTPUT_COLUMNS = [
    "Product", "V9-WMS", "V9-FG Code", "VX-FG Code", "VX-WMS",
    "MSKU", "Material", "Pack Style", "Board GSM", "Board Size",
    "Board Length/Cyl Circumference", "Board Type", "Cat", "Route", "Pack Size", "No.Of UPS"
]
REQUIRED_FILLED_MISSING_COLUMNS = [
    "Product", "MSKU", "Pack Style", "Board GSM", "Board Size",
    "Board Length/Cyl Circumference", "Board Type", "Cat", "Route",
    "Pack Size", "No.Of UPS"
]
missing_file_is_ready = bool(missing_mskus) and missing_file_ready(
    MISSING_FILE,
    missing_mskus,
    REQUIRED_FILLED_MISSING_COLUMNS,
)

if not missing_mskus:
    print("No missing MSKUs - nothing to write.")
elif missing_file_is_ready:
    print("Filled missing SKU file found. Continuing with master update.")
elif Path(MISSING_FILE).exists():
    print("Existing missing SKU file is not complete, so it was not overwritten.")
    print("Please fill the required columns in this file and run again:", MISSING_FILE)
    logging.info("Stopped because missing SKU file is incomplete: %s", MISSING_FILE)
    sys.exit(1)
else:
    # Collect dropdown options from system file
    dropdown_options = {}
    for col in DROPDOWN_COLS:
        if col in df_system.columns:
            vals = sorted(df_system[col].str.strip().replace("", None).dropna().unique().tolist())
            vals.append("Others")
            dropdown_options[col] = vals
            print(col + ": " + str(vals))
        else:
            print("WARNING: column not found in system file: " + col)

    # Full column list (main cols + Others free-text cols at end)
    others_cols = [c + " (Others)" for c in DROPDOWN_COLS if c in dropdown_options]
    all_cols    = OUTPUT_COLUMNS + others_cols
    others_set  = set(others_cols)
    col_idx     = {name: i for i, name in enumerate(all_cols)}

    wb = open_workbook(MISSING_FILE)
    ws = wb.add_worksheet("Missing MSKUs")

    # Formats
    hdr  = wb.add_format({"bold":True,"font_name":"Arial","font_size":11,"bg_color":"#2E4057","font_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
    ohdr = wb.add_format({"bold":True,"font_name":"Arial","font_size":11,"bg_color":"#E6A817","font_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
    dat  = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
    alt  = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#F0F4F8","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
    oth  = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#FFF3CD","valign":"vcenter","border":1,"border_color":"#CCCCCC"})

    # Header row
    ws.set_row(0, 24)
    for i, name in enumerate(all_cols):
        ws.write(0, i, name, ohdr if name in others_set else hdr)

    # Data rows
    for r, msku in enumerate(missing_mskus, start=1):
        ws.set_row(r, 18)
        rfmt = alt if r % 2 == 0 else dat
        for i, col in enumerate(all_cols):
            if col in others_set:
                ws.write(r, i, "", oth)
            elif col == "MSKU":
                ws.write(r, i, msku, rfmt)
            elif col == "Product":
                ws.write(r, i, "HLP", rfmt)
            else:
                ws.write(r, i, "", rfmt)

    # Dropdowns
    n = len(missing_mskus)
    for col_name, options in dropdown_options.items():
        c = col_idx[col_name]
        ws.data_validation(1, c, n, c, {
            "validate": "list",
            "source": options,
            "input_title": col_name,
            "input_message": "Pick a value. Choose Others for custom entry.",
            "error_title": "Invalid Input",
            "error_message": "Select from list or choose Others."
        })
        print("Dropdown added: " + col_name)

    # Column widths, freeze, autofilter
    for i, col in enumerate(all_cols):
        ws.set_column(i, i, max(len(col) + 4, 14))
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, n, len(all_cols) - 1)

    wb.close()
    print("Done -> " + MISSING_FILE)
    print("Rows written : " + str(n))
    print("Others cols  : " + str(others_cols))
    print("Product col pre-filled -> HLP for all rows.")
    print("Please fill the missing SKU file and run this script again.")
    logging.info("Stopped after creating missing SKU template: %s", MISSING_FILE)
    sys.exit(0)

# ---- EXTEND: Append filled missing SKUs to Master Sheet ---------------
# Reads the filled missing_msku.xlsx, resolves Others values,
# appends rows to master sheet in order, auto-fills S.No., overwrites master.

FILLED_MISSING_FILE = MISSING_FILE   # the filled missing file uploaded by user
MASTER_FILE         = SYSTEM_FILE             # same as system file - will be overwritten

# Columns that have Others logic
OTHERS_COLS = ["Pack Style", "Board Type", "Route"]

if not missing_mskus:
    df_filled = pd.DataFrame(columns=OUTPUT_COLUMNS)
else:
    require_file(FILLED_MISSING_FILE, "Filled missing SKU file")
    df_filled = pd.read_excel(FILLED_MISSING_FILE, dtype=str, engine="openpyxl").fillna("")
    current_missing_set = set(str(msku).strip() for msku in missing_mskus)
    df_filled["MSKU"] = df_filled["MSKU"].str.strip()
    df_filled = (
        df_filled[df_filled["MSKU"].isin(current_missing_set)]
        .drop_duplicates(subset=["MSKU"], keep="first")
        .copy()
    )
print("Filled missing file rows :", len(df_filled))
print("Columns :", df_filled.columns.tolist())

# Resolve Others: for each dropdown col, if value == "Others" replace with (Others) col value
for col in OTHERS_COLS:
    others_col = col + " (Others)"
    if col in df_filled.columns and others_col in df_filled.columns:
        mask = df_filled[col].str.strip().str.lower() == "others"
        df_filled.loc[mask, col] = df_filled.loc[mask, others_col].str.strip()
        resolved = mask.sum()
        print("Resolved Others in", col, ":", resolved, "rows")

# Keep only the 15 master columns (drop the (Others) helper cols)
MASTER_COLUMNS = [
    "Product", "V9-WMS", "V9-FG Code", "VX-FG Code", "VX-WMS",
    "MSKU", "Material", "Pack Style", "Board GSM", "Board Size",
    "Board Length/Cyl Circumference", "Board Type", "Cat", "Route", "Pack Size", "No.Of UPS"
]
df_to_add = df_filled[[c for c in MASTER_COLUMNS if c in df_filled.columns]].copy()

# Load master sheet
df_master = pd.read_excel(MASTER_FILE, dtype=str, engine="openpyxl").fillna("")
print("Master rows before :", len(df_master))

# Find current max S.No. and continue numbering
sno_col = "S.No."
if sno_col in df_master.columns:
    existing_snos = pd.to_numeric(df_master[sno_col], errors="coerce").dropna()
    start_sno = int(existing_snos.max()) + 1 if len(existing_snos) > 0 else 1
else:
    start_sno = 1
print("S.No. will start from :", start_sno)

# Add S.No. to new rows
df_to_add.insert(0, sno_col, range(start_sno, start_sno + len(df_to_add)))

# Ensure master has S.No. column; align columns
if sno_col not in df_master.columns:
    df_master.insert(0, sno_col, range(1, len(df_master) + 1))

if df_to_add.empty:
    df_updated = df_master.copy()
    print("No missing SKU rows to append. Master file left unchanged.")
else:
    # Append in exact order, reset index
    df_updated = pd.concat([df_master, df_to_add], ignore_index=True)

    # Overwrite master file
    write_excel(df_updated, MASTER_FILE, index=False, engine="openpyxl")
    print("Master rows after  :", len(df_updated))
    print("Done - master file updated ->", MASTER_FILE)
    df_updated.tail(len(df_to_add) + 2)

# ---- EXTEND: Create Program File from Updated Master Sheet ------------
# Creates a COPY of the updated master sheet, adds Program column.
# Program values matched from Input file Prog. column using MSKU as key.
# Unmatched MSKUs get Program = 0.
# CBO rows get Program of the HLP row immediately above them.

PROGRAM_FILE = rel_path("User Output", "Program File", "Program File.xlsx")   # output copy - original master untouched

# Load updated master (the overwritten master file)
df_program = pd.read_excel(MASTER_FILE, dtype=str, engine="openpyxl").fillna("")
print("Loaded updated master rows :", len(df_program))

# Load input file to get Prog. column
df_input_prog = pd.read_excel(INPUT_FILE, dtype=str, engine="openpyxl").fillna("")
print("Input file rows            :", len(df_input_prog))

# Build MSKU -> Prog. mapping from input file
# If MSKU appears more than once, sum all its Prog. values
prog_map = (
    df_input_prog
    .assign(Prog_num=pd.to_numeric(df_input_prog["Prog."].str.strip(), errors="coerce").fillna(0))
    .groupby(MSKU_COL)["Prog_num"]
    .sum()
    .astype(str)
    .to_dict()
)

# Step 1: Fill Program from input file using MSKU as key; default to 0
df_program["Program"] = (
    df_program[MSKU_COL]
    .str.strip()
    .map(prog_map)
    .fillna("0")
)
print("Program filled. Unmatched (set to 0) :",
      (df_program["Program"] == "0").sum())

# Step 2: CBO rule - for every CBO row, use Program of the HLP row above it
for i in range(1, len(df_program)):
    if df_program.loc[i, "Product"].strip().upper() == "CBO":
        # Walk upward to find the nearest HLP row
        for j in range(i - 1, -1, -1):
            if df_program.loc[j, "Product"].strip().upper() == "HLP":
                df_program.loc[i, "Program"] = df_program.loc[j, "Program"]
                break

cbo_count = (df_program["Product"].str.strip().str.upper() == "CBO").sum()
print("CBO rows updated with HLP Program   :", cbo_count)

# Save as Program File (copy - master file untouched)
write_excel(df_program, PROGRAM_FILE, index=False, engine="openpyxl")
print("Program File created ->", PROGRAM_FILE)
df_program[["S.No.", "MSKU", "Product", "Program"]].tail(10)

# ---- EXTEND: Add Finished Goods, WIP, Total Balance to Program File ----
PRINT_REVIEW_FILE = rel_path("User Input", "Print Review File", "Print Review V10 May, 26 Update.xlsm")
require_file(PRINT_REVIEW_FILE, "Print review file")
SHEETS_TO_READ    = ["10s HLC", "20s HLC & CBO"]

COL_SUMNSUM = "Sum/N.Sum"
COL_FGNUM   = "FG Number"
COL_MPF     = "MPF"
COL_TVT     = "TVT"
COL_UPF     = "UPF"
COL_NPF     = "NPF"
COL_OPSOH   = "Op SOH"

def clean_val(v):
    return str(v).replace(chr(10), " ").replace(chr(13), " ").strip()

df_prog = pd.read_excel(PROGRAM_FILE, dtype=str, engine="openpyxl").fillna("")
print("Program File rows :", len(df_prog))

fg_map = {}

for sheet in SHEETS_TO_READ:
    print("Reading sheet:", sheet)

    df_raw = pd.read_excel(
        PRINT_REVIEW_FILE, sheet_name=sheet,
        header=None, dtype=str, engine="openpyxl"
    ).fillna("")

    # Find the row containing MPF — that is the deepest sub-heading row
    mpf_row_idx = None
    for idx, row in df_raw.iterrows():
        if any(clean_val(v) == "MPF" for v in row.values):
            mpf_row_idx = idx
            break

    if mpf_row_idx is None:
        print("  WARNING: Could not find MPF row - skipping", sheet)
        continue
    print("  MPF row found at index:", mpf_row_idx)

    # Build final header per column:
    # Take the LAST non-empty value scanning top-down across all header rows
    # This means sub-heading (row 3) overrides heading (row 1/2) where present,
    # and heading value fills in where sub-heading cell is empty (e.g. Op SOH at col T)
    num_cols = df_raw.shape[1]
    final_headers = [""] * num_cols
    for col_i in range(num_cols):
        for row_i in range(mpf_row_idx + 1):  # scan rows 0 to mpf_row_idx inclusive
            v = clean_val(df_raw.iloc[row_i, col_i])
            if v and v != "nan":
                final_headers[col_i] = v  # keep overwriting — last non-empty wins

    print("  Final headers:", final_headers)

    # Make unique headers
    seen = {}
    unique_headers = []
    for h in final_headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(h + "_" + str(seen[h]))
        else:
            seen[h] = 0
            unique_headers.append(h)

    # Skip serial number row(s) after mpf_row_idx
    data_start = mpf_row_idx + 1
    for probe_idx in range(mpf_row_idx + 1, min(mpf_row_idx + 4, len(df_raw))):
        probe_vals = [clean_val(v) for v in df_raw.iloc[probe_idx].values if clean_val(v) not in ["", "nan"]]
        numeric_count = sum(1 for v in probe_vals if v.replace(".","",1).isdigit())
        if len(probe_vals) > 0 and numeric_count / len(probe_vals) > 0.7:
            data_start = probe_idx + 1
        else:
            break
    print("  Data starts at row index:", data_start)

    df_pr = df_raw.iloc[data_start:].copy()
    df_pr.columns = unique_headers
    df_pr = df_pr.reset_index(drop=True)

    missing_cols = [c for c in [COL_SUMNSUM, COL_FGNUM, COL_OPSOH, COL_MPF, COL_TVT, COL_UPF, COL_NPF] if c not in df_pr.columns]
    if missing_cols:
        print("  WARNING: Missing columns:", missing_cols, "- skipping sheet")
        continue

    mask = df_pr[COL_SUMNSUM].str.strip() != "Non-Sum"
    df_valid = df_pr[mask].copy()
    print("  Rows after Non-Sum filter:", len(df_valid))

    for col in [COL_MPF, COL_TVT, COL_UPF, COL_NPF, COL_OPSOH]:
        df_valid[col] = pd.to_numeric(df_valid[col], errors="coerce").fillna(0)

    # Store raw dispatch sum per FG Number for later calculation
    df_valid["_dispatch_sum"] = (
        df_valid[COL_MPF] + df_valid[COL_TVT] + df_valid[COL_UPF] + df_valid[COL_NPF]
    )

    for _, row in df_valid.iterrows():
        fg_num = clean_val(row[COL_FGNUM])
        if fg_num and fg_num != "nan":
            fg_map[fg_num] = {"op_soh": row[COL_OPSOH], "dispatch_sum": row["_dispatch_sum"]}

print("Total FG Number entries mapped:", len(fg_map))

# Build a Pack Size lookup for CBO rows: use Pack Size of HLP row immediately above
pack_size_for_cbo = {}
last_hlp_pack_size = None
for i in range(len(df_prog)):
    prod = df_prog.loc[i, "Product"].strip().upper()
    if prod == "HLP":
        last_hlp_pack_size = df_prog.loc[i, "Pack Size"].strip()
    elif prod == "CBO":
        pack_size_for_cbo[i] = last_hlp_pack_size

# Calculate Finished Goods row by row using Product and Pack Size
fg_values = []
matched = 0
for i, row in df_prog.iterrows():
    fg_num = row["VX-FG Code"].strip()
    if fg_num not in fg_map:
        fg_values.append(0)
        continue
    matched += 1
    op_soh       = fg_map[fg_num]["op_soh"]
    dispatch_sum = fg_map[fg_num]["dispatch_sum"]
    prod = row["Product"].strip().upper()
    if prod == "HLP":
        pack_size = pd.to_numeric(row["Pack Size"], errors="coerce")
        multiplier = 1000000
    else:  # CBO
        pack_size = pd.to_numeric(pack_size_for_cbo.get(i, 0), errors="coerce")
        multiplier = 10000000
    pack_size = pack_size if pd.notna(pack_size) and pack_size != 0 else 1
    fg_val = op_soh - ((dispatch_sum * pack_size) /multiplier)
    fg_values.append(fg_val)

df_prog["Finished Goods"] = fg_values
print("VX-FG Code matches found:", matched)

df_prog["WIP"] = 0

# ---- Required Cover Days ------------------------------------------
# HLP: Cat A=30, B=45, C=90 | CBO: same as HLP row above
cat_map = {"A": 30, "B": 45, "C": 90}
req_cover = []
last_hlp_req_cover = 0
for i in range(len(df_prog)):
    prod = df_prog.loc[i, "Product"].strip().upper()
    if prod == "HLP":
        cat = df_prog.loc[i, "Cat"].strip().upper()
        val = cat_map.get(cat, 0)
        last_hlp_req_cover = val
        req_cover.append(val)
    else:  # CBO
        req_cover.append(last_hlp_req_cover)
df_prog["Required Cover Days"] = req_cover

# ---- Actual Cover Days --------------------------------------------
# (Program / (Finished Goods + WIP)) * 30 | 0 if denominator is 0
prog_num = pd.to_numeric(df_prog["Program"],        errors="coerce").fillna(0)
fg_num   = pd.to_numeric(df_prog["Finished Goods"], errors="coerce").fillna(0)
wip_num  = pd.to_numeric(df_prog["WIP"],            errors="coerce").fillna(0)
denom    = fg_num + wip_num
df_prog["Actual Cover Days"] = (
    (denom.replace(0, float("nan")) / prog_num * 30)
    .fillna(0)
)

# ---- Additional Cover Required ------------------------------------
# Required - Actual, floor at 0
df_prog["Additional Cover Required"] = (
    (df_prog["Required Cover Days"] - df_prog["Actual Cover Days"])
    .clip(lower=0)
)

# ---- Total Balance To Print (modified) ----------------------------
# (Additional Cover Required / 30) * Program
df_prog["Total Balance to Print"] = (
    (((df_prog["Additional Cover Required"])+30) / 30) * prog_num
)

# ---- Printing Requirement In Tonnes --------------------------------------
# HLP: (GSM * Size * (Length/CylCirc) * TBP) / (PackSize * 1000 * 1000 * UPS)
# CBO: same but PackSize = HLP above & denominator uses 10000 instead of 1000
gsm_s    = pd.to_numeric(df_prog["Board GSM"],                    errors="coerce").fillna(0)
size_s   = pd.to_numeric(df_prog["Board Size"],                   errors="coerce").fillna(0)
lcc_s    = pd.to_numeric(df_prog["Board Length/Cyl Circumference"],errors="coerce").fillna(0)
tbp_s    = pd.to_numeric(df_prog["Total Balance to Print"],       errors="coerce").fillna(0)
ups_s    = pd.to_numeric(df_prog["No.Of UPS"],                    errors="coerce").fillna(1)
ups_s    = ups_s.replace(0, 1)  # avoid division by zero

# Pack size per row: HLP uses own, CBO uses HLP above (already in pack_size_for_cbo)
hlp_pack_s = pd.to_numeric(df_prog["Pack Size"], errors="coerce").fillna(1)

tonnes_vals = []
for i in range(len(df_prog)):
    prod = df_prog.loc[i, "Product"].strip().upper()
    gsm  = gsm_s.iloc[i]
    size = size_s.iloc[i]
    lcc  = lcc_s.iloc[i]
    tbp  = tbp_s.iloc[i]
    ups  = ups_s.iloc[i]
    if prod == "HLP":
        ps   = hlp_pack_s.iloc[i] if hlp_pack_s.iloc[i] != 0 else 1
        denom_t = ps * 1000 * 1000 * ups
    else:  # CBO
        ps_raw = pack_size_for_cbo.get(i, None)
        ps   = pd.to_numeric(ps_raw, errors="coerce") if ps_raw is not None else 1
        ps   = ps if pd.notna(ps) and ps != 0 else 1
        denom_t = ps * 10000 * 1000 * ups
    if denom_t == 0:
        tonnes_vals.append(0)
    else:
        tonnes_vals.append((gsm * size * lcc * tbp) / denom_t)

df_prog["Printing Requirement In Tonnes"] = tonnes_vals

write_excel(df_prog, PROGRAM_FILE, index=False, engine="openpyxl")
print("Program File updated ->", PROGRAM_FILE)
df_prog[["MSKU", "Product", "Program", "Finished Goods", "WIP",
         "Required Cover Days", "Actual Cover Days", "Additional Cover Required",
         "Total Balance to Print", "Printing Requirement In Tonnes"]].tail(10)



# ---- BOARD SUMMARY (with Gravure length-zeroing for grouping) --------------
BOARD_SUMMARY_FILE = rel_path("User Output", "Format Wise Summary", "Board Summary.xlsx")

df_prog = pd.read_excel(PROGRAM_FILE, dtype=str, engine="openpyxl").fillna("")

print("Program File rows loaded:", len(df_prog))

# Numeric columns needed for aggregation
NUM_COLS_AGG = ["Program", "Finished Goods", "WIP", "Total Balance to Print", "Printing Requirement In Tonnes"]
for col in NUM_COLS_AGG:
    df_prog[col] = pd.to_numeric(df_prog[col], errors="coerce").fillna(0)

# ---- CREATE GROUPING LENGTH COLUMN -----------------------------------------
# For Gravure / Gravure Sheeter rows: use "0" as length for grouping
# For all other routes: use actual Board Length/Cyl Circumference
# This does NOT modify df_prog — purely for grouping

GRAVURE_ROUTES = {"Gravure", "Gravure Sheeter"}

df_prog["_grp_length"] = df_prog.apply(
    lambda r: "0" if r["Route"].strip() in GRAVURE_ROUTES
              else str(r["Board Length/Cyl Circumference"]).strip(),
    axis=1
)

# ---- AGGREGATE -------------------------------------------------------------
GROUP_COLS = ["Pack Style", "Route", "Board GSM", "Board Size", "_grp_length", "Board Type"]

df_agg = (
    df_prog.groupby(GROUP_COLS, sort=False)[["Program", "Finished Goods", "WIP", "Printing Requirement In Tonnes"]]
    .sum()
    .reset_index()
    .rename(columns={"_grp_length": "Board Length/Cyl Circumference"})
)

# Balance To Print: sum of Total Balance to Print per group
tbp_agg = (
    df_prog.groupby(GROUP_COLS, sort=False)["Total Balance to Print"]
    .sum()
    .reset_index()
    .rename(columns={"_grp_length": "Board Length/Cyl Circumference", "Total Balance to Print": "Balance To Print"})
)
df_agg = df_agg.merge(tbp_agg, on=["Pack Style", "Route", "Board GSM", "Board Size", "Board Length/Cyl Circumference", "Board Type"], how="left")
df_agg["Balance To Print"] = df_agg["Balance To Print"].fillna(0)

# Final column order
df_agg = df_agg[[
    "Pack Style", "Route", "Board GSM", "Board Size",
    "Board Length/Cyl Circumference", "Board Type",
    "Program", "Finished Goods", "WIP", "Balance To Print", "Printing Requirement In Tonnes"
]]

# Sort
df_agg = df_agg.sort_values([
    "Pack Style", "Route", "Board GSM", "Board Size", "Board Length/Cyl Circumference"
]).reset_index(drop=True)

print("Summary rows:", len(df_agg))

# ---- WRITE FORMATTED EXCEL -------------------------------------------------
wb = open_workbook(BOARD_SUMMARY_FILE)
ws = wb.add_worksheet("Board Summary")

hdr_fmt   = wb.add_format({"bold":True,"font_name":"Arial","font_size":11,"bg_color":"#2E4057","font_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
ps_fmt    = wb.add_format({"bold":True,"font_name":"Arial","font_size":11,"bg_color":"#1B3A5C","font_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
ps_blank  = wb.add_format({"bg_color":"#1B3A5C","border":1,"border_color":"#CCCCCC"})
rt_fmt    = wb.add_format({"bold":True,"font_name":"Arial","font_size":10,"bg_color":"#D6E4F0","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
rt_blank  = wb.add_format({"bg_color":"#D6E4F0","border":1,"border_color":"#CCCCCC"})
gsm_fmt   = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#EBF5FB","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
gsm_blank = wb.add_format({"bg_color":"#EBF5FB","border":1,"border_color":"#CCCCCC"})
dat_fmt   = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
alt_fmt   = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#F0F4F8","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt   = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#FFFFFF","valign":"vcenter","num_format":"#,##0.00","border":1,"border_color":"#CCCCCC"})
num_alt   = wb.add_format({"font_name":"Arial","font_size":10,"bg_color":"#F0F4F8","valign":"vcenter","num_format":"#,##0.00","border":1,"border_color":"#CCCCCC"})
bal_pos   = wb.add_format({"bold":True,"font_name":"Arial","font_size":10,"bg_color":"#E8F5E9","font_color":"#1B5E20","valign":"vcenter","num_format":"#,##0.00","border":1,"border_color":"#CCCCCC"})
bal_neg   = wb.add_format({"bold":True,"font_name":"Arial","font_size":10,"bg_color":"#FFEBEE","font_color":"#B71C1C","valign":"vcenter","num_format":"#,##0.00","border":1,"border_color":"#CCCCCC"})

COLS     = ["Pack Style","Route","Board GSM","Board Size","Board Length/Cyl Circumference","Board Type","Program","Finished Goods","WIP","Balance To Print","Printing Requirement In Tonnes"]
NUM_COLS = {"Program","Finished Goods","WIP","Balance To Print","Printing Requirement In Tonnes"}

ws.set_row(0, 26)
for c, col in enumerate(COLS):
    ws.write(0, c, col, hdr_fmt)

row_num   = 1
prev_pack = None
prev_rt   = None
prev_gsm  = None
prev_size = None
prev_len  = None

for _, rec in df_agg.iterrows():
    ws.set_row(row_num, 18)
    pack  = rec["Pack Style"]
    route = rec["Route"]
    gsm   = rec["Board GSM"]

    size   = rec["Board Size"]
    length = rec["Board Length/Cyl Circumference"]

    is_new_pack = (pack, route) != (prev_pack, prev_rt)
    is_new_rt   = route  != prev_rt   or is_new_pack
    is_new_gsm  = gsm    != prev_gsm  or is_new_rt
    is_new_size = size   != prev_size or is_new_gsm
    is_new_len  = length != prev_len  or is_new_size

    show_route  = is_new_rt  or is_new_size
    show_gsm    = is_new_gsm or is_new_size
    rfmt = alt_fmt if row_num % 2 == 0 else dat_fmt
    nfmt = num_alt if row_num % 2 == 0 else num_fmt

    for c, col in enumerate(COLS):
        val = rec[col]
        if col == "Pack Style":
            ws.write(row_num, c, pack,  ps_fmt  if is_new_pack else ps_blank)
        elif col == "Route":
            ws.write(row_num, c, route if show_route else "", rt_fmt if show_route else rt_blank)
        elif col == "Board GSM":
            ws.write(row_num, c, gsm if show_gsm else "", gsm_fmt if show_gsm else gsm_blank)
        elif col == "Board Size":
            ws.write(row_num, c, size if is_new_size else "", rfmt)
        elif col == "Board Length/Cyl Circumference":
            ws.write(row_num, c, length if is_new_len else "", rfmt)
        elif col == "Balance To Print":
            ws.write(row_num, c, float(val), bal_neg if float(val) < 0 else bal_pos)
        elif col in NUM_COLS:
            ws.write(row_num, c, float(val), nfmt)
        else:
            ws.write(row_num, c, val, rfmt)

    prev_pack = pack
    prev_rt   = route
    prev_gsm  = gsm
    prev_size = size
    prev_len  = length
    row_num  += 1

col_widths = [18, 20, 12, 14, 14, 14, 15, 15, 10, 18, 22]
for c, w in enumerate(col_widths):
    ws.set_column(c, c, w)

ws.freeze_panes(1, 0)
ws.autofilter(0, 0, row_num - 1, len(COLS) - 1)
wb.close()

print("Board Summary File created ->", BOARD_SUMMARY_FILE)
print("Total summary rows:", len(df_agg))
df_agg.head(10)

AGEING_INPUT_FILE  = rel_path("User Input", "PPB RM", "Main.xlsx")  # Input file with Ageing Report sheet
AGEING_OUTPUT_FILE = rel_path("System Files", "System Generated PPB RM.xlsx")  # Output file
AGEING_SHEET       = "Ageing Report"    # Sheet name to read
MAT_GRP_COL        = "Material Group Desc."  # Column to filter on
FILTER_VALUE       = "Board-Tobacco"    # Value to keep
MATERIAL_COL       = "Material"         # Material code column
MAT_DESC_COL       = "Material Description"  # Material description column
require_file(AGEING_INPUT_FILE, "PPB RM ageing input file")


df_age = pd.read_excel(AGEING_INPUT_FILE, sheet_name=AGEING_SHEET, dtype=str, engine="openpyxl").fillna("")
print("Total rows loaded       :", len(df_age))
print("Columns                 :", df_age.columns.tolist())

df_bt = df_age[df_age[MAT_GRP_COL].str.strip().str.upper() == FILTER_VALUE.upper()].copy().reset_index(drop=True)
print("Board-Tobacco rows      :", len(df_bt))

# ---- SPLIT MATERIAL CODE INTO GSM / SIZE / LENGTH --------------------------
# 18-char: pos 6,7,8 = GSM | pos 12,13,14 = Size | pos 16,17,18 = Length
# 14-char: pos 6,7,8 = GSM | pos 12,13,14 = Size | Length = 0
# Anything else: GSM=0, Size=0, Length=0

def parse_material_code(code):
    c = str(code).strip()
    l = len(c)
    if l == 18:
        return c[5:8], c[11:14], c[15:18]
    elif l == 14:
        return c[5:8], c[11:14], "0"
    elif l == 15:
        return c[5:8], c[12:15], "0"
    else:
        return "0", "0", "0"

parsed = df_bt[MATERIAL_COL].apply(parse_material_code)
df_bt["Board GSM"]    = [x[0] for x in parsed]
df_bt["Board Size"]   = [x[1] for x in parsed]
df_bt["Board Length"] = [x[2] for x in parsed]

print("Sample parsed codes:")
print(df_bt[[MATERIAL_COL, "Board GSM", "Board Size", "Board Length"]].head(10).to_string(index=False))

# ---- CLASSIFY MATERIAL DESCRIPTION -----------------------------------------
def classify_board(name):
    n = str(name).upper()
    if "CROSS PILLAR" in n:                                          return "Transmet Cross"
    if "TRANSPOL" in n:                                             return "Transmet HOLO"
    if "TRANS" in n and any(x in n for x in ["PILLAR","FRESNEL","RAINBOW","LENS"]):
                                                                    return "Transmet HOLO"
    if ("TRANSMET" in n and any(x in n for x in ["SILVER","PLAIN"])) or "PLAIN SILVER TRANSMET" in n:
                                                                    return "Transmet Plain"
    if "FOIL" in n or "ALU FOIL" in n:                              return "Foil Board"
    if "INVERCOTE" in n or "LENATO" in n:                           return "Invercote Lenato"
    if any(x in n for x in ["CYBER XL","CY XL","CYXL","CYX"]):   return "CYXL"
    if any(x in n for x in ["CYPAK","CYCY","CY CY, CYBER CYPAK"]):             return "CYCY"
    if any(x in n for x in ["BCKI","BCK-I","BCFK-I","BCKE","BCFK","BCK"]):
                                                                    return "BCKI"
    if "CFK" in n:                                                  return "CFKI"
    if "BATABAK" in n:                                              return "Batabak"
    return "Others"

# Insert classification column immediately after Material Description
desc_idx = df_bt.columns.tolist().index(MAT_DESC_COL)
df_bt.insert(desc_idx + 1, "Board Type Category", df_bt[MAT_DESC_COL].apply(classify_board))

print("Classification distribution:")
print(df_bt["Board Type Category"].value_counts().to_string())


# ---- WRITE OUTPUT FILE ------------------------------------------------------
wb_out = open_workbook(AGEING_OUTPUT_FILE)
ws_out = wb_out.add_worksheet("Board Tobacco")

# Formats
hdr_fmt  = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,"bg_color":"#2E4057","font_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
new_fmt  = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,"bg_color":"#E6A817","font_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
dat_fmt  = wb_out.add_format({"font_name":"Arial","font_size":10,"bg_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
alt_fmt  = wb_out.add_format({"font_name":"Arial","font_size":10,"bg_color":"#F0F4F8","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
cat_fmt  = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,"bg_color":"#EAF3DE","font_color":"#27500A","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
cat_alt  = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,"bg_color":"#D4E8C2","font_color":"#27500A","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt  = wb_out.add_format({"font_name":"Arial","font_size":10,"bg_color":"#FFFFFF","valign":"vcenter","align":"center","border":1,"border_color":"#CCCCCC"})
num_alt  = wb_out.add_format({"font_name":"Arial","font_size":10,"bg_color":"#F0F4F8","valign":"vcenter","align":"center","border":1,"border_color":"#CCCCCC"})

NEW_COLS = {"Board GSM", "Board Size", "Board Length", "Board Type Category"}
cols     = df_bt.columns.tolist()

# Header row
ws_out.set_row(0, 24)
for ci, col in enumerate(cols):
    ws_out.write(0, ci, col, new_fmt if col in NEW_COLS else hdr_fmt)

# Data rows
for ri, (_, row_data) in enumerate(df_bt.iterrows(), start=1):
    ws_out.set_row(ri, 18)
    rfmt = alt_fmt if ri % 2 == 0 else dat_fmt
    nfmt = num_alt  if ri % 2 == 0 else num_fmt
    cf   = cat_alt  if ri % 2 == 0 else cat_fmt
    for ci, col in enumerate(cols):
        val = row_data[col]
        if col == "Board Type Category":
            ws_out.write(ri, ci, val, cf)
        elif col in {"Board GSM", "Board Size", "Board Length"}:
            ws_out.write(ri, ci, val, nfmt)
        else:
            ws_out.write(ri, ci, val, rfmt)

# Column widths
for ci, col in enumerate(cols):
    width = max(len(col) + 4, 14)
    if col in {"Board GSM", "Board Size", "Board Length"}: width = 13
    if col == "Board Type Category": width = 22
    if col == "Material Description": width = 50
    if col == MATERIAL_COL: width = 22
    ws_out.set_column(ci, ci, width)

ws_out.freeze_panes(1, 0)
ws_out.autofilter(0, 0, len(df_bt), len(cols) - 1)
wb_out.close()

print("Output file written ->", AGEING_OUTPUT_FILE)
print("Rows written         :", len(df_bt))
print("Columns              :", len(cols))

# ---- PPB RM BOARD TYPE SUMMARY ---------------------------------------------
import pandas as pd
import xlsxwriter

PPB_RM_SUMMARY_FILE = rel_path("System Files", "PPB RM System Summary.xlsx")

# ---- LOAD AGEING OUTPUT FILE -----------------------------------------------
df_ppb = pd.read_excel(AGEING_OUTPUT_FILE, dtype=str, engine="openpyxl").fillna("")

# Ensure numeric stock
df_ppb["_stock_num"] = pd.to_numeric(df_ppb["Total Stock(Batch-wise)"], errors="coerce").fillna(0)

# Strip whitespace from key columns
for col in ["Board Type Category", "Board GSM", "Board Size", "Board Length", "Plnt"]:
    df_ppb[col] = df_ppb[col].str.strip()

# ---- SORT ------------------------------------------------------------------
# Convert to numeric for sorting, keep original string for display
df_ppb["_gsm_n"]  = pd.to_numeric(df_ppb["Board GSM"],    errors="coerce").fillna(0)
df_ppb["_size_n"] = pd.to_numeric(df_ppb["Board Size"],   errors="coerce").fillna(0)
df_ppb["_len_n"]  = pd.to_numeric(df_ppb["Board Length"], errors="coerce").fillna(0)

df_ppb = df_ppb.sort_values(
    ["Board Type Category", "_gsm_n", "_size_n", "_len_n", "Plnt"]
).reset_index(drop=True)

# ---- BUILD OUTPUT ROWS -----------------------------------------------------
# Each output row: Board Type, Board GSM, Board Size, Board Length, Plant, Stock
# Structure:
#   [Section header row] Board Type spanning all columns
#   [Data rows]          repeated Board Type/GSM/Size/Length per plant
#   [Total row]          after each GSM+Size+Length group

GROUP_COLS = ["Board Type Category", "Board GSM", "Board Size", "Board Length"]

output_rows = []   # list of dicts with keys: type, board_type, gsm, size, length, plant, stock
                   # type: "header" | "data" | "total"

prev_board_type = None

# Group by Board Type first, then GSM+Size+Length, then Plant
for board_type, df_bt_grp in df_ppb.groupby("Board Type Category", sort=False):

    # Section header row for Board Type
    output_rows.append({
        "type": "header",
        "board_type": board_type,
        "gsm": "", "size": "", "length": "", "plant": "", "stock": ""
    })

    # Within board type, group by GSM + Size + Length
    for (gsm, size, length), df_grp in df_bt_grp.groupby(
        ["Board GSM", "Board Size", "Board Length"], sort=False
    ):
        # Within each GSM+Size+Length group, sum by Plant
        plant_totals = (
            df_grp.groupby("Plnt", sort=True)["_stock_num"]
            .sum()
            .reset_index()
            .rename(columns={"Plnt": "Plant", "_stock_num": "Stock"})
        )

        group_total = plant_totals["Stock"].sum()

        for _, pr in plant_totals.iterrows():
            output_rows.append({
                "type": "data",
                "board_type": board_type,
                "gsm":    gsm,
                "size":   size,
                "length": length,
                "plant":  pr["Plant"],
                "stock":  pr["Stock"]
            })

        # Total row after all plants for this group
        output_rows.append({
            "type": "total",
            "board_type": board_type,
            "gsm":    gsm,
            "size":   size,
            "length": length,
            "plant":  "Total",
            "stock":  group_total
        })

print("Total output rows (incl. headers & totals):", len(output_rows))

# ---- WRITE EXCEL -----------------------------------------------------------
wb_sum = open_workbook(PPB_RM_SUMMARY_FILE)
ws_sum = wb_sum.add_worksheet("PPB RM Summary")

COLS = ["Board Type", "Board GSM", "Board Size", "Board Length", "Plant", "Stock"]

# Formats
hdr_fmt   = wb_sum.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                "bg_color": "#2E4057", "font_color": "#FFFFFF",
                                "align": "center", "valign": "vcenter",
                                "border": 1, "border_color": "#CCCCCC"})

sec_fmt   = wb_sum.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                "bg_color": "#E6A817", "font_color": "#FFFFFF",
                                "align": "left", "valign": "vcenter",
                                "border": 1, "border_color": "#CCCCCC"})

sec_blank = wb_sum.add_format({"bg_color": "#E6A817",
                                "border": 1, "border_color": "#CCCCCC"})

dat_fmt   = wb_sum.add_format({"font_name": "Arial", "font_size": 10,
                                "bg_color": "#FFFFFF", "valign": "vcenter",
                                "border": 1, "border_color": "#CCCCCC"})

alt_fmt   = wb_sum.add_format({"font_name": "Arial", "font_size": 10,
                                "bg_color": "#F0F4F8", "valign": "vcenter",
                                "border": 1, "border_color": "#CCCCCC"})

tot_fmt   = wb_sum.add_format({"bold": True, "font_name": "Arial", "font_size": 10,
                                "bg_color": "#D6E4F0", "valign": "vcenter",
                                "border": 1, "border_color": "#CCCCCC"})

tot_num   = wb_sum.add_format({"bold": True, "font_name": "Arial", "font_size": 10,
                                "bg_color": "#D6E4F0", "valign": "vcenter",
                                "align": "center", "num_format": "#,##0.00",
                                "border": 1, "border_color": "#CCCCCC"})

num_fmt   = wb_sum.add_format({"font_name": "Arial", "font_size": 10,
                                "bg_color": "#FFFFFF", "valign": "vcenter",
                                "align": "center", "num_format": "#,##0.00",
                                "border": 1, "border_color": "#CCCCCC"})

num_alt   = wb_sum.add_format({"font_name": "Arial", "font_size": 10,
                                "bg_color": "#F0F4F8", "valign": "vcenter",
                                "align": "center", "num_format": "#,##0.00",
                                "border": 1, "border_color": "#CCCCCC"})

# Header row
ws_sum.set_row(0, 24)
for ci, col in enumerate(COLS):
    ws_sum.write(0, ci, col, hdr_fmt)

# Data rows
row_num    = 1
data_row_n = 0   # counter for alternating colors (only data rows, not headers/totals)

for rec in output_rows:

    ws_sum.set_row(row_num, 18)

    if rec["type"] == "header":
        # Board Type section header — value in col 0, blank cells for rest
        ws_sum.write(row_num, 0, rec["board_type"], sec_fmt)
        for ci in range(1, len(COLS)):
            ws_sum.write(row_num, ci, "", sec_blank)
        data_row_n = 0   # reset alternating counter per section

    elif rec["type"] == "data":
        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt
        ws_sum.write(row_num, 0, rec["board_type"], rfmt)
        ws_sum.write(row_num, 1, rec["gsm"],         rfmt)
        ws_sum.write(row_num, 2, rec["size"],         rfmt)
        ws_sum.write(row_num, 3, rec["length"],       rfmt)
        ws_sum.write(row_num, 4, rec["plant"],        rfmt)
        ws_sum.write(row_num, 5, rec["stock"],        nfmt)
        data_row_n += 1

    elif rec["type"] == "total":
        ws_sum.write(row_num, 0, rec["board_type"], tot_fmt)
        ws_sum.write(row_num, 1, rec["gsm"],         tot_fmt)
        ws_sum.write(row_num, 2, rec["size"],         tot_fmt)
        ws_sum.write(row_num, 3, rec["length"],       tot_fmt)
        ws_sum.write(row_num, 4, "Total",             tot_fmt)
        ws_sum.write(row_num, 5, rec["stock"],        tot_num)

    row_num += 1

# Column widths
col_widths = [22, 12, 12, 14, 10, 16]
for ci, w in enumerate(col_widths):
    ws_sum.set_column(ci, ci, w)

ws_sum.freeze_panes(1, 0)
ws_sum.autofilter(0, 0, row_num - 1, len(COLS) - 1)
wb_sum.close()

print("PPB RM Summary written ->", PPB_RM_SUMMARY_FILE)
print("Total rows written      :", row_num - 1)

# ---- PPB RM SUMMARY WITH BOARD SUMMARY LOOKUP ------------------------------

import pandas as pd
import xlsxwriter

# ---- LOAD BOTH FILES -------------------------------------------------------
df_ppb_raw = pd.read_excel(PPB_RM_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")
df_bs_raw  = pd.read_excel(BOARD_SUMMARY_FILE,  dtype=str, engine="openpyxl").fillna("")

# Strip key columns
for col in ["Board Type", "Board GSM", "Board Size", "Board Length",
            "Plant", "Stock"]:
    if col in df_ppb_raw.columns:
        df_ppb_raw[col] = df_ppb_raw[col].str.strip()

for col in ["Board Type", "Board GSM", "Board Size",
            "Board Length/Cyl Circumference",
            "Pack Style", "Route", "Printing Requirement In Tonnes"]:
    if col in df_bs_raw.columns:
        df_bs_raw[col] = df_bs_raw[col].str.strip()

# ---- BUILD SINGLE LOOKUP ---------------------------------------------------
# Key: (Board Type, Board GSM, Board Size, Board Length/Cyl Circumference)
# Value: list of match dicts {pack_style, route, tbt}
# Gravure rows naturally have length "0", Offset rows have non-zero length
# So PPB RM length "0" will only ever match Gravure rows and vice versa

bs_lookup = {}

for _, row in df_bs_raw.iterrows():
    btype  = row["Board Type"].strip()
    gsm    = row["Board GSM"].strip()
    size   = row["Board Size"].strip()
    length = row["Board Length/Cyl Circumference"].strip()
    ps     = row["Pack Style"].strip()
    route  = row["Route"].strip()
    tbt    = pd.to_numeric(row["Printing Requirement In Tonnes"], errors="coerce")
    tbt    = tbt if pd.notna(tbt) else 0

    # Skip header rows (GSM will be empty for section header rows)
    if gsm == "":
        continue

    key = (btype, gsm, size, length)
    bs_lookup.setdefault(key, []).append({
        "pack_style": ps,
        "route":      route,
        "tbt":        tbt
    })

print("Board Summary lookup keys:", len(bs_lookup))

# ---- PRE-COMPUTE FRACTIONS -------------------------------------------------
def attach_fractions(lookup):
    result = {}
    for key, matches in lookup.items():
        total_tbt = sum(m["tbt"] for m in matches)
        enriched  = []
        for m in matches:
            frac = (m["tbt"] / total_tbt) if total_tbt != 0 else 0
            enriched.append({**m, "fraction": frac})
        result[key] = enriched
    return result

bs_lookup = attach_fractions(bs_lookup)

# ---- DETERMINE ROW TYPES ---------------------------------------------------
def get_row_type(r):
    if r["Plant"].strip() == "Total":
        return "total"
    if r["Board GSM"].strip() == "" and r["Plant"].strip() == "":
        return "header"
    return "data"

# ---- REBUILD OUTPUT ROWS ---------------------------------------------------
PPB_COLS   = df_ppb_raw.columns.tolist()
output_rows = []

for _, row in df_ppb_raw.iterrows():
    rtype = get_row_type(row)
    base  = {col: row[col] for col in PPB_COLS}

    if rtype != "total":
        output_rows.append({**base, "row_type": rtype,
                             "pack_style": "", "route": "",
                             "tbt": "", "fraction": ""})
        continue

    # --- Total row: lookup ---
    btype  = row["Board Type"].strip()
    gsm    = row["Board GSM"].strip()
    size   = row["Board Size"].strip()
    length = row["Board Length"].strip()

    key     = (btype, gsm, size, length)
    matches = bs_lookup.get(key, [])

    if not matches:
        output_rows.append({**base, "row_type": "total",
                             "pack_style": "", "route": "",
                             "tbt": "", "fraction": 0})
    else:
        # First match on the Total row itself
        first = matches[0]
        output_rows.append({**base, "row_type": "total",
                             "pack_style": first["pack_style"],
                             "route":      first["route"],
                             "tbt":        first["tbt"],
                             "fraction":   first["fraction"]})

        # Additional matches as extra rows below
        for m in matches[1:]:
            extra = {col: "" for col in PPB_COLS}
            extra["Board Type"]   = btype
            extra["Board GSM"]    = gsm
            extra["Board Size"]   = size
            extra["Board Length"] = length
            extra["Plant"]        = ""
            extra["Stock"]        = ""
            output_rows.append({**extra, "row_type": "extra",
                                 "pack_style": m["pack_style"],
                                 "route":      m["route"],
                                 "tbt":        m["tbt"],
                                 "fraction":   m["fraction"]})

print("Total output rows:", len(output_rows))

# ---- WRITE EXCEL -----------------------------------------------------------
wb_out = open_workbook(PPB_RM_SUMMARY_FILE)
ws_out = wb_out.add_worksheet("PPB RM Summary")

ALL_COLS = ["Board Type", "Board GSM", "Board Size", "Board Length",
            "Plant", "Stock",
            "Pack Style", "Route", "Printing Requirement In Tonnes", "Fraction"]

# Formats
hdr_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
                              "bg_color":"#2E4057","font_color":"#FFFFFF",
                              "align":"center","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
sec_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
                              "bg_color":"#E6A817","font_color":"#FFFFFF",
                              "align":"left","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
sec_blk = wb_out.add_format({"bg_color":"#E6A817",
                              "border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
tot_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,
                              "bg_color":"#D6E4F0","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
tot_num = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,
                              "bg_color":"#D6E4F0","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
tot_pct = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,
                              "bg_color":"#D6E4F0","valign":"vcenter",
                              "align":"center","num_format":"0.00%",
                              "border":1,"border_color":"#CCCCCC"})
num_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
num_alt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
ext_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFF9E6","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
ext_num = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFF9E6","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
ext_pct = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFF9E6","valign":"vcenter",
                              "align":"center","num_format":"0.00%",
                              "border":1,"border_color":"#CCCCCC"})

# Header row
ws_out.set_row(0, 24)
for ci, col in enumerate(ALL_COLS):
    ws_out.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0

for rec in output_rows:
    ws_out.set_row(row_num, 18)
    rtype = rec["row_type"]

    if rtype == "header":
        ws_out.write(row_num, 0, rec["Board Type"], sec_fmt)
        for ci in range(1, len(ALL_COLS)):
            ws_out.write(row_num, ci, "", sec_blk)
        data_row_n = 0

    elif rtype == "data":
        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt
        ws_out.write(row_num, 0, rec["Board Type"],  rfmt)
        ws_out.write(row_num, 1, rec["Board GSM"],   rfmt)
        ws_out.write(row_num, 2, rec["Board Size"],  rfmt)
        ws_out.write(row_num, 3, rec["Board Length"],rfmt)
        ws_out.write(row_num, 4, rec["Plant"],       rfmt)
        stock_val = pd.to_numeric(rec["Stock"], errors="coerce")
        ws_out.write(row_num, 5,
                     stock_val if pd.notna(stock_val) else 0, nfmt)
        for ci in range(6, len(ALL_COLS)):
            ws_out.write(row_num, ci, "", rfmt)
        data_row_n += 1

    elif rtype == "total":
        ws_out.write(row_num, 0, rec["Board Type"],  tot_fmt)
        ws_out.write(row_num, 1, rec["Board GSM"],   tot_fmt)
        ws_out.write(row_num, 2, rec["Board Size"],  tot_fmt)
        ws_out.write(row_num, 3, rec["Board Length"],tot_fmt)
        ws_out.write(row_num, 4, "Total",            tot_fmt)
        stock_val = pd.to_numeric(rec["Stock"], errors="coerce")
        ws_out.write(row_num, 5,
                     stock_val if pd.notna(stock_val) else 0, tot_num)
        ws_out.write(row_num, 6, rec["pack_style"],  tot_fmt)
        ws_out.write(row_num, 7, rec["route"],       tot_fmt)
        tbt_val = rec["tbt"]
        ws_out.write(row_num, 8,
                     tbt_val if tbt_val != "" else 0, tot_num)
        frac_val = rec["fraction"]
        ws_out.write(row_num, 9,
                     frac_val if isinstance(frac_val, (int, float)) else 0,
                     tot_pct)

    elif rtype == "extra":
        ws_out.write(row_num, 0, rec["Board Type"],  ext_fmt)
        ws_out.write(row_num, 1, rec["Board GSM"],   ext_fmt)
        ws_out.write(row_num, 2, rec["Board Size"],  ext_fmt)
        ws_out.write(row_num, 3, rec["Board Length"],ext_fmt)
        ws_out.write(row_num, 4, "",                 ext_fmt)
        ws_out.write(row_num, 5, "",                 ext_fmt)
        ws_out.write(row_num, 6, rec["pack_style"],  ext_fmt)
        ws_out.write(row_num, 7, rec["route"],       ext_fmt)
        ws_out.write(row_num, 8, rec["tbt"],         ext_num)
        ws_out.write(row_num, 9, rec["fraction"],    ext_pct)

    row_num += 1

# Column widths
col_widths = [22, 12, 12, 14, 10, 16, 18, 20, 22, 12]
for ci, w in enumerate(col_widths):
    ws_out.set_column(ci, ci, w)

ws_out.freeze_panes(1, 0)
ws_out.autofilter(0, 0, row_num - 1, len(ALL_COLS) - 1)
wb_out.close()

print("PPB RM Summary (enriched) written ->", PPB_RM_SUMMARY_FILE)
print("Total rows written                 :", row_num - 1)

# ---- BOARD SUMMARY WITH PPB RM AVAILABLE -----------------------------------

import pandas as pd
import xlsxwriter

# ---- LOAD FILES ------------------------------------------------------------
df_bs  = pd.read_excel(BOARD_SUMMARY_FILE,  dtype=str, engine="openpyxl").fillna("")
df_ppb = pd.read_excel(PPB_RM_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")

# Strip all relevant columns
for col in df_bs.columns:
    df_bs[col] = df_bs[col].str.strip()
for col in df_ppb.columns:
    df_ppb[col] = df_ppb[col].str.strip()

# ---- BUILD PPB RM LOOKUPS --------------------------------------------------

# LOOKUP 1: (Pack Style, Route) -> list of fractions
# Key: (Pack Style, Route, Board Type, GSM, Size, Length) -> fraction
ps_route_fraction = {}

ppb_prev_btype  = ""
ppb_prev_gsm    = ""
ppb_prev_size   = ""
ppb_prev_length = ""

for _, row in df_ppb.iterrows():
    plant = row.get("Plant", "").strip()
    btype = row.get("Board Type",   "").strip()
    gsm   = row.get("Board GSM",    "").strip()
    size  = row.get("Board Size",   "").strip()
    length= row.get("Board Length", "").strip()

    # Carry forward board combo values
    btype  = btype  or ppb_prev_btype
    gsm    = gsm    or ppb_prev_gsm
    size   = size   or ppb_prev_size
    length = length or ppb_prev_length

    if row.get("Board Type",   "").strip(): ppb_prev_btype  = row.get("Board Type",   "").strip()
    if row.get("Board GSM",    "").strip(): ppb_prev_gsm    = row.get("Board GSM",    "").strip()
    if row.get("Board Size",   "").strip(): ppb_prev_size   = row.get("Board Size",   "").strip()
    if row.get("Board Length", "").strip(): ppb_prev_length = row.get("Board Length", "").strip()

    # Reset on section header
    if btype == "" and gsm == "" and plant == "":
        ppb_prev_btype = ppb_prev_gsm = ppb_prev_size = ppb_prev_length = ""
        continue

    # Only read from Total rows
    # Process Total rows AND extra match rows
    is_total = plant == "Total"
    is_extra = (plant == "" and
                row.get("Pack Style", "").strip() != "" and
                row.get("Fraction",   "").strip() != "")

    if not is_total and not is_extra:
        continue

    ps    = row.get("Pack Style", "").strip()
    route = row.get("Route",      "").strip()
    frac  = row.get("Fraction",   "").strip()
    if ps == "" or route == "" or frac == "":
        continue
    frac_val = pd.to_numeric(frac, errors="coerce")
    if pd.isna(frac_val):
        continue

    # Key includes board combo to avoid mixing fractions across different combos
    key = (ps, route, btype, gsm, size, length)
    ps_route_fraction[key] = frac_val

print("Pack Style+Route fraction keys:", len(ps_route_fraction))
print("Key check GS  :", ps_route_fraction.get(("CBO-KSFT16CP","Gravure Sheeter","CYCY","240","402","0"), "NOT FOUND"))
print("Key check Off :", ps_route_fraction.get(("CBO-KSFT16CP","Offset","CYCY","240","567","676"), "NOT FOUND"))

# LOOKUP 2: (Board Type, GSM, Size, Length) -> plant stocks and total stock
def get_ppb_row_type(r):
    if r.get("Plant", "").strip() == "Total":
        return "total"
    if r.get("Board GSM", "").strip() == "" and r.get("Plant", "").strip() == "":
        return "header"
    return "data"

combo_plant_stock = {}
combo_total_stock = {}

for _, row in df_ppb.iterrows():
    rtype  = get_ppb_row_type(row)
    btype  = row.get("Board Type",   "").strip()
    gsm    = row.get("Board GSM",    "").strip()
    size   = row.get("Board Size",   "").strip()
    length = row.get("Board Length", "").strip()
    plant  = row.get("Plant",        "").strip()
    stock  = pd.to_numeric(row.get("Stock", ""), errors="coerce")
    stock  = stock if pd.notna(stock) else 0

    if btype == "" or gsm == "" or size == "":
        continue

    key = (btype, gsm, size, length)

    if rtype == "data":
        combo_plant_stock.setdefault(key, []).append((plant, stock))
    elif rtype == "total":
        combo_total_stock[key] = stock

print("PPB RM combo keys:", len(combo_plant_stock))

# ---- DETERMINE BS ROW TYPE -------------------------------------------------
def get_bs_row_type(r):
    if (r.get("Board GSM", "").strip() == "" and
        r.get("Board Type", "").strip() == "" and
        r.get("Route", "").strip() == ""):
        return "header"
    return "data"

# ---- BUILD OUTPUT ROWS -----------------------------------------------------
BS_COLS      = df_bs.columns.tolist()
ALL_OUT_COLS = BS_COLS + ["PPB RM Available"]

output_rows = []

prev_gsm_val = ""
prev_btype   = ""

for _, bs_row in df_bs.iterrows():
    rtype = get_bs_row_type(bs_row)
    base  = {col: bs_row[col] for col in BS_COLS}
    base["PPB RM Available"] = ""

    if rtype == "header":
        output_rows.append({**base, "row_type": "header"})
        prev_gsm_val = ""
        prev_btype   = ""
        continue

    # Data row — append main row first
    output_rows.append({**base, "row_type": "data"})

    ps     = bs_row.get("Pack Style", "").strip()
    route  = bs_row.get("Route", "").strip()
    btype  = bs_row.get("Board Type", "").strip() or prev_btype
    gsm    = bs_row.get("Board GSM", "").strip()   or prev_gsm_val
    size   = bs_row.get("Board Size", "").strip()
    length = bs_row.get("Board Length/Cyl Circumference", "").strip()

    if bs_row.get("Board Type", "").strip(): prev_btype   = bs_row.get("Board Type", "").strip()
    if bs_row.get("Board GSM",  "").strip(): prev_gsm_val = bs_row.get("Board GSM",  "").strip()
    
    key_ps    = (ps, route, btype, gsm, size, length)
    key_combo = (btype, gsm, size, length)

    frac_val  = ps_route_fraction.get(key_ps, None)
    fractions = [frac_val] if frac_val is not None else []
    plant_rows  = combo_plant_stock.get(key_combo, [])
    total_stock = combo_total_stock.get(key_combo, 0)

    if not fractions or not plant_rows:
        sub = {col: "" for col in BS_COLS}
        sub["PPB RM Available"] = 0
        sub["_plant_label"]     = "No Stock"
        sub["row_type"]         = "sub_zero"
        output_rows.append(sub)
        continue

    # Compute per-plant PPB RM Available (divide by 1000 for tonnes)
    plant_avail = {}
    for (plant, stock) in plant_rows:
        avail = sum(stock * f for f in fractions) / 1000
        plant_avail[plant] = plant_avail.get(plant, 0) + avail

    total_avail = sum(total_stock * f for f in fractions) / 1000

    # Plant sub-rows (only non-zero)
    has_any = False
    for plant, avail in plant_avail.items():
        if avail == 0:
            continue
        sub = {col: "" for col in BS_COLS}
        sub["PPB RM Available"] = avail
        sub["_plant_label"]     = plant
        sub["row_type"]         = "sub"
        has_any = True
        output_rows.append(sub)

    if not has_any:
        sub = {col: "" for col in BS_COLS}
        sub["PPB RM Available"] = 0
        sub["_plant_label"]     = "No Stock"
        sub["row_type"]         = "sub_zero"
        output_rows.append(sub)
        continue

    # Total sub-row
    sub_tot = {col: "" for col in BS_COLS}
    sub_tot["PPB RM Available"] = total_avail
    sub_tot["_plant_label"]     = "Total"
    sub_tot["row_type"]         = "sub_total"
    output_rows.append(sub_tot)

print("Total output rows:", len(output_rows))

for i, r in enumerate(output_rows):
    bt   = str(r.get("Board Type",""))
    size = str(r.get("Board Size",""))
    ps   = str(r.get("Pack Style",""))
    if "DSFT10HL" in ps or "592" in size or (bt == "CFKI" and size in ["592",""]):
        print(i, r.get("row_type",""), ps, r.get("Route",""), bt, 
              r.get("Board GSM",""), size, r.get("Board Length/Cyl Circumference",""),
              r.get("_plant_label",""), r.get("PPB RM Available",""))

# ---- WRITE EXCEL -----------------------------------------------------------
wb_new = open_workbook(BOARD_SUMMARY_FILE)
ws_new = wb_new.add_worksheet("Board Summary")

COLS     = ALL_OUT_COLS
NUM_COLS = {"Program", "Finished Goods", "WIP",
            "Balance To Print", "Printing Requirement In Tonnes", "PPB RM Available"}

# Formats
hdr_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 11,
    "bg_color": "#2E4057", "font_color": "#FFFFFF",
    "align": "center", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
ps_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 11,
    "bg_color": "#1B3A5C", "font_color": "#FFFFFF",
    "valign": "vcenter", "border": 1, "border_color": "#CCCCCC"
})
rt_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#D6E4F0", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
gsm_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#EBF5FB", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
dat_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFFFFF", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
alt_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0F4F8", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
num_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFFFFF", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
num_alt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0F4F8", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
bal_pos = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#E8F5E9", "font_color": "#1B5E20",
    "valign": "vcenter", "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
bal_neg = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFEBEE", "font_color": "#B71C1C",
    "valign": "vcenter", "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
sub_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0FFF0", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
sub_num = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0FFF0", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
sub_tot_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#CCFFCC", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
sub_tot_num = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#CCFFCC", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
sub_zero_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFF0F0", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})

# Header row
ws_new.set_row(0, 26)
for ci, col in enumerate(COLS):
    ws_new.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0
prev_pack  = None
prev_rt    = None
prev_gsm   = None

for rec in output_rows:
    ws_new.set_row(row_num, 18)
    rtype = rec["row_type"]

    if rtype == "header":
        ws_new.write(row_num, 0, rec.get("Pack Style", ""), ps_fmt)
        for ci in range(1, len(COLS)):
            ws_new.write(row_num, ci, "", ps_fmt)
        data_row_n = 0
        prev_pack  = None
        prev_rt    = None
        prev_gsm   = None

    elif rtype == "data":
        pack  = rec.get("Pack Style", "")
        route = rec.get("Route", "")
        gsm   = rec.get("Board GSM", "")

        is_new_pack = (pack, route) != (prev_pack, prev_rt)
        is_new_rt   = route != prev_rt  or is_new_pack
        is_new_gsm  = gsm   != prev_gsm or is_new_rt

        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt

        for ci, col in enumerate(COLS):
            val = rec.get(col, "")
            if col == "Pack Style":
                ws_new.write(row_num, ci, pack, ps_fmt)
            elif col == "Route":
                ws_new.write(row_num, ci,
                             route if is_new_rt else "", rt_fmt)
            elif col == "Board GSM":
                ws_new.write(row_num, ci,
                             gsm if is_new_gsm else "", gsm_fmt)
            elif col == "Balance To Print":
                fval = pd.to_numeric(val, errors="coerce")
                fval = fval if pd.notna(fval) else 0
                ws_new.write(row_num, ci, fval,
                             bal_neg if fval < 0 else bal_pos)
            elif col in NUM_COLS:
                fval = pd.to_numeric(val, errors="coerce")
                ws_new.write(row_num, ci,
                             fval if pd.notna(fval) else 0, nfmt)
            else:
                ws_new.write(row_num, ci, val, rfmt)

        prev_pack  = pack
        prev_rt    = route
        prev_gsm   = gsm
        data_row_n += 1

    elif rtype == "sub":
        plant_label = rec.get("_plant_label", "")
        avail       = rec.get("PPB RM Available", 0)
        for ci, col in enumerate(COLS):
            if col == "PPB RM Available":
                ws_new.write(row_num, ci,
                             float(avail) if avail != "" else 0, sub_num)
            else:
                ws_new.write(row_num, ci, "", sub_fmt)
        ws_new.write(row_num, 0, plant_label, sub_fmt)

    elif rtype == "sub_total":
        plant_label = rec.get("_plant_label", "Total")
        avail       = rec.get("PPB RM Available", 0)
        for ci, col in enumerate(COLS):
            if col == "PPB RM Available":
                ws_new.write(row_num, ci,
                             float(avail) if avail != "" else 0, sub_tot_num)
            else:
                ws_new.write(row_num, ci, "", sub_tot_fmt)
        ws_new.write(row_num, 0, plant_label, sub_tot_fmt)

    elif rtype == "sub_zero":
        for ci, col in enumerate(COLS):
            if col == "PPB RM Available":
                ws_new.write(row_num, ci, 0, sub_zero_fmt)
            else:
                ws_new.write(row_num, ci, "", sub_zero_fmt)
        ws_new.write(row_num, 0, "No Stock", sub_zero_fmt)

    row_num += 1

# Column widths
col_widths = [18, 20, 12, 14, 14, 14, 15, 15, 10, 18, 22, 18]
for ci, w in enumerate(col_widths):
    ws_new.set_column(ci, ci, w)

ws_new.freeze_panes(1, 0)
ws_new.autofilter(0, 0, row_num - 1, len(COLS) - 1)
print("Sub rows for CFKI/200/592/0:")
for r in output_rows:
    if r.get("_plant_label","") in ["AHM1","MPF1","Total"] or r.get("PPB RM Available","") != "":
        print(r.get("row_type",""), r.get("_plant_label",""), r.get("PPB RM Available",""))
wb_new.close()

print("Board Summary (with PPB RM Available) written ->", BOARD_SUMMARY_FILE)
print("Total rows written:", row_num - 1)
print("combo_plant_stock for CFKI/200/592/0:")
print(combo_plant_stock.get(("CFKI","200","592","0"), []))

# ---- SYSTEM GENERATED PSPD RM REELS FILE -----------------------------------

PSPD_BHD_FILE    = rel_path("User Input", "PSPD RM", "Bhadrachalam.xlsx")
PSPD_GTR_FILE    = rel_path("User Input", "PSPD RM", "Guntur.xlsx")
PSPD_OUTPUT_FILE = rel_path("System Files", "System Generated PSPD RM Reels.xlsx")
require_file(PSPD_BHD_FILE, "PSPD Bhadrachalam input file")
require_file(PSPD_GTR_FILE, "PSPD Guntur input file")

COLS_NEEDED = ["Product Code", "GSM", "Width", "Length", "Net Wt.","Age"]

# ---- LOAD BHADRACHALAM (header row 1, data from row 3) ---------------------
df_bhd = pd.read_excel(
    PSPD_BHD_FILE,
    header=0,        # row 1 = index 0
    skiprows=[1],    # skip row 2 (empty)
    dtype=str,
    engine="openpyxl"
).fillna("")
df_bhd.columns = df_bhd.columns.str.strip()
print("Bhadrachalam rows loaded:", len(df_bhd))
print("Bhadrachalam columns    :", df_bhd.columns.tolist())

# ---- LOAD GUNTUR sheet 2 (header row 5, data from row 8) -------------------
df_gtr_raw = pd.read_excel(
    PSPD_GTR_FILE,
    sheet_name=1,    # 2nd sheet (0-indexed)
    header=None,
    dtype=str,
    engine="openpyxl"
).fillna("")

# Row index 4 = 5th row = headers
gtr_headers = [str(v).strip() for v in df_gtr_raw.iloc[4].values]
# Data starts from row index 7 = 8th row
df_gtr = df_gtr_raw.iloc[7:].copy()
df_gtr.columns = [str(c) for c in gtr_headers]
df_gtr = df_gtr.reset_index(drop=True)
df_gtr = df_gtr.astype(str)
df_gtr = df_gtr.reset_index(drop=True)
print("Guntur rows loaded      :", len(df_gtr))
print("Guntur columns          :", df_gtr.columns.tolist())

# ---- STRIP ALL VALUES ------------------------------------------------------
df_bhd = df_bhd.map(lambda x: str(x).strip())
df_gtr = df_gtr.map(lambda x: str(x).strip())
# ---- VERIFY REQUIRED COLUMNS EXIST ----------------------------------------
for src, df, name in [(COLS_NEEDED, df_bhd, "Bhadrachalam"),
                       (COLS_NEEDED, df_gtr, "Guntur")]:
    missing = [c for c in src if c not in df.columns]
    if missing:
        print(f"WARNING: {name} missing columns: {missing}")

# ---- EXTRACT NEEDED COLUMNS ------------------------------------------------
df_bhd_slim = df_bhd[COLS_NEEDED].copy()
df_gtr_slim = df_gtr[COLS_NEEDED].copy()

# Add source tag for traceability
df_bhd_slim["Source"] = "Bhadrachalam"
df_gtr_slim["Source"] = "Guntur"

# ---- COMBINE ---------------------------------------------------------------
df_combined = pd.concat([df_bhd_slim, df_gtr_slim], ignore_index=True)

# Drop rows where Product Code is blank or nan
df_combined = df_combined[
    ~df_combined["Product Code"].str.strip().isin(["", "nan", "None"])
].reset_index(drop=True)

print("Combined rows           :", len(df_combined))

# ---- CLASSIFY BOARD TYPE FROM PRODUCT CODE ---------------------------------
def classify_reel_product(code):
    c = str(code).strip().upper()
    if "CYBER CYPAK" in c or "CYCY" in c:
        return "CYCY"
    if "BCK" in c:
        return "BCKI"
    if "CFK" in c:
        return "CFKI"
    return "Others"

df_combined.insert(1, "Board Type Category",
                   df_combined["Product Code"].apply(classify_reel_product))

print("Classification distribution:")
print(df_combined["Board Type Category"].value_counts().to_string())

# ---- PARSE NUMERICS --------------------------------------------------------
df_combined["_gsm_raw"]   = pd.to_numeric(df_combined["GSM"],    errors="coerce").fillna(0)
df_combined["_width_raw"] = pd.to_numeric(df_combined["Width"].str.replace(r"[^\d.]", "", regex=True), errors="coerce").fillna(0)
df_combined["_len_raw"]   = pd.to_numeric(df_combined["Length"], errors="coerce").fillna(0)
df_combined["_netwt_raw"] = pd.to_numeric(df_combined["Net Wt."],errors="coerce").fillna(0)
df_combined["_age_raw"] = pd.to_numeric(df_combined["Age"],errors="coerce").fillna(0)

# ---- DERIVE BOARD COLUMNS --------------------------------------------------
# Width × 10 → Board Size (cm to mm)
# Length stays as-is (already mm)
# GSM as-is

df_combined["Board GSM"]    = df_combined["_gsm_raw"].apply(
    lambda x: str(int(x)) if x != 0 else "0"
)
df_combined["Board Size"]   = (df_combined["_width_raw"] * 10).apply(
    lambda x: str(int(x)) if x != 0 else "0"
)
df_combined["Board Length"] = df_combined["_len_raw"].apply(
    lambda x: "0" if x == 0 else str(int(x))
)
df_combined["Net Wt."]      = df_combined["_netwt_raw"]
df_combined["Age"]      = df_combined["_age_raw"]

# Drop temp columns
df_combined.drop(columns=["_gsm_raw", "_width_raw", "_len_raw", "_netwt_raw","_age_raw"],
                 inplace=True)

# ---- REORDER COLUMNS -------------------------------------------------------
# Final order: Product Code | Board Type Category | Board GSM | Board Size |
#              Board Length | Net Wt. | Source
final_cols = ["Product Code", "Board Type Category",
              "Board GSM", "Board Size", "Board Length",
              "Net Wt.", "Source","Age"]
# Keep any extra original columns at the end
extra_cols = [c for c in df_combined.columns if c not in final_cols]
df_combined = df_combined[final_cols + extra_cols]

print("Sample derived values:")
print(df_combined[["Product Code", "Board Type Category",
                   "Board GSM", "Board Size", "Board Length",
                   "Net Wt.","Age"]].head(10).to_string(index=False))

# ---- WRITE OUTPUT FILE -----------------------------------------------------
wb_out = open_workbook(PSPD_OUTPUT_FILE)
ws_out = wb_out.add_worksheet("PSPD RM Reels")

cols     = df_combined.columns.tolist()
NEW_COLS = {"Board GSM", "Board Size", "Board Length",
            "Board Type Category", "Net Wt.","Age"}

hdr_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
                              "bg_color":"#2E4057","font_color":"#FFFFFF",
                              "align":"center","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
new_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
                              "bg_color":"#E6A817","font_color":"#FFFFFF",
                              "align":"center","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
cat_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,
                              "bg_color":"#EAF3DE","font_color":"#27500A",
                              "valign":"vcenter","border":1,"border_color":"#CCCCCC"})
cat_alt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":10,
                              "bg_color":"#D4E8C2","font_color":"#27500A",
                              "valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
num_alt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
num_alt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})

# Header row
ws_out.set_row(0, 24)
for ci, col in enumerate(cols):
    ws_out.write(0, ci, col, new_fmt if col in NEW_COLS else hdr_fmt)

# Data rows
for ri, (_, row_data) in enumerate(df_combined.iterrows(), start=1):
    ws_out.set_row(ri, 18)
    rfmt = alt_fmt if ri % 2 == 0 else dat_fmt
    nfmt = num_alt if ri % 2 == 0 else num_fmt
    cf   = cat_alt if ri % 2 == 0 else cat_fmt
    for ci, col in enumerate(cols):
        val = row_data[col]
        if col == "Board Type Category":
            ws_out.write(ri, ci, val, cf)
        elif col in {"Board GSM", "Board Size", "Board Length","Age"}:
            ws_out.write(ri, ci, val, nfmt)
        elif col == "Net Wt.":
            fval = pd.to_numeric(val, errors="coerce")
            ws_out.write(ri, ci, fval if pd.notna(fval) else 0, nfmt)
        else:
            ws_out.write(ri, ci, val, rfmt)

# Column widths
for ci, col in enumerate(cols):
    width = max(len(col) + 4, 14)
    if col in {"Board GSM", "Board Size", "Board Length","Age"}: width = 13
    if col == "Board Type Category": width = 22
    if col == "Product Code": width = 35
    if col == "Source": width = 16
    ws_out.set_column(ci, ci, width)

ws_out.freeze_panes(1, 0)
ws_out.autofilter(0, 0, len(df_combined), len(cols) - 1)
wb_out.close()

print("Output file written ->", PSPD_OUTPUT_FILE)
print("Rows written        :", len(df_combined))
print("Columns             :", len(cols))

# ---- PSPD RM SYSTEM SUMMARY ------------------------------------------------

PSPD_SUMMARY_FILE = rel_path("System Files", "PSPD RM System Summary.xlsx")

# ---- LOAD PSPD OUTPUT FILE -------------------------------------------------
df_pspd_out = pd.read_excel(PSPD_OUTPUT_FILE, dtype=str, engine="openpyxl").fillna("")

for col in df_pspd_out.columns:
    df_pspd_out[col] = df_pspd_out[col].str.strip()

# Numeric stock
df_pspd_out["_stock_num"] = pd.to_numeric(
    df_pspd_out["Net Wt."], errors="coerce"
).fillna(0)

# Numeric sort keys
df_pspd_out["_gsm_n"]  = pd.to_numeric(df_pspd_out["Board GSM"],    errors="coerce").fillna(0)
df_pspd_out["_size_n"] = pd.to_numeric(df_pspd_out["Board Size"],   errors="coerce").fillna(0)
df_pspd_out["_len_n"]  = pd.to_numeric(df_pspd_out["Board Length"], errors="coerce").fillna(0)

df_pspd_out = df_pspd_out.sort_values(
    ["Board Type Category", "_gsm_n", "_size_n", "_len_n"]
).reset_index(drop=True)

# ---- BUILD OUTPUT ROWS -----------------------------------------------------
# Structure:
#   [Section header] Board Type spanning all columns
#   [Data row]       One row per GSM+Size+Length with Total Stock
# No plant breakdown — just one total per combination

output_rows = []

for board_type, df_bt in df_pspd_out.groupby("Board Type Category", sort=False):

    # Section header
    output_rows.append({
        "row_type":   "header",
        "board_type": board_type,
        "gsm":        "",
        "size":       "",
        "length":     "",
        "stock":      ""
    })

    # Group by GSM + Size + Length, sum stock
    for (gsm, size, length), df_grp in df_bt.groupby(
        ["Board GSM", "Board Size", "Board Length"], sort=False
    ):
        total_stock = df_grp["_stock_num"].sum()
        output_rows.append({
            "row_type":   "data",
            "board_type": board_type,
            "gsm":        gsm,
            "size":       size,
            "length":     length,
            "stock":      total_stock
        })

print("Total output rows:", len(output_rows))

# ---- WRITE EXCEL -----------------------------------------------------------
wb_sum = open_workbook(PSPD_SUMMARY_FILE)
ws_sum = wb_sum.add_worksheet("PSPD Summary")

COLS = ["Board Type", "Board GSM", "Board Size", "Board Length", "Total Stock"]

# Formats
hdr_fmt = wb_sum.add_format({
    "bold": True, "font_name": "Arial", "font_size": 11,
    "bg_color": "#2E4057", "font_color": "#FFFFFF",
    "align": "center", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
sec_fmt = wb_sum.add_format({
    "bold": True, "font_name": "Arial", "font_size": 11,
    "bg_color": "#E6A817", "font_color": "#FFFFFF",
    "align": "left", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
sec_blk = wb_sum.add_format({
    "bg_color": "#E6A817",
    "border": 1, "border_color": "#CCCCCC"
})
dat_fmt = wb_sum.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFFFFF", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
alt_fmt = wb_sum.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0F4F8", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
num_fmt = wb_sum.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFFFFF", "valign": "vcenter",
    "align": "center", "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
num_alt = wb_sum.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0F4F8", "valign": "vcenter",
    "align": "center", "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})

# Header row
ws_sum.set_row(0, 24)
for ci, col in enumerate(COLS):
    ws_sum.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0

for rec in output_rows:
    ws_sum.set_row(row_num, 18)

    if rec["row_type"] == "header":
        ws_sum.write(row_num, 0, rec["board_type"], sec_fmt)
        for ci in range(1, len(COLS)):
            ws_sum.write(row_num, ci, "", sec_blk)
        data_row_n = 0

    elif rec["row_type"] == "data":
        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt
        ws_sum.write(row_num, 0, rec["board_type"], rfmt)
        ws_sum.write(row_num, 1, rec["gsm"],         rfmt)
        ws_sum.write(row_num, 2, rec["size"],         rfmt)
        ws_sum.write(row_num, 3, rec["length"],       rfmt)
        ws_sum.write(row_num, 4, rec["stock"],        nfmt)
        data_row_n += 1

    row_num += 1

# Column widths
col_widths = [22, 12, 12, 14, 16]
for ci, w in enumerate(col_widths):
    ws_sum.set_column(ci, ci, w)

ws_sum.freeze_panes(1, 0)
ws_sum.autofilter(0, 0, row_num - 1, len(COLS) - 1)
wb_sum.close()

print("PSPD Summary written ->", PSPD_SUMMARY_FILE)
print("Total rows written   :", row_num - 1)

# ---- PSPD SUMMARY WITH BOARD SUMMARY LOOKUP --------------------------------

import pandas as pd
import xlsxwriter

# ---- LOAD BOTH FILES -------------------------------------------------------
df_pspd_sum = pd.read_excel(PSPD_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")
df_bs_raw   = pd.read_excel(BOARD_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")

for col in df_pspd_sum.columns:
    df_pspd_sum[col] = df_pspd_sum[col].str.strip()
for col in df_bs_raw.columns:
    df_bs_raw[col] = df_bs_raw[col].str.strip()

# ---- BUILD BOARD SUMMARY LOOKUP --------------------------------------------
# Key: (Board Type, Board GSM, Board Size, Board Length/Cyl Circumference)
# Value: list of {pack_style, route, tbt, fraction}

bs_lookup = {}

for _, row in df_bs_raw.iterrows():
    btype  = row.get("Board Type", "").strip()
    gsm    = row.get("Board GSM", "").strip()
    size   = row.get("Board Size", "").strip()
    length = row.get("Board Length/Cyl Circumference", "").strip()
    ps     = row.get("Pack Style", "").strip()
    route  = row.get("Route", "").strip()
    tbt    = pd.to_numeric(row.get("Printing Requirement In Tonnes", ""), errors="coerce")
    tbt    = tbt if pd.notna(tbt) else 0

    if gsm == "":
        continue

    key = (btype, gsm, size, length)
    bs_lookup.setdefault(key, []).append({
        "pack_style": ps,
        "route":      route,
        "tbt":        tbt
    })

# Attach fractions
def attach_fractions(lookup):
    result = {}
    for key, matches in lookup.items():
        total_tbt = sum(m["tbt"] for m in matches)
        enriched  = []
        for m in matches:
            frac = (m["tbt"] / total_tbt) if total_tbt != 0 else 0
            enriched.append({**m, "fraction": frac})
        result[key] = enriched
    return result

bs_lookup = attach_fractions(bs_lookup)
print("Board Summary lookup keys:", len(bs_lookup))

# ---- DETERMINE PSPD SUMMARY ROW TYPES -------------------------------------
def get_pspd_row_type(r):
    if r.get("Board GSM", "").strip() == "" and r.get("Total Stock", "").strip() == "":
        return "header"
    return "data"

# ---- BUILD OUTPUT ROWS -----------------------------------------------------
PSPD_COLS   = df_pspd_sum.columns.tolist()
output_rows = []

for _, row in df_pspd_sum.iterrows():
    rtype = get_pspd_row_type(row)
    base  = {col: row[col] for col in PSPD_COLS}

    if rtype == "header":
        output_rows.append({**base, "row_type": "header",
                             "pack_style": "", "route": "",
                             "tbt": "", "fraction": ""})
        continue

    # Data row — lookup
    btype  = row.get("Board Type", "").strip()
    gsm    = row.get("Board GSM",  "").strip()
    size   = row.get("Board Size", "").strip()
    length = row.get("Board Length", "").strip()

    key     = (btype, gsm, size, length)
    matches = bs_lookup.get(key, [])

    if not matches:
        output_rows.append({**base, "row_type": "data",
                             "pack_style": "", "route": "",
                             "tbt": "", "fraction": 0})
    else:
        # First match on data row itself
        first = matches[0]
        output_rows.append({**base, "row_type": "data",
                             "pack_style": first["pack_style"],
                             "route":      first["route"],
                             "tbt":        first["tbt"],
                             "fraction":   first["fraction"]})

        # Additional matches as extra rows below
        for m in matches[1:]:
            extra = {col: "" for col in PSPD_COLS}
            extra["Board Type"]   = btype
            extra["Board GSM"]    = gsm
            extra["Board Size"]   = size
            extra["Board Length"] = length
            extra["Total Stock"]  = ""
            output_rows.append({**extra, "row_type": "extra",
                                 "pack_style": m["pack_style"],
                                 "route":      m["route"],
                                 "tbt":        m["tbt"],
                                 "fraction":   m["fraction"]})

print("Total output rows:", len(output_rows))

# ---- WRITE EXCEL -----------------------------------------------------------
wb_out = open_workbook(PSPD_SUMMARY_FILE)
ws_out = wb_out.add_worksheet("PSPD Summary")

ALL_COLS = ["Board Type", "Board GSM", "Board Size", "Board Length",
            "Total Stock", "Pack Style", "Route",
            "Printing Requirement In Tonnes", "Fraction"]

# Formats
hdr_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
                              "bg_color":"#2E4057","font_color":"#FFFFFF",
                              "align":"center","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
sec_fmt = wb_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
                              "bg_color":"#E6A817","font_color":"#FFFFFF",
                              "align":"left","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
sec_blk = wb_out.add_format({"bg_color":"#E6A817",
                              "border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
num_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
num_alt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
pct_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFFFFF","valign":"vcenter",
                              "align":"center","num_format":"0.00%",
                              "border":1,"border_color":"#CCCCCC"})
pct_alt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#F0F4F8","valign":"vcenter",
                              "align":"center","num_format":"0.00%",
                              "border":1,"border_color":"#CCCCCC"})
ext_fmt = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFF9E6","valign":"vcenter",
                              "border":1,"border_color":"#CCCCCC"})
ext_num = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFF9E6","valign":"vcenter",
                              "align":"center","num_format":"#,##0.00",
                              "border":1,"border_color":"#CCCCCC"})
ext_pct = wb_out.add_format({"font_name":"Arial","font_size":10,
                              "bg_color":"#FFF9E6","valign":"vcenter",
                              "align":"center","num_format":"0.00%",
                              "border":1,"border_color":"#CCCCCC"})

# Header row
ws_out.set_row(0, 24)
for ci, col in enumerate(ALL_COLS):
    ws_out.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0

for rec in output_rows:
    ws_out.set_row(row_num, 18)
    rtype = rec["row_type"]

    if rtype == "header":
        ws_out.write(row_num, 0, rec.get("Board Type", ""), sec_fmt)
        for ci in range(1, len(ALL_COLS)):
            ws_out.write(row_num, ci, "", sec_blk)
        data_row_n = 0

    elif rtype == "data":
        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt
        pfmt = pct_alt if data_row_n % 2 == 0 else pct_fmt
        ws_out.write(row_num, 0, rec.get("Board Type", ""),  rfmt)
        ws_out.write(row_num, 1, rec.get("Board GSM", ""),   rfmt)
        ws_out.write(row_num, 2, rec.get("Board Size", ""),  rfmt)
        ws_out.write(row_num, 3, rec.get("Board Length", ""),rfmt)
        stock_val = pd.to_numeric(rec.get("Total Stock", ""), errors="coerce")
        ws_out.write(row_num, 4,
                     stock_val if pd.notna(stock_val) else 0, nfmt)
        ws_out.write(row_num, 5, rec.get("pack_style", ""),  rfmt)
        ws_out.write(row_num, 6, rec.get("route", ""),       rfmt)
        tbt_val = rec.get("tbt", "")
        ws_out.write(row_num, 7,
                     tbt_val if tbt_val != "" else 0, nfmt)
        frac_val = rec.get("fraction", 0)
        ws_out.write(row_num, 8,
                     frac_val if isinstance(frac_val, (int, float)) else 0,
                     pfmt)
        data_row_n += 1

    elif rtype == "extra":
        ws_out.write(row_num, 0, rec.get("Board Type", ""),  ext_fmt)
        ws_out.write(row_num, 1, rec.get("Board GSM", ""),   ext_fmt)
        ws_out.write(row_num, 2, rec.get("Board Size", ""),  ext_fmt)
        ws_out.write(row_num, 3, rec.get("Board Length", ""),ext_fmt)
        ws_out.write(row_num, 4, "",                          ext_fmt)
        ws_out.write(row_num, 5, rec.get("pack_style", ""),  ext_fmt)
        ws_out.write(row_num, 6, rec.get("route", ""),       ext_fmt)
        tbt_val = rec.get("tbt", 0)
        ws_out.write(row_num, 7,
                     tbt_val if tbt_val != "" else 0, ext_num)
        frac_val = rec.get("fraction", 0)
        ws_out.write(row_num, 8,
                     frac_val if isinstance(frac_val, (int, float)) else 0,
                     ext_pct)

    row_num += 1

col_widths = [22, 12, 12, 14, 16, 18, 20, 22, 12]
for ci, w in enumerate(col_widths):
    ws_out.set_column(ci, ci, w)

ws_out.freeze_panes(1, 0)
ws_out.autofilter(0, 0, row_num - 1, len(ALL_COLS) - 1)
wb_out.close()

print("PSPD Summary (enriched) written ->", PSPD_SUMMARY_FILE)
print("Total rows written               :", row_num - 1)

# ---- BOARD SUMMARY WITH PSPD AVAILABLE -------------------------------------

import pandas as pd
import xlsxwriter

# ---- LOAD FILES ------------------------------------------------------------
df_bs   = pd.read_excel(BOARD_SUMMARY_FILE,  dtype=str, engine="openpyxl").fillna("")
df_pspd = pd.read_excel(PSPD_SUMMARY_FILE,   dtype=str, engine="openpyxl").fillna("")

for col in df_bs.columns:
    df_bs[col] = df_bs[col].str.strip()
for col in df_pspd.columns:
    df_pspd[col] = df_pspd[col].str.strip()

# ---- BUILD PSPD LOOKUPS ONLY -----------------------------------------------
# LOOKUP 1: (Pack Style, Route) -> list of fractions
ps_route_fraction_pspd = {}

for _, row in df_pspd.iterrows():
    ps    = row.get("Pack Style", "").strip()
    route = row.get("Route", "").strip()
    frac  = row.get("Fraction", "").strip()
    if ps == "" or route == "" or frac == "":
        continue
    frac_val = pd.to_numeric(frac, errors="coerce")
    if pd.isna(frac_val):
        continue
    ps_route_fraction_pspd.setdefault((ps, route), []).append(frac_val)

# LOOKUP 2: (Board Type, GSM, Size, Length) -> total stock
def get_pspd_row_type(r):
    if r.get("Board GSM", "").strip() == "" and r.get("Total Stock", "").strip() == "":
        return "header"
    return "data"

pspd_total_stock = {}

for _, row in df_pspd.iterrows():
    rtype  = get_pspd_row_type(row)
    btype  = row.get("Board Type",   "").strip()
    gsm    = row.get("Board GSM",    "").strip()
    size   = row.get("Board Size",   "").strip()
    length = row.get("Board Length", "").strip()
    stock  = pd.to_numeric(row.get("Total Stock", ""), errors="coerce")
    stock  = (stock/1000) if pd.notna(stock) else 0

    if btype == "" or gsm == "" or size == "":
        continue

    if rtype == "data":
        key = (btype, gsm, size, length)
        pspd_total_stock[key] = pspd_total_stock.get(key, 0) + stock

print("PSPD fraction keys:", len(ps_route_fraction_pspd))
print("PSPD combo keys   :", len(pspd_total_stock))

# ---- DETERMINE BS ROW TYPE -------------------------------------------------
def get_bs_row_type(r):
    if r.get("Board GSM", "").strip() == "":
        return "header"
    return "data"

def get_bs_sub_type(r):
    # Identify existing sub-rows written by previous cell
    # They have blank Pack Style, Route, GSM etc. and a value in PPB RM Available
    if r.get("Pack Style", "").strip() == "" and r.get("PPB RM Available", "").strip() != "":
        plant = r.get("Pack Style", "")  # col 0 has plant label
        return "sub"
    return None

# ---- BUILD OUTPUT ROWS -----------------------------------------------------
# BS already has PPB RM Available column
# We just need to add PSPD Available alongside it
# Strategy: read BS rows as-is, identify sub-rows, add PSPD Available column

BS_COLS      = df_bs.columns.tolist()
# PSPD Available goes right after PPB RM Available
if "PPB RM Available" in BS_COLS:
    ppb_idx = BS_COLS.index("PPB RM Available")
    ALL_OUT_COLS = BS_COLS[:ppb_idx + 1] + ["PSPD Available"] + BS_COLS[ppb_idx + 1:]
else:
    ALL_OUT_COLS = BS_COLS + ["PSPD Available"]

output_rows = []

# Track current board combination for PSPD lookup
current_btype  = ""
current_gsm    = ""
current_size   = ""
current_length = ""
current_ps     = ""
current_route  = ""
current_pspd_avail_total = 0

for _, bs_row in df_bs.iterrows():
    base = {col: bs_row[col] for col in BS_COLS}
    base["PSPD Available"] = ""

    gsm_val   = bs_row.get("Board GSM", "").strip()
    pack_val  = bs_row.get("Pack Style", "").strip()
    ppb_val   = bs_row.get("PPB RM Available", "").strip()
    plant_col = bs_row.iloc[0]  # first column has plant label on sub-rows

    # Determine row type — use GSM alone as the gate (PPB may be filled on Total rows)
    if gsm_val == "":
        first_col_val = str(bs_row.iloc[0]).strip()
        ppb_num       = pd.to_numeric(ppb_val, errors="coerce")

        # Sub-rows: Total/No Stock labels, or PPB filled, or non-PackStyle label
        if first_col_val in ["Total", "No Stock"] or pd.notna(ppb_num) or (
            ppb_val != "" and first_col_val not in df_bs["Pack Style"].unique()
        ):
            # This is a sub-row — add PSPD Available on Total rows
            if first_col_val == "Total":
                base["PSPD Available"] = current_pspd_avail_total
            else:
                base["PSPD Available"] = ""  # plant rows — PSPD blank
            output_rows.append({**base, "row_type": "existing_sub", "_label": first_col_val})
        else:
            # Section header
            output_rows.append({**base, "row_type": "header"})

    else:
        # Main data row — update current combo and compute PSPD
        btype  = bs_row.get("Board Type", "").strip()
        gsm    = gsm_val
        size   = bs_row.get("Board Size", "").strip()
        length = bs_row.get("Board Length/Cyl Circumference", "").strip()
        ps     = bs_row.get("Pack Style", "").strip() or current_ps
        route  = bs_row.get("Route", "").strip() or current_route

        # Update current tracking
        if ps != "":
            current_ps    = ps
        if route != "":
            current_route = route
        current_btype  = btype  or current_btype
        current_gsm    = gsm
        current_size   = size
        current_length = length

        key_ps    = (current_ps, current_route)
        key_combo = (current_btype, current_gsm, current_size, current_length)

        fractions_pspd   = ps_route_fraction_pspd.get(key_ps, [])
        total_stock_pspd = pspd_total_stock.get(key_combo, 0)

        current_pspd_avail_total = (
            sum(total_stock_pspd * f for f in fractions_pspd)
            if fractions_pspd else 0
        )

        base["PSPD Available"] = current_pspd_avail_total  # FIX: write value onto data row
        output_rows.append({**base, "row_type": "data"})

print("Total output rows:", len(output_rows))

# ---- WRITE EXCEL -----------------------------------------------------------
wb_new = open_workbook(BOARD_SUMMARY_FILE)
ws_new = wb_new.add_worksheet("Board Summary")

COLS     = ALL_OUT_COLS
NUM_COLS = {"Program", "Finished Goods", "WIP", "Balance To Print",
            "Printing Requirement In Tonnes", "PPB RM Available", "PSPD Available"}

# Formats
hdr_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 11,
    "bg_color": "#2E4057", "font_color": "#FFFFFF",
    "align": "center", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
ps_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 11,
    "bg_color": "#1B3A5C", "font_color": "#FFFFFF",
    "valign": "vcenter", "border": 1, "border_color": "#CCCCCC"
})
rt_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#D6E4F0", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
gsm_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#EBF5FB", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
dat_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFFFFF", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
alt_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0F4F8", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
num_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFFFFF", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
num_alt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0F4F8", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
bal_pos = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#E8F5E9", "font_color": "#1B5E20",
    "valign": "vcenter", "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
bal_neg = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFEBEE", "font_color": "#B71C1C",
    "valign": "vcenter", "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
sub_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0FFF0", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
sub_num = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#F0FFF0", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
sub_tot_fmt = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#CCFFCC", "valign": "vcenter",
    "border": 1, "border_color": "#CCCCCC"
})
sub_tot_num = wb_new.add_format({
    "bold": True, "font_name": "Arial", "font_size": 10,
    "bg_color": "#CCFFCC", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})
sub_zero_fmt = wb_new.add_format({
    "font_name": "Arial", "font_size": 10,
    "bg_color": "#FFF0F0", "valign": "vcenter",
    "num_format": "#,##0.00",
    "border": 1, "border_color": "#CCCCCC"
})

# Header row
ws_new.set_row(0, 26)
for ci, col in enumerate(COLS):
    ws_new.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0
prev_pack  = None
prev_rt    = None
prev_gsm   = None

for rec in output_rows:
    ws_new.set_row(row_num, 18)
    rtype = rec["row_type"]

    if rtype == "header":
        ws_new.write(row_num, 0, rec.get("Pack Style", ""), ps_fmt)
        for ci in range(1, len(COLS)):
            ws_new.write(row_num, ci, "", ps_fmt)
        data_row_n = 0
        prev_pack = None
        prev_rt   = None
        prev_gsm  = None

    elif rtype == "data":
        pack  = rec.get("Pack Style", "")
        route = rec.get("Route", "")
        gsm   = rec.get("Board GSM", "")

        is_new_pack = (pack, route) != (prev_pack, prev_rt)
        is_new_rt   = route != prev_rt  or is_new_pack
        is_new_gsm  = gsm   != prev_gsm or is_new_rt

        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt

        for ci, col in enumerate(COLS):
            val = rec.get(col, "")
            if col == "Pack Style":
                ws_new.write(row_num, ci, pack, ps_fmt)
            elif col == "Route":
                ws_new.write(row_num, ci,
                             route if is_new_rt else "", rt_fmt)
            elif col == "Board GSM":
                ws_new.write(row_num, ci,
                             gsm if is_new_gsm else "", gsm_fmt)
            elif col == "Balance To Print":
                fval = pd.to_numeric(val, errors="coerce")
                fval = fval if pd.notna(fval) else 0
                ws_new.write(row_num, ci, fval,
                             bal_neg if fval < 0 else bal_pos)
            elif col in NUM_COLS:
                fval = pd.to_numeric(val, errors="coerce")
                ws_new.write(row_num, ci,
                             fval if pd.notna(fval) else 0, nfmt)
            else:
                ws_new.write(row_num, ci, val, rfmt)

        prev_pack  = pack
        prev_rt    = route
        prev_gsm   = gsm
        data_row_n += 1

    elif rtype == "existing_sub":
        # Preserve existing sub-row exactly, just add PSPD Available
        first_col = str(rec.get("_label", rec.get(BS_COLS[0], ""))).strip()
        is_total  = first_col == "Total"
        is_zero   = first_col == "No Stock"

        fmt     = sub_tot_fmt if is_total else (sub_zero_fmt if is_zero else sub_fmt)
        num_f   = sub_tot_num if is_total else (sub_zero_fmt if is_zero else sub_num)

        for ci, col in enumerate(COLS):
            val = rec.get(col, "")
            if col == "PSPD Available":
                pspd_val = rec.get("PSPD Available", "")
                if pspd_val != "":
                    ws_new.write(row_num, ci, float(pspd_val), num_f)
                else:
                    ws_new.write(row_num, ci, "", fmt)
            elif col in NUM_COLS:
                fval = pd.to_numeric(val, errors="coerce")
                ws_new.write(row_num, ci,
                             fval if pd.notna(fval) else 0, num_f)
            else:
                ws_new.write(row_num, ci, val, fmt)

    row_num += 1

# Column widths
col_widths = [18, 20, 12, 14, 14, 14, 15, 15, 10, 18, 22, 18, 18]
for ci, w in enumerate(col_widths):
    ws_new.set_column(ci, ci, w)

ws_new.freeze_panes(1, 0)
ws_new.autofilter(0, 0, row_num - 1, len(COLS) - 1)
wb_new.close()

print("Board Summary (with PSPD Available added) written ->", BOARD_SUMMARY_FILE)
print("Total rows written:", row_num - 1)

# ---- SYSTEM GENERATED PSPD SHEET STOCK FILE --------------------------------

PSPD_SHEET_INPUT_FILE  = rel_path("User Input", "PSPD RM", "PSPD Sheet Stocks.xlsx")
PSPD_SHEET_OUTPUT_FILE = rel_path("System Files", "System Generated PSPD RM Sheets.xlsx")
require_file(PSPD_SHEET_INPUT_FILE, "PSPD sheet stock input file")

df_pspd_sh = pd.read_excel(PSPD_SHEET_INPUT_FILE, dtype=str, engine="openpyxl").fillna("")
print("Total rows loaded :", len(df_pspd_sh))
print("Columns           :", df_pspd_sh.columns.tolist())

# Strip column names
df_pspd_sh.columns = df_pspd_sh.columns.str.strip()

for col in df_pspd_sh.columns:
    df_pspd_sh[col] = df_pspd_sh[col].str.strip()

# ---- CLASSIFY BOARD TYPE FROM PRODUCT CODE ---------------------------------
def classify_board_by_product_code(code):
    c = str(code).strip().upper()
    if "CYBER CYPAK" in c:
        return "CYCY"
    if any(k in c for k in ["BFK", "BCK", "BNK"]):
        return "BCKI"
    if any(k in c for k in ["CFK", "CFE", "CKE"]):
        return "CFKI"
    if "CYX" in c:
        return "CYXL"
    return "Others"

prod_idx = df_pspd_sh.columns.tolist().index("Product Code")
df_pspd_sh.insert(prod_idx + 1, "Board Type Category",
                  df_pspd_sh["Product Code"].apply(classify_board_by_product_code))

print("Classification distribution:")
print(df_pspd_sh["Board Type Category"].value_counts().to_string())

# ---- PARSE NUMERICS --------------------------------------------------------
# Helper to strip "cm" suffix before converting to numeric
def parse_cm(series):
    return pd.to_numeric(
        series.astype(str).str.replace("cm", "", case=False).str.strip(),
        errors="coerce"
    ).fillna(0)

df_pspd_sh["_avail_qty"] = pd.to_numeric(df_pspd_sh["Available Quantity"], errors="coerce").fillna(0)
df_pspd_sh["_gsm_raw"]   = pd.to_numeric(df_pspd_sh["GSM"],               errors="coerce").fillna(0)
df_pspd_sh["_width_raw"] = parse_cm(df_pspd_sh["Width"])
df_pspd_sh["_len_raw"]   = parse_cm(df_pspd_sh["Length"])

# ---- CALCULATE WEIGHT IN TONNES (using raw Width and Length BEFORE ×10) ----
df_pspd_sh["Gross weight"] = (
    df_pspd_sh["_avail_qty"] *
    df_pspd_sh["_gsm_raw"]   *
    df_pspd_sh["_width_raw"] *
    df_pspd_sh["_len_raw"]
) / (10_000_000 * 1000)

# ---- DERIVE BOARD GSM, BOARD SIZE, BOARD LENGTH (Width/Length ×10 AFTER weight) ---
df_pspd_sh["Board GSM"] = df_pspd_sh["_gsm_raw"].apply(
    lambda x: str(int(x)) if x != 0 else "0"
)
df_pspd_sh["Board Size"] = (df_pspd_sh["_width_raw"] * 10).apply(
    lambda x: str(int(x)) if x != 0 else "0"
)
df_pspd_sh["Board Length"] = (df_pspd_sh["_len_raw"] * 10).apply(
    lambda x: "0" if x == 0 else str(int(x))
)

# Drop temp columns
df_pspd_sh.drop(columns=["_avail_qty", "_gsm_raw", "_width_raw", "_len_raw"], inplace=True)

# ---- INSERT DERIVED COLUMNS AFTER Board Type Category IN CORRECT ORDER -----
# Insert in REVERSE order so final order is:
# Board Type Category | Board GSM | Board Size | Board Length | Gross weight
for derived_col in ["Gross weight", "Board Length", "Board Size", "Board GSM"]:
    if derived_col in df_pspd_sh.columns:
        col_data = df_pspd_sh.pop(derived_col)
        btc_idx  = df_pspd_sh.columns.tolist().index("Board Type Category")
        df_pspd_sh.insert(btc_idx + 1, derived_col, col_data)

print("Sample derived values:")
print(df_pspd_sh[["Product Code", "Board Type Category",
               "Board GSM", "Board Size", "Board Length",
               "Gross weight"]].head(10).to_string(index=False))

# ---- WRITE OUTPUT FILE -----------------------------------------------------
wb_pspd_sh = open_workbook(PSPD_SHEET_OUTPUT_FILE)
ws_pspd_sh = wb_pspd_sh.add_worksheet("PSPD Sheet Stock")

cols     = df_pspd_sh.columns.tolist()
NEW_COLS = {"Board GSM", "Board Size", "Board Length", "Board Type Category", "Gross weight"}

hdr_fmt = wb_pspd_sh.add_format({"bold":True,"font_name":"Arial","font_size":11,
                                  "bg_color":"#2E4057","font_color":"#FFFFFF",
                                  "align":"center","valign":"vcenter",
                                  "border":1,"border_color":"#CCCCCC"})
new_fmt = wb_pspd_sh.add_format({"bold":True,"font_name":"Arial","font_size":11,
                                  "bg_color":"#E6A817","font_color":"#FFFFFF",
                                  "align":"center","valign":"vcenter",
                                  "border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_pspd_sh.add_format({"font_name":"Arial","font_size":10,
                                  "bg_color":"#FFFFFF","valign":"vcenter",
                                  "border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_pspd_sh.add_format({"font_name":"Arial","font_size":10,
                                  "bg_color":"#F0F4F8","valign":"vcenter",
                                  "border":1,"border_color":"#CCCCCC"})
cat_fmt = wb_pspd_sh.add_format({"bold":True,"font_name":"Arial","font_size":10,
                                  "bg_color":"#EAF3DE","font_color":"#27500A",
                                  "valign":"vcenter","border":1,"border_color":"#CCCCCC"})
cat_alt = wb_pspd_sh.add_format({"bold":True,"font_name":"Arial","font_size":10,
                                  "bg_color":"#D4E8C2","font_color":"#27500A",
                                  "valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt = wb_pspd_sh.add_format({"font_name":"Arial","font_size":10,
                                  "bg_color":"#FFFFFF","valign":"vcenter",
                                  "align":"center","num_format":"#,##0.00",
                                  "border":1,"border_color":"#CCCCCC"})
num_alt = wb_pspd_sh.add_format({"font_name":"Arial","font_size":10,
                                  "bg_color":"#F0F4F8","valign":"vcenter",
                                  "align":"center","num_format":"#,##0.00",
                                  "border":1,"border_color":"#CCCCCC"})

ws_pspd_sh.set_row(0, 24)
for ci, col in enumerate(cols):
    ws_pspd_sh.write(0, ci, col, new_fmt if col in NEW_COLS else hdr_fmt)

for ri, (_, row_data) in enumerate(df_pspd_sh.iterrows(), start=1):
    ws_pspd_sh.set_row(ri, 18)
    rfmt = alt_fmt if ri % 2 == 0 else dat_fmt
    nfmt = num_alt if ri % 2 == 0 else num_fmt
    cf   = cat_alt if ri % 2 == 0 else cat_fmt
    for ci, col in enumerate(cols):
        val = row_data[col]
        if col == "Board Type Category":
            ws_pspd_sh.write(ri, ci, val, cf)
        elif col in {"Board GSM", "Board Size", "Board Length", "Gross weight"}:
            fval = pd.to_numeric(val, errors="coerce")
            ws_pspd_sh.write(ri, ci, fval if pd.notna(fval) else 0, nfmt)
        else:
            ws_pspd_sh.write(ri, ci, val, rfmt)

for ci, col in enumerate(cols):
    width = max(len(col) + 4, 14)
    if col in {"Board GSM", "Board Size", "Board Length"}: width = 13
    if col == "Board Type Category": width = 22
    if col == "Product Code": width = 40
    ws_pspd_sh.set_column(ci, ci, width)

ws_pspd_sh.freeze_panes(1, 0)
ws_pspd_sh.autofilter(0, 0, len(df_pspd_sh), len(cols) - 1)
wb_pspd_sh.close()

print("Sheet Stock output written ->", PSPD_SHEET_OUTPUT_FILE)
print("Rows written               :", len(df_pspd_sh))

# ---- PSPD SHEET RM SYSTEM SUMMARY ------------------------------------------

PSPD_SHEET_SUMMARY_FILE = rel_path("System Files", "PSPD Sheet RM System Summary.xlsx")

df_pspd_sh_out = pd.read_excel(PSPD_SHEET_OUTPUT_FILE, dtype=str, engine="openpyxl").fillna("")

for col in df_pspd_sh_out.columns:
    df_pspd_sh_out[col] = df_pspd_sh_out[col].str.strip()

df_pspd_sh_out["_stock_num"] = pd.to_numeric(
    df_pspd_sh_out["Gross weight"], errors="coerce"
).fillna(0)
df_pspd_sh_out["_gsm_n"]  = pd.to_numeric(df_pspd_sh_out["Board GSM"],    errors="coerce").fillna(0)
df_pspd_sh_out["_size_n"] = pd.to_numeric(df_pspd_sh_out["Board Size"],   errors="coerce").fillna(0)
df_pspd_sh_out["_len_n"]  = pd.to_numeric(df_pspd_sh_out["Board Length"], errors="coerce").fillna(0)

df_pspd_sh_out = df_pspd_sh_out.sort_values(
    ["Board Type Category", "_gsm_n", "_size_n", "_len_n"]
).reset_index(drop=True)

output_rows = []

for board_type, df_bt in df_pspd_sh_out.groupby("Board Type Category", sort=False):
    output_rows.append({
        "row_type": "header", "board_type": board_type,
        "gsm": "", "size": "", "length": "", "stock": ""
    })
    for (gsm, size, length), df_grp in df_bt.groupby(
        ["Board GSM", "Board Size", "Board Length"], sort=False
    ):
        total_stock = df_grp["_stock_num"].sum()
        output_rows.append({
            "row_type": "data", "board_type": board_type,
            "gsm": gsm, "size": size, "length": length, "stock": total_stock
        })

print("Total output rows:", len(output_rows))

wb_sh_sum = open_workbook(PSPD_SHEET_SUMMARY_FILE)
ws_sh_sum = wb_sh_sum.add_worksheet("PSPD Sheet Summary")

COLS = ["Board Type", "Board GSM", "Board Size", "Board Length", "Total Stock"]

hdr_fmt = wb_sh_sum.add_format({"bold":True,"font_name":"Arial","font_size":11,
    "bg_color":"#2E4057","font_color":"#FFFFFF","align":"center","valign":"vcenter",
    "border":1,"border_color":"#CCCCCC"})
sec_fmt = wb_sh_sum.add_format({"bold":True,"font_name":"Arial","font_size":11,
    "bg_color":"#E6A817","font_color":"#FFFFFF","align":"left","valign":"vcenter",
    "border":1,"border_color":"#CCCCCC"})
sec_blk = wb_sh_sum.add_format({"bg_color":"#E6A817",
    "border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_sh_sum.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_sh_sum.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt = wb_sh_sum.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","align":"center","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
num_alt = wb_sh_sum.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","align":"center","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})

ws_sh_sum.set_row(0, 24)
for ci, col in enumerate(COLS):
    ws_sh_sum.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0

for rec in output_rows:
    ws_sh_sum.set_row(row_num, 18)
    if rec["row_type"] == "header":
        ws_sh_sum.write(row_num, 0, rec["board_type"], sec_fmt)
        for ci in range(1, len(COLS)):
            ws_sh_sum.write(row_num, ci, "", sec_blk)
        data_row_n = 0
    elif rec["row_type"] == "data":
        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt
        ws_sh_sum.write(row_num, 0, rec["board_type"], rfmt)
        ws_sh_sum.write(row_num, 1, rec["gsm"],         rfmt)
        ws_sh_sum.write(row_num, 2, rec["size"],         rfmt)
        ws_sh_sum.write(row_num, 3, rec["length"],       rfmt)
        ws_sh_sum.write(row_num, 4, rec["stock"],        nfmt)
        data_row_n += 1
    row_num += 1

col_widths = [22, 12, 12, 14, 16]
for ci, w in enumerate(col_widths):
    ws_sh_sum.set_column(ci, ci, w)

ws_sh_sum.freeze_panes(1, 0)
ws_sh_sum.autofilter(0, 0, row_num - 1, len(COLS) - 1)
wb_sh_sum.close()

print("PSPD Sheet Summary written ->", PSPD_SHEET_SUMMARY_FILE)
print("Total rows written          :", row_num - 1)

# ---- PSPD SHEET SUMMARY WITH BOARD SUMMARY LOOKUP --------------------------

import pandas as pd
import xlsxwriter

df_pspd_sh_sum = pd.read_excel(PSPD_SHEET_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")
df_bs_raw      = pd.read_excel(BOARD_SUMMARY_FILE,       dtype=str, engine="openpyxl").fillna("")

for col in df_pspd_sh_sum.columns:
    df_pspd_sh_sum[col] = df_pspd_sh_sum[col].str.strip()
for col in df_bs_raw.columns:
    df_bs_raw[col] = df_bs_raw[col].str.strip()

# ---- BOARD SUMMARY LOOKUP (same logic as cell 15) --------------------------
bs_lookup_sh = {}
for _, row in df_bs_raw.iterrows():
    btype  = row.get("Board Type", "").strip()
    gsm    = row.get("Board GSM", "").strip()
    size   = row.get("Board Size", "").strip()
    length = row.get("Board Length/Cyl Circumference", "").strip()
    ps     = row.get("Pack Style", "").strip()
    route  = row.get("Route", "").strip()
    tbt    = pd.to_numeric(row.get("Printing Requirement In Tonnes", ""), errors="coerce")
    tbt    = tbt if pd.notna(tbt) else 0
    if gsm == "":
        continue
    key = (btype, gsm, size, length)
    bs_lookup_sh.setdefault(key, []).append({"pack_style": ps, "route": route, "tbt": tbt})

def attach_fractions_sh(lookup):
    result = {}
    for key, matches in lookup.items():
        total_tbt = sum(m["tbt"] for m in matches)
        enriched  = []
        for m in matches:
            frac = (m["tbt"] / total_tbt) if total_tbt != 0 else 0
            enriched.append({**m, "fraction": frac})
        result[key] = enriched
    return result

bs_lookup_sh = attach_fractions_sh(bs_lookup_sh)
print("Board Summary lookup keys:", len(bs_lookup_sh))

# ---- BUILD OUTPUT ROWS -----------------------------------------------------
def get_pspd_sh_row_type(r):
    if r.get("Board GSM", "").strip() == "" and r.get("Total Stock", "").strip() == "":
        return "header"
    return "data"

PSPD_SH_COLS = df_pspd_sh_sum.columns.tolist()
output_rows  = []

for _, row in df_pspd_sh_sum.iterrows():
    rtype = get_pspd_sh_row_type(row)
    base  = {col: row[col] for col in PSPD_SH_COLS}

    if rtype == "header":
        output_rows.append({**base, "row_type": "header",
                             "pack_style": "", "route": "", "tbt": "", "fraction": ""})
        continue

    btype  = row.get("Board Type",   "").strip()
    gsm    = row.get("Board GSM",    "").strip()
    size   = row.get("Board Size",   "").strip()
    length = row.get("Board Length", "").strip()

    key     = (btype, gsm, size, length)
    matches = bs_lookup_sh.get(key, [])

    if not matches:
        output_rows.append({**base, "row_type": "data",
                             "pack_style": "", "route": "", "tbt": "", "fraction": 0})
    else:
        first = matches[0]
        output_rows.append({**base, "row_type": "data",
                             "pack_style": first["pack_style"], "route": first["route"],
                             "tbt": first["tbt"], "fraction": first["fraction"]})
        for m in matches[1:]:
            extra = {col: "" for col in PSPD_SH_COLS}
            extra["Board Type"]   = btype
            extra["Board GSM"]    = gsm
            extra["Board Size"]   = size
            extra["Board Length"] = length
            extra["Total Stock"]  = ""
            output_rows.append({**extra, "row_type": "extra",
                                 "pack_style": m["pack_style"], "route": m["route"],
                                 "tbt": m["tbt"], "fraction": m["fraction"]})

print("Total output rows:", len(output_rows))

# ---- WRITE ENRICHED SHEET SUMMARY -----------------------------------------
wb_sh_out = open_workbook(PSPD_SHEET_SUMMARY_FILE)
ws_sh_out = wb_sh_out.add_worksheet("PSPD Sheet Summary")

ALL_COLS = ["Board Type", "Board GSM", "Board Size", "Board Length",
            "Total Stock", "Pack Style", "Route", "Printing Requirement In Tonnes", "Fraction"]

hdr_fmt = wb_sh_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
    "bg_color":"#2E4057","font_color":"#FFFFFF","align":"center","valign":"vcenter",
    "border":1,"border_color":"#CCCCCC"})
sec_fmt = wb_sh_out.add_format({"bold":True,"font_name":"Arial","font_size":11,
    "bg_color":"#E6A817","font_color":"#FFFFFF","align":"left","valign":"vcenter",
    "border":1,"border_color":"#CCCCCC"})
sec_blk = wb_sh_out.add_format({"bg_color":"#E6A817","border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","align":"center","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
num_alt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","align":"center","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
pct_fmt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","align":"center","num_format":"0.00%",
    "border":1,"border_color":"#CCCCCC"})
pct_alt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","align":"center","num_format":"0.00%",
    "border":1,"border_color":"#CCCCCC"})
ext_fmt = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFF9E6","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
ext_num = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFF9E6","valign":"vcenter","align":"center","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
ext_pct = wb_sh_out.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFF9E6","valign":"vcenter","align":"center","num_format":"0.00%",
    "border":1,"border_color":"#CCCCCC"})

ws_sh_out.set_row(0, 24)
for ci, col in enumerate(ALL_COLS):
    ws_sh_out.write(0, ci, col, hdr_fmt)

row_num = 1
data_row_n = 0

for rec in output_rows:
    ws_sh_out.set_row(row_num, 18)
    rtype = rec["row_type"]

    if rtype == "header":
        ws_sh_out.write(row_num, 0, rec.get("Board Type", ""), sec_fmt)
        for ci in range(1, len(ALL_COLS)):
            ws_sh_out.write(row_num, ci, "", sec_blk)
        data_row_n = 0

    elif rtype == "data":
        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt
        pfmt = pct_alt if data_row_n % 2 == 0 else pct_fmt
        ws_sh_out.write(row_num, 0, rec.get("Board Type", ""),  rfmt)
        ws_sh_out.write(row_num, 1, rec.get("Board GSM", ""),   rfmt)
        ws_sh_out.write(row_num, 2, rec.get("Board Size", ""),  rfmt)
        ws_sh_out.write(row_num, 3, rec.get("Board Length", ""),rfmt)
        stock_val = pd.to_numeric(rec.get("Total Stock", ""), errors="coerce")
        ws_sh_out.write(row_num, 4, stock_val if pd.notna(stock_val) else 0, nfmt)
        ws_sh_out.write(row_num, 5, rec.get("pack_style", ""), rfmt)
        ws_sh_out.write(row_num, 6, rec.get("route", ""),      rfmt)
        tbt_val = rec.get("tbt", "")
        ws_sh_out.write(row_num, 7, tbt_val if tbt_val != "" else 0, nfmt)
        frac_val = rec.get("fraction", 0)
        ws_sh_out.write(row_num, 8,
                        frac_val if isinstance(frac_val, (int, float)) else 0, pfmt)
        data_row_n += 1

    elif rtype == "extra":
        ws_sh_out.write(row_num, 0, rec.get("Board Type", ""),  ext_fmt)
        ws_sh_out.write(row_num, 1, rec.get("Board GSM", ""),   ext_fmt)
        ws_sh_out.write(row_num, 2, rec.get("Board Size", ""),  ext_fmt)
        ws_sh_out.write(row_num, 3, rec.get("Board Length", ""),ext_fmt)
        ws_sh_out.write(row_num, 4, "",                          ext_fmt)
        ws_sh_out.write(row_num, 5, rec.get("pack_style", ""), ext_fmt)
        ws_sh_out.write(row_num, 6, rec.get("route", ""),      ext_fmt)
        tbt_val = rec.get("tbt", 0)
        ws_sh_out.write(row_num, 7, tbt_val if tbt_val != "" else 0, ext_num)
        frac_val = rec.get("fraction", 0)
        ws_sh_out.write(row_num, 8,
                        frac_val if isinstance(frac_val, (int, float)) else 0, ext_pct)

    row_num += 1

col_widths = [22, 12, 12, 14, 16, 18, 20, 22, 12]
for ci, w in enumerate(col_widths):
    ws_sh_out.set_column(ci, ci, w)

ws_sh_out.freeze_panes(1, 0)
ws_sh_out.autofilter(0, 0, row_num - 1, len(ALL_COLS) - 1)
wb_sh_out.close()

print("PSPD Sheet Summary (enriched) written ->", PSPD_SHEET_SUMMARY_FILE)
print("Total rows written                     :", row_num - 1)

# ---- NOW COMBINE: add Sheet stock into PSPD Available in Board Summary ------
# Build Sheet lookups (same pattern as Cell 16 does for Reels)
df_pspd_sh_enriched = pd.read_excel(PSPD_SHEET_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")
for col in df_pspd_sh_enriched.columns:
    df_pspd_sh_enriched[col] = df_pspd_sh_enriched[col].str.strip()

ps_route_fraction_sh = {}
for _, row in df_pspd_sh_enriched.iterrows():
    ps    = row.get("Pack Style", "").strip()
    route = row.get("Route", "").strip()
    frac  = row.get("Fraction", "").strip()
    if ps == "" or route == "" or frac == "":
        continue
    frac_val = pd.to_numeric(frac, errors="coerce")
    if pd.isna(frac_val):
        continue
    ps_route_fraction_sh.setdefault((ps, route), []).append(frac_val)

def get_sh_row_type(r):
    if r.get("Board GSM", "").strip() == "" and r.get("Total Stock", "").strip() == "":
        return "header"
    return "data"

pspd_sh_total_stock = {}
for _, row in df_pspd_sh_enriched.iterrows():
    if get_sh_row_type(row) != "data":
        continue
    btype  = row.get("Board Type",   "").strip()
    gsm    = row.get("Board GSM",    "").strip()
    size   = row.get("Board Size",   "").strip()
    length = row.get("Board Length", "").strip()
    stock  = pd.to_numeric(row.get("Total Stock", ""), errors="coerce")
    stock  = stock if pd.notna(stock) else 0
    if btype == "" or gsm == "" or size == "":
        continue
    key = (btype, gsm, size, length)
    pspd_sh_total_stock[key] = pspd_sh_total_stock.get(key, 0) + stock

print("Sheet fraction keys:", len(ps_route_fraction_sh))
print("Sheet combo keys   :", len(pspd_sh_total_stock))

# ---- READ CURRENT BOARD SUMMARY (already has Reels PSPD Available) ---------
df_bs_current = pd.read_excel(BOARD_SUMMARY_FILE, dtype=str, engine="openpyxl").fillna("")
for col in df_bs_current.columns:
    df_bs_current[col] = df_bs_current[col].str.strip()

BS_COLS_CURR = df_bs_current.columns.tolist()

current_ps    = ""
current_route = ""
current_btype = ""
current_gsm   = ""

updated_rows = []

for _, bs_row in df_bs_current.iterrows():
    base    = {col: bs_row[col] for col in BS_COLS_CURR}
    gsm_val  = bs_row.get("Board GSM", "").strip()
    size_val = bs_row.get("Board Size", "").strip()

    if gsm_val == "" and size_val == "":
        base["_prog_in_tonnes"] = ""
        base["_pspd_total"]     = ""
        updated_rows.append(base)
        continue

    # Main data row — compute sheet contribution and add to existing PSPD Available
    btype  = bs_row.get("Board Type", "").strip()  or current_btype
    ps     = bs_row.get("Pack Style", "").strip()  or current_ps
    route  = bs_row.get("Route", "").strip()        or current_route
    gsm_val = bs_row.get("Board GSM", "").strip()  or current_gsm

    current_btype = btype
    current_ps    = ps
    current_route = route
    current_gsm   = gsm_val

    size   = bs_row.get("Board Size", "").strip()
    length = bs_row.get("Board Length/Cyl Circumference", "").strip()

    base["Board Type"] = btype
    base["Pack Style"] = ps
    base["Route"]      = route
    base["Board GSM"]  = gsm_val

    key_ps    = (current_ps, current_route)
    key_combo = (current_btype, size, length) if gsm_val == "" else (current_btype, gsm_val, size, length)

    fracs_sh  = ps_route_fraction_sh.get(key_ps, [])
    stk_sh    = pspd_sh_total_stock.get(key_combo, 0)
    sh_avail  = sum(stk_sh * f for f in fracs_sh) if fracs_sh else 0

    existing_pspd = pd.to_numeric(bs_row.get("PSPD Available", ""), errors="coerce")
    existing_pspd = existing_pspd if pd.notna(existing_pspd) else 0

    base["PSPD Available"] = existing_pspd + sh_avail
     # Pre-compute Program In Tonnes and RM Cover Days inline
    prog_val = pd.to_numeric(bs_row.get("Program", ""), errors="coerce")
    prog_val = prog_val if pd.notna(prog_val) else 0
    btp_val  = pd.to_numeric(bs_row.get("Balance To Print", ""), errors="coerce")
    btp_val  = btp_val if pd.notna(btp_val) else 0
    tbt_val  = pd.to_numeric(bs_row.get("Printing Requirement In Tonnes", ""), errors="coerce")
    tbt_val  = tbt_val if pd.notna(tbt_val) else 0

    prog_in_tonnes = (((prog_val / btp_val) * tbt_val)/30) if btp_val != 0 else 0

    # PPB Total comes from existing Board Summary Total sub-row
    # PSPD Total = existing_pspd + sh_avail (just computed above)
    # We store prog_in_tonnes now; PPB total will be picked up from next Total sub-row
    base["_prog_in_tonnes"] = prog_in_tonnes
    base["_pspd_total"]     = existing_pspd + sh_avail
    updated_rows.append(base)

# ---- BACK-FILL PPB Total into each data row --------------------------------
# Walk updated_rows, find Total sub-rows, read their PPB RM Available,
# assign back to the preceding data row

last_data_idx = None
for i, rec in enumerate(updated_rows):
    gsm_val   = str(rec.get("Board GSM", "")).strip()
    first_col = str(list(rec.values())[0]).strip()

    if gsm_val != "":
        # Main data row
        last_data_idx = i
        updated_rows[i]["_ppb_total"] = 0  # default until Total sub-row found

    elif first_col == "Total" and last_data_idx is not None:
        # Total sub-row — read PPB RM Available and assign to preceding data row
        ppb_total = pd.to_numeric(rec.get("PPB RM Available", ""), errors="coerce")
        ppb_total = ppb_total if pd.notna(ppb_total) else 0
        updated_rows[last_data_idx]["_ppb_total"] = ppb_total

    else:
        updated_rows[i]["_ppb_total"] = ""

# Now compute RM Cover Days for each data row
for i, rec in enumerate(updated_rows):
    gsm_val = str(rec.get("Board GSM", "")).strip()
    if gsm_val == "":
        updated_rows[i]["_rm_cover_days"] = ""
        continue
    pit   = rec.get("_prog_in_tonnes", 0)
    pit   = pit if isinstance(pit, (int, float)) else 0
    ppbt  = rec.get("_ppb_total", 0)
    ppbt  = ppbt if isinstance(ppbt, (int, float)) else 0
    pspdt = rec.get("_pspd_total", 0)
    pspdt = pspdt if isinstance(pspdt, (int, float)) else 0
    denom = ppbt + pspdt
    updated_rows[i]["_rm_cover_days"] = (denom / pit) if pit != 0 else 0

# ---- REWRITE BOARD SUMMARY WITH COMBINED PSPD AVAILABLE --------------------
wb_combined = open_workbook(BOARD_SUMMARY_FILE)
ws_combined = wb_combined.add_worksheet("Board Summary")

if "RM Cover Days" not in BS_COLS_CURR:
    BS_COLS_CURR = BS_COLS_CURR + ["RM Cover Days"]
COLS = BS_COLS_CURR

NUM_COLS = {"Program", "Finished Goods", "WIP", "Balance To Print",
            "Printing Requirement In Tonnes", "PPB RM Available", "PSPD Available","RM Cover Days"}

hdr_fmt = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":11,
    "bg_color":"#2E4057","font_color":"#FFFFFF","align":"center","valign":"vcenter",
    "border":1,"border_color":"#CCCCCC"})
ps_fmt = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":11,
    "bg_color":"#1B3A5C","font_color":"#FFFFFF","valign":"vcenter",
    "border":1,"border_color":"#CCCCCC"})
rt_fmt = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":10,
    "bg_color":"#D6E4F0","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
gsm_fmt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#EBF5FB","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
dat_fmt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
alt_fmt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
num_fmt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFFFFF","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
num_alt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0F4F8","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
bal_pos = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":10,
    "bg_color":"#E8F5E9","font_color":"#1B5E20","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
bal_neg = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":10,
    "bg_color":"#FFEBEE","font_color":"#B71C1C","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
sub_fmt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0FFF0","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
sub_num = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#F0FFF0","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
sub_tot_fmt = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":10,
    "bg_color":"#CCFFCC","valign":"vcenter","border":1,"border_color":"#CCCCCC"})
sub_tot_num = wb_combined.add_format({"bold":True,"font_name":"Arial","font_size":10,
    "bg_color":"#CCFFCC","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})
sub_zero_fmt = wb_combined.add_format({"font_name":"Arial","font_size":10,
    "bg_color":"#FFF0F0","valign":"vcenter","num_format":"#,##0.00",
    "border":1,"border_color":"#CCCCCC"})

ws_combined.set_row(0, 26)
for ci, col in enumerate(COLS):
    ws_combined.write(0, ci, col, hdr_fmt)

row_num    = 1
data_row_n = 0
prev_pack  = None
prev_rt    = None
prev_gsm   = None
last_pspd_avail = 0.0

for rec in updated_rows:
    ws_combined.set_row(row_num, 18)
    gsm_val = str(rec.get("Board GSM", "")).strip()

    if gsm_val == "":
        first_col = str(list(rec.values())[0]).strip()
        is_total  = first_col == "Total"
        is_zero   = first_col == "No Stock"
        pack_val  = str(rec.get("Pack Style", "")).strip()
        ppb_check     = pd.to_numeric(rec.get("PPB RM Available", ""), errors="coerce")
        known_packs   = set(r.get("Pack Style", "") for r in updated_rows if str(r.get("Board GSM","")).strip() != "")
        is_sub        = is_total or is_zero or pd.notna(ppb_check) or first_col not in known_packs

        if is_sub:
            fmt   = sub_tot_fmt if is_total else (sub_zero_fmt if is_zero else sub_fmt)
            num_f = sub_tot_num if is_total else (sub_zero_fmt if is_zero else sub_num)
            for ci, col in enumerate(COLS):
                val = rec.get(col, "")
                if col == "PSPD Available":
                    if is_total:
                        # Always write last_pspd_avail on Total row
                        ws_combined.write(row_num, ci, last_pspd_avail, num_f)
                    else:
                        ws_combined.write(row_num, ci, "", num_f)
                elif col == "PPB RM Available":
                    fval = pd.to_numeric(val, errors="coerce")
                    ws_combined.write(row_num, ci, fval if pd.notna(fval) else 0, num_f)
                elif ci == 0:
                    ws_combined.write(row_num, ci, first_col, fmt)
                else:
                    ws_combined.write(row_num, ci, "", fmt)
        else:
            ws_combined.write(row_num, 0, pack_val, ps_fmt)
            for ci in range(1, len(COLS)):
                ws_combined.write(row_num, ci, "", ps_fmt)
            data_row_n = 0
            prev_pack = None
            prev_rt   = None
            prev_gsm  = None

    else:
        pack  = rec.get("Pack Style", "") or prev_pack or ""
        route = rec.get("Route", "")      or prev_rt   or ""
        gsm   = gsm_val

        is_new_pack = (pack, route) != (prev_pack, prev_rt)
        is_new_rt   = route != prev_rt  or is_new_pack
        is_new_gsm  = gsm   != prev_gsm or is_new_rt

        rfmt = alt_fmt if data_row_n % 2 == 0 else dat_fmt
        nfmt = num_alt if data_row_n % 2 == 0 else num_fmt

        for ci, col in enumerate(COLS):
            val = rec.get(col, "")
            if col == "Pack Style":
                ws_combined.write(row_num, ci, pack, ps_fmt)
            elif col == "Route":
                ws_combined.write(row_num, ci, route if is_new_rt else "", rt_fmt)
            elif col == "Board GSM":
                ws_combined.write(row_num, ci, gsm if is_new_gsm else "", gsm_fmt)
            elif col == "Balance To Print":
                fval = pd.to_numeric(val, errors="coerce")
                fval = fval if pd.notna(fval) else 0
                ws_combined.write(row_num, ci, fval, bal_neg if fval < 0 else bal_pos)
            elif col == "PSPD Available":
                fval = pd.to_numeric(val, errors="coerce")
                fval = fval if pd.notna(fval) else 0
                last_pspd_avail = fval  # store for Total sub-row
                ws_combined.write(row_num, ci, fval, nfmt)
            elif col == "RM Cover Days":
                rm_val = rec.get("_rm_cover_days", 0)
                rm_val = rm_val if isinstance(rm_val, (int, float)) else 0
                ws_combined.write(row_num, ci, rm_val, nfmt)
            elif col in NUM_COLS:
                fval = pd.to_numeric(val, errors="coerce")
                ws_combined.write(row_num, ci, fval if pd.notna(fval) else 0, nfmt)
            else:
                ws_combined.write(row_num, ci, val, rfmt)

        prev_pack  = pack
        prev_rt    = route
        prev_gsm   = gsm
        data_row_n += 1

    row_num += 1
    

col_widths = [18, 20, 12, 14, 14, 14, 15, 15, 10, 18, 22, 18, 18, 18]
for ci, w in enumerate(col_widths):
    ws_combined.set_column(ci, ci, w)

ws_combined.freeze_panes(1, 0)
ws_combined.autofilter(0, 0, row_num - 1, len(COLS) - 1)
wb_combined.close()

print("Board Summary (Reels + Sheets combined) written ->", BOARD_SUMMARY_FILE)
print("Total rows written:", row_num - 1)
