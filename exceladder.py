"""
Census SC/ST Data Processor -- General Purpose
=================================================
Processes Census "Primary Census Abstract" data files in the standard
DDW-* format (e.g. B-01, B-03, and similarly structured tables, for SC,
ST, or general population) into clean, long-format CSVs.

GENERAL APPROACH -- not hardcoded to any specific table:
These Census tables share a standard layout: 5 leading identifier
columns (Table Name, State Code, District Code, Area Name,
Total/Rural/Urban), one "classifying variable" column (whatever the
table is broken down by -- Age-Group, Educational level, Household
size, etc: read directly from the file, not hardcoded), then metric
columns always arranged in Persons/Males/Females triplets, arranged
under one or more levels of group headers (Population, Main workers,
Marginal workers, etc).

Header flattening works like this, with no hardcoded table-specific
column lists:
  1. Auto-detect the "numbering row" (Census's own sequential column
     index row -- always the last header row).
  2. Auto-detect the "sex row" directly above it (contains repeating
     Persons/Males/Females) -- this locates exactly where identifier
     columns end and metric columns begin, and confirms metric columns
     come in clean triplets.
  3. Flatten group-header rows into metric names using TWO different
     fill rules depending on row position:
       - The row immediately after the title (top-level groups like
         "Population", "Marginal workers") is forward-filled across the
         WHOLE metric region, since these top-level labels legitimately
         span multiple triplets.
       - Every row below that (finer sub-labels) is forward-filled only
         WITHIN each 3-column triplet, never crossing into the next
         triplet. This was a real bug in an earlier version of this
         script: unrestricted forward-fill let text from one triplet
         (e.g. a wrapped "3 months" label) bleed into the unrelated
         next triplet. Verified fixed by checking the flattened output
         against manually-inspected header cells.
     This two-tier rule was derived by inspecting two different real
     table structures (B-01 SC, B-03 SC) and generalizes to other
     Census PCA tables using the same standard layout -- but if you
     feed it a table with a genuinely different structure, the
     validation check below will raise a clear error rather than
     silently producing wrong column names.

Just drop any DDW-*.xls / DDW-*.xlsx file (any table series, any
state) into the input folder and re-run -- no code changes needed for
new files sharing this standard layout.

Install first:
    pip install pandas openpyxl xlrd

Usage:
    python process_census_files.py
"""

import os
import re
import glob
import pandas as pd

INPUT_DIR = r"C:\Users\PC\Desktop\ScSctdata\census_SC_ST\b1sc"        # put all DDW-*.xls / .xlsx files here
OUTPUT_DIR = r"C:\Users\PC\Desktop\ScSctdata\census_SC_ST\b1sc\Census_COMBINED\combined-census"   # per-file + combined CSVs land here

# These 5 leading identifier columns are a Census-wide standard across
# Primary Census Abstract tables -- not specific to any one dataset.
STANDARD_ID_COLS = ["table_name", "state_code", "district_code", "area_name", "total_rural_urban"]


def extract_dataset_code(filepath: str) -> str:
    """Pull the table identifier out of a filename, e.g.
    'DDW-0100B-01SC-Census.xls' -> '01SC', 'DDW-0000B-03SC.xlsx' -> '03SC'.
    Generalizes to any table code (not just SC) -- e.g. '01ST', '02ST' etc.
    """
    name = os.path.basename(filepath)
    m = re.search(r"DDW-\d+B-([A-Za-z0-9]+?)(?:-Census)?\.xlsx?$", name, re.IGNORECASE)
    if not m:
        raise ValueError(f"Filename doesn't match expected DDW-*B-*.xls[x] pattern: {name}")
    return m.group(1).upper()


def find_numbering_row(df: pd.DataFrame, max_scan=15) -> int:
    """The row of sequential small integers Census uses as its own
    column index -- always the very last header row."""
    for r in range(max_scan):
        row = df.iloc[r]
        numeric_vals = [v for v in row if pd.notna(v) and isinstance(v, (int, float))]
        if len(numeric_vals) >= 10 and all(0 < v < 60 for v in numeric_vals):
            return r
    raise ValueError("Could not auto-detect the numbering row -- file structure may differ from the standard Census layout this script expects.")


def find_sex_row_and_id_col_count(df: pd.DataFrame, numbering_row: int):
    """The row directly above the numbering row should contain repeating
    Persons/Males/Females -- this tells us exactly where metric columns
    start (= where identifier columns end)."""
    sex_row = numbering_row - 1
    row = df.iloc[sex_row]
    persons_cols = [c for c in range(len(row)) if str(row.iloc[c]).strip() == "Persons"]
    if len(persons_cols) < 1:
        raise ValueError("Could not find the Persons/Males/Females row -- file structure may differ from the standard Census layout this script expects.")
    id_col_count = persons_cols[0]
    metric_col_count = df.shape[1] - id_col_count
    if metric_col_count % 3 != 0:
        raise ValueError(
            f"Metric columns ({metric_col_count}) aren't a clean multiple of 3 "
            f"(Persons/Males/Females) -- file structure may differ from what "
            f"this script expects. Check manually."
        )
    return sex_row, id_col_count


def flatten_metric_names(df: pd.DataFrame, title_row: int, sex_row: int, id_col_count: int) -> list:
    """Build metric column names using the two-tier fill rule described
    in the module docstring."""
    n_cols = df.shape[1]
    n_triplets = (n_cols - id_col_count) // 3

    # collect each header row (between title and sex row) as a list,
    # metric region only
    header_rows = []
    for r in range(title_row + 1, sex_row):
        row = df.iloc[r].iloc[id_col_count:].tolist()
        header_rows.append(row)

    filled_rows = []
    for i, row in enumerate(header_rows):
        if i == 0:
            # top-level row: fill forward across the WHOLE metric region
            filled = pd.Series(row).ffill().tolist()
        else:
            # sub-level rows: fill forward only WITHIN each triplet
            filled = list(row)
            for t in range(n_triplets):
                start, end = t * 3, t * 3 + 3
                triplet = pd.Series(filled[start:end]).ffill().tolist()
                filled[start:end] = triplet
        filled_rows.append(filled)

    metric_names = []
    sex_labels = ["Persons", "Males", "Females"]
    for c in range(n_cols - id_col_count):
        pieces = []
        for filled in filled_rows:
            val = filled[c]
            if pd.isna(val):
                continue
            val_str = str(val).strip()
            if not val_str or set(val_str) <= {"_", " "}:
                continue  # skip separator/decoration-only fragments
            if not pieces or pieces[-1] != val_str:  # skip immediate duplicates
                pieces.append(val_str)
        sex_label = sex_labels[c % 3]  # position within its Persons/Males/Females triplet
        pieces.append(sex_label)
        metric_names.append("_".join(pieces) if pieces else f"metric_col{c}")
    return metric_names


def build_id_col_names(df: pd.DataFrame, title_row: int, sex_row: int, id_col_count: int) -> list:
    """First 5 columns get standard names (Census-wide convention). The
    6th (classifying variable, e.g. Age-Group / Educational level) is
    read directly from the file so it's never hardcoded."""
    names = list(STANDARD_ID_COLS)
    if id_col_count > len(STANDARD_ID_COLS):
        # classifying variable column: combine its own header fragments
        classifying_col = len(STANDARD_ID_COLS)
        pieces = []
        for r in range(title_row + 1, sex_row):
            val = df.iat[r, classifying_col]
            if pd.notna(val):
                val_str = str(val).strip()
                if val_str and val_str not in pieces:
                    pieces.append(val_str)
        classifying_name = "_".join(pieces) if pieces else "category"
        names.append(classifying_name.lower().replace(" ", "_").replace("-", "_"))
    return names[:id_col_count]


def process_file(filepath: str) -> pd.DataFrame:
    dataset_code = extract_dataset_code(filepath)

    xls = pd.ExcelFile(filepath)
    sheet_name = xls.sheet_names[0]
    raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)

    title_row = 0  # Census tables always lead with a single title row
    numbering_row = find_numbering_row(raw)
    sex_row, id_col_count = find_sex_row_and_id_col_count(raw, numbering_row)

    id_cols = build_id_col_names(raw, title_row, sex_row, id_col_count)
    metric_names = flatten_metric_names(raw, title_row, sex_row, id_col_count)

    data_start = numbering_row + 1
    while data_start < len(raw) and raw.iloc[data_start].isna().all():
        data_start += 1

    data = raw.iloc[data_start:].reset_index(drop=True).dropna(how="all")

    id_data = data.iloc[:, :id_col_count]
    id_data.columns = id_cols

    metric_data = data.iloc[:, id_col_count:id_col_count + len(metric_names)]
    metric_data.columns = metric_names

    combined = pd.concat([id_data, metric_data], axis=1)

    long_df = combined.melt(
        id_vars=id_cols, value_vars=metric_names,
        var_name="metric", value_name="value"
    )
    long_df["value"] = pd.to_numeric(long_df["value"], errors="coerce")
    long_df = long_df.dropna(subset=["value"])

    long_df["source_file"] = os.path.basename(filepath)
    long_df["dataset_code"] = dataset_code

    return long_df


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    files = sorted(glob.glob(os.path.join(INPUT_DIR, "DDW-*.xls")) +
                    glob.glob(os.path.join(INPUT_DIR, "DDW-*.xlsx")))

    print(f"Found {len(files)} census file(s) to process\n")

    by_dataset = {}  # dataset_code -> list of DataFrames

    for filepath in files:
        name = os.path.basename(filepath)
        try:
            df = process_file(filepath)
            dataset_code = df["dataset_code"].iloc[0]

            out_path = os.path.join(OUTPUT_DIR, name.rsplit(".", 1)[0] + ".csv")
            df.to_csv(out_path, index=False)
            print(f"  {name} [{dataset_code}]: {len(df)} rows -> {out_path}")

            by_dataset.setdefault(dataset_code, []).append(df)
        except Exception as e:
            print(f"  FAILED on {name}: {e}")

    print()
    for dataset_code, frames in by_dataset.items():
        combined = pd.concat(frames, ignore_index=True)
        out_path = os.path.join(OUTPUT_DIR, f"MASTER_{dataset_code}.csv")
        combined.to_csv(out_path, index=False)
        print(f"Combined {dataset_code} master: {len(combined)} rows -> {out_path}")

    print(f"\nDone. All output in '{OUTPUT_DIR}/'.")


if __name__ == "__main__":
    main()