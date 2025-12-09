# app.py
import streamlit as st
import pandas as pd
import io
from typing import List, Set
import re

st.set_page_config(page_title="Surname Extractor — Master Sheet", layout="wide")

st.title("📥 Excel Surname Extractor — Drag & Drop up to 100 files")
st.markdown(
    "Upload your surnames list (or paste), then drag & drop Excel/CSV files to extract matching rows into one cleaned master sheet."
)

with st.expander("Instructions (open)", expanded=False):
    st.markdown(
        """
- Provide your list of surnames by uploading a text/csv/xlsx file or paste them (one per line or comma separated).
- Drag & drop up to 100 `.xlsx`, `.xls`, or `.csv` files in the file area.
- App will attempt to auto-detect common surname column names (e.g., surname, last_name, family, lname) and pre-select them.
- If names are in a full-name column, enable **Match only last-name token** to compare surname list to the last token of the name cell.
- Choose exact or substring matching (case-insensitive).
- Preview results, remove duplicates, and download as Excel or CSV.
- The app adds `__source_file` and `__sheet_name` columns for traceability.
"""
    )

# ---------- Step 1: Load surname list ----------
st.subheader("1) Provide surnames (500 names expected)")
col1, col2 = st.columns([2, 3])
with col1:
    uploaded_surnames_file = st.file_uploader(
        "Upload surnames file (.txt, .csv, .xlsx) — OR paste below",
        type=["txt", "csv", "xlsx"],
        accept_multiple_files=False,
        key="surnames_uploader",
    )
with col2:
    pasted = st.text_area(
        "Or paste surnames here (one per line or comma separated)", height=120, key="surnames_paste"
    )

def load_surnames_from_file(f) -> List[str]:
    try:
        fname = getattr(f, "name", "surnames")
        if fname.lower().endswith(".xlsx"):
            df = pd.read_excel(f, header=None)
            vals = df.stack().astype(str).tolist()
        else:
            # csv or txt
            # read bytes then decode if necessary
            data = f.read()
            text = data.decode("utf-8") if isinstance(data, (bytes, bytearray)) else str(data)
            vals = [s.strip() for s in text.replace(',', '\n').splitlines() if s.strip()]
        try:
            f.seek(0)
        except Exception:
            pass
        return vals
    except Exception as e:
        st.error(f"Failed to read surnames file: {e}")
        return []

surnames_list: List[str] = []
if uploaded_surnames_file is not None:
    surnames_list = load_surnames_from_file(uploaded_surnames_file)

if pasted and not surnames_list:
    items = [s.strip() for s in pasted.replace(',', '\n').splitlines() if s.strip()]
    surnames_list = items

if not surnames_list:
    st.warning("No surnames loaded yet — please upload a file or paste the surnames.")
else:
    # normalize surnames
    surnames_list = list(dict.fromkeys([s.strip() for s in surnames_list if s.strip()]))
    st.success(f"Loaded {len(surnames_list)} unique surnames.")
    if len(surnames_list) > 500:
        st.info("You uploaded more than 500 surnames — that's fine, the app will use all provided.")

# ---------- Step 2: Upload data files ----------
st.subheader("2) Drag & drop data files (.xlsx, .xls, .csv)")
uploaded_files = st.file_uploader(
    "Drop up to 100 files here (multi-file)", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="data_files"
)

# ---------- Settings ----------
st.subheader("3) Matching settings")
search_cols_input = st.text_input(
    "Specify column name(s) to search (comma separated). Leave blank to search all text columns:",
    value="",
    key="search_cols",
)
exact_match = st.checkbox("Exact match (cell == surname)", value=True, key="exact_match")
substring_match = st.checkbox("Substring match (cell contains surname)", value=False, key="substring_match")
match_last_token = st.checkbox("Match only last-name token (use when full names are in one column)", value=False, key="match_last_token")

# Auto-detect common surname columns if available
COMMON_SURNAME_COLS = [
    "surname","last_name","last name","lastname","family_name","family name","familyname",
    "lname","sur_name","sirname","surnames","last","lastname","name_last","family"
]
def detect_surname_columns(df: pd.DataFrame):
    cols = []
    df_cols = [c.lower() for c in df.columns]
    for candidate in COMMON_SURNAME_COLS:
        if candidate in df_cols:
            # return original-case column name(s)
            idx = df_cols.index(candidate)
            cols.append(df.columns[idx])
    return cols

def normalize_text(x):
    try:
        return str(x).strip().lower()
    except Exception:
        return ""

def surnames_set(surnames: List[str]) -> Set[str]:
    return set([s.strip().lower() for s in surnames if s and str(s).strip()])

TARGET_SURNAMES = surnames_set(surnames_list)

def last_name_token(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    # split on whitespace and common separators, keep last non-empty token
    tokens = re.split(r"[\s,;/\\]+", text.strip())
    tokens = [t for t in tokens if t]
    return tokens[-1].lower() if tokens else ""

def find_matches_in_dataframe(df: pd.DataFrame, search_columns: List[str], surnames: Set[str], exact: bool, substring: bool, last_token_only: bool):
    if df.empty:
        return pd.DataFrame()

    df_copy = df.copy()
    mask = pd.Series([False] * len(df_copy), index=df_copy.index)

    if not search_columns:
        text_cols = df_copy.select_dtypes(include=[object, "string"]).columns.tolist()
    else:
        text_cols = [c for c in search_columns if c in df_copy.columns]

    # If last_token_only is True and there is exactly one column specified, we will extract last token for that column
    for col in text_cols:
        col_series = df_copy[col].fillna("").map(normalize_text)
        if last_token_only:
            # use last token of the original (not normalized) value to be safer with punctuation
            col_series = df_copy[col].fillna("").map(last_name_token)

        if exact:
            mask = mask | col_series.isin(surnames)
        if substring:
            for s in surnames:
                if not s:
                    continue
                mask = mask | col_series.str.contains(s, na=False)
    return df_copy[mask]

# ---------- Main processing ----------
if uploaded_files and surnames_list:
    if len(uploaded_files) > 100:
        st.warning("You uploaded more than 100 files — only the first 100 will be processed.")
        uploaded_files = uploaded_files[:100]

    st.info(f"Processing {len(uploaded_files)} files. This may take a little while depending on file sizes.")

    results = []
    total_checked_rows = 0
    progress_bar = st.progress(0)

    # parsed user specified columns
    search_columns = [c.strip() for c in search_cols_input.split(",") if c.strip()]

    # If search_columns empty, attempt to auto-detect from the first sheet/file
    if not search_columns and uploaded_files:
        # peek into first file to detect common surname-like columns
        first = uploaded_files[0]
        fname = getattr(first, "name", "file0")
        try:
            if fname.lower().endswith(".csv"):
                peek_df = pd.read_csv(first, nrows=50)
            else:
                first.seek(0)
                sheets = pd.read_excel(first, sheet_name=None)
                # pick first non-empty sheet
                peek_df = None
                for sn, d in sheets.items():
                    if d is not None and not d.empty:
                        peek_df = d.head(50)
                        break
                if peek_df is None:
                    peek_df = pd.DataFrame()
        except Exception:
            try:
                first.seek(0)
            except Exception:
                pass
            peek_df = pd.DataFrame()

        detected = detect_surname_columns(peek_df) if not peek_df.empty else []
        if detected:
            # pre-fill search_columns with detected surname-like columns
            search_columns = detected
            st.info(f"Auto-detected surname column(s): {', '.join(detected)} — pre-selected.")

    for i, f in enumerate(uploaded_files):
        fname = getattr(f, "name", f"file_{i}")
        try:
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(f)
                total_checked_rows += len(df)
                matched = find_matches_in_dataframe(df, search_columns, TARGET_SURNAMES, exact_match, substring_match, match_last_token)
                if not matched.empty:
                    matched["__source_file"] = fname
                    matched["__sheet_name"] = "<csv>"
                    results.append(matched)
            else:
                try:
                    sheets = pd.read_excel(f, sheet_name=None)
                except Exception:
                    f.seek(0)
                    sheets = {"Sheet1": pd.read_excel(f)}
                for sheet_name, df in sheets.items():
                    if df is None or df.empty:
                        continue
                    total_checked_rows += len(df)
                    # if we have no explicit search_columns but matched was auto-detected earlier, use those; otherwise search all text columns
                    matched = find_matches_in_dataframe(df, search_columns, TARGET_SURNAMES, exact_match, substring_match, match_last_token)
                    if not matched.empty:
                        matched["__source_file"] = fname
                        matched["__sheet_name"] = sheet_name
                        results.append(matched)

        except Exception as e:
            st.error(f"Failed to process {fname}: {e}")

        progress_bar.progress((i + 1) / len(uploaded_files))

    if results:
        combined = pd.concat(results, ignore_index=True)
        combined = combined.drop_duplicates().reset_index(drop=True)

        st.success(f"Found {len(combined)} matching rows across files (checked ~{total_checked_rows} rows).")

        with st.expander("Preview first 200 matched rows", expanded=True):
            st.dataframe(combined.head(200))

        to_download_format = st.radio("Download format:", ("xlsx", "csv"), horizontal=True)
        if to_download_format == "xlsx":
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                combined.to_excel(writer, index=False, sheet_name="master")
            st.download_button(
                label="📥 Download Master Excel",
                data=buffer.getvalue(),
                file_name="master_surnames_master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            csv_bytes = combined.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Master CSV",
                data=csv_bytes,
                file_name="master_surnames_master.csv",
                mime="text/csv",
            )

    else:
        st.warning("No matches found in the uploaded files.")

else:
    if not surnames_list and not uploaded_files:
        st.info("Waiting for surnames list and data files to be uploaded.")
    elif not surnames_list:
        st.warning("Please load the surnames list before processing files.")
    elif not uploaded_files:
        st.warning("Please upload data files to process.")

st.markdown("---")
st.markdown("**Tips:** If your surname column has leading/trailing spaces or mixed-case, this app will normalize before matching. For full-name columns, enable 'Match only last-name token' to compare the final token. If you need accent-insensitive matching, augment `normalize_text()` to remove diacritics.")
