import streamlit as st
import pandas as pd
from striprtf.striprtf import rtf_to_text
from streamlit_sortables import sort_items
import re
from streamlit_gsheets import GSheetsConnection

st.title("Excel Processor with RTF Cleaning, Detail Extraction, Order Classification, and Google Sheet Data")

uploaded_files = st.file_uploader(
    "Upload 1 or 2 Excel files",
    accept_multiple_files=True,
    type=["xlsx", "xls"]
)

def load_file(file):
    df = pd.read_excel(file)
    for col in df.columns:
        if col.strip().lower() in ["doc no", "doc. no."]:
            df = df.rename(columns={col: "PI"})
    return df

def clean_rtf_column(df, column_name):
    if column_name in df.columns:
        df[column_name] = df[column_name].apply(
            lambda x: rtf_to_text(x) if isinstance(x, str) and x.strip().startswith("{\\rtf") else x
        )
    return df

def extract_details(text):
    if not isinstance(text, str):
        return pd.Series([None]*22)

    lines = [line.strip() for line in text.strip().splitlines() if line.strip()]

    detail_pairs = []
    remarks = [""] * 6
    remark_mode = False
    remark_idx = 0

    for line in lines:
        if remark_mode:
            if remark_idx < 6:
                if ":" in line:
                    parts = line.split(":", 1)
                    remark_text = parts[1].strip()
                else:
                    remark_text = line.strip()
                remarks[remark_idx] = remark_text
                remark_idx += 1
            continue

        if line.upper().startswith("ORDER REMARK"):
            remark_mode = True
            continue

        if ":" in line:
            parts = line.split(":", 1)
            detail = parts[0].strip()
            fabric = parts[1].strip()
            detail_pairs.append((detail, fabric))

    flat = []
    for i in range(10):
        if i < len(detail_pairs):
            flat.extend([detail_pairs[i][0], detail_pairs[i][1]])
        else:
            flat.extend(["", ""])

    return pd.Series(flat + remarks)

def classify_order_type(row):
    desc2_cols = [col for col in row.index if "Detail Description 2" in col]
    for col in desc2_cols:
        desc2 = str(row.get(col, "")).strip()
        if "(F)" in desc2:
            return pd.Series(["NEW ORDER", "FIXED PART"])
        if "(R)" in desc2:
            return pd.Series(["NEW ORDER", "REMOVABLE PART"])
        if "(I)" in desc2:
            return pd.Series(["NEW ORDER", "INNER PART"])

    item_code = str(row.get("Item Code", "")).strip()
    if item_code:
        first_letter = item_code[0].upper()
        if first_letter == "U":
            return pd.Series(["NEW ORDER", "CUSTOMADE"])
        elif first_letter == "Y":
            return pd.Series(["WARRANTY", "SERVICE"])

    return pd.Series(["NEW ORDER", "STANDARD"])

# Build column headers
detail_cols = []
for i in range(1, 11):
    detail_cols.extend([f"DETAIL {i}", f"FAB {i}"])
remark_cols = [f"REMARK ORDER {i}" for i in range(1, 7)]
all_extract_cols = detail_cols + remark_cols

# Default order
default_order = [
    "Doc Date",
    "PI",
    "Your Ref.",
    "Debtor Name_File1",
    "ORDER",
    "TYPE",
    "MODEL",
    "Qty",
    "DETAIL 1",
    "FAB 1",
    "DETAIL 2",
    "FAB 2",
    "DETAIL 3",
    "FAB 3",
    "DETAIL 4",
    "FAB 4",
    "DETAIL 5",
    "FAB 5",
    "DETAIL 6",
    "FAB 6",
    "DETAIL 7",
    "FAB 7",
    "DETAIL 8",
    "FAB 8",
    "DETAIL 9",
    "FAB 9",
    "DETAIL 10",
    "FAB 10",
    "REMARK ORDER 1",
    "REMARK ORDER 2",
    "REMARK ORDER 3",
    "REMARK ORDER 4",
    "REMARK ORDER 5",
    "REMARK ORDER 6",
    "REMARK DELIVERY"
]

# Google Sheets connection and caching
@st.cache_data(ttl=3000)
def load_google_sheet():
    conn = st.connection("gsheets", type=GSheetsConnection)
    return conn.read(worksheet="Sheet1")

try:
    # st.subheader("Google Sheet Data (Sheet1)")
    df_gsheets = load_google_sheet()
    df_gsheets.columns = [c.strip().lower() for c in df_gsheets.columns]
    # st.write("Normalized Google Sheet Columns:", df_gsheets.columns.tolist())
    # st.dataframe(df_gsheets)
except Exception as e:
    st.error(f"Error fetching Google Sheet: {e}")
    df_gsheets = pd.DataFrame()

if uploaded_files:
    if len(uploaded_files) == 1:
        df1 = load_file(uploaded_files[0])
        df1 = clean_rtf_column(df1, "Further Description")
        st.subheader("Uploaded File Data")
        st.dataframe(df1)

        st.subheader("Extracted Details Table")
        extracted = df1["Further Description"].apply(extract_details)
        extracted.columns = all_extract_cols
        st.dataframe(extracted)

        st.subheader("Order & Type Classification")
        order_type = df1.apply(classify_order_type, axis=1)
        order_type.columns = ["ORDER", "TYPE"]
        st.dataframe(order_type)

        combined_df = pd.concat([df1.reset_index(drop=True), extracted, order_type], axis=1)

        csv = combined_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Combined CSV",
            csv,
            "details_with_order_type.csv",
            "text/csv"
        )

    elif len(uploaded_files) == 2:
        df1 = load_file(uploaded_files[0])
        df2 = load_file(uploaded_files[1])

        df1 = clean_rtf_column(df1, "Further Description")
        df2 = clean_rtf_column(df2, "Further Description")

        merged_df = pd.merge(
            df1,
            df2,
            on="PI",
            how="outer",
            suffixes=("_File1", "_File2")
        )

        extracted = merged_df["Further Description"].apply(extract_details)
        extracted.columns = all_extract_cols

        order_type = merged_df.apply(classify_order_type, axis=1)
        order_type.columns = ["ORDER", "TYPE"]

        combined_df = pd.concat([merged_df.reset_index(drop=True), extracted.reset_index(drop=True), order_type], axis=1)
        filled_df = combined_df.copy()

        try:
            lookup_df = pd.DataFrame({
                "item_code_full": df_gsheets["item code"].astype(str).str.strip(),
                "model": df_gsheets["model"].astype(str).str.strip()
            })

            filled_df["Item Code"] = filled_df["Item Code"].astype(str).str.strip()
            filled_df["Detail Description 2"] = filled_df["Detail Description 2"].astype(str).str.strip()

            cross_item = filled_df.assign(key=1).merge(
                lookup_df.assign(key=1),
                on="key"
            ).drop("key", axis=1)

            mask_item = cross_item.apply(lambda x: x["item_code_full"].startswith(x["Item Code"]), axis=1)
            cross_item = cross_item[mask_item].sort_values("PI").drop_duplicates("PI")
            filled_df = filled_df.merge(
                cross_item[["PI", "item_code_full", "model"]],
                on="PI",
                how="left"
            ).rename(columns={"item_code_full": "ITEM_CODE_2", "model": "MODEL"})

            no_match_rows = filled_df["ITEM_CODE_2"].isna()
            if no_match_rows.any():
                filled_df.loc[no_match_rows, "Detail Description 2 Cleaned"] = (
                    filled_df.loc[no_match_rows, "Detail Description 2"]
                    .str.replace(r"\s*\([^)]*\)$", "", regex=True)
                    .str.strip()
                )

                to_match_df = filled_df.loc[no_match_rows].copy()
                cross_desc = to_match_df.assign(key=1).merge(
                    lookup_df.assign(key=1),
                    on="key"
                ).drop("key", axis=1)

                mask_desc = cross_desc.apply(
                    lambda x: x["item_code_full"].startswith(x["Detail Description 2 Cleaned"]),
                    axis=1
                )
                cross_desc = cross_desc[mask_desc].sort_values("PI").drop_duplicates("PI")

                filled_df.loc[no_match_rows, ["ITEM_CODE_2", "MODEL"]] = filled_df.loc[no_match_rows].merge(
                    cross_desc[["PI", "item_code_full", "model"]],
                    on="PI",
                    how="left"
                )[["item_code_full", "model"]].values

        except Exception as e:
            st.error(f"Error loading Google Sheet: {e}")

        filled_df["REMARK DELIVERY"] = filled_df["Further Description"].apply(
            lambda text: re.search(r"REMARK DELIVERY\s*:\s*(.*)", text, flags=re.IGNORECASE).group(1).strip()
            if isinstance(text, str) and re.search(r"REMARK DELIVERY\s*:\s*(.*)", text, flags=re.IGNORECASE)
            else ""
        )

        filled_df = filled_df[filled_df["PI"].notna() & (filled_df["PI"].astype(str).str.strip() != "")]

        with st.expander(label="Debugging Information", expanded=False):
            st.subheader("Google Sheet Data (Sheet1)")
            st.dataframe(df_gsheets)

            st.subheader("Merged Data (RTF Cleaned)")
            st.dataframe(merged_df)

            st.subheader("Extracted Details Table from Merged Data")
            st.dataframe(extracted)

            st.subheader("Order & Type Classification")
            st.dataframe(order_type)

            st.subheader("Table with Extracted REMARK DELIVERY")
            st.dataframe(filled_df[["PI", "REMARK DELIVERY"]])

            st.subheader("Table with Matched Item Code and Model (fast vectorized with cleaned fallback)")
            st.dataframe(filled_df)

            available_columns = [col for col in default_order if col in filled_df.columns]
            selected_columns = st.multiselect(
                "Choose columns (you can drop unwanted columns):",
                options=filled_df.columns.tolist(),
                default=available_columns
            )

            if selected_columns:
                st.subheader("Drag and Drop to Reorder Columns")
                reordered = sort_items(
                    items=selected_columns,
                    key="reorder"
                )

                reordered_df = filled_df[reordered]

                st.subheader("Final Reordered Data")
                st.dataframe(reordered_df)

        if selected_columns:
            renamed_df = reordered_df.rename(columns={
                "Doc Date": "TIMESTAMP",
                "PI": "PI NUMBER",
                "Your Ref.": "PO NUMBER",
                "Debtor Name_File1": "CUSTOMERS",
                "Qty": "QTY"
            })

            if "TIMESTAMP" in renamed_df.columns:
                renamed_df["TIMESTAMP"] = pd.to_datetime(renamed_df["TIMESTAMP"]).dt.strftime("%Y-%m-%d %H:%M:%S")

            st.subheader("Table with Renamed Columns")
            st.dataframe(renamed_df)

            copy_df = renamed_df.copy()
            if "TIMESTAMP" in copy_df.columns:
                cols = list(copy_df.columns)
                idx = cols.index("TIMESTAMP")
                cols.insert(idx + 1, " ")
                copy_df[" "] = ""
                copy_df = copy_df[cols]

            tsv_no_header = copy_df.to_csv(sep="\t", index=False, header=False)

            st.subheader("Copy Data Rows to Google Sheets (No Column Names, Extra Blank Column)")
            st.text_area(
                "1. Select all text below\n2. Copy (Ctrl+C)\n3. Paste into Google Sheets\n4. Use Data > Split text to columns > Tab",
                tsv_no_header,
                height=300
            )
        else:
            st.warning("Please select at least one column.")
