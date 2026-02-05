import streamlit as st
import pandas as pd
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import tempfile
import os

# ==================================================
# APP CONFIG
# ==================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

st.title("📊 Board Excel Intelligence Platform")
st.caption(
    "Universal Excel spreadsheet viewer with executive insights "
    "and Board-ready exports."
)

# ==================================================
# HELPERS
# ==================================================
def safe_display(df):
    return df.where(pd.notna(df), "")

def sanitize_columns(df):
    df.columns = [
        f"Column_{i}" if (c is None or (isinstance(c, float) and pd.isna(c))) else str(c)
        for i, c in enumerate(df.columns)
    ]
    return df

def is_numeric(series):
    return pd.api.types.is_numeric_dtype(series)

def needs_six_decimals(series):
    if not is_numeric(series):
        return False
    return any(abs(x - round(x, 2)) > 1e-6 for x in series.dropna())

def format_series(series, decimals):
    return series.apply(lambda x: f"{x:.{decimals}f}" if pd.notna(x) else "")

# ==================================================
# AGGRID BUILDER (USED BY BOTH VIEWS)
# ==================================================
def build_aggrid(df, raw_df, column_settings, height):
    gb = GridOptionsBuilder.from_dataframe(df)

    for col in df.columns:
        align_setting = column_settings[col]["align"]
        if align_setting == "Auto":
            align = "center" if is_numeric(raw_df[col]) else "left"
        else:
            align = align_setting.lower()

        gb.configure_column(
            col,
            wrapText=True,
            autoHeight=True,
            cellStyle={
                "textAlign": align,
                "whiteSpace": "normal",
                "lineHeight": "1.4",
                "borderRight": "1px solid #3a3a3a",
                "borderBottom": "1px solid #2a2a2a"
            }
        )

    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filter=True,
        wrapText=True,
        autoHeight=True
    )

    gb.configure_grid_options(
        domLayout="normal",
        suppressColumnVirtualisation=True,
        alwaysShowHorizontalScroll=True,
        rowHeight=38
    )

    AgGrid(
        df,
        gridOptions=gb.build(),
        height=height,
        theme="alpine",
        fit_columns_on_grid_load=True
    )

# ==================================================
# SIDEBAR — FILE NAVIGATION
# ==================================================
st.sidebar.header("📁 File Navigation")

uploaded_files = st.sidebar.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload one or more Excel files to begin.")
    st.stop()

file_map = {f.name: f for f in uploaded_files}

selected_file_name = st.sidebar.selectbox(
    "Select file",
    list(file_map.keys())
)

selected_file = file_map[selected_file_name]
xls = pd.ExcelFile(selected_file)

sheet_name = st.sidebar.selectbox(
    "Select sheet",
    xls.sheet_names
)

view_mode = st.sidebar.radio(
    "View Mode",
    ["Interactive Spreadsheet", "Executive View"]
)

# ==================================================
# READ DATA
# ==================================================
raw_df = pd.read_excel(selected_file, sheet_name=sheet_name)
raw_df = sanitize_columns(raw_df)
raw_df = safe_display(raw_df)

# ==================================================
# SIDEBAR — COLUMN FORMATTING
# ==================================================
st.sidebar.header("🧮 Column Formatting")
st.sidebar.caption("Unnamed / merged columns are auto-renamed safely.")

column_settings = {}

for idx, col in enumerate(raw_df.columns):
    with st.sidebar.expander(col):
        align = st.selectbox(
            "Alignment",
            ["Auto", "Left", "Center"],
            key=f"align_{idx}"
        )
        decimals = st.selectbox(
            "Decimals",
            ["Auto", 2, 6],
            key=f"dec_{idx}"
        )
        column_settings[col] = {"align": align, "decimals": decimals}

# ==================================================
# FORMAT DATA
# ==================================================
formatted_df = raw_df.copy()

for col in formatted_df.columns:
    settings = column_settings[col]
    dec_setting = settings["decimals"]

    if is_numeric(raw_df[col]):
        if dec_setting == "Auto":
            decimals = 6 if needs_six_decimals(raw_df[col]) else 2
        else:
            decimals = dec_setting
        formatted_df[col] = format_series(raw_df[col], decimals)
    else:
        formatted_df[col] = raw_df[col].astype(str)

# ==================================================
# INTERACTIVE VIEW
# ==================================================
if view_mode == "Interactive Spreadsheet":
    st.subheader(f"📄 Spreadsheet View – {sheet_name}")
    st.caption("Excel-like interactive spreadsheet. All columns fully visible.")

    build_aggrid(
        formatted_df,
        raw_df,
        column_settings,
        height=650
    )

# ==================================================
# EXECUTIVE VIEW
# ==================================================
else:
    st.subheader(f"📑 Executive View – {sheet_name}")
    st.caption("Board-ready executive spreadsheet with wrapped text and full visibility.")

    build_aggrid(
        formatted_df,
        raw_df,
        column_settings,
        height=520
    )

    st.markdown("---")
    st.subheader("📌 Executive Insights")

    numeric_df = raw_df.select_dtypes(include=np.number)

    st.info(f"Rows: {raw_df.shape[0]} | Columns: {raw_df.shape[1]}")

    if not numeric_df.empty:
        st.success(f"Numeric columns detected: {len(numeric_df.columns)}")
        st.success(f"Total numeric sum: {round(numeric_df.sum().sum(), 2)}")
        st.success(f"Maximum value: {round(numeric_df.max().max(), 2)}")
        st.success(f"Minimum value: {round(numeric_df.min().min(), 2)}")
    else:
        st.info("No numeric data detected in this sheet.")

    # PDF EXPORT
    if st.button("📄 Export Executive View as PDF"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            pdf_path = tmp.name

        c = canvas.Canvas(pdf_path, pagesize=A4)
        text = c.beginText(40, 800)
        text.setFont("Helvetica", 9)

        text.textLine(f"Executive View – {selected_file_name}")
        text.textLine(f"Sheet: {sheet_name}")
        text.textLine("")
        text.textLine(" | ".join(formatted_df.columns))
        text.textLine("-" * 120)

        for _, row in formatted_df.iterrows():
            text.textLine(" | ".join(map(str, row.values)))

        c.drawText(text)
        c.showPage()
        c.save()

        with open(pdf_path, "rb") as f:
            st.download_button(
                "Download PDF",
                data=f,
                file_name=f"{sheet_name}_Executive_View.pdf",
                mime="application/pdf"
            )

        os.remove(pdf_path)

# ==================================================
# CONSOLIDATED INSIGHTS
# ==================================================
st.markdown("---")
st.subheader("🌐 Consolidated Insights (All Uploaded Files)")

total_rows = 0
total_numeric_cols = 0
global_sum = 0
global_max = None

for f in uploaded_files:
    for sh in pd.ExcelFile(f).sheet_names:
        df = sanitize_columns(pd.read_excel(f, sh))
        total_rows += df.shape[0]

        num_df = df.select_dtypes(include=np.number)
        total_numeric_cols += num_df.shape[1]

        if not num_df.empty:
            global_sum += num_df.sum().sum()
            max_val = num_df.max().max()
            if global_max is None or max_val > global_max:
                global_max = max_val

st.success(f"Total rows across all files: {total_rows}")
st.success(f"Total numeric columns across all files: {total_numeric_cols}")
st.success(f"Global numeric sum: {round(global_sum, 2)}")

if global_max is not None:
    st.success(f"Highest value across all files: {round(global_max, 2)}")
else:
    st.info("No numeric values detected across uploaded files.")
