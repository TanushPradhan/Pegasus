import streamlit as st
import pandas as pd
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

st.title("📊 Board Excel Intelligence Platform")
st.caption("Universal Excel spreadsheet viewer with executive insights and Board-ready presentation")

# =====================================================
# SIDEBAR – FILE NAVIGATION
# =====================================================
st.sidebar.header("📁 File Navigation")

uploaded_files = st.sidebar.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Please upload one or more Excel files to continue.")
    st.stop()

file_names = [f.name for f in uploaded_files]
selected_file_name = st.sidebar.selectbox("Select file", file_names)
selected_file = next(f for f in uploaded_files if f.name == selected_file_name)

xls = pd.ExcelFile(selected_file)
sheet_name = st.sidebar.selectbox("Select sheet", xls.sheet_names)

# =====================================================
# VIEW MODE
# =====================================================
view_mode = st.sidebar.radio(
    "View Mode",
    ["Interactive Spreadsheet", "Executive View"]
)

# =====================================================
# HIGHLIGHTING CONTROLS (SIDEBAR ONLY)
# =====================================================
st.sidebar.header("🎨 Highlighting")

preset_highlight = st.sidebar.checkbox(
    "Board Highlight Preset (Costs & Totals)",
    value=False
)

highlight_color = st.sidebar.color_picker(
    "Highlight color",
    "#FFF3B0"
)

# =====================================================
# LOAD EXCEL (SAFE)
# =====================================================
raw_df = pd.read_excel(
    selected_file,
    sheet_name=sheet_name,
    header=0
)

df = raw_df.copy()
df = df.replace({np.nan: ""})

# =====================================================
# COLUMN ANALYSIS (AFTER df EXISTS)
# =====================================================
all_columns = df.columns.tolist()

numeric_columns = [
    col for col in df.columns
    if pd.to_numeric(df[col], errors="coerce").notna().sum() > 0
]

default_text_columns = [df.columns[0]] if len(df.columns) > 0 else []

highlight_columns = st.sidebar.multiselect(
    "Highlight columns (manual)",
    all_columns
)

# ✅ APPLY BOARD PRESET *AFTER* df EXISTS (FIXED)
if preset_highlight:
    highlight_columns = [
        col for col in df.columns
        if any(k in col.lower() for k in ["cost", "budget", "total", "sum"])
    ]

highlight_row = st.sidebar.number_input(
    "Highlight row (1-based, optional)",
    min_value=0,
    max_value=len(df),
    step=1
)

# =====================================================
# AGGRID CONFIG
# =====================================================
gb = GridOptionsBuilder.from_dataframe(df)

for col in df.columns:
    is_numeric = col in numeric_columns
    align = "center" if is_numeric else "left"

    style = {
        "textAlign": align,
        "whiteSpace": "normal",
        "lineHeight": "1.4",
        "borderRight": "1px solid #3a3a3a",
        "borderBottom": "1px solid #2a2a2a",
    }

    if col in highlight_columns:
        style["backgroundColor"] = highlight_color
        style["fontWeight"] = "600"

    gb.configure_column(
        col,
        wrapText=True,
        autoHeight=True,
        cellStyle=style
    )

# =====================================================
# ROW HIGHLIGHTING (SAFE)
# =====================================================
if highlight_row > 0 and highlight_row <= len(df):
    df["_row_flag"] = ""
    df.loc[highlight_row - 1, "_row_flag"] = "highlight"

    gb.configure_column("_row_flag", hide=True)

    gb.configure_grid_options(
        getRowStyle={
            "styleConditions": [
                {
                    "condition": "params.data._row_flag === 'highlight'",
                    "style": {
                        "backgroundColor": "#E3F2FD",
                        "fontWeight": "600"
                    }
                }
            ]
        }
    )

# =====================================================
# GRID OPTIONS
# =====================================================
gb.configure_grid_options(
    suppressColumnVirtualisation=True,
    alwaysShowHorizontalScroll=True,
    domLayout="normal"
)

grid_options = gb.build()

# =====================================================
# INTERACTIVE VIEW
# =====================================================
if view_mode == "Interactive Spreadsheet":
    st.subheader(f"📄 Spreadsheet View – {sheet_name}")

    AgGrid(
        df,
        gridOptions=grid_options,
        height=540,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        theme="streamlit"
    )

# =====================================================
# EXECUTIVE VIEW
# =====================================================
else:
    st.subheader(f"📊 Executive View – {sheet_name}")

    numeric_df = df[numeric_columns].apply(
        pd.to_numeric, errors="coerce"
    )

    total_sum = numeric_df.sum().sum() if not numeric_df.empty else 0
    max_value = numeric_df.max().max() if not numeric_df.empty else 0

    col1, col2 = st.columns(2)
    col1.metric("Total Numeric Sum", f"{round(total_sum, 2)}")
    col2.metric("Highest Value", f"{round(max_value, 2)}")

    st.markdown("---")

    AgGrid(
        df,
        gridOptions=grid_options,
        height=540,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        theme="streamlit"
    )
