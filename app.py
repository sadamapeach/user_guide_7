import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def format_rupiah_percent(x):
    if pd.isna(x):
        return ""                   # hilangkan None / NaN
    return f"{format_rupiah(x)}%"   # pakai format_rupiah + %

def highlight_min_cell(row):
    styles = []
    
    # Cari nilai minimum, abaikan NaN
    numeric_vals = row[row.apply(lambda x: isinstance(x, (int, float)))]
    if not numeric_vals.empty:
        min_val = numeric_vals.min()
    else:
        min_val = None

    # Buat style per cell
    for val in row:
        if val == min_val:
            styles.append("background-color: #C6EFCE; color: #006100;")
        else:
            styles.append("")
    return styles

st.markdown(
    """
    <div style="font-size:1.75rem; font-weight:700; margin-bottom:9px">
        🧑‍🏫 User Guide: Standard Deviation
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
)
st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

# Divider custom
st.markdown(
    """
    <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="
        display: flex;
        align-items: center;
        height: 65px;
        margin-bottom: 10px;
    ">
        <div style="text-align: justify; font-size: 15px;">
            <span style="color: #ED1C24; font-weight: 800;">
            Standard Deviation</span>
            measures the price variation across vendors to assess pricing stability 
            and the consistency of commercial offers.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px;">
            The input file required for this menu should be a 
            <span style="color: #FF69B4; font-weight: 500;">single file containing single sheet</span>, in eather 
            <span style="background:#C6EFCE; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> or 
            <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xls</span> format. 
            The table structure is as follows:
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["Scope", "Desc", "Vendor A", "Vendor B", "Vendor C", "Vendor D", "Vendor E"]
df = pd.DataFrame([[""] * len(columns) for _ in range(3)], columns=columns)

st.dataframe(df, hide_index=True)

# Buat DataFrame 1 row
st.markdown("""
<table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
    <tr>
        <td style="border: 1px solid gray; width: 15%;">Sheet1</td>
        <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">single sheet only</td>
    </tr>
</table>
""", unsafe_allow_html=True)

st.markdown("###### Description:")
st.markdown(
    """
    <div style="font-size:15px;">
        <ul>
            <li>
                <span style="display:inline-block; width:100px;">Scope & Desc</span>: non-numeric columns
            </li>
            <li>
                <span style="display:inline-block; width:100px;">Vendor A - E</span>: numeric columns
            </li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a 
            <span style="font-weight: bold;">dynamic table</span>, allowing users to enter any number of non-numeric
            and numeric columns. Users have the freedom to name the columns as they wish. The system logic relies on 
            <span style="font-weight: bold;"> column indices</span>, not specific column names.
        </div>
    """,
    unsafe_allow_html=True 
)

st.divider()
st.markdown("#### Constraint")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
            To ensure this menu works correctly, users need to follow certain rules regarding
            the dataset structure.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. COLUMN ORDER]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
            When creating tables, it is important to follow the specified column structure. Columns 
            <span style="font-weight: bold;">must</span> be arranged in the following order:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            Non-Numeric Columns → Numeric Columns
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px">
            this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
            <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. NUMBER COLUMN]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Please refer the table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["No", "Scope", "Desc", "Vendor A", "Vendor B", "Vendor C"]
data = [
    [1] + [""] * (len(columns) - 1),
    [2] + [""] * (len(columns) - 1),
    [3] + [""] * (len(columns) - 1)
]
df = pd.DataFrame(data, columns=columns)

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains a 
            <span style="font-weight: bold;">"No"</span> column. The "No" column is prohibited in this
            menu, as it will be treated as a numeric column by the system, which violates the constraint
            described in point 1.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables <span style="color: #FF69B4; font-weight: 700;">
            do not need to start from cell A1</span>. However, ensure
            that the cells above and to the left of the table are empty, as shown in the example below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["", "A", "B", "C", "D", "E", "F", "G"]

# Buat 5 baris kosong
df = pd.DataFrame([[""] * len(columns) for _ in range(6)], columns=columns)

# Isi kolom pertama dengan 1–7
df.iloc[:, 0] = [1, 2, 3, 4, 5, 6]

# Header bagian kedua
df.loc[1, ["B", "C", "D", "E", "F"]] = ["Scope", "UoM", "Vendor A", "Vendor B", "Vendor C"]

# Data Software & Hardware
df.loc[2, ["B", "C", "D", "E", "F"]] = ["WP1", "Site", "1.000", "2.000", "3.000"]
df.loc[3, ["B", "C", "D", "E", "F"]] = ["WP2", "Site", "4.800", "5.000", "5.200"]
df.loc[4, ["B", "C", "D", "E", "F"]] = ["WP3", "Site", "3.650", "3.450", "3.250"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top:-10px;">
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:blue-badge[4. TOTAL COLUMN & TOTAL ROW]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            You are not allowed to add a
            <span style="font-weight: 700;">TOTAL COLUMN</span> or
            <span style="font-weight: 700;">TOTAL ROW</span>!
            Please refer to the example table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Scope", "Vendor A", "Vendor B", "Vendor C", "TOTAL"]
data = [
    ["WP1", "1.000", "2.000", "3.000", "6.000"],
    ["WP2", "4.800", "5.000", "5.200", "15.000"],
    ["TOTAL", "5.800", "7.000", "8.200", "21.000"],
]
df = pd.DataFrame(data, columns=columns)

def red_highlight(row):
    styles = [""] * len(row)

    # Highlight ROW "TOTAL"
    if row["Scope"] == "TOTAL":
        styles = ["color: #FF4D4D;" for _ in row]
    else:
        # Highlight COLUMN "TOTAL"
        total_col_index = row.index.get_loc("TOTAL")
        styles[total_col_index] = "color: #FF4D4D;"

    return styles

df_styled = df.style.apply(red_highlight, axis=1)

st.dataframe(df_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not permitted</span>! 
            If you add it, the system will treat it as a regular row and include it in the calculations.
        </div>
    """,
    unsafe_allow_html=True
)


st.divider()

st.markdown("#### What is Displayed?")

# Path file Excel yang sudah ada
file_path = "dummy dataset.xlsx"

# Buka file sebagai binary
with open(file_path, "rb") as f:
    file_data = f.read()

# Markdown teks
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
        You can try this menu by downloading the dummy dataset using the button below: 
    </div>
    """,
    unsafe_allow_html=True
)

@st.fragment
def release_the_balloons():
    st.balloons()

# Download button untuk file Excel
st.download_button(
    label="Dummy Dataset",
    data=file_data,
    file_name="Dummy Dataset - Standard Deviation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    on_click=release_the_balloons,
    type="primary",
    use_container_width=True,
)

st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        Based on this dummy dataset, the menu will produce the following results.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. BIDDER'S RANK]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will rank the scope prices for each vendor and display them in a table as follows.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Scope", "Vendor A", "Vendor B", "Vendor C"]
data = [
    ["WP1", 1, 2, 3],
    ["WP2", 3, 1, 2],
    ["WP3", 3, 2, 1],
]
df_bid_rank = pd.DataFrame(data, columns=columns)
st.dataframe(df_bid_rank, hide_index=True)

st.write("")
st.markdown("**:orange-badge[2. RANK-1 DEVIATION (%)]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system then calculate the deviation from the lowest price (1st rank). A deviation value of  
            <span style="color: #FF69B4; font-weight: 700;">"0%"</span>
            indicates that the vendor is the one offering the lowest price.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Scope", "Vendor A", "Vendor B", "Vendor C"]
data = [
    ["WP1", 0, 27.35, 27.39],
    ["WP2", 22.57, 0, 0.07],
    ["WP3", 130.1, 66.56, 0],
]

df_rank_dev = pd.DataFrame(data, columns=columns)
num_cols = df_rank_dev.select_dtypes(include=["number"]).columns
format_dict = {col: format_rupiah_percent for col in num_cols}

df_rank_dev_styled = (
    df_rank_dev.style
    .format(format_dict)
    .apply(highlight_min_cell, axis=1)
)

st.dataframe(df_rank_dev_styled, hide_index=True)

st.write("")
st.markdown("**:yellow-badge[3. Summary Deviation (%)]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            After that, the system will generate a summary that helps users analyze each vendor's rank and 
            its deviation compared to the first-ranked bidder.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Scope", "1st Rank", "Best Price", "2nd Rank", "Dev. 2nd to 1st (%)", "3rd Rank", "Dev. 3rd to 1st (%)"]
data = [
    ["WP1", "Vendor A", 10310, "Vendor B", 27.35, "Vendor C", 27.39],
    ["WP2", "Vendor B", 14242, "Vendor C", 0.07, "Vendor A", 22.57],
    ["WP3", "Vendor C", 3242, "Vendor B", 66.56, "Vendor A", 130.1],
]

df_sum_dev = pd.DataFrame(data, columns=columns)
format_dict = {}

# Kolom "Best Price"
if "Best Price" in df_sum_dev.columns:
    format_dict["Best Price"] = format_rupiah

# Kolom deviasi (%)
for col in df_sum_dev.columns:
    if col.startswith("Dev. ") and col.endswith("(%)"):
        format_dict[col] = format_rupiah_percent

df_sum_dev_styled = df_sum_dev.style.format(format_dict)

st.dataframe(df_sum_dev_styled, hide_index=True)

st.write("")
st.markdown("**:green-badge[4. VISUALIZATION]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu visualizes the ranking for each scope, where the system loops through the tabs based on the number of scopes.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2, tab3 = st.tabs(["WP1", "WP2", "WP3"])

with tab1:
    st.image("assets/1.png")

with tab2:
    st.image("assets/2.png")

with tab3:
    st.image("assets/3.png")
    
st.write("")
st.markdown("**:blue-badge[5. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Super Button</span> 
            feature where all dataframes generated by the system can be downloaded as a single file with multiple sheets. 
            You can also customize the order of the sheets. The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

dataframes = {
    "Bidder's Rank": df_bid_rank,
    "Rank-1 Deviation (%)": df_rank_dev,
    "Summary Deviation (%)": df_sum_dev,
}

# Tampilkan multiselect
selected_sheets = st.multiselect(
    "Select sheets to download in a single Excel file:",
    options=list(dataframes.keys()),
    default=list(dataframes.keys())  # default semua dipilih
)

def generate_multi_sheet_excel(selected_sheets, df_dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet in selected_sheets:
            df = df_dict[sheet].copy()

            # ========= SPECIAL SHEET =========
            if sheet == "Rank-1 Deviation (%)":
                df_to_write = df.copy()
                numeric_cols = [
                    c for c in df_to_write.columns
                    if pd.api.types.is_numeric_dtype(df_to_write[c])
                ]

                df_to_write.to_excel(writer, index=False, sheet_name=sheet)
                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                fmt_pct = workbook.add_format({'num_format': '#,##0.0"%"'})
                fmt_min = workbook.add_format({
                    "bg_color": "#D9EAD3",
                    'num_format': '#,##0.0"%"'
                })

                # ===== WRITE + FORMAT CELL =====
                for r, row in enumerate(df_to_write.itertuples(index=False), start=1):
                    numeric_vals = {
                        i: val for i, val in enumerate(row)
                        if df_to_write.columns[i] in numeric_cols and pd.notna(val)
                    }
                    min_val = min(numeric_vals.values()) if numeric_vals else None

                    for c, col in enumerate(df_to_write.columns):
                        val = row[c]

                        if pd.isna(val):
                            worksheet.write_blank(r, c, None)
                        elif col in numeric_cols:
                            fmt = fmt_min if val == min_val else fmt_pct
                            worksheet.write_number(r, c, val, fmt)
                        else:
                            worksheet.write(r, c, val)

                # ===== AUTOFIT (KEEP FORMAT!) =====
                for i, col in enumerate(df_to_write.columns):
                    width = max(
                        len(str(col)),
                        df_to_write[col].astype(str).map(len).max()
                    ) + 2

                    worksheet.set_column(
                        i, i,
                        width,
                        fmt_pct if col in numeric_cols else None
                    )

                continue

            # ========= DEFAULT SHEETS =========
            df.to_excel(writer, index=False, sheet_name=sheet)
            workbook  = writer.book
            worksheet = writer.sheets[sheet]

            fmt_rp  = workbook.add_format({'num_format': '#,##0'})
            fmt_pct = workbook.add_format({'num_format': '#,##0.0"%"'})

            numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
            pct_cols = [c for c in df.columns if "%" in c]

            # ===== REWRITE CELLS (SAFE) =====
            for r, row in enumerate(df.itertuples(index=False), start=1):
                for c, col in enumerate(df.columns):
                    val = row[c]

                    if pd.isna(val) or (isinstance(val, float) and np.isinf(val)):
                        worksheet.write_blank(r, c, None)
                    elif col in pct_cols:
                        worksheet.write_number(r, c, val, fmt_pct)
                    elif col in numeric_cols:
                        worksheet.write_number(r, c, val, fmt_rp)
                    else:
                        worksheet.write(r, c, val)

            # ===== AUTOFIT (KEEP FORMAT!) =====
            for i, col in enumerate(df.columns):
                width = max(
                    len(str(col)),
                    df[col].astype(str).map(len).max()
                ) + 2

                if col in pct_cols:
                    worksheet.set_column(i, i, width, fmt_pct)
                elif col in numeric_cols:
                    worksheet.set_column(i, i, width, fmt_rp)
                else:
                    worksheet.set_column(i, i, width)

    output.seek(0)
    return output.getvalue()

# ---- DOWNLOAD BUTTON ----
if selected_sheets:
    excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

    st.download_button(
        label="Download",
        data=excel_bytes,
        file_name="Super Botton - Standard Deviation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.write("")
st.divider()

st.markdown("#### Video Tutorial")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">YouTube</span> link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/nS3xERgggqA?si=c3XHqEbMoLENDt62")