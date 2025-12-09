import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import date, datetime
import altair as alt
import os
import io

HISTORY_FILE = "history.csv"
DEFAULT_FILE = "data.xlsx"  # Primary Excel in repo

# ---------------------------
# Utilities
# ---------------------------
def clean_and_make_unique_columns(cols):
    cleaned = []
    for c in cols:
        s = "" if c is None else str(c)
        s = s.replace("\u00A0", " ")
        s = re.sub(r"\s+", " ", s).strip()
        cleaned.append(s)
    for i, s in enumerate(cleaned):
        if s == "" or s.lower().startswith("unnamed"):
            cleaned[i] = f"Column_{i}"
    unique = []
    counts = {}
    for s in cleaned:
        if s not in counts:
            counts[s] = 0
            unique.append(s)
        else:
            counts[s] += 1
            unique.append(f"{s}.{counts[s]}")
    return unique

def find_column(headers_map, possible_names):
    for name in possible_names:
        name_words = name.lower().split()
        for lower_col, actual_col in headers_map.items():
            if all(word in lower_col for word in name_words):
                return actual_col
    return None

def standardize_columns(df):
    orig_cols = list(df.columns)
    normalized = []
    for c in orig_cols:
        s = "" if c is None else str(c)
        s = s.replace("\u00A0", " ")
        s_norm = re.sub(r'[^a-z0-9]+', ' ', s.lower()).strip()
        normalized.append(s_norm)
    mapping = {}
    for orig, norm in zip(orig_cols, normalized):
        if "next" in norm and "service" in norm and ("hour" in norm or "hrs" in norm or "hr" in norm):
            mapping[orig] = "Next service hours"
        elif "run" in norm and ("hour" in norm or "hrs" in norm or "hr" in norm):
            mapping[orig] = "Run hours"
        elif ("last" in norm and "service" in norm and "hours" in norm) or ("hrs at last" in norm) or ("hrs last" in norm):
            mapping[orig] = "Hrs at last service"
        elif "dew" in norm and "inbuilt" in norm:
            mapping[orig] = "Dew point (inbuilt)"
        elif "dew" in norm and "external" in norm:
            mapping[orig] = "Dew point (external)"
        elif "element" in norm and ("temp" in norm or "temperature" in norm):
            mapping[orig] = "Element Temp"
        elif "oil" in norm and "level" in norm:
            mapping[orig] = "Oil Level"
        elif "oil" in norm and ("leak" in norm or "leakage" in norm):
            mapping[orig] = "Oil Leakage"
        else:
            mapping[orig] = orig
    df = df.rename(columns=mapping)
    return df

# ---------------------------
# Sidebar Options
# ---------------------------
st.sidebar.title("Excel Dashboard Options")
uploaded_file = st.sidebar.file_uploader("Upload your Excel workbook (optional)", type=["xlsx", "xls"])
save_history = st.sidebar.checkbox("Save snapshot to history.csv (archive)", value=False)

with st.sidebar.expander("Thresholds (hidden)"):
    DUE_SOON_HOURS = st.number_input("Due Soon Hours", value=336, step=1)
    DEWPOINT_INBUILT_MIN = st.number_input("Dew Point Inbuilt Min", value=3.0, format="%.2f")
    DEWPOINT_INBUILT_MAX = st.number_input("Dew Point Inbuilt Max", value=8.0, format="%.2f")
    DEWPOINT_EXTERNAL_MIN = st.number_input("Dew Point External Min", value=2.0, format="%.2f")
    DEWPOINT_EXTERNAL_MAX = st.number_input("Dew Point External Max", value=10.0, format="%.2f")
    ELEMENT_TEMP_WARNING_MIN = st.number_input("Element Temp Warning Min", value=100.0, format="%.1f")
    ELEMENT_TEMP_WARNING_MAX = st.number_input("Element Temp Warning Max", value=105.0, format="%.1f")
    ELEMENT_TEMP_TRIP_MIN = st.number_input("Element Temp Trip Min", value=110.0, format="%.1f")
    ELEMENT_TEMP_TRIP_MAX = st.number_input("Element Temp Trip Max", value=120.0, format="%.1f")

st.sidebar.subheader("Trend Chart Options")
trend_metric = st.sidebar.selectbox("Select metric to trend",
                                    ["Element Temp", "Dew point (inbuilt)", "Dew point (external)"])
trend_x_axis = st.sidebar.radio("X-axis for trend", ["Snapshot Date", "Run hours"])

# ---------------------------
# Main Dashboard
# ---------------------------
st.title("📊 Equipment Dashboard")

# ---------------------------
# Read uploaded file once (if any) and decide source
# ---------------------------
excel_bytes = None
if uploaded_file is not None:
    try:
        uploaded_file.seek(0)
        excel_bytes = uploaded_file.read()
    except Exception:
        excel_bytes = None

# Choose source and create xls_obj (keep same object for all reads)
if os.path.exists(DEFAULT_FILE):
    if excel_bytes is not None:
        source = "uploaded"
        xls_obj = io.BytesIO(excel_bytes)
    else:
        source = "repo"
        xls_obj = DEFAULT_FILE
else:
    if excel_bytes is not None:
        source = "uploaded"
        xls_obj = io.BytesIO(excel_bytes)
    else:
        st.error(f"No Excel file found. Please add '{DEFAULT_FILE}' to the repo or upload via sidebar.")
        st.stop()

# Show source message only once
if source == "repo":
    st.info(f"Using default Excel from repo: {DEFAULT_FILE}")
else:
    st.success("Using uploaded Excel file (overrides repo default).")

# ---------------------------
# Attempt to read sheet names
# ---------------------------
try:
    # If BytesIO, make sure pointer at start
    if isinstance(xls_obj, io.BytesIO):
        xls_obj.seek(0)
    xls = pd.ExcelFile(xls_obj)
    all_sheets = xls.sheet_names
except Exception as e:
    st.error(f"Unable to read Excel file: {e}")
    all_sheets = []

with st.sidebar.expander("Select Sheets"):
    selected_sheets = st.multiselect("Choose sheet(s) to display", options=all_sheets, default=all_sheets)

# ---------------------------
# Download button for current Excel
# ---------------------------
try:
    if isinstance(xls_obj, str):
        with open(xls_obj, "rb") as f:
            excel_data = f.read()
    else:
        xls_obj.seek(0)
        excel_data = xls_obj.read()

    st.download_button(
        label="📥 Download current Excel",
        data=excel_data,
        file_name="EquipmentDashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as e:
    st.error(f"Unable to prepare Excel for download: {e}")

# ---------------------------
# Process selected sheets (use the SAME xls_obj for all reads)
# ---------------------------
combined_rows = []
if selected_sheets:
    for sheet_name in selected_sheets:
        try:
            # Ensure file-like is rewound before each read
            if isinstance(xls_obj, io.BytesIO):
                xls_obj.seek(0)
                df = pd.read_excel(xls_obj, sheet_name=sheet_name, dtype=object)
            else:
                df = pd.read_excel(xls_obj, sheet_name=sheet_name, dtype=object)
        except Exception:
            # skip unreadable sheets
            continue

        if df is None or df.empty:
            continue

        df = standardize_columns(df)
        df.columns = clean_and_make_unique_columns(df.columns)
        df.reset_index(drop=True, inplace=True)
        df["Equipment"] = sheet_name if "Equipment" not in df.columns else df["Equipment"].fillna(sheet_name)

        headers_map = {col.lower(): col for col in df.columns}
        run_col = find_column(headers_map, ["Run hours"])
        next_col = find_column(headers_map, ["Next service hours"])
        dew_in_col = find_column(headers_map, ["Dew point (inbuilt)"])
        dew_ex_col = find_column(headers_map, ["Dew point (external)"])
        element_col = find_column(headers_map, ["Element Temp"])
        oil_level_col = find_column(headers_map, ["Oil Level"])
        oil_leak_col = find_column(headers_map, ["Oil Leakage"])

        for c in [run_col, next_col, dew_in_col, dew_ex_col, element_col]:
            if c and c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")

        possible_date_cols = ["Date", "Snapshot Date", "Date Recorded", "Reading Date", "Date/Time"]
        found_date_col = next((col for col in possible_date_cols if col in df.columns), None)

        if found_date_col:
            df["Snapshot Date"] = pd.to_datetime(df[found_date_col], errors="coerce", dayfirst=True)
        else:
            df["Snapshot Date"] = pd.to_datetime(date.today())

        df = df.dropna(subset=["Snapshot Date"])
        df = df.sort_values(by="Snapshot Date")

        for c in [run_col, next_col, dew_in_col, dew_ex_col, element_col]:
            if c and c in df.columns:
                df[c] = df[c].ffill().bfill()

        summary_df = df.groupby("Equipment").tail(1)
        sheet_canon = pd.DataFrame(index=summary_df.index)
        sheet_canon["Equipment"] = summary_df["Equipment"]
        for col in [run_col, next_col, dew_in_col, dew_ex_col, element_col]:
            if col in summary_df.columns:
                sheet_canon[col] = pd.to_numeric(summary_df[col], errors="coerce").round(2)
        sheet_canon["Oil Level"] = summary_df.get(oil_level_col, "")
        sheet_canon["Oil Leakage"] = summary_df.get(oil_leak_col, "")
        sheet_canon["Hours Left"] = (sheet_canon.get(next_col, 0) - sheet_canon.get(run_col, 0)).round(2)
        sheet_canon["Snapshot Date"] = summary_df["Snapshot Date"]
        sheet_canon["Date"] = summary_df["Snapshot Date"].dt.date

        combined_rows.append({"summary": sheet_canon, "trend": df})

# ---------------------------
# Summary, Issues, Trends
# ---------------------------
if combined_rows:
    summary_df_all = pd.concat([x["summary"] for x in combined_rows], ignore_index=True)
    summary_df_all.reset_index(drop=True, inplace=True)
    summary_df_all.columns = clean_and_make_unique_columns(summary_df_all.columns)

    def highlight(row):
        styles = [''] * len(row)
        for i, col in enumerate(row.index):
            val = row[col]
            if pd.isnull(val):
                continue
            try:
                if col == "Dew point (inbuilt)" and not (DEWPOINT_INBUILT_MIN <= val <= DEWPOINT_INBUILT_MAX):
                    styles[i] = "background-color: red"
                if col == "Dew point (external)" and not (DEWPOINT_EXTERNAL_MIN <= val <= DEWPOINT_EXTERNAL_MAX):
                    styles[i] = "background-color: red"
                if col == "Element Temp":
                    if ELEMENT_TEMP_TRIP_MIN <= val <= ELEMENT_TEMP_TRIP_MAX:
                        styles[i] = "background-color: red"
                    elif ELEMENT_TEMP_WARNING_MIN <= val <= ELEMENT_TEMP_WARNING_MAX:
                        styles[i] = "background-color: yellow"
                if col == "Hours Left":
                    if val <= 0:
                        styles[i] = "background-color: red"
                    elif val <= DUE_SOON_HOURS:
                        styles[i] = "background-color: yellow"
            except:
                pass
        return styles

    st.subheader("📋 Summary (latest row per equipment)")
    st.dataframe(summary_df_all.drop(columns=["Snapshot Date"], errors="ignore").style.apply(highlight, axis=1))

    # Equipment Issues
    issues = []
    for _, row in summary_df_all.iterrows():
        equip = row.get("Equipment", "Unknown")
        if "Hours Left" in row and pd.notnull(row["Hours Left"]):
            try:
                hl = float(row["Hours Left"])
                if hl <= 0:
                    issues.append({"msg": f"{equip}: Overdue", "color": "red"})
                elif hl <= DUE_SOON_HOURS:
                    issues.append({"msg": f"{equip}: Service due soon", "color": "yellow"})
            except:
                pass

        if "Dew point (inbuilt)" in row and pd.notnull(row["Dew point (inbuilt)"]):
            try:
                val = float(row["Dew point (inbuilt)"])
                if not (DEWPOINT_INBUILT_MIN <= val <= DEWPOINT_INBUILT_MAX):
                    issues.append({"msg": f"{equip}: Dewpoint out of range (inbuilt)", "color": "red"})
            except:
                pass

        if "Dew point (external)" in row and pd.notnull(row["Dew point (external)"]):
            try:
                val = float(row["Dew point (external)"])
                if not (DEWPOINT_EXTERNAL_MIN <= val <= DEWPOINT_EXTERNAL_MAX):
                    issues.append({"msg": f"{equip}: Dewpoint out of range (external)", "color": "red"})
            except:
                pass

        if "Element Temp" in row and pd.notnull(row["Element Temp"]):
            try:
                val = float(row["Element Temp"])
                if ELEMENT_TEMP_TRIP_MIN <= val <= ELEMENT_TEMP_TRIP_MAX:
                    issues.append({"msg": f"{equip}: Element Temp high high", "color": "red"})
                elif ELEMENT_TEMP_WARNING_MIN <= val <= ELEMENT_TEMP_WARNING_MAX:
                    issues.append({"msg": f"{equip}: Element Temp high", "color": "yellow"})
            except:
                pass

    if issues:
        st.subheader("⚠️ Equipment Issues")
        for issue in issues:
            msg = issue["msg"]
            color = issue.get("color", "normal")
            if color == "red":
                st.markdown(f"<span style='color:red; font-weight:bold;'>• {msg}</span>", unsafe_allow_html=True)
            elif color == "yellow":
                st.markdown(f"<span style='color:orange; font-weight:bold;'>• {msg}</span>", unsafe_allow_html=True)
            else:
                st.write(f"- {msg}")
    else:
        st.info("No issues detected.")

    if save_history:
        snapshot = summary_df_all.copy()
        snapshot["Archive Date"] = pd.to_datetime(datetime.now())
        if os.path.exists(HISTORY_FILE):
            history = pd.read_csv(HISTORY_FILE)
            snapshot = pd.concat([history, snapshot], ignore_index=True)
        snapshot.to_csv(HISTORY_FILE, index=False)
        st.success(f"Snapshot saved to {HISTORY_FILE}")

    # Trend Charts
    trend_df_all = pd.concat([x["trend"] for x in combined_rows], ignore_index=True)
    trend_df_all.reset_index(drop=True, inplace=True)
    trend_df_all.columns = clean_and_make_unique_columns(trend_df_all.columns)
    trend_df_all = trend_df_all.dropna(subset=["Snapshot Date"])

    min_date = trend_df_all["Snapshot Date"].min()
    max_date = trend_df_all["Snapshot Date"].max()
    trend_start = st.sidebar.date_input("Trend Start Date", value=min_date.date() if pd.notna(min_date) else date.today())
    trend_end = st.sidebar.date_input("Trend End Date", value=max_date.date() if pd.notna(max_date) else date.today())

    trend_df_all = trend_df_all[
        (trend_df_all["Snapshot Date"] >= pd.to_datetime(trend_start)) &
        (trend_df_all["Snapshot Date"] <= pd.to_datetime(trend_end))
    ]

    equipment_options = trend_df_all["Equipment"].unique().tolist()
    selected_equipment = st.sidebar.multiselect("Select Equipment for trend", options=equipment_options, default=equipment_options)
    x_col = "Snapshot Date" if trend_x_axis == "Snapshot Date" else "Run hours"

    if not trend_df_all.empty and trend_metric in trend_df_all.columns:
        st.subheader(f"📈 Trend of {trend_metric} per Equipment")
        for eq in selected_equipment:
            eq_df = trend_df_all[trend_df_all["Equipment"] == eq]
            if eq_df.empty:
                st.info(f"No data for {eq}")
                continue

            if "dew" in trend_metric.lower():
                min_val = DEWPOINT_INBUILT_MIN if "inbuilt" in trend_metric.lower() else DEWPOINT_EXTERNAL_MIN
                max_val = DEWPOINT_INBUILT_MAX if "inbuilt" in trend_metric.lower() else DEWPOINT_EXTERNAL_MAX
                eq_df["InRange"] = eq_df[trend_metric].between(min_val, max_val)
            elif "element" in trend_metric.lower():
                eq_df["InRange"] = ~(
                    ((eq_df[trend_metric] >= ELEMENT_TEMP_WARNING_MIN) & (eq_df[trend_metric] <= ELEMENT_TEMP_WARNING_MAX)) |
                    ((eq_df[trend_metric] >= ELEMENT_TEMP_TRIP_MIN) & (eq_df[trend_metric] <= ELEMENT_TEMP_TRIP_MAX))
                )
            else:
                eq_df["InRange"] = True

            line = alt.Chart(eq_df).mark_line(interpolate="monotone").encode(
                x=alt.X(x_col, type="temporal" if x_col=="Snapshot Date" else "quantitative"),
                y=alt.Y(trend_metric+":Q"),
                color=alt.condition(
                    alt.datum.InRange,
                    alt.value("steelblue"),
                    alt.value("red")
                )
            )
            points = alt.Chart(eq_df).mark_point(size=60).encode(
                x=alt.X(x_col, type="temporal" if x_col=="Snapshot Date" else "quantitative"),
                y=alt.Y(trend_metric+":Q"),
                color=alt.condition(
                    alt.datum.InRange,
                    alt.value("steelblue"),
                    alt.value("red")
                ),
                tooltip=["Snapshot Date", trend_metric]
            )

            chart = alt.layer(line, points).properties(title=f"{eq} - {trend_metric}")
            st.altair_chart(chart, use_container_width=True)
    else:
        st.subheader(f"📈 Trend of {trend_metric}")
        st.info("No trend data available for the selected metric / date range / equipment.")