import io
import re
import difflib
from datetime import datetime, date
import dateutil.parser
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="AutoExcel Enterprise - Adaptive Date & Number Engine",
    page_icon="⚡",
    layout="wide"
)

# ----------------- Cleaning & Parsing Helpers ----------------- #

def clean_text(val):
    if val is None:
        return ""
    text = str(val).strip().lower()
    return re.sub(r"[_\W]+", " ", text).strip()

def calculate_similarity(t_val, d_val):
    t_clean = clean_text(t_val)
    d_clean = clean_text(d_val)
    if not t_clean or not d_clean:
        return 0.0
    if t_clean == d_clean:
        return 1.0
    t_words = set(t_clean.split())
    d_words = set(d_clean.split())
    overlap = len(t_words & d_words) / max(len(t_words | d_words), 1)
    ratio = difflib.SequenceMatcher(None, t_clean, d_clean).ratio()
    return (overlap * 0.45) + (ratio * 0.55)

def parse_flexible_date(raw_val, day_first=True):
    """
    Parses dates across formats:
    - Raw Excel serial numbers (e.g., 45321 -> 2024-01-28)
    - Python datetime/date instances
    - Text strings with configurable day/month priority
    """
    if raw_val is None:
        return None
    
    # Already a datetime/date object
    if isinstance(raw_val, (datetime, date, pd.Timestamp)):
        return raw_val

    # Numeric Excel serial (integers/floats typical in raw exports)
    if isinstance(raw_val, (int, float)):
        try:
            return from_excel(raw_val)
        except Exception:
            pass

    raw_str = str(raw_val).strip()
    if not raw_str:
        return None

    # Check if string is a numeric serial
    if raw_str.replace(".", "", 1).isdigit() and len(raw_str) <= 7:
        try:
            return from_excel(float(raw_str))
        except Exception:
            pass

    # Try explicit common standard formats first
    common_formats = [
        "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", 
        "%Y/%m/%d", "%d.%m.%Y", "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y", "%Y%m%d"
    ]
    # Prioritize format based on day_first preference
    if not day_first:
        common_formats.insert(0, "%m/%d/%Y")
        common_formats.insert(1, "%m-%d-%Y")
    else:
        common_formats.insert(0, "%d/%m/%Y")
        common_formats.insert(1, "%d-%m-%Y")

    for fmt in common_formats:
        try:
            return datetime.strptime(raw_str, fmt)
        except ValueError:
            continue

    # Fallback to dateutil fuzzy parser with dayfirst flag
    try:
        return dateutil.parser.parse(raw_str, dayfirst=day_first, fuzzy=True)
    except Exception:
        return None

def extract_sheet_columns(ws, header_row=1):
    columns = []
    for cell in ws[header_row]:
        val = cell.value
        col_letter = get_column_letter(cell.column)
        name = str(val).strip() if val is not None else f"Unnamed_{col_letter}"
        columns.append({
            "col_idx": cell.column,
            "col_letter": col_letter,
            "name": name,
            "label": f"[{col_letter}] {name}"
        })
    return columns

# ----------------- Streamlit UI ----------------- #

st.title("⚡ AutoExcel Enterprise")
st.caption("Adaptive Date Parsing, Number Transformation, and Formula Injection")

with st.sidebar:
    st.header("⚙️ General Settings")
    data_header_row = st.number_input("Source Header Row", min_value=1, value=1)
    template_header_row = st.number_input("Template Header Row", min_value=1, value=1)
    similarity_threshold = st.slider("Auto-Match Confidence (%)", 10, 95, 30, 5) / 100.0
    
    st.subheader("🗓️ Global Date Behavior")
    global_day_first = st.radio(
        "Ambiguous Date Preference (e.g. 03/04/2024)",
        options=["DD/MM/YYYY (Day First)", "MM/DD/YYYY (Month First)"],
        index=0
    ) == "DD/MM/YYYY (Day First)"

col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    data_file = st.file_uploader("1. Upload Source Data (.xlsx)", type=["xlsx"])
with col_upload2:
    template_file = st.file_uploader("2. Upload Blank Template (.xlsx)", type=["xlsx"])

if data_file and template_file:
    wb_data = openpyxl.load_workbook(data_file, data_only=False)
    wb_template = openpyxl.load_workbook(template_file, data_only=False)

    s_col1, s_col2 = st.columns(2)
    with s_col1:
        data_sheet_name = st.selectbox("Source Sheet", wb_data.sheetnames)
        ws_data = wb_data[data_sheet_name]
    with s_col2:
        template_sheet_name = st.selectbox("Template Sheet", wb_template.sheetnames)
        ws_template = wb_template[template_sheet_name]

    data_columns = extract_sheet_columns(ws_data, data_header_row)
    template_columns = extract_sheet_columns(ws_template, template_header_row)

    dropdown_options = ["-- Skip / Do Not Populate --"] + [d["label"] for d in data_columns]
    auto_mappings = []

    for t_item in template_columns:
        best_score = 0.0
        best_col = None
        for d_item in data_columns:
            score = calculate_similarity(t_item["name"], d_item["name"])
            if score > best_score:
                best_score = score
                best_col = d_item

        is_matched = best_score >= similarity_threshold and best_col is not None
        auto_mappings.append({
            "t_idx": t_item["col_idx"],
            "t_letter": t_item["col_letter"],
            "t_name": t_item["name"],
            "default_source": best_col["label"] if is_matched else "-- Skip / Do Not Populate --",
            "score": best_score if is_matched else 0.0
        })

    st.subheader("🛠️ Column Rules & Transformations")
    st.caption("Customize formulas, dynamic dates, and number formats per column:")

    final_rules = []

    for idx, mapping in enumerate(auto_mappings):
        with st.container():
            c1, c2, c3, c4 = st.columns([2.5, 3, 2.5, 3])

            with c1:
                st.markdown(f"**[{mapping['t_letter']}] {mapping['t_name']}**")
                conf = int(mapping["score"] * 100)
                st.caption(f"Match: {conf}%" if conf > 0 else "Unmapped")

            with c2:
                default_idx = dropdown_options.index(mapping["default_source"]) if mapping["default_source"] in dropdown_options else 0
                selected_source = st.selectbox(
                    "Source Column",
                    options=dropdown_options,
                    index=default_idx,
                    key=f"src_{idx}",
                    label_visibility="collapsed"
                )

            # Auto-detect if column name suggests a date
            is_date_hint = any(w in mapping["t_name"].lower() for w in ["date", "dt", "time", "dob", "created", "day"])
            default_mode_idx = 2 if is_date_hint else 0

            with c3:
                transform_mode = st.selectbox(
                    "Transformation",
                    options=[
                        "Direct Dynamic Link (=Source!X)",
                        "Text to Number (Clean & Formula)",
                        "Date Normalizer & Formatter",
                        "VLOOKUP Dynamic Match",
                        "Static Direct Value"
                    ],
                    index=default_mode_idx,
                    key=f"mode_{idx}",
                    label_visibility="collapsed"
                )

            with c4:
                extra_opt = {}
                if transform_mode == "Date Normalizer & Formatter":
                    d_c1, d_c2 = st.columns(2)
                    with d_c1:
                        extra_opt["output_format"] = st.selectbox(
                            "Target Format",
                            options=[
                                "DD-MM-YYYY", 
                                "YYYY-MM-DD", 
                                "DD/MM/YYYY", 
                                "MM/DD/YYYY", 
                                "DD-MMM-YYYY", 
                                "YYYY/MM/DD",
                                "DD MMMM YYYY"
                            ],
                            key=f"dt_target_{idx}"
                        )
                    with d_c2:
                        extra_opt["formula_mode"] = st.checkbox(
                            "Formula Link",
                            value=False,
                            key=f"dt_formula_{idx}",
                            help="If checked, injects =DATEVALUE(...) link instead of static date object."
                        )
                elif transform_mode == "Text to Number (Clean & Formula)":
                    extra_opt["num_format"] = st.selectbox(
                        "Number Format",
                        options=["Standard (#,##0.00)", "Integer (#,##0)", "Currency ($#,##0.00)", "Percentage (0.0%)"],
                        key=f"num_fmt_{idx}"
                    )
                elif transform_mode == "VLOOKUP Dynamic Match":
                    extra_opt["lookup_key"] = st.selectbox(
                        "Lookup Key Col",
                        options=[d["col_letter"] for d in template_columns],
                        key=f"vlk_{idx}"
                    )

            if selected_source != "-- Skip / Do Not Populate --":
                src_letter = selected_source.split("]")[0].replace("[", "").strip()
                matched_source = next((d for d in data_columns if d["col_letter"] == src_letter), None)
                if matched_source:
                    final_rules.append({
                        "t_idx": mapping["t_idx"],
                        "t_letter": mapping["t_letter"],
                        "d_idx": matched_source["col_idx"],
                        "d_letter": matched_source["col_letter"],
                        "mode": transform_mode,
                        "opts": extra_opt
                    })
        st.write("---")

    if st.button("🚀 Process & Generate Excel", type="primary", use_container_width=True):
        max_data_row = ws_data.max_row
        row_count = max_data_row - data_header_row

        if row_count <= 0:
            st.error("No valid data rows found under header.")
        else:
            with st.spinner("Normalizing dates and formatting fields..."):
                out_wb = openpyxl.Workbook()
                out_wb.remove(out_wb.active)

                # 1. Copy Source Data tab
                ws_source_target = out_wb.create_sheet(title="SourceData")
                for row in ws_data.iter_rows(values_only=True):
                    ws_source_target.append(row)

                # 2. Replicate Template Header & Formatting
                ws_out = out_wb.create_sheet(title="FilledTemplate")
                for r in range(1, template_header_row + 1):
                    for cell in ws_template[r]:
                        target_cell = ws_out.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            target_cell.font = cell.font.copy()
                            target_cell.border = cell.border.copy()
                            target_cell.fill = cell.fill.copy()
                            target_cell.number_format = cell.number_format
                            target_cell.alignment = cell.alignment.copy()

                start_out = template_header_row + 1

                # 3. Populate Rows
                for offset in range(row_count):
                    curr_out = start_out + offset
                    curr_src = data_header_row + 1 + offset

                    for rule in final_rules:
                        t_col = rule["t_idx"]
                        d_letter = rule["d_letter"]
                        d_col = rule["d_idx"]
                        mode = rule["mode"]
                        opts = rule["opts"]
                        target_cell = ws_out.cell(row=curr_out, column=t_col)

                        if mode == "Direct Dynamic Link (=Source!X)":
                            target_cell.value = f"=SourceData!{d_letter}{curr_src}"

                        elif mode == "Text to Number (Clean & Formula)":
                            target_cell.value = f"=VALUE(SourceData!{d_letter}{curr_src})"
                            fmt_map = {
                                "Standard (#,##0.00)": "#,##0.00",
                                "Integer (#,##0)": "#,##0",
                                "Currency ($#,##0.00)": "$#,##0.00",
                                "Percentage (0.0%)": "0.0%"
                            }
                            target_cell.number_format = fmt_map.get(opts.get("num_format"), "#,##0.00")

                        elif mode == "Date Normalizer & Formatter":
                            target_fmt = opts.get("output_format", "DD-MM-YYYY")
                            # Map format string to Excel standard format codes
                            excel_fmt_map = {
                                "DD-MM-YYYY": "DD-MM-YYYY",
                                "YYYY-MM-DD": "YYYY-MM-DD",
                                "DD/MM/YYYY": "DD/MM/YYYY",
                                "MM/DD/YYYY": "MM/DD/YYYY",
                                "DD-MMM-YYYY": "DD-MMM-YYYY",
                                "YYYY/MM/DD": "YYYY/MM/DD",
                                "DD MMMM YYYY": "DD MMMM YYYY"
                            }
                            excel_date_fmt = excel_fmt_map.get(target_fmt, "DD-MM-YYYY")

                            if opts.get("formula_mode", False):
                                # Dynamic formula mode
                                target_cell.value = f"=DATEVALUE(TEXT(SourceData!{d_letter}{curr_src}, \"YYYY-MM-DD\"))"
                            else:
                                # Native date parse mode
                                raw_val = ws_data.cell(row=curr_src, column=d_col).value
                                parsed_dt = parse_flexible_date(raw_val, day_first=global_day_first)
                                if parsed_dt:
                                    # Write as pure date (removes timestamp if present)
                                    target_cell.value = parsed_dt.date() if isinstance(parsed_dt, datetime) else parsed_dt
                                else:
                                    target_cell.value = f"=SourceData!{d_letter}{curr_src}"

                            target_cell.number_format = excel_date_fmt

                        elif mode == "VLOOKUP Dynamic Match":
                            l_key = opts.get("lookup_key", "A")
                            max_col_let = get_column_letter(ws_data.max_column)
                            target_cell.value = f'=VLOOKUP({l_key}{curr_out}, SourceData!A:{max_col_let}, {d_col}, FALSE)'

                        elif mode == "Static Direct Value":
                            target_cell.value = ws_data.cell(row=curr_src, column=d_col).value

                out_stream = io.BytesIO()
                out_wb.save(out_stream)
                out_stream.seek(0)

                st.success(f"Workbook compiled successfully with {row_count} rows!")
                st.download_button(
                    label="⬇️ Download Processed Excel File",
                    data=out_stream,
                    file_name="auto_populated_formatted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )