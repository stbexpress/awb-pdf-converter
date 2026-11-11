import io
import re
import zipfile
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from jinja2 import Template
from weasyprint import HTML

st.set_page_config(page_title="IATA Air Waybill Generator", page_icon="✈️", layout="wide")

def excel_serial_to_date(val):
    """Convert Excel serial float/int OR YYYY-MM-DD timestamp string to DD-MMM-YYYY."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%b-%Y").upper()
        except ValueError:
            pass
    try:
        num = float(s)
        if 30000 < num < 70000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(num))
            return dt.strftime("%d-%b-%Y").upper()
    except ValueError:
        pass
    return s

def clean_cell(df, r, c):
    """Safely extract and strip a cell value from the DataFrame."""
    try:
        val = df.iat[r, c]
        if pd.isna(val):
            return ""
        s = str(val).strip()
        if s.endswith(".0") and s[:-2].isdigit():
            return s[:-2]
        return s
    except IndexError:
        return ""

def clean_multiline_spaces(text):
    """Collapse excessive horizontal padding spaces while preserving clean line breaks."""
    if not text:
        return ""
    lines = [re.sub(r"\s{2,}", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)

def find_row_containing(df, keyword, default_row):
    """Find the first row index containing a specific keyword."""
    for r in range(len(df)):
        for c in range(len(df.columns)):
            val = str(df.iat[r, c])
            if keyword.upper() in val.upper():
                return r
    return default_row

def parse_awb_excel(file_bytes, filename):
    """Extract all Air Waybill fields from your standard Excel/CSV template."""
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes), header=None)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)

    # Row 0: MAWB & HAWB
    prefix = clean_cell(df, 0, 0)
    origin_code = clean_cell(df, 0, 3)
    mawb_serial = clean_cell(df, 0, 7)
    hawb_no = clean_cell(df, 0, 60) or filename.rsplit(".", 1)
    awb_header = f"{prefix} - {origin_code} - {mawb_serial}".strip(" -")

    # Parties & Agents
    issued_by = clean_cell(df, 1, 46)
    shipper = clean_cell(df, 2, 0)
    consignee = clean_cell(df, 9, 0)
    issuing_agent = clean_cell(df, 14, 0)
    notify_party = clean_multiline_spaces(clean_cell(df, 14, 38))

    # Routing & Valuation
    airport_departure = clean_cell(df, 19, 0)
    to_dest = clean_cell(df, 21, 0)
    first_carrier = clean_cell(df, 21, 4)
    currency = clean_cell(df, 21, 37)
    declared_carriage = clean_cell(df, 21, 55) or "N.V.D"
    declared_customs = clean_cell(df, 21, 60)

    wt_val = clean_cell(df, 22, 42) or "PP"
    other_chgs = clean_cell(df, 22, 46) or wt_val

    airport_destination = clean_cell(df, 24, 0)
    flight_no = clean_cell(df, 24, 17)
    raw_flight_date = (
        excel_serial_to_date(clean_cell(df, 24, 29))
        or excel_serial_to_date(clean_cell(df, 24, 28))
        or excel_serial_to_date(clean_cell(df, 24, 30))
    )
    insurance = clean_cell(df, 24, 38) or "NIL"

    # Cartons, Weights & Rates
    total_cartons_text = clean_cell(df, 27, 0)
    pieces = clean_cell(df, 33, 0)
    gross_weight = clean_cell(df, 33, 4)
    weight_unit = clean_cell(df, 33, 10) or "KG"
    chargeable_weight = clean_cell(df, 33, 21)
    rate_charge = clean_cell(df, 33, 40) or "AS ARRANGED"

    # Shipping Marks & Description
    marks_raw = clean_cell(df, 36, 0)
    shipping_marks = marks_raw.replace("SHIPPING MARKS:", "").strip()
    goods_description = clean_multiline_spaces(clean_cell(df, 36, 17))

    # Dynamic lookup for Dhaka Agent & Execution Signatures
    agent_row = find_row_containing(df, "AGENT IN DHAKA", 54)
    dhaka_agent = clean_multiline_spaces(clean_cell(df, agent_row, 32))

    ship_sig_row = find_row_containing(df, "AS AGENT FOR CARRIER", 70)
    shipper_signature = clean_cell(df, ship_sig_row, 28)

    exec_row = ship_sig_row + 5
    exec_date = excel_serial_to_date(clean_cell(df, exec_row, 27)) or raw_flight_date
    flight_date = raw_flight_date or exec_date
    exec_place = clean_cell(df, exec_row, 46) or "HONG KONG"
    carrier_signature = clean_cell(df, exec_row, 55)

    is_collect = wt_val.upper() == "CC"
    return {
        "awb_header": awb_header,
        "hawb_no": hawb_no,
        "issued_by": issued_by,
        "shipper": shipper,
        "consignee": consignee,
        "issuing_agent": issuing_agent,
        "notify_party": notify_party,
        "airport_departure": airport_departure,
        "to_dest": to_dest,
        "first_carrier": first_carrier,
        "currency": currency,
        "wt_val": wt_val,
        "other_chgs": other_chgs,
        "declared_carriage": declared_carriage,
        "declared_customs": declared_customs,
        "airport_destination": airport_destination,
        "flight_no": flight_no,
        "flight_date": flight_date,
        "insurance": insurance,
        "total_cartons_text": total_cartons_text,
        "pieces": pieces,
        "gross_weight": gross_weight,
        "weight_unit": weight_unit,
        "chargeable_weight": chargeable_weight,
        "rate_charge": rate_charge,
        "shipping_marks": shipping_marks,
        "goods_description": goods_description,
        "dhaka_agent": dhaka_agent,
        "pp_weight": "" if is_collect else rate_charge,
        "cc_weight": rate_charge if is_collect else "",
        "pp_total": "" if is_collect else rate_charge,
        "cc_total": rate_charge if is_collect else "",
        "shipper_signature": shipper_signature,
        "exec_date": exec_date,
        "exec_place": exec_place,
        "carrier_signature": carrier_signature,
    }

# --- Web App User Interface ---
st.title("✈️ IATA Air Waybill (Excel → PDF) Converter")
st.markdown("Upload one or multiple shipment spreadsheets (`.xlsx`, `.xls`, or `.csv`) to generate single-page IATA Air Waybill PDFs.")

uploaded_files = st.file_uploader(
    "Drop your Air Waybill Excel/CSV files here",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

if uploaded_files:
    with open("template.html", "r", encoding="utf-8") as f:
        html_template = Template(f.read())

    generated_pdfs = []

    with st.spinner("Generating Air Waybill PDFs..."):
        for up_file in uploaded_files:
            try:
                data = parse_awb_excel(up_file.read(), up_file.name)
                rendered_html = html_template.render(**data)
                pdf_bytes = HTML(string=rendered_html).write_pdf()
                out_name = f"Air_Waybill_{data['hawb_no']}.pdf"
                generated_pdfs.append((out_name, pdf_bytes, data))
            except Exception as e:
                st.error(f"Error processing {up_file.name}: {e}")

    if generated_pdfs:
        st.success(f"Generated {len(generated_pdfs)} Air Waybill PDF(s) successfully!")

        if len(generated_pdfs) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for pdf_name, pdf_bytes, _ in generated_pdfs:
                    zf.writestr(pdf_name, pdf_bytes)
            st.download_button(
                label="📦 Download All PDFs as ZIP",
                data=zip_buffer.getvalue(),
                file_name="Air_Waybills_Batch.zip",
                mime="application/zip",
                type="primary"
            )

        st.subheader("Generated Air Waybills")
        for pdf_name, pdf_bytes, info in generated_pdfs:
            col1, col2, col3, col4 = st.columns()
            col1.write(f"**{info['hawb_no']}**")
            col2.write(info['shipper'].splitlines() if info['shipper'] else "")
            col3.write(f"{info['pieces']} CTNS | {info['gross_weight']} {info['weight_unit']}")
            col4.download_button(
                label="⬇️ Download PDF",
                data=pdf_bytes,
                file_name=pdf_name,
                mime="application/pdf",
                key=pdf_name
            )
