import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HG AMS ACAS Converter", page_icon="📄")
st.title("HG AMS ACAS Converter")

st.markdown(
    """
**What this app does**
1. Upload a source Excel file (e.g., `180-50508754 HG.xlsx`)
2. Generate a new Excel file with a fixed column order (hard-coded mapping)
3. Set `sender_state` to **GD** for all rows
4. Truncate `description` to the **left 75 characters**
5. Set `lastmile` to **CONSOL** for all rows
"""
)

# =====================================================
# Fixed column mapping (UPDATED: order 10 original -> "CITY NAME SHIPPER")
# Format: (order, original_column_name, new_column_name, default_value)
# - If original is empty -> output column is blank
# - If default_value is non-empty -> fill the whole column with that value
# =====================================================
COLUMNS = [
    (1, "BG Number", "consignor_item_id", ""),
    (2, "", "display_id", ""),
    (3, "Bag ID", "receptacle_id", ""),
    (4, "Tracking Number", "tracking_number", ""),
    (5, "SHIPPER", "sender_name", ""),
    (6, "", "sender_orgname", ""),
    (7, "SHIPPER ADDRESS", "sender_address1", ""),
    (8, "", "sender_address2", ""),
    (9, "", "sender_district", ""),
    (10, "CITY NAME SHIPPER", "sender_city", ""),  # <-- updated here
    (11, "", "sender_state", "GD"),
    (12, "", "sender_zip5", ""),
    (13, "", "sender_zip4", ""),
    (14, "COUNTRY CODE SHIPPER", "sender_country", ""),
    (15, "", "sender_phone", ""),
    (16, "", "sender_email", ""),
    (17, "", "sender_url", ""),
    (18, "Consignee Name", "recipient_name", ""),
    (19, "", "recipient_orgname", ""),
    (20, "Consignee Address", "recipient_address1", ""),
    (21, "", "recipient_address2", ""),
    (22, "", "recipient_district", ""),
    (23, "Consignee City", "recipient_city", ""),
    (24, "Consignee Province", "recipient_state", ""),
    (25, "Consignee Post Code", "recipient_zip5", ""),
    (26, "", "recipient_zip4", ""),
    (27, "Country of Destination", "recipient_country", ""),
    (28, "", "recipient_phone", ""),
    (29, "", "recipient_email", ""),
    (30, "", "recipient_addr_type", ""),
    (31, "", "return_name", ""),
    (32, "", "return_orgname", ""),
    (33, "", "return_address1", ""),
    (34, "", "return_address2", ""),
    (35, "", "return_district", ""),
    (36, "", "return_city", ""),
    (37, "", "return_state", ""),
    (38, "", "return_zip5", ""),
    (39, "", "return_zip4", ""),
    (40, "", "return_country", ""),
    (41, "", "return_phone", ""),
    (42, "", "return_email", ""),
    (43, "", "mail_type", ""),
    (44, "TOTAL QTY", "pieces", ""),
    (45, "WEIGHT", "weight", ""),
    (46, "", "length", ""),
    (47, "", "width", ""),
    (48, "", "height", ""),
    (49, "", "girth", ""),
    (50, "TOTAL DECLARE VALUE", "value", ""),
    (51, "", "machinable", ""),
    (52, "", "po_box_flag", ""),
    (53, "", "gift_flag", ""),
    (54, "", "commercial_flag", ""),
    (55, "", "customs_quantity_units", ""),
    (56, "", "dutiable", ""),
    (57, "", "duty_pay_by", ""),
    (58, "", "product", ""),
    (59, "PRODUCT DESCRIPTION", "description", ""),
    (60, "", "url", ""),
    (61, "", "sku", ""),
    (62, "COUNTRY CODE SHIPPER", "country_of_origin", ""),
    (63, "", "manufacturer", ""),
    (64, "HSCODE", "harmonization_code", ""),
    (65, "TOTAL DECLARE VALUE", "unit_value", ""),
    (66, "TOTAL QTY", "quantity", ""),
    (67, "TOTAL DECLARE VALUE", "total_value", ""),
    (68, "WEIGHT", "total_weight", ""),
    (69, "", "lastmile", "CONSOL"),
    (70, "", "item_id", ""),
    (71, "", "manufacture_name", ""),
    (72, "", "manufacture_address", ""),
    (73, "", "manufacture_city", ""),
    (74, "", "manufacture_state", ""),
    (75, "", "manufacture_zip_code", ""),
    (76, "", "manufacture_country", ""),
    (77, "", "manufacture_mid_code", ""),
    (78, "", "entry_no", ""),
    (79, "", "pga_product_code", ""),
    (80, "", "error_message", ""),
]

def normalize_col_name(x: str) -> str:
    return re.sub(r"\s+", " ", str(x).strip()).lower()

def extract_prefix(filename: str) -> str:
    base = re.sub(r"\.xlsx?$", "", filename, flags=re.I).strip()
    m = re.match(r"^(\d+(?:-\d+)+)", base)
    return m.group(1) if m else base

uploaded = st.file_uploader("Upload source .xlsx", type=["xlsx"])

if st.button("Generate Output File", type="primary", disabled=(uploaded is None)):
    try:
        src = pd.read_excel(uploaded, sheet_name=0, engine="openpyxl")
        src.columns = [str(c).strip() for c in src.columns]

        # Map normalized column name -> actual column name
        src_map = {normalize_col_name(c): c for c in src.columns}

        out = pd.DataFrame(index=src.index)

        # Build output columns in required order
        for _, orig, new, default in COLUMNS:
            if default:
                out[new] = default
            elif not orig:
                out[new] = ""
            else:
                key = normalize_col_name(orig)
                out[new] = src[src_map[key]] if key in src_map else ""

        # Enforced rules (always override)
        out["sender_state"] = "GD"
        out["lastmile"] = "CONSOL"

        if "description" in out.columns:
            out["description"] = (
                out["description"]
                .fillna("")
                .astype(str)
                .str.slice(0, 75)
            )

        prefix = extract_prefix(uploaded.name)
        output_filename = f"{prefix} - HG AMS ACAS.xlsx"

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            out.to_excel(writer, index=False, sheet_name="HG AMS ACAS")
        buf.seek(0)

        st.success("Done!")
        st.download_button(
            "Download output Excel",
            data=buf,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Preview (first 10 rows)"):
            st.dataframe(out.head(10), use_container_width=True)

    except Exception as e:
        st.exception(e)
