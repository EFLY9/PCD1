# =============================================
# NEA Gasmet OCR Tool – FINAL WORKING VERSION
# Double-click this file → drag screenshot → get Excel
# =============================================

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from PIL import Image
import pytesseract
import re
from datetime import datetime
import os
import shutil
import io

# ------------------- CONFIG -------------------
st.set_page_config(
    page_title="NEA Gasmet OCR",
    page_icon="https://www.nea.gov.sg/docs/default-source/corporate/nea_logo.jpg",
    layout="wide"
)

TEMPLATE_PATH = "Gasmet Reference Limits -Chemical Burning.xlsx"

EXPECTED_COMPONENTS = [
    "Water Vapour", "Carbon Dioxide", "Carbon Monoxide", "Nitrous Oxide",
    "Acrolein", "Phenol", "Styrene", "M-Xylene", "P-Xylene", "O-Xylene",
    "Ammonia", "Benzene", "Crotonaldehyde", "Formaldehyde", "Hydrogen Chloride",
    "Hydrogen Fluoride", "Naphthalene", "Ethyl Benzene", "Toluene", "Ethylene"
]

if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = pd.DataFrame(columns=["Component", "Concentration"])
if "working_file" not in st.session_state:
    st.session_state.working_file = None

# ------------------- TITLE -------------------
st.markdown("""
<div style="text-align: center; padding: 20px;">
    <img src="https://www.nea.gov.sg/docs/default-source/corporate/nea_logo.jpg" width="80">
    <h1 style="color: #1B5E20; display: inline; margin-left: 15px;">NEA Gasmet OCR Tool</h1>
</div>
<p style="text-align: center; color: #555; font-size: 1.2rem;">
    Drag your screenshot below → auto-fill Excel in 5 seconds
</p>
""", unsafe_allow_html=True)
st.markdown("---")

# ------------------- FUNCTIONS (all your original ones + fixes) -------------------
def create_working_copy():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # ← FIXED: ccyymmdd_hhmmss
    name = os.path.splitext(os.path.basename(TEMPLATE_PATH))[0]
    new_file = f"{name}_{timestamp}.xlsx"
    shutil.copy2(TEMPLATE_PATH, new_file)
    return new_file

def extract_text_from_image(image):
    image = image.convert('L')
    from PIL import ImageEnhance
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    custom_config = r'--oem 3 --psm 6'
    return pytesseract.image_to_string(image, config=custom_config)

def normalize_component_name(name):
    return re.sub(r'[^a-z]', '', str(name).lower())

def parse_extracted_text(text):
    expected_map = {normalize_component_name(c): c for c in EXPECTED_COMPONENTS}
    results = {}
    patterns = [
        r'([A-Za-z\s\-\(\)]{3,}?)\s*[:\.\s]\s*([0-9,.]+)',
        r'([A-Za-z\s\-\(\)]{3,}?)\s+([0-9,.]+)'
    ]
    for line in text.split('\n'):
        for pattern in patterns:
            for comp, val in re.findall(pattern, line):
                norm = normalize_component_name(comp)
                for key, orig in expected_map.items():
                    if key in norm or norm in key:
                        results[orig] = val.replace(',', '')
                        break
    data = [{"Component": c, "Concentration": v} for c, v in results.items()]
    return pd.DataFrame(data)

def check_column_status(path):
    wb = load_workbook(path)
    ws = wb.active
    b = any(ws[f"B{r}"].value for r in range(2, 22))
    c = any(ws[f"C{r}"].value for r in range(2, 22))
    wb.close()
    return b, c

def populate_column(path, df, col):
    wb = load_workbook(path)
    ws = wb.active
    map = {normalize_component_name(r["Component"]): str(r["Concentration"]) for _, r in df.iterrows()}
    updates = 0
    for row in range(2, 22):
        comp = ws[f"A{row}"].value
        if comp:
            key = normalize_component_name(comp)
            if key in map:
                cell = ws[f"{col}{row}"]
                val = map[key]
                try:
                    cell.value = float(re.sub(r'[^\d.]', '', val))
                except:
                    cell.value = val
                updates += 1
    wb.save(path)
    wb.close()
    return True, updates

# ------------------- DRAG & DROP (WORKS EVERYWHERE) -------------------
st.header("📸 Drag Your Screenshot Here")
uploaded = st.file_uploader(
    "Win+Shift+S → select table → drag the floating image here\n"
    "⌘+Shift+5 → select → drag thumbnail here",
    type=["png", "jpg", "jpeg"],
    label_visibility="collapsed"
)

if uploaded:
    image = Image.open(uploaded)
    st.image(image, use_column_width=True)
    with st.spinner("Reading..."):
        text = extract_text_from_image(image)
        df = parse_extracted_text(text)
        if not df.empty:
            st.session_state.extracted_data = df
            st.success(f"Extracted {len(df)} components")
            st.balloons()

# ------------------- POPULATE EXCEL (FINAL FIX) -------------------
if not st.session_state.extracted_data.empty:
    st.header("📥 Generate Excel")
    
    # Always create NEW file + copy previous data
    new_file = create_working_copy()
    if st.session_state.working_file and os.path.exists(st.session_state.working_file):
        old_wb = load_workbook(st.session_state.working_file)
        new_wb = load_workbook(new_file)
        old_ws = old_wb.active
        new_ws = new_wb.active
        for col in ["B", "C"]:
            for row in range(2, 22):
                if old_ws[f"{col}{row}"].value:
                    new_ws[f"{col}{row}"].value = old_ws[f"{col}{row}"].value
        new_wb.save(new_file)
        old_wb.close()
        new_wb.close()
    
    b_has, c_has = check_column_status(new_file)
    target = "B" if not b_has else "C" if not c_has else st.radio("Both full – choose:", ["B", "C"], horizontal=True)
    
    col1, col2, col3 = st.columns([2,2,1])
    with col1:
        st.metric("Components", len(st.session_state.extracted_data))
    with col2:
        st.metric("Target", f"Column {target}")
    with col3:
        if st.button("Generate File", type="primary"):
            success, count = populate_column(new_file, st.session_state.extracted_data, target)
            st.session_state.working_file = new_file
            st.success(f"Done! {count} values → Column {target}")
            with open(new_file, "rb") as f:
                st.download_button(
                    "⬇️ Download Excel",
                    f.read(),
                    file_name=os.path.basename(new_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            if b_has and c_has:
                st.balloons()

st.markdown("---")
st.caption("Made with ❤️ for NEA officers – no more typing!")
