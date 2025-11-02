import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Protection
from PIL import Image
import pytesseract
import re
from datetime import datetime
import os
import io
import shutil

# Template path
TEMPLATE_PATH = "Gasmet Reference Limits -Chemical Burning.xlsx"

# Expected components in exact order
EXPECTED_COMPONENTS = [
    "Water Vapour", "Carbon Dioxide", "Carbon Monoxide", "Nitrous Oxide",
    "Acrolein", "Phenol", "Styrene", "M-Xylene", "P-Xylene", "O-Xylene",
    "Ammonia", "Benzene", "Crotonaldehyde", "Formaldehyde", "Hydrogen Chloride",
    "Hydrogen Fluoride", "Naphthalene", "Ethyl Benzene", "Toluene", "Ethylene"
]

# Initialize session state
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = pd.DataFrame(columns=["Component", "Concentration"])
if "upload_count" not in st.session_state:
    st.session_state.upload_count = 0
if "working_file" not in st.session_state:
    st.session_state.working_file = None
if "last_uploaded_file" not in st.session_state:
    st.session_state.last_uploaded_file = None

st.set_page_config(page_title="Excel OCR Populator", page_icon="🧾", layout="wide")

st.title("🧾 Excel Template OCR Populator")
st.markdown("---")


# ==================== FUNCTIONS ====================

def clear_columns(excel_path, columns=['B', 'C']):
    """Clear specified columns in the Excel template."""
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        
        cleared_count = 0
        for col in columns:
            # Clear rows 2-21 (20 chemical components)
            for row_idx in range(2, 22):
                cell = ws[f"{col}{row_idx}"]
                if cell.value is not None:
                    cell.value = None
                    cleared_count += 1
        
        wb.save(excel_path)
        wb.close()
        return True, cleared_count
    except Exception as e:
        return False, str(e)


def create_working_copy():
    """Create a timestamped copy of the template with exact formatting."""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        directory = os.path.dirname(TEMPLATE_PATH)
        filename = os.path.basename(TEMPLATE_PATH)
        name, ext = os.path.splitext(filename)
        
        new_filename = f"{name}_{timestamp}{ext}"
        new_path = os.path.join(directory, new_filename)
        
        # Use shutil.copy2 to preserve all metadata and formatting
        shutil.copy2(TEMPLATE_PATH, new_path)
        
        return new_path
    except Exception as e:
        st.error(f"Error creating working copy: {e}")
        return None


def extract_text_from_image(image):
    """Extract text from uploaded image using optimized OCR."""
    try:
        # Convert to PIL Image if needed
        if not isinstance(image, Image.Image):
            image = Image.open(image)
        
        # Optimize image for OCR
        # Convert to grayscale
        image = image.convert('L')
        
        # Enhance contrast
        from PIL import ImageEnhance
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        
        # Perform OCR with custom config for better accuracy
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(image, config=custom_config)
        
        return text
    except Exception as e:
        st.error(f"OCR Error: {e}")
        return ""


def normalize_component_name(name):
    """Normalize component name for matching."""
    if not name:
        return ""
    return str(name).strip().lower().replace(" ", "").replace("-", "").replace("_", "")


def parse_extracted_text(text):
    """Parse OCR text to extract all 20 component-concentration pairs."""
    lines = text.strip().split('\n')
    
    # Create normalized map of expected components
    expected_map = {normalize_component_name(comp): comp for comp in EXPECTED_COMPONENTS}
    
    # Dictionary to store results - use dict to avoid duplicates
    results_dict = {}
    
    # Multiple patterns to catch different formats
    patterns = [
        r'([A-Za-z\s\(\)\-]+?)[\s\.:]+([0-9,.]+)',  # Component ... 123
        r'([A-Za-z\s\(\)\-]+?)\s+([0-9,.]+)\s*(?:ppm)?',  # Component 123 ppm
        r'([A-Za-z\s\(\)\-]+?)[:\s]+([0-9,.]+)',  # Component: 123
        r'([A-Za-z\s\(\)\-]{3,})\s*([0-9]+\.?[0-9]*)',  # More flexible pattern
    ]
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 3:
            continue
        
        for pattern in patterns:
            matches = re.finditer(pattern, line)
            for match in matches:
                component = match.group(1).strip()
                concentration = match.group(2).strip()
                
                # Skip if component is too short or contains common OCR noise
                if len(component) < 3:
                    continue
                if component.lower() in ['the', 'and', 'for', 'ppm', 'reading', 'st', 'nd', 'rd', 'th', 'min']:
                    continue
                
                # Normalize and try to match
                normalized = normalize_component_name(component)
                
                # Direct match
                if normalized in expected_map:
                    matched_component = expected_map[normalized]
                    if matched_component not in results_dict:
                        results_dict[matched_component] = concentration
                    continue
                
                # Fuzzy match - check if it's a substring or close match
                for exp_norm, exp_orig in expected_map.items():
                    if exp_orig in results_dict:  # Already matched
                        continue
                    
                    # Check various matching strategies
                    match_found = False
                    
                    # Strategy 1: Check if extracted contains expected or vice versa
                    if normalized in exp_norm or exp_norm in normalized:
                        if len(normalized) >= len(exp_norm) * 0.5:  # At least 50% match
                            results_dict[exp_orig] = concentration
                            match_found = True
                            break
                    
                    # Strategy 2: Check word-by-word match for multi-word components
                    component_words = set(component.lower().split())
                    expected_words = set(exp_orig.lower().split())
                    if component_words and expected_words:
                        overlap = len(component_words & expected_words)
                        if overlap > 0 and overlap >= len(expected_words) * 0.5:
                            results_dict[exp_orig] = concentration
                            match_found = True
                            break
                    
                    if match_found:
                        break
    
    # Convert dict to list of dicts for DataFrame
    data = [{"Component": comp, "Concentration": conc} for comp, conc in results_dict.items()]
    
    # If we didn't get all 20, add placeholders for missing ones
    matched_components = set(results_dict.keys())
    missing_components = set(EXPECTED_COMPONENTS) - matched_components
    
    if missing_components:
        st.warning(f"⚠️ {len(missing_components)} component(s) not auto-detected. Please add manually:")
        for comp in sorted(missing_components):
            st.caption(f"   • {comp}")
    
    return pd.DataFrame(data)


def check_column_status(excel_path):
    """Check which columns (B or C) have data."""
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        
        b_has_data = False
        c_has_data = False
        
        # Check rows 2-21 (20 chemicals)
        for row_idx in range(2, 22):
            if ws[f"B{row_idx}"].value is not None:
                b_has_data = True
            if ws[f"C{row_idx}"].value is not None:
                c_has_data = True
            
            if b_has_data and c_has_data:
                break
        
        wb.close()
        return b_has_data, c_has_data
    except Exception as e:
        st.error(f"Error checking columns: {e}")
        return False, False


def populate_column(excel_path, data_df, target_column):
    """Populate specified column with extracted data while preserving ALL formatting including conditional formatting."""
    try:
        # Load workbook without data_only to preserve formulas and conditional formatting
        wb = load_workbook(excel_path)
        ws = wb.active
        
        # Create a mapping of normalized component names to concentrations
        data_map = {}
        for _, row in data_df.iterrows():
            if pd.notna(row["Component"]) and str(row["Component"]).strip():
                key = normalize_component_name(row["Component"])
                value = str(row["Concentration"]).strip()
                data_map[key] = value
        
        updates = 0
        # Match components in column A with our data
        # Starting from row 2 (row 1 is header)
        for row_idx in range(2, 22):  # Rows 2-21 (20 chemicals)
            component_cell = ws[f"A{row_idx}"]
            if component_cell.value:
                # Normalize the component name from Excel
                component_key = normalize_component_name(component_cell.value)
                
                if component_key in data_map:
                    # Get target cell
                    target_cell = ws[f"{target_column}{row_idx}"]
                    
                    # Convert concentration to number if possible (for conditional formatting to work)
                    value = data_map[component_key]
                    try:
                        # Remove any text like "ppm" and commas, then convert to float
                        clean_value = re.sub(r'[^\d.]', '', value)
                        if clean_value:
                            numeric_value = float(clean_value)
                            target_cell.value = numeric_value
                        else:
                            target_cell.value = value
                    except (ValueError, AttributeError):
                        # If not a number, store as string
                        target_cell.value = value
                    
                    updates += 1
        
        # Save the workbook (openpyxl automatically preserves conditional formatting)
        wb.save(excel_path)
        wb.close()
        return True, updates
    except Exception as e:
        return False, str(e)


# ==================== UI ====================

# Section 1: Clear Template
st.header("1️⃣ Clear Template Columns")
col1, col2 = st.columns([3, 1])

with col1:
    st.info(f"📁 Template: `{os.path.basename(TEMPLATE_PATH)}`")

with col2:
    if st.button("🗑️ Clear Columns B & C", type="secondary"):
        if os.path.exists(TEMPLATE_PATH):
            success, result = clear_columns(TEMPLATE_PATH)
            if success:
                st.success(f"✅ Cleared {result} cells from columns B & C")
                st.session_state.upload_count = 0  # Reset upload count
                st.session_state.extracted_data = pd.DataFrame(columns=["Component", "Concentration"])
            else:
                st.error(f"❌ Error: {result}")
        else:
            st.error("❌ Template file not found!")

st.markdown("---")

# Section 2: Upload Image & Auto Extract
st.header("2️⃣ Upload Image (Auto Extract)")

uploaded_file = st.file_uploader(
    "Upload an image with component data",
    type=["png", "jpg", "jpeg", "tiff", "bmp"],
    help="Upload an image - text will be extracted automatically"
)

# Auto-extract when new file is uploaded
if uploaded_file is not None:
    # Check if this is a new upload
    file_id = f"{uploaded_file.name}_{uploaded_file.size}"
    
    if st.session_state.last_uploaded_file != file_id:
        st.session_state.last_uploaded_file = file_id
        
        with st.spinner("🔍 Extracting text from image..."):
            # Extract text
            text = extract_text_from_image(uploaded_file)
            
            if text:
                # Show raw text in expander
                with st.expander("📄 View Raw Extracted Text"):
                    st.text_area("Raw Text:", text, height=150)
                
                # Parse the text
                extracted_df = parse_extracted_text(text)
                
                if not extracted_df.empty:
                    st.session_state.extracted_data = extracted_df
                    st.success(f"✅ Automatically extracted {len(extracted_df)} components")
                else:
                    st.warning("⚠️ No component-concentration pairs found. Please enter data manually below.")
            else:
                st.error("❌ No text extracted from image")

st.markdown("---")

# Section 3: Edit Extracted Data
st.header("3️⃣ Review & Edit Extracted Data")

if not st.session_state.extracted_data.empty:
    st.info("✏️ Review and edit the extracted data if needed")
    
    # Show expected vs extracted
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Expected Components", len(EXPECTED_COMPONENTS))
    with col2:
        st.metric("Extracted Components", len(st.session_state.extracted_data))
    
    # Editable data editor
    edited_df = st.data_editor(
        st.session_state.extracted_data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Component": st.column_config.SelectboxColumn(
                "Component Name",
                options=EXPECTED_COMPONENTS,
                width="medium",
                required=True
            ),
            "Concentration": st.column_config.TextColumn("Concentration Value", width="medium")
        },
        hide_index=True
    )
    
    # Update session state with edited data
    st.session_state.extracted_data = edited_df
else:
    st.warning("📝 No data extracted yet. Upload an image above.")
    
    # Allow manual entry
    if st.button("➕ Add Manual Entry"):
        new_row = pd.DataFrame([{"Component": "", "Concentration": ""}])
        st.session_state.extracted_data = pd.concat(
            [st.session_state.extracted_data, new_row], 
            ignore_index=True
        )
        st.rerun()

st.markdown("---")

# Section 4: Populate Excel
st.header("4️⃣ Populate Excel File")

if not st.session_state.extracted_data.empty:
    # Check current column status
    b_has_data, c_has_data = check_column_status(TEMPLATE_PATH)
    
    # Determine target column
    if not b_has_data:
        target_column = "B"
        st.info("📍 Column B is empty → Data will populate to Column B (1st reading)")
    elif not c_has_data:
        target_column = "C"
        st.info("📍 Column B has data → Data will populate to Column C (2nd reading)")
    else:
        st.warning("⚠️ Both columns B and C have data. Clear the template first or choose manually:")
        target_column = st.radio("Select target column:", ["B", "C"], horizontal=True)
    
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        st.metric("Components Ready", len(st.session_state.extracted_data))
    
    with col2:
        st.metric("Target Column", f"{target_column} ({'1st' if target_column == 'B' else '2nd'} reading)")
    
    with col3:
        if st.button("📥 Populate Excel", type="primary"):
            with st.spinner("Creating new file and populating data..."):
                # Create working copy with exact formatting
                new_file = create_working_copy()
                
                if new_file:
                    # Populate the column
                    success, result = populate_column(
                        new_file, 
                        st.session_state.extracted_data, 
                        target_column
                    )
                    
                    if success:
                        st.session_state.working_file = new_file
                        st.session_state.upload_count += 1
                        st.success(f"✅ Successfully populated {result} rows in Column {target_column}")
                        st.success(f"📄 New file: `{os.path.basename(new_file)}`")
                        
                        # Show file location
                        st.info(f"📂 File saved to: `{new_file}`")
                        st.info("💡 Ready for next upload - upload another image to populate the other column")
                    else:
                        st.error(f"❌ Error populating data: {result}")
else:
    st.warning("⚠️ No data available to populate. Upload an image first.")

# Footer
st.markdown("---")
st.caption("💡 **Note:** Make sure Tesseract OCR is installed on your system")
st.caption("   • Mac: `brew install tesseract`")
st.caption("   • Linux: `sudo apt-get install tesseract-ocr`")
st.caption("   • Windows: Download from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)")