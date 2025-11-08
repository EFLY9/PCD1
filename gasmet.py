import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image, ImageEnhance, ImageOps
import pytesseract
import re
from datetime import datetime
import io
import shutil
import cv2

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

st.set_page_config(
    page_title="Excel OCR Populator",
    page_icon="https://www.sgcleaningxpert.com/wp-content/uploads/2017/10/NEA-Logo1.png",
    layout="wide"
)

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


# ==================== IMAGE CAPTURE, ENHANCEMENT & PREPROCESSING ====================

# Image Enhancement (Step 1)
def preprocess_image(image):
    """Enhance image for better OCR results."""
    image = image.convert('L')  # Convert to grayscale
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)  # Increase contrast
    image = ImageOps.autocontrast(image)  # Auto contrast to improve image quality
    return image

# OpenCV Preprocessing (Step 2)
def preprocess_with_opencv(image):
    """Preprocess image using OpenCV techniques like denoising and thresholding."""
    image_np = np.array(image)  # Convert PIL image to numpy array
    image_blurred = cv2.GaussianBlur(image_np, (5, 5), 0)  # Apply Gaussian blur to reduce noise
    _, image_thresh = cv2.threshold(image_blurred, 150, 255, cv2.THRESH_BINARY)  # Apply thresholding
    return image_thresh

# Streamlit UI for camera input (Image capture)
captured_image = st.camera_input("Capture an image")

if captured_image is not None:
    # Open the captured image using PIL
    image = Image.open(captured_image)

    # Step 1: Enhance the image quality
    enhanced_image = preprocess_image(image)

    # Step 2: Preprocess with OpenCV (thresholding & denoising)
    opencv_image = preprocess_with_opencv(enhanced_image)

    # Display the processed image for OCR
    st.image(opencv_image, caption="Processed Image", use_column_width=True)

    # Step 3: Perform OCR (using Tesseract for demonstration)
    with st.spinner("🔍 Extracting text from image..."):
        text = extract_text_from_image(opencv_image)
        if text:
            st.text_area("Extracted Text", text, height=150)

            # Parse and display extracted components
            extracted_df = parse_extracted_text(text)
            st.dataframe(extracted_df)
