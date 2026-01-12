#!/usr/bin/env python3
"""
JPL Report Generator - Web Application

A Streamlit web app for generating JPL reports from Excel data and Word templates.
"""

import streamlit as st
import os
import tempfile
import shutil
import zipfile
from io import BytesIO

# Import the fill_report function
from fill_report import fill_report

st.set_page_config(
    page_title="JPL Report Generator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 JPL Report Generator")
st.markdown("---")

# Create columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Excel File")
    excel_file = st.file_uploader(
        "Select Excel file with sample data",
        type=["xlsx"],
        help="Excel file containing 'To Word 1' and 'To Word 2' tabs"
    )

with col2:
    st.subheader("2. Upload Word Template")
    template_file = st.file_uploader(
        "Select Word template",
        type=["docx"],
        help="Word document template to fill"
    )

st.markdown("---")

# Particle data upload
st.subheader("3. Upload Particle Data (Optional)")
st.markdown("Upload a ZIP file containing all `*_processed` folders with `summary_export.csv` files")
particle_zip = st.file_uploader(
    "Select ZIP file with particle data folders",
    type=["zip"],
    help="ZIP file containing folders like 'JPL25-0172_Abx_BS2003ES01_..._processed' with summary_export.csv"
)

st.markdown("---")

# Output filename
st.subheader("4. Output Filename")
output_filename = st.text_input(
    "Enter output filename",
    value="JPL_Report_Filled.docx",
    help="This will also be used as the footer filename"
)

# Ensure .docx extension
if output_filename and not output_filename.endswith('.docx'):
    output_filename += '.docx'

st.markdown("---")

# Generate button
if st.button("🚀 Generate Report", type="primary", use_container_width=True):
    
    if not excel_file:
        st.error("Please upload an Excel file")
    elif not template_file:
        st.error("Please upload a Word template")
    elif not output_filename:
        st.error("Please enter an output filename")
    else:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                with st.spinner("Generating report..."):
                    
                    # Save uploaded Excel file
                    excel_path = os.path.join(temp_dir, "data.xlsx")
                    with open(excel_path, "wb") as f:
                        f.write(excel_file.getbuffer())
                    
                    # Save uploaded template
                    template_path = os.path.join(temp_dir, "template.docx")
                    with open(template_path, "wb") as f:
                        f.write(template_file.getbuffer())
                    
                    # Output path
                    output_path = os.path.join(temp_dir, output_filename)
                    
                    # Handle particle data
                    particle_data_path = None
                    if particle_zip:
                        # Extract ZIP file
                        particle_dir = os.path.join(temp_dir, "particle_data")
                        os.makedirs(particle_dir, exist_ok=True)
                        
                        with zipfile.ZipFile(BytesIO(particle_zip.getbuffer()), 'r') as z:
                            z.extractall(particle_dir)
                        
                        particle_data_path = particle_dir
                        st.info(f"Extracted particle data to temporary folder")
                    
                    # Progress display
                    progress_placeholder = st.empty()
                    log_placeholder = st.empty()
                    
                    # Capture output
                    import io
                    import sys
                    
                    # Redirect stdout to capture logs
                    old_stdout = sys.stdout
                    sys.stdout = log_capture = io.StringIO()
                    
                    try:
                        # Run the report generator
                        success = fill_report(
                            excel_path=excel_path,
                            template_path=template_path,
                            output_path=output_path,
                            particle_data_path=particle_data_path,
                            reorder_images=True,
                            footer_filename=output_filename
                        )
                    finally:
                        # Restore stdout
                        sys.stdout = old_stdout
                    
                    # Show logs
                    log_output = log_capture.getvalue()
                    with st.expander("📋 Processing Log", expanded=False):
                        st.code(log_output, language="text")
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ Report generated successfully!")
                        
                        # Read the generated file
                        with open(output_path, "rb") as f:
                            report_data = f.read()
                        
                        # Download button
                        st.download_button(
                            label="📥 Download Report",
                            data=report_data,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.error("❌ Failed to generate report. Check the log for details.")
                        
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                with st.expander("Error Details"):
                    st.code(traceback.format_exc())

# Instructions
st.markdown("---")
with st.expander("📖 Instructions"):
    st.markdown("""
    ### How to use this app:
    
    1. **Upload Excel File**: Select your Excel file containing sample data
       - Must have 'To Word 1' and 'To Word 2' tabs
    
    2. **Upload Word Template**: Select the JPL report template (.docx)
    
    3. **Upload Particle Data (Optional)**: 
       - Create a ZIP file containing all `*_processed` folders
       - Each folder should contain a `summary_export.csv` file
       - The folder names should match the batch names in your Excel
    
    4. **Enter Output Filename**: This will be used for both the output file and footer
    
    5. **Click Generate Report**: The app will process your data and generate the report
    
    6. **Download**: Click the download button to save your filled report
    
    ### Features:
    - ✅ Fills Sample Information Table
    - ✅ Fills FlowCam Table (Table 1)
    - ✅ Fills particle concentration and count data
    - ✅ Handles Orientation dropdowns
    - ✅ Handles EP/Seidenader selection
    - ✅ Reorders morphology images by sample sequence
    - ✅ Updates footer filename and removes highlight
    """)
