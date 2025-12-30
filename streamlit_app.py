"""
Streamlit UI for Fund Analysis Excel Processor

A simple web interface to upload Excel files, process them, and download the results.
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import os
from fund_analysis_processor import FundAnalysisProcessor
import logging

# Configure page
st.set_page_config(
    page_title="Fund Analysis Processor",
    page_icon="📊",
    layout="wide"
)

# Suppress unnecessary logs
logging.getLogger().setLevel(logging.WARNING)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    """Main Streamlit application."""
    
    # Header
    st.markdown('<div class="main-header">📊 Fund Analysis Excel Processor</div>', unsafe_allow_html=True)
    
    # Sidebar with instructions
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        1. **Upload Excel File**: Click "Browse files" to select your input Excel file
        2. **Required Columns**: Ensure your file has these columns:
           - Scheme Code, Scheme Name, Month, Month End
           - Instrument Name, Holding (%), Instrument Sector
           - Instrument SEBI Mcap, Instrument SEBI Mcap Type
           - NSE Symbol, Price
        3. **Process**: Click "Process File" to calculate derived columns and generate pivot tables
        4. **Download**: Use the download button to get your processed file
        
        **Output includes:**
        - Processed Data with calculated columns
        - Company Pivot Table
        - Sector Pivot Table
        - Market Cap Pivot Table
        """)
        
        st.markdown("---")
        st.markdown("### 📊 About")
        st.markdown("""
        This tool processes fund analysis data and calculates:
        - Start Price (previous month's price)
        - Monthly Stock Return %
        - Start wt% (previous month's holding)
        - Stock Monthly Contribution %
        """)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 Upload Excel File")
        
        # File uploader
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload your fund analysis Excel file with required columns"
        )
    
    with col2:
        st.header("⚙️ Processing")
        
        # Process button
        process_button = st.button(
            "🔄 Process File",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None
        )
    
    # Initialize session state
    if 'output_file_path' not in st.session_state:
        st.session_state.output_file_path = None
    if 'processing_error' not in st.session_state:
        st.session_state.processing_error = None
    if 'processing_success' not in st.session_state:
        st.session_state.processing_success = False
    
    # Process file when button is clicked
    if process_button and uploaded_file is not None:
        with st.spinner("Processing file... This may take a moment."):
            try:
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Save uploaded file temporarily
                    input_path = Path(temp_dir) / uploaded_file.name
                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Generate output path
                    output_path = Path(temp_dir) / f"{input_path.stem}_processed.xlsx"
                    
                    # Initialize and process
                    processor = FundAnalysisProcessor(
                        input_file_path=str(input_path),
                        output_file_path=str(output_path)
                    )
                    
                    # Run processing
                    result_path = processor.process()
                    
                    # Read output file to session state
                    with open(result_path, "rb") as f:
                        st.session_state.output_file_bytes = f.read()
                    
                    st.session_state.output_file_path = result_path
                    st.session_state.output_file_name = f"{input_path.stem}_processed.xlsx"
                    st.session_state.processing_success = True
                    st.session_state.processing_error = None
                    
                    st.success("✅ File processed successfully!")
                    
            except FileNotFoundError as e:
                st.session_state.processing_error = f"File not found: {str(e)}"
                st.session_state.processing_success = False
                st.error(f"❌ Error: {st.session_state.processing_error}")
                
            except ValueError as e:
                st.session_state.processing_error = f"Validation error: {str(e)}"
                st.session_state.processing_success = False
                st.error(f"❌ Error: {st.session_state.processing_error}")
                
            except Exception as e:
                st.session_state.processing_error = f"Processing error: {str(e)}"
                st.session_state.processing_success = False
                st.error(f"❌ Error: {st.session_state.processing_error}")
    
    # Display file info if uploaded
    if uploaded_file is not None:
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📄 File Name", uploaded_file.name)
        
        with col2:
            file_size = len(uploaded_file.getvalue()) / 1024  # KB
            st.metric("📦 File Size", f"{file_size:.2f} KB")
        
        with col3:
            st.metric("📅 Upload Status", "✅ Ready" if uploaded_file else "⏳ Waiting")
    
    # Show processing results
    if st.session_state.processing_success and 'output_file_bytes' in st.session_state:
        st.markdown("---")
        st.markdown('<div class="success-box">', unsafe_allow_html=True)
        st.success("🎉 Processing Complete!")
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Download button
        st.download_button(
            label="📥 Download Processed File",
            data=st.session_state.output_file_bytes,
            file_name=st.session_state.output_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        
        # Show file preview info
        with st.expander("📊 View Output File Information"):
            try:
                # Read the output file to show sheet names
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(st.session_state.output_file_bytes)
                    tmp_path = tmp_file.name
                
                try:
                    excel_file = pd.ExcelFile(tmp_path, engine='openpyxl')
                    st.write("**Output File Contains:**")
                    for i, sheet_name in enumerate(excel_file.sheet_names, 1):
                        df = pd.read_excel(tmp_path, sheet_name=sheet_name, engine='openpyxl')
                        st.write(f"{i}. **{sheet_name}** - {len(df)} rows × {len(df.columns)} columns")
                    excel_file.close()
                finally:
                    # Clean up temp file
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                        
            except Exception as e:
                st.warning(f"Could not preview file details: {str(e)}")
    
    # Show error if processing failed
    if st.session_state.processing_error and not st.session_state.processing_success:
        st.markdown("---")
        st.error(f"❌ Processing Failed: {st.session_state.processing_error}")
        st.info("💡 Please check that your file has all required columns and is in the correct format.")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666; padding: 1rem;'>"
        "Fund Analysis Excel Processor | Built with Streamlit"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()

