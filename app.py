import streamlit as st
import os
import tempfile
from pathlib import Path
from convert import convert 

st.set_page_config(page_title="Azure Cost Converter", page_icon="☁️")

st.title("☁️ Azure Cost Estimation Converter")
st.write("Upload your Azure Calculator export to mathematically deduct licenses and fetch RI pricing.")

currency = st.selectbox("Select Target Currency:", ["INR", "USD", "EUR", "GBP", "AUD"])
uploaded_file = st.file_uploader("Upload Azure Export (.xlsx)", type=["xlsx"])

MAX_FILE_SIZE_MB = 50

if uploaded_file is not None:
    # Security: File Size Limit Check
    if uploaded_file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
        st.error(f"File exceeds maximum allowed size of {MAX_FILE_SIZE_MB}MB.")
        st.stop()

    if st.button("Convert File"):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "temp_input.xlsx"
            output_path = Path(temp_dir) / "Processed_Cost_Estimation.xlsx"
            
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            with st.spinner(f"Querying Azure Retail Pricing API in {currency}... Please wait."):
                try:
                    convert(str(input_path), str(output_path), currency)
                    
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                        
                    st.success("Conversion Complete!")
                    st.download_button(
                        label="📥 Download Processed Estimate",
                        data=file_data,
                        file_name=f"Processed_Estimate_{currency}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    # Security: Do not expose raw internal Python exceptions to the UI
                    st.error("An error occurred during file processing. Please ensure the file is an unmodified Azure Calculator export.")
                    # Optionally log 'e' to a backend monitoring service here
