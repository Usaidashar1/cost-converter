import streamlit as st
import tempfile
import os
from pathlib import Path
from convert import convert

st.set_page_config(page_title="Azure Cost Converter", page_icon="☁️", layout="centered")

st.title("☁️ Azure Cost Estimation Converter")
st.write("Upload your Azure Calculator export to mathematically deduct licenses and fetch RI pricing.")

currency = st.selectbox("Select Target Currency:", ["INR", "USD", "EUR", "GBP", "AUD"])
uploaded_file = st.file_uploader("Upload Azure Export (.xlsx)", type=["xlsx"])

MAX_FILE_SIZE_MB = 50

if uploaded_file is not None:
    if uploaded_file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
        st.error(f"File exceeds maximum allowed size of {MAX_FILE_SIZE_MB}MB.")
        st.stop()

    if st.button("Convert File", type="primary"):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / f"input_{uploaded_file.file_id}.xlsx"
            output_path = Path(temp_dir) / f"Processed_Estimate_{currency}.xlsx"
            
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
                    st.error(f"Processing Error: {str(e)}")
                    st.exception(e) # Expose exact traceback for debugging
