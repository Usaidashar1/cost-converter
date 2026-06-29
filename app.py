import streamlit as st
import os
from convert import convert # This imports your perfectly working script!

st.set_page_config(page_title="Azure Cost Converter", page_icon="☁️")

st.title("☁️ Azure Cost Estimation Converter")
st.write("Upload your Azure Calculator export to mathematically deduct licenses and fetch RI pricing.")

# UI Controls
currency = st.selectbox("Select Target Currency:", ["INR", "USD", "EUR", "GBP", "AUD"])
uploaded_file = st.file_uploader("Upload Azure Export (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Convert File"):
        # Save the uploaded file temporarily
        input_path = "temp_input.xlsx"
        output_path = "Processed_Cost_Estimation.xlsx"
        
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        with st.spinner("Querying Azure Retail Pricing API... Please wait."):
            try:
                # Run your exact script logic
                convert(input_path, output_path, currency)
                
                # Provide the download button
                with open(output_path, "rb") as f:
                    st.success("Conversion Complete!")
                    st.download_button(
                        label="📥 Download Processed Estimate",
                        data=f,
                        file_name=f"Processed_Estimate_{currency}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"An error occurred: {e}")
