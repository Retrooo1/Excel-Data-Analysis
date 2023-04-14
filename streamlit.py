import streamlit as st
import pandas as pd
from main import run
import os
st.title("Excel Processor")
st.text("Testing1")
st.caption("Testing1")
# File uploader
uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")

if uploaded_file is not None:
    # Read the uploaded file
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    print("df ", df)
    # Process the data using the 'run' function from main.py
    run(df)

    # Create a download link for the output file
    # output_file = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')
    # processed_data.to_excel(output_file, index=False)
    # output_file.save()
    # Create a download link for the output file
    with open("Output101.xlsx", "rb") as excel_file:
        data = excel_file.read()
        st.download_button("Download output file", data, file_name="Output101.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Show a success message
    st.success("Output file generated successfully!")
