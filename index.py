import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Excel to PDF Converter", page_icon="📄", layout="wide")

# Custom CSS
st.markdown(
    """
    <style>
    .stApp {
        background-color: #f0f2f6;
        color: #000000;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Title and description
st.title("Excel to PDF Converter")
st.write("This app allows you to convert Excel files to PDF format.")

# Upload excel file
uploaded_files = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"], accept_multiple_files=True)

# Convert to PDF
if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file)
        else:
            st.error(f"Unsupported file type {file_ext}")
            continue

        # File details
        st.write(f"**Preview of the file: {file.name}**")
        st.dataframe(df.head())

        # Data cleaning options
        st.subheader("Data Cleaning Options")
        if st.checkbox(f"Remove duplicates from {file.name}"):
            if st.button(f"Apply duplicate removal for {file.name}"):
                df.drop_duplicates(inplace=True)
                st.success("Duplicates removed successfully")

        if st.checkbox(f"Fill missing values for {file.name}"):
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            df[numeric_cols] = df[numeric_cols].apply(lambda x: x.fillna(x.mean()))
            st.success("Missing values filled successfully")

        # Data transformation
        st.subheader("Data Transformation Options")
        columns = st.multiselect(f"Select columns to normalize in {file.name}", df.columns.tolist())

        if columns:
            for col in columns:
                if st.checkbox(f"Normalize {col} in {file.name}"):
                    df[col] = (df[col] - df[col].min()) / (df[col].max() - df[col].min())
                    st.success(f"{col} normalized successfully")

        # Data visualization
        st.subheader("Data Visualization")
        if st.checkbox(f"Show summary statistics for {file.name}"):
            st.bar_chart(df.select_dtypes(include=['number']).iloc[:, :2])

        # Conversion options
        st.subheader("Conversion Options")
        conversion_type = st.radio(f"Convert {file.name} to:", ['CSV', 'Excel'], key=file.name)

        if st.button(f"Convert {file.name} to {conversion_type}"):
            output = BytesIO()
            if conversion_type == 'CSV':
                df.to_csv(output, index=False)
                output.seek(0)
                st.download_button(label="Download CSV", data=output, file_name=f"{file.name}.csv", mime="text/csv")

            elif conversion_type == 'Excel':
                with BytesIO() as output:
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    df.to_excel(writer, index=False, sheet_name="Sheet1")
                    writer.close()
                    output.seek(0)
                    st.download_button(label="Download Excel", data=output, file_name=f"{file.name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.success("Conversion successful!")
