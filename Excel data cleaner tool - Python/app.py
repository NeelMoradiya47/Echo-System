import streamlit as st
import pandas as pd
import tempfile

# Import your cleaner functions
from cleaner import (
    standardize_columns,
    remove_empty_rows,
    remove_duplicates
)

st.set_page_config(page_title="Excel Cleaner", layout="wide")

st.title("Excel Data Cleaner Tool")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Original Data")
    st.dataframe(df.head())

    # Apply your actual cleaning pipeline
    df = standardize_columns(df)
    df = remove_empty_rows(df)
    df = remove_duplicates(df)

    st.subheader("Cleaned Data")
    st.dataframe(df.head())

    # Download option
    csv = df.to_csv(index=False).encode('utf-8')

    st.download_button(
        label="Download Cleaned File",
        data=csv,
        file_name='cleaned_data.csv',
        mime='text/csv'
    )
