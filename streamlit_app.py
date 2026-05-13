import streamlit as st
import pandas as pd
import re

st.title("OPAL-RT Lead Import Normalizer")

uploaded_file = st.file_uploader(
    "Upload CSV or Excel file",
    type=["csv", "xlsx"]
)

def clean_email(email):
    if pd.isna(email):
        return ""
    return str(email).strip().lower()

def is_valid_email(email):
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(pattern, email)

if uploaded_file:

    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Original Data")
    st.dataframe(df)

    # Auto rename columns
    column_mapping = {
        "fname": "First Name",
        "firstname": "First Name",
        "lname": "Last Name",
        "lastname": "Last Name",
        "mail": "Email",
        "email address": "Email",
        "org": "Company",
        "company name": "Company"
    }

    new_columns = {}

    for col in df.columns:
        normalized = col.strip().lower()

        if normalized in column_mapping:
            new_columns[col] = column_mapping[normalized]

    df.rename(columns=new_columns, inplace=True)

    # Clean emails
    if "Email" in df.columns:
        df["Email"] = df["Email"].apply(clean_email)

        df["Valid Email"] = df["Email"].apply(
            lambda x: "Yes" if is_valid_email(x) else "No"
        )

    # Remove duplicates
    if "Email" in df.columns:
        df = df.drop_duplicates(subset=["Email"])

    st.subheader("Cleaned Data")
    st.dataframe(df)

    # Download cleaned CSV
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Download Dynamics-Ready CSV",
        csv,
        "cleaned_leads.csv",
        "text/csv"
    )
