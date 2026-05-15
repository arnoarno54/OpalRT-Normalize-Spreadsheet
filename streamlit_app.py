# streamlit_app.py

import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.set_page_config(page_title="OPAL-RT Lead Import Normalizer", layout="wide")

# =========================
# OPAL-RT HEADER
# =========================

col1, col2 = st.columns([1, 4])

with col1:
    st.image(
        "https://www.opal-rt.com/wp-content/uploads/2022/09/opalrt-logo-black.png",
        width=220
    )

with col2:
    st.title("Dynamics Lead Import Normalizer")
    st.caption("Prepare CRM-ready lead imports for Microsoft Dynamics")

st.divider()

# =========================
# GLOBAL IMPORT SETTINGS
# =========================

st.subheader("Global Import Settings")
st.write("These values will be applied to all imported rows.")

colA, colB, colC = st.columns(3)

with colA:
    subject = st.text_input(
        "Subject *",
        value=f"{datetime.now().strftime('%Y%m')}Prospection"
    )

    lead_source = st.selectbox(
        "Lead Source *",
        [
            "Shows",
            "Web",
            "Prospection",
            "Webinar",
            "Referral",
            "Social Media",
            "Customer Portal",
            "SPS",
            "Others"
        ],
        index=2
    )

with colB:
    rating = st.selectbox(
        "Rating *",
        ["Cold", "Warm", "Hot"],
        index=0
    )

    allow_marketing = st.selectbox(
        "Allow Marketing Communication *",
        ["Yes", "No"],
        index=0
    )

with colC:
    source_campaign = st.text_input(
        "Source Campaign",
        value=""
    )

    description = st.text_area(
        "Description",
        height=120
    )

st.divider()

# =========================
# FILE UPLOAD
# =========================

uploaded_file = st.file_uploader(
    "Upload CSV or Excel File",
    type=["csv", "xlsx"]
)

# =========================
# HELPERS
# =========================

REQUIRED_COLUMNS = [
    "Subject",
    "First Name",
    "Last Name",
    "Job Title",
    "Company Name",
    "Country"
]

FINAL_COLUMNS = [
    "(Do Not Modify) Lead",
    "(Do Not Modify) Row Checksum",
    "(Do Not Modify) Modified On",
    "Subject",
    "First Name",
    "Last Name",
    "Job Title",
    "Company Name",
    "Email",
    "Business Phone",
    "Country",
    "State or Province",
    "Description",
    "Lead Source",
    "Rating",
    "Source Campaign",
    "Market Segment",
    "Main Application",
    "Industry Sector",
    "LinkedIn",
    "Allow Marketing Communication"
]

COLUMN_MAPPING = {
    "firstname": "First Name",
    "first name": "First Name",
    "fname": "First Name",
    "lastname": "Last Name",
    "last name": "Last Name",
    "surname": "Last Name",
    "lname": "Last Name",
    "jobtitle": "Job Title",
    "title": "Job Title",
    "company": "Company Name",
    "organization": "Company Name",
    "org": "Company Name",
    "email": "Email",
    "mail": "Email",
    "phone": "Business Phone",
    "telephone": "Business Phone",
    "mobile": "Business Phone",
    "country": "Country",
    "state": "State or Province",
    "province": "State or Province",
    "linkedin": "LinkedIn",
    "linkedin_url": "LinkedIn"
}

COUNTRY_NORMALIZATION = {
    "usa": "United States",
    "us": "United States",
    "u.s.a": "United States",
    "uk": "United Kingdom",
    "uae": "United Arab Emirates",
    "qc": "Quebec"
}

VALID_LEAD_SOURCES = [
    "Shows",
    "Web",
    "Prospection",
    "Webinar",
    "Referral",
    "Social Media",
    "Customer Portal",
    "SPS",
    "Others"
]

VALID_RATINGS = ["Cold", "Warm", "Hot"]
VALID_MARKETING = ["Yes", "No"]


def clean_text(value):
    if pd.isna(value):
        return ""

    value = str(value)
    value = value.strip()
    value = re.sub(r"\s+", " ", value)

    return value


def clean_email(email):
    if pd.isna(email):
        return ""

    return str(email).strip().lower()


def is_valid_email(email):
    if email == "":
        return True

    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email)


def normalize_country(country):
    if pd.isna(country):
        return ""

    cleaned = str(country).strip()
    lowered = cleaned.lower()

    if lowered in COUNTRY_NORMALIZATION:
        return COUNTRY_NORMALIZATION[lowered]

    return cleaned


# =========================
# PROCESS FILE
# =========================

if uploaded_file:

    # READ FILE
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Original Uploaded Data")
    st.dataframe(df)

    # REMOVE EMPTY/GHOST COLUMNS
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # STANDARDIZE HEADERS
    new_columns = {}

    for col in df.columns:
        normalized = str(col).strip().lower()

        if normalized in COLUMN_MAPPING:
            new_columns[col] = COLUMN_MAPPING[normalized]

    df.rename(columns=new_columns, inplace=True)

    # CLEAN TEXT
    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    # EMAIL CLEANUP
    if "Email" in df.columns:
        df["Email"] = df["Email"].apply(clean_email)

    # COUNTRY CLEANUP
    if "Country" in df.columns:
        df["Country"] = df["Country"].apply(normalize_country)

    # ADD MISSING COLUMNS
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # APPLY GLOBAL VALUES
    df["Subject"] = subject
    df["Description"] = description
    df["Lead Source"] = lead_source
    df["Rating"] = rating
    df["Source Campaign"] = source_campaign
    df["Allow Marketing Communication"] = allow_marketing

    # VALIDATION
    validation_errors = []

    for index, row in df.iterrows():

        # REQUIRED FIELDS
        for field in REQUIRED_COLUMNS:
            if str(row[field]).strip() == "":
                validation_errors.append(
                    f"Row {index + 2}: Missing required field -> {field}"
                )

        # EMAIL VALIDATION
        if not is_valid_email(row["Email"]):
            validation_errors.append(
                f"Row {index + 2}: Invalid email -> {row['Email']}"
            )

        # FIELD LENGTHS
        if len(str(row["First Name"])) > 50:
            validation_errors.append(
                f"Row {index + 2}: First Name exceeds 50 characters"
            )

        if len(str(row["Last Name"])) > 50:
            validation_errors.append(
                f"Row {index + 2}: Last Name exceeds 50 characters"
            )

        if len(str(row["Job Title"])) > 100:
            validation_errors.append(
                f"Row {index + 2}: Job Title exceeds 100 characters"
            )

        if len(str(row["Company Name"])) > 100:
            validation_errors.append(
                f"Row {index + 2}: Company Name exceeds 100 characters"
            )

    # REMOVE DUPLICATES
    if "Email" in df.columns:
        df = df.drop_duplicates(subset=["Email"])

    # FINAL COLUMN ORDER
    df = df[FINAL_COLUMNS]

    # RESULTS
    st.subheader("Dynamics-Ready Lead File")
    st.dataframe(df)

    # VALIDATION REPORT
    st.subheader("Validation Report")

    if len(validation_errors) == 0:
        st.success("No validation issues detected.")
    else:
        for error in validation_errors:
            st.warning(error)

    # EXPORT
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Dynamics-Ready CSV",
        data=csv,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv"
    )
