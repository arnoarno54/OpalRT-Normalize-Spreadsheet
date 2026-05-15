import streamlit as st
import pandas as pd
import re
from datetime import datetime

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="OPAL-RT Dynamics Lead Import Normalizer",
    layout="wide"
)

# =====================================================
# CUSTOM CSS / BRANDING
# =====================================================

st.markdown("""
<style>

.main {
    background-color: #f5f7fb;
}

h1 {
    color: #002B5C;
    font-weight: 700;
}

h2, h3 {
    color: #002B5C;
}

.stButton>button {
    background-color: #00A3E0;
    color: white;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}

.stDownloadButton>button {
    background-color: #00A3E0;
    color: white;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}

.block-container {
    padding-top: 2rem;
}

.required-box {
    background-color: #ffffff;
    border-left: 5px solid #00A3E0;
    padding: 15px;
    border-radius: 8px;
    margin-bottom: 25px;
}

.validation-box {
    background-color: #ffffff;
    border-left: 5px solid #002B5C;
    padding: 15px;
    border-radius: 8px;
    margin-bottom: 20px;
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# HEADER
# =====================================================

col1, col2 = st.columns([1, 4])

with col1:
    st.image(
        "https://www.opal-rt.com/wp-content/uploads/2022/09/OPALRT-logo.png",
        width=220
    )

with col2:
    st.title("Dynamics Lead Import Normalizer")
    st.markdown(
        "Prepare CRM-ready lead imports for Microsoft Dynamics"
    )

st.divider()

# =====================================================
# REQUIRED FIELD INFO
# =====================================================

st.markdown("""
<div class="required-box">

<h3>Mandatory Lead Fields</h3>

The following fields are required by the OPAL-RT Dynamics import process:

<ul>
<li><b>Subject</b></li>
<li><b>First Name</b></li>
<li><b>Last Name</b></li>
<li><b>Email</b></li>
<li><b>Company</b></li>
<li><b>Country/Region</b></li>
<li><b>State/Province</b></li>
<li><b>Market Segment</b></li>
<li><b>Main Application</b></li>
</ul>

These validations are automatically checked before export.

</div>
""", unsafe_allow_html=True)

# =====================================================
# GLOBAL IMPORT SETTINGS
# =====================================================

st.header("Global Import Settings")

st.write(
    "These values will be automatically applied to all imported rows."
)

colA, colB, colC = st.columns(3)

with colA:

    subject = st.text_input(
        "Subject *",
        value=f"{datetime.now().strftime('%Y%m')}Prospection"
    )

    lead_source = st.selectbox(
        "Lead Source",
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
        "Rating",
        [
            "Cold",
            "Warm",
            "Hot"
        ],
        index=0
    )

    allow_marketing = st.selectbox(
        "Allow Marketing Communication",
        [
            "Yes",
            "No"
        ],
        index=0
    )

with colC:

    source_campaign = st.text_input(
        "Source Campaign"
    )

    description = st.text_area(
        "Description",
        height=120
    )

st.divider()

# =====================================================
# FILE UPLOAD
# =====================================================

uploaded_file = st.file_uploader(
    "Upload CSV or Excel File",
    type=["csv", "xlsx"]
)

# =====================================================
# CONFIG
# =====================================================

REQUIRED_COLUMNS = [
    "Subject",
    "First Name",
    "Last Name",
    "Email",
    "Company Name",
    "Country",
    "State or Province",
    "Market Segment",
    "Main Application"
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
    "country/region": "Country",

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

# =====================================================
# HELPERS
# =====================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    value = str(value)
    value = value.strip()
    value = re.sub(r"\\s+", " ", value)

    return value


def clean_email(email):

    if pd.isna(email):
        return ""

    return str(email).strip().lower()


def is_valid_email(email):

    if email == "":
        return False

    pattern = r'^[\\w\\.-]+@[\\w\\.-]+\\.\\w+$'

    return re.match(pattern, email)


def normalize_country(country):

    if pd.isna(country):
        return ""

    cleaned = str(country).strip()

    lowered = cleaned.lower()

    if lowered in COUNTRY_NORMALIZATION:
        return COUNTRY_NORMALIZATION[lowered]

    return cleaned

# =====================================================
# PROCESS FILE
# =====================================================

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)

    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Original Uploaded Data")

    st.dataframe(df)

    # REMOVE GHOST COLUMNS
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # STANDARDIZE COLUMN NAMES
    new_columns = {}

    for col in df.columns:

        normalized = str(col).strip().lower()

        if normalized in COLUMN_MAPPING:
            new_columns[col] = COLUMN_MAPPING[normalized]

    df.rename(columns=new_columns, inplace=True)

    # CLEAN TEXT
    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    # EMAIL
    if "Email" in df.columns:
        df["Email"] = df["Email"].apply(clean_email)

    # COUNTRY
    if "Country" in df.columns:
        df["Country"] = df["Country"].apply(normalize_country)

    # CREATE MISSING COLUMNS
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

    # =====================================================
    # VALIDATION
    # =====================================================

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

        if len(str(row["Company Name"])) > 100:

            validation_errors.append(
                f"Row {index + 2}: Company Name exceeds 100 characters"
            )

    # REMOVE DUPLICATES
    if "Email" in df.columns:

        df = df.drop_duplicates(subset=["Email"])

    # FINAL COLUMN ORDER
    df = df[FINAL_COLUMNS]

    st.divider()

    # =====================================================
    # RESULTS
    # =====================================================

    st.subheader("Dynamics-Ready Import File")

    st.dataframe(df)

    # =====================================================
    # VALIDATION REPORT
    # =====================================================

    st.subheader("Validation Report")

    if len(validation_errors) == 0:

        st.success(
            "No validation issues detected. File is ready for Dynamics import."
        )

    else:

        st.markdown("""
        <div class="validation-box">
        <h4>Validation Issues Found</h4>
        </div>
        """, unsafe_allow_html=True)

        for error in validation_errors:
            st.warning(error)

    # =====================================================
    # EXPORT
    # =====================================================

    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Dynamics-Ready CSV",
        data=csv,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv"
    )
