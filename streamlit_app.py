# =========================================================
# DATA-CENTRES FILE MAPPING
# =========================================================
#
# Source File Columns Detected:
#
# Company Name        -> Company Name
# First Name          -> First Name
# Last Name           -> Last Name
# Job Title           -> Job Title
# Location            -> Country + State/Province
# LinkedIn Profile    -> LinkedIn
# Work Email          -> Email
# Mobile Phone        -> Business Phone
#
# This updated version automatically maps those fields.
#
# =========================================================

import streamlit as st
import pandas as pd
import re
from datetime import datetime

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Opal RT Spreadsheet Cleaner",
    layout="wide"
)

# =========================================================
# STYLING
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #F4F7FB;
}

.main .block-container {
    padding-top: 1.5rem;
    max-width: 1450px;
}

.hero-container {
    position: relative;
    border-radius: 20px;
    overflow: hidden;
    margin-bottom: 35px;
}

.hero-overlay {
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: linear-gradient(
        90deg,
        rgba(0,32,91,0.92) 0%,
        rgba(0,32,91,0.75) 40%,
        rgba(0,32,91,0.35) 100%
    );
}

.hero-content {
    position: absolute;
    top: 50%;
    left: 60px;
    transform: translateY(-50%);
    color: white;
    z-index: 2;
}

.hero-title {
    font-size: 52px;
    font-weight: 700;
    margin-bottom: 12px;
}

.hero-subtitle {
    font-size: 20px;
    opacity: 0.95;
}

.section-title {
    color: #071B4D;
    font-size: 34px;
    font-weight: 700;
    margin-bottom: 10px;
}

.stButton>button,
.stDownloadButton>button {
    background-color: #00AEEF;
    color: white;
    border-radius: 10px;
    border: none;
    padding: 12px 24px;
    font-weight: 600;
}

.validation-good {
    background: #DFF6E5;
    color: #146C2E;
    padding: 18px;
    border-radius: 10px;
    margin-top: 15px;
    font-weight: 600;
}

.validation-error {
    background: #FFE3E3;
    color: #B42318;
    padding: 18px;
    border-radius: 10px;
    margin-top: 15px;
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# HERO IMAGE
# =========================================================

st.markdown("""
<div class="hero-container">

<img src="https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg"
style="width:100%; height:420px; object-fit:cover;">

<div class="hero-overlay"></div>

<div class="hero-content">

<div class="hero-title">
Opal RT Spreadsheet Cleaner
</div>

<div class="hero-subtitle">
Prepare CRM-ready lead imports for Microsoft Dynamics
</div>

</div>

</div>
""", unsafe_allow_html=True)

# =========================================================
# DROPDOWNS
# =========================================================

MARKET_SEGMENT_APPLICATIONS = {

    "": [""],

    "Aerospace": [
        "",
        "Autonomous Systems (Aero)",
        "Avionics System",
        "Electrical Actuators and Servos",
        "EVTOL",
        "More Electrical Aircraft",
        "Onboard System"
    ],

    "Automotive": [
        "",
        "Charging",
        "EV/HEV Powertrain",
        "Full Vehicle Simulation",
        "ICE Powertrain"
    ],

    "Energy Conversion": [
        "",
        "Backup Power (UPS)",
        "Inverter/Converter",
        "Medium and Large Drive (>150KW)"
    ],

    "Marine, Railway, Off-Highway": [
        "",
        "BMS Control",
        "Grid Infrastructure",
        "Onboard Power System",
        "Propulsion Control"
    ],

    "Power System": [
        "",
        "Conventional Generation",
        "Converter-Based Energy Resource",
        "Distribution",
        "FACTS & HVDC",
        "Microgrid",
        "Substation",
        "Transmission"
    ]
}

INDUSTRY_SECTORS = [
    "",
    "Academic - Research or Post-graduate",
    "Academic - Undergraduate",
    "Consulting & Engineering Firm",
    "Defense",
    "Electrical Utility",
    "Manufacturer",
    "Other",
    "Research Lab - Industrial & Gov.",
    "Stock - Inventory"
]

# =========================================================
# IMPORT SETTINGS
# =========================================================

st.markdown(
    '<div class="section-title">Global Import Settings</div>',
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

with col1:

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

    market_segment = st.selectbox(
        "Market Segment",
        list(MARKET_SEGMENT_APPLICATIONS.keys())
    )

with col2:

    rating = st.selectbox(
        "Rating",
        ["Cold", "Warm", "Hot"],
        index=0
    )

    allow_marketing = st.selectbox(
        "Allow Marketing Communication",
        ["Yes", "No"],
        index=0
    )

    main_application = st.selectbox(
        "Main Application",
        MARKET_SEGMENT_APPLICATIONS[market_segment]
    )

with col3:

    industry_sector = st.selectbox(
        "Industry Sector",
        INDUSTRY_SECTORS
    )

    source_campaign = st.text_input(
        "Source Campaign"
    )

    description = st.text_area(
        "Description",
        height=120
    )

st.divider()

# =========================================================
# FILE UPLOAD
# =========================================================

uploaded_file = st.file_uploader(
    "Upload CSV or Excel File",
    type=["csv", "xlsx"]
)

# =========================================================
# FINAL COLUMNS
# =========================================================

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

# =========================================================
# COLUMN MAPPING
# =========================================================

COLUMN_MAPPING = {

    "first name": "First Name",
    "firstname": "First Name",
    "fname": "First Name",

    "last name": "Last Name",
    "lastname": "Last Name",
    "lname": "Last Name",

    "company": "Company Name",
    "company name": "Company Name",
    "organization": "Company Name",

    "job title": "Job Title",
    "title": "Job Title",

    "work email": "Email",
    "business email": "Email",
    "email": "Email",
    "email address": "Email",

    "mobile phone": "Business Phone",
    "phone": "Business Phone",
    "telephone": "Business Phone",

    "linkedin profile": "LinkedIn",
    "linkedin profile url": "LinkedIn",
    "linkedin": "LinkedIn",

    "location": "Location"
}

# =========================================================
# LOCATION PARSING
# =========================================================

US_STATES = [
    "Alabama","Alaska","Arizona","Arkansas","California",
    "Colorado","Connecticut","Delaware","Florida","Georgia",
    "Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas",
    "Kentucky","Louisiana","Maine","Maryland","Massachusetts",
    "Michigan","Minnesota","Mississippi","Missouri","Montana",
    "Nebraska","Nevada","New Hampshire","New Jersey","New Mexico",
    "New York","North Carolina","North Dakota","Ohio","Oklahoma",
    "Oregon","Pennsylvania","Rhode Island","South Carolina",
    "South Dakota","Tennessee","Texas","Utah","Vermont",
    "Virginia","Washington","West Virginia","Wisconsin","Wyoming"
]

CANADA_PROVINCES = [
    "Quebec","Ontario","British Columbia","Alberta",
    "Manitoba","Saskatchewan","Nova Scotia",
    "New Brunswick","Prince Edward Island",
    "Newfoundland and Labrador"
]

def parse_location(location):

    location = str(location)

    country = ""
    province = ""

    parts = [x.strip() for x in location.split(",")]

    if len(parts) > 0:

        country_candidate = parts[-1]

        if country_candidate != "":
            country = country_candidate

    for part in parts:

        if part in US_STATES:
            province = part

        if part in CANADA_PROVINCES:
            province = part

    return country, province

# =========================================================
# HELPERS
# =========================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    value = str(value).strip()

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

# =========================================================
# PROCESS FILE
# =========================================================

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)

    else:
        df = pd.read_excel(uploaded_file)

    # REMOVE GHOST COLUMNS
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

    # CREATE MISSING COLUMNS
    for col in FINAL_COLUMNS:

        if col not in df.columns:
            df[col] = ""

    # LOCATION PARSING
    if "Location" in df.columns:

        for idx, row in df.iterrows():

            country, province = parse_location(
                row["Location"]
            )

            if row["Country"] == "":
                df.at[idx, "Country"] = country

            if row["State or Province"] == "":
                df.at[idx, "State or Province"] = province

    # APPLY GLOBAL VALUES
    df["Subject"] = subject
    df["Description"] = description
    df["Lead Source"] = lead_source
    df["Rating"] = rating
    df["Source Campaign"] = source_campaign
    df["Allow Marketing Communication"] = allow_marketing

    if market_segment != "":
        df["Market Segment"] = market_segment

    if main_application != "":
        df["Main Application"] = main_application

    if industry_sector != "":
        df["Industry Sector"] = industry_sector

    # VALIDATION
    errors = []

    for idx, row in df.iterrows():

        if not is_valid_email(row["Email"]):

            errors.append(
                f"Row {idx + 2}: Invalid email -> {row['Email']}"
            )

    # REMOVE DUPLICATES
    if "Email" in df.columns:

        df = df.drop_duplicates(subset=["Email"])

    # FINAL ORDER
    df = df[FINAL_COLUMNS]

    st.divider()

    st.markdown(
        '<div class="section-title">Dynamics-Ready Import File</div>',
        unsafe_allow_html=True
    )

    st.dataframe(df, use_container_width=True)

    # VALIDATION
    if len(errors) == 0:

        st.markdown("""
        <div class="validation-good">
        File successfully normalized and ready for Dynamics import.
        </div>
        """, unsafe_allow_html=True)

    else:

        st.markdown("""
        <div class="validation-error">
        Validation errors detected.
        </div>
        """, unsafe_allow_html=True)

        for error in errors:
            st.error(error)

    # EXPORT
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Dynamics-Ready CSV",
        data=csv,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv"
    )

# =========================================================
# FOOTER
# =========================================================

st.markdown("---")

st.markdown(
    """
    <div style='text-align:center;
                color:#666;
                padding-top:20px;
                padding-bottom:30px;
                font-size:14px;'>

    Built by
    <a href="mailto:arnaud.joakim@opal-rt.com"
       style="color:#00AEEF;
              text-decoration:none;
              font-weight:600;">
       Arnaud Joakim
    </a>

    </div>
    """,
    unsafe_allow_html=True
)
