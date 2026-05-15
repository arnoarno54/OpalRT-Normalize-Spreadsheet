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

MARKET_SEGMENTS = [
    "",
    "Aerospace",
    "Automotive",
    "Energy Conversion",
    "Marine, Railway, Off-Highway",
    "Power System"
]

MAIN_APPLICATIONS = [
    "",
    "Autonomous Systems (Aero)",
    "Avionics System",
    "Electrical Actuators and Servos",
    "EVTOL",
    "More Electrical Aircraft",
    "Onboard System",
    "Charging",
    "EV/HEV Powertrain",
    "Full Vehicle Simulation",
    "ICE Powertrain",
    "Backup Power (UPS)",
    "Inverter/Converter",
    "Medium and Large Drive (>150KW)",
    "BMS Control",
    "Grid Infrastructure",
    "Onboard Power System",
    "Propulsion Control",
    "Conventional Generation",
    "Converter-Based Energy Resource",
    "Distribution",
    "FACTS & HVDC",
    "Microgrid",
    "Substation",
    "Transmission"
]

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
        MARKET_SEGMENTS
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
        MAIN_APPLICATIONS
    )

with col3:

    source_campaign = st.text_input(
        "Source Campaign"
    )

    industry_sector = st.selectbox(
        "Industry Sector",
        INDUSTRY_SECTORS
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

REQUIRED_FIELDS = [
    "Subject",
    "First Name",
    "Last Name",
    "Email",
    "Company Name",
    "Country"
]

COLUMN_MAPPING = {
    "firstname": "First Name",
    "first name": "First Name",
    "fname": "First Name",

    "lastname": "Last Name",
    "last name": "Last Name",
    "lname": "Last Name",

    "company": "Company Name",
    "organization": "Company Name",
    "org": "Company Name",

    "email": "Email",
    "mail": "Email",
    "email address": "Email",

    "phone": "Business Phone",
    "telephone": "Business Phone",
    "mobile": "Business Phone",

    "linkedin": "LinkedIn",
    "linkedin url": "LinkedIn",
    "linkedin profile": "LinkedIn",

    "country": "Country",
    "country/region": "Country",

    "state": "State or Province",
    "province": "State or Province",

    "location": "Location",
    "city": "Location"
}

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


def infer_location(location):

    location = str(location).lower()

    # Canada
    if "quebec" in location or "montreal" in location:
        return ("Canada", "Quebec")

    if "toronto" in location or "ontario" in location:
        return ("Canada", "Ontario")

    if "vancouver" in location:
        return ("Canada", "British Columbia")

    # USA
    if "california" in location:
        return ("United States", "California")

    if "texas" in location:
        return ("United States", "Texas")

    if "new york" in location:
        return ("United States", "New York")

    # Europe
    if "france" in location or "paris" in location:
        return ("France", "")

    if "germany" in location:
        return ("Germany", "")

    if "uk" in location or "united kingdom" in location or "london" in location:
        return ("United Kingdom", "")

    return ("", "")

# =========================================================
# PROCESS FILE
# =========================================================

if uploaded_file:

    # READ FILE
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

    # LOCATION INFERENCE
    if "Location" in df.columns:

        for idx, row in df.iterrows():

            if row["Country"] == "":

                country, province = infer_location(
                    row["Location"]
                )

                df.at[idx, "Country"] = country

                if province != "":
                    df.at[idx, "State or Province"] = province

    # APPLY GLOBAL SETTINGS
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

        for field in REQUIRED_FIELDS:

            if str(row[field]).strip() == "":

                errors.append(
                    f"Row {idx + 2}: Missing required field -> {field}"
                )

        if not is_valid_email(row["Email"]):

            errors.append(
                f"Row {idx + 2}: Invalid email -> {row['Email']}"
            )

    # REMOVE DUPLICATES
    if "Email" in df.columns:
        df = df.drop_duplicates(subset=["Email"])

    # FINAL ORDER
    df = df[FINAL_COLUMNS]

    st.markdown(
        '<div class="section-title">Dynamics-Ready Import File</div>',
        unsafe_allow_html=True
    )

    st.dataframe(df, use_container_width=True)

    # VALIDATION RESULTS
    if len(errors) == 0:

        st.markdown("""
        <div class="validation-good">
        File successfully normalized and ready for Dynamics import.
        </div>
        """, unsafe_allow_html=True)

    else:

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
