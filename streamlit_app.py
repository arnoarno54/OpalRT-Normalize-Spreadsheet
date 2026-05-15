import streamlit as st
import pandas as pd
import re
from datetime import datetime

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="OPAL-RT Dynamics Lead Import Normalizer",
    layout="wide"
)

# =========================================================
# CLEAN OPAL-RT BRANDING
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #F4F7FB;
}

.main .block-container {
    padding-top: 2rem;
    max-width: 1400px;
}

.hero-section {
    background: linear-gradient(135deg,#003B8E 0%, #071B4D 100%);
    padding: 45px;
    border-radius: 18px;
    color: white;
    margin-bottom: 35px;
}

.hero-title {
    font-size: 42px;
    font-weight: 700;
    margin-bottom: 10px;
}

.hero-subtitle {
    font-size: 18px;
    opacity: 0.9;
}

.info-box {
    background: white;
    padding: 25px;
    border-radius: 14px;
    border-left: 6px solid #00AEEF;
    margin-bottom: 25px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.04);
}

.section-title {
    color: #071B4D;
    font-size: 30px;
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

.stDownloadButton>button:hover {
    background-color: #0095d0;
    color: white;
}

.validation-good {
    background: #DFF6E5;
    color: #146C2E;
    padding: 18px;
    border-radius: 10px;
    margin-top: 15px;
}

.validation-bad {
    background: #FFE4E4;
    color: #8A1C1C;
    padding: 18px;
    border-radius: 10px;
    margin-top: 15px;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# HERO SECTION
# =========================================================

st.markdown("""
<div class="hero-section">

<img src="https://www.opal-rt.com/wp-content/uploads/2024/02/OPALRT-logo-white.png" width="240">

<div class="hero-title">
Dynamics Lead Import Normalizer
</div>

<div class="hero-subtitle">
Prepare CRM-ready lead imports for Microsoft Dynamics
</div>

</div>
""", unsafe_allow_html=True)

# =========================================================
# REQUIRED FIELD BOX
# =========================================================

st.markdown("""
<div class="info-box">

<h3 style="color:#071B4D;">
Mandatory Lead Fields
</h3>

<ul style="color:#222; font-size:16px;">
<li>Subject</li>
<li>First Name</li>
<li>Last Name</li>
<li>Email</li>
<li>Company</li>
<li>Country/Region</li>
<li>State/Province</li>
<li>Market Segment</li>
<li>Main Application</li>
</ul>

<p style="color:#555;">
These validations are automatically checked before export.
</p>

</div>
""", unsafe_allow_html=True)

# =========================================================
# DROPDOWN VALUES
# =========================================================

MARKET_SEGMENTS = [
    "Let AI figure it out",
    "Aerospace",
    "Automotive",
    "Energy Conversion",
    "Marine, Railway, Off-Highway",
    "Power System"
]

MAIN_APPLICATIONS = [
    "Let AI figure it out",
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

# =========================================================
# IMPORT SETTINGS
# =========================================================

st.markdown(
    '<div class="section-title">Global Import Settings</div>',
    unsafe_allow_html=True
)

st.write(
    "These values will automatically be applied to all imported rows."
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
        "Market Segment *",
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
        "Main Application *",
        MAIN_APPLICATIONS
    )

with col3:

    source_campaign = st.text_input(
        "Source Campaign"
    )

    description = st.text_area(
        "Description",
        height=140
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

COLUMN_MAPPING = {
    "firstname": "First Name",
    "first name": "First Name",
    "fname": "First Name",

    "lastname": "Last Name",
    "last name": "Last Name",

    "company": "Company Name",

    "email": "Email",

    "phone": "Business Phone",

    "country": "Country",
    "country/region": "Country",

    "state": "State or Province",
    "province": "State or Province",

    "linkedin": "LinkedIn"
}

# =========================================================
# HELPERS
# =========================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    value = str(value).strip()

    value = re.sub(r"\s+", " ", value)

    return value


def clean_email(email):

    if pd.isna(email):
        return ""

    return str(email).strip().lower()


def infer_market_segment(row):

    text = " ".join([
        str(row.get("Company Name", "")),
        str(row.get("Job Title", "")),
        str(row.get("Description", "")),
        str(row.get("Email", ""))
    ]).lower()

    if any(x in text for x in [
        "grid",
        "utility",
        "transmission",
        "distribution",
        "microgrid",
        "renewable",
        "hvdc"
    ]):
        return "Power System"

    if any(x in text for x in [
        "battery",
        "ev",
        "vehicle",
        "automotive"
    ]):
        return "Automotive"

    if any(x in text for x in [
        "aircraft",
        "flight",
        "avionics",
        "aerospace"
    ]):
        return "Aerospace"

    return "Power System"


def infer_main_application(row):

    text = " ".join([
        str(row.get("Company Name", "")),
        str(row.get("Job Title", "")),
        str(row.get("Description", "")),
        str(row.get("Email", ""))
    ]).lower()

    if "microgrid" in text:
        return "Microgrid"

    if "hvdc" in text:
        return "FACTS & HVDC"

    if "battery" in text:
        return "BMS Control"

    if "charging" in text:
        return "Charging"

    if "distribution" in text:
        return "Distribution"

    if "transmission" in text:
        return "Transmission"

    return "Grid Infrastructure"

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

    # APPLY GLOBAL VALUES
    df["Subject"] = subject
    df["Description"] = description
    df["Lead Source"] = lead_source
    df["Rating"] = rating
    df["Source Campaign"] = source_campaign
    df["Allow Marketing Communication"] = allow_marketing

    # AI INFERENCE
    if market_segment == "Let AI figure it out":

        df["Market Segment"] = df.apply(
            infer_market_segment,
            axis=1
        )

    else:
        df["Market Segment"] = market_segment

    if main_application == "Let AI figure it out":

        df["Main Application"] = df.apply(
            infer_main_application,
            axis=1
        )

    else:
        df["Main Application"] = main_application

    # FINAL ORDER
    df = df[FINAL_COLUMNS]

    st.divider()

    st.markdown(
        '<div class="section-title">Dynamics-Ready Import File</div>',
        unsafe_allow_html=True
    )

    st.dataframe(df, use_container_width=True)

    # EXPORT
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Dynamics-Ready CSV",
        data=csv,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv"
    )

    st.markdown("""
    <div class="validation-good">
    File successfully normalized and ready for Dynamics import.
    </div>
    """, unsafe_allow_html=True)
