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
# OPAL-RT BRANDING
# =========================================================

st.markdown("""
<style>

html, body, [class*="css"] {
    background-color: #071B4D;
    color: white;
}

.main {
    background: linear-gradient(180deg,#071B4D 0%, #0A2A7A 100%);
}

.block-container {
    padding-top: 2rem;
    max-width: 1500px;
}

h1 {
    color: white;
    font-size: 3rem;
    font-weight: 700;
}

h2, h3 {
    color: white;
}

p, li, label, div {
    color: white;
}

section[data-testid="stSidebar"] {
    background-color: #05163D;
}

.stTextInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"] {
    background-color: white !important;
    color: black !important;
    border-radius: 8px;
}

.stButton>button,
.stDownloadButton>button {
    background-color: #00AEEF;
    color: white;
    border-radius: 10px;
    border: none;
    font-weight: 600;
    padding: 12px 20px;
}

.required-box {
    background: rgba(255,255,255,0.08);
    border-left: 5px solid #00AEEF;
    padding: 20px;
    border-radius: 10px;
    margin-bottom: 30px;
}

.stDataFrame {
    background-color: white;
}

hr {
    border-color: rgba(255,255,255,0.2);
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================

col1, col2 = st.columns([1, 4])

with col1:
    st.image(
        "https://www.opal-rt.com/wp-content/uploads/2024/02/OPALRT-logo-white.png",
        width=240
    )

with col2:
    st.title("Dynamics Lead Import Normalizer")
    st.markdown(
        "Prepare CRM-ready lead imports for Microsoft Dynamics"
    )

st.markdown("---")

# =========================================================
# REQUIRED FIELDS
# =========================================================

st.markdown("""
<div class="required-box">

<h3>Mandatory Lead Fields</h3>

<ul>
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

These validations are automatically checked before export.

</div>
""", unsafe_allow_html=True)

# =========================================================
# MARKET SEGMENTS
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
# GLOBAL SETTINGS
# =========================================================

st.header("Global Import Settings")

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

    market_segment = st.selectbox(
        "Market Segment *",
        MARKET_SEGMENTS
    )

with colB:

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

with colC:

    source_campaign = st.text_input(
        "Source Campaign"
    )

    description = st.text_area(
        "Description",
        height=120
    )

st.markdown("---")

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
    "firstname": "First Name",
    "first name": "First Name",
    "fname": "First Name",

    "lastname": "Last Name",
    "last name": "Last Name",
    "surname": "Last Name",

    "company": "Company Name",
    "organization": "Company Name",

    "email": "Email",
    "mail": "Email",

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
        "power system",
        "microgrid",
        "hvdc",
        "renewable"
    ]):
        return "Power System"

    if any(x in text for x in [
        "automotive",
        "vehicle",
        "ev",
        "battery",
        "charging"
    ]):
        return "Automotive"

    if any(x in text for x in [
        "aircraft",
        "avionics",
        "aerospace",
        "flight"
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

    if "charging" in text:
        return "Charging"

    if "battery" in text:
        return "BMS Control"

    if "transmission" in text:
        return "Transmission"

    if "distribution" in text:
        return "Distribution"

    if "renewable" in text:
        return "Converter-Based Energy Resource"

    return "Grid Infrastructure"

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

    # STANDARDIZE COLUMNS
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

    # =====================================================
    # AI INFERENCE
    # =====================================================

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

    st.subheader("Dynamics-Ready Import File")

    st.dataframe(df)

    # EXPORT
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Dynamics-Ready CSV",
        data=csv,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv"
    )
