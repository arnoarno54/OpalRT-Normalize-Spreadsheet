import streamlit as st
import pandas as pd
import re
from datetime import datetime

# ----------------------------------------
# Page Setup
# ----------------------------------------
st.set_page_config(
    page_title="Opal RT Spreadsheet Cleaner",
    layout="wide"
)
st.markdown("""
<style>
.stApp { background-color: #F4F7FB; }
.main .block-container { padding-top: 1.5rem; max-width: 1450px; }
.hero-container {
    position: relative; border-radius: 20px; overflow: hidden; margin-bottom: 35px;
}
.hero-overlay {
    position: absolute; top: 0; left: 0; right: 0; bottom: 0;
    background: linear-gradient(90deg, rgba(0,32,91,0.92) 0%, rgba(0,32,91,0.75) 40%, rgba(0,32,91,0.35) 100%);
}
.hero-content {
    position: absolute; top: 50%; left: 60px; transform: translateY(-50%);
    color: white; z-index: 2;
}
.hero-title { font-size: 52px; font-weight: 700; margin-bottom: 12px; }
.hero-subtitle { font-size: 20px; opacity: 0.95; }
.section-title { color: #071B4D; font-size: 34px; font-weight: 700; margin-bottom: 10px; }
.stButton>button, .stDownloadButton>button {
    background-color: #00AEEF; color: white; border-radius: 10px; border: none;
    padding: 12px 24px; font-weight: 600;
}
.stButton>button:hover, .stDownloadButton>button:hover {
    background-color: #0093d1; color: white;
}
input, textarea { border-radius: 10px !important; }
.validation-good {
    background: #DFF6E5; color: #146C2E; padding: 18px; border-radius: 10px;
    margin-top: 15px; font-weight: 600;
}
.validation-error {
    background: #FFE3E3; color: #B42318; padding: 18px; border-radius: 10px;
    margin-top: 15px; font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------
# Hero Image and Title
# ----------------------------------------
st.markdown("""
<div class="hero-container">
  <img src="https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg"
       style="width:100%; height:420px; object-fit:cover;">
  <div class="hero-overlay"></div>
  <div class="hero-content">
    <div class="hero-title">Opal RT Spreadsheet Cleaner</div>
    <div class="hero-subtitle">Prepare CRM-ready lead imports for Microsoft Dynamics</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------
# Dropdown Definitions
# ----------------------------------------
MARKET_SEGMENT_APPLICATIONS = {
    "": [""],
    "Aerospace": [
        "", "Autonomous Systems (Aero)", "Avionics System",
        "Electrical Actuators and Servos", "EVTOL",
        "More Electrical Aircraft", "Onboard System",
        "Other (if nothing fits) Aero", "Propulsion and APU",
        "Testbench - Test Automation and Monitoring from RTS"
    ],
    "Automotive": [
        "", "Autonomous Systems (Auto)", "Body & Chassis",
        "Charging", "EV/HEV Powertrain", "Full Vehicle Simulation",
        "ICE Powertrain", "Other (if nothing fits) Auto"
    ],
    "Energy Conversion": [
        "", "Autonomous Systems (Energy Conversion)",
        "Backup Power (UPS)", "Inverter/Converter",
        "Medium and Large Drive (>150KW)",
        "Other (if nothing fits) EnergyConversion",
        "Small Drive (<150KW)"
    ],
    "Marine, Railway, Off-Highway": [
        "", "Autonomous Systems (Marine, Railway, Off-Highway)",
        "BMS Control", "Grid Infrastructure",
        "Onboard Power System", "Other (if nothing fits) Marine, Railway, Off-Highway",
        "Propulsion Control"
    ],
    "Power System": [
        "", "Autonomous Systems (Power Systems)",
        "Conventional Generation", "Converter-Based Energy Resource",
        "Distribution", "FACTS & HVDC", "Microgrid",
        "Other (if nothing fits) PowerSystem", "Substation", "Transmission"
    ]
}
INDUSTRY_SECTORS = [
    "", "Academic - Research or Post-graduate",
    "Academic - Undergraduate", "Consulting & Engineering Firm",
    "Defense", "Electrical Utility", "Manufacturer", "Other",
    "Research Lab - Industrial & Gov.", "Stock - Inventory"
]

# ----------------------------------------
# Global Import Settings (UI)
# ----------------------------------------
st.markdown('<div class="section-title">Global Import Settings</div>', unsafe_allow_html=True)
col1, col2, col3 = st.columns(3)
with col1:
    subject = st.text_input("Subject *", value=f"{datetime.now():%Y%m}Prospection")
    lead_source = st.selectbox("Lead Source", [
        "Shows", "Web", "Prospection", "Webinar", "Referral",
        "Social Media", "Customer Portal", "SPS", "Others"
    ], index=2)
    market_segment = st.selectbox("Market Segment", list(MARKET_SEGMENT_APPLICATIONS.keys()))
with col2:
    rating = st.selectbox("Rating", ["Cold", "Warm", "Hot"], index=0)
    allow_marketing = st.selectbox("Allow Marketing Communication", ["Yes", "No"], index=0)
    main_application = st.selectbox("Main Application", MARKET_SEGMENT_APPLICATIONS[market_segment])
with col3:
    industry_sector = st.selectbox("Industry Sector", INDUSTRY_SECTORS)
    source_campaign = st.text_input("Source Campaign")
    description = st.text_area("Description", height=120)
st.divider()

# ----------------------------------------
# File Upload
# ----------------------------------------
uploaded_file = st.file_uploader("Upload CSV or Excel File", type=["csv", "xlsx"])

# ----------------------------------------
# Final Columns (per import template)
# ----------------------------------------
FINAL_COLUMNS = [
    "(Do Not Modify) Lead", "(Do Not Modify) Row Checksum", "(Do Not Modify) Modified On",
    "Subject", "First Name", "Last Name", "Job Title", "Company Name",
    "Email", "Business Phone", "Country", "State or Province",
    "Description", "Lead Source", "Rating", "Source Campaign",
    "Market Segment", "Main Application", "Industry Sector", "LinkedIn",
    "Allow Marketing Communication"
]
REQUIRED_FIELDS = [
    "Subject", "First Name", "Last Name", "Email",
    "Company Name", "Country", "Market Segment", "Main Application"
]

# ----------------------------------------
# Header Normalization Mapping
# ----------------------------------------
COLUMN_MAPPING = {
    # Name
    "firstname": "First Name", "first name": "First Name", "fname": "First Name", "given name": "First Name",
    "lastname": "Last Name", "last name": "Last Name", "surname": "Last Name", "lname": "Last Name",
    # Company
    "company": "Company Name", "company name": "Company Name", "organization": "Company Name", "org": "Company Name",
    # Job Title
    "job title": "Job Title", "title": "Job Title", "position": "Job Title",
    # Email
    "email": "Email", "email address": "Email", "work email": "Email", "business email": "Email",
    # Phone
    "phone": "Business Phone", "telephone": "Business Phone", "mobile": "Business Phone", "work phone": "Business Phone",
    # LinkedIn
    "linkedin": "LinkedIn", "linkedin profile": "LinkedIn", "linkedin profile url": "LinkedIn",
    # Location
    "location": "Location", "hq location": "Location", "city": "Location"
}

# Lists for parsing location into country/state
US_STATES = [ "Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut","Delaware",
              "Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas","Kentucky",
              "Louisiana","Maine","Maryland","Massachusetts","Michigan","Minnesota","Mississippi",
              "Missouri","Montana","Nebraska","Nevada","New Hampshire","New Jersey","New Mexico",
              "New York","North Carolina","North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania",
              "Rhode Island","South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont",
              "Virginia","Washington","West Virginia","Wisconsin","Wyoming" ]
CANADA_PROVINCES = [
    "Quebec","Ontario","British Columbia","Alberta","Manitoba",
    "Saskatchewan","Nova Scotia","New Brunswick","Prince Edward Island","Newfoundland and Labrador"
]

def clean_text(value):
    """Trim whitespace and collapse spaces."""
    if pd.isna(value): return ""
    value = str(value).strip()
    return re.sub(r"\s+", " ", value)

def clean_email(email):
    """Lowercase and trim email."""
    if pd.isna(email): return ""
    return str(email).strip().lower()

def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return bool(email and re.match(pattern, email))

def parse_location(location):
    """Parse a location string into (country, province) if possible."""
    loc = str(location)
    parts = [p.strip() for p in loc.split(",") if p.strip()]
    country = parts[-1] if parts else ""
    province = ""
    for part in parts:
        if part in US_STATES or part in CANADA_PROVINCES:
            province = part
            break
    return country, province

# ----------------------------------------
# Process Uploaded File
# ----------------------------------------
if uploaded_file:
    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Remove extra unnamed cols
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # Normalize headers
    new_cols = {}
    for col in df.columns:
        norm = str(col).strip().lower()
        if norm in COLUMN_MAPPING:
            new_cols[col] = COLUMN_MAPPING[norm]
    df.rename(columns=new_cols, inplace=True)

    # Clean text fields
    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    # Clean email field
    if "Email" in df.columns:
        df["Email"] = df["Email"].apply(clean_email)

    # Create any missing final columns
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # Parse Location into Country/State
    if "Location" in df.columns:
        for idx, row in df.iterrows():
            country, province = parse_location(row["Location"])
            if row["Country"] == "":
                df.at[idx, "Country"] = country
            if row["State or Province"] == "":
                df.at[idx, "State or Province"] = province

    # Apply global UI values
    df["Subject"] = subject
    df["Description"] = description
    df["Lead Source"] = lead_source
    df["Rating"] = rating
    df["Source Campaign"] = source_campaign
    df["Allow Marketing Communication"] = allow_marketing
    if market_segment:
        df["Market Segment"] = market_segment
    if main_application:
        df["Main Application"] = main_application
    if industry_sector:
        df["Industry Sector"] = industry_sector

    # Validation
    errors = []
    for idx, row in df.iterrows():
        # Check required fields
        for field in REQUIRED_FIELDS:
            if str(row[field]).strip() == "":
                errors.append(f"Row {idx+2}: Missing required field -> {field}")
        # Check email format
        if not is_valid_email(row["Email"]):
            errors.append(f"Row {idx+2}: Invalid email -> {row['Email']}")

    # Drop duplicate emails
    if "Email" in df.columns:
        df = df.drop_duplicates(subset=["Email"])

    # Reorder columns to match template
    df = df[FINAL_COLUMNS]

    # Display final table
    st.divider()
    st.markdown('<div class="section-title">Dynamics-Ready Import File</div>', unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True)

    # Show validation messages
    if not errors:
        st.markdown('<div class="validation-good">File successfully normalized and ready for Dynamics import.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="validation-error">Validation errors detected.</div>', unsafe_allow_html=True)
        for err in errors:
            st.error(err)

    # Download CSV
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button("Download Dynamics-Ready CSV", data=csv_bytes, file_name="opalrt_dynamics_import.csv", mime="text/csv")

# ----------------------------------------
# Footer
# ----------------------------------------
st.markdown("---")
st.markdown("""
<div style="text-align:center; color:#666; padding-top:20px; padding-bottom:30px; font-size:14px;">
Built by <a href="mailto:arnaud.joakim@opal-rt.com" 
style="color:#00AEEF; text-decoration:none; font-weight:600;">Arnaud Joakim</a>
</div>
""", unsafe_allow_html=True)
