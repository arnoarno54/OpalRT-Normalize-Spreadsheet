import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# ─── Page Config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OPAL-RT Spreadsheet Cleaner",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS Styling ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

  html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
  }

  /* Hero banner */
  .hero-banner {
    background: linear-gradient(135deg, #003366 0%, #0055A5 50%, #0077CC 100%);
    padding: 0;
    border-radius: 12px;
    margin-bottom: 28px;
    overflow: hidden;
    position: relative;
    min-height: 160px;
    display: flex;
    align-items: center;
  }
  .hero-bg {
    position: absolute; top:0; left:0; width:100%; height:100%;
    background-image: url('https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg');
    background-size: cover;
    background-position: center;
    opacity: 0.18;
  }
  .hero-content {
    position: relative; z-index: 1;
    padding: 32px 40px;
  }
  .hero-title {
    color: #FFFFFF;
    font-size: 2rem;
    font-weight: 700;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
  }
  .hero-subtitle {
    color: #A8CFEE;
    font-size: 1rem;
    font-weight: 400;
    margin: 0;
  }
  .hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 20px;
    padding: 4px 14px;
    color: #D0E8FF;
    font-size: 0.75rem;
    font-weight: 500;
    margin-bottom: 10px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
  }

  /* Section cards */
  .section-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 24px 28px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  }
  .section-title {
    font-size: 1rem;
    font-weight: 600;
    color: #0055A5;
    margin-bottom: 16px;
    padding-bottom: 10px;
    border-bottom: 2px solid #E8F0FA;
    display: flex;
    align-items: center;
    gap: 8px;
  }

  /* Validation boxes */
  .success-box {
    background: #F0FDF4;
    border: 1px solid #86EFAC;
    border-left: 4px solid #22C55E;
    border-radius: 8px;
    padding: 16px 20px;
    color: #166534;
    font-weight: 500;
    margin: 16px 0;
  }
  .error-box {
    background: #FFF1F2;
    border: 1px solid #FECDD3;
    border-left: 4px solid #EF4444;
    border-radius: 8px;
    padding: 16px 20px;
    color: #9F1239;
    margin: 16px 0;
  }
  .error-box h4 {
    margin: 0 0 10px 0;
    font-size: 0.95rem;
    font-weight: 600;
  }
  .error-item {
    font-size: 0.85rem;
    padding: 4px 0;
    border-bottom: 1px solid rgba(239,68,68,0.1);
    color: #7F1D1D;
  }
  .warning-box {
    background: #FFFBEB;
    border: 1px solid #FDE68A;
    border-left: 4px solid #F59E0B;
    border-radius: 8px;
    padding: 14px 18px;
    color: #78350F;
    font-size: 0.875rem;
    margin: 12px 0;
  }

  /* Mandatory asterisk */
  .required-star { color: #EF4444; font-weight: 700; }

  /* Stats row */
  .stats-row {
    display: flex;
    gap: 16px;
    margin: 16px 0;
    flex-wrap: wrap;
  }
  .stat-chip {
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
    border-radius: 8px;
    padding: 10px 18px;
    text-align: center;
    min-width: 110px;
  }
  .stat-chip .stat-num {
    font-size: 1.5rem;
    font-weight: 700;
    color: #1D4ED8;
    display: block;
  }
  .stat-chip .stat-label {
    font-size: 0.72rem;
    color: #64748B;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }

  /* Footer */
  .footer {
    text-align: center;
    padding: 20px;
    color: #94A3B8;
    font-size: 0.8rem;
    border-top: 1px solid #E2E8F0;
    margin-top: 40px;
  }
  .footer a { color: #0055A5; text-decoration: none; }
  .footer a:hover { text-decoration: underline; }

  /* Override Streamlit button */
  .stDownloadButton > button {
    background: linear-gradient(135deg, #0055A5, #0077CC) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 28px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    transition: opacity 0.2s !important;
  }
  .stDownloadButton > button:hover { opacity: 0.9 !important; }

  div[data-testid="stFileUploader"] {
    border: 2px dashed #93C5FD;
    border-radius: 10px;
    padding: 12px;
    background: #F8FBFF;
  }

  /* Tab styling */
  .stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #F1F5F9;
    border-radius: 8px;
    padding: 4px;
  }
  .stTabs [data-baseweb="tab"] {
    border-radius: 6px;
    font-weight: 500;
    color: #475569;
  }
  .stTabs [aria-selected="true"] {
    background: #FFFFFF !important;
    color: #0055A5 !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  }
</style>
""", unsafe_allow_html=True)

# ─── Constants ───────────────────────────────────────────────────────────────
EXPORT_COLUMNS = [
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
    "Allow Marketing Communication",
]

MANDATORY_FIELDS = ["Subject", "First Name", "Last Name", "Email", "Company Name", "Country"]

LEAD_SOURCE_VALUES = ["", "Shows", "Web", "Prospection", "Webinar", "Referral",
                      "Social Media", "Customer Portal", "SPS", "Others"]

RATING_VALUES = ["", "Cold", "Warm", "Hot"]

ALLOW_MARKETING_VALUES = ["", "Yes", "No"]

INDUSTRY_SECTOR_VALUES = [
    "",
    "Academic - Research or Post-graduate",
    "Academic - Undergraduate",
    "Consulting & Engineering Firm",
    "Defense",
    "Electrical Utility",
    "Manufacturer",
    "Other",
    "Research Lab - Industrial & Gov.",
    "Stock - Inventory",
]

MARKET_SEGMENT_VALUES = [
    "",
    "Aerospace",
    "Automotive",
    "Energy Conversion",
    "Marine, Railway, Off-Highway",
    "Power System",
]

MAIN_APPLICATION_MAP = {
    "": [""],
    "Aerospace": [
        "",
        "Autonomous Systems (Aero)",
        "Avionics System",
        "Electrical Actuators and Servos",
        "EVTOL",
        "More Electrical Aircraft",
        "Onboard System",
        "Other (if nothing fits) Aero",
        "Propulsion and APU",
        "Testbench - Test Automation and Monitoring from RTS",
    ],
    "Automotive": [
        "",
        "Autonomous Systems (Auto)",
        "Body & Chassis",
        "Charging",
        "EV/HEV Powertrain",
        "Full Vehicle Simulation",
        "ICE Powertrain",
        "Other (if nothing fits) Auto",
    ],
    "Energy Conversion": [
        "",
        "Autonomous Systems (Energy Conversion)",
        "Backup Power (UPS)",
        "Inverter/Converter",
        "Medium and Large Drive (>150KW)",
        "Other (if nothing fits) EnergyConversion",
        "Small Drive (<150KW)",
    ],
    "Marine, Railway, Off-Highway": [
        "",
        "Autonomous Systems (Marine, Railway, Off-Highway)",
        "BMS Control",
        "Grid Infrastructure",
        "Onboard Power System",
        "Other (if nothing fits) Marine, Railway, Off-Highway",
        "Propulsion Control",
    ],
    "Power System": [
        "",
        "Autonomous Systems (Power Systems)",
        "Conventional Generation",
        "Converter-Based Energy Resource",
        "Distribution",
        "FACTS & HVDC",
        "Microgrid",
        "Other (if nothing fits) PowerSystem",
        "Substation",
        "Transmission",
    ],
}

FIELD_MAX_LENGTHS = {
    "First Name": 58,
    "Last Name": 50,
    "Company Name": 100,
    "Job Title": 100,
    "Email": 100,
    "LinkedIn": 500,
    "Description": 2000,
    "Subject": 300,
    "Business Phone": 50,
}

# ─── US States & Canadian Provinces ─────────────────────────────────────────
US_STATES = {
    "alabama": "Alabama", "alaska": "Alaska", "arizona": "Arizona", "arkansas": "Arkansas",
    "california": "California", "colorado": "Colorado", "connecticut": "Connecticut",
    "delaware": "Delaware", "florida": "Florida", "georgia": "Georgia", "hawaii": "Hawaii",
    "idaho": "Idaho", "illinois": "Illinois", "indiana": "Indiana", "iowa": "Iowa",
    "kansas": "Kansas", "kentucky": "Kentucky", "louisiana": "Louisiana", "maine": "Maine",
    "maryland": "Maryland", "massachusetts": "Massachusetts", "michigan": "Michigan",
    "minnesota": "Minnesota", "mississippi": "Mississippi", "missouri": "Missouri",
    "montana": "Montana", "nebraska": "Nebraska", "nevada": "Nevada",
    "new hampshire": "New Hampshire", "new jersey": "New Jersey", "new mexico": "New Mexico",
    "new york": "New York", "north carolina": "North Carolina", "north dakota": "North Dakota",
    "ohio": "Ohio", "oklahoma": "Oklahoma", "oregon": "Oregon", "pennsylvania": "Pennsylvania",
    "rhode island": "Rhode Island", "south carolina": "South Carolina",
    "south dakota": "South Dakota", "tennessee": "Tennessee", "texas": "Texas",
    "utah": "Utah", "vermont": "Vermont", "virginia": "Virginia", "washington": "Washington",
    "west virginia": "West Virginia", "wisconsin": "Wisconsin", "wyoming": "Wyoming",
    "district of columbia": "District of Columbia", "dc": "District of Columbia",
}

CA_PROVINCES = {
    "alberta": "Alberta", "british columbia": "British Columbia", "manitoba": "Manitoba",
    "new brunswick": "New Brunswick", "newfoundland": "Newfoundland and Labrador",
    "newfoundland and labrador": "Newfoundland and Labrador",
    "northwest territories": "Northwest Territories", "nova scotia": "Nova Scotia",
    "nunavut": "Nunavut", "ontario": "Ontario", "prince edward island": "Prince Edward Island",
    "quebec": "Quebec", "québec": "Quebec", "saskatchewan": "Saskatchewan", "yukon": "Yukon",
    "qc": "Quebec", "on": "Ontario", "bc": "British Columbia", "ab": "Alberta",
}

# Common country name normalizations
COUNTRY_ALIASES = {
    "usa": "United States", "us": "United States", "u.s.": "United States",
    "u.s.a.": "United States", "united states of america": "United States",
    "uk": "United Kingdom", "u.k.": "United Kingdom", "great britain": "United Kingdom",
    "england": "United Kingdom",
    "ca": "Canada", "can": "Canada",
    "fr": "France",
    "de": "Germany",
    "jp": "Japan",
    "cn": "China",
    "au": "Australia",
    "br": "Brazil",
    "in": "India",
    "mx": "Mexico",
    "kr": "South Korea",
    "se": "Sweden",
    "no": "Norway",
    "fi": "Finland",
    "dk": "Denmark",
    "nl": "Netherlands",
    "be": "Belgium",
    "ch": "Switzerland",
    "at": "Austria",
    "es": "Spain",
    "it": "Italy",
    "pt": "Portugal",
    "pl": "Poland",
    "cz": "Czech Republic",
    "ro": "Romania",
    "hu": "Hungary",
    "gr": "Greece",
    "tr": "Turkey",
    "sa": "Saudi Arabia",
    "ae": "United Arab Emirates",
    "uae": "United Arab Emirates",
    "sg": "Singapore",
    "hk": "Hong Kong",
    "tw": "Taiwan",
    "nz": "New Zealand",
    "za": "South Africa",
    "il": "Israel",
    "ir": "Iran",
    "pk": "Pakistan",
    "eg": "Egypt",
    "ng": "Nigeria",
    "ke": "Kenya",
    "th": "Thailand",
    "vn": "Vietnam",
    "id": "Indonesia",
    "my": "Malaysia",
    "ph": "Philippines",
    "ar": "Argentina",
    "cl": "Chile",
    "co": "Colombia",
    "pe": "Peru",
}

# ─── Helper Functions ─────────────────────────────────────────────────────────
def normalize_col(col_name):
    """Lowercase, strip, collapse spaces for fuzzy matching."""
    if not col_name or not isinstance(col_name, str):
        return ""
    return re.sub(r'\s+', ' ', col_name.lower().strip())


def detect_column(df_cols, patterns):
    """Find the first column matching any pattern (fuzzy)."""
    normalized = {normalize_col(c): c for c in df_cols}
    for pat in patterns:
        pat_n = normalize_col(pat)
        if pat_n in normalized:
            return normalized[pat_n]
    # Partial match fallback
    for pat in patterns:
        pat_n = normalize_col(pat)
        for nc, orig in normalized.items():
            if pat_n in nc or nc in pat_n:
                return orig
    return None


COLUMN_PATTERNS = {
    "First Name":       ["first name", "firstname", "fname", "given name", "prénom", "prenom", "first"],
    "Last Name":        ["last name", "lastname", "lname", "surname", "family name", "nom", "last"],
    "Company Name":     ["company name", "company", "organization", "organisation", "org", "employer",
                         "account name", "firm", "société", "societe"],
    "Job Title":        ["job title", "title", "position", "role", "function", "fonction",
                         "designation", "poste", "jobtitle"],
    "Email":            ["work email", "email address", "business email", "corporate email",
                         "email", "e-mail", "courriel", "mail"],
    "Business Phone":   ["mobile phone", "work phone", "business phone", "phone number",
                         "telephone", "phone", "mobile", "tel", "téléphone"],
    "Country":          ["country/region", "country", "pays", "nation"],
    "State or Province":["state/province", "state or province", "province", "state", "région", "region"],
    "LinkedIn":         ["linkedin profile url", "linkedin profile", "linkedin url",
                         "linkedin", "linked in"],
    "Location":         ["location", "city/state/country", "city, state", "address", "localisation"],
}


def parse_location(location_str):
    """
    Parse a freeform location string into (country, state_province).
    Handles formats like:
      Montreal, Quebec, Canada
      Dallas, Texas, United States
      Paris, France
      London, United Kingdom
      São Paulo, SP, Brazil
    """
    if not location_str or not isinstance(location_str, str):
        return "", ""

    parts = [p.strip() for p in location_str.split(",") if p.strip()]
    if not parts:
        return "", ""

    country = ""
    state_province = ""

    # Try to identify country from last part
    last = parts[-1].lower().strip()
    if last in COUNTRY_ALIASES:
        country = COUNTRY_ALIASES[last]
    else:
        # Title-case it as a country
        country = parts[-1].strip().title()

    # Now look for state/province in remaining parts
    if len(parts) >= 3:
        candidate = parts[-2].lower().strip()
        if candidate in US_STATES:
            state_province = US_STATES[candidate]
        elif candidate in CA_PROVINCES:
            state_province = CA_PROVINCES[candidate]
        else:
            # Check if it looks like a known state/province abbreviation
            if len(candidate) == 2 and candidate in (set(US_STATES) | set(CA_PROVINCES)):
                state_province = US_STATES.get(candidate) or CA_PROVINCES.get(candidate, "")
            else:
                state_province = parts[-2].strip()
    elif len(parts) == 2:
        # Could be "City, Country" — no state info
        candidate = parts[0].lower().strip()
        if candidate in US_STATES:
            state_province = US_STATES[candidate]
            country = country  # already set
        elif candidate in CA_PROVINCES:
            state_province = CA_PROVINCES[candidate]
        # else: first part is just city, no state

    return country, state_province


def normalize_country(val):
    """Normalize a country field value."""
    if not val or not isinstance(val, str):
        return ""
    v = val.strip()
    key = v.lower()
    if key in COUNTRY_ALIASES:
        return COUNTRY_ALIASES[key]
    return v.strip()


def normalize_state(val):
    if not val or not isinstance(val, str):
        return ""
    v = val.strip().lower()
    if v in US_STATES:
        return US_STATES[v]
    if v in CA_PROVINCES:
        return CA_PROVINCES[v]
    return val.strip()


def clean_text(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', s)
    s = re.sub(r' +', ' ', s)
    return s.strip()


def clean_email(val):
    v = clean_text(val).lower()
    return v


def validate_email(email):
    if not email:
        return False
    pattern = r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def process_dataframe(df_raw, global_settings, column_map):
    """
    Main processing pipeline:
    1. Map source columns to Dynamics columns
    2. Clean & normalize
    3. Apply global settings
    4. Return (clean_df, errors_list)
    """
    errors = []
    rows = []

    # Remove fully-empty rows
    df_raw = df_raw.dropna(how='all')
    # Remove duplicate emails (keep first)
    email_col_src = column_map.get("Email")
    if email_col_src and email_col_src in df_raw.columns:
        df_raw = df_raw.drop_duplicates(subset=[email_col_src], keep='first')

    location_col = None
    for c in df_raw.columns:
        if normalize_col(c) in ["location", "city/state/country", "city, state",
                                 "localisation", "address", "location/country"]:
            location_col = c
            break

    for idx, row in df_raw.iterrows():
        row_num = idx + 2  # human-readable (1-indexed + header)
        rec = {col: "" for col in EXPORT_COLUMNS}
        row_errors = []

        # ── Map columns ──────────────────────────────────────────────────
        for dyn_col, src_col in column_map.items():
            if src_col and src_col in df_raw.columns:
                val = clean_text(row.get(src_col, ""))
                if dyn_col == "Email":
                    val = clean_email(row.get(src_col, ""))
                elif dyn_col == "Country":
                    val = normalize_country(val)
                elif dyn_col == "State or Province":
                    val = normalize_state(val)
                rec[dyn_col] = val

        # ── Parse Location field ─────────────────────────────────────────
        if location_col and location_col in df_raw.columns:
            loc_val = clean_text(row.get(location_col, ""))
            if loc_val:
                parsed_country, parsed_state = parse_location(loc_val)
                # Only fill if not already populated from dedicated columns
                if not rec.get("Country") and parsed_country:
                    rec["Country"] = parsed_country
                if not rec.get("State or Province") and parsed_state:
                    rec["State or Province"] = parsed_state

        # ── Apply global settings (only if row doesn't already have value) ──
        for field, value in global_settings.items():
            if value:  # Only apply non-empty global settings
                if field in ("Market Segment", "Main Application", "Industry Sector",
                             "Lead Source", "Rating", "Allow Marketing Communication",
                             "Source Campaign", "Description"):
                    if not rec.get(field):
                        rec[field] = value
                elif field == "Subject":
                    rec[field] = value  # Subject always from global

        # ── Validations ──────────────────────────────────────────────────
        email_val = rec.get("Email", "")
        if email_val and not validate_email(email_val):
            row_errors.append(f"Row {row_num}: Invalid email → {email_val}")

        for mf in MANDATORY_FIELDS:
            if not rec.get(mf, "").strip():
                row_errors.append(f"Row {row_num}: Missing required field → {mf}")

        for field, max_len in FIELD_MAX_LENGTHS.items():
            val = rec.get(field, "")
            if val and len(val) > max_len:
                row_errors.append(
                    f"Row {row_num}: {field} exceeds {max_len} characters (has {len(val)})"
                )
                rec[field] = val[:max_len]

        errors.extend(row_errors)
        rows.append(rec)

    result_df = pd.DataFrame(rows, columns=EXPORT_COLUMNS)
    return result_df, errors


def detect_columns(df):
    """Auto-detect source columns for each Dynamics field."""
    mapping = {}
    for dyn_col, patterns in COLUMN_PATTERNS.items():
        if dyn_col == "Location":
            continue
        found = detect_column(df.columns.tolist(), patterns)
        mapping[dyn_col] = found
    return mapping


# ─── Hero Banner ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-banner">
  <div class="hero-bg"></div>
  <div class="hero-content">
    <div class="hero-badge">⚡ OPAL-RT Internal Tool</div>
    <h1 class="hero-title">OPAL-RT Spreadsheet Cleaner</h1>
    <p class="hero-subtitle">Prepare CRM-ready lead imports for Microsoft Dynamics</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📤 Upload & Configure", "🗺️ Column Mapping", "✅ Validate & Export"])

# ─── Session State ────────────────────────────────────────────────────────────
if "df_raw" not in st.session_state:
    st.session_state.df_raw = None
if "column_map" not in st.session_state:
    st.session_state.column_map = {}
if "processed_df" not in st.session_state:
    st.session_state.processed_df = None
if "errors" not in st.session_state:
    st.session_state.errors = []

# ─── Tab 1: Upload & Global Settings ────────────────────────────────────────
with tab1:
    col_upload, col_settings = st.columns([1, 1], gap="large")

    with col_upload:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📁 Upload Source File</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Upload CSV or Excel File",
            type=["csv", "xlsx"],
            help="Accepted formats: .csv, .xlsx",
            label_visibility="collapsed"
        )
        if uploaded:
            try:
                if uploaded.name.endswith(".csv"):
                    df_raw = pd.read_csv(uploaded, dtype=str)
                else:
                    df_raw = pd.read_excel(uploaded, dtype=str)

                # Remove unnamed / ghost columns
                df_raw = df_raw.loc[:, ~df_raw.columns.str.match(r'^Unnamed')]
                df_raw = df_raw.loc[:, df_raw.columns.notna()]
                df_raw = df_raw.loc[:, df_raw.columns.astype(str).str.strip() != ""]
                df_raw.columns = [str(c).strip() for c in df_raw.columns]

                st.session_state.df_raw = df_raw
                auto_map = detect_columns(df_raw)
                st.session_state.column_map = auto_map

                st.markdown(f"""
                <div class="stats-row">
                  <div class="stat-chip">
                    <span class="stat-num">{len(df_raw)}</span>
                    <span class="stat-label">Total Rows</span>
                  </div>
                  <div class="stat-chip">
                    <span class="stat-num">{len(df_raw.columns)}</span>
                    <span class="stat-label">Columns Found</span>
                  </div>
                  <div class="stat-chip">
                    <span class="stat-num">{sum(1 for v in auto_map.values() if v)}</span>
                    <span class="stat-label">Auto-Mapped</span>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("**Detected source columns:**")
                cols_preview = list(df_raw.columns[:20])
                st.code(", ".join(cols_preview) + ("..." if len(df_raw.columns) > 20 else ""))

            except Exception as e:
                st.error(f"Error reading file: {e}")
        else:
            st.markdown("""
            <div class="warning-box">
            📂 No file uploaded yet. Drag & drop or click to browse for a <strong>.csv</strong> or <strong>.xlsx</strong> file.
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col_settings:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚙️ Global Import Settings <span style="font-size:0.75rem;color:#64748B;font-weight:400;">(applied to all rows)</span></div>', unsafe_allow_html=True)

        default_subject = datetime.now().strftime("%Y%m") + "Prospection"

        st.markdown('<span class="required-star">*</span> Subject', unsafe_allow_html=True)
        gs_subject = st.text_input("Subject", value=default_subject, label_visibility="collapsed", key="gs_subject")

        col_a, col_b = st.columns(2)
        with col_a:
            gs_lead_source = st.selectbox("Lead Source", LEAD_SOURCE_VALUES, key="gs_lead_source")
            gs_rating = st.selectbox("Rating", RATING_VALUES, key="gs_rating")
            gs_allow_marketing = st.selectbox("Allow Marketing Communication",
                                              ALLOW_MARKETING_VALUES, key="gs_allow_marketing")
            gs_industry = st.selectbox("Industry Sector", INDUSTRY_SECTOR_VALUES, key="gs_industry")
        with col_b:
            gs_campaign = st.text_input("Source Campaign", key="gs_campaign")
            gs_market_segment = st.selectbox("Market Segment", MARKET_SEGMENT_VALUES, key="gs_market_segment")
            gs_main_app = st.selectbox(
                "Main Application",
                MAIN_APPLICATION_MAP.get(st.session_state.get("gs_market_segment", ""), [""]),
                key="gs_main_app"
            )
            gs_description = st.text_area("Description", height=68, key="gs_description")

        st.markdown('</div>', unsafe_allow_html=True)

# ─── Tab 2: Column Mapping ────────────────────────────────────────────────────
with tab2:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🗺️ Source → Dynamics Column Mapping</div>', unsafe_allow_html=True)

    if st.session_state.df_raw is None:
        st.markdown("""
        <div class="warning-box">⬆️ Please upload a file in the <strong>Upload & Configure</strong> tab first.</div>
        """, unsafe_allow_html=True)
    else:
        df = st.session_state.df_raw
        source_cols = ["(not mapped)"] + list(df.columns)

        st.markdown("""
        <div class="warning-box">
        ✨ Auto-detection has pre-filled these mappings. Review and adjust as needed.
        <strong>Location</strong> fields are automatically parsed into Country and State/Province.
        </div>
        """, unsafe_allow_html=True)

        updated_map = {}
        dynamics_fields = [
            "First Name", "Last Name", "Company Name", "Job Title",
            "Email", "Business Phone", "Country", "State or Province",
            "LinkedIn", "Location",
        ]

        mandatory_set = {"First Name", "Last Name", "Company Name", "Email", "Country"}
        cols_left, cols_right = st.columns(2)

        for i, dyn_field in enumerate(dynamics_fields):
            current = st.session_state.column_map.get(dyn_field, None)
            default_idx = 0
            if current and current in source_cols:
                default_idx = source_cols.index(current)

            star = " *" if dyn_field in mandatory_set else ""
            label = f"{dyn_field}{star}"

            if i % 2 == 0:
                with cols_left:
                    sel = st.selectbox(label, source_cols, index=default_idx, key=f"map_{dyn_field}")
            else:
                with cols_right:
                    sel = st.selectbox(label, source_cols, index=default_idx, key=f"map_{dyn_field}")

            updated_map[dyn_field] = None if sel == "(not mapped)" else sel

        if st.button("💾 Save Mapping", type="primary"):
            st.session_state.column_map = updated_map
            st.success("Column mapping saved!")

        # Preview
        st.markdown("---")
        st.markdown("**Preview of first 5 rows from source file:**")
        st.dataframe(df.head(5), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─── Tab 3: Validate & Export ─────────────────────────────────────────────────
with tab3:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">✅ Validate & Export</div>', unsafe_allow_html=True)

    if st.session_state.df_raw is None:
        st.markdown("""
        <div class="warning-box">⬆️ Please upload a file in the <strong>Upload & Configure</strong> tab first.</div>
        """, unsafe_allow_html=True)
    else:
        if st.button("🔄 Run Validation & Build Export", type="primary", use_container_width=True):
            # Collect global settings
            seg = st.session_state.get("gs_market_segment", "")
            main_app_options = MAIN_APPLICATION_MAP.get(seg, [""])
            raw_main_app = st.session_state.get("gs_main_app", "")
            main_app_val = raw_main_app if raw_main_app in main_app_options else ""

            global_settings = {
                "Subject": st.session_state.get("gs_subject", default_subject),
                "Lead Source": st.session_state.get("gs_lead_source", ""),
                "Rating": st.session_state.get("gs_rating", ""),
                "Allow Marketing Communication": st.session_state.get("gs_allow_marketing", ""),
                "Market Segment": seg,
                "Main Application": main_app_val,
                "Industry Sector": st.session_state.get("gs_industry", ""),
                "Source Campaign": st.session_state.get("gs_campaign", ""),
                "Description": st.session_state.get("gs_description", ""),
            }

            col_map = {k: v for k, v in st.session_state.column_map.items()
                       if k != "Location"}

            with st.spinner("Processing rows…"):
                processed_df, errors = process_dataframe(
                    st.session_state.df_raw.copy(),
                    global_settings,
                    col_map,
                )

            st.session_state.processed_df = processed_df
            st.session_state.errors = errors
            st.rerun()

        if st.session_state.processed_df is not None:
            processed_df = st.session_state.processed_df
            errors = st.session_state.errors

            # Stats
            total = len(processed_df)
            error_rows = set()
            for e in errors:
                m = re.match(r'Row (\d+):', e)
                if m:
                    error_rows.add(int(m.group(1)))
            clean_rows = total - len(error_rows)

            st.markdown(f"""
            <div class="stats-row">
              <div class="stat-chip">
                <span class="stat-num">{total}</span>
                <span class="stat-label">Total Rows</span>
              </div>
              <div class="stat-chip" style="background:#F0FDF4;border-color:#86EFAC;">
                <span class="stat-num" style="color:#16A34A;">{clean_rows}</span>
                <span class="stat-label">Clean Rows</span>
              </div>
              <div class="stat-chip" style="background:#FFF1F2;border-color:#FECDD3;">
                <span class="stat-num" style="color:#DC2626;">{len(error_rows)}</span>
                <span class="stat-label">Rows w/ Issues</span>
              </div>
              <div class="stat-chip">
                <span class="stat-num">{len(errors)}</span>
                <span class="stat-label">Total Warnings</span>
              </div>
            </div>
            """, unsafe_allow_html=True)

            if not errors:
                st.markdown("""
                <div class="success-box">
                ✅ File successfully normalized and ready for Dynamics import.
                All mandatory fields are present and all values pass validation.
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="error-box">
                  <h4>⚠️ {len(errors)} Validation Issue{"s" if len(errors) != 1 else ""} Found</h4>
                  {"".join(f'<div class="error-item">• {e}</div>' for e in errors[:50])}
                  {"<div class='error-item' style='color:#9F1239;font-style:italic;'>...and more. Fix source file and re-run.</div>" if len(errors) > 50 else ""}
                </div>
                """, unsafe_allow_html=True)

            # Preview export
            st.markdown("---")
            st.markdown("**Preview of export (first 5 rows):**")
            st.dataframe(processed_df.head(5), use_container_width=True)

            # Export
            csv_buffer = io.StringIO()
            processed_df.to_csv(csv_buffer, index=False, encoding="utf-8")
            csv_bytes = csv_buffer.getvalue().encode("utf-8")

            st.download_button(
                label="⬇️ Download opalrt_dynamics_import.csv",
                data=csv_bytes,
                file_name="opalrt_dynamics_import.csv",
                mime="text/csv",
                use_container_width=True,
            )

    st.markdown('</div>', unsafe_allow_html=True)

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  Built by <strong>Arnaud Joakim</strong> &nbsp;·&nbsp;
  <a href="mailto:arnaud.joakim@opal-rt.com">arnaud.joakim@opal-rt.com</a>
  &nbsp;·&nbsp; OPAL-RT Technologies © 2025
</div>
""", unsafe_allow_html=True)
