"""
Opal RT Spreadsheet Cleaner
Production-ready Streamlit application for normalizing messy lead spreadsheets
into Microsoft Dynamics-compatible CSV imports.
"""

from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl  # noqa: F401 - required by pandas Excel engine and deployment spec
import pandas as pd
import streamlit as st

APP_TITLE = "Opal RT Spreadsheet Cleaner"
APP_SUBTITLE = "Prepare CRM-ready lead imports for Microsoft Dynamics"
HERO_IMAGE = "https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg"
FOOTER_EMAIL = "arnaud.joakim@opal-rt.com"

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

REQUIRED_FIELDS = ["Subject", "First Name", "Last Name", "Email", "Company Name", "Country"]

FIELD_LENGTHS = {
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

LEAD_SOURCE_OPTIONS = ["", "Web", "Prospection", "Webinar", "Referral", "Social Media", "Customer Portal", "SPS", "Others"]
RATING_OPTIONS = ["", "Cold", "Warm", "Hot"]
ALLOW_MARKETING_OPTIONS = ["", "Yes", "No"]
MARKET_SEGMENT_OPTIONS = ["", "Aerospace", "Automotive", "Energy Conversion", "Marine, Railway, Off-Highway", "Power System"]
INDUSTRY_SECTOR_OPTIONS = [
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
MAIN_APPLICATION_OPTIONS = {
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
ALL_MAIN_APPLICATIONS = sorted({v for values in MAIN_APPLICATION_OPTIONS.values() for v in values if v})

COLUMN_ALIASES = {
    "First Name": ["first name", "firstname", "fname", "given name", "givenname", "forename"],
    "Last Name": ["last name", "lastname", "surname", "lname", "family name", "familyname"],
    "Job Title": ["job title", "title", "position", "role", "designation", "function"],
    "Company Name": ["company", "company name", "organization", "organisation", "org", "account", "business", "employer"],
    "Email": ["email", "email address", "work email", "business email", "corporate email", "e-mail", "mail"],
    "Business Phone": ["phone", "telephone", "mobile", "mobile phone", "work phone", "business phone", "cell", "cell phone", "phone number"],
    "LinkedIn": ["linkedin", "linkedin profile", "linkedin profile url", "linkedin url", "linked in", "linkedin profile link"],
    "Location": ["location", "hq location", "office location", "city", "address", "region"],
    "Country": ["country", "country/region", "country region", "nation", "location"],
    "State or Province": ["state", "province", "state/province", "state or province", "region", "location"],
    "Market Segment": ["market segment", "segment", "market", "business segment"],
    "Main Application": ["main application", "application", "primary application", "use case"],
    "Industry Sector": ["industry sector", "industry", "sector", "vertical"],
    "Source Campaign": ["source campaign", "campaign", "campaign source"],
    "Lead Source": ["lead source", "source", "origin"],
    "Rating": ["rating", "lead rating", "temperature"],
    "Description": ["description", "notes", "note", "comments", "comment", "details"],
    "Allow Marketing Communication": ["allow marketing communication", "marketing consent", "opt in", "opt-in", "newsletter", "consent"],
}

COUNTRY_ALIASES = {
    "usa": "United States",
    "us": "United States",
    "u.s.": "United States",
    "u.s.a.": "United States",
    "united states of america": "United States",
    "united states": "United States",
    "america": "United States",
    "can": "Canada",
    "ca": "Canada",
    "canada": "Canada",
}
US_STATES = {
    "alabama": "Alabama", "al": "Alabama", "alaska": "Alaska", "ak": "Alaska", "arizona": "Arizona", "az": "Arizona",
    "arkansas": "Arkansas", "ar": "Arkansas", "california": "California", "ca": "California", "colorado": "Colorado", "co": "Colorado",
    "connecticut": "Connecticut", "ct": "Connecticut", "delaware": "Delaware", "de": "Delaware", "florida": "Florida", "fl": "Florida",
    "georgia": "Georgia", "ga": "Georgia", "hawaii": "Hawaii", "hi": "Hawaii", "idaho": "Idaho", "id": "Idaho",
    "illinois": "Illinois", "il": "Illinois", "indiana": "Indiana", "in": "Indiana", "iowa": "Iowa", "ia": "Iowa",
    "kansas": "Kansas", "ks": "Kansas", "kentucky": "Kentucky", "ky": "Kentucky", "louisiana": "Louisiana", "la": "Louisiana",
    "maine": "Maine", "me": "Maine", "maryland": "Maryland", "md": "Maryland", "massachusetts": "Massachusetts", "ma": "Massachusetts",
    "michigan": "Michigan", "mi": "Michigan", "minnesota": "Minnesota", "mn": "Minnesota", "mississippi": "Mississippi", "ms": "Mississippi",
    "missouri": "Missouri", "mo": "Missouri", "montana": "Montana", "mt": "Montana", "nebraska": "Nebraska", "ne": "Nebraska",
    "nevada": "Nevada", "nv": "Nevada", "new hampshire": "New Hampshire", "nh": "New Hampshire", "new jersey": "New Jersey", "nj": "New Jersey",
    "new mexico": "New Mexico", "nm": "New Mexico", "new york": "New York", "ny": "New York", "north carolina": "North Carolina", "nc": "North Carolina",
    "north dakota": "North Dakota", "nd": "North Dakota", "ohio": "Ohio", "oh": "Ohio", "oklahoma": "Oklahoma", "ok": "Oklahoma",
    "oregon": "Oregon", "or": "Oregon", "pennsylvania": "Pennsylvania", "pa": "Pennsylvania", "rhode island": "Rhode Island", "ri": "Rhode Island",
    "south carolina": "South Carolina", "sc": "South Carolina", "south dakota": "South Dakota", "sd": "South Dakota", "tennessee": "Tennessee", "tn": "Tennessee",
    "texas": "Texas", "tx": "Texas", "utah": "Utah", "ut": "Utah", "vermont": "Vermont", "vt": "Vermont",
    "virginia": "Virginia", "va": "Virginia", "washington": "Washington", "wa": "Washington", "west virginia": "West Virginia", "wv": "West Virginia",
    "wisconsin": "Wisconsin", "wi": "Wisconsin", "wyoming": "Wyoming", "wy": "Wyoming",
}
CANADIAN_PROVINCES = {
    "alberta": "Alberta", "ab": "Alberta", "british columbia": "British Columbia", "bc": "British Columbia",
    "manitoba": "Manitoba", "mb": "Manitoba", "new brunswick": "New Brunswick", "nb": "New Brunswick",
    "newfoundland and labrador": "Newfoundland and Labrador", "nl": "Newfoundland and Labrador",
    "northwest territories": "Northwest Territories", "nt": "Northwest Territories", "nova scotia": "Nova Scotia", "ns": "Nova Scotia",
    "nunavut": "Nunavut", "nu": "Nunavut", "ontario": "Ontario", "on": "Ontario", "prince edward island": "Prince Edward Island", "pe": "Prince Edward Island",
    "quebec": "Quebec", "québec": "Quebec", "qc": "Quebec", "saskatchewan": "Saskatchewan", "sk": "Saskatchewan", "yukon": "Yukon", "yt": "Yukon",
}
CITY_STATE_HINTS = {
    "montreal": ("Canada", "Quebec"), "montréal": ("Canada", "Quebec"), "quebec city": ("Canada", "Quebec"),
    "toronto": ("Canada", "Ontario"), "ottawa": ("Canada", "Ontario"), "vancouver": ("Canada", "British Columbia"),
    "calgary": ("Canada", "Alberta"), "edmonton": ("Canada", "Alberta"), "winnipeg": ("Canada", "Manitoba"),
    "dallas": ("United States", "Texas"), "houston": ("United States", "Texas"), "austin": ("United States", "Texas"),
    "san francisco": ("United States", "California"), "los angeles": ("United States", "California"), "san diego": ("United States", "California"),
    "new york": ("United States", "New York"), "chicago": ("United States", "Illinois"), "boston": ("United States", "Massachusetts"),
    "seattle": ("United States", "Washington"), "detroit": ("United States", "Michigan"), "atlanta": ("United States", "Georgia"),
}

EMAIL_RE = re.compile(r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$", re.IGNORECASE)


def normalize_key(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.replace("&", " and ")
    text = re.sub(r"[^A-Za-z0-9]+", " ", text).strip().lower()
    return re.sub(r"\s+", " ", text)


def clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).replace("\ufeff", "").replace("\u200b", "").replace("\xa0", " ")
    text = "".join(ch for ch in text if ch.isprintable())
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", " ", text).strip()


def clean_email(value: object) -> str:
    return clean_text(value).lower()


def clean_phone(value: object) -> str:
    phone = clean_text(value)
    return re.sub(r"\s+", " ", phone)


def titleish(value: object) -> str:
    # Normalize spaces and hidden characters without changing professional capitalization.
    return clean_text(value)


def is_empty_series(series: pd.Series) -> bool:
    return series.fillna("").astype(str).map(clean_text).eq("").all()


def remove_ghost_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [clean_text(c) for c in cleaned.columns]
    keep_cols = []
    for col in cleaned.columns:
        key = normalize_key(col)
        if not key or key.startswith("unnamed"):
            continue
        if is_empty_series(cleaned[col]):
            continue
        keep_cols.append(col)
    return cleaned.loc[:, keep_cols].copy()


def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    detected: Dict[str, str] = {}
    normalized_columns = {col: normalize_key(col) for col in df.columns}
    used_cols = set()

    for target, aliases in COLUMN_ALIASES.items():
        best_col = None
        best_score = 0.0
        alias_keys = [normalize_key(a) for a in aliases]
        for col, col_key in normalized_columns.items():
            if col in used_cols and target not in {"Country", "State or Province"}:
                continue
            scores = []
            for alias in alias_keys:
                if col_key == alias:
                    scores.append(1.0)
                elif alias and (alias in col_key or col_key in alias):
                    scores.append(0.91)
                else:
                    scores.append(SequenceMatcher(None, col_key, alias).ratio())
            score = max(scores) if scores else 0
            if score > best_score:
                best_score = score
                best_col = col
        if best_col and best_score >= 0.78:
            detected[target] = best_col
            if target not in {"Country", "State or Province"}:
                used_cols.add(best_col)
    return detected


def canonical_from_options(value: object, options: Iterable[str]) -> str:
    text = clean_text(value)
    if not text:
        return ""
    key = normalize_key(text)
    lookup = {normalize_key(opt): opt for opt in options if opt}
    if key in lookup:
        return lookup[key]
    best = ""
    best_score = 0.0
    for opt_key, opt in lookup.items():
        score = SequenceMatcher(None, key, opt_key).ratio()
        if key in opt_key or opt_key in key:
            score = max(score, 0.92)
        if score > best_score:
            best = opt
            best_score = score
    return best if best_score >= 0.88 else text


def infer_market_from_main_application(main_application: str) -> str:
    main_key = normalize_key(main_application)
    if not main_key:
        return ""
    for segment, values in MAIN_APPLICATION_OPTIONS.items():
        if not segment:
            continue
        for app in values:
            if normalize_key(app) == main_key:
                return segment
    return ""


def parse_location(value: object) -> Tuple[str, str]:
    text = clean_text(value)
    if not text:
        return "", ""

    pieces = [clean_text(p) for p in re.split(r"[,;/|]+", text) if clean_text(p)]
    lowered = [p.lower() for p in pieces]
    all_tokens = [normalize_key(p) for p in pieces]
    full_key = normalize_key(text)

    country = ""
    state = ""

    # Explicit country tokens first.
    for token in all_tokens + [full_key]:
        if token in COUNTRY_ALIASES:
            country = COUNTRY_ALIASES[token]
            break

    # Explicit state/province tokens.
    for token in all_tokens:
        if token in CANADIAN_PROVINCES:
            state = CANADIAN_PROVINCES[token]
            country = country or "Canada"
            break
        if token in US_STATES:
            state = US_STATES[token]
            country = country or "United States"
            break

    # Text can be "Dallas Texas United States" without commas.
    if not state:
        for province_key, province in CANADIAN_PROVINCES.items():
            if re.search(rf"\b{re.escape(province_key)}\b", full_key):
                state = province
                country = country or "Canada"
                break
    if not state:
        for state_key, state_name in US_STATES.items():
            if len(state_key) > 2 and re.search(rf"\b{re.escape(state_key)}\b", full_key):
                state = state_name
                country = country or "United States"
                break

    # Known city hints for obvious locations.
    if not country or not state:
        for city_key, (hint_country, hint_state) in CITY_STATE_HINTS.items():
            if city_key in full_key:
                country = country or hint_country
                state = state or hint_state
                break

    # If the last token is country and previous token is a known state/province.
    if len(all_tokens) >= 2:
        prev = all_tokens[-2]
        if not state and prev in CANADIAN_PROVINCES:
            state = CANADIAN_PROVINCES[prev]
            country = country or "Canada"
        if not state and prev in US_STATES:
            state = US_STATES[prev]
            country = country or "United States"

    return country, state


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        raw = uploaded_file.getvalue()
        for encoding in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(raw), dtype=str, encoding=encoding)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(io.BytesIO(raw), dtype=str)
    return pd.read_excel(uploaded_file, dtype=str, engine="openpyxl")


def default_subject() -> str:
    return f"{datetime.now():%Y%m}Prospection"


def build_normalized_export(df: pd.DataFrame, detected: Dict[str, str], settings: Dict[str, str]) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    output = pd.DataFrame(columns=EXPORT_COLUMNS)
    output["(Do Not Modify) Lead"] = [""] * len(df)
    output["(Do Not Modify) Row Checksum"] = [""] * len(df)
    output["(Do Not Modify) Modified On"] = [""] * len(df)

    for col in EXPORT_COLUMNS[3:]:
        output[col] = ""

    mappings: Dict[str, Optional[str]] = {field: detected.get(field) for field in EXPORT_COLUMNS if field in detected}

    direct_fields = [
        "First Name", "Last Name", "Job Title", "Company Name", "Email", "Business Phone",
        "Country", "State or Province", "LinkedIn", "Lead Source", "Rating", "Source Campaign",
        "Description", "Market Segment", "Main Application", "Industry Sector", "Allow Marketing Communication",
    ]
    for field in direct_fields:
        source_col = detected.get(field)
        if source_col and source_col in df.columns:
            if field == "Email":
                output[field] = df[source_col].map(clean_email)
            elif field == "Business Phone":
                output[field] = df[source_col].map(clean_phone)
            else:
                output[field] = df[source_col].map(titleish)

    # Location parsing should fill Country/State only where those fields are blank.
    loc_col = detected.get("Location")
    if loc_col and loc_col in df.columns:
        for idx, value in df[loc_col].items():
            parsed_country, parsed_state = parse_location(value)
            if parsed_country and not clean_text(output.at[idx, "Country"]):
                output.at[idx, "Country"] = parsed_country
            if parsed_state and not clean_text(output.at[idx, "State or Province"]):
                output.at[idx, "State or Province"] = parsed_state
        mappings["Location"] = loc_col

    # Canonicalize CRM option fields found in the source file.
    output["Lead Source"] = output["Lead Source"].map(lambda v: canonical_from_options(v, LEAD_SOURCE_OPTIONS))
    output["Rating"] = output["Rating"].map(lambda v: canonical_from_options(v, RATING_OPTIONS))
    output["Allow Marketing Communication"] = output["Allow Marketing Communication"].map(lambda v: canonical_from_options(v, ALLOW_MARKETING_OPTIONS))
    output["Market Segment"] = output["Market Segment"].map(lambda v: canonical_from_options(v, MARKET_SEGMENT_OPTIONS))
    output["Main Application"] = output["Main Application"].map(lambda v: canonical_from_options(v, ALL_MAIN_APPLICATIONS))
    output["Industry Sector"] = output["Industry Sector"].map(lambda v: canonical_from_options(v, INDUSTRY_SECTOR_OPTIONS))

    # Infer segment from a valid source main application only when source segment is empty.
    for idx in output.index:
        if not clean_text(output.at[idx, "Market Segment"]):
            inferred = infer_market_from_main_application(output.at[idx, "Main Application"])
            if inferred:
                output.at[idx, "Market Segment"] = inferred

    # Global settings. Blank classification dropdowns intentionally do not fill values.
    always_apply = ["Subject", "Lead Source", "Rating", "Allow Marketing Communication", "Source Campaign", "Description"]
    for field in always_apply:
        value = clean_text(settings.get(field, ""))
        if value:
            output[field] = value

    for field in ["Market Segment", "Main Application", "Industry Sector"]:
        value = clean_text(settings.get(field, ""))
        if value:
            output[field] = value

    # Subject is mandatory and has a default; populate it even when not provided by source.
    if not clean_text(settings.get("Subject", "")):
        output["Subject"] = default_subject()

    # Final cleanup, preserving exact column order.
    for col in output.columns:
        if col == "Email":
            output[col] = output[col].map(clean_email)
        elif col == "Business Phone":
            output[col] = output[col].map(clean_phone)
        else:
            output[col] = output[col].map(clean_text)

    if "Email" in output.columns:
        before = len(output)
        output = output.drop_duplicates(subset=["Email"], keep="first", ignore_index=True)
        output.attrs["duplicates_removed"] = before - len(output)

    return output[EXPORT_COLUMNS], mappings


def validate_export(df: pd.DataFrame) -> List[str]:
    errors: List[str] = []
    for row_idx, row in df.iterrows():
        display_row = row_idx + 2  # Spreadsheet-like row number after header.
        for field in REQUIRED_FIELDS:
            if not clean_text(row.get(field, "")):
                errors.append(f"Row {display_row}: Missing required field → {field}")

        email = clean_text(row.get("Email", ""))
        if email and not EMAIL_RE.match(email):
            errors.append(f"Row {display_row}: Invalid email → {email}")

        for field, limit in FIELD_LENGTHS.items():
            value = clean_text(row.get(field, ""))
            if value and len(value) > limit:
                errors.append(f"Row {display_row}: {field} exceeds {limit} characters")

        segment = clean_text(row.get("Market Segment", ""))
        main_app = clean_text(row.get("Main Application", ""))
        industry = clean_text(row.get("Industry Sector", ""))
        if segment and segment not in MARKET_SEGMENT_OPTIONS:
            errors.append(f"Row {display_row}: Invalid Market Segment → {segment}")
        if industry and industry not in INDUSTRY_SECTOR_OPTIONS:
            errors.append(f"Row {display_row}: Invalid Industry Sector → {industry}")
        if main_app:
            valid_apps = MAIN_APPLICATION_OPTIONS.get(segment, [""])
            if segment and main_app not in valid_apps:
                errors.append(f"Row {display_row}: Main Application does not match Market Segment → {main_app}")
            elif not segment and main_app not in ALL_MAIN_APPLICATIONS:
                errors.append(f"Row {display_row}: Invalid Main Application → {main_app}")

    return errors


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --opal-blue: #0057A8; --opal-navy: #071D49; --opal-sky: #EAF4FF; --opal-gray: #F6F8FB; }
        .stApp { background: linear-gradient(180deg, #f8fbff 0%, #ffffff 42%, #f7f9fc 100%); }
        [data-testid="stHeader"] { background: rgba(248, 251, 255, 0.82); backdrop-filter: blur(8px); }
        .block-container { max-width: 1220px; padding-top: 1.2rem; padding-bottom: 3rem; }
        .hero {
            position: relative; overflow: hidden; border-radius: 28px; min-height: 310px; padding: 44px;
            background-image: linear-gradient(90deg, rgba(7,29,73,.94), rgba(0,87,168,.72), rgba(0,87,168,.18)), url('https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg');
            background-size: cover; background-position: center; box-shadow: 0 24px 70px rgba(7,29,73,.22); color: white;
        }
        .hero h1 { font-size: clamp(2.2rem, 5vw, 4.2rem); line-height: 1.02; margin: 0 0 12px 0; font-weight: 800; letter-spacing: -.04em; }
        .hero p { font-size: 1.25rem; max-width: 680px; opacity: .94; margin: 0; }
        .hero-badge { display: inline-flex; gap: 8px; align-items: center; padding: 8px 13px; border-radius: 999px; background: rgba(255,255,255,.16); border: 1px solid rgba(255,255,255,.24); margin-bottom: 22px; font-weight: 650; }
        .metric-card, .soft-card { border: 1px solid #e6ecf5; background: rgba(255,255,255,.92); border-radius: 22px; padding: 20px; box-shadow: 0 12px 32px rgba(7,29,73,.07); }
        .metric-card strong { color: var(--opal-blue); font-size: 1.85rem; display:block; }
        .section-title { font-size: 1.28rem; font-weight: 760; color: var(--opal-navy); margin: 14px 0 6px; }
        .required-star { color: #D92D20; font-weight: 800; }
        .footer { text-align:center; padding: 26px 0 12px; color: #5b667a; }
        .footer a { color: var(--opal-blue); font-weight: 700; text-decoration: none; }
        div.stButton > button, div.stDownloadButton > button { border-radius: 14px !important; border: 1px solid #0057A8 !important; background: #0057A8 !important; color: white !important; font-weight: 720 !important; min-height: 44px; box-shadow: 0 10px 22px rgba(0,87,168,.22); }
        div.stButton > button:hover, div.stDownloadButton > button:hover { background: #004987 !important; border-color: #004987 !important; }
        [data-testid="stFileUploader"] section { border-radius: 20px; border: 1.5px dashed #8dbbe8; background: #f6fbff; }
        div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, textarea { border-radius: 12px !important; }
        .stAlert { border-radius: 16px; }
        hr { border: none; border-top: 1px solid #e6ecf5; margin: 1.4rem 0; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero() -> None:
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-badge">OPAL-RT • Dynamics Import Utility</div>
            <h1>{APP_TITLE}</h1>
            <p>{APP_SUBTITLE}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_global_settings() -> Dict[str, str]:
    st.markdown('<div class="section-title">Global Import Settings</div>', unsafe_allow_html=True)
    st.caption("Values selected here are applied to every exported row. Classification fields can remain blank.")

    col1, col2, col3 = st.columns(3)
    with col1:
        subject = st.text_input("Subject *", value=default_subject(), help="Required. Default format: YYYYMMProspection.")
        lead_source = st.selectbox("Lead Source", LEAD_SOURCE_OPTIONS, index=LEAD_SOURCE_OPTIONS.index("Prospection"))
        market_segment = st.selectbox("Market Segment", MARKET_SEGMENT_OPTIONS, index=0)
    with col2:
        rating = st.selectbox("Rating", RATING_OPTIONS, index=0)
        allow_marketing = st.selectbox("Allow Marketing Communication", ALLOW_MARKETING_OPTIONS, index=0)
        main_options = MAIN_APPLICATION_OPTIONS.get(market_segment, [""])
        main_application = st.selectbox("Main Application", main_options, index=0)
    with col3:
        industry_sector = st.selectbox("Industry Sector", INDUSTRY_SECTOR_OPTIONS, index=0)
        source_campaign = st.text_input("Source Campaign", value="")
        description = st.text_area("Description", value="", height=92, max_chars=2000)

    return {
        "Subject": subject,
        "Lead Source": lead_source,
        "Rating": rating,
        "Allow Marketing Communication": allow_marketing,
        "Market Segment": market_segment,
        "Main Application": main_application,
        "Industry Sector": industry_sector,
        "Source Campaign": source_campaign,
        "Description": description,
    }


def render_mapping_summary(mappings: Dict[str, Optional[str]], detected: Dict[str, str]) -> None:
    rows = []
    for target in ["First Name", "Last Name", "Company Name", "Job Title", "Email", "Business Phone", "LinkedIn", "Location", "Country", "State or Province", "Market Segment", "Main Application", "Industry Sector"]:
        rows.append({"Dynamics Field": target, "Detected Source Column": mappings.get(target) or detected.get(target) or "—"})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🧹", layout="wide")
    inject_css()
    render_hero()

    st.write("")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="metric-card"><strong>CSV/XLSX</strong>Accepted upload formats</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="metric-card"><strong>21</strong>Exact Dynamics export columns</div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="metric-card"><strong>Pre-check</strong>CRM validation before import</div>', unsafe_allow_html=True)

    st.write("")
    settings = render_global_settings()

    st.markdown("---")
    st.markdown('<div class="section-title">Upload CSV or Excel File</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload CSV or Excel File", type=["csv", "xlsx"], label_visibility="collapsed")

    if not uploaded:
        st.info("Upload a lead spreadsheet to automatically clean, map, validate, and export it for Microsoft Dynamics.")
        st.markdown(
            """
            **Mandatory fields:** Subject*, First Name*, Last Name*, Email*, Company Name*, Country*  
            **Important:** if your source has a `Location` column, the app will use it to infer Country and State/Province when obvious.
            """
        )
    else:
        try:
            source_df = read_uploaded_file(uploaded)
            source_df = remove_ghost_columns(source_df)
            for col in source_df.columns:
                source_df[col] = source_df[col].map(clean_text)
        except Exception as exc:  # pragma: no cover - displayed in UI
            st.error(f"Could not read the uploaded file: {exc}")
            return

        if source_df.empty or len(source_df.columns) == 0:
            st.error("The uploaded file does not contain usable data after removing empty/unnamed columns.")
            return

        detected = detect_columns(source_df)
        normalized_df, mappings = build_normalized_export(source_df, detected, settings)
        errors = validate_export(normalized_df)
        duplicates_removed = int(normalized_df.attrs.get("duplicates_removed", 0))

        st.markdown('<div class="section-title">Detected Source Mapping</div>', unsafe_allow_html=True)
        render_mapping_summary(mappings, detected)

        left, right = st.columns([1.1, 1.4])
        with left:
            st.markdown('<div class="section-title">Validation Results</div>', unsafe_allow_html=True)
            if duplicates_removed:
                st.warning(f"Removed {duplicates_removed} duplicate row(s) by email, keeping the first occurrence.")
            if errors:
                st.error("Validation errors found. Fix these before importing into Dynamics.")
                st.code("\n".join(errors[:300]), language="text")
                if len(errors) > 300:
                    st.caption(f"Showing first 300 of {len(errors)} errors.")
            else:
                st.success("File successfully normalized and ready for Dynamics import.")

            st.download_button(
                "Download Dynamics-ready CSV",
                data=to_csv_bytes(normalized_df),
                file_name="opalrt_dynamics_import.csv",
                mime="text/csv",
                disabled=bool(errors),
                use_container_width=True,
            )
            if errors:
                st.caption("Download is disabled until validation errors are resolved.")

        with right:
            st.markdown('<div class="section-title">Dynamics Export Preview</div>', unsafe_allow_html=True)
            st.dataframe(normalized_df.head(100), use_container_width=True, hide_index=True)

        with st.expander("View cleaned source preview"):
            st.dataframe(source_df.head(100), use_container_width=True, hide_index=True)

    st.markdown(
        f"""
        <div class="footer">
            Built by <strong>Arnaud Joakim</strong> · <a href="mailto:{FOOTER_EMAIL}">{FOOTER_EMAIL}</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
