"""
Opal RT Spreadsheet Cleaner
Production-ready Streamlit application for converting messy lead spreadsheets
into Microsoft Dynamics-compatible CSV import files.
"""

from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl  # noqa: F401 - required for Excel support on Streamlit Cloud
import pandas as pd
import streamlit as st

APP_TITLE = "Opal RT Spreadsheet Cleaner"
APP_SUBTITLE = "Prepare CRM-ready lead imports for Microsoft Dynamics"
HERO_IMAGE_URL = "https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg"
CONTACT_EMAIL = "arnaud.joakim@opal-rt.com"

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

LEAD_SOURCE_VALUES = ["Web", "Prospection", "Webinar", "Referral", "Social Media", "Customer Portal", "SPS", "Others"]
RATING_VALUES = ["Cold", "Warm", "Hot"]
ALLOW_MARKETING_VALUES = ["Yes", "No"]
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
MARKET_SEGMENT_APPLICATIONS = {
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
MARKET_SEGMENT_VALUES = list(MARKET_SEGMENT_APPLICATIONS.keys())
ALL_MAIN_APPLICATIONS = sorted({v for values in MARKET_SEGMENT_APPLICATIONS.values() for v in values if v})

COLUMN_ALIASES = {
    "First Name": ["first name", "firstname", "fname", "given name", "givenname", "forename"],
    "Last Name": ["last name", "lastname", "surname", "lname", "family name", "familyname"],
    "Company Name": ["company", "company name", "organization", "organisation", "org", "account", "business", "employer"],
    "Job Title": ["job title", "title", "position", "role", "designation", "function"],
    "Email": ["email", "email address", "work email", "business email", "corporate email", "e-mail", "mail"],
    "Business Phone": ["phone", "telephone", "mobile", "mobile phone", "work phone", "business phone", "cell", "cell phone", "phone number"],
    "LinkedIn": ["linkedin", "linkedin profile", "linkedin profile url", "linkedin url", "linked in", "linkedin profile link"],
    "Location": ["location", "hq location", "office location", "city", "address", "region", "territory"],
    "Country": ["country", "country/region", "country region", "nation"],
    "State or Province": ["state", "province", "state/province", "state or province", "region", "territory"],
    "Market Segment": ["market segment", "segment", "market", "business segment"],
    "Main Application": ["main application", "application", "primary application", "use case"],
    "Industry Sector": ["industry sector", "industry", "sector", "vertical"],
    "Source Campaign": ["source campaign", "campaign", "campaign source"],
    "Lead Source": ["lead source", "source", "origin"],
    "Rating": ["rating", "lead rating", "temperature"],
    "Description": ["description", "notes", "note", "comments", "comment", "details"],
    "Allow Marketing Communication": ["allow marketing communication", "marketing consent", "opt in", "opt-in", "newsletter", "consent"],
}
LOCATION_HEADER_HINTS = ["location", "country", "country region", "state", "province", "state province", "state or province", "city", "hq location", "office location", "address", "region", "territory"]

COUNTRY_ALIASES = {
    "usa": "United States",
    "us": "United States",
    "u s": "United States",
    "u s a": "United States",
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
    "nunavut": "Nunavut", "nu": "Nunavut", "ontario": "Ontario", "on": "Ontario", "prince edward island": "Prince Edward Island", "pei": "Prince Edward Island", "pe": "Prince Edward Island",
    "quebec": "Quebec", "québec": "Quebec", "qc": "Quebec", "saskatchewan": "Saskatchewan", "sk": "Saskatchewan", "yukon": "Yukon", "yt": "Yukon",
}
CITY_HINTS = {
    "montreal": ("Canada", "Quebec"), "montréal": ("Canada", "Quebec"), "quebec city": ("Canada", "Quebec"),
    "toronto": ("Canada", "Ontario"), "ottawa": ("Canada", "Ontario"), "vancouver": ("Canada", "British Columbia"),
    "calgary": ("Canada", "Alberta"), "edmonton": ("Canada", "Alberta"), "winnipeg": ("Canada", "Manitoba"),
    "dallas": ("United States", "Texas"), "houston": ("United States", "Texas"), "austin": ("United States", "Texas"),
    "san francisco": ("United States", "California"), "los angeles": ("United States", "California"), "san diego": ("United States", "California"),
    "new york": ("United States", "New York"), "chicago": ("United States", "Illinois"), "boston": ("United States", "Massachusetts"),
    "seattle": ("United States", "Washington"), "detroit": ("United States", "Michigan"), "atlanta": ("United States", "Georgia"),
}

EMAIL_RE = re.compile(r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$", re.IGNORECASE)
HIDDEN_CHARS_RE = re.compile(r"[\u0000-\u001F\u007F\u200B\u200C\u200D\uFEFF]")
MULTISPACE_RE = re.compile(r"\s+")


def clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = HIDDEN_CHARS_RE.sub(" ", text).replace("\xa0", " ")
    return MULTISPACE_RE.sub(" ", text).strip()


def normalize_header(value: object) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.replace("&", " and ").lower()
    text = re.sub(r"[_\-\/]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return MULTISPACE_RE.sub(" ", text).strip()


def clean_email(value: object) -> str:
    return clean_text(value).lower()


def remove_ghost_columns(df: pd.DataFrame) -> pd.DataFrame:
    keep_cols = []
    for col in df.columns:
        header = clean_text(col)
        normalized_header = normalize_header(col)
        values_empty = df[col].map(clean_text).eq("").all()
        if header and not normalized_header.startswith("unnamed") and not values_empty:
            keep_cols.append(col)
    return df.loc[:, keep_cols].copy()


def score_alias(header_norm: str, alias_norm: str) -> int:
    if not header_norm or not alias_norm:
        return 0
    if header_norm == alias_norm:
        return 100
    if header_norm.replace(" ", "") == alias_norm.replace(" ", ""):
        return 96
    if alias_norm in header_norm:
        return 86
    h_tokens = set(header_norm.split())
    a_tokens = set(alias_norm.split())
    if a_tokens and a_tokens.issubset(h_tokens):
        return 80
    overlap = len(h_tokens & a_tokens)
    if overlap:
        return min(64, 24 * overlap)
    return int(SequenceMatcher(None, header_norm, alias_norm).ratio() * 54)


def detect_source_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    normalized_columns = {col: normalize_header(col) for col in df.columns}
    used: set[str] = set()
    mapping: Dict[str, Optional[str]] = {}
    priority = [
        "Email", "First Name", "Last Name", "Company Name", "Job Title", "Business Phone", "LinkedIn",
        "Country", "State or Province", "Location", "Market Segment", "Main Application", "Industry Sector",
        "Source Campaign", "Lead Source", "Rating", "Description", "Allow Marketing Communication",
    ]
    for target in priority:
        best_col = None
        best_score = 0
        for col, col_norm in normalized_columns.items():
            if col in used:
                continue
            for alias in COLUMN_ALIASES[target]:
                score = score_alias(col_norm, normalize_header(alias))
                if score > best_score:
                    best_col = col
                    best_score = score
        mapping[target] = best_col if best_score >= 60 else None
        if mapping[target]:
            used.add(mapping[target])
    return mapping


def likely_location_columns(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> List[str]:
    candidates: List[str] = []
    for target in ["Location", "Country", "State or Province"]:
        col = mapping.get(target)
        if col and col not in candidates:
            candidates.append(col)
    for col in df.columns:
        norm = normalize_header(col)
        if any(score_alias(norm, hint) >= 70 for hint in LOCATION_HEADER_HINTS) and col not in candidates:
            candidates.append(col)
    return candidates


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    suffix = uploaded_file.name.lower().rsplit(".", 1)[-1]
    data = uploaded_file.getvalue()
    if suffix == "csv":
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(data), dtype=str, encoding=encoding)
            except UnicodeDecodeError as exc:
                last_error = exc
        raise last_error or ValueError("Unable to decode CSV file.")
    if suffix == "xlsx":
        return pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl")
    raise ValueError("Unsupported file type. Please upload a .csv or .xlsx file.")


def value_from_mapping(df: pd.DataFrame, mapping: Dict[str, Optional[str]], target: str) -> pd.Series:
    col = mapping.get(target)
    if col and col in df.columns:
        return df[col].map(clean_email if target == "Email" else clean_text)
    return pd.Series([""] * len(df), index=df.index, dtype="object")


def canonical_choice(value: object, choices: Iterable[str]) -> str:
    text = clean_text(value)
    if not text:
        return ""
    choice_list = list(choices)
    norm_to_choice = {normalize_header(choice): choice for choice in choice_list}
    norm = normalize_header(text)
    if norm in norm_to_choice:
        return norm_to_choice[norm]
    best_choice = ""
    best_score = 0.0
    for choice in choice_list:
        if not choice:
            continue
        score = SequenceMatcher(None, norm, normalize_header(choice)).ratio()
        if score > best_score:
            best_choice = choice
            best_score = score
    return best_choice if best_score >= 0.86 else ""


def detect_country_from_text(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    pieces = [normalize_header(p) for p in re.split(r"[,|;/()\[\]\n]", text) if clean_text(p)]
    joined = normalize_header(text)
    for piece in reversed(pieces):
        if piece in COUNTRY_ALIASES:
            return COUNTRY_ALIASES[piece]
    for alias, country in sorted(COUNTRY_ALIASES.items(), key=lambda item: len(item[0]), reverse=True):
        if re.search(rf"\b{re.escape(alias)}\b", joined):
            return country
    return ""


def detect_state_from_text(value: object, country: str = "") -> str:
    text = clean_text(value)
    if not text:
        return ""
    parts = [normalize_header(p) for p in re.split(r"[,|;/()\[\]\n]", text) if clean_text(p)]
    joined = normalize_header(text)
    country_norm = normalize_header(country)
    maps = []
    if country_norm in {"canada", "ca"}:
        maps = [CANADIAN_PROVINCES]
    elif country_norm in {"united states", "united states of america", "usa", "us"}:
        maps = [US_STATES]
    else:
        maps = [CANADIAN_PROVINCES, US_STATES]
    for mapping in maps:
        for part in parts:
            if part in mapping:
                return mapping[part]
    for mapping in maps:
        for key, label in sorted(mapping.items(), key=lambda item: len(item[0]), reverse=True):
            if len(key) > 2 and re.search(rf"\b{re.escape(key)}\b", joined):
                return label
    return ""


def infer_country_from_state_or_province(value: object) -> str:
    key = normalize_header(value)
    if not key:
        return ""
    if key in CANADIAN_PROVINCES or any(normalize_header(v) == key for v in CANADIAN_PROVINCES.values()):
        return "Canada"
    if key in US_STATES or any(normalize_header(v) == key for v in US_STATES.values()):
        return "United States"
    return ""


def parse_location(value: object) -> Tuple[str, str]:
    text = clean_text(value)
    if not text:
        return "", ""
    country = detect_country_from_text(text)
    state = detect_state_from_text(text, country)
    if not country and state:
        country = infer_country_from_state_or_province(state)
    if not country and not state:
        normalized = normalize_header(text)
        for city, result in CITY_HINTS.items():
            if re.search(rf"\b{re.escape(normalize_header(city))}\b", normalized):
                country, state = result
                break
    return country, state


def fill_location_fields(output: pd.DataFrame, df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    country_col = mapping.get("Country")
    state_col = mapping.get("State or Province")
    if country_col and country_col in df.columns:
        output["Country"] = df[country_col].map(clean_text)
    if state_col and state_col in df.columns:
        output["State or Province"] = df[state_col].map(clean_text)

    for col in likely_location_columns(df, mapping):
        parsed = df[col].map(parse_location)
        countries = parsed.map(lambda pair: pair[0])
        states = parsed.map(lambda pair: pair[1])
        output["Country"] = output["Country"].where(output["Country"].astype(str).str.strip().ne(""), countries)
        output["State or Province"] = output["State or Province"].where(output["State or Province"].astype(str).str.strip().ne(""), states)

    # If a state/province column exists without a country column, infer only when the state/province is recognized.
    inferred_from_state = output["State or Province"].map(infer_country_from_state_or_province)
    output["Country"] = output["Country"].where(output["Country"].astype(str).str.strip().ne(""), inferred_from_state)
    return output


def canonical_market_segment(value: object) -> str:
    return canonical_choice(value, MARKET_SEGMENT_VALUES)


def canonical_main_application(value: object, market_segment: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    segment = canonical_market_segment(market_segment)
    allowed = MARKET_SEGMENT_APPLICATIONS.get(segment, [""])
    if segment:
        return canonical_choice(text, allowed)
    # If no segment is supplied, allow a match against any known application without inventing a segment.
    return canonical_choice(text, ALL_MAIN_APPLICATIONS)


def create_normalized_export(df_raw: pd.DataFrame, global_settings: Dict[str, str]) -> Tuple[pd.DataFrame, Dict[str, Optional[str]], int, List[str]]:
    df = remove_ghost_columns(df_raw).copy()
    df.columns = [clean_text(c) for c in df.columns]
    mapping = detect_source_columns(df)
    output = pd.DataFrame("", index=df.index, columns=EXPORT_COLUMNS, dtype="object")

    for field in [
        "First Name", "Last Name", "Job Title", "Company Name", "Email", "Business Phone", "Description", "LinkedIn",
        "Market Segment", "Main Application", "Industry Sector", "Source Campaign", "Lead Source", "Rating", "Allow Marketing Communication",
    ]:
        output[field] = value_from_mapping(df, mapping, field)

    output = fill_location_fields(output, df, mapping)

    for col in output.columns:
        output[col] = output[col].map(clean_email if col == "Email" else clean_text)

    # Canonicalize source-provided optional dropdowns. Blank or unmatched values remain blank.
    output["Market Segment"] = output["Market Segment"].map(canonical_market_segment)
    output["Industry Sector"] = output["Industry Sector"].map(lambda v: canonical_choice(v, INDUSTRY_SECTOR_VALUES))

    # Mandatory/global settings. These may be intentionally applied to every row.
    output["Subject"] = clean_text(global_settings.get("Subject", ""))
    output["Lead Source"] = clean_text(global_settings.get("Lead Source", "")) or output["Lead Source"]
    output["Rating"] = clean_text(global_settings.get("Rating", "")) or output["Rating"]
    output["Allow Marketing Communication"] = clean_text(global_settings.get("Allow Marketing Communication", "")) or output["Allow Marketing Communication"]

    # Optional settings must not overwrite source values unless the user explicitly entered/selected a non-blank value.
    for field in ["Source Campaign", "Description"]:
        selected = clean_text(global_settings.get(field, ""))
        if selected:
            output[field] = selected

    selected_market_segment = clean_text(global_settings.get("Market Segment", ""))
    selected_industry_sector = clean_text(global_settings.get("Industry Sector", ""))
    selected_main_application = clean_text(global_settings.get("Main Application", ""))
    if selected_market_segment:
        output["Market Segment"] = canonical_market_segment(selected_market_segment)
    if selected_industry_sector:
        output["Industry Sector"] = canonical_choice(selected_industry_sector, INDUSTRY_SECTOR_VALUES)
    if selected_main_application:
        output["Main Application"] = selected_main_application
    else:
        output["Main Application"] = [
            canonical_main_application(app, seg) for app, seg in zip(output["Main Application"], output["Market Segment"])
        ]

    before = len(output)
    has_email = output["Email"].astype(str).str.strip().ne("")
    duplicate_mask = has_email & output.duplicated(subset=["Email"], keep="first")
    output = output.loc[~duplicate_mask].reset_index(drop=True)
    duplicates_removed = before - len(output)

    location_cols = likely_location_columns(df, mapping)
    return output[EXPORT_COLUMNS], mapping, duplicates_removed, location_cols


def validate_export(df: pd.DataFrame) -> List[str]:
    errors: List[str] = []
    for idx, row in df.iterrows():
        row_number = idx + 2
        for field in REQUIRED_FIELDS:
            if clean_text(row.get(field, "")) == "":
                errors.append(f"Row {row_number}: Missing required field → {field}")

        email = clean_email(row.get("Email", ""))
        if email and not EMAIL_RE.match(email):
            errors.append(f"Row {row_number}: Invalid email → {email}")

        for field, max_len in FIELD_LENGTHS.items():
            value = clean_text(row.get(field, ""))
            if len(value) > max_len:
                errors.append(f"Row {row_number}: {field} exceeds {max_len} characters")

        market_segment = clean_text(row.get("Market Segment", ""))
        main_application = clean_text(row.get("Main Application", ""))
        if market_segment and market_segment not in MARKET_SEGMENT_APPLICATIONS:
            errors.append(f"Row {row_number}: Invalid Market Segment → {market_segment}")
        if main_application:
            valid_apps = MARKET_SEGMENT_APPLICATIONS.get(market_segment, [""])
            if market_segment and main_application not in valid_apps:
                errors.append(f"Row {row_number}: Main Application '{main_application}' is not valid for Market Segment '{market_segment}'")
            if not market_segment and main_application not in ALL_MAIN_APPLICATIONS:
                errors.append(f"Row {row_number}: Invalid Main Application → {main_application}")

        industry = clean_text(row.get("Industry Sector", ""))
        if industry and industry not in INDUSTRY_SECTOR_VALUES:
            errors.append(f"Row {row_number}: Invalid Industry Sector → {industry}")

        lead_source = clean_text(row.get("Lead Source", ""))
        if lead_source and lead_source not in LEAD_SOURCE_VALUES:
            errors.append(f"Row {row_number}: Invalid Lead Source → {lead_source}")

        rating = clean_text(row.get("Rating", ""))
        if rating and rating not in RATING_VALUES:
            errors.append(f"Row {row_number}: Invalid Rating → {rating}")

        allow = clean_text(row.get("Allow Marketing Communication", ""))
        if allow and allow not in ALLOW_MARKETING_VALUES:
            errors.append(f"Row {row_number}: Invalid Allow Marketing Communication → {allow}")
    return errors


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8").encode("utf-8")


def render_metric(label: str, value: object) -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title=APP_TITLE, page_icon="🔷", layout="wide")
st.markdown(
    f"""
    <style>
    :root {{
        --opal-blue: #005BAA;
        --opal-navy: #092A49;
        --opal-sky: #00A3E0;
        --opal-light: #F4F8FC;
        --opal-border: #D8E7F5;
        --opal-text: #172033;
    }}
    .stApp {{ background: linear-gradient(180deg, #F7FBFF 0%, #FFFFFF 42%, #F6F9FC 100%); color: var(--opal-text); }}
    .hero {{
        min-height: 285px;
        border-radius: 28px;
        overflow: hidden;
        margin: 0.25rem 0 1.6rem 0;
        background-image: linear-gradient(90deg, rgba(4, 31, 57, .92), rgba(0, 91, 170, .72), rgba(0, 163, 224, .22)), url('{HERO_IMAGE_URL}');
        background-size: cover;
        background-position: center;
        box-shadow: 0 20px 55px rgba(9, 42, 73, 0.16);
    }}
    .hero-content {{ padding: 3.1rem 3.4rem; max-width: 900px; }}
    .eyebrow {{
        display: inline-flex; align-items: center; gap: .5rem; padding: .35rem .75rem; border-radius: 999px;
        background: rgba(255,255,255,.16); color: #EAF7FF; border: 1px solid rgba(255,255,255,.25);
        font-size: .78rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase;
    }}
    .hero h1 {{ margin: 1rem 0 .35rem 0; font-size: 3.25rem; line-height: 1.02; color: white; font-weight: 800; }}
    .hero p {{ color: #E8F5FF; font-size: 1.22rem; margin: 0; max-width: 760px; }}
    .card {{ background: rgba(255,255,255,.92); border: 1px solid var(--opal-border); border-radius: 22px; padding: 1.1rem 1.2rem; box-shadow: 0 10px 28px rgba(9, 42, 73, 0.07); }}
    .metric-card {{ background: #FFFFFF; border: 1px solid var(--opal-border); border-radius: 18px; padding: 1rem 1.1rem; min-height: 95px; box-shadow: 0 8px 24px rgba(9, 42, 73, 0.06); }}
    .metric-label {{ color: #55708A; font-size: .78rem; text-transform: uppercase; letter-spacing: .06em; font-weight: 800; }}
    .metric-value {{ color: var(--opal-navy); font-size: 1.9rem; font-weight: 850; margin-top: .1rem; }}
    .footer {{ margin-top: 3rem; padding: 1.25rem 0 1.5rem 0; border-top: 1px solid #DDEBF6; text-align: center; color: #5B7084; font-size: .94rem; }}
    .footer a {{ color: var(--opal-blue); text-decoration: none; font-weight: 800; }}
    div.stButton > button, div.stDownloadButton > button {{ background: linear-gradient(90deg, #005BAA 0%, #008DD2 100%) !important; color: white !important; border: 0 !important; border-radius: 999px !important; padding: .72rem 1.25rem !important; font-weight: 800 !important; box-shadow: 0 8px 18px rgba(0, 91, 170, .22) !important; }}
    [data-testid="stFileUploader"] {{ border: 1px dashed #8ABDE7; border-radius: 20px; background: #F7FBFF; padding: 1rem; }}
    .mandatory-note {{ color: #57708A; font-size: .9rem; margin-bottom: 0.7rem; }}
    .stAlert {{ border-radius: 16px; }}
    </style>
    <div class="hero"><div class="hero-content"><div class="eyebrow">OPAL-RT · Dynamics CRM Import Utility</div><h1>{APP_TITLE}</h1><p>{APP_SUBTITLE}</p></div></div>
    """,
    unsafe_allow_html=True,
)

st.markdown("### Global Import Settings")
st.markdown('<div class="mandatory-note">Fields marked with * are mandatory for Dynamics import.</div>', unsafe_allow_html=True)

row1 = st.columns([1.4, 1, 1, 1])
with row1[0]:
    subject = st.text_input("Subject *", value=f"{datetime.now():%Y%m}Prospection", max_chars=300)
with row1[1]:
    lead_source = st.selectbox("Lead Source", LEAD_SOURCE_VALUES, index=LEAD_SOURCE_VALUES.index("Prospection"))
with row1[2]:
    rating = st.selectbox("Rating", RATING_VALUES, index=0)
with row1[3]:
    allow_marketing = st.selectbox("Allow Marketing Communication", ALLOW_MARKETING_VALUES, index=0)

row2 = st.columns([1, 1, 1, 1])
with row2[0]:
    market_segment = st.selectbox("Market Segment", MARKET_SEGMENT_VALUES, index=0)
with row2[1]:
    main_application = st.selectbox("Main Application", MARKET_SEGMENT_APPLICATIONS[market_segment], index=0)
with row2[2]:
    industry_sector = st.selectbox("Industry Sector", INDUSTRY_SECTOR_VALUES, index=0)
with row2[3]:
    source_campaign = st.text_input("Source Campaign", value="")
description = st.text_area("Description", value="", max_chars=2000, height=86)

st.divider()
left, right = st.columns([1, 1], gap="large")
with left:
    st.markdown("### Upload CSV or Excel File")
    uploaded_file = st.file_uploader("Accepted formats: .csv, .xlsx", type=["csv", "xlsx"])
with right:
    st.markdown("### Mandatory CRM Fields")
    st.markdown(
        """
        <div class="card">
        <b>Required:</b> Subject *, First Name *, Last Name *, Email *, Company Name *, Country *<br><br>
        <b>Location logic:</b> Any source column that looks like Location, Country, State, Province, City, Address, Region, or Territory is scanned for obvious country/state/province data.<br><br>
        <b>Optional CRM classifications:</b> Market Segment, Main Application, and Industry Sector remain blank unless supplied by the source file or selected above.
        </div>
        """,
        unsafe_allow_html=True,
    )

if uploaded_file:
    try:
        raw_df = read_uploaded_file(uploaded_file)
        normalized_df, detected_mapping, duplicates_removed, location_cols = create_normalized_export(
            raw_df,
            {
                "Subject": subject,
                "Lead Source": lead_source,
                "Rating": rating,
                "Allow Marketing Communication": allow_marketing,
                "Market Segment": market_segment,
                "Main Application": main_application,
                "Industry Sector": industry_sector,
                "Source Campaign": source_campaign,
                "Description": description,
            },
        )
        validation_errors = validate_export(normalized_df)

        metric_cols = st.columns(4)
        with metric_cols[0]:
            render_metric("Source Rows", len(raw_df))
        with metric_cols[1]:
            render_metric("Export Rows", len(normalized_df))
        with metric_cols[2]:
            render_metric("Duplicates Removed", duplicates_removed)
        with metric_cols[3]:
            render_metric("Validation Errors", len(validation_errors))

        st.markdown("### Source Column Detection")
        mapping_rows = []
        display_fields = ["First Name", "Last Name", "Company Name", "Job Title", "Email", "Business Phone", "LinkedIn", "Country", "State or Province", "Location", "Description", "Market Segment", "Main Application", "Industry Sector"]
        for target in display_fields:
            mapping_rows.append({"Dynamics Field": target, "Detected Source Column": detected_mapping.get(target) or "—"})
        st.dataframe(pd.DataFrame(mapping_rows), use_container_width=True, hide_index=True)
        if location_cols:
            st.caption("Location-aware columns scanned: " + ", ".join(str(c) for c in location_cols))

        if validation_errors:
            st.error("Validation errors found. Fix these before importing into Dynamics.")
            with st.expander("View row-level validation errors", expanded=True):
                for err in validation_errors[:500]:
                    st.write(f"• {err}")
                if len(validation_errors) > 500:
                    st.write(f"… and {len(validation_errors) - 500} more errors.")
        else:
            st.success("File successfully normalized and ready for Dynamics import.")

        st.markdown("### Dynamics-Ready Preview")
        st.dataframe(normalized_df.head(100), use_container_width=True, hide_index=True)
        st.download_button(
            label="Download Dynamics CSV",
            data=to_csv_bytes(normalized_df),
            file_name="opalrt_dynamics_import.csv",
            mime="text/csv",
            disabled=bool(validation_errors),
        )
        if validation_errors:
            st.info("The download button is disabled until validation errors are resolved.")
    except Exception as exc:
        st.error(f"Unable to process this file: {exc}")
else:
    st.info("Upload a CSV or XLSX file to start normalization and validation.")

st.markdown(
    f"""
    <div class="footer">Built by Arnaud Joakim · <a href="mailto:{CONTACT_EMAIL}">{CONTACT_EMAIL}</a></div>
    """,
    unsafe_allow_html=True,
)
