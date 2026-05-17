"""
Opal RT Spreadsheet Cleaner
Prepare CRM-ready lead imports for Microsoft Dynamics.

Built by Arnaud Joakim <arnaud.joakim@opal-rt.com>
"""
from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st

# =============================================================================
# CONSTANTS — derived from ImportLeadTemplate.xlsm
# =============================================================================

PAGE_TITLE = "Opal RT Spreadsheet Cleaner"
PAGE_SUBTITLE = "Prepare CRM-ready lead imports for Microsoft Dynamics"
HERO_IMAGE_URL = "https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg"

BRAND_PRIMARY = "#003E7E"
BRAND_ACCENT = "#0066CC"
BRAND_BG = "#F5F7FA"
BRAND_INK = "#0E1A2B"
BRAND_MUTED = "#5A6B82"

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
    "Allow Marketing Communication",
]

MANDATORY_FIELDS = ["Subject", "First Name", "Last Name", "Email", "Company Name", "Country"]

MAX_LENGTHS = {
    "Subject": 300,
    "First Name": 58,
    "Last Name": 50,
    "Company Name": 100,
    "Job Title": 100,
    "Email": 100,
    "Business Phone": 50,
    "LinkedIn": 500,
    "Description": 2000,
}

LEAD_SOURCES = ["", "Shows", "Web", "Prospection", "Webinar", "Referral",
                "Social Media", "Customer Portal", "SPS", "Others"]
RATINGS = ["", "Cold", "Warm", "Hot"]
ALLOW_MM_OPTIONS = ["", "Yes", "No"]

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
    "Stock - Inventory",
]

MARKET_SEGMENTS = ["", "Aerospace", "Automotive", "Energy Conversion",
                   "Marine, Railway, Off-Highway", "Power System"]

MAIN_APPS_BY_SEGMENT: dict[str, list[str]] = {
    "": [""],
    "Aerospace": [
        "", "Autonomous Systems (Aero)", "Avionics System",
        "Electrical Actuators and Servos", "EVTOL", "More Electrical Aircraft",
        "Onboard System", "Other (if nothing fits) Aero", "Propulsion and APU",
        "Testbench - Test Automation and Monitoring from RTS",
    ],
    "Automotive": [
        "", "Autonomous Systems (Auto)", "Body & Chassis", "Charging",
        "EV/HEV Powertrain", "Full Vehicle Simulation", "ICE Powertrain",
        "Other (if nothing fits) Auto",
    ],
    "Energy Conversion": [
        "", "Autonomous Systems (Energy Conversion)", "Backup Power (UPS)",
        "Inverter/Converter", "Medium and Large Drive (>150KW)",
        "Other (if nothing fits) EnergyConversion", "Small Drive (<150KW)",
    ],
    "Marine, Railway, Off-Highway": [
        "", "Autonomous Systems (Marine, Railway, Off-Highway)", "BMS Control",
        "Grid Infrastructure", "Onboard Power System",
        "Other (if nothing fits) Marine, Railway, Off-Highway", "Propulsion Control",
    ],
    "Power System": [
        "", "Autonomous Systems (Power Systems)", "Conventional Generation",
        "Converter-Based Energy Resource", "Distribution", "FACTS & HVDC",
        "Microgrid", "Other (if nothing fits) PowerSystem", "Substation", "Transmission",
    ],
}

# Full Dynamics country list (exact spellings from the template)
COUNTRIES = [
    "Afghanistan", "African Country (non-maghrebian)", "Aland Island", "Albania", "Algeria",
    "American Samoa", "Andorra", "Angola", "Anguilla", "Antartica", "Antigua and Barbuda",
    "Argentina", "Armenia", "Aruba", "Australia", "Austria", "Azerbaijan", "Bahamas",
    "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin",
    "Bermuda", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Bouvet Island",
    "Brazil", "British Indian Ocean Territory", "Brunei Darussalam", "Bulgaria", "Burkina Faso",
    "Burundi", "Cambodia", "Cameroon", "Canada", "Cape Verde", "Cayman Islands",
    "Central African Republic", "Chad", "Chile", "China", "Chrismas Island",
    "Cocos (Keeling) Islands", "Colombia", "Comoros", "Congo",
    "Congo, The democatic Republic of the", "Cook Islands", "Costa Rica", "Croatia", "Cuba",
    "Cyprus", "Czech Republic", "Denmark", "Djibouti", "Dominica", "Dominican Republic",
    "Egypt", "El Salvador", "Ecuador", "Equatorial Guinea", "Eritrea", "Estonia",
    "Ethiopia", "Falkland Islands (Malvinas)", "Faroe Island", "Fiji", "Finland", "France",
    "French Guiana", "French Polynesia", "French Southern Territories", "French-Guadeloupe",
    "French-Martinique", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Gibraltar",
    "Greece", "Greenland", "Grenada", "Guam", "Guatemala", "Guernser", "Guinea",
    "Guinea-Bissau", "Guyana", "Haiti", "Heard Island and McDonals Islands", "Honduras",
    "Hong Kong", "Hungary", "Iceland", "India", "Indonesia", "Iran (Islamic Republic of)",
    "Iraq", "Ireland", "Isle of Man", "Israel", "Italy", "Ivory Coast", "Jamaica", "Japan",
    "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan", "Lakshadweep",
    "Lao People's Democratic republic", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya",
    "Liechtenstein", "Lithuania", "Luxembourg", "Macao", "Macedonia", "Madagascar",
    "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania",
    "Mauritius", "Mayotte", "Mexico", "Micronesia, Federated States of", "Moldova",
    "Monaco", "Mongolia", "Montenegro", "Montserrat", "Morocco", "Mozambique", "Myanmar",
    "N/A", "Namibia", "Nepal", "Netherlands", "Netherlands Antilles", "New Caledonia",
    "New Zealand", "Nicaragua", "Niger", "Nigeria", "Niue", "Norfolk Iskand",
    "Northern Mariana Islands", "Norway", "Oman", "Pakistan", "Palau", "Palestine",
    "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Pitcairn", "Poland",
    "Portugal", "Puerto Rico", "Qatar", "Reunion", "Romania", "Russia", "Rwanda",
    "Saint Barthelemy", "Saint Helena", "Saint Kitts and Nevis", "Saint Lucia",
    "Saint Pierre and Miquelon", "Saint Vincent and the Grenadines", "Samoa", "San Marino",
    "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Shanghai",
    "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia",
    "South Africa", "South Georgia and the South Sandwich Islands", "South Korea", "Spain",
    "Sri Lanka", "St Martin", "Sudan", "Suriname", "Svalbard and Jan Mayen", "Swaziland",
    "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand",
    "Timor-Leste", "Togo", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan",
    "Turks and Caicos Island", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates",
    "United Kingdom", "United States", "Uruguay", "Uzbekistan", "Vanuatu",
    "Vatican City State", "Venezuela", "Vietnam", "Virgin Islands, British",
    "Virgin Islands, U.S.", "Wallis and Futuna", "Western Sahara", "Yemen", "Zambia",
    "Zimbabwe",
]

US_STATES = [
    "Alabama", "Alaska", "American Samoa", "Arizona", "Arkansas", "California",
    "Colorado", "Connecticut", "Delaware", "District of Columbia", "Florida", "Georgia",
    "Guam", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky",
    "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
    "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire",
    "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota",
    "Northern Mariana Islands", "Ohio", "Oklahoma", "Oregon", "Pennsylvania",
    "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee",
    "Texas", "United States Minor Outlying Islands", "Utah", "Vermont",
    "Virgin Islands, U.S.", "Virginia", "Washington", "West Virginia", "Wisconsin",
    "Wyoming",
]

CA_PROVINCES = [
    "Alberta", "British Columbia", "Manitoba", "New Brunswick",
    "Newfoundland and Labrador", "Northwest Territories", "Nova Scotia", "Nunavut",
    "Ontario", "Prince Edward Island", "Québec", "Saskatchewan", "Yukon Territory",
]

US_STATE_ABBR = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois",
    "IN": "Indiana", "IA": "Iowa", "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana",
    "ME": "Maine", "MD": "Maryland", "MA": "Massachusetts", "MI": "Michigan",
    "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri", "MT": "Montana",
    "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey",
    "NM": "New Mexico", "NY": "New York", "NC": "North Carolina", "ND": "North Dakota",
    "OH": "Ohio", "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania",
    "RI": "Rhode Island", "SC": "South Carolina", "SD": "South Dakota",
    "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont", "VA": "Virginia",
    "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
    "DC": "District of Columbia", "PR": "Puerto Rico", "GU": "Guam",
    "VI": "Virgin Islands, U.S.",
}

CA_PROV_ABBR = {
    "AB": "Alberta", "BC": "British Columbia", "MB": "Manitoba", "NB": "New Brunswick",
    "NL": "Newfoundland and Labrador", "NF": "Newfoundland and Labrador",
    "NS": "Nova Scotia", "NT": "Northwest Territories", "NU": "Nunavut",
    "ON": "Ontario", "PE": "Prince Edward Island", "PEI": "Prince Edward Island",
    "QC": "Québec", "QU": "Québec", "PQ": "Québec",
    "SK": "Saskatchewan", "YT": "Yukon Territory", "YK": "Yukon Territory",
}

# Common country aliases → canonical Dynamics country name
COUNTRY_ALIASES = {
    "usa": "United States", "us": "United States", "u.s.": "United States",
    "u.s.a.": "United States", "united states of america": "United States",
    "america": "United States", "united states": "United States",
    "uk": "United Kingdom", "u.k.": "United Kingdom", "great britain": "United Kingdom",
    "britain": "United Kingdom", "england": "United Kingdom", "scotland": "United Kingdom",
    "wales": "United Kingdom", "northern ireland": "United Kingdom",
    "united kingdom": "United Kingdom",
    "uae": "United Arab Emirates", "u.a.e.": "United Arab Emirates",
    "united arab emirates": "United Arab Emirates",
    "korea": "South Korea", "republic of korea": "South Korea",
    "south korea": "South Korea",
    "russia": "Russia", "russian federation": "Russia",
    "czechia": "Czech Republic", "czech republic": "Czech Republic",
    "iran": "Iran (Islamic Republic of)",
    "p.r.c.": "China", "prc": "China", "china": "China",
    "hong kong": "Hong Kong", "hk": "Hong Kong",
    "macao": "Macao", "macau": "Macao",
    "côte d'ivoire": "Ivory Coast", "cote d'ivoire": "Ivory Coast",
    "ivory coast": "Ivory Coast",
    "the netherlands": "Netherlands", "holland": "Netherlands",
    "netherlands": "Netherlands",
    "viet nam": "Vietnam", "vietnam": "Vietnam",
    "the philippines": "Philippines", "philippines": "Philippines",
    "deutschland": "Germany", "germany": "Germany",
    "españa": "Spain", "espana": "Spain", "spain": "Spain",
    "italia": "Italy", "italy": "Italy",
    "brasil": "Brazil", "brazil": "Brazil",
    "méxico": "Mexico", "mexico": "Mexico",
}

# Regions / provinces / states (non-US/CA) → inferred country.
# Used ONLY when no explicit country is present in the location string.
# Kept conservative — ambiguous names (e.g. Georgia, Victoria) are omitted.
REGION_TO_COUNTRY = {
    # Germany
    "bavaria": "Germany", "bayern": "Germany",
    "baden-württemberg": "Germany", "baden-wurttemberg": "Germany",
    "brandenburg": "Germany", "hamburg": "Germany",
    "hesse": "Germany", "hessen": "Germany",
    "lower saxony": "Germany", "niedersachsen": "Germany",
    "mecklenburg-vorpommern": "Germany",
    "north rhine-westphalia": "Germany", "nordrhein-westfalen": "Germany",
    "rhineland-palatinate": "Germany", "rheinland-pfalz": "Germany",
    "saarland": "Germany", "saxony": "Germany", "sachsen": "Germany",
    "saxony-anhalt": "Germany", "schleswig-holstein": "Germany",
    "thuringia": "Germany", "thüringen": "Germany",
    # France
    "île-de-france": "France", "ile-de-france": "France",
    "auvergne-rhône-alpes": "France", "auvergne-rhone-alpes": "France",
    "bretagne": "France", "brittany": "France",
    "normandie": "France", "normandy": "France",
    "occitanie": "France",
    "provence-alpes-côte d'azur": "France", "paca": "France",
    "nouvelle-aquitaine": "France",
    "centre-val de loire": "France",
    "bourgogne-franche-comté": "France", "bourgogne-franche-comte": "France",
    "grand est": "France", "pays de la loire": "France",
    "hauts-de-france": "France", "corse": "France", "corsica": "France",
    # UK constituent countries already handled in COUNTRY_ALIASES
    # Spain
    "catalonia": "Spain", "cataluña": "Spain", "catalunya": "Spain",
    "andalusia": "Spain", "andalucía": "Spain",
    "comunidad valenciana": "Spain", "valencian community": "Spain",
    "galicia": "Spain", "basque country": "Spain", "país vasco": "Spain",
    "pais vasco": "Spain",
    # Italy
    "lombardy": "Italy", "lombardia": "Italy",
    "lazio": "Italy", "tuscany": "Italy", "toscana": "Italy",
    "piedmont": "Italy", "piemonte": "Italy",
    "veneto": "Italy", "sicily": "Italy", "sicilia": "Italy",
    "sardinia": "Italy", "sardegna": "Italy", "emilia-romagna": "Italy",
    # Australia
    "new south wales": "Australia",
    "queensland": "Australia", "south australia": "Australia",
    "western australia": "Australia", "tasmania": "Australia",
    "northern territory": "Australia",
    "australian capital territory": "Australia", "act": "Australia",
    # India
    "maharashtra": "India", "karnataka": "India", "tamil nadu": "India",
    "gujarat": "India", "kerala": "India",
    "west bengal": "India", "uttar pradesh": "India", "punjab": "India",
    "haryana": "India", "rajasthan": "India", "telangana": "India",
    "andhra pradesh": "India",
    # Switzerland
    "zürich": "Switzerland", "zurich": "Switzerland",
    "geneva": "Switzerland", "genève": "Switzerland",
    "vaud": "Switzerland",
    # Brazil
    "são paulo": "Brazil", "sao paulo": "Brazil",
    "rio de janeiro": "Brazil", "minas gerais": "Brazil",
    # Netherlands
    "noord-holland": "Netherlands", "north holland": "Netherlands",
    "zuid-holland": "Netherlands", "south holland": "Netherlands",
    "utrecht": "Netherlands",
    # Belgium
    "flanders": "Belgium", "wallonia": "Belgium",
    "vlaanderen": "Belgium",
}

# Synonym groups for source column detection. Keys are canonical fields.
COLUMN_SYNONYMS = {
    "First Name": ["firstname", "first name", "fname", "given name", "givenname",
                   "prenom", "prénom", "first"],
    "Last Name": ["lastname", "last name", "lname", "surname", "family name",
                  "familyname", "nom", "last"],
    "Company Name": ["company name", "companyname", "company", "organization",
                     "organisation", "org", "employer", "entreprise", "société",
                     "societe", "account name", "accountname"],
    "Job Title": ["job title", "jobtitle", "title", "position", "role",
                  "poste", "fonction", "job"],
    "Email": ["email", "email address", "emailaddress", "work email", "business email",
              "corporate email", "professional email", "courriel", "mail", "e-mail",
              "primary email"],
    "Business Phone": ["business phone", "businessphone", "work phone", "workphone",
                       "phone", "telephone", "tel", "mobile", "mobile phone",
                       "mobilephone", "cell", "cell phone", "cellphone",
                       "office phone", "primary phone"],
    "LinkedIn": ["linkedin", "linkedin profile", "linkedin profile url",
                 "linkedin url", "li profile", "li url", "linkedin link"],
    "Country": ["country", "country/region", "country region", "countryregion",
                "nation", "pays"],
    "State or Province": ["state", "province", "state/province", "state or province",
                          "stateprovince", "region", "état", "etat"],
    "Location": ["location", "address", "full address", "city/state", "city, state",
                 "geo", "place"],
    "Description": ["description", "notes", "comments", "comment", "remarks", "remark"],
}

# =============================================================================
# HELPERS
# =============================================================================

EMAIL_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._%+\-]*@[A-Za-z0-9](?:[A-Za-z0-9\-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]*[A-Za-z0-9])?)*\.[A-Za-z]{2,63}$")
MOJIBAKE_HINT_RE = re.compile(r"Ã[\x80-\xBF]|Â[\x80-\xBF]|â€|�")
CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\uFEFF]")
MULTI_SPACE_RE = re.compile(r"\s+")


def fix_mojibake(s: str) -> str:
    """Repair common UTF-8-decoded-as-Latin1 corruption (Ã©→é, etc.)."""
    if not isinstance(s, str) or not s:
        return s
    if MOJIBAKE_HINT_RE.search(s):
        try:
            fixed = s.encode("latin-1", errors="strict").decode("utf-8", errors="strict")
            return fixed
        except (UnicodeDecodeError, UnicodeEncodeError):
            pass
    return s


def clean_text(value: Any) -> str:
    """Normalize a cell value: fix mojibake, strip control chars, collapse whitespace.

    If a replacement character (�) remains after mojibake repair, drop it
    rather than leaving garbage in the export.
    """
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    s = str(value)
    s = fix_mojibake(s)
    s = unicodedata.normalize("NFC", s)
    s = ZERO_WIDTH_RE.sub("", s)
    s = CONTROL_CHARS_RE.sub("", s)
    s = s.replace("\ufffd", "")  # leftover replacement char
    s = MULTI_SPACE_RE.sub(" ", s).strip()
    return s


def norm_header(s: Any) -> str:
    """Canonical form of a column header: lowercase alphanumeric only."""
    if s is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def norm_syn(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


# Pre-compute normalized synonym → canonical field map
_SYN_LOOKUP: dict[str, str] = {}
for canonical, syns in COLUMN_SYNONYMS.items():
    for syn in syns:
        _SYN_LOOKUP[norm_syn(syn)] = canonical


def detect_columns(df: pd.DataFrame) -> dict[str, str]:
    """Map canonical field → source column name. First match wins per field."""
    mapping: dict[str, str] = {}
    for src_col in df.columns:
        nh = norm_header(src_col)
        if not nh:
            continue
        canonical = _SYN_LOOKUP.get(nh)
        if canonical and canonical not in mapping:
            mapping[canonical] = src_col
    # Loose contains-fallback for headers that didn't exactly match
    used = set(mapping.values())
    for src_col in df.columns:
        if src_col in used:
            continue
        nh = norm_header(src_col)
        if not nh:
            continue
        for syn_norm, canonical in _SYN_LOOKUP.items():
            if canonical in mapping:
                continue
            # match if normalized header contains the synonym as a substring
            if syn_norm and syn_norm in nh:
                mapping[canonical] = src_col
                used.add(src_col)
                break
    return mapping


def is_ghost_column(series: pd.Series, name: Any) -> bool:
    """True if a column header is missing/unnamed or all values are empty."""
    if name is None:
        return True
    name_str = str(name).strip()
    if not name_str or name_str.lower().startswith("unnamed:") or name_str.lower() == "nan":
        return True
    cleaned = series.dropna().astype(str).map(str.strip)
    return cleaned.eq("").all() if not cleaned.empty else True


def valid_email(e: str) -> bool:
    if not e:
        return False
    if len(e) > 100:
        return False
    return bool(EMAIL_RE.match(e))


# Build fast lookup sets for country/state matching
_COUNTRY_BY_NORM = {norm_syn(c): c for c in COUNTRIES}
_COUNTRY_ALIAS_BY_NORM = {norm_syn(k): v for k, v in COUNTRY_ALIASES.items()}
_US_STATE_BY_NORM = {norm_syn(s): s for s in US_STATES}
_CA_PROV_BY_NORM = {norm_syn(p): p for p in CA_PROVINCES}
# Quebec without accent → Québec
_CA_PROV_BY_NORM[norm_syn("Quebec")] = "Québec"
_US_ABBR_BY_NORM = {norm_syn(k): v for k, v in US_STATE_ABBR.items()}
_CA_ABBR_BY_NORM = {norm_syn(k): v for k, v in CA_PROV_ABBR.items()}
_REGION_BY_NORM = {norm_syn(k): v for k, v in REGION_TO_COUNTRY.items()}


def match_country(token: str) -> str:
    """Return canonical country name or empty string."""
    n = norm_syn(token)
    if not n:
        return ""
    if n in _COUNTRY_BY_NORM:
        return _COUNTRY_BY_NORM[n]
    if n in _COUNTRY_ALIAS_BY_NORM:
        return _COUNTRY_ALIAS_BY_NORM[n]
    return ""


def match_us_state(token: str) -> str:
    n = norm_syn(token)
    if not n:
        return ""
    if n in _US_STATE_BY_NORM:
        return _US_STATE_BY_NORM[n]
    # 2-letter abbreviation must match exactly (avoid partial collisions)
    upper = token.strip().upper().rstrip(".")
    if upper in US_STATE_ABBR:
        return US_STATE_ABBR[upper]
    return ""


def match_ca_prov(token: str) -> str:
    n = norm_syn(token)
    if not n:
        return ""
    if n in _CA_PROV_BY_NORM:
        return _CA_PROV_BY_NORM[n]
    upper = token.strip().upper().rstrip(".")
    if upper in CA_PROV_ABBR:
        return CA_PROV_ABBR[upper]
    return ""


def match_region_country(token: str) -> str:
    """If a non-US/CA region name, return inferred country."""
    n = norm_syn(token)
    return _REGION_BY_NORM.get(n, "")


def parse_location(loc: str) -> tuple[str, str]:
    """Return (Country, State_or_Province) from a free-form location string.

    State/Province is only populated when Country is United States or Canada.
    Returns ("", "") if nothing confident can be parsed.
    """
    loc = clean_text(loc)
    if not loc:
        return "", ""

    # Split on commas, slashes, pipes, and semicolons
    parts = [p.strip() for p in re.split(r"[,/|;]", loc) if p.strip()]
    if not parts:
        return "", ""

    # 1) Try to find an explicit country, scanning right-to-left (country usually last)
    country = ""
    country_idx = -1
    for i in range(len(parts) - 1, -1, -1):
        c = match_country(parts[i])
        if c:
            country = c
            country_idx = i
            break

    state = ""

    if country == "United States":
        for i in range(len(parts)):
            if i == country_idx:
                continue
            s = match_us_state(parts[i])
            if s:
                state = s
                break
    elif country == "Canada":
        for i in range(len(parts)):
            if i == country_idx:
                continue
            s = match_ca_prov(parts[i])
            if s:
                state = s
                break
    elif country:
        # Other country: never populate State/Province
        state = ""
    else:
        # 2) No explicit country found. Try to infer from a US state / CA province.
        for p in parts:
            s = match_us_state(p)
            if s:
                country, state = "United States", s
                break
        if not country:
            for p in parts:
                s = match_ca_prov(p)
                if s:
                    country, state = "Canada", s
                    break
        # 3) Still nothing? Try non-US/CA region inference.
        if not country:
            for p in parts:
                inferred = match_region_country(p)
                if inferred:
                    country = inferred
                    break

    return country, state


def read_uploaded_file(uploaded) -> pd.DataFrame:
    """Read a CSV or XLSX upload into a DataFrame, dtype=str, trying encodings."""
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    if name.endswith(".csv"):
        last_err = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(data), dtype=str, encoding=enc,
                                   keep_default_na=False, na_values=[""])
            except (UnicodeDecodeError, pd.errors.ParserError) as e:
                last_err = e
                continue
        raise ValueError(f"Could not parse CSV (last error: {last_err})")
    elif name.endswith(".xlsx") or name.endswith(".xlsm"):
        return pd.read_excel(io.BytesIO(data), dtype=str,
                             engine="openpyxl", keep_default_na=False, na_values=[""])
    else:
        raise ValueError(f"Unsupported file type: {uploaded.name}")


def drop_ghost_columns(df: pd.DataFrame) -> pd.DataFrame:
    keep = [c for c in df.columns if not is_ghost_column(df[c], c)]
    return df[keep]


def coerce_industry_sector(value: str) -> str:
    """Map a free-text industry value to a canonical Industry Sector if possible."""
    if not value:
        return ""
    n = norm_syn(value)
    for canonical in INDUSTRY_SECTORS:
        if canonical and norm_syn(canonical) == n:
            return canonical
    # Loose substring match
    for canonical in INDUSTRY_SECTORS:
        if canonical and norm_syn(canonical) in n:
            return canonical
    return ""


def coerce_market_segment(value: str) -> str:
    if not value:
        return ""
    n = norm_syn(value)
    for canonical in MARKET_SEGMENTS:
        if canonical and norm_syn(canonical) == n:
            return canonical
    return ""


# =============================================================================
# CORE PROCESSING
# =============================================================================

def process(df: pd.DataFrame, settings: dict[str, str]) -> tuple[pd.DataFrame, dict[str, str], list[dict], int]:
    """Build the final Dynamics-shaped DataFrame.

    Returns: (final_df, column_mapping, errors, dropped_no_email_count)
    """
    # 1) Drop ghost / unnamed / empty columns
    df = drop_ghost_columns(df)

    # 2) Clean every cell (mojibake, whitespace, control chars)
    for col in df.columns:
        df[col] = df[col].map(clean_text)

    # 3) Auto-detect source columns
    mapping = detect_columns(df)

    # 4) Build final rows
    rows: list[dict[str, str]] = []
    for _, src in df.iterrows():
        out: dict[str, str] = {c: "" for c in FINAL_COLUMNS}

        # Dynamics-managed columns stay empty
        # Global settings — applied to every row
        out["Subject"] = settings.get("Subject", "")
        out["Lead Source"] = settings.get("Lead Source", "")
        out["Rating"] = settings.get("Rating", "")
        out["Source Campaign"] = settings.get("Source Campaign", "")
        out["Description"] = settings.get("Description", "")
        out["Allow Marketing Communication"] = settings.get("Allow Marketing Communication", "")
        out["Market Segment"] = settings.get("Market Segment", "")
        out["Main Application"] = settings.get("Main Application", "")
        out["Industry Sector"] = settings.get("Industry Sector", "")

        def src_val(canonical: str) -> str:
            col = mapping.get(canonical)
            if col is None or col not in src.index:
                return ""
            return clean_text(src[col])

        # Simple mapped fields
        out["First Name"] = src_val("First Name")
        out["Last Name"] = src_val("Last Name")
        out["Company Name"] = src_val("Company Name")
        out["Job Title"] = src_val("Job Title")
        out["LinkedIn"] = src_val("LinkedIn")
        out["Business Phone"] = src_val("Business Phone")

        # Email — lowercase
        email = src_val("Email")
        out["Email"] = email.lower() if email else ""

        # Description — only overwrite global if source has a value
        src_desc = src_val("Description")
        if src_desc:
            out["Description"] = src_desc

        # Country / State / Location — source data overrides global blanks; never
        # silently overwrites a non-empty user-set value with weaker data.
        src_country = src_val("Country")
        src_state = src_val("State or Province")
        src_location = src_val("Location")

        country = ""
        state = ""

        # Prefer an explicit Country column, but validate it
        if src_country:
            matched = match_country(src_country)
            if matched:
                country = matched
            else:
                # Maybe the "Country" column is actually a state/region
                if match_us_state(src_country):
                    country = "United States"
                    if not src_state:
                        state = match_us_state(src_country)
                elif match_ca_prov(src_country):
                    country = "Canada"
                    if not src_state:
                        state = match_ca_prov(src_country)
                else:
                    inferred = match_region_country(src_country)
                    if inferred:
                        country = inferred
                    # Otherwise leave country blank — never put a non-country here.

        # State column (only kept if US/CA)
        if src_state and not state:
            if country == "United States":
                state = match_us_state(src_state)
            elif country == "Canada":
                state = match_ca_prov(src_state)
            elif not country:
                # Try state → infer country
                us = match_us_state(src_state)
                if us:
                    country, state = "United States", us
                else:
                    ca = match_ca_prov(src_state)
                    if ca:
                        country, state = "Canada", ca

        # Location column — fill any gaps
        if src_location:
            loc_country, loc_state = parse_location(src_location)
            if not country and loc_country:
                country = loc_country
            if not state and country in ("United States", "Canada") and loc_state:
                state = loc_state

        # Final guarantee: State only allowed for US/CA
        if country not in ("United States", "Canada"):
            state = ""

        out["Country"] = country
        out["State or Province"] = state

        # Industry sector / market segment — if source has it, use it (only if valid)
        # but never silently overwrite a user dropdown selection.
        if not out["Industry Sector"]:
            for guess_col in ("industry sector", "industry", "sector", "industrysector"):
                for src_col in src.index:
                    if norm_header(src_col) == norm_syn(guess_col):
                        coerced = coerce_industry_sector(clean_text(src[src_col]))
                        if coerced:
                            out["Industry Sector"] = coerced
                        break
                if out["Industry Sector"]:
                    break

        if not out["Market Segment"]:
            for guess_col in ("market segment", "marketsegment", "segment"):
                for src_col in src.index:
                    if norm_header(src_col) == norm_syn(guess_col):
                        coerced = coerce_market_segment(clean_text(src[src_col]))
                        if coerced:
                            out["Market Segment"] = coerced
                        break
                if out["Market Segment"]:
                    break

        rows.append(out)

    final = pd.DataFrame(rows, columns=FINAL_COLUMNS)

    # 5) Drop rows with no email at all, then dedupe by email
    no_email_mask = final["Email"].astype(str).str.strip().eq("")
    dropped_no_email = int(no_email_mask.sum())
    final = final[~no_email_mask].reset_index(drop=True)

    # Deduplicate by email (keep first)
    before = len(final)
    final = final.drop_duplicates(subset=["Email"], keep="first").reset_index(drop=True)
    deduped = before - len(final)

    # 6) Validate
    errors = validate(final)

    # Stash dedupe info on errors for the UI
    if deduped:
        errors.append({
            "row": "—",
            "field": "Email",
            "type": "info",
            "message": f"Removed {deduped} duplicate row(s) by email.",
        })

    return final, mapping, errors, dropped_no_email


def validate(df: pd.DataFrame) -> list[dict]:
    """Per-row validation. Returns list of error dicts.

    Row numbers are 1-based as shown to the user (after the header row).
    """
    errs: list[dict] = []
    for i, row in df.iterrows():
        row_num = i + 2  # +1 for 1-based, +1 for header row

        # Mandatory fields
        for f in MANDATORY_FIELDS:
            val = str(row.get(f, "") or "").strip()
            if not val:
                errs.append({
                    "row": row_num,
                    "field": f,
                    "type": "missing",
                    "message": f"Missing required field → {f}",
                })

        # Email format
        email = str(row.get("Email", "") or "").strip()
        if email and not valid_email(email):
            errs.append({
                "row": row_num,
                "field": "Email",
                "type": "invalid_email",
                "message": f"Invalid email → {email}",
            })

        # Country must be a known country
        country = str(row.get("Country", "") or "").strip()
        if country and country not in COUNTRIES:
            errs.append({
                "row": row_num,
                "field": "Country",
                "type": "invalid_country",
                "message": f"Unrecognized country → {country}",
            })

        # Field length limits
        for f, limit in MAX_LENGTHS.items():
            val = str(row.get(f, "") or "")
            if len(val) > limit:
                errs.append({
                    "row": row_num,
                    "field": f,
                    "type": "length",
                    "message": f"{f} exceeds {limit} characters ({len(val)})",
                })
    return errs


# =============================================================================
# UI
# =============================================================================

st.set_page_config(page_title=PAGE_TITLE, page_icon="🟦", layout="wide",
                   initial_sidebar_state="collapsed")

# Minimal CSS — only touches the chrome, never the widget internals
st.markdown(f"""
<style>
.stApp {{ background-color: {BRAND_BG}; }}
.block-container {{ padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1180px; }}

/* Hero banner */
.opal-hero {{
    position: relative;
    border-radius: 16px;
    overflow: hidden;
    background: linear-gradient(135deg, {BRAND_PRIMARY} 0%, {BRAND_ACCENT} 100%);
    margin-bottom: 1.8rem;
    box-shadow: 0 6px 24px rgba(0, 62, 126, 0.18);
}}
.opal-hero-img {{
    background-image: linear-gradient(135deg, rgba(0,62,126,0.88), rgba(0,102,204,0.72)),
                      url('{HERO_IMAGE_URL}');
    background-size: cover;
    background-position: center;
    padding: 2.4rem 2.4rem 2.2rem 2.4rem;
    color: white;
}}
.opal-hero-img h1 {{
    color: white;
    font-size: 2.1rem;
    margin: 0 0 0.45rem 0;
    font-weight: 700;
    letter-spacing: -0.01em;
    line-height: 1.15;
}}
.opal-hero-img p {{
    color: rgba(255,255,255,0.92);
    margin: 0;
    font-size: 1.05rem;
    font-weight: 400;
}}
.opal-hero-badge {{
    display: inline-block;
    background: rgba(255,255,255,0.18);
    color: white;
    padding: 0.22rem 0.7rem;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 600;
    margin-bottom: 0.9rem;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}}

/* Section cards */
.opal-section {{
    background: white;
    border-radius: 14px;
    padding: 1.4rem 1.5rem;
    box-shadow: 0 1px 3px rgba(14,26,43,0.06);
    border: 1px solid rgba(14,26,43,0.06);
    margin-bottom: 1.1rem;
}}
.opal-section h3 {{
    color: {BRAND_PRIMARY};
    margin: 0 0 0.85rem 0;
    font-size: 1.08rem;
    font-weight: 700;
    letter-spacing: -0.005em;
}}
.opal-step {{
    display: inline-block;
    background: {BRAND_PRIMARY};
    color: white;
    width: 1.5rem;
    height: 1.5rem;
    border-radius: 50%;
    text-align: center;
    line-height: 1.5rem;
    font-size: 0.82rem;
    font-weight: 700;
    margin-right: 0.55rem;
    vertical-align: middle;
}}

/* Buttons */
.stButton > button, .stDownloadButton > button {{
    background-color: {BRAND_PRIMARY};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.55rem 1.4rem;
    font-weight: 600;
    transition: background-color 0.15s ease;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{
    background-color: {BRAND_ACCENT};
    color: white;
}}
.stDownloadButton > button {{
    background-color: {BRAND_ACCENT};
    padding: 0.7rem 1.8rem;
    font-size: 1rem;
}}
.stDownloadButton > button:hover {{
    background-color: {BRAND_PRIMARY};
}}

/* File uploader */
[data-testid="stFileUploader"] section {{
    background: rgba(0, 102, 204, 0.04);
    border-radius: 10px;
    border: 1.5px dashed rgba(0, 102, 204, 0.35);
}}

/* Footer */
.opal-footer {{
    margin-top: 2.5rem;
    padding: 1.2rem 0 0.5rem 0;
    border-top: 1px solid rgba(14,26,43,0.08);
    text-align: center;
    color: {BRAND_MUTED};
    font-size: 0.88rem;
}}
.opal-footer a {{ color: {BRAND_ACCENT}; text-decoration: none; font-weight: 600; }}
.opal-footer a:hover {{ text-decoration: underline; }}

/* Required asterisk */
.req {{ color: #C0392B; font-weight: 700; }}

/* Mapping chip */
.map-row {{ padding: 0.35rem 0; border-bottom: 1px solid rgba(14,26,43,0.05); }}
.map-row:last-child {{ border-bottom: none; }}
.map-src {{ color: {BRAND_PRIMARY}; font-weight: 600; }}
.map-arrow {{ color: {BRAND_MUTED}; margin: 0 0.5rem; }}
.map-dst {{ color: {BRAND_INK}; }}
.map-missing {{ color: #B85C00; font-style: italic; }}
</style>
""", unsafe_allow_html=True)


# ---- Hero ----
st.markdown(f"""
<div class="opal-hero">
  <div class="opal-hero-img">
    <span class="opal-hero-badge">OPAL-RT • Internal Tool</span>
    <h1>{PAGE_TITLE}</h1>
    <p>{PAGE_SUBTITLE}</p>
  </div>
</div>
""", unsafe_allow_html=True)


# ---- Step 1: Upload ----
st.markdown('<div class="opal-section">'
            '<h3><span class="opal-step">1</span>Upload your source spreadsheet</h3>',
            unsafe_allow_html=True)
st.caption("Accepted formats: **.csv**, **.xlsx**. Messy column names, extra columns, "
           "and encoding glitches are fine — they\u2019ll be cleaned automatically.")
uploaded = st.file_uploader(
    "Upload CSV or Excel File",
    type=["csv", "xlsx"],
    label_visibility="collapsed",
)
st.markdown("</div>", unsafe_allow_html=True)


# ---- Step 2: Global settings ----
st.markdown('<div class="opal-section">'
            '<h3><span class="opal-step">2</span>Global import settings</h3>'
            '<p style="margin:-0.3rem 0 1rem 0; color:#5A6B82; font-size:0.92rem;">'
            'These values are applied to every row of the export. Fields marked '
            '<span class="req">*</span> are mandatory in Dynamics.</p>',
            unsafe_allow_html=True)

default_subject = f"{datetime.now().strftime('%Y%m')}Prospection"

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('**Subject** <span class="req">*</span>', unsafe_allow_html=True)
    subject = st.text_input("Subject", value=default_subject,
                            label_visibility="collapsed", max_chars=300,
                            help="Default format: YYYYMMProspection")
with c2:
    st.markdown("**Lead Source**")
    lead_source = st.selectbox("Lead Source", LEAD_SOURCES,
                               index=LEAD_SOURCES.index("Prospection"),
                               label_visibility="collapsed")
with c3:
    st.markdown("**Rating**")
    rating = st.selectbox("Rating", RATINGS, index=0, label_visibility="collapsed")

c4, c5, c6 = st.columns(3)
with c4:
    st.markdown("**Allow Marketing Communication**")
    allow_mm = st.selectbox("Allow MM", ALLOW_MM_OPTIONS, index=0,
                            label_visibility="collapsed")
with c5:
    st.markdown("**Source Campaign**")
    source_campaign = st.text_input("Source Campaign", value="",
                                    label_visibility="collapsed",
                                    placeholder="(optional)")
with c6:
    st.markdown("**Industry Sector**")
    industry_sector = st.selectbox("Industry Sector", INDUSTRY_SECTORS, index=0,
                                   label_visibility="collapsed")

c7, c8 = st.columns(2)
with c7:
    st.markdown("**Market Segment**")
    market_segment = st.selectbox("Market Segment", MARKET_SEGMENTS, index=0,
                                  label_visibility="collapsed",
                                  key="market_segment_select")
with c8:
    st.markdown("**Main Application**")
    main_app_opts = MAIN_APPS_BY_SEGMENT.get(market_segment, [""])
    main_application = st.selectbox(
        "Main Application", main_app_opts, index=0,
        label_visibility="collapsed",
        disabled=(market_segment == ""),
        help="Becomes available once a Market Segment is selected.",
    )

st.markdown("**Description**")
description = st.text_area("Description", value="", label_visibility="collapsed",
                          max_chars=2000, height=80,
                          placeholder="Optional note applied to every lead "
                                      "(e.g. campaign context, source list).")
st.markdown("</div>", unsafe_allow_html=True)


# ---- Step 3+: Process ----
if uploaded is not None:
    try:
        raw_df = read_uploaded_file(uploaded)
    except Exception as e:
        st.error(f"❌ Could not read the file: {e}")
        st.stop()

    if raw_df.empty:
        st.warning("The uploaded file has no rows.")
        st.stop()

    settings = {
        "Subject": clean_text(subject),
        "Lead Source": lead_source,
        "Rating": rating,
        "Allow Marketing Communication": allow_mm,
        "Source Campaign": clean_text(source_campaign),
        "Industry Sector": industry_sector,
        "Market Segment": market_segment,
        "Main Application": main_application if market_segment else "",
        "Description": clean_text(description),
    }

    final_df, mapping, errors, dropped_no_email = process(raw_df.copy(), settings)

    # ---- Column mapping preview ----
    st.markdown('<div class="opal-section">'
                '<h3><span class="opal-step">3</span>Detected column mapping</h3>',
                unsafe_allow_html=True)

    detected_for = ["First Name", "Last Name", "Company Name", "Job Title", "Email",
                    "Business Phone", "LinkedIn", "Country", "State or Province",
                    "Location", "Description"]
    map_lines = []
    for canonical in detected_for:
        src = mapping.get(canonical)
        if src:
            map_lines.append(
                f'<div class="map-row"><span class="map-src">{src}</span>'
                f'<span class="map-arrow">→</span><span class="map-dst">{canonical}</span></div>'
            )
        else:
            map_lines.append(
                f'<div class="map-row"><span class="map-missing">(no source column)</span>'
                f'<span class="map-arrow">→</span><span class="map-dst">{canonical}</span></div>'
            )
    st.markdown("".join(map_lines), unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ---- Validation ----
    st.markdown('<div class="opal-section">'
                '<h3><span class="opal-step">4</span>Validation</h3>',
                unsafe_allow_html=True)

    blocking_errors = [e for e in errors if e.get("type") != "info"]
    info_msgs = [e for e in errors if e.get("type") == "info"]

    summary_cols = st.columns(4)
    summary_cols[0].metric("Rows in export", len(final_df))
    summary_cols[1].metric("Skipped (no email)", dropped_no_email)
    summary_cols[2].metric("Validation issues", len(blocking_errors))
    summary_cols[3].metric("Source columns", len(raw_df.columns))

    for m in info_msgs:
        st.info(m["message"])

    if not blocking_errors:
        st.success("✅ File successfully normalized and ready for Dynamics import.")
    else:
        st.error(f"⚠️ {len(blocking_errors)} validation issue(s) found. "
                 "Review the details below before importing.")
        err_df = pd.DataFrame(blocking_errors)[["row", "field", "message"]]
        err_df.columns = ["Row", "Field", "Issue"]
        st.dataframe(err_df, use_container_width=True, hide_index=True,
                     height=min(420, 48 + 35 * len(err_df)))

    st.markdown("</div>", unsafe_allow_html=True)

    # ---- Preview & Export ----
    st.markdown('<div class="opal-section">'
                '<h3><span class="opal-step">5</span>Preview & export</h3>',
                unsafe_allow_html=True)

    with st.expander(f"Preview first 20 rows of the Dynamics-ready file "
                     f"({len(final_df)} total)", expanded=False):
        st.dataframe(final_df.head(20), use_container_width=True, hide_index=True)

    csv_bytes = final_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label=f"⬇  Download opalrt_dynamics_import.csv ({len(final_df)} rows)",
        data=csv_bytes,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv",
        use_container_width=False,
    )
    st.caption("UTF-8 encoded · matches the exact 21-column Dynamics template structure.")
    st.markdown("</div>", unsafe_allow_html=True)

else:
    st.info("⬆️ Upload a CSV or XLSX above to begin. Global settings will be applied "
            "to every row.")


# ---- Footer ----
st.markdown(
    '<div class="opal-footer">Built by <strong>Arnaud Joakim</strong> · '
    '<a href="mailto:arnaud.joakim@opal-rt.com">arnaud.joakim@opal-rt.com</a></div>',
    unsafe_allow_html=True,
)
