"""
Opal RT Spreadsheet Cleaner
---------------------------
Internal tool for converting messy lead spreadsheets into a Dynamics-compatible
import CSV.

Built by Arnaud Joakim <arnaud.joakim@opal-rt.com>
"""

from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# PAGE CONFIG - must be the first Streamlit call
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Opal RT Spreadsheet Cleaner",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ===========================================================================
# CONSTANTS
# ===========================================================================

# Exact Dynamics export column order (do not change)
DYNAMICS_COLUMNS: List[str] = [
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

MANDATORY_FIELDS: List[str] = [
    "Subject",
    "First Name",
    "Last Name",
    "Email",
    "Company Name",
    "Country",
]

FIELD_MAX_LENGTHS: Dict[str, int] = {
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

# Dropdown options (exact values from ImportLeadTemplate.xlsm)
LEAD_SOURCE_OPTIONS = [
    "",
    "Shows",
    "Web",
    "Prospection",
    "Webinar",
    "Referral",
    "Social Media",
    "Customer Portal",
    "SPS",
    "Others",
]

RATING_OPTIONS = ["", "Cold", "Warm", "Hot"]

ALLOW_MARKETING_OPTIONS = ["", "Yes", "No"]

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

MARKET_SEGMENT_OPTIONS = [
    "",
    "Aerospace",
    "Automotive",
    "Energy Conversion",
    "Marine, Railway, Off-Highway",
    "Power System",
]

MAIN_APPLICATION_BY_SEGMENT: Dict[str, List[str]] = {
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

# Canonical Dynamics countries (244 - exact list from template)
COUNTRIES: List[str] = [
    "Afghanistan", "African Country (non-maghrebian)", "Aland Island", "Albania",
    "Algeria", "American Samoa", "Andorra", "Angola", "Anguilla", "Antartica",
    "Antigua and Barbuda", "Argentina", "Armenia", "Aruba", "Australia", "Austria",
    "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus",
    "Belgium", "Belize", "Benin", "Bermuda", "Bhutan", "Bolivia",
    "Bosnia and Herzegovina", "Botswana", "Bouvet Island", "Brazil",
    "British Indian Ocean Territory", "Brunei Darussalam", "Bulgaria",
    "Burkina Faso", "Burundi", "Cambodia", "Cameroon", "Canada", "Cape Verde",
    "Cayman Islands", "Central African Republic", "Chad", "Chile", "China",
    "Chrismas Island", "Cocos (Keeling) Islands", "Colombia", "Comoros", "Congo",
    "Congo, The democatic Republic of the", "Cook Islands", "Costa Rica",
    "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark", "Djibouti",
    "Dominica", "Dominican Republic", "Egypt", "El Salvador", "Ecuador",
    "Equatorial Guinea", "Eritrea", "Estonia", "Ethiopia",
    "Falkland Islands (Malvinas)", "Faroe Island", "Fiji", "Finland", "France",
    "French Guiana", "French Polynesia", "French Southern Territories",
    "French-Guadeloupe", "French-Martinique", "Gabon", "Gambia", "Georgia",
    "Germany", "Ghana", "Gibraltar", "Greece", "Greenland", "Grenada", "Guam",
    "Guatemala", "Guernser", "Guinea", "Guinea-Bissau", "Guyana", "Haiti",
    "Heard Island and McDonals Islands", "Honduras", "Hong Kong", "Hungary",
    "Iceland", "India", "Indonesia", "Iran (Islamic Republic of)", "Iraq",
    "Ireland", "Isle of Man", "Israel", "Italy", "Ivory Coast", "Jamaica",
    "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan",
    "Lakshadweep", "Lao People's Democratic republic", "Latvia", "Lebanon",
    "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg",
    "Macao", "Macedonia", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali",
    "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mayotte", "Mexico",
    "Micronesia, Federated States of", "Moldova", "Monaco", "Mongolia",
    "Montenegro", "Montserrat", "Morocco", "Mozambique", "Myanmar", "N/A",
    "Namibia", "Nepal", "Netherlands", "Netherlands Antilles", "New Caledonia",
    "New Zealand", "Nicaragua", "Niger", "Nigeria", "Niue", "Norfolk Iskand",
    "Northern Mariana Islands", "Norway", "Oman", "Pakistan", "Palau",
    "Palestine", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines",
    "Pitcairn", "Poland", "Portugal", "Puerto Rico", "Qatar", "Reunion",
    "Romania", "Russia", "Rwanda", "Saint Barthelemy", "Saint Helena",
    "Saint Kitts and Nevis", "Saint Lucia", "Saint Pierre and Miquelon",
    "Saint Vincent and the Grenadines", "Samoa", "San Marino",
    "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles",
    "Shanghai", "Sierra Leone", "Singapore", "Slovakia", "Slovenia",
    "Solomon Islands", "Somalia", "South Africa",
    "South Georgia and the South Sandwich Islands", "South Korea", "Spain",
    "Sri Lanka", "St Martin", "Sudan", "Suriname", "Svalbard and Jan Mayen",
    "Swaziland", "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan",
    "Tanzania", "Thailand", "Timor-Leste", "Togo", "Trinidad and Tobago",
    "Tunisia", "Turkey", "Turkmenistan", "Turks and Caicos Island", "Tuvalu",
    "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom",
    "United States", "Uruguay", "Uzbekistan", "Vanuatu", "Vatican City State",
    "Venezuela", "Vietnam", "Virgin Islands, British", "Virgin Islands, U.S.",
    "Wallis and Futuna", "Western Sahara", "Yemen", "Zambia", "Zimbabwe",
]

# Common country aliases → canonical name (lowercase keys)
COUNTRY_ALIASES: Dict[str, str] = {
    "usa": "United States",
    "u.s.a.": "United States",
    "u.s.a": "United States",
    "us": "United States",
    "u.s.": "United States",
    "u.s": "United States",
    "united states of america": "United States",
    "united states": "United States",
    "america": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "u.k": "United Kingdom",
    "great britain": "United Kingdom",
    "britain": "United Kingdom",
    "england": "United Kingdom",
    "scotland": "United Kingdom",
    "wales": "United Kingdom",
    "northern ireland": "United Kingdom",
    "uae": "United Arab Emirates",
    "u.a.e.": "United Arab Emirates",
    "u.a.e": "United Arab Emirates",
    "russian federation": "Russia",
    "russia": "Russia",
    "south korea": "South Korea",
    "korea, republic of": "South Korea",
    "republic of korea": "South Korea",
    "korea (south)": "South Korea",
    "rok": "South Korea",
    "viet nam": "Vietnam",
    "vietnam": "Vietnam",
    "ivory coast": "Ivory Coast",
    "côte d'ivoire": "Ivory Coast",
    "cote d'ivoire": "Ivory Coast",
    "czech republic": "Czech Republic",
    "czechia": "Czech Republic",
    "hong kong sar": "Hong Kong",
    "hk": "Hong Kong",
    "taiwan, province of china": "Taiwan",
    "republic of china": "Taiwan",
    "iran": "Iran (Islamic Republic of)",
    "islamic republic of iran": "Iran (Islamic Republic of)",
    "laos": "Lao People's Democratic republic",
    "macau": "Macao",
    "syrian arab republic": "Syria",
    "republic of moldova": "Moldova",
    "bolivia, plurinational state of": "Bolivia",
    "venezuela, bolivarian republic of": "Venezuela",
    "tanzania, united republic of": "Tanzania",
    "north macedonia": "Macedonia",
    "republic of north macedonia": "Macedonia",
    "swaziland": "Swaziland",
    "eswatini": "Swaziland",
    "myanmar (burma)": "Myanmar",
    "burma": "Myanmar",
    "the netherlands": "Netherlands",
    "holland": "Netherlands",
    "deutschland": "Germany",
    "espana": "Spain",
    "españa": "Spain",
    "italia": "Italy",
    "brasil": "Brazil",
    "republic of singapore": "Singapore",
    "kingdom of saudi arabia": "Saudi Arabia",
    "ksa": "Saudi Arabia",
    "people's republic of china": "China",
    "prc": "China",
    "mainland china": "China",
}

# US states (with abbreviations)
US_STATES_FULL: List[str] = [
    "Alabama", "Alaska", "American Samoa", "Arizona", "Arkansas", "California",
    "Colorado", "Connecticut", "Delaware", "District of Columbia", "Florida",
    "Georgia", "Guam", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa",
    "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts",
    "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska",
    "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York",
    "North Carolina", "North Dakota", "Northern Mariana Islands", "Ohio",
    "Oklahoma", "Oregon", "Pennsylvania", "Puerto Rico", "Rhode Island",
    "South Carolina", "South Dakota", "Tennessee", "Texas",
    "United States Minor Outlying Islands", "Utah", "Vermont",
    "Virgin Islands, U.S.", "Virginia", "Washington", "West Virginia",
    "Wisconsin", "Wyoming",
]

US_STATE_ABBR: Dict[str, str] = {
    "al": "Alabama", "ak": "Alaska", "as": "American Samoa", "az": "Arizona",
    "ar": "Arkansas", "ca": "California", "co": "Colorado", "ct": "Connecticut",
    "de": "Delaware", "dc": "District of Columbia", "fl": "Florida",
    "ga": "Georgia", "gu": "Guam", "hi": "Hawaii", "id": "Idaho",
    "il": "Illinois", "in": "Indiana", "ia": "Iowa", "ks": "Kansas",
    "ky": "Kentucky", "la": "Louisiana", "me": "Maine", "md": "Maryland",
    "ma": "Massachusetts", "mi": "Michigan", "mn": "Minnesota",
    "ms": "Mississippi", "mo": "Missouri", "mt": "Montana", "ne": "Nebraska",
    "nv": "Nevada", "nh": "New Hampshire", "nj": "New Jersey",
    "nm": "New Mexico", "ny": "New York", "nc": "North Carolina",
    "nd": "North Dakota", "mp": "Northern Mariana Islands", "oh": "Ohio",
    "ok": "Oklahoma", "or": "Oregon", "pa": "Pennsylvania", "pr": "Puerto Rico",
    "ri": "Rhode Island", "sc": "South Carolina", "sd": "South Dakota",
    "tn": "Tennessee", "tx": "Texas", "ut": "Utah", "vt": "Vermont",
    "vi": "Virgin Islands, U.S.", "va": "Virginia", "wa": "Washington",
    "wv": "West Virginia", "wi": "Wisconsin", "wy": "Wyoming",
}

# Canadian provinces (with abbreviations) - using template's spelling (Québec)
CA_PROVINCES_FULL: List[str] = [
    "Alberta", "British Columbia", "Manitoba", "New Brunswick",
    "Newfoundland and Labrador", "Northwest Territories", "Nova Scotia",
    "Nunavut", "Ontario", "Prince Edward Island", "Québec", "Saskatchewan",
    "Yukon Territory",
]

CA_PROVINCE_ABBR: Dict[str, str] = {
    "ab": "Alberta", "bc": "British Columbia", "mb": "Manitoba",
    "nb": "New Brunswick", "nl": "Newfoundland and Labrador",
    "nt": "Northwest Territories", "ns": "Nova Scotia", "nu": "Nunavut",
    "on": "Ontario", "pe": "Prince Edward Island", "pei": "Prince Edward Island",
    "qc": "Québec", "que": "Québec", "quebec": "Québec",
    "sk": "Saskatchewan", "yt": "Yukon Territory", "yk": "Yukon Territory",
}

# Source-column detection rules. Order matters for tie-breaking (specific first).
COLUMN_ALIASES: Dict[str, List[str]] = {
    "First Name": [
        "first name", "firstname", "first_name", "fname", "given name",
        "given_name", "givenname", "first", "prenom", "prénom", "nom prenom",
    ],
    "Last Name": [
        "last name", "lastname", "last_name", "lname", "surname",
        "family name", "family_name", "familyname", "last", "nom de famille",
    ],
    "Company Name": [
        "company name", "company", "companyname", "company_name", "organization",
        "organisation", "org", "employer", "business name", "account", "firm",
        "entreprise", "société", "societe",
    ],
    "Job Title": [
        "job title", "jobtitle", "job_title", "title", "position", "role",
        "job role", "designation", "fonction", "poste",
    ],
    "Email": [
        "email", "e-mail", "email address", "emailaddress", "email_address",
        "work email", "business email", "corporate email", "professional email",
        "primary email", "mail", "courriel", "adresse email",
    ],
    "Business Phone": [
        "business phone", "businessphone", "business_phone", "work phone",
        "workphone", "work_phone", "office phone", "company phone",
        "mobile phone", "mobilephone", "mobile_phone", "mobile", "cell",
        "cell phone", "cellphone", "phone", "telephone", "tel", "phone number",
        "contact number", "phone no", "primary phone", "téléphone",
    ],
    "LinkedIn": [
        "linkedin", "linkedin profile", "linkedin profile url",
        "linkedin url", "linkedin link", "linkedinprofile", "linkedin_profile",
        "linkedin profile name", "li profile", "li url",
    ],
    "Country": [
        "country", "country/region", "country region", "country_region",
        "nation", "pays",
    ],
    "State or Province": [
        "state or province", "state/province", "state province", "state",
        "province", "region", "state_province", "stateprovince",
    ],
    "Description": [
        "description", "notes", "note", "comments", "comment", "remarks",
        "details", "about",
    ],
}

# "Location" gets a separate slot because it's parsed differently
LOCATION_ALIASES: List[str] = [
    "location", "city", "addresse", "address", "city/state", "city, state",
    "geography", "geo", "city and country",
]

# ---------------------------------------------------------------------------
# City → (Country, State/Province) lookup
# ---------------------------------------------------------------------------
# Used as a final fallback in parse_location_string for LinkedIn-style strings
# like 'Greater Chicago Area', 'Greater Toulouse Metropolitan Area', or just
# 'Houston'. Keys are LOWERCASE and ASCII-folded (no accents).
#
# Rules of curation:
#   - State/Province only included for US and Canada (matches template).
#   - Ambiguous city names (Springfield, Cambridge, London-the-CA-city, etc.)
#     are intentionally OMITTED rather than guessed.
#   - For mildly ambiguous names (Birmingham, Manchester, Athens), the most
#     populous / globally-recognised version wins.
CITY_TO_GEO: Dict[str, Tuple[str, str]] = {
    # ---------- United States ----------
    "new york": ("United States", "New York"),
    "new york city": ("United States", "New York"),
    "nyc": ("United States", "New York"),
    "los angeles": ("United States", "California"),
    "chicago": ("United States", "Illinois"),
    "houston": ("United States", "Texas"),
    "phoenix": ("United States", "Arizona"),
    "philadelphia": ("United States", "Pennsylvania"),
    "san antonio": ("United States", "Texas"),
    "san diego": ("United States", "California"),
    "dallas": ("United States", "Texas"),
    "fort worth": ("United States", "Texas"),
    "san jose": ("United States", "California"),
    "austin": ("United States", "Texas"),
    "jacksonville": ("United States", "Florida"),
    "columbus": ("United States", "Ohio"),
    "charlotte": ("United States", "North Carolina"),
    "san francisco": ("United States", "California"),
    "silicon valley": ("United States", "California"),
    "indianapolis": ("United States", "Indiana"),
    "seattle": ("United States", "Washington"),
    "denver": ("United States", "Colorado"),
    "washington dc": ("United States", "District of Columbia"),
    "washington d.c.": ("United States", "District of Columbia"),
    "boston": ("United States", "Massachusetts"),
    "nashville": ("United States", "Tennessee"),
    "el paso": ("United States", "Texas"),
    "detroit": ("United States", "Michigan"),
    "memphis": ("United States", "Tennessee"),
    "oklahoma city": ("United States", "Oklahoma"),
    "las vegas": ("United States", "Nevada"),
    "louisville": ("United States", "Kentucky"),
    "baltimore": ("United States", "Maryland"),
    "milwaukee": ("United States", "Wisconsin"),
    "albuquerque": ("United States", "New Mexico"),
    "tucson": ("United States", "Arizona"),
    "fresno": ("United States", "California"),
    "sacramento": ("United States", "California"),
    "atlanta": ("United States", "Georgia"),
    "kansas city": ("United States", "Missouri"),
    "colorado springs": ("United States", "Colorado"),
    "miami": ("United States", "Florida"),
    "fort lauderdale": ("United States", "Florida"),
    "raleigh": ("United States", "North Carolina"),
    "omaha": ("United States", "Nebraska"),
    "long beach": ("United States", "California"),
    "virginia beach": ("United States", "Virginia"),
    "oakland": ("United States", "California"),
    "minneapolis": ("United States", "Minnesota"),
    "saint paul": ("United States", "Minnesota"),
    "st. paul": ("United States", "Minnesota"),
    "tulsa": ("United States", "Oklahoma"),
    "tampa": ("United States", "Florida"),
    "new orleans": ("United States", "Louisiana"),
    "cleveland": ("United States", "Ohio"),
    "pittsburgh": ("United States", "Pennsylvania"),
    "cincinnati": ("United States", "Ohio"),
    "saint louis": ("United States", "Missouri"),
    "st. louis": ("United States", "Missouri"),
    "orlando": ("United States", "Florida"),
    "salt lake city": ("United States", "Utah"),
    "buffalo": ("United States", "New York"),
    "anaheim": ("United States", "California"),
    "santa ana": ("United States", "California"),
    "irvine": ("United States", "California"),
    "berkeley": ("United States", "California"),
    "palo alto": ("United States", "California"),
    "mountain view": ("United States", "California"),
    "sunnyvale": ("United States", "California"),
    "santa clara": ("United States", "California"),
    "santa monica": ("United States", "California"),
    "cupertino": ("United States", "California"),
    "menlo park": ("United States", "California"),
    "redwood city": ("United States", "California"),
    "san mateo": ("United States", "California"),
    "santa barbara": ("United States", "California"),
    "burbank": ("United States", "California"),
    "pasadena": ("United States", "California"),
    "long island": ("United States", "New York"),
    "brooklyn": ("United States", "New York"),
    "queens": ("United States", "New York"),
    "manhattan": ("United States", "New York"),
    "bronx": ("United States", "New York"),
    "newark": ("United States", "New Jersey"),
    "jersey city": ("United States", "New Jersey"),
    "trenton": ("United States", "New Jersey"),
    "princeton": ("United States", "New Jersey"),
    "stamford": ("United States", "Connecticut"),
    "hartford": ("United States", "Connecticut"),
    "providence": ("United States", "Rhode Island"),
    "annandale": ("United States", "Virginia"),
    "arlington": ("United States", "Virginia"),
    "alexandria": ("United States", "Virginia"),
    "reston": ("United States", "Virginia"),
    "tysons": ("United States", "Virginia"),
    "fairfax": ("United States", "Virginia"),
    "bethesda": ("United States", "Maryland"),
    "rockville": ("United States", "Maryland"),
    "annapolis": ("United States", "Maryland"),
    "silver spring": ("United States", "Maryland"),
    "ann arbor": ("United States", "Michigan"),
    "grand rapids": ("United States", "Michigan"),
    "madison": ("United States", "Wisconsin"),
    "des moines": ("United States", "Iowa"),
    "boulder": ("United States", "Colorado"),
    "boise": ("United States", "Idaho"),
    "anchorage": ("United States", "Alaska"),
    "honolulu": ("United States", "Hawaii"),
    "research triangle": ("United States", "North Carolina"),
    "durham": ("United States", "North Carolina"),
    "chapel hill": ("United States", "North Carolina"),
    "winston-salem": ("United States", "North Carolina"),
    "asheville": ("United States", "North Carolina"),
    "savannah": ("United States", "Georgia"),
    "augusta": ("United States", "Georgia"),
    "tallahassee": ("United States", "Florida"),
    "gainesville": ("United States", "Florida"),
    "boca raton": ("United States", "Florida"),
    "naples fl": ("United States", "Florida"),
    "huntsville": ("United States", "Alabama"),
    "birmingham al": ("United States", "Alabama"),

    # ---------- Canada ----------
    "toronto": ("Canada", "Ontario"),
    "greater toronto": ("Canada", "Ontario"),
    "gta": ("Canada", "Ontario"),
    "ottawa": ("Canada", "Ontario"),
    "mississauga": ("Canada", "Ontario"),
    "brampton": ("Canada", "Ontario"),
    "markham": ("Canada", "Ontario"),
    "vaughan": ("Canada", "Ontario"),
    "kitchener": ("Canada", "Ontario"),
    "waterloo": ("Canada", "Ontario"),
    "windsor on": ("Canada", "Ontario"),
    "montreal": ("Canada", "Québec"),
    "greater montreal": ("Canada", "Québec"),
    "laval": ("Canada", "Québec"),
    "quebec city": ("Canada", "Québec"),
    "sherbrooke": ("Canada", "Québec"),
    "trois-rivieres": ("Canada", "Québec"),
    "gatineau": ("Canada", "Québec"),
    "vancouver": ("Canada", "British Columbia"),
    "burnaby": ("Canada", "British Columbia"),
    "richmond bc": ("Canada", "British Columbia"),
    "calgary": ("Canada", "Alberta"),
    "edmonton": ("Canada", "Alberta"),
    "winnipeg": ("Canada", "Manitoba"),
    "halifax": ("Canada", "Nova Scotia"),
    "saskatoon": ("Canada", "Saskatchewan"),
    "regina": ("Canada", "Saskatchewan"),
    "st. john's": ("Canada", "Newfoundland and Labrador"),
    "fredericton": ("Canada", "New Brunswick"),
    "moncton": ("Canada", "New Brunswick"),
    "charlottetown": ("Canada", "Prince Edward Island"),
    "whitehorse": ("Canada", "Yukon Territory"),

    # ---------- United Kingdom ----------
    "london": ("United Kingdom", ""),
    "greater london": ("United Kingdom", ""),
    "manchester": ("United Kingdom", ""),
    "birmingham": ("United Kingdom", ""),
    "glasgow": ("United Kingdom", ""),
    "liverpool": ("United Kingdom", ""),
    "edinburgh": ("United Kingdom", ""),
    "bristol": ("United Kingdom", ""),
    "leeds": ("United Kingdom", ""),
    "sheffield": ("United Kingdom", ""),
    "cardiff": ("United Kingdom", ""),
    "belfast": ("United Kingdom", ""),
    "newcastle": ("United Kingdom", ""),
    "nottingham": ("United Kingdom", ""),
    "southampton": ("United Kingdom", ""),
    "aberdeen": ("United Kingdom", ""),
    "oxford": ("United Kingdom", ""),
    "brighton": ("United Kingdom", ""),

    # ---------- France ----------
    "paris": ("France", ""),
    "ile-de-france": ("France", ""),
    "toulouse": ("France", ""),
    "lyon": ("France", ""),
    "marseille": ("France", ""),
    "bordeaux": ("France", ""),
    "lille": ("France", ""),
    "nantes": ("France", ""),
    "strasbourg": ("France", ""),
    "rennes": ("France", ""),
    "grenoble": ("France", ""),
    "montpellier": ("France", ""),
    "nice cote d'azur": ("France", ""),

    # ---------- Germany ----------
    "berlin": ("Germany", ""),
    "munich": ("Germany", ""),
    "munchen": ("Germany", ""),
    "hamburg": ("Germany", ""),
    "frankfurt": ("Germany", ""),
    "stuttgart": ("Germany", ""),
    "cologne": ("Germany", ""),
    "koln": ("Germany", ""),
    "dusseldorf": ("Germany", ""),
    "leipzig": ("Germany", ""),
    "dresden": ("Germany", ""),
    "bremen": ("Germany", ""),
    "hannover": ("Germany", ""),
    "nuremberg": ("Germany", ""),
    "nurnberg": ("Germany", ""),

    # ---------- Italy ----------
    "rome": ("Italy", ""),
    "roma": ("Italy", ""),
    "milan": ("Italy", ""),
    "milano": ("Italy", ""),
    "naples": ("Italy", ""),
    "napoli": ("Italy", ""),
    "turin": ("Italy", ""),
    "torino": ("Italy", ""),
    "florence": ("Italy", ""),
    "firenze": ("Italy", ""),
    "venice": ("Italy", ""),
    "venezia": ("Italy", ""),
    "bologna": ("Italy", ""),
    "genoa": ("Italy", ""),
    "genova": ("Italy", ""),

    # ---------- Spain ----------
    "madrid": ("Spain", ""),
    "barcelona": ("Spain", ""),
    "seville": ("Spain", ""),
    "sevilla": ("Spain", ""),
    "valencia": ("Spain", ""),
    "bilbao": ("Spain", ""),
    "zaragoza": ("Spain", ""),
    "malaga": ("Spain", ""),

    # ---------- Netherlands / Belgium / Switzerland / Austria ----------
    "amsterdam": ("Netherlands", ""),
    "rotterdam": ("Netherlands", ""),
    "the hague": ("Netherlands", ""),
    "utrecht": ("Netherlands", ""),
    "eindhoven": ("Netherlands", ""),
    "brussels": ("Belgium", ""),
    "antwerp": ("Belgium", ""),
    "ghent": ("Belgium", ""),
    "zurich": ("Switzerland", ""),
    "geneva": ("Switzerland", ""),
    "basel": ("Switzerland", ""),
    "bern": ("Switzerland", ""),
    "lausanne": ("Switzerland", ""),
    "vienna": ("Austria", ""),
    "graz": ("Austria", ""),
    "salzburg": ("Austria", ""),

    # ---------- Nordics ----------
    "stockholm": ("Sweden", ""),
    "gothenburg": ("Sweden", ""),
    "malmo": ("Sweden", ""),
    "oslo": ("Norway", ""),
    "bergen": ("Norway", ""),
    "copenhagen": ("Denmark", ""),
    "aarhus": ("Denmark", ""),
    "helsinki": ("Finland", ""),
    "tampere": ("Finland", ""),
    "reykjavik": ("Iceland", ""),

    # ---------- Ireland / Portugal / Greece ----------
    "dublin": ("Ireland", ""),
    "cork": ("Ireland", ""),
    "galway": ("Ireland", ""),
    "lisbon": ("Portugal", ""),
    "porto": ("Portugal", ""),
    "athens": ("Greece", ""),
    "thessaloniki": ("Greece", ""),

    # ---------- Eastern Europe ----------
    "warsaw": ("Poland", ""),
    "krakow": ("Poland", ""),
    "wroclaw": ("Poland", ""),
    "poznan": ("Poland", ""),
    "gdansk": ("Poland", ""),
    "prague": ("Czech Republic", ""),
    "praha": ("Czech Republic", ""),
    "brno": ("Czech Republic", ""),
    "budapest": ("Hungary", ""),
    "bucharest": ("Romania", ""),
    "cluj-napoca": ("Romania", ""),
    "sofia": ("Bulgaria", ""),
    "belgrade": ("Serbia", ""),
    "zagreb": ("Croatia", ""),
    "ljubljana": ("Slovenia", ""),
    "bratislava": ("Slovakia", ""),
    "tallinn": ("Estonia", ""),
    "riga": ("Latvia", ""),
    "vilnius": ("Lithuania", ""),
    "moscow": ("Russia", ""),
    "saint petersburg": ("Russia", ""),
    "novosibirsk": ("Russia", ""),
    "kyiv": ("Ukraine", ""),
    "kiev": ("Ukraine", ""),
    "lviv": ("Ukraine", ""),
    "minsk": ("Belarus", ""),

    # ---------- Middle East ----------
    "istanbul": ("Turkey", ""),
    "ankara": ("Turkey", ""),
    "izmir": ("Turkey", ""),
    "tel aviv": ("Israel", ""),
    "jerusalem": ("Israel", ""),
    "haifa": ("Israel", ""),
    "dubai": ("United Arab Emirates", ""),
    "abu dhabi": ("United Arab Emirates", ""),
    "sharjah": ("United Arab Emirates", ""),
    "doha": ("Qatar", ""),
    "riyadh": ("Saudi Arabia", ""),
    "jeddah": ("Saudi Arabia", ""),
    "mecca": ("Saudi Arabia", ""),
    "kuwait city": ("Kuwait", ""),
    "manama": ("Bahrain", ""),
    "muscat": ("Oman", ""),
    "amman": ("Jordan", ""),
    "beirut": ("Lebanon", ""),
    "tehran": ("Iran (Islamic Republic of)", ""),

    # ---------- Africa ----------
    "cairo": ("Egypt", ""),
    "alexandria eg": ("Egypt", ""),
    "casablanca": ("Morocco", ""),
    "rabat": ("Morocco", ""),
    "marrakech": ("Morocco", ""),
    "tunis": ("Tunisia", ""),
    "algiers": ("Algeria", ""),
    "lagos": ("Nigeria", ""),
    "abuja": ("Nigeria", ""),
    "nairobi": ("Kenya", ""),
    "johannesburg": ("South Africa", ""),
    "cape town": ("South Africa", ""),
    "pretoria": ("South Africa", ""),
    "durban": ("South Africa", ""),
    "addis ababa": ("Ethiopia", ""),
    "accra": ("Ghana", ""),
    "dakar": ("Senegal", ""),

    # ---------- Asia ----------
    "tokyo": ("Japan", ""),
    "osaka": ("Japan", ""),
    "kyoto": ("Japan", ""),
    "yokohama": ("Japan", ""),
    "nagoya": ("Japan", ""),
    "sapporo": ("Japan", ""),
    "fukuoka": ("Japan", ""),
    "seoul": ("South Korea", ""),
    "busan": ("South Korea", ""),
    "incheon": ("South Korea", ""),
    "daegu": ("South Korea", ""),
    "beijing": ("China", ""),
    "shenzhen": ("China", ""),
    "guangzhou": ("China", ""),
    "chengdu": ("China", ""),
    "hangzhou": ("China", ""),
    "nanjing": ("China", ""),
    "xian": ("China", ""),
    "wuhan": ("China", ""),
    "tianjin": ("China", ""),
    "taipei": ("Taiwan", ""),
    "kaohsiung": ("Taiwan", ""),
    "taichung": ("Taiwan", ""),
    "bangkok": ("Thailand", ""),
    "chiang mai": ("Thailand", ""),
    "kuala lumpur": ("Malaysia", ""),
    "penang": ("Malaysia", ""),
    "jakarta": ("Indonesia", ""),
    "surabaya": ("Indonesia", ""),
    "bali": ("Indonesia", ""),
    "manila": ("Philippines", ""),
    "cebu": ("Philippines", ""),
    "ho chi minh city": ("Vietnam", ""),
    "ho chi minh": ("Vietnam", ""),
    "hanoi": ("Vietnam", ""),
    "saigon": ("Vietnam", ""),
    "phnom penh": ("Cambodia", ""),
    "yangon": ("Myanmar", ""),
    "mumbai": ("India", ""),
    "bombay": ("India", ""),
    "delhi": ("India", ""),
    "new delhi": ("India", ""),
    "national capital region india": ("India", ""),
    "ncr india": ("India", ""),
    "bangalore": ("India", ""),
    "bengaluru": ("India", ""),
    "hyderabad": ("India", ""),
    "chennai": ("India", ""),
    "madras": ("India", ""),
    "kolkata": ("India", ""),
    "calcutta": ("India", ""),
    "pune": ("India", ""),
    "ahmedabad": ("India", ""),
    "gurgaon": ("India", ""),
    "gurugram": ("India", ""),
    "noida": ("India", ""),
    "karachi": ("Pakistan", ""),
    "lahore": ("Pakistan", ""),
    "islamabad": ("Pakistan", ""),
    "dhaka": ("Bangladesh", ""),
    "colombo": ("Sri Lanka", ""),
    "kathmandu": ("Nepal", ""),

    # ---------- Oceania ----------
    "sydney": ("Australia", ""),
    "melbourne": ("Australia", ""),
    "brisbane": ("Australia", ""),
    "perth": ("Australia", ""),
    "adelaide": ("Australia", ""),
    "canberra": ("Australia", ""),
    "gold coast": ("Australia", ""),
    "auckland": ("New Zealand", ""),
    "wellington": ("New Zealand", ""),
    "christchurch": ("New Zealand", ""),

    # ---------- Latin America ----------
    "mexico city": ("Mexico", ""),
    "guadalajara": ("Mexico", ""),
    "monterrey": ("Mexico", ""),
    "puebla": ("Mexico", ""),
    "tijuana": ("Mexico", ""),
    "queretaro": ("Mexico", ""),
    "sao paulo": ("Brazil", ""),
    "rio de janeiro": ("Brazil", ""),
    "brasilia": ("Brazil", ""),
    "belo horizonte": ("Brazil", ""),
    "porto alegre": ("Brazil", ""),
    "curitiba": ("Brazil", ""),
    "salvador": ("Brazil", ""),
    "recife": ("Brazil", ""),
    "fortaleza": ("Brazil", ""),
    "buenos aires": ("Argentina", ""),
    "cordoba": ("Argentina", ""),
    "santiago": ("Chile", ""),
    "valparaiso": ("Chile", ""),
    "lima": ("Peru", ""),
    "bogota": ("Colombia", ""),
    "medellin": ("Colombia", ""),
    "cali": ("Colombia", ""),
    "caracas": ("Venezuela", ""),
    "quito": ("Ecuador", ""),
    "guayaquil": ("Ecuador", ""),
    "la paz": ("Bolivia", ""),
    "asuncion": ("Paraguay", ""),
    "montevideo": ("Uruguay", ""),
    "panama city": ("Panama", ""),
    "san jose costa rica": ("Costa Rica", ""),
    "havana": ("Cuba", ""),
    "san salvador": ("El Salvador", ""),
    "tegucigalpa": ("Honduras", ""),
    "managua": ("Nicaragua", ""),
    "guatemala city": ("Guatemala", ""),
    "santo domingo": ("Dominican Republic", ""),
    "san juan pr": ("Puerto Rico", ""),
}

EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
)

# Brand palette
BRAND = {
    "navy": "#002F6C",
    "navy_dark": "#001A3F",
    "accent": "#0099D8",
    "accent_dark": "#007BB0",
    "bg": "#F4F6F9",
    "card": "#FFFFFF",
    "text": "#1A1A1A",
    "text_muted": "#5A6473",
    "border": "#E1E6EE",
    "success_bg": "#E8F6EC",
    "success_border": "#2E7D32",
    "error_bg": "#FDECEC",
    "error_border": "#C62828",
}

# ===========================================================================
# CUSTOM CSS / BRANDING
# ===========================================================================

CUSTOM_CSS = f"""
<style>
    /* ------- Base ------- */
    .stApp {{
        background: {BRAND['bg']};
    }}
    .block-container {{
        padding-top: 0rem !important;
        padding-bottom: 3rem !important;
        max-width: 1200px;
    }}
    /* Hide Streamlit chrome we don't need */
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header[data-testid="stHeader"] {{
        background: transparent;
    }}

    /* ------- Typography ------- */
    html, body, [class*="css"] {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, Roboto, Helvetica, Arial, sans-serif;
        color: {BRAND['text']};
    }}

    /* ------- Hero ------- */
    .hero {{
        position: relative;
        width: 100%;
        min-height: 220px;
        border-radius: 16px;
        overflow: hidden;
        margin-top: 1rem;
        margin-bottom: 1.5rem;
        background:
            linear-gradient(120deg, rgba(0,26,63,0.85) 0%, rgba(0,47,108,0.65) 50%, rgba(0,153,216,0.45) 100%),
            url("https://www.opal-rt.com/wp-content/uploads/2025/05/Hero-News-OPAL-RT.jpg") center/cover no-repeat,
            linear-gradient(120deg, {BRAND['navy_dark']} 0%, {BRAND['navy']} 50%, {BRAND['accent']} 100%);
        color: white;
        padding: 2.25rem 2rem;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-shadow: 0 10px 30px rgba(0, 47, 108, 0.18);
    }}
    .hero-eyebrow {{
        font-size: 0.75rem;
        letter-spacing: 0.18em;
        text-transform: uppercase;
        opacity: 0.85;
        margin-bottom: 0.5rem;
        font-weight: 600;
    }}
    .hero-title {{
        font-size: 2.2rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.15;
        letter-spacing: -0.01em;
    }}
    .hero-subtitle {{
        font-size: 1.05rem;
        margin-top: 0.6rem;
        opacity: 0.95;
        max-width: 720px;
        line-height: 1.5;
    }}
    @media (max-width: 640px) {{
        .hero {{ padding: 1.5rem 1.25rem; min-height: 180px; border-radius: 12px; }}
        .hero-title {{ font-size: 1.5rem; }}
        .hero-subtitle {{ font-size: 0.95rem; }}
    }}

    /* ------- Section cards ------- */
    .section-card {{
        background: {BRAND['card']};
        border: 1px solid {BRAND['border']};
        border-radius: 14px;
        padding: 1.5rem 1.5rem 1.25rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
    }}
    .section-card h3 {{
        margin: 0 0 0.25rem 0;
        font-size: 1.15rem;
        color: {BRAND['navy']};
        font-weight: 700;
    }}
    .section-card .section-hint {{
        font-size: 0.85rem;
        color: {BRAND['text_muted']};
        margin-bottom: 1rem;
    }}
    .req-asterisk {{
        color: {BRAND['accent']};
        font-weight: 700;
        margin-left: 2px;
    }}

    /* ------- Streamlit widgets ------- */
    /* Buttons */
    .stButton > button {{
        background: linear-gradient(180deg, {BRAND['accent']} 0%, {BRAND['accent_dark']} 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 1.25rem;
        font-weight: 600;
        font-size: 0.95rem;
        box-shadow: 0 2px 6px rgba(0, 153, 216, 0.25);
        transition: transform 0.04s ease, box-shadow 0.15s ease, filter 0.15s ease;
        width: 100%;
    }}
    .stButton > button:hover {{
        filter: brightness(1.05);
        box-shadow: 0 4px 14px rgba(0, 153, 216, 0.35);
    }}
    .stButton > button:active {{
        transform: translateY(1px);
    }}
    /* Download button (secondary look – navy) */
    .stDownloadButton > button {{
        background: linear-gradient(180deg, {BRAND['navy']} 0%, {BRAND['navy_dark']} 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.7rem 1.25rem;
        font-weight: 600;
        font-size: 1rem;
        width: 100%;
        box-shadow: 0 2px 6px rgba(0, 47, 108, 0.25);
    }}
    .stDownloadButton > button:hover {{
        filter: brightness(1.1);
    }}

    /* File uploader */
    [data-testid="stFileUploader"] section {{
        border: 2px dashed {BRAND['accent']};
        background: rgba(0, 153, 216, 0.04);
        border-radius: 12px;
        padding: 1rem;
    }}
    [data-testid="stFileUploader"] section:hover {{
        background: rgba(0, 153, 216, 0.08);
    }}

    /* Inputs / selects */
    .stTextInput input, .stTextArea textarea {{
        border-radius: 8px !important;
        border-color: {BRAND['border']} !important;
    }}
    .stTextInput input:focus, .stTextArea textarea:focus {{
        border-color: {BRAND['accent']} !important;
        box-shadow: 0 0 0 1px {BRAND['accent']} !important;
    }}
    /* Selectbox styling – keep the native dropdown chrome but tone the border */
    div[data-baseweb="select"] > div {{
        border-radius: 8px !important;
        border-color: {BRAND['border']} !important;
    }}

    /* Tables */
    .stDataFrame {{
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid {BRAND['border']};
    }}

    /* Custom status banners */
    .success-banner {{
        background: {BRAND['success_bg']};
        border-left: 5px solid {BRAND['success_border']};
        color: #1B5E20;
        padding: 1rem 1.25rem;
        border-radius: 10px;
        font-weight: 600;
        margin: 0.5rem 0 1rem 0;
    }}
    .error-banner {{
        background: {BRAND['error_bg']};
        border-left: 5px solid {BRAND['error_border']};
        color: #8B0000;
        padding: 1rem 1.25rem;
        border-radius: 10px;
        font-weight: 600;
        margin: 0.5rem 0 1rem 0;
    }}
    .stats-row {{
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
        margin: 0.5rem 0 1.25rem 0;
    }}
    .stat-pill {{
        background: white;
        border: 1px solid {BRAND['border']};
        border-radius: 10px;
        padding: 0.65rem 1rem;
        min-width: 120px;
        flex: 1 1 120px;
    }}
    .stat-pill .stat-label {{
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: {BRAND['text_muted']};
        font-weight: 600;
    }}
    .stat-pill .stat-value {{
        font-size: 1.4rem;
        font-weight: 700;
        color: {BRAND['navy']};
        margin-top: 0.15rem;
    }}

    /* Mapping list */
    .mapping-grid {{
        display: grid;
        grid-template-columns: 1fr;
        gap: 0.4rem;
        font-size: 0.9rem;
    }}
    @media (min-width: 720px) {{
        .mapping-grid {{ grid-template-columns: 1fr 1fr; }}
    }}
    .mapping-row {{
        background: #F8FAFC;
        border: 1px solid {BRAND['border']};
        border-radius: 8px;
        padding: 0.5rem 0.75rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }}
    .mapping-arrow {{ color: {BRAND['accent']}; font-weight: 700; }}
    .mapping-source {{ color: {BRAND['text_muted']}; }}
    .mapping-target {{ color: {BRAND['navy']}; font-weight: 600; }}
    .mapping-missing {{ color: #B26A00; font-style: italic; }}

    /* Footer */
    .footer {{
        text-align: center;
        margin-top: 2.5rem;
        padding: 1.25rem;
        color: {BRAND['text_muted']};
        font-size: 0.85rem;
        border-top: 1px solid {BRAND['border']};
    }}
    .footer a {{
        color: {BRAND['accent']};
        text-decoration: none;
        font-weight: 600;
    }}
    .footer a:hover {{ text-decoration: underline; }}

    /* Make columns stack on small screens */
    @media (max-width: 640px) {{
        [data-testid="column"] {{
            width: 100% !important;
            flex: 1 1 100% !important;
        }}
    }}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ===========================================================================
# HELPER FUNCTIONS
# ===========================================================================

def fix_encoding(value) -> str:
    """Repair common mojibake (UTF-8 decoded as Latin-1) and strip junk chars."""
    if value is None:
        return ""
    if not isinstance(value, str):
        value = str(value)

    s = value

    # Try the classic latin1→utf8 roundtrip if we see telltale "Ã" patterns
    if "Ã" in s or "Â" in s:
        try:
            fixed = s.encode("latin-1", errors="strict").decode("utf-8", errors="strict")
            # Only accept the fix if it removed the suspicious sequences
            if ("Ã" in s and "Ã" not in fixed) or ("Â" in s and "Â" not in fixed):
                s = fixed
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass

    # Hand-roll fallback for the most common Latin-letter mojibake patterns
    # that survive the roundtrip. Smart-quote / dash patterns are handled by
    # the encode/decode attempt above (we skip them here to keep the literals
    # plain ASCII-safe).
    mojibake_map = {
        "Ã©": "é", "Ã¨": "è", "Ãª": "ê", "Ã«": "ë",
        "Ã ": "à", "Ã¢": "â", "Ã¤": "ä", "Ã¡": "á", "Ã£": "ã",
        "Ã®": "î", "Ã¯": "ï", "Ã­": "í",
        "Ã´": "ô", "Ã¶": "ö", "Ã²": "ò", "Ã³": "ó",
        "Ã»": "û", "Ã¼": "ü", "Ã¹": "ù", "Ãº": "ú",
        "Ã±": "ñ", "Ã§": "ç", "Ã¿": "ÿ",
        "Ã‰": "É", "Ãˆ": "È", "ÃŠ": "Ê",
        "Ã€": "À", "Ã‚": "Â",
        "Ã\u201d": "Ô", "Ã™": "Ù", "Ã›": "Û",
        "Ã‡": "Ç", "Ã\u2018": "Ñ",
    }
    for bad, good in mojibake_map.items():
        if bad in s:
            s = s.replace(bad, good)

    # Strip replacement characters & zero-width junk
    s = s.replace("\ufffd", "")  # U+FFFD replacement
    s = s.replace("\u200b", "")  # zero-width space
    s = s.replace("\ufeff", "")  # BOM
    s = s.replace("\u00a0", " ")  # NBSP → space

    # Normalise unicode (combining accents → composed)
    s = unicodedata.normalize("NFC", s)
    return s


def _ascii_fold(s: str) -> str:
    """Strip accents/diacritics for substring matching.
    'São Paulo' → 'Sao Paulo', 'Montréal' → 'Montreal', 'Düsseldorf' → 'Dusseldorf'."""
    if not s:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )


def lookup_city_in_location(location: str) -> Tuple[str, str]:
    """Search a free-text location for any city in CITY_TO_GEO and return
    (country, state_or_province). Used as a final fallback for LinkedIn-style
    strings like 'Greater Chicago Area' or 'Greater Toulouse Metropolitan Area'
    where no country is explicitly named.

    Matches on word boundaries and prefers the longest city name (so 'New York'
    wins over 'York', 'Saint Petersburg' wins over 'Petersburg' if present)."""
    if not location:
        return "", ""
    haystack = _ascii_fold(location.lower())
    # Sort keys longest-first so multi-word cities win over single-word substrings
    for city in sorted(CITY_TO_GEO.keys(), key=len, reverse=True):
        if re.search(r"\b" + re.escape(city) + r"\b", haystack):
            return CITY_TO_GEO[city]
    return "", ""


def clean_text(value, lowercase: bool = False) -> str:
    """Trim, collapse whitespace, fix encoding."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = fix_encoding(str(value))
    s = re.sub(r"\s+", " ", s).strip()
    if lowercase:
        s = s.lower()
    return s


def normalize_header(name: str) -> str:
    """Normalize a column header for fuzzy matching."""
    if name is None:
        return ""
    s = fix_encoding(str(name)).lower().strip()
    # Drop non-alphanumerics for comparison
    return re.sub(r"[^a-z0-9]+", "", s)


def detect_source_column(
    source_columns: List[str], target_field: str
) -> Optional[str]:
    """Find the best matching source column for a Dynamics target field."""
    aliases = COLUMN_ALIASES.get(target_field, [])
    normalized_sources = {normalize_header(c): c for c in source_columns}

    # 1) exact normalized match
    for alias in aliases:
        key = normalize_header(alias)
        if key in normalized_sources:
            return normalized_sources[key]

    # 2) contains match (alias as substring of source column)
    for alias in aliases:
        key = normalize_header(alias)
        if not key:
            continue
        for norm_src, orig_src in normalized_sources.items():
            if key == norm_src or key in norm_src:
                return orig_src

    return None


def detect_location_column(source_columns: List[str]) -> Optional[str]:
    """Locate a 'Location'-style column for parsing."""
    normalized_sources = {normalize_header(c): c for c in source_columns}
    for alias in LOCATION_ALIASES:
        key = normalize_header(alias)
        if key in normalized_sources:
            return normalized_sources[key]
    # also accept things containing 'location'
    for norm_src, orig_src in normalized_sources.items():
        if "location" in norm_src:
            return orig_src
    return None


def normalize_country(raw: str) -> str:
    """Return canonical country name (from COUNTRIES list) or '' if unrecognised."""
    if not raw:
        return ""
    s = clean_text(raw)
    if not s:
        return ""
    low = s.lower().strip(" .,")
    # Direct alias hit
    if low in COUNTRY_ALIASES:
        return COUNTRY_ALIASES[low]
    # Case-insensitive direct match against canonical list
    for c in COUNTRIES:
        if c.lower() == low:
            return c
    # Try stripping common prefixes ("the ")
    if low.startswith("the "):
        return normalize_country(low[4:])
    return ""


def normalize_us_state(raw: str) -> str:
    """Return canonical US state name or '' if not a recognised US state."""
    if not raw:
        return ""
    s = clean_text(raw).strip(" .,")
    if not s:
        return ""
    low = s.lower()
    # Abbreviation (e.g. 'ca', 'n.y.')
    abbr_key = re.sub(r"[^a-z]", "", low)
    if abbr_key in US_STATE_ABBR:
        return US_STATE_ABBR[abbr_key]
    # Full name (case-insensitive)
    for st_name in US_STATES_FULL:
        if st_name.lower() == low:
            return st_name
    return ""


def normalize_ca_province(raw: str) -> str:
    """Return canonical Canadian province (with template spelling) or ''."""
    if not raw:
        return ""
    s = clean_text(raw).strip(" .,")
    if not s:
        return ""
    low = s.lower()
    abbr_key = re.sub(r"[^a-z]", "", low)
    if abbr_key in CA_PROVINCE_ABBR:
        return CA_PROVINCE_ABBR[abbr_key]
    for prov in CA_PROVINCES_FULL:
        if prov.lower() == low:
            return prov
    # Also accept 'Quebec' without accent
    if low == "quebec":
        return "Québec"
    return ""


def parse_location_string(loc: str) -> Tuple[str, str]:
    """Parse a free-text location into (country, state_or_province).
    Only US/Canada keep a state/province. Other countries return ('Country', '').
    Returns ('', '') if unable to determine confidently."""
    if not loc:
        return "", ""
    s = clean_text(loc)
    if not s:
        return "", ""

    # Normalise common alternate separators to commas so we can split uniformly.
    # Handles pipe, slash, semicolon, newline, and " - " separators.
    s_norm = re.sub(r"\s*[\|/;]\s*", ", ", s)
    s_norm = re.sub(r"\s*\n\s*", ", ", s_norm)
    s_norm = re.sub(r"\s+-\s+", ", ", s_norm)

    # Split on commas (and clean parts)
    parts = [p.strip() for p in s_norm.split(",") if p.strip()]

    if not parts:
        return "", ""

    # Try identifying a country anywhere in the parts (rightmost first)
    country = ""
    country_idx = -1
    for i in range(len(parts) - 1, -1, -1):
        c = normalize_country(parts[i])
        if c:
            country = c
            country_idx = i
            break

    state = ""

    if country == "United States":
        # Look for a US state in parts before the country position (or anywhere)
        search_range = parts[:country_idx] if country_idx >= 0 else parts
        for i in range(len(search_range) - 1, -1, -1):
            cand = normalize_us_state(search_range[i])
            if cand:
                state = cand
                break
    elif country == "Canada":
        search_range = parts[:country_idx] if country_idx >= 0 else parts
        for i in range(len(search_range) - 1, -1, -1):
            cand = normalize_ca_province(search_range[i])
            if cand:
                state = cand
                break

    # If no country found, try inferring from a US state or CA province
    if not country:
        for i in range(len(parts) - 1, -1, -1):
            cand = normalize_us_state(parts[i])
            if cand:
                country = "United States"
                state = cand
                break
            cand = normalize_ca_province(parts[i])
            if cand:
                country = "Canada"
                state = cand
                break

    # Last-resort substring scan for free-text locations like
    # "Greater New York City Area" or "San Francisco Bay Area, USA"
    if not country:
        low = s.lower()
        # Country aliases (longest first to avoid 'us' matching inside 'austin')
        for alias, canon in sorted(COUNTRY_ALIASES.items(), key=lambda kv: -len(kv[0])):
            if len(alias) < 4:
                continue  # avoid short tokens like 'us', 'uk' as substrings
            if re.search(r"\b" + re.escape(alias) + r"\b", low):
                country = canon
                break
        if not country:
            for canon in sorted(COUNTRIES, key=len, reverse=True):
                if len(canon) < 4:
                    continue
                if re.search(r"\b" + re.escape(canon.lower()) + r"\b", low):
                    country = canon
                    break

    # CITY lookup runs BEFORE the single-word state substring scan, because
    # multi-word city keys ('Washington DC', 'New York City', 'Tel Aviv') are
    # more specific than ambiguous single-word state names ('Washington').
    if not country:
        city_country, city_state = lookup_city_in_location(s)
        if city_country:
            country = city_country
            if city_state and city_country in ("United States", "Canada"):
                state = city_state

    # If still no country but the string mentions a known US state / CA province as a word
    if not country:
        for st_name in US_STATES_FULL:
            if re.search(r"\b" + re.escape(st_name.lower()) + r"\b", s.lower()):
                country = "United States"
                state = st_name
                break
    if not country:
        for prov in CA_PROVINCES_FULL:
            if re.search(r"\b" + re.escape(prov.lower()) + r"\b", s.lower()):
                country = "Canada"
                state = prov
                break

    return country, state


def validate_email(email: str) -> bool:
    if not email:
        return False
    return bool(EMAIL_REGEX.match(email))


def default_subject() -> str:
    """Default 'YYYYMMProspection' subject for the current month."""
    return datetime.now().strftime("%Y%m") + "Prospection"


def truncate_to(value: str, max_len: int) -> str:
    """Currently we DO NOT truncate – we surface a validation error instead.
    This helper exists in case future config wants to enable truncation."""
    return value if len(value) <= max_len else value[:max_len]


# ===========================================================================
# CORE PIPELINE
# ===========================================================================

def read_uploaded_file(uploaded) -> pd.DataFrame:
    """Read a CSV or Excel file uploaded via Streamlit into a pandas DataFrame."""
    name = uploaded.name.lower()
    raw = uploaded.read()
    bio = io.BytesIO(raw)

    if name.endswith((".xlsx", ".xlsm", ".xls")):
        # openpyxl is the supported engine
        engine = "openpyxl"
        df = pd.read_excel(bio, engine=engine, dtype=str)
    elif name.endswith(".csv"):
        # Try a couple of encodings (UTF-8 first, then cp1252 for legacy exports)
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                bio.seek(0)
                df = pd.read_csv(bio, dtype=str, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("Could not decode the CSV file with common encodings.")
    else:
        raise ValueError("Unsupported file format. Please upload .csv or .xlsx.")

    return df


def strip_ghost_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove ghost columns: anything fully empty. Unnamed columns are KEPT
    if they hold data (some exports leave headers blank or write placeholders
    like 'Column9' for valid country/email cells)."""
    drop_cols = []
    for col in df.columns:
        is_empty = df[col].isna().all() or (df[col].astype(str).str.strip() == "").all()
        if is_empty:
            drop_cols.append(col)
            continue
        # If it's an utterly nameless column AND essentially empty header AND looks
        # like a pandas-generated placeholder, we already kept it above because it
        # has data. We deliberately do NOT drop populated 'Unnamed: N' / 'ColumnN'
        # columns — the row-scan fallback in process_dataframe will mine them for
        # country / state values.
    return df.drop(columns=drop_cols)


def build_column_mapping(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Map Dynamics target fields → source column names found in df."""
    cols = list(df.columns)
    mapping: Dict[str, Optional[str]] = {}
    for target in COLUMN_ALIASES.keys():
        mapping[target] = detect_source_column(cols, target)
    mapping["__location__"] = detect_location_column(cols)
    return mapping


def process_dataframe(
    raw_df: pd.DataFrame,
    settings: Dict[str, str],
) -> Tuple[pd.DataFrame, List[Dict], Dict[str, Optional[str]], Dict[str, int]]:
    """Main pipeline. Returns:
        - output_df: ready-to-export DataFrame in DYNAMICS_COLUMNS order
        - errors: list of {row, field, issue} dicts
        - mapping: detected column mapping for transparency
        - stats: counts of rows in / out / dropped"""

    df = strip_ghost_columns(raw_df.copy())

    # Normalize header strings (preserve original for display but trim)
    df.columns = [fix_encoding(str(c)).strip() for c in df.columns]

    mapping = build_column_mapping(df)

    # Build the set of source columns we've already mapped to a target field.
    # The row-scan fallback for country/state uses this to avoid scanning the
    # email / phone / name / etc. columns for country values (which would risk
    # false positives), and instead only mines truly unmapped columns
    # (e.g. 'Column9', 'Column11' from sloppy exports).
    mapped_source_columns = set()
    for _tgt, _src in mapping.items():
        if _src:
            mapped_source_columns.add(_src)

    output_rows: List[Dict] = []
    errors: List[Dict] = []
    seen_emails: Dict[str, int] = {}

    stats = {
        "rows_in": len(df),
        "rows_out": 0,
        "rows_skipped_no_email": 0,
        "rows_duplicate_email": 0,
    }

    for idx, source_row in df.iterrows():
        # Source-spreadsheet row number for user-facing messages (header = row 1)
        source_row_num = int(idx) + 2

        def pull(target: str) -> str:
            src_col = mapping.get(target)
            if src_col is None or src_col not in df.columns:
                return ""
            return clean_text(source_row[src_col])

        # --- pull each field ---
        first_name = pull("First Name")
        last_name = pull("Last Name")
        company = pull("Company Name")
        job_title = pull("Job Title")
        email = pull("Email").lower()
        phone = pull("Business Phone")
        linkedin = pull("LinkedIn")
        description = pull("Description")
        country_raw = pull("Country")
        state_raw = pull("State or Province")
        location_raw = ""
        loc_col = mapping.get("__location__")
        if loc_col and loc_col in df.columns:
            location_raw = clean_text(source_row[loc_col])

        # --- skip rows with no email entirely (per spec addendum) ---
        if not email:
            stats["rows_skipped_no_email"] += 1
            continue

        # --- de-duplicate by email (case-insensitive) ---
        if email in seen_emails:
            stats["rows_duplicate_email"] += 1
            continue
        seen_emails[email] = source_row_num

        # --- resolve Country & State/Province ---
        country = normalize_country(country_raw) if country_raw else ""
        state = ""
        # Only treat State/Province as valid for US / Canada
        if state_raw:
            us_candidate = normalize_us_state(state_raw)
            ca_candidate = normalize_ca_province(state_raw)
            if us_candidate:
                state = us_candidate
                if not country:
                    country = "United States"
            elif ca_candidate:
                state = ca_candidate
                if not country:
                    country = "Canada"

        # If still missing, try the Location column
        if (not country or not state) and location_raw:
            loc_country, loc_state = parse_location_string(location_raw)
            if not country and loc_country:
                country = loc_country
            if not state and loc_state and country in ("United States", "Canada"):
                state = loc_state

        # LAST-RESORT FALLBACK: scan every unmapped column in this row for a
        # value that *exactly* matches a country / US state / CA province.
        # This rescues sloppy exports where country lives in unnamed columns
        # like 'Column9' or 'Unnamed: 10' with no recognisable header.
        if not country or not state:
            for col in df.columns:
                if country and state:
                    break
                if col in mapped_source_columns:
                    continue
                if col not in source_row.index:
                    continue
                val = clean_text(source_row[col])
                if not val or len(val) > 60:
                    continue
                if not country:
                    c_match = normalize_country(val)
                    if c_match:
                        country = c_match
                        continue
                if not state:
                    us_match = normalize_us_state(val)
                    if us_match:
                        state = us_match
                        if not country:
                            country = "United States"
                        continue
                    ca_match = normalize_ca_province(val)
                    if ca_match:
                        state = ca_match
                        if not country:
                            country = "Canada"
                        continue

        # Country may have been pulled but isn't on the canonical list — final guard
        if country and country not in COUNTRIES:
            # If it's actually a US state or CA province name that landed in the
            # country slot, recover gracefully
            recovered_state = normalize_us_state(country) or normalize_ca_province(country)
            if normalize_us_state(country):
                country, state = "United States", normalize_us_state(country)
            elif normalize_ca_province(country):
                country, state = "Canada", normalize_ca_province(country)
            else:
                country = ""

        # State only applies to US/Canada
        if country not in ("United States", "Canada"):
            state = ""

        # --- Marketing/segment/sector: only fill from explicit user override or
        #     a direct source-file match (we have no source aliases for those,
        #     so they stay blank unless the user picked them in global settings)
        market_segment = settings.get("market_segment", "") or ""
        main_application = settings.get("main_application", "") or ""
        industry_sector = settings.get("industry_sector", "") or ""

        # --- Subject (mandatory, from global settings) ---
        subject = settings.get("subject", "").strip() or default_subject()

        out = {
            "(Do Not Modify) Lead": "",
            "(Do Not Modify) Row Checksum": "",
            "(Do Not Modify) Modified On": "",
            "Subject": subject,
            "First Name": first_name,
            "Last Name": last_name,
            "Job Title": job_title,
            "Company Name": company,
            "Email": email,
            "Business Phone": phone,
            "Country": country,
            "State or Province": state,
            "Description": description or settings.get("description", ""),
            "Lead Source": settings.get("lead_source", ""),
            "Rating": settings.get("rating", ""),
            "Source Campaign": settings.get("source_campaign", ""),
            "Market Segment": market_segment,
            "Main Application": main_application,
            "Industry Sector": industry_sector,
            "LinkedIn": linkedin,
            "Allow Marketing Communication": settings.get("allow_marketing", ""),
        }

        # ----- Row-level validation -----
        # Mandatory checks
        for f in MANDATORY_FIELDS:
            if not out.get(f):
                # For Country, include diagnostic context so the user can see
                # which source values we tried and why they didn't resolve.
                if f == "Country":
                    diag_bits = []
                    if country_raw:
                        diag_bits.append(f"raw country '{country_raw[:40]}' not recognised")
                    if location_raw:
                        diag_bits.append(f"location '{location_raw[:60]}' not parseable")
                    if not country_raw and not location_raw:
                        diag_bits.append("no Country or Location column detected for this row")
                    diag = f" — {'; '.join(diag_bits)}" if diag_bits else ""
                    errors.append({
                        "row": source_row_num,
                        "field": f,
                        "issue": f"Missing required field → Country{diag}",
                    })
                else:
                    errors.append({
                        "row": source_row_num,
                        "field": f,
                        "issue": f"Missing required field → {f}",
                    })

        # Email format
        if email and not validate_email(email):
            errors.append({
                "row": source_row_num,
                "field": "Email",
                "issue": f"Invalid email → {email}",
            })

        # Length validation (errors only — we do not silently truncate)
        for f, max_len in FIELD_MAX_LENGTHS.items():
            val = out.get(f, "")
            if val and len(val) > max_len:
                errors.append({
                    "row": source_row_num,
                    "field": f,
                    "issue": f"{f} exceeds {max_len} characters ({len(val)})",
                })

        output_rows.append(out)

    output_df = pd.DataFrame(output_rows, columns=DYNAMICS_COLUMNS)
    stats["rows_out"] = len(output_df)
    return output_df, errors, mapping, stats


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Serialize DataFrame to UTF-8 (BOM) CSV bytes matching Dynamics expectations."""
    buf = io.StringIO()
    df.to_csv(buf, index=False, encoding="utf-8")
    # utf-8-sig BOM helps Excel/Dynamics correctly detect encoding
    return ("\ufeff" + buf.getvalue()).encode("utf-8")


# ===========================================================================
# UI
# ===========================================================================

def render_hero() -> None:
    st.markdown(
        """
        <div class="hero">
            <div class="hero-eyebrow">OPAL-RT • Internal Tool</div>
            <h1 class="hero-title">Opal RT Spreadsheet Cleaner</h1>
            <p class="hero-subtitle">
                Prepare CRM-ready lead imports for Microsoft Dynamics.
                Auto-detect columns, normalize data, validate every row,
                and export a clean import-ready file in seconds.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_footer() -> None:
    st.markdown(
        """
        <div class="footer">
            Built by <strong>Arnaud Joakim</strong> ·
            <a href="mailto:arnaud.joakim@opal-rt.com">arnaud.joakim@opal-rt.com</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_global_settings() -> Dict[str, str]:
    """Render the 'Global Import Settings' section and return the chosen values."""
    st.markdown(
        '<div class="section-card">'
        '<h3>① Global Import Settings</h3>'
        '<div class="section-hint">These values are applied to <strong>every row</strong> of the export. '
        'Fields marked with <span class="req-asterisk">*</span> are required by Dynamics.</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        subject = st.text_input(
            "Subject *",
            value=default_subject(),
            max_chars=300,
            help="Default format: YYYYMMProspection (e.g. 202605Prospection)",
        )
        lead_source = st.selectbox("Lead Source", options=LEAD_SOURCE_OPTIONS, index=0)
        rating = st.selectbox("Rating", options=RATING_OPTIONS, index=0)
        allow_marketing = st.selectbox(
            "Allow Marketing Communication",
            options=ALLOW_MARKETING_OPTIONS,
            index=0,
        )
        source_campaign = st.text_input(
            "Source Campaign",
            value="",
            help="Free-text campaign identifier (optional)",
        )

    with col2:
        market_segment = st.selectbox(
            "Market Segment",
            options=MARKET_SEGMENT_OPTIONS,
            index=0,
            key="market_segment_select",
        )
        # Main Application options are dependent on Market Segment selection
        main_app_options = MAIN_APPLICATION_BY_SEGMENT.get(market_segment, [""])
        main_application = st.selectbox(
            "Main Application",
            options=main_app_options,
            index=0,
            key="main_application_select",
            help="Options change based on Market Segment selection",
        )
        industry_sector = st.selectbox(
            "Industry Sector",
            options=INDUSTRY_SECTOR_OPTIONS,
            index=0,
        )
        description = st.text_area(
            "Description",
            value="",
            height=98,
            help="Optional default description applied when source row has none",
        )

    st.markdown("</div>", unsafe_allow_html=True)

    return {
        "subject": subject.strip(),
        "lead_source": lead_source,
        "rating": rating,
        "allow_marketing": allow_marketing,
        "source_campaign": source_campaign.strip(),
        "market_segment": market_segment,
        "main_application": main_application,
        "industry_sector": industry_sector,
        "description": description.strip(),
    }


def render_mapping_panel(mapping: Dict[str, Optional[str]]) -> None:
    """Show which source columns were auto-detected for each Dynamics field."""
    with st.expander("🔎 View detected column mapping", expanded=False):
        rows_html = []
        # Build mapping in display order excluding the Location-internal key
        for target in COLUMN_ALIASES.keys():
            src = mapping.get(target)
            if src:
                rows_html.append(
                    f'<div class="mapping-row">'
                    f'<span class="mapping-source">{src}</span>'
                    f'<span class="mapping-arrow">→</span>'
                    f'<span class="mapping-target">{target}</span>'
                    f'</div>'
                )
            else:
                rows_html.append(
                    f'<div class="mapping-row">'
                    f'<span class="mapping-missing">— not found —</span>'
                    f'<span class="mapping-arrow">→</span>'
                    f'<span class="mapping-target">{target}</span>'
                    f'</div>'
                )
        loc = mapping.get("__location__")
        if loc:
            rows_html.append(
                f'<div class="mapping-row">'
                f'<span class="mapping-source">{loc}</span>'
                f'<span class="mapping-arrow">→</span>'
                f'<span class="mapping-target">Country + State/Province (parsed)</span>'
                f'</div>'
            )
        st.markdown(
            f'<div class="mapping-grid">{"".join(rows_html)}</div>',
            unsafe_allow_html=True,
        )


def render_stats(stats: Dict[str, int], n_errors: int) -> None:
    st.markdown(
        f"""
        <div class="stats-row">
            <div class="stat-pill">
                <div class="stat-label">Rows in source</div>
                <div class="stat-value">{stats['rows_in']}</div>
            </div>
            <div class="stat-pill">
                <div class="stat-label">Rows exported</div>
                <div class="stat-value">{stats['rows_out']}</div>
            </div>
            <div class="stat-pill">
                <div class="stat-label">Skipped (no email)</div>
                <div class="stat-value">{stats['rows_skipped_no_email']}</div>
            </div>
            <div class="stat-pill">
                <div class="stat-label">Duplicates removed</div>
                <div class="stat-value">{stats['rows_duplicate_email']}</div>
            </div>
            <div class="stat-pill">
                <div class="stat-label">Validation issues</div>
                <div class="stat-value">{n_errors}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_errors(errors: List[Dict]) -> None:
    """Show row-level errors in a tidy red panel."""
    if not errors:
        return
    st.markdown(
        f'<div class="error-banner">⚠️ {len(errors)} validation issue(s) detected. '
        f'Row numbers refer to the <strong>source file</strong> '
        f'(row 1 = headers, row 2 = first data row). '
        f'You can still export — Dynamics may reject affected rows.</div>',
        unsafe_allow_html=True,
    )
    err_df = pd.DataFrame(errors).sort_values(["row", "field"]).reset_index(drop=True)
    err_df.columns = ["Source row", "Field", "Issue"]
    st.dataframe(err_df, use_container_width=True, hide_index=True)


# ===========================================================================
# MAIN APP
# ===========================================================================

def main() -> None:
    render_hero()

    # -------- Global settings panel --------
    settings = render_global_settings()

    # -------- File upload panel --------
    st.markdown(
        '<div class="section-card">'
        '<h3>② Upload Source File</h3>'
        '<div class="section-hint">CSV or Excel (.xlsx). The app will auto-detect '
        'and map columns to the Dynamics template.</div>',
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader(
        "Upload CSV or Excel File",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
        label_visibility="collapsed",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    if not uploaded:
        st.info(
            "📥 Upload a lead spreadsheet above to begin. "
            "Supported formats: **.csv** and **.xlsx**.",
            icon="ℹ️",
        )
        render_footer()
        return

    # -------- Process --------
    st.markdown(
        '<div class="section-card">'
        '<h3>③ Normalize & Validate</h3>'
        '<div class="section-hint">Click below to process the uploaded file. '
        'The app will clean encoding, normalize formatting, parse locations, '
        'remove duplicate emails, and validate every row.</div>',
        unsafe_allow_html=True,
    )
    process_clicked = st.button("🚀 Process file", type="primary")
    st.markdown("</div>", unsafe_allow_html=True)

    if not process_clicked and "processed" not in st.session_state:
        render_footer()
        return

    if process_clicked:
        try:
            with st.spinner("Reading file…"):
                raw_df = read_uploaded_file(uploaded)
            with st.spinner("Normalizing and validating data…"):
                output_df, errors, mapping, stats = process_dataframe(raw_df, settings)
            st.session_state["processed"] = {
                "output_df": output_df,
                "errors": errors,
                "mapping": mapping,
                "stats": stats,
                "source_name": uploaded.name,
            }
        except Exception as e:  # noqa: BLE001
            st.markdown(
                f'<div class="error-banner">❌ Failed to read or process file: {e}</div>',
                unsafe_allow_html=True,
            )
            render_footer()
            return

    # -------- Results --------
    state = st.session_state.get("processed")
    if not state:
        render_footer()
        return

    output_df = state["output_df"]
    errors = state["errors"]
    mapping = state["mapping"]
    stats = state["stats"]

    st.markdown(
        '<div class="section-card">'
        '<h3>④ Results</h3>',
        unsafe_allow_html=True,
    )

    render_stats(stats, len(errors))
    render_mapping_panel(mapping)

    if not errors:
        st.markdown(
            '<div class="success-banner">✅ File successfully normalized and ready for Dynamics import.</div>',
            unsafe_allow_html=True,
        )
    else:
        render_errors(errors)

    # Preview
    st.markdown("**Preview (first 50 rows of the export):**")
    st.dataframe(output_df.head(50), use_container_width=True, hide_index=True)

    # Download
    csv_bytes = df_to_csv_bytes(output_df)
    st.download_button(
        label="⬇️ Download opalrt_dynamics_import.csv",
        data=csv_bytes,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv",
    )

    st.markdown("</div>", unsafe_allow_html=True)
    render_footer()


if __name__ == "__main__":
    main()
