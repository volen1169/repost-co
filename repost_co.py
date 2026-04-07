"""
Sales Territory Dashboard — Updated
Fixes:
  1. Scroll-to-top  → st.components.v1.html (ไม่โดน sanitize)
  2. Map Plus Code  → openlocationcode.js (OLC) decode แทน Nominatim
  3. Save จริง      → upload ขึ้น SharePoint Graph API
  4. SharePoint     → โหลด/บันทึกแยกแผนก (CA CO PH PL PO SF)

pip install streamlit[auth] pandas plotly openpyxl msal requests

MICROSOFT 365 / OIDC CONFIG (.streamlit/secrets.toml example)
[auth]
redirect_uri = "http://localhost:8501/oauth2callback"
cookie_secret = "CHANGE_ME_TO_A_LONG_RANDOM_SECRET"
client_id = "YOUR_ENTRA_APP_CLIENT_ID"
client_secret = "YOUR_ENTRA_APP_CLIENT_SECRET"
server_metadata_url = "https://login.microsoftonline.com/organizations/v2.0/.well-known/openid-configuration"

auth_allowed_email_domains = ["optimal.co.th"]
"""

# ── Imports ───────────────────────────────────────────────────────────────────
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io
import requests
import traceback
import json
import os
import textwrap
import msal
from datetime import datetime

APP_ENV = os.getenv("SALES_DASHBOARD_ENV", "development")
ADMIN_PASSWORD = os.getenv("SALES_DASHBOARD_ADMIN_PASSWORD", "1234")

# ── Page config (ต้องเป็น command แรก) ───────────────────────────────────────
st.set_page_config(page_title="Sales Territory Dashboard", page_icon="📊", layout="wide")

# ═══════════════════════════════════════════════════════════════════════════════
# SharePoint Config — อ่านจาก st.secrets (ไม่ hardcode ใน code)
# ═══════════════════════════════════════════════════════════════════════════════
def _get_secret(key: str, fallback: str = "") -> str:
    """อ่านค่าจาก st.secrets ก่อน แล้ว fallback ไป env var"""
    try:
        val = st.secrets.get(key)
        if val:
            return str(val).strip()
    except Exception:
        pass
    return os.getenv(key, fallback).strip()

SP_TENANT_ID     = _get_secret("SP_TENANT_ID")
SP_CLIENT_ID     = _get_secret("SP_CLIENT_ID")
SP_CLIENT_SECRET = _get_secret("SP_CLIENT_SECRET")
SP_HOST          = _get_secret("SP_HOST",      "optimalcoth.sharepoint.com")
SP_SITE_PATH     = _get_secret("SP_SITE_PATH", "/sites/SalesTerritory")
GRAPH_BASE       = "https://graph.microsoft.com/v1.0"

DEPARTMENTS = ["CA", "CO", "PH", "PL", "PO", "SF"]

DEPARTMENT_LABELS = {
    "CA": "Care Solutions",
    "CO": "Colourant Solutions",
    "PH": "Personalcare & Homecare",
    "PL": "Petroleum&Lubricant Solutions",
    "PO": "Polymer Solutions",
    "SF": "Surface Solutions",
}

DEPT_GROUPS = {
    "CA": "OPT Care Solutions",
    "CO": "OPT Colourant Solutions",
    "PH": "OPT Personalcare & Homecare",
    "PL": "OPT Petroleum&Lubricant Solutions",
    "PO": "OPT Polymer Solutions",
    "SF": "OPT Surface Solutions",
}

ADMIN_EMAILS = {
    "it@optimal.co.th",
    "itsuppor1t@poonyaruk.co.th",
    "ratchaphruek.r@optimal.co.th",
}

HEAD_EMAIL_TO_DEPT = {
    # ตัวอย่าง
    # "manager.ca@optimal.co.th": "CA",
    "itsupport@poonyaruk.co.th":"CO",
}

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════
REGION_COLORS = {
    "เหนือ": "#4C9BE8", "ตะวันออกเฉียงเหนือ": "#F4A261",
    "ออก": "#2A9D8F", "ตก": "#E76F51",
    "ใต้": "#8338EC", "กลาง": "#E63946", "ไม่ระบุ": "#ADB5BD",
}
REGION_EN_TO_TH = {
    "North": "เหนือ", "Northeast": "ตะวันออกเฉียงเหนือ",
    "East": "ออก", "West": "ตก",
    "South": "ใต้", "Central": "กลาง", "Unknown": "ไม่ระบุ",
}
ALL_PROVINCES = {
    "กรุงเทพมหานคร": "Central", "นนทบุรี": "Central", "ปทุมธานี": "Central",
    "พระนครศรีอยุธยา": "Central", "อ่างทอง": "Central", "ลพบุรี": "Central",
    "สิงห์บุรี": "Central", "ชัยนาท": "Central", "สระบุรี": "Central",
    "สมุทรปราการ": "Central", "สุพรรณบุรี": "Central", "นครปฐม": "Central",
    "สมุทรสาคร": "Central", "สมุทรสงคราม": "Central",
    "เชียงใหม่": "North", "เชียงราย": "North", "ลำปาง": "North",
    "ลำพูน": "North", "แม่ฮ่องสอน": "North", "น่าน": "North",
    "พะเยา": "North", "แพร่": "North", "อุตรดิตถ์": "North",
    "นครสวรรค์": "North", "อุทัยธานี": "North", "กำแพงเพชร": "North",
    "ตาก": "North", "สุโขทัย": "North", "พิษณุโลก": "North",
    "พิจิตร": "North", "เพชรบูรณ์": "North",
    "ขอนแก่น": "Northeast", "นครราชสีมา": "Northeast", "อุดรธานี": "Northeast",
    "อุบลราชธานี": "Northeast", "ร้อยเอ็ด": "Northeast", "ชัยภูมิ": "Northeast",
    "เลย": "Northeast", "สกลนคร": "Northeast", "กาฬสินธุ์": "Northeast",
    "มหาสารคาม": "Northeast", "มุกดาหาร": "Northeast", "หนองคาย": "Northeast",
    "หนองบัวลำภู": "Northeast", "บึงกาฬ": "Northeast", "นครพนม": "Northeast",
    "ยโสธร": "Northeast", "อำนาจเจริญ": "Northeast", "ศรีสะเกษ": "Northeast",
    "สุรินทร์": "Northeast", "บุรีรัมย์": "Northeast",
    "ชลบุรี": "East", "ระยอง": "East", "ฉะเชิงเทรา": "East",
    "จันทบุรี": "East", "ตราด": "East", "ปราจีนบุรี": "East",
    "นครนายก": "East", "สระแก้ว": "East",
    "ราชบุรี": "West", "กาญจนบุรี": "West", "เพชรบุรี": "West",
    "ประจวบคีรีขันธ์": "West",
    "สงขลา": "South", "สุราษฎร์ธานี": "South", "นครศรีธรรมราช": "South",
    "ภูเก็ต": "South", "กระบี่": "South", "ชุมพร": "South",
    "ตรัง": "South", "พังงา": "South", "ระนอง": "South",
    "สตูล": "South", "พัทลุง": "South", "ปัตตานี": "South",
    "ยะลา": "South", "นราธิวาส": "South",
}
PROVINCE_KEYWORDS_EN = {
    "Bangkok": "กรุงเทพมหานคร", "BKK": "กรุงเทพมหานคร",
    "Ladkrabang": "กรุงเทพมหานคร", "Suan Luang": "กรุงเทพมหานคร",
    "Suan-Luang": "กรุงเทพมหานคร", "Bangchak": "กรุงเทพมหานคร",
    "Bangkhen": "กรุงเทพมหานคร", "Bangkhuntien": "กรุงเทพมหานคร",
    "Bangbon": "กรุงเทพมหานคร", "Bangna": "กรุงเทพมหานคร",
    "Nongkhaem": "กรุงเทพมหานคร", "Prawet": "กรุงเทพมหานคร",
    "Khan Na Yao": "กรุงเทพมหานคร", "Jomthong": "กรุงเทพมหานคร",
    "Rama IX": "กรุงเทพมหานคร", "Rama lll": "กรุงเทพมหานคร", "Rama III": "กรุงเทพมหานคร",
    "Phra Nakhon Si Ayutthaya": "พระนครศรีอยุธยา", "Ayutthaya": "พระนครศรีอยุธยา",
    "Samut Prakan": "สมุทรปราการ", "Samutprakarn": "สมุทรปราการ",
    "Samuthprakarn": "สมุทรปราการ", "Sumutprakarn": "สมุทรปราการ",
    "Nonthaburi": "นนทบุรี",
    "Pathum Thani": "ปทุมธานี", "Pathumthani": "ปทุมธานี",
    "PathumThani": "ปทุมธานี", "Patumthani": "ปทุมธานี",
    "Ang Thong": "อ่างทอง", "Lop Buri": "ลพบุรี", "Lopburi": "ลพบุรี",
    "Sing Buri": "สิงห์บุรี", "Singburi": "สิงห์บุรี",
    "Chai Nat": "ชัยนาท", "Chainat": "ชัยนาท", "Saraburi": "สระบุรี",
    "Suphan Buri": "สุพรรณบุรี", "Suphanburi": "สุพรรณบุรี",
    "Nakhon Pathom": "นครปฐม", "Nakornpathom": "นครปฐม",
    "Nakornprathom": "นครปฐม", "Nakhonpathom": "นครปฐม",
    "Nakornchaisri": "นครปฐม", "Sampran": "นครปฐม", "Sam Phran": "นครปฐม",
    "Samut Sakhon": "สมุทรสาคร", "Samutsakhon": "สมุทรสาคร",
    "Samuthsakorn": "สมุทรสาคร", "Samutsakorn": "สมุทรสาคร",
    "Kratumbaen": "สมุทรสาคร", "Krathum Baen": "สมุทรสาคร", "Krathumbaen": "สมุทรสาคร",
    "Samut Songkhram": "สมุทรสงคราม",
    "Chiang Mai": "เชียงใหม่", "Chiangmai": "เชียงใหม่",
    "Chiang Rai": "เชียงราย", "Chiangrai": "เชียงราย",
    "Lampang": "ลำปาง", "Lamphun": "ลำพูน",
    "Mae Hong Son": "แม่ฮ่องสอน", "Maehongson": "แม่ฮ่องสอน",
    "Nan": "น่าน", "Phayao": "พะเยา", "Phrae": "แพร่", "Uttaradit": "อุตรดิตถ์",
    "Nakhon Sawan": "นครสวรรค์", "Nakhonsawan": "นครสวรรค์",
    "Uthai Thani": "อุทัยธานี", "Kamphaeng Phet": "กำแพงเพชร",
    "Tak": "ตาก", "Sukhothai": "สุโขทัย", "Phitsanulok": "พิษณุโลก",
    "Phichit": "พิจิตร", "Phetchabun": "เพชรบูรณ์",
    "Khon Kaen": "ขอนแก่น", "Khonkaen": "ขอนแก่น",
    "Nakhon Ratchasima": "นครราชสีมา", "Korat": "นครราชสีมา",
    "Udon Thani": "อุดรธานี", "Udonthani": "อุดรธานี",
    "Ubon Ratchathani": "อุบลราชธานี",
    "Roi Et": "ร้อยเอ็ด", "Chaiyaphum": "ชัยภูมิ", "Loei": "เลย",
    "Sakon Nakhon": "สกลนคร", "Kalasin": "กาฬสินธุ์",
    "Maha Sarakham": "มหาสารคาม", "Mukdahan": "มุกดาหาร",
    "Nong Khai": "หนองคาย", "Nongkhai": "หนองคาย",
    "Nong Bua Lam Phu": "หนองบัวลำภู", "Bueng Kan": "บึงกาฬ",
    "Nakhon Phanom": "นครพนม", "Yasothon": "ยโสธร",
    "Amnat Charoen": "อำนาจเจริญ", "Si Sa Ket": "ศรีสะเกษ", "Sisaket": "ศรีสะเกษ",
    "Surin": "สุรินทร์", "Buri Ram": "บุรีรัมย์", "Buriram": "บุรีรัมย์",
    "Chonburi": "ชลบุรี", "Chon Buri": "ชลบุรี", "Cholburi": "ชลบุรี", "Pattaya": "ชลบุรี",
    "Rayong": "ระยอง", "Chachoengsao": "ฉะเชิงเทรา", "Chanthaburi": "จันทบุรี", "Trat": "ตราด",
    "Prachin Buri": "ปราจีนบุรี", "Prachinburi": "ปราจีนบุรี",
    "Nakhon Nayok": "นครนายก", "Sa Kaeo": "สระแก้ว", "Sakaeo": "สระแก้ว",
    "Ratchaburi": "ราชบุรี", "Banpong": "ราชบุรี",
    "Kanchanaburi": "กาญจนบุรี",
    "Phetchaburi": "เพชรบุรี", "Petchaburi": "เพชรบุรี",
    "Prachuap Khiri Khan": "ประจวบคีรีขันธ์", "Prachuap": "ประจวบคีรีขันธ์",
    "Songkhla": "สงขลา", "Surat Thani": "สุราษฎร์ธานี", "Suratthani": "สุราษฎร์ธานี",
    "Nakhon Si Thammarat": "นครศรีธรรมราช",
    "Phuket": "ภูเก็ต", "Krabi": "กระบี่", "Chumphon": "ชุมพร", "Trang": "ตรัง",
    "Phang Nga": "พังงา", "Ranong": "ระนอง", "Satun": "สตูล",
    "Phatthalung": "พัทลุง", "Pattani": "ปัตตานี", "Yala": "ยะลา", "Narathiwat": "นราธิวาส",
}
POSTCODE_MAP = {
    "10": "กรุงเทพมหานคร", "11": "นนทบุรี", "12": "ปทุมธานี", "13": "พระนครศรีอยุธยา",
    "14": "อ่างทอง", "15": "ลพบุรี", "16": "สิงห์บุรี", "17": "ชัยนาท", "18": "สระบุรี",
    "20": "ชลบุรี", "21": "ระยอง", "22": "จันทบุรี", "23": "ตราด", "24": "ฉะเชิงเทรา",
    "25": "ปราจีนบุรี", "26": "นครนายก", "27": "สระแก้ว",
    "30": "นครราชสีมา", "31": "บุรีรัมย์", "32": "สุรินทร์", "33": "ศรีสะเกษ",
    "34": "อุบลราชธานี", "35": "ยโสธร", "36": "ชัยภูมิ", "37": "อำนาจเจริญ",
    "38": "มุกดาหาร", "39": "หนองบัวลำภู",
    "40": "ขอนแก่น", "41": "อุดรธานี", "42": "เลย", "43": "หนองคาย",
    "44": "มหาสารคาม", "45": "ร้อยเอ็ด", "46": "กาฬสินธุ์", "47": "สกลนคร",
    "48": "นครพนม", "49": "มุกดาหาร",
    "50": "เชียงใหม่", "51": "ลำพูน", "52": "ลำปาง", "53": "อุตรดิตถ์",
    "54": "แพร่", "55": "น่าน", "56": "พะเยา", "57": "เชียงราย", "58": "แม่ฮ่องสอน",
    "60": "นครสวรรค์", "61": "อุทัยธานี", "62": "กำแพงเพชร", "63": "ตาก",
    "64": "สุโขทัย", "65": "พิษณุโลก", "66": "พิจิตร", "67": "เพชรบูรณ์",
    "70": "ราชบุรี", "71": "กาญจนบุรี", "72": "สุพรรณบุรี", "73": "นครปฐม",
    "74": "สมุทรสาคร", "75": "สมุทรสงคราม", "76": "เพชรบุรี", "77": "ประจวบคีรีขันธ์",
    "80": "นครศรีธรรมราช", "81": "กระบี่", "82": "พังงา", "83": "ภูเก็ต",
    "84": "สุราษฎร์ธานี", "85": "ระนอง", "86": "ชุมพร",
    "90": "สงขลา", "91": "สตูล", "92": "ตรัง", "93": "พัทลุง",
    "94": "ปัตตานี", "95": "ยะลา", "96": "นราธิวาส",
}
GRADE_BASE = {"A": 5_000_000, "A-": 4_000_000, "B": 3_000_000, "B-": 2_500_000,
              "C": 2_000_000, "C-": 1_500_000, "F": 800_000}
TEMPLATE_COLS = ["Customer Name", "Salesperson", "Industry", "Grade",
                 "Sales/Year", "Budget_kg", "Actual_kg", "LastYear_kg", "Plus_Code", "Address"]

PROVINCE_CENTERS = {
    "กรุงเทพมหานคร": (13.7563, 100.5018), "นนทบุรี": (13.8621, 100.5144),
    "ปทุมธานี": (14.0208, 100.5250), "พระนครศรีอยุธยา": (14.3532, 100.5689),
    "สมุทรปราการ": (13.5991, 100.5998), "นครปฐม": (13.8199, 100.0622),
    "สมุทรสาคร": (13.5475, 100.2744), "ชลบุรี": (13.3611, 100.9847),
    "ระยอง": (12.6814, 101.2813), "ฉะเชิงเทรา": (13.6904, 101.0779),
    "ปราจีนบุรี": (14.0509, 101.3704), "นครนายก": (14.2069, 101.2131),
    "สระแก้ว": (13.8240, 102.0646), "ราชบุรี": (13.5367, 99.8171),
    "กาญจนบุรี": (14.0228, 99.5328), "เพชรบุรี": (13.1119, 99.9447),
    "ประจวบคีรีขันธ์": (11.8124, 99.7973), "เชียงใหม่": (18.7883, 98.9853),
    "เชียงราย": (19.9105, 99.8406), "ขอนแก่น": (16.4322, 102.8236),
    "นครราชสีมา": (14.9799, 102.0977), "อุดรธานี": (17.4138, 102.7872),
    "อุบลราชธานี": (15.2448, 104.8473), "สงขลา": (7.1898, 100.5951),
    "สุราษฎร์ธานี": (9.1382, 99.3215), "นครศรีธรรมราช": (8.4304, 99.9631),
    "ภูเก็ต": (7.8804, 98.3923), "กระบี่": (8.0863, 98.9063)
}

REGION_CENTERS = {
    "Central": (13.7367, 100.5231), "East": (13.1500, 101.1000),
    "North": (18.7900, 98.9800), "Northeast": (16.4322, 102.8236),
    "West": (13.7000, 99.5000), "South": (8.4300, 99.9600),
    "Unknown": (13.6776, 100.6262)
}

def resolve_reference_latlng(province: str = "", region: str = "", address: str = ""):
    prov = str(province or "").strip()
    if prov and prov in PROVINCE_CENTERS:
        return PROVINCE_CENTERS[prov]
    reg = str(region or "").strip()
    if reg and reg in REGION_CENTERS:
        return REGION_CENTERS[reg]
    addr = str(address or "").strip()
    if addr:
        _sub, _dis, _prov, _reg = parse_address(addr)
        if _prov and _prov in PROVINCE_CENTERS:
            return PROVINCE_CENTERS[_prov]
        if _reg and _reg in REGION_CENTERS:
            return REGION_CENTERS[_reg]
    return REGION_CENTERS["Unknown"]


# ═══════════════════════════════════════════════════════════════════════════════
# Microsoft 365 Custom Auth Helpers (NO secrets.toml required)
# ═══════════════════════════════════════════════════════════════════════════════
APP_BASE_URL   = os.getenv("APP_BASE_URL", "http://localhost:8501").rstrip("/")
REDIRECT_URI   = os.getenv("REDIRECT_URI", f"{APP_BASE_URL}/oauth2callback")
TENANT_ID      = os.getenv("TENANT_ID", "").strip()
CLIENT_ID      = os.getenv("CLIENT_ID", "").strip()
CLIENT_SECRET  = os.getenv("CLIENT_SECRET", "").strip()
AUTHORITY      = f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else ""
OIDC_SCOPES = ["User.Read", "GroupMember.Read.All"]
AUTH_READY     = bool(TENANT_ID and CLIENT_ID and CLIENT_SECRET and REDIRECT_URI)
AUTH_COOKIE_DAYS = 7
AUTH_COOKIE_PREFIX = "salesdash_"


def _js_escape(v: str) -> str:
    return str(v or "").replace("\\", "\\\\").replace("'", "\\'").replace("\n", " ")

def _set_auth_cookies(email: str = "", name: str = "", role: str = "", dept: str = "", is_admin: bool = False):
    email_js = _js_escape(email)
    name_js = _js_escape(name)
    role_js = _js_escape(role)
    dept_js = _js_escape(dept)
    is_admin_js = "1" if is_admin else "0"
    components.html(
        f"""
        <script>
        (function() {{
            var d = new Date();
            d.setTime(d.getTime() + ({AUTH_COOKIE_DAYS}*24*60*60*1000));
            var expires = "; expires=" + d.toUTCString() + "; path=/; SameSite=Lax";
            document.cookie = "{AUTH_COOKIE_PREFIX}email=" + encodeURIComponent('{email_js}') + expires;
            document.cookie = "{AUTH_COOKIE_PREFIX}name=" + encodeURIComponent('{name_js}') + expires;
            document.cookie = "{AUTH_COOKIE_PREFIX}role=" + encodeURIComponent('{role_js}') + expires;
            document.cookie = "{AUTH_COOKIE_PREFIX}dept=" + encodeURIComponent('{dept_js}') + expires;
            document.cookie = "{AUTH_COOKIE_PREFIX}is_admin=" + encodeURIComponent('{is_admin_js}') + expires;
        }})();
        </script>
        """,
        height=0,
    )

def _clear_auth_cookies():
    components.html(
        f"""
        <script>
        (function() {{
            var names = ["email","name","role","dept","is_admin"];
            names.forEach(function(n) {{
                document.cookie = "{AUTH_COOKIE_PREFIX}" + n + "=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/; SameSite=Lax";
            }});
        }})();
        </script>
        """,
        height=0,
    )

def _restore_session_from_cookies():
    try:
        cookies = st.context.cookies
        email = str(cookies.get(f"{AUTH_COOKIE_PREFIX}email", "") or "").strip().lower()
        if not email or st.session_state.get("auth_user"):
            return
        name = str(cookies.get(f"{AUTH_COOKIE_PREFIX}name", "") or "").strip() or (email.split("@")[0] if "@" in email else email)
        role = str(cookies.get(f"{AUTH_COOKIE_PREFIX}role", "") or "").strip()
        dept = str(cookies.get(f"{AUTH_COOKIE_PREFIX}dept", "") or "").strip()
        is_admin_raw = str(cookies.get(f"{AUTH_COOKIE_PREFIX}is_admin", "") or "").strip()
        st.session_state["auth_user"] = {"email": email, "name": name}
        st.session_state["user_email"] = email
        st.session_state["user_name"] = name
        if role:
            st.session_state["user_role"] = role
        if dept:
            st.session_state["dept"] = dept
        st.session_state["is_admin"] = is_admin_raw in ("1", "true", "True", "yes", "on")
    except Exception:
        pass

def _msal_app():
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
    )

def _auth_configured() -> bool:
    return AUTH_READY

def _session_logged_in() -> bool:
    return bool(st.session_state.get("auth_user"))

def _auth_logout():
    for k in ["auth_user", "auth_access_token", "auth_id_token_claims", "oauth_state"]:
        if k in st.session_state:
            del st.session_state[k]
    _clear_auth_cookies()
    try:
        st.query_params.clear()
    except Exception:
        pass

def _build_login_url():
    app = _msal_app()
    state = os.urandom(16).hex()
    st.session_state["oauth_state"] = state
    return app.get_authorization_request_url(
        scopes=OIDC_SCOPES,
        redirect_uri=REDIRECT_URI,
        state=state,
        prompt="select_account",
    )

def _complete_login_from_query():
    if not AUTH_READY:
        return
    qp = st.query_params
    code = qp.get("code")
    if not code:
        return
    state = qp.get("state")
    expected_state = st.session_state.get("oauth_state")
    if expected_state and state and state != expected_state:
        st.error("Microsoft 365 login state mismatch")
        st.stop()
    app = _msal_app()
    result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=OIDC_SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    if "access_token" not in result:
        st.error("Microsoft 365 login failed: " + str(result.get("error_description", result.get("error", "Unknown error"))))
        st.stop()
    claims = result.get("id_token_claims", {}) or {}
    email = (
        claims.get("preferred_username")
        or claims.get("email")
        or claims.get("upn")
        or claims.get("unique_name")
        or ""
    ).strip().lower()
    name = (
        claims.get("name")
        or claims.get("given_name")
        or email
        or "Microsoft 365 User"
    ).strip()
    st.session_state["auth_access_token"] = result["access_token"]
    st.session_state["auth_id_token_claims"] = claims
    st.session_state["auth_user"] = {
        "email": email,
        "name": name,
    }
    _set_auth_cookies(email=email, name=name)
    try:
        st.query_params.clear()
    except Exception:
        pass

def _get_allowed_email_domains() -> list:
    raw = _get_secret("AUTH_ALLOWED_EMAIL_DOMAINS", "optimal.co.th,poonyaruk.co.th")
    return [x.strip().lower() for x in str(raw).split(",") if x.strip()]

def _get_user_email() -> str:
    return str((st.session_state.get("auth_user") or {}).get("email", "")).strip().lower()

def _get_user_name() -> str:
    return str((st.session_state.get("auth_user") or {}).get("name", "")).strip() or "Microsoft 365 User"

def _user_email_allowed() -> bool:
    email = _get_user_email()
    if not email:
        return False

    admin_emails = {e.strip().lower() for e in ADMIN_EMAILS}
    if email in admin_emails:
        return True

    domains = _get_allowed_email_domains()
    if not domains:
        return True

    return any(email.endswith("@" + d) for d in domains)

def _get_user_groups() -> list[str]:
    token = st.session_state.get("auth_access_token", "")
    if not token:
        return []
    headers = {"Authorization": f"Bearer {token}"}
    groups = []
    url = GRAPH_BASE + "/me/memberOf?$select=displayName"
    try:
        while url:
            r = requests.get(url, headers=headers, timeout=20)
            if r.status_code == 403:
                st.warning("ไม่สามารถอ่าน Microsoft 365 group ได้ กรุณาเพิ่มสิทธิ์ GroupMember.Read.All และกด Grant admin consent")
                return []
            r.raise_for_status()
            data = r.json()
            for item in data.get("value", []):
                name = str(item.get("displayName", "")).strip()
                if name:
                    groups.append(name)
            url = data.get("@odata.nextLink")
    except Exception as exc:
        st.warning("อ่าน group จาก Microsoft 365 ไม่สำเร็จ: " + str(exc))
        return []
    return groups

def _resolve_role_and_dept(email: str | None = None, user_groups: list | None = None):
    email = str(email or _get_user_email() or "").strip().lower()
    groups = set(str(g).strip() for g in (user_groups or _get_user_groups()) if str(g).strip())

    if email in {e.lower() for e in ADMIN_EMAILS}:
        return "admin", None

    user_depts = [dept for dept, group_name in DEPT_GROUPS.items() if group_name in groups]
    if not user_depts:
        return None, None

    head_map = {str(k).strip().lower(): str(v).strip().upper() for k, v in HEAD_EMAIL_TO_DEPT.items()}
    if email in head_map:
        head_dept = head_map[email]
        if head_dept in user_depts:
            return "manager", head_dept
        return None, None

    return "staff", user_depts[0]

# ═══════════════════════════════════════════════════════════════════════════════
# SharePoint Auth & API
# ═══════════════════════════════════════════════════════════════════════════════
import time as _time

_SP_CACHE = {"tok": None, "tok_exp": 0.0, "site_id": None, "drive_id": None}


def _get_token() -> str:
    now = _time.time()
    if _SP_CACHE["tok"] and _SP_CACHE["tok_exp"] > now + 120:
        return _SP_CACHE["tok"]
    url = ("https://login.microsoftonline.com/"
           + SP_TENANT_ID + "/oauth2/v2.0/token")
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     SP_CLIENT_ID,
        "client_secret": SP_CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=15)
    resp.raise_for_status()
    d = resp.json()
    if "access_token" not in d:
        raise ConnectionError(d.get("error_description", str(d))[:300])
    _SP_CACHE["tok"]     = d["access_token"]
    _SP_CACHE["tok_exp"] = now + float(d.get("expires_in", 3600))
    return _SP_CACHE["tok"]


def _gh() -> dict:
    return {"Authorization": "Bearer " + _get_token()}


def _get_site_drive() -> tuple:
    if _SP_CACHE["site_id"] and _SP_CACHE["drive_id"]:
        return _SP_CACHE["site_id"], _SP_CACHE["drive_id"]
    h = _gh()
    r = requests.get(
        GRAPH_BASE + "/sites/" + SP_HOST + ":" + SP_SITE_PATH,
        headers=h, timeout=15)
    r.raise_for_status()
    sid = r.json()["id"]
    r2 = requests.get(GRAPH_BASE + "/sites/" + sid + "/drives",
                      headers=h, timeout=15)
    r2.raise_for_status()
    drives = r2.json().get("value", [])
    if not drives:
        raise ValueError("No drives found in SharePoint site")
    did = next((d["id"] for d in drives
                if "document" in d.get("name", "").lower()), drives[0]["id"])
    _SP_CACHE["site_id"]  = sid
    _SP_CACHE["drive_id"] = did
    return sid, did


def sp_list_files(dept: str) -> list:
    h = _gh()
    sid, did = _get_site_drive()
    url = (GRAPH_BASE + "/sites/" + sid + "/drives/" + did
           + "/root:/" + dept + ":/children")
    r = requests.get(url, headers=h, timeout=15)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    return [f for f in r.json().get("value", [])
            if f["name"].lower().endswith((".xlsx", ".csv"))]


def sp_load(dept: str, fname: str) -> pd.DataFrame:
    h = _gh()
    sid, did = _get_site_drive()
    url = (GRAPH_BASE + "/sites/" + sid + "/drives/" + did
           + "/root:/" + dept + "/" + fname + ":/content")
    r = requests.get(url, headers=h, timeout=30)
    r.raise_for_status()
    raw = io.BytesIO(r.content)
    if fname.lower().endswith(".csv"):
        df_out = pd.read_csv(raw)
        df_out["Region_TH"] = ""
        return df_out
    return build_df_from_original(pd.read_excel(raw, sheet_name=None))


def sp_save(df: pd.DataFrame, dept: str, fname: str) -> bool:
    try:
        h = _gh()
        h["Content-Type"] = "application/octet-stream"
        sid, did = _get_site_drive()
        url = (GRAPH_BASE + "/sites/" + sid + "/drives/" + did
               + "/root:/" + dept + "/" + fname + ":/content")
        r = requests.put(url, headers=h, data=to_excel_bytes(df), timeout=30)
        if r.status_code in (200, 201):
            return True
        st.error("Upload failed HTTP " + str(r.status_code) + ": " + r.text[:200])
        return False
    except Exception as exc:
        import traceback
        st.error("sp_save error: " + str(exc))
        st.code(traceback.format_exc())
        return False

# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _scroll_top():
    components.html(
        """<script>
        function doScroll() {
            var SELS = [
                '[data-testid="stAppViewContainer"]',
                '[data-testid="stMain"]',
                'section.main', '.main', '.appview-container',
                '[data-testid="stAppViewBlockContainer"]',
                '.block-container',
                '[data-testid="stVerticalBlock"]'
            ];
            var doc = window.parent.document;
            SELS.forEach(function(s){
                var el = doc.querySelector(s);
                if(el){ el.scrollTop=0; el.scrollTo(0,0); }
            });
            doc.documentElement.scrollTop = 0;
            doc.body.scrollTop = 0;
            window.parent.scrollTo(0, 0);
        }
        doScroll();
        setTimeout(doScroll,  80);
        setTimeout(doScroll, 300);
        </script>""",
        height=0,
    )
def parse_address(address: str):
    clean = re.sub(r"\s+", " ", str(address)).strip()
    province = ""
    for kw, prov in sorted(PROVINCE_KEYWORDS_EN.items(), key=lambda x: -len(x[0])):
        if re.search(r"\b" + re.escape(kw) + r"\b", clean, re.IGNORECASE):
            province = prov
            break
    if not province:
        for pc in re.findall(r"\b(\d{5})\b", clean):
            if pc[:2] in POSTCODE_MAP:
                province = POSTCODE_MAP[pc[:2]]
                break
    region_en = ALL_PROVINCES.get(province, "Unknown")
    sub_district, district = "", ""
    th_sub = re.search(r"ต\.([^\s,]+)", clean)
    th_dis = re.search(r"อ\.([^\s,]+)", clean)
    if th_sub: sub_district = th_sub.group(1).strip()
    if th_dis: district = th_dis.group(1).strip()
    if not sub_district:
        m = re.search(r"([A-Za-z][A-Za-z\s\-]+?)\s+[Ss]ub[-\s]?district", clean)
        if m: sub_district = m.group(1).strip().split(",")[-1].strip()
    if not district:
        m = re.search(r"([A-Za-z][A-Za-z\s\-]+?)\s+[Dd]istrict", clean)
        if m:
            candidate = m.group(1).strip().split(",")[-1].strip()
            if not sub_district or sub_district.lower() not in candidate.lower():
                district = candidate
    if not sub_district:
        m = re.search(r"\bT\.?\s*([A-Za-z][A-Za-z\s\-]+?)(?=\s+A\.|\s*,|\s+\d{5}|$)", clean)
        if m: sub_district = m.group(1).strip().rstrip(",")
    if not district:
        m = re.search(r"\bA\.?\s*(?:Mueang\s*)?([A-Za-z][A-Za-z\s\-]+?)(?=\s*,|\s+\d{5}|$)", clean)
        if m: district = m.group(1).strip().rstrip(",")
    return sub_district, district, province, region_en


def extract_plus_code_and_address(value: str):
    s = str(value or "").strip()
    if not s:
        return "", ""
    s = re.sub(r"\s+", " ", s).strip()
    m = re.search(r'([23456789CFGHJMPQRVWX]{4,8}\+[23456789CFGHJMPQRVWX]{2,3})', s, re.IGNORECASE)
    if not m:
        return "", s
    code = m.group(1).upper()
    remainder = (s[:m.start()] + " " + s[m.end():]).strip(" ,;-\t")
    remainder = re.sub(r"\s+", " ", remainder).strip()
    return code, remainder


def clean_plus_code(value: str) -> str:
    code, _ = extract_plus_code_and_address(value)
    return code


def merge_address_parts(address: str, plus_code_value: str) -> str:
    addr = str(address or "").strip()
    _, plus_tail = extract_plus_code_and_address(plus_code_value)
    if addr:
        return addr
    return plus_tail


def build_df_from_original(xl):
    import numpy as np
    if "All Customer" in xl:
        raw = xl["All Customer"].copy()
        raw = raw.rename(columns={
            "Customer name":       "Customer Name",
            "Salesperson (2026)":  "Salesperson",
            "Business type":       "Industry",
            "Budget (kg/year)":    "Budget_kg",
            "Plus Codes":          "Plus_Code",
        })
        raw = raw[raw["Customer Name"].notna()].copy()
        raw = raw[~raw["Customer Name"].astype(str).str.strip().isin(
            ["", "Customer name", "Customer Name"])].copy()
        raw["Grade"] = raw["Grade"].astype(str).str.strip().replace({"f": "F", "F ": "F", "nan": ""})
        raw["Colourant"] = ""
        raw["Address"] = raw["Address"].astype(str).str.replace(r"\n.*", "", regex=True).str.strip()

        raw["Plus_Code_Raw"] = raw["Plus_Code"].astype(str).fillna("")
        raw["Plus_Code"] = raw["Plus_Code_Raw"].apply(clean_plus_code)
        raw["Address"] = raw.apply(lambda r: merge_address_parts(r.get("Address", ""), r.get("Plus_Code_Raw", "")), axis=1)
        raw = raw.drop(columns=["Plus_Code_Raw"])
        raw["Budget_kg"] = pd.to_numeric(raw.get("Budget_kg", 0), errors="coerce").fillna(0).astype(int)
        for _col in ["Actual (kg/year)", "Actual_kg", "Actual kg", "Actual"]:
            if _col in raw.columns:
                raw = raw.rename(columns={_col: "Actual_kg"}); break
        if "Actual_kg" not in raw.columns: raw["Actual_kg"] = 0
        raw["Actual_kg"] = pd.to_numeric(raw["Actual_kg"], errors="coerce").fillna(0).astype(int)
        for _col in ["Last Year (kg)", "LastYear_kg", "Last_Year_kg", "Last Year kg"]:
            if _col in raw.columns:
                raw = raw.rename(columns={_col: "LastYear_kg"}); break
        if "LastYear_kg" not in raw.columns: raw["LastYear_kg"] = 0
        raw["LastYear_kg"] = pd.to_numeric(raw["LastYear_kg"], errors="coerce").fillna(0).astype(int)
        if "Sales/Year" in raw.columns:
            raw["Sales/Year"] = pd.to_numeric(raw["Sales/Year"], errors="coerce").fillna(0)
    elif "Original" in xl:
        raw = xl["Original"].copy()
        raw.columns = raw.iloc[1]
        raw = raw.iloc[2:].reset_index(drop=True)
        raw.columns = ["col0", "No", "Customer Name", "col3", "col4", "Industry", "Address"]
        raw = raw[raw["Customer Name"].notna()].copy()
        raw = raw[~raw["Customer Name"].astype(str).str.strip().isin(
            ["", "Customer name", "Customer Name"])].copy()
        raw["Address"] = raw["Address"].astype(str).str.strip()
        raw["Address"] = raw["Address"].str.replace(r"\n.*", "", regex=True).str.strip()
        GRADE_VALS = {"A", "A-", "B", "B-", "C", "C-", "F"}
        grades, salespersons, colourants = [], [], []
        for _, row in raw.iterrows():
            c3 = str(row["col3"]).strip() if pd.notna(row["col3"]) else ""
            c4 = str(row["col4"]).strip() if pd.notna(row["col4"]) else ""
            if c3 in GRADE_VALS:
                grades.append(c3); salespersons.append(c4); colourants.append("")
            else:
                grades.append(""); salespersons.append(c3); colourants.append(c4)
        raw["Grade"] = grades
        raw["Salesperson"] = [s.strip() for s in salespersons]
        raw["Colourant"] = [s.strip() for s in colourants]
        raw = raw.drop(columns=["col3", "col4"])
        if "Plus_Code" not in raw.columns:
            raw["Plus_Code"] = ""
        raw["Plus_Code_Raw"] = raw["Plus_Code"].astype(str).fillna("")
        raw["Plus_Code"] = raw["Plus_Code_Raw"].apply(clean_plus_code)
        raw["Address"] = raw.apply(lambda r: merge_address_parts(r.get("Address", ""), r.get("Plus_Code_Raw", "")), axis=1)
        raw = raw.drop(columns=["Plus_Code_Raw"])
    else:
        raise ValueError("ไม่พบ sheet 'All Customer' หรือ 'Original' ในไฟล์")

    loc = raw["Address"].apply(
        lambda a: pd.Series(parse_address(a), index=["Sub-district", "District", "Province", "Region"])
    )
    df = pd.concat([raw.reset_index(drop=True), loc.reset_index(drop=True)], axis=1)
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated(keep="last")].copy()
    if "Region" not in df.columns:
        df["Region"] = "Unknown"
    region_series = df["Region"]
    if isinstance(region_series, pd.DataFrame):
        region_series = region_series.iloc[:, -1]
    df["Region"] = region_series.fillna("Unknown").astype(str).replace({"": "Unknown", "nan": "Unknown"})
    df["Region_TH"] = df["Region"].map(REGION_EN_TO_TH).fillna("ไม่ระบุ")
    if "Sales/Year" not in df.columns:
        df["Sales/Year"] = 0.0
    else:
        df["Sales/Year"] = pd.to_numeric(df["Sales/Year"], errors="coerce").fillna(0)
    if "Plus_Code"   not in df.columns: df["Plus_Code"]   = ""
    df["Plus_Code"] = df["Plus_Code"].apply(clean_plus_code)
    if "Budget_kg"   not in df.columns: df["Budget_kg"]   = 0
    if "Actual_kg"   not in df.columns: df["Actual_kg"]   = 0
    if "LastYear_kg" not in df.columns: df["LastYear_kg"] = 0
    return df


def to_excel_bytes(df_out: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_out.to_excel(w, index=False, sheet_name="All Customer")
    return buf.getvalue()


def make_template() -> bytes:
    return to_excel_bytes(pd.DataFrame(columns=TEMPLATE_COLS))


# ═══════════════════════════════════════════════════════════════════════════════
# PLUS CODE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

_OLC_ALPHABET = "23456789CFGHJMPQRVWX"
_OLC_RES = [20.0, 1.0, 0.05, 0.0025, 0.000125]

def _olc_idx(ch: str) -> int:
    return _OLC_ALPHABET.find(ch.upper())

def _olc_decode_full(code: str):
    s = str(code or "").strip().upper()
    if "+" not in s:
        return None
    s = s.replace("+", "")
    s = re.sub(r"0+$", "", s)
    if len(s) < 4:
        return None
    lat = 0.0
    lng = 0.0
    upto = min(len(s), 10)
    for i in range(0, upto, 2):
        res = _OLC_RES[i // 2]
        li = _olc_idx(s[i])
        ni = _olc_idx(s[i + 1]) if i + 1 < len(s) else 0
        if li < 0 or ni < 0:
            return None
        lat += li * res
        lng += ni * res
    finest = _OLC_RES[min((len(s) - 1) // 2, 4)]
    return (lat - 90 + finest / 2, lng - 180 + finest / 2)

def _olc_prefix4(lat: float, lng: float) -> str:
    la = lat + 90
    lo = lng + 180
    p1l = int(la // 20)
    r1l = la - p1l * 20
    p1g = int(lo // 20)
    r1g = lo - p1g * 20
    return (
        _OLC_ALPHABET[int(la // 20)] +
        _OLC_ALPHABET[int(lo // 20)] +
        _OLC_ALPHABET[int(r1l)] +
        _OLC_ALPHABET[int(r1g)]
    )

def _olc_recover(short_code: str, ref_lat: float, ref_lng: float):
    s = str(short_code or "").strip().upper().split()[0]
    if "+" not in s:
        return None
    before, tail = s.split("+", 1)
    if len(before) >= 8:
        return _olc_decode_full(s)
    pf_len = 8 - len(before)
    best = _olc_decode_full(_olc_prefix4(ref_lat, ref_lng)[:pf_len] + before + "+" + tail)
    if best is None:
        return None
    gs = _OLC_RES[pf_len // 2 - 1]
    best_dist = 10**18
    for dl in (-1, 0, 1):
        for dg in (-1, 0, 1):
            cand_prefix = _olc_prefix4(ref_lat + dl * gs, ref_lng + dg * gs)[:pf_len]
            cand = _olc_decode_full(cand_prefix + before + "+" + tail)
            if cand is None:
                continue
            dist = (cand[0] - ref_lat) ** 2 + (cand[1] - ref_lng) ** 2
            if dist < best_dist:
                best_dist = dist
                best = cand
    return best

def plus_code_to_coords(code: str, ref_lat: float = 13.6776, ref_lng: float = 100.6262):
    s = clean_plus_code(code)
    if not s or "+" not in s:
        return None
    before = s.split("+", 1)[0]
    if len(before) >= 8:
        return _olc_decode_full(s)
    return _olc_recover(s, ref_lat, ref_lng)



def get_secret_or_default(key: str, default_value: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default_value)


def append_audit_log(action: str, detail: str = "", dept: str = ""):
    try:
        log_path = "sales_dashboard_audit_log.csv"
        row = pd.DataFrame([{
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "department": dept or st.session_state.get("dept") or "",
            "action": action,
            "detail": detail,
            "is_admin": bool(st.session_state.get("is_admin", False)),
            "user_email": st.session_state.get("user_email") or _get_user_email() or "",
            "user_name": st.session_state.get("user_name") or _get_user_name() or "",
            "role": st.session_state.get("user_role") or "",
        }])
        if os.path.exists(log_path):
            old = pd.read_csv(log_path)
            old = pd.concat([old, row], ignore_index=True)
            old.to_csv(log_path, index=False, encoding="utf-8-sig")
        else:
            row.to_csv(log_path, index=False, encoding="utf-8-sig")
    except Exception:
        pass


def sp_upload_bytes(content_bytes: bytes, remote_path: str, content_type: str = "application/octet-stream") -> bool:
    try:
        h = _gh()
        h["Content-Type"] = content_type
        sid, did = _get_site_drive()
        safe_path = remote_path.strip("/").replace(" ", "%20")
        url = GRAPH_BASE + "/sites/" + sid + "/drives/" + did + "/root:/" + safe_path + ":/content"
        r = requests.put(url, headers=h, data=content_bytes, timeout=45)
        if r.status_code in (200, 201):
            return True
        st.error("SharePoint upload failed HTTP " + str(r.status_code) + ": " + r.text[:250])
        return False
    except Exception as exc:
        st.error("sp_upload_bytes error: " + str(exc))
        return False

def push_audit_log_to_sharepoint() -> bool:
    log_path = "sales_dashboard_audit_log.csv"
    if not os.path.exists(log_path):
        st.warning("ยังไม่มี audit log ในเครื่องให้ส่งขึ้น SharePoint")
        return False
    try:
        with open(log_path, "rb") as f:
            data = f.read()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dept = (st.session_state.get("dept") or "ALL").strip()
        remote_path = f"Logs/{dept}/sales_dashboard_audit_log_{stamp}.csv"
        ok = sp_upload_bytes(data, remote_path, "text/csv")
        if ok:
            append_audit_log("push_audit_log_sharepoint", remote_path, dept)
            st.success("✅ ส่ง audit log ขึ้น SharePoint สำเร็จ")
        return ok
    except Exception as exc:
        st.error("push_audit_log_to_sharepoint error: " + str(exc))
        return False

def build_executive_report_df(df_in: pd.DataFrame) -> pd.DataFrame:
    rep = df_in.copy()
    rep["Sales/Year"] = pd.to_numeric(rep.get("Sales/Year", 0), errors="coerce").fillna(0)
    rep["Budget_kg"] = pd.to_numeric(rep.get("Budget_kg", 0), errors="coerce").fillna(0)
    rep["Actual_kg"] = pd.to_numeric(rep.get("Actual_kg", 0), errors="coerce").fillna(0)
    rep["LastYear_kg"] = pd.to_numeric(rep.get("LastYear_kg", 0), errors="coerce").fillna(0)
    rep["gap_kg"] = (rep["Budget_kg"] - rep["Actual_kg"]).clip(lower=0)
    rep["achievement_pct"] = rep.apply(lambda r: (r["Actual_kg"] / r["Budget_kg"] * 100) if r["Budget_kg"] > 0 else 0, axis=1)
    rep["yoy_pct"] = rep.apply(lambda r: ((r["Actual_kg"] - r["LastYear_kg"]) / r["LastYear_kg"] * 100) if r["LastYear_kg"] > 0 else 0, axis=1)
    rep["opportunity_score"] = (
        rep["gap_kg"].rank(pct=True).fillna(0) * 45
        + (100 - rep["achievement_pct"].clip(upper=100)).rank(pct=True).fillna(0) * 35
        + rep["Sales/Year"].rank(pct=True).fillna(0) * 20
    ).round(1)
    cols = [
        "Customer Name", "Salesperson", "Industry", "Province", "Region_TH", "Grade",
        "Sales/Year", "Budget_kg", "Actual_kg", "LastYear_kg", "gap_kg",
        "achievement_pct", "yoy_pct", "opportunity_score", "Plus_Code"
    ]
    for c in cols:
        if c not in rep.columns:
            rep[c] = ""
    return rep[cols].sort_values(["opportunity_score", "gap_kg", "Sales/Year"], ascending=[False, False, False])

def to_excel_bytes_multi(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df_sheet in sheets.items():
            clean_name = re.sub(r"[:\\/?*\[\]]", "_", str(sheet_name))[:31]
            df_sheet.to_excel(writer, index=False, sheet_name=clean_name)
    return buf.getvalue()

def build_map_points(df_in: pd.DataFrame, ref_lat: float = 13.6776, ref_lng: float = 100.6262):
    map_points = []
    map_points_no_coords = []
    if df_in is None or df_in.empty:
        return "[]", "[]"
    for _, row in df_in.iterrows():
        name = str(row.get("Customer Name", "") or "").strip()
        salesperson = str(row.get("Salesperson", "") or "").strip()
        province = str(row.get("Province", "") or "").strip()
        region = str(row.get("Region", "") or "").strip()
        address = str(row.get("Address", "") or "").strip()
        plus_code = str(row.get("Plus_Code", "") or "").strip()
        row_ref_lat, row_ref_lng = resolve_reference_latlng(province, region, address)
        if plus_code and "+" in plus_code:
            query = f"{plus_code} {address} {province} Thailand".strip()
        else:
            query = f"{name} {address} {province} Thailand".strip()
        coords = plus_code_to_coords(plus_code, ref_lat=row_ref_lat, ref_lng=row_ref_lng) if plus_code else None
        if coords:
            map_points.append({
                "name": name,
                "plus_code": plus_code,
                "lat": round(float(coords[0]), 7),
                "lng": round(float(coords[1]), 7),
                "salesperson": salesperson,
                "province": province,
                "query": query,
            })
        else:
            map_points_no_coords.append({
                "name": name,
                "plus_code": plus_code,
                "salesperson": salesperson,
                "province": province,
                "query": query,
            })
    return json.dumps(map_points, ensure_ascii=False), json.dumps(map_points_no_coords, ensure_ascii=False)


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INIT
# ═══════════════════════════════════════════════════════════════════════════════

EMPTY_DF = pd.DataFrame(columns=TEMPLATE_COLS + [
    "Region_TH", "Region", "Sub-district", "District", "Province"])

for _k, _v in [("dept", None), ("sp_file", None), ("df", EMPTY_DF),
               ("is_admin", False), ("user_role", "staff"), ("user_email", ""), ("user_name", ""),
               ("edit_mode", "edit"), ("editing_idx", None), ("confirm_delete", False),
               ("last_refresh", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
               ("last_menu_logged", ""),
               ("sp_file_last_modified", ""),
               ("sp_file_etag", ""),
               ("live_sync_enabled", True),
               ("live_sync_interval", 60),
               ("is_dirty", False)]:
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ═══════════════════════════════════════════════════════════════════════════════
# UI Helpers
# ═══════════════════════════════════════════════════════════════════════════════
def _dept_label(dept_code: str | None) -> str:
    code = str(dept_code or "").strip()
    return DEPARTMENT_LABELS.get(code, code)

def _role_label():
    role = st.session_state.get("user_role", "")
    dept = _dept_label(st.session_state.get("dept", ""))
    if role == "admin":
        return "👑 Admin (ทุกแผนก)"
    elif role == "manager":
        return f"🧑‍💼 หัวหน้าแผนก ({dept})"
    elif role == "staff":
        return f"👨‍💻 พนักงาน ({dept})"
    return "❓ ไม่ทราบสิทธิ์"


def _can_view_dashboard():
    return st.session_state.get("user_role", "") in ["admin", "manager"]

def _can_view_customer_data():
    return st.session_state.get("user_role", "") in ["admin", "manager", "staff"]



def render_live_sync_autorefresh(enabled: bool, interval_sec: int = 60):
    if not enabled:
        return
    interval_ms = max(15, int(interval_sec)) * 1000
    components.html(f"""
    <script>
    (function() {{
        if (window.__salesdash_live_sync_timer__) return;
        window.__salesdash_live_sync_timer__ = true;
        setTimeout(function() {{
            try {{ window.parent.location.reload(); }} catch (e) {{ window.location.reload(); }}
        }}, {interval_ms});
    }})();
    </script>
    """, height=0)


def grade_weight(g: str) -> int:
    return {"A": 40, "A-": 34, "B": 26, "B-": 20, "C": 12, "C-": 8, "F": 3}.get(str(g or "").strip(), 0)


def build_ai_recommendations(df_in: pd.DataFrame, origin_lat: float, origin_lng: float) -> pd.DataFrame:
    if df_in is None or df_in.empty:
        return pd.DataFrame()
    work = df_in.copy()
    for c in ["Sales/Year", "Budget_kg", "Actual_kg", "LastYear_kg"]:
        work[c] = pd.to_numeric(work.get(c, 0), errors="coerce").fillna(0)
    work["Gap_kg"] = (work["Budget_kg"] - work["Actual_kg"]).clip(lower=0)
    work["Achievement_pct"] = work.apply(lambda r: (r["Actual_kg"] / r["Budget_kg"] * 100) if r["Budget_kg"] > 0 else 0, axis=1)
    lats, lngs, dists, reasons = [], [], [], []
    for _, row in work.iterrows():
        province = str(row.get("Province", "") or "").strip()
        region = str(row.get("Region", "") or "").strip()
        address = str(row.get("Address", "") or "").strip()
        plus_code = str(row.get("Plus_Code", "") or "").strip()
        ref_lat, ref_lng = resolve_reference_latlng(province, region, address)
        coords = plus_code_to_coords(plus_code, ref_lat=ref_lat, ref_lng=ref_lng) if plus_code else None
        if coords:
            lat, lng = float(coords[0]), float(coords[1])
            dist = (((lat-origin_lat)**2 + (lng-origin_lng)**2) ** 0.5) * 111
        else:
            lat, lng, dist = None, None, 9999
        lats.append(lat); lngs.append(lng); dists.append(dist)
        why = []
        if grade_weight(row.get("Grade", "")) >= 34: why.append("เกรดสูง")
        if float(row.get("Gap_kg", 0) or 0) > 0: why.append("gap สูง")
        if float(row.get("Achievement_pct", 0) or 0) < 50: why.append("achievement ต่ำ")
        if dist < 80: why.append("ใกล้เส้นทาง")
        reasons.append(", ".join(why[:3]) if why else "โอกาสทั่วไป")
    work["Latitude"] = lats
    work["Longitude"] = lngs
    work["Distance_km"] = dists
    work["GradeScore"] = work["Grade"].apply(grade_weight)
    work["GapScore"] = work["Gap_kg"].rank(pct=True).fillna(0) * 30
    work["SalesScore"] = work["Sales/Year"].rank(pct=True).fillna(0) * 15
    work["AchPenalty"] = (100 - work["Achievement_pct"].clip(upper=100)).rank(pct=True).fillna(0) * 20
    work["DistanceScore"] = work["Distance_km"].apply(lambda d: 15 if d <= 40 else 10 if d <= 100 else 6 if d <= 180 else 2 if d < 9999 else 0)
    work["AI_Score"] = (work["GradeScore"] + work["GapScore"] + work["SalesScore"] + work["AchPenalty"] + work["DistanceScore"]).round(1)
    work["AI_Reason"] = reasons
    return work.sort_values(["AI_Score", "Gap_kg", "Sales/Year"], ascending=[False, False, False])


def build_google_maps_route_link(origin_label: str, stops: list[tuple[float, float]]) -> str:
    import urllib.parse
    origin = urllib.parse.quote(origin_label)
    if not stops:
        return f"https://www.google.com/maps/dir/?api=1&origin={origin}"
    if len(stops) == 1:
        lat, lng = stops[0]
        dest = urllib.parse.quote(f"{lat},{lng}")
        return f"https://www.google.com/maps/dir/?api=1&origin={origin}&destination={dest}&travelmode=driving"
    dest = urllib.parse.quote(f"{stops[-1][0]},{stops[-1][1]}")
    waypoints = "|".join(f"{lat},{lng}" for lat, lng in stops[:-1][:8])
    wp = urllib.parse.quote(waypoints) if waypoints else ""
    url = f"https://www.google.com/maps/dir/?api=1&origin={origin}&destination={dest}&travelmode=driving"
    if wp:
        url += f"&waypoints={wp}"
    return url

def _can_edit_data():
    return st.session_state.get("user_role", "") in ["admin", "manager", "staff"]

def render_login_page(auth_ready: bool):
    st.markdown(textwrap.dedent("""
    <style>
    .stApp {
        background:
            radial-gradient(circle at 12% 18%, rgba(147, 197, 253, 0.34) 0%, transparent 24%),
            radial-gradient(circle at 85% 16%, rgba(191, 219, 254, 0.30) 0%, transparent 26%),
            radial-gradient(circle at 72% 78%, rgba(125, 211, 252, 0.22) 0%, transparent 28%),
            linear-gradient(135deg, #dbeafe 0%, #bfdbfe 38%, #93c5fd 100%);
    }
    [data-testid="stHeader"] { background: transparent; }
    .block-container {
        padding-top: 0.05rem !important;
        padding-bottom: 0.2rem !important;
        max-width: 1320px !important;
    }
    .login-shell { position: relative; min-height: 0; height: 0; }
    .login-orb {
        position: fixed; border-radius: 999px; filter: blur(56px); opacity: 0.45;
        pointer-events: none; z-index: 0; animation: floatOrb 11s ease-in-out infinite;
    }
    .login-orb.orb1 { width: 250px; height: 250px; left: 4%; top: 10%; background: rgba(96, 165, 250, 0.45); }
    .login-orb.orb2 { width: 320px; height: 320px; right: 6%; top: 16%; background: rgba(125, 211, 252, 0.35); animation-delay: 2s; }
    .login-orb.orb3 { width: 300px; height: 300px; left: 26%; bottom: 4%; background: rgba(191, 219, 254, 0.42); animation-delay: 4s; }
    @keyframes floatOrb {
        0% { transform: translate(0, 0) scale(1); }
        50% { transform: translate(14px, -18px) scale(1.05); }
        100% { transform: translate(0, 0) scale(1); }
    }
    .glass-card {
        position: relative; z-index: 1;
        background: rgba(255,255,255,0.28);
        border: 1px solid rgba(255,255,255,0.46);
        box-shadow: 0 18px 60px rgba(30, 64, 175, 0.16);
        border-radius: 30px;
        backdrop-filter: blur(18px);
        -webkit-backdrop-filter: blur(18px);
        padding: 28px;
    }
    .brand-row { display:flex; gap:18px; align-items:center; margin-bottom: 14px; }
    .brand-logo {
        width: 78px; height: 78px; border-radius: 22px; display:flex; align-items:center; justify-content:center;
        background: linear-gradient(135deg, #2563eb, #38bdf8); color:#fff; font-size: 34px; font-weight: 800;
        box-shadow: 0 12px 30px rgba(37, 99, 235, 0.25);
    }
    .brand-eyebrow { color: #1d4ed8; font-weight: 800; letter-spacing: .12em; font-size: 12px; text-transform: uppercase; }
    .brand-title { color: #0f172a; font-size: 34px; line-height: 1.05; font-weight: 800; margin: 2px 0 0 0; }
    .brand-sub { color: #334155; font-size: 15px; line-height: 1.75; margin-top: 8px; }
    .feature-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:14px; margin-top:22px; }
    .feature-item { border-radius: 20px; padding: 16px 14px; background: rgba(255,255,255,0.34); border: 1px solid rgba(255,255,255,0.52); min-height:112px; }
    .feature-icon { font-size:24px; margin-bottom:8px; }
    .feature-title { color:#0f172a; font-size:15px; font-weight:800; margin-bottom:4px; }
    .feature-text { color:#475569; font-size:12.5px; line-height:1.55; }
    .login-right-panel { padding: 10px 2px 0 2px; }
    .login-panel-title { color:#0f172a; font-size:26px; font-weight:800; margin-bottom:8px; }
    .login-panel-sub { color:#334155; font-size:14px; line-height:1.65; margin-bottom:18px; }
    .login-footer { text-align:center; color:#334155; font-size:12.5px; margin-top:16px; padding-bottom: 8px; }
    .login-footer a { color:#1d4ed8; text-decoration:none; font-weight:700; }
    .ms-login-link {
        display:block; text-align:center; padding:12px 16px; border-radius:14px; text-decoration:none; font-weight:800;
        background: linear-gradient(135deg, #2563eb, #3b82f6); color:white; border: 1px solid rgba(37,99,235,0.2);
        box-shadow: 0 10px 22px rgba(37,99,235,0.18);
    }
    .ms-login-link:hover { filter: brightness(1.03); }
    .loading-overlay {
        display:none; position: fixed; inset:0; background: rgba(219, 234, 254, 0.72); backdrop-filter: blur(8px);
        z-index: 99999; align-items:center; justify-content:center; flex-direction:column; gap:12px;
    }
    .loading-overlay.show { display:flex; }
    .loading-spinner {
        width:54px; height:54px; border-radius:999px; border:5px solid rgba(37,99,235,0.18); border-top-color:#2563eb;
        animation: spin 1s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .loading-text { color:#1e3a8a; font-weight:800; font-size:15px; }
    @media (max-width: 980px) { .feature-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
    @media (max-width: 640px) { .feature-grid { grid-template-columns: 1fr; } .brand-title { font-size:28px; } }
    
[data-testid="stAppViewBlockContainer"], .block-container {
    padding-top: 0.15rem !important;
    padding-bottom: 0 !important;
}
div[data-testid="column"] {
    padding-top: 0 !important;
}
.glass-card {
    margin-top: 0 !important;
}

</style>
<div class="loading-overlay" id="login-loading-overlay">
    <div class="loading-spinner"></div>
    <div class="loading-text">กำลังพาไปหน้า Microsoft 365...</div>
</div>
<script>
function showLoginLoading(){
    const el = window.parent.document.getElementById('login-loading-overlay') || document.getElementById('login-loading-overlay');
    if(el){ el.classList.add('show'); }
}
</script>
<div class="login-shell">
    <div class="login-orb orb1"></div>
    <div class="login-orb orb2"></div>
    <div class="login-orb orb3"></div>
</div>
    """), unsafe_allow_html=True)

    left, right = st.columns([1.35, 1])
    with left:
        st.markdown(textwrap.dedent("""
        <div class="glass-card">
            <div class="brand-row">
                <div class="brand-logo">📊</div>
                <div>
                    <div class="brand-eyebrow">Optimal Group Platform</div>
                    <div class="brand-title">Sales Territory Dashboard</div>
                    <div class="brand-sub">ระบบบริหารข้อมูลลูกค้า แผนที่ Budget การเข้าถึงข้อมูล และการส่งออกไฟล์ สำหรับทีมงานในแต่ละแผนกของบริษัท</div>
                </div>
            </div>
            <div class="feature-grid">
                <div class="feature-item"><div class="feature-icon">📊</div><div class="feature-title">Dashboard</div><div class="feature-text">ดูภาพรวมยอดขาย ลูกค้า และ Insight แยกตามแผนก</div></div>
                <div class="feature-item"><div class="feature-icon">🗺️</div><div class="feature-title">แผนที่</div><div class="feature-text">ดูหมุดลูกค้า เส้นทาง และข้อมูล Plus Code</div></div>
                <div class="feature-item"><div class="feature-icon">🎯</div><div class="feature-title">Budget</div><div class="feature-text">เปรียบเทียบ Budget และ Actual พร้อมวิเคราะห์ Gap</div></div>
                <div class="feature-item"><div class="feature-icon">☁️</div><div class="feature-title">SharePoint</div><div class="feature-text">โหลดและบันทึกไฟล์แยกตามแผนกจาก SharePoint</div></div>
                <div class="feature-item"><div class="feature-icon">🔐</div><div class="feature-title">สิทธิ์</div><div class="feature-text">ควบคุมการเข้าถึงตาม Microsoft 365, Group และ Role</div></div>
                <div class="feature-item"><div class="feature-icon">📤</div><div class="feature-title">Export</div><div class="feature-text">ส่งออก Template, ข้อมูลลูกค้า และ Audit Log ได้ทันที</div></div>
            </div>
        </div>
        """), unsafe_allow_html=True)

    with right:
        st.markdown('<div class="login-right-panel">', unsafe_allow_html=True)
        st.markdown('<div class="login-panel-title">เข้าสู่ระบบ</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-panel-sub">เข้าสู่ระบบด้วย Microsoft 365 เพื่อกำหนดสิทธิ์และแผนกอัตโนมัติจากบัญชีองค์กร</div>', unsafe_allow_html=True)
        st.markdown('#### Microsoft 365')
        if auth_ready:
            login_url = _build_login_url()
            st.markdown(
                f"""
                <a href="{login_url}" target="_self" onclick="showLoginLoading()" class="ms-login-link">
                    🔵 Sign in with Microsoft 365
                </a>
                """,
                unsafe_allow_html=True,
            )
            st.caption('ใช้บัญชีองค์กรเพื่อตรวจสอบสิทธิ์และดึงกลุ่มแผนกอัตโนมัติ')
        else:
            st.button('🔵 Microsoft 365 Not Configured', disabled=True, use_container_width=True)
            st.caption('ยังไม่ได้ตั้งค่า TENANT_ID / CLIENT_ID / CLIENT_SECRET / REDIRECT_URI')
        st.markdown(textwrap.dedent("""
        <div class="login-footer">
            Version 2026.04 • IT Support: <a href="mailto:it@optimal.co.th">it@optimal.co.th</a>
        </div>
        """), unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# LOGIN PAGE GATE
# ═══════════════════════════════════════════════════════════════════════════════
auth_ready = _auth_configured()
_restore_session_from_cookies()
_complete_login_from_query()
_restore_session_from_cookies()
is_logged_in = _session_logged_in()

if not st.session_state.dept and not (auth_ready and is_logged_in):
    render_login_page(auth_ready)
    st.stop()

if auth_ready and is_logged_in and not _user_email_allowed():
    st.title("⛔ ไม่ได้รับสิทธิ์เข้าใช้งาน")
    st.error("บัญชี Microsoft 365 นี้ไม่มีสิทธิ์เข้าใช้งานระบบ")
    st.caption("อนุญาตเฉพาะโดเมน: " + ", ".join(_get_allowed_email_domains()))
    if st.button("🚪 Log out"):
        _auth_logout()
    st.stop()

if auth_ready and is_logged_in:
    st.session_state.user_email = _get_user_email()
    st.session_state.user_name = _get_user_name()
    if st.session_state.get("auth_access_token"):
        user_groups = _get_user_groups()
        resolved_role, resolved_dept = _resolve_role_and_dept(st.session_state.user_email, user_groups)
    else:
        user_groups = []
        resolved_role = st.session_state.get("user_role")
        resolved_dept = st.session_state.get("dept")

    if not resolved_role:
        st.title("⛔ ไม่มีสิทธิ์เข้าใช้งาน")
        st.error("ไม่พบอีเมลนี้ในระบบสิทธิ์ หรือบัญชีนี้ไม่ได้อยู่ใน Group แผนกที่กำหนด")
        st.caption("ตรวจสอบว่า user อยู่ใน Group แผนกของ Microsoft 365 และถ้าเป็นหัวหน้าให้เพิ่ม email ใน HEAD_EMAIL_TO_DEPT")
        with st.expander("ดูข้อมูลสำหรับตรวจสอบ"):
            st.write({"email": st.session_state.user_email, "groups": user_groups})
        if st.button("🚪 Log out"):
            _auth_logout()
        st.stop()

    st.session_state.user_role = resolved_role
    st.session_state.is_admin = (resolved_role == "admin")
    _set_auth_cookies(
        email=st.session_state.get("user_email", ""),
        name=st.session_state.get("user_name", ""),
        role=resolved_role,
        dept=(resolved_dept or st.session_state.get("dept") or ""),
        is_admin=(resolved_role == "admin"),
    )

    target_dept = st.session_state.dept
    if resolved_role == "admin":
        if not target_dept:
            target_dept = DEPARTMENTS[0]
    else:
        target_dept = resolved_dept

    if st.session_state.dept != target_dept:
        st.session_state.dept = target_dept
        st.session_state.sp_file = None
        st.session_state.df = EMPTY_DF
        st.session_state.sp_file_last_modified = ""
        st.session_state.sp_file_etag = ""
        append_audit_log("login_role_resolved", f"m365 role={resolved_role} dept={target_dept}", target_dept or "")
        st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════

st.sidebar.image("https://img.icons8.com/fluency/96/combo-chart.png", width=80)
st.sidebar.title("📂 เมนูหลัก")

role_label = _role_label()
allowed_menus = ["🏢 ข้อมูลบริษัทลูกค้า", "🧠 Executive Report", "✏️ แก้ไข / เพิ่มข้อมูล"]
if _can_view_dashboard():
    allowed_menus.insert(0, "📊 Dashboard")

menu = st.sidebar.radio("", allowed_menus, label_visibility="collapsed")
if st.session_state.get("last_menu_logged") != menu:
    append_audit_log("page_view", menu, st.session_state.get("dept") or "")
    st.session_state["last_menu_logged"] = menu
st.sidebar.divider()

st.sidebar.subheader("🔐 บัญชีผู้ใช้งาน")
if auth_ready:
    st.sidebar.success(f"👤 {st.session_state.get('user_name') or _get_user_name()}")
    if st.session_state.get("user_email"):
        st.sidebar.caption(st.session_state.get("user_email"))
    st.sidebar.info(f"สิทธิ์: {role_label}")

    if st.session_state.get("user_role") == "admin":
        switch = st.sidebar.selectbox("เลือกแผนก", DEPARTMENTS,
                                      index=DEPARTMENTS.index(st.session_state.dept) if st.session_state.dept in DEPARTMENTS else 0,
                                      key="dept_switch_auth",
                                      format_func=lambda x: DEPARTMENT_LABELS.get(x, x))
        if switch != st.session_state.dept:
            st.session_state.dept = switch
            st.session_state.sp_file = None
            st.session_state.df = EMPTY_DF
            st.session_state.sp_file_last_modified = ""
            st.session_state.sp_file_etag = ""
            append_audit_log("switch_dept", f"admin switch to {switch}", switch)
            st.rerun()
        st.sidebar.success(f"📁 แผนกที่กำลังดู: **{_dept_label(st.session_state.dept)}**")
        st.sidebar.caption("สิทธิ์ Admin: ดูได้ทุกแผนก")
    else:
        st.sidebar.success(f"📁 แผนก: **{_dept_label(st.session_state.dept)}**")
        if _can_view_dashboard():
            st.sidebar.caption("สิทธิ์หัวหน้าแผนก: ดู Dashboard ได้เฉพาะแผนกตัวเอง")
        else:
            st.sidebar.caption("สิทธิ์ลูกทีม: ดูข้อมูลลูกค้าและแก้ไขข้อมูลเฉพาะแผนกตัวเอง")

    if st.sidebar.button("🚪 ออกจากระบบ", use_container_width=True):
        append_audit_log("logout", "m365 logout", st.session_state.get("dept") or "")
        for k in ["dept", "sp_file", "df", "is_admin", "user_role", "user_email", "user_name"]:
            st.session_state[k] = None if k not in ["df", "user_role", "user_email", "user_name"] else (EMPTY_DF if k=="df" else ("staff" if k=="user_role" else ""))
        _auth_logout()
else:
    st.sidebar.subheader("🔐 เข้าสู่ระบบแผนก")
    if not st.session_state.dept:
        sel_dept = st.sidebar.selectbox("เลือกแผนก", [""] + DEPARTMENTS, key="sel_dept_sb")
        admin_pw = st.sidebar.text_input("รหัส Admin (ว่าง = ดูแลแผนกตนเอง)", type="password", key="admin_pw_sb")
        if st.sidebar.button("เข้าสู่ระบบ", type="primary", use_container_width=True):
            if sel_dept:
                st.session_state.dept = sel_dept
                st.session_state.is_admin = (admin_pw == ADMIN_PASSWORD)
                st.session_state.user_role = "admin" if st.session_state.is_admin else "manager"
                st.session_state.user_name = "Local User"
                st.session_state.user_email = ""
                st.session_state.sp_file = None
                st.session_state.df = EMPTY_DF
                append_audit_log("login", f"login to {sel_dept}", sel_dept)
                st.rerun()
            else:
                st.sidebar.warning("กรุณาเลือกแผนก")
    else:
        st.sidebar.success(f"📁 แผนก: **{_dept_label(st.session_state.dept)}**")
        st.sidebar.info(f"สิทธิ์: {_role_label()}")
        if st.session_state.is_admin:
            switch = st.sidebar.selectbox("สลับแผนก", DEPARTMENTS,
                                          index=DEPARTMENTS.index(st.session_state.dept),
                                          key="dept_switch")
            if switch != st.session_state.dept:
                st.session_state.dept = switch
                st.session_state.sp_file = None
                st.session_state.df = EMPTY_DF
                append_audit_log("switch_dept", f"switch to {switch}", switch)
                st.rerun()
        if st.sidebar.button("🚪 ออกจากระบบ", use_container_width=True):
            append_audit_log("logout", "local logout", st.session_state.get("dept") or "")
            for k in ["dept", "sp_file", "df", "is_admin", "user_role", "user_email", "user_name"]:
                st.session_state[k] = None if k not in ["df", "user_role", "user_email", "user_name"] else (EMPTY_DF if k=="df" else ("staff" if k=="user_role" else ""))
            st.rerun()

# ── File Selector + SharePoint Load ──────────────────────────────────────────
st.sidebar.divider()
st.sidebar.subheader("📁 จัดการไฟล์")

if st.session_state.dept:
    try:
        files = sp_list_files(st.session_state.dept)
        if files:
            files = sorted(files, key=lambda f: f.get("lastModifiedDateTime", ""), reverse=True)
            fnames = [f["name"] for f in files]

            default_idx = 0
            if st.session_state.sp_file in fnames:
                default_idx = fnames.index(st.session_state.sp_file)

            chosen = st.sidebar.selectbox("ไฟล์ใน SharePoint", fnames, index=default_idx, key="file_sel")

            selected_meta = next((f for f in files if f["name"] == chosen), {})
            selected_modified = str(selected_meta.get("lastModifiedDateTime", "") or "")
            selected_etag = str(selected_meta.get("eTag", "") or "")

            prev_file = str(st.session_state.get("sp_file") or "")
            prev_modified = str(st.session_state.get("sp_file_last_modified") or "")
            prev_etag = str(st.session_state.get("sp_file_etag") or "")

            file_changed = prev_file != chosen
            version_changed = (
                (selected_modified and selected_modified != prev_modified)
                or (selected_etag and selected_etag != prev_etag)
            )

            df_current = st.session_state.get("df")
            auto_load_needed = (
                file_changed
                or version_changed
                or df_current is None
                or df_current.empty
            )

            if auto_load_needed:
                reason = []
                if file_changed:
                    reason.append("file_changed")
                if version_changed:
                    reason.append("version_changed")
                if df_current is None:
                    reason.append("df_none")
                elif getattr(df_current, "empty", False):
                    reason.append("df_empty")

                st.session_state.sp_file = chosen
                st.session_state.df = sp_load(st.session_state.dept, chosen)
                st.session_state.sp_file_last_modified = selected_modified
                st.session_state.sp_file_etag = selected_etag
                st.session_state.last_refresh = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                append_audit_log(
                    "load_sharepoint",
                    f"auto-load {chosen} | reason={','.join(reason) if reason else 'unknown'} | modified={selected_modified}",
                    st.session_state.dept
                )
                st.rerun()

            c1, c2 = st.sidebar.columns(2)

            with c1:
                if st.button("🔄 รีโหลดไฟล์", use_container_width=True):
                    st.session_state.df = sp_load(st.session_state.dept, chosen)
                    st.session_state.sp_file = chosen
                    st.session_state.sp_file_last_modified = selected_modified
                    st.session_state.sp_file_etag = selected_etag
                    st.session_state.last_refresh = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    append_audit_log("reload_sharepoint", chosen, st.session_state.dept)
                    st.rerun()

            with c2:
                if st.button("🧹 Force Refresh", use_container_width=True):
                    st.session_state.df = EMPTY_DF
                    st.session_state.sp_file = chosen
                    st.session_state.sp_file_last_modified = ""
                    st.session_state.sp_file_etag = ""
                    st.session_state.last_refresh = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    append_audit_log("force_refresh_prepare", chosen, st.session_state.dept)
                    st.rerun()

            if st.session_state.sp_file:
                st.sidebar.caption(
                    f"✅ โหลดแล้ว: **{st.session_state.sp_file}** ({len(st.session_state.df):,} ราย)"
                )
                if selected_modified:
                    st.sidebar.caption(f"🕒 SharePoint modified: {selected_modified}")
                st.sidebar.caption(f"🕒 App refresh: {st.session_state.last_refresh}")

        else:
            st.sidebar.info(f"ไม่พบไฟล์ในโฟลเดอร์ {st.session_state.dept}")
    except Exception as e:
        st.sidebar.error(f"SharePoint error: {e}")
        with st.sidebar.expander("🔍 รายละเอียด error (คลิกเพื่อดู)"):
            st.code(traceback.format_exc())

live_sync_enabled = st.sidebar.toggle("🔄 Live Sync SharePoint", value=st.session_state.get("live_sync_enabled", True))
st.session_state["live_sync_enabled"] = live_sync_enabled
live_sync_interval = st.sidebar.selectbox("ช่วงเวลา sync (วินาที)", [15, 30, 60, 120], index=[15,30,60,120].index(st.session_state.get("live_sync_interval", 60)) if st.session_state.get("live_sync_interval", 60) in [15,30,60,120] else 2)
st.session_state["live_sync_interval"] = int(live_sync_interval)
if live_sync_enabled:
    st.sidebar.caption(f"Live sync ทุก {live_sync_interval} วินาที เมื่อไม่ได้อยู่ระหว่างแก้ไข")

st.sidebar.download_button(
    "⬇️ ดาวน์โหลด Template (.xlsx)",
    data=make_template(),
    file_name="customer_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.sidebar.divider()
with st.sidebar.expander("🛡️ System / Production Status", expanded=False):
    st.write(f"**Environment:** {APP_ENV}")
    st.write(f"**Department:** {st.session_state.get('dept') or '-'}")
    st.write(f"**Current file:** {st.session_state.get('sp_file') or '-'}")
    st.write(f"**Last refresh:** {st.session_state.get('last_refresh') or '-'}")
    st.write(f"**Records in memory:** {len(st.session_state.get('df', EMPTY_DF)):,}")
    if os.path.exists("sales_dashboard_audit_log.csv"):
        try:
            _audit_df = pd.read_csv("sales_dashboard_audit_log.csv")
            st.write(f"**Audit rows:** {len(_audit_df):,}")
            if st.button("☁️ Push Audit Log to SharePoint", use_container_width=True):
                push_audit_log_to_sharepoint()
            st.download_button(
                "⬇️ Download Audit Log",
                data=_audit_df.to_csv(index=False, encoding="utf-8-sig"),
                file_name="sales_dashboard_audit_log.csv",
                mime="text/csv",
                use_container_width=True,
            )
        except Exception:
            pass

uploaded = st.sidebar.file_uploader("📤 อัปโหลดไฟล์ (xlsx / csv)", type=["xlsx", "csv"])
if uploaded:
    try:
        if uploaded.name.endswith(".csv"):
            raw = pd.read_csv(uploaded)
        else:
            raw = pd.read_excel(uploaded)
        if "Customer Name" in raw.columns or "Customer name" in raw.columns:
            raw = raw.rename(columns={"Customer name": "Customer Name",
                                       "Salesperson (2026)": "Salesperson",
                                       "Business type": "Industry"})
            for col in TEMPLATE_COLS:
                if col not in raw.columns: raw[col] = ""
            if "Plus_Code" in raw.columns:
                raw["Plus_Code"] = raw["Plus_Code"].apply(clean_plus_code)
            else:
                raw["Plus_Code"] = ""
            if "Address" not in raw.columns:
                raw["Address"] = ""
            raw["Address"] = raw.apply(lambda r: merge_address_parts(r.get("Address", ""), r.get("Plus_Code", "")), axis=1)
            def _enrich(row):
                addr_for_parse = merge_address_parts(row.get("Address", ""), row.get("Plus_Code", ""))
                if pd.notna(row.get("Province")) and str(row.get("Province")).strip(): return row
                sub, dis, prov, reg = parse_address(str(addr_for_parse))
                if not str(row.get("Sub-district", "")).strip(): row["Sub-district"] = sub
                if not str(row.get("District", "")).strip():     row["District"] = dis
                if not str(row.get("Province", "")).strip():     row["Province"] = prov
                if not str(row.get("Region", "")).strip():       row["Region"] = reg
                return row
            raw = raw.apply(_enrich, axis=1)
            raw["Region_TH"] = raw["Region"].map(REGION_EN_TO_TH).fillna("ไม่ระบุ")
            st.session_state.df = raw
            st.session_state.sp_file = uploaded.name
        else:
            raw_xl = pd.read_excel(uploaded, sheet_name=None)
            st.session_state.df = build_df_from_original(raw_xl)
            st.session_state.sp_file = uploaded.name
        append_audit_log("manual_upload", uploaded.name, st.session_state.get("dept", ""))
        st.sidebar.success(f"✅ โหลดสำเร็จ ({len(st.session_state.df):,} ราย)")
    except Exception as e:
        st.sidebar.error(f"❌ {e}")

if st.session_state.get("live_sync_enabled") and not st.session_state.get("editing_idx") and menu != "✏️ แก้ไข / เพิ่มข้อมูล":
    render_live_sync_autorefresh(True, int(st.session_state.get("live_sync_interval", 60)))

df = st.session_state.df

if not st.session_state.dept:
    st.title("📊 Sales Territory Dashboard")
    st.info("👈 กรุณาเลือกแผนกและเข้าสู่ระบบก่อนใช้งาน")
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════════
# MENU 1 – DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

if menu == "📊 Dashboard":
    if not _can_view_dashboard():
        st.error("คุณไม่มีสิทธิ์ดูหน้า Dashboard")
        st.stop()
    _scroll_top()
    st.title("📊 Sales Territory Dashboard")
    st.caption("ภาพรวมยอดขายและลูกค้าแยกตามภูมิภาค")

    if df.empty or "Customer Name" not in df.columns:
        st.info("📂 กรุณาโหลดไฟล์จาก SharePoint ก่อน (ด้านซ้าย)")
        st.stop()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("👥 ลูกค้าทั้งหมด", f"{len(df):,} ราย")
    k2.metric("💰 ยอดขายรวม", f"฿{df['Sales/Year'].sum()/1e6:,.1f} M")
    k3.metric("📈 เฉลี่ย/ลูกค้า", f"฿{df['Sales/Year'].mean()/1e6:,.2f} M")
    k4.metric("🗺️ ระบุภูมิภาคได้", f"{(df['Region'] != 'Unknown').sum():,} ราย")
    st.divider()

    rsum = (df.groupby("Region_TH")
            .agg(count=("Customer Name", "count"), total_sales=("Sales/Year", "sum"))
            .reset_index().rename(columns={"Region_TH": "region"})
            .sort_values("total_sales", ascending=False))

    cl, cr = st.columns([1.2, 1])
    with cl:
        st.subheader("💹 ยอดขายรวมแยกตามภูมิภาค")
        f1 = px.bar(rsum, x="region", y="total_sales", color="region",
                    color_discrete_map=REGION_COLORS,
                    text=rsum["total_sales"].apply(lambda v: f"฿{v/1e6:.1f}M"),
                    labels={"total_sales": "ยอดขาย (บาท)", "region": ""})
        f1.update_traces(textposition="outside")
        f1.update_layout(showlegend=False, yaxis_tickformat=",.0f",
                         plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                         margin=dict(t=20, b=10), height=380)
        st.plotly_chart(f1, use_container_width=True)
    with cr:
        st.subheader("🗺️ สัดส่วนลูกค้าตามภูมิภาค")
        f2 = px.pie(rsum, names="region", values="count", color="region",
                    color_discrete_map=REGION_COLORS, hole=0.45)
        f2.update_traces(textinfo="label+percent", pull=[0.03] * len(rsum))
        f2.update_layout(showlegend=False, margin=dict(t=20, b=10),
                         height=380, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(f2, use_container_width=True)

    st.subheader("📋 ตารางสรุปแยกภูมิภาค")
    tbl = rsum.rename(columns={"region": "ภูมิภาค", "count": "จำนวนลูกค้า",
                                "total_sales": "ยอดขายรวม"}).copy()
    tbl["ยอดขายรวม"] = tbl["ยอดขายรวม"].apply(lambda v: f"฿{v:,.0f}")
    st.dataframe(tbl, use_container_width=True, hide_index=True)
    st.divider()

    st.subheader("🏭 ยอดขายแยกตาม Industry")
    ind = (df.groupby("Industry").agg(total_sales=("Sales/Year", "sum"))
           .reset_index().sort_values("total_sales", ascending=True))
    f3 = px.bar(ind, x="total_sales", y="Industry", orientation="h",
                text=ind["total_sales"].apply(lambda v: f"฿{v/1e6:.1f}M"),
                color="total_sales", color_continuous_scale="Blues",
                labels={"total_sales": "ยอดขาย (บาท)"})
    f3.update_traces(textposition="outside")
    f3.update_layout(coloraxis_showscale=False, xaxis_tickformat=",.0f",
                     plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                     margin=dict(t=10, b=10), height=300)
    st.plotly_chart(f3, use_container_width=True)

    st.subheader("👤 ยอดขายแยกตาม Salesperson")
    sp_g = (df.groupby("Salesperson").agg(total_sales=("Sales/Year", "sum"))
            .reset_index().sort_values("total_sales", ascending=False))
    f4 = px.bar(sp_g, x="Salesperson", y="total_sales",
                text=sp_g["total_sales"].apply(lambda v: f"฿{v/1e6:.1f}M"),
                color="total_sales", color_continuous_scale="Teal",
                labels={"total_sales": "ยอดขาย (บาท)", "Salesperson": ""})
    f4.update_traces(textposition="outside")
    f4.update_layout(coloraxis_showscale=False, yaxis_tickformat=",.0f",
                     plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                     margin=dict(t=10, b=10), height=350)
    st.plotly_chart(f4, use_container_width=True)

    st.divider()
    st.subheader("🧠 Executive Insights / Opportunity Ranking")
    insight_df = df.copy()
    insight_df["Sales/Year"] = pd.to_numeric(insight_df.get("Sales/Year", 0), errors="coerce").fillna(0)
    insight_df["Budget_kg"] = pd.to_numeric(insight_df.get("Budget_kg", 0), errors="coerce").fillna(0)
    insight_df["Actual_kg"] = pd.to_numeric(insight_df.get("Actual_kg", 0), errors="coerce").fillna(0)
    insight_df["LastYear_kg"] = pd.to_numeric(insight_df.get("LastYear_kg", 0), errors="coerce").fillna(0)
    insight_df["gap_kg"] = (insight_df["Budget_kg"] - insight_df["Actual_kg"]).clip(lower=0)
    insight_df["achievement_pct"] = insight_df.apply(lambda r: (r["Actual_kg"] / r["Budget_kg"] * 100) if r["Budget_kg"] > 0 else 0, axis=1)
    insight_df["yoy_pct"] = insight_df.apply(lambda r: ((r["Actual_kg"] - r["LastYear_kg"]) / r["LastYear_kg"] * 100) if r["LastYear_kg"] > 0 else 0, axis=1)
    insight_df["opportunity_score"] = (
        insight_df["gap_kg"].rank(pct=True).fillna(0) * 45
        + (100 - insight_df["achievement_pct"].clip(upper=100)).rank(pct=True).fillna(0) * 35
        + insight_df["Sales/Year"].rank(pct=True).fillna(0) * 20
    ).round(1)

    ex1, ex2 = st.columns([1.25, 1])
    with ex1:
        top_opp = insight_df.sort_values(["opportunity_score", "gap_kg"], ascending=False).head(12).copy()
        top_opp_show = top_opp[["Customer Name", "Salesperson", "Industry", "gap_kg", "achievement_pct", "opportunity_score"]].copy()
        top_opp_show["gap_kg"] = top_opp_show["gap_kg"].apply(lambda v: f"{int(v):,} kg")
        top_opp_show["achievement_pct"] = top_opp_show["achievement_pct"].apply(lambda v: f"{v:.1f}%")
        top_opp_show["opportunity_score"] = top_opp_show["opportunity_score"].apply(lambda v: f"{v:.1f}")
        st.markdown("**🎯 ลูกค้าที่ควรเร่งตามมากที่สุด**")
        st.dataframe(top_opp_show.rename(columns={
            "gap_kg": "Gap", "achievement_pct": "Achievement", "opportunity_score": "Score"
        }), use_container_width=True, hide_index=True, height=360)
    with ex2:
        sp_perf = insight_df.groupby("Salesperson").agg(
            total_sales=("Sales/Year", "sum"),
            budget_kg=("Budget_kg", "sum"),
            actual_kg=("Actual_kg", "sum"),
            customers=("Customer Name", "count"),
            avg_score=("opportunity_score", "mean"),
        ).reset_index()
        sp_perf["achievement_pct"] = sp_perf.apply(lambda r: (r["actual_kg"] / r["budget_kg"] * 100) if r["budget_kg"] > 0 else 0, axis=1)
        fig_exec = px.scatter(
            sp_perf, x="achievement_pct", y="total_sales", size="customers",
            hover_name="Salesperson", text="Salesperson", color="avg_score",
            color_continuous_scale="Turbo", labels={"achievement_pct": "Achievement %", "total_sales": "Total Sales"}
        )
        fig_exec.update_traces(textposition="top center")
        fig_exec.update_layout(height=360, margin=dict(t=15, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_exec, use_container_width=True)

    if "Budget_kg" in df.columns and df["Budget_kg"].sum() > 0:
        st.divider()
        st.subheader("📦 Budget vs Actual & การวิเคราะห์ ปี 2026")
        bdf = df[df["Budget_kg"] > 0].copy()
        bdf["Actual_kg"] = pd.to_numeric(bdf.get("Actual_kg", 0), errors="coerce").fillna(0)

        total_budget = int(bdf["Budget_kg"].sum())
        total_actual = int(bdf["Actual_kg"].sum())
        gap  = total_actual - total_budget
        pct  = (total_actual / total_budget * 100) if total_budget > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("🎯 Budget รวม",    f"{total_budget:,} kg")
        k2.metric("✅ Actual รวม",    f"{total_actual:,} kg",  delta=f"{gap:+,} kg")
        k3.metric("📊 Achievement",   f"{pct:.1f}%")
        k4.metric("📋 บริษัทที่มีข้อมูล", f"{len(bdf):,} ราย")

        gc1, gc2 = st.columns(2)
        with gc1:
            st.markdown("**📊 Budget vs Actual แยก Industry**")
            b_ind = (bdf.groupby("Industry")
                     .agg(Budget=("Budget_kg", "sum"), Actual=("Actual_kg", "sum"))
                     .reset_index().sort_values("Budget", ascending=True))
            fig_ind = go.Figure()
            fig_ind.add_trace(go.Bar(name="Budget", y=b_ind["Industry"], x=b_ind["Budget"],
                                     orientation="h", marker_color="#fb923c",
                                     text=b_ind["Budget"].apply(lambda v: f"{int(v):,}"),
                                     textposition="outside"))
            fig_ind.add_trace(go.Bar(name="Actual", y=b_ind["Industry"], x=b_ind["Actual"],
                                     orientation="h", marker_color="#22c55e",
                                     text=b_ind["Actual"].apply(lambda v: f"{int(v):,}" if v > 0 else ""),
                                     textposition="outside"))
            fig_ind.update_layout(barmode="group", height=280, xaxis_tickformat=",.0f",
                                  plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                  margin=dict(t=10, b=10), legend=dict(orientation="h", y=1.1))
            st.plotly_chart(fig_ind, use_container_width=True)

        with gc2:
            st.markdown("**📊 Budget vs Actual แยก Salesperson**")
            b_sp2 = (bdf.groupby("Salesperson")
                     .agg(Budget=("Budget_kg", "sum"), Actual=("Actual_kg", "sum"))
                     .reset_index().sort_values("Budget", ascending=False))
            fig_sp2 = go.Figure()
            fig_sp2.add_trace(go.Bar(name="Budget", x=b_sp2["Salesperson"], y=b_sp2["Budget"],
                                     marker_color="#fb923c",
                                     text=b_sp2["Budget"].apply(lambda v: f"{int(v):,}"),
                                     textposition="outside"))
            fig_sp2.add_trace(go.Bar(name="Actual", x=b_sp2["Salesperson"], y=b_sp2["Actual"],
                                     marker_color="#22c55e",
                                     text=b_sp2["Actual"].apply(lambda v: f"{int(v):,}" if v > 0 else ""),
                                     textposition="outside"))
            fig_sp2.update_layout(barmode="group", height=280, yaxis_tickformat=",.0f",
                                  plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                  margin=dict(t=10, b=10), legend=dict(orientation="h", y=1.1))
            st.plotly_chart(fig_sp2, use_container_width=True)

        st.markdown("**🔍 Gap & Achievement Analysis รายบริษัท**")
        gap_df = bdf[["Customer Name", "Salesperson", "Industry", "Budget_kg", "Actual_kg"]].copy()
        gap_df["Gap (kg)"]     = gap_df["Actual_kg"] - gap_df["Budget_kg"]
        gap_df["Achievement%"] = gap_df.apply(
            lambda r: round(r["Actual_kg"] / r["Budget_kg"] * 100, 1) if r["Budget_kg"] > 0 else 0.0, axis=1)
        gap_df["สถานะ"] = gap_df.apply(
            lambda r: "✅ เกิน Budget" if r["Achievement%"] >= 100
            else ("🟡 ใกล้ถึง" if r["Achievement%"] >= 70 else "🔴 ต้องเร่ง"), axis=1)

        gauge_pct = min(pct, 150)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=pct,
            number={"suffix": "%", "font": {"size": 36}},
            delta={"reference": 100, "suffix": "%"},
            title={"text": "Achievement vs Budget (รวม)", "font": {"size": 14}},
            gauge={
                "axis": {"range": [0, 150], "ticksuffix": "%"},
                "bar": {"color": "#22c55e" if pct >= 100 else ("#f59e0b" if pct >= 70 else "#ef4444")},
                "steps": [{"range": [0, 70], "color": "#fee2e2"},
                           {"range": [70, 100], "color": "#fef9c3"},
                           {"range": [100, 150], "color": "#dcfce7"}],
                "threshold": {"line": {"color": "#1e3a5f", "width": 4}, "thickness": 0.75, "value": 100},
            }
        ))
        fig_gauge.update_layout(height=260, margin=dict(t=30, b=10, l=30, r=30),
                                paper_bgcolor="rgba(0,0,0,0)")

        gg1, gg2 = st.columns([1.2, 2])
        with gg1:
            st.plotly_chart(fig_gauge, use_container_width=True)
        with gg2:
            top_ach = gap_df[gap_df["Actual_kg"] > 0].nlargest(8, "Achievement%")[
                ["Customer Name", "Budget_kg", "Actual_kg", "Achievement%", "สถานะ"]].copy()
            top_ach["Budget_kg"]    = top_ach["Budget_kg"].apply(lambda v: f"{int(v):,} kg")
            top_ach["Actual_kg"]    = top_ach["Actual_kg"].apply(lambda v: f"{int(v):,} kg")
            top_ach["Achievement%"] = top_ach["Achievement%"].apply(lambda v: f"{v:.1f}%")
            st.markdown("**🏆 Top Achievement**")
            st.dataframe(top_ach.rename(columns={"Budget_kg": "Budget", "Actual_kg": "Actual"}),
                         use_container_width=True, hide_index=True, height=220)

        tab_all, tab_actual, tab_behind = st.tabs(["📋 ทั้งหมด", "✅ มี Actual", "🔴 ต้องเร่ง"])
        for tab, mask in [
            (tab_all,    slice(None)),
            (tab_actual, gap_df["Actual_kg"] > 0),
            (tab_behind, gap_df["Achievement%"] < 70),
        ]:
            with tab:
                show = gap_df[mask].sort_values("Achievement%", ascending=False).copy()
                show["Budget_kg"]    = show["Budget_kg"].apply(lambda v: f"{int(v):,}")
                show["Actual_kg"]    = show["Actual_kg"].apply(lambda v: f"{int(v):,}")
                show["Gap (kg)"]     = show["Gap (kg)"].apply(lambda v: f"{int(v):+,}")
                show["Achievement%"] = show["Achievement%"].apply(lambda v: f"{v:.1f}%")
                st.dataframe(show.rename(columns={"Budget_kg": "Budget (kg)", "Actual_kg": "Actual (kg)"}),
                             use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MENU 2 – CUSTOMER TABLE
# ═══════════════════════════════════════════════════════════════════════════════

elif menu == "🏢 ข้อมูลบริษัทลูกค้า":
    _scroll_top()
    st.title("🏢 ข้อมูลบริษัทลูกค้า")

    if df.empty or "Customer Name" not in df.columns:
        st.info("📂 กรุณาโหลดไฟล์จาก SharePoint ก่อน")
        st.stop()

    with st.expander("🔍 ตัวกรองข้อมูล", expanded=True):
        f1, f2, f3, f4 = st.columns(4)
        sel_reg = f1.selectbox("ภูมิภาค", ["ทั้งหมด"] + sorted(df["Region_TH"].dropna().astype(str).unique().tolist()))
        sel_ind = f2.selectbox("Industry", ["ทั้งหมด"] + sorted(df["Industry"].dropna().astype(str).unique().tolist()))
        sel_grd = f3.selectbox("Grade", ["ทั้งหมด"] + sorted(df["Grade"].dropna().astype(str).unique().tolist()))
        province_options = sorted([x for x in df.get("Province", pd.Series(dtype=str)).dropna().astype(str).unique().tolist() if str(x).strip()])
        sel_prov = f4.multiselect("Province", province_options)
        f5, f6 = st.columns([2, 2])
        sp_options = sorted([x for x in df["Salesperson"].dropna().astype(str).unique().tolist() if str(x).strip()])
        sel_sp_multi = f5.multiselect("Salesperson", sp_options)
        srch = f6.text_input("🔎 ค้นหาชื่อบริษัท / จังหวัด / Plus Code")

    flt = df.copy()
    if sel_reg != "ทั้งหมด": flt = flt[flt["Region_TH"] == sel_reg]
    if sel_ind != "ทั้งหมด": flt = flt[flt["Industry"] == sel_ind]
    if sel_grd != "ทั้งหมด": flt = flt[flt["Grade"] == sel_grd]
    if sel_prov: flt = flt[flt["Province"].astype(str).isin(sel_prov)]
    if sel_sp_multi: flt = flt[flt["Salesperson"].astype(str).isin(sel_sp_multi)]
    if srch:
        _mask = (
            flt["Customer Name"].astype(str).str.contains(srch, case=False, na=False)
            | flt.get("Province", pd.Series(index=flt.index, dtype=str)).astype(str).str.contains(srch, case=False, na=False)
            | flt.get("Plus_Code", pd.Series(index=flt.index, dtype=str)).astype(str).str.contains(srch, case=False, na=False)
        )
        flt = flt[_mask]

    GRADE_COLOR  = {"A": "#16a34a", "A-": "#22c55e", "B": "#2563eb", "B-": "#60a5fa",
                    "C": "#d97706", "C-": "#f59e0b", "F": "#dc2626"}
    REGION_BADGE = {"กลาง": "#e63946", "เหนือ": "#4c9be8", "ออก": "#2a9d8f",
                    "ตก": "#e76f51", "ใต้": "#8338ec", "ตะวันออกเฉียงเหนือ": "#f4a261",
                    "ไม่ระบุ": "#adb5bd"}

    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("📋 รายการที่พบ",           f"{len(flt):,} ราย")
    sc2.metric("💰 ยอดขายรวม (ที่กรอง)",   f"฿{flt['Sales/Year'].sum()/1e6:,.1f} M")
    sc3.metric("📊 เฉลี่ย/บริษัท",
               f"฿{(flt['Sales/Year'].sum()/len(flt)/1e6 if len(flt) else 0):,.2f} M")
    sc4.metric("📦 Budget รวม (kg/yr)",    f"{int(flt.get('Budget_kg', pd.Series(0)).sum()):,} kg")
    cex1, cex2 = st.columns([1,1])
    with cex1:
        st.download_button(
            "⬇️ Export Current Customer List (.csv)",
            data=flt.to_csv(index=False, encoding="utf-8-sig"),
            file_name=f"customer_list_{st.session_state.get('dept') or 'ALL'}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with cex2:
        if st.button("☁️ Upload Current Customer Export to SharePoint", use_container_width=True):
            remote_path = f"Reports/{st.session_state.get('dept') or 'ALL'}/customer_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            ok = sp_upload_bytes(flt.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"), remote_path, "text/csv")
            if ok:
                append_audit_log("upload_customer_export", remote_path, st.session_state.get("dept") or "")
                st.success("✅ ส่ง Customer Export ขึ้น SharePoint สำเร็จ")
    st.markdown("---")

    import urllib.parse

    def grade_badge(g):
        g = str(g).strip() if pd.notna(g) else ""
        c = GRADE_COLOR.get(g, "#6b7280")
        return f'<span style="background:{c};color:#fff;padding:2px 10px;border-radius:20px;font-size:12px;font-weight:700">{g or "—"}</span>'

    def region_badge(r):
        r = str(r).strip() if pd.notna(r) else ""
        c = REGION_BADGE.get(r, "#adb5bd")
        return f'<span style="background:{c}22;color:{c};border:1px solid {c};padding:2px 10px;border-radius:20px;font-size:12px;font-weight:600">{r or "—"}</span>'

    def fmt_sales(v):
        try:    return f'<span style="font-weight:600;color:#16a34a">฿{float(v):,.0f}</span>'
        except: return "—"

    def safe(v):
        s = str(v).strip() if pd.notna(v) else ""
        return s if s and s != "nan" else "—"

    LABEL_PAT = re.compile(
        r'(?:^|\n)\s*(Head\s*Office|Factory|Plant|Office|Branch|Warehouse|W/?H)\s*[:\)\s]*',
        re.IGNORECASE)
    LOC_ICONS = {"Head Office": "🏢", "Factory": "🏭", "Plant": "🏭",
                 "Office": "🏢", "Branch": "📌", "Warehouse": "📦", "W/H": "📦", "": "📍"}

    rows_html = ""
    for _, row in flt.iterrows():
        name      = safe(row.get("Customer Name"))
        sp        = safe(row.get("Salesperson"))
        ind       = safe(row.get("Industry"))
        grade     = grade_badge(row.get("Grade", ""))
        sales     = fmt_sales(row.get("Sales/Year", 0))
        prov      = safe(row.get("Province"))
        plus_code = safe(row.get("Plus_Code", ""))
        reg       = region_badge(row.get("Region_TH", ""))
        raw_address = str(row.get("Address", "") or "").strip()
        raw_province = "" if prov == "—" else prov
        raw_region = str(row.get("Region", "") or "").strip()
        row_ref_lat, row_ref_lng = resolve_reference_latlng(raw_province, raw_region, raw_address)
        row_coords = plus_code_to_coords(plus_code, ref_lat=row_ref_lat, ref_lng=row_ref_lng) if plus_code != "—" else None
        prefetched_js = f"[{float(row_coords[0]):.7f},{float(row_coords[1]):.7f}]" if row_coords else "null"

        raw_addr = str(row.get("Address", "")) if pd.notna(row.get("Address", "")) else ""
        locations = []
        if LABEL_PAT.search(raw_addr):
            parts = LABEL_PAT.split(raw_addr)
            i = 1
            while i < len(parts) - 1:
                lbl = parts[i].strip().title()
                txt = re.sub(
                    r'\n\s*(Tel|Fax|Email|E-mail|Website|www\.|TAX|Mobile|T\s*:|F\s*:|M\s*:|E\s*:)[^\n]*',
                    '', parts[i + 1].strip(), flags=re.IGNORECASE).strip()
                if txt: locations.append((lbl, txt))
                i += 2
        else:
            clean = re.sub(
                r'\n\s*(Tel|Fax|Email|E-mail|Website|www\.|TAX|Mobile|T\s*:|F\s*:|M\s*:|E\s*:)[^\n]*',
                '', raw_addr, flags=re.IGNORECASE).strip().split("\n")[0].strip()
            if clean: locations.append(("", clean))

        has_loc = bool(locations) or prov != "—" or (plus_code != "—" and len(plus_code) >= 6)

        if plus_code != "—" and len(plus_code) >= 6:
            map_query = plus_code
        elif prov != "—":
            map_query = f"{name} {prov} Thailand"
        else:
            map_query = f"{name} Thailand"

        loc_parts_html = []
        for lbl, txt in locations:
            icon = LOC_ICONS.get(lbl, "📌")
            lbl_html = (f"<span style='display:inline-block;background:#1e3a5f;color:#fff;"
                        f"font-size:10px;font-weight:700;padding:1px 8px;border-radius:10px;"
                        f"margin-bottom:3px'>{icon} {lbl}</span><br>") if lbl else ""
            if plus_code != "—" and len(plus_code) >= 6:
                q_loc = urllib.parse.quote(f"{plus_code} {txt} {prov if prov != '—' else ''} Thailand".strip())
            elif lbl and prov != "—":
                q_loc = urllib.parse.quote(f"{name} {lbl} {txt} {prov} Thailand")
            elif prov != "—":
                q_loc = urllib.parse.quote(f"{name} {txt} {prov} Thailand")
            else:
                q_loc = urllib.parse.quote(f"{name} {txt} Thailand")

            n_js  = name.replace("'", "`").replace('"', "`")
            btn_label = f"🗺️ {lbl}" if lbl else "🗺️ ดูแผนที่"
            btn = (f"<button onclick=\"event.stopPropagation();"
                   f"showMap('{q_loc}','{n_js}{(' - '+lbl) if lbl else ''}',event,true,{prefetched_js})\" "
                   f"style='margin-top:4px;font-size:10px;background:#2563eb;color:#fff;"
                   f"border:none;border-radius:6px;padding:2px 9px;cursor:pointer'>{btn_label}</button>")
            loc_parts_html.append(
                f"<div style='margin-bottom:6px;padding:6px 8px;background:#f8fafc;"
                f"border-left:3px solid #2563eb;border-radius:0 6px 6px 0'>"
                f"{lbl_html}<div style='font-size:11.5px;color:#1e293b;line-height:1.6'>{txt}</div>"
                f"{btn}</div>")

        pc_badge = (f"<div style='margin-bottom:5px'>"
                    f"<span style='background:#dcfce7;color:#166534;font-size:10.5px;"
                    f"font-weight:700;padding:2px 8px;border-radius:6px'>📌 {plus_code}</span>"
                    f"</div>") if (plus_code != "—" and len(plus_code) >= 6) else ""
        location_html = pc_badge + ("".join(loc_parts_html) if loc_parts_html
                                    else "<span style='color:#adb5bd'>—</span>")

        if has_loc:
            q_enc   = urllib.parse.quote(map_query)
            name_js = name.replace("'", "`")
            tr_attr = f"onclick=\"showMap('{q_enc}','{name_js}',event,true,{prefetched_js})\" class='clickable'"
            co_html = f"<div class='co has-map'>📍 {name}</div>"
        else:
            tr_attr = "class='no-map'"
            co_html = f"<div class='co no-loc'>{name}</div>"

        bkg_int = int(row.get("Budget_kg", 0) or 0)
        act_int = int(row.get("Actual_kg",  0) or 0)
        bkg_html = (f"<span style='font-size:10.5px;background:#fef3c7;color:#92400e;"
                    f"border-radius:6px;padding:1px 7px;font-weight:600'>🎯 {bkg_int:,} kg</span>"
                    if bkg_int > 0 else "")
        if act_int > 0 and bkg_int > 0:
            pct_a = act_int / bkg_int * 100
            ac = "#16a34a" if pct_a >= 100 else "#d97706" if pct_a >= 50 else "#dc2626"
            act_html = (f"<span style='font-size:10.5px;background:{ac}22;color:{ac};"
                        f"border-radius:6px;padding:1px 7px;font-weight:600'>"
                        f"✅ {act_int:,} kg ({pct_a:.0f}%)</span>")
        elif act_int > 0:
            act_html = (f"<span style='font-size:10.5px;background:#dcfce7;color:#16a34a;"
                        f"border-radius:6px;padding:1px 7px;font-weight:600'>✅ {act_int:,} kg</span>")
        else:
            act_html = ""

        rows_html += (
            f"<tr {tr_attr}>"
            f"<td>{co_html}<div class='sp'>👤 {sp}</div></td>"
            f"<td class='ind'>{ind}</td>"
            f"<td style='text-align:center'>{grade}</td>"
            f"<td class='sal'>{sales}<br>{bkg_html} {act_html}</td>"
            f"<td class='loc'>{location_html}</td>"
            f"<td style='text-align:center'>{reg}</td>"
            f"</tr>"
        )

    ORIGIN_PLUS_SHORT = "MJHG+2F"
    ORIGIN_LABEL      = "บริษัท ออฟติมอลเทค จำกัด"
    ORIGIN_LAT_FIXED  = 13.6776
    ORIGIN_LNG_FIXED  = 100.6262

    import urllib.parse as _up
    gmaps_origin = _up.quote(f"{ORIGIN_PLUS_SHORT} Bangkok Thailand")

    map_points_json, map_points_no_coords_json = build_map_points(
        flt, ref_lat=ORIGIN_LAT_FIXED, ref_lng=ORIGIN_LNG_FIXED
    )


    st.subheader("🚀 Sales Planning & AI Recommendations")
    ai_df = build_ai_recommendations(flt, ORIGIN_LAT_FIXED, ORIGIN_LNG_FIXED)
    ai_top = ai_df.head(20).copy()
    plan1, plan2, plan3 = st.columns([1.1, 1.1, 1.8])
    salesperson_filter = plan1.selectbox("เลือก Salesperson สำหรับแผนเข้า", ["ทั้งหมด"] + sorted([x for x in ai_df.get("Salesperson", pd.Series(dtype=str)).dropna().astype(str).unique().tolist() if str(x).strip()]))
    province_filter = plan2.selectbox("เลือกจังหวัดเป้าหมาย", ["ทั้งหมด"] + sorted([x for x in ai_df.get("Province", pd.Series(dtype=str)).dropna().astype(str).unique().tolist() if str(x).strip()]))
    visit_day = plan3.date_input("วางแผนเข้าพบวันที่", value=datetime.now().date())
    ai_view = ai_df.copy()
    if salesperson_filter != "ทั้งหมด":
        ai_view = ai_view[ai_view["Salesperson"].astype(str) == salesperson_filter]
    if province_filter != "ทั้งหมด":
        ai_view = ai_view[ai_view["Province"].astype(str) == province_filter]
    suggested_names = ai_view.head(8)["Customer Name"].astype(str).tolist()
    selected_targets = st.multiselect("ลูกค้าที่แนะนำให้เข้าในทริปนี้", suggested_names, default=suggested_names[:5])
    plan_df = ai_view[ai_view["Customer Name"].astype(str).isin(selected_targets)].copy()
    if not plan_df.empty:
        plan_df = plan_df.sort_values(["Province", "Distance_km", "AI_Score"], ascending=[True, True, False])
        stops = [(float(r["Latitude"]), float(r["Longitude"])) for _, r in plan_df.iterrows() if pd.notna(r.get("Latitude")) and pd.notna(r.get("Longitude"))]
        route_link = build_google_maps_route_link(f"{ORIGIN_PLUS_SHORT} Bangkok Thailand", stops)
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("ลูกค้าในแผน", f"{len(plan_df):,} ราย")
        kpi2.metric("ยอดขายรวมเป้าหมาย", f"฿{plan_df['Sales/Year'].sum()/1e6:,.1f}M")
        kpi3.metric("Gap รวม", f"{int(plan_df['Gap_kg'].sum()):,} kg")
        kpi4.metric("คะแนน AI เฉลี่ย", f"{plan_df['AI_Score'].mean():.1f}")
        st.markdown(f"**แผนเข้าเยี่ยมวันที่ {visit_day.strftime('%d/%m/%Y')}** — [เปิดเส้นทางใน Google Maps]({route_link})")
        show_plan = plan_df[[c for c in ["Customer Name","Salesperson","Province","Grade","AI_Score","AI_Reason","Distance_km","Gap_kg","Achievement_pct"] if c in plan_df.columns]].copy()
        if "Distance_km" in show_plan.columns:
            show_plan["Distance_km"] = show_plan["Distance_km"].apply(lambda v: f"{v:,.1f} km" if pd.notna(v) and v < 9999 else "—")
        if "Gap_kg" in show_plan.columns:
            show_plan["Gap_kg"] = show_plan["Gap_kg"].apply(lambda v: f"{int(v):,}")
        if "Achievement_pct" in show_plan.columns:
            show_plan["Achievement_pct"] = show_plan["Achievement_pct"].apply(lambda v: f"{v:.1f}%")
        st.dataframe(show_plan, use_container_width=True, hide_index=True)
    with st.expander("ดู Top AI Recommendations", expanded=False):
        ai_show = ai_top[[c for c in ["Customer Name","Salesperson","Province","Grade","AI_Score","AI_Reason","Gap_kg","Achievement_pct","Sales/Year"] if c in ai_top.columns]].copy()
        if "Gap_kg" in ai_show.columns:
            ai_show["Gap_kg"] = ai_show["Gap_kg"].apply(lambda v: f"{int(v):,}")
        if "Achievement_pct" in ai_show.columns:
            ai_show["Achievement_pct"] = ai_show["Achievement_pct"].apply(lambda v: f"{v:.1f}%")
        if "Sales/Year" in ai_show.columns:
            ai_show["Sales/Year"] = ai_show["Sales/Year"].apply(lambda v: f"฿{float(v):,.0f}")
        st.dataframe(ai_show, use_container_width=True, hide_index=True)

    map_ui1, map_ui2, map_ui3 = st.columns([1.2, 1.2, 2])
    default_view = map_ui1.selectbox("🗺️ Map View", ["Cluster", "Heatmap", "Hybrid"], index=2)
    auto_fit_bounds = map_ui2.toggle("Auto fit bounds", value=True)
    map_ui3.caption("โหมด Hybrid = แสดงทั้ง cluster และ heatmap พร้อมกัน")

    html_table = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
html,body{{font-family:'Sarabun',sans-serif;background:transparent;}}
.page{{display:flex;flex-direction:column;gap:10px;padding:4px;}}
.route-bar{{background:linear-gradient(135deg,#1e3a5f,#2563eb);border-radius:12px;
  padding:12px 16px;color:#fff;display:flex;flex-direction:column;gap:6px;}}
.route-title{{font-size:13px;font-weight:700;display:flex;align-items:center;gap:6px;}}
.route-steps{{display:flex;align-items:center;gap:8px;font-size:12px;flex-wrap:wrap;}}
.route-box{{background:rgba(255,255,255,.15);border-radius:8px;padding:4px 10px;
  font-size:12px;white-space:nowrap;}}
.route-arrow{{font-size:16px;opacity:.8;}}
.route-hint{{font-size:10.5px;opacity:.65;margin-top:2px;}}
.btn-gmaps{{display:inline-flex;align-items:center;gap:5px;margin-left:auto;
  background:#fff;color:#2563eb;border:none;border-radius:8px;
  padding:5px 12px;font-size:11.5px;font-weight:700;cursor:pointer;
  text-decoration:none;white-space:nowrap;}}
.btn-gmaps:hover{{background:#dbeafe;}}
.map-wrap{{border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;
  box-shadow:0 2px 12px rgba(0,0,0,0.08);}}
#leaflet-map{{width:100%;height:340px;}}
.wrap{{max-height:380px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:12px;
  box-shadow:0 2px 8px rgba(0,0,0,0.06);}}
table{{width:100%;border-collapse:collapse;}}
thead tr{{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:#fff;
  position:sticky;top:0;z-index:10;}}
thead th{{padding:11px 13px;text-align:left;font-size:12px;font-weight:600;white-space:nowrap;}}
tbody tr{{transition:background .12s;}}
tbody tr:nth-child(even){{background:#f8fafc;}}
tbody tr.clickable{{cursor:pointer;}}
tbody tr.clickable:hover{{background:#dbeafe;}}
tbody tr.clickable.active{{background:#bfdbfe!important;box-shadow:inset 3px 0 0 #2563eb;}}
tbody tr.no-map{{cursor:default;opacity:.55;}}
td{{padding:9px 13px;border-bottom:1px solid #f0f4f8;vertical-align:middle;}}
.co{{font-weight:600;font-size:13px;}}
.has-map{{color:#2563eb;}}
.no-loc{{color:#94a3b8;}}
.sp{{color:#64748b;font-size:11px;margin-top:2px;}}
.ind{{color:#374151;font-size:12px;}}
.sal{{font-weight:600;font-size:12px;text-align:right;white-space:nowrap;}}
.loc{{color:#374151;font-size:11.5px;word-break:break-word;min-width:160px;line-height:1.5;}}
.legend{{font-size:11px;color:#64748b;padding:5px 13px 7px;display:flex;gap:14px;background:#f8fafc;}}
</style>
</head>
<body><div class="page">

<div class="route-bar">
  <div class="route-title">🚗 เส้นทางการเดินทาง</div>
  <div class="route-steps">
    <span class="route-box">🏠 {ORIGIN_LABEL}</span>
    <span class="route-arrow">→</span>
    <span class="route-box" id="dest-label">กรุณาเลือกบริษัทปลายทาง</span>
    <a id="open-gmaps" href="#" target="_blank" class="btn-gmaps" style="display:none">
      🗺️ เปิด Google Maps
    </a>
  </div>
  <div class="route-hint" id="route-hint">⏳ กำลังโหลด Open Location Code library…</div>
</div>

<div class="map-wrap"><div id="leaflet-map"></div></div>\n<div id="marker-hint" style="font-size:11px;color:#64748b;padding:6px 4px 0 4px;"></div>

<div class="wrap">
  <div class="legend">
    <span>📍 <b style="color:#2563eb">สีน้ำเงิน</b> = คลิกเพื่อดูเส้นทาง</span>
    <span style="color:#94a3b8">⬜ สีเทา = ไม่มีข้อมูลที่อยู่</span>
  </div>
  <table>
    <thead><tr>
      <th>🏢 บริษัท / Salesperson</th><th>🏭 Industry</th>
      <th style="text-align:center">⭐ Grade</th>
      <th style="text-align:right">💰 Sales/Year</th>
      <th>📍 ที่ตั้ง</th><th style="text-align:center">🗺️ ภูมิภาค</th>
    </tr></thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>

</div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css"/>
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css"/>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
<script src="https://unpkg.com/leaflet.heat/dist/leaflet-heat.js"></script>

<script>
const ORIGIN_LAT   = {ORIGIN_LAT_FIXED};
const ORIGIN_LNG   = {ORIGIN_LNG_FIXED};
const ORIGIN_LABEL = "{ORIGIN_LABEL}";
const ORIGIN_SHORT = "{ORIGIN_PLUS_SHORT}";
const MAP_POINTS          = {map_points_json};
const MAP_POINTS_NO_COORD = {map_points_no_coords_json};
const DEFAULT_VIEW_MODE   = "{default_view}";
const AUTO_FIT_BOUNDS     = {str(auto_fit_bounds).lower()};

let destMarker = null, routeLayer = null, heatLayer = null;

var clusterGroup = L.markerClusterGroup({{
    spiderfyOnMaxZoom: true,
    showCoverageOnHover: false,
    maxClusterRadius: 45
}});

var map = L.map('leaflet-map').setView([ORIGIN_LAT, ORIGIN_LNG], 10);
L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
    attribution: '© OpenStreetMap contributors', maxZoom: 19
}}).addTo(map);
clusterGroup.addTo(map);

var originIcon = L.divIcon({{
    html: '<div style="background:#1e3a5f;color:#fff;border-radius:50%;width:40px;height:40px;'
        + 'display:flex;align-items:center;justify-content:center;font-size:22px;'
        + 'border:3px solid #fff;box-shadow:0 3px 12px rgba(0,0,0,.4)">🏠</div>',
    iconSize:[40,40], iconAnchor:[20,20]
}});
L.marker([ORIGIN_LAT, ORIGIN_LNG], {{icon: originIcon}})
    .addTo(map)
    .bindPopup('<b>' + ORIGIN_LABEL + '</b><br><small>' + ORIGIN_SHORT + ' — Bangna/Phra Khanong</small>')
    .openPopup();

document.getElementById('route-hint').textContent =
    '✅ แผนที่พร้อม — คลิกชื่อบริษัทในตารางเพื่อดูเส้นทาง';

function escapeHtml(s) {{
    return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;')
        .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}}

(function renderMarkers() {{
    clusterGroup.clearLayers();
    if (!MAP_POINTS || MAP_POINTS.length === 0) {{
        document.getElementById('marker-hint').textContent =
            'ℹ️ ไม่มีข้อมูล Plus Code ในชุดนี้ — คลิกชื่อบริษัทเพื่อค้นหาตำแหน่ง';
        return;
    }}
    var bounds = [];
    var heatData = [];
    MAP_POINTS.forEach(function(item) {{
        var coords = [Number(item.lat), Number(item.lng)];
        var popup  = '<b>' + escapeHtml(item.name) + '</b>'
            + '<br><small>👤 ' + escapeHtml(item.salesperson||'—') + '</small>'
            + '<br><small>📍 ' + escapeHtml(item.province||'—') + '</small>'
            + (item.plus_code ? '<br><small>📌 ' + escapeHtml(item.plus_code) + '</small>' : '');
        var mk = L.marker(coords);
        heatData.push([coords[0], coords[1], 0.7]);
        mk.bindPopup(popup);
        mk.on('click', function() {{
            showMap(
                encodeURIComponent(item.query||''),
                item.name||'',
                null, false,
                [Number(item.lat), Number(item.lng)]
            );
        }});
        clusterGroup.addLayer(mk);
        bounds.push(coords);
    }});
    if (heatLayer) {{ map.removeLayer(heatLayer); heatLayer = null; }}
    if (heatData.length) {{ heatLayer = L.heatLayer(heatData, {{radius: 28, blur: 20, maxZoom: 13}}); }}
    if (heatLayer && (DEFAULT_VIEW_MODE === 'Heatmap' || DEFAULT_VIEW_MODE === 'Hybrid')) {{ heatLayer.addTo(map); }}
    if (DEFAULT_VIEW_MODE === 'Heatmap' && clusterGroup) {{ try {{ map.removeLayer(clusterGroup); }} catch(e) {{}} }}
    var skipped = MAP_POINTS_NO_COORD.length;
    document.getElementById('marker-hint').textContent =
        '📌 ' + MAP_POINTS.length + ' หมุด'
        + (skipped > 0 ? ' • อีก ' + skipped + ' รายการไม่มี Plus Code (geocode เมื่อคลิก)' : '');
    if (AUTO_FIT_BOUNDS) {{
        if (bounds.length > 1) map.fitBounds(bounds, {{padding:[40,40]}});
        else if (bounds.length === 1) map.setView(bounds[0], 11);
    }}
}})();

var OLC_A = '23456789CFGHJMPQRVWX';
var OLC_R = [20.0, 1.0, 0.05, 0.0025, 0.000125];
function olcIdx(c) {{ return OLC_A.indexOf(c.toUpperCase()); }}
function olcDecodeFull(code) {{
    var d = code.toUpperCase().replace('+','').replace(/0+$/,'');
    if (d.length < 4) return null;
    var lat = 0, lng = 0;
    for (var i = 0; i < Math.min(d.length, 10); i += 2) {{
        var res = OLC_R[i >> 1];
        var li = olcIdx(d[i]), ni = (i+1 < d.length) ? olcIdx(d[i+1]) : 0;
        if (li < 0 || ni < 0) return null;
        lat += li * res; lng += ni * res;
    }}
    var finest = OLC_R[Math.min(Math.floor((d.length-1)/2), 4)];
    return [lat - 90 + finest/2, lng - 180 + finest/2];
}}
function olcPrefix4(lat, lng) {{
    var la=lat+90, lo=lng+180;
    var p1l=Math.floor(la/20), r1l=la-p1l*20;
    var p1g=Math.floor(lo/20), r1g=lo-p1g*20;
    return OLC_A[Math.floor(la/20)]+OLC_A[Math.floor(lo/20)]
          +OLC_A[Math.floor(r1l)]+OLC_A[Math.floor(r1g)];
}}
function olcRecover(shortCode, refLat, refLng) {{
    var s = String(shortCode||'').trim().split(/[ \t]+/)[0].toUpperCase();
    var pi = s.indexOf('+'); if (pi < 0) return null;
    var before = s.substring(0, pi), tail = s.substring(pi+1);
    if (before.length >= 8) return olcDecodeFull(s);
    var pfLen = 8 - before.length;
    var best = olcDecodeFull(olcPrefix4(refLat, refLng).substring(0,pfLen)+before+'+'+tail);
    var gs = OLC_R[pfLen/2 - 1], bestD = 1e9;
    for (var dl of [-1,0,1]) for (var dg of [-1,0,1]) {{
        var tp = olcPrefix4(refLat+dl*gs, refLng+dg*gs).substring(0,pfLen);
        var td = olcDecodeFull(tp+before+'+'+tail);
        if (!td) continue;
        var dist = Math.pow(td[0]-refLat,2)+Math.pow(td[1]-refLng,2);
        if (dist < bestD) {{ bestD=dist; best=td; }}
    }}
    return best;
}}
function plusCodeToCoords(code) {{
    try {{
        var s = String(code||'').trim(); if (!s || s.indexOf('+')<0) return null;
        var before = s.substring(0, s.indexOf('+'));
        return before.length >= 8 ? olcDecodeFull(s) : olcRecover(s, ORIGIN_LAT, ORIGIN_LNG);
    }} catch(e) {{ return null; }}
}}

async function geocode(query) {{
    var q = decodeURIComponent(String(query||'')).trim(); if (!q) return null;
    var pm = q.match(/([23456789CFGHJMPQRVWX]{{4,8}}[+][23456789CFGHJMPQRVWX]{{2,3}})/i);
    if (pm) {{ var c = plusCodeToCoords(pm[1]); if (c) return c; }}
    for (var attempt = 0; attempt < 2; attempt++) {{
        try {{
            var qe = encodeURIComponent(attempt===0 ? q : q+' Thailand');
            var r  = await fetch(
                'https://nominatim.openstreetmap.org/search?q='+qe+
                '&format=jsonv2&limit=1&countrycodes=th',
                {{headers:{{'Accept':'application/json','User-Agent':'STDashboard/1'}}}}
            );
            var d = await r.json();
            if (d.length > 0) return [parseFloat(d[0].lat), parseFloat(d[0].lon)];
        }} catch(e) {{ console.warn('Nominatim:', e.message); }}
    }}
    return null;
}}

async function drawRoute(dLat, dLng, destName) {{
    var url = 'https://router.project-osrm.org/route/v1/driving/'
        + ORIGIN_LNG+','+ORIGIN_LAT+';'+dLng+','+dLat
        + '?overview=full&geometries=geojson';
    try {{
        var ctrl = new AbortController();
        var timer = setTimeout(()=>ctrl.abort(), 10000);
        var r = await fetch(url, {{signal:ctrl.signal}});
        clearTimeout(timer);
        var d = await r.json();
        if (d.routes && d.routes.length > 0) {{
            var route = d.routes[0];
            var dist  = (route.distance/1000).toFixed(1);
            var mins  = Math.round(route.duration/60);
            var h = Math.floor(mins/60), m = mins%60;
            var ts = h>0 ? h+' ชม. '+m+' นาที' : mins+' นาที';
            if (routeLayer) map.removeLayer(routeLayer);
            routeLayer = L.geoJSON(route.geometry,
                {{style:{{color:'#2563eb',weight:5,opacity:.9}}}}).addTo(map);
            map.fitBounds(routeLayer.getBounds(), {{padding:[50,50]}});
            document.getElementById('route-hint').textContent =
                '🚗 '+dist+' กม. | ⏱ '+ts+'  ('+ORIGIN_LABEL+' → '+destName+')';
            return;
        }}
    }} catch(e) {{ console.warn('OSRM fail:', e.message||e); }}
    if (routeLayer) {{ map.removeLayer(routeLayer); routeLayer = null; }}
    routeLayer = L.polyline([[ORIGIN_LAT,ORIGIN_LNG],[dLat,dLng]],
        {{color:'#94a3b8',weight:3,dashArray:'8 5',opacity:.8}}).addTo(map);
    map.fitBounds(routeLayer.getBounds(), {{padding:[60,60]}});
    document.getElementById('route-hint').textContent =
        '⚠️ คำนวณเส้นทางอัตโนมัติไม่ได้ — กด "เปิด Google Maps"';
}}

async function showMap(destQuery, destName, e, drawRouteLine, prefetchedCoords) {{
    if (drawRouteLine === undefined) drawRouteLine = true;
    if (prefetchedCoords === undefined) prefetchedCoords = null;
    var name = decodeURIComponent(String(destName||''));
    document.getElementById('dest-label').textContent = '📍 '+name;
    document.getElementById('route-hint').textContent  = '🔍 กำลังค้นหาตำแหน่ง…';
    var rawDest = decodeURIComponent(String(destQuery||''));
    var btn = document.getElementById('open-gmaps');
    btn.href = 'https://www.google.com/maps/dir/?api=1'
        +'&origin='+encodeURIComponent(ORIGIN_SHORT+' Bangkok Thailand')
        +'&destination='+encodeURIComponent(rawDest)
        +'&travelmode=driving';
    btn.style.display = 'inline-flex';
    document.querySelectorAll('tbody tr').forEach(function(r){{r.classList.remove('active');}});
    var tr = e && (e.currentTarget || (e.target && e.target.closest('tr')));
    if (tr) tr.classList.add('active');
    if (destMarker) {{ map.removeLayer(destMarker); destMarker = null; }}
    if (routeLayer) {{ map.removeLayer(routeLayer); routeLayer = null; }}
    var coords = prefetchedCoords || (await geocode(rawDest));
    if (!coords) {{
        document.getElementById('route-hint').textContent = '❌ ไม่พบตำแหน่ง — กด "เปิด Google Maps"';
        return;
    }}
    var destIcon = L.divIcon({{
        html:'<div style="background:#dc2626;color:#fff;border-radius:50%;width:40px;height:40px;'
            +'display:flex;align-items:center;justify-content:center;font-size:22px;'
            +'border:3px solid #fff;box-shadow:0 3px 12px rgba(0,0,0,.4)">📍</div>',
        iconSize:[40,40], iconAnchor:[20,20]
    }});
    destMarker = L.marker(coords, {{icon:destIcon}})
        .addTo(map).bindPopup('<b>'+name+'</b>').openPopup();
    if (!drawRouteLine) {{
        map.setView(coords, 11);
        document.getElementById('route-hint').textContent =
            '📍 คลิกแถวในตารางเพื่อดูเส้นทาง';
        return;
    }}
    document.getElementById('route-hint').textContent = '⏳ กำลังคำนวณเส้นทาง…';
    await drawRoute(coords[0], coords[1], name);
}}
</script>
</body></html>"""

    components.html(html_table, height=800, scrolling=False)

    st.markdown("<br>", unsafe_allow_html=True)
    VIEW = ["Customer Name", "Salesperson", "Industry", "Grade", "Sales/Year",
            "Budget_kg", "Actual_kg", "Plus_Code", "Sub-district", "District", "Province", "Region_TH"]
    export_df = flt[[c for c in VIEW if c in flt.columns]].rename(columns={"Region_TH": "Region"})
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ CSV", data=export_df.to_csv(index=False, encoding="utf-8-sig"),
                       file_name="customers.csv", mime="text/csv")
    c2.download_button("⬇️ Excel", data=to_excel_bytes(export_df),
                       file_name="customers.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


elif menu == "🧠 Executive Report":
    _scroll_top()
    st.title("🧠 Executive Report Center")

    if df.empty or "Customer Name" not in df.columns:
        st.info("📂 กรุณาโหลดไฟล์จาก SharePoint ก่อน")
        st.stop()

    rep = build_executive_report_df(df)
    st.caption("สรุปผู้บริหารสำหรับการประชุม, export, และส่งรายงานต่อ")

    k1, k2, k3, k4 = st.columns(4)
    total_sales = float(pd.to_numeric(rep["Sales/Year"], errors="coerce").fillna(0).sum())
    total_budget = float(pd.to_numeric(rep["Budget_kg"], errors="coerce").fillna(0).sum())
    total_actual = float(pd.to_numeric(rep["Actual_kg"], errors="coerce").fillna(0).sum())
    avg_ach = float(pd.to_numeric(rep["achievement_pct"], errors="coerce").fillna(0).mean()) if len(rep) else 0
    k1.metric("💰 Total Sales", f"฿{total_sales/1e6:,.1f} M")
    k2.metric("🎯 Budget รวม", f"{int(total_budget):,} kg")
    k3.metric("✅ Actual รวม", f"{int(total_actual):,} kg")
    k4.metric("📈 Avg Achievement", f"{avg_ach:,.1f}%")

    c1, c2 = st.columns([1.1, 1])
    with c1:
        top_rep = rep.head(15).copy()
        st.markdown("**🎯 Top Opportunity Accounts**")
        show = top_rep[["Customer Name", "Salesperson", "Industry", "Province", "gap_kg", "achievement_pct", "opportunity_score"]].copy()
        show["gap_kg"] = show["gap_kg"].apply(lambda v: f"{int(v):,} kg")
        show["achievement_pct"] = show["achievement_pct"].apply(lambda v: f"{v:.1f}%")
        show["opportunity_score"] = show["opportunity_score"].apply(lambda v: f"{v:.1f}")
        st.dataframe(show.rename(columns={"gap_kg": "Gap", "achievement_pct": "Achievement", "opportunity_score": "Score"}), use_container_width=True, hide_index=True, height=430)

    with c2:
        by_sp = rep.groupby("Salesperson", dropna=False).agg(
            total_sales=("Sales/Year", "sum"),
            total_gap=("gap_kg", "sum"),
            avg_achievement=("achievement_pct", "mean"),
            customers=("Customer Name", "count")
        ).reset_index().sort_values("total_sales", ascending=False)
        st.markdown("**👤 Salesperson Performance**")
        fig_sp = px.bar(
            by_sp.head(12),
            x="Salesperson",
            y="total_sales",
            text=by_sp.head(12)["total_sales"].apply(lambda v: f"฿{v/1e6:.1f}M"),
            color="avg_achievement",
            color_continuous_scale="Blues",
            labels={"total_sales": "Total Sales", "avg_achievement": "Avg Achievement %"}
        )
        fig_sp.update_traces(textposition="outside")
        fig_sp.update_layout(height=430, coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_sp, use_container_width=True)

    r1, r2, r3 = st.columns(3)
    region_summary = rep.groupby("Region_TH", dropna=False).agg(
        customers=("Customer Name", "count"),
        sales=("Sales/Year", "sum"),
        budget=("Budget_kg", "sum"),
        actual=("Actual_kg", "sum")
    ).reset_index().sort_values("sales", ascending=False)
    industry_summary = rep.groupby("Industry", dropna=False).agg(
        customers=("Customer Name", "count"),
        sales=("Sales/Year", "sum")
    ).reset_index().sort_values("sales", ascending=False)
    province_summary = rep.groupby("Province", dropna=False).agg(
        customers=("Customer Name", "count"),
        sales=("Sales/Year", "sum")
    ).reset_index().sort_values("sales", ascending=False)

    with r1:
        st.markdown("**🗺️ Region Summary**")
        st.dataframe(region_summary, use_container_width=True, hide_index=True, height=280)
    with r2:
        st.markdown("**🏭 Industry Summary**")
        st.dataframe(industry_summary.head(15), use_container_width=True, hide_index=True, height=280)
    with r3:
        st.markdown("**📍 Province Summary**")
        st.dataframe(province_summary.head(15), use_container_width=True, hide_index=True, height=280)

    st.divider()
    st.subheader("📦 Export + Report")
    report_xlsx = to_excel_bytes_multi({
        "Executive Report": rep,
        "Region Summary": region_summary,
        "Industry Summary": industry_summary,
        "Province Summary": province_summary,
    })
    cexp1, cexp2, cexp3 = st.columns(3)
    with cexp1:
        st.download_button(
            "⬇️ Download Executive Excel",
            data=report_xlsx,
            file_name=f"executive_report_{st.session_state.get('dept') or 'ALL'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with cexp2:
        st.download_button(
            "⬇️ Download Executive CSV",
            data=rep.to_csv(index=False, encoding="utf-8-sig"),
            file_name=f"executive_report_{st.session_state.get('dept') or 'ALL'}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with cexp3:
        if st.button("☁️ Upload Executive Report to SharePoint", use_container_width=True):
            remote_path = f"Reports/{st.session_state.get('dept') or 'ALL'}/executive_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ok = sp_upload_bytes(report_xlsx, remote_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if ok:
                append_audit_log("upload_executive_report", remote_path, st.session_state.get("dept") or "")
                st.success("✅ ส่ง Executive Report ขึ้น SharePoint สำเร็จ")

    st.divider()
    st.subheader("🧭 Smart Customer Map Export")
    map_export = rep[["Customer Name", "Salesperson", "Province", "Region_TH", "Plus_Code", "Sales/Year", "opportunity_score"]].copy()
    st.download_button(
        "⬇️ Download Map Customer List",
        data=map_export.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"customer_map_export_{st.session_state.get('dept') or 'ALL'}.csv",
        mime="text/csv",
        use_container_width=False,
    )

# ═══════════════════════════════════════════════════════════════════════════════
# MENU 3 – EDIT / ADD
# ═══════════════════════════════════════════════════════════════════════════════

else:
    _scroll_top()
    st.title("✏️ แก้ไข / เพิ่มข้อมูลลูกค้า")

    if df.empty or "Customer Name" not in df.columns:
        st.info("📂 กรุณาโหลดไฟล์จาก SharePoint ก่อน")
        st.stop()

    tab_edit, tab_add = st.tabs(["📝 แก้ไขข้อมูล", "➕ เพิ่มลูกค้าใหม่"])
    GRADE_COLOR = {"A": "#16a34a", "A-": "#22c55e", "B": "#2563eb", "B-": "#60a5fa",
                   "C": "#d97706", "C-": "#f59e0b", "F": "#dc2626"}

    def _s(v):
        x = str(v).strip() if pd.notna(v) else ""
        return x if x and x != "nan" else ""

    def _commit_save(label: str = "บันทึก"):
        if st.session_state.sp_file and st.session_state.dept:
            with st.spinner("💾 กำลังบันทึกขึ้น SharePoint…"):
                ok = sp_save(st.session_state.df,
                             st.session_state.dept,
                             st.session_state.sp_file)
            if ok:
                append_audit_log("save_sharepoint", label, st.session_state.dept)
                st.success(f"✅ {label} สำเร็จ! (บันทึกขึ้น SharePoint แล้ว)")
            else:
                st.warning(f"⚠️ {label} ใน session แล้ว แต่ upload SharePoint ไม่สำเร็จ — ลอง Export แทน")
        else:
            append_audit_log("save_session", label, st.session_state.get("dept") or "")
            st.success(f"✅ {label} สำเร็จ! (บันทึกใน session — กรุณา Export เพื่อเก็บไฟล์)")

    with tab_edit:
        sc1, sc2, sc3 = st.columns([4, 1.2, 1.2])
        srch2 = sc1.text_input("", key="edit_srch", placeholder="🔎 ค้นหาชื่อบริษัท…",
                               label_visibility="collapsed")
        if sc2.button("✏️ แก้ไข",
                      type="primary" if st.session_state.edit_mode == "edit" else "secondary",
                      use_container_width=True):
            st.session_state.edit_mode = "edit"; st.session_state.confirm_delete = False; st.rerun()
        if sc3.button("🗑️ ลบ",
                      type="primary" if st.session_state.edit_mode == "delete" else "secondary",
                      use_container_width=True):
            st.session_state.edit_mode = "delete"; st.session_state.confirm_delete = False; st.rerun()

        mask   = (df["Customer Name"].str.contains(srch2, case=False, na=False).values
                  if srch2 else [True] * len(df))
        subset = df[mask].copy()
        subset["_orig_idx"] = subset.index
        subset = subset.reset_index(drop=True)
        orig_idx = subset["_orig_idx"].tolist()

        if subset.empty:
            st.info("ไม่พบข้อมูล")
        else:
            st.caption(f"พบ **{len(subset):,}** รายการ  •  "
                       f"{'✏️ โหมดแก้ไข' if st.session_state.edit_mode=='edit' else '🗑️ โหมดลบ'}")
            st.divider()

            if st.session_state.edit_mode == "delete":
                if ("del_checks" not in st.session_state or
                        len(st.session_state.del_checks) != len(subset)):
                    st.session_state.del_checks = [False] * len(subset)
                sa1, sa2, _ = st.columns([1.5, 1.8, 6])
                if sa1.button("☑️ เลือกทั้งหมด", use_container_width=True):
                    st.session_state.del_checks = [True] * len(subset); st.rerun()
                if sa2.button("⬜ ยกเลิกทั้งหมด", use_container_width=True):
                    st.session_state.del_checks = [False] * len(subset); st.rerun()

            if "editing_idx" not in st.session_state:
                st.session_state.editing_idx = None

            for i, row in subset.iterrows():
                orig_i    = orig_idx[i]
                name      = _s(row.get("Customer Name"))
                sp        = _s(row.get("Salesperson"))
                ind       = _s(row.get("Industry"))
                grade     = _s(row.get("Grade"))
                sales_v   = row.get("Sales/Year", 0)
                try:    sales_fmt = f"฿{int(round(float(sales_v))):,}"
                except: sales_fmt = "—"
                raw_addr  = _s(row.get("Address", ""))
                addr_line = raw_addr.split("\n")[0].strip()
                prov      = _s(row.get("Province"))
                g_color   = GRADE_COLOR.get(grade, "#6b7280")
                loc_txt   = addr_line if addr_line else (prov if prov else "—")

                card_tpl = (
                    '<div style="border:{border};background:{bg};border-radius:12px;'
                    'padding:12px 16px;margin-bottom:8px;box-shadow:0 1px 4px rgba(0,0,0,.06)">'
                    '<div style="display:flex;justify-content:space-between;align-items:center">'
                    '<span style="font-weight:700;font-size:14px;color:#1e293b">{name}</span>'
                    '<span style="font-weight:600;color:#16a34a;font-size:13px">{sales}</span></div>'
                    '<div style="margin-top:4px">'
                    '<span style="background:{gc};color:#fff;font-size:11px;font-weight:700;'
                    'padding:2px 9px;border-radius:12px">{grade}</span>'
                    '&nbsp;<span style="color:#64748b;font-size:12px">👤 {sp}&nbsp;|&nbsp;🏭 {ind}</span>'
                    '</div>'
                    '<div style="color:#475569;font-size:11.5px;margin-top:5px">📍 {loc}</div></div>'
                )

                if st.session_state.edit_mode == "delete":
                    checked = st.session_state.del_checks[i]
                    card = card_tpl.format(
                        border="2px solid #ef4444" if checked else "1px solid #e2e8f0",
                        bg="#fff5f5" if checked else "#ffffff",
                        name=name, sales=sales_fmt, gc=g_color,
                        grade=grade or "—", sp=sp, ind=ind, loc=loc_txt)
                    col_chk, col_card = st.columns([0.4, 11])
                    new_val = col_chk.checkbox("", value=checked,
                                               key=f"chk_{i}_{orig_i}",
                                               label_visibility="collapsed")
                    if new_val != checked:
                        st.session_state.del_checks[i] = new_val; st.rerun()
                    col_card.markdown(card, unsafe_allow_html=True)

                else:
                    is_open = (st.session_state.editing_idx == orig_i)
                    card = card_tpl.format(
                        border="2px solid #2563eb" if is_open else "1px solid #e2e8f0",
                        bg="#f0f7ff" if is_open else "#ffffff",
                        name=name, sales=sales_fmt, gc=g_color,
                        grade=grade or "—", sp=sp, ind=ind, loc=loc_txt)
                    col_card, col_btn = st.columns([10, 1.2])
                    col_card.markdown(card, unsafe_allow_html=True)
                    lbl = "✕ ปิด" if is_open else "✏️ แก้ไข"
                    if col_btn.button(lbl, key=f"ebtn_{i}_{orig_i}", use_container_width=True):
                        st.session_state.editing_idx = None if is_open else orig_i; st.rerun()

                    if is_open:
                        with st.form(key=f"form_edit_{orig_i}"):
                            st.markdown("##### ✏️ แก้ไขข้อมูล")
                            ef1, ef2, ef3 = st.columns([3, 3, 1.5])
                            new_name  = ef1.text_input("🏢 Customer Name", value=name)
                            new_sp    = ef2.text_input("👤 Salesperson",   value=sp)
                            gopts     = ["", "A", "A-", "B", "B-", "C", "C-", "F"]
                            new_grade = ef3.selectbox("⭐ Grade", gopts,
                                                      index=gopts.index(grade) if grade in gopts else 0)
                            ef4, ef5 = st.columns([3, 2])
                            new_ind   = ef4.text_input("🏭 Business Type", value=ind)
                            try:    sv = int(round(float(sales_v)))
                            except: sv = 0
                            new_sales = ef5.number_input("💰 Sales/Year (฿)", value=sv, min_value=0, step=100000)
                            pf1, pf2 = st.columns([2, 3])
                            cur_pc   = _s(row.get("Plus_Code", ""))
                            cur_bk   = int(row.get("Budget_kg",  0) or 0)
                            new_pc   = pf1.text_input("📌 Plus Code", value=cur_pc,
                                                       placeholder="เช่น MC8G+82 กรุงเทพมหานคร")
                            new_bkg  = pf2.number_input("🎯 Budget (kg/yr)", value=cur_bk,
                                                         min_value=0, step=100)
                            pf3, pf4 = st.columns([2, 3])
                            cur_act  = int(row.get("Actual_kg",  0) or 0)
                            cur_ly   = int(row.get("LastYear_kg", 0) or 0)
                            new_act  = pf3.number_input("✅ Actual (kg/yr)",  value=cur_act,
                                                         min_value=0, step=100)
                            new_ly   = pf4.number_input("📅 Last Year (kg)",  value=cur_ly,
                                                         min_value=0, step=100)
                            pm1, pm2, pm3 = st.columns(3)
                            if cur_bk > 0:
                                pm1.metric("📊 Achievement", f"{cur_act/cur_bk*100:.1f}%",
                                           delta=f"{cur_act-cur_bk:+,} kg")
                            if cur_ly > 0:
                                pm2.metric("📈 YoY Growth",
                                           f"{(cur_act-cur_ly)/cur_ly*100:+.1f}%",
                                           delta=f"{cur_act-cur_ly:+,} kg")
                            if cur_bk > 0:
                                pm3.metric("📉 Gap to Budget", f"{cur_bk-cur_act:,} kg",
                                           delta_color="inverse")
                            new_addr = st.text_area("📍 Address (เต็ม)", value=raw_addr, height=90)
                            saved = st.form_submit_button("💾 บันทึก", type="primary",
                                                           use_container_width=True)
                            if saved:
                                clean_pc = clean_plus_code(new_pc)
                                merged_addr = merge_address_parts(new_addr, new_pc)
                                st.session_state.df.at[orig_i, "Customer Name"] = new_name
                                st.session_state.df.at[orig_i, "Salesperson"]   = new_sp
                                st.session_state.df.at[orig_i, "Industry"]      = new_ind
                                st.session_state.df.at[orig_i, "Grade"]         = new_grade
                                st.session_state.df.at[orig_i, "Sales/Year"]    = new_sales
                                st.session_state.df.at[orig_i, "Plus_Code"]     = clean_pc
                                st.session_state.df.at[orig_i, "Budget_kg"]     = new_bkg
                                st.session_state.df.at[orig_i, "Actual_kg"]     = new_act
                                st.session_state.df.at[orig_i, "LastYear_kg"]   = new_ly
                                st.session_state.df.at[orig_i, "Address"]       = merged_addr
                                if merged_addr.strip():
                                    _sub, _dis, _prov, _reg = parse_address(merged_addr)
                                    st.session_state.df.at[orig_i, "Sub-district"] = _sub
                                    st.session_state.df.at[orig_i, "District"]     = _dis
                                    st.session_state.df.at[orig_i, "Province"]     = _prov
                                    st.session_state.df.at[orig_i, "Region"]       = _reg
                                else:
                                    st.session_state.df.at[orig_i, "Sub-district"] = ""
                                    st.session_state.df.at[orig_i, "District"]     = ""
                                    st.session_state.df.at[orig_i, "Province"]     = ""
                                    st.session_state.df.at[orig_i, "Region"]       = "Unknown"
                                st.session_state.df["Region_TH"] = (
                                    st.session_state.df["Region"].map(REGION_EN_TO_TH).fillna("ไม่ระบุ"))
                                st.session_state.editing_idx = None
                                _commit_save(f"บันทึก '{new_name}'")
                                st.rerun()

            if st.session_state.edit_mode == "delete":
                sel_idxs  = [orig_idx[i] for i, v in enumerate(st.session_state.del_checks) if v]
                sel_count = len(sel_idxs)
                sel_names = [_s(subset.iloc[i].get("Customer Name", ""))
                             for i, v in enumerate(st.session_state.del_checks) if v]
                st.divider()
                da, db = st.columns([2.5, 5])
                with da:
                    del_btn = st.button(
                        f"🗑️ ลบที่เลือก ({sel_count})" if sel_count else "🗑️ ลบที่เลือก",
                        type="primary" if sel_count else "secondary",
                        disabled=(sel_count == 0), use_container_width=True)
                with db:
                    if sel_count:
                        prev = ", ".join(sel_names[:3]) + (f" +{sel_count-3}" if sel_count > 3 else "")
                        st.warning(f"จะลบ: **{prev}**")
                if del_btn: st.session_state.confirm_delete = True
                if st.session_state.get("confirm_delete") and sel_count > 0:
                    st.error(f"⚠️ ยืนยันลบ **{sel_count} รายการ**? ไม่สามารถย้อนกลับได้")
                    cc1, cc2, _ = st.columns([1.5, 1.5, 5])
                    with cc1:
                        if st.button("✅ ยืนยัน", type="primary",
                                     use_container_width=True, key="confirm_yes"):
                            st.session_state.df = (
                                st.session_state.df.drop(index=sel_idxs).reset_index(drop=True))
                            st.session_state.del_checks  = []
                            st.session_state.confirm_delete = False
                            _commit_save(f"ลบ {sel_count} รายการ")
                            st.rerun()
                    with cc2:
                        if st.button("❌ ยกเลิก", use_container_width=True, key="confirm_no"):
                            st.session_state.confirm_delete = False; st.rerun()

    with tab_add:
        st.caption("กรอกข้อมูลลูกค้าใหม่ — ระบบจะ Parse Province จาก Address อัตโนมัติ")
        with st.form("form_add", clear_on_submit=True):
            r1c1, r1c2 = st.columns(2)
            n_name = r1c1.text_input("Customer Name *")
            n_sp   = r1c2.text_input("Salesperson")
            r2c1, r2c2, r2c3 = st.columns(3)
            n_ind   = r2c1.text_input("Industry")
            n_grade = r2c2.selectbox("Grade", ["", "A", "A-", "B", "B-", "C", "C-", "F"])
            n_sales = r2c3.number_input("Sales/Year (฿)", min_value=0.0, step=100_000.0)
            n_pc   = st.text_input("📌 Plus Code", placeholder="เช่น MJHG+2F กรุงเทพมหานคร")
            n_addr = st.text_area("Address (ระบุเพื่อ Auto-parse)")
            st.markdown("**หรือระบุที่อยู่เองด้านล่าง** (ว่าง = Auto-parse)")
            r3c1, r3c2, r3c3, r3c4 = st.columns(4)
            n_sub  = r3c1.text_input("Sub-district")
            n_dis  = r3c2.text_input("District")
            n_prov = r3c3.text_input("Province")
            n_reg  = r3c4.selectbox("Region", ["", "Central", "North", "Northeast",
                                                "East", "West", "South"])
            ok = st.form_submit_button("➕ เพิ่มลูกค้า", type="primary", use_container_width=True)

        if ok:
            if not n_name.strip():
                st.error("กรุณากรอก Customer Name")
            else:
                clean_pc = clean_plus_code(n_pc)
                merged_addr = merge_address_parts(n_addr, n_pc)
                auto_sub, auto_dis, auto_prov, auto_reg = parse_address(merged_addr)
                new_row = {
                    "Customer Name": n_name.strip(), "Salesperson": n_sp,
                    "Industry": n_ind, "Grade": n_grade, "Sales/Year": n_sales,
                    "Plus_Code": clean_pc, "Address": merged_addr,
                    "Sub-district": n_sub or auto_sub, "District": n_dis or auto_dis,
                    "Province": n_prov or auto_prov, "Region": n_reg or auto_reg,
                    "Budget_kg": 0, "Actual_kg": 0, "LastYear_kg": 0,
                }
                new_row["Region_TH"] = REGION_EN_TO_TH.get(new_row["Region"], "ไม่ระบุ")
                st.session_state.df = pd.concat(
                    [st.session_state.df, pd.DataFrame([new_row])], ignore_index=True)
                _commit_save(f"เพิ่ม '{n_name}'")
                st.rerun()

    st.divider()
    st.subheader("⬇️ Export ข้อมูลทั้งหมด")
    ec1, ec2 = st.columns(2)
    all_c   = [c for c in TEMPLATE_COLS if c in st.session_state.df.columns]
    exp_all = st.session_state.df[all_c]
    ec1.download_button("📥 Export CSV",   data=exp_all.to_csv(index=False, encoding="utf-8-sig"),
                        file_name="all_customers.csv", mime="text/csv", use_container_width=True)
    ec2.download_button("📥 Export Excel", data=to_excel_bytes(exp_all),
                        file_name="all_customers.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)