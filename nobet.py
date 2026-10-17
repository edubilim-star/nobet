import io
import random
import datetime
import re
import os
import html
import tempfile
import json
import base64
import urllib.error
import urllib.parse
import urllib.request
import hashlib
import unicodedata
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(
    page_title="Okul Nöbet Otomasyon Sistemi",
    layout="wide",
    menu_items={}
)
st.set_option("client.toolbarMode", "auto")

# --- VERİTABANI / SESSION STATE BAŞLANGICI ---
if "school_info" not in st.session_state:
    st.session_state.school_info = {
        "okul_adi": "MEHMET AKİF ERSOY İLKOKULU",
        "mudur_adi": "HAMDİ ALTINTAŞ",
        "not_metni": (
            "Nöbetçi öğretmenler dersler başlamadan 15 dakika önce nöbet defterini imzalayıp görev yerlerinde olacaklar, "
            "en son dersten 15 dakika sonra (öğrencilerin okuldan güvenli şekilde ayrılmaları için gerekli tedbirleri alarak "
            "öğrencileri evlerine gönderdikten sonra) nöbet defterini imzalayacaklardır."
        )
    }

if "locations" not in st.session_state:
    st.session_state.locations = pd.DataFrame({
        "Nöbet Yeri / Bölüm": [
            "BAHÇE", "SINIF 1.KAT", "SINIF 2.KAT", "SINIF 3.KAT", 
            "ANASINIFI KAPI", "EK BİNA 1.KAT", "EK BİNA 2.KAT"
        ],
        "Gerekli Öğretmen Sayısı": [2, 1, 1, 1, 2, 1, 1]
    })

WEEKDAYS = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
def search_normalize(value):
    text = str(value).casefold().replace("ı", "i")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", " ", text).strip()

DEFAULT_EXCEL_SECTIONS = [
    {"Excel Başlığı": "BAHÇE"},
    {"Excel Başlığı": "SINIF VE KORİDORLAR"},
    {"Excel Başlığı": "ANASINIFI KAPI"},
    {"Excel Başlığı": "EK BİNA"},
]

if "teachers_dict_list" not in st.session_state:
    default_teachers = [
        "KÜBRA BIYIKOĞLU", "MELİSA GÜVENÇ", "MERVE GÜLER", "MELİKE ÖDEN",
        "EMEL KAYABAŞI", "AYSEL GÜL", "RABİA ŞAHİN BAKRAKTAR", "EBRU BAĞCI",
        "İBRAHİM KILIÇ", "YUNUS KARACA", "SEHER TAŞ", "ALEV AYÇA AÇIKEL",
        "ŞULE YÜKSELAY", "MERVE AYDEMİR KESİKBAŞ", "MERVE ÜLKER KÖKEN",
        "AYŞEGÜL SEVİM BOZĞOLU", "ESMANUR SEZER", "ÖZGÜR ESER", "DİLARA SAVAŞKAN",
        "MÜNEVER ALICI", "HATİCE BİRER OKAY", "FATMA ÖZMEN", "AYŞE ZORKOL",
        "ÖZGE KEPENEK", "BEYZA NUR ÇAĞLAR SEZGİN", "AYŞEGÜL SÜER", "BERNA ÜLKER",
        "FERİDE ÖZDEMİR", "KÜBRA EVREN", "HATİCE BETÜL TURAN",
        "NİHAL GÖKSU"
    ]
    t_list = []
    for t in default_teachers:
        t_list.append({
            "Ad Soyad": t,
            "Sabit Yer": "Yok",
            "Hariç Tutulan Bölümler": [],
            "İzinli Günler": [],
            "Haftada 2 Nöbet İstiyor": False
        })
    st.session_state.teachers_dict_list = t_list
st.session_state.teachers_dict_list = [
    t for t in st.session_state.teachers_dict_list
    if str(t.get("Ad Soyad", "")).strip().upper() != "AYŞENUR AKBOĞA"
]

if "holidays" not in st.session_state:
    st.session_state.holidays = pd.DataFrame(columns=["Tarih", "Tatil Açıklaması"])

def normalized_holidays(frame):
    """Return holiday rows with date objects, valid descriptions, and no duplicates."""
    if frame is None or frame.empty:
        return pd.DataFrame(columns=["Tarih", "Tatil Açıklaması"])
    result = frame.reindex(columns=["Tarih", "Tatil Açıklaması"]).copy()
    result["Tarih"] = pd.to_datetime(result["Tarih"], errors="coerce").dt.date
    result["Tatil Açıklaması"] = result["Tatil Açıklaması"].fillna("").astype(str).str.strip()
    result = result.dropna(subset=["Tarih"]).drop_duplicates(subset=["Tarih"], keep="last")
    return result.reset_index(drop=True)

def is_holiday(frame, day):
    frame = normalized_holidays(frame)
    return frame[frame["Tarih"] == day]

if "mudur_yardimcilari" not in st.session_state:
    st.session_state.mudur_yardimcilari = ["AYŞENUR AKBOĞA"]

if "history" not in st.session_state:
    st.session_state.history = pd.DataFrame(columns=["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi", "Kaynak"])
else:
    if "Tarih" not in st.session_state.history.columns:
        st.session_state.history["Tarih"] = pd.NaT
    st.session_state.history = st.session_state.history.drop(
        st.session_state.history[st.session_state.history.get("Öğretmen", pd.Series(dtype=str)).astype(str).str.strip().str.upper() == "AYŞENUR AKBOĞA"].index,
        errors="ignore"
    )

if "current_schedule" not in st.session_state:
    st.session_state.current_schedule = None
if "saved_schedules" not in st.session_state:
    st.session_state.saved_schedules = {}
if "mudur_yardimcisi_sayilari" not in st.session_state:
    st.session_state.mudur_yardimcisi_sayilari = {day: 1 for day in WEEKDAYS}
if "mudur_yardimcisi_gecmisi" not in st.session_state:
    st.session_state.mudur_yardimcisi_gecmisi = {}
if "excel_headings" not in st.session_state:
    st.session_state.excel_headings = {"Gün Başlığı": "NÖBET GÜNÜ", "Müdür Yardımcısı Başlığı": "MÜDÜR YRD"}
if "processed_uploads" not in st.session_state:
    st.session_state.processed_uploads = set()

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nobet_verileri.json")

def github_storage_settings():
    """Read optional GitHub persistence settings from Streamlit secrets."""
    try:
        secrets = st.secrets
        token = str(secrets.get("GITHUB_TOKEN", "")).strip()
        repository = str(secrets.get("GITHUB_REPOSITORY", "")).strip()
        branch = str(secrets.get("GITHUB_BRANCH", "main")).strip() or "main"
        data_path = str(secrets.get("GITHUB_DATA_PATH", "nobet_verileri.json")).strip()
    except Exception:
        # Local development can continue without a secrets.toml file.
        return None
    if not token and not repository:
        return None
    if not token or not repository or "/" not in repository:
        raise ValueError("GitHub Secrets içinde GITHUB_TOKEN ve sahip/depo biçiminde GITHUB_REPOSITORY birlikte tanımlanmalı.")
    return {"token": token, "repository": repository, "branch": branch, "path": data_path}

def github_contents_request(settings, method="GET", payload=None):
    """Call GitHub's Contents API for the persistent JSON file."""
    encoded_path = urllib.parse.quote(settings["path"].strip("/"), safe="/")
    url = f"https://api.github.com/repos/{settings['repository']}/contents/{encoded_path}"
    if method == "GET":
        url += "?ref=" + urllib.parse.quote(settings["branch"], safe="")
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    request = urllib.request.Request(
        url,
        data=body,
        method=method,
        headers={
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {settings['token']}",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "okul-nobet-streamlit",
            **({"Content-Type": "application/json"} if body is not None else {}),
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            response_body = response.read()
            return json.loads(response_body.decode("utf-8")) if response_body else {}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        # Do not surface request headers or the token in user-facing errors.
        raise RuntimeError(f"GitHub API isteği başarısız (HTTP {exc.code}): {detail[:300]}") from None
    except (urllib.error.URLError, TimeoutError) as exc:
        raise RuntimeError(f"GitHub'a bağlanılamadı: {getattr(exc, 'reason', exc)}") from None

def read_github_app_data(settings):
    try:
        result = github_contents_request(settings)
    except RuntimeError as exc:
        if "HTTP 404" in str(exc):
            return {}
        raise
    content = result.get("content", "")
    if result.get("encoding") != "base64" or not content:
        return {}
    return json.loads(base64.b64decode(content).decode("utf-8"))

def write_github_app_data(settings, data):
    # Skip redundant commits caused by normal Streamlit reruns.
    try:
        current = github_contents_request(settings)
    except RuntimeError as exc:
        if "HTTP 404" not in str(exc):
            raise
        current = {}
    content = json.dumps(data, ensure_ascii=False, indent=2)
    current_content = current.get("content", "")
    if current_content and current.get("encoding") == "base64":
        try:
            existing = base64.b64decode(current_content).decode("utf-8")
            if json.loads(existing) == data:
                return False
        except (ValueError, json.JSONDecodeError, UnicodeDecodeError):
            pass
    payload = {
        "message": "Nöbet uygulaması verilerini güncelle",
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": settings["branch"],
    }
    if current.get("sha"):
        payload["sha"] = current["sha"]
    github_contents_request(settings, method="PUT", payload=payload)
    return True

def _records_for_json(frame):
    if frame is None or frame.empty:
        return []
    return json.loads(frame.to_json(orient="records", date_format="iso"))

def save_app_data():
    data = {
        "school_info": st.session_state.school_info,
        "locations": _records_for_json(st.session_state.locations),
        "excel_sections": _records_for_json(st.session_state.excel_sections),
        "excel_headings": st.session_state.excel_headings,
        "teachers_dict_list": st.session_state.teachers_dict_list,
        "holidays": _records_for_json(st.session_state.holidays),
        "mudur_yardimcilari": st.session_state.mudur_yardimcilari,
        "mudur_yardimcisi_sayilari": st.session_state.mudur_yardimcisi_sayilari,
        "mudur_yardimcisi_gecmisi": st.session_state.mudur_yardimcisi_gecmisi,
        "history": _records_for_json(st.session_state.history),
        "current_schedule": _records_for_json(st.session_state.current_schedule),
        "saved_schedules": {
            name: _records_for_json(schedule if isinstance(schedule, pd.DataFrame) else pd.DataFrame(schedule))
            for name, schedule in st.session_state.saved_schedules.items()
        },
    }
    temp_path = DATA_FILE + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    os.replace(temp_path, DATA_FILE)
    settings = st.session_state.get("_github_storage_settings")
    if settings:
        try:
            write_github_app_data(settings, data)
            st.session_state["_github_storage_error"] = None
        except Exception as exc:
            st.session_state["_github_storage_error"] = str(exc)

def _frame_from_records(records, columns):
    frame = pd.DataFrame(records or [])
    if frame.empty:
        return pd.DataFrame(columns=columns)
    for column in columns:
        if column not in frame.columns:
            frame[column] = pd.NA
    if "Tarih" in frame.columns:
        frame["Tarih"] = pd.to_datetime(frame["Tarih"], errors="coerce").dt.date
    return frame.reindex(columns=columns)

if not st.session_state.get("_persistence_loaded", False):
    saved_data = {}
    try:
        settings = github_storage_settings()
        st.session_state["_github_storage_settings"] = settings
        if settings:
            saved_data = read_github_app_data(settings)
            # Seed a newly configured GitHub store from an existing local file.
            if not saved_data:
                try:
                    with open(DATA_FILE, "r", encoding="utf-8") as file:
                        saved_data = json.load(file)
                except (OSError, json.JSONDecodeError):
                    saved_data = {}
        else:
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as file:
                    saved_data = json.load(file)
            except (OSError, json.JSONDecodeError):
                saved_data = {}
    except Exception as exc:
        st.session_state["_github_storage_settings"] = None
        st.session_state["_github_storage_error"] = str(exc)
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as file:
                saved_data = json.load(file)
        except (OSError, json.JSONDecodeError):
            saved_data = {}
    for key in ("school_info", "teachers_dict_list", "mudur_yardimcilari", "mudur_yardimcisi_sayilari", "mudur_yardimcisi_gecmisi"):
        if key in saved_data:
            st.session_state[key] = saved_data[key]
    if "locations" in saved_data:
        st.session_state.locations = pd.DataFrame(saved_data["locations"])
    if "excel_sections" in saved_data:
        st.session_state.excel_sections = pd.DataFrame(saved_data["excel_sections"])
    if "excel_headings" in saved_data:
        st.session_state.excel_headings = saved_data["excel_headings"]
    if "holidays" in saved_data:
        st.session_state.holidays = normalized_holidays(pd.DataFrame(saved_data["holidays"]))
    if "history" in saved_data:
        st.session_state.history = _frame_from_records(
            saved_data["history"], ["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi", "Kaynak"]
        )
    if "current_schedule" in saved_data and saved_data["current_schedule"]:
        st.session_state.current_schedule = _frame_from_records(
            saved_data["current_schedule"], ["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi", "Kaynak"]
        )
    st.session_state.saved_schedules = {
        name: pd.DataFrame(records) for name, records in saved_data.get("saved_schedules", {}).items()
    }
    if st.session_state.saved_schedules:
        dated_schedules = []
        for name, frame in st.session_state.saved_schedules.items():
            if "Tarih" in frame.columns:
                dates = pd.to_datetime(frame["Tarih"], errors="coerce").dropna()
                if not dates.empty:
                    dated_schedules.append((dates.max(), name, frame))
        if dated_schedules:
            _, latest_name, latest_frame = max(dated_schedules, key=lambda item: item[0])
            st.session_state.current_schedule = _frame_from_records(
                latest_frame.to_dict("records"),
                ["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi", "Kaynak"]
            )
    st.session_state.teachers_dict_list = [
        teacher for teacher in st.session_state.teachers_dict_list
        if str(teacher.get("Ad Soyad", "")).strip().upper() != "AYŞENUR AKBOĞA"
    ]
    st.session_state["_persistence_loaded"] = True

if "Kaynak" not in st.session_state.history.columns:
    st.session_state.history["Kaynak"] = st.session_state.history["Tarih/Hafta"].apply(
        lambda name: "Oluşturulan" if str(name) in st.session_state.saved_schedules else "Eski Liste"
    )
else:
    missing_source = st.session_state.history["Kaynak"].isna()
    st.session_state.history.loc[missing_source, "Kaynak"] = st.session_state.history.loc[
        missing_source, "Tarih/Hafta"
    ].apply(lambda name: "Oluşturulan" if str(name) in st.session_state.saved_schedules else "Eski Liste")

if "Aktif Günler" not in st.session_state.locations.columns:
    st.session_state.locations["Aktif Günler"] = [WEEKDAYS.copy() for _ in range(len(st.session_state.locations))]
else:
    st.session_state.locations["Aktif Günler"] = st.session_state.locations["Aktif Günler"].apply(
        lambda days: [day for day in days if day in WEEKDAYS]
        if isinstance(days, (list, tuple, set)) else WEEKDAYS.copy()
    )
if "Excel'de Yer Başlığını Göster" not in st.session_state.locations.columns:
    st.session_state.locations["Excel'de Yer Başlığını Göster"] = True
else:
    st.session_state.locations["Excel'de Yer Başlığını Göster"] = st.session_state.locations["Excel'de Yer Başlığını Göster"].fillna(True).astype(bool)
if "Excel Sırası" in st.session_state.locations.columns:
    st.session_state.locations = st.session_state.locations.drop(columns=["Excel Sırası"])

if "excel_sections" not in st.session_state:
    st.session_state.excel_sections = pd.DataFrame(DEFAULT_EXCEL_SECTIONS)
if st.session_state.excel_sections.empty or "Excel Başlığı" not in st.session_state.excel_sections.columns:
    st.session_state.excel_sections = pd.DataFrame(DEFAULT_EXCEL_SECTIONS)

def guess_excel_section(location_name, headers=None):
    headers = headers or st.session_state.excel_sections["Excel Başlığı"].astype(str).tolist()
    normalized = search_normalize(location_name)
    preferred = None
    if "bahce" in normalized:
        preferred = "BAHÇE"
    elif "anasinifi" in normalized:
        preferred = "ANASINIFI KAPI"
    elif "ek bina" in normalized:
        preferred = "EK BİNA"
    elif any(word in normalized for word in ("sinif", "kat", "koridor", "giris")):
        preferred = "SINIF VE KORİDORLAR"
    return preferred if preferred in headers else (headers[0] if headers else "")

if "Excel Bölümü" not in st.session_state.locations.columns:
    st.session_state.locations["Excel Bölümü"] = st.session_state.locations["Nöbet Yeri / Bölüm"].apply(guess_excel_section)
else:
    excel_headers = st.session_state.excel_sections["Excel Başlığı"].astype(str).tolist() if "Excel Başlığı" in st.session_state.excel_sections.columns else [row["Excel Başlığı"] for row in DEFAULT_EXCEL_SECTIONS]
    st.session_state.locations["Excel Bölümü"] = st.session_state.locations.apply(
        lambda row: row["Excel Bölümü"] if row["Excel Bölümü"] in excel_headers else guess_excel_section(row["Nöbet Yeri / Bölüm"], excel_headers),
        axis=1
    )

st.markdown("""
<style>
div[data-testid="stAppViewBlockContainer"], .block-container {
  padding-top: 0.45rem !important;
}
.app-footer { margin-top: 1.5rem; padding: 0.75rem 0 0.3rem; border-top: 1px solid rgba(128, 128, 128, 0.25); color: #64748b; font-size: 0.85rem; font-weight: 600; letter-spacing: 0.04em; text-align: left; }
@media (max-width: 700px) {
  .block-container { padding-left: 0.75rem; padding-right: 0.75rem; padding-top: 0.35rem !important; }
  [data-testid="stDataFrame"] { max-width: 100%; overflow-x: auto; }
}
</style>
""", unsafe_allow_html=True)

# --- UYGULAMA GEZİNME MENÜSÜ ---
MENU_OPTIONS = [
    "📅 Nöbet Listesi Oluştur",
    "👥 Öğretmen Yönetimi",
    "📍 Nöbet Yerleri ve Kontenjan",
    "⚙️ Okul, Tatil ve Ayarlar",
    "📂 Eski Listeler & Arşiv",
    "📊 İstatistikler",
]
if st.session_state.get("main_navigation") not in MENU_OPTIONS:
    st.session_state.main_navigation = MENU_OPTIONS[0]

nav_columns = st.columns(len(MENU_OPTIONS))
for nav_index, nav_label in enumerate(MENU_OPTIONS):
    with nav_columns[nav_index]:
        if st.button(
            nav_label,
            key=f"main_navigation_button_{nav_index}",
            type="primary" if st.session_state.main_navigation == nav_label else "secondary",
            use_container_width=True,
        ):
            st.session_state.main_navigation = nav_label
            st.rerun()

main_menu = st.session_state.main_navigation

# ================= EXCEL ÜRETME FONKSİYONU =================
def create_custom_excel(df, hafta_baslik, my_listesi, start_date, school_name, principal_name, note_text, holidays_df, locations_df, sections_df, headings):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nöbet Çizelgesi"
    ws.views.sheetView[0].showGridLines = True

    section_headers = [str(value).strip() for value in sections_df["Excel Başlığı"].tolist() if str(value).strip()]
    if not section_headers:
        section_headers = ["BAHÇE", "SINIF VE KORİDORLAR", "ANASINIFI KAPI", "EK BİNA"]
    location_groups = {}
    for original_index, row in locations_df.iterrows():
        location_name = str(row.get("Nöbet Yeri / Bölüm", "")).strip()
        if not location_name or location_name.casefold() == "nan":
            continue
        section = str(row.get("Excel Bölümü", "")).strip()
        if section not in section_headers:
            section = guess_excel_section(location_name, section_headers)
        show_location_title = row.get("Excel'de Yer Başlığını Göster", True)
        if pd.isna(show_location_title):
            show_location_title = True
        location_groups[location_name] = (section, original_index, bool(show_location_title))
    first_data_col = 2
    deputy_col = first_data_col + len(section_headers)
    last_col = deputy_col

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws["A1"] = school_name
    ws["A1"].font = Font(name="Calibri", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws["A2"] = f"{hafta_baslik} NÖBET ÇİZELGESİ"
    ws["A2"].font = Font(name="Calibri", size=12, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=3, column=1, value=headings.get("Gün Başlığı", "NÖBET GÜNÜ"))
    for index, section_header in enumerate(section_headers):
        ws.cell(row=3, column=first_data_col + index, value=section_header)
    ws.cell(row=3, column=deputy_col, value=headings.get("Müdür Yardımcısı Başlığı", "MÜDÜR YRD"))

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000")
    )

    for col in range(1, last_col + 1):
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    gunler = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
    start_row = 4
    for idx, gun in enumerate(gunler):
        row = start_row + idx
        gun_tarihi = start_date + datetime.timedelta(days=idx)
        tarih_str = gun_tarihi.strftime("%d.%m.%Y")
        
        tatil_aciklama = ""
        if not holidays_df.empty and "Tarih" in holidays_df.columns:
            match_tatil = is_holiday(holidays_df, gun_tarihi)
            if not match_tatil.empty:
                tatil_aciklama = match_tatil["Tatil Açıklaması"].values[0]

        gun_baslik = f"{gun}\n({tarih_str})"
        gun_df = df[df["Gün"] == gun]

        ws.cell(row=row, column=1, value=gun_baslik).font = Font(name="Calibri", size=10, bold=True)
        for index, section_header in enumerate(section_headers):
            group_locations = [
                (name, order, show_title) for name, (group, order, show_title) in location_groups.items()
                if search_normalize(group) == search_normalize(section_header)
            ]
            group_locations.sort(key=lambda item: item[1])
            if tatil_aciklama:
                section_value = f"*** TATİL ({tatil_aciklama}) ***" if index == 0 else ""
            else:
                lines = []
                for location, _, show_title in group_locations:
                    matches = gun_df[gun_df["Bölüm/Yer"] == location]["Öğretmen"].tolist()
                    clean_matches = [name.split(" (")[0] for name in matches]
                    if clean_matches:
                        if show_title:
                            lines.append(f"{location}: {clean_matches[0]}" + ("\n" + "\n".join(clean_matches[1:]) if len(clean_matches) > 1 else ""))
                        else:
                            lines.append("\n".join(clean_matches))
                section_value = "\n".join(lines)
            ws.cell(row=row, column=first_data_col + index, value=section_value)

        assistants = my_listesi.get(gun, []) if isinstance(my_listesi, dict) else my_listesi
        ws.cell(row=row, column=deputy_col, value="\n".join(assistants))
        
        ws.row_dimensions[row].height = 70
        
        for col in range(1, last_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if col in [1, deputy_col] else "left", vertical="center", wrap_text=True)
            cell.font = Font(name="Calibri", size=10, bold=(col==1))

    footer_row = start_row + len(gunler) + 1
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row+1, end_column=last_col)
    note_cell = ws.cell(row=footer_row, column=1)
    note_cell.value = note_text
    note_cell.font = Font(name="Calibri", size=9, italic=True)
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    sig_row = footer_row + 3
    ws.cell(row=sig_row, column=last_col, value=principal_name).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=sig_row, column=last_col).alignment = Alignment(horizontal="center")
    ws.cell(row=sig_row+1, column=last_col, value="OKUL MÜDÜRÜ").font = Font(name="Calibri", size=10, bold=True)
    ws.cell(row=sig_row+1, column=last_col).alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 18
    for col in range(first_data_col, deputy_col):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 32
    ws.column_dimensions[openpyxl.utils.get_column_letter(deputy_col)].width = 24
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Saving directly to BytesIO can fail with lxml.etree.SerialisationError(IO_WRITE)
    # in some Windows/openpyxl combinations. Save to a real temporary file, then
    # return a fresh in-memory stream for Streamlit's download button.
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_path = temp_file.name
        wb.save(temp_path)
        with open(temp_path, "rb") as saved_file:
            output = io.BytesIO(saved_file.read())
        output.seek(0)
        return output
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


# ================= PDF ÜRETME FONKSİYONU =================
def create_pdf(df, hafta_baslik, my_listesi, school_name, principal_name, note_text, start_date, holidays_df, locations_df=None):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    bundled_fonts = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    font_candidates = [
        (os.path.join(bundled_fonts, "NotoSans-Regular.ttf"), os.path.join(bundled_fonts, "NotoSans-Bold.ttf")),
        (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
        (r"C:\Windows\Fonts\calibri.ttf", r"C:\Windows\Fonts\calibrib.ttf"),
        (r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\segoeuib.ttf"),
        (r"C:\Windows\Fonts\tahoma.ttf", r"C:\Windows\Fonts\tahomabd.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    regular_font, bold_font = "Helvetica", "Helvetica-Bold"
    unicode_font_found = False
    for regular_path, bold_path in font_candidates:
        if os.path.exists(regular_path) and os.path.exists(bold_path):
            pdfmetrics.registerFont(TTFont("SchoolUnicode", regular_path))
            pdfmetrics.registerFont(TTFont("SchoolUnicode-Bold", bold_path))
            pdfmetrics.registerFontFamily(
                "SchoolUnicode",
                normal="SchoolUnicode",
                bold="SchoolUnicode-Bold",
                italic="SchoolUnicode",
                boldItalic="SchoolUnicode-Bold",
            )
            regular_font, bold_font = "SchoolUnicode", "SchoolUnicode-Bold"
            unicode_font_found = True
            break
    if not unicode_font_found:
        raise RuntimeError("Türkçe karakter destekli yazı tipi bulunamadı. Uygulamayla birlikte fonts/NotoSans-Regular.ttf ve fonts/NotoSans-Bold.ttf dosyalarını yükleyin.")
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=bold_font, fontSize=14, alignment=1)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName=regular_font, fontSize=9)
    right_style = ParagraphStyle('RightStyle', parent=cell_style, alignment=2)
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontName=bold_font, fontSize=10, textColor=colors.whitesmoke, alignment=1)

    def safe_text(txt):
        # Escape text for ReportLab Paragraph while preserving Turkish Unicode.
        return html.escape(str(txt), quote=False).replace("\n", "<br/>")

    elements.append(Paragraph(safe_text(school_name), title_style))
    elements.append(Paragraph(safe_text(f"{hafta_baslik} NÖBET ÇİZELGESİ"), title_style))
    elements.append(Spacer(1, 10))

    table_data = [[
        Paragraph("NÖBET GÜNÜ", header_style),
        Paragraph("BAHÇE", header_style),
        Paragraph("SINIF VE KORİDORLAR", header_style),
        Paragraph("ANASINIFI KAPI", header_style),
        Paragraph("EK BİNA", header_style),
        Paragraph("MÜDÜR YRD", header_style)
    ]]

    gunler = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
    for idx, gun in enumerate(gunler):
        gun_tarihi = start_date + datetime.timedelta(days=idx)
        tatil_aciklama = ""
        if not holidays_df.empty and "Tarih" in holidays_df.columns:
            match_tatil = is_holiday(holidays_df, gun_tarihi)
            if not match_tatil.empty:
                tatil_aciklama = match_tatil["Tatil Açıklaması"].values[0]

        gun_df = df[df["Gün"] == gun]
        def get_all_t(b_adi):
            if tatil_aciklama:
                return safe_text(f"*** TATİL ({tatil_aciklama}) ***" if b_adi == "BAHÇE" else "")
            matches = gun_df[gun_df["Bölüm/Yer"] == b_adi]["Öğretmen"].tolist()
            clean_matches = [safe_text(m.split(" (")[0]) for m in matches]
            return "<br/>".join(clean_matches)

        def get_section_text(section_name):
            if tatil_aciklama:
                return safe_text(f"*** TATİL ({tatil_aciklama}) ***" if section_name == "BAHÇE" else "")
            if locations_df is None or locations_df.empty or "Excel Bölümü" not in locations_df.columns:
                return ""
            section_locations = locations_df[
                locations_df["Excel Bölümü"].astype(str).map(search_normalize) == search_normalize(section_name)
            ]
            parts = []
            for _, location_row in section_locations.iterrows():
                location = str(location_row.get("Nöbet Yeri / Bölüm", "")).strip()
                if not location:
                    continue
                names = gun_df[gun_df["Bölüm/Yer"].astype(str) == location]["Öğretmen"].astype(str).tolist()
                names = [name.split(" (")[0] for name in names]
                if not names:
                    continue
                show_title = location_row.get("Excel'de Yer Başlığını Göster", True)
                if pd.isna(show_title):
                    show_title = True
                if bool(show_title):
                    parts.append(f"<b>{safe_text(location)}:</b> {safe_text(names[0])}" + "".join(f"<br/>{safe_text(name)}" for name in names[1:]))
                else:
                    parts.append("<br/>".join(safe_text(name) for name in names))
            return "<br/>".join(parts)

        sinif_text = get_section_text("SINIF VE KORİDORLAR")
        ek_text = get_section_text("EK BİNA")

        assistants = my_listesi.get(gun, []) if isinstance(my_listesi, dict) else my_listesi
        my_str = "<br/>".join(safe_text(name) for name in assistants)
        table_data.append([
            Paragraph(f"<b>{safe_text(gun)}</b><br/>({gun_tarihi.strftime('%d.%m.%Y')})", cell_style),
            Paragraph(get_all_t('BAHÇE'), cell_style),
            Paragraph(sinif_text, cell_style),
            Paragraph(get_section_text('ANASINIFI KAPI'), cell_style),
            Paragraph(ek_text, cell_style),
            Paragraph(my_str, cell_style)
        ])

    t = Table(table_data, colWidths=[90, 120, 150, 120, 120, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4F81BD')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    
    elements.append(t)
    elements.append(Spacer(1, 15))
    elements.append(Paragraph(safe_text(note_text), cell_style))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(f"<b>{safe_text(principal_name)}</b>", right_style))
    elements.append(Paragraph("OKUL MÜDÜRÜ", right_style))

    doc.build(elements)
    pdf_buffer.seek(0)
    return pdf_buffer


def create_schedule_image(df, hafta_baslik, my_listesi, start_date, school_name, principal_name, holidays_df, locations_df, sections_df):
    """Render the weekly duty schedule as a downloadable Turkish-capable PNG."""
    section_headers = sections_df["Excel Başlığı"].astype(str).tolist() if sections_df is not None and "Excel Başlığı" in sections_df else ["BAHÇE", "SINIF VE KORİDORLAR", "ANASINIFI KAPI", "EK BİNA"]
    if not section_headers:
        section_headers = ["BAHÇE", "SINIF VE KORİDORLAR", "ANASINIFI KAPI", "EK BİNA"]
    columns = ["NÖBET GÜNÜ", *section_headers, "MÜDÜR YRD"]
    width = 2600
    margin, title_height = 60, 180
    table_width = width - margin * 2
    widths = [250] + [int((table_width - 250 - 250) / len(section_headers))] * len(section_headers) + [250]
    widths[1] += table_width - sum(widths)
    bundled_fonts = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    font_paths = [os.path.join(bundled_fonts, "NotoSans-Regular.ttf"), r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\calibri.ttf", r"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
    bold_paths = [os.path.join(bundled_fonts, "NotoSans-Bold.ttf"), r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\calibrib.ttf", r"/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
    regular_path = next((path for path in font_paths if os.path.exists(path)), None)
    bold_path = next((path for path in bold_paths if os.path.exists(path)), regular_path)
    if not regular_path:
        raise RuntimeError("PNG oluşturmak için Unicode destekli yazı tipi bulunamadı. Uygulamayla birlikte fonts/NotoSans-Regular.ttf ve fonts/NotoSans-Bold.ttf dosyalarını yükleyin.")
    title_font = ImageFont.truetype(bold_path, 42)
    header_font = ImageFont.truetype(bold_path, 28)
    body_font = ImageFont.truetype(regular_path, 25)
    day_font = ImageFont.truetype(bold_path, 25)
    measure_draw = ImageDraw.Draw(Image.new("RGB", (1, 1), "white"))

    def wrap_to_width(text, font, max_width):
        output = []
        source_lines = [line.strip() for line in str(text).splitlines() if line.strip()]
        for source_line in source_lines or [""]:
            words = source_line.split()
            current = ""
            for word in words:
                candidate = f"{current} {word}".strip()
                if current and measure_draw.textbbox((0, 0), candidate, font=font)[2] > max_width:
                    output.append(current)
                    current = word
                else:
                    current = candidate
                if measure_draw.textbbox((0, 0), current, font=font)[2] > max_width:
                    fragment = ""
                    for character in current:
                        if fragment and measure_draw.textbbox((0, 0), fragment + character, font=font)[2] > max_width:
                            output.append(fragment)
                            fragment = character
                        else:
                            fragment += character
                    current = fragment
            output.append(current)
        return output

    location_info = {}
    for order, (_, loc_row) in enumerate(locations_df.iterrows()):
        loc = str(loc_row.get("Nöbet Yeri / Bölüm", "")).strip()
        if not loc:
            continue
        section = str(loc_row.get("Excel Bölümü", "")).strip()
        if section not in section_headers:
            section = guess_excel_section(loc, section_headers)
        show_title_value = loc_row.get("Excel'de Yer Başlığını Göster", True)
        show_title = True if pd.isna(show_title_value) else bool(show_title_value)
        location_info[loc] = (section, order, show_title)

    holidays = normalized_holidays(holidays_df)
    days = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
    daily_values = []
    for day_idx, day in enumerate(days):
        day_date = start_date + datetime.timedelta(days=day_idx)
        holiday_rows = holidays[holidays["Tarih"] == day_date] if not holidays.empty else pd.DataFrame()
        values = [f"{day}\n({day_date.strftime('%d.%m.%Y')})"]
        row_df = df[df["Gün"] == day]
        for section in section_headers:
            if not holiday_rows.empty:
                values.append(f"*** TATİL ({holiday_rows.iloc[0]['Tatil Açıklaması']}) ***" if section == section_headers[0] else "")
                continue
            content = []
            for place, (place_section, _, show_title) in sorted(location_info.items(), key=lambda item: item[1][1]):
                if search_normalize(place_section) != search_normalize(section):
                    continue
                teachers = row_df[row_df["Bölüm/Yer"].astype(str) == place]["Öğretmen"].astype(str).tolist()
                teachers = [name.split(" (")[0] for name in teachers]
                if teachers:
                    content.append((f"{place}: {teachers[0]}\n" + "\n".join(teachers[1:])) if show_title else "\n".join(teachers))
            values.append("\n".join(content))
        assistants = my_listesi.get(day, []) if isinstance(my_listesi, dict) else my_listesi
        values.append("\n".join(assistants))
        daily_values.append(values)

    header_line_height = measure_draw.textbbox((0, 0), "Ay", font=header_font)[3] + 8
    header_height = max(100, max(len(wrap_to_width(label, header_font, widths[idx] - 24)) for idx, label in enumerate(columns)) * header_line_height + 24)
    body_line_height = measure_draw.textbbox((0, 0), "Ay", font=body_font)[3] + 2
    row_heights = []
    for values in daily_values:
        line_count = max(
            len(wrap_to_width(value, day_font if col_idx == 0 else body_font, widths[col_idx] - 24))
            for col_idx, value in enumerate(values)
        )
        row_heights.append(max(150, line_count * body_line_height + 28))
    footer_height = 120
    height = title_height + header_height + sum(row_heights) + footer_height
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    draw.text((width // 2, 30), str(school_name), fill="#1f2937", font=title_font, anchor="mt")
    draw.text((width // 2, 95), f"{hafta_baslik} NÖBET ÇİZELGESİ", fill="#1f2937", font=header_font, anchor="mt")
    x_positions = [margin]
    for col_width in widths:
        x_positions.append(x_positions[-1] + col_width)
    table_top = title_height
    for col_idx, heading in enumerate(columns):
        x0, x1 = x_positions[col_idx], x_positions[col_idx + 1]
        draw.rectangle((x0, table_top, x1, table_top + header_height), fill="#4F81BD", outline="#263238", width=2)
        lines = wrap_to_width(heading, header_font, x1 - x0 - 24)
        y = table_top + max(12, (header_height - len(lines) * header_line_height) // 2)
        for line in lines:
            draw.text(((x0 + x1) // 2, y), line, fill="white", font=header_font, anchor="mt")
            y += header_line_height

    top = table_top + header_height
    for day_idx, values in enumerate(daily_values):
        row_height = row_heights[day_idx]
        for col_idx, value in enumerate(values):
            x0, x1 = x_positions[col_idx], x_positions[col_idx + 1]
            draw.rectangle((x0, top, x1, top + row_height), fill="white", outline="#374151", width=2)
            lines = wrap_to_width(value, day_font if col_idx == 0 else body_font, x1 - x0 - 24)
            y = top + 12
            for line in lines:
                draw.text((x0 + 12, y), line, fill="#111827", font=day_font if col_idx == 0 else body_font)
                y += body_line_height
        top += row_height

    footer_y = top + 35
    draw.text((width - margin, footer_y), str(principal_name), fill="#111827", font=body_font, anchor="rt")
    draw.text((width - margin, footer_y + 40), "OKUL MÜDÜRÜ", fill="#111827", font=body_font, anchor="rt")
    output = io.BytesIO()
    image.save(output, format="PNG")
    output.seek(0)
    return output


# ================= 1. SEKME: NÖBET LİSTESİ OLUŞTUR =================
if main_menu == "📅 Nöbet Listesi Oluştur":
    st.header("Haftalık Nöbet Çizelgesi ve Dağıtım Sihirbazı")

    def distribute_deputies(start_date, daily_counts, deputy_names, saved_deputy_history, active_days=None):
        days = WEEKDAYS
        active_days = set(days if active_days is None else active_days)
        names = list(dict.fromkeys(str(name).strip() for name in deputy_names if str(name).strip()))
        if not names:
            return {day: [] for day in days}, False
        prior_start = start_date - datetime.timedelta(days=7)
        prior_entry = None
        for entry in saved_deputy_history.values():
            try:
                entry_start = datetime.date.fromisoformat(entry.get("start_date", ""))
            except (TypeError, ValueError):
                continue
            if entry_start == prior_start:
                prior_entry = entry
                break
        pointer = 0
        if prior_entry:
            prior_assignments = prior_entry.get("assignments", {})
            last_prior_duty = next((prior_assignments.get(day, []) for day in reversed(days) if prior_assignments.get(day)), [])
            if last_prior_duty:
                pointer = (names.index(last_prior_duty[-1]) + 1) % len(names) if last_prior_duty[-1] in names else 0
        assignments = {}
        previous_day = set()
        consecutive_fallback = False
        for day in days:
            if day not in active_days:
                assignments[day] = []
                continue
            requested = max(0, min(int(daily_counts.get(day, 1)), len(names)))
            rotated = names[pointer:] + names[:pointer]
            available = [name for name in rotated if name not in previous_day]
            chosen = available[:requested]
            if len(chosen) < requested:
                consecutive_fallback = True
                chosen.extend(name for name in rotated if name not in chosen and len(chosen) < requested)
            assignments[day] = chosen
            previous_day = set(chosen)
            if names:
                pointer = (pointer + max(requested, 1)) % len(names)
        return assignments, consecutive_fallback
    
    col1, col2 = st.columns(2)
    with col1:
        default_start_date = datetime.date(2026, 10, 19)
        if st.session_state.current_schedule is not None and not st.session_state.current_schedule.empty and "Tarih" in st.session_state.current_schedule:
            schedule_dates = pd.to_datetime(st.session_state.current_schedule["Tarih"], errors="coerce").dropna()
            if not schedule_dates.empty:
                latest_date = schedule_dates.max().date()
                default_start_date = latest_date - datetime.timedelta(days=latest_date.weekday())
        baslangic_tarihi = st.date_input("Haftanın Başlangıç Tarihi (Pazartesi)", value=default_start_date)
    with col2:
        bitis_tarihi = baslangic_tarihi + datetime.timedelta(days=4)
        aylar = {1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN", 5: "MAYIS", 6: "HAZİRAN", 7: "TEMMUZ", 8: "AĞUSTOS", 9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"}
        default_donem = f"{baslangic_tarihi.day}-{bitis_tarihi.day} {aylar[bitis_tarihi.month]}"
        
        hafta_adi = st.text_input("Nöbet Dönem Adı (Otomatik Gelir)", value=default_donem)

    generate_left, generate_center, generate_right = st.columns([1, 1, 1], gap="small")
    with generate_center:
        generate_clicked = st.button(
            "🎲 Akıllı Kurallı Dağıtımla Çizelgeyi Oluştur",
            type="primary",
            use_container_width=True,
        )
    if "overwrite_pending" not in st.session_state:
        st.session_state.overwrite_pending = False
    if "overwrite_week" not in st.session_state:
        st.session_state.overwrite_week = ""

    week_dates = {
        baslangic_tarihi + datetime.timedelta(days=offset)
        for offset in range(5)
    }
    existing_weeks = set()
    for saved_name, saved_frame in st.session_state.saved_schedules.items():
        if str(saved_name).strip().casefold() == hafta_adi.strip().casefold():
            existing_weeks.add(str(saved_name))
            continue
        if isinstance(saved_frame, pd.DataFrame) and "Tarih" in saved_frame.columns:
            saved_dates = pd.to_datetime(saved_frame["Tarih"], errors="coerce").dt.date
            if any(day in week_dates for day in saved_dates.dropna()):
                existing_weeks.add(str(saved_name))
    if not st.session_state.history.empty and {"Tarih/Hafta", "Tarih"}.issubset(st.session_state.history.columns):
        history_dates = pd.to_datetime(st.session_state.history["Tarih"], errors="coerce").dt.date
        matching_history = st.session_state.history[history_dates.isin(week_dates)]
        existing_weeks.update(matching_history["Tarih/Hafta"].dropna().astype(str).tolist())

    if generate_clicked:
        if existing_weeks:
            st.session_state.overwrite_pending = True
            st.session_state.overwrite_week = hafta_adi
        else:
            st.session_state.overwrite_pending = False

    overwrite_confirmed = False
    if st.session_state.overwrite_pending and st.session_state.overwrite_week == hafta_adi:
        st.warning(
            f"{', '.join(sorted(existing_weeks)) or hafta_adi} için kayıtlı bir nöbet listesi var. "
            "Devam ederseniz bu haftanın eski verileri yeni çizelgeyle değiştirilecek."
        )
        confirm_col, cancel_col = st.columns(2)
        with confirm_col:
            overwrite_confirmed = st.button("Evet, üzerine yaz", type="primary", key="confirm_overwrite_schedule")
        with cancel_col:
            if st.button("İptal", key="cancel_overwrite_schedule"):
                st.session_state.overwrite_pending = False
                st.session_state.overwrite_week = ""
                st.rerun()

    if (generate_clicked and not existing_weeks) or overwrite_confirmed:
        st.session_state.overwrite_pending = False
        st.session_state.overwrite_week = ""
        teachers_list = st.session_state.teachers_dict_list
        locs_df = st.session_state.locations
        holidays_df = st.session_state.holidays
        
        gunler = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
        
        holidays_df = normalized_holidays(holidays_df)
        tatil_gunleri = []
        for idx, g in enumerate(gunler):
            g_tarih = baslangic_tarihi + datetime.timedelta(days=idx)
            if not is_holiday(holidays_df, g_tarih).empty:
                tatil_gunleri.append(g)

        schedule_records = []
        active_gunler = [g for g in gunler if g not in tatil_gunleri]
        
        if not teachers_list:
            st.error("Nöbet dağıtımı yapılamadı: öğretmen listesi boş.")
            st.stop()

        # Normalize legacy values once, so Turkish casing and missing fields do not bypass rules.
        for teacher in teachers_list:
            teacher.setdefault("Sabit Yer", "Yok")
            teacher.setdefault("Hariç Tutulan Bölümler", [])
            teacher.setdefault("İzinli Günler", [])
            teacher.setdefault("Haftada 2 Nöbet İstiyor", False)

        # Build all slots first and solve the full week together. A greedy choice can
        # consume a teacher needed for a more restricted section later in the week.
        slots = []
        for gun in active_gunler:
            for _, loc_row in locs_df.iterrows():
                bolge = str(loc_row["Nöbet Yeri / Bölüm"]).strip()
                try:
                    raw_capacity = loc_row.get("Gerekli Öğretmen Sayısı", 1)
                    if pd.isna(raw_capacity) or str(raw_capacity).strip() == "":
                        capacity = 1
                    else:
                        numeric_capacity = float(raw_capacity)
                        if not numeric_capacity.is_integer() or numeric_capacity < 0:
                            raise ValueError
                        capacity = int(numeric_capacity)
                except (TypeError, ValueError):
                    st.error(f"'{bolge}' için gerekli öğretmen sayısı geçersiz.")
                    st.stop()
                active_days = loc_row.get("Aktif Günler", WEEKDAYS)
                if not isinstance(active_days, (list, tuple, set)):
                    active_days = []
                if gun in active_days:
                    slots.extend((gun, bolge) for _ in range(capacity))

        teacher_by_name = {t["Ad Soyad"]: t for t in teachers_list}
        teacher_names = list(teacher_by_name)
        # This preference influences who gets a second assignment; it is not a hard cap.
        # A hard cap of one made ordinary staffing needs impossible in many schools.
        max_weekly = {name: len(slots) for name in teacher_names}

        def eligible(name, slot):
            info = teacher_by_name[name]
            gun, bolge = slot
            izinler = {str(x).strip().casefold() for x in info["İzinli Günler"]}
            haricler = {str(x).strip().casefold() for x in info["Hariç Tutulan Bölümler"]}
            return (
                gun.casefold() not in izinler
                and bolge.casefold() not in haricler
                and info["Sabit Yer"] in ("Yok", bolge)
            )

        # Assign the most constrained slots first; backtracking preserves feasibility.
        slot_candidates = [
            [name for name in teacher_names if eligible(name, slot)]
            for slot in slots
        ]
        slot_order = sorted(range(len(slots)), key=lambda i: len(slot_candidates[i]))
        assignments = [None] * len(slots)
        counts = {name: 0 for name in teacher_names}
        # Avoid repeating a teacher's previous week's location. In the fallback,
        # allow it only for a teacher whose fixed location makes them required there.
        history = st.session_state.history.copy()
        prior_history = history[~history["Tarih/Hafta"].astype(str).isin(existing_weeks | {hafta_adi})] if "Tarih/Hafta" in history.columns else history
        if "Tarih" in prior_history.columns:
            prior_dates = pd.to_datetime(prior_history["Tarih"], errors="coerce").dt.date
            prior_history = prior_history.loc[~prior_dates.isin(week_dates)]
        previous_start = baslangic_tarihi - datetime.timedelta(days=7)
        previous_end = previous_start + datetime.timedelta(days=4)
        if "Tarih" in history.columns:
            history["_date"] = pd.to_datetime(history["Tarih"], errors="coerce").dt.date
            previous_rows = history[
                history["_date"].apply(
                    lambda d: pd.notna(d)
                    and isinstance(d, datetime.date)
                    and previous_start <= d <= previous_end
                )
            ]
        else:
            previous_rows = pd.DataFrame()
        previous_places = {}
        if not previous_rows.empty:
            for _, row in previous_rows.iterrows():
                previous_places.setdefault(row["Öğretmen"], set()).add(str(row["Bölüm/Yer"]).strip().casefold())
        prior_weekly_counts = prior_history.groupby(["Tarih/Hafta", "Öğretmen"]).size() if {"Tarih/Hafta", "Öğretmen"}.issubset(prior_history.columns) else pd.Series(dtype=int)
        historic_double_weeks = prior_weekly_counts[prior_weekly_counts > 1].reset_index().groupby("Öğretmen").size().to_dict() if not prior_weekly_counts.empty else {}
        all_place_counts = {}
        if not prior_history.empty and {"Öğretmen", "Bölüm/Yer"}.issubset(prior_history.columns):
            known_place_history = prior_history[prior_history["Bölüm/Yer"].isin(locs_df["Nöbet Yeri / Bölüm"].astype(str))]
            all_place_counts = known_place_history.groupby(["Öğretmen", "Bölüm/Yer"]).size().to_dict()
        any_double_preference = any(bool(info["Haftada 2 Nöbet İstiyor"]) for info in teacher_by_name.values())
        active_day_order = {day: index for index, day in enumerate(active_gunler)}

        node_budget = 1500000
        visited_nodes = [0]

        def try_solve(allow_mandatory_previous_place):
            counts.clear()
            counts.update({name: 0 for name in teacher_names})
            assignments[:] = [None] * len(slots)
            visited_nodes[0] = 0

            def solve(pos):
                visited_nodes[0] += 1
                if visited_nodes[0] > node_budget:
                    return False
                if pos == len(slot_order):
                    return True
                slot_idx = slot_order[pos]
                slot_gun, slot_place = slots[slot_idx]
                slot_day_index = active_day_order.get(slot_gun, -1)
                candidates = []
                for name in slot_candidates[slot_idx]:
                    if counts[name] >= max_weekly[name]:
                        continue
                    # A teacher may never receive two locations on the same day.
                    already_on_same_day = any(
                        assigned_name == name and slots[other_idx][0] == slot_gun
                        for other_idx, assigned_name in enumerate(assignments)
                        if assigned_name is not None
                    )
                    if already_on_same_day:
                        continue
                    # Holidays are removed from active_gunler, so duty days on
                    # either side of a holiday are treated as consecutive duties.
                    has_adjacent_day = any(
                        assigned_name == name
                        and abs(active_day_order.get(slots[other_idx][0], -99) - slot_day_index) == 1
                        for other_idx, assigned_name in enumerate(assignments)
                        if assigned_name is not None
                    )
                    if has_adjacent_day:
                        continue
                    repeats_previous_place = slot_place.casefold() in previous_places.get(name, set())
                    if repeats_previous_place and not (
                        allow_mandatory_previous_place
                        and teacher_by_name[name]["Sabit Yer"] == slot_place
                    ):
                        continue
                    candidates.append(name)
                random.shuffle(candidates)
                candidates.sort(key=lambda name: (
                    counts[name],
                    historic_double_weeks.get(name, 0) if counts[name] > 0 else 0,
                    (0 if teacher_by_name[name]["Haftada 2 Nöbet İstiyor"] else 1)
                    if any_double_preference and counts[name] > 0 else 0,
                    all_place_counts.get((name, slot_place), 0)
                ))
                for name in candidates:
                    counts[name] += 1
                    assignments[slot_idx] = name
                    if solve(pos + 1):
                        return True
                    assignments[slot_idx] = None
                    counts[name] -= 1
                return False
            return solve(0)

        solved_without_repeat = try_solve(allow_mandatory_previous_place=False)
        repeated_place_fallback = False
        fallback_solved = False
        if not solved_without_repeat:
            fallback_solved = try_solve(allow_mandatory_previous_place=True)
            repeated_place_fallback = fallback_solved and any(
                place.casefold() in previous_places.get(name, set())
                for (day, place), name in zip(slots, assignments)
            )

        if not solved_without_repeat and not fallback_solved:
            impossible = []
            for slot, candidates in zip(slots, slot_candidates):
                if not candidates:
                    impossible.append(f"{slot[0]} / {slot[1]} (uygun öğretmen: 0)")
            if impossible:
                detail = "; ".join(dict.fromkeys(impossible))
                st.error(f"Kurallara göre doldurulamayan görevler: {detail}. Hariç tutma, izin ve sabit yer ayarlarını gözden geçirin.")
            elif visited_nodes[0] > node_budget:
                st.error("Öğretmen, gün ve yer kurallarıyla uygun çizelge bulunamadı. Kısıtlar veya kontenjanlar mevcut öğretmen sayısına göre çok sıkı olabilir.")
            else:
                st.error("Tüm nöbet yerlerini mevcut öğretmen ve haftalık nöbet sınırlarıyla doldurmak mümkün değil. İzin, sabit yer, hariç tutma veya kontenjanları gözden geçirin.")
            st.stop()

        deputy_assignments, _deputy_adjacent_fallback = distribute_deputies(
            baslangic_tarihi,
            {
                day: 0 if day in tatil_gunleri else st.session_state.mudur_yardimcisi_sayilari.get(day, 1)
                for day in gunler
            },
            st.session_state.mudur_yardimcilari,
            st.session_state.mudur_yardimcisi_gecmisi,
            active_gunler,
        )
        st.session_state.mudur_yardimcisi_gecmisi[hafta_adi] = {
            "start_date": baslangic_tarihi.isoformat(),
            "assignments": deputy_assignments,
        }
        ogretmen_haftalik_sayac = dict(counts)
        for (gun, bolge), secilen_ogretmen in zip(slots, assignments):
            day_index = gunler.index(gun)
            schedule_records.append({
                "Tarih/Hafta": hafta_adi,
                "Tarih": baslangic_tarihi + datetime.timedelta(days=day_index),
                "Gün": gun,
                "Bölüm/Yer": bolge,
                "Öğretmen": secilen_ogretmen,
                "Çift Nöbetçi mi": counts[secilen_ogretmen] > 1,
                "Kaynak": "Oluşturulan"
            })
        
        sched_df = pd.DataFrame(schedule_records)
        st.session_state.current_schedule = sched_df
        
        if not st.session_state.history.empty:
            history_dates = pd.to_datetime(st.session_state.history["Tarih"], errors="coerce").dt.date
            replace_rows = st.session_state.history["Tarih/Hafta"].astype(str).isin(existing_weeks | {hafta_adi})
            replace_rows |= history_dates.isin(week_dates)
            st.session_state.history = st.session_state.history.loc[~replace_rows].copy()
        for old_name in existing_weeks:
            if old_name != hafta_adi:
                st.session_state.saved_schedules.pop(old_name, None)
                st.session_state.mudur_yardimcisi_gecmisi.pop(old_name, None)
        
        st.session_state.history = pd.concat([st.session_state.history, sched_df], ignore_index=True)
        st.session_state.saved_schedules[hafta_adi] = sched_df.copy()
        save_app_data()
        st.success(f"✨ '{hafta_adi}' dönemi çizelgesi başarıyla oluşturuldu!")
        if repeated_place_fallback:
            st.warning("Öğretmen sayısı ve diğer kurallar nedeniyle bazı öğretmenlere geçen haftaki nöbet yeri yeniden verildi.")

    if st.session_state.current_schedule is not None:
        st.subheader("Oluşturulan Çizelge Önizlemesi")
        
        preview_df = st.session_state.current_schedule.copy()
        schedule_name_for_export = str(preview_df["Tarih/Hafta"].dropna().iloc[0]) if "Tarih/Hafta" in preview_df and not preview_df["Tarih/Hafta"].dropna().empty else hafta_adi
        schedule_start_for_export = baslangic_tarihi
        if "Tarih" in preview_df.columns:
            preview_dates = pd.to_datetime(preview_df["Tarih"], errors="coerce").dropna()
            if not preview_dates.empty:
                first_date = preview_dates.min().date()
                schedule_start_for_export = first_date - datetime.timedelta(days=first_date.weekday())
        current_deputy_assignments = st.session_state.mudur_yardimcisi_gecmisi.get(schedule_name_for_export, {}).get("assignments", {})
        hist_df = st.session_state.history
        
        def format_teacher_name_with_history(ogretmen_adi):
            toplam_nobet = len(hist_df[hist_df["Öğretmen"] == ogretmen_adi])
            teacher_history = hist_df[hist_df["Öğretmen"] == ogretmen_adi]
            cift_nobet = int((teacher_history.groupby("Tarih/Hafta").size() > 1).sum()) if "Tarih/Hafta" in teacher_history.columns else 0
            return f"{ogretmen_adi} (Toplam Nöbet: {toplam_nobet} ; Çift Nöbet Haftası: {cift_nobet})"

        preview_df["Öğretmen"] = preview_df["Öğretmen"].apply(format_teacher_name_with_history)

        st.markdown("""
        <style>
        [data-testid="stHorizontalBlock"]:has([data-testid="stDownloadButton"]) {
            column-gap: 0 !important;
            gap: 0 !important;
        }
        [data-testid="stHorizontalBlock"]:has([data-testid="stDownloadButton"]) [data-testid="stDownloadButton"] button {
            border-radius: 0 !important;
        }
        </style>
        """, unsafe_allow_html=True)
        col_spacer, col_excel, col_pdf, col_img = st.columns([6, 2, 2, 2], gap="small")
        with col_excel:
            xlsx_data = create_custom_excel(
                st.session_state.current_schedule, schedule_name_for_export,
                current_deputy_assignments, schedule_start_for_export,
                st.session_state.school_info["okul_adi"], st.session_state.school_info["mudur_adi"],
                st.session_state.school_info["not_metni"], st.session_state.holidays,
                st.session_state.locations, st.session_state.excel_sections,
                st.session_state.excel_headings
            )
            st.download_button(
                label="📥 Excel Olarak İndir",
                data=xlsx_data,
                file_name=f"{schedule_name_for_export}_Nobet_Cizelgesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        with col_pdf:
            pdf_data = create_pdf(
                st.session_state.current_schedule, schedule_name_for_export,
                current_deputy_assignments,
                st.session_state.school_info["okul_adi"], st.session_state.school_info["mudur_adi"],
                st.session_state.school_info["not_metni"], schedule_start_for_export, st.session_state.holidays,
                st.session_state.locations
            )
            st.download_button(
                label="📄 PDF Olarak İndir",
                data=pdf_data,
                file_name=f"{schedule_name_for_export}_Nobet_Cizelgesi.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )
        with col_img:
            image_data = create_schedule_image(
                st.session_state.current_schedule, schedule_name_for_export,
                current_deputy_assignments, schedule_start_for_export,
                st.session_state.school_info["okul_adi"], st.session_state.school_info["mudur_adi"],
                st.session_state.holidays, st.session_state.locations, st.session_state.excel_sections
            )
            st.download_button(
                label="🖼️ Resim Olarak İndir",
                data=image_data,
                file_name=f"{schedule_name_for_export}_Nobet_Cizelgesi.png",
                mime="image/png",
                type="secondary",
                use_container_width=True
            )
        st.dataframe(preview_df, use_container_width=True)


# ================= 2. SEKME: ÖĞRETMEN YÖNETİMİ =================
elif main_menu == "👥 Öğretmen Yönetimi":
    st.header("👥 Öğretmen Yönetimi ve Kısıtlar Paneli")
    save_top_left, save_top_center, save_top_right = st.columns([1, 2, 1])
    with save_top_center:
        if st.button("💾 Öğretmen ayarlarını kaydet", key="save_teacher_settings_top", type="primary", use_container_width=True):
            save_app_data()
            st.success("Öğretmen listesi ve kısıt ayarları kaydedildi.")

    tum_bolgeler = st.session_state.locations["Nöbet Yeri / Bölüm"].tolist()
    yer_listesi = ["Yok"] + tum_bolgeler
    gun_listesi = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]

    # Rerun sonrası mesaj göster
    if "_tab2_mesaj" in st.session_state:
        st.success(st.session_state.pop("_tab2_mesaj"))

    # --- TOPLU İŞLEMLER PANELİ (FORM YOK, NORMAL BUTON) ---
    with st.expander("⚡ Toplu Kısıtlama Uygulama Paneli (Tüm Öğretmenler İçin)", expanded=True):
        st.markdown("Aşağıdan bir bölüm seçerek **tüm öğretmenlere tek tıkla** hariç tutma kısıtı ekleyebilir veya kaldırabilirsiniz.")

        select_left, select_center, select_right = st.columns([1, 2, 1])
        with select_center:
            toplu_secilen_bolge = st.selectbox(
                "",
                tum_bolgeler,
                key="toplu_bolge_secim_input",
                label_visibility="collapsed"
            )
        button_left, button_add, button_clear, button_right = st.columns([1, 2, 2, 1])
        with button_add:
            if st.button("➕ Tüm Öğretmenlerin Hariç Tutulanlarına Ekle",
                         key="btn_toplu_ekle",
                         type="primary",
                         use_container_width=True):
                # Update both the source records and the keyed multiselect state.
                # Streamlit keeps keyed widget values across reruns, so changing only
                # teachers_dict_list would leave the visible controls (and next edits)
                # on their previous values.
                yeni_liste = []
                eklenen_sayi = 0
                for idx, t in enumerate(st.session_state.teachers_dict_list):
                    yeni_t = dict(t)
                    haric = list(yeni_t.get("Hariç Tutulan Bölümler", []))
                    haric = [h for h in haric if h in tum_bolgeler]
                    if toplu_secilen_bolge not in haric:
                        haric.append(toplu_secilen_bolge)
                        eklenen_sayi += 1
                    yeni_t["Hariç Tutulan Bölümler"] = haric
                    yeni_liste.append(yeni_t)
                    st.session_state[f"haric_{t['Ad Soyad']}"] = haric
                st.session_state.teachers_dict_list = yeni_liste
                st.session_state["_tab2_mesaj"] = f"✅ '{toplu_secilen_bolge}' {len(yeni_liste)} öğretmene uygulandı ({eklenen_sayi} yeni kısıt eklendi)."
                save_app_data()
                st.rerun()

        with button_clear:
            if st.button("🧹 Tüm Öğretmenlerin Hariç Tutulanlarını Temizle",
                         key="btn_toplu_temizle",
                         use_container_width=True):
                yeni_liste = []
                for idx, t in enumerate(st.session_state.teachers_dict_list):
                    yeni_t = dict(t)
                    yeni_t["Hariç Tutulan Bölümler"] = []
                    yeni_liste.append(yeni_t)
                    st.session_state[f"haric_{t['Ad Soyad']}"] = []
                st.session_state.teachers_dict_list = yeni_liste
                st.session_state["_tab2_mesaj"] = "✅ Tüm öğretmenlerin hariç tutulan bölümleri temizlendi!"
                save_app_data()
                st.rerun()

    st.markdown("---")

    # --- Yeni Öğretmen Ekleme ---
    with st.expander("➕ Yeni Öğretmen Ekle"):
        col_y1, col_y2 = st.columns([4, 1])
        with col_y1:
            yeni_ogretmen_adi = st.text_input("Öğretmen Adı Soyadı", key="yeni_ogretmen_input", label_visibility="collapsed", placeholder="Öğretmen Adı Soyadı")
        with col_y2:
            if st.button("Ekle", key="btn_ogretmen_ekle", use_container_width=True):
                if yeni_ogretmen_adi and yeni_ogretmen_adi.strip():
                    yeni_isim = yeni_ogretmen_adi.strip().upper()
                    mevcut_isimler = [t["Ad Soyad"] for t in st.session_state.teachers_dict_list]
                    if yeni_isim not in mevcut_isimler:
                        yeni_liste = list(st.session_state.teachers_dict_list)
                        yeni_liste.append({
                            "Ad Soyad": yeni_isim,
                            "Sabit Yer": "Yok",
                            "Hariç Tutulan Bölümler": [],
                            "İzinli Günler": [],
                            "Haftada 2 Nöbet İstiyor": False
                        })
                        st.session_state.teachers_dict_list = yeni_liste
                        st.session_state["_tab2_mesaj"] = f"✅ {yeni_isim} başarıyla eklendi!"
                        save_app_data()
                        st.rerun()
                    else:
                        st.warning(f"{yeni_isim} zaten listede var.")

    with st.expander("➕ Toplu öğretmen ekle"):
        bulk_teacher_text = st.text_area(
            "İsimleri satır satır veya virgülle ayırarak yazın",
            key="bulk_teacher_names",
            placeholder="ÖRNEK ÖĞRETMEN\nDİĞER ÖĞRETMEN"
        )
        if st.button("Toplu eklemeyi uygula", key="bulk_teacher_add", use_container_width=True):
            raw_names = re.split(r"[\n,;]+", bulk_teacher_text)
            existing_names = {str(t.get("Ad Soyad", "")).strip().casefold() for t in st.session_state.teachers_dict_list}
            added_names = []
            for raw_name in raw_names:
                name = raw_name.strip().upper()
                if name and name.casefold() not in existing_names:
                    st.session_state.teachers_dict_list.append({
                        "Ad Soyad": name,
                        "Sabit Yer": "Yok",
                        "Hariç Tutulan Bölümler": [],
                        "İzinli Günler": [],
                        "Haftada 2 Nöbet İstiyor": False,
                    })
                    existing_names.add(name.casefold())
                    added_names.append(name)
            save_app_data()
            st.success(f"{len(added_names)} öğretmen eklendi.")
            st.rerun()

    st.markdown("---")
    st.subheader("📋 Öğretmen Listesi ve Kısıt Ayarları")
    st.caption("Nöbet sayaçları arşivdeki çizelgelere göre hesaplanır. Listeyi ada veya nöbet sayılarına göre sıralayabilirsiniz.")

    if "confirm_delete_all_teachers" not in st.session_state:
        st.session_state.confirm_delete_all_teachers = False
    if st.button("🗑️ Tüm öğretmenleri sil", key="delete_all_teachers_start", use_container_width=True):
        st.session_state.confirm_delete_all_teachers = True
    if st.session_state.confirm_delete_all_teachers:
        st.warning("Tüm öğretmenler listesinden silinecek. Bu işlem geri alınamaz.")
        del_col, keep_col = st.columns(2)
        with del_col:
            if st.button("Evet, tümünü sil", key="delete_all_teachers_confirm", type="primary", use_container_width=True):
                for teacher in st.session_state.teachers_dict_list:
                    teacher_name = teacher.get("Ad Soyad", "")
                    for prefix in ("sabit_", "haric_", "izin_", "cift_"):
                        st.session_state.pop(f"{prefix}{teacher_name}", None)
                st.session_state.teachers_dict_list = []
                st.session_state.confirm_delete_all_teachers = False
                save_app_data()
                st.rerun()
        with keep_col:
            if st.button("İptal", key="delete_all_teachers_cancel", use_container_width=True):
                st.session_state.confirm_delete_all_teachers = False
                st.rerun()

    sort_mode = st.selectbox(
        "Öğretmenleri sırala",
        ["Ada göre (A-Z)", "Toplam nöbet sayısına göre (çoktan aza)", "Çift nöbet sayısına göre (çoktan aza)"],
        key="teacher_sort_mode"
    )

    # --- Öğretmenler Listesi (Normal widget'lar, form yok) ---
    silinecek_indexler = []
    degisiklik_var = False

    # Anlık snapshot al
    mevcut_liste = st.session_state.teachers_dict_list

    hist_for_counts = st.session_state.history
    st.metric("Toplam öğretmen sayısı", len(mevcut_liste))
    total_duty_counts = {}
    double_duty_counts = {}
    if not hist_for_counts.empty and "Öğretmen" in hist_for_counts.columns:
        total_duty_counts = hist_for_counts.groupby("Öğretmen").size().to_dict()
        if "Tarih/Hafta" in hist_for_counts.columns:
            double_weeks = hist_for_counts.groupby(["Tarih/Hafta", "Öğretmen"]).size()
            double_duty_counts = double_weeks[double_weeks > 1].reset_index().groupby("Öğretmen").size().to_dict()

    indexed_teachers = list(enumerate(mevcut_liste))
    if sort_mode == "Toplam nöbet sayısına göre (çoktan aza)":
        indexed_teachers.sort(key=lambda pair: (-total_duty_counts.get(pair[1]["Ad Soyad"], 0), pair[1]["Ad Soyad"].casefold()))
    elif sort_mode == "Çift nöbet sayısına göre (çoktan aza)":
        indexed_teachers.sort(key=lambda pair: (-double_duty_counts.get(pair[1]["Ad Soyad"], 0), pair[1]["Ad Soyad"].casefold()))
    else:
        indexed_teachers.sort(key=lambda pair: pair[1]["Ad Soyad"].casefold())

    for idx, t_data in indexed_teachers:
        t_data.setdefault("Sabit Yer", "Yok")
        t_data.setdefault("Hariç Tutulan Bölümler", [])
        t_data.setdefault("İzinli Günler", [])
        t_data.setdefault("Haftada 2 Nöbet İstiyor", False)
        teacher_name = t_data["Ad Soyad"]
        total_duties = total_duty_counts.get(teacher_name, 0)
        double_weeks_count = double_duty_counts.get(teacher_name, 0)
        with st.expander(f"{teacher_name}  ·  Toplam nöbet: {total_duties}  ·  Çift nöbet haftası: {double_weeks_count}"):
            st.caption("Bir haftadaki çift nöbet, sayaçta bir hafta olarak sayılır.")

            curr_sabit = t_data["Sabit Yer"]
            sabit_idx = yer_listesi.index(curr_sabit) if curr_sabit in yer_listesi else 0
            yeni_sabit = st.selectbox(
                "Sabit nöbet yeri", yer_listesi, index=sabit_idx,
                key=f"sabit_{teacher_name}"
            )
            if yeni_sabit != t_data["Sabit Yer"]:
                t_data["Sabit Yer"] = yeni_sabit
                degisiklik_var = True

            curr_haric = [h for h in t_data["Hariç Tutulan Bölümler"] if h in tum_bolgeler]
            yeni_haric = st.multiselect(
                "Hariç tutulacak yerler", tum_bolgeler, default=curr_haric,
                key=f"haric_{teacher_name}",
                placeholder="Hariç bölümler..."
            )
            if yeni_haric != t_data["Hariç Tutulan Bölümler"]:
                t_data["Hariç Tutulan Bölümler"] = yeni_haric
                degisiklik_var = True

            curr_izin = [i for i in t_data["İzinli Günler"] if i in gun_listesi]
            yeni_izin = st.multiselect(
                "İzinli günler", gun_listesi, default=curr_izin,
                key=f"izin_{teacher_name}",
                placeholder="İzinli günler..."
            )
            if yeni_izin != t_data["İzinli Günler"]:
                t_data["İzinli Günler"] = yeni_izin
                degisiklik_var = True

            yeni_cift = st.checkbox(
                "Gerekirse haftada 2 nöbet alabilir", value=t_data["Haftada 2 Nöbet İstiyor"],
                key=f"cift_{teacher_name}"
            )
            if yeni_cift != t_data["Haftada 2 Nöbet İstiyor"]:
                t_data["Haftada 2 Nöbet İstiyor"] = yeni_cift
                degisiklik_var = True

            if st.button("🗑️ Öğretmeni sil", key=f"del_t_{teacher_name}", use_container_width=True):
                silinecek_indexler.append(idx)

    # Silme işlemi
    if silinecek_indexler:
        yeni_liste = [
            t for i, t in enumerate(st.session_state.teachers_dict_list)
            if i not in silinecek_indexler
        ]
        st.session_state.teachers_dict_list = yeni_liste
        st.session_state["_tab2_mesaj"] = f"🗑️ {len(silinecek_indexler)} öğretmen silindi!"
        save_app_data()
        st.rerun()

    # Değişiklik varsa session_state'i yeniden ata
    if degisiklik_var:
        st.session_state.teachers_dict_list = list(st.session_state.teachers_dict_list)
        save_app_data()
# ================= 3. SEKME: NÖBET YERLERİ VE KONTENJAN =================
elif main_menu == "📍 Nöbet Yerleri ve Kontenjan":
    st.header("Nöbet Yerleri ve Kişi Sayısı Yönetimi")
    st.caption("Her yer için aktif günleri seçin. Seçilmeyen günlerde bu yerde nöbet görevi oluşturulmaz.")
    excel_headers = st.session_state.excel_sections["Excel Başlığı"].astype(str).tolist()

    edited_locs = st.data_editor(
        st.session_state.locations,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Nöbet Yeri / Bölüm": st.column_config.TextColumn("Nöbet Yeri / Bölüm", required=True),
            "Gerekli Öğretmen Sayısı": st.column_config.NumberColumn("Gerekli Öğretmen Sayısı", min_value=0, step=1, default=1),
            "Aktif Günler": st.column_config.MultiselectColumn("Aktif Günler", options=WEEKDAYS),
            "Excel Bölümü": st.column_config.SelectboxColumn("Excel Bölümü", options=excel_headers),
            "Excel'de Yer Başlığını Göster": st.column_config.CheckboxColumn("Excel'de yer adını göster", default=True)
        }
    )
    if "Gerekli Öğretmen Sayısı" in edited_locs.columns:
        edited_locs["Gerekli Öğretmen Sayısı"] = pd.to_numeric(
            edited_locs["Gerekli Öğretmen Sayısı"], errors="coerce"
        ).fillna(1).astype(int)
    edited_locs["Excel Bölümü"] = edited_locs.apply(
        lambda row: row["Excel Bölümü"]
        if isinstance(row.get("Excel Bölümü"), str) and row["Excel Bölümü"] in excel_headers
        else guess_excel_section(row.get("Nöbet Yeri / Bölüm", ""), excel_headers),
        axis=1
    )
    st.session_state.locations = edited_locs.reset_index(drop=True)
    st.caption("Aynı Excel başlığı altındaki yerleri seçip yukarı/aşağı taşıyarak sıralayın. Yer adını gizlerseniz Excel'de yalnızca öğretmen adları görünür.")
    order_section = st.selectbox("Sıralanacak Excel başlığı", excel_headers, key="location_order_section")
    section_positions = [
        idx for idx, value in st.session_state.locations["Excel Bölümü"].items()
        if str(value).strip() == order_section
    ]
    if section_positions:
        selected_location = st.selectbox(
            "Yer", [str(st.session_state.locations.at[idx, "Nöbet Yeri / Bölüm"]) for idx in section_positions],
            key="location_order_item"
        )
        selected_idx = next(
            idx for idx in section_positions
            if str(st.session_state.locations.at[idx, "Nöbet Yeri / Bölüm"]) == selected_location
        )
        move_up, move_down = st.columns(2)
        with move_up:
            if st.button("⬆️ Yukarı taşı", key="location_order_up", disabled=section_positions.index(selected_idx) == 0):
                ordered = st.session_state.locations.reset_index(drop=True)
                positions = [idx for idx, value in ordered["Excel Bölümü"].items() if str(value).strip() == order_section]
                pos = positions.index(selected_idx)
                swap_pos = positions[pos - 1]
                row_values = ordered.iloc[[selected_idx, swap_pos]].to_numpy().copy()
                ordered.iloc[selected_idx] = row_values[1]
                ordered.iloc[swap_pos] = row_values[0]
                st.session_state.locations = ordered
                save_app_data()
                st.rerun()
        with move_down:
            if st.button("⬇️ Aşağı taşı", key="location_order_down", disabled=section_positions.index(selected_idx) == len(section_positions) - 1):
                ordered = st.session_state.locations.reset_index(drop=True)
                positions = [idx for idx, value in ordered["Excel Bölümü"].items() if str(value).strip() == order_section]
                pos = positions.index(selected_idx)
                swap_pos = positions[pos + 1]
                row_values = ordered.iloc[[selected_idx, swap_pos]].to_numpy().copy()
                ordered.iloc[selected_idx] = row_values[1]
                ordered.iloc[swap_pos] = row_values[0]
                st.session_state.locations = ordered
                save_app_data()
                st.rerun()

    st.markdown("---")
    st.subheader("Excel çizelgesi başlıkları")
    st.caption("Başlıkları değiştirebilir, yeni başlık ekleyebilir veya satır silebilirsiniz. Her nöbet yerini yukarıdaki Excel bölümüne bağlayın.")
    old_excel_headers = excel_headers.copy()
    edited_sections = st.data_editor(
        st.session_state.excel_sections,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={"Excel Başlığı": st.column_config.TextColumn("Excel Sütun Başlığı", required=True)},
        key="excel_sections_editor"
    )
    edited_sections["Excel Başlığı"] = edited_sections["Excel Başlığı"].fillna("").astype(str).str.strip()
    edited_sections = edited_sections[edited_sections["Excel Başlığı"] != ""].drop_duplicates(subset=["Excel Başlığı"])
    if edited_sections.empty:
        edited_sections = pd.DataFrame(DEFAULT_EXCEL_SECTIONS)
    st.session_state.excel_sections = edited_sections.reset_index(drop=True)
    new_excel_headers = st.session_state.excel_sections["Excel Başlığı"].astype(str).tolist()
    renamed_headers = dict(zip(old_excel_headers, new_excel_headers))
    st.session_state.locations["Excel Bölümü"] = st.session_state.locations.apply(
        lambda row: renamed_headers.get(str(row.get("Excel Bölümü", "")), str(row.get("Excel Bölümü", "")))
        if renamed_headers.get(str(row.get("Excel Bölümü", "")), str(row.get("Excel Bölümü", ""))) in new_excel_headers
        else guess_excel_section(row.get("Nöbet Yeri / Bölüm", ""), new_excel_headers),
        axis=1
    )
    heading_col1, heading_col2 = st.columns(2)
    with heading_col1:
        st.session_state.excel_headings["Gün Başlığı"] = st.text_input(
            "Gün sütununun başlığı",
            value=st.session_state.excel_headings.get("Gün Başlığı", "NÖBET GÜNÜ"),
            key="excel_day_heading"
        )
    with heading_col2:
        st.session_state.excel_headings["Müdür Yardımcısı Başlığı"] = st.text_input(
            "Müdür yardımcısı sütununun başlığı",
            value=st.session_state.excel_headings.get("Müdür Yardımcısı Başlığı", "MÜDÜR YRD"),
            key="excel_deputy_heading"
        )
# ================= 4. SEKME: OKUL, TATİL VE AYARLAR =================
elif main_menu == "⚙️ Okul, Tatil ve Ayarlar":
    st.header("Okul, Yönetici ve Tatil Günü Ayarları")
    okul_adi_input = st.text_input("Okul Adı", value=st.session_state.school_info["okul_adi"])
    mudur_adi_input = st.text_input("Okul Müdürü Ad Soyad", value=st.session_state.school_info["mudur_adi"])
    not_metni_input = st.text_area("Çizelge Altı Notu (Açıklama Metni)", value=st.session_state.school_info["not_metni"], height=100)
    
    if st.button("Bilgileri Kaydet"):
        st.session_state.school_info["okul_adi"] = okul_adi_input
        st.session_state.school_info["mudur_adi"] = mudur_adi_input
        st.session_state.school_info["not_metni"] = not_metni_input
        save_app_data()
        st.success("Okul bilgileri güncellendi!")

    st.markdown("---")
    st.subheader("📅 Okulun Tatil Olduğu Günler (Takvimden Seçim)")
    
    # Date input provides a predictable calendar picker; the editor below is only for
    # reviewing/removing saved days and avoids fragile blank rows in data_editor.
    holiday_date = st.date_input("Tatil tarihini takvimden seçin", value=None, key="holiday_date_picker")
    holiday_note = st.text_input("Tatil açıklaması (isteğe bağlı)", key="holiday_note_input", placeholder="Örn. Resmî tatil")
    if st.button("📅 Seçili tarihi tatil olarak kaydet", key="holiday_add", type="primary"):
        if holiday_date is None:
            st.warning("Önce takvimden bir tarih seçin.")
        else:
            holidays = normalized_holidays(st.session_state.holidays)
            holidays = holidays[holidays["Tarih"] != holiday_date]
            holidays.loc[len(holidays)] = [holiday_date, holiday_note.strip()]
            st.session_state.holidays = normalized_holidays(holidays)
            st.session_state.pop("holiday_date_picker", None)
            save_app_data()
            st.success(f"{holiday_date.strftime('%d.%m.%Y')} tatil günlerine kaydedildi.")
            st.rerun()

    holidays_view = normalized_holidays(st.session_state.holidays)
    if not holidays_view.empty:
        st.caption("Kaydedilmiş tatil günleri. Açıklamayı düzenleyebilir veya satırları seçip silebilirsiniz.")
        edited_holidays = st.data_editor(
            holidays_view,
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Tarih": st.column_config.DateColumn("Tatil Tarihi", format="DD.MM.YYYY", disabled=True),
                "Tatil Açıklaması": st.column_config.TextColumn("Tatil Açıklaması")
            }
        )
        if st.button("💾 Tatil açıklamalarını kaydet", key="holiday_save"):
            st.session_state.holidays = normalized_holidays(edited_holidays)
            save_app_data()
            st.success("Tatil kayıtları güncellendi.")
            st.rerun()
        to_delete = st.multiselect("Silinecek tatil tarihleri", holidays_view["Tarih"].tolist(), format_func=lambda d: d.strftime("%d.%m.%Y"), key="holiday_delete_select")
        if st.button("🗑️ Seçili tatil günlerini sil", key="holiday_delete", disabled=not to_delete):
            st.session_state.holidays = holidays_view[~holidays_view["Tarih"].isin(to_delete)].reset_index(drop=True)
            save_app_data()
            st.success(f"{len(to_delete)} tatil kaydı silindi.")
            st.rerun()
    else:
        st.info("Henüz tatil günü eklenmedi.")

    st.markdown("---")
    st.subheader("Müdür Yardımcıları Listesi")
    my_input = st.text_input("Yeni Müdür Yardımcısı Ekle", key="yeni_my_input")
    if st.button("Müdür Yardımcısı Ekle", key="btn_my_ekle"):
        if my_input and my_input not in st.session_state.mudur_yardimcilari:
            st.session_state.mudur_yardimcilari.append(my_input)
            save_app_data()
            st.success(f"{my_input} eklendi.")
            st.rerun()

    for idx, my in enumerate(st.session_state.mudur_yardimcilari):
        col_a, col_b = st.columns([4, 1])
        with col_a:
            st.write(f"- {my}")
        with col_b:
            if st.button("Sil", key=f"del_my_{idx}"):
                st.session_state.mudur_yardimcilari.pop(idx)
                save_app_data()
                st.rerun()

    st.subheader("Günlük müdür yardımcısı nöbet sayısı")
    st.caption("Aynı müdür yardımcısı mümkün olduğunca arka arkaya günlerde görevlendirilmez; dağıtım önceki haftanın sırasını izler.")
    count_cols = st.columns(2)
    for day_index, day in enumerate(WEEKDAYS):
        with count_cols[day_index % 2]:
            current_count = int(st.session_state.mudur_yardimcisi_sayilari.get(day, 1))
            current_count = max(0, min(current_count, len(st.session_state.mudur_yardimcilari)))
            st.session_state.mudur_yardimcisi_sayilari[day] = st.number_input(
                f"{day} günü kişi sayısı",
                min_value=0,
                max_value=len(st.session_state.mudur_yardimcilari),
                value=current_count,
                step=1,
                key=f"deputy_count_{day}"
            )


# ================= 5. SEKME: ESKİ LİSTELER & ARŞİV =================
elif main_menu == "📂 Eski Listeler & Arşiv":
    st.header("Eski Listeler ve Nöbet Arşivi")
    st.markdown("Geçmiş Word veya bu uygulamanın oluşturduğu Excel çizelgelerini yükleyin. Excel aktarımı tarih, gün ve nöbet yerlerini arşive işler.")
    
    uploaded_files = st.file_uploader("Eski nöbet listelerini seçin (Word veya oluşturulan Excel)", type=["docx", "xlsx"], accept_multiple_files=True)
    if uploaded_files and st.button("🔄 Seçili dosyaları yeniden tara", key="rescan_archives"):
        for file in uploaded_files:
            st.session_state.processed_uploads.discard(hashlib.sha256(file.getvalue()).hexdigest())
    
    if uploaded_files:
        known_names = [t["Ad Soyad"].strip().upper() for t in st.session_state.teachers_dict_list if t.get("Ad Soyad")]
        if "Öğretmen" in st.session_state.history.columns:
            known_names.extend(st.session_state.history["Öğretmen"].dropna().astype(str).str.strip().str.upper().tolist())
        known_names.append("AYŞENUR AKBOĞA")  # Keep removed staff matchable in historical files.
        valid_teachers = sorted(set(known_names), key=len, reverse=True)
        known_locations = {
            search_normalize(str(name)): str(name)
            for name in st.session_state.locations["Nöbet Yeri / Bölüm"].dropna().tolist()
        }
        
        for uploaded_file in uploaded_files:
            try:
                upload_signature = hashlib.sha256(uploaded_file.getvalue()).hexdigest()
                if upload_signature in st.session_state.processed_uploads:
                    continue
                parsed_records = []
                hafta_adi_str = os.path.splitext(uploaded_file.name)[0].upper()
                teacher_occurrences = []
                deputy_occurrences = {day: [] for day in WEEKDAYS}
                imported_start_date = None
                assistant_names = list(dict.fromkeys(
                    list(st.session_state.mudur_yardimcilari)
                    + [name for entry in st.session_state.mudur_yardimcisi_gecmisi.values()
                       for day_names in entry.get("assignments", {}).values() for name in day_names]
                ))
                if uploaded_file.name.lower().endswith(".docx"):
                    doc = Document(uploaded_file)
                    current_gun = "BİLİNMEYEN GÜN"
                    current_place = "ESKİ ARŞİV"
                    day_lookup = {search_normalize(day): day for day in WEEKDAYS}

                    def day_in_text(text):
                        normalized = search_normalize(text)
                        for day_key, day_value in day_lookup.items():
                            if re.search(rf"\b{re.escape(day_key)}\b", normalized):
                                return day_value
                        return None

                    def add_word_teachers(text, day_value, place_value="ESKİ ARŞİV"):
                        normalized_line = f" {search_normalize(text)} "
                        for assistant in assistant_names:
                            normalized_assistant = search_normalize(assistant)
                            if day_value in deputy_occurrences and normalized_assistant and f" {normalized_assistant} " in normalized_line:
                                deputy_occurrences[day_value].append(assistant)
                        for teacher in valid_teachers:
                            normalized_teacher = search_normalize(teacher)
                            if normalized_teacher and f" {normalized_teacher} " in normalized_line:
                                teacher_occurrences.append({
                                    "Tarih/Hafta": hafta_adi_str,
                                    "Tarih": pd.NaT,
                                    "Gün": day_value,
                                    "Bölüm/Yer": place_value,
                                    "Öğretmen": teacher,
                                })

                    def location_in_text(text):
                        normalized = search_normalize(text)
                        matches = [
                            (len(normalized_name), place) for normalized_name, place in known_locations.items()
                            if normalized_name and f" {normalized_name} " in f" {normalized} "
                        ]
                        return max(matches)[1] if matches else None

                    for paragraph in doc.paragraphs:
                        text = paragraph.text.strip()
                        if not text:
                            continue
                        found_day = day_in_text(text)
                        if found_day:
                            current_gun = found_day
                        found_place = location_in_text(text)
                        if found_place:
                            current_place = found_place
                        add_word_teachers(text, current_gun, current_place)

                    for table in doc.tables:
                        column_days = {}
                        column_places = {}
                        for row in table.rows:
                            cells = [cell.text.strip() for cell in row.cells]
                            detected = {index: day_in_text(value) for index, value in enumerate(cells)}
                            detected = {index: day for index, day in detected.items() if day}
                            detected_places = {index: location_in_text(value) for index, value in enumerate(cells)}
                            detected_places = {index: place for index, place in detected_places.items() if place}
                            column_places.update(detected_places)
                            if detected:
                                column_days.update(detected)
                                current_gun = next(iter(detected.values()))
                            for index, text in enumerate(cells):
                                day_value = day_in_text(text) or column_days.get(index, current_gun)
                                if day_value:
                                    current_gun = day_value
                                place_value = detected_places.get(index, column_places.get(index, location_in_text(text) or "ESKİ ARŞİV"))
                                add_word_teachers(text, day_value, place_value)
                else:
                    sheet = openpyxl.load_workbook(uploaded_file, data_only=True, read_only=True).active
                    if sheet["A2"].value:
                        hafta_adi_str = str(sheet["A2"].value).replace(" NÖBET ÇİZELGESİ", "").strip().upper()
                    excel_place_map = {2: "BAHÇE", 3: "SINIF", 4: "ANASINIFI KAPI", 5: "EK BİNA"}
                    floor_locations = {
                        (3, "1"): "SINIF 1.KAT", (3, "2"): "SINIF 2.KAT", (3, "3"): "SINIF 3.KAT",
                        (5, "1"): "EK BİNA 1.KAT", (5, "2"): "EK BİNA 2.KAT",
                    }
                    known_locations = {
                        search_normalize(str(name)): str(name)
                        for name in st.session_state.locations["Nöbet Yeri / Bölüm"].dropna().tolist()
                    }
                    configured_deputy_heading = search_normalize(
                        st.session_state.excel_headings.get("Müdür Yardımcısı Başlığı", "MÜDÜR YRD")
                    )
                    deputy_col_idx = next(
                        (
                            col for col in range(1, sheet.max_column + 1)
                            if search_normalize(sheet.cell(3, col).value or "") == configured_deputy_heading
                            or "mudur yrd" in search_normalize(sheet.cell(3, col).value or "")
                        ),
                        sheet.max_column,
                    )
                    for row_idx in range(4, 9):
                        day_cell = str(sheet.cell(row_idx, 1).value or "")
                        day_match = re.search(r"(PAZARTESİ|SALI|ÇARŞAMBA|PERŞEMBE|CUMA)", day_cell.upper())
                        if not day_match:
                            continue
                        day_name = day_match.group(1)
                        date_match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", day_cell)
                        duty_date = datetime.date(int(date_match.group(3)), int(date_match.group(2)), int(date_match.group(1))) if date_match else pd.NaT
                        if date_match and imported_start_date is None:
                            imported_start_date = duty_date - datetime.timedelta(days=duty_date.weekday())
                        deputy_text = str(sheet.cell(row_idx, deputy_col_idx).value or "")
                        normalized_deputy_text = f" {search_normalize(deputy_text)} "
                        for assistant in assistant_names:
                            normalized_assistant = search_normalize(assistant)
                            if normalized_assistant and f" {normalized_assistant} " in normalized_deputy_text:
                                deputy_occurrences[day_name].append(assistant)
                        for col_idx in range(2, sheet.max_column + 1):
                            heading = search_normalize(sheet.cell(3, col_idx).value or "")
                            if col_idx == deputy_col_idx or "mudur yrd" in heading:
                                continue
                            cell_text = str(sheet.cell(row_idx, col_idx).value or "")
                            current_place = None
                            for line in cell_text.splitlines():
                                prefix, separator, teacher_text = line.partition(":")
                                place = known_locations.get(search_normalize(prefix)) if separator else None
                                if place is not None:
                                    current_place = place
                                elif current_place is not None:
                                    # In newly generated workbooks, each teacher for a
                                    # location is on its own line after the location label.
                                    place = current_place
                                    teacher_text = line
                                if place is None and col_idx in excel_place_map:
                                    place = excel_place_map[col_idx]
                                    floor_match = re.search(r"([123])\s*\.\s*KAT", line.upper())
                                    if floor_match and col_idx in (3, 5):
                                        place = floor_locations.get((col_idx, floor_match.group(1)), place)
                                if place is None:
                                    continue
                                searchable_teacher_text = f" {search_normalize(teacher_text if separator else line)} "
                                for teacher in valid_teachers:
                                    normalized_teacher = search_normalize(teacher)
                                    if normalized_teacher and f" {normalized_teacher} " in searchable_teacher_text:
                                        teacher_occurrences.append({"Tarih/Hafta": hafta_adi_str, "Tarih": duty_date, "Gün": day_name, "Bölüm/Yer": place, "Öğretmen": teacher})

                temp = pd.DataFrame(teacher_occurrences).drop_duplicates(subset=["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen"])
                weekly_teacher_counts = temp.groupby("Öğretmen").size().to_dict() if not temp.empty else {}
                if not temp.empty:
                    temp["Çift Nöbetçi mi"] = temp["Öğretmen"].map(lambda name: weekly_teacher_counts.get(name, 0) > 1)
                    parsed_records = temp.to_dict("records")
                
                if parsed_records:
                    temp_df = pd.DataFrame(parsed_records).reindex(columns=["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi"])
                    temp_df["Kaynak"] = "Eski Liste"
                    temp_df = temp_df.drop_duplicates()
                    st.session_state.history = st.session_state.history[st.session_state.history["Tarih/Hafta"] != hafta_adi_str]
                    st.session_state.history = pd.concat([st.session_state.history, temp_df], ignore_index=True)
                    st.session_state.processed_uploads.add(upload_signature)
                    save_app_data()
                    st.success(f"'{uploaded_file.name}' okundu ve arşive işlendi!")
                normalized_deputies = {
                    day: list(dict.fromkeys(names)) for day, names in deputy_occurrences.items()
                }
                if any(normalized_deputies.values()):
                    st.session_state.mudur_yardimcisi_gecmisi[hafta_adi_str] = {
                        "start_date": imported_start_date.isoformat() if imported_start_date else "",
                        "assignments": normalized_deputies,
                    }
                    if not parsed_records:
                        st.session_state.processed_uploads.add(upload_signature)
                        save_app_data()
                        st.success(f"'{uploaded_file.name}' müdür yardımcısı görevleriyle arşive işlendi!")
            except Exception as e:
                st.error(f"Dosya okunurken hata: {e}")

    st.subheader("Kayıtlı Listeler ve Eski Nöbet Arşivi")
    if (not st.session_state.history.empty and "Tarih/Hafta" in st.session_state.history.columns) or st.session_state.mudur_yardimcisi_gecmisi:
        archive_names = set(st.session_state.mudur_yardimcisi_gecmisi.keys())
        if not st.session_state.history.empty and "Tarih/Hafta" in st.session_state.history.columns:
            archive_names.update(st.session_state.history["Tarih/Hafta"].dropna().astype(str).unique().tolist())
        archive_names = sorted(archive_names)
        if archive_names:
            with st.expander("Arşiv kayıtlarını yönet", expanded=False):
                selected_archive = st.selectbox("Silmek istediğiniz liste veya dönemi seçin", archive_names, key="archive_to_delete")
                if st.button("🗑️ Seçili listeyi sil", key="delete_archive", type="secondary", use_container_width=True):
                    st.session_state.history = st.session_state.history[
                        st.session_state.history["Tarih/Hafta"].astype(str) != selected_archive
                    ].reset_index(drop=True)
                    st.session_state.saved_schedules.pop(selected_archive, None)
                    st.session_state.mudur_yardimcisi_gecmisi.pop(selected_archive, None)
                    if st.session_state.current_schedule is not None and not st.session_state.current_schedule.empty:
                        if (st.session_state.current_schedule["Tarih/Hafta"].astype(str) == selected_archive).any():
                            st.session_state.current_schedule = None
                    save_app_data()
                    st.success(f"'{selected_archive}' silindi; ilgili kayıtlar istatistiklerden kaldırıldı.")
                    st.rerun()
                if st.button("🧹 Eski listelerin tamamını sil", key="delete_all_archives", use_container_width=True):
                    st.session_state.history = st.session_state.history[
                        st.session_state.history["Kaynak"] == "Oluşturulan"
                    ].reset_index(drop=True)
                    retained_periods = set(st.session_state.history["Tarih/Hafta"].astype(str))
                    st.session_state.saved_schedules = {}
                    # Keep generated schedules; remove only imported old-list data.
                    # Their saved copies are reconstructed from the remaining history.
                    for period in retained_periods:
                        period_df = st.session_state.history[st.session_state.history["Tarih/Hafta"].astype(str) == period].copy()
                        st.session_state.saved_schedules[period] = period_df
                    st.session_state.mudur_yardimcisi_gecmisi = {
                        period: entry for period, entry in st.session_state.mudur_yardimcisi_gecmisi.items()
                        if period in retained_periods
                    }
                    if st.session_state.current_schedule is not None and not st.session_state.current_schedule.empty:
                        if not (st.session_state.current_schedule["Tarih/Hafta"].astype(str).isin(retained_periods)).any():
                            st.session_state.current_schedule = None
                    save_app_data()
                    st.success("Eski Word/Excel listeleri silindi; bu listelerin istatistikleri kaldırıldı.")
                    st.rerun()

    st.dataframe(st.session_state.history, use_container_width=True)


# ================= 6. SEKME: İSTATİSTİKLER =================
elif main_menu == "📊 İstatistikler":
    st.header("Öğretmen Nöbet Dağılımı ve Çift Nöbet İstatistikleri")
    if not st.session_state.history.empty:
        hist = st.session_state.history
        
        toplam_stats = hist.groupby("Öğretmen").size().reset_index(name="Toplam Nöbet Sayısı")
        weekly_counts = hist.groupby(["Tarih/Hafta", "Öğretmen"]).size().rename("Haftalık Nöbet").reset_index()
        cift_df = weekly_counts[weekly_counts["Haftalık Nöbet"] > 1]

        if not cift_df.empty:
            cift_stats = cift_df.groupby("Öğretmen").size().reset_index(name="Çift Nöbet Haftası Sayısı")
            final_stats = pd.merge(toplam_stats, cift_stats, on="Öğretmen", how="left")
        else:
            final_stats = toplam_stats
            final_stats["Çift Nöbet Haftası Sayısı"] = 0

        if "Çift Nöbet Sayısı" in final_stats.columns:
            final_stats = final_stats.drop(columns=["Çift Nöbet Sayısı"])
        final_stats["Çift Nöbet Haftası Sayısı"] = final_stats["Çift Nöbet Haftası Sayısı"].fillna(0).astype(int)

        if "Bölüm/Yer" in hist.columns:
            place_stats = hist.pivot_table(
                index="Öğretmen", columns="Bölüm/Yer", values="Tarih/Hafta",
                aggfunc="count", fill_value=0
            )
            place_stats.columns = [str(place) for place in place_stats.columns]
            place_stats = place_stats.reset_index()
            final_stats = pd.merge(final_stats, place_stats, on="Öğretmen", how="left")
            place_columns = [column for column in final_stats.columns if column.endswith(" nöbeti")]
            final_stats[place_columns] = final_stats[place_columns].fillna(0).astype(int)
        final_stats = final_stats.sort_values("Öğretmen", key=lambda column: column.astype(str).str.casefold())
        final_stats.insert(0, "No", range(1, len(final_stats) + 1))
        st.dataframe(final_stats, use_container_width=True, hide_index=True)
    else:
        st.warning("Henüz sistemde işlenmiş nöbet verisi bulunmuyor.")

    st.subheader("Müdür Yardımcısı Nöbet İstatistikleri")
    deputy_totals = {}
    for entry in st.session_state.mudur_yardimcisi_gecmisi.values():
        assignments_by_day = entry.get("assignments", {})
        for day_names in assignments_by_day.values():
            for name in set(day_names):
                deputy_totals[name] = deputy_totals.get(name, 0) + 1
    if deputy_totals:
        deputy_stats = pd.DataFrame([
            {"Müdür Yardımcısı": name, "Toplam Görev Günü": count}
            for name, count in deputy_totals.items()
        ]).sort_values(["Toplam Görev Günü", "Müdür Yardımcısı"], ascending=[False, True])
        deputy_stats.insert(0, "No", range(1, len(deputy_stats) + 1))
        st.dataframe(
            deputy_stats,
            use_container_width=True,
            hide_index=True,
            column_config={"No": st.column_config.NumberColumn("No", width="small")},
        )
    else:
        st.info("Müdür yardımcısı görev geçmişi henüz yok.")

# Persist edits to locations, school information, and any other session data that
# changed during this run. Button actions that rerun save immediately before rerun.
save_app_data()
if st.session_state.get("_github_storage_error"):
    st.warning(
        "GitHub kayıt işlemi tamamlanamadı. Bu oturumdaki veriler yalnızca geçici yerel dosyaya yazılabilir. "
        f"Ayarları kontrol edin. Ayrıntı: {st.session_state['_github_storage_error']}"
    )
st.markdown('<div class="app-footer">KILIÇ BİLİŞİM</div>', unsafe_allow_html=True)

