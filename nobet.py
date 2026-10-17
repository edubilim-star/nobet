import io
import random
import datetime
import re
import os
import html
import tempfile
import pandas as pd
import streamlit as st
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
    layout="wide"
)

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
    st.session_state.history = pd.DataFrame(columns=["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi"])
else:
    if "Tarih" not in st.session_state.history.columns:
        st.session_state.history["Tarih"] = pd.NaT
    st.session_state.history = st.session_state.history.drop(
        st.session_state.history[st.session_state.history.get("Öğretmen", pd.Series(dtype=str)).astype(str).str.strip().str.upper() == "AYŞENUR AKBOĞA"].index,
        errors="ignore"
    )

if "current_schedule" not in st.session_state:
    st.session_state.current_schedule = None

st.title("🏫 Gelişmiş Okul Nöbet Dağıtım ve Yönetim Sistemi")
st.markdown("---")
st.markdown("""
<style>
@media (max-width: 700px) {
  .block-container { padding-left: 0.75rem; padding-right: 0.75rem; }
  [data-testid="stHorizontalBlock"] { flex-wrap: wrap; gap: 0.5rem; }
  [data-testid="column"] { flex: 1 1 100% !important; min-width: 100% !important; }
  [data-testid="stDataFrame"] { max-width: 100%; overflow-x: auto; }
}
</style>
""", unsafe_allow_html=True)

# --- MENÜ SEKMELERİ ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📅 Nöbet Listesi Oluştur",
    "👥 Öğretmen Yönetimi",
    "📍 Nöbet Yerleri ve Kontenjan",
    "⚙️ Okul, Tatil ve Ayarlar",
    "📂 Eski Listeler & Arşiv",
    "📊 İstatistikler"
])

# ================= EXCEL ÜRETME FONKSİYONU =================
def create_custom_excel(df, hafta_baslik, my_listesi, start_date, school_name, principal_name, note_text, holidays_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nöbet Çizelgesi"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:G1")
    ws["A1"] = school_name
    ws["A1"].font = Font(name="Calibri", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:G2")
    ws["A2"] = f"{hafta_baslik} NÖBET ÇİZELGESİ"
    ws["A2"].font = Font(name="Calibri", size=12, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A3"] = "NÖBET GÜNÜ"
    ws["B3"] = "BAHÇE"
    ws["C3"] = "SINIF VE KORİDORLAR"
    ws["D3"] = "ANASINIFI KAPI"
    ws["E3"] = "EK BİNA"
    ws.merge_cells("F3:G3")
    ws["F3"] = "MÜDÜR YRD"

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000")
    )

    for col in range(1, 8):
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    gunler = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
    start_row = 4
    my_str = ", ".join(my_listesi)

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

        def get_all_t(b_adi):
            if tatil_aciklama:
                return f"*** TATİL ({tatil_aciklama}) ***" if b_adi == "BAHÇE" else ""
            matches = gun_df[gun_df["Bölüm/Yer"] == b_adi]["Öğretmen"].tolist()
            clean_matches = [m.split(" (")[0] for m in matches]
            return "\n".join(clean_matches)

        def get_kat_t(kat_adi):
            if tatil_aciklama:
                return ""
            matches = gun_df[gun_df["Bölüm/Yer"] == kat_adi]["Öğretmen"].tolist()
            return matches[0].split(" (")[0] if matches else ""

        ws.cell(row=row, column=1, value=gun_baslik).font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=row, column=2, value=get_all_t('BAHÇE'))
        ws.cell(row=row, column=3, value=f"1.KAT: {get_kat_t('SINIF 1.KAT')}\n2.KAT: {get_kat_t('SINIF 2.KAT')}\n3.KAT: {get_kat_t('SINIF 3.KAT')}")
        ws.cell(row=row, column=4, value=get_all_t('ANASINIFI KAPI'))
        ws.cell(row=row, column=5, value=f"1.KAT: {get_kat_t('EK BİNA 1.KAT')}\n2.KAT: {get_kat_t('EK BİNA 2.KAT')}")
        
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=7)
        ws.cell(row=row, column=6, value=my_str)
        
        ws.row_dimensions[row].height = 70
        
        for col in range(1, 8):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if col in [1, 6, 7] else "left", vertical="center", wrap_text=True)
            cell.font = Font(name="Calibri", size=10, bold=(col==1))

    footer_row = start_row + len(gunler) + 1
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row+1, end_column=7)
    note_cell = ws.cell(row=footer_row, column=1)
    note_cell.value = note_text
    note_cell.font = Font(name="Calibri", size=9, italic=True)
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    sig_row = footer_row + 3
    ws.merge_cells(start_row=sig_row, start_column=5, end_row=sig_row, end_column=7)
    ws.merge_cells(start_row=sig_row+1, start_column=5, end_row=sig_row+1, end_column=7)

    ws.cell(row=sig_row, column=5, value=principal_name).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=sig_row, column=5).alignment = Alignment(horizontal="center")
    ws.cell(row=sig_row+1, column=5, value="OKUL MÜDÜRÜ").font = Font(name="Calibri", size=10, bold=True)
    ws.cell(row=sig_row+1, column=5).alignment = Alignment(horizontal="center")

    col_widths = {"A": 18, "B": 26, "C": 36, "D": 28, "E": 26, "F": 18, "G": 2}
    for col_letter, w in col_widths.items():
        ws.column_dimensions[col_letter].width = w

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
def create_pdf(df, hafta_baslik, my_listesi, school_name, principal_name, note_text, start_date, holidays_df):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    font_candidates = [
        (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
        (r"C:\Windows\Fonts\calibri.ttf", r"C:\Windows\Fonts\calibrib.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    regular_font, bold_font = "Helvetica", "Helvetica-Bold"
    for regular_path, bold_path in font_candidates:
        if os.path.exists(regular_path) and os.path.exists(bold_path):
            pdfmetrics.registerFont(TTFont("SchoolUnicode", regular_path))
            pdfmetrics.registerFont(TTFont("SchoolUnicode-Bold", bold_path))
            regular_font, bold_font = "SchoolUnicode", "SchoolUnicode-Bold"
            break
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=bold_font, fontSize=14, alignment=1)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName=regular_font, fontSize=9)
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
    my_str = ", ".join([safe_text(m) for m in my_listesi])

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

        def get_kat_t(kat_adi):
            if tatil_aciklama:
                return ""
            m = gun_df[gun_df["Bölüm/Yer"] == kat_adi]["Öğretmen"].tolist()
            return safe_text(m[0].split(" (")[0]) if m else ""

        sinif_text = f"1.KAT: {get_kat_t('SINIF 1.KAT')}<br/>2.KAT: {get_kat_t('SINIF 2.KAT')}<br/>3.KAT: {get_kat_t('SINIF 3.KAT')}"
        ek_text = f"1.KAT: {get_kat_t('EK BİNA 1.KAT')}<br/>2.KAT: {get_kat_t('EK BİNA 2.KAT')}"

        table_data.append([
            Paragraph(f"<b>{safe_text(gun)}</b><br/>({gun_tarihi.strftime('%d.%m.%Y')})", cell_style),
            Paragraph(get_all_t('BAHÇE'), cell_style),
            Paragraph(sinif_text, cell_style),
            Paragraph(get_all_t('ANASINIFI KAPI'), cell_style),
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
    elements.append(Paragraph(f"<b>OKUL MÜDÜRÜ:</b> {safe_text(principal_name)}", cell_style))

    doc.build(elements)
    pdf_buffer.seek(0)
    return pdf_buffer


# ================= 1. SEKME: NÖBET LİSTESİ OLUŞTUR =================
with tab1:
    st.header("Haftalık Nöbet Çizelgesi ve Dağıtım Sihirbazı")
    
    col1, col2 = st.columns(2)
    with col1:
        baslangic_tarihi = st.date_input("Haftanın Başlangıç Tarihi (Pazartesi)", value=datetime.date(2026, 10, 19))
    with col2:
        bitis_tarihi = baslangic_tarihi + datetime.timedelta(days=4)
        aylar = {1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN", 5: "MAYIS", 6: "HAZİRAN", 7: "TEMMUZ", 8: "AĞUSTOS", 9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"}
        default_donem = f"{baslangic_tarihi.day}-{bitis_tarihi.day} {aylar[bitis_tarihi.month]}"
        
        hafta_adi = st.text_input("Nöbet Dönem Adı (Otomatik Gelir)", value=default_donem)

    st.write("Seçilen Müdür Yardımcıları:", ", ".join(st.session_state.mudur_yardimcilari))

    if st.button("🎲 Akıllı Kurallı Dağıtımla Çizelgeyi Oluştur", type="primary"):
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
                    capacity = max(0, int(loc_row["Gerekli Öğretmen Sayısı"]))
                except (TypeError, ValueError):
                    st.error(f"'{bolge}' için gerekli öğretmen sayısı geçersiz.")
                    st.stop()
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
        # Avoid repeating a teacher's previous week's location when a valid
        # assignment exists; relax that preference only if staffing is insufficient.
        history = st.session_state.history.copy()
        previous_start = baslangic_tarihi - datetime.timedelta(days=7)
        previous_end = previous_start + datetime.timedelta(days=4)
        if "Tarih" in history.columns:
            history["_date"] = pd.to_datetime(history["Tarih"], errors="coerce").dt.date
            previous_rows = history[
                history["_date"].apply(lambda d: isinstance(d, datetime.date) and previous_start <= d <= previous_end)
            ]
        else:
            previous_rows = pd.DataFrame()
        previous_places = {}
        if not previous_rows.empty:
            for _, row in previous_rows.iterrows():
                previous_places.setdefault(row["Öğretmen"], set()).add(str(row["Bölüm/Yer"]).strip().casefold())
        all_place_counts = {}
        if not history.empty and {"Öğretmen", "Bölüm/Yer"}.issubset(history.columns):
            known_place_history = history[history["Bölüm/Yer"].isin(locs_df["Nöbet Yeri / Bölüm"].astype(str))]
            all_place_counts = known_place_history.groupby(["Öğretmen", "Bölüm/Yer"]).size().to_dict()

        node_budget = 1500000
        visited_nodes = [0]

        def try_solve(avoid_previous_place):
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
                slot_day_index = gunler.index(slot_gun)
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
                    # Never schedule the same teacher on adjacent active weekdays.
                    has_adjacent_day = any(
                        assigned_name == name
                        and abs(gunler.index(slots[other_idx][0]) - slot_day_index) == 1
                        for other_idx, assigned_name in enumerate(assignments)
                        if assigned_name is not None
                    )
                    if has_adjacent_day:
                        continue
                    if avoid_previous_place and slot_place.casefold() in previous_places.get(name, set()):
                        continue
                    candidates.append(name)
                random.shuffle(candidates)
                candidates.sort(key=lambda name: (
                    counts[name] > 0 and not teacher_by_name[name]["Haftada 2 Nöbet İstiyor"],
                    counts[name],
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

        solved_without_repeat = try_solve(avoid_previous_place=True)
        repeated_place_fallback = False
        if not solved_without_repeat:
            repeated_place_fallback = try_solve(avoid_previous_place=False)

        if not solved_without_repeat and not repeated_place_fallback:
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

        ogretmen_haftalik_sayac = dict(counts)
        for (gun, bolge), secilen_ogretmen in zip(slots, assignments):
            day_index = gunler.index(gun)
            schedule_records.append({
                "Tarih/Hafta": hafta_adi,
                "Tarih": baslangic_tarihi + datetime.timedelta(days=day_index),
                "Gün": gun,
                "Bölüm/Yer": bolge,
                "Öğretmen": secilen_ogretmen,
                "Çift Nöbetçi mi": counts[secilen_ogretmen] > 1
            })
        
        sched_df = pd.DataFrame(schedule_records)
        st.session_state.current_schedule = sched_df
        
        if not st.session_state.history.empty:
            st.session_state.history = st.session_state.history[st.session_state.history["Tarih/Hafta"] != hafta_adi]
        
        st.session_state.history = pd.concat([st.session_state.history, sched_df], ignore_index=True)
        st.success(f"✨ '{hafta_adi}' dönemi çizelgesi başarıyla oluşturuldu!")
        if repeated_place_fallback:
            st.warning("Öğretmen sayısı ve diğer kurallar nedeniyle bazı öğretmenlere geçen haftaki nöbet yeri yeniden verildi.")

    if st.session_state.current_schedule is not None:
        st.subheader("Oluşturulan Çizelge Önizlemesi (Toplam ve Çift Nöbet İstatistikleri İle)")
        
        preview_df = st.session_state.current_schedule.copy()
        hist_df = st.session_state.history
        
        def format_teacher_name_with_history(ogretmen_adi):
            toplam_nobet = len(hist_df[hist_df["Öğretmen"] == ogretmen_adi])
            teacher_history = hist_df[hist_df["Öğretmen"] == ogretmen_adi]
            cift_nobet = int((teacher_history.groupby("Tarih/Hafta").size() > 1).sum()) if "Tarih/Hafta" in teacher_history.columns else 0
            return f"{ogretmen_adi} (Toplam Nöbet: {toplam_nobet} ; Çift Nöbet Haftası: {cift_nobet})"

        preview_df["Öğretmen"] = preview_df["Öğretmen"].apply(format_teacher_name_with_history)
        st.dataframe(preview_df, use_container_width=True)

        col_a, col_b = st.columns(2)
        with col_a:
            xlsx_data = create_custom_excel(
                st.session_state.current_schedule, hafta_adi, 
                st.session_state.mudur_yardimcilari, baslangic_tarihi,
                st.session_state.school_info["okul_adi"], st.session_state.school_info["mudur_adi"],
                st.session_state.school_info["not_metni"], st.session_state.holidays
            )
            st.download_button(
                label="📥 Excel Olarak İndir",
                data=xlsx_data,
                file_name=f"{hafta_adi}_Nobet_Cizelgesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        with col_b:
            pdf_data = create_pdf(
                st.session_state.current_schedule, hafta_adi,
                st.session_state.mudur_yardimcilari,
                st.session_state.school_info["okul_adi"], st.session_state.school_info["mudur_adi"],
                st.session_state.school_info["not_metni"], baslangic_tarihi, st.session_state.holidays
            )
            st.download_button(
                label="📄 PDF Olarak İndir",
                data=pdf_data,
                file_name=f"{hafta_adi}_Nobet_Cizelgesi.pdf",
                mime="application/pdf",
                type="secondary"
            )


# ================= 2. SEKME: ÖĞRETMEN YÖNETİMİ =================
with tab2:
    st.header("👥 Öğretmen Yönetimi ve Kısıtlar Paneli")

    tum_bolgeler = st.session_state.locations["Nöbet Yeri / Bölüm"].tolist()
    yer_listesi = ["Yok"] + tum_bolgeler
    gun_listesi = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]

    # Rerun sonrası mesaj göster
    if "_tab2_mesaj" in st.session_state:
        st.success(st.session_state.pop("_tab2_mesaj"))

    # --- TOPLU İŞLEMLER PANELİ (FORM YOK, NORMAL BUTON) ---
    with st.expander("⚡ Toplu Kısıtlama Uygulama Paneli (Tüm Öğretmenler İçin)", expanded=True):
        st.markdown("Aşağıdan bir bölüm seçerek **tüm öğretmenlere tek tıkla** hariç tutma kısıtı ekleyebilir veya kaldırabilirsiniz.")

        col_top1, col_top2 = st.columns(2)
        with col_top1:
            toplu_secilen_bolge = st.selectbox(
                "Toplu İşlem İçin Bölüm Seçin",
                tum_bolgeler,
                key="toplu_bolge_secim_input"
            )
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
                st.rerun()

        with col_top2:
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
                        st.rerun()
                    else:
                        st.warning(f"{yeni_isim} zaten listede var.")

    st.markdown("---")
    st.subheader("📋 Öğretmen Listesi ve Kısıt Ayarları")
    st.caption("Nöbet sayaçları arşivdeki çizelgelere göre hesaplanır. Listeyi ada veya nöbet sayılarına göre sıralayabilirsiniz.")

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
        st.rerun()

    # Değişiklik varsa session_state'i yeniden ata
    if degisiklik_var:
        st.session_state.teachers_dict_list = list(st.session_state.teachers_dict_list)
# ================= 3. SEKME: NÖBET YERLERİ VE KONTENJAN =================
with tab3:
    st.header("Nöbet Yerleri ve Kişi Sayısı Yönetimi")
    edited_locs = st.data_editor(st.session_state.locations, num_rows="dynamic", use_container_width=True)
    st.session_state.locations = edited_locs


# ================= 4. SEKME: OKUL, TATİL VE AYARLAR =================
with tab4:
    st.header("Okul, Yönetici ve Tatil Günü Ayarları")
    okul_adi_input = st.text_input("Okul Adı", value=st.session_state.school_info["okul_adi"])
    mudur_adi_input = st.text_input("Okul Müdürü Ad Soyad", value=st.session_state.school_info["mudur_adi"])
    not_metni_input = st.text_area("Çizelge Altı Notu (Açıklama Metni)", value=st.session_state.school_info["not_metni"], height=100)
    
    if st.button("Bilgileri Kaydet"):
        st.session_state.school_info["okul_adi"] = okul_adi_input
        st.session_state.school_info["mudur_adi"] = mudur_adi_input
        st.session_state.school_info["not_metni"] = not_metni_input
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
            st.success("Tatil kayıtları güncellendi.")
            st.rerun()
        to_delete = st.multiselect("Silinecek tatil tarihleri", holidays_view["Tarih"].tolist(), format_func=lambda d: d.strftime("%d.%m.%Y"), key="holiday_delete_select")
        if st.button("🗑️ Seçili tatil günlerini sil", key="holiday_delete", disabled=not to_delete):
            st.session_state.holidays = holidays_view[~holidays_view["Tarih"].isin(to_delete)].reset_index(drop=True)
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
            st.success(f"{my_input} eklendi.")
            st.rerun()

    for idx, my in enumerate(st.session_state.mudur_yardimcilari):
        col_a, col_b = st.columns([4, 1])
        with col_a:
            st.write(f"- {my}")
        with col_b:
            if st.button("Sil", key=f"del_my_{idx}"):
                st.session_state.mudur_yardimcilari.pop(idx)
                st.rerun()


# ================= 5. SEKME: ESKİ LİSTELER & ARŞİV =================
with tab5:
    st.header("Eski Listeler ve Nöbet Arşivi")
    st.markdown("Geçmiş Word veya bu uygulamanın oluşturduğu Excel çizelgelerini yükleyin. Excel aktarımı tarih, gün ve nöbet yerlerini arşive işler.")
    
    uploaded_files = st.file_uploader("Eski nöbet listelerini seçin (Word veya oluşturulan Excel)", type=["docx", "xlsx"], accept_multiple_files=True)
    
    if uploaded_files:
        valid_teachers = set([t["Ad Soyad"].upper() for t in st.session_state.teachers_dict_list])
        
        for uploaded_file in uploaded_files:
            try:
                parsed_records = []
                hafta_adi_str = os.path.splitext(uploaded_file.name)[0].upper()
                teacher_occurrences = []
                if uploaded_file.name.lower().endswith(".docx"):
                    doc = Document(uploaded_file)
                    current_gun = "BİLİNMEYEN GÜN"
                    all_texts = [p.text.strip() for p in doc.paragraphs if p.text]
                    all_texts.extend(cell.text.strip() for table in doc.tables for row in table.rows for cell in row.cells if cell.text)
                    seen_occurrences = set()
                    for text_block in all_texts:
                        for line in text_block.split("\n"):
                            clean_line = line.strip().upper()
                            if clean_line in ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]:
                                current_gun = clean_line
                                continue
                            for teacher in valid_teachers:
                                if teacher in clean_line:
                                    occurrence = (current_gun, clean_line, teacher)
                                    if occurrence not in seen_occurrences:
                                        teacher_occurrences.append({"Tarih/Hafta": hafta_adi_str, "Tarih": pd.NaT, "Gün": current_gun, "Bölüm/Yer": "ESKİ ARŞİV", "Öğretmen": teacher})
                                        seen_occurrences.add(occurrence)
                else:
                    sheet = openpyxl.load_workbook(uploaded_file, data_only=True, read_only=True).active
                    if sheet["A2"].value:
                        hafta_adi_str = str(sheet["A2"].value).replace(" NÖBET ÇİZELGESİ", "").strip().upper()
                    excel_place_map = {2: "BAHÇE", 3: "SINIF", 4: "ANASINIFI KAPI", 5: "EK BİNA"}
                    floor_locations = {
                        (3, "1"): "SINIF 1.KAT", (3, "2"): "SINIF 2.KAT", (3, "3"): "SINIF 3.KAT",
                        (5, "1"): "EK BİNA 1.KAT", (5, "2"): "EK BİNA 2.KAT",
                    }
                    for row_idx in range(4, 9):
                        day_cell = str(sheet.cell(row_idx, 1).value or "")
                        day_match = re.search(r"(PAZARTESİ|SALI|ÇARŞAMBA|PERŞEMBE|CUMA)", day_cell.upper())
                        if not day_match:
                            continue
                        day_name = day_match.group(1)
                        date_match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", day_cell)
                        duty_date = datetime.date(int(date_match.group(3)), int(date_match.group(2)), int(date_match.group(1))) if date_match else pd.NaT
                        for col_idx, base_place in excel_place_map.items():
                            cell_text = str(sheet.cell(row_idx, col_idx).value or "").upper()
                            for teacher in valid_teachers:
                                for line in cell_text.splitlines():
                                    if teacher not in line:
                                        continue
                                    place = base_place
                                    if col_idx in (3, 5):
                                        floor_match = re.search(r"([123])\s*\.\s*KAT", line)
                                        if floor_match:
                                            place = floor_locations.get((col_idx, floor_match.group(1)), base_place)
                                    teacher_occurrences.append({"Tarih/Hafta": hafta_adi_str, "Tarih": duty_date, "Gün": day_name, "Bölüm/Yer": place, "Öğretmen": teacher})

                temp = pd.DataFrame(teacher_occurrences).drop_duplicates(subset=["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen"])
                weekly_teacher_counts = temp.groupby("Öğretmen").size().to_dict() if not temp.empty else {}
                if not temp.empty:
                    temp["Çift Nöbetçi mi"] = temp["Öğretmen"].map(lambda name: weekly_teacher_counts.get(name, 0) > 1)
                    parsed_records = temp.to_dict("records")
                
                if parsed_records:
                    temp_df = pd.DataFrame(parsed_records).reindex(columns=["Tarih/Hafta", "Tarih", "Gün", "Bölüm/Yer", "Öğretmen", "Çift Nöbetçi mi"]).drop_duplicates()
                    st.session_state.history = st.session_state.history[st.session_state.history["Tarih/Hafta"] != hafta_adi_str]
                    st.session_state.history = pd.concat([st.session_state.history, temp_df], ignore_index=True)
                    st.success(f"'{uploaded_file.name}' okundu ve arşive işlendi!")
            except Exception as e:
                st.error(f"Dosya okunurken hata: {e}")

    st.subheader("Kayıtlı Tüm Nöbet Geçmişi Arşivi")
    st.dataframe(st.session_state.history, use_container_width=True)


# ================= 6. SEKME: İSTATİSTİKLER =================
with tab6:
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

        st.dataframe(final_stats, use_container_width=True)
    else:
        st.warning("Henüz sistemde işlenmiş nöbet verisi bulunmuyor.")
