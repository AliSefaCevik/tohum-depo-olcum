import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import os

# Mobil Uyumlu Sayfa Yapısı
st.set_page_config(page_title="Tohum Depo Ölçüm", page_icon="🌾", layout="centered")

# --- OTOMATİK ŞABLON ÜRETİCİ FONKSİYON ---
def sablon_olustur():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Depo Ölçüm Raporu"
    ws.views.sheetView[0].showGridLines = True

    # Renkler ve Fontlar
    HEADER_FILL = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")
    SUBHEADER_FILL = PatternFill(start_color="EAF0F6", end_color="EAF0F6", fill_type="solid")
    Z_CELL_FILL = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F9FBFC", end_color="F9FBFC", fill_type="solid")

    font_title = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    font_sub = Font(name="Arial", size=10, bold=True, color="1A335E")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_regular = Font(name="Arial", size=10)
    font_italic = Font(name="Arial", size=9, italic=True, color="555555")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    thin_line = Side(border_style="thin", color="D0D7DE")
    thick_line = Side(border_style="medium", color="2B4C7E")
    border_all_thin = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)

    # Başlık Alanı
    ws.merge_cells("A1:G2")
    title_cell = ws["A1"]
    title_cell.value = "TOHUM DEPOSU SICAKLIK VE NEM TAKİP RAPORU (Z-ŞEMASI)"
    title_cell.font = font_title
    title_cell.fill = HEADER_FILL
    title_cell.alignment = align_center
    for r in [1, 2]:
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = HEADER_FILL

    # Bilgi Satırı
    ws["A3"] = "DEPO / BÖLÜM:"
    ws["C3"] = "ORTAM SICAKLIĞI:"
    ws["E3"] = "ÖLÇÜM SAATİ:"
    ws["G3"] = "NEM (%):"
    for col in ["A", "C", "E", "G"]:
        ws[f"{col}3"].font = font_sub
        ws[f"{col}3"].alignment = align_right
    for col in ["B", "D", "F"]:
        ws[f"{col}3"].font = font_bold
        ws[f"{col}3"].alignment = align_center

    # Şema Başlığı
    ws.merge_cells("A4:G4")
    ws["A4"] = "YIĞIN DİP SICAKLIKLARI GÖRSEL DAĞILIM ŞEMASI (Z)"
    ws["A4"].font = font_sub
    ws["A4"].fill = SUBHEADER_FILL
    ws["A4"].alignment = align_center
    for c in range(1, 8):
        ws.cell(row=4, column=c).fill = SUBHEADER_FILL

    # Z Şeması Yapısı
    z_points = {
        "B6": "1. ÜST SOL", "D6": "2. ÜST ORTA", "F6": "3. ÜST SAĞ",
        "D9": "4. TAM ORTA",
        "B12": "5. ALT SOL", "D12": "6. ALT ORTA", "F12": "7. ALT SAĞ"
    }
    for col in ["B", "C", "D", "E", "F"]:
        ws[f"{col}5"].border = Border(top=thick_line)
        ws[f"{col}12"].border = Border(bottom=thick_line)

    for cell_ref, label in z_points.items():
        v_cell = ws[cell_ref]
        v_cell.value = 0.0
        v_cell.font = font_bold
        v_cell.alignment = align_center
        v_cell.fill = Z_CELL_FILL
        v_cell.border = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)
        
        col_letter = cell_ref[0]
        row_num = int(cell_ref[1:])
        l_cell = ws[f"{col_letter}{row_num-1}"]
        l_cell.value = label
        l_cell.font = font_italic
        l_cell.alignment = align_center

    # Özet Liste Başlığı
    ws.merge_cells("A14:G14")
    ws["A14"] = "ÖLÇÜM VERİLERİ ÖZET LİSTESİ"
    ws["A14"].font = font_sub
    ws["A14"].fill = SUBHEADER_FILL
    ws["A14"].alignment = align_center
    for c in range(1, 8):
        ws.cell(row=14, column=c).fill = SUBHEADER_FILL

    headers = ["Nokta No / Tanımı", "Ölçülen Değer (°C)", "Durum / Limit", "Tarih", "Saat", "Ölçen Personel", "Açıklama"]
    for idx, h in enumerate(headers, start=1):
        c = ws.cell(row=15, column=idx)
        c.value = h
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = align_center

    point_names = [
        "Ortam Sıcaklığı", "1. Üst Sol Noktası", "2. Üst Orta Noktası", "3. Üst Sağ Noktası",
        "4. Tam Orta Noktası", "5. Alt Sol Noktası", "6. Alt Orta Noktası", "7. Alt Sağ Noktası"
    ]
    for idx, name in enumerate(point_names, start=16):
        ws.cell(row=idx, column=1, value=name).font = font_bold
        ws.cell(row=idx, column=2, value="-").alignment = align_center
        ws.cell(row=idx, column=3, value="Normal").alignment = align_center
        ws.cell(row=idx, column=4, value="-").alignment = align_center
        ws.cell(row=idx, column=5, value="-").alignment = align_center
        ws.cell(row=idx, column=6, value="-")
        ws.cell(row=idx, column=7, value="Z-Şeması Takibi")
        
        row_fill = ZEBRA_FILL if idx % 2 == 0 else PatternFill(fill_type=None)
        for col in range(1, 8):
            cell = ws.cell(row=idx, column=col)
            cell.border = border_all_thin
            if row_fill.fill_type:
                cell.fill = row_fill
            cell.font = font_regular

    widths = {"A": 22, "B": 15, "C": 18, "D": 15, "E": 15, "F": 15, "G": 16}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    wb.save("sablon.xlsx")

# Şablon yoksa üret
if not os.path.exists("sablon.xlsx"):
    sablon_olustur()

# --- STREAMLIT ARAYÜZÜ ---
st.title("🌾 Tohum Depo Ölçüm Sitesi")
st.write("Klavyeden hızlı değer girmeli güncel versiyon.")

if 'olcumler' not in st.session_state:
    st.session_state.olcumler = {}

# --- 1. ADIM: GENEL BİLGİLER ---
st.subheader("1. Genel Bilgiler")
col_a, col_b = st.columns(2)
with col_a:
    depo_no = st.selectbox("Depo / Bölüm No", [f"Depo {i}" for i in range(1, 21)], index=9)
    tohum_cinsi = st.selectbox("Tohum Cinsi", ["Buğday", "Arpa", "Mısır", "Ayçiçeği", "Nohut"])
with col_b:
    olcen_kisi = st.text_input("Ölçen Kişi", "Ali Sefa")
    saat_araligi = st.text_input("Saat Aralığı", datetime.now().strftime("%H:%M"))

st.markdown("---")

# --- 2. ADIM: VERİ GİRİŞ ALANI (KLAVYEDEN SAYI GİRMELİ) ---
st.subheader("2. Sıcaklık ve Nem Değerleri")

noktalar = [
    "Ortam Sıcaklığı", "Ortam Nemi (%)",
    "1. Üst Sol (Z)", "2. Üst Orta (Z)", "3. Üst Sağ (Z)", 
    "4. TAM ORTA (Z)", 
    "5. Alt Sol (Z)", "6. Alt Orta (Z)", "7. Alt Sağ (Z)"
]

secilen_nokta = st.selectbox("Değer gireceğiniz noktayı seçin:", noktalar)

# İŞTE O KRİTİK DEĞİŞİKLİK: Kaydırmalı liste yerine kutuya klavyeden yazma alanı
if "Nemi" in secilen_nokta:
    # Nem için tam sayı girişi (0-100 arası)
    deger = st.number_input(f"{secilen_nokta} Değerini Yazın", min_value=0, max_value=100, value=35, step=1)
else:
    # Sıcaklık için ondalıklı sayı girişi (15.0 - 45.0 arası, 0.1 hassasiyet)
    deger = st.number_input(f"{secilen_nokta} Değerini Yazın (°C)", min_value=15.0, max_value=45.0, value=25.0, step=0.1, format="%.1f")

if st.button("📌 Değeri Hafızaya Kaydet", use_container_width=True):
    st.session_state.olcumler[secilen_nokta] = deger
    st.success(f"✔️ {secilen_nokta} için {deger} hafızaya alındı!")

# Canlı Durum Tablosu
st.markdown("### 📊 Mevcut Giriş Durumu")
for n in noktalar:
    if n in st.session_state.olcumler:
        birim = " %" if "Nemi" in n else " °C"
        st.write(f"✅ **{n}:** {st.session_state.olcumler[n]}{birim}")
    else:
        st.write(f"❌ **{n}:** Veri bekleniyor...")

st.markdown("---")

# --- 3. ADIM: EXCEL'E YAZMA VE İNDİRME ---
if st.button("💾 EXCEL ŞABLONUNA AKTAR VE DOSYAYI İNDİR", use_container_width=True):
    eksikler = [n for n in noktalar if n not in st.session_state.olcumler]
    
    if eksikler:
        st.error(f"Hata! Tüm noktaları doldurmalısın. Eksikler: {', '.join(eksikler)}")
    else:
        wb = openpyxl.load_workbook("sablon.xlsx")
        ws = wb.active
        tarih_str = datetime.now().strftime("%d.%m.%Y")
        
        # Üst Bilgileri Doldurma
        ws["B3"] = depo_no
        ws["D3"] = float(st.session_state.olcumler["Ortam Sıcaklığı"])
        ws["F3"] = saat_araligi
        ws["G3"] = f"NEM: %{st.session_state.olcumler['Ortam Nemi (%)']}"
        
        # Z Şeması hücreleri
        ws["B6"] = float(st.session_state.olcumler["1. Üst Sol (Z)"])
        ws["D6"] = float(st.session_state.olcumler["2. Üst Orta (Z)"])
        ws["F6"] = float(st.session_state.olcumler["3. Üst Sağ (Z)"])
        ws["D9"] = float(st.session_state.olcumler["4. TAM ORTA (Z)"])
        ws["B12"] = float(st.session_state.olcumler["5. Alt Sol (Z)"])
        ws["D12"] = float(st.session_state.olcumler["6. Alt Orta (Z)"])
        ws["F12"] = float(st.session_state.olcumler["7. Alt Sağ (Z)"])
        
        # Özet Tabloyu Doldurma (Satır 16-23)
        mapping = {
            16: ("Ortam Sıcaklığı", "Ortam Sıcaklığı"),
            17: ("1. Üst Sol Noktası", "1. Üst Sol (Z)"),
            18: ("2. Üst Orta Noktası", "2. Üst Orta (Z)"),
            19: ("3. Üst Sağ Noktası", "3. Üst Sağ (Z)"),
            20: ("4. Tam Orta Noktası", "4. TAM ORTA (Z)"),
            21: ("5. Alt Sol Noktası", "5. Alt Sol (Z)"),
            22: ("6. Alt Orta Noktası", "6. Alt Orta (Z)"),
            23: ("7. Alt Sağ Noktası", "7. Alt Sağ (Z)")
        }
        
        for row_idx, (label, state_key) in mapping.items():
            ws.cell(row=row_idx, column=2, value=float(st.session_state.olcumler[state_key]))
            ws.cell(row=row_idx, column=4, value=tarih_str)
            ws.cell(row=row_idx, column=5, value=saat_araligi)
            ws.cell(row=row_idx, column=6, value=olcen_kisi)
        
        cikti_adi = f"Olcum_Raporu_{depo_no}_{tarih_str.replace('.', '_')}.xlsx"
        wb.save(cikti_adi)
        
        with open(cikti_adi, "rb") as file:
            st.download_button(
                label="📥 DOLDURULMUŞ EXCEL'İ İNDİR",
                data=file,
                file_name=cikti_adi,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.balloons()
