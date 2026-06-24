import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

st.set_page_config(page_title="Tohum Depo Ölçüm Merkezi", page_icon="🌾", layout="centered")

# --- HEM Z-ŞEMASINI HEM ÖZET TABLOYU TEK SAYFADA ÜRETEN MOTOR ---
def mükemmel_excel_olustur():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fabrika Genel Raporu"
    ws.views.sheetView[0].showGridLines = True

    # Tasarım Renkleri
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
    thin_line = Side(border_style="thin", color="D0D7DE")
    border_all_thin = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)

    # 1. BAŞLIK
    ws.merge_cells("A1:M1")
    ws["A1"] = "FABRİKA GENELİ TOHUM DEPOLARI SICAKLIK VE NEM TAKİP RAPORU"
    ws["A1"].font = font_title
    ws["A1"].fill = HEADER_FILL
    ws["A1"].alignment = align_center
    for c in range(1, 14):
        ws.cell(row=1, column=c).fill = HEADER_FILL

    # 2. ÜST KISIM: Z-ŞEMASI GÖRSEL ALANI
    ws.merge_cells("A3:M3")
    ws["A3"] = "SEÇİLİ DEPONUN GÖRSEL YIĞIN DİP SICAKLIK DAĞILIMI (Z-ŞEMASI)"
    ws["A3"].font = font_sub
    ws["A3"].fill = SUBHEADER_FILL
    ws["A3"].alignment = align_center

    z_sablonu = {
        "B5": "1. ÜST SOL", "G5": "2. ÜST ORTA", "L5": "3. ÜST SAĞ",
        "G8": "4. TAM ORTA",
        "B11": "5. ALT SOL", "G11": "6. ALT ORTA", "L11": "7. ALT SAĞ"
    }
    
    for cell_ref, label in z_sablonu.items():
        ws[cell_ref] = "-"
        ws[cell_ref].font = font_bold
        ws[cell_ref].alignment = align_center
        ws[cell_ref].fill = Z_CELL_FILL
        ws[cell_ref].border = border_all_thin
        
        col_letter = cell_ref[0]
        row_num = int(cell_ref[1:])
        ws[f"{col_letter}{row_num-1}"] = label
        ws[f"{col_letter}{row_num-1}"].font = font_italic
        ws[f"{col_letter}{row_num-1}"].alignment = align_center

    # 3. ALT KISIM: ÖZET LİSTE
    ws.merge_cells("A14:M14")
    ws["A14"] = "TÜM FABRİKA DEPOLARI ANLIK ÖZET TABLOTSU"
    ws["A14"].font = font_sub
    ws["A14"].fill = SUBHEADER_FILL
    ws["A14"].alignment = align_center

    headers = [
        "Depo No", "Ortam Sıcaklığı (°C)", "Ortam Nemi (%)", 
        "1. Üst Sol (Z)", "2. Üst Orta (Z)", "3. Üst Sağ (Z)", 
        "4. Tam Orta (Z)", "5. Alt Sol (Z)", "6. Alt Orta (Z)", "7. Alt Sağ (Z)",
        "Ölçüm Tarihi", "Ölçüm Saati", "Ölçen Personel"
    ]
    
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=15, column=idx, value=h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = align_center
        cell.border = border_all_thin
    ws.row_dimensions[15].height = 25

    for d_id in range(1, 13):
        row_idx = 15 + d_id
        ws.row_dimensions[row_idx].height = 22
        ws.cell(row=row_idx, column=1, value=f"Depo {d_id}").font = font_bold
        ws.cell(row=row_idx, column=1).alignment = align_center
        
        for col_idx in range(2, 14):
            c = ws.cell(row=row_idx, column=col_idx, value="-")
            c.alignment = align_center
            c.font = font_regular
            c.border = border_all_thin
            if d_id % 2 == 0:
                c.fill = ZEBRA_FILL
                ws.cell(row=row_idx, column=1).fill = ZEBRA_FILL

    widths = [12, 18, 15, 14, 14, 14, 14, 14, 14, 14, 14, 14, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save("mükemmel_merkezi_rapor.xlsx")

if not os.path.exists("mükemmel_merkezi_rapor.xlsx"):
    mükemmel_excel_olustur()

# --- HAFIZA MOTORU ---
if 'fabrika_verisi' not in st.session_state:
    st.session_state.fabrika_verisi = {f"Depo {i}": {} for i in range(1, 13)}

st.title("🌾 Z-Şemalı Fabrika Ölçüm İstasyonu")
st.write("Hem kağıttaki Z görseli hem de büyük özet tablo tek ekranda!")

# --- GİRİŞ ALANI ---
st.subheader("1. Genel Bilgiler")
col_a, col_b = st.columns(2)
with col_a:
    depo_no = st.selectbox("Ölçülen Depo:", [f"Depo {i}" for i in range(1, 13)])
    olcen_kisi = st.text_input("Ölçen Personel", "Ali Sefa")
with col_b:
    saat_araligi = st.text_input("Ölçüm Saati", datetime.now().strftime("%H:%M"))

st.markdown("---")

# --- Z SEÇİM VE DEĞER GİRİŞİ ---
st.subheader(f"2. {depo_no} Sıcaklık ve Nem Girişi")

# Hatalı olan o1 ve o2 kısımları jilet gibi temizlendi kanka:
st.write("📍 **Depo İçi Hücre Pozisyonları (Z Düzeni):**")
c1, c2, c3 = st.columns(3)
with c1: st.info("1. Üst Sol")
with c2: st.info("2. Üst Orta")
with c3: st.info("3. Üst Sağ")

c4, c5, c6 = st.columns(3)
with c4: st.write("")
with c5: st.warning("4. TAM ORTA")
with c6: st.write("")

c7, c8, c9 = st.columns(3)
with c7: st.info("5. Alt Sol")
with c8: st.info("6. Alt Orta")
with c9: st.info("7. Alt Sağ")

noktalar = [
    "Ortam Sıcaklığı", "Ortam Nemi (%)",
    "1. Üst Sol (Z)", "2. Üst Orta (Z)", "3. Üst Sağ (Z)", 
    "4. TAM ORTA (Z)", 
    "5. Alt Sol (Z)", "6. Alt Orta (Z)", "7. Alt Sağ (Z)"
]

secilen_nokta = st.selectbox("Değerini Gireceğiniz Noktayı Seçin:", noktalar)

if "Nemi" in secilen_nokta:
    deger = st.number_input(f"{secilen_nokta} Değeri", min_value=0, max_value=100, value=35, step=1)
else:
    deger = st.number_input(f"{secilen_nokta} Değeri (°C)", min_value=15.0, max_value=45.0, value=25.0, step=0.1, format="%.1f")

if st.button(f"📌 {depo_no} Verisini Hafızaya Ekle", use_container_width=True):
    st.session_state.fabrika_verisi[depo_no][secilen_nokta] = deger
    st.session_state.fabrika_verisi[depo_no]["Zaman"] = saat_araligi
    st.session_state.fabrika_verisi[depo_no]["Kisi"] = olcen_kisi
    st.success(f"✔️ {depo_no} - {secilen_nokta} başarıyla kaydedildi!")

# Mevcut Deponun Durumu
st.markdown(f"#### 📊 {depo_no} Güncel Durumu")
for n in noktalar:
    if n in st.session_state.fabrika_verisi[depo_no]:
        birim = " %" if "Nemi" in n else " °C"
        st.write(f"✅ **{n}:** {st.session_state.fabrika_verisi[depo_no][n]}{birim}")
    else:
        st.write(f"❌ **{n}:** Eksik...")

st.markdown("---")

# --- RAPORU KAYDET VE İNDİR ---
st.subheader("3. Raporu Tamamla")

if st.button("💾 TÜM FABRİKAYI Z-ŞEMALI EXCEL'E AKTAR VE İNDİR", use_container_width=True):
    wb = openpyxl.load_workbook("mükemmel_merkezi_rapor.xlsx")
    ws = wb["Fabrika Genel Raporu"]
    tarih_str = datetime.now().strftime("%d.%m.%Y")
    
    kirmizi_dolgu = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    kirmizi_yazi = Font(name="Arial", size=10, bold=True, color="9C0006")
    
    aktif_depo_sayisi = 0
    
    for d_id in range(1, 13):
        d_name = f"Depo {d_id}"
        d_data = st.session_state.fabrika_verisi[d_name]
        
        if len(d_data) >= len(noktalar):
            aktif_depo_sayisi += 1
            row_idx = 15 + d_id
            
            # Üst taraftaki Z Görsel Şemasına en son seçili deponun değerlerini basalım
            if d_name == depo_no:
                ws["B5"] = float(d_data["1. Üst Sol (Z)"])
                ws["G5"] = float(d_data["2. Üst Orta (Z)"])
                ws["L5"] = float(d_data["3. Üst Sağ (Z)"])
                ws["G8"] = float(d_data["4. TAM ORTA (Z)"])
                ws["B11"] = float(d_data["5. Alt Sol (Z)"])
                ws["G11"] = float(d_data["6. Alt Orta (Z)"])
                ws["L11"] = float(d_data["7. Alt Sağ (Z)"])
            
            # Alt taraftaki Özet Listeyi Doldurma
            ws.cell(row=row_idx, column=2, value=float(d_data["Ortam Sıcaklığı"]))
            ws.cell(row=row_idx, column=3, value=f"%{d_data['Ortam Nemi (%)']}")
            
            z_degerleri = [
                float(d_data["1. Üst Sol (Z)"]),
                float(d_data["2. Üst Orta (Z)"]),
                float(d_data["3. Üst Sağ (Z)"]),
                float(d_data["4. TAM ORTA (Z)"]),
                float(d_data["5. Alt Sol (Z)"]),
                float(d_data["6. Alt Orta (Z)"]),
                float(d_data["7. Alt Sağ (Z)"])
            ]
            
            for i, val in enumerate(z_degerleri, start=4):
                ws.cell(row=row_idx, column=i, value=val)
                
            ws.cell(row=row_idx, column=11, value=tarih_str)
            ws.cell(row=row_idx, column=12, value=d_data.get("Zaman", saat_araligi))
            ws.cell(row=row_idx, column=13, value=d_data.get("Kisi", olcen_kisi))
            
            # En yüksek olanı "Yüksek" olarak işaretleme ve kırmızı yapma
            en_yuksek = max(z_degerleri)
            for i, val in enumerate(z_degerleri, start=4):
                if val == en_yuksek:
                    cell = ws.cell(row=row_idx, column=i)
                    cell.value = f"{val} (Yüksek)"
                    cell.fill = kirmizi_dolgu
                    cell.font = kirmizi_yazi

    if aktif_depo_sayisi == 0:
        st.error("Rapor üretmek için en az 1 deponun tüm verilerini girmelisin kanka!")
    else:
        cikti_adi = f"Z_Semali_Fabrika_Raporu_{tarih_str.replace('.', '_')}.xlsx"
        wb.save(cikti_adi)
        
        with open(cikti_adi, "rb") as file:
            st.download_button(
                label="📥 Z-ŞEMALI GENEL FABRİKA RAPORUNU İNDİR",
                data=file,
                file_name=cikti_adi,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.balloons()
