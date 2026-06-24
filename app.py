import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import os

st.set_page_config(page_title="Tohum Depo Ölçüm Merkezi", page_icon="🌾", layout="centered")

# --- TEK EXCELDE ÇOKLU SAYFA (DEPO) ÜRETEN MOTOR ---
def merkezi_excel_olustur():
    wb = openpyxl.Workbook()
    # İlk açılıştaki boş sayfayı silelim, temiz isimlerle baştan açacağız
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # 12 Depo için aynı şablonu ayrı sayfalar olarak oluşturuyoruz
    for d_id in range(1, 13):
        ws = wb.create_sheet(title=f"Depo {d_id}")
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
        align_right = Alignment(horizontal="right", vertical="center")
        thin_line = Side(border_style="thin", color="D0D7DE")
        thick_line = Side(border_style="medium", color="2B4C7E")
        border_all_thin = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)

        # Başlıklar
        ws.merge_cells("A1:G2")
        ws["A1"] = f"DEPO {d_id} - SICAKLIK VE NEM TAKİP RAPORU"
        ws["A1"].font = font_title
        ws["A1"].fill = HEADER_FILL
        ws["A1"].alignment = align_center
        for r in [1, 2]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = HEADER_FILL

        ws["A3"] = "DEPO / BÖLÜM:"
        ws["B3"] = f"Depo {d_id}"
        ws["C3"] = "ORTAM SICAKLIĞI:"
        ws["E3"] = "ÖLÇÜM SAATİ:"
        ws["G3"] = "NEM (%):"
        for col in ["A", "C", "E", "G"]:
            ws[f"{col}3"].font = font_sub
            ws[f"{col}3"].alignment = align_right
        ws["B3"].font = font_bold
        ws["B3"].alignment = align_center

        ws.merge_cells("A4:G4")
        ws["A4"] = "YIĞIN DİP SICAKLIKLARI GÖRSEL DAĞILIM ŞEMASI (Z)"
        ws["A4"].font = font_sub
        ws["A4"].fill = SUBHEADER_FILL
        ws["A4"].alignment = align_center
        for c in range(1, 8):
            ws.cell(row=4, column=c).fill = SUBHEADER_FILL

        z_points = {
            "B6": "1. ÜST SOL", "D6": "2. ÜST ORTA", "F6": "3. ÜST SAĞ",
            "D9": "4. TAM ORTA",
            "B12": "5. ALT SOL", "D12": "6. ALT ORTA", "F12": "7. ALT SAĞ"
        }
        for col in ["B", "C", "D", "E", "F"]:
            ws[f"{col}5"].border = Border(top=thick_line)
            ws[f"{col}12"].border = Border(bottom=thick_line)

        for cell_ref, label in z_points.items():
            ws[cell_ref] = 0.0
            ws[cell_ref].font = font_bold
            ws[cell_ref].alignment = align_center
            ws[cell_ref].fill = Z_CELL_FILL
            ws[cell_ref].border = border_all_thin
            
            l_cell = ws[f"{cell_ref[0]}{int(cell_ref[1:])-1}"]
            l_cell.value = label
            l_cell.font = font_italic
            l_cell.alignment = align_center

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
            ws.cell(row=idx, column=2, value="-")
            ws.cell(row=idx, column=3, value="Normal")
            ws.cell(row=idx, column=4, value="-")
            ws.cell(row=idx, column=5, value="-")
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

    wb.save("merkezi_sablon.xlsx")

if not os.path.exists("merkezi_sablon.xlsx"):
    merkezi_excel_olustur()

# --- STREAMLIT SESSİON STATE YAPISI ---
if 'fabrika_verisi' not in st.session_state:
    st.session_state.fabrika_verisi = {f"Depo {i}": {} for i in range(1, 13)}

st.title("🌾 Merkezi Tohum Ölçüm İstasyonu")
st.write("Tüm depolar tek bir Excel dosyasında toplanıyor kanka!")

# --- 1. ADIM: GENEL BİLGİLER ---
st.subheader("1. Genel Bilgiler")
col_a, col_b = st.columns(2)
with col_a:
    depo_no = st.selectbox("Hangi Depoyu Ölçüyorsunuz?",
