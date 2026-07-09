import os
from fpdf import FPDF

# Paths to Windows fonts for Turkish character support
FONT_REGULAR = r"C:\Windows\Fonts\arial.ttf"
FONT_BOLD = r"C:\Windows\Fonts\arialbd.ttf"

class WorkOrderPDF(FPDF):
    def __init__(self, data):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.data = data
        
        # Load fonts if they exist, otherwise fallback to Helvetica (which lacks full Turkish support but prevents crash)
        if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
            self.add_font("ArialTR", "", FONT_REGULAR)
            self.add_font("ArialTR", "B", FONT_BOLD)
            self.font_name = "ArialTR"
        else:
            self.font_name = "Helvetica"

    def header(self):
        # Draw a beautiful top border/bar
        self.set_fill_color(31, 78, 121)  # Deep corporate blue
        self.rect(0, 0, 210, 8, 'F')
        self.ln(5)

    def draw_section_title(self, title):
        self.set_font(self.font_name, "B", 12)
        self.set_text_color(31, 78, 121)
        self.cell(0, 8, title, ln=True)
        # Draw a thin horizontal line below section title
        self.set_draw_color(31, 78, 121)
        self.set_line_width(0.5)
        self.line(self.get_x(), self.get_y(), 200, self.get_y())
        self.ln(3)

    def draw_field(self, label, value, width):
        # Background for label
        self.set_fill_color(240, 240, 240)
        self.set_font(self.font_name, "B", 9)
        self.set_text_color(50, 50, 50)
        self.cell(width * 0.4, 7, f" {label}", border=1, fill=True)
        
        # Value cell
        self.set_font(self.font_name, "", 9)
        self.set_text_color(0, 0, 0)
        val_str = str(value) if value is not None and str(value) != 'nan' else '-'
        self.cell(width * 0.6, 7, f" {val_str}", border=1, ln=False)

    def generate(self, output_path):
        self.add_page()
        
        # MAIN HEADER CARD
        self.set_fill_color(245, 248, 250)
        self.rect(10, 12, 190, 20, 'F')
        self.set_draw_color(180, 180, 180)
        self.rect(10, 12, 190, 20, 'D')
        
        # Title text
        self.set_y(15)
        self.set_font(self.font_name, "B", 18)
        self.set_text_color(31, 78, 121)
        self.cell(110, 14, "  İMALAT İŞ EMRİ", ln=False)
        
        # Date & Shift on the right side of the card
        self.set_font(self.font_name, "", 9)
        self.set_text_color(100, 100, 100)
        import datetime
        now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        self.cell(80, 6, f"Yazdırma Tarihi: {now_str}", align="R", ln=True)
        
        self.ln(8)
        
        # 1. KEY INFO (LOT & MACHINE) - Large bold cards
        self.set_draw_color(31, 78, 121)
        self.set_line_width(0.8)
        
        # LOT CARD
        self.set_fill_color(230, 240, 250)
        self.rect(10, 38, 92, 16, 'DF')
        self.set_y(40)
        self.set_x(12)
        self.set_font(self.font_name, "B", 10)
        self.set_text_color(100, 100, 100)
        self.cell(90, 4, "LOT NO", ln=True)
        self.set_x(12)
        self.set_font(self.font_name, "B", 14)
        self.set_text_color(31, 78, 121)
        self.cell(90, 8, str(self.data.get("LOT", "-")), ln=False)
        
        # MACHINE CARD
        self.set_fill_color(230, 240, 250)
        self.rect(108, 38, 92, 16, 'DF')
        self.set_y(40)
        self.set_x(110)
        self.set_font(self.font_name, "B", 10)
        self.set_text_color(100, 100, 100)
        self.cell(90, 4, "MAKİNE", ln=True)
        self.set_x(110)
        self.set_font(self.font_name, "B", 14)
        self.set_text_color(31, 78, 121)
        self.cell(90, 8, str(self.data.get("MAKİNE", "-")), ln=True)
        
        self.ln(12)
        self.set_line_width(0.2)
        self.set_draw_color(180, 180, 180)
        
        # 2. PRODUCT DETAILS SECTION
        self.draw_section_title("Ürün Teknik Özellikleri")
        
        # Row 1
        self.draw_field("Standart", self.data.get("STANDART"), 63)
        self.draw_field("Kalite", self.data.get("KALİTE"), 63)
        self.draw_field("Marka", self.data.get("MARKA"), 64)
        self.ln(7)
        
        # Row 2
        self.draw_field("Çap (Dia)", self.data.get("ÇAP"), 63)
        self.draw_field("Boy (Len)", self.data.get("BOY"), 63)
        self.draw_field("Diş Adımı", self.data.get("DİŞADIM"), 64)
        self.ln(12)
        
        # 3. QUANTITY DETAILS SECTION
        self.draw_section_title("Miktar ve Ağırlık Bilgileri")
        
        # Row 1
        self.draw_field("Hedef Adet", self.data.get("ADET"), 63)
        self.draw_field("Hedef Ağırlık (KG)", self.data.get("KG"), 63)
        self.draw_field("Kalan Adet", self.data.get("KALAN ADET"), 64)
        self.ln(7)
        
        # Row 2
        self.draw_field("Net Gram", self.data.get("NET GRAM"), 63)
        self.draw_field("Brüt Gram", self.data.get("BRÜT GRAM"), 63)
        self.draw_field("Üretilen Adet", self.data.get("ÜRETİLEN ADET"), 64)
        self.ln(12)
        
        # 4. RAW MATERIAL & COATING SECTION
        self.draw_section_title("Malzeme ve Kaplama Bilgileri")
        
        # Row 1
        self.draw_field("Hammadde", self.data.get("HAMMADDE"), 95)
        self.draw_field("Malzeme (Mat)", self.data.get("MATERIAL"), 95)
        self.ln(7)
        
        # Row 2
        self.draw_field("Kaplama Tipi", self.data.get("KAPLAMATIPI"), 95)
        self.draw_field("Kaplama Standardı", self.data.get("KAPLAMASTANDART"), 95)
        self.ln(12)
        
        # 5. DESCRIPTION AND NOTES SECTION
        self.draw_section_title("Açıklama ve Özel Notlar")
        
        self.set_fill_color(255, 255, 255)
        # Description box
        self.set_font(self.font_name, "B", 9)
        self.cell(40, 8, "Genel Açıklama", border=1, fill=True)
        self.set_font(self.font_name, "", 9)
        desc_val = str(self.data.get("AÇIKLAMA", "-")) if self.data.get("AÇIKLAMA") else "-"
        self.cell(150, 8, f" {desc_val}", border=1, ln=True)
        
        # Notes box
        self.set_font(self.font_name, "B", 9)
        self.cell(40, 8, "Özel Not", border=1, fill=True)
        self.set_font(self.font_name, "", 9)
        note_val = str(self.data.get("NOT", "-")) if self.data.get("NOT") else "-"
        self.cell(150, 8, f" {note_val}", border=1, ln=True)
        
        self.ln(15)
        
        # 6. SIGNATURE BLOCK
        self.set_fill_color(245, 248, 250)
        self.set_font(self.font_name, "B", 9)
        self.set_text_color(50, 50, 50)
        
        # Drawing a signature table
        self.cell(63, 6, " Hazırlayan / Planlama", border=1, fill=True, align="C")
        self.cell(63, 6, " Operatör / Üretim", border=1, fill=True, align="C")
        self.cell(64, 6, " Onay / Kalite Kontrol", border=1, fill=True, align="C")
        self.ln(6)
        
        # Signature empty spaces
        self.cell(63, 20, "", border=1)
        self.cell(63, 20, "", border=1)
        self.cell(64, 20, "", border=1)
        
        # Output
        self.output(output_path)

def generate_pdf(data, output_path):
    pdf = WorkOrderPDF(data)
    pdf.generate(output_path)
