from fpdf import FPDF
from typing import List, Sequence, Optional

class PowerPDF(FPDF):
    def __init__(
        self,
        title: str,
        period: str = "",
        issued: str = "",
        orientation="P",
        unit="mm",
        format="A4",
        margin=12,
    ):
        super().__init__(orientation=orientation, unit=unit, format=format)
        self.title_text = title
        self.period_text = period
        self.issued_text = issued

        # ฟอนต์ไทย
        self.add_font("TH", "",   "fonts/THSarabunNew.ttf",        uni=True) 
        self.add_font("TH", "B",  "fonts/THSarabunNew-Bold.ttf",   uni=True) 
        self.add_font("TH", "I",  "fonts/THSarabunNew-Italic.ttf", uni=True)  
        self.set_font("TH", size=12)

        # margin + auto page break
        self.set_auto_page_break(auto=True, margin=margin)
        self.set_margins(margin, margin, margin)

        # ค่าสี/เส้นเริ่มต้น
        self.set_draw_color(154, 160, 166)  # เส้นกรอบเทา
        self.set_line_width(0.3)

    # ====== Header / Footer ======
    def header(self):
        self.image('images/tbkk-logo.png', x=10, y=3, w=24)
        self.image('images/tbkGroup-logo.png', x=170, y=0, w=30)
        # self.image('images/tbkGroup-logo.png', )
        self.set_font("TH", "B", size=20)
        self.set_text_color(31, 31, 31)
        self.text(35, 10, "บริษัท ทีบีเคเค ( ประเทศไทย ) จำกัด")
        self.text(35, 18, "TBKK ( Thailand ) Co., Ltd.")

        self.set_font("TH", "B", size=17)
        self.set_xy(120, 18)
        self.multi_cell(80, 6, self.title_text, align="C")
        self.ln(1)

        self.set_font("TH", size=12)
        self.set_text_color(95, 99, 104)

        self.text(12, 28, "สำนักงานใหญ่ 700/1017 หมู่ที่ 9 ตำบลมาบโป่ง อำเภอพานทอง จังหวัดชลบุรี 20160")
        self.text(12, 34, "Head Office 700/1017, Moo 9, TB.Mappong, AP.Panthong, Cholburi 20160")
        self.text(12, 40, "Tel 66(0)38 109 360-7 | Fax 66(0)38 109 368")

        self.text(125, 36, f"การใช้ไฟฟ้า {self.period_text}")
        self.text(125, 41, f"บิลออกให้วันที่ {self.issued_text}")
        self.ln(14)
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-10)
        self.set_font("TH", size=9)
        self.set_text_color(95, 99, 104)
        self.cell(0, 6, f"หน้า {self.page_no()}", align="R")

    # ====== Helpers ======
    def add_section_title(self, text: str):
        self.set_font("TH", size=12)
        self.set_fill_color(245, 247, 250)
        self.cell(0, 8, text, ln=1, fill=True)
        self.ln(1)

    def table(
        self,
        headers: Sequence[str],
        rows: List[Sequence],
        col_widths: Optional[Sequence[float]] = None,
        aligns: Optional[Sequence[str]] = None,
        header_fill=(232, 240, 254),
        header_text=(31, 31, 31),
        grid=True,
    ):
        """
        วาดตารางอย่างง่าย
        - headers: รายชื่อหัวคอลัมน์
        - rows: list ของแต่ละแถว (list/tuple)
        - col_widths: ความกว้างคอลัมน์ (ถ้าไม่ใส่จะหารจากหน้า)
        - aligns: การจัดแนวต่อคอลัมน์ เช่น ["L","R","R"]
        """
        ncols = len(headers)
        page_width = self.w - self.l_margin - self.r_margin
        if not col_widths:
            col_widths = [page_width / ncols] * ncols
        if not aligns:
            aligns = ["L"] * ncols

        # --- Header row ---
        self.set_font("TH", size=11)
        self.set_fill_color(*header_fill)
        self.set_text_color(*header_text)
        for i, h in enumerate(headers):
            self.cell(col_widths[i], 8, str(h), border=1, align="C", fill=True)
        self.ln(8)
        self.set_text_color(0, 0, 0)

        # --- Body rows ---
        self.set_font("TH", size=11)
        for r in rows:
            # ตรวจ page-break: ถ้าใกล้ขอบล่าง ให้ขึ้นหน้าใหม่และวาดหัวตารางซ้ำ
            if self.get_y() > self.h - self.b_margin - 10:
                self.add_page()
                # วาด header ซ้ำ
                self.set_fill_color(*header_fill)
                self.set_text_color(*header_text)
                for i, h in enumerate(headers):
                    self.cell(col_widths[i], 8, str(h), border=1, align="C", fill=True)
                self.ln(8)
                self.set_text_color(0, 0, 0)

            for i, v in enumerate(r):
                align = aligns[i] if i < len(aligns) else "L"
                txt = str(v)
                self.cell(col_widths[i], 8, txt, border=1, align=align)
            self.ln(8)

        # เส้นกริด = ใช้ border=1 ไปกับทุก cell แล้ว (grid=True)
        # ถ้าต้องเส้นรอบนอกหนากว่า:
        if grid:
            x0 = self.get_x() - sum(col_widths)
            y0 = self.get_y() - 8 * (len(rows)) - 8  # รวม header
            self.set_line_width(0.7)
            self.rect(x0, y0, sum(col_widths), 8 * (len(rows) + 1))
            self.set_line_width(0.3)