import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    pdf = FPDF(orientation='P', unit="in", format='letter')
    pdf.add_page()
    
    filename = Path(filepath).stem
    invoice_no, inv_date = filename.split("-")
    
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=2, h=.3, txt=f"Invoice Number: {invoice_no}", ln=1)
    
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=2, h=.3, txt=f"Date: {inv_date}")
 
    
    
    
    
    pdf.output(f"pdf_output/Inv{invoice_no}.pdf")