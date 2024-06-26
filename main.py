import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    
    pdf = FPDF(orientation='P', unit="in", format='letter')
    pdf.add_page()
    
    # Add company name and logo
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=2, h=.1, txt="", ln=1)
    pdf.cell(w=6, h=.4, txt="Voima Technical Services LLC", align='C', ln=1)
    pdf.image("voima_logo.png", w = 0.75, x=3, y=.65)

    # Get invoice number and date from filename
    filename = Path(filepath).stem
    invoice_no, inv_date = filename.split("-")
    
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=2, h=.3, txt="", ln=1)
    pdf.cell(w=2, h=.3, txt=f"Invoice Number: {invoice_no}", ln=1)
    
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=2, h=.3, txt=f"Date: {inv_date}", ln=2)
 
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    header_columns = df.columns
    header_columns = [item.replace("_", " ").title() for item in header_columns]
    pdf.set_font(family="Times", size=9, style="B")
    pdf.set_text_color(0,0,0)
    pdf.cell(w=1, h=.3, txt=header_columns[0],border=1)
    pdf.cell(w=2, h=.3, txt=header_columns[1], border=1)
    pdf.cell(w=1, h=.3, txt="Quantity", border=1, align='C')
    pdf.cell(w=1, h=.3, txt=header_columns[3], border=1, align='R')
    pdf.cell(w=1, h=.3, txt=header_columns[4], border=1, ln=1, align='R')

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(200,0,50)
        pdf.cell(w=1, h=.3, txt=str(row["product_id"]),border=1)
        pdf.cell(w=2, h=.3, txt=str(row["product_name"]), border=1)
        pdf.cell(w=1, h=.3, txt=str(row["amount_purchased"]), border=1, align='C')
        pdf.cell(w=1, h=.3, txt="${:,.2f}".format(row["price_per_unit"]), border=1, align='R')
        pdf.cell(w=1, h=.3, txt="${:,.2f}".format(row["total_price"]), border=1, ln=1, align='R')
    
    price_sum = df["total_price"].sum()

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0,0,0)
    pdf.cell(w=1, h=.3, txt="",border=1)
    pdf.cell(w=2, h=.3, txt="", border=1)
    pdf.cell(w=1, h=.3, txt="", border=1, align='C')
    pdf.cell(w=1, h=.3, txt="Total", border=1, align='R')
    pdf.cell(w=1, h=.3, txt=f"${price_sum:,.2f}", border=1, ln=1, align='R')

    # Add total cost sentence    
    pdf.cell(w=6, h=.3, txt="",ln = 1)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=6, h=.5, txt=f"The total amount due is ${price_sum:,.2f}",ln = 2)



    pdf.output(f"pdf_output/Inv{invoice_no}.pdf")