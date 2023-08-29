import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path
# get the filepath
paths = glob.glob("invoices/*.xlsx")
for i in paths:
    # Read the excel file
    df = pd.read_excel(i, sheet_name="Sheet 1")

    # extract the filename from path
    filename = Path(i).stem

    # extract date and invoice number from filename
    invoice_nr, date = filename.split("-")

    # create pdf file
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.add_page()

    # Add number and date
    pdf.set_font(family="times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.:{invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Add a header
    headers = df.columns
    headers = [head.replace("_", " ").title() for head in headers]
    pdf.set_font(family="Times",size=10,style="B")
    pdf.cell(w=30, h=12, txt=f"{headers[0]}", border=1, ln=0)
    pdf.cell(w=70, h=12, txt=f"{headers[1]}", border=1, ln=0)
    pdf.cell(w=35, h=12, txt=f"{headers[2]}", border=1, ln=0)
    pdf.cell(w=30, h=12, txt=f"{headers[3]}", border=1, ln=0)
    pdf.cell(w=30, h=12, txt=f"{headers[4]}", border=1, ln=1)


    # Add rows
    for index,row in df.iterrows():
        pdf.set_font(family="times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30,h=12, txt=f"{row['product_id']}", border=1,ln=0)
        pdf.cell(w=70, h=12, txt=f"{row['product_name']}", border=1, ln=0)
        pdf.cell(w=35, h=12, txt=f"{row['amount_purchased']}", border=1, ln=0)
        pdf.cell(w=30, h=12, txt=f"{row['price_per_unit']}", border=1, ln=0)
        pdf.cell(w=30, h=12, txt=f"{row['total_price']}", border=1, ln=1)



    # Add total sum row
    total_sum = df['total_price'].sum()
    pdf.cell(w=30, h=12, txt="", border=1, ln=0)
    pdf.cell(w=70, h=12, txt="", border=1, ln=0)
    pdf.cell(w=35, h=12, txt="", border=1, ln=0)
    pdf.cell(w=30, h=12, txt="", border=1, ln=0)
    pdf.cell(w=30, h=12, txt=f"{total_sum}", border=1, ln=1)


    # Add footer(total price and company name)
    pdf.set_font(family="Times",size=14,style="B")
    pdf.cell(w=30,h=10,txt=f"The total due amount is {total_sum} Euros.",ln=1)
    pdf.cell(w=27,h=10,txt='PythonHow',ln=0)
    pdf.image(w=12,name="pythonhow.png")

    # create output
    pdf.output(f"pdffiles/{filename}.pdf")
