import pandas as pd
import glob # glob used to create lists
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    
    # Take the name of the file and uses separator to break it in two
    filename = Path(filepath).stem
    invoice_nr, date=filename.split("-")
    
    pdf.set_font(family="Times", size=16, style="B")
    # Creating the cell to put the invoice number into the PDF file
    pdf.cell(w=50, h=8, txt=f"invoice nr.{invoice_nr}", ln=1)
    
    pdf.set_font(family="Times", size=16, style="B")
    # Creating the cell to put the date into the PDF file
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    
    
    # Reading the excel file, specifing the Sheet number to read from
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Adding a table header
    column = df.columns
    # Using list comprehension to remove underscore and capitalize table heder names
    column = [item.replace("_", " ").title() for item in column]
    pdf.set_font(family="Times", size=10, style="BI")
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt=column[0], border=1)
    pdf.cell(w=70, h=8, txt=column[1], border=1)
    pdf.cell(w=30, h=8, txt=column[2],border=1)
    pdf.cell(w=30, h=8, txt=column[3], border=1)
    pdf.cell(w=30, h=8, txt=column[4], border=1, ln=1)        
    
    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
    
    # Adding total sum
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    
    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")    
    pdf.cell(w=50, h=8, txt=f"The total price is {total_sum}", ln=1)
    
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")    
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)
    
    # Putting output inside the loop to create seperate pages
    pdf.output(f"PDFs/{filename}.pdf")
    
    # ROWS and COLUMNS are the properties of the dataframe (from excel file)
