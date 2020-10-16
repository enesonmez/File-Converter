import os
import sys
from fpdf import FPDF


def TXTtoPDF(_in,_out):
    file_in = os.path.abspath(_in)
    file_out = os.path.abspath(_out)
    file = open(file_in,'r')

    pdf = FPDF(format='A4')
    pdf.add_page()

    for text in file:
        if len(text) <= 30:#title
            pdf.set_font('Times', 'B', size=15)
            pdf.cell(w=200, h=10, txt=text, ln=1, align='C')
        else:#paragraph
            pdf.set_font('Times', size=12)
            pdf.multi_cell(w=0, h=10, txt=text, align='L')
    
    pdf.output(file_out)
    

if __name__ == "__main__":
    TXTtoPDF(sys.argv[1],sys.argv[2])