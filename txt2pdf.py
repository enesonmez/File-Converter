#DejaVuSansCondensed.ttf
#DejaVuSansCondensed.pkl
#DejaVuSansCondensed.cw127.pkl
#The above files help this file run. The purpose of these files is to use unicode characters.
#If you do not want to use these files, activate the comment line and delete the add_font line. Replace 'Dejavu' in set_font with 'Times'.
import os
import sys
from fpdf import FPDF

def TXTtoPDF(_in,_out):
    file_in = os.path.abspath(_in)
    file_out = os.path.abspath(_out)
    file = open(file_in,'r',encoding='utf-8',errors='ignore')
    
    pdf = FPDF(format='A4')
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
    for text in file:
        #text=text.encode('latin1','ignore').decode('iso-8859-1')
        text = u'' + text
        if len(text) <= 30:#title
            pdf.set_font('DejaVu', '', size=15)
            pdf.multi_cell(w=200, h=10, txt=text, align='C')
        else:#paragraph
            pdf.set_font('DejaVu', size=12)
            pdf.multi_cell(w=0, h=10, txt=text, align='L')
    
    pdf.output(file_out)
    

if __name__ == "__main__":
    TXTtoPDF(sys.argv[1],sys.argv[2])