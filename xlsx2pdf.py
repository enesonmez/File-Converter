import os
import sys
import comtypes.client

def xlsx2pdf(_in,_out):
    pdf_format_key = 17
    file_in = os.path.abspath(_in)
    file_out = os.path.abspath(_out)
    excel = comtypes.client.CreateObject('Excel.Application')
    doc = excel.Workbooks.Open(file_in)
    doc.ExportAsFixedFormat(0, file_out, 1, 0)
    doc.Close()
    excel.Quit()

if __name__ == "__main__":
    xlsx2pdf(sys.argv[1],sys.argv[2])