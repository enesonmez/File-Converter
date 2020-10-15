import os
import sys
import comtypes.client

def docx2pdf(_in,_out):
    pdf_format_key = 17
    file_in = os.path.abspath(_in)
    file_out = os.path.abspath(_out)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(file_in)
    doc.SaveAs(file_out, FileFormat=pdf_format_key)
    doc.Close()
    word.Quit()

if __name__ == "__main__":
    docx2pdf(sys.argv[1],sys.argv[2])