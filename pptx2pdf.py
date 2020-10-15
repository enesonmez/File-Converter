import os
import sys
import comtypes.client

def PPTXtoPDF(_in,_out, formatType = 32):
    file_in = os.path.abspath(_in)
    file_out = os.path.abspath(_out)
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

    pdf = powerpoint.Presentations.Open(file_in, WithWindow=False)
    pdf.SaveAs(file_out, formatType) # formatType = 32 for ppt to pdf
    pdf.Close()
    powerpoint.Quit()

if __name__ == "__main__":
    PPTXtoPDF(sys.argv[1],sys.argv[2])