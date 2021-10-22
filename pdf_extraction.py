from typing import Text
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import pandas as pd
import datetime
import os

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()
    
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

def generate_fee_master(dir):
    frames = list()
    year = str(datetime.datetime.today().year)
    for filepath in os.listdir(dir):
        if filepath.endswith('FEES.pdf'):
            filepath = os.path.join(dir, filepath)
            txt = convert_pdf_to_txt(filepath)
            df = {'Date': list(), 'Amount': list(), 'File': list()}
            for x in txt.split('\n'):
                if x.endswith(year):
                    df['Date'].append(year)
                if x.startswith('$'):
                    try:
                        amt, file_no = x.split()
                        df['Amount'].append(float(amt.strip('$').replace(',', '')))
                        df['File'].append(file_no)
                    except ValueError:
                        ...
            frames.append(pd.DataFrame(df))
    return pd.concat(frames, ignore_index=True)

if __name__ == "__main__":
    df = generate_fee_master(r'/Users/owendunfee/Developer/Python/stc/Fees xfers 10_19_2021')
    print(df)