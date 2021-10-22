import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from shutil import copyfile
import codecs
import os

class PDFApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDFBot")
        self.resizable(False, False)
        self.main_screen()

    def main_screen(self):
        self.folder_button = tk.Button(self, text="Select Folder",
                                      command=self.select_folder)
        self.folder_label = tk.Label(self, text='')
        self.generate_fees = tk.Button(self, text='Generate Fees', state=tk.DISABLED,
                                        command=self.generate_pdf)
        self.read_pdf = tk.Button(self, text='Scan PDF',
                                command=self.read_pdf)

        self.folder_button.grid(row=0, column=0)
        self.folder_label.grid(row=0, column=1)
        self.generate_fees.grid(row=1, column=0)
        self.read_pdf.grid(row=1,column=1)

    def switch_button(self):
        if self.folder:
            self.generate_fees['state'] = tk.ACTIVE

    def select_folder(self):
        self.folder = fd.askdirectory(title='Load Folder')
        self.folder_label['text'] = self.folder.split('/')[-1]
        self.switch_button()

    def clear_attributes(self):
        self.folder_label['text'] = ''
        self.folder = None

    def generate_pdf(self):
        fees, other = [], []
        for path, folders, files in os.walk(self.folder):
            for file in files:
                if file.endswith('FEES.pdf'):
                    fees.append(file.split()[-1])
                else:
                    other.append(file.split()[-1])
        if 'XFER.pdf' in set(other):
            self.generate_fee_files(self.folder)
            showinfo(title='Complete', message='PDFs Generated')
            self.clear_attributes()
        else:
            self.generate_crystal(self.folder)

    def generate_fee_files(self, folder):
        save_folder = fd.askdirectory()
        if save_folder:

            beg = sorted([file for file in os.listdir(folder) if file.endswith('XFER.pdf')])

            for i, file in enumerate(beg):
                pdf1 = open(os.path.join(folder, file), 'rb')
                fileparts = file.partition('BANK XFER')
                pdf2filename = fileparts[0] + 'FEES.pdf'
                try:
                    pdf2 = open(os.path.join(folder, pdf2filename), 'rb')

                    pdf1Reader = PdfFileReader(pdf1)
                    pdf2Reader = PdfFileReader(pdf2)

                    pdfWriter = PdfFileWriter()

                    for pageNum in range(pdf1Reader.numPages):
                        pageObj = pdf1Reader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)

                    for pageNum in range(pdf2Reader.numPages):
                        pageObj = pdf2Reader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)

                    pdfOutputFile = open(os.path.join(save_folder, pdf2filename), 'wb')
                    pdfWriter.write(pdfOutputFile) 

                    pdfOutputFile.close()
                    pdf1.close()
                    pdf2.close()
                except FileNotFoundError:
                    copyfile(os.path.join(folder), os.path.join(save_folder, pdf2filename))


    def generate_crystal(self, folder):
        save_folder = fd.askdirectory()

        beg = sorted([file for file in os.listdir(folder) if not file.endswith('FEES.pdf') and file.endswith('.pdf')])
        end = sorted([file for file in os.listdir(folder) if file.endswith('FEES.pdf') and file.endswith('.pdf')])

        for i, file in enumerate(beg):
            if file.split()[0] == end[i].split()[0]:
                pdf1 = open(os.path.join(folder, file), 'rb')
                pdf2 = open(os.path.join(folder, end[i]), 'rb')

                pdf1Reader = PdfFileReader(pdf1)
                pdf2Reader = PdfFileReader(pdf2)

                pdfWriter = PdfFileWriter()

                for pageNum in range(pdf1Reader.numPages):
                    pageObj = pdf1Reader.getPage(pageNum)
                    pdfWriter.addPage(pageObj)

                for pageNum in range(pdf2Reader.numPages):
                    pageObj = pdf2Reader.getPage(pageNum)
                    pdfWriter.addPage(pageObj)

                pdfOutputFile = open(os.path.join(save_folder, beg[i]), 'wb')
                pdfWriter.write(pdfOutputFile)

                pdfOutputFile.close()
                pdf1.close()
                pdf2.close()

    def read_pdf(self):
        ...


if __name__ == "__main__":
    app = PDFApp()
    app.mainloop()