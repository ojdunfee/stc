import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo
import os
import cash


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sheetbot")
        self.geometry('220x78')
        self.resizable(False, False)
        self.main_screen()

    def main_screen(self):
        self.cash_file = tk.Button(self, text='...',
                                  command=self.choose_cash_file)
        self.cash_label = tk.Label(self, text="Cash Receipts")
        self.cash_entry = tk.Label(self, bg='white', justify=tk.RIGHT,
                                   state=tk.DISABLED, width=15)

        self.fee_file = tk.Button(self, text='...',
                                  command=self.choose_fee_file)
        self.fee_label = tk.Label(self, text="TD Fees")
        self.fee_entry = tk.Label(self, bg='white', justify=tk.RIGHT,
                                  state=tk.DISABLED, width=15)
        
        self.generate = tk.Button(self, text='Generate', width=30,
                                  state=tk.DISABLED, command=self.generate_report)

        self.cash_label.grid(row=0, column=0)
        self.cash_file.grid(row=0, column=2)
        self.cash_entry.grid(row=0, column=1)

        self.fee_label.grid(row=1, column=0)
        self.fee_entry.grid(row=1, column=1)
        self.fee_file.grid(row=1, column=2)

        self.generate.grid(row=2, columnspan=3)

    def switch_button(self):
        if self.cash_filename and self.fee_filename:
            self.generate['state'] = tk.ACTIVE

    def choose_cash_file(self):
        self.cash_filename = fd.askopenfilename(title="Open File",
                                        initialdir=os.getcwd(),
                                        filetypes=[('Excel', '*.xls *.xlsx')])
        if self.cash_filename:
            file = self.cash_filename.split('/')[-1]
            self.cash_entry['text'] = file
            try:
                self.switch_button()
            except AttributeError:
                pass

    def choose_fee_file(self):
        self.fee_filename = fd.askopenfilename(title="Open File",
                                    initialdir=os.getcwd(),
                                    filetypes=[('Excel', '*.xls *.xlsx')])
        if self.fee_filename:
            file = self.fee_filename.split('/')[-1]
            self.fee_entry['text'] = file
            try:
                self.switch_button()
            except AttributeError:
                pass

    def generate_report(self):
        fname = cash.get_filename(self.cash_filename)
        errors = cash.check_sheet(self.cash_filename, self.fee_filename)
        revisions = cash.get_revisions(self.cash_filename)
        if errors or revisions:
            errors += revisions
            root = tk.Tk()
            root.geometry('250x250')
            root.title('Errors')
            sbar = tk.Scrollbar(root)
            listbox = tk.Listbox(root)
            sbar.config(command=listbox.yview)
            listbox.config(yscrollcommand=sbar.set)
            sbar.pack(side=tk.RIGHT, fill=tk.Y)
            listbox.pack(side=tk.LEFT, expand=tk.YES, fill=tk.BOTH)
            [listbox.insert(i, errors[i]) for i in range(len(errors))]
        else:
            save_file = fd.asksaveasfilename(
                filetypes=[('Excel', '*.xlsx')],
                initialfile=fname)
            try:
                cash.generate_journal(cash_file=self.cash_filename, save_file=save_file,  fee_sheet=self.fee_filename)
                showinfo(title='Complete', message='{} Generated'.format(fname))
                self.clear_selected_files()
            except FileNotFoundError:
                ...

    def clear_selected_files(self):
        self.cash_entry['text'] = ''
        self.fee_entry['text'] = ''
        self.cash_file = None
        self.fee_file = None


if __name__ == "__main__":
    app = App()
    app.mainloop()