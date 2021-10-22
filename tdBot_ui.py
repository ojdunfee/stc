import tkinter as tk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo
import os
import tdBot


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CCBot")
        self.geometry("227x50")
        self.resizable(False, False)
        self.main_screen()
    
    def main_screen(self):
        self.statement_dir_button = tk.Button(self, text='...',
                                                command=self.choose_card_directory)
        self.dir_label = tk.Label(self, text='Statement Folder')
        self.dir_entry = tk.Label(self, bg='white', justify=tk.RIGHT,
                                  state=tk.DISABLED, width=15)
        self.generate_button = tk.Button(self, text='Generate', width=31,
                                        state=tk.DISABLED, command=self.generate_report)
    
        self.dir_label.grid(row=0, column=0)
        self.statement_dir_button.grid(row=0, column=2)
        self.dir_entry.grid(row=0, column=1)

        self.generate_button.grid(row=1, columnspan=3)
    
    def choose_card_directory(self):
        self.card_dir = fd.askdirectory()
        if self.card_dir:
            self.dir_entry['text'] = self.card_dir.split('/')[-1]
            self.generate_button['state'] = tk.ACTIVE

    def clear_selected_dir(self):
        self.dir_entry['text'] = ''
        self.card_dir = None

    def generate_report(self):
        filename = tdBot.get_save_file(self.card_dir)
        savefile = fd.asksaveasfilename(
            filetypes=[('Excel', '*.xlsx')],
            initialfile=filename
        )
        if savefile:
            tdBot.generate_journal(self.card_dir, savefile)
            showinfo(title='Complete', message='{} Generated'.format(savefile.split('/')[-1]))
            self.clear_selected_dir()

if __name__ == "__main__":
    app = App()
    app.mainloop()