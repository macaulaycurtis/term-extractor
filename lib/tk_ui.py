import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from pathlib import Path
from unicodedata import east_asian_width
from configparser import ConfigParser
from lib.text_extractor import TextExtractor
from lib.substring_analyser import SubstringAnalyser
from threading import Thread
from win32com import client

class GUI(tk.Frame):

    def __init__(self, root):
        tk.Frame.__init__(self, root)
        self.pack(fill='both', expand=True)
        self.root = root
        self.setup_attributes()
        self.setup_bindings()
        self.create_widgets()
        self.create_menus()
        self.setup_context_menu()
        self.setup_window()

    def setup_attributes(self):
        self.config_parser = ConfigParser()
        self.config_parser.read('config.ini', encoding='utf-8-sig')
        self.config = self.config_parser['USER']

        self.last_loc = self.config['input_path']
        self.files = []

        # OPTION VARIABLES #
        self.spaced = tk.BooleanVar()
        self.spaced.set(self.config.getboolean('spaced'))
        self.spaced.trace('w', lambda *args: self.change_option('spaced', self.spaced))
        self.min_occurrences = tk.IntVar()
        self.min_occurrences.set(self.config.getint('min_occurrences'))
        self.min_occurrences.trace('w', lambda *args: self.change_option('min_occurrences', self.min_occurrences))
        self.min_length = tk.IntVar()
        self.min_length.set(self.config.getint('min_length'))
        self.min_length.trace('w', lambda *args: self.change_option('min_length', self.min_length))

    def setup_bindings(self):
        self.root.bind('<Control-o>', self.open)
        self.root.bind('<Control-q>', self.exit)
        self.root.bind('<Control-w>', self.exit)

    def create_widgets(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # BUTTON FRAME #
        button_frame = tk.Frame(self)
        button_frame.grid(row=0, column=0, sticky='ew')
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        
        open_button = tk.Button(button_frame, text='Add file(s)...', command=self.open)
        open_button.grid(row=0, column=0, sticky='ew')
        self.go_button = tk.Button(button_frame, text='Start', command=self.execute, state='disabled')
        self.go_button.grid(row=0, column=1, sticky='ew')

        # OPTIONS FRAME #
        options_frame = tk.Frame(button_frame)
        options_frame.grid(row=1, column=0, columnspan=2, sticky='ew')
        spaced_toggle = tk.Checkbutton(options_frame, text='Spaced text (Non-Japanese)', padx=2, variable=self.spaced)
        spaced_toggle.grid(row=0, column=0)
        occ_selector = tk.Spinbox(options_frame, from_=1, to=100, increment=1
                              , textvariable=self.min_occurrences
                              , exportselection=True, width=3
                              , state='readonly', readonlybackground='white')
        occ_selector.grid(row=0, column=1)
        occ_label = tk.Label(options_frame, text='Minimum occurrences')
        occ_label.grid(row=0, column=2)
        len_selector = tk.Spinbox(options_frame, from_=1, to=100, increment=1
                              , textvariable=self.min_length
                              , exportselection=True, width=3
                              , state='readonly', readonlybackground='white')
        len_selector.grid(row=0, column=3)
        len_label = tk.Label(options_frame, text='Minimum length')
        len_label.grid(row=0, column=4)
        

        # FILE LIST #
        input_frame = tk.Frame(self)
        input_frame.grid(row=1, column=0, sticky='nesw')
        input_frame.grid_columnconfigure(0, weight=1)
        input_frame.grid_rowconfigure(1, weight=1)
        
        input_label = tk.Label(input_frame, text='Files:')
        input_label.grid(row=0, column=0, sticky='w')

        self.input_box = tk.Listbox(input_frame
                                    , listvariable=tk.StringVar()
                                    , activestyle='none'
                                    , selectmode='extended')
        self.input_box.grid(row=1, column=0, sticky='nesw')
        self.input_box.bind('<Delete>', self.delete)


        # NAVIGATION #
        xscroll = tk.Scrollbar(input_frame, orient='horizontal')
        xscroll.grid(row=2,column=0, sticky='ew')
        self.input_box.configure(xscrollcommand=xscroll.set)
        xscroll['command'] = self.input_box.xview
        yscroll = tk.Scrollbar(input_frame, orient='vertical')
        yscroll.grid(row=1,column=1, sticky='ns')
        self.input_box.configure(yscrollcommand=yscroll.set)
        yscroll['command'] = self.input_box.yview
        size_grip = ttk.Sizegrip(input_frame)
        size_grip.grid(column=1, row=2, sticky='se')

        # PROGRESS BAR #
        input_frame.grid_rowconfigure(1, weight=1)
        self.progress_bar = ttk.Progressbar(input_frame, mode='indeterminate', orient='horizontal')
        self.progress_bar.grid(column=0, row=3, sticky='ew', columnspan=2)

    def create_menus(self):
        menu_bar = tk.Menu(self)

        # FILE MENU #
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='File', menu=file_menu)
        file_menu.add_command(label='Add file(s)...', command=self.open)
        file_menu.add_command(label='Exit', command=self.exit)

        # HELP MENU #
        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='Help', menu=help_menu)
        help_menu.add_command(label='Licence', command=self.licence)

        self.root.config(menu=menu_bar)

    def setup_window(self):
        self.root.title(self.config['title'])
        self.root.minsize(self.config['minimum_x'], self.config['minimum_y'])
        self.root.update()

    def execute(self):
        if len(self.files) == 0:
            return
        for c in self.files[0]['text'][:25]:
            if (east_asian_width(c) in 'FW') and (self.spaced.get()):
                spaced = messagebox.askyesno(
                    title='Check language'
                    , message='Text appears to be unspaced (Japanese). Set to unspaced processing?')
                self.spaced.set(not spaced)
                break
            elif (east_asian_width(c) in 'HNNa') and (c != '\n') and (not self.spaced.get()):
                spaced = messagebox.askyesno(
                    title='Check language'
                    , message='Text appears to be spaced (not Japanese). Set to spaced processing?')
                self.spaced.set(spaced)
                break
        t = Thread(target=self.create_sa)
        t.start()

    def create_sa(self):
        self.go_button.config(state='disabled')
        self.progress_bar.start()
        self.sa = SubstringAnalyser(spaced=self.spaced.get()
                                , min_occurrences=self.min_occurrences.get()
                                , min_length=self.min_length.get())
        self.sa.load(list(
            [(f['path'].name, f['text']) for f in self.files]
            ))
        self.sa.load_common()
        self.progress_bar.stop()
        self.save()
        

    def setup_context_menu(self):
        
        def select_and_context(e):
            context_menu = tk.Menu(tearoff=0)
            context_menu.add_command(label='Add', command=self.open)
            context_menu.add_command(label='Remove', command=self.delete)
            nearest = self.input_box.nearest(e.y_root - self.input_box.winfo_rooty())
            if nearest not in self.input_box.curselection():
                self.input_box.selection_clear(0, 'end')
                self.input_box.selection_set(nearest)
            if self.input_box.curselection() == ():
                context_menu.entryconfig(1, state='disabled')
            try:
                context_menu.tk_popup(e.x_root+40, e.y_root+10,entry="0")
            finally:
                context_menu.grab_release()
    
        self.input_box.bind('<Button-3>', select_and_context)
        self.input_box.bind('<Delete>', self.delete)

    def open(self, event=None):
        filepath = filedialog.askopenfilename(multiple=True, initialdir=(self.last_loc))
        if filepath == '': return
        if isinstance(filepath, str):
            filepath = (filepath,)
        for f in filepath:
            f = Path(f)
            self.extract_text(f)
            self.last_loc = str(f.root)

    def extract_text(self, f, password=''):
        try:
            text = str(TextExtractor(f, password))
            self.files.append({'path' : f, 'text' : text})
            self.input_box.insert('end', '{} ({})'.format(f.name, f))
            self.go_button.config(state='active')
        except Exception as e:
            if e.args[0] == 'Incorrect password.':
                password = simpledialog.askstring('Incorrect password', 'Enter password for {}:'.format(f.name))
                if password == None:
                    return
                self.extract_text(f, password)
            else:
                messagebox.showerror('Error opening file', 'Error opening file "{}":\n{}'.format(f.name, e))

    def delete(self, event=None):
        while self.input_box.curselection() != ():
            i = self.input_box.curselection()[0]
            self.input_box.delete(i)
            self.files.pop(i)
            if len(self.files) == 0:
                self.go_button.config(state='disabled')

    def save(self, event=None):
        filepath = filedialog.asksaveasfilename(initialdir=(self.config['output_path']), initialfile='extracted_terms.xlsx'
                                                , defaultextension='.xlsx', filetypes=(('xlsx', '*.xlsx'),))
        if filepath == '': return
        filepath = Path(filepath)
        retry = True
        while retry:
            try:
                self.progress_bar.start()
                self.sa.save_output(filepath)
                retry = False
            except PermissionError:
                if messagebox.askretrycancel('Permission denied', 'File is open in another program.'):
                    continue
                else:
                    return
            finally:
                self.progress_bar.stop()
                self.go_button.config(state='active')
        yesopen = messagebox.askyesno(title='Output', message='Output saved to {}. Open in Excel?'.format(filepath.name))
        if yesopen:
            excel = client.DispatchEx('Excel.Application')
            excel.Visible = 1
            wb = excel.Workbooks.Open(filepath)
            
    def change_option(self, option, var):
        '''Saves changes to options to the config file.'''
        self.config_parser.set('USER', option, str(var.get()))
        with Path('config.ini').open('w', encoding='utf-8') as conf:
            self.config_parser.write(conf)

    def licence(self, event=None):
        licence = str
        with Path('licence.txt').open('r', encoding='utf-8-sig') as f:
            licence = f.read()
        messagebox.showinfo(title='Licence', message=licence)

    def exit(self, event=None):
        self.root.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    ui = GUI(root)
    root.mainloop()
