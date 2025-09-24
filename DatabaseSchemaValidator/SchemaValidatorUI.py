import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import importlib
import os
import glob
from PIL import Image, ImageTk

# Try to load a true black theme (fallback to azure-dark if not found)
def try_load_theme(root):
    try:
        style = ttk.Style()
        root.tk.call('source', os.path.join(os.path.dirname(__file__), 'SchemaValidatorUI_style_black.tcl'))
        style.theme_use('black')
    except Exception:
        try:
            style = ttk.Style()
            root.tk.call('source', os.path.join(os.path.dirname(__file__), 'SchemaValidatorUI_style.tcl'))
            style.theme_use('azure-dark')
        except Exception:
            pass

def import_config():
    import importlib.util
    import sys
    config_path = os.path.join(os.path.dirname(__file__), 'config.py')
    spec = importlib.util.spec_from_file_location('config', config_path)
    config = importlib.util.module_from_spec(spec)
    sys.modules['config'] = config
    spec.loader.exec_module(config)
    return config

def run_validation():
    import importlib.util
    import sys
    script_path = os.path.join(os.path.dirname(__file__), 'SchemaValidatior.py')
    spec = importlib.util.spec_from_file_location('SchemaValidatior', script_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules['SchemaValidatior'] = module
    spec.loader.exec_module(module)
    module.main()

def find_latest_reports():
    reports_dir = os.path.join(os.path.dirname(__file__), 'SchemaValidationReports')
    if not os.path.exists(reports_dir):
        return []
    files = glob.glob(os.path.join(reports_dir, '*.xlsx'))
    files.sort(key=os.path.getmtime, reverse=True)
    return files[:10]

def load_icon(name, size=(32, 32)):
    path = os.path.join(os.path.dirname(__file__), 'assets', name)
    try:
        img = Image.open(path).resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None

class ConfigUI(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style='Black.TFrame')
        self.parent = parent
        self.config = import_config()
        # Remove unsupported background option for ttk.Frame
        self.create_widgets()

    def create_widgets(self):
        self.configure(style='Black.TFrame')
        # Outer frame with border
        outer = tk.Frame(self, bg='#2e2d2d', highlightbackground='#4a3d0e', highlightthickness=2, bd=0)
        outer.pack(fill='both', expand=True, padx=0, pady=0)  # Remove extra padding
        title = tk.Label(outer, text='Database Connection Configuration', font=(None, 13, 'bold'), fg='#fff', bg='#2e2d2d', anchor='w')
        title.pack(pady=(10, 5), anchor='w')
        # Add space before SQL Server Config
        tk.Frame(outer, height=6, bg='#2e2d2d').pack()
        # SQL Server Config section (no border)
        sql_frame = tk.Frame(outer, bg='#2e2d2d')
        sql_frame.pack(fill='x', padx=0, pady=0)
        sql_title = tk.Label(sql_frame, text='SQL Server Config', fg='#bbb', bg='#2e2d2d', font=(None, 11, 'bold'))
        sql_title.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 4))
        self.sql_vars = {}
        # Windows Auth checkbox
        self.windows_auth_var = tk.BooleanVar(value=self.config.SQL_SERVER_CONFIG.get('windows_auth', False))
        def on_windows_auth_toggle():
            wa = self.windows_auth_var.get()
            # Enable/disable username/password fields
            for key in ['username', 'password']:
                entry = self.sql_entries.get(key)
                if entry:
                    entry.config(state='disabled' if wa else 'normal')
        wa_cb = tk.Checkbutton(sql_frame, text='Use Windows Authentication', variable=self.windows_auth_var, bg='#2e2d2d', fg='#fff', selectcolor='#2e2d2d', command=on_windows_auth_toggle)
        wa_cb.grid(row=1, column=0, columnspan=2, sticky='w', padx=5, pady=(2, 6))
        self.sql_entries = {}
        row_offset = 2
        for i, (key, val) in enumerate(self.config.SQL_SERVER_CONFIG.items()):
            if key == 'windows_auth':
                continue  # Already handled
            tk.Label(sql_frame, text=key.capitalize(), fg='#fff', bg='#2e2d2d').grid(row=i+row_offset, column=0, sticky='w', padx=5, pady=2)
            var = tk.StringVar(value=val)
            entry = tk.Entry(sql_frame, textvariable=var, width=40, fg='#fff', bg='#383838', insertbackground='#fff',
                             highlightbackground='#4a3d0e', highlightcolor='#4a3d0e', highlightthickness=1, relief='flat',
                             show='*' if 'password' in key.lower() else '')
            # Disable database and driver fields, but keep color the same
            if key.lower() in ['database', 'driver']:
                entry.config(state='normal', disabledbackground='#383838', disabledforeground='#fff')
                entry.config(state='disabled')
            entry.grid(row=i+row_offset, column=1, padx=5, pady=2)
            self.sql_vars[key] = var
            self.sql_entries[key] = entry
        # Initial enable/disable
        on_windows_auth_toggle()
        # Add space before Postgres Config
        tk.Frame(outer, height=6, bg='#2e2d2d').pack()
        # Postgres Config section (no border)
        pg_frame = tk.Frame(outer, bg='#2e2d2d')
        pg_frame.pack(fill='x', padx=0, pady=0)
        pg_title = tk.Label(pg_frame, text='Postgres Config', fg='#bbb', bg='#2e2d2d', font=(None, 11, 'bold'))
        pg_title.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 4))
        self.pg_vars = {}
        for i, (key, val) in enumerate(self.config.POSTGRES_CONFIG.items()):
            tk.Label(pg_frame, text=key.capitalize(), fg='#fff', bg='#2e2d2d').grid(row=i+1, column=0, sticky='w', padx=5, pady=2)
            var = tk.StringVar(value=val)
            entry = tk.Entry(pg_frame, textvariable=var, width=40, fg='#fff', bg='#383838', insertbackground='#fff',
                             highlightbackground='#4a3d0e', highlightcolor='#4a3d0e', highlightthickness=1, relief='flat',
                             show='*' if 'password' in key.lower() else '')
            # Disable database and port fields, but keep color the same
            if key.lower() in ['database', 'port']:
                entry.config(state='normal', disabledbackground='#383838', disabledforeground='#fff')
                entry.config(state='disabled')
            entry.grid(row=i+1, column=1, padx=5, pady=2)
            self.pg_vars[key] = var
        # Add space before DB List
        tk.Frame(outer, height=6, bg='#2e2d2d').pack()
        # DB List section (no border)
        db_frame = tk.Frame(outer, bg='#2e2d2d')
        db_frame.pack(fill='x', padx=0, pady=0)
        db_title = tk.Label(db_frame, text='Database List (comma separated)', fg='#bbb', bg='#2e2d2d', font=(None, 11, 'bold'))
        db_title.pack(anchor='w', pady=(0, 4))
        self.db_list_var = tk.StringVar(value=','.join(self.config.DB_LIST))
        tk.Entry(db_frame, textvariable=self.db_list_var, width=50, fg='#fff', bg='#383838', insertbackground='#fff',
                 highlightbackground='#4a3d0e', highlightcolor='#4a3d0e', highlightthickness=1, relief='flat').pack(padx=5, pady=2)
        # Save Config button (styled like validation button)
        self.save_btn = tk.Button(outer, text='Save Config', command=self.save_config, fg='#fff', bg='#8c6916',
                                 activebackground='#4a3d0e', highlightbackground='#4a3d0e', highlightcolor='#4a3d0e', highlightthickness=1, relief='flat')
        self.save_btn.pack(pady=15)
        self.status_var = tk.StringVar(value='Ready')
        self.status_label = tk.Label(outer, textvariable=self.status_var, fg='#fff', bg='#2e2d2d', font=(None, 11, 'bold'))
        self.status_label.pack(pady=5)

    def save_config(self):
        sql_conf = {k: v.get() for k, v in self.sql_vars.items()}
        sql_conf['windows_auth'] = self.windows_auth_var.get()
        pg_conf = {k: v.get() for k, v in self.pg_vars.items()}
        db_list = [db.strip() for db in self.db_list_var.get().split(',') if db.strip()]
        config_code = f"""# config.py\n\n# Database connection settings\nSQL_SERVER_CONFIG = {repr(sql_conf)}\n\nPOSTGRES_CONFIG = {repr(pg_conf)}\n\n# List of databases to validate\nDB_LIST = {repr(db_list)}\n"""
        config_path = os.path.join(os.path.dirname(__file__), 'config.py')
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(config_code)
        self.status_var.set('✔ Config saved!')
        self.status_label.configure(fg='#fff', bg='#2d2d2d')
        CustomMessage(self, 'Configuration saved to config.py.', 'Success', 'success')
        self.config = import_config()

class ValidationUI(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style='Black.TFrame')
        self.parent = parent
        self.validation_thread = None
        self.validation_in_progress = False
        self.icons = {
            'check': '\u2714',  # Unicode checkmark
            'excel': load_icon('excel.png', size=(28, 28)),
            'view': load_icon('eye.png', size=(18, 18)),
            'delete': load_icon('delete.png', size=(18, 18))
        }
        self.config = import_config()
        self.create_widgets()

    def create_widgets(self):
        self.configure(style='Black.TFrame')
        outer = tk.Frame(self, bg='#2e2d2d', highlightbackground='#4a3d0e', highlightthickness=2, bd=0)
        outer.pack(fill='both', expand=True, padx=0, pady=0)
        # Use grid for proportional layout
        outer.grid_rowconfigure(0, weight=2, minsize=110)  # Increased height for validation section
        outer.grid_rowconfigure(1, weight=3)
        outer.grid_columnconfigure(0, weight=1)
        # Top: Validate & Generate Report button/status (increased height)
        top_frame = tk.Frame(outer, bg='#2e2d2d')
        top_frame.grid(row=0, column=0, sticky='nsew')
        self.validate_btn = tk.Button(
            top_frame,
            text='  Validate & Generate Report',
            font=(None, 13, 'bold'),
            fg='#fff',
            bg='#8c6916',
            activebackground='#4a3d0e',
            relief='flat',
            padx=18, pady=10,
            compound='left',
            command=self.run_validation_thread
        )
        if self.icons.get('excel'):
            self.validate_btn.config(image=self.icons['excel'])
        self.validate_btn.pack(pady=(10, 4))
        self.status_var = tk.StringVar(value='')
        self.status_label = tk.Label(top_frame, textvariable=self.status_var, fg='#8c6916', bg='#2d2d2d', font=(None, 12, 'bold'))
        self.status_label.pack(pady=(2, 6))
        self.spinner_label = tk.Label(top_frame, text='', fg='#8c6916', bg='#2d2d2d', font=(None, 18, 'bold'))
        self.spinner_label.pack(pady=(0, 4))
        self.spinner_running = False
        # Lower section: results/messages/reports
        lower_frame = tk.Frame(outer, bg='#2e2d2d')
        lower_frame.grid(row=1, column=0, sticky='nsew')
        self.db_results_frame = tk.Frame(lower_frame, bg='#2e2d2d')
        self.db_results_frame.pack(fill='both', expand=True)

    def run_validation_thread(self):
        self.config = import_config()  # Reload config to get latest DB_LIST
        self.status_var.set('Report generation in progress')
        self.status_label.configure(fg='#8c6916')
        self.validate_btn.config(state='disabled')
        for widget in self.db_results_frame.winfo_children():
            widget.destroy()
        # Show spinner
        self.spinner_running = True
        self._animate_spinner()
        t = threading.Thread(target=self._run_validation)
        t.start()
        self.validation_thread = t

    def _animate_spinner(self):
        # Simple text-based spinner animation
        if not self.spinner_running:
            self.spinner_label.config(text='')
            return
        spinner_chars = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        if not hasattr(self, '_spinner_index'):
            self._spinner_index = 0
        self.spinner_label.config(text=spinner_chars[self._spinner_index] + '  Validating...')
        self._spinner_index = (self._spinner_index + 1) % len(spinner_chars)
        self.after(120, self._animate_spinner)

    def _run_validation(self):
        # Only run validation once for all databases
        try:
            run_validation()
            db_list = getattr(self.config, 'DB_LIST', [])
            db_results = [(db, True, '') for db in db_list]
        except Exception as e:
            db_list = getattr(self.config, 'DB_LIST', [])
            db_results = [(db, False, str(e)) for db in db_list]
        self.after(0, lambda: self._show_db_results(db_results))

    def _show_db_results(self, db_results):
        # Clear previous results/messages
        for widget in self.db_results_frame.winfo_children():
            widget.destroy()
        # Remove spinner
        self.spinner_running = False
        self.spinner_label.config(text='')
        # Remove any previous error area
        if hasattr(self, 'error_area') and self.error_area:
            self.error_area.destroy()
            self.error_area = None
        # Collect error messages from db_results
        error_msgs = [err for db, success, err in db_results if not success and err]
        db_list = set(getattr(self.config, 'DB_LIST', []))
        success_dbs = [db for db, success, err in db_results if success and db in db_list]
        # Show a single, centered success message if any DB succeeded
        if success_dbs:
            frame = tk.Frame(self.db_results_frame, bg='#2e2d2d', height=40)
            frame.pack(fill='x', pady=8)
            msg_label = tk.Label(frame, text="Report generated for selected database(s)", font=(None, 11, 'bold'), fg='#fff', bg='#2e2d2d', anchor='center', justify='center')
            msg_label.pack(side='left', padx=(0, 6), expand=True)
            check_label = tk.Label(frame, text=self.icons['check'], font=(None, 15, 'bold'), fg='#4caf50', bg='#2e2d2d', anchor='center', justify='center')
            check_label.pack(side='left')
            frame.pack(anchor='center')
            # Remove success message after 10 seconds
            self.after(10000, lambda: frame.destroy())
        # Refresh reports with a short delay to ensure file is written
        if success_dbs:
            self.after(1000, lambda: getattr(self.winfo_toplevel(), 'refresh_reports', lambda: None)())
        # Show error message with copy button if error(s) exist
        if error_msgs:
            self.error_area = tk.Frame(self.db_results_frame, bg='#2e2d2d', height=40)
            self.error_area.pack(fill='x', pady=(4, 0))
            self.error_area.pack_propagate(False)
            row = tk.Frame(self.error_area, bg='#2e2d2d')
            row.pack(fill='x', pady=4)
            error_label = tk.Label(row, text='Validation failed! Copy the error here.', font=(None, 9), fg='#ff4444', bg='#2e2d2d', anchor='center', justify='center')
            error_label.pack(side='left', padx=(0, 10), expand=True)
            def copy_error():
                self.error_area.clipboard_clear()
                self.error_area.clipboard_append('\n'.join(error_msgs))
            copy_btn = tk.Button(row, text='Copy Errors', font=(None, 9, 'bold'), fg='#fff', bg='#888', activebackground='#aaa', relief='flat', command=copy_error, padx=10, pady=2, bd=0)
            copy_btn.pack(side='left', padx=(0, 12))
            copy_btn.configure(highlightbackground='#bbb', highlightthickness=1)
            copy_btn.configure(borderwidth=0)
            copy_btn.configure(cursor='hand2')
            copy_btn.configure(relief='flat')
            copy_btn.configure(overrelief='ridge')
            copy_btn.configure(bg='#888', fg='#fff')
            copy_btn.configure(activebackground='#aaa', activeforeground='#fff')
            copy_btn.configure(highlightcolor='#fff')
            copy_btn.configure(highlightbackground='#bbb')
            # Rounded corners (simulate with padding)
            copy_btn.configure(padx=10, pady=2)
            # Remove error area after 10 seconds
            self.after(10000, lambda: (self.error_area.destroy() if hasattr(self, 'error_area') and self.error_area else None))
            # Re-enable button immediately
            self.validate_btn.config(state='normal')
            self.status_var.set('')
            self.status_label.configure(fg='#fff')
        else:
            self.status_var.set('')
            self.status_label.configure(fg='#fff')
            self.validate_btn.config(state='normal')
        # Add a short delay before refreshing recent reports to ensure file is written
        if success_dbs:
            self.after(700, lambda: hasattr(self.parent.master, 'refresh_reports') and self.parent.master.refresh_reports())

class AboutUI(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style='Black.TFrame')
        self.create_widgets()

    def create_widgets(self):
        self.configure(style='Black.TFrame')
        outer = tk.Frame(self, bg='#2d2d2d', highlightbackground='#4a3d0e', highlightthickness=2)
        outer.pack(fill='both', expand=True, padx=30, pady=30)
        title = tk.Label(outer, text='About', font=(None, 15, 'bold'), fg='#fff', bg='#2d2d2d', anchor='w')
        title.pack(pady=(10, 5), anchor='w')
        info = (
            'Schema Validator : SQL Server vs PostgreSQL\n'
            'Version: 2025.09.23\n'
            '\n'
            'A modern tool for comparing SQL Server and PostgreSQL schemas,\n'
            'with robust logic for columns, triggers, constraints, and more.\n'
            '\n'
            'Recent updates:\n'
            '- Improved matching for columns and triggers\n'
            '- Modernized UI\n'
            '- Bugfixes and usability improvements\n'
            '\n'
            'For support, contact the development team.'
        )
        tk.Label(outer, text=info, font=(None, 11), fg='#fff', bg='#2d2d2d', justify='left', anchor='w').pack(padx=20, pady=10, anchor='w')

class CustomMessage(tk.Toplevel):
    def __init__(self, parent, message, title, msgtype, ask_yes_no=False):
        super().__init__(parent)
        self.title(title)
        self.configure(bg='#2d2d2d')
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        icon_color = {'success': '#fff', 'error': '#fff', 'warning': '#fff'}.get(msgtype, '#fff')
        tk.Label(self, text=title, font=(None, 14, 'bold'), fg=icon_color, bg='#2d2d2d').pack(pady=(15, 5), padx=20)
        tk.Label(self, text=message, font=(None, 11), fg='#fff', bg='#2d2d2d', wraplength=400, justify='center').pack(pady=(0, 15), padx=20)
        btn_frame = tk.Frame(self, bg='#2d2d2d')
        btn_frame.pack(pady=(0, 15))
        if ask_yes_no:
            yes_btn = tk.Button(btn_frame, text='Yes', width=10, bg='#383838', fg='#fff', activebackground='#2d2d2d', command=self._yes)
            yes_btn.pack(side='left', padx=10)
            no_btn = tk.Button(btn_frame, text='No', width=10, bg='#383838', fg='#fff', activebackground='#2d2d2d', command=self._no)
            no_btn.pack(side='left', padx=10)
            self.wait_window()
        else:
            ok_btn = tk.Button(btn_frame, text='OK', width=12, bg='#383838', fg='#fff', activebackground='#2d2d2d', command=self._ok)
            ok_btn.pack()
            self.wait_window()
    def _ok(self):
        self.destroy()
        self.result = True
    def _yes(self):
        self.destroy()
        self.result = True
    def _no(self):
        self.destroy()
        self.result = False
    def __bool__(self):
        return bool(self.result)

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Schema Validator : SQL vs PostgreSQL')
        self.geometry('1100x700')
        self.resizable(False, False)
        try_load_theme(self)
        # Main background
        self.configure(bg='#383838', highlightbackground='#383838', highlightthickness=0, bd=0)
        self.bg_frame = tk.Frame(self, bg='#383838', highlightbackground='#383838', highlightthickness=0, bd=0)
        self.bg_frame.pack(fill='both', expand=True)
        # Header with title, subtitle, and About/folder icons
        header_frame = tk.Frame(self.bg_frame, bg='#383838')
        header_frame.pack(side='top', fill='x', pady=(0, 8))
        # Left-aligned header and subtitle
        header_inner = tk.Frame(header_frame, bg='#383838')
        header_inner.pack(side='left', fill='y', padx=(32, 0))
        self.header = tk.Label(header_inner, text='Database Schema Validator', font=(None, 22, 'bold'), fg='#ebe9e8', bg='#383838', bd=0, anchor='w', justify='left')
        self.header.pack(anchor='w', pady=(18, 0))
        self.subtitle = tk.Label(header_inner, text='SQL Server vs PostgreSQL', font=(None, 11, 'italic'), fg='#bbb', bg='#383838', anchor='w', justify='left')
        self.subtitle.pack(anchor='w', pady=(1, 8))
        # Right-side icon buttons
        icon_btn_frame = tk.Frame(header_frame, bg='#383838')
        icon_btn_frame.pack(side='right', padx=10, pady=8)
        folder_icon = load_icon('folder.png', size=(24, 24))
        # Folder icon + text button with rounded white border
        if folder_icon:
            self.open_folder_btn = tk.Button(
                icon_btn_frame, image=folder_icon, text=' Open Reports', compound='left', bg='#383838', fg='#bbb',
                font=(None, 10, 'bold'), bd=0, highlightthickness=2, highlightbackground='#fff', activebackground='#383838',
                command=self.open_report_folder, relief='flat')
            self.open_folder_btn.image = folder_icon
        else:
            self.open_folder_btn = tk.Button(
                icon_btn_frame, text='Open Reports', bg='#383838', fg='#bbb', font=(None, 10, 'bold'), bd=0,
                highlightthickness=2, highlightbackground='#fff', activebackground='#383838', command=self.open_report_folder, relief='flat')
        self.open_folder_btn.pack(side='right', padx=(0, 8))
        # About buttons with rounded white border
        about_icon = load_icon('about.png', size=(28, 28))
        paper_icon = load_icon('paper.png', size=(24, 24))
        self.about_btn = tk.Button(
            icon_btn_frame, image=about_icon, bg='#383838', bd=0, highlightthickness=2, highlightbackground='#fff',
            activebackground='#383838', command=self.show_about_window, relief='flat')
        self.about_btn.image = about_icon
        if paper_icon:
            self.about_paper_btn = tk.Button(
                icon_btn_frame, image=paper_icon, bg='#383838', bd=0, highlightthickness=2, highlightbackground='#fff',
                activebackground='#383838', command=self.show_about_window, relief='flat')
            self.about_paper_btn.image = paper_icon
        else:
            self.about_paper_btn = tk.Button(
                icon_btn_frame, text='About', bg='#383838', fg='#bbb', font=(None, 10, 'bold'), bd=0,
                highlightthickness=2, highlightbackground='#fff', activebackground='#383838', command=self.show_about_window, relief='flat')
        self.about_btn.pack(side='right', padx=(0, 8))
        self.about_paper_btn.pack(side='right', padx=(0, 8))
        # Main content area
        self.content = tk.Frame(self.bg_frame, bg='#383838', highlightbackground='#383838', highlightthickness=0, bd=0)
        self.content.pack(fill='both', expand=True)
        # Left: Config UI with shadow
        self.config_shadow = tk.Frame(self.content, bg='#222', highlightthickness=0, bd=0)
        self.config_shadow.pack(side='left', fill='both', expand=True, padx=(32, 12), pady=22)
        self.config_frame = tk.Frame(self.config_shadow, bg='#2e2d2d', highlightbackground='#4a3d0e', highlightthickness=2, bd=0)
        self.config_frame.pack(fill='both', expand=True, padx=4, pady=4)
        self.config_ui = ConfigUI(self.config_frame)
        self.config_ui.pack(fill='both', expand=True)
        # Right: Validation UI (top) and ResultSet UI (bottom) with shadow
        self.right_shadow = tk.Frame(self.content, bg='#222', highlightthickness=0, bd=0)
        self.right_shadow.pack(side='right', fill='both', expand=True, padx=(12, 32), pady=22)
        self.right_frame = tk.Frame(self.right_shadow, bg='#2e2d2d', highlightbackground='#4a3d0e', highlightthickness=2, bd=0)
        self.right_frame.pack(fill='both', expand=True, padx=4, pady=4)
        self.validation_ui = ValidationUI(self.right_frame)
        self.validation_ui.pack(side='top', fill='x', padx=10, pady=(10, 5))
        # ResultSet UI placeholder (bottom) with Recent Reports (no label/text)
        self.resultset_frame = tk.Frame(self.right_frame, bg='#383838', highlightbackground='#4a3d0e', highlightthickness=1, height=220, bd=0)
        self.resultset_frame.pack(side='bottom', fill='both', expand=True, padx=10, pady=(5, 10))
        # Recent Reports in ResultSet UI
        reports_header = tk.Frame(self.resultset_frame, bg='#383838')
        reports_header.pack(fill='x', padx=10, pady=(8, 2))
        self.reports_label = tk.Label(reports_header, text='Recent Reports', font=(None, 11, 'bold'), fg='#bbb', bg='#383838', anchor='w')
        self.reports_label.pack(side='left', anchor='w')
        refresh_icon = load_icon('refresh.png', size=(18, 18))
        if refresh_icon:
            self.refresh_btn = tk.Button(reports_header, image=refresh_icon, command=self.refresh_reports, bg='#383838', bd=0, highlightthickness=0, activebackground='#2e2d2d', width=22, height=22)
            self.refresh_btn.image = refresh_icon
        else:
            self.refresh_btn = tk.Button(reports_header, text='⟳', command=self.refresh_reports, font=(None, 11, 'bold'), fg='#bbb', bg='#383838', bd=0, highlightthickness=0, activebackground='#2e2d2d', width=2, height=1)
        self.refresh_btn.pack(side='right', anchor='e', padx=(0, 2))
        self.reports_frame = tk.Frame(self.resultset_frame, bg='#383838')
        self.reports_frame.pack(fill='x', padx=10)
        self.refresh_reports()

    def show_about_window(self):
        win = tk.Toplevel(self)
        win.title('About Database Schema Validator')
        win.configure(bg='#2d2d2d')
        win.geometry('480x200')
        win.resizable(False, False)
        tk.Label(win, text='About Database Schema Validator', font=(None, 12, 'bold'), fg='#fff', bg='#2d2d2d').pack(pady=(24, 10))
        about_text = (
            'This application allows you to validate and compare database schemas between SQL Server and PostgreSQL. '
            'It provides a modern, user-friendly interface to configure connections, run schema validations, and generate detailed Excel reports. '
            'Recent reports are easily accessible for review and management.'
        )
        tk.Label(win, text=about_text, font=(None, 11), fg='#bbb', bg='#2d2d2d', wraplength=420, justify='left').pack(padx=28, pady=(0, 18))

    def refresh_reports(self):
        # Clear previous
        for widget in self.reports_frame.winfo_children():
            widget.destroy()
        files = find_latest_reports()[:8]
        icons = {
            'view': load_icon('eye.png', size=(18, 18)),
            'delete': load_icon('delete.png', size=(18, 18))
        }
        for f in files:
            row = tk.Frame(self.reports_frame, bg='#383838')
            row.pack(fill='x', pady=2)
            fname = os.path.basename(f)
            label = tk.Label(row, text=fname, font=(None, 9), bg='#383838', fg='#fff', anchor='w')
            label.pack(side='left', fill='x', expand=True, padx=2)
            view_btn = tk.Button(row, text='View', command=lambda f=f: os.startfile(f), fg='#fff', bg='#383838',
                                 activebackground='#2e2d2d', highlightbackground='#4a3d0e', highlightcolor='#4a3d0e', highlightthickness=1, relief='flat', font=(None, 9))
            if icons.get('view'):
                view_btn.config(image=icons['view'], compound='left', padx=4)
            view_btn.pack(side='right', padx=2)
            del_btn = tk.Button(row, text='Delete', command=lambda f=f: self.delete_report_file(f), fg='#fff', bg='#383838',
                                activebackground='#2e2d2d', highlightbackground='#4a3d0e', highlightcolor='#4a3d0e', highlightthickness=1, relief='flat', font=(None, 9))
            if icons.get('delete'):
                del_btn.config(image=icons['delete'], compound='left', padx=4)
            del_btn.pack(side='right', padx=2)

    def delete_report_file(self, file):
        if CustomMessage(self, f'Delete report file?\n{os.path.basename(file)}', 'Delete', 'warning', ask_yes_no=True):
            try:
                os.remove(file)
                self.refresh_reports()
            except Exception as e:
                CustomMessage(self, f'Could not delete file:\n{e}', 'Error', 'error')

    def show_page(self, name):
        pass  # No-op, as there are no pages now

    def open_report_folder(self):
        import subprocess
        folder = os.path.join(os.path.dirname(__file__), 'SchemaValidationReports')
        if not os.path.exists(folder):
            os.makedirs(folder)
        subprocess.Popen(f'explorer "{folder}"')

if __name__ == '__main__':
    app = MainApp()
    app.mainloop()
