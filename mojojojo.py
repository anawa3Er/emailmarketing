import re
import os
import json
import csv
import smtplib
import threading
import time
import gc
from datetime import datetime
from collections import defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
from tkinter import *
from threading import Lock, Thread
from ttkbootstrap import Style
import pandas as pd
from tkinter.colorchooser import askcolor
from tkinter import font as tkfont

class QuantumEmailSuite:
    def __init__(self, root):
        self.root = root
        self.style = Style(theme='cyborg')
        self.root.title("Quantum Email Suite Pro")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)  # Minimum window size
        
        # Configure grid for main window
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Header
        self.header = ttk.Label(
            root, 
            text="QUANTUM EMAIL SUITE PRO", 
            font=('Arial', 16, 'bold'),
            foreground='#00ff99',
            background='#1a1a1a')
        self.header.grid(row=0, column=0, sticky='ew', pady=(0, 5))
        
        # Border
        self.style.configure('Border.TFrame', background='#00ff99')
        border = ttk.Frame(root, height=2, style='Border.TFrame')
        border.grid(row=1, column=0, sticky='ew', pady=(0, 5))
        
        # Create main notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=2, column=0, sticky='nsew', padx=10, pady=(0, 5))
        
        # Initialize modules
        self.email_cleaner = EmailCleanerModule(self)
        self.email_editor = EmailEditorModule(self)
        self.email_sender = EmailSenderModule(self)
        
        # Add modules as tabs
        self.notebook.add(self.email_cleaner.frame, text="Email Cleaner")
        self.notebook.add(self.email_editor.frame, text="Email Editor")
        self.notebook.add(self.email_sender.frame, text="Email Sender")
        
        # Status bar
        self.status_var = StringVar()
        self.status_bar = ttk.Label(
            root, 
            textvariable=self.status_var, 
            relief='sunken',
            background='#1a1a1a',
            foreground='#00ff99',
            font=('Helvetica', 9))
        self.status_bar.grid(row=3, column=0, sticky='ew')

    def update_status(self, message):
        self.status_var.set(message)
        self.root.update()

    def on_closing(self):
        """Handle application closing"""
        if hasattr(self, 'email_sender'):
            self.email_sender.save_config()
        self.root.destroy()


class EmailCleanerModule:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent.notebook)
        
        # Configuration
        self.MAX_DISPLAY_ITEMS = 1000
        self.MAX_EMAILS_IN_MEMORY = 1000000
        self.FILE_CHUNK_SIZE = 50000
        
        # Data storage
        self.loaded_files = []
        self.email_db = defaultdict(list)
        self.clean_emails = set()
        self.lock = Lock()
        self.processing = False
        
        # Create UI
        self.create_ui()
    
    def create_ui(self):
        # Configure grid for resizing
        self.frame.grid_rowconfigure(3, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        
        # File Loading Section
        load_frame = ttk.LabelFrame(self.frame, text="File Loading", padding=10)
        load_frame.grid(row=0, column=0, sticky='ew', pady=5)
        load_frame.grid_columnconfigure(0, weight=1)
        
        # File list with scrollbar
        file_list_container = ttk.Frame(load_frame)
        file_list_container.grid(row=0, column=0, sticky='nsew', pady=5)
        file_list_container.grid_columnconfigure(0, weight=1)
        
        self.file_listbox = Listbox(file_list_container, height=8, selectmode=EXTENDED)
        self.file_listbox.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(file_list_container, orient='vertical', command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        # Load buttons
        btn_frame = ttk.Frame(load_frame)
        btn_frame.grid(row=1, column=0, sticky='ew')
        
        self.add_btn = ttk.Button(btn_frame, text="Add Files", command=self.add_files)
        self.add_btn.pack(side='left', padx=5)
        
        self.remove_btn = ttk.Button(btn_frame, text="Remove Selected", command=self.remove_selected_files)
        self.remove_btn.pack(side='left', padx=5)
        
        self.clear_btn = ttk.Button(btn_frame, text="Clear All", command=self.clear_files)
        self.clear_btn.pack(side='left', padx=5)
        
        self.load_btn = ttk.Button(btn_frame, text="Load Emails", command=self.load_emails_thread)
        self.load_btn.pack(side='right', padx=5)

        # Processing Section
        process_frame = ttk.LabelFrame(self.frame, text="Email Processing", padding=10)
        process_frame.grid(row=1, column=0, sticky='ew', pady=5)
        process_frame.grid_columnconfigure(0, weight=1)

        # Stats display
        stats_frame = ttk.Frame(process_frame)
        stats_frame.grid(row=0, column=0, sticky='ew', pady=5)
        
        # Create stats labels
        self.total_files_var = StringVar(value="0")
        self.total_emails_var = StringVar(value="0")
        self.unique_emails_var = StringVar(value="0")
        self.valid_emails_var = StringVar(value="0")
        
        stats_labels = [
            ("Total Files:", self.total_files_var),
            ("Total Emails:", self.total_emails_var),
            ("Unique Emails:", self.unique_emails_var),
            ("Valid Emails:", self.valid_emails_var)
        ]
        
        for i, (text, var) in enumerate(stats_labels):
            ttk.Label(stats_frame, text=text).grid(row=0, column=i*2, sticky='e', padx=5)
            ttk.Label(stats_frame, textvariable=var).grid(row=0, column=i*2+1, sticky='w', padx=5)

        # Action buttons
        action_frame = ttk.Frame(process_frame)
        action_frame.grid(row=1, column=0, sticky='ew', pady=5)
        
        buttons = [
            ("Remove Duplicates", self.remove_duplicates),
            ("Remove Invalid", self.remove_invalid),
            ("Clean All", self.clean_all),
            ("Export Clean List", self.export_clean_list)
        ]
        
        for i, (text, command) in enumerate(buttons):
            if i < len(buttons) - 1:
                ttk.Button(action_frame, text=text, command=command).pack(side='left', padx=5)
            else:
                ttk.Button(action_frame, text=text, command=command).pack(side='right', padx=5)

        # Results Display
        results_frame = ttk.LabelFrame(self.frame, text="Email Results", padding=10)
        results_frame.grid(row=2, column=0, sticky='nsew', pady=5)
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        # Treeview with scrollbars
        self.tree = ttk.Treeview(results_frame, columns=("email", "status", "sources"), show="headings")
        self.tree.heading("email", text="Email")
        self.tree.heading("status", text="Status")
        self.tree.heading("sources", text="Source Files")
        self.tree.column("email", width=300, stretch=True)
        self.tree.column("status", width=100, stretch=False)
        self.tree.column("sources", width=400, stretch=True)
        
        yscroll = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        # Progress bar
        self.progress = ttk.Progressbar(self.frame, mode="determinate")
        self.progress.grid(row=3, column=0, sticky='ew', padx=10, pady=5)

    def add_files(self):
        filetypes = [
            ("Text files", "*.txt"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xls *.xlsx *.ods"),
            ("All files", "*.*")
        ]
        files = filedialog.askopenfilenames(filetypes=filetypes)
        if files:
            for file in files:
                if file not in self.loaded_files:
                    self.loaded_files.append(file)
                    self.file_listbox.insert(END, file)
            self.update_stats()

    def remove_selected_files(self):
        selected = self.file_listbox.curselection()
        for i in reversed(selected):
            self.loaded_files.pop(i)
            self.file_listbox.delete(i)
        self.update_stats()

    def clear_files(self):
        self.loaded_files = []
        self.file_listbox.delete(0, END)
        self.email_db.clear()
        self.clean_emails.clear()
        self.update_stats()
        self.update_display()
        self.parent.update_status("Cleared all files and data")

    def update_stats(self):
        self.total_files_var.set(str(len(self.loaded_files)))
        self.total_emails_var.set(str(sum(len(sources) for sources in self.email_db.values())))
        self.unique_emails_var.set(str(len(self.email_db)))
        valid_count = sum(1 for email in self.email_db if self.is_valid_email(email))
        self.valid_emails_var.set(str(valid_count))

    def update_display(self):
        self.tree.delete(*self.tree.get_children())
        count = 0
        for email, sources in list(self.email_db.items())[:self.MAX_DISPLAY_ITEMS]:
            status = "Valid" if self.is_valid_email(email) else "Invalid"
            self.tree.insert("", "end", values=(email, status, ", ".join(sources)))
            count += 1
        if len(self.email_db) > self.MAX_DISPLAY_ITEMS:
            self.tree.insert("", "end", values=(f"... showing {self.MAX_DISPLAY_ITEMS} of {len(self.email_db)} emails", "", ""))

    def is_valid_email(self, email):
        """Basic email validation"""
        regex = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
        return re.fullmatch(regex, email) is not None

    def load_emails_thread(self):
        if not self.loaded_files:
            messagebox.showwarning("No Files", "Please add files first")
            return
        
        if self.processing:
            messagebox.showwarning("Processing", "Already processing files")
            return
            
        self.processing = True
        self.parent.update_status("Loading emails...")
        self.progress["value"] = 0
        self.load_btn.config(state=DISABLED)
        
        # Start processing in a separate thread
        Thread(target=self.load_emails, daemon=True).start()

    def load_emails(self):
        try:
            total_files = len(self.loaded_files)
            email_count = 0
            
            for i, filepath in enumerate(self.loaded_files):
                if not self.processing:  # Check if stopped
                    break
                    
                self.parent.update_status(f"Processing file {i+1}/{total_files}: {os.path.basename(filepath)}")
                self.progress["value"] = (i / total_files) * 100
                self.parent.root.update()
                
                try:
                    if filepath.endswith('.csv'):
                        # Process CSV in chunks
                        for chunk in pd.read_csv(filepath, chunksize=self.FILE_CHUNK_SIZE, dtype=str, engine='c'):
                            for col in chunk.columns:
                                for item in chunk[col].dropna():
                                    if isinstance(item, str):
                                        for email in self.extract_emails(item):
                                            with self.lock:
                                                if email not in self.email_db:
                                                    self.email_db[email] = []
                                                if filepath not in self.email_db[email]:
                                                    self.email_db[email].append(filepath)
                                            email_count += 1
                            del chunk
                            gc.collect()
                            
                    elif filepath.endswith(('.xls', '.xlsx', '.ods')):
                        # Process Excel file
                        xl = pd.ExcelFile(filepath)
                        for sheet in xl.sheet_names:
                            df = xl.parse(sheet, dtype=str)
                            for col in df.columns:
                                for item in df[col].dropna():
                                    if isinstance(item, str):
                                        for email in self.extract_emails(item):
                                            with self.lock:
                                                if email not in self.email_db:
                                                    self.email_db[email] = []
                                                if filepath not in self.email_db[email]:
                                                    self.email_db[email].append(filepath)
                                            email_count += 1
                            del df
                            gc.collect()
                        xl.close()
                        
                    else:  # Text file
                        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                            buffer = ""
                            while True:
                                chunk = f.read(self.FILE_CHUNK_SIZE)
                                if not chunk:
                                    break
                                buffer += chunk
                                lines = buffer.split('\n')
                                buffer = lines.pop()  # Save incomplete line
                                
                                for line in lines:
                                    for email in self.extract_emails(line):
                                        with self.lock:
                                            if email not in self.email_db:
                                                self.email_db[email] = []
                                            if filepath not in self.email_db[email]:
                                                self.email_db[email].append(filepath)
                                        email_count += 1
                                
                                if email_count >= self.MAX_EMAILS_IN_MEMORY:
                                    self.parent.update_status(f"Memory limit reached ({self.MAX_EMAILS_IN_MEMORY} emails)")
                                    break
                                
                            if buffer:  # Process remaining content
                                for email in self.extract_emails(buffer):
                                    with self.lock:
                                        if email not in self.email_db:
                                            self.email_db[email] = []
                                        if filepath not in self.email_db[email]:
                                            self.email_db[email].append(filepath)
                                    email_count += 1
                                    
                except Exception as e:
                    self.parent.update_status(f"Error processing {filepath}: {str(e)}")
                    continue
                    
                if email_count >= self.MAX_EMAILS_IN_MEMORY:
                    break
                    
            self.parent.update_status(f"Loaded {email_count} emails from {min(i+1, total_files)} files")
            
        except Exception as e:
            self.parent.update_status(f"Error: {str(e)}")
        finally:
            self.processing = False
            self.load_btn.config(state=NORMAL)
            self.progress["value"] = 100
            self.update_stats()
            self.update_display()
            gc.collect()

    def extract_emails(self, text):
        """Extract unique emails from text"""
        emails = set()
        regex = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
        for match in re.finditer(regex, text):
            email = match.group(0).lower().strip()
            if email:
                emails.add(email)
        return emails

    def remove_duplicates(self):
        """Keep only one copy of each email"""
        if not self.email_db:
            messagebox.showwarning("No Data", "No emails to process")
            return
            
        self.parent.update_status("Removing duplicates...")
        self.parent.root.update()
        
        # Already unique in our database structure
        self.parent.update_status(f"Found {len(self.email_db)} unique emails")
        self.update_stats()
        self.update_display()

    def remove_invalid(self):
        """Remove emails with invalid syntax"""
        if not self.email_db:
            messagebox.showwarning("No Data", "No emails to process")
            return
            
        self.parent.update_status("Removing invalid emails...")
        self.parent.root.update()
        
        initial_count = len(self.email_db)
        invalid_emails = [email for email in self.email_db if not self.is_valid_email(email)]
        
        for email in invalid_emails:
            del self.email_db[email]
            
        removed = initial_count - len(self.email_db)
        self.parent.update_status(f"Removed {removed} invalid emails, kept {len(self.email_db)} valid ones")
        self.update_stats()
        self.update_display()

    def clean_all(self):
        """Remove duplicates and invalid emails in one step"""
        self.remove_duplicates()
        self.remove_invalid()

    def export_clean_list(self):
        if not self.email_db:
            messagebox.showwarning("No Data", "No emails to export")
            return
            
        filetypes = [
            ("CSV files", "*.csv"),
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ]
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=filetypes,
            title="Save Clean Email List"
        )
        
        if not filename:
            return
            
        try:
            self.parent.update_status("Exporting clean email list...")
            self.progress["value"] = 0
            self.parent.root.update()
            
            valid_emails = [email for email in self.email_db if self.is_valid_email(email)]
            total = len(valid_emails)
            
            if filename.endswith('.csv'):
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(['email'])  # header
                    
                    # Write in chunks to handle large files
                    chunk_size = 10000
                    for i in range(0, total, chunk_size):
                        chunk = valid_emails[i:i+chunk_size]
                        writer.writerows([[email] for email in chunk])
                        self.progress["value"] = (i / total) * 100
                        self.parent.root.update()
                        
            else:  # Text file
                with open(filename, 'w', encoding='utf-8') as f:
                    # Write in chunks
                    chunk_size = 10000
                    for i in range(0, total, chunk_size):
                        chunk = valid_emails[i:i+chunk_size]
                        f.write('\n'.join(chunk) + '\n')
                        self.progress["value"] = (i / total) * 100
                        self.parent.root.update()
                        
            self.parent.update_status(f"Exported {total} clean emails to {filename}")
            self.progress["value"] = 100
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {str(e)}")
            self.parent.update_status("Export failed")
        finally:
            gc.collect()

class EmailEditorModule:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent.notebook)
        
        # Configure grid for resizing
        self.frame.grid_rowconfigure(1, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        
        # Current font settings
        self.current_font_family = 'Arial'
        self.current_font_size = 12
        self.current_font_weight = 'normal'
        self.current_font_slant = 'roman'
        self.current_font_underline = False
        self.current_font_color = '#000000'
        self.current_bg_color = '#ffffff'
        self.current_align = 'left'
        
        # Create advanced editor UI
        self.create_advanced_editor_ui()
    
    def create_advanced_editor_ui(self):
        # Main container with grid configuration
        main_container = ttk.Frame(self.frame)
        main_container.pack(fill='both', expand=True, padx=5, pady=5)
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # Toolbar Frame - Ribbon Style
        toolbar_frame = ttk.Frame(main_container)  # Correct variable name
        toolbar_frame.grid(row=0, column=0, sticky='ew', pady=(0, 5))

        # Create ribbon tabs - FIXED THE TYPO HERE (tooler_frame -> toolbar_frame)
        ribbon_notebook = ttk.Notebook(toolbar_frame)  # This is the corrected line
        ribbon_notebook.pack(fill='x', expand=True)

        # Home Tab
        home_tab = ttk.Frame(ribbon_notebook)
        ribbon_notebook.add(home_tab, text="Home")
        
        # Clipboard group
        clipboard_group = ttk.LabelFrame(home_tab, text="Clipboard", padding=5)
        clipboard_group.pack(side='left', fill='y', padx=5)
        
        ttk.Button(clipboard_group, text="Paste", command=lambda: self.editor.event_generate('<<Paste>>')).pack(side='left', padx=2)
        ttk.Button(clipboard_group, text="Copy", command=lambda: self.editor.event_generate('<<Copy>>')).pack(side='left', padx=2)
        ttk.Button(clipboard_group, text="Cut", command=lambda: self.editor.event_generate('<<Cut>>')).pack(side='left', padx=2)
        
        # Font group
        font_group = ttk.LabelFrame(home_tab, text="Font", padding=5)
        font_group.pack(side='left', fill='y', padx=5)
        
        # Font family
        font_families = sorted(tkfont.families())
        self.font_family = ttk.Combobox(font_group, values=font_families, width=15)
        self.font_family.set('Arial')
        self.font_family.bind('<<ComboboxSelected>>', self.change_font_family)
        self.font_family.pack(side='left', padx=2)
        
        # Font size
        self.font_size = ttk.Combobox(font_group, values=[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72], width=3)
        self.font_size.set('12')
        self.font_size.bind('<<ComboboxSelected>>', self.change_font_size)
        self.font_size.pack(side='left', padx=2)
        
        # Font style buttons
        style_btn_frame = ttk.Frame(font_group)
        style_btn_frame.pack(side='left', padx=5)
        
        self.bold_btn = ttk.Button(style_btn_frame, text="B", width=2, command=self.toggle_bold)
        self.bold_btn.pack(side='left', padx=1)
        
        self.italic_btn = ttk.Button(style_btn_frame, text="I", width=2, command=self.toggle_italic)
        self.italic_btn.pack(side='left', padx=1)
        
        self.underline_btn = ttk.Button(style_btn_frame, text="U", width=2, command=self.toggle_underline)
        self.underline_btn.pack(side='left', padx=1)
        
        # Font color
        self.color_btn = ttk.Button(font_group, text="Color", command=self.choose_font_color)
        self.color_btn.pack(side='left', padx=2)
        
        # Background color
        self.bg_color_btn = ttk.Button(font_group, text="BG Color", command=self.choose_bg_color)
        self.bg_color_btn.pack(side='left', padx=2)
        
        # Paragraph group
        paragraph_group = ttk.LabelFrame(home_tab, text="Paragraph", padding=5)
        paragraph_group.pack(side='left', fill='y', padx=5)
        
        self.align_left_btn = ttk.Button(paragraph_group, text="Left", command=lambda: self.set_alignment('left'))
        self.align_left_btn.pack(side='left', padx=1)
        
        self.align_center_btn = ttk.Button(paragraph_group, text="Center", command=lambda: self.set_alignment('center'))
        self.align_center_btn.pack(side='left', padx=1)
        
        self.align_right_btn = ttk.Button(paragraph_group, text="Right", command=lambda: self.set_alignment('right'))
        self.align_right_btn.pack(side='left', padx=1)
        
        # Lists
        self.bullet_list_btn = ttk.Button(paragraph_group, text="• List", command=self.insert_bullet)
        self.bullet_list_btn.pack(side='left', padx=1)
        
        # Insert Tab
        insert_tab = ttk.Frame(ribbon_notebook)
        ribbon_notebook.add(insert_tab, text="Insert")
        
        # Insert group
        insert_group = ttk.LabelFrame(insert_tab, text="Insert", padding=5)
        insert_group.pack(side='left', fill='y', padx=5)
        
        ttk.Button(insert_group, text="Image", command=self.insert_image).pack(side='left', padx=2)
        ttk.Button(insert_group, text="Hyperlink", command=self.insert_hyperlink).pack(side='left', padx=2)
        ttk.Button(insert_group, text="Table", command=self.insert_table).pack(side='left', padx=2)
        
        # Text editor with scrollbars
        editor_frame = ttk.Frame(main_container)
        editor_frame.grid(row=1, column=0, sticky='nsew')
        editor_frame.grid_rowconfigure(0, weight=1)
        editor_frame.grid_columnconfigure(0, weight=1)

        self.editor = scrolledtext.ScrolledText(
            editor_frame,
            wrap='word',
            font=('Arial', 12),
            undo=True,
            maxundo=-1,
            autoseparators=True,
            padx=10,
            pady=10
        )
        self.editor.grid(row=0, column=0, sticky='nsew')
        
        # Configure tags for formatting
        self.editor.tag_configure('bold', font=('Arial', 12, 'bold'))
        self.editor.tag_configure('italic', font=('Arial', 12, 'italic'))
        self.editor.tag_configure('underline', underline=True)
        self.editor.tag_configure('center', justify='center')
        self.editor.tag_configure('right', justify='right')
        self.editor.tag_configure('left', justify='left')
        
        # Status bar
        status_frame = ttk.Frame(main_container)
        status_frame.grid(row=2, column=0, sticky='ew', pady=(5, 0))
        
        self.word_count_label = ttk.Label(status_frame, text="Words: 0")
        self.word_count_label.pack(side='left', padx=5)
        
        self.char_count_label = ttk.Label(status_frame, text="Chars: 0")
        self.char_count_label.pack(side='left', padx=5)
        
        # Export buttons
        export_frame = ttk.Frame(main_container)
        export_frame.grid(row=3, column=0, sticky='e', pady=5)
        
        ttk.Button(export_frame, text="Export as HTML", command=self.export_html).pack(side='left', padx=5)
        ttk.Button(export_frame, text="Export as TXT", command=self.export_txt).pack(side='left', padx=5)
    
    def update_counts(self, event=None):
        """Update word and character counts"""
        content = self.editor.get('1.0', 'end-1c')
        words = len(content.split())
        chars = len(content)
        self.word_count_label.config(text=f"Words: {words}")
        self.char_count_label.config(text=f"Chars: {chars}")
    
    def update_font_display(self, event=None):
        """Update font display based on current selection"""
        try:
            # Get current tags at the insertion cursor
            tags = self.editor.tag_names("insert")
            
            # Update button states based on tags
            self.bold_btn.state(['!pressed'] if 'bold' in tags else ['pressed'])
            self.italic_btn.state(['!pressed'] if 'italic' in tags else ['pressed'])
            self.underline_btn.state(['!pressed'] if 'underline' in tags else ['pressed'])
            
            # Update alignment buttons
            if 'center' in tags:
                self.current_align = 'center'
                self.align_center_btn.state(['pressed'])
                self.align_left_btn.state(['!pressed'])
                self.align_right_btn.state(['!pressed'])
            elif 'right' in tags:
                self.current_align = 'right'
                self.align_right_btn.state(['pressed'])
                self.align_left_btn.state(['!pressed'])
                self.align_center_btn.state(['!pressed'])
            else:
                self.current_align = 'left'
                self.align_left_btn.state(['pressed'])
                self.align_center_btn.state(['!pressed'])
                self.align_right_btn.state(['!pressed'])
                
        except Exception as e:
            print(f"Error updating font display: {e}")
    
    def change_font_family(self, event=None):
        """Change font family for selection or new text"""
        self.current_font_family = self.font_family.get()
        self.apply_font()
    
    def change_font_size(self, event=None):
        """Change font size for selection or new text"""
        try:
            self.current_font_size = int(self.font_size.get())
            self.apply_font()
        except ValueError:
            pass
    
    def toggle_bold(self):
        """Toggle bold style"""
        self.current_font_weight = 'bold' if self.current_font_weight == 'normal' else 'normal'
        self.apply_font()
    
    def toggle_italic(self):
        """Toggle italic style"""
        self.current_font_slant = 'italic' if self.current_font_slant == 'roman' else 'roman'
        self.apply_font()
    
    def toggle_underline(self):
        """Toggle underline style"""
        self.current_font_underline = not self.current_font_underline
        self.apply_font()
    
    def choose_font_color(self):
        """Choose font color"""
        color = askcolor(title="Select Font Color")
        if color[1]:  # color[1] is the hex code
            self.current_font_color = color[1]
            self.apply_font()
    
    def choose_bg_color(self):
        """Choose background color"""
        color = askcolor(title="Select Background Color")
        if color[1]:
            self.current_bg_color = color[1]
            self.apply_font()
    
    def set_alignment(self, align):
        """Set text alignment"""
        self.current_align = align
        try:
            # Remove all alignment tags from selection
            for tag in ['left', 'center', 'right']:
                self.editor.tag_remove(tag, "sel.first", "sel.last")
            
            # Apply new alignment tag
            if self.current_align != 'left':  # left is default
                self.editor.tag_add(self.current_align, "sel.first", "sel.last")
        except:
            pass
    
    def apply_font(self):
        """Apply current font settings to selection"""
        try:
            # Create or update font configuration
            font_config = (self.current_font_family, self.current_font_size, 
                          self.current_font_weight, self.current_font_slant)
            
            # Create a unique tag name based on font settings
            tag_name = f"font_{self.current_font_family}_{self.current_font_size}_{self.current_font_weight}_{self.current_font_slant}"
            
            # Configure the tag if it doesn't exist
            if tag_name not in self.editor.tag_names():
                self.editor.tag_configure(tag_name, 
                                        font=font_config,
                                        underline=self.current_font_underline,
                                        foreground=self.current_font_color,
                                        background=self.current_bg_color)
            
            # Apply the tag to the current selection
            self.editor.tag_add(tag_name, "sel.first", "sel.last")
            
            # Update button states
            self.update_font_display()
            
        except Exception as e:
            print(f"Error applying font: {e}")
    
    def insert_bullet(self):
        """Insert bullet point at current position"""
        self.editor.insert('insert', '• ')
    
    def insert_image(self):
        """Insert image placeholder"""
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg *.gif")])
        if file_path:
            self.editor.insert('insert', f'\n[Image: {os.path.basename(file_path)}]\n')
    
    def insert_hyperlink(self):
        """Insert hyperlink"""
        url = simpledialog.askstring("Insert Hyperlink", "Enter URL:")
        if url:
            text = simpledialog.askstring("Insert Hyperlink", "Enter display text:", initialvalue=url)
            if text:
                self.editor.insert('insert', f'<a href="{url}">{text}</a>')
    
    def insert_table(self):
        """Insert table placeholder"""
        rows = simpledialog.askinteger("Insert Table", "Number of rows:", minvalue=1, maxvalue=20)
        cols = simpledialog.askinteger("Insert Table", "Number of columns:", minvalue=1, maxvalue=10)
        
        if rows and cols:
            table = "\n" + "+-----" * cols + "+\n"
            for _ in range(rows):
                table += "|     " * cols + "|\n"
                table += "+-----" * cols + "+\n"
            self.editor.insert('insert', table)
    
    def export_html(self):
        """Export content as HTML"""
        text = self.editor.get('1.0', 'end-1c')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML Files", "*.html")]
        )
        if file_path:
            try:
                # Enhanced HTML conversion with basic styling
                html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>Exported Email</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
        .bold {{ font-weight: bold; }}
        .italic {{ font-style: italic; }}
        .underline {{ text-decoration: underline; }}
        .center {{ text-align: center; }}
        .right {{ text-align: right; }}
    </style>
</head>
<body>
{text}
</body>
</html>"""
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                messagebox.showinfo("Success", f"Exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")
    
    def export_txt(self):
        """Export content as plain text"""
        text = self.editor.get('1.0', 'end-1c')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(text)
                messagebox.showinfo("Success", f"Exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")


class EmailSenderModule:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent.notebook)
        
        # Configure grid for resizing
        self.frame.grid_rowconfigure(2, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        
        # Initialize variables
        self.email_list = []
        self.smtp_accounts = []
        self.current_account_index = 0
        self.email_counters = {}
        self.sending_active = False
        self.pause_sending = False
        self.last_sent_times = {}
        
        # Load configuration
        self.config_file = "quantum_email_config.json"
        self.load_config()
        
        # Create UI with notebook for tabs
        self.create_ui_with_tabs()
    
    def create_ui_with_tabs(self):
        # Create notebook for tabs
        self.module_notebook = ttk.Notebook(self.frame)
        self.module_notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create tabs
        self.create_compose_tab()
        self.create_smtp_tab()
        self.create_status_tab()
        
        # Add tabs to notebook
        self.module_notebook.add(self.compose_tab, text="Compose")
        self.module_notebook.add(self.smtp_tab, text="SMTP Accounts")
        self.module_notebook.add(self.status_tab, text="Sending Status")
    
    def create_compose_tab(self):
        self.compose_tab = ttk.Frame(self.module_notebook)
        
        # Configure grid
        self.compose_tab.grid_rowconfigure(1, weight=1)
        self.compose_tab.grid_columnconfigure(0, weight=1)
        
        # Email List Section
        email_list_frame = ttk.LabelFrame(self.compose_tab, text="Email List", padding=10)
        email_list_frame.grid(row=0, column=0, sticky='ew', pady=5)
        email_list_frame.grid_columnconfigure(0, weight=1)
        
        # Email List with scrollbar
        list_container = ttk.Frame(email_list_frame)
        list_container.grid(row=0, column=0, sticky='nsew')
        list_container.grid_columnconfigure(0, weight=1)
        
        self.email_listbox = Listbox(list_container, height=8, selectmode=EXTENDED)
        self.email_listbox.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(list_container, orient='vertical', command=self.email_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.email_listbox.config(yscrollcommand=scrollbar.set)
        
        # Load button
        ttk.Button(
            email_list_frame, 
            text="Load Email List (CSV)", 
            command=self.load_email_list,
            cursor="hand2"
        ).grid(row=1, column=0, sticky='w', pady=5)
        
        # Email Content Section
        content_frame = ttk.LabelFrame(self.compose_tab, text="Email Content", padding=10)
        content_frame.grid(row=1, column=0, sticky='nsew', pady=5)
        content_frame.grid_rowconfigure(1, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)
        
        # Subject Field
        ttk.Label(content_frame, text="Subject:").grid(row=0, column=0, sticky='w', pady=2)
        self.subject_entry = ttk.Entry(content_frame)
        self.subject_entry.grid(row=0, column=1, sticky='ew', pady=2)
        content_frame.grid_columnconfigure(1, weight=1)
        
        # Message Body with scrollbars
        editor_frame = ttk.Frame(content_frame)
        editor_frame.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=5)
        editor_frame.grid_rowconfigure(0, weight=1)
        editor_frame.grid_columnconfigure(0, weight=1)
        
        self.message_text = scrolledtext.ScrolledText(
            editor_frame, 
            wrap='word', 
            height=15,
            font=('Arial', 11)
        )
        self.message_text.grid(row=0, column=0, sticky='nsew')
        
        # Control Buttons
        btn_frame = ttk.Frame(self.compose_tab)
        btn_frame.grid(row=2, column=0, sticky='e', pady=5)
        
        self.btn_start = ttk.Button(
            btn_frame, 
            text="Start Sending", 
            command=self.start_sending,
            style='success.TButton',
            cursor="hand2"
        )
        self.btn_start.pack(side='left', padx=5)
        
        self.btn_pause = ttk.Button(
            btn_frame, 
            text="Pause", 
            command=self.toggle_pause,
            style='warning.TButton',
            cursor="hand2",
            state='disabled'
        )
        self.btn_pause.pack(side='left', padx=5)
        
        self.btn_stop = ttk.Button(
            btn_frame, 
            text="Stop", 
            command=self.stop_sending,
            style='danger.TButton',
            cursor="hand2",
            state='disabled'
        )
        self.btn_stop.pack(side='left', padx=5)
    
    def create_smtp_tab(self):
        self.smtp_tab = ttk.Frame(self.module_notebook)
        
        # Configure grid
        self.smtp_tab.grid_rowconfigure(1, weight=1)
        self.smtp_tab.grid_columnconfigure(0, weight=1)
        
        # SMTP Accounts List
        list_frame = ttk.LabelFrame(self.smtp_tab, text="SMTP Accounts", padding=10)
        list_frame.grid(row=0, column=0, sticky='nsew', pady=5)
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Listbox with scrollbar
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky='nsew')
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)
        
        self.smtp_listbox = Listbox(list_container, height=10, selectmode=SINGLE)
        self.smtp_listbox.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(list_container, orient='vertical', command=self.smtp_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.smtp_listbox.config(yscrollcommand=scrollbar.set)
        
        # Account Management Buttons
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=1, column=0, sticky='ew', pady=5)
        
        ttk.Button(
            btn_frame, 
            text="Add Account", 
            command=self.add_smtp_account
        ).pack(side='left', padx=2)
        
        ttk.Button(
            btn_frame, 
            text="Remove Account", 
            command=self.remove_smtp_account
        ).pack(side='left', padx=2)
        
        ttk.Button(
            btn_frame, 
            text="Edit Account", 
            command=self.edit_smtp_account
        ).pack(side='left', padx=2)
        
        ttk.Button(
            btn_frame, 
            text="Test Connection", 
            command=self.test_connection,
            style='info.TButton'
        ).pack(side='right', padx=2)
        
        # SMTP Configuration Fields
        config_frame = ttk.LabelFrame(self.smtp_tab, text="SMTP Configuration", padding=10)
        config_frame.grid(row=1, column=0, sticky='nsew', pady=5)
        config_frame.grid_columnconfigure(1, weight=1)
        
        # Server
        ttk.Label(config_frame, text="SMTP Server:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        self.smtp_server = ttk.Entry(config_frame)
        self.smtp_server.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        
        # Port
        ttk.Label(config_frame, text="Port:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        self.smtp_port = ttk.Entry(config_frame)
        self.smtp_port.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        self.smtp_port.insert(0, "587")  # Default port
        
        # Email
        ttk.Label(config_frame, text="Email:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        self.smtp_email = ttk.Entry(config_frame)
        self.smtp_email.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        
        # Display Name
        ttk.Label(config_frame, text="Display Name:").grid(row=3, column=0, sticky='e', padx=5, pady=2)
        self.smtp_display_name = ttk.Entry(config_frame)
        self.smtp_display_name.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        
        # Password
        ttk.Label(config_frame, text="Password:").grid(row=4, column=0, sticky='e', padx=5, pady=2)
        self.smtp_password = ttk.Entry(config_frame, show="•")
        self.smtp_password.grid(row=4, column=1, sticky='ew', padx=5, pady=2)
        
        # Daily Limit
        ttk.Label(config_frame, text="Daily Limit:").grid(row=5, column=0, sticky='e', padx=5, pady=2)
        self.daily_limit = ttk.Entry(config_frame)
        self.daily_limit.grid(row=5, column=1, sticky='ew', padx=5, pady=2)
        self.daily_limit.insert(0, "100")  # Default limit
        
        # Delay between emails
        ttk.Label(config_frame, text="Delay (seconds):").grid(row=6, column=0, sticky='e', padx=5, pady=2)
        self.email_delay = ttk.Entry(config_frame)
        self.email_delay.grid(row=6, column=1, sticky='ew', padx=5, pady=2)
        self.email_delay.insert(0, "5")  # Default delay
    
    def create_status_tab(self):
        self.status_tab = ttk.Frame(self.module_notebook)
        
        # Configure grid
        self.status_tab.grid_rowconfigure(0, weight=1)
        self.status_tab.grid_columnconfigure(0, weight=1)
        
        # Status Treeview
        columns = ("#", "Email", "Status", "Timestamp", "SMTP Account", "Details")
        self.status_tree = ttk.Treeview(
            self.status_tab, 
            columns=columns, 
            show="headings",
            selectmode='extended',
            height=20
        )
        
        for col in columns:
            self.status_tree.heading(col, text=col)
            self.status_tree.column(col, width=100, anchor='w')
        
        self.status_tree.column("#", width=40)
        self.status_tree.column("Email", width=180)
        self.status_tree.column("Status", width=100)
        self.status_tree.column("Timestamp", width=140)
        self.status_tree.column("SMTP Account", width=150)
        self.status_tree.column("Details", width=300)
        
        # Add scrollbars
        yscroll = ttk.Scrollbar(self.status_tab, orient="vertical", command=self.status_tree.yview)
        xscroll = ttk.Scrollbar(self.status_tab, orient="horizontal", command=self.status_tree.xview)
        self.status_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        
        self.status_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        
        # Action buttons
        btn_frame = ttk.Frame(self.status_tab)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=5)
        
        ttk.Button(
            btn_frame, 
            text="Clear Status", 
            command=self.clear_status
        ).pack(side='left', padx=5)
        
        ttk.Button(
            btn_frame, 
            text="Export to CSV", 
            command=self.export_status
        ).pack(side='left', padx=5)
        
        # Progress bar
        self.progress_var = DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.status_tab, 
            variable=self.progress_var, 
            maximum=100
        )
        self.progress_bar.grid(row=3, column=0, columnspan=2, sticky='ew', padx=5, pady=5)
        
        # Progress label
        self.progress_label = ttk.Label(self.status_tab, text="Ready")
        self.progress_label.grid(row=4, column=0, columnspan=2, sticky='ew', padx=5, pady=2)
    
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.smtp_accounts = config.get('smtp_accounts', [])
                    self.last_sent_times = config.get('last_sent_times', {})
                    
                    # Initialize counters
                    for i, account in enumerate(self.smtp_accounts):
                        self.email_counters[i] = account.get('sent_today', 0)
                        self.smtp_listbox.insert('end', self.format_account_display(account))
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load config: {str(e)}")

    def save_config(self):
        try:
            # Update sent counts in accounts before saving
            for i, account in enumerate(self.smtp_accounts):
                account['sent_today'] = self.email_counters.get(i, 0)
            
            config = {
                'smtp_accounts': self.smtp_accounts,
                'last_sent_times': self.last_sent_times
            }
            
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {str(e)}")

    def format_account_display(self, account):
        return f"{account['email']} (Limit: {account['limit']}/day, Sent: {account.get('sent_today', 0)})"

    def load_email_list(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    reader = csv.reader(f)
                    self.email_list = [row[0].strip() for row in reader if row and row[0].strip()]
                
                # Update listbox
                self.email_listbox.delete(0, 'end')
                for email in self.email_list[:100]:  # Show first 100 as preview
                    self.email_listbox.insert('end', email)
                if len(self.email_list) > 100:
                    self.email_listbox.insert('end', f"... and {len(self.email_list) - 100} more")
                
                self.parent.update_status(f"Loaded {len(self.email_list)} emails from {os.path.basename(file_path)}")
                self.update_progress()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    def add_smtp_account(self):
        # Get all field values
        server = self.smtp_server.get().strip()
        port = self.smtp_port.get().strip()
        email = self.smtp_email.get().strip()
        display_name = self.smtp_display_name.get().strip()
        password = self.smtp_password.get().strip()
        limit = self.daily_limit.get().strip()
        delay = self.email_delay.get().strip()
        
        # Validate fields
        if not all([server, port, email, password, limit, delay]):
            messagebox.showerror("Error", "All fields are required!")
            return
            
        try:
            limit = int(limit)
            if limit <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Daily limit must be a positive number")
            return
            
        try:
            port = int(port)
            if port <= 0 or port > 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Port must be a valid number (1-65535)")
            return
            
        try:
            delay = int(delay)
            if delay < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Delay must be a positive number")
            return
            
        # Add account
        account = {
            'server': server,
            'port': port,
            'email': email,
            'display_name': display_name,
            'password': password,
            'limit': limit,
            'delay': delay,
            'sent_today': 0
        }
        
        self.smtp_accounts.append(account)
        index = len(self.smtp_accounts) - 1
        self.email_counters[index] = 0
        self.smtp_listbox.insert('end', self.format_account_display(account))
        
        # Clear fields
        self.smtp_password.delete(0, 'end')
        
        # Save config
        self.save_config()

    def remove_smtp_account(self):
        if selected := self.smtp_listbox.curselection():
            index = selected[0]
            self.smtp_listbox.delete(index)
            del self.smtp_accounts[index]
            
            # Rebuild counters
            new_counters = {}
            for i, account in enumerate(self.smtp_accounts):
                new_counters[i] = self.email_counters.get(i, 0)
            self.email_counters = new_counters
            
            self.save_config()

    def edit_smtp_account(self):
        if selected := self.smtp_listbox.curselection():
            index = selected[0]
            account = self.smtp_accounts[index]
            
            # Fill fields with selected account
            self.smtp_server.delete(0, 'end')
            self.smtp_server.insert(0, account['server'])
            
            self.smtp_port.delete(0, 'end')
            self.smtp_port.insert(0, str(account['port']))
            
            self.smtp_email.delete(0, 'end')
            self.smtp_email.insert(0, account['email'])
            
            self.smtp_display_name.delete(0, 'end')
            self.smtp_display_name.insert(0, account.get('display_name', ''))
            
            self.smtp_password.delete(0, 'end')
            self.smtp_password.insert(0, account['password'])
            
            self.daily_limit.delete(0, 'end')
            self.daily_limit.insert(0, str(account['limit']))
            
            self.email_delay.delete(0, 'end')
            self.email_delay.insert(0, str(account.get('delay', 5)))
            
            # Remove the account (will be re-added if save is clicked)
            self.smtp_listbox.delete(index)
            del self.smtp_accounts[index]
            del self.email_counters[index]

    def test_connection(self):
        # Get current values
        server = self.smtp_server.get().strip()
        port = self.smtp_port.get().strip()
        email = self.smtp_email.get().strip()
        password = self.smtp_password.get().strip()
        
        if not all([server, port, email, password]):
            messagebox.showerror("Error", "All fields are required to test connection")
            return
            
        try:
            port = int(port)
        except ValueError:
            messagebox.showerror("Error", "Port must be a valid number")
            return
            
        server_obj = None
        try:
            if port == 465:
                server_obj = smtplib.SMTP_SSL(server, port, timeout=10)
            else:
                server_obj = smtplib.SMTP(server, port, timeout=10)
                server_obj.starttls()
                
            server_obj.login(email, password)
            messagebox.showinfo("Success", "Connection successful!")
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror("Error", "Authentication failed. Check your email and password.")
        except Exception as e:
            messagebox.showerror("Error", f"Connection failed: {str(e)}")
        finally:
            if server_obj:
                server_obj.quit()

    def start_sending(self):
        if not self.email_list:
            messagebox.showerror("Error", "Please load email list first!")
            return
        if not self.smtp_accounts:
            messagebox.showerror("Error", "Please add SMTP accounts first!")
            return
        if not self.subject_entry.get().strip():
            messagebox.showerror("Error", "Please enter a subject!")
            return
        if not self.message_text.get("1.0", "end-1c").strip():
            messagebox.showerror("Error", "Please enter a message!")
            return
        
        # Check if any accounts have available quota
        available_accounts = False
        for i, account in enumerate(self.smtp_accounts):
            if self.check_account_limit(i):
                available_accounts = True
                break
        
        if not available_accounts:
            messagebox.showinfo("Limit Reached", 
                "All accounts have reached their daily limits. Please add more accounts or wait 24 hours.")
            return
        
        # Update UI
        self.btn_start['state'] = 'disabled'
        self.btn_pause['state'] = 'normal'
        self.btn_stop['state'] = 'normal'
        self.sending_active = True
        self.pause_sending = False
        
        # Start sending thread
        thread = threading.Thread(target=self.send_emails)
        thread.daemon = True
        thread.start()

    def toggle_pause(self):
        self.pause_sending = not self.pause_sending
        if self.pause_sending:
            self.btn_pause.config(text="Resume", style='success.TButton')
            self.parent.update_status("Sending paused")
        else:
            self.btn_pause.config(text="Pause", style='warning.TButton')
            self.parent.update_status("Resuming sending...")

    def stop_sending(self):
        self.sending_active = False
        self.btn_start['state'] = 'normal'
        self.btn_pause['state'] = 'disabled'
        self.btn_stop['state'] = 'disabled'
        self.parent.update_status("Sending stopped by user")

    def send_emails(self):
        total_sent = 0
        total_emails = len(self.email_list)
        
        try:
            for i, email in enumerate(self.email_list):
                if not self.sending_active:
                    break
                
                while self.pause_sending and self.sending_active:
                    time.sleep(1)
                
                # Find next available account
                account_index = self.get_next_available_account()
                if account_index is None:
                    messagebox.showinfo("Limit Reached", 
                        "All accounts have reached their daily limits. Please add more accounts or wait 24 hours.")
                    break
                
                account = self.smtp_accounts[account_index]
                
                # Create message
                msg = MIMEMultipart('alternative')
                msg['Subject'] = self.subject_entry.get()
                
                # Set From address
                if account.get('display_name'):
                    msg['From'] = f"{account['display_name']} <{account['email']}>"
                else:
                    msg['From'] = account['email']
                
                msg['To'] = email
                
                # Add message body
                text_content = self.message_text.get("1.0", "end-1c").strip()
                msg.attach(MIMEText(text_content, 'plain'))
                
                try:
                    # Connect to server
                    if account['port'] == 465:
                        server = smtplib.SMTP_SSL(account['server'], account['port'])
                    else:
                        server = smtplib.SMTP(account['server'], account['port'])
                        server.starttls()
                    
                    server.login(account['email'], account['password'])
                    server.sendmail(account['email'], [email], msg.as_string())
                    server.quit()
                    
                    # Update counters
                    total_sent += 1
                    self.email_counters[account_index] += 1
                    
                    # Update last sent time
                    self.last_sent_times[account_index] = datetime.now().timestamp()
                    
                    # Update status
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.status_tree.insert('', 'end', values=(
                        i + 1,
                        email,
                        "Sent",
                        timestamp,
                        account['email'],
                        ""
                    ))
                    
                    # Update UI
                    self.parent.update_status(f"Sent: {total_sent}/{total_emails} | Using: {account['email']} ({self.email_counters[account_index]}/{account['limit']})")
                    self.update_progress(i + 1, total_emails)
                    
                    # Update account display
                    self.smtp_listbox.delete(account_index)
                    self.smtp_listbox.insert(account_index, self.format_account_display(account))
                    
                    # Save progress
                    self.save_config()
                    
                    # Apply delay between emails
                    time.sleep(account.get('delay', 5))
                    
                except Exception as e:
                    error_msg = str(e)
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.status_tree.insert('', 'end', values=(
                        i + 1,
                        email,
                        "Failed",
                        timestamp,
                        account['email'],
                        error_msg
                    ))
                    continue
            
            # Sending complete
            self.parent.update_status(f"Finished sending {total_sent} emails")
            self.btn_start['state'] = 'normal'
            self.btn_pause['state'] = 'disabled'
            self.btn_stop['state'] = 'disabled'
            self.sending_active = False
            
            if total_sent == total_emails:
                messagebox.showinfo("Complete", f"Successfully sent all {total_sent} emails!")
            else:
                messagebox.showinfo("Stopped", f"Sent {total_sent} out of {total_emails} emails")
            
        except Exception as e:
            messagebox.showerror("Error", f"Sending failed: {str(e)}")
            self.btn_start['state'] = 'normal'
            self.btn_pause['state'] = 'disabled'
            self.btn_stop['state'] = 'disabled'
            self.sending_active = False

    def get_next_available_account(self):
        """Find the next available account that hasn't reached its daily limit"""
        for i in range(len(self.smtp_accounts)):
            if self.check_account_limit(i):
                return i
        
        # If all accounts are at limit, check if any have passed 24 hours
        for i in range(len(self.smtp_accounts)):
            last_sent = self.last_sent_times.get(i, 0)
            if datetime.now().timestamp() - last_sent > 24 * 60 * 60:
                # Reset counter for this account
                self.email_counters[i] = 0
                account = self.smtp_accounts[i]
                account['sent_today'] = 0
                self.smtp_listbox.delete(i)
                self.smtp_listbox.insert(i, self.format_account_display(account))
                return i
        
        return None

    def check_account_limit(self, account_index):
        """Check if account hasn't reached its daily limit"""
        account = self.smtp_accounts[account_index]
        sent_today = self.email_counters.get(account_index, 0)
        
        # Check if we've passed 24 hours since last send
        last_sent = self.last_sent_times.get(account_index, 0)
        if datetime.now().timestamp() - last_sent > 24 * 60 * 60:
            # Reset counter
            self.email_counters[account_index] = 0
            account['sent_today'] = 0
            self.smtp_listbox.delete(account_index)
            self.smtp_listbox.insert(account_index, self.format_account_display(account))
            return True
        
        return sent_today < account['limit']

    def update_progress(self, current=0, total=1):
        """Update progress bar and label"""
        if total == 0:
            percent = 0
        else:
            percent = (current / total) * 100
        
        self.progress_var.set(percent)
        self.progress_label.config(text=f"{current}/{total} ({percent:.1f}%)")

    def clear_status(self):
        """Clear the status treeview"""
        for item in self.status_tree.get_children():
            self.status_tree.delete(item)
        self.update_progress(0, 1)

    def export_status(self):
        """Export status to CSV file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["#", "Email", "Status", "Timestamp", "SMTP Account", "Details"])
                    
                    for item in self.status_tree.get_children():
                        writer.writerow(self.status_tree.item(item)['values'])
                
                messagebox.showinfo("Success", f"Status exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export status: {str(e)}")

    def on_closing(self):
        """Handle module closing"""
        self.save_config()


if __name__ == "__main__":
    root = Tk()
    app = QuantumEmailSuite(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()