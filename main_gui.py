import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import webbrowser
from tkinter.scrolledtext import ScrolledText
import logging
import traceback
import sys

import modified_excel
import summary_report_tool

# Setup logger
logger = logging.getLogger("AppLogger")
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler("app.log", encoding="utf-8")
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Font definitions
FONT_FAMILY = "Segoe UI Variable, Segoe UI, Calibri, Arial"
HEADER_FONT = ("Segoe UI Variable", 38, "bold")
SUBHEADER_FONT = ("Segoe UI Variable", 20, "bold")
LABEL_FONT = ("Segoe UI Variable", 13, "bold")
ENTRY_FONT = ("Segoe UI Variable", 12)
BUTTON_FONT = ("Segoe UI Variable", 13, "bold")
CONSOLE_FONT = ("Consolas", 12)
FOOTER_FONT = ("Segoe UI Variable", 12, "bold")

class PhenomenalGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HCCB Excel Report & Summary Generator")
        self.geometry("1100x700")
        self.configure(bg="#e9ecef")
        self.resizable(True, True)
        # Set window icon
        try:
            if sys.platform.startswith('win'):
                self.iconbitmap("app_icon.ico")
            else:
                icon = tk.PhotoImage(file="app_icon.png")
                self.iconphoto(True, icon)
        except Exception as e:
            print("Icon not set:", e)

        # Colors and fonts
        HEADER_BG = "#2c3e50"
        HEADER_FG = "#ffffff"
        PRIMARY = "#007acc"
        CARD_BG = "#f7fafd"
        CARD_BORDER = "#d0d7de"
        BUTTON_BG = PRIMARY
        BUTTON_FG = "#ffffff"
        BUTTON_ACTIVE = "#005999"
        CONSOLE_BG = "#23272e"
        CONSOLE_FG = "#00ff99"
        FOOTER_BG = "#2c3e50"
        FOOTER_FG = "#ffffff"

        # Header
        header = tk.Frame(self, bg=HEADER_BG, height=120)
        header.pack(side=tk.TOP, fill=tk.X)
        header.grid_propagate(False)
        tk.Label(header, text="TNBT", bg=HEADER_BG, fg=HEADER_FG, font=HEADER_FONT).pack(pady=(10,0))
        tk.Label(header, text="The Next Big Thing", bg=HEADER_BG, fg=HEADER_FG, font=SUBHEADER_FONT).pack(pady=(0,10))

        # Main split area
        main = tk.Frame(self, bg="#e9ecef")
        main.pack(fill=tk.BOTH, expand=True)
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=2)

        # Left: Card for input/actions
        card = tk.Frame(main, bg=CARD_BG, bd=2, relief=tk.GROOVE, highlightbackground=CARD_BORDER, highlightthickness=1)
        card.grid(row=0, column=0, sticky="nsew", padx=(30,15), pady=30)
        card.grid_rowconfigure(6, weight=1)
        card.grid_columnconfigure(1, weight=1)

        # Input fields
        tk.Label(card, text="Input Directory", bg=CARD_BG, fg=PRIMARY, font=LABEL_FONT).grid(row=0, column=0, sticky="w", pady=(10,2), padx=10, columnspan=2)
        self.dump_folder_var = tk.StringVar()
        self._folder_picker(card, "Dump Folder:", self.dump_folder_var, row=1)

        tk.Label(card, text="Plan File", bg=CARD_BG, fg=PRIMARY, font=LABEL_FONT).grid(row=2, column=0, sticky="w", pady=(20,2), padx=10, columnspan=2)
        self.plan_var = tk.StringVar()
        self._file_picker(card, "Plan File:", self.plan_var, row=3)

        # Action buttons
        self.run_modified_btn = tk.Button(card, text="Generate Report Excel", bg=BUTTON_BG, fg=BUTTON_FG, font=BUTTON_FONT, activebackground=BUTTON_ACTIVE, relief=tk.FLAT, command=self.run_modified_excel, padx=8, pady=8, bd=0, highlightthickness=0)
        self.run_modified_btn.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=(30,5))
        self.open_modified_btn = tk.Button(card, text="Open Report Output", bg=BUTTON_BG, fg=BUTTON_FG, font=BUTTON_FONT, activebackground=BUTTON_ACTIVE, relief=tk.FLAT, command=self.open_modified_output, state=tk.DISABLED, padx=8, pady=8, bd=0, highlightthickness=0)
        self.open_modified_btn.grid(row=5, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self.run_summary_btn = tk.Button(card, text="Generate Summary Report", bg=BUTTON_BG, fg=BUTTON_FG, font=BUTTON_FONT, activebackground=BUTTON_ACTIVE, relief=tk.FLAT, command=self.run_summary_report, padx=8, pady=8, bd=0, highlightthickness=0)
        self.run_summary_btn.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self.open_summary_btn = tk.Button(card, text="Open Summary Report", bg=BUTTON_BG, fg=BUTTON_FG, font=BUTTON_FONT, activebackground=BUTTON_ACTIVE, relief=tk.FLAT, command=self.open_summary_output, state=tk.DISABLED, padx=8, pady=8, bd=0, highlightthickness=0)
        self.open_summary_btn.grid(row=7, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        # Right: Console output
        console_card = tk.Frame(main, bg="#fff", bd=2, relief=tk.GROOVE, highlightbackground=CARD_BORDER, highlightthickness=1)
        console_card.grid(row=0, column=1, sticky="nsew", padx=(10,30), pady=30)
        console_card.grid_rowconfigure(1, weight=1)
        console_card.grid_columnconfigure(0, weight=1)
        tk.Label(console_card, text="Console Output", bg="#fff", fg=PRIMARY, font=LABEL_FONT).grid(row=0, column=0, sticky="w", padx=10, pady=(10,2))
        self.console = ScrolledText(console_card, height=20, bg=CONSOLE_BG, fg=CONSOLE_FG, font=CONSOLE_FONT, borderwidth=0, relief="flat", wrap=tk.WORD, padx=8, pady=8)
        self.console.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.console.config(state=tk.DISABLED)

        # Footer
        footer = tk.Frame(self, bg=FOOTER_BG, height=30)
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(footer, text="Developed by Rishav Raj", bg=FOOTER_BG, fg=FOOTER_FG, font=FOOTER_FONT).pack(pady=2)

        # Status bar (optional)
        self.status_var = tk.StringVar(value="Ready.")
        status_bar = tk.Label(self, textvariable=self.status_var, font=("Segoe UI Variable", 11, "italic"), anchor="w", bg="#23272e", fg="#00ff99", padx=8, pady=4)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.create_gui_layout()
        self.modified_output = None
        self.summary_output = None

    def create_gui_layout(self):
        pass

    def _folder_picker(self, parent, label, var, row):
        tk.Label(parent, text=label, bg=parent["bg"], font=LABEL_FONT).grid(row=row, column=0, sticky="w", padx=10, pady=4)
        entry = tk.Entry(parent, textvariable=var, font=ENTRY_FONT, bg="#fff")
        entry.grid(row=row, column=1, sticky="ew", padx=(0,10), pady=4)
        parent.grid_columnconfigure(1, weight=1)
        def browse():
            folder = filedialog.askdirectory(title=label)
            if folder:
                var.set(folder)
        tk.Button(parent, text="Browse", bg="#007acc", fg="#fff", font=BUTTON_FONT, activebackground="#005999", relief=tk.FLAT, command=browse, padx=6, pady=4, bd=0, highlightthickness=0).grid(row=row, column=2, padx=(0,10), pady=4)

    def _file_picker(self, parent, label, var, row):
        tk.Label(parent, text=label, bg=parent["bg"], font=LABEL_FONT).grid(row=row, column=0, sticky="w", padx=10, pady=4)
        entry = tk.Entry(parent, textvariable=var, font=ENTRY_FONT, bg="#fff")
        entry.grid(row=row, column=1, sticky="ew", padx=(0,10), pady=4)
        parent.grid_columnconfigure(1, weight=1)
        def browse():
            file = filedialog.askopenfilename(title=label, filetypes=[("Excel files", "*.xlsx")])
            if file:
                var.set(file)
        tk.Button(parent, text="Browse", bg="#007acc", fg="#fff", font=BUTTON_FONT, activebackground="#005999", relief=tk.FLAT, command=browse, padx=6, pady=4, bd=0, highlightthickness=0).grid(row=row, column=2, padx=(0,10), pady=4)

    def log(self, message, level="info"):
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)
        if level == "error":
            logger.error(message)
        elif level == "warning":
            logger.warning(message)
        else:
            logger.info(message)

    def set_status(self, message, level="info"):
        self.status_var.set(message)
        self.log(message, level=level)

    def run_modified_excel(self):
        dump_folder = self.dump_folder_var.get()
        plan_file = self.plan_var.get()
        if not dump_folder or not plan_file:
            messagebox.showerror("Missing Input", "Please select both the dump folder and plan file.")
            return
        self.set_status("Running modified_excel.py...")
        self.run_modified_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._run_modified_excel_thread, args=(dump_folder, plan_file)).start()

    def _run_modified_excel_thread(self, dump_folder, plan_file):
        try:
            template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report_template.xlsx")
            modified_excel.get_resource_path = lambda x: template_path if 'template' in x else x
            self.set_status(f"Selected dump folder: {dump_folder}")
            self.set_status(f"Selected plan file: {plan_file}")
            modified_excel.select_excel_files(dump_folder, plan_file)
            files = [f for f in os.listdir('.') if f.startswith('template_copy_') and f.endswith('.xlsx')]
            if files:
                latest = max(files, key=os.path.getctime)
                self.modified_output = os.path.abspath(latest)
                self.open_modified_btn.config(state=tk.NORMAL)
                self.set_status(f"Report Excel generated: {latest}")
            else:
                self.set_status("No output file generated.")
        except Exception as e:
            err_msg = f"Error: {e}\n{traceback.format_exc()}"
            self.set_status(err_msg, level="error")
            messagebox.showerror("Error", str(e))
            logger.exception("Exception in run_modified_excel thread")
        finally:
            self.run_modified_btn.config(state=tk.NORMAL)

    def open_modified_output(self):
        try:
            if self.modified_output and os.path.exists(self.modified_output):
                webbrowser.open(self.modified_output)
                self.set_status(f"Opened: {self.modified_output}")
        except Exception as e:
            err_msg = f"Error: {e}\n{traceback.format_exc()}"
            self.set_status(err_msg, level="error")
            logger.exception("Exception in open_modified_output")

    def run_summary_report(self):
        if not self.modified_output:
            messagebox.showerror("Missing Output", "Please generate the modified Excel first.")
            return
        self.set_status("Running summary_report_tool.py...")
        self.run_summary_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._run_summary_report_thread, args=(self.modified_output,)).start()

    def _run_summary_report_thread(self, modified_file):
        try:
            template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "summary_report_template.xlsx")
            summary_report_tool.resource_path = lambda x: template_path if 'template' in x else x
            self.set_status(f"Selected modified file: {modified_file}")
            summary_report_tool.extract_distributor_data(modified_file)
            self.summary_output = modified_file
            self.open_summary_btn.config(state=tk.NORMAL)
            self.set_status(f"Summary Report generated: {os.path.basename(modified_file)}")
        except Exception as e:
            err_msg = f"Error: {e}\n{traceback.format_exc()}"
            self.set_status(err_msg, level="error")
            messagebox.showerror("Error", str(e))
            logger.exception("Exception in run_summary_report thread")
        finally:
            self.run_summary_btn.config(state=tk.NORMAL)

    def open_summary_output(self):
        try:
            if self.summary_output and os.path.exists(self.summary_output):
                webbrowser.open(self.summary_output)
                self.set_status(f"Opened: {self.summary_output}")
        except Exception as e:
            err_msg = f"Error: {e}\n{traceback.format_exc()}"
            self.set_status(err_msg, level="error")
            logger.exception("Exception in open_summary_output")

if __name__ == "__main__":
    app = PhenomenalGUI()
    app.mainloop() 