# python
import os
import re
import sys
import json
import time
import shutil
import tempfile
import platform
import subprocess
import configparser
import tkinter as tk
from threading import Thread
from datetime import datetime
from tkinter import filedialog, messagebox
from datetime import datetime
import tempfile
import math

import requests
import gspread
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound, APIError
from google.oauth2 import service_account

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from dotenv import load_dotenv

# Importa la nuova classe Updater (presupponendo che il file updater.py esista)
from updater import Updater

# --- Configurazione e Dipendenze ---
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

load_dotenv()

CURRENT_VERSION = "1.8.2"

# regex globali (case-insensitive dove opportuno)
EMAIL_RE = re.compile(r'[\w\.+-]+@[\w\.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?', re.I)
PHONE_RE = re.compile(r'(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{0,4}')
DATE_RE = re.compile(r'\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b')
LEADING_PREFIX_RE = re.compile(r'^\s*\d+\s*-\s*')
CHANNEL_LINE_RE = re.compile(r'(SISTEMA INVIO ONLINE|SITO WEB|FACEBOOK|INSTAGRAM|MAIL DOMANDE|INVI|GIAN|JAN|INVI O|MIDDLE|SHORT|NEW|DRAGOTA|TIKTOK|WHATSAPP|ONLINE|EMAIL_INTERNAL|DIGITURBO|PIGITURBO|MECCANISMO TRAMITE UNICO)', re.I)

# parole/piccoli token che possono far parte del cognome (particle) e che non sono occupazione
SURNAME_PARTICLES = {'di','da','de','del','della','la','le','van','von','mac','mc','san','st','d\''}

class BotApp(ttk.Frame):
    """
    Classe principale dell'applicazione, gestisce l'intera interfaccia utente
    e la logica del bot.
    """

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(fill=BOTH, expand=True)

        self.user_data_path = self._get_user_data_path()
        self.config_file = os.path.join(self.user_data_path, 'config.ini')

        # Inizializza TUTTE le variabili qui in modo uniforme
        self.file_path_var = tk.StringVar()
        self.svuota_file_var = tk.BooleanVar()
        self.tutorial_state_var = tk.BooleanVar()
        self.update_available = tk.BooleanVar(value=False)
        
        # Variabili per la modalità sheets - usa nomi coerenti
        self.source_mode = tk.StringVar(value="file")
        self.source_spreadsheet_var = tk.StringVar()  # Sorgente sheets
        self.source_worksheet_var = tk.StringVar()    # Sorgente sheets
        self.dest_spreadsheet_var = tk.StringVar()    # Destinazione
        self.dest_worksheet_var = tk.StringVar()      # Destinazione
        
        # Mantieni le variabili legacy per compatibilità
        self.spreadsheet_name_var = self.dest_spreadsheet_var  # Alias per compatibilità
        self.worksheet_name_var = self.dest_worksheet_var      # Alias per compatibilità
        self.src_spreadsheet_var = self.source_spreadsheet_var # Alias per compatibilità
        self.src_worksheet_var = self.source_worksheet_var     # Alias per compatibilità
        
        self.running_thread = None
        self.headers_set = False 
        self.create_gui_elements()
        self.updater = Updater(current_version=CURRENT_VERSION, log_callback=self._log_message)

        self._create_user_config_directory()
        self.tutorial_state_var.set(self._load_config_boolean('SETTINGS', 'tutorial_shown', False))
        
        self.load_initial_configuration()

        if not self._load_config_boolean('SETTINGS', 'tutorial_shown', False):
            self.master.after(500, self.show_tutorial_window)
            self._save_config_value('SETTINGS', 'tutorial_shown', True)

        self.master.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Avvia la verifica degli aggiornamenti in un thread separato
        Thread(target=lambda: self.updater.check_for_updates(on_update_available=lambda: self.update_available.set(True)), daemon=True).start()
        try:
            self._load_local_settings()
        except Exception:
            pass

    def _get_local_settings_path(self):
        # Assicurati che self.user_data_path esista nella tua classe.
        # Se non esiste, usa la home
        base = getattr(self, 'user_data_path', None) or os.path.expanduser("~")
        os.makedirs(base, exist_ok=True)
        return os.path.join(base, "mxttjaw_bot_settings.json")

    def _persist_local_settings(self):
        """
        Salva localmente le impostazioni in un file JSON dentro self.user_data_path per persistenza tra sessioni.
        """
        try:
            cfg = {
                "source_mode": getattr(self, "source_mode", tk.StringVar(value="file")).get() if hasattr(self, "source_mode") else "",
                "file_path": getattr(self, "file_path_var", tk.StringVar()).get() if hasattr(self, "file_path_var") else "",
                "source_spreadsheet": getattr(self, "source_spreadsheet_var", tk.StringVar()).get() if hasattr(self, "source_spreadsheet_var") else "",
                "source_worksheet": getattr(self, "source_worksheet_var", tk.StringVar()).get() if hasattr(self, "source_worksheet_var") else "",
                "dest_spreadsheet": getattr(self, "dest_spreadsheet_var", tk.StringVar()).get() if hasattr(self, "dest_spreadsheet_var") else "",
                "dest_worksheet": getattr(self, "dest_worksheet_var", tk.StringVar()).get() if hasattr(self, "dest_worksheet_var") else ""
            }
            p = self._get_local_settings_path()
            with open(p, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh, ensure_ascii=False, indent=2)
            self._log_message(f"Impostazioni locali salvate: {p}")
        except Exception as e:
            self._log_message(f"Errore salvando impostazioni locali: {e}")

    def _load_local_settings(self):
        """
        Carica le impostazioni locali, se esistono, e le ripopola nelle variabili
        dell'interfaccia (self.*_var). Chiamalo in inizializzazione della GUI.
        """
        try:
            p = self._get_local_settings_path()
            if not os.path.exists(p):
                return
            with open(p, "r", encoding="utf-8") as fh:
                cfg = json.load(fh)
            # usa set() sulle StringVar se esistono
            if hasattr(self, "source_mode") and cfg.get("source_mode") is not None:
                try: self.source_mode.set(cfg.get("source_mode",""))
                except: pass
            if hasattr(self, "file_path_var") and cfg.get("file_path") is not None:
                try: self.file_path_var.set(cfg.get("file_path",""))
                except: pass
            if hasattr(self, "source_spreadsheet_var") and cfg.get("source_spreadsheet") is not None:
                try: self.source_spreadsheet_var.set(cfg.get("source_spreadsheet",""))
                except: pass
            if hasattr(self, "source_worksheet_var") and cfg.get("source_worksheet") is not None:
                try: self.source_worksheet_var.set(cfg.get("source_worksheet",""))
                except: pass
            if hasattr(self, "dest_spreadsheet_var") and cfg.get("dest_spreadsheet") is not None:
                try: self.dest_spreadsheet_var.set(cfg.get("dest_spreadsheet",""))
                except: pass
            if hasattr(self, "dest_worksheet_var") and cfg.get("dest_worksheet") is not None:
                try: self.dest_worksheet_var.set(cfg.get("dest_worksheet",""))
                except: pass

            self._log_message(f"Impostazioni locali caricate: {p}")
        except Exception as e:
            # non bloccante
            try: self._log_message(f"Errore caricando impostazioni locali: {e}")
            except: pass

    # --- Metodi per la gestione della GUI ---

    def create_gui_elements(self):
        """Costruisce tutti i widget dell'interfaccia utente."""
        self.master.title(f"Bot Email - Gestione Dati (v{CURRENT_VERSION})")
        self.master.geometry("700x460")
        self.master.resizable(True, True)
        
        self.top_menu_frame = ttk.Frame(self, height=36)
        self.top_menu_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)

        self.main_content_frame = ttk.Frame(self)
        self.main_content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self.main_content_frame.grid_rowconfigure(1, weight=1)
        self.main_content_frame.grid_columnconfigure(0, weight=1)

        self._create_dropdown_menu()
        self._create_status_area()
        self._create_file_selection_area()
        self._create_log_area()
        self._create_action_buttons()

    def _create_dropdown_menu(self):
        """Crea il menu a tendina usando un pulsante e un widget Menu."""
        menu_btn = ttk.Button(self.top_menu_frame, text="☰", width=3, bootstyle="secondary")
        menu_btn.pack(side=tk.LEFT)

        menu = tk.Menu(self.master, tearoff=0)
        menu.add_command(label="Avvia Bot", command=self.start_bot_thread)
        menu.add_command(label="Opzioni", command=self.open_options_window)
        menu.add_command(label="Tutorial", command=self.show_tutorial_window)
        menu.add_command(label="Controlla Aggiornamenti", command=lambda: Thread(target=lambda: self.updater.check_for_updates(on_update_available=lambda: self.update_available.set(True)), daemon=True).start())
        menu.add_separator()
        menu.add_command(label="Esci", command=self.master.quit)
        
        def show_menu():
            """Mostra il menu a tendina nella posizione corretta."""
            try:
                menu.tk_popup(menu_btn.winfo_rootx(), menu_btn.winfo_rooty() + menu_btn.winfo_height())
            finally:
                menu.grab_release()
        
        menu_btn.config(command=show_menu)

    def _create_status_area(self):
        """Area di stato a destra del top frame: versione e stato aggiornamenti."""
        status_frame = ttk.Frame(self.top_menu_frame)
        status_frame.pack(side=tk.RIGHT)

        self.status_label = ttk.Label(status_frame, text=f"v{CURRENT_VERSION}", bootstyle="muted")
        self.status_label.pack(side=tk.RIGHT, padx=(8,0))

        self.update_progress = ttk.Progressbar(status_frame, length=180, mode='indeterminate')
        self.update_progress.pack(side=tk.RIGHT, padx=(0,8))
        self.update_progress.pack_forget()

    def _toggle_source_mode(self):
        """Mostra o nasconde i campi in base alla sorgente selezionata (file/sheets)."""
        mode = self.source_mode.get()

        if mode == "file":
            # Mostra sezione file, nascondi sezione sheets
            self.path_frame.pack(fill=tk.X, pady=(0, 5))
            self.ss_frame.pack_forget()
            self.ws_frame.pack_forget()
        else:
            # Mostra sezione sheets, nascondi sezione file
            self.path_frame.pack_forget()
            self.ss_frame.pack(fill=tk.X, pady=(5, 5))
            self.ws_frame.pack(fill=tk.X, pady=(5, 5))

    # def _create_file_selection_area(self):
    #     """Crea il frame per la selezione del file e i campi per foglio di calcolo/lavoro."""
    #     file_frame = ttk.LabelFrame(self.main_content_frame, text="Configurazione Dati e Fogli", padding=(15, 10))
    #     file_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 15))
        
    #     # Path + Source mode
    #     path_frame = ttk.Frame(file_frame)
    #     path_frame.pack(fill=tk.X, pady=(0, 5))

    #     ttk.Label(path_frame, text="Sorgente dati:").pack(side=tk.LEFT)dd
    #     ttk.OptionMenu(
    #         path_frame,
    #         self.source_mode,
    #         "file",
    #         "file",
    #         "sheets"
    #     ).pack(side=tk.LEFT, padx=(5, 15))

    #     # File chooser (per modalità file)
    #     ttk.Label(path_frame, text="File Dati:").pack(side=tk.LEFT)
    #     self.file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, state='readonly')
    #     self.file_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
    #     self.browse_button = ttk.Button(path_frame, text="Scegli File", command=self.select_file, bootstyle="primary")
    #     self.browse_button.pack(side=tk.RIGHT)

    #     # ============================================================
    #     #  Sorgente Google Sheets (solo per modalità 'sheets')
    #     #  -> l'utente può comunque compilare questi campi; se la modalità
    #     #     è 'sheets' verranno usati come sorgente
    #     # ============================================================
    #     src_frame = ttk.Frame(file_frame)
    #     src_frame.pack(fill=tk.X, pady=(5, 5))

    #     ttk.Label(src_frame, text="Spreadsheet Sorgente (solo 'sheets'):").pack(side=tk.LEFT)
    #     self.src_ss_entry = ttk.Entry(src_frame, textvariable=self.src_spreadsheet_var)
    #     self.src_ss_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

    #     src_ws_frame = ttk.Frame(file_frame)
    #     src_ws_frame.pack(fill=tk.X, pady=(5, 5))

    #     ttk.Label(src_ws_frame, text="Worksheet Sorgente:").pack(side=tk.LEFT)
    #     self.src_ws_entry = ttk.Entry(src_ws_frame, textvariable=self.src_worksheet_var)
    #     self.src_ws_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

    #     # ============================================================
    #     #  Destinazione Google Sheets (usata sia in modalità 'file' che 'sheets')
    #     # ============================================================
    #     ss_frame = ttk.Frame(file_frame)
    #     ss_frame.pack(fill=tk.X, pady=(5, 5))
        
    #     ttk.Label(ss_frame, text="Nome Foglio di Calcolo (destinazione):").pack(side=tk.LEFT)
    #     self.ss_entry = ttk.Entry(ss_frame, textvariable=self.spreadsheet_name_var)
    #     self.ss_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        
    #     ws_frame = ttk.Frame(file_frame)
    #     ws_frame.pack(fill=tk.X, pady=(5, 5))
        
    #     ttk.Label(ws_frame, text="Nome Foglio di Lavoro (destinazione):").pack(side=tk.LEFT)
    #     self.ws_entry = ttk.Entry(ws_frame, textvariable=self.worksheet_name_var)
    #     self.ws_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

    def _create_file_selection_area(self):
        """Crea il frame per la selezione del file e i campi per foglio di calcolo/lavoro (sorgente + destinazione)."""
        file_frame = ttk.LabelFrame(self.main_content_frame, text="Configurazione Dati e Fogli", padding=(15, 10))
        file_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 15))

        # RIMUOVI queste righe - le variabili sono già inizializzate in __init__
        # source_mode_default = self._load_config_string('SETTINGS', 'source_mode', 'file') if hasattr(self, '_load_config_string') else 'file'
        # self.source_mode = tk.StringVar(value=source_mode_default)
        # self.file_path_var = getattr(self, 'file_path_var', tk.StringVar())
        # self.source_spreadsheet_var = getattr(self, 'source_spreadsheet_var', tk.StringVar(value=self._load_config_string('SHEETS', 'source_spreadsheet', '') if hasattr(self,'_load_config_string') else ''))
        # self.source_worksheet_var = getattr(self, 'source_worksheet_var', tk.StringVar(value=self._load_config_string('SHEETS', 'source_worksheet', '') if hasattr(self,'_load_config_string') else ''))
        # self.dest_spreadsheet_var = getattr(self, 'spreadsheet_name_var', tk.StringVar(value=self._load_config_string('SHEETS', 'dest_spreadsheet', '') if hasattr(self,'_load_config_string') else ''))
        # self.dest_worksheet_var = getattr(self, 'worksheet_name_var', tk.StringVar(value=self._load_config_string('SHEETS', 'dest_worksheet', '') if hasattr(self,'_load_config_string') else ''))

        # Row: source mode
        source_frame = ttk.Frame(file_frame)
        source_frame.pack(fill=tk.X, pady=(0,5))
        ttk.Label(source_frame, text="Sorgente dati:").pack(side=tk.LEFT)
        ttk.OptionMenu(source_frame, self.source_mode, self.source_mode.get(), "file", "sheets").pack(side=tk.LEFT, padx=5)

        # Row: file chooser (when source == file)
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(path_frame, text="File Dati (sorgente):").pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, state='readonly')
        self.file_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        self.browse_button = ttk.Button(path_frame, text="Scegli File", command=self.select_file, bootstyle="primary")
        self.browse_button.pack(side=tk.RIGHT)

        # Row: source spreadsheet + worksheet (used when source_mode == sheets)
        src_sheet_frame = ttk.Frame(file_frame)
        src_sheet_frame.pack(fill=tk.X, pady=(5, 5))
        ttk.Label(src_sheet_frame, text="Spreadsheet sorgente (nome):").pack(side=tk.LEFT)
        self.src_ss_entry = ttk.Entry(src_sheet_frame, textvariable=self.source_spreadsheet_var)
        self.src_ss_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Label(src_sheet_frame, text="Worksheet sorgente:").pack(side=tk.LEFT, padx=(10,0))
        self.src_ws_entry = ttk.Entry(src_sheet_frame, textvariable=self.source_worksheet_var, width=20)
        self.src_ws_entry.pack(side=tk.LEFT, padx=(5,0))

        # Row: destination spreadsheet + worksheet (where to write)
        dst_sheet_frame = ttk.Frame(file_frame)
        dst_sheet_frame.pack(fill=tk.X, pady=(5, 5))
        ttk.Label(dst_sheet_frame, text="Spreadsheet destinazione (nome):").pack(side=tk.LEFT)
        self.dst_ss_entry = ttk.Entry(dst_sheet_frame, textvariable=self.dest_spreadsheet_var)
        self.dst_ss_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Label(dst_sheet_frame, text="Worksheet destinazione:").pack(side=tk.LEFT, padx=(10,0))
        self.dst_ws_entry = ttk.Entry(dst_sheet_frame, textvariable=self.dest_worksheet_var, width=20)
        self.dst_ws_entry.pack(side=tk.LEFT, padx=(5,0))

        # Aggiungi questa funzione per salvare automaticamente
        def save_all_settings():
            """Salva tutte le impostazioni correnti."""
            try:
                self._save_config_value('SETTINGS', 'source_mode', self.source_mode.get())
                self._save_config_value('SHEETS', 'source_spreadsheet', self.source_spreadsheet_var.get())
                self._save_config_value('SHEETS', 'source_worksheet', self.source_worksheet_var.get())
                self._save_config_value('SHEETS', 'dest_spreadsheet', self.dest_spreadsheet_var.get())
                self._save_config_value('SHEETS', 'dest_worksheet', self.dest_worksheet_var.get())
                self._save_config_value('PATHS', 'dati_emails_path', self.file_path_var.get() or '')
            except Exception as e:
                print(f"Errore durante il salvataggio delle impostazioni: {e}")

        def on_source_change(*args):
            mode = self.source_mode.get()
            if mode == 'file':
                self.file_entry.configure(state='readonly')
                self.browse_button.configure(state='normal')
                self.src_ss_entry.configure(state='disabled')
                self.src_ws_entry.configure(state='disabled')
            else:
                self.file_entry.configure(state='disabled')
                self.browse_button.configure(state='disabled')
                self.src_ss_entry.configure(state='normal')
                self.src_ws_entry.configure(state='normal')
            
            save_all_settings()

        # Aggiungi tracciamento per salvare automaticamente quando i valori cambiano
        self.source_mode.trace_add('write', on_source_change)
        self.source_spreadsheet_var.trace_add('write', lambda *args: save_all_settings())
        self.source_worksheet_var.trace_add('write', lambda *args: save_all_settings())
        self.dest_spreadsheet_var.trace_add('write', lambda *args: save_all_settings())
        self.dest_worksheet_var.trace_add('write', lambda *args: save_all_settings())
        
        on_source_change()
    
    # Aggiungi questo metodo alla classe BotApp
    def _on_closing(self):
        """Salva le impostazioni quando l'applicazione viene chiusa."""
        try:
            self._save_config_value('SETTINGS', 'source_mode', self.source_mode.get())
            self._save_config_value('SHEETS', 'source_spreadsheet', self.source_spreadsheet_var.get())
            self._save_config_value('SHEETS', 'source_worksheet', self.source_worksheet_var.get())
            self._save_config_value('SHEETS', 'dest_spreadsheet', self.dest_spreadsheet_var.get())
            self._save_config_value('SHEETS', 'dest_worksheet', self.dest_worksheet_var.get())
            self._save_config_value('PATHS', 'dati_emails_path', self.file_path_var.get() or '')
        except Exception as e:
            print(f"Errore durante il salvataggio finale: {e}")
        
        self.master.quit()

  
    




    def _create_log_area(self):
        """Crea il frame per i messaggi di log."""
        log_frame = ttk.LabelFrame(self.main_content_frame, text="Log e Stato", padding=(15, 10))
        log_frame.grid(row=1, column=0, sticky=tk.NSEW, pady=(0, 15))
        
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.log_area = tk.Text(log_frame, wrap=WORD, yscrollcommand=scrollbar.set, state='disabled', font=("Helvetica", 10), relief=tk.FLAT)
        self.log_area.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.log_area.yview)

    def _create_action_buttons(self):
        """Crea i pulsanti per l'avvio e l'anteprima del bot."""
        btn_frame = ttk.Frame(self.main_content_frame)
        btn_frame.grid(row=2, column=0, pady=(15, 0), sticky=tk.EW)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        self.preview_button = ttk.Button(btn_frame, text="Anteprima Dati", command=self.show_data_preview, bootstyle="info")
        self.preview_button.grid(row=0, column=0, padx=(0, 5), sticky=tk.E)

        self.run_button = ttk.Button(btn_frame, text="Avvia Bot", command=self.start_bot_thread, bootstyle="success")
        self.run_button.grid(row=0, column=1, padx=(5, 0), sticky=tk.W)

        self.update_btn = ttk.Button(self.main_content_frame, text="Aggiorna Bot", command=self.start_update, bootstyle="warning")
        self.update_btn.grid(row=3, column=0, pady=10, sticky=tk.EW)
        self.update_btn.grid_remove() 
        
        self.update_available.trace_add('write', self.handle_update_button_visibility)

    def handle_update_button_visibility(self, *args):
        """Mostra o nasconde il pulsante di aggiornamento in base allo stato."""
        if self.update_available.get():
            self.update_btn.grid()
            self.status_label.config(text="Aggiornamento disponibile")
        else:
            self.update_btn.grid_remove()
            self.status_label.config(text=f"v{CURRENT_VERSION}")

    def start_bot_thread(self):
        """Avvia l'esecuzione del bot in un thread separato per non bloccare la GUI."""
        if self.running_thread and self.running_thread.is_alive():
            messagebox.showwarning("Attenzione", "Il bot è già in esecuzione.")
            return
        
        self.running_thread = Thread(target=self.run_bot, daemon=True)
        self.running_thread.start()

    def _get_user_data_path(self):
        """Restituisce il percorso della directory dati specifica per l'utente,
        in base al sistema operativo."""
        if platform.system() == "Windows":
            return os.path.join(os.environ["APPDATA"], "DataToSheets")
        elif platform.system() == "Darwin":
            return os.path.join(os.path.expanduser("~"), "Library", "Application Support", "DataToSheets")
        else: # Linux
            return os.path.join(os.path.expanduser("~"), ".config", "DataToSheets")

    def _create_user_config_directory(self):
        """Crea la directory utente e i file di configurazione essenziali se non esistono."""
        if not os.path.exists(self.user_data_path):
            os.makedirs(self.user_data_path)
            self._log_message(f"Creata cartella di configurazione utente in: {self.user_data_path}")

        env_path = os.path.join(self.user_data_path, '.env')
        if not os.path.exists(env_path):
            with open(env_path, 'w') as f:
                f.write("GOOGLE_CREDENTIALS_FILE = credentials.json\n")
                f.write("SPREADSHEET_NAME = nome_del_tuo_foglio_di_calcolo\n")
            self._log_message(f"Creato file .env di esempio in {self.user_data_path}.")

    def load_initial_configuration(self):
        """Carica tutte le configurazioni all'avvio."""
        # Carica percorso file
        self.file_path_var.set(self._load_config_string('PATHS', 'dati_emails_path', ''))
        self.svuota_file_var.set(self._load_config_boolean('SETTINGS', 'svuota_file', False))
        
        # Carica modalità sorgente
        source_mode = self._load_config_string('SETTINGS', 'source_mode', 'file')
        self.source_mode.set(source_mode)
        
        # Carica configurazioni sheets SORGENTE
        source_spreadsheet = self._load_config_string('SHEETS', 'source_spreadsheet', '')
        self.source_spreadsheet_var.set(source_spreadsheet)
        
        source_worksheet = self._load_config_string('SHEETS', 'source_worksheet', '')
        self.source_worksheet_var.set(source_worksheet)
        
        # Carica configurazioni sheets DESTINAZIONE
        dest_spreadsheet = self._load_config_string('SHEETS', 'dest_spreadsheet', '')
        if not dest_spreadsheet:
            # Fallback per compatibilità con versioni precedenti
            spreadsheet_name_from_env = os.getenv("SPREADSHEET_NAME", "")
            dest_spreadsheet = self._load_config_string('SHEETS', 'spreadsheet_name', spreadsheet_name_from_env or '')
        self.dest_spreadsheet_var.set(dest_spreadsheet)
        
        dest_worksheet = self._load_config_string('SHEETS', 'dest_worksheet', '')
        if not dest_worksheet:
            # Fallback: usa il mese corrente come default
            nomi_mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", 
                        "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
            nome_foglio_mese = nomi_mesi[datetime.now().month - 1]
            dest_worksheet = self._load_config_string('SHEETS', 'worksheet_name', nome_foglio_mese)
        self.dest_worksheet_var.set(dest_worksheet)
        
        # Carica tema
        tema_salvato = self._load_config_string('SETTINGS', 'theme', 'lumen')
        self.master.style.theme_use(tema_salvato)
        
        if self.file_path_var.get() and os.path.exists(self.file_path_var.get()):
            self._log_message("Percorso file caricato automaticamente.")
        else:
            self._log_message("Nessun percorso file salvato. Seleziona il file dei dati.")

    def _log_message(self, message):
        """Aggiorna la Textbox della GUI con un messaggio e un timestamp."""
        try:
            # Decodifica eventuali caratteri speciali
            if isinstance(message, bytes):
                message = message.decode('utf-8', errors='replace')
            message = message.replace('Ã¨', 'è').replace('Ã', 'à') 
        except Exception:
            pass
        
        timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        full_message = f"{timestamp} {message}\n"
        
        try:
            self.log_area.configure(state='normal')
            self.log_area.insert(tk.END, full_message)
            self.log_area.configure(state='disabled')
            self.log_area.see(tk.END)
        except Exception:
            print(full_message, end='')

    def _load_config(self):
        """Carica la configurazione da un file."""
        config = configparser.ConfigParser()
        config.read(self.config_file)
        return config

    def _save_config(self, config):
        """Salva la configurazione corrente in un file."""
        try:
            with open(self.config_file, 'w') as configfile:
                config.write(configfile)
        except Exception as e:
            self._log_message(f"Errore durante il salvataggio della configurazione: {e}")

    def _save_config_value(self, section, key, value):
        """Salva un singolo valore nel file di configurazione."""
        config = self._load_config()
        if section not in config:
            config[section] = {}
        config[section][key] = str(value)
        self._save_config(config)

    def _load_config_string(self, section, key, fallback=None):
        """Carica una stringa da un file di configurazione."""
        config = self._load_config()
        return config.get(section, key, fallback=fallback)

    def _load_config_boolean(self, section, key, fallback=False):
        """Carica un booleano da un file di configurazione."""
        config = self._load_config()
        return config.getboolean(section, key, fallback=fallback)
        
    def _create_options_window(self, options_window):
        """Costruisce la finestra delle opzioni."""
        options_window.title("Opzioni")
        options_window.geometry("400x200")
        options_window.grab_set()

        options_frame = ttk.Frame(options_window, padding=20)
        options_frame.pack(fill=BOTH, expand=True)

        temi = self.master.style.theme_names()
        theme_var = tk.StringVar(value=self.master.style.theme_use())
        
        ttk.Label(options_frame, text="Seleziona tema:").pack(anchor=W)
        theme_menu = ttk.OptionMenu(options_frame, theme_var, theme_var.get(), *temi)
        theme_menu.pack(fill=X, pady=(0, 10))

        def on_theme_change(*args):
            """Cambia il tema dell'applicazione."""
            nuovo_tema = theme_var.get()
            self.master.style.theme_use(nuovo_tema)
            self._save_config_value('SETTINGS', 'theme', nuovo_tema)

        theme_var.trace_add('write', on_theme_change)

        # Opzione svuota file
        svuota_frame = ttk.Frame(options_frame)
        svuota_frame.pack(fill=X, pady=(0, 10))
        ttk.Checkbutton(svuota_frame, text="Svuota il file di testo dopo l'invio", variable=self.svuota_file_var).pack(side=LEFT)
        
        def on_svuota_change(*args):
            """Salva lo stato della checkbox."""
            self._save_config_value('SETTINGS', 'svuota_file', self.svuota_file_var.get())
            
        self.svuota_file_var.trace_add('write', on_svuota_change)
        
        ttk.Button(options_frame, text="Chiudi", command=options_window.destroy, bootstyle="info").pack(pady=10)

    def open_options_window(self):
        """Apre la finestra delle opzioni."""
        options_window = tk.Toplevel(self.master)
        self._create_options_window(options_window)
    
    def show_tutorial_window(self):
        """Mostra una finestra di tutorial con le istruzioni per la configurazione."""
        tutorial_window = tk.Toplevel(self.master)
        tutorial_window.title("Tutorial: Configurazione Bot")
        tutorial_window.geometry("700x500")
        tutorial_window.transient(self.master)
        tutorial_window.grab_set()

        tutorial_frame = ttk.Frame(tutorial_window, padding=20)
        tutorial_frame.pack(fill=tk.BOTH, expand=True)
        
        tutorial_text = tk.Text(tutorial_frame, wrap=tk.WORD, font=("Helvetica", 10), relief=tk.FLAT, borderwidth=0)
        tutorial_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        tutorial_text.insert(tk.END, "Benvenuto nel tutorial di configurazione!\n\n", ("bold"))
        tutorial_text.insert(tk.END, "Per garantire che ogni utente abbia la propria configurazione, il bot ora salva i file in una cartella dedicata del tuo sistema. L'eseguibile, invece, può essere posizionato dove preferisci.\n\n")

        tutorial_text.insert(tk.END, "Passo 1: Abilitare l'API di Google Sheets\n", ("bold"))
        tutorial_text.insert(tk.END, "Vai su Google Cloud Console (https://console.cloud.google.com/) e assicurati di essere nel progetto corretto. Cerca 'Google Sheets API' e abilitala.\n\n")

        tutorial_text.insert(tk.END, "Passo 2: Creare un account di servizio\n", ("bold"))
        tutorial_text.insert(tk.END, "1. Sempre nella Cloud Console, vai su 'IAM e Amministrazione' -> 'Account di servizio'.\n")
        tutorial_text.insert(tk.END, "2. Clicca su 'Crea account di servizio'.\n")
        tutorial_text.insert(tk.END, "3. Assegna un nome (es. 'bot-sheets-service').\n")
        tutorial_text.insert(tk.END, "4. Concedi i permessi necessari (es. 'Editor').\n\n")

        tutorial_text.insert(tk.END, "Passo 3: Scaricare la chiave privata\n", ("bold"))
        tutorial_text.insert(tk.END, "1. Clicca sul nome dell'account di servizio appena creato.\n")
        tutorial_text.insert(tk.END, "2. Vai alla scheda 'Chiavi' e clicca su 'Aggiungi chiave' -> 'Crea nuova chiave'.\n")
        tutorial_text.insert(tk.END, "3. Scegli 'JSON' come tipo di chiave. Il file verrà scaricato automaticamente.\n\n")

        tutorial_text.insert(tk.END, "Passo 4: Caricare il file delle credenziali\n", ("bold"))
        tutorial_text.insert(tk.END, "Ora clicca sul pulsante 'Seleziona File Credenziali' qui sotto per caricare il file JSON appena scaricato. Il bot lo rinominerà in 'credentials.json' e lo posizionerà nella cartella corretta:\n")
        tutorial_text.insert(tk.END, f"{self.user_data_path}\n\n", ("bold"))

        tutorial_text.insert(tk.END, "Passo 5: Condividere il foglio di calcolo\n", ("bold"))
        tutorial_text.insert(tk.END, "Apri il tuo foglio di Google Sheets e clicca su 'Condividi'. Incolla l'indirizzo email dell'account di servizio (lo trovi nel file JSON) e concedi il permesso 'Editor'.\n")
        tutorial_text.configure(state='disabled')
        
        bottom_frame = ttk.Frame(tutorial_window)
        bottom_frame.pack(pady=(0, 10))
        
        creds_btn = ttk.Button(bottom_frame, text="Seleziona File Credenziali", command=self.handle_credentials_file, bootstyle="success")
        creds_btn.pack(side=tk.LEFT, padx=(0, 20))
        
        close_btn = ttk.Button(bottom_frame, text="Chiudi", command=lambda: tutorial_window.destroy(), bootstyle="primary")
        close_btn.pack(side=tk.LEFT)

    def handle_credentials_file(self):
        """Permette all'utente di selezionare il file JSON e lo sposta/rinomina."""
        source_path = filedialog.askopenfilename(
            title="Seleziona il file 'credentials.json' appena scaricato",
            filetypes=[("File JSON", "*.json")]
        )
        if not source_path:
            return

        destination_path = os.path.join(self.user_data_path, "credentials.json")
        os.makedirs(self.user_data_path, exist_ok=True) 

        try:
            # Se già esiste, chiedi conferma prima di sovrascrivere
            if os.path.exists(destination_path):
                overwrite = messagebox.askyesno(
                    "Sovrascrivere file?",
                    f"Esiste già un file 'credentials.json' in:\n{destination_path}\n\nVuoi sostituirlo?"
                )
                if not overwrite:
                    return

            shutil.copyfile(source_path, destination_path)
            self.credentials_path = destination_path
            self._log_message(f"File credenziali copiato e rinominato con successo in: {destination_path}")
            messagebox.showinfo(
                "Successo",
                f"File 'credentials.json' configurato correttamente!\n"
                f"Ora puoi condividere il tuo foglio di calcolo con l'indirizzo email di servizio.\n"
                f"Il file è stato salvato in:\n{self.user_data_path}"
            )
        except Exception as e:
            self._log_message(f"Errore durante la gestione del file credenziali: {e}")
            messagebox.showerror("Errore", f"Errore durante la gestione del file: {e}. Controlla i permessi della cartella.")


    def connect_to_sheets(self):
        """Connessione a Google Sheets tramite gspread."""
        try:
            creds_path = os.path.join(self.user_data_path, os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json"))
            if not os.path.exists(creds_path):
                self._log_message(f"ERRORE: File di credenziali non trovato a: {creds_path}")
                messagebox.showerror("Errore Autenticazione", "File di credenziali 'credentials.json' non trovato.")
                return None
                
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            creds = service_account.Credentials.from_service_account_file(creds_path, scopes=scopes)
            return gspread.authorize(creds)
        except Exception as e:
            self._log_message(f"Errore di connessione a Google Sheets: {e}")
            messagebox.showerror("Errore", f"Impossibile connettersi a Google Sheets: {e}")
            return None
        
    def select_file(self):
        """Permette all'utente di selezionare un file e salva il percorso."""
        file_path = filedialog.askopenfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self._save_config_value('PATHS', 'dati_emails_path', file_path)
            self._log_message(f"File selezionato: {os.path.basename(file_path)}")

    def leggi_da_google_sheets(self, spreadsheet_name: str, worksheet_name: str) -> list:
        """
        Legge righe da uno spreadsheet Google e restituisce una lista di stringhe
        (ogni stringa è l'unione delle celle non vuote della riga).
        Usa gspread con service account (self.credentials_path) se possibile.
        """
        import os
        try:
            self._log_message(f"Lettura dal foglio sorgente: '{spreadsheet_name}' / '{worksheet_name}'")
        except Exception:
            pass

        # connect
        gc = None
        try:
            if hasattr(self, 'connect_to_sheets'):
                gc = self.connect_to_sheets()
        except Exception:
            gc = None

        # se connect_to_sheets non ha funzionato, prova a creare client da credentials.json
        if gc is None:
            try:
                import gspread
                cred_path = getattr(self, 'credentials_path', None)
                if cred_path and os.path.exists(cred_path):
                    gc = gspread.service_account(filename=cred_path)
                else:
                    # tenta il metodo senza file (se l'ambiente è già autenticato)
                    try:
                        gc = gspread.oauth()
                    except Exception:
                        gc = None
            except Exception:
                gc = None

        if not gc:
            msg = "Impossibile creare il client gspread (controlla credentials.json o connect_to_sheets)."
            try:
                self._log_message(f"leggi_da_google_sheets: {msg}")
                messagebox.showerror("Errore", msg)
            except Exception:
                pass
            return []

        # open spreadsheet (by name o by key)
        try:
            try:
                sh = gc.open(spreadsheet_name)
            except Exception:
                # fallback: prova open_by_key (l'utente potrebbe aver passato l'ID)
                try:
                    sh = gc.open_by_key(spreadsheet_name)
                except Exception as e:
                    raise RuntimeError(f"Impossibile aprire lo spreadsheet '{spreadsheet_name}': {e}")

            # worksheet
            try:
                ws = sh.worksheet(worksheet_name) if worksheet_name else sh.sheet1
            except Exception:
                # fallback case-insensitive match
                found = None
                for w in sh.worksheets():
                    if w.title.strip().lower() == (worksheet_name or '').strip().lower():
                        found = w
                        break
                if found:
                    ws = found
                else:
                    # se non trovato, usa sheet1 come fallback (ma avvisa)
                    ws = sh.sheet1

            all_values = ws.get_all_values()
            righe = []
            for row in all_values:
                # unisci tutte le celle non vuote con uno spazio
                testo = ' '.join([c.strip() for c in row if c and str(c).strip()])
                if testo:
                    righe.append(testo)
            self._log_message(f"leggi_da_google_sheets: trovate {len(righe)} righe non vuote.")
            return righe

        except Exception as e:
            self._log_message(f"Errore durante la lettura da Sheets: {e}")
            try:
                messagebox.showerror("Errore", f"Impossibile leggere dallo spreadsheet: {e}")
            except Exception:
                pass
            return []
        
    # def leggi_sheet_sorgente(self, spreadsheet_name: str, worksheet_name: str) -> list:
    #     """Wrapper legacy compatibile: chiama leggi_da_google_sheets."""
    #     return self.leggi_da_google_sheets(spreadsheet_name, worksheet_name)


    # def leggi_file_dati(self, file_path):
    #     """
    #     Legge un file di testo e restituisce lista di righe parsate:
    #     [nome, cognome, eta, occupazione, email, telefono]

    #     Migliorata la rimozione delle "code canale" (es. SISTEMA INVIO, MECCANISMO TRAMITE UNICO,
    #     DIGITURBO / PIGITURBO, NEW ..., MIDDLE SHORT ..., FACEBOOK, etc.), incluse varianti
    #     come 'PIGITURBO X' o parole che finiscono in 'turbo'. Questa porzione viene rimossa
    #     prima del parsing per evitare che finisca in cognome/occupazione.
    #     """
    #     import re, os, shutil

    #     dati_emails = []
    #     try:
    #         with open(file_path, 'r', encoding='utf-8') as f:
    #             contenuto = f.read()

    #         blocchi = re.split(r'\n{2,}', contenuto.strip())

    #         email_pattern = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', re.IGNORECASE)
    #         phone_pattern = re.compile(r'\+?[0-9][0-9 ()\-\.]{5,}[0-9]')

    #         COUNTRY_CODES = [
    #             '1242','1246','1264','1268','1284','1340','1345','1441','1649','1664','1721','1758',
    #             '1767','1784','1787','1809','1829','1849','380','385','386','234','91','39','44','33',
    #             '34','49','36','30','31','32','52','54','55','56','57','58','60','61','62','63','64',
    #             '65','66','7','1','20','27','86','81','82','84','90','92','98'
    #         ]
    #         COUNTRY_CODES = sorted(set(COUNTRY_CODES), key=lambda x: -len(x))

    #         def format_phone(phone_raw):
    #             if not phone_raw or not phone_raw.strip():
    #                 return ''
    #             s = phone_raw.strip()
    #             digits = re.sub(r'[^0-9]', '', s)

    #             if s.startswith('+') and digits:
    #                 matched = None
    #                 for cc in COUNTRY_CODES:
    #                     if digits.startswith(cc):
    #                         rest = digits[len(cc):]
    #                         if len(rest) >= 6:
    #                             matched = cc
    #                             break
    #                 if not matched:
    #                     m_space = re.match(r'^\+([^\s]+)\s', s)
    #                     if m_space:
    #                         cc_candidate = re.sub(r'[^0-9]', '', m_space.group(1))
    #                         if cc_candidate and digits.startswith(cc_candidate) and len(digits[len(cc_candidate):]) >= 6:
    #                             matched = cc_candidate
    #                 if not matched:
    #                     if len(digits) > 2 and len(digits[2:]) >= 6:
    #                         matched = digits[:2]
    #                     elif len(digits) > 1 and len(digits[1:]) >= 6:
    #                         matched = digits[:1]
    #                     else:
    #                         matched = ''
    #                 if matched:
    #                     rest = digits[len(matched):]
    #                     return f'+{matched} {rest}'
    #                 else:
    #                     return '+' + digits
    #             else:
    #                 return digits

    #         # keywords utili a riconoscere occupazioni
    #         OCCUPATION_KEYWORDS = {
    #             'data','scientist','developer','engineer','ingegnere','analyst','analista',
    #             'manager','director','cto','ceo','founder','owner','product','designer','marketing',
    #             'sales','consultant','consulente','responsabile','head','research','ricerca',
    #             'operations','operation','specialist','teacher','professor','prof','doctor','dr',
    #             'student','studente','amministratore','admin','support','supporto','dev','architect',
    #             'impiegato','impiegata','operaio','operaia','libero','professionista'
    #         }

    #         def token_is_occupation(tok):
    #             if not tok:
    #                 return False
    #             t = re.sub(r'[^A-Za-zÀ-ÿ0-9]', '', tok).lower()
    #             if not t:
    #                 return False
    #             if t in OCCUPATION_KEYWORDS:
    #                 return True
    #             for kw in OCCUPATION_KEYWORDS:
    #                 if kw in t:
    #                     return True
    #             return False

    #         # ----------------- regex robusta ---------------
    #         # include: SISTEMA DI INVIO, MECCANISMO TRAMITE UNICO, DIGITURBO/PIGITURBO, parole *turbo*,
    #         # NEW ..., MIDDLE/SHORT, FACEBOOK, INSTAGRAM, TIKTOK, WHATSAPP, SITO WEB, ONLINE, EMAIL_INTERNAL, ecc.
    #         CHANNEL_REMOVE_RE = re.compile(
    #             r'\b(?:'
    #             r'SISTEMA(?:\s+DI)?\s+INVIO|'          # SISTEMA DI INVIO / SISTEMA INVIO
    #             r'MECCANISMO\s+TRAMITE\s+UNICO|'      # MECCANISMO TRAMITE UNICO
    #             r'(?:DI)?DIGITURBO(?:\s+X)?|'         # DIGITURBO / (possibile prefisso)
    #             r'PIGITURBO(?:\s+X)?|'                # PIGITURBO (typo variante)
    #             r'\w*turbo\b|'                        # catch-all parole che finiscono con 'turbo'
    #             r'NEW(?:\s+[A-ZÀ-Ö0-9][A-ZÀ-Ö0-9\w-]+)?|'  # NEW ... (seguìto da parola)
    #             r'MIDDLE(?:\s+SHORT(?:\s+[A-ZÀ-Ö0-9\w-]+)?)?|' # MIDDLE / MIDDLE SHORT / MIDDLE SHORT GIAN
    #             r'SHORT|'                             
    #             r'DRAGOTA|'                            # esempio presente negli screencap
    #             r'FACEBOOK|INSTAGRAM|TIKTOK|WHATSAPP|'
    #             r'SITO\s+WEB|SITO|ONLINE|EMAIL(?:_INTERNAL)?|FACEBOOK'
    #             r')\b.*',
    #             re.IGNORECASE | re.UNICODE
    #         )

    #         def strip_channels(s: str) -> str:
    #             # rimuove dalla prima occorrenza di un marker fino a fine riga
    #             if not s:
    #                 return ''
    #             s2 = CHANNEL_REMOVE_RE.sub(' ', s)
    #             s2 = re.sub(r'\s{2,}', ' ', s2).strip()
    #             return s2

    #         # -----------------------------------------------------------------------

    #         for blocco in blocchi:
    #             if not blocco.strip():
    #                 continue
    #             righe = [line.strip() for line in blocco.split('\n') if line.strip()]
    #             text = ' '.join(righe)

    #             email_match = email_pattern.search(text)
    #             email = email_match.group(0).strip() if email_match else ''

    #             phone_match = phone_pattern.search(text)
    #             telefono_raw = phone_match.group(0).strip() if phone_match else ''
    #             telefono_formatted = format_phone(telefono_raw)

    #             # prepare text_for_age: remove phone/email and channel tails
    #             text_for_age = text
    #             if telefono_raw:
    #                 text_for_age = re.sub(re.escape(telefono_raw), ' ', text_for_age)
    #                 text_for_age = re.sub(r'\+\d[\d\s\-\.\(\)]{4,}\d', ' ', text_for_age)
    #             if email:
    #                 text_for_age = re.sub(re.escape(email), ' ', text_for_age)
    #             text_for_age = strip_channels(text_for_age)

    #             # find age
    #             age_match = re.search(r'\b([1-9][0-9]{0,2})\b', text_for_age)
    #             age = ''
    #             start_age = end_age = None
    #             if age_match:
    #                 candidate = int(age_match.group(1))
    #                 if 10 <= candidate <= 120:
    #                     age = str(candidate)
    #                     start_age = age_match.start(1)
    #                     end_age = age_match.end(1)

    #             # first line cleaned for parsing name/occupation
    #             prima_riga = righe[0] if righe else ''
    #             prima_riga = strip_channels(prima_riga)

    #             fullname = ''
    #             occupation = ''

    #             if age:
    #                 # split around age
    #                 m_in_first = re.search(r'\b' + re.escape(age) + r'\b', prima_riga)
    #                 if m_in_first:
    #                     before = prima_riga[:m_in_first.start()].strip()
    #                     after = prima_riga[m_in_first.end():].strip()
    #                 else:
    #                     before = text_for_age[:start_age].strip()
    #                     after = text_for_age[end_age:].strip()
    #                 fullname = before
    #                 occupation = strip_channels(after)
    #             else:
    #                 # heuristics when age missing
    #                 prima_riga_clean = re.sub(r'\s*-\s*', ' - ', prima_riga).strip()
    #                 tokens = [t for t in re.split(r'\s+', prima_riga_clean) if t != '']

    #                 occ_idx = None
    #                 for i, tok in enumerate(tokens):
    #                     if token_is_occupation(tok):
    #                         occ_idx = i
    #                         break

    #                 if occ_idx is not None and occ_idx >= 1:
    #                     fullname = ' '.join(tokens[:occ_idx]).strip()
    #                     occupation = ' '.join(tokens[occ_idx:]).strip()
    #                 else:
    #                     if len(tokens) <= 2:
    #                         fullname = ' '.join(tokens).strip()
    #                         occupation = ''
    #                     else:
    #                         if '-' in tokens:
    #                             dash_idx = tokens.index('-')
    #                             if dash_idx >= 1:
    #                                 fullname = ' '.join(tokens[:dash_idx]).strip()
    #                                 occupation = ' '.join(tokens[dash_idx+1:]).strip()
    #                             else:
    #                                 fullname = ' '.join(tokens[:2]).strip()
    #                                 occupation = ' '.join(tokens[2:]).strip()
    #                         else:
    #                             # fallback: primi due token nome+cognome, resto occupazione
    #                             fullname = ' '.join(tokens[:2]).strip()
    #                             occupation = ' '.join(tokens[2:]).strip()

    #                 occupation = strip_channels(occupation)

    #             # split fullname into nome / cognome
    #             nome, cognome = '', ''
    #             if fullname:
    #                 parts = fullname.split()
    #                 if len(parts) == 1:
    #                     nome = parts[0]
    #                 else:
    #                     nome = parts[0]
    #                     cognome = ' '.join(parts[1:])

    #             # fallback: email/telefono da righe successive
    #             if not email and len(righe) > 1 and '@' in righe[1]:
    #                 email = righe[1].strip()
    #             if (not telefono_formatted) and len(righe) > 2:
    #                 telefono_formatted = format_phone(righe[2])

    #             if telefono_formatted and (len(re.sub(r'[^0-9]', '', telefono_formatted)) < 6):
    #                 self._log_message(f"Telefono sospetto trovato: {telefono_formatted} nel blocco: {prima_riga}")

    #             dati = [nome or '', cognome or '', age or '', occupation or '', email or '', telefono_formatted or '']
    #             dati_emails.append(dati)

    #         return dati_emails

    #     except FileNotFoundError:
    #         self._log_message(f"Errore: File '{file_path}' non trovato.")
    #         return []
    #     except Exception as e:
    #         self._log_message(f"Errore durante la lettura del file: {e}")
    #         return []



    def _clean_channel_and_dates(s: str) -> str:
        # rimuove porzioni come "SISTEMA INVIO ONLINE - FACEBOOK ..." fino a prima di una mail o numero
        # strategy: truncate everything from a channel marker until the end of the segment (but stop if email/phone remains)
        # Prima estraiamo email/phone altrove; quindi qui puliamo in modo conservativo.
        s = re.sub(DATE_RE, ' ', s)
        # rimuovi ocurrencees dei marker e tutto ciò che li segue se non c'è una mail dopo nel token
        s = re.sub(r'(?i)\bSISTEMA INVIO ONLINE\b.*', ' ', s)
        s = re.sub(r'(?i)\bSITO WEB\b.*', ' ', s)
        s = re.sub(r'(?i)\bFACEBOOK\b.*', ' ', s)
        s = re.sub(r'(?i)\bINSTAGRAM\b.*', ' ', s)
        s = re.sub(r'(?i)\bMAIL DOMANDE\b.*', ' ', s)
        return s

    def _insert_space_between_digits_and_letters(s: str) -> str:
        # trasforma "20Dipendente" -> "20 Dipendente" e "Operaio40" -> "Operaio 40"
        s = re.sub(r'(?<=\d)(?=[A-Za-zÀ-ÖØ-öø-ÿ])', ' ', s)
        s = re.sub(r'(?<=[A-Za-zÀ-ÖØ-öø-ÿ])(?=\d)', ' ', s)
        return s

    def choose_best_email(candidates):
        if not candidates:
            return None
        # preferisco email con lettere nel local-part e più lunga
        def score(e):
            local = e.split('@',1)[0]
            letters = sum(1 for c in local if c.isalpha())
            return (letters, len(e))
        return max(candidates, key=score)

    # def parse_record_text(block_text: str):
    #     """
    #     Input: un blocco di testo che rappresenta un record (più linee unite).
    #     Output: dict {name, surname, age, occupation, email, phone}
    #     """
    #     original = block_text.strip()
    #     text = original

    #     # 1) rimuovi prefisso numerico "6-" o "10-" all'inizio
    #     text = re.sub(LEADING_PREFIX_RE, '', text)

    #     # 2) estrai email (PRIMA)
    #     emails = EMAIL_RE.findall(text)
    #     email = choose_best_email(emails)
    #     if email:
    #         # rimuovo solo esatta occorrenza dell'email per non rompere altre parti
    #         text = re.sub(re.escape(email), ' ', text, count=1, flags=re.I)

    #     # 3) estrai telefono (PRIMA)
    #     phones = PHONE_RE.findall(text)
    #     phone = phones[0].strip() if phones else ''
    #     if phone:
    #         text = re.sub(re.escape(phone), ' ', text, count=1)

    #     # 4) rimuovi date e grandi marcatori (conservativo)
    #     text = _clean_channel_and_dates(text)

    #     # 5) ins. spazi fra numeri e lettere per gestire "18studente"
    #     text = _insert_space_between_digits_and_letters(text)

    #     # 6) normalizza spazi e rimuovi caratteri di punteggiatura inutili (ma non @ .)
    #     text = re.sub(r'[\"“”\'\t]', ' ', text)
    #     text = re.sub(r'[,:;]+', ' ', text)
    #     text = re.sub(r'\s{2,}', ' ', text).strip()

    #     # 7) tokenizza
    #     tokens = text.split()
    #     if not tokens:
    #         return {'name':'','surname':'','age':'','occupation':'','email':email or '','phone':phone or ''}

    #     # 8) trova age (primo token che è un numero plausibile 10..120)
    #     age_idx = None
    #     for i,tok in enumerate(tokens):
    #         if tok.isdigit():
    #             val = int(tok)
    #             if 10 <= val <= 120:
    #                 age_idx = i
    #                 break

    #     name = ''
    #     surname = ''
    #     age = ''
    #     occupation = ''

    #     if age_idx is not None:
    #         age = tokens[age_idx]
    #         name_tokens = tokens[:age_idx]
    #         occ_tokens = tokens[age_idx+1:]
    #         # name + surname: name = primo token, surname = resto dei name_tokens
    #         if name_tokens:
    #             name = name_tokens[0]
    #             surname = ' '.join(name_tokens[1:]) if len(name_tokens) > 1 else ''
    #         occupation = ' '.join(occ_tokens).strip()
    #     else:
    #         # NO AGE: heuristica
    #         # - il primo token è il nome
    #         # - successivi token capitalizzati o particles fanno parte del cognome
    #         # - il primo token che appare "basso" (inizia lowercase) e non è particle è probabile inizio occupazione
    #         name = tokens[0]
    #         if len(tokens) == 1:
    #             surname = ''
    #             occupation = ''
    #         else:
    #             # costruisco surname dalla posizione 1 in avanti finché token sembra cognome
    #             surname_parts = []
    #             occ_start = None
    #             for i in range(1, len(tokens)):
    #                 tok = tokens[i]
    #                 lower = tok.lower()
    #                 # se token è particle o è capitalizzato (inizia con maiuscola), consideralo parte del cognome
    #                 if lower in SURNAME_PARTICLES or tok[0].isupper():
    #                     surname_parts.append(tok)
    #                     continue
    #                 # se token è tutto maiuscolo o ha caratteri numerici stray -> probabilmente noise -> ignora
    #                 # se token in channel markers -> stop
    #                 if CHANNEL_LINE_RE.search(tok):
    #                     occ_start = i
    #                     break
    #                 # se token è minuscolo e non è una particle -> probabile inizio occ
    #                 if tok[0].islower() and lower not in SURNAME_PARTICLES:
    #                     occ_start = i
    #                     break
    #                 # fallback: if token looks like a typical occupation word (manager, operaio, infermiere...) check
    #                 # but non hard-code too much: if following token is lowercase, assume occupation
    #                 if i+1 < len(tokens) and tokens[i+1][0].islower():
    #                     occ_start = i
    #                     break
    #                 # else, assume it's surname (capitalized)
    #                 surname_parts.append(tok)

    #             if surname_parts:
    #                 surname = ' '.join(surname_parts).strip()
    #                 if occ_start is None:
    #                     # anything after surname_parts is occupation
    #                     rest_idx = 1 + len(surname_parts)
    #                     occupation = ' '.join(tokens[rest_idx:]).strip()
    #                 else:
    #                     occupation = ' '.join(tokens[occ_start:]).strip()
    #             else:
    #                 # no surname parts decided: everything after first token is either surname or occupation
    #                 # fallback: put next token into surname, rest into occupation
    #                 if len(tokens) >= 2:
    #                     surname = tokens[1]
    #                     occupation = ' '.join(tokens[2:]).strip()
    #                 else:
    #                     surname = ''
    #                     occupation = ''

    #     # pulizie finali: strip ed eventuali doppie spazi
    #     for k in ('name','surname','age','occupation','email','phone'):
    #         if k in locals():
    #             pass
    #     parsed = {
    #         'name': (name or '').strip(),
    #         'surname': (surname or '').strip(),
    #         'age': (age or '').strip(),
    #         'occupation': (occupation or '').strip(),
    #         'email': (email or '').strip(),
    #         'phone': (phone or '').strip()
    #     }
    #     return parsed

    # def leggi_file_dati(self, file_path):
    #     """
    #     Funzione che legge il file .txt e ritorna lista di dict (uno per persona).
    #     Strategie:
    #     - suddivide il file in blocchi separati da blank-line (piu affidabile coi tuoi sample)
    #     - per ogni blocco usa parse_record_text
    #     """
    #     results = []
    #     try:
    #         with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
    #             content = f.read()
    #     except Exception as e:
    #         self._log_message(f"Errore apertura file: {e}")
    #         return results

    #     # Normalizza separatori: usa doppio newline come separatore record; fallback: ogni riga che contiene phone/email/numero va considerata record.
    #     blocks = re.split(r'\n\s*\n', content.strip())
    #     for b in blocks:
    #         # se blocco multilinea, uniscilo in una singola linea per parsing
    #         collapsed = ' '.join(line.strip() for line in b.splitlines() if line.strip())
    #         if not collapsed:
    #             continue
    #         # se il blocco è troppo corto e la riga successiva contiene telefono/email, potremmo voler unire con next;
    #         parsed = parse_record_text(collapsed)
    #         # validità minima: almeno name o email o phone
    #         if parsed['name'] or parsed['email'] or parsed['phone']:
    #             results.append(parsed)
    #     return results

    def _pulizia_avanzata_testo(self, testo, email, telefono):
        """Pulizia aggressiva del testo per rimuovere pattern indesiderati."""
        if not testo:
            return ""
        
        # 1. Rimuovi email e telefono
        if email:
            testo = re.sub(re.escape(email), '', testo, flags=re.IGNORECASE)
        if telefono:
            testo = re.sub(re.escape(telefono), '', testo)
        
        # 2. Rimuovi punti isolati che rompono il parsing
        testo = self._rimuovi_punti_isolati(testo)
        
        # 3. Rimuovi date in tutti i formati (gg/mm/aaaa, gg-mm-aa, etc.)
        testo = re.sub(r'\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b', '', testo)
        
        # 4. Rimuovi marcatori di canale e system info
        pattern_canali = [
            r'SISTEMA\s+INVIO\s+ONLINE[^\.]*\.?\s*',
            r'SITO\s+WEB[^\.]*\.?\s*', 
            r'FACEBOOK[^\.]*\.?\s*',
            r'INSTAGRAM[^\.]*\.?\s*',
            r'MAIL\s+DOMANDE[^\.]*\.?\s*',
            r'INVI[^\.]*\.?\s*',
            r'GIAN[^\.]*\.?\s*',
            r'JAN[^\.]*\.?\s*',
            r'INVI\s+O[^\.]*\.?\s*',
            r'MIDDLE[^\.]*\.?\s*',
            r'SHORT[^\.]*\.?\s*',
            r'NEW[^\.]*\.?\s*',
            r'DRAGOTA[^\.]*\.?\s*',
            r'TIKTOK[^\.]*\.?\s*',
            r'WHATSAPP[^\.]*\.?\s*',
            r'ONLINE[^\.]*\.?\s*',
            r'EMAIL_INTERNAL[^\.]*\.?\s*',
            r'DIGITURBO[^\.]*\.?\s*',
            r'PIGITURBO[^\.]*\.?\s*',
            r'MECCANISMO\s+TRAMITE\s+UNICO[^\.]*\.?\s*',
            r'INVIO[^\.]*\.?\s*',
            r'SISTEMA[^\.]*\.?\s*'
        ]
        
        for pattern in pattern_canali:
            testo = re.sub(pattern, '', testo, flags=re.IGNORECASE)
        
        # 5. Rimuovi numeri residui (10, 9, etc.) che non sono età
        testo = re.sub(r'\b(10|09|9|11|12)(?![0-9])\b', '', testo)
        
        # 6. Rimuovi prefissi numerici (es: "6-", "10-")
        testo = re.sub(r'^\s*\d+\s*-\s*', '', testo)
        testo = re.sub(r'\s+\d+\s*-\s*', ' ', testo)
        
        # 7. Separa età da occupazione quando sono attaccati
        testo = self._separa_eta_da_occupazione(testo)
        
        # 8. Corregge la formattazione dei nomi
        testo = self._correggi_formattazione_nomi(testo)
        
        # 9. Normalizza spazi tra numeri e lettere
        testo = self._normalizza_spazi_num_lettere(testo)
        
        # 10. Normalizza spazi
        testo = re.sub(r'\s+', ' ', testo).strip()
        
        return testo
    
    def _normalizza_spazi_num_lettere(self, testo):
        """Inserisce spazi tra numeri e lettere dove mancano."""
        if not testo:
            return testo
        
        # Separa numeri da lettere (es. "58Libero" -> "58 Libero", "20Studente" -> "20 Studente")
        testo = re.sub(r'(\b\d{2,3})([A-Za-zÀ-ÿ])', r'\1 \2', testo)
        
        # Separa lettere da numeri (es. "pendente2005" -> "pendente 2005")
        testo = re.sub(r'([A-Za-zÀ-ÿ])(\d{2,})', r'\1 \2', testo)
        
        # Separa numeri da lettere per singoli caratteri (es. "Lavoro3" -> "Lavoro 3")
        testo = re.sub(r'([A-Za-zÀ-ÿ])(\d)', r'\1 \2', testo)
        testo = re.sub(r'(\d)([A-Za-zÀ-ÿ])', r'\1 \2', testo)
        
        return testo
    
    def _valida_e_corregge_campi(self, nome, cognome, eta, occupazione, email, telefono):
        """Validazione e correzione finale dei campi estratti."""

        # Pulizia base
        if eta and eta.isdigit():
            age_val = int(eta)
            if not (15 <= age_val <= 120):
                eta = ""

        if nome and nome.startswith('-'):
            nome = nome[1:].strip()

        if cognome and any(char.isdigit() for char in cognome):
            cognome = re.sub(r'\s*\d+$', '', cognome)
            cognome = re.sub(r'^\d+\s*', '', cognome)

        
        nome_ok = bool(nome.strip())
        cognome_ok = bool(cognome.strip())
        altri_dati = any([
            email.strip(),
            telefono.strip(),
            eta.strip(),
            occupazione.strip()
        ])

        # Se manca tutto tranne nome o cognome singolo -> SCARTA
        if (nome_ok != cognome_ok) and not altri_dati:
            # ritorna tutto vuoto per segnalare che questo record è da ignorare
            return "", "", "", "", "", ""

        # Se non c’è neanche nome né cognome, scarta comunque
        if not nome_ok and not cognome_ok:
            return "", "", "", "", "", ""

        return nome, cognome, eta, occupazione, email, telefono

    def _estrai_campi_strutturati(self, testo):
        """Estrazione robusta di nome, cognome, età e occupazione."""
        if not testo:
            return "", "", "", ""
        
        # Inizializza campi
        nome, cognome, eta, occupazione = "", "", "", ""
        
        # Tokenizza il testo
        tokens = testo.split()
        if not tokens:
            return nome, cognome, eta, occupazione

        # CERCA ETÀ PRIMA - strategia più aggressiva
        found_age_index = -1
        age_candidates = []
        
        for i, token in enumerate(tokens):
            # Cerca numeri puri come età
            if token.isdigit():
                age_val = int(token)
                if 15 <= age_val <= 120:  # Range più realistico per età
                    age_candidates.append((i, token, age_val))
            
            # Cerca pattern come "33." già puliti o "58Libero" già separati
            else:
                # Se il token è un numero seguito da lettere (es: "33Segretaria" ma dovrebbe essere già separato)
                match = re.match(r'^(\d{2,3})([A-Za-zÀ-ÿ].*)$', token)
                if match:
                    age_val = int(match.group(1))
                    if 15 <= age_val <= 120:
                        # Separa il token e aggiorna la lista dei tokens
                        tokens[i] = match.group(1)
                        tokens.insert(i + 1, match.group(2))
                        age_candidates.append((i, match.group(1), age_val))
                        break

        # Seleziona il candidato età più probabile (il primo che incontriamo)
        if age_candidates:
            found_age_index, eta, age_val = age_candidates[0]
        
        # Se abbiamo trovato un'età
        if found_age_index >= 0:
            # Tutto prima dell'età è nome+cognome
            if found_age_index > 0:
                parte_nome_cognome = tokens[:found_age_index]
                if parte_nome_cognome:
                    nome = parte_nome_cognome[0]
                    cognome = ' '.join(parte_nome_cognome[1:]) if len(parte_nome_cognome) > 1 else ""
            
            # Tutto dopo l'età è occupazione
            if found_age_index + 1 < len(tokens):
                occupazione = ' '.join(tokens[found_age_index+1:])
        else:
            # STRATEGIA DI FALLBACK: cerca pattern comuni senza età esplicita
            if len(tokens) >= 2:
                nome = tokens[0]
                
                # Se il secondo token sembra un cognome (inizia con maiuscola) e non è un numero
                if len(tokens) > 1 and tokens[1][0].isupper() and not tokens[1].isdigit():
                    cognome = tokens[1]
                    occupazione = ' '.join(tokens[2:]) if len(tokens) > 2 else ""
                else:
                    # Altrimenti, tutto il resto è occupazione
                    occupazione = ' '.join(tokens[1:]) if len(tokens) > 1 else ""
            else:
                nome = tokens[0] if tokens else ""
        
        return nome, cognome, eta, occupazione

    def _pulizia_campo(self, campo):
        """Pulizia finale di un singolo campo."""
        if not campo:
            return ""
        
        # Rimuovi trattini all'inizio e spazi extra
        campo = re.sub(r'^[-\s]+', '', campo)
        campo = re.sub(r'[-\s]+$', '', campo)
        
        # Rimuovi numeri isolati all'inizio/fine
        campo = re.sub(r'^\d+\s*', '', campo)
        campo = re.sub(r'\s*\d+$', '', campo)
        
        # Rimuovi caratteri speciali problematici ma mantieni trattini interni
        campo = re.sub(r'[^\w\sÀ-ÿ\-]', '', campo)
        
        # Normalizza spazi
        campo = re.sub(r'\s+', ' ', campo).strip()
        
        return campo

    def _pulizia_email(self, email):
        """Pulizia specifica per email."""
        if not email:
            return ""
        
        # Rimuovi numeri residui dopo l'email
        email = re.sub(r'(\.[0-9]{2,})+$', '', email)
        email = re.sub(r'\s+[0-9]+\s*$', '', email)
        
        return email.strip()

    def _pulizia_telefono(self, telefono):
        """Pulizia specifica per telefono."""
        if not telefono:
            return ""
        
        # Rimuovi numeri residui dopo il telefono
        telefono = re.sub(r'\s+[0-9]+\s*$', '', telefono)
        telefono = re.sub(r'\.\d+$', '', telefono)
        
        return telefono.strip()

    def leggi_file_dati(self, file_path):
        """
        Legge un file di testo e restituisce lista di righe parsate correttamente.
        Versione migliorata con regex più robuste.
        """
        dati_emails = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                contenuto = f.read()

            # Divide in blocchi separati da righe vuote
            blocchi = re.split(r'\n{2,}', contenuto.strip())
            
            for blocco in blocchi:
                if not blocco.strip():
                    continue
                    
                righe = [line.strip() for line in blocco.split('\n') if line.strip()]
                if not righe:
                    continue
                    
                # Unisci le righe per il parsing
                testo_completo = ' '.join(righe)
                
                # FASE 1: Estrazione pulita di email e telefono
                email_match = EMAIL_RE.search(testo_completo)
                email = email_match.group(0) if email_match else ''
                
                phone_match = PHONE_RE.search(testo_completo)
                telefono = phone_match.group(0) if phone_match else ''
                
                # FASE 2: Pulizia aggressiva del testo
                testo_pulito = self._pulizia_avanzata_testo(testo_completo, email, telefono)
                
                # FASE 3: Estrazione strutturata dei campi
                nome, cognome, eta, occupazione = self._estrai_campi_strutturati(testo_pulito)
                
                # FASE 4: Validazione e correzione finale
                nome, cognome, eta, occupazione, email, telefono = self._valida_e_corregge_campi(
                    nome, cognome, eta, occupazione, email, telefono
                )
                
                # FASE 5: Pulizia finale dei campi
                nome = self._pulizia_campo(nome)
                cognome = self._pulizia_campo(cognome)
                occupazione = self._pulizia_campo(occupazione)
                email = self._pulizia_email(email)
                telefono = self._pulizia_telefono(telefono)
                
                dati = [nome, cognome, eta, occupazione, email, telefono]
                # se TUTTI i campi sono vuoti -> scarta
                if not any(field and str(field).strip() for field in dati):
                    continue
                dati_emails.append(dati)
                
            return dati_emails

        except Exception as e:
            self._log_message(f"Errore durante la lettura del file: {e}")
            return []

    def leggi_sheet_sorgente(self, spreadsheet_name: str, worksheet_name: str) -> list:
        """Legge righe da Google Sheets e restituisce lista di stringhe."""
        gc = self.connect_to_sheets()
        if not gc:
            return []

        try:
            sh = gc.open(spreadsheet_name)
            ws = sh.worksheet(worksheet_name)
            all_values = ws.get_all_values()
            
            righe = []
            for row in all_values:
                # Unisci le celle non vuote
                testo = ' '.join([str(cell).strip() for cell in row if str(cell).strip()])
                if testo:
                    righe.append(testo)
                    
            self._log_message(f"Trovate {len(righe)} righe nel foglio sorgente")
            return righe
            
        except Exception as e:
            self._log_message(f"Errore lettura sheet sorgente: {e}")
            return []

    # def _scrivi_dati_con_chunking(self, dati_completi):
    #     """Scrive i dati in Google Sheets usando chunking."""
    #     gc = self.connect_to_sheets()
    #     if not gc:
    #         messagebox.showerror("Errore", "Impossibile connettersi a Google Sheets.")
    #         return

    #     spreadsheet_name = self.spreadsheet_name_var.get().strip()
    #     worksheet_name = self.worksheet_name_var.get().strip()

    #     if not spreadsheet_name:
    #         messagebox.showerror("Errore", "Inserisci il nome del foglio di calcolo.")
    #         return

    #     try:
    #         # Apri o crea spreadsheet
    #         try:
    #             sh = gc.open(spreadsheet_name)
    #         except SpreadsheetNotFound:
    #             sh = gc.create(spreadsheet_name)
    #             self._log_message(f"Creato nuovo spreadsheet: {spreadsheet_name}")

    #         # Apri o crea worksheet
    #         try:
    #             ws = sh.worksheet(worksheet_name)
    #             # Chiedi conferma prima di sovrascrivere
    #             if not messagebox.askyesno("Conferma", f"Il worksheet '{worksheet_name}' esiste già. Sovrascrivere?"):
    #                 return
    #             ws.clear()
    #         except WorksheetNotFound:
    #             ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=7)

    #         # Intestazioni con checkbox
    #         headers = ["✅", "Nome", "Cognome", "Età", "Occupazione", "Email", "Numero di telefono"]
    #         ws.update('A1:G1', [headers])

    #         # Scrivi i dati in chunk
    #         WRITE_CHUNK_SIZE = 2000
    #         for i in range(0, len(dati_completi), WRITE_CHUNK_SIZE):
    #             chunk_dati = dati_completi[i:i + WRITE_CHUNK_SIZE]
    #             start_row = i + 2  # +2 perché la riga 1 sono le headers
                
    #             # Prepara i dati per la scrittura, aggiungendo la colonna checkbox (False)
    #             valori = []
    #             for riga in chunk_dati:
    #                 # riga è [nome, cognome, eta, occupazione, email, telefono]
    #                 # lo trasformiamo in [False, nome, cognome, eta, occupazione, email, telefono]
    #                 valori.append([False] + riga)
                
    #             range_start = f'A{start_row}'
    #             ws.update(range_start, valori)
    #             self._log_message(f"Scritto chunk {i//WRITE_CHUNK_SIZE + 1} (righe {start_row}-{start_row + len(chunk_dati) - 1})")

    #         self._log_message(f"Scrittura completata: {len(dati_completi)} record")
            
    #         # Applica formattazione (checkbox e stile)
    #         try:
    #             self.apply_sheet_styling(sh, ws)
    #         except Exception as e:
    #             self._log_message(f"Avviso: impossibile applicare la formattazione: {e}")

    #     except Exception as e:
    #         self._log_message(f"Errore durante la scrittura: {e}")
    #         raise

    def _scrivi_dati_con_chunking(self, dati_completi, dest_spreadsheet, dest_worksheet):
        """Scrive i dati in Google Sheets DESTINAZIONE usando chunking."""
        gc = self.connect_to_sheets()
        if not gc:
            messagebox.showerror("Errore", "Impossibile connettersi a Google Sheets.")
            return

        # Usa i parametri espliciti per la destinazione
        spreadsheet_name = dest_spreadsheet
        worksheet_name = dest_worksheet

        if not spreadsheet_name:
            messagebox.showerror("Errore", "Nome del foglio di calcolo di destinazione mancante.")
            return

        try:
            # Apri o crea spreadsheet DESTINAZIONE
            try:
                sh = gc.open(spreadsheet_name)
                self._log_message(f"Aperto spreadsheet destinazione: {spreadsheet_name}")
            except SpreadsheetNotFound:
                sh = gc.create(spreadsheet_name)
                self._log_message(f"Creato nuovo spreadsheet destinazione: {spreadsheet_name}")

            # Gestisci worksheet DESTINAZIONE
            if not worksheet_name:
                worksheet_name = datetime.now().strftime("%B %Y")
                self._log_message(f"Usando nome worksheet predefinito: {worksheet_name}")

            try:
                ws = sh.worksheet(worksheet_name)
                self._log_message(f"Worksheet destinazione '{worksheet_name}' esistente - richiesta conferma sovrascrittura")
                if not messagebox.askyesno("Conferma", f"Il worksheet '{worksheet_name}' esiste già in '{spreadsheet_name}'. Sovrascrivere tutto il contenuto?"):
                    self._log_message("Operazione annullata dall'utente.")
                    return
                ws.clear()
                self._log_message(f"Worksheet '{worksheet_name}' pulito")
            except WorksheetNotFound:
                ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=7)
                self._log_message(f"Creato nuovo worksheet destinazione: {worksheet_name}")

            # Intestazioni con checkbox
            headers = ["✅", "Nome", "Cognome", "Età", "Occupazione", "Email", "Telefono"]
            ws.update('A1:G1', [headers])
            self._log_message("Intestazioni scritte")

            # Scrivi i dati in chunk
            WRITE_CHUNK_SIZE = 2000
            total_written = 0
            
            for i in range(0, len(dati_completi), WRITE_CHUNK_SIZE):
                chunk_dati = dati_completi[i:i + WRITE_CHUNK_SIZE]
                start_row = i + 2  # +2 perché la riga 1 sono le headers
                
                # Prepara i dati per la scrittura, aggiungendo la colonna checkbox (False)
                valori = []
                for riga in chunk_dati:
                    # riga è [nome, cognome, eta, occupazione, email, telefono]
                    # trasformiamo in [False, nome, cognome, eta, occupazione, email, telefono]
                    valori.append([False] + riga)
                
                range_start = f'A{start_row}'
                ws.update(range_start, valori)
                total_written += len(chunk_dati)
                self._log_message(f"Scritto chunk {i//WRITE_CHUNK_SIZE + 1}: righe {start_row}-{start_row + len(chunk_dati) - 1}")

            self._log_message(f"Scrittura completata: {total_written} record in '{spreadsheet_name}/{worksheet_name}'")
            
            # Applica formattazione (checkbox e stile)
            try:
                self.apply_sheet_styling(sh, ws)
                self._log_message("Formattazione applicata con successo")
            except Exception as e:
                self._log_message(f"Avviso: impossibile applicare la formattazione: {e}")

        except Exception as e:
            self._log_message(f"Errore durante la scrittura in destinazione: {e}")
            raise

    def _separa_eta_da_occupazione(self, testo):
        """Separa l'età dall'occupazione quando sono attaccati (es: '58Libero' -> '58 Libero')."""
        if not testo:
            return testo
        
        # Pattern per età seguita da occupazione (es: "58Libero", "20Studente")
        testo = re.sub(r'(\b\d{2})\s*([A-Za-zÀ-ÿ])', r'\1 \2', testo)
        
        # Pattern per occupazione seguita da età (meno comune, ma per completezza)
        testo = re.sub(r'([A-Za-zÀ-ÿ])\s*(\d{2}\b)', r'\1 \2', testo)
        
        return testo

    def _rimuovi_punti_isolati(self, testo):
        """Rimuove punti isolati che rompono il parsing (es: '33. Segretaria' -> '33 Segretaria')."""
        if not testo:
            return testo
        
        # Rimuove punti che separano età da occupazione
        testo = re.sub(r'(\b\d{1,3})\.\s+', r'\1 ', testo)
        
        # Rimuove punti alla fine di numeri
        testo = re.sub(r'(\b\d{1,3})\.', r'\1', testo)
        
        return testo

    def _correggi_formattazione_nomi(self, testo):
        """Corregge i formati dei nomi problematici."""
        if not testo:
            return testo
        
        # Rimuove trattini all'inizio dei nomi (es: "-Andrew" -> "Andrew")
        testo = re.sub(r'^\s*-\s*', '', testo)
        
        # Corregge spazi mancanti dopo i trattini nei nomi
        testo = re.sub(r'(\w)-\s*(\w)', r'\1 - \2', testo)
        
        return testo

    # def run_bot(self):
    #     """Logica principale del bot con chunking reintrodotto."""
    #     try:
    #         source_mode = self.source_mode.get()
    #         self._log_message(f"Modalità sorgente: {source_mode}")

    #         # Salva configurazione
    #         self._save_config_value('PATHS', 'dati_emails_path', self.file_path_var.get() or '')
    #         self._save_config_value('SHEETS', 'spreadsheet_name', self.spreadsheet_name_var.get() or '')
    #         self._save_config_value('SHEETS', 'worksheet_name', self.worksheet_name_var.get() or '')
    #         self._save_config_value('SETTINGS', 'source_mode', source_mode)

    #         if source_mode == 'file':
    #             file_path = self.file_path_var.get()
    #             if not file_path or not os.path.exists(file_path):
    #                 messagebox.showerror("Errore", "Seleziona un file valido.")
    #                 return
                
    #             self._log_message("Lettura e elaborazione del file dati...")
    #             dati_completi = self.leggi_file_dati(file_path)
                
    #         else:  # modalità sheets
    #             ss_name = self.spreadsheet_name_var.get().strip()
    #             ws_name = self.worksheet_name_var.get().strip()
    #             if not ss_name:
    #                 messagebox.showerror("Errore", "Inserisci il nome del Google Spreadsheet.")
    #                 return
                
    #             self._log_message(f"Lettura da Sheet: {ss_name} / {ws_name}")
    #             righe_sorgente = self.leggi_sheet_sorgente(ss_name, ws_name)
                
    #             # Processa i dati sorgente in chunk
    #             CHUNK_SIZE = 5000
    #             dati_completi = []
    #             for i in range(0, len(righe_sorgente), CHUNK_SIZE):
    #                 chunk = righe_sorgente[i:i + CHUNK_SIZE]
    #                 self._log_message(f"Processando chunk {i//CHUNK_SIZE + 1}...")
                    
    #                 # Crea file temporaneo per il parsing
    #                 with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', delete=False, suffix='.txt') as tmp:
    #                     tmp.write("\n\n".join(chunk))
    #                     tmp_path = tmp.name
                    
    #                 try:
    #                     dati_chunk = self.leggi_file_dati(tmp_path)
    #                     dati_completi.extend(dati_chunk)
    #                 finally:
    #                     os.unlink(tmp_path)

    #         if not dati_completi:
    #             messagebox.showinfo("Info", "Nessun dato valido trovato.")
    #             return

    #         self._log_message(f"Trovati {len(dati_completi)} record validi.")
            
    #         # Scrivi in Google Sheets con chunking
    #         self._scrivi_dati_con_chunking(dati_completi)

    #         # Svuota file se richiesto
    #         if self.svuota_file_var.get() and source_mode == 'file':
    #             self.svuota_file_dati()

    #         self._log_message("Processo completato con successo.")
    #         messagebox.showinfo("Successo", "Operazione completata.")

    #     except Exception as e:
    #         self._log_message(f"Errore critico: {e}")
    #         messagebox.showerror("Errore", f"Errore durante l'esecuzione: {e}")

    def _get_source_dest_info(self):
        """Restituisce informazioni chiare su sorgente e destinazione per logging."""
        source_mode = self.source_mode.get()
        
        if source_mode == 'file':
            source_info = f"File: {os.path.basename(self.file_path_var.get()) if self.file_path_var.get() else 'N/A'}"
        else:
            source_info = f"Sheet: {self.source_spreadsheet_var.get() or 'N/A'}/{self.source_worksheet_var.get() or 'N/A'}"
        
        dest_info = f"Sheet: {self.dest_spreadsheet_var.get() or 'N/A'}/{self.dest_worksheet_var.get() or 'N/A'}"
        
        return source_info, dest_info

    def run_bot(self):
        """Logica principale del bot con chiara separazione sorgente/destinazione."""

        source_info, dest_info = self._get_source_dest_info()
        self._log_message(f"Configurazione: SORGENTE={source_info}, DESTINAZIONE={dest_info}")

        try:
            source_mode = self.source_mode.get()
            self._log_message(f"Modalità sorgente: {source_mode}")

            # Salva configurazione
            self._save_config_value('PATHS', 'dati_emails_path', self.file_path_var.get() or '')
            self._save_config_value('SHEETS', 'source_spreadsheet', self.source_spreadsheet_var.get() or '')
            self._save_config_value('SHEETS', 'source_worksheet', self.source_worksheet_var.get() or '')
            self._save_config_value('SHEETS', 'dest_spreadsheet', self.dest_spreadsheet_var.get() or '')
            self._save_config_value('SHEETS', 'dest_worksheet', self.dest_worksheet_var.get() or '')
            self._save_config_value('SETTINGS', 'source_mode', source_mode)

            dati_completi = []
            
            try:
                self._persist_local_settings()
            except Exception:
                pass     
               
            if source_mode == 'file':
                # Modalità FILE: leggi da file locale
                file_path = self.file_path_var.get()
                if not file_path or not os.path.exists(file_path):
                    messagebox.showerror("Errore", "Seleziona un file valido.")
                    return
                
                self._log_message(f"Lettura da file: {os.path.basename(file_path)}")
                dati_completi = self.leggi_file_dati(file_path)
                
            else:
                # Modalità SHEETS: leggi da Google Sheets sorgente
                src_spreadsheet = self.source_spreadsheet_var.get().strip()
                src_worksheet = self.source_worksheet_var.get().strip()
                
                if not src_spreadsheet:
                    messagebox.showerror("Errore", "Inserisci il nome del Google Spreadsheet SORGENTE.")
                    return
                
                self._log_message(f"Lettura da Sheet SORGENTE: {src_spreadsheet} / {src_worksheet}")
                righe_sorgente = self.leggi_sheet_sorgente(src_spreadsheet, src_worksheet)
                
                if not righe_sorgente:
                    messagebox.showinfo("Info", "Nessun dato trovato nello sheet sorgente.")
                    return

                # Processa i dati sorgente in chunk
                CHUNK_SIZE = 5000
                for i in range(0, len(righe_sorgente), CHUNK_SIZE):
                    chunk = righe_sorgente[i:i + CHUNK_SIZE]
                    self._log_message(f"Processando chunk {i//CHUNK_SIZE + 1}...")
                    
                    # Crea file temporaneo per il parsing
                    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', delete=False, suffix='.txt') as tmp:
                        tmp.write("\n\n".join(chunk))
                        tmp_path = tmp.name
                    
                    try:
                        dati_chunk = self.leggi_file_dati(tmp_path)
                        dati_completi.extend(dati_chunk)
                        self._log_message(f"Chunk {i//CHUNK_SIZE + 1}: elaborati {len(dati_chunk)} record")
                    finally:
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass

            if not dati_completi:
                messagebox.showinfo("Info", "Nessun dato valido trovato.")
                return

            self._log_message(f"Trovati {len(dati_completi)} record validi totali.")
            
            # Prendi destinazione
            dest_spreadsheet = self.dest_spreadsheet_var.get().strip()
            dest_worksheet = self.dest_worksheet_var.get().strip()
            
            if not dest_spreadsheet:
                messagebox.showerror("Errore", "Inserisci il nome del Google Spreadsheet DESTINAZIONE.")
                return

            # Scrivi in Google Sheets DESTINAZIONE
            self._scrivi_dati_con_chunking(dati_completi, dest_spreadsheet, dest_worksheet)

            # Svuota file se richiesto (solo in modalità file)
            if self.svuota_file_var.get() and source_mode == 'file':
                self.svuota_file_dati()

            self._log_message("Processo completato con successo.")
            messagebox.showinfo("Successo", f"Operazione completata: {len(dati_completi)} record processati.")

        except Exception as e:
            self._log_message(f"Errore critico: {e}")
            messagebox.showerror("Errore", f"Errore durante l'esecuzione: {e}")

    def apply_sheet_styling(self, sh, ws):
        """
        Applica checkbox in colonna A e styling:
        - header bold + background
        - bordi su tabella A1:G{row_count}
        - data validation (checkbox) su A2:A{row_count}
        - conditional formatting: se A è TRUE -> strikethrough + sfondo grigio su A..G
        sh: gspread.Spreadsheet
        ws: gspread.Worksheet
        """
        try:
            import os
            from googleapiclient.discovery import build
            from google.oauth2 import service_account

            # --- trova credentials.json in percorsi probabili ---
            possible = []
            if hasattr(self, 'credentials_path') and self.credentials_path:
                possible.append(self.credentials_path)
            possible.append(os.path.join(os.getcwd(), 'credentials.json'))
            possible.append(os.path.join(os.path.dirname(__file__), 'credentials.json'))
            if hasattr(self, 'user_data_path') and self.user_data_path:
                possible.append(os.path.join(self.user_data_path, 'credentials.json'))

            credentials_path = next((p for p in possible if p and os.path.exists(p)), None)
            if not credentials_path:
                self._log_message("apply_sheet_styling: nessuna credentials.json trovata nei percorsi probabili; salto styling avanzato.")
                return

            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
            creds = service_account.Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
            service = build('sheets', 'v4', credentials=creds)

            sheet_id = ws._properties.get('sheetId')
            spreadsheet_id = sh.id

            # prendi numero righe attuali: valori totali
            values = ws.get_all_values()
            row_count = max(1, len(values))
            # A=checkbox, B..G = Nome,Cognome,Età,Occupazione,Email,Telefono -> total columns = 7 (A..G)
            col_count = 7

            requests = []

            # 1) header style A1:G1
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": col_count
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.94, "green": 0.94, "blue": 0.96},
                            "horizontalAlignment": "CENTER",
                            "textFormat": {"bold": True, "fontSize": 11}
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            })

            # 2) bordi su tutta la tabella A1:G{row_count}
            requests.append({
                "updateBorders": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": row_count,
                        "startColumnIndex": 0,
                        "endColumnIndex": col_count
                    },
                    "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "left": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "right": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "innerHorizontal": {"style": "SOLID", "width": 1, "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
                    "innerVertical": {"style": "SOLID", "width": 1, "color": {"red": 0.85, "green": 0.85, "blue": 0.85}}
                }
            })

            # 3) set dataValidation (checkbox) per A2:A{row_count} usando repeatCell
            if row_count > 1:
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 1,
                            "endRowIndex": row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        },
                        "cell": {
                            # userEnteredValue opzionale; dataValidation definisce l'interfaccia checkbox
                            "userEnteredValue": {"boolValue": False},
                            "dataValidation": {
                                "condition": {"type": "BOOLEAN"},
                                "showCustomUi": True,
                                "strict": False
                            }
                        },
                        "fields": "userEnteredValue,dataValidation"
                    }
                })

                # 4) conditional formatting rule: se colonna A TRUE -> strike + background grey (applies to A..G)
                requests.append({
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{
                                "sheetId": sheet_id,
                                "startRowIndex": 1,
                                "endRowIndex": row_count,
                                "startColumnIndex": 0,
                                "endColumnIndex": col_count
                            }],
                            "booleanRule": {
                                "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": "=$A2=TRUE"}]},
                                "format": {
                                    "backgroundColor": {"red": 0.92, "green": 0.92, "blue": 0.92},
                                    "textFormat": {"foregroundColor": {"red": 0.45, "green": 0.45, "blue": 0.45}, "strikethrough": True}
                                }
                            }
                        },
                        "index": 0
                    }
                })

            # INVIA BATCH (header style, bordi, dataValidation e conditional formatting)
            if requests:
                body = {"requests": requests}
                service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

            # 5) infine, metti i valori FALSE in colonna A (A2..A{row_count}) con values().update
            if row_count > 1:
                check_values = [["FALSE"] for _ in range(row_count - 1)]
                service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=f"A2:A{row_count}",
                    valueInputOption="USER_ENTERED",
                    body={"values": check_values}
                ).execute()

            self._log_message(f"Stile applicato al foglio '{ws.title}' (rows={row_count}).")
        except Exception as e:
            self._log_message(f"apply_sheet_styling: errore applicando stile: {e}")

    

    def _ensure_spreadsheet(self, gc, spreadsheet_name):
        """
        Restituisce l'oggetto Spreadsheet (gspread.Spreadsheet).
        Se non esiste lo crea.
        """
        try:
            sh = gc.open(spreadsheet_name)
            return sh
        except Exception as e:
            # Spreadsheet non trovato -> crealo
            try:
                sh = gc.create(spreadsheet_name)
                self._log_message(f"Creato nuovo Spreadsheet: {spreadsheet_name}")
                return sh
            except Exception as ce:
                self._log_message(f"Errore creating spreadsheet '{spreadsheet_name}': {ce}")
                raise

    def _get_worksheet_or_none(self, sh, ws_name):
        try:
            if not ws_name:
                return None
            return sh.worksheet(ws_name)
        except WorksheetNotFound:
            return None

    def _create_worksheet(self, sh, ws_name, rows=1000, cols=10):
        """
        Crea una worksheet nel Spreadsheet sh. Ritorna la worksheet.
        """
        ws = sh.add_worksheet(title=ws_name, rows=str(max(100, rows)), cols=str(max(6, cols)))
        self._log_message(f"Worksheet '{ws_name}' creata nello spreadsheet '{sh.title}'")
        return ws

    def _confirm_clear_or_create_ws(self, sh, ws_name, total_rows_estimate=0):
        """
        Se la worksheet esiste chiede conferma per sovrascrivere (clear).
        Se non esiste crea una nuova worksheet con nome ws_name.
        Ritorna la worksheet pronta per la scrittura o None se user cancella.
        """
        existing = self._get_worksheet_or_none(sh, ws_name)
        if existing:
            msg = (f"Il worksheet '{ws_name}' esiste già in '{sh.title}'.\n"
                f"Se confermi, tutto il contenuto presente verrà cancellato e sovrascritto.\n\n"
                "Vuoi proseguire?")
            if not messagebox.askyesno("Conferma sovrascrittura", msg):
                self._log_message("Utente ha annullato l'operazione di sovrascrittura.")
                return None
            # clear existing:
            existing.clear()
            self._log_message(f"Worksheet '{ws_name}' pulito (clear) prima della scrittura.")
            return existing
        else:
            # create worksheet with reasonable size
            rows = max(1000, total_rows_estimate + 10)
            cols = 10
            ws = self._create_worksheet(sh, ws_name, rows=rows, cols=cols)
            return ws

    def _read_rows_from_worksheet(self, sh, ws_name):
        """
        Legge tutte le righe dal worksheet e ritorna una lista di stringhe (una per riga),
        dove ogni riga è l'unione delle celle non vuote separate da spazio.
        """
        try:
            ws = sh.worksheet(ws_name)
        except WorksheetNotFound:
            raise
        all_values = ws.get_all_values()
        righe = []
        for row in all_values:
            testo = ' '.join([c for c in row if c and c.strip()])
            if testo:
                righe.append(testo)
        return righe

    # def _batch_write_worksheet(self, ws, rows_of_lists, start_row=1, batch_size=5000):
    #     """
    #     Scrive rows_of_lists (lista di liste) su worksheet ws partendo da start_row in batch.
    #     rows_of_lists: [[col1, col2, ...], ...]
    #     """
    #     total = len(rows_of_lists)
    #     if total == 0:
    #         return

    #     # gspread expects a 2D list. We'll upload in chunks.
    #     for i in range(0, total, batch_size):
    #         chunk = rows_of_lists[i:i+batch_size]
    #         top_row = start_row + i
    #         # compute range top-left cell
    #         cell = f"A{top_row}"
    #         # update() will fill the block
    #         ws.update(cell, chunk)
    #         self._log_message(f"Aggiornate righe {top_row} - {top_row + len(chunk) - 1} su '{ws.title}'")

    # ---------------- core: rielabora_sheets ----------------
    # def rielabora_sheets(self, src_spreadsheet, dest_spreadsheet, src_worksheet, dest_worksheet):
    #     """
    #     Legge dati da src_spreadsheet/src_worksheet, li riformatta usando leggi_file_dati
    #     (usando il tmp file trick), e scrive i risultati in dest_spreadsheet/dest_worksheet.
    #     Se il dest worksheet esiste chiede conferma per cancellare.
    #     """
    #     gc = self.connect_to_sheets()
    #     if not gc:
    #         messagebox.showerror("Errore", "Impossibile connettersi a Google Sheets.")
    #         return

    #     # --- apri sorgente
    #     try:
    #         sh_src = gc.open(src_spreadsheet)
    #     except SpreadsheetNotFound:
    #         self._log_message(f"Spreadsheet sorgente '{src_spreadsheet}' non trovato.")
    #         messagebox.showerror("Errore", f"Spreadsheet sorgente '{src_spreadsheet}' non trovato.")
    #         return
    #     except Exception as e:
    #         self._log_message(f"Errore aprendo spreadsheet sorgente: {e}")
    #         messagebox.showerror("Errore", f"Errore aprendo spreadsheet sorgente: {e}")
    #         return

    #     # read from source worksheet
    #     try:
    #         righe = self._read_rows_from_worksheet(sh_src, src_worksheet)
    #     except WorksheetNotFound:
    #         self._log_message(f"Worksheet sorgente '{src_worksheet}' non trovato nello spreadsheet '{src_spreadsheet}'.")
    #         messagebox.showerror("Errore", f"Worksheet sorgente '{src_worksheet}' non trovato.")
    #         return
    #     except Exception as e:
    #         self._log_message(f"Errore leggendo worksheet sorgente: {e}")
    #         messagebox.showerror("Errore", f"Errore leggendo worksheet sorgente: {e}")
    #         return

    #     if not righe:
    #         self._log_message("Nessuna riga da rielaborare nel foglio sorgente.")
    #         messagebox.showinfo("Info", "Nessuna riga valida nel foglio sorgente.")
    #         return

    #     self._log_message(f"Trovate {len(righe)} righe nel foglio sorgente. Procedo al parsing in blocchi...")

    #     # process in chunks to avoid mem/CPU peaks
    #     CHUNK_SIZE = 5000
    #     parsed_all = []
    #     for i in range(0, len(righe), CHUNK_SIZE):
    #         block = righe[i:i+CHUNK_SIZE]
    #         # scrivo temporaneo e riutilizzo leggi_file_dati
    #         with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as tmp:
    #             tmp.write("\n\n".join(block))
    #             tmp_path = tmp.name
    #         try:
    #             parsed = self.leggi_file_dati(tmp_path)
    #         finally:
    #             try:
    #                 os.remove(tmp_path)
    #             except Exception:
    #                 pass
    #         parsed_all.extend(parsed)
    #         self._log_message(f"Parsed chunk {i}..{i+len(block)-1} -> {len(parsed)} records")

    #     self._log_message(f"Parsing completato. Totale records: {len(parsed_all)}")

    #     # --- apri o crea dest spreadsheet
    #     try:
    #         sh_dest = self._ensure_spreadsheet(gc, dest_spreadsheet)
    #     except Exception as e:
    #         messagebox.showerror("Errore", f"Impossibile accedere/creare spreadsheet destinazione: {e}")
    #         return

    #     # if destination worksheet name empty -> fallback to month name (es. "August 2025")
    #     if not dest_worksheet or not dest_worksheet.strip():
    #         dest_worksheet = datetime.now().strftime("%B %Y")

    #     # check worksheet confirm/clear/create
    #     ws_dest = self._confirm_clear_or_create_ws(sh_dest, dest_worksheet, total_rows_estimate=len(parsed_all))
    #     if ws_dest is None:
    #         # user cancelled
    #         messagebox.showinfo("Annullato", "Operazione annullata dall'utente.")
    #         return

    #     # Prepare rows: header + data (nome, cognome, eta, occupazione, email, telefono)
    #     header = ["Nome", "Cognome", "Età", "Occupazione", "Email", "Telefono"]
    #     values = [header]
    #     for rec in parsed_all:
    #         # rec expected to be [nome, cognome, eta, occupation, email, telefono]
    #         values.append([rec[0], rec[1], rec[2], rec[3], rec[4], rec[5]])

    #     # write in batch (inizio riga 1)
    #     self._log_message(f"Scrittura di {len(values)-1} record su '{sh_dest.title}' -> '{ws_dest.title}'")
    #     try:
    #         self._batch_write_worksheet(ws_dest, values, start_row=1, batch_size=2000)
    #     except Exception as e:
    #         self._log_message(f"Errore durante la scrittura su Sheets: {e}")
    #         messagebox.showerror("Errore", f"Errore durante la scrittura su Sheets: {e}")
    #         return

    #     try:
    #         self.apply_sheet_styling(sh_dest, ws_dest)
    #     except Exception as e:
    #         self._log_message(f"Errore nell'applicare la formattazione: {e}")

    #     self._log_message("Scrittura completata con successo.")
    #     messagebox.showinfo("Completato", f"Rielaborazione eseguita: {len(values)-1} record scritti in '{sh_dest.title}'/'{ws_dest.title}'.")

    # --- helper: converte indice colonna in lettera (A, B, ..., Z, AA, AB, ...)


    # --- helper: normalizza un singolo record in una lista di 6 valori
    def _normalize_parsed_record(self, rec):
        """
        Accetta:
        - dict con chiavi 'name','surname','age','occupation','email','phone'
        - list/tuple con elementi in ordine possibile
        - stringa (raw) -> prova a riparsearla con parse_record_text se disponibile,
            altrimenti mette la stringa come 'name' e lascia gli altri vuoti.
        Ritorna sempre: [name, surname, age, occupation, email, phone]
        """
        # se dict
        if isinstance(rec, dict):
            return [
                rec.get('name', '') or '',
                rec.get('surname', '') or '',
                rec.get('age', '') or '',
                rec.get('occupation', '') or '',
                rec.get('email', '') or '',
                rec.get('phone', '') or ''
            ]
        # se list/tuple
        if isinstance(rec, (list, tuple)):
            out = []
            for i in range(6):
                out.append(str(rec[i]) if i < len(rec) and rec[i] is not None else '')
            return out
        # se stringa: prova a usare parse_record_text se esiste
        if isinstance(rec, str):
            try:
                parsed = None
                if 'parse_record_text' in globals() and callable(globals()['parse_record_text']):
                    parsed = globals()['parse_record_text'](rec)
                elif hasattr(self, 'parse_record_text') and callable(getattr(self, 'parse_record_text')):
                    parsed = getattr(self, 'parse_record_text')(rec)
                if parsed and isinstance(parsed, dict):
                    return [
                        parsed.get('name', '') or '',
                        parsed.get('surname', '') or '',
                        parsed.get('age', '') or '',
                        parsed.get('occupation', '') or '',
                        parsed.get('email', '') or '',
                        parsed.get('phone', '') or ''
                    ]
            except Exception:
                # fallthrough: non possiamo parse -> metti la stringa come name
                pass
            # fallback: metti tutta la stringa in 'name'
            return [rec.strip(), '', '', '', '', '']
        # fallback generico
        return ['', '', '', '', '', '']

    # --- versione robusta e difensiva di _batch_write_worksheet
    # def _batch_write_worksheet(self, ws, rows_of_lists, start_row=1, batch_size=5000):
    #     """
    #     Scrive rows_of_lists (lista di liste) su worksheet ws partendo da start_row in batch.
    #     Questo metodo:
    #     - verifica e normalizza il formato (tutti i record diventano liste)
    #     - calcola il numero di colonne effettivo come max(len(row)) tra le righe
    #     - usa update(range, values) con un range A{start}:<lastcol>{endrow}
    #     - evita il problema dei singoli caratteri se passi per errore stringhe
    #     """
    #     # difesa: assicurati che rows_of_lists sia una lista
    #     if not isinstance(rows_of_lists, (list, tuple)):
    #         raise ValueError("_batch_write_worksheet: rows_of_lists deve essere una lista di righe")

    #     # normalizza ogni riga: se la riga è stringa o dict la convertiamo in lista di celle
    #     normalized = []
    #     for r in rows_of_lists:
    #         if isinstance(r, (list, tuple)):
    #             # converti ogni cella in str (None->'')
    #             normalized.append(['' if c is None else str(c) for c in r])
    #         else:
    #             # r potrebbe essere dict o string -> prova normalizzazione
    #             normalized.append(self._normalize_parsed_record(r))

    #     total = len(normalized)
    #     if total == 0:
    #         return

    #     # determina numero colonne effettive (minimo 1)
    #     max_cols = max(len(row) for row in normalized)
    #     if max_cols == 0:
    #         max_cols = 1

    #     # upload in chunk
    #     for i in range(0, total, batch_size):
    #         chunk = normalized[i:i + batch_size]
    #         top_row = start_row + i
    #         bottom_row = top_row + len(chunk) - 1
    #         last_col_letter = _col_index_to_letter(max_cols - 1)
    #         range_str = f"A{top_row}:{last_col_letter}{bottom_row}"
    #         # ATTENZIONE: gspread.update accetta una 2D-list corrispondente al range
    #         ws.update(range_str, chunk)
    #         self._log_message(f"Aggiornate righe {top_row} - {bottom_row} su '{ws.title}' (cols={max_cols})")

    # def rielabora_sheets(self, src_spreadsheet, dest_spreadsheet, src_worksheet, dest_worksheet):
    #     gc = self.connect_to_sheets()
    #     if not gc:
    #         messagebox.showerror("Errore", "Impossibile connettersi a Google Sheets.")
    #         return

    #     # --- apri sorgente
    #     try:
    #         sh_src = gc.open(src_spreadsheet)
    #     except SpreadsheetNotFound:
    #         self._log_message(f"Spreadsheet sorgente '{src_spreadsheet}' non trovato.")
    #         messagebox.showerror("Errore", f"Spreadsheet sorgente '{src_spreadsheet}' non trovato.")
    #         return
    #     except Exception as e:
    #         self._log_message(f"Errore aprendo spreadsheet sorgente: {e}")
    #         messagebox.showerror("Errore", f"Errore aprendo spreadsheet sorgente: {e}")
    #         return

    #     # read from source worksheet (uso corretto di self)
    #     try:
    #         righe = self._read_rows_from_worksheet(sh_src, src_worksheet)
    #     except WorksheetNotFound:
    #         self._log_message(f"Worksheet sorgente '{src_worksheet}' non trovato nello spreadsheet '{src_spreadsheet}'.")
    #         messagebox.showerror("Errore", f"Worksheet sorgente '{src_worksheet}' non trovato.")
    #         return
    #     except Exception as e:
    #         self._log_message(f"Errore leggendo worksheet sorgente: {e}")
    #         messagebox.showerror("Errore", f"Errore leggendo worksheet sorgente: {e}")
    #         return

    #     if not righe:
    #         self._log_message("Nessuna riga da rielaborare nel foglio sorgente.")
    #         messagebox.showinfo("Info", "Nessuna riga valida nel foglio sorgente.")
    #         return

    #     # process in chunks...
    #     CHUNK_SIZE = 5000
    #     parsed_all = []
    #     for i in range(0, len(righe), CHUNK_SIZE):
    #         block = righe[i:i+CHUNK_SIZE]
    #         with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as tmp:
    #             tmp.write("\n\n".join(block))
    #             tmp_path = tmp.name
    #         try:
    #             parsed = self.leggi_file_dati(tmp_path)
    #         finally:
    #             try:
    #                 os.remove(tmp_path)
    #             except Exception:
    #                 pass
    #         parsed_all.extend(parsed)
    #         self._log_message(f"Parsed chunk {i}..{i+len(block)-1} -> {len(parsed)} records")

    #     # --- apri o crea dest spreadsheet (usa metodo di istanza)
    #     try:
    #         sh_dest = self._ensure_spreadsheet(gc, dest_spreadsheet)
    #     except Exception as e:
    #         messagebox.showerror("Errore", f"Impossibile accedere/creare spreadsheet destinazione: {e}")
    #         return

    #     if not dest_worksheet or not dest_worksheet.strip():
    #         dest_worksheet = datetime.now().strftime("%B %Y")

    #     ws_dest = self._confirm_clear_or_create_ws(sh_dest, dest_worksheet, total_rows_estimate=len(parsed_all))
    #     if ws_dest is None:
    #         messagebox.showinfo("Annullato", "Operazione annullata dall'utente.")
    #         return

    #     # Prepare rows: header + data
    #     header = ["Nome", "Cognome", "Età", "Occupazione", "Email", "Telefono"]
    #     values = [header] + [[r[0], r[1], r[2], r[3], r[4], r[5]] for r in parsed_all]

    #     # scrivi in batch usando metodo d'istanza
    #     try:
    #         self._batch_write_worksheet(ws_dest, values, start_row=1, batch_size=2000)
    #     except Exception as e:
    #         self._log_message(f"Errore durante la scrittura su Sheets: {e}")
    #         messagebox.showerror("Errore", f"Errore durante la scrittura su Sheets: {e}")
    #         return

    #     try:
    #         self.apply_sheet_styling(sh_dest, ws_dest)
    #     except Exception as e:
    #         self._log_message(f"Errore nell'applicare la formattazione: {e}")

    #     self._log_message("Scrittura completata con successo.")
    #     messagebox.showinfo("Completato", f"Rielaborazione eseguita: {len(values)-1} record scritti in '{sh_dest.title}'/'{ws_dest.title}'.")


    #     def leggi_da_google_sheets(self, spreadsheet_name, worksheet_name):
    #         """Legge i dati disordinati da un foglio di Google Sheets e li restituisce come lista di righe testuali."""
    #         self._log_message(f"Lettura dal foglio '{spreadsheet_name}' / '{worksheet_name}'...")
    #         gc = self.connect_to_sheets()
    #         if not gc:
    #             return []

    #         try:
    #             sh = gc.open(spreadsheet_name)
    #             ws = sh.worksheet(worksheet_name)
    #             all_values = ws.get_all_values()

    #             righe = []
    #             for row in all_values:
    #                 # unisci tutte le celle non vuote in un'unica stringa
    #                 testo = ' '.join([c for c in row if c.strip()])
    #                 if testo:
    #                     righe.append(testo)
    #             self._log_message(f"Trovate {len(righe)} righe non vuote nel foglio.")
    #             return righe
    #         except Exception as e:
    #             self._log_message(f"Errore durante la lettura da Sheets: {e}")
    #             messagebox.showerror("Errore", f"Impossibile leggere dal foglio: {e}")
    #             return []


    #     def rielabora_sheets(self, sorgente, destinazione, ws_sorgente, ws_destinazione):
    #         """
    #         Legge da un foglio sorgente, riformatta i dati e li esporta su un nuovo foglio.
    #         Esegue il processo in blocchi per non superare i limiti API.
    #         """
    #         dati_raw = self.leggi_da_google_sheets(sorgente, ws_sorgente)
    #         if not dati_raw:
    #             return

    #         self._log_message("Rielaborazione dei dati in corso...")
    #         blocchi = [dati_raw[i:i+5000] for i in range(0, len(dati_raw), 5000)]

    #         gc = self.connect_to_sheets()
    #         if not gc:
    #             return

    #         try:
    #             try:
    #                 sh_dest = gc.open(destinazione)
    #             except gspread.SpreadsheetNotFound:
    #                 self._log_message(f"Creazione del foglio '{destinazione}'...")
    #                 sh_dest = gc.create(destinazione)

    #             try:
    #                 ws_dest = sh_dest.worksheet(ws_destinazione)
    #             except WorksheetNotFound:
    #                 ws_dest = sh_dest.add_worksheet(title=ws_destinazione, rows=100, cols=7)

    #             all_parsed = []
    #             for blocco in blocchi:
    #                 # simula il parsing da file
    #                 with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as tmp:
    #                     tmp.write("\n\n".join(blocco))
    #                     tmp_path = tmp.name
    #                 dati_parsati = self.leggi_file_dati(tmp_path)
    #                 os.remove(tmp_path)
    #                 all_parsed.extend(dati_parsati)

    #             self._log_message(f"Totale righe formattate: {len(all_parsed)}")
    #             self.invia_a_google_sheets(all_parsed)
    #         except Exception as e:
    #             self._log_message(f"Errore durante la rielaborazione: {e}")


    # def write_parsed_to_destination(self, parsed_records, dest_spreadsheet, dest_worksheet):
    #     """
    #     Scrive parsed_records (lista di [nome,cognome,eta,occupazione,email,telefono]) su dest_spreadsheet/dest_worksheet.
    #     Comportamento: se worksheet esiste chiedi conferma e se sì clear + write, altrimenti crea.
    #     """
    #     if not parsed_records:
    #         messagebox.showinfo("Info", "Nessun record da scrivere.")
    #         return

    #     gc = self.connect_to_sheets()
    #     if not gc:
    #         messagebox.showerror("Errore", "Impossibile connettersi a Google Sheets.")
    #         return

    #     try:
    #         sh_dest = self._ensure_spreadsheet(gc, dest_spreadsheet)
    #     except Exception as e:
    #         messagebox.showerror("Errore", f"Impossibile accedere/creare spreadsheet destinazione: {e}")
    #         return

    #     if not dest_worksheet or not dest_worksheet.strip():
    #         dest_worksheet = datetime.now().strftime("%B %Y")

    #     ws_dest = self._confirm_clear_or_create_ws(sh_dest, dest_worksheet, total_rows_estimate=len(parsed_records))
    #     if ws_dest is None:
    #         messagebox.showinfo("Annullato", "Operazione annullata dall'utente.")
    #         return

    #     header = ["Nome", "Cognome", "Età", "Occupazione", "Email", "Telefono"]
    #     values = [header] + [[r[0], r[1], r[2], r[3], r[4], r[5]] for r in parsed_records]

    #     try:
    #         self._batch_write_worksheet(ws_dest, values, start_row=1, batch_size=2000)
    #     except Exception as e:
    #         self._log_message(f"Errore scrittura su Sheets: {e}")
    #         messagebox.showerror("Errore", f"Errore scrittura su Sheets: {e}")
    #         return

    #     self._log_message(f"Sono stati scritti {len(values)-1} record in '{sh_dest.title}/{ws_dest.title}'")
    #     messagebox.showinfo("Completato", f"Dati scritti in: {sh_dest.title} / {ws_dest.title}")





    

    def show_data_preview(self):
        """Mostra l'anteprima dei dati in una nuova finestra."""
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Errore", "Seleziona prima un file di testo.")
            return

        dati = self.leggi_file_dati(file_path)
        if not dati:
            messagebox.showinfo("Anteprima Dati", "Nessun dato valido trovato nel file.")
            return
            
        preview_text = "Anteprima dei dati da inviare:\n\n"
        for riga in dati:
            preview_text += f"Nome: {riga[0]}, Cognome: {riga[1]}, Età: {riga[2]}, Occupazione: {riga[3]}, Email: {riga[4]}, Telefono: {riga[5]}\n"
        
        preview_window = tk.Toplevel(self.master)
        preview_window.title("Anteprima Dati")
        preview_window.geometry("600x400")
        preview_window.grab_set()

        text_widget = tk.Text(preview_window, wrap=WORD)
        text_widget.pack(fill=BOTH, expand=True, padx=10, pady=10)
        text_widget.insert(tk.END, preview_text)
        text_widget.configure(state='disabled')

        ttk.Button(preview_window, text="Chiudi", command=preview_window.destroy, bootstyle="info").pack(pady=10)
    
    def _configura_checkbox(self, worksheet, start_row, num_rows):
        """Configura le checkbox nella colonna A"""
        try:
            if num_rows <= 0:
                return
                
            requests = {
                "requests": [{
                    "setDataValidation": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": start_row + num_rows - 1,
                            "startColumnIndex": 0, 
                            "endColumnIndex": 1
                        },
                        "rule": {
                            "condition": {
                                "type": "BOOLEAN"
                            },
                            "inputMessage": "Seleziona/Deseleziona",
                            "strict": True,
                            "showCustomUi": True  
                        }
                    }
                }]
            }
            
            
            worksheet.spreadsheet.batch_update(requests)
            
        except Exception as e:
            self._log_message(f"Errore durante la configurazione delle checkbox: {str(e)}")

    def invia_a_google_sheets(self, dati_emails):
        """Invia i dati a Google Sheets con checkbox funzionanti e formattazione corretta"""
        self._log_message("Connessione a Google Sheets...")
        gc = self.connect_to_sheets()
        if not gc:
            return

        spreadsheet_name = self.spreadsheet_name_var.get()
        worksheet_name = self.worksheet_name_var.get()

        if not spreadsheet_name or not worksheet_name:
            self._log_message("ERRORE: Inserisci il nome del foglio di calcolo e del foglio di lavoro.")
            messagebox.showerror("Errore", "Per favore, inserisci il nome del Foglio di Calcolo e del Foglio di Lavoro.")
            return

        try:
            # Apri o crea il foglio di calcolo
            try:
                sh = gc.open(spreadsheet_name)
            except gspread.SpreadsheetNotFound:
                self._log_message(f"Creazione del foglio di calcolo '{spreadsheet_name}'...")
                sh = gc.create(spreadsheet_name)
                if os.getenv("GSPREAD_EMAIL"):
                    sh.share(os.getenv("GSPREAD_EMAIL"), perm_type='user', role='writer')

            # Apri o crea il foglio di lavoro
            try:
                worksheet = sh.worksheet(worksheet_name)
                existing_data = worksheet.get_all_values()
            except WorksheetNotFound:
                self._log_message(f"Creazione del foglio di lavoro '{worksheet_name}'...")
                worksheet = sh.add_worksheet(title=worksheet_name, rows=len(dati_emails)+10, cols=7)  
                existing_data = []
            
            # Intestazioni corrette della tabella (A1:G1)
            headers = ["✅", "Nome", "Cognome", "Età", "Occupazione", "Email", "Numero di telefono"]
            
            # Se il foglio è vuoto, aggiungi le intestazioni
            if not existing_data:
                worksheet.update(values=[headers], range_name='A1:G1')
            
            # Prepara i nuovi dati (partendo da A2)
            new_rows = []
            for riga in dati_emails:
                new_row = [
                    False,  # Checkbox inizialmente non selezionata (colonna A)
                    riga[0],  # Nome (colonna B)
                    riga[1],  # Cognome (colonna C)
                    riga[2],  # Età (colonna D)
                    riga[3],  # Occupazione (colonna E)
                    riga[4],  # Email (colonna F)
                    riga[5]  # Telefono (colonna G)
                ]
                new_rows.append(new_row)
            
            # Trova la prima riga vuota (partendo da A2)
            next_row = len(existing_data) + 1 if existing_data else 2
            
            # Aggiungi i nuovi dati a partire da A2
            if new_rows:
                worksheet.update(values=new_rows, range_name=f'A{next_row}:G{next_row + len(new_rows) - 1}')
                self._log_message(f"Aggiunte {len(new_rows)} righe alla tabella")
            
            # Configura le checkbox nella colonna A
            self._configura_checkbox(worksheet, next_row, len(new_rows))
            
            # Formattazione della tabella
            self._formatta_tabella(worksheet, next_row + len(new_rows) - 1)
            
            # Applica la formattazione condizionale
            self._applica_formattazione_condizionale(worksheet, next_row, len(new_rows))
            
            return worksheet

        except Exception as e:
            self._log_message(f"Errore durante l'invio a Google Sheets: {e}")
            messagebox.showerror("Errore", f"Si è verificato un errore: {e}")
            return None

    
    def _applica_formattazione_condizionale(self, worksheet, start_row, num_rows):
        """Applica la formattazione condizionale corretta (solo fino alla colonna G)"""
        try:
            if num_rows <= 0:
                return
                
            sheet_id = worksheet.id
            end_row = start_row + num_rows - 1  
            
            # Crea la richiesta di formattazione condizionale
            requests = {
                "requests": [{
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{
                                "sheetId": sheet_id,
                                "startRowIndex": start_row - 1,  # -1 perché l'index è 0-based
                                "endRowIndex": end_row,          # Fino all'ultima riga dei dati (esclusa)
                                "startColumnIndex": 0,          # Dalla colonna A
                                "endColumnIndex": 7             # Fino a colonna G (esclusa)
                            }],
                            "booleanRule": {
                                "condition": {
                                    "type": "CUSTOM_FORMULA",
                                    "values": [{"userEnteredValue": "=INDIRECT(\"A\"&ROW())=TRUE"}]
                                },
                                "format": {
                                    "textFormat": {
                                        "strikethrough": True,
                                        "foregroundColor": {"red": 0.6, "green": 0.6, "blue": 0.6}
                                    },
                                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95}
                                }
                            }
                        },
                        "index": 0
                    }
                }]
            }
            
            # Invia la richiesta
            worksheet.spreadsheet.batch_update(requests)
            
        except Exception as e:
            self._log_message(f"Errore durante la formattazione condizionale: {str(e)}")

    def _formatta_tabella(self, worksheet, last_row):
        """Applica la formattazione base alla tabella"""
        try:
            # Formatta le intestazioni (A1:G1)
            requests = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 7
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0},
                                "horizontalAlignment": "CENTER"
                            }
                        },
                        "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)"
                    }
                },
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": last_row,
                            "startColumnIndex": 0,
                            "endColumnIndex": 7
                        },
                        "top": {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                        "left": {"style": "SOLID"},
                        "right": {"style": "SOLID"},
                        "innerHorizontal": {"style": "SOLID"},
                        "innerVertical": {"style": "SOLID"}
                    }
                },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": worksheet.id,
                                "dimension": "COLUMNS",
                                "startIndex": 0,
                                "endIndex": 1
                            },
                            "properties": {
                                "pixelSize": 60  # Larghezza colonna checkbox
                            },
                            "fields": "pixelSize"
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": worksheet.id,
                                "dimension": "COLUMNS",
                                "startIndex": 1,
                                "endIndex": 7
                            },
                            "properties": {
                                "pixelSize": 120  # Larghezza colonne dati
                            },
                            "fields": "pixelSize"
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": worksheet.id,
                                "dimension": "COLUMNS",
                                "startIndex": 5,
                                "endIndex": 6
                            },
                            "properties": {
                                "pixelSize": 180  # Larghezza colonna email
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
            }
            
            worksheet.spreadsheet.batch_update(requests)
            
        except Exception as e:
            self._log_message(f"Errore durante la formattazione della tabella: {str(e)}")

    def svuota_file_dati(self):
        """Svuota il contenuto del file di dati se l'opzione è attiva."""
        if self.svuota_file_var.get():
            file_path = self.file_path_var.get()
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write("")
                    self._log_message("File dati svuotato con successo.")
                except Exception as e:
                    self._log_message(f"Errore durante lo svuotamento del file: {e}")

    # def run_bot(self):
    #     """Logica principale del bot — gestisce input file e input sheets, e scrittura su spreadsheet di destinazione."""
    #     try:
    #         mode = self.source_mode.get() if hasattr(self, "source_mode") else "file"
    #         if mode == "sheets":
    #             # leggi dal foglio sorgente e scrivi su destinazione
    #             src_spreadsheet = self.src_spreadsheet_var.get().strip()
    #             src_worksheet = self.src_worksheet_var.get().strip()
    #             dest_spreadsheet = self.spreadsheet_name_var.get().strip()
    #             dest_worksheet = self.worksheet_name_var.get().strip()

    #             if not src_spreadsheet or not src_worksheet:
    #                 messagebox.showerror("Errore", "Per favore specifica il Spreadsheet sorgente e il Worksheet sorgente.")
    #                 return
    #             if not dest_spreadsheet:
    #                 messagebox.showerror("Errore", "Per favore specifica il Spreadsheet di destinazione.")
    #                 return

    #             self._log_message(f"Avvio rielaborazione: '{src_spreadsheet}/{src_worksheet}' -> '{dest_spreadsheet}/{dest_worksheet or '(mese corrente)'}'")
    #             self.rielabora_sheets(src_spreadsheet, dest_spreadsheet, src_worksheet, dest_worksheet)
    #             return

    #         # ------------------- modalità file -------------------
    #         file_path = self.file_path_var.get().strip()
    #         if not file_path or not os.path.exists(file_path):
    #             messagebox.showerror("Errore", "Per favore, seleziona un file dati valido.")
    #             return

    #         dest_spreadsheet = self.spreadsheet_name_var.get().strip()
    #         dest_worksheet = self.worksheet_name_var.get().strip()
    #         if not dest_spreadsheet:
    #             messagebox.showerror("Errore", "Per favore specifica il Spreadsheet di destinazione.")
    #             return

    #         self._log_message(f"Lettura file: {file_path}")
    #         parsed = self.leggi_file_dati(file_path)
    #         if not parsed:
    #             messagebox.showinfo("Info", "Nessun record valido da processare.")
    #             return

    #         self._log_message(f"Parsed {len(parsed)} record. Avvio scrittura su '{dest_spreadsheet}/{dest_worksheet or '(mese corrente)'}'...")
    #         self.write_parsed_to_destination(parsed, dest_spreadsheet, dest_worksheet)

    #         # svuota file se richiesto
    #         if getattr(self, "svuota_file_var", tk.BooleanVar()) and self.svuota_file_var.get():
    #             self.svuota_file_dati()

    #     except Exception as e:
    #         self._log_message(f"Errore critico in run_bot: {e}")
    #         messagebox.showerror("Errore critico", f"Errore: {e}")

    # def run_bot(self):
    #     """Logica principale del bot (file o sheets come sorgente)."""
    #     source_mode = self.source_mode.get() if hasattr(self, 'source_mode') else 'file'
    #     self._log_message(f"Modalità sorgente: {source_mode}")

    #     # Salva le scelte correnti (persistenza)
    #     try:
    #         self._save_config_value('PATHS', 'dati_emails_path', self.file_path_var.get() or '')
    #         self._save_config_value('SHEETS', 'spreadsheet_name', self.spreadsheet_name_var.get() or '')
    #         self._save_config_value('SHEETS', 'worksheet_name', self.worksheet_name_var.get() or '')
    #         self._save_config_value('SETTINGS', 'source_mode', source_mode)
    #     except Exception:
    #         pass

    #     dati_emails = []
    #     if source_mode == 'file':
    #         file_path = self.file_path_var.get()
    #         if not file_path or not os.path.exists(file_path):
    #             self._log_message("ERRORE: Seleziona un file valido prima di avviare il bot.")
    #             messagebox.showerror("Errore", "Per favore, seleziona un file dati valido.")
    #             return
    #         self._log_message("Lettura e elaborazione del file dati...")
    #         dati_emails = self.leggi_file_dati(file_path)
    #     else:
    #         # sheets: si assume che siano già configurate le credenziali e il nome foglio/worksheet
    #         ss_name = self.spreadsheet_name_var.get().strip()
    #         ws_name = self.worksheet_name_var.get().strip()
    #         if not ss_name:
    #             self._log_message("ERRORE: Inserisci il nome del Google Spreadsheet sorgente.")
    #             messagebox.showerror("Errore", "Inserisci il nome del Google Spreadsheet sorgente.")
    #             return
    #         try:
    #             self._log_message(f"Lettura righe dallo Sheet sorgente: {ss_name} / {ws_name}")
    #             dati_emails = self.leggi_sheet_sorgente(ss_name, ws_name)
    #         except Exception as e:
    #             self._log_message(f"Errore lettura sheet sorgente: {e}")
    #             messagebox.showerror("Errore", f"Impossibile leggere lo Sheet sorgente: {e}")
    #             return

    #     if not dati_emails:
    #         self._log_message("Nessun dato valido trovato. Operazione completata.")
    #         messagebox.showinfo("Bot", "Nessun dato valido da inviare. Operazione completata.")
    #         return

    #     self._log_message(f"{len(dati_emails)} blocchi di dati validi trovati.")
    #     try:
    #         self.invia_a_google_sheets(dati_emails)
    #     except Exception as e:
    #         self._log_message(f"Errore invio a Google Sheets: {e}")
    #         messagebox.showerror("Errore", f"Errore durante l'invio a Google Sheets: {e}")
    #         return

    #     if self.svuota_file_var.get() and source_mode == 'file':
    #         self.svuota_file_dati()

    #     self._log_message("Processo bot completato.")
    #     messagebox.showinfo("Bot", "Operazione completata con successo.")



    # --- Nuove funzioni UI/UX: aggiornamento in place con progress bar e disabilitazione UI ---
    def start_update(self):
        """Mostra conferma all'utente e avvia l'aggiornamento mostrando progress."""
        if not messagebox.askyesno("Aggiornamento", "È stato trovato un aggiornamento. Vuoi installarlo ora? L'app si chiuderà." ):
            return

        # Disabilita i controlli principali per evitare azioni concorrenti
        self.disable_ui_for_update()

        # Mostra progress e testo
        self.status_label.config(text="Download e installazione in corso...")
        self.update_progress.pack(side=tk.RIGHT, padx=(0,8))
        self.update_progress.start(10)

        # Avvia l'update in thread (Updater terminerà il processo se tutto ok)
        Thread(target=self._run_update_thread, daemon=True).start()

    def _run_update_thread(self):
        try:
            self.updater.update_app()
        except Exception as e:
            # Se rientra qui vuol dire che l'update ha fallito senza terminare il processo
            self._log_message(f"Errore aggiornamento: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Errore durante l'aggiornamento: {e}")
            # ripristina UI
            self.update_progress.stop()
            self.update_progress.pack_forget()
            self.enable_ui_after_update()

    def disable_ui_for_update(self):
        try:
            self.run_button.configure(state='disabled')
            self.preview_button.configure(state='disabled')
            self.browse_button.configure(state='disabled')
            self.ss_entry.configure(state='disabled')
            self.ws_entry.configure(state='disabled')
        except Exception:
            pass

    def enable_ui_after_update(self):
        try:
            self.run_button.configure(state='normal')
            self.preview_button.configure(state='normal')
            self.browse_button.configure(state='normal')
            self.ss_entry.configure(state='normal')
            self.ws_entry.configure(state='normal')
            self.status_label.config(text=f"v{CURRENT_VERSION}")
        except Exception:
            pass

def _col_index_to_letter(n):
    """0-based index -> column letters (0 -> 'A')."""
    letters = ''
    while True:
        letters = chr(ord('A') + (n % 26)) + letters
        n = n // 26 - 1
        if n < 0:
            break
    return letters

# --- Punto di Ingresso dell'applicazione ---
if __name__ == "__main__":
    try:
        app = ttk.Window(themename="lumen")
        BotApp(app)
        app.mainloop()
    except Exception as e:
        import traceback
        print(f"Errore critico: {e}")
        traceback.print_exc()