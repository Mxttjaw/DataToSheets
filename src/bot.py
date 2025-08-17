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

import requests
import gspread
from gspread.exceptions import WorksheetNotFound, APIError
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

CURRENT_VERSION = "1.0.6"

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

        self.file_path_var = tk.StringVar()
        self.svuota_file_var = tk.BooleanVar()
        self.spreadsheet_name_var = tk.StringVar()
        self.worksheet_name_var = tk.StringVar()
        self.tutorial_state_var = tk.BooleanVar()
        self.update_available = tk.BooleanVar(value=False)
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

        # Avvia la verifica degli aggiornamenti in un thread separato
        Thread(target=lambda: self.updater.check_for_updates(on_update_available=lambda: self.update_available.set(True)), daemon=True).start()

    # --- Metodi per la gestione della GUI ---

    def create_gui_elements(self):
        """Costruisce tutti i widget dell'interfaccia utente."""
        self.master.title(f"Bot Email - Gestione Dati (v{CURRENT_VERSION})")
        self.master.geometry("700x460")
        self.master.resizable(True, True)
        
        # Top frame: menu + status
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

        # Progressbar nascosta che appare durante download/aggiornamento
        self.update_progress = ttk.Progressbar(status_frame, length=180, mode='indeterminate')
        self.update_progress.pack(side=tk.RIGHT, padx=(0,8))
        self.update_progress.pack_forget()

    def _create_file_selection_area(self):
        """Crea il frame per la selezione del file e i campi per foglio di calcolo/lavoro."""
        file_frame = ttk.LabelFrame(self.main_content_frame, text="Configurazione Dati e Fogli", padding=(15, 10))
        file_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 15))
        
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(path_frame, text="File Dati:").pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, state='readonly')
        self.file_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        
        self.browse_button = ttk.Button(path_frame, text="Scegli File", command=self.select_file, bootstyle="primary")
        self.browse_button.pack(side=tk.RIGHT)
        
        ss_frame = ttk.Frame(file_frame)
        ss_frame.pack(fill=tk.X, pady=(5, 5))
        
        ttk.Label(ss_frame, text="Nome Foglio di Calcolo:").pack(side=tk.LEFT)
        self.ss_entry = ttk.Entry(ss_frame, textvariable=self.spreadsheet_name_var)
        self.ss_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        
        ws_frame = ttk.Frame(file_frame)
        ws_frame.pack(fill=tk.X, pady=(5, 5))
        
        ttk.Label(ws_frame, text="Nome Foglio di Lavoro:").pack(side=tk.LEFT)
        self.ws_entry = ttk.Entry(ws_frame, textvariable=self.worksheet_name_var)
        self.ws_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

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
            # Mostra pulsante e aggiorna la status bar
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

    # --- Logica di Configurazione e Stato ---

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
        """Carica il percorso del file e le opzioni salvate all'avvio."""
        self.file_path_var.set(self._load_config_string('PATHS', 'dati_emails_path'))
        self.svuota_file_var.set(self._load_config_boolean('SETTINGS', 'svuota_file', False))
        
        spreadsheet_name_from_env = os.getenv("SPREADSHEET_NAME")
        self.spreadsheet_name_var.set(self._load_config_string('SHEETS', 'spreadsheet_name', spreadsheet_name_from_env or ''))
        
        nomi_mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
        nome_foglio_mese = nomi_mesi[datetime.now().month - 1]
        self.worksheet_name_var.set(self._load_config_string('SHEETS', 'worksheet_name', nome_foglio_mese))
        
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
            message = message.replace('Ã¨', 'è').replace('Ã', 'à')  # Correzione specifica
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

    # --- Metodi per la gestione della configurazione (configparser) ---

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

        # Temi
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
        """Apre la finestra del tutorial."""
        tutorial_window = tk.Toplevel(self.master)
        tutorial_window.title("Tutorial")
        tutorial_window.geometry("500x300")
        tutorial_window.resizable(False, False)
        tutorial_window.grab_set()

        tutorial_frame = ttk.Frame(tutorial_window, padding=20)
        tutorial_frame.pack(fill=BOTH, expand=True)

        tutorial_text = (
            "Benvenuto nel Bot per la gestione dei dati su Google Sheets!\n\n"
            "Questo bot ti aiuta a estrarre i dati da un file di testo e a caricarli "
            "automaticamente su un foglio di calcolo Google Sheets.\n\n"
            "Passaggi:\n"
            "1. Clicca su 'Scegli File' per selezionare il file di testo contenente i dati.\n"
            "2. Inserisci il nome del Foglio di Calcolo e del Foglio di Lavoro.\n"
            "3. Clicca su 'Anteprima Dati' per vedere cosa verrà inviato.\n"
            "4. Clicca su 'Avvia Bot' per caricare i dati su Google Sheets.\n"
            "5. Nelle 'Opzioni' puoi cambiare il tema e attivare lo svuotamento automatico del file.\n\n"
            "Assicurati di avere il file 'credentials.json' e il file '.env' configurati correttamente."
        )
        
        ttk.Label(tutorial_frame, text=tutorial_text, justify=LEFT, wraplength=450).pack(fill=BOTH, expand=True)
        ttk.Button(tutorial_frame, text="Chiudi", command=tutorial_window.destroy, bootstyle="info").pack(pady=10)

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
    
    def leggi_file_dati(self, file_path):
        """
        Legge un file di testo, lo divide in blocchi e ritorna i dati estratti in modo robusto.
        Rimuove sempre la porzione 'SISTEMA ...' (es. 'SISTEMA INVIO ONLINE - SITO WEB' o
        'SISTEMA DI INVIO ...') prima del parsing così quella parte non viene mai inviata su Google Sheets.
        """
        dati_emails = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                contenuto = f.read()

            blocchi = re.split(r'\n{2,}', contenuto.strip())

            email_pattern = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', re.IGNORECASE)
            phone_pattern = re.compile(r'\+?[0-9][0-9 ()\-\.]{5,}[0-9]')

            COUNTRY_CODES = [
                '1242','1246','1264','1268','1284','1340','1345','1441','1649','1664','1721','1758',
                '1767','1784','1787','1809','1829','1849','380','385','386','234','91','39','44','33',
                '34','49','36','30','31','32','52','54','55','56','57','58','60','61','62','63','64',
                '65','66','7','1','20','27','86','81','82','84','90','92','98'
            ]
            COUNTRY_CODES = sorted(set(COUNTRY_CODES), key=lambda x: -len(x))

            def format_phone(phone_raw):
                if not phone_raw or not phone_raw.strip():
                    return ''
                s = phone_raw.strip()
                digits = re.sub(r'[^0-9]', '', s)

                if s.startswith('+') and digits:
                    matched = None
                    for cc in COUNTRY_CODES:
                        if digits.startswith(cc):
                            rest = digits[len(cc):]
                            if len(rest) >= 6:
                                matched = cc
                                break
                    if not matched:
                        m_space = re.match(r'^\+([^\s]+)\s', s)
                        if m_space:
                            cc_candidate = re.sub(r'[^0-9]', '', m_space.group(1))
                            if cc_candidate and digits.startswith(cc_candidate) and len(digits[len(cc_candidate):]) >= 6:
                                matched = cc_candidate
                    if not matched:
                        if len(digits) > 2 and len(digits[2:]) >= 6:
                            matched = digits[:2]
                        elif len(digits) > 1 and len(digits[1:]) >= 6:
                            matched = digits[:1]
                        else:
                            matched = ''
                    if matched:
                        rest = digits[len(matched):]
                        return f'+{matched} {rest}'
                    else:
                        return '+' + digits
                else:
                    return digits

            # parole chiave occupazione
            OCCUPATION_KEYWORDS = {
                'data','scientist','developer','engineer','ingegnere','analyst','analista',
                'manager','director','cto','ceo','founder','owner','product','designer','marketing',
                'sales','consultant','consulente','responsabile','head','research','ricerca',
                'system','sistema','invio','online','site','sito','web','interno','internal',
                'operations','operation','specialist','teacher','professor','prof','doctor','dr',
                'student','studente','amministratore','admin','support','supporto','dev','architect'
            }

            def token_is_occupation(tok):
                if not tok:
                    return False
                t = re.sub(r'[^A-Za-z0-9]', '', tok).lower()
                if not t:
                    return False
                if t in OCCUPATION_KEYWORDS:
                    return True
                for kw in OCCUPATION_KEYWORDS:
                    if kw in t:
                        return True
                return False

            # pattern per rimuovere la porzione "SISTEMA ...", compresi "SISTEMA DI INVIO"
            SYSTEM_REMOVE_RE = re.compile(r'\bSISTEMA(?:\s+DI)?\s+INVIO\b.*', re.IGNORECASE)
            # rimuove anche tag di tipo "EMAIL_INTERNAL", "SITO WEB", "ONLINE" eventuali rimasugli
            TRAILING_TAGS_RE = re.compile(r'\b(EMAIL_INTERNAL|SITO\s+WEB|SITO|ONLINE|SITE|WEB|INTERNAL)\b.*', re.IGNORECASE)

            for blocco in blocchi:
                if not blocco.strip():
                    continue
                righe = [line.strip() for line in blocco.split('\n') if line.strip()]
                text = ' '.join(righe)

                email_match = email_pattern.search(text)
                email = email_match.group(0).strip() if email_match else ''

                phone_match = phone_pattern.search(text)
                telefono_raw = phone_match.group(0).strip() if phone_match else ''
                telefono_formatted = format_phone(telefono_raw)

                # Rimuovo email e telefono dal testo prima di cercare l'età
                text_for_age = text
                if telefono_raw:
                    text_for_age = re.sub(re.escape(telefono_raw), ' ', text_for_age)
                    text_for_age = re.sub(r'\+\d[\d\s\-\.\(\)]{4,}\d', ' ', text_for_age)
                if email:
                    text_for_age = re.sub(re.escape(email), ' ', text_for_age)

                # Rimuovo sempre la porzione 'SISTEMA ...' anche dal text_for_age così non influenza nulla
                text_for_age = SYSTEM_REMOVE_RE.sub(' ', text_for_age)
                text_for_age = TRAILING_TAGS_RE.sub(' ', text_for_age)

                # ricerca età dopo aver tolto telefono/email e 'SISTEMA...'
                age_match = re.search(r'\b([1-9][0-9]{0,2})\b', text_for_age)
                age = ''
                start_age = end_age = None
                if age_match:
                    candidate = int(age_match.group(1))
                    if 10 <= candidate <= 120:
                        age = str(candidate)
                        start_age = age_match.start(1)
                        end_age = age_match.end(1)

                prima_riga = righe[0] if righe else ''
                # importante: rimuovo la porzione 'SISTEMA...' dalla prima riga PRIMA di tokenizzare
                prima_riga = SYSTEM_REMOVE_RE.sub(' ', prima_riga).strip()
                prima_riga = TRAILING_TAGS_RE.sub(' ', prima_riga).strip()

                fullname = ''
                occupation = ''

                if age:
                    # se l'età è presente, dividiamo attorno alla sua posizione
                    m_in_first = re.search(r'\b' + re.escape(age) + r'\b', prima_riga)
                    if m_in_first:
                        before = prima_riga[:m_in_first.start()].strip()
                        after = prima_riga[m_in_first.end():].strip()
                    else:
                        before = text_for_age[:start_age].strip()
                        after = text_for_age[end_age:].strip()
                    fullname = before
                    # rimuovo eventuale "SISTEMA..." in after per sicurezza
                    after = SYSTEM_REMOVE_RE.sub(' ', after)
                    after = TRAILING_TAGS_RE.sub(' ', after)
                    sys_match = re.search(r'\bSISTEMA\s+INVIO\b', after, re.IGNORECASE)
                    if sys_match:
                        occupation = after[:sys_match.start()].strip()
                    else:
                        occupation = after
                else:
                    # Se manca età: cerchiamo indice in cui inizia l'occupazione nella prima riga
                    prima_riga_clean = prima_riga.strip()
                    prima_riga_clean = re.sub(r'\s*-\s*', ' - ', prima_riga_clean).strip()
                    tokens = [t for t in re.split(r'\s+', prima_riga_clean) if t != '']

                    occ_idx = None
                    for i, tok in enumerate(tokens):
                        if token_is_occupation(tok):
                            occ_idx = i
                            break
                    if occ_idx is not None and occ_idx >= 1:
                        fullname = ' '.join(tokens[:occ_idx]).strip()
                        occupation = ' '.join(tokens[occ_idx:]).strip()
                    else:
                        if len(tokens) <= 2:
                            fullname = ' '.join(tokens).strip()
                            occupation = ''
                        else:
                            if '-' in tokens:
                                dash_idx = tokens.index('-')
                                if dash_idx >= 1:
                                    fullname = ' '.join(tokens[:dash_idx]).strip()
                                    occupation = ' '.join(tokens[dash_idx+1:]).strip()
                                else:
                                    fullname = ' '.join(tokens[:2]).strip()
                                    occupation = ' '.join(tokens[2:]).strip()
                            else:
                                fullname = ' '.join(tokens[:2]).strip()
                                occupation = ' '.join(tokens[2:]).strip()

                # pulizia definitiva: rimuovo sempre 'SISTEMA...' e tag finali dall'occupation
                occupation = SYSTEM_REMOVE_RE.sub(' ', occupation).strip()
                occupation = TRAILING_TAGS_RE.sub(' ', occupation).strip()
                occupation = re.sub(r'\s{2,}', ' ', occupation).strip()

                # split fullname in nome / cognome (nome = primo token, cognome = resto)
                nome, cognome = '', ''
                if fullname:
                    parts = fullname.split()
                    if len(parts) == 1:
                        nome = parts[0]
                        cognome = ''
                    else:
                        nome = parts[0]
                        cognome = ' '.join(parts[1:])

                # fallback: email e telefono da righe specifiche (se non già trovati)
                if not email and len(righe) > 1 and '@' in righe[1]:
                    email = righe[1].strip()
                if (not telefono_formatted) and len(righe) > 2:
                    telefono_formatted = format_phone(righe[2])

                if telefono_formatted and (len(re.sub(r'[^0-9]', '', telefono_formatted)) < 6):
                    self._log_message(f"Telefono sospetto trovato: {telefono_formatted} nel blocco: {prima_riga}")

                dati = [nome or '', cognome or '', age or '', occupation or '', email or '', telefono_formatted or '']
                dati_emails.append(dati)

            return dati_emails

        except FileNotFoundError:
            self._log_message(f"Errore: File '{file_path}' non trovato.")
            return []
        except Exception as e:
            self._log_message(f"Errore durante la lettura del file: {e}")
            return []









    

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

    def run_bot(self):
        """Logica principale del bot."""
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            self._log_message("ERRORE: Seleziona un file valido prima di avviare il bot.")
            messagebox.showerror("Errore", "Per favore, seleziona un file dati valido.")
            return

        self._log_message("Lettura e elaborazione del file dati...")
        dati_emails = self.leggi_file_dati(file_path)

        if not dati_emails:
            self._log_message("Nessun dato valido trovato nel file. Operazione completata.")
            messagebox.showinfo("Bot", "Nessun dato valido da inviare. Operazione completata.")
            return

        self._log_message(f"{len(dati_emails)} blocchi di dati validi trovati.")
        
        self.invia_a_google_sheets(dati_emails)
        
        if self.svuota_file_var.get():
            self.svuota_file_dati()
            
        self._log_message("Processo bot completato.")
        messagebox.showinfo("Bot", "Operazione completata con successo.")

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
