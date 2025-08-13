# python
import os
import re
import sys
import json
import time
import shutil
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

# --- Configurazione e Dipendenze ---
# Imposta il percorso di base come la directory dell'eseguibile per trovare i file di configurazione.
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Carica le variabili d'ambiente
load_dotenv()

CURRENT_VERSION = "0.9.9"


class BotApp(ttk.Frame):
    """
    Classe principale dell'applicazione, gestisce l'intera interfaccia utente
    e la logica del bot. Ora supporta la gestione di file utente-specifici.
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

        # FIX: Ho spostato la creazione della GUI come prima cosa in __init__
        # per evitare che qualsiasi chiamata a _log_message fallisca.
        self.create_gui_elements()
        
        # Le seguenti funzioni ora vengono chiamate solo dopo che la GUI è pronta.
        self._create_user_config_directory()
        self.tutorial_state_var.set(self._load_config_boolean('SETTINGS', 'tutorial_shown', False))
        
        self.load_initial_configuration()

        if not self._load_config_boolean('SETTINGS', 'tutorial_shown', False):
            self.master.after(500, self.show_tutorial_window)
            self._save_config_value('SETTINGS', 'tutorial_shown', True)

        Thread(target=self.check_for_updates, daemon=True).start()

    # --- Metodi per la gestione della GUI ---

    def create_gui_elements(self):
        """Costruisce tutti i widget dell'interfaccia utente."""
        self.master.title(f"Bot Email - Gestione Dati (v{CURRENT_VERSION})")
        self.master.geometry("600x400")
        self.master.resizable(True, True)
        
        self.top_menu_frame = ttk.Frame(self, height=30)
        self.top_menu_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        
        self.main_content_frame = ttk.Frame(self)
        self.main_content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self.main_content_frame.grid_rowconfigure(1, weight=1)
        self.main_content_frame.grid_columnconfigure(0, weight=1)

        self._create_dropdown_menu()
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
        menu.add_command(label="Controlla Aggiornamenti", command=lambda: Thread(target=self.check_for_updates, daemon=True).start())
        menu.add_separator()
        menu.add_command(label="Esci", command=self.master.quit)
        
        def show_menu():
            """Mostra il menu a tendina nella posizione corretta."""
            try:
                menu.tk_popup(menu_btn.winfo_rootx(), menu_btn.winfo_rooty() + menu_btn.winfo_height())
            finally:
                menu.grab_release()
        
        menu_btn.config(command=show_menu)

    def _create_file_selection_area(self):
        """Crea il frame per la selezione del file e i campi per foglio di calcolo/lavoro."""
        file_frame = ttk.LabelFrame(self.main_content_frame, text="Configurazione Dati e Fogli", padding=(15, 10))
        file_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 15))
        
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(path_frame, text="File Dati:").pack(side=tk.LEFT)
        file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, state='readonly')
        file_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        
        browse_button = ttk.Button(path_frame, text="Scegli File", command=self.select_file, bootstyle="primary")
        browse_button.pack(side=tk.RIGHT)
        
        ss_frame = ttk.Frame(file_frame)
        ss_frame.pack(fill=tk.X, pady=(5, 5))
        
        ttk.Label(ss_frame, text="Nome Foglio di Calcolo:").pack(side=tk.LEFT)
        ss_entry = ttk.Entry(ss_frame, textvariable=self.spreadsheet_name_var)
        ss_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        
        ws_frame = ttk.Frame(file_frame)
        ws_frame.pack(fill=tk.X, pady=(5, 5))
        
        ttk.Label(ws_frame, text="Nome Foglio di Lavoro:").pack(side=tk.LEFT)
        ws_entry = ttk.Entry(ws_frame, textvariable=self.worksheet_name_var)
        ws_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

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

        preview_button = ttk.Button(btn_frame, text="Anteprima Dati", command=self.show_data_preview, bootstyle="info")
        preview_button.grid(row=0, column=0, padx=(0, 5), sticky=tk.E)

        run_button = ttk.Button(btn_frame, text="Avvia Bot", command=self.start_bot_thread, bootstyle="success")
        run_button.grid(row=0, column=1, padx=(5, 0), sticky=tk.W)

        self.update_btn = ttk.Button(self.main_content_frame, text="Aggiorna Bot", command=self.update_bot, bootstyle="warning")
        self.update_btn.grid(row=3, column=0, pady=10, sticky=tk.EW)
        self.update_btn.grid_remove() 
        
        self.update_available.trace_add('write', self.handle_update_button_visibility)
    
    def handle_update_button_visibility(self, *args):
        """Mostra o nasconde il pulsante di aggiornamento in base allo stato."""
        if self.update_available.get():
            self.update_btn.grid()
        else:
            self.update_btn.grid_remove()

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

        # Crea un file .env di esempio
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
        timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, f"{timestamp} {message}\n")
        self.log_area.configure(state='disabled')
        self.log_area.see(tk.END)
    
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

    # --- Logica di Aggiornamento ---

    def check_for_updates(self):
        """
        Controlla se è disponibile un aggiornamento del bot.
        """
        self._log_message("Controllo aggiornamenti in corso...")
        github_repo_url = "https://api.github.com/repos/Mxttjaw/DataToSheets/releases/latest"
        
        try:
            response = requests.get(github_repo_url, timeout=10)
            response.raise_for_status() 
            latest_release = response.json()
            latest_version = latest_release.get("tag_name", "v0.0.0").lstrip('v')
            
            if latest_version > CURRENT_VERSION:
                self.update_available.set(True)
                self._log_message(f"Aggiornamento disponibile! Versione attuale: {CURRENT_VERSION}, ultima versione: {latest_version}.")
                messagebox.showinfo("Aggiornamento Disponibile", 
                    f"Una nuova versione del bot ({latest_version}) è disponibile per il download.\n\n"
                    "Clicca su 'Aggiorna Bot' per installare la nuova versione.")
            else:
                self.update_available.set(False)
                self._log_message("Il bot è già aggiornato all'ultima versione.")
        
        except requests.exceptions.RequestException as e:
            self._log_message(f"Errore durante il controllo degli aggiornamenti: {e}")
            self.update_available.set(False)


    def update_bot(self):
        """
        Inizializza il processo di aggiornamento.
        """
        self._log_message("Avvio del processo di aggiornamento...")
        
        download_url = self._get_download_url()
        if not download_url:
            messagebox.showerror("Errore Aggiornamento", "URL di download non trovato per il tuo sistema operativo.")
            return

        try:
            # Scarica la nuova versione
            self._log_message(f"Download in corso da: {download_url}")
            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()
            
            # Crea un nome temporaneo per il nuovo file
            temp_dir = tempfile.gettempdir()
            if platform.system() == "Windows":
                new_exe_name = "DataToSheets_new.exe"
            else:
                new_exe_name = "DataToSheets_new"
                
            new_exe_path = os.path.join(temp_dir, new_exe_name)
            
            with open(new_exe_path, "wb") as f:
                shutil.copyfileobj(response.raw, f)
            
            self._log_message("Nuova versione scaricata con successo.")
            
            # Rendi il file eseguibile (solo su Unix)
            if platform.system() != "Windows":
                os.chmod(new_exe_path, 0o755)
            
            # Avvia lo script di aggiornamento
            self.launch_updater_script(new_exe_path)
            self.master.destroy() 
        
        except Exception as e:
            self._log_message(f"Errore durante il download o il lancio dell'updater: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Non è stato possibile aggiornare il bot: {e}")



    def _get_latest_version(self):
        """
        Ottiene l'ultima versione disponibile dall'API di GitHub
        """
        try:
            api_url = "https://api.github.com/repos/Mxttjaw/DataToSheets/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            return response.json().get('tag_name', '').lstrip('v')
        except Exception as e:
            self._log_message(f"Errore nel recupero dell'ultima versione: {str(e)}")
            return None

    def _get_download_url(self):
        """
        Restituisce l'URL di download corretto per l'aggiornamento in base al sistema operativo.
        Formato URL: https://github.com/[user]/[repo]/releases/download/[tag]/[filename]
        """
        try:
            # Ottieni l'ultima versione disponibile
            latest_version = self._get_latest_version()  # Assicurati di avere questa funzione
            if not latest_version:
                self._log_message("Impossibile determinare l'ultima versione disponibile")
                return None

            # Costruisci la base dell'URL
            repo_url = "https://github.com/Mxttjaw/DataToSheets/releases/download"
            tag = latest_version.lstrip('v')  # Rimuove 'v' dal tag se presente

            # Determina il nome del file in base all'OS e all'architettura
            system = platform.system()
            machine = platform.machine().lower()

            if system == "Windows":
                filename = "DataToSheets-Windows.exe"
            elif system == "Darwin":
                filename = "DataToSheets-macOS"
            else:  # Linux e altri
                filename = "DataToSheets-Linux"

            download_url = f"{repo_url}/{latest_version}/{filename}"
            
            # Debug (puoi rimuoverlo in produzione)
            self._log_message(f"URL di download generato: {download_url}")
            self._log_message(f"Sistema: {system}, Architettura: {machine}")
            
            return download_url

        except Exception as e:
            self._log_message(f"Errore nella generazione dell'URL di download: {str(e)}")
            return None

    def launch_updater_script(self, new_exe_path):
        """
        Crea e avvia lo script temporaneo che si occuperà di sostituire il file.
        """
        current_exe_path = sys.executable
        temp_dir = tempfile.gettempdir()
        
        if platform.system() == "Windows":
            updater_script = os.path.join(temp_dir, "DataToSheets_updater.bat")
            script_content = f"""
            @echo off
            timeout /t 3 /nobreak >nul
            taskkill /F /PID {os.getpid()} >nul 2>&1
            move /Y "{new_exe_path}" "{current_exe_path}" >nul
            start "" "{current_exe_path}"
            del "%~f0"
            """
        else:  # macOS e Linux
            updater_script = os.path.join(temp_dir, "DataToSheets_updater.sh")
            script_content = f"""#!/bin/bash
            sleep 3
            kill -9 {os.getpid()} >/dev/null 2>&1
            mv -f "{new_exe_path}" "{current_exe_path}" >/dev/null 2>&1
            chmod +x "{current_exe_path}"
            nohup "{current_exe_path}" >/dev/null 2>&1 &
            rm -f "$0"
            """
        
        try:
            with open(updater_script, "w") as f:
                f.write(script_content)
            
            if platform.system() != "Windows":
                os.chmod(updater_script, 0o755)
            
            # Avvia lo script di aggiornamento
            if platform.system() == "Windows":
                subprocess.Popen([updater_script], creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                subprocess.Popen([updater_script], start_new_session=True)
        except Exception as e:
            self._log_message(f"Errore durante la creazione dello script di aggiornamento: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Errore durante l'avvio del processo di aggiornamento: {e}")

    # --- Metodi per la gestione dei file e dell'interfaccia utente ---
    
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
        
        try:
            shutil.copyfile(source_path, destination_path)
            self._log_message(f"File credenziali copiato e rinominato con successo in: {destination_path}")
            messagebox.showinfo("Successo", f"File 'credentials.json' configurato correttamente! Ora puoi condividere il tuo foglio di calcolo con l'indirizzo email di servizio per completare la configurazione. Lo trovi in: {self.user_data_path}")
        except Exception as e:
            self._log_message(f"Errore durante la gestione del file credenziali: {e}")
            messagebox.showerror("Errore", f"Errore durante la gestione del file: {e}. Controlla i permessi della cartella.")
    
    def select_file(self):
        """Apre una finestra di dialogo per selezionare il file di testo e lo salva."""
        file_path = filedialog.askopenfilename(
            title="Seleziona il file dei dati",
            filetypes=[("File di testo", "*.txt")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self._log_message(f"File selezionato: {file_path}")
            self._save_config_value('PATHS', 'dati_emails_path', file_path)

    def show_data_preview(self):
        """Mostra un'anteprima dei dati estratti dal file in una nuova finestra."""
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Attenzione", "Seleziona un file valido prima di visualizzare l'anteprima.")
            return

        dati_emails = self.processa_file_testo(file_path)
        if not dati_emails:
            messagebox.showinfo("Nessun Dato", "Nessun dato valido trovato nel file per l'anteprima.")
            return

        preview_window = tk.Toplevel(self.master)
        preview_window.title("Anteprima Dati")
        preview_window.geometry("500x300")
        preview_window.transient(self.master)

        text_area = tk.Text(preview_window, wrap=tk.WORD)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for riga in dati_emails:
            text_area.insert(tk.END, " | ".join(riga) + "\n")
            text_area.insert(tk.END, "-"*50 + "\n")

        text_area.configure(state='disabled')

    def open_options_window(self):
        """Apre una nuova finestra per le opzioni con il pulsante 'Applica Modifiche'."""
        opzioni_window = tk.Toplevel(self.master)
        opzioni_window.title("Opzioni")
        opzioni_window.geometry("300x150")
        opzioni_window.transient(self.master)
        opzioni_window.grab_set()

        tema_originale = self.master.style.theme_use()
        svuota_originale_state = self.svuota_file_var.get()
        
        opzioni_frame = ttk.Frame(opzioni_window, padding=15)
        opzioni_frame.pack(fill=BOTH, expand=True)

        svuota_check = ttk.Checkbutton(opzioni_frame,
                                        text="Svuota file dati dopo l'invio",
                                        variable=self.svuota_file_var,
                                        bootstyle="round-toggle")
        svuota_check.pack(pady=(0, 10))

        tema_frame = ttk.Frame(opzioni_frame)
        tema_frame.pack(fill=tk.X)

        tema_label = ttk.Label(tema_frame, text="Scegli Tema:")
        tema_label.pack(side=tk.LEFT, padx=(0, 10))
        
        temi = self.master.style.theme_names()
        tema_combo = ttk.Combobox(tema_frame, values=temi, state="readonly")
        tema_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tema_combo.set(tema_originale)
        
        def on_tema_scelto(event):
            """Cambia il tema dell'applicazione in tempo reale."""
            tema_selezionato = tema_combo.get()
            self.master.style.theme_use(tema_selezionato)

        tema_combo.bind("<<ComboboxSelected>>", on_tema_scelto)
        
        def apply_options():
            """Applica e salva le modifiche delle opzioni."""
            tema_selezionato = tema_combo.get()
            self._save_config_value('SETTINGS', 'theme', tema_selezionato)
            self._save_config_value('SETTINGS', 'svuota_file', self.svuota_file_var.get())
            messagebox.showinfo("Opzioni salvate", "Le modifiche sono state salvate con successo!")
            opzioni_window.destroy()

        def reset_options():
            """Annulla le modifiche e chiude la finestra."""
            self.master.style.theme_use(tema_originale)
            self.svuota_file_var.set(svuota_originale_state)
            opzioni_window.destroy()

        btn_frame = ttk.Frame(opzioni_window)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        apply_btn = ttk.Button(btn_frame, text="Applica", command=apply_options, bootstyle="success")
        apply_btn.grid(row=0, column=0, padx=(0, 5))
        
        cancel_btn = ttk.Button(btn_frame, text="Annulla", command=reset_options, bootstyle="danger")
        cancel_btn.grid(row=0, column=1, padx=(5, 0))

    def run_bot(self):
        """Esegue la logica principale del bot."""
        self._log_message("Avvio del bot...")
        file_path = self.file_path_var.get()

        if not file_path or not os.path.exists(file_path):
            self._log_message("Errore: Percorso del file non valido. Seleziona un file e riprova.")
            messagebox.showerror("Errore", "Seleziona un file valido prima di avviare il bot.")
            return

        spreadsheet_name = self.spreadsheet_name_var.get()
        worksheet_name = self.worksheet_name_var.get()
        svuota_file = self.svuota_file_var.get()

        if not spreadsheet_name:
            self._log_message("Errore: Il nome del foglio di calcolo non può essere vuoto.")
            messagebox.showerror("Errore", "Inserisci il nome del foglio di calcolo.")
            return

        if not worksheet_name:
            self._log_message("Errore: Il nome del foglio di lavoro non può essere vuoto.")
            messagebox.showerror("Errore", "Inserisci il nome del foglio di lavoro.")
            return

        dati_emails = self.processa_file_testo(file_path)
        
        if dati_emails:
            self._log_message(f"Trovati {len(dati_emails)} nuovi record da scrivere.")
            self.scrivi_su_sheets(dati_emails, spreadsheet_name, worksheet_name)
            
            if svuota_file:
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write('')
                    self._log_message(f"File '{os.path.basename(file_path)}' svuotato con successo.")
                except Exception as e:
                    self._log_message(f"Errore durante lo svuotamento del file: {e}")
            
        else:
            self._log_message("Nessun nuovo dato da scrivere.")
        
        self._log_message("Bot completato.")

    # --- Logica del bot originale, ora come metodi della classe ---

    def get_google_sheet_client(self):
        """Si connette a Google Sheets usando le credenziali di servizio."""
        try:
            creds_file_name = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
            creds_path = os.path.join(self.user_data_path, creds_file_name)
            
            if not os.path.exists(creds_path):
                self._log_message(f"Errore: File '{creds_file_name}' non trovato nella cartella di configurazione. Assicurati di averlo caricato tramite il tutorial.")
                messagebox.showerror("Errore Credenziali", f"File credenziali non trovato. Per favore, segui il tutorial per configurare il file '{creds_file_name}' in: {self.user_data_path}")
                return None
            
            creds = service_account.Credentials.from_service_account_file(
                creds_path,
                scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            )
            client = gspread.authorize(creds)
            self._log_message("Autenticazione con Google Sheets riuscita.")
            return client
        except Exception as e:
            self._log_message(f"Errore di autenticazione con Google Sheets: {e}")
            messagebox.showerror("Errore di Autenticazione", f"Verifica che il file delle credenziali sia corretto e che le API siano abilitate: {e}")
            return None

    def scrivi_su_sheets(self, dati_emails, spreadsheet_name, worksheet_name):
        """Scrive più righe di dati nel foglio di Google Sheets specificato."""
        client = self.get_google_sheet_client()
        if not client:
            return

        self._log_message(f"Tentativo di scrittura nel foglio di calcolo: '{spreadsheet_name}'")
        self._log_message(f"Tentativo di scrittura nel foglio di lavoro: '{worksheet_name}'")

        try:
            sheet = client.open(spreadsheet_name).worksheet(worksheet_name)
        except WorksheetNotFound:
            self._log_message(f"Il foglio di lavoro '{worksheet_name}' non è stato trovato. Creazione in corso...")
            try:
                workbook = client.open(spreadsheet_name)
                sheet = workbook.add_worksheet(title=worksheet_name, rows="100", cols="20")
                self._log_message(f"Foglio '{worksheet_name}' creato con successo.")
            except Exception as e:
                self._log_message(f"Errore durante la creazione del foglio '{worksheet_name}': {e}")
                messagebox.showerror("Errore Creazione Foglio", f"Non è stato possibile creare il foglio di lavoro. Verifica che il nome del foglio di calcolo principale sia corretto e che l'account di servizio abbia i permessi di 'Editor': {e}")
                return
        except gspread.exceptions.SpreadsheetNotFound:
            self._log_message(f"Errore: Il foglio di calcolo '{spreadsheet_name}' non è stato trovato o l'account di servizio non ha i permessi necessari.")
            messagebox.showerror("Errore Foglio di Calcolo", f"Il foglio di calcolo '{spreadsheet_name}' non è stato trovato. Assicurati che il nome sia corretto e che l'account di servizio abbia i permessi di 'Editor'.")
            return
        except APIError as e:
            self._log_message(f"Errore API durante l'accesso a Google Sheets: {e}")
            messagebox.showerror("Errore API", f"Si è verificato un errore con l'API di Google Sheets. Controlla i permessi e la connessione a Internet. Dettagli: {e}")
            return

        try:
            if dati_emails:
                sheet.append_rows(dati_emails)
                self._log_message(f"Dati scritti con successo: {len(dati_emails)} righe nel foglio '{worksheet_name}'.")
            else:
                self._log_message(f"Nessun dato da scrivere nel foglio '{worksheet_name}'.")

        except Exception as e:
            self._log_message(f"Errore durante la scrittura su Google Sheets: {e}")
            messagebox.showerror("Errore Scrittura Dati", f"Si è verificato un errore durante la scrittura dei dati. Dettagli: {e}")
            
    def pulisci_testo(self, testo):
        """Rimuove caratteri indesiderati e spazi extra."""
        if testo:
            # Sostituisce più spazi con uno solo, rimuove spazi all'inizio e alla fine
            return re.sub(r'\s+', ' ', testo).strip()
        return ""

    def estrai_dati_da_testo(self, blocco_testo):
        """Estrae i dati da un blocco di testo formattato in righe separate."""
        dati = {
            "Nome": "",
            "Cognome": "",
            "Eta": "",
            "Occupazione": "",
            "Email Mittente": "",
            "Numero di Telefono": "",
            "Richiesta": "",
        }
        
        # Dividi il blocco di testo in righe
        righe = blocco_testo.strip().split('\n')
        
        if not righe:
            return None

        # Estrazione dati
        if len(righe) > 0:
            # Riga 1: Nome, Cognome e Età
            riga1 = self.pulisci_testo(righe[0])
            parti_riga1 = riga1.split()
            if len(parti_riga1) >= 3 and parti_riga1[-1].isdigit():
                dati["Eta"] = parti_riga1[-1]
                dati["Nome"] = " ".join(parti_riga1[:-2])
                dati["Cognome"] = parti_riga1[-2]
            elif len(parti_riga1) >= 2:
                dati["Nome"] = " ".join(parti_riga1[:-1])
                dati["Cognome"] = parti_riga1[-1]

        if len(righe) > 1:
            # Riga 2: Occupazione
            dati["Occupazione"] = self.pulisci_testo(righe[1])
            
        if len(righe) > 2:
            # Riga 3: Email e Numero di Telefono
            riga3 = self.pulisci_testo(righe[2])
            match_email = re.search(r'[\w\.-]+@[\w\.-]+', riga3)
            if match_email:
                dati["Email Mittente"] = match_email.group(0)
            
            match_tel = re.search(r'\+?\d[\d\s-]{7,}\d', riga3)
            if match_tel:
                dati["Numero di Telefono"] = match_tel.group(0)

        if len(righe) > 3:
            # Riga 4: Richiesta
            richiesta = " ".join(righe[3:]).strip()
            dati["Richiesta"] = richiesta if richiesta else "Nessuna richiesta specificata"

        return [
            dati["Nome"],
            dati["Cognome"],
            dati["Eta"],
            dati["Occupazione"],
            dati["Email Mittente"],
            dati["Numero di Telefono"],
            dati["Richiesta"],
            datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        ]

    def processa_file_testo(self, file_path):
        """
        Legge un file di testo, lo divide in blocchi basati su una riga vuota
        e estrae i dati da ogni blocco.
        """
        dati_emails = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                contenuto = f.read()
            
            # Divide il contenuto in blocchi basati su righe vuote o sequenze di trattini
            blocchi = re.split(r'\n{2,}|\n-{5,}\n', contenuto.strip())
            
            for blocco in blocchi:
                if blocco.strip():
                    dati = self.estrai_dati_da_testo(blocco)
                    if dati and dati[4]: # Controllo che l'email non sia vuota
                        dati_emails.append(dati)
                    else:
                        self._log_message(f"Blocco ignorato, dati insufficienti: {blocco.strip()[:50]}...")
                        
            return dati_emails

        except FileNotFoundError:
            self._log_message(f"Errore: File '{file_path}' non trovato.")
            return []
        except Exception as e:
            self._log_message(f"Errore durante la lettura del file: {e}")
            return []


# --- Punto di Ingresso dell'applicazione ---
if __name__ == "__main__":
    try:
        app = ttk.Window(themename="lumen")
        BotApp(app)
        app.mainloop()
    except Exception as e:
        import traceback
        print("Errore critico all'avvio dell'applicazione:")
        traceback.print_exc()