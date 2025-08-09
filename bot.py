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
from gspread.exceptions import WorksheetNotFound
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

load_dotenv()

CONFIG_FILE = 'config.ini'

CURRENT_VERSION = "1.0.0"


class BotApp(ttk.Frame):
    """
    Classe principale dell'applicazione, gestisce l'intera interfaccia utente
    e la logica del bot.
    """

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(fill=BOTH, expand=True)

        self.file_path_var = tk.StringVar()
        self.svuota_file_var = tk.BooleanVar()
        self.spreadsheet_name_var = tk.StringVar()
        self.worksheet_name_var = tk.StringVar()
        self.tutorial_state_var = tk.BooleanVar()
        self.update_available = tk.BooleanVar(value=False)
        self.running_thread = None

        self.tutorial_state_var.set(self._load_config_boolean('SETTINGS', 'tutorial_shown', False))
        
        self.top_menu_frame = ttk.Frame(self, height=30)
        self.top_menu_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        
        self.main_content_frame = ttk.Frame(self)
        self.main_content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self.main_content_frame.grid_rowconfigure(1, weight=1)
        self.main_content_frame.grid_columnconfigure(0, weight=1)

        self.create_gui_elements()
        
        self.load_initial_configuration()

        if self._load_config_boolean('SETTINGS', 'first_run', True):
            self.master.after(200, self.show_tutorial_window)

        Thread(target=self.check_for_updates, daemon=True).start()

    # --- Metodi per la gestione della GUI ---

    def create_gui_elements(self):
        """Costruisce tutti i widget dell'interfaccia utente."""
        self.master.title(f"Bot Email - Gestione Dati (v{CURRENT_VERSION})")
        self.master.geometry("600x400")
        self.master.resizable(True, True)
        
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

    def load_initial_configuration(self):
        """Carica il percorso del file e le opzioni salvate all'avvio."""
        self.file_path_var.set(self._load_config_string('PATHS', 'dati_emails_path'))
        self.svuota_file_var.set(self._load_config_boolean('SETTINGS', 'svuota_file', False))
        self.spreadsheet_name_var.set(self._load_config_string('SHEETS', 'spreadsheet_name', os.getenv("SPREADSHEET_NAME")))
        
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
        config.read(CONFIG_FILE)
        return config

    def _save_config(self, config):
        """Salva la configurazione corrente in un file."""
        try:
            with open(CONFIG_FILE, 'w') as configfile:
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
        github_repo_url = "https://api.github.com/repos/tuo_utente/tuo_repo_bot/releases/latest"
        
        try:
            response = requests.get(github_repo_url, timeout=10)
            response.raise_for_status() 
            latest_release = response.json()
            latest_version = latest_release.get("tag_name", "v0.0.0").lstrip('v')
            
            if latest_version > CURRENT_VERSION:
                self.update_available.set(True)
                self._log_message(f"Aggiornamento disponibile! Versione attuale: {CURRENT_VERSION}, ultima versione: {latest_version}.")
                messagebox.showinfo("Aggiornamento Disponibile", f"Una nuova versione del bot ({latest_version}) è disponibile per il download.")
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
            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()
            
            new_exe_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "new_bot")
            with open(new_exe_path, "wb") as f:
                shutil.copyfileobj(response.raw, f)
            
            self._log_message("Nuova versione scaricata con successo.")
            
            self.launch_updater_script(new_exe_path)
            self.master.destroy() 
        
        except Exception as e:
            self._log_message(f"Errore durante il download o il lancio dell'updater: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Non è stato possibile aggiornare il bot: {e}")

    def _get_download_url(self):
        """Restituisce l'URL di download corretto in base al sistema operativo."""
        github_release_url = "https://github.com/tuo_utente/tuo_repo_bot/releases/download/v1.1.0"
        
        system = platform.system()
        if system == "Windows":
            return f"{github_release_url}/tuo_bot.exe"
        elif system == "Darwin":
            return f"{github_release_url}/tuo_bot_macos"
        elif system == "Linux":
            return f"{github_release_url}/tuo_bot_linux"
        else:
            return None

    def launch_updater_script(self, new_exe_path):
        """
        Crea e avvia lo script temporaneo che si occuperà di sostituire il file.
        """
        current_exe_path = sys.executable
        if platform.system() == "Windows":
            updater_script = "updater.bat"
            script_content = f"""
            @echo off
            timeout /t 5 /nobreak
            copy "{new_exe_path}" "{current_exe_path}"
            del "{new_exe_path}"
            start "" "{current_exe_path}"
            exit
            """
            with open(updater_script, "w") as f:
                f.write(script_content)
            subprocess.Popen([updater_script], creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            updater_script = "updater.sh"
            script_content = f"""
            #!/bin/bash
            sleep 5
            mv "{new_exe_path}" "{current_exe_path}"
            chmod +x "{current_exe_path}"
            nohup "{current_exe_path}" &
            exit
            """
            with open(updater_script, "w") as f:
                f.write(script_content)
            os.chmod(updater_script, 0o755) 
            subprocess.Popen(["bash", updater_script])

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
        tutorial_text.insert(tk.END, "Segui questi semplici passi per collegare il bot a Google Sheets e ai tuoi file dati.\n\n")

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
        tutorial_text.insert(tk.END, "Ora clicca sul pulsante 'Seleziona File Credenziali' qui sotto per caricare il file JSON appena scaricato. Il bot lo rinominerà in 'credentials.json' e lo posizionerà nella cartella corretta.\n\n")

        tutorial_text.insert(tk.END, "Passo 5: Condividere il foglio di calcolo\n", ("bold"))
        tutorial_text.insert(tk.END, "Apri il tuo foglio di Google Sheets e clicca su 'Condividi'. Incolla l'indirizzo email dell'account di servizio (lo trovi nel file JSON) e concedi il permesso 'Editor'.\n")
        tutorial_text.configure(state='disabled')
        
        bottom_frame = ttk.Frame(tutorial_window)
        bottom_frame.pack(pady=(0, 10))

        tutorial_check = ttk.Checkbutton(bottom_frame,
                                          text="Non mostrare più all'avvio",
                                          variable=self.tutorial_state_var,
                                          bootstyle="round-toggle")
        tutorial_check.pack(side=tk.LEFT, padx=(0, 20))
        
        creds_btn = ttk.Button(bottom_frame, text="Seleziona File Credenziali", command=self.handle_credentials_file, bootstyle="success")
        creds_btn.pack(side=tk.LEFT, padx=(0, 20))
        
        close_btn = ttk.Button(bottom_frame, text="Chiudi", command=lambda: self._close_tutorial_window(tutorial_window), bootstyle="primary")
        close_btn.pack(side=tk.LEFT)

    def _close_tutorial_window(self, window):
        """Salva lo stato del tutorial e chiude la finestra."""
        self._save_config_value('SETTINGS', 'tutorial_shown', self.tutorial_state_var.get())
        if self._load_config_boolean('SETTINGS', 'first_run', True):
            self._save_config_value('SETTINGS', 'first_run', False)
        window.destroy()
        
    def handle_credentials_file(self):
        """Permette all'utente di selezionare il file JSON e lo sposta/rinomina."""
        source_path = filedialog.askopenfilename(
            title="Seleziona il file 'credentials.json' appena scaricato",
            filetypes=[("File JSON", "*.json")]
        )
        if not source_path:
            return

        destination_dir = os.path.dirname(os.path.abspath(__file__))
        destination_path = os.path.join(destination_dir, "credentials.json")
        
        try:
            shutil.copyfile(source_path, destination_path)
            self._log_message(f"File credenziali copiato e rinominato con successo in: {destination_path}")
            messagebox.showinfo("Successo", "File 'credentials.json' configurato correttamente! Ora puoi condividere il tuo foglio di calcolo con l'indirizzo email di servizio per completare la configurazione.")
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

        dati_emails = self.process_text_file(file_path)
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
        
        applica_btn = ttk.Button(opzioni_window, text="Applica Modifiche",
                                 command=lambda: self._save_options_and_close(tema_combo.get(), opzioni_window),
                                 bootstyle="success")
        applica_btn.pack(pady=10)
        
        opzioni_window.protocol("WM_DELETE_WINDOW", lambda: self._revert_options_and_close(tema_originale, svuota_originale_state, opzioni_window))

    def _save_options_and_close(self, tema_selezionato, window):
        """Salva tutte le opzioni e chiude la finestra."""
        self._save_config_value('SETTINGS', 'svuota_file', self.svuota_file_var.get())
        self._save_config_value('SETTINGS', 'theme', tema_selezionato)
        self._save_config_value('SHEETS', 'spreadsheet_name', self.spreadsheet_name_var.get())
        self._save_config_value('SHEETS', 'worksheet_name', self.worksheet_name_var.get())
        
        self._log_message("Modifiche alle opzioni applicate e salvate.")
        window.destroy()

    def _revert_options_and_close(self, tema_originale, svuota_originale_state, window):
        """Ripristina le opzioni allo stato iniziale e chiude la finestra."""
        self.master.style.theme_use(tema_originale)
        self.svuota_file_var.set(svuota_originale_state)
        
        self._log_message("Modifiche alle opzioni annullate.")
        window.destroy()
    
    # --- Logica Principale del Bot ---

    def run_bot(self):
        """Funzione principale per avviare la logica del bot."""
        file_path = self.file_path_var.get()
        spreadsheet_name = self.spreadsheet_name_var.get()
        worksheet_name = self.worksheet_name_var.get()
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Attenzione", "Percorso file non valido. Seleziona un file prima di avviare il bot.")
            return
        
        if not spreadsheet_name or not worksheet_name:
            messagebox.showwarning("Attenzione", "Inserisci il nome del foglio di calcolo e del foglio di lavoro.")
            return
        
        self._log_message("Avvio del bot...")
        
        dati_emails = self.process_text_file(file_path)
        
        if not dati_emails:
            self._log_message("Nessun nuovo dato da scrivere.")
            return

        self._create_progress_bar()
        try:
            self.write_to_sheets(dati_emails, spreadsheet_name, worksheet_name)
        finally:
            self.progress_bar.stop()
            self.progress_window.destroy()
    
    def process_text_file(self, file_path):
        """
        Legge il file di testo, estrae le email e altri dati.
        Ritorna una lista di liste con i dati estratti.
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            dati_estratti = []
            for line in lines:
                match = re.match(r'(.*@.*\..*)\s(.*?)$', line.strip())
                if match:
                    email, altri_dati = match.groups()
                    dati_estratti.append([email, altri_dati])
            
            return dati_estratti
        except Exception as e:
            self._log_message(f"Errore durante la lettura del file: {e}")
            messagebox.showerror("Errore File", f"Errore durante la lettura del file '{file_path}': {e}")
            return []

    def get_google_sheet_client(self):
        """
        Si connette a Google Sheets usando le credenziali di servizio.
        Il file `credentials.json` deve essere nella stessa cartella del bot.
        """
        try:
            creds_file_path = "credentials.json"
            if not os.path.exists(creds_file_path):
                self._log_message(f"Errore: File credenziali non trovato: '{creds_file_path}'. Esegui il tutorial per configurarlo.")
                return None
        
            creds = service_account.Credentials.from_service_account_file(
                creds_file_path,
                scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            )
            client = gspread.authorize(creds)
            self._log_message("Autenticazione con Google Sheets riuscita.")
            return client
        except Exception as e:
            self._log_message(f"Errore di autenticazione con Google Sheets: {e}")
            messagebox.showerror("Errore Credenziali", f"Errore di autenticazione con Google Sheets. Controlla il file 'credentials.json': {e}")
            return None

    def write_to_sheets(self, dati_emails, spreadsheet_name, worksheet_name):
        """
        Scrive i dati estratti in un foglio di Google Sheets.
        Se il foglio di lavoro non esiste, lo crea.
        """
        client = self.get_google_sheet_client()
        if not client:
            return

        try:
            spreadsheet = client.open(spreadsheet_name)
            self._log_message(f"Foglio di calcolo '{spreadsheet_name}' aperto con successo.")
        except gspread.exceptions.SpreadsheetNotFound:
            self._log_message(f"Errore: Foglio di calcolo '{spreadsheet_name}' non trovato.")
            messagebox.showerror("Errore", f"Foglio di calcolo '{spreadsheet_name}' non trovato. Controlla il nome o i permessi di condivisione.")
            return
        
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
            self._log_message(f"Foglio di lavoro '{worksheet_name}' trovato.")
        except WorksheetNotFound:
            self._log_message(f"Foglio di lavoro '{worksheet_name}' non trovato. Creazione in corso...")
            try:
                worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="100", cols="20")
                self._log_message(f"Foglio di lavoro '{worksheet_name}' creato con successo.")
            except Exception as e:
                self._log_message(f"Errore durante la creazione del foglio di lavoro: {e}")
                messagebox.showerror("Errore", f"Impossibile creare il foglio di lavoro '{worksheet_name}'.")
                return

        try:
            worksheet.append_rows(dati_emails)
            self._log_message(f"Dati scritti con successo nel foglio di lavoro '{worksheet_name}'.")
            
            if self.svuota_file_var.get():
                self._clear_data_file(self.file_path_var.get())
            
            messagebox.showinfo("Successo", "Dati inviati correttamente a Google Sheets!")
        except Exception as e:
            self._log_message(f"Errore durante la scrittura dei dati su Google Sheets: {e}")
            messagebox.showerror("Errore Scrittura", f"Errore durante la scrittura dei dati: {e}")
            
    def _create_progress_bar(self):
        """Crea e mostra una finestra con una progress bar."""
        self.progress_window = tk.Toplevel(self.master)
        self.progress_window.title("In lavorazione...")
        self.progress_window.geometry("300x100")
        self.progress_window.transient(self.master)
        self.progress_window.grab_set()

        frame = ttk.Frame(self.progress_window, padding=15)
        frame.pack(fill=BOTH, expand=True)

        label = ttk.Label(frame, text="Invio dati a Google Sheets...", font=("Helvetica", 10))
        label.pack(pady=(0, 10))

        self.progress_bar = ttk.Progressbar(frame, mode="indeterminate", bootstyle="success")
        self.progress_bar.pack(fill=tk.X, expand=True)
        self.progress_bar.start()

    def _clear_data_file(self, file_path):
        """Svuota il contenuto del file specificato."""
        if file_path and os.path.exists(file_path):
            try:
                with open(file_path, 'w') as f:
                    f.truncate(0)
                self._log_message(f"Contenuto del file '{file_path}' svuotato con successo.")
            except Exception as e:
                self._log_message(f"Errore durante lo svuotamento del file: {e}")
                messagebox.showerror("Errore Svuotamento", f"Errore durante lo svuotamento del file: {e}")

# --- Avvio dell'applicazione ---
# Esegue il codice solo quando lo script viene lanciato direttamente.
if __name__ == "__main__":
    app_root = ttk.Window(themename="lumen")
    app = BotApp(master=app_root)
    app.mainloop()

