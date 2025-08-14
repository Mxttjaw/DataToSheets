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

CURRENT_VERSION = "1.0.2"

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
        self.headers_set = False # Variabile di stato per la gestione delle intestazioni

        # Crea un'istanza della classe Updater e passa il metodo di log
        self.updater = Updater(current_version=CURRENT_VERSION, log_callback=self._log_message)

        self.create_gui_elements()
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
        # Sostituisce la vecchia logica con la chiamata all'Updater
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

        self.update_btn = ttk.Button(self.main_content_frame, text="Aggiorna Bot", command=lambda: Thread(target=self.updater.update_app, daemon=True).start(), bootstyle="warning")
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
        Legge un file di testo, lo divide in blocchi basati su una riga vuota
        e estrae i dati dal formato specifico dell'utente.
        """
        dati_emails = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                contenuto = f.read()
            
            # Divide il contenuto in blocchi basati su righe vuote
            blocchi = re.split(r'\n{2,}', contenuto.strip())
            
            for blocco in blocchi:
                if blocco.strip():
                    righe = [line.strip() for line in blocco.split('\n') if line.strip()]
                    
                    # Estrae i dati dalla prima riga
                    riga_uno = righe[0].split()
                    nome = riga_uno[0]
                    cognome = riga_uno[1]
                    eta = riga_uno[2]
                    occupazione = riga_uno[3]

                    # Estrae l'email dalla seconda riga e il telefono dalla terza
                    email = righe[1]
                    telefono = righe[2]
                    
                    dati = [nome, cognome, eta, occupazione, email, telefono]
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

    def invia_a_google_sheets(self, dati_emails):
        """Invia i dati a Google Sheets e restituisce il foglio di lavoro aggiornato."""
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
            sh = gc.open(spreadsheet_name)
        except gspread.SpreadsheetNotFound:
            self._log_message(f"Creazione del foglio di calcolo '{spreadsheet_name}'...")
            try:
                sh = gc.create(spreadsheet_name)
                # Assicurati di condividere il foglio con l'email del servizio
                sh.share(os.getenv("GSPREAD_EMAIL"), perm_type='user', role='writer')
            except Exception as e:
                self._log_message(f"Errore nella creazione del foglio di calcolo: {e}")
                messagebox.showerror("Errore", f"Impossibile creare il foglio di calcolo: {e}")
                return

        try:
            worksheet = sh.worksheet(worksheet_name)
        except WorksheetNotFound:
            self._log_message(f"Creazione del foglio di lavoro '{worksheet_name}'...")
            worksheet = sh.add_worksheet(title=worksheet_name, rows=100, cols=20)
            # Aggiunge le intestazioni
            worksheet.append_row(["Nome", "Cognome", "Età", "Occupazione", "Email", "Numero di Telefono"])
        
        self._log_message(f"Apertura del foglio di lavoro '{worksheet_name}'...")
        
        # Trova la prima riga vuota per iniziare a scrivere i dati
        next_row = len(worksheet.get_all_values()) + 1
        
        self._log_message(f"Invio di {len(dati_emails)} righe a Google Sheets a partire dalla riga {next_row}...")
        
        cell_list = []
        for i, riga in enumerate(dati_emails):
            # Mappa i dati alle colonne richieste
            # Nome in colonna B, Cognome in C, Età in D, ecc.
            # L'indice 1 corrisponde alla colonna B in gspread (che è 1-based)
            cell_list.append(gspread.Cell(next_row + i, 2, riga[0])) # Nome
            cell_list.append(gspread.Cell(next_row + i, 3, riga[1])) # Cognome
            cell_list.append(gspread.Cell(next_row + i, 4, riga[2])) # Età
            cell_list.append(gspread.Cell(next_row + i, 5, riga[3])) # Occupazione
            cell_list.append(gspread.Cell(next_row + i, 6, riga[4])) # Email
            cell_list.append(gspread.Cell(next_row + i, 7, riga[5])) # Telefono
            cell_list.append(gspread.Cell(next_row + i, 8, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))) # Data Inserimento

        worksheet.update_cells(cell_list)
        self._log_message("Dati inviati con successo!")
        
        return worksheet

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

