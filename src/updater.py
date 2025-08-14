# python
import os
import sys
import time
import shutil
import platform
import tempfile
import subprocess
import requests
from threading import Thread
from tkinter import messagebox

class Updater:
    """
    Gestisce la logica di verifica e installazione degli aggiornamenti
    per l'applicazione, in modo compatibile con tutti i sistemi operativi.
    """
    def __init__(self, current_version, log_callback=None):
        self.current_version = current_version
        self.log = log_callback if log_callback else print
        self.github_repo_url = "https://api.github.com/repos/Mxttjaw/DataToSheets/releases"

    def _parse_version_string(self, version_tag):
        """Estrae la parte numerica da una stringa di versione come 'v1.0.1'."""
        return version_tag.lstrip('v')

    def check_for_updates(self, on_update_available=None):
        """
        Controlla la disponibilità di un aggiornamento su GitHub.
        on_update_available: callback da chiamare se un aggiornamento è disponibile.
        """
        self.log("Controllo aggiornamenti in corso...")
        try:
            response = requests.get(self.github_repo_url, timeout=10)
            response.raise_for_status()
            all_releases = response.json()
            
            if not all_releases:
                self.log("Nessuna release trovata sul repository GitHub.")
                return

            releases = [r for r in all_releases if "test" not in r.get("tag_name", "")]
            if not releases:
                self.log("Nessuna release stabile trovata.")
                return

            releases.sort(key=lambda r: self._parse_version_string(r.get("tag_name", "v0.0.0")), reverse=True)
            latest_release = releases[0]
            latest_tag = latest_release.get("tag_name", "v0.0.0")
            latest_version = self._parse_version_string(latest_tag)
            current_version = self._parse_version_string(self.current_version)

            if latest_version > current_version:
                self.log(f"Aggiornamento disponibile! Versione attuale: {current_version}, ultima versione: {latest_version}.")
                if on_update_available:
                    on_update_available()
            else:
                self.log("Il bot è già aggiornato all'ultima versione.")
        except requests.exceptions.RequestException as e:
            self.log(f"Errore durante la verifica aggiornamenti: {e}")
        except Exception as e:
            self.log(f"Errore imprevisto durante la verifica aggiornamenti: {e}")

    def update_app(self):
        """
        Scarica e installa l'aggiornamento.
        """
        self.log("Avvio del processo di aggiornamento...")
        try:
            response = requests.get(self.github_repo_url, timeout=10)
            response.raise_for_status()
            all_releases = response.json()
            releases = [r for r in all_releases if "test" not in r.get("tag_name", "")]
            releases.sort(key=lambda r: self._parse_version_string(r.get("tag_name", "v0.0.0")), reverse=True)
            latest_release = releases[0]

            download_url = None
            for asset in latest_release.get("assets", []):
                if asset.get("name") == os.path.basename(sys.argv[0]): 
                    download_url = asset.get("browser_download_url")
                    break

            if not download_url:
                self.log("Errore: Impossibile trovare il file di aggiornamento nel rilascio.")
                messagebox.showerror("Errore Aggiornamento", "Impossibile trovare il file di aggiornamento.")
                return
            
            self.log(f"Download dell'aggiornamento da: {download_url}")
            response = requests.get(download_url)
            response.raise_for_status()

            new_file_path = os.path.join(tempfile.gettempdir(), f"bot_new_{int(time.time())}.py")
            with open(new_file_path, 'wb') as f:
                f.write(response.content)
            
            self.log(f"Aggiornamento scaricato in: {new_file_path}")
            current_script_path = os.path.abspath(sys.argv[0])
            self.log(f"Percorso del file corrente: {current_script_path}")
            
            shutil.move(new_file_path, current_script_path)
            self.log("Aggiornamento completato. Il file è stato sostituito.")

            self._restart_app(current_script_path)

        except requests.exceptions.RequestException as e:
            self.log(f"Errore durante il download dell'aggiornamento: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Errore durante il download: {e}")
        except Exception as e:
            self.log(f"Errore imprevisto durante l'aggiornamento: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Errore imprevisto: {e}")

    def _restart_app(self, script_path):
        """
        Riavvia l'applicazione in modo robusto.
        """
        python_executable = sys.executable
        self.log(f"Riavvio dell'applicazione con: {python_executable} {script_path}")
        
        if platform.system() == 'Windows':
            subprocess.Popen([python_executable, script_path])
        else:
            subprocess.Popen([python_executable, script_path], start_new_session=True)
        
        sys.exit(0)
