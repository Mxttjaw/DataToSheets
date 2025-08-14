# python
import os
import sys
import time
import platform
import tempfile
import subprocess
import requests
from tkinter import messagebox

class Updater:
    """
    Updater robusto multi-OS progettato per:
    - funzionare con PyInstaller (onefile) e con script interpretati
    - lanciare un helper che aspetta il PID corretto, effettua il mv e riavvia
    - evitare blocchi se update chiamato da thread (usa os._exit per terminare)
    """
    def __init__(self, current_version, log_callback=None):
        self.current_version = current_version
        self.log = log_callback if log_callback else print
        self.github_repo_url = "https://api.github.com/repos/Mxttjaw/DataToSheets/releases"

        # Costruiamo più "candidati" per il path dell'eseguibile e scegliamo quello che esiste.
        cand_exec = None
        cand_argv = None
        try:
            cand_exec = os.path.realpath(sys.executable)
        except Exception:
            cand_exec = None

        try:
            if sys.argv and len(sys.argv) > 0:
                arg0 = sys.argv[0]
                if os.path.isabs(arg0):
                    cand_argv = os.path.realpath(arg0)
                else:
                    # risolvi arg0 rispetto alla current working dir (anche se potrebbe essere stata cambiata)
                    cand_argv = os.path.realpath(os.path.join(os.getcwd(), arg0))
        except Exception:
            cand_argv = None

        # Altri possibili fallback: basename nell'attuale cwd
        cand_basename = None
        try:
            base = os.path.basename(sys.argv[0]) if sys.argv and len(sys.argv) > 0 else None
            if base:
                cand_basename = os.path.realpath(os.path.join(os.getcwd(), base))
        except Exception:
            cand_basename = None

        candidates = []
        for c in (cand_argv, cand_exec, cand_basename):
            if c and c not in candidates:
                candidates.append(c)

        # Scegli il primo candidato che esiste realmente sul FS
        chosen = None
        for c in candidates:
            try:
                if c and os.path.exists(c):
                    chosen = c
                    break
            except Exception:
                continue

        # Se non troviamo nulla, usiamo cand_exec (o un fallback)
        if not chosen:
            chosen = cand_exec or cand_argv or (os.path.abspath(sys.argv[0]) if sys.argv and len(sys.argv) > 0 else sys.executable)

        # Normalizziamo il path e lo salviamo
        try:
            chosen = os.path.realpath(chosen)
        except Exception:
            pass

        self.current_executable_path = chosen
        self.log(f"Updater: candidates={candidates}, chosen_current_executable_path={self.current_executable_path}")


    def _parse_version_string(self, version_tag):
        try:
            return tuple(map(int, version_tag.lstrip('v').split('.')))
        except (ValueError, AttributeError):
            self.log(f"Avviso: Impossibile parsare la versione '{version_tag}'.")
            return (0, 0, 0)

    def check_for_updates(self, on_update_available=None):
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
                self.log(f"Aggiornamento disponibile! Versione attuale: {self.current_version}, ultima versione: {latest_tag}.")
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
        Scarica l'asset dal release più recente e avvia lo script helper.
        Nota: questa funzione può essere chiamata anche da thread; terminare
        il processo principale viene effettuato con os._exit(0) dopo aver avviato l'helper.
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
                asset_name = asset.get("name", "")
                if platform.system() == 'Windows' and asset_name.endswith(".exe"):
                    download_url = asset.get("browser_download_url")
                    break
                elif platform.system() == 'Linux' and 'Linux' in asset_name:
                    download_url = asset.get("browser_download_url")
                    break
                elif platform.system() == 'Darwin' and 'macOS' in asset_name:
                    download_url = asset.get("browser_download_url")
                    break

            if not download_url:
                self.log("Errore: Impossibile trovare il file di aggiornamento nel rilascio.")
                messagebox.showerror("Errore Aggiornamento", "Impossibile trovare il file di aggiornamento.")
                return

            self.log(f"Download dell'aggiornamento da: {download_url}")
            response = requests.get(download_url, timeout=60)
            response.raise_for_status()

            new_executable_name = os.path.basename(self.current_executable_path) + "_new"
            new_executable_path = os.path.join(tempfile.gettempdir(), new_executable_name)

            with open(new_executable_path, 'wb') as f:
                f.write(response.content)

            # Assicuriamoci che sia eseguibile (utile su Unix)
            try:
                os.chmod(new_executable_path, 0o755)
            except Exception:
                pass

            self.log(f"Aggiornamento scaricato in: {new_executable_path}")

            # Lancia helper (passiamo PID, OLD_PATH, NEW_PATH, e gli eventuali argomenti)
            self._launch_update_script(new_executable_path)

        except requests.exceptions.RequestException as e:
            self.log(f"Errore durante il download dell'aggiornamento: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Errore durante il download: {e}")
        except Exception as e:
            self.log(f"Errore imprevisto durante l'aggiornamento: {e}")
            messagebox.showerror("Errore Aggiornamento", f"Errore imprevisto: {e}")

    def _launch_update_script(self, new_executable_path, wait_timeout=60):
        """
        Crea e avvia lo script helper. Dopo l'avvio dello helper termina
        immediatamente il processo principale (os._exit) per assicurare che il file
        sia liberato e possa essere sovrascritto.
        """
        current_system = platform.system()
        current_pid = os.getpid()
        ts = int(time.time())

        if current_system in ['Linux', 'Darwin']:
            self.log("Creazione dello script helper per sistemi Unix.")
            logpath = os.path.join(tempfile.gettempdir(), f"update_helper_{ts}.log")
            script_content = f"""#!/bin/bash
# helper update (unix)
OLD_PID="{current_pid}"
OLD_PATH="$2"
NEW_PATH="$3"
LOG="{logpath}"
echo "helper start: $(date) PID $OLD_PID OLD_PATH=$OLD_PATH NEW_PATH=$NEW_PATH" >> "$LOG"

# Wait until PID exits (with timeout)
timeout={int(wait_timeout)}
elapsed=0
while kill -0 "$OLD_PID" 2>/dev/null; do
    sleep 1
    elapsed=$((elapsed+1))
    if [ "$elapsed" -ge "$timeout" ]; then
        echo "Timeout waiting for PID $OLD_PID after $elapsed seconds" >> "$LOG"
        exit 1
    fi
done

echo "Processo principale terminato. Tentativo di sostituzione..." >> "$LOG"

# Retry move (backup on first fail)
for i in 1 2 3; do
    if mv -f "$NEW_PATH" "$OLD_PATH" 2>>"$LOG"; then
        chmod +x "$OLD_PATH" 2>>"$LOG" || true
        echo "Move succeeded on attempt $i" >> "$LOG"
        break
    else
        echo "Move attempt $i failed" >> "$LOG"
        sleep 1
    fi
    if [ $i -eq 3 ]; then
        echo "Failed to move after retries" >> "$LOG"
        exit 1
    fi
done

# Avvia il nuovo eseguibile in background in modo detached (compatibile PyInstaller onefile)
if command -v setsid >/dev/null 2>&1; then
    setsid "$OLD_PATH" "${{@:4}}" >/dev/null 2>&1 &
elif command -v nohup >/dev/null 2>&1; then
    nohup "$OLD_PATH" "${{@:4}}" >/dev/null 2>&1 &
else
    "$OLD_PATH" "${{@:4}}" >/dev/null 2>&1 &
fi

echo "Started new process, exiting helper." >> "$LOG"
exit 0
"""
            script_path = os.path.join(tempfile.gettempdir(), f"update_helper_{ts}.sh")
            with open(script_path, 'w', newline='\n') as f:
                f.write(script_content)
            os.chmod(script_path, 0o755)

            # Avviamo lo script: passiamo (script_path, pid, old_path, new_path, ...argv)
            try:
                subprocess.Popen(
                    [script_path, str(current_pid), self.current_executable_path, new_executable_path] + sys.argv[1:],
                    preexec_fn=os.setsid,
                    close_fds=True
                )
            except Exception as e:
                self.log(f"Errore avviando lo script helper: {e}")
                messagebox.showerror("Errore Aggiornamento", f"Errore avviando lo script helper: {e}")
                return

            # Termina immediatamente il processo principale (affinché il file venga liberato)
            self.log("Helper avviato — esco immediatamente per permettere l'aggiornamento.")
            os._exit(0)

        elif current_system == 'Windows':
            self.log("Creazione del file batch helper per Windows.")
            logpath = os.path.join(tempfile.gettempdir(), f"update_helper_{ts}.log").replace('\\', '\\\\')
            script_content = f"""@echo off
setlocal
set LOG={logpath}
echo helper start: %date% %time% PID %1 OLD_PATH=%2 NEW_PATH=%3 >> "%LOG%"
set OLD_PID=%1
set OLD_PATH=%2
set NEW_PATH=%3

:waitloop
timeout /t 1 > nul
tasklist /FI "PID eq %OLD_PID%" | findstr /I "%OLD_PID%" > nul
if %ERRORLEVEL%==0 goto waitloop

REM Retry move
for /L %%i in (1,1,5) do (
    move /Y "%NEW_PATH%" "%OLD_PATH%" >> "%LOG%" 2>&1 && goto moved || timeout /t 1 > nul
)
echo Move failed after retries >> "%LOG%"
exit /b 1

:moved
start "" "%OLD_PATH%" %*
exit /b 0
"""
            script_path = os.path.join(tempfile.gettempdir(), f"update_helper_{ts}.bat")
            with open(script_path, 'w', newline='\r\n') as f:
                f.write(script_content)

            try:
                subprocess.Popen(
                    [script_path, str(current_pid), self.current_executable_path, new_executable_path] + sys.argv[1:],
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
                    close_fds=True
                )
            except Exception as e:
                self.log(f"Errore avviando il batch helper: {e}")
                messagebox.showerror("Errore Aggiornamento", f"Errore avviando il batch helper: {e}")
                return

            self.log("Helper batch avviato — esco immediatamente per permettere l'aggiornamento.")
            os._exit(0)

        else:
            self.log(f"Sistema operativo non supportato: {current_system}")
            messagebox.showerror("Errore Aggiornamento", f"Il sistema operativo {current_system} non è supportato.")
            return
