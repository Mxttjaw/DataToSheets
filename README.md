🤖 Bot per l'Automazione dei Dati da Email
Questo progetto è un'applicazione desktop creata in Python che automatizza il processo di estrazione di informazioni (come nomi, cognomi, email, telefoni e date) da file di testo grezzi. I dati estratti vengono poi caricati e organizzati automaticamente in un foglio di calcolo di Google Sheets. L'applicazione offre un'interfaccia utente grafica (GUI) intuitiva che permette di configurare il processo e monitorarne l'avanzamento.

✨ Caratteristiche Principali
Estrazione Dati Automatica: Analizza file di testo per identificare e raccogliere informazioni specifiche.

Integrazione con Google Sheets: Carica i dati estratti direttamente in un foglio di calcolo pre-configurato.

Interfaccia Utente Grafica (GUI): Un'applicazione desktop facile da usare con cui interagire.

Gestione Configurazione: Mantiene le impostazioni di configurazione in un file dedicato (config.ini).

Compatibilità Multi-piattaforma: Può essere eseguito come applicazione nativa su Windows e macOS.

🚀 Installazione e Avvio
Prerequisiti
Assicurati di avere installato Python 3.8 o superiore e git.

1. Clonare il Repository
   Apri il terminale o il prompt dei comandi e clona il progetto:

git clone https://github.com/tuo-username/tuo-repo.git
cd tuo-repo

2. Configurazione dell'Ambiente
   Crea un ambiente virtuale (consigliato) e installa tutte le dipendenze necessarie dal file requirements.txt.

python -m venv venv

# Su Windows

venv\Scripts\activate

# Su macOS/Linux

source venv/bin/activate
pip install -r requirements.txt

3. Configurazione di Google Sheets
   Segui queste istruzioni per configurare l'integrazione con Google Sheets:

Crea un progetto in Google Cloud Console.

Abilita le API Google Sheets API e Google Drive API.

Crea un account di servizio e scarica il file JSON con le credenziali.

Rinomina il file JSON in service_account.json e posizionalo nella cartella del tuo progetto.

Condividi il tuo Google Sheet con l'indirizzo email dell'account di servizio (si trova nel file service_account.json).

⚙️ Utilizzo del Bot
Avviare l'Applicazione:

python src/bot.py

Interfaccia Utente:

"Seleziona File": Clicca su questo pulsante per scegliere il file di testo da cui estrarre i dati.

"Avvia": Avvia il processo di estrazione e caricamento dei dati.

Log: La finestra in basso mostrerà lo stato di avanzamento e i messaggi di errore.

📦 Creazione di un Eseguibile
Se vuoi creare un'applicazione autonoma che non richieda l'installazione di Python, puoi usare PyInstaller.

1. Installa PyInstaller
   Assicurati di essere nel tuo ambiente virtuale e installa PyInstaller:

pip install pyinstaller

2. Prepara le Icone
   Posiziona i file delle icone (.ico per Windows e .icns per macOS) nella cartella assets/icons/.

tuo-repo/
├── src/
│ └── bot.py
├── assets/
│ ├── icons/
│ │ ├── bot_icon.ico
│ │ └── bot_icon.icns
...

3. Esegui il Comando di Creazione
   Per Windows:

pyinstaller --onefile --windowed --icon=assets/icons/bot_icon.ico src/bot.py

Per macOS:

pyinstaller --onefile --windowed --icon=assets/icons/bot_icon.icns src/bot.py

Per Linux:

pyinstaller --onefile --name DataToSheets-Linux --add-data="src/assets/icons/\*:assets/icons" --additional-hooks-dir=hooks --hidden-import='PIL.\_tkinter_finder' --hidden-import='PIL.Image' --hidden-import='PIL.ImageTk' src/bot.py
·
Il file eseguibile (.exe o .app) si troverà nella cartella dist/.

📁 Struttura del Progetto
.
├── src/
│ └── bot.py
├── assets/
│ ├── icons/
│ │ ├── bot_icon.ico
│ │ └── bot_icon.icns
├── venv/
├── requirements.txt
├── .env
├── config.ini
└── README.md

📄 Licenza
Questo progetto è distribuito con licenza MIT. Vedi il file LICENSE per maggiori dettagli.
