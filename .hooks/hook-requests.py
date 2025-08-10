from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Questo hook garantisce che l'intero pacchetto 'requests'
# e i suoi sub-moduli vengano inclusi nell'eseguibile finale.

# Colleziona tutti i sottomoduli del pacchetto requests
hiddenimports = collect_submodules('requests')

# Colleziona i file di dati necessari (come i certificati SSL di certifi)
datas = collect_data_files('requests')

# Aggiunge i pacchetti associati a requests per una maggiore sicurezza
# che tutto venga incluso.
hiddenimports += [
    'requests.auth',
    'requests.compat',
    'requests.cookies',
    'requests.exceptions',
    'requests.models',
    'requests.structures',
    'requests.status_codes',
    'requests.utils',
    'idna',
    'certifi',
    'urllib3',
    'chardet'
]
