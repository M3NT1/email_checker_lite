import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
from exchangelib import Credentials, Account, DELEGATE, Configuration, Q, Folder
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.errors import ErrorNonExistentMailbox, UnauthorizedError, ErrorTimeoutExpired
from datetime import datetime, timedelta
import pytz
import urllib3
import warnings
import threading

# Kikapcsoljuk az SSL figyelmeztetéseket
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", category=UserWarning, module="exchangelib")

# Növeljük az XML olvasó mélységi korlátját és a timeout értéket
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
BaseProtocol.TIMEOUT = 300  # Növeljük a timeout értéket 300 másodpercre

# Fájlnevek
SETTINGS_FILE = 'beallitasok.json'

# Alapértelmezett adatok
default_settings = {
    "credentials": {
        "username": "",
        "password": "",
        "server": "",
        "email": ""
    },
    "folders": [],
    "subjects": []
}

# JSON fájl betöltése vagy létrehozása
def load_or_create_json(file_path, default_data):
    if not os.path.exists(file_path):
        with open(file_path, 'w', encoding='utf-8') as file:
            json.dump(default_data, file, indent=4, ensure_ascii=False)
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

settings_data = load_or_create_json(SETTINGS_FILE, default_settings)

# Globális változók
root = None
account = None

# Exchange szerverhez való csatlakozás
def connect_to_exchange():
    global account, settings_data
    credentials = Credentials(username=settings_data['credentials']['username'], password=settings_data['credentials']['password'])
    config = Configuration(server=settings_data['credentials']['server'], credentials=credentials)
    config.verify_ssl = False
    config.max_connections = 10

    try:
        account = Account(primary_smtp_address=settings_data['credentials']['email'], config=config, autodiscover=False,
                          access_type=DELEGATE)
    except Exception as e:
        messagebox.showerror("Hiba", f"Nem sikerült csatlakozni az Exchange szerverhez: {str(e)}")
        return False

    return True

# Mappák betöltése
def load_folders():
    global account, settings_data
    folders = []
    try:
        inbox = account.inbox
        folders.append(inbox)
        for folder in account.root.walk():
            if not folder.is_distinguished:
                folders.append(folder)
    except Exception as e:
        print(f"Hiba történt a mappák lekérdezése során: {str(e)}")

    settings_data['folders'] = [{"name": folder.name, "enabled": False} for folder in folders]
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as file:
        json.dump(settings_data, file, indent=4, ensure_ascii=False)

    messagebox.showinfo("Információ", "Mappák betöltve és elmentve a beallitasok.json fájlba.")

# E-mailek keresése
def search_emails():
    global account, settings_data

    date_str = date_entry.get()
    search_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    start = timezone.localize(datetime.combine(search_date, datetime.min.time()))
    end = timezone.localize(datetime.combine(search_date + timedelta(days=1), datetime.min.time()))

    subjects_to_search = settings_data["subjects"]

    results = {}
    for subject in subjects_to_search:
        query = Q(datetime_received__range=(start, end)) & Q(subject__icontains=subject)
        total_count = 0
        for folder_info in settings_data['folders']:
            if folder_info['enabled']:
                try:
                    folder = get_folder_by_name(account, folder_info['name'])
                    if folder:
                        emails = folder.filter(query)
                        count = emails.count()
                        total_count += count
                        print(f"Tárgy: {subject}, Mappa: {folder_info['name']}, Találatok száma: {count}")
                        for email in emails:
                            read_status = "Olvasott" if email.is_read else "Olvasatlan"
                            print(f"  - Feladó: {email.sender.email_address}, Időpont: {email.datetime_received}, Állapot: {read_status}")
                            print(f"    Tárgy: {email.subject}")
                except ErrorTimeoutExpired as e:
                    print(f"Időtúllépés történt a {folder_info['name']} mappa keresése során: {str(e)}")
                except Exception as e:
                    print(f"Hiba történt a {folder_info['name']} mappa keresése során: {str(e)}")

        results[subject] = total_count

    # Eredmények megjelenítése új ablakban
    result_window = tk.Toplevel(root)
    result_window.title("Keresési eredmények")
    result_window.geometry("500x400")

    result_text = tk.Text(result_window, wrap=tk.WORD, font=("Helvetica", 12))
    result_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    for subject, count in results.items():
        if count > 0:
            result_text.insert(tk.END, f"✔ {subject} (Találatok száma: {count})\n")
        else:
            result_text.insert(tk.END, f"✘ {subject} (Nincs találat)\n")

    result_text.config(state=tk.DISABLED)

# Mappa lekérése név alapján
def get_folder_by_name(account, folder_name):
    for folder in account.root.walk():
        if folder.name == folder_name:
            return folder
    return None

# GUI létrehozása
def create_gui():
    global root, date_entry

    root = tk.Tk()
    root.title("Egyszerűsített E-mail kereső")
    root.geometry("400x300")

    main_frame = tk.Frame(root, padx=10, pady=10)
    main_frame.pack(fill=tk.BOTH, expand=True)

    date_label = tk.Label(main_frame, text="Dátum (YYYY-MM-DD):", font=("Helvetica", 12))
    date_label.grid(row=0, column=0, sticky="w")

    date_entry = tk.Entry(main_frame, font=("Helvetica", 12))
    date_entry.grid(row=0, column=1, sticky="ew", pady=(0, 10))
    date_entry.insert(0, datetime.now().strftime('%Y-%m-%d'))

    load_folders_button = tk.Button(main_frame, text="Mappák betöltése", command=load_folders)
    load_folders_button.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

    search_emails_button = tk.Button(main_frame, text="E-mailek keresése", command=search_emails)
    search_emails_button.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 10))

    root.mainloop()

# Fő program
if __name__ == "__main__":
    timezone = pytz.timezone('Europe/Budapest')

    if connect_to_exchange():
        create_gui()
    else:
        print("Nem sikerült csatlakozni az Exchange szerverhez.")
