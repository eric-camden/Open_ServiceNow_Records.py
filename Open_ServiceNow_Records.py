# Imports
import tkinter as tk  # GUI Library
from tkinter import simpledialog, messagebox  # GUI Dialogs
import configparser  # Configuration file handling
import os  # OS-level file operations
import webbrowser  # Opens URLs in the default browser
import pyperclip  # Clipboard access
import re  # Regular expressions for matching patterns
import pystray  # System tray integration
from PIL import Image  # Image processing (for tray icon)
import threading  # Allows concurrent execution of tasks
import keyboard  # Global hotkey listener
import queue  # Communication queue between threads
from watchdog.observers import Observer  # Watches for file changes
from watchdog.events import FileSystemEventHandler  # Event handler for file changes
import sys  # System-level operations

# Paths & Defaults
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'CompanySettings.ini')
DEFAULT_COMPANY = 'mycompany'
DEFAULT_TICKET_LENGTH = 7

# Load configuration
config = configparser.ConfigParser()
if not os.path.exists(CONFIG_FILE):
    config['Settings'] = {
        'CompanyName': DEFAULT_COMPANY,
        'NeverAskAgain': '0',
        'TicketLength': str(DEFAULT_TICKET_LENGTH),
        'Hotkey': 'ctrl+shift+o'
    }
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)
else:
    config.read(CONFIG_FILE)

# Global variables initialized from config
company_name = config.get('Settings', 'CompanyName', fallback=DEFAULT_COMPANY)
never_ask_again = config.getint('Settings', 'NeverAskAgain', fallback=0)
ticket_length = config.getint('Settings', 'TicketLength', fallback=DEFAULT_TICKET_LENGTH)
hotkey = config.get('Settings', 'Hotkey', fallback='ctrl+shift+o')

# Reload the configuration dynamically
def reload_config():
    global company_name, never_ask_again, ticket_length, hotkey
    config.read(CONFIG_FILE)
    company_name = config.get('Settings', 'CompanyName', fallback=DEFAULT_COMPANY)
    never_ask_again = config.getint('Settings', 'NeverAskAgain', fallback=0)
    ticket_length = config.getint('Settings', 'TicketLength', fallback=DEFAULT_TICKET_LENGTH)
    hotkey = config.get('Settings', 'Hotkey', fallback='ctrl+shift+o')
    keyboard.clear_all_hotkeys()
    keyboard.add_hotkey(hotkey, lambda: q.put('open_interface'))

# Watches configuration file for changes
class ConfigChangeHandler(FileSystemEventHandler):
    def __init__(self, filepath):
        self.filepath = filepath

    def on_modified(self, event):
        if os.path.normpath(event.src_path) == os.path.normpath(self.filepath):
            reload_config()
            messagebox.showinfo("Config Reloaded", "Configuration was updated automatically.")

# Save settings to config file
def save_settings(new_company, new_ticket_length, never_ask, new_hotkey):
    config['Settings']['CompanyName'] = new_company
    config['Settings']['TicketLength'] = str(new_ticket_length)
    config['Settings']['NeverAskAgain'] = '1' if never_ask else '0'
    config['Settings']['Hotkey'] = new_hotkey
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

# Reset config file by deleting it
def reset_config():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
    messagebox.showinfo("Reset", "Configuration reset! Relaunch the script to enter new values.")

# Open a specific ServiceNow record based on clipboard or user input
def open_record(record_type, root):
    clipboard = pyperclip.paste()
    regex = rf"\b{record_type}\d{{{ticket_length}}}\b"
    match = re.search(regex, clipboard)
    default_value = match.group(0) if match else ''

    user_input = simpledialog.askstring("Input", f"Enter a {record_type} number (full or short):", initialvalue=default_value, parent=root)
    if user_input is None:
        messagebox.showinfo("Canceled", "Operation canceled.")
        return

    record_id = None
    match = re.search(regex, user_input)
    if match:
        record_id = match.group(0)
    else:
        digits = re.search(rf"\b\d{{1,{ticket_length}}}\b", user_input)
        if digits:
            num = digits.group(0).zfill(ticket_length)
            record_id = record_type + num

    if not record_id:
        messagebox.showwarning("Invalid", f"Invalid input for {record_type}!")
        return

    record_urls = {
        "INC": "incident.do?sysparm_query=number=",
        "CHG": "change_request.do?sysparm_query=number=",
        "CTASK": "change_task.do?sysparm_query=number=",
        "SCTASK": "sc_task.do?sysparm_query=number=",
        "REQ": "sc_request.do?sysparm_query=number=",
        "RITM": "sc_req_item.do?sysparm_query=number=",
        "TASK": "task.do?sysparm_query=number=",
        "KB": "kb_knowledge.do?sysparm_query=number="
    }

    url_path = record_urls.get(record_type)
    full_url = f"https://{company_name}.service-now.com/nav_to.do?uri={url_path}{record_id}"
    webbrowser.open(full_url)

def open_lookup(lookup_type, root):
    user_input = simpledialog.askstring("Lookup", f"Lookup {lookup_type} in ServiceNow (e.g. email for users):", parent=root)
    if not user_input:
        messagebox.showinfo("Canceled", "Operation canceled.")
        return

    lookup_urls = {
        "User": "sys_user.do?sysparm_query=email=",
        "Group": "sys_user_group.do?sysparm_query=name=",
        "CI": "cmdb_ci.do?sysparm_query=name="
    }

    url_path = lookup_urls.get(lookup_type)
    if not url_path:
        messagebox.showwarning("Unsupported", f"Unsupported lookup type: {lookup_type}")
        return

    full_url = f"https://{company_name}.service-now.com/nav_to.do?uri={url_path}{user_input}"
    webbrowser.open(full_url)

def ticket_interface():
    win = tk.Toplevel(root)
    win.title("ServiceNow Ticket Interface")

    button_data = [
        ('INC Ticket', 'INC'),
        ('CTASK', 'CTASK'),
        ('CHG Change', 'CHG'),
        ('SCTASK', 'SCTASK'),
        ('REQ', 'REQ'),
        ('RITM', 'RITM'),
        ('TASK', 'TASK'),
        ('KB', 'KB'),
        ('User Lookup', 'User'),
        ('Group Lookup', 'Group'),
        ('CI Lookup', 'CI')
    ]

    record_types = {'INC', 'CTASK', 'CHG', 'SCTASK', 'REQ', 'RITM', 'TASK', 'KB'}

    clipboard = pyperclip.paste().strip().upper()

    # Check clipboard for a known record number
    for record_type in record_types:
        regex = rf"\b{record_type}\d{{{ticket_length}}}\b"
        match = re.fullmatch(regex, clipboard)
        if match:
            win.destroy()
            open_record(record_type, root)
            return

    def make_button(text, command, row):
        def wrapped_command():
            command()
            win.destroy()
        tk.Button(win, text=text, width=20, command=wrapped_command).grid(row=row, column=0, padx=10, pady=5)

    # No auto-selection occurred; display buttons normally
    for idx, (btn_text, record_type) in enumerate(button_data):
        if record_type in record_types:
            make_button(btn_text, lambda rt=record_type: open_record(rt, win), idx)
        else:
            make_button(btn_text, lambda lt=record_type: open_lookup(lt, win), idx)

    tk.Button(win, text="Exit", width=20, command=win.destroy).grid(row=len(button_data), column=0, padx=10, pady=10)


def show_setup_window():
    def save(event=None):
        new_company = company_entry.get()
        new_hotkey = hotkey_entry.get()
        try:
            new_ticket_length = int(ticket_length_entry.get())
            if not (7 <= new_ticket_length <= 10):
                messagebox.showwarning("Invalid", "Ticket length must be between 7-10.")
                return
            save_settings(new_company, new_ticket_length, never_ask_var.get(), new_hotkey)
            messagebox.showinfo("Saved", f"Company: {new_company}, Length: {new_ticket_length}, Hotkey: {new_hotkey}")
            setup_win.destroy()
            messagebox.showinfo(
                "ServiceNow Helper",
                f"ServiceNow helper running.\nUse {new_hotkey} to open ticket interface.\nRight-click tray icon for more options."
            )
        except ValueError:
            messagebox.showwarning("Invalid", "Ticket length must be a number.")

    setup_win = tk.Toplevel(root)
    setup_win.title("Company Setup")

    tk.Label(setup_win, text="Customize Your ServiceNow Settings").grid(row=0, column=0, columnspan=2, pady=5)

    tk.Label(setup_win, text="Company name:").grid(row=1, column=0, sticky="e")
    company_entry = tk.Entry(setup_win)
    company_entry.insert(0, company_name)
    company_entry.grid(row=1, column=1)

    tk.Label(setup_win, text="Ticket length (7-10):").grid(row=2, column=0, sticky="e")
    ticket_length_entry = tk.Entry(setup_win)
    ticket_length_entry.insert(0, str(ticket_length))
    ticket_length_entry.grid(row=2, column=1)

    tk.Label(setup_win, text="Hotkey (e.g., ctrl+shift+o):").grid(row=3, column=0, sticky="e")
    hotkey_entry = tk.Entry(setup_win)
    hotkey_entry.insert(0, hotkey)
    hotkey_entry.grid(row=3, column=1)

    never_ask_var = tk.IntVar(value=never_ask_again)
    tk.Checkbutton(setup_win, text="Never ask again", variable=never_ask_var).grid(row=4, columnspan=2)

    tk.Button(setup_win, text="Save", command=save).grid(row=5, column=0)
    tk.Button(setup_win, text="Cancel", command=setup_win.destroy).grid(row=5, column=1)

    setup_win.bind('<Return>', save)  # BIND ENTER key to save


def create_tray():
    def on_clicked(icon, item):
        if item.text == "Reset Settings":
            reset_config()
        elif item.text == "Reload":
            messagebox.showinfo("Reload", "Reloading configuration.")
            reload_config()
        elif item.text == "Exit":
            icon.stop()
            root.quit()

    def reload_config():
        global company_name, never_ask_again, ticket_length, hotkey
        config.read(CONFIG_FILE)
        company_name = config.get('Settings', 'CompanyName', fallback=DEFAULT_COMPANY)
        never_ask_again = config.getint('Settings', 'NeverAskAgain', fallback=0)
        ticket_length = config.getint('Settings', 'TicketLength', fallback=DEFAULT_TICKET_LENGTH)
        hotkey = config.get('Settings', 'Hotkey', fallback='ctrl+shift+o')
        keyboard.clear_all_hotkeys()
        keyboard.add_hotkey(hotkey, lambda: q.put('open_interface'))


    menu = pystray.Menu(
        pystray.MenuItem('Reload', on_clicked),
        pystray.MenuItem('Reset Settings', on_clicked),
        pystray.MenuItem('Exit', on_clicked)
    )

    icon = pystray.Icon(
        "SNow",
        Image.new('RGB', (64, 64), color='black'),
        "ServiceNow Helper",
        menu=menu
    )

    icon.run()



def listen_hotkey():
    def on_hotkey():
        q.put('open_interface')

    keyboard.add_hotkey(hotkey, on_hotkey)
    keyboard.wait()


def poll_queue():
    try:
        task = q.get_nowait()
        if task == 'open_interface':
            ticket_interface()
    except queue.Empty:
        pass
    root.after(100, poll_queue)

import sys  # add this import at the top

def start_config_watcher():
    event_handler = ConfigChangeHandler(CONFIG_FILE)
    observer = Observer()
    observer.schedule(event_handler, path=os.path.dirname(CONFIG_FILE), recursive=False)
    observer.start()

# MAIN APPLICATION EXECUTION
root = tk.Tk()
root.withdraw()  # Hide the main Tkinter window
q = queue.Queue()  # Communication queue between threads

# Show setup or directly run depending on user's previous choice
if never_ask_again == 0:
    show_setup_window()
else:
    messagebox.showinfo("ServiceNow Helper", f"ServiceNow helper running.\nUse {hotkey} to open ticket interface.\nRight-click tray icon for more options.")

# Start threads for tray icon, hotkey listening, and config file watcher
threading.Thread(target=create_tray, daemon=True).start()
threading.Thread(target=listen_hotkey, daemon=True).start()
threading.Thread(target=start_config_watcher, daemon=True).start()

# Continuously poll the queue for events
poll_queue()

# Start the GUI main event loop
root.mainloop()
sys.exit()
