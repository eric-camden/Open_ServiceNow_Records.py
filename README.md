# ServiceNow Helper

A Python desktop assistant to quickly open ServiceNow records from your clipboard or custom input, using a system tray icon and hotkey. Designed to save time for IT professionals who frequently work in ServiceNow.

---

## 🚀 Features

- ⌨️ Global hotkey to trigger the interface (configurable)
- 📋 Automatically reads ServiceNow record numbers from your clipboard
- 🧠 Recognizes record types like `INC`, `CHG`, `RITM`, `CTASK`, etc.
- 🔎 Lookup ServiceNow users, groups, and configuration items
- 💾 Remembers settings using a local `.ini` config file
- 🔁 Auto-reloads config on change using `watchdog`
- 🖱️ System tray menu with **Reload**, **Reset Settings**, and **Exit**

---

## 🖥️ Demo

When you press `Ctrl+Shift+O` (or your configured hotkey), the app:

1. Checks your clipboard for a ServiceNow record (e.g. `INC1234567`).
2. If a match is found, opens the record directly in your browser.
3. If no match, it shows a window to manually enter or look up records.

---

## 🧰 Requirements

Install these Python packages:

```bash
pip install pyperclip pystray pillow keyboard watchdog
