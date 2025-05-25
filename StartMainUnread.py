import tkinter as tk
from tkinter import filedialog, messagebox
import json
from pathlib import Path
import win32com.client
import os
from PyPDF2 import PdfMerger

CONFIG_FILE = "config.json"

def check_and_open_outlook():
    try:
        # Try to connect to an existing Outlook instance
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("Outlook is already running.")
    except Exception:
        # If Outlook is not running, open it
        print("Outlook is not running. Attempting to open it...")
        os.system("start outlook")
        try:
            # Reconnect to Outlook after opening it
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            print("Outlook has been opened successfully.")
        except Exception as e:
            print(f"Failed to open Outlook: {e}")

def load_config():
    default_config = {
        "SAVE_FOLDER": "",
        "QR_PDF": "",
        "FORWARD_TO": "",
        "INBOX_ID": 6,
        "SUBFOLDER_NAME": "Invoices"
    }
    if Path(CONFIG_FILE).exists():
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
        # Ensure all default keys are present
        for key, value in default_config.items():
            config.setdefault(key, value)
        return config
    return default_config

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)

def process_invoices():
    config = load_config()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(config.get("INBOX_ID", 6))
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Could not connect to Outlook:\n{e}")
        return

    # Найти подпапку
    subfolder_name = config.get("SUBFOLDER_NAME", "").lower()
    invoices_folder = None
    for folder in inbox.Folders:
        if folder.Name.lower() == subfolder_name:
            invoices_folder = folder
            break

    if not invoices_folder:
        messagebox.showerror("Folder Error", f"Subfolder '{subfolder_name}' not found in Inbox.")
        return

    # Обработка только непрочитанных писем
    items = invoices_folder.Items.Restrict("[Unread]=true")
    if items.Count == 0:
        messagebox.showinfo("No Emails", "There are no unread emails found. Process is completed.")
        return

    for item in items:
        if item.Class == 43 and ("invoice" in item.Subject.lower() or "invoice" in item.Body.lower()):
            attachments = item.Attachments
            for i in range(1, attachments.Count + 1):
                att = attachments.Item(i)
                if att.FileName.lower().endswith(".pdf"):
                    saved_pdf = os.path.join(config["SAVE_FOLDER"], att.FileName)
                    att.SaveAsFile(saved_pdf)

                    base_name = Path(att.FileName).stem
                    ext = Path(att.FileName).suffix
                    counter = 1
                    output_pdf = os.path.join(config["SAVE_FOLDER"], f"merged_{base_name}_{counter}{ext}")
                    # Увеличиваем индекс, если файл уже существует
                    while os.path.exists(output_pdf):
                        counter += 1
                        output_pdf = os.path.join(config["SAVE_FOLDER"], f"merged_{base_name}_{counter}{ext}")
                    merger = PdfMerger()
                    merger.append(saved_pdf)
                    merger.append(config["QR_PDF"])
                    merger.write(output_pdf)
                    merger.close()

                    forward = item.Forward()
                    forward.Recipients.Add(config["FORWARD_TO"])
                    forward.Subject = item.Subject
                    forward.Attachments.Add(output_pdf)
                    forward.Send()

            item.Unread = False
            item.Save()

    messagebox.showinfo("Completed", "Processing completed successfully.")


def run_gui():
    config = load_config()

    def choose_folder():
        path = filedialog.askdirectory()
        if path:
            entry_save_folder.delete(0, tk.END)
            entry_save_folder.insert(0, path)

    def choose_qr_pdf():
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            entry_qr_pdf.delete(0, tk.END)
            entry_qr_pdf.insert(0, path)

    def on_save():
        try:
            config["SAVE_FOLDER"] = entry_save_folder.get()
            config["QR_PDF"] = entry_qr_pdf.get()
            config["FORWARD_TO"] = entry_forward_to.get()
            config["INBOX_ID"] = int(entry_inbox_id.get())
            config["SUBFOLDER_NAME"] = entry_subfolder.get()
            save_config(config)
            messagebox.showinfo("Saved", "Settings have been saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error during saving settings: {e}")

    def on_run():
        try:
            check_and_open_outlook()
            process_invoices()
        except Exception as e:
            messagebox.showerror("Error", f"Error during processing: {e}")

    root = tk.Tk()
    root.title("Invoice Processing Settings")

    # Set fixed window size
    window_width = 580
    window_height = 190
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height - window_height) / 2)
    position_left = int((screen_width - window_width) / 2)
    root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")
    root.resizable(False, False)  # Disable resizing

    tk.Label(root, text="Folder to save PDF:").grid(row=0, column=0, sticky="w")
    entry_save_folder = tk.Entry(root, width=50)
    entry_save_folder.insert(0, config["SAVE_FOLDER"])
    entry_save_folder.grid(row=0, column=1)
    tk.Button(root, text="Choose", command=choose_folder, width=8).grid(row=0, column=2, padx=(5, 10))

    tk.Label(root, text="QR PDF файл:").grid(row=1, column=0, sticky="w")
    entry_qr_pdf = tk.Entry(root, width=50)
    entry_qr_pdf.insert(0, config["QR_PDF"])
    entry_qr_pdf.grid(row=1, column=1)
    tk.Button(root, text="Choose", command=choose_qr_pdf, width=8).grid(row=1, column=2, padx=(5, 10))

    tk.Label(root, text="Email to forward:").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
    entry_forward_to = tk.Entry(root, width=50)
    entry_forward_to.grid(row=2, column=1, columnspan=1, padx=(10, 10), pady=5)
    entry_forward_to.insert(0, config["FORWARD_TO"])

    tk.Label(root, text="Inbox Folder ID (default is 6):").grid(row=3, column=0, sticky="w")
    entry_inbox_id = tk.Entry(root, width=10)
    entry_inbox_id.insert(0, str(config["INBOX_ID"]))
    entry_inbox_id.grid(row=3, column=1, sticky="w", padx=(10, 10), pady=5)

    tk.Label(root, text="Subfolder name (e.g. 'Invoices')").grid(row=4, column=0, sticky="w")
    entry_subfolder = tk.Entry(root, width=50)
    entry_subfolder.insert(0, config.get("SUBFOLDER_NAME", "Invoices"))
    entry_subfolder.grid(row=4, column=1, padx=(10, 10), pady=5)

    tk.Button(root, text="Save settings", command=on_save).grid(row=5, column=1, sticky="w", padx=(30, 30), pady=10)
    tk.Button(root, text="▶ Start Process", command=on_run).grid(row=5, column=1, sticky="w", padx=(120, 120), pady=10)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
