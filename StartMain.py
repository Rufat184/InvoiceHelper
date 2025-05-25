import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import platform
import subprocess
from pathlib import Path
import win32com.client
from PyPDF2 import PdfMerger

CONFIG_FILE = "config.json"
merged_files = []  # Данные для таблицы: список кортежей (from_email, file_path)

def open_file_cross_platform(file_path):
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", file_path])
        else:  # Linux
            subprocess.run(["xdg-open", file_path])
    except Exception as e:
        messagebox.showerror("Open File Error", f"Cannot open file: {e}")

def load_config():
    default = {
        "SAVE_FOLDER": "",
        "QR_PDF": "",
        "FORWARD_TO": "",
        "INBOX_ID": 6,
        "SUBFOLDER_NAME": "Invoices"
    }
    if Path(CONFIG_FILE).exists():
        with open(CONFIG_FILE, "r") as f:
            cfg = json.load(f)
        for k in default:
            cfg.setdefault(k, default[k])
        return cfg
    return default

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)

def check_and_open_outlook():
    try:
        win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception:
        os.system("start outlook")

def process_invoices(table, config):
    merged_files.clear()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(config["INBOX_ID"])
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Could not connect to Outlook:\n{e}")
        return

    subfolder = None
    for f in inbox.Folders:
        if f.Name.lower() == config["SUBFOLDER_NAME"].lower():
            subfolder = f
            break

    if not subfolder:
        messagebox.showerror("Folder Error", "Subfolder not found.")
        return

    items = subfolder.Items.Restrict("[Unread]=true")
    if items.Count == 0:
        messagebox.showinfo("Done", "There are no unread emails found.")
        return

    for item in items:
        if item.Class == 43 and ("invoice" in item.Subject.lower() or "invoice" in item.Body.lower()):
            sender = item.SenderEmailAddress
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)
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

                    merged_files.append((sender, output_pdf))
                    table.insert("", "end", values=(sender, os.path.basename(output_pdf), output_pdf))

            item.Unread = False
            item.Save()

    messagebox.showinfo("Done", "Processing complete.")

def forward_selected(config, table):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    app = win32com.client.Dispatch("Outlook.Application")
    for row in table.selection():
        file_path = table.item(row)["values"][2]
        subject = "Invoice"
        forward = app.CreateItem(0)
        forward.To = config["FORWARD_TO"]
        forward.Subject = subject
        forward.Body = "Please find the invoice attached."
        forward.Attachments.Add(file_path)
        forward.Send()
    messagebox.showinfo("Sent", "Selected invoices forwarded.")

def run_gui():
    config = load_config()

    root = tk.Tk()
    root.title("Invoice Handler")
    window_width = 550
    window_height = 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height - window_height) / 2)
    position_left = int((screen_width - window_width) / 2)
    root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")
    root.resizable(False, True)

    def choose_folder():
        path = filedialog.askdirectory()
        if path:
            entry_save.delete(0, tk.END)
            entry_save.insert(0, path)

    def choose_qr():
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            entry_qr.delete(0, tk.END)
            entry_qr.insert(0, path)

    def save_settings():
        config["SAVE_FOLDER"] = entry_save.get()
        config["QR_PDF"] = entry_qr.get()
        config["FORWARD_TO"] = entry_email.get()
        config["INBOX_ID"] = int(entry_inbox.get())
        config["SUBFOLDER_NAME"] = entry_subfolder.get()
        save_config(config)
        messagebox.showinfo("Saved", "Settings saved.")

    def start_process():
        check_and_open_outlook()
        for i in tree.get_children():
            tree.delete(i)
        process_invoices(tree, config)

    def on_double_click(event):
        selected = tree.selection()
        if selected:
            file_path = tree.item(selected[0])["values"][2]
            open_file_cross_platform(file_path)

    def delete_selected():
        for item in tree.selection():
            file_path = tree.item(item)["values"][2]
            try:
                os.remove(file_path)
            except Exception:
                pass
            tree.delete(item)

    tk.Label(root, text="Save Folder:").grid(row=0, column=0, sticky="w")
    entry_save = tk.Entry(root, width=50)
    entry_save.grid(row=0, column=1)
    entry_save.insert(0, config["SAVE_FOLDER"])
    tk.Button(root, text="Choose Folder", command=choose_folder, width=15).grid(row=0, column=2,  padx=(5, 10))

    tk.Label(root, text="QR PDF File:").grid(row=1, column=0, sticky="w")
    entry_qr = tk.Entry(root, width=50)
    entry_qr.grid(row=1, column=1)
    entry_qr.insert(0, config["QR_PDF"])
    tk.Button(root, text="Choose PDF File", command=choose_qr, width=15).grid(row=1, column=2, padx=(5, 10), pady=5)

    tk.Label(root, text="Forward To Email:").grid(row=2, column=0, sticky="w")
    entry_email = tk.Entry(root, width=50)
    entry_email.grid(row=2, column=1)
    entry_email.insert(0, config["FORWARD_TO"])

    tk.Label(root, text="Inbox Folder ID:").grid(row=3, column=0, sticky="w")
    entry_inbox = tk.Entry(root, width=10)
    entry_inbox.grid(row=3, column=1, sticky="w", padx=(0, 10), pady=5)
    entry_inbox.insert(0, str(config["INBOX_ID"]))

    tk.Label(root, text="Subfolder Name:").grid(row=4, column=0, sticky="w")
    entry_subfolder = tk.Entry(root, width=50)
    entry_subfolder.grid(row=4, column=1)
    entry_subfolder.insert(0, config["SUBFOLDER_NAME"])

    tk.Button(root, text="💾 Save settings", command=save_settings).grid(row=5, column=1, sticky="w", pady=5)
    tk.Button(root, text="▶ Start Process", command=start_process).grid(row=5, column=1, sticky="e", pady=5)

    # Table
    table_frame = tk.Frame(root)
    root.grid_rowconfigure(6, weight=1)
    root.grid_columnconfigure(1, weight=1)

    table_frame.grid(row=6, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")

    scrollbar = tk.Scrollbar(table_frame)
    scrollbar.pack(side="right", fill="y")

    tree = ttk.Treeview(table_frame, columns=("email", "file"), show="headings", height=8, yscrollcommand=scrollbar.set)
    scrollbar.config(command=tree.yview)

    tree.heading("email", text="From Email")
    tree.heading("file", text="Merged PDF File")
    tree.column("email", width=220)
    tree.column("file", width=250)
    tree.pack(side="left", fill="both", expand=True)
    tree.bind("<Double-1>", on_double_click)

    tk.Button(root, text="🗑 Delete Selected", command=delete_selected).grid(row=7, column=0, padx=10, sticky="w", pady=10)
    tk.Button(root, text="📤 Send Selected", command=lambda: forward_selected(config, tree)).grid(row=7, column=2, sticky="e", pady=10, padx=(5, 28))

    root.mainloop()

if __name__ == "__main__":
    run_gui()
