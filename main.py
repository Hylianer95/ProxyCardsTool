import os
import re
import sys
from typing import Optional
from io import BytesIO
from tkinter import (
    Tk, Label, Button, Text, Scrollbar, filedialog, messagebox,
    Entry, StringVar, END, BooleanVar, Checkbutton
)

import requests
from PIL import Image

BASE_URL = "https://static.dotgg.gg/onepiece/card/"

def resource_path(relative_path: str) -> str:
    """
    Ermöglicht Zugriff auf Ressourcen, auch wenn es als .exe (PyInstaller) gebaut wird.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def download_card(card_code: str, save_dir: str, copy_index: Optional[int] = None) -> str:
    """Download card image and save as PNG. Supports multiple copies."""
    url = f"{BASE_URL}{card_code}.webp"
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(f"Image not found at {url}")

    raw = BytesIO(resp.content)
    os.makedirs(save_dir, exist_ok=True)

    if copy_index:
        out_file = os.path.join(save_dir, f"{card_code}_{copy_index}.png")
    else:
        out_file = os.path.join(save_dir, f"{card_code}.png")

    with Image.open(raw) as im:
        im = im.convert("RGBA")
        im.save(out_file, "PNG")

    return out_file

def start_gui():
    root = Tk()
    root.title("OnePiece Deck Downloader")

    # eigenes Fenstericon (muss app.ico im selben Ordner liegen)
    try:
        root.iconbitmap(resource_path("app.ico"))
    except Exception:
        pass

    # Kopf-Label (ohne Beispiel im Text)
    Label(root, text="Enter your deck list:\n").pack()

    text_box = Text(root, height=20, width=40)
    text_box.pack(side="left", padx=10, pady=10)

    scrollbar = Scrollbar(root, command=text_box.yview)
    scrollbar.pack(side="right", fill="y")
    text_box.config(yscrollcommand=scrollbar.set)

    Label(root, text="Save folder:").pack(pady=(10, 0))
    folder_var = StringVar(value=os.getcwd())
    Entry(root, textvariable=folder_var, width=40).pack(padx=10, pady=5)

    def browse_folder():
        selected = filedialog.askdirectory(title="Select save folder")
        if selected:
            folder_var.set(selected)

    Button(root, text="Browse", command=browse_folder).pack()

    multiply_var = BooleanVar(value=False)
    status_label = Label(root, text="")
    status_label.pack(pady=(5, 0))

    def normalize_folder(path_raw: str) -> str:
        p = path_raw.strip().strip('"').strip("'")
        return os.path.normpath(os.path.abspath(p))

    def on_download():
        lines = text_box.get("1.0", END).strip().splitlines()
        if not lines:
            messagebox.showerror("Error", "Please enter a deck list!")
            return

        save_dir_raw = folder_var.get()
        if not save_dir_raw.strip():
            messagebox.showerror("Error", "Please select a save folder!")
            return

        save_dir = normalize_folder(save_dir_raw)
        success, failed = [], []
        total = len(lines)

        for idx, line in enumerate(lines, start=1):
            line = line.strip()
            if not line:
                continue

            qty_match = re.match(r"(\d+)x\s*([A-Z]+\d{2}-\d{3})", line, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
                code = qty_match.group(2)
            else:
                code_match = re.search(r"([A-Z]+\d{2}-\d{3})", line)
                if not code_match:
                    continue
                code = code_match.group(1)
                qty = 1

            effective_qty = qty if multiply_var.get() else 1
            status_label.config(text=f"Downloading {idx}/{total}: {code} …")
            root.update_idletasks()

            try:
                for i in range(effective_qty):
                    path = download_card(
                        code, save_dir,
                        copy_index=(i + 1) if effective_qty > 1 else None
                    )
                    success.append(path)
            except Exception:
                failed.append(f"{qty}x{code}")

        status_label.config(text="Download finished.")
        msg = f"{len(success)} images saved in:\n{save_dir}"
        if failed:
            msg += f"\nFailed: {', '.join(failed)}"
        messagebox.showinfo("Done", msg)

    Button(root, text="Download Deck", command=on_download).pack(pady=(10, 0))
    Checkbutton(root, text="Download Cards multiply", variable=multiply_var).pack(pady=(4, 4))

    # Beispiel unten bleibt
    example_text = (
        "e.g.\n"
        "1xEB02-010\n"
        "4xST18-001\n"
        "4xST18-004\n"
        "2xOP05-070\n"
        "4xEB02-035\n"
        "3xEB02-061\n"
        "4xOP07-064\n"
        "3xST18-005\n"
        "4xEB02-017\n"
        "3xEB02-019\n"
        "1xOP03-072\n"
        "4xOP05-076\n"
        "4xEB02-041\n"
        "4xOP09-078\n"
        "3xOP01-119\n"
        "3xOP12-037"
    )
    Label(root, text=example_text, justify="left").pack(padx=10, pady=(0, 10), anchor="w")

    root.mainloop()

if __name__ == "__main__":
    try:
        start_gui()
    except Exception as e:
        messagebox.showerror("Fatal Error", str(e))
