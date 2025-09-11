# main_local_search_a4_formats.py
import os, re, math, sys
from typing import Optional, List, Tuple
from io import BytesIO
from tkinter import (
    Tk, Label, Button, Text, Scrollbar, filedialog, messagebox,
    Entry, StringVar, END, BooleanVar, Toplevel, Canvas, Frame,
    IntVar, Radiobutton, DISABLED, NORMAL, LabelFrame, Checkbutton, OptionMenu
)

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from concurrent.futures import ThreadPoolExecutor, as_completed
from PIL import Image, ImageTk, ImageOps, ImageFilter, Image as PILImage, ImageDraw

DOTGG_BASE = "https://static.dotgg.gg/onepiece/card/"
LIMITLESS_BASE = "https://limitlesstcg.nyc3.cdn.digitaloceanspaces.com/one-piece/"
MAX_PARALLEL, TIMEOUT_SEC = 16, 8
THUMB_W, THUMB_H, THUMB_COLS = 150, 210, 4
P_MIN, P_MAX = 1, 10
IMAGE_EXTS = {".png",".webp",".jpg",".jpeg",".bmp"}

def make_session() -> requests.Session:
    s = requests.Session()
    retries = Retry(total=2, backoff_factor=0.2,
                    status_forcelist=(429,500,502,503,504),
                    allowed_methods=frozenset(["GET"]))
    adapter = HTTPAdapter(pool_connections=MAX_PARALLEL, pool_maxsize=MAX_PARALLEL, max_retries=retries)
    s.mount("http://", adapter); s.mount("https://", adapter)
    s.headers.update({"User-Agent":"OP-Card-Grabber/6.0 (+Tkinter)","Accept":"image/webp,image/*;q=0.8,*/*;q=0.5"})
    return s
SESSION = make_session()

def request_ok(url: str) -> Optional[bytes]:
    try:
        r = SESSION.get(url, timeout=TIMEOUT_SEC)
        if r.status_code == 200 and r.content:
            return r.content
    except Exception:
        pass
    return None

def candidates_dotgg(card_code: str) -> List[str]:
    code_dash = card_code.strip().upper()
    code_us = code_dash.replace("-", "_")
    urls = [f"{DOTGG_BASE}{code_dash}.webp"]
    urls += [f"{DOTGG_BASE}{code_dash}_p{i}.webp" for i in range(P_MIN, P_MAX+1)]
    urls += [f"{DOTGG_BASE}{code_us}.webp"]
    urls += [f"{DOTGG_BASE}{code_us}_p{i}.webp" for i in range(P_MIN, P_MAX+1)]
    seen, out = set(), []
    for u in urls:
        if u not in seen: seen.add(u); out.append(u)
    return out

def candidates_limitless(card_code: str) -> List[str]:
    code = card_code.strip().upper()
    m = re.match(r"^([A-Z]+\d{2})-", code)
    set_folder = m.group(1) if m else code.split("-")[0]
    base = f"{LIMITLESS_BASE}{set_folder}/"
    urls = [f"{base}{code}_EN.webp"]
    urls += [f"{base}{code}_p{i}_EN.webp" for i in range(P_MIN, P_MAX+1)]
    return urls

def candidates_local_by_code(card_code: str, root_dir: str) -> List[str]:
    code = card_code.strip().upper()
    hits: List[str] = []
    for base, _dirs, files in os.walk(root_dir):
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext in IMAGE_EXTS and code in name.upper():
                hits.append(os.path.join(base, name))
    def sort_key(p: str):
        n = os.path.basename(p).upper()
        plain = any(n == f"{code}{ext}".upper() for ext in (".png",".webp",".jpg",".jpeg"))
        return (0 if plain else 1, n)
    hits.sort(key=sort_key)
    return hits

def candidates_local_substr(term: str, root_dir: str) -> List[str]:
    needle = term.strip().upper()
    hits: List[str] = []
    for base, _dirs, files in os.walk(root_dir):
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext in IMAGE_EXTS and needle in name.upper():
                hits.append(os.path.join(base, name))
    hits.sort(key=lambda p: os.path.basename(p).upper())
    return hits

def probe_urls_in_parallel(urls: List[str]) -> List[Tuple[str, bytes]]:
    results: List[Optional[Tuple[str, bytes]]] = [None] * len(urls)
    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as pool:
        futmap = {pool.submit(request_ok, u): (i, u) for i, u in enumerate(urls)}
        for fut in as_completed(futmap):
            i, u = futmap[fut]; data = fut.result()
            if data: results[i] = (u, data)
    return [x for x in results if x is not None]

def probe_all_arts(card_code: str, source: str, local_dir: Optional[str] = None) -> List[Tuple[str, bytes]]:
    if source == "dotgg":
        return probe_urls_in_parallel(candidates_dotgg(card_code))
    if source == "limitless":
        return probe_urls_in_parallel(candidates_limitless(card_code))
    if source == "local":
        if not local_dir or not os.path.isdir(local_dir): return []
        out: List[Tuple[str, bytes]] = []
        for p in candidates_local_by_code(card_code, local_dir):
            try:
                with open(p, "rb") as f: out.append((p, f.read()))
            except Exception: pass
        return out
    return []

def make_padded_thumb(img_bytes: bytes, w: int, h: int) -> ImageTk.PhotoImage:
    with PILImage.open(BytesIO(img_bytes)) as im:
        im = im.convert("RGBA")
        im.thumbnail((w, h), PILImage.LANCZOS)
        canvas = PILImage.new("RGBA", (w, h), (255, 255, 255, 0))
        x = (w - im.width) // 2; y = (h - im.height) // 2
        canvas.paste(im, (x, y))
        return ImageTk.PhotoImage(canvas)

def pick_art_popup(root: Tk, title_text: str, found_images: List[Tuple[str, bytes]]) -> Optional[bytes]:
    if not found_images: return None
    win = Toplevel(root); win.title(f"Select art for {title_text}")
    # Try to set window icon for the picker, ignore if not found
    try:
        win.iconbitmap(resource_path("app.ico"))
    except Exception:
        pass
    win.geometry("900x580")
    win.rowconfigure(0, weight=1); win.columnconfigure(0, weight=1)
    canvas = Canvas(win, highlightthickness=0, bd=0)
    canvas.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
    scr = Scrollbar(win, orient="vertical", command=canvas.yview); scr.grid(row=0, column=1, sticky="ns")
    canvas.configure(yscrollcommand=scr.set)
    frame = Frame(canvas, bd=0, highlightthickness=0)
    inner = canvas.create_window((0, 0), window=frame, anchor="nw")
    thumbs: List[ImageTk.PhotoImage] = []; chosen = {"data": None}
    for idx, (url, data) in enumerate(found_images):
        try:
            tkimg = make_padded_thumb(data, THUMB_W, THUMB_H)
        except Exception:
            continue
        thumbs.append(tkimg)
        r, c = divmod(idx, THUMB_COLS)
        cell = Frame(frame, bd=1, relief="groove", width=THUMB_W+8, height=THUMB_H+38, highlightthickness=0)
        cell.grid(row=r, column=c, padx=3, pady=3, sticky="n")
        cell.grid_propagate(False)
        lbl = Label(cell, image=tkimg); lbl.pack(padx=3, pady=2)
        label_text = url if os.path.isabs(url) else url.rsplit("/", 1)[-1]
        cap = Label(cell, text=os.path.basename(label_text)); cap.pack(padx=2, pady=(0,2))
        def on_dbl(_e, b=data): chosen["data"] = b; win.destroy()
        lbl.bind("<Double-Button-1>", on_dbl); cell.bind("<Double-Button-1>", on_dbl)
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(inner, width=e.width))
    win.transient(root); win.grab_set(); root.wait_window(win)
    return chosen["data"]

def save_png(image_bytes: bytes, out_path: str,
             add_border: bool = True, border_px: int = 50,
             do_upscale: bool = False, min_height_px: int = 1500,
             dpi: int = 300) -> None:
    with PILImage.open(BytesIO(image_bytes)) as im:
        im = im.convert("RGBA")
        if do_upscale:
            w, h = im.size
            if h < min_height_px:
                scale = float(min_height_px) / float(h)
                im = im.resize((max(1, int(round(w * scale))), min_height_px), PILImage.LANCZOS)
                im = im.filter(ImageFilter.SHARPEN)
        if add_border and border_px > 0:
            im = ImageOps.expand(im, border=border_px, fill="white")
        im.save(out_path, "PNG", optimize=True, dpi=(dpi, dpi))

def mm_to_px(mm: float, dpi: int) -> int:
    return int(round(mm / 25.4 * dpi))

def draw_crop_marks(draw: ImageDraw.ImageDraw, x: int, y: int, w: int, h: int,
                    length_px: int, gap_px: int, stroke_px: int = 1, color=(0,0,0)) -> None:
    draw.line([(x - gap_px - length_px, y - gap_px), (x - gap_px, y - gap_px)], fill=color, width=stroke_px)
    draw.line([(x - gap_px, y - gap_px - length_px), (x - gap_px, y - gap_px)], fill=color, width=stroke_px)
    draw.line([(x + w + gap_px, y - gap_px), (x + w + gap_px + length_px, y - gap_px)], fill=color, width=stroke_px)
    draw.line([(x + w + gap_px, y - gap_px - length_px), (x + w + gap_px, y - gap_px)], fill=color, width=stroke_px)
    draw.line([(x - gap_px - length_px, y + h + gap_px), (x - gap_px, y + h + gap_px)], fill=color, width=stroke_px)
    draw.line([(x - gap_px, y + h + gap_px), (x - gap_px, y + h + gap_px + length_px)], fill=color, width=stroke_px)
    draw.line([(x + w + gap_px, y + h + gap_px), (x + w + gap_px + length_px, y + h + gap_px)], fill=color, width=stroke_px)
    draw.line([(x + w + gap_px, y + h + gap_px), (x + w + gap_px, y + h + gap_px + length_px)], fill=color, width=stroke_px)

def build_a4_pages(images: List[bytes], dpi: int,
                   card_w_mm: float, card_h_mm: float,
                   margin_x_mm: float, margin_y_mm: float,
                   gap_x_mm: float, gap_y_mm: float,
                   crop_marks: bool = False, crop_len_mm: float = 2.5, crop_gap_mm: float = 0.8,
                   crop_stroke_px: int = 1) -> List[PILImage.Image]:
    page_w = mm_to_px(210, dpi); page_h = mm_to_px(297, dpi)
    card_w = mm_to_px(card_w_mm, dpi); card_h = mm_to_px(card_h_mm, dpi)
    margin_x = mm_to_px(margin_x_mm, dpi); margin_y = mm_to_px(margin_y_mm, dpi)
    gap_x = mm_to_px(gap_x_mm, dpi); gap_y = mm_to_px(gap_y_mm, dpi)
    crop_len = mm_to_px(crop_len_mm, dpi); crop_gap = mm_to_px(crop_gap_mm, dpi)

    cols, rows = 3, 3
    needed_w = 2*margin_x + cols*card_w + (cols-1)*gap_x
    needed_h = 2*margin_y + rows*card_h + (rows-1)*gap_y
    if needed_w > page_w and (cols-1) > 0:
        over = needed_w - page_w; gap_x = max(0, gap_x - max(0, over // (cols-1)))
    if needed_h > page_h and (rows-1) > 0:
        over = needed_h - page_h; gap_y = max(0, gap_y - max(0, over // (rows-1)))

    pages: List[PILImage.Image] = []
    for page_idx in range(0, math.ceil(len(images)/9)):
        chunk = images[page_idx*9:(page_idx+1)*9]
        page = PILImage.new("RGB", (page_w, page_h), "white")
        draw = ImageDraw.Draw(page)
        for i, img_bytes in enumerate(chunk):
            with PILImage.open(BytesIO(img_bytes)) as im:
                im = im.convert("RGB")
                im = im.resize((card_w, card_h), PILImage.LANCZOS)
                r = i // cols; c = i % cols
                x = margin_x + c * (card_w + gap_x)
                y = margin_y + r * (card_h + gap_y)
                page.paste(im, (x, y))
                if crop_marks:
                    draw_crop_marks(draw, x, y, card_w, card_h, crop_len, crop_gap, crop_stroke_px)
        pages.append(page)
    return pages

def ensure_unique_suffix(path: str) -> str:
    base, ext = os.path.splitext(path)
    candidate = path; n = 1
    while os.path.exists(candidate): candidate = f"{base} ({n}){ext}"; n += 1
    return candidate

def next_unique(target_dir: str, name: str, ext: str) -> str:
    return ensure_unique_suffix(os.path.join(target_dir, f"{name}{ext}"))

def candidates_best_first_dotgg_or_limitless(code, source):
    return (candidates_dotgg(code) if source == "dotgg" else candidates_limitless(code))

def download_card_default(card_code: str, source: str, local_dir: Optional[str]) -> bytes:
    if source == "local":
        paths = candidates_local_by_code(card_code, local_dir or ".")
        if not paths: raise RuntimeError(f"No local image found for {card_code}.")
        with open(paths[0], "rb") as f: return f.read()
    else:
        for url in candidates_best_first_dotgg_or_limitless(card_code, source):
            data = request_ok(url)
            if data: return data
    raise RuntimeError(f"No image found for {card_code} on source '{source}'.")

def looks_like_code(s: str) -> bool:
    return re.fullmatch(r"[A-Z]+\d{2}-\d{3}", s.strip().upper()) is not None

def resource_path(rel: str) -> str:
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel)

def start_gui():
    root = Tk()
    root.title("ProxyCardMaker")
    # Try to set main window icon from app.ico (works in PyInstaller via resource_path)
    try:
        root.iconbitmap(resource_path("app.ico"))
    except Exception:
        try:
            from tkinter import PhotoImage
            root.iconphoto(True, PhotoImage(file=resource_path("app.ico")))
        except Exception:
            pass
    try:
        from tkinter import font as tkfont
        tkfont.nametofont("TkDefaultFont").configure(family="Segoe UI", size=9)
    except Exception:
        pass

    root.minsize(760, 520)
    root.grid_rowconfigure(1, weight=1)
    root.grid_columnconfigure(0, weight=0)
    RIGHT_MIN = 430
    root.grid_columnconfigure(1, weight=0, minsize=RIGHT_MIN)

    Label(root, text="Enter your deck list:", font=("Segoe UI", 9, "bold")).grid(
        row=0, column=0, sticky="w", padx=2, pady=(3, 0)
    )

    right = Frame(root, bd=0, highlightthickness=0)
    right.grid(row=0, column=1, rowspan=2, padx=0, pady=(3,6), sticky="n")

    lf_padx, lf_pady = 4, 1
    row_gap = dict(pady=(1,0))

    source_box = LabelFrame(right, text="Source & Local folder", padx=lf_padx, pady=lf_pady, bd=1)
    source_box.grid(row=0, column=0, sticky="ew", **row_gap)
    source_var = StringVar(value="local")
    Radiobutton(source_box, text="OnePiece.gg",  variable=source_var, value="dotgg", anchor="w").grid(row=0, column=0, sticky="w")
    Radiobutton(source_box, text="limitlesstcg", variable=source_var, value="limitless", anchor="w").grid(row=1, column=0, sticky="w")
    Radiobutton(source_box, text="Local folder", variable=source_var, value="local", anchor="w").grid(row=2, column=0, sticky="w")
    Label(source_box, text="Folder:").grid(row=3, column=0, sticky="w", pady=(1,0))
    local_dir_var = StringVar(value=r"C:\User\Desktop\folder")
    Entry(source_box, textvariable=local_dir_var, width=36).grid(row=4, column=0, sticky="w")
    Button(source_box, text="Browse", width=6,
           command=lambda: local_dir_var.set(filedialog.askdirectory() or local_dir_var.get())
    ).grid(row=4, column=1, padx=(4,0), sticky="e")

    save_box = LabelFrame(right, text="Save folder", padx=lf_padx, pady=lf_pady, bd=1)
    save_box.grid(row=1, column=0, sticky="ew", **row_gap)
    folder_var = StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "KartenDownloader"))
    Entry(save_box, textvariable=folder_var, width=36).grid(row=0, column=0, sticky="w")
    Button(save_box, text="Browse", width=6,
           command=lambda: folder_var.set(filedialog.askdirectory() or folder_var.get())
    ).grid(row=0, column=1, padx=(4,0), sticky="e")

    out_box = LabelFrame(right, text="Output mode", padx=lf_padx, pady=lf_pady, bd=1)
    out_box.grid(row=2, column=0, sticky="ew", **row_gap)
    out_mode = StringVar(value="a4sheet")
    Radiobutton(out_box, text="Save individual images (PNG)", variable=out_mode, value="images", anchor="w").grid(row=0, column=0, sticky="w")
    row_a4 = Frame(out_box); row_a4.grid(row=1, column=0, sticky="w")
    Radiobutton(row_a4, text="Save A4 sheet 3×3", variable=out_mode, value="a4sheet", anchor="w").pack(side="left")
    Label(row_a4, text="Format:").pack(side="left", padx=(8,2))
    a4_fmt = StringVar(value="PDF")
    OptionMenu(row_a4, a4_fmt, "PDF", "PNG", "JPG").pack(side="left")
    dpi_row = Frame(out_box); dpi_row.grid(row=2, column=0, sticky="w", pady=(0,0))
    Label(dpi_row, text="DPI (applies to both):").pack(side="left")
    dpi_var = IntVar(value=1500)
    Entry(dpi_row, textvariable=dpi_var, width=5).pack(side="left", padx=(4,0))

    a4_box = LabelFrame(right, text="A4 layout (mm)", padx=max(lf_padx-2,0), pady=0, bd=1)
    a4_box.grid(row=3, column=0, sticky="ew", **row_gap)
    Label(a4_box, text="Card W×H").grid(row=0, column=0, sticky="w")
    card_w_mm = IntVar(value=63); card_h_mm = IntVar(value=88)
    row_wh = Frame(a4_box); row_wh.grid(row=1, column=0, columnspan=2, sticky="w")
    Entry(row_wh, textvariable=card_w_mm, width=4).pack(side="left")
    Entry(row_wh, textvariable=card_h_mm, width=4).pack(side="left", padx=(1,0))
    Label(a4_box, text="Margins L/R,T/B").grid(row=2, column=0, sticky="w")
    margin_x_mm = IntVar(value=7); margin_y_mm = IntVar(value=13)
    row_mg = Frame(a4_box); row_mg.grid(row=3, column=0, columnspan=2, sticky="w")
    Entry(row_mg, textvariable=margin_x_mm, width=4).pack(side="left")
    Entry(row_mg, textvariable=margin_y_mm, width=4).pack(side="left", padx=(1,0))
    Label(a4_box, text="Gaps H,V").grid(row=4, column=0, sticky="w")
    gap_x_mm = IntVar(value=3); gap_y_mm = IntVar(value=3)
    row_gp = Frame(a4_box); row_gp.grid(row=5, column=0, columnspan=2, sticky="w")
    Entry(row_gp, textvariable=gap_x_mm, width=4).pack(side="left")
    Entry(row_gp, textvariable=gap_y_mm, width=4).pack(side="left", padx=(1,0))

    crop_box = Frame(a4_box); crop_box.grid(row=6, column=0, columnspan=2, sticky="w", pady=(2,0))
    crop_var = BooleanVar(value=True)
    Checkbutton(crop_box, text="Add corner crop marks", variable=crop_var).pack(side="left")
    Label(crop_box, text="len(mm)").pack(side="left", padx=(6,2))
    crop_len_var = IntVar(value=300)
    Entry(crop_box, textvariable=crop_len_var, width=3).pack(side="left")
    Label(crop_box, text="gap(mm)").pack(side="left", padx=(6,2))
    crop_gap_var = IntVar(value=0)
    Entry(crop_box, textvariable=crop_gap_var, width=3).pack(side="left")

    img_box = LabelFrame(right, text="Image processing", padx=lf_padx, pady=lf_pady, bd=1)
    img_box.grid(row=4, column=0, sticky="ew", **row_gap)
    border_var = BooleanVar(value=False); border_px_var = IntVar(value=50)
    def on_toggle_border(): border_entry.configure(state=(NORMAL if border_var.get() else DISABLED))
    Checkbutton(img_box, text="Add white border", variable=border_var, command=on_toggle_border).grid(row=0, column=0, sticky="w")
    row_b = Frame(img_box); row_b.grid(row=1, column=0, sticky="w")
    Label(row_b, text="Border px:").pack(side="left")
    border_entry = Entry(row_b, textvariable=border_px_var, width=5); border_entry.pack(side="left", padx=(4,0))
    on_toggle_border()
    upscale_var = BooleanVar(value=False); min_height_var = IntVar(value=1500)
    row_u = Frame(img_box); row_u.grid(row=2, column=0, sticky="w", pady=(1,0))
    Checkbutton(row_u, text="Upscale (min h)", variable=upscale_var).pack(side="left")
    Entry(row_u, textvariable=min_height_var, width=7).pack(side="left", padx=(6,0))

    opt_box = LabelFrame(right, text="Options", padx=lf_padx, pady=lf_pady, bd=1)
    opt_box.grid(row=5, column=0, sticky="ew", **row_gap)
    multiply_var = BooleanVar(value=True); choose_art_var = BooleanVar(value=True); overwrite_var = BooleanVar(value=False)
    Checkbutton(opt_box, text="Download cards multiply", variable=multiply_var).grid(row=0, column=0, sticky="w")
    Checkbutton(opt_box, text="I want to select picture art", variable=choose_art_var).grid(row=1, column=0, sticky="w")
    Checkbutton(opt_box, text="Overwrite existing files", variable=overwrite_var).grid(row=2, column=0, sticky="w")

    status_label = Label(right, text="", fg="#007a33"); status_label.grid(row=6, column=0, sticky="w", pady=(1,2))
    Button(right, text="Download Deck", command=lambda: on_download(), width=30).grid(row=7, column=0, sticky="ew")

    left = Frame(root, bd=0, highlightthickness=0)
    left.grid(row=1, column=0, sticky="nsw", padx=2, pady=(2,4))
    text_frame = Frame(left, bd=0, highlightthickness=0); text_frame.grid(row=0, column=0, sticky="nsw")
    text_box = Text(text_frame, height=20, width=36); text_box.grid(row=0, column=0, sticky="n")
    text_scroll = Scrollbar(text_frame, command=text_box.yview); text_scroll.grid(row=0, column=1, sticky="ns")
    text_box.config(yscrollcommand=text_scroll.set)
    example = ("Example:\n"
               "1xOP11-040\n"
               "2xOP05-067\n"
               "4xST18-001\n"
               "4xEB01-061\n"
               "4xOP10-072\n"
               "2xOP07-064\n"
               "51xCard_back\n"
               "3xKuzan")
    Label(left, text=example, justify="left", fg="#666").grid(row=1, column=0, sticky="nw", pady=(2,0))

    def normalize_folder(p: str) -> str:
        p = p.strip().strip('"').strip("'"); return os.path.normpath(os.path.abspath(p))

    def ensure_unique_suffix(path: str) -> str:
        base, ext = os.path.splitext(path); candidate = path; n = 1
        while os.path.exists(candidate): candidate = f"{base} ({n}){ext}"; n += 1
        return candidate
    def ensure_unique(path: str) -> str: return path if overwrite_var.get() else ensure_unique_suffix(path)

    def on_download():
        lines = text_box.get("1.0", END).strip().splitlines()
        if not lines:
            messagebox.showerror("Error", "Please enter a deck list!"); return
        save_dir_raw = folder_var.get().strip()
        if not save_dir_raw:
            messagebox.showerror("Error", "Please select a save folder!"); return
        target_dir = normalize_folder(save_dir_raw); os.makedirs(target_dir, exist_ok=True)

        success, failed = [], []
        total = len([ln for ln in lines if ln.strip()])
        src = source_var.get()
        local_dir = normalize_folder(local_dir_var.get()) if src == "local" else None
        collected_for_a4: List[bytes] = []
        first_title_for_sheet: Optional[str] = None

        for idx, line in enumerate(lines, start=1):
            txt = line.strip()
            if not txt:
                continue
            m = re.match(r'\s*(\d+)\s*x\s*(.+?)\s*$', txt, re.I)
            if m:
                qty = int(m.group(1)); term = m.group(2).strip()
            else:
                qty = 1; term = txt
            is_code = looks_like_code(term)
            display = term.upper() if is_code else term
            if not first_title_for_sheet:
                first_title_for_sheet = display

            effective_qty = qty if multiply_var.get() else 1
            status_label.config(text=f"Processing {idx}/{total}: {display} …"); right.update_idletasks()

            try:
                if choose_art_var.get():
                    if is_code:
                        variants = probe_all_arts(term.upper(), src, local_dir)
                    else:
                        if src != "local":
                            raise RuntimeError("Free-text search is only supported with 'Local folder'.")
                        paths = candidates_local_substr(term, local_dir or ".")
                        variants = [(p, open(p, "rb").read()) for p in paths]
                    if not variants:
                        raise RuntimeError(f"No images found for {display}.")
                    chosen = variants[0][1] if len(variants) == 1 else (pick_art_popup(root, display, variants) or None)
                    if chosen is None:
                        continue
                    img_bytes = chosen
                else:
                    if is_code:
                        img_bytes = download_card_default(term.upper(), src, local_dir)
                    else:
                        if src != "local":
                            raise RuntimeError("Free-text search is only supported with 'Local folder'.")
                        paths = candidates_local_substr(term, local_dir or ".")
                        if not paths:
                            raise RuntimeError(f"No images found for '{display}'.")
                        with open(paths[0], "rb") as f:
                            img_bytes = f.read()

                if out_mode.get() == "images":
                    for i in range(effective_qty):
                        base = f"{display}_{i+1}" if effective_qty > 1 else display
                        safe = re.sub(r'[^A-Za-z0-9_-]+', '_', base)[:60]
                        out_file = ensure_unique(os.path.join(target_dir, f"{safe}.png"))
                        save_png(img_bytes, out_file,
                                 add_border=border_var.get(),
                                 border_px=max(0, int(border_px_var.get())),
                                 do_upscale=upscale_var.get(),
                                 min_height_px=max(1, int(min_height_var.get())),
                                 dpi=max(72, int(dpi_var.get())))
                        success.append(out_file)
                else:
                    for _ in range(effective_qty):
                        collected_for_a4.append(img_bytes)

            except Exception as e:
                failed.append(f"{qty}x{display}")
                print(f"[ERROR] {display}: {e}")

        if out_mode.get() == "a4sheet" and collected_for_a4:
            dpi = max(72, int(dpi_var.get()))
            pages = build_a4_pages(
                collected_for_a4, dpi=dpi,
                card_w_mm=float(card_w_mm.get()), card_h_mm=float(card_h_mm.get()),
                margin_x_mm=float(margin_x_mm.get()), margin_y_mm=float(margin_y_mm.get()),
                gap_x_mm=float(gap_x_mm.get()), gap_y_mm=float(gap_y_mm.get()),
                crop_marks=bool(crop_var.get()),
                crop_len_mm=float(crop_len_var.get()), crop_gap_mm=float(crop_gap_var.get())
            )
            if pages:
                base_name = re.sub(r'[^A-Za-z0-9_-]+', '_', first_title_for_sheet or "sheet")
                fmt = a4_fmt.get().upper()
                if fmt == "PDF":
                    path = next_unique(target_dir, f"A4_{base_name}", ".pdf")
                    first, rest = pages[0].convert("RGB"), [p.convert("RGB") for p in pages[1:]]
                    first.save(path, "PDF", resolution=dpi, save_all=True, append_images=rest)
                    success.append(path)
                else:
                    for i, p in enumerate(pages, 1):
                        ext = ".png" if fmt == "PNG" else ".jpg"
                        path = next_unique(target_dir, f"A4_{base_name}_{i:03d}", ext)
                        if fmt == "PNG":
                            p.save(path, "PNG", optimize=True)
                        else:
                            p.save(path, "JPEG", quality=95, subsampling=0, optimize=True)
                        success.append(path)

        status_label.config(text="Download finished.")
        msg = f"{len(success)} file(s) saved in:\\n{target_dir}"
        if failed: msg += f"\\nFailed: {', '.join(failed)}"
        messagebox.showinfo("Done", msg)

    root.mainloop()

if __name__ == "__main__":
    try:
        start_gui()
    except Exception as e:
        messagebox.showerror("Fatal Error", str(e))
