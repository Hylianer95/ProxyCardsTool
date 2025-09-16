# main_multi_game_fixed10.py
import os, re, math, sys, html
from typing import Optional, List, Tuple, Dict
from io import BytesIO
from tkinter import (
    Tk, Label, Button, Text, Scrollbar, filedialog, messagebox,
    Entry, StringVar, END, BooleanVar, Toplevel, Canvas, Frame,
    IntVar, Radiobutton, DISABLED, NORMAL, LabelFrame, Checkbutton, OptionMenu
)

import requests
import webbrowser
import json
import glob
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from concurrent.futures import ThreadPoolExecutor, as_completed
from PIL import Image, ImageTk, ImageOps, ImageFilter, Image as PILImage, ImageDraw

DOTGG_BASE = "https://static.dotgg.gg/onepiece/card/"
LIMITLESS_BASE = "https://limitlesstcg.nyc3.cdn.digitaloceanspaces.com/one-piece/"
YGOPRODECK_API = "https://db.ygoprodeck.com/api/v7/cardinfo.php"
SCRYFALL_NAMED = "https://api.scryfall.com/cards/named"
SCRYFALL_SEARCH = "https://api.scryfall.com/cards/search"
PKMNCARDS_SEARCH = "https://pkmncards.com/?s="

MAX_PARALLEL, TIMEOUT_SEC = 16, 10
THUMB_W, THUMB_H, THUMB_COLS = 150, 210, 6
P_MIN, P_MAX = 1, 10
IMAGE_EXTS = {".png",".webp",".jpg",".jpeg",".bmp"}

GAMES = ["One Piece", "Yu-Gi-Oh!", "Pokémon", "MTG"]
SOURCES_BY_GAME: Dict[str, List[Tuple[str,str]]] = {
    "One Piece": [("dotgg","OnePiece.gg"), ("limitless","limitlesstcg"), ("local","Local folder")],
    "Yu-Gi-Oh!": [("ygoprodeck","YGOPRODeck"), ("local","Local folder")],
    "Pokémon":   [("pkmncards","PKMNCards"), ("local","Local folder")],
    "MTG":       [("scryfall","Scryfall"), ("local","Local folder")],
}

def make_session() -> requests.Session:
    s = requests.Session()
    retries = Retry(total=2, backoff_factor=0.3, status_forcelist=(429,500,502,503,504),
                    allowed_methods=frozenset(["GET"]))
    adapter = HTTPAdapter(pool_connections=MAX_PARALLEL, pool_maxsize=MAX_PARALLEL, max_retries=retries)
    s.mount("http://", adapter); s.mount("https://", adapter)
    s.headers.update({"User-Agent":"MultiGame-ProxyMaker/1.4 (+Tkinter)",
                      "Accept":"text/html,application/json,image/webp,image/*;q=0.8,*/*;q=0.5"})
    return s
SESSION = make_session()

def http_get(url: str, params: Optional[dict]=None) -> Optional[requests.Response]:
    try:
        r = SESSION.get(url, timeout=TIMEOUT_SEC, params=params)
        if r.status_code == 200: return r
    except Exception: pass
    return None

def request_ok(url: str, params: Optional[dict]=None) -> Optional[bytes]:
    r = http_get(url, params=params); return (r.content if r and r.content else None)

def request_ok_pkmn(url: str) -> Optional[bytes]:
    """Download with Referer to avoid CDN anti-hotlink 1x1 thumbnails."""
    try:
        r = SESSION.get(url, timeout=TIMEOUT_SEC, headers={'Referer': 'https://pkmncards.com/'})
        if r.status_code == 200 and r.content:
            return r.content
    except Exception:
        pass
    return None

# -------- Local file search (all games) --------
def candidates_local_by_code_or_name(term: str, root_dir: str) -> List[str]:
    needle = term.strip().upper()
    hits: List[str] = []
    for base, _dirs, files in os.walk(root_dir):
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext in IMAGE_EXTS and needle in name.upper():
                hits.append(os.path.join(base, name))
    hits.sort(key=lambda p: os.path.basename(p).upper())
    return hits

# -------- One Piece (code-based) --------
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

def probe_urls_in_parallel(urls: List[str]) -> List[Tuple[str, bytes]]:
    results: List[Optional[Tuple[str, bytes]]] = [None] * len(urls)
    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as pool:
        futmap = {pool.submit(request_ok, u): (i, u) for i, u in enumerate(urls)}
        for fut in as_completed(futmap):
            i, u = futmap[fut]; data = fut.result()
            if data: results[i] = (u, data)
    return [x for x in results if x is not None]

# -------- Yu-Gi-Oh! --------

# -------- Yu-Gi-Oh! --------
def fetch_ygo_images(name_term: str) -> List[Tuple[str, bytes]]:
    out: List[Tuple[str, bytes]] = []
    name_clean = name_term.strip()

    def collect_from_card(card: dict):
        imgs = card.get("card_images", []) or []
        for ci in imgs:
            url = ci.get("image_url") or ci.get("image_url_small")
            if not url: continue
            b = request_ok(url)
            if b: out.append((url, b))
        try:
            cid = card.get("id"); nm = str(card.get("name",""))
            if cid and nm:
                slug = re.sub(r'[^a-z0-9]+', '-', nm.lower()); slug = re.sub(r'-+', '-', slug).strip('-')
                rp = http_get(f"https://ygoprodeck.com/card/{slug}-{cid}")
                if rp and rp.text:
                    html_text = rp.text
                    for m in re.finditer(r'<img[^>]+class\s*=\s*(?:"|\')[^"\']*variant-artwork[^"\']*(?:"|\')[^>]*>', html_text, flags=re.I|re.DOTALL):
                        tag = m.group(0)
                        msrc = re.search(r'\s(?:src|data-src|data-lazy-src)\s*=\s*(?:"([^"]+)"|\'([^\']+)\')', tag, flags=re.I)
                        url2 = msrc.group(1) if (msrc and msrc.group(1) is not None) else (msrc.group(2) if msrc else None)
                        if not url2:
                            mset = re.search(r'\s(?:srcset|data-srcset)\s*=\s*"([^"]+)"', tag, flags=re.I)
                            if mset: url2 = mset.group(1).split(",")[0].split()[0]
                        if not url2: continue
                        b2 = request_ok(url2)
                        if b2: out.append((url2, b2))
        except Exception: pass

    r = http_get(YGOPRODECK_API, params={"name": name_clean})
    if r:
        try:
            data = r.json(); cards = data.get("data") or []
            for card in cards:
                if str(card.get("name","")).lower() == name_clean.lower():
                    collect_from_card(card)
                    if out: return out
        except Exception: pass

    r = http_get(YGOPRODECK_API, params={"fname": name_clean})
    if r:
        try:
            data = r.json(); cards = data.get("data", [])
            if not cards: return out
            exact = None
            for c in cards:
                if str(c.get("name","")).lower() == name_clean.lower():
                    exact = c; break
            collect_from_card(exact or cards[0])
        except Exception: pass
    return out

# -------- MTG (Scryfall prints) --------
def fetch_mtg_images(name_term: str) -> List[Tuple[str, bytes]]:
    out: List[Tuple[str, bytes]] = []
    query = f'!"{name_term}" include:extras'
    params = {"q": query, "unique": "prints", "order": "released"}
    r = http_get(SCRYFALL_SEARCH, params=params)
    def add_from_list(obj):
        for card in obj.get("data", []):
            uris = card.get("image_uris") or {}
            url = uris.get("png") or uris.get("large") or uris.get("normal") or uris.get("small")
            if not url and card.get("card_faces"):
                for face in card["card_faces"]:
                    uris = face.get("image_uris") or {}
                    url = uris.get("png") or uris.get("large") or uris.get("normal") or uris.get("small")
                    if url:
                        b = request_ok(url)
                        if b: out.append((url, b))
                        url = None
                        break
                continue
            if url:
                b = request_ok(url)
                if b: out.append((url, b))
    if r:
        try:
            j = r.json()
            add_from_list(j)
            while j.get("has_more") and len(out) < 60:
                next_url = j.get("next_page")
                if not next_url: break
                r2 = http_get(next_url)
                if not r2: break
                j = r2.json()
                add_from_list(j)
        except Exception:
            pass
    if not out:
        r = http_get(SCRYFALL_NAMED, params={"fuzzy": name_term})
        if r:
            try:
                j = r.json()
                uris = j.get("image_uris") or {}
                for key in ("png","large","normal","small"):
                    if uris.get(key):
                        b = request_ok(uris[key])
                        if b: out.append((uris[key], b)); break
                if not out and j.get("card_faces"):
                    for face in j["card_faces"]:
                        uris = face.get("image_uris") or {}
                        for key in ("png","large","normal","small"):
                            if uris.get(key):
                                b = request_ok(uris[key])
                                if b: out.append((uris[key], b))
            except Exception:
                pass
    return out

# -------- Pokémon (PKMNCards, multi-art with filtering) --------
def _first_href_in_search(html_text: str, term: str) -> Optional[str]:
    m = re.search(r'<a\s+href="([^"]+)"[^>]*class="[^"]*entry-title-link[^"]*"[^>]*>', html_text, re.I)
    if m: return html.unescape(m.group(1))
    m = re.search(r'<a\s+href="([^"]+)"[^>]*rel="bookmark"[^>]*>', html_text, re.I)
    if m: return html.unescape(m.group(1))
    m = re.search(r'<a\s+href="(https?://[^"]+/card/[^"]+)"', html_text, re.I)
    if m: return html.unescape(m.group(1))
    safe = re.escape(term.strip())
    m = re.search(rf'<a\s+href="([^"]*{safe}[^"]*)"', html_text, re.I)
    return html.unescape(m.group(1)) if m else None

def _first_upload_image(html_text: str) -> Optional[str]:
    m = re.search(r'(https?://[^"]+/wp-content/uploads/[^"]+\.(?:png|jpg|jpeg|webp))', html_text, re.I)
    return html.unescape(m.group(1)) if m else None

def _og_image(html_text: str) -> Optional[str]:
    m = re.search(r'<meta\s+property="og:image"\s+content="([^"]+)"', html_text, re.I)
    if m: return html.unescape(m.group(1))
    m = re.search(r"<meta\s+property='og:image'\s+content='([^']+)'", html_text, re.I)
    if m: return html.unescape(m.group(1))
    return None

def fetch_pokemon_images(name_term: str) -> List[Tuple[str, bytes]]:
    """Collect images from PKMNCards search grid.
    - Strict: ALL keywords in the query must appear in the filename/URL (order-free).
    - Convert thumbnail URLs (…-150x150.jpg) to original full images by stripping the -WxH suffix.
    - Download with Referer to avoid 1x1 anti-hotlink placeholders.
    - Fallback: first result page og:image/upload.
    """
    out: List[Tuple[str, bytes]] = []
    term = name_term.strip().lower()
    # Build keywords from the query (keep short tokens like 'ex', 'gx', 'us', 'promo')
    keywords = re.findall(r'[a-z0-9]+', term)

    r = http_get(PKMNCARDS_SEARCH + requests.utils.quote(term))
    if not r:
        return out
    html_text = r.text

    # Grab all candidate image URLs from the search page
    urls = re.findall(r'(https?://pkmncards\.com/wp-content/uploads/[^"]+\.(?:png|jpg|jpeg|webp))', html_text, re.I)
    seen = set()
    filtered: List[str] = []

    def canonicalize(u: str) -> str:
        # Turn ...-200x300.jpg into ....jpg (full image)
        return re.sub(r'-(\d+)x(\d+)(\.(?:png|jpe?g|webp))$', r'\3', u, flags=re.I)

    for u in urls:
        base = os.path.basename(u).lower()
        if 'cropped-' in base or base.startswith('crop-') or base.startswith('cropped-'):
            continue
        # Normalize and check ALL keywords appear
        base_norm = re.sub(r'[^a-z0-9]+', '', base)
        if keywords and not all(k in base_norm for k in keywords):
            continue
        u_full = canonicalize(u)
        if u_full not in seen:
            seen.add(u_full)
            filtered.append(u_full)

    # Fetch images with Referer header
    for u in filtered[:60]:
        b = request_ok_pkmn(u)
        if b:
            out.append((u, b))

    # Fallback: open first result page and try main image
    if not out:
        href = _first_href_in_search(html_text, term)
        if href:
            r2 = http_get(href)
            if r2:
                img = _first_upload_image(r2.text) or _og_image(r2.text)
                if img:
                    img = canonicalize(img)
                    b = request_ok_pkmn(img)
                    if b:
                        out.append((img, b))
    return out

# -------- Image utils --------
def make_padded_thumb(img_bytes: bytes, w: int, h: int) -> ImageTk.PhotoImage:
    with PILImage.open(BytesIO(img_bytes)) as im:
        im = im.convert("RGBA"); im.thumbnail((w, h), PILImage.LANCZOS)
        canvas = PILImage.new("RGBA", (w, h), (255, 255, 255, 0))
        x = (w - im.width) // 2; y = (h - im.height) // 2
        canvas.paste(im, (x, y))
        return ImageTk.PhotoImage(canvas)

def save_png(image_bytes: bytes, out_path: str,
             add_border: bool = True, border_px: int = 50, border_color: str = 'white',
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
            im = ImageOps.expand(im, border=border_px, fill=border_color)
        im.save(out_path, "PNG", optimize=True, dpi=(dpi, dpi))

def mm_to_px(mm: float, dpi: int) -> int:
    return int(round(mm / 25.4 * dpi))

def draw_crop_marks(draw: ImageDraw.ImageDraw, x: int, y: int, w: int, h: int,
                    length_px: int, gap_px: int, stroke_px: int = 1, color=(0,0,0),
                    border_px: int = 0, hide_under_border: bool = False) -> None:
    """
    Draws external corner crop marks around the card rectangle.
    If hide_under_border=True and border_px>0, segments are CLIPPED so they only appear
    in the future border area and are therefore fully covered by the border.
    """
    s = max(1, int(stroke_px)); L = max(0, int(length_px)); G = max(0, int(gap_px))
    Xl = int(x); Xr = int(x + w); Yt = int(y); Yb = int(y + h)

    # Helper: draw solid segments as rectangles (pixel-perfect)
    def hseg(x1, x2, yy):
        if x2 <= x1: return
        y1 = yy - (s // 2); y2 = y1 + s - 1
        draw.rectangle([int(x1), int(y1), int(x2), int(y2)], fill=color)
    def vseg(xx, y1, y2):
        if y2 <= y1: return
        x1 = xx - (s // 2); x2 = x1 + s - 1
        draw.rectangle([int(x1), int(y1), int(x2), int(y2)], fill=color)

    # Compute unclipped segments (left/top corner shown; others analogous)
    # Top-left horizontal: y = Yt - G, x in [Xl - G - L, Xl - G]
    # Top-left vertical:   x = Xl - G, y in [Yt - G - L, Yt - G]
    segs = [
        ("h", Xl - G - L, Xl - G, Yt - G),          # TL horizontal
        ("v", Xl - G, Yt - G - L, Yt - G),          # TL vertical
        ("h", Xr + G, Xr + G + L, Yt - G),          # TR horizontal
        ("v", Xr + G, Yt - G - L, Yt - G),          # TR vertical
        ("h", Xl - G - L, Xl - G, Yb + G),          # BL horizontal
        ("v", Xl - G, Yb + G, Yb + G + L),          # BL vertical
        ("h", Xr + G, Xr + G + L, Yb + G),          # BR horizontal
        ("v", Xr + G, Yb + G, Yb + G + L),          # BR vertical
    ]

    if hide_under_border and border_px > 0:
        # Clamp to the border "ring" around the card: expand by border_px
        left_min  = Xl - border_px
        left_max  = Xl
        right_min = Xr
        right_max = Xr + border_px
        top_min   = Yt - border_px
        top_max   = Yt
        bot_min   = Yb
        bot_max   = Yb + border_px

        clipped = []
        for kind, a1, a2, b in segs:
            if kind == "h":
                # Horizontal segment at y=b, range [a1, a2]
                # Keep only if b lies within top/bottom border stripe
                if top_min <= b <= top_max:
                    # Clamp x to left ring (<=Xl) or right ring (>=Xr)
                    x1 = max(a1, left_min)
                    x2 = min(a2, left_max)
                    if x2 > x1: clipped.append(("h", x1, x2, b))
                if bot_min <= b <= bot_max:
                    x1 = max(a1, right_min)
                    x2 = min(a2, right_max)
                    if x2 > x1: clipped.append(("h", x1, x2, b))
            else: # "v"
                # Vertical segment at x=a1, range [a2, b] => we stored (x, y1, y2)
                x = a1; y1 = a2; y2 = b
                if left_min <= x <= left_max:
                    yy1 = max(y1, top_min)
                    yy2 = min(y2, top_max)
                    if yy2 > yy1: clipped.append(("v", x, yy1, yy2))
                if right_min <= x <= right_max:
                    yy1 = max(y1, bot_min)
                    yy2 = min(y2, bot_max)
                    if yy2 > yy1: clipped.append(("v", x, yy1, yy2))
        segs = clipped

    # Draw
    for kind, a1, a2, b in segs:
        if kind == "h":
            hseg(a1, a2, b)
        else:
            vseg(a1, a2, b)



def build_a4_pages(images: List[bytes], dpi: int,
                   card_w_mm: float, card_h_mm: float,
                   margin_x_mm: float, margin_y_mm: float,
                   gap_x_mm: float, gap_y_mm: float,
                   crop_marks: bool = False, crop_len_mm: float = 2.5, crop_gap_mm: float = 0.8,
                   crop_stroke_px: int = 1, crop_color=(0,0,0),
                   add_border: bool = False, border_px: int = 0, border_color: str = 'white') -> List[PILImage.Image]:
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
        positions = []
        # collect positions for this page
        for i, img_bytes in enumerate(chunk):
            r = i // cols; c = i % cols
            x = margin_x + c * (card_w + gap_x)
            y = margin_y + r * (card_h + gap_y)
            positions.append((i, img_bytes, x, y))

        # PASS 1 — draw all crop marks (under everything)
        if crop_marks:
            for _, __, x, y in positions:
                draw_crop_marks(draw, x, y, card_w, card_h, crop_len, crop_gap, crop_stroke_px, color=crop_color)

        # PASS 2 — paste all images on top (with border applied and offset)
        for _, img_bytes, x, y in positions:
            with PILImage.open(BytesIO(img_bytes)) as im:
                im = im.convert("RGB").resize((card_w, card_h), PILImage.LANCZOS)
                off = 0
                if add_border and border_px > 0:
                    im = ImageOps.expand(im, border=border_px, fill=border_color)
                    off = border_px
                page.paste(im, (x - off, y - off))
        pages.append(page)
    return pages

def ensure_unique_suffix(path: str) -> str:
    base, ext = os.path.splitext(path); candidate = path; n = 1
    while os.path.exists(candidate): candidate = f"{base} ({n}){ext}"; n += 1
    return candidate

def next_unique(target_dir: str, name: str, ext: str) -> str:
    return ensure_unique_suffix(os.path.join(target_dir, f"{name}{ext}"))

def looks_like_op_code(s: str) -> bool:
    return re.fullmatch(r"[A-Z]+\d{2}-\d{3}", s.strip().upper()) is not None

def probe_all_arts(game: str, term: str, source: str, local_dir: Optional[str] = None) -> List[Tuple[str, bytes]]:
    term = term.strip()
    if source == "local":
        if not local_dir or not os.path.isdir(local_dir): return []
        out: List[Tuple[str, bytes]] = []
        for p in candidates_local_by_code_or_name(term, local_dir):
            try:
                with open(p, "rb") as f: out.append((p, f.read()))
            except Exception: pass
        return out
    if game == "One Piece":
        if looks_like_op_code(term):
            urls = (candidates_dotgg(term) if source == "dotgg" else candidates_limitless(term))
            return probe_urls_in_parallel(urls)
        else:
            return []
    elif game == "Yu-Gi-Oh!":
        if source == "ygoprodeck": return fetch_ygo_images(term)
    elif game == "Pokémon":
        if source == "pkmncards": return fetch_pokemon_images(term)
    elif game == "MTG":
        if source == "scryfall": return fetch_mtg_images(term)
    return []

def download_card_default(game: str, term: str, source: str, local_dir: Optional[str]) -> bytes:
    term = term.strip()
    if source == "local":
        paths = candidates_local_by_code_or_name(term, local_dir or ".")
        if not paths: raise RuntimeError(f"No local image found for '{term}'.")
        with open(paths[0], "rb") as f: return f.read()
    if game == "One Piece":
        if not looks_like_op_code(term): raise RuntimeError("For One Piece online sources, please use a code like OP11-040.")
        urls = (candidates_dotgg(term) if source == "dotgg" else candidates_limitless(term))
        for url in urls:
            data = request_ok(url)
            if data: return data
        raise RuntimeError(f"No image found for code {term} on source '{source}'.")
    if game == "Yu-Gi-Oh!" and source == "ygoprodeck":
        imgs = fetch_ygo_images(term)
        if imgs: return imgs[0][1]
        raise RuntimeError(f"No image found for Yu-Gi-Oh! name '{term}'.")
    if game == "Pokémon" and source == "pkmncards":
        imgs = fetch_pokemon_images(term)
        if imgs: return imgs[0][1]
        raise RuntimeError(f"No image found for Pokémon name '{term}'.")
    if game == "MTG" and source == "scryfall":
        imgs = fetch_mtg_images(term)
        if imgs: return imgs[0][1]
        raise RuntimeError(f"No image found for MTG name '{term}'.")
    raise RuntimeError("Unsupported combination.")

def resource_path(rel: str) -> str:
    try:
        base = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel)


# --- Settings save/load as file (no profiles) ---
def _collect_settings_dict():
    try:
        data = {
            "game": game_var.get(),
            "source": source_var.get(),
            "local_dir": local_dir_var.get(),
            "save_folder": folder_var.get(),
            "out_mode": out_mode.get(),
            "a4_fmt": a4_fmt.get(),
            "dpi": int(dpi_var.get() or 300),
            "card_w_mm": int(card_w_mm.get() or 63),
            "card_h_mm": int(card_h_mm.get() or 88),
            "margin_x_mm": int(margin_x_mm.get() or 7),
            "margin_y_mm": int(margin_y_mm.get() or 13),
            "gap_x_mm": int(gap_x_mm.get() or 3),
            "gap_y_mm": int(gap_y_mm.get() or 3),
            "crop_enabled": bool(crop_var.get()),
            "crop_len_mm": int(crop_len_var.get() or 5),
            "crop_gap_mm": int(crop_gap_var.get() or 0),
            "crop_stroke_px": int(crop_stroke_px_var.get() or 1),
            "crop_color": crop_color_var.get(),
            "border": bool(border_var.get()),
            "border_px": int(border_px_var.get() or 0),
            "border_color": border_color_var.get(),
            "upscale": bool(upscale_var.get()),
            "min_height": int(min_height_var.get() or 1500),
            "multiply": bool(multiply_var.get()),
            "choose_art": bool(choose_art_var.get()),
            "overwrite": bool(overwrite_var.get()),
        }
    except Exception:
        data = {}
    return data

def _apply_settings_dict(data: dict) -> bool:
    try:
        if data.get("game") in GAMES:
            game_var.set(data["game"]); on_game_change()
        if data.get("source") in [v for (v, _l) in SOURCES_BY_GAME.get(game_var.get(), [])]:
            source_var.set(data["source"])
        local_dir_var.set(data.get("local_dir", local_dir_var.get()))
        folder_var.set(data.get("save_folder", folder_var.get()))
        out_mode.set(data.get("out_mode", out_mode.get()))
        a4_fmt.set(data.get("a4_fmt", a4_fmt.get()))
        for var, key in [(dpi_var,"dpi"), (card_w_mm,"card_w_mm"), (card_h_mm,"card_h_mm"),
                         (margin_x_mm,"margin_x_mm"), (margin_y_mm,"margin_y_mm"),
                         (gap_x_mm,"gap_x_mm"), (gap_y_mm,"gap_y_mm"),
                         (crop_len_var,"crop_len_mm"), (crop_gap_var,"crop_gap_mm"),
                         (border_px_var,"border_px"), (min_height_var,"min_height")]:
            try: var.set(int(data.get(key, var.get())))
            except Exception: pass
        for var, key in [(crop_var,"crop_enabled"), (border_var,"border"),
                         (upscale_var,"upscale"), (multiply_var,"multiply"),
                         (choose_art_var,"choose_art"), (overwrite_var,"overwrite")]:
            try: var.set(bool(data.get(key, var.get())))
            except Exception: pass
        try: crop_stroke_px_var.set(int(data.get("crop_stroke_px", crop_stroke_px_var.get())))
        except Exception: pass
        try: crop_color_var.set(data.get("crop_color", crop_color_var.get()))
        except Exception: pass
        try: border_color_var.set(data.get("border_color", border_color_var.get()))
        except Exception: pass
        return True
    except Exception:
        return False

def save_settings_as_dialog():
    from tkinter.filedialog import asksaveasfilename
    path = asksaveasfilename(title="Save settings as...",
                             defaultextension=".json",
                             filetypes=[("JSON files","*.json")])
    if not path: 
        return False
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(_collect_settings_dict(), f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

def load_settings_from_dialog():
    from tkinter.filedialog import askopenfilename
    path = askopenfilename(title="Load settings from...",
                           filetypes=[("JSON files","*.json")])
    if not path:
        return False
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return _apply_settings_dict(data)
    except Exception:
        return False
def start_gui():
    root = Tk()
    selected_profile_var = StringVar(value="default")
    root.title("ProxyCardsTool (PCT)")
    try:
        root.iconbitmap(resource_path("app.ico"))
    except Exception:
        try:
            from tkinter import PhotoImage
            root.iconphoto(True, PhotoImage(file=resource_path("app.ico")))
        except Exception: pass
    try:
        from tkinter import font as tkfont
        tkfont.nametofont("TkDefaultFont").configure(family="Segoe UI", size=9)
    except Exception: pass

    root.minsize(820, 560)
    root.grid_rowconfigure(1, weight=1)
    root.grid_columnconfigure(0, weight=0)
    RIGHT_MIN = 470
    root.grid_columnconfigure(1, weight=0, minsize=RIGHT_MIN)

    Label(root, text="Enter your list (codes or names):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=2, pady=(3, 0))

    right = Frame(root, bd=0, highlightthickness=0)
    right.grid(row=0, column=1, rowspan=2, padx=0, pady=(3,6), sticky="n")

    lf_padx, lf_pady = 4, 1
    row_gap = dict(pady=(1,0))

    source_box = LabelFrame(right, text="Game & Source", padx=lf_padx, pady=lf_pady, bd=1)
    source_box.grid(row=0, column=0, sticky="ew", **row_gap)
    source_box.grid_columnconfigure(0, weight=0)
    source_box.grid_columnconfigure(1, weight=0)

    game_var = StringVar(value=GAMES[0])
    Label(source_box, text="Game:").grid(row=0, column=0, sticky="w", padx=(0,0))
    OptionMenu(source_box, game_var, *GAMES).grid(row=0, column=0, sticky="w", padx=(60,0))

    Label(source_box, text="Source:").grid(row=1, column=0, sticky="w", pady=(3,0))
    source_var = StringVar(value=SOURCES_BY_GAME[game_var.get()][0][0])
    source_buttons_frame = Frame(source_box); source_buttons_frame.grid(row=2, column=0, columnspan=2, sticky="w")

    def render_source_radios():
        for w in list(source_buttons_frame.children.values()): w.destroy()
        for i, (val, label) in enumerate(SOURCES_BY_GAME[game_var.get()]):
            Radiobutton(source_buttons_frame, text=label, variable=source_var, value=val, anchor="w").grid(row=i, column=0, sticky="w")

    def on_game_change(*_):
        first_src = SOURCES_BY_GAME[game_var.get()][0][0]
        source_var.set(first_src); render_source_radios()
    render_source_radios(); game_var.trace_add("write", on_game_change)

    Label(source_box, text="Local folder (all games):").grid(row=3, column=0, columnspan=2, sticky="w", pady=(6,0))
    local_dir_var = StringVar(value=r"C:\User\Desktop\folder")
    Entry(source_box, textvariable=local_dir_var, width=36).grid(row=4, column=0, sticky="w")
    Button(source_box, text="Browse", width=6, command=lambda: local_dir_var.set(filedialog.askdirectory(initialdir=local_dir_var.get()))).grid(row=4, column=1, padx=(2,0), sticky="w")

    save_box = LabelFrame(right, text="Save folder", padx=lf_padx, pady=lf_pady, bd=1)
    save_box.grid(row=1, column=0, sticky="ew", **row_gap)
    folder_var = StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "Cards"))
    Entry(save_box, textvariable=folder_var, width=36).grid(row=0, column=0, sticky="w")
    Button(save_box, text="Browse", width=6, command=lambda: folder_var.set(filedialog.askdirectory() or folder_var.get())).grid(row=0, column=1, padx=(4,0), sticky="e")

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
    Label(crop_box, text="thick(px)").pack(side="left", padx=(6,2))
    crop_stroke_px_var = IntVar(value=1)
    Entry(crop_box, textvariable=crop_stroke_px_var, width=3).pack(side="left")
    Label(crop_box, text="color").pack(side="left", padx=(6,2))
    crop_color_var = StringVar(value="Black")
    OptionMenu(crop_box, crop_color_var, "Black", "Red", "Green", "Blue").pack(side="left")

    img_box = LabelFrame(right, text="Image processing", padx=lf_padx, pady=lf_pady, bd=1)
    img_box.grid(row=4, column=0, sticky="ew", **row_gap)
    border_var = BooleanVar(value=False); border_px_var = IntVar(value=50)
    def on_toggle_border(): border_entry.configure(state=(NORMAL if border_var.get() else DISABLED))
    Checkbutton(img_box, text="Add border", variable=border_var, command=on_toggle_border).grid(row=0, column=0, sticky="w")
    row_b = Frame(img_box); row_b.grid(row=1, column=0, sticky="w")
    Label(row_b, text="Border px:").pack(side="left")
    border_entry = Entry(row_b, textvariable=border_px_var, width=5); border_entry.pack(side="left", padx=(4,6))
    Label(row_b, text="color").pack(side="left")
    border_color_var = StringVar(value="white")
    OptionMenu(row_b, border_color_var, "white", "black").pack(side="left")
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
    Button(right, text="Download List", command=lambda: on_download(), width=30).grid(row=7, column=0, sticky="ew")
    # --- settings persistence (moved above first call) ---
    SETTINGS_FILE = os.path.join(os.path.expanduser("~"), "mg_pcm_settings.json")

    def save_settings_to_file(path: str = None):
        name = selected_profile_var.get() if "selected_profile_var" in globals() else DEFAULT_PROFILE
        return save_settings_to_profile(name)

    # --- settings profiles ---
    SETTINGS_DIR = os.path.join(os.path.expanduser("~"), "mg_pcm_profiles")
    DEFAULT_PROFILE = "default"
    PROFILE_EXT = ".json"
    LAST_FILE = os.path.join(SETTINGS_DIR, "_last_profile.json")

    def _ensure_settings_dir():
        try:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
        except Exception:
            pass

    def _sanitize_name(name: str) -> str:
        name = (name or "").strip()
        if not name:
            return DEFAULT_PROFILE
        for ch in '<>:"/\\|?*':
            name = name.replace(ch, '_')
        return name

    def _profile_path(name: str) -> str:
        _ensure_settings_dir()
        return os.path.join(SETTINGS_DIR, _sanitize_name(name) + PROFILE_EXT)

    def _list_profiles() -> list:
        _ensure_settings_dir()
        files = glob.glob(os.path.join(SETTINGS_DIR, "*" + PROFILE_EXT))
        names = [os.path.splitext(os.path.basename(f))[0] for f in files]
        if DEFAULT_PROFILE not in names:
            names.append(DEFAULT_PROFILE)
        return sorted(set(names))

    def _save_last_profile(name: str):
        try:
            _ensure_settings_dir()
            with open(LAST_FILE, "w", encoding="utf-8") as f:
                json.dump({"last": _sanitize_name(name)}, f)
        except Exception:
            pass

    def _load_last_profile() -> str:
        try:
            with open(LAST_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return _sanitize_name(data.get("last", DEFAULT_PROFILE))
        except Exception:
            return DEFAULT_PROFILE

    def save_settings_to_profile(profile_name: str) -> bool:
        path = _profile_path(profile_name)
        try:
            data = {
                "game": game_var.get(),
                "source": source_var.get(),
                "local_dir": local_dir_var.get(),
                "save_folder": folder_var.get(),
                "out_mode": out_mode.get(),
                "a4_fmt": a4_fmt.get(),
                "dpi": int(dpi_var.get() or 300),
                "card_w_mm": int(card_w_mm.get() or 63),
                "card_h_mm": int(card_h_mm.get() or 88),
                "margin_x_mm": int(margin_x_mm.get() or 7),
                "margin_y_mm": int(margin_y_mm.get() or 13),
                "gap_x_mm": int(gap_x_mm.get() or 3),
                "gap_y_mm": int(gap_y_mm.get() or 3),
                "crop_enabled": bool(crop_var.get()),
                "crop_len_mm": int(crop_len_var.get() or 5),
                "crop_gap_mm": int(crop_gap_var.get() or 0),
                "crop_stroke_px": int(crop_stroke_px_var.get() or 1),
                "crop_color": crop_color_var.get(),
                "border": bool(border_var.get()),
                "border_px": int(border_px_var.get() or 0),
                "border_color": border_color_var.get(),
                "upscale": bool(upscale_var.get()),
                "min_height": int(min_height_var.get() or 1500),
                "multiply": bool(multiply_var.get()),
                "choose_art": bool(choose_art_var.get()),
                "overwrite": bool(overwrite_var.get()),
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            _save_last_profile(profile_name)
            return True
        except Exception:
            return False

    def load_settings_from_profile(profile_name: str) -> bool:
        path = _profile_path(profile_name)
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            legacy = os.path.join(os.path.expanduser("~"), "mg_pcm_settings.json")
            if os.path.exists(legacy):
                try:
                    with open(legacy, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    return False
            else:
                return False

        try:
            if data.get("game") in GAMES:
                game_var.set(data["game"]); on_game_change()
            if data.get("source") in [v for (v, _l) in SOURCES_BY_GAME.get(game_var.get(), [])]:
                source_var.set(data["source"])
            local_dir_var.set(data.get("local_dir", local_dir_var.get()))
            folder_var.set(data.get("save_folder", folder_var.get()))
            out_mode.set(data.get("out_mode", out_mode.get()))
            a4_fmt.set(data.get("a4_fmt", a4_fmt.get()))
            for var, key in [(dpi_var,"dpi"), (card_w_mm,"card_w_mm"), (card_h_mm,"card_h_mm"),
                             (margin_x_mm,"margin_x_mm"), (margin_y_mm,"margin_y_mm"),
                             (gap_x_mm,"gap_x_mm"), (gap_y_mm,"gap_y_mm"),
                             (crop_len_var,"crop_len_mm"), (crop_gap_var,"crop_gap_mm"),
                             (border_px_var,"border_px"), (min_height_var,"min_height")]:
                try: var.set(int(data.get(key, var.get())))
                except Exception: pass
            for var, key in [(crop_var,"crop_enabled"), (border_var,"border"),
                             (upscale_var,"upscale"), (multiply_var,"multiply"),
                             (choose_art_var,"choose_art"), (overwrite_var,"overwrite")]:
                try: var.set(bool(data.get(key, var.get())))
                except Exception: pass
            try: crop_stroke_px_var.set(int(data.get("crop_stroke_px", crop_stroke_px_var.get())))
            except Exception: pass
            try: crop_color_var.set(data.get("crop_color", crop_color_var.get()))
            except Exception: pass
            try: border_color_var.set(data.get("border_color", border_color_var.get()))
            except Exception: pass
            _save_last_profile(profile_name)
            return True
        except Exception:
            return False

    def delete_settings_profile(profile_name: str) -> bool:
        try:
            path = _profile_path(profile_name)
            if os.path.exists(path):
                os.remove(path)
            return True
        except Exception:
            return False
    # --- settings profiles (helpers) ---
    SETTINGS_DIR = os.path.join(os.path.expanduser("~"), "mg_pcm_profiles")
    DEFAULT_PROFILE = "default"
    PROFILE_EXT = ".json"
    LAST_FILE = os.path.join(SETTINGS_DIR, "_last_profile.json")

    def _ensure_settings_dir():
        try:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
        except Exception:
            pass

    def _sanitize_name(name: str) -> str:
        name = (name or "").strip()
        if not name:
            return DEFAULT_PROFILE
        for ch in '<>:"/\\|?*':
            name = name.replace(ch, '_')
        return name

    def _profile_path(name: str) -> str:
        _ensure_settings_dir()
        return os.path.join(SETTINGS_DIR, _sanitize_name(name) + PROFILE_EXT)

    def _list_profiles() -> list:
        _ensure_settings_dir()
        files = glob.glob(os.path.join(SETTINGS_DIR, "*" + PROFILE_EXT))
        names = [os.path.splitext(os.path.basename(f))[0] for f in files]
        if DEFAULT_PROFILE not in names:
            names.append(DEFAULT_PROFILE)
        return sorted(set(names))

    def _save_last_profile(name: str):
        try:
            _ensure_settings_dir()
            with open(LAST_FILE, "w", encoding="utf-8") as f:
                json.dump({"last": _sanitize_name(name)}, f)
        except Exception:
            pass

    def _load_last_profile() -> str:
        try:
            with open(LAST_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return _sanitize_name(data.get("last", DEFAULT_PROFILE))
        except Exception:
            return DEFAULT_PROFILE

    def save_settings_to_profile(profile_name: str) -> bool:
        path = _profile_path(profile_name)
        try:
            data = {
                "game": game_var.get(),
                "source": source_var.get(),
                "local_dir": local_dir_var.get(),
                "save_folder": folder_var.get(),
                "out_mode": out_mode.get(),
                "a4_fmt": a4_fmt.get(),
                "dpi": int(dpi_var.get() or 300),
                "card_w_mm": int(card_w_mm.get() or 63),
                "card_h_mm": int(card_h_mm.get() or 88),
                "margin_x_mm": int(margin_x_mm.get() or 7),
                "margin_y_mm": int(margin_y_mm.get() or 13),
                "gap_x_mm": int(gap_x_mm.get() or 3),
                "gap_y_mm": int(gap_y_mm.get() or 3),
                "crop_enabled": bool(crop_var.get()),
                "crop_len_mm": int(crop_len_var.get() or 5),
                "crop_gap_mm": int(crop_gap_var.get() or 0),
                "crop_stroke_px": int(crop_stroke_px_var.get() or 1),
                "crop_color": crop_color_var.get(),
                "border": bool(border_var.get()),
                "border_px": int(border_px_var.get() or 0),
                "border_color": border_color_var.get(),
                "upscale": bool(upscale_var.get()),
                "min_height": int(min_height_var.get() or 1500),
                "multiply": bool(multiply_var.get()),
                "choose_art": bool(choose_art_var.get()),
                "overwrite": bool(overwrite_var.get()),
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            _save_last_profile(profile_name)
            return True
        except Exception:
            return False

    def load_settings_from_profile(profile_name: str) -> bool:
        path = _profile_path(profile_name)
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            legacy = os.path.join(os.path.expanduser("~"), "mg_pcm_settings.json")
            if os.path.exists(legacy):
                try:
                    with open(legacy, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    return False
            else:
                return False

        try:
            if data.get("game") in GAMES:
                game_var.set(data["game"]); on_game_change()
            if data.get("source") in [v for (v, _l) in SOURCES_BY_GAME.get(game_var.get(), [])]:
                source_var.set(data["source"])
            local_dir_var.set(data.get("local_dir", local_dir_var.get()))
            folder_var.set(data.get("save_folder", folder_var.get()))
            out_mode.set(data.get("out_mode", out_mode.get()))
            a4_fmt.set(data.get("a4_fmt", a4_fmt.get()))
            for var, key in [(dpi_var,"dpi"), (card_w_mm,"card_w_mm"), (card_h_mm,"card_h_mm"),
                             (margin_x_mm,"margin_x_mm"), (margin_y_mm,"margin_y_mm"),
                             (gap_x_mm,"gap_x_mm"), (gap_y_mm,"gap_y_mm"),
                             (crop_len_var,"crop_len_mm"), (crop_gap_var,"crop_gap_mm"),
                             (border_px_var,"border_px"), (min_height_var,"min_height")]:
                try: var.set(int(data.get(key, var.get())))
                except Exception: pass
            for var, key in [(crop_var,"crop_enabled"), (border_var,"border"),
                             (upscale_var,"upscale"), (multiply_var,"multiply"),
                             (choose_art_var,"choose_art"), (overwrite_var,"overwrite")]:
                try: var.set(bool(data.get(key, var.get())))
                except Exception: pass
            try: crop_stroke_px_var.set(int(data.get("crop_stroke_px", crop_stroke_px_var.get())))
            except Exception: pass
            try: crop_color_var.set(data.get("crop_color", crop_color_var.get()))
            except Exception: pass
            try: border_color_var.set(data.get("border_color", border_color_var.get()))
            except Exception: pass
            _save_last_profile(profile_name)
            return True
        except Exception:
            return False

    def delete_settings_profile(profile_name: str) -> bool:
        try:
            path = _profile_path(profile_name)
            if os.path.exists(path):
                os.remove(path)
            return True
        except Exception:
            return False





    def _on_close():
        try: save_settings_to_file()
        finally: root.destroy()
    root.protocol("WM_DELETE_WINDOW", _on_close)
    # auto-load on start
    load_settings_from_profile(selected_profile_var.get())

    left = Frame(root, bd=0, highlightthickness=0)
    left.grid(row=1, column=0, sticky="nsw", padx=2, pady=(2,4))
    text_frame = Frame(left, bd=0, highlightthickness=0); text_frame.grid(row=0, column=0, sticky="nsw")
    text_box = Text(text_frame, height=20, width=40); text_box.grid(row=0, column=0, sticky="n")
    text_scroll = Scrollbar(text_frame, command=text_box.yview); text_scroll.grid(row=0, column=1, sticky="ns")
    text_box.config(yscrollcommand=text_scroll.set)
    example = (
        "Examples:\n"
        "- One Piece (codes): 4xOP11-040\n"
        "- Yu-Gi-Oh! (names): 3 Dark Magician\n"
        "- Pokémon (names): 2x Charizard\n"
        "- MTG (names): 4 Lightning Bolt\n"
        "(you can write 4x or just 4 for the needed amount)"
        "\n\nLocal folder (for all games):\n"
        "- Searches the selected folder recursively\n"
        "- Partial matches in filenames are enough\n"
        '- Example: "char" matches "Charizard.png" (case-insensitive)'
    )
    Label(left, text=example, justify="left", fg="#666").grid(row=1, column=0, sticky="nw", pady=(2,0))

    from tkinter import font as _tkfont
    try: _big = _tkfont.Font(family="Segoe UI", size=11, weight="bold")
    except Exception: _big = None
    Label(left, text="SEARCH IN ONLINE SOURCES\nIN ENGLISH CARD-NAMES FOR BEST RESULTS", font=_big, fg="#333").grid(row=2, column=0, sticky="nw", pady=(6,0))

    def normalize_folder(p: str) -> str:
        p = p.strip().strip('"').strip("'")
        return os.path.normpath(os.path.abspath(p))

    def ensure_unique(path: str) -> str:
        if overwrite_var.get(): return path
        base, ext = os.path.splitext(path); candidate = path; n = 1
        while os.path.exists(candidate): candidate = f"{base} ({n}){ext}"; n += 1
        return candidate

    def on_download():
        lines = text_box.get("1.0", END).strip().splitlines()
        if not lines:
            messagebox.showerror("Error", "Please enter a list!"); return
        save_dir_raw = folder_var.get().strip()
        if not save_dir_raw:
            messagebox.showerror("Error", "Please select a save folder!"); return
        target_dir = normalize_folder(save_dir_raw); os.makedirs(target_dir, exist_ok=True)

        success, failed = [], []
        total = len([ln for ln in lines if ln.strip()])
        src = source_var.get()
        game = game_var.get()
        local_dir = normalize_folder(local_dir_var.get()) if src == "local" else None
        collected_for_a4: List[bytes] = []
        first_title_for_sheet: Optional[str] = None

        for idx, line in enumerate(lines, start=1):
            txt = line.strip()
            if not txt: continue
            m = re.match(r'\s*(\d+)\s*(?:[x×]\s*)?(.+?)\s*$', txt, re.I)
            if m: qty = int(m.group(1)); term = m.group(2).strip()
            else: qty = 1; term = txt

            display = term
            if game == "One Piece" and looks_like_op_code(term): display = term.upper()
            if not first_title_for_sheet: first_title_for_sheet = display
            effective_qty = qty if multiply_var.get() else 1
            status_label.config(text=f"Processing {idx}/{total}: {display} …"); right.update_idletasks()

            try:
                if choose_art_var.get():
                    variants = probe_all_arts(game, term, src, local_dir)
                    if not variants:
                        raise RuntimeError(f"No images found for {display}.\n(Hint: OP needs codes; others use names.)")
                    chosen = variants[0][1] if len(variants) == 1 else (pick_art_popup(root, display, variants) or None)
                    if chosen is None: continue
                    img_bytes = chosen
                else:
                    img_bytes = download_card_default(game, term, src, local_dir)

                if out_mode.get() == "images":
                    for i in range(effective_qty):
                        base = f"{display}_{i+1}" if effective_qty > 1 else display
                        safe = re.sub(r'[^A-Za-z0-9_-]+', '_', base)[:60]
                        out_file = ensure_unique(os.path.join(target_dir, f"{safe}.png"))
                        save_png(img_bytes, out_file,
                                 add_border=border_var.get(),
                                 border_px=max(0, int(border_px_var.get())), border_color=border_color_var.get(),
                                 do_upscale=upscale_var.get(),
                                 min_height_px=max(1, int(min_height_var.get())),
                                 dpi=max(72, int(dpi_var.get())))
                        success.append(out_file)
                else:
                    for _ in range(effective_qty): collected_for_a4.append(img_bytes)

            except Exception as e:
                failed.append(f"{qty}x{display}"); print(f"[ERROR] {display}: {e}")

        if out_mode.get() == "a4sheet" and collected_for_a4:
            dpi = max(72, int(dpi_var.get()))
            def _crop_color_rgb(name: str):
                return {
                    "Black": (0,0,0),
                    "Red": (230,0,0),
                    "Green": (0,170,0),
                    "Blue": (30,90,255),
                }.get(name, (0,0,0))
            pages = build_a4_pages(collected_for_a4, dpi=dpi,
                card_w_mm=float(card_w_mm.get()), card_h_mm=float(card_h_mm.get()),
                margin_x_mm=float(margin_x_mm.get()), margin_y_mm=float(margin_y_mm.get()),
                gap_x_mm=float(gap_x_mm.get()), gap_y_mm=float(gap_y_mm.get()),
                crop_marks=bool(crop_var.get()),
                crop_len_mm=float(crop_len_var.get()), crop_gap_mm=float(crop_gap_var.get()),
                crop_stroke_px=int(crop_stroke_px_var.get()), crop_color=_crop_color_rgb(crop_color_var.get()),
                add_border=border_var.get(), border_px=int(border_px_var.get()), border_color=border_color_var.get()
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
                        if fmt == "PNG": p.save(path, "PNG", optimize=True)
                        else: p.save(path, "JPEG", quality=95, subsampling=0, optimize=True)
                        success.append(path)

        status_label.config(text="Download finished.")
        msg = f"{len(success)} file(s) saved in:\n{target_dir}"
        if failed: msg += f"\nFailed: {', '.join(failed)}"
        messagebox.showinfo("Done", msg)

    def pick_art_popup(root_win: Tk, title_text: str, found_images: List[Tuple[str, bytes]]) -> Optional[bytes]:
        if not found_images: return None
        win = Toplevel(root_win); win.title(f"Select art for {title_text}")
        try: win.iconbitmap(resource_path("app.ico"))
        except Exception: pass

        win.geometry("1000x700")
        try: win.state('zoomed')
        except Exception:
            try: win.attributes('-zoomed', True)
            except Exception:
                try:
                    sw = root_win.winfo_screenwidth(); sh = root_win.winfo_screenheight()
                    win.geometry(f"{sw}x{sh}+0+0")
                except Exception: pass
        win.minsize(1000, 700)
        win.rowconfigure(0, weight=1); win.columnconfigure(0, weight=1)

        canvas = Canvas(win, highlightthickness=0, bd=0); canvas.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scr = Scrollbar(win, orient="vertical", command=canvas.yview); scr.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scr.set)

        frame = Frame(canvas, bd=0, highlightthickness=0)
        inner = canvas.create_window((0, 0), window=frame, anchor="nw")

        def _on_mousewheel(event):
            try:
                if event.delta > 0: canvas.yview_scroll(-3, "units")
                elif event.delta < 0: canvas.yview_scroll(3, "units")
            except Exception: pass
            try:
                if getattr(event, "num", None) == 4: canvas.yview_scroll(-3, "units")
                elif getattr(event, "num", None) == 5: canvas.yview_scroll(3, "units")
            except Exception: pass
        try:
            win.bind_all("<MouseWheel>", _on_mousewheel)
            win.bind_all("<Button-4>", _on_mousewheel)
            win.bind_all("<Button-5>", _on_mousewheel)
        except Exception: pass

        thumbs: List[ImageTk.PhotoImage] = []; chosen = {"data": None}
        for idx, (url, data) in enumerate(found_images):
            try: tkimg = make_padded_thumb(data, THUMB_W, THUMB_H)
            except Exception: continue
            thumbs.append(tkimg)
            r, c = divmod(idx, THUMB_COLS)
            cell = Frame(frame, bd=1, relief="groove", width=THUMB_W+8, height=THUMB_H+38, highlightthickness=0)
            cell.grid(row=r, column=c, padx=6, pady=6, sticky="n"); cell.grid_propagate(False)
            lbl = Label(cell, image=tkimg); lbl.image = tkimg; lbl.pack(padx=3, pady=2)
            label_text = url if os.path.isabs(url) else url.rsplit("/", 1)[-1]
            cap = Label(cell, text=os.path.basename(label_text)); cap.pack(padx=2, pady=(0,2))
            def on_dbl(_e, b=data): chosen["data"] = b; win.destroy()
            lbl.bind("<Double-Button-1>", on_dbl); cell.bind("<Double-Button-1>", on_dbl)

        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(inner, width=e.width))
        win.transient(root_win); win.grab_set(); root.wait_window(win)
        return chosen["data"]

    # Settings (left bottom)
    settings_box = LabelFrame(left, text="Settings", padx=4, pady=2, bd=1)
    settings_box.grid(row=3, column=0, sticky="nw", padx=2, pady=(2,2))
    Button(settings_box, text="Save As", width=10, command=save_settings_as_dialog).pack(side="left", padx=(2,6))
    Button(settings_box, text="Load", width=10, command=load_settings_from_dialog).pack(side="left", padx=(0,2))

    # Neue Links-Box unter Settings
    links_box = Frame(left)
    links_box.grid(row=4, column=0, sticky="nw", padx=2, pady=(100,2))

    Button(links_box, text="GitHub", width=10, command=lambda: webbrowser.open("https://github.com/Hylianer95/ProxyCardsTool")).pack(side="left", padx=2)
    Button(links_box, text="Discord", width=10, command=lambda: webbrowser.open("https://discord.gg/ShyFdE49fg")).pack(side="left", padx=2)
    Button(links_box, text="Cheat", width=10, command=lambda: webbrowser.open("https://youtu.be/dQw4w9WgXcQ?si=WtdIz9vxCAaGoF0g")).pack(side="left", padx=2)


    root.mainloop()
if __name__ == "__main__":
    print("Launching GUI…")
    start_gui()
