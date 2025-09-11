
# ProxyCardMaker 

A compact desktop tool to collect and export **One Piece TCG** card art from deck lists.  
It ships with a lightweight **Tkinter** GUI and can save **individual PNGs** or **ready‑to‑print A4 sheets (3×3)**.

> For personal/educational use only.

---

## ✨ What’s new

- **Sources**: OnePiece.gg, limitlesstcg, or **Local folder**.
- **Smart local search**: Besides card codes (e.g. `OP12-025`) you can put **free text** into the deck list (e.g. `3xKuzan`, `51xCard_back`).  
  In *Local folder* mode all matching filenames are found (case‑insensitive).
- **Art picker**: Enable *__I want to select picture art__* to pick from all found variants/thumbnails (web or local).
- **A4 export (3×3)** with **format selection**: **PDF**, **PNG**, or **JPG**.
- **DPI** setting that **applies to both** output modes (PNG and A4).
- **Precise layout controls** for A4: **Card W×H (mm)**, **Margins (L/R, T/B)**, **Gaps (H, V)**.
- **Corner crop marks** (length & gap in mm) for easier cutting.
- **Download cards multiply** toggle (respects quantities from the deck list).
- **Overwrite existing files** toggle (otherwise unique file names are created).
- **Image processing** options: *Add white border*, *Upscale (min height)*.

---

## 🖥️ UI overview

1. **Source & Local folder**  
   Choose *OnePiece.gg*, *limitlesstcg*, or *Local folder* (choose a directory with images).  
   - If a line looks like a code (`OP12-025`), it fetches by code.  
   - Otherwise it searches locally by **substring** (e.g. `Kuzan`, `Card_back`).

2. **Save folder**  
   Target directory for all outputs.

3. **Output mode**  
   - **Save individual images (PNG)** – saves each (selected) art as a PNG.  
   - **Save A4 sheet 3×3** – choose **PDF / PNG / JPG**; 
   - *DPI applies to both*.

4. **A4 layout (mm)**  
   - **Card W×H** (default **63 × 88 mm**)  
   - **Margins** L/R, T/B (mm)  
   - **Gaps** H, V (mm)  
   - **Crop marks**: enable + set **len(mm)** & **gap(mm)**

5. **Image processing**  
   - **Add white border** (px)  
   - **Upscale (min height)** to ensure a minimum image height before placing

6. **Options**  
   - **Download cards multiply** – obeys quantities like `4xOP01-025`  
   - **I want to select picture art** – shows a thumbnail picker if multiple arts exist  
   - **Overwrite existing files** – otherwise unique names like `name (1).png`

---

## 📥 Usage

1) **Enter your deck list** (one item per line)  
   You can mix codes and free text:
   ```text
   1xOP11-040
   2xOP05-067
   4xST18-001
   4xEB01-061
   4xOP10-072
   2xOP07-064
   3xKuzan
   51xCard_back
   ```

2) **Pick Source & folders**  
   - For **Local folder** point to the directory that contains your images.

3) **Choose the output mode**  
   - **PNG** per card *or* **A4 sheet 3×3** as **PDF/PNG/JPG**.  
   - Set **DPI** (recommended **300**).

4) **Adjust A4 layout** (optional)  
   - Card size, margins, gaps, crop‑marks.

5) **Download Deck**  
   - If *Select picture art* is on, choose the variant(s) you want.

### Result
- **PNG mode**: individual files are written to the save folder.  
- **A4 mode**:  
  - **PDF** → one multi‑page file  
  - **PNG/JPG** → `A4_<name>_001.png/.jpg`, `A4_<name>_002...`

---

## 🖨️ Printing tips (true scale)

- In the print dialog use **Actual size / 100%** (no “shrink to fit”).  
- Paper size **A4**, no extra scaling in the driver.  
- If your printer still overscales a little, reduce Card WxH in settings .

----

## ❓ Troubleshooting

- **No images found**: check Internet (for web sources), spelling of codes, or the local folder path & filenames.  
- **Output too large/small** when printing: verify **DPI**, and that your viewer uses **100%** scale.  
- **Existing files get new suffixes**: disable **Overwrite existing files** if you want to truly overwrite.

---



## 🙏 Credits
- Data/Images: **OnePiece.gg**, **limitlesstcg**  
- Libraries: **Pillow**, **Requests**, **Tkinter**
