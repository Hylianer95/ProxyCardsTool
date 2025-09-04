# OnePiece Deck Downloader 🏴‍☠️📥

A simple desktop tool to download card images for the **One Piece TCG** from deck lists.  
It comes with a minimal GUI (built in **Tkinter**) and allows you to export your decks as card images.

---

## ✨ Features
- Paste your deck list directly into the app
- Automatic recognition of card codes (`4xST18-001`, `OP03-072`, etc.)
- Option to **download each card only once** or **as many copies as listed**
- Status display while downloading
- Output as `.png` images in your chosen folder
- Custom pirate-style app icon 🏴‍☠️

---

## 📥 Usage

1. **Prepare your deck list**  
   Example format:
```
1xEB02-010
4xST18-001
4xST18-004
2xOP05-070
4xEB02-035
3xEB02-061
4xOP07-064
3xST18-005
```


2. **Run the program**
- Enter your deck list in the text box
- Select the save folder
- Press **Download Deck**
- (Optional) enable the checkbox *Download Cards multiply* if you want multiple copies saved

3. **Result**
- All images are saved as `.png` files inside your chosen folder
- If multiply mode is enabled, extra copies get suffixes (`_1`, `_2`, …)
