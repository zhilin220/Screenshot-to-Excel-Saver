# 📸 Tax Screenshot Capture Tool

> Capture any window + taskbar to a structured Excel spreadsheet with a single hotkey — built for freelancers and self-employed folks tracking tax records.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-lightblue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)

---

## ✨ Features

- **One-hotkey capture** — press `Ctrl+Shift+S` from any window to snapshot it instantly
- **Window + Taskbar** — automatically detects and includes the Windows taskbar (bottom, top, left, or right position) in every screenshot
- **GUI settings panel** — browse and set your Excel file path and temp folder without touching the code
- **Auto Excel logging** — timestamps, window title, screenshot, notes, amount, and category columns created automatically
- **Category dropdowns** — built-in Excel data validation with `Income / Expense / Deduction / Asset / Other`
- **Live log panel** — see capture events, saves, and errors directly in the GUI
- **Smart image sizing** — screenshots are resized to fit neatly in Excel (max 600px wide)
- **Self-cleaning** — temporary PNG files are deleted after being embedded in Excel

---

## 🖥️ Screenshot

```
┌─────────────────────────────────────────────────────────┐
│  📸 Tax Screenshot Capture Tool                         │
│  Configure paths then click Start.                      │
├─────────────────────────────────────────────────────────┤
│  Excel File (.xlsx): [C:\Users\...\tax_records.xlsx] [Browse…] │
│  Temp Folder:        [C:\Users\...\Python           ] [Browse…] │
├─────────────────────────────────────────────────────────┤
│  Log:  ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ │
│        ✅ Listener started.                              │
│        📸 Capturing at 14:32:01…                        │
│        ✅ Capture saved to Excel.                        │
├─────────────────────────────────────────────────────────┤
│  [ ▶ Start Listener ]  [ ⏹ Stop ]                       │
│  🟢 Listening — press Ctrl+Shift+S to capture           │
└─────────────────────────────────────────────────────────┘
```

---

## 📋 Requirements

- **Windows 10 or 11** (uses Windows API for window/taskbar detection)
- **Python 3.8+**

### Python Dependencies

```
pyautogui
keyboard
openpyxl
Pillow
pygetwindow
```

Install all at once:

```bash
pip install pyautogui keyboard openpyxl Pillow pygetwindow
```

> The script will also attempt to auto-install missing packages on first run.

---

## 🚀 Getting Started

### 1. Clone or download

```bash
git clone https://github.com/yourname/tax-screenshot-capture.git
cd tax-screenshot-capture
```

Or just download `tax_screenshot_capture.py` directly.

### 2. Install dependencies

```bash
pip install pyautogui keyboard openpyxl Pillow pygetwindow
```

### 3. Run the script

```bash
python tax_screenshot_capture.py
```

### 4. Configure paths in the GUI

- Click **Browse…** next to **Excel File** to choose where your `.xlsx` will be saved (or type a new filename — it will be created automatically).
- Click **Browse…** next to **Temp Folder** to choose where temporary PNGs are stored during processing.

### 5. Start the listener

Click **▶ Start Listener**. The status bar turns green.

### 6. Capture!

Switch to any window (browser, invoice, bank statement, etc.) and press:

```
Ctrl + Shift + S
```

The screenshot is captured and saved to your Excel file automatically.

---

## 📊 Excel Output Structure

| Column | Content | Fill in? |
|--------|---------|----------|
| A — Timestamp | Auto-filled date & time | ✅ Auto |
| B — Window Title | Active window name + dimensions | ✅ Auto |
| C — Description/Notes | Placeholder text | ✏️ You fill in |
| D — Screenshot | Embedded image | ✅ Auto |
| E — Amount | Dollar amount | ✏️ You fill in |
| F — Category | Dropdown selector | ✏️ You fill in |

**Category options** (dropdown): `Income`, `Expense`, `Deduction`, `Asset`, `Other`

---

## ⚙️ Configuration

You can change the **default** paths at the top of the script:

```python
DEFAULT_EXCEL_FILE = r"C:\Users\yourname\Documents\tax_records.xlsx"
DEFAULT_TEMP_FOLDER = r"C:\Users\yourname\Documents\temp"
```

These are just defaults — the GUI lets you override them at runtime without editing code.

---

## 🔧 Troubleshooting

### Black / blank screenshots

1. **Disable hardware acceleration** in Chrome or Edge:
   - Chrome: `Settings → System → Use hardware acceleration when available → Off`
   - Edge: `Settings → System and performance → Use hardware acceleration → Off`
2. **Run as Administrator** — right-click the script or terminal and choose *Run as administrator*.
3. Make sure the target window is **not minimized** when you press the hotkey.

### Hotkey not working

- Another application may have claimed `Ctrl+Shift+S`. Try closing screen capture apps (Snip & Sketch, ShareX, etc.).
- Run the script as Administrator to gain global hotkey access.

### Excel file locked

- Close the Excel file before triggering a capture. openpyxl cannot write to a file that Excel has open.

### `DwmGetWindowAttribute` errors

- This is a Windows API call that may fail on very old Windows 10 builds. The script automatically falls back to `GetWindowRect` and then to a full-screen capture.

---

## 🗂️ Project Structure

```
tax-screenshot-capture/
├── tax_screenshot_capture.py   # Main script (GUI + capture logic)
└── README.md
```

---

## 🛣️ Roadmap

- [ ] System tray icon (minimize to tray)
- [ ] Configurable hotkey via GUI
- [ ] Auto-open Excel after capture
- [ ] Multi-monitor support improvements
- [ ] Export to Google Sheets

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you'd like to change.

1. Fork the repo
2. Create your feature branch: `git checkout -b feature/my-feature`
3. Commit your changes: `git commit -m 'Add my feature'`
4. Push to the branch: `git push origin feature/my-feature`
5. Open a Pull Request

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 🙏 Acknowledgements

- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file manipulation
- [Pillow](https://python-pillow.org/) — Screenshot capture
- [keyboard](https://github.com/boppreh/keyboard) — Global hotkey listener
- [pygetwindow](https://github.com/asweigart/PyGetWindow) — Window management
- Windows DWM API — Accurate window bounds detection
