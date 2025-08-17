# Python-GUI-Media-Organizer
Python GUI to organize your photos and videos into its each creation date time folder



# 📦 Ragilmalik’s Media Sorter — *Image & Video Organizer* ✨

> Sort thousands of photos & videos into a clean, date-based archive in minutes — safely, beautifully, and with zero drama.

---

## 🌟 Highlights

* 🎨 **Sleek UI** — Pure **Black** / **White** themes with tasteful **cyan/blue outlines** on buttons, boxes & tables.
  *No colored fonts, checkers, or tick marks.*
* 🗂️ **Smart date folders** — `Year → Month → Day` (e.g., `2025/August/17`), with **English** or **Indonesian** month names.
* 🧠 **True “Date Taken”** — Reads EXIF & common video tags; falls back to file system times when needed.
* 🧪 **Dry-Run Preview** — Simulate the whole operation first; no files copied.
* 🔁 **Resume** — Skip already-copied files on subsequent runs.
* 🧬 **Duplicate control**

  * Detection: **Off** (default) or **Exact (MD5)**
  * On hit: **Skip**, **Keep both**, or **Move to folder** (inside the same date folder; user-named)
* 🧰 **Long-path support** on Windows (handles very long file paths gracefully).
* ⚡ **Auto-tuned performance** — Chooses a safe, fast worker count based on your CPU, file sizes & network conditions.
* 📊 **Pro-grade Excel log** (`.xlsx`) with rich metadata and a bold summary at the end.

---

## 🔧 Installation

> Requires **Python 3.10+** (Windows/macOS/Linux)

```bash
pip install PySide6 openpyxl Pillow mutagen
```

Optional (for richer HEIC/HEIF metadata on some files):

```bash
pip install pillow-heif
```

---

## ▶️ Run (Developer Mode)

```bash
python media_sorter.py
```

---

## 🧱 Build a *single-file* EXE (Windows)

**Copy–paste this single line into CMD:**

```cmd
pyinstaller --onefile --noconfirm --clean --noconsole --name MediaSorter --collect-all PySide6 --hidden-import shiboken6 --hidden-import PySide6 media_sorter.py
```

* Output will be at: `dist\MediaSorter.exe`
* It’s **folderless** and portable. You can rename the EXE if you like.

---

## 🖥️ Quick Start

1. **Source Folder** — point to your messy media folder.
2. **Destination Folder** — where sorted copies will be created.
3. Options:

   * **Include subfolders (recursive)**
   * **Append date to filename** → `DD-MM-YYYY-{original_filename}`
   * **Resume (skip already-copied)**
   * **Simulate (dry-run)**
4. **Month folder language** — *English* 🇬🇧 or *Indonesian* 🇮🇩
5. **Folder template** — choose one:

   * `YYYY/MonthName/DD` *(default)*
   * `YYYY/MM/DD`
   * `YYYY/Mon/DD` *(abbr.)*
6. **Duplicates**

   * **Detect:** Off / Exact (MD5)
   * **When found:**

     * **Skip** (don’t copy dupes)
     * **Keep both** (auto-unique naming)
     * **Move to folder** (inside the date folder; name defaults to `Duplicates`, but you can change it)
7. Click **Preview** to simulate, then **Start** to copy for real.
8. When done, click **Open Last Saved Log File** to view your Excel log.

---

## 📘 Excel Log Details

**File name (saved in destination root):**

* Real copy → `Sort_Log_YYYYMMDD_HHMMSS.xlsx`
* Dry-run → `Simulation_Log_YYYYMMDD_HHMMSS.xlsx`

**Columns (headers are bold):**

* **Timestamp** (`DD:MMMM:YYYY HH:MM:SS`)
* **Source folder**
* **Destination Folder**
* **Filename**
* **New Filename** *(empty if rename disabled)*
* **Creation Date** (`DD-MM-YYYY`)
* **Creation Time** (`HH:MM:SS`)
* **Action** *(Copied / Skipped / Duplicate-Skipped / Duplicate→<YourFolder> / Simulated-…)*
* **Size (bytes)**
* **Width / Height** *(images)*
* **Duration (sec)** *(videos)*
* **Camera Make / Camera Model / Lens** *(if available)*
* **GPS Lat / GPS Lon** *(if available)*
* **Duplicate Of** *(original’s path inside your destination archive)*

At the end of the sheet you’ll see a **bold Summary** row (totals, mode, etc.).

---

## 🧠 Auto-Tune Performance (How it works)

* Looks at your **CPU count** and a **sample of file sizes**.
* Adjusts parallel workers to keep disks & networks happy:

  * **More** workers for lots of small files.
  * **Fewer** for giant files or network shares (UNC paths).
* Result: **fast**, but **stable** — no thrashing.

---

## 🛡️ Safety & Design Choices

* **Copy, not move** — your originals remain untouched.
* **Resume** is based on a small index file stored in the destination root (`.media_sorter_index.json`).
* **Duplicates** (when detection is on):

  * **Skip** → do nothing for the duplicate.
  * **Keep both** → copy with a unique name.
  * **Move to folder** → copy into `.../Year/Month/Day/<YourFolder>/`.
* **Long-paths** handled on Windows using the `\\?\` prefix under the hood.

---

## 🎛️ UI Notes

* Themes are **pure Black** / **pure White**.
* **Cyan (dark)** and **Blue (light)** accents are used **only as outlines**:

  * Buttons, group boxes, dropdown borders, table gridlines.
  * **Not** used for fonts, checkers, or tick marks.
* Dropdown text has comfy padding and alignment.
* Log table columns auto-size smartly.
  You can **resize rows** by dragging the left row header to your liking.

---

## ❓ FAQ

**Q: Does the app change my file metadata?**
A: No. It reads EXIF/video metadata to determine **Date Taken**, but **never writes** metadata.

**Q: What if EXIF is missing or broken?**
A: The app falls back to the file’s creation time, or modified time if needed.

**Q: HEIC/HEIF support?**
A: Files copy regardless. For better EXIF on HEIC, install `pillow-heif`. Some HEICs may not expose all tags—handled gracefully.

**Q: Can I run this on a network share (NAS)?**
A: Yes. Auto-tune will reduce worker count for UNC paths to keep things stable.

**Q: Windows long paths?**
A: Supported transparently.

---

## 🧭 Roadmap (ideas)

* Custom filename templates (tokens like `{YYYY}`, `{Mon}`, `{HHmm}`).
* Built-in thumbnail preview for a selected file.
* Per-extension rules (e.g., different folders for RAWs).

> Have ideas? Open an issue! We’re all ears. 👂

---

## 📝 License

**MIT** — free to use, modify, and share.
If this project saves you hours (it will), a ⭐ on GitHub makes our day.

---

## 🚀 Final Nudge

Run a **Preview** on your messiest folder.
If the plan looks perfect, hit **Start** and watch chaos turn into order.
*Your future self will thank you.* 🧹📁

<img width="1365" height="727" alt="Screenshot_1" src="https://github.com/user-attachments/assets/ed43b3f9-6330-4b94-bbf4-293a1e4370b2" />

