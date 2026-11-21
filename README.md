# Chord Sheet Maker

Turns an Ultimate Guitar chord sheet into a clean, editable Word document (and optionally a PDF), with the chords bold and aligned above the right words. It can also transpose the chords afterwards, including in documents you made by hand.

It comes as a small desktop app for Mac and Windows (no Python needed), and as a command-line script.

> This is an unofficial personal tool and is not affiliated with Ultimate Guitar. Their terms of use restrict automated access, so use it for your own personal chord sheets.

> Note: this tool is about 80% vibe-coded with Claude :)

---

## What it does

- **Fetches a chord sheet** from an Ultimate Guitar URL (or from a page you saved in your browser) and builds a Word document.
- **Matches a fixed layout:**
  - Title as "Song Title – Artist", Arial 18, bold.
  - Everything else Arial, bold chords, section labels like `[Chorus]` in plain text.
  - A4 page with 1" margins.
- **Keeps chords in the right place** above the text, even in a proportional font like Arial.
- **Keeps sections together:** if a verse or chorus doesn't fit on the page, the whole section moves to the next page.
- **Converts capo versions to the original key.** Ultimate Guitar shows the guitar chords plus a capo number. The tool raises every chord by the capo number so you get the chords of the original song.
- **Removes junk before the song.** Everything before the first `[Section]` line (tabber notes, tuning info, etc.) is dropped.
- **Transposes** while creating the document, or later on any existing document.
- **Auto-fits the text size** between 8 and 12 pt to use the fewest pages, or lets you pick a fixed size.
- **Exports a PDF** with the same layout, if you want one.
- **Follows your system's dark mode** in the app (Windows and Mac).

---

## Download

Go to the **[Releases page](../../releases/latest)** of this repository and download:

- **Mac with Apple Silicon (M1 or newer):** `ChordSheetMaker-mac-apple-silicon.zip`. Unzip it and open **Chord Sheet Maker.app**.
- **Mac with an Intel processor:** `ChordSheetMaker-mac-intel.zip`. Unzip it and open **Chord Sheet Maker.app**.
- **Windows:** `ChordSheetMaker-windows.exe`. Double-click it.

Not sure which Mac you have? Click the Apple menu > **About This Mac**: it says "Chip" for Apple Silicon and "Processor" for Intel. You don't need Python or a GitHub account.

### First launch

The apps are not signed with a paid developer certificate, so the operating system will warn you the first time.

- **Mac:** double-click the app once, then open **System Settings > Privacy & Security** and click **Open Anyway** at the bottom. Alternatively, run this once in Terminal:
  ```
  xattr -cr "/Applications/Chord Sheet Maker.app"
  ```
- **Windows:** when SmartScreen appears, click **More info**, then **Run anyway**.

---

## Using the app

### New chord sheet tab

1. Paste an Ultimate Guitar chords URL.
2. Choose how many semitones to transpose (0 means leave as is) and whether chord names use sharps, flats, or automatic.
3. Tick or untick the options:
   - **Convert capo chords to the original key** (on by default)
   - **Auto-fit text size (8-12 pt)** or turn it off and type your own size (6 to 16 pt)
   - **Also save a PDF**
4. Click **Create Word document** and choose where to save it.

If the download is blocked (you get an error like "HTTP 403"), open the page in your browser, save it (Ctrl/Cmd+S, "Web Page, HTML Only"), and use **Choose file…** to select the saved page instead of pasting a URL.

### Transpose a document tab

Choose any chord-sheet Word document, set the number of semitones, and save a transposed copy (with an optional PDF). This works on documents made by this tool and on your own, as long as the chord lines are **bold**.

---

## How it works

**1. Getting the sheet.** The tool downloads the page and reads the chord data that Ultimate Guitar embeds in it (title, artist, capo and the sheet text with `[ch]` chord markers). Everything before the first `[Section]` line is discarded after the capo is read from it.

**2. Capo to original key.** A capo on fret *n* means the written chord shapes sound *n* semitones higher. The tool shifts every chord up by *n*, plus any transposition you ask for. Songs played in a lowered tuning are not corrected for.

**3. Flats or sharps (Automatic).** The tool guesses the key the chords fit best. It uses flats for F, B♭, E♭ and A♭ major (and their relative minors) and sharps for everything else, in line with Ultimate Guitar.

**4. Chords above the right words.** Chords are not lined up with spaces, since spaces don't work in a proportional font. Instead each chord gets a Word **tab stop** at the exact horizontal position of its syllable, measured with Arial's letter widths. This keeps the chords where they belong, and means transposing never shifts them sideways.

**5. Sections stay together.** Every line of a section is set to "keep with next", so Word moves the whole section to the next page if it doesn't fit. This keeps working after you edit the document (unless a section is longer than one page).

**6. Auto-fit text size.** The tool tries every size from 12 down to 8 pt in half-point steps, works out how many pages each would need with the same layout rules, and picks the size with the fewest pages (the largest size wins a tie). Sizes where a line would be too wide for the page, and would wrap and break the chord alignment, are skipped when possible. The title always stays at 18 pt.

**7. Transposing an existing document.** The tool finds the chord lines (paragraphs where all the text is bold and every word is a chord) and rewrites only those chords. In older documents that use spaces for alignment, it adjusts the spacing so the next chord stays in the same column.

**8. PDF.** The PDF is drawn directly, not converted from Word, so it works without Microsoft Word installed. It uses Helvetica, which has exactly Arial's letter widths. Characters outside Western European text may show as "?" in the PDF.

---

## Things to know

- **The PDF is a snapshot.** If you edit the Word file afterwards, use Word's *Save as PDF* to refresh it.
- **Page breaks may differ slightly in Word.** The tool predicts page counts with its own calculation, which has matched in testing, but Word can occasionally be one line off.
- **Editing lyrics:** chords don't move with the text. If you change a lyric line, adjust the chord positions with Word's tab stops (Paragraph > Tabs) or retype the chord line.
- **If Ultimate Guitar changes its website,** the downloader may need a small update. The saved-page option usually still works.
- **Dialogs:** the file-picker and message dialogs are the operating system's own and don't follow the dark theme on Windows.

---

## Repository contents

```
ug2docx.py                  core: download, parse, transpose, Word and PDF output
ug2docx_gui.py              the desktop window (needs ug2docx.py next to it)
assets/AppIcon.iconset/     the app icon (the workflow builds .icns and .ico from it)
.github/workflows/          the workflow that builds the Mac and Windows apps
build_mac_app.sh            optional: build the Mac app locally (needs Python once)
```

## Requirements (only if running from source)

Python 3.8 or newer, with `python-docx`, `pillow`, `requests`, `reportlab`, and optionally `cloudscraper` (helps if Ultimate Guitar blocks plain downloads). The window also needs Tk, which is included with the Python from python.org.