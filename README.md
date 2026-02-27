# ADSJ Grade encoder

Chrome extension for OLFU AIMS grading sheet that helps you:

- Unlock PRELIM / MIDTERM / FINAL grade cells
- Generate Excel template
- Upload grades by selected period (Prelim, Midterm, Final)
- Encode and persist grades with retry logic

## Features

- Floating action button on AIMS grading page
- Unlock controls for period columns and hidden Save/Finalize buttons
- Template with dedicated columns: `PRELIM`, `MIDTERM`, `FINAL`
- Upload period selector (choose which period to apply)
- Per-cell backend save attempts with retries

## Requirements

- Google Chrome
- Access to `https://sis.fatima.edu.ph/*`

## Setup (Step by Step)

1. Open Chrome and go to `chrome://extensions`.
2. Turn on **Developer mode** (top-right).
3. Click **Load unpacked**.
4. Select this folder: `grade_encoder`.
5. Confirm extension is loaded as **ADSJ Grade encoder**.

## How to Use (Step by Step)

1. Open AIMS grading sheet page (`https://sis.fatima.edu.ph/11/113`) and select a schedule.
2. Click the floating button `📊` (bottom-right).
3. Click **Unlock Everything**.
4. Click **Download Template**.
5. Fill grades in Excel (`PRELIM`, `MIDTERM`, `FINAL` columns as needed).
6. In the panel, choose upload target period:
   - `PRELIM` or
   - `MIDTERM` or
   - `FINAL`
7. Upload the filled file.
8. Click **Apply Grades Now**.
9. Wait for completion status.
10. Click **Save** in AIMS.

## Notes

- Encoding is slower by design because it saves/retries per cell for reliability.
- If some rows still fail, rerun **Apply Grades Now** for the same period.

## Project Files

- `manifest.json` — extension metadata and permissions
- `content.js` — injects scripts into AIMS page context
- `encoder.js` — main UI/logic (unlock, template, upload, apply)
- `xlsx.min.js` — SheetJS library for Excel processing

