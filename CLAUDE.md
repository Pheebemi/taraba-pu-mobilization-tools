# Job: APC polling unit agent forms → Taraba PU Mobilization Teams sheet

This folder holds phone photos of handwritten forms. Your job is to read every
photo and record each person on it in `raw.csv`. At the end, `merge.py`
builds the final Excel sheet. Work through the photos in order, and always
resume from where the progress files say you stopped.

## Files
- `pu images/`: the original iPhone photos (note the space in the folder name, so always quote the path). **Never modify, move or delete these.**
- `jpg/`: JPG copies of the photos. Read these, because you cannot open HEIC.
- `raw.csv`: your output, one row per person (create it with the header below if it doesn't exist).
- `done.txt`: one line per photo you have finished, including photos with no people on them.
- `skipped.txt`: photos that are not a form or are completely unreadable, with a reason.
- `merge.py`: builds `Taraba_PU_Mobilization_Teams_FILLED.xlsx` from raw.csv. Run it with `python3 merge.py`.
- `Taraba_PU_Mobilization_Teams_Template.xlsx`: the final column layout. Don't edit it.

## Step 0, once: convert photos
If `jpg/` is missing or has fewer files than `pu images/`, convert the missing ones (macOS `sips`, already installed):
```
mkdir -p jpg
find "pu images" -maxdepth 1 -type f \( -iname '*.heic' -o -iname '*.jpg' -o -iname '*.jpeg' -o -iname '*.png' \) | while IFS= read -r f; do
  b=$(basename "${f%.*}"); [ -e "jpg/$b.jpg" ] || sips -s format jpeg -Z 2000 "$f" --out "jpg/$b.jpg" >/dev/null
done
```
Also make sure openpyxl is available for merge.py: `python3 -c "import openpyxl" || pip3 install --user openpyxl`.

## What the forms look like
Each photo is one printed page with this heading:
- "LIST OF APC POLLING UNIT AGENTS FOR **GUBERNATORIAL** ELECTIONS. <X> WARD." → form = `GOV`
- "LIST OF APC POLLING UNIT AGENTS FOR **STATE HOUSE OF ASSEMBLY** ELECTIONS. <X> WARD." → form = `HOA`

Each page has **2 polling units** (a printed name such as "DOROFI DISPENSARY P.U"). Under each PU are
**2 numbered boxes**: `1.` and `2.`, each with NAME, ACCT, BANK, PHONE handwritten, plus a passport photo.
Many boxes are empty. Ignore BANK (we don't record it) and ignore the passport photos.

## raw.csv format
Header (exactly):
```
image,form,ward,pu_name,pu_note,slot,name,account,phone,confidence,notes
```
- `image`: the jpg filename, e.g. `IMG_9991.jpg`
- `form`: `GOV` or `HOA`
- `ward`: the ward name from the heading, without the word WARD, e.g. `Kabri`
- `pu_name`: the **printed** PU name exactly as printed, e.g. `DOROFI MARKET (2A) P.U`
- `pu_note`: anything **handwritten** next to the PU name, e.g. `(1B)`; otherwise empty
- `slot`: `1` or `2` (the box number)
- `name`: the person's name in Title Case
- `account`: the digits exactly as written. Never add, drop or "fix" digits, even if the count looks wrong.
- `phone`: the digits exactly as written
- `confidence`: `high` or `low`
- `notes`: short reason whenever confidence is low or something is odd

Only write rows for boxes that have a name or number filled in. **Skip empty boxes.**
Quote fields that contain commas (use Python's csv module or careful quoting).

## Reading rules (important: this is handwriting)
- Zoom in. If a name or number is hard to read, crop that area of the jpg at full size (Python + Pillow, saved to a temp file outside jpg/) and look again before deciding.
- Copy what is written. Don't correct spellings or guess at names.
- If part of a value is unclear, give your best reading, set `confidence=low`, and explain in `notes` (e.g. "surname squeezed at edge, read as MANAS…").
- If a value is completely illegible, leave it empty, set `confidence=low`, and note it.
- Watch digits that look alike: 1/7, 4/9, 5/6, 0/6, 3/8, 2/7. Any doubt means `low`.
- Account numbers should be 10 digits and phones 11 (starting 0). If the count is off, still write exactly what's there and mark it `low`.
- Something crossed out or overwritten: use the final value and mention it in notes.

## Workflow
1. Run Step 0 if needed.
2. List the jpgs in `jpg/` in name order, then skip any already listed in `done.txt` or `skipped.txt`.
3. For each photo: read it, append its rows to `raw.csv`, **then** append the filename to `done.txt`. Do one photo at a time, so a stop never loses or doubles work.
4. After every 25 photos, print a one-line progress count (done / total, rows so far, low-confidence so far).
5. When I say "merge", or when all photos are done, run `python3 merge.py` and tell me the summary it prints.
6. If you are asked to do "the first N", stop after N photos and show me those rows as a table.

If a photo is not one of these forms, is upside-down beyond reading, or is a duplicate shot of a page
already done (same ward + same PUs + same form type), write it to `skipped.txt` with the reason instead.
