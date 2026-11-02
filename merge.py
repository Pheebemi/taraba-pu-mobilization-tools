"""
Build the filled Taraba PU Mobilization Teams sheet from raw.csv.

    python3 merge.py

Reads raw.csv (one row per person, written while reading the form photos),
pairs each PU's GOVERNORSHIP form (Primary + Relief agent) with its
HOUSE OF ASSEMBLY form (Canvasser 1 + 2) by ward + PU name, and writes
Taraba_PU_Mobilization_Teams_FILLED.xlsx using the template's columns.
A "Needs review" sheet lists everything worth checking by hand.
PU codes use Taraba's 34 prefix and are assigned alphabetically from the
LGAs, wards, and PUs present in raw.csv.
Safe to run as often as you like; it rebuilds the output each time.
"""

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

HERE = Path(__file__).resolve().parent
RAW = HERE / "raw.csv"
TEMPLATE = HERE / "Taraba_PU_Mobilization_Teams_Template.xlsx"
OUT = HERE / "Taraba_PU_Mobilization_Teams_FILLED.xlsx"

# form + slot on the paper  ->  person block in the template
ROLES = {
    ("GOV", "PL"): "1. Party Leader",
    ("GOV", "1"): "2. Primary Agent",
    ("GOV", "2"): "2. Relief Agent",
    ("HOA", "1"): "3. Canvasser 1",
    ("HOA", "2"): "3. Canvasser 2",
}


def key(text: str) -> str:
    """Normalise a ward / PU name for matching: case, spacing, 'P.U', 'WARD', punctuation."""
    t = (text or "").upper()
    t = re.sub(r"\bP\s*[./]?\s*U\b\.?", " ", t)   # P.U / PU / P/U
    t = re.sub(r"\bWARD\b", " ", t)
    return re.sub(r"[^A-Z0-9]", "", t)


def clean_name(text: str) -> str:
    t = re.sub(r"\bP\s*[./]?\s*U\b\.?", " ", text or "")
    return re.sub(r"\s+", " ", t).strip(" .-").title()


def digits(text: str) -> str:
    return re.sub(r"\D", "", text or "")


def fix_phone(text: str) -> str:
    d = digits(text)
    if d.startswith("234") and len(d) == 13:
        d = "0" + d[3:]
    return d


def assign_pu_codes(pus: dict) -> None:
    """Assign Taraba hierarchical codes alphabetically within the current dataset."""
    lgas = sorted(
        {p["lga"] for p in pus.values()},
        key=lambda value: (key(value), value.casefold()),
    )
    for lga_index, lga in enumerate(lgas, start=1):
        lga_pus = [p for p in pus.values() if p["lga"] == lga]
        wards = sorted(
            {p["ward"] for p in lga_pus},
            key=lambda value: (key(value), value.casefold()),
        )
        for ward_index, ward in enumerate(wards, start=1):
            ward_pus = [p for p in lga_pus if p["ward"] == ward]
            ward_pus.sort(key=lambda p: (key(p["pu"]), p["pu"].casefold()))
            for pu_index, pu in enumerate(ward_pus, start=1):
                pu["code"] = f"34-{lga_index:02d}-{ward_index:02d}-{pu_index:02d}"


KWESATI_HOA_PUS = {
    1: "ANDE-USSA II POLLING UNIT",
    2: "ANDE-USSA II POLLING UNIT (second sheet)",
    3: "AYIYI POLLING UNIT",
    4: "KUSANGSANG POLLING UNIT",
    5: "KWESATI PRI. SCHOOL UNIT",
    6: "TUTUWA POLLING UNIT",
    7: "KUNDAK-KWAE POLLING UNIT",
    8: "MATERNITY POLLING UNIT",
    9: "GDSS KWESATI POLLING UNIT",
    10: "NYITEN POLLING UNIT",
    11: "ATSIRIXOM POLLING UNIT",
    12: "MBITU/NYIDO POLLING UNIT",
    13: "AIAHA POLLING UNIT",
    14: "RIKUMCWO POLLING UNIT",
    15: "RIKWETAM POLLING UNIT",
    16: "KUTUKOK POLLING UNIT",
    17: "RIPEACHI POLLING UNIT",
}


def main() -> None:
    if not RAW.exists():
        sys.exit("raw.csv not found - nothing to merge yet.")
    with RAW.open(newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    ussa_gov_by_code = defaultdict(set)
    for r in rows:
        if (r.get("lga", "").strip().casefold() == "ussa"
                and (r.get("form") or "").strip().upper() == "GOV"):
            match = re.search(r"(\d{1,4})\s*$", r.get("pu_name", ""))
            if match:
                ussa_gov_by_code[(key(r.get("ward", "")), int(match.group(1)))].add(r["pu_name"])

    review = []  # (lga, ward, pu, role, image, issue)
    pus = {}     # (lga_key, ward_key, pu_key) -> {"lga", "ward", "pu", "forms": set(), "people": {role: row}}
    seen = defaultdict(list)

    for r in rows:
        form = (r.get("form") or "").strip().upper()
        slot = (r.get("slot") or "").strip()
        lga, ward, pu = r.get("lga", ""), r.get("ward", ""), r.get("pu_name", "")
        img = r.get("image", "")
        if (form == "HOA" and img.startswith("Ussa-HOS/")
                and lga.strip().casefold() == "ussa"):
            match = re.search(r"(\d{1,4})\s*$", pu)
            if match:
                code = int(match.group(1))
                if key(ward) == key("Kwesati"):
                    pu = KWESATI_HOA_PUS.get(code, pu)
                else:
                    candidates = ussa_gov_by_code.get((key(ward), code), set())
                    if len(candidates) == 1:
                        pu = next(iter(candidates))
        role = ROLES.get((form, slot))
        if not role:
            review.append((lga, ward, pu, f"{form} slot {slot}", img, "Unknown form type or slot number"))
            continue

        k = (key(lga), key(ward), key(pu))
        p = pus.setdefault(k, {"lga": clean_name(lga), "ward": clean_name(ward), "pu": clean_name(pu),
                               "code": "", "forms": set(), "people": {}, "notes": set()})
        pu_code = (r.get("pu_code") or "").strip()
        if pu_code:
            if p["code"] and p["code"] != pu_code:
                review.append((lga, ward, pu, "PU Code", img,
                               f"Conflicting PU code {pu_code}; kept {p['code']}"))
            else:
                p["code"] = pu_code
        p["forms"].add(form)
        if r.get("pu_note"):
            p["notes"].add(r["pu_note"].strip())

        seen[(k, role)].append(img)
        if role in p["people"]:
            review.append((lga, ward, pu, role, img,
                           f"Duplicate: this PU/role also appears in {p['people'][role]['image']} (kept the first)"))
            continue

        acct, phone = digits(r.get("account", "")), fix_phone(r.get("phone", ""))
        person = {"image": img, "name": (r.get("name") or "").strip().title(),
                  "acct": acct, "phone": phone,
                  "email": (r.get("email") or "").strip(),
                  "dob": (r.get("dob") or "").strip(),
                  "state_origin": (r.get("state_origin") or "").strip().title(),
                  "origin_lga": (r.get("origin_lga") or "").strip().title()}
        p["people"][role] = person

        if not person["name"]:
            review.append((lga, ward, pu, role, img, "Name missing or unreadable"))
        if acct and len(acct) != 10:
            review.append((lga, ward, pu, role, img, f"Account number has {len(acct)} digits (should be 10): {acct}"))
        if not acct:
            review.append((lga, ward, pu, role, img, "Account number missing"))
        if phone and not (len(phone) == 11 and phone.startswith("0")):
            review.append((lga, ward, pu, role, img, f"Phone looks wrong ({len(phone)} digits): {phone}"))
        if (r.get("confidence") or "").strip().lower() == "low":
            review.append((lga, ward, pu, role, img, "Low confidence: " + (r.get("notes") or "check against photo")))
        elif r.get("notes"):
            review.append((lga, ward, pu, role, img, "Note: " + r["notes"]))

    for p in pus.values():
        for form, label in (("GOV", "Governorship"), ("HOA", "House of Assembly")):
            if form not in p["forms"]:
                review.append((p["lga"], p["ward"], p["pu"], "-", "-",
                               f"No {label} form found for this PU (name may be spelled differently)"))

    assign_pu_codes(pus)

    # ---- write the workbook, keeping the template's header row ----
    if TEMPLATE.exists():
        wb = load_workbook(TEMPLATE)
        ws = wb.worksheets[0]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)   # drop the sample row
    else:
        sys.exit(f"Template not found: {TEMPLATE.name}")
    headers = [c.value for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    for p in sorted(pus.values(), key=lambda p: (p["lga"], p["ward"], p["pu"])):
        row = [None] * len(headers)
        pu_name = p["pu"]
        for n in sorted(p["notes"]):
            if n.upper().strip("() ") not in pu_name.upper():
                pu_name += f" {n if n.startswith('(') else '(' + n + ')'}"
        row[col["PU Name"]] = pu_name
        if "PU Code" in col:
            row[col["PU Code"]] = p["code"] or None
        row[col["Ward"]] = p["ward"]
        row[col["LGA"]] = p["lga"]
        for role, person in p["people"].items():
            row[col[f"{role} Name"]] = person["name"] or None
            row[col[f"{role} Phone"]] = person["phone"] or None
            row[col[f"{role} Account Number"]] = person["acct"] or None
            if f"{role} Email" in col:
                row[col[f"{role} Email"]] = person.get("email") or None
                row[col[f"{role} Date of Birth"]] = person.get("dob") or None
                row[col[f"{role} State of Origin"]] = person.get("state_origin") or None
                row[col[f"{role} Origin LGA"]] = person.get("origin_lga") or None
        ws.append(row)

    # numbers stored as text so leading zeros survive
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.value is not None:
                c.value = str(c.value)
                c.number_format = "@"

    rv = wb.create_sheet("Needs review")
    rv.append(["LGA", "Ward", "PU Name", "Role", "Image", "Issue"])
    for item in review:
        rv.append(list(item))
    for sheet in (ws, rv):
        for c in sheet[1]:
            c.font = Font(bold=True)
        sheet.freeze_panes = "A2"
        for column in sheet.columns:
            width = max(len(str(c.value or "")) for c in column)
            sheet.column_dimensions[column[0].column_letter].width = min(max(width + 2, 12), 60)

    wb.save(OUT)
    print(f"{len(pus)} polling units, {len(rows)} people -> {OUT.name}")
    print(f"{len(review)} items on the 'Needs review' sheet")


if __name__ == "__main__":
    main()
