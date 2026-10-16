#!/usr/bin/env python3
"""
Build the Soloh Margin Dashboard from the Excel workbook.

Usage:
    python3 build.py --password "YOUR-PASSWORD" [--xlsx data/OB_Candidates_Pay_Packages_Soloh.xlsx]

What it does:
  1. Reads every sheet of the workbook and maps columns BY HEADER NAME (not position),
     so adding / reordering columns does not break the build.
  2. Recomputes the workbook's formula columns (weeks, MSP fee, NBO, gross / net revenue,
     net profit) using the same formulas and the settings on the "Client Fees" sheet,
     and compares them with the values Excel last saved. Any mismatch is reported.
  3. Runs data-quality checks (blanks, duplicates, odd values) without deleting anything.
  4. Encrypts the resulting dataset with the password (PBKDF2-SHA256 + AES-256-GCM)
     and writes index.html from template.html. Only the encrypted blob is published.
"""
import argparse, base64, datetime as dt, json, os, re, secrets, sys
from zoneinfo import ZoneInfo

import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

HERE = os.path.dirname(os.path.abspath(__file__))
ET = ZoneInfo("America/New_York")
PBKDF2_ITER = 310000

# ---------------------------------------------------------------- column mapping
# Each dashboard field lists header aliases (lower-case, punctuation-insensitive).
FIELD_ALIASES = {
    "candidate":  ["candidate", "candidate name", "name", "worker"],
    "client":     ["client", "health system"],
    "facility":   ["facility / unit", "facility/unit", "facility", "unit"],
    "req":        ["req / eng id", "req id", "req", "job id"],
    "start":      ["start", "start date"],
    "end":        ["end", "end date"],
    "weeks":      ["no. of weeks", "weeks", "number of weeks"],
    "shift":      ["shift type", "shift"],
    "hrs_shift":  ["hours / shift", "hours per shift"],
    "guar_hrs":   ["guaranteed hrs / wk", "guaranteed hours", "guaranteed hrs"],
    "w2":         ["w2 hourly rate", "w2 rate", "hourly rate"],
    "ot":         ["ot rate", "overtime rate"],
    "stipend":    ["weekly stipend", "stipend"],
    "bill":       ["bill rate"],
    "bill_after": ["bill rate after norton +$5 (requested)", "bill rate after"],
    "msp":        ["msp fee %", "msp fee"],
    "nbo":        ["nbo (hrs)", "nbo"],
    "gross_rev":  ["gross revenue"],
    "net_rev":    ["net revenue"],
    "net_profit": ["net profit"],
    "reason":     ["status / reason", "status", "reason"],
}

def norm(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())

def map_headers(headers):
    mapping, unmapped = {}, []
    nh = [norm(h) for h in headers]
    for field, aliases in FIELD_ALIASES.items():
        for a in aliases:  # exact match first
            if a in nh and nh.index(a) not in mapping.values():
                mapping[field] = nh.index(a); break
    for i, h in enumerate(headers):
        if h is not None and i not in mapping.values():
            unmapped.append(str(h))
    return mapping, unmapped

def num(v):
    if v is None or v == "": return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace("$", "").replace(",", "").strip()
    pct = s.endswith("%")
    s = s.rstrip("%")
    try:
        f = float(s); return f / 100 if pct else f
    except ValueError:
        return None

def date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, (int, float)):  # Excel serial
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    if isinstance(v, str):
        for f in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
            try: return dt.datetime.strptime(v.strip(), f).date()
            except ValueError: pass
    return None

def text(v):
    return None if v is None else re.sub(r"\s+", " ", str(v)).strip() or None

# ---------------------------------------------------------------- derived labels
UNIT_RULES = [
    (r"\bct tech\b", "CT Tech (Allied)"), (r"\bicu\b", "ICU"), (r"\bpcu\b", "PCU"),
    (r"\ber\b|\bed\b|emergency", "ER"), (r"med ?surg", "Med Surg"),
    (r"\btele\b|telemetry", "Telemetry"),
]
def derive_unit(fac):
    s = norm(fac)
    for pat, label in UNIT_RULES:
        if re.search(pat, s): return label
    return "Not specified"

def derive_facility(fac):
    if not fac: return None
    first = fac.split(",")[0]
    first = re.sub(r"\s*\(.*?\)", "", first)
    first = re.sub(r"\s+in\s+.*$", "", first, flags=re.I)
    return first.strip()

# ---------------------------------------------------------------- read workbook
def read_sheet(ws_f, ws_v):
    rows_f = list(ws_f.iter_rows(values_only=True))
    rows_v = list(ws_v.iter_rows(values_only=True))
    if not rows_f: return [], {}, []
    headers = list(rows_f[0])
    mapping, unmapped = map_headers(headers)
    out = []
    for r_i in range(1, len(rows_f)):
        rf, rv = rows_f[r_i], rows_v[r_i]
        if all(c is None or str(c).strip() == "" for c in rf):
            continue
        rec = {"_row": r_i + 1, "_cached": {}, "_formulas": {}, "_extra": {}}
        for field, ci in mapping.items():
            fv, vv = rf[ci] if ci < len(rf) else None, rv[ci] if ci < len(rv) else None
            if isinstance(fv, str) and fv.startswith("="):
                rec["_formulas"][field] = fv
                rec["_cached"][field] = vv
                rec[field] = vv
            else:
                rec[field] = fv
        for i, h in enumerate(headers):
            if h is not None and i not in mapping.values():
                rec["_extra"][str(h)] = rv[i] if i < len(rv) else None
        out.append(rec)
    return out, mapping, unmapped

def read_fees(ws):
    fees, settings, notes = {}, {}, []
    rows = list(ws.iter_rows(values_only=True))
    in_table = True
    for r in rows[1:]:
        a = text(r[0]) if r else None
        if a is None:
            in_table = False; continue
        if in_table and num(r[1]) is not None:
            fees[a] = {"msp": num(r[1]), "nbo": num(r[2]) or 0, "source": text(r[3]) if len(r) > 3 else None}
            continue
        in_table = False
        la = a.lower()
        if "payroll load" in la: settings["payroll_load"] = num(r[1]); settings["payroll_load_note"] = text(r[2])
        elif "onboarding cost" in la or "recruiting" in la and num(r[1]) is not None:
            settings["onboarding_cost"] = num(r[1]) or 0; settings["onboarding_note"] = text(r[2])
        elif la not in ("soloh calculator v1.1 settings",) and not la.startswith("soloh calculator"):
            notes.append(a)
        else:
            settings["settings_title"] = a
    return fees, settings, notes

def recompute(rec, fees, settings, dq, sheet):
    """Recompute formula fields; compare with cached Excel values."""
    s, e = date(rec.get("start")), date(rec.get("end"))
    rec["start"], rec["end"] = s, e
    weeks = round((e - s).days / 7) if s and e else num(rec.get("weeks"))
    fee = fees.get(text(rec.get("client")) or "", None)
    msp = fee["msp"] if fee else num(rec.get("msp"))
    nbo = fee["nbo"] if fee else num(rec.get("nbo"))
    if fee is None and text(rec.get("client")):
        dq.append({"sev": "warn", "sheet": sheet, "row": rec["_row"], "who": text(rec.get("candidate")),
                   "msg": f"Client '{text(rec.get('client'))}' not found on Client Fees sheet; MSP/NBO taken from row."})
    G, W2, OT = num(rec.get("guar_hrs")), num(rec.get("w2")), num(rec.get("ot"))
    stip, bill = num(rec.get("stipend")), num(rec.get("bill"))
    load = settings.get("payroll_load") or 0
    onb = settings.get("onboarding_cost") or 0
    calc = {"weeks": weeks, "msp": msp, "nbo": nbo}
    if None not in (bill, G, weeks):
        calc["gross_rev"] = bill * G * weeks
        calc["net_rev"] = calc["gross_rev"] * (1 - (msp or 0))
        if None not in (W2, OT, stip):
            pay = weeks * (min(G, 40) * W2 + max(G - 40, 0) * OT) * (1 + load)
            calc["net_profit"] = calc["net_rev"] - pay - weeks * stip - (nbo or 0) * bill * (1 - (msp or 0)) - onb
    for k, v in calc.items():
        cached = num(rec["_cached"].get(k)) if k in rec["_cached"] else None
        if k in rec["_formulas"] and cached is not None and v is not None and abs(cached - v) > 0.01:
            dq.append({"sev": "error", "sheet": sheet, "row": rec["_row"], "who": text(rec.get("candidate")),
                       "msg": f"{k}: Excel shows {cached:,.2f} but recalculation gives {v:,.2f}. Using recalculated value."})
        if v is not None:
            rec[k] = v
        elif k in rec:
            rec[k] = num(rec[k])

def finalize(rec, status):
    G = num(rec.get("guar_hrs")); W2 = num(rec.get("w2")); OT = num(rec.get("ot")); stip = num(rec.get("stipend"))
    bill = num(rec.get("bill")); weeks = num(rec.get("weeks"))
    wk_taxable = (min(G, 40) * W2 + max(G - 40, 0) * OT) if None not in (G, W2, OT) else None
    out = {
        "row": rec["_row"], "status": status,
        "candidate": text(rec.get("candidate")), "client": text(rec.get("client")),
        "facilityUnit": text(rec.get("facility")), "facility": derive_facility(text(rec.get("facility"))),
        "unit": derive_unit(text(rec.get("facility"))), "req": text(rec.get("req")),
        "start": rec["start"].isoformat() if rec.get("start") else None,
        "end": rec["end"].isoformat() if rec.get("end") else None,
        "weeks": weeks, "shift": text(rec.get("shift")), "hrsShift": num(rec.get("hrs_shift")),
        "guarHrs": G, "w2": W2, "ot": OT, "stipend": stip, "bill": bill,
        "billAfter": num(rec.get("bill_after")), "msp": num(rec.get("msp")), "nbo": num(rec.get("nbo")),
        "grossRev": num(rec.get("gross_rev")), "netRev": num(rec.get("net_rev")),
        "netProfit": num(rec.get("net_profit")), "reason": text(rec.get("reason")),
        "extra": {k: (v.isoformat() if isinstance(v, (dt.date, dt.datetime)) else v) for k, v in rec["_extra"].items()},
    }
    out["weeklyTaxable"] = wk_taxable
    out["weeklyGross"] = wk_taxable + stip if wk_taxable is not None and stip is not None else None
    out["contractGrossPay"] = out["weeklyGross"] * weeks if out["weeklyGross"] is not None and weeks else None
    out["weeklyBill"] = bill * G if bill is not None and G is not None else None
    out["weeklyProfit"] = out["netProfit"] / weeks if out["netProfit"] is not None and weeks else None
    out["marginPct"] = out["netProfit"] / out["grossRev"] if out["netProfit"] is not None and out["grossRev"] else None
    return out

def quality(records, dq):
    seen = {}
    for r in records:
        who, sheet = r["candidate"], r["_sheet"]
        if not who:
            dq.append({"sev": "error", "sheet": sheet, "row": r["row"], "who": None, "msg": "Candidate name is blank."})
        elif len(who.split()) < 2:
            dq.append({"sev": "warn", "sheet": sheet, "row": r["row"], "who": who, "msg": "Only a first name is recorded (last name missing)."})
        key = (norm(who), norm(r["client"]), r["start"])
        if key in seen:
            dq.append({"sev": "warn", "sheet": sheet, "row": r["row"], "who": who, "msg": f"Possible duplicate of row {seen[key]} (same name, client and start). Counted once in Unique Candidates."})
        else:
            seen[key] = r["row"]
        for f, label in [("start", "Start date"), ("end", "End date"), ("w2", "W2 rate"), ("bill", "Bill rate"), ("stipend", "Weekly stipend"), ("guarHrs", "Guaranteed hours"), ("shift", "Shift type")]:
            if r.get(f) in (None, ""):
                dq.append({"sev": "warn", "sheet": sheet, "row": r["row"], "who": who, "msg": f"{label} is blank."})
        if r["start"] and r["end"] and r["end"] <= r["start"]:
            dq.append({"sev": "error", "sheet": sheet, "row": r["row"], "who": who, "msg": "End date is on or before start date."})
        if r["w2"] and r["ot"] and abs(r["ot"] - 1.5 * r["w2"]) > 0.05:
            dq.append({"sev": "info", "sheet": sheet, "row": r["row"], "who": who,
                       "msg": f"OT rate ${r['ot']:.2f} is {r['ot']/r['w2']:.2f}x the W2 rate ${r['w2']:.2f} (not 1.5x). "
                              + ("Affects profit, since guaranteed hours exceed 40." if (r['guarHrs'] or 0) > 40
                                 else "No effect on the numbers (guaranteed hours are 40 or less).")})
        if r["unit"] == "Not specified":
            dq.append({"sev": "info", "sheet": sheet, "row": r["row"], "who": who, "msg": f"No unit/specialty found in '{r['facilityUnit']}'."})
        if r["status"] == "Active" and r["netProfit"] is not None and r["netProfit"] < 0:
            dq.append({"sev": "error", "sheet": sheet, "row": r["row"], "who": who, "msg": "Negative net profit."})
    active = [r for r in records if r["status"] == "Active" and r["marginPct"] is not None]
    if len(active) >= 4:
        ms = sorted(r["marginPct"] for r in active)
        med = ms[len(ms) // 2]
        for r in active:
            if r["marginPct"] > med * 1.8:
                dq.append({"sev": "info", "sheet": r["_sheet"], "row": r["row"], "who": r["candidate"],
                           "msg": f"Margin {r['marginPct']*100:.1f}% is well above the group median {med*100:.1f}%. Worth confirming the pay terms."})

# ---------------------------------------------------------------- encrypt
def encrypt(payload: bytes, password: str):
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=PBKDF2_ITER).derive(password.encode())
    ct = AESGCM(key).encrypt(iv, payload, None)
    b = lambda x: base64.b64encode(x).decode()
    return {"v": 1, "kdf": "PBKDF2-SHA256", "iter": PBKDF2_ITER, "salt": b(salt), "iv": b(iv), "ct": b(ct)}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default=os.path.join(HERE, "data", "OB_Candidates_Pay_Packages_Soloh.xlsx"))
    ap.add_argument("--password", default=os.environ.get("DASHBOARD_PASSWORD"))
    ap.add_argument("--out", default=os.path.join(HERE, "index.html"))
    ap.add_argument("--dump-json", help="also write the decrypted dataset here (never commit it)")
    a = ap.parse_args()
    if not a.password:
        sys.exit("Password required: --password or DASHBOARD_PASSWORD env var")

    wb_f = openpyxl.load_workbook(a.xlsx, data_only=False)
    wb_v = openpyxl.load_workbook(a.xlsx, data_only=True)
    dq, sheets_info = [], []
    fees, settings, notes = ({}, {}, [])
    fee_sheet = next((n for n in wb_f.sheetnames if "fee" in n.lower()), None)
    if fee_sheet:
        fees, settings, notes = read_fees(wb_v[fee_sheet])

    records = []
    for name in wb_f.sheetnames:
        if name == fee_sheet:
            sheets_info.append({"name": name, "role": "Client fees & calculator settings", "rows": len(fees), "unmapped": []}); continue
        status = "Cancelled" if re.search(r"cancel|declin", name, re.I) else "Active"
        rows, mapping, unmapped = read_sheet(wb_f[name], wb_v[name])
        if "candidate" not in mapping:
            sheets_info.append({"name": name, "role": "Ignored (no Candidate column)", "rows": len(rows), "unmapped": unmapped}); continue
        sheets_info.append({"name": name, "role": f"{status} offers", "rows": len(rows), "unmapped": unmapped,
                            "mapped": sorted(mapping.keys())})
        for rec in rows:
            recompute(rec, fees, settings, dq, name)
            f = finalize(rec, status); f["_sheet"] = name
            records.append(f)
    quality(records, dq)
    for r in records:
        r["sheet"] = r.pop("_sheet")

    mtime = dt.datetime.fromtimestamp(os.path.getmtime(a.xlsx), ET)
    data = {
        "workbook": os.path.basename(a.xlsx),
        "builtAt": dt.datetime.now(ET).strftime("%a %m/%d/%Y %I:%M %p ET"),
        "fileModified": mtime.strftime("%m/%d/%Y %I:%M %p ET"),
        "records": records, "fees": fees, "settings": settings, "feeNotes": notes,
        "sheets": sheets_info, "quality": dq,
    }
    if a.dump_json:
        with open(a.dump_json, "w") as fh: json.dump(data, fh, indent=1, default=str)
    blob = encrypt(json.dumps(data, default=str).encode(), a.password)
    with open(os.path.join(HERE, "template.html")) as fh:
        html = fh.read()
    html = html.replace("/*__ENCRYPTED_DATA__*/null", json.dumps(blob))
    with open(a.out, "w") as fh:
        fh.write(html)
    act = sum(1 for r in records if r["status"] == "Active")
    print(f"Built {a.out}: {act} active, {len(records)-act} cancelled, {len(dq)} data-quality notes.")

if __name__ == "__main__":
    main()
