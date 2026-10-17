#!/usr/bin/env python3
"""
Lock / unlock private files (the pay-package workbook, the automation state) so they can live
in this public repo. Same scheme as the dashboard: PBKDF2-SHA256 (310,000 rounds) + AES-256-GCM.

    python3 automation/vault.py lock   data/OB_Candidates_Pay_Packages_Soloh.xlsx vault/workbook.xlsx.enc
    python3 automation/vault.py unlock vault/workbook.xlsx.enc data/OB_Candidates_Pay_Packages_Soloh.xlsx

Password: --password or the DASHBOARD_PASSWORD environment variable (the Margin Dashboard password).
"""
import argparse, base64, json, os, secrets, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

ITER = 310000
b64e = lambda b: base64.b64encode(b).decode()
b64d = base64.b64decode

def key(pw, salt, it=ITER):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=it).derive(pw.encode())

def lock_bytes(data: bytes, pw: str) -> bytes:
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    ct = AESGCM(key(pw, salt)).encrypt(iv, data, None)
    return json.dumps({"v": 1, "kdf": "PBKDF2-SHA256", "iter": ITER, "salt": b64e(salt), "iv": b64e(iv), "ct": b64e(ct)}).encode()

def unlock_bytes(blob: bytes, pw: str) -> bytes:
    b = json.loads(blob)
    try:
        return AESGCM(key(pw, b64d(b["salt"]), b["iter"])).decrypt(b64d(b["iv"]), b64d(b["ct"]), None)
    except Exception:
        sys.exit("Could not unlock: wrong password or damaged file.")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=["lock", "unlock"])
    ap.add_argument("src"); ap.add_argument("dst")
    ap.add_argument("--password", default=os.environ.get("DASHBOARD_PASSWORD"))
    a = ap.parse_args()
    if not a.password: sys.exit("Password required: --password or DASHBOARD_PASSWORD")
    data = open(a.src, "rb").read()
    out = lock_bytes(data, a.password) if a.mode == "lock" else unlock_bytes(data, a.password)
    os.makedirs(os.path.dirname(os.path.abspath(a.dst)), exist_ok=True)
    open(a.dst, "wb").write(out)
    print(f"{a.mode}ed {a.src} -> {a.dst}")

if __name__ == "__main__":
    main()
