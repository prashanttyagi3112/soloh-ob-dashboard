#!/usr/bin/env python3
"""
Remember Prashant's "don't add" answers so the automation stops asking about that person.
Stored encrypted in vault/state.json.enc (Margin Dashboard password).

    python3 automation/state.py show
    python3 automation/state.py skip-margin "Full Name" "reason"   # never propose for the Margin Dashboard
    python3 automation/state.py skip-soloh  "Full Name" "reason"   # keep off the Soloh Margin Dashboard
    python3 automation/state.py unskip "Full Name"
"""
import argparse, datetime as dt, json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vault import lock_bytes, unlock_bytes
P = os.path.join(os.path.dirname(HERE), "vault", "state.json.enc")

def load(pw):
    if not os.path.exists(P): return {"skipMargin": {}, "skipSoloh": {}}
    return json.loads(unlock_bytes(open(P, "rb").read(), pw))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["show", "skip-margin", "skip-soloh", "unskip"])
    ap.add_argument("name", nargs="?"); ap.add_argument("reason", nargs="?", default="")
    ap.add_argument("--password", default=os.environ.get("DASHBOARD_PASSWORD"))
    a = ap.parse_args()
    if not a.password: sys.exit("Password required")
    s = load(a.password); s.setdefault("skipMargin", {}); s.setdefault("skipSoloh", {})
    if a.cmd == "show":
        print(json.dumps(s, indent=1)); return
    if not a.name: sys.exit("name required")
    if a.cmd == "unskip":
        for k in ("skipMargin", "skipSoloh"): s[k].pop(a.name, None)
    else:
        s["skipMargin" if a.cmd == "skip-margin" else "skipSoloh"][a.name] = {"reason": a.reason, "on": dt.date.today().isoformat()}
    os.makedirs(os.path.dirname(P), exist_ok=True)
    open(P, "wb").write(lock_bytes(json.dumps(s).encode(), a.password))
    print(f"{a.cmd} {a.name}: saved")

if __name__ == "__main__":
    main()
