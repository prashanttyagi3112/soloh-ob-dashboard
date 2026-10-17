#!/usr/bin/env python3
"""
Find what is missing between the three Soloh dashboards. Read-only: changes nothing.

  Compliance Board (onboarding cards)  ->  Margin Dashboard  ->  Soloh Margin Dashboard

    python3 automation/sync_check.py \
        --board ../soloh-compliance-board/index.html \
        --soloh ../soloh-margin-dashboard/index.html \
        --password "$DASHBOARD_PASSWORD" --soloh-password "$SOLOH_PASSWORD" [--json out.json]

Reports:
  board_not_in_margin   active board cards with no row on the Margin Dashboard (to be proposed)
  margin_not_in_soloh   Margin Dashboard candidates not yet on the Soloh Margin Dashboard
  declined_still_active board-declined candidates still counted as active on the Margin Dashboard
  possible_matches      names that only partly match (check before treating as new)
Names Prashant said to skip (vault/state.json.enc) are left out.
"""
import argparse, json, os, re, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vault import unlock_bytes

ROOT = os.path.dirname(HERE)

def norm(s):
    s = re.sub(r"\(.*?\)", " ", str(s or "")).lower()
    return " ".join(re.sub(r"[^a-z ]", " ", s).split())

def same(a, b):
    a, b = norm(a).split(), norm(b).split()
    if not a or not b: return False
    return a == b or (a[0] == b[0] and a[-1] == b[-1])

def partial(a, b):
    a, b = norm(a).split(), norm(b).split()
    return bool(a and b) and (a[0] == b[0] or a[-1] == b[-1])

def board_data(path):
    html = open(path, encoding="utf-8").read()
    def grab(name):
        m = re.search(r"const %s = (\[.*?\n\]);" % name, html, re.S)
        if not m: sys.exit(f"Could not find const {name} in {path}")
        return m.group(1)
    js = "const C=%s;const EX=%s;process.stdout.write(JSON.stringify({C,EX}))" % (grab("C"), grab("EX"))
    out = subprocess.run(["node", "-e", js], capture_output=True, text=True)
    if out.returncode: sys.exit("Could not read board data: " + out.stderr[:500])
    return json.loads(out.stdout)

def dashboard(path, pw):
    html = open(path, encoding="utf-8").read()
    m = re.search(r'const ENC = (\{"v": 1.*?\});', html, re.S)
    if not m: sys.exit(f"Encrypted data not found in {path}")
    return json.loads(unlock_bytes(m.group(1).encode(), pw))

def state(pw):
    p = os.path.join(ROOT, "vault", "state.json.enc")
    if not os.path.exists(p): return {"skipMargin": {}, "skipSoloh": {}}
    s = json.loads(unlock_bytes(open(p, "rb").read(), pw))
    s.setdefault("skipMargin", {}); s.setdefault("skipSoloh", {})
    return s

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--board", default=os.path.join(ROOT, "..", "soloh-compliance-board", "index.html"))
    ap.add_argument("--margin", default=os.path.join(ROOT, "index.html"))
    ap.add_argument("--soloh", default=os.path.join(ROOT, "..", "soloh-margin-dashboard", "index.html"))
    ap.add_argument("--password", default=os.environ.get("DASHBOARD_PASSWORD"))
    ap.add_argument("--soloh-password", default=os.environ.get("SOLOH_PASSWORD"))
    ap.add_argument("--json")
    a = ap.parse_args()
    if not a.password or not a.soloh_password: sys.exit("Both passwords are required.")

    B = board_data(a.board)
    M = dashboard(a.margin, a.password)["records"]
    S = dashboard(a.soloh, a.soloh_password)["records"]
    st = state(a.password)
    skipM = [k for k in st["skipMargin"]]; skipS = [k for k in st["skipSoloh"]]

    res = {"board_not_in_margin": [], "possible_matches": [], "margin_not_in_soloh": [], "declined_still_active": []}
    for c in B["C"]:
        if any(same(c["name"], m["candidate"]) for m in M): continue
        if any(same(c["name"], k) for k in skipM): continue
        near = [m["candidate"] for m in M if partial(c["name"], m["candidate"])]
        item = {k: c.get(k) for k in ("id", "name", "team", "rec", "role", "fac", "src", "start")}
        if near: item["partlyMatches"] = near; res["possible_matches"].append(item)
        res["board_not_in_margin"].append(item)
    for m in M:
        if m["status"] == "Cancelled": continue
        if not any(same(m["candidate"], s["candidate"]) for s in S) and not any(same(m["candidate"], k) for k in skipS):
            res["margin_not_in_soloh"].append({k: m.get(k) for k in ("candidate", "client", "facilityUnit", "start", "end", "status", "team", "recruiter")})
    for x in B["EX"]:
        for m in M:
            if same(x["name"], m["candidate"]) and m["status"] != "Cancelled":
                res["declined_still_active"].append({"name": m["candidate"], "marginStatus": m["status"], "boardDetail": re.sub("<[^>]+>", "", x.get("detail", ""))})

    if a.json:
        with open(a.json, "w") as fh: json.dump(res, fh, indent=1)
    print(f"Board active cards: {len(B['C'])} | Margin records: {len(M)} | Soloh records: {len(S)}")
    for k, v in res.items():
        names = [i.get("name") or i.get("candidate") for i in v]
        print(f"{k}: {len(v)}" + (" -> " + "; ".join(names) if names else ""))

if __name__ == "__main__":
    main()
