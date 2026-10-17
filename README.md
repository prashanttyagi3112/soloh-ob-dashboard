# Margin Dashboard (Soloh Partners)

Password-protected Margin Dashboard for Soloh Partners: candidate pay packages, client bill rates, MSP fees, revenue and margins.
It is built from `OB_Candidates_Pay_Packages_Soloh.xlsx`.

## How the protection works
- The candidate data is **encrypted** (AES-256-GCM, key from the password via PBKDF2-SHA256, 310,000 rounds)
  and stored inside `index.html`. The page shows nothing readable until the right password is entered.
- The raw Excel workbook is **never committed** (see `.gitignore`), so the repo can safely be public.
  That matters because GitHub Pages on a free account needs a public repo.
- Anyone with the password can view the site. To revoke access, rebuild with a new password and push.

## Files
| File | Purpose |
|---|---|
| `index.html` | The published dashboard (generated, includes encrypted data) |
| `template.html` | Dashboard source (layout, charts, logic). Edit this, not `index.html` |
| `build.py` | Reads the workbook, recalculates and checks formulas, flags data issues, encrypts, writes `index.html` |
| `vendor/chart.umd.min.js` | Chart.js 4.4.1 (bundled, so the site has no CDN dependency) |

## Updating with a new workbook
```bash
pip install openpyxl cryptography
cp /path/to/OB_Candidates_Pay_Packages_Soloh.xlsx data/
python3 build.py --password "THE-PASSWORD"
git add index.html && git commit -m "Update data" && git push
```
GitHub Pages republishes within about a minute.

`build.py` matches columns **by header name**, so reordering columns or adding new ones is fine.
Extra columns appear under "Other columns" in each candidate's detail panel. A sheet whose name contains
"Cancel" or "Declin" is treated as cancelled offers, and the sheet with "Fee" in its name supplies MSP %,
NBO hours and the calculator settings.

## Calculations (same as the workbook)
- Weeks = round((End − Start) / 7)
- Gross revenue = Bill rate × Guaranteed hrs/wk × Weeks
- Net revenue = Gross revenue × (1 − MSP %)
- Net profit = Net revenue − Weeks × (min(G,40) × W2 + max(G−40,0) × OT) × (1 + payroll load)
  − Weeks × Stipend − NBO hrs × Bill × (1 − MSP %) − onboarding cost
- Weekly gross pay (nurse) = min(G,40) × W2 + max(G−40,0) × OT + Stipend

## Candidate status: Onboarding, Started, Cancelled
Status comes from the workbook sheet a candidate sits on:
- **OB Candidates** → Onboarding (accepted, not started yet)
- **Started Candidates** → Started
- **Cancelled Offers** → Cancelled (kept out of the active totals)

When the Soloh Compliance Board marks a candidate as started, move their row and rebuild:
```bash
python3 move_candidate.py --xlsx data/OB_Candidates_Pay_Packages_Soloh.xlsx --name "Full Name" --to "Started Candidates"
python3 build.py --password "THE-PASSWORD"
```
`move_candidate.py` keeps every formula pointing at its own row, so the calculations move with the candidate.

## Team and Recruiter
Each candidate sheet has **Team** and **Recruiter** columns (copied from the Soloh Compliance Board).
They drive the Team and Recruiter filters, the recruiter line under each name, the cards and the detail panel.
New candidates get both values from the Compliance Board when they are added.

## Automation: new board cards → Margin Dashboard → Soloh Margin Dashboard
A scheduled Claude task ("Margin sync") runs right after every Soloh Compliance Board update.
It never adds anyone on its own: it proposes, asks Prashant, and only adds after he says yes.

| File | Purpose |
|---|---|
| `vault/workbook.xlsx.enc` | The pay-package workbook, encrypted with the dashboard password (raw .xlsx is never committed) |
| `vault/state.json.enc` | Encrypted list of people Prashant said not to add (so he isn't asked again) |
| `automation/vault.py` | `lock` / `unlock` a file with the dashboard password |
| `automation/sync_check.py` | Read-only: board cards missing from this dashboard, and this dashboard's candidates missing from the Soloh Margin Dashboard |
| `automation/add_candidate.py` | Adds one approved candidate row; formula columns are copied and renumbered from the row above |
| `automation/state.py` | `skip-margin`, `skip-soloh`, `unskip`, `show` |
