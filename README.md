# Soloh OB Pay Packages Dashboard

Password-protected dashboard for Soloh Partners' accepted (OB) candidates and their pay packages.
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
