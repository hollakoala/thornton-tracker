# Thornton Tracker (cloud, GitHub Actions)

Daily automated tracker for Creekside unit prices at Thornton Place.
Runs on GitHub's servers on a schedule - no local machine or Cowork
session needs to be on.

## One-time setup

1. Upload your current `Thornton Tracker.xlsx` into `data/Thornton_Tracker.xlsx`
   in this repo (exact filename matters). This seeds price history - the
   script will refuse to run if this file is missing, rather than starting
   a fresh tracker from scratch.
2. Make sure GitHub Pages is enabled: **Settings > Pages > Source: Deploy
   from a branch > Branch: main, folder: /docs**.
3. That's it. The workflow in `.github/workflows/track.yml` runs daily at
   13:00 UTC (5am Pacific) and will:
   - scrape thornton-place.com for current Creekside listings
   - update `data/Thornton_Tracker.xlsx` (same formulas/layout as before)
   - append today's prices to `data/price-history.csv`
   - rebuild `docs/index.html` (your live dashboard)
   - commit everything back to the repo automatically

## Checking on it

- **Live dashboard**: `https://<your-username>.github.io/<repo-name>/`
- **Run manually**: repo's *Actions* tab -> "Track Thornton Place Creekside
  Prices" -> *Run workflow*
- **See what changed on any given day**: click into `data/price-history.csv`
  in the repo and use GitHub's built-in history/diff view on the file
- **If a run fails**: GitHub emails the account that owns the repo automatically

## Notes

- Website structure changes could break the scraper silently returning
  zero units - the script treats an empty scrape as a no-op (nothing
  written) rather than assuming every unit went off-market. Check the
  Actions log if the dashboard stops updating.
- Spec mismatches (sqft, etc.) are flagged in the Actions log, never
  auto-corrected in the tracker - same as the original routine.
