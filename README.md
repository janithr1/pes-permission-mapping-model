# PES Field Mapping — auto-deployed viewer

Excel is the source of truth. Push a new copy of `pes-field-mapping-workbook.xlsx`
to `main` and GitHub Actions rebuilds `data.json` and redeploys the viewer to
GitHub Pages automatically. No manual export step.

## One-time setup (5 minutes)

1. Create a GitHub repo and push everything in this folder to `main`.
2. Repo → **Settings → Pages → Build and deployment → Source** → select
   **GitHub Actions** (not "Deploy from a branch").
3. Push once (or run the workflow manually: **Actions → Build and deploy PES
   viewer → Run workflow**). First run creates the site at
   `https://<your-org>.github.io/<repo>/`.

## Day to day

1. Edit `pes-field-mapping-workbook.xlsx` locally (Excel, LibreOffice, whatever).
2. Commit and push it back to `main`, same filename, same path.
3. Actions runs automatically, converts it, redeploys. Usually live in under
   a minute — check the Actions tab for progress or failures.

## Files

- `pes-field-mapping-workbook.xlsx` — the workbook. Keep tab names exactly as
  they are (`Application`, `Contract`, `Monitoring Visit`, `Remediation`,
  `Payment`, `Beneficiary`, `Project & Activity`, `Files`) and keep the column
  order on row 1 — the converter reads by column position.
- `scripts/xlsx_to_json.py` — the conversion script. Edit this if you add or
  reorder columns.
- `template/index.html` — the viewer page. Fetches `data.json` at load; no
  external dependencies, no build step beyond the copy.
- `.github/workflows/deploy.yml` — the pipeline. Triggers on any push that
  touches the workbook, the scripts, or the template.

## What this version can't do

This is the static/GitHub Pages build: read-only in the browser — no
in-page editing, no shared "Saved" state (those need the claude.ai runtime).
Editing happens in Excel; the page always reflects whatever was last pushed.
CSV export per module/tab still works client-side.
