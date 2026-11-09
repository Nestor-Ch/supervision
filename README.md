# Супервізія: опитування педагогів — dashboard

Static Quarto dashboard (no server), published with GitHub Pages from the `docs/` folder.

## Publish
1. Edit `index.qmd` / `inputs/dashboard_data/sheet_classification.xlsx`, then `python check_sheet.py` (fix every `[ERROR]`/`[DATA]`).
2. `quarto render index.qmd` → writes `docs/index.html` (keep `docs/.nojekyll`).
3. Commit and push; on GitHub: Settings → Pages → *Deploy from a branch* → `main` / `/docs`.

## Privacy
The repository is public. Only aggregated tables (`inputs/dashboard_data/*.parquet`) and the page sheet are committed.
Respondent-level exports (`inputs/supervision_df_*`, `*readable*`) are git-ignored — never commit them.

The page loads Plotly.js from a CDN, so viewers need internet access.
