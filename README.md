# LUXESHADE Quote Builder

Tablet-friendly web app for quoting window treatments (curtains, blinds, motorized shades). Plain HTML/CSS/JS with no build step, backed by Supabase.

## Features

- Sign in with Google (Supabase OAuth) to save, load, and delete quotes.
- Load the material catalog from Supabase, or import a CSV pricelist.
- Select client materials, set asking prices, and see the pocket (margin).
- Encode measurements in millimeters:
  - square footage is calculated with a 15 sq ft minimum per line
  - per-line type, control (`SPLIT` / `FULL`), panel count, and motor quantity for motorized items
  - free lines, lines excluded from the discount, and note rows
  - drag to reorder, and `COPY ABOVE` for sheer/blackout pairs
- Totals: subtotal, discount, additional charges (installation, steaming), final total, and a 50/50 downpayment and balance split.
- PDF output (pdfmake, generated in the browser):
  - contract/quote and acknowledgement receipt
  - brands: LuxeShade, NDS Trading, Kurtina Kultura, or unbranded (Admin Tools option)
- Local draft kept in browser storage, plus a save reminder (autosave is currently disabled).

## Files

- `index.html`: app shell
- `styles.css`: tablet-first UI styling
- `app.js`: state, calculations, rendering, Supabase access, PDF building, CSV parsing
- `data/pricelist.csv`: bundled sample pricelist
- `assets/`: brand logos
- `docs/`: business registration documents
- `supabase/`: `schema.sql` plus incremental `*_migration.sql` scripts
- `.cursor/rules/version-rollover.mdc`: version numbering rule

## How To Run

Serve the folder with any static server and open it in the browser:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

Notes:

- Opening `index.html` directly (`file://`) works for CSV upload, but the bundled sample button and Google sign-in may be blocked.
- `supabase-js` and `pdfmake` load from the jsDelivr CDN, so an internet connection is required.

## Supabase Setup

1. Create a project and run `supabase/schema.sql` in the SQL editor.
2. Run the `supabase/*_migration.sql` scripts as needed (they are applied manually). `quote_security_migration.sql` sets owner-based row level security for quotes.
3. Enable the Google auth provider.
4. Set `SUPABASE_URL` and `SUPABASE_ANON_KEY` near the top of `app.js`.
5. Optionally seed the catalog from `supabase/material_catalog_import.csv`.

Tables: `material_catalog`, `quotes`, `quote_materials`, `quote_measurements`.

## Pricing

- Material options come from `CATEGORY`.
- Retail prices come from `FABRIC WITH ACETATE SQFT`.
- `Subtotal` = sum of line costs
- `Final Total` = subtotal - discount (plus additional charges)
- `50% Downpayment` = final total / 2
- `50% Remaining Balance` = final total / 2

## Versioning

The version appears in the page title, asset query parameters, and the visible label in `index.html`. Patch rolls over base 10 into minor (`0.0.11` is written `0.1.1`); see `.cursor/rules/version-rollover.mdc`.
