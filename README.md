# LuxeShade Quote Builder

Tablet-friendly web app for Part 1 of the quotation workflow:

- import a CSV source pricelist
- select client materials
- set asking prices and see pocket
- encode measurements in millimeters
- calculate subtotal, discount, final total, and 50/50 payment split

## Files

- `index.html`: app shell
- `styles.css`: tablet-first UI styling
- `app.js`: CSV parsing, state, calculations, and rendering
- `data/pricelist.csv`: bundled sample pricelist based on your current source file

## How To Run

### Simplest option

Open `index.html` directly in your browser.

Notes:

- CSV upload works when opening the file directly.
- The `Load Bundled Sample Pricelist` button may be blocked by the browser when using `file://`.

### Better option

Serve the folder with a simple local web server, then open it in the browser.

Examples:

- Python: `python -m http.server 8000`
- Then open `http://localhost:8000`

If you use a local server, the bundled sample pricelist button should work.

## Current Formula Logic

Square footage:

```text
ROUND(((width_mm / 1000) * (height_mm / 1000)) * 10.76)
```

Row cost:

```text
rounded_sqft * asking_price
```

Pocket:

```text
asking_price - retail_price
```

Totals:

- `Subtotal` = sum of row costs
- `Final Total` = subtotal - discount
- `50% Downpayment` = final total / 2
- `50% Remaining Balance` = final total / 2

## Notes

- Material options come from `CATEGORY`.
- Retail prices come from `FABRIC WITH ACETATE SQFT`.
- The app stores its state in browser local storage so reloads do not wipe the current quote immediately.
