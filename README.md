# Today’s Basket — Dashboard V4.0

## What changed

- Login welcome label is **Store Dashboard & Analysis**.
- Successful login includes a smooth **4.3-second** loading screen, followed by a **0.26-second** transition into the dashboard.
- The progress indicator, status text and a modern door/arrow logout icon are updated.
- The prior full dashboard remains intact: GRN, Indent, Data Analysis, Smart Filters, and Excel tools.

## Start

Extract all files into one folder and open `index.html` in current Chrome, or serve with `python3 -m http.server 8000` and visit `http://localhost:8000`.

All login checks are browser-side demo checks, **not secure authentication**. Do not publish sensitive reports behind this login. Use server authentication when going live.

Excel upload depends on SheetJS served from a public CDN, so Internet is required unless that dependency is bundled separately.
