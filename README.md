# Today’s Basket — Store Dashboard & Analysis (V4.2)

This is the **separated-file** edition of the V4.2 browser prototype. It has the same dashboard as the single-file version.

## Contents

- `index.html` — page layout, login portal, and embedded logo images (Base64).
- `style.css` — fonts, visual styling, layout, and animations.
- `app.js` — login flow, navigation, report parsing, and dashboard functionality.
- `README.md` — these instructions.

## Run locally

Keep all four files in the **same directory**. Double-click `index.html` in Finder to open it in Chrome, or serve the directory locally if your browser restricts file access. Excel parsing uses the SheetJS library loaded over the internet. Google Fonts also require internet.

## Publish to GitHub Pages

1. Create or open your repository.
2. Upload all four files to the repository's top level (**not a ZIP** and not inside a subfolder).
3. Commit the files.
4. Open **Settings → Pages**, select **Deploy from a branch → main → /(root)**, then Save.
5. Once deployed, open `https://YOUR-USERNAME.github.io/YOUR-REPOSITORY/`.

## Security

This is a prototype: its authentication is client-side. Publishing on a public GitHub Pages repository makes HTML, CSS, and JavaScript publicly accessible. **Do not use this version to safeguard real employee credentials or confidential data.** Move authentication and sensitive data handling to a secure backend before production usage.
