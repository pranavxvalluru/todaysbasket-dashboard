# Today's Basket — Dashboard V4.9

This is the next release of the **Today's Basket Store Dashboard & Analysis** website. It continues using your existing Supabase project and private report-storage backend.

## Changes in V4.9

- Replaced text symbols in the navigation sidebar with consistent inline SVG icons, refined active/hover states and subtle icon-press animations. No icon font or PNG dependency.
- Removed the `TODAY'S BASKET` eyebrow line above **Welcome back** on the login card; the brand/logo in the left panel is unchanged.
- Added a dedicated `reset-password.html` callback page and `reset-password.js` handler to let employees choose a new password after clicking their email reset link.
- Added helpful messages for invalid, expired and already-used password reset links, and retained the original login page, employee-ID sign-in, reports and analytics.
- **No existing Supabase accounts, reports or database data have been changed** in this release.

## Upload to GitHub Pages

Place **all eight** website files at the repository root (beside `index.html`):

- `index.html`
- `style.css`
- `app.js`
- `reset-password.html`  **(new, required)**
- `reset-password.js` **(new, required)**
- `terms.html`
- `privacy.html`
- `README.md`

Do not upload the ZIP as a file. Unzip it and upload its contents. GitHub Pages should be set to deploy from the `main` branch and `/ (root)`. Wait for GitHub Pages to finish deployment and force-refresh your browser.

## CRITICAL: Supabase password reset setup

If an employee receives a reset email but clicking its link opens an error, it is often because **Supabase Authentication → URL Configuration** still points at localhost or does not include your exact deployed callback URL. This account-level setting must be corrected in Supabase; changing HTML alone is not sufficient.

1. Copy your actual live GitHub Pages website address from your browser, for example:
   `https://YOUR-USERNAME.github.io/todays-basket-dashboard/`
2. Open your Supabase project at:
   `https://supabase.com/dashboard/project/kggypoblcagwqqtshixd/auth/url-configuration`
3. Set **Site URL** to the full live website address ending in `/` (e.g., the URL above).
4. Under **Redirect URLs**, click **Add URL**, and add the exact callback path:
   `https://YOUR-USERNAME.github.io/todays-basket-dashboard/reset-password.html`
   Replace `YOUR-USERNAME` and the repository name with your actual URL. Don't leave this example value unchanged.
5. Save the Supabase URL Configuration.
6. Once the NEW GitHub deployment and Supabase URL settings are in place, go to your live website → **Forgot password?**, enter the email registered with that employee's Supabase account, and request a **fresh** link. Old links may already be expired or may point to the wrong callback.
7. Open the newest email link **once**, enter a new strong password on the reset page, and then sign in with Employee ID / Username and that password.

**If it still fails:** look at the browser address bar and error message. `otp_expired` / `access_denied` usually means the link expired, was reused, or is invalid. A redirect to `localhost:3000` or the Supabase dashboard normally means your Site URL or Redirect URLs are not configured correctly. Check that the email template (if customized) uses Supabase's confirmation URL instead of manually building a link with the Site URL. The current release cannot change those project settings automatically.

> **Do not use your Supabase service-role key or database password in GitHub/HTML.** The project uses the browser-safe publishable key; private report access is enforced by Supabase authentication and policies.

## Important security / rollout notes

- The website remains static HTML + CSS + JS hosted on GitHub Pages; Supabase handles employee authentication and private Excel storage.
- Users sign in with Employee ID / username. Forgot Password correctly requires the email associated with the user's Supabase Auth account.
- The reset page is available to any visitor but only accepts a valid Supabase recovery session. The password is saved by Supabase, not the website.
- An SMTP provider may be necessary for reliable email delivery as staff count and reset frequency grow.
- Enable leaked-password protection in Supabase Auth if available, and use unique passwords of at least 12 characters.

## Developer tests

Static DOM, navigation animations, sidebar position, mobile submenu and reset-page valid/expired-link flows were tested with Chromium and mocked Supabase sessions. Live email delivery, real confirmation URL redirects and real password changes were **not tested on the production account** because they require the exact GitHub Pages URL and a real employee email link.


## V5.0 changes
- Indent Control category chips remain on **one horizontally scrolling row** even on desktop.
- Data Analysis has a Brand filter applied across transactions, charts, and SKUs. The filter reads real Brand/BrandName fields from the Total Sales Report; missing brands are marked *Unspecified brand*, never invented from product names.
- Password recovery shows a green success notification when a reset request is accepted. Local preview directs employees to the hosted dashboard without the former error sentence.
- After updating GitHub, hard-refresh to clear cached JavaScript/CSS.


## V5.1 corrections
- Indent category filtering uses a compact native searchable-by-keyboard dropdown, alphabetically ordered with counts; no giant chip row.
- SKU/Product cells have a dedicated readable column, with UOM on a separate label.
- Removed Subcategory from Data Analysis smart filters.
- Fixed missing Brand values in the background parser for large Total Sales reports. Brand is read from the actual spreadsheet Brand column; existing cached reports created by older versions need to be uploaded again to populate this new field.
- Password-reset modal fields stack and align consistently.
