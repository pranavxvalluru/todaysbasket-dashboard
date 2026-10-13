# Today's Basket — Store Dashboard & Analysis
## Version 5.7 · GitHub Pages + Supabase

This is the **V5.7** website release. The interface is built with HTML, CSS and JavaScript. GitHub Pages hosts the website; Supabase handles employee sign-in, private report storage and database permissions.

### Previous login layout improvements (V5.3)
- Centered **Forgot password?** immediately below the Sign in button.
- Positioned **Terms & Conditions** on the lower left and **Privacy Policy** on the lower right of the login card.
- Matched the headings on the policy pages to the login links.
- Kept existing reports, filters, calculations, navigation and Supabase integration unchanged.

### Other features currently included
- Employee ID/username sign-in using Supabase authentication.
- Employee-specific saved reports with upload timestamps and reloading after sign-in.
- Administrator-managed official reports, with separate employee report workspaces.
- GRN report navigation: month → GRN date → complete GRN number → related product lines.
- Data Analysis with Month, Floor, Category, Brand and SKU filters.
- Stock Requirements and Indent Workspace, including a compact category selector.
- Account recovery and change-password UI.
- Private report replacement/cleanup logic and report-storage monitoring for administrators.

**Important:** Features requiring live backend access still need end-to-end testing on your deployed site. Browser-only tests do not confirm successful uploads, password-reset emails, or database permissions.

## Project files

| File | Purpose |
| --- | --- |
| `index.html` | Login screen and application layout |
| `style.css` | Styles, responsiveness and animations |
| `app.js` | Dashboard navigation, Excel parsing and Supabase integration |
| `reset-password.html` | Password reset page opened from the recovery email |
| `reset-password.js` | Password reset and session callback handling |
| `terms.html` | Draft Terms & Conditions |
| `privacy.html` | Draft Privacy Policy |
| `README.md` | Installation, configuration and release notes |

## Deploy to GitHub Pages

1. Extract the ZIP.
2. Upload **all eight files** to the root of your existing GitHub repository (not the ZIP itself).
3. Replace the older versions and commit your changes.
4. In GitHub: **Settings → Pages → Deploy from a branch → main → / (root)**.
5. Open the published `.github.io` link and refresh the page after deployment. You may need a hard refresh to clear cached CSS or JavaScript.

**Do not upload confidential inventory or sales Excel reports to the public GitHub repository.** Store them only in the private Supabase buckets through the application, subject to access controls.

## Supabase connection

The website is configured to use the existing **Today's Basket** Supabase project. Do not put any database password, secret API key or `service_role` key in website files. The browser-safe publishable key is not an administrator secret; report access depends on Supabase authentication and Row Level Security.

Employee account records and designations are managed in Supabase. Each person should have their **own** account and password, not a shared employee login. The role displayed in the interface is distinct from the designation/title.

### Forgot Password: required deployment settings

Open Supabase **Authentication → URL Configuration**:
https://supabase.com/dashboard/project/kggypoblcagwqqtshixd/auth/url-configuration

- **Site URL:** Set to your exact deployed website root URL, for example `https://YOUR-USERNAME.github.io/todays-basket-dashboard/`.
- **Redirect URLs:** Allow the exact callback URL, for example `https://YOUR-USERNAME.github.io/todays-basket-dashboard/reset-password.html`.
- Replace the placeholders with your actual GitHub username and repository name.

Then:
1. Visit the **published** website and click **Forgot password?**.
2. Enter the email address registered in that employee's Supabase Auth account.
3. Open the **newest** reset email link once.
4. Set a strong, unique password.
5. Sign in with the employee ID/username and the newly chosen password.

Expired or reused links must be requested again. If the browser redirects to localhost or an unexpected website, verify Supabase Site URL, Redirect URLs and any customized email template.

## Data handling and retention

- **Official reports:** Admin-published data is stored in a separate private Supabase bucket.
- **My Reports:** Employee uploads are scoped to the signed-in employee and restored after login.
- **Replacement:** The latest successful upload is made active first; older workbook files may then be removed by cleanup. This is not an archival backup system.
- **Timestamps:** Upload date/time is shown for each saved report.
- **Large reports:** Browser parsing may take time; test your actual inventory and annual sales workbooks before staff rollout.
- **Brand values:** If a Total Sales report was imported before the Brand parser fix in V5.1, re-upload it to restore real Brand filters.
- The Total Sales Report is used for **Data Analysis**; the 30-day Summary Sales by Quantity report is used for **Indent Workspace**. Keep them separate.

## Terms & Conditions and Privacy Policy

`terms.html` and `privacy.html` are **draft documents**, not approved corporate policies. Before rollout, management/legal must confirm the correct business entity, contact details, permitted use, retention periods, employee notice/consent where required, incident response, and applicable law. Do not present draft text as a legally certified policy.

## Deployment verification checklist

- [ ] The website opens at the correct GitHub Pages URL.
- [ ] Login, Logout and Forgot Password work with **real** Supabase accounts.
- [ ] Password-reset emails return to the correct reset page.
- [ ] Admin can publish an official report and see it after logging back in.
- [ ] Employee can upload a personal report and see it again after logging out and back in.
- [ ] An employee cannot access another employee's private reports or administrator-only controls.
- [ ] GRN grouping uses the **GRNDate** field plus full GRN number (not a partial serial number).
- [ ] Month, Floor, Category, Brand and SKU filters calculate correctly.
- [ ] Store/Warehouse Stock Requirements and indent quantities match original reports.
- [ ] Report replacement does not remove a working file when the new upload fails.
- [ ] Layout is checked on a 13-inch laptop, larger desktop, mobile, Chrome, Safari and Edge.
- [ ] Actual report data is not committed to the public GitHub repo.

## Version history

- **V4.9:** Premium sidebar icons and dedicated password-reset page.
- **V5.0:** Brand filtering groundwork, category filter refinements and reset notices.
- **V5.1:** Real Brand parsing in the large-report path, compact indent category selector, product-name layout and aligned reset fields.
- **V5.2:** Expanded internal policies, responsive safeguards and account-flow fixes.
- **V5.3:** Centered password-recovery link and revised Terms & Conditions / Privacy Policy login layout.

### Support and changes

Edit `index.html`, `style.css`, and `app.js` for future UI features; update GitHub Pages to publish changes. Supabase report data and employee accounts are stored separately and do not need to be recreated after a website update.

## V5.4 — Employee Management

- Added **Employee Management** to the sidebar, visible only to the signed-in administrator profile **15907**.
- Added staff list, designation and account-role editing, and active/inactive controls.
- **Register employee** links an existing Supabase Auth user by email or sends an email invitation to a new user. The user chooses a password privately. Administrator cannot see passwords.
- Unauthorized navigation shows an **Access Restricted** page with a Back to Dashboard button. The server registration endpoint requires a valid signed-in JWT and checks the administrator's Supabase profile; ordinary staff cannot use that endpoint.
- The new `register-employee` Edge Function was deployed to the existing Supabase project. Company data is not bundled in this ZIP.
- An inactive user's already-issued token may remain active until its expiry; for urgent revocation, an administrator should revoke that user's sessions in Supabase Auth.
- The application is still in development. Run live account invitation and staff-permission tests on the published site before enabling broad staff access. Manager role alone does not grant Employee Management access.

## V5.5 — Supabase Password Update Flow

- **Change Password** updates Supabase Auth directly using `sb.auth.updateUser({password})` after verifying the currently signed-in employee and their current password with an isolated authentication client. It does not alter `staff_profiles`, Employee ID, role, or reports.
- Prevents an employee from choosing the same password as their current password; new passwords must contain 12+ characters, letters, a digit and a symbol, and match confirmation.
- After a successful change, clears password inputs, signs out and prompts login with the new password. A success message refers explicitly to Supabase.
- **Forgot Password** requests a Supabase recovery email sent to the user's registered email and opens `reset-password.html`. The reset page calls Supabase Auth directly and signs out before returning to the login page.
- Password fields are **never saved to browser storage or dashboard tables**. The employee ID remains unchanged after a password update.

### Live authentication tests required

1. On the deployed GitHub Pages URL, log in as a test employee, select Change Password, verify an incorrect current password is rejected, and confirm the same password is not accepted as a new one.
2. Save a valid new password and confirm the account can sign in with **Employee ID + new password**, while the old password fails.
3. Test Forgot Password with a registered email, open the freshly generated reset link, set a different strong password and verify login.
4. Verify Supabase **Authentication → URL Configuration** allows the exact deployed `reset-password.html` redirect.
5. Repeat using a non-admin staff account without altering access permissions or stored reports.

**Security and release limitation:** This ZIP was checked statically; real Supabase authentication, email delivery, permission enforcement, and all browsers cannot be certified without an authorized live end-to-end session.


## V5.6 — Daily Sales table width fix
- Fixed the large empty area beside the **Data Analysis → Daily sales** date table. The Date, Sales Value, Reported Qty and Inspect day columns now span the full available report card width.
- The table remains horizontally scrollable on narrow devices.
- Other inventory, GRN and SKU tables retain their original sizing.
- No sales data, dashboard calculations, user accounts, or Supabase backend settings were changed.


## V5.7 — cross-screen layout quality review
- Expanded Data Analysis **SKU performance** table to occupy its full report panel instead of ending at a fixed minimum width.
- Smart Filters now auto-fit the available width rather than reserving an empty column on wide monitors.
- Standard report table wrappers use the panel width; wide operational grids (Indent, GRN, employee directory) retain horizontal scrolling with readable minimum column sizing.
- Small screen action groups and panel headers wrap instead of extending beyond their cards.
- Audited at 360, 390, 600, 768, 1024, 1280, 1366, 1440, 1920, and 2560px in headless Chromium. Sample imported sales transactions used for populated Data Analysis scenarios. Checked all eight project files for ZIP integrity, JS syntax, local references and basic HTML markup validation.

### Boundaries of these checks
A viewport sweep is not a guarantee that every Excel report, live account, device or browser will be bug-free. **Do not call the deployment certified** until real employee/admin logins, Supabase storage permissions, password reset emails, report replacement and representative large workbooks are tested on the final GitHub Pages URL. Safari, Firefox and Edge were not independently exercised in this build. The existing Supabase security advisories should also be reviewed before broad staff rollout.
