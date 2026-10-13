# Today's Basket V4.8 — Staff Security & Policies

## GitHub deployment
Upload `index.html`, `app.js`, `style.css`, `terms.html` and `privacy.html` together to the repository root.

## Changes
- Employees see their own private workbook list but not cloud-storage totals or organization-wide usage. Those UI panels and cleanup controls are admin-only.
- Normal login uses Employee ID/username + Supabase password.
- Forgot Password requires the employee's **registered email address**; Supabase emails a reset link.
- Change Password is available from the authenticated dashboard header. Requires current password and new strong password (12+ characters, letters, numbers and symbols).
- Terms and privacy disclosures are included as **draft documents**. Obtain company/legal approval and add administrator contact details before production.
- Existing V4.7 report upload and analysis features remain unchanged.

## CRITICAL: configure password recovery redirect
In Supabase Dashboard → Authentication → URL Configuration, add your exact GitHub Pages URL (e.g., `https://YOUR-USERNAME.github.io/todays-basket-dashboard/`) to **Redirect URLs**. Supabase's email provider must be configured and able to deliver password reset emails. Test the reset flow from an incognito browser. Password recovery uses `detectSessionInUrl:true`.

## Employee 15269 setup
The Harish account could not be assigned automatically because the connected database write for that user was blocked. Verify the user exists in Authentication → Users and copy that user UUID. Run `LINK_EMPLOYEE_15269.sql` using the Supabase SQL Editor, after verifying the intended user owns that email. The SQL links that existing Auth user to ID 15269 if neither the account nor ID already has a staff profile. If a profile exists, review and update it manually rather than creating a duplicate. **Do not use password `1234`**. In Authentication → Users, trigger a reset/invite flow so Harish sets a secure unique password. If the row already exists, edit it rather than adding a duplicate.

## Security / admin-only storage stats
The V4.8 HTML hides storage size and cleanup controls from non-admin users, and the My Reports page still shows only each employee's own files. **REQUIRED**: Run `ADMIN_STORAGE_SECURITY.sql` in the Supabase SQL Editor *before* employee deployment. The existing RPC still gives non-admins access to some storage totals unless this SQL is executed. UI hiding alone is not secure.

## Known limits
The password reset redirect must be configured manually in Supabase. Reports remain subject to the plan's quotas and testing. Terms and policy drafts are not legal advice. No initial employee password has been assigned by this release.
